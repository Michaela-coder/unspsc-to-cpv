
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UNSPSC → CPV Matcher (GitHub-ready)

- Reads UNSPSC input, multilingual CPV dictionary, and full UNSPSC names
- Matches via three signals: fuzzy, TF‑IDF, and semantic embeddings
- Decision logic prioritizes semantic score (score_sem), TF‑IDF is second
- Exports Excel with tabs: ALL, 1_confident, 2_review, 3_no_match + CSV mapping
- Optional DeepL translation for Czech fallbacks (if DEEPL_API_KEY in .env)
- All paths can be provided via CLI; defaults use current working directory
- Python 3.10+ recommended

Usage (examples):
    python unspsc_to_cpv.py \
        --input unspsc_input.csv \
        --cpv cpv_2008_cz.xlsx \
        --unspsc-full unspsc_full.xlsx \
        --out-xlsx unspsc_to_cpv.xlsx \
        --use-deepl

Threshold philosophy (defaults):
    - SEMANTIC is primary signal:
        * confident  : score_sem ≥ 85
        * review     : 70 ≤ score_sem < 85
        * tie-breaker: must beat others by SEM_MARGIN (default 5) unless ≥ 85
    - TF‑IDF is secondary:
        * confident  : score_tfidf ≥ 88
        * review     : 75 ≤ score_tfidf < 88
    - Fuzzy is only a fallback (uses traditional HIGH/LOW = 85/70).
    You can tune thresholds from CLI.

Author: Michaela-friendly edition 🫶
"""
from __future__ import annotations

import os, re, sys, logging
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Tuple, List

import numpy as np
import pandas as pd
from rapidfuzz import process, fuzz

# ---------- Logging ----------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
log = logging.getLogger("unspsc2cpv")

# ---------- CLI ----------
import argparse

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Match UNSPSC to CPV with fuzzy, TF‑IDF, and semantic signals.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    # Files
    p.add_argument("--input", dest="input_file", default="unspsc_input.csv",
                   help="CSV: first col = UNSPSC code, second col = short EN description")
    p.add_argument("--cpv", dest="cpv_file", default="cpv_2008_cz.xlsx",
                   help="CPV dictionary (XLSX/XLS) with CODE + EN + CS columns (or similar)")
    p.add_argument("--unspsc-full", dest="unspsc_full", default="unspsc_full.xlsx",
                   help="Official UNSPSC file (code + full EN name)")
    p.add_argument("--out-xlsx", dest="out_xlsx", default="unspsc_to_cpv.xlsx",
                   help="Output Excel path")
    p.add_argument("--out-csv", dest="out_csv", default="unspsc_to_cpv_mapping.csv",
                   help="Optional CSV mapping path")
    # Thresholds
    p.add_argument("--high", type=int, default=85, help="Generic HIGH threshold (fuzzy)")
    p.add_argument("--low",  type=int, default=70, help="Generic LOW threshold (fuzzy)")
    p.add_argument("--sem-confident", type=int, default=85, help="Semantic confident threshold")
    p.add_argument("--sem-review",    type=int, default=70, help="Semantic review threshold")
    p.add_argument("--tfidf-confident", type=int, default=88, help="TF‑IDF confident threshold")
    p.add_argument("--tfidf-review",    type=int, default=75, help="TF‑IDF review threshold")
    p.add_argument("--sem-margin", type=int, default=5, help="How much semantic must beat others (unless ≥ sem-confident)")
    # Matching knobs
    p.add_argument("--topk-tfidf-for-sem", type=int, default=50,
                   help="How many TF‑IDF candidates to re‑rank with semantic")
    p.add_argument("--use-deepl", action="store_true", help="Use DeepL if DEEPL_API_KEY exists in .env")
    return p

# ---------- Helpers ----------
def normalize_unspsc(code) -> str:
    s = str(code)
    s = re.sub(r"\D", "", s)
    return s[:8].zfill(8)

def normalize_cpv(code) -> Optional[str]:
    if pd.isna(code): return None
    s = str(code).strip().replace(" ", "")
    m = re.search(r"(\d{8})(?:-(\d))?$", s)
    if not m: return None
    return m.group(1) + ("-" + m.group(2) if m.group(2) else "")

def engine_for(path: Path):
    return "xlrd" if path.suffix.lower() == ".xls" else None

def load_cpv_multilang(path: Path) -> pd.DataFrame:
    """Read CPV XLSX/XLS (CODE, EN, CS, …) -> cpv_code, cpv_en, cpv_cs."""
    df = pd.read_excel(path, dtype=str, engine=engine_for(path))
    # normalize headers
    cols = {c: re.sub(r"\s+", "", str(c)).upper() for c in df.columns}
    df = df.rename(columns=cols)
    code_col = next((c for c in df.columns if c in {"CODE","CPVCODE","KOD","KÓD"}), None)
    en_col   = next((c for c in df.columns if c in {"EN","ENGLISH"}), None)
    cs_col   = next((c for c in df.columns if c in {"CS","CZECH","CZE"}), None)
    if code_col is None:
        for c in df.columns:
            if df[c].astype(str).str.match(r"^\d{8}-\d$").sum() > 50:
                code_col = c; break
    if code_col is None:
        raise RuntimeError("CPV file: could not detect code column (e.g., CODE).")
    out = pd.DataFrame()
    out["cpv_code"] = df[code_col].map(normalize_cpv)
    out["cpv_en"] = df[en_col].astype(str) if en_col in df.columns else None
    out["cpv_cs"] = df[cs_col].astype(str) if cs_col in df.columns else None
    out = out.dropna(subset=["cpv_code"]).drop_duplicates(subset=["cpv_code"])
    return out

def try_import_deepl(base_dir: Path):
    try:
        from dotenv import load_dotenv
        load_dotenv(base_dir / ".env")
        import deepl
        key = os.getenv("DEEPL_API_KEY")
        if key:
            return deepl.Translator(key)
    except Exception:
        pass
    return None

def norm_text(s: str, allow_cz=False) -> str:
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9á-ž\s\-\/&]" if allow_cz else r"[^a-z0-9\s\-\/&]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# small synonymization (optional, for precision)
SYNONYMS = {
    r"\blaptop(s)?\b": "portable computer",
    r"\bnotebook(s)?\b": "portable computer",
    r"\bplant nutrient(s)?\b": "fertilizer",
    r"\bherbicide(s)?\b": "weed killer",
    r"\brepellent(s)?\b": "repellent",
    r"\btermite shield(s)?\b": "termite barrier",
}
def apply_synonyms(s: str) -> str:
    t = s
    for pat, repl in SYNONYMS.items():
        t = re.sub(pat, repl, t)
    return t

# ---------- Core ----------
@dataclass
class Thresholds:
    high: int = 85
    low: int  = 70
    sem_confident: int = 85
    sem_review: int    = 70
    tfidf_confident: int = 88
    tfidf_review: int    = 75
    sem_margin: int      = 5

def status_from_score_generic(score: int, high: int, low: int) -> str:
    if score >= high: return "confident"
    if score >= low:  return "review"
    return "no_match"

def main(argv=None):
    args = build_parser().parse_args(argv)
    base_dir = Path.cwd()
    th = Thresholds(
        high=args.high, low=args.low,
        sem_confident=args.sem_confident, sem_review=args.sem_review,
        tfidf_confident=args.tfidf_confident, tfidf_review=args.tfidf_review,
        sem_margin=args.sem_margin
    )

    input_file  = Path(args.input_file)
    cpv_file    = Path(args.cpv_file)
    unspsc_full = Path(args.unspsc_full)
    out_xlsx    = Path(args.out_xlsx)
    out_csv     = Path(args.out_csv)

    # --- Load UNSPSC (CSV) ---
    log.info("Loading UNSPSC input CSV: %s", input_file)
    uns = pd.read_csv(input_file, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
    if uns.shape[1] < 2:
        raise ValueError("CSV must have at least 2 columns: code and description.")
    uns = uns.rename(columns={uns.columns[0]: "unspsc_code_raw", uns.columns[1]: "unspsc_en_short"})
    uns["unspsc_code"] = uns["unspsc_code_raw"].map(normalize_unspsc)
    uns["unspsc_en_short"] = uns["unspsc_en_short"].fillna("").astype(str).str.strip()

    # --- Load full UNSPSC ---
    log.info("Loading full UNSPSC dictionary: %s", unspsc_full)
    full_uns = pd.read_excel(unspsc_full, dtype=str)
    if full_uns.shape[1] < 2:
        raise ValueError("unspsc_full must have at least 2 columns (code, full EN name).")
    full_uns = full_uns.rename(columns={full_uns.columns[0]: "unspsc_code", full_uns.columns[1]: "unspsc_full_en"})
    full_uns["unspsc_code"] = full_uns["unspsc_code"].astype(str).str.replace(r"\D","",regex=True).str.zfill(8)
    full_uns["unspsc_full_en"] = full_uns["unspsc_full_en"].fillna("").astype(str).str.strip()

    # join full name and build UNSPSC text
    uns = uns.merge(full_uns[["unspsc_code","unspsc_full_en"]], on="unspsc_code", how="left")
    uns["unspsc_en"] = np.where(uns["unspsc_full_en"].notna() & (uns["unspsc_full_en"]!=""),
                                uns["unspsc_full_en"], uns["unspsc_en_short"])

    log.info("UNSPSC rows: %d", len(uns))

    # --- Load CPV ---
    log.info("Loading CPV multilingual: %s", cpv_file)
    cpv = load_cpv_multilang(cpv_file)
    log.info("CPV rows: %d", len(cpv))

    # --- Prepare texts for matching ---
    match_on_cs = cpv["cpv_en"].isna().all()
    if match_on_cs:
        log.info("Note: EN is missing in CPV; matching on Czech labels (cpv_cs).")
        translator = try_import_deepl(base_dir) if args.use_deepl else None
        if translator:
            log.info("Translating UNSPSC EN→CS via DeepL…")
            def tr_en_cs(t):
                try:
                    return translator.translate_text(t, source_lang="EN", target_lang="CS").text if t else ""
                except Exception:
                    return t or ""
            uns["unspsc_for_match"] = uns["unspsc_en"].astype(str).apply(tr_en_cs)
        else:
            uns["unspsc_for_match"] = uns["unspsc_en"]
        cpv["label_for_match"] = cpv["cpv_cs"].fillna("")
        allow_cz = True
    else:
        uns["unspsc_for_match"] = uns["unspsc_en"]
        cpv["label_for_match"] = cpv["cpv_en"].fillna("")
        allow_cz = False

    # normalize & synonyms
    uns["unspsc_norm"] = uns["unspsc_for_match"].apply(lambda s: apply_synonyms(norm_text(s, allow_cz)))
    cpv["cpv_norm"]     = cpv["label_for_match"].apply(lambda s: apply_synonyms(norm_text(s, allow_cz)))

    # --- Fuzzy ---
    log.info("Matching (fuzzy)…")
    cpv_labels_norm = cpv["cpv_norm"].tolist()
    def fuzzy_best(desc_norm: str):
        if not desc_norm: return (None, None, 0)
        m = process.extractOne(desc_norm, cpv_labels_norm, scorer=fuzz.WRatio)
        if not m: return (None, None, 0)
        _, score, idx = m
        return (cpv.iloc[idx]["cpv_code"], cpv.iloc[idx]["label_for_match"], int(score))
    f_rows = [fuzzy_best(t) for t in uns["unspsc_norm"]]

    # --- TF‑IDF ---
    log.info("Building TF‑IDF and computing cosine similarities…")
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    vectorizer = TfidfVectorizer(min_df=1, ngram_range=(1,2))
    X_cpv = vectorizer.fit_transform(cpv["cpv_norm"])
    X_uns = vectorizer.transform(uns["unspsc_norm"])

    best_idx_tfidf, best_sim_tfidf, topk_idx_list = [], [], []
    batch = 1500
    for i in range(0, X_uns.shape[0], batch):
        sims = cosine_similarity(X_uns[i:i+batch], X_cpv)
        idx1 = sims.argmax(axis=1)
        sim1 = sims.max(axis=1)
        best_idx_tfidf.extend(idx1.tolist())
        best_sim_tfidf.extend(sim1.tolist())
        K = min(args.topk_tfidf_for_sem, sims.shape[1])
        part_idx = np.argpartition(sims, -K, axis=1)[:, -K:]
        row_sorted = np.take_along_axis(part_idx, np.argsort(np.take_along_axis(sims, part_idx, axis=1), axis=1)[:, ::-1], axis=1)
        topk_idx_list.extend(row_sorted.tolist())

    # --- Semantic ---
    log.info("Loading embedding model (first run may download ~90 MB)…")
    from sentence_transformers import SentenceTransformer, util
    model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

    cpv_text_for_sem = cpv["cpv_en"].fillna(cpv["cpv_cs"].fillna("")).tolist()
    uns_text_for_sem = uns["unspsc_en"].tolist()

    cpv_emb = model.encode(cpv_text_for_sem, convert_to_tensor=True, normalize_embeddings=True, show_progress_bar=True)

    sem_best_idx, sem_best_score = [], []
    B = 256
    for i in range(0, len(uns_text_for_sem), B):
        emb_uns = model.encode(uns_text_for_sem[i:i+B], convert_to_tensor=True, normalize_embeddings=True)
        for r, emb in enumerate(emb_uns):
            cand_idx = topk_idx_list[i + r]
            cand_matrix = cpv_emb[cand_idx]
            sims = util.cos_sim(emb, cand_matrix)  # (1, K)
            j = int(np.argmax(sims.cpu().numpy()))
            sem_best_idx.append(cand_idx[j])
            score = float(sims[0][j].cpu().item())
            sem_best_score.append(round(max(0.0, min(1.0, score)) * 100))

    # --- Compose results ---
    matches = uns[["unspsc_code", "unspsc_en_short", "unspsc_full_en", "unspsc_en"]].copy()

    # fuzzy
    matches["cpv_code_fuzz"]  = [f_rows[k][0] for k in range(len(f_rows))]
    matches["cpv_label_fuzz"] = [f_rows[k][1] for k in range(len(f_rows))]
    matches["score_fuzz"]     = [f_rows[k][2] for k in range(len(f_rows))]

    # tfidf
    matches["cpv_code_tfidf"]  = [cpv.iloc[j]["cpv_code"] for j in best_idx_tfidf]
    matches["cpv_label_tfidf"] = [cpv.iloc[j]["label_for_match"] for j in best_idx_tfidf]
    matches["score_tfidf"]     = [round(s*100) for s in best_sim_tfidf]

    # semantic
    matches["cpv_code_sem"]  = [cpv.iloc[j]["cpv_code"] for j in sem_best_idx]
    matches["cpv_label_sem"] = [cpv.iloc[j]["cpv_en"] for j in sem_best_idx]  # semantic runs on EN
    matches["score_sem"]     = sem_best_score

    # --- Decision function (semantic-first) ---
    def decide(row):
        fuzz_s  = int(row.get("score_fuzz") or 0)
        tfidf_s = int(row.get("score_tfidf") or 0)
        sem_s   = int(row.get("score_sem") or 0)

        # 1) If semantic is very strong, take it.
        if sem_s >= th.sem_confident:
            return pd.Series({"method":"semantic","cpv_code":row["cpv_code_sem"],"cpv_label_used":row["cpv_label_sem"],
                              "score":sem_s, "status":"confident"})

        # 2) If semantic is decent AND clearly better than others, mark as review.
        if sem_s >= th.sem_review and sem_s >= max(fuzz_s, tfidf_s) + th.sem_margin:
            return pd.Series({"method":"semantic","cpv_code":row["cpv_code_sem"],"cpv_label_used":row["cpv_label_sem"],
                              "score":sem_s, "status":"review"})

        # 3) Consider TF‑IDF as second best.
        if tfidf_s >= th.tfidf_confident and tfidf_s >= fuzz_s + 5:
            return pd.Series({"method":"tfidf","cpv_code":row["cpv_code_tfidf"],"cpv_label_used":row["cpv_label_tfidf"],
                              "score":tfidf_s, "status":"confident"})
        if tfidf_s >= th.tfidf_review:
            return pd.Series({"method":"tfidf","cpv_code":row["cpv_code_tfidf"],"cpv_label_used":row["cpv_label_tfidf"],
                              "score":tfidf_s, "status":"review"})

        # 4) Fallback to better of fuzzy/tfidf with generic thresholds.
        base = ("tfidf", row["cpv_code_tfidf"], row["cpv_label_tfidf"], tfidf_s) if tfidf_s >= fuzz_s \
               else ("fuzz",  row["cpv_code_fuzz"],  row["cpv_label_fuzz"],  fuzz_s)
        status = status_from_score_generic(base[3], th.high, th.low)
        return pd.Series({"method":base[0], "cpv_code":base[1], "cpv_label_used":base[2], "score":int(base[3]), "status":status})

    chosen = matches.apply(decide, axis=1)
    matches = pd.concat([matches, chosen], axis=1)

    # attach both language variants
    matches = matches.merge(cpv[["cpv_code","cpv_en","cpv_cs"]].drop_duplicates(), on="cpv_code", how="left")

    # --- Translation fallback (CZ) ---
    translator = try_import_deepl(base_dir) if args.use_deepl else None
    def translate_en_to_cs(text: str) -> Optional[str]:
        if not text or translator is None: return None
        try:
            return translator.translate_text(text, source_lang="EN", target_lang="CS").text
        except Exception:
            return None
    needs_translation = (matches["status"] == "no_match") | (matches["cpv_cs"].isna())
    matches.loc[needs_translation, "fallback_cs"] = matches.loc[needs_translation, "unspsc_en"].apply(translate_en_to_cs)

    # --- Export ---
    confident = matches[matches["status"]=="confident"].copy()
    review    = matches[matches["status"]=="review"].copy()
    no_match  = matches[matches["status"]=="no_match"].copy()

    cols = ["unspsc_code","unspsc_en_short","unspsc_full_en","unspsc_en",
            "cpv_code","cpv_en","cpv_cs","cpv_label_used",
            "method","score","status","fallback_cs",
            "cpv_code_sem","score_sem","cpv_code_tfidf","score_tfidf","cpv_code_fuzz","score_fuzz"]

    log.info("Writing Excel: %s", out_xlsx)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        matches[cols].to_excel(xw, sheet_name="ALL", index=False)
        confident[cols].to_excel(xw, sheet_name="1_confident", index=False)
        review[cols].to_excel(xw, sheet_name="2_review", index=False)
        no_match[cols].to_excel(xw, sheet_name="3_no_match", index=False)

    if out_csv:
        mapping = matches[["unspsc_code","unspsc_en","cpv_code","cpv_en","cpv_cs","score","method","status"]]
        mapping.to_csv(out_csv, index=False, encoding="utf-8")
        log.info("Mapping CSV written: %s", out_csv)

    log.info("Done → %s", out_xlsx)
    log.info("confident / review / no_match: %s %s %s", confident.shape, review.shape, no_match.shape)

if __name__ == "__main__":
    try:
        import pandas as pd  # noqa
    except Exception as e:
        log.error("Pandas missing? pip install -r requirements.txt")
        raise
    main()

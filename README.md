
# UNSPSC → CPV Matcher

Python nástroj pro mapování kódů **UNSPSC** na **CPV** s využitím tří signálů:
- **Semantic embeddings** (Sentence-Transformers) – hlavní a nejspolehlivější signál
- **TF‑IDF kosinová podobnost** – druhé vodítko
- **Fuzzy matching** (RapidFuzz) – fallback

Výstupem je Excel se záložkami `ALL`, `1_confident`, `2_review`, `3_no_match` a volitelně CSV mapping.

> Ověřeno na Pythonu 3.10+

---

## Vlastnosti
- **Priorita semantic skóre** (`score_sem`) při rozhodování
- Nastavitelné prahy pro semantic/TF‑IDF/fuzzy z příkazové řádky
- Volitelný překlad přes **DeepL** (pokud v `.env` existuje `DEEPL_API_KEY`)
- Čisté CLI, logging, žádné pevné cesty – vhodné pro GitHub/CI

---

## Instalace

Doporučeno v čistém virtuálním prostředí:
```bash
python -m venv .venv
# Windows PowerShell
. .\.venv\Scripts\Activate.ps1
# macOS/Linux
# source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
```

> **Pozn.:** Balíček `torch` (PyTorch) si sentence-transformers stáhne jako závislost. V případě potíží nainstalujte PyTorch pro svou platformu dle oficiálního návodu: https://pytorch.org/get-started/locally/

---

## Vstupní data

- `unspsc_input.csv` – CSV se **dvěma sloupci**:  
  1. `UNSPSC code` (8-místný, libovolné formátování, čísla se automaticky očistí)  
  2. `Short EN description` (zkrácený popis v angličtině)

- `cpv_2008_cz.xlsx` – CPV slovník (XLSX/XLS), obsahuje sloupce s kódem (`CODE` / `CPV CODE` / `KÓD`) a ideálně **EN** i **CS** názvy. Pokud EN chybí, nástroj páruje na **české** názvy (volitelně překládá UNSPSC EN→CS přes DeepL).

- `unspsc_full.xlsx` – oficiální UNSPSC (kód + **plný EN název**). První sloupec kód, druhý text – hlavičky jsou toleranční.

---

## Rychlý start

```bash
python unspsc_to_cpv.py \
  --input unspsc_input.csv \
  --cpv cpv_2008_cz.xlsx \
  --unspsc-full unspsc_full.xlsx \
  --out-xlsx unspsc_to_cpv.xlsx \
  --use-deepl
```

### Prahy a parametrizace
- Semantic (primární): `--sem-confident 85`, `--sem-review 70`, `--sem-margin 5`
- TF‑IDF (sekundární): `--tfidf-confident 88`, `--tfidf-review 75`
- Fuzzy (fallback): `--high 85`, `--low 70`

Příklad přísnější konfigurace:
```bash
python unspsc_to_cpv.py \
  --input unspsc_input.csv \
  --cpv cpv_2008_cz.xlsx \
  --unspsc-full unspsc_full.xlsx \
  --sem-confident 88 --sem-review 72 --sem-margin 6 \
  --tfidf-confident 90 --tfidf-review 78 \
  --out-xlsx unspsc_to_cpv.xlsx
```

---

## DeepL překlad (volitelně)
V kořenové složce projektu vytvořte soubor `.env`:
```
DEEPL_API_KEY=your_deepl_api_key_here
```
A spusťte skript s `--use-deepl`. Překlad se použije pro:
- Párování na **české** názvy CPV, pokud v CPV chybí EN
- `fallback_cs` u řádků `no_match` nebo tam, kde chybí `cpv_cs`

---

## Výstupy
- **Excel**: `unspsc_to_cpv.xlsx` se záložkami
  - `ALL` – kompletní výsledky se skóre a metodou (`semantic`/`tfidf`/`fuzz`)
  - `1_confident`, `2_review`, `3_no_match`
- **CSV**: `unspsc_to_cpv_mapping.csv` – zkrácená mapovací tabulka

Důležité sloupce:
- `score_sem`, `score_tfidf`, `score_fuzz`
- `method` (využitá metoda při rozhodnutí)
- `status` (`confident` / `review` / `no_match`)

---

## Tipy k výkonu
- První běh stáhne model (~90 MB)
- Parametr `--topk-tfidf-for-sem` (default 50) omezuje počet kandidátů pro reranking embeddingy – vyšší hodnota zvýší přesnost i nároky

---

## Troubleshooting
- **Torch install**: postupujte dle oficiální příručky pro vaši platformu (CUDA vs CPU)
- **Excel engine**: pro `.xls` soubory je potřeba `xlrd` (viz `requirements.txt`)
- **Paměť/rychlost**: snižte `--topk-tfidf-for-sem` nebo běžte po dávkách menších datasetů

---

## Licence
MIT 


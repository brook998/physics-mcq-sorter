# 🫡 Physics MCQ Sorter Web App

Upload up to **21 Physics question-paper PDFs at once** (or ZIP files containing up to 21 PDFs total). The app extracts English Section-A MCQs, classifies them into the 15 Class 12 Physics chapters, and produces chapter-wise DOCX files inside a ZIP.

## What it does

- Accepts PDF files directly or ZIP files containing PDFs.
- Maximum **21 PDF papers per processing run**.
- Extracts English MCQs and filters Devanagari/Hindi text.
- Uses OCR only as a fallback for scanned/image-only pages.
- Preserves equations, symbols, fractions, diagrams and options by placing a crop of the original PDF in the DOCX.
- Sorts MCQs into the 15-chapter NCERT Class 12 Physics structure.
- Creates `Physics_MCQ_Chapter_Wise.zip` plus a classification-review CSV.
- Does not intentionally persist uploaded papers in the app code.

## Speed / fidelity choices

The app first uses the PDF's native text layer. It avoids unnecessary OCR on normal text PDFs and uses a lighter output render scale to reduce processing time and output size while keeping equations readable.

Scanned papers can still use OCR by leaving **Use OCR fallback** enabled. OCR requires a Tesseract binary on the host.

## Run locally

Python 3.10+ is recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

For OCR on Debian/Ubuntu:

```bash
sudo apt-get update && sudo apt-get install -y tesseract-ocr
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository and upload the contents of this folder.
2. Open Streamlit Community Cloud and create an app from that GitHub repository.
3. Select `app.py` as the entry point and deploy.
4. Streamlit gives you an HTTPS `streamlit.app` address you can open on your phone.

### OCR note for cloud deployment

The normal text-PDF workflow does not depend on Tesseract. Some cloud hosts do not provide the Tesseract system binary automatically, so scanned/image-only PDFs may need a Docker deployment. The app fails gracefully rather than crashing when OCR is unavailable.

## Docker deployment

The included `Dockerfile` installs Tesseract so the OCR fallback works on Docker-capable hosts.

```bash
docker build -t physics-mcq-sorter .
docker run -p 8501:8501 physics-mcq-sorter
```

Then open `http://localhost:8501`.

## Classification

The classifier is offline and explainable. It scores chapter-specific Physics terms and reports a confidence value. Review lower-confidence rows in the CSV before treating the classification as authoritative.

## DOCX output formatting (v3)

- Hindi/Devanagari question regions are rejected at the question level, rather than only removing Hindi characters from the text.
- Only English MCQ regions with recognizable options are kept.
- The visible `Source: ...` line is removed from the DOCX. Source/page/question metadata remains in `classification_review.csv`.
- Question numbering is renumbered sequentially within each chapter (1, 2, 3, ...).
- The detected paper set code (for example `55/4/1`) is placed unobtrusively at the bottom-right of each MCQ image.
- Chapter keyword matching uses word boundaries to avoid false matches such as `led` being found inside words like `doubled`.


## v4.0 update
- Forces a fresh processor reload on app startup and displays the engine version.
- Removes Devanagari/Hindi text regions from bilingual question crops while preserving English content.
- Removes the old `Source:` line from generated DOCX files.
- Keeps continuous chapter-wise numbering and set code placement.


## v5.0 update
- Fixes the major bilingual-paper filtering bug: English characters anywhere in an MCQ no longer make it an English MCQ.
- Language classification is based primarily on the question/assertion/reason stem before the option list.
- Single Latin characters commonly used as physics variables (q, E, r, V, etc.) are ignored as English-language evidence.
- Hindi/Devanagari-dominant stems are rejected even when their options contain English words or (A)-(D).
- Mixed-language stems use conservative thresholds to avoid accidentally admitting Hindi questions.
- Duplicate same-number candidates are ranked by genuine English-stem signal rather than raw Latin-character count across the whole region.

from __future__ import annotations

import importlib
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import processor
importlib.reload(processor)
from processor import APP_ENGINE_VERSION, CHAPTERS, MAX_PDFS, collect_uploaded_files, extract_pdfs, make_bundle

st.set_page_config(page_title="Physics MCQ Sorter", page_icon="🫡", layout="wide")
st.title("🫡 Physics MCQ Sorter")
st.caption("Upload question papers → extract English MCQs → classify into 15 Class 12 Physics chapters → download chapter-wise DOCX files")
st.caption(f"Processing engine: {APP_ENGINE_VERSION}")

with st.expander("How it works", expanded=True):
    st.markdown(
        "**Fidelity first:** each extracted MCQ is rendered from the original PDF page and inserted as an image in the DOCX. "
        "That keeps physics symbols, fractions, equations and diagrams from being mangled by plain-text extraction. "
        "The chapter classifier runs locally using explainable physics keywords and provides a confidence score for review."
    )

uploads = st.file_uploader(
    f"Upload up to {MAX_PDFS} PDF papers at once, or a ZIP containing up to {MAX_PDFS} PDFs",
    type=["pdf", "zip"],
    accept_multiple_files=True,
    help=f"Maximum {MAX_PDFS} PDF papers per run. ZIP files are counted by the number of PDFs inside them.",
)
use_ocr = st.checkbox("Use OCR fallback for scanned/image-only PDFs", value=True,
                     help="OCR is slower. Keep this on for scanned papers; for normal text PDFs, the app avoids OCR automatically.")

if uploads and len(uploads) > MAX_PDFS:
    st.error(f"Please upload at most {MAX_PDFS} files at a time. ZIP contents are counted separately by PDF.")

if st.button("🚀 Process papers", type="primary", disabled=not uploads or len(uploads) > MAX_PDFS):
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        progress = st.progress(0, text="Preparing files…")
        status = st.empty()
        try:
            pdfs = collect_uploaded_files(uploads, workdir, max_pdfs=MAX_PDFS)
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        if not pdfs:
            st.error("No PDF files were found in the uploaded files.")
            st.stop()

        status.info(f"Found {len(pdfs)} PDF file(s). Processing…")
        progress.progress(5, text=f"Found {len(pdfs)} PDF file(s)…")

        try:
            def on_progress(done: int, total: int) -> None:
                pct = 5 + int(80 * done / max(1, total))
                progress.progress(min(pct, 85), text=f"Scanning paper {done}/{total}…")

            questions = extract_pdfs(pdfs, use_ocr=use_ocr, progress_callback=on_progress)
        except Exception as exc:
            progress.empty()
            status.empty()
            st.exception(exc)
            st.stop()

        if not questions:
            progress.empty()
            status.empty()
            st.warning("No MCQs were detected. Check that the papers contain readable Section-A MCQs numbered 1–16.")
            st.stop()

        progress.progress(92, text="Building chapter-wise DOCX files…")
        bundle, csv_bytes, df = make_bundle(questions)
        progress.progress(100, text="Finished")
        status.success(f"Done — detected {len(questions)} English MCQs from {len(pdfs)} paper(s).")

        counts = {n: int((df["chapter"] == n).sum()) for n, _ in CHAPTERS}
        c1, c2, c3 = st.columns(3)
        c1.metric("MCQs", len(questions))
        c2.metric("Chapters with MCQs", sum(v > 0 for v in counts.values()))
        c3.metric("Low confidence", int((df["confidence"] < 0.55).sum()))

        table = pd.DataFrame([{"Chapter": f"{n}. {name}", "MCQs": counts[n]} for n, name in CHAPTERS])
        st.dataframe(table, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download all chapter-wise DOCX files (ZIP)",
            data=bundle,
            file_name="Physics_MCQ_Chapter_Wise.zip",
            mime="application/zip",
            type="primary",
        )
        st.download_button(
            "⬇️ Download classification review CSV",
            data=csv_bytes,
            file_name="classification_review.csv",
            mime="text/csv",
        )

        with st.expander("Review extracted questions / confidence"):
            st.dataframe(df, use_container_width=True, hide_index=True)

st.divider()
st.caption("The app processes uploads in the active session and does not intentionally persist paper files or generated documents.")

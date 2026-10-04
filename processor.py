from __future__ import annotations

import io
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image, ImageOps, ImageDraw, ImageFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

APP_ENGINE_VERSION = "v4.0"
MAX_PDFS = 21
MAX_TOTAL_UPLOAD_MB = 500
IMAGE_SCALE = 2.0
OCR_SCALE = 2.0

CHAPTERS = [
    (1, "Electric Charges and Fields"),
    (2, "Electrostatic Potential and Capacitance"),
    (3, "Current Electricity"),
    (4, "Moving Charges and Magnetism"),
    (5, "Magnetism and Matter"),
    (6, "Electromagnetic Induction"),
    (7, "Alternating Current"),
    (8, "Electromagnetic Waves"),
    (9, "Ray Optics and Optical Instruments"),
    (10, "Wave Optics"),
    (11, "Dual Nature of Radiation and Matter"),
    (12, "Atoms"),
    (13, "Nuclei"),
    (14, "Semiconductor Electronics"),
    (15, "Communication Systems"),
]

KEYWORDS = {
    1: [("electric flux", 5), ("gauss law", 6), ("gaussian", 4), ("coulomb", 5), ("electric field", 2), ("field due to", 2), ("charge density", 3), ("line charge", 4), ("surface charge", 4), ("point charge", 3), ("electric dipole", 3), ("dipole moment", 3), ("superposition", 3)],
    2: [("capacitance", 7), ("capacitor", 5), ("dielectric", 5), ("electric potential", 4), ("potential difference", 4), ("equipotential", 5), ("potential energy", 3), ("combination of capacitors", 6), ("parallel plate", 3), ("stored energy", 3)],
    3: [("current electricity", 6), ("drift velocity", 6), ("drift speed", 5), ("resistivity", 5), ("resistance", 4), ("ohm's law", 6), ("ohm law", 6), ("kirchhoff", 6), ("wheatstone", 5), ("meter bridge", 5), ("potentiometer", 5), ("current density", 5), ("emf", 3), ("internal resistance", 5), ("cell", 1)],
    4: [("moving charges", 6), ("magnetic force", 5), ("lorentz", 6), ("biot-savart", 7), ("ampere's law", 6), ("ampere law", 6), ("cyclotron", 7), ("galvanometer", 6), ("moving coil", 5), ("force on a current", 5), ("charged particle", 2), ("magnetic field due to", 3), ("velocity selector", 6)],
    5: [("magnetism and matter", 8), ("magnetic material", 6), ("magnetic susceptibility", 7), ("diamagnetic", 7), ("paramagnetic", 7), ("ferromagnetic", 7), ("bar magnet", 6), ("earth's magnetism", 8), ("earth magnetism", 7), ("magnetic elements", 6), ("magnetic field intensity", 5), ("hysteresis", 6)],
    6: [("electromagnetic induction", 8), ("faraday", 7), ("lentz", 7), ("induced emf", 7), ("induced current", 6), ("self inductance", 7), ("mutual inductance", 7), ("inductance", 4), ("eddy current", 6), ("magnetic flux", 3), ("flux change", 4)],
    7: [("alternating current", 8), ("alternating voltage", 7), ("rms", 5), ("reactance", 6), ("impedance", 6), ("phasor", 6), ("transformer", 7), ("power factor", 6), ("resonance", 4), ("inductive reactance", 7), ("capacitive reactance", 7)],
    8: [("electromagnetic waves", 9), ("electromagnetic wave", 9), ("microwave", 7), ("radio wave", 7), ("infrared", 6), ("ultraviolet", 6), ("x-ray", 7), ("x ray", 7), ("gamma ray", 7), ("gamma rays", 7), ("spectrum", 2), ("wave propagation", 5)],
    9: [("ray optics", 8), ("geometrical optics", 8), ("mirror", 5), ("spherical mirror", 7), ("lens", 5), ("lens maker", 7), ("refraction", 6), ("refractive index", 7), ("total internal reflection", 8), ("prism", 6), ("optical instrument", 7), ("microscope", 7), ("telescope", 7), ("myopia", 7), ("hypermetropia", 7)],
    10: [("wave optics", 8), ("interference", 7), ("diffraction", 7), ("young's", 7), ("double slit", 7), ("ydse", 7), ("coherent", 5), ("fringe width", 7), ("polarisation", 7), ("polarization", 7), ("biprism", 6)],
    11: [("dual nature", 9), ("photoelectric", 8), ("photoelectric effect", 9), ("work function", 7), ("stopping potential", 8), ("de broglie", 8), ("matter wave", 7), ("threshold frequency", 7), ("photoelectron", 7)],
    12: [("bohr", 8), ("hydrogen spectrum", 8), ("spectral series", 6), ("energy level", 5), ("energy levels", 5), ("radius of orbit", 7), ("atomic spectra", 7), ("rydberg", 7), ("ground state", 4), ("excited state", 4)],
    13: [("nuclei", 3), ("nuclear", 4), ("radioactive", 8), ("radioactivity", 8), ("half life", 8), ("binding energy", 8), ("mass defect", 8), ("fission", 8), ("fusion", 8), ("nuclear force", 7), ("decay constant", 8)],
    14: [("semiconductor", 8), ("semiconductors", 8), ("p-n junction", 9), ("pn junction", 9), ("diode", 7), ("zener", 8), ("transistor", 8), ("rectifier", 8), ("logic gate", 8), ("logic gates", 8), ("junction diode", 8), ("digital electronics", 7), ("led", 6), ("photodiode", 7)],
    15: [("communication system", 9), ("communication systems", 9), ("modulation", 8), ("amplitude modulation", 9), ("bandwidth", 8), ("antenna", 8), ("propagation", 4), ("carrier wave", 8), ("communication", 2)],
}

Q_START_RE = re.compile(r"^\s*(?:Q(?:uestion)?\s*)?(1[0-6]|[1-9])\s*[\.)](?:\s+|$)", re.I)
OPTION_RE = re.compile(r"(?:\([A-Da-d]\)|\b[A-Da-d][\.)])")
SET_CODE_RE = re.compile(r"\b\d{2}/\d{1,2}/\d{1,2}\b")


def english_ratio(text: str) -> float:
    latin = len(re.findall(r"[A-Za-z]", text))
    dev = len(re.findall(r"[\u0900-\u097F]", text))
    return latin / (latin + dev) if latin + dev else 0.0


def devanagari_ratio(text: str) -> float:
    latin = len(re.findall(r"[A-Za-z]", text))
    dev = len(re.findall(r"[\u0900-\u097F]", text))
    return dev / (latin + dev) if latin + dev else 0.0


def _is_footer_line(line: str) -> bool:
    compact = re.sub(r"\s+", " ", line or "").strip()
    if not compact:
        return True
    if SET_CODE_RE.fullmatch(compact):
        return True
    if re.fullmatch(r"(?:P\.?\s*T\.?\s*O\.?|Page\s+\d+(?:\s+of\s+\d+)?)", compact, re.I):
        return True
    return False


def clean_english(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[\u0900-\u097F]", " ", text)
    lines = []
    for raw in text.splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if not line or _is_footer_line(line):
            continue
        if re.fullmatch(r"(?:page\s*)?\d+", line, re.I):
            continue
        if len(line) < 32 and re.search(r"(?:set|code|series)\s*[:\-]?\s*[A-Z0-9]+$", line, re.I):
            continue
        lines.append(line)
    return "\n".join(lines)


def chapter_classify(text: str) -> tuple[int, float, float]:
    t = re.sub(r"\s+", " ", text.lower())
    scores = {n: 0.0 for n, _ in CHAPTERS}
    for n, _ in CHAPTERS:
        for phrase, weight in KEYWORDS[n]:
            pattern = r"(?<![a-z0-9])" + re.escape(phrase) + r"(?![a-z0-9])"
            if re.search(pattern, t):
                scores[n] += weight
    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    best_n, best = ranked[0]
    second = ranked[1][1] if len(ranked) > 1 else 0.0
    confidence = 0.15 if not best else 0.25 + min(0.70, best / (best + 7.0))
    if best and best - second >= 5:
        confidence = min(0.98, confidence + 0.12)
    elif best and best - second <= 1:
        confidence = max(0.20, confidence - 0.10)
    return best_n, confidence, best - second


def _text_blocks(page) -> list[dict]:
    # Work line-by-line rather than block-by-block. Many PDFs put the entire
    # question section in one text block, while each MCQ number starts on its
    # own line. Keeping line geometry lets us find those starts reliably.
    lines_out = []
    raw = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT).get("blocks", [])
    for b in raw:
        if "lines" not in b:
            continue
        for line in b["lines"]:
            spans = line.get("spans", [])
            text = "".join(span.get("text", "") for span in spans).strip()
            if not text:
                continue
            xs0 = [span["bbox"][0] for span in spans]
            ys0 = [span["bbox"][1] for span in spans]
            xs1 = [span["bbox"][2] for span in spans]
            ys1 = [span["bbox"][3] for span in spans]
            lines_out.append({
                "x0": min(xs0), "y0": min(ys0), "x1": max(xs1), "y1": max(ys1),
                "text": text, "font_size": max((span.get("size", 11) for span in spans), default=11)
            })
    return sorted(lines_out, key=lambda z: (z["y0"], z["x0"]))


def _candidate_starts(blocks: list[dict]) -> list[dict]:
    starts = []
    for b in blocks:
        text = b["text"].strip()
        m = Q_START_RE.match(text)
        if not m:
            continue
        q = int(m.group(1))
        prefix = m.group(1) + (m.group(0)[-1] if m.group(0).strip() else ".")
        try:
            mask_width = float(fitz.get_text_length(prefix, fontname="helv", fontsize=float(b.get("font_size", 11)))) + 1.5
        except Exception:
            mask_width = max(18.0, 7.0 * len(prefix))
        starts.append({**b, "q": q, "number_mask_width": mask_width})
    return sorted(starts, key=lambda z: (z["y0"], z["x0"]))


def _region_text(blocks: list[dict], y0: float, y1: float) -> str:
    parts = []
    for b in blocks:
        if b["y1"] <= y0 or b["y0"] >= y1:
            continue
        if not _is_footer_line(b["text"]):
            parts.append(b["text"])
    return "\n".join(parts)


def _looks_like_english_mcq(region_text: str) -> bool:
    latin = len(re.findall(r"[A-Za-z]", region_text))
    dev = len(re.findall(r"[\u0900-\u097F]", region_text))
    if latin < 20:
        return False
    total = latin + dev
    er = latin / total if total else 0.0
    dr = dev / total if total else 0.0
    if er < 0.62 or dr > 0.18:
        return False
    # A true MCQ should have at least two recognizable option markers.
    options = OPTION_RE.findall(region_text)
    return len(options) >= 2


def _extract_set_code_from_text(text: str) -> str | None:
    m = SET_CODE_RE.search(text or "")
    return m.group(0) if m else None


def _detect_set_code(doc, use_ocr: bool = False) -> str | None:
    # Set code is normally printed in the page header/footer on every page.
    # Check a few early pages first and then the final pages, without OCR by default
    # to keep normal text-PDF processing fast.
    page_indexes = list(range(min(4, len(doc))))
    if len(doc) > 4:
        page_indexes += list(range(max(0, len(doc)-2), len(doc)))
    seen = set()
    for pno in page_indexes:
        if pno in seen:
            continue
        seen.add(pno)
        try:
            code = _extract_set_code_from_text(doc[pno].get_text("text") or "")
        except Exception:
            code = None
        if code:
            return code
    return None


def _ocr_page(page, scale: float = OCR_SCALE) -> tuple[list[dict], str]:
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    try:
        langs = pytesseract.get_languages(config="")
        ocr_lang = "eng+hin" if "hin" in langs and "eng" in langs else "eng"
    except Exception:
        ocr_lang = "eng"
    data = pytesseract.image_to_data(img, lang=ocr_lang, output_type=pytesseract.Output.DICT, config="--psm 6")
    lines = {}
    for i, txt in enumerate(data["text"]):
        txt = (txt or "").strip()
        if not txt:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append((data["left"][i], data["top"][i], data["width"][i], data["height"][i], txt))
    blocks, texts = [], []
    for words in lines.values():
        x0 = min(w[0] for w in words) / scale
        y0 = min(w[1] for w in words) / scale
        x1 = max(w[0] + w[2] for w in words) / scale
        y1 = max(w[1] + w[3] for w in words) / scale
        text = " ".join(w[4] for w in sorted(words, key=lambda z: z[0]))
        blocks.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": text})
        texts.append(text)
    return blocks, "\n".join(texts)


def _page_has_english(text: str) -> bool:
    return english_ratio(text) >= 0.45 and len(re.findall(r"[A-Za-z]", text)) >= 25


def _render_crop(page, y0: float, y1: float, scale: float = IMAGE_SCALE,
                 normalize_number: str | None = None, number_x: float | None = None,
                 number_y: float | None = None, set_code: str | None = None, hindi_boxes: list[tuple[float,float,float,float]] | None = None) -> bytes:
    rect = page.rect
    clip = fitz.Rect(0, max(0, y0), rect.width, min(rect.height, y1))
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    border = 8
    img = ImageOps.expand(img, border=border, fill="white")

    draw = ImageDraw.Draw(img)

    # Remove Hindi/Devanagari text from bilingual papers while preserving the
    # surrounding English question, options, symbols and diagrams.
    if hindi_boxes:
        for bx0, by0, bx1, by1 in hindi_boxes:
            hx0 = border + int(max(0, bx0) * scale)
            hy0 = border + int(max(0, (by0 - y0)) * scale)
            hx1 = border + int(max(0, bx1) * scale)
            hy1 = border + int(max(0, (by1 - y0)) * scale)
            draw.rectangle([hx0, hy0, hx1, hy1], fill="white")
    # Normalize the original paper question number so chapter-wise documents
    # always show one consistent 1, 2, 3, ... sequence.
    if normalize_number and number_x is not None and number_y is not None:
        x = border + max(0, number_x * scale)
        y = border + max(0, (number_y - y0) * scale)
        cover_w = int(42 * scale)
        cover_h = int(26 * scale)
        draw.rectangle([max(0, int(x-3)), max(0, int(y-4)),
                        min(img.width, int(x + cover_w)), min(img.height, int(y + cover_h))], fill="white")
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", max(18, int(11 * scale)))
        except Exception:
            font = ImageFont.load_default()
        draw.text((int(x), int(y)), normalize_number, fill="black", font=font)

    # Put the source set code where a mark/annotation is normally placed:
    # unobtrusively at the bottom-right of the MCQ image.
    if set_code:
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", max(14, int(8.5 * scale)))
        except Exception:
            font = ImageFont.load_default()
        label = set_code
        bbox = draw.textbbox((0, 0), label, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        pad = int(5 * scale)
        x = img.width - tw - pad
        y = img.height - th - pad
        draw.rectangle([x-pad//2, y-pad//2, img.width, img.height], fill="white")
        draw.text((x, y), label, fill="black", font=font)

    out = io.BytesIO()
    img.save(out, format="PNG", optimize=False, compress_level=3)
    return out.getvalue()


def extract_page_mcqs(page, paper: str, page_num: int, use_ocr: bool = True,
                      set_code: str | None = None, chapter_sequence_start: int = 1) -> list[dict]:
    blocks = _text_blocks(page)
    text = "\n".join(b["text"] for b in blocks)
    if use_ocr and (not _page_has_english(text) or len(blocks) < 4):
        try:
            ocr_blocks, ocr_text = _ocr_page(page)
            if ocr_blocks and _page_has_english(ocr_text):
                blocks, text = ocr_blocks, ocr_text
        except Exception:
            pass
    if not _page_has_english(text):
        return []

    starts = _candidate_starts(blocks)
    if not starts:
        return []

    accepted = []
    for i, s in enumerate(starts):
        # Region boundary is the next detected numbered question on the page.
        next_y = starts[i + 1]["y0"] if i + 1 < len(starts) else page.rect.height * 0.97
        region_text = _region_text(blocks, s["y0"] - 2, next_y - 2)
        if not _looks_like_english_mcq(region_text):
            continue
        # Prefer the English candidate if a paper contains Hindi+English versions
        # with the same question number. The region-level language check above
        # makes this deterministic.
        accepted.append((s, next_y, region_text))

    # Remove duplicate question-number starts only after language filtering.
    # If the same number appears twice, keep the candidate with the strongest
    # English ratio / most option markers.
    best = {}
    for s, next_y, region_text in accepted:
        latin = len(re.findall(r"[A-Za-z]", region_text))
        dev = len(re.findall(r"[\u0900-\u097F]", region_text))
        er = latin / (latin + dev) if latin + dev else 0.0
        opt_count = len(OPTION_RE.findall(region_text))
        score = (er, opt_count, len(region_text))
        prev = best.get(s["q"])
        if prev is None or score > prev[0]:
            best[s["q"]] = (score, s, next_y, region_text)

    results = []
    for s, next_y, region_text in sorted(((v[1], v[2], v[3]) for v in best.values()), key=lambda z: z[0]["y0"]):
        y0 = max(0, s["y0"] - 5)
        # Keep the crop tightly bounded to the accepted question region.
        if next_y < page.rect.height * 0.97:
            y1 = max(s["y1"] + 10, next_y - 7)
        else:
            region_blocks = [b for b in blocks if b["y1"] > s["y0"] and not _is_footer_line(b["text"])]
            max_y = max((b["y1"] for b in region_blocks), default=page.rect.height * 0.95)
            y1 = min(page.rect.height * 0.97, max_y + 10)

        q_lines = []
        for b in blocks:
            if b["y1"] <= y0 or b["y0"] >= y1 or _is_footer_line(b["text"]):
                continue
            cleaned = clean_english(b["text"])
            if cleaned:
                q_lines.append(cleaned)
        q_text = clean_english("\n".join(q_lines))
        if not q_text:
            continue
        chapter, conf, gap = chapter_classify(q_text)
        results.append({
            "paper": paper,
            "page": page_num,
            "number": s["q"],
            "text": q_text,
            "image": _render_crop(
                page, y0, y1,
                normalize_number=None,  # assigned in make_docx after chapter sorting
                number_x=s["x0"], number_y=s["y0"], set_code=set_code,
                hindi_boxes=[(b["x0"] - 1, b["y0"] - 1, b["x1"] + 1, b["y1"] + 1)
                             for b in blocks
                             if b["y1"] > y0 and b["y0"] < y1
                             and re.search(r"[\u0900-\u097F]", b["text"])
                             and not _is_footer_line(b["text"])]
            ),
            "chapter": chapter,
            "confidence": conf,
            "gap": gap,
            "number_x": s["x0"],
            "number_y": s["y0"],
            "number_mask_width": s.get("number_mask_width", 24.0),
            "crop_y0": y0,
            "crop_y1": y1,
            "set_code": set_code or "",
        })
    return results


def extract_pdfs(
    input_paths: Iterable[Path],
    use_ocr: bool = True,
    progress_callback: Callable[[int, int], None] | None = None,
) -> list[dict]:
    paths = list(input_paths)
    found = []
    total = len(paths)
    for i, pdf_path in enumerate(paths, start=1):
        doc = fitz.open(pdf_path)
        set_code = _detect_set_code(doc, use_ocr=False)
        try:
            for pno, page in enumerate(doc, start=1):
                # Cheap text-only gate first. Most question-paper pages contain native text;
                # this avoids expensive OCR/image work on obvious cover/footer pages when possible.
                plain = page.get_text("text") or ""
                likely_mcq_text = bool(re.search(r"\b(?:Q(?:uestion)?\s*)?(?:1[0-6]|[1-9])\s*[\.)]", plain, re.I))
                if not likely_mcq_text:
                    # Native-text page with no MCQ numbering: skip immediately.
                    # Empty-text pages are considered for OCR only when they actually contain images,
                    # which avoids OCR work on blank pages.
                    if plain.strip() or not page.get_images(full=True):
                        continue
                found.extend(extract_page_mcqs(page, pdf_path.name, pno, use_ocr=use_ocr, set_code=set_code))
        finally:
            doc.close()
        if progress_callback:
            progress_callback(i, total)

    # Deduplicate exact paper/page/question repetitions while keeping the richest crop/text.
    unique = {}
    for q in found:
        key = (q["paper"], q["page"], q["number"])
        prev = unique.get(key)
        if prev is None or len(q["text"]) > len(prev["text"]):
            unique[key] = q
    return list(unique.values())


def make_docx(chapter_num: int, chapter_name: str, questions: list[dict]) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.55)
    sec.right_margin = Inches(0.55)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run(f"Class 12 Physics — Chapter {chapter_num}: {chapter_name}")
    r.bold = True
    r.font.size = Pt(15)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run(f"English MCQs • {len(questions)} questions")
    sr.italic = True
    sr.font.size = Pt(9)

    for idx, q in enumerate(questions, 1):
        # Re-render the crop with uniform chapter-wise numbering. The original
        # paper number is masked and replaced; source metadata remains available
        # in the review CSV instead of cluttering the DOCX.
        image_bytes = q["image"]
        if q.get("number_x") is not None and q.get("number_y") is not None:
            try:
                # We need the original page to mask the number, so if image_bytes
                # is already rendered, keep it unchanged. The current crop builder
                # has already removed the footer and placed the set code.
                img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
                draw = ImageDraw.Draw(img)
                # Mask the original paper question number at its measured
                # position, then write the new chapter-wise sequence there.
                scale = IMAGE_SCALE
                border = 8
                x = border + float(q.get("number_x", 0)) * scale
                y = border + (float(q.get("number_y", q.get("crop_y0", 0))) - float(q.get("crop_y0", 0))) * scale
                cover_w = int(float(q.get("number_mask_width", 24.0)) * scale)
                cover_h = int(max(18.0, float(q.get("number_mask_width", 24.0)) * 0.75) * scale)
                x0 = max(0, int(x - 3))
                y0 = max(0, int(y - 4))
                draw.rectangle([x0, y0, min(img.width, int(x + cover_w)), min(img.height, int(y + cover_h))], fill="white")
                try:
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", max(20, int(11 * scale)))
                except Exception:
                    font = ImageFont.load_default()
                draw.text((int(x), int(y)), f"{idx}.", fill="black", font=font)
                out = io.BytesIO(); img.save(out, format="PNG", optimize=False, compress_level=3); image_bytes = out.getvalue()
            except Exception:
                pass

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.keep_together = True
        p.paragraph_format.keep_with_next = False
        p.add_run().add_picture(io.BytesIO(image_bytes), width=Inches(6.25))

    doc.core_properties.title = f"Physics Chapter {chapter_num} MCQs"
    doc.core_properties.subject = "English MCQs sorted by NCERT chapter"
    doc.core_properties.author = "Physics MCQ Sorter"
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def make_bundle(questions: list[dict]) -> tuple[bytes, bytes, pd.DataFrame]:
    rows = [{
        "paper": q["paper"], "page": q["page"], "question": q["number"],
        "chapter": q["chapter"], "chapter_name": dict(CHAPTERS)[q["chapter"]],
        "confidence": round(q["confidence"], 3), "score_gap": round(q["gap"], 2),
    } for q in questions]
    df = pd.DataFrame(rows, columns=["paper", "page", "question", "chapter", "chapter_name", "confidence", "score_gap"])
    if not df.empty:
        df = df.sort_values(["chapter", "paper", "page", "question"]).reset_index(drop=True)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        manifest = {
            "total_mcqs": len(questions),
            "chapters": {str(n): {"name": name, "count": int((df["chapter"] == n).sum()) if not df.empty else 0} for n, name in CHAPTERS},
        }
        z.writestr("manifest.json", json.dumps(manifest, indent=2))
        z.writestr("classification_review.csv", df.to_csv(index=False))
        for n, name in CHAPTERS:
            qs = [q for q in questions if q["chapter"] == n]
            fname = f"Chapter_{n:02d}_{re.sub(r'[^A-Za-z0-9]+', '_', name).strip('_')}.docx"
            z.writestr(fname, make_docx(n, name, qs))
    return zip_buf.getvalue(), df.to_csv(index=False).encode("utf-8"), df


def collect_uploaded_files(uploaded_files, workdir: Path, max_pdfs: int = MAX_PDFS) -> list[Path]:
    pdfs = []
    total_bytes = 0

    def add_pdf_bytes(filename: str, data: bytes) -> None:
        nonlocal total_bytes
        if len(pdfs) >= max_pdfs:
            raise ValueError(f"Upload limit reached: at most {max_pdfs} PDF papers can be processed per run.")
        total_bytes += len(data)
        if total_bytes > MAX_TOTAL_UPLOAD_MB * 1024 * 1024:
            raise ValueError(f"Total uploaded paper data is too large. Keep the run under {MAX_TOTAL_UPLOAD_MB} MB.")

        safe_name = Path(filename).name or f"paper_{len(pdfs)+1}.pdf"
        out = workdir / safe_name
        if out.exists():
            stem, suffix = out.stem, out.suffix
            k = 2
            while (workdir / f"{stem}_{k}{suffix}").exists():
                k += 1
            out = workdir / f"{stem}_{k}{suffix}"
        out.write_bytes(data)
        pdfs.append(out)

    for uf in uploaded_files:
        name = Path(uf.name).name
        data = uf.getvalue()
        if name.lower().endswith(".pdf"):
            add_pdf_bytes(name, data)
        elif name.lower().endswith(".zip"):
            total_bytes += len(data)
            if total_bytes > MAX_TOTAL_UPLOAD_MB * 1024 * 1024:
                raise ValueError(f"Total uploaded paper data is too large. Keep the run under {MAX_TOTAL_UPLOAD_MB} MB.")
            # Read ZIP from memory so the app does not need an extra extraction tree.
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for member in z.infolist():
                    if member.is_dir() or not Path(member.filename).name.lower().endswith(".pdf"):
                        continue
                    add_pdf_bytes(Path(member.filename).name, z.read(member))
    return pdfs


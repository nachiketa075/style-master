"""
Extracts header / left-panel metadata fields from a jewelry CAD sheet:
Company, Date, Design Code, Product Type, Total Components,
Component Details, Two Tone, Metal Details, Designer.

Works purely from reconstructed text LINES (word boxes grouped by
vertical position) plus their pixel position on the page, so it does
not depend on any single fixed crop region.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import numpy as np

import config
from ocr_engine import (
    Word,
    group_words_into_lines,
    line_text,
    ocr_cell,
    ocr_region_lines,
    ocr_region_lines_with_bounds,
    ocr_words,
)

_DESIGN_CODE_RE = re.compile(r"\b([A-Za-z0-9][A-Za-z0-9\-]{3,})\s*\(\s*([A-Za-z0-9#\-]{3,})\s*\)")
_DATE_RE = re.compile(r"Date\s*[:\-]?\s*([0-9]{1,2}[\-/][A-Za-z]{3,9}[\-/][0-9]{2,4})", re.IGNORECASE)
_TOTAL_COMPONENTS_RE = re.compile(r"Total\s*Number\s*of\s*Components\s*=?\s*(\d+)", re.IGNORECASE)
_TWO_TONE_RE = re.compile(r"Two\s*tone\s*=\s*([A-Za-z]+)", re.IGNORECASE)
_METAL_LABEL_RE = re.compile(r"Metal\s*details", re.IGNORECASE)
_COMPONENT_LINE_RE = re.compile(r"^([A-Za-z][A-Za-z\s]*?)[\.\s]{2,}\s*(\d+)\s*$")
_COMPONENT_LINE_RE_TIGHT = re.compile(r"^([A-Za-z][A-Za-z\s]*?)\s*[\.]{2,}\s*(\d+)\s*$")
# Dot-leaders (e.g. "Bell..............1") were found to sometimes OCR
# with a stray "0"/"o"/"O" mixed into the leader itself (a dot
# misread as a zero), which the two patterns above - deliberately
# strict about "only dots/spaces" - correctly refuse to match, per
# Rule 26 (never guess). This third, looser pattern is the intentional
# fallback for exactly that case: it still requires a real letter-only
# name and a real trailing digit run, just tolerates 0/o/O as leader
# filler alongside dots.
_COMPONENT_LINE_RE_LOOSE = re.compile(r"^([A-Za-z][A-Za-z\s]*?)[\.\s0oO]{2,}(\d+)\s*$")


@dataclass
class SheetMetadata:
    company: str = ""
    date: str = ""
    design_code: str = ""
    design_code_2: str = ""
    product_type: str = ""
    total_components: str = ""
    component_details: str = ""
    extra_components: str = ""
    two_tone: str = ""
    designer: str = ""
    components: list[tuple[str, str]] = field(default_factory=list)
    # One entry per logical Metal Details statement (e.g. "Target 10kt
    # Wt =0.34gms" and "Achieve 10kt Wt =0.34gms (Without Chain)" are
    # TWO separate entries) - main.py generates one Excel row per
    # entry, repeating every other field. Never combined into a single
    # "|"-joined string: that would merge distinct business records.
    metal_details_list: list[str] = field(default_factory=list)


def _clean(s: str) -> str:
    # Curly braces never legitimately appear in this data; Tesseract
    # occasionally misreads a round bracket as one (a known, narrowly-
    # scoped glyph confusion, same reasoning as the "|"/"1" fix in
    # table_extractor.py - not a guess at new content).
    s = s.replace("{", "(").replace("}", ")")
    return re.sub(r"\s+", " ", s).strip()


def extract_two_tone(text: str) -> str:
    """Dedicated Two Tone extractor (case-insensitive). Matches
    'Two tone=<value>' / 'Two Tone = <value>' / '*Two tone=<value>'
    etc.; the '*' label marker is never part of the value. Common
    Yes/No variants are normalized to title case; any other legitimate
    value is preserved exactly as OCR'd rather than guessed at.
    Returns "" if no Two Tone label is present in `text` - a sheet
    that genuinely has no Two Tone field must not get one invented."""
    m = _TWO_TONE_RE.search(text)
    if not m:
        return ""
    value = _clean(m.group(1))
    lowered = value.lower()
    if lowered == "no":
        return "No"
    if lowered == "yes":
        return "Yes"
    return value


def _split_metal_detail_entries(block_text: str) -> list[str]:
    """Split a Metal Details text block into separate logical
    statements. A line that contains '=' (e.g. "Target 10kt Wt
    =0.34gms") starts a NEW entry; a line without '=' (e.g.
    "(Without Chain)") is a wrapped CONTINUATION of the previous
    entry, not a new one - so a single statement that wraps across
    visual lines stays one record, while genuinely distinct
    statements each become their own."""
    entries: list[str] = []
    current: list[str] = []
    for raw_line in block_text.splitlines():
        stripped = raw_line.strip().strip("|").strip()
        if not stripped or stripped.startswith("*"):
            continue
        if "=" in stripped:
            if current:
                entries.append(_clean(" ".join(current)))
            current = [stripped]
        elif current:
            current.append(stripped)
        else:
            current = [stripped]
    if current:
        entries.append(_clean(" ".join(current)))
    return entries


def _recover_component_name(crop_bgr: np.ndarray, band_y0: int, band_y1: int) -> str:
    """Word-level fallback for a component-list line whose single-
    line OCR pass (psm 7) produced nonsense with no usable name/qty
    match at all (seen on a real sample: a clean "pri..... 2" line
    read back as "ere"). Re-OCR'ing the same band word-by-word (psm
    11) - the same isolate-and-re-OCR strategy already used for table
    cells and the design-code line elsewhere in this codebase -
    recovers at least the item name even when the line's single-line
    pass fails outright. Only letters-only words above the same
    confidence bar already reserved for critical fields are trusted,
    so a stray low-confidence hallucination can't be mistaken for a
    real component name."""
    band = crop_bgr[band_y0:band_y1, :]
    if band.size == 0:
        return ""
    words = ocr_words(band, psm=11)
    name_tokens = [
        w.text
        for w in words
        if w.conf >= config.MIN_CONFIDENCE_FOR_CRITICAL_FIELDS and re.fullmatch(r"[A-Za-z]+", w.text)
    ]
    if not name_tokens:
        return ""
    return _clean(" ".join(name_tokens))


def _backfill_component_quantities(meta: "SheetMetadata", pending_names: list[str]) -> None:
    """A component line whose OCR recovered a real item name (via
    _recover_component_name) but no usable quantity digit is only
    ever filled in when the sheet's own printed Total Number of
    Components unambiguously determines it - i.e. exactly one
    component on the whole sheet is missing its quantity, and the
    remaining amount (the printed total minus every quantity that WAS
    read cleanly) is a sensible positive number. This is arithmetic
    derived from a value already trusted elsewhere in this file
    (never guess independently of the printed total), not a guess at
    what the missing digit might have been - if more than one line is
    ambiguous, or the numbers don't cleanly reconcile, nothing is
    added rather than risk fabricating a business record."""
    if len(pending_names) != 1 or not meta.total_components.strip().isdigit():
        return
    try:
        known_sum = sum(int(q) for _n, q in meta.components)
    except ValueError:
        return
    remaining = int(meta.total_components) - known_sum
    if remaining > 0:
        meta.components.append((pending_names[0], str(remaining)))


def extract_metadata(
    img_bgr: np.ndarray,
    words: list[Word],
    image_width: int,
    image_height: int,
    table_left_boundary: float | None = None,
) -> SheetMetadata:
    # Company/Date/Design Code can legitimately appear anywhere on the
    # page (top header band, or centered/right at the bottom), so they
    # are matched against ALL reconstructed lines.
    lines = group_words_into_lines(words)
    line_strings = [line_text(l) for l in lines]

    # Total Components / component list / Two Tone / Metal details /
    # Designer live in the LEFT info panel only. If we know where the
    # table starts (table_left_boundary), restrict these fields to
    # words left of it so same-row table/table-title text can never
    # bleed into them (mirrors the fix in table_extractor.py).
    if table_left_boundary is not None:
        left_words = [w for w in words if w.cx < table_left_boundary]
        left_lines = group_words_into_lines(left_words)
        left_line_strings = [line_text(l) for l in left_lines]
    else:
        left_lines, left_line_strings = lines, line_strings

    meta = SheetMetadata()

    # ---- Company + Date: both live in the top header band ----------
    top_band = [
        (i, l, s)
        for i, (l, s) in enumerate(zip(lines, line_strings))
        if l and (l[0].top / max(image_height, 1)) < 0.08
    ]
    date_line_idx = None
    date_line_str = ""
    for i, l, s in top_band:
        m = _DATE_RE.search(s)
        if m:
            meta.date = _clean(m.group(1))
            date_line_idx = i
            date_line_str = s
            break
    if not meta.date:
        # date label may sit outside the strict 8% band on tall sheets
        for i, s in enumerate(line_strings):
            m = _DATE_RE.search(s)
            if m:
                meta.date = _clean(m.group(1))
                date_line_idx = i
                date_line_str = s
                break

    # Company and "Date :" can end up reconstructed as ONE line when
    # they sit on the same visual row far apart in X (header spans the
    # full page width). Strip the "Date ..." portion off first so it
    # never contaminates the company name.
    if date_line_str:
        label_m = re.search(r"\bDate\b", date_line_str, re.IGNORECASE)
        if label_m:
            prefix = date_line_str[: label_m.start()].strip()
            if prefix:
                meta.company = _clean(prefix)

    if not meta.company:
        company_candidates = [(i, l, s) for i, l, s in top_band if i != date_line_idx and s.strip()]
        if company_candidates:
            company_candidates.sort(key=lambda t: t[1][0].left)
            meta.company = _clean(company_candidates[0][2])

    # ---- Design code: pattern CODE(CODE2) ---------------------------
    # Design codes are small, dense strings (Rule 5: must be preserved
    # exactly) where the page-wide OCR pass is prone to character-level
    # noise like O/0 confusion. Re-OCR each candidate line in isolation
    # first - proven far more reliable than the tiled full-page text,
    # same as the component-list fix above - and only fall back to the
    # noisy full-page text if isolated re-OCR finds nothing.
    code_counts: dict[str, int] = {}
    candidate_lines = [l for l, s in zip(lines, line_strings) if "(" in s]
    for l in candidate_lines:
        # A near-zero-confidence word anywhere on the same visual row
        # (e.g. a hallucinated OCR artifact off near a page border,
        # far from the real code) would otherwise stretch this line's
        # bounding box out to include it, and re-OCR'ing that much
        # wider/mostly-blank crop was observed to hallucinate an
        # entirely fabricated prefix onto the real code (e.g. a real
        # "MBS061780G-E(REVN1087)" read back as "KNODIL
        # MBS061780G-E(REVN1087)"). Restrict the bbox to words that
        # meet the same confidence bar already reserved in config for
        # critical code/number fields; fall back to the full line only
        # if every word on it happens to be low-confidence.
        confident_words = [w for w in l if w.conf >= config.MIN_CONFIDENCE_FOR_CRITICAL_FIELDS] or l
        top = max(min(w.top for w in confident_words) - 4, 0)
        bottom = min(max(w.bottom for w in confident_words) + 4, image_height)
        left = max(min(w.left for w in confident_words) - 4, 0)
        right = min(max(w.right for w in confident_words) + 4, image_width)
        if bottom <= top or right <= left:
            continue
        clean_text = ocr_cell(img_bgr[top:bottom, left:right], psm=7)
        for m in _DESIGN_CODE_RE.finditer(clean_text.replace(" ", "")):
            code = f"{m.group(1)}({m.group(2)})"
            code_counts[code] = code_counts.get(code, 0) + 1

    if not code_counts:
        for s in line_strings:
            for m in _DESIGN_CODE_RE.finditer(s.replace(" ", "")):
                code = f"{m.group(1)}({m.group(2)})"
                code_counts[code] = code_counts.get(code, 0) + 1

    if code_counts:
        best_code = max(code_counts.items(), key=lambda kv: kv[1])[0]
        meta.design_code = best_code.split("(", 1)[0]
        meta.design_code_2 = best_code.split("(", 1)[1][:-1]

    # ---- Total Number of Components ---------------------------------
    tc_idx = None
    for i, s in enumerate(left_line_strings):
        m = _TOTAL_COMPONENTS_RE.search(s)
        if m:
            meta.total_components = m.group(1)
            tc_idx = i
            break

    # ---- Metal details label position (needed as a boundary below) -
    metal_idx = None
    for i, s in enumerate(left_line_strings):
        if _METAL_LABEL_RE.search(s):
            metal_idx = i
            break

    crop_x1 = int(table_left_boundary) if table_left_boundary is not None else int(image_width * 0.22)
    crop_x1 = min(crop_x1, image_width)

    # ---- Component list + Two Tone ----------------------------------
    # Both live in the gap between "Total Number of Components" and
    # "Metal details" (e.g. "Bracelet..... 1" then "*Two tone=No").
    # This is small, faint text (dot-leaders, and on one real sample a
    # cyan/low-contrast color) that the page-wide OCR pass regularly
    # drops or garbles - re-OCR the whole gap region in isolation in
    # ONE pass, proven far more reliable than reusing the already
    # -tiled full-page words here.
    if tc_idx is not None:
        tc_line = left_lines[tc_idx]
        region_y0 = max(w.bottom for w in tc_line) + 2
        if metal_idx is not None:
            region_y1 = max(min(w.top for w in left_lines[metal_idx]) - 2, region_y0)
        else:
            region_y1 = min(region_y0 + 250, image_height)
        region_y1 = min(region_y1, image_height)

        block_text = ""
        lines_with_bounds: list[tuple[str, int, int]] = []
        crop = None
        if region_y1 > region_y0 and crop_x1 > 0:
            crop = img_bgr[region_y0:region_y1, 0:crop_x1]
            lines_with_bounds = ocr_region_lines_with_bounds(crop)
            block_text = "\n".join(text for text, _y0, _y1 in lines_with_bounds if text)

        pending_name_only: list[str] = []
        for raw_line, band_y0, band_y1 in lines_with_bounds:
            stripped = raw_line.strip().strip("|").strip()
            if not stripped or stripped.startswith("*"):
                continue
            m = (
                _COMPONENT_LINE_RE_TIGHT.match(stripped)
                or _COMPONENT_LINE_RE.match(stripped)
                or _COMPONENT_LINE_RE_LOOSE.match(stripped)
            )
            if m:
                name = _clean(m.group(1)).rstrip(".")
                qty = m.group(2)
                meta.components.append((name, qty))
                continue

            # This line's own single-line OCR pass came back as
            # nonsense (no name/qty match at all) - try to at least
            # recover the item name via word-level re-OCR before
            # giving up on the line entirely.
            if crop is not None:
                recovered_name = _recover_component_name(crop, band_y0, band_y1)
                if recovered_name:
                    pending_name_only.append(recovered_name)

        if pending_name_only:
            _backfill_component_quantities(meta, pending_name_only)

        meta.two_tone = extract_two_tone(block_text)

    if meta.components:
        first_n, first_q = meta.components[0]
        meta.component_details = f"{first_n}: {first_q}"
        meta.extra_components = "; ".join(f"{n}: {q}" for n, q in meta.components[1:])

    # ---- Product type (derived from components / known vocabulary) -
    meta.product_type = _derive_product_type(meta.components)

    # ---- Metal details: one logical statement = one list entry -----
    # (Rule: "Target 10kt Wt =0.34gms" and "Achieve 10kt Wt =0.34gms
    # (Without Chain)" are TWO separate records, never joined into one
    # string - main.py generates one Excel row per entry.)
    if metal_idx is not None:
        region_y0 = max(w.bottom for w in left_lines[metal_idx]) + 2
        # Bound below the Designer name's own box (Rule: don't sweep
        # it into Metal Details) without needing that box's exact
        # position yet - it always sits in the bottom ~6% of the page.
        region_y1 = min(region_y0 + 260, image_height, int(image_height * 0.94))

        block_text = ""
        if region_y1 > region_y0 and crop_x1 > 0:
            crop = img_bgr[region_y0:region_y1, 0:crop_x1]
            block_text = "\n".join(ocr_region_lines(crop))

        meta.metal_details_list = _split_metal_detail_entries(block_text)

    # ---- Designer: short line, bottom-left corner of the sheet -----
    bottom_left_candidates = [
        (l, s)
        for l, s in zip(left_lines, left_line_strings)
        if l
        and (l[0].top / max(image_height, 1)) > 0.90
        and (l[0].left / max(image_width, 1)) < 0.20
        and not _DESIGN_CODE_RE.search(s.replace(" ", ""))
        and len(s.strip()) > 0
    ]
    if bottom_left_candidates:
        bottom_left_candidates.sort(key=lambda t: -t[0][0].top)
        meta.designer = _clean(bottom_left_candidates[0][1])

    return meta


def _derive_product_type(components: list[tuple[str, str]]) -> str:
    if not components:
        return ""
    name = components[0][0].strip().lower()
    # Exact match wins outright, before any substring comparison: "ring"
    # is a literal substring of "earring"/"earrings", so without this an
    # exact "Ring" reading would otherwise be overridden below by the
    # longer word that merely contains it.
    for known in config.KNOWN_COMPONENT_TYPES:
        if known.lower() == name:
            return known
    # Word-boundary match (not raw substring containment) against the
    # rest of the name, e.g. "Diamond Ring" or OCR noise like "Rings":
    # a plain substring test would let "ring" match inside
    # "earring"/"earrings" (and "rings" match inside "earrings", since
    # "earrings" literally contains "rings"), misclassifying a Ring
    # sheet as Earrings. \b requires the known word to appear as its
    # own token, which "ring" never does inside "earring"/"earrings" -
    # there's no boundary between "ear" and "ring" in that token.
    best: str | None = None
    for known in config.KNOWN_COMPONENT_TYPES:
        k = known.lower().rstrip("s")
        if re.search(rf"\b{re.escape(k)}s?\b", name):
            if best is None or len(k) > len(best.lower().rstrip("s")):
                best = known
    if best is not None:
        return best
    return components[0][0].strip().title()

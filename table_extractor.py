"""
Extracts the stone/component table (Location, Shape, MM Size, Sieve
Size, Per Stone Wt, Pcs, Total Wt, Setting Type, Setting) from a
jewelry CAD sheet.

Primary strategy (Rule 16): detect the table's actual grid lines with
OpenCV morphology, then OCR each cell in isolation. Testing against
real sample sheets showed that OCR'ing a whole row (or a multi-row
strip) at once regularly drops or garbles cells even at high
resolution, while the exact same text OCR'd as an isolated single-cell
crop reads cleanly - so per-cell OCR against the real grid is what
actually gets the accuracy the spec requires.

Fallback strategy (Rule 16's explicit fallback): if grid-line
detection doesn't confidently find the table (thin/broken borders,
skewed scan, etc.), fall back to fuzzy-matching the header labels
against reconstructed OCR text lines and assigning words to their
nearest column-anchor X position.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher

import cv2
import numpy as np

import config
from ocr_engine import Word, group_words_into_lines, line_text, ocr_cell

_TOTAL_MARKER = "total"

# Columns that are always purely numeric on these sheets. Tesseract
# frequently misreads a lone digit "1" as a pipe/I/l in narrow numeric
# cells; normalizing ONLY a single stray character in these numeric
# columns is a known, well-scoped OCR-glyph fix, not a data guess.
_NUMERIC_COLUMNS = {"MM Size", "Pcs", "Per Stone Wt", "Total Wt"}
_AMBIGUOUS_ONE_RE = re.compile(r"^[|Il]$")

# Columns that only ever hold a plain digit run (optionally with a
# decimal point) - never "x" pairs like MM Size can. Used to scope the
# digit-whitelist re-OCR recovery below.
_DIGITS_ONLY_COLUMNS = {"Pcs", "Per Stone Wt", "Total Wt"}
_HAS_DIGIT_RE = re.compile(r"[0-9]")


def _is_zero_shaped(crop_bgr: np.ndarray) -> bool:
    """Direct topological evidence that a single glyph is "0": a "0"
    is a closed loop, so its binarized blob has exactly one enclosed
    hole occupying a substantial share of the glyph's area - unlike
    "i"/"I"/"l"/"1", which have no enclosed hole at all. Bundled
    tessdata here is LSTM-only, so a digit-whitelist re-OCR (tested
    directly) cannot force a legacy-engine classification and returns
    empty or a spurious leading period instead of "0"; shape evidence
    from the pixels themselves is what actually resolves this
    misread, the same evidence-over-guessing approach already used for
    MM Size decimal recovery."""
    if crop_bgr.size == 0:
        return False
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    up = cv2.resize(gray, None, fx=6, fy=6, interpolation=cv2.INTER_LANCZOS4)
    _, binary = cv2.threshold(up, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    contours, hierarchy = cv2.findContours(binary, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    if hierarchy is None:
        return False
    areas = [cv2.contourArea(c) for c in contours]
    for i, area in enumerate(areas):
        if area < 200:
            continue
        # A child contour (hierarchy[i][2] != -1) is a hole inside
        # this blob. Require it to be a real enclosed loop (>15% of
        # the parent's area), not stray noise inside the glyph.
        child_idx = hierarchy[0][i][2]
        if child_idx != -1 and areas[child_idx] > area * 0.15:
            return True
    return False


def _count_full_glyphs(crop_bgr: np.ndarray) -> int | None:
    """Count the digit-sized ink blobs in a cell crop via connected-
    component analysis, excluding small decimal-point-shaped blobs
    (same short/squarish shape test as _recover_missing_decimals).
    This is direct pixel evidence for how many characters a digits-only
    cell actually contains, independent of whatever OCR returns for it.
    Returns None when the crop has no usable ink to count."""
    if crop_bgr.size == 0:
        return None
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    up = cv2.resize(gray, None, fx=8, fy=8, interpolation=cv2.INTER_LANCZOS4)
    _, binary = cv2.threshold(up, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, _labels, stats, _centroids = cv2.connectedComponentsWithStats(binary, connectivity=8)
    comps = [
        (stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT])
        for i in range(1, n)
        if stats[i, cv2.CC_STAT_AREA] >= 15
    ]
    if not comps:
        return None
    max_h = max(h for _w, h in comps)
    return sum(1 for w, h in comps if not (h < max_h * 0.35 and w < max_h * 0.5))


# Digits-only columns are schema-guaranteed to hold nothing but digits
# and an optional decimal point, so a character whitelist is safe here
# (unlike MM Size, which can legitimately contain "x"). Tested directly
# against real misread cells: whitelisting alone fixes single-glyph
# letter-for-digit misreads (e.g. "4" read as "A"), but multi-digit
# cells sometimes still come out wrong under PSM 7 (single line) while
# reading correctly under PSM 8 (single word), and vice versa for
# single-glyph cells - so both are tried and the pixel-evidence glyph
# count (_count_full_glyphs) picks whichever candidate actually has the
# right number of digits, rather than trusting either PSM blindly.
_DIGIT_WHITELIST_CONFIG = "-c tessedit_char_whitelist=0123456789."


def _recover_digit_cell(crop_bgr: np.ndarray, label: str, text: str) -> str:
    if label not in _DIGITS_ONLY_COLUMNS:
        return text

    expected = _count_full_glyphs(crop_bgr)
    if not expected:
        return text

    for psm in (7, 8):
        candidate = ocr_cell(crop_bgr, psm=psm, extra_config=_DIGIT_WHITELIST_CONFIG).strip()
        if sum(1 for c in candidate if c.isdigit()) == expected:
            return candidate

    # Neither whitelisted candidate resolved it. A single glyph with NO
    # digits at all (e.g. a lone "0" misread as "i") is still worth one
    # more direct pixel-shape check (see _is_zero_shaped) - tested
    # directly, the digit whitelist alone cannot force this bundled
    # LSTM-only tessdata to reclassify that specific glyph. Anything
    # else is too ambiguous to guess, so the original OCR text is left
    # untouched rather than blanked or invented.
    stripped = text.strip()
    if expected == 1 and stripped and not _HAS_DIGIT_RE.search(stripped) and _is_zero_shaped(crop_bgr):
        return "0"
    return text


def _normalize_cell(label: str, text: str) -> str:
    if label in _NUMERIC_COLUMNS and _AMBIGUOUS_ONE_RE.match(text.strip()):
        return "1"
    if label == "Setting Type" and text.strip():
        best_match, best_score = None, 0.0
        for known in config.KNOWN_SETTING_TYPES:
            score = _similarity(text, known)
            if score > best_score:
                best_match, best_score = known, score
        if best_match is not None and best_score >= 0.8:
            return best_match
    return text


# --------------------------------------------------------------------------
# MM Size decimal recovery
# --------------------------------------------------------------------------
# Tesseract was directly tested (every PSM mode, OEM 1/3, upscale
# 3x-10x, binarization, dilation, a "0123456789.xX" character
# whitelist) against a real sample cell containing "1.8x1.8" and
# returned "18x18" in every single configuration - it reliably drops
# the tiny decimal-point glyph in this font/size, which is a genuine
# Tesseract text-recognition limitation, not a bug in this codebase's
# string handling (nothing here strips "." - grep finds no such logic).
#
# Root-cause fix: measure the actual pixels for decimal-point-shaped
# blobs via connected-component analysis (they are small, roughly
# square, and sit near the cell's baseline - clearly distinguishable
# from full-height digit glyphs) and re-insert "." at the x-position
# the blob evidence indicates. This is direct image evidence, not a
# guess, and it deliberately backs off - returning the OCR text
# UNCHANGED - whenever that evidence is not unambiguous.
_PURE_DIGIT_RUN_RE = re.compile(r"^[0-9]+([xX][0-9]+)?$")

# MM Size is a free-form measurement string, never a number: it can be
# a plain value ("0.85", "1"), a paired dimension ("1.8x1.8"), and
# future sheets may introduce other legitimate formats. It must never
# be parsed with int()/float() or written to Excel as a numeric cell.
_MM_SIZE_VALID_RE = re.compile(r"^\d+(\.\d+)?\s*[xX]\s*\d+(\.\d+)?$|^\d+(\.\d+)?$")


def validate_mm_size(value: str) -> bool:
    """True if value is a well-formed MM Size measurement: a plain
    integer/decimal ("0.85", "0.9", "1", "1.1") or a decimal-x-decimal
    pair ("1.8x1.8", "2.25x2.25"). Used to sanity-check extraction
    output, never to rewrite a value - an invalid result should be
    left as OCR produced it, not guessed at."""
    return bool(_MM_SIZE_VALID_RE.match(value.strip()))


def _recover_missing_decimals(crop_bgr: np.ndarray, ocr_text: str) -> str:
    text = ocr_text.strip()
    if "." in text or not text or not _PURE_DIGIT_RUN_RE.match(text):
        # Already has a decimal, or isn't a plain digit/x run - not a
        # shape this recovers; e.g. "FALSE", "6.5-7.0", "" pass through.
        return ocr_text

    if crop_bgr.size == 0:
        return ocr_text
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    scale = 8
    up = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_LANCZOS4)
    _, binary = cv2.threshold(up, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, _labels, stats, centroids = cv2.connectedComponentsWithStats(binary, connectivity=8)
    if n <= 1:
        return ocr_text

    components = [
        (centroids[i][0], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT])
        for i in range(1, n)
        if stats[i, cv2.CC_STAT_AREA] >= 15  # drop single-pixel noise specks
    ]
    if not components:
        return ocr_text

    max_h = max(h for _cx, _top, _w, h in components)
    img_h = up.shape[0]

    glyph_xs: list[float] = []
    period_xs: list[float] = []
    for cx, top, w, h in components:
        # A decimal point: much shorter than a full digit glyph, not
        # much wider than it is tall, and sitting on the baseline
        # (its bottom edge is in the lower part of the cell) rather
        # than spanning from the top like digits/letters do.
        if h < max_h * 0.35 and w < max_h * 0.5 and (top + h) > img_h * 0.55:
            period_xs.append(cx)
        else:
            glyph_xs.append(cx)

    if not period_xs:
        return ocr_text  # no blob evidence of a dropped decimal

    glyph_xs.sort()
    if len(glyph_xs) != len(text):
        # Blob count doesn't line up 1:1 with the OCR'd characters -
        # the evidence is ambiguous, so don't guess where "." goes.
        return ocr_text

    period_xs.sort()
    result: list[str] = []
    p_idx = 0
    for i, gx in enumerate(glyph_xs):
        while p_idx < len(period_xs) and period_xs[p_idx] < gx:
            result.append(".")
            p_idx += 1
        result.append(text[i])
    while p_idx < len(period_xs):
        result.append(".")
        p_idx += 1

    recovered = "".join(result)
    return recovered if validate_mm_size(recovered) else ocr_text


# --------------------------------------------------------------------------
# MM Size "?"-for-decimal-point misread recovery
# --------------------------------------------------------------------------
# Distinct failure mode from the dropped-decimal case above: instead of
# dropping the tiny decimal-point glyph entirely, Tesseract sometimes
# classifies it as a literal "?" character - confirmed directly against
# a real cell printed as "4.0x2.9" (both decimal points clearly visible
# on inspection of the upscaled crop), which OCR read as "40x?2.9": the
# first "." was dropped and a spurious "?" was inserted elsewhere in
# the string, so naively swapping "?" -> "." in place lands in the
# wrong position ("40x.2.9", not a valid MM Size shape) and gets
# rejected by validate_mm_size. "?" is schema-invalid in this column
# (_MM_SIZE_VALID_RE never allows it), so its position in the OCR text
# carries no information - the fix has to fall back to the same
# pixel-evidence reconstruction as _recover_missing_decimals: strip
# every "?" and "." from the OCR text down to a bare digit/x run, then
# rebuild the decimal(s) purely from period-shaped blobs.
#
# One extra wrinkle over the dropped-decimal case: here two glyphs
# ("x" and the following digit) were touching in the source image and
# came out as a single wide connected-component blob, so the simple
# "one component per character" assumption doesn't hold. A blob much
# wider than a normal single glyph is evidence it actually contains
# multiple touching glyphs, so it's split into round(width /
# reference_width) evenly-spaced slots (reference_width = the
# narrowest non-period blob, since merged blobs are only ever wider,
# never narrower, than a single glyph).
def _recover_question_mark_decimal(crop_bgr: np.ndarray, ocr_text: str) -> str:
    text = ocr_text.strip()
    if "?" not in text or crop_bgr.size == 0:
        return ocr_text

    bare = text.replace("?", "").replace(".", "")
    if not bare or not _PURE_DIGIT_RUN_RE.match(bare):
        return ocr_text

    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    scale = 8
    up = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_LANCZOS4)
    _, binary = cv2.threshold(up, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, _labels, stats, _centroids = cv2.connectedComponentsWithStats(binary, connectivity=8)
    if n <= 1:
        return ocr_text

    components = [
        (stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT])
        for i in range(1, n)
        if stats[i, cv2.CC_STAT_AREA] >= 15
    ]
    if not components:
        return ocr_text

    max_h = max(h for _left, _top, _w, h in components)
    img_h = up.shape[0]

    period_xs: list[float] = []
    glyph_widths: list[int] = []
    non_period: list[tuple[int, int, int, int]] = []
    for left, top, w, h in components:
        if h < max_h * 0.35 and w < max_h * 0.5 and (top + h) > img_h * 0.55:
            period_xs.append(left + w / 2)
        else:
            non_period.append((left, top, w, h))
            glyph_widths.append(w)

    if not period_xs or not non_period:
        return ocr_text  # no blob evidence of a misread decimal

    reference_w = min(glyph_widths)
    glyph_xs: list[float] = []
    for left, _top, w, _h in non_period:
        slots = max(1, round(w / reference_w))
        for i in range(slots):
            glyph_xs.append(left + (i + 0.5) * w / slots)

    if len(glyph_xs) != len(bare):
        # Slot count doesn't line up 1:1 with the stripped-down
        # character count - the evidence is ambiguous, don't guess.
        return ocr_text

    glyph_xs.sort()
    period_xs.sort()
    result: list[str] = []
    p_idx = 0
    for i, gx in enumerate(glyph_xs):
        while p_idx < len(period_xs) and period_xs[p_idx] < gx:
            result.append(".")
            p_idx += 1
        result.append(bare[i])
    while p_idx < len(period_xs):
        result.append(".")
        p_idx += 1

    recovered = "".join(result)
    return recovered if validate_mm_size(recovered) else ocr_text


def _has_ink(crop_bgr: np.ndarray, min_ratio: float = 0.003) -> bool:
    """True if a crop has enough dark pixels to plausibly contain real
    text, vs. a blank/near-blank cell that Tesseract can still return
    hallucinated garbage text for."""
    if crop_bgr.size == 0 or crop_bgr.shape[0] < 3 or crop_bgr.shape[1] < 3:
        return False
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    dark_ratio = (gray < 200).sum() / gray.size
    return dark_ratio >= min_ratio


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, _norm(a), _norm(b)).ratio()


@dataclass
class TableResult:
    rows: list[dict] = field(default_factory=list)
    total_pcs: str = ""
    total_wt: str = ""
    header_found: bool = False
    # X pixel coordinate separating the left info panel from the table
    # region. Words with cx < left_boundary are NOT part of the table.
    # Exposed so metadata_extractor can use the same split and avoid
    # the same left/right bleed-through in the other direction.
    left_boundary: float | None = None
    # The table's own merged title cell directly above the column
    # header row (e.g. "Engagement Ring", "Bracelet", "Pendant") - a
    # more specific sub-category than Product Type, which is derived
    # separately from the component list. Blank when no such title
    # band is found, never guessed.
    sub_category: str = ""


# ==========================================================================
# Grid-line detection (primary strategy)
# ==========================================================================


def _cluster_peaks(profile: np.ndarray, threshold: float, min_gap: int = 5) -> list[int]:
    idxs = [i for i, v in enumerate(profile) if v >= threshold]
    if not idxs:
        return []
    groups: list[list[int]] = [[idxs[0]]]
    for i in idxs[1:]:
        if i - groups[-1][-1] <= min_gap:
            groups[-1].append(i)
        else:
            groups.append([i])
    return [int(sum(g) / len(g)) for g in groups]


def _binary_image(img_bgr: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    return binary


def _detect_row_lines(binary: np.ndarray, width: int) -> list[int]:
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 20, 20), 1))
    horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)
    profile = (horiz > 0).sum(axis=1)
    return _cluster_peaks(profile, threshold=width * 0.25)


def _detect_col_lines(binary: np.ndarray, y0: int, y1: int) -> list[int]:
    band_h = y1 - y0
    if band_h <= 4:
        return []
    sub = binary[y0:y1, :]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(int(band_h * 0.5), 8)))
    vert = cv2.morphologyEx(sub, cv2.MORPH_OPEN, kernel, iterations=1)
    profile = (vert > 0).sum(axis=0)
    for frac in (0.6, 0.5, 0.4, 0.3):
        xs = _cluster_peaks(profile, threshold=band_h * frac)
        if len(xs) >= 6:
            return xs
    return _cluster_peaks(profile, threshold=band_h * 0.2)


def _match_header_columns(
    img_bgr: np.ndarray, y0: int, y1: int, col_xs: list[int]
) -> dict[str, tuple[int, int]]:
    """OCR each candidate column segment in the header band and fuzzy
    match it to a known label. Returns {label: (x0, x1)}. Segments that
    don't match any label (page-border slivers, dividers with no real
    column) are simply dropped, so the exact number of detected lines
    doesn't need to equal 9+1."""
    candidates: list[tuple[int, int, str]] = []
    for x0, x1 in zip(col_xs, col_xs[1:]):
        if x1 - x0 < 15:
            continue
        crop = img_bgr[y0 + 3 : y1 - 3, x0 + 4 : x1 - 4]
        text = ocr_cell(crop, psm=7)
        if text:
            candidates.append((x0, x1, text))

    assigned: dict[str, tuple[int, int]] = {}
    used_segments: set[int] = set()
    for label in config.TABLE_HEADER_LABELS:
        best_score = 0.0
        best_idx = None
        for i, (x0, x1, text) in enumerate(candidates):
            if i in used_segments:
                continue
            score = _similarity(text, label)
            if score > best_score:
                best_score = score
                best_idx = i
        if best_idx is not None and best_score >= 0.6:
            x0, x1, _ = candidates[best_idx]
            assigned[label] = (x0, x1)
            used_segments.add(best_idx)

    return assigned


def _find_header_band_by_grid_text(
    img_bgr: np.ndarray, binary: np.ndarray, row_ys: list[int]
) -> tuple[tuple[int, int], dict[str, tuple[int, int]]] | None:
    """Scan every candidate row band's own grid lines directly (per-
    cell OCR via _match_header_columns) for the one that best matches
    the known table column headers, independent of any full-page word
    pass. This is the same reliable per-cell-OCR strategy already used
    for data cells below - used here as the primary header locator so
    detection no longer depends on the word-based rough pass, which
    was observed to sometimes drop the header row's text entirely from
    the full-page tiled OCR (a known Tesseract failure mode on these
    dense sheets, same class of issue documented in ocr_engine's tiling
    docstring) even though the exact same cells read cleanly once
    OCR'd in isolation."""
    best_band = None
    best_columns: dict[str, tuple[int, int]] = {}
    for y0, y1 in zip(row_ys, row_ys[1:]):
        band_h = y1 - y0
        if not (15 <= band_h <= 100):
            continue
        col_xs = _detect_col_lines(binary, y0, y1)
        if len(col_xs) < 6:
            continue
        columns = _match_header_columns(img_bgr, y0, y1, col_xs)
        if len(columns) > len(best_columns):
            best_columns = columns
            best_band = (y0, y1)

    if best_band is None or len(best_columns) < 4:
        return None
    return best_band, best_columns


def _extract_table_via_grid(img_bgr: np.ndarray, approx_header_y: float | None) -> TableResult | None:
    height, width = img_bgr.shape[:2]
    binary = _binary_image(img_bgr)
    row_ys = _detect_row_lines(binary, width)
    if len(row_ys) < 4:
        return None

    # Find the row-line pair that brackets the rough header Y position
    # (from the word-based pass) with a plausible row height.
    header_band = None
    columns: dict[str, tuple[int, int]] = {}
    if approx_header_y is not None:
        best_dist = None
        for y0, y1 in zip(row_ys, row_ys[1:]):
            band_h = y1 - y0
            if not (15 <= band_h <= 100):
                continue
            mid = (y0 + y1) / 2.0
            dist = abs(mid - approx_header_y)
            if dist < 60 and (best_dist is None or dist < best_dist):
                best_dist = dist
                header_band = (y0, y1)

        if header_band is not None:
            col_xs = _detect_col_lines(binary, header_band[0], header_band[1])
            if len(col_xs) >= 6:
                columns = _match_header_columns(img_bgr, header_band[0], header_band[1], col_xs)
            if len(columns) < 4:
                # The word-based hint either found no band or pointed
                # at a band whose cells don't actually read as the
                # table header - fall through to the direct scan below
                # rather than giving up (Rule 16's grid-detection
                # strategy should not depend on a noisy full-page OCR
                # pass having found the header text).
                header_band = None
                columns = {}

    if header_band is None:
        found = _find_header_band_by_grid_text(img_bgr, binary, row_ys)
        if found is None:
            return None
        header_band, columns = found

    header_y0, header_y1 = header_band

    # Merge row-separator lines that fall well under a real row's
    # height into the previous line: these are duplicate/noise line
    # detections (anti-aliasing, a stray text stroke) rather than
    # genuine extra rows, and left unmerged they shred a single data
    # row's text across multiple garbled fragment "rows". The header
    # band's own height is a reliable per-image row-height reference
    # since it was just confirmed against the word-based header pass.
    header_h = header_y1 - header_y0
    min_row_gap = header_h * 0.5
    merged_row_ys = [row_ys[0]]
    for y in row_ys[1:]:
        # The header's own two boundary lines are never merged away -
        # they are the confirmed anchor the rest of this function
        # (and the row_ys.index(header_y0) lookup below) relies on.
        if y not in (header_y0, header_y1) and y - merged_row_ys[-1] < min_row_gap:
            continue
        merged_row_ys.append(y)
    row_ys = merged_row_ys

    result = TableResult(header_found=True)
    sorted_cols = sorted(columns.items(), key=lambda kv: kv[1][0])
    result.left_boundary = float(sorted_cols[0][1][0]) - 40.0
    table_x0 = min(x0 for x0, _x1 in columns.values())
    table_x1 = max(x1 for _x0, x1 in columns.values())

    header_idx_in_rows = row_ys.index(header_y0)

    # The table's merged title cell, if present, sits in the row band
    # immediately above the header band (e.g. "Engagement Ring" over
    # "Location | Shape | ..."). Only trust it when that band has a
    # plausible single-text-line height - never guessed from OCR text
    # found anywhere else on the sheet. Some real sheets have no title
    # at all (an empty merged cell there) - Tesseract was observed to
    # hallucinate long garbage strings ("eee eee...") from a blank
    # crop, so require actual ink in the band before even attempting
    # OCR rather than trusting whatever text comes back from nothing.
    if header_idx_in_rows >= 1:
        title_y0 = row_ys[header_idx_in_rows - 1]
        title_h = header_y0 - title_y0
        if 10 <= title_h <= 100:
            title_crop = img_bgr[title_y0 + 3 : header_y0 - 3, table_x0 + 4 : table_x1 - 4]
            if _has_ink(title_crop):
                result.sub_category = ocr_cell(title_crop, psm=7)

    pcs_running_sum = 0.0
    for y0, y1 in zip(row_ys[header_idx_in_rows + 1 :], row_ys[header_idx_in_rows + 2 :]):
        band_h = y1 - y0
        if band_h > 100 or band_h < 10:
            # Not a plausible data-row height any more - the table has ended.
            break

        # The Total row's label often spans several MERGED columns with
        # no internal divider, so "Total" can get chopped apart by
        # per-column slicing (e.g. "T" / "" / "tal"). OCR the row as one
        # unsliced blob just to test for the marker word, independent
        # of the per-column cells used for the actual field values.
        row_blob = img_bgr[y0 + 3 : y1 - 3, table_x0 + 4 : table_x1 - 4]
        is_total_row = _TOTAL_MARKER in ocr_cell(row_blob, psm=6).lower()

        row: dict[str, str] = {}
        for label, (cx0, cx1) in columns.items():
            crop = img_bgr[y0 + 3 : y1 - 3, cx0 + 4 : cx1 - 4]
            cell_text = ocr_cell(crop, psm=7)
            if label == "MM Size":
                cell_text = _recover_missing_decimals(crop, cell_text)
                cell_text = _recover_question_mark_decimal(crop, cell_text)
            elif label in _DIGITS_ONLY_COLUMNS:
                cell_text = _recover_digit_cell(crop, label, cell_text)
            row[label] = _normalize_cell(label, cell_text)

        full_row = {label: row.get(label, "") for label in config.TABLE_COLUMNS}

        # Backup signal: when the merged "Total" label OCR's too badly
        # for even the whole-row-blob text check to catch it (seen on
        # a real sample - it read as unrelated fragments, e.g. "re" /
        # "Bn" / "OS" / "Be" spilled into every column instead of
        # coming back blank), a row whose Pcs exactly equals the
        # running sum of every prior row's Pcs IS, in practice, the
        # Total row - but ONLY when the row's own cells don't actually
        # look like a real stone row (a well-formed MM Size, or a
        # recognized Setting Type). Without that corroboration, a
        # genuine data row that happens to repeat an earlier row's Pcs
        # count (e.g. two different stone positions both with Pcs=2,
        # each with perfectly valid MM Size/Setting Type values) is a
        # real coincidence that this signal must NOT swallow - doing
        # so was observed to silently drop every remaining stone row
        # on a real sample sheet.
        if not is_total_row and result.rows and pcs_running_sum > 0:
            mm_size = full_row.get("MM Size", "").strip()
            setting_type = full_row.get("Setting Type", "").strip()
            looks_like_real_row = (mm_size and validate_mm_size(mm_size)) or (
                setting_type in config.KNOWN_SETTING_TYPES
            )
            if not looks_like_real_row:
                try:
                    row_pcs = float(full_row.get("Pcs", "").strip())
                except ValueError:
                    row_pcs = None
                if row_pcs is not None and abs(row_pcs - pcs_running_sum) < 0.01:
                    is_total_row = True

        if is_total_row:
            result.total_pcs = full_row.get("Pcs", "")
            result.total_wt = full_row.get("Total Wt", "")
            break

        if any(v for v in full_row.values()):
            result.rows.append(full_row)
            try:
                pcs_running_sum += float(full_row.get("Pcs", "").strip())
            except ValueError:
                pass

    return result


# ==========================================================================
# Word-bounding-box fallback strategy
# ==========================================================================


def _find_header_line(lines: list[list[Word]]):
    """Return (line_index, {label: x_center}) for the best-matching
    header row, or None if no line looks like the table header."""
    best_idx = None
    best_score = 0
    best_anchors: dict[str, float] = {}

    for idx, line in enumerate(lines):
        if len(line) < 3:
            continue
        anchors: dict[str, float] = {}
        used_indices: set[int] = set()
        matched = 0

        for label in config.TABLE_HEADER_LABELS:
            best_span_score = 0.0
            best_span_cx = None
            best_span_indices: tuple[int, ...] | None = None

            for span_len in (1, 2, 3):
                for start in range(len(line) - span_len + 1):
                    span_indices = tuple(range(start, start + span_len))
                    if used_indices.intersection(span_indices):
                        continue
                    span = line[start : start + span_len]
                    span_text = " ".join(w.text for w in span)
                    score = _similarity(span_text, label)
                    if score > best_span_score:
                        best_span_score = score
                        best_span_cx = sum(w.cx for w in span) / len(span)
                        best_span_indices = span_indices

            if best_span_score >= 0.72 and best_span_cx is not None:
                anchors[label] = best_span_cx
                used_indices.update(best_span_indices)
                matched += 1

        if matched > best_score:
            best_score = matched
            best_idx = idx
            best_anchors = anchors

    if best_idx is None or best_score < 4:
        return None
    return best_idx, best_anchors


def _rough_header_position(words: list[Word]) -> tuple[float | None, dict[str, float] | None]:
    """Word-based pass used both to seed the grid detector with an
    approximate header Y, and as the fallback table extractor."""
    lines = group_words_into_lines(words)
    header = _find_header_line(lines)
    if header is None:
        return None, None
    header_idx, column_anchors = header
    if not column_anchors:
        return None, None
    header_line = lines[header_idx]
    approx_y = sum(w.cy for w in header_line) / len(header_line)
    return approx_y, column_anchors


def _extract_table_via_anchors(words: list[Word], column_anchors: dict[str, float]) -> TableResult:
    result = TableResult(header_found=True)

    anchor_values = sorted(column_anchors.values())
    gaps = [b - a for a, b in zip(anchor_values, anchor_values[1:])]
    margin = (min(gaps) / 2.0) if gaps else 60.0
    left_boundary = anchor_values[0] - margin
    result.left_boundary = left_boundary

    table_words = [w for w in words if w.cx >= left_boundary]
    lines = group_words_into_lines(table_words)
    header2 = _find_header_line(lines)
    header_idx = header2[0] if header2 is not None else -1

    for line in lines[header_idx + 1 :]:
        if not line:
            continue
        line_str = line_text(line).strip()
        if not line_str:
            continue

        is_total_row = _TOTAL_MARKER in line_str.lower()

        cells: dict[str, list[str]] = {label: [] for label in column_anchors}
        for w in line:
            nearest_label = min(column_anchors, key=lambda lbl: abs(column_anchors[lbl] - w.cx))
            cells[nearest_label].append(w.text)

        row = {label: " ".join(cells.get(label, [])).strip() for label in config.TABLE_COLUMNS}

        if is_total_row:
            result.total_pcs = row.get("Pcs", "")
            result.total_wt = row.get("Total Wt", "")
            break

        if any(v for v in row.values()):
            result.rows.append(row)

    return result


# ==========================================================================
# Entry point
# ==========================================================================


def extract_table(img_bgr: np.ndarray, words: list[Word]) -> TableResult:
    approx_header_y, column_anchors = _rough_header_position(words)

    grid_result = _extract_table_via_grid(img_bgr, approx_header_y)
    if grid_result is not None and grid_result.rows:
        return grid_result

    if column_anchors:
        return _extract_table_via_anchors(words, column_anchors)

    return grid_result if grid_result is not None else TableResult()

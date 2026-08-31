"""
Tesseract OCR wrapper.

Produces word-level bounding boxes (left/top/width/height/conf) for a
whole image. All higher-level modules (metadata_extractor,
table_extractor) work from these word boxes rather than raw dumped
text, so layout position is always available for column/row mapping.
"""

from __future__ import annotations

import re
import statistics
import subprocess
from dataclasses import dataclass

import cv2
import numpy as np
import pytesseract
from pytesseract import Output

import config

pytesseract.pytesseract.tesseract_cmd = config.TESSERACT_CMD

# Every OCR call here launches tesseract.exe as a subprocess (some
# extraction paths do this 50-100+ times per image, e.g. one call per
# table cell). pytesseract hides the child window via STARTUPINFO/
# SW_HIDE but never sets CREATE_NO_WINDOW, so it still allocates a
# console for each child. From a --windowed (console-less) parent
# process, this was observed to crash the whole app partway through a
# large image's tiled OCR pass. CREATE_NO_WINDOW skips allocating a
# console for the child entirely, which is what we want either way.
if hasattr(subprocess, "CREATE_NO_WINDOW"):
    _original_subprocess_args = pytesseract.pytesseract.subprocess_args

    def _patched_subprocess_args(include_stdout: bool = True) -> dict:
        kwargs = _original_subprocess_args(include_stdout)
        kwargs["creationflags"] = kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
        return kwargs

    pytesseract.pytesseract.subprocess_args = _patched_subprocess_args


@dataclass
class Word:
    text: str
    left: int
    top: int
    width: int
    height: int
    conf: float

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    @property
    def cx(self) -> float:
        return self.left + self.width / 2.0

    @property
    def cy(self) -> float:
        return self.top + self.height / 2.0


def load_image_bgr(path: str) -> np.ndarray:
    # cv2.imread mishandles some non-ASCII / spaced Windows paths, so
    # read via numpy buffer instead.
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Could not read image: {path}")
    return img


def preprocess_for_ocr(img_bgr: np.ndarray) -> np.ndarray:
    """Upscale while KEEPING color (colored text in these sheets
    carries meaning and must remain OCR-able, Rule 14).

    LANCZOS4 is used because these sheets pack small, dense table text
    against thin grid lines: CUBIC/LINEAR interpolation and CLAHE
    contrast enhancement were both tested and found to blur/distort
    that text badly enough that Tesseract dropped whole header cells
    (e.g. "Location", "Shape", "Pcs", "Setting" vanished entirely on a
    real sample sheet) while LANCZOS4 alone read them cleanly.
    """
    scale = config.OCR_UPSCALE_FACTOR
    if scale != 1.0:
        img_bgr = cv2.resize(
            img_bgr, None, fx=scale, fy=scale, interpolation=cv2.INTER_LANCZOS4
        )
    return img_bgr


# Tokens that are pure grid-line/border OCR noise (table cell borders
# misread as pipes/brackets) and never carry real cell data on these
# sheets. Punctuation that CAN legitimately appear in values (-, ., =,
# #, /) is intentionally excluded.
_NOISE_TOKEN_RE = re.compile(r"^[|\[\]{}~`^¦‖]+$")


def _words_from_tesseract_dict(data: dict, scale: float) -> list[Word]:
    words: list[Word] = []
    n = len(data["text"])
    for i in range(n):
        text = data["text"][i].strip()
        if not text or _NOISE_TOKEN_RE.match(text):
            continue
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1.0
        if conf < 0:
            continue
        words.append(
            Word(
                text=text,
                left=int(round(data["left"][i] / scale)),
                top=int(round(data["top"][i] / scale)),
                width=int(round(data["width"][i] / scale)),
                height=int(round(data["height"][i] / scale)),
                conf=conf,
            )
        )
    return words


def ocr_words(img_bgr: np.ndarray, psm: int = 11) -> list[Word]:
    """Run OCR on a single (BGR) image/region and return word boxes in
    the ORIGINAL (pre-preprocess) pixel coordinate space of img_bgr."""
    processed = preprocess_for_ocr(img_bgr)
    scale = config.OCR_UPSCALE_FACTOR
    cfg = f"--oem 3 --psm {psm}"
    data = pytesseract.image_to_data(processed, config=cfg, output_type=Output.DICT)
    return _words_from_tesseract_dict(data, scale)


def ocr_words_tiled(
    img_bgr: np.ndarray, psm: int = 11, tile_height: int = 60, overlap: int = 20
) -> list[Word]:
    """OCR the page in horizontal strips instead of one giant upscaled
    image.

    These CAD sheets pack small, dense text (stone tables) next to
    large near-blank CAD renders. At full-page scale, upscaling for
    OCR balloons a ~1850x1150 sheet to tens of megapixels, and
    Tesseract was observed to silently drop entire header cells
    ("Location", "Shape", "Pcs", "Setting"...) that it reads correctly
    when the same text is OCR'd in isolation. Tiling keeps each OCR
    call's upscaled size small and reliable.

    Each tile overlaps its neighbors; only words whose vertical center
    falls in a tile's non-overlapping "core" band are kept, so every
    row is attributed to exactly one tile (whichever one has that row
    away from its seam) without duplicates or seam-cut misses.
    """
    height = img_bgr.shape[0]
    if height <= tile_height:
        return ocr_words(img_bgr, psm=psm)

    step = max(tile_height - overlap, 1)
    all_words: list[Word] = []
    y = 0
    while y < height:
        y_end = min(y + tile_height, height)
        tile = img_bgr[y:y_end, :]
        if tile.shape[0] < 5:
            break

        core_top = y if y == 0 else y + overlap // 2
        core_bottom = y_end if y_end >= height else y_end - overlap // 2

        for w in ocr_words(tile, psm=psm):
            cy = y + w.cy
            if core_top <= cy < core_bottom:
                all_words.append(
                    Word(text=w.text, left=w.left, top=w.top + y, width=w.width, height=w.height, conf=w.conf)
                )

        if y_end >= height:
            break
        y += step

    return all_words


def ocr_image_file(path: str, psm: int = 11) -> list[Word]:
    img = load_image_bgr(path)
    return ocr_words(img, psm=psm)


def ocr_text(img_bgr: np.ndarray, psm: int = 7) -> str:
    """Plain text OCR for a small, already-cropped region (e.g. one
    table cell). PSM 7 = treat region as a single text line."""
    processed = preprocess_for_ocr(img_bgr)
    cfg = f"--oem 3 --psm {psm}"
    return pytesseract.image_to_string(processed, config=cfg).strip()


def ocr_cell(img_bgr: np.ndarray, psm: int = 7, upscale: float = 6.0, extra_config: str = "") -> str:
    """OCR one already-isolated table cell crop. Table cells are tiny
    (often well under 150x40px), so they get their own, more aggressive
    upscale independent of config.OCR_UPSCALE_FACTOR (proven reliable
    in testing: isolated single-line crops read cleanly where the same
    text embedded in a larger multi-row image did not)."""
    if img_bgr.size == 0 or img_bgr.shape[0] < 3 or img_bgr.shape[1] < 3:
        return ""
    up = cv2.resize(img_bgr, None, fx=upscale, fy=upscale, interpolation=cv2.INTER_LANCZOS4)
    cfg = f"--oem 3 --psm {psm}"
    if extra_config:
        cfg = f"{cfg} {extra_config}"
    return pytesseract.image_to_string(up, config=cfg).strip()


def ocr_region_lines(img_bgr: np.ndarray, min_ink_ratio: float = 0.01, min_gap: int = 2) -> list[str]:
    """Text-only convenience wrapper around ocr_region_lines_with_bounds
    (see there for the full docstring) for callers that don't need
    each line's pixel bounds."""
    return [text for text, _y0, _y1 in ocr_region_lines_with_bounds(img_bgr, min_ink_ratio, min_gap) if text]


def ocr_region_lines_with_bounds(
    img_bgr: np.ndarray, min_ink_ratio: float = 0.01, min_gap: int = 2
) -> list[tuple[str, int, int]]:
    """OCR a small multi-line metadata region (e.g. the left info
    panel's "Total Components / Two Tone" gap) by first splitting it
    into individual text-line bands via a row-wise ink projection,
    then OCR'ing each band on its own and returning one (text, y0, y1)
    tuple per line - the (y0, y1) band bounds (in img_bgr's own
    coordinate space) let a caller re-attempt a specific line with
    direct pixel access if its returned text doesn't parse as
    expected, rather than only ever seeing the flattened string.

    Testing found that a single OCR call over a region spanning
    multiple separated text lines (with a blank box or gap between
    them) can silently DROP an entire line, even though that exact
    same line reads perfectly when cropped and OCR'd by itself - the
    same class of "OCR a whole block at once is unreliable" issue
    already fixed for the main page and for table cells elsewhere in
    this codebase. Splitting into single-line bands first sidesteps it.

    Uses a "distance from white" grayscale (the min of the B/G/R
    channels) rather than luminance, because it was found that
    standard grayscale gives light/saturated colors like cyan far
    less contrast against a white background than darker colors like
    red or black get, even though both are clearly legible to a human.
    """
    if img_bgr.size == 0:
        return []

    # Vertical border/divider lines often sit right at the crop's left
    # and/or right edge and contribute a constant handful of dark
    # pixels to EVERY row, which never dips back below a naive
    # threshold - the whole region then looks like one unbroken band.
    # Trim a margin before computing the ink profile (but OCR the
    # full-width band once found, so no real text is lost).
    margin = min(10, img_bgr.shape[1] // 4)
    inner = img_bgr[:, margin : img_bgr.shape[1] - margin] if img_bgr.shape[1] > 2 * margin else img_bgr

    b, g, r = cv2.split(inner.astype(np.int16))
    min_channel = np.minimum(np.minimum(b, g), r).astype(np.uint8)
    dark = min_channel < 200

    # An interior vertical border (e.g. a left-panel box edge that
    # happens to fall inside the crop rather than right at its edge,
    # which the margin trim above doesn't reach) is dark in nearly
    # every row - something real text never does across a multi-line
    # region. Drop such columns before building the row-wise ink
    # profile so one stray border line can't make the whole region
    # look like a single unbroken band and swallow the per-line split.
    row_count = dark.shape[0]
    if row_count > 0:
        col_dark_ratio = dark.sum(axis=0) / row_count
        text_cols = col_dark_ratio < 0.8
        if text_cols.any():
            dark = dark[:, text_cols]

    ink = dark.sum(axis=1)
    threshold = max(int(dark.shape[1] * min_ink_ratio), 1)

    bands: list[tuple[int, int]] = []
    in_band = False
    start = 0
    for y, v in enumerate(ink):
        if v >= threshold and not in_band:
            in_band = True
            start = y
        elif v < threshold and in_band:
            in_band = False
            bands.append((start, y))
    if in_band:
        bands.append((start, len(ink)))

    merged: list[list[int]] = []
    for band_start, band_end in bands:
        if merged and band_start - merged[-1][1] <= min_gap:
            merged[-1][1] = band_end
        else:
            merged.append([band_start, band_end])

    # Thin horizontal divider/border lines crossing the crop show up
    # as their own 1-3px "band" (far shorter than any real glyph) -
    # drop those before OCR'ing rather than risk them becoming a
    # spurious garbage "line" that a caller might treat as real text.
    min_band_height = 10
    merged = [b for b in merged if b[1] - b[0] >= min_band_height]

    lines: list[tuple[str, int, int]] = []
    for band_start, band_end in merged:
        pad = 3
        y0 = max(band_start - pad, 0)
        y1 = min(band_end + pad, img_bgr.shape[0])
        # A lower upscale than ocr_cell's table-cell default (6x) was
        # found to read this metadata-panel text more reliably - at
        # 6x, some smaller-font lines (e.g. a dot-leader "Bell.....1"
        # entry) blurred into garbage that 3-4x read correctly.
        text = ocr_cell(img_bgr[y0:y1, :], psm=7, upscale=4.0)
        lines.append((text, y0, y1))
    return lines


def group_words_into_lines(words: list[Word], y_tolerance_ratio: float = 0.6) -> list[list[Word]]:
    """Cluster words into visual text lines using vertical position,
    independent of Tesseract's own block/line numbering (which is
    unreliable on these multi-region sheets)."""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w.cy, w.left))
    median_h = statistics.median(w.height for w in sorted_words) or 10

    lines: list[list[Word]] = []
    current: list[Word] = [sorted_words[0]]
    current_cy = sorted_words[0].cy

    for w in sorted_words[1:]:
        if abs(w.cy - current_cy) <= median_h * y_tolerance_ratio:
            current.append(w)
            current_cy = statistics.mean(x.cy for x in current)
        else:
            lines.append(sorted(current, key=lambda x: x.left))
            current = [w]
            current_cy = w.cy
    lines.append(sorted(current, key=lambda x: x.left))
    return lines


def line_text(line: list[Word]) -> str:
    return " ".join(w.text for w in line)

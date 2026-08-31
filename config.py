"""
Central configuration for the JPG TO EXCEL CONVERTER.

Everything that is likely to need tuning later (fixed Excel columns,
known section labels, table column headers, OCR paths, etc.) lives here
so the rest of the codebase never hard-codes it.
"""

import os
import sys


# --------------------------------------------------------------------------
# Tesseract OCR location
# --------------------------------------------------------------------------
# When the app is frozen with PyInstaller we ship the Tesseract-OCR folder
# next to the EXE (see build.bat) so the target machine needs nothing
# pre-installed. When running from source, fall back to the normal
# Program Files install used during development.

def _tesseract_cmd() -> str:
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
        bundled = os.path.join(base, "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(bundled):
            return bundled
    default = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default):
        return default
    return "tesseract"  # rely on PATH as a last resort


TESSERACT_CMD = _tesseract_cmd()

# OCR is run at this DPI-equivalent scale factor to improve small-text
# accuracy (CAD sheets are dense, low-contrast scans/exports).
OCR_UPSCALE_FACTOR = 4.0

# Minimum OCR confidence (0-100) below which a word is still kept for
# table/metadata assembly but is NOT allowed to overwrite a code/number
# field with a low-confidence guess (see Rule 26: never invent data).
MIN_CONFIDENCE_FOR_CRITICAL_FIELDS = 40


# --------------------------------------------------------------------------
# Fixed Excel output schema (Rule 9 / Rule 10)
# --------------------------------------------------------------------------
# This list is the single source of truth for the output workbook's
# columns. Order matters. Do not let OCR create/rename/reorder headers.

FIXED_HEADERS = [
    "Company",
    "Date",
    "Design Code",
    "Product Type",
    "Total Components",
    "Component Details",
    "EXTRA COMPONENTS",
    "Two Tone",
    "Metal Details",
    "Designer",
    "Location",
    "Shape",
    "MM Size",
    "Sieve Size",
    "Per Stone Wt",
    "Pcs",
    "Total Wt",
    "All Total Pcs",
    "All Total Wt",
    "Setting Type",
    "Setting",
    "Design Code 2",
    "Sub Category",
]

# Columns that belong to the stone/component table region (used by the
# table extractor to know which fixed columns it is allowed to fill).
# NOTE: "All Total Pcs"/"All Total Wt" are NOT here - they come from the
# table's single TOTAL row (one value per image), not per-row cells, and
# are populated like the other per-image metadata fields below.
TABLE_COLUMNS = [
    "Location",
    "Shape",
    "MM Size",
    "Sieve Size",
    "Per Stone Wt",
    "Pcs",
    "Total Wt",
    "Setting Type",
    "Setting",
]

# Columns that belong to per-image metadata (repeated on every row of
# that image, Rule 8). Includes the table's printed Total row values
# (All Total Pcs / All Total Wt), which are one value per image, not
# per stone row, and so repeat across every row the same way.
METADATA_COLUMNS = [
    "Company",
    "Date",
    "Design Code",
    "Product Type",
    "Total Components",
    "Component Details",
    "EXTRA COMPONENTS",
    "Two Tone",
    "Metal Details",
    "Designer",
    "All Total Pcs",
    "All Total Wt",
    "Design Code 2",
    "Sub Category",
]

WORKSHEET_NAME = "Extracted Data"
DEBUG_WORKSHEET_NAME = "OCR Debug"

# Set True to also emit the raw-OCR debug worksheet in the workbook.
DEBUG_MODE = False


# --------------------------------------------------------------------------
# Known label vocabulary used by the metadata extractor.
# --------------------------------------------------------------------------
# These are recognition anchors only (fuzzy-matched against OCR text);
# they do NOT restrict what values can be captured after the label.

LABEL_TOTAL_COMPONENTS = "Total Number of Components"
LABEL_TWO_TONE = "Two tone"
LABEL_METAL_DETAILS = "Metal details"
LABEL_DATE = "Date"

# Recognized product/component type words. This list only *helps* the
# detector recognize common types faster; unrecognized words found next
# to a quantity (e.g. "Bell...1") are still captured dynamically.
KNOWN_COMPONENT_TYPES = [
    "Bracelet",
    "Pendant",
    "Necklace",
    "Ring",
    "Earring",
    "Earrings",
    "Bangle",
    "Chain",
    "Bell",
    "Charm",
    "Anklet",
    "Brooch",
]

# Words that mark a table row as a TOTAL row rather than a data row
# (Rule 11). Matched case-insensitively against the first non-empty
# cell of a row.
TOTAL_ROW_MARKERS = ["total", "grand total"]

# Values that must be preserved exactly as-is and never blanked out
# (Rule 12).
PRESERVE_LITERAL_VALUES = {"false", "true"}

# Known stone-setting vocabulary for the "Setting Type" column. Tesseract
# has been observed to consistently misread "Prong" as "Prona" in these
# sheets (a font/glyph confusion, not random noise), so close OCR output
# is snapped to the canonical spelling from this small, controlled list -
# not a guess, a spell-correction against a known domain vocabulary.
KNOWN_SETTING_TYPES = [
    "Prong",
    "Split Prong",
    "Micro Prong",
    "Claw-Prong",
    "Bezel",
    "Pave",
    "Micro Pave",
    "Channel",
    "Flush Set",
    "Tension",
    "Cluster",
    "Halo",
    "Bead Set",
]

# Stone table header labels used to locate the table's header row and
# anchor column X-positions. Matched fuzzily (case-insensitive, partial).
TABLE_HEADER_LABELS = [
    "Location",
    "Shape",
    "MM Size",
    "Sieve Size",
    "Per Stone Wt",
    "Pcs",
    "Total Wt",
    "Setting Type",
    "Setting",
]


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def default_output_dir() -> str:
    """Folder the generated workbook is written to."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(base, "Output")
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


OUTPUT_FILENAME_PREFIX = "Extracted_Data"

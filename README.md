# JPG to Excel Converter

Converts jewelry CAD/design specification sheet JPGs into a single,
structured Excel workbook using layout-aware OCR (not a raw text dump).

## What it does

```
JPG IMAGE(S)
    -> OCR (Tesseract, per-word bounding boxes, tiled across the page)
    -> layout split (left info panel vs. CAD/table region)
    -> metadata extraction (company, date, design code, components,
       two tone, metal details, designer) - critical fields (design
       code, component list) are re-OCR'd in isolation for accuracy
    -> stone/component table extraction: real grid-line detection
       (OpenCV) + per-cell OCR, with a bounding-box-anchor fallback
       if a sheet's borders aren't detected confidently
    -> Total row detected separately (text marker + numeric-sum
       cross-check) and never added as a stone row
    -> one workbook, one fixed header, all images appended in order
    -> Excel opens automatically, maximized
```

Building this against the three real sample sheets (bracelet, pendant,
necklace) surfaced two non-obvious realities worth knowing about this
document type, both already handled by the code above:

1. **OCR'ing a whole row/region at once is unreliable.** Tesseract
   regularly dropped or garbled entire header cells (e.g. "Location",
   "Pcs", "Setting" vanished outright) when run across a full page or
   even a ~100px-tall multi-row strip, but read the exact same text
   correctly when it was cropped down to one row or one cell in
   isolation. `table_extractor.py` therefore detects the table's real
   grid lines and OCRs each cell on its own; `metadata_extractor.py`
   does the same for the design code and component list.
2. **The left info panel and the table sit at overlapping Y-positions**
   on these sheets, so naive top-to-bottom text reconstruction merges
   them (e.g. "Two tone=no" ending up inside a table row). Both
   extractors split words by X-position before reconstructing lines.

## Project layout

| File                    | Purpose                                                        |
|--------------------------|-----------------------------------------------------------------|
| `config.py`              | Fixed Excel headers, label vocabulary, OCR/Tesseract settings  |
| `ocr_engine.py`          | Tesseract wrapper: image preprocessing, word boxes, line grouping |
| `metadata_extractor.py`  | Company / Date / Design Code / Components / Two Tone / Metal / Designer |
| `table_extractor.py`     | Stone/component table -> rows, using column-anchor matching     |
| `excel_export.py`        | Builds the formatted workbook, saves it, opens + maximizes Excel |
| `ui.py`                  | Simple Tkinter UI (Select Images / Convert to Excel)            |
| `main.py`                | Pipeline orchestration + entry point                            |
| `build.bat`              | Builds the standalone EXE with PyInstaller                      |

## Running from source

```
venv\Scripts\python.exe main.py
```

Or headless (prints results, still writes the workbook, useful for
testing without the GUI):

```
venv\Scripts\python.exe main.py path\to\image1.jpg path\to\image2.jpg
```

## Building the EXE

Requires Python (already set up in `venv/`) and a Tesseract-OCR
install at `C:\Program Files\Tesseract-OCR` (used only to source the
files that get bundled — the target machine needs neither).

```
build.bat
```

Output: `dist\JPG_To_Excel_Converter\JPG_To_Excel_Converter.exe`, with
`Tesseract-OCR\` bundled alongside it. Ship the whole
`dist\JPG_To_Excel_Converter` folder — the EXE needs the files next to
it (icons/DLLs from `--onedir`, plus the bundled `Tesseract-OCR`
folder).

## Configuration

All fixed Excel columns live in `config.py` as `FIXED_HEADERS`. Change
them there — OCR never creates or renames columns.

Set `config.DEBUG_MODE = True` to also emit an "OCR Debug" worksheet
with every raw OCR word, its bounding box, and confidence.

## Known limitations

- OCR accuracy depends on image resolution/clarity. Low-confidence or
  unrecognized fields are left blank rather than guessed (never
  invents data). A well-scoped exception: a handful of known,
  narrowly-targeted OCR glyph confusions are normalized rather than
  left broken - a lone "|"/"I"/"l" in a numeric column becomes "1",
  and Setting Type values are snapped to the closest entry in
  `config.KNOWN_SETTING_TYPES` (e.g. "Prona" -> "Prong") when the
  match is close (>=0.8 similarity). Both are spell-correction against
  a controlled vocabulary, not guessing new data.
- Residual single-character OCR noise is possible on very small text
  (e.g. an "O"/"0" mix-up inside a design code's second parenthesized
  segment, or a dropped decimal point in a densely-packed cell like
  "1.8x1.8"). These are left as OCR reads them rather than
  auto-"corrected", per the accuracy-over-correction rule for codes
  and numbers - verify design codes on critical documents.
- If a grid-line-based table isn't confidently detected (broken/faint
  borders, heavy skew), the table extractor falls back to matching
  header-label words by position; if even that fails to match at
  least 4 of the 9 known columns, the image contributes a
  metadata-only row with blank table columns rather than risk
  mismatched columns.
- Validated end-to-end against the three real sample sheets (bracelet,
  pendant, necklace) in `samples/` and against a synthetic mock sheet.
  New sheet layouts/fonts should be spot-checked before relying on
  this for production data entry.
- The built EXE explicitly patches `pytesseract`'s subprocess launch
  with `CREATE_NO_WINDOW` and gives `sys.stdin/stdout/stderr` real
  null-device file objects at startup - a `--windowed` PyInstaller
  build has none of the three by default, which is otherwise a subtle
  source of crashes for any GUI app that shells out to a CLI tool
  (`ocr_engine.py`, top of `main.py`).

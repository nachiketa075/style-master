"""
JPG TO EXCEL CONVERTER - pipeline orchestration + entry point.

Flow (per spec):
  JPG -> OCR -> layout detection -> metadata extraction ->
  table extraction -> fixed Excel rows -> workbook -> open + maximize.

Running this file directly launches the Tkinter UI. It can also be
run headless for testing:  python main.py image1.jpg image2.jpg
"""

from __future__ import annotations

import os
import sys
import traceback
from dataclasses import dataclass

# A --windowed PyInstaller build has no real console, so sys.stdin /
# sys.stdout / sys.stderr are all None. An accidental print() or
# library warning would crash the app outright, so give all three
# real (null) file objects.
if getattr(sys, "frozen", False):
    if sys.stdin is None:
        sys.stdin = open(os.devnull, "r")
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")

import config
import excel_export
import table_extractor
from metadata_extractor import extract_metadata
from ocr_engine import Word, load_image_bgr, ocr_words_tiled


@dataclass
class ImageResult:
    path: str
    ok: bool
    row_count: int = 0
    error: str = ""


def _rows_for_image(image_path: str, debug_words_out: list[dict] | None = None) -> tuple[list[dict], ImageResult]:
    img = load_image_bgr(image_path)
    height, width = img.shape[:2]

    words: list[Word] = ocr_words_tiled(img, psm=11)
    if debug_words_out is not None:
        name = os.path.basename(image_path)
        for w in words:
            debug_words_out.append(
                {
                    "Image": name,
                    "Text": w.text,
                    "Left": w.left,
                    "Top": w.top,
                    "Width": w.width,
                    "Height": w.height,
                    "Confidence": w.conf,
                }
            )

    if not words:
        return [], ImageResult(path=image_path, ok=False, error="No text detected in image.")

    table = table_extractor.extract_table(img, words)
    meta = extract_metadata(img, words, width, height, table_left_boundary=table.left_boundary)

    base_row = {
        "Company": meta.company,
        "Date": meta.date,
        "Design Code": meta.design_code.lower(),
        "Design Code 2": meta.design_code_2,
        "Sub Category": table.sub_category.lower(),
        "Product Type": meta.product_type.lower(),
        "Total Components": meta.total_components,
        "Component Details": meta.component_details,
        "EXTRA COMPONENTS": meta.extra_components,
        "Two Tone": meta.two_tone,
        "Designer": meta.designer,
        # The table's own printed TOTAL row (Rule: read the printed
        # total, never compute it independently unless OCR truly can't
        # find it) - one value per image, repeated on every row exactly
        # like the other metadata fields above.
        "All Total Pcs": table.total_pcs,
        "All Total Wt": table.total_wt,
    }

    # Product-level and component-level fields are kept in separate
    # rows rather than merged: the PRODUCT row (metadata only, no
    # Location/Shape/etc.) always comes first, immediately followed by
    # one row per detected component/stone-table row (Location/Shape/
    # etc. only, no Company/Date/etc.) - dynamic in count, never
    # hardcoded. One logical Metal Details statement (Rule: never
    # joined into one cell) = one product row, so a sheet with several
    # separate Metal Details entries repeats the product+component
    # block once per entry; the common case of a single entry yields
    # exactly one PRODUCT row followed by the component rows.
    stone_rows = table.rows
    metal_entries = meta.metal_details_list if meta.metal_details_list else [""]

    rows: list[dict] = []
    for metal_entry in metal_entries:
        product_row = dict(base_row)
        product_row["Metal Details"] = metal_entry
        rows.append(product_row)
        for stone_row in stone_rows:
            rows.append(dict(stone_row))

    return rows, ImageResult(path=image_path, ok=True, row_count=len(rows))


def process_images(image_paths: list[str], progress_cb=None) -> tuple[str, list[ImageResult]]:
    """Run the full pipeline over all selected images and produce one
    workbook. Returns (output_file_path, per_image_results)."""
    if not image_paths:
        raise ValueError("Please select at least one image.")

    all_rows: list[dict] = []
    results: list[ImageResult] = []
    debug_words: list[dict] = [] if config.DEBUG_MODE else None  # type: ignore[assignment]

    total = len(image_paths)
    for i, path in enumerate(image_paths, start=1):
        if progress_cb:
            progress_cb(i, total, os.path.basename(path))
        try:
            rows, result = _rows_for_image(path, debug_words)
            all_rows.extend(rows)
            results.append(result)
        except Exception as exc:  # noqa: BLE001 - report, don't crash the batch
            traceback.print_exc()
            results.append(ImageResult(path=path, ok=False, error=str(exc)))

    if not all_rows:
        raise RuntimeError("Unable to extract data from the selected image(s).")

    wb = excel_export.build_workbook(all_rows)
    if config.DEBUG_MODE and debug_words:
        excel_export.write_debug_sheet(wb, debug_words)

    out_dir = config.default_output_dir()
    out_path = excel_export.unique_output_path(out_dir)
    excel_export.save_workbook(wb, out_path)

    return out_path, results


def _run_headless(paths: list[str]) -> None:
    out_path, results = process_images(paths)
    for r in results:
        status = "OK" if r.ok else f"FAILED: {r.error}"
        print(f"{os.path.basename(r.path)}: {status} ({r.row_count} rows)")
    print(f"\nSaved: {out_path}")


def main() -> None:
    if len(sys.argv) > 1:
        _run_headless(sys.argv[1:])
        return
    import ui

    ui.launch_app(process_images)


def _crash_guarded_main() -> None:
    try:
        main()
    except Exception:
        try:
            log_path = os.path.join(
                os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else ".",
                "crash_log.txt",
            )
            with open(log_path, "w") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass
        raise


if __name__ == "__main__":
    _crash_guarded_main()

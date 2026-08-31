"""
Simple Tkinter UI for the JPG TO EXCEL CONVERTER (Rule 20).

    JPG TO EXCEL CONVERTER
    Select your JPG images
        [ SELECT IMAGES ]
    Selected Images: ...
        [ CONVERT TO EXCEL ]
    Status: Ready

No dashboards, no menus, no dark theme - just the two actions the
user needs.
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import excel_export

_BG = "#F4F6F8"
_ACCENT = "#1F4E78"
_TEXT = "#222222"


def launch_app(process_images_fn) -> None:
    root = tk.Tk()
    root.title("JPG to Excel Converter")
    root.geometry("520x460")
    root.minsize(480, 420)
    root.configure(bg=_BG)

    selected_paths: list[str] = []

    title = tk.Label(
        root,
        text="JPG TO EXCEL CONVERTER",
        font=("Segoe UI", 16, "bold"),
        bg=_BG,
        fg=_ACCENT,
    )
    title.pack(pady=(20, 4))

    subtitle = tk.Label(
        root,
        text="Select your JPG images",
        font=("Segoe UI", 10),
        bg=_BG,
        fg=_TEXT,
    )
    subtitle.pack(pady=(0, 12))

    select_btn = tk.Button(
        root,
        text="SELECT IMAGES",
        font=("Segoe UI", 10, "bold"),
        bg=_ACCENT,
        fg="white",
        activebackground="#163A57",
        activeforeground="white",
        relief="flat",
        padx=18,
        pady=8,
        cursor="hand2",
    )
    select_btn.pack(pady=(0, 14))

    list_frame = tk.Frame(root, bg="white", highlightbackground="#CCCCCC", highlightthickness=1)
    list_frame.pack(padx=24, pady=(0, 14), fill="both", expand=True)

    list_label = tk.Label(list_frame, text="Selected Images:", font=("Segoe UI", 9, "bold"), bg="white", anchor="w")
    list_label.pack(fill="x", padx=8, pady=(6, 0))

    listbox = tk.Listbox(list_frame, font=("Segoe UI", 9), activestyle="none", selectmode="extended")
    listbox.pack(fill="both", expand=True, padx=8, pady=8)

    convert_btn = tk.Button(
        root,
        text="CONVERT TO EXCEL",
        font=("Segoe UI", 10, "bold"),
        bg="#2E7D32",
        fg="white",
        activebackground="#1B5E20",
        activeforeground="white",
        relief="flat",
        padx=18,
        pady=8,
        cursor="hand2",
        state="disabled",
    )
    convert_btn.pack(pady=(0, 10))

    status_frame = tk.Frame(root, bg=_BG)
    status_frame.pack(pady=(0, 16), fill="x", padx=24)

    status_caption = tk.Label(status_frame, text="Status:", font=("Segoe UI", 9, "bold"), bg=_BG, fg=_TEXT)
    status_caption.pack(anchor="w")

    status_label = tk.Label(status_frame, text="Ready", font=("Segoe UI", 9), bg=_BG, fg="#555555", anchor="w")
    status_label.pack(anchor="w")

    def set_status(text: str) -> None:
        status_label.config(text=text)
        root.update_idletasks()

    def refresh_listbox() -> None:
        listbox.delete(0, tk.END)
        for p in selected_paths:
            listbox.insert(tk.END, os.path.basename(p))
        convert_btn.config(state="normal" if selected_paths else "disabled")

    def on_select_images() -> None:
        paths = filedialog.askopenfilenames(
            title="Select JPG images",
            filetypes=[("JPEG images", "*.jpg *.jpeg"), ("All images", "*.jpg *.jpeg *.png"), ("All files", "*.*")],
        )
        if not paths:
            return
        selected_paths.clear()
        selected_paths.extend(paths)
        refresh_listbox()
        set_status(f"{len(selected_paths)} image(s) selected. Ready to convert.")

    def on_progress(index: int, total: int, filename: str) -> None:
        root.after(0, lambda: set_status(f"Processing {index} of {total}: {filename}"))

    def convert_worker() -> None:
        try:
            out_path, results = process_images_fn(list(selected_paths), progress_cb=on_progress)
        except ValueError as exc:
            root.after(0, lambda: _finish_with_error(str(exc)))
            return
        except RuntimeError as exc:
            root.after(0, lambda: _finish_with_error(str(exc)))
            return
        except Exception:
            root.after(0, lambda: _finish_with_error("Unable to create Excel file."))
            return

        failed = [r for r in results if not r.ok]

        def _finish_ok():
            set_status("Opening Excel...")
            try:
                excel_ready = excel_export.open_and_maximize_excel(out_path)
            except Exception:
                excel_ready = False

            if excel_ready:
                # Excel is confirmed open, visible, and maximized - the
                # converter's job is done. Close completely (no lingering
                # window, tray icon, or background thread/process) so
                # only Excel is left on screen.
                root.destroy()
                os._exit(0)
                return

            # Excel could not be confirmed open - never close silently.
            # The workbook is already saved either way.
            if failed:
                names = ", ".join(os.path.basename(r.path) for r in failed)
                set_status(f"Excel could not be opened. Some images had issues: {names}")
            else:
                set_status("Excel could not be opened. The exported file has been saved successfully.")
            messagebox.showwarning(
                "Excel did not open",
                "Excel could not be opened automatically.\n\n"
                f"The exported file has been saved successfully:\n{out_path}\n\n"
                "You can open it manually.",
            )
            if failed:
                names = ", ".join(os.path.basename(r.path) for r in failed)
                messagebox.showwarning(
                    "Partially completed",
                    f"Could not extract data from: {names}",
                )
            convert_btn.config(state="normal")
            select_btn.config(state="normal")

        root.after(0, _finish_ok)

    def _finish_with_error(message: str) -> None:
        set_status("Ready")
        convert_btn.config(state="normal")
        select_btn.config(state="normal")
        messagebox.showerror("Error", message)

    def on_convert() -> None:
        if not selected_paths:
            messagebox.showerror("Error", "Please select at least one image.")
            return
        convert_btn.config(state="disabled")
        select_btn.config(state="disabled")
        set_status("Starting conversion...")
        threading.Thread(target=convert_worker, daemon=True).start()

    select_btn.config(command=on_select_images)
    convert_btn.config(command=on_convert)

    root.mainloop()

"""
clean_pdf_fields.py
Flattens all form fields in every PDF in INPUT_DIR, then overwrites
each file in-place with the same name.

Usage:
    python clean_pdf_fields.py

Requirements:
    pip install pymupdf
"""

import shutil
from pathlib import Path
import fitz  # PyMuPDF

# ── Configuration ────────────────────────────────────────────────────────────
INPUT_DIR = r"C:\Users\horizon5\Desktop\New folder"
# ─────────────────────────────────────────────────────────────────────────────


def flatten_pdf(pdf_path: Path) -> None:
    tmp_path = pdf_path.with_suffix(".tmp.pdf")

    try:
        doc = fitz.open(str(pdf_path))

        for page in doc:
            # Flatten all widgets on the page into static content
            page.widgets()  # ensure widgets are loaded
            annots_to_flatten = [w for w in page.widgets()]
            for widget in annots_to_flatten:
                try:
                    page.delete_widget(widget)
                except Exception:
                    pass

        doc.save(str(tmp_path), garbage=4, deflate=True, no_new_id=True)
        doc.close()

        shutil.move(str(tmp_path), str(pdf_path))
        print(f"[OK]   {pdf_path.name}")

    except Exception as e:
        print(f"[WARN] Skipped {pdf_path.name}: {e}")
        if tmp_path.exists():
            tmp_path.unlink()


def main() -> None:
    folder = Path(INPUT_DIR)

    if not folder.exists():
        print(f"[ERROR] Directory not found: {INPUT_DIR}")
        return

    pdf_files = sorted(folder.glob("*.pdf"))

    if not pdf_files:
        print("[INFO] No PDF files found.")
        return

    print(f"[INFO] Flattening {len(pdf_files)} PDF(s) in:\n       {INPUT_DIR}\n")

    for pdf_path in pdf_files:
        flatten_pdf(pdf_path)

    print("\n[DONE] All files flattened in-place.")


if __name__ == "__main__":
    main()

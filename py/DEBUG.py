import fitz
from pathlib import Path
from tabulate import tabulate

input_dir = Path.home() / "Downloads"
pdf_files = input_dir.glob("*.pdf")
output_dir = Path.home() / "Desktop" / "Coordinates"
output_dir.mkdir(parents=True, exist_ok=True)


def get_text(doc, mode="blocks"):
    field_dict = {}
    for page_number in range(doc.page_count):
        page = doc[page_number]
        wlist = page.get_text(mode)
        field_dict[page_number + 1] = [
            [list(filter(None, item[4].split("\n"))), item[:4]] for item in wlist
        ]
    return field_dict


def write_txt_to_file(dir_path, field_dict):
    with open(dir_path, "w", encoding="utf-8") as file:
        for page, data in field_dict.items():
            file.write(f"Page: {page}\n")
            try:
                file.write(
                    f"{tabulate(data, ['Keywords', 'Coordinates'], tablefmt='grid', maxcolwidths=[50, None])}\n"
                )
            except IndexError:
                continue


def main():
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            write_txt_to_file(output_dir / f"(Block Coordinates) {pdf.stem}.txt", get_text(doc))
            write_txt_to_file(output_dir / f"(Word Coordinates) {pdf.stem}.txt", get_text(doc, "words"))


if __name__ == "__main__":
    main()
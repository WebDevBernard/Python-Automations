import os
from pathlib import Path

def unique_file_name(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while Path(path).is_file():
        path = f"{filename} ({counter}){extension}"
        counter += 1
    return path

def build_index(doc):

    page_index = {}
    text_to_location = []

    for page_num, page in enumerate(doc):
        blocks = []
        for block_idx, block in enumerate(page.get_text("blocks")):
            coords = tuple(block[:4])
            text_lines = block[4].split("\n")
            blocks.append({"words": text_lines, "coords": coords})

            for line_idx, line in enumerate(text_lines):
                text_to_location.append({
                    "normalized": line.lower(),
                    "page": page_num,
                    "block": block_idx,
                    "line": line_idx,
                    "text": line,
                    "coords": coords
                })

        page_index[page_num] = blocks

    return page_index, text_to_location
import fitz  # PyMuPDF
from pathlib import Path


def _gray_header_bands(page, drawings):
    bands = []
    for d in drawings:
        if not d.get("fill"):
            continue
        fill = d["fill"]
        r, g, b = fill[:3] if len(fill) >= 3 else (fill[0],) * 3
        if abs(r - g) < 0.05 and abs(g - b) < 0.05 and 0.5 < r < 0.95:
            bands.append(fitz.Rect(d["rect"]))
    return sorted(bands, key=lambda b: b.y0)


def _existing_vertical_lines(drawings, y0, y1, x0, x1):
    xs = []
    for d in drawings:
        if d.get("color") is None:
            continue
        for item in d["items"]:
            op = item[0]
            args = item[1:]
            if op == "l" and len(args) == 2:
                p1, p2 = args[0], args[1]
                if abs(p1.x - p2.x) < 0.5 and abs(p1.y - p2.y) > 3:
                    ly0, ly1 = sorted((p1.y, p2.y))
                    if x0 - 1 <= p1.x <= x1 + 1 and ly1 >= y0 - 1 and ly0 <= y1 + 1:
                        xs.append(p1.x)
            elif op == "re" and args and isinstance(args[0], fitz.Rect):
                r = args[0]
                if abs(r.width) < 1 and r.height > 3:
                    if x0 - 1 <= r.x0 <= x1 + 1 and r.y1 >= y0 - 1 and r.y0 <= y1 + 1:
                        xs.append(r.x0)
    return sorted(set(round(x, 1) for x in xs))


def _column_bounds_from_words(band_words, band_x0, band_x1, gap=10.0):
    words = sorted(band_words, key=lambda w: w[0])
    clusters = []
    for w in words:
        if clusters and w[0] <= clusters[-1][1] + gap:
            clusters[-1][1] = max(clusters[-1][1], w[2])
        else:
            clusters.append([w[0], w[2]])
    if not clusters:
        return None
    start = max(band_x0, clusters[0][0] - 20)
    end = min(band_x1, clusters[-1][1] + 20)
    bounds = [start]
    for i in range(len(clusters) - 1):
        bounds.append((clusters[i][1] + clusters[i + 1][0]) / 2)
    bounds.append(end)
    return bounds


def _resolve_boundary_collisions(col_bounds, words, min_gap=1.0):
    col_bounds = col_bounds[:]
    for i in range(1, len(col_bounds) - 1):
        x = col_bounds[i]
        for w in words:
            x0, y0, x1, y1 = w[0], w[1], w[2], w[3]
            if x0 < x < x1:
                dist_left = x - x0
                dist_right = x1 - x
                new_x = (x0 - min_gap) if dist_left < dist_right else (x1 + min_gap)
                new_x = max(new_x, col_bounds[i - 1] + min_gap)
                new_x = min(new_x, col_bounds[i + 1] - min_gap)
                col_bounds[i] = new_x
                x = new_x
    return col_bounds


def find_and_draw_table(
    pdf_path,
    output_image,
    keywords,
    page_num=0,
    row_tolerance=3,
    zoom=2.0,
    top_header_keywords=None,
    stop_keywords=None,
):
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    drawings = page.get_drawings()
    words = page.get_text("words")

    header_keywords = list(top_header_keywords or []) + list(keywords)

    kw_rects = {}
    for kw in header_keywords:
        matches = page.search_for(kw)
        if matches:
            kw_rects[kw] = matches[0]
        else:
            print(f"  Warning: keyword '{kw}' not found")

    if len(kw_rects) < 2:
        raise ValueError("Need at least 2 matched headings")

    stop_terms = stop_keywords or ["Outstanding Balance", "Customer Original"]
    stop_y = None
    for term in stop_terms:
        matches = page.search_for(term)
        if matches:
            candidate = max(r.y1 for r in matches) + 4
            stop_y = candidate if stop_y is None else min(stop_y, candidate)

    table_bottom = stop_y if stop_y is not None else page.rect.height - 36

    bands = _gray_header_bands(page, drawings)
    sections = []
    for rect in kw_rects.values():
        for b in bands:
            if b.y0 - 3 <= rect.y0 and rect.y1 <= b.y1 + 3:
                sec = next((s for s in sections if s["band"] == b), None)
                if sec is None:
                    sections.append({"band": b, "rects": [rect]})
                else:
                    sec["rects"].append(rect)
                break

    if not sections:
        rects = list(kw_rects.values())
        y0 = min(r.y0 for r in rects) - 3
        y1 = max(r.y1 for r in rects)
        x0 = min(r.x0 for r in rects)
        x1 = max(r.x1 for r in rects)
        sections.append({"band": fitz.Rect(x0, y0, x1, y1), "rects": rects})

    sections.sort(key=lambda s: s["band"].y0)

    shape = page.new_shape()

    for i, sec in enumerate(sections):
        band = sec["band"]
        sec_top = band.y0

        if i + 1 < len(sections):
            limit = sections[i + 1]["band"].y0
        else:
            limit = table_bottom

        body_words = [w for w in words if w[1] >= band.y1 + 2 and w[3] <= limit]
        if body_words:
            sec_bottom = min(max(w[3] for w in body_words) + 3, limit)
        else:
            sec_bottom = band.y1

        vlines = _existing_vertical_lines(drawings, sec_top, sec_bottom, band.x0, band.x1)
        if vlines:
            col_bounds = sorted(set(c for c in [band.x0] + vlines + [band.x1] if band.x0 <= c <= band.x1))
        else:
            band_words = [
                w
                for w in words
                if band.y0 - 2 <= w[1] and w[3] <= band.y1 + 2 and band.x0 - 2 <= w[0] and w[2] <= band.x1 + 2
            ]
            col_bounds = _column_bounds_from_words(band_words, band.x0, band.x1)
            if col_bounds is None:
                rects = sorted(sec["rects"], key=lambda r: r.x0)
                col_bounds = [rects[0].x0 - 20]
                for j in range(len(rects) - 1):
                    col_bounds.append((rects[j].x1 + rects[j + 1].x0) / 2)
                col_bounds.append(rects[-1].x1 + 20)
                col_bounds[0] = max(col_bounds[0], band.x0)
                col_bounds[-1] = min(col_bounds[-1], band.x1)

        section_words = [
            w
            for w in words
            if sec_top - 2 <= w[1] and w[3] <= sec_bottom + 2 and band.x0 - 2 <= w[0] and w[2] <= band.x1 + 2
        ]
        col_bounds = _resolve_boundary_collisions(col_bounds, section_words)

        rows_y = sorted({round(w[1]) for w in body_words})
        grouped = []
        for y in rows_y:
            if not grouped or y - grouped[-1] > row_tolerance:
                grouped.append(y)

        row_lines = [sec_top, band.y1] + [y - 3 for y in grouped] + [sec_bottom]
        row_lines = sorted(set(row_lines))
        deduped = [row_lines[0]]
        for y in row_lines[1:]:
            if y - deduped[-1] > 4:
                deduped.append(y)
        row_lines = deduped

        for x in col_bounds:
            shape.draw_line((x, sec_top), (x, sec_bottom))
        for y in row_lines:
            shape.draw_line((col_bounds[0], y), (col_bounds[-1], y))

    shape.finish(color=(1, 0, 0), width=0.8)
    shape.commit()

    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(output_image)

    doc.close()
    print(f"  Saved annotated image to {output_image}")


if __name__ == "__main__":
    downloads = Path.home() / "Downloads"

    pdf_files = sorted(downloads.rglob("*.pdf"))

    if not pdf_files:
        print(f"No PDFs found in {downloads} or its subfolders")
    else:
        input_pdf = pdf_files[0]
        print(f"Using: {input_pdf}")

        output_image = downloads / f"{input_pdf.stem}_gridded.png"

        with fitz.open(input_pdf) as doc:
            first_page_text = doc[0].get_text()

        invoice_stop = "Transaction Amount:"
        if invoice_stop in first_page_text:
            keywords = [
                "Transaction",
                "Due",
                "Description",
                "Taxe",
                "Tax Amount",
                "Amount",
            ]
            top_header_keywords = ["Company Name", "Policy", "From", "To"]
            stop_keywords = [invoice_stop]
        else:
            keywords = [
                "Date",
                "Transaction",
                "Invoice",
                "Policy",
                "Description",
                "Amount",
            ]
            top_header_keywords = None
            stop_keywords = None

        find_and_draw_table(
            pdf_path=str(input_pdf),
            output_image=str(output_image),
            keywords=keywords,
            page_num=0,
            top_header_keywords=top_header_keywords,
            stop_keywords=stop_keywords,
        )
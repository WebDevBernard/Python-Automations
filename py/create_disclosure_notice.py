import math
import re
import time
import traceback
from datetime import datetime
from pathlib import Path

import fitz
import pandas as pd

from utils import (
    _split_mailing_address,
    address_one_title_case,
    address_two_title_case,
    format_amount,
    load_producer_mapping,
    parse_date,
    parse_date_to_dt,
    smart_title,
    to_float,
    write_to_new_docx,
)
from constants import REGEX_PATTERNS, get_insurer

DATE_FORMAT = "%B %d, %Y"

TABLE_KEYWORDS = ["Transaction", "Due", "Description", "% Taxe", "Tax Amount", "Amount"]

_STOP_TERMS = ["Transaction Amount", "Outstanding Balance", "Customer Original"]
_EXCLUDE_TERMS = ["balance forward", "outstanding balance", "customer original"]

_POLICY_NUMBER_RE = re.compile(r"\b(?:[A-Z]{1,2}\d{6,}[A-Z]{0,4}|\d{8,})\b")


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


def _header_columns(band_words, band_x0, band_x1, gap=10.0):
    words = sorted(band_words, key=lambda w: w[0])
    clusters = []
    for w in words:
        if clusters and w[0] <= clusters[-1][1] + gap:
            clusters[-1][1] = max(clusters[-1][1], w[2])
            clusters[-1][2].append(w)
        else:
            clusters.append([w[0], w[2], [w]])
    if not clusters:
        return None, None

    col_names = []
    for c in clusters:
        c[2].sort(key=lambda w: (round(w[1]), w[0]))
        col_names.append(" ".join(w[4] for w in c[2]))

    start = max(band_x0, clusters[0][0] - 20)
    end = min(band_x1, clusters[-1][1] + 20)
    col_bounds = [start]
    for i in range(len(clusters) - 1):
        col_bounds.append((clusters[i][1] + clusters[i + 1][0]) / 2)
    col_bounds.append(end)

    return col_names, col_bounds


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


def _header_geometry(page, drawings, keywords):
    kw_rects = {}
    for kw in keywords:
        matches = page.search_for(kw)
        if matches:
            kw_rects[kw] = matches[0]

    if len(kw_rects) < 2:
        return None

    rects = list(kw_rects.values())
    words = page.get_text("words")
    bands = _gray_header_bands(page, drawings)

    header_band = None
    best = 0
    for b in bands:
        contained = sum(1 for r in rects if b.y0 - 3 <= r.y0 and r.y1 <= b.y1 + 3)
        if contained > best:
            best = contained
            header_band = b

    if header_band is None:
        header_band = fitz.Rect(
            min(r.x0 for r in rects) - 20,
            min(r.y0 for r in rects) - 3,
            max(r.x1 for r in rects) + 20,
            max(r.y1 for r in rects),
        )

    header_bottom = header_band.y1

    band_words = [
        w
        for w in words
        if header_band.y0 - 2 <= w[1]
        and w[3] <= header_bottom + 2
        and header_band.x0 - 2 <= w[0]
        and w[2] <= header_band.x1 + 2
    ]

    col_names, col_bounds = _header_columns(
        band_words, header_band.x0, header_band.x1
    )

    if col_names is None:
        ordered = sorted(kw_rects.items(), key=lambda kv: kv[1].x0)
        col_names = [kw for kw, _ in ordered]
        ordered_rects = [r for _, r in ordered]
        col_bounds = [header_band.x0]
        for i in range(len(ordered_rects) - 1):
            col_bounds.append(
                (ordered_rects[i].x1 + ordered_rects[i + 1].x0) / 2
            )
        col_bounds.append(header_band.x1)
        col_bounds[0] = max(col_bounds[0], header_band.x0)
        col_bounds[-1] = min(col_bounds[-1], header_band.x1)

    n_cols = len(col_names)

    col_bounds = _resolve_boundary_collisions(col_bounds, band_words)

    return {
        "col_names": col_names,
        "rects": rects,
        "n_cols": n_cols,
        "header_bottom": header_bottom,
        "col_bounds": col_bounds,
    }


def _words_to_rows(
    body_words, col_bounds, n_cols, header_bottom_ref=None, row_tolerance=3
):
    if not body_words:
        return []

    rows_y = sorted({round(w[1]) for w in body_words})
    grouped = []
    for y in rows_y:
        if not grouped or y - grouped[-1] > row_tolerance:
            grouped.append(y)

    start_ref = header_bottom_ref if header_bottom_ref is not None else grouped[0] - 3
    row_bounds_y = (
        [start_ref] + [y - 3 for y in grouped] + [max(w[3] for w in body_words) + 4]
    )
    row_bounds_y = sorted(set(row_bounds_y))

    deduped = [row_bounds_y[0]]
    for y in row_bounds_y[1:]:
        if y - deduped[-1] > 2:
            deduped.append(y)
    row_bounds_y = deduped

    table_data = []
    for y_top, y_bottom in zip(row_bounds_y[:-1], row_bounds_y[1:]):
        row_cells = ["" for _ in range(n_cols)]
        for w in body_words:
            x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
            wcy = (y0 + y1) / 2
            if not (y_top <= wcy < y_bottom):
                continue
            wcx = (x0 + x1) / 2
            for i in range(n_cols):
                if col_bounds[i] <= wcx < col_bounds[i + 1]:
                    row_cells[i] = (row_cells[i] + " " + text).strip()
                    break
        if any(row_cells):
            table_data.append(row_cells)

    return table_data


def extract_table_from_pdf(pdf_path, keywords, header_page=0):
    doc = fitz.open(pdf_path)
    try:
        if header_page >= len(doc):
            return None

        first_page = doc[header_page]
        info = _header_geometry(first_page, first_page.get_drawings(), keywords)
        if info is None:
            print(f"  Skipping {Path(pdf_path).name}: fewer than 2 headers matched")
            return None

        col_names = info["col_names"]
        n_cols = info["n_cols"]
        col_bounds = info["col_bounds"]

        page_word_sets = []
        y_start = info["header_bottom"] + 2

        for pnum in range(header_page, len(doc)):
            page = doc[pnum]
            page_info = _header_geometry(page, page.get_drawings(), keywords)

            ref = None
            if page_info is not None:
                y_start = page_info["header_bottom"] + 2
                ref = y_start

            table_bottom = page.rect.height - 36
            for term in _STOP_TERMS:
                matches = page.search_for(term)
                if matches:
                    table_bottom = min(table_bottom, max(r.y0 for r in matches))

            words = page.get_text("words")
            body_words = [w for w in words if w[1] >= y_start and w[3] <= table_bottom]
            page_word_sets.append((body_words, ref))
            y_start = 20
    finally:
        doc.close()

    all_body_words = [w for body_words, _ in page_word_sets for w in body_words]
    if all_body_words:
        refined = _column_bounds_from_words(
            all_body_words, col_bounds[0], col_bounds[-1]
        )
        if refined is not None and len(refined) == n_cols + 1:
            col_bounds = refined

    all_table_data = []
    for body_words, ref in page_word_sets:
        all_table_data.extend(_words_to_rows(body_words, col_bounds, n_cols, ref))

    if not all_table_data:
        print(f"  Skipping {Path(pdf_path).name}: no data rows extracted")
        return None

    df = pd.DataFrame(all_table_data, columns=col_names)

    mask = ~df.apply(
        lambda row: any(
            term in " ".join(str(v) for v in row).lower() for term in _EXCLUDE_TERMS
        ),
        axis=1,
    )
    df = df[mask].reset_index(drop=True)

    return df


# Address lines start with a street number, "Unit", or "PO Box" (same rule as
# the auto renewal letter). Postal codes mark the end of the address.
_ADDRESS_RE = REGEX_PATTERNS["address"]
_POSTAL_RE = REGEX_PATTERNS["postal_code"]

# Producer line, e.g. "EPS530 BY NT" or "BCS358 BY NT NT".
_PRODUCER_RE = re.compile(r"^\S+\s+BY\s+\S+", re.IGNORECASE)


def _is_label_line(line, label):
    """True if the line is a '<label>:' label (colon required, so bare table
    headers like 'From' inside 'Term From To' are not treated as labels)."""
    return bool(
        re.match(rf"^\s*{label}\s*:\s*$", line, re.IGNORECASE)
        or re.match(rf"^\s*{label}\s*:\s*(.+)$", line, re.IGNORECASE)
    )


def extract_label_block(page, label):
    """Extract all text lines following '<label>:' on a page.

    Collects lines from the label through the address, which ends at the
    postal code. Only blocks in the same column as the label are considered,
    so a side-by-side column (e.g. invoice details) is not mixed in. The
    caller splits the result into name / address / producer using regex rules.
    """
    blocks = sorted(page.get_text("blocks"), key=lambda b: (b[1], b[0]))

    for i, block in enumerate(blocks):
        lines = [ln.strip() for ln in block[4].split("\n")]
        label_index = -1
        for idx, ln in enumerate(lines):
            if _is_label_line(ln, label):
                label_index = idx
                break
        if label_index == -1:
            continue

        label_x0 = block[0]

        content = []
        for ln in lines[label_index:]:
            if re.match(rf"^\s*{label}\s*:?\s*$", ln, re.IGNORECASE):
                continue
            m = re.match(rf"^\s*{label}\s*:\s*(.*)$", ln, re.IGNORECASE)
            if m is not None:
                rest = m.group(1).strip()
                if rest:
                    content.append(rest)
                continue
            if ln:
                content.append(ln)
            if _POSTAL_RE.search(ln):
                break

        # Content continues in the following blocks (same column) until the
        # postal code is reached.
        for nxt in blocks[i + 1 :]:
            if nxt[0] > label_x0 + 200:
                continue
            nxt_lines = [l.strip() for l in nxt[4].split("\n") if l.strip()]
            if not nxt_lines:
                continue
            for ln in nxt_lines:
                content.append(ln)
                if _POSTAL_RE.search(ln):
                    break
            if content and _POSTAL_RE.search(content[-1]):
                break

        return content if content else None

    return None


def extract_block(doc, label):
    """Find a '<label>:' block on a page and return its lines."""
    for page in doc:
        lines = extract_label_block(page, label)
        if lines:
            return lines
    return None


def extract_from_block(doc):
    """Find the 'From:' block used for the name."""
    return extract_block(doc, "From")


def extract_to_block(doc):
    """Find the 'To:' block used for the address."""
    return extract_block(doc, "To")


def extract_policy_number(doc):
    """Find a policy number (letter + 6+ digits) in the statement text."""
    for page in doc:
        m = _POLICY_NUMBER_RE.search(page.get_text("text"))
        if m:
            return m.group(0)
    return ""


def extract_insurer(doc):
    """Extract the insurer name from under the 'Company Name' field.

    When the field holds the broker (Horizon West) instead of an actual
    insurer, an empty string is returned so the caller falls back to
    get_insurer() and matches the insurer from the policy number.
    """
    for page in doc:
        matches = page.search_for("Company Name")
        if not matches:
            continue
        r = matches[0]
        clip = fitz.Rect(r.x0 - 10, r.y1, r.x0 + 180, r.y1 + 25)
        text = page.get_text("text", clip=clip).strip()
        if text:
            name = smart_title(" ".join(text.split()))
            if "horizon west" in name.lower():
                return ""
            return name
    return ""


def extract_producer_line(doc):
    """Find the producer line (e.g. 'CHES01 BY NT') anywhere in the statement."""
    for page in doc:
        for ln in page.get_text("text").split("\n"):
            ln = ln.strip()
            if _PRODUCER_RE.match(ln):
                return ln
    return None


def extract_producer_code(producer_line):
    """Producer code is the 3rd whitespace-separated token in the producer line.

    Handles short lines like "EPS530 BY NT" (3 tokens) and even "BY NT" (2).
    """
    if not producer_line:
        return None
    parts = producer_line.split()
    if len(parts) >= 3:
        return parts[2]
    if len(parts) == 2:
        return parts[1]
    return None


def _split_block_lines(lines):
    """Split raw 'To:'/'From:' lines into (name, address, producer_line).

    - Address starts at the first line that looks like a street address
      (begins with a number, 'Unit', or 'PO Box' - same rule as the auto
      renewal letter), so extra name lines before it are kept as the name.
    - The producer line (e.g. 'EPS530 BY NT') ends the address block.
    """
    producer_line = None
    producer_index = -1
    for i in range(len(lines) - 1, -1, -1):
        if _PRODUCER_RE.match(lines[i]):
            producer_index = i
            producer_line = lines[i]
            break

    body = lines[:producer_index] if producer_index != -1 else lines

    address_start = -1
    for i, ln in enumerate(body):
        if _ADDRESS_RE.search(ln):
            address_start = i
            break

    if address_start == -1:
        return body[:1], body[1:], producer_line

    # Address ends at the line containing the postal code (inclusive), matching
    # the auto renewal letter; otherwise it runs to the end of the block.
    address_end = len(body)
    for i in range(address_start, len(body)):
        if _POSTAL_RE.search(body[i]):
            address_end = i + 1
            break

    return body[:address_start], body[address_start:address_end], producer_line


def _build_block(lines):
    name_lines, address_lines, producer_line = _split_block_lines(lines)
    return {
        "name": " ".join(name_lines) if name_lines else "",
        "address": [l for l in address_lines if l],
        "producer_line": producer_line,
    }


def parse_blocks(from_lines):
    """Build the name/address dict from the 'From:' block."""
    if not from_lines:
        return None, "Could not find a 'From:' name/address block in the statement."

    from_block = _build_block(from_lines)
    return from_block, "Using 'From:' block."


def _find_column(df, keyword):
    for col in df.columns:
        if keyword in str(col).lower():
            return col
    return None


def _find_amount_column(df):
    for col in df.columns:
        if str(col).strip().lower() == "amount":
            return col
    return _find_column(df, "amount")


def _earliest_date(frame, date_col):
    if not date_col:
        return None
    dates = []
    for value in frame[date_col]:
        dt = parse_date_to_dt(value)
        if dt:
            dates.append(dt)
    return min(dates) if dates else None


def _policy_row(policy_number, prem_total, effective_dt):
    return {
        "policy_number": policy_number,
        "prem_amt": format_amount(prem_total) if prem_total else "",
        "effective_date": (
            parse_date(effective_dt)
            if effective_dt
            else datetime.today().strftime(DATE_FORMAT)
        ),
        "insurer": get_insurer(policy_number),
    }


def build_policy_rows(df):
    """Build a single disclosure-notice row for the policy with the largest
    premium. Rows sharing a policy number are summed together first. When no
    policy column exists (e.g. an invoice), all amounts are treated as one
    policy and summed into a single row."""
    if df is None or df.empty:
        return None

    policy_col = _find_column(df, "policy")
    date_col = _find_column(df, "date") or _find_column(df, "due")
    amount_col = _find_amount_column(df)

    if policy_col:
        candidates = []
        for policy, group in df.groupby(policy_col, dropna=False):
            if isinstance(policy, float) and math.isnan(policy):
                continue
            policy_number = str(policy).strip()
            if not policy_number:
                continue
            prem_total = (
                group[amount_col].apply(to_float).sum() if amount_col else 0.0
            )
            effective_dt = _earliest_date(group, date_col)
            candidates.append(
                (prem_total, _policy_row(policy_number, prem_total, effective_dt))
            )
        if not candidates:
            return None
        candidates.sort(key=lambda item: item[0], reverse=True)
        return [candidates[0][1]]

    prem_total = df[amount_col].apply(to_float).sum() if amount_col else 0.0
    effective_dt = _earliest_date(df, date_col)
    return [_policy_row("", prem_total, effective_dt)]


def generate_letters(pdf_path, output_dir, producer_mapping=None):
    """Extract name/address + transactions from a statement and write letters."""
    producer_mapping = producer_mapping or {}

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    from_lines = extract_from_block(doc)
    to_lines = extract_to_block(doc)
    policy_number = extract_policy_number(doc)
    insurer = extract_insurer(doc)
    producer_line = extract_producer_line(doc)
    doc.close()

    block, message = parse_blocks(from_lines)
    if not block:
        return 0

    named_insured = smart_title(block["name"])

    to_block = parse_blocks(to_lines)[0] if to_lines else None
    address_block = to_block or block
    mailing_address = _split_mailing_address(", ".join(address_block["address"]))
    address_parts = mailing_address.split("\n")
    if len(address_parts) >= 2:
        mailing_address = "\n".join(
            [address_one_title_case(address_parts[0])]
            + [address_two_title_case(part) for part in address_parts[1:-1]]
            + [address_parts[-1]]
        )

    template_name = (
        "Strata Disclosure Notice"
        if "strata plan" in named_insured.lower()
        else "Disclosure Notice"
    )
    template_path = Path.cwd() / "assets" / f"{template_name}.docx"

    producer_code = extract_producer_code(producer_line or block.get("producer_line"))
    producer_name = (
        producer_mapping.get(producer_code.lower(), "") if producer_code else ""
    )

    df = extract_table_from_pdf(pdf_path, TABLE_KEYWORDS)
    rows = build_policy_rows(df)

    if not rows:
        rows = [
            {
                "policy_number": "",
                "prem_amt": "",
                "effective_date": datetime.today().strftime(DATE_FORMAT),
                "insurer": "",
            }
        ]
        print(
            "Warning: no transaction rows extracted; generating a single letter "
            "from name/address only."
        )

    success = 0
    for row in rows:
        row["named_insured"] = named_insured
        row["mailing_address"] = mailing_address
        row["producer_name"] = producer_name
        if policy_number:
            row["policy_number"] = policy_number
        row["insurer"] = insurer or get_insurer(policy_number)
        suffix = template_name
        try:
            if write_to_new_docx(
                template_path=template_path,
                data=row,
                output_dir=output_dir,
                output_suffix=suffix,
            ):
                success += 1
        except Exception as e:
            print(f"Failed for {row.get('policy_number', 'unknown')}: {e}")
            traceback.print_exc()

    return success


def _exit_countdown(seconds=3):
    print("\nExiting in ", end="")
    for i in range(seconds, 0, -1):
        print(f"{i} ", end="", flush=True)
        time.sleep(1)
    print()


def _find_statements(downloads_dir, count=2):
    pdfs = sorted(
        downloads_dir.glob("*.pdf"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return pdfs[:count]


def create_disclosure_notice(config_data=None):
    downloads_dir = Path.home() / "Downloads"
    pdf_paths = _find_statements(downloads_dir)
    if not pdf_paths:
        print("No PDF files found in Downloads folder.")
        _exit_countdown()
        return

    output_dir = Path.home() / "Desktop"
    producer_mapping = load_producer_mapping("config.xlsx")

    success = 0
    for pdf_path in pdf_paths:
        success += generate_letters(pdf_path, output_dir, producer_mapping)

    print(
        f"******** Disclosure Letter from Invoice or Statement ran successfully: "
        f"{success} letter(s) generated ********"
    )
    _exit_countdown()


if __name__ == "__main__":
    create_disclosure_notice()

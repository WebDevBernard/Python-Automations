import fitz  # PyMuPDF
import pandas as pd
import re
import sys
import time
import os
from pathlib import Path
from itertools import combinations


# ---------------- Setup ----------------
POLICY_PATTERN = re.compile(r"[A-Z0-9]{6,}")
DEBUG = False
DRAW_TABLE = False
DRAW_COLUMNS = False
MIN_ROW_FREQUENCY = 0.3
X_ROUNDING = 10
COLUMN_MERGE_DISTANCE = 40


# ---------------- Utility Functions ----------------
def progressbar(it, prefix="", size=60, out=sys.stdout):
    """Display a progress bar with time estimation."""
    count = len(it)
    start = time.time()

    def show(j):
        x = int(size * j / count)
        remaining = ((time.time() - start) / j) * (count - j) if j else 0
        mins, sec = divmod(remaining, 60)
        time_str = f"{int(mins):02}:{sec:03.1f}"
        print(
            f"{prefix}[{'█' * x}{'.' * (size - x)}] {j}/{count} Est wait {time_str}",
            end="\r",
            file=out,
            flush=True,
        )

    if len(it) > 0:
        show(0.1)
        for i, item in enumerate(it):
            yield item
            show(i + 1)
        print(flush=True, file=out)


def safe_filename(name: str) -> str:
    """Remove invalid filename characters."""
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def unique_file_name(path: str) -> str:
    """Generate a unique filename by appending (n) if file exists."""
    directory = os.path.dirname(path)
    filename, extension = os.path.splitext(os.path.basename(path))
    filename = safe_filename(filename)
    base_name = re.sub(r"\s*\(\d+\)$", "", filename)
    counter = 1
    new_path = os.path.join(directory, f"{base_name}{extension}")
    while Path(new_path).is_file():
        new_path = os.path.join(directory, f"{base_name} ({counter}){extension}")
        counter += 1
    return new_path


# ---------------- Enhanced Table Extraction ----------------
def detect_columns(words):
    """Find columns by looking for x-positions that repeat across many rows."""
    if not words:
        return []

    rows = []
    for w in sorted(words, key=lambda x: x[1]):
        y_center = (w[1] + w[3]) / 2
        added = False
        for row in rows:
            if abs(y_center - row["y"]) <= 4:
                row["words"].append(w)
                row["y"] = sum((ww[1] + ww[3]) / 2 for ww in row["words"]) / len(
                    row["words"]
                )
                added = True
                break
        if not added:
            rows.append({"y": y_center, "words": [w]})

    x_counts = {}
    for row in rows:
        seen_x = set()
        for w in row["words"]:
            x = round(w[0] / X_ROUNDING) * X_ROUNDING
            seen_x.add(x)
        for x in seen_x:
            x_counts[x] = x_counts.get(x, 0) + 1

    threshold = max(2, len(rows) * MIN_ROW_FREQUENCY)
    columns = sorted([x for x, count in x_counts.items() if count >= threshold])

    merged = []
    for x in columns:
        if not merged or (x - merged[-1]) >= COLUMN_MERGE_DISTANCE:
            merged.append(x)
        else:
            merged[-1] = (merged[-1] + x) / 2

    return merged


def build_table(words, columns):
    """Assign words to columns and build table rows."""
    rows = []
    for w in sorted(words, key=lambda x: x[1]):
        y_center = (w[1] + w[3]) / 2
        added = False
        for row in rows:
            if abs(y_center - row["y"]) <= 4:
                row["words"].append(w)
                added = True
                break
        if not added:
            rows.append({"y": y_center, "words": [w]})

    table = []
    for row in sorted(rows, key=lambda r: r["y"]):
        col_data = {i: [] for i in range(len(columns))}
        for w in row["words"]:
            distances = [abs(w[0] - col) for col in columns]
            closest = distances.index(min(distances))
            if distances[closest] < 40:
                col_data[closest].append(w[4])
        table.append([" ".join(col_data[i]) for i in range(len(columns))])

    return table


def extract_table_from_bbox(page, bbox):
    """Extract table using enhanced column detection."""
    x0, y0, x1, y1 = bbox
    words = page.get_text("words")
    table_words = [
        w for w in words if x0 - 1 <= w[0] <= x1 + 1 and y0 - 1 <= w[1] <= y1 + 1
    ]

    if not table_words:
        return pd.DataFrame(), []

    columns = detect_columns(table_words)
    if not columns:
        return pd.DataFrame(), []

    if DEBUG:
        print(f"  Detected {len(columns)} columns at x={[int(x) for x in columns]}")

    table_data = build_table(table_words, columns)

    if not table_data or len(table_data) < 2:
        return pd.DataFrame(), columns

    headers = table_data[0]
    headers = [str(h).strip() if h else f"Col_{i}" for i, h in enumerate(headers)]

    seen = {}
    unique_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)

    try:
        df = pd.DataFrame(table_data[1:], columns=unique_headers)
        return df, columns
    except Exception as e:
        if DEBUG:
            print(f"  Error creating DataFrame: {e}")
        return pd.DataFrame(), columns


# ---------------- PDF Column Finders ----------------
def find_policy_column(df, min_unique=2):
    """Find the column containing policy numbers."""
    if df.empty or len(df.columns) == 0:
        return None

    best_col, max_unique = None, 0
    for col in df.columns:
        try:
            series = df[col]
            if not isinstance(series, pd.Series):
                continue

            def extract_policy(text):
                text = str(text).strip()
                match = POLICY_PATTERN.search(text)
                return match.group(0) if match else None

            policy_numbers = series.apply(extract_policy)
            unique_matches = policy_numbers.dropna().nunique()

            if unique_matches >= min_unique and unique_matches > max_unique:
                best_col, max_unique = col, unique_matches
        except Exception as e:
            if DEBUG:
                print(f"  Error checking column '{col}': {e}")
            continue

    return df.columns.get_loc(best_col) if best_col else None


def find_premium_column(df):
    """Find the column containing premium amounts."""
    if df.empty or len(df.columns) == 0:
        return None

    candidates = {}
    for col in df.columns:
        try:
            series = df[col]
            if not isinstance(series, pd.Series):
                continue

            col_clean = (
                series.astype(str).str.strip().str.replace(r"[%\s$,]", "", regex=True)
            )
            col_numeric = pd.to_numeric(col_clean, errors="coerce")
            col_numeric = col_numeric.where(col_clean.str.endswith(".00"))
            count = col_numeric.notna().sum()
            if count > 0:
                candidates[col] = {"count": count, "max_value": col_numeric.max()}
        except Exception as e:
            if DEBUG:
                print(f"  Error checking column '{col}': {e}")
            continue

    if not candidates:
        return None

    return df.columns.get_loc(
        max(candidates.items(), key=lambda x: x[1]["max_value"])[0]
    )


# ---------------- PDF Data Extraction ----------------
def is_intact_pdf(doc):
    """Check if PDF contains 'intact' (case-insensitive)."""
    for page in doc:
        text = page.get_text().lower()
        if "intact" in text:
            return True
    return False


def extract_policies_and_premiums(pdf_path):
    """Extract policies and premiums using enhanced table detection."""
    all_data = []
    doc = fitz.open(pdf_path)
    is_intact = is_intact_pdf(doc)

    if DEBUG and is_intact:
        print(f"  Detected Intact insurance PDF")

    for page_num, page in enumerate(doc):
        table_finder = page.find_tables(strategy="text")
        detected_tables = list(table_finder) if table_finder else []

        if DEBUG:
            print(f"Page {page_num + 1}: Found {len(detected_tables)} table regions")

        for table in detected_tables:
            df, columns = extract_table_from_bbox(page, table.bbox)

            if df.empty:
                continue

            if DEBUG:
                print(f"  Extracted DataFrame shape: {df.shape}")
                print(f"  Column names: {list(df.columns)}")

            policy_idx = find_policy_column(df)
            premium_idx = find_premium_column(df)

            if policy_idx is None or premium_idx is None:
                if DEBUG:
                    print(f"  Could not find policy/premium columns")
                continue

            if DEBUG:
                print(f"  ✓ Policy column index: {policy_idx}")
                print(f"  ✓ Premium column index: {premium_idx}")

            def clean_policy(text):
                original_text = str(text).strip()
                text = original_text

                if is_intact:
                    text = text.replace(" ", "")
                    if len(text) > 3 and text[:3].isalpha():
                        text = text[3:]

                match = POLICY_PATTERN.search(text)
                if not match:
                    return None

                policy = match.group(0)

                if is_intact and policy.endswith("H"):
                    policy = policy[:-1]

                return policy

            df = df.rename(
                columns={
                    df.columns[policy_idx]: "policy_number",
                    df.columns[premium_idx]: "premium",
                }
            )[["policy_number", "premium"]]

            df["policy_number"] = df["policy_number"].apply(clean_policy)
            df = df.dropna()

            df["premium"] = pd.to_numeric(
                df["premium"].astype(str).str.replace(r"[$,]", "", regex=True),
                errors="coerce",
            ).fillna(0)

            all_data.extend(
                [
                    {
                        "pdf": pdf_path.name,
                        "page": page_num + 1,
                        "policy_number": row["policy_number"],
                        "premium": row["premium"],
                    }
                    for _, row in df.iterrows()
                ]
            )

    doc.close()
    return pd.DataFrame(all_data)


# ---------------- Matching PDFs ----------------
def find_matching_pdfs_by_policy(folder):
    """Find PDFs that share policy numbers."""
    pdf_files = list(folder.glob("*.pdf"))
    if not pdf_files:
        return []

    pdf_policies = {}
    for pdf in pdf_files:
        df = extract_policies_and_premiums(pdf)
        if not df.empty:
            pdf_policies[pdf] = set(df["policy_number"].unique())

    return [
        (p1, p2)
        for p1, p2 in combinations(pdf_policies.keys(), 2)
        if pdf_policies[p1] & pdf_policies[p2]
    ]


# ---------------- Aggregation & Balancing ----------------
def aggregate_pdfs(pdf_paths):
    """Aggregate policies and premiums by policy number."""
    aggregates = {}
    for pdf_path in pdf_paths:
        df = extract_policies_and_premiums(pdf_path)
        if df.empty:
            print(f"No data found in {pdf_path.name}.")
            continue
        df_agg = df.groupby("policy_number", as_index=False)["premium"].sum()
        df_agg["balanced"] = False
        aggregates[pdf_path.name] = df_agg
    return aggregates


def update_balanced_status(df1, df2):
    """Mark policies as balanced if premiums match between PDFs."""
    map1, map2 = (
        df1.set_index("policy_number")["premium"],
        df2.set_index("policy_number")["premium"],
    )
    df1["balanced"] |= df1["policy_number"].map(map2) == df1["premium"]
    df2["balanced"] |= df2["policy_number"].map(map1) == df2["premium"]


# ---------------- PDF Highlighting ----------------
def draw_debug_lines(page, bbox, columns):
    """Draw column lines on PDF for visualization."""
    y0, y1 = bbox[1], bbox[3]
    for x in columns:
        line = fitz.Rect(x, y0, x + 1, y1)
        annot = page.add_rect_annot(line)
        annot.set_colors(stroke=(0, 0, 1))
        annot.set_border(width=0.5)
        annot.update()


def highlight_unbalanced_policies(pdf_path, df, output_folder):
    """Highlight unbalanced policies in the PDF."""
    unbalanced = set(df[(~df["balanced"]) & (df["premium"] > 0)]["policy_number"])

    doc = fitz.open(pdf_path)
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)

        table_finder = page.find_tables(strategy="text")
        tables = list(table_finder) if table_finder else []

        for table in tables:
            if DRAW_TABLE:
                rect = fitz.Rect(*table.bbox)
                annot = page.add_rect_annot(rect)
                annot.set_colors(stroke=(1, 0, 0))
                annot.set_border(width=1)
                annot.update()

            if DRAW_COLUMNS:
                df_table, columns = extract_table_from_bbox(page, table.bbox)
                if columns:
                    draw_debug_lines(page, table.bbox, columns)

        if unbalanced:
            for w in page.get_text("words"):
                word_text = w[4].strip()
                match = POLICY_PATTERN.search(word_text)
                if match and match.group(0) in unbalanced:
                    highlight = page.add_highlight_annot(fitz.Rect(w[:4]))
                    highlight.set_colors(stroke=(1, 1, 0))
                    highlight.update()

    output_path = unique_file_name(
        str(output_folder / f"{pdf_path.stem}_highlighted.pdf")
    )
    doc.save(output_path)
    doc.close()
    print(f"Processed and saved highlighted PDF: {Path(output_path).name}")


# ---------------- Main Reconciler Function ----------------
def reconciller(config_data=None):
    """
    Main reconciliation function that processes PDFs from Downloads folder
    and saves highlighted results to Desktop.

    Args:
        config_data: Optional dict with configuration (input_folder, output_folder)
    """
    # Setup folders
    input_folder = Path.home() / "Downloads"
    output_folder = Path.home() / "Desktop"

    if config_data:
        input_folder = Path(config_data.get("input_folder", input_folder))
        output_folder = Path(config_data.get("output_folder", output_folder))

    output_folder.mkdir(parents=True, exist_ok=True)

    print("Searching for matching PDFs...")
    matching_pairs = find_matching_pdfs_by_policy(input_folder)

    if not matching_pairs:
        print("No matching PDFs found. Exiting.")
        return

    print(f"Found {len(matching_pairs)} matching pair(s).\n")

    all_pdfs = {pdf for pair in matching_pairs for pdf in pair}
    aggregates = aggregate_pdfs(all_pdfs)

    # Update balanced status for all matched pairs
    for pdf1, pdf2 in matching_pairs:
        update_balanced_status(aggregates[pdf1.name], aggregates[pdf2.name])

    for pdf_path in all_pdfs:
        highlight_unbalanced_policies(
            pdf_path, aggregates[pdf_path.name], output_folder
        )

    # Print summary
    # print("\n" + "=" * 60)
    # print("RECONCILIATION SUMMARY")
    # print("=" * 60)
    # for pdf_name, df in aggregates.items():
    # print(f"\nAggregated results for {pdf_name}:")
    # print(df.to_string(index=False))


if __name__ == "__main__":
    reconciller()

import fitz  # PyMuPDF
import pandas as pd
import re
from pathlib import Path
from itertools import combinations


# ---------------- Setup ----------------
POLICY_PATTERN = re.compile(r"^[A-Z0-9]{6,}$")
INPUT_FOLDER = Path("../input/")
OUTPUT_FOLDER = Path("../output/")
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


# ---------------- PDF Column Finders ----------------
def find_policy_column(df, min_unique=2):
    best_col, max_unique = None, 0
    for col in df.columns:
        matches = (
            df[col]
            .astype(str)
            .str.strip()
            .apply(lambda x: bool(POLICY_PATTERN.fullmatch(x)))
        )
        unique_matches = df[col][matches].nunique()
        if unique_matches >= min_unique and unique_matches > max_unique:
            best_col, max_unique = col, unique_matches
    return df.columns.get_loc(best_col) if best_col else None


def find_premium_column(df):
    candidates = {}
    for col in df.columns:
        try:
            col_clean = (
                df[col].astype(str).str.strip().str.replace(r"[%\s$,]", "", regex=True)
            )
            col_numeric = pd.to_numeric(col_clean, errors="coerce")
            col_numeric = col_numeric.where(col_clean.str.endswith(".00"))
            count = col_numeric.notna().sum()
            if count > 0:
                candidates[col] = {"count": count, "max_value": col_numeric.max()}
        except Exception:
            continue
    if not candidates:
        return None
    return df.columns.get_loc(
        max(candidates.items(), key=lambda x: x[1]["max_value"])[0]
    )


# ---------------- PDF Data Extraction ----------------
def extract_policies_and_premiums(pdf_path):
    all_data = []
    doc = fitz.open(pdf_path)
    for page_num, page in enumerate(doc):
        for table in page.find_tables(strategy="text") or []:
            df = table.to_pandas()
            if df.empty:
                continue
            policy_idx, premium_idx = find_policy_column(df), find_premium_column(df)
            if policy_idx is None or premium_idx is None:
                continue

            df = df.rename(
                columns={
                    df.columns[policy_idx]: "policy_number",
                    df.columns[premium_idx]: "premium",
                }
            )[["policy_number", "premium"]].dropna()
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
    return pd.DataFrame(all_data)


# ---------------- Matching PDFs ----------------
def find_matching_pdfs_by_policy(folder):
    pdf_files = list(folder.glob("*.pdf"))
    if not pdf_files:
        return []

    pdf_policies = {
        pdf: set(extract_policies_and_premiums(pdf)["policy_number"].unique())
        for pdf in pdf_files
    }
    return [
        (p1, p2)
        for p1, p2 in combinations(pdf_files, 2)
        if pdf_policies[p1] & pdf_policies[p2]
    ]


# ---------------- Aggregation & Balancing ----------------
def aggregate_pdfs(pdf_paths):
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
    map1, map2 = (
        df1.set_index("policy_number")["premium"],
        df2.set_index("policy_number")["premium"],
    )
    df1["balanced"] |= df1["policy_number"].map(map2) == df1["premium"]
    df2["balanced"] |= df2["policy_number"].map(map1) == df2["premium"]


# ---------------- PDF Highlighting ----------------
def highlight_unbalanced_policies(pdf_path, df):
    # Only consider unbalanced policies with premium > 0
    unbalanced = set(df[(~df["balanced"]) & (df["premium"] > 0)]["policy_number"])
    if not unbalanced:
        return

    doc = fitz.open(pdf_path)
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)

        # Draw table bounding boxes
        for table in page.find_tables(strategy="text"):
            rect = fitz.Rect(*table.bbox)
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=(1, 0, 0))
            annot.set_border(width=1)
            annot.update()

        # Highlight unbalanced policies
        for w in page.get_text("words"):
            if w[4].strip() in unbalanced:
                highlight = page.add_highlight_annot(fitz.Rect(w[:4]))
                highlight.set_colors(stroke=(1, 1, 0))
                highlight.update()

    output_path = OUTPUT_FOLDER / f"{pdf_path.stem}_highlighted.pdf"
    doc.save(output_path)
    doc.close()
    print(f"Processed and saved highlighted PDF: {output_path}")


# ---------------- Main ----------------
matching_pairs = find_matching_pdfs_by_policy(INPUT_FOLDER)
if not matching_pairs:
    print("No matching PDFs found. Exiting.")
    exit()

all_pdfs = {pdf for pair in matching_pairs for pdf in pair}
aggregates = aggregate_pdfs(all_pdfs)

# Update balanced status for all matched pairs
for pdf1, pdf2 in matching_pairs:
    update_balanced_status(aggregates[pdf1.name], aggregates[pdf2.name])

# Highlight unbalanced policies
for pdf_path in all_pdfs:
    highlight_unbalanced_policies(pdf_path, aggregates[pdf_path.name])

# Print summary
for pdf_name, df in aggregates.items():
    print(f"\nAggregated results for {pdf_name}:")
    print(df)
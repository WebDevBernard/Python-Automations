import fitz
import re
from pathlib import Path

CWD = Path.cwd()
INPUT = CWD / "input"
OUTPUT = CWD / "output"
CONFIG = CWD / "config.txt"

MONEY_RE = re.compile(r"^\$[\d,\.]+$|^[\d,\.]+%$|^\$?[\d,\.]+%?$")


def load_words(config_path: Path) -> list[str]:
    text = config_path.read_text(encoding="utf-8")
    return [w.strip() for w in text.split(",") if w.strip()]


def redact_pdf(src: Path, dst: Path, words: list[str]):
    doc = fitz.open(src)
    redacted = []
    found = False

    for page in doc:
        if found:
            break
        for phrase in words:
            if found:
                break

            instances = page.search_for(phrase, quads=False)
            if not instances:
                continue

            # Verify exact match — extract the text at that rect and compare
            for match_rect in instances:
                extracted = page.get_textbox(match_rect).strip()
                if extracted.lower() != phrase.lower():
                    continue  # substring match, not exact — skip

                right_region = fitz.Rect(
                    match_rect.x1,
                    match_rect.y0 - 2,
                    match_rect.x1 + 200,
                    match_rect.y1 + 2,
                )
                nearby_words = page.get_text("words", clip=right_region)

                if nearby_words:
                    w = nearby_words[0]
                    next_text = w[4].strip()

                    if not MONEY_RE.match(next_text):
                        print(
                            f"    SKIPPED p{page.number + 1}: '{phrase}' — '{next_text}' is not $ or %"
                        )
                        found = True
                        break

                    page.add_redact_annot(match_rect, fill=(1, 1, 1))
                    page.add_redact_annot(
                        fitz.Rect(w[0], w[1], w[2], w[3]), fill=(1, 1, 1)
                    )
                    redacted.append(
                        f"    p{page.number + 1}: '{phrase}' + '{next_text}'"
                    )
                else:
                    print(
                        f"    SKIPPED p{page.number + 1}: '{phrase}' — no word beside it"
                    )

                found = True
                break

        page.apply_redactions()

    doc.save(dst)
    doc.close()
    return redacted


def main():
    OUTPUT.mkdir(exist_ok=True)

    if not CONFIG.exists():
        print(f"config.txt not found at {CONFIG}")
        return

    words = load_words(CONFIG)
    if not words:
        print("No words found in config.txt")
        return

    print(f"Loaded {len(words)} word(s): {words}\n")

    pdfs = sorted(INPUT.rglob("*.pdf"))
    if not pdfs:
        print(f"No PDFs found in {INPUT}")
        return

    print(f"Found {len(pdfs)} PDF(s)\n")

    ok = 0
    for src in pdfs:
        relative = src.relative_to(INPUT)
        dst = OUTPUT / relative
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            hits = redact_pdf(src, dst, words)
            if hits:
                print(f"  ✓  {src.name}")
                for h in hits:
                    print(h)
            else:
                print(f"  -  {src.name}  (no matches)")
            ok += 1
        except Exception as exc:
            print(f"  ✗  {src.name}  — {exc}")

    print(f"\n  Processed: {ok}")
    print(f"  Saved to: {OUTPUT}")


if __name__ == "__main__":
    main()

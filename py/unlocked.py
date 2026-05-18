import pikepdf
from pathlib import Path

DOWNLOADS = Path.home() / "Downloads"
OUTPUT_DIR = DOWNLOADS / "unlocked"


def strip_pdf(src: Path, dst: Path):
    with pikepdf.open(src) as pdf:

        # Remove top-level restrictions
        for key in ["/Perms", "/NeedsRendering", "/OpenAction", "/AA", "/Names"]:
            if key in pdf.Root:
                del pdf.Root[key]

        if "/AcroForm" in pdf.Root:
            acroform = pdf.Root["/AcroForm"]

            if "/SigFlags" in acroform:
                del acroform["/SigFlags"]

            # Remove XFA — this is the only thing blocking save
            if "/XFA" in acroform:
                del acroform["/XFA"]

            # Remove sig fields
            if "/Fields" in acroform:
                clean = []
                for ref in acroform["/Fields"]:
                    try:
                        field = pdf.get_object(ref.objgen)
                        if field.get("/FT") != "/Sig":
                            clean.append(ref)
                    except:
                        clean.append(ref)
                acroform["/Fields"] = pikepdf.Array(clean)

        # Scan every object — find date widget fields (had <dateTimeEdit/> in XFA)
        # and give them AcroForm date format+keystroke actions so they still
        # validate MM/DD/YYYY input
        for obj in pdf.objects:
            try:
                if not isinstance(obj, pikepdf.Dictionary):
                    continue
                if str(obj.get("/Subtype", "")) != "/Widget":
                    continue
                if str(obj.get("/FT", "")) != "/Tx":
                    continue

                tm = str(obj.get("/TM", ""))
                if "Date" not in tm:
                    continue

                # Replace XFA date widget with AcroForm date format actions
                obj["/AA"] = pikepdf.Dictionary(
                    F=pikepdf.Dictionary(
                        Type=pikepdf.Name("/Action"),
                        S=pikepdf.Name("/JavaScript"),
                        JS=pikepdf.String('AFDate_FormatEx("mm/dd/yyyy");'),
                    ),
                    K=pikepdf.Dictionary(
                        Type=pikepdf.Name("/Action"),
                        S=pikepdf.Name("/JavaScript"),
                        JS=pikepdf.String('AFDate_KeystrokeEx("mm/dd/yyyy");'),
                    ),
                )
                # Mark as date field
                obj["/V"] = pikepdf.String("")

            except:
                pass

        pdf.save(dst)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"Scanning: {DOWNLOADS}")
    print(f"Output:   {OUTPUT_DIR}\n")

    pdfs = sorted(DOWNLOADS.rglob("*.pdf"))
    pdfs = [p for p in pdfs if OUTPUT_DIR not in p.parents]

    if not pdfs:
        print("No PDFs found.")
        return

    print(f"Found {len(pdfs)} PDF(s):\n")

    ok = fail = 0
    for src in pdfs:
        dst = OUTPUT_DIR / src.relative_to(DOWNLOADS)
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            strip_pdf(src, dst)
            print(f"  ✓  {src.name}")
            ok += 1
        except Exception as exc:
            print(f"  ✗  {src.name}  — {exc}")
            fail += 1

    print(f"\n  Unlocked: {ok}   Skipped: {fail}")
    print(f"  Saved to: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()

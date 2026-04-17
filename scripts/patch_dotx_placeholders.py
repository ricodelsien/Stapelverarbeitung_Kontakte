#!/usr/bin/env python3
"""Fügt docxtemplater-Platzhalter in eine Word-Vorlage (.dotx) ein — ein Tag pro <w:t>-Lauf."""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path


def patch_document_xml(xml: str) -> str:
    # 1) Vier Zeilen Briefkopf → ein Absatz mit {Briefkopf_Block} (linebreaks in der App)
    block = (
        r'<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
        r'<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t>{Briefkopf_Block}</w:t></w:r></w:p>'
    )
    old_addr = (
        r'<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
        r'<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t>Behörde, Organisation usw\.</w:t></w:r></w:p>'
        r'<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
        r'<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t>Vorname Name</w:t></w:r></w:p>'
        r'<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
        r'<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t>Straße Hausnummer</w:t></w:r></w:p>'
        r'<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
        r'<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t>PLZ ORT</w:t></w:r></w:p>'
    )
    xml2, n = re.subn(old_addr, block, xml, count=1)
    if n != 1:
        raise SystemExit(
            "Adressblock nicht gefunden (Vorlage abweichend?). Erwartet vier feste Zeilen „Behörde…“ bis „PLZ ORT“."
        )

    # 2) Anrede: „Sehr geehrte“ → vollständige {Briefanrede}
    xml2, n = re.subn(r"<w:t>Sehr geehrte</w:t>", r"<w:t>{Briefanrede}</w:t>", xml2, count=1)
    if n != 1:
        raise SystemExit("Text „Sehr geehrte“ nicht gefunden.")

    # 3) Rechtsbündiges Datum: Feld / Lesezeichen entfernen, Platzhalter aus der App
    date_para = (
        r'<w:p w:rsidR="003C17ED" w:rsidRPr="009D52CE" w:rsidRDefault="00DE5D7D" w:rsidP="00B53C3A">'
        r'<w:pPr><w:jc w:val="right"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r[^>]*><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t xml:space="preserve">Berlin, </w:t></w:r>.*?</w:p>'
    )
    date_repl = (
        r'<w:p w:rsidR="003C17ED" w:rsidRPr="009D52CE" w:rsidRDefault="00DE5D7D" w:rsidP="00B53C3A">'
        r'<w:pPr><w:jc w:val="right"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r'<w:t xml:space="preserve">Berlin, </w:t></w:r>'
        r'<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
        r"<w:t>{Briefdatum}</w:t></w:r></w:p>"
    )
    xml3, n = re.subn(date_para, date_repl, xml2, count=1, flags=re.DOTALL)
    if n != 1:
        raise SystemExit("Datumszeile „Berlin, …“ nicht gefunden.")

    return xml3


def patch_dotx(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".patching")
    shutil.copy2(src, tmp)
    try:
        with zipfile.ZipFile(tmp, "r") as zin:
            data = zin.read("word/document.xml").decode("utf-8")
            patched = patch_document_xml(data)
            buf = patched.encode("utf-8")
            # Zip neu schreiben (gleiche Reihenfolge der übrigen Dateien)
            with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, buf)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        tmp.unlink()
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise


def main() -> None:
    base = Path("/Users/nicosiedler/Desktop/2026")
    candidates = sorted(base.glob("Einladung_*Vorlage.dotx"))
    if not candidates:
        print("Keine Einladung_*Vorlage.dotx in", base, file=sys.stderr)
        sys.exit(1)
    src = candidates[0]
    bak = src.with_suffix(src.suffix + ".vor_platz_backup")
    out = src.with_name(src.stem + "_mit_Platzhaltern.dotx")
    shutil.copy2(src, bak)
    patch_dotx(src, out)
    print("OK:", out)
    print("Backup:", bak)


if __name__ == "__main__":
    main()

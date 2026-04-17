#!/usr/bin/env python3
"""Ersetzt {Briefkopf_Block} durch docxtemplater-Absatzschleife {#Briefkopf_zeilen}{.}{/Briefkopf_zeilen}."""
from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

OLD_PARA = (
    '<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
    '<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
    '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
    "<w:t>{Briefkopf_Block}</w:t></w:r></w:p>"
)

RUN_OPEN = (
    '<w:p w:rsidR="00B53C3A" w:rsidRPr="00B53C3A" w:rsidRDefault="0090173E" w:rsidP="00B53C3A">'
    '<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:pPr>'
    '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr>'
    "<w:t>"
)
RUN_CLOSE = "</w:t></w:r></w:p>"


def patch_document_xml(xml: str) -> str:
    if OLD_PARA not in xml:
        raise SystemExit("Erwarteter Absatz mit {Briefkopf_Block} nicht gefunden (Vorlage abweichend?).")
    loop = (
        RUN_OPEN
        + "{#Briefkopf_zeilen}"
        + RUN_CLOSE
        + RUN_OPEN
        + "{.}"
        + RUN_CLOSE
        + RUN_OPEN
        + "{/Briefkopf_zeilen}"
        + RUN_CLOSE
    )
    return xml.replace(OLD_PARA, loop, 1)


def patch_dotx(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".patching")
    shutil.copy2(src, tmp)
    try:
        with zipfile.ZipFile(tmp, "r") as zin:
            data = zin.read("word/document.xml").decode("utf-8")
            patched = patch_document_xml(data)
            buf = patched.encode("utf-8")
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
    files = [
        base / "Einladung_AVöD_Auftaktveranstaltung_20260708_Vorlage.dotx",
        base / "Einladung_AVöD_Auftaktveranstaltung_20260708_Vorlage_mit_Platzhaltern.dotx",
    ]
    for src in files:
        if not src.exists():
            print("Überspringe (fehlt):", src, file=sys.stderr)
            continue
        bak = src.with_suffix(src.suffix + ".vor_briefkopf_loop_backup")
        shutil.copy2(src, bak)
        patch_dotx(src, src)
        print("OK:", src.name)
        print("     Backup:", bak.name)


if __name__ == "__main__":
    main()

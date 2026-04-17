#!/usr/bin/env python3
"""
Verdichtet die Fensteradresse in der Serienbrief-Vorlage: kleinere Schrift (10 pt statt 12 pt),
weniger Abstand zwischen den Adresszeilen — optisch näher an klassischen DIN-Briefen.

Erwartet die docxtemplater-Schleife {#Briefkopf_zeilen} … {.} … {/Briefkopf_zeilen} in word/document.xml.
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

# Drei aufeinanderfolgende Absätze der Briefkopf-Schleife (non-greedy pro Absatz)
BRIEFKOPF_LOOP_RE = re.compile(
    r"<w:p\b[^>]*>.*?\{#Briefkopf_zeilen\}.*?</w:p>\s*"
    r"<w:p\b[^>]*>.*?\{\.\}.*?</w:p>\s*"
    r"<w:p\b[^>]*>.*?\{/Briefkopf_zeilen\}.*?</w:p>",
    re.DOTALL,
)


def patch_briefkopf_block_compact(xml: str) -> tuple[str, bool]:
    m = BRIEFKOPF_LOOP_RE.search(xml)
    if not m:
        return xml, False
    block = m.group(0)
    # 12 pt → 10 pt (Word: halbe Punkte)
    patched = block.replace('w:sz w:val="24"', 'w:sz w:val="20"')
    # Nach Absatz 0 twips — enger als Word-Standard
    patched = re.sub(
        r"(<w:pPr>)(?!\s*<w:spacing\b)",
        r'\1<w:spacing w:before="0" w:after="20"/>',
        patched,
    )
    return xml[: m.start()] + patched + xml[m.end() :], True


def patch_dotx(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".patching")
    shutil.copy2(src, tmp)
    try:
        with zipfile.ZipFile(tmp, "r") as zin:
            data = zin.read("word/document.xml").decode("utf-8")
            patched, ok = patch_briefkopf_block_compact(data)
            if not ok:
                raise SystemExit(
                    "Briefkopf-Schleife {#Briefkopf_zeilen} nicht gefunden — Vorlage abweichend?"
                )
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
        bak = src.with_suffix(src.suffix + ".vor_briefkopf_compact_backup")
        shutil.copy2(src, bak)
        patch_dotx(src, src)
        print("OK:", src.name)
        print("     Backup:", bak.name)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Macht die E-Mail-Word-Vorlage „E-Mail Vorlage Einladung zur Abschluss“ docxtemplater-kompatibel:

1) Ersetzt die kaputte Anrede „Sehr geehrte “ + leerem Absatz durch {Briefanrede},
2) Fügt einen Absatz mit {Veranstaltung_Terminzeile} ein (Termin aus der App).

Wendet auf .dotx und .docx im gleichen Ordner an (falls vorhanden).
"""
from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

# Aus aktueller Vorlage (Desktop/2026)
OLD_SALUTATION = (
    '<w:t xml:space="preserve">Sehr geehrte </w:t></w:r></w:p><w:p w14:paraId="203BCD9A" w14:textId="77777777" '
    'w:rsidR="008673B8" w:rsidRPr="008673B8" w:rsidRDefault="008673B8" w:rsidP="008673B8"/><w:p w14:paraId="314B5F84" '
    'w14:textId="77777777" w:rsidR="008673B8" w:rsidRDefault="008673B8" w:rsidP="008673B8">'
    '<w:r w:rsidRPr="008673B8"><w:t xml:space="preserve">im Namen des </w:t></w:r>'
)

NEW_SALUTATION = (
    '<w:t>{Briefanrede},</w:t></w:r></w:p><w:p w14:paraId="314B5F84" w14:textId="77777777" '
    'w:rsidR="008673B8" w:rsidRDefault="008673B8" w:rsidP="008673B8">'
    '<w:r w:rsidRPr="008673B8"><w:t xml:space="preserve">im Namen des </w:t></w:r>'
)

# Leerer Absatz nach dem Absatz „Teilnehmerinnen und Teilnehmer.“ → Terminzeile
OLD_EMPTY_AFTER_EVENT = (
    '<w:p w14:paraId="53FAD3DA" w14:textId="77777777" w:rsidR="008673B8" w:rsidRPr="008673B8" '
    'w:rsidRDefault="008673B8" w:rsidP="008673B8"/>'
)

NEW_TERMIN_PARA = (
    '<w:p w14:paraId="53FAD3DA" w14:textId="77777777" w:rsidR="008673B8" w:rsidRPr="008673B8" '
    'w:rsidRDefault="008673B8" w:rsidP="008673B8"><w:r w:rsidRPr="008673B8">'
    '<w:t>{Veranstaltung_Terminzeile}</w:t></w:r></w:p>'
)


def patch_document_xml(xml: str) -> tuple[str, list[str]]:
    msgs: list[str] = []
    if OLD_SALUTATION not in xml:
        msgs.append("Hinweis: Anrede-Block nicht exakt wie erwartet — evtl. Vorlage schon geändert.")
    else:
        xml = xml.replace(OLD_SALUTATION, NEW_SALUTATION, 1)
        msgs.append("OK: Anrede → {Briefanrede},")
    if OLD_EMPTY_AFTER_EVENT not in xml:
        msgs.append("Hinweis: leerer Absatz (53FAD3DA) nicht gefunden — Terminzeile nicht eingefügt.")
    else:
        xml = xml.replace(OLD_EMPTY_AFTER_EVENT, NEW_TERMIN_PARA, 1)
        msgs.append("OK: Terminzeile → {Veranstaltung_Terminzeile}")
    return xml, msgs


def patch_ooxml(path: Path) -> None:
    if not path.exists():
        print("Überspringe (fehlt):", path, file=sys.stderr)
        return
    bak = path.with_suffix(path.suffix + ".vor_email_platzhalter_backup")
    shutil.copy2(path, bak)
    tmp = path.with_suffix(path.suffix + ".patching")
    shutil.copy2(path, tmp)
    try:
        with zipfile.ZipFile(tmp, "r") as zin:
            data = zin.read("word/document.xml").decode("utf-8")
            patched, msgs = patch_document_xml(data)
            buf = patched.encode("utf-8")
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, buf)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        tmp.unlink()
        print(path.name)
        for m in msgs:
            print(" ", m)
        print(" Backup:", bak.name)
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise


def main() -> None:
    base = Path("/Users/nicosiedler/Desktop/2026")
    for name in (
        "E-Mail Vorlage Einladung zur Abschluss.dotx",
        "E-Mail Vorlage Einladung zur Abschluss.docx",
    ):
        patch_ooxml(base / name)


if __name__ == "__main__":
    main()

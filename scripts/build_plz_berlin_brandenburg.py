#!/usr/bin/env python3
"""Lädt PLZ→Ort für Berlin (11) und Brandenburg (12) über OpenPLZ API (curl), schreibt plz-berlin-brandenburg.json."""

from __future__ import annotations

import json
import subprocess
import time
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "plz-berlin-brandenburg.json"


def fetch_page(regex: str, page: int) -> list | None:
    from urllib.parse import urlencode

    url = "https://openplzapi.org/de/Localities?" + urlencode(
        {"postalCode": regex, "page": page, "pageSize": 50}
    )
    p = subprocess.run(
        ["curl", "-sL", "--max-time", "120", url],
        capture_output=True,
        text=True,
    )
    if p.returncode != 0:
        print("curl error", p.stderr[:300])
        return None
    try:
        data = json.loads(p.stdout)
    except json.JSONDecodeError:
        print("bad json", p.stdout[:300])
        return None
    if isinstance(data, dict) and data.get("status") in (400, "400"):
        return None
    if not isinstance(data, list):
        return None
    return data


def main() -> None:
    out: dict[str, str] = {}
    for prefix in range(1, 20):
        regex = f"^{prefix:02d}"
        page = 1
        while True:
            data = fetch_page(regex, page)
            if not data:
                break
            for loc in data:
                if not isinstance(loc, dict):
                    continue
                if loc.get("federalState", {}).get("key") not in ("11", "12"):
                    continue
                pc = loc.get("postalCode")
                name = loc.get("name")
                if pc and name and pc not in out:
                    out[pc] = name
            if len(data) < 50:
                break
            page += 1
            time.sleep(0.06)
        print(regex, "→", len(out), "PLZ")
    OUT.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print("Geschrieben:", OUT, "Anzahl:", len(out))


if __name__ == "__main__":
    main()

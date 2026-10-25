"""Document numbers: IN[YYMMDD-NN] and DO[YYMMDD-NN].

A small JSON register remembers the last sequence used for each date, so a
second PO processed on the same date gets -02, the next -03, and so on.
"""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path


def invoice_no(d: date, seq: int) -> str:
    return f"IN{d:%y%m%d}-{seq:02d}"


def do_no(d: date, seq: int) -> str:
    return f"DO{d:%y%m%d}-{seq:02d}"


class Register:
    def __init__(self, path: Path):
        self.path = Path(path)

    def _load(self) -> dict:
        if not self.path.exists():
            return {}
        try:
            return json.loads(self.path.read_text())
        except (OSError, json.JSONDecodeError):
            return {}

    def last_seq(self, d: date) -> int:
        return int(self._load().get(d.isoformat(), {}).get("last_seq", 0))

    def next_seq(self, d: date) -> int:
        return self.last_seq(d) + 1

    def record(self, d: date, seq: int, po_number: str) -> None:
        data = self._load()
        entry = data.setdefault(d.isoformat(), {"last_seq": 0, "issued": []})
        entry["last_seq"] = max(int(entry.get("last_seq", 0)), seq)
        entry.setdefault("issued", []).append({"seq": seq, "po_number": po_number})
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(data, indent=2))

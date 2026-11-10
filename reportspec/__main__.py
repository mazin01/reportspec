"""reportspec SPEC.json --data NAME=FILE [--data ...] [--pdf out.pdf] [--xlsx out.xlsx] [--html out.html]

FILE is .csv (any encoding nuqta can sniff) or .json (a list of objects)."""
import argparse
import csv
import io
import json
import sys
from pathlib import Path

import nuqta as nq

from . import SpecError, render_html, render_pdf, render_xlsx


def _load(path: Path) -> list[dict]:
    raw = path.read_bytes()
    if path.suffix.lower() == ".json":
        return json.loads(raw.decode("utf-8-sig"))
    text = raw.decode(nq.sniff_encoding(raw[:1_000_000]), errors="replace")
    return list(csv.DictReader(io.StringIO(text)))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="reportspec", description="Render a JSON report spec to PDF / Excel / HTML")
    ap.add_argument("spec")
    ap.add_argument("--data", action="append", default=[], metavar="NAME=FILE")
    ap.add_argument("--pdf")
    ap.add_argument("--xlsx")
    ap.add_argument("--html")
    a = ap.parse_args(argv)
    if not (a.pdf or a.xlsx or a.html):
        ap.error("ask for at least one of --pdf, --xlsx, --html")
    try:
        spec = json.loads(Path(a.spec).read_text(encoding="utf-8-sig"))
        data = {}
        for item in a.data:
            name, _, file = item.partition("=")
            if not file:
                ap.error(f"--data expects NAME=FILE, got {item!r}")
            data[name] = _load(Path(file))
        if a.html:
            Path(a.html).write_text(render_html(spec, data), encoding="utf-8")
        if a.xlsx:
            render_xlsx(spec, data, a.xlsx)
        if a.pdf:
            render_pdf(spec, data, a.pdf)
    except SpecError as e:
        print(json.dumps({"error": e.code, "message": str(e), **e.details}, ensure_ascii=False), file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())

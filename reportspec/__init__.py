"""reportspec: one JSON spec → a PDF and an Excel report. Arabic/RTL-ready.

    import reportspec as rs

    spec = {
        "title": "Monthly sales",
        "sections": [
            {"type": "metrics", "items": [{"label": "Revenue", "data": "orders", "column": "amount", "fn": "sum"}]},
            {"type": "table", "data": "orders", "number_columns": ["amount"], "totals": True},
        ],
    }
    rs.render_pdf(spec, {"orders": rows}, "report.pdf")     # WeasyPrint
    rs.render_xlsx(spec, {"orders": rows}, "report.xlsx")   # xlsxwriter
    html = rs.render_html(spec, {"orders": rows})            # no native libraries needed
"""
from __future__ import annotations

import html
import math
import re
from datetime import date, datetime
from pathlib import Path
from string import Template

import nuqta as nq

__all__ = ["render_html", "render_pdf", "render_xlsx", "validate", "SpecError"]
__version__ = "0.1.0"

FONTS = Path(__file__).parent / "fonts"
SECTION_TYPES = ("metrics", "table", "text")
FUNCTIONS = ("sum", "avg", "min", "max", "count", "rows")
HIGHLIGHTS = {  # name: (PDF background, Excel fill)
    "green": ("#dcfce7", "#C6EFCE"), "yellow": ("#fef9c3", "#FFEB9C"), "red": ("#fee2e2", "#FFC7CE"),
    "orange": ("#ffedd5", "#FFCC99"), "blue": ("#dbeafe", "#9DC3E6"),
}
LABELS = {
    "ltr": {"page": "Page", "of": "of", "total": "Total"},
    "rtl": {"page": "صفحة", "of": "من", "total": "الإجمالي"},
}
WIDE_TABLE = 7  # a table with this many columns switches the PDF to landscape


class SpecError(ValueError):
    """Invalid spec or data, with a stable `code` (e.g. COLUMN_NOT_FOUND)."""

    def __init__(self, code: str, message: str, **details):
        super().__init__(f"{code}: {message}")
        self.code, self.details = code, details


# ── data & validation ──────────────────────────────────────────────────────────

def _records(data: dict, name, where: str) -> list[dict]:
    if name not in data:
        raise SpecError("MISSING_DATA", f"{where}: no dataset named {name!r}", available=sorted(data))
    rows = data[name]
    if hasattr(rows, "to_dict"):  # pandas DataFrame
        rows = rows.to_dict("records")
    rows = list(rows)
    if not all(isinstance(r, dict) for r in rows):
        raise SpecError("BAD_DATA", f"{where}: dataset {name!r} must be a list of dicts or a DataFrame")
    return rows


def _columns(rows: list[dict], wanted, where: str) -> list[str]:
    present = list(dict.fromkeys(k for r in rows for k in r))
    if wanted is None:
        return present
    missing = [c for c in wanted if c not in present]
    if missing and rows:
        raise SpecError("COLUMN_NOT_FOUND", f"{where}: columns {missing} not in the data", available=present)
    return list(wanted)


def validate(spec: dict, data: dict | None = None) -> None:
    """Raise SpecError if the spec (and, when given, the data it names) is unusable."""
    if not isinstance(spec, dict) or not isinstance(spec.get("sections", []), list):
        raise SpecError("BAD_SPEC", "spec must be a dict with a 'sections' list")
    if spec.get("direction", "auto") not in ("auto", "rtl", "ltr"):
        raise SpecError("BAD_SPEC", "direction must be 'auto', 'rtl' or 'ltr'")
    for i, s in enumerate(spec.get("sections", [])):
        where = f"sections[{i}]"
        if not isinstance(s, dict) or s.get("type") not in SECTION_TYPES:
            raise SpecError("UNKNOWN_SECTION_TYPE", f"{where}: type must be one of {SECTION_TYPES}",
                            got=s.get("type") if isinstance(s, dict) else repr(s))
        if s["type"] == "metrics":
            for j, m in enumerate(s.get("items", [])):
                fn = m.get("fn", "sum")
                if fn not in FUNCTIONS:
                    raise SpecError("BAD_SPEC", f"{where}.items[{j}]: fn must be one of {FUNCTIONS}")
                if data is not None and "value" not in m:
                    rows = _records(data, m.get("data"), f"{where}.items[{j}]")
                    if fn != "rows":
                        _columns(rows, [m.get("column")], f"{where}.items[{j}]")
        elif s["type"] == "table":
            hl = s.get("highlight") or {}
            for color in hl.get("map", {}).values():
                if color not in HIGHLIGHTS:
                    raise SpecError("BAD_SPEC", f"{where}: highlight colours are {sorted(HIGHLIGHTS)}", got=color)
            if data is not None:
                rows = _records(data, s.get("data"), where)
                _columns(rows, s.get("columns"), where)
                named = s.get("number_columns", []) + s.get("percent_columns", []) + \
                    ([hl["column"]] if hl.get("column") else [])
                _columns(rows, named, where)


# ── values ─────────────────────────────────────────────────────────────────────

def _number(v) -> float | None:
    if isinstance(v, bool):
        return None
    n = float(v) if isinstance(v, (int, float)) else nq.to_number(v)
    return n if n is not None and math.isfinite(n) else None


def _fmt(n: float, decimals: int) -> str:
    return f"{n:,.{decimals}f}"


def _text(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, datetime):
        return v.isoformat(sep=" ", timespec="minutes")
    if isinstance(v, date):
        return v.isoformat()
    return str(v)


def _metric(m: dict, data: dict) -> tuple[str, float | None, int]:
    """→ (display text, number or None, decimals)."""
    decimals = m.get("decimals", 2)
    if "value" in m:
        n = _number(m["value"])
        return (_fmt(n, decimals) if n is not None else _text(m["value"])), n, decimals
    rows = _records(data, m.get("data"), "metric")
    fn = m.get("fn", "sum")
    if fn == "rows":
        return str(len(rows)), float(len(rows)), 0
    nums = [n for n in (_number(r.get(m["column"])) for r in rows) if n is not None]
    if fn == "count":
        return str(len(nums)), float(len(nums)), 0
    if not nums:
        return "—", None, decimals
    n = {"sum": math.fsum(nums), "avg": math.fsum(nums) / len(nums), "min": min(nums), "max": max(nums)}[fn]
    return _fmt(n, decimals), n, decimals


def _direction(spec: dict, data: dict) -> str:
    """Explicit, or RTL when the title, a section title or a column name reads right-to-left."""
    if spec.get("direction", "auto") != "auto":
        return spec["direction"]
    sample = [spec.get("title", "")] + [s.get("title", "") for s in spec.get("sections", [])]
    for s in spec.get("sections", []):
        if s["type"] == "table":
            sample += _columns(_records(data, s.get("data"), "table"), s.get("columns"), "table")
    return "rtl" if any(nq.text_direction(str(t)) == "rtl" for t in sample if t) else "ltr"


def _table(s: dict, data: dict):
    """→ (columns, rows, totals, decimals per column).
    A row is (cells, highlight colour); a cell is (text, number or None)."""
    rows = _records(data, s.get("data"), "table")
    cols = _columns(rows, s.get("columns"), "table")
    nums, pcts = set(s.get("number_columns", [])), set(s.get("percent_columns", []))
    parsed = {c: [_number(r.get(c)) if c in nums or c in pcts else None for r in rows] for c in cols}
    # Explicit "decimals" wins; otherwise whole-number columns (quantities, counts) get none.
    decs = [s.get("decimals", 0 if all(n is None or n.is_integer() for n in parsed[c]) else 2) for c in cols]
    hl = s.get("highlight") or {}
    hl_map = {k.lower(): v for k, v in hl.get("map", {}).items()}
    out = []
    for y, r in enumerate(rows):
        cells = []
        for i, c in enumerate(cols):
            n = parsed[c][y]
            text = _text(r.get(c)) if n is None else _fmt(n, decs[i]) + ("%" if c in pcts else "")
            cells.append((text, n))
        color = hl_map.get(_text(r.get(hl["column"])).lower()) if hl.get("column") else None
        out.append((cells, color))
    totals = None
    if s.get("totals"):
        totals = []
        for i, c in enumerate(cols):
            t = math.fsum(n or 0 for n in parsed[c]) if c in nums else None
            totals.append((_fmt(t, decs[i]), t) if t is not None else ("", None))
    return cols, out, totals, decs


# ── HTML / PDF ─────────────────────────────────────────────────────────────────

CSS = Template("""
@font-face { font-family: "Plex Arabic"; src: url("$regular"); }
@font-face { font-family: "Plex Arabic"; font-weight: bold; src: url("$bold"); }
@page { size: $page; margin: 14mm 12mm 16mm;
  @bottom-center { content: "$page_label " counter(page) " $of_label " counter(pages);
                   font: 8pt "Plex Arabic"; color: #94a3b8; }
  @bottom-right { content: "$footer"; font: 8pt "Plex Arabic"; color: #94a3b8; } }
* { box-sizing: border-box; }
body { font: 9pt/1.45 "Plex Arabic", sans-serif; color: #0f172a; margin: 0; }
header { border-bottom: 1px solid #e2e8f0; padding-bottom: 6pt; margin-bottom: 8pt; }
h1 { font-size: 17pt; margin: 0 0 2pt; } .subtitle { color: #64748b; margin: 0; }
h2 { font-size: 11pt; margin: 12pt 0 5pt; color: #1e293b; }
.metrics { display: table; border-spacing: 6pt 0; margin: 0 -6pt; }
.metric { display: table-cell; border: 0.5pt solid #e2e8f0; border-radius: 4pt; padding: 6pt 10pt; min-width: 110pt; }
.metric .label { color: #64748b; font-size: 8pt; } .metric .value { font-size: 14pt; font-weight: bold; }
table { width: 100%; border-collapse: collapse; } thead { display: table-header-group; }
th { background: #1e293b; color: #fff; padding: 4pt 5pt; font-size: 8.5pt; text-align: start; }
td { padding: 3pt 5pt; border-bottom: 0.5pt solid #e2e8f0; font-size: 8pt; }
tbody tr:nth-child(even) td { background: #f8fafc; }
td.num { text-align: end; white-space: nowrap; direction: ltr; }
.neg { color: #cc0000; } tr.totals td { font-weight: bold; background: #e2e8f0; border-top: 1pt solid #94a3b8; }
tr { break-inside: avoid; } p.text { margin: 4pt 0; white-space: pre-wrap; }
""")


_DATE_RE = re.compile(r"[0-9\u0660-\u0669]{1,4}([-/.])[0-9\u0660-\u0669]{1,2}\1[0-9\u0660-\u0669]{1,4}")


def _isolate(text: str) -> str:
    """Wrap dates in Unicode isolates (LRI…PDI). In RTL text a date after an Arabic
    word becomes an "Arabic number" and "-" stops holding it together, so
    2025-03-31 would display as 31-03-2025."""
    return _DATE_RE.sub(lambda m: "\u2066" + m.group() + "\u2069", text)


def _dir(text: str, doc_dir: str) -> str:
    """A dir attribute only where the text runs against the document direction."""
    d = nq.text_direction(text) if text.strip() else doc_dir
    return "" if d == doc_dir else f' dir="{d}"'


def _td(text: str, n, doc_dir: str, red: bool) -> str:
    if n is None:
        shown = _isolate(text) if doc_dir == "rtl" else text
        return f"<td{_dir(text, doc_dir)}>{html.escape(shown)}</td>"
    return f'<td class="{"num neg" if red and n < 0 else "num"}">{html.escape(text)}</td>'


def render_html(spec: dict, data: dict | None = None) -> str:
    """The report as a standalone HTML document (what the PDF is made from)."""
    data = data or {}
    validate(spec, data)
    d = _direction(spec, data)
    labels = {**LABELS[d], **spec.get("labels", {})}
    e = (lambda t: html.escape(_isolate(t))) if d == "rtl" else html.escape
    wide = any(s["type"] == "table" and len(_table(s, data)[0]) >= WIDE_TABLE for s in spec.get("sections", []))
    css_text = lambda s: e(s).replace('"', '\\"')  # noqa: E731  (inside a CSS string)
    css = CSS.substitute(regular=(FONTS / "IBMPlexSansArabic-Regular.ttf").as_uri(),
                         bold=(FONTS / "IBMPlexSansArabic-Bold.ttf").as_uri(),
                         page=spec.get("page", "A4 landscape" if wide else "A4"),
                         page_label=css_text(labels["page"]), of_label=css_text(labels["of"]),
                         footer=css_text(spec.get("footer", "")))
    title = spec.get("title", "")
    out = [f'<!doctype html><html dir="{d}"><head><meta charset="utf-8"><title>{e(title)}</title>'
           f"<style>{css}</style></head><body><header><h1{_dir(title, d)}>{e(title)}</h1>"]
    if spec.get("subtitle"):
        out.append(f'<p class="subtitle"{_dir(spec["subtitle"], d)}>{e(spec["subtitle"])}</p>')
    out.append("</header>")
    for s in spec.get("sections", []):
        if s.get("title"):
            out.append(f'<h2{_dir(s["title"], d)}>{e(s["title"])}</h2>')
        if s["type"] == "text":
            body = s.get("body", "")
            out.append(f'<p class="text"{_dir(body, d)}>{e(body)}</p>')
        elif s["type"] == "metrics":
            out.append('<div class="metrics">')
            for m in s.get("items", []):
                text, n, _ = _metric(m, data)
                label = m.get("label", "")
                neg = " neg" if n is not None and n < 0 else ""
                out.append(f'<div class="metric"><div class="label"{_dir(label, d)}>{e(label)}</div>'
                           f'<div class="value{neg}" dir="ltr">{e(text)}</div></div>')
            out.append("</div>")
        else:
            cols, rows, totals, _ = _table(s, data)
            red = s.get("negative_red", True)
            out.append("<table><thead><tr>" + "".join(f"<th{_dir(c, d)}>{e(c)}</th>" for c in cols)
                       + "</tr></thead><tbody>")
            for cells, color in rows:
                style = f' style="background:{HIGHLIGHTS[color][0]}"' if color else ""
                out.append(f"<tr{style}>" + "".join(_td(t, n, d, red) for t, n in cells) + "</tr>")
            if totals:
                if totals[0][1] is None:
                    totals[0] = (labels["total"], None)
                out.append('<tr class="totals">' + "".join(_td(t, n, d, red) for t, n in totals) + "</tr>")
            out.append("</tbody></table>")
    out.append("</body></html>")
    return "".join(out)


def render_pdf(spec: dict, data: dict | None, out) -> Path:
    """Write the PDF with WeasyPrint (needs Pango: see README)."""
    from weasyprint import HTML

    out = Path(out)
    HTML(string=render_html(spec, data)).write_pdf(out)
    return out


# ── Excel ──────────────────────────────────────────────────────────────────────

def render_xlsx(spec: dict, data: dict | None, out) -> Path:
    """Write an .xlsx: a 'Summary' sheet (title, metrics, text) and one sheet per table.
    Numbers are written as numbers with formats; RTL documents get right-to-left sheets."""
    import xlsxwriter

    data = data or {}
    validate(spec, data)
    d = _direction(spec, data)
    labels = {**LABELS[d], **spec.get("labels", {})}
    out = Path(out)
    wb = xlsxwriter.Workbook(str(out))
    formats: dict = {}

    def fmt(decimals: int | None = None, pct: bool = False, red: bool = False, **props):
        """Cached formats; decimals=None means text."""
        key = (decimals, pct, red, tuple(sorted(props.items())))
        if key not in formats:
            spec_ = {"valign": "top", **props}
            if decimals is None:
                spec_["text_wrap"] = True
            else:
                body = "#,##0" + ("." + "0" * decimals if decimals else "")
                if pct:
                    body = body.replace("#,##", "") + "%"
                spec_["num_format"] = f"{body};[Red]-{body}" if red else body
            formats[key] = wb.add_format(spec_)
        return formats[key]

    used: set = set()

    def sheet(name: str):
        base = "".join(ch for ch in name if ch not in "[]:*?/\\")[:31] or "Sheet"
        n, i = base, 2
        while n.lower() in used:
            n, i = f"{base[:28]}_{i}", i + 1
        used.add(n.lower())
        ws = wb.add_worksheet(n)
        if d == "rtl":
            ws.right_to_left()
        return ws

    summary = sheet("Summary")
    summary.set_column(0, 0, 36)
    summary.set_column(1, 1, 20)
    summary.write(0, 0, spec.get("title", ""), fmt(bold=True, font_size=14))
    r = 1
    if spec.get("subtitle"):
        summary.write(r, 0, spec["subtitle"], fmt(font_color="#64748B"))
        r += 1
    for i, s in enumerate(spec.get("sections", [])):
        if s["type"] == "table":
            _xlsx_table(sheet(s.get("title") or f"Table {i + 1}"), s, data, labels, fmt)
            continue
        r += 1
        if s.get("title"):
            summary.write(r, 0, s["title"], fmt(bold=True, font_size=11))
            r += 1
        if s["type"] == "text":
            summary.write(r, 0, s.get("body", ""), fmt())
            r += 1
            continue
        for m in s.get("items", []):
            text, n, decimals = _metric(m, data)
            summary.write(r, 0, m.get("label", ""), fmt(bold=True))
            if n is None:
                summary.write_string(r, 1, text)
            else:
                summary.write_number(r, 1, n, fmt(decimals, red=True))
            r += 1
    wb.close()
    return out


def _xlsx_table(ws, s: dict, data: dict, labels: dict, fmt) -> None:
    cols, rows, totals, decs = _table(s, data)
    pcts = set(s.get("percent_columns", []))
    red = s.get("negative_red", True)
    head = fmt(bold=True, font_color="#FFFFFF", bg_color="#1E293B")
    for c, name in enumerate(cols):
        ws.write_string(0, c, name, head)
        width = max([len(name)] + [len(row[0][c][0]) for row in rows[:200]])
        ws.set_column(c, c, min(max(10, width + 2), 50))
    for y, (cells, color) in enumerate(rows, 1):
        bg = {"bg_color": HIGHLIGHTS[color][1]} if color else {}
        for c, (text, n) in enumerate(cells):
            if n is None:
                ws.write_string(y, c, text, fmt(**bg))
            elif cols[c] in pcts:
                ws.write_number(y, c, n / 100, fmt(decs[c], pct=True, **bg))
            else:
                ws.write_number(y, c, n, fmt(decs[c], red=red, **bg))
    if totals:
        y = len(rows) + 1
        for c, (_, n) in enumerate(totals):
            if n is None:
                ws.write_string(y, c, labels["total"] if c == 0 else "", fmt(bold=True, bg_color="#E2E8F0", top=1))
            else:
                ws.write_number(y, c, n, fmt(decs[c], red=red, bold=True, bg_color="#E2E8F0", top=1))
    ws.freeze_panes(1, 0)
    if cols:
        ws.autofilter(0, 0, max(len(rows), 1), len(cols) - 1)

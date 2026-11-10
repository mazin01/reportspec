import re

import pytest

import reportspec as rs
from reportspec import SpecError

try:  # WeasyPrint needs Pango; without it (e.g. plain Windows) the PDF tests skip
    import weasyprint  # noqa: F401
    HAVE_PDF = True
except (ImportError, OSError):
    HAVE_PDF = False

ORDERS = [
    {"ref": "ORD-1", "item": "Printer paper", "qty": 4, "amount": "1,250.50", "status": "paid"},
    {"ref": "ORD-2", "item": "Desk lamp", "qty": 1, "amount": "(120.00)", "status": "refund"},
    {"ref": "ORD-3", "item": "Stapler", "qty": 2, "amount": "35", "status": "overdue"},
]
SPEC = {
    "title": "Monthly orders",
    "subtitle": "North branch",
    "footer": "Generated 2025-03-31",
    "sections": [
        {"type": "metrics", "title": "Summary", "items": [
            {"label": "Revenue", "data": "orders", "column": "amount", "fn": "sum"},
            {"label": "Orders", "data": "orders", "fn": "rows"},
            {"label": "Target", "value": 2000},
        ]},
        {"type": "table", "title": "Orders", "data": "orders", "number_columns": ["qty", "amount"],
         "totals": True, "highlight": {"column": "status", "map": {"overdue": "red", "refund": "yellow"}}},
        {"type": "text", "title": "Notes", "body": "Refunds are shown as negative amounts."},
    ],
}
AR_ORDERS = [{"المرجع": "ORD-1", "البيان": "ورق طباعة", "المبلغ": "١٬٢٥٠٫٥٠"},
             {"المرجع": "ORD-2", "البيان": "Desk lamp", "المبلغ": "99"}]
AR_SPEC = {"title": "تقرير الطلبات", "sections": [
    {"type": "metrics", "items": [{"label": "الإجمالي", "data": "o", "column": "المبلغ", "fn": "sum"}]},
    {"type": "table", "data": "o", "number_columns": ["المبلغ"], "totals": True}]}


# ── HTML (no native libraries needed) ──────────────────────────────────────────

def test_html_numbers_totals_highlights_and_escaping():
    page = rs.render_html({**SPEC, "title": "Orders <Q1> & more"}, {"orders": ORDERS})
    assert "Orders &lt;Q1&gt; &amp; more" in page
    assert '<td class="num">1,250.50</td>' in page
    assert '<td class="num neg">-120.00</td>' in page                 # "(120.00)" parsed as negative
    assert '<tr class="totals"><td>Total</td>' in page and ">1,165.50<" in page
    assert 'style="background:#fee2e2"' in page and 'style="background:#fef9c3"' in page
    assert ">1,165.50</div>" in page and ">3</div>" in page and ">2,000.00</div>" in page  # metrics
    assert '<html dir="ltr">' in page and "size: A4;" in page


def test_html_arabic_document_is_rtl_with_arabic_labels():
    page = rs.render_html(AR_SPEC, {"o": AR_ORDERS})
    assert '<html dir="rtl">' in page
    assert ">الإجمالي</td>" in page                    # totals label follows the document language
    assert '"صفحة " counter(page) " من "' in page        # page footer too
    assert '<td dir="ltr">Desk lamp</td>' in page       # an English cell inside an RTL table
    assert ">1,349.50<" in page                          # Arabic-Indic digits and separators summed


def test_wide_tables_go_landscape_and_labels_can_be_overridden():
    row = {f"c{i}": i for i in range(rs.WIDE_TABLE)}
    page = rs.render_html({"title": "W", "labels": {"page": "P."}, "sections": [{"type": "table", "data": "d"}]},
                          {"d": [row]})
    assert "size: A4 landscape;" in page and '"P. " counter(page)' in page


@pytest.mark.parametrize("spec,data,code", [
    ({"sections": [{"type": "chart"}]}, {}, "UNKNOWN_SECTION_TYPE"),
    ({"sections": [{"type": "table", "data": "nope"}]}, {}, "MISSING_DATA"),
    ({"sections": [{"type": "table", "data": "d", "number_columns": ["zzz"]}]}, {"d": [{"a": 1}]}, "COLUMN_NOT_FOUND"),
    ({"sections": [{"type": "metrics", "items": [{"data": "d", "column": "a", "fn": "median"}]}]}, {"d": []}, "BAD_SPEC"),
    ({"sections": [{"type": "table", "data": "d", "highlight": {"column": "a", "map": {"x": "pink"}}}]},
     {"d": [{"a": 1}]}, "BAD_SPEC"),
    ({"direction": "up", "sections": []}, {}, "BAD_SPEC"),
    ({"sections": [{"type": "table", "data": "d"}]}, {"d": ["not a dict"]}, "BAD_DATA"),
])
def test_errors_have_stable_codes(spec, data, code):
    with pytest.raises(SpecError) as e:
        rs.render_html(spec, data)
    assert e.value.code == code


def test_dataframe_input():
    pd = pytest.importorskip("pandas")
    page = rs.render_html({"sections": [{"type": "table", "data": "d", "number_columns": ["x"]}]},
                          {"d": pd.DataFrame({"x": [1.5, 2.5]})})
    assert ">2.50<" in page


# ── Excel ──────────────────────────────────────────────────────────────────────

def test_xlsx_numbers_formats_totals_and_sheets(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    out = rs.render_xlsx(SPEC, {"orders": ORDERS}, tmp_path / "r.xlsx")
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Summary", "Orders"]
    s = wb["Summary"]
    assert s["A1"].value == "Monthly orders"
    values = {s.cell(r, 1).value: s.cell(r, 2).value for r in range(1, s.max_row + 1)}
    assert values["Revenue"] == pytest.approx(1165.5) and values["Orders"] == 3
    o = wb["Orders"]
    assert [c.value for c in o[1]] == ["ref", "item", "qty", "amount", "status"]
    assert o["D2"].value == 1250.5 and o["D3"].value == -120.0       # real numbers, not text
    assert "[Red]" in o["D3"].number_format
    assert o["A5"].value == "Total" and o["D5"].value == pytest.approx(1165.5)
    assert o["A4"].fill.fgColor.rgb.endswith("FFC7CE")                # overdue → red
    assert o.freeze_panes == "A2"


def test_xlsx_rtl_sheets(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.load_workbook(rs.render_xlsx(AR_SPEC, {"o": AR_ORDERS}, tmp_path / "ar.xlsx"))
    sheet = wb.worksheets[1]
    assert sheet.sheet_view.rightToLeft
    assert sheet["C2"].value == 1250.5 and sheet["A4"].value == "الإجمالي"


# ── PDF (WeasyPrint; runs where Pango is installed: Linux, WSL, CI) ─────────────

pdf = pytest.mark.skipif(not HAVE_PDF, reason="WeasyPrint's native libraries (Pango) aren't available")


@pdf
def test_pdf_renders_with_bundled_font_and_text(tmp_path):
    pypdf = pytest.importorskip("pypdf")
    out = rs.render_pdf(SPEC, {"orders": ORDERS}, tmp_path / "r.pdf")
    reader = pypdf.PdfReader(out)
    text = "".join(p.extract_text() for p in reader.pages)
    assert "Monthly orders" in text and "1,250.50" in text and "Page 1 of 1" in text
    fonts = {f["/BaseFont"] for p in reader.pages for f in p["/Resources"]["/Font"].values()}
    assert fonts and all("Plex-Arabic" in str(f) for f in fonts), fonts  # embedded under its @font-face name


@pdf
def test_pdf_arabic_report(tmp_path):
    pypdf = pytest.importorskip("pypdf")
    out = rs.render_pdf(AR_SPEC, {"o": AR_ORDERS}, tmp_path / "ar.pdf")
    text = "".join(p.extract_text() for p in pypdf.PdfReader(out).pages)
    words = set(re.findall(r"[؀-ۿ]+", text))
    # Extractors may return Arabic in visual order; accept either reading of each word.
    for word in ("تقرير", "الطلبات", "الإجمالي"):
        assert word in words or word[::-1] in words, (word, words)
    assert "1,349.50" in text


def test_whole_number_columns_get_no_decimals_unless_asked():
    page = rs.render_html(SPEC, {"orders": ORDERS})
    assert '<td class="num">4</td>' in page and ">7</td>" in page  # qty and its total
    page = rs.render_html({"sections": [{**SPEC["sections"][1], "decimals": 1}]}, {"orders": ORDERS})
    assert '<td class="num">4.0</td>' in page


def test_dates_in_rtl_text_are_isolated():
    page = rs.render_html({**AR_SPEC, "footer": "أنشئ في 2025-03-31"}, {"o": AR_ORDERS})
    lri, pdi = chr(0x2066), chr(0x2069)
    assert f"{lri}2025-03-31{pdi}" in page
    assert lri not in rs.render_html({**SPEC, "footer": "Generated 2025-03-31"}, {"orders": ORDERS})

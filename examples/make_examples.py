"""Render the example reports in this folder (English and Arabic).

    python examples/make_examples.py    # PDFs need WeasyPrint's Pango (Linux, WSL, macOS)
"""
from pathlib import Path

import reportspec as rs

OUT = Path(__file__).parent

orders = [
    {"ref": f"ORD-{1000 + i}", "item": item, "qty": q, "amount": amt, "status": st}
    for i, (item, q, amt, st) in enumerate([
        ("Printer paper A4", 4, "1,250.50", "paid"), ("Desk lamp", 1, "(120.00)", "refund"),
        ("Stapler", 2, "35", "overdue"), ("Office chair", 3, "2,400.00", "paid"),
        ("Monitor stand", 5, "640.25", "paid"),
    ])
]
en = {"title": "Monthly orders", "subtitle": "North branch · March 2025", "footer": "Generated 2025-03-31",
      "sections": [
          {"type": "metrics", "title": "Summary", "items": [
              {"label": "Revenue", "data": "orders", "column": "amount", "fn": "sum"},
              {"label": "Orders", "data": "orders", "fn": "rows"},
              {"label": "Average order", "data": "orders", "column": "amount", "fn": "avg"}]},
          {"type": "table", "title": "Orders", "data": "orders", "number_columns": ["qty", "amount"], "totals": True,
           "highlight": {"column": "status", "map": {"overdue": "red", "refund": "yellow"}}},
          {"type": "text", "title": "Notes", "body": "Refunds are shown as negative amounts."}]}

ar_rows = [
    {"التاريخ": "2025-03-02", "المرجع": "ORD-1000", "البيان": "ورق طباعة (عبوة)", "المبلغ": "١٬٢٥٠٫٥٠", "الحالة": "مدفوع"},
    {"التاريخ": "2025-03-05", "المرجع": "ORD-1001", "البيان": "Desk lamp مصباح", "المبلغ": "(120.00)", "الحالة": "مسترد"},
    {"التاريخ": "2025-03-09", "المرجع": "ORD-1002", "البيان": "دباسة", "المبلغ": "35", "الحالة": "متأخر"},
]
ar = {"title": "تقرير الطلبات الشهري", "subtitle": "فرع الشمال · مارس ٢٠٢٥", "footer": "أُنشئ في 2025-03-31",
      "sections": [
          {"type": "metrics", "title": "الملخص", "items": [
              {"label": "الإيرادات", "data": "o", "column": "المبلغ", "fn": "sum"},
              {"label": "عدد الطلبات", "data": "o", "fn": "rows"}]},
          {"type": "table", "title": "الطلبات", "data": "o", "number_columns": ["المبلغ"], "totals": True,
           "highlight": {"column": "الحالة", "map": {"متأخر": "red", "مسترد": "yellow"}}},
          {"type": "text", "title": "ملاحظات", "body": "المبالغ المستردة تظهر بالسالب. Amounts in SAR."}]}

for name, spec, data in (("orders_en", en, {"orders": orders}), ("orders_ar", ar, {"o": ar_rows})):
    rs.render_xlsx(spec, data, OUT / f"{name}.xlsx")
    try:
        rs.render_pdf(spec, data, OUT / f"{name}.pdf")
    except OSError as exc:  # WeasyPrint without Pango
        print(f"{name}.pdf skipped: {exc.__class__.__name__}")
print("done")

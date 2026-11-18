# -*- coding: utf-8 -*-
"""تقرير مدير المبيعات: الترتيب الحرفي للأعمدة، دقة الحسابات، الفلاتر، المركز الأول، والتصدير بالتمييز الأصفر."""
import os
import time

import pytest

from conftest import ROOT, buy, sell

COLS = ["Salesman Name", "CUSTOMER CODE", "CUSTOMER NAME", "DISTRIBUTOR SKU CODE", "PRODUCT DESCRIPTION", "SALES TYPE",
        "INVOICE DATE", "QNT", "Price", "Sales Amount", "INVOICE_NO", "Net Cost", "Profit", "GP", "GP%", "GPM%", "COGS"]
P = dict(d_from="2026-01-01", d_to="2026-12-31")


@pytest.fixture
def data(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    F.save_rep(None, dict(name="خالد", commission_pct=1, active=1))
    F.save_partner_record("customer", cust[0], cust[1], rep="خالد", rep_date="2026-01-01")
    other = (F.save_partner_record("customer", None, "عميل آخر"), "عميل آخر")
    buy(F, supp, "ELH-T1", 100, 10, d="2026-02-01")
    si, _, _ = sell(F, cust, "ELH-T1", 10, 25, d="2026-02-05", discount=10)      # 240 صافي، تكلفة 100
    sell(F, other, "ELH-T1", 5, 20, d="2026-02-06")                               # 100، تكلفة 50
    F.save_return(F.SALES_RETURN, "2026-02-07", cust, ("ELH-T1", "صنف أول"), 2, 24)  # -48، تكلفة -20
    F.save_post_discount("sales", "2026-02-08", cust, 12, invoice_id=si)          # إشعار دائن -12
    return F, cust, other, si


def test_first_in_catalog_and_exact_columns(erp):
    F = erp
    assert F.REPORT_CATALOG[0][1][0][0] == "sales_mgr"
    assert F.REPORT_GROUP_OF["sales_mgr"] == "reports.salesmgr" and "reports.salesmgr" in F.ALL_PERMISSION_KEYS
    rep = F.run_report("sales_mgr", P)
    assert [h for h, _ in rep.columns] == COLS
    assert [COLS[i] for i in rep.highlight] == ["Sales Amount", "GP"] and rep.ltr


def test_calculations(data):
    F, cust, other, si = data
    rep = F.run_report("sales_mgr", P)
    by = {(r[1], r[5]): r for r in rep.rows}
    s = by[(cust[0], "Sales")]
    assert s[0] == "خالد" and s[3] == "ELH-T1" and s[6] == "2026-02-05"
    assert (s[7], s[8], s[9]) == (10, 24, 240)                 # QNT × Price = Sales Amount (بعد الخصم الخاص)
    assert (s[11], s[12], s[13], s[16]) == (10, 14, 140, 100)  # Net Cost, Profit/unit, GP, COGS
    assert s[14] == pytest.approx(140.0) and s[15] == pytest.approx(140 / 240 * 100)
    r = by[(cust[0], "Return")]
    assert r[7] == -2 and r[9] == -48 and r[16] == -20 and r[13] == -28 and r[10].startswith("RET-")
    c = by[(cust[0], "Credit Note")]
    assert c[9] == -12 and c[16] == 0 and c[7] is None
    assert by[(other[0], "Sales")][0] == F.NO_REP
    # الإجماليات = صافي المبيعات ومجمل الربح في قائمة الدخل
    t = rep.total
    assert t[9] == pytest.approx(240 + 100 - 48 - 12) and t[16] == pytest.approx(100 + 50 - 20)
    assert t[13] == pytest.approx(t[9] - t[16]) and t[15] == pytest.approx(t[13] / t[9] * 100)
    txt = F.build_income_report("2026-01-01", "2026-12-31")
    assert "280.00" in txt and "150.00" in txt              # صافي المبيعات 280 ومجمل الربح 150


def test_filters(data):
    F, cust, other, si = data
    assert {r[1] for r in F.run_report("sales_mgr", dict(P, customer=other[0])).rows} == {other[0]}
    assert {r[0] for r in F.run_report("sales_mgr", dict(P, rep="خالد")).rows} == {"خالد"}
    assert not F.run_report("sales_mgr", dict(d_from="2026-03-01", d_to="2026-03-31")).rows


def test_exports_keep_order_and_yellow(data, tmp_path):
    F = data[0]
    from openpyxl import load_workbook
    rep = F.run_report("sales_mgr", P)
    x = str(tmp_path / "smr.xlsx")
    F.export_report_xlsx(x, rep)
    ws = load_workbook(x).active
    head = next(r for r in ws.iter_rows(values_only=True) if r and r[0] == "Salesman Name")
    assert list(head[:17]) == COLS and not ws.sheet_view.rightToLeft
    hrow = next(c.row for c in ws["A"] if c.value == "Salesman Name")
    assert ws.cell(hrow, 10).fill.fgColor.rgb.endswith("FACC15") and ws.cell(hrow + 1, 14).fill.fgColor.rgb.endswith("FEF08A")
    assert not ws.cell(hrow + 1, 9).fill.fgColor.rgb.endswith("FEF08A")
    pdf = str(tmp_path / "smr.pdf")
    F.build_report_pdf(pdf, rep)
    assert os.path.getsize(pdf) > 2000


def test_web_shows_it_first_with_filters(data):
    F = data[0]
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(os.path.join(ROOT, "streamlit_app.py"), default_timeout=120)
    at.session_state["user"] = F.authenticate(F.DEFAULT_ADMIN_USER, "Elhagry2012@")
    at.session_state["last_seen"] = time.time()
    at.session_state["nav"] = "reports"
    at.run()
    assert not at.exception, [e.message for e in at.exception]
    assert at.selectbox(key="rep_key").value == "sales_mgr"          # أول تقرير يظهر افتراضياً
    assert any("Sales Manager Report" in m.value for m in at.markdown)
    at.selectbox(key="smr_rep").set_value("خالد").run()
    assert not at.exception


def test_desktop_marks_highlighted_columns(data, monkeypatch):
    tk = pytest.importorskip("tkinter")
    F = data[0]
    from elhagry_desktop import app as A
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("لا توجد شاشة متاحة لـ Tk")
    root.withdraw()
    errors = []
    monkeypatch.setattr(A, "handle_error", lambda e: errors.append(e))
    try:
        erp = A.ElhagryERP(root)
        erp.show_page("reports")
        erp.on_reports_shown()
        assert erp.rep_catalog.selection() == ("sales_mgr",)
        erp.rep_from_e.delete(0, "end")
        erp.rep_from_e.insert(0, "2026-01-01")
        erp.rep_rep_cb.set("خالد")
        erp.rep_current = "sales_mgr"
        erp.run_current_report()
        assert not errors, errors
        assert str(erp.rep_tree.heading("c9", "text")).startswith("🟨") and str(erp.rep_tree.heading("c13", "text")).startswith("🟨")
        assert all(erp.rep_tree.item(i, "values")[0] in ("خالد", "الإجمالي") for i in erp.rep_tree.get_children())
    finally:
        root.destroy()

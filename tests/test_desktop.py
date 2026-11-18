# -*- coding: utf-8 -*-
"""واجهة سطح المكتب: بناء كل الصفحات (ومنها الجديدة) وتشغيل محمّلاتها دون أخطاء. تُتخطى بدون Tk/شاشة."""
import pytest

from conftest import buy, sell

tk = pytest.importorskip("tkinter")


@pytest.fixture
def app(world, monkeypatch):
    from elhagry_desktop import app as A
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("لا توجد شاشة متاحة لـ Tk")
    root.withdraw()
    errors = []
    monkeypatch.setattr(A, "handle_error", lambda e: errors.append(e))
    for name in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(A.messagebox, name, lambda *a, **k: None)
    monkeypatch.setattr(A.messagebox, "askyesno", lambda *a, **k: True)
    erp = A.ElhagryERP(root)
    yield erp, errors, world
    try:
        erp.shutdown()
    except Exception:  # noqa: BLE001
        pass
    root.destroy()


def test_all_pages_build_and_load(app):
    erp, errors, w = app
    F = w["F"]
    buy(F, w["supp"], "ELH-T1", 20, 10)
    sell(F, w["cust"], "ELH-T1", 5, 20)
    for key in erp.frames:
        erp.show_page(key)
        erp.root.update_idletasks()
        loader = erp.tab_loaders.get(str(erp.frames[key]))
        if loader:
            loader()
    assert not errors, errors


def test_new_pages_actions(app):
    erp, errors, w = app
    F = w["F"]
    buy(F, w["supp"], "ELH-T1", 20, 10)
    sell(F, w["cust"], "ELH-T1", 5, 20)
    erp.load_ledger_page()
    assert "متوازن ✓" in erp.gl_status_lbl.cget("text")
    for i in range(len(erp.gl_rep_cb["values"])):
        erp.gl_rep_cb.current(i)
        erp.run_ledger_report()
    erp.run_cash_forecast()
    erp.load_planning_page()
    erp.bud_line_cb.set(F.BUILTIN_LINES[F.LINE_SALES])
    erp.bud_month_e[0].insert(0, "1000")
    erp.save_budget_line()
    assert F.budget_year_matrix(int(F.today()[:4]))[F.LINE_SALES][0] == 1000
    erp.run_bva()
    erp.run_commissions()
    erp.run_credit_report()
    erp.load_audit_page()
    assert "سليمة" in erp.au_chain_lbl.cget("text")
    assert erp.au_box["tree"].get_children()
    assert not errors, errors


def test_customer_credit_fields_and_block_override(app):
    erp, errors, w = app
    F = w["F"]
    code = w["cust"][0]
    st = erp.pp["customer"]
    erp.load_partner_page("customer")
    st["tree"].selection_set(code)
    erp._pp_edit("customer")
    erp._pp_set(st["cr_limit"], "100")
    erp._pp_set(st["cr_days"], "15")
    st["cr_policy"].set(F.CREDIT_POLICIES[F.POLICY_BLOCK])
    erp._pp_save("customer")
    with F.db() as conn:
        assert F.partner_credit_terms(conn, code)[:3] == (100, 15, F.POLICY_BLOCK)
    buy(F, w["supp"], "ELH-T1", 50, 10)
    # منع: المدير يوافق صراحة (askyesno = True في الاختبار) فيُحفظ الأمر بتجاوز الحد
    oid = erp.with_credit(F.create_order, "sales", F.today(), w["cust"], "", [F.make_line("ELH-T1", "x", "KG", 10, 20)])
    assert oid
    assert not errors, errors

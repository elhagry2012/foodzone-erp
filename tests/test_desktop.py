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


def test_dashboard_sources_and_hw_pl_compare_windows(app):
    erp, errors, w = app
    erp.show_page("dashboard")
    erp.root.update_idletasks()
    assert "OpEx" in erp.dash_formula.cget("text") or "بيعية" in erp.dash_formula.cget("text")
    erp.open_dashboard_sources()
    erp.show_page("import")
    erp.root.update_idletasks()
    for loader in erp.tab_loaders.values():
        try:
            loader()
        except Exception:  # noqa: BLE001
            pass
    if getattr(erp, "hw", None) and erp.hw.get("tree") is not None:
        erp._hw_show_pl_compare()
        erp._hw_agg_refresh()
    assert not errors, errors


def test_voucher_attachments_window_and_security_form(app, tmp_path, monkeypatch):
    erp, errors, w = app
    F = w["F"]
    from elhagry_core import engine as E
    monkeypatch.setattr(E, "ATTACH_DIR", str(tmp_path / "att"))
    vid = F.save_voucher("قبض", "2026-01-11", "صندوق نقدية الرياض", F.CAT_PARTNER, w["cust"], 50)
    F.add_attachment("voucher", vid, "r.pdf", b"%PDF")
    erp.show_page("cash")
    erp.root.update_idletasks()
    erp.load_cash_transactions()
    assert erp.cash_tree.item(str(vid))["values"][-1] == "📎1"
    erp.cash_tree.selection_set(str(vid))
    erp.open_voucher_attachments_selected()
    erp.show_page("company")
    erp.root.update_idletasks()
    for name in ("admin_pwd_e", "reset_pwd_e"):
        if hasattr(erp, name):
            assert getattr(erp, name).get() == ""        # كلمات المرور لا تُعرض أبداً
    assert not errors, errors


def test_forgot_password_dialog_and_admin_requests_panel(app):
    erp, errors, w = app
    F = w["F"]
    from elhagry_desktop import app as A
    F.save_user_account(None, "clerk.x", "موظف", "Clerk-1234", "Clerk-1234", False, True, ["dashboard"])
    login = A.LoginWindow.__new__(A.LoginWindow)        # نموذج الدخول بدون فتح نافذة رئيسية جديدة
    login.win = tk.Toplevel(erp.root)
    login.user_e = tk.Entry(login.win)
    login.user_e.insert(0, "clerk.x")
    dlg = login.forgot_password()
    dlg.update_idletasks()
    dlg.destroy()
    login.win.destroy()
    F.request_password_reset("clerk.x", "من الاختبار")
    erp.show_page("users")
    erp.root.update_idletasks()
    erp.load_pw_requests()
    rows = erp.pwreq_tree.get_children()
    assert len(rows) == 1
    erp.pwreq_tree.selection_set(rows[0])
    erp.pwreq_pw1.insert(0, "Brand-New-1")
    erp.pwreq_pw2.insert(0, "Brand-New-1")
    erp.resolve_pw_request(True)
    assert F.authenticate("clerk.x", "Brand-New-1")
    assert not errors, errors

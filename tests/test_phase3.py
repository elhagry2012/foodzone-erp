# -*- coding: utf-8 -*-
"""المرحلة الثالثة: العمولات المرنة، التقارير التفاعلية والسداد السريع، الدورة المستندية المتصلة، ومحاذاة الجداول."""
import os
import time

import pytest

from conftest import ROOT, buy, sell

D = "2026-03-10"


# ============================================================ المحرك
def _rep_world(F, cust):
    F.save_rep(None, dict(name="سالم", commission_pct=2.0, active=1))
    F.save_partner_record("customer", cust[0], cust[1], rep="سالم", rep_date="2026-01-01")


def test_commission_rates_uniform_override_and_save(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    _rep_world(F, cust)
    buy(F, supp, "ELH-T1", 50, 10, d="2026-03-01")
    si, _, _ = sell(F, cust, "ELH-T1", 10, 100, d=D)                 # صافي 1,000
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 575, ("sales", si))
    p = dict(d_from="2026-03-01", d_to="2026-03-31")
    get = lambda rows: next(r for r in rows if r[1] == "سالم")       # noqa: E731
    assert get(F.commission_rows(p)[0])[6] == 20.0                   # نسبة السجل 2%
    assert get(F.commission_rows(p, uniform=5)[0])[6] == 50.0        # موحدة 5%
    assert get(F.commission_rows(p, rates={"سالم": 3}, uniform=5)[0])[6] == 30.0   # المباشر يتقدم على الموحدة
    assert get(F.commission_rows(p, uniform=10, base="collections")[0])[6] == 57.5
    with pytest.raises(F.UserError):
        F.commission_rows(p, uniform=150)
    assert F.save_commission_rates({"سالم": 4.5, "غير موجود": 1}) == 1
    assert get(F.commission_rows(p)[0])[2] == 4.5
    rep = F.run_report("commissions", dict(p, uniform=5))
    assert "نسبة موحدة 5%" in rep.notes[0] and rep.total[6] == 50.0


def test_payment_prefill_and_invoice_lookup(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 10, 100, d=D)                 # 1,150
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 150, ("sales", si))
    F.save_post_discount("sales", D, cust, 100, invoice_id=si)       # 115 إشعار دائن
    pre = F.invoice_payment_prefill("sales", si)
    assert pre["remaining"] == 885.0 and pre["t_type"] == "قبض" and pre["alloc"] == ("sales", si)
    assert F.find_invoice_id("sales", pre["no"]) == si
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", pre["category"], pre["party"], pre["remaining"], pre["alloc"])
    assert F.invoice_payment_prefill("sales", si)["remaining"] == 0.0


# ============================================================ الويب
st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(ROOT, "streamlit_app.py")


def _at(F, page, **state):
    at = st_testing.AppTest.from_file(APP, default_timeout=120)
    at.session_state["user"] = F.authenticate(F.DEFAULT_ADMIN_USER, "Elhagry2012@")
    at.session_state["last_seen"] = time.time()
    at.session_state["nav"] = page
    for k, v in state.items():
        at.session_state[k] = v
    at.run()
    assert not at.exception, [e.message for e in at.exception]
    return at


def _btn(at, text):
    return next(b for b in at.button if text in (b.label or ""))


def test_web_order_to_delivery_to_invoice(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    oid = F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 8, 30)])
    at = _at(F, "sales", flow_view_sales="ord", ord_last_sales=oid)
    _btn(at, f"ترحيل الأمر {oid} إلى إذن تسليم").click().run()
    assert at.session_state["flow_view_sales"] == "ful" and at.session_state["ful_order_sales"] == oid
    assert not at.exception
    _btn(at, "حفظ إذن التسليم").click().run()
    with F.db() as conn:
        fid, = conn.execute("SELECT delivery_id FROM sales_deliveries WHERE order_id=?", (oid,)).fetchone()
        assert conn.execute("SELECT SUM(quantity) FROM sales_delivery_lines WHERE delivery_id=?", (fid,)).fetchone()[0] == 8
    _btn(at, f"إصدار فاتورة من إذن التسليم رقم {fid}").click().run()
    assert at.session_state["flow_view_sales"] == "review"
    assert any("مراجعة" in m.value for m in at.markdown)
    _btn(at, "اعتماد وإصدار الفاتورة").click().run()
    assert not at.exception
    with F.db() as conn:
        assert conn.execute("SELECT invoice_id FROM sales_deliveries WHERE delivery_id=?", (fid,)).fetchone()[0]


def test_web_drill_report_statement_invoice_pay(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 10, 100, d=D)
    at = _at(F, "reports", rep_group="💳 الذمم والمديونيات", rep_key="rec")
    pick = next(s for s in at.selectbox if s.key and s.key.startswith("drill_pick_rec"))
    pick.set_value(0).run()
    _btn(at, "فتح في تبويب").click().run()
    # بدون نوافذ منبثقة: الكشف يُفتح كتبويب مجاور والتقرير باقٍ في تبويبه
    assert at.session_state["nav"] == "reports" and at.session_state["ws_active"] == "statement"
    assert at.session_state["stmt_party"] == cust[0]
    at = _at(F, "statement", stmt_party=cust[0])
    inv_pick = next(s for s in at.selectbox if s.key == f"stmt_inv_pick_{cust[0]}")
    keys = F.fetch_partner_statement(cust[0], True)[5]
    inv_pick.set_value(keys.index(("sales", si))).run()
    _btn(at, "عرض ↗").click().run()
    assert not at.exception
    pay = _btn(at, "سداد هذه الفاتورة")
    pay.click().run()
    assert at.session_state["nav"] == "statement" and at.session_state["ws_active"] == "pay"
    assert at.number_input(key="cv_amount").value == 1150.0
    assert tuple(at.session_state["cv_inv"]) == ("sales", si)
    _btn(at, "حفظ السند").click().run()
    assert not at.exception
    with F.db() as conn:
        assert conn.execute("SELECT status FROM sales_invoices WHERE invoice_id=?", (si,)).fetchone()[0] != "مفتوحة"


def test_web_interactive_commissions(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    _rep_world(F, cust)
    buy(F, supp, "ELH-T1", 50, 10, d=F.today())
    sell(F, cust, "ELH-T1", 10, 100, d=F.today())
    at = _at(F, "planning", com_rates={"سالم": 7.0})
    assert any(m.label == "إجمالي العمولات المستحقة" and m.value == "70.00" for m in at.metric)
    at.toggle(key="com_use_uni").set_value(True).run()
    at.number_input(key="com_uni").set_value(1.0).run()
    assert any(m.label == "إجمالي العمولات المستحقة" and m.value == "70.00" for m in at.metric)   # المباشر يتقدم
    _btn(at, "إلغاء التعديلات").click().run()
    assert any(m.label == "إجمالي العمولات المستحقة" and m.value == "10.00" for m in at.metric)   # 1% موحدة


def test_web_tables_aligned_wrapped_and_mobile_cards():
    src = open(APP, encoding="utf-8").read()
    assert 'alignment="right"' in src and 'row_height' in src and "fz-cards" in src
    assert '[class*="st-key-fzt_c_"] [data-testid="stDataFrame"],' in src


# ============================================================ سطح المكتب
tk = pytest.importorskip("tkinter")


@pytest.fixture
def desk(world, monkeypatch):
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
    monkeypatch.setattr(A.messagebox, "askyesno", lambda *a, **k: False)
    erp = A.ElhagryERP(root)
    yield erp, errors, world, A
    for w in root.winfo_children():
        if isinstance(w, tk.Toplevel):
            w.destroy()
    root.destroy()


def test_desktop_headers_aligned_with_data(desk):
    erp, errors, w, A = desk
    t = erp.pp["customer"]["tree"]
    for col in t["columns"]:
        assert str(t.heading(col, "anchor")) == str(t.column(col, "anchor"))
    A.configure_tree(erp.rep_tree, [("الحساب", "t"), ("المبلغ", "m")])
    assert str(erp.rep_tree.heading("c0", "anchor")) == "e" and str(erp.rep_tree.heading("c1", "anchor")) == "center"
    erp.rep_tree.insert("", "end", values=("نص طويل جداً " * 12, "1,000.00"))
    erp.rep_tree._autofit()
    assert int(erp.rep_tree.column("c0", "width")) > 400        # العمود اتسع للنص


def test_desktop_order_to_fulfil_to_invoice(desk):
    erp, errors, w, A = desk
    F = w["F"]
    buy(F, w["supp"], "ELH-T1", 50, 10)
    oid = F.create_order("sales", F.today(), w["cust"], "", [F.make_line("ELH-T1", "x", "KG", 6, 30)])
    erp.refresh_flow("sales")
    erp.order_go_fulfil("sales", oid)
    rows = erp.flow["sales"]["f_rows"]
    assert rows and rows[0]["qty_e"].get() == "6"                      # الكميات منقولة تلقائياً
    erp.save_fulfilment("sales")
    with F.db() as conn:
        fid, = conn.execute("SELECT delivery_id FROM sales_deliveries WHERE order_id=?", (oid,)).fetchone()
    erp.goto_invoice_for("sales", fid)
    assert erp.flow["sales"]["u_tree"].selection() == (f"sales:{fid}",)
    review = [x for x in erp.root.winfo_children() if isinstance(x, tk.Toplevel)]
    assert review, "review window should open"
    erp.make_flow_invoice("sales")
    with F.db() as conn:
        assert conn.execute("SELECT invoice_id FROM sales_deliveries WHERE delivery_id=?", (fid,)).fetchone()[0]
    assert not errors, errors


def test_desktop_drill_and_quick_pay(desk):
    erp, errors, w, A = desk
    F = w["F"]
    buy(F, w["supp"], "ELH-T1", 50, 10)
    si, _, _ = sell(F, w["cust"], "ELH-T1", 4, 50)
    erp.open_statement_for(w["cust"][0], "customer")
    assert erp.stmt_tree.get_children()
    erp.prefill_payment("sales", si)
    assert erp.c_amount.get() == "230.00" and erp.open_inv_cb.selected() == ("sales", si)
    erp.save_cash_transaction()
    with F.db() as conn:
        assert conn.execute("SELECT status FROM sales_invoices WHERE invoice_id=?", (si,)).fetchone()[0] != "مفتوحة"
    erp.rep_current = "rec"
    erp.run_current_report()
    assert erp.rep_tree.get_children() is not None
    assert not errors, errors


def test_desktop_interactive_commissions(desk):
    erp, errors, w, A = desk
    F, cust = w["F"], w["cust"]
    _rep_world(F, cust)
    buy(F, w["supp"], "ELH-T1", 50, 10, d=F.today())
    sell(F, cust, "ELH-T1", 10, 100, d=F.today())
    erp.com_uniform_e.insert(0, "3")
    erp.run_commissions()
    assert "30.00" in erp.com_total_lbl.cget("text")
    erp._com_rates["سالم"] = 6.0
    erp.run_commissions()
    assert "60.00" in erp.com_total_lbl.cget("text")
    assert not errors, errors


def test_web_order_cart_entry(world):
    """إدخال أمر بيع من الويب بنموذج «إضافة صنف» المريح للمس (بدون جدول عريض)."""
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    at = _at(F, "sales", flow_view_sales="ord")
    at.selectbox(key="ord_item_sales").set_value("ELH-T1").run()
    at.number_input(key="ord_qty_sales").set_value(4.0).run()
    at.number_input(key="ord_price_sales_ELH-T1").set_value(25.0).run()
    at.text_input(key="ord_disc_sales").set_value("10%").run()
    _btn(at, "إضافة للأمر").click().run()
    assert not at.exception and len(at.session_state["ord_cart_sales"]) == 1
    at.selectbox(key="ord_party_sales").set_value(cust[0]).run()
    _btn(at, "حفظ أمر البيع").click().run()
    assert not at.exception
    with F.db() as conn:
        q, price, disc = conn.execute("SELECT quantity, unit_price, discount_amount FROM sales_order_lines").fetchone()
    assert (q, price, disc) == (4, 25, 10)
    assert any("ترحيل الأمر" in (b.label or "") for b in at.button)


def test_desktop_quick_pay_survives_event_loop_and_trail(desk):
    """السبب الجذري: حدث تغيير الصفحة المؤجل كان يمسح الطرف والفاتورة بعد التعبئة. نُفرغ طابور الأحداث كما في التشغيل الفعلي."""
    erp, errors, w, A = desk
    F = w["F"]
    buy(F, w["supp"], "ELH-T1", 50, 10)
    with F.db() as conn:
        pid = conn.execute("SELECT MAX(purchase_id) FROM purchase_invoices WHERE supplier_code=?",
                           (w["supp"][0],)).fetchone()[0]
    erp.show_page("reports")
    erp.rep_current = "pay"
    erp._trail_push_report()
    erp.open_statement_for(w["supp"][0], "supplier")
    erp.root.update()
    erp.show_transaction_details(("purchase", pid))
    erp.prefill_payment("purchase", pid)
    erp.root.update()                                   # معالجة <<NotebookTabChanged>> المؤجل
    assert erp.party_name_cb.selected()[0] == w["supp"][0]
    assert erp.open_inv_cb.selected() == ("purchase", pid)
    assert erp.c_amount.get() and float(erp.c_amount.get()) > 0
    assert [k for k, _, _ in erp.trail] == ["report", "statement", "invoice", "pay"]
    erp.trail_open(1)                                   # الرجوع لكشف المورد بنقرة
    erp.root.update()
    assert erp._active_page == "statement" and erp.stmt_shown_code == w["supp"][0]
    assert len(erp.trail) == 4
    assert not errors, errors

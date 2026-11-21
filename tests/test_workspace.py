# -*- coding: utf-8 -*-
"""مساحة العمل بتبويبات متجاورة (بدون نوافذ منبثقة) + السداد الذكي + بقاء الحالة عند التنقل بين الشاشات."""
import os
import time

import pytest

from conftest import ROOT, buy, sell

st_testing = pytest.importorskip("streamlit.testing.v1")
APP = os.path.join(ROOT, "streamlit_app.py")
D = "2026-03-10"


def _at(F, page, **state):
    at = st_testing.AppTest.from_file(APP, default_timeout=120)
    at.session_state["user"] = F.authenticate(F.main_admin_user(), os.environ["ELHAGRY_ADMIN_PASSWORD"])
    at.session_state["last_seen"] = time.time()
    at.session_state["nav"] = page
    for k, v in state.items():
        at.session_state[k] = v
    at.run()
    _ok(at)
    return at


def _ok(at):
    assert not at.exception, [e.message for e in at.exception]
    assert not [e.value for e in at.error], [e.value for e in at.error]
    assert not [w.value for w in at.warning if "Session State" in (w.value or "")]


def _btn(at, text):
    return next(b for b in at.button if text in (b.label or ""))


def _ws_labels(at):
    return [t.label for t in at.tabs]


def test_no_dialogs_in_reports_or_invoices():
    src = open(APP, encoding="utf-8").read()
    assert "invoice_dialog" not in src
    assert src.count("@st.dialog") == 1          # تفاصيل الصنف فقط (خارج مسار التقارير والفواتير)


def test_report_statement_invoice_pay_stay_side_by_side(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 10, 100, d=D)                      # 1,150
    at = _at(F, "reports", rep_group="💳 الذمم والمديونيات", rep_key="rec")
    next(s for s in at.selectbox if (s.key or "").startswith("drill_pick_rec")).set_value(0).run()
    _btn(at, "فتح في تبويب").click().run()
    _ok(at)
    labels = _ws_labels(at)
    assert at.session_state["nav"] == "reports"
    assert labels[0].startswith("📊") and cust[1] in labels[1]

    keys = F.fetch_partner_statement(cust[0], True)[5]
    next(s for s in at.selectbox if s.key == f"stmt_inv_pick_{cust[0]}").set_value(keys.index(("sales", si))).run()
    _btn(at, "عرض ↗").click().run()
    _ok(at)
    assert at.session_state["ws_active"] == "invoice" and len(_ws_labels(at)) == 3

    _btn(at, "سداد هذه الفاتورة").click().run()
    _ok(at)
    assert at.session_state["nav"] == "reports" and at.session_state["ws_active"] == "pay"
    assert len(_ws_labels(at)) == 4                                        # التقرير + الكشف + الفاتورة + السداد
    # التقرير والكشف ما زالا محفوظين بحالتهما
    assert at.session_state["rep_key"] == "rec" and at.session_state["stmt_party"] == cust[0]
    # الطرف + الفاتورة + المبلغ + البيان معبأة تلقائياً
    assert tuple(at.session_state["cv_party"])[0] == cust[0]
    assert tuple(at.session_state["cv_inv"]) == ("sales", si)
    assert at.number_input(key="cv_amount").value == 1150.0
    assert F.invoice_number("sales", si) in at.text_input(key="cv_desc").value
    assert any(m.label == "المتبقي بعد السداد" for m in at.metric)

    _btn(at, "حفظ السند").click().run()
    _ok(at)
    assert F.invoice_payment_prefill("sales", si)["remaining"] == 0.0
    assert at.session_state["ws_active"] == "invoice"                      # يعود لتبويب الفاتورة لعرض حالتها
    assert any("مسددة بالكامل" in s.value for s in at.success)
    assert len(_ws_labels(at)) == 4


def test_closing_tabs(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 2, 50, d=D)
    at = _at(F, "statement", stmt_party=cust[0], ws_open=["statement", "invoice"], inv_view=("sales", si))
    assert len(_ws_labels(at)) == 2
    at.button(key="ws_x_invoice").click().run()
    _ok(at)
    assert not at.tabs
    at = _at(F, "statement", stmt_party=cust[0], ws_open=["invoice"], inv_view=("sales", si))
    at.button(key="ws_close_all").click().run()
    _ok(at)
    assert not at.tabs


def test_quick_pay_selects_party_by_code_even_after_rename(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 4, 50, d=D)                        # 230
    F.save_partner_record("customer", cust[0], cust[1] + " (معدل)")       # اسم الطرف تغيّر بعد الفاتورة
    at = _at(F, "statement", stmt_party=cust[0], ws_open=["invoice"], inv_view=("sales", si))
    _btn(at, "سداد هذه الفاتورة").click().run()
    _ok(at)
    assert tuple(at.session_state["cv_party"]) == (cust[0], cust[1] + " (معدل)")
    assert tuple(at.session_state["cv_inv"]) == ("sales", si)
    assert at.number_input(key="cv_amount").value == 230.0


def test_flow_invoice_inline_and_pay_returns(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 2, 50, d=F.today())                # 115
    at = _at(F, "sales", flow_view_sales="inv", inv_inline_sales=si)
    assert any("فاتورة ضريبية" in m.value for m in at.markdown)            # التفاصيل داخل الصفحة لا منبثقة
    _btn(at, "سداد هذه الفاتورة").click().run()
    _ok(at)
    assert at.session_state["nav"] == "cash" and tuple(at.session_state["cv_inv"]) == ("sales", si)
    assert at.number_input(key="cv_amount").value == 115.0
    at.button(key="cash_back").click().run()
    _ok(at)
    assert at.session_state["nav"] == "sales" and at.session_state["inv_inline_sales"] == si


def test_state_survives_leaving_the_page(world):
    F, cust = world["F"], world["cust"]
    at = _at(F, "statement", stmt_party=cust[0])
    at.session_state["nav"] = "dashboard"
    at.run()
    _ok(at)
    at.session_state["nav"] = "statement"
    at.run()
    _ok(at)
    assert at.session_state["stmt_party"] == cust[0]


def test_choosing_invoice_fills_amount_and_description(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 2, 50, d=D)                        # 115
    with F.db() as conn:
        party = conn.execute("SELECT partner_code, name FROM business_partners WHERE partner_code=?", (cust[0],)).fetchone()
    at = _at(F, "cash", cv_type="قبض", cv_cat=F.CAT_PARTNER, cv_party=tuple(party))
    at.selectbox(key="cv_inv").set_value(("sales", si)).run()
    _ok(at)
    assert at.number_input(key="cv_amount").value == 115.0
    assert F.invoice_number("sales", si) in at.text_input(key="cv_desc").value

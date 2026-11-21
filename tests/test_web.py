# -*- coding: utf-8 -*-
"""واجهة الويب: كل الصفحات (ومنها الجديدة) تعمل بدون استثناءات، والتصميم المتجاوب مضمّن."""
import os

import pytest

from conftest import ROOT, buy, sell

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

APP = os.path.join(ROOT, "streamlit_app.py")
PAGES = ["dashboard", "customers", "suppliers", "reps", "import", "inventory", "sales", "purchase", "statement",
         "discounts", "income", "reports", "journal", "ledger", "planning", "coa", "company", "users", "audit", "cash"]


@pytest.fixture
def web(world):
    F = world["F"]
    buy(F, world["supp"], "ELH-T1", 20, 10)
    sell(F, world["cust"], "ELH-T1", 5, 20)
    sell(F, world["cust"], "ELH-T1", 2, 20, invoice=False)
    F.save_partner_record("customer", world["cust"][0], world["cust"][1], credit=(100, 10, "warn"))
    user = F.authenticate(F.main_admin_user(), os.environ["ELHAGRY_ADMIN_PASSWORD"])
    return F, user


def _run(user, page, **state):
    import time
    at = AppTest.from_file(APP, default_timeout=120)
    at.session_state["user"] = user
    at.session_state["last_seen"] = time.time()
    at.session_state["nav"] = page
    for k, v in state.items():
        at.session_state[k] = v
    at.run()
    return at


@pytest.mark.parametrize("page", PAGES)
def test_every_page_renders(web, page):
    F, user = web
    at = _run(user, page)
    assert not at.exception, [e.message for e in at.exception]
    assert not [e.value for e in at.error], [e.value for e in at.error]


def test_ledger_page_shows_balanced_and_bs(web):
    F, user = web
    at = _run(user, "ledger")
    assert any("متوازن" in m.value for m in at.metric)
    assert any("الميزانية العمومية" in m.value for m in at.markdown)


def test_planning_tabs_and_audit_entries(web):
    F, user = web
    at = _run(user, "planning")
    labels = [t.label for t in at.tabs]
    assert labels[:4] == ["💧 توقع التدفق النقدي", "🎯 الموازنة التقديرية", "💼 عمولات المناديب", "🛡️ حدود الائتمان"]
    at = _run(user, "audit")
    assert any("سلسلة البصمات سليمة" in s.value for s in at.success)


def test_audit_records_web_user(web):
    F, user = web
    at = _run(user, "customers")
    assert not at.exception
    # المستخدم يُسجَّل لكل عملية في جلسة الويب
    F.set_current_user(user["username"], "web")
    F.save_partner_record("customer", "C0001", "اسم من الويب")
    row = F.list_audit_log(user=user["username"])[0]
    assert row[3] == "web"


def test_responsive_css_present():
    src = open(APP, encoding="utf-8").read()
    for needle in ("@media (max-width: 1024px)", "@media (max-width: 640px)", "overflow-x:auto",
                   "-webkit-overflow-scrolling:touch", 'flex-direction:column', "fz-scroll"):
        assert needle in src, needle


def test_cash_allocation_grid_auto_and_save(web, world):
    """توزيع السند على عدة فواتير من واجهة الويب: توزيع تلقائي ثم حفظ سند واحد."""
    F, user = web
    sell(F, world["cust"], "ELH-T1", 5, 20, d="2026-01-12")
    cust = tuple(world["cust"])
    at = _run(user, "cash", cv_type="قبض", cv_cat=F.CAT_PARTNER, cv_party=cust, cv_amount=150.0, cv_spread=True)
    assert not at.exception, [e.message for e in at.exception]
    next(b for b in at.button if b.key == "cv_auto").click().run()
    assert not at.exception, [e.message for e in at.exception]
    plan = at.session_state["cv_alloc_plan"][1]
    assert round(sum(plan.values()), 2) == 150.0 and len(plan) == 2
    next(b for b in at.button if b.key == "cv_save").click().run()
    assert not at.exception and not [e.value for e in at.error], [e.value for e in at.error]
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM advance_allocations").fetchone()[0] == 2
        assert conn.execute("SELECT status FROM sales_invoices ORDER BY invoice_id LIMIT 1").fetchone()[0] == "مغلقة"


@pytest.mark.parametrize("key", ["pt_cust", "pt_supp", "pt_net"])
@pytest.mark.parametrize("mode", ["summary", "detail"])
def test_party_reports_render_in_web(web, key, mode):
    F, user = web
    group = next(g for g, items in F.REPORT_CATALOG if any(k == key for k, *_ in items))
    at = _run(user, "reports", rep_group=group, rep_key=key, **{f"pt_mode_{key}": mode})
    assert not at.exception, [e.message for e in at.exception]
    assert not [e.value for e in at.error], [e.value for e in at.error]
    word = {"summary": "إجماليات", "detail": "تفاصيل"}[mode]
    assert any(word in m.value for m in at.markdown if m.value.startswith("###")), mode

# -*- coding: utf-8 -*-
"""التخطيط المالي: توقع التدفق النقدي، الموازنة التقديرية مقابل الفعلي، وعمولات المناديب."""
from datetime import date, timedelta

import pytest

from conftest import buy, sell


def test_cash_flow_forecast_uses_due_dates(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    t = date.today()
    d_inv = (t - timedelta(days=5)).isoformat()
    buy(F, supp, "ELH-T1", 100, 10, d=(t - timedelta(days=40)).isoformat())   # 1,150 مستحق منذ 10 أيام (30 يوم افتراضي)
    F.save_partner_record("customer", cust[0], cust[1], credit=(0, 20, ""))
    sell(F, cust, "ELH-T1", 10, 50, d=d_inv)                                   # 575 تستحق بعد 15 يوماً
    rows, cash_now = F.cash_flow_forecast((t + timedelta(days=120)).isoformat(), "month")
    assert rows[0][0].startswith("متأخرات") and rows[0][2] == 1150.0
    due = (t - timedelta(days=5) + timedelta(days=20)).strftime("%Y-%m")
    assert any(r[0] == due and r[1] == 575.0 for r in rows)
    assert rows[-1][5] == pytest.approx(cash_now - 1150 + 575)
    det = F.run_report("due_inv", dict(d_to=t.isoformat()))
    assert {r[0] for r in det.rows} == {"تحصيل من عميل", "سداد لمورد"}
    weekly = F.run_report("cf_forecast", dict(d_to=(t + timedelta(days=60)).isoformat(), group="day"))
    assert weekly.rows and "أسبوعي" in weekly.subtitle


def test_forecast_excludes_paid_and_includes_budget(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    t = date.today()
    buy(F, supp, "ELH-T1", 10, 10, d=t.isoformat())
    si, _, _ = sell(F, cust, "ELH-T1", 2, 50, d=t.isoformat())
    F.save_voucher("قبض", t.isoformat(), "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 115, ("sales", si))
    nxt = (t.replace(day=1) + timedelta(days=40)).strftime("%Y-%m")
    F.save_budget(nxt, F.LINE_EXPENSES, 3000)
    rows, _ = F.cash_flow_forecast((t + timedelta(days=100)).isoformat(), "month")
    assert sum(r[1] for r in rows) == 0.0                     # الفاتورة مسددة بالكامل
    assert any(r[0] == nxt and r[3] == pytest.approx(3000) for r in rows)


def test_budget_vs_actual(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 100, 10, d="2026-03-01")
    sell(F, cust, "ELH-T1", 10, 100, d="2026-03-10")          # صافي مبيعات مارس 1,000
    F.save_budget_year(2026, F.LINE_SALES, [None, None, 800] + [None] * 9)
    with F.db() as conn:
        code, name = conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE account_type='مصروفات' "
                                  "AND COALESCE(is_group,0)=0 AND account_code LIKE '43%' LIMIT 1").fetchone()
    F.save_budget("2026-03", f"acc:{code}", 100)
    F.save_voucher("صرف", "2026-03-15", "صندوق نقدية الرياض", F.CAT_EXPENSE, (code, name), 150)
    rep = F.run_report("bva", dict(d_from="2026-03-01", d_to="2026-03-31"))
    by = {(r[0], r[1]): r for r in rep.rows}
    s = by[(F.BUILTIN_LINES[F.LINE_SALES], "2026-03")]
    assert s[2:5] == (800, 1000, 200) and s[6] == "ملائم ✓"
    e = next(r for r in rep.rows if r[0].startswith(code) and r[1] == "2026-03")
    assert e[3] == 150 and e[6] == "غير ملائم ✗"
    assert F.budget_year_matrix(2026)[F.LINE_SALES][2] == 800
    with pytest.raises(F.UserError):
        F.save_budget("2026-13", F.LINE_SALES, 1)
    with pytest.raises(F.UserError):
        F.save_budget("2026-03", "acc:21101", 1)               # ليس حساب إيراد/مصروف
    F.delete_budget("2026-03", F.LINE_SALES)
    assert F.budget_year_matrix(2026).get(F.LINE_SALES, [None] * 12)[2] is None


def test_rep_commissions(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    F.save_rep(None, dict(name="أحمد", commission_pct=2.5, active=1))
    F.save_partner_record("customer", cust[0], cust[1], rep="أحمد", rep_date="2026-01-01")
    buy(F, supp, "ELH-T1", 100, 10, d="2026-04-01")
    si, _, _ = sell(F, cust, "ELH-T1", 20, 50, d="2026-04-05")      # صافي 1,000
    F.save_voucher("قبض", "2026-04-06", "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 460, ("sales", si))
    p = dict(d_from="2026-04-01", d_to="2026-04-30")
    row = next(r for r in F.run_report("commissions", p).rows if r[1] == "أحمد")
    assert row[2] == 2.5 and row[3] == 1000 and row[6] == 25.0
    F.set_commission_base("collections")
    row = next(r for r in F.run_report("commissions", p).rows if r[1] == "أحمد")
    assert row[5] == 460 and row[6] == 11.5
    with pytest.raises(F.UserError):
        F.set_commission_base("x")

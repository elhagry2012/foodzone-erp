# -*- coding: utf-8 -*-
"""الحد الائتماني للعملاء: التنبيه، المنع، تجاوز المدير، وتقرير الاستخدام."""
import pytest

from conftest import buy, sell

D = "2026-02-01"


def _set(F, cust, limit, policy="", days=""):
    F.save_partner_record("customer", cust[0], cust[1], credit=(limit, days, policy))


def test_warn_policy_allows_but_reports_exceeded(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 100, 5)
    _set(F, cust, 1000, "warn")
    sell(F, cust, "ELH-T1", 50, 20, d=D)          # 1,150 > 1,000 لكن السياسة تنبيه فقط
    st = F.credit_status(cust[0], 0.0, "order")
    assert st["exceeded"] and st["balance"] == 1150.0
    assert "تجاوز الحد الائتماني" in st["message"]


def test_block_policy_rejects_order_and_delivery(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 100, 5)
    _set(F, cust, 1000, "block")
    with pytest.raises(F.CreditLimitError):
        F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 50, 20)])
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM sales_orders").fetchone()[0] == 0     # لم يُحفظ شيء
    oid = F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 50, 20)], allow_over_credit=True)
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM sales_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    with pytest.raises(F.CreditLimitError):
        F.fulfil_order("sales", oid, D, "", [(lid, 50, None)])
    with F.db() as conn:
        assert conn.execute("SELECT quantity FROM inventory WHERE item_code='ELH-T1'").fetchone()[0] == 100   # تراجع كامل
    did = F.fulfil_order("sales", oid, D, "", [(lid, 50, None)], allow_over_credit=True)
    iid, total = F.invoice_from_fulfilments("sales", cust[0], [did], D)    # الفوترة لا تُمنع أبداً
    assert total == 1150.0


def test_within_limit_and_unlimited(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 100, 5)
    _set(F, cust, 5000, "block")
    sell(F, cust, "ELH-T1", 10, 20, d=D)
    st = F.credit_status(cust[0], 0.0, "order")
    assert not st["exceeded"] and st["available"] == pytest.approx(5000 - 230)
    _set(F, cust, 0, "block")                          # صفر = بدون حد
    assert not F.credit_status(cust[0], 10**7, "order")["exceeded"]


def test_open_orders_and_collections_count(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 100, 5)
    _set(F, cust, 2000, "block")
    si, _, _ = sell(F, cust, "ELH-T1", 40, 20, d=D)                       # 920
    F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 40, 20)])   # أمر مفتوح 920
    with pytest.raises(F.CreditLimitError):                                  # 920 + 920 + 230 > 2000
        F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 10, 20)])
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 500, ("sales", si))
    F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 10, 20)])    # بعد التحصيل يُقبل


def test_credit_terms_validation_and_report(world):
    F, cust = world["F"], world["cust"]
    with pytest.raises(F.UserError):
        _set(F, cust, -1)
    with pytest.raises(F.UserError):
        _set(F, cust, 100, "bogus")
    _set(F, cust, 100, "", 45)
    with F.db() as conn:
        limit, days, policy, stored = F.partner_credit_terms(conn, cust[0])
    assert (limit, days, policy, stored) == (100, 45, F.default_policy(), "")
    F.set_credit_defaults("block", 60)
    assert F.default_policy() == "block" and F.default_credit_days() == 60
    rep = F.run_report("credit", dict(d_from=D, d_to=D))
    assert any(r[0] == cust[0] for r in rep.rows)

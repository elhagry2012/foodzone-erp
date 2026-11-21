# -*- coding: utf-8 -*-
"""كارت الصنف التفصيلي (HansaWorld + النظام)، تطابق لوحة المؤشرات مع قائمة الدخل، ومسح سجل التدقيق عند التصفير."""
import sqlite3

import pytest

from conftest import buy, sell
from test_hansa_alloc_reset import _write_hw_files, _write_tb


def test_item_card_shows_each_system_move_with_ref_and_party(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 10, 10)
    sell(F, cust, "ELH-T1", 2, 100)
    card = F.item_card("ELH-T1", "2000-01-01", "2099-12-31")
    assert [m["qty"] for m in card["moves"]] == [10, -2] and card["final"] == 8
    buy_m, sell_m = card["moves"]
    assert buy_m["party_code"] == supp[0] and buy_m["ref"] != "-" and sell_m["party_code"] == cust[0]
    assert buy_m["key"][0] in ("purchase", "receipt") and sell_m["key"][0] in ("sales", "delivery")
    assert sell_m["qty_out"] == 2 and sell_m["balance"] == 8 and sell_m["cost"] == 10 and sell_m["value"] == 20
    # فترة لاحقة: الرصيد المرحّل يشمل ما قبلها
    later = F.item_card("ELH-T1", "2099-01-01", "2099-12-31")
    assert later["opening"] == 8 and later["moves"] == [] and later["opening_label"] == "رصيد مرحّل"
    rows, final = F.fetch_item_card("ELH-T1", "2000-01-01", "2099-12-31")
    assert len(rows) == 3 and final == 8


def test_item_card_includes_hansaworld_history_without_double_counting(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    H.run_hansa_import(**_write_hw_files(tmp_path))
    card = F.item_card("100001", "1900-01-01", "2099-12-31")
    assert card["has_hw"] and card["opening"] == 0
    assert sorted(m["qty"] for m in card["moves"]) == [-30, -5, 100]
    assert card["final"] == 65                                    # = الرصيد الافتتاحي من ملف التكلفة (لا تكرار)
    by_ref = {m["ref"]: m for m in card["moves"]}
    rec, deld, dup = by_ref["25000010"], by_ref["25000020"], by_ref["25000021"]
    assert rec["src"] == "HansaWorld" and rec["type"] == "استلام من مورد" and rec["ref"] == "25000010"
    assert rec["party_code"] == "S0001" and rec["link"] == "PINV 25000010" and rec["key"][0] == "purchase"
    assert deld["party_code"] == "C0001" and deld["key"][0] in ("sales", "partner") and deld["qty_out"] == 30
    assert deld["cost"] == 20.5 and deld["value"] == round(30 * 20.5, 2)
    assert "مكرر" in dup["notes"]
    first = F.item_first_move_date("100001")
    assert first == min(m["date"] for m in card["moves"])
    # فلتر يبدأ بعد أول حركة: ما قبله يصبح رصيداً مرحّلاً والرصيد النهائي لا يتغير
    part = F.item_card("100001", rec["date"], "2099-12-31")
    before = sum(m["qty"] for m in card["moves"] if m["date"] < rec["date"])
    assert part["opening"] == before and part["final"] == 65 and part["opening_label"] == "رصيد مرحّل"


def test_dashboard_matches_income_figures(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    H.run_hansa_import(**_write_hw_files(tmp_path), balance=_write_tb(tmp_path), tb_as_of="2025-09-30")
    H.post_tb_expense_aggregate()
    m = F.fetch_dashboard_metrics()
    f = F.income_figures(F.DASH_FROM, F.today())
    for k_dash, k_inc in (("sales", "net_sales"), ("cogs", "cogs"), ("gross", "gross"), ("other_income", "total_other"),
                          ("expenses", "opex"), ("depreciation", "dep"), ("profit", "net_profit")):
        assert round(m[k_dash], 2) == round(f[k_inc], 2), k_dash
    assert round(m["gross"] + m["other_income"] - m["expenses"] - m["depreciation"], 2) == round(m["profit"], 2)
    st = H.tb_expense_status()
    assert st["posted"] and st["posted"][2] == round(m["expenses"], 2) and st["total"] == 882.5


def test_reset_can_clear_audit_log(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    from elhagry_core import audit as A
    buy(F, supp, "ELH-T1", 10, 10)
    sell(F, cust, "ELH-T1", 2, 100)
    res = F.reset_transactions("2026")                      # بدون مسح: السجل يحتفظ بحذف المستندات وحدث التصفير
    assert res["audit_cleared"] is None
    with F.db() as conn:
        kept = conn.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]
    assert kept > 1
    buy(F, supp, "ELH-T1", 5, 10)
    res = F.reset_transactions("2026", clear_audit=True)
    assert res["audit_cleared"] >= kept and res["backup"]
    with F.db() as conn:
        rows = conn.execute("SELECT log_id, doc_type, new_values FROM audit_log").fetchall()
        assert len(rows) == 1 and rows[0][0] == 1 and "مسح سجل التدقيق" in rows[0][1]
        with pytest.raises(sqlite3.DatabaseError):          # منع الحذف ما زال مفعلاً
            conn.execute("DELETE FROM audit_log")
    ok, n, bad = A.verify_audit_chain()
    assert ok and n == 1 and bad is None
    # النسخة الاحتياطية قبل المسح تحتفظ بالسجل القديم
    old = sqlite3.connect(res["backup"])
    try:
        assert old.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0] >= kept
    finally:
        old.close()

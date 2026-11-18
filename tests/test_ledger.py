# -*- coding: utf-8 -*-
"""الترحيل الآلي إلى دفتر الأستاذ العام: التوازن، المطابقة مع الدفاتر المساعدة، والتحديث الفوري مع أي تعديل."""
import pytest

from conftest import buy, sell

D = "2026-01-31"


def _gl(F, src=None):
    with F.db() as conn:
        sql = ("SELECT e.source_type, l.account_code, SUM(l.debit), SUM(l.credit) FROM gl_lines l JOIN gl_entries e "
               "ON e.gl_id=l.gl_id " + ("WHERE e.source_type=? " if src else "") + "GROUP BY 1, 2")
        return conn.execute(sql, (src,) if src else ()).fetchall()


def _acc(F, key):
    with F.db() as conn:
        return F.gl_accounts(conn)[key]


def _bal(F, key, as_of=None):
    info, data = F.gl_balances("0000-01-01", as_of or F.today())
    o, dr, cr = data.get(_acc(F, key), [0, 0, 0])
    return round(o + dr - cr, 2)


def _assert_reconciled(F, as_of=None):
    for name, gl, sub, diff in F.ledger_reconciliation(as_of or F.today()):
        assert abs(diff) < 0.02, f"{name}: GL={gl} sub={sub}"
    st = F.ledger_status()
    assert st["balanced"] and st["pending"] == 0 and not st["last_error"]
    bs = F.balance_sheet(as_of or F.today())
    assert bs["balanced"], bs["totals"]


def test_full_cycle_posts_and_reconciles(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    pi = buy(F, supp, "ELH-T1", 100, 10)
    si, did, _ = sell(F, cust, "ELH-T1", 30, 20, discount=15)
    # فاتورة المبيعات: عملاء = الإجمالي، إيراد = الصافي بعد الخصم، ضريبة 15%
    net = 30 * 20 - 15
    assert _bal(F, "gl_acc_ar") == round(net * 1.15, 2)
    assert _bal(F, "gl_acc_sales") == -net
    assert _bal(F, "gl_acc_vat_out") == -round(net * 0.15, 2)
    assert _bal(F, "gl_acc_cogs") == 300.0            # 30 × 10
    assert _bal(F, "gl_acc_stock") == 700.0           # 1000 - 300
    assert _bal(F, "gl_acc_ap") == -1150.0
    assert _bal(F, "gl_acc_vat_in") == 150.0
    assert _bal(F, "gl_acc_grni") == 0.0 and _bal(F, "gl_acc_gdni") == 0.0
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 300, ("sales", si))
    F.save_voucher("صرف", D, "بنك الراجحى", F.CAT_PARTNER, supp, 500, ("purchase", pi))
    _assert_reconciled(F)


def test_unbilled_delivery_and_receipt_use_clearing_accounts(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 8)
    oid = F.create_order("purchase", D, supp, "", [F.make_line("ELH-T2", "x", "KG", 10, 4)])
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM purchase_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    F.fulfil_order("purchase", oid, D, "", [(lid, 10, None)])            # استلام بدون فاتورة
    sell(F, cust, "ELH-T1", 5, 20, invoice=False)                      # تسليم بدون فاتورة
    assert _bal(F, "gl_acc_grni") == -40.0
    assert _bal(F, "gl_acc_gdni") == 40.0                               # 5 × 8 تكلفة
    assert _bal(F, "gl_acc_cogs") == 0.0                                # التكلفة تُعترف بها عند الفوترة (مطابقة قائمة الدخل)
    _assert_reconciled(F)


def test_returns_and_purchase_price_difference(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 10, 10)
    sell(F, cust, "ELH-T1", 4, 30)
    F.save_return(F.SALES_RETURN, D, cust, ("ELH-T1", "صنف أول"), 1, 30)
    F.save_return(F.PURCHASE_RETURN, D, supp, ("ELH-T1", "صنف أول"), 2, 25)   # تكلفته 10 أُرجع بـ 25
    assert _bal(F, "gl_acc_sales_ret") == 30.0
    assert _bal(F, "gl_acc_pret_diff") == -30.0        # (25 - 10) × 2 ربح فروق مردودات
    _assert_reconciled(F)


def test_expense_voucher_posts_to_chosen_account(world):
    F = world["F"]
    with F.db() as conn:
        code, name = conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE account_type='مصروفات' "
                                  "AND COALESCE(is_group,0)=0 AND account_code LIKE '43%' LIMIT 1").fetchone()
    F.save_voucher("صرف", D, "صندوق نقدية الرياض", F.CAT_EXPENSE, (code, name), 120)
    info, data = F.gl_balances("0000-01-01", F.today())
    assert round(data[code][1], 2) == 120.0
    _assert_reconciled(F)


def test_unknown_cash_account_goes_to_suspense(world):
    F, cust = world["F"], world["cust"]
    with F.db() as conn:
        conn.execute("INSERT INTO cash_transactions (trans_type, trans_date, account_bank, category, party_code, party_name, "
                     "amount) VALUES ('قبض', ?, 'خزينة غير معرّفة', ?, ?, ?, 50)", (D, F.CAT_PARTNER, cust[0], cust[1]))
    assert _bal(F, "gl_acc_suspense") == 50.0
    names = [r[0] for r in F.integrity_check() if r[0] == "تنبيه" and "المعلق" in r[2]]
    assert names, "integrity check should warn about the suspense balance"


def test_edit_and_delete_invoice_repost_immediately(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 20, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 10, 20)
    head, lines = F.load_invoice("sales", si)
    lines = [dict(l, price=25.0) for l in lines]
    F.update_invoice("sales", si, dict(date=D, party=cust, notes="معدلة"), lines)
    assert _bal(F, "gl_acc_sales") == -250.0
    _assert_reconciled(F)
    F.delete_invoice("sales", si)
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM gl_entries WHERE source_type='sales_invoice'").fetchone()[0] == 0
    assert _bal(F, "gl_acc_ar") == 0.0 and _bal(F, "gl_acc_gdni") == 100.0   # الإذن عاد «غير مفوتر»
    _assert_reconciled(F)


def test_post_discount_and_openings_and_depreciation(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    pi = buy(F, supp, "ELH-T1", 10, 10)
    F.save_post_discount("purchase", D, supp, 20, invoice_id=pi, method=F.DISC_STOCK, item=None)
    F.save_post_discount("sales", D, cust, 10)
    F.save_partner_record("customer", cust[0], cust[1], opening=(500, 0, "2026-01-01"))
    with F.db() as conn:
        conn.execute("INSERT INTO fixed_assets (asset_code, asset_name, purchase_date, purchase_cost, salvage_value, "
                     "useful_life_years, monthly_depreciation) VALUES ('AST-T','سيارة','2026-01-15',12000,0,1,1000)")
    F.sync_ledger()
    with F.db() as conn:
        n_dep = conn.execute("SELECT COUNT(*) FROM gl_entries WHERE source_type='asset'").fetchone()[0]
    import datetime as _dt
    t = _dt.date.today()
    expected = min(max((t.year - 2026) * 12 + t.month - 1, 0), 12)
    assert n_dep == expected
    assert _bal(F, "gl_acc_dep_acc") == -1000.0 * expected
    _assert_reconciled(F)


def test_incremental_equals_full_rebuild(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 40, 9)
    buy(F, supp, "ELH-T2", 15, 3, d="2026-02-01")
    sell(F, cust, "ELH-T1", 7, 20)
    sell(F, cust, "ELH-T2", 5, 9, d="2026-02-03")
    F.save_return(F.SALES_RETURN, D, cust, ("ELH-T1", "صنف أول"), 2, 20)
    before = sorted(_gl(F))
    F.rebuild_ledger()
    assert sorted(_gl(F)) == before


def test_trial_balance_report_uses_full_ledger(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 10, 10)
    sell(F, cust, "ELH-T1", 2, 30)
    rep = F.run_report("tb", dict(d_from="2026-01-01", d_to=F.today()))
    assert "متوازن ✓" in rep.notes[-1]
    tot = rep.total
    assert abs(tot[5] - tot[6]) < 0.01 and tot[5] > 0
    for key in ("gl_tb", "bs", "gl_pl", "gl_journal", "gl_recon"):
        r = F.run_report(key, dict(d_from="2026-01-01", d_to=F.today()))
        assert r.rows, key


def test_posting_failure_never_blocks_saving(world, monkeypatch):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    from elhagry_core import ledger

    def boom(*a, **k):
        raise RuntimeError("posting bug")

    monkeypatch.setitem(ledger._POSTERS, "purchase_invoice", boom)
    pi = buy(F, supp, "ELH-T1", 5, 10)
    with F.db() as conn:
        assert conn.execute("SELECT 1 FROM purchase_invoices WHERE purchase_id=?", (pi,)).fetchone()   # المستند حُفظ
        assert conn.execute("SELECT COUNT(*) FROM gl_queue").fetchone()[0] >= 1                         # وينتظر الترحيل
    assert "posting bug" in F.ledger_status()["last_error"]
    monkeypatch.undo()
    F.sync_ledger()
    assert F.ledger_status()["pending"] == 0
    _assert_reconciled(F)


def test_gl_account_mapping_validation(world):
    F = world["F"]
    with pytest.raises(F.UserError):
        F.set_gl_accounts({"gl_acc_sales": "31100"})     # حساب تجميعي
    with pytest.raises(F.UserError):
        F.set_gl_accounts({"gl_acc_sales": "00000"})     # غير موجود
    F.set_gl_accounts({"gl_acc_sales": "31120"})
    assert _acc(F, "gl_acc_sales") == "31120"


def test_text_expense_voucher_maps_to_expense_account(world):
    """سند مصروفات ببند نصي بدون كود (مثل «مواصلات ونقل»): يُرحَّل لحساب مصروف وليس للحساب المعلق،
    فيطابق صافي ربح الدفتر العام قائمة الدخل (حالة حقيقية من بيانات الشركة)."""
    F = world["F"]
    F.save_voucher("صرف", D, "صندوق نقدية الرياض", F.CAT_EXPENSE, (None, "مواصلات ونقل"), 200)
    assert _bal(F, "gl_acc_suspense") == 0.0
    assert _bal(F, "gl_acc_exp_general") == 200.0
    _assert_reconciled(F)

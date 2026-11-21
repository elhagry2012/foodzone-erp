# -*- coding: utf-8 -*-
"""تصفير الحركات بكلمة المرور، سند واحد موزع على عدة فواتير، ومستورد HansaWorld (ملفات صغيرة بنفس الهيكل الفعلي)."""
import datetime as dt

import pandas as pd
import pytest

from conftest import buy, sell


# ------------------------------------------------------------------ تصفير الحركات
def test_reset_requires_password_and_clears_allocations(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 10, 10)
    iid, _, _ = sell(F, cust, "ELH-T1", 2, 100)
    vid = F.save_voucher("قبض", "2026-01-11", "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 50)
    F.allocate_advance(vid, "sales", iid, 50)
    with pytest.raises(F.UserError, match="غير صحيحة"):
        F.reset_transactions("0000")
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM sales_invoices").fetchone()[0] == 1   # لم يُحذف شيء
    res = F.reset_transactions("2026")
    assert res["backup"] and res["deleted"]["advance_allocations"] == 1
    with F.db() as conn:
        for t in ("sales_invoices", "cash_transactions", "advance_allocations", "purchase_invoices"):
            assert conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] == 0
        assert conn.execute("SELECT quantity FROM inventory WHERE item_code='ELH-T1'").fetchone()[0] == 0
    # كلمة المرور قابلة للتغيير
    F.set_reset_password("9999")
    with pytest.raises(F.UserError):
        F.reset_transactions("2026")
    F.reset_transactions("9999")


def test_reset_admin_only(world):
    F = world["F"]
    F.set_current_user("clerk", "test", user=dict(is_admin=0, perms=set()))
    try:
        with pytest.raises(F.UserError, match="لمدير النظام"):
            F.reset_transactions("2026")
    finally:
        F.set_current_user("tester", "test")


# ------------------------------------------------------------------ توزيع السند على الفواتير
def test_fifo_allocation_and_allocated_voucher(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    a, _, _ = sell(F, cust, "ELH-T1", 10, 100, d="2026-01-10")   # 1150
    b, _, _ = sell(F, cust, "ELH-T1", 10, 100, d="2026-01-12")   # 1150
    c, _, _ = sell(F, cust, "ELH-T1", 10, 100, d="2026-01-15")   # 1150
    inv = F.party_open_invoices("sales", cust[0])
    assert [r["inv_id"] for r in inv] == [a, b, c]
    plan = F.fifo_allocation(inv, 1500)
    assert plan == {a: 1150.0, b: 350.0}
    with pytest.raises(F.UserError, match="أكبر من مبلغ السند"):
        F.save_allocated_voucher("قبض", "2026-01-20", "بنك الراجحى", cust, 1000, plan)
    with pytest.raises(F.UserError, match="أكبر من متبقيها"):
        F.save_allocated_voucher("قبض", "2026-01-20", "بنك الراجحى", cust, 5000, {a: 1200})
    vid, used = F.save_allocated_voucher("قبض", "2026-01-20", "بنك الراجحى", cust, 1600, plan)
    assert used == 1500
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM cash_transactions WHERE trans_id=?", (vid,)).fetchone()[0] == 1
        st = dict(conn.execute("SELECT invoice_id, status FROM sales_invoices").fetchall())
        paid = dict(conn.execute("SELECT invoice_id, paid_amount FROM sales_invoices").fetchall())
    assert st[a] == "مغلقة" and st[b] == "مفتوحة" and paid[b] == 350
    assert F.party_advance_available(cust[0], "sales") == 100     # الباقي تحت الحساب
    assert sorted(x[3] for x in F.voucher_allocations(vid)) == [350, 1150]
    # الرصيد والدفتر العام: سند واحد على البنك بقيمة 1600
    bal = {r["code"]: r["balance"] for r in F.partner_balances("9999-12-31")}
    assert round(bal[cust[0]], 2) == round(3 * 1150 - 1600, 2)
    F.delete_voucher(vid)
    assert all(r["paid"] == 0 for r in F.party_open_invoices("sales", cust[0]))


def test_allocated_voucher_rejects_other_party(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    other = (F.save_partner_record("customer", None, "عميل آخر"), "عميل آخر")
    buy(F, supp, "ELH-T1", 20, 10)
    a, _, _ = sell(F, cust, "ELH-T1", 5, 100)
    with pytest.raises(F.UserError, match="لا تخص"):
        F.save_allocated_voucher("قبض", "2026-01-20", "بنك الراجحى", other, 100, {a: 100})


# ------------------------------------------------------------------ قراءة HansaWorld
def test_detect_swapped_dates():
    from elhagry_core import hansa_import as H
    vals = ["16/08/2025", dt.datetime(2025, 8, 9), "21/09/2025", dt.datetime(2025, 4, 10)]
    assert H.detect_swapped_dates(vals)
    assert [H.parse_hw_date(v, True) for v in vals] == ["2025-08-16", "2025-09-08", "2025-09-21", "2025-10-04"]
    # ملف تواريخه كلها حقيقية (فيها يوم > 12) لا تُعكس
    assert not H.detect_swapped_dates([dt.datetime(2025, 9, 21), dt.datetime(2025, 10, 1)])
    assert H.parse_hw_date(45900.0) == "2025-08-31"   # رقم إكسيل تسلسلي


def test_simulate_allocations_unapply_and_cap():
    from elhagry_core import hansa_import as H
    i1, i2 = ("sales", "A"), ("sales", "B")
    vs = [dict(key="R1", party="p", date="2026-01-01", cash=100, doc="R1", lines=[(i1, 100)]),
          # R2 يلغي 40 من A ويضعها مع نقديته على B
          dict(key="R2", party="p", date="2026-01-02", cash=60, doc="R2", lines=[(i1, -40), (i2, 100)])]
    alloc, probs = H.simulate_allocations(vs, {i1: 100, i2: 100})
    assert alloc[("R1", i1)][0] == 60 and alloc[("R2", i2)][0] == 60 and alloc[("R1", i2)][0] == 40
    assert not probs
    # تسوية زائدة: السقف على الحالة النهائية والزائد يبقى تحت الحساب
    alloc, probs = H.simulate_allocations([dict(key="R", party="p", date="d", cash=150, doc="R", lines=[(i1, 150)])],
                                          {i1: 100})
    assert alloc[("R", i1)][0] == 100 and probs


def _write_hw_files(tmp_path):
    cust = pd.DataFrame([
        ["C0001", "شركة الأولى", dt.datetime(2025, 9, 1), "INV", 10000001, None, 1150, 0, "SAR", 1150],
        ["C0001", "شركة الأولى", dt.datetime(2025, 9, 2), "INV", 10000002, None, 2300, 0, "SAR", 3450],
        ["C0001", "شركة الأولى", dt.datetime(2025, 9, 3), "CRED", 10000003, 10000002, 0, 115, "SAR", 3335],
        ["C0001", "شركة الأولى", dt.datetime(2025, 9, 20), "REC", 25000001, 10000001, 0, 1150, "SAR", 2185],
        ["C0001", "شركة الأولى", dt.datetime(2025, 9, 20), "REC", 25000001, 10000002, 0, 1000, "SAR", 1185],
        ["C0002", "مؤسسة الثانية", dt.datetime(2025, 9, 5), "INV", 10000004, None, 500, 0, "SAR", 500],
    ], columns=["Customer Code", "Customer Name", "Trans Date", "Type", "Doc No", "Applied To Doc", "Debit (SAR)",
                "Credit (SAR)", "Currency", "Running Balance (SAR)"])
    supp = pd.DataFrame([
        ["C0001", "شركة الأولى", 310000000000003, "SAR", dt.datetime(2025, 9, 1), "PINV", "Purchase Invoice", 25000010,
         None, 0, 4600, 4600, 3],
        ["C0001", "شركة الأولى", 310000000000003, "SAR", dt.datetime(2025, 9, 9), "PAYM", "Payment", 25000011,
         25000010, 4000, 0, 600, 4],
        ["Net Change", None, None, None, None, None, None, None, None, 4000, 4600, 600, 7],
    ], columns=["Supplier Code", "Supplier Name", "VAT No", "Currency", "Trans Date", "Type", "Type Description", "Doc No",
                "Applied To Invoice", "Debit (SAR)", "Credit (SAR)", "Balance (SAR)", "Source Row"])
    hist = pd.DataFrame([
        ["100001", "لحم تجريبي", "KG", "16/08/2025", "G.Rec.", "25000010", "شركة الأولى", "MAIN", 100, 100],
        ["100001", "لحم تجريبي", "KG", dt.datetime(2025, 9, 1), "Deld", "25000020", "شركة الأولى", "MAIN", -30, 70],
        ["100001", "لحم تجريبي", "KG", dt.datetime(2025, 9, 1), "Deld", "25000021", "مكرر", "MAIN", -5, 65],
        ["100001", "لحم تجريبي", "KG", None, "End Balance", None, None, None, 0, 65],
        ["100002", "دجاج تجريبي", "KG", "20/09/2025", "G.Rec.", "25000012", "شركة الأولى", "MAIN", 10, 10],
        ["100002", "دجاج تجريبي", "KG", "21/09/2025", "Deld", "25000022", "مؤسسة الثانية", "MAIN", -10, 0],
        ["100002", "دجاج تجريبي", "KG", None, "End Balance", None, None, None, 0, 0],
        [None, None, None, None, None, None, None, None, None, 65],
    ], columns=["Item", "Name", "Unit", "Date", "Type", "Reference", "Cust/Supp.", "Location", "Qty", "Balance"])
    cost = pd.DataFrame([["100001", "لحم تجريبي", 65, 20.5, 21, -2.4, -0.5]],
                        columns=["Item", "Name", "In Stock", "Cost Pr.", "Last Price", "%", "Diff."])
    paths = {}
    for key, df, sheet in (("customers", cust, "ورقة1"), ("suppliers", supp, "Supplier Transactions"),
                           ("history", hist, "Item History"), ("costs", cost, "Item Cost Price Analysis")):
        paths[key] = str(tmp_path / f"{key}.xlsx")
        with pd.ExcelWriter(paths[key]) as xw:
            df.to_excel(xw, sheet_name=sheet, index=False)
    return paths


def test_hansa_import_end_to_end(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    # شريك موجود مسبقاً بنفس الاسم لكن بكود مختلف (ترقيم قديم) يُعاد ترقيمه لكود HansaWorld
    old = F.save_partner_record("customer", None, "مؤسسة الثانية", phone="0500000000")
    paths = _write_hw_files(tmp_path)
    dry = H.run_hansa_import(**paths, dry_run=True)
    assert dry["dry_run"] and not [r for r in dry["report"].rows if str(r[5]).startswith("❌")], dry["report"].rows
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM sales_invoices").fetchone()[0] == 0     # المعاينة لا تحفظ
    out = H.run_hansa_import(**paths)
    bad = [r for r in out["report"].rows if str(r[5]).startswith("❌")]
    assert not bad, bad
    assert "حركات الأصناف" in out["swapped"]
    with F.db() as conn:
        p = {r[0]: r[1:] for r in conn.execute("SELECT partner_code, name, is_customer, is_supplier, legacy_code, "
                                               "vat_number, phone FROM business_partners")}
        assert p["C0001"][1:4] == (1, 0, "C0001") and p["S0001"][1:5] == (0, 1, "C0001", "310000000000003")
        assert p["C0002"][0] == "مؤسسة الثانية" and p["C0002"][5] == "0500000000"   # احتفظ بالهاتف
        assert old == "C0001" and p["C0001"][0] == "شركة الأولى"   # الكود القديم أُعطي لصاحبه في HansaWorld
        st = dict(conn.execute("SELECT invoice_no, status FROM sales_invoices").fetchall())
        assert st == {"10000001": "مغلقة", "10000002": "مفتوحة", "10000004": "مفتوحة"}
        n_vouchers = conn.execute("SELECT COUNT(*) FROM cash_transactions").fetchone()[0]
        assert n_vouchers == 2                                  # REC واحد (سطران) + PAYM واحد
        assert conn.execute("SELECT opening_qty, opening_cost, last_purchase_price FROM inventory WHERE item_code='100001'"
                            ).fetchone() == (65, 20.5, 21)
        h = conn.execute("SELECT move_date, party_text, is_duplicate, link_doc_no FROM hw_item_history ORDER BY hist_id"
                         ).fetchall()
        assert h[1][0] == "2025-01-09"                          # خلية إكسيل مقلوبة (شهر/يوم) صُححت
        assert h[2][1:3] == (None, 1)
        assert h[0][3] == "25000010"                            # استلام مربوط بفاتورة المورد بالمرجع والطرف
    rem = {r["no"]: r["remaining"] for r in F.party_open_invoices("sales", "C0001")}
    assert rem == {"10000002": 1185.0}
    # إعادة الاستيراد لا تكرر شيئاً
    again = H.run_hansa_import(**paths)
    assert again["stats"].get("sales_invoices", 0) == 0 and again["stats"]["skipped_existing"] > 0
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM cash_transactions").fetchone()[0] == n_vouchers
        assert conn.execute("SELECT COUNT(*) FROM hw_item_history").fetchone()[0] == 5


def _write_tb(tmp_path):
    rows = [["Account", None, None, "B/Fwd Bal", "Per Deb", "Per Cred", "Acc Deb", "Acc Cred", "Balance", None],
            ["11111", "صندوق نقدية الرياض", None, None, 1000, 400, 1000, 400, 600, None],
            ["11121", "بنك الراجحى", None, None, 100, 300, 100, 300, None, 200],
            ["11250", "عملاء جملة", None, None, 3950, 2265, 3950, 2265, 1685, None],
            ["11310", "مخزون سلعى مجمدات", None, None, 1332.5, 0, 1332.5, 0, 1332.5, None],
            ["Group '1'", None, None, None, 6382.5, 2965, 6382.5, 2965, 3417.5, None],
            ["22101", "موردين محليين", None, None, 4000, 4600, 4000, 4600, None, 600],
            ["Group '2'", None, None, None, 4000, 4600, 4000, 4600, None, 600],
            ["43101", "رواتب و اجور - بيعية", None, None, 500, None, 500, None, 500, None],
            ["41125", "ايجار سيارات", None, None, 300, None, 300, None, 300, None],
            ["55016", "وقود ومحروقات", None, None, 82.5, None, 82.5, None, 82.5, None],
            ["31110", "ايراد مبيعات مجمدات", None, None, 0, 3700, 0, 3700, None, 3700],
            [None, None, None, None, 11265, 11265, 11265, 11265, 5300, 5300]]
    path = str(tmp_path / "tb.xlsx")
    pd.DataFrame(rows[1:], columns=rows[0]).to_excel(path, sheet_name="Analytical Balance Report", index=False)
    return path


def test_trial_balance_reference_and_expenses(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    paths = _write_hw_files(tmp_path)
    tb = H.read_balance_file(_write_tb(tmp_path))
    assert tb["accounts"]["11121"]["balance"] == -200 and tb["accounts"]["22101"]["balance"] == -600
    assert not tb["problems"], tb["problems"]                         # متوازن ومجموعاته صحيحة
    out = H.run_hansa_import(**paths, balance=_write_tb(tmp_path), tb_as_of="2025-09-30", apply_cash_openings=True)
    rows = {r[1]: r for r in out["report"].rows if r[0] == "ميزان المراجعة"}
    assert rows["الموردون 22101 (دائن) ↔ أرصدة الموردين بالنظام"][5].startswith("✅")
    assert next(r for k, r in rows.items() if k.startswith("العملاء 11250"))[5].startswith("✅")
    assert rows["الخزينة/البنك 11121 بنك الراجحى"][3] == -200                  # سحب على المكشوف
    assert rows["الخزينة/البنك 11111 صندوق نقدية الرياض"][5].startswith("✅")
    with F.db() as conn:
        assert conn.execute("SELECT parent_code FROM chart_of_accounts WHERE account_code='41125'").fetchone()[0] == "41100"
        assert conn.execute("SELECT parent_code FROM chart_of_accounts WHERE account_code='55016'").fetchone()[0] == "55000"
        assert conn.execute("SELECT COUNT(*) FROM hw_tb_reference").fetchone()[0] == 9
    rep = H.rep_hw_expenses({})
    assert {r[1] for r in rep.rows} == {"43101", "41125", "55016"} and rep.total[3] == 882.5
    # حركات مصاريف تفصيلية لاحقاً تطابق الرصيد المرجعي ولا تتكرر
    exp = str(tmp_path / "exp.xlsx")
    pd.DataFrame([["2025-09-10", "43101", 500, "", "رواتب", "E1"], ["2025-09-11", "41125", 300, "", "سيارات", "E2"]],
                 columns=H.EXP_TEMPLATE).to_excel(exp, index=False)
    dry = H.import_expense_movements(exp, dry_run=True)
    assert dry["stats"]["expenses"] == 2
    res = H.import_expense_movements(exp)
    ok = {r[1].split()[0]: r[5] for r in res["report"].rows}
    assert ok["43101"].startswith("✅") and ok["41125"].startswith("✅") and ok["55016"].startswith("❌")
    assert H.import_expense_movements(exp)["stats"].get("expenses", 0) == 0


def _cust_file(tmp_path, rows, name="c.xlsx"):
    cols = ["Customer Code", "Customer Name", "Trans Date", "Type", "Doc No", "Applied To Doc", "Debit (SAR)",
            "Credit (SAR)", "Currency", "Running Balance (SAR)"]
    path = str(tmp_path / name)
    with pd.ExcelWriter(path) as xw:
        pd.DataFrame(rows, columns=cols).to_excel(xw, sheet_name="ورقة1", index=False)
    return path


def test_credit_note_consumed_by_receipt_is_split(erp, tmp_path):
    """REC يستهلك إشعاراً دائناً تحت الحساب لسداد فاتورتين ← إشعارات فرعية /1 و /2 مربوطة بالفواتير، والرصيد ثابت."""
    F = erp
    from elhagry_core import hansa_import as H
    d = dt.datetime
    n = "شركة الإشعار"
    rows = [["C0009", n, d(2025, 9, 1), "CRED", 10000100, None, 0, 460, "SAR", -460],        # إشعار تحت الحساب
            ["C0009", n, d(2025, 9, 2), "INV", 10000101, None, 345, 0, "SAR", -115],
            ["C0009", n, d(2025, 9, 3), "INV", 10000102, None, 230, 0, "SAR", 115],
            ["C0009", n, d(2025, 9, 4), "INV", 10000103, None, 115, 0, "SAR", 230],
            ["C0009", n, d(2025, 9, 9), "REC", 25000100, 10000100, 460, 0, "SAR", 690],        # استهلاك الإشعار
            ["C0009", n, d(2025, 9, 9), "REC", 25000100, 10000101, 0, 345, "SAR", 345],
            ["C0009", n, d(2025, 9, 9), "REC", 25000100, "10000102/1", 0, 230, "SAR", 115],     # مرجع بلاحقة «/»
            ["C0009", n, d(2025, 9, 9), "REC", 25000100, 10000103, 0, 115, "SAR", 0]]
    path = _cust_file(tmp_path, rows)
    for _ in range(2):   # التشغيل الثاني يعيد نفس التقسيم بلا تكرار
        out = H.run_hansa_import(customers=path)
        assert not [r for r in out["report"].rows if str(r[5]).startswith("❌")], out["report"].rows
        with F.db() as conn:
            notes = conn.execute("SELECT discount_no, total_amount, vat_amount, invoice_id IS NOT NULL FROM post_discounts "
                                 "ORDER BY discount_no").fetchall()
            st = {r[0]: r[1] for r in conn.execute("SELECT invoice_no, status FROM sales_invoices")}
            vouchers = conn.execute("SELECT amount FROM cash_transactions").fetchall()
        assert vouchers == [(230.0,)]                                     # النقدية الفعلية فقط
        assert [x[0] for x in notes] == ["10000100/1", "10000100/2"]      # استُهلك كله: الأصل أصبح /1
        assert round(sum(x[1] for x in notes), 2) == 460 and round(sum(x[2] for x in notes), 2) == 60
        assert all(x[3] for x in notes) and set(st.values()) == {"مغلقة"}
    bal = {b["code"]: b["balance"] for b in F.partner_balances("9999-12-31")}
    assert abs(bal["C0009"]) < 0.005


def test_credit_note_partial_use_keeps_remainder_on_account(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    d = dt.datetime
    n = "عميل الباقي"
    rows = [["C0010", n, d(2025, 9, 1), "CRED", 10000200, None, 0, 1000, "SAR", -1000],
            ["C0010", n, d(2025, 9, 2), "INV", 10000201, None, 400, 0, "SAR", -600],
            ["C0010", n, d(2025, 9, 9), "REC", 25000200, 10000200, 400, 0, "SAR", -200],
            ["C0010", n, d(2025, 9, 9), "REC", 25000200, 10000201, 0, 400, "SAR", -600]]
    H.run_hansa_import(customers=_cust_file(tmp_path, rows))
    with F.db() as conn:
        notes = dict(conn.execute("SELECT discount_no, total_amount FROM post_discounts").fetchall())
        assert conn.execute("SELECT COUNT(*) FROM cash_transactions").fetchone()[0] == 0   # لا نقدية في السند
    assert notes == {"10000200": 600.0, "10000200/1": 400.0}
    assert H.find_doc({"10000201": 1}, "10000201/3") == ("10000201", 1)


def test_tb_expense_aggregate_idempotent(erp, tmp_path):
    F = erp
    from elhagry_core import hansa_import as H
    H.run_hansa_import(**_write_hw_files(tmp_path), balance=_write_tb(tmp_path), tb_as_of="2025-09-30")

    def lines():
        with F.db() as conn:
            return dict(conn.execute("SELECT l.account_code, l.debit - l.credit FROM journal_lines l JOIN journal_entries e "
                                     "ON e.entry_id=l.entry_id WHERE e.source_type='hw_tb_expenses'").fetchall())
    dry = H.post_tb_expense_aggregate(dry_run=True)
    assert dry["total"] == 882.5 and lines() == {}
    first = H.post_tb_expense_aggregate()
    assert H.post_tb_expense_aggregate()["entry_id"] == first["entry_id"]           # إعادة الترحيل لا تكرر
    assert lines() == {"43101": 500, "41125": 300, "55016": 82.5, "99999": -882.5}
    with pytest.raises(F.UserError, match="HansaWorld"):
        F.delete_journal_entry(first["entry_id"])                                       # قيد نظام محمي
    exp = str(tmp_path / "e.xlsx")
    pd.DataFrame([["2025-09-10", "43101", 200, "", "رواتب", "E1"]], columns=H.EXP_TEMPLATE).to_excel(exp, index=False)
    H.import_expense_movements(exp)
    assert lines()["43101"] == 300                                                      # الفرق فقط بعد التفصيلي
    assert all(abs(r[5]) < 0.01 for r in H.rep_hw_expenses({}).rows)                 # كل حساب = رصيده المرجعي
    assert H.post_tb_expense_aggregate(remove=True)["removed"] and lines() == {}


def test_dashboard_profit_uses_periodic_cogs_for_hansaworld(erp, tmp_path):
    """فواتير HansaWorld بلا بنود: تكلفة المبيعات = المشتريات − إشعارات الموردين − المخزون الختامي، والخصم المكتسب
    المستورد لا يُضاف مرة ثانية كإيراد، ومصاريف الميزان المرحّلة تظهر في المصروفات."""
    F = erp
    from elhagry_core import hansa_import as H
    H.run_hansa_import(**_write_hw_files(tmp_path), balance=_write_tb(tmp_path), tb_as_of="2025-09-30")
    H.post_tb_expense_aggregate()
    m = F.fetch_dashboard_metrics()
    net_sales = round((1150 + 2300 + 500 - 115) / 1.15, 2)          # فواتير − إشعار العميل (قبل الضريبة)
    cogs = round(4600 / 1.15 - 65 * 20.5, 2)                         # مشتريات − مخزون ختامي (لا إشعارات موردين)
    assert round(m["sales"], 2) == net_sales
    assert round(m["gross"], 2) == round(net_sales - cogs, 2)
    assert round(m["expenses"], 2) == 882.5
    assert round(m["profit"], 2) == round(net_sales - cogs - 882.5, 2)
    txt = F.build_income_report("2025-01-01", "2025-12-31")
    assert "جرد دوري" in txt and f"{m['profit']:,.2f}" in txt

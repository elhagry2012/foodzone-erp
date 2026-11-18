# -*- coding: utf-8 -*-
"""المرحلة الرابعة: مرفقات فاتورة المورد على استلام المخزن، الأدوار والصلاحيات التفصيلية (عرض/إنشاء) المفروضة
في المحرك، تسوية الدفعات المقدمة مع الفواتير اللاحقة، وتقرير الفواتير المفتوحة."""
import pytest

from conftest import buy, sell
from test_phase3 import desk  # noqa: F401  (fixture واجهة سطح المكتب)

D = "2026-03-10"
PDF = b"%PDF-1.4 supplier invoice"


def _receipt(F, supp, qty=5, cost=10, d=D):
    oid = F.create_order("purchase", d, supp, "", [F.make_line("ELH-T1", "x", "KG", qty, cost)])
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM purchase_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    return oid, F.fulfil_order("purchase", oid, d, "", [(lid, qty, None)])


def _as(F, perms):
    """تشغيل العمليات كمستخدم غير مدير بصلاحيات محددة (فحص الصلاحية داخل المحرك)."""
    F.set_current_user("u1", "test", dict(is_admin=False, perms=set(perms)))


# ============================================================ 1) مرفقات استلام المخزن
def test_receipt_attachment_flows_to_invoice_and_supplier_view(world):
    F, supp = world["F"], world["supp"]
    _, rid = _receipt(F, supp)
    aid = F.add_receipt_attachment(rid, "supplier_inv_771.pdf", PDF)
    assert [a[0] for a in F.list_attachments("receipt", rid)] == [aid]
    inv, _ = F.invoice_from_fulfilments("purchase", supp[0], [rid], D)
    # تظهر تلقائياً على الفاتورة الصادرة من الإذن وفي عداد المرفقات
    assert aid in [a[0] for a in F.list_attachments("purchase", inv)]
    assert F.attachment_counts("purchase")[inv] == 1
    rows = F.supplier_attachments(supp[0])
    assert len(rows) == 1 and rows[0][1] == "إذن استلام" and rows[0][2] == rid
    assert F.get_attachment(aid)[2] == PDF


def test_receipt_attachment_requires_receipt_permission(world):
    F, supp = world["F"], world["supp"]
    _, rid = _receipt(F, supp)
    _as(F, ["purchase", "purchase.receipt"])                  # عرض الاستلام فقط
    with pytest.raises(F.UserError):
        F.add_receipt_attachment(rid, "x.pdf", PDF)
    _as(F, ["purchase", "purchase.receipt", "purchase.receipt_create"])
    assert F.add_receipt_attachment(rid, "x.pdf", PDF)


# ============================================================ 2) الصلاحيات والأدوار
def test_view_vs_create_enforced_in_engine(world):
    F, supp = world["F"], world["supp"]
    oid, _ = _receipt(F, supp, qty=10)
    _as(F, ["purchase", "purchase.orders"])                    # يرى أوامر الشراء ولا ينشئها
    with pytest.raises(F.UserError, match="أوامر الشراء: إنشاء"):
        F.create_order("purchase", D, supp, "", [F.make_line("ELH-T1", "x", "KG", 1, 10)])
    _as(F, ["purchase", "purchase.receipt"])                   # يرى الاستلام ولا ينفذه
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM purchase_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    with pytest.raises(F.UserError, match="استلام المخزن: تنفيذ"):
        F.fulfil_order("purchase", oid, D, "", [(lid, 1, None)])
    _as(F, ["purchase", "purchase.orders"])                    # بدون الخزينة: لا سندات
    with pytest.raises(F.UserError):
        F.save_voucher("صرف", D, "صندوق نقدية الرياض", F.CAT_PARTNER, supp, 10)


def test_create_implies_view_and_old_users_keep_create(world):
    F = world["F"]
    assert set(F.normalize_permissions(["purchase", "purchase.orders_create"])) == {
        "purchase", "purchase.orders", "purchase.orders_create"}
    with F.db() as conn:   # مستخدم قديم كان يملك «أوامر الشراء» قبل فصل العرض عن الإنشاء
        conn.execute("INSERT INTO users (username, password_hash, permissions) VALUES ('old', 'x', 'purchase,purchase.orders')")
        conn.execute("DELETE FROM app_settings WHERE key=?", (F.PERM_SPLIT_FLAG,))
        F.setup_controls(conn)
        perms = conn.execute("SELECT permissions FROM users WHERE username='old'").fetchone()[0].split(",")
    assert "purchase.orders_create" in perms


def test_roles_seeded_and_role_update_propagates(world):
    F = world["F"]
    roles = {n: (rid, perms) for rid, n, _d, perms, _c in F.list_roles()}
    assert {"أمين المخزن", "مسؤول المشتريات", "أمين الخزينة", "المحاسب"} <= set(roles)
    keeper_id, keeper = roles["أمين المخزن"]
    assert "purchase.receipt_create" in keeper and "purchase.orders_create" not in keeper and "cash" not in keeper
    F.save_user_account(None, "store1", "أمين", "Abc12345", "Abc12345", 0, 1, keeper, role_id=keeper_id)
    with F.db() as conn:
        uid, = conn.execute("SELECT user_id FROM users WHERE username='store1'").fetchone()
    assert F.user_role_id(uid) == keeper_id
    _, n = F.save_role(keeper_id, "أمين المخزن", keeper + ["expiry", "inventory"] + ["sales", "sales.delivery"])
    assert n == 1
    rid2, _ = F.save_role(None, "مراجع", ["reports", "reports.debts"])
    F.delete_role(rid2)
    assert "مراجع" not in [r[1] for r in F.list_roles()]


# ============================================================ 3) الدفعات المقدمة والفواتير المفتوحة
def test_advance_settles_later_invoice_without_new_entry(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 50, 10)
    adv = F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 500, None, "دفعة مقدمة")
    si, _, _ = sell(F, cust, "ELH-T1", 10, 100, d=D)                   # 1,150
    bal_before = F.fetch_partner_statement(cust[0], True)[4]
    assert F.party_advance_available(cust[0], "sales") == 500
    rows = F.open_invoices_rows("sales")
    assert rows[0][9] == "غير مسددة" and rows[0][11] == 500            # الدفعة المتاحة ظاهرة في التقرير
    done = F.settle_invoice_from_advances("sales", si)
    assert done == [(adv, 500.0)]
    assert F.invoice_payment_prefill("sales", si)["remaining"] == 650.0
    assert F.fetch_partner_statement(cust[0], True)[4] == bal_before      # مطابقة فقط: الرصيد لم يتغير
    row = F.open_invoices_rows("sales")[0]
    assert row[6] == 500 and row[8] == 650 and row[9] == "مسددة جزئياً" and row[11] == 0
    assert not [r for r in F.integrity_check() if r[0] == "خطأ"], [r for r in F.integrity_check() if r[0] == "خطأ"]
    # إكمال السداد بسند يغلق الفاتورة وتختفي من التقرير
    F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 650, ("sales", si))
    assert not F.open_invoices_rows("sales")
    # حذف سند الدفعة المقدمة يعيد الفاتورة مفتوحة بقيمتها
    F.delete_voucher(adv)
    assert F.invoice_payment_prefill("sales", si)["remaining"] == 500.0


def test_auto_settle_fifo_and_limits(world):
    F, supp = world["F"], world["supp"]
    pi1 = buy(F, supp, "ELH-T1", 10, 10, d="2026-01-05")                 # 115
    pi2 = buy(F, supp, "ELH-T1", 10, 10, d="2026-01-06")                 # 115
    adv = F.save_voucher("صرف", D, "صندوق نقدية الرياض", F.CAT_PARTNER, supp, 200)
    out = F.auto_settle_party(supp[0], "purchase")
    assert [round(a, 2) for _, _, a in out] == [115.0, 85.0]              # الأقدم أولاً
    assert F.invoice_payment_prefill("purchase", pi1)["remaining"] == 0
    assert F.invoice_payment_prefill("purchase", pi2)["remaining"] == 30.0
    with pytest.raises(F.UserError):                                    # لا متاح بعد الاستهلاك
        F.settle_invoice_from_advances("purchase", pi2)
    with pytest.raises(F.UserError):                                    # لا يقل مبلغ الدفعة عما سُوّي منها
        F.update_voucher(adv, D, "صندوق نقدية الرياض", 150, "")
    a_id = F.list_allocations(trans_id=adv)[-1][0]
    F.delete_allocation(a_id)
    assert F.invoice_payment_prefill("purchase", pi2)["remaining"] == 115.0
    with pytest.raises(F.UserError):                                    # نوع السند لا يناسب
        F.allocate_advance(adv, "sales", pi2)
    rep = F.run_report("open_inv_s", dict(d_to=F.today()))
    assert len(rep.rows) == 1 and rep.rows[0][11] == 85.0
    assert F.run_report("advances", {}).rows[0][7] == 85.0


# ============================================================ 4) واجهة سطح المكتب
def test_desktop_receipt_attach_roles_and_settlement(desk, monkeypatch):
    erp, errors, w, A = desk
    F = w["F"]
    oid = F.create_order("purchase", D, w["supp"], "", [F.make_line("ELH-T1", "x", "KG", 5, 10)])
    erp.refresh_flow("purchase")
    st = erp.flow["purchase"]
    label = next(lbl for lbl, v in st["f_order_cb"]._map.items() if v == oid)
    st["f_order_cb"].set(label)
    erp.load_fulfil_lines("purchase")
    import tempfile, os
    path = os.path.join(tempfile.mkdtemp(), "inv_771.pdf")
    with open(path, "wb") as fh:
        fh.write(PDF)
    monkeypatch.setattr(A.filedialog, "askopenfilenames", lambda **k: (path,))
    erp.pick_receipt_files()
    assert "1 ملف" in st["f_files_lbl"].cget("text")
    erp.save_fulfilment("purchase")
    with F.db() as conn:
        rid = conn.execute("SELECT MAX(receipt_id) FROM purchase_receipts").fetchone()[0]
    assert len(F.list_attachments("receipt", rid)) == 1
    assert st["f_tree"].item(str(rid), "values")[-1] == "📎 1"
    win, t = erp.open_supplier_attachments(w["supp"][0])
    assert len(t.get_children()) == 1
    win.destroy()
    # الأدوار في شاشة المستخدمين
    erp.users_load_roles(None)
    role_label = next(lbl for lbl in erp._roles if lbl.startswith("أمين المخزن"))
    erp.u_role_cb.set(role_label)
    erp.users_apply_role()
    assert erp.u_perm_vars["purchase.receipt_create"].get() == 1 and erp.u_perm_vars["cash"].get() == 0
    # تسوية الدفعات المقدمة من الخزينة
    F.invoice_from_fulfilments("purchase", w["supp"][0], [rid], D)          # 57.50
    F.save_voucher("صرف", D, "صندوق نقدية الرياض", F.CAT_PARTNER, w["supp"], 40)
    monkeypatch.setattr(A.messagebox, "askyesno", lambda *a, **k: True)
    erp.show_page("cash")
    erp.c_type.set(next(v for v in erp.c_type["values"] if v.startswith("صرف")))
    erp.party_category_cb.set(F.CAT_PARTNER)
    erp.update_party_options()
    erp.party_name_cb.set(next(lbl for lbl, v in erp.party_name_cb._map.items() if v and v[0] == w["supp"][0]))
    done = erp.settle_party_advances()
    assert done and round(done[0][2], 2) == 40.0
    assert not errors, errors


def test_desktop_order_party_fills_default_rep(desk):
    """اختيار العميل في أمر البيع يملأ مندوبه الافتراضي (كان يمرر (الكود، الاسم) لـ SQLite فيفشل)."""
    erp, errors, w, A = desk
    F = w["F"]
    F.save_rep(None, dict(name="سالم", commission_pct=2.0, active=1))
    F.save_partner_record("customer", w["cust"][0], w["cust"][1], rep="سالم", rep_date="2026-01-01")
    erp.refresh_flow("sales")
    st = erp.flow["sales"]
    st["party_cb"].set(next(lbl for lbl, v in st["party_cb"]._map.items() if v and v[0] == w["cust"][0]))
    st["party_cb"].event_generate("<<ComboboxSelected>>")
    erp.root.update()
    assert st["o_rep"].get() == "سالم"
    assert not errors, errors


# ============================================================ 5) توريث المندوب في السندات والعمولات
def test_voucher_inherits_rep_and_commissions_use_it(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    for n in ("سالم", "خالد"):
        F.save_rep(None, dict(name=n, commission_pct=2.0, active=1))
    buy(F, supp, "ELH-T1", 50, 10)
    # الفاتورة مسجلة بمندوب بينما بطاقة العميل بلا مندوب (الحالة الفعلية في البيانات)
    oid = F.create_order("sales", D, cust, "", [F.make_line("ELH-T1", "x", "KG", 10, 100)], rep="سالم")
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM sales_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    did = F.fulfil_order("sales", oid, D, "", [(lid, 10, None)])
    si, _ = F.invoice_from_fulfilments("sales", cust[0], [did], D)      # 1,150
    v1 = F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 1150, ("sales", si))
    v2 = F.save_voucher("قبض", D, "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 300)   # تحت الحساب
    with F.db() as conn:
        reps = dict(conn.execute("SELECT trans_id, rep_name FROM cash_transactions"))
    assert reps[v1] == "سالم"                      # من الفاتورة المسددة
    assert reps[v2] is None                         # الطرف بلا مندوب مسجل
    perf = {r[0]: r for r in F.rep_rep_performance(dict(d_from="2026-01-01", d_to="2026-12-31")).rows}
    assert perf["سالم"][4] == 1150 and perf[F.NO_REP][4] == 300
    # اقتراح ربط العميل بمندوب فواتيره ← توريث السند القديم بلا مندوب تلقائياً
    sug = F.suggest_partner_reps()
    assert [(s[0], s[2]) for s in sug] == [(cust[0], "سالم")]
    F.apply_partner_reps(sug)
    with F.db() as conn:
        assert conn.execute("SELECT rep_name FROM cash_transactions WHERE trans_id=?", (v2,)).fetchone()[0] == "سالم"
    rows = {r[1]: r for r in F.commission_rows(dict(d_from="2026-01-01", d_to="2026-12-31"), base="collections")[0]}
    assert rows["سالم"][4] == 1450 and rows["سالم"][6] == 29.0   # 2% من التحصيلات
    assert F.NO_REP not in rows or rows[F.NO_REP][4] == 0     # لا تحصيلات بدون مندوب (يبقى فقط شراء بلا مندوب)
    # سند جديد لعميل له مندوب في بطاقته يُورَّث منه مباشرة
    with F.db() as conn:
        assert F.voucher_rep_info(conn, cust[0], D) == ("سالم", "من بطاقة الطرف")


def test_desktop_cash_shows_inherited_rep(desk):
    erp, errors, w, A = desk
    F = w["F"]
    F.save_rep(None, dict(name="سالم", commission_pct=2.0, active=1))
    F.save_partner_record("customer", w["cust"][0], w["cust"][1], rep="سالم", rep_date="2026-01-01")
    erp.show_page("cash")
    erp.c_type.set(next(v for v in erp.c_type["values"] if v.startswith("قبض")))
    erp.party_category_cb.set(F.CAT_PARTNER)
    erp.update_party_options()
    erp.party_name_cb.set(next(lbl for lbl, v in erp.party_name_cb._map.items() if v and v[0] == w["cust"][0]))
    erp.refresh_open_invoices()
    assert "سالم" in erp.cash_rep_lbl.cget("text")
    assert not errors, errors

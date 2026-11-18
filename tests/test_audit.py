# -*- coding: utf-8 -*-
"""سجل التدقيق: التسجيل الآلي لكل تعديل/حذف، عدم القابلية للتعديل، وكشف التلاعب."""
import json
import sqlite3

import pytest

from conftest import buy, sell


def _log(F):
    return F.list_audit_log()


def test_partner_update_logged_with_before_after_and_user(world):
    F, cust = world["F"], world["cust"]
    F.set_current_user("ahmed", "web")
    F.save_partner_record("customer", cust[0], "اسم جديد", phone="0500")
    row = next(r for r in _log(F) if r[5] == "عميل / مورد" and r[4] == "تعديل")
    assert row[2] == "ahmed" and row[3] == "web" and row[6] == cust[0]
    old, new = json.loads(row[8]), json.loads(row[9])
    assert old["name"] == cust[1] and new["name"] == "اسم جديد" and new["phone"] == "0500"
    assert "name" in row[7] and "phone" in row[7]


def test_voucher_delete_and_invoice_line_snapshot(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 20, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 5, 20)
    vid = F.save_voucher("قبض", "2026-01-20", "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 50, ("sales", si))
    F.delete_voucher(vid)
    d = next(r for r in _log(F) if r[5] == "سند قبض/صرف" and r[4] == "حذف")
    assert json.loads(d[8])["amount"] == 50 and d[9] == ""
    _, lines = F.load_invoice("sales", si)
    F.update_invoice("sales", si, dict(date="2026-01-10", party=cust, notes=""), [dict(l, price=22.0) for l in lines])
    snap = next(r for r in _log(F) if r[5] == "بنود فاتورة مبيعات")
    assert json.loads(snap[8])[0]["unit_price"] == 20 and json.loads(snap[9])[0]["unit_price"] == 22
    head = next(r for r in _log(F) if r[5] == "فاتورة مبيعات" and r[4] == "تعديل")
    assert "subtotal" in head[7]
    F.delete_invoice("sales", si)
    assert any(r[5] == "فاتورة مبيعات" and r[4] == "حذف" for r in _log(F))
    assert any(r[5] == "بنود فاتورة مبيعات" and r[4] == "حذف" for r in _log(F))


def test_automatic_changes_are_not_noise(world):
    """تغيّر المسدد/الحالة/الكميات المنفذة/أرصدة المخزون آلياً ليس تعديلاً إدارياً."""
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 20, 10)
    si, _, _ = sell(F, cust, "ELH-T1", 5, 20)
    F.save_voucher("قبض", "2026-01-20", "صندوق نقدية الرياض", F.CAT_PARTNER, cust, 115, ("sales", si))
    assert not [r for r in _log(F) if r[5] in ("فاتورة مبيعات", "صنف", "أمر بيع", "بند أمر بيع")]


def test_password_is_masked(erp):
    F = erp
    F.save_user_account(None, "clerk1", "موظف", "Abcdef12!", "Abcdef12!", False, True, {"statement"})
    with F.db() as conn:
        uid = conn.execute("SELECT user_id FROM users WHERE username='clerk1'").fetchone()[0]
    F.save_user_account(uid, "clerk1", "موظف", "Xyzabc34!", "Xyzabc34!", False, True, {"statement"})
    row = next(r for r in _log(F) if r[5] == "مستخدم")
    assert "pbkdf2" not in (row[8] + row[9]).lower() and "مخفي" in row[9]


def test_log_is_immutable_and_chain_detects_tampering(world):
    F, cust = world["F"], world["cust"]
    F.save_partner_record("customer", cust[0], "تعديل 1")
    F.save_partner_record("customer", cust[0], "تعديل 2")
    assert F.verify_audit_chain()[0]
    with F.db() as conn:
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute("UPDATE audit_log SET username='x'")
    with F.db() as conn:
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute("DELETE FROM audit_log")
    # محاولة تلاعب خارجية: إسقاط المشغلات ثم تعديل سطر → السلسلة تكشفه
    raw = sqlite3.connect(F.engine.DB_FILE)
    raw.execute("DROP TRIGGER audit_log_no_update")
    raw.execute("UPDATE audit_log SET username='hacker' WHERE log_id=(SELECT MIN(log_id) FROM audit_log)")
    raw.commit()
    raw.close()
    ok, n, bad = F.verify_audit_chain()
    assert not ok and bad is not None
    F.init_db()          # يعيد إنشاء الحماية
    with F.db() as conn:
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute("UPDATE audit_log SET username='x'")


def test_audit_report_and_filters(world):
    F, cust = world["F"], world["cust"]
    F.set_current_user("u1", "desktop")
    F.save_partner_record("customer", cust[0], "x1")
    F.set_current_user("u2", "web")
    F.save_partner_record("customer", cust[0], "x2")
    assert all(r[2] == "u2" for r in F.list_audit_log(user="u2"))
    users, types = F.audit_filter_options()
    assert {"u1", "u2"} <= set(users) and "عميل / مورد" in types
    rep = F.audit_report(q="x2")
    assert rep.rows and "سليمة" in rep.notes[-1]

# -*- coding: utf-8 -*-
"""«نسيت كلمة المرور؟»: طلب معلق من شاشة الدخول، مراجعة المدير (تعيين كلمة مشفرة / رفض)، بلا كشف بيانات أو تغيير تلقائي."""
import os

import pytest


def _user(F, name="clerk.one", pw="Clerk-1234"):
    return F.save_user_account(None, name, "موظف", pw, pw, False, True, ["dashboard"])


def _hash(F, name):
    with F.db() as conn:
        return conn.execute("SELECT password_hash FROM users WHERE username=?", (name,)).fetchone()[0]


def test_request_is_pending_and_changes_nothing(erp):
    F = erp
    _user(F)
    before = _hash(F, "clerk.one")
    msg = F.request_password_reset("CLERK.ONE", "نسيت كلمتي", source="desktop")
    assert msg == F.PW_REQ_SENT_MSG and _hash(F, "clerk.one") == before            # لا تغيير تلقائي
    assert F.authenticate("clerk.one", "Clerk-1234")
    # اسم غير موجود: نفس الرسالة تماماً (لا كشف لأسماء المستخدمين) ولا يُحفظ شيء
    assert F.request_password_reset("no.such.user", "x") == msg
    F.request_password_reset("clerk.one", "تحديث الملاحظة")                         # طلب معلق واحد لكل مستخدم
    rows = F.list_password_requests(F.PW_REQ_PENDING)
    assert [(r[1], r[3], r[4]) for r in rows] == [("clerk.one", "تحديث الملاحظة", "معلق")]
    assert F.pending_password_request_count() == 1
    with pytest.raises(F.UserError):
        F.request_password_reset("  ")
    with F.db() as conn:
        audit = conn.execute("SELECT doc_ref, new_values FROM audit_log WHERE table_name='password_reset_requests'").fetchall()
    assert len(audit) == 2 and all(a[0] == "clerk.one" for a in audit)


def test_admin_completes_or_rejects(erp):
    F = erp
    _user(F)
    _user(F, "clerk.two", "Clerk-5678")
    F.request_password_reset("clerk.one")
    F.request_password_reset("clerk.two")
    req = {r[1]: r[0] for r in F.list_password_requests()}
    with F.db() as conn:   # قفل مؤقت قائم يُفك عند الإكمال
        conn.execute("UPDATE users SET failed_attempts=3, locked_until='2999-01-01 00:00:00' WHERE username='clerk.one'")
    with pytest.raises(F.UserError, match="غير متطابقتين"):
        F.resolve_password_request(req["clerk.one"], True, "New-Pass-11", "New-Pass-12")
    with pytest.raises(F.UserError, match="8 أحرف"):
        F.resolve_password_request(req["clerk.one"], True, "short1", "short1")
    F.set_current_user("clerk.two", "test", user=dict(is_admin=0, perms={"dashboard"}))
    try:
        with pytest.raises(F.UserError, match="لمدير النظام"):
            F.resolve_password_request(req["clerk.one"], True, "New-Pass-11", "New-Pass-11")
    finally:
        F.set_current_user("tester", "test")
    two_before = _hash(F, "clerk.two")
    F.resolve_password_request(req["clerk.one"], True, "New-Pass-11", "New-Pass-11", "تم بالهاتف")
    F.resolve_password_request(req["clerk.two"], False, admin_note="ليس صاحب الحساب")
    assert F.authenticate("clerk.one", "New-Pass-11") and _hash(F, "clerk.one").startswith("pbkdf2_sha256$")
    assert _hash(F, "clerk.two") == two_before and F.authenticate("clerk.two", "Clerk-5678")
    st = {r[1]: (r[4], r[8], r[9]) for r in F.list_password_requests()}
    assert st == {"clerk.one": ("مكتمل", "tester", "تم بالهاتف"), "clerk.two": ("مرفوض", "tester", "ليس صاحب الحساب")}
    with pytest.raises(F.UserError, match="تمت معالجته"):
        F.resolve_password_request(req["clerk.two"], True, "New-Pass-11", "New-Pass-11")
    assert F.pending_password_request_count() == 0
    with F.db() as conn:
        log = " ".join(v or "" for (v,) in conn.execute("SELECT new_values FROM audit_log"))
    assert "New-Pass-11" not in log and "Clerk-5678" not in log                       # لا كلمات مرور في السجل


def test_main_admin_password_untouched_by_requests(erp):
    F = erp
    main = F.main_admin_user()
    F.request_password_reset(main, "اختبار")
    assert F.authenticate(main, os.environ["ELHAGRY_ADMIN_PASSWORD"])

# -*- coding: utf-8 -*-
"""تأمين بيانات الدخول (لا قيم ثابتة، كلمات مرور مشفرة، ترحيل آمن بلا فقد بيانات) ومرفقات سندات القبض والصرف كملفات."""
import os

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ------------------------------------------------------------------ بيانات الدخول
def test_no_hardcoded_credentials_in_source():
    leaked = []
    for folder, _, files in os.walk(ROOT):
        if any(x in folder for x in (".git", "__pycache__", ".claude", "backups", "tests")):
            continue
        for f in files:
            if f.endswith((".py", ".md", ".toml", ".bat", ".pat", ".example")):
                txt = open(os.path.join(folder, f), encoding="utf-8", errors="ignore").read()
                if "Elhagry2012" in txt or "DEFAULT_ADMIN_PASSWORD" in txt:
                    leaked.append(f)
    assert not leaked, leaked
    ign = open(os.path.join(ROOT, ".gitignore"), encoding="utf-8").read().split()
    for rule in (".env", "*.db", "backups/", "attachments/", "initial_admin_password.txt"):
        assert rule in ign, rule
    assert os.path.exists(os.path.join(ROOT, ".env.example"))


def test_plaintext_admin_settings_are_hashed_without_changing_password(erp):
    F = erp
    with F.db() as conn:
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('admin_password', '4455')")
        conn.execute("DELETE FROM app_settings WHERE key='reset_password'")
    os.environ.pop("ELHAGRY_RESET_PASSWORD")
    try:
        F.init_db()                                           # ترحيل قاعدة قائمة
    finally:
        os.environ["ELHAGRY_RESET_PASSWORD"] = "2026"
    with F.db() as conn:
        vals = dict(conn.execute("SELECT key, value FROM app_settings WHERE key LIKE '%password%'").fetchall())
    assert vals["admin_password"].startswith("pbkdf2_sha256$") and "4455" not in vals["admin_password"]
    assert F.verify_modify_password("4455") and not F.verify_modify_password("2026")
    with pytest.raises(F.UserError, match="لم تُضبط"):          # لا كلمة تصفير افتراضية
        F.reset_transactions("2026")
    F.set_reset_password("Reset-77")
    assert F.get_setting("reset_password", "").startswith("pbkdf2_sha256$")
    with pytest.raises(F.UserError, match="غير صحيحة"):
        F.reset_transactions("2026")
    assert F.reset_transactions("Reset-77")["backup"]


def test_existing_admin_password_kept_and_changeable(erp):
    F = erp
    main = F.main_admin_user()
    assert main == os.environ["ELHAGRY_ADMIN_USER"]
    with F.db() as conn:
        uid = conn.execute("SELECT user_id FROM users WHERE username=?", (main,)).fetchone()[0]
    F.save_user_account(uid, main, "مدير", "NewPass-123", "NewPass-123", True, True, [])   # المدير يغيّر كلمته
    F.init_db()                                                # التشغيل التالي لا يعيد فرض أي كلمة مرور
    assert F.authenticate(main, "NewPass-123")["is_admin"]
    with pytest.raises(F.UserError):
        F.authenticate(main, os.environ["ELHAGRY_ADMIN_PASSWORD"])
    with pytest.raises(F.UserError, match="المدير الرئيسي"):
        F.delete_user_account(uid)
    with pytest.raises(F.UserError, match="المدير الرئيسي"):
        F.save_user_account(uid, "other.name", "مدير", "", "", True, True, [])
    os.environ["ELHAGRY_ADMIN_RESET"] = "1"                    # استعادة الدخول من البيئة
    try:
        F.init_db()
    finally:
        os.environ.pop("ELHAGRY_ADMIN_RESET")
    assert F.authenticate(main, os.environ["ELHAGRY_ADMIN_PASSWORD"])


def test_fresh_database_without_env_gets_random_password_file(erp, tmp_path, monkeypatch):
    F = erp
    from elhagry_core import engine as E
    monkeypatch.setattr(E, "DB_FILE", str(tmp_path / "fresh.db"))
    monkeypatch.setattr(E, "DATA_DIR", str(tmp_path))
    monkeypatch.delenv("ELHAGRY_ADMIN_USER")
    monkeypatch.delenv("ELHAGRY_ADMIN_PASSWORD")
    monkeypatch.delenv("ELHAGRY_RESET_PASSWORD")
    monkeypatch.delenv("ELHAGRY_MODIFY_PASSWORD")
    F.init_db()
    path = tmp_path / E.INITIAL_PASSWORD_FILE
    lines = path.read_text(encoding="utf-8").splitlines()
    user, pw = lines[0].split(": ", 1)[1], lines[1].split(": ", 1)[1]
    assert user == "admin" == F.main_admin_user() and len(pw) >= 12
    assert F.authenticate(user, pw)["is_admin"]
    with F.db() as conn:
        assert not conn.execute("SELECT 1 FROM app_settings WHERE key IN ('admin_password','reset_password')").fetchone()


def test_load_env_file(tmp_path, monkeypatch):
    from elhagry_core import engine as E
    env = tmp_path / ".env"
    env.write_text("# تعليق\nELHAGRY_T1=abc\nexport ELHAGRY_T2 = 'x y'\nELHAGRY_T3=\"q\"\nbad line\n", encoding="utf-8")
    monkeypatch.delenv("ELHAGRY_NO_DOTENV")
    monkeypatch.setenv("ELHAGRY_T3", "kept")                 # متغير البيئة الموجود له الأولوية
    for k in ("ELHAGRY_T1", "ELHAGRY_T2"):
        monkeypatch.delenv(k, raising=False)
    E.load_env_file(str(env))
    try:
        assert (os.environ["ELHAGRY_T1"], os.environ["ELHAGRY_T2"], os.environ["ELHAGRY_T3"]) == ("abc", "x y", "kept")
    finally:
        os.environ.pop("ELHAGRY_T1", None)
        os.environ.pop("ELHAGRY_T2", None)


# ------------------------------------------------------------------ مرفقات السندات
@pytest.fixture
def att_dir(world, tmp_path, monkeypatch):
    from elhagry_core import engine as E
    monkeypatch.setattr(E, "ATTACH_DIR", str(tmp_path / "attachments"))
    return world, E


def test_voucher_attachments_saved_as_files(att_dir):
    w, E = att_dir
    F = w["F"]
    vid = F.save_voucher("قبض", "2026-01-11", "صندوق نقدية الرياض", F.CAT_PARTNER, w["cust"], 50)
    pdf = b"%PDF-1.4 receipt"
    aid = F.add_attachment("voucher", vid, "إيصال: بنك 1.pdf", pdf, "tester")
    img = F.add_attachment("voucher", vid, "photo.JPG", b"\xff\xd8\xff jpeg", "tester")
    with F.db() as conn:
        rel, blob = conn.execute("SELECT file_path, length(data) FROM attachments WHERE attach_id=?", (aid,)).fetchone()
    path = os.path.join(E.ATTACH_DIR, *rel.split("/"))
    assert rel.startswith(f"voucher/{vid}/") and blob == 0 and open(path, "rb").read() == pdf
    assert ":" not in os.path.basename(path) and "/" not in os.path.basename(path)
    assert F.get_attachment(aid) == ("إيصال: بنك 1.pdf", "application/pdf", pdf)
    assert F.get_attachment(img)[1] == "image/jpeg"
    assert F.attachment_counts("voucher") == {vid: 2}
    assert [r[0] for r in F.list_attachments("voucher", vid)] == [aid, img]
    with pytest.raises(F.UserError, match="غير مدعوم"):
        F.add_attachment("voucher", vid, "x.exe", b"MZ")
    with pytest.raises(F.UserError, match="غير موجود"):
        F.add_attachment("voucher", 99999, "a.pdf", pdf)
    F.delete_attachment(img)                                  # الحذف ينقل الملف للأرشيف ولا يتلفه
    assert F.attachment_counts("voucher") == {vid: 1}
    deleted = os.path.join(E.ATTACH_DIR, "_deleted")
    assert any(f.endswith("photo.JPG") for _, _, fs in os.walk(deleted) for f in fs)
    F.delete_voucher(vid)                                     # حذف السند ← مرفقاته تُؤرشف
    assert F.attachment_counts("voucher") == {} and not os.path.exists(path)
    assert sum(len(fs) for _, _, fs in os.walk(deleted)) == 2


def test_legacy_blob_attachment_still_readable_and_permissions(att_dir):
    w, E = att_dir
    F = w["F"]
    vid = F.save_voucher("صرف", "2026-01-11", "صندوق نقدية الرياض", F.CAT_PARTNER, w["supp"], 20)
    with F.db() as conn:   # مرفق قديم محفوظ داخل القاعدة (قبل مجلد المرفقات)
        old = conn.execute("INSERT INTO attachments (doc_type, doc_id, file_name, mime, size, data) VALUES "
                           "('voucher', ?, 'old.png', 'image/png', 3, ?)", (vid, b"PNG")).lastrowid
    assert F.get_attachment(old) == ("old.png", "image/png", b"PNG")
    F.set_current_user("clerk", "test", user=dict(is_admin=0, perms={"sales"}))
    try:
        with pytest.raises(F.UserError, match="صلاحية"):
            F.add_attachment("voucher", vid, "a.pdf", b"%PDF")
    finally:
        F.set_current_user("tester", "test")
    F.add_attachment("voucher", vid, "a.pdf", b"%PDF")
    res = F.reset_transactions("2026")                       # التصفير يؤرشف الملفات ولا يتلفها
    assert res["backup"] and F.attachment_counts("voucher") == {}
    assert any(d.startswith("reset_") for d in os.listdir(os.path.join(E.ATTACH_DIR, "_deleted")))

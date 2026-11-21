# -*- coding: utf-8 -*-
"""
سجل التدقيق والرقابة الداخلية (Audit Log).

- كل تعديل أو حذف على المستندات والبيانات الأساسية يُسجَّل تلقائياً بمشغلات SQLite (Triggers)، فلا يمكن لأي شاشة
  (سطح المكتب أو الويب) أو أي كود أن يعدّل/يحذف دون أن يُسجَّل: اسم المستخدم، التوقيت، الواجهة، نوع المستند ورقمه،
  القيم قبل التعديل وبعده، والحقول التي تغيّرت.
- بنود المستندات المركبة (بنود الفواتير وأسطر القيود) تُسجَّل كلقطة كاملة قبل/بعد من المحرك نفسه (audit_event).
- السجل غير قابل للتعديل أو الحذف: مشغلات تمنع UPDATE و DELETE على الجدول نفسه.
- سلسلة بصمات SHA-256 (كل سطر يحمل بصمة السطر السابق): أي تلاعب خارجي بالملف (حذف/تعديل سطر أو إسقاط المشغلات)
  يكشفه verify_audit_chain().
"""
import hashlib
import json
import sqlite3
from datetime import datetime

from . import engine as E
from .engine import db, Report

AUDIT_SQL = """
CREATE TABLE IF NOT EXISTS audit_log (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    logged_at TEXT NOT NULL,
    username TEXT,
    source TEXT,
    action TEXT NOT NULL,
    doc_type TEXT,
    table_name TEXT,
    doc_ref TEXT,
    old_values TEXT,
    new_values TEXT,
    changed_fields TEXT,
    row_hash TEXT
);
CREATE INDEX IF NOT EXISTS idx_audit_at ON audit_log(logged_at);
CREATE INDEX IF NOT EXISTS idx_audit_doc ON audit_log(table_name, doc_ref);
DROP TRIGGER IF EXISTS audit_log_no_update;
CREATE TRIGGER audit_log_no_update BEFORE UPDATE ON audit_log
BEGIN SELECT RAISE(ABORT, 'سجل التدقيق غير قابل للتعديل'); END;
DROP TRIGGER IF EXISTS audit_log_no_delete;
CREATE TRIGGER audit_log_no_delete BEFORE DELETE ON audit_log
BEGIN SELECT RAISE(ABORT, 'سجل التدقيق غير قابل للحذف'); END;
"""

ACT_UPDATE, ACT_DELETE, ACT_EVENT = "تعديل", "حذف", "حدث إداري"

# الجدول ← (اسم المستند، تعبير الرقم المرجعي، أعمدة لا تُعد تعديلاً إدارياً (تتغير آلياً)، أعمدة تُخفى قيمتها، شرط إضافي)
AUDITED = {
    "sales_invoices": ("فاتورة مبيعات", "COALESCE(R.invoice_no, R.invoice_id)", {"paid_amount", "status"}, set(), None),
    "purchase_invoices": ("فاتورة مشتريات", "COALESCE(R.invoice_no, R.purchase_id)", {"paid_amount", "status"}, set(), None),
    "returns_invoices": ("مرتجع", "R.return_id", set(), set(), None),
    "cash_transactions": ("سند قبض/صرف", "R.trans_id", set(), set(), None),
    "journal_entries": ("قيد يومية", "R.entry_id", set(), set(), "R.source_type IS NULL"),
    "post_discounts": ("خصم لاحق / إشعار دائن", "COALESCE(R.discount_no, R.discount_id)",
                       {"journal_entry_id", "stock_applied"}, set(), None),
    "sales_orders": ("أمر بيع", "R.order_id", {"status"}, set(), None),
    "purchase_orders": ("أمر شراء", "R.order_id", {"status"}, set(), None),
    "sales_order_lines": ("بند أمر بيع", "R.order_id || '/' || R.line_id", {"delivered_qty", "invoiced_qty"}, set(), None),
    "purchase_order_lines": ("بند أمر شراء", "R.order_id || '/' || R.line_id", {"received_qty", "invoiced_qty"}, set(), None),
    "sales_deliveries": ("إذن تسليم", "R.delivery_id", {"invoice_id"}, set(), None),
    "purchase_receipts": ("إذن استلام", "R.receipt_id", {"invoice_id"}, set(), None),
    "sales_delivery_lines": ("بند إذن تسليم", "R.delivery_id || '/' || R.line_id", {"unit_cost"}, set(), None),
    "purchase_receipt_lines": ("بند إذن استلام", "R.receipt_id || '/' || R.line_id", set(), set(), None),
    "business_partners": ("عميل / مورد", "R.partner_code", set(), set(), None),
    "inventory": ("صنف", "R.item_code", {"quantity", "cost_price"}, set(), None),
    "fixed_assets": ("أصل ثابت", "R.asset_code", set(), set(), None),
    "chart_of_accounts": ("حساب بالدليل", "R.account_code", {"is_group"}, set(), None),
    "users": ("مستخدم", "R.username", {"failed_attempts", "locked_until", "last_login"}, {"password_hash"}, None),
    "sales_reps": ("مندوب", "COALESCE(R.rep_code, R.name)", set(), set(), None),
    "company_profile": ("بيانات المنشأة", "R.id", set(), set(), None),
    "daily_closures": ("إغلاق يومي", "R.closure_date", set(), set(), None),
    "attachments": ("مرفق", "R.doc_type || '#' || R.doc_id || ' ' || R.file_name", set(), {"data"}, None),
    "budgets": ("موازنة تقديرية", "R.period || ' ' || R.line_key", set(), set(), None),
    "sequences": ("تسلسل ترقيم", "R.name", {"next_value"}, set(), None),
}
SECRET_TEXT = "*** (مخفي)"


def _hash(prev, payload):
    return hashlib.sha256(((prev or "") + "|" + (payload or "")).encode("utf-8")).hexdigest()


def _sql_quote(s):
    return "'" + str(s).replace("'", "''") + "'"


def _json_expr(alias, cols, secret):
    parts = []
    for c in cols:
        if c in secret:
            parts.append(f"{_sql_quote(c)}, CASE WHEN {alias}.{c} IS NULL THEN NULL ELSE {_sql_quote(SECRET_TEXT)} END")
        else:
            parts.append(f"{_sql_quote(c)}, {alias}.{c}")
    return f"json_object({', '.join(parts)})" if parts else "'{}'"


def _create_triggers(conn):
    conn.executescript(AUDIT_SQL)
    for table, (label, ref, ignore, secret, cond) in AUDITED.items():
        cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]
        for kind in ("u", "d"):
            conn.execute(f"DROP TRIGGER IF EXISTS audit_{table}_{kind}")
        if not cols:
            continue
        tracked = [c for c in cols if c not in ignore]
        changed = " || ".join(f"CASE WHEN OLD.{c} IS NOT NEW.{c} THEN {_sql_quote(c + ', ')} ELSE '' END" for c in tracked)
        when_u = " OR ".join(f"OLD.{c} IS NOT NEW.{c}" for c in tracked) or "0"
        cond_old = f" AND ({cond.replace('R.', 'OLD.')})" if cond else ""

        def body(action, old_json, new_json, ch, ref_alias):
            r = ref.replace("R.", f"{ref_alias}.")
            return (f"INSERT INTO audit_log (logged_at, username, source, action, doc_type, table_name, doc_ref, old_values, "
                    f"new_values, changed_fields, row_hash) "
                    f"SELECT ts, u, s, a, dt, tn, rf, o, n, ch, audit_hash(prev, ts||'|'||u||'|'||a||'|'||tn||'|'||"
                    f"COALESCE(rf,'')||'|'||COALESCE(o,'')||'|'||COALESCE(n,'')) FROM ("
                    f"SELECT strftime('%Y-%m-%d %H:%M:%S','now','localtime') AS ts, app_user() AS u, app_source() AS s, "
                    f"{_sql_quote(action)} AS a, {_sql_quote(label)} AS dt, {_sql_quote(table)} AS tn, CAST({r} AS TEXT) AS rf, "
                    f"{old_json} AS o, {new_json} AS n, {ch} AS ch, "
                    f"(SELECT row_hash FROM audit_log ORDER BY log_id DESC LIMIT 1) AS prev);")

        conn.execute(f"CREATE TRIGGER audit_{table}_u AFTER UPDATE ON {table} WHEN ({when_u}){cond_old} BEGIN "
                     + body(ACT_UPDATE, _json_expr("OLD", tracked, secret), _json_expr("NEW", tracked, secret),
                            f"rtrim({changed or chr(39) * 2}, ', ')", "NEW") + " END")
        conn.execute(f"CREATE TRIGGER audit_{table}_d AFTER DELETE ON {table} WHEN 1{cond_old} BEGIN "
                     + body(ACT_DELETE, _json_expr("OLD", cols, secret), "NULL", "NULL", "OLD") + " END")


def setup_audit(conn):
    _create_triggers(conn)


def _register(conn):
    conn.create_function("audit_hash", 2, _hash)


E.CONNECT_HOOKS.append(_register)
E.INIT_HOOKS.append(setup_audit)   # آخر وحدة تُهيأ: بعد إنشاء كل الجداول


def audit_event(conn, action, doc_type, doc_ref="", old=None, new=None, table_name="", changed=""):
    """تسجيل حدث من المحرك (لقطة بنود مستند قبل/بعد، تصفير، استعادة، إعادة ترحيل...). داخل نفس المعاملة."""
    try:
        prev = conn.execute("SELECT row_hash FROM audit_log ORDER BY log_id DESC LIMIT 1").fetchone()
    except sqlite3.OperationalError:
        return
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    u, src = E.current_user_name(), E.current_user_source()
    o = json.dumps(old, ensure_ascii=False, default=str) if old is not None else None
    n = json.dumps(new, ensure_ascii=False, default=str) if new is not None else None
    ref = str(doc_ref or "")
    h = _hash(prev[0] if prev else None, f"{ts}|{u}|{action}|{table_name}|{ref}|{o or ''}|{n or ''}")
    conn.execute("INSERT INTO audit_log (logged_at, username, source, action, doc_type, table_name, doc_ref, old_values, "
                 "new_values, changed_fields, row_hash) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                 (ts, u, src, action, doc_type, table_name, ref, o, n, changed, h))


AUDIT_GUARD_SQL = """
CREATE TRIGGER IF NOT EXISTS audit_log_no_update BEFORE UPDATE ON audit_log
BEGIN SELECT RAISE(ABORT, 'سجل التدقيق غير قابل للتعديل'); END;
CREATE TRIGGER IF NOT EXISTS audit_log_no_delete BEFORE DELETE ON audit_log
BEGIN SELECT RAISE(ABORT, 'سجل التدقيق غير قابل للحذف'); END;
"""


def clear_audit_log(conn, backup_name, deleted=None):
    """مسح سجل التدقيق بالكامل عند التصفير فقط (داخل معاملة التصفير وبعد نسخته الاحتياطية).
    يُرفع مشغل منع الحذف مؤقتاً ثم يُعاد فوراً، ويُصفَّر الترقيم، ويُكتب سطر افتتاحي يبدأ سلسلة بصمات جديدة.
    يرجع عدد السطور الممسوحة."""
    (n,) = conn.execute("SELECT COUNT(*) FROM audit_log").fetchone()
    conn.execute("DROP TRIGGER IF EXISTS audit_log_no_delete")
    conn.execute("DELETE FROM audit_log")
    for stmt in AUDIT_GUARD_SQL.strip().split("END;"):
        if stmt.strip():
            conn.execute(stmt + "END;")
    if conn.execute("SELECT 1 FROM sqlite_master WHERE name='sqlite_sequence'").fetchone():
        conn.execute("DELETE FROM sqlite_sequence WHERE name='audit_log'")
    audit_event(conn, ACT_EVENT, "تصفير الحركات ومسح سجل التدقيق", "بداية سجل جديد",
                new={"backup": backup_name, "audit_rows_cleared": n,
                     "documents_deleted": sum((deleted or {}).values())})
    return n


def lines_snapshot(conn, table, key_col, key, cols):
    rows = conn.execute(f"SELECT {', '.join(cols)} FROM {table} WHERE {key_col}=? ORDER BY rowid", (key,)).fetchall()
    return [dict(zip(cols, r)) for r in rows]


def verify_audit_chain():
    """يتحقق من سلسلة البصمات: (سليم؟، عدد السجلات، أول رقم سجل مخالف أو None)."""
    with db() as conn:
        rows = conn.execute("SELECT log_id, logged_at, username, action, table_name, doc_ref, old_values, new_values, row_hash "
                            "FROM audit_log ORDER BY log_id").fetchall()
        triggers = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='trigger' AND name LIKE 'audit_log_no_%'")}
    prev, expect_id = None, None
    for lid, ts, u, a, tn, rf, o, n, h in rows:
        if expect_id is not None and lid != expect_id:
            return False, len(rows), lid
        calc = _hash(prev, f"{ts}|{u}|{a}|{tn or ''}|{rf or ''}|{o or ''}|{n or ''}")
        if calc != h:
            return False, len(rows), lid
        prev, expect_id = h, lid + 1
    if {"audit_log_no_update", "audit_log_no_delete"} - triggers:
        return False, len(rows), None
    return True, len(rows), None


def _diff_text(old, new):
    try:
        o = json.loads(old) if old else {}
        n = json.loads(new) if new else {}
    except (TypeError, ValueError):
        return ""
    if not isinstance(o, dict) or not isinstance(n, dict):
        return ""
    return " | ".join(f"{k}: {o.get(k)} ← {n.get(k)}" for k in n if o.get(k) != n.get(k))


def list_audit_log(d_from=None, d_to=None, user="", doc_type="", action="", q="", limit=2000):
    """[(رقم، التوقيت، المستخدم، الواجهة، العملية، نوع المستند، الرقم، الحقول المتغيرة، قبل، بعد)] الأحدث أولاً."""
    where, args = ["1=1"], []
    if d_from:
        where.append("logged_at >= ?")
        args.append(d_from)
    if d_to:
        where.append("logged_at <= ?")
        args.append(d_to + " 23:59:59")
    for col, v in (("username", user), ("doc_type", doc_type), ("action", action)):
        if v:
            where.append(f"{col} = ?")
            args.append(v)
    if q:
        where.append("(COALESCE(doc_ref,'') LIKE ? OR COALESCE(old_values,'') LIKE ? OR COALESCE(new_values,'') LIKE ?)")
        args += [f"%{q}%"] * 3
    with db() as conn:
        try:
            rows = conn.execute(f"SELECT log_id, logged_at, COALESCE(username,''), COALESCE(source,''), action, "
                                f"COALESCE(doc_type,''), COALESCE(doc_ref,''), COALESCE(changed_fields,''), "
                                f"COALESCE(old_values,''), COALESCE(new_values,'') FROM audit_log WHERE {' AND '.join(where)} "
                                f"ORDER BY log_id DESC LIMIT ?", args + [int(limit)]).fetchall()
        except sqlite3.OperationalError:
            return []
    return rows


def audit_filter_options():
    with db() as conn:
        try:
            users = [r[0] for r in conn.execute("SELECT DISTINCT username FROM audit_log WHERE username IS NOT NULL ORDER BY 1")]
            types = [r[0] for r in conn.execute("SELECT DISTINCT doc_type FROM audit_log WHERE doc_type IS NOT NULL ORDER BY 1")]
        except sqlite3.OperationalError:
            return [], []
    return users, types


def audit_report(**filters):
    rows = list_audit_log(**filters)
    out = [(r[0], r[1], r[2], {"web": "الويب", "desktop": "سطح المكتب"}.get(r[3], r[3] or "-"), r[4], r[5], r[6],
            r[7] or ("كل الحقول" if r[4] == ACT_DELETE else ""), _diff_text(r[8], r[9]) or r[8][:300], r[9][:300])
           for r in rows]
    cols = [("رقم السجل", "i"), ("التوقيت", "t"), ("المستخدم", "t"), ("الواجهة", "t"), ("العملية", "t"), ("نوع المستند", "t"),
            ("رقم المستند", "t"), ("الحقول المتغيرة", "t"), ("القيمة قبل ← بعد", "t"), ("القيمة بعد (كاملة)", "t")]
    ok, n, bad = verify_audit_chain()
    notes = ["السجل غير قابل للتعديل أو الحذف، ومحمي بسلسلة بصمات SHA-256.",
             f"فحص سلامة السلسلة: {'سليمة ✓' if ok else 'تنبيه: تلاعب محتمل عند السجل ' + str(bad)} ({n} سجل)."]
    return Report("سجل التدقيق والرقابة الداخلية", cols, out, None, notes, "")

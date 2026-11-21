# -*- coding: utf-8 -*-
"""
الترحيل الآلي إلى دفتر الأستاذ العام (Automated General Ledger Posting).

الفكرة:
- الدفتر العام = القيود اليومية (اليدوية + قيود الخصم اللاحق الآلية في journal_lines)
                 + قيود الترحيل الآلي لكل المستندات (gl_entries / gl_lines).
- قيود الترحيل الآلي تُبنى من المستند نفسه (مصدر الحقيقة) ولا تُكتب يدوياً، فلا يمكن أن تختلف عنه.
- مشغلات SQLite (Triggers) تسجّل أي مستند أُضيف/عُدّل/حُذف في طابور gl_queue، ويُرحَّل الطابور داخل نفس المعاملة
  قبل الحفظ (db() ← COMMIT_HOOKS)، فالترحيل فوري وذري مع المستند، ويشمل التعديلات غير المباشرة
  (مثل إعادة احتساب التكلفة بالمتوسط المرجح recalc_items بعد تعديل إداري).
- لا تُضاف قيود الترحيل إلى journal_lines عمداً: قائمة الدخل وتقارير الخزينة الحالية تُحسب من الدفاتر المساعدة
  + القيود اليدوية، فإضافتها هناك تعني ازدواجاً.

قواعد الترحيل (القيد المزدوج):
    فاتورة مبيعات      من ح/ العملاء (الإجمالي)       إلى ح/ إيراد المبيعات (الصافي) + ح/ ضريبة المخرجات
                        من ح/ تكلفة المبيعات            إلى ح/ بضاعة مسلّمة لم تُفوتر (بنود الأذون) أو ح/ المخزون (بنود مباشرة)
    إذن تسليم           من ح/ بضاعة مسلّمة لم تُفوتر     إلى ح/ المخزون (بالتكلفة)
    إذن استلام          من ح/ المخزون (بالتكلفة)         إلى ح/ بضاعة مستلمة لم تصل فاتورتها
    فاتورة مشتريات     من ح/ بضاعة مستلمة لم تصل فاتورتها (أو المخزون للمباشر) + ح/ ضريبة المدخلات  إلى ح/ الموردين
    مرتجع مبيعات       من ح/ مردودات المبيعات + ح/ ضريبة المخرجات  إلى ح/ العملاء؛  من ح/ المخزون إلى ح/ تكلفة المبيعات
    مرتجع مشتريات      من ح/ الموردين  إلى ح/ المخزون (متوسط التكلفة) + ح/ ضريبة المدخلات ± ح/ فروق أسعار المردودات
    سند قبض / صرف      النقدية/البنك مقابل العميل أو المورد أو حساب المصروف/الأصل/العهدة المختار
    أرصدة افتتاحية      العملاء/الموردون/المخزون مقابل ح/ الوسيط الافتتاحي
    الإهلاك             من ح/ مصروف الإهلاك  إلى ح/ مجمع الإهلاك (شهرياً بداية كل شهر يستحق فيه)
"""
import sqlite3
from collections import defaultdict
from datetime import date, datetime

from . import engine as E
from .engine import (UserError, db, CAT_PARTNER, SALES_RETURN, TYPE_ASSET, TYPE_LIAB, TYPE_EQUITY, TYPE_REV, TYPE_EXP,
                     Report, today, _account_maps, _tree_report_rows, register_report_group)

GL_VERSION = "2"   # تغيير قواعد الترحيل = إعادة ترحيل شاملة تلقائية عند أول تشغيل

GL_SQL = """
CREATE TABLE IF NOT EXISTS gl_entries (
    gl_id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    source_type TEXT NOT NULL,
    source_id INTEGER NOT NULL,
    ref TEXT,
    description TEXT,
    posted_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS gl_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    gl_id INTEGER NOT NULL REFERENCES gl_entries(gl_id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    debit REAL NOT NULL DEFAULT 0 CHECK (debit >= 0),
    credit REAL NOT NULL DEFAULT 0 CHECK (credit >= 0),
    memo TEXT,
    partner_code TEXT
);
CREATE TABLE IF NOT EXISTS gl_queue (
    source_type TEXT NOT NULL,
    source_id INTEGER NOT NULL,
    PRIMARY KEY (source_type, source_id)
) WITHOUT ROWID;
CREATE INDEX IF NOT EXISTS idx_gl_src ON gl_entries(source_type, source_id);
CREATE INDEX IF NOT EXISTS idx_gl_date ON gl_entries(entry_date);
CREATE INDEX IF NOT EXISTS idx_gll_entry ON gl_lines(gl_id);
CREATE INDEX IF NOT EXISTS idx_gll_acc ON gl_lines(account_code);
"""

SRC_LABEL = {
    "sales_invoice": "فاتورة مبيعات", "purchase_invoice": "فاتورة مشتريات", "delivery": "إذن تسليم",
    "receipt": "إذن استلام", "return": "مرتجع", "voucher": "سند", "partner_opening": "رصيد افتتاحي طرف",
    "item_opening": "رصيد افتتاحي صنف", "asset": "إهلاك أصل", "journal": "قيد يومية", "post_discount": "خصم لاحق (قيد آلي)",
}

# (الجدول، نوع المصدر، تعبير المعرّف، أعمدة التحديث المؤثرة أو None = كل الأعمدة)
_TRIGGERS = [
    ("sales_invoices", "sales_invoice", "invoice_id",
     "invoice_date, customer_code, customer_name, subtotal, vat_amount, total_amount, invoice_no"),
    ("sales_invoice_lines", "sales_invoice", "invoice_id", None),
    ("purchase_invoices", "purchase_invoice", "purchase_id",
     "purchase_date, supplier_code, supplier_name, subtotal, vat_amount, total_amount, invoice_no"),
    ("purchase_invoice_lines", "purchase_invoice", "purchase_id", None),
    ("sales_deliveries", "delivery", "delivery_id", "delivery_date, customer_code, customer_name"),
    ("sales_delivery_lines", "delivery", "delivery_id", None),
    ("purchase_receipts", "receipt", "receipt_id", "receipt_date, supplier_code, supplier_name"),
    ("purchase_receipt_lines", "receipt", "receipt_id", None),
    ("returns_invoices", "return", "return_id", None),
    ("cash_transactions", "voucher", "trans_id", None),
    ("business_partners", "partner_opening", "partner_id",
     "opening_debit, opening_credit, opening_date, is_customer, is_supplier, partner_code"),
    ("inventory", "item_opening", "rowid", "opening_qty, opening_cost, item_code"),
    ("fixed_assets", "asset", "asset_id", None),
]

# حسابات الترحيل الآلي (قابلة للتغيير من الإعدادات): المفتاح ← (الكود الافتراضي، الوصف، النوع، الحساب الأب عند الإنشاء)
GL_ACCOUNT_DEFAULTS = {
    "gl_acc_ar": ("11250", "العملاء (ذمم مدينة)", TYPE_ASSET, "11200"),
    "gl_acc_ap": ("22101", "الموردون (ذمم دائنة)", TYPE_LIAB, "22100"),
    "gl_acc_stock": ("11310", "المخزون", TYPE_ASSET, "11300"),
    "gl_acc_gdni": ("", "بضاعة مسلّمة لم تُفوتر (وسيط تسليم)", TYPE_ASSET, "11300"),
    "gl_acc_grni": ("", "بضاعة مستلمة لم تصل فاتورتها (وسيط استلام)", TYPE_LIAB, "22100"),
    "gl_acc_sales": ("31110", "إيراد المبيعات", TYPE_REV, "31100"),
    "gl_acc_sales_ret": ("31200", "مردودات المبيعات", TYPE_REV, "31000"),
    "gl_acc_cogs": ("41110", "تكلفة المبيعات", TYPE_EXP, "41100"),
    "gl_acc_pret_diff": ("41201", "فروق أسعار مردودات المشتريات", TYPE_EXP, "41200"),
    "gl_acc_vat_in": ("22501", "ضريبة القيمة المضافة - مدخلات", TYPE_LIAB, "22500"),
    "gl_acc_vat_out": ("22502", "ضريبة القيمة المضافة - مخرجات", TYPE_LIAB, "22500"),
    "gl_acc_opening": ("99999", "وسيط افتتاحي (أرصدة أول المدة)", TYPE_EQUITY, None),
    "gl_acc_dep_exp": ("42108", "مصروف الإهلاك", TYPE_EXP, "42100"),
    "gl_acc_dep_acc": ("22208", "مجمع الإهلاك", TYPE_LIAB, "22200"),
    "gl_acc_exp_general": ("", "مصروفات عمومية - سندات بدون حساب محدد", TYPE_EXP, "40000"),
    "gl_acc_suspense": ("", "حساب معلق - ترحيل آلي (حسابات غير معرّفة)", TYPE_LIAB, None),
}
SUSPENSE_ROOT_CODE = "99998"
# حسابات تُعرض في الميزانية مخصومة من الأصول رغم تصنيفها (مثل مجمعات الإهلاك 222xx المصنفة خصوماً في الدليل المستورد)
CONTRA_ASSET_SETTING = "bs_contra_prefixes"
CONTRA_ASSET_DEFAULT = "222"


# ============================================================
# البنية: الجداول والمشغلات والحسابات
# ============================================================
def _create_triggers(conn):
    for table, src, idcol, upd_cols in _TRIGGERS:
        if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone():
            continue
        tcols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        cols = [c.strip() for c in upd_cols.split(",")] if upd_cols else []
        cols = [c for c in cols if c in tcols]
        base = f"gl_q_{table}"
        for kind in ("i", "u", "d"):
            conn.execute(f"DROP TRIGGER IF EXISTS {base}_{kind}")
        conn.execute(f"CREATE TRIGGER {base}_i AFTER INSERT ON {table} BEGIN "
                     f"INSERT OR IGNORE INTO gl_queue VALUES ('{src}', NEW.{idcol}); END")
        of = f" OF {', '.join(cols)}" if cols else ""
        conn.execute(f"CREATE TRIGGER {base}_u AFTER UPDATE{of} ON {table} BEGIN "
                     f"INSERT OR IGNORE INTO gl_queue VALUES ('{src}', OLD.{idcol}); "
                     f"INSERT OR IGNORE INTO gl_queue VALUES ('{src}', NEW.{idcol}); END")
        conn.execute(f"CREATE TRIGGER {base}_d AFTER DELETE ON {table} BEGIN "
                     f"INSERT OR IGNORE INTO gl_queue VALUES ('{src}', OLD.{idcol}); END")


def _ensure_account(conn, key):
    """كود الحساب المستخدم لمفتاح الترحيل: الإعداد المحفوظ ← الافتراضي ← إنشاء حساب نظامي فرعي تلقائياً."""
    default, label, typ, parent = GL_ACCOUNT_DEFAULTS[key]
    row = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
    for code in ((row[0] if row else None), default):
        if not code:
            continue
        acc = conn.execute("SELECT COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone()
        if acc is not None and not acc[0]:
            if not row or row[0] != code:
                conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (key, code))
            return code
    # إنشاء حساب نظامي: تحت الأب المقترح إن وُجد كحساب، وإلا في جذر الدليل
    par = conn.execute("SELECT account_type FROM chart_of_accounts WHERE account_code=?", (parent,)).fetchone() if parent else None
    if par is not None:
        conn.execute("UPDATE chart_of_accounts SET is_group=1 WHERE account_code=?", (parent,))
        code = E.next_child_account_code(conn, parent)
    else:
        code = default or SUSPENSE_ROOT_CODE
        while conn.execute("SELECT 1 FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone():
            code = str(int(code) - 1) if code.isdigit() else code + "1"
        parent = None
    conn.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                 "VALUES (?,?,?,?,0)", (code, label, typ, parent))
    conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (key, code))
    return code


def gl_accounts(conn):
    return {k: _ensure_account(conn, k) for k in GL_ACCOUNT_DEFAULTS}


def set_gl_accounts(mapping):
    """حفظ حسابات الترحيل الآلي (حسابات فرعية موجودة فقط) ثم إعادة ترحيل كل المستندات."""
    with db() as conn:
        for key, code in mapping.items():
            if key not in GL_ACCOUNT_DEFAULTS:
                continue
            code = (code or "").strip()
            label = GL_ACCOUNT_DEFAULTS[key][1]
            acc = conn.execute("SELECT COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone()
            if acc is None:
                raise UserError(f"حساب «{label}»: الكود {code} غير موجود في دليل الحسابات.")
            if acc[0]:
                raise UserError(f"حساب «{label}»: {code} حساب تجميعي. اختر حساباً فرعياً.")
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (key, code))
        _rebuild(conn)


def setup_ledger(conn):
    """يُستدعى من init_db: الجداول، المشغلات، الحسابات النظامية، وأول ترحيل شامل (أو عند تغيير قواعد الترحيل)."""
    conn.executescript(GL_SQL)
    _create_triggers(conn)
    gl_accounts(conn)
    v = conn.execute("SELECT value FROM app_settings WHERE key='gl_version'").fetchone()
    if not v or v[0] != GL_VERSION:
        conn.execute("INSERT OR IGNORE INTO gl_queue VALUES ('all', 0)")
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_version', ?)", (GL_VERSION,))


# ============================================================
# قواعد الترحيل لكل مستند
# ============================================================
def _r(x):
    return round(x or 0.0, 2)


class _Entry:
    def __init__(self, d, ref, desc):
        self.date, self.ref, self.desc, self.lines = d, ref, desc, []

    def dr(self, acc, amt, memo="", partner=None):
        amt = _r(amt)
        if amt > 0.005:
            self.lines.append((acc, amt, 0.0, memo, partner))
        elif amt < -0.005:
            self.lines.append((acc, 0.0, -amt, memo, partner))

    def cr(self, acc, amt, memo="", partner=None):
        self.dr(acc, -(amt or 0.0), memo, partner)


def _cash_code_map(conn):
    return E.cash_accounts(conn)


def _leaf_codes(conn):
    return {c for (c,) in conn.execute("SELECT account_code FROM chart_of_accounts WHERE COALESCE(is_group,0)=0")}


def _post_sales_invoice(conn, A, sid):
    h = conn.execute("SELECT invoice_date, customer_code, customer_name, vat_amount, total_amount, "
                     "COALESCE(invoice_no, invoice_id) FROM sales_invoices WHERE invoice_id=?", (sid,)).fetchone()
    if h is None:
        return []
    d, code, name, vat, total, no = h
    e = _Entry(d, str(no), f"فاتورة مبيعات {no} - {name}")
    e.dr(A["gl_acc_ar"], total, f"مديونية العميل {name}", code)
    e.cr(A["gl_acc_sales"], _r(total) - _r(vat), "إيراد المبيعات (بعد الخصم الخاص)")
    e.cr(A["gl_acc_vat_out"], vat, "ضريبة المخرجات 15%")
    linked, direct = conn.execute(
        "SELECT COALESCE(SUM(CASE WHEN delivery_line_id IS NOT NULL THEN quantity*unit_cost END),0), "
        "COALESCE(SUM(CASE WHEN delivery_line_id IS NULL THEN quantity*unit_cost END),0) "
        "FROM sales_invoice_lines WHERE invoice_id=?", (sid,)).fetchone()
    e.dr(A["gl_acc_cogs"], _r(linked) + _r(direct), "تكلفة البضاعة المباعة")
    e.cr(A["gl_acc_gdni"], linked, "تحويل تكلفة البضاعة المسلّمة")
    e.cr(A["gl_acc_stock"], direct, "صرف من المخزون (بنود مباشرة)")
    return [e]


def _post_delivery(conn, A, did):
    h = conn.execute("SELECT delivery_date, customer_name FROM sales_deliveries WHERE delivery_id=?", (did,)).fetchone()
    if h is None:
        return []
    (cost,) = conn.execute("SELECT COALESCE(SUM(quantity*unit_cost),0) FROM sales_delivery_lines WHERE delivery_id=?",
                           (did,)).fetchone()
    e = _Entry(h[0], f"تسليم-{did}", f"إذن تسليم {did} - {h[1]}")
    e.dr(A["gl_acc_gdni"], cost, "بضاعة مسلّمة للعميل لم تُفوتر (بالتكلفة)")
    e.cr(A["gl_acc_stock"], cost, "صرف من المخزون")
    return [e]


def _post_receipt(conn, A, rid):
    h = conn.execute("SELECT receipt_date, supplier_name FROM purchase_receipts WHERE receipt_id=?", (rid,)).fetchone()
    if h is None:
        return []
    (cost,) = conn.execute("SELECT COALESCE(SUM(quantity*unit_cost),0) FROM purchase_receipt_lines WHERE receipt_id=?",
                           (rid,)).fetchone()
    e = _Entry(h[0], f"استلام-{rid}", f"إذن استلام {rid} - {h[1]}")
    e.dr(A["gl_acc_stock"], cost, "إضافة للمخزون (بسعر أمر الشراء)")
    e.cr(A["gl_acc_grni"], cost, "بضاعة مستلمة لم تصل فاتورتها")
    return [e]


def _post_purchase_invoice(conn, A, pid):
    h = conn.execute("SELECT purchase_date, supplier_code, supplier_name, vat_amount, total_amount, "
                     "COALESCE(invoice_no, purchase_id), COALESCE(source_type,'') FROM purchase_invoices WHERE purchase_id=?",
                     (pid,)).fetchone()
    if h is None:
        return []
    d, code, name, vat, total, no, src = h
    (linked,) = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM purchase_invoice_lines WHERE purchase_id=? "
                             "AND receipt_line_id IS NOT NULL", (pid,)).fetchone()
    net = _r(total) - _r(vat)
    e = _Entry(d, str(no), f"فاتورة مشتريات {no} - {name}")
    e.dr(A["gl_acc_grni"], linked, "تسوية البضاعة المستلمة")
    if src == "HansaWorld":
        # فاتورة تاريخية مستوردة بالإجمالي فقط (بلا أصناف): المخزون الفعلي أُدخل كرصيد افتتاحي من ملف التكلفة،
        # فتُحمَّل المشتريات على تكلفة المبيعات حتى لا يتضخم حساب المخزون بمشتريات بيعت بالفعل.
        e.dr(A["gl_acc_cogs"], net - _r(linked), "مشتريات تاريخية مستوردة (HansaWorld)")
    else:
        e.dr(A["gl_acc_stock"], net - _r(linked), "مشتريات مباشرة للمخزون")
    e.dr(A["gl_acc_vat_in"], vat, "ضريبة المدخلات 15%")
    e.cr(A["gl_acc_ap"], total, f"مستحقات المورد {name}", code)
    return [e]


def _post_return(conn, A, rid):
    h = conn.execute("SELECT return_date, return_type, party_code, party_name, item_name, quantity, unit_cost, "
                     "vat_amount, total_amount FROM returns_invoices WHERE return_id=?", (rid,)).fetchone()
    if h is None:
        return []
    d, rtype, code, name, item, qty, ucost, vat, total = h
    cost = _r((qty or 0) * (ucost or 0))
    net = _r(total) - _r(vat)
    e = _Entry(d, f"مرتجع-{rid}", f"{rtype} {rid} - {name} - {item}")
    if rtype == SALES_RETURN:
        e.dr(A["gl_acc_sales_ret"], net, "مردودات المبيعات")
        e.dr(A["gl_acc_vat_out"], vat, "عكس ضريبة المخرجات")
        e.cr(A["gl_acc_ar"], total, f"تخفيض مديونية العميل {name}", code)
        e.dr(A["gl_acc_stock"], cost, "إعادة البضاعة للمخزون بالتكلفة")
        e.cr(A["gl_acc_cogs"], cost, "عكس تكلفة المبيعات")
    else:
        e.dr(A["gl_acc_ap"], total, f"تخفيض مستحقات المورد {name}", code)
        e.cr(A["gl_acc_stock"], cost, "إخراج من المخزون بمتوسط التكلفة")
        e.cr(A["gl_acc_vat_in"], vat, "عكس ضريبة المدخلات")
        e.cr(A["gl_acc_pret_diff"], net - cost, "فرق سعر الإرجاع عن متوسط التكلفة")
    return [e]


def _partner_side(conn, code, alloc, ttype, tid=None):
    r = conn.execute("SELECT is_customer, is_supplier FROM business_partners WHERE partner_code=?", (code,)).fetchone()
    if not alloc and tid is not None:   # سند موزّع على عدة فواتير: نوع فواتيره يحدد جهة الذمم
        k = conn.execute("SELECT kind FROM advance_allocations WHERE trans_id=? ORDER BY alloc_id LIMIT 1",
                         (tid,)).fetchone()
        alloc = k[0] if k else alloc
    if alloc == "sales":
        return "c"
    if alloc == "purchase":
        return "s"
    if r is not None:
        c, s = bool(r[0]), bool(r[1])
        if c and not s:
            return "c"
        if s and not c:
            return "s"
    return "c" if ttype == "قبض" else "s"


def _expense_by_name(conn, name):
    """سند مصروفات ببند نصي (مثل «مواصلات ونقل») بدون كود: حساب مصروف فرعي بنفس الاسم إن وُجد."""
    target = E.norm_name(name or "")
    if not target:
        return None
    for code, n in conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE account_type=? "
                                "AND COALESCE(is_group,0)=0", (TYPE_EXP,)):
        if E.norm_name(n) == target:
            return code
    return None


def _post_voucher(conn, A, tid, cash_map=None, leaves=None):
    h = conn.execute("SELECT trans_type, trans_date, account_bank, category, party_code, party_name, amount, "
                     "COALESCE(alloc_type,''), COALESCE(description,'') FROM cash_transactions WHERE trans_id=?",
                     (tid,)).fetchone()
    if h is None:
        return []
    ttype, d, bank, cat, pcode, pname, amt, alloc, desc = h
    cash_map = cash_map if cash_map is not None else _cash_code_map(conn)
    leaves = leaves if leaves is not None else _leaf_codes(conn)
    cash = cash_map.get(bank)
    if cash not in leaves:
        cash = A["gl_acc_suspense"]
    partner = None
    if cat == CAT_PARTNER and pcode:
        side = _partner_side(conn, pcode, alloc, ttype, tid)
        counter, partner = (A["gl_acc_ar"], pcode) if side == "c" else (A["gl_acc_ap"], pcode)
    elif pcode in leaves:
        counter = pcode
    elif cat == E.CAT_EXPENSE:   # مصروف ببند نصي: حساب بنفس الاسم وإلا حساب المصروفات العمومية
        counter = _expense_by_name(conn, pname) or A["gl_acc_exp_general"]
    else:
        counter = A["gl_acc_suspense"]
    memo = f"{pname}" + (f" - {desc}" if desc else "")
    e = _Entry(d, f"سند-{tid}", f"سند {ttype} {tid} - {pname}")
    if ttype == "قبض":
        e.dr(cash, amt, f"وارد إلى {bank}")
        e.cr(counter, amt, memo, partner)
    else:
        e.dr(counter, amt, memo, partner)
        e.cr(cash, amt, f"منصرف من {bank}")
    return [e]


def _opening_date(conn):
    return E.system_start_date(conn) or today()


def _post_partner_opening(conn, A, pid):
    h = conn.execute("SELECT partner_code, name, is_customer, is_supplier, COALESCE(opening_debit,0), "
                     "COALESCE(opening_credit,0), COALESCE(opening_date,'') FROM business_partners WHERE partner_id=?",
                     (pid,)).fetchone()
    if h is None:
        return []
    code, name, is_c, is_s, od, oc, odate = h
    net = _r(od) - _r(oc)
    if abs(net) < 0.005:
        return []
    acc = A["gl_acc_ap"] if (is_s and not is_c) else A["gl_acc_ar"]
    e = _Entry(odate or _opening_date(conn), f"افتتاحي-{code}", f"رصيد افتتاحي - {name} ({code})")
    e.dr(acc, net, f"رصيد أول المدة {name}", code)
    e.cr(A["gl_acc_opening"], net, "وسيط افتتاحي")
    return [e]


def _post_item_opening(conn, A, rid):
    h = conn.execute("SELECT item_code, item_name, COALESCE(opening_qty,0), COALESCE(opening_cost,0) FROM inventory "
                     "WHERE rowid=?", (rid,)).fetchone()
    if h is None:
        return []
    val = _r(h[2] * h[3])
    if abs(val) < 0.005:
        return []
    e = _Entry(_opening_date(conn), f"افتتاحي-{h[0]}", f"رصيد افتتاحي مخزون - {h[1]}")
    e.dr(A["gl_acc_stock"], val, f"{h[2]:g} × {h[3]:g}")
    e.cr(A["gl_acc_opening"], val, "وسيط افتتاحي")
    return [e]


def _month_start_after(d, months):
    y, m = d.year, d.month + months
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return date(y, m, 1)


def _post_asset(conn, A, aid, as_of=None):
    """قيد إهلاك لكل شهر مستحق حتى اليوم. الإهلاك يبدأ من الشهر التالي لشهر الشراء (نفس قاعدة قائمة الدخل)."""
    h = conn.execute("SELECT asset_code, asset_name, purchase_date, purchase_cost, COALESCE(salvage_value,0), "
                     "useful_life_years FROM fixed_assets WHERE asset_id=?", (aid,)).fetchone()
    if h is None:
        return []
    code, name, p_date, cost, salvage, life = h
    try:
        p = datetime.strptime(p_date, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return []
    as_of = as_of or date.today()
    out = []
    prev = 0.0
    for i in range(1, int(life) * 12 + 1):
        d = _month_start_after(p, i)
        if d > as_of:
            break
        acc = round(E.accumulated_depreciation(cost, salvage, life, p_date, d), 2)
        amt = acc - prev
        prev = acc
        e = _Entry(d.isoformat(), f"{code}/{i}", f"إهلاك {name} ({code}) - شهر {d:%Y-%m}")
        e.dr(A["gl_acc_dep_exp"], amt, "مصروف الإهلاك الشهري")
        e.cr(A["gl_acc_dep_acc"], amt, "مجمع الإهلاك")
        out.append(e)
    return out


_POSTERS = {
    "sales_invoice": _post_sales_invoice, "purchase_invoice": _post_purchase_invoice, "delivery": _post_delivery,
    "receipt": _post_receipt, "return": _post_return, "voucher": _post_voucher,
    "partner_opening": _post_partner_opening, "item_opening": _post_item_opening, "asset": _post_asset,
}
_SOURCES_SQL = {
    "sales_invoice": "SELECT invoice_id FROM sales_invoices", "purchase_invoice": "SELECT purchase_id FROM purchase_invoices",
    "delivery": "SELECT delivery_id FROM sales_deliveries", "receipt": "SELECT receipt_id FROM purchase_receipts",
    "return": "SELECT return_id FROM returns_invoices", "voucher": "SELECT trans_id FROM cash_transactions",
    "partner_opening": "SELECT partner_id FROM business_partners WHERE COALESCE(opening_debit,0)!=0 OR COALESCE(opening_credit,0)!=0",
    "item_opening": "SELECT rowid FROM inventory WHERE COALESCE(opening_qty,0)*COALESCE(opening_cost,0) != 0",
    "asset": "SELECT asset_id FROM fixed_assets",
}


def _write(conn, src, sid, entries):
    conn.execute("DELETE FROM gl_lines WHERE gl_id IN (SELECT gl_id FROM gl_entries WHERE source_type=? AND source_id=?)",
                 (src, sid))
    conn.execute("DELETE FROM gl_entries WHERE source_type=? AND source_id=?", (src, sid))
    for e in entries:
        if not e.lines:
            continue
        dr = round(sum(x[1] for x in e.lines), 2)
        cr = round(sum(x[2] for x in e.lines), 2)
        if abs(dr - cr) > 0.005:   # حماية: أي فرق تقريب يذهب للحساب المعلق حتى يبقى الدفتر متوازناً دائماً
            e.lines.append((_ensure_account(conn, "gl_acc_suspense"), max(cr - dr, 0.0), max(dr - cr, 0.0),
                            "فرق تقريب", None))
        gid = conn.execute("INSERT INTO gl_entries (entry_date, source_type, source_id, ref, description) VALUES (?,?,?,?,?)",
                           (e.date, src, sid, e.ref, e.desc)).lastrowid
        conn.executemany("INSERT INTO gl_lines (gl_id, account_code, debit, credit, memo, partner_code) VALUES (?,?,?,?,?,?)",
                         [(gid, a, round(d, 2), round(c, 2), m, p) for a, d, c, m, p in e.lines])


def post_document(conn, src, sid, A=None):
    A = A or gl_accounts(conn)
    _write(conn, src, sid, _POSTERS[src](conn, A, sid))


def _rebuild(conn):
    A = gl_accounts(conn)
    conn.execute("DELETE FROM gl_lines")
    conn.execute("DELETE FROM gl_entries")
    cash_map, leaves = _cash_code_map(conn), _leaf_codes(conn)
    for src, sql in _SOURCES_SQL.items():
        for (sid,) in conn.execute(sql).fetchall():
            if src == "voucher":
                entries = _post_voucher(conn, A, sid, cash_map, leaves)
            else:
                entries = _POSTERS[src](conn, A, sid)
            _write(conn, src, sid, entries)
    conn.execute("DELETE FROM gl_queue")
    conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_rebuilt_at', ?)",
                 (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
    conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_dep_month', ?)", (date.today().strftime("%Y-%m"),))
    conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_open_date', ?)", (_opening_date(conn),))


def process_gl_queue(conn):
    """يرحّل المستندات المعلقة في الطابور. يُستدعى تلقائياً قبل حفظ أي معاملة (COMMIT_HOOKS).
    أي خطأ في الترحيل لا يُفشل حفظ المستند: يُتراجع عن الترحيل فقط ويبقى المستند في الطابور لإعادة المحاولة."""
    try:
        pending = conn.execute("SELECT source_type, source_id FROM gl_queue").fetchall()
    except sqlite3.OperationalError:   # قاعدة قديمة قبل إنشاء جداول الدفتر العام
        return
    if not pending:
        return
    conn.execute("SAVEPOINT gl_post")
    try:
        if any(s == "all" for s, _ in pending):
            _rebuild(conn)
        else:
            A = gl_accounts(conn)
            for src, sid in pending:
                if src in _POSTERS:
                    post_document(conn, src, sid, A)
            conn.executemany("DELETE FROM gl_queue WHERE source_type=? AND source_id=?", pending)
        conn.execute("DELETE FROM app_settings WHERE key='gl_last_error'")   # الترحيل نجح: لا يوجد خطأ معلق
        conn.execute("RELEASE SAVEPOINT gl_post")
    except Exception as exc:  # noqa: BLE001
        conn.execute("ROLLBACK TO SAVEPOINT gl_post")
        conn.execute("RELEASE SAVEPOINT gl_post")
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_last_error', ?)",
                     (f"{datetime.now():%Y-%m-%d %H:%M} {type(exc).__name__}: {exc}",))


E.COMMIT_HOOKS.append(process_gl_queue)
E.INIT_HOOKS.append(setup_ledger)


def sync_ledger():
    """قبل أي تقرير من الدفتر العام: ترحيل المعلق + إهلاك الشهر الجديد + تاريخ الأرصدة الافتتاحية إن تغيّر."""
    with db() as conn:
        try:
            dep_m = conn.execute("SELECT value FROM app_settings WHERE key='gl_dep_month'").fetchone()
            open_d = conn.execute("SELECT value FROM app_settings WHERE key='gl_open_date'").fetchone()
        except sqlite3.OperationalError:
            return
        cur_m = date.today().strftime("%Y-%m")
        if not dep_m or dep_m[0] != cur_m:
            conn.execute("INSERT INTO gl_queue SELECT 'asset', asset_id FROM fixed_assets WHERE 1 "
                         "ON CONFLICT DO NOTHING")
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_dep_month', ?)", (cur_m,))
        od = _opening_date(conn)
        if not open_d or open_d[0] != od:
            conn.execute("INSERT INTO gl_queue SELECT 'item_opening', rowid FROM inventory WHERE 1 ON CONFLICT DO NOTHING")
            conn.execute("INSERT INTO gl_queue SELECT 'partner_opening', partner_id FROM business_partners "
                         "WHERE COALESCE(opening_date,'')='' ON CONFLICT DO NOTHING")
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('gl_open_date', ?)", (od,))


def rebuild_ledger():
    """إعادة ترحيل كل المستندات من الصفر (آمنة؛ القيود الآلية مشتقة من المستندات)."""
    with db() as conn:
        _rebuild(conn)
    return ledger_status()


def ledger_status():
    with db() as conn:
        g = lambda k: (conn.execute("SELECT value FROM app_settings WHERE key=?", (k,)).fetchone() or [""])[0]
        n_auto, = conn.execute("SELECT COUNT(*) FROM gl_entries").fetchone()
        n_man, = conn.execute("SELECT COUNT(*) FROM journal_entries").fetchone()
        pend, = conn.execute("SELECT COUNT(*) FROM gl_queue").fetchone()
        by_src = conn.execute("SELECT source_type, COUNT(*), COALESCE(SUM(l.d),0) FROM gl_entries e JOIN "
                              "(SELECT gl_id, SUM(debit) d FROM gl_lines GROUP BY gl_id) l ON l.gl_id=e.gl_id "
                              "GROUP BY source_type ORDER BY 2 DESC").fetchall()
        dr, cr = conn.execute("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(credit),0) FROM gl_lines").fetchone()
        jdr, jcr = conn.execute("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(credit),0) FROM journal_lines").fetchone()
        unbal = conn.execute("SELECT COUNT(*) FROM (SELECT gl_id FROM gl_lines GROUP BY gl_id "
                             "HAVING ABS(SUM(debit)-SUM(credit)) > 0.005)").fetchone()[0]
        susp = gl_accounts(conn)["gl_acc_suspense"]
        sbal, = conn.execute("SELECT COALESCE(SUM(debit-credit),0) FROM gl_lines WHERE account_code=?", (susp,)).fetchone()
        rebuilt_at, last_error = g("gl_rebuilt_at"), g("gl_last_error")
    return dict(auto_entries=n_auto, manual_entries=n_man, pending=pend, rebuilt_at=rebuilt_at,
                last_error=last_error, by_source=[(SRC_LABEL.get(s, s), n, v) for s, n, v in by_src],
                debit=dr + jdr, credit=cr + jcr, balanced=abs((dr + jdr) - (cr + jcr)) < 0.01 and unbal == 0,
                suspense_code=susp, suspense_balance=sbal)


# ============================================================
# القراءة من الدفتر العام الكامل
# ============================================================
_GL_UNION = ("SELECT e.entry_date AS d, l.account_code AS acc, l.debit AS dr, l.credit AS cr, "
             "'قيد-' || e.entry_id AS ref, COALESCE(NULLIF(l.memo,''), e.description, '') AS memo, "
             "CASE WHEN e.source_type IS NULL THEN 'journal' ELSE e.source_type END AS src, NULL AS partner "
             "FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id "
             "UNION ALL "
             "SELECT g.entry_date, l.account_code, l.debit, l.credit, g.ref, "
             "COALESCE(NULLIF(l.memo,''), g.description, ''), g.source_type, l.partner_code "
             "FROM gl_lines l JOIN gl_entries g ON g.gl_id=l.gl_id")


def gl_balances(d_from, d_to):
    """{الكود: [أول المدة، مدين الفترة، دائن الفترة]} من الدفتر العام الكامل (يدوي + آلي)."""
    sync_ledger()
    with db() as conn:
        info, _ = _account_maps(conn)
        data = defaultdict(lambda: [0.0, 0.0, 0.0])
        for acc, before, dr, cr in conn.execute(
                f"SELECT acc, SUM(CASE WHEN d < ? THEN dr - cr ELSE 0 END), SUM(CASE WHEN d >= ? THEN dr ELSE 0 END), "
                f"SUM(CASE WHEN d >= ? THEN cr ELSE 0 END) FROM ({_GL_UNION}) WHERE d <= ? GROUP BY acc",
                (d_from, d_from, d_from, d_to)):
            data[acc] = [before or 0.0, dr or 0.0, cr or 0.0]
    return info, data


def gl_account_statement(code, d_from, d_to):
    """كشف حساب من الدفتر العام الكامل (يشمل كل فروع الحساب لو كان تجميعياً)."""
    sync_ledger()
    with db() as conn:
        info, _ = _account_maps(conn)
        if code not in info:
            raise UserError("الحساب غير موجود.")
        codes = set(E.account_descendants(conn, code))
        ph = ",".join("?" * len(codes))
        moves = conn.execute(f"SELECT d, ref, acc, memo, src, dr, cr FROM ({_GL_UNION}) WHERE acc IN ({ph}) "
                             f"ORDER BY d, ref", list(codes)).fetchall()
    opening = sum(m[5] - m[6] for m in moves if m[0] < d_from)
    rows, bal, tdr, tcr = [], opening, 0.0, 0.0
    for d, ref, acc, memo, src, dr, cr in moves:
        if d < d_from or d > d_to:
            continue
        bal += dr - cr
        tdr += dr
        tcr += cr
        rows.append((d, ref, SRC_LABEL.get(src, src), f"{acc} - {info.get(acc, {}).get('name', acc)}", memo, dr, cr, bal))
    return dict(code=code, name=info[code]["name"], opening=opening, debit=tdr, credit=tcr, closing=bal, rows=rows)


def _p_from_to(p):
    return p.get("d_from") or "0000-01-01", p.get("d_to") or today()


def rep_gl_trial_balance(p):
    a, b = _p_from_to(p)
    info, data = gl_balances(a, b)
    rows, _ = _tree_report_rows(info, data, two_sided=True)
    pos = lambda x: x if x > 0.005 else 0.0
    leaves = [(pos(o), pos(-o), dr, cr, pos(o + dr - cr), pos(-(o + dr - cr))) for o, dr, cr in data.values()]
    sums = [sum(r[i] for r in leaves) for i in range(6)]
    ok = abs(sums[2] - sums[3]) < 0.01 and abs(sums[4] - sums[5]) < 0.01
    cols = [("رقم الحساب", "t"), ("اسم الحساب", "t"), ("المستوى / النوع", "t"), ("أول المدة مدين", "m"), ("أول المدة دائن", "m"),
            ("حركة مدين", "m"), ("حركة دائن", "m"), ("آخر المدة مدين", "m"), ("آخر المدة دائن", "m")]
    notes = ["الدفتر العام الكامل: القيود اليدوية + الترحيل الآلي لكل المستندات (فواتير، أذون، سندات، مرتجعات، خصومات، "
             "تكلفة المبيعات، أرصدة افتتاحية، إهلاك).",
             "الميزان متوازن ✓" if ok else "تنبيه: الميزان غير متوازن!"]
    return Report("ميزان المراجعة الشامل (الدفتر العام)", cols, rows, ("", "الإجمالي (الحسابات الفرعية)", "", *sums),
                  notes, f"من {a} إلى {b}")


def _contra_prefixes():
    v = E.get_setting(CONTRA_ASSET_SETTING, CONTRA_ASSET_DEFAULT)
    return tuple(x.strip() for x in v.split(",") if x.strip())


def balance_sheet(as_of):
    """الميزانية العمومية من الدفتر العام: {أصول، خصوم، حقوق ملكية، ربح الفترة}، كل قسم [(الكود، الاسم، الرصيد)]."""
    info, data = gl_balances("0000-01-01", as_of)
    contra = _contra_prefixes()
    sec = {TYPE_ASSET: [], TYPE_LIAB: [], TYPE_EQUITY: []}
    pl = 0.0
    for acc, (o, dr, cr) in sorted(data.items()):
        bal = o + dr - cr                     # مدين موجب
        if abs(bal) < 0.005:
            continue
        typ = info.get(acc, {}).get("type", TYPE_LIAB)
        name = info.get(acc, {}).get("name", f"{acc} (غير موجود بالدليل)")
        if typ in (TYPE_REV, TYPE_EXP):
            pl += -bal
        elif typ == TYPE_ASSET or (contra and acc.startswith(contra)):
            sec[TYPE_ASSET].append((acc, name, bal))
        elif typ == TYPE_LIAB:
            sec[TYPE_LIAB].append((acc, name, -bal))
        else:
            sec[TYPE_EQUITY].append((acc, name, -bal))
    tot = {k: sum(r[2] for r in v) for k, v in sec.items()}
    return dict(sections=sec, totals=tot, profit=pl, as_of=as_of,
                balanced=abs(tot[TYPE_ASSET] - (tot[TYPE_LIAB] + tot[TYPE_EQUITY] + pl)) < 0.01)


def rep_balance_sheet(p):
    _, b = _p_from_to(p)
    bs = balance_sheet(b)
    info, _ = gl_balances("0000-01-01", b)
    rows = []

    def parent_name(code):
        par = info.get(code, {}).get("parent")
        return info.get(par, {}).get("name", "") if par else ""

    for typ, title in ((TYPE_ASSET, "الأصول"), (TYPE_LIAB, "الخصوم (الالتزامات)"), (TYPE_EQUITY, "حقوق الملكية")):
        rows.append(("", f"▸ {title}", "", None))
        for code, name, val in bs["sections"][typ]:
            rows.append((code, "    " + name, parent_name(code), val))
        if typ == TYPE_EQUITY:
            rows.append(("", "    صافي ربح/خسارة الفترات (من الدفتر العام، لم يُقفل بعد)", "", bs["profit"]))
        total = bs["totals"][typ] + (bs["profit"] if typ == TYPE_EQUITY else 0.0)
        rows.append(("", f"إجمالي {title}", "", total))
    le = bs["totals"][TYPE_LIAB] + bs["totals"][TYPE_EQUITY] + bs["profit"]
    rows.append(("", "إجمالي الخصوم وحقوق الملكية", "", le))
    cols = [("الكود", "t"), ("البند", "t"), ("الحساب الرئيسي", "t"), ("الرصيد", "m")]
    notes = [("الميزانية متوازنة ✓ (الأصول = الخصوم + حقوق الملكية)" if bs["balanced"] else
              "تنبيه: الميزانية غير متوازنة! راجع «حالة الترحيل» و«مطابقة الدفتر العام»."),
             f"حسابات تبدأ بـ ({', '.join(_contra_prefixes()) or '-'}) تُعرض مخصومة من الأصول (مجمعات الإهلاك) "
             "حتى لو صُنفت خصوماً في الدليل. الرصيد السالب في الأصول = حساب مقابل (Contra).",
             "الأرصدة من الدفتر العام الكامل: القيود اليدوية + الترحيل الآلي لكل المستندات."]
    return Report("الميزانية العمومية (المركز المالي)", cols, rows, None, notes, f"في {b}")


def rep_gl_income(p):
    """قائمة الدخل من الدفتر العام (بالحسابات) — للمطابقة مع قائمة الدخل التحليلية."""
    a, b = _p_from_to(p)
    info, data = gl_balances(a, b)
    rev, exp = [], []
    for acc, (_o, dr, cr) in sorted(data.items()):
        typ = info.get(acc, {}).get("type")
        if typ == TYPE_REV and abs(cr - dr) > 0.005:
            rev.append((acc, info[acc]["name"], cr - dr))
        elif typ == TYPE_EXP and abs(dr - cr) > 0.005:
            exp.append((acc, info[acc]["name"], dr - cr))
    tr, te = sum(r[2] for r in rev), sum(r[2] for r in exp)
    rows = [("", "▸ الإيرادات (دائن - مدين)", None)] + [(c, "    " + n, v) for c, n, v in rev] + \
           [("", "إجمالي الإيرادات", tr), ("", "▸ المصروفات والتكاليف (مدين - دائن)", None)] + \
           [(c, "    " + n, v) for c, n, v in exp] + [("", "إجمالي المصروفات والتكاليف", te), ("", "صافي الربح / (الخسارة)", tr - te)]
    notes = ["محسوبة من الدفتر العام. يجب أن يطابق صافي الربح قائمة الدخل التحليلية للفترة نفسها (راجع تقرير المطابقة)."]
    return Report("قائمة الدخل من الدفتر العام", [("الكود", "t"), ("الحساب", "t"), ("المبلغ", "m")], rows, None, notes,
                  f"من {a} إلى {b}")


def rep_gl_journal(p):
    """دفتر اليومية العام: كل القيود (اليدوية والآلية) في الفترة، سطراً بسطر."""
    a, b = _p_from_to(p)
    sync_ledger()
    with db() as conn:
        info, _ = _account_maps(conn)
        rows = conn.execute(f"SELECT d, ref, src, acc, memo, dr, cr FROM ({_GL_UNION}) WHERE d BETWEEN ? AND ? "
                            f"ORDER BY d, ref", (a, b)).fetchall()
    out = [(d, ref, SRC_LABEL.get(src, src), acc, info.get(acc, {}).get("name", acc), memo, dr or 0.0, cr or 0.0)
           for d, ref, src, acc, memo, dr, cr in rows]
    cols = [("التاريخ", "t"), ("المرجع", "t"), ("المصدر", "t"), ("الحساب", "t"), ("اسم الحساب", "t"), ("البيان", "t"),
            ("مدين", "m"), ("دائن", "m")]
    total = ("الإجمالي", "", "", "", "", "", sum(r[6] for r in out), sum(r[7] for r in out))
    return Report("دفتر اليومية العام (يدوي + آلي)", cols, out, total,
                  ["كل مستند يُرحَّل تلقائياً لحظة حفظه أو تعديله أو حذفه."], f"من {a} إلى {b}")


def ledger_reconciliation(as_of):
    """مطابقة حسابات المراقبة في الدفتر العام مع الدفاتر المساعدة. [(البند، الدفتر العام، الدفتر المساعد، الفرق)]"""
    info, data = gl_balances("0000-01-01", as_of)
    bal = lambda c: sum(v[0] + v[1] - v[2] for k, v in data.items() if k == c)
    with db() as conn:
        A = gl_accounts(conn)
        (stock,) = conn.execute("SELECT COALESCE(SUM(quantity*cost_price),0) FROM inventory").fetchone()
        cash = sum(E.cash_balances(conn, as_of).values())
        cash_codes = set(E.cash_accounts(conn).values())
        (gdni_sub,) = conn.execute("SELECT COALESCE(SUM(l.quantity*l.unit_cost),0) FROM sales_delivery_lines l "
                                   "JOIN sales_deliveries d ON d.delivery_id=l.delivery_id "
                                   "WHERE d.invoice_id IS NULL AND d.delivery_date<=?", (as_of,)).fetchone()
        (grni_sub,) = conn.execute("SELECT COALESCE(SUM(l.quantity*l.unit_cost),0) FROM purchase_receipt_lines l "
                                   "JOIN purchase_receipts r ON r.receipt_id=l.receipt_id "
                                   "WHERE r.invoice_id IS NULL AND r.receipt_date<=?", (as_of,)).fetchone()
    partners = sum(r["balance"] for r in E.partner_balances(as_of))
    gl_cash = sum(v[0] + v[1] - v[2] for k, v in data.items() if k in cash_codes)
    out = [
        ("العملاء والموردون (صافي: لنا - علينا)", bal(A["gl_acc_ar"]) + bal(A["gl_acc_ap"]), partners),
        ("المخزون (الدفتر العام مقابل تقييم المخزون الحالي)", bal(A["gl_acc_stock"]), stock),
        ("النقدية والبنوك", gl_cash, cash),
        ("بضاعة مسلّمة لم تُفوتر", bal(A["gl_acc_gdni"]), gdni_sub),
        ("بضاعة مستلمة لم تصل فاتورتها", -bal(A["gl_acc_grni"]), grni_sub),
    ]
    try:
        ga, gb = f"{as_of[:4]}-01-01", as_of
        gl_pl = rep_gl_income(dict(d_from=ga, d_to=gb)).rows[-1][2]
        out.append((f"صافي ربح السنة حتى {as_of} (الدفتر العام مقابل قائمة الدخل)", gl_pl, _income_net_profit(ga, gb)))
    except Exception:  # noqa: BLE001
        pass
    out.append(("الحساب المعلق (يجب أن يكون صفراً)", bal(A["gl_acc_suspense"]), 0.0))
    return [(n, g, s, g - s) for n, g, s in out]


def _income_net_profit(d_from, d_to):
    """صافي الربح من قائمة الدخل التحليلية (نفس معادلة build_income_report)."""
    txt = E.build_income_report(d_from, d_to)
    for line in txt.splitlines():
        if line.strip().startswith("صافي الربح"):
            num = line.split("صافي الربح", 1)[1].replace("ر.س", "").replace(",", "").strip()
            return float(num)
    return 0.0


def rep_gl_reconcile(p):
    _, b = _p_from_to(p)
    rows = ledger_reconciliation(b)
    notes = ["الفرق الصفري = الدفتر العام يطابق الدفاتر المساعدة. فروق تقريب صغيرة في المخزون طبيعية (متوسط مرجح بأربعة أرقام).",
             "فرق صافي الربح ينشأ فقط من قيود يدوية على حسابات يحسبها النظام آلياً (تكلفة المبيعات، مردودات المشتريات، الإهلاك).",
             "فرق النقدية ينشأ من سندات على خزينة غير معرّفة بالدليل (تُرحَّل للحساب المعلق)."]
    return Report("مطابقة الدفتر العام مع الدفاتر المساعدة", [("البند", "t"), ("الدفتر العام", "m"), ("الدفتر المساعد", "m"),
                                                              ("الفرق", "m")], rows, None, notes, f"في {b}")


register_report_group("ledger", "📚 الدفتر العام والقوائم المالية", [
    ("gl_tb", "ميزان المراجعة الشامل (الدفتر العام)", rep_gl_trial_balance,
     "كل المستندات مرحّلة آلياً + القيود اليدوية، مع تجميع الحسابات الفرعية تحت الرئيسية."),
    ("bs", "الميزانية العمومية (المركز المالي)", rep_balance_sheet,
     "الأصول والخصوم وحقوق الملكية في تاريخ (إلى) من الدفتر العام، مع فحص التوازن."),
    ("gl_pl", "قائمة الدخل من الدفتر العام", rep_gl_income, "الإيرادات والمصروفات بالحسابات من الدفتر العام للفترة."),
    ("gl_journal", "دفتر اليومية العام (يدوي + آلي)", rep_gl_journal, "كل أسطر القيود في الفترة بمصدرها ومرجعها."),
    ("gl_recon", "مطابقة الدفتر العام مع الدفاتر المساعدة", rep_gl_reconcile,
     "العملاء/الموردون، المخزون، النقدية، الوسطاء، وصافي الربح: الدفتر العام مقابل المصدر."),
])

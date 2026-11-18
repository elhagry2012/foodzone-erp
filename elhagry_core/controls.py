# -*- coding: utf-8 -*-
"""
الرقابة والدورة المستندية المتقدمة (المرحلة الرابعة):
    🧑‍💼 الأدوار الوظيفية (مجموعات صلاحيات) + فصل «العرض» عن «الإنشاء/التنفيذ» في الأوامر والاستلام والتسليم
    📎 مرفقات فاتورة المورد على «استلام المخزن» + كل مرفقات المورد في مكان واحد
    🔗 تسوية الدفعات المقدمة / تحت الحساب مع الفواتير اللاحقة (مطابقة فقط بلا قيد جديد)
    📑 تقرير الفواتير المفتوحة (غير المسددة والمسددة جزئياً) للعملاء والموردين

ملاحظة محاسبية: الدفعة المقدمة مسجلة أصلاً كحركة على حساب الطرف في الذمم (دائن للعميل / مدين للمورد)،
فتسويتها مع فاتورة لاحقة لا تغيّر أي رصيد ولا تحتاج قيداً؛ هي مطابقة تغلق الفاتورة وتستهلك الدفعة.
"""
from datetime import datetime

from . import engine as E
from .engine import (CAT_PARTNER, FLOW, REPORT_CATALOG, REPORT_FIN_REQUIRES, REPORT_GROUP_OF, REPORT_INDEX,
                     SUB_PERMISSIONS, Report, UserError, current_user_name, db, fmt, normalize_permissions,
                     parse_date, refresh_invoice_paid, require_perm, today)

# ============================================================
# 🧑‍💼 الأدوار والصلاحيات
# ============================================================
PERM_SPLIT_FLAG = "perm_split_create_v1"
# صلاحية العرض القديمة ← صلاحية الإنشاء/التنفيذ الجديدة (المستخدمون الحاليون لا يفقدون قدرتهم على الإنشاء)
PERM_SPLIT = {"purchase.orders": "purchase.orders_create", "purchase.receipt": "purchase.receipt_create",
              "sales.orders": "sales.orders_create", "sales.delivery": "sales.delivery_create"}


def _report_subs():
    return [k for k, _ in SUB_PERMISSIONS["reports"]]


def default_roles():
    """أدوار جاهزة تُنشأ مرة واحدة (يمكن تعديلها أو حذفها من شاشة المستخدمين)."""
    return [
        ("أمين المخزن", "استلام البضاعة وتسليمها وإرفاق فاتورة المورد — بدون أسعار شراء أو خزينة أو أرصدة",
         ["inventory", "expiry", "purchase", "purchase.receipt", "purchase.receipt_create",
          "sales", "sales.delivery", "sales.delivery_create"]),
        ("مسؤول المشتريات", "إنشاء أوامر الشراء ومتابعة الاستلام وفواتير الموردين",
         ["suppliers", "inventory", "purchase", "purchase.orders", "purchase.orders_create", "purchase.receipt",
          "purchase.invoices", "purchase.supplier_fin", "statement"]),
        ("أمين الخزينة", "سندات القبض والصرف وتسوية الدفعات المقدمة وكشوف الحساب",
         ["cash", "cash.receipt", "cash.payment", "cash.register", "cash.print", "statement"]),
        ("المحاسب", "الخزينة والقيود والتقارير المالية والكشوف — بدون إنشاء أوامر أو تنفيذ استلام",
         ["dashboard", "cash", "cash.receipt", "cash.payment", "cash.register", "cash.print", "statement", "income",
          "reports", *_report_subs(), "journal", "discounts", "ledger", "coa",
          "sales", "sales.invoices", "sales.discount", "sales.customer_fin",
          "purchase", "purchase.invoices", "purchase.discount", "purchase.supplier_fin"]),
    ]


def setup_controls(conn):
    # 1) ترحيل الصلاحيات: من يملك «العرض» تفصيلياً يحصل على «الإنشاء» المقابل (كما كان يعمل قبل الفصل)
    done = conn.execute("SELECT value FROM app_settings WHERE key=?", (PERM_SPLIT_FLAG,)).fetchone()
    if not done:
        for uid, perms in conn.execute("SELECT user_id, permissions FROM users WHERE is_admin=0").fetchall():
            ps = [p for p in (perms or "").split(",") if p]
            add = [new for old, new in PERM_SPLIT.items() if old in ps and new not in ps]
            if add:
                conn.execute("UPDATE users SET permissions=? WHERE user_id=?", (",".join(ps + add), uid))
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, '1')", (PERM_SPLIT_FLAG,))
    # 2) الأدوار الجاهزة (مرة واحدة)
    if not conn.execute("SELECT 1 FROM user_roles LIMIT 1").fetchone() and not conn.execute(
            "SELECT value FROM app_settings WHERE key='roles_seeded'").fetchone():
        for name, desc, perms in default_roles():
            conn.execute("INSERT OR IGNORE INTO user_roles (name, description, permissions) VALUES (?,?,?)",
                         (name, desc, ",".join(normalize_permissions(perms))))
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES ('roles_seeded', '1')")


    # 3) توريث المندوب للسندات القديمة التي ليس لها مندوب
    E.stamp_voucher_reps(conn)


E.INIT_HOOKS.append(setup_controls)


def suggest_partner_reps():
    """أطراف بطاقتها بلا مندوب بينما فواتيرها مسجلة بمندوب: اقتراح المندوب الأكثر تكراراً في فواتيرها
    (عند التساوي: مندوب أحدث فاتورة).
    [(الكود، الاسم، المندوب المقترح، عدد فواتيره، أول تاريخ فاتورة)]"""
    out = []
    with db() as conn:
        for kind in ("sales", "purchase"):
            c = FLOW[kind]
            col = "sales_rep" if kind == "sales" else "purchase_rep"
            for code, name, rep, n, first in conn.execute(
                    f"SELECT h.{c['pcode']}, b.name, h.{col}, COUNT(*), MIN(h.{c['idate']}) FROM {c['inv']} h "
                    f"JOIN business_partners b ON b.partner_code=h.{c['pcode']} "
                    f"WHERE COALESCE(b.rep_name,'')='' AND COALESCE(h.{col},'')!='' "
                    f"AND NOT EXISTS (SELECT 1 FROM partner_rep_history x WHERE x.partner_code=h.{c['pcode']}) "
                    f"GROUP BY h.{c['pcode']}, h.{col} ORDER BY h.{c['pcode']}, COUNT(*) DESC, MAX(h.{c['idate']}) DESC"):
                if not any(o[0] == code for o in out):
                    out.append((code, name, rep, n, first))
    return out


def apply_partner_reps(suggestions):
    """تسجيل المندوب المقترح في بطاقة الطرف اعتباراً من أول فاتورة له (فتُنسب له تحصيلاته السابقة واللاحقة)."""
    require_perm("reps", "ربط الأطراف بالمناديب")
    with db() as conn:
        for code, _name, rep, _n, first in suggestions:
            E.set_partner_rep(conn, code, rep, first, current_user_name())
    return len(suggestions)


def list_roles():
    """[(role_id, الاسم، الوصف، [الصلاحيات]، عدد المستخدمين)]"""
    with db() as conn:
        return [(rid, n, d or "", [p for p in (ps or "").split(",") if p], cnt) for rid, n, d, ps, cnt in conn.execute(
            "SELECT r.role_id, r.name, r.description, r.permissions, "
            "(SELECT COUNT(*) FROM users u WHERE u.role_id=r.role_id) FROM user_roles r ORDER BY r.name")]


def role_permissions(role_id):
    with db() as conn:
        r = conn.execute("SELECT permissions FROM user_roles WHERE role_id=?", (role_id,)).fetchone()
    if r is None:
        raise UserError("الدور غير موجود.")
    return [p for p in (r[0] or "").split(",") if p]


def save_role(role_id, name, perms, description=""):
    """إضافة/تعديل دور. التعديل يُطبَّق فوراً على كل مستخدمي الدور (غير المديرين). يرجع (role_id، عدد المستخدمين المحدَّثين)."""
    name = (name or "").strip()
    if not name:
        raise UserError("أدخل اسم الدور.")
    perms = normalize_permissions(perms)
    if not perms:
        raise UserError("اختر صلاحية واحدة على الأقل للدور.")
    with db() as conn:
        if conn.execute("SELECT 1 FROM user_roles WHERE name=? AND role_id != ?", (name, role_id or -1)).fetchone():
            raise UserError("يوجد دور بنفس الاسم.")
        if role_id:
            conn.execute("UPDATE user_roles SET name=?, description=?, permissions=? WHERE role_id=?",
                         (name, description, ",".join(perms), role_id))
        else:
            role_id = conn.execute("INSERT INTO user_roles (name, description, permissions) VALUES (?,?,?)",
                                   (name, description, ",".join(perms))).lastrowid
        n = conn.execute("UPDATE users SET permissions=? WHERE role_id=? AND is_admin=0",
                         (",".join(perms), role_id)).rowcount
    return role_id, n


def delete_role(role_id):
    """حذف الدور فقط: مستخدموه يحتفظون بصلاحياتهم الحالية كصلاحيات فردية."""
    with db() as conn:
        conn.execute("UPDATE users SET role_id=NULL WHERE role_id=?", (role_id,))
        conn.execute("DELETE FROM user_roles WHERE role_id=?", (role_id,))


def user_role_id(user_id):
    with db() as conn:
        r = conn.execute("SELECT role_id FROM users WHERE user_id=?", (user_id,)).fetchone()
    return r[0] if r else None


# ============================================================
# 📎 مرفقات فاتورة المورد على استلام المخزن
# ============================================================
def add_receipt_attachment(receipt_id, file_name, data, notes=""):
    """رفع صورة/PDF فاتورة المورد الورقية على إذن الاستلام (تظهر أيضاً على الفاتورة التي تصدر منه)."""
    require_perm("purchase.receipt_create", "إرفاق فاتورة المورد على الاستلام")
    return E.add_attachment("receipt", receipt_id, file_name, data, current_user_name(), notes)


def supplier_attachments(supplier_code):
    """كل مرفقات المورد: على أذون الاستلام وعلى فواتير المشتريات.
    [(attach_id، المستند، رقم الإذن، رقم الفاتورة، التاريخ، اسم الملف، الحجم، رفعه، وقت الرفع)]"""
    with db() as conn:
        rows = conn.execute(
            "SELECT a.attach_id, 'إذن استلام', r.receipt_id, "
            "COALESCE((SELECT COALESCE(invoice_no, purchase_id) FROM purchase_invoices WHERE purchase_id=r.invoice_id),'-'), "
            "r.receipt_date, a.file_name, a.size, COALESCE(a.uploaded_by,''), a.uploaded_at "
            "FROM attachments a JOIN purchase_receipts r ON a.doc_type='receipt' AND r.receipt_id=a.doc_id "
            "WHERE r.supplier_code=?", (supplier_code,)).fetchall()
        rows += conn.execute(
            "SELECT a.attach_id, 'فاتورة مشتريات', '-', COALESCE(p.invoice_no, p.purchase_id), p.purchase_date, a.file_name, "
            "a.size, COALESCE(a.uploaded_by,''), a.uploaded_at "
            "FROM attachments a JOIN purchase_invoices p ON a.doc_type='purchase' AND p.purchase_id=a.doc_id "
            "WHERE p.supplier_code=?", (supplier_code,)).fetchall()
    return sorted(rows, key=lambda r: (r[4] or "", r[0]), reverse=True)


# ============================================================
# 🔗 الدفعات المقدمة / تحت الحساب وتسويتها مع الفواتير
# ============================================================
def _voucher_type(kind):
    return "قبض" if kind == "sales" else "صرف"


def list_advances(party_code=None, kind=None, only_open=True):
    """سندات الأطراف غير المربوطة بفاتورة (دفعات مقدمة / تحت الحساب) والمتاح منها للتسوية.
    [dict(trans_id, date, kind, party_code, party_name, amount, allocated, available, description)]"""
    kinds = [kind] if kind else ["sales", "purchase"]
    out = []
    with db() as conn:
        for k in kinds:
            sql = ("SELECT t.trans_id, t.trans_date, t.party_code, t.party_name, t.amount, COALESCE(t.description,''), "
                   "COALESCE((SELECT SUM(a.amount) FROM advance_allocations a WHERE a.trans_id=t.trans_id),0) "
                   "FROM cash_transactions t WHERE t.category=? AND t.trans_type=? AND COALESCE(t.alloc_type,'')='' "
                   "AND t.party_code IS NOT NULL")
            args = [CAT_PARTNER, _voucher_type(k)]
            if party_code:
                sql += " AND t.party_code=?"
                args.append(party_code)
            for tid, d, pc, pn, amt, desc, used in conn.execute(sql + " ORDER BY t.trans_date, t.trans_id", args):
                avail = round(amt - used, 2)
                if only_open and avail <= 0.005:
                    continue
                out.append(dict(trans_id=tid, date=d, kind=k, party_code=pc, party_name=pn, amount=amt,
                                allocated=round(used, 2), available=max(avail, 0.0), description=desc))
    return out


def party_advance_available(party_code, kind):
    return round(sum(a["available"] for a in list_advances(party_code, kind)), 2)


def _invoice_remaining(conn, kind, inv_id):
    c = FLOW[kind]
    h = conn.execute(f"SELECT total_amount, paid_amount, {c['pcode']}, status, COALESCE(invoice_no, {c['iid']}), {c['idate']} "
                     f"FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
    if h is None:
        raise UserError("الفاتورة غير موجودة.")
    rem = round(h[0] - h[1] - E.invoice_discount_total(conn, kind, inv_id), 2)
    return max(rem, 0.0), h


def allocate_advance(trans_id, kind, inv_id, amount=None, a_date=None):
    """تسوية جزء من دفعة مقدمة (أو كلها) على فاتورة لاحقة لنفس الطرف. amount=None = أقصى مبلغ ممكن."""
    require_perm("cash.receipt" if kind == "sales" else "cash.payment", "تسوية دفعة مقدمة")
    a_date = a_date or today()
    if not parse_date(a_date):
        raise UserError("تاريخ التسوية غير صحيح.")
    with db() as conn:
        t = conn.execute("SELECT trans_type, category, party_code, amount, COALESCE(alloc_type,'') FROM cash_transactions "
                         "WHERE trans_id=?", (trans_id,)).fetchone()
        if t is None:
            raise UserError("سند الدفعة غير موجود.")
        if t[1] != CAT_PARTNER or t[4]:
            raise UserError("هذا السند ليس دفعة مقدمة/تحت الحساب (مربوط بفاتورة أو ليس لطرف).")
        if t[0] != _voucher_type(kind):
            raise UserError("نوع الدفعة لا يناسب الفاتورة (قبض العميل لفواتير المبيعات، وصرف المورد لفواتير المشتريات).")
        rem, h = _invoice_remaining(conn, kind, inv_id)
        if h[2] != t[2]:
            raise UserError("الدفعة والفاتورة لطرفين مختلفين.")
        if h[3] != "مفتوحة" or rem <= 0.005:
            raise UserError(f"الفاتورة {h[4]} مسددة بالكامل.")
        avail = round(t[3] - E.advance_allocated_from(conn, trans_id), 2)
        if avail <= 0.005:
            raise UserError("لا يوجد رصيد متاح في هذه الدفعة.")
        amount = round(min(avail, rem) if amount is None else amount, 2)
        if amount <= 0:
            raise UserError("أدخل مبلغ تسوية أكبر من الصفر.")
        if amount > avail + 0.005:
            raise UserError(f"المبلغ أكبر من المتاح في الدفعة ({fmt(avail)}).")
        if amount > rem + 0.005:
            raise UserError(f"المبلغ أكبر من المتبقي على الفاتورة {h[4]} ({fmt(rem)}).")
        aid = conn.execute("INSERT INTO advance_allocations (trans_id, kind, invoice_id, amount, alloc_date, created_by) "
                           "VALUES (?,?,?,?,?,?)", (trans_id, kind, inv_id, amount, a_date, current_user_name())).lastrowid
        refresh_invoice_paid(conn, kind, inv_id)
    return aid, amount


def settle_invoice_from_advances(kind, inv_id):
    """تسوية فاتورة واحدة من دفعات طرفها المتاحة (الأقدم أولاً). يرجع [(سند، مبلغ)]."""
    with db() as conn:
        _, h = _invoice_remaining(conn, kind, inv_id)
    done = []
    for adv in list_advances(h[2], kind):
        with db() as conn:
            rem, _ = _invoice_remaining(conn, kind, inv_id)
        if rem <= 0.005:
            break
        _, amt = allocate_advance(adv["trans_id"], kind, inv_id, min(adv["available"], rem))
        done.append((adv["trans_id"], amt))
    if not done:
        raise UserError("لا توجد دفعات مقدمة متاحة لهذا الطرف.")
    return done


def auto_settle_party(party_code, kind):
    """تسوية كل الدفعات المتاحة للطرف على فواتيره المفتوحة (FIFO: أقدم دفعة على أقدم فاتورة).
    يرجع [(سند، رقم الفاتورة الظاهر، مبلغ)]."""
    c = FLOW[kind]
    with db() as conn:
        invs = conn.execute(f"SELECT {c['iid']}, COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} WHERE {c['pcode']}=? "
                            f"AND status='مفتوحة' ORDER BY {c['idate']}, {c['iid']}", (party_code,)).fetchall()
    out = []
    for inv_id, no in invs:
        try:
            for tid, amt in settle_invoice_from_advances(kind, inv_id):
                out.append((tid, no, amt))
        except UserError:
            if not list_advances(party_code, kind):
                break
    if not out:
        raise UserError("لا توجد دفعات مقدمة متاحة أو فواتير مفتوحة لهذا الطرف.")
    return out


def list_allocations(kind=None, inv_id=None, trans_id=None):
    sql = ("SELECT alloc_id, trans_id, kind, invoice_id, amount, alloc_date, COALESCE(created_by,'') "
           "FROM advance_allocations WHERE 1=1")
    args = []
    for col, v in (("kind", kind), ("invoice_id", inv_id), ("trans_id", trans_id)):
        if v is not None:
            sql += f" AND {col}=?"
            args.append(v)
    with db() as conn:
        return conn.execute(sql + " ORDER BY alloc_id", args).fetchall()


def delete_allocation(alloc_id):
    """إلغاء تسوية: الفاتورة تعود مفتوحة بقيمتها والدفعة تعود متاحة."""
    with db() as conn:
        r = conn.execute("SELECT kind, invoice_id FROM advance_allocations WHERE alloc_id=?", (alloc_id,)).fetchone()
        if r is None:
            raise UserError("التسوية غير موجودة.")
        require_perm("cash.receipt" if r[0] == "sales" else "cash.payment", "إلغاء تسوية دفعة مقدمة")
        conn.execute("DELETE FROM advance_allocations WHERE alloc_id=?", (alloc_id,))
        refresh_invoice_paid(conn, r[0], r[1])


# ============================================================
# 📑 تقرير الفواتير المفتوحة (غير المسددة / المسددة جزئياً) + تقرير الدفعات المقدمة
# ============================================================
OPEN_INV_COLUMNS = [("رقم الفاتورة", "t"), ("التاريخ", "t"), ("الكود", "t"), ("الطرف", "t"), ("الإجمالي", "m"),
                    ("مسدد بسندات", "m"), ("تسوية دفعات مقدمة", "m"), ("خصومات لاحقة", "m"), ("المتبقي", "m"),
                    ("الحالة", "t"), ("العمر (يوم)", "i"), ("دفعات متاحة للطرف", "m")]


def open_invoices_rows(kind, d_to=None, party_code=None):
    c = FLOW[kind]
    d_to = d_to or today()
    as_of = datetime.strptime(d_to, "%Y-%m-%d").date()
    adv = {}
    for a in list_advances(kind=kind):
        adv[a["party_code"]] = adv.get(a["party_code"], 0.0) + a["available"]
    sql = (f"SELECT h.{c['iid']}, COALESCE(h.invoice_no, h.{c['iid']}), h.{c['idate']}, h.{c['pcode']}, h.{c['pname']}, "
           f"h.total_amount, "
           f"(SELECT COALESCE(SUM(amount),0) FROM cash_transactions WHERE alloc_type=? AND alloc_id=h.{c['iid']}), "
           f"(SELECT COALESCE(SUM(amount),0) FROM advance_allocations WHERE kind=? AND invoice_id=h.{c['iid']}), "
           f"(SELECT COALESCE(SUM(total_amount),0) FROM post_discounts WHERE kind=? AND invoice_id=h.{c['iid']}) "
           f"FROM {c['inv']} h WHERE h.status='مفتوحة' AND h.{c['idate']}<=?")
    args = [kind, kind, kind, d_to]
    if party_code:
        sql += f" AND h.{c['pcode']}=?"
        args.append(party_code)
    rows = []
    with db() as conn:
        for iid, no, d, pc, pn, tot, cash, alloc, disc in conn.execute(sql + f" ORDER BY h.{c['idate']}, h.{c['iid']}", args):
            rem = round(tot - cash - alloc - disc, 2)
            if rem <= 0.005:
                continue
            status = "غير مسددة" if cash + alloc + disc <= 0.005 else "مسددة جزئياً"
            age = (as_of - datetime.strptime(d, "%Y-%m-%d").date()).days
            rows.append((str(no), d, pc, pn, tot, cash, alloc, disc, rem, status, age, round(adv.get(pc, 0.0), 2)))
    return rows


def _rep_open_invoices(kind):
    def run(p):
        rows = open_invoices_rows(kind, p.get("d_to"))
        total = ("", "الإجمالي", "", f"{len(rows)} فاتورة", *[sum(r[i] for r in rows) for i in range(4, 9)], "", None, None)
        unpaid = sum(1 for r in rows if r[9] == "غير مسددة")
        notes = [f"غير مسددة: {unpaid}   |   مسددة جزئياً: {len(rows) - unpaid}   |   إجمالي المتبقي: "
                 f"{fmt(sum(r[8] for r in rows))} ر.س",
                 "«دفعات متاحة للطرف» = دفعات مقدمة/تحت الحساب لم تُسوَّ بعد: سوّها من تفاصيل الفاتورة أو شاشة الخزينة.",
                 "الحالة لحظية (حتى الآن)؛ «حتى تاريخ» يحدد الفواتير المصدرة حتى ذلك اليوم وعمرها."]
        title = "الفواتير المفتوحة للعملاء (غير المسددة والمسددة جزئياً)" if kind == "sales" else \
            "الفواتير المفتوحة للموردين (غير المسددة والمسددة جزئياً)"
        return Report(title, OPEN_INV_COLUMNS, rows, total, notes, f"حتى {p.get('d_to')}", chart=(3, 8))
    return run


def rep_advances(p):
    rows = [(a["trans_id"], a["date"], "عميل (قبض)" if a["kind"] == "sales" else "مورد (صرف)", a["party_code"],
             a["party_name"], a["amount"], a["allocated"], a["available"], a["description"]) for a in list_advances()]
    cols = [("رقم السند", "i"), ("التاريخ", "t"), ("النوع", "t"), ("الكود", "t"), ("الطرف", "t"), ("مبلغ الدفعة", "m"),
            ("ما سُوّي منها", "m"), ("المتاح للتسوية", "m"), ("البيان", "t")]
    total = ("", "الإجمالي", "", "", "", *[sum(r[i] for r in rows) for i in (5, 6, 7)], "")
    return Report("الدفعات المقدمة / تحت الحساب غير المسوّاة", cols, rows, total,
                  ["سندات قبض/صرف لأطراف غير مربوطة بفاتورة. سوّها على الفواتير اللاحقة من شاشة الخزينة."], "", chart=(4, 7))


def _register_reports():
    """إضافة التقارير لمجموعة «الذمم والمديونيات» (نفس صلاحية reports.debts) مع شرط البيانات المالية للطرف."""
    group = next(items for title, items in REPORT_CATALOG if "الذمم" in title)
    new = [("open_inv_c", "الفواتير المفتوحة للعملاء (غير مسددة / جزئياً)", _rep_open_invoices("sales"),
            "كل فاتورة مبيعات لم تُسدد بالكامل: المسدد بسندات وبتسوية الدفعات المقدمة والخصومات والمتبقي والعمر."),
           ("open_inv_s", "الفواتير المفتوحة للموردين (غير مسددة / جزئياً)", _rep_open_invoices("purchase"),
            "كل فاتورة مورد لم تُسدد بالكامل، مع الدفعات المقدمة المتاحة لكل مورد."),
           ("advances", "الدفعات المقدمة / تحت الحساب غير المسوّاة", rep_advances,
            "سندات العملاء والموردين غير المربوطة بفاتورة والمتاح منها للتسوية.")]
    for key, title, func, desc in new:
        if key in REPORT_INDEX:
            continue
        group.append((key, title, func, desc))
        REPORT_INDEX[key] = (title, func, desc)
        REPORT_GROUP_OF[key] = "reports.debts"
    REPORT_FIN_REQUIRES.update({"open_inv_c": "sales.customer_fin", "open_inv_s": "purchase.supplier_fin"})


_register_reports()


def open_invoice_drill(key, row):
    """سطر تقرير الفواتير المفتوحة ← (النوع، المعرّف الداخلي) لفتح تفاصيل الفاتورة."""
    kind = "sales" if key == "open_inv_c" else "purchase"
    inv_id = E.find_invoice_id(kind, row[0])
    return (kind, inv_id) if inv_id else None


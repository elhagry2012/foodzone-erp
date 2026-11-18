# -*- coding: utf-8 -*-
"""
التخطيط المالي والعمولات:

1) توقع التدفق النقدي (Cash Flow Forecast): التحصيلات المتوقعة من فواتير العملاء المفتوحة والمدفوعات المتوقعة
   لفواتير الموردين بحسب تاريخ الاستحقاق (تاريخ الفاتورة + مدة السداد للطرف أو الافتراضية)، مع المتأخرات،
   والمصروفات التقديرية من الموازنة، ورصيد النقدية المتوقع تراكمياً.
2) الموازنة التقديرية مقابل الفعلي: موازنة شهرية لصافي المبيعات، والمصروفات التشغيلية، وأي حساب إيراد/مصروف بالدليل.
3) عمولات المناديب: نسبة العمولة من سجل المناديب × أساس الاحتساب (صافي المبيعات أو التحصيلات) حسب فترة الإسناد الفعلية.
"""
from collections import defaultdict
from datetime import date, datetime, timedelta

from . import engine as E
from .engine import UserError, db, Report, SALES_RETURN, TYPE_REV, TYPE_EXP, today, month_keys, register_report_group, \
    get_setting, set_setting
from .credit import default_credit_days

BUDGET_SQL = """
CREATE TABLE IF NOT EXISTS budgets (
    budget_id INTEGER PRIMARY KEY AUTOINCREMENT,
    period TEXT NOT NULL,
    line_key TEXT NOT NULL,
    amount REAL NOT NULL DEFAULT 0,
    notes TEXT,
    updated_by TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (period, line_key)
);
CREATE INDEX IF NOT EXISTS idx_budget_period ON budgets(period);
"""
LINE_SALES, LINE_EXPENSES = "sales", "expenses"
BUILTIN_LINES = {LINE_SALES: "صافي المبيعات (بعد المرتجعات والإشعارات الدائنة)",
                 LINE_EXPENSES: "إجمالي المصروفات التشغيلية (سندات + قيود)"}
COMMISSION_BASE_SETTING = "commission_base"
COMMISSION_BASES = {"sales": "صافي المبيعات (قبل الضريبة)", "collections": "التحصيلات الفعلية"}


def setup_planning(conn):
    conn.executescript(BUDGET_SQL)


E.INIT_HOOKS.append(setup_planning)


# ============================================================
# الموازنة التقديرية
# ============================================================
def budget_line_label(conn, key):
    if key in BUILTIN_LINES:
        return BUILTIN_LINES[key]
    if key.startswith("acc:"):
        code = key[4:]
        r = conn.execute("SELECT account_name FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone()
        return f"{code} - {r[0] if r else '(حساب غير موجود)'}"
    return key


def budget_line_choices():
    """[(المفتاح، الوصف)]: البنود الأساسية + كل حسابات الإيرادات والمصروفات (الرئيسية والفرعية)."""
    with db() as conn:
        accs = conn.execute("SELECT account_code, account_name, account_type, COALESCE(is_group,0) FROM chart_of_accounts "
                            "WHERE account_type IN (?,?) ORDER BY account_code", (TYPE_REV, TYPE_EXP)).fetchall()
    out = list(BUILTIN_LINES.items())
    out += [(f"acc:{c}", f"{c} - {n} ({t}{' • رئيسي' if g else ''})") for c, n, t, g in accs]
    return out


def _check_period(period):
    try:
        datetime.strptime(period + "-01", "%Y-%m-%d")
    except (TypeError, ValueError):
        raise UserError("الفترة يجب أن تكون بصيغة YYYY-MM.")


def save_budget(period, line_key, amount, notes="", user=""):
    _check_period(period)
    if amount is None or amount < 0:
        raise UserError("مبلغ الموازنة رقم صفر أو أكثر.")
    with db() as conn:
        if line_key not in BUILTIN_LINES:
            code = line_key[4:] if line_key.startswith("acc:") else None
            r = conn.execute("SELECT account_type FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone() if code else None
            if r is None or r[0] not in (TYPE_REV, TYPE_EXP):
                raise UserError("بند الموازنة يجب أن يكون صافي المبيعات أو المصروفات أو حساب إيراد/مصروف بالدليل.")
        conn.execute("INSERT INTO budgets (period, line_key, amount, notes, updated_by, updated_at) VALUES (?,?,?,?,?,?) "
                     "ON CONFLICT(period, line_key) DO UPDATE SET amount=excluded.amount, notes=excluded.notes, "
                     "updated_by=excluded.updated_by, updated_at=excluded.updated_at",
                     (period, line_key, round(amount, 2), notes or "", user or E.current_user_name(),
                      datetime.now().strftime("%Y-%m-%d %H:%M:%S")))


def save_budget_year(year, line_key, amounts, user=""):
    """amounts = 12 مبلغاً (يناير..ديسمبر). None = بلا تغيير."""
    if len(amounts) != 12:
        raise UserError("أدخل 12 شهراً.")
    for m, a in enumerate(amounts, start=1):
        if a is not None:
            save_budget(f"{int(year):04d}-{m:02d}", line_key, float(a), user=user)


def delete_budget(period, line_key):
    with db() as conn:
        conn.execute("DELETE FROM budgets WHERE period=? AND line_key=?", (period, line_key))


def list_budgets(year=None):
    """[(الفترة، المفتاح، الوصف، المبلغ، بواسطة، التحديث)]"""
    with db() as conn:
        rows = conn.execute("SELECT period, line_key, amount, COALESCE(updated_by,''), COALESCE(updated_at,'') FROM budgets "
                            + ("WHERE period LIKE ? " if year else "") + "ORDER BY line_key, period",
                            ((f"{int(year):04d}-%",) if year else ())).fetchall()
        return [(p, k, budget_line_label(conn, k), a, u, t) for p, k, a, u, t in rows]


def budget_year_matrix(year):
    """{المفتاح: [12 مبلغاً أو None]} لموازنة سنة."""
    out = defaultdict(lambda: [None] * 12)
    for p, k, _l, a, _u, _t in list_budgets(year):
        out[k][int(p[5:7]) - 1] = a
    return dict(out)


def _actual_monthly(conn, keys, k1, k2):
    """الفعلي لكل بند لكل شهر: {المفتاح: {YYYY-MM: مبلغ}}."""
    out = {k: defaultdict(float) for k in keys}
    d1, d2 = f"{k1}-01", E.month_end(k2).isoformat()
    if LINE_SALES in keys:
        o = out[LINE_SALES]
        for m, v in conn.execute("SELECT substr(invoice_date,1,7), SUM(subtotal) FROM sales_invoices WHERE invoice_date "
                                 "BETWEEN ? AND ? GROUP BY 1", (d1, d2)):
            o[m] += v or 0.0
        for m, v in conn.execute("SELECT substr(return_date,1,7), SUM(subtotal) FROM returns_invoices WHERE return_type=? "
                                 "AND return_date BETWEEN ? AND ? GROUP BY 1", (SALES_RETURN, d1, d2)):
            o[m] -= v or 0.0
        for m, v in conn.execute("SELECT substr(discount_date,1,7), SUM(subtotal) FROM post_discounts WHERE kind='sales' "
                                 "AND discount_date BETWEEN ? AND ? GROUP BY 1", (d1, d2)):
            o[m] -= v or 0.0
    if LINE_EXPENSES in keys:
        for k in month_keys(d1, d2):
            out[LINE_EXPENSES][k] = sum(r[2] for r in E.expense_breakdown(conn, f"{k}-01", E.month_end(k).isoformat()))
    acc_keys = [k for k in keys if k.startswith("acc:")]
    if acc_keys:
        from .ledger import _GL_UNION   # الحسابات من الدفتر العام الكامل
        info, _ = E._account_maps(conn)
        for key in acc_keys:
            code = key[4:]
            codes = E.account_descendants(conn, code)
            typ = info.get(code, {}).get("type")
            ph = ",".join("?" * len(codes))
            for m, dr, cr in conn.execute(f"SELECT substr(d,1,7), SUM(dr), SUM(cr) FROM ({_GL_UNION}) WHERE acc IN ({ph}) "
                                          f"AND d BETWEEN ? AND ? GROUP BY 1", codes + [d1, d2]):
                out[key][m] += ((cr or 0) - (dr or 0)) if typ == TYPE_REV else ((dr or 0) - (cr or 0))
    return out


def _is_revenue(conn, key):
    if key == LINE_SALES:
        return True
    if key.startswith("acc:"):
        r = conn.execute("SELECT account_type FROM chart_of_accounts WHERE account_code=?", (key[4:],)).fetchone()
        return bool(r and r[0] == TYPE_REV)
    return False


def rep_budget_vs_actual(p):
    a, b = p.get("d_from") or f"{date.today().year}-01-01", p.get("d_to") or today()
    k1, k2 = a[:7], b[:7]
    months = month_keys(a, b)
    from .ledger import sync_ledger
    sync_ledger()
    with db() as conn:
        budget = defaultdict(dict)
        for per, key, amt in conn.execute("SELECT period, line_key, amount FROM budgets WHERE period BETWEEN ? AND ?", (k1, k2)):
            budget[key][per] = amt
        keys = list(budget) or [LINE_SALES, LINE_EXPENSES]
        for k in (LINE_SALES, LINE_EXPENSES):
            if k not in keys:
                keys.append(k)
        actual = _actual_monthly(conn, keys, k1, k2)
        labels = {k: budget_line_label(conn, k) for k in keys}
        revenue = {k: _is_revenue(conn, k) for k in keys}
    rows = []
    order = sorted(keys, key=lambda k: (k not in BUILTIN_LINES, k))
    for key in order:
        tb = ta = 0.0
        for m in months:
            bv, av = budget[key].get(m), actual[key].get(m, 0.0)
            if bv is None and abs(av) < 0.005:
                continue
            tb += bv or 0.0
            ta += av
            rows.append(_bva_row(labels[key], m, bv, av, revenue[key]))
        rows.append(_bva_row(f"◀ إجمالي: {labels[key]}", f"{k1} → {k2}", tb, ta, revenue[key]))
    cols = [("البند", "t"), ("الشهر", "t"), ("الموازنة التقديرية", "m"), ("الفعلي", "m"), ("الانحراف (فعلي - تقديري)", "m"),
            ("نسبة التنفيذ", "p"), ("التقييم", "t")]
    notes = ["الإيرادات: تجاوز الموازنة = انحراف ملائم ✓. المصروفات: تجاوز الموازنة = انحراف غير ملائم ✗.",
             "صافي المبيعات من الفواتير بعد المرتجعات والإشعارات الدائنة (مثل قائمة الدخل). المصروفات التشغيلية من "
             "سندات الصرف والقيود (عدا التكلفة والإهلاك). بنود الحسابات من الدفتر العام الكامل (تشمل فروعها).",
             "تُدخل الموازنة من صفحة «التخطيط المالي ← الموازنة التقديرية»."]
    return Report("الموازنة التقديرية مقابل الفعلي", cols, rows, None, notes, f"من {a} إلى {b}")


def _bva_row(label, month, bv, av, is_rev):
    if bv is None:
        return (label, month, None, av, None, None, "بدون موازنة")
    var = av - bv
    pct = (av / bv * 100) if bv else None
    if abs(var) < 0.005:
        ev = "مطابق"
    elif (var > 0) == is_rev:
        ev = "ملائم ✓"
    else:
        ev = "غير ملائم ✗"
    return (label, month, bv, av, var, pct, ev)


# ============================================================
# توقع التدفق النقدي
# ============================================================
def _due(inv_date, days):
    try:
        return (datetime.strptime(inv_date[:10], "%Y-%m-%d").date() + timedelta(days=int(days or 0))).isoformat()
    except (TypeError, ValueError):
        return inv_date


def open_invoices_due(as_of=None):
    """الفواتير المفتوحة بتاريخ الاستحقاق: [(النوع c/s، الرقم، التاريخ، الاستحقاق، الكود، الطرف، المتبقي)]"""
    as_of = as_of or today()
    dflt = default_credit_days()
    out = []
    with db() as conn:
        terms = {c: d for c, d in conn.execute("SELECT partner_code, credit_days FROM business_partners")}
        for kind, sql in (("c", "SELECT invoice_id, COALESCE(invoice_no, invoice_id), invoice_date, customer_code, customer_name, "
                                "total_amount, paid_amount FROM sales_invoices"),
                          ("s", "SELECT purchase_id, COALESCE(invoice_no, purchase_id), purchase_date, supplier_code, supplier_name, "
                                "total_amount, paid_amount FROM purchase_invoices")):
            for iid, no, d, code, name, total, paid in conn.execute(sql).fetchall():
                disc = E.invoice_discount_total(conn, "sales" if kind == "c" else "purchase", iid)
                rem = round((total or 0) - (paid or 0) - (disc or 0), 2)
                if rem <= 0.005:
                    continue
                days = terms.get(code)
                out.append((kind, str(no), d, _due(d, dflt if days is None else days), code, name, rem))
    return out


def _bucket(d, gran):
    dt = datetime.strptime(d, "%Y-%m-%d").date()
    if gran == "month":
        return dt.strftime("%Y-%m")
    start = dt - timedelta(days=(dt.weekday() + 2) % 7)   # الأسبوع يبدأ السبت
    return start.isoformat()


def cash_flow_forecast(horizon_to=None, gran="month", include_budget=True):
    """[(الفترة، تحصيلات متوقعة، مدفوعات للموردين، مصروفات تقديرية، الصافي، رصيد النقدية المتوقع)] + الرصيد الحالي."""
    t = date.today()
    horizon_to = horizon_to or (t + timedelta(days=90)).isoformat()
    inv = open_invoices_due()
    with db() as conn:
        cash_now = round(sum(E.cash_balances(conn).values()), 2)
        budget = {}
        if include_budget:
            budget = {p: a for p, a in conn.execute("SELECT period, amount FROM budgets WHERE line_key=?", (LINE_EXPENSES,))}
    flows = defaultdict(lambda: [0.0, 0.0, 0.0])
    overdue = [0.0, 0.0, 0.0]
    for kind, _no, _d, due, *_rest, rem in inv:
        if due > horizon_to:
            continue
        slot = overdue if due < t.isoformat() else flows[_bucket(due, gran)]
        slot[0 if kind == "c" else 1] += rem
    if include_budget and budget:
        for k in month_keys(t.isoformat(), horizon_to):
            amt = budget.get(k)
            if not amt:
                continue
            # الشهر الحالي: الجزء المتبقي فقط (نسبة الأيام المتبقية)
            me = E.month_end(k)
            ms = date(int(k[:4]), int(k[5:7]), 1)
            share = 1.0 if ms > t else max((me - t).days + 1, 0) / ((me - ms).days + 1)
            if gran == "month":
                flows[k][2] += amt * share
            else:   # توزيع على الأسابيع بالتساوي حسب الأيام
                d = max(ms, t)
                end = min(me, datetime.strptime(horizon_to, "%Y-%m-%d").date())
                n = (end - d).days + 1
                while d <= end:
                    flows[_bucket(d.isoformat(), gran)][2] += amt * share / max(n, 1)
                    d += timedelta(days=1)
    rows, bal = [], cash_now
    if any(abs(x) > 0.005 for x in overdue):
        net = overdue[0] - overdue[1] - overdue[2]
        bal += net
        rows.append(("متأخرات مستحقة قبل اليوم", overdue[0], overdue[1], overdue[2], net, bal))
    for k in sorted(flows):
        i, o, e = flows[k]
        net = i - o - e
        bal += net
        label = k if gran == "month" else f"أسبوع {k}"
        rows.append((label, i, o, e, net, bal))
    return rows, cash_now


def rep_cash_forecast(p):
    t = date.today()
    horizon = p.get("d_to") or today()
    if horizon <= t.isoformat():
        horizon = (t + timedelta(days=90)).isoformat()
    gran = "month" if p.get("group", "month") == "month" else "week"
    rows, cash_now = cash_flow_forecast(horizon, gran)
    total = ("الإجمالي", sum(r[1] for r in rows), sum(r[2] for r in rows), sum(r[3] for r in rows),
             sum(r[4] for r in rows), rows[-1][5] if rows else cash_now)
    cols = [("الفترة (الاستحقاق)", "t"), ("تحصيلات متوقعة من العملاء", "m"), ("مدفوعات متوقعة للموردين", "m"),
            ("مصروفات تقديرية (الموازنة)", "m"), ("صافي التدفق", "m"), ("رصيد النقدية المتوقع", "m")]
    neg = [r[0] for r in rows if r[5] < -0.005]
    notes = [f"رصيد النقدية والبنوك الحالي: {cash_now:,.2f}. الاستحقاق = تاريخ الفاتورة + مدة السداد للطرف "
             f"(الافتراضية {default_credit_days()} يوماً).",
             "المتبقي = الإجمالي - المسدد - الخصومات اللاحقة. المصروفات التقديرية من بند «المصروفات التشغيلية» بالموازنة.",
             ("⚠ عجز نقدي متوقع في: " + "، ".join(neg)) if neg else "لا يوجد عجز نقدي متوقع خلال الفترة ✓"]
    return Report("توقع التدفق النقدي (حسب تواريخ الاستحقاق)", cols, rows, total, notes,
                  f"من اليوم حتى {horizon} ({'شهري' if gran == 'month' else 'أسبوعي'})", chart=(0, 5))


def rep_due_invoices(p):
    horizon = p.get("d_to") or today()
    t = today()
    rows = []
    for kind, no, d, due, code, name, rem in sorted(open_invoices_due(), key=lambda r: r[3]):
        late = (datetime.strptime(t, "%Y-%m-%d").date() - datetime.strptime(due, "%Y-%m-%d").date()).days
        rows.append(("تحصيل من عميل" if kind == "c" else "سداد لمورد", no, d, due, code, name, rem,
                     late if late > 0 else 0, "متأخر" if late > 0 else ("مستحق خلال الفترة" if due <= horizon else "لاحقاً")))
    cols = [("النوع", "t"), ("رقم الفاتورة", "t"), ("التاريخ", "t"), ("الاستحقاق", "t"), ("الكود", "t"), ("الطرف", "t"),
            ("المتبقي", "m"), ("أيام التأخير", "i"), ("الحالة", "t")]
    return Report("الفواتير المفتوحة بتواريخ الاستحقاق", cols, rows, None,
                  ["تفاصيل توقع التدفق النقدي: كل فاتورة مفتوحة بتاريخ استحقاقها والمتبقي منها."], f"حتى {horizon}")


# ============================================================
# عمولات المناديب
# ============================================================
def commission_base():
    v = get_setting(COMMISSION_BASE_SETTING, "sales")
    return v if v in COMMISSION_BASES else "sales"


def set_commission_base(base):
    if base not in COMMISSION_BASES:
        raise UserError("أساس العمولة غير صحيح.")
    set_setting(COMMISSION_BASE_SETTING, base)


def commission_rows(p, rates=None, uniform=None, base=None):
    """احتساب العمولات لحظياً. rates = {اسم المندوب: نسبة} تتجاوز سجل المناديب، uniform = نسبة موحدة للكل
    (تُطبق على من ليس له نسبة في rates)، base = sales | collections (الافتراضي من الإعدادات).
    يرجع [(الكود، المندوب، النسبة، صافي المبيعات، التحصيلات، أساس الاحتساب، العمولة، ملاحظات)]."""
    perf = E.rep_rep_performance(p)
    base = base if base in COMMISSION_BASES else commission_base()
    rates = rates or {}
    with db() as conn:
        reps = {n: (c or "", pct or 0.0, act) for n, c, pct, act in
                conn.execute("SELECT name, rep_code, commission_pct, active FROM sales_reps")}
    rows = []
    for r in perf.rows:
        name, sales, coll = r[0], r[3], r[4]
        code, pct, active = reps.get(name, ("", 0.0, 1))
        if name in rates and rates[name] is not None:
            pct = float(rates[name])
        elif uniform is not None:
            pct = float(uniform)
        if pct < 0 or pct > 100:
            raise UserError(f"نسبة عمولة المندوب «{name}» يجب أن تكون بين 0 و 100.")
        base_val = sales if base == "sales" else coll
        comm = round(max(base_val, 0.0) * (pct or 0.0) / 100, 2)
        note = ("غير مسجل بسجل المناديب" if name not in reps and name != E.NO_REP else
                "بدون مندوب" if name == E.NO_REP else ("موقوف" if not active else ""))
        rows.append((code, name, pct, sales, coll, base_val, comm, note))
    rows.sort(key=lambda r: -r[6])
    return rows, base


def save_commission_rates(rates):
    """حفظ نسب العمولة المعدلة في سجل المناديب (للمسجلين فقط). يرجع عدد المناديب المحدثين."""
    n = 0
    with db() as conn:
        for name, pct in rates.items():
            if pct is None:
                continue
            if pct < 0 or pct > 100:
                raise UserError(f"نسبة عمولة المندوب «{name}» يجب أن تكون بين 0 و 100.")
            n += conn.execute("UPDATE sales_reps SET commission_pct=? WHERE name=?", (round(float(pct), 4), name)).rowcount
    return n


def rep_commissions(p):
    a, b = p["d_from"], p["d_to"]
    rows, base = commission_rows(p, p.get("rates"), p.get("uniform"), p.get("base"))
    cols = [("الكود", "t"), ("المندوب", "t"), ("نسبة العمولة", "p"), ("صافي المبيعات", "m"), ("التحصيلات", "m"),
            ("أساس الاحتساب", "m"), ("العمولة المستحقة", "m"), ("ملاحظات", "t")]
    total = ("الإجمالي", "", None, sum(r[3] for r in rows), sum(r[4] for r in rows), sum(r[5] for r in rows),
             sum(r[6] for r in rows), "")
    src = ("نسبة موحدة " + f"{p['uniform']:g}%" if p.get("uniform") is not None else "نسب سجل المناديب")
    notes = [f"أساس الاحتساب: {COMMISSION_BASES[base]}  |  مصدر النسب: {src}"
             + (" مع تعديلات مباشرة لبعض المناديب" if p.get("rates") else "") + ".",
             "العمولة = أساس الاحتساب × نسبة العمولة. الإسناد حسب فترة تواجد المندوب الفعلية "
             "مع كل عميل (بدون تداخل بين مندوبين).", "صافي المبيعات بعد المرتجعات والإشعارات الدائنة، وقبل الضريبة."]
    return Report("عمولات المناديب المستحقة", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 6))


def _credit_report(p):
    from .credit import rep_credit_limits
    return rep_credit_limits(p)


register_report_group("plan", "🧭 التخطيط المالي والعمولات", [
    ("cf_forecast", "توقع التدفق النقدي", rep_cash_forecast,
     "التحصيلات والمدفوعات المتوقعة حسب تواريخ استحقاق الفواتير + المصروفات التقديرية، حتى تاريخ (إلى). التجميع: شهري/أسبوعي."),
    ("due_inv", "الفواتير المفتوحة بتواريخ الاستحقاق", rep_due_invoices, "تفاصيل التوقع: كل فاتورة عميل/مورد مفتوحة واستحقاقها."),
    ("bva", "الموازنة التقديرية مقابل الفعلي", rep_budget_vs_actual, "المبيعات والمصروفات والحسابات: تقديري مقابل فعلي لكل شهر."),
    ("commissions", "عمولات المناديب", rep_commissions, "احتساب آلي للعمولة من نسبة كل مندوب وأدائه في الفترة."),
    ("credit", "حدود الائتمان واستخدامها", _credit_report, "تعرض كل عميل مقابل حده الائتماني والمتاح ونسبة الاستخدام."),
])


# -*- coding: utf-8 -*-
"""
الحد الائتماني للعملاء (Customer Credit Limit).

- لكل عميل: الحد الائتماني (0 = بدون حد)، مدة السداد بالأيام (تُستخدم في توقع التدفق النقدي)، وسياسة التجاوز:
    warn  = تنبيه فوري ويُسمح بالمتابعة
    block = منع البيع الآجل (أمر البيع / إذن التسليم) — يتجاوزه مدير النظام فقط بموافقة صريحة
- التعرض الائتماني = رصيد العميل المستحق + تسليمات لم تُفوتر بعد (شاملة الضريبة) + المتبقي من أوامر البيع المفتوحة
  (عند إنشاء أمر جديد) + قيمة المستند الجديد.
- الفحص داخل المحرك نفسه (create_order / fulfil_order) فيُطبَّق على كل الواجهات.
- الفوترة من أذون التسليم لا تُمنع أبداً (البضاعة خرجت والفاتورة الضريبية إلزامية) لكن يظهر تنبيه.
"""
from .engine import UserError, db, CAT_PARTNER, SALES_RETURN, PURCHASE_RETURN, VAT_RATE, Report, get_setting, set_setting, \
    OPEN_BLOCKED, today

POLICY_WARN, POLICY_BLOCK = "warn", "block"
CREDIT_POLICIES = {POLICY_WARN: "تنبيه فقط (يُسمح بالمتابعة)", POLICY_BLOCK: "منع البيع الآجل عند التجاوز"}
DEFAULT_POLICY_SETTING = "credit_policy_default"
DEFAULT_DAYS_SETTING = "credit_days_default"
DEFAULT_CREDIT_DAYS = 30
STAGE_LABEL = {"order": "أمر بيع", "delivery": "إذن تسليم", "invoice": "فاتورة"}


class CreditLimitError(UserError):
    """تجاوز الحد الائتماني لعميل سياسته «منع البيع الآجل»."""

    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


def default_policy():
    v = get_setting(DEFAULT_POLICY_SETTING, POLICY_WARN)
    return v if v in CREDIT_POLICIES else POLICY_WARN


def default_credit_days():
    try:
        return int(float(get_setting(DEFAULT_DAYS_SETTING, str(DEFAULT_CREDIT_DAYS))))
    except ValueError:
        return DEFAULT_CREDIT_DAYS


def set_credit_defaults(policy=None, days=None):
    if policy is not None:
        if policy not in CREDIT_POLICIES:
            raise UserError("سياسة الائتمان غير صحيحة.")
        set_setting(DEFAULT_POLICY_SETTING, policy)
    if days is not None:
        if int(days) < 0:
            raise UserError("مدة السداد لا تكون سالبة.")
        set_setting(DEFAULT_DAYS_SETTING, str(int(days)))


def set_partner_credit(conn, code, limit=None, days=None, policy=None):
    """تعيين الحد الائتماني/مدة السداد/السياسة لعميل (None = بلا تغيير، "" للسياسة = الافتراضية)."""
    if limit is not None:
        if limit < 0:
            raise UserError("الحد الائتماني لا يكون سالباً (0 = بدون حد).")
        conn.execute("UPDATE business_partners SET credit_limit=? WHERE partner_code=?", (round(limit, 2), code))
    if days is not None:
        if days != "" and int(days) < 0:
            raise UserError("مدة السداد لا تكون سالبة.")
        conn.execute("UPDATE business_partners SET credit_days=? WHERE partner_code=?",
                     (None if days == "" else int(days), code))
    if policy is not None:
        if policy and policy not in CREDIT_POLICIES:
            raise UserError("سياسة الائتمان غير صحيحة.")
        conn.execute("UPDATE business_partners SET credit_policy=? WHERE partner_code=?", (policy or None, code))


def partner_credit_terms(conn, code):
    """(الحد، مدة السداد الفعلية، السياسة الفعلية، السياسة المخزنة)"""
    r = conn.execute("SELECT COALESCE(credit_limit,0), credit_days, COALESCE(credit_policy,'') FROM business_partners "
                     "WHERE partner_code=?", (code,)).fetchone()
    if r is None:
        return 0.0, default_credit_days(), default_policy(), ""
    return (r[0] or 0.0, r[1] if r[1] is not None else default_credit_days(),
            r[2] if r[2] in CREDIT_POLICIES else default_policy(), r[2] or "")


def _balance(conn, code):
    """رصيد الطرف (مدين موجب) بنفس معادلة partner_balances لكن لطرف واحد وحتى الآن."""
    one = lambda sql, *a: conn.execute(sql, a).fetchone()[0] or 0.0  # noqa: E731
    od, oc = conn.execute("SELECT COALESCE(opening_debit,0), COALESCE(opening_credit,0) FROM business_partners "
                          "WHERE partner_code=?", (code,)).fetchone() or (0.0, 0.0)
    sales = one("SELECT SUM(total_amount) FROM sales_invoices WHERE customer_code=?", code)
    purch = one("SELECT SUM(total_amount) FROM purchase_invoices WHERE supplier_code=?", code)
    ret_s = one("SELECT SUM(total_amount) FROM returns_invoices WHERE party_code=? AND return_type=?", code, SALES_RETURN)
    ret_p = one("SELECT SUM(total_amount) FROM returns_invoices WHERE party_code=? AND return_type=?", code, PURCHASE_RETURN)
    rec = one("SELECT SUM(amount) FROM cash_transactions WHERE party_code=? AND category=? AND trans_type='قبض'", code, CAT_PARTNER)
    pay = one("SELECT SUM(amount) FROM cash_transactions WHERE party_code=? AND category=? AND trans_type='صرف'", code, CAT_PARTNER)
    disc_s = one("SELECT SUM(total_amount) FROM post_discounts WHERE party_code=? AND kind='sales'", code)
    disc_p = one("SELECT SUM(total_amount) FROM post_discounts WHERE party_code=? AND kind='purchase'", code)
    return (od - oc) + sales + ret_p + pay + disc_p - purch - ret_s - rec - disc_s


def _unbilled(conn, code):
    v = conn.execute("SELECT COALESCE(SUM(l.quantity*l.unit_price - COALESCE(l.discount_amount,0)),0) FROM sales_delivery_lines l "
                     "JOIN sales_deliveries d ON d.delivery_id=l.delivery_id WHERE d.customer_code=? AND d.invoice_id IS NULL",
                     (code,)).fetchone()[0] or 0.0
    return round(v * (1 + VAT_RATE), 2)


def _open_orders(conn, code, exclude_order=None):
    ph = ",".join("?" * len(OPEN_BLOCKED))
    v = conn.execute(f"SELECT COALESCE(SUM((l.quantity - l.delivered_qty) * (l.unit_price - COALESCE(l.discount_amount,0)/l.quantity)),0) "
                     f"FROM sales_order_lines l JOIN sales_orders o ON o.order_id=l.order_id WHERE o.customer_code=? "
                     f"AND o.status NOT IN ({ph}) AND l.quantity > l.delivered_qty AND o.order_id != ?",
                     (code, *OPEN_BLOCKED, exclude_order or -1)).fetchone()[0] or 0.0
    return round(v * (1 + VAT_RATE), 2)


def credit_status_conn(conn, code, amount=0.0, stage="order", exclude_order=None):
    limit, days, policy, _ = partner_credit_terms(conn, code)
    bal = round(_balance(conn, code), 2)
    unb = _unbilled(conn, code)
    opn = _open_orders(conn, code, exclude_order) if stage == "order" else 0.0
    exposure = round(max(bal, 0.0) + unb + opn, 2)
    after = round(exposure + (amount or 0.0), 2)
    exceeded = limit > 0.005 and after > limit + 0.005
    name = (conn.execute("SELECT name FROM business_partners WHERE partner_code=?", (code,)).fetchone() or [code])[0]
    msg = ""
    if exceeded:
        msg = (f"تجاوز الحد الائتماني للعميل {name} ({code}):\n"
               f"الحد المسموح: {limit:,.2f}\nالرصيد المستحق: {bal:,.2f}\nتسليمات لم تُفوتر: {unb:,.2f}"
               + (f"\nأوامر بيع مفتوحة: {opn:,.2f}" if stage == "order" else "")
               + f"\nقيمة {STAGE_LABEL.get(stage, 'المستند')} الحالي: {amount or 0:,.2f}\n"
               f"الإجمالي بعد العملية: {after:,.2f}  (تجاوز بمقدار {after - limit:,.2f})")
    return dict(code=code, name=name, limit=limit, days=days, policy=policy, balance=bal, unbilled=unb, open_orders=opn,
                exposure=exposure, amount=amount or 0.0, after=after, available=(limit - after) if limit > 0.005 else None,
                exceeded=exceeded, usage=(after / limit * 100) if limit > 0.005 else None, message=msg, stage=stage)


def credit_status(code, amount=0.0, stage="order", exclude_order=None):
    with db() as conn:
        return credit_status_conn(conn, code, amount, stage, exclude_order)


def check_credit(conn, code, amount, stage, allow_over=False, exclude_order=None):
    """يُستدعى من المحرك قبل حفظ البيع الآجل. سياسة «منع» + تجاوز = خطأ CreditLimitError (إلا بإذن صريح)."""
    st = credit_status_conn(conn, code, amount, stage, exclude_order)
    if st["exceeded"] and st["policy"] == POLICY_BLOCK and not allow_over and stage != "invoice":
        raise CreditLimitError(st["message"] + "\n\nسياسة العميل: منع البيع الآجل عند التجاوز. "
                               "حصّل جزءاً من المديونية أو ارفع الحد الائتماني، أو اطلب موافقة مدير النظام.", st)
    return st


def rep_credit_limits(p):
    with db() as conn:
        codes = [r[0] for r in conn.execute("SELECT partner_code FROM business_partners WHERE is_customer=1 ORDER BY name")]
        rows = []
        for code in codes:
            st = credit_status_conn(conn, code, 0.0, "order")
            if st["limit"] <= 0.005 and abs(st["balance"]) < 0.005 and st["unbilled"] < 0.005 and st["open_orders"] < 0.005:
                continue
            state = ("بدون حد" if st["limit"] <= 0.005 else "متجاوز ⛔" if st["exceeded"] or st["after"] > st["limit"]
                     else "قريب من الحد ⚠" if st["usage"] and st["usage"] >= 80 else "ضمن الحد ✓")
            rows.append((code, st["name"], st["limit"] or None, st["days"], CREDIT_POLICIES[st["policy"]].split(" (")[0],
                         st["balance"], st["unbilled"], st["open_orders"], st["exposure"],
                         st["available"], st["usage"], state))
    rows.sort(key=lambda r: (r[10] is None, -(r[10] or 0)))
    cols = [("الكود", "t"), ("العميل", "t"), ("الحد الائتماني", "m"), ("مدة السداد (يوم)", "i"), ("السياسة", "t"),
            ("الرصيد المستحق", "m"), ("تسليمات لم تُفوتر", "m"), ("أوامر مفتوحة", "m"), ("إجمالي التعرض", "m"),
            ("المتاح", "m"), ("نسبة الاستخدام", "p"), ("الحالة", "t")]
    total = ("الإجمالي", "", sum(r[2] or 0 for r in rows), None, "", sum(r[5] for r in rows), sum(r[6] for r in rows),
             sum(r[7] for r in rows), sum(r[8] for r in rows), None, None, "")
    notes = ["إجمالي التعرض = الرصيد المستحق + تسليمات لم تُفوتر (شاملة الضريبة) + المتبقي من أوامر البيع المفتوحة.",
             f"السياسة الافتراضية: {CREDIT_POLICIES[default_policy()]}، ومدة السداد الافتراضية {default_credit_days()} يوماً.",
             "الحد صفر = بدون حد ائتماني. الأرصدة لحظية حتى اليوم."]
    return Report("حدود الائتمان واستخدامها (العملاء)", cols, rows, total, notes, f"حتى {today()}", chart=(1, 8))

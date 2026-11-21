# -*- coding: utf-8 -*-
"""
مستورد HansaWorld (الترحيل من النظام السابق) — مطابق لهيكل الملفات الفعلية المصدَّرة من 2025-07 حتى 2026-11-19.

الملفات (كلها اختيارية، ويمكن الاستيراد على مراحل):
    1) العملاء  «تفاصيل العملاء.xlsx» شيت «ورقة1»:
       Customer Code, Customer Name, Trans Date, Type (INV/REC/CRED/CASH), Doc No, Applied To Doc,
       Debit (SAR), Credit (SAR), Currency, Running Balance (SAR)
    2) الموردون «حركات الموردين بالتفصيل ....xlsx» شيت «Supplier Transactions»:
       Supplier Code, Supplier Name, VAT No, Currency, Trans Date, Type (PINV/PAYM/CRED), Type Description, Doc No,
       Applied To Invoice, Debit (SAR), Credit (SAR), Balance (SAR), Source Row   (+ سطر «Net Change» الإجمالي)
    3) حركات الأصناف «حركات الاصناف بالتفصيل ....xlsx» شيت «Item History»:
       Item, Name, Unit, Date, Type, Reference, Cust/Supp., Location, Qty, Balance   (+ سطور End Balance + سطر إجمالي)
    4) التكلفة والرصيد «تكلفه الخامات ....xlsx» شيت «Item Cost Price Analysis»: Item, Name, In Stock, Cost Pr., Last Price

قواعد الترحيل:
    - الشركاء: كود HansaWorld هو المرجع. العملاء بنفس الكود (C0001)، والموردون كيان منفصل S0001 (لأن HansaWorld يستخدم
      نفس الكود C0001 للعميل والمورد)، ويُحفظ كود HansaWorld في legacy_code. الشركاء الموجودون بالنظام يُطابَقون بالاسم
      (ثم الرقم الضريبي) ويُعاد ترقيمهم لكود HansaWorld مع الاحتفاظ بهاتفهم ومندوبهم وحدهم الائتماني.
    - INV/CASH ← فاتورة مبيعات بالإجمالي (التاريخ ورقم المستند الأصليان)، PINV ← فاتورة مشتريات.
      CRED ← إشعار دائن / خصم لاحق مربوط بالفاتورة (Applied To). الضريبة مستخرجة من الإجمالي بنسبة 15%.
    - REC/PAYM ← سند واحد لكل رقم مستند على «حساب وسيط ترحيل HansaWorld» (أو حساب تختاره)، وتوزيعه على الفواتير
      من Applied To (سطور الخصم السالبة تُلغي تسويات سابقة وتعيد توزيعها) — نفس جدول توزيع السندات في شاشة الخزينة.
    - الأصناف: تعريف كل الأصناف، والرصيد الافتتاحي = In Stock بتكلفة Cost Pr. (والباقي صفر)، وآخر سعر شراء = Last Price.
      حركات الأصناف التفصيلية تُحفظ في hw_item_history للمرجعية والمطابقة (لا تُكرر حركة المخزون).
    - منع التكرار: hw_import_map (مستند لكل دور ونوع ورقم وطرف) و row_key فريد لكل حركة صنف — إعادة الاستيراد آمنة.
    - المعاينة (dry_run) تنفذ كل شيء داخل معاملة ثم تتراجع عنها، وتعرض نفس الإحصاءات وتقرير المطابقة.
"""
import math
import re
import sqlite3
from collections import defaultdict
from datetime import date, datetime, timedelta

import pandas as pd

from . import engine as E
from .engine import CAT_PARTNER, Report, UserError, db, fmt, norm_name

HW_SOURCE = "HansaWorld"
HW_CASH_ACCOUNT = "وسيط ترحيل HansaWorld"
HW_CASH_PARENT = "11120"          # النقدية بالبنوك (حساب الوسيط يظهر ضمن حسابات الخزينة)
HW_DONE_SETTING = "hw_import_last"
VAT_RATE = 0.15
DUP_MARK = "مكرر"

HW_SQL = """
CREATE TABLE IF NOT EXISTS hw_import_map (
    map_id INTEGER PRIMARY KEY AUTOINCREMENT,
    role TEXT NOT NULL,
    doc_type TEXT NOT NULL,
    doc_no TEXT NOT NULL,
    party_code TEXT NOT NULL,
    target TEXT NOT NULL,
    target_id INTEGER,
    doc_date TEXT,
    amount REAL,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (role, doc_type, doc_no, party_code)
);
CREATE INDEX IF NOT EXISTS idx_hwmap_target ON hw_import_map(target, target_id);
CREATE TABLE IF NOT EXISTS hw_item_history (
    hist_id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_key TEXT UNIQUE NOT NULL,
    item_code TEXT NOT NULL,
    item_name TEXT,
    unit TEXT,
    move_date TEXT NOT NULL,
    move_type TEXT NOT NULL,
    reference TEXT,
    party_text TEXT,
    party_code TEXT,
    is_duplicate INTEGER NOT NULL DEFAULT 0,
    location TEXT,
    qty REAL NOT NULL,
    balance REAL,
    link_doc_type TEXT,
    link_doc_no TEXT,
    link_method TEXT,
    src_row INTEGER
);
CREATE INDEX IF NOT EXISTS idx_hwhist_item ON hw_item_history(item_code, move_date);
CREATE INDEX IF NOT EXISTS idx_hwhist_ref ON hw_item_history(reference);
-- ميزان المراجعة التحليلي من HansaWorld: أرصدة مرجعية للمطابقة فقط (لا تُرحَّل كقيود)
CREATE TABLE IF NOT EXISTS hw_tb_reference (
    account_code TEXT PRIMARY KEY,
    account_name TEXT,
    group_code TEXT,
    per_debit REAL NOT NULL DEFAULT 0,
    per_credit REAL NOT NULL DEFAULT 0,
    balance REAL NOT NULL DEFAULT 0,
    as_of TEXT,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP
);
-- قوائم دخل HansaWorld لفترات جزئية (مثل 2026-01-01 → 2026-11-19): تقسم فترة الميزان إلى فترات مرجعية.
-- amount مدين موجب (الإيرادات سالبة) كما في تقرير Profit & Loss.
CREATE TABLE IF NOT EXISTS hw_pl_reference (
    period_from TEXT NOT NULL,
    period_to TEXT NOT NULL,
    account_code TEXT NOT NULL,
    account_name TEXT,
    amount REAL NOT NULL DEFAULT 0,
    imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (period_from, period_to, account_code)
);
"""


def setup_hansa(conn):
    conn.executescript(HW_SQL)


if setup_hansa not in E.INIT_HOOKS:
    E.INIT_HOOKS.append(setup_hansa)

CUST_COLS = ["Customer Code", "Customer Name", "Trans Date", "Type", "Doc No", "Applied To Doc", "Debit (SAR)",
             "Credit (SAR)", "Currency", "Running Balance (SAR)"]
SUPP_COLS = ["Supplier Code", "Supplier Name", "VAT No", "Currency", "Trans Date", "Type", "Type Description", "Doc No",
             "Applied To Invoice", "Debit (SAR)", "Credit (SAR)", "Balance (SAR)", "Source Row"]
HIST_COLS = ["Item", "Name", "Unit", "Date", "Type", "Reference", "Cust/Supp.", "Location", "Qty", "Balance"]
COST_COLS = ["Item", "Name", "In Stock", "Cost Pr.", "Last Price"]
CUST_TYPES = ("INV", "REC", "CRED", "CASH")
SUPP_TYPES = ("PINV", "PAYM", "CRED")
HIST_END = "End Balance"
# نوع حركة الصنف ← (دور الطرف، أنواع المستندات المحتملة للربط)
HIST_LINKS = {"Deld": ("customer", ("INV", "CASH")), "Cust. Ret.": ("customer", ("CRED",)),
              "G.Rec.": ("supplier", ("PINV",)), "Ret. to Supp.": ("supplier", ("CRED",)), "St. Depr.": (None, ())}

STAT_LABELS = dict(
    customers="عملاء", suppliers="موردون", partners_created="شركاء جدد", partners_recoded="شركاء أُعيد ترقيمهم",
    partners_evicted="شركاء قدامى نُقلوا لكود OLD-", items="أصناف", items_created="أصناف جديدة",
    items_stock="أصناف لها رصيد", sales_invoices="فواتير مبيعات", sales_credits="إشعارات دائنة للعملاء",
    receipts="سندات قبض", customer_refunds="سندات صرف (رد للعميل)", purchase_invoices="فواتير مشتريات",
    supplier_credits="إشعارات خصم موردين", payments="سندات صرف للموردين", supplier_refunds="سندات قبض (رد من المورد)",
    allocations="سطور توزيع على الفواتير", credit_splits="إشعارات دائنة فرعية (تقسيم)", history_rows="حركات أصناف", history_linked="حركات مربوطة بمستند", tb_accounts="حسابات الميزان",
    tb_accounts_created="حسابات عُرّفت بالدليل", cash_openings="أرصدة افتتاحية للخزينة", expenses="سندات مصاريف",
    expense_total="إجمالي المصاريف", bad="سطور مرفوضة",
    skipped_existing="مستندات موجودة مسبقاً (تخطي)")


# ============================================================
# قراءة الملفات
# ============================================================
def _blank(v):
    return v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and not v.strip()) \
        or v is pd.NaT


def _txt(v):
    if _blank(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r"\s+", " ", str(v)).strip()


def _num(v):
    if _blank(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = E.normalize_digits(str(v)).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _vat_text(v):
    """الرقم الضريبي (15 رقماً) قد يُقرأ من إكسيل كرقم عشري علمي: نعيده نصاً كاملاً."""
    if _blank(v):
        return None
    if isinstance(v, float):
        return str(int(round(v)))
    return re.sub(r"\D", "", E.normalize_digits(str(v))) or None


_DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y", "%d.%m.%Y",
                 "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d")


def _as_date(v):
    """(التاريخ، هل كان نصاً). الخلايا النصية بصيغة يوم/شهر/سنة (dayfirst=True)، وأرقام إكسيل التسلسلية مدعومة."""
    if _blank(v):
        return None, False
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime().date(), False
    if isinstance(v, datetime):
        return v.date(), False
    if isinstance(v, date):
        return v, False
    if isinstance(v, (int, float)):
        if 20000 < v < 80000:
            return date(1899, 12, 30) + timedelta(days=int(v)), False
        return None, False
    s = E.normalize_digits(str(v)).strip()
    for f in _DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date(), True
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).date(), True
    except (ValueError, TypeError):
        return None, True


def detect_swapped_dates(values):
    """إكسيل يحوّل النص «08/09/2025» (يوم/شهر) إلى تاريخ حقيقي بقراءة شهر/يوم ← 2025-08-09 خطأ، ويترك «16/08/2025»
    نصاً لأنه لا يصلح شهر/يوم. العلامة المميزة: الملف فيه تواريخ نصية يوم/شهر، وكل الخلايا المحوّلة يومها ≤ 12.
    عندها نعكس اليوم والشهر في الخلايا المحوّلة. يرجع True لو اكتُشف هذا النمط."""
    has_text = False
    real = []
    for v in values:
        d, was_text = _as_date(v)
        if d is None:
            continue
        if was_text:
            has_text = True
        else:
            real.append(d)
    return has_text and bool(real) and all(d.day <= 12 for d in real)


def parse_hw_date(v, swap=False):
    d, was_text = _as_date(v)
    if d is None:
        return None
    if swap and not was_text and d.day <= 12:
        d = date(d.year, d.day, d.month)
    return d.isoformat()


def _read_sheet(src, sheet, required, label):
    if src is None:
        return None
    try:
        xl = pd.ExcelFile(src)
    except Exception as e:  # noqa: BLE001
        raise UserError(f"تعذر فتح ملف {label}: {e}")
    name = sheet if sheet in xl.sheet_names else None
    if name is None:   # الشيت بالاسم المتوقع غير موجود: أول شيت فيه الأعمدة المطلوبة
        for sh in xl.sheet_names:
            cols = {str(c).strip().lower() for c in xl.parse(sh, nrows=0).columns}
            if all(r.lower() in cols for r in required):
                name = sh
                break
    if name is None:
        raise UserError(f"ملف {label}: لم أجد الشيت «{sheet}» ولا شيتاً يحتوي الأعمدة المطلوبة:\n" + "، ".join(required))
    df = xl.parse(name, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    lower = {c.lower(): c for c in df.columns}
    missing = [r for r in required if r.lower() not in lower]
    if missing:
        raise UserError(f"ملف {label} (شيت {name}): أعمدة ناقصة: " + "، ".join(missing))
    df = df.rename(columns={lower[r.lower()]: r for r in required})
    return df


def read_customer_file(src):
    df = _read_sheet(src, "ورقة1", CUST_COLS, "العملاء")
    if df is None:
        return None
    swap = detect_swapped_dates(df["Trans Date"])
    recs, problems, master, last_bal, order = [], [], {}, {}, []
    for i, r in enumerate(df.to_dict("records"), start=2):
        code = _txt(r["Customer Code"])
        if not code:
            continue
        typ = (_txt(r["Type"]) or "").upper()
        if typ not in CUST_TYPES:
            problems.append(f"العملاء سطر {i}: نوع حركة غير معروف «{r['Type']}» (تم تجاهله).")
            continue
        d = parse_hw_date(r["Trans Date"], swap)
        doc = _txt(r["Doc No"])
        if not d or not doc:
            problems.append(f"العملاء سطر {i}: تاريخ أو رقم مستند غير صالح (تم تجاهله).")
            continue
        cur = (_txt(r["Currency"]) or "SAR").upper()
        if cur != "SAR":
            problems.append(f"العملاء سطر {i}: العملة {cur} — استُخدمت المبالغ بالريال (SAR) من الملف.")
        name = _txt(r["Customer Name"]) or code
        if code not in master:
            order.append(code)
        master.setdefault(code, name)
        last_bal[code] = round(_num(r["Running Balance (SAR)"]), 2)
        recs.append(dict(row=i, code=code, date=d, type=typ, doc=doc, applied=_txt(r["Applied To Doc"]),
                         debit=round(_num(r["Debit (SAR)"]), 2), credit=round(_num(r["Credit (SAR)"]), 2)))
    return dict(recs=recs, master=master, order=order, final=last_bal, problems=problems, swapped=swap)


def read_supplier_file(src):
    df = _read_sheet(src, "Supplier Transactions", SUPP_COLS, "الموردين")
    if df is None:
        return None
    swap = detect_swapped_dates(df["Trans Date"])
    recs, problems, master, last_bal, order, net_change = [], [], {}, {}, [], None
    for i, r in enumerate(df.to_dict("records"), start=2):
        code = _txt(r["Supplier Code"])
        if not code:
            continue
        if code.lower().startswith("net change"):
            net_change = dict(debit=round(_num(r["Debit (SAR)"]), 2), credit=round(_num(r["Credit (SAR)"]), 2),
                              balance=round(_num(r["Balance (SAR)"]), 2))
            continue
        typ = (_txt(r["Type"]) or "").upper()
        if typ not in SUPP_TYPES:
            problems.append(f"الموردين سطر {i}: نوع حركة غير معروف «{r['Type']}» (تم تجاهله).")
            continue
        d = parse_hw_date(r["Trans Date"], swap)
        doc = _txt(r["Doc No"])
        if not d or not doc:
            problems.append(f"الموردين سطر {i}: تاريخ أو رقم مستند غير صالح (تم تجاهله).")
            continue
        name = _txt(r["Supplier Name"]) or code
        if code not in master:
            order.append(code)
            master[code] = (name, _vat_text(r["VAT No"]))
        elif not master[code][1] and _vat_text(r["VAT No"]):
            master[code] = (master[code][0], _vat_text(r["VAT No"]))
        last_bal[code] = round(_num(r["Balance (SAR)"]), 2)
        applied = _txt(r["Applied To Invoice"]) if typ != "PINV" else None
        if typ == "PINV" and _txt(r["Applied To Invoice"]):
            problems.append(f"الموردين سطر {i}: فاتورة شراء {doc} عليها «Applied To Invoice = "
                            f"{_txt(r['Applied To Invoice'])}» (لا معنى له لفاتورة؛ تم تجاهل الربط).")
        recs.append(dict(row=i, src_row=_txt(r["Source Row"]), code=code, date=d, type=typ, doc=doc, applied=applied,
                         debit=round(_num(r["Debit (SAR)"]), 2), credit=round(_num(r["Credit (SAR)"]), 2)))
    return dict(recs=recs, master=master, order=order, final=last_bal, problems=problems, swapped=swap,
                net_change=net_change)


def read_item_history(src):
    df = _read_sheet(src, "Item History", HIST_COLS, "حركات الأصناف")
    if df is None:
        return None
    body = df[df["Type"].map(lambda v: _txt(v) not in (None, HIST_END))]
    swap = detect_swapped_dates(body["Date"])
    moves, problems, items, end_bal, grand = [], [], {}, {}, None
    for i, r in enumerate(df.to_dict("records"), start=2):
        code, typ = _txt(r["Item"]), _txt(r["Type"])
        if not code:
            if typ is None and not _blank(r["Balance"]):
                grand = round(_num(r["Balance"]), 3)   # السطر الإجمالي الأخير
            continue
        items.setdefault(code, (_txt(r["Name"]) or code, _txt(r["Unit"]) or "KG"))
        if typ == HIST_END:
            end_bal[code] = round(_num(r["Balance"]), 3)
            continue
        if not typ:
            continue
        d = parse_hw_date(r["Date"], swap)
        if not d:
            problems.append(f"حركات الأصناف سطر {i}: تاريخ غير صالح «{r['Date']}» (تم تجاهله).")
            continue
        party = _txt(r["Cust/Supp."])
        dup = party == DUP_MARK
        moves.append(dict(row=i, item=code, date=d, type=typ, ref=_txt(r["Reference"]), party=None if dup else party,
                          dup=dup, loc=_txt(r["Location"]), qty=_num(r["Qty"]), balance=_num(r["Balance"])))
    # فحص الرصيد الجاري لكل صنف: مجموع الكميات بترتيب الملف = عمود Balance
    run, bad = defaultdict(float), 0
    for m in moves:
        run[m["item"]] += m["qty"]
        if abs(run[m["item"]] - m["balance"]) > 0.01:
            bad += 1
            run[m["item"]] = m["balance"]
    if bad:
        problems.append(f"حركات الأصناف: {bad} سطر رصيده الجاري لا يساوي مجموع الكميات قبله (اعتُمد رصيد الملف).")
    return dict(moves=moves, items=items, end=end_bal, grand=grand, problems=problems, swapped=swap)


def read_cost_file(src):
    df = _read_sheet(src, "Item Cost Price Analysis", COST_COLS, "التكلفة والرصيد")
    if df is None:
        return None
    out, problems = {}, []
    for i, r in enumerate(df.to_dict("records"), start=2):
        code = _txt(r["Item"])
        if not code:
            continue
        qty, cost = _num(r["In Stock"]), _num(r["Cost Pr."])
        if qty < 0 or cost < 0:
            problems.append(f"ملف التكلفة سطر {i}: رصيد أو تكلفة سالبة للصنف {code} (تم تجاهله).")
            continue
        out[code] = dict(name=_txt(r["Name"]), qty=round(qty, 3), cost=round(cost, 5), last=round(_num(r["Last Price"]), 5))
    return dict(items=out, problems=problems)


# ============================================================
# توزيع السندات على الفواتير (يحاكي Applied To في HansaWorld بما فيها إلغاء التسويات)
# ============================================================
def simulate_allocations(vouchers, caps=None):
    """vouchers بالترتيب الزمني: dict(key, party, date, cash, lines=[(invoice_key, مبلغ ±)]).
    cash = ما يتيحه السند من نقدية للتوزيع (صفر لسند رد أو إعادة توزيع فقط).
    السطر السالب يلغي تسوية سابقة على الفاتورة (الأحدث أولاً) فيعود مبلغها «تحت الحساب» لسندها الأصلي؛
    السطر الموجب يُموَّل من نقدية السند نفسه، ثم من الإشعارات الدائنة التي استهلكها نفس المستند (credits)، ثم مما حرره
    نفس المستند، ثم من أقدم رصيد غير موزع لنفس الطرف. credits = [(مفتاح الإشعار، المبلغ المستهلك)].
    caps {invoice: أقصى ما يمكن تسويته عليها} (الإجمالي بعد الإشعارات الدائنة): ما يزيد يبقى تحت الحساب لسنده.
    يرجع (allocations {(voucher, invoice): [مبلغ، تاريخ]}, مشاكل)."""
    caps = caps or {}
    stacks = defaultdict(list)      # الفاتورة ← [[السند، المبلغ]] بترتيب التسوية
    on_inv = defaultdict(float)
    used = defaultdict(float)
    cash = {}
    party_vouchers = defaultdict(list)
    alloc = {}
    problems = []

    def add(vk, inv, amt, d):
        if amt <= 0.005:
            return
        stacks[inv].append([vk, amt])
        used[vk] += amt
        on_inv[inv] += amt
        cur = alloc.setdefault((vk, inv), [0.0, d])
        cur[0] = round(cur[0] + amt, 2)
        cur[1] = d

    for v in vouchers:
        vk = v["key"]
        cash[vk] = round(max(v["cash"], 0.0), 2)
        party_vouchers[v["party"]].append(vk)
        credits = []
        for ck, camt in v.get("credits", ()):
            cash[ck] = round(cash.get(ck, 0.0) + max(camt, 0.0), 2)
            credits.append(ck)
        freed = []
        for inv, amt in v["lines"]:
            if amt >= -0.005:
                continue
            need = round(-amt, 2)
            st = stacks[inv]
            while need > 0.005 and st:
                ok, a = st[-1]
                take = round(min(a, need), 2)
                st[-1][1] = round(a - take, 2)
                if st[-1][1] <= 0.005:
                    st.pop()
                used[ok] -= take
                on_inv[inv] -= take
                alloc[(ok, inv)][0] = round(alloc[(ok, inv)][0] - take, 2)
                freed.append([ok, take])
                need = round(need - take, 2)
            if need > 0.005:
                problems.append(f"المستند {v['doc']}: سطر عكسي {fmt(need)} على {inv[1]} بلا تسوية سابقة يلغيها "
                                "(رد/تسوية عكسية) — أثره على رصيد الطرف مسجل، والفاتورة لا يزيد متبقيها عن إجماليها.")
        for inv, amt in v["lines"]:
            if amt <= 0.005:
                continue
            need = round(amt, 2)
            take = round(min(cash[vk] - used[vk], need), 2)
            if take > 0.005:
                add(vk, inv, take, v["date"])
                need = round(need - take, 2)
            for ck in credits:
                if need <= 0.005:
                    break
                take = round(min(cash[ck] - used[ck], need), 2)
                if take > 0.005:
                    add(ck, inv, take, v["date"])
                    need = round(need - take, 2)
            for f in freed:
                if need <= 0.005:
                    break
                take = round(min(f[1], need), 2)
                if take > 0.005:
                    add(f[0], inv, take, v["date"])
                    f[1] = round(f[1] - take, 2)
                    need = round(need - take, 2)
            for ok in party_vouchers[v["party"]]:
                if need <= 0.005:
                    break
                spare = round(cash[ok] - used[ok], 2)
                take = round(min(spare, need), 2)
                if take > 0.005:
                    add(ok, inv, take, v["date"])
                    need = round(need - take, 2)
            if need > 0.005:
                problems.append(f"المستند {v['doc']}: {fmt(need)} على الفاتورة {inv[1]} سُدد في HansaWorld من مصدر غير نقدي "
                                "(مثل إشعار دائن تحت الحساب) — بقي مفتوحاً على الفاتورة ويقابله الإشعار الدائن في رصيد الطرف.")
    # HansaWorld يسمح بتسوية زائدة مؤقتة تصححها سطور عكسية لاحقة، فالسقف يُطبق على الحالة النهائية فقط:
    # ما يزيد عن المتاح للفاتورة يُرفع من أحدث تسوية ويبقى تحت الحساب في سنده.
    for inv, cap in caps.items():
        extra = round(on_inv[inv] - cap, 2)
        if extra <= 0.005:
            continue
        problems.append(f"الفاتورة {inv[1]}: سُوّي عليها في HansaWorld {fmt(extra)} أكثر من إجماليها — "
                        "بقي الزائد تحت الحساب في السند.")
        st = stacks[inv]
        while extra > 0.005 and st:
            ok, a = st[-1]
            take = round(min(a, extra), 2)
            st[-1][1] = round(a - take, 2)
            if st[-1][1] <= 0.005:
                st.pop()
            alloc[(ok, inv)][0] = round(alloc[(ok, inv)][0] - take, 2)
            on_inv[inv] -= take
            extra = round(extra - take, 2)
    return {k: v for k, v in alloc.items() if v[0] > 0.005}, problems


# ============================================================
# الشركاء: مطابقة وإعادة ترقيم وفصل العملاء عن الموردين
# ============================================================
# جداول تحمل كود الشريك بدون مستندات (المستندات نفسها تمنع إعادة الترقيم): تُنقل مع الكود
PARTNER_SOFT_COLUMNS = [("gl_lines", "partner_code")]


def supplier_code_for(hw_code):
    m = re.fullmatch(r"[Cc](\d+)", hw_code or "")
    if m:
        return "S" + m.group(1)
    if re.fullmatch(r"[Ss]\d+", hw_code or ""):
        return hw_code.upper()
    return "S-" + hw_code


def _rename_partners(conn, moves):
    """moves {الكود الحالي: الكود الجديد} (قد تتبادل الأكواد فيما بينها). سجل فترات المناديب مرتبط بمفتاح أجنبي،
    فيُرفع جانباً ثم يُعاد بالأكواد الجديدة بعد إعادة الترقيم (بدون أي مخالفة للمفاتيح الأجنبية)."""
    if not moves:
        return
    olds = list(moves)
    ph = ",".join("?" * len(olds))
    hist = conn.execute(f"SELECT partner_code, rep_name, start_date, end_date, changed_by, changed_at FROM partner_rep_history "
                        f"WHERE partner_code IN ({ph})", olds).fetchall()
    conn.execute(f"DELETE FROM partner_rep_history WHERE partner_code IN ({ph})", olds)
    tmp = {old: f"__HW_TMP_{i}" for i, old in enumerate(olds)}
    for step in (tmp, {tmp[o]: moves[o] for o in olds}):
        for a, b in step.items():
            conn.execute("UPDATE business_partners SET partner_code=? WHERE partner_code=?", (b, a))
            for table, col in PARTNER_SOFT_COLUMNS:
                try:
                    conn.execute(f"UPDATE {table} SET {col}=? WHERE {col}=?", (b, a))
                except E.sqlite3.OperationalError:
                    pass
    conn.executemany("INSERT INTO partner_rep_history (partner_code, rep_name, start_date, end_date, changed_by, changed_at) "
                     "VALUES (?,?,?,?,?,?)", [(moves[h[0]], *h[1:]) for h in hist])


def sync_partners(conn, customers, suppliers, stats, problems):
    """customers {كود: اسم}، suppliers {كود: (اسم، رقم ضريبي)} ← {('customer'|'supplier', كود HW): كود النظام}."""
    rows = conn.execute("SELECT partner_code, name, is_customer, is_supplier, COALESCE(legacy_code,''), "
                        "COALESCE(vat_number,''), phone, email, address, cr_number, rep_name FROM business_partners").fetchall()
    db_p = {r[0]: dict(code=r[0], name=r[1], norm=norm_name(r[1]), c=bool(r[2]), s=bool(r[3]), legacy=r[4],
                       vat=E.normalize_digits(r[5]).strip(), extra=r[6:]) for r in rows}
    by_name = defaultdict(list)
    for p in db_p.values():
        by_name[p["norm"]].append(p["code"])
    entities = [("customer", k, k, n, None) for k, n in customers.items()] + \
               [("supplier", k, supplier_code_for(k), n, vat) for k, (n, vat) in suppliers.items()]
    claimed = {}          # كود النظام الحالي ← الكيان
    claim_of = {}         # الكيان (role, hw) ← كود النظام الحالي
    flag = {"customer": "c", "supplier": "s"}
    # 1) نفس الكود والاسم/المرجع  2) الكود القديم  3) الاسم الموحد  4) الرقم الضريبي
    passes = [
        lambda e: [e[2]] if e[2] in db_p and (db_p[e[2]]["norm"] == norm_name(e[3]) or db_p[e[2]]["legacy"] == e[1]) else [],
        lambda e: [p["code"] for p in db_p.values() if p["legacy"] == e[1] and p[flag[e[0]]]],
        lambda e: sorted(by_name.get(norm_name(e[3]), []), key=lambda c: (not db_p[c][flag[e[0]]], c)),
        lambda e: [p["code"] for p in db_p.values() if e[4] and len(e[4]) >= 10 and p["vat"] == e[4]],
    ]
    for pick in passes:
        for e in entities:
            key = (e[0], e[1])
            if key in claim_of:
                continue
            for code in pick(e):
                if code not in claimed:
                    claimed[code] = e
                    claim_of[key] = code
                    break
    target = {(e[0], e[1]): e[2] for e in entities}
    wanted = set(target.values())
    renames = {old: target[key] for key, old in claim_of.items() if old != target[key]}
    evict = {code: f"OLD-{code}" for code in db_p if code in wanted and code not in claimed}
    for code in evict:
        n = 1
        while evict[code] in db_p or evict[code] in wanted:
            n += 1
            evict[code] = f"OLD{n}-{code}"
    blocked = [(old, db_p[old]["name"], E.partner_reference_count(conn, old)) for old in list(renames) + list(evict)]
    blocked = [b for b in blocked if b[2]]
    if blocked:
        raise UserError("لا يمكن إعادة ترقيم شركاء عليهم مستندات في النظام (يجب أن يأخذوا أكواد HansaWorld):\n" +
                        "\n".join(f"• {c} - {n} ({k} مستند)" for c, n, k in blocked[:10]) +
                        ("\n..." if len(blocked) > 10 else "") +
                        "\n\nنفّذ «تصفير الحركات» أولاً (يحتفظ بالشركاء والأصناف) ثم أعد الاستيراد.")
    _rename_partners(conn, {**renames, **evict})
    stats["partners_recoded"] = len(renames)
    stats["partners_evicted"] = len(evict)
    if evict:
        problems.append(f"{len(evict)} شريك قديم بالنظام غير موجود في ملفات HansaWorld كان يشغل كوداً مطلوباً، "
                        "فنُقل لكود يبدأ بـ OLD- (لم يُحذف): " + "، ".join(f"{k}→{v}" for k, v in list(evict.items())[:8]))
    # الكيانات بلا شريك مطابق تأخذ بيانات الاتصال من نسختها في الدور الآخر (نفس الاسم) إن وُجدت
    sibling_extra = {}
    for key, old in claim_of.items():
        sibling_extra[norm_name(db_p[old]["name"])] = db_p[old]["extra"]
    result = {}
    created = 0
    for role, hw, new, name, vat in entities:
        is_c, is_s = (1, 0) if role == "customer" else (0, 1)
        if (role, hw) in claim_of:
            conn.execute("UPDATE business_partners SET name=?, is_customer=?, is_supplier=?, legacy_code=?, "
                         "vat_number=COALESCE(?, vat_number) WHERE partner_code=?", (name, is_c, is_s, hw, vat, new))
        else:
            ph, em, ad, cr, rep = sibling_extra.get(norm_name(name), (None,) * 5)
            conn.execute("INSERT INTO business_partners (partner_code, name, phone, email, address, cr_number, vat_number, "
                         "is_customer, is_supplier, legacy_code, rep_name) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                         (new, name, ph, em, ad, cr, vat, is_c, is_s, hw, rep))
            created += 1
        result[(role, hw)] = new
    stats["partners_created"] = created
    stats["customers"] = len(customers)
    stats["suppliers"] = len(suppliers)
    for seq, prefix in (("customer", "C"), ("supplier", "S")):
        row = conn.execute("SELECT prefix FROM sequences WHERE name=?", (seq,)).fetchone()
        if row and row[0] == prefix:
            conn.execute("UPDATE sequences SET next_value=MAX(next_value, ?) WHERE name=?",
                         (E.max_used_serial(conn, seq, prefix) + 1, seq))
    return result


# ============================================================
# الأصناف والتكلفة
# ============================================================
def sync_items(conn, hist, costs, stats, problems):
    items = dict(hist["items"]) if hist else {}
    cost = costs["items"] if costs else {}
    for code, c in cost.items():
        items.setdefault(code, (c["name"] or code, "KG"))
    existing = {r[0] for r in conn.execute("SELECT item_code FROM inventory")}
    created = 0
    for code, (name, unit) in items.items():
        name = E.clean_name(name) or code
        if code in existing:
            conn.execute("UPDATE inventory SET item_name=?, unit=?, legacy_code=COALESCE(legacy_code, ?) WHERE item_code=?",
                         (name, unit or "KG", code, code))
        else:
            conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, cost_price, "
                         "legacy_code) VALUES (?,?,?,0,0,0,0,?)", (code, name, unit or "KG", code))
            created += 1
    if costs:
        for code in items:
            c = cost.get(code)
            if c:
                conn.execute("UPDATE inventory SET opening_qty=?, opening_cost=?, last_purchase_price=? WHERE item_code=?",
                             (c["qty"], c["cost"], c["last"], code))
            else:   # صنف بلا رصيد في ملف التكلفة ← رصيده صفر في HansaWorld
                conn.execute("UPDATE inventory SET opening_qty=0 WHERE item_code=?", (code,))
        stats["items_stock"] = sum(1 for c in cost.values() if c["qty"] > 0)
    E.recalc_items(conn, list(items), allow_negative=True)
    stats["items"] = len(items)
    stats["items_created"] = created


# ============================================================
# المستندات
# ============================================================
def _split_vat(total):
    vat = round(total * VAT_RATE / (1 + VAT_RATE), 2)
    return round(total - vat, 2), vat


def _mapped(conn, role):
    return {(t, d, p): (tg, tid) for t, d, p, tg, tid in
            conn.execute("SELECT doc_type, doc_no, party_code, target, target_id FROM hw_import_map WHERE role=?", (role,))}


def _map(conn, role, typ, doc, hw_code, target, tid, d, amount):
    conn.execute("INSERT INTO hw_import_map (role, doc_type, doc_no, party_code, target, target_id, doc_date, amount) "
                 "VALUES (?,?,?,?,?,?,?,?)", (role, typ, doc, hw_code, target, tid, d, amount))


def _cash_account(conn, name):
    """حساب الخزينة للسندات المستوردة: المختار، وإلا حساب «وسيط ترحيل HansaWorld» (يُنشأ تحت النقدية بالبنوك)."""
    names = E.cash_account_names(conn)
    if name and name != HW_CASH_ACCOUNT:
        if name not in names:
            raise UserError(f"حساب الخزينة «{name}» غير معرّف.")
        return name
    if HW_CASH_ACCOUNT not in names:
        parent = HW_CASH_PARENT if conn.execute("SELECT 1 FROM chart_of_accounts WHERE account_code=? AND is_group=1",
                                                (HW_CASH_PARENT,)).fetchone() else None
        if parent is None:
            raise UserError(f"الحساب التجميعي {HW_CASH_PARENT} (النقدية بالبنوك) غير موجود بالدليل؛ اختر حساب خزينة للسندات.")
        code = E.next_child_account_code(conn, parent)
        conn.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                     "VALUES (?,?,?,?,0)", (code, HW_CASH_ACCOUNT, E.TYPE_ASSET, parent))
    return HW_CASH_ACCOUNT


def find_doc(mapping, ref):
    """البحث عن مستند بالرقم كما هو، ثم بالمرجع الأساسي قبل «/» (مثل 10001230/1 ← 10001230). يرجع (الرقم، القيمة)."""
    if not ref:
        return None, None
    if ref in mapping:
        return ref, mapping[ref]
    base = ref.split("/")[0].strip()
    return (base, mapping[base]) if base and base != ref and base in mapping else (None, None)


def import_party_docs(conn, role, data, codes, cash_account, stats, problems, voucher_book):
    """ترحيل مستندات العملاء أو الموردين بالترتيب الزمني. يُرجع خريطة الفواتير {رقم المستند: (inv_id, كود HW)}."""
    is_c = role == "customer"
    kind = "sales" if is_c else "purchase"
    inv_types = ("INV", "CASH") if is_c else ("PINV",)
    pay_type = "REC" if is_c else "PAYM"
    c = E.FLOW[kind]
    done = _mapped(conn, role)
    names = {code: conn.execute("SELECT name FROM business_partners WHERE partner_code=?", (codes[(role, code)],)).fetchone()[0]
             for code in data["master"]}
    recs = sorted(data["recs"], key=lambda r: (r["date"], r["row"]))
    invoices = {}
    creds = {}     # الإشعارات الدائنة: رقم ← (discount_id، كود HW، تحت الحساب؟)
    for (t, d, p), (tg, tid) in done.items():
        if t in inv_types:
            invoices[d] = (tid, p)
        elif t == "CRED" and tid:
            row = conn.execute("SELECT invoice_id FROM post_discounts WHERE discount_id=?", (tid,)).fetchone()
            creds[d] = (tid, p, row is not None and row[0] is None)
    # 1) الفواتير
    for r in recs:
        if r["type"] not in inv_types:
            continue
        key = (r["type"], r["doc"], r["code"])
        total = round(r["debit"] if is_c else r["credit"], 2)
        if key in done:
            stats["skipped_existing"] += 1
            continue
        if total <= 0:
            problems.append(f"{r['type']} {r['doc']} ({r['code']}) بقيمة صفر أو سالبة — تم تجاهله.")
            continue
        if conn.execute(f"SELECT 1 FROM {c['inv']} WHERE invoice_no=?", (r["doc"],)).fetchone():
            problems.append(f"رقم الفاتورة {r['doc']} مستخدم لفاتورة أخرى بالنظام (ليست من HansaWorld) — لم تُستورد.")
            continue
        sub, vat = _split_vat(total)
        code = codes[(role, r["code"])]
        note = f"مستورد من HansaWorld: {r['type']} {r['doc']}"
        if is_c:
            iid = conn.execute("INSERT INTO sales_invoices (invoice_date, customer_code, customer_name, subtotal, vat_amount, "
                               "total_amount, paid_amount, status, notes, invoice_no, source_type) "
                               "VALUES (?,?,?,?,?,?,0,'مفتوحة',?,?,?)",
                               (r["date"], code, names[r["code"]], sub, vat, total, note, r["doc"], HW_SOURCE)).lastrowid
        else:
            iid = conn.execute("INSERT INTO purchase_invoices (purchase_date, supplier_code, supplier_name, subtotal, vat_amount, "
                               "total_amount, paid_amount, status, notes, invoice_no, source_type, supplier_ref) "
                               "VALUES (?,?,?,?,?,?,0,'مفتوحة',?,?,?,?)",
                               (r["date"], code, names[r["code"]], sub, vat, total, note, r["doc"], HW_SOURCE,
                                r["doc"])).lastrowid
        _map(conn, role, r["type"], r["doc"], r["code"], c["inv"], iid, r["date"], total)
        invoices[r["doc"]] = (iid, r["code"])
        stats["sales_invoices" if is_c else "purchase_invoices"] += 1
    # 2) الإشعارات الدائنة / خصومات الموردين
    for r in recs:
        if r["type"] != "CRED":
            continue
        key = ("CRED", r["doc"], r["code"])
        total = round((r["credit"] - r["debit"]) if is_c else (r["debit"] - r["credit"]), 2)
        if key in done:
            stats["skipped_existing"] += 1
            continue
        if total <= 0:
            problems.append(f"إشعار {r['doc']} ({r['code']}) بقيمة صفر أو عكسية ({fmt(total)}) — تم تجاهله.")
            continue
        inv_id = None
        if r["applied"]:
            _, hit = find_doc(invoices, r["applied"])
            if hit and hit[1] == r["code"]:
                inv_id = hit[0]
            else:
                problems.append(f"إشعار {r['doc']}: الفاتورة المربوطة {r['applied']} غير موجودة لنفس الطرف — سُجّل تحت الحساب.")
        sub, vat = _split_vat(total)
        no = r["doc"] if not conn.execute("SELECT 1 FROM post_discounts WHERE discount_no=?", (r["doc"],)).fetchone() \
            else f"HW-{'C' if is_c else 'S'}-{r['doc']}"
        did = conn.execute("INSERT INTO post_discounts (discount_no, kind, discount_date, party_code, party_name, invoice_id, "
                           "method, subtotal, vat_amount, total_amount, reason, notes, created_by) "
                           "VALUES (?,?,?,?,?,?,'pl',?,?,?,?,?,?)",
                           (no, kind, r["date"], codes[(role, r["code"])], names[r["code"]], inv_id, sub, vat, total,
                            "إشعار دائن HansaWorld" if is_c else "إشعار خصم مورد HansaWorld",
                            f"مستورد من HansaWorld: CRED {r['doc']}" + (f" على {r['applied']}" if r["applied"] else ""),
                            E.current_user_name())).lastrowid
        E.post_discount_journal(conn, did)
        _map(conn, role, "CRED", r["doc"], r["code"], "post_discounts", did, r["date"], total)
        creds[r["doc"]] = (did, r["code"], inv_id is None)
        stats["sales_credits" if is_c else "supplier_credits"] += 1
    # «تحت الحساب» يُحدَّد من الملف المصدر (بلا Applied To لفاتورة نفس الطرف)، لا من قاعدة البيانات:
    # فعند إعادة الاستيراد يكون الإشعار مقسّماً ومربوطاً بفواتير من التشغيل السابق.
    for r in recs:
        if r["type"] == "CRED" and r["doc"] in creds:
            _, hit = find_doc(invoices, r["applied"])
            creds[r["doc"]] = (creds[r["doc"]][0], creds[r["doc"]][1], not (hit and hit[1] == r["code"]))
    # 3) السندات: سند واحد لكل رقم مستند (قد يُوزع على عدة فواتير)
    groups = {}
    for r in recs:
        if r["type"] == pay_type or (is_c and r["type"] == "CASH"):
            gk = (r["type"], r["doc"], r["code"])
            g = groups.setdefault(gk, dict(type=r["type"], doc=r["doc"], code=r["code"], date=r["date"], row=r["row"],
                                           lines=defaultdict(float), credits=defaultdict(float), net=0.0))
            g["date"] = min(g["date"], r["date"])
            if r["type"] == "CASH":   # فاتورة نقدية: تحصيل فوري بقيمتها على نفسها
                g["net"] += r["credit"]
                g["lines"][r["doc"]] += r["credit"]
                continue
            amt = (r["credit"] - r["debit"]) if is_c else (r["debit"] - r["credit"])
            g["net"] += amt
            target = r["applied"]
            if target:
                inv_no, hit = find_doc(invoices, target)
                cred_no, chit = find_doc(creds, target)
                if hit and hit[1] == r["code"]:
                    g["lines"][inv_no] += amt
                elif chit and chit[1] == r["code"] and chit[2] and amt < 0:
                    # سطر عكسي على إشعار دائن تحت الحساب = استخدام الإشعار في سداد فواتير نفس المستند
                    g["credits"][cred_no] += -amt
                else:
                    problems.append(f"{r['type']} {r['doc']}: الفاتورة المسددة {target} غير موجودة لنفس الطرف "
                                    f"— المبلغ {fmt(amt)} بقي تحت الحساب.")
    for gk, g in sorted(groups.items(), key=lambda kv: (kv[1]["date"], kv[1]["row"])):
        net = round(g["net"], 2)
        key = ("CASH-REC" if g["type"] == "CASH" else g["type"], g["doc"], g["code"])
        code = codes[(role, g["code"])]
        tid = done.get(key, (None, None))[1]
        if key in done:
            stats["skipped_existing"] += 1
        elif abs(net) > 0.005:
            in_cash = net > 0
            t_type = ("قبض" if in_cash else "صرف") if is_c else ("صرف" if in_cash else "قبض")
            applied = [k for k, v in g["lines"].items() if abs(v) > 0.005]
            desc = (f"HansaWorld {g['type']} {g['doc']}" + (f" — على: {'، '.join(applied[:6])}" if applied else "") +
                    ("..." if len(applied) > 6 else ""))
            tid = conn.execute("INSERT INTO cash_transactions (trans_type, trans_date, account_bank, category, party_code, "
                               "party_name, amount, alloc_type, alloc_id, description) VALUES (?,?,?,?,?,?,?,NULL,NULL,?)",
                               (t_type, g["date"], cash_account, CAT_PARTNER, code, names[g["code"]], abs(net),
                                desc)).lastrowid
            _map(conn, role, key[0], g["doc"], g["code"], "cash_transactions", tid, g["date"], net)
            if in_cash:
                stats["receipts" if is_c else "payments"] += 1
            else:
                stats["customer_refunds" if is_c else "supplier_refunds"] += 1
        else:   # إعادة توزيع فقط (لا نقدية): يُسجَّل للمطابقة ومنع التكرار
            _map(conn, role, key[0], g["doc"], g["code"], "reallocation", None, g["date"], 0.0)
        voucher_book.append(dict(key=(kind, g["type"], g["doc"], g["code"]), tid=tid, party=(kind, g["code"]),
                                 date=g["date"], doc=g["doc"], cash=max(net, 0.0),
                                 lines=sorted(((("inv", d), round(a, 2)) for d, a in g["lines"].items()),
                                              key=lambda x: x[1]), invoices=invoices, kind=kind, creds=creds,
                                 credits=[(("CREDIT", d), round(a, 2)) for d, a in g["credits"].items() if a > 0.005]))
    return invoices


def write_allocations(conn, voucher_book, stats, problems):
    """إعادة كتابة توزيع كل سندات HansaWorld (حتمي وقابل للتكرار) ثم تحديث المسدد وحالة الفواتير."""
    for kind in ("sales", "purchase"):
        vb = [v for v in voucher_book if v["kind"] == kind]
        if not vb:
            continue
        vb.sort(key=lambda v: v["date"])   # ترتيب ثابت: نفس اليوم بترتيب سطور الملف
        invs = vb[0]["invoices"]
        creds = vb[0]["creds"]
        role = "customer" if kind == "sales" else "supplier"
        undo_credit_splits(conn, role, kind)      # يُعاد بناء التقسيم من الصفر في كل تشغيل (حتمي بلا تكرار)
        tids = {v["key"]: v["tid"] for v in vb}
        hw_tids = [t for t in tids.values() if t]
        hw_set = set(hw_tids)
        c = E.FLOW[kind]
        caps = {}
        for doc, (iid, _) in invs.items():   # المتاح للتسوية = الإجمالي - الإشعارات - مدفوعات غير مستوردة
            tot = conn.execute(f"SELECT total_amount FROM {c['inv']} WHERE {c['iid']}=?", (iid,)).fetchone()[0]
            direct = conn.execute("SELECT COALESCE(SUM(amount),0) FROM cash_transactions WHERE alloc_type=? AND alloc_id=?",
                                  (kind, iid)).fetchone()[0]
            other = sum(a for t, a in conn.execute("SELECT trans_id, amount FROM advance_allocations WHERE kind=? AND "
                                                   "invoice_id=?", (kind, iid)) if t not in hw_set)
            caps[(kind, doc)] = round(tot - E.invoice_discount_total(conn, kind, iid) - direct - other, 2)
        sim, probs = simulate_allocations([dict(key=v["key"], party=v["party"], date=v["date"], cash=v["cash"], doc=v["doc"],
                                                lines=[((kind, d), a) for (_, d), a in v["lines"]],
                                                credits=v["credits"]) for v in vb], caps)
        problems.extend(probs)
        touched = set()
        for i in range(0, len(hw_tids), 500):
            chunk = hw_tids[i:i + 500]
            for (k, iid) in conn.execute(f"SELECT kind, invoice_id FROM advance_allocations WHERE trans_id IN "
                                         f"({','.join('?' * len(chunk))})", chunk).fetchall():
                touched.add(iid)
            conn.execute(f"DELETE FROM advance_allocations WHERE trans_id IN ({','.join('?' * len(chunk))})", chunk)
        n = 0
        splits = defaultdict(list)
        for (vk, (_, doc)), (amt, d) in sim.items():
            if vk[0] == "CREDIT":
                splits[vk[1]].append((doc, amt))
                continue
            tid = tids.get(vk)
            if not tid:
                problems.append(f"توزيع {fmt(amt)} على الفاتورة {doc} بلا سند نقدي مقابل — تم تجاهله.")
                continue
            iid = invs[doc][0]
            conn.execute("INSERT INTO advance_allocations (trans_id, kind, invoice_id, amount, alloc_date, created_by) "
                         "VALUES (?,?,?,?,?,?)", (tid, kind, iid, amt, d, "HansaWorld"))
            touched.add(iid)
            n += 1
        stats["allocations"] += n
        apply_credit_splits(conn, role, kind, splits, creds, invs, stats, problems)
        for iid, _ in invs.values():
            touched.add(iid)
        for iid in touched:
            E.refresh_invoice_paid(conn, kind, iid)


def _discount_set(conn, did, no, inv_id, total, vat):
    conn.execute("UPDATE post_discounts SET discount_no=?, invoice_id=?, subtotal=?, vat_amount=?, total_amount=? "
                 "WHERE discount_id=?", (no, inv_id, round(total - vat, 2), vat, total, did))
    E.post_discount_journal(conn, did)


def _discount_drop(conn, did):
    j = conn.execute("SELECT journal_entry_id FROM post_discounts WHERE discount_id=?", (did,)).fetchone()
    if j and j[0]:
        conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (j[0],))
        conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (j[0],))
    conn.execute("DELETE FROM post_discounts WHERE discount_id=?", (did,))


def undo_credit_splits(conn, role, kind):
    """إرجاع كل إشعار دائن قسّمه الاستيراد إلى أصله: رقمه وقيمته الكاملة تحت الحساب (ثم يُعاد التقسيم)."""
    parts = conn.execute("SELECT map_id, doc_no, party_code, target_id, target FROM hw_import_map WHERE role=? "
                         "AND doc_type='CRED-SPLIT'", (role,)).fetchall()
    bases = set()
    for map_id, no, party, did, target in parts:
        if target != "post_discounts-base":   # السطر الأصلي نفسه لا يُحذف بل يُعاد لقيمته ورقمه
            _discount_drop(conn, did)
        conn.execute("DELETE FROM hw_import_map WHERE map_id=?", (map_id,))
        bases.add((no.split("/")[0], party))
    for base, party in bases:
        r = conn.execute("SELECT target_id, amount FROM hw_import_map WHERE role=? AND doc_type='CRED' AND doc_no=? "
                         "AND party_code=?", (role, base, party)).fetchone()
        if not r:
            continue
        did, total = r
        no = conn.execute("SELECT discount_no FROM post_discounts WHERE discount_id=?", (did,)).fetchone()[0]
        _discount_set(conn, did, no.split("/")[0], None, round(total, 2), _split_vat(total)[1])


def apply_credit_splits(conn, role, kind, splits, creds, invs, stats, problems):
    """splits {رقم الإشعار: [(رقم الفاتورة، المبلغ)]}: الإشعار الدائن المستخدم في سداد عدة فواتير يُقسَّم لإشعارات فرعية
    «رقم/1، رقم/2...» كل منها مربوط بفاتورته (بترتيب تاريخ الفاتورة). الباقي غير المستخدم يبقى تحت الحساب برقم الإشعار
    الأصلي، ولو استُهلك كله يصبح سطره الأصلي هو «/1». مجموع القيمة والضريبة يساوي الإشعار الأصلي تماماً."""
    c = E.FLOW[kind]
    for base, pieces in splits.items():
        did, hw_code, _ = creds[base]
        row = conn.execute("SELECT discount_no, discount_date, party_code, party_name, total_amount, vat_amount, reason "
                           "FROM post_discounts WHERE discount_id=?", (did,)).fetchone()
        no, d, pcode, pname, total, vat, reason = row
        dates = {doc: conn.execute(f"SELECT {c['idate']} FROM {c['inv']} WHERE {c['iid']}=?", (invs[doc][0],)).fetchone()[0]
                 for doc, _ in pieces}
        pieces = sorted(((doc, round(a, 2)) for doc, a in pieces if a > 0.005), key=lambda x: (dates[x[0]], x[0]))
        used = round(sum(a for _, a in pieces), 2)
        if used > total + 0.005:
            problems.append(f"الإشعار {no}: الموزع منه ({fmt(used)}) أكبر من قيمته ({fmt(total)}) — لم يُقسَّم.")
            continue
        rest = round(total - used, 2)
        vats = [_split_vat(a)[1] for _, a in pieces]
        if rest > 0.005:
            rest_vat = round(vat - sum(vats), 2)
        else:
            vats[-1] = round(vat - sum(vats[:-1]), 2)   # آخر جزء يأخذ فرق تقريب الضريبة
        out = []
        for i, ((doc, amt), v) in enumerate(zip(pieces, vats), start=1):
            pno = f"{no}/{i}"
            if i == 1 and rest <= 0.005:
                _discount_set(conn, did, pno, invs[doc][0], amt, v)
                pid = did
            else:
                pid = conn.execute("INSERT INTO post_discounts (discount_no, kind, discount_date, party_code, party_name, "
                                   "invoice_id, method, subtotal, vat_amount, total_amount, reason, notes, created_by) "
                                   "VALUES (?,?,?,?,?,?,'pl',?,?,?,?,?,?)",
                                   (pno, kind, d, pcode, pname, invs[doc][0], round(amt - v, 2), v, amt, reason,
                                    f"جزء من الإشعار {no} استُخدم في سداد الفاتورة {doc} (HansaWorld)",
                                    "HansaWorld")).lastrowid
                E.post_discount_journal(conn, pid)
                _map(conn, role, "CRED-SPLIT", pno, hw_code, "post_discounts", pid, d, amt)
            out.append(f"{pno} = {fmt(amt)} على {doc}")
        if rest > 0.005:
            _discount_set(conn, did, no, None, rest, rest_vat)
            out.append(f"الباقي {fmt(rest)} تحت الحساب برقم {no}")
        elif len(pieces) >= 1:
            _map(conn, role, "CRED-SPLIT", f"{no}/1#base", hw_code, "post_discounts-base", did, d, pieces[0][1])
        stats["credit_splits"] += len(pieces)
        problems.append(f"الإشعار الدائن {no} ({fmt(total)}) استُخدم في HansaWorld لسداد فواتير، فقُسّم: " + "، ".join(out) + ".")


def import_item_history(conn, hist, cust, supp, codes, inv_maps, stats):
    """حفظ حركات الأصناف التفصيلية مع تحديد الطرف وربط الحركة بمستندها حيثما أمكن."""
    name_to = {"customer": {}, "supplier": {}}
    if cust:
        for k, n in cust["master"].items():
            name_to["customer"].setdefault(norm_name(n), k)
    if supp:
        for k, (n, _) in supp["master"].items():
            name_to["supplier"].setdefault(norm_name(n), k)
    docs = {"customer": defaultdict(list), "supplier": defaultdict(list)}   # (كود HW، التاريخ) ← [(النوع، الرقم)]
    by_no = {"customer": {}, "supplier": {}}
    for role, data in (("customer", cust), ("supplier", supp)):
        if not data:
            continue
        for r in data["recs"]:
            if r["type"] in ("INV", "CASH", "PINV", "CRED"):
                docs[role][(r["code"], r["date"])].append((r["type"], r["doc"]))
                by_no[role][(r["type"], r["doc"])] = r["code"]
    # «مكرر»: الطرف يُستنتج من حركة أخرى بنفس المرجع والنوع
    ref_party = {}
    for m in hist["moves"]:
        if m["party"] and m["ref"]:
            ref_party.setdefault((m["type"], m["ref"]), m["party"])
    seen = defaultdict(int)
    linked = inserted = 0
    for m in hist["moves"]:
        role, link_types = HIST_LINKS.get(m["type"], (None, ()))
        party_text = m["party"] or (ref_party.get((m["type"], m["ref"])) if m["dup"] else None)
        hw = name_to.get(role, {}).get(norm_name(party_text)) if role and party_text else None
        party_code = codes.get((role, hw)) if hw else None
        l_type = l_no = method = None
        if hw:
            for t in link_types:
                if m["ref"] and by_no[role].get((t, m["ref"])) == hw:
                    l_type, l_no, method = t, m["ref"], "رقم المرجع + الطرف"
                    break
            if not l_no:
                cands = [d for d in docs[role].get((hw, m["date"]), []) if d[0] in link_types]
                if len(cands) == 1:
                    (l_type, l_no), method = cands[0], "الطرف + التاريخ (مستند وحيد)"
        base = f"{m['item']}|{m['date']}|{m['type']}|{m['ref'] or ''}|{m['qty']:.4f}|{m['balance']:.4f}"
        seen[base] += 1
        cur = conn.execute("INSERT OR IGNORE INTO hw_item_history (row_key, item_code, item_name, unit, move_date, move_type, "
                           "reference, party_text, party_code, is_duplicate, location, qty, balance, link_doc_type, "
                           "link_doc_no, link_method, src_row) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                           (f"{base}#{seen[base]}", m["item"], hist["items"].get(m["item"], ("", ""))[0],
                            hist["items"].get(m["item"], ("", "KG"))[1], m["date"], m["type"], m["ref"], party_text,
                            party_code, 1 if m["dup"] else 0, m["loc"], m["qty"], m["balance"], l_type, l_no, method,
                            m["row"]))
        if cur.rowcount:
            inserted += 1
            linked += 1 if l_no else 0
        else:
            stats["skipped_existing"] += 1
    stats["history_rows"] = inserted
    stats["history_linked"] = linked


# ============================================================
# تقرير المطابقة
# ============================================================
RECON_COLUMNS = [("القسم", "t"), ("البند", "t"), ("HansaWorld", "m"), ("النظام", "m"), ("الفرق", "m"), ("الحالة", "t")]


def _row(section, label, hw, sysv, tol=0.01, kind="m"):
    diff = round((sysv or 0) - (hw or 0), 3)
    return (section, label, hw, sysv, diff, "✅ مطابق" if abs(diff) <= tol else "❌ فرق")


def build_reconciliation(conn, cust, supp, hist, costs, codes):
    """يرجع (Report الملخص، {اسم: DataFrame} للتفاصيل). يقارن ما في قاعدة البيانات الآن بأرقام HansaWorld."""
    rows, details = [], {}
    bal = {b["code"]: b for b in E.partner_balances("9999-12-31", conn=conn)}
    for role, data, label in (("customer", cust, "العملاء"), ("supplier", supp, "الموردون")):
        if not data:
            continue
        kind = "sales" if role == "customer" else "purchase"
        c = E.FLOW[kind]
        sign = 1 if role == "customer" else -1      # رصيد المورد في النظام سالب (دائن)
        det, bad = [], 0
        hw_total = sys_total = 0.0
        for hw in data["order"]:
            code = codes[(role, hw)]
            b = bal.get(code, {}).get("balance", 0.0) * sign
            h = data["final"].get(hw, 0.0)
            hw_total += h
            sys_total += b
            ok = abs(b - h) <= 0.01
            bad += 0 if ok else 1
            det.append((hw, code, data["master"][hw] if role == "customer" else data["master"][hw][0], h, round(b, 2),
                        round(b - h, 2), "مطابق" if ok else "فرق"))
        details[label] = pd.DataFrame(det, columns=["كود HansaWorld", "كود النظام", "الاسم", "رصيد HansaWorld",
                                                    "رصيد النظام", "الفرق", "الحالة"])
        rows.append(_row(label, f"إجمالي الأرصدة ({len(data['order'])} طرف)", round(hw_total, 2), round(sys_total, 2)))
        rows.append((label, "أطراف أرصدتها غير مطابقة", 0, bad, bad, "✅ مطابق" if not bad else "❌ فرق"))
        inv_types = ("INV", "CASH") if role == "customer" else ("PINV",)
        hw_inv = sum((r["debit"] if role == "customer" else r["credit"]) for r in data["recs"] if r["type"] in inv_types)
        sys_inv = conn.execute(f"SELECT COALESCE(SUM(total_amount),0), COUNT(*) FROM {c['inv']} WHERE source_type=?",
                               (HW_SOURCE,)).fetchone()
        n_hw_inv = len({r["doc"] for r in data["recs"] if r["type"] in inv_types
                        and (r["debit"] if role == "customer" else r["credit"]) > 0})
        rows.append(_row(label, "إجمالي الفواتير", round(hw_inv, 2), round(sys_inv[0], 2)))
        rows.append((label, "عدد الفواتير (غير الصفرية)", n_hw_inv, sys_inv[1], sys_inv[1] - n_hw_inv,
                     "✅ مطابق" if sys_inv[1] == n_hw_inv else "❌ فرق"))
        hw_cred = sum(((r["credit"] - r["debit"]) if role == "customer" else (r["debit"] - r["credit"]))
                      for r in data["recs"] if r["type"] == "CRED")
        sys_cred = conn.execute("SELECT COALESCE(SUM(d.total_amount),0) FROM post_discounts d JOIN hw_import_map m "
                                "ON m.target='post_discounts' AND m.target_id=d.discount_id WHERE d.kind=?",
                                (kind,)).fetchone()[0]
        rows.append(_row(label, "إجمالي الإشعارات الدائنة", round(hw_cred, 2), round(sys_cred, 2)))
        pay = "REC" if role == "customer" else "PAYM"
        hw_pay = sum(((r["credit"] - r["debit"]) if role == "customer" else (r["debit"] - r["credit"]))
                     for r in data["recs"] if r["type"] == pay)
        in_t, out_t = ("قبض", "صرف") if role == "customer" else ("صرف", "قبض")
        sys_pay = conn.execute(
            "SELECT COALESCE(SUM(CASE WHEN t.trans_type=? THEN t.amount ELSE -t.amount END),0) FROM cash_transactions t "
            "JOIN hw_import_map m ON m.target='cash_transactions' AND m.target_id=t.trans_id "
            "WHERE m.role=? AND m.doc_type=?", (in_t, role, pay)).fetchone()[0]
        rows.append(_row(label, "صافي التحصيلات" if role == "customer" else "صافي المدفوعات", round(hw_pay, 2),
                         round(sys_pay, 2)))
        # المتبقي المتوقع لكل فاتورة من Applied To مقابل المتبقي الفعلي في النظام
        applied = defaultdict(float)
        inv_nos = {r["doc"]: True for r in data["recs"] if r["type"] in inv_types}
        for r in data["recs"]:
            if r["type"] == "CASH":
                applied[r["doc"]] += r["credit"]
            elif r["applied"] and r["type"] in (pay, "CRED"):
                applied[find_doc(inv_nos, r["applied"])[0] or r["applied"]] += (r["credit"] - r["debit"]) if role == "customer" else (r["debit"] - r["credit"])
        exp_rem = sys_rem = 0.0
        mism = []
        for no, iid, tot, paid in conn.execute(
                f"SELECT h.invoice_no, h.{c['iid']}, h.total_amount, h.paid_amount FROM {c['inv']} h WHERE h.source_type=?",
                (HW_SOURCE,)).fetchall():
            e = round(min(max(tot - applied.get(no, 0.0), 0.0), tot), 2)   # الفاتورة لا يزيد متبقيها عن إجماليها
            s = max(round(tot - paid - E.invoice_discount_total(conn, kind, iid), 2), 0.0)   # نفس تعريف النظام للمتبقي
            exp_rem += e
            sys_rem += s
            if abs(e - s) > 0.011:   # هللة واحدة = فرق تقريب في HansaWorld
                mism.append((no, tot, e, s, round(s - e, 2)))
        rows.append(_row(label, "إجمالي المتبقي على الفواتير (حسب Applied To)", round(exp_rem, 2), round(sys_rem, 2),
                         tol=0.011 * max(len(mism), 1)))
        rows.append((label, "فواتير متبقيها غير مطابق", 0, len(mism), len(mism), "✅ مطابق" if not mism else "❌ فرق"))
        if mism:
            details[f"{label} - فواتير"] = pd.DataFrame(mism, columns=["الفاتورة", "الإجمالي", "المتبقي المتوقع",
                                                                       "المتبقي بالنظام", "الفرق"])
        if role == "supplier" and data.get("net_change"):
            nc = data["net_change"]
            rows.append(_row(label, "سطر Net Change (دائن - مدين)", round(nc["credit"] - nc["debit"], 2),
                             round(sys_total, 2)))
    if hist or costs:
        inv = {r[0]: r[1:] for r in conn.execute("SELECT item_code, item_name, quantity, cost_price, opening_qty, "
                                                  "opening_cost FROM inventory")}
        cost = costs["items"] if costs else {}
        end = hist["end"] if hist else {}
        codes_i = list(dict.fromkeys(list(end) + list(cost)))
        det, bad = [], 0
        q_hw = q_sys = v_hw = v_sys = 0.0
        for code in codes_i:
            name, q, cp = (inv.get(code) or (None, 0.0, 0.0))[:3]
            eb = end.get(code)
            cs = cost.get(code, {})
            ref_q = cs.get("qty", eb if eb is not None else 0.0)
            ok = abs((q or 0) - ref_q) <= 0.001 and (eb is None or abs(eb - ref_q) <= 0.001)
            bad += 0 if ok else 1
            q_hw += ref_q
            q_sys += q or 0
            v_hw += ref_q * cs.get("cost", 0.0)
            v_sys += (q or 0) * (cp or 0)
            if ref_q or q or (eb or 0):
                det.append((code, name, eb, cs.get("qty"), round(q or 0, 3), cs.get("cost"), round(cp or 0, 5),
                            round((q or 0) * (cp or 0), 2), "مطابق" if ok else "فرق"))
        details["المخزون"] = pd.DataFrame(det, columns=["الصنف", "الاسم", "End Balance (الحركات)", "In Stock (التكلفة)",
                                                       "الرصيد بالنظام", "Cost Pr.", "التكلفة بالنظام", "القيمة بالنظام",
                                                       "الحالة"])
        rows.append(_row("المخزون", "إجمالي الكميات", round(q_hw, 3), round(q_sys, 3), tol=0.001))
        if hist and hist.get("grand") is not None:
            rows.append(_row("المخزون", "السطر الإجمالي الأخير لملف الحركات", hist["grand"], round(q_sys, 3), tol=0.001))
        rows.append(_row("المخزون", "قيمة المخزون (الرصيد × Cost Pr.)", round(v_hw, 2), round(v_sys, 2)))
        rows.append(("المخزون", "أصناف رصيدها غير مطابق", 0, bad, bad, "✅ مطابق" if not bad else "❌ فرق"))
        rows.append(("المخزون", "عدد الأصناف المعرّفة", len(set(end) | set(cost)),
                     sum(1 for k in set(end) | set(cost) if k in inv), 0, "✅ مطابق"
                     if all(k in inv for k in set(end) | set(cost)) else "❌ فرق"))
        if hist:
            n, lk, dup = conn.execute("SELECT COUNT(*), SUM(link_doc_no IS NOT NULL), SUM(is_duplicate) FROM hw_item_history"
                                      ).fetchone()
            rows.append(("المخزون", "حركات الأصناف المحفوظة", len(hist["moves"]), n or 0, (n or 0) - len(hist["moves"]),
                         "✅ مطابق" if (n or 0) == len(hist["moves"]) else "❌ فرق"))
            rows.append(("المخزون", "حركات مربوطة بمستند / «مكرر»", None, lk or 0, None, f"منها «مكرر»: {dup or 0}"))
    failed = sum(1 for r in rows if str(r[5]).startswith("❌"))
    notes = ["المقارنة بين أرقام ملفات HansaWorld المرجعية وما هو مسجل في قاعدة البيانات الآن (الأرصدة من نفس دالة كشوف "
             "الحسابات). رصيد المورد يظهر موجباً = مستحق له.",
             "السندات المستوردة مسجلة على حساب «" + HW_CASH_ACCOUNT + "» ما لم تختر حساباً آخر؛ رصيده = صافي التحصيلات "
             "والمدفوعات التاريخية ويُسوّى بقيد مع الأرصدة الفعلية للبنوك.",
             "فواتير المشتريات التاريخية مُحمّلة على تكلفة المبيعات، والمخزون الختامي أُدرج رصيداً افتتاحياً (على الوسيط "
             "الافتتاحي 99999). لتصحيح تكلفة المبيعات التاريخية: قيد من ح/ الوسيط الافتتاحي إلى ح/ تكلفة المبيعات بقيمة "
             "المخزون الختامي.",
             f"النتيجة: {len(rows) - failed} بند مطابق، {failed} بند به فرق."]
    rep = Report("تقرير مطابقة الاستيراد من HansaWorld", RECON_COLUMNS, rows, None, notes,
                 f"حتى {max([r['date'] for d in (cust, supp) if d for r in d['recs']] or [E.today()])}")
    return rep, details


def export_reconciliation_xlsx(path, rep, details):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(rep.rows, columns=[c for c, _ in rep.columns]).to_excel(xw, sheet_name="الملخص", index=False)
        for name, df in details.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
        pd.DataFrame({"ملاحظات": rep.notes}).to_excel(xw, sheet_name="ملاحظات", index=False)
    return path


# ============================================================
# نقطة التشغيل
# ============================================================
def run_hansa_import(customers=None, suppliers=None, history=None, costs=None, cash_account=None, dry_run=False,
                     balance=None, tb_as_of=None, apply_cash_openings=False):
    """استيراد ملفات HansaWorld (أي مجموعة منها). يرجع dict(stats, problems, report, details, dry_run, swapped).
    balance: ميزان المراجعة التحليلي (مرجع المطابقة العامة + تعريف حسابات المصاريف وأرصدتها المرجعية).
    apply_cash_openings: تسجيل أرصدة الصناديق والبنوك من الميزان كأرصدة افتتاحية بتاريخ tb_as_of."""
    E.require_perm("import", "استيراد بيانات HansaWorld")
    cust = read_customer_file(customers)
    supp = read_supplier_file(suppliers)
    hist = read_item_history(history)
    cost = read_cost_file(costs)
    tb = read_balance_file(balance)
    if not any((cust, supp, hist, cost, tb)):
        raise UserError("اختر ملفاً واحداً على الأقل من ملفات HansaWorld.")
    if apply_cash_openings and not tb:
        raise UserError("تسجيل أرصدة الخزينة والبنوك يحتاج ملف ميزان المراجعة.")
    problems = []
    swapped = []
    for label, d in (("العملاء", cust), ("الموردين", supp), ("حركات الأصناف", hist), ("التكلفة", cost),
                     ("ميزان المراجعة", tb)):
        if d:
            problems.extend(d["problems"])
            if d.get("swapped"):
                swapped.append(label)
    if swapped:
        problems.insert(0, "تم اكتشاف تواريخ حوّلها إكسيل بقراءة شهر/يوم في ملف " + "، ".join(swapped) +
                        " (مثل 08/09/2025 ← 2025-08-09) وتم تصحيحها إلى يوم/شهر، مع قراءة التواريخ النصية بصيغة يوم/شهر/سنة.")
    stats = defaultdict(int)
    if not dry_run:
        E.backup_database("before_hansaworld")
    out = {}
    try:
        with db() as conn:
            codes = sync_partners(conn, cust["master"] if cust else {}, supp["master"] if supp else {}, stats, problems) \
                if (cust or supp) else {}
            # الأطراف غير المعرّفة في هذا التشغيل (استيراد على مراحل): من الكود القديم المحفوظ
            for role, flagcol in (("customer", "is_customer"), ("supplier", "is_supplier")):
                for code, legacy in conn.execute(f"SELECT partner_code, legacy_code FROM business_partners WHERE {flagcol}=1 "
                                                 "AND COALESCE(legacy_code,'')!=''"):
                    codes.setdefault((role, legacy), code)
            if hist or cost:
                sync_items(conn, hist, cost, stats, problems)
            book = []
            inv_maps = {}
            if cust or supp:
                acct = _cash_account(conn, cash_account)
                for role, data in (("customer", cust), ("supplier", supp)):
                    if data:
                        inv_maps[role] = import_party_docs(conn, role, data, codes, acct, stats, problems, book)
                write_allocations(conn, book, stats, problems)
                E.stamp_voucher_reps(conn)
            if hist:
                import_item_history(conn, hist, cust, supp, codes, inv_maps, stats)
            as_of = tb_as_of or _cutover_date(conn)
            if tb:
                sync_tb_accounts(conn, tb, as_of, stats, problems)
                if apply_cash_openings:
                    apply_tb_cash_openings(conn, tb, as_of, stats, problems)
            if cust or supp or cost or tb:
                sync_hw_closing_stock_entry(conn)
            rep, details = build_reconciliation(conn, cust, supp, hist, cost, codes)
            if tb:
                tb_rows, tb_details = build_tb_reconciliation(conn, tb, cust, supp, cost)
                rep.rows.extend(tb_rows)
                details.update(tb_details)
                failed = sum(1 for r in rep.rows if str(r[5]).startswith("❌"))
                rep.notes[-1] = f"النتيجة: {len(rep.rows) - failed} بند مطابق، {failed} بند به فرق."
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)",
                         (HW_DONE_SETTING, f"{datetime.now():%Y-%m-%d %H:%M} — {rep.notes[-1]}"))
            out = dict(stats=dict(stats), problems=problems, report=rep, details=details, dry_run=dry_run, swapped=swapped)
            if dry_run:
                raise E._DryRun()
    except E._DryRun:
        pass
    return out


def stats_text(stats):
    return "  •  ".join(f"{STAT_LABELS.get(k, k)}: {v:,}" for k, v in stats.items() if v)


def last_import_summary():
    """ملخص آخر استيراد HansaWorld (التاريخ ونتيجة المطابقة)، أو نص فارغ."""
    return E.get_setting(HW_DONE_SETTING, "")


def item_history_rows(item_code=None, d_from=None, d_to=None):
    sql = ("SELECT move_date, item_code, item_name, move_type, reference, COALESCE(party_text,''), COALESCE(party_code,''), "
           "qty, balance, COALESCE(link_doc_type || ' ' || link_doc_no, ''), COALESCE(link_method,''), is_duplicate "
           "FROM hw_item_history WHERE 1=1")
    args = []
    for cond, v in (("item_code=?", item_code), ("move_date>=?", d_from), ("move_date<=?", d_to)):
        if v:
            sql += f" AND {cond}"
            args.append(v)
    with db() as conn:
        return conn.execute(sql + " ORDER BY item_code, move_date, src_row", args).fetchall()


def rep_hw_item_history(p):
    rows = [(*r[:11], "نعم" if r[11] else "") for r in item_history_rows(p.get("item"), p.get("d_from"), p.get("d_to"))]
    cols = [("التاريخ", "t"), ("الصنف", "t"), ("الاسم", "t"), ("النوع", "t"), ("المرجع", "t"), ("الطرف (HansaWorld)", "t"),
            ("كود الطرف", "t"), ("الكمية", "q"), ("الرصيد", "q"), ("المستند المربوط", "t"), ("طريقة الربط", "t"),
            ("مكرر", "t")]
    return Report("حركات الأصناف المستوردة من HansaWorld", cols, rows, None,
                  ["Deld تسليم للعميل، G.Rec. استلام من مورد، Cust. Ret. مرتجع عميل، Ret. to Supp. مرتجع لمورد، "
                   "St. Depr. إهلاك/تالف مخزني. «مكرر» = سطر عليه علامة مكرر في HansaWorld (يُستنتج طرفه من نفس المرجع)."],
                  f"{p.get('d_from') or ''} - {p.get('d_to') or ''}")


def _register_reports():
    group = next((items for title, items in E.REPORT_CATALOG if "المخزون" in title), None)
    if group is None or "hw_hist" in E.REPORT_INDEX:
        return
    item = ("hw_hist", "حركات الأصناف المستوردة من HansaWorld", rep_hw_item_history,
            "كل حركات الأصناف التاريخية (تسليم/استلام/مرتجعات) مع الطرف والمستند المربوط بها.")
    group.append(item)
    E.REPORT_INDEX[item[0]] = item[1:]
    E.REPORT_GROUP_OF[item[0]] = "reports.stock"


_register_reports()


# ============================================================
# ميزان المراجعة التحليلي (Analytical Balance Report)
# ============================================================
TB_COLS = ["Account", "Per Deb", "Per Cred", "Balance"]
TB_CUSTOMER_RANGE = (11210, 11289)     # حسابات العملاء (مطاعم، جملة، تجزئة، كاش فان)
TB_EXPENSE_GROUPS = [("مصاريف بيعية وتسويقية (43xxx)", lambda c: c.startswith("43")),
                     ("مصاريف إدارية وعمومية (5xxxx)", lambda c: c.startswith("5")),
                     ("إيجار سيارات (41125)", lambda c: c == "41125")]


def is_tb_expense(code):
    """حسابات المصاريف المرجعية (البيعية، الإدارية، إيجار السيارات) — بدون تكلفة المبيعات ومردودات المشتريات."""
    return any(cond(code) for _, cond in TB_EXPENSE_GROUPS)


def read_balance_file(src):
    """ميزان HansaWorld: Account | الاسم | B/Fwd Bal | Per Deb | Per Cred | Acc Deb | Acc Cred | Balance (مدين) | (دائن).
    عمود الرصيد مقسوم لعمودين والثاني بلا عنوان: الأول مدين والذي يليه دائن. سطور Group والإجمالي الأخير للتحقق فقط."""
    df = _read_sheet(src, "Analytical Balance Report", TB_COLS, "ميزان المراجعة")
    if df is None:
        return None
    cols = list(df.columns)
    c_acc, c_bal = cols.index("Account"), cols.index("Balance")
    c_name = cols[c_acc + 1]
    c_cr = cols[c_bal + 1] if c_bal + 1 < len(cols) else None
    accounts, groups, problems, grand = {}, {}, [], None
    for i, r in enumerate(df.to_dict("records"), start=2):
        acc = _txt(r["Account"])
        bal = round(_num(r["Balance"]) - (_num(r[c_cr]) if c_cr else 0.0), 2)
        if acc and acc.lower().startswith("group"):
            groups[re.sub(r"\D", "", acc)] = bal
            continue
        if not acc:
            if not _blank(r["Per Deb"]):
                grand = (round(_num(r["Per Deb"]), 2), round(_num(r["Per Cred"]), 2))
            continue
        if not re.fullmatch(r"\d+", acc):
            problems.append(f"ميزان المراجعة سطر {i}: كود حساب غير رقمي «{acc}» (تم تجاهله).")
            continue
        accounts[acc] = dict(name=(_txt(r[c_name]) or acc).strip(), debit=round(_num(r["Per Deb"]), 2),
                             credit=round(_num(r["Per Cred"]), 2), balance=bal, row=i)
    tot_d = round(sum(a["debit"] for a in accounts.values()), 2)
    tot_c = round(sum(a["credit"] for a in accounts.values()), 2)
    if abs(tot_d - tot_c) > 0.01:
        problems.append(f"ميزان المراجعة غير متوازن: المدين {fmt(tot_d)} والدائن {fmt(tot_c)}.")
    if grand and (abs(grand[0] - tot_d) > 0.01 or abs(grand[1] - tot_c) > 0.01):
        problems.append(f"مجموع حسابات الميزان ({fmt(tot_d)} / {fmt(tot_c)}) لا يساوي سطر الإجمالي "
                        f"({fmt(grand[0])} / {fmt(grand[1])}).")
    for g, gb in groups.items():
        sb = round(sum(a["balance"] for k, a in accounts.items() if k.startswith(g)), 2)
        if abs(sb - gb) > 0.01:
            problems.append(f"ميزان المراجعة: رصيد المجموعة {g} ({fmt(gb)}) لا يساوي مجموع حساباتها ({fmt(sb)}).")
    return dict(accounts=accounts, groups=groups, grand=grand, totals=(tot_d, tot_c), problems=problems)


def _cutover_date(conn):
    d = conn.execute("SELECT MAX(doc_date) FROM hw_import_map").fetchone()[0]
    return d or E.today()


def _tb_parent(conn, code):
    """أقرب حساب تجميعي بالدليل يبدأ به الكود (41125 ← 41100، 55014 ← 55000)."""
    best = None
    for (g,) in conn.execute("SELECT account_code FROM chart_of_accounts WHERE COALESCE(is_group,0)=1"):
        stem = g.rstrip("0")
        if stem and code.startswith(stem) and g != code and (best is None or len(stem) > len(best.rstrip("0"))):
            best = g
    return best


def sync_tb_accounts(conn, tb, as_of, stats, problems):
    """تعريف حسابات الميزان غير الموجودة بالدليل (تحت أقرب حساب تجميعي) وحفظ أرصدتها المرجعية."""
    chart = {c: (n, g) for c, n, g in conn.execute(
        "SELECT account_code, account_name, COALESCE(is_group,0) FROM chart_of_accounts")}
    created = []
    for code, a in tb["accounts"].items():
        if code not in chart:
            parent = _tb_parent(conn, code)
            if parent is None:
                problems.append(f"الحساب {code} - {a['name']} غير موجود بالدليل ولا يوجد حساب تجميعي مناسب له؛ لم يُعرَّف.")
                continue
            conn.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                         "VALUES (?,?,?,?,0)", (code, E.clean_name(a["name"]), E.classify_account(code), parent))
            created.append(f"{code} {a['name']} (تحت {parent})")
        elif chart[code][1]:
            problems.append(f"الحساب {code} في الميزان حساب تجميعي في دليل النظام؛ يُقارن بمجموع فروعه.")
        group = next((g for g in sorted(tb["groups"], key=len, reverse=True) if code.startswith(g)), None)
        conn.execute("INSERT OR REPLACE INTO hw_tb_reference (account_code, account_name, group_code, per_debit, per_credit, "
                     "balance, as_of) VALUES (?,?,?,?,?,?,?)",
                     (code, a["name"], group, a["debit"], a["credit"], a["balance"], as_of))
    stats["tb_accounts"] = len(tb["accounts"])
    stats["tb_accounts_created"] = len(created)
    if created:
        problems.append("حسابات جديدة عُرّفت بالدليل من الميزان: " + "، ".join(created))


def _tb_cash_names(conn, tb):
    """{كود حساب نقدية بالميزان: اسمه في النظام}."""
    names = {c: n for n, c in E.cash_accounts(conn).items()}
    return {code: names.get(code) or a["name"] for code, a in tb["accounts"].items() if code.startswith(E.CASH_PREFIX)}


def apply_tb_cash_openings(conn, tb, as_of, stats, problems):
    names = _tb_cash_names(conn, tb)
    amounts = {names[c]: tb["accounts"][c]["balance"] for c in names}
    E.set_cash_opening(as_of, amounts, conn=conn)
    stats["cash_openings"] = sum(1 for v in amounts.values() if abs(v) > 0.005)
    problems.append(f"سُجّلت أرصدة الخزينة والبنوك من الميزان كأرصدة افتتاحية بتاريخ {as_of} (قيد واحد على الوسيط الافتتاحي "
                    "99999 يستبدل أي أرصدة افتتاحية سابقة للخزينة).")


def gl_activity(conn, codes=None):
    """صافي (مدين - دائن) كل حساب في الدفتر العام الآن (قيود يدوية + ترحيل آلي)، داخل نفس المعاملة."""
    from . import ledger
    ledger.process_gl_queue(conn)
    out = defaultdict(float)
    for acc, v in conn.execute(f"SELECT acc, SUM(dr - cr) FROM ({ledger._GL_UNION}) GROUP BY acc"):
        if codes is None or acc in codes:
            out[acc] = round(v or 0.0, 2)
    return out


def build_tb_reconciliation(conn, tb, cust, supp, cost):
    """المطابقة العامة مع ميزان HansaWorld: الموردون، العملاء، المخزون، الصناديق والبنوك، ومجموعات المصاريف."""
    rows, details = [], {}
    acc = tb["accounts"]
    sec = "ميزان المراجعة"
    tot_d, tot_c = tb["totals"]
    rows.append((sec, "توازن الميزان (المدين = الدائن)", tot_d, tot_c, round(tot_c - tot_d, 2),
                 "✅ مطابق" if abs(tot_d - tot_c) <= 0.01 else "❌ فرق"))
    bal = E.partner_balances("9999-12-31", conn=conn)
    flags = {c: (bool(ic), bool(is_)) for c, ic, is_ in conn.execute(
        "SELECT partner_code, is_customer, is_supplier FROM business_partners")}
    ar = round(sum(b["balance"] for b in bal if flags.get(b["code"], (True, False)) == (True, False)), 2)
    ap = round(-sum(b["balance"] for b in bal if flags.get(b["code"], (False, True)) == (False, True)), 2)
    if "22101" in acc:
        rows.append(_row(sec, "الموردون 22101 (دائن) ↔ أرصدة الموردين بالنظام", round(-acc["22101"]["balance"], 2), ap))
        if supp:
            rows.append(_row(sec, "الموردون 22101 ↔ ملف حركات الموردين", round(-acc["22101"]["balance"], 2),
                             round(sum(supp["final"].values()), 2)))
    cust_codes = [c for c in acc if TB_CUSTOMER_RANGE[0] <= int(c) <= TB_CUSTOMER_RANGE[1]]
    if cust_codes:
        tb_ar = round(sum(acc[c]["balance"] for c in cust_codes), 2)
        label = f"العملاء {min(cust_codes)}–{max(cust_codes)} ({len(cust_codes)} حسابات)"
        rows.append(_row(sec, f"{label} ↔ أرصدة العملاء بالنظام", tb_ar, ar))
        if cust:
            rows.append(_row(sec, f"{label} ↔ ملف تفاصيل العملاء", tb_ar, round(sum(cust["final"].values()), 2)))
        details["الميزان - العملاء"] = pd.DataFrame([(c, acc[c]["name"], acc[c]["balance"]) for c in cust_codes],
                                                   columns=["الحساب", "الاسم", "الرصيد بالميزان"])
    if "11310" in acc:
        stock = round(conn.execute("SELECT COALESCE(SUM(quantity * cost_price),0) FROM inventory").fetchone()[0], 2)
        rows.append(_row(sec, "المخزون 11310 ↔ قيمة المخزون بالنظام (الرصيد × التكلفة)", acc["11310"]["balance"], stock))
        if cost:
            rows.append(_row(sec, "المخزون 11310 ↔ ملف التكلفة (In Stock × Cost Pr.)", acc["11310"]["balance"],
                             round(sum(c["qty"] * c["cost"] for c in cost["items"].values()), 2)))
    cash = E.cash_balances(conn)
    for code, name in sorted(_tb_cash_names(conn, tb).items()):
        rows.append(_row(sec, f"الخزينة/البنك {code} {name}", acc[code]["balance"], round(cash.get(name, 0.0), 2)))
    clr = round(cash.get(HW_CASH_ACCOUNT, 0.0), 2)
    if abs(clr) > 0.005:
        rows.append((sec, f"«{HW_CASH_ACCOUNT}» (صافي تحصيلات ومدفوعات الشركاء التاريخية — ليس في الميزان)", None, clr,
                     None, "ℹ️ للتسوية"))
    act = gl_activity(conn)
    det = []
    for label, cond in TB_EXPENSE_GROUPS:
        codes = sorted(c for c in acc if cond(c))
        if not codes:
            continue
        ref = round(sum(acc[c]["balance"] for c in codes), 2)
        sysv = round(sum(act.get(c, 0.0) for c in codes), 2)
        st = "✅ مطابق" if abs(ref - sysv) <= 0.01 else ("📌 مرجع فقط (لم تُستورد الحركات)" if abs(sysv) <= 0.005 else "❌ فرق")
        rows.append((sec, f"{label} — {len(codes)} حساب", ref, sysv, round(sysv - ref, 2), st))
        det += [(label, c, acc[c]["name"], acc[c]["balance"], act.get(c, 0.0), round(act.get(c, 0.0) - acc[c]["balance"], 2))
                for c in codes]
    if det:
        details["الميزان - المصاريف"] = pd.DataFrame(det, columns=["المجموعة", "الحساب", "الاسم", "الرصيد المرجعي",
                                                                  "المسجل بالنظام", "الفرق"])
    return rows, details


# ============================================================
# حركات المصاريف التفصيلية (عند تصديرها لاحقاً من HansaWorld أو من أي نظام)
# ============================================================
EXP_COLS = dict(date=["التاريخ", "Date", "Trans Date", "Trans. Date"],
                account=["كود الحساب", "الحساب", "Account", "Acc", "Account Code"],
                amount=["المبلغ", "Amount", "Debit", "Debit (SAR)", "Per Deb", "مدين"],
                credit=["دائن", "Credit", "Credit (SAR)", "Per Cred"],
                cash=["الخزينة / البنك", "الخزينة", "البنك", "Cash Account", "Bank", "Paid From"],
                desc=["البيان", "Description", "Text", "Comment", "Memo"],
                doc=["رقم المستند", "Doc No", "Number", "No", "Ref", "Reference"])
EXP_TEMPLATE = ["التاريخ", "كود الحساب", "المبلغ", "الخزينة / البنك", "البيان", "رقم المستند"]


def write_expense_template(path):
    pd.DataFrame([["2026-11-19", "52014", 1000, HW_CASH_ACCOUNT, "إيجار مكتب نوفمبر", "EXP-0001"]],
                 columns=EXP_TEMPLATE).to_excel(path, index=False)
    return path


def import_expense_movements(src, cash_account=None, dry_run=False):
    """استيراد حركات مصاريف تفصيلية: سند صرف لكل سطر على حساب المصروف (يرحَّل آلياً للدفتر العام)، بدون تكرار،
    ثم مقارنة كل حساب بالرصيد المرجعي من الميزان. يرجع dict(stats, problems, report, details, dry_run)."""
    E.require_perm("import", "استيراد حركات المصاريف")
    if src is None:
        raise UserError("اختر ملف حركات المصاريف.")
    try:
        is_csv = str(getattr(src, "name", src)).lower().endswith(".csv")
        df = pd.read_csv(src, dtype=object) if is_csv else pd.read_excel(src, dtype=object)
    except Exception as e:  # noqa: BLE001
        raise UserError(f"تعذر قراءة ملف المصاريف: {e}")
    df.columns = [str(c).strip() for c in df.columns]
    c = {k: E._pick_col(df, v) for k, v in EXP_COLS.items()}
    missing = [lbl for k, lbl in (("date", "التاريخ"), ("account", "كود الحساب"), ("amount", "المبلغ")) if c[k] is None]
    if missing:
        raise UserError("ملف المصاريف ينقصه عمود: " + "، ".join(missing) + "\nالأعمدة المقبولة: " + "، ".join(EXP_TEMPLATE))
    swap = detect_swapped_dates(df[c["date"]])
    stats, problems = defaultdict(int), []
    total = 0.0
    if not dry_run:
        E.backup_database("before_expenses")
    out = {}
    try:
        with db() as conn:
            default_cash = _cash_account(conn, cash_account)
            cash_names = E.cash_account_names(conn)
            code_to_cash = {code: n for n, code in E.cash_accounts(conn).items()}
            chart = {k: (n, t, g) for k, n, t, g in conn.execute(
                "SELECT account_code, account_name, account_type, COALESCE(is_group,0) FROM chart_of_accounts")}
            done = {d for (d,) in conn.execute("SELECT doc_no FROM hw_import_map WHERE role='expense'")}
            seen = defaultdict(int)
            for i, r in enumerate(df.to_dict("records"), start=2):
                code = _txt(r[c["account"]])
                amt = round(_num(r[c["amount"]]) - (_num(r[c["credit"]]) if c["credit"] else 0.0), 2)
                if not code and not amt:
                    continue
                d = parse_hw_date(r[c["date"]], swap)
                bad = None
                if not d:
                    bad = "تاريخ غير صالح"
                elif code not in chart or chart[code][2]:
                    bad = f"الحساب {code} غير موجود أو تجميعي بالدليل"
                elif chart[code][1] != E.TYPE_EXP:
                    bad = f"الحساب {code} ليس حساب مصروفات ({chart[code][1]})"
                elif amt <= 0:
                    bad = f"مبلغ دائن/سالب ({fmt(amt)}) على {code} — سجّل التسوية بقيد يومية يدوي"
                bank = _txt(r[c["cash"]]) if c["cash"] else None
                bank = code_to_cash.get(bank, bank) if bank else default_cash
                if not bad and bank not in cash_names:
                    bad = f"حساب الخزينة «{bank}» غير معرّف"
                if bad:
                    problems.append(f"سطر {i}: {bad}.")
                    stats["bad"] += 1
                    continue
                doc = _txt(r[c["doc"]]) if c["doc"] else None
                base = doc or f"{d}|{code}|{amt:.2f}"
                seen[base] += 1
                key = base if (doc and seen[base] == 1) else f"{base}#{seen[base]}"
                if key in done:
                    stats["skipped_existing"] += 1
                    continue
                desc = (_txt(r[c["desc"]]) if c["desc"] else None) or f"مصروف {chart[code][0]}"
                tid = conn.execute("INSERT INTO cash_transactions (trans_type, trans_date, account_bank, category, party_code, "
                                   "party_name, amount, description) VALUES ('صرف',?,?,?,?,?,?,?)",
                                   (d, bank, E.CAT_EXPENSE, code, chart[code][0], amt,
                                    desc + (f" ({doc})" if doc else ""))).lastrowid
                _map(conn, "expense", "EXP", key, code, "cash_transactions", tid, d, amt)
                done.add(key)
                stats["expenses"] += 1
                total += amt
            if stats.get("expenses") and _resync_tb_aggregate(conn):
                problems.append(f"أُعيد بناء قيد «{TB_EXP_LABEL}» ليطرح الحركات التفصيلية الجديدة (لا تكرار للمصروف).")
            ref = {k: (n, b) for k, n, b in conn.execute(
                "SELECT r.account_code, r.account_name, r.balance FROM hw_tb_reference r JOIN chart_of_accounts a "
                "ON a.account_code=r.account_code WHERE a.account_type=?", (E.TYPE_EXP,)) if is_tb_expense(k)}
            for k in {code for (code,) in conn.execute("SELECT party_code FROM hw_import_map WHERE role='expense'")}:
                ref.setdefault(k, (chart[k][0], 0.0))   # حساب عليه حركات بلا رصيد مرجعي
            act = gl_activity(conn, set(ref))
            rows = [("المصاريف", f"{k} {n}", b, act.get(k, 0.0), round(act.get(k, 0.0) - b, 2),
                     "✅ مطابق" if abs(act.get(k, 0.0) - b) <= 0.01 else "❌ فرق") for k, (n, b) in sorted(ref.items())]
            failed = sum(1 for r_ in rows if r_[5].startswith("❌"))
            notes = ["مقارنة المسجل بالنظام على كل حساب مصروف بالرصيد المرجعي من ميزان HansaWorld." if ref else
                     "لا توجد أرصدة مرجعية: استورد ملف ميزان المراجعة أولاً للمقارنة.",
                     f"النتيجة: {len(rows) - failed} بند مطابق، {failed} بند به فرق."]
            st_ = dict(stats)
            if total:
                st_["expense_total"] = round(total, 2)
            out = dict(stats=st_, problems=problems, dry_run=dry_run, swapped=["المصاريف"] if swap else [],
                       report=Report("مطابقة المصاريف مع ميزان HansaWorld", RECON_COLUMNS, rows, None, notes, ""),
                       details={})
            if dry_run:
                raise E._DryRun()
    except E._DryRun:
        pass
    return out


def rep_hw_expenses(p):
    """تقرير: المصاريف المسجلة بالنظام مقابل الأرصدة المرجعية من ميزان HansaWorld."""
    with db() as conn:
        ref = conn.execute("SELECT r.account_code, r.account_name, r.balance, COALESCE(r.as_of,'') FROM hw_tb_reference r "
                           "JOIN chart_of_accounts c ON c.account_code=r.account_code WHERE c.account_type=? "
                           "ORDER BY r.account_code", (E.TYPE_EXP,)).fetchall()
        ref = [r for r in ref if is_tb_expense(r[0])]
        act = gl_activity(conn, {r[0] for r in ref})
    rows = []
    for code, name, b, _ in ref:
        group = next(lbl for lbl, cond in TB_EXPENSE_GROUPS if cond(code))
        rows.append((group, code, name, b, act.get(code, 0.0), round(act.get(code, 0.0) - b, 2)))
    cols = [("المجموعة", "t"), ("الحساب", "t"), ("الاسم", "t"), ("الرصيد المرجعي (HansaWorld)", "m"),
            ("المسجل بالنظام", "m"), ("الفرق", "m")]
    total = ("الإجمالي", "", "", *[round(sum(r[i] for r in rows), 2) for i in (3, 4, 5)])
    return Report("المصاريف مقابل ميزان HansaWorld المرجعي", cols, rows, total,
                  ["الرصيد المرجعي من ملف ميزان المراجعة التحليلي (لا يُرحَّل كقيد). يتطابق المسجل عند استيراد حركات "
                   "المصاريف التفصيلية من مركز الاستيراد."], f"حتى {ref[0][3] if ref else ''}")


def _register_tb_report():
    group = next((items for title, items in E.REPORT_CATALOG if "الخزينة" in title), None)
    if group is None or "hw_exp" in E.REPORT_INDEX:
        return
    item = ("hw_exp", "المصاريف مقابل ميزان HansaWorld المرجعي", rep_hw_expenses,
            "كل حساب مصروف في ميزان HansaWorld: الرصيد المرجعي والمسجل بالنظام والفرق.")
    group.append(item)
    E.REPORT_INDEX[item[0]] = item[1:]
    E.REPORT_GROUP_OF[item[0]] = "reports.cash"


_register_tb_report()


# ============================================================
# ترحيل أرصدة المصاريف المرجعية كتجميع فترة (قيد نظام واحد قابل لإعادة البناء)
# ============================================================
TB_EXP_SOURCE = "hw_tb_expenses"
TB_EXP_LABEL = "تجميع فترة من ميزان HansaWorld"


TB_OTHER_REV_LABEL = "إيرادات أخرى (33xxx) — فروق العملة والخدمات"


def is_tb_aggregate(code):
    """الحسابات التي يُرحّلها قيد التجميع: المصاريف المرجعية + الإيرادات الأخرى 33xxx (المبيعات 31xxx من الفواتير، وتكلفة
    المبيعات 41110/41201 تُسوّى تلقائياً بقيد تسوية التكلفة)."""
    return is_tb_expense(code) or str(code).startswith(E.HW_OTHER_REV_PREFIX)


def _tb_group_label(code):
    if str(code).startswith(E.HW_OTHER_REV_PREFIX):
        return TB_OTHER_REV_LABEL
    return next(lbl for lbl, cond in TB_EXPENSE_GROUPS if cond(code))


def _tb_aggregate_entries(conn):
    """[(رقم القيد، التاريخ، الحساب المقابل)] لقيود التجميع (قيد لكل فترة مرجعية) بترتيب التاريخ."""
    out = []
    for eid, d in conn.execute("SELECT entry_id, entry_date FROM journal_entries WHERE source_type=? "
                               "ORDER BY entry_date, entry_id", (TB_EXP_SOURCE,)).fetchall():
        counter = conn.execute("SELECT account_code FROM journal_lines WHERE entry_id=? AND memo LIKE 'مقابل%' LIMIT 1",
                               (eid,)).fetchone()
        out.append((eid, d, counter[0] if counter else None))
    return out


def _tb_aggregate_entry(conn):
    e = _tb_aggregate_entries(conn)
    return e[0] if e else (None, None, None)


def _gl_activity_between(conn, codes, d_from, d_to):
    """صافي (مدين − دائن) لكل حساب في الدفتر العام بين تاريخين (حركات تفصيلية فقط بعد حذف سطور قيد التجميع)."""
    from . import ledger
    codes = list(codes)
    if not codes:
        return {}
    ledger.process_gl_queue(conn)
    ph = ",".join("?" * len(codes))
    return {acc: round(v or 0.0, 2) for acc, v in conn.execute(
        f"SELECT acc, SUM(dr - cr) FROM ({ledger._GL_UNION}) WHERE d BETWEEN ? AND ? AND acc IN ({ph}) GROUP BY acc",
        [d_from, d_to, *codes])}


def _write_tb_aggregate(conn, counter, as_of=None):
    """يبني قيود التجميع من الصفر: قيد لكل فترة مرجعية (الميزان، أو الميزان مقسوماً بقوائم الدخل المستوردة) بتاريخ آخر
    يوم فيها. لكل حساب = رصيد الفترة المرجعي − ما سُجّل عليه من حركات تفصيلية داخل نفس الفترة (فلا يتكرر المصروف،
    ولا تؤثر مصروفات التشغيل الجديدة بعد الترحيل). as_of يغيّر تاريخ القيد فقط لو كانت فترة واحدة.
    يرجع (أرقام القيود، الصفوف)."""
    old = _tb_aggregate_entries(conn)
    for eid, _, _ in old:
        conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (eid,))
    acc = conn.execute("SELECT account_name, COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?",
                       (counter,)).fetchone()
    if acc is None or acc[1]:
        raise UserError(f"الحساب المقابل {counter} غير موجود أو تجميعي بالدليل.")
    segs = E.hw_ref_segments(conn)
    if not any(is_tb_aggregate(c) for sg in segs for c in sg["amounts"]):
        raise UserError("لا توجد أرصدة مصاريف مرجعية: استورد ملف ميزان المراجعة أولاً.")
    if as_of and len(segs) == 1:
        segs = [dict(segs[0], d_to=as_of)]
    rows, entries = [], []
    for i, sg in enumerate(segs):
        ref = sorted((c, n, a) for c, (n, a) in sg["amounts"].items() if is_tb_aggregate(c))
        if not ref:
            continue
        period = f"{sg['d_from']} → {sg['d_to']}"
        act = _gl_activity_between(conn, [c for c, _, _ in ref], "0000-01-01" if i == 0 else sg["d_from"], sg["d_to"])
        lines = []
        for code, name, bal in ref:
            detail = round(act.get(code, 0.0), 2)
            post = round(bal - detail, 2)
            rows.append((f"{_tb_group_label(code)} | {period}" if len(segs) > 1 else _tb_group_label(code),
                         f"{code} {name}", bal, detail, post, "✅" if abs(post) <= 0.005 else "📥 يُرحَّل"))
            if abs(post) > 0.005:
                lines.append((code, name, max(post, 0.0), max(-post, 0.0),
                              f"{TB_EXP_LABEL} ({period}): رصيد مرجعي {fmt(bal)} − حركات تفصيلية {fmt(detail)}"))
        if lines:
            entries.append((sg, lines))
    eids = []
    for j, (sg, lines) in enumerate(entries):
        total = round(sum(d - c for _, _, d, c, _ in lines), 2)
        desc = (f"قيد نظام — {TB_EXP_LABEL} للفترة {sg['d_from']} → {sg['d_to']}: أرصدة المصاريف البيعية والإدارية "
                f"وإيجار السيارات والإيرادات الأخرى (ليست حركات تفصيلية)")
        if j < len(old):
            eid = old[j][0]
            conn.execute("UPDATE journal_entries SET entry_date=?, description=?, source_id=? WHERE entry_id=?",
                         (sg["d_to"], desc, j, eid))
        else:
            eid = conn.execute("INSERT INTO journal_entries (entry_date, description, source_type, source_id) "
                               "VALUES (?,?,?,?)", (sg["d_to"], desc, TB_EXP_SOURCE, j)).lastrowid
        for code, name, d, c, memo in lines:
            conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                         "VALUES (?,?,?,?,?,?)", (eid, code, name, d, c, memo))
        conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                     "VALUES (?,?,?,?,?,?)", (eid, counter, acc[0], max(-total, 0.0), max(total, 0.0),
                                              f"مقابل {TB_EXP_LABEL}"))
        eids.append(eid)
    for eid, _, _ in old[len(entries):]:
        conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (eid,))
    return eids, rows


def _is_rev_row(r):
    return str(r[1]).startswith(E.HW_OTHER_REV_PREFIX)


def post_tb_expense_aggregate(counter=None, as_of=None, dry_run=False, remove=False):
    """ترحيل (أو إعادة بناء، أو إلغاء) قيود تجميع المصاريف والإيرادات الأخرى من مراجع HansaWorld: قيد لكل فترة مرجعية.
    قابل لإعادة التشغيل بلا تكرار. يرجع dict(entry_id, entry_ids, total (المصاريف), other_income, report, dry_run, removed)."""
    E.require_perm("import", "ترحيل أرصدة المصاريف المرجعية")
    if as_of and not E.parse_date(as_of):
        raise UserError("تاريخ القيد غير صحيح (YYYY-MM-DD).")
    if not dry_run:
        E.backup_database("before_tb_expenses")
    out = {}
    try:
        with db() as conn:
            old = _tb_aggregate_entries(conn)
            if remove:
                for eid, _, _ in old:
                    conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (eid,))
                    conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (eid,))
                out = dict(entry_id=None, entry_ids=[], total=0.0, other_income=0.0, removed=bool(old),
                           dry_run=dry_run, report=None)
            else:
                counter = counter or (old[0][2] if old else None) or E.OPENING_ACCOUNT
                eids, rows = _write_tb_aggregate(conn, counter, as_of)
                dates = [d for (d,) in conn.execute(
                    f"SELECT entry_date FROM journal_entries WHERE entry_id IN ({','.join('?' * len(eids)) or 'NULL'}) "
                    "ORDER BY entry_date", eids)]
                total = round(sum(r[4] for r in rows if not _is_rev_row(r)), 2)
                other = round(-sum(r[4] for r in rows if _is_rev_row(r)), 2)
                tot = ("الإجمالي (المصاريف)", "", round(sum(r[2] for r in rows if not _is_rev_row(r)), 2),
                       round(sum(r[3] for r in rows if not _is_rev_row(r)), 2), total, "")
                when = "، ".join(dates) or "-"
                notes = [f"قيود نظام بتاريخ {when} (لا تُعدَّل ولا تُحذف من شاشة القيود): قيد لكل فترة مرجعية من HansaWorld، "
                         f"من ح/ كل مصروف إلى ح/ {counter}، ومن ح/ {counter} إلى ح/ الإيرادات الأخرى (33xxx).",
                         "الفترات: الميزان كاملاً، أو مقسوماً بقوائم الدخل المستوردة (قائمة 2026 = أرقامها، والباقي = الميزان − "
                         "القوائم) فتظهر مصروفات كل سنة في سنتها.",
                         "يُرحَّل الفرق فقط بين الرصيد المرجعي وما سُجّل من حركات تفصيلية داخل نفس الفترة؛ استيراد حركات تفصيلية "
                         "لاحقاً يعيد بناء القيود تلقائياً فلا يتكرر المصروف."]
                if abs(other) > 0.005:
                    notes.append(f"منها إيرادات أخرى (فروق عملة وخدمات 33xxx): {fmt(other)} ر.س.")
                cols = [("المجموعة", "t"), ("الحساب", "t"), ("الرصيد المرجعي", "m"), ("حركات تفصيلية بالنظام", "m"),
                        ("المرحّل في قيد التجميع", "m"), ("الحالة", "t")]
                out = dict(entry_id=eids[0] if eids else None, entry_ids=eids, total=total, other_income=other,
                           removed=False, dry_run=dry_run, counter=counter, as_of=when,
                           report=Report(f"ترحيل أرصدة المصاريف والإيرادات الأخرى — {TB_EXP_LABEL}", cols, rows, tot,
                                         notes, f"حتى {dates[-1] if dates else '-'}"))
            if dry_run:
                raise E._DryRun()
    except E._DryRun:
        pass
    return out


def _resync_tb_aggregate(conn):
    """بعد استيراد حركات مصاريف تفصيلية: لو قيد التجميع موجود يُعاد بناؤه حتى لا يتكرر المصروف."""
    old = _tb_aggregate_entries(conn)
    if old:
        _write_tb_aggregate(conn, old[0][2] or E.OPENING_ACCOUNT, old[0][1] if len(old) == 1 else None)
        return True
    return False



# ============================================================
# تسوية تكلفة المبيعات بالمخزون الختامي في الدفتر العام (جرد دوري لفترة HansaWorld)
# ============================================================
HW_CLOSING_SOURCE = "hw_cogs_closing"


def sync_hw_closing_stock_entry(conn):
    """فواتير مشتريات HansaWorld تُرحَّل لتكلفة المبيعات، والمخزون الختامي يُرحَّل رصيداً افتتاحياً على الوسيط 99999.
    قيود النظام هنا (تُعاد بنائها مع كل استيراد) تجعل ربح الدفتر العام = قائمة الدخل ولوحة المؤشرات:
      1) طرح المخزون الختامي من تكلفة المبيعات في تاريخ الترحيل: من ح/ الوسيط الافتتاحي إلى ح/ تكلفة المبيعات.
      2) لكل فترة مرجعية من HansaWorld (الميزان/قوائم الدخل): تسوية التكلفة لتساوي دفتر HansaWorld (41110 + 41201)
         في آخر يوم من الفترة (engine.hw_cogs_ref_adjustments) — نفس الأرقام التي تحسبها قائمة الدخل.
    يرجع رقم قيد المخزون الختامي (أو None)."""
    from . import ledger
    cutover = conn.execute("SELECT MAX(doc_date) FROM hw_import_map WHERE role IN ('customer','supplier')").fetchone()[0]
    for (old,) in conn.execute("SELECT entry_id FROM journal_entries WHERE source_type=?", (HW_CLOSING_SOURCE,)).fetchall():
        conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (old,))
        conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (old,))
    if not cutover:
        return None
    A = ledger.gl_accounts(conn)
    cogs, opening = A["gl_acc_cogs"], A["gl_acc_opening"]
    names = {c: n for c, n in conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE account_code IN (?,?)",
                                           (cogs, opening))}
    if len(names) < 2:
        return None

    def post(day, desc, amount, memo_open, memo_cogs):
        """amount موجب = تخفيض تكلفة المبيعات (من ح/ الوسيط إلى ح/ التكلفة)، وسالب = زيادتها."""
        eid = conn.execute("INSERT INTO journal_entries (entry_date, description, source_type, source_id) VALUES (?,?,?,0)",
                           (day, desc, HW_CLOSING_SOURCE)).lastrowid
        conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) VALUES (?,?,?,?,?,?)",
                     (eid, opening, names[opening], max(amount, 0.0), max(-amount, 0.0), memo_open))
        conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) VALUES (?,?,?,?,?,?)",
                     (eid, cogs, names[cogs], max(-amount, 0.0), max(amount, 0.0), memo_cogs))
        return eid

    (closing,) = conn.execute("SELECT COALESCE(SUM(opening_qty * opening_cost), 0) FROM inventory").fetchone()
    closing = round(closing or 0.0, 2)
    eid = None
    if closing > 0.005:
        eid = post(cutover, f"قيد نظام — تسوية تكلفة مبيعات فترة HansaWorld بالمخزون الختامي في {cutover} (جرد دوري)", closing,
                   "المخزون الختامي لفترة HansaWorld (مقابل رصيده الافتتاحي)", "تخفيض تكلفة المبيعات بالمخزون الختامي (جرد دوري)")
    for a in E.hw_cogs_ref_adjustments(conn):
        if abs(a["adj"]) <= 0.005:
            continue
        sg = a["seg"]
        post(a["date"], f"قيد نظام — تسوية تكلفة مبيعات HansaWorld للفترة {sg['d_from']} → {sg['d_to']} لتساوي دفترها "
                        f"(41110 + 41201): {fmt(a['ref'])} − جرد دوري {fmt(a['periodic'])}", -a["adj"],
             "مقابل تسوية تكلفة المبيعات لدفتر HansaWorld",
             f"تسوية تكلفة المبيعات لدفتر HansaWorld ({sg['d_from']} → {sg['d_to']})")
    return eid


def tb_expense_status():
    """حالة أرصدة المصاريف المرجعية لعرضها بوضوح في تبويب الاستيراد:
    dict(groups=[(المجموعة، عدد الحسابات، الرصيد المرجعي)], total, other_income, as_of,
         posted=(أرقام القيود، التواريخ، مبلغ المصاريف) أو None، segments=[(من، إلى، المصدر، مجاميع الفترة)])."""
    with db() as conn:
        try:
            ref = conn.execute("SELECT account_code, balance, as_of FROM hw_tb_reference").fetchall()
        except sqlite3.OperationalError:
            ref = []
        other = round(-sum(r[1] or 0.0 for r in ref if str(r[0]).startswith(E.HW_OTHER_REV_PREFIX)), 2)
        ref = [r for r in ref if is_tb_expense(r[0])]
        groups = []
        for label, cond in TB_EXPENSE_GROUPS:
            rows = [r for r in ref if cond(r[0])]
            groups.append((label, len(rows), round(sum(r[1] or 0.0 for r in rows), 2)))
        entries = _tb_aggregate_entries(conn)
        posted = None
        if entries:
            ids = [e[0] for e in entries]
            (amt,) = conn.execute(f"SELECT COALESCE(SUM(debit - credit), 0) FROM journal_lines l JOIN chart_of_accounts a "
                                  f"ON a.account_code=l.account_code WHERE l.entry_id IN ({','.join('?' * len(ids))}) "
                                  f"AND a.account_type=?", (*ids, E.TYPE_EXP)).fetchone()
            posted = ("، ".join(str(i) for i in ids), "، ".join(e[1] for e in entries), round(amt or 0.0, 2))
        segments = [(sg["d_from"], sg["d_to"], sg["src"], E.hw_segment_totals(sg["amounts"]))
                    for sg in E.hw_ref_segments(conn)]
    return dict(groups=groups, total=round(sum(g[2] for g in groups), 2), other_income=other,
                as_of=max((r[2] or "" for r in ref), default="") or None, posted=posted, segments=segments)


# ============================================================
# قائمة دخل HansaWorld لفترة (Profit & Loss) — مرجع يقسم فترة الميزان، ومطابقة قائمة الدخل بالنظام معها
# ============================================================
PL_PERIOD_RE = re.compile(r"(?<!Budget )Period\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})\s*:\s*(\d{1,2}/\d{1,2}/\d{4})", re.I)
PL_CODE_RE = re.compile(r"^[345]\d{4,5}$")
PL_AMOUNT_RE = re.compile(r"^-?\(?-?[\d,]*\d\.\d+\)?$")
PL_SRC_LABELS = {"pl": "قائمة دخل HansaWorld", "tb_rest": "الميزان − قوائم الدخل", "tb": "ميزان المراجعة"}


def _pl_amount(t):
    neg = t.startswith("(") and t.endswith(")")
    v = float(t.strip("()").replace(",", ""))
    return -v if neg else v


def _pl_date(t):
    d, m, y = (int(x) for x in t.split("/"))
    return date(y, m, d).isoformat()


def _pl_rows(src):
    """سطور القائمة كقوائم خلايا: من Excel (كل الشيتات) أو CSV أو نص منسوخ من ملف PDF (سطر لكل حساب)."""
    name = str(getattr(src, "name", src)).lower()
    if name.endswith(".txt"):
        raw = src.read() if hasattr(src, "read") else open(src, "rb").read()
        text = raw.decode("utf-8-sig") if isinstance(raw, bytes) else raw
        return [ln.split() for ln in text.splitlines() if ln.strip()]
    try:
        if name.endswith(".csv"):
            sheets = {"csv": pd.read_csv(src, header=None, dtype=object)}
        else:
            sheets = pd.read_excel(src, sheet_name=None, header=None, dtype=object)
    except Exception as e:  # noqa: BLE001
        raise UserError(f"تعذر قراءة ملف قائمة الدخل: {e}")
    rows = []
    for df in sheets.values():
        for rec in df.itertuples(index=False):
            cells = []
            for v in rec:
                if v is None or (isinstance(v, float) and math.isnan(v)):
                    continue
                if isinstance(v, (datetime, date, pd.Timestamp)):
                    cells.append(pd.Timestamp(v).strftime("%d/%m/%Y"))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    cells.append(float(v))
                elif str(v).strip():
                    cells.append(str(v).strip())
            if cells:
                rows.append(cells)
    return rows


def parse_pl_rows(rows):
    """تحليل تقرير Profit & Loss من HansaWorld: الفترة («Period dd/mm/yyyy : dd/mm/yyyy»)، ولكل حساب أول مبلغ
    (عمود Period)، وسطر صافي الربح للتحقق (مجموع الحسابات = صافي الربح بنفس الإشارة).
    يرجع dict(period=(من، إلى) أو None، accounts={الكود: (الاسم، المبلغ)}، net، problems)."""
    period, accounts, net, problems = None, {}, None, []
    for cells in rows:
        text = " ".join(f"{c:.2f}" if isinstance(c, float) else str(c) for c in cells)
        m = PL_PERIOD_RE.search(text)
        if m and period is None:
            try:
                period = (_pl_date(m.group(1)), _pl_date(m.group(2)))
            except ValueError:
                problems.append(f"تاريخ الفترة غير صالح: {m.group(0)}")
        code, amts, words = None, [], []
        for c in cells:
            if isinstance(c, float):
                if code is None and c.is_integer() and PL_CODE_RE.match(str(int(c))):
                    code = str(int(c))
                else:
                    amts.append(c)
                continue
            t = str(c).strip()
            if code is None and PL_CODE_RE.match(t):
                code = t
            elif PL_AMOUNT_RE.match(t):
                amts.append(_pl_amount(t))
            else:
                words.append(t)
        low = text.lower()
        if "net profit" in low and amts:
            net = round(amts[0], 2)
        if code and amts and "total" not in low:
            accounts[code] = (" ".join(words).strip() or code, round(amts[0], 2))
    if accounts and net is not None:
        tot = round(sum(a for _, a in accounts.values()), 2)
        if abs(tot - net) > 0.01:
            problems.append(f"مجموع الحسابات ({fmt(tot)}) لا يساوي سطر صافي الربح بالقائمة ({fmt(net)}).")
    return dict(period=period, accounts=accounts, net=net, problems=problems)


def import_pl_reference(src, period_from=None, period_to=None, dry_run=False):
    """استيراد قائمة دخل HansaWorld لفترة كمرجع (لا تُرحَّل كقيود بنفسها): تقسم فترة الميزان إلى فترات، فتُعاد بناء قيود
    تجميع المصاريف وتسوية تكلفة المبيعات لكل فترة. إعادة استيراد نفس الفترة تستبدلها. الفترات المتداخلة مرفوضة.
    يرجع dict(period, accounts, totals, problems, report, dry_run)."""
    E.require_perm("import", "استيراد قائمة دخل HansaWorld")
    if src is None:
        raise UserError("اختر ملف قائمة الدخل (Excel / CSV / نص منسوخ من PDF).")
    pl = parse_pl_rows(_pl_rows(src))
    f = E.parse_date(period_from) if period_from else (pl["period"][0] if pl["period"] else None)
    t = E.parse_date(period_to) if period_to else (pl["period"][1] if pl["period"] else None)
    if not f or not t:
        raise UserError("لم أجد فترة القائمة في الملف (Period dd/mm/yyyy : dd/mm/yyyy): أدخل التاريخين.")
    if f > t:
        raise UserError("بداية فترة قائمة الدخل بعد نهايتها.")
    if not pl["accounts"]:
        raise UserError("لم أجد حسابات بقائمة الدخل (سطور تحتوي كود حساب 3xxxx/4xxxx/5xxxx ومبلغاً).")
    totals = E.hw_segment_totals(pl["accounts"])
    cols = [("المجموعة", "t"), ("الحساب", "t"), ("المبلغ", "m")]
    prev_rows = [("إيرادات" if c[:1] == "3" else ("تكلفة المبيعات" if c.startswith(E.HW_COGS_PREFIXES) else "مصروفات"),
                  f"{c} {n}", a) for c, (n, a) in sorted(pl["accounts"].items())]
    preview = Report(f"قائمة دخل HansaWorld {f} → {t}", cols, prev_rows, ("صافي (مدين موجب)", "",
                     round(sum(a for _, a in pl["accounts"].values()), 2)),
                     [f"المبيعات {fmt(totals['sales'])} • تكلفة المبيعات {fmt(totals['cogs'])} • مجمل الربح {fmt(totals['gross'])}"
                      f" • المصروفات {fmt(totals['opex'])} • صافي الربح {fmt(totals['net_profit'])}"] + pl["problems"], f"{f} → {t}")
    out = dict(period=(f, t), accounts=len(pl["accounts"]), totals=totals, problems=list(pl["problems"]),
               report=preview, dry_run=dry_run)
    if dry_run:
        return out
    E.backup_database("before_hw_pl")
    with db() as conn:
        setup_hansa(conn)
        conn.execute("DELETE FROM hw_pl_reference WHERE period_from=? AND period_to=?", (f, t))
        clash = conn.execute("SELECT DISTINCT period_from, period_to FROM hw_pl_reference WHERE NOT (period_to < ? OR "
                             "period_from > ?)", (f, t)).fetchall()
        if clash:
            raise UserError("فترة القائمة تتداخل مع قائمة دخل مستوردة سابقاً: " +
                            "، ".join(f"{a} → {b}" for a, b in clash) + ". احذفها أولاً.")
        conn.executemany("INSERT INTO hw_pl_reference (period_from, period_to, account_code, account_name, amount) "
                         "VALUES (?,?,?,?,?)", [(f, t, c, n, a) for c, (n, a) in pl["accounts"].items()])
        _resync_after_reference(conn, out["problems"])
    out["report"] = rep_hw_pl_compare({})
    return out


def remove_pl_reference(period_from, period_to):
    E.require_perm("import", "حذف قائمة دخل HansaWorld المرجعية")
    E.backup_database("before_hw_pl_remove")
    problems = []
    with db() as conn:
        n = conn.execute("DELETE FROM hw_pl_reference WHERE period_from=? AND period_to=?",
                         (period_from, period_to)).rowcount
        if n:
            _resync_after_reference(conn, problems)
    return n


def list_pl_references():
    with db() as conn:
        return _safe_pl_list(conn)


def _safe_pl_list(conn):
    try:
        return conn.execute("SELECT period_from, period_to, COUNT(*), ROUND(SUM(amount), 2) FROM hw_pl_reference "
                            "GROUP BY 1, 2 ORDER BY 1").fetchall()
    except sqlite3.OperationalError:
        return []


def _resync_after_reference(conn, problems):
    """بعد تغيير المراجع: إعادة بناء تسويات تكلفة المبيعات، وقيود تجميع المصاريف إن كانت مُرحّلة."""
    sync_hw_closing_stock_entry(conn)
    if _resync_tb_aggregate(conn):
        problems.append(f"أُعيد بناء قيود «{TB_EXP_LABEL}» لكل فترة مرجعية.")


def hw_pl_compare_rows():
    """[(الفترة، البند، HansaWorld، النظام، الفرق، الحالة)] لكل فترة مرجعية + الإجمالي التراكمي، وملاحظات."""
    with db() as conn:
        segs = E.hw_ref_segments(conn)
        adjs = E.hw_cogs_ref_adjustments(conn)
    rows, notes = [], []
    blocks = [(f"{sg['d_from']} → {sg['d_to']} ({PL_SRC_LABELS[sg['src']]})", E.DASH_FROM if i == 0 else sg["d_from"],
               sg["d_to"], sg["amounts"]) for i, sg in enumerate(segs)]
    if len(segs) > 1:
        allacc = {}
        for sg in segs:
            for c, (n, a) in sg["amounts"].items():
                allacc[c] = (n, round(allacc.get(c, (n, 0.0))[1] + a, 2))
        blocks.append((f"التراكمي {segs[0]['d_from']} → {segs[-1]['d_to']} (ميزان المراجعة)", E.DASH_FROM,
                       segs[-1]["d_to"], allacc))
    for label, a, b, amounts in blocks:
        t = E.hw_segment_totals(amounts)
        f = E.income_figures(a, b)
        og = f["opex_groups"]
        items = [("المبيعات (31xxx) ↔ صافي المبيعات", t["sales"], f["net_sales"]),
                 ("الإيرادات الأخرى (33xxx فروق عملة وخدمات)", t["other_income"], f["total_other"]),
                 ("تكلفة المبيعات (41110 + 41201)", t["cogs"], f["cogs"]),
                 ("مجمل الربح", t["gross"], f["gross"]),
                 ("مصروفات بيعية وتسويقية (43xxx)", t["selling"], og["selling"]),
                 ("مصروفات إدارية وعمومية (5xxxx)", t["admin"], og["admin"]),
                 ("مصروفات أخرى (41125 إيجار سيارات)", t["other_exp"], og["other"])]
        if abs(f["dep"]) > 0.005:
            items.append(("مصروف الإهلاك (بالنظام فقط)", 0.0, f["dep"]))
        items.append(("صافي الربح", t["net_profit"], f["net_profit"]))
        sales_diff = round(f["net_sales"] - t["sales"], 2)
        for name, hwv, sysv in items:
            diff = round(sysv - hwv, 2)
            if abs(diff) <= 0.01:
                st = "✅ مطابق"
            elif name.startswith("المبيعات") and abs(diff) <= max(1.0, abs(hwv) * 0.0001):
                st = "ℹ️ فرق دفتر HansaWorld عن ملف العملاء"
            elif name in ("مجمل الربح", "صافي الربح") and abs(diff - sales_diff) <= 0.01:
                st = "ℹ️ = فرق المبيعات"
            elif abs(sysv) <= 0.005 and abs(hwv) > 0.005 and "تكلفة" not in name and "مجمل" not in name:
                st = "⏳ غير مُرحّل"
            else:
                st = "❌ فرق"
            rows.append((label, name, round(hwv, 2), round(sysv, 2), diff, st))
    if not segs:
        notes.append("لا توجد مراجع: استورد «ميزان المراجعة التحليلي» من HansaWorld (وقوائم الدخل لفترات إن وجدت).")
    for a in adjs:
        sg = a["seg"]
        notes.append(f"تكلفة {sg['d_from']} → {sg['d_to']}: جرد دوري = مشتريات {fmt(a['purchases'])} − إشعارات موردين "
                     f"{fmt(a['credits'])} − مخزون ختامي {fmt(a['closing'])} = {fmt(a['periodic'])}؛ دفتر HansaWorld "
                     f"{fmt(a['ref'])} ← تسوية {fmt(a['adj'])} بتاريخ {a['date']}.")
    if adjs:
        notes.append("سبب التسوية: فواتير HansaWorld بلا بنود أصناف فتُحسب تكلفتها بالجرد الدوري، وهو لا يعرف مخزون بداية "
                     "الفترات الوسيطة ولا البضاعة المستلمة غير المفوترة (22380) ولا فروق مردودات المشتريات (41201)؛ "
                     "فتُعتمد تكلفة دفتر HansaWorld لكل فترة.")
    notes.append("فرق المبيعات الصغير: المبيعات بالنظام من ملف حركات العملاء (الفواتير والإشعارات بأرصدتها المطابقة)، "
                 "بينما دفتر HansaWorld العام (31110 + ضريبة 22502) أعلى منه بنفس فرق حساب العملاء 11210–11280 مع ملف "
                 "العملاء (515.92 في الميزان الحالي = إيراد + ضريبته). فرق داخلي في HansaWorld بين دفتره وكشف العملاء، فلا "
                 "يُفرض تعديل حتى لا تختل أرصدة العملاء.")
    return rows, notes


def rep_hw_pl_compare(p):
    rows, notes = hw_pl_compare_rows()
    cols = [("الفترة", "t"), ("البند", "t"), ("HansaWorld", "m"), ("النظام", "m"), ("الفرق (النظام − HansaWorld)", "m"),
            ("الحالة", "t")]
    return Report("مطابقة قائمة الدخل مع HansaWorld", cols, rows, None, notes, "حسب الفترات المرجعية")


def _register_pl_report():
    group = next((items for title, items in E.REPORT_CATALOG if "الملخص التنفيذي" in title), None)
    if group is None or "hw_pl" in E.REPORT_INDEX:
        return
    item = ("hw_pl", "مطابقة قائمة الدخل مع HansaWorld", rep_hw_pl_compare,
            "لكل فترة مرجعية (الميزان وقوائم الدخل المستوردة): المبيعات والتكلفة ومجمل الربح والمصروفات البيعية "
            "والإدارية وصافي الربح بالنظام مقابل HansaWorld، والفرق وسببه.")
    group.append(item)
    E.REPORT_INDEX[item[0]] = item[1:]
    E.REPORT_GROUP_OF[item[0]] = E.REPORT_GROUP_OF.get("kpi", "reports.exec")


_register_pl_report()

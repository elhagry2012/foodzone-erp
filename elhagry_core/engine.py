# -*- coding: utf-8 -*-
"""
نظام الحاجري المالي | ELHAGRY Financial Systems — المحرك المحاسبي (elhagry_core.engine)

لا يعتمد على أي واجهة: تستخدمه واجهة سطح المكتب (elhagry_desktop) وواجهة الويب (streamlit_app.py).

المكتبات المطلوبة:
    pip install -r requirements.txt

التشغيل:
    واجهة سطح المكتب (Tkinter):   python foodzone_erp.py
    واجهة الويب (Streamlit):       streamlit run streamlit_app.py

الملفات بجوار البرنامج (المسارات نسبية لمجلد البرنامج):
    charts of account.xlsx   دليل الحسابات (يُستورد عند أول تشغيل)
    elhagry.png              شعار الحاجري (شاشة الدخول)
    Amiri.ttf               خط عربي للـ PDF (أو يُستخدم Tahoma/Arial من ويندوز)
    accounting.db           قاعدة البيانات (تُنشأ تلقائياً؛ يمكن تغيير مكانها بمتغير البيئة ELHAGRY_DB)
"""
import base64
import contextvars
import hashlib
import hmac
import io
import math
import re
import sys
import os
import sqlite3
from collections import defaultdict
from contextlib import contextmanager, nullcontext
from datetime import datetime, date, timedelta
from xml.sax.saxutils import escape

import pandas as pd
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.barcode.qr import QrCodeWidget

# دعم تشكيل الحروف العربية داخل ملفات الـ PDF
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_AR_LIBS = True
except Exception:
    HAS_AR_LIBS = False

# ============================================================
# الإعدادات العامة
# ============================================================
# كل المسارات نسبية لمجلد البرنامج (وليس لمجلد التشغيل الحالي) حتى تعمل محلياً وعلى Streamlit Cloud بنفس الشكل.
# عند التجميع بـ PyInstaller تُحفظ البيانات بجوار ملف exe. يمكن تغيير مكان قاعدة البيانات بمتغير البيئة ELHAGRY_DB.
# المحرك داخل الحزمة elhagry_core/؛ مجلد البرنامج = المجلد الأعلى منها (حيث قاعدة البيانات والدليل والشعارات)
APP_DIR = (os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False)
           else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def load_env_file(*paths):
    """يحمّل الإعدادات الحساسة من ملف .env (KEY=VALUE) دون استبدال متغيرات البيئة الموجودة فعلاً.
    الملف محلي فقط (في .gitignore) — انظر .env.example. ELHAGRY_NO_DOTENV=1 يعطّل التحميل (الاختبارات)."""
    if os.environ.get("ELHAGRY_NO_DOTENV") == "1":
        return
    for path in paths:
        try:
            with open(path, encoding="utf-8-sig") as fh:
                lines = fh.read().splitlines()
        except OSError:
            continue
        for ln in lines:
            ln = ln.strip()
            if not ln or ln.startswith("#") or "=" not in ln:
                continue
            k, v = ln.split("=", 1)
            k, v = k.strip(), v.strip()
            if k.startswith("export "):
                k = k[7:].strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
                v = v[1:-1]
            if k:
                os.environ.setdefault(k, v)


load_env_file(os.path.join(APP_DIR, ".env"))
DATA_DIR = os.environ.get("ELHAGRY_DATA_DIR") or APP_DIR
if os.path.abspath(DATA_DIR) != os.path.abspath(APP_DIR):
    load_env_file(os.path.join(DATA_DIR, ".env"))
DB_FILE = os.environ.get("ELHAGRY_DB") or os.path.join(DATA_DIR, 'accounting.db')
EXCEL_PATH = os.path.join(APP_DIR, 'charts of account.xlsx')
BACKUP_DIR = os.path.join(DATA_DIR, 'backups')
# المرفقات تُحفظ كملفات في مجلد (خارج git) وقاعدة البيانات تحفظ بياناتها ومسارها فقط
ATTACH_DIR = os.environ.get("ELHAGRY_ATTACH_DIR") or os.path.join(DATA_DIR, 'attachments')
VAT_RATE = 0.15
EPS = 1e-6

SCHEMA_VERSION = 12

BANK_ACCOUNTS = ["صندوق نقدية الرياض", "بنك الراجحى", "بنك الانماء", "البنك الاهلى"]
# ربط حسابات الخزينة بأكوادها في دليل الحسابات (احتياطي لو الاسم غير موجود بالدليل)
DEFAULT_CASH_CODES = {"صندوق نقدية الرياض": "11111", "بنك الراجحى": "11121", "بنك الانماء": "11122", "البنك الاهلى": "11123"}
CASH_PREFIX = "111"        # النقدية بالصندوق والبنوك
CASH_BOX_PREFIX = "1111"   # الصناديق (الباقي بنوك)
CAT_PARTNER = "شريك تجاري"
CAT_EXPENSE = "مصروفات عمومية"
CAT_ASSET = "أصول ثابتة"
CAT_EMPLOYEE = "ذمم وعهد العاملين"
EMPLOYEE_PREFIXES = ("1141", "1142")   # 11410 ذمم العاملين، 11420 عهد العاملين وحساباتهما الفرعية

# أنواع الحسابات القياسية
TYPE_ASSET, TYPE_LIAB, TYPE_EQUITY, TYPE_REV, TYPE_EXP = "أصول", "خصوم", "حقوق ملكية", "إيرادات", "مصروفات"
ACCOUNT_TYPES = [TYPE_ASSET, TYPE_LIAB, TYPE_EQUITY, TYPE_REV, TYPE_EXP]
# حسابات يحسبها النظام تلقائياً (تكلفة المبيعات، مردودات المشتريات، الإهلاك): لا تُضاف من القيود لقائمة الدخل منعاً للازدواج
SYSTEM_PL_PREFIXES = ("4111", "412", "42")
# أكواد الشركاء القديمة المؤرشفة (OLD-...) بعد إعادة الترقيم بأكواد HansaWorld: تبقى بقاعدة البيانات للمرجعية والمطابقة
# ولا تظهر في القوائم والتقارير ولا تدخل في ربط العميل بالمورد.
ARCHIVED_PARTNER_PREFIX = "OLD-"
ACTIVE_PARTNER_SQL = f"partner_code NOT LIKE '{ARCHIVED_PARTNER_PREFIX}%'"
# حسابات لا تُصرف نقداً بسند مصروفات
NON_CASH_EXPENSE_PREFIXES = SYSTEM_PL_PREFIXES + ("413",)
DEFAULT_EXPENSES = ["إيجار", "رواتب وأجور", "كهرباء ومياه", "مواصلات ونقل",
                    "صيانة", "اتصالات وإنترنت", "مصروفات متنوعة"]
SALES_RETURN = "مرتجع مبيعات"
PURCHASE_RETURN = "مرتجع مشتريات"

MIGRATION_NOTE = None

# ============================================================
# المستخدمون والصلاحيات
# ============================================================
BRAND_AR = "نظام الحاجري المالي"
BRAND_EN = "ELHAGRY Financial Systems"
BRAND_FULL = f"{BRAND_AR} | {BRAND_EN}"
# لا توجد بيانات دخول مكتوبة في الكود: المدير الرئيسي يُحفظ اسمه في app_settings (main_admin_user)، وكلمة مروره مشفرة
# (PBKDF2) في جدول المستخدمين. لقاعدة جديدة: ELHAGRY_ADMIN_USER / ELHAGRY_ADMIN_PASSWORD من .env أو Secrets،
# وإلا تُولَّد كلمة مرور عشوائية وتُكتب مرة واحدة في initial_admin_password.txt بمجلد البيانات.
MAIN_ADMIN_KEY = "main_admin_user"
INITIAL_PASSWORD_FILE = "initial_admin_password.txt"
PBKDF2_ITERATIONS = 200_000
MAX_FAILED_LOGINS = 5
LOCKOUT_MINUTES = 5
# (مفتاح الصلاحية = مفتاح الصفحة، الوصف)
PERMISSIONS = [
    ("dashboard", "لوحة المؤشرات"), ("customers", "العملاء"), ("suppliers", "الموردون"),
    ("import", "مركز استيراد البيانات"),
    ("inventory", "المخزون وكارت الصنف"),
    ("sales", "المبيعات"), ("purchase", "المشتريات"), ("cash", "الخزينة والبنوك والسداد"),
    ("returns", "المرتجعات"), ("assets", "الأصول الثابتة"), ("statement", "كشوف الحساب"),
    ("income", "قائمة الدخل"), ("reports", "التقارير وتحليل القرار"), ("journal", "القيود اليومية"),
    ("discounts", "سجل الخصومات اللاحقة وأثرها (الحسابات)"), ("reps", "سجل المناديب"),
    ("coa", "دليل الحسابات"), ("company", "بيانات الشركة والأمان"), ("closure", "الإغلاق اليومي"),
    ("expiry", "الصلاحية والأصناف الراكدة"), ("modify", "تعديل وحذف المستندات والأرصدة (حساسة)"),
    ("ledger", "الدفتر العام والقوائم المالية (حالة الترحيل الآلي)"),
    ("planning", "التخطيط المالي (التدفق النقدي، الموازنة، العمولات، حدود الائتمان)"),
    ("audit", "سجل التدقيق والرقابة الداخلية"),
]
PERMISSION_KEYS = [k for k, _ in PERMISSIONS]


class UserError(Exception):
    """خطأ متوقع يُعرض للمستخدم كتنبيه (وليس كعطل في البرنامج)."""


# ============================================================
# الخط العربي وتجهيز النصوص للـ PDF
# ============================================================
FONT_NAME = 'Helvetica'


def register_arabic_font():
    global FONT_NAME
    candidates = [
        os.path.join(APP_DIR, 'Amiri.ttf'),
        os.path.join(APP_DIR, 'Amiri-Regular.ttf'),
        'C:/Windows/Fonts/Tahoma.ttf',
        'C:/Windows/Fonts/Arial.ttf',
        'C:/Windows/Fonts/trado.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('ArabicFont', path))
                FONT_NAME = 'ArabicFont'
                return
            except Exception as e:
                print(f"تعذر تسجيل الخط {path}: {e}")
    print("تحذير: لم يتم العثور على خط عربي، ضع ملف Amiri-Regular.ttf بجوار البرنامج.")


register_arabic_font()


_AR_RE = re.compile('[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]')


def ar(text):
    """تشكيل الحروف العربية وترتيبها من اليمين لليسار للـ PDF.
    الاتجاه الأساسي RTL يُفرض صراحةً، والنصوص بدون حروف عربية (أرقام، تواريخ، إنجليزي) تُترك كما هي."""
    text = '' if text is None else str(text)
    if not HAS_AR_LIBS or not _AR_RE.search(text):
        return text
    try:
        shaped = arabic_reshaper.reshape(text)
        try:
            return get_display(shaped, base_dir='R')
        except TypeError:  # إصدار قديم من المكتبة لا يدعم base_dir
            return get_display(shaped)
    except Exception:
        return text


def P(text, style):
    """فقرة PDF آمنة: تشكيل عربي + تهريب رموز XML + دعم تعدد الأسطر."""
    lines = ('' if text is None else str(text)).split('\n')
    return Paragraph('<br/>'.join(escape(ar(line)) for line in lines), style)


def pdf_styles():
    def st(name, size=9, align=2, color="#0f172a"):
        return ParagraphStyle(name, fontName=FONT_NAME, fontSize=size,
                              leading=size * 1.6, alignment=align,
                              textColor=colors.HexColor(color))
    return {
        'right': st('right', 9, 2),
        'sub': st('sub', 9, 2, "#475569"),
        'cell': st('cell', 8, 1),
        'head': st('head', 8, 1, "#ffffff"),
        'title': st('title', 14, 1),
        'left': st('left', 9, 0),
    }


def rtl(seq):
    """عكس ترتيب الأعمدة ليظهر أول عمود في أقصى اليمين."""
    return list(reversed(seq))


# ============================================================
# أدوات مساعدة عامة
# ============================================================
AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩٫٬", "0123456789.,")


def normalize_digits(s):
    return str(s).translate(AR_DIGITS)


def parse_float(s):
    try:
        v = float(normalize_digits(s).replace(",", "").strip())
    except (ValueError, TypeError):
        return None
    return v if math.isfinite(v) else None


def parse_date(s):
    try:
        return datetime.strptime(normalize_digits(s).strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def today():
    return datetime.now().strftime("%Y-%m-%d")


def fmt(x):
    return f"{(x or 0):,.2f}"


def make_line(code, name, unit, qty, price, discount=0.0):
    """سطر فاتورة/أمر. discount = الخصم الخاص على السطر (مبلغ قبل الضريبة).
    المبلغ الأساسي = الكمية × السعر - الخصم، والضريبة 15% تُحسب على المبلغ بعد الخصم."""
    qty = round(qty, 3)
    gross = round(qty * price, 2)
    discount = round(discount or 0.0, 2)
    if discount < 0:
        raise UserError("الخصم لا يمكن أن يكون سالباً.")
    if discount > gross + 0.005:
        raise UserError(f"الخصم ({fmt(discount)}) أكبر من قيمة السطر ({fmt(gross)}).")
    discount = min(discount, gross)
    subtotal = round(gross - discount, 2)
    vat = round(subtotal * VAT_RATE, 2)
    return {"code": code, "name": name, "unit": unit or "", "qty": qty, "price": price, "gross": gross,
            "discount": discount, "subtotal": subtotal, "vat": vat, "total": round(subtotal + vat, 2)}


def parse_discount(text, gross):
    """الخصم الخاص: رقم = مبلغ خصم على السطر كله، أو نسبة مثل 10% من قيمة السطر."""
    t = normalize_digits(text or "").replace("٪", "%").replace(",", "").strip()
    if not t:
        return 0.0
    if t.endswith("%"):
        pct = parse_float(t[:-1])
        if pct is None or pct < 0 or pct > 100:
            raise UserError("نسبة الخصم يجب أن تكون بين 0% و 100%.")
        return round(gross * pct / 100, 2)
    v = parse_float(t)
    if v is None or v < 0:
        raise UserError("أدخل الخصم الخاص كمبلغ (مثل 50) أو نسبة (مثل 10%).")
    return round(v, 2)


# ============================================================
# قاعدة البيانات
# ============================================================
# مهلة انتظار القفل: في الويب يعمل أكثر من مستخدم على نفس الملف؛ ننتظر بدلاً من فشل فوري «database is locked»
DB_TIMEOUT = 30


# المستخدم الحالي لسجل التدقيق: لكل جلسة/خيط قيمته (الويب: جلسة Streamlit، سطح المكتب: المستخدم المسجّل)
_CURRENT_USER = contextvars.ContextVar("elhagry_current_user", default=None)
# دوال تُستدعى قبل حفظ أي معاملة فيها تعديل (مثل الترحيل الآلي للدفتر العام) وبعد تهيئة قاعدة البيانات
COMMIT_HOOKS = []
INIT_HOOKS = []
CONNECT_HOOKS = []   # تسجيل دوال SQL إضافية على كل اتصال (مثل بصمة سجل التدقيق)


_CURRENT_PERMS = contextvars.ContextVar("elhagry_current_perms", default=None)


def set_current_user(username, source=None, user=None):
    """تعيين المستخدم الحالي (يُسجَّل في سجل التدقيق مع كل تعديل/حذف). source: desktop / web.
    user: قاموس الجلسة (is_admin, perms) لفحص الصلاحيات داخل المحرك نفسه (لا تكفي إخفاء الأزرار)."""
    _CURRENT_USER.set((username or None, source))
    _CURRENT_PERMS.set(user)


def require_perm(key, action=""):
    """يمنع العملية لو المستخدم الحالي غير مخوّل. بدون مستخدم مسجّل (النظام/الاستيراد/الاختبارات) = مسموح."""
    u = _CURRENT_PERMS.get()
    if u is None or user_can(u, key):
        return
    raise UserError(f"ليس لديك صلاحية «{SUB_LABELS.get(key) or dict(PERMISSIONS).get(key, key)}»"
                    + (f" لتنفيذ: {action}" if action else "") + ".\nراجع مدير النظام.")


def current_user_name():
    v = _CURRENT_USER.get()
    return (v[0] if v and v[0] else "النظام")


def current_user_source():
    v = _CURRENT_USER.get()
    return (v[1] if v and v[1] else "")


def _register_functions(conn):
    conn.create_function("app_user", 0, current_user_name)
    conn.create_function("app_source", 0, current_user_source)
    for hook in CONNECT_HOOKS:
        hook(conn)


@contextmanager
def db():
    """اتصال آمن: حفظ تلقائي عند النجاح، تراجع كامل عند أي خطأ، وإغلاق دائماً.
    قبل الحفظ تُنفَّذ COMMIT_HOOKS داخل نفس المعاملة (الترحيل الآلي للدفتر العام)، فالمستند وقيده يُحفظان معاً."""
    conn = sqlite3.connect(DB_FILE, timeout=DB_TIMEOUT)
    conn.execute("PRAGMA foreign_keys = ON")
    _register_functions(conn)
    try:
        yield conn
        if conn.in_transaction:
            for hook in COMMIT_HOOKS:
                hook(conn)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS company_profile (
    id INTEGER PRIMARY KEY,
    company_name TEXT, cr_number TEXT, vat_number TEXT,
    phone TEXT, address TEXT, rep_name TEXT
);

CREATE TABLE IF NOT EXISTS chart_of_accounts (
    account_code TEXT PRIMARY KEY,
    account_name TEXT NOT NULL,
    account_type TEXT NOT NULL,
    parent_code TEXT
);

CREATE TABLE IF NOT EXISTS business_partners (
    partner_id INTEGER PRIMARY KEY AUTOINCREMENT,
    partner_code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    phone TEXT, email TEXT, address TEXT,
    cr_number TEXT, vat_number TEXT,
    is_customer INTEGER DEFAULT 1,
    is_supplier INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS inventory (
    item_code TEXT PRIMARY KEY,
    item_name TEXT NOT NULL,
    unit TEXT DEFAULT 'KG',
    opening_qty REAL DEFAULT 0,
    opening_cost REAL DEFAULT 0,
    quantity REAL DEFAULT 0,
    cost_price REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS sales_invoices (
    invoice_id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_date TEXT NOT NULL,
    customer_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    customer_name TEXT NOT NULL,
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    paid_amount REAL NOT NULL DEFAULT 0 CHECK (paid_amount >= 0),
    status TEXT NOT NULL DEFAULT 'مفتوحة',
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS sales_invoice_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_id INTEGER NOT NULL REFERENCES sales_invoices(invoice_id) ON DELETE CASCADE,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_price REAL NOT NULL CHECK (unit_price >= 0),
    unit_cost REAL NOT NULL DEFAULT 0,
    subtotal REAL NOT NULL,
    vat_amount REAL NOT NULL,
    total_amount REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS purchase_invoices (
    purchase_id INTEGER PRIMARY KEY AUTOINCREMENT,
    purchase_date TEXT NOT NULL,
    supplier_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    supplier_name TEXT NOT NULL,
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    paid_amount REAL NOT NULL DEFAULT 0 CHECK (paid_amount >= 0),
    status TEXT NOT NULL DEFAULT 'مفتوحة',
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS purchase_invoice_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    purchase_id INTEGER NOT NULL REFERENCES purchase_invoices(purchase_id) ON DELETE CASCADE,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_cost REAL NOT NULL CHECK (unit_cost >= 0),
    subtotal REAL NOT NULL,
    vat_amount REAL NOT NULL,
    total_amount REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS returns_invoices (
    return_id INTEGER PRIMARY KEY AUTOINCREMENT,
    return_date TEXT NOT NULL,
    return_type TEXT NOT NULL CHECK (return_type IN ('مرتجع مبيعات','مرتجع مشتريات')),
    party_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    party_name TEXT NOT NULL,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT NOT NULL,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_price REAL NOT NULL CHECK (unit_price >= 0),
    unit_cost REAL NOT NULL DEFAULT 0,
    subtotal REAL NOT NULL,
    vat_amount REAL NOT NULL,
    total_amount REAL NOT NULL,
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS cash_transactions (
    trans_id INTEGER PRIMARY KEY AUTOINCREMENT,
    trans_type TEXT NOT NULL CHECK (trans_type IN ('قبض','صرف')),
    trans_date TEXT NOT NULL,
    account_bank TEXT NOT NULL,
    category TEXT NOT NULL,
    party_code TEXT,
    party_name TEXT NOT NULL,
    amount REAL NOT NULL CHECK (amount > 0),
    alloc_type TEXT,
    alloc_id INTEGER,
    description TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS fixed_assets (
    asset_id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_code TEXT UNIQUE NOT NULL,
    asset_name TEXT NOT NULL,
    purchase_date TEXT NOT NULL,
    purchase_cost REAL NOT NULL CHECK (purchase_cost > 0),
    salvage_value REAL DEFAULT 0,
    useful_life_years INTEGER NOT NULL CHECK (useful_life_years > 0),
    monthly_depreciation REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS daily_closures (
    closure_id INTEGER PRIMARY KEY AUTOINCREMENT,
    closure_date TEXT UNIQUE NOT NULL,
    total_sales REAL,
    total_purchases REAL,
    net_cash_movement REAL,
    closed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_sal_lines_inv ON sales_invoice_lines(invoice_id);
CREATE INDEX IF NOT EXISTS idx_sal_lines_item ON sales_invoice_lines(item_code);
CREATE INDEX IF NOT EXISTS idx_pur_lines_inv ON purchase_invoice_lines(purchase_id);
CREATE INDEX IF NOT EXISTS idx_pur_lines_item ON purchase_invoice_lines(item_code);
CREATE INDEX IF NOT EXISTS idx_sal_cust ON sales_invoices(customer_code);
CREATE INDEX IF NOT EXISTS idx_pur_supp ON purchase_invoices(supplier_code);
CREATE INDEX IF NOT EXISTS idx_cash_party ON cash_transactions(party_code);
CREATE INDEX IF NOT EXISTS idx_ret_party ON returns_invoices(party_code);
"""


def archive_legacy_db():
    """لو الملف القديم بالهيكل القديم (فاتورة = صنف واحد) نحفظه ونبدأ بقاعدة جديدة."""
    if not os.path.exists(DB_FILE):
        return None
    conn = sqlite3.connect(DB_FILE)
    try:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(sales_invoices)")]
    finally:
        conn.close()
    if cols and 'customer_code' not in cols:
        name = os.path.join(os.path.dirname(os.path.abspath(DB_FILE)), f"accounting_legacy_{datetime.now():%Y%m%d_%H%M%S}.db")
        os.rename(DB_FILE, name)
        return name
    return None


def import_from_legacy(legacy_file):
    """نقل الشركاء والأصناف (بأرصدتها الحالية كرصيد افتتاحي) وبيانات الشركة."""
    src = sqlite3.connect(legacy_file)
    try:
        with db() as conn:
            try:
                for row in src.execute(
                        "SELECT partner_code, name, phone, email, address, cr_number, vat_number, "
                        "is_customer, is_supplier FROM business_partners"):
                    conn.execute(
                        "INSERT OR IGNORE INTO business_partners (partner_code, name, phone, email, "
                        "address, cr_number, vat_number, is_customer, is_supplier) "
                        "VALUES (?,?,?,?,?,?,?,?,?)", row)
            except Exception as e:
                print(f"تعذر نقل الشركاء: {e}")
            try:
                for code, name, qty, cost in src.execute(
                        "SELECT item_code, item_name, quantity, cost_price FROM inventory"):
                    qty = max(qty or 0.0, 0.0)
                    cost = cost or 0.0
                    conn.execute(
                        "INSERT OR IGNORE INTO inventory (item_code, item_name, unit, opening_qty, "
                        "opening_cost, quantity, cost_price) VALUES (?,?,?,?,?,?,?)",
                        (code, name, 'KG', qty, cost, qty, cost))
            except Exception as e:
                print(f"تعذر نقل الأصناف: {e}")
            try:
                comp = src.execute("SELECT company_name, cr_number, vat_number, phone, address "
                                   "FROM company_profile WHERE id = 1").fetchone()
                if comp:
                    conn.execute("UPDATE company_profile SET company_name=?, cr_number=?, vat_number=?, "
                                 "phone=?, address=? WHERE id=1", comp)
            except Exception as e:
                print(f"تعذر نقل بيانات الشركة: {e}")
    finally:
        src.close()


_EXCEL_TYPES = (("asset", TYPE_ASSET), ("liabilit", TYPE_LIAB), ("equity", TYPE_EQUITY),
                ("income", TYPE_REV), ("revenue", TYPE_REV), ("expense", TYPE_EXP))


def classify_account(code, hint=""):
    """تصنيف الحساب حسب ترقيم الدليل: 1 أصول، 21 حقوق ملكية، 2 خصوم، 3 إيرادات، 4/5 مصروفات، 9 حقوق ملكية.
    عمود Type في ملف الإكسيل غير موثوق (يصنّف رأس المال والأرباح المرحلة كخصوم)، فيُستخدم فقط لو الترقيم غير معروف."""
    code = str(code).strip()
    if code[:1].isdigit():
        first = code[0]
        if first == "1":
            return TYPE_ASSET
        if first == "2":
            return TYPE_EQUITY if code.startswith("21") else TYPE_LIAB
        if first == "3":
            return TYPE_REV
        if first in "45":
            return TYPE_EXP
        if first == "9":
            return TYPE_EQUITY
    if hint in ACCOUNT_TYPES:
        return hint
    h = str(hint).lower()
    for key, typ in _EXCEL_TYPES:
        if key in h:
            return typ
    return TYPE_ASSET


def _parent_from_groups(code, group_codes):
    """أقرب حساب رئيسي يبدأ كود الحساب ببادئته (10000 → 1، 11400 → 114)."""
    best, best_len = None, -1
    for g in group_codes:
        prefix = g.rstrip("0") or g
        if g != code and code.startswith(prefix) and len(prefix) > best_len:
            best, best_len = g, len(prefix)
    return best


def import_chart_of_accounts():
    """قراءة دليل الحسابات من الإكسيل وتحديثه (upsert).
    لا يحذف الحسابات المضافة يدوياً من البرنامج. يرجع عدد الحسابات المقروءة، أو None لو تعذرت القراءة."""
    if not os.path.exists(EXCEL_PATH):
        return None
    try:
        df = pd.read_excel(EXCEL_PATH, header=1)
    except Exception as e:
        print(f"خطأ في قراءة ملف الإكسيل: {e}")
        return None
    if 'Account' not in df.columns or 'Name' not in df.columns:
        print("ملف الإكسيل لا يحتوي على الأعمدة Account و Name")
        return None

    rows = []
    for _, row in df.iterrows():
        code = str(row['Account']).strip() if pd.notna(row['Account']) else ''
        name = str(row['Name']).strip() if pd.notna(row['Name']) else ''
        if code.endswith('.0'):
            code = code[:-2]
        if not code or not name or code.lower() == 'nan' or not code.replace('.', '').isdigit():
            continue
        val_e = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ''
        val_f = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ''
        # في الملف: العمود E = 1 يعني حساب رئيسي (تجميعي) ونوعه في العمود F، غير ذلك النوع في العمود E
        is_group = val_e in ("1", "1.0")
        rows.append([code, name, classify_account(code, f"{val_e} {val_f}"), None, int(is_group)])

    if not rows:
        return None
    groups = [r[0] for r in rows if r[4]]
    for r in rows:
        r[3] = _parent_from_groups(r[0], groups)
    with db() as conn:
        conn.executemany(
            "INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
            "VALUES (?,?,?,?,?) ON CONFLICT(account_code) DO UPDATE SET account_name=excluded.account_name, "
            "account_type=excluded.account_type, parent_code=excluded.parent_code, is_group=excluded.is_group",
            rows)
        normalize_account_types(conn)
    return len(rows)


def normalize_account_types(conn):
    """تصحيح أي نوع حساب غير قياسي (مثل «ذمم العاملين» أو «أصول ثابتة») وتعليم الحسابات التي لها فروع كحسابات رئيسية."""
    for code, typ in conn.execute("SELECT account_code, account_type FROM chart_of_accounts").fetchall():
        if typ not in ACCOUNT_TYPES:
            conn.execute("UPDATE chart_of_accounts SET account_type=? WHERE account_code=?",
                         (classify_account(code, typ), code))
    conn.execute("UPDATE chart_of_accounts SET is_group=1 WHERE account_code IN "
                 "(SELECT parent_code FROM chart_of_accounts WHERE parent_code IS NOT NULL)")


# ============================================================
# هيكل البيانات - الإصدار 3: أمر ← تسليم/استلام ← فاتورة (مع دعم التنفيذ الجزئي)
# ============================================================
SCHEMA_SQL_V3 = """
CREATE TABLE IF NOT EXISTS sales_orders (
    order_id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_date TEXT NOT NULL,
    expected_date TEXT,
    customer_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    customer_name TEXT NOT NULL,
    sales_rep TEXT,
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'جديد',
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sales_order_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL REFERENCES sales_orders(order_id) ON DELETE CASCADE,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_price REAL NOT NULL CHECK (unit_price >= 0),
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    delivered_qty REAL NOT NULL DEFAULT 0,
    invoiced_qty REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS sales_deliveries (
    delivery_id INTEGER PRIMARY KEY AUTOINCREMENT,
    delivery_date TEXT NOT NULL,
    order_id INTEGER NOT NULL REFERENCES sales_orders(order_id),
    customer_code TEXT NOT NULL,
    customer_name TEXT NOT NULL,
    notes TEXT,
    invoice_id INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sales_delivery_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    delivery_id INTEGER NOT NULL REFERENCES sales_deliveries(delivery_id) ON DELETE CASCADE,
    order_line_id INTEGER NOT NULL REFERENCES sales_order_lines(line_id),
    item_code TEXT NOT NULL,
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_price REAL NOT NULL,
    unit_cost REAL NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS purchase_orders (
    order_id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_date TEXT NOT NULL,
    expected_date TEXT,
    supplier_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    supplier_name TEXT NOT NULL,
    purchase_rep TEXT,
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'جديد',
    notes TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS purchase_order_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER NOT NULL REFERENCES purchase_orders(order_id) ON DELETE CASCADE,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_cost REAL NOT NULL CHECK (unit_cost >= 0),
    subtotal REAL NOT NULL DEFAULT 0,
    vat_amount REAL NOT NULL DEFAULT 0,
    total_amount REAL NOT NULL DEFAULT 0,
    received_qty REAL NOT NULL DEFAULT 0,
    invoiced_qty REAL NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS purchase_receipts (
    receipt_id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_date TEXT NOT NULL,
    order_id INTEGER NOT NULL REFERENCES purchase_orders(order_id),
    supplier_code TEXT NOT NULL,
    supplier_name TEXT NOT NULL,
    supplier_ref TEXT,
    notes TEXT,
    invoice_id INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS purchase_receipt_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_id INTEGER NOT NULL REFERENCES purchase_receipts(receipt_id) ON DELETE CASCADE,
    order_line_id INTEGER NOT NULL REFERENCES purchase_order_lines(line_id),
    item_code TEXT NOT NULL,
    item_name TEXT NOT NULL,
    unit TEXT,
    quantity REAL NOT NULL CHECK (quantity > 0),
    unit_cost REAL NOT NULL CHECK (unit_cost >= 0)
);

CREATE TABLE IF NOT EXISTS journal_entries (
    entry_id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    description TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS journal_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id INTEGER NOT NULL REFERENCES journal_entries(entry_id) ON DELETE CASCADE,
    account_code TEXT NOT NULL,
    account_name TEXT NOT NULL,
    debit REAL NOT NULL DEFAULT 0,
    credit REAL NOT NULL DEFAULT 0,
    memo TEXT
);
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE INDEX IF NOT EXISTS idx_so_cust ON sales_orders(customer_code);
CREATE INDEX IF NOT EXISTS idx_sol_order ON sales_order_lines(order_id);
CREATE INDEX IF NOT EXISTS idx_sdl_deliv ON sales_delivery_lines(delivery_id);
CREATE INDEX IF NOT EXISTS idx_sdl_item ON sales_delivery_lines(item_code);
CREATE INDEX IF NOT EXISTS idx_po_supp ON purchase_orders(supplier_code);
CREATE INDEX IF NOT EXISTS idx_pol_order ON purchase_order_lines(order_id);
CREATE INDEX IF NOT EXISTS idx_prl_rec ON purchase_receipt_lines(receipt_id);
CREATE INDEX IF NOT EXISTS idx_prl_item ON purchase_receipt_lines(item_code);
CREATE INDEX IF NOT EXISTS idx_jl_entry ON journal_lines(entry_id);
"""

# أعمدة تُضاف للجداول القديمة (تغطي قاعدة النسخة السابقة وقاعدة النسخة الأحدث). آمنة للتكرار.
NEW_COLUMNS = [
    ("users", "role_id", "INTEGER"),
    ("attachments", "file_path", "TEXT"),   # مسار الملف داخل مجلد المرفقات (المرفقات القديمة تبقى داخل القاعدة)
    ("cash_transactions", "rep_name", "TEXT"),   # المندوب الموروث تلقائياً (للعمولات على التحصيلات)
    ("sales_invoice_lines", "delivery_line_id", "INTEGER"),
    ("purchase_invoice_lines", "receipt_line_id", "INTEGER"),
    ("sales_invoices", "source_type", "TEXT DEFAULT 'مباشر'"),
    ("purchase_invoices", "source_type", "TEXT DEFAULT 'مباشر'"),
    ("purchase_invoices", "supplier_ref", "TEXT"),
    ("sales_invoices", "order_id", "INTEGER"),
    ("purchase_invoices", "order_id", "INTEGER"),
    ("sales_invoices", "sales_rep", "TEXT"),
    ("purchase_invoices", "purchase_rep", "TEXT"),
    ("sales_orders", "expected_date", "TEXT"),
    ("sales_orders", "sales_rep", "TEXT"),
    ("purchase_orders", "expected_date", "TEXT"),
    ("purchase_orders", "purchase_rep", "TEXT"),
    ("sales_order_lines", "delivered_qty", "REAL NOT NULL DEFAULT 0"),
    ("sales_order_lines", "invoiced_qty", "REAL NOT NULL DEFAULT 0"),
    ("purchase_order_lines", "received_qty", "REAL NOT NULL DEFAULT 0"),
    ("purchase_order_lines", "invoiced_qty", "REAL NOT NULL DEFAULT 0"),
    ("chart_of_accounts", "is_group", "INTEGER NOT NULL DEFAULT 0"),
    # الخصم الخاص على مستوى الصنف (مبلغ قبل الضريبة)
    ("sales_invoice_lines", "discount_amount", "REAL NOT NULL DEFAULT 0"),
    ("sales_invoices", "discount_amount", "REAL NOT NULL DEFAULT 0"),
    ("sales_order_lines", "discount_amount", "REAL NOT NULL DEFAULT 0"),
    ("sales_delivery_lines", "discount_amount", "REAL NOT NULL DEFAULT 0"),
    # أرصدة افتتاحية للعملاء والموردين (ما عليه / ما له)
    ("business_partners", "opening_debit", "REAL NOT NULL DEFAULT 0"),
    ("business_partners", "opening_credit", "REAL NOT NULL DEFAULT 0"),
    ("business_partners", "opening_date", "TEXT"),
    # تاريخ الإنتاج والانتهاء لكل صنف في المشتريات
    ("purchase_invoice_lines", "production_date", "TEXT"),
    ("purchase_invoice_lines", "expiry_date", "TEXT"),
    ("purchase_receipt_lines", "production_date", "TEXT"),
    ("purchase_receipt_lines", "expiry_date", "TEXT"),
    # رقم الفاتورة التسلسلي المستقل
    ("sales_invoices", "invoice_no", "TEXT"),
    ("purchase_invoices", "invoice_no", "TEXT"),
    # بيانات الأصناف الموحدة: سعر البيع، الباركود، والكود في النظام السابق (مرجع الترحيل)
    ("inventory", "selling_price", "REAL NOT NULL DEFAULT 0"),
    ("inventory", "barcode", "TEXT"),
    ("inventory", "legacy_code", "TEXT"),
    ("business_partners", "legacy_code", "TEXT"),
    ("business_partners", "rep_name", "TEXT"),
    # بيانات المنشأة المصدرة للفواتير (ثنائية اللغة + البنك)
    ("company_profile", "company_name_en", "TEXT"),
    ("company_profile", "address_en", "TEXT"),
    ("company_profile", "bank_name", "TEXT"),
    ("company_profile", "iban", "TEXT"),
    # قيود يولّدها النظام تلقائياً (مثل قيود الخصم اللاحق) مرتبطة بمستندها ولا تُعدَّل يدوياً
    ("journal_entries", "source_type", "TEXT"),
    ("journal_entries", "source_id", "INTEGER"),
    # الحد الائتماني وشروط السداد للعملاء (0 = بدون حد) — سياسة التجاوز: warn تنبيه / block منع البيع الآجل
    ("business_partners", "credit_limit", "REAL NOT NULL DEFAULT 0"),
    ("business_partners", "credit_days", "INTEGER"),
    ("business_partners", "credit_policy", "TEXT"),
    # آخر سعر شراء للصنف (من ملف تكلفة HansaWorld «Last Price») — مرجع للتسعير فقط، لا يدخل في متوسط التكلفة
    ("inventory", "last_purchase_price", "REAL NOT NULL DEFAULT 0"),
]

# الخصم اللاحق بعد الشراء/البيع (إشعار خصم من المورد / إشعار دائن للعميل).
# method: stock = تخفيض تكلفة المخزون (مشتريات فقط) | pl = أثر مباشر على قائمة الدخل (خصم مكتسب / خصم مسموح به)
# post_discount_lines: توزيع خصم «تخفيض التكلفة» على الأصناف، و applied = الجزء الذي امتصه المخزون المتبقي فعلاً
# (الباقي يخص بضاعة سبق بيعها فيظهر كخصم مكتسب في قائمة الدخل).
POST_DISCOUNT_SQL = """
CREATE TABLE IF NOT EXISTS post_discounts (
    discount_id INTEGER PRIMARY KEY AUTOINCREMENT,
    discount_no TEXT UNIQUE,
    kind TEXT NOT NULL CHECK (kind IN ('purchase','sales')),
    discount_date TEXT NOT NULL,
    party_code TEXT NOT NULL REFERENCES business_partners(partner_code),
    party_name TEXT NOT NULL,
    invoice_id INTEGER,
    method TEXT NOT NULL DEFAULT 'pl' CHECK (method IN ('stock','pl')),
    subtotal REAL NOT NULL CHECK (subtotal > 0),
    vat_amount REAL NOT NULL DEFAULT 0 CHECK (vat_amount >= 0),
    total_amount REAL NOT NULL,
    stock_applied REAL NOT NULL DEFAULT 0,
    journal_entry_id INTEGER,
    reason TEXT,
    notes TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS post_discount_lines (
    line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    discount_id INTEGER NOT NULL REFERENCES post_discounts(discount_id) ON DELETE CASCADE,
    item_code TEXT NOT NULL REFERENCES inventory(item_code),
    item_name TEXT,
    amount REAL NOT NULL CHECK (amount > 0),
    applied REAL NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_pdisc_party ON post_discounts(kind, party_code);
CREATE INDEX IF NOT EXISTS idx_pdisc_inv ON post_discounts(kind, invoice_id);
CREATE INDEX IF NOT EXISTS idx_pdisc_lines_item ON post_discount_lines(item_code);
"""

SEQUENCES_SQL = """
CREATE TABLE IF NOT EXISTS sequences (
    name TEXT PRIMARY KEY,
    label TEXT,
    prefix TEXT NOT NULL DEFAULT '',
    next_value INTEGER NOT NULL DEFAULT 1,
    width INTEGER NOT NULL DEFAULT 4
);
"""

USERS_SQL = """
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    full_name TEXT,
    password_hash TEXT NOT NULL,
    is_admin INTEGER NOT NULL DEFAULT 0,
    permissions TEXT NOT NULL DEFAULT '',
    active INTEGER NOT NULL DEFAULT 1,
    failed_attempts INTEGER NOT NULL DEFAULT 0,
    locked_until TEXT,
    last_login TEXT,
    role_id INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
"""

# حركات المخزون الموحدة: الفاتورة المباشرة + التسليم/الاستلام.
# بنود الفواتير المولّدة من تسليم/استلام مستبعدة (حركتها مسجلة وقت التنفيذ) لتجنب الازدواج.
VIEWS_SQL = """
DROP VIEW IF EXISTS v_stock_in;
CREATE VIEW v_stock_in AS
    SELECT h.purchase_date AS mdate, h.created_at AS created_at, l.item_code AS item_code,
           l.quantity AS qty, l.unit_cost AS unit_price, l.subtotal AS value,
           h.supplier_name AS party, 'مشتريات (فاتورة)' AS mtype, h.notes AS notes,
           'فاتورة مشتريات ' || COALESCE(h.invoice_no, h.purchase_id) AS ref, h.supplier_code AS party_code,
           l.unit_cost AS unit_cost, 'purchase' AS doc_kind, h.purchase_id AS doc_id
    FROM purchase_invoice_lines l JOIN purchase_invoices h ON h.purchase_id = l.purchase_id
    WHERE l.receipt_line_id IS NULL
    UNION ALL
    SELECT r.receipt_date, r.created_at, l.item_code, l.quantity, l.unit_cost,
           ROUND(l.quantity * l.unit_cost, 2), r.supplier_name, 'استلام مخزني', r.notes,
           'إذن استلام ' || r.receipt_id, r.supplier_code, l.unit_cost, 'receipt', r.receipt_id
    FROM purchase_receipt_lines l JOIN purchase_receipts r ON r.receipt_id = l.receipt_id;

DROP VIEW IF EXISTS v_stock_out;
CREATE VIEW v_stock_out AS
    SELECT h.invoice_date AS mdate, h.created_at AS created_at, l.item_code AS item_code,
           l.quantity AS qty, l.unit_price AS unit_price, l.subtotal AS value,
           h.customer_name AS party, 'مبيعات (فاتورة)' AS mtype, h.notes AS notes,
           'فاتورة مبيعات ' || COALESCE(h.invoice_no, h.invoice_id) AS ref, h.customer_code AS party_code,
           l.unit_cost AS unit_cost, 'sales' AS doc_kind, h.invoice_id AS doc_id
    FROM sales_invoice_lines l JOIN sales_invoices h ON h.invoice_id = l.invoice_id
    WHERE l.delivery_line_id IS NULL
    UNION ALL
    SELECT d.delivery_date, d.created_at, l.item_code, l.quantity, l.unit_price,
           ROUND(l.quantity * l.unit_price - COALESCE(l.discount_amount, 0), 2), d.customer_name, 'تسليم للعميل', d.notes,
           'إذن تسليم ' || d.delivery_id, d.customer_code, l.unit_cost, 'delivery', d.delivery_id
    FROM sales_delivery_lines l JOIN sales_deliveries d ON d.delivery_id = l.delivery_id;

DROP VIEW IF EXISTS v_batches;
CREATE VIEW v_batches AS
    SELECT h.purchase_date AS mdate, h.created_at AS created_at, l.item_code AS item_code, l.quantity AS qty,
           l.unit_cost AS unit_cost, l.production_date AS production_date, l.expiry_date AS expiry_date,
           'فاتورة ' || COALESCE(h.invoice_no, h.purchase_id) AS ref, h.supplier_name AS party
    FROM purchase_invoice_lines l JOIN purchase_invoices h ON h.purchase_id = l.purchase_id
    WHERE l.receipt_line_id IS NULL
    UNION ALL
    SELECT r.receipt_date, r.created_at, l.item_code, l.quantity, l.unit_cost, l.production_date, l.expiry_date,
           'إذن استلام ' || r.receipt_id, r.supplier_name
    FROM purchase_receipt_lines l JOIN purchase_receipts r ON r.receipt_id = l.receipt_id;
"""


def migrate_schema(conn):
    for table, col, ddl in NEW_COLUMNS:
        cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]
        if cols and col not in cols:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")


def convert_legacy_orders(conn):
    """أوامر النسخة السابقة كانت تُنفَّذ كاملة دفعة واحدة؛ نحوّلها لكميات منفّذة حتى لا يُعاد صرفها."""
    conn.execute("UPDATE sales_order_lines SET delivered_qty = quantity WHERE delivered_qty = 0 AND order_id IN "
                 "(SELECT order_id FROM sales_orders WHERE status IN ('تم التسليم وصرف المخزن','مفوترة بالكامل'))")
    conn.execute("UPDATE sales_order_lines SET invoiced_qty = quantity WHERE invoiced_qty = 0 AND order_id IN "
                 "(SELECT order_id FROM sales_orders WHERE status = 'مفوترة بالكامل')")
    conn.execute("UPDATE sales_orders SET status = 'مسلّم' WHERE status = 'تم التسليم وصرف المخزن'")
    conn.execute("UPDATE sales_orders SET status = 'مفوتر' WHERE status = 'مفوترة بالكامل'")
    conn.execute("UPDATE purchase_order_lines SET received_qty = quantity, invoiced_qty = quantity WHERE order_id IN "
                 "(SELECT order_id FROM purchase_orders WHERE status = 'مستلم بالمخزن ومفوتر')")
    conn.execute("UPDATE purchase_orders SET status = 'مفوتر' WHERE status = 'مستلم بالمخزن ومفوتر'")


def get_setting(key, default=""):
    with db() as conn:
        row = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
    return row[0] if row and row[0] is not None else default


def set_setting(key, value):
    with db() as conn:
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (key, value))


def _backup_before_upgrade():
    """نسخة احتياطية قبل ترقية قاعدة بيانات موجودة إلى الإصدار 3."""
    if not os.path.exists(DB_FILE):
        return
    conn = sqlite3.connect(DB_FILE)
    try:
        version = conn.execute("PRAGMA user_version").fetchone()[0]
        has_tables = conn.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table'").fetchone()[0] > 0
    finally:
        conn.close()
    if has_tables and version < SCHEMA_VERSION:
        backup_database(f"before_v{SCHEMA_VERSION}")


def init_db():
    global MIGRATION_NOTE
    legacy = archive_legacy_db()
    if not legacy:
        _backup_before_upgrade()
    with db() as conn:
        old_version = conn.execute("PRAGMA user_version").fetchone()[0]
        conn.executescript(SCHEMA_SQL)
        conn.executescript(SCHEMA_SQL_V3)
        conn.executescript(POST_DISCOUNT_SQL)
        conn.executescript(ENTERPRISE_SQL)
        migrate_schema(conn)
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_inventory_barcode ON inventory(barcode) "
                     "WHERE barcode IS NOT NULL AND barcode != ''")
        convert_legacy_orders(conn)
        conn.executescript(VIEWS_SQL)
        conn.executescript(USERS_SQL)
        ensure_main_admin(conn)
        conn.executescript(SEQUENCES_SQL)
        init_sequences(conn)
        if old_version < 8:  # ترقيم موحد جديد: ELH-XXXX للأصناف، C-XXXX للعملاء، S-XXXX للموردين (الأكواد القائمة لا تتغير)
            for name, prefix in (("item", "ELH-"), ("customer", "C"), ("supplier", "S")):
                conn.execute("UPDATE sequences SET prefix=?, width=4, next_value=? WHERE name=?",
                             (prefix, max_used_serial(conn, name, prefix) + 1, name))
        if old_version < 10:
            # الأكواد التلقائية تكمل من آخر رقم مسجل: C#### للعملاء، S#### للموردين (بدون شرطة)
            for name, old, new in (("customer", "C-", "C"), ("supplier", "S-", "S")):
                if conn.execute("SELECT prefix FROM sequences WHERE name=?", (name,)).fetchone()[0] in (old, new):
                    conn.execute("UPDATE sequences SET prefix=?, width=4, next_value=? WHERE name=?",
                                 (new, max_used_serial(conn, name, new) + 1, name))
            # صفحة «الشركاء» انقسمت إلى «العملاء» و«الموردين»: من كان لديه الصلاحية القديمة يأخذ الصفحتين
            for uid, perms in conn.execute("SELECT user_id, permissions FROM users").fetchall():
                ps = [x for x in (perms or "").split(",") if x]
                if "entities" in ps:
                    ps = [x for x in ps if x != "entities"] + ["customers", "suppliers"]
                    conn.execute("UPDATE users SET permissions=? WHERE user_id=?", (",".join(dict.fromkeys(ps)), uid))
        if old_version < 11:
            seed_reps_from_history(conn)
            # صفحة «سجل المناديب» الجديدة: من يعدّل العملاء والموردين يمكنه إدارة المناديب
            for uid, perms in conn.execute("SELECT user_id, permissions FROM users").fetchall():
                ps = [x for x in (perms or "").split(",") if x]
                if ("customers" in ps or "suppliers" in ps) and "reps" not in ps:
                    conn.execute("UPDATE users SET permissions=? WHERE user_id=?", (",".join(ps + ["reps"]), uid))
        conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        secure_secret_settings(conn)
        if conn.execute("SELECT COUNT(*) FROM company_profile").fetchone()[0] == 0:
            conn.execute(f"INSERT INTO company_profile (id, {', '.join(COMPANY_FIELDS)}) VALUES (1, {', '.join('?' * len(COMPANY_FIELDS))})",
                         [COMPANY_DEFAULTS[k] for k in COMPANY_FIELDS])
        elif old_version < 9:
            # الجهة المصدرة للفواتير: استكمال البيانات القانونية للمنشأة (الحقول الفارغة فقط، أو استبدال الاسم المؤقت)
            row = conn.execute(f"SELECT {', '.join(COMPANY_FIELDS)} FROM company_profile WHERE id=1").fetchone()
            cur = dict(zip(COMPANY_FIELDS, row))
            if (cur.get("company_name") or "").startswith("اسم المنشأة"):
                cur["company_name"] = ""
            sets = {k: COMPANY_DEFAULTS[k] for k in COMPANY_FIELDS if not (cur.get(k) or "").strip() and k != "rep_name"}
            if sets:
                conn.execute(f"UPDATE company_profile SET {', '.join(f'{k}=?' for k in sets)} WHERE id=1", list(sets.values()))
    if legacy:
        import_from_legacy(legacy)
        MIGRATION_NOTE = (f"تم تحديث هيكل قاعدة البيانات لدعم الفواتير متعددة الأصناف.\n\n"
                          f"النسخة القديمة محفوظة باسم:\n{legacy}\n\n"
                          f"تم نقل الشركاء والأصناف (بأرصدتها الحالية كرصيد افتتاحي) وبيانات الشركة.\n"
                          f"الفواتير والحركات القديمة لم تُنقل.")
    # الدليل يُستورد من الإكسيل فقط عند أول تشغيل أو عند الترقية (لتصحيح التصنيفات)،
    # ولا يُمسح عند كل تشغيل حتى لا تضيع الحسابات المضافة يدوياً (مثل حسابات العاملين الفرعية).
    with db() as conn:
        coa_empty = conn.execute("SELECT COUNT(*) FROM chart_of_accounts").fetchone()[0] == 0
    if coa_empty or old_version < 4:
        import_chart_of_accounts()
    with db() as conn:
        normalize_account_types(conn)
    # الوحدات الإضافية (الدفتر العام، سجل التدقيق، ...) تنشئ جداولها ومشغلاتها هنا
    with db() as conn:
        for hook in INIT_HOOKS:
            hook(conn)
    with db() as conn:   # executescript داخل التهيئة يحفظ ما قبله؛ نضمن تنفيذ الترحيل المعلق الآن
        for hook in COMMIT_HOOKS:
            hook(conn)


# ============================================================
# تسجيل الدخول وكلمات المرور (PBKDF2-SHA256 مع salt عشوائي، لا تُخزَّن كلمة المرور نفسها)
# ============================================================
def hash_password(password, salt=None, iterations=PBKDF2_ITERATIONS):
    salt = salt or os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt.hex()}${dk.hex()}"


def verify_password(password, stored):
    try:
        algo, iterations, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), int(iterations))
        return hmac.compare_digest(dk.hex(), hash_hex)
    except (ValueError, AttributeError):
        return False


def main_admin_user():
    """اسم المدير الرئيسي (محمي من الحذف وتغيير الصلاحيات) كما هو محفوظ في الإعدادات."""
    try:
        return get_setting(MAIN_ADMIN_KEY, "") or ""
    except sqlite3.Error:
        return ""


def _new_random_password():
    import secrets
    return secrets.token_urlsafe(12) + "7a"   # حروف وأرقام دائماً (شرط validate_new_password)


def ensure_main_admin(conn):
    """يحدد المدير الرئيسي بدون أي بيانات مكتوبة في الكود ولا يغيّر كلمة مرور موجودة:
    - قاعدة قائمة: المدير المحفوظ في الإعدادات، وإلا ELHAGRY_ADMIN_USER لو موجود، وإلا أقدم مدير.
    - قاعدة جديدة بلا مدير: يُنشأ ELHAGRY_ADMIN_USER (أو admin) بكلمة ELHAGRY_ADMIN_PASSWORD، أو بكلمة عشوائية
      تُكتب مرة واحدة في initial_admin_password.txt (في .gitignore) ويجب تغييرها بعد أول دخول.
    - استعادة الدخول: ELHAGRY_ADMIN_RESET=1 مع ELHAGRY_ADMIN_PASSWORD يضبط كلمة مرور المدير الرئيسي (مشفرة)."""
    def exists(u):
        return bool(u) and conn.execute("SELECT 1 FROM users WHERE username=?", (u,)).fetchone() is not None
    row = conn.execute("SELECT value FROM app_settings WHERE key=?", (MAIN_ADMIN_KEY,)).fetchone()
    main = row[0] if row and exists(row[0]) else None
    env_user = (os.environ.get("ELHAGRY_ADMIN_USER") or "").strip()
    env_pw = os.environ.get("ELHAGRY_ADMIN_PASSWORD") or ""
    if not main:
        if exists(env_user):
            main = env_user
        else:
            r = conn.execute("SELECT username FROM users WHERE is_admin=1 ORDER BY user_id LIMIT 1").fetchone()
            main = r[0] if r else None
    if not main:
        main = env_user or "admin"
        validate_username(main)
        pw = env_pw or _new_random_password()
        if env_pw:
            validate_new_password(env_pw)
        conn.execute("INSERT INTO users (username, full_name, password_hash, is_admin, permissions) VALUES (?,?,?,1,?)",
                     (main, "مدير النظام", hash_password(pw), ",".join(ALL_PERMISSION_KEYS)))
        if not env_pw:
            path = os.path.join(DATA_DIR, INITIAL_PASSWORD_FILE)
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(f"المدير الرئيسي: {main}\nكلمة المرور المؤقتة: {pw}\n"
                         "غيّرها بعد أول دخول ثم احذف هذا الملف. (الملف في .gitignore)\n")
    conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (MAIN_ADMIN_KEY, main))
    conn.execute("UPDATE users SET is_admin=1, active=1, permissions=? WHERE username=?", (",".join(ALL_PERMISSION_KEYS), main))
    if env_pw and os.environ.get("ELHAGRY_ADMIN_RESET") == "1":
        validate_new_password(env_pw)
        conn.execute("UPDATE users SET password_hash=?, failed_attempts=0, locked_until=NULL WHERE username=?",
                     (hash_password(env_pw), main))
    return main


# كلمات المرور الإدارية في الإعدادات (لا قيم افتراضية): تُخزَّن مشفرة PBKDF2، وتُقرأ أول مرة من البيئة إن وُجدت
SECRET_SETTINGS = {"admin_password": ("ELHAGRY_MODIFY_PASSWORD", "كلمة مرور المدير (فتح الأيام المغلقة وتجاوز الحد الائتماني)"),
                   "reset_password": ("ELHAGRY_RESET_PASSWORD", "كلمة مرور تصفير الحركات")}
SECRET_MIN_LEN = 4


def _is_hashed(value):
    return str(value or "").startswith("pbkdf2_sha256$")


def secure_secret_settings(conn):
    """ترحيل آمن: أي كلمة مرور إدارية محفوظة كنص صريح تُستبدل ببصمتها المشفرة (نفس كلمة المرور تبقى صالحة،
    لا تُمسح أي بيانات). غير المضبوطة تُؤخذ من متغير البيئة إن وُجد، وإلا تبقى غير مضبوطة حتى يضبطها المدير."""
    for key, (env, _) in SECRET_SETTINGS.items():
        r = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        val = (r[0] if r else "") or ""
        if val and not _is_hashed(val):
            conn.execute("UPDATE app_settings SET value=? WHERE key=?", (hash_password(val), key))
        elif not val and os.environ.get(env):
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)",
                         (key, hash_password(os.environ[env])))


def secret_is_set(key):
    return _is_hashed(get_setting(key, ""))


def set_secret_setting(key, value):
    value = (value or "").strip()
    if len(value) < SECRET_MIN_LEN:
        raise UserError(f"{SECRET_SETTINGS[key][1]} يجب ألا تقل عن {SECRET_MIN_LEN} أحرف/أرقام.")
    set_setting(key, hash_password(value))


def verify_secret_setting(key, value):
    """True لو كلمة المرور صحيحة. يرفع UserError لو لم تُضبط بعد (لا توجد قيمة افتراضية)."""
    stored = get_setting(key, "")
    if not _is_hashed(stored):
        raise UserError(f"لم تُضبط {SECRET_SETTINGS[key][1]} بعد.\n"
                        "يضبطها مدير النظام من «بيانات الشركة والأمان» أو من متغير البيئة "
                        f"{SECRET_SETTINGS[key][0]} في ملف .env.")
    return verify_password((value or "").strip(), stored)


def verify_modify_password(value):
    return verify_secret_setting("admin_password", value)


def set_modify_password(value):
    set_secret_setting("admin_password", value)


def validate_new_password(password):
    if len(password) < 8:
        raise UserError("كلمة المرور يجب أن تكون 8 أحرف على الأقل.")
    if not (re.search(r"[A-Za-z]", password) and re.search(r"\d", password)):
        raise UserError("كلمة المرور يجب أن تحتوي على حروف إنجليزية وأرقام.")


def validate_username(username):
    if not re.fullmatch(r"[A-Za-z0-9.\-]{3,30}", username or ""):
        raise UserError("اسم المستخدم من 3 إلى 30 حرفاً: حروف إنجليزية وأرقام ونقطة وشرطة فقط (بدون مسافات).")


def user_record(row):
    """صف (user_id, username, full_name, is_admin, permissions) → قاموس جلسة المستخدم."""
    uid, username, full_name, is_admin, perms = row
    perm_set = set(ALL_PERMISSION_KEYS) if is_admin else {p for p in (perms or "").split(",") if p in ALL_PERMISSION_KEYS}
    return dict(user_id=uid, username=username, full_name=full_name or username, is_admin=bool(is_admin), perms=perm_set)


def authenticate(username, password):
    """يرجع قاموس المستخدم عند النجاح، أو يرفع UserError برسالة عامة (لا تكشف أي الحقلين خاطئ)."""
    username = (username or "").strip()
    generic = "اسم المستخدم أو كلمة المرور غير صحيحة."
    now = datetime.now()
    stamp = now.strftime("%Y-%m-%d %H:%M:%S")
    with db() as conn:
        row = conn.execute("SELECT user_id, password_hash, active, failed_attempts, locked_until FROM users "
                           "WHERE username=?", (username,)).fetchone()
        if row is None:
            raise UserError(generic)
        uid, stored, active, failed, locked_until = row
        if not active:
            raise UserError("هذا المستخدم موقوف. راجع مدير النظام.")
        if locked_until and locked_until > stamp:
            raise UserError(f"تم إيقاف الدخول مؤقتاً بسبب محاولات خاطئة متكررة.\nحاول مرة أخرى بعد الساعة {locked_until[11:16]}.")
        failure = None
        if not verify_password(password or "", stored):
            failed = (failed or 0) + 1
            lock = None
            if failed >= MAX_FAILED_LOGINS:
                lock = (now + timedelta(minutes=LOCKOUT_MINUTES)).strftime("%Y-%m-%d %H:%M:%S")
                failed = 0
            conn.execute("UPDATE users SET failed_attempts=?, locked_until=? WHERE user_id=?", (failed, lock, uid))
            failure = (f"محاولات خاطئة متكررة. تم إيقاف الدخول لمدة {LOCKOUT_MINUTES} دقائق." if lock else
                       f"{generic}\n(المحاولات المتبقية قبل الإيقاف المؤقت: {MAX_FAILED_LOGINS - failed})")
        else:
            conn.execute("UPDATE users SET failed_attempts=0, locked_until=NULL, last_login=? WHERE user_id=?", (stamp, uid))
            rec = conn.execute("SELECT user_id, username, full_name, is_admin, permissions FROM users WHERE user_id=?",
                               (uid,)).fetchone()
    # الخطأ يُرفع بعد حفظ عداد المحاولات (رفعه داخل الاتصال يلغي الحفظ)
    if failure:
        raise UserError(failure)
    return user_record(rec)


def active_admin_count(conn, exclude_id=None):
    return conn.execute("SELECT COUNT(*) FROM users WHERE is_admin=1 AND active=1 AND user_id != ?",
                        (exclude_id or -1,)).fetchone()[0]


_KEEP = object()


def normalize_permissions(perms):
    """تنظيف قائمة الصلاحيات (مشتركة بين المستخدمين والأدوار): مفاتيح معروفة فقط، الإنشاء/التنفيذ يستلزم العرض،
    والصلاحيات الفرعية تتبع صفحتها الأم، وكل صفحة لها مديول تفصيلي تحتاج صلاحية فرعية واحدة على الأقل."""
    perms = [k for k in perms if k in ALL_PERMISSION_KEYS]
    perms = perms + [k[:-7] for k in perms if k.endswith("_create") and k[:-7] in SUB_PARENT and k[:-7] not in perms]
    pages = [k for k in PERMISSION_KEYS if k in perms]
    subs_all = [k for k in perms if k in SUB_PARENT and SUB_PARENT[k] in pages]
    for parent in (p_ for p_ in pages if p_ in SUB_PERMISSIONS):
        if not any(s_ in subs_all for s_, _ in SUB_PERMISSIONS[parent]):
            raise UserError(f"حددت صفحة «{dict(PERMISSIONS)[parent]}» بدون أي صلاحية تفصيلية داخلها.\n"
                            "اختر صلاحية فرعية واحدة على الأقل.")
    return pages + subs_all


def save_user_account(uid, uname, full, pw, pw2, is_admin, active, perms, current_user_id=None, role_id=_KEEP):
    """إضافة/تعديل مستخدم (مشتركة بين واجهة سطح المكتب والويب). perms = مفاتيح الصفحات + المفاتيح التفصيلية المختارة.
    يحمي حساب المدير الرئيسي (اسمه وصلاحياته ثابتة) ويضمن بقاء مدير نشط واحد على الأقل."""
    uname = (uname or "").strip()
    validate_username(uname)
    full = (full or "").strip()
    perms = list(ALL_PERMISSION_KEYS) if is_admin else normalize_permissions(perms)
    if uid is None and not pw:
        raise UserError("أدخل كلمة مرور للمستخدم الجديد.")
    if pw or pw2:
        if pw != pw2:
            raise UserError("كلمتا المرور غير متطابقتين.")
        validate_new_password(pw)
    if not perms:
        raise UserError("اختر صلاحية واحدة على الأقل.")
    is_admin, active = int(bool(is_admin)), int(bool(active))
    with db() as conn:
        if conn.execute("SELECT 1 FROM users WHERE lower(username)=lower(?) AND user_id != ?",
                        (uname, uid if uid is not None else -1)).fetchone():
            raise UserError("اسم المستخدم مستخدم بالفعل.")
        if uid is not None:
            old = conn.execute("SELECT is_admin, active, username FROM users WHERE user_id=?", (uid,)).fetchone()
            if old is None:
                raise UserError("المستخدم غير موجود.")
            main = main_admin_user()
            if old[2] == main and (uname != main or not is_admin or not active):
                raise UserError(f"حساب المدير الرئيسي «{main}» لا يمكن تغيير اسمه أو صلاحياته أو إيقافه.")
            if uid == current_user_id and (not is_admin or not active):
                raise UserError("لا يمكنك إلغاء صلاحية المدير أو إيقاف حسابك الحالي.")
            if old[0] and old[1] and (not is_admin or not active) and active_admin_count(conn, uid) == 0:
                raise UserError("يجب أن يبقى مدير نشط واحد على الأقل.")
            conn.execute("UPDATE users SET username=?, full_name=?, is_admin=?, permissions=?, active=? WHERE user_id=?",
                         (uname, full, is_admin, ",".join(perms), active, uid))
            if pw:
                conn.execute("UPDATE users SET password_hash=?, failed_attempts=0, locked_until=NULL WHERE user_id=?",
                             (hash_password(pw), uid))
            if role_id is not _KEEP:
                conn.execute("UPDATE users SET role_id=? WHERE user_id=?", (role_id, uid))
            return f"تم حفظ تعديلات المستخدم {uname}." + ("\nتم تغيير كلمة المرور." if pw else "")
        conn.execute("INSERT INTO users (username, full_name, password_hash, is_admin, permissions, active, role_id) "
                     "VALUES (?,?,?,?,?,?,?)", (uname, full, hash_password(pw), is_admin, ",".join(perms), active,
                                                None if role_id is _KEEP else role_id))
        return f"تم إضافة المستخدم {uname}."


def delete_user_account(uid, current_user_id=None):
    if uid == current_user_id:
        raise UserError("لا يمكنك حذف حسابك الحالي.")
    with db() as conn:
        r = conn.execute("SELECT username, is_admin, active FROM users WHERE user_id=?", (uid,)).fetchone()
        if r is None:
            raise UserError("المستخدم غير موجود.")
        if r[0] == main_admin_user():
            raise UserError(f"لا يمكن حذف حساب المدير الرئيسي «{r[0]}».")
        if r[1] and r[2] and active_admin_count(conn, uid) == 0:
            raise UserError("لا يمكن حذف آخر مدير نشط.")
        conn.execute("DELETE FROM users WHERE user_id=?", (uid,))
    return r[0]


def list_users():
    with db() as conn:
        return conn.execute("SELECT user_id, username, COALESCE(full_name,''), is_admin, active, permissions, "
                            "COALESCE(last_login,''), failed_attempts, COALESCE(locked_until,'') FROM users "
                            "ORDER BY is_admin DESC, username").fetchall()


# ============================================================
# النسخ الاحتياطي
# ============================================================
def backup_database(reason="manual"):
    if not os.path.exists(DB_FILE):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    path = os.path.join(BACKUP_DIR, f"accounting_{reason}_{datetime.now():%Y%m%d_%H%M%S}.db")
    src = sqlite3.connect(DB_FILE)
    dst = sqlite3.connect(path)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    return path


# ============================================================
# تصفير الحركات (للمدير فقط وبكلمة مرور التصفير) — منطق موحد لسطح المكتب والويب
# ============================================================
RESET_PASSWORD_KEY = "reset_password"
# الترتيب مهم: الأبناء قبل الآباء حتى لا يرفض SQLite الحذف بسبب المفاتيح الأجنبية
RESET_TABLES = ("advance_allocations",
                "sales_delivery_lines", "purchase_receipt_lines", "sales_invoice_lines", "purchase_invoice_lines",
                "sales_deliveries", "purchase_receipts", "sales_invoices", "purchase_invoices",
                "sales_order_lines", "purchase_order_lines", "sales_orders", "purchase_orders",
                "returns_invoices", "cash_transactions", "daily_closures", "post_discount_lines",
                "post_discounts", "journal_lines", "journal_entries",
                # سجل استيراد HansaWorld (يُعاد الاستيراد بعد التصفير بدون أن يُعتبر مكرراً)
                "hw_import_map", "hw_item_history")
RESET_SEQUENCES = ("sales_invoice", "purchase_invoice", "purchase_discount", "sales_discount")


def reset_password_is_set():
    return secret_is_set(RESET_PASSWORD_KEY)


def set_reset_password(new_password):
    set_secret_setting(RESET_PASSWORD_KEY, new_password)


def reset_transactions(password, clear_audit=False):
    """حذف كل المستندات والحركات وإرجاع المخزون للرصيد الافتتاحي (دليل الحسابات والشركاء والأصناف والأصول والمستخدمين تبقى).
    يتطلب: مستخدماً مديراً (أو تشغيلاً من النظام بلا جلسة)، وكلمة مرور التصفير (الافتراضية 2026)، ونسخة احتياطية ناجحة أولاً.
    clear_audit=True: يمسح سجل التدقيق والرقابة أيضاً (بعد النسخة الاحتياطية التي تحفظه كاملاً) ويبدأ سلسلة بصمات جديدة
    بسطر افتتاحي واحد يوثّق المسح (المستخدم، التوقيت، عدد السطور الممسوحة، اسم النسخة الاحتياطية).
    يرجع dict(backup=مسار النسخة، deleted={الجدول: عدد السطور المحذوفة}، audit_cleared=عدد سطور السجل الممسوحة أو None)."""
    u = _CURRENT_PERMS.get()
    if u is not None and not u.get("is_admin"):
        raise UserError("تصفير الحركات متاح لمدير النظام فقط.")
    if not verify_secret_setting(RESET_PASSWORD_KEY, password):
        raise UserError("كلمة مرور التصفير غير صحيحة. لم يتم حذف أي بيانات.\n"
                        "(يمكن للمدير تغييرها من الإعدادات ← كلمة مرور تصفير الحركات)")
    try:
        backup = backup_database("before_reset")
    except Exception as e:  # noqa: BLE001
        raise UserError(f"تعذر أخذ نسخة احتياطية قبل التصفير، فتم إلغاء العملية حفاظاً على البيانات.\nالسبب: {e}")
    if not backup:
        raise UserError(f"ملف قاعدة البيانات غير موجود:\n{DB_FILE}")
    from . import audit as _audit   # استيراد متأخر: audit يعتمد على engine
    deleted, audit_cleared = {}, None
    try:
        with db() as conn:
            existing = {n for (n,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if not clear_audit:
                _audit.audit_event(conn, _audit.ACT_EVENT, "تصفير الحركات", "كل المستندات",
                                   new={"backup": os.path.basename(backup)})
            for table in RESET_TABLES:
                if table in existing:
                    deleted[table] = conn.execute(f"DELETE FROM {table}").rowcount
            # مرفقات المستندات المحذوفة: بدون حذفها ستظهر على المستندات الجديدة التي تأخذ نفس الأرقام
            rels = [p for (p,) in conn.execute(f"SELECT file_path FROM attachments WHERE doc_type IN "
                                               f"({','.join('?' * len(ATTACH_DOCS))})", list(ATTACH_DOCS))]
            conn.execute(f"DELETE FROM attachments WHERE doc_type IN ({','.join('?' * len(ATTACH_DOCS))})",
                         list(ATTACH_DOCS))
            _archive_attachment_files(rels, f"reset_{datetime.now():%Y%m%d_%H%M%S}")
            conn.execute("UPDATE inventory SET quantity = opening_qty, cost_price = opening_cost")
            conn.execute(f"UPDATE sequences SET next_value=1 WHERE name IN ({','.join('?' * len(RESET_SEQUENCES))})",
                         RESET_SEQUENCES)
            conn.execute("DELETE FROM app_settings WHERE key=?", (CASH_OPENING_SETTING,))
            if "sqlite_sequence" in existing:
                names = [t for t in RESET_TABLES if t in existing]
                conn.execute(f"DELETE FROM sqlite_sequence WHERE name IN ({','.join('?' * len(names))})", names)
            left = {t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in RESET_TABLES if t in existing}
            left = {t: n for t, n in left.items() if n}
            if left:
                raise UserError("لم يكتمل التصفير (بقيت سطور في: " + "، ".join(f"{t}={n}" for t, n in left.items()) + ").")
            if clear_audit and "audit_log" in existing:
                # آخر خطوة داخل نفس المعاملة: مشغلات حذف المستندات أعلاه تكتب في السجل، فيُمسح بعدها
                audit_cleared = _audit.clear_audit_log(conn, os.path.basename(backup), deleted)
    except sqlite3.OperationalError as e:
        if "locked" in str(e).lower() or "busy" in str(e).lower():
            raise UserError("قاعدة البيانات مشغولة الآن بعملية أخرى (نافذة ثانية من البرنامج أو واجهة الويب تكتب عليها).\n"
                            "أغلق الواجهات الأخرى أو انتظر لحظات ثم أعد المحاولة. لم يتم حذف أي بيانات.")
        raise UserError(f"تعذر التصفير بسبب خطأ في قاعدة البيانات، ولم يُحذف شيء (تم التراجع الكامل).\nالتفاصيل: {e}")
    except sqlite3.IntegrityError as e:
        raise UserError(f"تعذر التصفير بسبب ارتباطات بين الجداول، ولم يُحذف شيء (تم التراجع الكامل).\nالتفاصيل: {e}")
    return dict(backup=backup, deleted=deleted, audit_cleared=audit_cleared)


def auto_backup_daily():
    """نسخة احتياطية تلقائية مرة يومياً مع الاحتفاظ بآخر 15 نسخة."""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        prefix = f"accounting_auto_{datetime.now():%Y%m%d}"
        if not any(f.startswith(prefix) for f in os.listdir(BACKUP_DIR)):
            backup_database("auto")
        autos = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("accounting_auto_"))
        for old in autos[:-15]:
            os.remove(os.path.join(BACKUP_DIR, old))
    except Exception as e:
        print(f"تعذر النسخ الاحتياطي التلقائي: {e}")


# ============================================================
# منطق التقارير (منفصل عن الواجهة)
# ============================================================
def months_elapsed(purchase_date, as_of, life_years):
    """الإهلاك يبدأ من الشهر التالي لشهر الشراء."""
    p = datetime.strptime(purchase_date, "%Y-%m-%d").date()
    months = (as_of.year - p.year) * 12 + (as_of.month - p.month)
    return max(0, min(months, life_years * 12))


def accumulated_depreciation(cost, salvage, life_years, purchase_date, as_of):
    monthly = (cost - (salvage or 0.0)) / (life_years * 12)
    return monthly * months_elapsed(purchase_date, as_of, life_years)


# ============================================================
# الخزينة والقيود: أرصدة موحدة من السندات + القيود اليومية
# ============================================================
def cash_accounts(conn):
    """{اسم حساب الخزينة: كوده}. من دليل الحسابات (فرع 111) + الحسابات الافتراضية."""
    out = {}
    for name, code in conn.execute("SELECT account_name, account_code FROM chart_of_accounts WHERE account_code LIKE ? "
                                   "AND COALESCE(is_group,0)=0 ORDER BY account_code", (CASH_PREFIX + "%",)):
        out.setdefault(name, code)
    for name, code in DEFAULT_CASH_CODES.items():
        out.setdefault(name, code)
    return out


def cash_account_names(conn):
    """أسماء حسابات الخزينة المتاحة للسندات: الحسابات الأساسية أولاً ثم بقية فرع النقدية بالدليل."""
    names = list(BANK_ACCOUNTS)
    names += [n for n in cash_accounts(conn) if n not in names]
    return names


def is_cash_box(name, code):
    return (code or "").startswith(CASH_BOX_PREFIX) or "صندوق" in (name or "")


def cash_movements(conn, d_from=None, d_to=None):
    """{حساب الخزينة: [رصيد قبل d_from، المقبوضات، المدفوعات]} خلال الفترة.
    يشمل سندات القبض/الصرف + أسطر القيود اليومية على حسابات النقدية والبنوك (المدين وارد، الدائن منصرف)."""
    d_from = d_from or "0000-01-01"
    d_to = d_to or "9999-12-31"
    res = defaultdict(lambda: [0.0, 0.0, 0.0])
    for acc, before, inflow, outflow in conn.execute(
            "SELECT account_bank, "
            "SUM(CASE WHEN trans_date < ? THEN (CASE WHEN trans_type='قبض' THEN amount ELSE -amount END) ELSE 0 END), "
            "SUM(CASE WHEN trans_date >= ? AND trans_type='قبض' THEN amount ELSE 0 END), "
            "SUM(CASE WHEN trans_date >= ? AND trans_type='صرف' THEN amount ELSE 0 END) "
            "FROM cash_transactions WHERE trans_date <= ? GROUP BY account_bank", (d_from, d_from, d_from, d_to)):
        r = res[acc]
        r[0] += before or 0.0
        r[1] += inflow or 0.0
        r[2] += outflow or 0.0
    accounts = cash_accounts(conn)
    by_code = {code: name for name, code in accounts.items()}
    if by_code:
        ph = ",".join("?" * len(by_code))
        for code, before, dr, cr in conn.execute(
                f"SELECT l.account_code, SUM(CASE WHEN e.entry_date < ? THEN l.debit - l.credit ELSE 0 END), "
                f"SUM(CASE WHEN e.entry_date >= ? THEN l.debit ELSE 0 END), "
                f"SUM(CASE WHEN e.entry_date >= ? THEN l.credit ELSE 0 END) "
                f"FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id "
                f"WHERE e.entry_date <= ? AND l.account_code IN ({ph}) GROUP BY l.account_code",
                [d_from, d_from, d_from, d_to] + list(by_code)):
            r = res[by_code[code]]
            r[0] += before or 0.0
            r[1] += dr or 0.0
            r[2] += cr or 0.0
    return res


def cash_balances(conn, as_of=None):
    """{حساب الخزينة: الرصيد حتى تاريخ} (سندات + قيود)."""
    return {acc: o + i - x for acc, (o, i, x) in cash_movements(conn, None, as_of).items()}


def journal_cash_by_month(conn, k1, k2):
    """صافي حركة القيود اليومية على حسابات الخزينة لكل شهر (مدين - دائن)."""
    codes = list(cash_accounts(conn).values())
    if not codes:
        return {}
    ph = ",".join("?" * len(codes))
    return {k: v or 0.0 for k, v in conn.execute(
        f"SELECT substr(e.entry_date,1,7), SUM(l.debit - l.credit) FROM journal_lines l "
        f"JOIN journal_entries e ON e.entry_id=l.entry_id WHERE l.account_code IN ({ph}) "
        f"AND substr(e.entry_date,1,7) BETWEEN ? AND ? GROUP BY 1", codes + [k1, k2])}


def journal_pl(conn, d_from, d_to, by_month=False):
    """أثر القيود اليومية على قائمة الدخل: [(الشهر أو ''، النوع، الكود، الاسم، المبلغ)].
    المبلغ = مدين - دائن للمصروفات، ودائن - مدين للإيرادات. حسابات تكلفة المبيعات ومردودات المشتريات
    والإهلاك مستبعدة لأن النظام يحسبها تلقائياً من المخزون والأصول."""
    key = "substr(e.entry_date,1,7)" if by_month else "''"
    out = []
    for k, typ, code, name, d, c in conn.execute(
            f"SELECT {key}, a.account_type, l.account_code, MAX(a.account_name), SUM(l.debit), SUM(l.credit) "
            f"FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id "
            f"JOIN chart_of_accounts a ON a.account_code=l.account_code "
            f"WHERE e.entry_date BETWEEN ? AND ? AND a.account_type IN (?,?) "
            f"AND (e.source_type IS NULL OR e.source_type='hw_tb_expenses') GROUP BY 1, 2, 3",
            (d_from, d_to, TYPE_EXP, TYPE_REV)):
        if code.startswith(SYSTEM_PL_PREFIXES):
            continue
        amt = ((d or 0.0) - (c or 0.0)) if typ == TYPE_EXP else ((c or 0.0) - (d or 0.0))
        if abs(amt) > EPS:
            out.append((k, typ, code, name, amt))
    return out


def journal_pl_totals(conn, d_from, d_to):
    """(مصروفات القيود، إيرادات القيود) للفترة."""
    rows = journal_pl(conn, d_from, d_to)
    return (sum(r[4] for r in rows if r[1] == TYPE_EXP), sum(r[4] for r in rows if r[1] == TYPE_REV))


def expense_breakdown(conn, d_from, d_to):
    """المصروفات بحسب البند: سندات الصرف المصنفة مصروفات + القيود اليومية على حسابات المصروفات.
    يرجع [(البند، عدد الحركات، المبلغ)] مرتبة تنازلياً."""
    agg = defaultdict(lambda: [0, 0.0])
    for name, n, amt in conn.execute(
            "SELECT party_name, COUNT(*), SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND category=? "
            "AND trans_date BETWEEN ? AND ? GROUP BY party_name", (CAT_EXPENSE, d_from, d_to)):
        agg[name][0] += n
        agg[name][1] += amt or 0.0
    for _, typ, _, name, amt in journal_pl(conn, d_from, d_to):
        if typ == TYPE_EXP:
            agg[name][0] += 1
            agg[name][1] += amt
    return sorted(((k, v[0], v[1]) for k, v in agg.items()), key=lambda r: r[2], reverse=True)


HW_MOVE_LABELS = {"Deld": "تسليم للعميل", "G.Rec.": "استلام من مورد", "Cust. Ret.": "مرتجع من عميل",
                  "Ret. to Supp.": "مرتجع لمورد", "St. Depr.": "إهلاك / تالف مخزني"}
HW_SUPPLIER_MOVES = ("G.Rec.", "Ret. to Supp.")
ITEM_CARD_SOURCES = ("HansaWorld", "النظام")


def item_card(item_code, d_from, d_to):
    """كارت الصنف التفصيلي: كل حركة على حدة (حركات HansaWorld التاريخية + مستندات النظام) برصيد متحرك.
    حركات HansaWorld (hw_item_history) تبدأ من رصيد صفر وتنتهي برصيد ملف التكلفة الذي صار الرصيد الافتتاحي للصنف،
    فلا تُكرَّر الكمية: الرصيد الأساسي قبلها = الرصيد الافتتاحي − مجموع كمياتها (صفر عادةً).
    يرجع dict(opening, final, moves=[dict(...)], cost, has_hw, opening_label). لكل حركة: date, src, type, ref, link,
    party, party_code, qty_in, qty_out, qty, balance, cost, value (بالتكلفة)، doc_price, doc_value, notes,
    key (لفتح المستند أو كشف حساب الطرف)."""
    with db() as conn:
        row = conn.execute("SELECT COALESCE(opening_qty,0), COALESCE(opening_cost,0), COALESCE(cost_price,0) "
                           "FROM inventory WHERE item_code=?", (item_code,)).fetchone()
        if row is None:
            raise UserError("الصنف غير موجود.")
        opening_qty, opening_cost, avg_cost = row
        names = dict(conn.execute("SELECT partner_code, name FROM business_partners"))

        def one(sql, params):
            return conn.execute(sql, params).fetchone()[0] or 0.0

        try:
            hw_rows = conn.execute(
                "SELECT move_date, src_row, move_type, COALESCE(reference,''), COALESCE(party_text,''), "
                "COALESCE(party_code,''), qty, COALESCE(link_doc_type,''), COALESCE(link_doc_no,''), is_duplicate, "
                "COALESCE(location,'') FROM hw_item_history WHERE item_code=? ORDER BY move_date, src_row",
                (item_code,)).fetchall()
        except sqlite3.OperationalError:
            hw_rows = []
        hw_docs = {}
        if hw_rows:   # المستند المربوط في HansaWorld ← مستنده المستورد في النظام (لفتحه من الكارت)
            kinds = {"sales_invoices": "sales", "purchase_invoices": "purchase", "post_discounts": "discount"}
            for role, typ, no, target, tid in conn.execute(
                    "SELECT role, doc_type, doc_no, target, target_id FROM hw_import_map WHERE target IN (?,?,?)",
                    tuple(kinds)):
                hw_docs.setdefault((role, typ, str(no)), (kinds[target], tid))
        hw_total = sum(r[6] or 0.0 for r in hw_rows)
        base = round(opening_qty - hw_total, 6) if hw_rows else opening_qty
        sys_prior = (one("SELECT SUM(qty) FROM v_stock_in WHERE item_code=? AND mdate < ?", (item_code, d_from))
                     - one("SELECT SUM(qty) FROM v_stock_out WHERE item_code=? AND mdate < ?", (item_code, d_from))
                     + one("SELECT SUM(quantity) FROM returns_invoices WHERE item_code=? AND return_type=? "
                           "AND return_date < ?", (item_code, SALES_RETURN, d_from))
                     - one("SELECT SUM(quantity) FROM returns_invoices WHERE item_code=? AND return_type=? "
                           "AND return_date < ?", (item_code, PURCHASE_RETURN, d_from)))
        running = base + sum(r[6] or 0.0 for r in hw_rows if r[0] < d_from) + sys_prior

        moves = []

        def add(d, order, src, typ, ref, link, pcode, ptext, qty, cost, doc_price, doc_value, notes, key):
            moves.append(dict(date=d, _order=order, src=src, type=typ, ref=str(ref) if ref else "-", link=link or "",
                              party_code=pcode or "", party=(names.get(pcode) if pcode else None) or ptext or "-",
                              qty=qty, cost=cost, doc_price=doc_price, doc_value=doc_value, notes=notes or "",
                              key=key))

        for d, srow, typ, ref, ptext, pcode, qty, ltype, lno, dup, loc in hw_rows:
            if not (d_from <= d <= d_to):
                continue
            notes = " • ".join(x for x in (f"موقع {loc}" if loc else "", "مكرر في HansaWorld" if dup else "") if x)
            add(d, (0, srow or 0), "HansaWorld", HW_MOVE_LABELS.get(typ, typ), ref,
                f"{ltype} {lno}" if lno else "", pcode, ptext, qty or 0.0, opening_cost, None, None, notes,
                hw_docs.get(("supplier" if typ in HW_SUPPLIER_MOVES else "customer", ltype, lno))
                or (("partner", pcode) if pcode else None))
        for view, sign in (("v_stock_in", 1), ("v_stock_out", -1)):
            for r in conn.execute(f"SELECT mdate, created_at, mtype, party, qty, unit_price, value, notes, ref, party_code, "
                                  f"unit_cost, doc_kind, doc_id FROM {view} WHERE item_code=? AND mdate BETWEEN ? AND ?",
                                  (item_code, d_from, d_to)):
                add(r[0], (1, r[1] or ""), "النظام", r[2], r[8], "", r[9], r[3], sign * (r[4] or 0.0), r[10] or 0.0,
                    r[5], r[6], r[7], (r[11], r[12]))
        for r in conn.execute(
                "SELECT return_date, created_at, return_type, party_name, quantity, unit_price, subtotal, notes, return_id, "
                "party_code, unit_cost FROM returns_invoices WHERE item_code=? AND return_date BETWEEN ? AND ?",
                (item_code, d_from, d_to)):
            is_add = r[2] == SALES_RETURN
            add(r[0], (1, r[1] or ""), "النظام", f"{r[2]} ({'إضافة' if is_add else 'خصم'})", f"مرتجع {r[8]}", "", r[9],
                r[3], (r[4] or 0.0) * (1 if is_add else -1), r[10] or 0.0, r[5], r[6], r[7], ("return", r[8]))
        for r in conn.execute(
                "SELECT d.discount_date, d.created_at, d.discount_no, d.party_name, l.amount, l.applied, "
                "COALESCE(d.reason,''), d.party_code, d.discount_id FROM post_discount_lines l "
                "JOIN post_discounts d ON d.discount_id=l.discount_id WHERE l.item_code=? AND d.discount_date BETWEEN ? AND ?",
                (item_code, d_from, d_to)):
            note = (r[6] or "") + (f" | يخص مبيعات سابقة: {fmt(r[4] - r[5])}" if r[4] - r[5] > 0.005 else "")
            add(r[0], (1, r[1] or ""), "النظام", "خصم لاحق (تخفيض تكلفة بدون كمية)", r[2], "", r[7], r[3], 0.0, 0.0,
                None, -(r[5] or 0.0), note, ("discount", r[8]))

    moves.sort(key=lambda m: (m["date"], m["_order"]))
    opening = running
    for m in moves:
        running += m["qty"]
        m.pop("_order")
        m["qty_in"] = m["qty"] if m["qty"] > 0 else 0.0
        m["qty_out"] = -m["qty"] if m["qty"] < 0 else 0.0
        m["balance"] = running
        m["value"] = (round(abs(m["qty"]) * m["cost"], 2) if m["qty"] else
                      (m["doc_value"] if m["doc_value"] is not None else 0.0))
    first = item_first_move_date(item_code)
    return dict(opening=opening, final=running, moves=moves, cost=avg_cost, has_hw=bool(hw_rows),
                opening_label="رصيد مرحّل" if first and first < d_from else "رصيد أول المدة")


def item_card_rows(card, moves=None, d_from=""):
    """صفوف العرض/الطباعة المختصرة (8 أعمدة) من item_card — تُستخدم في الويب وPDF."""
    rows = [(d_from, card["opening_label"], "-", "-", f"{card['opening']:,.3f}", "-", "-", "-")]
    for m in (card["moves"] if moves is None else moves):
        q = m["qty"]
        qty_str = f"+{q:g}" if q > 0 else (f"{q:g}" if q < 0 else "0")
        typ = f"{m['type']} ({m['src']})" if m["src"] == "HansaWorld" else m["type"]
        party = f"{m['party_code']} - {m['party']}" if m["party_code"] else m["party"]
        notes = " • ".join(x for x in (f"مرجع {m['ref']}" if m["ref"] != "-" else "",
                                        f"مربوط {m['link']}" if m["link"] else "", m["notes"]) if x) or "-"
        rows.append((m["date"], typ, party, qty_str, f"{m['balance']:,.3f}", fmt(m["cost"]), fmt(m["value"]), notes))
    return rows


def fetch_item_card(item_code, d_from, d_to):
    """يرجع (الصفوف المنسقة، الرصيد النهائي كرقم). كل الحركات التفصيلية (HansaWorld + الفواتير المباشرة والاستلام/التسليم
    والمرتجعات والخصومات اللاحقة). السعر = التكلفة، والقيمة = الكمية × التكلفة."""
    card = item_card(item_code, d_from, d_to)
    return item_card_rows(card, d_from=d_from), card["final"]


def item_first_move_date(item_code):
    """تاريخ أول حركة للصنف (HansaWorld أو النظام) — بداية الفلتر الافتراضية في كارت الصنف، أو None."""
    sqls = ["SELECT MIN(mdate) FROM v_stock_in WHERE item_code=?", "SELECT MIN(mdate) FROM v_stock_out WHERE item_code=?",
            "SELECT MIN(return_date) FROM returns_invoices WHERE item_code=?",
            "SELECT MIN(move_date) FROM hw_item_history WHERE item_code=?"]
    found = []
    with db() as conn:
        for sql in sqls:
            try:
                v = conn.execute(sql, (item_code,)).fetchone()[0]
            except sqlite3.OperationalError:
                continue
            if v:
                found.append(v)
    return min(found) if found else None


def _partner_entries(conn, codes):
    """كل حركات أكواد الطرف: [(date, created_at, ref, type, desc, debit, credit, key, side)].
    side: c = حركة عميل، s = حركة مورد، b = مشتركة."""
    multi = len(codes) > 1
    entries = []
    for code in codes:
        tag = f"[{code}] " if multi else ""
        pr = conn.execute("SELECT is_customer, is_supplier, COALESCE(opening_debit,0), COALESCE(opening_credit,0), "
                          "COALESCE(opening_date,'') FROM business_partners WHERE partner_code=?", (code,)).fetchone()
        if pr is None:
            continue
        is_c, is_s = bool(pr[0]), bool(pr[1])
        own = "b" if (is_c and is_s) else ("c" if is_c else ("s" if is_s else "b"))
        if pr[2] > 0.005 or pr[3] > 0.005:
            entries.append((pr[4], '', "افتتاحي", "رصيد افتتاحي", f"{tag}رصيد أول المدة (ما عليه / ما له)", pr[2], pr[3],
                            ("opening", code), own))
        for r in conn.execute("SELECT COALESCE(invoice_no, invoice_id), invoice_date, created_at, total_amount, notes, "
                              "invoice_id FROM sales_invoices WHERE customer_code=?", (code,)):
            entries.append((r[1], r[2] or '', str(r[0]), "فاتورة مبيعات", tag + (r[4] or "-"), r[3], 0.0, ("sales", r[5]), "c"))
        for r in conn.execute("SELECT COALESCE(invoice_no, purchase_id), purchase_date, created_at, total_amount, notes, "
                              "purchase_id, COALESCE(source_type,'مباشر') FROM purchase_invoices WHERE supplier_code=?",
                              (code,)):
            typ = "فاتورة شراء مباشر" if r[6] == "مباشر" else "فاتورة مشتريات"
            entries.append((r[1], r[2] or '', str(r[0]), typ, tag + (r[4] or "-"), 0.0, r[3], ("purchase", r[5]), "s"))
        for r in conn.execute("SELECT return_id, return_date, created_at, return_type, item_name, total_amount, notes "
                              "FROM returns_invoices WHERE party_code=?", (code,)):
            desc = tag + f"{r[4]}" + (f" - {r[6]}" if r[6] else "")
            if r[3] == SALES_RETURN:
                entries.append((r[1], r[2] or '', f"مرتجع-{r[0]}", r[3], desc, 0.0, r[5], ("return", r[0]), "c"))
            else:
                entries.append((r[1], r[2] or '', f"مرتجع-{r[0]}", r[3], desc, r[5], 0.0, ("return", r[0]), "s"))
        for r in conn.execute("SELECT discount_id, discount_no, discount_date, created_at, kind, total_amount, "
                              "COALESCE(reason,''), COALESCE(notes,'') FROM post_discounts WHERE party_code=?", (code,)):
            desc = tag + (" - ".join(x for x in (r[6], r[7]) if x) or "-")
            if r[4] == "purchase":   # إشعار خصم من المورد: يخفض ما له علينا (مدين)
                entries.append((r[2], r[3] or '', r[1] or f"PD-{r[0]}", "خصم لاحق من المورد", desc, r[5], 0.0,
                                ("discount", r[0]), "s"))
            else:                    # إشعار دائن للعميل: يخفض ما عليه لنا (دائن)
                entries.append((r[2], r[3] or '', r[1] or f"CN-{r[0]}", "إشعار دائن (خصم لاحق)", desc, 0.0, r[5],
                                ("discount", r[0]), "c"))
        for r in conn.execute("SELECT trans_id, trans_date, created_at, trans_type, description, amount, "
                              "COALESCE(alloc_type,'') FROM cash_transactions WHERE party_code=? AND category=?",
                              (code, CAT_PARTNER)):
            side = ("c" if r[6] == "sales" else "s" if r[6] == "purchase" else
                    own if own in ("c", "s") else ("c" if r[3] == "قبض" else "s"))
            if r[3] == "قبض":
                entries.append((r[1], r[2] or '', f"سند-{r[0]}", "سند قبض", tag + (r[4] or "-"), 0.0, r[5], ("cash", r[0]), side))
            else:
                entries.append((r[1], r[2] or '', f"سند-{r[0]}", "سند صرف", tag + (r[4] or "-"), r[5], 0.0, ("cash", r[0]), side))
    return entries


def _scope_filter(entries, scope):
    allowed = {"customer": ("c", "b"), "supplier": ("s", "b")}.get(scope, ("c", "s", "b"))
    return [e for e in entries if e[8] in allowed]


def _split_period(entries, d_from=None, d_to=None):
    """(رصيد ما قبل الفترة، حركات الفترة). الرصيد الافتتاحي بدون تاريخ يُعتبر أقدم حركة.
    بدون تواريخ = كل الحركات من بداية العمل حتى الآن (رصيد لحظي)."""
    prior, inside = 0.0, []
    for e in entries:
        d = e[0] or ""
        if d_from and d < d_from:
            prior += e[5] - e[6]
        elif not d_to or d <= d_to:
            inside.append(e)
    return round(prior, 2), inside


def period_label(d_from=None, d_to=None):
    if d_from and d_to:
        return f"الفترة من {d_from} إلى {d_to}"
    if d_from:
        return f"من {d_from} حتى الآن"
    if d_to:
        return f"من بداية العمل حتى {d_to}"
    return f"رصيد لحظي تراكمي من بداية العمل حتى الآن ({today()})"


def fetch_partner_statement(partner_code, with_keys=False, scope="both", linked=True, d_from=None, d_to=None):
    """كشف الحساب الموحد. الرصيد = مدين - دائن (الموجب = الطرف مدين لنا).
    linked=True: يجمع كل أكواد نفس الكيان (عميل C#### ومورد S#### بنفس الاسم/الرقم الضريبي) لإجراء المقاصة.
    scope: both (الصافي) | customer (حركات العميل فقط) | supplier (حركات المورد فقط).
    d_from/d_to: نطاق زمني اختياري؛ ما قبل d_from يظهر سطراً واحداً «رصيد ما قبل الفترة».
    with_keys=True يضيف مفاتيح المصدر لكل سطر لفتح تفاصيل الحركة."""
    d_from, d_to = (parse_date(d_from or "") or None), (parse_date(d_to or "") or None)
    if d_from and d_to and d_from > d_to:
        raise UserError("تاريخ البداية بعد تاريخ النهاية.")
    with db() as conn:
        info = conn.execute("SELECT name, phone, cr_number, vat_number, address FROM business_partners "
                            "WHERE partner_code=?", (partner_code,)).fetchone()
        if not info:
            raise UserError("الشريك غير موجود")
        codes = partner_link_map(conn).get(partner_code, [partner_code]) if linked else [partner_code]
        entries = _scope_filter(_partner_entries(conn, codes), scope)
    entries.sort(key=lambda e: (e[0] or '', e[1]))
    prior, entries = _split_period(entries, d_from, d_to)
    rows, keys, balance, tot_d, tot_c = [], [], 0.0, 0.0, 0.0
    if d_from and abs(prior) > 0.005:
        entries.insert(0, (d_from, "", "سابق", "رصيد سابق", f"رصيد ما قبل {d_from} (مُرحَّل)",
                           prior if prior > 0 else 0.0, -prior if prior < 0 else 0.0, None, "b"))
    for d, _, ref, typ, desc, debit, credit, key, _side in entries:
        keys.append(key)
        balance += debit - credit
        tot_d += debit
        tot_c += credit
        side = "مدين" if balance > 0.005 else ("دائن" if balance < -0.005 else "")
        rows.append((ref, d or "-", typ, desc, fmt(debit), fmt(credit), f"{abs(balance):,.2f} {side}".strip()))
    if with_keys:
        return info, rows, tot_d, tot_c, balance, keys
    return info, rows, tot_d, tot_c, balance


def partners_balance_summary(scope="both", d_from=None, d_to=None, q="", hide_zero=False):
    """العرض الإجمالي: سطر واحد لكل عميل/مورد بإجمالي رصيده (بدون تفاصيل الحركات).
    scope=customer: كل عميل بحركاته كعميل | supplier: كل مورد | both: كل كيان (العميل والمورد المرتبطان سطر واحد صافٍ).
    يرجع dict(rows=[dict(code, codes, name, role, phone, prior, debit, credit, balance, moves, last)], totals=...)."""
    d_from, d_to = (parse_date(d_from or "") or None), (parse_date(d_to or "") or None)
    if d_from and d_to and d_from > d_to:
        raise UserError("تاريخ البداية بعد تاريخ النهاية.")
    qn = norm_name(q or "").strip()
    with db() as conn:
        parts = {r[0]: r[1:] for r in conn.execute(
            "SELECT partner_code, name, is_customer, is_supplier, COALESCE(phone,''), COALESCE(vat_number,'') "
            "FROM business_partners ORDER BY partner_code")}
        links = partner_link_map(conn) if scope == "both" else {}
        groups, seen = [], set()
        for code, (name, is_c, is_s, _ph, _vat) in parts.items():
            if scope == "customer" and not is_c or scope == "supplier" and not is_s or code in seen:
                continue
            codes = links.get(code, [code]) if scope == "both" else [code]
            seen.update(codes)
            groups.append(codes)
        out = []
        for codes in groups:
            first = codes[0]
            name, _c, _s, phone, vat = parts[first]
            if qn and not any(qn in norm_name(f"{c} {parts[c][0]} {parts[c][3]} {parts[c][4]}") for c in codes if c in parts):
                continue
            entries = _scope_filter(_partner_entries(conn, codes), scope)
            prior, inside = _split_period(entries, d_from, d_to)
            debit = round(sum(e[5] for e in inside), 2)
            credit = round(sum(e[6] for e in inside), 2)
            bal = round(prior + debit - credit, 2)
            if not inside and abs(prior) < 0.005:
                continue  # طرف بدون أي حركة أو رصيد في هذا النطاق
            if hide_zero and abs(bal) < 0.005:
                continue
            roles = set()
            for c in codes:
                if parts.get(c):
                    if parts[c][1]:
                        roles.add("عميل")
                    if parts[c][2]:
                        roles.add("مورد")
            role = {"customer": "عميل", "supplier": "مورد"}.get(scope)
            if not role:
                role = ("عميل ومورد (مقاصة)" if roles == {"عميل", "مورد"} else (roles.pop() if roles else "-"))
            out.append(dict(code=first, codes=codes, name=name, role=role, phone=phone, prior=prior, debit=debit,
                            credit=credit, balance=bal, moves=len(inside),
                            last=max((e[0] or "" for e in inside), default="") or "-"))
    out.sort(key=lambda r: (-abs(r["balance"]), r["name"]))
    totals = dict(prior=round(sum(r["prior"] for r in out), 2), debit=round(sum(r["debit"] for r in out), 2),
                  credit=round(sum(r["credit"] for r in out), 2),
                  receivable=round(sum(r["balance"] for r in out if r["balance"] > 0), 2),
                  payable=round(-sum(r["balance"] for r in out if r["balance"] < 0), 2))
    totals["net"] = round(totals["receivable"] - totals["payable"], 2)
    return dict(rows=out, totals=totals)


def balance_text(b):
    return f"{abs(b):,.2f} " + ("مدين (عليه)" if b > 0.005 else ("دائن (له)" if b < -0.005 else "")).strip()


SUMMARY_HEADERS = ["الكود", "الاسم", "الصفة", "رصيد سابق", "مدين الفترة", "دائن الفترة", "الرصيد", "عدد الحركات",
                   "آخر حركة"]


def summary_table_rows(res, with_prior=True):
    return [(" + ".join(r["codes"]), r["name"], r["role"], balance_text(r["prior"]) if with_prior else "",
             fmt(r["debit"]), fmt(r["credit"]), balance_text(r["balance"]), r["moves"], r["last"]) for r in res["rows"]]


def company_name():
    """اسم المنشأة المسجل في بيانات الشركة (يظهر على المستندات: الفواتير والكشوف والتقارير)."""
    try:
        with db() as conn:
            r = conn.execute("SELECT company_name FROM company_profile WHERE id=1").fetchone()
        return (r[0] if r and r[0] else "") or BRAND_AR
    except sqlite3.Error:
        return BRAND_AR


def income_figures(d_from, d_to, dep_to=None):
    """أرقام قائمة الدخل للفترة — المصدر الوحيد لقائمة الدخل ولوحة المؤشرات حتى يتطابقا دائماً:
        صافي المبيعات = المبيعات قبل الضريبة − مرتجعات المبيعات − الإشعارات الدائنة للعملاء
        تكلفة المبيعات = تكلفة بنود الفواتير الجديدة (متوسط مرجح) − تكلفة المرتجعات + تكلفة فترة HansaWorld (جرد دوري)
        مجمل الربح = صافي المبيعات − تكلفة المبيعات
        فترة HansaWorld: تُسوّى تكلفتها لتساوي دفتر HansaWorld (41110 + 41201) لكل فترة مرجعية (الميزان/قوائم الدخل)
        المصروفات التشغيلية = سندات صرف المصروفات + القيود اليومية على حسابات المصروفات (ومنها قيد تجميع أرصدة الميزان)
            مصنفة: بيعية وتسويقية (43) / إدارية وعمومية (5) / أخرى (41125 إيجار السيارات …)
        صافي الربح = مجمل الربح + الإيرادات الأخرى − المصروفات التشغيلية − الإهلاك.
    dep_to: نهاية فترة الإهلاك (الافتراضي d_to)."""
    from_dt = datetime.strptime(d_from, "%Y-%m-%d").date()
    to_dt = datetime.strptime(dep_to or d_to, "%Y-%m-%d").date()
    with db() as conn:
        def one(sql, params=()):
            row = conn.execute(sql, params).fetchone()
            return tuple((v or 0.0) for v in row)

        sales_sub, sales_vat = one("SELECT SUM(subtotal), SUM(vat_amount) FROM sales_invoices "
                                   "WHERE invoice_date BETWEEN ? AND ?", (d_from, d_to))
        (sales_disc,) = one("SELECT SUM(l.discount_amount) FROM sales_invoice_lines l JOIN sales_invoices h "
                            "ON h.invoice_id=l.invoice_id WHERE h.invoice_date BETWEEN ? AND ?", (d_from, d_to))
        rs_sub, rs_vat, rs_cost = one(
            "SELECT SUM(subtotal), SUM(vat_amount), SUM(quantity*unit_cost) FROM returns_invoices "
            "WHERE return_type=? AND return_date BETWEEN ? AND ?", (SALES_RETURN, d_from, d_to))
        (sold_cost,) = one("SELECT SUM(l.quantity*l.unit_cost) FROM sales_invoice_lines l "
                           "JOIN sales_invoices h ON h.invoice_id=l.invoice_id "
                           "WHERE h.invoice_date BETWEEN ? AND ?", (d_from, d_to))
        pur_sub, pur_vat = one("SELECT SUM(subtotal), SUM(vat_amount) FROM purchase_invoices "
                               "WHERE purchase_date BETWEEN ? AND ?", (d_from, d_to))
        rp_sub, rp_vat, rp_cost = one("SELECT SUM(subtotal), SUM(vat_amount), SUM(quantity*unit_cost) FROM returns_invoices "
                                      "WHERE return_type=? AND return_date BETWEEN ? AND ?",
                                      (PURCHASE_RETURN, d_from, d_to))
        expenses = [(name, amt) for name, _, amt in expense_breakdown(conn, d_from, d_to)]
        opex_groups = expense_groups(conn, d_from, d_to)
        other_income = [(name, amt) for _, typ, _, name, amt in journal_pl(conn, d_from, d_to) if typ == TYPE_REV]
        dsc = discount_totals(conn, d_from, d_to)[""]
        hw = hw_cogs(conn, d_from, d_to)[""]
        earned = dsc["earned"] - hw["credits"]
        if earned > 0.005:
            other_income.insert(0, ("خصم مكتسب من الموردين (خصومات لاحقة)", earned))
        # المرتجع للمورد يخفض المخزون بمتوسط التكلفة ويخفض رصيد المورد بسعر الإرجاع: الفرق ربح/خسارة فعلية
        rp_var = round(rp_sub - rp_cost, 2)
        if abs(rp_var) > 0.005:
            other_income.append(("فروق أسعار مردودات المشتريات (سعر الإرجاع - متوسط التكلفة)", rp_var))
        assets = conn.execute("SELECT purchase_cost, salvage_value, useful_life_years, purchase_date "
                              "FROM fixed_assets").fetchall()
        (stock_value,) = one("SELECT SUM(quantity*cost_price) FROM inventory")

    dep = 0.0
    for cost, salvage, life, p_date in assets:
        dep += (accumulated_depreciation(cost, salvage, life, p_date, to_dt)
                - accumulated_depreciation(cost, salvage, life, p_date, from_dt - timedelta(days=1)))

    net_sales = sales_sub - rs_sub - dsc["sales_sub"]
    cogs_lines = sold_cost - rs_cost
    cogs = cogs_lines + hw["cogs"]
    gross = net_sales - cogs
    total_other = sum(e[1] for e in other_income)
    opex = sum(e[1] for e in expenses)
    net_profit = gross + total_other - opex - dep
    out_vat = sales_vat - rs_vat - dsc["sales_vat"]
    in_vat = pur_vat - rp_vat - dsc["pur_vat"]
    return dict(sales_sub=sales_sub, sales_disc=sales_disc, rs_sub=rs_sub, dsc=dsc, hw=hw, net_sales=net_sales,
                cogs_lines=cogs_lines, cogs=cogs, gross=gross, other_income=other_income, total_other=total_other,
                expenses=expenses, opex=opex, opex_groups=opex_groups, dep=dep, total_exp=opex + dep, net_profit=net_profit,
                net_purchases=pur_sub - rp_sub - dsc["pur_sub"], stock_value=stock_value,
                out_vat=out_vat, in_vat=in_vat, net_vat=out_vat - in_vat)


def build_income_report(d_from, d_to):
    f = income_figures(d_from, d_to)
    sales_sub, sales_disc, rs_sub, dsc, hw = f["sales_sub"], f["sales_disc"], f["rs_sub"], f["dsc"], f["hw"]
    net_sales, cogs_lines, cogs, gross = f["net_sales"], f["cogs_lines"], f["cogs"], f["gross"]
    other_income, total_other, expenses, dep = f["other_income"], f["total_other"], f["expenses"], f["dep"]
    total_exp, net_profit, stock_value = f["total_exp"], f["net_profit"], f["stock_value"]
    out_vat, in_vat, net_vat = f["out_vat"], f["in_vat"], f["net_vat"]

    def line(label, value, indent=1):
        return f"{' ' * indent}{label:<46}{value:>14,.2f} ر.س"

    sep = "=" * 66
    thin = "-" * 66
    out = [sep, f"              {company_name()}",
           "          قائمة الدخل وإقرار ضريبة القيمة المضافة",
           f"          الفترة: من {d_from} إلى {d_to}", sep,
           *([line("إجمالي المبيعات قبل الخصم الخاص", sales_sub + sales_disc),
               line("(-) الخصم الخاص على الأصناف", sales_disc)] if sales_disc > 0.005 else []),
           line("إجمالي المبيعات (قبل الضريبة، بعد الخصم)", sales_sub),
           line("(-) مرتجعات المبيعات", rs_sub),
           *([line("(-) خصومات لاحقة للعملاء (إشعارات دائنة)", dsc["sales_sub"])] if dsc["sales_sub"] > 0.005 else []),
           line("صافي المبيعات", net_sales), thin,
           *([line("مشتريات فترة HansaWorld (قبل الضريبة)", hw["purchases"], 3),
              line("(-) إشعارات خصم الموردين المستوردة", hw["credits"], 3),
              line("(-) المخزون الختامي في تاريخ الترحيل", hw["closing"], 3),
              *([line("± تسوية لدفتر HansaWorld (41110 + 41201)", hw["ref_adj"], 3)] if abs(hw["ref_adj"]) > 0.005 else []),
              line("= تكلفة مبيعات فترة HansaWorld (جرد دوري" + (" مُسوّى لدفترها)" if abs(hw["ref_adj"]) > 0.005 else ")"),
                   hw["cogs"], 3)]
             if abs(hw["purchases"]) + abs(hw["closing"]) + abs(hw["ref_adj"]) > 0.005 else []),
           *([line("تكلفة مبيعات المستندات الجديدة (بالمتوسط المرجح)", cogs_lines, 3)]
             if abs(hw["cogs"]) > 0.005 and abs(cogs_lines) > 0.005 else []),
           line("(-) تكلفة البضاعة المباعة" + (" (إجمالي)" if abs(hw["cogs"]) > 0.005 else " (بالمتوسط المرجح)"), cogs),
           line("مجمل الربح", gross), thin]
    if other_income:
        out.append(" إيرادات أخرى (من القيود اليومية):")
        for name, amt in other_income:
            out.append(line(f"+ {name}", amt, 3))
        out += [line("إجمالي الإيرادات الأخرى", total_other), thin]
    out.append(" المصروفات التشغيلية (سندات الصرف + القيود اليومية):")
    if expenses:
        for name, amt in expenses:
            out.append(line(f"- {name}", amt, 3))
    else:
        out.append("   - لا توجد مصروفات مسجلة في الفترة")
    og = f["opex_groups"]
    if sum(1 for v in og.values() if abs(v) > 0.005) > 1:
        out += [line(f"منها {OPEX_GROUP_LABELS[k]}", og[k], 3) for k in ("selling", "admin", "other") if abs(og[k]) > 0.005]
    out += [line("إجمالي المصروفات التشغيلية", total_exp - dep), thin,
            line("مصروف الإهلاك (إهلاك الأصول الثابتة)", dep), thin,
            line("إجمالي المصروفات (التشغيلية + الإهلاك)", total_exp), sep,
            line("صافي الربح", net_profit), sep,
            " معلومات إضافية:",
            line("مشتريات الفترة (قبل الضريبة، بعد المرتجعات والخصومات)", f["net_purchases"]),
            *([line("منها خصم لاحق خُفّضت به تكلفة المخزون", dsc["stock"])] if dsc["stock"] > 0.005 else []),
            line("قيمة المخزون الحالية (بالتكلفة)", stock_value), sep,
            " إقرار ضريبة القيمة المضافة (15%):",
            line("ضريبة المخرجات (المبيعات - المرتجعات والإشعارات)", out_vat),
            line("ضريبة المدخلات (المشتريات - المرتجعات والخصومات)", in_vat), thin,
            line("صافي الضريبة " + ("المستحقة لهيئة الزكاة والضريبة والجمارك" if net_vat >= 0
                                  else "المستردة / رصيد دائن"), abs(net_vat)), sep,
            " ملاحظة: ضريبة المدخلات لا تشمل ضريبة المصروفات العامة المدفوعة بسندات.",
            " ملاحظة: تكلفة المبيعات والإهلاك تُحسب تلقائياً، فلا تُسجَّل بقيود يدوية (تُستبعد منعاً للازدواج).", ""]
    return "\n".join(out)


# ============================================================
# ملفات PDF
# ============================================================
def zatca_qr_base64(seller, vat_no, timestamp, total, vat):
    """TLV مشفر Base64 حسب متطلبات هيئة الزكاة والضريبة والجمارك (المرحلة الأولى)."""
    payload = b''
    for tag, val in enumerate([seller, vat_no, timestamp, f"{total:.2f}", f"{vat:.2f}"], start=1):
        vb = str(val).encode('utf-8')[:255]
        payload += bytes([tag, len(vb)]) + vb
    return base64.b64encode(payload).decode()


def qr_drawing(data, size=95):
    widget = QrCodeWidget(data)
    b = widget.getBounds()
    bw, bh = b[2] - b[0], b[3] - b[1]
    d = Drawing(size, size, transform=[size / bw, 0, 0, size / bh, 0, 0])
    d.add(widget)
    return d


# ============================================================
# الجهة المصدرة للمستندات (الفواتير والإشعارات): البيانات القانونية للمنشأة من «بيانات الشركة»
# (هوية البرنامج «نظام الحاجري المالي» منفصلة عنها ولا تظهر إلا كسطر صغير في تذييل المستند)
# ============================================================
COMPANY_DEFAULTS = dict(
    company_name="شركة فود زون للتجارة", company_name_en="Food Zone Company",
    vat_number="314482150400003", cr_number="7017428496", phone="+966 56 3390 246",
    address="حي السلي، الرياض، المملكة العربية السعودية", address_en="Al Sulay District, Riyadh, Kingdom of Saudi Arabia",
    bank_name="بنك الرياض", iban="SA5920000002304289769940", rep_name="")
COMPANY_FIELDS = ("company_name", "company_name_en", "vat_number", "cr_number", "phone", "address", "address_en",
                  "bank_name", "iban", "rep_name")
COMPANY_LOGO_FILES = ("company_logo.png", "company_logo.jpg", "لوجو.jpeg", "لوجو.jpg", "لوجو.png")


def company_info():
    """بيانات المنشأة المصدرة للفواتير كقاموس (الحقول الفارغة تأخذ القيم القانونية الافتراضية)."""
    with db() as conn:
        r = conn.execute(f"SELECT {', '.join(COMPANY_FIELDS)} FROM company_profile WHERE id=1").fetchone()
    info = dict(zip(COMPANY_FIELDS, r)) if r else {}
    return {k: (info.get(k) or COMPANY_DEFAULTS[k]) for k in COMPANY_FIELDS}


def save_company_info(values):
    """حفظ بيانات المنشأة (للفواتير). الرقم الضريبي 15 رقماً يبدأ وينتهي بـ 3، والآيبان سعودي SA + 22 رقماً."""
    v = {k: (values.get(k) or "").strip() for k in COMPANY_FIELDS}
    if not v["company_name"]:
        raise UserError("اسم المنشأة مطلوب.")
    vat = normalize_digits(v["vat_number"]).replace(" ", "")
    if vat and not re.fullmatch(r"3\d{13}3", vat):
        raise UserError("الرقم الضريبي يجب أن يكون 15 رقماً يبدأ وينتهي بالرقم 3.")
    iban = normalize_digits(v["iban"]).replace(" ", "").upper()
    if iban and not re.fullmatch(r"SA\d{22}", iban):
        raise UserError("رقم الآيبان السعودي يبدأ بـ SA ويليه 22 رقماً.")
    v["vat_number"], v["iban"] = vat, iban
    with db() as conn:
        conn.execute(f"UPDATE company_profile SET {', '.join(f'{k}=?' for k in COMPANY_FIELDS)} WHERE id=1",
                     [v[k] for k in COMPANY_FIELDS])


def format_iban(iban):
    iban = (iban or "").replace(" ", "")
    return " ".join(iban[i:i + 4] for i in range(0, len(iban), 4))


# ---------------------------------------------------------- التفقيط (المبلغ كتابةً بالعربية)
_ONES = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة", "عشرة",
         "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر", "خمسة عشر", "ستة عشر", "سبعة عشر", "ثمانية عشر", "تسعة عشر"]
_TENS = ["", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون", "ثمانون", "تسعون"]
_HUNDREDS = ["", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة"]
# (مفرد، مثنى، جمع 3-10، تمييز 11-99)
_SCALES = [("", "", "", ""), ("ألف", "ألفان", "آلاف", "ألفاً"), ("مليون", "مليونان", "ملايين", "مليوناً"),
           ("مليار", "ملياران", "مليارات", "ملياراً")]


def _words_999(n):
    parts = []
    h, r = divmod(n, 100)
    if h:
        parts.append(_HUNDREDS[h])
    if r:
        if r < 20:
            parts.append(_ONES[r])
        else:
            t, o = divmod(r, 10)
            parts.append(f"{_ONES[o]} و{_TENS[t]}" if o else _TENS[t])
    return " و".join(parts)


def number_to_arabic_words(n):
    n = int(n)
    if n == 0:
        return "صفر"
    groups = []
    while n:
        n, g = divmod(n, 1000)
        groups.append(g)
    parts = []
    for idx in range(len(groups) - 1, -1, -1):
        g = groups[idx]
        if not g:
            continue
        if idx == 0:
            parts.append(_words_999(g))
            continue
        one, two, plural, acc = _SCALES[idx]
        if g == 1:
            parts.append(one)
        elif g == 2:
            parts.append(two)
        elif 3 <= g <= 10:
            parts.append(f"{_ONES[g]} {plural}")
        else:
            r = g % 100
            word = acc if 11 <= r <= 99 else (plural if 3 <= r <= 10 else one)
            parts.append(f"{_words_999(g)} {word}")
    return " و".join(parts)


def tafqeet(amount, currency="ريال سعودي", fraction="هللة"):
    """تفقيط المبلغ: فقط ألف ومائتان وخمسون ريال سعودي وخمس وعشرون هللة لا غير."""
    amount = round(abs(float(amount or 0)) + 1e-9, 2)
    whole = int(amount)
    halala = int(round((amount - whole) * 100))
    if halala == 100:
        whole, halala = whole + 1, 0
    text = f"فقط {number_to_arabic_words(whole)} {currency}"
    if halala:
        text += f" و{number_to_arabic_words(halala)} {fraction}"
    return text + " لا غير"


# ---------------------------------------------------------- تصميم المستند
DOC_INK, DOC_ACCENT, DOC_ACCENT2, DOC_SOFT, DOC_LINE = "#0f172a", "#1e3a8a", "#0d9488", "#f1f5fb", "#d5deeb"


def _doc_styles():
    def s(name, size=8.5, align=2, color=DOC_INK, leading=1.45):
        return ParagraphStyle(name, fontName=FONT_NAME, fontSize=size, leading=size * leading, alignment=align,
                              textColor=colors.HexColor(color))
    return dict(r=s("r"), l=s("l", align=0), c=s("c", align=1), muted_r=s("mr", 7.5, 2, "#64748b"),
                muted_l=s("ml", 7.5, 0, "#64748b"), muted_c=s("mc", 7, 1, "#64748b"),
                name_r=s("nr", 13, 2, DOC_ACCENT), name_l=s("nl", 12, 0, DOC_ACCENT),
                band_r=s("br", 14, 2, "#ffffff"), band_l=s("bl", 12, 0, "#ffffff"),
                head=s("h", 7.2, 1, "#ffffff", 1.3), cell=s("cell", 8, 1), cell_r=s("cellr", 8, 2),
                tot_r=s("tr", 9, 2), tot_l=s("tl", 9, 0), big_r=s("bgr", 10.5, 2, "#ffffff"),
                big_l=s("bgl", 10.5, 0, "#ffffff"), sec=s("sec", 9, 2, DOC_ACCENT), words=s("w", 9, 2, DOC_INK))


_WATERMARK_CACHE = {}


def _watermark_image():
    """شعار المنشأة كعلامة مائية شفافة هادئة (حواف متلاشية) لخلفية الفاتورة. يرجع ImageReader أو None."""
    path = find_app_file(COMPANY_LOGO_FILES)
    if not path:
        return None
    key = (path, os.path.getmtime(path))
    if key not in _WATERMARK_CACHE:
        try:
            from reportlab.lib.utils import ImageReader
            from PIL import ImageEnhance
            im = Image.open(path).convert("RGB")
            im.thumbnail((560, 560))
            w, h = im.size
            # فصل الشعار عن خلفيته (خشب/لون داكن أو أبيض): الشعار = البكسلات المختلفة بوضوح عن لون الخلفية
            corners = [im.getpixel(xy) for xy in ((0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1))]
            bg = tuple(sum(c[k] for c in corners) / 4 for k in range(3))
            bg_v = max(bg)
            hsv = im.convert("HSV")
            px_src, px_hsv = im.load(), hsv.load()
            alpha = Image.new("L", (w, h), 0)
            px = alpha.load()
            cx, cy, rmax = w / 2, h / 2, min(w, h) / 2
            for y in range(h):
                for x in range(w):
                    r_, g_, b_ = px_src[x, y]
                    if bg_v < 160:   # خلفية داكنة: الشعار أفتح وأكثر تشبعاً
                        v = px_hsv[x, y][2]
                        m = max(0.0, min(1.0, (v - max(bg_v + 110, 165)) / 45.0))
                    else:            # خلفية فاتحة: الشعار أغمق من الخلفية
                        d = (abs(r_ - bg[0]) + abs(g_ - bg[1]) + abs(b_ - bg[2])) / 3
                        m = max(0.0, min(1.0, (d - 25) / 60.0))
                    rr = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5 / rmax
                    edge = 1.0 if rr < 0.85 else max(0.0, (1.0 - rr) / 0.15)
                    px[x, y] = int(13 * m * edge)  # أقصى عتامة ≈5% فقط → لا تؤثر على قراءة النصوص
            im = ImageEnhance.Color(im).enhance(0.6)  # ألوان الشعار هادئة
            im.putalpha(alpha)
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            _WATERMARK_CACHE[key] = ImageReader(buf)
        except Exception:
            _WATERMARK_CACHE[key] = None
    return _WATERMARK_CACHE[key]


def _company_logo_flowable(size=64):
    path = find_app_file(COMPANY_LOGO_FILES)
    if not path:
        return ""
    try:
        from reportlab.platypus import Image as RLImage
        return RLImage(path, width=size, height=size)
    except Exception:
        return ""


def _kv_table(rows, S, width):
    """جدول بيانات (القيمة يساراً والتسمية العربية/الإنجليزية يميناً)."""
    data = [[P(v, S["r"]), P(k, S["muted_r"])] for k, v in rows]
    t = Table(data, colWidths=[width * 0.6, width * 0.4])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
                           ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor(DOC_LINE))]))
    return t


def _box(title, inner, S, width):
    t = Table([[P(title, S["sec"])], [inner]], colWidths=[width])
    t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(DOC_LINE)),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(DOC_SOFT)),
                           ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor(DOC_ACCENT2)),
                           ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                           ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
    return t


def _render_document(path, d):
    """d: title_ar, title_en, meta (قائمة)، party_title، party (قائمة)، columns [(عنوان، عرض)]، rows، totals
    [(تسمية، مبلغ، مميز؟)]، grand (مبلغ للتفقيط)، qr (نص TLV أو None)، notes، signatures، footer_note."""
    comp = company_info()
    S = _doc_styles()
    W = 535
    doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=26, bottomMargin=40,
                            title=f"{d['title_en']} {d['number']}", author=comp["company_name_en"])
    el = []

    # 1) رأس المنشأة: عربي يميناً، الشعار وسطاً، إنجليزي يساراً
    right = [P(comp["company_name"], S["name_r"]), P(comp["address"], S["muted_r"]),
             P(f"الرقم الضريبي: {comp['vat_number']}", S["r"]), P(f"السجل التجاري: {comp['cr_number']}", S["r"])]
    left = [P(comp["company_name_en"], S["name_l"]), P(comp["address_en"], S["muted_l"]),
            P(f"VAT No: {comp['vat_number']}", S["l"]), P(f"C.R No: {comp['cr_number']}", S["l"])]
    if comp["phone"]:  # الهاتف بالجانب الإنجليزي فقط (الأرقام المتعددة المقاطع تنعكس داخل سطر عربي)
        left.append(P(f"Tel: {comp['phone']}", S["muted_l"]))
    if d.get("qr"):  # رمز الزكاة بارز في أعلى الفاتورة (المكان المخصص له)
        center = [qr_drawing(d["qr"], 92), P("ZATCA QR", S["muted_c"])]
    else:
        center = _company_logo_flowable()
    head = Table([[left, center, right]], colWidths=[215, 105, 215])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 0), (1, 0), "CENTER"),
                              ("LINEBELOW", (0, 0), (-1, 0), 2, colors.HexColor(DOC_ACCENT))]))
    el += [head, Spacer(1, 8)]

    # 2) شريط عنوان المستند
    band = Table([[P(d["title_en"], S["band_l"]), P(d["title_ar"], S["band_r"])]], colWidths=[W / 2, W / 2])
    band.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(DOC_ACCENT)),
                              ("LINEAFTER", (0, 0), (0, 0), 0, colors.white),
                              ("LINEBELOW", (0, 0), (-1, -1), 3, colors.HexColor(DOC_ACCENT2)),
                              ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                              ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10)]))
    el += [band, Spacer(1, 8)]

    # 3) بيانات المستند وبيانات الطرف
    half = (W - 10) / 2
    info = Table([[_box(d["party_title"], _kv_table(d["party"], S, half - 12), S, half), "",
                   _box("بيانات المستند | Document Details", _kv_table(d["meta"], S, half - 12), S, half)]],
                 colWidths=[half, 10, half])
    info.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    el += [info, Spacer(1, 10)]

    # 4) جدول البنود
    heads = [P(h, S["head"]) for h, _ in d["columns"]]
    widths = [w for _, w in d["columns"]]
    data = [rtl(heads)]
    desc_idx = d.get("desc_col", 2)
    for r in d["rows"]:
        data.append(rtl([P(v, S["cell_r"] if i == desc_idx else S["cell"]) for i, v in enumerate(r)]))
    items = Table(data, colWidths=rtl(widths), repeatRows=1)
    style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(DOC_INK)), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor(DOC_LINE)),
             ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(DOC_LINE)),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f8fafc")))
    items.setStyle(TableStyle(style))
    el += [items, Spacer(1, 10)]

    # 5) الإجماليات + رمز QR
    tot_rows = []
    hl = []
    for i, (label, amount, strong) in enumerate(d["totals"]):
        st_l, st_r = (S["big_l"], S["big_r"]) if strong else (S["tot_l"], S["tot_r"])
        tot_rows.append([P(f"{fmt(amount)} SAR", st_l), P(label, st_r)])
        if strong:
            hl.append(i)
    tot = Table(tot_rows, colWidths=[110, 190])
    tstyle = [("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor(DOC_LINE)),
              ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(DOC_LINE)),
              ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
              ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
    for i in hl:
        tstyle.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor(DOC_ACCENT2)))
    tot.setStyle(TableStyle(tstyle))
    if d.get("qr"):
        qr_cell = [P("رمز الاستجابة السريعة (ZATCA) في أعلى الفاتورة\nيمكن مسحه للتحقق من بيانات الفاتورة الضريبية",
                     S["muted_c"])]
    else:
        qr_cell = [P(d.get("no_qr_note", ""), S["muted_c"])]
    bottom = Table([[qr_cell, "", tot]], colWidths=[150, W - 450, 300])
    bottom.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (0, 0), "CENTER"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    el += [bottom, Spacer(1, 8)]

    # 6) التفقيط
    words = Table([[P(f"المبلغ كتابةً: {tafqeet(d['grand'])}", S["words"])]], colWidths=[W])
    words.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#ecfdf5")),
                               ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#99f6e4")),
                               ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 8)]))
    el += [words, Spacer(1, 8)]

    # 7) البيانات البنكية والملاحظات
    if d.get("show_bank", True) and comp["iban"]:
        bank = _kv_table([("البنك | Bank", comp["bank_name"]), ("رقم الآيبان | IBAN", format_iban(comp["iban"])),
                          ("اسم المستفيد | Beneficiary", comp["company_name"]),
                          ("Beneficiary Name", comp["company_name_en"])],
                         S, W - 12)
        el += [_box("البيانات البنكية | Bank Details", bank, S, W), Spacer(1, 8)]
    if d.get("notes"):
        el += [P(f"ملاحظات: {d['notes']}", S["muted_r"]), Spacer(1, 6)]

    # 8) التوقيعات
    sigs = d.get("signatures") or ["المستلم | Received by", "المحاسب | Accountant", "الختم والتوقيع | Stamp & Signature"]
    sig = Table([[P(x, S["c"]) for x in rtl(sigs)], [P("........................................", S["c"])] * len(sigs)],
                colWidths=[W / len(sigs)] * len(sigs))
    sig.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 10)]))
    el += [Spacer(1, 6), sig]

    footer_note = d.get("footer_note", "")

    watermark = _watermark_image()

    def _footer(canvas, doc_):
        canvas.saveState()
        if watermark is not None:
            # منتصف الصفحة تماماً، بعرض ≈ 80% من الصفحة مع الحفاظ على نسبة الشعار
            iw, ih = watermark.getSize()
            size_w = A4[0] * 0.8
            size_h = size_w * ih / iw
            if size_h > A4[1] * 0.6:
                size_h = A4[1] * 0.6
                size_w = size_h * iw / ih
            canvas.drawImage(watermark, (A4[0] - size_w) / 2, (A4[1] - size_h) / 2, size_w, size_h, mask="auto")
        canvas.setStrokeColor(colors.HexColor(DOC_LINE))
        canvas.setLineWidth(0.6)
        canvas.line(30, 30, A4[0] - 30, 30)
        canvas.setFont(FONT_NAME, 7)
        canvas.setFillColor(colors.HexColor("#64748b"))
        canvas.drawString(30, 19, f"Page {doc_.page}  |  {d['title_en']} {d['number']}")
        canvas.drawRightString(A4[0] - 30, 19, ar(footer_note or f"صادر بواسطة {BRAND_FULL}"))
        canvas.restoreState()

    doc.build(el, onFirstPage=_footer, onLaterPages=_footer)


def _zatca_ts(d, created):
    t = (created or "")[11:19] or "00:00:00"
    return f"{d}T{t}Z", t


def build_invoice_pdf(path, kind, inv_id):
    """فاتورة ضريبية (مبيعات) أو فاتورة مشتريات بتصميم احترافي ثنائي اللغة."""
    comp = company_info()
    c = FLOW[kind]
    sales = kind == "sales"
    with db() as conn:
        h = conn.execute(
            f"SELECT COALESCE(invoice_no, {c['iid']}), {c['idate']}, {c['pcode']}, {c['pname']}, subtotal, vat_amount, "
            f"total_amount, paid_amount, COALESCE(notes,''), created_at, "
            + ("COALESCE(sales_rep,''), ''" if sales else "COALESCE(purchase_rep,''), COALESCE(supplier_ref,'')")
            + f" FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
        if not h:
            raise UserError("الفاتورة غير موجودة.")
        if sales:
            lines = conn.execute("SELECT item_code, item_name, quantity, unit, unit_price, COALESCE(discount_amount,0), "
                                 "subtotal, vat_amount, total_amount FROM sales_invoice_lines WHERE invoice_id=? "
                                 "ORDER BY line_id", (inv_id,)).fetchall()
        else:
            lines = conn.execute("SELECT item_code, item_name, quantity, unit, unit_cost, 0, subtotal, vat_amount, "
                                 "total_amount FROM purchase_invoice_lines WHERE purchase_id=? ORDER BY line_id",
                                 (inv_id,)).fetchall()
        party = conn.execute("SELECT name, COALESCE(vat_number,''), COALESCE(cr_number,''), COALESCE(address,''), "
                             "COALESCE(phone,'') FROM business_partners WHERE partner_code=?", (h[2],)).fetchone()
        post = conn.execute("SELECT COALESCE(SUM(total_amount),0) FROM post_discounts WHERE kind=? AND invoice_id=?",
                            (kind, inv_id)).fetchone()[0]
    p_name, p_vat, p_cr, p_addr, p_phone = party or (h[3], "", "", "", "")
    ts, t = _zatca_ts(h[1], h[9])
    settled = h[7] + post
    status = "مسددة | Paid" if settled >= h[6] - 0.005 else ("آجلة | Credit" if h[7] < 0.005 else "مسددة جزئياً | Partial")
    meta = [("رقم الفاتورة | Invoice No", h[0]), ("تاريخ الإصدار | Issue Date", h[1]), ("الوقت | Time", t),
            ("حالة السداد | Payment", status)]
    if h[10]:
        meta.append(("المندوب | Sales Rep" if sales else "المسؤول | Officer", h[10]))
    if not sales and h[11]:
        meta.append(("رقم فاتورة المورد | Supplier Ref", h[11]))
    party_rows = [("الاسم | Name", f"{p_name} ({h[2]})"), ("الرقم الضريبي | VAT No", p_vat or "غير مسجل"),
                  ("السجل التجاري | C.R No", p_cr or "-"), ("العنوان | Address", p_addr or "-")]
    if p_phone:
        party_rows.append(("الهاتف | Phone", p_phone))
    if sales:
        columns = [("م\n#", 20), ("الكود\nCode", 46), ("الصنف\nDescription", 120), ("الكمية\nQty", 40),
                   ("الوحدة\nUnit", 34), ("سعر الوحدة\nUnit Price", 50), ("الخصم\nDiscount", 42),
                   ("الخاضع للضريبة\nTaxable", 60), ("الضريبة 15%\nVAT", 55), ("الإجمالي\nTotal", 68)]
    else:
        columns = [("م\n#", 20), ("الكود\nCode", 50), ("الصنف\nDescription", 158), ("الكمية\nQty", 42),
                   ("الوحدة\nUnit", 36), ("التكلفة\nUnit Cost", 52), ("قبل الضريبة\nNet", 60),
                   ("الضريبة 15%\nVAT", 54), ("الإجمالي\nTotal", 63)]
    rows, gross, disc = [], 0.0, 0.0
    for i, (code, name, qty, unit, price, dsc, sub, vat, tot) in enumerate(lines, start=1):
        gross += sub + dsc
        disc += dsc
        r = [str(i), code, name, f"{qty:g}", unit or "-", fmt(price)]
        if sales:
            r.append(fmt(dsc) if dsc > 0.005 else "-")
        rows.append(r + [fmt(sub), fmt(vat), fmt(tot)])
    totals = []
    if disc > 0.005:
        totals += [("الإجمالي قبل الخصم | Gross", gross, False), ("(-) الخصم | Discount", disc, False)]
    totals += [("الإجمالي الخاضع للضريبة | Taxable Amount", h[4], False),
               ("ضريبة القيمة المضافة 15% | VAT 15%", h[5], False),
               ("الإجمالي شامل الضريبة | Total incl. VAT", h[6], True)]
    if post > 0.005:
        totals += [("(-) إشعارات دائنة / خصومات لاحقة | Credit Notes", post, False),
                   ("صافي المستحق | Net Due", h[6] - post, False)]
    if h[7] > 0.005:
        totals += [("المسدد | Paid", h[7], False), ("المتبقي | Balance Due", max(h[6] - post - h[7], 0.0), False)]
    _render_document(path, dict(
        title_ar="فاتورة ضريبية" if sales else "فاتورة مشتريات", title_en="TAX INVOICE" if sales else "PURCHASE INVOICE",
        number=h[0], meta=meta, party_title="بيانات العميل | Customer" if sales else "بيانات المورد | Supplier",
        party=party_rows, columns=columns, rows=rows, totals=totals, grand=h[6], notes=h[8],
        qr=zatca_qr_base64(comp["company_name"], comp["vat_number"], ts, h[6], h[5]) if sales else None,
        no_qr_note="" if sales else "مستند داخلي لقيد فاتورة المورد\nInternal purchase record",
        show_bank=sales,
        signatures=None if sales else ["أمين المخزن | Storekeeper", "المحاسب | Accountant", "المدير | Manager"]))


def build_sales_invoice_pdf(path, invoice_id):
    build_invoice_pdf(path, "sales", invoice_id)


def build_purchase_invoice_pdf(path, purchase_id):
    build_invoice_pdf(path, "purchase", purchase_id)


def build_post_discount_pdf(path, discount_id):
    """إشعار دائن ضريبي للعميل (برمز QR) أو إشعار خصم من المورد (مستند داخلي) بنفس التصميم."""
    d, lines = load_post_discount(discount_id)
    comp = company_info()
    sales = d["kind"] == DISC_SALES
    with db() as conn:
        party = conn.execute("SELECT COALESCE(vat_number,''), COALESCE(cr_number,''), COALESCE(address,'') "
                             "FROM business_partners WHERE partner_code=?", (d["party_code"],)).fetchone() or ("", "", "")
    ts, t = _zatca_ts(d["discount_date"], d["created_at"])
    meta = [("رقم الإشعار | Note No", d["discount_no"]), ("تاريخ الإصدار | Issue Date", d["discount_date"]),
            ("الوقت | Time", t), ("الفاتورة الأصلية | Original Invoice", d["invoice_no"] or "على الحساب | On account"),
            ("السبب | Reason", d["reason"] or "-")]
    if not sales:
        meta.append(("المعالجة | Treatment", DISC_METHOD_LABEL[d["method"]]))
    party_rows = [("الاسم | Name", f"{d['party_name']} ({d['party_code']})"),
                  ("الرقم الضريبي | VAT No", party[0] or "غير مسجل"), ("السجل التجاري | C.R No", party[1] or "-"),
                  ("العنوان | Address", party[2] or "-")]
    vat_rate = (d["vat_amount"] / d["subtotal"]) if d["subtotal"] else 0
    if lines:
        columns = [("م\n#", 24), ("الكود\nCode", 60), ("الصنف\nDescription", 211), ("قبل الضريبة\nNet", 80),
                   ("الضريبة\nVAT", 75), ("الإجمالي\nTotal", 85)]
        rows = [[str(i), code, name, fmt(a), fmt(a * vat_rate), fmt(a * (1 + vat_rate))]
                for i, (code, name, a, _ap) in enumerate(lines, start=1)]
    else:
        columns = [("م\n#", 24), ("البيان\nDescription", 271), ("قبل الضريبة\nNet", 80), ("الضريبة\nVAT", 75),
                   ("الإجمالي\nTotal", 85)]
        desc = ("خصم / تسوية لاحقة" + (f" على الفاتورة {d['invoice_no']}" if d["invoice_no"] else "")
                + (f" - {d['reason']}" if d["reason"] else ""))
        rows = [["1", desc, fmt(d["subtotal"]), fmt(d["vat_amount"]), fmt(d["total_amount"])]]
    totals = [("قيمة الخصم قبل الضريبة | Net Amount", d["subtotal"], False),
              ("ضريبة القيمة المضافة | VAT", d["vat_amount"], False),
              ("إجمالي الإشعار | Total Credit", d["total_amount"], True)]
    _render_document(path, dict(
        title_ar="إشعار دائن ضريبي" if sales else "إشعار خصم من مورد",
        title_en="TAX CREDIT NOTE" if sales else "SUPPLIER DISCOUNT NOTE",
        number=d["discount_no"], meta=meta, party_title="بيانات العميل | Customer" if sales else "بيانات المورد | Supplier",
        party=party_rows, columns=columns, rows=rows, totals=totals, grand=d["total_amount"], notes=d["notes"],
        desc_col=1 if not lines else 2,
        qr=zatca_qr_base64(comp["company_name"], comp["vat_number"], ts, d["total_amount"], d["vat_amount"]) if sales else None,
        no_qr_note="" if sales else "مستند داخلي لقيد خصم المورد\nInternal supplier discount record",
        show_bank=False,
        signatures=["المحاسب | Accountant", "المدير المالي | Finance Manager", "الختم | Stamp"]))


def _simple_report_pdf(path, title_lines, headers, rows, col_widths, footer_lines=()):
    S = pdf_styles()
    with db() as conn:
        comp = conn.execute("SELECT company_name FROM company_profile WHERE id=1").fetchone()
    doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    el = [P(comp[0], S['title']), Spacer(1, 8)]
    for ln in title_lines:
        el += [P(ln, S['right']), Spacer(1, 3)]
    el.append(Spacer(1, 6))
    data = [rtl([P(h, S['head']) for h in headers])]
    for r in rows:
        data.append(rtl([P(v, S['cell']) for v in r]))
    t = Table(data, colWidths=rtl(col_widths), repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
                           ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                           ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1"))]))
    el += [t, Spacer(1, 12)]
    for ln in footer_lines:
        el += [P(ln, S['right']), Spacer(1, 3)]
    doc.build(el)


def build_statement_pdf(path, partner_code, scope="both", d_from=None, d_to=None):
    info, rows, tot_d, tot_c, bal = fetch_partner_statement(partner_code, scope=scope, d_from=d_from, d_to=d_to)
    side = "مدين" if bal > 0.005 else ("دائن" if bal < -0.005 else "")
    codes = linked_partner_codes(partner_code)
    _simple_report_pdf(
        path,
        [f"كشف حساب موحد: {info[0]} ({' + '.join(codes)})",
         f"العرض: {STATEMENT_SCOPES.get(scope, scope)}   |   {period_label(d_from, d_to)}",
         f"هاتف: {info[1] or '-'}   |   الرقم الضريبي: {info[3] or '-'}   |   تاريخ الكشف: {today()}"],
        ["المرجع", "التاريخ", "النوع", "البيان", "مدين", "دائن", "الرصيد"], rows,
        [60, 60, 75, 140, 65, 65, 90],
        [f"إجمالي المدين: {fmt(tot_d)}   |   إجمالي الدائن: {fmt(tot_c)}   |   "
         f"الرصيد النهائي: {abs(bal):,.2f} {side}"])


def build_balances_summary_pdf(path, scope="both", d_from=None, d_to=None, q="", hide_zero=False):
    """ملخص أرصدة العملاء/الموردين: سطر واحد لكل طرف."""
    res = partners_balance_summary(scope, d_from, d_to, q, hide_zero)
    t = res["totals"]
    rows = [r[:3] + r[3:7] for r in summary_table_rows(res, bool(d_from))]
    _simple_report_pdf(
        path,
        [f"ملخص أرصدة {SUMMARY_SCOPES.get(scope, scope)}", period_label(d_from, d_to),
         f"عدد الأطراف: {len(res['rows'])}   |   تاريخ الطباعة: {today()}"],
        SUMMARY_HEADERS[:7], rows, [55, 140, 70, 75, 70, 70, 85],
        [f"إجمالي المدين (لنا): {fmt(t['receivable'])}   |   إجمالي الدائن (علينا): {fmt(t['payable'])}   |   "
         f"الصافي: {balance_text(t['net'])}"])


def build_item_card_pdf(path, item_code, item_name, rows, final_text):
    _simple_report_pdf(
        path,
        [f"كارت صنف / Item Movement Report: {item_name} ({item_code})"],
        ["التاريخ", "نوع الحركة", "الطرف", "الكمية", "الرصيد", "التكلفة", "القيمة بالتكلفة"],
        [r[:7] for r in rows], [62, 100, 110, 55, 65, 60, 70], [final_text])


# ============================================================
# دورة المستندات: أمر ← تسليم/استلام ← فاتورة
# ============================================================
OPEN_BLOCKED = ("مفوتر", "مغلق", "ملغي")   # حالات لا يمكن تنفيذ الأمر فيها

FLOW = {
    "sales": dict(
        order="sales_orders", olines="sales_order_lines", ful="sales_deliveries", flines="sales_delivery_lines",
        fid="delivery_id", done="delivered_qty", pcode="customer_code", pname="customer_name", rep="sales_rep",
        oprice="unit_price", fdate="delivery_date", fprice="unit_price", inv="sales_invoices",
        ilines="sales_invoice_lines", iid="invoice_id", link="delivery_line_id", idate="invoice_date",
        partial="مسلّم جزئياً", full="مسلّم", src="تسليم",
        fval="(l.quantity*l.unit_price - COALESCE(l.discount_amount,0))", fdisc="COALESCE(discount_amount,0)",
        onet="(l.unit_price - COALESCE(l.discount_amount,0) / l.quantity)"),
    "purchase": dict(
        order="purchase_orders", olines="purchase_order_lines", ful="purchase_receipts", flines="purchase_receipt_lines",
        fid="receipt_id", done="received_qty", pcode="supplier_code", pname="supplier_name", rep="purchase_rep",
        oprice="unit_cost", fdate="receipt_date", fprice="unit_cost", inv="purchase_invoices",
        ilines="purchase_invoice_lines", iid="purchase_id", link="receipt_line_id", idate="purchase_date",
        partial="مستلم جزئياً", full="مستلم", src="استلام",
        fval="(l.quantity*l.unit_cost)", fdisc="0", onet="l.unit_cost"),
}


def stock_add(conn, code, qty, unit_cost):
    """إضافة للمخزن مع تحديث متوسط التكلفة المرجح."""
    row = conn.execute("SELECT quantity, cost_price FROM inventory WHERE item_code=?", (code,)).fetchone()
    if row is None:
        raise UserError(f"الصنف {code} غير موجود بالمخزن.")
    old_q, old_c = max(row[0] or 0.0, 0.0), row[1] or 0.0
    new_q = old_q + qty
    new_c = (old_q * old_c + qty * unit_cost) / new_q if new_q > 0 else unit_cost
    conn.execute("UPDATE inventory SET quantity = quantity + ?, cost_price = ? WHERE item_code=?",
                 (qty, round(new_c, 4), code))


def stock_remove(conn, code, qty):
    """صرف من المخزن ويرجع متوسط التكلفة وقت الصرف."""
    row = conn.execute("SELECT item_name, quantity, cost_price FROM inventory WHERE item_code=?", (code,)).fetchone()
    if row is None:
        raise UserError(f"الصنف {code} غير موجود بالمخزن.")
    if (row[1] or 0.0) + EPS < qty:
        raise UserError(f"الكمية غير كافية للصنف {row[0]}. المتاح: {row[1]:g}")
    conn.execute("UPDATE inventory SET quantity = quantity - ? WHERE item_code=?", (qty, code))
    return row[2] or 0.0


def refresh_order_status(conn, kind, order_id):
    c = FLOW[kind]
    cur = conn.execute(f"SELECT status FROM {c['order']} WHERE order_id=?", (order_id,)).fetchone()
    if cur is None or cur[0] in ("ملغي", "مغلق"):
        return
    q, d, i = conn.execute(f"SELECT COALESCE(SUM(quantity),0), COALESCE(SUM({c['done']}),0), "
                           f"COALESCE(SUM(invoiced_qty),0) FROM {c['olines']} WHERE order_id=?", (order_id,)).fetchone()
    if d <= EPS:
        status = "جديد"
    elif d + EPS < q:
        status = c["partial"]
    else:
        status = c["full"] if i + EPS < q else "مفوتر"
    conn.execute(f"UPDATE {c['order']} SET status=? WHERE order_id=?", (status, order_id))


def create_order(kind, o_date, party, notes, lines, expected=None, rep="", allow_over_credit=False):
    require_perm(f"{kind}.orders_create", "إنشاء أمر " + ("بيع" if kind == "sales" else "شراء"))
    """party = (code, name). lines = قائمة make_line. لا يؤثر على المخزون.
    أمر البيع يخضع للحد الائتماني للعميل (سياسة «منع» ترفضه عند التجاوز إلا بـ allow_over_credit)."""
    c = FLOW[kind]
    if not lines:
        raise UserError("أضف صنفاً واحداً على الأقل للأمر.")
    sub = round(sum(l["subtotal"] for l in lines), 2)
    vat = round(sum(l["vat"] for l in lines), 2)
    with db() as conn:
        if kind == "sales":
            from .credit import check_credit
            check_credit(conn, party[0], round(sub + vat, 2), "order", allow_over_credit)
        rep = ensure_rep(conn, rep)
        cur = conn.execute(f"INSERT INTO {c['order']} (order_date, expected_date, {c['pcode']}, {c['pname']}, {c['rep']}, "
                           f"subtotal, vat_amount, total_amount, notes) VALUES (?,?,?,?,?,?,?,?,?)",
                           (o_date, expected, party[0], party[1], rep, sub, vat, round(sub + vat, 2), notes))
        oid = cur.lastrowid
        for l in lines:
            if kind == "sales":
                conn.execute("INSERT INTO sales_order_lines (order_id, item_code, item_name, unit, quantity, unit_price, "
                             "discount_amount, subtotal, vat_amount, total_amount) VALUES (?,?,?,?,?,?,?,?,?,?)",
                             (oid, l["code"], l["name"], l["unit"], l["qty"], l["price"], l.get("discount", 0.0),
                              l["subtotal"], l["vat"], l["total"]))
            else:
                conn.execute(f"INSERT INTO {c['olines']} (order_id, item_code, item_name, unit, quantity, {c['oprice']}, "
                             f"subtotal, vat_amount, total_amount) VALUES (?,?,?,?,?,?,?,?,?)",
                             (oid, l["code"], l["name"], l["unit"], l["qty"], l["price"], l["subtotal"], l["vat"], l["total"]))
    return oid


def close_order(kind, order_id):
    c = FLOW[kind]
    with db() as conn:
        st = conn.execute(f"SELECT status FROM {c['order']} WHERE order_id=?", (order_id,)).fetchone()
        if st is None:
            raise UserError("الأمر غير موجود.")
        if st[0] in OPEN_BLOCKED:
            raise UserError(f"الأمر حالته بالفعل: {st[0]}")
        conn.execute(f"UPDATE {c['order']} SET status='مغلق' WHERE order_id=?", (order_id,))


def cancel_order(kind, order_id):
    c = FLOW[kind]
    with db() as conn:
        st = conn.execute(f"SELECT status FROM {c['order']} WHERE order_id=?", (order_id,)).fetchone()
        if st is None:
            raise UserError("الأمر غير موجود.")
        done = conn.execute(f"SELECT COALESCE(SUM({c['done']}),0) FROM {c['olines']} WHERE order_id=?",
                            (order_id,)).fetchone()[0]
        if done > EPS:
            raise UserError("لا يمكن إلغاء أمر نُفّذ منه جزء. استخدم «إغلاق الأمر» لإنهاء الكمية المتبقية.")
        if st[0] == "ملغي":
            raise UserError("الأمر ملغي بالفعل.")
        conn.execute(f"UPDATE {c['order']} SET status='ملغي' WHERE order_id=?", (order_id,))


def fulfil_order(kind, order_id, f_date, notes, entries, supplier_ref="", allow_over_credit=False):
    """تسليم (مبيعات: يخصم من المخزن) أو استلام (مشتريات: يضيف للمخزن).
    entries = [(order_line_id, qty, price_override_or_None)]. يدعم التنفيذ الجزئي على أكثر من دفعة."""
    c = FLOW[kind]
    require_perm(f"{kind}.{'delivery' if kind == 'sales' else 'receipt'}_create",
                 "تنفيذ " + ("إذن تسليم" if kind == "sales" else "استلام المخزن"))
    entries = [(e[0], round(e[1], 3), e[2], (e[3] if len(e) > 3 else None) or None, (e[4] if len(e) > 4 else None) or None)
               for e in entries if e[1] is not None and e[1] > EPS]
    if not entries:
        raise UserError("أدخل كمية واحدة على الأقل.")
    with db() as conn:
        o = conn.execute(f"SELECT {c['pcode']}, {c['pname']}, status FROM {c['order']} WHERE order_id=?",
                         (order_id,)).fetchone()
        if o is None:
            raise UserError("الأمر غير موجود.")
        if o[2] in OPEN_BLOCKED:
            raise UserError(f"لا يمكن التنفيذ على أمر حالته: {o[2]}")
        if kind == "sales":
            cur = conn.execute("INSERT INTO sales_deliveries (delivery_date, order_id, customer_code, customer_name, notes) "
                               "VALUES (?,?,?,?,?)", (f_date, order_id, o[0], o[1], notes))
        else:
            cur = conn.execute("INSERT INTO purchase_receipts (receipt_date, order_id, supplier_code, supplier_name, "
                               "supplier_ref, notes) VALUES (?,?,?,?,?,?)", (f_date, order_id, o[0], o[1], supplier_ref, notes))
        fid = cur.lastrowid
        for lid, qty, price_override, prod_date, exp_date in entries:
            ol = conn.execute(f"SELECT item_code, item_name, unit, quantity, {c['oprice']}, {c['done']} FROM {c['olines']} "
                              f"WHERE line_id=? AND order_id=?", (lid, order_id)).fetchone()
            if ol is None:
                raise UserError("بند غير تابع لهذا الأمر.")
            remaining = ol[3] - ol[5]
            if qty > remaining + EPS:
                raise UserError(f"الكمية أكبر من المتبقي للصنف {ol[1]} (المتبقي {remaining:g}).")
            if kind == "sales":
                cost = stock_remove(conn, ol[0], qty)
                o_disc = conn.execute("SELECT COALESCE(discount_amount,0) FROM sales_order_lines WHERE line_id=?",
                                      (lid,)).fetchone()[0]
                done_disc = conn.execute("SELECT COALESCE(SUM(discount_amount),0) FROM sales_delivery_lines "
                                         "WHERE order_line_id=?", (lid,)).fetchone()[0]
                disc = (o_disc - done_disc) if qty >= remaining - EPS else o_disc * qty / ol[3]
                disc = max(round(disc, 2), 0.0)
                conn.execute("INSERT INTO sales_delivery_lines (delivery_id, order_line_id, item_code, item_name, unit, "
                             "quantity, unit_price, unit_cost, discount_amount) VALUES (?,?,?,?,?,?,?,?,?)",
                             (fid, lid, ol[0], ol[1], ol[2], qty, ol[4], cost, disc))
            else:
                # سعر الاستلام = سعر أمر الشراء المعتمد دائماً (لا تعديل للسعر في إذن الاستلام)
                price = ol[4]
                if price < 0:
                    raise UserError("السعر لا يمكن أن يكون سالباً.")
                validate_batch_dates(prod_date, exp_date, ol[1])
                stock_add(conn, ol[0], qty, price)
                conn.execute("INSERT INTO purchase_receipt_lines (receipt_id, order_line_id, item_code, item_name, unit, "
                             "quantity, unit_cost, production_date, expiry_date) VALUES (?,?,?,?,?,?,?,?,?)",
                             (fid, lid, ol[0], ol[1], ol[2], qty, price, prod_date, exp_date))
            conn.execute(f"UPDATE {c['olines']} SET {c['done']} = {c['done']} + ? WHERE line_id=?", (qty, lid))
        refresh_order_status(conn, kind, order_id)
        if kind == "sales":   # التسليم = البيع الآجل الفعلي: الإذن نفسه محسوب ضمن «تسليمات لم تُفوتر»
            from .credit import check_credit
            check_credit(conn, o[0], 0.0, "delivery", allow_over_credit)
    return fid


def invoice_from_fulfilments(kind, partner_code, ful_ids, inv_date, notes="", supplier_ref=""):
    """إصدار فاتورة من تسليمات/استلامات غير مفوترة لنفس الشريك. لا يمس المخزن (تحرك وقت التنفيذ).
    يرجع (رقم الفاتورة، الإجمالي)."""
    c = FLOW[kind]
    ful_ids = list(dict.fromkeys(ful_ids))
    if not ful_ids:
        raise UserError("اختر مستنداً واحداً على الأقل.")
    ph = ",".join("?" * len(ful_ids))
    with db() as conn:
        heads = conn.execute(f"SELECT {c['fid']}, {c['pcode']}, {c['pname']}, invoice_id, order_id FROM {c['ful']} "
                             f"WHERE {c['fid']} IN ({ph})", ful_ids).fetchall()
        if len(heads) != len(ful_ids):
            raise UserError("بعض المستندات المحددة غير موجودة.")
        for h in heads:
            if h[1] != partner_code:
                raise UserError("المستندات المحددة تخص أكثر من شريك.")
            if h[3] is not None:
                raise UserError(f"المستند رقم {h[0]} مفوتر بالفعل (فاتورة {h[3]}).")
        party_name = heads[0][2]
        order_ids = sorted({h[4] for h in heads})
        reps = [r[0] for r in conn.execute(
            f"SELECT DISTINCT {c['rep']} FROM {c['order']} WHERE order_id IN ({','.join('?' * len(order_ids))}) "
            f"AND {c['rep']} IS NOT NULL AND {c['rep']} != ''", order_ids)]
        rep = "، ".join(reps)
        if kind == "sales":
            raw = conn.execute(f"SELECT line_id, order_line_id, item_code, item_name, unit, quantity, unit_price, unit_cost, "
                               f"COALESCE(discount_amount,0), NULL, NULL FROM sales_delivery_lines WHERE delivery_id IN ({ph}) ORDER BY delivery_id, line_id",
                               ful_ids).fetchall()
        else:
            raw = conn.execute(f"SELECT line_id, order_line_id, item_code, item_name, unit, quantity, unit_cost, unit_cost, 0, "
                               f"production_date, expiry_date FROM purchase_receipt_lines WHERE receipt_id IN ({ph}) ORDER BY receipt_id, line_id",
                               ful_ids).fetchall()
        lines = []
        for lid, olid, code, name, unit, qty, price, cost, disc, prod_date, exp_date in raw:
            L = make_line(code, name, unit, qty, price, disc)
            L.update(link=lid, olid=olid, cost=cost, production=prod_date, expiry=exp_date)
            lines.append(L)
        if not lines:
            raise UserError("المستندات المحددة لا تحتوي على بنود.")
        subtotal = round(sum(l["subtotal"] for l in lines), 2)
        vat = round(sum(l["vat"] for l in lines), 2)
        total = round(subtotal + vat, 2)
        discount = round(sum(l["discount"] for l in lines), 2)
        order_ref = order_ids[0] if len(order_ids) == 1 else None
        if kind == "sales":
            cur = conn.execute("INSERT INTO sales_invoices (order_id, invoice_date, customer_code, customer_name, sales_rep, "
                               "subtotal, vat_amount, total_amount, notes, source_type, discount_amount, invoice_no) "
                               "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                               (order_ref, inv_date, partner_code, party_name, rep, subtotal, vat, total, notes, c["src"],
                                discount, next_serial(conn, "sales_invoice")))
        else:
            cur = conn.execute("INSERT INTO purchase_invoices (order_id, purchase_date, supplier_code, supplier_name, "
                               "purchase_rep, subtotal, vat_amount, total_amount, notes, source_type, supplier_ref, invoice_no) "
                               "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                               (order_ref, inv_date, partner_code, party_name, rep, subtotal, vat, total, notes,
                                c["src"], supplier_ref, next_serial(conn, "purchase_invoice")))
        iid = cur.lastrowid
        for L in lines:
            if kind == "sales":
                conn.execute("INSERT INTO sales_invoice_lines (invoice_id, item_code, item_name, unit, quantity, unit_price, "
                             "unit_cost, discount_amount, subtotal, vat_amount, total_amount, delivery_line_id) "
                             "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                             (iid, L["code"], L["name"], L["unit"], L["qty"], L["price"], L["cost"], L["discount"],
                              L["subtotal"], L["vat"], L["total"], L["link"]))
            else:
                conn.execute("INSERT INTO purchase_invoice_lines (purchase_id, item_code, item_name, unit, quantity, unit_cost, "
                             "subtotal, vat_amount, total_amount, receipt_line_id, production_date, expiry_date) "
                             "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                             (iid, L["code"], L["name"], L["unit"], L["qty"], L["price"],
                              L["subtotal"], L["vat"], L["total"], L["link"], L["production"], L["expiry"]))
            conn.execute(f"UPDATE {c['olines']} SET invoiced_qty = invoiced_qty + ? WHERE line_id=?", (L["qty"], L["olid"]))
        conn.execute(f"UPDATE {c['ful']} SET invoice_id=? WHERE {c['fid']} IN ({ph})", [iid] + ful_ids)
        for oid in order_ids:
            refresh_order_status(conn, kind, oid)
    return iid, total


def build_order_pdf(path, kind, oid):
    """أمر بيع (للمندوب) أو أمر شراء (للمورد)."""
    c = FLOW[kind]
    title = "أمر بيع" if kind == "sales" else "أمر شراء"
    party_label = "العميل" if kind == "sales" else "المورد"
    with db() as conn:
        h = conn.execute(f"SELECT order_date, expected_date, {c['pname']}, {c['rep']}, total_amount, notes, status "
                         f"FROM {c['order']} WHERE order_id=?", (oid,)).fetchone()
        if h is None:
            raise UserError("الأمر غير موجود.")
        disc_col = "COALESCE(discount_amount,0)" if kind == "sales" else "0"
        lines = conn.execute(f"SELECT item_name, quantity, unit, {c['oprice']}, {disc_col}, total_amount FROM {c['olines']} "
                             f"WHERE order_id=? ORDER BY line_id", (oid,)).fetchall()
    head = [f"{title} رقم {oid}   |   التاريخ: {h[0]}   |   الحالة: {h[6]}", f"{party_label}: {h[2]}"]
    if h[3]:
        head.append(("المندوب: " if kind == "sales" else "المسؤول: ") + h[3])
    if h[1]:
        head.append(f"تاريخ التسليم المتوقع: {h[1]}")
    if h[5]:
        head.append(f"ملاحظات: {h[5]}")
    if kind == "sales":
        rows = [(str(i), n, f"{q:g}", u or "-", fmt(p), fmt(d) if d > 0.005 else "-", fmt(t))
                for i, (n, q, u, p, d, t) in enumerate(lines, 1)]
        headers, widths = ["م", "الصنف", "الكمية", "الوحدة", "السعر", "الخصم الخاص", "الإجمالي شامل الضريبة"], \
            [28, 190, 50, 45, 65, 65, 97]
    else:
        rows = [(str(i), n, f"{q:g}", u or "-", fmt(p), fmt(t)) for i, (n, q, u, p, _, t) in enumerate(lines, 1)]
        headers, widths = ["م", "الصنف", "الكمية", "الوحدة", "السعر", "الإجمالي شامل الضريبة"], [30, 215, 55, 50, 70, 110]
    _simple_report_pdf(path, head, headers, rows, widths, [f"إجمالي الأمر شامل الضريبة: {fmt(h[4])} ر.س"])


def build_fulfilment_pdf(path, kind, fid):
    """إذن تسليم (مبيعات) أو إذن استلام (مشتريات) بدون أسعار."""
    c = FLOW[kind]
    title = "إذن تسليم بضاعة" if kind == "sales" else "إذن استلام بضاعة"
    with db() as conn:
        h = conn.execute(f"SELECT {c['fdate']}, order_id, {c['pname']}, notes FROM {c['ful']} WHERE {c['fid']}=?", (fid,)).fetchone()
        if h is None:
            raise UserError("المستند غير موجود.")
        ref = ""
        if kind == "purchase":
            ref = conn.execute("SELECT supplier_ref FROM purchase_receipts WHERE receipt_id=?", (fid,)).fetchone()[0] or ""
        lines = conn.execute(f"SELECT item_name, unit, quantity FROM {c['flines']} WHERE {c['fid']}=? ORDER BY line_id", (fid,)).fetchall()
    party_label = "العميل" if kind == "sales" else "المورد"
    head = [f"{title} رقم {fid}   |   التاريخ: {h[0]}   |   على أمر رقم {h[1]}", f"{party_label}: {h[2]}"]
    if ref:
        head.append(f"مرجع المورد: {ref}")
    if h[3]:
        head.append(f"ملاحظات: {h[3]}")
    rows = [(str(i), name, unit or "-", f"{qty:g}") for i, (name, unit, qty) in enumerate(lines, 1)]
    sign = "المستلم: ........................     التوقيع: ........................     التاريخ: ................"
    _simple_report_pdf(path, head, ["م", "الصنف", "الوحدة", "الكمية"], rows, [35, 330, 70, 90], [sign])


def period_is_closed(d):
    with db() as conn:
        return conn.execute("SELECT 1 FROM daily_closures WHERE closure_date=?", (d,)).fetchone() is not None


# ============================================================
# استيراد من الإكسيل (شركاء + أصناف)
# ============================================================
_NAME_NORM = str.maketrans({'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ى': 'ي', 'ة': 'ه', 'ؤ': 'و', 'ئ': 'ي', 'ـ': ''})


def clean_name(s):
    return re.sub(r'\s+', ' ', str(s).replace('\u200f', '').replace('\u200e', '')).strip()


def norm_name(s):
    """مفتاح مطابقة الأسماء: يتجاهل المسافات الزائدة وهمزات الألف والتاء المربوطة وحجم الحروف."""
    return clean_name(s).translate(_NAME_NORM).casefold()


def _find_col(df, candidates, fallback_idx=None):
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    if fallback_idx is not None and df.shape[1] > fallback_idx:
        return df.columns[fallback_idx]
    return None


def _cell(row, col):
    if col is None:
        return ""
    v = row[col]
    return "" if pd.isna(v) else str(v).strip()


def import_partners_from_df(df, role, merge_roles=True, dry_run=False):
    """role = 'customer' | 'supplier'. الأسماء المتطابقة تُدمج في شريك واحد (عميل ومورد معاً) إذا merge_roles."""
    name_col = _find_col(df, ['Name Contact', 'Name', 'الاسم', 'اسم', 'اسم العميل', 'اسم المورد'],
                         1 if df.shape[1] > 1 else 0)
    phone_col = _find_col(df, ['الهاتف', 'Phone', 'Mobile', 'جوال', 'تليفون'])
    vat_col = _find_col(df, ['الرقم الضريبي', 'VAT', 'VAT Number', 'VAT No'])
    cr_col = _find_col(df, ['السجل التجاري', 'CR', 'CR Number'])
    addr_col = _find_col(df, ['العنوان', 'Address'])
    code_col = _find_col(df, PARTNER_CODE_COLS)
    flag = "is_customer" if role == "customer" else "is_supplier"
    stats = dict(added=0, merged=0, existing_same_role=0, skipped_other_role=0, duplicates=0, empty=0)
    with db() as conn:
        existing, used_codes = {}, set()
        for code, name, ic, is_ in conn.execute("SELECT partner_code, name, is_customer, is_supplier FROM business_partners"):
            existing[norm_name(name)] = (code, ic, is_)
            used_codes.add(code)
        seen = set()
        for _, row in df.iterrows():
            name = clean_name(_cell(row, name_col))
            if not name:
                stats["empty"] += 1
                continue
            key = norm_name(name)
            if key in seen:
                stats["duplicates"] += 1
                continue
            seen.add(key)
            if key in existing:
                code, ic, is_ = existing[key]
                if (ic if role == "customer" else is_):
                    stats["existing_same_role"] += 1
                elif merge_roles:
                    stats["merged"] += 1
                    if not dry_run:
                        conn.execute(f"UPDATE business_partners SET {flag}=1 WHERE partner_code=?", (code,))
                else:
                    stats["skipped_other_role"] += 1
                continue
            stats["added"] += 1
            if dry_run:
                continue
            code = _cell(row, code_col)
            if not code or code in used_codes:
                code = next_serial(conn, role)
            used_codes.add(code)
            conn.execute("INSERT INTO business_partners (partner_code, name, phone, address, cr_number, vat_number, "
                         "is_customer, is_supplier) VALUES (?,?,?,?,?,?,?,?)",
                         (code, name, _cell(row, phone_col), _cell(row, addr_col), _cell(row, cr_col),
                          normalize_digits(_cell(row, vat_col)), int(role == "customer"), int(role == "supplier")))
            existing[key] = (code, role == "customer", role == "supplier")
    return stats


def import_items_from_df(df, dry_run=False, update_opening=False):
    """أعمدة: الصنف (إجباري) + الكود (لو فارغ يُرقَّم تلقائياً). اختيارياً: الوحدة، الكمية، التكلفة (رصيد افتتاحي).
    update_opening=True: يحدّث الرصيد والتكلفة الافتتاحية للأصناف الموجودة أيضاً ثم يعيد احتساب أرصدتها."""
    name_col = _find_col(df, ['صنف', 'اسم الصنف', 'الصنف', 'Item', 'Name', 'الاسم', 'الوصف', 'Description'], 1)
    code_col = _find_col(df, ['كود', 'كود الصنف', 'الكود', 'Code', 'Item Code', 'No', 'رقم'], 0)
    if code_col == name_col:
        code_col = None
    unit_col = _find_col(df, ['الوحدة', 'وحدة', 'Unit', 'UOM'])
    qty_col = _find_col(df, ['الكمية', 'الرصيد', 'رصيد', 'رصيد افتتاحي', 'Qty', 'Quantity'])
    cost_col = _find_col(df, ['التكلفة', 'سعر التكلفة', 'تكلفة الوحدة', 'Cost'])
    stats = dict(added=0, existing=0, opening_updated=0, duplicates=0, empty=0, auto_codes=0)
    touched = []
    with db() as conn:
        existing = {r[0] for r in conn.execute("SELECT item_code FROM inventory")}
        seen = set()
        for _, row in df.iterrows():
            code = _cell(row, code_col)
            if code.endswith(".0"):
                code = code[:-2]
            name = clean_name(_cell(row, name_col))
            if not name:
                stats["empty"] += 1
                continue
            if code and code in seen:
                stats["duplicates"] += 1
                continue
            seen.add(code)
            qty = parse_float(_cell(row, qty_col)) if qty_col is not None else None
            cost = parse_float(_cell(row, cost_col)) if cost_col is not None else None
            if code in existing:
                if update_opening and (qty is not None or cost is not None):
                    stats["opening_updated"] += 1
                    if not dry_run:
                        conn.execute("UPDATE inventory SET opening_qty=COALESCE(?, opening_qty), "
                                     "opening_cost=COALESCE(?, opening_cost) WHERE item_code=?",
                                     (None if qty is None else max(qty, 0.0), None if cost is None else max(cost, 0.0), code))
                        touched.append(code)
                else:
                    stats["existing"] += 1
                continue
            stats["added"] += 1
            if dry_run:
                continue
            if not code:
                code = next_serial(conn, "item")
                stats["auto_codes"] += 1
            unit = _cell(row, unit_col) or "KG"
            qty, cost = max(qty or 0.0, 0.0), max(cost or 0.0, 0.0)
            conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, cost_price) "
                         "VALUES (?,?,?,?,?,?,?)", (code, name, unit, qty, cost, qty, cost))
            existing.add(code)
        if touched:
            recalc_items(conn, touched, allow_negative=True)
    return stats


def item_has_movements(conn, code):
    n = conn.execute("SELECT COUNT(*) FROM v_stock_in WHERE item_code=?", (code,)).fetchone()[0]
    n += conn.execute("SELECT COUNT(*) FROM v_stock_out WHERE item_code=?", (code,)).fetchone()[0]
    n += conn.execute("SELECT COUNT(*) FROM returns_invoices WHERE item_code=?", (code,)).fetchone()[0]
    return n > 0


# ============================================================
# الترقيم التسلسلي المستقل (أصناف، عملاء، موردون، فواتير مبيعات، فواتير مشتريات)
# ============================================================
# (الاسم، الوصف، البادئة الافتراضية، عدد الأرقام، الجدول، العمود)
SEQUENCE_DEFS = [
    ("item", "الأصناف", "ELH-", 4, "inventory", "item_code"),
    ("customer", "العملاء", "C", 4, "business_partners", "partner_code"),
    ("supplier", "الموردين", "S", 4, "business_partners", "partner_code"),
    ("sales_invoice", "فواتير المبيعات", "SI-", 6, "sales_invoices", "invoice_no"),
    ("purchase_invoice", "فواتير المشتريات", "PI-", 6, "purchase_invoices", "invoice_no"),
    ("purchase_discount", "إشعارات خصم الموردين (خصم لاحق)", "PD-", 6, "post_discounts", "discount_no"),
    ("sales_discount", "إشعارات دائنة للعملاء (خصم لاحق)", "CN-", 6, "post_discounts", "discount_no"),
]
SEQ_TARGET = {n: (t, c) for n, _, _, _, t, c in SEQUENCE_DEFS}
PARTNER_CODE_COLS = ['No', 'Code', 'كود', 'الكود', 'رقم', 'الرقم', 'كود العميل', 'كود المورد']


def _format_serial(prefix, n, width):
    return f"{prefix}{n:0{width}d}"


def max_used_serial(conn, name, prefix):
    table, col = SEQ_TARGET[name]
    pat = re.compile(re.escape(prefix) + r"(\d+)")
    best = 0
    for (v,) in conn.execute(f"SELECT {col} FROM {table} WHERE {col} IS NOT NULL"):
        m = pat.fullmatch(str(v))
        if m:
            best = max(best, int(m.group(1)))
    return best


def peek_serial(conn, name):
    """الرقم التالي بدون حجزه: (النص، الرقم). يتخطى أي رقم مستخدم يدوياً."""
    prefix, n, width = conn.execute("SELECT prefix, next_value, width FROM sequences WHERE name=?", (name,)).fetchone()
    table, col = SEQ_TARGET[name]
    while conn.execute(f"SELECT 1 FROM {table} WHERE {col}=?", (_format_serial(prefix, n, width),)).fetchone():
        n += 1
    return _format_serial(prefix, n, width), n


def next_serial(conn, name):
    """يحجز الرقم التالي في نفس المعاملة (لا تكرار، وكل تسلسل مستقل عن الآخر)."""
    code, n = peek_serial(conn, name)
    conn.execute("UPDATE sequences SET next_value=? WHERE name=?", (n + 1, name))
    return code


def init_sequences(conn):
    for name, label, prefix, width, _, _ in SEQUENCE_DEFS:
        if conn.execute("SELECT 1 FROM sequences WHERE name=?", (name,)).fetchone() is None:
            conn.execute("INSERT INTO sequences (name, label, prefix, next_value, width) VALUES (?,?,?,?,?)",
                         (name, label, prefix, max_used_serial(conn, name, prefix) + 1, width))
    for name, table, idcol in (("sales_invoice", "sales_invoices", "invoice_id"),
                               ("purchase_invoice", "purchase_invoices", "purchase_id")):
        for (iid,) in conn.execute(f"SELECT {idcol} FROM {table} WHERE invoice_no IS NULL OR invoice_no='' "
                                   f"ORDER BY {idcol}").fetchall():
            conn.execute(f"UPDATE {table} SET invoice_no=? WHERE {idcol}=?", (next_serial(conn, name), iid))
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_sales_invoice_no ON sales_invoices(invoice_no)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_purchase_invoice_no ON purchase_invoices(invoice_no)")


def update_sequence(name, prefix, next_value, width):
    prefix = (prefix or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9\-/_.]{0,10}", prefix):
        raise UserError("البادئة حتى 10 أحرف: حروف إنجليزية وأرقام و - / _ . فقط.")
    if not (1 <= width <= 12):
        raise UserError("عدد الأرقام بين 1 و 12.")
    if next_value < 1:
        raise UserError("الرقم التالي يجب أن يكون 1 أو أكثر.")
    with db() as conn:
        used = max_used_serial(conn, name, prefix)
        if next_value <= used:
            raise UserError(f"الرقم التالي ({next_value}) يجب أن يكون أكبر من آخر رقم مستخدم بنفس البادئة ({used}).")
        conn.execute("UPDATE sequences SET prefix=?, next_value=?, width=? WHERE name=?", (prefix, next_value, width, name))


def invoice_number(kind, inv_id):
    table, idcol = ("sales_invoices", "invoice_id") if kind == "sales" else ("purchase_invoices", "purchase_id")
    with db() as conn:
        r = conn.execute(f"SELECT COALESCE(invoice_no, {idcol}) FROM {table} WHERE {idcol}=?", (inv_id,)).fetchone()
    return str(r[0]) if r else str(inv_id)


# ============================================================
# ملفات الرفع (Excel / CSV) والنماذج
# ============================================================
def read_table_file(path):
    """قراءة Excel أو CSV كنصوص (CSV بترميز UTF-8 أو Windows-1256 العربي)."""
    if path.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1256"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise UserError("تعذرت قراءة ملف CSV. احفظه بترميز UTF-8.")
    return pd.read_excel(path, dtype=str)


UPLOAD_TEMPLATES = {
    "customer_openings": ["No", "Name Contact", "مدين (عليه)", "دائن (له)"],
    "supplier_openings": ["No", "Name Contact", "مدين (عليه)", "دائن (له)"],
    "items": ["كود", "صنف", "الوحدة", "الكمية", "التكلفة"],
}


def write_upload_template(path, kind):
    cols = UPLOAD_TEMPLATES[kind]
    sample = {"customer_openings": [["C0001", "عميل تجريبي", "1500", "0"]],
              "supplier_openings": [["S0001", "مورد تجريبي", "0", "2300"]],
              "items": [["", "صنف تجريبي (الكود يُرقَّم تلقائياً لو فارغ)", "KG", "100", "25.5"]]}[kind]
    df = pd.DataFrame(sample, columns=cols)
    if path.lower().endswith(".csv"):
        df.to_csv(path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(path, index=False)


# ============================================================
# الأرصدة الافتتاحية: العملاء والموردون + الخزينة والبنوك
# ============================================================
OPENING_ACCOUNT = "99999"           # وسيط افتتاحي (حقوق ملكية)
CASH_OPENING_SETTING = "cash_opening_entry_id"
OPEN_DEBIT_COLS = ['مدين', 'عليه', 'ما عليه', 'مدين (عليه)', 'Debit', 'Dr']
OPEN_CREDIT_COLS = ['دائن', 'له', 'ما له', 'دائن (له)', 'Credit', 'Cr']
OPEN_BALANCE_COLS = ['الرصيد', 'رصيد', 'الرصيد الافتتاحي', 'رصيد افتتاحي', 'Balance', 'Opening Balance']


def _net_opening(debit, credit):
    net = round((debit or 0.0) - (credit or 0.0), 2)
    return (net, 0.0) if net > 0 else (0.0, -net)


def set_partner_opening(conn, code, debit, credit, as_of):
    d, c = _net_opening(debit, credit)
    conn.execute("UPDATE business_partners SET opening_debit=?, opening_credit=?, opening_date=? WHERE partner_code=?",
                 (d, c, as_of if (d or c) else None, code))


def import_partner_openings_from_df(df, role, as_of, dry_run=False):
    """رفع الأرصدة الافتتاحية: أعمدة (مدين/عليه) و(دائن/له)، أو عمود رصيد واحد
    (للعملاء الموجب = عليه، وللموردين الموجب = له). المطابقة بالكود ثم بالاسم، والشريك غير الموجود يُضاف."""
    code_col = _find_col(df, PARTNER_CODE_COLS)
    name_col = _find_col(df, ['Name Contact', 'Name', 'الاسم', 'اسم', 'اسم العميل', 'اسم المورد', 'العميل', 'المورد'])
    debit_col, credit_col = _find_col(df, OPEN_DEBIT_COLS), _find_col(df, OPEN_CREDIT_COLS)
    bal_col = _find_col(df, OPEN_BALANCE_COLS)
    if code_col is None and name_col is None:
        raise UserError("الملف يجب أن يحتوي على عمود الكود (No / كود) أو الاسم (Name Contact / الاسم).")
    if debit_col is None and credit_col is None and bal_col is None:
        raise UserError("الملف يجب أن يحتوي على عمودي (مدين / دائن) أو (عليه / له) أو عمود (الرصيد).")
    flag = "is_customer" if role == "customer" else "is_supplier"
    stats = dict(updated=0, created=0, empty=0, bad=0, total_debit=0.0, total_credit=0.0)
    problems = []
    with db() as conn:
        by_code = {c for (c,) in conn.execute("SELECT partner_code FROM business_partners")}
        by_name = {norm_name(n): c for c, n in conn.execute("SELECT partner_code, name FROM business_partners")}
        for idx, (_, row) in enumerate(df.iterrows(), start=2):
            code = _cell(row, code_col)
            if code.endswith(".0"):
                code = code[:-2]
            name = clean_name(_cell(row, name_col))
            if not code and not name:
                stats["empty"] += 1
                continue
            try:
                def num(col):
                    raw = _cell(row, col)
                    if not raw:
                        return 0.0
                    v = parse_float(raw)
                    if v is None:
                        raise ValueError(raw)
                    return v
                debit, credit = num(debit_col), num(credit_col)
                if bal_col is not None and not (debit or credit):
                    b = num(bal_col)
                    if (b > 0) == (role == "customer"):
                        debit = abs(b)
                    else:
                        credit = abs(b)
            except ValueError as e:
                stats["bad"] += 1
                problems.append(f"سطر {idx}: قيمة غير رقمية ({e})")
                continue
            debit, credit = _net_opening(debit, credit)
            target = code if code in by_code else by_name.get(norm_name(name)) if name else None
            if target is None:
                if not name:
                    stats["bad"] += 1
                    problems.append(f"سطر {idx}: الكود {code} غير موجود ولا يوجد اسم لإضافته")
                    continue
                stats["created"] += 1
                if not dry_run:
                    target = code if code else next_serial(conn, role)
                    conn.execute("INSERT INTO business_partners (partner_code, name, is_customer, is_supplier) VALUES (?,?,?,?)",
                                 (target, name, int(role == "customer"), int(role == "supplier")))
                    by_code.add(target)
                    by_name[norm_name(name)] = target
            else:
                stats["updated"] += 1
                if not dry_run:
                    conn.execute(f"UPDATE business_partners SET {flag}=1 WHERE partner_code=?", (target,))
            stats["total_debit"] += debit
            stats["total_credit"] += credit
            if not dry_run:
                set_partner_opening(conn, target, debit, credit, as_of)
    return stats, problems


def get_cash_opening():
    """(التاريخ، {اسم الحساب: المبلغ}) من قيد الأرصدة الافتتاحية للخزينة والبنوك إن وُجد."""
    with db() as conn:
        row = conn.execute("SELECT value FROM app_settings WHERE key=?", (CASH_OPENING_SETTING,)).fetchone()
        if not row or not row[0]:
            return None, {}
        head = conn.execute("SELECT entry_date FROM journal_entries WHERE entry_id=?", (int(row[0]),)).fetchone()
        if head is None:
            return None, {}
        by_code = {code: name for name, code in cash_accounts(conn).items()}
        amounts = {by_code[c]: d - cr for c, d, cr in conn.execute(
            "SELECT account_code, debit, credit FROM journal_lines WHERE entry_id=?", (int(row[0]),)) if c in by_code}
    return head[0], amounts


def set_cash_opening(as_of, amounts, conn=None):
    """الأرصدة الافتتاحية للخزينة والبنوك كقيد واحد: من ح/ النقدية والبنوك إلى ح/ وسيط افتتاحي (99999).
    إعادة الحفظ تستبدل القيد السابق، والسالب (سحب على المكشوف) يُرحَّل دائناً. conn: معاملة مفتوحة (مثل الاستيراد)."""
    with (nullcontext(conn) if conn is not None else db()) as conn:
        accounts = cash_accounts(conn)
        old = conn.execute("SELECT value FROM app_settings WHERE key=?", (CASH_OPENING_SETTING,)).fetchone()
        if old and old[0]:
            conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (int(old[0]),))
            conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (int(old[0]),))
            conn.execute("DELETE FROM app_settings WHERE key=?", (CASH_OPENING_SETTING,))
        lines = [(accounts.get(n, n), n, round(a, 2)) for n, a in amounts.items() if abs(a or 0) > 0.005]
        if not lines:
            return None
        for code, name, _ in lines:
            if not conn.execute("SELECT 1 FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone():
                raise UserError(f"الحساب {name} غير موجود في دليل الحسابات.")
        acc = conn.execute("SELECT account_name FROM chart_of_accounts WHERE account_code=?", (OPENING_ACCOUNT,)).fetchone()
        if acc is None:
            conn.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                         "VALUES (?,?,?,NULL,0)", (OPENING_ACCOUNT, "وسيط افتتاحى", TYPE_EQUITY))
            acc = ("وسيط افتتاحى",)
        eid = conn.execute("INSERT INTO journal_entries (entry_date, description) VALUES (?,?)",
                           (as_of, "أرصدة افتتاحية - الخزينة والبنوك")).lastrowid
        for code, name, a in lines:
            conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                         "VALUES (?,?,?,?,?,?)", (eid, code, name, max(a, 0.0), max(-a, 0.0), "رصيد افتتاحي"))
        total = round(sum(a for _, _, a in lines), 2)
        if abs(total) > 0.005:
            conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                         "VALUES (?,?,?,?,?,?)", (eid, OPENING_ACCOUNT, acc[0], max(-total, 0.0), max(total, 0.0),
                                                  "مقابل الأرصدة الافتتاحية"))
        conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (CASH_OPENING_SETTING, str(eid)))
    return eid


# ============================================================
# مركز استيراد البيانات والترحيل (Excel / CSV): الأصناف، أرصدة العملاء، أرصدة الموردين، الخزينة والبنوك
# ============================================================
class _DryRun(Exception):
    """تُرفع داخل اتصال المعاينة لإلغاء كل التغييرات (حتى أرقام التسلسل) بعد حساب النتائج."""


IMPORT_KINDS = {
    "items": dict(
        title="استيراد الأصناف والمخزون",
        columns=["اسم الصنف", "الكود القديم", "الوحدة", "الكمية الافتتاحية", "تكلفة الوحدة", "سعر البيع", "الباركود"],
        sample=[["دجاج مجمد 1000 جم", "OLD-115", "كرتونة", "120", "45.50", "58", "6281000000017"],
                ["لحم بقري مفروم", "OLD-220", "KG", "35.5", "29", "37.5", ""]],
        help="المطابقة بالباركود ثم الكود القديم ثم اسم الصنف. الصنف الجديد يأخذ كوداً موحداً ELH-XXXX "
             "ويُحفظ الكود القديم كمرجع. الكمية والتكلفة تُسجَّلان رصيداً افتتاحياً."),
    "customers": dict(
        title="أرصدة العملاء أول المدة",
        columns=["الكود القديم", "اسم العميل", "الهاتف", "البريد الإلكتروني", "العنوان", "الرقم الضريبي",
                 "السجل التجاري", "المندوب", "مدين (عليه)", "دائن (له)"],
        sample=[["CUS-77", "مطعم النخبة", "0551234567", "info@example.com", "الرياض - حي العليا", "300000000000003",
                 "1010000000", "أحمد", "1500", "0"]],
        help="المطابقة بالكود ثم الاسم بين العملاء فقط. العميل الجديد يأخذ كوداً تلقائياً C#### يكمل آخر رقم مسجل. "
             "المندوب يُسجَّل اعتباراً من تاريخ الرصيد الافتتاحي. "
             "مدين (عليه) = مستحق لنا على العميل، دائن (له) = دفعة مقدمة من العميل."),
    "suppliers": dict(
        title="أرصدة الموردين أول المدة",
        columns=["الكود القديم", "اسم المورد", "الهاتف", "البريد الإلكتروني", "العنوان", "الرقم الضريبي",
                 "السجل التجاري", "المندوب", "مدين (عليه)", "دائن (له)"],
        sample=[["SUP-12", "مصنع الأغذية الحديثة", "0112345678", "", "جدة", "310000000000003", "4030000000", "خالد", "0",
                 "2300"]],
        help="منفصل تماماً عن العملاء: المطابقة بين الموردين فقط، والمورد الجديد يأخذ كوداً تلقائياً S#### حتى لو تشابه اسمه "
             "مع عميل (ويُربطان تلقائياً في كشف الحساب الموحد). "
             "دائن (له) = مستحق للمورد علينا، مدين (عليه) = دفعة مقدمة منا للمورد."),
    "cash": dict(
        title="أرصدة الخزينة والبنوك",
        columns=["الحساب", "الرصيد الافتتاحي"],
        sample=[["صندوق نقدية الرياض", "15000"], ["بنك الراجحى", "250000"]],
        help="اسم الحساب أو كوده من دليل الحسابات (فرع النقدية 111). يُسجَّل قيد افتتاحي واحد مقابل ح/ وسيط افتتاحي."),
}

_COLS_ITEM = dict(
    name=["اسم الصنف", "الصنف", "صنف", "الاسم", "اسم", "Item", "Item Name", "Name", "Description", "الوصف"],
    legacy=["الكود القديم", "كود قديم", "الكود", "كود", "كود الصنف", "رقم الصنف", "Code", "Item Code", "Old Code",
            "Legacy Code", "SKU"],
    unit=["الوحدة", "وحدة", "Unit", "UOM"],
    qty=["الكمية الافتتاحية", "الكمية", "الرصيد", "رصيد", "رصيد افتتاحي", "Qty", "Quantity", "Opening Qty"],
    cost=["تكلفة الوحدة", "التكلفة", "سعر التكلفة", "تكلفة", "Cost", "Unit Cost"],
    price=["سعر البيع", "البيع", "سعر", "Price", "Selling Price", "Sale Price"],
    barcode=["الباركود", "باركود", "Barcode", "EAN", "UPC"],
)
_COLS_PARTY = dict(
    legacy=["الكود القديم", "كود قديم", "الكود", "كود", "رقم", "الرقم", "كود العميل", "كود المورد", "No", "Code",
            "Old Code", "Legacy Code"],
    name=["اسم العميل", "اسم المورد", "الاسم", "اسم", "العميل", "المورد", "Name", "Name Contact", "Customer", "Supplier"],
    phone=["الهاتف", "هاتف", "الجوال", "جوال", "رقم الجوال", "تليفون", "Phone", "Mobile", "Tel"],
    email=["البريد الإلكتروني", "البريد الالكتروني", "البريد", "ايميل", "Email", "E-mail"],
    address=["العنوان", "عنوان", "المدينة", "Address", "City"],
    vat=["الرقم الضريبي", "رقم ضريبي", "VAT", "VAT Number", "VAT No", "Tax Number"],
    cr=["السجل التجاري", "سجل تجاري", "CR", "CR Number", "Commercial Register"],
    rep=["المندوب", "اسم المندوب", "مندوب المبيعات", "Sales Rep", "Rep", "Salesman"],
    debit=OPEN_DEBIT_COLS + ["مدين (عليه)", "ما عليه"],
    credit=OPEN_CREDIT_COLS + ["دائن (له)", "ما له"],
    balance=OPEN_BALANCE_COLS,
)
_COLS_CASH = dict(account=["الحساب", "اسم الحساب", "الخزينة", "البنك", "كود الحساب", "Account", "Account Name"],
                  balance=["الرصيد الافتتاحي", "الرصيد", "رصيد", "المبلغ", "Balance", "Amount"],
                  debit=OPEN_DEBIT_COLS, credit=OPEN_CREDIT_COLS)

IMPORT_STAT_LABELS = dict(
    rows="صفوف الملف", added="جديد", updated="محدَّث", recoded="أُعيد ترقيمه إلى ELH", opening_set="رصيد افتتاحي مسجل",
    duplicates="مكرر داخل الملف", empty="صفوف فارغة", bad="صفوف بها أخطاء", total_debit="إجمالي المدين",
    total_credit="إجمالي الدائن", stock_value="قيمة المخزون الافتتاحي", accounts="حسابات", total="الإجمالي")


def _col_key(s):
    return re.sub(r"[\s_\-()/.:]+", "", search_norm(s))


def _pick_col(df, candidates):
    """مطابقة مرنة لاسم العمود (تجاهل المسافات والأقواس والهمزات وحالة الأحرف)."""
    cols = {_col_key(c): c for c in df.columns}
    for cand in candidates:
        k = _col_key(cand)
        if k in cols:
            return cols[k]
    return None


def _num(row, col, label, problems, row_no, allow_negative=False):
    """(القيمة أو None لو الخلية فارغة، هل صالحة)."""
    raw = _cell(row, col)
    if raw == "":
        return None, True
    v = parse_float(raw)
    if v is None or (v < 0 and not allow_negative):
        problems.append(f"صف {row_no}: قيمة «{label}» غير صحيحة ({raw}).")
        return None, False
    return v, True


def _clean_code(v):
    v = (v or "").strip()
    return v[:-2] if v.endswith(".0") and v[:-2].isdigit() else v


def write_import_template(path, kind):
    spec = IMPORT_KINDS[kind]
    df = pd.DataFrame(spec["sample"], columns=spec["columns"])
    if path.lower().endswith(".csv"):
        df.to_csv(path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(path, index=False, sheet_name=spec["title"][:31])


def _rename_item(conn, old_code, new_code):
    """تغيير كود صنف ونقله لكل المستندات داخل نفس الاتصال."""
    conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, cost_price, "
                 "selling_price, barcode, legacy_code) SELECT ?, item_name, unit, opening_qty, opening_cost, quantity, "
                 "cost_price, selling_price, NULL, legacy_code FROM inventory WHERE item_code=?", (new_code, old_code))
    barcode = conn.execute("SELECT barcode FROM inventory WHERE item_code=?", (old_code,)).fetchone()[0]
    for t in ITEM_REF_TABLES:
        conn.execute(f"UPDATE {t} SET item_code=? WHERE item_code=?", (new_code, old_code))
    conn.execute("DELETE FROM inventory WHERE item_code=?", (old_code,))
    conn.execute("UPDATE inventory SET barcode=? WHERE item_code=?", (barcode, new_code))


def import_items_center(df, dry_run=False, update_existing=True, recode_matched=False):
    """استيراد الأصناف: المطابقة بالباركود ← الكود القديم ← الاسم. الجديد يأخذ ELH-XXXX والكود القديم يُحفظ مرجعاً.
    update_existing: تحديث بيانات الأصناف المطابقة (الوحدة/السعر/الباركود/الرصيد الافتتاحي).
    recode_matched: توحيد أكواد الأصناف المطابقة إلى ELH-XXXX (يُنقل الكود لكل المستندات ويُحفظ القديم مرجعاً)."""
    c = {k: _pick_col(df, v) for k, v in _COLS_ITEM.items()}
    if c["name"] is None:
        raise UserError("الملف يجب أن يحتوي على عمود «اسم الصنف».")
    stats = dict(rows=len(df), added=0, updated=0, recoded=0, duplicates=0, empty=0, bad=0, stock_value=0.0)
    problems, preview, touched = [], [], []
    try:
        with db() as conn:
            by_name, by_legacy, by_barcode, name_of = {}, {}, {}, {}
            for code, name, legacy, bc in conn.execute("SELECT item_code, item_name, legacy_code, barcode FROM inventory"):
                by_name.setdefault(norm_name(name), code)
                name_of[code] = norm_name(name)
                by_legacy[code.lower()] = code
                if legacy:
                    by_legacy.setdefault(str(legacy).lower(), code)
                if bc:
                    by_barcode[str(bc)] = code
            seen, seen_bc = set(), set()
            for idx, (_, row) in enumerate(df.iterrows(), start=2):
                name = clean_name(_cell(row, c["name"]))
                legacy = _clean_code(_cell(row, c["legacy"]))
                barcode = _clean_code(_cell(row, c["barcode"]))
                if not name:
                    if legacy or barcode:
                        problems.append(f"صف {idx}: اسم الصنف فارغ.")
                        stats["bad"] += 1
                    else:
                        stats["empty"] += 1
                    continue
                key = norm_name(name)
                if key in seen:
                    stats["duplicates"] += 1
                    preview.append((idx, "مكرر - تخطي", "", name, legacy))
                    continue
                seen.add(key)
                qty, ok1 = _num(row, c["qty"], "الكمية", problems, idx)
                cost, ok2 = _num(row, c["cost"], "التكلفة", problems, idx)
                price, ok3 = _num(row, c["price"], "سعر البيع", problems, idx)
                if not (ok1 and ok2 and ok3):
                    stats["bad"] += 1
                    continue
                unit = _cell(row, c["unit"])
                if barcode and barcode in seen_bc:
                    problems.append(f"صف {idx}: الباركود {barcode} مكرر داخل الملف؛ لم يُسجَّل لهذا الصنف.")
                    barcode = ""
                code = (by_legacy.get(legacy.lower()) if legacy else None) or by_name.get(key)
                bc_code = by_barcode.get(barcode) if barcode else None
                if bc_code and code is None and name_of.get(bc_code) == key:
                    code = bc_code
                elif bc_code and bc_code != code:
                    # الباركود لا يفرض المطابقة مع صنف باسم مختلف (حماية من الكتابة فوق صنف آخر بالخطأ)
                    problems.append(f"صف {idx}: الباركود {barcode} مسجل للصنف {bc_code} باسم مختلف؛ لم يُسجَّل لهذا الصنف.")
                    barcode = ""
                if barcode:
                    seen_bc.add(barcode)
                if code is None:
                    code = next_serial(conn, "item")
                    conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, "
                                 "cost_price, selling_price, barcode, legacy_code) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                 (code, name, unit or "KG", qty or 0.0, cost or 0.0, qty or 0.0, cost or 0.0, price or 0.0,
                                  barcode or None, legacy or None))
                    by_name[key] = code
                    name_of[code] = key
                    if legacy:
                        by_legacy[legacy.lower()] = code
                    if barcode:
                        by_barcode[barcode] = code
                    stats["added"] += 1
                    stats["stock_value"] += (qty or 0.0) * (cost or 0.0)
                    preview.append((idx, "جديد", code, name, legacy))
                    continue
                action = "مطابق"
                if recode_matched and not code.startswith("ELH-"):
                    new_code = next_serial(conn, "item")
                    old = code
                    _rename_item(conn, old, new_code)
                    conn.execute("UPDATE inventory SET legacy_code=COALESCE(NULLIF(legacy_code,''), ?) WHERE item_code=?",
                                 (old, new_code))
                    by_name[key] = new_code
                    name_of[new_code] = key
                    by_legacy[old.lower()] = new_code
                    code = new_code
                    stats["recoded"] += 1
                    action = f"أُعيد ترقيمه (كان {old})"
                    touched.append(code)
                if update_existing:
                    sets, vals = [], []
                    for col, v in (("unit", unit or None), ("selling_price", price), ("barcode", barcode or None),
                                   ("opening_qty", qty), ("opening_cost", cost)):
                        if v is not None:
                            sets.append(f"{col}=?")
                            vals.append(v)
                    if legacy:
                        sets.append("legacy_code=COALESCE(NULLIF(legacy_code,''), ?)")
                        vals.append(legacy)
                    if sets:
                        conn.execute(f"UPDATE inventory SET {', '.join(sets)} WHERE item_code=?", vals + [code])
                        if barcode:
                            by_barcode[barcode] = code
                        if qty is not None or cost is not None:
                            touched.append(code)
                            o = conn.execute("SELECT opening_qty, opening_cost FROM inventory WHERE item_code=?",
                                             (code,)).fetchone()
                            stats["stock_value"] += (o[0] or 0) * (o[1] or 0)
                        stats["updated"] += 1
                        action = "تحديث" if action == "مطابق" else action + " + تحديث"
                preview.append((idx, action, code, name, legacy))
            if touched:
                recalc_items(conn, touched, allow_negative=True)
            if dry_run:
                raise _DryRun()
    except _DryRun:
        pass
    return dict(stats=stats, problems=problems, preview=preview)


def import_parties_center(df, role, as_of, dry_run=False):
    """استيراد العملاء أو الموردين مع رصيد أول المدة. role = customer (C-XXXX) | supplier (S-XXXX).
    المطابقة داخل نفس الدور فقط (بالكود ثم الكود القديم ثم الاسم) فلا يختلط العملاء بالموردين."""
    if not parse_date(as_of or ""):
        raise UserError("تاريخ الرصيد الافتتاحي غير صحيح (YYYY-MM-DD).")
    c = {k: _pick_col(df, v) for k, v in _COLS_PARTY.items()}
    if c["name"] is None and c["legacy"] is None:
        raise UserError("الملف يجب أن يحتوي على عمود الاسم (أو الكود).")
    flag = "is_customer" if role == "customer" else "is_supplier"
    seq = "customer" if role == "customer" else "supplier"
    who = "العميل" if role == "customer" else "المورد"
    stats = dict(rows=len(df), added=0, updated=0, opening_set=0, duplicates=0, empty=0, bad=0,
                 total_debit=0.0, total_credit=0.0)
    problems, preview = [], []
    try:
        with db() as conn:
            by_code, by_legacy, by_name, both = {}, {}, {}, set()
            for code, name, legacy, ic, is_ in conn.execute(
                    f"SELECT partner_code, name, legacy_code, is_customer, is_supplier FROM business_partners WHERE {flag}=1"):
                by_code[code.lower()] = code
                if legacy:
                    by_legacy.setdefault(str(legacy).lower(), code)
                by_name.setdefault(norm_name(name), code)
                if ic and is_:
                    both.add(code)
            seen = set()
            for idx, (_, row) in enumerate(df.iterrows(), start=2):
                name = clean_name(_cell(row, c["name"]))
                legacy = _clean_code(_cell(row, c["legacy"]))
                if not name and not legacy:
                    stats["empty"] += 1
                    continue
                key = (legacy.lower() if legacy else "") + "|" + norm_name(name)
                if key in seen:
                    stats["duplicates"] += 1
                    preview.append((idx, "مكرر - تخطي", "", name, legacy, 0.0, 0.0))
                    continue
                seen.add(key)
                debit, ok1 = _num(row, c["debit"], "مدين", problems, idx)
                credit, ok2 = _num(row, c["credit"], "دائن", problems, idx)
                bal, ok3 = _num(row, c["balance"], "الرصيد", problems, idx, allow_negative=True)
                if not (ok1 and ok2 and ok3):
                    stats["bad"] += 1
                    continue
                if debit is None and credit is None and bal is not None:  # عمود رصيد واحد
                    pos_is_debit = role == "customer"
                    debit, credit = ((bal, 0.0) if bal >= 0 else (0.0, -bal)) if pos_is_debit else \
                                    ((0.0, bal) if bal >= 0 else (-bal, 0.0))
                code = (by_code.get(legacy.lower()) if legacy else None) or \
                       (by_legacy.get(legacy.lower()) if legacy else None) or (by_name.get(norm_name(name)) if name else None)
                fields = {k: _cell(row, c[k]) for k in ("phone", "email", "address", "vat", "cr")}
                rep_name = clean_name(_cell(row, c["rep"]))
                if code is None:
                    if not name:
                        problems.append(f"صف {idx}: {who} غير موجود والاسم فارغ.")
                        stats["bad"] += 1
                        continue
                    code = next_serial(conn, seq)
                    conn.execute(
                        "INSERT INTO business_partners (partner_code, name, phone, email, address, vat_number, cr_number, "
                        "is_customer, is_supplier, legacy_code) VALUES (?,?,?,?,?,?,?,?,?,?)",
                        (code, name, fields["phone"], fields["email"], fields["address"], fields["vat"], fields["cr"],
                         1 if role == "customer" else 0, 1 if role == "supplier" else 0, legacy or None))
                    by_name[norm_name(name)] = code
                    by_code[code.lower()] = code
                    if legacy:
                        by_legacy[legacy.lower()] = code
                    stats["added"] += 1
                    action = "جديد"
                else:
                    sets, vals = [], []
                    for col, k in (("phone", "phone"), ("email", "email"), ("address", "address"), ("vat_number", "vat"),
                                   ("cr_number", "cr")):
                        if fields[k]:
                            sets.append(f"{col}=?")
                            vals.append(fields[k])
                    if legacy:
                        sets.append("legacy_code=COALESCE(NULLIF(legacy_code,''), ?)")
                        vals.append(legacy)
                    if sets:
                        conn.execute(f"UPDATE business_partners SET {', '.join(sets)} WHERE partner_code=?", vals + [code])
                    stats["updated"] += 1
                    action = "تحديث"
                    if code in both and (debit or credit):
                        problems.append(f"صف {idx}: «{name or code}» مسجل عميلاً ومورداً معاً؛ رصيده الافتتاحي مشترك بين الدورين.")
                if rep_name:
                    set_partner_rep(conn, code, rep_name, as_of)
                if debit is not None or credit is not None:
                    set_partner_opening(conn, code, debit or 0.0, credit or 0.0, as_of)
                    d, cr = _net_opening(debit, credit)
                    stats["opening_set"] += 1
                    stats["total_debit"] += d
                    stats["total_credit"] += cr
                else:
                    d = cr = 0.0
                preview.append((idx, action, code, name, legacy, d, cr))
            if dry_run:
                raise _DryRun()
    except _DryRun:
        pass
    return dict(stats=stats, problems=problems, preview=preview)


def import_cash_center(df, as_of, dry_run=False):
    """أرصدة الخزينة والبنوك من ملف: (الحساب، الرصيد) أو (مدين، دائن). تُدمج مع الأرصدة الافتتاحية الحالية."""
    if not parse_date(as_of or ""):
        raise UserError("تاريخ الرصيد الافتتاحي غير صحيح (YYYY-MM-DD).")
    c = {k: _pick_col(df, v) for k, v in _COLS_CASH.items()}
    if c["account"] is None:
        raise UserError("الملف يجب أن يحتوي على عمود «الحساب».")
    with db() as conn:
        accounts = cash_accounts(conn)
    by_code = {code: name for name, code in accounts.items()}
    by_norm = {norm_name(n): n for n in accounts}
    _, current = get_cash_opening()
    amounts = dict(current)
    stats = dict(rows=len(df), accounts=0, bad=0, empty=0, total=0.0)
    problems, preview = [], []
    for idx, (_, row) in enumerate(df.iterrows(), start=2):
        acc = _clean_code(_cell(row, c["account"]))
        if not acc:
            stats["empty"] += 1
            continue
        name = by_code.get(acc) or by_norm.get(norm_name(acc))
        if name is None:
            problems.append(f"صف {idx}: الحساب «{acc}» غير موجود ضمن حسابات الخزينة والبنوك (فرع {CASH_PREFIX}).")
            stats["bad"] += 1
            continue
        bal, ok1 = _num(row, c["balance"], "الرصيد", problems, idx, allow_negative=True)
        d, ok2 = _num(row, c["debit"], "مدين", problems, idx)
        cr, ok3 = _num(row, c["credit"], "دائن", problems, idx)
        if not (ok1 and ok2 and ok3):
            stats["bad"] += 1
            continue
        amount = bal if bal is not None else (d or 0.0) - (cr or 0.0)
        amounts[name] = round(amount, 2)
        stats["accounts"] += 1
        preview.append((idx, "تحديث" if name in current else "جديد", accounts[name], name, "", amount))
    stats["total"] = round(sum(amounts.values()), 2)
    if not dry_run and stats["accounts"]:
        set_cash_opening(as_of, amounts)
    return dict(stats=stats, problems=problems, preview=preview)


def run_import(kind, df, dry_run=False, as_of=None, **opts):
    if df is None or df.empty:
        raise UserError("الملف فارغ.")
    if kind == "items":
        return import_items_center(df, dry_run=dry_run, **opts)
    if kind in ("customers", "suppliers"):
        return import_parties_center(df, "customer" if kind == "customers" else "supplier", as_of, dry_run=dry_run)
    if kind == "cash":
        return import_cash_center(df, as_of, dry_run=dry_run)
    raise UserError("نوع استيراد غير معروف.")


def import_preview_columns(kind):
    if kind == "items":
        return ["الصف", "الإجراء", "الكود الموحد", "اسم الصنف", "الكود القديم"]
    if kind == "cash":
        return ["الصف", "الإجراء", "كود الحساب", "الحساب", "", "الرصيد"]
    return ["الصف", "الإجراء", "الكود", "الاسم", "الكود القديم", "مدين (عليه)", "دائن (له)"]


def import_stats_text(stats):
    parts = []
    for k, v in stats.items():
        if isinstance(v, float):
            parts.append(f"{IMPORT_STAT_LABELS.get(k, k)}: {v:,.2f}")
        elif v:
            parts.append(f"{IMPORT_STAT_LABELS.get(k, k)}: {v}")
    return "   |   ".join(parts)


# ============================================================
# صلاحيات المدير الكاملة: تعديل وحذف أي مستند مع إعادة احتساب المخزون
# ============================================================
class NegativeStockError(UserError):
    """التعديل يؤدي لرصيد مخزون سالب؛ المدير يستطيع التأكيد والمتابعة."""


ITEM_REF_TABLES = ["sales_invoice_lines", "purchase_invoice_lines", "returns_invoices", "sales_order_lines",
                   "purchase_order_lines", "sales_delivery_lines", "purchase_receipt_lines", "post_discount_lines"]


def validate_batch_dates(prod, exp, label=""):
    where = f" للصنف «{label}»" if label else ""
    for v in (prod, exp):
        if v and not parse_date(v):
            raise UserError(f"صيغة التاريخ غير صحيحة ({v}){where}. استخدم YYYY-MM-DD")
    if prod and exp and exp < prod:
        raise UserError(f"تاريخ الانتهاء قبل تاريخ الإنتاج{where}.")


def recalc_items(conn, codes, allow_negative=False):
    """إعادة بناء رصيد ومتوسط تكلفة كل صنف من الصفر بتسلسل الحركات زمنياً، وتحديث تكلفة المبيعات والمرتجعات.
    تُستخدم بعد أي تعديل أو حذف إداري حتى تبقى أرصدة المخزون وتكلفة البضاعة المباعة صحيحة."""
    negatives = []
    disc_lines = set()
    for code in sorted({c for c in codes if c}):
        row = conn.execute("SELECT item_name, opening_qty, opening_cost FROM inventory WHERE item_code=?", (code,)).fetchone()
        if row is None:
            continue
        name, q, c = row[0], row[1] or 0.0, row[2] or 0.0
        moves = []
        queries = [
            ("pin", 0, "SELECT l.line_id, h.purchase_date, h.created_at, l.quantity, l.unit_cost FROM purchase_invoice_lines l "
                       "JOIN purchase_invoices h ON h.purchase_id=l.purchase_id WHERE l.item_code=? AND l.receipt_line_id IS NULL"),
            ("rin", 0, "SELECT l.line_id, r.receipt_date, r.created_at, l.quantity, l.unit_cost FROM purchase_receipt_lines l "
                       "JOIN purchase_receipts r ON r.receipt_id=l.receipt_id WHERE l.item_code=?"),
            ("srin", 0, f"SELECT return_id, return_date, created_at, quantity, unit_cost FROM returns_invoices "
                        f"WHERE item_code=? AND return_type='{SALES_RETURN}'"),
            ("sout", 1, "SELECT l.line_id, h.invoice_date, h.created_at, l.quantity, 0 FROM sales_invoice_lines l "
                        "JOIN sales_invoices h ON h.invoice_id=l.invoice_id WHERE l.item_code=? AND l.delivery_line_id IS NULL"),
            ("dout", 1, "SELECT l.line_id, d.delivery_date, d.created_at, l.quantity, 0 FROM sales_delivery_lines l "
                        "JOIN sales_deliveries d ON d.delivery_id=l.delivery_id WHERE l.item_code=?"),
            ("prout", 1, f"SELECT return_id, return_date, created_at, quantity, 0 FROM returns_invoices "
                         f"WHERE item_code=? AND return_type='{PURCHASE_RETURN}'"),
            # خصم لاحق من المورد بطريقة «تخفيض تكلفة المخزون»: يخفض قيمة الرصيد الموجود دون تغيير الكمية
            ("disc", 0, "SELECT l.line_id, d.discount_date, d.created_at, l.amount, 0 FROM post_discount_lines l "
                        "JOIN post_discounts d ON d.discount_id=l.discount_id WHERE l.item_code=? AND d.kind='purchase' "
                        "AND d.method='stock'"),
        ]
        for kind, order, sql in queries:
            for rid, d, created, qty, cost in conn.execute(sql, (code,)):
                moves.append((d or "", created or "", order, rid, kind, qty or 0.0, cost or 0.0))
        moves.sort()
        for _, _, _, rid, kind, qty, cost in moves:
            if kind == "disc":
                on_hand = max(q, 0.0)
                value = on_hand * c
                applied = round(min(qty, value), 2) if on_hand > EPS else 0.0
                if on_hand > EPS:
                    c = round(max(value - applied, 0.0) / on_hand, 4)
                conn.execute("UPDATE post_discount_lines SET applied=? WHERE line_id=?", (applied, rid))
                disc_lines.add(rid)
                continue
            if kind in ("pin", "rin", "srin"):
                old_q = max(q, 0.0)
                new_q = old_q + qty
                c = round((old_q * c + qty * cost) / new_q, 4) if new_q > 0 else cost
                q += qty
            else:
                if kind == "sout":
                    conn.execute("UPDATE sales_invoice_lines SET unit_cost=? WHERE line_id=?", (c, rid))
                elif kind == "dout":
                    conn.execute("UPDATE sales_delivery_lines SET unit_cost=? WHERE line_id=?", (c, rid))
                    conn.execute("UPDATE sales_invoice_lines SET unit_cost=? WHERE delivery_line_id=?", (c, rid))
                else:
                    conn.execute("UPDATE returns_invoices SET unit_cost=? WHERE return_id=?", (c, rid))
                q -= qty
        q = round(q, 3)
        conn.execute("UPDATE inventory SET quantity=?, cost_price=? WHERE item_code=?", (q, c, code))
        if q < -EPS:
            negatives.append(f"{name} ({code}): {q:g}")
    if disc_lines:
        refresh_discount_headers(conn, disc_lines)
    if negatives and not allow_negative:
        raise NegativeStockError("هذه العملية تجعل رصيد المخزون سالباً للأصناف:\n" + "\n".join(negatives[:10]))


def advance_allocated_to(conn, kind, inv_id):
    """ما سُوّي على الفاتورة من دفعات مقدمة/تحت الحساب."""
    try:
        return conn.execute("SELECT COALESCE(SUM(amount),0) FROM advance_allocations WHERE kind=? AND invoice_id=?",
                            (kind, inv_id)).fetchone()[0]
    except sqlite3.OperationalError:   # قاعدة قديمة قبل التهيئة
        return 0.0


def advance_allocated_from(conn, trans_id):
    """ما سُوّي من سند دفعة مقدمة على فواتير."""
    try:
        return conn.execute("SELECT COALESCE(SUM(amount),0) FROM advance_allocations WHERE trans_id=?",
                            (trans_id,)).fetchone()[0]
    except sqlite3.OperationalError:
        return 0.0


def refresh_invoice_paid(conn, kind, inv_id):
    """المسدد = مجموع السندات المربوطة بالفاتورة، والحالة تُحدَّث تبعاً."""
    c = FLOW[kind]
    paid = conn.execute("SELECT COALESCE(SUM(amount),0) FROM cash_transactions WHERE alloc_type=? AND alloc_id=?",
                        (kind, inv_id)).fetchone()[0] + advance_allocated_to(conn, kind, inv_id)
    total = conn.execute(f"SELECT total_amount FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
    if total is None:
        return
    # الخصومات اللاحقة المربوطة بالفاتورة تُسوّي جزءاً منها (المسدد نقداً يبقى كما هو)
    settled = paid + invoice_discount_total(conn, kind, inv_id)
    status = "مغلقة" if settled >= total[0] - 0.005 and total[0] > 0.005 else "مفتوحة"
    conn.execute(f"UPDATE {c['inv']} SET paid_amount=?, status=? WHERE {c['iid']}=?", (round(paid, 2), status, inv_id))


def _lines_snapshot(conn, table, key_col, key):
    cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]
    return [dict(zip(cols, r)) for r in conn.execute(f"SELECT * FROM {table} WHERE {key_col}=? ORDER BY rowid", (key,))]


def _audit_lines(conn, action, label, ref, table, key_col, key, before):
    """سجل التدقيق لبنود المستندات المركبة (تُحذف وتُعاد كتابتها عند التعديل): لقطة كاملة قبل/بعد."""
    from .audit import audit_event
    after = None if action == "حذف" else _lines_snapshot(conn, table, key_col, key)
    strip = lambda rows: [{k: v for k, v in r.items() if k != "line_id"} for r in (rows or [])]  # noqa: E731
    if after is not None and strip(before) == strip(after):
        return
    audit_event(conn, action, label, ref, old=before, new=after, table_name=table)


def delete_invoice(kind, inv_id, allow_negative=False):
    """حذف فاتورة مبيعات/مشتريات: البنود المباشرة تُعكس من المخزون، وأذون التسليم/الاستلام تعود «غير مفوترة»،
    والسندات المربوطة تبقى كدفعات على حساب الشريك (يُلغى ربطها بالفاتورة فقط)."""
    c = FLOW[kind]
    with db() as conn:
        if conn.execute(f"SELECT 1 FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone() is None:
            raise UserError("الفاتورة غير موجودة.")
        n_disc = conn.execute("SELECT COUNT(*) FROM post_discounts WHERE kind=? AND invoice_id=?", (kind, inv_id)).fetchone()[0]
        if n_disc:
            raise UserError(f"على هذه الفاتورة {n_disc} إشعار خصم لاحق مربوط بها.\n"
                            "احذف إشعارات الخصم أولاً (من تبويب «الخصم اللاحق») ثم احذف الفاتورة.")
        rows = conn.execute(f"SELECT item_code, quantity, {c['link']} FROM {c['ilines']} WHERE {c['iid']}=?",
                            (inv_id,)).fetchall()
        no = conn.execute(f"SELECT COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()[0]
        _audit_lines(conn, "حذف", f"بنود فاتورة {'مبيعات' if kind == 'sales' else 'مشتريات'}", no, c["ilines"], c["iid"],
                     inv_id, _lines_snapshot(conn, c["ilines"], c["iid"], inv_id))
        orders = set()
        for _, qty, link in rows:
            if link is None:
                continue
            ol = conn.execute(f"SELECT l.order_line_id, o.order_id FROM {c['flines']} l JOIN {c['olines']} o "
                              f"ON o.line_id=l.order_line_id WHERE l.line_id=?", (link,)).fetchone()
            if ol:
                conn.execute(f"UPDATE {c['olines']} SET invoiced_qty = MAX(invoiced_qty - ?, 0) WHERE line_id=?", (qty, ol[0]))
                orders.add(ol[1])
        conn.execute(f"UPDATE {c['ful']} SET invoice_id=NULL WHERE invoice_id=?", (inv_id,))
        conn.execute("UPDATE cash_transactions SET alloc_type=NULL, alloc_id=NULL WHERE alloc_type=? AND alloc_id=?",
                     (kind, inv_id))
        conn.execute("DELETE FROM advance_allocations WHERE kind=? AND invoice_id=?", (kind, inv_id))
        conn.execute(f"DELETE FROM {c['ilines']} WHERE {c['iid']}=?", (inv_id,))
        conn.execute(f"DELETE FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,))
        if kind == "purchase":
            _delete_doc_attachments(conn, "purchase", inv_id)
        for oid in orders:
            refresh_order_status(conn, kind, oid)
        recalc_items(conn, [r[0] for r in rows], allow_negative)


def load_invoice(kind, inv_id):
    """بيانات الفاتورة للمحرر: (الرأس، البنود). linked=True للبند الصادر من إذن تسليم/استلام."""
    c = FLOW[kind]
    sales = kind == "sales"
    ref_col = "''" if sales else "COALESCE(supplier_ref,'')"
    with db() as conn:
        h = conn.execute(f"SELECT {c['idate']}, {c['pcode']}, {c['pname']}, COALESCE(notes,''), COALESCE(invoice_no,''), "
                         f"{ref_col} FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
        if h is None:
            raise UserError("الفاتورة غير موجودة.")
        price_col = "unit_price" if sales else "unit_cost"
        extra = "COALESCE(discount_amount,0), NULL, NULL" if sales else "0, production_date, expiry_date"
        lines = [dict(line_id=r[0], code=r[1], name=r[2], unit=r[3] or "", qty=r[4], price=r[5], discount=r[6],
                      production=r[7] or "", expiry=r[8] or "", linked=r[9] is not None)
                 for r in conn.execute(f"SELECT line_id, item_code, item_name, unit, quantity, {price_col}, {extra}, {c['link']} "
                                       f"FROM {c['ilines']} WHERE {c['iid']}=? ORDER BY line_id", (inv_id,))]
    return dict(date=h[0], party=(h[1], h[2]), notes=h[3], invoice_no=h[4], supplier_ref=h[5] or ""), lines


def update_invoice(kind, inv_id, header, lines, allow_negative=False):
    """تعديل فاتورة بالكامل (المدير): الرأس والبنود (إضافة/تعديل/حذف) مع إعادة احتساب المخزون والتكلفة.
    في الفواتير الصادرة من أذون تسليم/استلام: يمكن تعديل الأسعار والخصم والتواريخ، أما الكميات والأصناف فمرتبطة بالأذون."""
    c = FLOW[kind]
    sales = kind == "sales"
    if not lines:
        raise UserError("الفاتورة يجب أن تحتوي على صنف واحد على الأقل. لإلغائها استخدم «حذف الفاتورة».")
    if not parse_date(header.get("date") or ""):
        raise UserError("تاريخ الفاتورة غير صحيح (YYYY-MM-DD).")
    party = header["party"]
    with db() as conn:
        head = conn.execute(f"SELECT {c['pcode']} FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
        if head is None:
            raise UserError("الفاتورة غير موجودة.")
        old = {r[0]: r for r in conn.execute(f"SELECT line_id, item_code, quantity, {c['link']} FROM {c['ilines']} "
                                             f"WHERE {c['iid']}=?", (inv_id,))}
        before_lines = _lines_snapshot(conn, c["ilines"], c["iid"], inv_id)
        linked = {lid for lid, r in old.items() if r[3] is not None}
        kept = {l.get("line_id") for l in lines if l.get("line_id") in old}
        disc_sub = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM post_discounts WHERE kind=? AND invoice_id=?",
                                (kind, inv_id)).fetchone()[0]
        if disc_sub > EPS and party[0] != head[0]:
            raise UserError("لا يمكن تغيير الطرف في فاتورة عليها إشعارات خصم لاحق. احذف الإشعارات أولاً.")
        if linked:
            if party[0] != head[0]:
                raise UserError("لا يمكن تغيير الطرف في فاتورة صادرة من أذون تسليم/استلام. احذفها ثم أعد إصدارها.")
            if linked - kept:
                raise UserError("بنود الأذون لا تُحذف من الفاتورة منفردة. لإلغاء ربط الأذون احذف الفاتورة كاملة.")
        built = []
        for l in lines:
            if sales:
                L = make_line(l["code"], l["name"], l.get("unit", ""), l["qty"], l["price"], l.get("discount", 0.0))
            else:
                L = make_line(l["code"], l["name"], l.get("unit", ""), l["qty"], l["price"])
                validate_batch_dates(l.get("production"), l.get("expiry"), l["name"])
            lid = l.get("line_id") if l.get("line_id") in old else None
            if lid in linked and (l["code"] != old[lid][1] or abs(L["qty"] - old[lid][2]) > EPS):
                raise UserError(f"صنف وكمية البند «{l['name']}» مرتبطان بإذن تسليم/استلام ولا يُعدَّلان من الفاتورة.")
            if L["qty"] <= 0 or L["price"] < 0:
                raise UserError(f"كمية/سعر غير صحيح للبند «{l['name']}».")
            L.update(line_id=lid, production=l.get("production") or None, expiry=l.get("expiry") or None)
            built.append(L)
        for lid in set(old) - kept:
            conn.execute(f"DELETE FROM {c['ilines']} WHERE line_id=?", (lid,))
        for L in built:
            if sales:
                vals = (L["code"], L["name"], L["unit"], L["qty"], L["price"], L["discount"], L["subtotal"], L["vat"], L["total"])
                if L["line_id"]:
                    conn.execute("UPDATE sales_invoice_lines SET item_code=?, item_name=?, unit=?, quantity=?, unit_price=?, "
                                 "discount_amount=?, subtotal=?, vat_amount=?, total_amount=? WHERE line_id=?",
                                 vals + (L["line_id"],))
                    link = old[L["line_id"]][3]
                    if link is not None:
                        conn.execute("UPDATE sales_delivery_lines SET unit_price=?, discount_amount=? WHERE line_id=?",
                                     (L["price"], L["discount"], link))
                else:
                    conn.execute("INSERT INTO sales_invoice_lines (item_code, item_name, unit, quantity, unit_price, "
                                 "discount_amount, subtotal, vat_amount, total_amount, invoice_id, unit_cost) "
                                 "VALUES (?,?,?,?,?,?,?,?,?,?,0)", vals + (inv_id,))
            else:
                vals = (L["code"], L["name"], L["unit"], L["qty"], L["price"], L["subtotal"], L["vat"], L["total"],
                        L["production"], L["expiry"])
                if L["line_id"]:
                    conn.execute("UPDATE purchase_invoice_lines SET item_code=?, item_name=?, unit=?, quantity=?, unit_cost=?, "
                                 "subtotal=?, vat_amount=?, total_amount=?, production_date=?, expiry_date=? WHERE line_id=?",
                                 vals + (L["line_id"],))
                    link = old[L["line_id"]][3]
                    if link is not None:
                        conn.execute("UPDATE purchase_receipt_lines SET unit_cost=?, production_date=?, expiry_date=? "
                                     "WHERE line_id=?", (L["price"], L["production"], L["expiry"], link))
                else:
                    conn.execute("INSERT INTO purchase_invoice_lines (item_code, item_name, unit, quantity, unit_cost, "
                                 "subtotal, vat_amount, total_amount, production_date, expiry_date, purchase_id) "
                                 "VALUES (?,?,?,?,?,?,?,?,?,?,?)", vals + (inv_id,))
        subtotal = round(sum(L["subtotal"] for L in built), 2)
        vat = round(sum(L["vat"] for L in built), 2)
        if disc_sub > subtotal + 0.005:
            raise UserError(f"إجمالي الفاتورة بعد التعديل ({fmt(subtotal)} قبل الضريبة) أقل من الخصومات اللاحقة المربوطة بها "
                            f"({fmt(disc_sub)}). عدّل أو احذف إشعارات الخصم أولاً.")
        if party[0] != head[0]:  # السندات المربوطة تخص الطرف القديم
            conn.execute("UPDATE cash_transactions SET alloc_type=NULL, alloc_id=NULL WHERE alloc_type=? AND alloc_id=?",
                         (kind, inv_id))
        if sales:
            conn.execute("UPDATE sales_invoices SET invoice_date=?, customer_code=?, customer_name=?, notes=?, subtotal=?, "
                         "vat_amount=?, total_amount=?, discount_amount=? WHERE invoice_id=?",
                         (header["date"], party[0], party[1], header.get("notes", ""), subtotal, vat, round(subtotal + vat, 2),
                          round(sum(L["discount"] for L in built), 2), inv_id))
        else:
            conn.execute("UPDATE purchase_invoices SET purchase_date=?, supplier_code=?, supplier_name=?, notes=?, "
                         "supplier_ref=?, subtotal=?, vat_amount=?, total_amount=? WHERE purchase_id=?",
                         (header["date"], party[0], party[1], header.get("notes", ""), header.get("supplier_ref", ""),
                          subtotal, vat, round(subtotal + vat, 2), inv_id))
        refresh_invoice_paid(conn, kind, inv_id)
        no = conn.execute(f"SELECT COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()[0]
        _audit_lines(conn, "تعديل", f"بنود فاتورة {'مبيعات' if sales else 'مشتريات'}", no, c["ilines"], c["iid"], inv_id,
                     before_lines)
        moved = []
        if not sales:  # إعادة توزيع خصومات «تخفيض التكلفة» على بنود الفاتورة بعد تعديلها
            moved = reallocate_invoice_discounts(conn, inv_id)
        recalc_items(conn, [r[1] for r in old.values()] + [L["code"] for L in built] + moved, allow_negative)
    return round(subtotal + vat, 2)


def save_return(r_type, r_date, party, item, qty, price, notes="", return_id=None, allow_negative=False):
    """تسجيل مرتجع جديد أو تعديل مرتجع قائم (return_id)، ثم إعادة احتساب رصيد وتكلفة الصنف."""
    if r_type not in (SALES_RETURN, PURCHASE_RETURN):
        raise UserError("نوع المرتجع غير صحيح.")
    if not parse_date(r_date or ""):
        raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
    if qty is None or qty <= 0 or price is None or price <= 0:
        raise UserError("أدخل كمية وسعراً صحيحين أكبر من الصفر.")
    qty = round(qty, 3)
    p_code, p_name = party
    i_code, i_name = item[0], item[1]
    subtotal = round(qty * price, 2)
    vat = round(subtotal * VAT_RATE, 2)
    total = round(subtotal + vat, 2)
    with db() as conn:
        old_item = None
        if return_id is not None:
            old = conn.execute("SELECT item_code FROM returns_invoices WHERE return_id=?", (return_id,)).fetchone()
            if old is None:
                raise UserError("المرتجع غير موجود.")
            old_item = old[0]
        if conn.execute("SELECT 1 FROM inventory WHERE item_code=?", (i_code,)).fetchone() is None:
            raise UserError(f"الصنف {i_code} غير موجود بالمخزن.")
        excl = return_id if return_id is not None else -1
        returned = conn.execute("SELECT COALESCE(SUM(quantity),0) FROM returns_invoices WHERE return_type=? AND party_code=? "
                                "AND item_code=? AND return_id != ?", (r_type, p_code, i_code, excl)).fetchone()[0]
        if r_type == SALES_RETURN:
            sold, cost_sum = conn.execute(
                "SELECT COALESCE(SUM(l.quantity),0), COALESCE(SUM(l.quantity*l.unit_cost),0) FROM sales_invoice_lines l "
                "JOIN sales_invoices h ON h.invoice_id=l.invoice_id WHERE h.customer_code=? AND l.item_code=?",
                (p_code, i_code)).fetchone()
            if qty > sold - returned + EPS:
                raise UserError(f"الكمية المرتجعة أكبر من المسموح.\nالمباع لهذا العميل: {sold:g}   |   "
                                f"سبق إرجاعه: {returned:g}   |   المتاح للإرجاع: {sold - returned:g}")
            avg = conn.execute("SELECT cost_price FROM inventory WHERE item_code=?", (i_code,)).fetchone()[0] or 0.0
            ret_cost = (cost_sum / sold) if sold > 0 else avg  # تكلفة البيع الأصلية
        else:
            bought = conn.execute(
                "SELECT COALESCE(SUM(l.quantity),0) FROM purchase_invoice_lines l JOIN purchase_invoices h "
                "ON h.purchase_id=l.purchase_id WHERE h.supplier_code=? AND l.item_code=?", (p_code, i_code)).fetchone()[0]
            if qty > bought - returned + EPS:
                raise UserError(f"الكمية المرتجعة أكبر من المسموح.\nالمشترى من هذا المورد: {bought:g}   |   "
                                f"سبق إرجاعه: {returned:g}   |   المتاح للإرجاع: {bought - returned:g}")
            ret_cost = 0.0  # يُحدَّد بمتوسط التكلفة وقت المرتجع عند إعادة الاحتساب
        vals = (r_date, r_type, p_code, p_name, i_code, i_name, qty, price, ret_cost, subtotal, vat, total, notes)
        if return_id is None:
            return_id = conn.execute(
                "INSERT INTO returns_invoices (return_date, return_type, party_code, party_name, item_code, item_name, "
                "quantity, unit_price, unit_cost, subtotal, vat_amount, total_amount, notes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                vals).lastrowid
        else:
            conn.execute("UPDATE returns_invoices SET return_date=?, return_type=?, party_code=?, party_name=?, item_code=?, "
                         "item_name=?, quantity=?, unit_price=?, unit_cost=?, subtotal=?, vat_amount=?, total_amount=?, notes=? "
                         "WHERE return_id=?", vals + (return_id,))
        recalc_items(conn, [i_code, old_item], allow_negative)
    return return_id, total


def delete_return(return_id, allow_negative=False):
    with db() as conn:
        r = conn.execute("SELECT item_code FROM returns_invoices WHERE return_id=?", (return_id,)).fetchone()
        if r is None:
            raise UserError("المرتجع غير موجود.")
        conn.execute("DELETE FROM returns_invoices WHERE return_id=?", (return_id,))
        recalc_items(conn, [r[0]], allow_negative)


# ============================================================
# الخصم اللاحق بعد الشراء / بعد البيع (إشعارات الخصم والإشعارات الدائنة) - آلية متماثلة للموردين والعملاء
# ============================================================
DISC_PURCHASE, DISC_SALES = "purchase", "sales"
DISC_STOCK, DISC_PL = "stock", "pl"
DISC_KIND_LABEL = {DISC_PURCHASE: "إشعار خصم من مورد (خصم لاحق على المشتريات)",
                   DISC_SALES: "إشعار دائن لعميل (خصم لاحق على المبيعات)"}
DISC_KIND_SHORT = {DISC_PURCHASE: "خصم لاحق من مورد", DISC_SALES: "إشعار دائن لعميل"}
DISC_METHOD_LABEL = {DISC_STOCK: "تخفيض تكلفة المخزون (البضاعة الموجودة)",
                     DISC_PL: "خصم مباشر بقائمة الدخل"}
DISC_SEQ = {DISC_PURCHASE: "purchase_discount", DISC_SALES: "sales_discount"}
DISC_SOURCE = "post_discount"
# حسابات القيد التلقائي (قابلة للتغيير من الإعدادات): المفتاح ← (الكود الافتراضي، الوصف)
DISC_ACCOUNT_DEFAULTS = {
    "disc_acc_ap": ("22101", "الموردين (ذمم دائنة)"),
    "disc_acc_ar": ("11250", "العملاء (ذمم مدينة)"),
    "disc_acc_stock": ("11310", "المخزون"),
    "disc_acc_earned": ("41301", "خصم مكتسب"),
    "disc_acc_allowed": ("31300", "خصم مسموح به"),
    "disc_acc_vat_in": ("22501", "ضريبة القيمة المضافة - مدخلات (مدينة)"),
    "disc_acc_vat_out": ("22502", "ضريبة القيمة المضافة - مخرجات (دائنة)"),
}


def invoice_discount_total(conn, kind, inv_id):
    """إجمالي الخصومات اللاحقة (شامل الضريبة) المربوطة بفاتورة."""
    try:
        return conn.execute("SELECT COALESCE(SUM(total_amount),0) FROM post_discounts WHERE kind=? AND invoice_id=?",
                            (kind, inv_id)).fetchone()[0] or 0.0
    except sqlite3.OperationalError:  # قاعدة قديمة قبل إنشاء الجدول
        return 0.0


def disc_accounts(conn):
    out = {}
    for key, (default, _) in DISC_ACCOUNT_DEFAULTS.items():
        r = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        out[key] = (r[0] if r and r[0] else default)
    return out


def set_disc_accounts(mapping):
    """حفظ حسابات قيود الخصم اللاحق بعد التحقق من وجودها وأنها حسابات فرعية (غير تجميعية)، ثم إعادة بناء القيود."""
    with db() as conn:
        for key, code in mapping.items():
            if key not in DISC_ACCOUNT_DEFAULTS:
                continue
            code = (code or "").strip()
            acc = conn.execute("SELECT COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?", (code,)).fetchone()
            label = DISC_ACCOUNT_DEFAULTS[key][1]
            if acc is None:
                raise UserError(f"حساب «{label}»: الكود {code} غير موجود في دليل الحسابات.")
            if acc[0]:
                raise UserError(f"حساب «{label}»: {code} حساب تجميعي. اختر حساباً فرعياً.")
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?,?)", (key, code))
        repost_all_discount_journals(conn)


def _discount_journal_lines(conn, d):
    """d = قاموس رأس الإشعار. يرجع [(كود، مدين، دائن، بيان)].
    مورد: من ح/ الموردين (الإجمالي) إلى ح/ المخزون (الممتص) + ح/ الخصم المكتسب (الباقي) + ح/ ضريبة المدخلات.
    عميل: من ح/ الخصم المسموح به + ح/ ضريبة المخرجات إلى ح/ العملاء (الإجمالي)."""
    acc = disc_accounts(conn)
    lines = []
    if d["kind"] == DISC_PURCHASE:
        stock = round(d["stock_applied"] if d["method"] == DISC_STOCK else 0.0, 2)
        earned = round(d["subtotal"] - stock, 2)
        lines.append((acc["disc_acc_ap"], d["total_amount"], 0.0, f"تخفيض مستحقات المورد {d['party_name']}"))
        if stock > 0.005:
            lines.append((acc["disc_acc_stock"], 0.0, stock, "تخفيض تكلفة المخزون المتبقي"))
        if earned > 0.005:
            lines.append((acc["disc_acc_earned"], 0.0, earned, "خصم مكتسب من المورد"))
        if d["vat_amount"] > 0.005:
            lines.append((acc["disc_acc_vat_in"], 0.0, d["vat_amount"], "تسوية ضريبة المدخلات"))
    else:
        lines.append((acc["disc_acc_allowed"], d["subtotal"], 0.0, f"خصم مسموح به للعميل {d['party_name']}"))
        if d["vat_amount"] > 0.005:
            lines.append((acc["disc_acc_vat_out"], d["vat_amount"], 0.0, "تسوية ضريبة المخرجات"))
        lines.append((acc["disc_acc_ar"], 0.0, d["total_amount"], f"تخفيض مديونية العميل {d['party_name']}"))
    return lines


_DISC_KEYS = ("discount_id", "discount_no", "kind", "discount_date", "party_code", "party_name", "invoice_id", "method",
              "subtotal", "vat_amount", "total_amount", "stock_applied", "journal_entry_id", "reason", "notes",
              "created_by", "created_at")


def _discount_row(conn, discount_id):
    r = conn.execute("SELECT discount_id, discount_no, kind, discount_date, party_code, party_name, invoice_id, method, "
                     "subtotal, vat_amount, total_amount, stock_applied, journal_entry_id, COALESCE(reason,''), "
                     "COALESCE(notes,''), COALESCE(created_by,''), created_at FROM post_discounts WHERE discount_id=?",
                     (discount_id,)).fetchone()
    return dict(zip(_DISC_KEYS, r)) if r else None


def post_discount_journal(conn, discount_id):
    """إنشاء/تحديث القيد اليومي التلقائي للإشعار. يرجع رقم القيد، أو None لو حسابات القيد غير معرّفة بالدليل
    (الأرصدة والتقارير تبقى صحيحة لأنها تُحسب من الإشعار نفسه)."""
    d = _discount_row(conn, discount_id)
    if d is None:
        return None
    lines = _discount_journal_lines(conn, d)
    names = {}
    for code, *_ in lines:
        acc = conn.execute("SELECT account_name, COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?",
                           (code,)).fetchone()
        if acc is None or acc[1]:
            names = None
            break
        names[code] = acc[0]
    entry_id = d["journal_entry_id"]
    if entry_id and conn.execute("SELECT 1 FROM journal_entries WHERE entry_id=?", (entry_id,)).fetchone() is None:
        entry_id = None
    if names is None:
        if entry_id:
            conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (entry_id,))
            conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (entry_id,))
        conn.execute("UPDATE post_discounts SET journal_entry_id=NULL WHERE discount_id=?", (discount_id,))
        return None
    desc = f"قيد آلي: {DISC_KIND_SHORT[d['kind']]} رقم {d['discount_no']} - {d['party_name']}"
    if entry_id is None:
        entry_id = conn.execute("INSERT INTO journal_entries (entry_date, description, source_type, source_id) "
                                "VALUES (?,?,?,?)", (d["discount_date"], desc, DISC_SOURCE, discount_id)).lastrowid
        conn.execute("UPDATE post_discounts SET journal_entry_id=? WHERE discount_id=?", (entry_id, discount_id))
    else:
        conn.execute("UPDATE journal_entries SET entry_date=?, description=?, source_type=?, source_id=? WHERE entry_id=?",
                     (d["discount_date"], desc, DISC_SOURCE, discount_id, entry_id))
        conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (entry_id,))
    for code, debit, credit, memo in lines:
        conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                     "VALUES (?,?,?,?,?,?)", (entry_id, code, names[code], round(debit, 2), round(credit, 2), memo))
    return entry_id


def repost_all_discount_journals(conn):
    for (did,) in conn.execute("SELECT discount_id FROM post_discounts").fetchall():
        post_discount_journal(conn, did)


def refresh_discount_headers(conn, line_ids):
    """بعد إعادة احتساب المخزون: تحديث الجزء الممتص بالمخزون لكل إشعار وإعادة بناء قيده."""
    ids = list(line_ids)
    dids = set()
    for i in range(0, len(ids), 500):
        chunk = ids[i:i + 500]
        ph = ",".join("?" * len(chunk))
        dids.update(r[0] for r in conn.execute(
            f"SELECT DISTINCT discount_id FROM post_discount_lines WHERE line_id IN ({ph})", chunk))
    for did in dids:
        conn.execute("UPDATE post_discounts SET stock_applied=(SELECT ROUND(COALESCE(SUM(applied),0),2) "
                     "FROM post_discount_lines WHERE discount_id=?) WHERE discount_id=?", (did, did))
        post_discount_journal(conn, did)


def _allocate_discount(conn, invoice_id, item, subtotal):
    """توزيع خصم «تخفيض التكلفة» على الأصناف: نسبةً لقيمة بنود فاتورة الشراء، أو على صنف واحد محدد."""
    if invoice_id:
        rows = conn.execute("SELECT item_code, MAX(item_name), SUM(subtotal) FROM purchase_invoice_lines WHERE purchase_id=? "
                            "GROUP BY item_code ORDER BY MIN(line_id)", (invoice_id,)).fetchall()
        rows = [r for r in rows if (r[2] or 0) > 0]
        if item:
            rows = [r for r in rows if r[0] == item[0]]
            if not rows:
                raise UserError("الصنف المختار غير موجود في بنود فاتورة الشراء المحددة.")
        base = sum(r[2] for r in rows)
        if base <= 0:
            raise UserError("لا توجد بنود بقيمة في الفاتورة لتوزيع الخصم عليها.")
        out, left = [], round(subtotal, 2)
        for i, (code, name, val) in enumerate(rows):
            amt = left if i == len(rows) - 1 else round(subtotal * val / base, 2)
            left = round(left - amt, 2)
            if amt > 0.004:
                out.append((code, name, amt))
        return out
    if not item:
        raise UserError("طريقة «تخفيض تكلفة المخزون» تحتاج فاتورة شراء مرتبطة أو صنفاً محدداً.")
    r = conn.execute("SELECT item_code, item_name FROM inventory WHERE item_code=?", (item[0],)).fetchone()
    if r is None:
        raise UserError(f"الصنف {item[0]} غير موجود بالمخزن.")
    return [(r[0], r[1], round(subtotal, 2))]


def reallocate_invoice_discounts(conn, purchase_id):
    """بعد تعديل فاتورة شراء: إعادة توزيع خصوماتها (تخفيض التكلفة) على بنودها الجديدة. يرجع أكواد الأصناف المتأثرة."""
    codes = []
    for did, sub in conn.execute("SELECT discount_id, subtotal FROM post_discounts WHERE kind='purchase' AND method='stock' "
                                 "AND invoice_id=?", (purchase_id,)).fetchall():
        old = [r[0] for r in conn.execute("SELECT item_code FROM post_discount_lines WHERE discount_id=?", (did,))]
        single = set(old)
        item = (next(iter(single)),) if len(single) == 1 else None
        try:
            alloc = _allocate_discount(conn, purchase_id, item, sub)
        except UserError:
            alloc = _allocate_discount(conn, purchase_id, None, sub)
        codes += old
        conn.execute("DELETE FROM post_discount_lines WHERE discount_id=?", (did,))
        for code, name, amt in alloc:
            conn.execute("INSERT INTO post_discount_lines (discount_id, item_code, item_name, amount) VALUES (?,?,?,?)",
                         (did, code, name, amt))
            codes.append(code)
    return codes


def save_post_discount(kind, d_date, party, subtotal, apply_vat=True, invoice_id=None, method=DISC_PL, item=None,
                       reason="", notes="", discount_id=None, user="", allow_closed=False):
    """تسجيل أو تعديل خصم لاحق.
    kind: purchase (إشعار خصم من مورد يخفض ما علينا له) | sales (إشعار دائن لعميل يخفض ما عليه لنا).
    subtotal: قيمة الخصم قبل الضريبة. apply_vat: تسوية ضريبة القيمة المضافة 15% (تُلغى تلقائياً لو الفاتورة بدون ضريبة).
    method (للمشتريات فقط): stock = تخفيض تكلفة المخزون | pl = خصم مكتسب مباشر. خصم المبيعات لا يمس المخزون أبداً.
    يرجع (المعرّف، الرقم التسلسلي، الإجمالي شامل الضريبة)."""
    if kind not in (DISC_PURCHASE, DISC_SALES):
        raise UserError("نوع الخصم غير صحيح.")
    d_date = parse_date(d_date or "")
    if not d_date:
        raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
    if subtotal is None or subtotal <= 0:
        raise UserError("قيمة الخصم يجب أن تكون أكبر من الصفر.")
    subtotal = round(subtotal, 2)
    if not party or not party[0]:
        raise UserError("اختر المورد/العميل.")
    method = DISC_PL if kind == DISC_SALES else (method if method in (DISC_STOCK, DISC_PL) else DISC_PL)
    invoice_id = int(invoice_id) if invoice_id else None
    if period_is_closed(d_date) and not allow_closed:
        raise UserError(f"يوم {d_date} مغلق (تم الإغلاق اليومي). لا يمكن التسجيل بتاريخه.")
    c = FLOW[kind]
    with db() as conn:
        p = conn.execute("SELECT name, is_customer, is_supplier FROM business_partners WHERE partner_code=?",
                         (party[0],)).fetchone()
        if p is None:
            raise UserError("الطرف غير موجود.")
        if kind == DISC_PURCHASE and not p[2]:
            raise UserError(f"«{p[0]}» غير معرّف كمورد.")
        if kind == DISC_SALES and not p[1]:
            raise UserError(f"«{p[0]}» غير معرّف كعميل.")
        party_name = p[0]
        old = None
        if discount_id is not None:
            old = _discount_row(conn, discount_id)
            if old is None or old["kind"] != kind:
                raise UserError("الإشعار غير موجود.")
            if period_is_closed(old["discount_date"]) and not allow_closed:
                raise UserError(f"الإشعار مسجل بيوم مغلق ({old['discount_date']}) ولا يمكن تعديله.")
        vat = round(subtotal * VAT_RATE, 2) if apply_vat else 0.0
        if invoice_id:
            inv = conn.execute(f"SELECT {c['pcode']}, {c['idate']}, subtotal, vat_amount, COALESCE(invoice_no, {c['iid']}) "
                               f"FROM {c['inv']} WHERE {c['iid']}=?", (invoice_id,)).fetchone()
            if inv is None:
                raise UserError("الفاتورة المرتبطة غير موجودة.")
            if inv[0] != party[0]:
                raise UserError("الفاتورة المختارة لا تخص هذا الطرف.")
            if d_date < inv[1]:
                raise UserError(f"تاريخ الخصم ({d_date}) قبل تاريخ الفاتورة ({inv[1]}).")
            prev = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM post_discounts WHERE kind=? AND invoice_id=? "
                                "AND discount_id != ?", (kind, invoice_id, discount_id or -1)).fetchone()[0]
            if prev + subtotal > (inv[2] or 0) + 0.005:
                raise UserError(f"إجمالي الخصومات على الفاتورة {inv[4]} سيتجاوز قيمتها.\n"
                                f"قيمة الفاتورة قبل الضريبة: {fmt(inv[2])}   |   خصومات سابقة: {fmt(prev)}   |   "
                                f"المتاح: {fmt(max((inv[2] or 0) - prev, 0))}")
            if (inv[3] or 0) <= 0.005:
                vat = 0.0  # الفاتورة الأصلية بدون ضريبة
        total = round(subtotal + vat, 2)
        alloc = (_allocate_discount(conn, invoice_id, item, subtotal) if method == DISC_STOCK else [])
        old_codes = []
        if old is None:
            no = next_serial(conn, DISC_SEQ[kind])
            discount_id = conn.execute(
                "INSERT INTO post_discounts (discount_no, kind, discount_date, party_code, party_name, invoice_id, method, "
                "subtotal, vat_amount, total_amount, reason, notes, created_by) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (no, kind, d_date, party[0], party_name, invoice_id, method, subtotal, vat, total, reason, notes, user)
            ).lastrowid
        else:
            no = old["discount_no"]
            old_codes = [r[0] for r in conn.execute("SELECT item_code FROM post_discount_lines WHERE discount_id=?",
                                                    (discount_id,))]
            conn.execute("DELETE FROM post_discount_lines WHERE discount_id=?", (discount_id,))
            conn.execute("UPDATE post_discounts SET discount_date=?, party_code=?, party_name=?, invoice_id=?, method=?, "
                         "subtotal=?, vat_amount=?, total_amount=?, stock_applied=0, reason=?, notes=? WHERE discount_id=?",
                         (d_date, party[0], party_name, invoice_id, method, subtotal, vat, total, reason, notes, discount_id))
        for code, name, amt in alloc:
            conn.execute("INSERT INTO post_discount_lines (discount_id, item_code, item_name, amount) VALUES (?,?,?,?)",
                         (discount_id, code, name, amt))
        codes = old_codes + [a[0] for a in alloc]
        if codes:  # الخصم لا يغير الكميات، فلا يُمنع بسبب أرصدة سالبة قائمة أصلاً
            recalc_items(conn, codes, allow_negative=True)
        post_discount_journal(conn, discount_id)
        for inv_id in {invoice_id, old["invoice_id"] if old else None} - {None}:
            refresh_invoice_paid(conn, kind, inv_id)
    return discount_id, no, total


def delete_post_discount(discount_id, allow_closed=False):
    with db() as conn:
        d = _discount_row(conn, discount_id)
        if d is None:
            raise UserError("الإشعار غير موجود.")
        if period_is_closed(d["discount_date"]) and not allow_closed:
            raise UserError(f"الإشعار مسجل بيوم مغلق ({d['discount_date']}) ولا يمكن حذفه.")
        codes = [r[0] for r in conn.execute("SELECT item_code FROM post_discount_lines WHERE discount_id=?", (discount_id,))]
        conn.execute("DELETE FROM journal_lines WHERE entry_id IN (SELECT entry_id FROM journal_entries "
                     "WHERE source_type=? AND source_id=?)", (DISC_SOURCE, discount_id))
        conn.execute("DELETE FROM journal_entries WHERE source_type=? AND source_id=?", (DISC_SOURCE, discount_id))
        conn.execute("DELETE FROM post_discount_lines WHERE discount_id=?", (discount_id,))
        conn.execute("DELETE FROM post_discounts WHERE discount_id=?", (discount_id,))
        if codes:
            recalc_items(conn, codes, allow_negative=True)
        if d["invoice_id"]:
            refresh_invoice_paid(conn, d["kind"], d["invoice_id"])


def load_post_discount(discount_id):
    """(رأس الإشعار كقاموس + invoice_no، بنود التوزيع [(كود، اسم، المبلغ، الممتص بالمخزون)])."""
    with db() as conn:
        d = _discount_row(conn, discount_id)
        if d is None:
            raise UserError("الإشعار غير موجود (ربما حُذف).")
        lines = conn.execute("SELECT item_code, COALESCE(item_name,''), amount, applied FROM post_discount_lines "
                             "WHERE discount_id=? ORDER BY line_id", (discount_id,)).fetchall()
        d["invoice_no"] = ""
        if d["invoice_id"]:
            c = FLOW[d["kind"]]
            r = conn.execute(f"SELECT COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} WHERE {c['iid']}=?",
                             (d["invoice_id"],)).fetchone()
            d["invoice_no"] = str(r[0]) if r else str(d["invoice_id"])
    return d, lines


def list_post_discounts(kind=None, party_code=None, d_from=None, d_to=None):
    """سجل الإشعارات: [(id, no, kind, date, party_code, party_name, invoice_no, method, subtotal, vat, total, applied, reason)]."""
    sql = ("SELECT d.discount_id, d.discount_no, d.kind, d.discount_date, d.party_code, d.party_name, "
           "COALESCE(CASE WHEN d.kind='sales' THEN (SELECT COALESCE(invoice_no, invoice_id) FROM sales_invoices "
           "WHERE invoice_id=d.invoice_id) ELSE (SELECT COALESCE(invoice_no, purchase_id) FROM purchase_invoices "
           "WHERE purchase_id=d.invoice_id) END, ''), d.method, d.subtotal, d.vat_amount, d.total_amount, d.stock_applied, "
           "COALESCE(d.reason,'') FROM post_discounts d WHERE 1=1")
    params = []
    for cond, val in (("d.kind=?", kind), ("d.party_code=?", party_code), ("d.discount_date>=?", d_from),
                      ("d.discount_date<=?", d_to)):
        if val:
            sql += f" AND {cond}"
            params.append(val)
    with db() as conn:
        return conn.execute(sql + " ORDER BY d.discount_date DESC, d.discount_id DESC", params).fetchall()


def discount_invoice_choices(kind, party_code):
    """فواتير الطرف المتاحة للربط: [(id, no, date, subtotal, total, خصومات سابقة قبل الضريبة)]."""
    c = FLOW[kind]
    with db() as conn:
        return conn.execute(
            f"SELECT h.{c['iid']}, COALESCE(h.invoice_no, h.{c['iid']}), h.{c['idate']}, h.subtotal, h.total_amount, "
            f"COALESCE((SELECT SUM(subtotal) FROM post_discounts d WHERE d.kind=? AND d.invoice_id=h.{c['iid']}),0) "
            f"FROM {c['inv']} h WHERE h.{c['pcode']}=? ORDER BY h.{c['idate']} DESC, h.{c['iid']} DESC",
            (kind, party_code)).fetchall()


HW_SOURCE = "HansaWorld"


def hw_cogs(conn, d_from, d_to, by_month=False):
    """تكلفة مبيعات الفترة التاريخية المستوردة من HansaWorld بطريقة الجرد الدوري، لأن فواتيرها بالإجمالي بلا بنود أصناف
    (فلا توجد تكلفة صنف لكل بيعة):
        تكلفة المبيعات = مشتريات HansaWorld − إشعارات خصم الموردين المستوردة − المخزون الختامي في تاريخ الترحيل
    (مخزون أول مدة HansaWorld = صفر لأن الملفات تبدأ من الصفر). المخزون الختامي = الرصيد الافتتاحي للأصناف في النظام
    ويُطرح في الفترة/الشهر الذي يقع فيه تاريخ الترحيل. إشعارات الموردين المستوردة تدخل هنا (تسوية مشتريات) فتُستبعد من
    «الخصم المكتسب» في الإيرادات الأخرى حتى لا تُحسب مرتين.
    يرجع {مفتاح الشهر أو '': dict(purchases, credits, closing, cogs)} — فارغ لو لم يُستورد شيء من HansaWorld."""
    out = defaultdict(lambda: dict(purchases=0.0, credits=0.0, closing=0.0, ref_adj=0.0, cogs=0.0))
    try:
        cutover = conn.execute("SELECT MAX(doc_date) FROM hw_import_map WHERE role IN ('customer','supplier')").fetchone()[0]
    except sqlite3.OperationalError:
        return out
    if not cutover:
        return out
    key = "substr({},1,7)" if by_month else "''"
    for k, v in conn.execute(f"SELECT {key.format('purchase_date')}, SUM(subtotal) FROM purchase_invoices "
                             f"WHERE source_type=? AND purchase_date BETWEEN ? AND ? GROUP BY 1", (HW_SOURCE, d_from, d_to)):
        out[k]["purchases"] += v or 0.0
    for k, v in conn.execute(f"SELECT {key.format('discount_date')}, SUM(subtotal) FROM post_discounts WHERE kind='purchase' "
                             f"AND discount_date BETWEEN ? AND ? AND discount_id IN (SELECT target_id FROM hw_import_map "
                             f"WHERE role='supplier' AND target='post_discounts') GROUP BY 1", (d_from, d_to)):
        out[k]["credits"] += v or 0.0
    if d_from <= cutover <= d_to:
        (closing,) = conn.execute("SELECT COALESCE(SUM(opening_qty * opening_cost), 0) FROM inventory").fetchone()
        out[cutover[:7] if by_month else ""]["closing"] += closing or 0.0
    for a in hw_cogs_ref_adjustments(conn):
        if d_from <= a["date"] <= d_to and abs(a["adj"]) > 0.005:
            out[a["date"][:7] if by_month else ""]["ref_adj"] += a["adj"]
    for o in out.values():
        o["cogs"] = round(o["purchases"] - o["credits"] - o["closing"] + o["ref_adj"], 2)
    return out


# ============================================================
# المراجع المحاسبية من HansaWorld (ميزان المراجعة + قوائم الدخل لفترات) — تقسيم فترة HansaWorld إلى فترات مرجعية
# ============================================================
HW_COGS_PREFIXES = ("4111", "412")   # تكلفة المبيعات 41110 ومردود/فروق المشتريات 41201 في دفتر HansaWorld
HW_SALES_PREFIX, HW_OTHER_REV_PREFIX = "31", "33"


def _safe_rows(conn, sql, params=()):
    try:
        return conn.execute(sql, params).fetchall()
    except sqlite3.OperationalError:
        return []


def _shift_day(d, days):
    return (datetime.strptime(d, "%Y-%m-%d").date() + timedelta(days=days)).strftime("%Y-%m-%d")


def hw_ref_segments(conn):
    """فترات HansaWorld المرجعية لحسابات قائمة الدخل (3xxxx/4xxxx/5xxxx):
    - ميزان المراجعة التحليلي = تراكمي من أول مستند HansaWorld حتى تاريخ الميزان.
    - كل قائمة دخل مستوردة لفترة (مثل 2026-01-01 → 2026-11-19) = فترة بأرقامها كما هي.
    - الباقي (الميزان − مجموع القوائم) = الفترة غير المغطاة (مثل 2025-07-31 → 2025-12-31).
    يرجع [dict(d_from, d_to, src: tb|pl|tb_rest, amounts={الكود: (الاسم، المبلغ مدين موجب)})] مرتبة، أو [] بلا ميزان."""
    tb = [r for r in _safe_rows(conn, "SELECT account_code, account_name, balance, as_of FROM hw_tb_reference")
          if str(r[0])[:1] in ("3", "4", "5")]
    if not tb:
        return []
    tb_to = max((r[3] or "" for r in tb), default="") or today()
    first = _safe_rows(conn, "SELECT MIN(doc_date) FROM hw_import_map WHERE role IN ('customer','supplier')")
    tb_from = min((first[0][0] if first and first[0][0] else None) or "1900-01-01", tb_to)
    pl = defaultdict(dict)
    for f, t, code, name, amt in _safe_rows(conn, "SELECT period_from, period_to, account_code, account_name, amount "
                                                  "FROM hw_pl_reference"):
        pl[(f, t)][code] = (name, round(amt or 0.0, 2))
    segs = [dict(d_from=f, d_to=t, src="pl", amounts=acc) for (f, t), acc in sorted(pl.items())]
    rest = {code: (name, round(bal or 0.0, 2)) for code, name, bal, _ in tb}
    for sg in segs:
        for code, (name, amt) in sg["amounts"].items():
            n0, b0 = rest.get(code, (name, 0.0))
            rest[code] = (n0, round(b0 - amt, 2))
    gaps, cur = [], tb_from
    for sg in segs:
        if sg["d_from"] > cur:
            gaps.append((cur, _shift_day(sg["d_from"], -1)))
        cur = max(cur, _shift_day(sg["d_to"], 1))
    if cur <= tb_to:
        gaps.append((cur, tb_to))
    if gaps:   # باقي الميزان يُنسب لأول فترة غير مغطاة (الحالة المعتادة: فترة واحدة قبل أول قائمة دخل)
        segs.append(dict(d_from=gaps[0][0], d_to=gaps[0][1], src="tb_rest" if pl else "tb",
                         amounts={c: v for c, v in rest.items() if abs(v[1]) > 0.005 or not pl}))
    return sorted(segs, key=lambda x: x["d_from"])


def hw_segment_totals(amounts):
    """تجميع حسابات فترة مرجعية بمفاهيم قائمة الدخل (موجبة): المبيعات، الإيرادات الأخرى، تكلفة المبيعات،
    المصروفات البيعية (43)، الإدارية (5)، أخرى (41125 وغيرها)، مجمل الربح وصافي الربح."""
    t = dict(sales=0.0, other_income=0.0, cogs=0.0, selling=0.0, admin=0.0, other_exp=0.0)
    for code, (_, amt) in amounts.items():
        if code.startswith(HW_SALES_PREFIX):
            t["sales"] -= amt
        elif code.startswith("3"):
            t["other_income"] -= amt
        elif code.startswith(HW_COGS_PREFIXES):
            t["cogs"] += amt
        elif code.startswith("43"):
            t["selling"] += amt
        elif code.startswith("5"):
            t["admin"] += amt
        elif code.startswith("4"):
            t["other_exp"] += amt
    t = {k: round(v, 2) for k, v in t.items()}
    t["opex"] = round(t["selling"] + t["admin"] + t["other_exp"], 2)
    t["gross"] = round(t["sales"] - t["cogs"], 2)
    t["net_profit"] = round(t["gross"] + t["other_income"] - t["opex"], 2)
    return t


def hw_cogs_ref_adjustments(conn):
    """تسوية تكلفة مبيعات كل فترة مرجعية من HansaWorld لتساوي دفتر HansaWorld (41110 + 41201):
    التسوية = تكلفة الدفتر للفترة − تكلفة الجرد الدوري للفترة (مشتريات − إشعارات موردين − المخزون الختامي إن وقع فيها)،
    وتُسجَّل في آخر يوم من الفترة. سببها: الجرد الدوري لا يعرف مخزون بداية الفترة الوسيطة (مثل 2026-01-01)
    ولا البضاعة المستلمة غير المفوترة (مشتريات مستحقة 22380) ولا فروق مردودات المشتريات.
    يرجع [] لو الميزان لا يحتوي حسابات تكلفة المبيعات أو كان تاريخ الترحيل بعد آخر فترة مرجعية."""
    segs = hw_ref_segments(conn)
    if not segs or not any(c.startswith(HW_COGS_PREFIXES) for sg in segs for c in sg["amounts"]):
        return []
    row = _safe_rows(conn, "SELECT MAX(doc_date) FROM hw_import_map WHERE role IN ('customer','supplier')")
    cutover = row[0][0] if row else None
    if not cutover or cutover > segs[-1]["d_to"]:
        return []
    (closing,) = conn.execute("SELECT COALESCE(SUM(opening_qty * opening_cost), 0) FROM inventory").fetchone()
    out = []
    for i, sg in enumerate(segs):
        a, b = ("0000-01-01" if i == 0 else sg["d_from"]), sg["d_to"]
        (pur,) = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM purchase_invoices WHERE source_type=? "
                              "AND purchase_date BETWEEN ? AND ?", (HW_SOURCE, a, b)).fetchone()
        (cr,) = conn.execute("SELECT COALESCE(SUM(subtotal),0) FROM post_discounts WHERE kind='purchase' AND discount_date "
                             "BETWEEN ? AND ? AND discount_id IN (SELECT target_id FROM hw_import_map WHERE role='supplier' "
                             "AND target='post_discounts')", (a, b)).fetchone()
        cl = (closing or 0.0) if a <= cutover <= b else 0.0
        ref = round(sum(amt for c, (_, amt) in sg["amounts"].items() if c.startswith(HW_COGS_PREFIXES)), 2)
        periodic = round(pur - cr - cl, 2)
        out.append(dict(seg=sg, date=b, ref=ref, periodic=periodic, purchases=round(pur, 2), credits=round(cr, 2),
                        closing=round(cl, 2), adj=round(ref - periodic, 2)))
    return out


# ============================================================
# تصنيف المصروفات التشغيلية: بيعية وتسويقية (43) / إدارية وعمومية (5 وسندات المصروفات العامة) / أخرى (41125 …)
# ============================================================
OPEX_GROUP_LABELS = {"selling": "مصروفات بيعية وتسويقية", "admin": "مصروفات إدارية وعمومية",
                     "other": "مصروفات أخرى (إيجار سيارات وغيرها)"}


def opex_group_of(code):
    code = str(code or "")
    if code.startswith("43"):
        return "selling"
    if code.startswith("4") and not code.startswith("40"):
        return "other"
    return "admin"


def expense_groups(conn, d_from, d_to):
    """{selling, admin, other}: نفس مصادر expense_breakdown (سندات الصرف + القيود) مصنفة بكود الحساب."""
    from . import ledger
    out = dict(selling=0.0, admin=0.0, other=0.0)
    chart = {c for (c,) in conn.execute("SELECT account_code FROM chart_of_accounts")}
    for code, name, amt in conn.execute(
            "SELECT party_code, party_name, SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND category=? "
            "AND trans_date BETWEEN ? AND ? GROUP BY party_code, party_name", (CAT_EXPENSE, d_from, d_to)):
        acc = code if code in chart else ledger._expense_by_name(conn, name)
        out[opex_group_of(acc)] += amt or 0.0
    for _, typ, code, _, amt in journal_pl(conn, d_from, d_to):
        if typ == TYPE_EXP:
            out[opex_group_of(code)] += amt
    return out


def discount_totals(conn, d_from, d_to, by_month=False):
    """أثر الخصومات اللاحقة على قائمة الدخل والضريبة: {مفتاح الشهر أو '': dict(...)}.
    sales_sub/sales_vat = إشعارات العملاء (تخفض صافي المبيعات وضريبة المخرجات).
    pur_sub/pur_vat = إشعارات الموردين (تخفض ضريبة المدخلات). earned = الخصم المكتسب الظاهر بالأرباح
    (الطريقة المباشرة + الجزء من خصم التكلفة الذي يخص بضاعة سبق بيعها). stock = الجزء الذي خفّض تكلفة المخزون."""
    key = "substr(discount_date,1,7)" if by_month else "''"
    out = defaultdict(lambda: dict(sales_sub=0.0, sales_vat=0.0, pur_sub=0.0, pur_vat=0.0, earned=0.0, stock=0.0))
    for k, kind, method, sub, vat, applied in conn.execute(
            f"SELECT {key}, kind, method, SUM(subtotal), SUM(vat_amount), SUM(stock_applied) FROM post_discounts "
            f"WHERE discount_date BETWEEN ? AND ? GROUP BY 1, 2, 3", (d_from, d_to)):
        o = out[k]
        sub, vat, applied = sub or 0.0, vat or 0.0, applied or 0.0
        if kind == DISC_SALES:
            o["sales_sub"] += sub
            o["sales_vat"] += vat
        else:
            o["pur_sub"] += sub
            o["pur_vat"] += vat
            stock = applied if method == DISC_STOCK else 0.0
            o["stock"] += stock
            o["earned"] += sub - stock
    return out


def post_discount_register(kind=None, d_from=None, d_to=None, party_code=None):
    """سجل الخصومات اللاحقة مع أثر كل إشعار: على رصيد الطرف، الضريبة، تكلفة المخزون، وقائمة الدخل."""
    out = []
    for (did, no, k, d, pcode, pname, inv_no, method, sub, vat, tot, applied, reason) in list_post_discounts(
            kind, party_code, d_from, d_to):
        stock = applied if (k == DISC_PURCHASE and method == DISC_STOCK) else 0.0
        pl = (sub - stock) if k == DISC_PURCHASE else -sub
        with db() as conn:
            je = conn.execute("SELECT journal_entry_id FROM post_discounts WHERE discount_id=?", (did,)).fetchone()[0]
        out.append(dict(id=did, no=no, kind=k, date=d, party_code=pcode, party=pname, invoice=inv_no or "",
                        method=("إشعار دائن (يخفض المبيعات)" if k == DISC_SALES else
                                ("تخفيض تكلفة المخزون" if method == DISC_STOCK else "خصم مكتسب مباشر")),
                        subtotal=sub, vat=vat, total=tot, stock=stock, pl=pl, journal=je, reason=reason,
                        balance_effect=("يخفض ما للمورد علينا" if k == DISC_PURCHASE else "يخفض ما على العميل"),
                        vat_effect=("يخفض ضريبة المدخلات" if k == DISC_PURCHASE else "يخفض ضريبة المخرجات")))
    return out


def rep_post_discounts(p):
    data = post_discount_register(None, p["d_from"], p["d_to"])
    rows = [(r["no"], r["date"], "مورد" if r["kind"] == DISC_PURCHASE else "عميل", f"{r['party_code']} - {r['party']}",
             r["invoice"] or "-", r["method"], r["subtotal"], r["vat"], r["total"], r["stock"], r["pl"],
             str(r["journal"] or "-")) for r in data]
    cols = [("رقم الإشعار", "t"), ("التاريخ", "t"), ("الطرف", "t"), ("الاسم", "t"), ("الفاتورة", "t"), ("المعالجة", "t"),
            ("قبل الضريبة", "m"), ("الضريبة", "m"), ("الإجمالي", "m"), ("خفض تكلفة المخزون", "m"),
            ("أثر الأرباح (+/-)", "m"), ("القيد", "t")]
    total = ("", "", "", "الإجمالي", "", "", *[sum(r[i] for r in rows) for i in range(6, 11)], "")
    sup = [r for r in data if r["kind"] == DISC_PURCHASE]
    cus = [r for r in data if r["kind"] == DISC_SALES]
    notes = [f"خصومات الموردين: {len(sup)} إشعار بإجمالي {fmt(sum(r['total'] for r in sup))} ر.س "
             f"(منها {fmt(sum(r['stock'] for r in sup))} خفّضت تكلفة المخزون و{fmt(sum(r['pl'] for r in sup))} خصم مكتسب بالأرباح)",
             f"الإشعارات الدائنة للعملاء: {len(cus)} إشعار بإجمالي {fmt(sum(r['total'] for r in cus))} ر.س "
             f"(خفّضت صافي المبيعات بـ {fmt(sum(r['subtotal'] for r in cus))})",
             "أثر الضريبة: خصومات الموردين تخفض ضريبة المدخلات، والإشعارات الدائنة تخفض ضريبة المخرجات."]
    return Report("الخصومات اللاحقة والإشعارات الدائنة وأثرها", cols, rows, total, notes, f"من {p['d_from']} إلى {p['d_to']}",
                  chart=(3, 8))


# ============================================================
# المناديب (سجل تاريخي بالفترات) + المرفقات
# ============================================================
ENTERPRISE_SQL = """
CREATE TABLE IF NOT EXISTS partner_rep_history (
    hist_id INTEGER PRIMARY KEY AUTOINCREMENT,
    partner_code TEXT NOT NULL REFERENCES business_partners(partner_code) ON DELETE CASCADE,
    rep_name TEXT NOT NULL,
    start_date TEXT NOT NULL,
    end_date TEXT,
    changed_by TEXT,
    changed_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_rephist_partner ON partner_rep_history(partner_code, start_date);
CREATE TABLE IF NOT EXISTS attachments (
    attach_id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_type TEXT NOT NULL,
    doc_id INTEGER NOT NULL,
    file_name TEXT NOT NULL,
    mime TEXT,
    size INTEGER NOT NULL DEFAULT 0,
    data BLOB NOT NULL,
    notes TEXT,
    uploaded_by TEXT,
    uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
    file_path TEXT
);
CREATE INDEX IF NOT EXISTS idx_attach_doc ON attachments(doc_type, doc_id);
-- تسوية الدفعات المقدمة / تحت الحساب مع الفواتير اللاحقة (مطابقة فقط: لا قيد جديد، الطرف نفسه في نفس حساب الذمم)
CREATE TABLE IF NOT EXISTS advance_allocations (
    alloc_id INTEGER PRIMARY KEY AUTOINCREMENT,
    trans_id INTEGER NOT NULL,
    kind TEXT NOT NULL CHECK (kind IN ('sales','purchase')),
    invoice_id INTEGER NOT NULL,
    amount REAL NOT NULL CHECK (amount > 0),
    alloc_date TEXT NOT NULL,
    created_by TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_advalloc_inv ON advance_allocations(kind, invoice_id);
CREATE INDEX IF NOT EXISTS idx_advalloc_trans ON advance_allocations(trans_id);
-- الأدوار الوظيفية (مجموعات صلاحيات): حفظ الدور يحدّث صلاحيات كل مستخدميه
CREATE TABLE IF NOT EXISTS user_roles (
    role_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    description TEXT,
    permissions TEXT NOT NULL DEFAULT '',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sales_reps (
    rep_id INTEGER PRIMARY KEY AUTOINCREMENT,
    rep_code TEXT UNIQUE,
    name TEXT NOT NULL UNIQUE,
    phone TEXT,
    email TEXT,
    national_id TEXT,
    region TEXT,
    job_title TEXT,
    hire_date TEXT,
    commission_pct REAL NOT NULL DEFAULT 0,
    active INTEGER NOT NULL DEFAULT 1,
    notes TEXT,
    created_by TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
"""
NO_REP = "(بدون مندوب)"


def _day_before(d):
    return (datetime.strptime(d, "%Y-%m-%d").date() - timedelta(days=1)).strftime("%Y-%m-%d")


def set_partner_rep(conn, code, rep, effective=None, user=""):
    """تعيين/تغيير مندوب الطرف اعتباراً من تاريخ: تُغلق فترة المندوب السابق في اليوم السابق وتُفتح فترة جديدة.
    التغيير بنفس تاريخ بداية الفترة الحالية = تصحيح لها. يرجع True لو تغيّر شيء."""
    rep = ensure_rep(conn, rep, user)
    effective = parse_date(effective or "") or today()
    cur = conn.execute("SELECT hist_id, rep_name, start_date FROM partner_rep_history WHERE partner_code=? AND end_date IS NULL "
                       "ORDER BY start_date DESC, hist_id DESC LIMIT 1", (code,)).fetchone()
    if (cur and cur[1] == rep) or (not cur and not rep):
        conn.execute("UPDATE business_partners SET rep_name=? WHERE partner_code=?", (rep, code))
        return False
    if cur:
        if effective < cur[2]:
            raise UserError(f"تاريخ تغيير المندوب ({effective}) قبل بداية فترة المندوب الحالي «{cur[1]}» ({cur[2]}).")
        if effective == cur[2]:
            conn.execute("DELETE FROM partner_rep_history WHERE hist_id=?", (cur[0],))
            prev = conn.execute("SELECT hist_id, rep_name FROM partner_rep_history WHERE partner_code=? AND end_date=? "
                                "ORDER BY hist_id DESC LIMIT 1", (code, _day_before(effective))).fetchone()
            if prev and prev[1] == rep:  # رجوع لنفس المندوب السابق = إعادة فتح فترته
                conn.execute("UPDATE partner_rep_history SET end_date=NULL WHERE hist_id=?", (prev[0],))
                rep_open = False
            else:
                rep_open = bool(rep)
        else:
            conn.execute("UPDATE partner_rep_history SET end_date=? WHERE hist_id=?", (_day_before(effective), cur[0]))
            rep_open = bool(rep)
    else:
        rep_open = True
    if rep and rep_open:
        conn.execute("INSERT INTO partner_rep_history (partner_code, rep_name, start_date, changed_by) VALUES (?,?,?,?)",
                     (code, rep, effective, user))
    conn.execute("UPDATE business_partners SET rep_name=? WHERE partner_code=?", (rep, code))
    stamp_voucher_reps(conn, code)
    return True


def partner_rep_history(code):
    with db() as conn:
        return conn.execute("SELECT rep_name, start_date, COALESCE(end_date,''), COALESCE(changed_by,''), changed_at "
                            "FROM partner_rep_history WHERE partner_code=? ORDER BY start_date, hist_id", (code,)).fetchall()


# ============================================================
# سجل المناديب (Representatives Directory)
# ============================================================
REP_REF_COLUMNS = [("partner_rep_history", "rep_name"), ("business_partners", "rep_name"), ("sales_orders", "sales_rep"),
                   ("purchase_orders", "purchase_rep"), ("sales_invoices", "sales_rep"),
                   ("purchase_invoices", "purchase_rep"), ("company_profile", "rep_name")]
REP_FIELDS = ["name", "phone", "email", "national_id", "region", "job_title", "hire_date", "commission_pct", "active", "notes"]
REP_LABELS = dict(code="الكود", name="اسم المندوب", phone="رقم الهاتف", email="البريد الإلكتروني", national_id="رقم الهوية / الإقامة",
                  region="المنطقة / خط السير", job_title="المسمى الوظيفي", hire_date="تاريخ التعيين",
                  commission_pct="نسبة العمولة %", active="نشط", notes="ملاحظات")


def _rep_ref_columns(conn):
    out = []
    for t, col in REP_REF_COLUMNS:
        if col in {r[1] for r in conn.execute(f"PRAGMA table_info({t})")}:
            out.append((t, col))
    return out


def ensure_rep(conn, name, user=""):
    """يسجل المندوب في سجل المناديب لو غير موجود (من الاستيراد أو البيانات القديمة). يرجع الاسم النظيف."""
    name = clean_name(name or "")
    if not name or name == NO_REP:
        return ""
    if conn.execute("SELECT 1 FROM sales_reps WHERE name=?", (name,)).fetchone() is None:
        cur = conn.execute("INSERT INTO sales_reps (name, created_by) VALUES (?,?)", (name, user))
        conn.execute("UPDATE sales_reps SET rep_code=? WHERE rep_id=?", (f"R{cur.lastrowid:03d}", cur.lastrowid))
    return name


def seed_reps_from_history(conn):
    """ترقية v11: كل الأسماء المستخدمة سابقاً (بطاقات الأطراف، الأوامر، الفواتير، سجل الفترات) تُضاف لسجل المناديب."""
    for t, col in _rep_ref_columns(conn):
        for (n,) in conn.execute(f"SELECT DISTINCT {col} FROM {t} WHERE COALESCE({col},'')!=''").fetchall():
            ensure_rep(conn, n, "ترقية")


def list_reps(active_only=False, q=""):
    """[(rep_id, code, name, phone, email, national_id, region, job_title, hire_date, commission, active, notes,
    عدد العملاء/الموردين الحاليين)]"""
    qn = norm_name(q or "").strip()
    with db() as conn:
        rows = conn.execute(
            "SELECT r.rep_id, COALESCE(r.rep_code,''), r.name, COALESCE(r.phone,''), COALESCE(r.email,''), "
            "COALESCE(r.national_id,''), COALESCE(r.region,''), COALESCE(r.job_title,''), COALESCE(r.hire_date,''), "
            "r.commission_pct, r.active, COALESCE(r.notes,''), "
            "(SELECT COUNT(*) FROM business_partners b WHERE b.rep_name=r.name) "
            "FROM sales_reps r " + ("WHERE r.active=1 " if active_only else "") + "ORDER BY r.active DESC, r.name").fetchall()
    if qn:
        rows = [r for r in rows if qn in norm_name(" ".join(str(x) for x in r[1:8]))]
    return rows


def rep_names(include=None):
    """أسماء المناديب النشطين من سجل المناديب (للقوائم المنسدلة). include: اسم حالي يُضاف حتى لو موقوف."""
    with db() as conn:
        names = [r[0] for r in conn.execute("SELECT name FROM sales_reps WHERE active=1 ORDER BY name")]
    if include and include not in names:
        names.append(include)
    return names


def save_rep(rep_id, data, user=""):
    """إضافة (rep_id=None) أو تعديل مندوب. تغيير الاسم ينتقل تلقائياً لكل المستندات والسجلات المرتبطة به."""
    d = {k: data.get(k) for k in REP_FIELDS}
    d["name"] = clean_name(d["name"] or "")
    if not d["name"]:
        raise UserError("اسم المندوب مطلوب.")
    if d["name"] == NO_REP:
        raise UserError("اسم غير مسموح.")
    hd = (d.get("hire_date") or "").strip()
    if hd and not parse_date(hd):
        raise UserError("تاريخ التعيين بصيغة YYYY-MM-DD أو اتركه فارغاً.")
    d["hire_date"] = parse_date(hd) if hd else ""
    try:
        d["commission_pct"] = float(normalize_digits(str(d.get("commission_pct") or 0)) or 0)
    except ValueError:
        raise UserError("نسبة العمولة يجب أن تكون رقماً.")
    if not 0 <= d["commission_pct"] <= 100:
        raise UserError("نسبة العمولة بين 0 و 100.")
    d["active"] = 1 if d.get("active") in (None, True, 1, "1") else 0
    for k in ("phone", "email", "national_id", "region", "job_title", "notes"):
        d[k] = (d.get(k) or "").strip()
    with db() as conn:
        dup = conn.execute("SELECT rep_id FROM sales_reps WHERE name=? AND rep_id!=?", (d["name"], rep_id or 0)).fetchone()
        if dup:
            raise UserError(f"يوجد مندوب مسجل بنفس الاسم «{d['name']}».")
        if rep_id:
            old = conn.execute("SELECT name FROM sales_reps WHERE rep_id=?", (rep_id,)).fetchone()
            if not old:
                raise UserError("المندوب غير موجود.")
            conn.execute(f"UPDATE sales_reps SET {', '.join(f'{k}=?' for k in REP_FIELDS)} WHERE rep_id=?",
                         [d[k] for k in REP_FIELDS] + [rep_id])
            if old[0] != d["name"]:
                for t, col in _rep_ref_columns(conn):
                    conn.execute(f"UPDATE {t} SET {col}=? WHERE {col}=?", (d["name"], old[0]))
            return rep_id
        cur = conn.execute(f"INSERT INTO sales_reps ({', '.join(REP_FIELDS)}, created_by) VALUES "
                           f"({', '.join('?' * len(REP_FIELDS))}, ?)", [d[k] for k in REP_FIELDS] + [user])
        rid = cur.lastrowid
        conn.execute("UPDATE sales_reps SET rep_code=? WHERE rep_id=?", (f"R{rid:03d}", rid))
        return rid


def delete_rep(rep_id):
    """حذف مندوب غير مستخدم. المندوب المرتبط بمستندات أو أطراف يُوقف (غير نشط) بدلاً من الحذف للحفاظ على التاريخ.
    يرجع "deleted" أو "deactivated"."""
    with db() as conn:
        r = conn.execute("SELECT name FROM sales_reps WHERE rep_id=?", (rep_id,)).fetchone()
        if not r:
            raise UserError("المندوب غير موجود.")
        used = any(conn.execute(f"SELECT 1 FROM {t} WHERE {col}=? LIMIT 1", (r[0],)).fetchone()
                   for t, col in _rep_ref_columns(conn))
        if used:
            conn.execute("UPDATE sales_reps SET active=0 WHERE rep_id=?", (rep_id,))
            return "deactivated"
        conn.execute("DELETE FROM sales_reps WHERE rep_id=?", (rep_id,))
        return "deleted"


class RepResolver:
    """إسناد الحركة لمندوب حسب فترة تواجده الفعلية مع الطرف بتاريخ الحركة (منع تداخل العمولات).
    الأولوية: سجل الفترات ← مندوب المستند ← المندوب الافتراضي للطرف (لو لا يوجد سجل) ← بدون مندوب."""

    def __init__(self, conn):
        self.hist = defaultdict(list)
        for code, rep, s, e in conn.execute("SELECT partner_code, rep_name, start_date, end_date FROM partner_rep_history"):
            self.hist[code].append((s, e, rep))
        self.default = dict(conn.execute("SELECT partner_code, COALESCE(rep_name,'') FROM business_partners"))

    def __call__(self, code, d, doc_rep=None):
        periods = self.hist.get(code)
        if periods:
            for s, e, rep in periods:
                if s <= d and (not e or e >= d):
                    return rep
            return doc_rep or NO_REP
        return doc_rep or self.default.get(code) or NO_REP


def voucher_rep_info(conn, party_code, t_date, alloc_type=None, alloc_id=None, who=None):
    """المندوب الموروث تلقائياً لسند قبض/صرف لطرف ومصدره (لا يختاره المستخدم يدوياً):
    1) سند مربوط بفاتورة ← مندوب الفاتورة كما نُسبت له المبيعة/المشتراة (فترة المندوب بتاريخ الفاتورة ثم مندوب المستند)
    2) وإلا ← مندوب الطرف بتاريخ السند (سجل فترات المناديب ثم المندوب الحالي في بطاقته)
    3) وإلا ← بدون مندوب (None). يرجع (المندوب أو None، وصف المصدر)."""
    if not party_code:
        return None, ""
    who = who or RepResolver(conn)
    if alloc_type in ("sales", "purchase") and alloc_id:
        c = FLOW[alloc_type]
        col = "sales_rep" if alloc_type == "sales" else "purchase_rep"
        r = conn.execute(f"SELECT {c['pcode']}, {c['idate']}, COALESCE({col},''), COALESCE(invoice_no, {c['iid']}) "
                         f"FROM {c['inv']} WHERE {c['iid']}=?", (alloc_id,)).fetchone()
        if r:
            rep = who(r[0], r[1], r[2] or None)
            if rep != NO_REP:
                return rep, f"من الفاتورة {r[3]}"
    rep = who(party_code, t_date)
    return (None, "الطرف بلا مندوب مسجل") if rep == NO_REP else (rep, "من بطاقة الطرف")


def voucher_rep(conn, party_code, t_date, alloc_type=None, alloc_id=None, who=None):
    return voucher_rep_info(conn, party_code, t_date, alloc_type, alloc_id, who)[0]


def stamp_voucher_reps(conn, party_code=None):
    """توريث المندوب للسندات التي لم يُحدَّد لها مندوب بعد (السندات القديمة، أو بعد تسجيل مندوب للطرف لاحقاً)."""
    who = RepResolver(conn)
    sql = ("SELECT trans_id, party_code, trans_date, alloc_type, alloc_id FROM cash_transactions "
           "WHERE category=? AND party_code IS NOT NULL AND COALESCE(rep_name,'')=''")
    args = [CAT_PARTNER]
    if party_code:
        sql += " AND party_code=?"
        args.append(party_code)
    n = 0
    for tid, code, d, at, ai in conn.execute(sql, args).fetchall():
        rep = voucher_rep(conn, code, d, at, ai, who)
        if rep:
            conn.execute("UPDATE cash_transactions SET rep_name=? WHERE trans_id=?", (rep, tid))
            n += 1
    return n


def rep_rep_performance(p):
    """أداء المناديب: صافي المبيعات (بعد المرتجعات والإشعارات الدائنة)، صافي المشتريات، التحصيلات والمدفوعات."""
    a, b = p["d_from"], p["d_to"]
    agg = defaultdict(lambda: dict(sales=0.0, purch=0.0, coll=0.0, pay=0.0, n_s=0, n_p=0, cust=set(), supp=set()))
    with db() as conn:
        who = RepResolver(conn)
        for code, d, sub, rep in conn.execute("SELECT customer_code, invoice_date, subtotal, sales_rep FROM sales_invoices "
                                              "WHERE invoice_date BETWEEN ? AND ?", (a, b)):
            g = agg[who(code, d, rep)]
            g["sales"] += sub or 0
            g["n_s"] += 1
            g["cust"].add(code)
        for code, d, sub, rt in conn.execute("SELECT party_code, return_date, subtotal, return_type FROM returns_invoices "
                                             "WHERE return_date BETWEEN ? AND ?", (a, b)):
            agg[who(code, d)]["sales" if rt == SALES_RETURN else "purch"] -= sub or 0
        for code, d, sub, k in conn.execute("SELECT party_code, discount_date, subtotal, kind FROM post_discounts "
                                            "WHERE discount_date BETWEEN ? AND ?", (a, b)):
            agg[who(code, d)]["sales" if k == DISC_SALES else "purch"] -= sub or 0
        for code, d, sub, rep in conn.execute("SELECT supplier_code, purchase_date, subtotal, purchase_rep FROM purchase_invoices "
                                              "WHERE purchase_date BETWEEN ? AND ?", (a, b)):
            g = agg[who(code, d, rep)]
            g["purch"] += sub or 0
            g["n_p"] += 1
            g["supp"].add(code)
        for code, d, amt, t, at, ai, rep in conn.execute(
                "SELECT party_code, trans_date, amount, trans_type, alloc_type, alloc_id, COALESCE(rep_name,'') "
                "FROM cash_transactions WHERE category=? AND party_code IS NOT NULL AND trans_date BETWEEN ? AND ?",
                (CAT_PARTNER, a, b)):
            rep = rep or voucher_rep(conn, code, d, at, ai, who) or NO_REP
            agg[rep]["coll" if t == "قبض" else "pay"] += amt or 0
    rows = []
    for rep, g in agg.items():
        if not any(abs(g[k]) > 0.005 for k in ("sales", "purch", "coll", "pay")):
            continue
        rows.append((rep, len(g["cust"]), g["n_s"], g["sales"], g["coll"], _pct(g["coll"], g["sales"] * (1 + VAT_RATE)),
                     len(g["supp"]), g["n_p"], g["purch"], g["pay"]))
    rows.sort(key=lambda r: r[3], reverse=True)
    cols = [("المندوب", "t"), ("عدد العملاء", "i"), ("فواتير البيع", "i"), ("صافي المبيعات", "m"), ("التحصيلات", "m"),
            ("نسبة التحصيل", "p"), ("عدد الموردين", "i"), ("فواتير الشراء", "i"), ("صافي المشتريات", "m"),
            ("المدفوعات للموردين", "m")]
    total = ("الإجمالي", None, sum(r[2] for r in rows), sum(r[3] for r in rows), sum(r[4] for r in rows), None, None,
             sum(r[7] for r in rows), sum(r[8] for r in rows), sum(r[9] for r in rows))
    notes = ["الإسناد حسب فترة تواجد المندوب الفعلية مع العميل/المورد بتاريخ كل حركة (سجل تغيير المناديب)، ثم مندوب المستند.",
             "التحصيلات والمدفوعات: مندوب الفاتورة المسددة (نفس من نُسبت له المبيعة)، وإلا مندوب الطرف بتاريخ السند — "
             "يُورَّث تلقائياً عند حفظ السند.",
             "صافي المبيعات/المشتريات بعد المرتجعات والخصومات اللاحقة. نسبة التحصيل = التحصيلات ÷ صافي المبيعات شاملة الضريبة."]
    return Report("أداء المناديب (مبيعات، مشتريات، تحصيلات)", cols, rows, total, notes, f"من {a} إلى {b}", chart=(0, 3))


# ---------------------------------------------------------- المرفقات (فواتير الموردين، مستندات القيود، سندات القبض والصرف)
# الملف نفسه في مجلد المرفقات ATTACH_DIR/<نوع المستند>/<رقمه>/<معرّف>_<الاسم> (خارج git)، والقاعدة تحفظ بياناته ومساره.
# المرفقات القديمة المحفوظة داخل القاعدة (data) تُقرأ كما هي. الحذف لا يتلف الملفات: تُنقل إلى ATTACH_DIR/_deleted/.
ATTACH_MAX_MB = 10
ATTACH_EXTS = (".pdf", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".heic", ".xlsx", ".xls",
               ".csv", ".doc", ".docx", ".txt")
ATTACH_DOCS = {"purchase": ("purchase_invoices", "purchase_id", "فاتورة مشتريات"),
               "receipt": ("purchase_receipts", "receipt_id", "إذن استلام (فاتورة المورد)"),
               "journal": ("journal_entries", "entry_id", "قيد يومية"),
               "voucher": ("cash_transactions", "trans_id", "سند قبض / صرف")}
# صلاحيات الرفع لكل نوع (يكفي واحدة منها)
ATTACH_PERMS = {"receipt": ("purchase.receipt_create",), "voucher": ("cash.receipt", "cash.payment", "cash.register")}


def _attach_abs(rel):
    return os.path.join(ATTACH_DIR, *str(rel).split("/"))


def _safe_file_name(name):
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip(" .")
    return name[:120] or "file"


def _archive_attachment_files(rels, tag):
    """نقل ملفات مرفقات محذوفة إلى ATTACH_DIR/_deleted/<tag>/ بدل حذفها (لا تُتلف أي بيانات)."""
    for rel in rels:
        if not rel:
            continue
        src = _attach_abs(rel)
        if not os.path.exists(src):
            continue
        dst = os.path.join(ATTACH_DIR, "_deleted", tag, *str(rel).split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if os.path.exists(dst):
            dst = f"{dst}.{datetime.now():%H%M%S%f}"
        os.replace(src, dst)


def _find_archived(rel):
    """ملف نُقل للأرشيف ثم أُلغيت عملية الحذف (تراجع المعاملة): يُعاد لمكانه."""
    root = os.path.join(ATTACH_DIR, "_deleted")
    if not os.path.isdir(root):
        return None
    for tag in sorted(os.listdir(root), reverse=True):
        cand = os.path.join(root, tag, *str(rel).split("/"))
        if os.path.exists(cand):
            dst = _attach_abs(rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            os.replace(cand, dst)
            return dst
    return None


def add_attachment(doc_type, doc_id, file_name, data, user="", notes=""):
    import mimetypes
    if doc_type not in ATTACH_DOCS:
        raise UserError("نوع المستند غير مدعوم للمرفقات.")
    file_name = os.path.basename(file_name or "").strip()
    if not file_name or not data:
        raise UserError("الملف فارغ.")
    if not file_name.lower().endswith(ATTACH_EXTS):
        raise UserError("نوع الملف غير مدعوم. المسموح: PDF والصور (JPG/PNG...) وملفات Excel/Word/CSV/TXT.")
    if len(data) > ATTACH_MAX_MB * 1024 * 1024:
        raise UserError(f"حجم الملف أكبر من {ATTACH_MAX_MB} ميجابايت.")
    table, idcol, label = ATTACH_DOCS[doc_type]
    keys = ATTACH_PERMS.get(doc_type)
    if keys:
        u = _CURRENT_PERMS.get()
        if u is not None and not any(user_can(u, k) for k in keys):
            require_perm(keys[0], f"إرفاق مستند على {label}")
    import uuid
    rel = f"{doc_type}/{int(doc_id)}/{uuid.uuid4().hex[:12]}_{_safe_file_name(file_name)}"
    with db() as conn:
        if conn.execute(f"SELECT 1 FROM {table} WHERE {idcol}=?", (doc_id,)).fetchone() is None:
            raise UserError(f"{label} غير موجود.")
        path = _attach_abs(rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(data)
        try:
            return conn.execute("INSERT INTO attachments (doc_type, doc_id, file_name, mime, size, data, notes, uploaded_by, "
                                "file_path) VALUES (?,?,?,?,?,?,?,?,?)",
                                (doc_type, int(doc_id), file_name,
                                 mimetypes.guess_type(file_name)[0] or "application/octet-stream", len(data),
                                 sqlite3.Binary(b""), notes, user, rel)).lastrowid
        except Exception:
            os.remove(path)
            raise


def list_attachments(doc_type, doc_id):
    """مرفقات المستند. فاتورة المشتريات تعرض أيضاً مرفقات أذون الاستلام التي صدرت منها (فاتورة المورد المرفوعة عند الاستلام)."""
    with db() as conn:
        rows = conn.execute("SELECT attach_id, file_name, size, mime, COALESCE(uploaded_by,''), uploaded_at FROM attachments "
                            "WHERE doc_type=? AND doc_id=? ORDER BY attach_id", (doc_type, int(doc_id))).fetchall()
        if doc_type == "purchase":
            rows += conn.execute(
                "SELECT a.attach_id, '[إذن ' || a.doc_id || '] ' || a.file_name, a.size, a.mime, COALESCE(a.uploaded_by,''), "
                "a.uploaded_at FROM attachments a JOIN purchase_receipts r ON r.receipt_id=a.doc_id "
                "WHERE a.doc_type='receipt' AND r.invoice_id=? ORDER BY a.attach_id", (int(doc_id),)).fetchall()
        return rows


def get_attachment(attach_id):
    with db() as conn:
        r = conn.execute("SELECT file_name, mime, data, file_path FROM attachments WHERE attach_id=?",
                         (attach_id,)).fetchone()
    if r is None:
        raise UserError("المرفق غير موجود.")
    if r[3]:
        path = _attach_abs(r[3])
        if not os.path.exists(path):
            path = _find_archived(r[3])
        if not path:
            raise UserError(f"ملف المرفق غير موجود في مجلد المرفقات:\n{_attach_abs(r[3])}")
        with open(path, "rb") as fh:
            return r[0], r[1], fh.read()
    return r[0], r[1], bytes(r[2] or b"")


def delete_attachment(attach_id):
    with db() as conn:
        r = conn.execute("SELECT file_path FROM attachments WHERE attach_id=?", (attach_id,)).fetchone()
        conn.execute("DELETE FROM attachments WHERE attach_id=?", (attach_id,))
    if r:
        _archive_attachment_files([r[0]], f"{datetime.now():%Y%m%d_%H%M%S}")


def attachment_counts(doc_type):
    with db() as conn:
        out = dict(conn.execute("SELECT doc_id, COUNT(*) FROM attachments WHERE doc_type=? GROUP BY doc_id", (doc_type,)))
        if doc_type == "purchase":   # + مرفقات أذون الاستلام المفوترة في الفاتورة
            for inv, n in conn.execute("SELECT r.invoice_id, COUNT(*) FROM attachments a JOIN purchase_receipts r "
                                       "ON r.receipt_id=a.doc_id WHERE a.doc_type='receipt' AND r.invoice_id IS NOT NULL "
                                       "GROUP BY r.invoice_id"):
                out[inv] = out.get(inv, 0) + n
        return out


def _delete_doc_attachments(conn, doc_type, doc_id):
    rels = [p for (p,) in conn.execute("SELECT file_path FROM attachments WHERE doc_type=? AND doc_id=?", (doc_type, doc_id))]
    conn.execute("DELETE FROM attachments WHERE doc_type=? AND doc_id=?", (doc_type, doc_id))
    _archive_attachment_files(rels, f"{datetime.now():%Y%m%d_%H%M%S}")


# ============================================================
# العملاء والموردون: حفظ موحد، الكيان المزدوج، وكشف الحساب الموحد بالتصفية
# ============================================================
ROLE_LABEL = {"customer": "عميل", "supplier": "مورد"}


def save_partner_record(role, code, name, phone="", email="", address="", cr="", vat="", rep="", rep_date=None,
                        opening=None, user="", credit=None):
    """إضافة/تعديل عميل أو مورد. الكود يُولَّد تلقائياً (C#### / S####) ولا يُكتب يدوياً.
    الاسم لا يتكرر داخل نفس الشجرة؛ ونفس الاسم/الرقم الضريبي في الشجرة الأخرى = كيان مزدوج يُربط تلقائياً بالكشف الموحد.
    opening = (مدين، دائن، التاريخ) أو None لعدم التغيير.
    credit = (الحد الائتماني، مدة السداد بالأيام أو ""، السياسة warn/block أو "" للافتراضية) أو None لعدم التغيير."""
    if role not in ROLE_LABEL:
        raise UserError("نوع الطرف غير صحيح.")
    name = clean_name(name or "")
    if not name:
        raise UserError("أدخل اسم " + ("العميل" if role == "customer" else "المورد") + ".")
    vat = normalize_digits(vat or "").replace(" ", "").strip()
    flag = "is_customer" if role == "customer" else "is_supplier"
    with db() as conn:
        for c_, n_ in conn.execute(f"SELECT partner_code, name FROM business_partners WHERE {flag}=1 AND partner_code != ?",
                                   (code or "",)):
            if norm_name(n_) == norm_name(name):
                raise UserError(f"يوجد {ROLE_LABEL[role]} بنفس الاسم بالفعل (الكود {c_}).")
        fields = (name, (phone or "").strip(), (email or "").strip(), (address or "").strip(), (cr or "").strip(), vat)
        if code:
            if conn.execute("SELECT 1 FROM business_partners WHERE partner_code=?", (code,)).fetchone() is None:
                raise UserError("الطرف غير موجود.")
            conn.execute(f"UPDATE business_partners SET name=?, phone=?, email=?, address=?, cr_number=?, vat_number=?, {flag}=1 "
                         f"WHERE partner_code=?", fields + (code,))
        else:
            code = next_serial(conn, role)
            conn.execute("INSERT INTO business_partners (partner_code, name, phone, email, address, cr_number, vat_number, "
                         "is_customer, is_supplier) VALUES (?,?,?,?,?,?,?,?,?)",
                         (code,) + fields + ((1, 0) if role == "customer" else (0, 1)))
        set_partner_rep(conn, code, rep, rep_date, user)
        if opening is not None:
            d, c, od = opening
            if (d or c) and not parse_date(od or ""):
                raise UserError("أدخل تاريخ الرصيد الافتتاحي (YYYY-MM-DD).")
            set_partner_opening(conn, code, d or 0.0, c or 0.0, parse_date(od or "") or None)
        if credit is not None:
            from .credit import set_partner_credit
            set_partner_credit(conn, code, *credit)
    return code


def list_partners(role, q=""):
    """صفوف صفحة العملاء أو الموردين: (الكود، الاسم، الهاتف، الرقم الضريبي، المندوب، الافتتاحي، الأكواد المرتبطة، القديم)."""
    flag = "is_customer" if role == "customer" else "is_supplier"
    with db() as conn:
        rows = conn.execute(f"SELECT partner_code, name, COALESCE(phone,''), COALESCE(vat_number,''), COALESCE(rep_name,''), "
                            f"COALESCE(opening_debit,0) - COALESCE(opening_credit,0), COALESCE(legacy_code,''), "
                            f"COALESCE(address,''), COALESCE(cr_number,''), COALESCE(email,'') "
                            f"FROM business_partners WHERE {flag}=1 AND {ACTIVE_PARTNER_SQL} ORDER BY name").fetchall()
        links = partner_link_map(conn)
    if q:
        nq = search_norm(q)
        rows = [r for r in rows if nq in search_norm(" ".join(str(x) for x in r))]
    return [r + (" ، ".join(c for c in links.get(r[0], []) if c != r[0]),) for r in rows]


def partner_link_map(conn):
    """{الكود: [كل أكواد نفس الكيان]} بمطابقة الاسم الموحد أو الرقم الضريبي (15 رقماً) بين الشجرتين."""
    rows = conn.execute("SELECT partner_code, name, COALESCE(vat_number,'') FROM business_partners "
                        f"WHERE {ACTIVE_PARTNER_SQL}").fetchall()
    groups = {}
    parent = {c: c for c, _, _ in rows}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for c, n, v in rows:
        for key in (("n", norm_name(n)), ("v", normalize_digits(v).replace(" ", "")) if len(normalize_digits(v).strip()) >= 10 else None):
            if not key or not key[1]:
                continue
            if key in groups:
                parent[find(c)] = find(groups[key])
            else:
                groups[key] = c
    out = defaultdict(list)
    for c, _, _ in rows:
        out[find(c)].append(c)
    return {c: sorted(out[find(c)]) for c, _, _ in rows if len(out[find(c)]) > 1}


def linked_partner_codes(code):
    with db() as conn:
        return partner_link_map(conn).get(code, [code])


def statement_parties(code):
    """[(الكود، الاسم، الدور)] لكل الأكواد المرتبطة بالكيان (للعرض في رأس الكشف الموحد)."""
    codes = linked_partner_codes(code)
    with db() as conn:
        out = []
        for c in codes:
            r = conn.execute("SELECT name, is_customer, is_supplier FROM business_partners WHERE partner_code=?", (c,)).fetchone()
            if r:
                out.append((c, r[0], " و".join(lbl for lbl, f in (("عميل", r[1]), ("مورد", r[2])) if f) or "-"))
    return out


STATEMENT_SCOPES = {"both": "العملاء والموردون معاً (صافي المقاصة)", "customer": "حركات العميل فقط",
                    "supplier": "حركات المورد فقط"}
# مسميات نفس النطاق في العرض الإجمالي (بدون اختيار طرف)
SUMMARY_SCOPES = {"both": "العملاء والموردين معاً", "customer": "العملاء فقط", "supplier": "الموردين فقط"}


# ============================================================
# شجرة الحسابات: حسابات فرعية هرمية، كشف حساب تحليلي، وتجميع تلقائي تحت الحسابات الرئيسية
# ============================================================
def _account_maps(conn):
    rows = conn.execute("SELECT account_code, account_name, account_type, parent_code, COALESCE(is_group,0) "
                        "FROM chart_of_accounts").fetchall()
    info = {c: dict(name=n, type=t, parent=p, group=bool(g)) for c, n, t, p, g in rows}
    children = defaultdict(list)
    for c, d in info.items():
        if d["parent"] and d["parent"] in info and d["parent"] != c:
            children[d["parent"]].append(c)
    return info, children


def account_descendants(conn, code):
    _, children = _account_maps(conn)
    out, stack, seen = [], [code], set()
    while stack:
        c = stack.pop()
        if c in seen:
            continue
        seen.add(c)
        out.append(c)
        stack.extend(children.get(c, []))
    return out


def _account_has_postings(conn, code, name):
    if conn.execute("SELECT 1 FROM journal_lines WHERE account_code=? LIMIT 1", (code,)).fetchone():
        return True
    if conn.execute("SELECT 1 FROM cash_transactions WHERE party_code=? AND category != ? LIMIT 1",
                    (code, CAT_PARTNER)).fetchone():
        return True
    if code.startswith(CASH_PREFIX) and conn.execute("SELECT 1 FROM cash_transactions WHERE account_bank=? LIMIT 1",
                                                     (name,)).fetchone():
        return True
    return False


def next_child_account_code(conn, parent):
    existing = {r[0] for r in conn.execute("SELECT account_code FROM chart_of_accounts")}
    kids = [r[0] for r in conn.execute("SELECT account_code FROM chart_of_accounts WHERE parent_code=?", (parent,))
            if r[0].isdigit()]
    nxt = str(max(int(c) for c in kids) + 1) if kids else f"{parent}01"
    while nxt in existing:
        nxt = str(int(nxt) + 1)
    return nxt


def create_sub_account(parent_code, name, acc_type=None, is_group=False):
    """إنشاء حساب فرعي تحت أي حساب بكود تلقائي (مثل تفريع بنك، خزينة فرعية، سلفة/عهدة باسم موظف).
    الحساب الأب يتحول لحساب تجميعي، بشرط ألا تكون عليه حركات مباشرة."""
    name = clean_name(name or "")
    if not name:
        raise UserError("أدخل اسم الحساب الفرعي.")
    with db() as conn:
        p = conn.execute("SELECT account_name, account_type, COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?",
                         (parent_code,)).fetchone()
        if p is None:
            raise UserError("الحساب الرئيسي غير موجود.")
        if conn.execute("SELECT 1 FROM chart_of_accounts WHERE parent_code=? AND account_name=?", (parent_code, name)).fetchone():
            raise UserError("يوجد حساب فرعي بنفس الاسم تحت هذا الحساب.")
        if not p[2]:
            if _account_has_postings(conn, parent_code, p[0]):
                raise UserError(f"الحساب «{p[0]}» عليه حركات مباشرة (قيود أو سندات)، فلا يمكن تحويله لحساب رئيسي.\n"
                                "أنشئ الحساب الفرعي تحت الحساب الأعلى منه، أو انقل الحركات بقيد تسوية أولاً.")
            conn.execute("UPDATE chart_of_accounts SET is_group=1 WHERE account_code=?", (parent_code,))
        code = next_child_account_code(conn, parent_code)
        acc_type = acc_type if acc_type in ACCOUNT_TYPES else p[1]
        conn.execute("INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                     "VALUES (?,?,?,?,?)", (code, name, acc_type, parent_code, 1 if is_group else 0))
    return code


def account_tree_rows():
    """[(الكود، الاسم، النوع، المستوى، تجميعي؟)] بترتيب الشجرة (الأب ثم أبناؤه)."""
    with db() as conn:
        info, children = _account_maps(conn)
    roots = sorted(c for c, d in info.items() if not d["parent"] or d["parent"] not in info)
    out = []

    def walk(c, lvl):
        d = info[c]
        out.append((c, d["name"], d["type"], lvl, d["group"] or bool(children.get(c))))
        for k in sorted(children.get(c, [])):
            walk(k, lvl + 1)

    for r in roots:
        walk(r, 0)
    return out


def _voucher_account_moves(conn, info):
    """حركات سندات الخزينة كحركات حسابات: حساب الخزينة/البنك (قبض مدين، صرف دائن)
    + الحساب المقابل غير الشريك (مصروف/سلفة/عهدة: صرف مدين، قبض دائن). [(الكود، التاريخ، مرجع، بيان، مدين، دائن)]"""
    by_name = {}
    for c, d in info.items():
        if c.startswith(CASH_PREFIX) and not d["group"]:
            by_name.setdefault(d["name"], c)
    for n, c in DEFAULT_CASH_CODES.items():
        by_name.setdefault(n, c)
    out = []
    for tid, d, t, bank, cat, pcode, pname, amt, desc in conn.execute(
            "SELECT trans_id, trans_date, trans_type, account_bank, category, party_code, party_name, amount, "
            "COALESCE(description,'') FROM cash_transactions"):
        memo = f"سند {t} - {pname}" + (f" - {desc}" if desc else "")
        cash_code = by_name.get(bank)
        if cash_code:
            out.append((cash_code, d, f"سند-{tid}", memo, amt if t == "قبض" else 0.0, amt if t == "صرف" else 0.0))
        if cat != CAT_PARTNER and pcode and pcode in info:
            out.append((pcode, d, f"سند-{tid}", f"سند {t} من {bank}" + (f" - {desc}" if desc else ""),
                        amt if t == "صرف" else 0.0, amt if t == "قبض" else 0.0))
    return out


def account_statement(code, d_from, d_to):
    """كشف حساب تحليلي لحساب (وكل فروعه): رصيد أول المدة، الحركات (قيود + سندات)، المسحوب/المسدد والمتبقي."""
    with db() as conn:
        info, _ = _account_maps(conn)
        if code not in info:
            raise UserError("الحساب غير موجود.")
        codes = set(account_descendants(conn, code))
        moves = []
        for eid, d, desc, acc, dr, cr, memo in conn.execute(
                f"SELECT e.entry_id, e.entry_date, COALESCE(e.description,''), l.account_code, l.debit, l.credit, "
                f"COALESCE(l.memo,'') FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id "
                f"WHERE l.account_code IN ({','.join('?' * len(codes))})", list(codes)):
            moves.append((d, f"قيد-{eid}", memo or desc, acc, dr or 0.0, cr or 0.0))
        for acc, d, ref, memo, dr, cr in _voucher_account_moves(conn, info):
            if acc in codes:
                moves.append((d, ref, memo, acc, dr, cr))
    moves.sort(key=lambda m: (m[0], m[1]))
    opening = sum(m[4] - m[5] for m in moves if m[0] < d_from)
    rows, bal, tdr, tcr = [], opening, 0.0, 0.0
    for d, ref, memo, acc, dr, cr in moves:
        if d < d_from or d > d_to:
            continue
        bal += dr - cr
        tdr += dr
        tcr += cr
        rows.append((d, ref, f"{acc} - {info.get(acc, {}).get('name', acc)}", memo, dr, cr, bal))
    return dict(code=code, name=info[code]["name"], type=info[code]["type"], group=len(codes) > 1, opening=opening,
                debit=tdr, credit=tcr, closing=bal, rows=rows, accounts=len(codes))


def account_balances(d_from, d_to, include_vouchers=True):
    """{الكود: [أول المدة، مدين، دائن]} لكل حساب فرعي من القيود (+ السندات اختيارياً)، قبل التجميع."""
    with db() as conn:
        info, _ = _account_maps(conn)
        data = defaultdict(lambda: [0.0, 0.0, 0.0])
        for acc, before, dr, cr in conn.execute(
                "SELECT l.account_code, SUM(CASE WHEN e.entry_date < ? THEN l.debit - l.credit ELSE 0 END), "
                "SUM(CASE WHEN e.entry_date >= ? THEN l.debit ELSE 0 END), SUM(CASE WHEN e.entry_date >= ? THEN l.credit ELSE 0 END) "
                "FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id WHERE e.entry_date <= ? "
                "GROUP BY l.account_code", (d_from, d_from, d_from, d_to)):
            r = data[acc]
            r[0] += before or 0.0
            r[1] += dr or 0.0
            r[2] += cr or 0.0
        if include_vouchers:
            for acc, d, _ref, _m, dr, cr in _voucher_account_moves(conn, info):
                if d > d_to:
                    continue
                r = data[acc]
                if d < d_from:
                    r[0] += dr - cr
                else:
                    r[1] += dr
                    r[2] += cr
    return info, data


def _rollup(info, data):
    """تجميع أرصدة الحسابات الفرعية تحت كل الحسابات الرئيسية (الأب، والجد، ...)."""
    total = defaultdict(lambda: [0.0, 0.0, 0.0])
    for acc, vals in data.items():
        seen = set()
        c = acc
        while c and c not in seen:
            seen.add(c)
            t = total[c]
            for i in range(3):
                t[i] += vals[i]
            c = info.get(c, {}).get("parent")
    return total


def _tree_order(info, keys):
    children = defaultdict(list)
    for c, d in info.items():
        if d["parent"] in info and d["parent"] != c:
            children[d["parent"]].append(c)
    roots = sorted(c for c in keys if not info.get(c, {}).get("parent") or info[c]["parent"] not in info)
    out = []

    def walk(c, lvl):
        if c in keys:
            out.append((c, lvl))
        for k in sorted(children.get(c, [])):
            walk(k, lvl + 1)

    for r in roots:
        walk(r, 0)
    listed = {c for c, _ in out}
    out += [(c, 0) for c in sorted(keys) if c not in listed]  # أكواد غير موجودة بالدليل
    return out


def _tree_report_rows(info, data, two_sided):
    total = _rollup(info, data)
    keys = {c for c, v in total.items() if any(abs(x) > 0.005 for x in v)}
    rows = []
    for c, lvl in _tree_order(info, keys):
        o, dr, cr = total[c]
        close = o + dr - cr
        is_group = c not in data or info.get(c, {}).get("group") or any(
            k for k in keys if info.get(k, {}).get("parent") == c)
        name = ("    " * lvl) + ("▸ " if is_group else "") + info.get(c, {}).get("name", c)
        typ = ("تجميعي" if is_group else "فرعي") + f" • {info.get(c, {}).get('type', '')}"
        if two_sided:
            pos = lambda x: x if x > 0.005 else 0.0
            rows.append((c, name, typ, pos(o), pos(-o), dr, cr, pos(close), pos(-close)))
        else:
            rows.append((c, name, typ, o, dr, cr, close))
    return rows, keys


def rep_trial_balance(p):
    """ميزان مراجعة القيود اليومية بشجرة الحسابات: الحسابات الفرعية تُجمَّع تلقائياً تحت الرئيسية.
    الإجمالي من الحسابات الفرعية فقط (منعاً للتكرار)."""
    if p.get("source", "full") == "full":   # الافتراضي: الدفتر العام الكامل (المستندات مرحّلة آلياً + القيود)
        from .ledger import rep_gl_trial_balance
        return rep_gl_trial_balance(p)
    a, b = p["d_from"], p["d_to"]
    info, data = account_balances(a, b, include_vouchers=False)
    rows, _ = _tree_report_rows(info, data, two_sided=True)
    leaves = []
    for acc, (o, dr, cr) in data.items():
        close = o + dr - cr
        pos = lambda x: x if x > 0.005 else 0.0
        leaves.append((pos(o), pos(-o), dr, cr, pos(close), pos(-close)))
    cols = [("رقم الحساب", "t"), ("اسم الحساب", "t"), ("المستوى / النوع", "t"), ("أول المدة مدين", "m"), ("أول المدة دائن", "m"),
            ("حركة مدين", "m"), ("حركة دائن", "m"), ("آخر المدة مدين", "m"), ("آخر المدة دائن", "m")]
    sums = [sum(r[i] for r in leaves) for i in range(6)]
    total = ("", "الإجمالي (الحسابات الفرعية)", "", *sums)
    ok = abs(sums[2] - sums[3]) < 0.01 and abs(sums[4] - sums[5]) < 0.01
    notes = ["الحسابات الرئيسية (▸) تعرض مجموع حساباتها الفرعية تلقائياً. القيود اليومية فقط (يدوية + آلية).",
             "الميزان متوازن ✓" if ok else "تنبيه: الميزان غير متوازن!"]
    return Report("ميزان المراجعة (شجرة الحسابات)", cols, rows, total, notes, f"من {a} إلى {b}")


def rep_account_tree(p):
    """أرصدة الحسابات الفرعية تحليلياً (القيود + سندات الخزينة) مجمعة تحت الحسابات الرئيسية."""
    a, b = p["d_from"], p["d_to"]
    info, data = account_balances(a, b, include_vouchers=True)
    rows, _ = _tree_report_rows(info, data, two_sided=False)
    cols = [("الكود", "t"), ("الحساب", "t"), ("المستوى / النوع", "t"), ("رصيد أول المدة", "m"), ("المسحوب / مدين", "m"),
            ("المسدد / دائن", "m"), ("المتبقي (الرصيد)", "m")]
    notes = ["يشمل القيود اليومية وسندات الخزينة: حسابات النقدية والبنوك، والمصروفات، وسلف وعهد العاملين بأسمائهم.",
             "الرصيد الموجب = مدين (مثلاً سلفة على الموظف أو نقدية متاحة)، والسالب = دائن. الحسابات الرئيسية (▸) مجمعة تلقائياً."]
    return Report("أرصدة الحسابات الفرعية (تحليلي مجمع)", cols, rows, None, notes, f"من {a} إلى {b}")


# ============================================================
# المستندات غير المفوترة (مبيعات + مشتريات في جدول واحد)
# ============================================================
def all_unbilled(kind=None, q=""):
    """كل أذون التسليم والاستلام التي لم تصدر لها فاتورة، الأقدم أولاً (منعاً للنسيان)."""
    out = []
    today_d = date.today()
    with db() as conn:
        for k in ([kind] if kind else ["sales", "purchase"]):
            c = FLOW[k]
            for fid, d, oid, pcode, pname, n, sub in conn.execute(
                    f"SELECT f.{c['fid']}, f.{c['fdate']}, f.order_id, f.{c['pcode']}, f.{c['pname']}, COUNT(l.line_id), "
                    f"COALESCE(SUM({c['fval']}),0) FROM {c['ful']} f JOIN {c['flines']} l ON l.{c['fid']}=f.{c['fid']} "
                    f"WHERE f.invoice_id IS NULL GROUP BY f.{c['fid']}"):
                vat = round(sub * VAT_RATE, 2)
                age = (today_d - datetime.strptime(d, "%Y-%m-%d").date()).days if parse_date(d or "") else 0
                out.append(dict(kind=k, id=fid, date=d, order=oid, party_code=pcode, party=pname, lines=n, subtotal=sub,
                                vat=vat, total=round(sub + vat, 2), age=age))
    if q:
        nq = search_norm(q)
        out = [r for r in out if nq in search_norm(f"{r['party_code']} {r['party']} {r['id']} {r['order']} {r['date']}")]
    out.sort(key=lambda r: (r["date"] or "", r["id"]))
    return out


def invoice_unbilled_selection(kind, ids, inv_date, notes="", supplier_ref="", allow_closed=False):
    """إصدار فاتورة من أذون محددة في الجدول الموحد (يجب أن تكون لنفس النوع ونفس الطرف).
    allow_closed: السماح بيوم مغلق (المدير، أو بعد تأكيد كلمة مرور الإغلاق في سطح المكتب)."""
    if not parse_date(inv_date or ""):
        raise UserError("تاريخ الفاتورة غير صحيح (YYYY-MM-DD).")
    if not allow_closed and period_is_closed(inv_date):
        raise UserError(f"يوم {inv_date} مغلق (تم الإغلاق اليومي). لا يمكن إصدار فاتورة بتاريخه إلا بواسطة المدير.")
    rows = {r["id"]: r for r in all_unbilled(kind)}
    picked = [rows[i] for i in ids if i in rows]
    if not picked:
        raise UserError("حدد إذناً واحداً على الأقل من الجدول.")
    parties = {r["party_code"] for r in picked}
    if len(parties) > 1:
        raise UserError("الأذون المحددة لأطراف مختلفة. الفاتورة الواحدة تكون لطرف واحد.")
    return invoice_from_fulfilments(kind, parties.pop(), [r["id"] for r in picked], inv_date, notes, supplier_ref)


# ============================================================
# تفاصيل الصنف (للنوافذ المنبثقة)
# ============================================================
def item_details(code):
    with db() as conn:
        r = conn.execute("SELECT item_code, item_name, unit, quantity, cost_price, COALESCE(selling_price,0), "
                         "COALESCE(barcode,''), COALESCE(legacy_code,''), opening_qty, opening_cost FROM inventory "
                         "WHERE item_code=?", (code,)).fetchone()
        if r is None:
            raise UserError("الصنف غير موجود.")
        sold = conn.execute("SELECT COALESCE(SUM(l.quantity),0), COALESCE(SUM(l.subtotal),0) FROM sales_invoice_lines l "
                            "WHERE l.item_code=?", (code,)).fetchone()
        bought = conn.execute("SELECT COALESCE(SUM(quantity),0), COALESCE(SUM(subtotal),0), MAX(unit_cost) "
                              "FROM purchase_invoice_lines WHERE item_code=?", (code,)).fetchone()
        last_cost = conn.execute("SELECT l.unit_cost FROM purchase_invoice_lines l JOIN purchase_invoices h "
                                 "ON h.purchase_id=l.purchase_id WHERE l.item_code=? ORDER BY h.purchase_date DESC, "
                                 "l.line_id DESC LIMIT 1", (code,)).fetchone()
    rows, final = fetch_item_card(code, "0000-01-01", "9999-12-31")
    keys = ("code", "name", "unit", "qty", "cost", "price", "barcode", "legacy", "open_qty", "open_cost")
    d = dict(zip(keys, r))
    d.update(value=(r[3] or 0) * (r[4] or 0), sold_qty=sold[0], sold_value=sold[1], bought_qty=bought[0],
             bought_value=bought[1], last_cost=last_cost[0] if last_cost else None,
             margin=_pct((d["price"] or 0) - (d["cost"] or 0), d["price"] or 0) if d["price"] else None,
             moves=rows[1:][-25:], balance=final)
    return d


def item_reference_count(conn, code):
    return sum(conn.execute(f"SELECT COUNT(*) FROM {t} WHERE item_code=?", (code,)).fetchone()[0] for t in ITEM_REF_TABLES)


def update_item(old_code, new_code, name, unit, opening_qty, opening_cost, allow_negative=False):
    """تعديل الصنف (المدير): الكود (يُنقل لكل المستندات)، الاسم، الوحدة، والرصيد الافتتاحي حتى لو عليه حركات."""
    new_code = (new_code or "").strip()
    name = clean_name(name or "")
    if not new_code or not name:
        raise UserError("الكود والاسم مطلوبان.")
    if opening_qty is None or opening_qty < 0 or opening_cost is None or opening_cost < 0:
        raise UserError("الرصيد والتكلفة الافتتاحية أرقام صفر أو أكثر.")
    with db() as conn:
        if conn.execute("SELECT 1 FROM inventory WHERE item_code=?", (old_code,)).fetchone() is None:
            raise UserError("الصنف غير موجود.")
        if new_code != old_code:
            if conn.execute("SELECT 1 FROM inventory WHERE item_code=?", (new_code,)).fetchone():
                raise UserError(f"الكود {new_code} مستخدم لصنف آخر.")
            _rename_item(conn, old_code, new_code)
        conn.execute("UPDATE inventory SET item_name=?, unit=?, opening_qty=?, opening_cost=? WHERE item_code=?",
                     (name, unit or "KG", opening_qty, opening_cost, new_code))
        recalc_items(conn, [new_code], allow_negative)


def delete_item(code):
    with db() as conn:
        n = item_reference_count(conn, code)
        if n:
            raise UserError(f"الصنف مستخدم في {n} بند (فواتير/أوامر/أذون/مرتجعات).\n"
                            "احذف هذه المستندات أولاً، أو عدّل الصنف بدلاً من حذفه.")
        conn.execute("DELETE FROM inventory WHERE item_code=?", (code,))


def partner_reference_count(conn, code):
    checks = [("sales_invoices", "customer_code"), ("purchase_invoices", "supplier_code"), ("returns_invoices", "party_code"),
              ("sales_orders", "customer_code"), ("purchase_orders", "supplier_code"),
              ("sales_deliveries", "customer_code"), ("purchase_receipts", "supplier_code"),
              ("post_discounts", "party_code")]
    n = sum(conn.execute(f"SELECT COUNT(*) FROM {t} WHERE {col}=?", (code,)).fetchone()[0] for t, col in checks)
    n += conn.execute("SELECT COUNT(*) FROM cash_transactions WHERE party_code=? AND category=?",
                      (code, CAT_PARTNER)).fetchone()[0]
    return n


def delete_partner(code):
    with db() as conn:
        n = partner_reference_count(conn, code)
        if n:
            raise UserError(f"الشريك عليه {n} مستند (فواتير/أوامر/مرتجعات/سندات).\nاحذف هذه المستندات أولاً.")
        conn.execute("DELETE FROM business_partners WHERE partner_code=?", (code,))


def save_journal(e_date, desc, lines, entry_id=None):
    """حفظ قيد جديد أو تعديل قيد قائم. lines = [(الكود، الاسم، مدين، دائن، ملاحظات)]."""
    if not parse_date(e_date or ""):
        raise UserError("تاريخ القيد غير صحيح.")
    if len(lines) < 2:
        raise UserError("القيد يحتاج طرفين على الأقل (مدين ودائن).")
    for i, l in enumerate(lines, 1):
        dr, cr = l[2] or 0.0, l[3] or 0.0
        if dr < 0 or cr < 0:
            raise UserError(f"السطر {i}: لا تُقبل مبالغ سالبة في القيد (استخدم الجانب المقابل).")
        if dr > 0.005 and cr > 0.005:
            raise UserError(f"السطر {i}: السطر الواحد يكون مديناً أو دائناً فقط، وليس الاثنين معاً.")
        if dr <= 0.005 and cr <= 0.005:
            raise UserError(f"السطر {i}: سطر بدون مبلغ. احذفه أو أدخل قيمة.")
    td = round(sum(l[2] for l in lines), 2)
    tc = round(sum(l[3] for l in lines), 2)
    if td <= 0.005:
        raise UserError("لا يمكن حفظ قيد بقيمة صفرية.")
    if abs(td - tc) > 0.005:
        raise UserError(f"القيد غير متوازن!\nإجمالي المدين: {fmt(td)}\nإجمالي الدائن: {fmt(tc)}")
    before_lines = None
    with db() as conn:
        for acc_code, *_ in lines:
            acc = conn.execute("SELECT COALESCE(is_group,0) FROM chart_of_accounts WHERE account_code=?", (acc_code,)).fetchone()
            if acc is None:
                raise UserError(f"الحساب {acc_code} غير موجود في دليل الحسابات.")
            if acc[0]:
                raise UserError(f"الحساب {acc_code} حساب رئيسي (تجميعي). رحّل على أحد حساباته الفرعية.")
        if entry_id is None:
            entry_id = conn.execute("INSERT INTO journal_entries (entry_date, description) VALUES (?, ?)", (e_date, desc)).lastrowid
        else:
            if conn.execute("SELECT 1 FROM journal_entries WHERE entry_id=?", (entry_id,)).fetchone() is None:
                raise UserError("القيد غير موجود.")
            _block_system_entry(conn, entry_id)
            before_lines = _lines_snapshot(conn, "journal_lines", "entry_id", entry_id)
            conn.execute("UPDATE journal_entries SET entry_date=?, description=? WHERE entry_id=?", (e_date, desc, entry_id))
            conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (entry_id,))
        for acc_code, acc_name, debit, credit, memo in lines:
            conn.execute("INSERT INTO journal_lines (entry_id, account_code, account_name, debit, credit, memo) "
                         "VALUES (?,?,?,?,?,?)", (entry_id, acc_code, acc_name, debit, credit, memo))
        if before_lines is not None:
            _audit_lines(conn, "تعديل", "أسطر قيد يومية", entry_id, "journal_lines", "entry_id", entry_id, before_lines)
    return entry_id, td


def journal_entry_source(conn, entry_id):
    r = conn.execute("SELECT source_type, source_id FROM journal_entries WHERE entry_id=?", (entry_id,)).fetchone()
    return (r[0], r[1]) if r and r[0] else None


def _block_system_entry(conn, entry_id):
    src = journal_entry_source(conn, entry_id)
    if src and src[0] in ("hw_tb_expenses", "hw_cogs_closing"):
        what = ("«تجميع فترة من ميزان HansaWorld» لأرصدة المصاريف" if src[0] == "hw_tb_expenses"
                else "«تسوية تكلفة مبيعات فترة HansaWorld بالمخزون الختامي»")
        raise UserError(f"هذا قيد نظام {what}.\n"
                        "يُعاد بناؤه أو يُلغى من مركز الاستيراد ← استيراد HansaWorld فقط.")
    if src:
        raise UserError("هذا قيد آلي ينشئه النظام من مستند (إشعار خصم لاحق) ويُحدَّث معه تلقائياً.\n"
                        "لتعديله أو حذفه عدّل/احذف الإشعار نفسه من تبويب «الخصم اللاحق».")


def delete_journal_entry(entry_id):
    with db() as conn:
        _block_system_entry(conn, entry_id)
        _audit_lines(conn, "حذف", "أسطر قيد يومية", entry_id, "journal_lines", "entry_id", entry_id,
                     _lines_snapshot(conn, "journal_lines", "entry_id", entry_id))
        _delete_doc_attachments(conn, "journal", entry_id)
        conn.execute("DELETE FROM journal_lines WHERE entry_id=?", (entry_id,))
        conn.execute("DELETE FROM journal_entries WHERE entry_id=?", (entry_id,))
        conn.execute("DELETE FROM app_settings WHERE key=? AND value=?", (CASH_OPENING_SETTING, str(entry_id)))


def delete_voucher(trans_id):
    with db() as conn:
        r = conn.execute("SELECT alloc_type, alloc_id FROM cash_transactions WHERE trans_id=?", (trans_id,)).fetchone()
        if r is None:
            raise UserError("السند غير موجود.")
        linked = conn.execute("SELECT DISTINCT kind, invoice_id FROM advance_allocations WHERE trans_id=?",
                              (trans_id,)).fetchall()
        conn.execute("DELETE FROM advance_allocations WHERE trans_id=?", (trans_id,))
        conn.execute("DELETE FROM cash_transactions WHERE trans_id=?", (trans_id,))
        _delete_doc_attachments(conn, "voucher", trans_id)
        if r[0] in ("sales", "purchase") and r[1]:
            refresh_invoice_paid(conn, r[0], r[1])
        for k_, i_ in linked:   # الفواتير التي سُوّيت من هذه الدفعة تعود مفتوحة بقيمتها
            refresh_invoice_paid(conn, k_, i_)


def invoice_payment_prefill(kind, inv_id):
    """بيانات «سداد هذه الفاتورة» الجاهزة لشاشة الخزينة: نوع السند، الطرف، المتبقي بعد المسدد والخصومات اللاحقة."""
    c = FLOW[kind]
    with db() as conn:
        h = conn.execute(f"SELECT COALESCE(invoice_no, {c['iid']}), {c['pcode']}, {c['pname']}, total_amount, paid_amount "
                         f"FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
        if h is None:
            raise UserError("الفاتورة غير موجودة.")
        remaining = round(h[3] - h[4] - invoice_discount_total(conn, kind, inv_id), 2)
    return dict(kind=kind, inv_id=inv_id, no=str(h[0]), t_type="قبض" if kind == "sales" else "صرف", category=CAT_PARTNER,
                party=(h[1], h[2]), remaining=max(remaining, 0.0), alloc=(kind, inv_id),
                description=f"سداد فاتورة {'مبيعات' if kind == 'sales' else 'مشتريات'} رقم {h[0]}")


def find_invoice_id(kind, number):
    """رقم الفاتورة الظاهر (التسلسلي أو الداخلي) ← المعرّف الداخلي."""
    c = FLOW[kind]
    with db() as conn:
        r = conn.execute(f"SELECT {c['iid']} FROM {c['inv']} WHERE invoice_no=? OR CAST({c['iid']} AS TEXT)=? "
                         f"ORDER BY (invoice_no=?) DESC LIMIT 1", (str(number), str(number), str(number))).fetchone()
    return r[0] if r else None


def check_voucher_header(t_type, t_date, category, party, amount):
    """تحقق رأس السند المشترك بين السند العادي والسند الموزّع على عدة فواتير. يرجع المبلغ مقرباً لهللتين."""
    if t_type not in ("قبض", "صرف"):
        raise UserError("نوع السند غير صحيح.")
    require_perm("cash.receipt" if t_type == "قبض" else "cash.payment", f"سند {t_type}")
    if not parse_date(t_date or ""):
        raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
    if t_date > today():
        raise UserError("لا يمكن تسجيل سند بتاريخ مستقبلي.")
    if not party or not party[1]:
        raise UserError("اختر الشريك / الحساب من القائمة.")
    if amount is None or amount <= 0:
        raise UserError("أدخل مبلغاً صحيحاً أكبر من الصفر.")
    if category in (CAT_EXPENSE, CAT_ASSET) and t_type == "قبض":
        raise UserError("المصروفات والأصول الثابتة تُسجّل بسند صرف فقط.")
    if category == CAT_EMPLOYEE and not party[0]:
        raise UserError("اختر حساب العامل من القائمة (حسابات ذمم/عهد العاملين).")
    return round(amount, 2)


def save_voucher(t_type, t_date, account_bank, category, party, amount, alloc=None, description=""):
    """تسجيل سند قبض/صرف (منقول من شاشة الخزينة إلى المحرك). party = (الكود، الاسم)، alloc = (sales|purchase، رقم الفاتورة).
    التحقق: نوع السند يناسب الفاتورة، الفاتورة تخص الطرف ومفتوحة، والمبلغ لا يتجاوز المتبقي بعد الخصومات اللاحقة."""
    amount = check_voucher_header(t_type, t_date, category, party, amount)
    with db() as conn:
        alloc_type = alloc_id = None
        if alloc and category == CAT_PARTNER:
            alloc_type, alloc_id = alloc
            if (alloc_type == "sales") != (t_type == "قبض"):
                raise UserError("نوع السند لا يناسب الفاتورة المحددة (القبض لفواتير المبيعات والصرف لفواتير المشتريات).")
            table, idcol, pcol = (("sales_invoices", "invoice_id", "customer_code") if alloc_type == "sales"
                                  else ("purchase_invoices", "purchase_id", "supplier_code"))
            row = conn.execute(f"SELECT total_amount, paid_amount, {pcol}, status FROM {table} WHERE {idcol}=?",
                               (alloc_id,)).fetchone()
            if not row:
                raise UserError("الفاتورة المحددة غير موجودة.")
            if row[2] != party[0]:
                raise UserError("الفاتورة المحددة لا تخص الشريك المختار. أعد اختيار الشريك ثم الفاتورة.")
            if row[3] != "مفتوحة":
                raise UserError("الفاتورة المحددة مسددة بالكامل بالفعل.")
            remaining = row[0] - row[1] - invoice_discount_total(conn, alloc_type, alloc_id)
            if amount > remaining + 0.005:
                raise UserError(f"المبلغ ({fmt(amount)}) أكبر من المتبقي على الفاتورة بعد الخصومات اللاحقة "
                                f"({fmt(remaining)}).")
        rep = voucher_rep(conn, party[0], t_date, alloc_type, alloc_id) if category == CAT_PARTNER else None
        vid = conn.execute(
            "INSERT INTO cash_transactions (trans_type, trans_date, account_bank, category, party_code, party_name, "
            "amount, alloc_type, alloc_id, description, rep_name) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (t_type, t_date, account_bank, category, party[0], party[1], amount, alloc_type, alloc_id,
             (description or "").strip(), rep)).lastrowid
        if alloc_type:
            refresh_invoice_paid(conn, alloc_type, alloc_id)
    return vid


def update_voucher(trans_id, t_date, account_bank, amount, description):
    if not parse_date(t_date or ""):
        raise UserError("تاريخ السند غير صحيح.")
    if amount is None or amount <= 0:
        raise UserError("المبلغ يجب أن يكون أكبر من الصفر.")
    with db() as conn:
        r = conn.execute("SELECT alloc_type, alloc_id FROM cash_transactions WHERE trans_id=?", (trans_id,)).fetchone()
        if r is None:
            raise UserError("السند غير موجود.")
        if r[0] in ("sales", "purchase") and r[1]:
            c = FLOW[r[0]]
            inv = conn.execute(f"SELECT total_amount, COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} WHERE {c['iid']}=?",
                               (r[1],)).fetchone()
            if inv:
                others = conn.execute("SELECT COALESCE(SUM(amount),0) FROM cash_transactions WHERE alloc_type=? AND alloc_id=? "
                                      "AND trans_id!=?", (r[0], r[1], trans_id)).fetchone()[0] \
                    + advance_allocated_to(conn, r[0], r[1])
                remaining = inv[0] - others - invoice_discount_total(conn, r[0], r[1])
                if amount > remaining + 0.005:
                    raise UserError(f"المبلغ ({fmt(amount)}) أكبر من المتبقي على الفاتورة {inv[1]} ({fmt(max(remaining, 0))}) "
                                    "بعد السندات الأخرى والخصومات اللاحقة.\nسجّل الزيادة كدفعة مقدمة بسند مستقل غير مربوط بفاتورة.")
        used = advance_allocated_from(conn, trans_id)
        if not r[0] and amount < used - 0.005:
            raise UserError(f"سُوّي من هذه الدفعة {fmt(used)} على فواتير؛ لا يمكن تخفيض مبلغها لأقل من ذلك.\n"
                            "ألغِ بعض التسويات أولاً.")
        conn.execute("UPDATE cash_transactions SET trans_date=?, account_bank=?, amount=?, description=? WHERE trans_id=?",
                     (t_date, account_bank, round(amount, 2), description, trans_id))
        if r[0] in ("sales", "purchase") and r[1]:
            refresh_invoice_paid(conn, r[0], r[1])


# ============================================================
# فحص سلامة البيانات (Data Integrity Audit) — قراءة فقط، لا يغيّر أي بيان
# ============================================================
def integrity_check():
    """فحص محاسبي ومخزني شامل لقاعدة البيانات. يرجع [(المستوى، المجال، الوصف، العدد، أمثلة)].
    المستوى: «خطأ» يحتاج إصلاحاً، «تنبيه» يحتاج مراجعة، «سليم» لا مشكلة."""
    out = []

    def add(level, area, desc, rows):
        rows = list(rows)
        out.append((level if rows else "سليم", area, desc, len(rows), ", ".join(str(r) for r in rows[:8])))

    conn = sqlite3.connect(DB_FILE, timeout=DB_TIMEOUT)
    _register_functions(conn)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        q = lambda sql, p=(): conn.execute(sql, p).fetchall()  # noqa: E731
        # 1) سلامة ملف قاعدة البيانات
        res = q("PRAGMA integrity_check")
        add("خطأ", "قاعدة البيانات", "سلامة ملف SQLite (PRAGMA integrity_check)",
            [r[0] for r in res if r[0] != "ok"])
        add("خطأ", "قاعدة البيانات", "مراجع مفقودة (Foreign keys)", [f"{r[0]}#{r[1]}" for r in q("PRAGMA foreign_key_check")])
        # 2) الفواتير: الرأس = مجموع البنود
        for kind, c in FLOW.items():
            label = "المبيعات" if kind == "sales" else "المشتريات"
            add("خطأ", f"فواتير {label}", "إجمالي الفاتورة لا يساوي مجموع بنودها (قبل الضريبة / الضريبة / الإجمالي)",
                [r[0] for r in q(f"SELECT COALESCE(h.invoice_no, h.{c['iid']}) FROM {c['inv']} h LEFT JOIN "
                                 f"(SELECT {c['iid']} k, SUM(subtotal) s, SUM(vat_amount) v, SUM(total_amount) t FROM {c['ilines']} "
                                 f"GROUP BY {c['iid']}) l ON l.k=h.{c['iid']} "
                                 f"WHERE COALESCE(h.source_type,'') != 'HansaWorld' AND (ABS(h.subtotal-COALESCE(l.s,0))>0.05 "
                                 f"OR ABS(h.vat_amount-COALESCE(l.v,0))>0.05 OR ABS(h.total_amount-COALESCE(l.t,0))>0.05)")])
            add("خطأ", f"فواتير {label}", "ضريبة البند لا تساوي 15% من قيمته",
                [r[0] for r in q(f"SELECT DISTINCT {c['iid']} FROM {c['ilines']} "
                                 f"WHERE ABS(vat_amount - ROUND(subtotal*{VAT_RATE},2)) > 0.011")])
            add("خطأ", f"فواتير {label}", "المسدد المسجل لا يساوي مجموع السندات المربوطة",
                [r[0] for r in q(f"SELECT COALESCE(h.invoice_no, h.{c['iid']}) FROM {c['inv']} h WHERE ABS(h.paid_amount - "
                                 f"(SELECT COALESCE(SUM(amount),0) FROM cash_transactions WHERE alloc_type=? AND alloc_id=h.{c['iid']})"
                                 f" - (SELECT COALESCE(SUM(amount),0) FROM advance_allocations WHERE kind=? AND invoice_id=h.{c['iid']}))"
                                 f" > 0.01", (kind, kind))])
            add("تنبيه", f"فواتير {label}", "مسدد أكثر من صافي الفاتورة (بعد الخصومات اللاحقة)",
                [r[0] for r in q(f"SELECT COALESCE(h.invoice_no, h.{c['iid']}) FROM {c['inv']} h WHERE h.paid_amount + "
                                 f"(SELECT COALESCE(SUM(total_amount),0) FROM post_discounts WHERE kind=? AND invoice_id=h.{c['iid']})"
                                 f" > h.total_amount + 0.01", (kind,))])
            add("خطأ", f"فواتير {label}", "فاتورة بدون أي بند",
                # فواتير HansaWorld التاريخية مستوردة بالإجمالي فقط (الأصناف في سجل حركات الأصناف المستوردة)
                [r[0] for r in q(f"SELECT COALESCE(invoice_no, {c['iid']}) FROM {c['inv']} h WHERE NOT EXISTS "
                                 f"(SELECT 1 FROM {c['ilines']} l WHERE l.{c['iid']}=h.{c['iid']}) "
                                 f"AND COALESCE(h.source_type,'') != 'HansaWorld'")])
            add("خطأ", f"أوامر {label}", "الكمية المنفذة أو المفوترة تتجاوز كمية الأمر",
                [r[0] for r in q(f"SELECT DISTINCT order_id FROM {c['olines']} WHERE {c['done']} > quantity + {EPS} "
                                 f"OR invoiced_qty > {c['done']} + {EPS}")])
        # 3) السندات
        add("خطأ", "الخزينة", "سند مربوط بفاتورة غير موجودة",
            [r[0] for r in q("SELECT trans_id FROM cash_transactions t WHERE (alloc_type='sales' AND NOT EXISTS "
                             "(SELECT 1 FROM sales_invoices WHERE invoice_id=t.alloc_id)) OR (alloc_type='purchase' AND NOT EXISTS "
                             "(SELECT 1 FROM purchase_invoices WHERE purchase_id=t.alloc_id))")])
        add("خطأ", "الخزينة", "سند بمبلغ صفر أو سالب", [r[0] for r in q("SELECT trans_id FROM cash_transactions WHERE amount <= 0")])
        add("تنبيه", "الخزينة", "سند شريك تجاري لطرف غير مسجل",
            [r[0] for r in q("SELECT trans_id FROM cash_transactions t WHERE category=? AND NOT EXISTS "
                             "(SELECT 1 FROM business_partners b WHERE b.partner_code=t.party_code)", (CAT_PARTNER,))])
        # 4) القيود اليومية
        add("خطأ", "القيود اليومية", "قيد غير متوازن (مدين ≠ دائن)",
            [r[0] for r in q("SELECT entry_id FROM journal_lines GROUP BY entry_id HAVING ABS(SUM(debit)-SUM(credit)) > 0.005")])
        add("خطأ", "القيود اليومية", "قيد بدون أسطر أو بقيمة صفرية",
            [r[0] for r in q("SELECT e.entry_id FROM journal_entries e LEFT JOIN journal_lines l ON l.entry_id=e.entry_id "
                             "GROUP BY e.entry_id HAVING COALESCE(SUM(l.debit),0) < 0.005")])
        add("خطأ", "القيود اليومية", "سطر قيد مدين ودائن معاً أو بقيمة سالبة",
            [r[0] for r in q("SELECT DISTINCT entry_id FROM journal_lines WHERE debit < 0 OR credit < 0 OR "
                             "(debit > 0.005 AND credit > 0.005)")])
        add("خطأ", "القيود اليومية", "سطر على حساب غير موجود في الدليل",
            [r[0] for r in q("SELECT DISTINCT l.entry_id FROM journal_lines l LEFT JOIN chart_of_accounts a "
                             "ON a.account_code=l.account_code WHERE a.account_code IS NULL")])
        add("تنبيه", "القيود اليومية", "ترحيل على حساب تجميعي (رئيسي) بدلاً من حساب فرعي",
            [r[0] for r in q("SELECT DISTINCT l.entry_id FROM journal_lines l JOIN chart_of_accounts a "
                             "ON a.account_code=l.account_code WHERE COALESCE(a.is_group,0)=1")])
        tb = q("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(credit),0) FROM journal_lines")[0]
        add("خطأ", "ميزان المراجعة", "إجمالي المدين لا يساوي إجمالي الدائن",
            [f"{tb[0]:,.2f} ≠ {tb[1]:,.2f}"] if abs(tb[0] - tb[1]) > 0.01 else [])
        # 5) الخصومات اللاحقة
        add("خطأ", "الخصومات اللاحقة", "إشعار بدون قيد آلي",
            [r[0] for r in q("SELECT COALESCE(discount_no, discount_id) FROM post_discounts d WHERE journal_entry_id IS NULL "
                             "OR NOT EXISTS (SELECT 1 FROM journal_entries e WHERE e.entry_id=d.journal_entry_id)")])
        # 5ب) الدفتر العام (الترحيل الآلي)
        if q("SELECT 1 FROM sqlite_master WHERE type='table' AND name='gl_entries'"):
            add("خطأ", "الدفتر العام", "قيد ترحيل آلي غير متوازن",
                [r[0] for r in q("SELECT e.ref FROM gl_entries e JOIN gl_lines l ON l.gl_id=e.gl_id GROUP BY e.gl_id "
                                 "HAVING ABS(SUM(l.debit)-SUM(l.credit)) > 0.005")])
            add("تنبيه", "الدفتر العام", "مستندات بانتظار الترحيل (تُرحَّل تلقائياً عند فتح أي تقرير من الدفتر العام)",
                [f"{r[0]}#{r[1]}" for r in q("SELECT source_type, source_id FROM gl_queue")])
            add("تنبيه", "الدفتر العام", "ترحيل على حساب غير موجود في الدليل",
                sorted({r[0] for r in q("SELECT DISTINCT l.account_code FROM gl_lines l LEFT JOIN chart_of_accounts a "
                                        "ON a.account_code=l.account_code WHERE a.account_code IS NULL")}))
            susp = q("SELECT value FROM app_settings WHERE key='gl_acc_suspense'")
            if susp:
                sb = q("SELECT COALESCE(SUM(debit-credit),0) FROM gl_lines WHERE account_code=?", (susp[0][0],))[0][0]
                add("تنبيه", "الدفتر العام", "رصيد في الحساب المعلق (سندات على خزينة/حساب غير معرّف بالدليل)",
                    [f"{susp[0][0]}: {sb:,.2f}"] if abs(sb) > 0.005 else [])
        # 6) المخزون: الرصيد المخزن = إعادة البناء من الحركات (تجربة داخل معاملة ثم تراجع)
        stored = {r[0]: (r[1] or 0.0, r[2] or 0.0) for r in q("SELECT item_code, quantity, cost_price FROM inventory")}
        try:
            recalc_items(conn, list(stored), allow_negative=True)
            rebuilt = {r[0]: (r[1] or 0.0, r[2] or 0.0) for r in q("SELECT item_code, quantity, cost_price FROM inventory")}
        finally:
            conn.rollback()
        add("خطأ", "المخزون", "رصيد الصنف المسجل يختلف عن رصيد الحركات (يُصلح بإعادة الاحتساب)",
            [f"{k}: {stored[k][0]:g}→{rebuilt[k][0]:g}" for k in stored if abs(stored[k][0] - rebuilt[k][0]) > 0.001])
        add("تنبيه", "المخزون", "متوسط التكلفة المسجل يختلف عن إعادة الاحتساب",
            [k for k in stored if abs(stored[k][1] - rebuilt[k][1]) > 0.01])
        add("خطأ", "المخزون", "رصيد مخزون سالب", [f"{r[0]}: {r[1]:g}" for r in q("SELECT item_code, quantity FROM inventory "
                                                                            f"WHERE quantity < -{EPS}")])
        add("تنبيه", "المخزون", "صنف بتكلفة صفرية مع وجود رصيد",
            [r[0] for r in q(f"SELECT item_code FROM inventory WHERE quantity > {EPS} AND COALESCE(cost_price,0) <= 0")])
        # 7) الترقيم
        for name, table, col in (("sales_invoice", "sales_invoices", "invoice_no"),
                                 ("purchase_invoice", "purchase_invoices", "invoice_no")):
            add("خطأ", "الترقيم", f"أرقام فواتير مكررة ({table})",
                [r[0] for r in q(f"SELECT {col} FROM {table} WHERE COALESCE({col},'')!='' GROUP BY {col} HAVING COUNT(*)>1")])
        add("تنبيه", "العملاء والموردون", "أسماء مكررة داخل نفس الشجرة",
            [r[0] for r in q("SELECT name FROM business_partners GROUP BY name, is_customer, is_supplier HAVING COUNT(*)>1")])
        add("تنبيه", "سجل المناديب", "مناديب بأسماء متقاربة (قد تكون نفس الشخص)",
            [" / ".join(v) for v in _similar_rep_names(conn)])
    finally:
        conn.close()
    order = {"خطأ": 0, "تنبيه": 1, "سليم": 2}
    out.sort(key=lambda r: order[r[0]])
    return out


def _similar_rep_names(conn):
    groups = defaultdict(list)
    for (n,) in conn.execute("SELECT name FROM sales_reps"):
        groups[norm_name(n).replace(" ", "")].append(n)
    return [v for v in groups.values() if len(v) > 1]


def repair_inventory():
    """إصلاح آمن: إعادة بناء أرصدة ومتوسط تكلفة كل الأصناف من الحركات (مع نسخة احتياطية قبلها)."""
    backup_database("before_repair")
    with db() as conn:
        recalc_items(conn, [r[0] for r in conn.execute("SELECT item_code FROM inventory")], allow_negative=True)


def repair_invoice_paid():
    """إصلاح آمن: إعادة احتساب المسدد وحالة كل الفواتير من السندات المربوطة."""
    backup_database("before_repair")
    with db() as conn:
        for kind, c in FLOW.items():
            for (iid,) in conn.execute(f"SELECT {c['iid']} FROM {c['inv']}").fetchall():
                refresh_invoice_paid(conn, kind, iid)


# ============================================================
# محرك التقارير (منفصل عن الواجهة، قابل للتصدير Excel / PDF)
# ============================================================
class Report:
    """columns = [(العنوان، النوع)] حيث النوع: t نص، m مبلغ، q كمية، p نسبة %، i عدد صحيح.
    rows = قائمة tuples بقيم خام. total = صف إجمالي اختياري. chart = (فهرس التسمية، فهرس القيمة)."""

    def __init__(self, title, columns, rows, total=None, notes=None, subtitle="", chart=None, highlight=(), ltr=False):
        self.title, self.columns, self.rows = title, columns, rows
        self.total, self.notes, self.subtitle, self.chart = total, notes or [], subtitle, chart
        # highlight: فهارس أعمدة تُميَّز بلون أصفر (العرض والتصدير). ltr: ترتيب الأعمدة من اليسار (تقارير بعناوين إنجليزية)
        self.highlight, self.ltr = tuple(highlight), ltr
        # row_styles: {فهرس السطر: group (رأس طرف) | sub (إجمالي فرعي) | shared (طرف مشترك) | net (صافي مقاصة)}
        self.row_styles = {}


def fmt_cell(v, kind):
    if v is None or v == "":
        return ""
    if isinstance(v, str):
        return v
    if kind == "m":
        return f"{v:,.2f}"
    if kind == "q":
        return f"{int(v):,}" if abs(v - round(v)) < 1e-9 else f"{v:,.3f}".rstrip("0").rstrip(".")
    if kind == "p":
        return f"{v:,.1f}%"
    if kind == "i":
        return f"{int(v):,}"
    return str(v)


def month_keys(d_from, d_to):
    y, m = int(d_from[:4]), int(d_from[5:7])
    y2, m2 = int(d_to[:4]), int(d_to[5:7])
    keys = []
    while (y, m) <= (y2, m2):
        keys.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return keys


def month_end(key):
    y, m = int(key[:4]), int(key[5:7])
    return date(y + (m == 12), (m % 12) + 1, 1) - timedelta(days=1)


def _pe(col, group):
    return f"substr({col},1,7)" if group == "month" else f"substr({col},1,10)"


def _share_rows(rows, value_idx):
    """يضيف عمودي النسبة والتراكمي لصفوف مرتبة تنازلياً بحسب value_idx."""
    total = sum(max(r[value_idx], 0) for r in rows)
    out, cum = [], 0.0
    for r in rows:
        share = (max(r[value_idx], 0) / total * 100) if total > 0 else 0.0
        cum += share
        out.append(tuple(r) + (share, cum))
    return out


def _pct(a, b):
    return (a / b * 100) if b else 0.0


def partner_balances(as_of, conn=None):
    """أرصدة كل شريك حتى تاريخ معين. الرصيد = مدين - دائن (الموجب = ذمة مدينة على الطرف).
    conn: اتصال مفتوح (مثل معاملة استيراد لم تُحفظ بعد) بدلاً من اتصال جديد."""
    comp = defaultdict(dict)
    queries = [
        ("sales", "SELECT customer_code, SUM(total_amount) FROM sales_invoices WHERE invoice_date<=? GROUP BY customer_code"),
        ("purch", "SELECT supplier_code, SUM(total_amount) FROM purchase_invoices WHERE purchase_date<=? GROUP BY supplier_code"),
        ("ret_s", f"SELECT party_code, SUM(total_amount) FROM returns_invoices WHERE return_type='{SALES_RETURN}' "
                  "AND return_date<=? GROUP BY party_code"),
        ("ret_p", f"SELECT party_code, SUM(total_amount) FROM returns_invoices WHERE return_type='{PURCHASE_RETURN}' "
                  "AND return_date<=? GROUP BY party_code"),
        ("rec", f"SELECT party_code, SUM(amount) FROM cash_transactions WHERE trans_type='قبض' AND party_code IS NOT NULL "
                f"AND category='{CAT_PARTNER}' AND trans_date<=? GROUP BY party_code"),
        ("pay", f"SELECT party_code, SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND party_code IS NOT NULL "
                f"AND category='{CAT_PARTNER}' AND trans_date<=? GROUP BY party_code"),
        ("disc_p", "SELECT party_code, SUM(total_amount) FROM post_discounts WHERE kind='purchase' AND discount_date<=? "
                   "GROUP BY party_code"),
        ("disc_s", "SELECT party_code, SUM(total_amount) FROM post_discounts WHERE kind='sales' AND discount_date<=? "
                   "GROUP BY party_code"),
    ]
    with (nullcontext(conn) if conn is not None else db()) as conn:
        for key, sql in queries:
            for code, v in conn.execute(sql, (as_of,)):
                comp[code][key] = v or 0.0
        for code, od, oc in conn.execute(
                "SELECT partner_code, COALESCE(opening_debit,0), COALESCE(opening_credit,0) FROM business_partners "
                "WHERE (opening_debit > 0 OR opening_credit > 0) AND COALESCE(opening_date,'0000-01-01') <= ?", (as_of,)):
            comp[code]["open"] = od - oc
        info = {c: (n, ph or "") for c, n, ph in conn.execute("SELECT partner_code, name, phone FROM business_partners")}
    out = []
    for code, c in comp.items():
        g = lambda k: c.get(k, 0.0)
        out.append(dict(code=code, name=info.get(code, (code, ""))[0], phone=info.get(code, ("", ""))[1], opening=g("open"),
                        sales=g("sales"), purch=g("purch"), ret_s=g("ret_s"), ret_p=g("ret_p"), rec=g("rec"), pay=g("pay"),
                        disc_s=g("disc_s"), disc_p=g("disc_p"),
                        balance=(g("open") + g("sales") + g("ret_p") + g("pay") + g("disc_p")
                                 - g("purch") - g("ret_s") - g("rec") - g("disc_s"))))
    return out


def rep_receivables(p):
    data = [r for r in partner_balances(p["d_to"]) if r["balance"] > 0.005]
    data.sort(key=lambda r: r["balance"], reverse=True)
    rows = [(r["code"], r["name"], r["phone"], r["opening"], r["sales"], r["ret_s"] + r["disc_s"], r["rec"],
             r["purch"] - r["ret_p"] - r["pay"] - r["disc_p"], r["balance"]) for r in data]
    rows = _share_rows(rows, 8)
    cols = [("الكود", "t"), ("العميل", "t"), ("الهاتف", "t"), ("رصيد افتتاحي", "m"), ("فواتير المبيعات", "m"),
            ("(-) مرتجعات وإشعارات دائنة", "m"), ("(-) مقبوضات", "m"), ("(-) صافي مشتريات منه", "m"), ("الرصيد المدين", "m"),
            ("النسبة", "p"), ("التراكمي", "p")]
    tot = sum(r[8] for r in rows)
    total = ("", "الإجمالي", "", *[sum(r[i] for r in rows) for i in range(3, 8)], tot, None, None)
    notes = [f"عدد العملاء المدينين: {len(rows)}   |   إجمالي الذمم المدينة: {fmt(tot)} ر.س"]
    if rows:
        top3 = sum(r[8] for r in rows[:3])
        notes.append(f"أكبر 3 مدينين يمثلون {_pct(top3, tot):.1f}% من إجمالي المديونية")
    return Report("الذمم المدينة / أرصدة العملاء", cols, rows, total, notes, f"الرصيد حتى {p['d_to']}", chart=(1, 8))


def rep_payables(p):
    data = [r for r in partner_balances(p["d_to"]) if r["balance"] < -0.005]
    data.sort(key=lambda r: r["balance"])
    rows = [(r["code"], r["name"], r["phone"], -r["opening"], r["purch"], r["ret_p"] + r["disc_p"], r["pay"],
             r["sales"] - r["ret_s"] - r["rec"] - r["disc_s"], -r["balance"]) for r in data]
    rows = _share_rows(rows, 8)
    cols = [("الكود", "t"), ("المورد", "t"), ("الهاتف", "t"), ("رصيد افتتاحي (له)", "m"), ("فواتير المشتريات", "m"),
            ("(-) مرتجعات وخصومات لاحقة", "m"), ("(-) مدفوعات", "m"), ("(-) صافي مبيعات له", "m"), ("الرصيد الدائن", "m"),
            ("النسبة", "p"), ("التراكمي", "p")]
    tot = sum(r[8] for r in rows)
    total = ("", "الإجمالي", "", *[sum(r[i] for r in rows) for i in range(3, 8)], tot, None, None)
    notes = [f"عدد الموردين الدائنين: {len(rows)}   |   إجمالي الذمم الدائنة: {fmt(tot)} ر.س"]
    return Report("الذمم الدائنة / أرصدة الموردين", cols, rows, total, notes, f"الرصيد حتى {p['d_to']}", chart=(1, 8))


# ---------------------------------------------------------- تقارير العملاء والموردين: إجماليات / تفاصيل مجمّعة / مقاصة
PT_MODES = {"summary": "إجماليات", "detail": "تفاصيل مجمّعة"}
PT_DETAIL_COLS = [("الكود", "t"), ("الاسم", "t"), ("التاريخ", "t"), ("رقم المستند", "t"), ("نوع الحركة", "t"),
                  ("البيان", "t"), ("مدين (عليه)", "m"), ("دائن (له)", "m"), ("الرصيد", "m")]


def _pt_opts(p):
    mode = p.get("mode") or "summary"
    if mode not in PT_MODES:
        raise UserError("وضع التقرير غير معروف (إجماليات أو تفاصيل).")
    return mode, bool(p.get("nonzero")), bool(p.get("shared_only"))


def _pt_parties(conn):
    """{الكود: (الاسم، عميل؟، مورد؟، الهاتف، الرقم الضريبي)} للأطراف الظاهرة (بدون الأكواد القديمة المؤرشفة)."""
    return {r[0]: r[1:] for r in conn.execute(
        "SELECT partner_code, name, is_customer, is_supplier, COALESCE(phone,''), COALESCE(vat_number,'') "
        f"FROM business_partners WHERE {ACTIVE_PARTNER_SQL} ORDER BY partner_code")}


def shared_partners(conn, parts=None):
    """الأطراف المشتركة بين العملاء والموردين: {الكود: (أكواد الطرف الآخر، السبب)}.
    السبب: «نفس الكود» (كود واحد عميل ومورد معاً) أو «نفس الاسم» أو «نفس الرقم الضريبي» (كودان مرتبطان)."""
    parts = parts if parts is not None else _pt_parties(conn)
    links = partner_link_map(conn)
    out = {}
    for code, (name, is_c, is_s, _ph, vat) in parts.items():
        group = [c for c in links.get(code, [code]) if c in parts]
        others = [c for c in group if c != code and (parts[c][2] if is_c else parts[c][1])]
        if is_c and is_s:
            out[code] = (others, "نفس الكود")
        elif others:
            same_name = any(norm_name(parts[c][0]) == norm_name(name) for c in others)
            out[code] = (others, "نفس الاسم" if same_name else "نفس الرقم الضريبي")
    return out


def _pt_side_entries(conn, code, side_of_b):
    """حركات كود واحد مقسومة: {'c': حركات كعميل، 's': حركات كمورد}. الحركة المشتركة (b) تذهب للجانب side_of_b."""
    out = {"c": [], "s": []}
    for e in _partner_entries(conn, [code]):
        out[side_of_b if e[8] == "b" else e[8]].append(e)
    for v in out.values():
        v.sort(key=lambda e: (e[0] or "", e[1]))
    return out


def _pt_block(rows, styles, entries, d_from, d_to, code, name, sign=1, head=None, sub_label=None):
    """يضيف حركات طرف (أو جانب منه) برصيد متحرك وإجمالي فرعي. sign=-1 لعرض رصيد المورد موجباً (له).
    يرجع (رصيد سابق، مدين، دائن، الرصيد، عدد الحركات)."""
    prior, inside = _split_period(entries, d_from, d_to)
    if head:
        styles[len(rows)] = "group"
        rows.append((code, head, "", "", "", "", None, None, None))
    bal = prior * sign
    if abs(prior) > 0.005:
        rows.append(("", "", d_from or "", "سابق", "رصيد ما قبل الفترة", f"مُرحَّل حتى {d_from}", None, None, bal))
    deb = cred = 0.0
    for d, _, ref, typ, desc, debit, credit, *_ in inside:
        bal += (debit - credit) * sign
        deb += debit
        cred += credit
        rows.append(("", "", d or "-", ref, typ, desc, debit or None, credit or None, round(bal, 2)))
    styles[len(rows)] = "sub"
    rows.append((code, sub_label or f"إجمالي {name}", "", "", f"{len(inside)} حركة", "", round(deb, 2), round(cred, 2),
                 round(bal, 2)))
    return prior * sign, deb, cred, round(bal, 2), len(inside)


def _rep_party_role(p, role):
    mode, nonzero, shared_only = _pt_opts(p)
    who, many = ("العميل", "العملاء") if role == "customer" else ("المورد", "الموردين")
    sign = 1 if role == "customer" else -1      # الرصيد الموجب: عليه للعميل، له للمورد
    d_from, d_to = p.get("d_from"), p.get("d_to")
    with db() as conn:
        parts = _pt_parties(conn)
        shared = shared_partners(conn, parts)
        codes = [c for c, v in parts.items() if (v[1] if role == "customer" else v[2])]
        if shared_only:
            codes = [c for c in codes if c in shared]
        side = "c" if role == "customer" else "s"
        data = {c: _pt_side_entries(conn, c, side)[side] for c in codes}
    period = period_label(d_from, d_to)
    filt = (" • الأرصدة غير الصفرية فقط" if nonzero else "") + (" • الأطراف المشتركة فقط" if shared_only else "")
    bal_name = "الرصيد (عليه)" if role == "customer" else "الرصيد (له)"
    rows, styles = [], {}
    if mode == "summary":
        for code in codes:
            prior, inside = _split_period(data[code], d_from, d_to)
            deb = round(sum(e[5] for e in inside), 2)
            cred = round(sum(e[6] for e in inside), 2)
            bal = round((prior + deb - cred) * sign, 2)
            if not inside and abs(prior) < 0.005 or nonzero and abs(bal) < 0.005:
                continue
            sh = shared.get(code)
            if sh:
                styles[len(rows)] = "shared"
            rows.append((code, parts[code][0], parts[code][3] or "-", round(prior * sign, 2), deb, cred, bal,
                         "مدين (عليه)" if bal * sign > 0.005 else ("دائن (له)" if bal * sign < -0.005 else "-"), len(inside),
                         max((e[0] or "" for e in inside), default="") or "-",
                         f"{sh[1]}: {' ، '.join(sh[0]) or code}" if sh else ""))
        cols = [("الكود", "t"), (who, "t"), ("الهاتف", "t"), ("رصيد سابق", "m"), ("مدين الفترة", "m"), ("دائن الفترة", "m"),
                (bal_name, "m"), ("الحالة", "t"), ("عدد الحركات", "i"), ("آخر حركة", "t"), ("مشترك مع", "t")]
        total = ("", "الإجمالي", "", *[round(sum(r[i] for r in rows), 2) for i in (3, 4, 5, 6)], "",
                 sum(r[8] for r in rows), "", "")
        notes = [f"عدد {many}: {len(rows)}   |   صافي الأرصدة: {fmt(total[6])} ر.س" + filt,
                 "«مشترك مع» = نفس الكيان عميل ومورد (نفس الكود أو مرتبط بالاسم/الرقم الضريبي) — راجع تقرير المقاصة."]
        rep = Report(f"تقرير {many} — إجماليات", cols, rows, total, notes, period)
    else:
        tot = [0.0, 0.0, 0.0, 0]
        for code in codes:
            prior, inside = _split_period(data[code], d_from, d_to)
            bal = (prior + sum(e[5] - e[6] for e in inside)) * sign
            if not inside and abs(prior) < 0.005 or nonzero and abs(bal) < 0.005:
                continue
            sh = shared.get(code)
            head = parts[code][0] + (f"   ⇄ مشترك ({sh[1]}: {' ، '.join(sh[0]) or code})" if sh else "")
            _, deb, cred, bal, n = _pt_block(rows, styles, data[code], d_from, d_to, code, parts[code][0], sign, head=head)
            tot[0] += deb
            tot[1] += cred
            tot[2] += bal
            tot[3] += 1
        cols = PT_DETAIL_COLS[:-1] + [(bal_name, "m")]
        total = ("", f"الإجمالي العام ({tot[3]} {who if tot[3] == 1 else many})", "", "", "", "", round(tot[0], 2),
                 round(tot[1], 2), round(tot[2], 2))
        notes = [f"كل {who} بكل حركاته ثم إجمالي فرعي، ثم الإجمالي العام." + filt,
                 "الرصيد متحرك داخل كل طرف ويبدأ من رصيد ما قبل الفترة."]
        rep = Report(f"تقرير {many} — تفاصيل مجمّعة", cols, rows, total, notes, period)
    rep.row_styles = styles
    return rep


def rep_customers_report(p):
    return _rep_party_role(p, "customer")


def rep_suppliers_report(p):
    return _rep_party_role(p, "supplier")


def rep_partner_netting(p):
    """العملاء والموردون معاً للمقاصة: لكل كيان رصيده كعميل (عليه) ورصيده كمورد (له) وصافي المقاصة."""
    mode, nonzero, shared_only = _pt_opts(p)
    d_from, d_to = p.get("d_from"), p.get("d_to")
    with db() as conn:
        parts = _pt_parties(conn)
        shared = shared_partners(conn, parts)
        links = partner_link_map(conn)
        groups, seen = [], set()
        for code in parts:
            if code in seen:
                continue
            g = [c for c in links.get(code, [code]) if c in parts]
            seen.update(g)
            groups.append(g)
        sides = {c: _pt_side_entries(conn, c, "c" if parts[c][1] else "s") for c in parts}
    period = period_label(d_from, d_to)
    filt = (" • الأرصدة غير الصفرية فقط" if nonzero else "") + (" • الأطراف المشتركة فقط" if shared_only else "")
    items = []
    for g in groups:
        reason = next((shared[c][1] for c in g if c in shared), "")
        if shared_only and not reason:
            continue
        c_codes = [c for c in g if parts[c][1]]
        s_codes = [c for c in g if parts[c][2]]
        c_ent = sorted((e for c in g for e in sides[c]["c"]), key=lambda e: (e[0] or "", e[1]))
        s_ent = sorted((e for c in g for e in sides[c]["s"]), key=lambda e: (e[0] or "", e[1]))
        cp, ci = _split_period(c_ent, d_from, d_to)
        sp, si = _split_period(s_ent, d_from, d_to)
        c_bal = round(cp + sum(e[5] - e[6] for e in ci), 2)
        s_bal = round(-(sp + sum(e[5] - e[6] for e in si)), 2)
        if not (ci or si) and abs(cp) < 0.005 and abs(sp) < 0.005:
            continue
        if nonzero and abs(c_bal) < 0.005 and abs(s_bal) < 0.005:
            continue
        items.append(dict(g=g, c=c_codes, s=s_codes, name=parts[g[0]][0], reason=reason, c_ent=c_ent, s_ent=s_ent,
                          c_bal=c_bal, s_bal=s_bal, net=round(c_bal - s_bal, 2)))
    items.sort(key=lambda it: (not it["reason"], it["g"][0]))
    rows, styles = [], {}
    if mode == "summary":
        for it in items:
            if it["reason"]:
                styles[len(rows)] = "shared"
            net = it["net"]
            rows.append((" ، ".join(it["c"]) or "-", " ، ".join(it["s"]) or "-", it["name"], it["reason"] or "-",
                         it["c_bal"], it["s_bal"], net,
                         "عليه (مدين لنا)" if net > 0.005 else ("له (دائن)" if net < -0.005 else "متوازن")))
        cols = [("كود العميل", "t"), ("كود المورد", "t"), ("الاسم", "t"), ("مشترك (السبب)", "t"),
                ("رصيد العميل (عليه)", "m"), ("رصيد المورد (له)", "m"), ("صافي المقاصة", "m"), ("الوضع", "t")]
        n_sh = sum(1 for it in items if it["reason"])
        total = ("", "", "الإجمالي", f"{n_sh} مشترك", *[round(sum(r[i] for r in rows), 2) for i in (4, 5, 6)], "")
        notes = [f"عدد الأطراف: {len(rows)} (منها {n_sh} مشترك بين العملاء والموردين، مظللة في الأعلى)" + filt,
                 "صافي المقاصة = رصيد العميل (عليه) − رصيد المورد (له): الموجب مستحق لنا والسالب مستحق له.",
                 "الأطراف المشتركة تُكتشف بنفس الكود أو بالاسم الموحد أو بالرقم الضريبي."]
        rep = Report("العملاء والموردون — المقاصة (إجماليات)", cols, rows, total, notes, period)
    else:
        tot_c = tot_s = 0.0
        for it in items:
            codes = " ، ".join(it["g"])
            styles[len(rows)] = "shared" if it["reason"] else "group"
            rows.append((codes, it["name"] + (f"   ⇄ مشترك ({it['reason']})" if it["reason"] else ""), "", "", "", "",
                         None, None, None))
            if it["c"]:
                _pt_block(rows, styles, it["c_ent"], d_from, d_to, " ، ".join(it["c"]), it["name"],
                          head="— حركاته كعميل", sub_label="إجمالي كعميل (عليه)")
            if it["s"]:
                _pt_block(rows, styles, it["s_ent"], d_from, d_to, " ، ".join(it["s"]), it["name"], sign=-1,
                          head="— حركاته كمورد", sub_label="إجمالي كمورد (له)")
            styles[len(rows)] = "net"
            net = it["net"]
            rows.append((codes, "صافي المقاصة", "", "", "عليه" if net > 0.005 else ("له" if net < -0.005 else "متوازن"),
                         f"عليه {fmt(it['c_bal'])} − له {fmt(it['s_bal'])}", None, None, net))
            tot_c += it["c_bal"]
            tot_s += it["s_bal"]
        cols = PT_DETAIL_COLS[:-1] + [("الرصيد / الصافي", "m")]
        total = ("", f"الإجمالي العام ({len(items)} طرف)", "", "", "", f"عليه {fmt(tot_c)} − له {fmt(tot_s)}",
                 None, None, round(tot_c - tot_s, 2))
        notes = ["لكل طرف: حركاته كعميل (الرصيد الموجب عليه) ثم حركاته كمورد (الرصيد الموجب له) ثم صافي المقاصة." + filt,
                 "الأطراف المشتركة (نفس الكود أو الاسم أو الرقم الضريبي) مظللة وتظهر أولاً."]
        rep = Report("العملاء والموردون — المقاصة (تفاصيل)", cols, rows, total, notes, period)
    rep.row_styles = styles
    return rep


def rep_aging(kind, p):
    """أعمار الديون: يُخصم المسدد والمرتجعات من أقدم فاتورة أولاً (FIFO)."""
    as_of = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    if kind == "customer":
        inv_sql = ("SELECT customer_code, invoice_date, total_amount FROM sales_invoices WHERE invoice_date<=? "
                   "ORDER BY invoice_date, invoice_id")
        cred = [f"SELECT party_code, SUM(amount) FROM cash_transactions WHERE trans_type='قبض' AND party_code IS NOT NULL "
                f"AND category='{CAT_PARTNER}' AND trans_date<=? GROUP BY party_code",
                f"SELECT party_code, SUM(total_amount) FROM returns_invoices WHERE return_type='{SALES_RETURN}' "
                "AND return_date<=? GROUP BY party_code",
                "SELECT party_code, SUM(total_amount) FROM post_discounts WHERE kind='sales' AND discount_date<=? "
                "GROUP BY party_code"]
        title, who, excess_label = "أعمار ديون العملاء", "العميل", "رصيد دائن (دفعات مقدمة)"
    else:
        inv_sql = ("SELECT supplier_code, purchase_date, total_amount FROM purchase_invoices WHERE purchase_date<=? "
                   "ORDER BY purchase_date, purchase_id")
        cred = [f"SELECT party_code, SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND party_code IS NOT NULL "
                f"AND category='{CAT_PARTNER}' AND trans_date<=? GROUP BY party_code",
                f"SELECT party_code, SUM(total_amount) FROM returns_invoices WHERE return_type='{PURCHASE_RETURN}' "
                "AND return_date<=? GROUP BY party_code",
                "SELECT party_code, SUM(total_amount) FROM post_discounts WHERE kind='purchase' AND discount_date<=? "
                "GROUP BY party_code"]
        title, who, excess_label = "أعمار ديون الموردين", "المورد", "رصيد مدين (دفعات مقدمة منا)"
    invoices, credits = defaultdict(list), defaultdict(float)
    with db() as conn:
        for code, d, amt in conn.execute(inv_sql, (p["d_to"],)):
            invoices[code].append((d, amt))
        # الرصيد الافتتاحي: المدين يُعامل كأقدم فاتورة، والدائن كدفعة مقدمة
        own, other = ("opening_debit", "opening_credit") if kind == "customer" else ("opening_credit", "opening_debit")
        for code, od, amt_own, amt_other in conn.execute(
                f"SELECT partner_code, COALESCE(opening_date,''), COALESCE({own},0), COALESCE({other},0) "
                f"FROM business_partners WHERE ({own} > 0 OR {other} > 0) AND COALESCE(opening_date,'0000-01-01') <= ?",
                (p["d_to"],)):
            if amt_own > 0.005:
                invoices[code].insert(0, (od or "2000-01-01", amt_own))
            if amt_other > 0.005:
                credits[code] += amt_other
        for sql in cred:
            for code, v in conn.execute(sql, (p["d_to"],)):
                credits[code] += v or 0.0
        names = dict(conn.execute("SELECT partner_code, name FROM business_partners"))
    rows = []
    for code, lst in invoices.items():
        rem_credit = credits.get(code, 0.0)
        buckets = [0.0, 0.0, 0.0, 0.0]
        for d, amt in lst:
            applied = min(rem_credit, amt)
            rem_credit -= applied
            open_amt = amt - applied
            if open_amt > 0.005:
                age = (as_of - datetime.strptime(d, "%Y-%m-%d").date()).days
                buckets[0 if age <= 30 else 1 if age <= 60 else 2 if age <= 90 else 3] += open_amt
        total_open = sum(buckets)
        if total_open > 0.005 or rem_credit > 0.005:
            rows.append((code, names.get(code, code), *buckets, total_open, max(rem_credit, 0.0)))
    for code, cr in credits.items():
        if code not in invoices and cr > 0.005 and code in names:
            rows.append((code, names[code], 0.0, 0.0, 0.0, 0.0, 0.0, cr))
    rows.sort(key=lambda r: r[6], reverse=True)
    cols = [("الكود", "t"), (who, "t"), ("0-30 يوم", "m"), ("31-60 يوم", "m"), ("61-90 يوم", "m"),
            ("أكثر من 90 يوم", "m"), ("إجمالي المستحق", "m"), (excess_label, "m")]
    sums = [sum(r[i] for r in rows) for i in range(2, 8)]
    total = ("", "الإجمالي", *sums)
    notes = ["يُخصم المسدد والمرتجعات والخصومات اللاحقة من أقدم فاتورة أولاً. الفئات حسب عمر الفاتورة من تاريخ الإصدار حتى تاريخ التقرير."]
    if sums[4] > 0:
        notes.append(f"الديون الأقدم من 90 يوماً: {fmt(sums[3])} ر.س ({_pct(sums[3], sums[4]):.1f}% من الإجمالي) - تحتاج متابعة")
    return Report(title, cols, rows, total, notes, f"حتى {p['d_to']}", chart=(1, 6))


def rep_sales_period(p):
    g, a, b = p["group"], p["d_from"], p["d_to"]
    with db() as conn:
        s = {k: (n, sub or 0, vat or 0) for k, n, sub, vat in conn.execute(
            f"SELECT {_pe('invoice_date', g)}, COUNT(*), SUM(subtotal), SUM(vat_amount) FROM sales_invoices "
            f"WHERE invoice_date BETWEEN ? AND ? GROUP BY 1", (a, b))}
        rt = {k: (sub or 0, cost or 0) for k, sub, cost in conn.execute(
            f"SELECT {_pe('return_date', g)}, SUM(subtotal), SUM(quantity*unit_cost) FROM returns_invoices "
            f"WHERE return_type=? AND return_date BETWEEN ? AND ? GROUP BY 1", (SALES_RETURN, a, b))}
        cg = {k: c or 0 for k, c in conn.execute(
            f"SELECT {_pe('h.invoice_date', g)}, SUM(l.quantity*l.unit_cost) FROM sales_invoice_lines l "
            f"JOIN sales_invoices h ON h.invoice_id=l.invoice_id WHERE h.invoice_date BETWEEN ? AND ? GROUP BY 1", (a, b))}
    rows, prev = [], None
    for k in sorted(set(s) | set(rt)):
        n, sub, vat = s.get(k, (0, 0, 0))
        rsub, rcost = rt.get(k, (0, 0))
        net, cost = sub - rsub, cg.get(k, 0) - rcost
        profit = net - cost
        growth = _pct(net - prev, prev) if prev else None
        rows.append((k, n, sub, rsub, net, cost, profit, _pct(profit, net), growth, vat))
        prev = net
    cols = [("الفترة", "t"), ("عدد الفواتير", "i"), ("المبيعات (قبل الضريبة)", "m"), ("المرتجعات", "m"), ("صافي المبيعات", "m"),
            ("التكلفة", "m"), ("مجمل الربح", "m"), ("الهامش", "p"), ("التغير عن السابق", "p"), ("الضريبة", "m")]
    t = [sum(r[i] for r in rows) for i in (1, 2, 3, 4, 5, 6, 9)]
    total = ("الإجمالي", t[0], t[1], t[2], t[3], t[4], t[5], _pct(t[5], t[3]), None, t[6])
    notes = []
    if rows:
        best = max(rows, key=lambda r: r[4])
        notes.append(f"أعلى فترة مبيعاً: {best[0]} بصافي {fmt(best[4])} ر.س   |   متوسط الفترة: {fmt(t[3] / len(rows))} ر.س")
    return Report("المبيعات والإيرادات حسب الفترة", cols, rows, total, notes, f"من {a} إلى {b}", chart=(0, 4))


def rep_sales_customer(p):
    a, b = p["d_from"], p["d_to"]
    with db() as conn:
        s = {c: (n, name, sub or 0) for c, name, n, sub in conn.execute(
            "SELECT customer_code, MAX(customer_name), COUNT(*), SUM(subtotal) FROM sales_invoices "
            "WHERE invoice_date BETWEEN ? AND ? GROUP BY customer_code", (a, b))}
        cg = {c: v or 0 for c, v in conn.execute(
            "SELECT h.customer_code, SUM(l.quantity*l.unit_cost) FROM sales_invoice_lines l JOIN sales_invoices h "
            "ON h.invoice_id=l.invoice_id WHERE h.invoice_date BETWEEN ? AND ? GROUP BY h.customer_code", (a, b))}
        rt = {c: [name, sub or 0, cost or 0] for c, name, sub, cost in conn.execute(
            "SELECT party_code, MAX(party_name), SUM(subtotal), SUM(quantity*unit_cost) FROM returns_invoices "
            "WHERE return_type=? AND return_date BETWEEN ? AND ? GROUP BY party_code", (SALES_RETURN, a, b))}
        for c, name, sub in conn.execute("SELECT party_code, MAX(party_name), SUM(subtotal) FROM post_discounts "
                                         "WHERE kind='sales' AND discount_date BETWEEN ? AND ? GROUP BY party_code", (a, b)):
            rt.setdefault(c, [name, 0, 0])[1] += sub or 0
    rows = []
    for c in set(s) | set(rt):
        n, name, sub = s.get(c, (0, rt[c][0] if c in rt else c, 0))
        rname, rsub, rcost = rt.get(c, (name, 0, 0))
        net, cost = sub - rsub, cg.get(c, 0) - rcost
        rows.append((c, name, n, sub, rsub, net, cost, net - cost, _pct(net - cost, net)))
    rows.sort(key=lambda r: r[5], reverse=True)
    rows = _share_rows([r[:9] for r in rows], 5)
    cols = [("الكود", "t"), ("العميل", "t"), ("عدد الفواتير", "i"), ("المبيعات", "m"), ("المرتجعات والإشعارات الدائنة", "m"),
            ("صافي المبيعات", "m"),
            ("التكلفة", "m"), ("الربح", "m"), ("الهامش", "p"), ("النسبة", "p"), ("التراكمي", "p")]
    t = [sum(r[i] for r in rows) for i in (2, 3, 4, 5, 6, 7)]
    total = ("", "الإجمالي", t[0], t[1], t[2], t[3], t[4], t[5], _pct(t[5], t[3]), None, None)
    notes = [f"عدد العملاء المشترين: {len(rows)}"]
    if rows and t[3] > 0:
        n80 = next((i + 1 for i, r in enumerate(rows) if r[10] >= 80), len(rows))
        notes.append(f"{n80} عميل فقط يحققون 80% من صافي المبيعات (قاعدة باريتو) - ركّز عليهم")
    return Report("المبيعات حسب العميل", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 5))


def rep_sales_item(p):
    a, b = p["d_from"], p["d_to"]
    with db() as conn:
        s = {c: (name, unit, q or 0, sub or 0, cost or 0) for c, name, unit, q, sub, cost in conn.execute(
            "SELECT l.item_code, MAX(l.item_name), MAX(l.unit), SUM(l.quantity), SUM(l.subtotal), SUM(l.quantity*l.unit_cost) "
            "FROM sales_invoice_lines l JOIN sales_invoices h ON h.invoice_id=l.invoice_id "
            "WHERE h.invoice_date BETWEEN ? AND ? GROUP BY l.item_code", (a, b))}
        rt = {c: (name, q or 0, sub or 0, cost or 0) for c, name, q, sub, cost in conn.execute(
            "SELECT item_code, MAX(item_name), SUM(quantity), SUM(subtotal), SUM(quantity*unit_cost) FROM returns_invoices "
            "WHERE return_type=? AND return_date BETWEEN ? AND ? GROUP BY item_code", (SALES_RETURN, a, b))}
    rows = []
    for c in set(s) | set(rt):
        name, unit, q, sub, cost = s.get(c, (rt[c][0] if c in rt else c, "", 0, 0, 0))
        _, rq, rsub, rcost = rt.get(c, (name, 0, 0, 0))
        net, ncost = sub - rsub, cost - rcost
        rows.append((c, name, unit or "", q - rq, (sub / q) if q else 0.0, sub, rsub, net, ncost, net - ncost, _pct(net - ncost, net)))
    rows.sort(key=lambda r: r[7], reverse=True)
    rows = _share_rows(rows, 7)
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("الكمية الصافية", "q"), ("متوسط سعر البيع", "m"),
            ("المبيعات", "m"), ("المرتجعات", "m"), ("صافي المبيعات", "m"), ("التكلفة", "m"), ("الربح", "m"),
            ("الهامش", "p"), ("النسبة", "p"), ("التراكمي", "p")]
    t = [sum(r[i] for r in rows) for i in (5, 6, 7, 8, 9)]
    total = ("", "الإجمالي", "", None, None, t[0], t[1], t[2], t[3], t[4], _pct(t[4], t[2]), None, None)
    notes = [f"عدد الأصناف المباعة: {len(rows)}"]
    losers = [r for r in rows if r[9] < 0]
    if losers:
        notes.append(f"تنبيه: {len(losers)} صنف مباع بخسارة (سعر أقل من التكلفة) - راجع التسعير")
    return Report("المبيعات والربحية حسب الصنف", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 7))


def _depreciation_in_month(assets, key):
    m_end = month_end(key)
    prev_end = date(int(key[:4]), int(key[5:7]), 1) - timedelta(days=1)
    return sum(accumulated_depreciation(c, s, l, d, m_end) - accumulated_depreciation(c, s, l, d, prev_end)
               for c, s, l, d in assets)


def rep_monthly_pnl(p):
    keys = month_keys(p["d_from"], p["d_to"])
    k1, k2 = keys[0], keys[-1]
    with db() as conn:
        sales = {k: v or 0 for k, v in conn.execute(
            "SELECT substr(invoice_date,1,7), SUM(subtotal) FROM sales_invoices WHERE substr(invoice_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2))}
        cg = {k: v or 0 for k, v in conn.execute(
            "SELECT substr(h.invoice_date,1,7), SUM(l.quantity*l.unit_cost) FROM sales_invoice_lines l JOIN sales_invoices h "
            "ON h.invoice_id=l.invoice_id WHERE substr(h.invoice_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2))}
        rt = {k: (sub or 0, cost or 0) for k, sub, cost in conn.execute(
            "SELECT substr(return_date,1,7), SUM(subtotal), SUM(quantity*unit_cost) FROM returns_invoices WHERE return_type=? "
            "AND substr(return_date,1,7) BETWEEN ? AND ? GROUP BY 1", (SALES_RETURN, k1, k2))}
        ex = defaultdict(float, {k: v or 0 for k, v in conn.execute(
            "SELECT substr(trans_date,1,7), SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND category=? "
            "AND substr(trans_date,1,7) BETWEEN ? AND ? GROUP BY 1", (CAT_EXPENSE, k1, k2))})
        oi = defaultdict(float)
        for k, typ, _, _, amt in journal_pl(conn, f"{k1}-01", month_end(k2).strftime("%Y-%m-%d"), by_month=True):
            (ex if typ == TYPE_EXP else oi)[k] += amt
        dsc = discount_totals(conn, f"{k1}-01", month_end(k2).strftime("%Y-%m-%d"), by_month=True)
        hw = hw_cogs(conn, f"{k1}-01", month_end(k2).strftime("%Y-%m-%d"), by_month=True)
        for k, v in list(dsc.items()):
            oi[k] += v["earned"]
        for k, v in list(hw.items()):
            oi[k] -= v["credits"]
        assets = conn.execute("SELECT purchase_cost, salvage_value, useful_life_years, purchase_date FROM fixed_assets").fetchall()
    rows = []
    for k in keys:
        rsub, rcost = rt.get(k, (0, 0))
        net = sales.get(k, 0) - rsub - (dsc[k]["sales_sub"] if k in dsc else 0)
        cogs = cg.get(k, 0) - rcost + (hw[k]["cogs"] if k in hw else 0)
        gross = net - cogs
        dep = _depreciation_in_month(assets, k)
        expenses, other = ex.get(k, 0), oi.get(k, 0)
        profit = gross + other - expenses - dep
        rows.append((k, net, cogs, gross, _pct(gross, net), other, expenses, dep, profit, _pct(profit, net)))
    cols = [("الشهر", "t"), ("صافي المبيعات", "m"), ("تكلفة المبيعات", "m"), ("مجمل الربح", "m"), ("هامش المجمل", "p"),
            ("إيرادات أخرى", "m"), ("المصروفات التشغيلية", "m"), ("مصروف الإهلاك", "m"), ("صافي الربح", "m"), ("هامش الصافي", "p")]
    t = [sum(r[i] for r in rows) for i in (1, 2, 3, 5, 6, 7, 8)]
    total = ("الإجمالي", t[0], t[1], t[2], _pct(t[2], t[0]), t[3], t[4], t[5], t[6], _pct(t[6], t[0]))
    notes = ["تُحسب الأشهر كاملة. المصروفات = سندات الصرف المصنفة «مصروفات عمومية» + القيود اليومية على حسابات المصروفات. "
             "صافي المبيعات بعد المرتجعات والإشعارات الدائنة. الإيرادات الأخرى من القيود اليومية + الخصم المكتسب من الموردين. "
             "الربح = مجمل الربح + الإيرادات الأخرى - المصروفات - الإهلاك."]
    if hw:
        notes.append("فترة HansaWorld المستوردة: تكلفة المبيعات بالجرد الدوري = المشتريات − إشعارات خصم الموردين، ويُطرح "
                     "المخزون الختامي في شهر الترحيل (لا توجد أرصدة مخزون شهرية في الملفات، فتوزيع التكلفة على الأشهر تقريبي).")
    loss = [r[0] for r in rows if r[8] < 0]
    if loss:
        notes.append("أشهر بخسارة: " + "، ".join(loss))
    return Report("قائمة الدخل الشهرية", cols, rows, total, notes, f"من {k1} إلى {k2}", chart=(0, 8))


def rep_collection(p):
    keys = month_keys(p["d_from"], p["d_to"])
    k1, k2 = keys[0], keys[-1]
    with db() as conn:
        billed = {k: v or 0 for k, v in conn.execute(
            "SELECT substr(invoice_date,1,7), SUM(total_amount) FROM sales_invoices WHERE substr(invoice_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2))}
        rets = {k: v or 0 for k, v in conn.execute(
            "SELECT substr(return_date,1,7), SUM(total_amount) FROM returns_invoices WHERE return_type=? "
            "AND substr(return_date,1,7) BETWEEN ? AND ? GROUP BY 1", (SALES_RETURN, k1, k2))}
        coll = {k: v or 0 for k, v in conn.execute(
            "SELECT substr(trans_date,1,7), SUM(amount) FROM cash_transactions WHERE trans_type='قبض' AND category=? "
            "AND party_code IS NOT NULL AND substr(trans_date,1,7) BETWEEN ? AND ? GROUP BY 1", (CAT_PARTNER, k1, k2))}
    rows = []
    for k in keys:
        net = billed.get(k, 0) - rets.get(k, 0)
        rows.append((k, billed.get(k, 0), rets.get(k, 0), net, coll.get(k, 0), net - coll.get(k, 0), _pct(coll.get(k, 0), net)))
    cols = [("الشهر", "t"), ("فواتير المبيعات (شامل الضريبة)", "m"), ("المرتجعات", "m"), ("صافي الفوترة", "m"),
            ("المحصَّل", "m"), ("الفرق (زيادة الذمم)", "m"), ("نسبة التحصيل", "p")]
    t = [sum(r[i] for r in rows) for i in (1, 2, 3, 4, 5)]
    total = ("الإجمالي", t[0], t[1], t[2], t[3], t[4], _pct(t[3], t[2]))
    notes = ["نسبة التحصيل أقل من 100% تعني أن الذمم المدينة تزيد (السيولة تتأخر عن المبيعات)."]
    return Report("الإيرادات مقابل التحصيل", cols, rows, total, notes, f"من {k1} إلى {k2}", chart=(0, 3))


def rep_pur_supplier(p):
    a, b = p["d_from"], p["d_to"]
    with db() as conn:
        s = {c: (name, n, sub or 0, vat or 0) for c, name, n, sub, vat in conn.execute(
            "SELECT supplier_code, MAX(supplier_name), COUNT(*), SUM(subtotal), SUM(vat_amount) FROM purchase_invoices "
            "WHERE purchase_date BETWEEN ? AND ? GROUP BY supplier_code", (a, b))}
        rt = {c: [name, sub or 0] for c, name, sub in conn.execute(
            "SELECT party_code, MAX(party_name), SUM(subtotal) FROM returns_invoices WHERE return_type=? "
            "AND return_date BETWEEN ? AND ? GROUP BY party_code", (PURCHASE_RETURN, a, b))}
        for c, name, sub in conn.execute("SELECT party_code, MAX(party_name), SUM(subtotal) FROM post_discounts "
                                         "WHERE kind='purchase' AND discount_date BETWEEN ? AND ? GROUP BY party_code", (a, b)):
            rt.setdefault(c, [name, 0])[1] += sub or 0
    rows = []
    for c in set(s) | set(rt):
        name, n, sub, vat = s.get(c, (rt[c][0] if c in rt else c, 0, 0, 0))
        rsub = rt.get(c, (name, 0))[1]
        rows.append((c, name, n, sub, rsub, sub - rsub, vat))
    rows.sort(key=lambda r: r[5], reverse=True)
    rows = _share_rows(rows, 5)
    cols = [("الكود", "t"), ("المورد", "t"), ("عدد الفواتير", "i"), ("المشتريات (قبل الضريبة)", "m"),
            ("المرتجعات والخصومات اللاحقة", "m"),
            ("صافي المشتريات", "m"), ("الضريبة", "m"), ("النسبة", "p"), ("التراكمي", "p")]
    t = [sum(r[i] for r in rows) for i in (2, 3, 4, 5, 6)]
    total = ("", "الإجمالي", t[0], t[1], t[2], t[3], t[4], None, None)
    notes = [f"عدد الموردين المتعامل معهم: {len(rows)}"]
    if rows and t[3] > 0 and rows[0][7] > 40:
        notes.append(f"تنبيه تركّز: المورد «{rows[0][1]}» يمثل {rows[0][7]:.1f}% من المشتريات - فكّر في تنويع الموردين")
    return Report("المشتريات حسب المورد", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 5))


def rep_pur_item(p):
    a, b = p["d_from"], p["d_to"]
    agg = {}
    with db() as conn:
        for code, name, unit, d, q, cost, sub in conn.execute(
                "SELECT l.item_code, l.item_name, l.unit, h.purchase_date, l.quantity, l.unit_cost, l.subtotal "
                "FROM purchase_invoice_lines l JOIN purchase_invoices h ON h.purchase_id=l.purchase_id "
                "WHERE h.purchase_date BETWEEN ? AND ? ORDER BY h.purchase_date, h.purchase_id, l.line_id", (a, b)):
            r = agg.setdefault(code, dict(name=name, unit=unit or "", q=0.0, v=0.0, lo=cost, hi=cost, last=cost, ld=d))
            r["q"] += q
            r["v"] += sub
            r["lo"], r["hi"], r["last"], r["ld"] = min(r["lo"], cost), max(r["hi"], cost), cost, d
    rows = [(c, r["name"], r["unit"], r["q"], r["v"], (r["v"] / r["q"]) if r["q"] else 0.0, r["lo"], r["hi"],
             _pct(r["hi"] - r["lo"], r["lo"]), r["last"], r["ld"]) for c, r in agg.items()]
    rows.sort(key=lambda r: r[4], reverse=True)
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("الكمية", "q"), ("قيمة المشتريات", "m"), ("متوسط التكلفة", "m"),
            ("أقل تكلفة", "m"), ("أعلى تكلفة", "m"), ("تذبذب السعر", "p"), ("آخر تكلفة", "m"), ("تاريخ آخر شراء", "t")]
    total = ("", "الإجمالي", "", None, sum(r[4] for r in rows), None, None, None, None, None, "")
    notes = ["تذبذب السعر = (أعلى - أقل) ÷ أقل. القيم المرتفعة تستدعي التفاوض أو تغيير المورد."]
    return Report("المشتريات وتكلفة الأصناف", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 4))


def rep_open_orders(p):
    as_of = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    late = p["late_days"]
    rows = []
    with db() as conn:
        for kind, label in (("sales", "بيع"), ("purchase", "شراء")):
            c = FLOW[kind]
            for oid, d, party, item, q, done, price in conn.execute(
                    f"SELECT o.order_id, o.order_date, o.{c['pname']}, l.item_name, l.quantity, l.{c['done']}, {c['onet']} "
                    f"FROM {c['olines']} l JOIN {c['order']} o ON o.order_id=l.order_id "
                    f"WHERE o.status NOT IN ('مفوتر','مغلق','ملغي') AND l.quantity - l.{c['done']} > {EPS} "
                    f"AND o.order_date <= ?", (p["d_to"],)):
                rem = q - done
                age = (as_of - datetime.strptime(d, "%Y-%m-%d").date()).days
                rows.append((label, oid, d, party, item, q, done, rem, rem * price, age, "متأخر" if age > late else "في الموعد"))
    rows.sort(key=lambda r: r[9], reverse=True)
    cols = [("النوع", "t"), ("رقم الأمر", "i"), ("التاريخ", "t"), ("الطرف", "t"), ("الصنف", "t"), ("المطلوب", "q"),
            ("المنفّذ", "q"), ("المتبقي", "q"), ("قيمة المتبقي (قبل الضريبة)", "m"), ("العمر (يوم)", "i"), ("الحالة", "t")]
    sv = sum(r[8] for r in rows if r[0] == "بيع")
    pv = sum(r[8] for r in rows if r[0] == "شراء")
    notes = [f"أوامر بيع مفتوحة بقيمة {fmt(sv)} ر.س (مبيعات متوقعة)   |   أوامر شراء مفتوحة بقيمة {fmt(pv)} ر.س (التزامات قادمة)",
             f"عدد البنود المتأخرة (أكثر من {late} يوم): {sum(1 for r in rows if r[10] == 'متأخر')}"]
    return Report("الأوامر المفتوحة (بيع وشراء)", cols, rows, None, notes, f"حتى {p['d_to']}")


def rep_not_invoiced(p):
    as_of = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    rows = []
    with db() as conn:
        for kind, label in (("sales", "تسليم غير مفوتر"), ("purchase", "استلام غير مفوتر")):
            c = FLOW[kind]
            for fid, d, party, oid, n, val in conn.execute(
                    f"SELECT f.{c['fid']}, f.{c['fdate']}, f.{c['pname']}, f.order_id, COUNT(l.line_id), "
                    f"SUM({c['fval']}) FROM {c['ful']} f JOIN {c['flines']} l ON l.{c['fid']}=f.{c['fid']} "
                    f"WHERE f.invoice_id IS NULL AND f.{c['fdate']} <= ? GROUP BY f.{c['fid']}", (p["d_to"],)):
                age = (as_of - datetime.strptime(d, "%Y-%m-%d").date()).days
                rows.append((label, fid, d, party, oid, n, val or 0, (val or 0) * (1 + VAT_RATE), age))
    rows.sort(key=lambda r: r[8], reverse=True)
    cols = [("النوع", "t"), ("رقم المستند", "i"), ("التاريخ", "t"), ("الطرف", "t"), ("رقم الأمر", "i"), ("عدد البنود", "i"),
            ("القيمة (قبل الضريبة)", "m"), ("القيمة (شامل الضريبة)", "m"), ("العمر (يوم)", "i")]
    sv = sum(r[7] for r in rows if r[0].startswith("تسليم"))
    pv = sum(r[7] for r in rows if r[0].startswith("استلام"))
    notes = [f"بضاعة سُلّمت للعملاء ولم تُفوتر: {fmt(sv)} ر.س (إيراد يجب إصدار فاتورته)   |   "
             f"بضاعة استُلمت ولم تصل فاتورتها: {fmt(pv)} ر.س (التزام للموردين غير مسجل)"]
    return Report("تسليمات واستلامات غير مفوترة", cols, rows, None, notes, f"حتى {p['d_to']}")


def rep_stock_val(p):
    with db() as conn:
        data = conn.execute("SELECT item_code, item_name, unit, quantity, cost_price, quantity*cost_price FROM inventory "
                            "WHERE ABS(quantity) > ? ORDER BY quantity*cost_price DESC", (EPS,)).fetchall()
    rows = _share_rows([tuple(r) for r in data], 5)
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("الرصيد", "q"), ("متوسط التكلفة", "m"), ("القيمة", "m"),
            ("النسبة", "p"), ("التراكمي", "p")]
    tot = sum(r[5] for r in rows)
    total = ("", "الإجمالي", "", None, None, tot, None, None)
    notes = [f"عدد الأصناف بالمخزن: {len(rows)}   |   قيمة المخزون: {fmt(tot)} ر.س   (الأرصدة الحالية وقت التشغيل)"]
    if rows and tot > 0:
        n80 = next((i + 1 for i, r in enumerate(rows) if r[7] >= 80), len(rows))
        notes.append(f"{n80} صنف يمثلون 80% من قيمة المخزون (تحليل ABC) - هذه الأصناف تستحق أشد رقابة")
    neg = [r for r in rows if r[3] < 0]
    if neg:
        notes.append(f"تنبيه: {len(neg)} صنف برصيد سالب - راجع الحركات")
    return Report("تقييم المخزون", cols, rows, total, notes, "الرصيد الحالي", chart=(1, 5))


def rep_stock_move(p):
    a, b = p["d_from"], p["d_to"]

    def agg(conn, cols, table, date_col, extra=""):
        out = {}
        sql = (f"SELECT {cols[0]}, SUM(CASE WHEN {date_col} < ? THEN {cols[1]} ELSE 0 END), "
               f"SUM(CASE WHEN {date_col} BETWEEN ? AND ? THEN {cols[1]} ELSE 0 END) FROM {table} "
               f"WHERE {date_col} <= ? {extra} GROUP BY {cols[0]}")
        for code, before, during in conn.execute(sql, (a, a, b, b)):
            out[code] = (before or 0, during or 0)
        return out

    with db() as conn:
        items = conn.execute("SELECT item_code, item_name, unit, opening_qty FROM inventory ORDER BY item_name").fetchall()
        pin = agg(conn, ("item_code", "qty"), "v_stock_in", "mdate")
        pout = agg(conn, ("item_code", "qty"), "v_stock_out", "mdate")
        rs = agg(conn, ("item_code", "quantity"), "returns_invoices", "return_date", f"AND return_type='{SALES_RETURN}'")
        rp = agg(conn, ("item_code", "quantity"), "returns_invoices", "return_date", f"AND return_type='{PURCHASE_RETURN}'")
    rows = []
    for code, name, unit, opening in items:
        g = lambda d: d.get(code, (0, 0))
        before = (opening or 0) + g(pin)[0] + g(rs)[0] - g(pout)[0] - g(rp)[0]
        closing = before + g(pin)[1] + g(rs)[1] - g(pout)[1] - g(rp)[1]
        if any(abs(x) > EPS for x in (before, g(pin)[1], g(rs)[1], g(pout)[1], g(rp)[1], closing)):
            rows.append((code, name, unit or "", before, g(pin)[1], g(rs)[1], g(pout)[1], g(rp)[1], closing))
    rows.sort(key=lambda r: r[6], reverse=True)
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("رصيد أول المدة", "q"), ("مشتريات/استلام", "q"),
            ("مرتجع مبيعات", "q"), ("مبيعات/تسليم", "q"), ("مرتجع مشتريات", "q"), ("رصيد آخر المدة", "q")]
    total = ("", "الإجمالي", "", *[sum(r[i] for r in rows) for i in range(3, 9)])
    notes = ["مرتبة بحسب الكمية المباعة/المسلّمة خلال الفترة (الأسرع حركة أولاً)."]
    return Report("حركة المخزون (ملخص الفترة)", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 6))


STALE_DAYS_DEFAULT = 30  # الصنف راكد فقط إذا تجاوزت آخر حركة له 30 يوماً (أكثر من 30 وليس 30)


def last_item_movements(conn, as_of):
    """آخر حركة لكل صنف حتى as_of: {الكود: (التاريخ، نوع الحركة)}. تشمل الشراء والاستلام والبيع والتسليم والمرتجعات."""
    last = {}
    for sql, label in (
            ("SELECT item_code, MAX(mdate) FROM v_stock_in WHERE mdate <= ? GROUP BY item_code", "شراء"),
            ("SELECT item_code, MAX(mdate) FROM v_stock_out WHERE mdate <= ? GROUP BY item_code", "بيع"),
            ("SELECT item_code, MAX(return_date) FROM returns_invoices WHERE return_date <= ? GROUP BY item_code", "مرتجع")):
        for code, d in conn.execute(sql, (as_of,)):
            if d and (code not in last or d > last[code][0]):
                last[code] = (d, label)
    return last


def system_start_date(conn):
    """أقدم تاريخ مسجل في النظام (أول حركة مخزنية أو أول رصيد افتتاحي) - مرجع الأصناف ذات الرصيد الافتتاحي بلا حركة."""
    dates = []
    for sql in ("SELECT MIN(mdate) FROM v_stock_in", "SELECT MIN(mdate) FROM v_stock_out",
                "SELECT MIN(return_date) FROM returns_invoices",
                "SELECT MIN(opening_date) FROM business_partners WHERE COALESCE(opening_date,'') != ''"):
        try:
            d = conn.execute(sql).fetchone()[0]
        except sqlite3.Error:
            d = None
        if d:
            dates.append(d[:10])
    return min(dates) if dates else None


def stagnant_items(conn, as_of, days=STALE_DAYS_DEFAULT):
    """أصناف لها رصيد ولم تتحرك (لا شراء ولا بيع ولا مرتجع) لأكثر من days يوماً تماماً.
    الصنف المشترى حديثاً لا يُعتبر راكداً حتى لو لم يُبع بعد. الصنف بلا أي حركة (رصيد افتتاحي فقط)
    يُقاس من أول تاريخ مسجل في النظام، فلا يظهر كراكد في أول days يوماً من تشغيل النظام."""
    ref = datetime.strptime(as_of, "%Y-%m-%d").date()
    last = last_item_movements(conn, as_of)
    start = system_start_date(conn)
    out = []
    for code, name, unit, q, cost, val in conn.execute(
            "SELECT item_code, item_name, unit, quantity, cost_price, quantity*cost_price FROM inventory WHERE quantity > ?", (EPS,)):
        if code in last:
            ld, kind = last[code]
            shown = f"{ld} ({kind})"
        elif start:
            ld, shown = start, "رصيد افتتاحي بلا حركة"
        else:
            continue  # لا يوجد أي تاريخ مرجعي: لا نحكم بالركود (منعاً للتنبيه الخاطئ)
        try:
            since = (ref - datetime.strptime(ld[:10], "%Y-%m-%d").date()).days
        except ValueError:
            continue
        if since > days:
            out.append((code, name, unit or "", q, cost, val, shown, since))
    out.sort(key=lambda r: r[5], reverse=True)
    return out


def rep_dormant(p):
    days = p["stale_days"]
    with db() as conn:
        rows = stagnant_items(conn, p["d_to"], days)
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("الرصيد", "q"), ("متوسط التكلفة", "m"), ("القيمة المحجوزة", "m"),
            ("آخر حركة", "t"), ("أيام بلا حركة", "i")]
    tot = sum(r[5] for r in rows)
    total = ("", "الإجمالي", "", None, None, tot, "", None)
    notes = [f"أصناف راكدة (بلا أي حركة شراء أو بيع أو مرتجع لأكثر من {days} يوماً): {len(rows)} صنف برأس مال محبوس {fmt(tot)} ر.س.",
             "المجمدات تتأثر بتاريخ الصلاحية: فكّر في عروض تصريف أو إيقاف الشراء لهذه الأصناف."]
    return Report("الأصناف الراكدة", cols, rows, total, notes, f"حتى {p['d_to']}", chart=(1, 5))


def stock_batches(conn):
    """توزيع الرصيد الحالي لكل صنف على دفعات الشراء بافتراض صرف الأقدم أولاً (FIFO):
    الرصيد المتبقي ينتمي لأحدث الدفعات. ما يزيد عن الدفعات المسجلة = رصيد افتتاحي بلا تاريخ صلاحية."""
    items = {c: (n, u, q, cost) for c, n, u, q, cost in conn.execute(
        "SELECT item_code, item_name, unit, quantity, cost_price FROM inventory WHERE quantity > ?", (EPS,))}
    batches = defaultdict(list)
    for row in conn.execute("SELECT item_code, mdate, created_at, qty, production_date, expiry_date, ref, party "
                            "FROM v_batches ORDER BY mdate DESC, created_at DESC"):
        if row[0] in items:
            batches[row[0]].append(row)
    out = []
    for code, (name, unit, q, cost) in items.items():
        rem = q
        for _, mdate, _, bq, prod, exp, ref, party in batches[code]:
            if rem <= EPS:
                break
            take = min(rem, bq or 0.0)
            rem -= take
            if take > EPS:
                out.append(dict(code=code, name=name, unit=unit or "", ref=ref, party=party or "-", mdate=mdate,
                                production=prod or "", expiry=exp or "", qty=take, value=take * (cost or 0.0)))
        if rem > EPS:
            out.append(dict(code=code, name=name, unit=unit or "", ref="رصيد افتتاحي / غير محدد", party="-", mdate="",
                            production="", expiry="", qty=rem, value=rem * (cost or 0.0)))
    return out


def expiry_rows(conn, as_of, alert_days):
    """دفعات المخزون التي لها تاريخ انتهاء، مرتبة بالأقرب انتهاءً، مع الحالة."""
    rows = []
    for b in stock_batches(conn):
        if not b["expiry"]:
            continue
        try:
            days = (datetime.strptime(b["expiry"], "%Y-%m-%d").date() - as_of).days
        except ValueError:
            continue
        status = "منتهي الصلاحية" if days < 0 else ("قريب الانتهاء" if days <= alert_days else "سليم")
        rows.append(dict(b, days=days, status=status))
    rows.sort(key=lambda r: r["days"])
    return rows


def fetch_expiry_summary(alert_days=30, stale_days=STALE_DAYS_DEFAULT):
    with db() as conn:
        ex = expiry_rows(conn, date.today(), alert_days)
        st = stagnant_items(conn, today(), stale_days)
    expired = [r for r in ex if r["days"] < 0]
    near = [r for r in ex if 0 <= r["days"] <= alert_days]
    mid = [r for r in ex if alert_days < r["days"] <= max(90, alert_days)]
    val = lambda rows: sum(r["value"] for r in rows)
    return dict(rows=ex, expired=expired, near=near, mid=mid, stagnant=st, alert_days=alert_days, stale_days=stale_days,
                expired_value=val(expired), near_value=val(near), mid_value=val(mid),
                stagnant_value=sum(r[5] for r in st))


def rep_expiry(p):
    alert = int(p.get("expiry_days", 30))
    as_of = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    with db() as conn:
        data = expiry_rows(conn, as_of, alert)
    rows = [(r["code"], r["name"], r["unit"], r["ref"], r["party"], r["production"] or "-", r["expiry"], r["days"],
             r["qty"], r["value"], r["status"]) for r in data]
    cols = [("الكود", "t"), ("الصنف", "t"), ("الوحدة", "t"), ("الدفعة", "t"), ("المورد", "t"), ("تاريخ الإنتاج", "t"),
            ("تاريخ الانتهاء", "t"), ("الأيام المتبقية", "i"), ("الكمية المتبقية", "q"), ("القيمة بالتكلفة", "m"), ("الحالة", "t")]
    total = ("", "الإجمالي", "", "", "", "", "", None, None, sum(r[9] for r in rows), "")
    expired = [r for r in rows if r[7] < 0]
    near = [r for r in rows if 0 <= r[7] <= alert]
    notes = [f"منتهية الصلاحية: {len(expired)} دفعة بقيمة {fmt(sum(r[9] for r in expired))} ر.س   |   "
             f"تنتهي خلال {alert} يوماً: {len(near)} دفعة بقيمة {fmt(sum(r[9] for r in near))} ر.س",
             "الكمية المتبقية لكل دفعة محسوبة بافتراض صرف الأقدم أولاً (FIFO). الدفعات بدون تاريخ انتهاء لا تظهر هنا."]
    return Report("الصلاحية وقرب الانتهاء", cols, rows, total, notes, f"حتى {p['d_to']} • تنبيه {alert} يوماً", chart=(1, 9))


def rep_inactive_customers(p):
    ref = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    days = p["stale_days"]
    bal = {r["code"]: r["balance"] for r in partner_balances(p["d_to"])}
    with db() as conn:
        data = conn.execute("SELECT customer_code, MAX(customer_name), MAX(invoice_date), COUNT(*), SUM(subtotal) FROM sales_invoices "
                            "WHERE invoice_date <= ? GROUP BY customer_code", (p["d_to"],)).fetchall()
        phones = dict(conn.execute("SELECT partner_code, COALESCE(phone,'') FROM business_partners"))
    rows = []
    for code, name, last, n, total in data:
        since = (ref - datetime.strptime(last, "%Y-%m-%d").date()).days
        if since >= days:
            rows.append((code, name, phones.get(code, ""), last, since, n, total or 0, max(bal.get(code, 0), 0)))
    rows.sort(key=lambda r: r[6], reverse=True)
    cols = [("الكود", "t"), ("العميل", "t"), ("الهاتف", "t"), ("آخر فاتورة", "t"), ("أيام الانقطاع", "i"),
            ("عدد الفواتير السابقة", "i"), ("إجمالي مشترياته", "m"), ("رصيده المستحق", "m")]
    total = ("", "الإجمالي", "", "", None, sum(r[5] for r in rows), sum(r[6] for r in rows), sum(r[7] for r in rows))
    notes = [f"عملاء توقفوا عن الشراء منذ {days} يوماً أو أكثر: {len(rows)} عميل - مرتبون بقيمة مشترياتهم السابقة، تواصل مع الأكبر أولاً."]
    return Report("العملاء المنقطعون", cols, rows, total, notes, f"حتى {p['d_to']}")


def rep_returns(p):
    a, b = p["d_from"], p["d_to"]
    with db() as conn:
        rets = conn.execute("SELECT return_type, item_code, MAX(item_name), COUNT(*), SUM(quantity), SUM(total_amount) FROM returns_invoices "
                            "WHERE return_date BETWEEN ? AND ? GROUP BY return_type, item_code", (a, b)).fetchall()
        sold = dict(conn.execute("SELECT l.item_code, SUM(l.quantity) FROM sales_invoice_lines l JOIN sales_invoices h "
                                 "ON h.invoice_id=l.invoice_id WHERE h.invoice_date BETWEEN ? AND ? GROUP BY l.item_code", (a, b)))
        bought = dict(conn.execute("SELECT l.item_code, SUM(l.quantity) FROM purchase_invoice_lines l JOIN purchase_invoices h "
                                   "ON h.purchase_id=l.purchase_id WHERE h.purchase_date BETWEEN ? AND ? GROUP BY l.item_code", (a, b)))
    rows = []
    for typ, code, name, n, q, val in rets:
        base = (sold if typ == SALES_RETURN else bought).get(code, 0) or 0
        rows.append((typ, code, name, n, q, val or 0, base, _pct(q, base)))
    rows.sort(key=lambda r: r[5], reverse=True)
    cols = [("النوع", "t"), ("الكود", "t"), ("الصنف", "t"), ("عدد المرتجعات", "i"), ("الكمية المرتجعة", "q"),
            ("القيمة (شامل الضريبة)", "m"), ("الكمية المباعة/المشتراة", "q"), ("نسبة المرتجع", "p")]
    sv = sum(r[5] for r in rows if r[0] == SALES_RETURN)
    pv = sum(r[5] for r in rows if r[0] == PURCHASE_RETURN)
    notes = [f"مرتجعات المبيعات: {fmt(sv)} ر.س   |   مرتجعات المشتريات: {fmt(pv)} ر.س",
             "نسبة مرتجع مبيعات عالية لصنف = مشكلة جودة/تخزين/توصيل. نسبة مرتجع مشتريات عالية = مشكلة عند المورد."]
    return Report("المرتجعات حسب الصنف", cols, rows, None, notes, f"من {a} إلى {b}", chart=(2, 5))


def rep_cash_accounts(p):
    a, b = p["d_from"], p["d_to"]
    with db() as conn:
        data = cash_movements(conn, a, b)
        names = cash_account_names(conn)
    for acc in BANK_ACCOUNTS:
        data.setdefault(acc, [0.0, 0.0, 0.0])  # يظهر الحساب الأساسي حتى لو بدون حركة
    order = {n: i for i, n in enumerate(names)}
    rows = [(acc, o, i, x, o + i - x) for acc, (o, i, x) in sorted(data.items(), key=lambda kv: (order.get(kv[0], 999), kv[0]))]
    cols = [("الحساب", "t"), ("رصيد أول المدة", "m"), ("المقبوضات", "m"), ("المدفوعات", "m"), ("رصيد آخر المدة", "m")]
    total = ("الإجمالي", *[sum(r[i] for r in rows) for i in range(1, 5)])
    notes = ["الرصيد = سندات القبض والصرف + القيود اليومية على حسابات النقدية والبنوك (مثل قيد إثبات رأس المال)."]
    return Report("حركة الخزينة والبنوك", cols, rows, total, notes, f"من {a} إلى {b}", chart=(0, 4))


def rep_cash_flow(p):
    keys = month_keys(p["d_from"], p["d_to"])
    k1, k2 = keys[0], keys[-1]
    with db() as conn:
        opening = sum(cash_balances(conn, (date(int(k1[:4]), int(k1[5:7]), 1) - timedelta(days=1)).strftime("%Y-%m-%d")).values())
        data = defaultdict(float)
        for k, typ, cat, v in conn.execute(
                "SELECT substr(trans_date,1,7), trans_type, category, SUM(amount) FROM cash_transactions "
                "WHERE substr(trans_date,1,7) BETWEEN ? AND ? GROUP BY 1, 2, 3", (k1, k2)):
            data[(k, typ, cat)] += v or 0
        jv = journal_cash_by_month(conn, k1, k2)
    rows, bal = [], opening
    for k in keys:
        inc = sum(v for (m, typ, _), v in data.items() if m == k and typ == "قبض")
        pay, exp, ast = data[(k, "صرف", CAT_PARTNER)], data[(k, "صرف", CAT_EXPENSE)], data[(k, "صرف", CAT_ASSET)]
        emp = data[(k, "صرف", CAT_EMPLOYEE)]
        j = jv.get(k, 0.0)
        net = inc - pay - exp - ast - emp + j
        bal += net
        rows.append((k, inc, pay, exp, ast, emp, j, net, bal))
    cols = [("الشهر", "t"), ("المقبوضات", "m"), ("مدفوعات للموردين/الشركاء", "m"), ("مصروفات", "m"), ("شراء أصول", "m"),
            ("سلف وعهد العاملين", "m"), ("صافي القيود اليومية", "m"), ("صافي الحركة", "m"), ("الرصيد التراكمي", "m")]
    total = ("الإجمالي", *[sum(r[i] for r in rows) for i in range(1, 8)], rows[-1][8] if rows else opening)
    notes = [f"الرصيد قبل بداية الفترة: {fmt(opening)} ر.س (سندات + قيود)",
             "صافي القيود اليومية = حركات النقدية والبنوك المسجلة بقيود (رأس مال، تحويلات، تسويات...)."]
    neg = [r[0] for r in rows if r[7] < 0]
    if neg:
        notes.append("أشهر بتدفق نقدي سالب: " + "، ".join(neg))
    return Report("التدفق النقدي الشهري", cols, rows, total, notes, f"من {k1} إلى {k2}", chart=(0, 7))


def rep_expenses(p):
    a, b = p["d_from"], p["d_to"]
    months = max(len(month_keys(a, b)), 1)
    with db() as conn:
        data = expense_breakdown(conn, a, b)
    rows = _share_rows([(n, c, v, v / months) for n, c, v in data], 2)
    cols = [("بند المصروف", "t"), ("عدد الحركات", "i"), ("الإجمالي", "m"), ("المتوسط الشهري", "m"), ("النسبة", "p"), ("التراكمي", "p")]
    total = ("الإجمالي", sum(r[1] for r in rows), sum(r[2] for r in rows), sum(r[3] for r in rows), None, None)
    notes = ["يشمل سندات الصرف المصنفة «مصروفات عمومية» والقيود اليومية على حسابات المصروفات (عدا تكلفة المبيعات والإهلاك)."]
    return Report("تحليل المصروفات", cols, rows, total, notes, f"من {a} إلى {b}", chart=(0, 2))


def employee_accounts(conn):
    """[(الكود، الاسم)] لحسابات ذمم وعهد العاملين القابلة للترحيل (الفرعية فقط)."""
    like = " OR ".join("account_code LIKE ?" for _ in EMPLOYEE_PREFIXES)
    return conn.execute(f"SELECT account_code, account_name FROM chart_of_accounts WHERE ({like}) "
                        f"AND COALESCE(is_group,0)=0 ORDER BY account_code",
                        [f"{p}%" for p in EMPLOYEE_PREFIXES]).fetchall()


def rep_employees(p):
    """ذمم وعهد العاملين: رصيد كل حساب من سندات الخزينة + القيود اليومية. الموجب = العامل مدين للشركة."""
    a, b = p["d_from"], p["d_to"]
    data = defaultdict(lambda: [0.0, 0.0, 0.0])  # قبل الفترة، مدين، دائن
    with db() as conn:
        names = dict(conn.execute("SELECT account_code, account_name FROM chart_of_accounts"))
        for code, before, dr, cr in conn.execute(
                "SELECT party_code, SUM(CASE WHEN trans_date < ? THEN (CASE WHEN trans_type='صرف' THEN amount ELSE -amount END) ELSE 0 END), "
                "SUM(CASE WHEN trans_date >= ? AND trans_type='صرف' THEN amount ELSE 0 END), "
                "SUM(CASE WHEN trans_date >= ? AND trans_type='قبض' THEN amount ELSE 0 END) "
                "FROM cash_transactions WHERE category=? AND trans_date <= ? GROUP BY party_code",
                (a, a, a, CAT_EMPLOYEE, b)):
            r = data[code]
            r[0] += before or 0.0
            r[1] += dr or 0.0
            r[2] += cr or 0.0
        like = " OR ".join("l.account_code LIKE ?" for _ in EMPLOYEE_PREFIXES)
        for code, before, dr, cr in conn.execute(
                f"SELECT l.account_code, SUM(CASE WHEN e.entry_date < ? THEN l.debit - l.credit ELSE 0 END), "
                f"SUM(CASE WHEN e.entry_date >= ? THEN l.debit ELSE 0 END), SUM(CASE WHEN e.entry_date >= ? THEN l.credit ELSE 0 END) "
                f"FROM journal_lines l JOIN journal_entries e ON e.entry_id=l.entry_id "
                f"WHERE e.entry_date <= ? AND ({like}) GROUP BY l.account_code",
                [a, a, a, b] + [f"{x}%" for x in EMPLOYEE_PREFIXES]):
            r = data[code]
            r[0] += before or 0.0
            r[1] += dr or 0.0
            r[2] += cr or 0.0
    rows = []
    for code, (o, dr, cr) in sorted(data.items(), key=lambda kv: kv[0] or ""):
        close = o + dr - cr
        if any(abs(x) > 0.005 for x in (o, dr, cr, close)):
            side = "مدين (على العامل)" if close > 0.005 else ("دائن (للعامل)" if close < -0.005 else "مسدد")
            rows.append((code or "-", names.get(code, code or "-"), o, dr, cr, close, side))
    cols = [("الكود", "t"), ("الحساب / العامل", "t"), ("رصيد أول المدة", "m"), ("سلف وعهد مصروفة (مدين)", "m"),
            ("مسترد / مسوّى (دائن)", "m"), ("الرصيد", "m"), ("الحالة", "t")]
    total = ("", "الإجمالي", *[sum(r[i] for r in rows) for i in range(2, 6)], "")
    notes = ["المدين = سندات صرف (سلفة/عهدة) + قيود مدينة. الدائن = سندات قبض (رد السلفة) + قيود دائنة (تسوية العهدة بمصروفات أو خصم من الراتب).",
             "لتسوية عهدة بفواتير مصروفات: قيد يومية (من ح/ المصروف إلى ح/ عهدة العامل)."]
    return Report("ذمم وعهد العاملين", cols, rows, total, notes, f"من {a} إلى {b}", chart=(1, 5))


def rep_vat(p):
    keys = month_keys(p["d_from"], p["d_to"])
    k1, k2 = keys[0], keys[-1]
    with db() as conn:
        def month_sum(sql, params):
            return {k: v or 0 for k, v in conn.execute(sql, params)}
        so = month_sum("SELECT substr(invoice_date,1,7), SUM(vat_amount) FROM sales_invoices "
                       "WHERE substr(invoice_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2))
        sr = month_sum("SELECT substr(return_date,1,7), SUM(vat_amount) FROM returns_invoices WHERE return_type=? "
                       "AND substr(return_date,1,7) BETWEEN ? AND ? GROUP BY 1", (SALES_RETURN, k1, k2))
        pi = month_sum("SELECT substr(purchase_date,1,7), SUM(vat_amount) FROM purchase_invoices "
                       "WHERE substr(purchase_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2))
        pr = month_sum("SELECT substr(return_date,1,7), SUM(vat_amount) FROM returns_invoices WHERE return_type=? "
                       "AND substr(return_date,1,7) BETWEEN ? AND ? GROUP BY 1", (PURCHASE_RETURN, k1, k2))
        # الإشعارات الدائنة للعملاء تخفض ضريبة المخرجات، وخصومات الموردين اللاحقة تخفض ضريبة المدخلات
        for k, v in month_sum("SELECT substr(discount_date,1,7), SUM(vat_amount) FROM post_discounts WHERE kind='sales' "
                              "AND substr(discount_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2)).items():
            sr[k] = sr.get(k, 0) + v
        for k, v in month_sum("SELECT substr(discount_date,1,7), SUM(vat_amount) FROM post_discounts WHERE kind='purchase' "
                              "AND substr(discount_date,1,7) BETWEEN ? AND ? GROUP BY 1", (k1, k2)).items():
            pr[k] = pr.get(k, 0) + v
    rows = []
    for k in keys:
        out_net, in_net = so.get(k, 0) - sr.get(k, 0), pi.get(k, 0) - pr.get(k, 0)
        rows.append((k, so.get(k, 0), sr.get(k, 0), out_net, pi.get(k, 0), pr.get(k, 0), in_net, out_net - in_net))
    cols = [("الشهر", "t"), ("ضريبة المبيعات", "m"), ("(-) مرتجعات وإشعارات دائنة", "m"), ("ضريبة المخرجات", "m"),
            ("ضريبة المشتريات", "m"), ("(-) مرتجعات وخصومات لاحقة", "m"), ("ضريبة المدخلات", "m"), ("الصافي المستحق", "m")]
    total = ("الإجمالي", *[sum(r[i] for r in rows) for i in range(1, 8)])
    notes = ["الصافي الموجب = مستحق لهيئة الزكاة والضريبة والجمارك، والسالب = رصيد دائن يُرحَّل/يُسترد.",
             "لا يشمل ضريبة المصروفات العامة المدفوعة بسندات (غير مسجلة كفواتير)."]
    return Report("ضريبة القيمة المضافة (شهرياً)", cols, rows, total, notes, f"من {k1} إلى {k2}", chart=(0, 7))


def rep_assets(p):
    as_of = datetime.strptime(p["d_to"], "%Y-%m-%d").date()
    with db() as conn:
        data = conn.execute("SELECT asset_code, asset_name, purchase_date, purchase_cost, salvage_value, useful_life_years, "
                            "monthly_depreciation FROM fixed_assets ORDER BY asset_id").fetchall()
    rows = []
    for code, name, d, cost, salv, life, monthly in data:
        acc = accumulated_depreciation(cost, salv, life, d, as_of)
        rows.append((code, name, d, cost, salv or 0, life, monthly, acc, cost - acc))
    cols = [("الكود", "t"), ("الأصل", "t"), ("تاريخ الشراء", "t"), ("التكلفة", "m"), ("الخردة", "m"), ("العمر (سنة)", "i"),
            ("إهلاك شهري", "m"), ("مجمع الإهلاك", "m"), ("القيمة الدفترية", "m")]
    total = ("", "الإجمالي", "", sum(r[3] for r in rows), sum(r[4] for r in rows), None, sum(r[6] for r in rows),
             sum(r[7] for r in rows), sum(r[8] for r in rows))
    return Report("سجل الأصول الثابتة والإهلاك", cols, rows, total, [], f"حتى {p['d_to']}", chart=(1, 8))


def rep_kpi(p):
    a, b = p["d_from"], p["d_to"]
    to_dt, from_dt = datetime.strptime(b, "%Y-%m-%d").date(), datetime.strptime(a, "%Y-%m-%d").date()
    with db() as conn:
        def one(sql, params=()):
            return tuple((v or 0.0) for v in conn.execute(sql, params).fetchone())
        s_sub, s_n = one("SELECT SUM(subtotal), COUNT(*) FROM sales_invoices WHERE invoice_date BETWEEN ? AND ?", (a, b))
        r_sub, r_cost = one("SELECT SUM(subtotal), SUM(quantity*unit_cost) FROM returns_invoices WHERE return_type=? "
                            "AND return_date BETWEEN ? AND ?", (SALES_RETURN, a, b))
        (cogs,) = one("SELECT SUM(l.quantity*l.unit_cost) FROM sales_invoice_lines l JOIN sales_invoices h ON h.invoice_id=l.invoice_id "
                      "WHERE h.invoice_date BETWEEN ? AND ?", (a, b))
        (pur,) = one("SELECT SUM(subtotal) FROM purchase_invoices WHERE purchase_date BETWEEN ? AND ?", (a, b))
        (exp,) = one("SELECT SUM(amount) FROM cash_transactions WHERE trans_type='صرف' AND category=? AND trans_date BETWEEN ? AND ?",
                     (CAT_EXPENSE, a, b))
        j_exp, other_inc = journal_pl_totals(conn, a, b)
        exp += j_exp
        dsc = discount_totals(conn, a, b)[""]
        hw = hw_cogs(conn, a, b)[""]
        other_inc += dsc["earned"] - hw["credits"]
        cogs += hw["cogs"]
        assets = conn.execute("SELECT purchase_cost, salvage_value, useful_life_years, purchase_date FROM fixed_assets").fetchall()
        (stock,) = one("SELECT SUM(quantity*cost_price) FROM inventory")
        cash = sum(cash_balances(conn, b).values())
        (so_val,) = one("SELECT SUM((l.quantity - l.delivered_qty) * (l.unit_price - COALESCE(l.discount_amount,0) / l.quantity)) FROM sales_order_lines l JOIN sales_orders o "
                        "ON o.order_id=l.order_id WHERE o.status NOT IN ('مفوتر','مغلق','ملغي')")
        (po_val,) = one("SELECT SUM((l.quantity - l.received_qty) * l.unit_cost) FROM purchase_order_lines l JOIN purchase_orders o "
                        "ON o.order_id=l.order_id WHERE o.status NOT IN ('مفوتر','مغلق','ملغي')")
        (undeliv,) = one("SELECT SUM(l.quantity*l.unit_price - COALESCE(l.discount_amount,0)) FROM sales_deliveries d JOIN sales_delivery_lines l ON l.delivery_id=d.delivery_id "
                         "WHERE d.invoice_id IS NULL")
    dep = sum(accumulated_depreciation(c, s, l, d, to_dt) - accumulated_depreciation(c, s, l, d, from_dt - timedelta(days=1))
              for c, s, l, d in assets)
    bals = partner_balances(b)
    recv = sum(r["balance"] for r in bals if r["balance"] > 0.005)
    payb = -sum(r["balance"] for r in bals if r["balance"] < -0.005)
    net_sales = s_sub - r_sub - dsc["sales_sub"]
    gross = net_sales - (cogs - r_cost)
    profit = gross + other_inc - exp - dep
    days = (to_dt - from_dt).days + 1
    dso = (recv / (net_sales * (1 + VAT_RATE)) * days) if net_sales > 0 else 0.0
    R = "ر.س"
    rows = [
        ("الأداء", "صافي المبيعات (قبل الضريبة)", net_sales, R),
        ("الأداء", "مجمل الربح", gross, R),
        ("الأداء", "هامش مجمل الربح", _pct(gross, net_sales), "%"),
        ("الأداء", "إيرادات أخرى (قيود + خصم مكتسب)", other_inc, R),
        ("الأداء", "خصومات لاحقة ممنوحة للعملاء (إشعارات دائنة)", dsc["sales_sub"], R),
        ("الأداء", "خصومات لاحقة مستلمة من الموردين", dsc["pur_sub"], R),
        ("الأداء", "المصروفات التشغيلية", exp, R),
        ("الأداء", "مصروف الإهلاك", dep, R),
        ("الأداء", "صافي الربح", profit, R),
        ("الأداء", "هامش صافي الربح", _pct(profit, net_sales), "%"),
        ("الأداء", "عدد فواتير المبيعات", s_n, "فاتورة"),
        ("الأداء", "متوسط قيمة الفاتورة", s_sub / s_n if s_n else 0.0, R),
        ("الأداء", "مشتريات الفترة (قبل الضريبة)", pur, R),
        ("السيولة والذمم", "الذمم المدينة / أرصدة العملاء", recv, R),
        ("السيولة والذمم", "الذمم الدائنة / أرصدة الموردين", payb, R),
        ("السيولة والذمم", "رصيد الخزينة والبنوك (سندات + قيود)", cash, R),
        ("السيولة والذمم", "صافي المركز (خزينة + ذمم مدينة - ذمم دائنة)", cash + recv - payb, R),
        ("السيولة والذمم", "فترة التحصيل التقريبية (DSO)", dso, "يوم"),
        ("المخزون", "قيمة المخزون الحالية", stock, R),
        ("المخزون", "بضاعة مسلّمة لم تُفوتر", undeliv, R),
        ("الأوامر", "أوامر بيع مفتوحة (متبقي)", so_val, R),
        ("الأوامر", "أوامر شراء مفتوحة (متبقي)", po_val, R),
    ]
    notes = ["DSO = متوسط عدد الأيام لتحصيل المبيعات (ذمم العملاء ÷ المبيعات شاملة الضريبة × أيام الفترة).",
             "الأرقام التفصيلية في بقية التقارير. استخدم الفترة السريعة لمقارنة الشهر الحالي بالسابق."]
    return Report("الملخص التنفيذي (مؤشرات القرار)", [("القسم", "t"), ("المؤشر", "t"), ("القيمة", "m"), ("الوحدة", "t")],
                  rows, None, notes, f"من {a} إلى {b}")


# (المجموعة، [(المفتاح، العنوان، الدالة، الوصف)])
SALES_MGR_COLUMNS = [
    ("Salesman Name", "t"), ("CUSTOMER CODE", "t"), ("CUSTOMER NAME", "t"), ("DISTRIBUTOR SKU CODE", "t"),
    ("PRODUCT DESCRIPTION", "t"), ("SALES TYPE", "t"), ("INVOICE DATE", "t"), ("QNT", "q"), ("Price", "m"),
    ("Sales Amount", "m"), ("INVOICE_NO", "t"), ("Net Cost", "m"), ("Profit", "m"), ("GP", "m"), ("GP%", "p"),
    ("GPM%", "p"), ("COGS", "m"),
]
SALES_TYPE_SALE, SALES_TYPE_RETURN, SALES_TYPE_CREDIT = "Sales", "Return", "Credit Note"


def sales_manager_rows(d_from, d_to, rep=None, customer=None):
    """سطور تقرير مدير المبيعات بالترتيب المطلوب حرفياً، من الحركات الفعلية:
    - بنود فواتير المبيعات (Sales)، والمرتجعات (Return: كمية وقيمة وتكلفة بالسالب)،
      والإشعارات الدائنة للعملاء (Credit Note: قيمة بالسالب بدون كمية ولا تكلفة) حتى يطابق الصافي قائمة الدخل.
    - Price = صافي سعر الوحدة بعد الخصم الخاص، فيتحقق دائماً: QNT × Price = Sales Amount (قبل الضريبة).
    - Net Cost = تكلفة الوحدة (متوسط مرجح وقت الصرف)، COGS = QNT × Net Cost.
    - Profit = ربح الوحدة (Price - Net Cost)، GP = Sales Amount - COGS.
    - GP% = GP ÷ COGS (نسبة الربح على التكلفة)، GPM% = GP ÷ Sales Amount (هامش الربح من المبيعات).
    - المندوب حسب فترة إسناده الفعلية للعميل بتاريخ الحركة (نفس منطق تقرير أداء المناديب)."""
    out = []
    with db() as conn:
        who = RepResolver(conn)
        lines = conn.execute(
            "SELECT h.invoice_date, h.customer_code, h.customer_name, l.item_code, l.item_name, l.quantity, l.subtotal, "
            "l.unit_cost, COALESCE(h.invoice_no, h.invoice_id), h.sales_rep, h.invoice_id, l.line_id "
            "FROM sales_invoice_lines l JOIN sales_invoices h ON h.invoice_id=l.invoice_id "
            "WHERE h.invoice_date BETWEEN ? AND ?" + (" AND h.customer_code=?" if customer else "") +
            " ORDER BY h.invoice_date, h.invoice_id, l.line_id", (d_from, d_to) + ((customer,) if customer else ())).fetchall()
        for d, code, name, sku, desc, qty, sub, ucost, no, doc_rep, _iid, _lid in lines:
            out.append(dict(rep=who(code, d, doc_rep), code=code, name=name, sku=sku, desc=desc, type=SALES_TYPE_SALE,
                            date=d, qty=qty or 0.0, amount=sub or 0.0, no=str(no), ucost=ucost or 0.0))
        rets = conn.execute(
            "SELECT return_date, party_code, party_name, item_code, item_name, quantity, subtotal, unit_cost, return_id "
            "FROM returns_invoices WHERE return_type=? AND return_date BETWEEN ? AND ?" +
            (" AND party_code=?" if customer else "") + " ORDER BY return_date, return_id",
            (SALES_RETURN, d_from, d_to) + ((customer,) if customer else ())).fetchall()
        for d, code, name, sku, desc, qty, sub, ucost, rid in rets:
            out.append(dict(rep=who(code, d), code=code, name=name, sku=sku, desc=desc, type=SALES_TYPE_RETURN, date=d,
                            qty=-(qty or 0.0), amount=-(sub or 0.0), no=f"RET-{rid}", ucost=ucost or 0.0))
        cns = conn.execute(
            "SELECT d.discount_date, d.party_code, d.party_name, d.subtotal, COALESCE(d.discount_no, d.discount_id), "
            "COALESCE(i.invoice_no, CAST(d.invoice_id AS TEXT), ''), COALESCE(d.reason,'') FROM post_discounts d "
            "LEFT JOIN sales_invoices i ON i.invoice_id=d.invoice_id WHERE d.kind='sales' AND d.discount_date BETWEEN ? AND ?" +
            (" AND d.party_code=?" if customer else "") + " ORDER BY d.discount_date, d.discount_id",
            (d_from, d_to) + ((customer,) if customer else ())).fetchall()
        for d, code, name, sub, no, inv_no, reason in cns:
            desc = "إشعار دائن / خصم لاحق" + (f" على فاتورة {inv_no}" if inv_no else "") + (f" - {reason}" if reason else "")
            out.append(dict(rep=who(code, d), code=code, name=name, sku="", desc=desc, type=SALES_TYPE_CREDIT, date=d,
                            qty=0.0, amount=-(sub or 0.0), no=str(no), ucost=0.0))
    if rep:
        out = [r for r in out if r["rep"] == rep]
    out.sort(key=lambda r: (r["rep"], r["name"], r["date"], r["no"]))
    return out


def rep_sales_manager(p):
    a, b = p["d_from"], p["d_to"]
    data = sales_manager_rows(a, b, p.get("rep") or None, p.get("customer") or None)
    rows = []
    for r in data:
        qty, amount = r["qty"], round(r["amount"], 2)
        price = (amount / qty) if abs(qty) > EPS else None
        cogs = round(qty * r["ucost"], 2)
        gp = round(amount - cogs, 2)
        unit_profit = (price - r["ucost"]) if price is not None else None
        rows.append((r["rep"], r["code"], r["name"], r["sku"], r["desc"], r["type"], r["date"], qty if abs(qty) > EPS else None,
                     price, amount, r["no"], r["ucost"] if abs(qty) > EPS else None, unit_profit, gp,
                     _pct(gp, cogs) if abs(cogs) > EPS else None, _pct(gp, amount) if abs(amount) > EPS else None, cogs))
    t_qty = sum(r[7] or 0 for r in rows)
    t_amount = sum(r[9] for r in rows)
    t_cogs = sum(r[16] for r in rows)
    t_gp = t_amount - t_cogs
    total = ("الإجمالي", "", f"{len({r[1] for r in rows})} عميل", "", "", "", "", t_qty, None, t_amount, f"{len(rows)} سطر",
             None, None, t_gp, _pct(t_gp, t_cogs) if t_cogs else None, _pct(t_gp, t_amount) if t_amount else None, t_cogs)
    flt = []
    if p.get("rep"):
        flt.append(f"المندوب: {p['rep']}")
    if p.get("customer"):
        flt.append(f"العميل: {p['customer']}")
    notes = ["Sales Amount = QNT × Price (صافي قبل الضريبة وبعد الخصم الخاص). COGS = QNT × Net Cost (متوسط التكلفة وقت الصرف).",
             "Profit = ربح الوحدة (Price − Net Cost)، GP = Sales Amount − COGS، GP% = GP ÷ COGS، GPM% = GP ÷ Sales Amount.",
             "المرتجعات (Return) بالسالب، والإشعارات الدائنة (Credit Note) تخفض المبيعات بلا كمية ولا تكلفة؛ "
             "فيطابق الإجمالي صافي المبيعات ومجمل الربح في قائمة الدخل للفترة.",
             "المندوب حسب فترة إسناده الفعلية للعميل بتاريخ كل حركة."]
    return Report("تقرير مدير المبيعات | Sales Manager Report", SALES_MGR_COLUMNS, rows, total, notes,
                  f"من {a} إلى {b}" + (" | " + " | ".join(flt) if flt else ""), chart=None,
                  highlight=(9, 13), ltr=True)


REPORT_CATALOG = [
    ("📋 تقرير مدير المبيعات", [
        ("sales_mgr", "تقرير مدير المبيعات (Sales Manager Report)", rep_sales_manager,
         "سطر لكل صنف في كل فاتورة/مرتجع/إشعار دائن: المندوب، العميل، الصنف، الكمية، السعر، المبيعات، التكلفة، الربح، "
         "GP وGP% وGPM%. فلترة بالفترة والمندوب والعميل."),
    ]),
    ("🎯 الملخص التنفيذي", [
        ("kpi", "مؤشرات القرار", rep_kpi, "ملخص الأداء والسيولة والذمم والمخزون والأوامر في صفحة واحدة. يستخدم الفترة."),
        ("dash_src", "مصادر أرقام لوحة المؤشرات", lambda p: rep_dash_sources(p),
         "كل رقم في لوحة المؤشرات: قيمته ومصدره ومعادلته (تراكمي من بداية العمل حتى اليوم)."),
    ]),
    ("💳 الذمم والمديونيات", [
        ("rec", "الذمم المدينة / أرصدة العملاء", rep_receivables, "كل عميل مدين لك بحسب الرصيد حتى تاريخ (إلى). يخصم المقبوضات والمرتجعات."),
        ("pay", "الذمم الدائنة / أرصدة الموردين", rep_payables, "كل مورد تدين له بحسب الرصيد حتى تاريخ (إلى)."),
        ("age_c", "أعمار ديون العملاء", lambda p: rep_aging("customer", p), "تحليل الديون بحسب عمر الفاتورة (0-30 / 31-60 / 61-90 / +90)."),
        ("age_s", "أعمار ديون الموردين", lambda p: rep_aging("supplier", p), "الفواتير المستحقة للموردين بحسب عمرها."),
        ("pt_cust", "تقرير العملاء (إجماليات / تفاصيل مجمّعة)", rep_customers_report,
         "إجماليات: سطر لكل عميل. تفاصيل مجمّعة: كل عميل بكل حركاته ثم إجمالي فرعي، ثم إجمالي عام. يستخدم الفترة و«العرض»."),
        ("pt_supp", "تقرير الموردين (إجماليات / تفاصيل مجمّعة)", rep_suppliers_report,
         "إجماليات: سطر لكل مورد. تفاصيل مجمّعة: كل مورد بكل حركاته ثم إجمالي فرعي، ثم إجمالي عام. يستخدم الفترة و«العرض»."),
        ("pt_net", "العملاء والموردون معاً (المقاصة)", rep_partner_netting,
         "رصيد الطرف كعميل مقابل رصيده كمورد وصافي المقاصة، مع إبراز الأطراف المشتركة (نفس الكود أو الاسم أو الرقم الضريبي)."),
    ]),
    ("📈 المبيعات والإيرادات", [
        ("s_period", "المبيعات حسب الفترة", rep_sales_period, "شهري أو يومي، مع المرتجعات والتكلفة والهامش والنمو. يستخدم التجميع."),
        ("s_cust", "المبيعات حسب العميل", rep_sales_customer, "ترتيب العملاء بصافي المبيعات والربح والنسبة التراكمية (باريتو)."),
        ("s_item", "المبيعات والربحية حسب الصنف", rep_sales_item, "الكمية والمبيعات والتكلفة والربح لكل صنف."),
        ("s_rep", "أداء المناديب (مبيعات، مشتريات، تحصيلات)", rep_rep_performance,
         "لكل مندوب: صافي المبيعات والمشتريات والتحصيلات حسب فترة تواجده الفعلية مع كل عميل/مورد."),
        ("pnl", "قائمة الدخل الشهرية", rep_monthly_pnl, "المبيعات والتكلفة والمصروفات التشغيلية ومصروف الإهلاك (كل منهما مستقل) وصافي الربح لكل شهر."),
        ("coll", "الإيرادات مقابل التحصيل", rep_collection, "الفوترة الشهرية مقابل المحصَّل فعلاً ونسبة التحصيل."),
        ("inactive", "العملاء المنقطعون", rep_inactive_customers, "عملاء توقفوا عن الشراء. يستخدم «أيام الركود»."),
    ]),
    ("🛒 المشتريات", [
        ("p_supp", "المشتريات حسب المورد", rep_pur_supplier, "صافي المشتريات لكل مورد وتركّز الاعتماد عليه."),
        ("p_item", "المشتريات وتكلفة الأصناف", rep_pur_item, "متوسط وأقل وأعلى وآخر تكلفة لكل صنف وتذبذب السعر."),
    ]),
    ("📝 دورة الأوامر", [
        ("open_ord", "الأوامر المفتوحة (بيع وشراء)", rep_open_orders, "ما لم يُنفَّذ بعد. يستخدم «مهلة التأخر»."),
        ("unbilled", "تسليمات واستلامات غير مفوترة", rep_not_invoiced, "بضاعة تحركت ولم تُصدر لها فاتورة."),
    ]),
    ("🏭 المخزون", [
        ("stock", "تقييم المخزون (ABC)", rep_stock_val, "قيمة كل صنف ونسبته وتراكمي النسب."),
        ("move", "حركة المخزون (ملخص الفترة)", rep_stock_move, "رصيد أول وآخر المدة والوارد والصادر لكل صنف."),
        ("returns", "المرتجعات حسب الصنف", rep_returns, "نسبة المرتجع من المباع/المشترى لكل صنف."),
    ]),
    ("⏳ الصلاحية والركود", [
        ("expiry", "الصلاحية وقرب الانتهاء", rep_expiry, "دفعات المخزون بتاريخ الانتهاء والأيام المتبقية. يستخدم «تنبيه الصلاحية»."),
        ("dormant", "الأصناف الراكدة", rep_dormant, "أصناف لها رصيد وبلا أي حركة (شراء/بيع/مرتجع) لأكثر من «أيام الركود» (الافتراضي 30)."),
    ]),
    ("💰 الخزينة والمصروفات", [
        ("cash_acc", "حركة الخزينة والبنوك", rep_cash_accounts, "رصيد أول وآخر المدة والمقبوضات والمدفوعات لكل حساب."),
        ("cash_flow", "التدفق النقدي الشهري", rep_cash_flow, "تفصيل المقبوضات والمدفوعات والمصروفات لكل شهر."),
        ("exp", "تحليل المصروفات", rep_expenses, "المصروفات بحسب البند ونسبة كل بند (سندات + قيود)."),
        ("emp", "ذمم وعهد العاملين", rep_employees, "رصيد كل عامل من السلف والعهد المصروفة والمستردة والمسوّاة بقيود."),
    ]),
    ("🧾 الضريبة والأصول والقيود", [
        ("vat", "ضريبة القيمة المضافة (شهرياً)", rep_vat, "ضريبة المخرجات والمدخلات والصافي المستحق لكل شهر."),
        ("assets", "سجل الأصول الثابتة", rep_assets, "التكلفة ومجمع الإهلاك والقيمة الدفترية حتى تاريخ (إلى)."),
        ("tb", "ميزان المراجعة (شجرة الحسابات)", rep_trial_balance,
         "ميزان الدفتر العام الكامل (كل المستندات مرحّلة آلياً + القيود اليومية) مع تجميع الحسابات الفرعية تحت الرئيسية."),
        ("acct_tree", "أرصدة الحسابات الفرعية (تحليلي)", rep_account_tree,
         "المسحوب والمسدد والمتبقي لكل حساب فرعي (بنوك، خزائن، سلف وعهد العاملين) من القيود والسندات، مجمعة تحت الرئيسية."),
        ("pdisc", "الخصومات اللاحقة والإشعارات الدائنة", rep_post_discounts,
         "كل خصم لاحق من الموردين وإشعار دائن للعملاء وأثره على المخزون والأرباح والضريبة."),
    ]),
]
REPORT_INDEX = {key: (title, func, desc) for _, items in REPORT_CATALOG for key, title, func, desc in items}

# ============================================================
# الصلاحيات التفصيلية داخل المديولات (مفتاح فرعي = «الصفحة.الإجراء»)
# ============================================================
REPORT_GROUP_KEYS = ["salesmgr", "exec", "debts", "sales", "purch", "orders", "stock", "expiry", "cash", "tax"]
assert len(REPORT_GROUP_KEYS) == len(REPORT_CATALOG)
REPORT_GROUP_OF = {key: f"reports.{g}" for g, (_, items) in zip(REPORT_GROUP_KEYS, REPORT_CATALOG) for key, *_ in items}
SUB_PERMISSIONS = {
    "sales": [("sales.orders", "أوامر البيع (عرض)"), ("sales.orders_create", "أوامر البيع: إنشاء وترحيل"),
              ("sales.delivery", "أذون التسليم (عرض)"), ("sales.delivery_create", "أذون التسليم: تنفيذ وخصم المخزن"),
              ("sales.invoices", "الفواتير الضريبية (السجل والإصدار من الأذون)"),
              ("sales.discount", "الخصم اللاحق / الإشعارات الدائنة للعملاء"),
              ("sales.customer_fin", "البيانات المالية للعملاء (الأرصدة والكشوف والأعمار)")],
    "purchase": [("purchase.orders", "أوامر الشراء (عرض)"), ("purchase.orders_create", "أوامر الشراء: إنشاء وترحيل"),
                 ("purchase.receipt", "استلام المخزن (عرض)"),
                 ("purchase.receipt_create", "استلام المخزن: تنفيذ واعتماد + إرفاق فاتورة المورد"),
                 ("purchase.invoices", "فواتير الموردين (السجل والتسجيل من الأذون)"),
                 ("purchase.discount", "الخصم اللاحق من الموردين (إشعارات الخصم)"),
                 ("purchase.supplier_fin", "البيانات المالية للموردين (الأرصدة والكشوف والأعمار)")],
    "cash": [("cash.receipt", "سندات القبض"), ("cash.payment", "سندات الصرف"),
             ("cash.register", "سجل السندات وأرصدة الخزينة والبنوك"), ("cash.print", "طباعة ملخص اليوم PDF")],
    "reports": [(f"reports.{g}", title.split(" ", 1)[-1]) for g, (title, _) in zip(REPORT_GROUP_KEYS, REPORT_CATALOG)],
}
SUB_PARENT = {k: parent for parent, subs in SUB_PERMISSIONS.items() for k, _ in subs}
SUB_LABELS = {k: label for subs in SUB_PERMISSIONS.values() for k, label in subs}
ALL_PERMISSION_KEYS = PERMISSION_KEYS + list(SUB_PARENT)
# تقارير تكشف أرصدة الشركاء: تحتاج صلاحية البيانات المالية للطرف المعني
REPORT_FIN_REQUIRES = {"rec": "sales.customer_fin", "age_c": "sales.customer_fin",
                       "pay": "purchase.supplier_fin", "age_s": "purchase.supplier_fin",
                       "pt_cust": "sales.customer_fin", "pt_supp": "purchase.supplier_fin",
                       "pt_net": ("sales.customer_fin", "purchase.supplier_fin")}   # tuple = يلزم الكل
# تقارير تقبل خيار «العرض» (إجماليات / تفاصيل مجمّعة) وفلتري «غير الصفرية» و«المشتركة»
PARTY_REPORTS = ("pt_cust", "pt_supp", "pt_net")


def report_fin_needs(key):
    """صلاحيات البيانات المالية المطلوبة للتقرير كأدوار: [] أو ["customer"] أو ["customer", "supplier"]."""
    need = REPORT_FIN_REQUIRES.get(key)
    needs = need if isinstance(need, tuple) else ((need,) if need else ())
    return ["customer" if n.startswith("sales") else "supplier" for n in needs]


def register_report_group(group_key, title, items):
    """إضافة مجموعة تقارير من وحدة أخرى (الدفتر العام، التخطيط المالي...) مع صلاحيتها الفرعية «reports.<group>».
    items = [(key, title, func, desc)]. التعديل في نفس القوائم (in place) حتى تراها كل الواجهات."""
    if group_key in REPORT_GROUP_KEYS:
        return
    REPORT_CATALOG.append((title, items))
    REPORT_GROUP_KEYS.append(group_key)
    perm = f"reports.{group_key}"
    for key, t, func, desc in items:
        REPORT_INDEX[key] = (t, func, desc)
        REPORT_GROUP_OF[key] = perm
    label = title.split(" ", 1)[-1]
    SUB_PERMISSIONS["reports"].append((perm, label))
    SUB_PARENT[perm] = "reports"
    SUB_LABELS[perm] = label
    if perm not in ALL_PERMISSION_KEYS:
        ALL_PERMISSION_KEYS.append(perm)


def user_can(user, key):
    """هل يملك المستخدم الصلاحية؟ المفاتيح الفرعية «module.action» تتبع صلاحية الصفحة الأم:
    - بدون الصفحة الأم = ممنوع.
    - الصفحة الأم بلا أي مفتاح فرعي مخزّن (مستخدم قديم) = كل الإجراءات داخلها مسموحة.
    - غير ذلك = المسموح هو المفاتيح الفرعية المختارة فقط."""
    if user["is_admin"]:
        return True
    perms = user["perms"]
    if key == "users":
        return False
    parent = SUB_PARENT.get(key)
    if parent is None:
        return key in perms
    if parent not in perms:
        return False
    if not any(s in perms for s, _ in SUB_PERMISSIONS[parent]):
        return True
    return key in perms


def user_can_fin(user, role):
    """رؤية أرصدة العملاء/الموردين. لا تُقيَّد إلا إذا أُسندت للمستخدم صفحة المبيعات/المشتريات بصلاحيات تفصيلية
    لا تتضمن البيانات المالية (صفحات الكشوف والتقارير لها صلاحياتها المستقلة)."""
    key = "sales.customer_fin" if role == "customer" else "purchase.supplier_fin"
    if user["is_admin"]:
        return True
    parent = SUB_PARENT[key]
    if parent not in user["perms"]:
        return True
    return user_can(user, key)


def run_report(key, params):
    return REPORT_INDEX[key][1](params)


# ---------------------------------------------------------- التصدير
def export_report_xlsx(path, rep):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    with db() as conn:
        company = conn.execute("SELECT company_name FROM company_profile WHERE id=1").fetchone()[0]
    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r'[\[\]:*?/\\]', '', rep.title)[:31] or "تقرير"
    ws.sheet_view.rightToLeft = not getattr(rep, "ltr", False)
    ncols = len(rep.columns)
    hl = {i + 1 for i in getattr(rep, "highlight", ())}
    for text, size in ((company, 14), (rep.title, 13), (rep.subtitle, 10)):
        ws.append([text])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=size)
    ws.append([])
    ws.append([h for h, _ in rep.columns])
    head_row = ws.max_row   # بعد الإضافة: سطر فارغ لا يزيد max_row فكان التنسيق يقع على سطر خاطئ
    for c in range(1, ncols + 1):
        cell = ws.cell(row=head_row, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E293B")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if c in hl:   # عمود مميز (مثل Sales Amount و GP): عنوان أصفر داكن
            cell.fill = PatternFill("solid", fgColor="FACC15")
            cell.font = Font(bold=True, color="111827")
    formats = {"m": "#,##0.00", "q": "#,##0.###", "p": '0.0"%"', "i": "#,##0"}

    def put(row, bold=False, fill=None):
        ws.append(list(row))
        r = ws.max_row
        for c, (_, kind) in enumerate(rep.columns, start=1):
            cell = ws.cell(row=r, column=c)
            if kind in formats and not isinstance(cell.value, str):
                cell.number_format = formats[kind]
            if bold:
                cell.font = Font(bold=True)
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
            elif c in hl:
                cell.fill = PatternFill("solid", fgColor="FEF08A")

    fills = {"group": "E0E7FF", "sub": "F1F5F9", "shared": "FCE7F3", "net": "DCFCE7"}
    styles = getattr(rep, "row_styles", {}) or {}
    for i, row in enumerate(rep.rows):
        stl = styles.get(i)
        put(row, bold=bool(stl), fill=fills.get(stl))
    if rep.total:
        put(rep.total, bold=True, fill="FDE68A")
    ws.append([])
    for n in rep.notes:
        ws.append([n])
    for c in range(1, ncols + 1):
        longest = max([len(str(rep.columns[c - 1][0]))] + [len(fmt_cell(r[c - 1], rep.columns[c - 1][1])) for r in rep.rows[:200]])
        ws.column_dimensions[get_column_letter(c)].width = min(max(longest + 3, 11), 45)
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    wb.save(path)


def build_report_pdf(path, rep):
    S = pdf_styles()
    with db() as conn:
        company = conn.execute("SELECT company_name FROM company_profile WHERE id=1").fetchone()[0]
    ncols = len(rep.columns)
    ltr = getattr(rep, "ltr", False)
    order = (lambda seq: list(seq)) if ltr else rtl
    page = landscape(A3) if ncols >= 14 else (landscape(A4) if ncols > 5 else A4)
    if ncols >= 14:   # تقرير عريض: خط أصغر حتى تظهر كل الأعمدة بدون قص
        S = dict(S, cell=ParagraphStyle("cell_s", parent=S["cell"], fontSize=7, leading=9),
                 head=ParagraphStyle("head_s", parent=S["head"], fontSize=7, leading=9))
    margin = 20
    doc = SimpleDocTemplate(path, pagesize=page, rightMargin=margin, leftMargin=margin, topMargin=28, bottomMargin=28)
    avail = page[0] - 2 * margin
    weight = {"t": 2.6, "m": 1.6, "q": 1.3, "p": 1.35, "i": 1.0}
    ws = [weight[k] for _, k in rep.columns]
    widths = [avail * w / sum(ws) for w in ws]
    el = [P(company, S["title"]), Spacer(1, 4), P(rep.title, S["title"]), Spacer(1, 3), P(rep.subtitle, S["sub"]),
          P(f"تاريخ الطباعة: {today()}", S["sub"]), Spacer(1, 8)]
    data = [order([P(h, S["head"]) for h, _ in rep.columns])]
    for row in rep.rows:
        data.append(order([P(fmt_cell(v, k), S["cell"]) for v, (_, k) in zip(row, rep.columns)]))
    last = len(data) - 1
    if rep.total:
        data.append(order([P(fmt_cell(v, k), S["cell"]) for v, (_, k) in zip(rep.total, rep.columns)]))
    t = Table(data, colWidths=order(widths), repeatRows=1)
    style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e293b")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
             ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")])]
    for i in getattr(rep, "highlight", ()):
        col = i if ltr else ncols - 1 - i
        style += [("BACKGROUND", (col, 0), (col, 0), colors.HexColor("#ca8a04")),
                  ("BACKGROUND", (col, 1), (col, last), colors.HexColor("#fef08a"))]
    fills = {"group": "#e0e7ff", "sub": "#f1f5f9", "shared": "#fce7f3", "net": "#dcfce7"}
    for i, stl in (getattr(rep, "row_styles", {}) or {}).items():
        if stl in fills and i < len(rep.rows):
            style.append(("BACKGROUND", (0, i + 1), (-1, i + 1), colors.HexColor(fills[stl])))
    if rep.total:
        style.append(("BACKGROUND", (0, last + 1), (-1, last + 1), colors.HexColor("#fde68a")))
    t.setStyle(TableStyle(style))
    el += [t, Spacer(1, 10)]
    for n in rep.notes:
        el += [P(n, S["right"]), Spacer(1, 2)]
    doc.build(el)



# ============================================================
# عناصر الواجهة المساعدة
# ============================================================
_AR_NORMALIZE = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ة": "ه", "ى": "ي", "ئ": "ي", "ؤ": "و",
                                "ـ": None, "\u064b": None, "\u064c": None, "\u064d": None, "\u064e": None,
                                "\u064f": None, "\u0650": None, "\u0651": None, "\u0652": None,
                                "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4", "٥": "5", "٦": "6", "٧": "7",
                                "٨": "8", "٩": "9"})


def search_norm(text):
    """توحيد النص للبحث: الهمزات والتاء المربوطة والألف المقصورة والتشكيل والأرقام العربية."""
    return " ".join(str(text or "").translate(_AR_NORMALIZE).lower().split())




# ============================================================
# أدوات مشتركة بين الواجهتين (لوحة المؤشرات والشعار) — بدون Tkinter
# ============================================================
DASH_FROM = "1900-01-01"   # بداية «من أول العمل» للوحة المؤشرات


def fetch_dashboard_metrics():
    """مؤشرات اللوحة الرئيسية: تراكمي من بداية العمل حتى اليوم، بدون ضريبة. أرقام الأرباح والمصروفات من income_figures
    (نفس معادلات قائمة الدخل حرفياً) فتتطابق اللوحة مع قائمة الدخل لنفس الفترة:
    مجمل الربح = صافي المبيعات − تكلفة المبيعات؛ صافي الربح = مجمل الربح + إيرادات أخرى − المصروفات التشغيلية − الإهلاك."""
    to = today()
    f = income_figures(DASH_FROM, to)
    with db() as conn:
        balances = cash_balances(conn)
        codes = cash_accounts(conn)
        bc = conn.execute("SELECT customer_name, SUM(subtotal) FROM sales_invoices GROUP BY customer_code ORDER BY SUM(subtotal) DESC LIMIT 1").fetchone()
        bs = conn.execute("SELECT supplier_name, SUM(subtotal) FROM purchase_invoices GROUP BY supplier_code ORDER BY SUM(subtotal) DESC LIMIT 1").fetchone()
        bi = conn.execute("SELECT item_name, SUM(quantity) FROM sales_invoice_lines GROUP BY item_code ORDER BY SUM(quantity) DESC LIMIT 1").fetchone()
    net_sales, stock = f["net_sales"], f["stock_value"]
    bals = partner_balances(to)
    cash_box = sum(v for k, v in balances.items() if is_cash_box(k, codes.get(k)))
    banks = sum(v for k, v in balances.items() if not is_cash_box(k, codes.get(k)))
    recv = sum(r["balance"] for r in bals if r["balance"] > 0.005)
    payb = -sum(r["balance"] for r in bals if r["balance"] < -0.005)
    return {
        "sales": net_sales, "purchases": f["net_purchases"], "cogs": f["cogs"], "gross": f["gross"],
        "other_income": f["total_other"], "expenses": f["opex"], "depreciation": f["dep"], "profit": f["net_profit"],
        "opex_selling": f["opex_groups"]["selling"], "opex_admin": f["opex_groups"]["admin"],
        "opex_other": f["opex_groups"]["other"], "hw_ref_adj": f["hw"]["ref_adj"], "warnings": dashboard_warnings(),
        # المصروفات التشغيلية ومصروف الإهلاك منفصلان
        "cash_box": cash_box, "banks": banks, "cash_total": cash_box + banks,
        "net_position": cash_box + banks + recv - payb,
        "recv": recv, "payb": payb, "stock": stock,
        # (الاسم، القيمة) وتُنسَّق عند العرض بحسب لغة اللوحة
        "best_cust": (bc[0], bc[1] or 0.0) if bc else None,
        "best_supp": (bs[0], bs[1] or 0.0) if bs else None,
        "best_item": (bi[0], bi[1] or 0.0) if bi else None,
    }



def dashboard_warnings():
    """تنبيهات تؤثر على صحة أرقام الأرباح في اللوحة (تظهر تحت معادلة الأرباح)."""
    out = []
    with db() as conn:
        refs = _safe_rows(conn, "SELECT COUNT(*) FROM hw_tb_reference WHERE substr(account_code,1,2)='43' "
                                "OR substr(account_code,1,1)='5' OR account_code='41125'")
        posted = _safe_rows(conn, "SELECT COUNT(*) FROM journal_entries WHERE source_type='hw_tb_expenses'")
        docs = _safe_rows(conn, "SELECT COUNT(*) FROM hw_import_map WHERE role IN ('customer','supplier')")
        segs = hw_ref_segments(conn) if docs and docs[0][0] else []
        adj = hw_cogs_ref_adjustments(conn) if segs else []
    if refs and refs[0][0] and not (posted and posted[0][0]):
        out.append("مصروفات ميزان HansaWorld المرجعية (43xxx/5xxxx/41125) غير مُرحّلة بعد: المصروفات وصافي الربح ناقصة. "
                   "رحّلها من مركز الاستيراد ← استيراد HansaWorld.")
    if segs and not adj and any(c.startswith(HW_COGS_PREFIXES) for sg in segs for c in sg["amounts"]):
        out.append("تاريخ آخر مستند HansaWorld بعد تاريخ الميزان المرجعي: تكلفة المبيعات بالجرد الدوري دون تسوية لدفتر "
                   "HansaWorld. استورد ميزاناً محدّثاً.")
    return out


def dashboard_sources():
    """مصدر كل رقم في لوحة المؤشرات ومعادلته (تقرير «مصادر أرقام لوحة المؤشرات»)."""
    m = fetch_dashboard_metrics()
    f = income_figures(DASH_FROM, today())
    hw, dsc = f["hw"], f["dsc"]
    hw_txt = ""
    if abs(hw["purchases"]) + abs(hw["closing"]) + abs(hw["ref_adj"]) > 0.005:
        hw_txt = (f"فترة HansaWorld (جرد دوري): مشتريات {fmt(hw['purchases'])} − إشعارات موردين {fmt(hw['credits'])} − "
                  f"مخزون ختامي {fmt(hw['closing'])}" +
                  (f" ± تسوية لدفتر HansaWorld (41110+41201) {fmt(hw['ref_adj'])}" if abs(hw["ref_adj"]) > 0.005 else "") +
                  f" = {fmt(hw['cogs'])}")
    og = f["opex_groups"]
    rows = [
        ("صافي المبيعات", m["sales"], f"فواتير البيع قبل الضريبة {fmt(f['sales_sub'])} − مرتجعات المبيعات {fmt(f['rs_sub'])} − "
                                      f"إشعارات دائنة للعملاء {fmt(dsc['sales_sub'])} (sales_invoices / returns_invoices / post_discounts)"),
        ("تكلفة المبيعات", m["cogs"], "؛ ".join(x for x in (
            f"تكلفة بنود الفواتير الجديدة بالمتوسط المرجح {fmt(f['cogs_lines'])}" if abs(f["cogs_lines"]) > 0.005 else "",
            hw_txt) if x) or "تكلفة بنود الفواتير بالمتوسط المرجح − تكلفة المرتجعات"),
        ("مجمل الربح", m["gross"], "صافي المبيعات − تكلفة المبيعات"),
        ("إيرادات أخرى", m["other_income"], "القيود اليومية على حسابات الإيرادات (ومنها 33xxx فروق العملة من ميزان HansaWorld) "
                                           "+ الخصم المكتسب من الموردين + فروق أسعار مردودات المشتريات"),
        ("المصروفات التشغيلية", m["expenses"],
         f"سندات صرف المصروفات + القيود اليومية على حسابات المصروفات (ومنها قيد تجميع ميزان HansaWorld): "
         f"بيعية وتسويقية (43) {fmt(og['selling'])} + إدارية وعمومية (5) {fmt(og['admin'])} + أخرى (41125 …) {fmt(og['other'])}"),
        ("مصروف الإهلاك", m["depreciation"], "إهلاك الأصول الثابتة بالقسط الثابت من تاريخ الشراء حتى اليوم (fixed_assets)"),
        ("صافي الربح", m["profit"], "مجمل الربح + إيرادات أخرى − المصروفات التشغيلية − الإهلاك"),
        ("صافي المشتريات", m["purchases"], "فواتير الشراء قبل الضريبة − مرتجعات المشتريات − إشعارات خصم الموردين"),
        ("قيمة المخزون", m["stock"], "Σ (الرصيد الحالي × متوسط التكلفة) لكل صنف (inventory)"),
        ("الصناديق", m["cash_box"], "أرصدة حسابات النقدية: الرصيد الافتتاحي + المقبوضات − المدفوعات"),
        ("البنوك", m["banks"], "أرصدة الحسابات البنكية: الرصيد الافتتاحي + المقبوضات − المدفوعات"),
        ("الذمم المدينة", m["recv"], "مجموع الأرصدة المدينة للعملاء (كشف حساب كل طرف حتى اليوم)"),
        ("الذمم الدائنة", m["payb"], "مجموع الأرصدة الدائنة للموردين (كشف حساب كل طرف حتى اليوم)"),
        ("صافي المركز", m["net_position"], "النقدية والبنوك + الذمم المدينة − الذمم الدائنة"),
    ]
    notes = ["كل الأرقام تراكمية من بداية العمل حتى اليوم وقبل الضريبة، ومن نفس دالة قائمة الدخل (income_figures) "
             "فتتطابق اللوحة مع قائمة الدخل لنفس الفترة."] + m["warnings"]
    return rows, notes


def rep_dash_sources(p):
    rows, notes = dashboard_sources()
    return Report("مصادر أرقام لوحة المؤشرات", [("البند", "t"), ("القيمة", "m"), ("المصدر والمعادلة", "t")], rows, None,
                  notes, f"من بداية العمل حتى {today()}")


AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
EN_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DASH_REFRESH_MS = 30_000
CHART_GRID, CHART_TEXT, CHART_INK = "#e2e8f0", "#64748b", "#0f172a"
CLR_SALES, CLR_PURCH, CLR_EXP, CLR_CASH, CLR_NEG = "#2563eb", "#f59e0b", "#e11d48", "#0d9488", "#dc2626"

# ألوان ولغة الرسوم الحالية. اللوحة تبدّلها مؤقتاً (داكن/فاتح، عربي/إنجليزي) أثناء الرسم ثم تعيدها للافتراضي.
CHART_PAL_DEFAULT = dict(grid=CHART_GRID, text=CHART_TEXT, ink=CHART_INK, label="#334155", axis="#94a3b8",
                         tip_bg=CHART_INK, tip_fg="white", neg=CLR_NEG, rtl=True, lang="ar")
CHART_PAL = dict(CHART_PAL_DEFAULT)


def _cur():
    return "ر.س" if CHART_PAL["lang"] == "ar" else "SAR"


def fetch_dashboard_charts(months=6):
    """بيانات الرسوم: المبيعات مقابل المشتريات شهرياً، توزيع المصروفات التشغيلية هذه السنة، أرصدة الخزينة والبنوك.
    الأشهر تُرجع كمفاتيح YYYY-MM وتُترجم أسماؤها عند العرض بحسب لغة اللوحة."""
    t = date.today()
    y, m = t.year, t.month - (months - 1)
    while m <= 0:
        m += 12
        y -= 1
    keys = month_keys(f"{y:04d}-{m:02d}-01", t.strftime("%Y-%m-%d"))
    k1, k2 = keys[0], keys[-1]
    with db() as conn:
        def monthly(sql, *params):
            return {k: v or 0.0 for k, v in conn.execute(sql, params)}
        sales = monthly("SELECT substr(invoice_date,1,7), SUM(subtotal) FROM sales_invoices "
                        "WHERE substr(invoice_date,1,7) BETWEEN ? AND ? GROUP BY 1", k1, k2)
        purch = monthly("SELECT substr(purchase_date,1,7), SUM(subtotal) FROM purchase_invoices "
                        "WHERE substr(purchase_date,1,7) BETWEEN ? AND ? GROUP BY 1", k1, k2)
        rets = defaultdict(dict)
        for typ, k, v in conn.execute("SELECT return_type, substr(return_date,1,7), SUM(subtotal) FROM returns_invoices "
                                      "WHERE substr(return_date,1,7) BETWEEN ? AND ? GROUP BY 1, 2", (k1, k2)):
            rets[typ][k] = v or 0.0
        exp = expense_breakdown(conn, f"{t.year}-01-01", t.strftime("%Y-%m-%d"))
        bal = cash_balances(conn)
    exp_items = [(n, v) for n, _, v in exp if v > 0.005]
    if len(exp_items) > 6:
        exp_items = exp_items[:5] + [(None, sum(v for _, v in exp_items[5:]))]  # None = «بنود أخرى»
    cash_items = sorted(((k, v) for k, v in bal.items() if abs(v) > 0.005), key=lambda kv: -kv[1])
    return dict(
        flow=(keys, [sales.get(k, 0.0) - rets[SALES_RETURN].get(k, 0.0) for k in keys],
              [purch.get(k, 0.0) - rets[PURCHASE_RETURN].get(k, 0.0) for k in keys]),
        exp=exp_items, cash=cash_items)



ELHAGRY_LOGO_FILES = ("elhagry.png", "الحاجري.png", "الحاجري.jpg", "الحاجري.jpeg", "elhagry.jpg", "elhagry.jpeg")
LOGIN_PANEL_BG = "#0b1e3f"   # كحلي لوحة الشعار (نفس لون التصميم)
LOGIN_FORM_BG = "#eef3fa"    # الخلفية الفاتحة لنموذج الدخول
LOGIN_ACCENT = "#14b8a6"     # الفيروزي (لون التركيز والإبراز)


def _app_dirs():
    dirs = []
    for d in (getattr(sys, "_MEIPASS", None), os.path.dirname(os.path.abspath(sys.argv[0] or ".")),
              APP_DIR, os.path.dirname(os.path.abspath(__file__)), os.getcwd()):
        if d and d not in dirs:
            dirs.append(d)
    return dirs


def find_app_file(names):
    """يبحث عن أول ملف موجود بالأسماء المعطاة في: مجلد البرنامج، مجلد التشغيل الحالي، ومجلد الحزمة (PyInstaller)."""
    for name in names:
        for d in _app_dirs():
            path = os.path.join(d, name)
            if os.path.isfile(path):
                return path
    return None


def find_logo_file(names=ELHAGRY_LOGO_FILES):
    """أحدث ملف شعار موجود (حتى يظهر الشعار الجديد فور وضعه في مجلد البرنامج)."""
    found = {os.path.normcase(os.path.abspath(os.path.join(d, n)))
             for n in names for d in _app_dirs() if os.path.isfile(os.path.join(d, n))}
    return max(found, key=os.path.getmtime) if found else None


def _color_dist(a, b):
    return max(abs(x - y) for x, y in zip(a, b))


def extract_logo_image(im):
    """يرجع (صورة الشعار، لون الخلفية).
    لو الصورة تصميم شاشة كامل بلوحة جانبية: تُقتطع اللوحة، ثم الجزء الأكبر من المحتوى داخلها (الرمز والاسم)
    دون النصوص الصغيرة البعيدة في الأسفل (مثل سطر الإصدار)."""
    im = im.convert("RGBA")
    flat = Image.new("RGB", im.size, (255, 255, 255))
    flat.paste(im, mask=im.split()[3])
    im = flat
    w, h = im.size
    px = im.load()
    # 1) لوحة جانبية بلون مختلف عن باقي الصورة (يمين أو يسار)
    for side in ("right", "left"):
        edge_x = w - 1 if side == "right" else 0
        bg = px[edge_x, 2]
        if any(_color_dist(px[edge_x, y], bg) > 12 for y in (h // 4, h // 2, h - 3)):
            continue  # الحافة ليست لوناً واحداً → ليست لوحة
        rng = range(w - 1, -1, -1) if side == "right" else range(w)
        border = None
        for x in rng:
            if _color_dist(px[x, 2], bg) > 40 and _color_dist(px[x, h - 3], bg) > 40:
                border = x
                break
        if border is not None and 0.15 * w < abs(border - edge_x) < 0.7 * w:
            box = (border + 1, 0, w, h) if side == "right" else (0, 0, border, h)
            im = im.crop(box)
            break
    w, h = im.size
    px = im.load()
    bg = px[w - 1, 1]
    # 2) صفوف المحتوى → مجموعات تفصلها فراغات كبيرة؛ نأخذ المجموعة الأكبر (الشعار)
    step = max(1, w // 300)
    rows = []
    for y in range(h):
        xs = [x for x in range(0, w, step) if _color_dist(px[x, y], bg) > 40]
        rows.append((min(xs), max(xs), len(xs)) if xs else None)
    gap_limit = max(12, int(h * 0.08))
    bands, cur, blank = [], None, 0
    for y, r in enumerate(rows):
        if r:
            if cur is None or blank > gap_limit:
                cur = [y, y, r[0], r[1], 0]
                bands.append(cur)
            cur[1], cur[2], cur[3] = y, min(cur[2], r[0]), max(cur[3], r[1])
            cur[4] += r[2]
            blank = 0
        else:
            blank += 1
    if bands:
        y0, y1, x0, x1, _ = max(bands, key=lambda b: b[4])
        pad = max(10, int(min(w, h) * 0.04))
        im = im.crop((max(0, x0 - pad), max(0, y0 - pad), min(w, x1 + step + pad), min(h, y1 + 1 + pad)))
    return im, "#%02x%02x%02x" % bg

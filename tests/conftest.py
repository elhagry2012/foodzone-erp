# -*- coding: utf-8 -*-
"""إعداد الاختبارات: كل اختبار يعمل على قاعدة بيانات مؤقتة جديدة (لا تُمس قاعدة البيانات الحقيقية أبداً)."""
import os
import sys
import tempfile

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_TMP = tempfile.mkdtemp(prefix="elhagry_tests_")
os.environ["ELHAGRY_DB"] = os.path.join(_TMP, "boot.db")
os.environ["ELHAGRY_DATA_DIR"] = _TMP
# بيانات دخول الاختبار من البيئة (لا توجد بيانات ثابتة في الكود)، و .env الحقيقي لا يُحمَّل في الاختبارات
os.environ["ELHAGRY_NO_DOTENV"] = "1"
os.environ["ELHAGRY_ADMIN_USER"] = "test.admin"
os.environ["ELHAGRY_ADMIN_PASSWORD"] = TEST_ADMIN_PASSWORD = "Test-Admin-2026"
os.environ["ELHAGRY_RESET_PASSWORD"] = "2026"
os.environ["ELHAGRY_MODIFY_PASSWORD"] = "2026"
for _k in ("ELHAGRY_ADMIN_RESET", "ELHAGRY_ATTACH_DIR"):
    os.environ.pop(_k, None)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import elhagry_core as F  # noqa: E402
from elhagry_core import engine as E  # noqa: E402

_TEMPLATE = {}


@pytest.fixture
def erp(tmp_path):
    """قاعدة بيانات جديدة مهيأة بدليل الحسابات الافتراضي. تُنسخ من قالب مُهيأ مرة واحدة لتسريع الاختبارات."""
    import shutil
    if "path" not in _TEMPLATE:
        tpl = os.path.join(_TMP, "template.db")
        E.DB_FILE = tpl
        E.BACKUP_DIR = os.path.join(_TMP, "backups")
        F.init_db()
        _TEMPLATE["path"] = tpl
    path = str(tmp_path / "test.db")
    shutil.copy(_TEMPLATE["path"], path)
    E.DB_FILE = path
    E.BACKUP_DIR = str(tmp_path / "backups")
    E.set_current_user("tester", "test")
    F.init_db()
    yield F
    E.set_current_user(None)


@pytest.fixture
def world(erp):
    """بيانات أساسية: عميل، مورد، صنفان."""
    F = erp
    cust = F.save_partner_record("customer", None, "عميل الاختبار")
    supp = F.save_partner_record("supplier", None, "مورد الاختبار")
    with F.db() as conn:
        for code, name in (("ELH-T1", "صنف أول"), ("ELH-T2", "صنف ثان")):
            conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, cost_price) "
                         "VALUES (?,?,?,?,?,?,?)", (code, name, "KG", 0, 0, 0, 0))
    return dict(F=F, cust=(cust, "عميل الاختبار"), supp=(supp, "مورد الاختبار"))


def buy(F, supp, item, qty, cost, d="2026-01-05"):
    """أمر شراء ← استلام ← فاتورة. يرجع رقم الفاتورة."""
    oid = F.create_order("purchase", d, supp, "", [F.make_line(item, item, "KG", qty, cost)])
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM purchase_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    rid = F.fulfil_order("purchase", oid, d, "", [(lid, qty, None)])
    iid, _ = F.invoice_from_fulfilments("purchase", supp[0], [rid], d)
    return iid


def sell(F, cust, item, qty, price, d="2026-01-10", invoice=True, discount=0.0, **kw):
    """أمر بيع ← تسليم ← فاتورة (اختياري). يرجع (رقم الفاتورة أو None، رقم الإذن، رقم الأمر)."""
    oid = F.create_order("sales", d, cust, "", [F.make_line(item, item, "KG", qty, price, discount)], **kw)
    with F.db() as conn:
        lid = conn.execute("SELECT line_id FROM sales_order_lines WHERE order_id=?", (oid,)).fetchone()[0]
    did = F.fulfil_order("sales", oid, d, "", [(lid, qty, None)], **kw)
    iid = F.invoice_from_fulfilments("sales", cust[0], [did], d)[0] if invoice else None
    return iid, did, oid

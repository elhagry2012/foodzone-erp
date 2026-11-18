# -*- coding: utf-8 -*-
"""
elhagry_core — المحرك المحاسبي لنظام الحاجري المالي (بدون أي واجهة).

الوحدات:
    engine    المحرك الأساسي: قاعدة البيانات، الدورة المستندية، المخزون، الشركاء، التقارير، PDF/Excel
    ledger    الترحيل الآلي لكل المستندات إلى دفتر الأستاذ العام + ميزان المراجعة الشامل والميزانية العمومية
    credit    الحد الائتماني للعملاء (تنبيه / منع البيع الآجل)
    planning  توقع التدفق النقدي، الموازنة التقديرية مقابل الفعلي، عمولات المناديب
    controls  الأدوار والصلاحيات، مرفقات استلام المخزن، تسوية الدفعات المقدمة، تقرير الفواتير المفتوحة
    audit     سجل التدقيق والرقابة الداخلية (غير قابل للتعديل أو الحذف)

الاستخدام:
    import elhagry_core as F
    F.init_db(); F.run_report("gl_tb", {...})
"""
from .engine import *  # noqa: F401,F403
from . import engine  # noqa: F401
from .ledger import *  # noqa: F401,F403
from .credit import *  # noqa: F401,F403
from .planning import *  # noqa: F401,F403
from .controls import *  # noqa: F401,F403
from .audit import *  # noqa: F401,F403   (آخر وحدة: مشغلات التدقيق تُنشأ بعد كل الجداول)
from . import audit, ledger, credit, planning, controls  # noqa: F401

# إعدادات تتغير وقت التشغيل (مسار قاعدة البيانات في الاختبارات، رسالة الترحيل، الخط): لا تُنسخ في الحزمة
# حتى تُقرأ دائماً حية من المحرك عبر __getattr__ (مثل F.DB_FILE).
_LIVE = ("DB_FILE", "BACKUP_DIR", "DATA_DIR", "MIGRATION_NOTE", "FONT_NAME")
for _n in _LIVE:
    globals().pop(_n, None)


def __getattr__(name):
    """الأسماء الداخلية (التي تبدأ بـ _) والقيم الحية للمحرك تُقرأ منه مباشرة."""
    for mod in (engine, ledger, credit, planning, controls, audit):
        if hasattr(mod, name):
            return getattr(mod, name)
    raise AttributeError(name)

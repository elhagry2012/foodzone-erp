# -*- coding: utf-8 -*-
"""
نظام الحاجري المالي | ELHAGRY Financial Systems

نقطة تشغيل واجهة سطح المكتب + واجهة توافق للأكواد القديمة.

هيكل المشروع:
    elhagry_core/       المحرك المحاسبي (بدون واجهة): engine, ledger, credit, planning, audit
    elhagry_desktop/    واجهة سطح المكتب (Tkinter)
    streamlit_app.py    واجهة الويب (Streamlit)
    tests/              الاختبارات الآلية (python -m pytest)

التشغيل:
    واجهة سطح المكتب (Tkinter):   python foodzone_erp.py
    واجهة الويب (Streamlit):       streamlit run streamlit_app.py

التوافق: `import foodzone_erp as F` ما زال يعمل؛ كل اسم يُقرأ حياً من المحرك (elhagry_core)
ثم من واجهة سطح المكتب (مثل ElhagryERP و LoginWindow).
"""
import elhagry_core as _core


def __getattr__(name):
    try:
        return getattr(_core, name)
    except AttributeError:
        pass
    from elhagry_desktop import app as _desktop
    return getattr(_desktop, name)


def __dir__():
    return sorted(set(dir(_core)))


def main():
    from elhagry_desktop.app import main as _main
    _main()


if __name__ == "__main__":
    main()

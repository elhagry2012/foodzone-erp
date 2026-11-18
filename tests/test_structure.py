# -*- coding: utf-8 -*-
"""هيكلة الكود: المحرك مستقل عن الواجهة، واجهة التوافق، الصلاحيات الجديدة، كل التقارير تعمل، والكود نظيف."""
import os
import subprocess
import sys

import pytest

from conftest import ROOT, buy, sell


def test_engine_has_no_ui_dependency():
    code = ("import sys; import elhagry_core as F; "
            "assert 'tkinter' not in sys.modules, 'engine imported tkinter'; "
            "assert 'streamlit' not in sys.modules; print('ok')")
    env = dict(os.environ, PYTHONPATH=ROOT)
    out = subprocess.run([sys.executable, "-c", code], cwd=ROOT, env=env, capture_output=True, text=True, timeout=120)
    assert out.returncode == 0 and "ok" in out.stdout, out.stderr


def test_compat_facade(erp):
    import foodzone_erp as legacy
    from elhagry_core import engine
    assert legacy.init_db is engine.init_db
    assert legacy.DB_FILE == engine.DB_FILE           # قيمة حية من المحرك
    assert legacy._account_maps is engine._account_maps
    assert legacy.ElhagryERP.__name__ == "ElhagryERP"
    assert legacy.rep_gl_trial_balance is erp.rep_gl_trial_balance


def test_paths_resolve_to_project_root():
    from elhagry_core import engine
    assert os.path.normcase(engine.APP_DIR) == os.path.normcase(ROOT)
    assert os.path.isfile(engine.EXCEL_PATH)
    assert engine.find_logo_file() is not None


def test_new_permissions_and_report_groups(erp):
    F = erp
    for k in ("ledger", "planning", "audit", "reports.ledger", "reports.plan"):
        assert k in F.ALL_PERMISSION_KEYS
    admin = F.authenticate(F.DEFAULT_ADMIN_USER, "Elhagry2012@")
    assert {"ledger", "planning", "audit", "reports.plan"} <= set(admin["perms"])
    assert len(F.REPORT_GROUP_KEYS) == len(F.REPORT_CATALOG)


def test_every_report_runs(world):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 30, 10)
    sell(F, cust, "ELH-T1", 5, 25)
    sell(F, cust, "ELH-T1", 2, 25, invoice=False)
    F.save_return(F.SALES_RETURN, "2026-01-15", cust, ("ELH-T1", "صنف أول"), 1, 25)
    p = dict(d_from="2026-01-01", d_to=F.today(), group="month", stale_days=30, late_days=7, expiry_days=30)
    for key in F.REPORT_INDEX:
        rep = F.run_report(key, p)
        assert rep.title and isinstance(rep.rows, list), key
        for row in rep.rows:
            assert len(row) == len(rep.columns), key


def test_exports_for_new_reports(world, tmp_path):
    F, cust, supp = world["F"], world["cust"], world["supp"]
    buy(F, supp, "ELH-T1", 10, 10)
    sell(F, cust, "ELH-T1", 3, 30)
    p = dict(d_from="2026-01-01", d_to=F.today())
    for key in ("bs", "gl_tb", "cf_forecast", "bva", "commissions", "credit"):
        rep = F.run_report(key, p)
        F.export_report_xlsx(str(tmp_path / f"{key}.xlsx"), rep)
        F.build_report_pdf(str(tmp_path / f"{key}.pdf"), rep)
        assert (tmp_path / f"{key}.pdf").stat().st_size > 1000
    F.build_report_pdf(str(tmp_path / "audit.pdf"), F.audit_report())


def test_code_is_lint_clean():
    pyflakes = pytest.importorskip("pyflakes.api")
    from pyflakes import reporter as rep_mod
    import io
    files = [os.path.join(ROOT, "foodzone_erp.py"), os.path.join(ROOT, "streamlit_app.py")]
    for pkg in ("elhagry_core", "elhagry_desktop"):
        files += [os.path.join(ROOT, pkg, f) for f in os.listdir(os.path.join(ROOT, pkg)) if f.endswith(".py")
                  and f != "__init__.py"]
    out, err = io.StringIO(), io.StringIO()
    n = sum(pyflakes.checkPath(f, rep_mod.Reporter(out, err)) for f in files)
    assert n == 0, out.getvalue() + err.getvalue()

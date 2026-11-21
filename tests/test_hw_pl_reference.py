# -*- coding: utf-8 -*-
"""قائمة دخل HansaWorld لفترة: تقسيم فترة الميزان، تسوية تكلفة المبيعات لدفتر HansaWorld (41110 + 41201)،
قيود تجميع المصاريف والإيرادات الأخرى لكل فترة، ومطابقة لوحة المؤشرات وقائمة الدخل والدفتر العام."""
import openpyxl
import pandas as pd
import pytest

from test_hansa_alloc_reset import _write_hw_files

# ميزان المراجعة (حسابات قائمة الدخل + حساب موازن): المبيعات 3700، فروق عملة 50، تكلفة 2500 + مردود مشتريات 200،
# بيعية 500، إدارية 300 ← صافي الربح = 3700 + 50 − 2700 − 800 = 250
TB_PL = [("31110", "ايراد مبيعات مجمدات", 0, 3700), ("33100", "فروق عملة", 0, 50),
         ("41110", "تكلفة المبيعات", 2500, 0), ("41201", "مردود مشتريات", 200, 0),
         ("43101", "رواتب و اجور - بيعية", 500, 0), ("52014", "ايجارات - ادارية", 300, 0)]
PL_TEXT = """Profit & Loss Standard ERP, Print date: 01/10/2025 10:00 Page: 1
Food Zone Company Period 15/09/2025 : 30/09/2025
Last Reg Date 01/10/2025 Budget Period 01/01/2025 : 31/12/2025
Revenues
-1,000.00 -1,000.00 ايراد مبيعات مجمدات 31110
-20.00 -20.00 فروق عملة 33100
Total Revenues -1,020.00 -1,020.00
800.00 800.00 تكلفة المبيعات 41110
100.00 100.00 ايجارات - ادارية 52014
Gross Profit -120.00 -120.00
Net Profit Or Loos After Zakat -120.00 -120.00
"""


def _write_tb(tmp_path):
    head = ["Account", None, None, "B/Fwd Bal", "Per Deb", "Per Cred", "Acc Deb", "Acc Cred", "Balance", None]
    rows = []
    for code, name, d, c in TB_PL:
        bal = d - c
        rows.append([code, name, None, None, d or None, c or None, d or None, c or None,
                     bal if bal > 0 else None, -bal if bal < 0 else None])
    net = sum(d - c for _, _, d, c in TB_PL)      # −250 (دائن) ← موازن مدين في الصندوق
    rows.append(["11111", "صندوق نقدية الرياض", None, None, -net, None, -net, None, -net, None])
    path = str(tmp_path / "tb_pl.xlsx")
    pd.DataFrame(rows, columns=head).to_excel(path, sheet_name="Analytical Balance Report", index=False)
    return path


def _write_pl_xlsx(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Food Zone Company", None, "Period 15/09/2025 : 30/09/2025"])
    ws.append(["Account", "Name", "Period", "Accumulated"])
    for code, name, amt in ((31110, "ايراد مبيعات مجمدات", -1000.0), (33100, "فروق عملة", -20.0),
                            (41110, "تكلفة المبيعات", 800.0), (52014, "ايجارات - ادارية", 100.0)):
        ws.append([code, name, amt, amt])
    ws.append(["Net Profit Or Loos After Zakat", None, -120.0, -120.0])
    path = str(tmp_path / "pl.xlsx")
    wb.save(path)
    return path


@pytest.fixture
def hw(erp, tmp_path):
    from elhagry_core import hansa_import as H
    H.run_hansa_import(**_write_hw_files(tmp_path), balance=_write_tb(tmp_path), tb_as_of="2025-09-30")
    return erp, H, tmp_path


def test_parse_pl_text_and_excel(hw):
    F, H, tmp = hw
    t = H.parse_pl_rows([ln.split() for ln in PL_TEXT.splitlines()])
    x = H.parse_pl_rows(H._pl_rows(_write_pl_xlsx(tmp)))
    for p in (t, x):
        assert p["period"] == ("2025-09-15", "2025-09-30") and not p["problems"]   # Budget Period تُتجاهل
        assert {c: a for c, (_, a) in p["accounts"].items()} == {"31110": -1000, "33100": -20, "41110": 800, "52014": 100}
        assert p["net"] == -120
    bad = H.parse_pl_rows([["Period", "01/01/2025", ":", "31/01/2025"], ["5.00", "x", "52014"], ["Net", "Profit", "9.00"]])
    assert bad["problems"]                                         # المجموع لا يساوي صافي الربح


def test_cogs_adjusted_to_hansaworld_ledger_per_period(hw):
    F, H, tmp = hw
    # الميزان وحده: فترة واحدة، والتكلفة التراكمية = 41110 + 41201 بدل الجرد الدوري (4000 − 1332.5)
    with F.db() as conn:
        assert [(s["d_from"], s["d_to"], s["src"]) for s in F.hw_ref_segments(conn)] == [("2025-09-01", "2025-09-30", "tb")]
    assert round(F.income_figures(F.DASH_FROM, F.today())["cogs"], 2) == 2700
    path = tmp / "pl.txt"
    path.write_text(PL_TEXT, encoding="utf-8")
    dry = H.import_pl_reference(str(path), dry_run=True)
    assert dry["period"] == ("2025-09-15", "2025-09-30") and dry["totals"]["net_profit"] == 120
    with F.db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM hw_pl_reference").fetchone()[0] == 0     # المعاينة لا تحفظ
    H.import_pl_reference(str(path))
    with F.db() as conn:
        segs = F.hw_ref_segments(conn)
        adj = {a["date"]: a["adj"] for a in F.hw_cogs_ref_adjustments(conn)}
    assert [(s["d_from"], s["d_to"], s["src"]) for s in segs] == [("2025-09-01", "2025-09-14", "tb_rest"),
                                                                  ("2025-09-15", "2025-09-30", "pl")]
    assert segs[0]["amounts"]["41201"][1] == 200 and segs[0]["amounts"]["41110"][1] == 1700
    assert adj == {"2025-09-14": 1900 - 4000, "2025-09-30": 800 + 1332.5}    # دفتر − جرد دوري لكل فترة
    assert round(F.income_figures(F.DASH_FROM, "2025-09-14")["cogs"], 2) == 1900
    assert round(F.income_figures("2025-09-15", "2025-09-30")["cogs"], 2) == 800
    assert round(F.income_figures(F.DASH_FROM, F.today())["cogs"], 2) == 2700
    with pytest.raises(F.UserError, match="تتداخل"):
        H.import_pl_reference(str(path), "2025-09-20", "2025-09-25")


def test_aggregate_per_period_with_other_income_and_dashboard(hw):
    F, H, tmp = hw
    from elhagry_core import ledger as L
    path = tmp / "pl.txt"
    path.write_text(PL_TEXT, encoding="utf-8")
    H.import_pl_reference(str(path))
    m0 = F.fetch_dashboard_metrics()
    assert any("غير مُرحّلة" in w for w in m0["warnings"])
    out = H.post_tb_expense_aggregate()
    assert out["total"] == 800 and out["other_income"] == 50 and len(out["entry_ids"]) == 2
    assert H.post_tb_expense_aggregate()["entry_ids"] == out["entry_ids"]            # إعادة الترحيل لا تكرر
    with F.db() as conn:
        lines = conn.execute("SELECT e.entry_date, l.account_code, ROUND(l.debit - l.credit, 2) FROM journal_lines l "
                             "JOIN journal_entries e ON e.entry_id=l.entry_id WHERE e.source_type='hw_tb_expenses' "
                             "ORDER BY 1, 2").fetchall()
    assert lines == [("2025-09-14", "33100", -30), ("2025-09-14", "43101", 500), ("2025-09-14", "52014", 200),
                     ("2025-09-14", "99999", -670), ("2025-09-30", "33100", -20), ("2025-09-30", "52014", 100),
                     ("2025-09-30", "99999", -80)]
    f2 = F.income_figures("2025-09-15", "2025-09-30")
    assert round(f2["total_other"], 2) == 20 and round(f2["opex"], 2) == 100 and round(f2["opex_groups"]["admin"], 2) == 100
    m = F.fetch_dashboard_metrics()
    assert not m["warnings"]
    assert (round(m["cogs"], 2), round(m["other_income"], 2), round(m["expenses"], 2)) == (2700, 50, 800)
    assert (round(m["opex_selling"], 2), round(m["opex_admin"], 2), round(m["opex_other"], 2)) == (500, 300, 0)
    assert round(m["profit"], 2) == round(m["sales"] - 2700 + 50 - 800, 2)
    bs = L.balance_sheet(F.today())
    assert bs["balanced"] and round(bs["profit"], 2) == round(m["profit"], 2)           # الدفتر العام = اللوحة
    rows = {(r[0][:23], r[1]): r for r in H.rep_hw_pl_compare({}).rows}
    seg2 = "2025-09-15 → 2025-09-30"
    for item in ("تكلفة المبيعات (41110 + 41201)", "الإيرادات الأخرى (33xxx فروق عملة وخدمات)",
                 "مصروفات إدارية وعمومية (5xxxx)"):
        assert rows[(seg2, item)][5] == "✅ مطابق", rows[(seg2, item)]
    src = dict((r[0], r[2]) for r in F.rep_dash_sources({}).rows)
    assert "تسوية لدفتر HansaWorld" in src["تكلفة المبيعات"] and "بيعية وتسويقية (43) 500.00" in src["المصروفات التشغيلية"]
    # حذف القائمة يعيد فترة واحدة ويعيد بناء القيود بلا تكرار
    assert H.remove_pl_reference("2025-09-15", "2025-09-30") == 4                 # أسطر الحسابات الأربعة
    st = H.tb_expense_status()
    assert len(st["segments"]) == 1 and st["posted"][2] == 800
    assert round(F.income_figures(F.DASH_FROM, F.today())["cogs"], 2) == 2700


def test_no_cogs_adjustment_when_tb_lacks_cogs_accounts(erp, tmp_path):
    """ميزان بلا حسابات تكلفة (مثل ملف الاختبار القديم) ← يبقى الجرد الدوري كما هو."""
    F = erp
    from elhagry_core import hansa_import as H
    from test_hansa_alloc_reset import _write_tb as tb_old
    H.run_hansa_import(**_write_hw_files(tmp_path), balance=tb_old(tmp_path), tb_as_of="2025-09-30")
    with F.db() as conn:
        assert F.hw_cogs_ref_adjustments(conn) == []
    assert round(F.income_figures(F.DASH_FROM, F.today())["cogs"], 2) == round(4600 / 1.15 - 65 * 20.5, 2)

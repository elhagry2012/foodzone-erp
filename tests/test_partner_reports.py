# -*- coding: utf-8 -*-
"""تقارير العملاء والموردين (إجماليات / تفاصيل مجمّعة) وتقرير المقاصة، وإخفاء الأكواد القديمة من العرض."""
import pytest

from conftest import buy, sell


@pytest.fixture
def parties(world):
    """عميل عادي، مورد عادي، وكيان مشترك (نفس الاسم عميلاً C ومورداً S) له حركات في الجانبين."""
    F, cust, supp = world["F"], world["cust"], world["supp"]
    both_c = F.save_partner_record("customer", None, "شركة مشتركة")
    both_s = F.save_partner_record("supplier", None, "شركة مشتركة")
    buy(F, supp, "ELH-T1", 10, 10)                       # مورد الاختبار: له 115
    buy(F, (both_s, "شركة مشتركة"), "ELH-T1", 10, 20)    # المشترك كمورد: له 230
    sell(F, cust, "ELH-T1", 2, 100)                       # عميل الاختبار: عليه 230
    sell(F, (both_c, "شركة مشتركة"), "ELH-T1", 1, 100)    # المشترك كعميل: عليه 115
    with F.db() as conn:   # طرف قديم مؤرشف بعد إعادة الترقيم: لا يظهر
        conn.execute("INSERT INTO business_partners (partner_code, name, is_customer, is_supplier) "
                     "VALUES ('OLD-C0099', 'شركة مشتركة', 1, 1)")
    return dict(F=F, cust=cust[0], supp=supp[0], both_c=both_c, both_s=both_s)


P = dict(d_from="2026-01-01", d_to="2026-12-31")


def test_customers_summary_and_detail(parties):
    F, c, bc = parties["F"], parties["cust"], parties["both_c"]
    s = F.run_report("pt_cust", dict(P, mode="summary"))
    by = {r[0]: r for r in s.rows}
    assert set(by) == {c, bc}                               # بدون الكود القديم OLD-
    assert round(by[c][6], 2) == 230 and round(by[bc][6], 2) == 115
    assert by[bc][10].startswith("نفس الاسم") and not by[c][10]
    assert round(s.total[6], 2) == 345
    d = F.run_report("pt_cust", dict(P, mode="detail"))
    codes = [r[0] for i, r in enumerate(d.rows) if d.row_styles.get(i) == "group"]
    assert codes == sorted([c, bc])                          # كل عميل ثم الذي يليه بترتيب الكود
    subs = [r for i, r in enumerate(d.rows) if d.row_styles.get(i) == "sub"]
    assert [round(r[8], 2) for r in subs] == [round(by[k][6], 2) for k in sorted([c, bc])]
    assert round(d.total[8], 2) == 345 and round(d.total[6], 2) == 345
    sh = F.run_report("pt_cust", dict(P, mode="summary", shared_only=True))
    assert [r[0] for r in sh.rows] == [bc]


def test_suppliers_detail_shows_balance_as_payable(parties):
    F = parties["F"]
    d = F.run_report("pt_supp", dict(P, mode="detail"))
    subs = {r[0]: r[8] for i, r in enumerate(d.rows) if d.row_styles.get(i) == "sub"}
    assert {k: round(v, 2) for k, v in subs.items()} == {parties["supp"]: 115, parties["both_s"]: 230}
    assert round(d.total[8], 2) == 345
    s = F.run_report("pt_supp", dict(P, mode="summary"))
    assert {r[0]: r[7] for r in s.rows}[parties["supp"]] == "دائن (له)"


def test_partner_netting_summary_and_detail(parties):
    F, bc, bs = parties["F"], parties["both_c"], parties["both_s"]
    s = F.run_report("pt_net", dict(P, mode="summary"))
    first = s.rows[0]                                        # المشترك أولاً ومظلل
    assert (first[0], first[1], first[3]) == (bc, bs, "نفس الاسم") and s.row_styles.get(0) == "shared"
    assert (round(first[4], 2), round(first[5], 2), round(first[6], 2)) == (115, 230, -115)
    assert len(s.rows) == 3 and round(s.total[6], 2) == round(345 - 345, 2)
    assert all("OLD-" not in str(r[0]) + str(r[1]) for r in s.rows)
    only = F.run_report("pt_net", dict(P, mode="summary", shared_only=True))
    assert len(only.rows) == 1
    d = F.run_report("pt_net", dict(P, mode="detail", shared_only=True))
    labels = [r[1] for r in d.rows if r[1]]
    assert "— حركاته كعميل" in labels and "— حركاته كمورد" in labels and "صافي المقاصة" in labels
    net = next(r for i, r in enumerate(d.rows) if d.row_styles.get(i) == "net")
    assert round(net[8], 2) == -115 and net[4] == "له"


def test_nonzero_filter_hides_settled_party(parties):
    F, c = parties["F"], parties["cust"]
    F.save_voucher("قبض", "2026-02-01", "صندوق نقدية الرياض", F.CAT_PARTNER, (c, "عميل الاختبار"), 230)
    all_rows = F.run_report("pt_cust", dict(P, mode="summary"))
    nz = F.run_report("pt_cust", dict(P, mode="summary", nonzero=True))
    assert c in [r[0] for r in all_rows.rows] and c not in [r[0] for r in nz.rows]


def test_prior_balance_carried_into_period(parties):
    F, c = parties["F"], parties["cust"]
    d = F.run_report("pt_cust", dict(d_from="2026-06-01", d_to="2026-12-31", mode="detail"))
    i = next(i for i, r in enumerate(d.rows) if r[0] == c and d.row_styles.get(i) == "group")
    assert d.rows[i + 1][4] == "رصيد ما قبل الفترة" and round(d.rows[i + 1][8], 2) == 230


def test_party_reports_export(parties, tmp_path):
    F = parties["F"]
    for key in F.PARTY_REPORTS:
        for mode in F.PT_MODES:
            rep = F.run_report(key, dict(P, mode=mode))
            F.export_report_xlsx(str(tmp_path / f"{key}_{mode}.xlsx"), rep)
            F.build_report_pdf(str(tmp_path / f"{key}_{mode}.pdf"), rep)
    assert F.report_fin_needs("pt_net") == ["customer", "supplier"] and F.report_fin_needs("pt_cust") == ["customer"]


def test_archived_and_legacy_codes_hidden_from_lists(parties):
    F = parties["F"]
    with F.db() as conn:
        conn.execute("UPDATE business_partners SET legacy_code='C0001' WHERE partner_code=?", (parties["both_s"],))
        links = F.partner_link_map(conn)
    codes = [r[0] for r in F.list_partners("customer")]
    assert "OLD-C0099" not in codes and all("OLD-" not in c for c in links)
    assert sorted(links[parties["both_c"]]) == sorted([parties["both_c"], parties["both_s"]])
    # الكود القديم يبقى في الخلفية (للبحث والمطابقة) ولا يظهر كعمود
    assert [r[0] for r in F.list_partners("supplier", "C0001")] == [parties["both_s"]]

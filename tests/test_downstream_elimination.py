import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from downstream_elimination import (
    combined_individual_entry,
    compute_downstream,
    textbook_example_7_12_all_sold_same_year,
    textbook_example_7_12_sold_70,
    textbook_example_7_12_unsold,
    textbook_example_ppe_year1,
)


class DownstreamInventoryTests(unittest.TestCase):
    def test_textbook_unsold(self):
        result = textbook_example_7_12_unsold()
        self.assertEqual(result.gross_profit, 400)
        self.assertEqual(result.unrealized_gross_profit, 400)
        self.assertEqual(result.investor_share_unrealized, 80)
        self.assertEqual(result.share_of_book_ni, 400)
        self.assertEqual(result.individual_investment_income, 320)
        combined = combined_individual_entry(result)
        self.assertIsNotNone(combined)
        self.assertTrue(combined.is_balanced())
        cons = result.consolidation_entries[0]
        self.assertTrue(cons.is_balanced())
        self.assertEqual(cons.lines[0].debit, 200)
        self.assertEqual(cons.lines[1].credit, 120)
        self.assertEqual(cons.lines[2].credit, 80)
        # 合并净利润与个别报表一致：收入成本重分类不改变净利润
        self.assertEqual(
            result.individual_investment_income + 80 - (200 - 120),
            result.individual_investment_income,
        )

    def test_textbook_sold_70_percent(self):
        result = textbook_example_7_12_sold_70()
        self.assertEqual(result.unrealized_gross_profit, 120)  # 400 * 30%
        self.assertEqual(result.investor_share_unrealized, 24)
        self.assertEqual(result.individual_investment_income, 376)
        cons = result.consolidation_entries[0]
        self.assertEqual(cons.lines[0].debit, 200)
        self.assertEqual(cons.lines[1].credit, 120)
        self.assertEqual(cons.lines[2].credit, 80)

    def test_textbook_all_sold_same_year(self):
        result = textbook_example_7_12_all_sold_same_year()
        self.assertEqual(result.unrealized_gross_profit, 0)
        self.assertEqual(result.investor_share_unrealized, 0)
        self.assertEqual(result.individual_investment_income, 400)
        cons = result.consolidation_entries[0]
        self.assertEqual(cons.lines[2].credit, 80)

    def test_subsequent_year_inventory_realization(self):
        result = compute_downstream(
            ownership=0.20,
            selling_price=1000,
            cost=600,
            remaining_ratio=0.0,
            associate_book_ni=3600,
            sale_year=False,
        )
        self.assertEqual(result.opening_unrealized_gross_profit, 400)
        self.assertEqual(result.investor_share_unrealized, 0)
        self.assertEqual(result.period_investor_deferral, -80)
        self.assertEqual(result.individual_investment_income, 800)
        self.assertIn("以后年度", result.consolidation_entries[0].title)

    def test_only_ownership_percentage_eliminated(self):
        result = compute_downstream(
            ownership=0.30,
            selling_price=1000,
            cost=700,
            remaining_ratio=1.0,
            associate_book_ni=0,
        )
        self.assertEqual(result.investor_share_unrealized, 90)
        self.assertEqual(result.group_keep_gross_profit, 210)

    def test_impairment_loss_not_eliminated(self):
        result = compute_downstream(
            ownership=0.25,
            selling_price=80,
            cost=100,
            remaining_ratio=1.0,
            associate_book_ni=400,
            is_impairment_loss=True,
        )
        self.assertEqual(result.investor_share_unrealized, 0)
        self.assertEqual(result.individual_investment_income, 100)
        self.assertEqual(result.consolidation_entries, [])

    def test_business_exception_full_recognition(self):
        result = compute_downstream(
            ownership=0.40,
            selling_price=500,
            cost=200,
            remaining_ratio=1.0,
            associate_book_ni=0,
            constitutes_business=True,
        )
        self.assertEqual(result.investor_share_unrealized, 0)
        self.assertEqual(result.individual_entries, [])

    def test_deferred_tax(self):
        result = compute_downstream(
            ownership=0.20,
            selling_price=1000,
            cost=600,
            remaining_ratio=1.0,
            associate_book_ni=2000,
            tax_rate=0.25,
        )
        self.assertEqual(result.deferred_tax_asset, 20)
        tax_entry = result.individual_entries[-1]
        self.assertEqual(tax_entry.lines[0].account, "递延所得税资产")
        self.assertEqual(tax_entry.lines[0].debit, 20)


class DownstreamPPETests(unittest.TestCase):
    def test_ppe_year1_one_month(self):
        result = textbook_example_ppe_year1()
        self.assertEqual(result.gross_profit, 300)
        self.assertEqual(result.excess_depreciation, 2.5)
        self.assertEqual(result.unrealized_gross_profit, 297.5)
        self.assertEqual(result.investor_share_unrealized, 59.5)
        self.assertEqual(result.individual_investment_income, 140.5)
        cons = result.consolidation_entries[0]
        self.assertEqual(cons.lines[0].debit, 180)
        self.assertEqual(cons.lines[1].credit, 120)
        self.assertEqual(cons.lines[2].credit, 60)

    def test_ppe_year2_depreciation_realization(self):
        result = compute_downstream(
            ownership=0.20,
            selling_price=900,
            cost=600,
            associate_book_ni=1000,
            asset_type="ppe",
            useful_life_years=10,
            months_used_in_year=12,
            cumulative_months_used=13,
            sale_year=False,
        )
        # 累计多提折旧 2.5 + 30 = 32.5，未实现 267.5，投资方份额 53.5
        self.assertEqual(result.excess_depreciation, 30)
        self.assertEqual(result.unrealized_gross_profit, 267.5)
        self.assertEqual(result.investor_share_unrealized, 53.5)
        self.assertEqual(result.period_investor_deferral, -6)
        self.assertEqual(result.individual_investment_income, 206)

    def test_journal_entries_balanced(self):
        for fn in (
            textbook_example_7_12_unsold,
            textbook_example_7_12_sold_70,
            textbook_example_7_12_all_sold_same_year,
            textbook_example_ppe_year1,
        ):
            result = fn()
            for entry in result.all_entries():
                if entry.lines:
                    self.assertTrue(entry.is_balanced(), entry.title)


if __name__ == "__main__":
    unittest.main()

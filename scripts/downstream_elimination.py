"""销售至联营/合营企业顺流交易抵消核算引擎。

依据：
- 《企业会计准则第2号——长期股权投资》（财会〔2014〕14号）
- 《企业会计准则第33号——合并财务报表》（财会〔2014〕10号）
- 财政部会计准则委员会 2023-01-05 问答
- IAS 28 Investments in Associates and Joint Ventures
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, List, Optional


def r2(value: float) -> float:
    return round(float(value) + (1e-12 if value >= 0 else -1e-12), 2)


def r4(value: float) -> float:
    return round(float(value) + (1e-12 if value >= 0 else -1e-12), 4)


@dataclass(frozen=True)
class JournalLine:
    account: str
    debit: float = 0.0
    credit: float = 0.0
    note: str = ""

    def balanced_amount(self) -> float:
        return r2(self.debit) - r2(self.credit)


@dataclass
class JournalEntry:
    title: str
    books: str
    lines: List[JournalLine]
    explanation: str = ""

    def is_balanced(self) -> bool:
        return abs(sum(line.balanced_amount() for line in self.lines)) < 0.005

    def as_text(self) -> str:
        rows = [f"【{self.books}】{self.title}"]
        for line in self.lines:
            if line.debit:
                rows.append(f"  借：{line.account}  {fmt_amount(line.debit)}")
            if line.credit:
                prefix = "    " if any(x.debit for x in self.lines) else "  "
                rows.append(f"{prefix}贷：{line.account}  {fmt_amount(line.credit)}")
        if self.explanation:
            rows.append(f"  说明：{self.explanation}")
        return "\n".join(rows)


@dataclass
class DownstreamResult:
    """一笔顺流交易在某一报告期的完整核算结果。"""

    asset_type: str
    ownership: float
    selling_price: float
    cost: float
    gross_profit: float
    remaining_ratio: float
    realized_ratio: float
    associate_book_ni: float
    excess_depreciation: float
    opening_unrealized_gross_profit: float
    unrealized_gross_profit: float
    investor_share_unrealized: float
    period_investor_deferral: float
    share_of_book_ni: float
    individual_investment_income: float
    individual_lt_investment_adj: float
    deferred_tax_asset: float
    deferred_tax_expense: float
    tax_rate: float
    sale_year: bool
    individual_entries: List[JournalEntry] = field(default_factory=list)
    consolidation_entries: List[JournalEntry] = field(default_factory=list)
    narrative: List[str] = field(default_factory=list)
    checklist: List[str] = field(default_factory=list)

    @property
    def group_keep_gross_profit(self) -> float:
        """顺流交易中可立即确认的部分：归属于联营其他股东的毛利。"""
        return r2(self.gross_profit * (1.0 - self.ownership))

    def all_entries(self) -> List[JournalEntry]:
        return list(self.individual_entries) + list(self.consolidation_entries)


def fmt_amount(value: float) -> str:
    return f"{value:,.2f}"


def fmt_pct(value: float) -> str:
    return f"{value * 100:.2f}%"


def _entry(title: str, books: str, lines: Iterable[JournalLine], explanation: str = "") -> JournalEntry:
    entry = JournalEntry(title=title, books=books, lines=list(lines), explanation=explanation)
    if not entry.is_balanced():
        raise ValueError(f"分录不平衡：{title}")
    return entry


def compute_downstream(
    *,
    ownership: float,
    selling_price: float,
    cost: float,
    remaining_ratio: float = 1.0,
    associate_book_ni: float = 0.0,
    asset_type: str = "inventory",
    useful_life_years: float = 0.0,
    months_used_in_year: float = 0.0,
    cumulative_months_used: Optional[float] = None,
    opening_unrealized_gross_profit: Optional[float] = None,
    tax_rate: float = 0.0,
    residual_value: float = 0.0,
    sale_year: bool = True,
    is_impairment_loss: bool = False,
    constitutes_business: bool = False,
    unit: str = "万元",
) -> DownstreamResult:
    """计算销售至联营公司顺流交易的抵消核算。

    remaining_ratio: 期末仍由联营企业持有、尚未对外部独立第三方出售或耗用的比例。
    固定资产场景下，remaining_ratio 一般取 1，未实现损益通过折旧实现。
    opening_unrealized_gross_profit: 期初尚未实现的内部毛利（存货后续年度必填；PPE 可按累计折旧推算）。
    sale_year: 是否为顺流销售发生当年。后续年度合并报表通常不再重分类收入成本。
    """
    if not 0 < ownership <= 1:
        raise ValueError("持股比例必须在 (0, 1] 之间")
    if remaining_ratio < 0 or remaining_ratio > 1:
        raise ValueError("未实现比例必须在 [0, 1] 之间")
    if asset_type not in {"inventory", "ppe"}:
        raise ValueError("asset_type 只能是 inventory 或 ppe")

    gross_profit = r2(selling_price - cost)
    excess_depreciation = 0.0
    realized_ratio = r2(1.0 - remaining_ratio)
    opening_unrealized = 0.0 if opening_unrealized_gross_profit is None else r2(opening_unrealized_gross_profit)

    if constitutes_business:
        unrealized_gp = 0.0
        remaining_ratio = 0.0
        realized_ratio = 1.0
        opening_unrealized = 0.0
    elif is_impairment_loss and gross_profit < 0:
        unrealized_gp = 0.0
        remaining_ratio = 0.0
        realized_ratio = 1.0
        opening_unrealized = 0.0
    elif asset_type == "ppe":
        if useful_life_years <= 0:
            raise ValueError("固定资产顺流交易必须提供预计使用寿命")
        depreciable_gp = r2(gross_profit - residual_value)
        annual_excess = r2(depreciable_gp / useful_life_years)
        months = months_used_in_year if sale_year else (months_used_in_year or 12.0)
        if cumulative_months_used is None:
            cumulative_months_used = months if sale_year else months
        life_months = useful_life_years * 12.0
        cumulative_months_used = min(max(cumulative_months_used, 0.0), life_months)
        current_year_excess = r2(annual_excess * (months / 12.0))
        cumulative_excess = r2(annual_excess * (cumulative_months_used / 12.0))
        remaining_unrealized_gp = r2(max(gross_profit - cumulative_excess, 0.0))
        excess_depreciation = current_year_excess
        unrealized_gp = remaining_unrealized_gp
        if opening_unrealized_gross_profit is None:
            opening_excess = r2(max(cumulative_excess - current_year_excess, 0.0))
            opening_unrealized = 0.0 if sale_year else r2(max(gross_profit - opening_excess, 0.0))
        remaining_ratio = 0.0 if abs(gross_profit) < 0.005 else r4(unrealized_gp / gross_profit)
        realized_ratio = r4(1.0 - remaining_ratio)
    else:
        unrealized_gp = r2(gross_profit * remaining_ratio)
        if opening_unrealized_gross_profit is None:
            opening_unrealized = 0.0 if sale_year else r2(gross_profit)
        excess_depreciation = 0.0

    investor_share = r2(unrealized_gp * ownership)
    period_investor_deferral = r2((unrealized_gp - opening_unrealized) * ownership)
    share_of_book_ni = r2(associate_book_ni * ownership)
    # 投资收益 = 应享联营账面净利润份额 − 本期未实现变动中归属于投资方的部分
    individual_ii = r2(share_of_book_ni - period_investor_deferral)
    individual_lt_adj = individual_ii
    ending_dta = r2(investor_share * tax_rate) if tax_rate and investor_share > 0 else 0.0
    opening_dta = (
        r2(r2(opening_unrealized * ownership) * tax_rate) if tax_rate and opening_unrealized > 0 else 0.0
    )
    deferred_tax = ending_dta
    deferred_tax_expense = r2(opening_dta - ending_dta)

    individual_entries: List[JournalEntry] = []
    consolidation_entries: List[JournalEntry] = []
    narrative: List[str] = []
    checklist: List[str] = []

    unit_note = f"（单位：{unit}）"
    narrative.append(
        f"顺流交易毛利 {fmt_amount(gross_profit)}{unit_note}，持股比例 {fmt_pct(ownership)}。"
    )
    if constitutes_business:
        narrative.append("该交易构成业务，未实现损益不予递延，投资方全额确认处置损益。")
    elif is_impairment_loss and gross_profit < 0:
        narrative.append("内部交易损失属于所转让资产发生减值，损失全额确认，不作抵销。")
    else:
        narrative.append(
            f"期末未实现对外部独立第三方出售/耗用的内部毛利 {fmt_amount(unrealized_gp)}，"
            f"其中归属于投资方、仍递延在长期股权投资中的余额 = {fmt_amount(unrealized_gp)} × {fmt_pct(ownership)} = {fmt_amount(investor_share)}。"
        )
        if abs(opening_unrealized) > 0.005:
            narrative.append(
                f"期初未实现内部毛利 {fmt_amount(opening_unrealized)}，本期变动 {fmt_amount(unrealized_gp - opening_unrealized)}，"
                f"归属于投资方的本期递延（负数表示转回实现）= {fmt_amount(period_investor_deferral)}。"
            )
        narrative.append(
            f"归属于联营企业其他股东的毛利 {fmt_amount(gross_profit * (1.0 - ownership))} 可立即确认，不抵销。"
        )

    if associate_book_ni or share_of_book_ni:
        individual_entries.append(
            _entry(
                title="按联营企业账面净利润确认权益法投资收益",
                books="个别财务报表",
                lines=[
                    JournalLine("长期股权投资——损益调整", debit=share_of_book_ni),
                    JournalLine("投资收益", credit=share_of_book_ni),
                ],
                explanation=f"联营企业账面净利润 {fmt_amount(associate_book_ni)} × {fmt_pct(ownership)}",
            )
        )

    if period_investor_deferral > 0.005:
        individual_entries.append(
            _entry(
                title="抵销顺流交易未实现损益中归属于投资方的部分",
                books="个别财务报表",
                lines=[
                    JournalLine("投资收益", debit=period_investor_deferral),
                    JournalLine("长期股权投资——损益调整", credit=period_investor_deferral),
                ],
                explanation="只抵持股比例对应部分，不是全额抵销。",
            )
        )
    elif period_investor_deferral < -0.005:
        realized = abs(period_investor_deferral)
        individual_entries.append(
            _entry(
                title="转回已实现的顺流未实现损益",
                books="个别财务报表",
                lines=[
                    JournalLine("长期股权投资——损益调整", debit=realized),
                    JournalLine("投资收益", credit=realized),
                ],
                explanation="联营企业已对外部出售、耗用或计提折旧，上期递延的内部损益在本期实现。",
            )
        )
    elif not constitutes_business and not (is_impairment_loss and gross_profit < 0) and abs(unrealized_gp) < 0.005:
        narrative.append("期末内部商品已全部对外部独立第三方出售或已提足折旧，本期无需再递延未实现损益。")

    dta_move = r2(ending_dta - opening_dta)
    if dta_move > 0.005:
        individual_entries.append(
            _entry(
                title="未实现内部交易损益的递延所得税",
                books="个别财务报表（常见实务处理）",
                lines=[
                    JournalLine("递延所得税资产", debit=dta_move),
                    JournalLine("所得税费用", credit=dta_move),
                ],
                explanation=(
                    f"会计上递延投资方份额 {fmt_amount(investor_share)}，"
                    f"可抵扣暂时性差异 × {fmt_pct(tax_rate)}。"
                ),
            )
        )
    elif dta_move < -0.005:
        individual_entries.append(
            _entry(
                title="转回未实现损益对应的递延所得税",
                books="个别财务报表（常见实务处理）",
                lines=[
                    JournalLine("所得税费用", debit=abs(dta_move)),
                    JournalLine("递延所得税资产", credit=abs(dta_move)),
                ],
                explanation="未实现损益在本期实现，相应转回递延所得税资产。",
            )
        )

    if sale_year and not constitutes_business and not (is_impairment_loss and gross_profit < 0):
        rev = r2(selling_price * ownership)
        cogs = r2(cost * ownership)
        ii_plug = r2(rev - cogs)
        if asset_type == "inventory":
            consolidation_entries.append(
                _entry(
                    title="合并报表：按持股比例重分类顺流收入与成本",
                    books="合并财务报表工作底稿",
                    lines=[
                        JournalLine("营业收入", debit=rev, note="售价 × 持股比例"),
                        JournalLine("营业成本", credit=cogs, note="成本 × 持股比例"),
                        JournalLine("投资收益", credit=ii_plug, note="毛利 × 持股比例"),
                    ],
                    explanation=(
                        "无论联营企业当年对外出售多少，销售当年合并报表均按持股比例全额重分类收入成本；"
                        "该分录不改变合并净利润，只改变列报项目。个别报表已通过投资收益递延未实现部分。"
                    ),
                )
            )
        else:
            # 若投资方出售的是存货、联营作为固定资产，仍走收入成本；
            # 若投资方出售的是固定资产，则抵销资产处置收益。
            consolidation_entries.append(
                _entry(
                    title="合并报表：按持股比例重分类顺流处置损益",
                    books="合并财务报表工作底稿",
                    lines=[
                        JournalLine("营业收入", debit=rev, note="如原列报为收入"),
                        JournalLine("营业成本", credit=cogs),
                        JournalLine("投资收益", credit=ii_plug),
                    ],
                    explanation=(
                        "若投资方原作为固定资产处置，可将借方改为“资产处置收益”"
                        f"（{fmt_amount(gross_profit * ownership)}），不再拆收入成本。"
                    ),
                )
            )
        narrative.append(
            "合并报表在个别报表已抵销的基础上，把记入投资收益的抵销重分类到营业收入/营业成本（或资产处置收益）。"
        )
    elif not sale_year:
        narrative.append("以后年度：个别报表将上期递延的未实现损益随对外出售或折旧逐步转回；合并报表通常不再重做收入成本抵销。")
        consolidation_entries.append(
            JournalEntry(
                title="合并报表：以后年度一般无需再调整",
                books="合并财务报表工作底稿",
                lines=[],
                explanation="上期损益类项目已转入留存收益；未实现余额已保留在长期股权投资账面价值中。",
            )
        )

    checklist = [
        "确认对被投资方是重大影响或共同控制（联营/合营），而非控制（子公司）。",
        "确认交易方向为投资方 → 联营/合营企业（顺流），且交易不构成业务。",
        "只按持股比例抵销未实现损益，不按 100% 抵销收入成本。",
        "未实现的判断时点：联营企业是否已对外部独立第三方出售或耗用（含折旧）。",
        "减值迹象导致的内部交易损失不得抵销。",
        "个别报表：借投资收益 / 贷长期股权投资——损益调整。",
        "合并报表（销售当年）：借营业收入 / 贷营业成本、投资收益（重分类，不改变净利润）。",
        "逆流交易不要套用本套分录：逆流在合并报表应贷存货等资产，同时借长期股权投资。",
    ]
    if tax_rate:
        checklist.append("评估是否就递延的未实现损益确认递延所得税资产，并关注长期股权投资相关暂时性差异的确认条件。")

    return DownstreamResult(
        asset_type=asset_type,
        ownership=ownership,
        selling_price=selling_price,
        cost=cost,
        gross_profit=gross_profit,
        remaining_ratio=remaining_ratio,
        realized_ratio=realized_ratio,
        associate_book_ni=associate_book_ni,
        excess_depreciation=excess_depreciation,
        opening_unrealized_gross_profit=opening_unrealized,
        unrealized_gross_profit=unrealized_gp,
        investor_share_unrealized=investor_share,
        period_investor_deferral=period_investor_deferral,
        share_of_book_ni=share_of_book_ni,
        individual_investment_income=individual_ii,
        individual_lt_investment_adj=individual_lt_adj,
        deferred_tax_asset=deferred_tax,
        deferred_tax_expense=deferred_tax_expense,
        tax_rate=tax_rate,
        sale_year=sale_year,
        individual_entries=individual_entries,
        consolidation_entries=consolidation_entries,
        narrative=narrative,
        checklist=checklist,
    )


def combined_individual_entry(result: DownstreamResult) -> Optional[JournalEntry]:
    """将“确认份额 + 抵销未实现”合并为一笔权益法分录。"""
    if abs(result.individual_investment_income) < 0.005:
        return None
    if result.individual_investment_income >= 0:
        return _entry(
            title="权益法确认投资收益（已扣未实现顺流损益）",
            books="个别财务报表",
            lines=[
                JournalLine("长期股权投资——损益调整", debit=result.individual_investment_income),
                JournalLine("投资收益", credit=result.individual_investment_income),
            ],
            explanation=(
                f"联营账面净利润份额 {fmt_amount(result.share_of_book_ni)}"
                f" − 本期未实现变动×持股比例 {fmt_amount(result.period_investor_deferral)}"
            ),
        )
    amount = abs(result.individual_investment_income)
    return _entry(
        title="权益法确认投资损失（已扣未实现顺流损益）",
        books="个别财务报表",
        lines=[
            JournalLine("投资收益", debit=amount),
            JournalLine("长期股权投资——损益调整", credit=amount),
        ],
    )


def textbook_example_7_12_unsold() -> DownstreamResult:
    """CPA 教材典型例题：存货全部未售。"""
    return compute_downstream(
        ownership=0.20,
        selling_price=1000,
        cost=600,
        remaining_ratio=1.0,
        associate_book_ni=2000,
        asset_type="inventory",
        sale_year=True,
    )


def textbook_example_7_12_sold_70() -> DownstreamResult:
    return compute_downstream(
        ownership=0.20,
        selling_price=1000,
        cost=600,
        remaining_ratio=0.30,
        associate_book_ni=2000,
        asset_type="inventory",
        sale_year=True,
    )


def textbook_example_7_12_all_sold_same_year() -> DownstreamResult:
    return compute_downstream(
        ownership=0.20,
        selling_price=1000,
        cost=600,
        remaining_ratio=0.0,
        associate_book_ni=2000,
        asset_type="inventory",
        sale_year=True,
    )


def textbook_example_ppe_year1() -> DownstreamResult:
    """存货售予联营后作为固定资产：售价 900，成本 600，寿命 10 年，当年使用 1 个月。"""
    return compute_downstream(
        ownership=0.20,
        selling_price=900,
        cost=600,
        remaining_ratio=1.0,
        associate_book_ni=1000,
        asset_type="ppe",
        useful_life_years=10,
        months_used_in_year=1,
        sale_year=True,
    )


def print_result(result: DownstreamResult) -> None:
    print("=" * 64)
    print("销售至联营公司顺流交易抵消 · 核算结果")
    print("=" * 64)
    for line in result.narrative:
        print(f"- {line}")
    print()
    print(f"联营账面净利润份额：{fmt_amount(result.share_of_book_ni)}")
    print(f"本期递延投资方份额：{fmt_amount(result.period_investor_deferral)}")
    print(f"个别报表投资收益：{fmt_amount(result.individual_investment_income)}")
    print(f"期末仍递延余额：{fmt_amount(result.investor_share_unrealized)}")
    if result.deferred_tax_asset:
        print(f"递延所得税资产余额：{fmt_amount(result.deferred_tax_asset)}")
    print()
    for entry in result.all_entries():
        print(entry.as_text())
        print()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="销售至联营公司顺流交易抵消核算")
    parser.add_argument("--ownership", type=float, required=True, help="持股比例，如 0.2")
    parser.add_argument("--price", type=float, required=True, help="售价")
    parser.add_argument("--cost", type=float, required=True, help="账面成本")
    parser.add_argument("--remain", type=float, default=1.0, help="期末未对外出售比例")
    parser.add_argument("--ni", type=float, default=0.0, help="联营企业账面净利润")
    parser.add_argument("--asset", choices=["inventory", "ppe"], default="inventory")
    parser.add_argument("--life", type=float, default=0.0, help="固定资产预计使用寿命（年）")
    parser.add_argument("--months", type=float, default=0.0, help="当年已使用月数")
    parser.add_argument("--cum-months", type=float, default=None, help="累计已使用月数")
    parser.add_argument("--opening", type=float, default=None, help="期初未实现内部毛利")
    parser.add_argument("--tax", type=float, default=0.0, help="所得税税率")
    parser.add_argument("--subsequent", action="store_true", help="以后年度（非销售当年）")
    parser.add_argument("--impairment", action="store_true", help="内部损失属于资产减值")
    parser.add_argument("--business", action="store_true", help="交易构成业务")
    args = parser.parse_args()
    out = compute_downstream(
        ownership=args.ownership,
        selling_price=args.price,
        cost=args.cost,
        remaining_ratio=args.remain,
        associate_book_ni=args.ni,
        asset_type=args.asset,
        useful_life_years=args.life,
        months_used_in_year=args.months,
        cumulative_months_used=args.cum_months,
        opening_unrealized_gross_profit=args.opening,
        tax_rate=args.tax,
        sale_year=not args.subsequent,
        is_impairment_loss=args.impairment,
        constitutes_business=args.business,
    )
    print_result(out)

#!/usr/bin/env python3
"""生成《销售至联营公司顺流交易抵消——会计核算》Word / PDF / HTML。"""

from __future__ import annotations

import html
import shutil
from pathlib import Path
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from downstream_elimination import (  # noqa: E402
    DownstreamResult,
    JournalEntry,
    combined_individual_entry,
    compute_downstream,
    fmt_amount,
    textbook_example_7_12_all_sold_same_year,
    textbook_example_7_12_sold_70,
    textbook_example_7_12_unsold,
    textbook_example_ppe_year1,
)

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
NAVY = RGBColor(0x1A, 0x36, 0x5D)
ACCENT = RGBColor(0xB8, 0x86, 0x0B)
TEAL = colors.HexColor("#1a365d")
GOLD = colors.HexColor("#b8860b")
PAPER = colors.HexColor("#f6f1e7")
INK = colors.HexColor("#1f2933")


def entry_text(entry: JournalEntry) -> str:
    return entry.as_text()


def examples() -> dict[str, DownstreamResult]:
    y2 = compute_downstream(
        ownership=0.20,
        selling_price=1000,
        cost=600,
        remaining_ratio=0.0,
        associate_book_ni=3600,
        sale_year=False,
        opening_unrealized_gross_profit=400,
    )
    ppe2 = compute_downstream(
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
    return {
        "unsold": textbook_example_7_12_unsold(),
        "sold70": textbook_example_7_12_sold_70(),
        "sold100": textbook_example_7_12_all_sold_same_year(),
        "year2": y2,
        "ppe1": textbook_example_ppe_year1(),
        "ppe2": ppe2,
    }


def set_run_font(run, size=11, bold=False, color=None, name="WenQuanYi Micro Hei"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def shade_cell(cell, fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def add_para(doc, text, *, size=11, bold=False, color=None, space_after=8, align=None, first_line=True):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.35
    if first_line and align is None:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_heading_cn(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(text)
    set_run_font(run, size=16 if level == 1 else 13, bold=True, color=NAVY)
    return p


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run_font(run, size=9, bold=True, color=RGBColor(255, 255, 255))
        shade_cell(cell, "1A365D")
    for r_idx, row in enumerate(rows):
        fill = "F7F3EA" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            set_run_font(run, size=9)
            shade_cell(cell, fill)
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def add_entry_box(doc, entry: JournalEntry):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run(entry_text(entry))
    set_run_font(run, size=10, color=RGBColor(0x33, 0x41, 0x55))
    p.paragraph_format.space_after = Pt(8)


def build_docx(ex: dict[str, DownstreamResult], path: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cap.add_run("会计核算技术备忘录")
    set_run_font(r, size=12, bold=True, color=ACCENT)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("销售至联营公司顺流交易抵消")
    set_run_font(r, size=22, bold=True, color=NAVY)

    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run("个别财务报表权益法调整 · 合并财务报表重分类 · 连续编制")
    set_run_font(r, size=11, color=RGBColor(0x55, 0x64, 0x74))

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run("依据 CAS 2 / CAS 33 / 会计准则委员会 2023-01-05 问答 / IAS 28    日期：2026年9月21日")
    set_run_font(r, size=9, color=RGBColor(0x71, 0x80, 0x96))

    add_heading_cn(doc, "一、事项概述")
    add_para(
        doc,
        "投资方向联营企业或合营企业出售资产，称为顺流交易。联营、合营企业不纳入合并范围，不能像母子公司那样全额抵销内部收入、成本和存货。会计核算的核心是：投资方个别报表已按售价确认全部利润，但从“投资方与其联营构成的经济主体”看，只有归属于联营其他股东的部分真正实现；归属于投资方持股比例的未实现内部交易损益必须递延，待联营企业对外部独立第三方出售、耗用或计提折旧后再确认。",
    )
    add_para(
        doc,
        "本备忘录给出可直接入账的分录、合并工作底稿调整、连续年度处理，以及存货、固定资产、所得税、减值与“构成业务”等例外。",
    )

    add_heading_cn(doc, "二、准则依据")
    add_para(doc, "《企业会计准则第2号——长期股权投资》（财会〔2014〕14号）：投资方与联营、合营企业之间未实现内部交易损益，按应享有比例抵销后确认投资损益。属于资产减值损失的内部交易损失，应当全额确认。", first_line=True)
    add_para(doc, "《企业会计准则第33号——合并财务报表》（财会〔2014〕10号）：编制合并财务报表时，应在个别报表抵销的基础上，对顺流未实现收入、成本或资产处置损益中归属于投资方的部分再作列报调整。")
    add_para(doc, "财政部会计准则委员会 2023年1月5日问答：顺流交易在合并报表中抵销有关未实现的收入和成本（或资产处置损益）中归属于投资方的部分，并相应调整投资收益；逆流交易则抵销资产账面价值中包含的未实现损益，并调整长期股权投资。属于减值的内部损失，合并报表亦不得抵销。")
    add_para(doc, "IAS 28 第28段：投资者与联营或合营企业之间的顺流、逆流交易产生的损益，仅确认归属于其他投资者的部分。未实现损失若表明资产减值，则不应抵销。构成业务的资产转让，处置损益原则上全额确认。")

    add_heading_cn(doc, "三、核心原则")
    add_table(
        doc,
        ["原则", "含义"],
        [
            ["只抵“自己的那一份”", "应抵销金额 = 未实现内部毛利 × 持股比例，不是 100%。"],
            ["其他股东份额立即实现", "售给联营，相当于部分售给外部股东，该部分毛利可确认。"],
            ["未实现看联营是否出手", "存货对外出售、固定资产折旧/摊销，内部损益才逐步实现。"],
            ["个表改投资收益，合表改列报", "个别报表借投资收益、贷长期股权投资；合并报表销售当年重分类收入成本。"],
            ["减值与构成业务例外", "减值损失不抵销；转让构成业务的，损益原则上全额确认。"],
        ],
        col_widths=[4.2, 12.0],
    )

    add_heading_cn(doc, "四、核算路径")
    add_para(doc, "第一步，投资方个别报表按销售准则确认对联营的全部收入、成本和毛利（这是对外销售，法律主体层面成立）。", first_line=True)
    add_para(doc, "第二步，按权益法确认应享联营企业账面净利润份额：借“长期股权投资——损益调整”，贷“投资收益”。联营企业按购入成本入账，其账面净利润通常不含投资方的顺流毛利。")
    add_para(doc, "第三步，计算期末未实现内部毛利，乘持股比例，递延归属于投资方的部分：借“投资收益”，贷“长期股权投资——损益调整”。以后年度随实现转回。")
    add_para(doc, "第四步，若投资方另有子公司需编合并报表，销售当年在工作底稿将第三步记入投资收益的影响，重分类为冲减营业收入、营业成本（或资产处置收益）。该步不改变合并净利润。")

    add_heading_cn(doc, "五、个别财务报表核算")
    add_para(doc, "关键公式：", first_line=False)
    add_para(doc, "未实现内部毛利（存货）＝（售价 − 成本）× 期末未对外部出售比例", first_line=False, bold=True, size=10)
    add_para(doc, "本期应递延（或转回）＝（期末未实现内部毛利 − 期初未实现内部毛利）× 持股比例", first_line=False, bold=True, size=10)
    add_para(doc, "本期投资收益＝联营账面净利润 × 持股比例 − 本期应递延", first_line=False, bold=True, size=10)
    add_para(doc, "固定资产则把“未对外出售比例”换成折旧实现：每年因内部加价多提的折旧＝内部毛利 / 预计使用寿命（不考虑残值时），未实现余额＝内部毛利 − 累计多提折旧。")

    u = ex["unsold"]
    add_para(doc, "标准分录（可合并为一笔）：", first_line=False)
    add_entry_box(doc, u.individual_entries[0])
    add_entry_box(doc, u.individual_entries[1])
    comb = combined_individual_entry(u)
    if comb:
        add_para(doc, "合并为一笔：", first_line=False)
        add_entry_box(doc, comb)

    add_heading_cn(doc, "六、合并财务报表调整")
    add_para(
        doc,
        "联营企业不进合并报表，其存货无法直接冲减。销售当年，准则要求在个别报表已递延的基础上，把投资方多确认的收入、成本按持股比例抵销，并把个别报表已冲减的投资收益加回。无论联营当年对外出售多少，销售当年合并调整分录形式相同：",
    )
    add_entry_box(doc, u.consolidation_entries[0])
    add_para(
        doc,
        "该分录借贷平衡，对合并净利润的影响为零，只改变列报：营业收入、营业成本减少，投资收益增加。以后年度商品已在上期销售，合并报表一般不再重做收入成本抵销；未实现余额继续体现在长期股权投资账面价值中，待实现年度通过个别报表投资收益转回。",
    )

    add_heading_cn(doc, "七、完整例题（存货）")
    add_para(doc, "甲公司持有联营乙公司 20% 股权，能够施加重大影响。20×7 年甲公司将账面价值 600 万元的商品以 1,000 万元出售给乙公司。乙公司 20×7 年账面净利润 2,000 万元。不考虑所得税。单位：万元。", first_line=True)

    add_heading_cn(doc, "（一）年末商品全部未对外部出售", 2)
    add_table(
        doc,
        ["项目", "计算", "金额"],
        [
            ["顺流毛利", "1,000 − 600", "400"],
            ["应递延（投资方份额）", "400 × 20%", "80"],
            ["可立即确认（其他股东）", "400 × 80%", "320"],
            ["联营净利润份额", "2,000 × 20%", "400"],
            ["个别报表投资收益", "400 − 80", "320"],
            ["合并重分类：营业收入", "1,000 × 20%", "200"],
            ["合并重分类：营业成本", "600 × 20%", "120"],
            ["合并重分类：投资收益", "400 × 20%", "80"],
        ],
        col_widths=[5.5, 6.5, 4.2],
    )
    add_para(doc, "个别报表投资收益 320；合并报表投资收益列报 = 320 + 80 = 400，同时营业收入、营业成本各减 200、120。合并净利润仍为个别报表净利润。", first_line=True)

    add_heading_cn(doc, "（二）年末已对外出售 70%，30% 形成存货", 2)
    s70 = ex["sold70"]
    add_para(doc, f"未实现毛利 = 400 × 30% = 120，应递延 = 24。个别报表投资收益 = {fmt_amount(s70.individual_investment_income)}。销售当年合并分录仍为：借营业收入 200，贷营业成本 120、投资收益 80。", first_line=True)
    add_entry_box(doc, combined_individual_entry(s70))
    add_entry_box(doc, s70.consolidation_entries[0])

    add_heading_cn(doc, "（三）当年全部对外部出售", 2)
    s100 = ex["sold100"]
    add_para(doc, f"未实现为零，个别报表投资收益 = {fmt_amount(s100.individual_investment_income)}（即 2,000 × 20%）。合并报表仍按持股比例重分类收入成本，贷记投资收益 80，使集团口径下对乙公司投资收益反映“乙公司若按甲公司成本结转”的份额。", first_line=True)

    add_heading_cn(doc, "（四）次年全部对外出售（连续编制）", 2)
    y2 = ex["year2"]
    add_para(doc, "设 20×8 年乙公司账面净利润 3,600 万元，并将上年内部商品全部售出。个别报表应把上年递延的 80 转回：投资收益 = 3,600 × 20% + 80 = 800。合并报表无收入成本调整分录。", first_line=True)
    add_entry_box(doc, combined_individual_entry(y2))

    add_heading_cn(doc, "八、固定资产顺流交易")
    p1 = ex["ppe1"]
    add_para(doc, "甲公司 20×7 年 11 月将成本 600 万元的商品以 900 万元售给乙公司，乙公司作为管理用固定资产，寿命 10 年、净残值 0、年限平均法。乙公司当年净利润 1,000 万元。内部毛利 300 万元，当年多提折旧 300 / 10 / 12 = 2.5 万元。", first_line=True)
    add_para(doc, f"期末未实现 = 300 − 2.5 = 297.5；应递延 = 59.5；个别报表投资收益 = {fmt_amount(p1.individual_investment_income)}。", first_line=True)
    add_entry_box(doc, combined_individual_entry(p1))
    add_entry_box(doc, p1.consolidation_entries[0])
    add_para(doc, "若投资方原列报为固定资产处置，合并报表借方改为“资产处置收益”60，不再拆收入成本。次年按全年多提折旧 30 转回实现：本期转回投资方份额 6，个别报表投资收益 = 1,000 × 20% + 6 = 206；合并报表以后年度不再重分类收入成本。", first_line=True)

    add_heading_cn(doc, "九、顺流、逆流、子公司内部交易对照")
    add_table(
        doc,
        ["项目", "顺流（售给联营）", "逆流（联营售给投资方）", "售给子公司"],
        [
            ["未实现利润存在于", "联营企业存货/固定资产（不进合表）", "投资方存货/固定资产（进合表）", "购货方存货/固定资产"],
            ["抵销比例", "持股比例", "持股比例", "全额 100%"],
            ["个别报表", "借投资收益 / 贷长期股权投资", "同左", "一般不作此笔（成本法或内部未抵）"],
            ["合并报表", "借营业收入 / 贷营业成本、投资收益", "借长期股权投资 / 贷存货等", "借营业收入 / 贷营业成本、存货"],
            ["少数股东", "不涉及 NCI（联营不合并）", "不涉及 NCI", "未实现损益在母与 NCI 之间分摊"],
        ],
    )

    add_heading_cn(doc, "十、例外与特殊事项")
    add_heading_cn(doc, "（一）减值损失不抵销", 2)
    add_para(doc, "顺流售价低于成本且表明资产发生减值的，内部损失全额确认，不得按持股比例递延。理由：损失已经实现，递延会高估资产和利润。")
    add_heading_cn(doc, "（二）构成业务", 2)
    add_para(doc, "投出或出售的资产构成业务的，相关损益原则上全额确认，不适用未实现损益递延。判断应结合《企业会计准则第20号——企业合并》中“业务”的定义，并关注分步交易是否应作为一揽子交易。")
    add_heading_cn(doc, "（三）抵销额超过长期股权投资账面价值", 2)
    add_para(doc, "IAS 28 相关项目曾讨论：顺流未实现收益的抵销额若大于长期股权投资账面价值，超额部分应列报为递延收益（负债），而不是把投资冲成负数后停止抵销。中国实务宜谨慎：先将长期股权投资冲至零，超额部分确认为递延收益，待联营产生利润或资产对外部实现时再释放。")
    add_heading_cn(doc, "（四）所得税", 2)
    add_para(doc, "投资方计税时通常已就全部销售利润纳税，会计上却递延了持股比例部分，形成可抵扣暂时性差异。常见分录：借“递延所得税资产”（应递延金额 × 税率），贷“所得税费用”。实现年度转回。确认前需评估暂时性差异转回的可能性，以及 CAS 18 / IAS 12 对联营投资相关暂时性差异的限制条件。")

    add_heading_cn(doc, "十一、工作底稿")
    add_table(
        doc,
        ["工作底稿栏", "取数", "注意"],
        [
            ["内部销售收入", "投资方对联营营业收入", "按客户明细或关联方披露核对"],
            ["内部销售成本", "对应结转营业成本", "含运费资本化等调整"],
            ["期末联营持有比例", "联营存货进销存 / 固定资产台账", "须区分已对第三方出售"],
            ["持股比例", "考虑稀释、潜在表决权后的权益法比例", "与损益分配比例一致"],
            ["联营账面净利润", "联营当期净利润", "先按投资时点公允价值调整"],
            ["期初未实现余额", "上年工作底稿期末数", "连续编制的起点"],
            ["税率", "投资方适用税率", "亏损或税收优惠时可能不确认 DTA"],
        ],
    )

    add_heading_cn(doc, "十二、检查清单与常见错误")
    for item in u.checklist:
        add_para(doc, "□ " + item, first_line=False, size=10)
    add_para(doc, "常见错误：", first_line=False, bold=True)
    add_para(doc, "1. 按 100% 抵销对联营的收入成本，误把联营当子公司。", first_line=False)
    add_para(doc, "2. 顺流、逆流分录用反：顺流合表不冲存货，逆流合表不冲投资方收入。", first_line=False)
    add_para(doc, "3. 个别报表已递延后，合并报表又贷记长期股权投资，造成重复抵销。", first_line=False)
    add_para(doc, "4. 忽略以后年度转回，未实现损益被永久挂在长期股权投资。", first_line=False)
    add_para(doc, "5. 固定资产只抵初始毛利、不调折旧实现。", first_line=False)
    add_para(doc, "6. 把减值损失当成普通未实现损失予以递延。", first_line=False)

    add_heading_cn(doc, "十三、分录速查")
    add_table(
        doc,
        ["时点", "个别报表", "合并工作底稿"],
        [
            ["销售当年，存货未售完", "借：投资收益；贷：长期股权投资（未实现×持股）", "借：营业收入（售价×持股）；贷：营业成本、投资收益"],
            ["销售当年，存货已售完", "只确认联营净利润份额，不再递延", "仍按售价/成本×持股重分类收入成本"],
            ["以后年度存货售出", "借：长期股权投资；贷：投资收益（转回）", "一般无分录"],
            ["固定资产销售当年", "递延（毛利−当年多提折旧）×持股", "借收入/处置收益，贷成本与投资收益（毛利×持股）"],
            ["固定资产以后年度", "按当年多提折旧×持股转回", "一般无分录"],
        ],
    )

    note = doc.add_paragraph()
    note.paragraph_format.first_line_indent = Cm(0)
    r = note.add_run("本备忘录供会计核算与合并报表编制使用，不构成审计意见。金额单位除特别注明外为万元。配套计算：python3 scripts/downstream_elimination.py --ownership 0.2 --price 1000 --cost 600 --remain 1 --ni 2000")
    set_run_font(r, size=9, color=RGBColor(0x71, 0x80, 0x96))

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def register_pdf_font() -> str:
    pdfmetrics.registerFont(TTFont("WQ", FONT_PATH, subfontIndex=0))
    return "WQ"


def pdf_styles(font: str):
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CoverKicker", fontName=font, fontSize=11, textColor=GOLD, alignment=TA_CENTER, spaceAfter=8))
    styles.add(ParagraphStyle(name="CoverTitle", fontName=font, fontSize=20, leading=28, textColor=TEAL, alignment=TA_CENTER, spaceAfter=10))
    styles.add(ParagraphStyle(name="CoverSub", fontName=font, fontSize=10, leading=16, textColor=colors.HexColor("#556474"), alignment=TA_CENTER, spaceAfter=6))
    styles.add(ParagraphStyle(name="H1CN", fontName=font, fontSize=13, leading=20, textColor=TEAL, spaceBefore=14, spaceAfter=6))
    styles.add(ParagraphStyle(name="H2CN", fontName=font, fontSize=11, leading=16, textColor=colors.HexColor("#2c5282"), spaceBefore=10, spaceAfter=4))
    styles.add(ParagraphStyle(name="BodyCN", fontName=font, fontSize=9.5, leading=16, textColor=INK, alignment=TA_JUSTIFY, firstLineIndent=18, spaceAfter=6))
    styles.add(ParagraphStyle(name="BodyLeft", fontName=font, fontSize=9.5, leading=16, textColor=INK, alignment=TA_LEFT, spaceAfter=4))
    styles.add(ParagraphStyle(name="Entry", fontName=font, fontSize=8.5, leading=13, textColor=colors.HexColor("#334155"), backColor=colors.HexColor("#f4efe4"), leftIndent=8, rightIndent=8, spaceBefore=4, spaceAfter=6))
    styles.add(ParagraphStyle(name="Th", fontName=font, fontSize=8, leading=12, textColor=colors.white, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="Td", fontName=font, fontSize=8, leading=12, textColor=INK))
    styles.add(ParagraphStyle(name="FooterCN", fontName=font, fontSize=8, textColor=colors.HexColor("#718096"), alignment=TA_CENTER))
    return styles


def pdf_table(headers, rows, styles, widths):
    data = [[Paragraph(h, styles["Th"]) for h in headers]]
    for row in rows:
        data.append([Paragraph(str(c), styles["Td"]) for c in row])
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), TEAL),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "WQ"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#d6d3cd")),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f7f3ea"), colors.white]),
            ]
        )
    )
    return table


def build_pdf(ex: dict[str, DownstreamResult], path: Path) -> None:
    font = register_pdf_font()
    styles = pdf_styles(font)
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.6 * cm,
        title="销售至联营公司顺流交易抵消——会计核算",
        author="会计核算技术备忘录",
    )
    story = []
    story.append(Paragraph("会计核算技术备忘录", styles["CoverKicker"]))
    story.append(Paragraph("销售至联营公司顺流交易抵消", styles["CoverTitle"]))
    story.append(Paragraph("个别财务报表权益法调整 · 合并财务报表重分类 · 连续编制", styles["CoverSub"]))
    story.append(Paragraph("CAS 2 / CAS 33 / 会计准则委员会 2023-01-05 问答 / IAS 28　·　2026年9月21日", styles["CoverSub"]))
    story.append(Spacer(1, 8))

    def h1(t):
        story.append(Paragraph(t, styles["H1CN"]))

    def h2(t):
        story.append(Paragraph(t, styles["H2CN"]))

    def p(t):
        story.append(Paragraph(html.escape(t), styles["BodyCN"]))

    def box(entry: JournalEntry):
        story.append(Preformatted(entry_text(entry), styles["Entry"]))

    h1("一、事项概述")
    p("投资方向联营企业或合营企业出售资产，称为顺流交易。联营、合营企业不纳入合并范围，不能像母子公司那样全额抵销内部收入、成本和存货。会计核算的核心是：投资方个别报表已按售价确认全部利润，但从投资方与其联营构成的经济主体看，只有归属于联营其他股东的部分真正实现；归属于投资方持股比例的未实现内部交易损益必须递延，待联营企业对外部独立第三方出售、耗用或计提折旧后再确认。")

    h1("二、准则依据")
    p("CAS 2：未实现内部交易损益按应享有比例抵销后确认投资损益；属于资产减值损失的内部损失全额确认。")
    p("CAS 33 及会计准则委员会 2023-01-05 问答：合并报表在个别报表抵销基础上，对顺流未实现收入和成本（或资产处置损益）中归属于投资方的部分予以抵销并调整投资收益；逆流则抵销资产账面价值中的未实现损益并调整长期股权投资。")
    p("IAS 28.28：仅确认归属于其他投资者的损益；减值损失不抵销；构成业务的转让原则上全额确认损益。")

    h1("三、核心原则")
    story.append(
        pdf_table(
            ["原则", "含义"],
            [
                ["只抵自己的那一份", "应抵销金额 = 未实现内部毛利 × 持股比例，不是 100%。"],
                ["其他股东份额立即实现", "售给联营相当于部分售给外部股东，该部分毛利可确认。"],
                ["未实现看联营是否出手", "存货对外出售或固定资产折旧后，内部损益才逐步实现。"],
                ["个表改投资收益，合表改列报", "个别报表借投资收益、贷长期股权投资；合并报表销售当年重分类收入成本。"],
                ["减值与构成业务例外", "减值损失不抵销；构成业务的转让原则上全额确认损益。"],
            ],
            styles,
            [4.2 * cm, 12.3 * cm],
        )
    )

    h1("四、关键公式")
    p("未实现内部毛利（存货）＝（售价 − 成本）× 期末未对外部出售比例。")
    p("本期应递延（负数即转回）＝（期末未实现 − 期初未实现）× 持股比例。")
    p("本期投资收益＝联营账面净利润 × 持股比例 − 本期应递延。")
    p("固定资产：每年多提折旧＝内部毛利 / 使用寿命；未实现余额＝内部毛利 − 累计多提折旧。")

    u = ex["unsold"]
    h1("五、个别报表与合并报表分录（教材典型例）")
    p("甲持有乙 20%。甲将成本 600 的商品以 1,000 售给乙，乙年末未对外部出售。乙当年净利润 2,000。单位：万元。")
    box(u.individual_entries[0])
    box(u.individual_entries[1])
    comb = combined_individual_entry(u)
    if comb:
        p("亦可合并为一笔：借长期股权投资 320，贷投资收益 320。")
    p("合并工作底稿（销售当年，无论对外出售多少，形式相同）：")
    box(u.consolidation_entries[0])
    p("合并净利润不变。列报变为：营业收入、营业成本按持股比例减少，投资收益加回毛利×持股比例。")

    h1("六、存货三种情形与次年转回")
    story.append(
        pdf_table(
            ["情形", "未实现毛利", "应递延", "个别报表投资收益", "合并重分类"],
            [
                ["年末全部未售", "400", "80", "320", "借收入 200 / 贷成本 120、投资收益 80"],
                ["对外出售 70%", "120", "24", "376", "分录形式同上"],
                ["当年全部外售", "0", "0", "400", "分录形式同上"],
                ["次年全部外售（乙净利润 3,600）", "0", "转回 80", "800", "无收入成本分录"],
            ],
            styles,
            [4.0 * cm, 2.2 * cm, 2.2 * cm, 3.4 * cm, 4.7 * cm],
        )
    )

    h1("七、固定资产")
    p("11 月售出，售价 900、成本 600，乙作为固定资产，寿命 10 年。当年多提折旧 2.5。个别报表投资收益 =（1,000 − 297.5）× 20% = 140.5。合并报表借营业收入 180，贷营业成本 120、投资收益 60。次年全年多提折旧 30，转回 6，投资收益 = 206。")

    h1("八、对照表")
    story.append(
        pdf_table(
            ["项目", "顺流（售给联营）", "逆流", "售给子公司"],
            [
                ["未实现存在于", "联营资产（不进合表）", "投资方资产（进合表）", "购货方资产"],
                ["抵销比例", "持股比例", "持股比例", "100%"],
                ["合并调整", "借收入 / 贷成本、投资收益", "借长期股权投资 / 贷存货", "借收入 / 贷成本、存货"],
            ],
            styles,
            [3.2 * cm, 4.8 * cm, 4.3 * cm, 4.2 * cm],
        )
    )

    h1("九、例外")
    p("减值损失不抵销。构成业务的资产转让，损益原则上全额确认。抵销额超过长期股权投资账面价值时，宜将投资冲至零，超额列报递延收益。所得税方面，会计递延而税法已征税的部分，通常确认递延所得税资产（应递延金额 × 税率），实现时转回。")

    h1("十、检查清单")
    for item in u.checklist:
        story.append(Paragraph("□ " + html.escape(item), styles["BodyLeft"]))
    p("切勿：按 100% 抵销联营收入；顺流合表去冲存货；个表已抵后再贷长期股权投资造成重复；忘记以后年度转回；固定资产不调折旧；把减值损失递延。")
    p("命令行核算：python3 scripts/downstream_elimination.py --ownership 0.2 --price 1000 --cost 600 --remain 1 --ni 2000")

    def on_page(canvas, doc_):
        canvas.saveState()
        canvas.setFillColor(TEAL)
        canvas.rect(0, A4[1] - 12, A4[0], 12, fill=1, stroke=0)
        canvas.setFillColor(GOLD)
        canvas.rect(0, 0, A4[0], 8, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#718096"))
        canvas.setFont("WQ", 8)
        canvas.drawString(1.8 * cm, 1.0 * cm, "销售至联营公司顺流交易抵消 · 会计核算")
        canvas.drawRightString(A4[0] - 1.8 * cm, 1.0 * cm, f"{doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)


def esc(text: str) -> str:
    return html.escape(text)


def entry_html(entry: JournalEntry) -> str:
    return f'<pre class="entry">{esc(entry_text(entry))}</pre>'


def build_html(ex: dict[str, DownstreamResult], path: Path) -> None:
    u, s70, s100, y2, p1 = ex["unsold"], ex["sold70"], ex["sold100"], ex["year2"], ex["ppe1"]
    html_doc = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>销售至联营公司顺流交易抵消——会计核算</title>
<style>
:root {{
  --navy:#1a365d; --gold:#b8860b; --ink:#1f2933; --muted:#5c6b7a;
  --paper:#fbf7f0; --card:#fffdf8; --line:#e4d9c5;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0; font-family:"Source Han Serif SC","Noto Serif SC","Songti SC","WenQuanYi Micro Hei",serif;
  color:var(--ink); background:linear-gradient(180deg,#eef2f6 0%, #f7f3ea 220px, var(--paper) 220px);
  line-height:1.65;
}}
.wrap {{ max-width:980px; margin:0 auto; padding:32px 20px 80px; }}
header.hero {{
  background:var(--navy); color:#f8f4ea; padding:36px 32px; border-radius:18px;
  box-shadow:0 12px 40px rgba(26,54,93,.22);
}}
header.hero .kicker {{ letter-spacing:.18em; font-size:12px; color:#e4c56a; }}
header.hero h1 {{ font-size:30px; margin:10px 0 8px; line-height:1.3; }}
header.hero p {{ margin:0; color:#d7dee8; font-size:15px; }}
nav.toc {{
  display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:8px;
  margin:22px 0 28px;
}}
nav.toc a {{
  display:block; background:var(--card); border:1px solid var(--line); border-radius:10px;
  padding:10px 12px; color:var(--navy); text-decoration:none; font-size:13px;
}}
nav.toc a:hover {{ border-color:var(--gold); }}
section {{
  background:var(--card); border:1px solid var(--line); border-radius:16px;
  padding:22px 24px; margin:0 0 16px;
}}
h2 {{ color:var(--navy); font-size:20px; margin:0 0 12px; }}
h3 {{ color:#2c5282; font-size:16px; margin:18px 0 8px; }}
p {{ margin:0 0 10px; text-align:justify; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; margin:10px 0 14px; }}
th {{ background:var(--navy); color:#fff; padding:8px; text-align:left; font-weight:600; }}
td {{ border-bottom:1px solid var(--line); padding:8px; vertical-align:top; }}
tr:nth-child(even) td {{ background:#f7f3ea; }}
pre.entry {{
  background:#1a365d; color:#f6e7c1; padding:12px 14px; border-radius:10px;
  font-family:ui-monospace,"Sarasa Mono SC",monospace; font-size:13px; overflow:auto;
  white-space:pre-wrap;
}}
.formula {{
  background:#fff8e8; border-left:4px solid var(--gold); padding:10px 14px; margin:10px 0;
  font-family:ui-monospace,monospace; font-size:13px;
}}
.calc {{
  display:grid; grid-template-columns:1fr 1fr; gap:16px; align-items:start;
}}
label {{ display:block; font-size:12px; color:var(--muted); margin:8px 0 4px; }}
input, select {{
  width:100%; padding:8px 10px; border:1px solid var(--line); border-radius:8px;
  font-size:14px; background:#fff;
}}
button {{
  margin-top:14px; background:var(--navy); color:#fff; border:0; border-radius:10px;
  padding:10px 16px; font-size:14px; cursor:pointer;
}}
button:hover {{ background:#234876; }}
#calc-out {{ background:#1a365d; color:#f6e7c1; min-height:180px; padding:14px; border-radius:10px;
  white-space:pre-wrap; font-family:ui-monospace,monospace; font-size:13px; }}
.kpi {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:8px; margin:10px 0 14px; }}
.kpi div {{ background:#f7f3ea; border-radius:10px; padding:10px; }}
.kpi b {{ display:block; color:var(--navy); font-size:18px; }}
.kpi span {{ font-size:12px; color:var(--muted); }}
footer {{ text-align:center; color:var(--muted); font-size:12px; margin-top:24px; }}
@media (max-width:800px) {{ .calc {{ grid-template-columns:1fr; }} header.hero h1 {{ font-size:24px; }} }}
@media print {{
  body {{ background:#fff; }}
  .calc-card {{ break-inside:avoid; }}
  header.hero {{ box-shadow:none; }}
  nav.toc {{ display:none; }}
}}
</style>
</head>
<body>
<div class="wrap">
<header class="hero">
  <div class="kicker">会计核算技术备忘录　TM-EQ-2026-001</div>
  <h1>销售至联营公司顺流交易抵消</h1>
  <p>个别财务报表权益法调整 · 合并财务报表重分类 · 连续编制<br/>
  CAS 2 / CAS 33 / 会计准则委员会 2023-01-05 问答 / IAS 28　·　2026年9月21日</p>
</header>
<nav class="toc">
  <a href="#s1">1. 事项与原则</a>
  <a href="#s2">2. 核算路径与公式</a>
  <a href="#s3">3. 个别报表 / 合并报表</a>
  <a href="#s4">4. 存货例题</a>
  <a href="#s5">5. 固定资产</a>
  <a href="#s6">6. 对照与例外</a>
  <a href="#s7">7. 在线核算器</a>
</nav>

<section id="s1">
<h2>一、事项概述与核心原则</h2>
<p>投资方向联营企业或合营企业出售资产，称为<strong>顺流交易</strong>。联营、合营企业不纳入合并范围，不能像母子公司那样按 100% 抵销内部收入、成本和存货。投资方法律主体报表已确认全部销售利润，但从“投资方与其联营”这一经济主体看，只有归属于联营<strong>其他股东</strong>的部分真正实现；归属于投资方持股比例的未实现内部交易损益必须递延，待联营企业对外部独立第三方出售、耗用或计提折旧后再确认。</p>
<table>
<tr><th>原则</th><th>含义</th></tr>
<tr><td>只抵自己的那一份</td><td>应抵销金额 = 未实现内部毛利 × 持股比例，不是 100%。</td></tr>
<tr><td>其他股东份额立即实现</td><td>售给联营，相当于部分售给外部股东，该部分毛利可确认。</td></tr>
<tr><td>未实现看联营是否出手</td><td>存货对外出售、固定资产折旧后，内部损益才逐步实现。</td></tr>
<tr><td>个表改投资收益，合表改列报</td><td>个别报表借投资收益、贷长期股权投资；合并报表销售当年重分类收入成本，<strong>不改变净利润</strong>。</td></tr>
<tr><td>减值与构成业务例外</td><td>减值损失不抵销；构成业务的转让原则上全额确认损益。</td></tr>
</table>
<p>准则依据：CAS 2（财会〔2014〕14号）、CAS 33（财会〔2014〕10号）、财政部会计准则委员会 2023-01-05 问答、IAS 28 第 28 段。</p>
</section>

<section id="s2">
<h2>二、核算路径与公式</h2>
<p>① 个别报表按销售准则确认对联营的全部收入、成本、毛利。② 权益法确认应享联营账面净利润份额。③ 按持股比例递延（或转回）未实现内部损益。④ 若需编制合并报表，销售当年把记入投资收益的影响重分类为冲减营业收入与营业成本（或资产处置收益）。</p>
<div class="formula">未实现内部毛利（存货）＝（售价 − 成本）× 期末未对外部出售比例</div>
<div class="formula">本期应递延（负数=转回）＝（期末未实现 − 期初未实现）× 持股比例</div>
<div class="formula">本期投资收益＝联营账面净利润 × 持股比例 − 本期应递延</div>
<div class="formula">固定资产未实现余额＝内部毛利 − 累计多提折旧；年多提折旧＝内部毛利 / 使用寿命</div>
</section>

<section id="s3">
<h2>三、个别报表与合并报表</h2>
<p>教材典型例：持股 20%，售价 1,000，成本 600，乙年末未售出，乙净利润 2,000（万元）。个别报表投资收益 {fmt_amount(u.individual_investment_income)}。</p>
{entry_html(u.individual_entries[0])}
{entry_html(u.individual_entries[1])}
<p>合并工作底稿——销售当年无论对外出售多少，分录形式相同；这是重分类，合并净利润不变：</p>
{entry_html(u.consolidation_entries[0])}
<p>切勿在合并报表再次贷记长期股权投资，否则重复抵销。以后年度一般不再重做收入成本分录。</p>
</section>

<section id="s4">
<h2>四、存货例题</h2>
<div class="kpi">
  <div><span>全部未售·投资收益</span><b>{fmt_amount(u.individual_investment_income)}</b></div>
  <div><span>外售 70%·投资收益</span><b>{fmt_amount(s70.individual_investment_income)}</b></div>
  <div><span>当年售尽·投资收益</span><b>{fmt_amount(s100.individual_investment_income)}</b></div>
  <div><span>次年售尽·投资收益</span><b>{fmt_amount(y2.individual_investment_income)}</b></div>
</div>
<table>
<tr><th>情形</th><th>未实现毛利</th><th>应递延</th><th>个别报表投资收益</th><th>合并重分类</th></tr>
<tr><td>年末全部未售</td><td>400</td><td>80</td><td>320</td><td>借收入 200 / 贷成本 120、投资收益 80</td></tr>
<tr><td>对外出售 70%</td><td>120</td><td>24</td><td>376</td><td>分录形式同上</td></tr>
<tr><td>当年全部外售</td><td>0</td><td>0</td><td>400</td><td>分录形式同上</td></tr>
<tr><td>次年全部外售（乙净利润 3,600）</td><td>0</td><td>转回 80</td><td>800</td><td>无收入成本分录</td></tr>
</table>
<h3>70% 外售时个别报表（合并一笔）</h3>
{entry_html(combined_individual_entry(s70))}
<h3>次年转回</h3>
{entry_html(combined_individual_entry(y2))}
</section>

<section id="s5">
<h2>五、固定资产顺流交易</h2>
<p>11 月将成本 600 的商品以 900 售给乙，乙作管理用固定资产，寿命 10 年、残值 0。乙净利润 1,000。内部毛利 300，当年多提折旧 2.5。期末未实现 297.5，应递延 59.5，投资收益 <strong>{fmt_amount(p1.individual_investment_income)}</strong>。</p>
{entry_html(combined_individual_entry(p1))}
{entry_html(p1.consolidation_entries[0])}
<p>若投资方原列报固定资产处置，合并借方改为“资产处置收益”60。次年全年多提折旧 30，转回投资方份额 6，投资收益 = 206；合并报表以后年度不再重分类收入成本。</p>
</section>

<section id="s6">
<h2>六、对照、例外与检查清单</h2>
<table>
<tr><th>项目</th><th>顺流（售给联营）</th><th>逆流</th><th>售给子公司</th></tr>
<tr><td>未实现存在于</td><td>联营资产（不进合表）</td><td>投资方资产（进合表）</td><td>购货方资产</td></tr>
<tr><td>抵销比例</td><td>持股比例</td><td>持股比例</td><td>100%</td></tr>
<tr><td>合并调整</td><td>借收入 / 贷成本、投资收益</td><td>借长期股权投资 / 贷存货</td><td>借收入 / 贷成本、存货</td></tr>
</table>
<p><strong>减值：</strong>售价低于成本且表明资产减值的，损失全额确认，不抵销。<strong>构成业务：</strong>损益原则上全额确认。<strong>超额抵销：</strong>宜将长期股权投资冲至零，超额列报递延收益。<strong>所得税：</strong>借递延所得税资产（应递延 × 税率），贷所得税费用；实现时转回。</p>
<ul>
{''.join(f'<li>{esc(item)}</li>' for item in u.checklist)}
</ul>
<p>常见错误：按 100% 抵联营收入；顺流合表去冲存货；个表已抵后再贷长期股权投资；忘记以后年度转回；固定资产不调折旧；把减值损失递延。</p>
</section>

<section id="s7" class="calc-card">
<h2>七、在线核算器</h2>
<p>输入本企业数据，即时生成个别报表与合并报表分录。单位与输入保持一致。</p>
<div class="calc">
  <form id="f">
    <label>持股比例（如 0.2）</label><input id="ownership" type="number" step="0.01" min="0.01" max="1" value="0.2" />
    <label>售价</label><input id="price" type="number" step="0.01" value="1000" />
    <label>成本</label><input id="cost" type="number" step="0.01" value="600" />
    <label>期末未对外出售比例（存货）</label><input id="remain" type="number" step="0.01" min="0" max="1" value="1" />
    <label>联营企业账面净利润</label><input id="ni" type="number" step="0.01" value="2000" />
    <label>资产类型</label>
    <select id="asset"><option value="inventory">存货</option><option value="ppe">固定资产</option></select>
    <label>固定资产寿命（年）</label><input id="life" type="number" step="0.5" value="10" />
    <label>当年使用月数</label><input id="months" type="number" step="1" value="1" />
    <label>累计使用月数（以后年度）</label><input id="cum" type="number" step="1" value="" placeholder="销售当年可空" />
    <label>期初未实现毛利（以后年度存货）</label><input id="opening" type="number" step="0.01" value="" placeholder="销售当年为 0" />
    <label>所得税税率</label><input id="tax" type="number" step="0.01" min="0" max="1" value="0" />
    <label>期间</label>
    <select id="period"><option value="sale">销售当年</option><option value="later">以后年度</option></select>
    <label>例外</label>
    <select id="exception"><option value="">无</option><option value="impairment">减值损失不抵销</option><option value="business">构成业务，全额确认</option></select>
    <button type="submit">生成分录</button>
  </form>
  <div id="calc-out">填写左侧数据后点击生成分录。</div>
</div>
</section>
<footer>本备忘录供会计核算与合并报表编制使用，不构成审计意见。命令行：python3 scripts/downstream_elimination.py --ownership 0.2 --price 1000 --cost 600 --remain 1 --ni 2000</footer>
</div>
<script>
function r2(v){{ return Math.round((Number(v)+Number.EPSILON)*100)/100; }}
function fmt(v){{ return r2(v).toLocaleString("zh-CN", {{minimumFractionDigits:2, maximumFractionDigits:2}}); }}
function compute(p){{
  const own = Number(p.ownership), price = Number(p.price), cost = Number(p.cost);
  const remain = Number(p.remain), ni = Number(p.ni), tax = Number(p.tax)||0;
  const saleYear = p.period === "sale";
  let gp = r2(price - cost);
  let opening = p.opening === "" || p.opening === null ? (saleYear ? 0 : gp) : Number(p.opening);
  let unreal = 0, excess = 0;
  if (p.exception === "business" || p.exception === "impairment") {{
    unreal = 0; opening = 0;
  }} else if (p.asset === "ppe") {{
    const life = Number(p.life)||10;
    const months = saleYear ? (Number(p.months)||0) : (Number(p.months)||12);
    let cum = p.cum === "" || p.cum === null ? months : Number(p.cum);
    const annual = r2(gp / life);
    excess = r2(annual * (months/12));
    const cumEx = r2(annual * (cum/12));
    unreal = r2(Math.max(gp - cumEx, 0));
    if (p.opening === "" || p.opening === null) {{
      const openEx = r2(Math.max(cumEx - excess, 0));
      opening = saleYear ? 0 : r2(Math.max(gp - openEx, 0));
    }}
  }} else {{
    unreal = r2(gp * remain);
    if (p.opening === "" || p.opening === null) opening = saleYear ? 0 : gp;
  }}
  const investorShare = r2(unreal * own);
  const periodDef = r2((unreal - opening) * own);
  const shareNI = r2(ni * own);
  const ii = r2(shareNI - periodDef);
  const dtaEnd = tax && investorShare>0 ? r2(investorShare * tax) : 0;
  const dtaOpen = tax && opening>0 ? r2(r2(opening*own)*tax) : 0;
  const dtaMove = r2(dtaEnd - dtaOpen);
  let lines = [];
  lines.push("毛利 " + fmt(gp) + "　期末未实现 " + fmt(unreal) + "　本期递延（投资方） " + fmt(periodDef));
  lines.push("联营净利润份额 " + fmt(shareNI) + "　个别报表投资收益 " + fmt(ii));
  if (excess) lines.push("本期多提折旧 " + fmt(excess));
  lines.push("");
  lines.push("【个别财务报表】");
  if (shareNI) {{
    lines.push("  借：长期股权投资——损益调整  " + fmt(shareNI));
    lines.push("    贷：投资收益  " + fmt(shareNI));
  }}
  if (periodDef > 0.005) {{
    lines.push("  借：投资收益  " + fmt(periodDef));
    lines.push("    贷：长期股权投资——损益调整  " + fmt(periodDef));
  }} else if (periodDef < -0.005) {{
    lines.push("  借：长期股权投资——损益调整  " + fmt(Math.abs(periodDef)));
    lines.push("    贷：投资收益  " + fmt(Math.abs(periodDef)));
  }}
  if (dtaMove > 0.005) {{
    lines.push("  借：递延所得税资产  " + fmt(dtaMove));
    lines.push("    贷：所得税费用  " + fmt(dtaMove));
  }} else if (dtaMove < -0.005) {{
    lines.push("  借：所得税费用  " + fmt(Math.abs(dtaMove)));
    lines.push("    贷：递延所得税资产  " + fmt(Math.abs(dtaMove)));
  }}
  lines.push("");
  lines.push("【合并财务报表工作底稿】");
  if (saleYear && p.exception === "") {{
    const rev = r2(price*own), cogs = r2(cost*own), plug = r2(rev-cogs);
    lines.push("  借：营业收入  " + fmt(rev));
    lines.push("    贷：营业成本  " + fmt(cogs));
    lines.push("    贷：投资收益  " + fmt(plug));
    lines.push("  （重分类，不改变合并净利润）");
  }} else {{
    lines.push("  以后年度或例外事项：通常无需收入成本重分类。");
  }}
  return lines.join("\\n");
}}
document.getElementById("f").addEventListener("submit", function(e){{
  e.preventDefault();
  const ids = ["ownership","price","cost","remain","ni","asset","life","months","cum","opening","tax","period","exception"];
  const p = {{}};
  ids.forEach(id => p[id] = document.getElementById(id).value);
  document.getElementById("calc-out").textContent = compute(p);
}});
</script>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html_doc, encoding="utf-8")


def main() -> None:
    REPORTS.mkdir(parents=True, exist_ok=True)
    ex = examples()
    html_path = REPORTS / "销售至联营公司顺流交易抵消-会计核算.html"
    docx_path = REPORTS / "销售至联营公司顺流交易抵消-会计核算.docx"
    pdf_path = REPORTS / "销售至联营公司顺流交易抵消-会计核算.pdf"
    build_html(ex, html_path)
    build_docx(ex, docx_path)
    build_pdf(ex, pdf_path)
    shutil.copyfile(html_path, REPORTS / "associate-downstream-elimination.html")
    shutil.copyfile(docx_path, REPORTS / "associate-downstream-elimination.docx")
    shutil.copyfile(pdf_path, REPORTS / "associate-downstream-elimination.pdf")
    print("Wrote:")
    for p in (html_path, docx_path, pdf_path):
        print(f"  {p}  ({p.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

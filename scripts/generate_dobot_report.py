#!/usr/bin/env python3
"""Generate 越疆科技财报分析 PDF report for iPhone download."""

from fpdf import FPDF
from pathlib import Path

FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
OUTPUT = Path("/workspace/reports/越疆科技财报与股价分析报告.pdf")


class ReportPDF(FPDF):
    def epw(self):
        return self.w - self.l_margin - self.r_margin

    def header(self):
        self.set_font("cn", "", 9)
        self.set_text_color(120, 120, 120)
        self.cell(0, 8, "越疆科技（02432.HK）财报与股价分析报告", align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("cn", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"第 {self.page_no()} 页 | 数据截至 2026年7月 | 仅供研究参考，不构成投资建议", align="C")

    def section_title(self, title: str):
        self.ln(4)
        self.set_font("cn", "B", 14)
        self.set_text_color(25, 55, 95)
        self.multi_cell(self.epw(), 8, title)
        self.ln(2)

    def sub_title(self, title: str):
        self.set_font("cn", "B", 11)
        self.set_text_color(40, 40, 40)
        self.multi_cell(self.epw(), 7, title)
        self.ln(1)

    def body(self, text: str):
        self.set_font("cn", "", 10)
        self.set_text_color(30, 30, 30)
        self.multi_cell(self.epw(), 6, text)
        self.ln(2)

    def bullet(self, text: str):
        self.set_font("cn", "", 10)
        self.set_text_color(30, 30, 30)
        self.multi_cell(self.epw(), 6, f"• {text}")
        self.ln(1)

    def table_row(self, cols, widths, bold=False):
        style = "B" if bold else ""
        self.set_font("cn", style, 9)
        for i, (col, w) in enumerate(zip(cols, widths)):
            self.cell(w, 7, col, border=1, align="C")
        self.ln()


def build_report():
    pdf = ReportPDF()
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_font("cn", "", FONT_PATH)
    pdf.add_font("cn", "B", FONT_PATH)
    pdf.add_page()
    w = pdf.epw()

    # Cover
    pdf.ln(20)
    pdf.set_font("cn", "B", 22)
    pdf.set_text_color(25, 55, 95)
    pdf.multi_cell(w, 12, "越疆科技（02432.HK）", align="C")
    pdf.set_font("cn", "B", 18)
    pdf.multi_cell(w, 10, "财报分析与股价预测报告", align="C")
    pdf.ln(8)
    pdf.set_font("cn", "", 11)
    pdf.set_text_color(80, 80, 80)
    pdf.multi_cell(w, 7, "深圳市越疆科技股份有限公司", align="C")
    pdf.multi_cell(w, 7, "协作机器人第一股 · 具身智能平台", align="C")
    pdf.ln(12)
    pdf.set_font("cn", "", 10)
    pdf.multi_cell(w, 6, "报告日期：2026年7月30日", align="C")
    pdf.multi_cell(w, 6, "数据来源：港交所年报、业绩预告、券商研报、公开市场数据", align="C")
    pdf.ln(20)
    pdf.set_font("cn", "", 9)
    pdf.set_text_color(150, 50, 50)
    pdf.multi_cell(w, 6, "免责声明：本报告基于公开信息整理，仅供研究参考，不构成任何投资建议。投资有风险，决策需谨慎。", align="C")

  # Part 1: Company overview
    pdf.add_page()
    pdf.section_title("一、公司概况")
    pdf.body(
        "越疆科技（02432.HK）成立于2015年，总部位于深圳，2024年12月23日在港交所主板上市，"
        "是「协作机器人第一股」及18C章特专科技公司。主营业务为协作机器人与具身智能机器人"
        "（机械臂、人形、多足）的研发、生产与销售。"
    )
    pdf.sub_title("基本信息")
    widths = [45, 125]
    pdf.table_row(["项目", "内容"], widths, bold=True)
    rows = [
        ("股票代码", "02432.HK"),
        ("上市时间", "2024年12月23日"),
        ("发行价", "18.8 港元/股"),
        ("创始人", "刘培超"),
        ("行业地位", "2023年协作机器人出货量全球第二、中国第一"),
        ("累计出货", "突破10万台（2025年）"),
        ("全球化", "业务覆盖100+国家和地区"),
    ]
    for r in rows:
        pdf.table_row(r, widths)

    pdf.sub_title("战略定位")
    pdf.bullet("双轮驱动：协作机器人智能化升级 + 具身智能机器人创新")
    pdf.bullet("全形态产品矩阵：机械臂 + 人形机器人 + 多足机器人")
    pdf.bullet("全栈自研：本体、力控、感知、端云协同、具身大模型")

    # Part 2: Financial performance
    pdf.section_title("二、财务业绩分析（2021–2025）")
    pdf.sub_title("营收与毛利")
    widths = [28, 28, 28, 28, 28, 28]
    pdf.table_row(["年度", "营收(亿)", "增速", "毛利(亿)", "毛利率", "净亏损(亿)"], widths, bold=True)
    fin_data = [
        ("2021", "1.74", "—", "0.88", "50.6%", "0.42"),
        ("2022", "2.41", "+38.5%", "0.98", "40.7%", "0.52"),
        ("2023", "2.87", "+19.0%", "1.25", "43.5%", "1.03"),
        ("2024", "3.74", "+30.3%", "1.74", "46.6%", "0.95"),
        ("2025", "4.92", "+31.7%", "2.27", "46.1%", "0.84"),
    ]
    for row in fin_data:
        pdf.table_row(row, widths)

    pdf.body(
        "近五年营收复合增速约28%，毛利率稳定在46%左右。2025年净亏损0.84亿元，同比收窄11.9%；"
        "但经调整净亏损0.50亿元，同比扩大36.1%，反映剔除上市费用后经营性亏损仍在加大。"
    )

    pdf.sub_title("2025年产品结构")
    widths = [50, 35, 25, 35]
    pdf.table_row(["产品类别", "收入(亿元)", "占比", "同比增速"], widths, bold=True)
    products = [
        ("六轴协作机器人", "3.02", "61.4%", "+44.7%"),
        ("四轴协作机器人", "0.93", "18.9%", "-3.0%"),
        ("复合机器人", "0.68", "13.7%", "+27.3%"),
        ("具身智能机器人", "0.20", "4.1%", "+418.8%"),
    ]
    for row in products:
        pdf.table_row(row, widths)

    pdf.sub_title("应用场景分布（2025）")
    widths = [60, 40, 40]
    pdf.table_row(["场景", "收入(亿元)", "同比增速"], widths, bold=True)
    for row in [("工业", "2.79", "+39.4%"), ("教育", "1.67", "+13.7%"), ("商业", "0.45", "+75.7%")]:
        pdf.table_row(row, widths)

    # Part 3: Cost structure
    pdf.section_title("三、费用结构与盈利能力")
    widths = [40, 35, 35, 40]
    pdf.table_row(["费用项", "2025(亿)", "同比增速", "占营收比"], widths, bold=True)
    for row in [
        ("研发费用", "1.15", "+59.7%", "23.3%"),
        ("销售及经销", "1.82", "+32.1%", "37.0%"),
        ("行政费用", "0.73", "-17.6%", "14.8%"),
    ]:
        pdf.table_row(row, widths)

    pdf.body(
        "三项费用合计约3.70亿元，远超2.27亿元毛利，是持续亏损的直接原因。"
        "2025年研发投入1.15亿元，具身智能占研发总额39.3%；员工从560人增至768人。"
    )

    # Part 4: Balance sheet & cash flow
    pdf.section_title("四、资产负债与现金流")
    widths = [55, 45, 45]
    pdf.table_row(["项目", "2025年末", "2024年末"], widths, bold=True)
    for row in [
        ("总资产(亿元)", "31.03", "14.95"),
        ("股东权益(亿元)", "26.23", "9.67"),
        ("现金及等价物(亿元)", "5.80", "8.84"),
        ("银行借款(亿元)", "0.72", "2.18"),
        ("资本负债比率", "15.5%", "35.3%"),
        ("经营现金流净额(万元)", "-4,259", "-9,168"),
    ]:
        pdf.table_row(row, widths)

    pdf.body(
        "股东权益大增主要来自2025年两次配售融资约17.93亿港元。经营现金流仍为负但改善54%，"
        "主营业务尚未自我造血，依赖外部融资，但现金储备充裕。"
    )

    # Part 5: 2026 H1 preview
    pdf.section_title("五、2026年上半年业绩预告")
    pdf.body("公司于2026年7月14日发布未经审计营运数据预告：")
    widths = [55, 45, 45]
    pdf.table_row(["指标", "2026H1预告", "同比变化"], widths, bold=True)
    for row in [
        ("营业收入", "3.0–3.3亿元", "+94.7%至+114.1%"),
        ("毛利额", "1.4–1.7亿元", "+84.7%至+124.3%"),
        ("归母净利润", "-0.9至-1.2亿元", "亏损扩大"),
        ("经调整净利润*", "-0.35至-0.65亿元", "剔除汇兑及股份支付"),
    ]:
        pdf.table_row(row, widths)
    pdf.body("*经调整净利润剔除汇兑损失与股份支付影响。亏损扩大主因战略性投入加大及汇兑损失。")

    pdf.sub_title("重要进展")
    pdf.bullet("2026年7月22日：创业板IPO获上市委审议通过（大湾区首单H回A）")
    pdf.bullet("协作机器人在手订单超1亿元；具身智能在手订单超6000万元")
    pdf.bullet("2026H1具身智能出货金额超4000万元")

    # Part 6: Valuation & stock price
    pdf.add_page()
    pdf.section_title("六、估值与股价分析")
    pdf.sub_title("当前定价（2026年7月下旬）")
    widths = [50, 90]
    pdf.table_row(["指标", "数值"], widths, bold=True)
    for row in [
        ("股价", "约24.8–25.8港元"),
        ("总市值", "约109–113亿港元"),
        ("52周区间", "22.52 – 64.50港元"),
        ("市销率PS(2025)", "约22倍"),
        ("市净率PB", "约3.9倍"),
        ("分析师平均目标价", "59.5港元（10家机构）"),
        ("目标价区间", "50 – 69港元"),
    ]:
        pdf.table_row(row, widths)

    pdf.body("股价自64.50港元高点回落约60%，估值已从高位明显压缩。")

    pdf.sub_title("营收预测（综合）")
    widths = [35, 45, 70]
    pdf.table_row(["年度", "预测营收(亿元)", "来源/假设"], widths, bold=True)
    for row in [
        ("2025(实际)", "4.92", "已披露"),
        ("2026E", "6.3–7.5", "H1预告+光大证券7.5亿"),
        ("2027E", "10.8", "光大证券预测"),
        ("2028E", "15.3", "光大证券；公司预计扭亏"),
    ]:
        pdf.table_row(row, widths)

    pdf.sub_title("可比公司估值")
    widths = [40, 40, 40, 40]
    pdf.table_row(["公司", "市值(亿港元)", "2025营收(亿)", "隐含PS"], widths, bold=True)
    for row in [
        ("优必选09880", "~430", "~20", "~20x"),
        ("越疆02432", "~110", "4.92", "~22x"),
        ("极智嘉02590", "~122", "—", "—"),
    ]:
        pdf.table_row(row, widths)

    # Part 7: Price forecast
    pdf.section_title("七、股价预测（情景分析）")
    pdf.body("基于市销率（PS）估值法，越疆尚未盈利，市场主要采用PS估值。")

    pdf.sub_title("分时段预测")
    widths = [35, 35, 70]
    pdf.table_row(["时间维度", "基准预测", "预测区间"], widths, bold=True)
    for row in [
        ("3个月(2026Q4)", "30港元", "25 – 38港元"),
        ("6个月(2027Q1)", "38港元", "28 – 45港元"),
        ("12个月(2027中)", "48港元", "35 – 58港元"),
        ("18个月(2027底)", "55港元", "40 – 65港元"),
        ("2028年(扭亏年)", "58港元", "45 – 75港元"),
    ]:
        pdf.table_row(row, widths)

    pdf.sub_title("情景估值对照")
    widths = [30, 35, 30, 35]
    pdf.table_row(["情景", "营收基准", "PS倍数", "目标价"], widths, bold=True)
    for row in [
        ("悲观", "2026E 6.5亿", "10x", "17港元"),
        ("保守", "2026E 7.5亿", "12x", "22港元"),
        ("中性", "2027E 10.8亿", "15x", "40港元"),
        ("乐观", "2027E 10.8亿", "20x", "53港元"),
        ("激进", "2028E 15.3亿", "22x", "83港元"),
    ]:
        pdf.table_row(row, widths)

    pdf.body(
        "基准情景下，12个月目标价约48港元，较现价有约90%上行空间。"
        "分析师59.5港元目标价对应2027年营收约22–24倍PS。"
    )

    # Part 8: Catalysts and risks
    pdf.section_title("八、催化剂与风险提示")
    pdf.sub_title("看多因素")
    pdf.bullet("2026H1营收预告翻倍增长，商业化能力验证")
    pdf.bullet("协作机器人出货量全球第一，累计超10万台")
    pdf.bullet("创业板过会（86天），H回A稀缺标的溢价")
    pdf.bullet("现金5.8亿+融资充裕，无流动性风险")
    pdf.bullet("10家机构80%给予买入评级")

    pdf.sub_title("看空因素")
    pdf.bullet("经调整净亏损扩大，2026H1亏损9000万–1.2亿")
    pdf.bullet("研发费用2026年预计超2亿元（同比翻倍）")
    pdf.bullet("股价已从高点缩水60%，板块情绪低迷")
    pdf.bullet("协作机器人价格战持续，毛利率承压")
    pdf.bullet("港股机器人板块整体回调（优必选从161跌至85港元）")

    pdf.sub_title("关键观察节点")
    pdf.bullet("2026年8–9月：H1正式财报披露")
    pdf.bullet("2026年Q4：A股IPO注册与发行定价")
    pdf.bullet("2026年底：全年营收能否达7亿+、具身智能破1亿")
    pdf.bullet("2027年：期间费用率是否下降、亏损是否收窄")
    pdf.bullet("2028年：能否如期扭亏（存在延后至2029年风险）")

    pdf.section_title("九、综合结论")
    pdf.body(
        "越疆科技是一家高增长、高投入、尚未盈利的协作机器人龙头，正向具身智能平台转型。"
        "2025年营收增长31.7%、具身智能收入暴涨418%，但研发投入激增59.7%导致经调整亏损扩大。"
        "资金储备充裕，短期无流动性风险，但盈利拐点取决于具身智能商业化进度与费用管控能力。"
    )
    pdf.body(
        "股价预测：基准情景12个月目标价约48港元（区间35–58港元）。"
        "若协作机器人增速低于预期10%或具身智能商业化慢于预期20%，"
        "扭亏可能推迟至2029年，股价或回落至30–40港元区间。"
    )

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(OUTPUT))
    return OUTPUT


if __name__ == "__main__":
    path = build_report()
    print(f"Generated: {path} ({path.stat().st_size / 1024:.1f} KB)")

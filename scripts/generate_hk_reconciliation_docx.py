#!/usr/bin/env python3
"""Generate Hong Kong-style bilingual Statement of Account / Reconciliation Letter (.docx).

Layout is tuned so the full template typically fits on two A4 pages when printed
(default margins, PMingLiU/Times New Roman, normal zoom).
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def set_cell_shading(cell, fill_hex: str) -> None:
    from docx.oxml import OxmlElement

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def add_compact_paragraph(
    doc: Document,
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    size_pt: float = 10,
    western: str = "PMingLiU",
    align=None,
) -> None:
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.08
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size_pt)
    r.font.name = western
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")


def tighten_document(doc: Document) -> None:
    """Reduce spacing on all paragraphs (including inside tables)."""
    for p in doc.paragraphs:
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(2)
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = 1.08

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                    pf.line_spacing = 1.05


def set_table_body_font(table, size_pt: float = 9) -> None:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(size_pt)
                    r.font.name = "PMingLiU"
                    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")


def main() -> None:
    out = Path("/workspace/HK_Statement_of_Account_Reconciliation_Template.docx")
    doc = Document()

    sec = doc.sections[0]
    sec.top_margin = Cm(1.7)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "PMingLiU"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(2)
    style.paragraph_format.space_before = Pt(0)

    # Title (compact — avoid built-in Heading 0 large spacing)
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_after = Pt(4)
    t.paragraph_format.space_before = Pt(0)
    rt = t.add_run("對賬函")
    rt.bold = True
    rt.font.size = Pt(15)
    rt.font.name = "PMingLiU"
    rt._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")

    st = doc.add_paragraph()
    st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.space_before = Pt(0)
    rs = st.add_run("Statement of Account / Reconciliation Letter")
    rs.bold = True
    rs.font.size = Pt(10.5)
    rs.font.name = "Times New Roman"
    rs._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")

    meta = doc.add_table(rows=5, cols=2)
    rows_meta = [
        ("Our Ref. 本函編號", "AR-REC-2026-001（請修改）"),
        ("Date 日期", "2026年5月26日（26 May 2026）"),
        ("To 致", "（客戶公司中英文全名）"),
        ("Attn. 聯絡人", "（姓名及職銜）"),
        ("Address 地址", "（客戶註冊或通訊地址）"),
    ]
    for i, (a, b) in enumerate(rows_meta):
        meta.rows[i].cells[0].text = a
        meta.rows[i].cells[1].text = b
        set_cell_shading(meta.rows[i].cells[0], "E7E6E6")

    add_compact_paragraph(doc, "Re: 主旨　Subject", bold=True, size_pt=10)
    add_compact_paragraph(
        doc,
        "Account Reconciliation as at [日期]　截至[日期]之帳目核對",
        size_pt=9.5,
    )

    add_compact_paragraph(doc, "Dear Sirs / Madam　敬啟者：", bold=True, size_pt=10)

    add_compact_paragraph(
        doc,
        "本公司為核對雙方帳目及確保記錄一致，謹將截至 ____年____月____日 止貴我雙方往來帳項之結餘列示如下，敬請查核。"
        "如與貴司帳冊記錄相符，請於下方「客戶確認」欄簽署及蓋上公司印章（或授權簽章）後，於 ____年____月____日或之前傳真／電郵／郵寄回本司，以便存檔。",
        size_pt=9.5,
    )
    pe = doc.add_paragraph()
    pe.paragraph_format.space_before = Pt(0)
    pe.paragraph_format.space_after = Pt(2)
    pe.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pe.paragraph_format.line_spacing = 1.08
    re = pe.add_run(
        "For the purpose of reconciling our records, please find below the balance of our account with you as at [date]. "
        "Kindly verify the figures against your books. If the balance agrees with your records, please sign and chop "
        "the confirmation section below and return the same to us by fax / email / post on or before [date]."
    )
    re.font.name = "Times New Roman"
    re.font.size = Pt(9)
    re.italic = True

    add_compact_paragraph(
        doc,
        "若有不符或須調整事項，請於上述期限前以書面列明差異及依據（例如發票／收據／付款憑證編號），以便雙方盡快處理。",
        size_pt=9.5,
    )
    pe2 = doc.add_paragraph()
    pe2.paragraph_format.space_before = Pt(0)
    pe2.paragraph_format.space_after = Pt(2)
    pe2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pe2.paragraph_format.line_spacing = 1.08
    re2 = pe2.add_run(
        "Should there be any discrepancy, please notify us in writing before the aforesaid date, stating the nature of "
        "the difference and supporting particulars (e.g. invoice / receipt / payment reference)."
    )
    re2.font.name = "Times New Roman"
    re2.font.size = Pt(9)
    re2.italic = True

    add_compact_paragraph(doc, "帳目結餘摘要　Account Summary", bold=True, size_pt=10)

    t_sum = doc.add_table(rows=3, cols=2)
    t_sum.rows[0].cells[0].text = "說明 Description"
    t_sum.rows[0].cells[1].text = "金額（港幣 HKD）Amount (HKD)"
    t_sum.rows[1].cells[0].text = "截至上述日期之應收／應付結餘（刪去不適用）Balance as at [date] (AR / AP)"
    t_sum.rows[1].cells[1].text = "$ ________________"
    t_sum.rows[2].cells[0].text = "其中：已開立但未到期款項（如適用）Including amounts not yet due (if any)"
    t_sum.rows[2].cells[1].text = "$ ________________"
    for c in t_sum.rows[0].cells:
        set_cell_shading(c, "D9D9D9")

    add_compact_paragraph(
        doc,
        "明細附表（可另附附件）　Detailed listing — may attach schedule",
        bold=True,
        size_pt=10,
    )

    # Header + 1 blank line + total (compact)
    t_det = doc.add_table(rows=3, cols=7)
    hdr = [
        "日期 Date",
        "類別 Type",
        "編號 Ref.",
        "摘要 Particulars",
        "借方 Dr",
        "貸方 Cr",
        "結餘 Bal.",
    ]
    for j, text in enumerate(hdr):
        t_det.rows[0].cells[j].text = text
        set_cell_shading(t_det.rows[0].cells[j], "D9D9D9")
    for j in range(7):
        t_det.rows[1].cells[j].text = ""
    t_det.rows[2].cells[0].merge(t_det.rows[2].cells[5])
    t_det.rows[2].cells[0].text = "結餘 Total"
    t_det.rows[2].cells[6].text = "$ ________________"
    set_cell_shading(t_det.rows[2].cells[0], "F2F2F2")

    add_compact_paragraph(
        doc,
        "註：預設港幣 HKD；若合約約定其他貨幣請改列並註明。",
        size_pt=8,
    )

    add_compact_paragraph(
        doc,
        "（可選）法律／私隱提示　(Optional) Legal and Privacy Notice",
        bold=True,
        size_pt=9.5,
    )
    add_compact_paragraph(
        doc,
        "1）「視為確認」類條款可能具法律效果，採用前請內部法務／律師審閱。2）聯絡資料僅作對帳用途，並按香港《個人資料（私隱）條例》處理。",
        size_pt=8.5,
    )
    pe3 = doc.add_paragraph()
    pe3.paragraph_format.space_before = Pt(0)
    pe3.paragraph_format.space_after = Pt(2)
    pe3.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pe3.paragraph_format.line_spacing = 1.08
    re3 = pe3.add_run(
        "Note: (1) Any “deemed acceptance” or similar wording may have legal effect — obtain internal legal / lawyer "
        "review before use. (2) Contact details herein are for reconciliation purposes only and will be handled in "
        "accordance with the Personal Data (Privacy) Ordinance (Cap. 486) of Hong Kong."
    )
    re3.font.name = "Times New Roman"
    re3.font.size = Pt(8)
    re3.italic = True

    add_compact_paragraph(doc, "發函公司資料　Our Company Particulars", bold=True, size_pt=10)

    # Merged contact row to save vertical space
    co = doc.add_table(rows=5, cols=2)
    co_rows = [
        ("公司中文名稱", ""),
        ("英文名稱 English name", ""),
        ("註冊／營業地址 Address", "（香港地址）"),
        ("BR 商業登記證號碼 / CR 公司註冊編號（如適用）", "（例：BR …；CR …）"),
        ("聯絡 Tel／Fax／Email", ""),
    ]
    for i, (a, b) in enumerate(co_rows):
        co.rows[i].cells[0].text = a
        co.rows[i].cells[1].text = b
        set_cell_shading(co.rows[i].cells[0], "E7E6E6")

    add_compact_paragraph(doc, "客戶確認（回函用）　Customer Confirmation", bold=True, size_pt=10)

    add_compact_paragraph(
        doc,
        "茲確認截至 ____年____月____日，上述結餘 HKD $ ________________ 與敝公司帳冊記錄　相符 □　不符 □（不符請另紙說明）。",
        size_pt=9.5,
    )
    pe4 = doc.add_paragraph()
    pe4.paragraph_format.space_before = Pt(0)
    pe4.paragraph_format.space_after = Pt(2)
    pe4.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pe4.paragraph_format.line_spacing = 1.08
    re4 = pe4.add_run(
        "We confirm that, as at [date], the above balance of HKD $________ agrees □ / does not agree □ with our records "
        "(if not agreed, please attach particulars)."
    )
    re4.font.name = "Times New Roman"
    re4.font.size = Pt(9)
    re4.italic = True

    conf = doc.add_table(rows=4, cols=2)
    conf_data = [
        ("授權簽署 Authorized Signature", ""),
        ("姓名 Name／職銜 Title", ""),
        ("公司印章 Company Chop", "（蓋印處）"),
        ("日期 Date", ""),
    ]
    for i, (a, b) in enumerate(conf_data):
        conf.rows[i].cells[0].text = a
        conf.rows[i].cells[1].text = b
        set_cell_shading(conf.rows[i].cells[0], "E7E6E6")
    conf.rows[2].cells[1].text = "（蓋印處）"

    add_compact_paragraph(doc, "Yours faithfully　此致　商祺", bold=True, size_pt=10)
    add_compact_paragraph(doc, "_________________________　（簽署人姓名、職銜、公司名稱及印章）", size_pt=9.5)

    for table in doc.tables:
        set_table_body_font(table, 8.5)
    tighten_document(doc)

    doc.save(out)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

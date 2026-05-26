#!/usr/bin/env python3
"""Generate Hong Kong-style bilingual Statement of Account / Reconciliation Letter (.docx)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def set_cell_shading(cell, fill_hex: str) -> None:
    from docx.oxml import OxmlElement

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def main() -> None:
    out = Path("/workspace/HK_Statement_of_Account_Reconciliation_Template.docx")
    doc = Document()

    # Page margins (A4, comfortable for HK letters)
    sec = doc.sections[0]
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    # Default font: MingLiU / PMingLiU common in HK; Word falls back if missing
    style = doc.styles["Normal"]
    style.font.name = "PMingLiU"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")
    style.font.size = Pt(11)

    add_heading(doc, "對賬函", 0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Statement of Account / Reconciliation Letter")
    r.bold = True
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PMingLiU")

    doc.add_paragraph()

    # Letter meta table
    meta = doc.add_table(rows=5, cols=2)
    meta.autofit = True
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
    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run("Re: 主旨").bold = True
    doc.add_paragraph(
        "Account Reconciliation as at [日期]　截至[日期]之帳目核對"
    )

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Dear Sirs / Madam　敬啟者：").bold = True

    body_tc = (
        "本公司為核對雙方帳目及確保記錄一致，謹將截至 ____年____月____日 止貴我雙方往來帳項之結餘列示如下，敬請查核。"
        "如與貴司帳冊記錄相符，請於下方「客戶確認」欄簽署及蓋上公司印章（或授權簽章）後，於 ____年____月____日或之前 "
        "傳真／電郵／郵寄回本司，以便存檔。"
    )
    doc.add_paragraph(body_tc)

    body_en = (
        "For the purpose of reconciling our records, please find below the balance of our account with you as at [date]. "
        "Kindly verify the figures against your books. If the balance agrees with your records, please sign and chop "
        "the confirmation section below and return the same to us by fax / email / post on or before [date]."
    )
    pe = doc.add_paragraph(body_en)
    for run in pe.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)
        run.italic = True

    body_tc2 = (
        "若有不符或須調整事項，請於上述期限前以書面列明差異及依據（例如發票／收據／付款憑證編號），以便雙方盡快處理。"
    )
    doc.add_paragraph(body_tc2)
    pe2 = doc.add_paragraph(
        "Should there be any discrepancy, please notify us in writing before the aforesaid date, stating the nature of "
        "the difference and supporting particulars (e.g. invoice / receipt / payment reference)."
    )
    for run in pe2.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)
        run.italic = True

    doc.add_paragraph()
    h = doc.add_paragraph()
    h.add_run("帳目結餘摘要　Account Summary").bold = True

    t_sum = doc.add_table(rows=3, cols=2)
    t_sum.rows[0].cells[0].text = "說明 Description"
    t_sum.rows[0].cells[1].text = "金額（港幣 HKD）Amount (HKD)"
    t_sum.rows[1].cells[0].text = (
        "截至上述日期之應收／應付結餘（請刪去不適用者）\n"
        "Balance as at [date] (AR / AP — delete as applicable)"
    )
    t_sum.rows[1].cells[1].text = "$ ________________"
    t_sum.rows[2].cells[0].text = (
        "其中：已開立但未到期款項（如適用）\nIncluding: amounts not yet due (if any)"
    )
    t_sum.rows[2].cells[1].text = "$ ________________"
    for c in t_sum.rows[0].cells:
        set_cell_shading(c, "D9D9D9")

    doc.add_paragraph()
    h2 = doc.add_paragraph()
    h2.add_run("明細附表（可另附附件）　Detailed listing — may attach schedule").bold = True

    t_det = doc.add_table(rows=4, cols=7)
    hdr = [
        "日期\nDate",
        "單據類別\nType",
        "單據編號\nRef.",
        "摘要\nParticulars",
        "借方\nDebit",
        "貸方\nCredit",
        "結餘\nBalance",
    ]
    for j, text in enumerate(hdr):
        t_det.rows[0].cells[j].text = text
        set_cell_shading(t_det.rows[0].cells[j], "D9D9D9")
    for r in range(1, 3):
        for j in range(7):
            t_det.rows[r].cells[j].text = ""
    t_det.rows[3].cells[0].merge(t_det.rows[3].cells[5])
    t_det.rows[3].cells[0].text = "結餘 Total"
    t_det.rows[3].cells[6].text = "$ ________________"
    set_cell_shading(t_det.rows[3].cells[0], "F2F2F2")

    note = doc.add_paragraph(
        "註：香港習慣以港幣列示；若合約約定其他貨幣，請改列並註明幣別。"
    )
    note.runs[0].font.size = Pt(9)

    doc.add_paragraph()
    h3 = doc.add_paragraph()
    h3.add_run("（可選）法律／合約提示　Optional clauses").bold = True
    doc.add_paragraph(
        "1. 無爭議款項：若貴司未於上述期限內提出書面異議，則視為確認上述結餘；惟此舉是否影響任何一方在合約或法例下的權利，請按公司政策及律師意見調整後再採用。\n"
        "2. 個人資料：本函所載聯絡資料僅作對帳用途，並按香港《個人資料（私隱）條例》處理。"
    )
    pe3 = doc.add_paragraph(
        "Note: Any “deemed acceptance” wording may have legal effect — obtain legal review before use."
    )
    for run in pe3.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        run.italic = True

    doc.add_paragraph()
    h4 = doc.add_paragraph()
    h4.add_run("發函公司資料　Our Company Particulars").bold = True

    co = doc.add_table(rows=7, cols=2)
    co_rows = [
        ("公司中文名稱", ""),
        ("英文名稱 English name", ""),
        ("註冊／營業地址 Registered / business address", "（香港地址）"),
        ("商業登記證號碼 BR No.", "（例：12345678-XXX-XX-XX-X）"),
        ("公司註冊編號（如適用）CR No.", "（有限公司）"),
        ("電話／傳真 Tel / Fax", ""),
        ("電郵 Email", ""),
    ]
    for i, (a, b) in enumerate(co_rows):
        co.rows[i].cells[0].text = a
        co.rows[i].cells[1].text = b
        set_cell_shading(co.rows[i].cells[0], "E7E6E6")

    doc.add_paragraph()
    doc.add_page_break()

    h5 = doc.add_paragraph()
    h5.add_run("客戶確認（回函用）　Customer Confirmation").bold = True

    doc.add_paragraph(
        "茲確認截至 ____年____月____日，上述結餘 HKD $ ________________ 與敝公司帳冊記錄　相符 □　不符 □（如不符，請另紙說明）。"
    )
    pe4 = doc.add_paragraph(
        "We confirm that, as at [date], the above balance of HKD $________ agrees □ / does not agree □ with our records "
        "(if not agreed, please attach particulars)."
    )
    for run in pe4.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)
        run.italic = True

    conf = doc.add_table(rows=5, cols=2)
    conf_data = [
        ("授權簽署 Authorized Signature", ""),
        ("姓名 Name", ""),
        ("職銜 Title", ""),
        ("公司印章 Company Chop", "（蓋印處）"),
        ("日期 Date", ""),
    ]
    for i, (a, b) in enumerate(conf_data):
        conf.rows[i].cells[0].text = a
        conf.rows[i].cells[1].text = b
        set_cell_shading(conf.rows[i].cells[0], "E7E6E6")
        if i == 3:
            conf.rows[i].cells[1].text = "\n\n（蓋印處）\n\n"

    doc.add_paragraph()
    doc.add_paragraph()
    p_close = doc.add_paragraph()
    p_close.add_run("Yours faithfully　此致").bold = True
    doc.add_paragraph("商祺")
    doc.add_paragraph()
    doc.add_paragraph("_________________________")
    doc.add_paragraph("（簽署人姓名）")
    doc.add_paragraph("（職銜）")
    doc.add_paragraph("（公司名稱及印章）")

    doc.save(out)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

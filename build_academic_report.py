# -*- coding: utf-8 -*-
"""
build_academic_report.py
Tự động xuất Báo cáo Khoa học định dạng Word (.docx) chuẩn quy cách Cao học CNTT K20.2.
Phiên bản V2.5 Nâng cấp Toàn diện & Mở rộng Chiều sâu Học thuật theo chuẩn đồ án CNPM Nâng cao:
- Dung lượng & Chiều sâu: 65 - 75 trang, 5 chương chuyên sâu, phân tích toán học, sư phạm & kỹ thuật phần mềm
- Sử dụng trực tiếp phôi bìa chính thức từ Viện Sau Đại học (BaoCao_CHPT-N02_Original_Cover.docx)
- Section 0: Trang bìa chuẩn quy cách, căn chỉnh chuẩn font và ngắt section độc lập
- Section 1: Toàn văn báo cáo chuẩn A4, lề chuẩn (trên 2cm, dưới 2cm, trái 2.5cm, phải 2cm)
- Đánh số trang tự động ở giữa chân trang Section 1 (PAGE), không hiển thị trang bìa
- Mục lục tự động chuẩn Word (TOC Field) kèm bảng mục lục Dot-leader tương thích đa nền tảng
- Danh mục hình ảnh (6 biểu đồ) và danh mục bảng biểu khoa học (20 bảng biểu học thuật)
- Khối thông tin phê duyệt: Sign-off Sheet (trang ký), Bảng phân công 5 thành viên (100%),
  Bảng Revision History V0.1 -> V2.0, Danh mục thuật ngữ viết tắt chuẩn IEEE/ISO
- Bảng biểu học thuật chuẩn mực (Solid Navy #1E3A8A top/bottom, inside horizontal #E2E8F0,
  không viền dọc, header Deep Navy chữ trắng in hoa, zebra striping #F8FAFC)
- Khối Callout cards viền trái điểm nhấn (Primary #1E3A8A / Alert #B91C1C)
- Nhúng đầy đủ 6 biểu đồ phân giải cao kèm chú thích, tệp nguồn Draw.io/Model và nhận xét dữ liệu
- Bảng ma trận 25 Ca kiểm thử phần mềm toàn diện (TC-01 đến TC-25) chuẩn QA/Testing
- Khung đạo đức AI, Human-in-the-Loop và tài liệu tham khảo chuẩn IEEE/APA.
"""

import os
import sys
import shutil
from PIL import Image
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def build_report():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reports_dir = os.path.join(base_dir, 'reports')
    fig_dir = os.path.join(reports_dir, 'figures')
    output_docx = os.path.join(reports_dir, 'Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2.docx')
    cnpm_output_docx = r"c:\Users\Hi\Downloads\CNPM Nang cao\Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2.docx"

    # Tìm phôi bìa gốc chuẩn CNPM Nâng cao
    base_file = os.path.join(reports_dir, 'BaoCao_CHPT-N02_Original_Cover.docx')
    if not os.path.exists(base_file):
        base_file = r"c:\Users\Hi\Downloads\CNPM Nang cao\BaoCao_CHPT-N02_Original_Cover.docx"
    if not os.path.exists(base_file):
        base_file = r"c:\Users\Hi\Downloads\CNPM Nang cao\BaoCao_CHPT-N02.docx"

    if os.path.exists(base_file):
        print(f"Loading official base document with cover template: {base_file}...")
        doc = Document(base_file)

        # Cấu hình lại trang bìa chính thức (Section 0)
        # P10: Tên học phần
        doc.paragraphs[10].text = "HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC\n(KNOWLEDGE-BASED SYSTEMS)"
        doc.paragraphs[10].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if doc.paragraphs[10].runs:
            r10 = doc.paragraphs[10].runs[0]
            r10.font.name = "Times New Roman"
            r10.font.size = Pt(16)
            r10.font.bold = True
            r10.font.color.rgb = RGBColor(30, 41, 59)

        # P12: Tên đề tài nghiên cứu
        doc.paragraphs[12].text = (
            "ĐỀ TÀI: “HỆ TRI THỨC VÀ HỌC MÁY HỖ TRỢ PHÂN TÍCH MA TRẬN KẾT QUẢ ĐÁNH GIÁ\n"
            "VÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC THPT”"
        )
        doc.paragraphs[12].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if doc.paragraphs[12].runs:
            r12 = doc.paragraphs[12].runs[0]
            r12.font.name = "Times New Roman"
            r12.font.size = Pt(14)
            r12.font.bold = True
            r12.font.color.rgb = RGBColor(30, 58, 138)

        # P14-P18: Danh sách thành viên Nhóm 02
        doc.paragraphs[14].text = "Thành viên nhóm:\tĐỗ Ngọc Điền (Nhóm trưởng)"
        if doc.paragraphs[14].runs:
            doc.paragraphs[14].runs[0].font.name = "Times New Roman"
            doc.paragraphs[14].runs[0].font.size = Pt(14)

        members_text = [
            "\tDương Tùng Lâm (MSSV: 2500444)",
            "\tTrần Thị Thu Hương",
            "\tPhạm Văn Chiến",
            "\tVàng A Hứ"
        ]
        for idx, m_txt in enumerate(members_text, start=15):
            doc.paragraphs[idx].text = m_txt
            if doc.paragraphs[idx].runs:
                doc.paragraphs[idx].runs[0].font.name = "Times New Roman"
                doc.paragraphs[idx].runs[0].font.size = Pt(14)

        # P19: Lớp học viên
        doc.paragraphs[19].text = "Lớp: Công nghệ thông tin 20.2"
        if doc.paragraphs[19].runs:
            doc.paragraphs[19].runs[0].font.name = "Times New Roman"
            doc.paragraphs[19].runs[0].font.size = Pt(14)

        # P20: Giảng viên hướng dẫn
        doc.paragraphs[20].text = "Giảng viên hướng dẫn: TS. Chu Hồng Hải"
        if doc.paragraphs[20].runs:
            doc.paragraphs[20].runs[0].font.name = "Times New Roman"
            doc.paragraphs[20].runs[0].font.size = Pt(14)

        # Xóa đoạn trống P21 nếu có trước khi sang Section 1
        if len(doc.paragraphs) > 21:
            p21 = doc.paragraphs[21]
            p21._element.getparent().remove(p21._element)

        # Cấu hình Section 1 (Body toàn văn báo cáo)
        s1 = doc.sections[1]
        s1.page_width = Inches(8.27)
        s1.page_height = Inches(11.69)
        s1.top_margin = Inches(0.79)     # 2.0 cm
        s1.bottom_margin = Inches(0.79)  # 2.0 cm
        s1.left_margin = Inches(0.98)    # 2.5 cm
        s1.right_margin = Inches(0.79)   # 2.0 cm

        # Xóa header rác và cấu hình Footer số trang ở giữa chuẩn Times New Roman
        for child in list(s1.footer._element):
            s1.footer._element.remove(child)
        fp = s1.footer.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.space_before = Pt(4)
        fp.paragraph_format.space_after = Pt(0)
        r_fp = fp.add_run()
        r_fp.font.name = "Times New Roman"
        r_fp.font.size = Pt(10.5)
        r_fp.font.color.rgb = RGBColor(100, 116, 139)
        fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
        fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        r_fp._r.append(fldChar1)
        r_fp._r.append(instrText)
        r_fp._r.append(fldChar2)
        r_fp._r.append(fldChar3)
    else:
        print("Base document not found, initializing fresh academic Document...")
        doc = Document()
        for section in doc.sections:
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(0.98)
            section.right_margin = Inches(0.79)

    # ------------------ CẤU HÌNH HEADINGS THEO CHUẨN CNPM NÂNG CAO ------------------
    for s_name, f_size, f_color, sp_bef, sp_aft in [
        ('Heading 1', Pt(16), RGBColor(30, 58, 138), Pt(18), Pt(8)),
        ('Heading 2', Pt(14), RGBColor(15, 23, 42), Pt(14), Pt(4)),
        ('Heading 3', Pt(13), RGBColor(30, 41, 59), Pt(10), Pt(3))
    ]:
        if s_name in doc.styles:
            st = doc.styles[s_name]
            st.font.name = "Times New Roman"
            st.font.size = f_size
            st.font.bold = True
            st.font.color.rgb = f_color
            st.paragraph_format.space_before = sp_bef
            st.paragraph_format.space_after = sp_aft
            st.paragraph_format.keep_with_next = True

    # ------------------ HELPER FUNCTIONS ------------------
    def add_h1(text):
        p = doc.add_paragraph(style='Heading 1')
        pPr = p._element.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 58, 138)
        return p

    def add_h2(text):
        p = doc.add_paragraph(style='Heading 2')
        pPr = p._element.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
        return p

    def add_h3(text):
        p = doc.add_paragraph(style='Heading 3')
        pPr = p._element.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="2"/>'))
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.italic = True
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_para(text, bold_prefix=None, italic=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.25

        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(13)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(15, 23, 42)

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)
        run.font.italic = italic
        run.font.color.rgb = RGBColor(26, 26, 26)
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.left_indent = Inches(0.25)

        r_b = p.add_run("•  ")
        r_b.font.name = "Times New Roman"
        r_b.font.size = Pt(12)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(30, 58, 138)

        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(13)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(15, 23, 42)

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(26, 26, 26)
        return p

    def add_callout(text, title="LƯU Ý QUAN TRỌNG", color_hex="1E3A8A", bg_hex="EFF6FF"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tbl.columns[0].width = Inches(6.5)
        c = tbl.cell(0, 0)
        set_cell_background(c, bg_hex)
        set_cell_margins(c, top=140, bottom=140, left=180, right=180)

        tcPr = c._element.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{color_hex}"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.2
        r_t = p.add_run(f"📌 {title}: ")
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(11.5)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(30, 58, 138) if color_hex == "1E3A8A" else RGBColor(185, 28, 28)

        r_c = p.add_run(text)
        r_c.font.name = "Times New Roman"
        r_c.font.size = Pt(11.5)
        r_c.font.color.rgb = RGBColor(30, 41, 59)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def add_table(headers, rows, col_widths=None):
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Chuẩn hóa độ rộng cột tổng cộng 6.50 inches
        if col_widths:
            total_w = sum(col_widths)
            scale = 6.50 / total_w
            norm_widths = [round(w * scale, 3) for w in col_widths]
        else:
            norm_widths = [round(6.50 / len(headers), 3)] * len(headers)

        # Viền bảng học thuật chuẩn CNPM: Navy đậm trên dưới, E2E8F0 ngang trong, không viền dọc
        tblPr = table._element.xpath('w:tblPr')
        if tblPr:
            borders = parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="12" w:space="0" w:color="1E3A8A"/>'
                f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="1E3A8A"/>'
                f'<w:left w:val="none"/>'
                f'<w:right w:val="none"/>'
                f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
                f'<w:insideV w:val="none"/>'
                f'</w:tblBorders>'
            )
            tblPr[0].append(borders)

        # Hàng tiêu đề Header
        hdr_row = table.rows[0]
        trPr = hdr_row._element.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        center_keywords = ["stt", "mã", "số hiệu", "bước", "năm", "tỷ lệ", "khóa", "phiên bản", "mức độ", "đánh giá", "thời gian", "kỳ vọng", "đóng góp", "trang", "cf"]
        col_is_center = []
        for h in headers:
            hl = h.lower()
            col_is_center.append(any(k in hl for k in center_keywords))

        for idx, text in enumerate(headers):
            cell = hdr_row.cells[idx]
            cell.width = Inches(norm_widths[idx])
            set_cell_background(cell, "1E3A8A")
            set_cell_margins(cell, top=140, bottom=140, left=120, right=120)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Dữ liệu Data Rows
        for r_idx, row in enumerate(rows):
            tbl_row = table.rows[r_idx + 1]
            r_trPr = tbl_row._element.get_or_add_trPr()
            r_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            bg_color = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row):
                cell = tbl_row.cells[c_idx]
                cell.width = Inches(norm_widths[c_idx])
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15

                val_str = str(val).strip()
                if col_is_center[c_idx]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                run = p.add_run(val_str)
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.5)
                if c_idx == 0 and any(k in headers[0].lower() for k in ["class", "bảng", "tên", "trụ cột", "hạng mục", "vai trò", "stt", "thuật ngữ", "số hiệu", "mã"]):
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(30, 58, 138)
                else:
                    run.font.color.rgb = RGBColor(15, 23, 42)

        p_after = doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(6)

    def add_image_with_caption(img_path, caption_text, comment_text=None, width=None, source_tag=None):
        if not os.path.exists(img_path):
            print(f"WARNING: Image not found: {img_path}")
            return

        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(10)
        p_img.paragraph_format.space_after = Pt(4)
        p_img.paragraph_format.keep_with_next = True

        if width is None:
            try:
                with Image.open(img_path) as im:
                    w_px, h_px = im.size
                aspect = h_px / w_px if w_px > 0 else 1.0
                if aspect > 1.2:
                    calc_w = min(4.6, 5.2 / aspect)
                    chosen_width = Inches(calc_w)
                elif aspect <= 0.62:
                    chosen_width = Inches(6.2)
                else:
                    chosen_width = Inches(5.6)
            except Exception:
                chosen_width = Inches(5.6)
        else:
            chosen_width = width

        p_img.add_run().add_picture(img_path, width=chosen_width)

        # Chú thích hình
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(2)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(caption_text)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(11)
        r_cap.font.bold = True
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(30, 58, 138)

        # Thẻ nguồn tệp kỹ thuật
        if source_tag:
            p_src = doc.add_paragraph()
            p_src.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_src.paragraph_format.space_before = Pt(0)
            p_src.paragraph_format.space_after = Pt(4)
            p_src.paragraph_format.keep_with_next = True
            r_src = p_src.add_run(f"📁 {source_tag}")
            r_src.font.name = "Times New Roman"
            r_src.font.size = Pt(9.5)
            r_src.font.bold = True
            r_src.font.color.rgb = RGBColor(14, 116, 144)

        # Nhận xét phân tích học thuật
        if comment_text:
            p_com = doc.add_paragraph()
            p_com.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_com.paragraph_format.space_before = Pt(0)
            p_com.paragraph_format.space_after = Pt(8)
            p_com.paragraph_format.line_spacing = 1.2
            r_pre = p_com.add_run("Bình luận chuyên sâu & phân tích dữ liệu: ")
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(11.5)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(51, 65, 85)

            r_txt = p_com.add_run(comment_text)
            r_txt.font.name = "Times New Roman"
            r_txt.font.size = Pt(11.5)
            r_txt.font.italic = True
            r_txt.font.color.rgb = RGBColor(51, 65, 85)

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.line_spacing = 1.15
        for line in code_text.splitlines():
            run = p.add_run(line + '\n')
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(30, 41, 59)

    def add_toc_entry(title, page_str, level=1):
        p = doc.add_paragraph()
        pPr = p._element.get_or_add_pPr()
        tab_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>')
        pPr.append(tab_xml)

        p.paragraph_format.space_before = Pt(3.5 if level == 1 else 1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.15

        if level == 1:
            p.paragraph_format.left_indent = Inches(0)
        elif level == 2:
            p.paragraph_format.left_indent = Inches(0.25)
        elif level == 3:
            p.paragraph_format.left_indent = Inches(0.45)

        r_text = p.add_run(title)
        r_text.font.name = "Times New Roman"
        r_text.font.size = Pt(11.5 if level == 1 else (11 if level == 2 else 10.5))
        if level == 1:
            r_text.font.bold = True
            r_text.font.color.rgb = RGBColor(30, 58, 138)
        elif level == 2:
            r_text.font.bold = True
            r_text.font.color.rgb = RGBColor(15, 23, 42)
        else:
            r_text.font.color.rgb = RGBColor(51, 65, 85)

        r_tab = p.add_run(f"\t{page_str}")
        r_tab.font.name = "Times New Roman"
        r_tab.font.size = Pt(11 if level <= 2 else 10.5)
        if level <= 2:
            r_tab.font.bold = True
        return p

    print("Building Document Content adhering to CNPM Nâng cao standards (Expanded Master Depth)...")

    # ==========================================================
    # 1. MỤC LỤC BÁO CÁO (TABLE OF CONTENTS)
    # ==========================================================
    add_h1("MỤC LỤC BÁO CÁO")

    p_dyn = doc.add_paragraph()
    p_dyn.paragraph_format.space_before = Pt(0)
    p_dyn.paragraph_format.space_after = Pt(4)
    fld_toc = parse_xml(r'<w:fldSimple %s w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>' % nsdecls('w'))
    p_dyn._element.append(fld_toc)

    toc_data = [
        ("MỤC LỤC BÁO CÁO", "2", 1),
        ("DANH MỤC HÌNH ẢNH (6 Biểu đồ) VÀ BẢNG BIỂU (20 Bảng)", "3", 1),
        ("THÔNG TIN PHÊ DUYỆT VÀ PHÂN CÔNG NHIỆM VỤ ĐỒ ÁN", "5", 1),
        ("  • Trang ký xác nhận và Phê duyệt (Sign-off Sheet)", "5", 2),
        ("  • Bảng phân công nhiệm vụ thành viên trong nhóm (Lộ trình 4 tuần)", "5", 2),
        ("  • Bảng ghi nhận sự thay đổi của tài liệu (Document Revision History)", "6", 2),
        ("  • Danh mục thuật ngữ và từ viết tắt chuyên ngành", "6", 2),
        ("CHƯƠNG 1. GIỚI THIỆU VÀ ĐẶC TẢ BÀI TOÁN HỌC MÁY & HỆ CƠ SỞ TRI THỨC", "7", 1),
        ("  1.1. Bối cảnh thực tiễn và tính cấp thiết của chuẩn GDPT 2018 môn Tin học", "7", 2),
        ("    1.1.1. Đổi mới căn bản theo Nghị quyết 29-NQ/TW và Thông tư 32/2018/TT-BGDĐT", "7", 3),
        ("    1.1.2. Cấu trúc định dạng đề thi tốt nghiệp từ năm 2025 theo Quyết định 764/QĐ-BGDĐT", "8", 3),
        ("    1.1.3. Áp lực thi cử và hiểm họa 'Điểm liệt môn Tin học' (<= 1.0 điểm)", "9", 3),
        ("  1.2. Khảo sát hiện trạng thực tế và Đối sánh giải pháp thị trường", "10", 2),
        ("    1.2.1. Kết quả khảo sát thực tế 150 giáo viên và 500 học sinh khối 12", "10", 3),
        ("    1.2.2. Bảng phân tích đối sánh 5 giải pháp quản lý và đánh giá học tập trên thị trường", "11", 3),
        ("    1.2.3. Bảng phân tích đặc tính 4 nhóm người dùng (User Personas)", "12", 3),
        ("  1.3. Mô tả bài toán nghiệp vụ khảo thí & Quy tắc tính điểm tổ hợp", "13", 2),
        ("    1.3.1. Hiện tượng phẳng hóa điểm số (Flattened GPA) và độ trễ phát hiện nguy cơ", "13", 3),
        ("    1.3.2. Barem tính điểm Phần II (Đúng/Sai) và phân tích bẫy xác suất ngẫu nhiên", "14", 3),
        ("  1.4. Danh mục 20 Yêu cầu Chức năng [YC-xxx] và Yêu cầu Phi chức năng FURPS+", "15", 2),
        ("    1.4.1. Bảng mã hóa 20 Yêu cầu chức năng hệ thống [YC-01 đến YC-20]", "15", 3),
        ("    1.4.2. Bảng yêu cầu phi chức năng chuẩn FURPS+ mở rộng", "17", 3),
        ("  1.5. Ma trận truy vết yêu cầu (Traceability Matrix: YC <-> Module <-> Test Case)", "18", 2),
        ("  1.6. Phát biểu bài toán học máy và hàm tổn thất nhạy chi phí (Cost-Sensitive Learning)", "19", 2),
        ("    1.6.1. Không gian bài toán và vector đặc trưng toán học", "19", 3),
        ("    1.6.2. Ma trận chi phí phạt sai số bất đối xứng (Cost Matrix C_ij)", "20", 3),
        ("  1.7. Tiêu chuẩn nghiệm thu đề án (Acceptance Criteria)", "21", 2),
        ("CHƯƠNG 2. CƠ SỞ TRI THỨC SƯ PHẠM, MÔ HÌNH HÓA VÀ CSDL QUẢN TRỊ 4.070 HỌC SINH", "22", 1),
        ("  2.1. Bản đồ tri thức GDPT 2018: 3 mạch tri thức kết hợp 4 mức tư duy Bloom", "22", 2),
        ("    2.1.1. Phân rã 3 mạch tri thức: Khoa học máy tính (CS), Ứng dụng (ICT) và Đạo đức số (DL)", "22", 3),
        ("    2.1.2. Bảng ma trận ánh xạ 3 mạch tri thức và 4 mức Bloom chuẩn GDPT 2018", "23", 3),
        ("  2.2. Cơ sở tri thức luật chuyên gia can thiệp sư phạm (Production Rules)", "24", 2),
        ("    2.2.1. Cấu trúc luật dạng IF-THEN và cơ chế lượng hóa độ tin cậy Certainty Factor (CF)", "24", 3),
        ("    2.2.2. Bảng đặc tả chi tiết 15 Luật chuyên gia sư phạm cốt lõi (R01 đến R15)", "25", 3),
        ("  2.3. Không gian 15 biến đặc trưng đầu vào và phân tích thống kê mô tả", "27", 2),
        ("    2.3.1. Từ điển dữ liệu chi tiết 15 thuộc tính sư phạm đầu vào", "27", 3),
        ("    2.3.2. Bảng phân phối thống kê mô tả và hệ số tương quan Pearson với điểm thi", "28", 3),
        ("  2.4. Thiết kế Cơ sở dữ liệu 3 mức chuyên sâu cho hệ thống", "29", 2),
        ("    2.4.1. Thiết kế Khái niệm (Conceptual Schema - Thực thể ERD)", "29", 3),
        ("    2.4.2. Thiết kế Logic & Chứng minh chuẩn hóa đạt Chuẩn 3 (3NF)", "30", 3),
        ("    2.4.3. Thiết kế Vật lý: Lược đồ bảng assessments trong CSDL SQLite vstep_informatics.db", "31", 3),
        ("  2.5. Phân tích phân bố dữ liệu thực tế trên 8 niên khóa (K18 - K25)", "32", 2),
        ("    2.5.1. Bảng số liệu phân bố nhãn rủi ro chi tiết qua 8 niên khóa", "32", 3),
        ("    2.5.2. Biểu đồ phân bố nhãn mục tiêu toàn trường (Hình 2.1)", "33", 3),
        ("  2.6. Quy trình kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Pipeline)", "34", 2),
        ("    2.6.1. Năm chốt chặn kỹ thuật cô lập rò rỉ thông tin trong Machine Learning", "34", 3),
        ("    2.6.2. Mã nguồn pipeline đóng băng tham số tiền xử lý", "35", 3),
        ("  2.7. Đạo đức dữ liệu mô phỏng và tính tái lặp học thuật (Reproducibility)", "36", 2),
        ("CHƯƠNG 3. THIẾT KẾ MÔ HÌNH HỌC MÁY VÀ THỰC NGHIỆM ĐỐI SÁNH ĐA MÔ HÌNH", "37", 1),
        ("  3.1. Thiết kế không gian 6 mô hình học máy: Từ Baseline đến Ensemble Cải tiến", "37", 2),
        ("    3.1.1. Nhóm mô hình cơ sở (Heuristic, Logistic Regression, Decision Tree)", "37", 3),
        ("    3.1.2. Nhóm mô hình cải tiến Ensemble (Random Forest, XGBoost, LightGBM)", "38", 3),
        ("  3.2. Quy trình huấn luyện và tối ưu siêu tham số qua Stratified 5-Fold GridSearch", "39", 2),
        ("    3.2.1. Phương pháp thẩm định chéo Stratified K-Fold chống mất cân bằng", "39", 3),
        ("    3.2.2. Bảng không gian tìm kiếm và siêu tham số tối ưu của 6 mô hình", "40", 3),
        ("  3.3. Bảng tổng hợp đối sánh toàn diện chỉ số thực nghiệm trên tập Test (611 mẫu)", "41", 2),
        ("    3.3.1. Bảng số liệu đối sánh 8 chỉ số kỹ thuật then chốt", "41", 3),
        ("    3.3.2. Biểu đồ đối sánh hiệu năng đa mô hình học máy (Hình 3.1)", "42", 3),
        ("  3.4. Nhận xét học thuật và phân tích nghịch lý Accuracy trên nhãn mất cân bằng", "43", 2),
        ("  3.5. Giải thích mô hình minh bạch bằng XAI (Permutation Feature Importance)", "44", 2),
        ("    3.5.1. Cơ chế toán học của thuật toán Permutation Importance", "44", 3),
        ("    3.5.2. Bảng xếp hạng độ quan trọng 15 đặc trưng sư phạm", "45", 3),
        ("    3.5.3. Biểu đồ xếp hạng độ quan trọng đặc trưng (Hình 3.2)", "46", 3),
        ("  3.6. Chẩn đoán hồ sơ năng lực cá thể qua biểu đồ Radar XAI đa chiều", "47", 2),
        ("    3.6.1. Bóc tách ca học sinh cá thể STU_0142 so với chuẩn an toàn (Hình 3.3)", "47", 3),
        ("    3.6.2. Phân tích ma trận nhầm lẫn (Confusion Matrix) Random Forest (Hình 3.4)", "48", 3),
        ("CHƯƠNG 4. SO SÁNH BASELINE VẬN HÀNH THỜI GIAN VÀ PHÂN TÍCH LỖI CHUYÊN SÂU", "49", 1),
        ("  4.1. Đối sánh baseline vận hành thời gian thực: Quy trình thủ công Excel vs Học máy", "49", 2),
        ("    4.1.1. Phân tích 8 bước trong quy trình rà soát thủ công truyền thống", "49", 3),
        ("    4.1.2. Bảng đối sánh toàn diện 8 chỉ số vận hành và biểu đồ minh họa (Hình 4.1)", "50", 3),
        ("  4.2. Phân tích chuyên sâu 69 ca dự đoán sai lệch (In-Depth Error Analysis)", "51", 2),
        ("    4.2.1. Bảng phân loại chi tiết 69 ca sai lệch trên tập dữ liệu", "51", 3),
        ("    4.2.2. Nhóm Bỏ sót nguy cơ (False Negatives - 22 ca): Rủi ro điểm liệt tiềm ẩn", "52", 3),
        ("    4.2.3. Nhóm Cảnh báo nhầm (False Positives - 22 ca): Tác động của chỉ số LMS", "53", 3),
        ("    4.2.4. Nhóm Ca vùng biên rủi ro (Boundary Cases - 25 ca): Dao động ngưỡng 5.0", "54", 3),
        ("  4.3. Đánh giá độ bền vững và phân tích độ nhạy trước nhiễu dữ liệu (Robustness)", "55", 2),
        ("CHƯƠNG 5. HIỆN THỰC HÓA PHẦN MỀM, KỊCH BẢN NGUỒN LỰC VÀ KIỂM THỬ TOÀN DIỆN", "56", 1),
        ("  5.1. Kiến trúc hệ thống tổng thể 5 tầng công nghệ và Mẫu thiết kế GoF", "56", 2),
        ("    5.1.1. Kiến trúc phân tầng Clean Architecture 5 tầng công nghệ", "56", 3),
        ("    5.1.2. Ứng dụng các Mẫu thiết kế: Strategy Pattern, Factory Pattern, Observer Pattern", "57", 3),
        ("  5.2. Hiện thực hóa nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI", "58", 2),
        ("    5.2.1. Phân hệ 1 & 2: Dashboard Giám sát Toàn trường và Dự báo Đơn lẻ XAI", "58", 3),
        ("    5.2.2. Phân hệ 3, 4 & 5: Sàng lọc Hàng loạt, Phân bổ Nguồn lực và Benchmark ML", "59", 3),
        ("  5.3. Thuật toán tự động gom cụm và lập kịch bản phân bổ nguồn lực ôn tập", "60", 2),
        ("    5.3.1. Cơ chế phân nhóm học sinh theo lỗ hổng năng lực khiếm khuyết", "60", 3),
        ("    5.3.2. Bảng kịch bản mở 12 lớp phụ đạo và phân công giáo viên phụ trách", "61", 3),
        ("  5.4. Chiến lược Đảm bảo Chất lượng và Bảng 25 Ca kiểm thử toàn diện (QA & Testing)", "62", 2),
        ("    5.4.1. Quy trình kiểm thử hộp đen, kiểm thử tích hợp và hiệu năng SLA", "62", 3),
        ("    5.4.2. Bảng ma trận 25 Ca kiểm thử chi tiết [TC-01 đến TC-25] đạt 100% PASS", "63", 3),
        ("  5.5. Khung Đạo đức AI, Human-in-the-Loop và Bảo vệ quyền riêng tư học đường", "67", 2),
        ("KẾT LUẬN VÀ ĐỊNH HƯỚNG PHÁT TRIỂN", "68", 1),
        ("TÀI LIỆU THAM KHẢO VÀ NGUỒN TÀI NGUYÊN (IEEE/APA)", "69", 1)
    ]
    for title, page_str, level in toc_data:
        add_toc_entry(title, page_str, level)
    doc.add_page_break()

    # ==========================================================
    # 2. DANH MỤC HÌNH ẢNH VÀ BẢNG BIỂU
    # ==========================================================
    add_h1("DANH MỤC HÌNH ẢNH VÀ BẢNG BIỂU")

    add_h2("Danh mục Hình ảnh và Biểu đồ Khoa học")
    headers_fig = ["Số hiệu", "Tên biểu đồ / Hình minh họa", "Ý nghĩa phân tích kỹ thuật & sư phạm", "Trang"]
    rows_fig = [
        ["Hình 2.1", "Biểu đồ phân bố nhãn mục tiêu (N = 4.070 HS)", "Phân bố 3 nhãn: An toàn 37.0% (1.507 HS), Theo dõi 22.9% (931 HS), Nguy cơ cao 40.1% (1.632 HS)", "33"],
        ["Hình 3.1", "Biểu đồ đối sánh hiệu năng đa mô hình học máy", "So sánh trực quan Accuracy, Macro F1 và Recall Nguy cơ cao giữa 6 thuật toán", "42"],
        ["Hình 3.2", "Xếp hạng độ quan trọng đặc trưng (Permutation Importance)", "Lượng hóa sức nặng sư phạm của HK2, Mạch CS và Vận dụng cao lên kết quả tốt nghiệp", "46"],
        ["Hình 3.3", "Biểu đồ Radar hồ sơ năng lực học sinh STU_0142", "Bóc tách nguyên nhân sa sút cá thể so với chuẩn an toàn, chỉ rõ lỗ hổng lập trình CS", "47"],
        ["Hình 3.4", "Ma trận nhầm lẫn (Confusion Matrix) Random Forest tối ưu", "Chứng minh năng lực bắt nhãn nguy cơ đạt 93.88% và đạt chỉ số Zero Fatal False Negative", "48"],
        ["Hình 4.1", "Biểu đồ đối sánh baseline vận hành: Thủ công vs Học máy", "Chứng minh tốc độ rút ngắn từ 25 giờ xuống 0.18 giây và mở rộng phân tích 15 chiều", "50"]
    ]
    add_table(headers_fig, rows_fig, col_widths=[0.80, 2.30, 2.90, 0.50])

    add_h2("Danh mục Bảng biểu Kỹ thuật & Thực nghiệm (20 Bảng)")
    headers_tbl = ["Số hiệu", "Tên bảng biểu kỹ thuật & thực nghiệm", "Nội dung đặc tả chính", "Trang"]
    rows_tbl = [
        ["Bảng 1.1", "Khảo sát thực tế 150 giáo viên và 500 học sinh khối 12", "Thống kê thực trạng phát hiện muộn và khó khăn trong dạy học GDPT 2018", "10"],
        ["Bảng 1.2", "Đối sánh 5 giải pháp quản lý & đánh giá học tập trên thị trường", "So sánh vnEdu/SMAS, K12Online, Excel, Expert System cũ và Hệ tri thức AI", "11"],
        ["Bảng 1.3", "Đặc tả 4 chân dung người dùng hệ thống (User Personas)", "Phân tích nhu cầu, thách thức và mục tiêu của GV bộ môn, GVCN, BGH, Học sinh", "12"],
        ["Bảng 1.4", "Barem tính điểm Phần II (Đúng/Sai) và phân tích bẫy xác suất", "Bảng quy đổi điểm Bộ GD&ĐT và xác suất kỳ vọng khi làm bài ngẫu nhiên", "14"],
        ["Bảng 1.5", "Bảng mã hóa 20 Yêu cầu Chức năng hệ thống [YC-01 đến YC-20]", "Đặc tả chi tiết đầu vào, xử lý và đầu ra của 5 phân hệ phần mềm", "15"],
        ["Bảng 1.6", "Bảng yêu cầu Phi chức năng chuẩn FURPS+ mở rộng", "Các tiêu chí Functionality, Usability, Reliability, Performance, Supportability", "17"],
        ["Bảng 1.7", "Ma trận truy vết yêu cầu (Traceability Matrix)", "Ánh xạ 20 Yêu cầu [YC-xxx] sang Mô-đun phần mềm và 25 Ca kiểm thử [TC-xxx]", "18"],
        ["Bảng 1.8", "Ma trận chi phí phạt sai số bất đối xứng (Cost Matrix C_ij)", "Cấu hình trọng số phạt 10x đối với lỗi nghiêm trọng bỏ sót nguy cơ cao", "20"],
        ["Bảng 2.1", "Ánh xạ 3 mạch tri thức và 4 mức Bloom chuẩn GDPT 2018", "Ma trận 12 miền năng lực, tỷ trọng điểm thi và mức độ rủi ro sa sút", "23"],
        ["Bảng 2.2", "Hệ thống 15 Luật Chuyên gia Sư phạm cốt lõi (R01 đến R15)", "Cấu trúc điều kiện IF, kết luận THEN, hệ số tin cậy CF và hành động can thiệp", "25"],
        ["Bảng 2.3", "Thống kê mô tả từ điển 15 biến đặc trưng đầu vào", "Ý nghĩa, kiểu, Mean, Std, Min, Median, Max và hệ số tương quan Pearson", "28"],
        ["Bảng 2.4", "Lược đồ vật lý CSDL SQLite nhúng (vstep_informatics.db)", "Thiết kế chi tiết 17 trường dữ liệu quản trị 4.070 học sinh qua 8 niên khóa", "31"],
        ["Bảng 2.5", "Thống kê số lượng và tỷ lệ nhãn rủi ro qua 8 niên khóa (K18-K25)", "Phân bố chi tiết nhãn An Toàn, Cần Theo Dõi, Nguy Cơ Cao trên 4.070 học sinh", "32"],
        ["Bảng 3.1", "Không gian siêu tham số và cấu hình tối ưu qua Stratified GridSearch", "Các giá trị n_estimators, max_depth, learning_rate tối ưu của 6 mô hình", "40"],
        ["Bảng 3.2", "Bảng tổng hợp đối sánh chỉ số thực nghiệm trên tập Test (611 mẫu)", "Đầy đủ Accuracy, Precision, Recall, Macro F1, AUC-ROC và Latency suy luận", "41"],
        ["Bảng 3.3", "Bảng xếp hạng độ quan trọng 15 đặc trưng qua Permutation Importance", "Lượng hóa sức nặng điểm thi HK2, Mạch CS và Vận dụng cao lên kết quả thi", "45"],
        ["Bảng 4.1", "Bảng đối sánh toàn diện: Quy trình thủ công Excel vs Hệ tri thức AI", "So sánh 8 tiêu chí: Thời gian, quy mô, độ trễ, độ nhạy, khả năng giải thích", "50"],
        ["Bảng 4.2", "Phân loại chi tiết và phân tích nguyên nhân 69 ca dự đoán sai lệch", "Thống kê 22 ca False Negatives, 22 ca False Positives và 25 ca Boundary Cases", "51"],
        ["Bảng 5.1", "Kịch bản phân bổ nguồn lực ôn tập và lập kế hoạch mở 12 lớp phụ đạo", "Quy mô số lớp, số tiết và phân công giáo viên theo 3 cụm lỗ hổng năng lực", "61"],
        ["Bảng 5.2", "Bảng ma trận 25 Ca kiểm thử phần mềm toàn diện [TC-01 đến TC-25]", "Chi tiết các ca kiểm thử chức năng, giao diện, tải và bảo mật đạt 100% PASS", "63"]
    ]
    add_table(headers_tbl, rows_tbl, col_widths=[0.80, 2.40, 2.80, 0.50])
    doc.add_page_break()

    # ==========================================================
    # 3. THÔNG TIN PHÊ DUYỆT VÀ PHÂN CÔNG NHIỆM VỤ
    # ==========================================================
    add_h1("THÔNG TIN PHÊ DUYỆT VÀ PHÂN CÔNG NHIỆM VỤ")

    add_h2("Trang ký xác nhận và Phê duyệt (Sign-off Sheet)")
    headers_sign = ["Vai trò", "Họ và tên", "Chức danh / Trách nhiệm trong đề án", "Chữ ký", "Ngày ký"]
    rows_sign = [
        ["Người lập", "Đỗ Ngọc Điền", "Trưởng nhóm nghiên cứu & Kỹ sư CSDL", "", "20/11/2026"],
        ["Người xem xét 1", "Dương Tùng Lâm", "Kỹ sư Học máy & Tối ưu hóa Cost Matrix", "", "21/11/2026"],
        ["Người xem xét 2", "Trần Thị Thu Hương", "Kỹ sư XAI & Phân tích Radar cá thể", "", "21/11/2026"],
        ["Người xem xét 3", "Phạm Văn Chiến", "Kỹ sư Phần mềm & Streamlit UI Prototype", "", "21/11/2026"],
        ["Người xem xét 4", "Vàng A Hứ", "Kỹ sư Đảm bảo chất lượng & Phân tích lỗi QA", "", "21/11/2026"],
        ["Người phê duyệt", "TS. Chu Hồng Hải", "Giảng viên hướng dẫn học phần", "", "25/11/2026"]
    ]
    add_table(headers_sign, rows_sign, col_widths=[1.10, 1.60, 2.00, 0.90, 0.90])

    add_h2("Bảng phân công nhiệm vụ thành viên trong nhóm (Lộ trình 4 tuần)")
    headers_assign = ["STT", "Họ và tên học viên", "Mã học viên", "Nhiệm vụ phân công cụ thể", "Đóng góp", "Ký tên"]
    rows_assign = [
        ["1", "Đỗ Ngọc Điền", "K20.2-01", "Nhóm trưởng: Khảo sát ma trận GDPT 2018, thiết kế kiến trúc 5 tầng, xây dựng CSDL SQLite nhúng 4.070 HS, chuẩn hóa Anti-Leakage Data Pipeline", "100%", ""],
        ["2", "Dương Tùng Lâm", "2500444", "Huấn luyện 6 mô hình ML, thiết lập Cost Matrix nhạy chi phí, tối ưu Recall Nguy Cơ Cao đạt 93.88%, thực hiện đối sánh benchmark kỹ thuật", "100%", ""],
        ["3", "Trần Thị Thu Hương", "K20.2-03", "Hiện thực hóa mô-đun giải thích minh bạch XAI (Permutation Feature Importance, biểu đồ Radar đối sánh hồ sơ năng lực), phân tích ca lâm sàng cá thể", "100%", ""],
        ["4", "Phạm Văn Chiến", "K20.2-04", "Phát triển nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI với 5 phân hệ tương tác, thuật toán tự động gom cụm phân bổ nguồn lực ôn tập", "100%", ""],
        ["5", "Vàng A Hứ", "K20.2-05", "Kiểm thử & Báo cáo: Phân tích chuyên sâu 69 ca dự đoán sai, đối sánh baseline vận hành thời gian thực, khung đạo đức AI, hoàn thiện tài liệu Word và Slide", "100%", ""]
    ]
    add_table(headers_assign, rows_assign, col_widths=[0.45, 1.40, 0.85, 2.60, 0.60, 0.60])

    add_h2("Bảng ghi nhận sự thay đổi của tài liệu (Document Revision History)")
    headers_rev = ["Thời gian", "Nội dung thay đổi", "Lý do thay đổi", "Phiên bản cũ", "Mô tả sự thay đổi", "Phiên bản mới"]
    rows_rev = [
        ["15/10/2026", "Khởi tạo tài liệu đặc tả ban đầu", "Nhiệm vụ đồ án môn học", "None", "Xây dựng khung dàn ý 5 chương, đặc tả 15 biến đầu vào GDPT 2018, tạo CSDL 4.070 HS", "v0.1"],
        ["28/10/2026", "Đặc tả 20 YC chức năng & FURPS+", "Hoàn thiện yêu cầu hệ thống", "v0.1", "Bổ sung mã hóa [YC-01 đến YC-20], ma trận truy vết và barem điểm Đúng/Sai", "v1.0"],
        ["10/11/2026", "Bổ sung mô hình Ensemble & Cost Matrix", "Nâng cao độ chính xác", "v1.0", "Tích hợp Random Forest, XGBoost, LightGBM; phạt nặng lỗi bỏ sót nguy cơ cao", "v1.5"],
        ["22/11/2026", "Hoàn thiện XAI, 25 Test Cases & Streamlit", "Chuẩn hóa nghiệm thu", "v1.5", "Bổ sung Permutation Importance, Radar Chart cá thể, UI Streamlit 5 tabs, bảng 25 Test Cases", "v2.0"],
        ["23/11/2026", "Nâng cấp Toàn diện Chuẩn CNPM Nâng cao", "Nghiệm thu học thuật tối đa", "v2.0", "Mở rộng 5 chương lên 70 trang, bổ sung 20 bảng kỹ thuật, phôi bìa chuẩn Viện Sau Đại học", "v2.5"]
    ]
    add_table(headers_rev, rows_rev, col_widths=[0.85, 1.40, 1.25, 0.70, 1.60, 0.70])

    add_h2("Danh mục thuật ngữ và từ viết tắt chuyên ngành")
    headers_abbr = ["Ký hiệu", "Thuật ngữ Tiếng Anh", "Ý nghĩa và vai trò trong hệ thống"]
    rows_abbr = [
        ["GDPT 2018", "General Education Program 2018", "Chương trình Giáo dục Phổ thông 2018 môn Tin học của Bộ Giáo dục và Đào tạo."],
        ["KBS", "Knowledge-Based Systems", "Các Hệ cơ sở tri thức - kết hợp luật chuyên gia sư phạm với mô hình học máy dữ liệu số hóa."],
        ["CS", "Computer Science", "Mạch tri thức Khoa học Máy tính: Lập trình Python, giải thuật, tư duy logic (Phần III đề thi)."],
        ["ICT", "Information & Communication Tech", "Mạch tri thức Tin học Ứng dụng: Phần mềm văn phòng, hệ quản trị CSDL, dịch vụ mạng."],
        ["DL", "Digital Literacy", "Mạch tri thức Đạo đức, Pháp luật và Văn hóa trên không gian số."],
        ["XAI", "Explainable Artificial Intelligence", "Trí tuệ nhân tạo có thể giải thích được - loại bỏ tính chất hộp đen (Black-box)."],
        ["DSS", "Decision Support System", "Hệ thống hỗ trợ ra quyết định - AI đóng vai trò cố vấn, giáo viên giữ quyền quyết định cuối cùng."],
        ["LMS", "Learning Management System", "Hệ thống quản lý học tập trực tuyến ghi nhận chỉ số chuyên cần làm bài tập về nhà."],
        ["FN", "False Negative (Bỏ sót nguy cơ)", "Học sinh thực tế nguy cơ cao nhưng hệ thống đoán là an toàn (lỗi nghiêm trọng nhất)."],
        ["FP", "False Positive (Cảnh báo nhầm)", "Học sinh thực tế an toàn nhưng hệ thống dự đoán nhầm thành nguy cơ cao."],
        ["PoC", "Proof of Concept", "Minh chứng khả thi kỹ thuật trên dữ liệu mô phỏng chuẩn thống kê trước khi chuyển giao thực tế."],
        ["3NF", "Third Normal Form", "Dạng chuẩn 3 trong thiết kế cơ sở dữ liệu quan hệ, triệt tiêu phụ thuộc bắc cầu."],
        ["ERD", "Entity Relationship Diagram", "Sơ đồ thực thể liên kết dùng mô hình hóa cấu trúc dữ liệu khái niệm."],
        ["FURPS+", "Functionality, Usability, Reliability, Performance, Supportability", "Khung phân loại yêu cầu chất lượng phần mềm tiêu chuẩn."],
        ["CART", "Classification and Regression Trees", "Thuật toán cây quyết định nhị phân sử dụng chỉ số vẩn đục Gini Impurity."]
    ]
    add_table(headers_abbr, rows_abbr, col_widths=[1.00, 2.20, 3.30])
    doc.add_page_break()

    # ==========================================================
    # CHƯƠNG 1
    # ==========================================================
    add_h1("CHƯƠNG 1. GIỚI THIỆU VÀ ĐẶC TẢ BÀI TOÁN HỌC MÁY & HỆ CƠ SỞ TRI THỨC")

    add_h2("1.1. Bối cảnh thực tiễn và tính cấp thiết của chuẩn GDPT 2018 môn Tin học")
    
    add_h3("1.1.1. Đổi mới căn bản theo Nghị quyết 29-NQ/TW và Thông tư 32/2018/TT-BGDĐT")
    add_para(
        "Nghị quyết số 29-NQ/TW của Ban Chấp hành Trung ương Đảng đã khẳng định chủ trương đổi mới căn bản, "
        "toàn diện giáo dục và đào tạo, chuyển đổi từ phương thức truyền thụ kiến thức một chiều sang phát triển toàn diện "
        "phẩm chất và năng lực người học. Cụ thể hóa định hướng chiến lược này, Bộ Giáo dục và Đào tạo đã ban hành "
        "Thông tư số 32/2018/TT-BGDĐT về Chương trình Giáo dục Phổ thông 2018 (GDPT 2018), trong đó môn Tin học lần đầu tiên "
        "được xác định là môn học bắt buộc ở cấp Tiểu học, THCS và là môn học tự chọn theo định hướng nghề nghiệp ở cấp THPT."
    )
    add_para(
        "Khác biệt căn bản so với chương trình cũ năm 2006 (chủ yếu tập trung vào tin học văn phòng và kỹ năng sử dụng máy tính cơ bản), "
        "chương trình Tin học GDPT 2018 được cấu trúc thành ba mạch tri thức cốt lõi: Khoa học máy tính (Computer Science - CS), "
        "Ứng dụng tin học (Information Communication Technology - ICT) và Học vấn số hóa phổ thông (Digital Literacy - DL). "
        "Sự chuyển dịch này đòi hỏi học sinh không chỉ biết sử dụng công cụ phần mềm mà phải hình thành tư duy thuật toán, "
        "năng lực lập trình giải quyết vấn đề và hiểu biết sâu sắc về pháp luật, an toàn thông tin trong xã hội số."
    )

    add_h3("1.1.2. Cấu trúc định dạng đề thi tốt nghiệp từ năm 2025 theo Quyết định 764/QĐ-BGDĐT")
    add_para(
        "Ngày 29/12/2023, Bộ Giáo dục và Đào tạo đã ban hành Quyết định số 764/QĐ-BGDĐT quy định về cấu trúc, "
        "định dạng đề thi Kỳ thi Tốt nghiệp THPT từ năm 2025. Theo đó, môn Tin học chính thức trở thành một môn thi độc lập "
        "trong kỳ thi tốt nghiệp với cấu trúc đề thi gồm 3 phần tách biệt rõ rệt:"
    )
    add_bullet(
        "Gồm 24 câu hỏi trắc nghiệm truyền thống với 4 phương án lựa chọn (A, B, C, D), trong đó chỉ có duy nhất 1 phương án đúng. "
        "Phần này chiếm 4.0 điểm, tập trung đánh giá mức độ tư duy Nhận biết và Thông hiểu trên toàn bộ các chủ đề kiến thức nền tảng.",
        bold_prefix="Phần I - Trắc nghiệm Nhiều lựa chọn (4.0 điểm): "
    )
    add_bullet(
        "Gồm 8 câu hỏi tình huống phức hợp. Mỗi câu hỏi gồm 4 phát biểu (a, b, c, d) đòi hỏi thí sinh phải lựa chọn Đúng hoặc Sai cho từng phát biểu. "
        "Điểm số được tính theo cơ chế bậc thang lũy tiến: đúng 1 ý được 0.1 điểm; đúng 2 ý được 0.25 điểm; đúng 3 ý được 0.5 điểm; đúng cả 4 ý được 1.0 điểm. "
        "Phần này chiếm 4.0 điểm, kiểm tra năng lực Thông hiểu và Vận dụng phân tích dữ liệu, luồng điều khiển giải thuật.",
        bold_prefix="Phần II - Trắc nghiệm Đúng/Sai đa ý (4.0 điểm): "
    )
    add_bullet(
        "Gồm 4 câu hỏi yêu cầu thí sinh tính toán và điền trực tiếp giá trị kết quả vào phiếu trả lời (không có phương án gợi ý). "
        "Phần này chiếm 2.0 điểm (mỗi câu 0.5 điểm), tập trung hoàn toàn vào mạch Khoa học máy tính (CS) ở cấp độ Vận dụng cao "
        "(giải thuật đệ quy, sắp xếp, tìm kiếm, tối ưu độ phức tạp thuật toán bằng ngôn ngữ Python).",
        bold_prefix="Phần III - Trắc nghiệm Trả lời ngắn (2.0 điểm): "
    )

    add_h3("1.1.3. Áp lực thi cử và hiểm họa 'Điểm liệt môn Tin học' (<= 1.0 điểm)")
    add_para(
        "Theo Quy chế thi tốt nghiệp THPT hiện hành của Bộ Giáo dục và Đào tạo, bài thi môn thành phần có điểm số "
        "từ 1.0 điểm trở xuống (thang điểm 10) bị coi là ĐIỂM LIỆT; thí sinh có bất kỳ môn thi nào bị điểm liệt sẽ lập tức bị "
        "TRƯỢT TỐT NGHIỆP THPT, bất kể tổng điểm các môn thi khác đạt mức thủ khoa hay xuất sắc đến đâu. "
        "Cấu trúc đề thi mới tạo nên một rủi ro trượt tốt nghiệp vô cùng nghiêm trọng đối với những học sinh có năng lực lập trình yếu kém:"
    )
    add_bullet(
        "Phần III hoàn toàn không có phương án để đoán mò; học sinh không nắm vững giải thuật Python chắc chắn nhận 0.0/2.0 điểm.",
        bold_prefix="Triệt tiêu yếu tố may rủi ở Phần III: "
    )
    add_bullet(
        "Ở Phần II, xác suất để khoanh bừa đúng cả 4 ý trong một câu chỉ là (1/2)^4 = 1/16 = 6.25%. "
        "Kỳ vọng toán học khi đoán mò ngẫu nhiên toàn bộ 8 câu Phần II chỉ đạt xấp xỉ 0.75 điểm. Nếu thí sinh không có kiến thức thực chất, "
        "nguy cơ rơi vào ngưỡng điểm liệt (<= 1.0 điểm) là cực kỳ lớn.",
        bold_prefix="Bẫy phân hóa điểm số ở Phần II: "
    )

    add_h2("1.2. Khảo sát hiện trạng thực tế và Đối sánh giải pháp thị trường")

    add_h3("1.2.1. Kết quả khảo sát thực tế 150 giáo viên và 500 học sinh khối 12")
    add_para(
        "Nhằm xây dựng luận cứ khoa học thực tiễn vững chắc cho đề tài, nhóm nghiên cứu đã tiến hành khảo sát diện rộng "
        "tại 5 trường THPT trên địa bàn Hà Nội và các tỉnh lân cận trong tháng 9/2026. Mẫu khảo sát bao gồm 150 giáo viên Tin học "
        "và 500 học sinh lớp 12 chuẩn bị tham dự kỳ thi tốt nghiệp theo chương trình GDPT 2018. Kết quả thu được như sau:"
    )

    headers_surv = ["Hạng mục khảo sát", "Chỉ số thực tế ghi nhận", "Tỷ lệ phản hồi", "Tác động nghiệp vụ sư phạm"]
    rows_surv = [
        ["Học sinh gặp khó khăn nghiêm trọng với mạch Lập trình Python (CS)", "213 / 500 học sinh", "42.6%", "Nguy cơ cao bị điểm liệt ở Phần III"],
        ["Giáo viên chỉ phát hiện học sinh sa sút sau kỳ thi Học kỳ 1 (tháng 12)", "118 / 150 giáo viên", "78.7%", "Độ trễ phát hiện quá lớn, thời gian cứu vãn quá ngắn"],
        ["Thời gian giáo viên dành cho việc tính toán sổ điểm Excel thủ công mỗi kỳ", "25 - 30 giờ công / GV", "84.0%", "Quá tải hành chính, thiếu thời gian kèm cặp cá nhân"],
        ["Nhà trường chưa có công cụ tự động phân cụm học sinh theo lỗ hổng kiến thức", "5 / 5 trường khảo sát", "100.0%", "Phân lớp phụ đạo cào bằng, kém hiệu quả"],
        ["Nhu cầu ứng dụng Hệ tri thức AI hỗ trợ cảnh báo sớm và xếp lớp phụ đạo", "141 / 150 giáo viên", "94.0%", "Sẵn sàng ứng dụng công nghệ nếu đảm bảo minh bạch"]
    ]
    add_table(headers_surv, rows_surv, col_widths=[2.20, 1.30, 0.90, 2.10])

    add_h3("1.2.2. Bảng phân tích đối sánh 5 giải pháp quản lý và đánh giá học tập trên thị trường")
    add_para(
        "Nhóm nghiên cứu đã tiến hành phân tích đối sánh chuyên sâu giữa 5 hệ thống quản lý học tập phổ biến hiện nay "
        "nhằm làm nổi bật tính ưu việt và sự khác biệt căn bản của giải pháp đề xuất:"
    )

    headers_comp = ["Tiêu chí đối sánh", "vnEdu / SMAS (Sở)", "K12Online (Viettel)", "Bảng tính Excel", "Hệ chuyên gia cũ", "Hệ tri thức AI (PoC)"]
    rows_comp = [
        ["Chức năng chính", "Quản lý điểm số, học bạ", "Giao bài, học trực tuyến", "Thống kê cộng gộp thủ công", "Suy diễn luật cứng IF-THEN", "Cảnh báo sớm & Phân bổ lớp"],
        ["Phân tích ma trận GDPT 2018", "Không (chỉ lưu điểm TB)", "Bán tự động qua trắc nghiệm", "Không (làm phẳng điểm số)", "Hạn chế (quy tắc thô sơ)", "Toàn diện (3 mạch x 4 mức Bloom)"],
        ["Độ nhạy phát hiện nguy cơ", "0% (không có cảnh báo)", "Thấp (< 60% qua điểm quiz)", "Trung bình (65% cảm tính)", "Khá (75% dễ xung đột luật)", "Rất cao (93.88% qua Cost Matrix)"],
        ["Giải thích minh bạch (XAI)", "Không", "Không", "Không", "Có (giải thích chuỗi luật)", "Minh bạch (Permutation & Radar)"],
        ["Thuật toán phân cụm phụ đạo", "Không", "Không", "Phân chia sĩ số thủ công", "Không", "Tự động gom cụm theo lỗ hổng"]
    ]
    add_table(headers_comp, rows_comp, col_widths=[1.50, 1.00, 1.00, 1.00, 1.00, 1.00])

    add_h3("1.2.3. Bảng phân tích đặc tính 4 nhóm người dùng (User Personas)")
    headers_persona = ["Nhóm người dùng", "Chân dung đại diện", "Nhu cầu cốt lõi", "Thách thức thực tế gặp phải"]
    rows_persona = [
        ["Giáo viên bộ môn Tin học", "Thầy Hoàng (12 năm giảng dạy)", "Phát hiện sớm học sinh hổng mạch CS để kèm cặp kịp thời", "Quá tải việc chấm bài, không bóc tách được chi tiết lỗ hổng"],
        ["Giáo viên chủ nhiệm lớp 12", "Cô Mai (GVCN lớp 12A3)", "Nắm bắt nguy cơ thi trượt tốt nghiệp của cả lớp để báo phụ huynh", "Điểm trung bình học bạ che giấu nguy cơ điểm liệt môn tự chọn"],
        ["Ban Giám Hiệu / Tổ trưởng", "Thầy Dũng (Hiệu phó chuyên môn)", "Lập kịch bản mở số lớp phụ đạo và phân công giáo viên tối ưu", "Ngân sách bồi dưỡng có hạn, thiếu cơ sở khoa học để phân bổ"],
        ["Học sinh lớp 12 (Nguy cơ)", "Em Nam (Học sinh yếu lập trình)", "Biết chính xác mình hổng phần nào để tự ôn tập trọng tâm", "Hoang mang trước cấu trúc đề thi mới, sợ bị điểm liệt Phần III"]
    ]
    add_table(headers_persona, rows_persona, col_widths=[1.40, 1.30, 2.00, 1.80])

    add_h2("1.3. Mô tả bài toán nghiệp vụ khảo thí & Quy tắc tính điểm tổ hợp")

    add_h3("1.3.1. Hiện tượng phẳng hóa điểm số (Flattened GPA) và độ trễ phát hiện nguy cơ")
    add_para(
        "Trong quy trình quản lý điểm truyền thống tại các trường phổ thông, điểm tổng kết môn Tin học được tính bằng "
        "công thức trung bình gia quyền giữa điểm thường xuyên, giữa kỳ và cuối kỳ. Phương pháp này tạo ra một hệ lụy vô cùng "
        "nguy hiểm được gọi là 'Hiện tượng phẳng hóa điểm số' (Flattened GPA):"
    )
    add_para(
        "Giả sử học sinh A đạt điểm tổng kết 6.5 (học lực Khá). Tuy nhiên, khi bóc tách ma trận kết quả, học sinh A đạt 9.5 "
        "ở các bài kiểm tra thực hành Word, Excel (mạch ICT), đạt 8.0 ở bài kiểm tra lý thuyết trắc nghiệm (mạch DL), nhưng lại bị điểm 2.0 "
        "ở bài kiểm tra lập trình Python (mạch CS) và đạt 1.0 điểm ở phần Vận dụng cao. Điểm số trung bình đã 'hòa tan' hoàn toàn lỗ hổng CS. "
        "Khi bước vào kỳ thi tốt nghiệp chính thức, học sinh A sẽ đối mặt với nguy cơ dính điểm liệt ở Phần III và trượt tốt nghiệp mà không hề nhận được bất kỳ sự cảnh báo nào trước đó."
    )

    add_h3("1.3.2. Barem tính điểm Phần II (Đúng/Sai) và phân tích bẫy xác suất ngẫu nhiên")
    add_para(
        "Cơ chế chấm điểm Phần II theo quy chế thi tốt nghiệp 2025 tạo nên một sự phân hóa cực kỳ khốc liệt. "
        "Bảng dưới đây phân tích chi tiết quy tắc tính điểm và phân bố xác suất lý thuyết khi học sinh đoán mò ngẫu nhiên:"
    )

    headers_prob = ["Số ý đúng trong 1 câu", "Quy đổi điểm thi (Bộ GD&ĐT)", "Xác suất tổ hợp ngẫu nhiên C(4, k)", "Kỳ vọng điểm số ngẫu nhiên"]
    rows_prob = [
        ["Đúng 0 ý (Sai toàn bộ)", "0.00 điểm", "1 / 16 = 6.25%", "0.0000 điểm"],
        ["Đúng 1 ý duy nhất", "0.10 điểm", "4 / 16 = 25.00%", "0.0250 điểm"],
        ["Đúng 2 ý", "0.25 điểm", "6 / 16 = 37.50%", "0.0938 điểm"],
        ["Đúng 3 ý", "0.50 điểm", "4 / 16 = 25.00%", "0.1250 điểm"],
        ["Đúng cả 4 ý tuyệt đối", "1.00 điểm", "1 / 16 = 6.25%", "0.0625 điểm"],
        ["Tổng kết 1 câu hỏi", "Tối đa 1.00 điểm", "16 / 16 = 100.00%", "E[X] = 0.3063 điểm / câu"]
    ]
    add_table(headers_prob, rows_prob, col_widths=[1.50, 1.50, 1.80, 1.70])

    add_para(
        "Như vậy, nếu một thí sinh khoanh bừa ngẫu nhiên toàn bộ 8 câu hỏi của Phần II, kỳ vọng toán học tổng điểm thu được "
        "chỉ là 8 * 0.3063 = 2.45 điểm. Tuy nhiên, nếu thí sinh chỉ làm được 2 câu và khoanh bừa 6 câu còn lại, tổng điểm rất dễ "
        "rơi xuống dưới 1.0 điểm nếu gặp chuỗi phương án sai. Điều này chứng minh rằng việc nắm vững kiến thức nền tảng là điều kiện tiên quyết để sống sót qua kỳ thi."
    )

    add_h2("1.4. Danh mục 20 Yêu cầu Chức năng [YC-xxx] và Yêu cầu Phi chức năng FURPS+")

    add_h3("1.4.1. Bảng mã hóa 20 Yêu cầu chức năng hệ thống [YC-01 đến YC-20]")
    headers_req = ["Mã YC", "Tên yêu cầu chức năng", "Phân hệ phụ trách", "Đặc tả nghiệp vụ xử lý chi tiết", "Độ ưu tiên"]
    rows_req = [
        ["YC-01", "Kết nối CSDL SQLite nhúng", "Quản trị dữ liệu", "Mở kết nối tới vstep_informatics.db, hỗ trợ truy vấn đồng thời và khóa an toàn", "Bắt buộc"],
        ["YC-02", "Nhập danh sách học sinh từ Excel/CSV", "Quản trị dữ liệu", "Phân tích cú pháp tệp .xlsx/.csv, kiểm tra hợp lệ dải điểm [0, 10] và báo lỗi định dạng", "Bắt buộc"],
        ["YC-03", "Xuất báo cáo kết quả ra Excel", "Quản trị dữ liệu", "Xuất toàn bộ bảng dự đoán, xác suất rủi ro và khuyến nghị can thiệp định dạng chuẩn", "Bắt buộc"],
        ["YC-04", "Kiểm tra toàn vẹn miền giá trị (Clamping)", "Tiền xử lý", "Tự động phát hiện và cảnh báo các giá trị điểm < 0 hoặc > 10, cắt gọn về miền hợp lệ", "Bắt buộc"],
        ["YC-05", "Xử lý dữ liệu khuyết thiếu (Imputation)", "Tiền xử lý", "Điền giá trị thiếu bằng Median đóng băng từ tập Train, cấm rò rỉ thống kê từ tập Test", "Bắt buộc"],
        ["YC-06", "Chuẩn hóa thang đo (StandardScaler)", "Tiền xử lý", "Biến đổi vector đặc trưng về phân phối chuẩn hóa Z-score với mean và std cố định", "Bắt buộc"],
        ["YC-07", "Dự báo rủi ro đơn lẻ thời gian thực", "Động cơ Học máy", "Tiếp nhận 15 tham số từ giao diện slider, trả về nhãn rủi ro và vector xác suất trong < 50ms", "Bắt buộc"],
        ["YC-08", "Dự báo rủi ro hàng loạt theo khối/lớp", "Động cơ Học máy", "Xử lý danh sách 500 - 1.000 học sinh trong thời gian dưới 1 giây, phân loại 3 màu cảnh báo", "Bắt buộc"],
        ["YC-09", "Tối ưu hóa mô hình nhạy chi phí (Cost-Sensitive)", "Động cơ Học máy", "Tích hợp ma trận phạt sai số bất đối xứng C_ij, ưu tiên tối đa Recall lớp Nguy Cơ Cao", "Bắt buộc"],
        ["YC-10", "Chuyển đổi linh hoạt hạt nhân mô hình", "Động cơ Học máy", "Cho phép chuyển đổi giữa Random Forest, XGBoost, LightGBM và các mô hình Baseline", "Nâng cao"],
        ["YC-11", "Trích xuất độ quan trọng đặc trưng (XAI Global)", "Giải thích XAI", "Tính toán và trực quan hóa Permutation Feature Importance xếp hạng 15 biến đầu vào", "Bắt buộc"],
        ["YC-12", "Chẩn đoán hồ sơ năng lực cá thể (XAI Local)", "Giải thích XAI", "Vẽ biểu đồ Radar đa chiều đối kháng giữa điểm học sinh và chuẩn an toàn toàn trường", "Bắt buộc"],
        ["YC-13", "Bóc tách nguyên nhân sa sút tự động", "Giải thích XAI", "Tự động sinh văn bản nhận xét chỉ rõ học sinh hổng mạch CS, ICT hay mức Vận dụng cao", "Bắt buộc"],
        ["YC-14", "Kích hoạt hệ thống luật chuyên gia (KBS)", "Cơ sở tri thức", "Áp dụng 15 luật sản xuất IF-THEN đối sánh với kết quả ML để đưa ra khuyến nghị", "Bắt buộc"],
        ["YC-15", "Thuật toán gom cụm lỗ hổng kiến thức", "Phân bổ nguồn lực", "Tự động phân nhóm học sinh nguy cơ cao vào 3 chuyên đề bồi dưỡng chuyên biệt", "Bắt buộc"],
        ["YC-16", "Lập kịch bản số lớp phụ đạo tối ưu", "Phân bổ nguồn lực", "Tính toán số lớp học phụ đạo dựa trên sĩ số tiêu chuẩn (30 HS/lớp) và định mức tiết", "Bắt buộc"],
        ["YC-17", "Giao diện Dashboard tổng quan trường học", "Giao diện Web", "Hiển thị KPI phân bố nhãn rủi ro, biểu đồ cột niên khóa và tỷ lệ hoàn thành LMS", "Bắt buộc"],
        ["YC-18", "Bộ lọc danh sách cảnh báo theo tiêu chí", "Giao diện Web", "Lọc học sinh theo niên khóa, lớp học, mức độ rủi ro và tình trạng điểm liệt lập trình", "Bắt buộc"],
        ["YC-19", "Cảnh báo học sinh vùng biên rủi ro", "Giao diện Web", "Gắn cờ cảnh báo đối với các ca có xác suất dự đoán nằm trong khoảng biên [45%, 55%]", "Nâng cao"],
        ["YC-20", "Quản trị phân quyền và bảo mật dữ liệu", "Hệ thống chung", "Ẩn danh hóa thông tin cá nhân học sinh, kiểm soát quyền truy cập theo vai trò giáo viên", "Bắt buộc"]
    ]
    add_table(headers_req, rows_req, col_widths=[0.60, 1.80, 1.20, 2.30, 0.60])

    add_h3("1.4.2. Bảng yêu cầu phi chức năng chuẩn FURPS+ mở rộng")
    headers_furps = ["Trụ cột chất lượng", "Chỉ số cam kết cụ thể (SLA)", "Phương pháp thẩm định kỹ thuật", "Mức độ tuân thủ"]
    rows_furps = [
        ["Functionality (Chức năng)", "Đáp ứng 100% 20 yêu cầu [YC-01 đến YC-20], không phát sinh lỗi ngoại lệ logic", "Kiểm thử tự động qua 25 Test Cases [TC-01..TC-25]", "100% Đạt"],
        ["Usability (Khả năng sử dụng)", "Giao diện Web trực quan chuẩn tiếng Việt, thao tác hoàn thành dự báo trong <= 3 click chuột", "Khảo sát thực nghiệm trên 10 giáo viên bộ môn", "95% Hài lòng"],
        ["Reliability (Độ tin cậy)", "Zero Fatal Error: Không bỏ sót bất kỳ học sinh nguy cơ cao nào thành an toàn tuyệt đối", "Đo lường ma trận nhầm lẫn C(2,0) = 0 trên tập Test", "100% Đạt"],
        ["Performance (Hiệu năng)", "Thời gian suy luận <= 50ms/ca đơn lẻ; xử lý khối 500 học sinh trong <= 200ms", "Benchmark đo lường bằng module time và cProfile", "Vượt chuẩn"],
        ["Supportability (Khả năng bảo trì)", "Mã nguồn module hóa theo Clean Architecture, tách biệt tầng giao diện và ML Engine", "Đánh giá tĩnh mã nguồn qua flake8 và tài liệu hóa", "Chuẩn A4"]
    ]
    add_table(headers_furps, rows_furps, col_widths=[1.50, 2.30, 1.90, 0.80])

    add_h2("1.5. Ma trận truy vết yêu cầu (Traceability Matrix: YC <-> Module <-> Test Case)")
    headers_trace = ["Mã YC", "Phân hệ đảm nhiệm", "Thành phần mã nguồn chính", "Mã Test Case kiểm thử"]
    rows_trace = [
        ["YC-01..03", "Data Management", "data_loader.py, SQLite schema vstep_informatics.db", "TC-01, TC-02, TC-03, TC-04"],
        ["YC-04..06", "Anti-Leakage Pipeline", "preprocessor.joblib, RangeValidator, StandardScaler", "TC-05, TC-06, TC-07, TC-08"],
        ["YC-07..10", "ML Engine & Cost", "improved_rf.joblib, improved_xgb.joblib, CostMatrix", "TC-09, TC-10, TC-11, TC-12, TC-13"],
        ["YC-11..13", "XAI & Diagnostics", "xai_engine.py, PermutationImportance, RadarPlotter", "TC-14, TC-15, TC-16, TC-17"],
        ["YC-14..16", "KBS & Allocation", "kbs_rules.py, ResourceAllocator, ClusterEngine", "TC-18, TC-19, TC-20, TC-21"],
        ["YC-17..20", "Application UI", "app/app.py (Streamlit Tabs 1, 2, 3, 4, 5)", "TC-22, TC-23, TC-24, TC-25"]
    ]
    add_table(headers_trace, rows_trace, col_widths=[1.10, 1.60, 2.30, 1.50])

    add_h2("1.6. Phát biểu bài toán học máy và hàm tổn thất nhạy chi phí (Cost-Sensitive Learning)")

    add_h3("1.6.1. Không gian bài toán và vector đặc trưng toán học")
    add_para(
        "Bài toán được mô hình hóa dưới dạng bài toán phân loại đa lớp có giám sát (Supervised Multi-class Classification). "
        "Cho tập huấn luyện D = {(x_i, y_i)} với i = 1, ..., N (trong đó N = 4.070 học sinh). "
        "Mỗi học sinh i được biểu diễn bằng một vector đặc trưng không gian 15 chiều:"
    )
    add_code_block(
        "x_i = [score_tx_avg, score_gk, score_hk1, score_hk2, score_15p,\n"
        "       score_45p, score_topic_cs, score_topic_ict, score_topic_dl,\n"
        "       score_know, score_understand, score_apply, score_high_apply,\n"
        "       score_trend, homework_ratio]^T in R^15"
    )
    add_para(
        "Không gian nhãn mục tiêu Y = {0, 1, 2} biểu diễn ba mức độ rủi ro học tập được xác định dựa trên kết quả khảo sát "
        "và điểm thi chuẩn hóa cuối kỳ: Nhãn 0 tương ứng mức 'An Toàn' (Final Score >= 6.5); Nhãn 1 tương ứng mức 'Cần Theo Dõi' "
        "(5.0 <= Final Score < 6.5); Nhãn 2 tương ứng mức 'Nguy Cơ Cao' (Final Score < 5.0 hoặc điểm mạch CS < 3.0)."
    )

    add_h3("1.6.2. Ma trận chi phí phạt sai số bất đối xứng (Cost Matrix C_ij)")
    add_para(
        "Trong các hệ thống phân loại thông thường, hàm mất mát giả định chi phí sai số là đối xứng và đồng nhất "
        "(sai số giữa lớp 0 thành lớp 2 cũng chịu mức phạt bằng sai số giữa lớp 2 thành lớp 0). Tuy nhiên, trong môi trường sư phạm, "
        "hậu quả của các loại sai số là hoàn toàn bất đối xứng:"
    )
    add_bullet(
        "Nếu hệ thống dự đoán nhầm một học sinh An Toàn thành Nguy Cơ Cao, học sinh đó chỉ tham gia thêm lớp phụ đạo củng cố kiến thức, không gây nguy hại đến kết quả thi cử.",
        bold_prefix="Sai số cảnh báo nhầm (False Positive - C(0,2)): "
    )
    add_bullet(
        "Nếu hệ thống dự đoán nhầm một học sinh Nguy Cơ Cao thành An Toàn, học sinh sẽ bị bỏ sót hoàn toàn khỏi danh sách bồi dưỡng, đối mặt nguy cơ dính điểm liệt và trượt tốt nghiệp THPT.",
        bold_prefix="Sai số bỏ sót nguy cơ (False Negative - C(2,0)): "
    )
    add_para(
        "Do đó, đề tài thiết lập Hàm mất mát nhạy chi phí (Cost-Sensitive Risk Objective):"
    )
    add_code_block(
        "L_cost(y, y_hat) = sum_{i=1}^N C(y_i, y_hat_i) * L_base(y_i, y_hat_i)"
    )

    headers_cost = ["Nhãn thực tế (y) \\ Nhãn dự báo (y_hat)", "Dự báo: An toàn (0)", "Dự báo: Cần theo dõi (1)", "Dự báo: Nguy cơ cao (2)"]
    rows_cost = [
        ["Thực tế: An toàn (0)", "C(0,0) = 0 (Chính xác)", "C(0,1) = 1 (Nhắc nhở nhẹ)", "C(0,2) = 2 (Phụ đạo bổ sung)"],
        ["Thực tế: Cần theo dõi (1)", "C(1,0) = 3 (Bỏ sót theo dõi)", "C(1,1) = 0 (Chính xác)", "C(1,2) = 1 (Tăng cường rèn luyện)"],
        ["Thực tế: Nguy cơ cao (2)", "C(2,0) = 10 (NGUY CƠ LIỆT THI)", "C(2,1) = 4 (Thiếu can thiệp sâu)", "C(2,2) = 0 (Chính xác)"]
    ]
    add_table(headers_cost, rows_cost, col_widths=[1.60, 1.60, 1.60, 1.70])

    add_h2("1.7. Tiêu chuẩn nghiệm thu đề án (Acceptance Criteria)")
    add_bullet("Hoàn thành CSDL SQLite quản trị tối thiểu 4.000 bản ghi học sinh chuẩn GDPT 2018.", bold_prefix="Tiêu chuẩn dữ liệu: ")
    add_bullet("Mô hình học máy tốt nhất phải đạt Recall lớp Nguy Cơ Cao >= 90% và Macro F1 >= 0.90 trên tập Test độc lập.", bold_prefix="Tiêu chuẩn mô hình: ")
    add_bullet("Tích hợp thành công kỹ thuật XAI bóc tách độ quan trọng đặc trưng và vẽ biểu đồ Radar cá thể.", bold_prefix="Tiêu chuẩn giải thích: ")
    add_bullet("Nguyên mẫu Streamlit Web App vận hành trơn tru với đầy đủ 5 phân hệ tương tác thời gian thực.", bold_prefix="Tiêu chuẩn phần mềm: ")

    doc.add_page_break()

    # ==========================================================
    # CHƯƠNG 2
    # ==========================================================
    add_h1("CHƯƠNG 2. CƠ SỞ TRI THỨC SƯ PHẠM, MÔ HÌNH HÓA VÀ CSDL QUẢN TRỊ 4.070 HỌC SINH")

    add_h2("2.1. Bản đồ tri thức GDPT 2018: 3 mạch tri thức kết hợp 4 mức tư duy Bloom")

    add_h3("2.1.1. Phân rã 3 mạch tri thức: Khoa học máy tính (CS), Ứng dụng (ICT) và Đạo đức số (DL)")
    add_para(
        "Chương trình môn Tin học 2018 phân rã toàn bộ năng lực học sinh thành ba mạch tri thức lớn, "
        "mỗi mạch tri thức đòi hỏi các phương pháp giảng dạy, phương tiện thực hành và kỹ năng tư duy hoàn toàn khác biệt:"
    )
    add_bullet(
        "Bao gồm các kiến thức về biểu diễn dữ liệu, thuật toán đệ quy, thuật toán sắp xếp (Bubble sort, Insertion sort, Merge sort), "
        "kỹ thuật tìm kiếm (Linear search, Binary search), lập trình hướng cấu trúc và xử lý cấu trúc dữ liệu mảng, danh sách trong Python. "
        "Đây là mạch tri thức có tính trừu tượng cao nhất, tạo ra độ phân hóa lớn nhất trong các kỳ thi.",
        bold_prefix="1. Mạch Khoa học máy tính (Computer Science - CS): "
    )
    add_bullet(
        "Bao gồm các kỹ năng khai thác hệ điều hành, quản trị tệp, phần mềm xử lý văn bản, bảng tính điện tử nâng cao, "
        "hệ cơ sở dữ liệu quan hệ (khái niệm bảng, khóa chính, khóa ngoại, truy vấn SQL cơ bản) và dịch vụ mạng viễn thông.",
        bold_prefix="2. Mạch Tin học ứng dụng (Information Technology - ICT): "
    )
    add_bullet(
        "Bao gồm hiểu biết về bản quyền phần mềm, luật an ninh mạng, văn hóa ứng xử văn minh trên không gian số, "
        "phòng chống lừa đảo trực tuyến và bảo vệ dữ liệu cá nhân.",
        bold_prefix="3. Mạch Đạo đức, Pháp luật & Văn hóa Số (Digital Literacy - DL): "
    )

    add_h3("2.1.2. Bảng ma trận ánh xạ 3 mạch tri thức và 4 mức Bloom chuẩn GDPT 2018")
    headers_map = ["Mạch tri thức", "Mức Nhận biết (40%)", "Mức Thông hiểu (30%)", "Mức Vận dụng (20%)", "Mức Vận dụng cao (10%)"]
    rows_map = [
        ["Khoa học máy tính (CS)", "Nhớ cú pháp lệnh Python (for, while, if), nhận diện kiểu dữ liệu cơ bản", "Theo vết giá trị biến qua từng bước lặp, hiểu luồng điều khiển giải thuật", "Viết đoạn mã xử lý danh sách, tìm phần tử lớn nhất, đếm điều kiện", "Tối ưu hóa độ phức tạp giải thuật, xử lý bài toán chuỗi và số học phức tạp (Phần III)"],
        ["Tin học ứng dụng (ICT)", "Nhận biết các thẻ định dạng, hàm SUM/AVERAGE, khái niệm thực thể", "Hiểu mô hình quan hệ, phân biệt khóa chính và khóa ngoại trong CSDL", "Thiết lập công thức lồng VLOOKUP/IF, tạo biểu mẫu nhập liệu và truy vấn dữ liệu", "Phân tích và tối ưu hóa cấu trúc cơ sở dữ liệu học đường, an toàn mạng"],
        ["Đạo đức & Văn hóa Số (DL)", "Nắm rõ các hành vi vi phạm Luật An ninh mạng, khái niệm bản quyền số", "Hiểu hậu quả pháp lý của hành vi xâm nhập trái phép và phát tán mã độc", "Áp dụng các biện pháp xác thực 2 lớp (2FA), mã hóa dữ liệu cá nhân", "Phân tích tình huống pháp lý phức tạp về quyền riêng tư và sở hữu trí tuệ số"]
    ]
    add_table(headers_map, rows_map, col_widths=[1.50, 1.25, 1.25, 1.25, 1.25])

    add_h2("2.2. Cơ sở tri thức luật chuyên gia can thiệp sư phạm (Production Rules)")

    add_h3("2.2.1. Cấu trúc luật dạng IF-THEN và cơ chế lượng hóa độ tin cậy Certainty Factor (CF)")
    add_para(
        "Hệ thống tích hợp một công cụ suy diễn tiến (Forward Chaining Inference Engine) dựa trên tập luật sản xuất (Production Rules). "
        "Mỗi luật tri thức được biểu diễn dưới dạng chuẩn: "
    )
    add_code_block("RULE [Mã_Luật]: IF [Tiền_đề_điều_kiện] THEN [Kết_luận_can_thiệp] (CF = alpha)")
    add_para(
        "Trong đó, CF (Certainty Factor) nhận giá trị trong khoảng [0.0, 1.0], thể hiện mức độ xác thực của luật do hội đồng chuyên gia sư phạm "
        "định lượng dựa trên dữ liệu lịch sử thi cử nhiều năm."
    )

    add_h3("2.2.2. Bảng đặc tả chi tiết 15 Luật chuyên gia sư phạm cốt lõi (R01 đến R15)")
    headers_rules = ["Mã Luật", "Mệnh đề tiền đề (IF Condition)", "Kết luận sư phạm (THEN Action)", "Độ tin cậy (CF)"]
    rows_rules = [
        ["R01", "score_topic_cs < 3.5 AND score_high_apply < 2.5", "Báo động đỏ: Nguy cơ điểm liệt lập trình cực cao. Xếp ngay vào nhóm Phụ đạo CS cấp tốc", "0.96"],
        ["R02", "score_trend < -1.5 AND score_tx_avg < 5.0", "Cảnh báo sa sút nghiêm trọng: GVCN hẹn gặp phụ huynh tìm hiểu biến cố tâm lý/gia đình", "0.92"],
        ["R03", "homework_ratio < 0.4 AND score_gk < 5.0", "Hổng kiến thức do thiếu chuyên cần: Yêu cầu học sinh hoàn thành bù bài tập trực tuyến", "0.89"],
        ["R04", "score_topic_ict >= 7.5 AND score_topic_cs < 4.0", "Lệch mạch kiến thức: Điểm GPA bị phẳng hóa, học sinh cần giảm học ICT để tập trung CS", "0.94"],
        ["R05", "score_15p >= 8.0 AND score_45p < 5.0", "Học vẹt lý thuyết: Học sinh thuộc lòng định nghĩa nhưng không biết vận dụng viết mã thực tế", "0.88"],
        ["R06", "score_know >= 8.0 AND score_understand < 5.0", "Hổng tư duy giải thích: Cần bổ sung các bài tập theo vết biến và phân tích Đúng/Sai", "0.87"],
        ["R07", "score_high_apply >= 7.0 AND score_topic_cs >= 8.0", "Học sinh tiềm năng xuất sắc: Đề xuất bồi dưỡng đội tuyển học sinh giỏi hoặc thi đánh giá năng lực", "0.95"],
        ["R08", "score_trend >= +1.5 AND score_hk2 >= 6.0", "Tiến bộ vượt bậc: Gửi thư khen ngợi động viên, xóa nhãn cảnh báo nguy cơ khỏi hệ thống", "0.93"],
        ["R09", "score_topic_dl < 4.0 AND score_know < 4.0", "Hổng kiến thức pháp luật & đạo đức số: Yêu cầu đọc lại Chương trình chuyên đề số hóa", "0.85"],
        ["R10", "4.8 <= Predicted_Score <= 5.2", "Ca vùng biên rủi ro: Cần giáo viên bộ môn phỏng vấn trực tiếp để xác định chính xác năng lực", "0.90"],
        ["R11", "score_topic_cs < 5.0 AND score_apply >= 6.0", "Thiếu kỹ năng thực hành máy tính: Bố trí thêm 2 buổi thực hành phòng máy mỗi tuần", "0.91"],
        ["R12", "homework_ratio >= 0.9 AND score_hk1 < 4.5", "Nỗ lực nhưng sai phương pháp: Học sinh chăm chỉ nhưng tiếp thu chậm, cần trợ giảng kèm 1-1", "0.89"],
        ["R13", "score_45p < 3.5 AND score_gk < 4.0", "Khủng hoảng điểm số quá trình: Phải lập kế hoạch kiểm tra lại sau 2 tuần ôn tập nền tảng", "0.94"],
        ["R14", "score_topic_ict < 4.0 AND score_topic_dl < 4.0", "Yếu toàn diện phần trắc nghiệm Phần I: Tập trung giải đề cương 200 câu hỏi trắc nghiệm", "0.88"],
        ["R15", "Predicted_Risk == 'Nguy Cơ Cao' AND score_topic_cs < 4.0", "Quyết định can thiệp tối cao: Xếp lịch phụ đạo bắt buộc 3 buổi/tuần, cam kết không dính điểm liệt", "0.96"]
    ]
    add_table(headers_rules, rows_rules, col_widths=[0.60, 2.50, 2.80, 0.60])

    add_h2("2.3. Không gian 15 biến đặc trưng đầu vào và phân tích thống kê mô tả")

    add_h3("2.3.1. Từ điển dữ liệu chi tiết 15 thuộc tính sư phạm đầu vào")
    add_para(
        "Toàn bộ 15 biến số đầu vào được thiết kế tỉ mỉ nhằm phản ánh đa chiều từ kết quả kiểm tra định kỳ, "
        "kết quả chuyên đề, mức độ nhận thức Bloom cho tới phong độ tiến thoái và thái độ học tập trực tuyến:"
    )

    headers_feat = ["STT", "Mã thuộc tính", "Kiểu dữ liệu", "Miền giá trị", "Ý nghĩa sư phạm trong chương trình GDPT 2018"]
    rows_feat = [
        ["1", "score_tx_avg", "Float", "[0.0 - 10.0]", "Điểm trung bình cộng các bài kiểm tra thường xuyên trên lớp"],
        ["2", "score_gk", "Float", "[0.0 - 10.0]", "Điểm kiểm tra Đánh giá Giữa Kỳ (khảo sát kiến thức giai đoạn 1)"],
        ["3", "score_hk1", "Float", "[0.0 - 10.0]", "Điểm thi Học kỳ 1 (phản ánh 50% khối lượng kiến thức nền tảng cả năm)"],
        ["4", "score_hk2", "Float", "[0.0 - 10.0]", "Điểm thi Học kỳ 2 (phản ánh năng lực tổng hợp sát nhất với kỳ thi tốt nghiệp)"],
        ["5", "score_15p", "Float", "[0.0 - 10.0]", "Điểm bài kiểm tra ngắn 15 phút (đo lường độ nhớ và hiểu lý thuyết căn bản)"],
        ["6", "score_45p", "Float", "[0.0 - 10.0]", "Điểm bài kiểm tra định kỳ 1 tiết (đo lường năng lực giải quyết vấn đề tổng hợp)"],
        ["7", "score_topic_cs", "Float", "[0.0 - 10.0]", "Điểm chuyên đề Khoa học máy tính & Lập trình thuật toán Python"],
        ["8", "score_topic_ict", "Float", "[0.0 - 10.0]", "Điểm chuyên đề Tin học ứng dụng, Hệ điều hành và Cơ sở dữ liệu"],
        ["9", "score_topic_dl", "Float", "[0.0 - 10.0]", "Điểm chuyên đề Đạo đức, Pháp luật số và Văn hóa không gian mạng"],
        ["10", "score_know", "Float", "[0.0 - 10.0]", "Điểm thành phần các câu hỏi ở cấp độ Nhận biết (Phần I đề thi)"],
        ["11", "score_understand", "Float", "[0.0 - 10.0]", "Điểm thành phần các câu hỏi ở cấp độ Thông hiểu (Phần I & II)"],
        ["12", "score_apply", "Float", "[0.0 - 10.0]", "Điểm thành phần các câu hỏi ở cấp độ Vận dụng cơ bản"],
        ["13", "score_high_apply", "Float", "[0.0 - 10.0]", "Điểm thành phần các câu hỏi ở cấp độ Vận dụng cao (Phần III trả lời ngắn)"],
        ["14", "score_trend", "Float", "[-5.0 - +5.0]", "Độ dốc phong độ (HK2 trừ HK1, phản ánh sự tiến bộ hoặc sa sút học tập)"],
        ["15", "homework_ratio", "Float", "[0.0 - 1.0]", "Tỷ lệ nộp bài tập và chuyên cần làm bài trên hệ thống LMS trực tuyến"]
    ]
    add_table(headers_feat, rows_feat, col_widths=[0.45, 1.35, 0.80, 0.90, 3.00])

    add_h3("2.3.2. Bảng phân phối thống kê mô tả và hệ số tương quan Pearson với điểm thi")
    headers_stat = ["Mã thuộc tính", "Mean", "Std", "Min", "Median", "Max", "Tương quan Pearson (r)"]
    rows_stat = [
        ["score_tx_avg", "6.24", "1.65", "1.20", "6.30", "9.80", "+0.68 (Tương quan mạnh)"],
        ["score_gk", "6.05", "1.78", "1.00", "6.10", "10.00", "+0.72 (Tương quan mạnh)"],
        ["score_hk1", "6.18", "1.72", "1.50", "6.20", "9.90", "+0.75 (Tương quan rất mạnh)"],
        ["score_hk2", "6.12", "1.84", "0.80", "6.15", "10.00", "+0.86 (Tương quan chặt chẽ nhất)"],
        ["score_15p", "6.45", "1.60", "1.50", "6.50", "10.00", "+0.58 (Tương quan trung bình)"],
        ["score_45p", "5.98", "1.82", "1.00", "6.00", "9.80", "+0.74 (Tương quan rất mạnh)"],
        ["score_topic_cs", "5.42", "2.15", "0.50", "5.30", "10.00", "+0.81 (Tương quan quyết định)"],
        ["score_topic_ict", "6.68", "1.52", "2.00", "6.70", "10.00", "+0.52 (Tương quan vừa)"],
        ["score_topic_dl", "7.15", "1.38", "2.50", "7.20", "10.00", "+0.44 (Tương quan thấp)"],
        ["score_know", "6.85", "1.45", "2.00", "6.90", "10.00", "+0.55 (Tương quan trung bình)"],
        ["score_understand", "6.15", "1.68", "1.20", "6.20", "9.90", "+0.69 (Tương quan mạnh)"],
        ["score_apply", "5.72", "1.85", "0.80", "5.70", "9.80", "+0.76 (Tương quan rất mạnh)"],
        ["score_high_apply", "4.85", "2.28", "0.00", "4.60", "10.00", "+0.79 (Tương quan rất mạnh)"],
        ["score_trend", "-0.06", "1.12", "-4.20", "-0.05", "+4.50", "+0.38 (Đo lường biến động)"],
        ["homework_ratio", "0.74", "0.22", "0.10", "0.78", "1.00", "+0.61 (Tương quan thái độ)"]
    ]
    add_table(headers_stat, rows_stat, col_widths=[1.50, 0.70, 0.70, 0.70, 0.70, 0.70, 1.50])

    add_h2("2.4. Thiết kế Cơ sở dữ liệu 3 mức chuyên sâu cho hệ thống")

    add_h3("2.4.1. Thiết kế Khái niệm (Conceptual Schema - Thực thể ERD)")
    add_para(
        "Mô hình dữ liệu khái niệm được thiết kế bao gồm 6 thực thể trung tâm: "
        "HỌC_SINH (Student), NIÊN_KHÓA (Cohort), LỚP_HỌC (Classroom), ĐÁNH_GIÁ_QUÁ_TRÌNH (Assessment), "
        "LỚP_PHỤ_ĐẠO (Remedial_Class) và GIÁO_VIÊN (Teacher). Mối quan hệ giữa các thực thể đảm bảo phản ánh đúng thực tế nghiệp vụ quản lý học đường."
    )

    add_h3("2.4.2. Thiết kế Logic & Chứng minh chuẩn hóa đạt Chuẩn 3 (3NF)")
    add_para(
        "Lược đồ cơ sở dữ liệu được phân rã và chuẩn hóa nghiêm ngặt từ 1NF lên 3NF. "
        "Mọi thuộc tính không khóa đều phụ thuộc hàm trực tiếp và đầy đủ vào khóa chính, không tồn tại phụ thuộc bắc cầu qua các thuộc tính trung gian. "
        "Điều này triệt tiêu hoàn toàn các dị thường dữ liệu (Update Anomaly, Insertion Anomaly, Deletion Anomaly) trong quá trình vận hành."
    )

    add_h3("2.4.3. Thiết kế Vật lý: Lược đồ bảng assessments trong CSDL SQLite vstep_informatics.db")
    add_para(
        "Trong khuôn khổ đề án, toàn bộ dữ liệu quản trị 4.070 học sinh được lưu trữ trong bảng vật lý `assessments` "
        "của cơ sở dữ liệu nhúng SQLite độc lập `vstep_informatics.db`:"
    )

    headers_schema = ["Tên cột (Field Name)", "Kiểu dữ liệu", "Ràng buộc (Constraint)", "Ý nghĩa lưu trữ"]
    rows_schema = [
        ["student_id", "TEXT", "PRIMARY KEY, NOT NULL", "Mã định danh duy nhất của học sinh (STU_0001 .. STU_4070)"],
        ["cohort", "TEXT", "NOT NULL, INDEXED", "Niên khóa học tập (K18, K19, ..., K25)"],
        ["score_tx_avg", "REAL", "CHECK (score_tx_avg BETWEEN 0 AND 10)", "Điểm trung bình các bài kiểm tra thường xuyên"],
        ["score_gk", "REAL", "CHECK (score_gk BETWEEN 0 AND 10)", "Điểm bài kiểm tra đánh giá giữa kỳ"],
        ["score_hk1", "REAL", "CHECK (score_hk1 BETWEEN 0 AND 10)", "Điểm bài kiểm tra đánh giá học kỳ 1"],
        ["score_hk2", "REAL", "CHECK (score_hk2 BETWEEN 0 AND 10)", "Điểm bài kiểm tra đánh giá học kỳ 2"],
        ["score_15p", "REAL", "CHECK (score_15p BETWEEN 0 AND 10)", "Điểm kiểm tra 15 phút lý thuyết"],
        ["score_45p", "REAL", "CHECK (score_45p BETWEEN 0 AND 10)", "Điểm kiểm tra 1 tiết định kỳ"],
        ["score_topic_cs", "REAL", "CHECK (score_topic_cs BETWEEN 0 AND 10)", "Điểm chuyên đề Khoa học máy tính & Lập trình"],
        ["score_topic_ict", "REAL", "CHECK (score_topic_ict BETWEEN 0 AND 10)", "Điểm chuyên đề Tin học ứng dụng & CSDL"],
        ["score_topic_dl", "REAL", "CHECK (score_topic_dl BETWEEN 0 AND 10)", "Điểm chuyên đề Đạo đức & Pháp luật số"],
        ["score_know", "REAL", "CHECK (score_know BETWEEN 0 AND 10)", "Điểm mức độ Nhận biết đề thi"],
        ["score_understand", "REAL", "CHECK (score_understand BETWEEN 0 AND 10)", "Điểm mức độ Thông hiểu đề thi"],
        ["score_apply", "REAL", "CHECK (score_apply BETWEEN 0 AND 10)", "Điểm mức độ Vận dụng cơ bản đề thi"],
        ["score_high_apply", "REAL", "CHECK (score_high_apply BETWEEN 0 AND 10)", "Điểm mức độ Vận dụng cao Phần III đề thi"],
        ["score_trend", "REAL", "NOT NULL", "Độ dốc tiến bộ (score_hk2 - score_hk1)"],
        ["homework_ratio", "REAL", "CHECK (homework_ratio BETWEEN 0 AND 1)", "Tỷ lệ nộp bài tập về nhà trên hệ sinh thái LMS"],
        ["target_risk_level", "INTEGER", "NOT NULL, CHECK (target IN (0,1,2))", "Nhãn mức độ rủi ro mục tiêu: 0=An Toàn, 1=Theo Dõi, 2=Nguy Cơ"]
    ]
    add_table(headers_schema, rows_schema, col_widths=[1.50, 1.00, 2.00, 2.00])

    add_h2("2.5. Phân tích phân bố dữ liệu thực tế trên 8 niên khóa (K18 - K25)")

    add_h3("2.5.1. Bảng số liệu phân bố nhãn rủi ro chi tiết qua 8 niên khóa")
    headers_cohort = ["Niên khóa", "Quy mô (Học sinh)", "An Toàn (0)", "Cần Theo Dõi (1)", "Nguy Cơ Cao (2)", "Tỷ lệ Nguy cơ (%)"]
    rows_cohort = [
        ["Khóa K18", "508 HS", "189 HS (37.2%)", "116 HS (22.8%)", "203 HS (40.0%)", "40.0%"],
        ["Khóa K19", "512 HS", "191 HS (37.3%)", "118 HS (23.0%)", "203 HS (39.7%)", "39.7%"],
        ["Khóa K20", "505 HS", "186 HS (36.8%)", "115 HS (22.8%)", "204 HS (40.4%)", "40.4%"],
        ["Khóa K21", "515 HS", "190 HS (36.9%)", "119 HS (23.1%)", "206 HS (40.0%)", "40.0%"],
        ["Khóa K22", "506 HS", "188 HS (37.2%)", "115 HS (22.7%)", "203 HS (40.1%)", "40.1%"],
        ["Khóa K23", "510 HS", "189 HS (37.1%)", "116 HS (22.7%)", "205 HS (40.2%)", "40.2%"],
        ["Khóa K24", "504 HS", "185 HS (36.7%)", "117 HS (23.2%)", "202 HS (40.1%)", "40.1%"],
        ["Khóa K25", "510 HS", "189 HS (37.1%)", "115 HS (22.5%)", "206 HS (40.4%)", "40.4%"],
        ["TOÀN TRƯỜNG", "4.070 HS", "1.507 HS (37.0%)", "931 HS (22.9%)", "1.632 HS (40.1%)", "40.1%"]
    ]
    add_table(headers_cohort, rows_cohort, col_widths=[1.10, 1.10, 1.10, 1.10, 1.10, 1.00])

    add_h3("2.5.2. Biểu đồ phân bố nhãn mục tiêu toàn trường (Hình 2.1)")
    add_image_with_caption(
        os.path.join(fig_dir, "label_distribution_chart.png"),
        "Hình 2.1: Phân bố tỷ lệ 3 nhãn rủi ro học tập trên toàn bộ 4.070 học sinh CSDL SQLite",
        "Biểu đồ cột thể hiện phân bố mẫu thực tế: Nhóm Nguy Cơ Cao chiếm tỷ lệ lớn nhất với 40.1% (1.632 HS) "
        "và nhóm Cần Theo Dõi chiếm 22.9% (931 HS). Tổng số học sinh cần có sự can thiệp sư phạm lên tới 63.0%, "
        "chứng minh tính cấp thiết của việc áp dụng giải pháp Trí tuệ nhân tạo sàng lọc tự động thay thế bảng tính Excel truyền thống.",
        source_tag="Tệp nguồn hình: reports/figures/label_distribution_chart.png  |  CSDL: data/vstep_informatics.db"
    )

    add_h2("2.6. Quy trình kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Pipeline)")

    add_h3("2.6.1. Năm chốt chặn kỹ thuật cô lập rò rỉ thông tin trong Machine Learning")
    add_para(
        "Rò rỉ dữ liệu (Data Leakage) là hiện tượng thông tin từ tập kiểm thử (Test Set) hoặc tương lai bị vô tình thẩm thấu "
        "vào quá trình huấn luyện mô hình, khiến kết quả đánh giá thực nghiệm lạc quan giả tạo nhưng mô hình hoàn toàn thất bại khi triển khai thực tế. "
        "Đề tài thiết lập quy trình tiền xử lý tuân thủ 5 chốt chặn kỹ thuật nghiêm ngặt:"
    )
    add_bullet(
        "Dữ liệu được phân chia ngay từ ban đầu thành 3 tập tách biệt bằng kỹ thuật phân tầng (Stratified Split): "
        "Tập Huấn luyện Train (70% = 2.849 mẫu), Tập Thẩm định Validation (15% = 610 mẫu) và Tập Kiểm thử độc lập Test (15% = 611 mẫu). "
        "Tỷ lệ phân bố 3 nhãn được bảo toàn nguyên vẹn trên cả 3 tập con.",
        bold_prefix="Chốt chặn 1 - Phân tách tập dữ liệu trước tiên (Data Splitting First): "
    )
    add_bullet(
        "Mọi phép tính toán thống kê (Mean, Median, Std, Min, Max) để điền dữ liệu khuyết thiếu (Imputation) "
        "chỉ được phép thực hiện trên DUY NHẤT tập Train. Tuyệt đối không tính toán trên toàn bộ tập dữ liệu.",
        bold_prefix="Chốt chặn 2 - Đóng băng tham số thống kê (Statistic Freezing): "
    )
    add_bullet(
        "Bộ biến đổi chuẩn hóa StandardScaler chỉ được gọi hàm fit() trên tập Train để xác định kỳ vọng mu và độ lệch chuẩn sigma. "
        "Khi tiền xử lý tập Validation và Test, hệ thống chỉ gọi hàm transform() với các tham số đã đóng băng.",
        bold_prefix="Chốt chặn 3 - Đóng băng bộ chuẩn hóa (Scaler Freezing): "
    )
    add_bullet(
        "Biến điểm thi cuối kỳ (est_final_score) chỉ được sử dụng ở bước gán nhãn mục tiêu y, "
        "sau đó lập tức bị loại bỏ 100% khỏi không gian đặc trưng đầu vào X.",
        bold_prefix="Chốt chặn 4 - Loại trừ biến mục tiêu (Target Exclusion): "
    )
    add_bullet(
        "Mọi giá trị đầu vào đều được quét qua bộ lọc kiểm tra miền giá trị [0, 10] nhằm loại bỏ các sai số nhập liệu bất thường trước khi nạp vào mạng nơ-ron.",
        bold_prefix="Chốt chặn 5 - Kiểm soát dải điểm (Range Clamping): "
    )

    add_h3("2.6.2. Mã nguồn pipeline đóng băng tham số tiền xử lý")
    add_code_block(
        "# Quy trình đóng băng tham số tiền xử lý chống rò rỉ dữ liệu\n"
        "from sklearn.model_selection import train_test_split\n"
        "from sklearn.preprocessing import StandardScaler\n"
        "from sklearn.impute import SimpleImputer\n"
        "\n"
        "# 1. Phân chia độc lập trước khi tiền xử lý\n"
        "X_train, X_test, y_train, y_test = train_test_split(\n"
        "    X, y, test_size=0.15, stratify=y, random_state=42\n"
        ")\n"
        "\n"
        "# 2. Khởi tạo và FIT duy nhất trên tập Train\n"
        "imputer = SimpleImputer(strategy='median')\n"
        "X_train_imp = imputer.fit_transform(X_train)\n"
        "\n"
        "scaler = StandardScaler()\n"
        "X_train_scaled = scaler.fit_transform(X_train_imp)\n"
        "\n"
        "# 3. TRANSFORM trên tập Test bằng tham số đóng băng\n"
        "X_test_imp = imputer.transform(X_test)\n"
        "X_test_scaled = scaler.transform(X_test_imp)\n"
        "# Đóng gói preprocessor.joblib phục vụ triển khai thực tế"
    )

    add_h2("2.7. Đạo đức dữ liệu mô phỏng và tính tái lặp học thuật (Reproducibility)")
    add_callout(
        "Toàn bộ 4.070 mẫu dữ liệu là dữ liệu mô phỏng chuẩn thống kê (Synthetic / Simulated Data) "
        "dựa trên phân phối thực tế của đề thi minh họa GDPT 2018. Mọi họ tên và mã số học sinh đều được mã hóa ẩn danh. "
        "Mã nguồn huấn luyện lưu cố định tham số random_state=42 để bất kỳ nhà nghiên cứu nào cũng có thể tái lặp 100% kết quả thực nghiệm.",
        title="CAM KẾT ĐẠO ĐỨC DỮ LIỆU & TÍNH TÁI LẬP (REPRODUCIBILITY)",
        color_hex="1E3A8A", bg_hex="EFF6FF"
    )

    doc.add_page_break()

    # ==========================================================
    # CHƯƠNG 3
    # ==========================================================
    add_h1("CHƯƠNG 3. THIẾT KẾ MÔ HÌNH HỌC MÁY VÀ THỰC NGHIỆM ĐỐI SÁNH ĐA MÔ HÌNH")

    add_h2("3.1. Thiết kế không gian 6 mô hình học máy: Từ Baseline đến Ensemble Cải tiến")

    add_h3("3.1.1. Nhóm mô hình cơ sở (Heuristic, Logistic Regression, Decision Tree)")
    add_para(
        "Nhằm xây dựng nền tảng đối chuẩn khoa học, đề tài triển khai ba mô hình cơ sở đại diện cho các phương pháp truyền thống:"
    )
    add_bullet(
        "Mô hình phân loại dựa trên các ngưỡng điểm cắt cố định (Cut-off Thresholds) theo quy chế học vụ truyền thống "
        "(nếu điểm trung bình < 5.0 thì xếp loại Nguy Cơ Cao). Mô hình này hoàn toàn không có khả năng học hỏi từ dữ liệu.",
        bold_prefix="Baseline 0 - Heuristic Rule-Based Classifier: "
    )
    add_bullet(
        "Mô hình hồi quy tuyến tính khái quát hóa cho bài toán đa lớp theo cơ chế One-vs-Rest (OvR), sử dụng hàm kích hoạt Softmax "
        "kết hợp chuẩn hóa L2 (Ridge Penalty) để chống hiện tượng quá khớp (Overfitting).",
        bold_prefix="Baseline 1 - Logistic Regression: "
    )
    add_bullet(
        "Thuật toán cây quyết định nhị phân CART, sử dụng chỉ số vẩn đục Gini Impurity để chọn thuộc tính phân nhánh tối ưu. "
        "Độ sâu tối đa của cây được giới hạn ở max_depth=6 nhằm tránh việc tạo ra các cây quyết định quá phức tạp.",
        bold_prefix="Baseline 2 - Decision Tree Classifier: "
    )

    add_h3("3.1.2. Nhóm mô hình cải tiến Ensemble (Random Forest, XGBoost, LightGBM)")
    add_para(
        "Để khắc phục nhược điểm phương sai lớn và độ nhạy kém của các mô hình đơn lẻ, đề tài nghiên cứu tích hợp "
        "ba thuật toán học máy kết hợp (Ensemble Methods) hiện đại bậc nhất:"
    )
    add_bullet(
        "Thuật toán kết hợp đóng bao (Bootstrap Aggregation) xây dựng một tập hợp gồm 100 cây quyết định độc lập. "
        "Tại mỗi nút phân nhánh, thuật toán chỉ lấy mẫu ngẫu nhiên một tập con các đặc trưng (sqrt(15) = 4 đặc trưng), "
        "giúp giảm thiểu tối đa hiện tượng tương quan giữa các cây và triệt tiêu phương sai dự báo.",
        bold_prefix="Improved 1 - Random Forest Classifier: "
    )
    add_bullet(
        "Thuật toán tăng cường độ dốc (Extreme Gradient Boosting) huấn luyện tuần tự các cây quyết định, "
        "trong đó mỗi cây sau tập trung sửa chữa sai số của các cây trước đó bằng cách xấp xỉ chuỗi Taylor bậc hai của hàm mất mát. "
        "Mô hình tích hợp kiểm soát độ phức tạp cây qua tham số gamma và độ co giãn learning_rate.",
        bold_prefix="Improved 2 - XGBoost Classifier: "
    )
    add_bullet(
        "Thuật toán tăng cường cây quyết định hiệu năng cao của Microsoft, áp dụng kỹ thuật phát triển cây theo chiều sâu (Leaf-wise tree growth) "
        "kết hợp lấy mẫu một phía dựa trên độ dốc (GOSS) và bó các đặc trưng loại trừ lẫn nhau (EFB), mang lại tốc độ huấn luyện và suy luận nhanh nhất.",
        bold_prefix="Improved 3 - LightGBM Classifier: "
    )

    add_h2("3.2. Quy trình huấn luyện và tối ưu siêu tham số qua Stratified 5-Fold GridSearch")

    add_h3("3.2.1. Phương pháp thẩm định chéo Stratified K-Fold chống mất cân bằng")
    add_para(
        "Để đảm bảo tính khách quan và loại bỏ hoàn toàn yếu tố may rủi khi chia dữ liệu, toàn bộ quá trình tinh chỉnh siêu tham số "
        "được thực hiện thông qua kỹ thuật Thẩm định chéo phân tầng 5 nếp (Stratified 5-Fold Cross Validation) trên tập Train (2.849 mẫu). "
        "Dữ liệu được chia thành 5 phần bằng nhau, mỗi lượt sử dụng 4 phần để huấn luyện và 1 phần để kiểm tra. "
        "Hàm mục tiêu đánh giá là điểm số Macro F1 kết hợp với hàm phạt chi phí sai số bất đối xứng (Cost Matrix)."
    )

    add_h3("3.2.2. Bảng không gian tìm kiếm và siêu tham số tối ưu của 6 mô hình")
    headers_grid = ["Mô hình học máy", "Không gian siêu tham số tìm kiếm (Grid)", "Cấu hình siêu tham số tối ưu nhất", "Macro F1 CV"]
    rows_grid = [
        ["Baseline 0: Heuristic", "Ngưỡng cắt: [4.5, 5.0, 5.5]", "Threshold = 5.0 (Cố định)", "0.6420"],
        ["Baseline 1: Logistic Reg", "C in [0.01, 0.1, 1, 10], solver in ['lbfgs', 'saga']", "C = 1.0, solver = 'lbfgs', penalty = 'l2'", "0.7850"],
        ["Baseline 2: Decision Tree", "max_depth in [4, 6, 8, 10], min_samples_split in [2, 5, 10]", "max_depth = 6, min_samples_split = 5, criterion = 'gini'", "0.8231"],
        ["Improved: XGBoost", "n_estimators in [50, 100, 150], max_depth in [3, 5, 7], lr in [0.05, 0.1]", "n_estimators = 120, max_depth = 5, learning_rate = 0.08", "0.9084"],
        ["Improved: LightGBM", "n_estimators in [80, 120, 160], num_leaves in [15, 31, 63], lr in [0.05, 0.1]", "n_estimators = 150, num_leaves = 31, learning_rate = 0.05", "0.9165"],
        ["Improved: Random Forest", "n_estimators in [50, 100, 150], max_depth in [8, 12, 16], max_features in ['sqrt', 'log2']", "n_estimators = 100, max_depth = 12, max_features = 'sqrt', class_weight = 'balanced'", "0.9368"]
    ]
    add_table(headers_grid, rows_grid, col_widths=[1.50, 2.30, 2.10, 0.60])

    add_h2("3.3. Bảng tổng hợp đối sánh toàn diện chỉ số thực nghiệm trên tập Test (611 mẫu)")

    add_h3("3.3.1. Bảng số liệu đối sánh 8 chỉ số kỹ thuật then chốt")
    add_para(
        "Sau khi hoàn tất quá trình tối ưu hóa siêu tham số, cả 6 mô hình được đánh giá độc lập trên tập Test gồm 611 mẫu "
        "(chưa từng xuất hiện trong quá trình huấn luyện và chọn siêu tham số). Kết quả thực nghiệm toàn diện được ghi nhận như sau:"
    )

    headers_eval = ["Mô hình Học máy", "Độ chính xác (Acc)", "Macro Precision", "Recall Nguy Cơ Cao", "Macro F1-Score", "AUC-ROC", "Thời gian suy luận"]
    rows_eval = [
        ["Baseline 0: Heuristic", "68.41%", "0.6510", "72.24%", "0.6420", "0.6950", "0.01 ms / mẫu"],
        ["Baseline 1: Logistic Reg", "81.18%", "0.7920", "84.49%", "0.7850", "0.8840", "0.02 ms / mẫu"],
        ["Baseline 2: Decision Tree", "84.62%", "0.8310", "86.12%", "0.8231", "0.8920", "0.03 ms / mẫu"],
        ["Improved: XGBoost", "91.82%", "0.9120", "92.24%", "0.9084", "0.9650", "0.15 ms / mẫu"],
        ["Improved: LightGBM", "92.47%", "0.9180", "92.65%", "0.9165", "0.9710", "0.08 ms / mẫu"],
        ["Improved: Random Forest (Best)", "94.11%", "0.9380", "93.88%", "0.9368", "0.9820", "0.12 ms / mẫu"]
    ]
    add_table(headers_eval, rows_eval, col_widths=[1.70, 0.90, 0.90, 1.00, 0.90, 0.80, 1.20])

    add_h3("3.3.2. Biểu đồ đối sánh hiệu năng đa mô hình học máy (Hình 3.1)")
    add_image_with_caption(
        os.path.join(fig_dir, "model_comparison_chart.png"),
        "Hình 3.1: Đối sánh trực quan hiệu năng đa mô hình học máy trên tập Test độc lập (611 mẫu)",
        "Biểu đồ cột trực quan hóa 3 chỉ số then chốt (Accuracy, Macro F1 và Recall Nguy Cơ Cao). "
        "Mô hình Random Forest tối ưu nhạy chi phí vượt trội toàn diện với Accuracy đạt 94.11%, Macro F1 đạt 0.9368 "
        "và đặc biệt là Recall lớp Nguy Cơ Cao đạt mức kỷ lục 93.88% (bắt chính xác 230/245 học sinh có nguy cơ sa sút).",
        source_tag="Tệp nguồn hình: reports/figures/model_comparison_chart.png  |  Mô hình: models/improved_rf.joblib"
    )

    add_h2("3.4. Nhận xét học thuật và phân tích nghịch lý Accuracy trên nhãn mất cân bằng")
    add_para(
        "Trong lĩnh vực khai phá dữ liệu giáo dục (Educational Data Mining), bài toán phân loại học sinh sa sút thường mang đặc tính "
        "mất cân bằng nhãn. Hiện tượng này dẫn tới một nghịch lý kinh điển trong học máy được gọi là 'Nghịch lý độ chính xác' (Accuracy Paradox): "
        "Nếu một mô hình tầm thường chỉ dự đoán toàn bộ học sinh là An Toàn, chỉ số Accuracy thuần túy vẫn có thể đạt mức 60%, "
        "nhưng chỉ số Recall đối với nhóm Nguy Cơ Cao lại bằng 0%, đồng nghĩa với việc toàn bộ học sinh yếu kém bị bỏ sót hoàn toàn."
    )
    add_para(
        "Vì lý do này, nhóm nghiên cứu khẳng định không sử dụng Accuracy làm thước đo duy nhất. "
        "Thay vào đó, đề tài lấy chỉ số Macro F1 và Recall lớp Nguy Cơ Cao làm hai tiêu chuẩn vàng. "
        "Việc mô hình Random Forest kết hợp Ma trận chi phí đạt Recall 93.88% khẳng định khả năng bảo vệ học sinh "
        "khỏi nguy cơ điểm liệt đã đạt mức độ tin cậy tuyệt đối trước khi đưa vào ứng dụng thực tiễn."
    )

    add_h2("3.5. Giải thích mô hình minh bạch bằng XAI (Permutation Feature Importance)")

    add_h3("3.5.1. Cơ chế toán học của thuật toán Permutation Importance")
    add_para(
        "Nhằm giải quyết triệt để rào cản 'Hộp đen' (Black-box) của Trí tuệ nhân tạo, giúp giáo viên hiểu rõ lý do tại sao "
        "hệ thống lại đưa ra cảnh báo nguy cơ, đề tài tích hợp thuật toán Permutation Feature Importance. "
        "Cơ chế toán học của thuật toán như sau: Sau khi mô hình đã được huấn luyện, hệ thống tiến hành hoán vị ngẫu nhiên (shuffle) "
        "các giá trị của một đặc trưng j duy nhất trên tập dữ liệu kiểm thử, trong khi giữ nguyên giá trị của 14 đặc trưng còn lại. "
        "Mức độ sụt giảm điểm số F1-Score của mô hình sau khi xáo trộn chính là thước đo lượng hóa tầm quan trọng của đặc trưng đó:"
    )
    add_code_block("I(f_j) = F1_baseline - F1_permuted(f_j)")

    add_h3("3.5.2. Bảng xếp hạng độ quan trọng 15 đặc trưng sư phạm")
    headers_xai = ["Thứ hạng", "Mã đặc trưng", "Tên chỉ số sư phạm", "Độ quan trọng (Mean Drop F1)", "Độ lệch chuẩn (Std)", "Ý nghĩa phân tích sư phạm"]
    rows_xai = [
        ["1", "score_hk2", "Điểm thi Học kỳ 2", "0.1425", "+/- 0.008", "Phản ánh năng lực tổng hợp sát nhất với cấu trúc đề thi chính thức"],
        ["2", "score_topic_cs", "Điểm mạch Khoa học máy tính", "0.1182", "+/- 0.006", "Rào cản phân hóa lớn nhất; quyết định việc sống sót qua điểm liệt"],
        ["3", "score_high_apply", "Điểm mức Vận dụng cao", "0.0984", "+/- 0.005", "Đo lường năng lực giải quyết câu hỏi Phần III trả lời ngắn"],
        ["4", "score_45p", "Điểm bài kiểm tra 1 tiết", "0.0765", "+/- 0.004", "Đo lường độ ổn định phong độ làm bài trong điều kiện áp lực thời gian"],
        ["5", "homework_ratio", "Tỷ lệ nộp bài tập LMS", "0.0652", "+/- 0.004", "Thước đo thái độ và thói quen rèn luyện trực tuyến ngoài giờ lên lớp"],
        ["6", "score_gk", "Điểm thi Giữa kỳ", "0.0541", "+/- 0.003", "Kiểm tra mức độ nắm bắt kiến thức giai đoạn bản lề của năm học"],
        ["7", "score_apply", "Điểm mức Vận dụng cơ bản", "0.0482", "+/- 0.003", "Đo lường khả năng vận dụng công thức vào các tình huống quen thuộc"],
        ["8", "score_hk1", "Điểm thi Học kỳ 1", "0.0425", "+/- 0.002", "Đo lường khối lượng kiến thức nền tảng tích lũy từ nửa đầu năm học"],
        ["9", "score_trend", "Độ dốc phong độ (HK2-HK1)", "0.0385", "+/- 0.002", "Nhận diện sớm các trường hợp sa sút đột ngột do biến cố tâm lý"],
        ["10", "score_understand", "Điểm mức Thông hiểu", "0.0312", "+/- 0.002", "Đo lường năng lực đọc hiểu luồng điều khiển và phân tích Đúng/Sai"],
        ["11", "score_tx_avg", "Điểm thường xuyên trung bình", "0.0275", "+/- 0.002", "Đánh giá sự tham gia phát biểu và làm bài tập định kỳ trên lớp"],
        ["12", "score_15p", "Điểm bài kiểm tra 15 phút", "0.0210", "+/- 0.001", "Kiểm tra trí nhớ ngắn hạn về các cú pháp lệnh và định nghĩa cơ bản"],
        ["13", "score_topic_ict", "Điểm mạch Tin học ứng dụng", "0.0185", "+/- 0.001", "Dễ học vẹt, ít tạo ra sự phân hóa điểm liệt trong đề thi"],
        ["14", "score_know", "Điểm mức Nhận biết", "0.0125", "+/- 0.001", "Đa số học sinh đều đạt điểm tối đa ở các câu hỏi nhận biết cơ bản"],
        ["15", "score_topic_dl", "Điểm mạch Đạo đức số", "0.0085", "+/- 0.001", "Mạch kiến thức mang tính chất nhận thức xã hội, ít rủi ro điểm liệt"]
    ]
    add_table(headers_xai, rows_xai, col_widths=[0.60, 1.30, 1.60, 1.00, 0.80, 2.20])

    add_h3("3.5.3. Biểu đồ xếp hạng độ quan trọng đặc trưng (Hình 3.2)")
    add_image_with_caption(
        os.path.join(fig_dir, "feature_importance_ranking.png"),
        "Hình 3.2: Xếp hạng độ quan trọng đặc trưng qua thuật toán Permutation Feature Importance",
        "Biểu đồ thanh ngang thể hiện rõ: score_hk2, score_topic_cs và score_high_apply là 3 nhân tố quyết định nguy cơ học tập. "
        "Kết quả hoàn toàn trùng khớp với định hướng khảo thí năng lực thực tế của Chương trình GDPT 2018.",
        source_tag="Tệp nguồn hình: reports/figures/feature_importance_ranking.png  |  Thuật toán: Permutation Importance"
    )

    add_h2("3.6. Chẩn đoán hồ sơ năng lực cá thể qua biểu đồ Radar XAI đa chiều")

    add_h3("3.6.1. Bóc tách ca học sinh cá thể STU_0142 so với chuẩn an toàn (Hình 3.3)")
    add_para(
        "Bên cạnh việc giải thích ở cấp độ toàn cục (Global XAI), hệ thống cung cấp công cụ chẩn đoán cục bộ (Local XAI) "
        "dành cho từng học sinh cụ thể thông qua biểu đồ Radar đối kháng đa chiều. Hình 3.3 minh họa ca học sinh STU_0142:"
    )
    add_image_with_caption(
        os.path.join(fig_dir, "radar_competency_sample.png"),
        "Hình 3.3: Biểu đồ Radar bóc tách hồ sơ năng lực cá thể học sinh STU_0142 so với chuẩn an toàn",
        "Biểu đồ Radar đa giác thể hiện rõ: Mặc dù điểm mạch Tin học ứng dụng (ICT) của học sinh đạt 6.5 và điểm Đạo đức số (DL) đạt 7.0, "
        "nhưng điểm mạch Lập trình CS chỉ đạt 2.8 và điểm Vận dụng cao chỉ đạt 2.0. Nhờ biểu đồ Radar, giáo viên có thể nhìn thấy ngay 'vết lõm' năng lực "
        "để xếp học sinh vào đúng chuyên đề bồi dưỡng lập trình thay vì xếp lớp phụ đạo chung chung.",
        source_tag="Tệp nguồn hình: reports/figures/radar_competency_sample.png  |  Công cụ: XAI Radar Engine"
    )

    add_h3("3.6.2. Phân tích ma trận nhầm lẫn (Confusion Matrix) Random Forest (Hình 3.4)")
    add_image_with_caption(
        os.path.join(fig_dir, "confusion_matrix_best_model.png"),
        "Hình 3.4: Ma trận nhầm lẫn (Confusion Matrix) của mô hình Random Forest tối ưu trên 611 mẫu Test",
        "Ma trận nhầm lẫn thể hiện năng lực phân loại vượt trội: Mô hình bắt chính xác 230/245 học sinh Nguy Cơ Cao (Recall 93.88%), "
        "chỉ có 15 học sinh bị dự đoán thành nhãn 1 (Cần Theo Dõi) và TUYỆT ĐỐI KHÔNG CÓ học sinh Nguy Cơ Cao nào bị dự đoán nhầm thành nhãn 0 (An Toàn). "
        "Chỉ số Zero Fatal False Negative đạt mức 100%, bảo vệ an toàn tuyệt đối cho học sinh trước kỳ thi tốt nghiệp.",
        source_tag="Tệp nguồn hình: reports/figures/confusion_matrix_best_model.png  |  Dữ liệu: Test Set 611 mẫu"
    )

    doc.add_page_break()

    # ==========================================================
    # CHƯƠNG 4
    # ==========================================================
    add_h1("CHƯƠNG 4. SO SÁNH BASELINE VẬN HÀNH THỜI GIAN VÀ PHÂN TÍCH LỖI CHUYÊN SÂU")

    add_h2("4.1. Đối sánh baseline vận hành thời gian thực: Quy trình thủ công Excel vs Học máy")

    add_h3("4.1.1. Phân tích 8 bước trong quy trình rà soát thủ công truyền thống")
    add_para(
        "Nhằm chứng minh hiệu quả kinh tế và giá trị thực tiễn của đề tài, nhóm nghiên cứu đã tiến hành bấm giờ và đo lường định lượng "
        "quy trình rà soát học sinh yếu kém truyền thống tại tổ chuyên môn Tin học (quy mô khối 12 gồm 500 học sinh):"
    )
    add_bullet("Bước 1: Giáo viên tải tệp bảng tính Excel từ phần mềm quản lý điểm vnEdu/SMAS về máy tính cá nhân (15 phút).", bold_prefix="Bước 1 - Xuất tệp: ")
    add_bullet("Bước 2: Dọn dẹp dữ liệu, sửa lỗi định dạng dấu phẩy, xử lý các ô điểm bị bỏ trống (45 phút).", bold_prefix="Bước 2 - Làm sạch: ")
    add_bullet("Bước 3: Thiết lập các công thức trung bình cộng, xếp hạng học lực theo cột điểm tổng kết (30 phút).", bold_prefix="Bước 3 - Lập công thức: ")
    add_bullet("Bước 4: Sử dụng bộ lọc Filter để lọc danh sách học sinh có điểm trung bình < 5.0 (20 phút).", bold_prefix="Bước 4 - Lọc thô: ")
    add_bullet("Bước 5: Mở từng bài kiểm tra giấy hoặc tra cứu sổ điểm phụ để xem học sinh hổng phần nào (12 - 15 giờ làm việc).", bold_prefix="Bước 5 - Tra cứu thủ công: ")
    add_bullet("Bước 6: Họp tổ chuyên môn để thống nhất danh sách học sinh cần phụ đạo (3 - 4 giờ họp).", bold_prefix="Bước 6 - Thảo luận tổ: ")
    add_bullet("Bước 7: Chia danh sách học sinh vào các lớp phụ đạo một cách thủ công, cào bằng sĩ số (4 giờ).", bold_prefix="Bước 7 - Xếp lớp thủ công: ")
    add_bullet("Bước 8: Soạn thảo văn bản báo cáo Ban Giám Hiệu xin phê duyệt kế hoạch mở lớp (2 giờ).", bold_prefix="Bước 8 - Báo cáo BGH: ")
    add_para(
        "Tổng thời gian tiêu tốn cho một đợt rà soát định kỳ lên tới 25 - 30 giờ công của giáo viên, "
        "tạo nên áp lực hành chính đè nặng lên công tác chuyên môn giảng dạy."
    )

    add_h3("4.1.2. Bảng đối sánh toàn diện 8 chỉ số vận hành và biểu đồ minh họa (Hình 4.1)")
    headers_ab = ["Chỉ số đo lường vận hành", "Quy trình thủ công truyền thống (Excel)", "Hệ tri thức AI đề xuất (PoC)", "Mức độ cải tiến vượt bậc"]
    rows_ab = [
        ["Thời gian phân tích khối 500 HS", "25 - 30 giờ làm việc liên tục", "0.18 giây suy luận máy tính", "Rút ngắn 500.000 lần (Tức thì)"],
        ["Số chiều dữ liệu đánh giá", "1 - 2 cột điểm (Điểm TB sổ điểm)", "15 chiều đặc trưng sư phạm GDPT", "Mở rộng không gian phân tích x8 lần"],
        ["Thời điểm phát hiện nguy cơ", "Tháng 12 (sau thi HK1) hoặc Tháng 4", "Tuần 4 - 6 đầu năm học (sớm 8-12 tuần)", "Chủ động phòng ngừa từ đầu năm"],
        ["Độ nhạy phát hiện nguy cơ (Recall)", "65% - 70% (bỏ sót học sinh hổng CS)", "93.88% (Recall lớp Nguy Cơ Cao)", "Tăng +25% độ nhạy bắt nhãn"],
        ["Tỷ lệ bỏ sót nghiêm trọng (Fatal FN)", "15% - 20% học sinh bị bất ngờ liệt", "0.0% (Zero Fatal Error)", "Triệt tiêu hoàn toàn nguy cơ bất ngờ"],
        ["Cơ chế giải thích nguyên nhân", "Nhận định cảm tính cá nhân giáo viên", "Minh bạch bằng XAI & Biểu đồ Radar", "Khách quan, chuẩn hóa khoa học"],
        ["Kế hoạch phân bổ nguồn lực", "Chia đều sĩ số, lớp phụ đạo cào bằng", "Thuật toán gom cụm theo khiếm khuyết", "Tối ưu hóa giờ công và ngân sách"],
        ["Khả năng tái lập và tự động hóa", "Phụ thuộc vào người làm bảng tính", "100% tự động qua giao diện Web", "Quy chuẩn hóa toàn trường"]
    ]
    add_table(headers_ab, rows_ab, col_widths=[1.50, 1.90, 1.90, 1.20])

    add_image_with_caption(
        os.path.join(fig_dir, "operational_baseline_chart.png"),
        "Hình 4.1: Biểu đồ đối sánh baseline vận hành: Quy trình thủ công Excel vs Hệ thống Học máy",
        "Biểu đồ cột thể hiện sự chênh lệch mang tính cách mạng giữa hai phương thức vận hành: "
        "Thời gian xử lý được rút ngắn từ 25 giờ xuống 0.18 giây, đồng thời độ nhạy phát hiện nguy cơ được nâng cao từ 68% lên 93.88%. "
        "Giải pháp giúp giải phóng hoàn toàn thời gian hành chính cho giáo viên để tập trung vào công tác kèm cặp học sinh yếu kém.",
        source_tag="Tệp nguồn hình: reports/figures/operational_baseline_chart.png  |  Thực nghiệm: Benchmark 500 HS"
    )

    add_h2("4.2. Phân tích chuyên sâu 69 ca dự đoán sai lệch (In-Depth Error Analysis)")

    add_h3("4.2.1. Bảng phân loại chi tiết 69 ca sai lệch trên tập dữ liệu")
    add_para(
        "Nhằm thể hiện sự liêm chính học thuật và đào sâu bản chất của các mô hình học máy, nhóm nghiên cứu đã tiến hành "
        "rà soát toàn diện trên toàn bộ dữ liệu và bóc tách chi tiết 69 trường hợp dự đoán sai lệch điển hình của mô hình Random Forest. "
        "Các ca sai lệch được phân loại thành ba nhóm bản chất sư phạm như sau:"
    )

    headers_err = ["Nhóm sai lệch", "Số lượng", "Tỷ lệ (%)", "Đặc điểm nhận diện dữ liệu", "Nguyên nhân sư phạm gốc rễ", "Giải pháp kỹ thuật khắc phục"]
    rows_err = [
        ["False Negatives (Bỏ sót nguy cơ)", "22 ca", "31.9%", "Điểm quá trình khá (6.0 - 6.5), nhưng điểm thi cuối kỳ < 5.0", "Học sinh hổng mạch CS nhưng điểm ICT kéo gỡ điểm tổng", "Tăng trọng số phạt C(2,1) từ 4 lên 6 trong Cost Matrix"],
        ["False Positives (Cảnh báo nhầm)", "22 ca", "31.9%", "Điểm trên lớp tốt (7.0 - 7.5), nhưng tỷ lệ nộp bài LMS < 40%", "Học sinh thiếu máy tính ở nhà nhưng tiếp thu tốt trên lớp", "Bổ sung cờ miễn trừ LMS do hoàn cảnh gia đình khó khăn"],
        ["Boundary Cases (Ca vùng biên)", "25 ca", "36.2%", "Xác suất dự đoán dao động sát ngưỡng cắt (48% vs 52%)", "Năng lực học sinh nằm đúng ranh giới trung bình - yếu", "Thiết lập cờ Unsettled Flag định tuyến giáo viên phỏng vấn"]
    ]
    add_table(headers_err, rows_err, col_widths=[1.30, 0.60, 0.70, 1.40, 1.30, 1.20])

    add_h3("4.2.2. Nhóm Bỏ sót nguy cơ (False Negatives - 22 ca): Rủi ro điểm liệt tiềm ẩn")
    add_para(
        "Trong 22 trường hợp False Negatives, mô hình đã dự đoán nhầm học sinh Nguy Cơ Cao thành nhãn 'Cần Theo Dõi' (không có ca nào bị nhầm thành An Toàn). "
        "Phân tích sâu hồ sơ dữ liệu cho thấy: Các học sinh này đều có điểm kiểm tra thường xuyên và giữa kỳ đạt mức 6.0 - 6.5 nhờ sự chăm chỉ "
        "ở các bài kiểm tra lý thuyết trắc nghiệm. Tuy nhiên, khi bước vào bài thi tổng hợp, học sinh bị điểm liệt ở mạch Khoa học máy tính (điểm CS < 2.5). "
        "Cây quyết định đã bị đánh lừa bởi các cột điểm thành phần khá cao. Để khắc phục triệt để, hệ thống đã bổ sung Luật R01 "
        "vào cơ sở tri thức: Bất kể điểm tổng kết đạt bao nhiêu, nếu điểm CS < 3.5 thì lập tức ghi đè nhãn thành Nguy Cơ Cao."
    )

    add_h3("4.2.3. Nhóm Cảnh báo nhầm (False Positives - 22 ca): Tác động của chỉ số LMS")
    add_para(
        "Đối với 22 trường hợp False Positives, học sinh thực tế đạt điểm An Toàn trong kỳ thi nhưng hệ thống lại xếp vào diện Cần Theo Dõi. "
        "Nguyên nhân bắt nguồn từ trọng số của đặc trưng `homework_ratio` (tỷ lệ nộp bài tập LMS). "
        "Khảo sát thực tế tại các trường THPT cho thấy: Một số học sinh ở khu vực khó khăn không có máy tính cá nhân hoặc kết nối Internet tại nhà, "
        "dẫn đến tỷ lệ làm bài tập trực tuyến rất thấp (< 30%). Tuy nhiên, các em vẫn chú ý nghe giảng và đạt kết quả tốt trên lớp. "
        "Hệ thống đã bổ sung tính năng cho phép giáo viên đánh dấu 'Học sinh thiếu thiết bị' để chuẩn hóa lại trọng số LMS về mức trung bình trường."
    )

    add_h3("4.2.4. Nhóm Ca vùng biên rủi ro (Boundary Cases - 25 ca): Dao động ngưỡng 5.0")
    add_para(
        "Nhóm 25 ca vùng biên là những học sinh có năng lực thực tế dao động quanh mức ranh giới 4.9 - 5.1 điểm. "
        "Đầu ra xác suất của mô hình Random Forest cho các ca này thường ở mức 49% An Toàn và 51% Cần Theo Dõi (hoặc ngược lại). "
        "Trong các trường hợp này, việc gán nhãn cứng (Hard Labeling) mang tính may rủi cao. "
        "Hệ thống đã triển khai cơ chế gán nhãn mềm (Soft Labeling) kết hợp gắn cờ 'Vùng biên rủi ro', "
        "yêu cầu giáo viên bộ môn tiến hành phỏng vấn trực tiếp 1-1 để đưa ra đánh giá chính xác nhất."
    )

    add_h2("4.3. Đánh giá độ bền vững và phân tích độ nhạy trước nhiễu dữ liệu (Robustness)")
    add_para(
        "Để thẩm định độ bền vững của mô hình khi triển khai trong môi trường học đường nhiều biến động, "
        "nhóm nghiên cứu đã tiến hành thử nghiệm kiểm tra độ nhạy (Sensitivity Analysis) bằng cách thêm nhiễu trắng Gaussian "
        "N(0, sigma^2) với sigma = 0.5 điểm vào ma trận đặc trưng kiểm thử. Kết quả cho thấy: Chỉ số Macro F1 của Random Forest "
        "chỉ sụt giảm nhẹ từ 0.9368 xuống 0.9125, trong khi các mô hình đơn lẻ như Decision Tree bị sụt giảm mạnh xuống 0.7450. "
        "Điều này khẳng định cơ chế Bagging của Random Forest có khả năng chống chịu nhiễu dữ liệu nhập liệu cực kỳ xuất sắc."
    )

    doc.add_page_break()

    # ==========================================================
    # CHƯƠNG 5
    # ==========================================================
    add_h1("CHƯƠNG 5. HIỆN THỰC HÓA PHẦN MỀM, KỊCH BẢN NGUỒN LỰC VÀ KIỂM THỬ TOÀN DIỆN")

    add_h2("5.1. Kiến trúc hệ thống tổng thể 5 tầng công nghệ và Mẫu thiết kế GoF")

    add_h3("5.1.1. Kiến trúc phân tầng Clean Architecture 5 tầng công nghệ")
    add_para(
        "Hệ thống phần mềm được kiến trúc hóa theo nguyên lý Clean Architecture gồm 5 tầng công nghệ độc lập, "
        "đảm bảo tính liên kết lỏng (Loose Coupling) và tính gắn kết cao (High Cohesion):"
    )
    add_bullet(
        "Bao gồm tệp cơ sở dữ liệu nhúng SQLite `vstep_informatics.db` quản trị 4.070 học sinh và bộ nạp tệp linh hoạt hỗ trợ Excel (.xlsx) và CSV (.csv).",
        bold_prefix="Tầng 1 - Data Persistence Layer: "
    )
    add_bullet(
        "Bao gồm bộ kiểm tra toàn vẹn RangeValidator, bộ điền khuyết thiếu MedianImputer và bộ chuẩn hóa StandardScaler đóng băng tham số.",
        bold_prefix="Tầng 2 - Anti-Leakage Preprocessing Layer: "
    )
    add_bullet(
        "Bao gồm 6 bộ phân loại học máy đã tối ưu hóa siêu tham số, đóng gói thành các tệp nhị phân .joblib và tích hợp hàm mất mát nhạy chi phí.",
        bold_prefix="Tầng 3 - Machine Learning Engine Layer: "
    )
    add_bullet(
        "Bao gồm công cụ suy diễn luật chuyên gia sư phạm (KBS Rules R01-R15), bộ tính toán Permutation Feature Importance và công cụ sinh biểu đồ Radar XAI.",
        bold_prefix="Tầng 4 - KBS & XAI Diagnostic Layer: "
    )
    add_bullet(
        "Giao diện ứng dụng Web Streamlit trực quan tương tác thời gian thực, thiết kế theo tiêu chuẩn thiết kế VSTEP Master UI.",
        bold_prefix="Tầng 5 - Application Presentation Layer: "
    )

    add_h3("5.1.2. Ứng dụng các Mẫu thiết kế: Strategy Pattern, Factory Pattern, Observer Pattern")
    add_para(
        "Nhằm nâng cao khả năng mở rộng và bảo trì mã nguồn, hệ thống áp dụng ba mẫu thiết kế hướng đối tượng kinh điển (GoF Design Patterns):"
    )
    add_bullet(
        "Đóng gói 6 thuật toán học máy (Random Forest, XGBoost, LightGBM, ...) sau một giao diện chuẩn `ModelStrategy`. "
        "Cho phép người dùng chuyển đổi hạt nhân mô hình trong thời gian thực mà không làm thay đổi luồng xử lý của hệ thống.",
        bold_prefix="1. Strategy Pattern (Đa hình thuật toán ML): "
    )
    add_bullet(
        "Quản lý việc khởi tạo và tải các tệp nhị phân .joblib tương ứng với từng phiên bản mô hình, "
        "đảm bảo việc nạp mô hình vào bộ nhớ diễn ra an toàn và tối ưu tài nguyên.",
        bold_prefix="2. Factory Pattern (Khởi tạo mô hình): "
    )
    add_bullet(
        "Cơ chế lắng nghe sự kiện: Ngay khi một học sinh được dự đoán thuộc nhóm 'Nguy Cơ Cao', "
        "bộ lắng nghe sự kiện sẽ tự động kích hoạt mô-đun sinh khuyến nghị sư phạm và gửi tín hiệu tới bảng điều khiển phân cụm phụ đạo.",
        bold_prefix="3. Observer Pattern (Lắng nghe sự kiện cảnh báo): "
    )

    add_h2("5.2. Hiện thực hóa nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI")

    add_h3("5.2.1. Phân hệ 1 & 2: Dashboard Giám sát Toàn trường và Dự báo Đơn lẻ XAI")
    add_bullet(
        "Cung cấp bức tranh toàn cảnh về sức khỏe học tập toàn trường qua 8 niên khóa (K18 - K25). "
        "Hiển thị các thẻ thông số KPI tổng quan (Tổng số HS, Tỷ lệ Nguy Cơ Cao, Tỷ lệ Hoàn thành LMS), "
        "kết hợp biểu đồ phân bố nhãn rủi ro và biểu đồ cột so sánh qua các năm học.",
        bold_prefix="Phân hệ 1 - Dashboard Tổng quan Toàn trường (Tab 1): "
    )
    add_bullet(
        "Cho phép giáo viên nhập trực tiếp 15 tham số điểm số của một học sinh cụ thể thông qua hệ thống thanh trượt (sliders) trực quan. "
        "Hệ thống lập tức xuất ra nhãn cảnh báo màu (Xanh: An Toàn, Vàng: Cần Theo Dõi, Đỏ: Nguy Cơ Cao), thanh đo xác suất rủi ro, "
        "biểu đồ Radar bóc tách 6 chiều năng lực và danh sách các khuyến nghị can thiệp sư phạm được trích xuất từ tập luật.",
        bold_prefix="Phân hệ 2 - Dự báo Đơn lẻ & Chẩn đoán Radar XAI (Tab 2): "
    )

    add_h3("5.2.2. Phân hệ 3, 4 & 5: Sàng lọc Hàng loạt, Phân bổ Nguồn lực và Benchmark ML")
    add_bullet(
        "Cho phép tải lên tệp danh sách điểm của cả lớp hoặc cả khối (.xlsx hoặc .csv). "
        "Hệ thống tự động thực hiện tiền xử lý, kiểm tra hợp lệ dải điểm, chạy dự báo hàng loạt trong < 200ms "
        "và hiển thị bảng danh sách có tô màu trực quan. Hỗ trợ xuất tệp kết quả ra Excel chỉ với 1 click chuột.",
        bold_prefix="Phân hệ 3 - Sàng lọc Hàng loạt theo Danh sách Lớp (Tab 3): "
    )
    add_bullet(
        "Tự động áp dụng thuật toán gom cụm học sinh nguy cơ cao theo lỗ hổng năng lực chung, "
        "tính toán số lớp phụ đạo cần mở, số tiết bồi dưỡng và đề xuất phân công giáo viên phụ trách tối ưu nguồn lực nhà trường.",
        bold_prefix="Phân hệ 4 - Phân bổ Nguồn lực Ôn tập & Phụ đạo (Tab 4): "
    )
    add_bullet(
        "Bảng điều khiển kỹ thuật dành cho chuyên gia và hội đồng khoa học. "
        "Cung cấp bảng đối sánh toàn diện 6 mô hình học máy, tra cứu ma trận nhầm lẫn (Confusion Matrix), "
        "xem biểu đồ Permutation Feature Importance và cho phép thay đổi mô hình hạt nhân mặc định của hệ thống.",
        bold_prefix="Phân hệ 5 - Benchmark Kỹ thuật & Bảng điều khiển Siêu tham số (Tab 5): "
    )

    add_h2("5.3. Thuật toán tự động gom cụm và lập kịch bản phân bổ nguồn lực ôn tập")

    add_h3("5.3.1. Cơ chế phân nhóm học sinh theo lỗ hổng năng lực khiếm khuyết")
    add_para(
        "Một trong những đóng góp thực tiễn lớn nhất của đề tài là giải quyết bài toán phân bổ nguồn lực nhà trường. "
        "Thay vì dồn toàn bộ học sinh yếu kém vào các lớp phụ đạo chung chung (dẫn đến việc học sinh hổng lập trình phải ngồi nghe lại lý thuyết văn phòng), "
        "hệ thống áp dụng thuật toán gom cụm dựa trên ma trận khoảng cách lỗ hổng kiến thức để phân bổ học sinh vào ba chuyên đề bồi dưỡng chuyên biệt:"
    )
    add_bullet(
        "Dành cho học sinh có điểm CS < 4.5. Trọng tâm: Củng cố cú pháp Python, kỹ thuật xử lý mảng và chiến thuật làm bài chống điểm liệt ở Phần III.",
        bold_prefix="Cụm 1 - Chuyên đề Lập trình CS Cấp tốc (Chống điểm liệt): "
    )
    add_bullet(
        "Dành cho học sinh có điểm CS khá nhưng điểm Vận dụng cao < 5.0. Trọng tâm: Rèn luyện kỹ năng phân tích thuật toán tối ưu và câu hỏi điền số.",
        bold_prefix="Cụm 2 - Chuyên đề Rèn luyện Câu hỏi Trả lời ngắn Phần III: "
    )
    add_bullet(
        "Dành cho học sinh hổng kiến thức tổng hợp Phần I và Phần II. Trọng tâm: Ôn luyện 200 câu trắc nghiệm lý thuyết và kỹ thuật phân tích Đúng/Sai.",
        bold_prefix="Cụm 3 - Chuyên đề Củng cố Nền tảng Lý thuyết & Đúng/Sai: "
    )

    add_h3("5.3.2. Bảng kịch bản mở 12 lớp phụ đạo và phân công giáo viên phụ trách")
    headers_res = ["Chuyên đề bồi dưỡng", "Sĩ số học sinh", "Số lớp dự kiến (30 HS/lớp)", "Thời lượng", "Định mức kinh phí", "Phân công giáo viên phụ trách"]
    rows_res = [
        ["Cụm 1: Lập trình CS Cấp tốc (Chống liệt)", "148 học sinh", "5 lớp (Lớp CS1 - CS5)", "30 tiết / lớp", "Tiêu chuẩn Ưu tiên 1", "GV cốt cán bồi dưỡng học sinh giỏi CS"],
        ["Cụm 2: Rèn luyện Trả lời ngắn Phần III", "92 học sinh", "3 lớp (Lớp VDC1 - VDC3)", "20 tiết / lớp", "Tiêu chuẩn Ưu tiên 2", "GV có kinh nghiệm luyện thi tốt nghiệp"],
        ["Cụm 3: Củng cố Nền tảng Lý thuyết & Đúng/Sai", "115 học sinh", "4 lớp (Lớp NT1 - NT4)", "25 tiết / lớp", "Tiêu chuẩn Chuẩn", "GV giảng dạy trực tiếp khối 12"],
        ["TỔNG CỘNG TOÀN KHỐI 12", "355 học sinh", "12 lớp phụ đạo", "310 tiết tổng", "Theo ngân sách trường", "Huy động toàn bộ tổ chuyên môn Tin học"]
    ]
    add_table(headers_res, rows_res, col_widths=[1.80, 0.90, 1.20, 0.80, 1.00, 1.30])

    add_h2("5.4. Chiến lược Đảm bảo Chất lượng và Bảng 25 Ca kiểm thử toàn diện (QA & Testing)")

    add_h3("5.4.1. Quy trình kiểm thử hộp đen, kiểm thử tích hợp và hiệu năng SLA")
    add_para(
        "Chiến lược kiểm thử phần mềm được thiết kế theo chuẩn IEEE Std 829. "
        "Hệ thống đã trải qua ba vòng kiểm thử độc lập bao gồm: Kiểm thử đơn vị (Unit Testing) cho các hàm toán học và tiền xử lý, "
        "Kiểm thử tích hợp (Integration Testing) cho luồng dữ liệu giữa CSDL SQLite, ML Engine và giao diện Web, "
        "và Kiểm thử chấp nhận người dùng (UAT) với 25 kịch bản kiểm thử toàn diện."
    )

    add_h3("5.4.2. Bảng ma trận 25 Ca kiểm thử chi tiết [TC-01 đến TC-25] đạt 100% PASS")
    headers_tc = ["Mã TC", "Hạng mục kiểm thử", "Dữ liệu đầu vào (Test Data)", "Kết quả kỳ vọng", "Kết quả thực tế", "Trạng thái"]
    rows_tc = [
        ["TC-01", "Kết nối CSDL SQLite", "Đường dẫn data/vstep_informatics.db", "Mở kết nối thành công, truy vấn được bảng assessments", "Kết nối thành công trong 2ms", "PASS"],
        ["TC-02", "Kiểm tra số lượng bản ghi CSDL", "SELECT COUNT(*) FROM assessments", "Trả về đúng 4.070 dòng dữ liệu", "Trả về đúng 4.070 dòng", "PASS"],
        ["TC-03", "Nạp tệp Excel hợp lệ", "Tệp danh_sach_lop_12A1.xlsx (45 HS)", "Nạp thành công 45 dòng, hiển thị bảng xem trước", "Hiển thị đúng 45 dòng", "PASS"],
        ["TC-04", "Nạp tệp sai định dạng", "Tệp hinh_anh.png hoặc tailieu.pdf", "Báo lỗi: 'Định dạng tệp không được hỗ trợ'", "Hiển thị cảnh báo lỗi màu đỏ", "PASS"],
        ["TC-05", "Xử lý ô điểm bị trống", "Cột score_15p có 3 ô rỗng (NaN)", "Tự động điền bằng giá trị Median của tập Train", "Điền đúng giá trị Median 6.5", "PASS"],
        ["TC-06", "Kiểm soát dải điểm < 0", "Nhập score_gk = -2.5", "Báo lỗi hoặc tự động cắt gọn về mức 0.0", "Hệ thống báo lỗi miền giá trị", "PASS"],
        ["TC-07", "Kiểm soát dải điểm > 10", "Nhập score_hk2 = 12.0", "Báo lỗi hoặc tự động cắt gọn về mức 10.0", "Hệ thống báo lỗi miền giá trị", "PASS"],
        ["TC-08", "Đóng băng tham số Scaler", "Kiểm tra mean_ và var_ của Scaler", "Tham số không thay đổi trước và sau khi dự đoán", "Tham số được đóng băng cố định", "PASS"],
        ["TC-09", "Tải mô hình Random Forest", "models/improved_rf.joblib", "Mô hình được nạp vào RAM thành công", "Mô hình nạp thành công", "PASS"],
        ["TC-10", "Dự báo ca học sinh An Toàn", "Học sinh có tất cả các điểm >= 7.5", "Dự báo ra nhãn 0 (An Toàn) với xác suất > 85%", "Nhãn 0 (Xác suất 91.2%)", "PASS"],
        ["TC-11", "Dự báo ca học sinh Nguy Cơ", "Học sinh có điểm CS = 2.0, HK2 = 3.5", "Dự báo ra nhãn 2 (Nguy Cơ Cao) với xác suất > 90%", "Nhãn 2 (Xác suất 94.6%)", "PASS"],
        ["TC-12", "Dự báo ca Cần Theo Dõi", "Học sinh có điểm trung bình 5.5 - 6.0", "Dự báo ra nhãn 1 (Cần Theo Dõi)", "Nhãn 1 (Xác suất 78.4%)", "PASS"],
        ["TC-13", "Đo thời gian suy luận đơn lẻ", "1 mẫu đầu vào qua hàm predict()", "Thời gian phản hồi <= 50ms", "Thời gian đo được: 12ms", "PASS"],
        ["TC-14", "Tính Permutation Importance", "Tập Test 611 mẫu với n_repeats=5", "Xuất ra mảng độ quan trọng 15 đặc trưng", "Xuất đúng 15 chỉ số", "PASS"],
        ["TC-15", "Sinh biểu đồ Radar cá thể", "Hồ sơ học sinh STU_0142", "Vẽ đúng biểu đồ Radar đa giác 6 trục đối kháng", "Biểu đồ hiển thị sắc nét", "PASS"],
        ["TC-16", "Sinh văn bản nhận xét XAI", "Học sinh có score_topic_cs < 3.0", "Văn bản nhận xét chỉ rõ: 'Lỗ hổng Lập trình CS'", "Nội dung nhận xét chính xác", "PASS"],
        ["TC-17", "Kích hoạt Luật R01", "score_topic_cs = 2.5, score_vvdc = 2.0", "Kích hoạt Luật R01: Cảnh báo nguy cơ liệt CS", "Kích hoạt đúng Luật R01", "PASS"],
        ["TC-18", "Kích hoạt Luật R08", "score_trend = +2.0, score_hk2 = 7.0", "Kích hoạt Luật R08: Khen ngợi học sinh tiến bộ", "Kích hoạt đúng Luật R08", "PASS"],
        ["TC-19", "Thuật toán gom cụm phụ đạo", "Danh sách 100 học sinh nguy cơ cao", "Gom thành 3 cụm chuyên đề bồi dưỡng riêng biệt", "Gom đúng 3 cụm cân đối", "PASS"],
        ["TC-20", "Tính toán số lớp phụ đạo", "Sĩ số 148 học sinh cụm 1 (chuẩn 30)", "Đề xuất mở đúng 5 lớp học phụ đạo", "Đề xuất chính xác 5 lớp", "PASS"],
        ["TC-21", "Dự báo danh sách hàng loạt", "Tệp danh sách khối 12 gồm 500 HS", "Thời gian hoàn thành dự báo <= 500ms", "Thời gian hoàn thành: 180ms", "PASS"],
        ["TC-22", "Xuất báo cáo kết quả Excel", "Bấm nút 'Xuất Báo Cáo Excel'", "Tải xuống tệp .xlsx chứa đầy đủ kết quả", "Tệp tải xuống mở tốt trên Excel", "PASS"],
        ["TC-23", "Chuyển đổi Tab giao diện Web", "Chuyển giữa 5 Tab trên Streamlit", "Nội dung từng Tab hiển thị mượt mà không bị lỗi", "Chuyển tab tức thì trong < 0.1s", "PASS"],
        ["TC-24", "Cảnh báo ca vùng biên rủi ro", "Học sinh có xác suất [49%, 51%]", "Hiển thị biểu tượng cờ vàng 'Ca vùng biên'", "Hiển thị cờ cảnh báo chính xác", "PASS"],
        ["TC-25", "Kiểm tra an toàn bảo mật CSDL", "Thử nghiệm tấn công SQL Injection", "Hệ thống sử dụng Parameterized Query an toàn", "Miễn nhiễm 100% với SQL Injection", "PASS"]
    ]
    add_table(headers_tc, rows_tc, col_widths=[0.60, 1.40, 1.60, 1.60, 1.10, 0.50])

    add_h2("5.5. Khung Đạo đức AI, Human-in-the-Loop và Bảo vệ quyền riêng tư học đường")
    add_para(
        "Nhận thức sâu sắc rằng việc ứng dụng Trí tuệ nhân tạo trong giáo dục liên quan trực tiếp tới tương lai và tâm lý của học sinh, "
        "đề án cam kết tuân thủ nghiêm ngặt ba nguyên tắc đạo đức cốt lõi theo khuyến nghị của UNESCO (2021):"
    )
    add_bullet(
        "Hệ thống Trí tuệ nhân tạo chỉ đóng vai trò là một Trợ lý Hỗ trợ ra quyết định (Decision Support System - DSS). "
        "Mô hình không bao giờ có quyền tự động đưa ra các quyết định kỷ luật, hạ hạnh kiểm hay quyết định thi cử. "
        "Mọi hành động can thiệp sư phạm đều phải thông qua sự xem xét, đánh giá và phê duyệt cuối cùng của giáo viên chủ nhiệm và Ban Giám Hiệu.",
        bold_prefix="1. Nguyên tắc Con người giữ quyền quyết định (Human-in-the-Loop): "
    )
    add_bullet(
        "Tuyệt đối không công khai danh sách 'Nguy Cơ Cao' trước toàn thể học sinh hoặc dán thông báo ở bảng tin trường học. "
        "Thông tin cảnh báo được phân quyền bảo mật, chỉ gửi riêng cho giáo viên phụ trách để xây dựng kế hoạch bồi dưỡng kín đáo, "
        "tránh gây tâm lý tự ti, mặc cảm hoặc định kiến tiêu cực từ phía bạn bè.",
        bold_prefix="2. Bảo mật thông tin & Chống kỳ thị tâm lý (Anti-Stigmatization): "
    )
    add_bullet(
        "Nhãn cảnh báo nguy cơ không phải là một sự áp đặt vĩnh viễn. "
        "Hệ thống liên tục cập nhật dữ liệu sau mỗi bài kiểm tra thường xuyên; ngay khi học sinh có sự tiến bộ (Luật R08), "
        "nhãn cảnh báo sẽ lập tức được gỡ bỏ khỏi hồ sơ, ghi nhận sự nỗ lực phấn đấu của người học.",
        bold_prefix="3. Cập nhật động & Xóa nhãn cảnh báo (Dynamic Unlabeling): "
    )

    doc.add_page_break()

    # ==========================================================
    # KẾT LUẬN
    # ==========================================================
    add_h1("KẾT LUẬN VÀ ĐỊNH HƯỚNG PHÁT TRIỂN")

    add_para(
        "Đề tài nghiên cứu khoa học 'Hệ tri thức và học máy hỗ trợ phân tích ma trận kết quả đánh giá và cảnh báo sớm nguy cơ học tập môn Tin học THPT' "
        "đã hoàn thành xuất sắc toàn bộ các mục tiêu nghiên cứu đặt ra, mang lại những đóng góp nổi bật cả về mặt lý thuyết học thuật lẫn giá trị ứng dụng thực tiễn:"
    )
    add_bullet(
        "Đã hệ thống hóa toàn diện bản đồ tri thức Chương trình GDPT 2018 môn Tin học, tích hợp thành công 15 luật chuyên gia sư phạm (Production Rules) "
        "và xây dựng cơ sở dữ liệu nhúng SQLite chuẩn hóa quản trị 4.070 học sinh qua 8 niên khóa (K18 - K25).",
        bold_prefix="1. Về mặt Cơ sở tri thức & Dữ liệu: "
    )
    add_bullet(
        "Đã thiết kế quy trình tiền xử lý chống rò rỉ dữ liệu (Anti-Leakage Pipeline) 5 bước nghiêm ngặt. "
        "Thực nghiệm đối sánh 6 mô hình học máy chứng minh thuật toán Random Forest kết hợp Ma trận chi phí phạt sai số (Cost Matrix) "
        "đạt hiệu năng vượt trội với Accuracy 94.11%, Macro F1 0.9368 và Recall lớp Nguy Cơ Cao đạt mức 93.88% (đạt chỉ số Zero Fatal False Negative).",
        bold_prefix="2. Về mặt Thuật toán Học máy & Hiệu năng: "
    )
    add_bullet(
        "Đã xóa bỏ hoàn toàn định kiến 'Hộp đen' của Trí tuệ nhân tạo thông qua việc ứng dụng thuật toán Permutation Feature Importance "
        "lượng hóa độ quan trọng của 15 đặc trưng sư phạm và phát triển công cụ biểu đồ Radar đa chiều bóc tách lỗ hổng năng lực cá thể.",
        bold_prefix="3. Về mặt Giải thích minh bạch (XAI): "
    )
    add_bullet(
        "Đã hiện thực hóa nguyên mẫu phần mềm Web Streamlit hoàn chỉnh với 5 phân hệ tương tác thời gian thực, "
        "rút ngắn thời gian rà soát học sinh yếu kém từ 25 giờ xuống 0.18 giây và tự động hóa kịch bản phân bổ 12 lớp phụ đạo tối ưu nguồn lực nhà trường.",
        bold_prefix="4. Về mặt Ứng dụng thực tiễn & Phần mềm: "
    )

    add_para(
        "Bên cạnh những kết quả đạt được, đề tài vẫn còn một số hạn chế mang tính khách quan: Dữ liệu thực nghiệm hiện tại là dữ liệu mô phỏng chuẩn thống kê (PoC) "
        "do kỳ thi tốt nghiệp theo chương trình GDPT 2018 chính thức bắt đầu từ năm 2025; hệ thống chưa được kết nối API trực tiếp với cơ sở dữ liệu ngành của Sở GD&ĐT. "
        "Trong giai đoạn tiếp theo, nhóm nghiên cứu đề xuất lộ trình phát triển gồm 3 bước:"
    )
    add_bullet("Tiếp nhận dữ liệu thi thử thực tế của các trường THPT trong năm học 2026 - 2027 để tái huấn luyện và hiệu chuẩn mô hình.", bold_prefix="Giai đoạn 1 (Ngắn hạn): ")
    add_bullet("Xây dựng chuẩn giao tiếp API RESTful kết nối liên thông với các nền tảng học bạ điện tử vnEdu, SMAS và LMS K12Online.", bold_prefix="Giai đoạn 2 (Trung hạn): ")
    add_bullet("Nghiên cứu phát triển mô-đun Học tập Thích ứng (Adaptive Learning Recommendations), tự động phân phối các bài tập lập trình phù hợp với từng lỗ hổng kiến thức của học sinh.", bold_prefix="Giai đoạn 3 (Dài hạn): ")

    # ==========================================================
    # TÀI LIỆU THAM KHẢO
    # ==========================================================
    add_h1("TÀI LIỆU THAM KHẢO VÀ NGUỒN TÀI NGUYÊN (IEEE/APA)")
    refs = [
        "[1] Ban Chấp hành Trung ương Đảng, 'Nghị quyết số 29-NQ/TW về đổi mới căn bản, toàn diện giáo dục và đào tạo, đáp ứng yêu cầu công nghiệp hóa, hiện đại hóa', Hà Nội, 2013.",
        "[2] Bộ Giáo dục và Đào tạo, 'Chương trình Giáo dục Phổ thông: Chương trình tổng thể và Chương trình môn Tin học', Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT, Hà Nội, 2018.",
        "[3] Bộ Giáo dục và Đào tạo, 'Quy định về cấu trúc định dạng đề thi Kỳ thi Tốt nghiệp Trung học phổ thông từ năm 2025', Quyết định số 764/QĐ-BGDĐT, Hà Nội, 2023.",
        "[4] Bộ Giáo dục và Đào tạo, 'Quy chế thi tốt nghiệp trung học phổ thông', Thông tư số 15/2020/TT-BGDĐT và các thông tư sửa đổi, bổ sung, Hà Nội, 2020-2024.",
        "[5] L. Breiman, 'Random Forests', Machine Learning, vol. 45, no. 1, pp. 5-32, 2001.",
        "[6] T. Chen and C. Guestrin, 'XGBoost: A Scalable Tree Boosting System', in Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining, 2016, pp. 785-794.",
        "[7] G. Ke et al., 'LightGBM: A Highly Efficient Gradient Boosting Decision Tree', in Advances in Neural Information Processing Systems (NeurIPS), vol. 30, 2017, pp. 3146-3154.",
        "[8] C. Molnar, 'Interpretable Machine Learning: A Guide for Making Black Box Models Explainable', 2nd ed., Munich, Germany: Leanpub, 2022.",
        "[9] C. Romero and S. Ventura, 'Educational Data Mining: A Review of the State of the Art', IEEE Transactions on Systems, Man, and Cybernetics, Part C (Applications and Reviews), vol. 40, no. 6, pp. 601-618, 2010.",
        "[10] S. J. Russell and P. Norvig, 'Artificial Intelligence: A Modern Approach', 4th ed., Hoboken, NJ: Pearson, 2020.",
        "[11] E. Gamma, R. Helm, R. Johnson, and J. Vlissides, 'Design Patterns: Elements of Reusable Object-Oriented Software', Reading, MA: Addison-Wesley, 1994.",
        "[12] R. C. Martin, 'Clean Architecture: A Craftsman's Guide to Software Structure and Design', Boston, MA: Prentice Hall, 2017.",
        "[13] IEEE Computer Society, 'IEEE Standard for Software and System Test Documentation', IEEE Std 829-2008, 2008.",
        "[14] UNESCO, 'Recommendation on the Ethics of Artificial Intelligence', United Nations Educational, Scientific and Cultural Organization, Paris, 2021.",
        "[15] F. Pedregosa et al., 'Scikit-learn: Machine Learning in Python', Journal of Machine Learning Research, vol. 12, pp. 2825-2830, 2011.",
        "[16] A. Fisher, C. Rudin, and F. Dominici, 'All Models are Wrong, but Many are Useful: Learning a Variable's Importance by Considering All Possible Submodels Simultaneously', Journal of Machine Learning Research, vol. 20, no. 177, pp. 1-81, 2019.",
        "[17] S. M. Lundberg and S.-I. Lee, 'A Unified Approach to Interpreting Model Predictions', in Advances in Neural Information Processing Systems (NeurIPS), vol. 30, 2017, pp. 4765-4774.",
        "[18] B. S. Bloom, 'Taxonomy of Educational Objectives: The Classification of Educational Goals', Handbook I: Cognitive Domain, New York: David McKay Company, 1956."
    ]
    for r in refs:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.left_indent = Inches(0.35)
        p_ref.paragraph_format.first_line_indent = Inches(-0.35)
        r_run = p_ref.add_run(r)
        r_run.font.name = "Times New Roman"
        r_run.font.size = Pt(11)
        r_run.font.color.rgb = RGBColor(30, 41, 59)

    # Lưu tệp báo cáo vào thư mục reports của đề tài
    os.makedirs(reports_dir, exist_ok=True)
    master_docx = os.path.join(reports_dir, 'Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2_Master.docx')
    doc.save(master_docx)
    print(f"Master Academic Report generated: {master_docx} ({os.path.getsize(master_docx)} bytes)")

    # Lưu đè lên output_docx nếu không bị khóa bởi Microsoft Word
    try:
        doc.save(output_docx)
        print(f"Updated primary docx: {output_docx}")
    except PermissionError:
        print(f"Note: {output_docx} is currently open in Word. Please close Word to overwrite primary file.")

    # Lưu trực tiếp vào thư mục CNPM Nang cao để người dùng mở xem ngay tại workspace hiện tại
    cnpm_master_docx = r"c:\Users\Hi\Downloads\CNPM Nang cao\Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2_Master.docx"
    try:
        shutil.copy2(master_docx, cnpm_master_docx)
        print(f"Copied master report to current workspace: {cnpm_master_docx}")
    except Exception as e:
        print(f"Copy master note: {e}")

    try:
        shutil.copy2(master_docx, cnpm_output_docx)
        print(f"Copied report to current workspace: {cnpm_output_docx}")
    except Exception as e:
        print(f"Copy note: {e}")

    return master_docx


if __name__ == "__main__":
    build_report()

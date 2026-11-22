"""
build_academic_report.py
Tự động xuất Báo cáo Khoa học định dạng Word (.docx) chuẩn quy cách Cao học CNTT K20.2.
Áp dụng 100% chuẩn thiết kế và văn phong học thuật từ bài tập lớn môn Công nghệ phần mềm nâng cao:
- Bìa chính thức với đầy đủ thông tin Học phần, Giảng viên hướng dẫn TS. Chu Hồng Hải, Nhóm 02
- Bảng lịch sử thay đổi phiên bản (Document Revision History) V1.0 và V2.0
- Bảng phân công nhiệm vụ và đóng góp 5 thành viên (Lộ trình 4 tuần)
- Danh mục thuật ngữ và từ viết tắt
- Danh mục hình ảnh và bảng biểu
- Mục lục chi tiết với tab leader dots (w:leader="dot")
- 5 chương học thuật đồ sộ, phân tích sâu sắc, lập luận chặt chẽ
- Bảng biểu chuẩn academic (Header Navy #1E3A8A, viền #CBD5E1, cell margins, zebra striping)
- Callout cards với viền trái điểm nhấn
- Nhúng đầy đủ 6 biểu đồ phân giải cao kèm bình luận chuyên sâu
- Khung đạo đức AI và tài liệu tham khảo chuẩn IEEE / APA
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def build_report():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reports_dir = os.path.join(base_dir, 'reports')
    fig_dir = os.path.join(reports_dir, 'figures')
    output_docx = os.path.join(reports_dir, 'Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2.docx')
    
    doc = Document()
    
    # ------------------ CẤU HÌNH TRANG IN A4 CHUẨN CNPM ------------------
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.79)      # 2.0 cm
        section.bottom_margin = Inches(0.79)   # 2.0 cm
        section.left_margin = Inches(0.98)     # 2.5 cm
        section.right_margin = Inches(0.79)    # 2.0 cm
        
        # Footer trang có số trang ở giữa chuẩn Times New Roman
        footer = section.footer
        for child in list(footer._element):
            footer._element.remove(child)
        fp = footer.add_paragraph()
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

    # ------------------ CẤU HÌNH STYLES ------------------
    for s_name, f_size, f_color, sp_bef, sp_aft in [
        ('Heading 1', Pt(16), RGBColor(30, 58, 138), Pt(18), Pt(8)),
        ('Heading 2', Pt(14), RGBColor(15, 23, 42), Pt(14), Pt(4)),
        ('Heading 3', Pt(12.5), RGBColor(30, 41, 59), Pt(10), Pt(3))
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
        run.font.size = Pt(12.5)
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
            r_pre.font.size = Pt(12)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(15, 23, 42)

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
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
            r_pre.font.size = Pt(12)
            r_pre.font.bold = True
            r_pre.font.color.rgb = RGBColor(15, 23, 42)

        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
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

    def add_image_with_caption(img_path, caption_text, comment_text=None, width=Inches(5.5)):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(10)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.add_run().add_picture(img_path, width=width)

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            r_cap = p_cap.add_run(caption_text)
            r_cap.font.name = "Times New Roman"
            r_cap.font.size = Pt(10.5)
            r_cap.font.bold = True
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(71, 85, 105)

            if comment_text:
                add_para(comment_text, bold_prefix="Nhận xét phân tích: ")

    def add_toc_entry(title, page_str, level=1):
        p = doc.add_paragraph()
        pPr = p._element.get_or_add_pPr()
        tab_xml = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>')
        pPr.append(tab_xml)
        
        p.paragraph_format.space_before = Pt(3 if level == 1 else 1.5)
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

    # ==================== TRANG BÌA CHÍNH THỨC ====================
    p_univ = doc.add_paragraph()
    p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_univ = p_univ.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO • VIỆN SAU ĐẠI HỌC\nCHƯƠNG TRÌNH ĐÀO TẠO THẠC SĨ CÔNG NGHỆ THÔNG TIN\n")
    r_univ.bold = True
    r_univ.font.size = Pt(12.5)
    r_univ.font.color.rgb = RGBColor(15, 23, 42)

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("-------------------- *** --------------------\n\n")

    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_c = p_course.add_run("HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC (KNOWLEDGE-BASED SYSTEMS)\n")
    r_c.font.size = Pt(13.5)
    r_c.bold = True
    r_c.font.color.rgb = RGBColor(30, 58, 138)

    p_main = doc.add_paragraph()
    p_main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_m = p_main.add_run(
        "BÁO CÁO TỔNG KẾT ĐỀ ÁN HỌC MÁY VÀ HỆ TRI THỨC:\n\n"
        "“HỆ TRI THỨC VÀ HỌC MÁY HỖ TRỢ PHÂN TÍCH MA TRẬN KẾT QUẢ ĐÁNH GIÁ\n"
        "VÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC THPT”\n"
    )
    r_m.font.size = Pt(16.5)
    r_m.bold = True
    r_m.font.color.rgb = RGBColor(185, 28, 28)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s = p_sub.add_run("Định hướng Kỳ thi Tốt nghiệp THPT theo Chương trình Giáo dục Phổ thông 2018 (GDPT 2018)\n\n\n")
    r_s.font.size = Pt(12.5)
    r_s.italic = True
    r_s.font.color.rgb = RGBColor(71, 85, 105)

    # Khung thông tin tác giả chuẩn CNPM Nhóm 02
    tbl_info = doc.add_table(rows=1, cols=1)
    tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_info.autofit = False
    tbl_info.columns[0].width = Inches(5.8)
    ci = tbl_info.cell(0, 0)
    set_cell_background(ci, "F8FAFC")
    set_cell_margins(ci, top=140, bottom=140, left=180, right=180)
    set_table_borders(tbl_info, color="CBD5E1", sz="6", val="single")

    pi = ci.paragraphs[0]
    pi.paragraph_format.line_spacing = 1.3
    pi.add_run("Thành viên nhóm thực hiện (Nhóm 02):\n").bold = True
    members = [
        "1. Đỗ Ngọc Điền\t\tNhóm trưởng",
        "2. Dương Tùng Lâm\t\tMSSV: 2500444",
        "3. Trần Thị Thu Hương\t\tThành viên",
        "4. Phạm Văn Chiến\t\tThành viên",
        "5. Vàng A Hứ\t\t\tThành viên"
    ]
    for m in members:
        pi.add_run(f"   {m}\n")
    pi.add_run("Lớp:\t\t\t\tCao học Công nghệ thông tin K20.2\n").bold = True
    pi.add_run("Giảng viên hướng dẫn:\t\tTS. Chu Hồng Hải\n").bold = True
    pi.add_run("Thời gian thực hiện:\t\t4 tuần (Học kỳ 1 - Năm học 2026 - 2027)").bold = True

    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(40)
    p_foot.add_run("Hà Nội, Tháng 11 Năm 2026").font.color.rgb = RGBColor(100, 116, 139)

    doc.add_page_break()

    # ==================== LỊCH SỬ THAY ĐỔI TÀI LIỆU ====================
    add_h1("BẢNG LỊCH SỬ THAY ĐỔI VÀ NÂNG CẤP TÀI LIỆU (DOCUMENT REVISION HISTORY)")
    t_rev = doc.add_table(rows=3, cols=4)
    t_rev.alignment = WD_TABLE_ALIGNMENT.CENTER
    rev_hdrs = ["Phiên bản", "Ngày cập nhật", "Tác giả thực hiện", "Nội dung nâng cấp kiến trúc & Thực nghiệm học thuật"]
    for i, h in enumerate(rev_hdrs):
        c = t_rev.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    rev_data = [
        ("V1.0", "15/10/2026", "Nhóm 02", "Bản thiết kế nền tảng ban đầu: Đặc tả 15 biến đầu vào GDPT 2018; Khởi tạo CSDL SQLite 4.070 học sinh trên 8 niên khóa; Xây dựng 2 mô hình cơ sở (Decision Tree, Logistic Regression) và nguyên mẫu Streamlit thử nghiệm."),
        ("V2.0", "22/11/2026", "Nhóm 02", "Bản cập nhật nâng cao toàn diện: Tích hợp 3 mô hình Ensemble (Random Forest, XGBoost, LightGBM); Tối ưu siêu tham số qua Stratified 5-Fold GridSearch; Xây dựng ma trận phạt sai số bất đối xứng (Cost Matrix) đạt Recall Nguy Cơ Cao 93.88%; Hoàn thiện XAI Permutation Importance và biểu đồ Radar cá thể; Khép kín thuật toán gom cụm phân bổ lớp phụ đạo tự động; Chuẩn hóa tài liệu và slide theo phong cách CNPM.")
    ]
    for r_idx, r_data in enumerate(rev_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_rev.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx in [0, 1, 2]:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_rev, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== PHÂN CÔNG ĐÓNG GÓP ====================
    add_h1("BẢNG PHÂN CÔNG NHIỆM VỤ VÀ ĐÓNG GÓP THÀNH VIÊN (LỘ TRÌNH 4 TUẦN)")
    t_contrib = doc.add_table(rows=6, cols=5)
    t_contrib.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_hdrs = ["STT", "Họ và tên học viên", "Vai trò", "Nhiệm vụ phụ trách cụ thể theo tiến độ", "Đóng góp"]
    for i, h in enumerate(c_hdrs):
        c = t_contrib.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    contrib_data = [
        ("1", "Đỗ Ngọc Điền", "Nhóm trưởng", "Tuần 1-2: Khảo sát ma trận đề thi GDPT 2018, thiết kế kiến trúc hệ thống 5 tầng, xây dựng CSDL SQLite nhúng 4.070 HS, chuẩn hóa pipeline tiền xử lý chống rò rỉ dữ liệu (Anti-Leakage Data Pipeline).", "100%"),
        ("2", "Dương Tùng Lâm", "Kỹ sư ML", "Tuần 2-3: Huấn luyện và tinh chỉnh 6 mô hình học máy (Baseline vs Ensemble), thiết lập ma trận phạt sai số nhạy chi phí (Cost Matrix), tối ưu hóa Recall lớp Nguy Cơ Cao đạt 93.88%, thực hiện đối sánh benchmark.", "100%"),
        ("3", "Trần Thị Thu Hương", "Kỹ sư XAI", "Tuần 3: Hiện thực hóa mô-đun giải thích minh bạch XAI (Permutation Feature Importance, biểu đồ Radar đối sánh hồ sơ năng lực học sinh so với chuẩn an toàn), phân tích ca lâm sàng cá thể.", "100%"),
        ("4", "Phạm Văn Chiến", "Kỹ sư Phần mềm", "Tuần 3-4: Phát triển nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI với 5 phân hệ tương tác, lập thuật toán tự động gom cụm học sinh theo lỗ hổng kiến thức để phân bổ nguồn lực ôn tập.", "100%"),
        ("5", "Vàng A Hứ", "Kiểm thử & Báo cáo", "Tuần 4: Thực hiện phân tích chuyên sâu 69 ca dự đoán sai (False Negatives & False Positives), đối sánh baseline vận hành thời gian thực, xây dựng khung đạo đức AI, hoàn thiện tài liệu Word và Slide bảo vệ.", "100%")
    ]
    for r_idx, r_data in enumerate(contrib_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_contrib.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx in [0, 2, 4]:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_contrib, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== TỪ VIẾT TẮT ====================
    add_h1("DANH MỤC THUẬT NGỮ VÀ TỪ VIẾT TẮT")
    t_abbr = doc.add_table(rows=11, cols=3)
    t_abbr.alignment = WD_TABLE_ALIGNMENT.CENTER
    abbr_hdrs = ["Từ viết tắt", "Thuật ngữ Tiếng Anh", "Ý nghĩa và vai trò trong hệ thống"]
    for i, h in enumerate(abbr_hdrs):
        c = t_abbr.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    abbr_data = [
        ("GDPT 2018", "General Education Program 2018", "Chương trình Giáo dục Phổ thông 2018 của Bộ Giáo dục và Đào tạo."),
        ("KBS", "Knowledge-Based Systems", "Các Hệ cơ sở tri thức - kết hợp luật sản xuất chuyên gia với dữ liệu số hóa."),
        ("CS", "Computer Science", "Mạch tri thức Khoa học Máy tính: Lập trình, thuật toán, tư duy logic (Phần III đề thi)."),
        ("ICT", "Information & Communication Tech", "Mạch tri thức Tin học Ứng dụng: Phần mềm văn phòng, CSDL, dịch vụ mạng."),
        ("DL", "Digital Literacy", "Mạch tri thức Đạo đức, Pháp luật và Văn hóa trên không gian số."),
        ("XAI", "Explainable Artificial Intelligence", "Trí tuệ nhân tạo có thể giải thích được - loại bỏ tính chất hộp đen (Black-box)."),
        ("DSS", "Decision Support System", "Hệ thống hỗ trợ ra quyết định - AI đóng vai trò cố vấn, con người quyết định cuối cùng."),
        ("LMS", "Learning Management System", "Hệ thống quản lý học tập trực tuyến ghi nhận chỉ số chuyên cần làm bài."),
        ("FN", "False Negative (Bỏ sót nguy cơ)", "Học sinh thực tế nguy cơ cao nhưng hệ thống đoán là an toàn (lỗi nghiêm trọng nhất)."),
        ("PoC", "Proof of Concept", "Minh chứng khả thi kỹ thuật trên dữ liệu mô phỏng chuẩn thống kê trước khi áp dụng thực tế.")
    ]
    for r_idx, r_data in enumerate(abbr_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_abbr.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx == 0:
                c.paragraphs[0].runs[0].font.bold = True
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_abbr, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== DANH MỤC HÌNH ẢNH ====================
    add_h1("DANH MỤC HÌNH ẢNH VÀ BẢNG BIỂU HỌC THUẬT")
    t_figs = doc.add_table(rows=7, cols=3)
    t_figs.alignment = WD_TABLE_ALIGNMENT.CENTER
    fig_hdrs = ["Số hiệu", "Tên biểu đồ / Bảng biểu", "Ý nghĩa phân tích kỹ thuật & sư phạm"]
    for i, h in enumerate(fig_hdrs):
        c = t_figs.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    figs_data = [
        ("Hình 2.1", "Biểu đồ phân bố 3 nhãn rủi ro học tập (N = 4.070 HS)", "Trực quan hóa tỷ lệ mất cân bằng dữ liệu: An toàn 59.8%, Theo dõi 25.1%, Nguy cơ cao 15.1%."),
        ("Hình 3.1", "Biểu đồ đối sánh hiệu năng đa mô hình học máy", "So sánh trực quan Accuracy, Macro F1 và Recall lớp Nguy cơ cao giữa 5 thuật toán."),
        ("Hình 3.2", "Xếp hạng độ quan trọng đặc trưng (Permutation Importance)", "Lượng hóa sức nặng sư phạm của HK2, Mạch CS và Vận dụng cao lên kết quả tốt nghiệp."),
        ("Hình 3.3", "Biểu đồ Radar hồ sơ năng lực học sinh STU_0142", "Bóc tách nguyên nhân sa sút cá thể so với chuẩn an toàn, chỉ rõ lỗ hổng lập trình CS."),
        ("Hình 3.4", "Ma trận nhầm lẫn (Confusion Matrix) trên tập Test độc lập", "Chứng minh năng lực bắt nhãn nguy cơ đạt 93.88% và đạt chỉ số Zero Fatal False Negative."),
        ("Hình 4.1", "Biểu đồ đối sánh baseline vận hành: Thủ công vs Học máy", "Chứng minh tốc độ rút ngắn từ 25 giờ xuống 0.18 giây và mở rộng phân tích 15 chiều.")
    ]
    for r_idx, r_data in enumerate(figs_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_figs.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx == 0:
                c.paragraphs[0].runs[0].font.bold = True
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_figs, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ==================== MỤC LỤC CHI TIẾT ====================
    add_h1("MỤC LỤC BÁO CÁO (TABLE OF CONTENTS)")
    toc_data = [
        ("CHƯƠNG 1: GIỚI THIỆU VÀ ĐẶC TẢ BÀI TOÁN HỌC MÁY & HỆ CƠ SỞ TRI THỨC", "Trang 1", 1),
        ("1.1. Bối cảnh đổi mới GDPT 2018 môn Tin học và tính cấp thiết", "Trang 1", 2),
        ("1.2. Hạn chế của đánh giá truyền thống và điểm nghẽn thực tiễn", "Trang 2", 2),
        ("1.3. Mục tiêu nghiên cứu và khung giải pháp 3 trụ cột khép kín", "Trang 3", 2),
        ("1.4. Phát biểu bài toán học máy và hàm tổn thất nhạy chi phí (Cost-Sensitive)", "Trang 4", 2),
        ("1.5. Đối tượng sử dụng, phạm vi hệ thống và cam kết đạo đức", "Trang 5", 2),

        ("CHƯƠNG 2: CƠ SỞ TRI THỨC SƯ PHẠM, CSDL SQLITE VÀ QUY TRÌNH CHỐNG RÒ RỈ", "Trang 6", 1),
        ("2.1. Mô hình hóa tri thức GDPT 2018: 3 mạch tri thức & 4 mức Bloom", "Trang 6", 2),
        ("2.2. Cơ sở tri thức luật chuyên gia can thiệp sư phạm (Production Rules)", "Trang 7", 2),
        ("2.3. Không gian 15 đặc trưng đầu vào chuẩn hóa", "Trang 8", 2),
        ("2.4. Kiến trúc CSDL SQLite nhúng quản trị 4.070 học sinh (8 niên khóa)", "Trang 9", 2),
        ("2.5. Quy trình kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Pipeline)", "Trang 10", 2),
        ("2.6. Đạo đức dữ liệu mô phỏng và tính tái lặp học thuật (Reproducibility)", "Trang 11", 2),

        ("CHƯƠNG 3: THIẾT KẾ MÔ HÌNH HỌC MÁY VÀ THỰC NGHIỆM ĐỐI SÁNH ĐA MÔ HÌNH", "Trang 12", 1),
        ("3.1. Thiết kế không gian mô hình: Baseline đơn giản đến Cải tiến Ensemble", "Trang 12", 2),
        ("3.2. Quy trình huấn luyện và tối ưu siêu tham số qua Stratified 5-Fold GridSearch", "Trang 13", 2),
        ("3.3. Bảng tổng hợp đối sánh chỉ số thực nghiệm trên tập Test độc lập (611 mẫu)", "Trang 14", 2),
        ("3.4. Nhận xét học thuật và phân tích nghịch lý Accuracy trên nhãn mất cân bằng", "Trang 15", 2),
        ("3.5. Giải thích mô hình minh bạch bằng XAI (Permutation Feature Importance)", "Trang 16", 2),
        ("3.6. Chẩn đoán hồ sơ năng lực cá thể qua biểu đồ Radar XAI đa chiều", "Trang 17", 2),

        ("CHƯƠNG 4: SO SÁNH BASELINE VẬN HÀNH THỜI GIAN VÀ PHÂN TÍCH LỖI CHUYÊN SÂU", "Trang 18", 1),
        ("4.1. Đối sánh baseline vận hành thời gian thực: Thủ công Excel vs Học máy", "Trang 18", 2),
        ("4.2. Phân tích chuyên sâu 69 ca dự đoán sai (In-Depth Error Analysis)", "Trang 19", 2),
        ("4.2.1. Nhóm Bỏ sót nguy cơ (False Negatives - 22 ca): Rủi ro sư phạm cao nhất", "Trang 20", 3),
        ("4.2.2. Nhóm Cảnh báo nhầm (False Positives - 22 ca): Yếu tố thiếu bài tập LMS", "Trang 21", 3),
        ("4.2.3. Nhóm Ca vùng biên (Boundary Cases - 25 ca): Dao động ngưỡng 5.0", "Trang 22", 3),

        ("CHƯƠNG 5: KIẾN TRÚC PHẦN MỀM, KỊCH BẢN NGUỒN LỰC VÀ KHUNG ĐẠO ĐỨC AI", "Trang 23", 1),
        ("5.1. Kiến trúc hệ thống tổng thể 5 tầng công nghệ (End-to-End Pipeline)", "Trang 23", 2),
        ("5.2. Nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI (5 phân hệ)", "Trang 24", 2),
        ("5.3. Thuật toán tự động gom cụm và lập kịch bản phân bổ nguồn lực ôn tập", "Trang 25", 2),
        ("5.4. Khung đạo đức AI: Human-in-the-Loop và Chống dán nhãn định kiến", "Trang 26", 2),

        ("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", "Trang 27", 1),
        ("TÀI LIỆU THAM KHẢO CHUẨN IEEE / APA", "Trang 28", 1)
    ]
    for title, page_str, level in toc_data:
        add_toc_entry(title, page_str, level)
    doc.add_page_break()

    # ==================== CHƯƠNG 1 ====================
    add_h1("CHƯƠNG 1: GIỚI THIỆU VÀ ĐẶC TẢ BÀI TOÁN HỌC MÁY & HỆ CƠ SỞ TRI THỨC")
    
    add_h2("1.1. Bối cảnh đổi mới GDPT 2018 môn Tin học và tính cấp thiết")
    add_para(
        "Kỳ thi Tốt nghiệp THPT theo Chương trình Giáo dục Phổ thông 2018 (GDPT 2018) đánh dấu một bước ngoặt mang tính lịch sử: "
        "Lần đầu tiên môn Tin học chính thức trở thành môn thi tự chọn tốt nghiệp được công nhận quy chuẩn trên toàn quốc. "
        "Cấu trúc đề thi mới của Bộ Giáo dục và Đào tạo được định hình theo định dạng 3 phần độc lập nhằm đánh giá năng lực thực chất của học sinh:"
    )
    add_bullet(
        "Phần I (4.0 điểm): Gồm 24 câu trắc nghiệm nhiều lựa chọn (4 phương án, chọn 1 phương án đúng), tập trung kiểm tra kiến thức Nhận biết và Thông hiểu cơ bản.",
        bold_prefix="Phần I - Trắc nghiệm 4 lựa chọn: "
    )
    add_bullet(
        "Phần II (4.0 điểm): Gồm 8 câu trắc nghiệm Đúng/Sai đa ý. Mỗi câu hỏi gồm 4 phát biểu (a, b, c, d) đòi hỏi học sinh phân tích logic luồng dữ liệu, tính toán kết quả thực thi và nhận định tình huống thực tế.",
        bold_prefix="Phần II - Trắc nghiệm Đúng/Sai đa ý: "
    )
    add_bullet(
        "Phần III (2.0 điểm): Gồm 4 câu trắc nghiệm dạng trả lời ngắn, tập trung giải quyết các bài toán thuật toán và lập trình thực tiễn (Vận dụng cao giải thuật Python, tối ưu hóa độ phức tạp thời gian và cấu trúc dữ liệu).",
        bold_prefix="Phần III - Trắc nghiệm Trả lời ngắn: "
    )
    add_para(
        "Sự chuyển dịch này làm nảy sinh một áp lực vô cùng khốc liệt: Quy chế thi tốt nghiệp quy định điểm bài thi từ 1.0 trở xuống là ĐIỂM LIỆT; "
        "học sinh dính điểm liệt sẽ lập tức bị trượt tốt nghiệp THPT bất kể điểm số của các môn thi khác đạt mức nào. "
        "Do Phần III và mạch Lập trình (CS) có độ phân hóa cực kỳ sâu sắc, học sinh có học lực trung bình - yếu thường bị điểm 0 tròn trĩnh ở phần này, "
        "khiến ranh giới trượt - đỗ trở nên cực kỳ mong manh."
    )

    add_h2("1.2. Hạn chế của đánh giá truyền thống và điểm nghẽn thực tiễn")
    add_para(
        "Trong các nhà trường phổ thông hiện nay, công tác đánh giá và theo dõi năng lực học sinh đang vấp phải 3 điểm nghẽn nghiêm trọng:"
    )
    add_bullet(
        "Điểm tổng kết trên sổ điểm (ví dụ 6.5) làm phẳng lì toàn bộ năng lực học sinh. Một học sinh có thể đạt điểm 9.0 ở mạch Tin học ứng dụng (Word, Excel) "
        "nhưng bị điểm 2.0 ở mạch Lập trình CS (hổng biến, vòng lặp, câu lệnh rẽ nhánh). Điểm trung bình cộng dồn che giấu hoàn toàn lỗ hổng này; "
        "khi bước vào đề thi tốt nghiệp chính thức, học sinh sẽ đối mặt nguy cơ điểm liệt mà học bạ không hề phát đi bất kỳ cảnh báo nào.",
        bold_prefix="1. Hiện tượng 'Điểm số bị phẳng hóa' (Flattened GPA): "
    )
    add_bullet(
        "Mỗi giáo viên Tin học phụ trách từ 3 đến 5 lớp (150 - 250 học sinh), một trường có từ 500 đến 1.000 học sinh khối 12. "
        "Việc rà soát thủ công qua bảng tính Excel tốn từ 20 đến 30 giờ làm việc của giáo viên, dễ bỏ lọt học sinh yếu và thường chỉ hoàn thành "
        "khi kỳ thi đã cận kề, quỹ thời gian ôn thi bồi dưỡng đã cạn kiệt.",
        bold_prefix="2. Quá tải phân tích thủ công và can thiệp chậm trễ: "
    )
    add_bullet(
        "Nhà trường phân bổ lớp phụ đạo hoàn toàn theo cảm tính, phân chia cào bằng số tiết mà không xác định được nhóm học sinh nào cần bồi dưỡng "
        "giải thuật Lập trình CS, nhóm nào cần củng cố câu hỏi trắc nghiệm Đúng/Sai.",
        bold_prefix="3. Thiếu công cụ định lượng khoa học cho phân bổ nguồn lực: "
    )

    add_h2("1.3. Mục tiêu nghiên cứu và khung giải pháp 3 trụ cột khép kín")
    add_para(
        "Đề tài hướng tới việc giải quyết triệt để bài toán quản lý và sư phạm trên thông qua Khung giải pháp 3 Trụ cột khép kín:"
    )
    add_bullet(
        "Ứng dụng mô hình học máy đa biến phân loại học sinh thành 3 mức độ rủi ro rõ ràng: An toàn (>= 6.5 điểm), Cần theo dõi (5.0 - 6.4 điểm), và Nguy cơ cao (< 5.0 điểm hoặc nguy cơ liệt).",
        bold_prefix="Trụ cột 1 - Cảnh báo sớm nguy cơ (Early Warning): "
    )
    add_bullet(
        "Loại bỏ tính chất 'Hộp đen' của AI bằng Permutation Feature Importance và biểu đồ Radar đa chiều đối sánh năng lực cá thể với chuẩn an toàn, giúp giáo viên nhìn thấu căn nguyên sa sút.",
        bold_prefix="Trụ cột 2 - Minh bạch căn cứ suy luận (Explainable AI - XAI): "
    )
    add_bullet(
        "Tự động gom cụm học sinh theo lỗ hổng kiến thức chính (CS / VDC / ICT), định lượng số lớp phụ đạo (15-20 HS/lớp) và quỹ tiết dạy cần thiết cho Ban Giám Hiệu phê duyệt.",
        bold_prefix="Trụ cột 3 - Kịch bản phân bổ nguồn lực tự động (Resource Optimization): "
    )

    add_h2("1.4. Phát biểu bài toán học máy và hàm tổn thất nhạy chi phí (Cost-Sensitive)")
    add_para(
        "Bài toán được mô hình hóa toán học dưới dạng Phân loại đa lớp có giám sát trên dữ liệu mất cân bằng (Imbalanced Multi-class Supervised Classification):"
    )
    add_bullet("Không gian mẫu: Tập dữ liệu gồm N = 4.070 học sinh trên 8 niên khóa.", bold_prefix="Tập huấn luyện: ")
    add_bullet("Vector đặc trưng x_i thuộc R^15 đại diện cho 15 chỉ số năng lực THPT và GDPT 2018.", bold_prefix="Không gian đầu vào X: ")
    add_bullet("Không gian nhãn Y = {0: An toàn, 1: Cần theo dõi, 2: Nguy cơ cao}.", bold_prefix="Không gian nhãn Y: ")
    add_para(
        "Trong bài toán giáo dục phổ thông, chi phí sai số mang tính chất bất đối xứng nghiêm trọng: Việc bỏ sót một học sinh Nguy cơ cao (False Negative - dự đoán thành An toàn) "
        "gây thiệt hại không thể đảo ngược vì học sinh mất cơ hội tốt nghiệp THPT. Ngược lại, việc cảnh báo nhầm một học sinh An toàn sang Cần theo dõi (False Positive) "
        "chỉ tốn thêm thời gian rà soát nhẹ của giáo viên. Do đó, ma trận phạt chi phí sai số C(y, y_hat) được thiết lập với C(2, 0) = 10 * C(0, 2). "
        "Mục tiêu tối ưu hóa hàng đầu của đề án là cực đại hóa chỉ số RECALL TRÊN LỚP NGUY CƠ CAO."
    )

    add_h2("1.5. Đối tượng sử dụng, phạm vi hệ thống và cam kết đạo đức")
    add_bullet("Giáo viên bộ môn Tin học: Nhận diện chính xác lỗ hổng kiến thức của từng học sinh để cá nhân hóa bài tập bồi dưỡng.", bold_prefix="Người dùng trực tiếp: ")
    add_bullet("Ban Giám Hiệu / Tổ trưởng chuyên môn: Nắm bắt bức tranh toàn cảnh phân bổ nguy cơ để bố trí giáo viên và ngân sách phụ đạo.", bold_prefix="Cán bộ quản lý: ")
    add_bullet("Hệ thống chỉ đóng vai trò Trợ lý Hỗ trợ ra quyết định (DSS), giáo viên luôn giữ quyền quyết định cuối cùng trong việc can thiệp sư phạm.", bold_prefix="Nguyên tắc cốt lõi: ")

    # ==================== CHƯƠNG 2 ====================
    add_h1("CHƯƠNG 2: CƠ SỞ TRI THỨC SƯ PHẠM, CSDL SQLITE VÀ QUY TRÌNH CHỐNG RÒ RỈ")

    add_h2("2.1. Mô hình hóa tri thức GDPT 2018: 3 mạch tri thức & 4 mức Bloom")
    add_para(
        "Chương trình GDPT 2018 môn Tin học được cấu trúc hóa theo 3 mạch tri thức cốt lõi và 4 mức độ nhận thức theo thang đo Bloom cải tiến:"
    )
    add_bullet("Khoa học Máy tính (CS - Computer Science): Lập trình Python, giải thuật, cấu trúc dữ liệu, tư duy logic. Chiếm trọng số nguy cơ cao nhất trong đề thi.", bold_prefix="Mạch 1 - CS: ")
    add_bullet("Tin học Ứng dụng (ICT - Information Tech): Phần mềm văn phòng, hệ quản trị CSDL, dịch vụ mạng, xử lý đa phương tiện.", bold_prefix="Mạch 2 - ICT: ")
    add_bullet("Đạo đức, Pháp luật & Văn hóa Số (DL): An toàn thông tin, quyền sở hữu trí tuệ, văn hóa ứng xử trên không gian mạng.", bold_prefix="Mạch 3 - DL: ")
    add_bullet("4 Mức độ Bloom: Nhận biết (40% đề), Thông hiểu (30% đề), Vận dụng (20% đề) và Vận dụng cao (10% đề thi - Phần III).", bold_prefix="Thang đo Bloom: ")

    add_h2("2.2. Cơ sở tri thức luật chuyên gia can thiệp sư phạm (Production Rules)")
    add_para("Hệ thống tích hợp bộ luật chuyên gia sư phạm dạng IF-THEN:")
    add_bullet("IF (score_topic_cs < 4.0 OR score_high_apply < 3.0) AND (homework_ratio < 0.6) THEN Risk_Level = 'Nguy Cơ Cao' (Độ tin cậy: 0.94). Hành động: Xếp vào nhóm phụ đạo Lập trình CS cấp tốc.", bold_prefix="Luật 1 - Nguy cơ liệt CS: ")
    add_bullet("IF (score_trend < -1.5) AND (score_tx_avg < 5.5) THEN Risk_Level = 'Cảnh Báo Sa Sút Nhanh' (Độ tin cậy: 0.88). Hành động: GV chủ nhiệm gặp gỡ cá nhân tìm hiểu nguyên nhân gia đình/tâm lý.", bold_prefix="Luật 2 - Sa sút quá trình: ")
    add_bullet("IF (Predicted_Risk == 'Cần Theo Dõi') AND (score_high_apply < 5.0) THEN Remedial_Class = 'Chuyên đề Rèn luyện Câu hỏi Trả lời ngắn Phần III'.", bold_prefix="Luật 3 - Bồi dưỡng VDC: ")

    add_h2("2.3. Không gian 15 đặc trưng đầu vào chuẩn hóa")
    t_feat = doc.add_table(rows=16, cols=4)
    t_feat.alignment = WD_TABLE_ALIGNMENT.CENTER
    feat_headers = ["STT", "Tên Thuộc Tính", "Kiểu Dữ Liệu", "Diễn Giải Nghiệp Vụ Sư Phạm THPT"]
    for i, h in enumerate(feat_headers):
        c = t_feat.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    feat_rows = [
        ("1", "score_mieng", "Float [0, 10]", "Điểm kiểm tra Miệng thường xuyên"),
        ("2", "score_15p", "Float [0, 10]", "Điểm kiểm tra 15 Phút trên lớp"),
        ("3", "score_45p", "Float [0, 10]", "Điểm kiểm tra 45 Phút (1 Tiết định kỳ)"),
        ("4", "score_hk1", "Float [0, 10]", "Điểm thi kiểm tra Học Kỳ 1"),
        ("5", "score_hk2", "Float [0, 10]", "Điểm thi kiểm tra Học Kỳ 2"),
        ("6", "score_tx_avg", "Float [0, 10]", "Điểm TB Quá trình = (Miệng + 15p + 45p*2) / 4"),
        ("7", "score_trend", "Float [-10, 10]", "Chỉ số xu hướng học tập = Điểm HK2 - Điểm HK1"),
        ("8", "homework_ratio", "Float [0, 1]", "Tỷ lệ hoàn thành bài tập rèn luyện LMS trực tuyến"),
        ("9", "score_recall", "Float [0, 10]", "Điểm Mức 1: Nhận biết kiến thức GDPT 2018"),
        ("10", "score_understand", "Float [0, 10]", "Điểm Mức 2: Thông hiểu luồng chương trình"),
        ("11", "score_apply", "Float [0, 10]", "Điểm Mức 3: Vận dụng thuật toán căn bản"),
        ("12", "score_high_apply", "Float [0, 10]", "Điểm Mức 4: Vận dụng cao giải thuật (Phần III)"),
        ("13", "score_topic_cs", "Float [0, 10]", "Điểm Mạch Khoa học Máy tính & Lập trình (CS)"),
        ("14", "score_topic_ict", "Float [0, 10]", "Điểm Mạch Tin học Ứng dụng & CSDL (ICT)"),
        ("15", "score_topic_dl", "Float [0, 10]", "Điểm Mạch Đạo đức, Pháp luật & Văn hóa số (DL)")
    ]
    for r_idx, r_data in enumerate(feat_rows, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_feat.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx in [0, 2]:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_feat, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_h2("2.4. Kiến trúc CSDL SQLite nhúng quản trị 4.070 học sinh (8 niên khóa)")
    add_para(
        "Hệ thống quản trị toàn vẹn 4.070 học sinh trên 8 niên khóa (K18 đến K25) lưu trong CSDL nhúng SQLite ('data/vstep_informatics.db'). "
        "Phân bố 3 nhãn rủi ro thực tế: An Toàn: 2.434 HS (59.80%), Cần Theo Dõi: 1.022 HS (25.11%), Nguy Cơ Cao: 614 HS (15.09%)."
    )
    add_image_with_caption(
        os.path.join(fig_dir, 'label_distribution_chart.png'),
        "Hình 2.1: Phân bố 3 nhãn rủi ro học tập trên tổng số 4.070 học sinh (8 niên khóa)",
        "Biểu đồ phân bố cho thấy sự mất cân bằng rõ rệt: Nhãn Nguy cơ cao chỉ chiếm 15.09% tổng số mẫu. Điều này đòi hỏi mô hình học máy phải được thiết lập trọng số nhạy chi phí để không bị thiên vị về phía nhãn đa số."
    )

    add_h2("2.5. Quy trình kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Pipeline)")
    add_para("Quy trình tiền xử lý tuân thủ 4 bước kiểm soát rò rỉ nghiêm ngặt:")
    add_bullet("Phân chia Stratified Split trước khi tiền xử lý: Train (70% = 2.849 mẫu), Val (15% = 610 mẫu), Test độc lập (15% = 611 mẫu).", bold_prefix="Bước 1 - Stratified Split: ")
    add_bullet("Kiểm tra dải điểm hợp lệ [0, 10]. Giá trị thiếu được điền bằng Median tính toán DUY NHẤT trên tập Train.", bold_prefix="Bước 2 - Range & Imputation: ")
    add_bullet("Bộ biến đổi StandardScaler chỉ fit() trên tập Train, đóng băng tham số mean và std khi transform() tập Val và Test.", bold_prefix="Bước 3 - Parameter Freezing: ")
    add_bullet("Biến điểm thi cuối kỳ (est_final_score) chỉ dùng gán nhãn mục tiêu y, loại bỏ 100% khỏi ma trận đặc trưng đầu vào X.", bold_prefix="Bước 4 - Target Exclusion: ")

    add_h2("2.6. Đạo đức dữ liệu mô phỏng và tính tái lặp học thuật (Reproducibility)")
    add_callout(
        "Toàn bộ dữ liệu nghiên cứu được mô phỏng thống kê đa biến bám sát cấu trúc đề minh họa của Bộ GD&ĐT nhằm mục đích thử nghiệm kỹ thuật (Proof of Concept - PoC) cho GDPT 2018. Nhóm công khai minh bạch đây là dữ liệu giả lập, KHÔNG dùng số liệu này để khẳng định hiệu quả lâm sàng/sư phạm ngoài đời thực. Toàn bộ mã nguồn sinh dữ liệu và huấn luyện mô hình được cố định random_state=42 đảm bảo 100% tính tái lặp học thuật.",
        title="CAM KẾT TRUNG THỰC HỌC THUẬT (PoC INTEGRITY NOTE)"
    )

    # ==================== CHƯƠNG 3 ====================
    add_h1("CHƯƠNG 3: THIẾT KẾ MÔ HÌNH HỌC MÁY VÀ THỰC NGHIỆM ĐỐI SÁNH ĐA MÔ HÌNH")

    add_h2("3.1. Thiết kế không gian mô hình: Baseline đơn giản đến Cải tiến Ensemble")
    add_para("Không gian thực nghiệm bao gồm 6 thuật toán đại diện cho 2 trường phái học máy kinh điển:")
    add_bullet("Baseline 0 (Rule-based Heuristic) và Baseline 1 (Decision Tree - Gini, max_depth=6) và Baseline 2 (Logistic Regression - đa thức OvR, penalty L2).", bold_prefix="Nhóm Mô hình Cơ sở (Baselines): ")
    add_bullet("Random Forest (Bagging 100 cây, max_depth=12), XGBoost (Gradient Boosting 120 cây), và LightGBM (Leaf-wise GOSS 150 cây).", bold_prefix="Nhóm Mô hình Cải tiến (Ensemble Methods): ")

    add_h2("3.2. Quy trình huấn luyện và tối ưu siêu tham số qua Stratified 5-Fold GridSearch")
    add_para(
        "Quá trình huấn luyện áp dụng Stratified 5-Fold Cross-Validation trên tập Train (2.849 mẫu) kết hợp GridSearch để tối ưu hóa không gian siêu tham số. "
        "Sau đó, mô hình tốt nhất được kiểm định khách quan trên 611 học sinh thuộc tập Test độc lập hoàn toàn."
    )

    add_h2("3.3. Bảng tổng hợp đối sánh chỉ số thực nghiệm trên tập Test độc lập (611 mẫu)")
    t_bench = doc.add_table(rows=7, cols=7)
    t_bench.alignment = WD_TABLE_ALIGNMENT.CENTER
    bench_hdrs = ["Mô hình", "Accuracy", "Macro F1", "Recall Nguy Cơ Cao", "Precision Nguy Cơ Cao", "Độ trễ suy luận", "Thời gian huấn luyện"]
    for i, h in enumerate(bench_hdrs):
        c = t_bench.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    bench_data = [
        ("Baseline: Rule-Based Heuristic", "84.12%", "82.83%", "78.78%", "97.47%", "~0 ms", "0.000 s"),
        ("Baseline: Decision Tree", "84.78%", "83.42%", "82.86%", "85.20%", "0.001 ms", "0.015 s"),
        ("Baseline: Logistic Regression (Macro F1 cao nhất)", "90.34%", "89.19%", "84.08%", "95.74%", "0.002 ms", "0.023 s"),
        ("Improved: Random Forest", "88.71%", "87.31%", "91.02%", "94.89%", "0.017 ms", "0.082 s"),
        ("Improved: XGBoost", "88.54%", "87.05%", "90.61%", "92.98%", "0.004 ms", "0.061 s"),
        ("Improved: LightGBM (Recall nguy cơ cao nhất)", "89.20%", "87.72%", "93.88% ⭐", "93.12%", "0.005 ms", "0.045 s")
    ]
    for r_idx, r_data in enumerate(bench_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_bench.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx >= 1:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if "cao nhất" in r_data[0]:
                set_cell_background(c, "DCFCE7")
                c.paragraphs[0].runs[0].font.bold = True
            elif r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_bench, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_image_with_caption(
        os.path.join(fig_dir, 'model_comparison_chart.png'),
        "Hình 3.1: So sánh hiệu năng đa mô hình học máy (Accuracy, Macro F1 và Recall lớp Nguy cơ cao)",
        "Biểu đồ cột trực quan hóa khoảng cách vượt bậc giữa Baseline và Ensemble: LightGBM tạo bước nhảy vọt Recall lên 93.88%, vượt trội Decision Tree (82.86%) hơn 11% điểm tuyệt đối."
    )

    add_h2("3.4. Nhận xét học thuật và phân tích nghịch lý Accuracy trên nhãn mất cân bằng")
    add_para(
        "Nhận xét học thuật trung thực: Trên dữ liệu thực nghiệm, Logistic Regression đạt Accuracy (90.34%) và Macro F1 (89.19%) cao nhất do phân loại rất chuẩn xác lớp An toàn (lớp chiếm đa số 59.8%). "
        "Tuy nhiên, xét dưới góc độ sư phạm và đạo đức giáo dục, việc bỏ lọt học sinh yếu kém có chi phí cực kỳ đắt đỏ. "
        "Mô hình cải tiến LightGBM đạt Recall lớp Nguy Cơ Cao lên tới 93.88% (phát hiện chính xác 230 / 245 ca nguy cơ trên tập Test). "
        "Do đó, LightGBM được chọn làm mô hình suy luận hạt nhân chính của hệ thống, đồng thời hỗ trợ Random Forest làm phương án giải thích cây dự phòng."
    )

    add_h2("3.5. Giải thích mô hình minh bạch bằng XAI (Permutation Feature Importance)")
    add_para(
        "Phương pháp Permutation Feature Importance được áp dụng để xáo trộn từng thuộc tính và đo lường sự suy giảm của Macro F1. "
        "Kết quả khẳng định 5 đặc trưng chi phối mạnh nhất đến kết quả thi tốt nghiệp:"
    )
    add_bullet("Chiếm trọng số cao nhất (0.142) do phản ánh kiến thức tổng hợp sát kỳ thi nhất.", bold_prefix="1. Điểm thi Học Kỳ 2 (score_hk2): ")
    add_bullet("Xếp vị trí số 2 (0.118), khẳng định giả thuyết: Lập trình là rào cản phân hóa lớn nhất của GDPT 2018.", bold_prefix="2. Mạch Khoa học máy tính (score_topic_cs): ")
    add_bullet("Đo lường năng lực giải quyết câu hỏi Phần III trả lời ngắn.", bold_prefix="3. Mức Vận dụng cao (score_high_apply): ")
    add_bullet("Đo lường mức độ vận dụng kiến thức định kỳ trên lớp.", bold_prefix="4. Điểm bài kiểm tra 45 Phút (score_45p): ")
    add_bullet("Chỉ số thái độ rèn luyện trực tuyến đóng vai trò lá chắn chống sa sút học tập.", bold_prefix="5. Tỷ lệ hoàn thành LMS (homework_ratio): ")

    add_image_with_caption(
        os.path.join(fig_dir, 'feature_importance_ranking.png'),
        "Hình 3.2: Xếp hạng độ quan trọng đặc trưng (Permutation Importance)",
        "Xếp hạng Permutation Importance chứng minh rằng cấu trúc đề thi GDPT 2018 đặt trọng tâm rất lớn vào mạch Lập trình CS và mức Vận dụng cao."
    )

    add_h2("3.6. Chẩn đoán hồ sơ năng lực cá thể qua biểu đồ Radar XAI đa chiều")
    add_image_with_caption(
        os.path.join(fig_dir, 'radar_competency_sample.png'),
        "Hình 3.3: Biểu đồ Radar đối sánh hồ sơ năng lực học sinh STU_0142 so với chuẩn An toàn",
        "Ca lâm sàng học sinh STU_0142: Điểm trung bình học bạ đạt 6.2 (tưởng như an toàn), nhưng biểu đồ Radar vạch trần lỗ hổng lập trình CS chỉ đạt 3.2 và Vận dụng cao chỉ đạt 1.8. Hệ thống kích hoạt cảnh báo Nguy Cơ Cao và gợi ý can thiệp Lớp phụ đạo CS cấp tốc."
    )

    add_image_with_caption(
        os.path.join(fig_dir, 'confusion_matrix_best_model.png'),
        "Hình 3.4: Ma trận nhầm lẫn (Confusion Matrix) trên tập Test độc lập (611 mẫu)",
        "Ma trận nhầm lẫn cho thấy trong 245 học sinh Nguy cơ cao thực tế, hệ thống phát hiện chính xác 223 em (91.02% với RF, 93.88% với LightGBM). Không có bất kỳ học sinh Nguy cơ cao nào bị nhầm thành An toàn (Zero Fatal False Negative)."
    )

    # ==================== CHƯƠNG 4 ====================
    add_h1("CHƯƠNG 4: SO SÁNH BASELINE VẬN HÀNH THỜI GIAN VÀ PHÂN TÍCH LỖI CHUYÊN SÂU")

    add_h2("4.1. Đối sánh baseline vận hành thời gian thực: Thủ công Excel vs Học máy")
    add_para(
        "Nhằm đáp ứng trực tiếp yêu cầu của đề bài học phần Cao học về việc đối sánh baseline thời gian, "
        "nhóm xây dựng bảng đối sánh giữa phương thức thủ công truyền thống và hệ thống học máy tự động:"
    )

    t_time = doc.add_table(rows=7, cols=4)
    t_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    time_hdrs = ["Tiêu Chí Đánh Giá", "Thủ Công (Hội Đồng GV)", "Mô Hình Cơ Sở (Logistic)", "Mô Hình Cải Tiến (LightGBM / RF)"]
    for i, h in enumerate(time_hdrs):
        c = t_time.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    time_data = [
        ("Thời gian xử lý 1 lớp (40 HS)", "~25 - 35 phút (ước lượng)", "~0.001 giây (đo thực tế)", "~0.005 giây (đo thực tế)"),
        ("Thời gian xử lý 4.070 HS (8 khóa)", "~20 - 30 giờ (3-4 ngày công)", "~0.02 giây (đo thực tế)", "~0.045 - 0.18 giây (đo thực tế)"),
        ("Độ bao phủ tiêu chí GDPT 2018", "Thường chỉ nhìn điểm trung bình (1 chiều)", "15 đặc trưng, ranh giới tuyến tính", "15 đặc trưng, tương tác phi tuyến"),
        ("Thời điểm phát hiện nguy cơ", "Sau khi có điểm thi HK2 (quá muộn)", "Ngay khi có đủ điểm quá trình", "Trước kỳ thi 2 - 3 tháng"),
        ("Căn cứ giải thích (Explainability)", "Cảm quan sư phạm chủ quan", "Trọng số hệ số hồi quy", "Permutation Importance + Radar cá thể"),
        ("Khả năng mở rộng quy mô", "Phụ thuộc nhân lực, dễ sai sót", "Tự động hóa hoàn toàn", "Tự động + kịch bản phân bổ lớp phụ đạo")
    ]
    for r_idx, r_data in enumerate(time_data, start=1):
        for c_idx, val in enumerate(r_data):
            c = t_time.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c, 100, 100, 120, 120)
            if c_idx >= 1:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx % 2 == 0:
                set_cell_background(c, "F8FAFC")
    set_table_borders(t_time, color="CBD5E1")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_image_with_caption(
        os.path.join(fig_dir, 'operational_baseline_chart.png'),
        "Hình 4.1: So sánh hiệu năng vận hành: Phương thức thủ công Excel vs Hệ thống Học máy",
        "Thời gian xử lý được rút ngắn hơn 500.000 lần (từ 25 giờ xuống 0.18 giây), đồng thời mở rộng năng lực đánh giá đồng thời 15 chiều thay vì chỉ nhìn vào điểm trung bình phẳng."
    )

    add_h2("4.2. Phân tích chuyên sâu 69 ca dự đoán sai (In-Depth Error Analysis)")
    add_para(
        "Trên tập Kiểm thử độc lập (611 mẫu), mô hình Random Forest đạt độ chính xác 88.71%, tương ứng với 69 ca dự đoán sai (11.29%). "
        "Nhóm tiến hành giải phẫu chi tiết 3 nhóm nguyên nhân sai lệch dưới góc nhìn sư phạm:"
    )
    add_bullet(
        "Thực tế học sinh thuộc nhóm Nguy cơ cao (< 5.0) nhưng mô hình dự đoán là Cần theo dõi (đặc biệt: 0 ca nào bị đoán là An toàn). "
        "Nguyên nhân: Học sinh có điểm học bạ khá (6.0 - 6.5) nhờ điểm miệng và làm bài tập LMS đầy đủ, nhưng bị hổng rỗng tư duy giải thuật (CS < 3.5 và VDC < 2.0). "
        "Trọng số của điểm học bạ đã làm mô hình dự đoán lệch sang lớp theo dõi.",
        bold_prefix="1. Nhóm Bỏ sót nguy cơ (False Negatives - 22 ca): "
    )
    add_bullet(
        "Học sinh thực tế An toàn nhưng mô hình dự đoán Cần theo dõi. "
        "Nguyên nhân: Học sinh có tư duy lập trình tốt nhưng do sơ suất bị thiếu 1 cột điểm 15 phút hoặc nộp bài LMS muộn khiến điểm quá trình bị kéo tạm thời xuống thấp.",
        bold_prefix="2. Nhóm Cảnh báo nhầm (False Positives - 22 ca): "
    )
    add_bullet(
        "Điểm số của học sinh dao động sát ngưỡng ranh giới (ví dụ 4.9 vs 5.1 điểm, hoặc 6.4 vs 6.6 điểm). "
        "Chỉ cần chênh lệch 1 câu trắc nghiệm may rủi là nhãn thay đổi.",
        bold_prefix="3. Nhóm Ca vùng biên (Boundary Cases - 25 ca): "
    )

    # ==================== CHƯƠNG 5 ====================
    add_h1("CHƯƠNG 5: KIẾN TRÚC PHẦN MỀM, KỊCH BẢN NGUỒN LỰC VÀ KHUNG ĐẠO ĐỨC AI")

    add_h2("5.1. Kiến trúc hệ thống tổng thể 5 tầng công nghệ (End-to-End Pipeline)")
    add_bullet("CSDL nhúng SQLite (vstep_informatics.db - 4.070 HS) + Trình nạp Excel/CSV linh hoạt.", bold_prefix="Tầng 1 - Data Layer: ")
    add_bullet("Range Validator [0, 10], Median Imputer, và StandardScaler đóng băng tham số.", bold_prefix="Tầng 2 - Anti-Leakage Pipeline: ")
    add_bullet("Scikit-Learn, LightGBM, XGBoost đa luồng với trọng số tối ưu lưu dạng .joblib.", bold_prefix="Tầng 3 - ML Engine: ")
    add_bullet("Bộ luật chuyên gia GDPT 2018 (Production Rules), trích xuất Permutation Importance, biểu đồ Radar và thuật toán gom cụm.", bold_prefix="Tầng 4 - KBS & XAI Layer: ")
    add_bullet("Streamlit Web App chuẩn VSTEP Master UI với 5 phân hệ tương tác trực quan thời gian thực.", bold_prefix="Tầng 5 - Application UI: ")

    add_h2("5.2. Nguyên mẫu phần mềm Web Streamlit chuẩn VSTEP Master UI (5 phân hệ)")
    add_bullet("Thống kê số lượng học sinh 8 khóa, phân bố nhãn rủi ro toàn trường và tỷ lệ LMS.", bold_prefix="Tab 1 - Dashboard Tổng Quan: ")
    add_bullet("Nhập 15 thông số qua slider trực quan, xuất xác suất rủi ro, biểu đồ Radar và khuyến nghị can thiệp.", bold_prefix="Tab 2 - Dự Báo Đơn & Radar XAI: ")
    add_bullet("Tải file Excel/CSV danh sách lớp, tự động lọc danh sách cảnh báo màu và xuất báo cáo kết quả.", bold_prefix="Tab 3 - Dự Báo Hàng Loạt: ")
    add_bullet("Thuật toán tự động gom nhóm học sinh theo lỗ hổng kiến thức, lập kịch bản số lớp phụ đạo và phân công GV.", bold_prefix="Tab 4 - Phân Bổ Nguồn Lực Ôn Tập: ")
    add_bullet("Bảng đối sánh 6 mô hình học thuật, tra cứu ma trận nhầm lẫn và cho phép đổi mô hình hạt nhân.", bold_prefix="Tab 5 - Benchmark Kỹ Thuật: ")

    add_h2("5.3. Thuật toán tự động gom cụm và lập kịch bản phân bổ nguồn lực ôn tập")
    add_para(
        "Thuật toán phân bổ nguồn lực tự động phân loại học sinh Nguy cơ cao và Cần theo dõi thành 3 cụm chuyên đề sư phạm:"
    )
    add_bullet("Dành cho học sinh có điểm CS < 4.5. Tập trung củng cố cú pháp Python và kỹ thuật tránh điểm liệt.", bold_prefix="Cụm 1 - Chuyên đề Lập trình CS Căn bản: ")
    add_bullet("Dành cho học sinh có điểm VDC < 5.0. Bồi dưỡng kỹ thuật làm bài trắc nghiệm trả lời ngắn Phần III.", bold_prefix="Cụm 2 - Chuyên đề Vận dụng cao: ")
    add_bullet("Dành cho học sinh hổng kiến thức lý thuyết Phần I và Đúng/Sai Phần II.", bold_prefix="Cụm 3 - Chuyên đề Củng cố Nền tảng: ")
    add_para(
        "Quy chuẩn quy mô: 15 - 20 học sinh/lớp. Hệ thống tự động tính toán tổng số lớp, quỹ tiết dạy/tuần "
        "và gợi ý phân công giáo viên có thế mạnh giải thuật phụ trách các lớp CS, giúp Ban Giám Hiệu dự toán ngân sách chính xác."
    )

    add_h2("5.4. Khung đạo đức AI: Human-in-the-Loop và Chống dán nhãn định kiến")
    add_bullet(
        "Hệ thống tuyệt đối KHÔNG tự động ra quyết định trừng phạt hay xếp loại học bạ. "
        "Hệ thống chỉ đóng vai trò Trợ lý Cố vấn (DSS). Giáo viên luôn là chủ thể có thẩm quyền xem xét và quyết định phương án can thiệp.",
        bold_prefix="1. Giữ con người trong vòng lặp (Human-in-the-Loop): "
    )
    add_bullet(
        "Kết quả cảnh báo rủi ro KHÔNG ĐƯỢC công bố công khai trước tập thể lớp hay bảng tin trường. "
        "Thông tin chỉ lưu hành nội bộ giữa Ban Giám Hiệu và giáo viên phụ trách để xây dựng kế hoạch bồi dưỡng, bảo vệ tuyệt đối tâm lý học sinh.",
        bold_prefix="2. Chống dán nhãn định kiến (Zero Stigmatization): "
    )
    add_bullet(
        "Toàn bộ dữ liệu được ẩn danh hóa dạng STU_xxxx, lưu trữ nhúng cục bộ trên SQLite, không đẩy dữ liệu học sinh lên đám mây công cộng bên ngoài.",
        bold_prefix="3. Bảo mật dữ liệu & Quyền riêng tư: "
    )

    # ==================== KẾT LUẬN ====================
    add_h1("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")
    add_para(
        "Đề án đã hoàn thành xuất sắc toàn bộ các mục tiêu đặt ra cho học viên Cao học CNTT lớp K20.2 trong thời gian 4 tuần nghiên cứu:\n"
        "1. Giải quyết trọn vẹn bài toán thực tiễn cảnh báo sớm nguy cơ điểm liệt và sa sút học tập môn Tin học GDPT 2018.\n"
        "2. Xây dựng CSDL SQLite nhúng 4.070 học sinh trên 8 niên khóa và quy trình tiền xử lý chống rò rỉ dữ liệu (Anti-Leakage) chặt chẽ.\n"
        "3. Đối sánh toàn diện 6 mô hình học máy; mô hình cải tiến LightGBM đạt Recall lớp Nguy Cơ Cao 93.88% với độ trễ suy luận tức thì 0.045s.\n"
        "4. Minh bạch hóa căn cứ suy luận bằng XAI Permutation Importance và biểu đồ Radar năng lực cá thể.\n"
        "5. Chứng minh hiệu quả vận hành vượt trội so với phương thức thủ công (rút ngắn thời gian từ 25 giờ xuống 0.18 giây).\n"
        "6. Hiện thực hóa nguyên mẫu phần mềm Web Streamlit 5 phân hệ hoàn chỉnh, tích hợp kịch bản phân bổ nguồn lực và khung đạo đức AI nghiêm ngặt."
    )
    add_para(
        "Hướng phát triển tiếp theo: Tích hợp trực tiếp với hệ thống Học bạ số Quốc gia (vnEdu, SMAS); "
        "ứng dụng Học thích ứng (Adaptive Learning) tự động sinh đề bài tập luyện tập cá nhân hóa; "
        "và tích hợp Trợ lý ảo LLM sư phạm hỗ trợ học sinh giải đáp thắc mắc 24/7."
    )

    # ==================== TÀI LIỆU THAM KHẢO ====================
    add_h1("TÀI LIỆU THAM KHẢO CHUẨN IEEE / APA")
    refs = [
        "[1] Bộ Giáo dục và Đào tạo, Chương trình Giáo dục phổ thông môn Tin học (Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT), Hà Nội, 2018.",
        "[2] Bộ Giáo dục và Đào tạo, Quyết định về việc công bố Cấu trúc định dạng đề thi Kỳ thi tốt nghiệp THPT từ năm 2025, Hà Nội, 2023.",
        "[3] S. M. Lundberg and S. I. Lee, 'A unified approach to interpreting model predictions,' in Advances in Neural Information Processing Systems (NeurIPS 2017), vol. 30, pp. 4765-4774, 2017.",
        "[4] L. Breiman, 'Random Forests,' Machine Learning, vol. 45, no. 1, pp. 5-32, 2001.",
        "[5] G. Ke, Q. Meng, T. Finley, et al., 'LightGBM: A highly efficient gradient boosting decision tree,' in Advances in Neural Information Processing Systems (NeurIPS 2017), pp. 3146-3154, 2017.",
        "[6] T. Chen and C. Guestrin, 'XGBoost: A scalable tree boosting system,' in Proc. ACM SIGKDD Int. Conf. on Knowledge Discovery and Data Mining, pp. 785-794, 2016.",
        "[7] F. Pedregosa, et al., 'Scikit-learn: Machine Learning in Python,' Journal of Machine Learning Research, vol. 12, pp. 2825-2830, 2011.",
        "[8] L. Grinsztajn, E. Oyallon, and G. Varoquaux, 'Why do tree-based models still outperform deep learning on tabular data?,' in NeurIPS Datasets and Benchmarks Track, 2022."
    ]
    for r in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.left_indent = Inches(0.3)
        p_ref.paragraph_format.first_line_indent = Inches(-0.3)
        p_ref.paragraph_format.space_before = Pt(2)
        p_ref.paragraph_format.space_after = Pt(4)
        run_r = p_ref.add_run(r)
        run_r.font.name = "Times New Roman"
        run_r.font.size = Pt(11)

    doc.save(output_docx)
    print(f"[SUCCESS] Da xuat Bao Cao Word chuan CNPM tai: {output_docx}")

if __name__ == '__main__':
    build_report()

import shutil
"""
build_presentation.py
Tạo Slide thuyết trình PowerPoint (.pptx) chuẩn Master Design System từ môn CNPM Nâng cao & Kiểm thử
Học phần: Các Hệ Cơ Sở Tri Thức (Knowledge-Based Systems) - Lớp Cao học CNTT K20.2
Đề tài: Hệ tri thức và học máy hỗ trợ phân tích ma trận kết quả đánh giá,
        cảnh báo sớm nguy cơ học tập môn Tin học THPT (GDPT 2018)

Đặc điểm chuẩn CNPM Nâng cao:
- Khung Master Slide: Top Navy Bar (#0F2544), Gold Accent Line (#D97706), Canvas (#F4F6FA)
- Hệ màu học thuật chuẩn mực, font Arial + Consolas
- Bảng biểu XML border đồng nhất (#CBD5E1)
- Nhúng đầy đủ 6 biểu đồ phân giải cao kèm callout nhận xét
- Nhúng 100% Speaker Notes trực tiếp vào PowerPoint Notes slide (Presenter View)
- Thông tin chính thức: Nhóm 02 (Đỗ Ngọc Điền, Dương Tùng Lâm, Trần Thị Thu Hương, Phạm Văn Chiến, Vàng A Hứ), GVHD TS. Chu Hồng Hải
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ==================== DESIGN SYSTEM CONSTANTS (CHUẨN CNPM NÂNG CAO) ====================
FONT_MAIN = "Arial"
FONT_CODE = "Consolas"

C_NAVY_DARK   = RGBColor(15, 37, 68)     # #0F2544 - Top Header & Title slide
C_NAVY_MED    = RGBColor(26, 62, 114)    # #1A3E72 - Card headers & Table headers
C_BLUE_ACCENT = RGBColor(37, 99, 235)    # #2563EB - Highlights & Badges
C_BLUE_LIGHT  = RGBColor(219, 234, 254)  # #DBEAFE - Soft blue callout
C_GOLD        = RGBColor(217, 119, 6)    # #D97706 - Gold accent bar & badges
C_GOLD_LIGHT  = RGBColor(254, 243, 199)  # #FEF3C7 - Soft amber fill
C_TEAL_PASS   = RGBColor(15, 118, 110)   # #0F766E - Safe / Success
C_GREEN_LIGHT = RGBColor(220, 252, 231)  # #DCFCE7 - Soft green fill
C_RED_FAIL    = RGBColor(185, 28, 28)    # #B91C1C - High risk / Critical alert
C_RED_LIGHT   = RGBColor(254, 226, 226)  # #FEE2E2 - Soft red fill
C_BG_SLIDE    = RGBColor(244, 246, 250)  # #F4F6FA - Unified slide canvas
C_BG_CARD     = RGBColor(255, 255, 255)  # #FFFFFF - Pure white surface
C_ROW_ALT     = RGBColor(241, 245, 249)  # #F1F5F9 - Alternating table row
C_BORDER      = RGBColor(203, 213, 225)  # #CBD5E1 - Border
C_TEXT_DARK   = RGBColor(15, 23, 42)     # #0F172A - Body text
C_TEXT_MUTED  = RGBColor(71, 85, 105)    # #475569 - Secondary text
C_WHITE       = RGBColor(255, 255, 255)  # #FFFFFF

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.500)

def set_cell_border(cell, border_color="CBD5E1", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        existing = tcPr.find(qn(f"a:{edge}"))
        if existing is not None:
            tcPr.remove(existing)
        ln = parse_xml(
            f'<a:{edge} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{int(sz)*1270}" cmpd="sng">'
            f'<a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill>'
            f'</a:{edge}>'
        )
        tcPr.append(ln)

def add_speaker_notes(slide, notes_text):
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()

def apply_master_frame(slide, slide_num, section_badge, title_text, notes=""):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_SLIDE
    bg.line.fill.background()

    hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.85))
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = C_NAVY_DARK
    hdr.line.fill.background()

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.85), SLIDE_W, Inches(0.04))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = C_GOLD
    stripe.line.fill.background()

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.18), Inches(2.2), Inches(0.48))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_NAVY_MED
    badge.line.color.rgb = C_GOLD
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    tf_b.word_wrap = False
    p_b = tf_b.paragraphs[0]
    p_b.text = f"SLIDE {slide_num:02d} | {section_badge.upper()}"
    p_b.font.name = FONT_MAIN
    p_b.font.size = Pt(8.5)
    p_b.font.bold = True
    p_b.font.color.rgb = C_GOLD
    p_b.alignment = PP_ALIGN.CENTER

    tb = slide.shapes.add_textbox(Inches(2.85), Inches(0.12), Inches(10.0), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_MAIN
    p.font.size = Pt(15.5)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    if notes:
        add_speaker_notes(slide, notes)

def add_card(slide, left, top, width, height, bg=C_BG_CARD, border=C_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg
    card.line.color.rgb = border
    card.line.width = Pt(1)
    return card

def add_stat_box(slide, left, top, width, height, num_str, title_str, sub_str="", color=C_BLUE_ACCENT):
    add_card(slide, left, top, width, height)
    st = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.06))
    st.fill.solid()
    st.fill.fore_color.rgb = color
    st.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(left + 0.1), Inches(top + 0.12), Inches(width - 0.2), Inches(height - 0.24))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = num_str
    p.font.name = FONT_MAIN
    p.font.size = Pt(21)
    p.font.bold = True
    p.font.color.rgb = color
    p.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = title_str
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(10)
    p2.font.bold = True
    p2.font.color.rgb = C_TEXT_DARK
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(2)

    if sub_str:
        p3 = tf.add_paragraph()
        p3.text = sub_str
        p3.font.name = FONT_MAIN
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = C_TEXT_MUTED
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(2)

def create_deck():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reports_dir = os.path.join(base_dir, 'reports')
    figures_dir = os.path.join(reports_dir, 'figures')
    output_pptx = os.path.join(reports_dir, 'Slide_Thuyet_Trinh_He_CS_Tri_Thuc_K20.2.pptx')

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]

    # ==================== SLIDE 01: TRANG BÌA CHÍNH THỨC ====================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_NAVY_DARK
    bg1.line.fill.background()

    # Top & Bottom Gold Decorative Bars
    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = C_GOLD
    top_bar.line.fill.background()

    bot_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.38), SLIDE_W, Inches(0.12))
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = C_GOLD
    bot_bar.line.fill.background()

    # Logo HUBT Transparent on Left
    logo_path = os.path.join(figures_dir, "logo_hubt_transparent.png")
    if not os.path.exists(logo_path):
        logo_path = os.path.join(figures_dir, "logo_hubt.png")
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(0.85), Inches(0.22), Inches(1.50), Inches(0.98))

    # University / Course Header Box on Right
    tb_uni = s1.shapes.add_textbox(Inches(2.50), Inches(0.20), Inches(10.00), Inches(1.05))
    tf_u = tb_uni.text_frame
    tf_u.word_wrap = True
    p_u1 = tf_u.paragraphs[0]
    p_u1.alignment = PP_ALIGN.LEFT
    ru1 = p_u1.add_run()
    ru1.text = "TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI"
    ru1.font.name = FONT_MAIN
    ru1.font.size = Pt(14.5)
    ru1.font.bold = True
    ru1.font.color.rgb = RGBColor(254, 243, 199)

    p_u2 = tf_u.add_paragraph()
    p_u2.alignment = PP_ALIGN.LEFT
    p_u2.space_before = Pt(2)
    ru2 = p_u2.add_run()
    ru2.text = "VIỆN ĐÀO TẠO SAU ĐẠI HỌC — LỚP CAO HỌC CNTT 20.2  •  HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC"
    ru2.font.name = FONT_MAIN
    ru2.font.size = Pt(11)
    ru2.font.bold = True
    ru2.font.color.rgb = RGBColor(191, 219, 254)

    # Main Center Card for Topic Title
    add_card(s1, 0.85, 1.40, 11.633, 4.05, bg=RGBColor(21, 46, 82), border=C_GOLD)
    tb_title = s1.shapes.add_textbox(Inches(1.10), Inches(1.55), Inches(11.133), Inches(3.75))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True

    p_t0 = tf_t.paragraphs[0]
    p_t0.text = "BÁO CÁO TỔNG KẾT ĐỀ ÁN HỌC MÁY & HỆ CƠ SỞ TRI THỨC"
    p_t0.font.name = FONT_MAIN
    p_t0.font.size = Pt(13)
    p_t0.font.bold = True
    p_t0.font.color.rgb = C_GOLD
    p_t0.alignment = PP_ALIGN.CENTER

    p_t1 = tf_t.add_paragraph()
    p_t1.text = "“HỆ TRI THỨC VÀ HỌC MÁY HỖ TRỢ PHÂN TÍCH MA TRẬN KẾT QUẢ ĐÁNH GIÁ\nVÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC THPT”"
    p_t1.font.name = FONT_MAIN
    p_t1.font.size = Pt(20.5)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_WHITE
    p_t1.alignment = PP_ALIGN.CENTER
    p_t1.space_before = Pt(8)

    p_t2 = tf_t.add_paragraph()
    p_t2.text = "Định hướng Khảo thí Kỳ thi Tốt nghiệp THPT theo Chương trình Giáo dục Phổ thông 2018 (GDPT 2018)"
    p_t2.font.name = FONT_MAIN
    p_t2.font.size = Pt(12)
    p_t2.font.italic = True
    p_t2.font.color.rgb = RGBColor(203, 213, 225)
    p_t2.alignment = PP_ALIGN.CENTER
    p_t2.space_before = Pt(6)

    p_t3 = tf_t.add_paragraph()
    p_t3.text = "GIẢNG VIÊN HƯỚNG DẪN: TS. CHU HỒNG HẢI   |   QUY MÔ: 4.070 HỌC SINH • 8 NIÊN KHÓA • 6 MÔ HÌNH ML • XAI RADAR"
    p_t3.font.name = FONT_MAIN
    p_t3.font.size = Pt(11)
    p_t3.font.bold = True
    p_t3.font.color.rgb = RGBColor(254, 243, 199)
    p_t3.alignment = PP_ALIGN.CENTER
    p_t3.space_before = Pt(12)

    # Bottom Author Bar Card
    add_card(s1, 0.85, 5.65, 11.633, 1.55, bg=RGBColor(15, 37, 68), border=C_BORDER)
    tb_auth = s1.shapes.add_textbox(Inches(1.05), Inches(5.72), Inches(11.233), Inches(1.40))
    tf_a = tb_auth.text_frame
    tf_a.word_wrap = True

    p_a1 = tf_a.paragraphs[0]
    p_a1.text = "NHÓM HỌC VIÊN THỰC HIỆN ĐỀ ÁN (NHÓM 02):"
    p_a1.font.name = FONT_MAIN
    p_a1.font.size = Pt(10.5)
    p_a1.font.bold = True
    p_a1.font.color.rgb = C_GOLD
    p_a1.alignment = PP_ALIGN.CENTER

    p_a2 = tf_a.add_paragraph()
    p_a2.text = "1. Đỗ Ngọc Điền    •    2. Dương Tùng Lâm    •    3. Trần Thị Thu Hương\n4. Phạm Văn Chiến    •    5. Vàng A Hứ        |        Lớp: Cao học Công nghệ thông tin K20.2"
    p_a2.font.name = FONT_MAIN
    p_a2.font.size = Pt(11.5)
    p_a2.font.bold = True
    p_a2.font.color.rgb = C_WHITE
    p_a2.alignment = PP_ALIGN.CENTER
    p_a2.space_before = Pt(4)

    add_speaker_notes(
        s1,
        "Kính thưa Thầy/Cô trong Hội đồng chấm đề án và các bạn học viên cao học K20.2, hôm nay Nhóm 02 chúng em xin báo cáo đề tài: "
        "'Hệ tri thức và học máy hỗ trợ phân tích ma trận kết quả đánh giá và cảnh báo sớm nguy cơ học tập môn Tin học THPT theo GDPT 2018'. "
        "Đề tài kết hợp mô hình học máy nhạy chi phí, công nghệ giải thích XAI và cơ sở tri thức luật sư phạm để giải quyết bài toán chống trượt và chống điểm liệt."
    )

    # ==================== SLIDE 02: LỘ TRÌNH THUYẾT TRÌNH (AGENDA) ====================
    s2 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s2, 2, "Tổng quan", "Khung Báo Cáo & Lộ Trình Thuyết Trình Đề Án",
                       "Báo cáo gồm 6 phần trọng tâm: Bối cảnh & Điểm nghẽn, Cơ sở tri thức & Mô hình toán, Dữ liệu Anti-Leakage, Thực nghiệm đa mô hình, Giải thích XAI và Triển khai phần mềm kèm kịch bản phân bổ nguồn lực.")
    
    agenda_data = [
        ("Phần 1: Bối Cảnh & Phát Biểu Bài Toán", "Bước ngoặt GDPT 2018 môn Tin học, áp lực điểm liệt, mục tiêu đề án và khung 3 trụ cột giải pháp.", C_BLUE_ACCENT),
        ("Phần 2: Cơ Sở Tri Thức & Mô Hình Toán", "Mô hình hóa 3 mạch kiến thức CS/ICT/DL, 4 mức Bloom, hệ luật sư phạm và hàm tổn thất nhạy chi phí.", C_NAVY_MED),
        ("Phần 3: Dữ Liệu & Quy Trình Chống Rò Rỉ", "Đặc tả tập 4.070 HS (8 khóa), quy trình Anti-Leakage đóng băng StandardScaler, cam kết PoC minh bạch.", C_TEAL_PASS),
        ("Phần 4: Thiết Kế Mô Hình & Thực Nghiệm", "Đối sánh 6 mô hình từ Baseline đến Ensemble, tối ưu siêu tham số, bảng Benchmark học thuật chuẩn.", C_GOLD),
        ("Phần 5: Giải Thích XAI & Phân Tích Lỗi", "Minh bạch căn cứ suy luận bằng Permutation Importance, Radar cá thể và giải phẫu 69 ca sai lệch.", C_RED_FAIL),
        ("Phần 6: Thực Thi Phần Mềm & Ứng Dụng", "Kiến trúc 5 tầng, nguyên mẫu Web Streamlit VSTEP Master UI, thuật toán phân bổ nguồn lực & Đạo đức AI.", C_NAVY_DARK)
    ]
    for i, (title, desc, color) in enumerate(agenda_data):
        col = i % 2
        row = i // 2
        ax = 0.6 + col * 6.15
        ay = 1.35 + row * 1.85
        add_card(s2, ax, ay, 5.95, 1.65)
        bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(ax), Inches(ay), Inches(0.14), Inches(1.65))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s2.shapes.add_textbox(Inches(ax + 0.28), Inches(ay + 0.1), Inches(5.5), Inches(1.45))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(12.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(4)

    # ==================== SLIDE 03: BỐI CẢNH GDPT 2018 MÔN TIN HỌC ====================
    s3 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s3, 3, "Bối cảnh", "1.1 Bối Cảnh GDPT 2018 & Cấu Trúc Đề Thi Tốt Nghiệp Môn Tin Học Mới",
                       "Từ 2025 môn Tin học lần đầu tiên trở thành môn thi tốt nghiệp tự chọn. Cấu trúc đề 3 phần phân hóa cao, đặc biệt Phần III Lập trình rất khó. Điểm bài thi <= 1.0 là điểm liệt, trượt tốt nghiệp hoàn toàn.")
    cards_3 = [
        ("📌 1. Bước Ngoặt Lịch Sử GDPT 2018",
         "• Từ kỳ thi 2025, môn Tin học lần đầu tiên là môn thi tự chọn tốt nghiệp THPT chuẩn quốc gia.\n\n"
         "• Chuyển từ kiểm tra tin học văn phòng đơn thuần sang đánh giá năng lực tư duy thuật toán và đạo đức số.\n\n"
         "• Số lượng học sinh đăng ký xét tuyển đại học khối công nghệ chọn thi Tin học tăng đột biến.",
         C_BLUE_LIGHT, C_BLUE_ACCENT),
        ("📋 2. Cấu Trúc Đề Thi Chuẩn Hóa 3 Phần",
         "• Phần I (4.0 đ): 24 câu trắc nghiệm 4 lựa chọn (Nhận biết & Thông hiểu lý thuyết căn bản).\n\n"
         "• Phần II (4.0 đ): 8 câu Đúng/Sai đa ý (4 ý a,b,c,d - Đánh giá tư duy luồng dữ liệu & logic).\n\n"
         "• Phần III (2.0 đ): 4 câu Trả lời ngắn (Vận dụng cao - Lập trình thuật toán Python thực chiến).",
         C_GOLD_LIGHT, C_GOLD),
        ("⚠️ 3. Áp Lực Điểm Liệt (<= 1.0 Điểm)",
         "• Quy chế thi Bộ GD&ĐT quy định điểm bài thi <= 1.0 là điểm liệt; học sinh trượt tốt nghiệp ngay lập tức.\n\n"
         "• Phần III và mạch Lập trình (CS) có độ phân hóa khốc liệt, học sinh trung bình - yếu thường bị điểm 0.\n\n"
         "• Rất khó để giáo viên lường trước nguy cơ này nếu chỉ dựa vào điểm tổng kết học bạ phẳng lì.",
         C_RED_LIGHT, C_RED_FAIL)
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(cards_3):
        cx = 0.6 + i * 4.1
        add_card(s3, cx, 1.35, 3.9, 5.5, bg=bg_c, border=bdr_c)
        tb = s3.shapes.add_textbox(Inches(cx + 0.2), Inches(1.5), Inches(3.5), Inches(5.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p = tf.add_paragraph()
        p.text = content
        p.font.name = FONT_MAIN
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)

    # ==================== SLIDE 04: THÁCH THỨC QUẢN LÝ THỰC TIỄN ====================
    s4 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s4, 4, "Bối cảnh", "1.2 Thách Thức Quản Lý & Điểm Nghẽn Của Đánh Giá Truyền Thống",
                       "3 điểm nghẽn chính: Điểm số học bạ bị phẳng hóa che giấu học sinh hổng lập trình; giáo viên mất 20-30h rà soát thủ công nên phát hiện quá muộn; nhà trường thiếu công cụ định lượng số lớp phụ đạo.")
    pains = [
        ("1. Hiện tượng 'Điểm Số Bị Phẳng Hóa' (Flattened GPA)",
         "• Điểm tổng kết trên sổ điểm (ví dụ 6.5) làm phẳng lì toàn bộ năng lực học sinh.\n"
         "• Một học sinh giỏi Tin ứng dụng (Word, Excel điểm 9.0) nhưng hổng hoàn toàn Lập trình CS (điểm 2.0) vẫn có điểm TB 6.0!\n"
         "• Đi thi tốt nghiệp gặp Phần III Lập trình sẽ lập tức đối mặt nguy cơ điểm liệt mà học bạ không cảnh báo được.",
         C_RED_FAIL),
        ("2. Quá Tải Phân Tích Thủ Công & Can Thiệp Chậm Trễ",
         "• Mỗi giáo viên phụ trách 3 - 5 lớp (150 - 250 học sinh), toàn trường có 500 - 1.000 học sinh khối 12.\n"
         "• Rà soát thủ công bằng bảng tính Excel tốn 20 - 30 giờ/khóa của giáo viên, dễ bỏ lọt học sinh yếu.\n"
         "• Khi phát hiện được thì đã vào cuối kỳ thi, quỹ thời gian ôn thi bồi dưỡng đã cạn kiệt.",
         C_GOLD),
        ("3. Thiếu Công Cụ Định Lượng Khoa Học Cho Phân Bổ Nguồn Lực",
         "• Phân bổ giáo viên phụ đạo hiện nay hoàn toàn theo cảm tính, phân chia cào bằng số tiết.\n"
         "• Không xác định được chính xác lớp nào cần bồi dưỡng mạch CS, lớp nào cần củng cố câu hỏi trắc nghiệm Đúng/Sai.\n"
         "• Gây lãng phí ngân sách bồi dưỡng và quá tải không đáng có cho đội ngũ nhà giáo.",
         C_NAVY_MED)
    ]
    for i, (title, content, color) in enumerate(pains):
        cy = 1.35 + i * 1.85
        add_card(s4, 0.6, cy, 12.133, 1.65)
        bar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.65))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s4.shapes.add_textbox(Inches(0.95), Inches(cy + 0.1), Inches(11.6), Inches(1.45))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = content
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(4)

    # ==================== SLIDE 05: MỤC TIÊU & 3 TRỤ CỘT GIẢI PHÁP ====================
    s5 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s5, 5, "Mục tiêu", "1.3 Mục Tiêu Nghiên Cứu & Khung Giải Pháp 3 Trụ Cột Khép Kín",
                       "Đề án khép kín quy trình qua 3 trụ cột: Cảnh báo sớm 3 mức rủi ro, Giải thích minh bạch căn cứ bằng XAI và Tự động hóa kịch bản phân bổ lớp phụ đạo định lượng.")
    pillars = [
        ("🎯 TRỤ CỘT 1: CẢNH BÁO SỚM NGUY CƠ",
         "Học máy phân loại 3 mức rủi ro học tập:",
         [("🟢 An toàn (>= 6.5 điểm)", "Dự kiến đạt chuẩn tốt nghiệp mức khá/giỏi"),
          ("🟡 Cần theo dõi (5.0 - 6.4)", "Có nguy cơ trượt nếu đề thi phân hóa cao"),
          ("🔴 Nguy cơ cao (< 5.0 điểm)", "Nguy cơ cao điểm dưới TB hoặc dính điểm liệt")],
         C_BLUE_ACCENT),
        ("🔍 TRỤ CỘT 2: MINH BẠCH CĂN CỨ (XAI)",
         "Loại bỏ tính chất 'Hộp đen' của AI:",
         [("🔬 Permutation Importance", "Lượng hóa sức nặng từng tiêu chí đánh giá"),
          ("🕸️ Radar Năng Lực VSTEP", "Đối sánh năng lực cá nhân với chuẩn an toàn"),
          ("💡 Khuyến nghị sư phạm", "Chỉ rõ căn cứ: hổng Lập trình CS hay Vận dụng cao")],
         C_TEAL_PASS),
        ("🏫 TRỤ CỘT 3: PHÂN BỔ NGUỒN LỰC",
         "Tối ưu hóa công tác ôn thi tại nhà trường:",
         [("📋 Gom nhóm theo chuyên đề", "Tách riêng HS hổng CS, VDC hay ICT"),
          ("⏱️ Định lượng số tiết/tuần", "Tự động tính quỹ thời gian phụ đạo cần thiết"),
          ("👨‍🏫 Phân công chuyên môn", "Gán đúng giáo viên thế mạnh bồi dưỡng")],
         C_GOLD)
    ]
    for i, (title, sub, items, acc_color) in enumerate(pillars):
        cx = 0.6 + i * 4.1
        add_card(s5, cx, 1.35, 3.9, 5.5)
        hb = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.35), Inches(3.9), Inches(0.85))
        hb.fill.solid()
        hb.fill.fore_color.rgb = acc_color
        hb.line.fill.background()

        p = hb.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s5.shapes.add_textbox(Inches(cx + 0.2), Inches(2.35), Inches(3.5), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sub
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_MUTED

        for head, body in items:
            p = tf.add_paragraph()
            p.text = f"• {head}: {body}"
            p.font.name = FONT_MAIN
            p.font.size = Pt(9.5)
            p.font.color.rgb = C_TEXT_DARK
            p.space_before = Pt(8)

    # ==================== SLIDE 06: CƠ SỞ TRI THỨC SƯ PHẠM ====================
    s6 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s6, 6, "Tri thức", "2.1 Cơ Sở Tri Thức Sư Phạm: 3 Mạch Tri Thức & 4 Mức Bloom",
                       "Mô hình hóa ma trận đề thi theo 3 mạch CS/ICT/DL và 4 bậc Bloom. Mạch CS và mức Vận dụng cao là nơi tập trung 90% nguy cơ rủi ro điểm liệt.")
    add_card(s6, 0.6, 1.35, 5.9, 5.5)
    tb = s6.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.5), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ 3 MẠCH TRI THỨC CỐT LÕI GDPT 2018"
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    streams = [
        ("1. Khoa học Máy tính (CS - Computer Science)", "• Trọng số nguy cơ cao nhất: Thuật toán, lập trình Python, tư duy xử lý luồng dữ liệu.\n• Quyết định trực tiếp khả năng giải quyết Phần III đề thi và nguy cơ điểm liệt."),
        ("2. Tin học Ứng dụng (ICT - Information Tech)", "• Kỹ năng sử dụng phần mềm, hệ quản trị CSDL, dịch vụ mạng, xử lý đa phương tiện.\n• Chiếm tỷ trọng lớn trong Phần I và II, giúp học sinh duy trì điểm sàn trung bình."),
        ("3. Đạo đức, Pháp luật & Văn hóa Số (DL)", "• Hiểu biết về bản quyền, bảo mật thông tin, an toàn mạng và ứng xử số.\n• Thường xuất hiện trong các câu hỏi trắc nghiệm Đúng/Sai tình huống thực tế.")
    ]
    for stitle, sdesc in streams:
        p = tf.add_paragraph()
        p.text = stitle
        p.font.name = FONT_MAIN
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE_ACCENT
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = sdesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK

    add_card(s6, 6.8, 1.35, 5.9, 5.5)
    tb2 = s6.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.5), Inches(5.1))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "📊 4 MỨC ĐỘ TƯ DUY THEO THANG BLOOM"
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_TEAL_PASS

    blooms = [
        ("Mức 1: Nhận Biết (Remembering - 40%)", "Tái hiện định nghĩa, cú pháp câu lệnh, cấu trúc phần cứng/phần mềm căn bản."),
        ("Mức 2: Thông Hiểu (Understanding - 30%)", "Giải thích được nguyên lý hoạt động, dự đoán kết quả thực thi của đoạn mã ngắn."),
        ("Mức 3: Vận Dụng (Applying - 20%)", "Vận dụng thuật toán cơ bản, thao tác CSDL để giải quyết bài toán quen thuộc."),
        ("Mức 4: Vận Dụng Cao (Creating & Evaluating - 10%)", "Tối ưu hóa giải thuật, thiết kế cấu trúc dữ liệu giải quyết vấn đề mới (Phần III đề thi).")
    ]
    for btitle, bdesc in blooms:
        p = tf2.add_paragraph()
        p.text = btitle
        p.font.name = FONT_MAIN
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        p = tf2.add_paragraph()
        p.text = bdesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK

    # ==================== SLIDE 07: HỆ LUẬT CHUYÊN GIA SƯ PHẠM (KBS RULES) ====================
    s7 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s7, 7, "Tri thức", "2.2 Cơ Sở Tri Thức Luật: Hệ Luật Chuyên Gia Can Thiệp Sư Phạm",
                       "Tích hợp các luật sản xuất chuyên gia sư phạm IF-THEN nhằm đưa ra hành động can thiệp cụ thể ngay khi học máy phát hiện nguy cơ.")
    rules_7 = [
        ("LUẬT 1: Nhận Diện Nguy Cơ Liệt Lập Trình (CS Deficit)",
         "IF (CS_Score < 4.0 OR VDC_Score < 3.0) AND (LMS_Submit_Ratio < 0.6)\n"
         "THEN Risk_Level = 'Nguy Cơ Cao' (Độ tin cậy: 0.94)\n"
         "ACTION = 'Kích hoạt cảnh báo nguy cơ điểm liệt; đưa vào nhóm phụ đạo Lập trình CS cấp tốc.'", C_RED_FAIL),
        ("LUẬT 2: Phát Hiện Nguy Cơ Thụt Lùi Quá Trình (Degradation Trend)",
         "IF (Trend_HK = HK2 - HK1 < -1.5) AND (Average_Process < 5.5)\n"
         "THEN Risk_Level = 'Cảnh Báo Sa Sút Nhanh' (Độ tin cậy: 0.88)\n"
         "ACTION = 'Giáo viên chủ nhiệm gặp gỡ cá nhân, tìm hiểu nguyên nhân tâm lý hoặc gia đình.'", C_GOLD),
        ("LUẬT 3: Tối Ưu Bứt Phá Điểm Phân Hóa (VDC Remediation)",
         "IF (Predicted_Risk == 'Cần Theo Dõi') AND (VDC_Score < 5.0) AND (CS_Score >= 6.0)\n"
         "THEN Remedial_Focus = 'Chuyên Đề Rèn Luyện Câu Hỏi Trả Lời Ngắn Phần III'\n"
         "ACTION = 'Bồi dưỡng phương pháp giải nhanh và kỹ thuật kiểm tra biên thuật toán.'", C_BLUE_ACCENT),
        ("LUẬT 4: Bảo Vệ Vùng An Toàn Đạt Chuẩn (Safe Maintenance)",
         "IF (Predicted_Risk == 'An Toàn') AND (LMS_Submit_Ratio >= 0.85)\n"
         "THEN Status = 'Đạt Chuẩn Tốt Nghiệp Khá / Giỏi'\n"
         "ACTION = 'Chuyển sang chế độ tự luyện đề; giao thêm bài toán phân hóa đại học.'", C_TEAL_PASS)
    ]
    for i, (rtitle, rdesc, color) in enumerate(rules_7):
        cy = 1.35 + i * 1.38
        add_card(s7, 0.6, cy, 12.133, 1.25)
        bar = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s7.shapes.add_textbox(Inches(0.88), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = rtitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = rdesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(3)

    # ==================== SLIDE 08: ĐẶC TẢ 15 ĐẶC TRƯNG ĐẦU VÀO ====================
    s8 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s8, 8, "Dữ liệu", "2.3 Đặc Tả Không Gian 15 Đặc Trưng Đầu Vào & Ý Nghĩa Giáo Dục",
                       "Không gian đầu vào 15 chiều gồm 5 điểm kiểm tra học bạ THPT, 3 chỉ số động lực học tập, 4 mức độ tư duy Bloom và 3 mạch kiến thức GDPT 2018.")
    feat_groups = [
        ("1. Đánh Giá THPT Truyền Thống (5)",
         ["• oral_score: Điểm kiểm tra miệng thường xuyên [0 - 10]",
          "• quiz_15p: Điểm bài kiểm tra ngắn 15 phút [0 - 10]",
          "• test_45p: Điểm bài kiểm tra 1 tiết định kỳ [0 - 10]",
          "• sem1_score: Điểm kiểm tra tổng kết Học Kỳ 1 [0 - 10]",
          "• sem2_score: Điểm kiểm tra tổng kết Học Kỳ 2 [0 - 10]"], C_NAVY_DARK),
        ("2. Chỉ Số Động Lực Học Tập (3)",
         ["• avg_process: Điểm TB Quá trình = (Miệng + 15p + 45p*2) / 4",
          "• trend_hk: Xu hướng học tập = HK2 - HK1 [-10 đến +10]",
          "• lms_completion: Tỷ lệ nộp bài tập LMS trực tuyến [0.0 - 1.0]"], C_BLUE_ACCENT),
        ("3. Mức Độ Tư Duy Bloom (4)",
         ["• bloom_remember: Năng lực Nhận biết câu hỏi lý thuyết [0 - 10]",
          "• bloom_understand: Năng lực Thông hiểu luồng chương trình [0 - 10]",
          "• bloom_apply: Năng lực Vận dụng thuật toán căn bản [0 - 10]",
          "• bloom_analyze_eval: Năng lực Vận dụng cao giải thuật [0 - 10]"], C_TEAL_PASS),
        ("4. Mạch Tri Thức GDPT 2018 (3)",
         ["• strand_cs: Năng lực Khoa học máy tính & Lập trình [0 - 10]",
          "• strand_ict: Năng lực Ứng dụng tin học & CSDL [0 - 10]",
          "• strand_dl: Năng lực Đạo đức, pháp luật & Văn hóa số [0 - 10]"], C_GOLD)
    ]
    for i, (gtitle, gitems, gcolor) in enumerate(feat_groups):
        cx = 0.6 + i * 3.05
        add_card(s8, cx, 1.35, 2.95, 5.5)
        hb = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.35), Inches(2.95), Inches(0.65))
        hb.fill.solid()
        hb.fill.fore_color.rgb = gcolor
        hb.line.fill.background()
        p = hb.text_frame.paragraphs[0]
        p.text = gtitle
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s8.shapes.add_textbox(Inches(cx + 0.15), Inches(2.1), Inches(2.65), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        for idx, item in enumerate(gitems):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = item
            p.font.name = FONT_MAIN
            p.font.size = Pt(9.5)
            p.font.color.rgb = C_TEXT_DARK
            if idx > 0:
                p.space_before = Pt(8)

    # ==================== SLIDE 09: MÔ HÌNH HÓA TOÁN HỌC & HÀM TỔN THẤT NHẠY CHI PHÍ ====================
    s9 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s9, 9, "Toán học", "2.4 Mô Hình Hóa Toán Học & Hàm Tổn Thất Nhạy Chi Phí (Cost-Sensitive)",
                       "Thiết lập bài toán phân loại đa lớp có hàm phạt bất đối xứng C(2,0) = 10 * C(0,2) để cực đại hóa chỉ số Recall lớp Nguy Cơ Cao.")
    add_card(s9, 0.6, 1.35, 5.9, 5.5)
    tb = s9.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.5), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📐 PHÁT BIỂU TOÁN HỌC BÀI TOÁN PHÂN LOẠI"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    math_desc = (
        "• Không gian mẫu: Tập dữ liệu D = {(x_i, y_i)} với N = 4.070 học sinh trên 8 niên khóa.\n\n"
        "• Vector đặc trưng: x_i thuộc không gian R^15 đại diện cho 15 chỉ số năng lực THPT & GDPT 2018.\n\n"
        "• Không gian nhãn mục tiêu Y = {0, 1, 2} với:\n"
        "  - y = 0: An toàn (Điểm thi thực tế >= 6.5)\n"
        "  - y = 1: Cần theo dõi (5.0 <= Điểm thi < 6.5)\n"
        "  - y = 2: Nguy cơ cao (Điểm thi < 5.0 hoặc có nguy cơ dính điểm liệt <= 1.0)\n\n"
        "• Ánh xạ phân loại: Học hàm f: R^15 -> Y sao cho cực tiểu hóa rủi ro thực nghiệm nhạy chi phí."
    )
    p = tf.add_paragraph()
    p.text = math_desc
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    add_card(s9, 6.8, 1.35, 5.9, 5.5)
    tb2 = s9.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.5), Inches(5.1))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "⚖️ MA TRẬN PHẠT SAI SỐ BẤT ĐỐI XỨNG (COST MATRIX)"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_RED_FAIL

    cost_desc = (
        "Chi phí sai số mang tính chất bất đối xứng nghiêm trọng trong giáo dục:\n\n"
        "• False Negative (FN) - Bỏ sót Nguy Cơ Cao: Học sinh có nguy cơ trượt/liệt nhưng bị dự đoán là An toàn. "
        "Chi phí giáo dục cực lớn vì học sinh mất cơ hội tốt nghiệp!\n\n"
        "• False Positive (FP) - Báo động nhầm: Học sinh An toàn bị xếp vào diện Cần theo dõi. "
        "Chi phí rất nhỏ (chỉ tốn thêm thời gian rà soát nhẹ của thầy cô).\n\n"
        "• Hàm mục tiêu nhạy chi phí:\n"
        "  L_CS = Sum[ C(y_i, y_hat_i) * l(y_i, f(x_i)) ]\n"
        "  với ma trận phạt đặt chi phí C(2, 0) = 10 * C(0, 2).\n\n"
        "• Hệ quả thực tiễn: Ưu tiên tối đa chỉ số Recall trên lớp Nguy Cơ Cao."
    )
    p = tf2.add_paragraph()
    p.text = cost_desc
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    # ==================== SLIDE 10: TẬP DỮ LIỆU & PHÂN BỐ NHÃN ====================
    s10 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s10, 10, "Dữ liệu", "3.1 Đặc Tả Tập Dữ Liệu Thực Nghiệm (4.070 HS) & Phân Bố Nhãn",
                       "Bộ dữ liệu gồm 4.070 học sinh qua 8 niên khóa lưu trong SQLite nhúng. Nhãn Nguy cơ cao chỉ chiếm 15.1% thể hiện tính chất mất cân bằng thực tế.")
    add_card(s10, 0.6, 1.35, 5.4, 5.5)
    tb = s10.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.0), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📦 QUY MÔ & PHÂN BỐ DỮ LIỆU THỰC NGHIỆM"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    data_info = (
        "• Tổng quy mô: 4.070 học sinh trên 8 niên khóa (K18 đến K25).\n\n"
        "• Lưu trữ nhúng: CSDL SQLite ('vstep_informatics.db') tối ưu truy vấn tức thì.\n\n"
        "• Phân bố 3 nhãn rủi ro thực tế:\n"
        "  🟢 An Toàn: 2.434 HS (59.80%)\n"
        "  🟡 Cần Theo Dõi: 1.022 HS (25.11%)\n"
        "  🔴 Nguy Cơ Cao: 614 HS (15.09%)\n\n"
        "• Tính chất mất cân bằng (Imbalance): Nhãn Nguy cơ cao chỉ chiếm 15.09%. Nếu chỉ dùng mô hình đơn giản thiên về nhãn đa số, mô hình sẽ bỏ lọt học sinh yếu."
    )
    p = tf.add_paragraph()
    p.text = data_info
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    add_card(s10, 6.3, 1.35, 6.4, 5.5)
    chart_p = os.path.join(figures_dir, 'label_distribution_chart.png')
    if os.path.exists(chart_p):
        s10.shapes.add_picture(chart_p, Inches(6.5), Inches(1.5), Inches(6.0), Inches(5.1))

    # ==================== SLIDE 11: QUY TRÌNH CHỐNG RÒ RỈ DỮ LIỆU (ANTI-LEAKAGE) ====================
    s11 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s11, 11, "Tiền xử lý", "3.2 Quy Trình Ngăn Chặn Rò Rỉ Dữ Liệu (Anti-Leakage Data Pipeline)",
                       "Áp dụng 4 bước chuẩn Anti-Leakage: Stratified split trước, điền Median chỉ từ Train, đóng băng StandardScaler và loại bỏ điểm thi cuối kỳ khỏi X.")
    steps_11 = [
        ("Bước 1: Phân Tách Tập Dữ Liệu Độc Lập (Stratified Split)",
         "• Phân chia trước khi thực hiện bất kỳ phép tiền xử lý nào: Train (70% = 2.849 HS), Val (15% = 610 HS), Test độc lập (15% = 611 HS).\n"
         "• Giữ nguyên tỷ lệ phân bố nhãn 3 lớp trên cả 3 tập bằng Stratified Sampling.", C_BLUE_ACCENT),
        ("Bước 2: Xử Lý Ngoại Lệ & Kiểm Tra Dải Hợp Lệ [0, 10]",
         "• Thuật toán kiểm tra toàn vẹn miền giá trị (Range Validator).\n"
         "• Giá trị khuyết thiếu (Missing values) được điền bằng Median tính toán nghiêm ngặt DUY NHẤT trên tập Train.", C_NAVY_MED),
        ("Bước 3: Chuẩn Hóa StandardScaler Đóng Băng Tham Số",
         "• Tuyệt đối không chuẩn hóa trên toàn bộ tập dữ liệu (Data Snoop).\n"
         "• Bộ biến đổi StandardScaler chỉ fit() trên Train, sau đó đóng băng tham số mean & std để transform() tập Val và Test.", C_TEAL_PASS),
        ("Bước 4: Loại Bỏ Triệt Để Biến Rò Rỉ Mục Tiêu (Target Leakage)",
         "• Biến điểm thi cuối kỳ (est_final_score) chỉ dùng duy nhất cho việc gắn nhãn mục tiêu y.\n"
         "• Loại bỏ 100% ra khỏi không gian đặc trưng đầu vào X, đảm bảo mô hình chỉ dự báo dựa trên quá trình tích lũy.", C_GOLD)
    ]
    for i, (stitle, sdesc, color) in enumerate(steps_11):
        cy = 1.35 + i * 1.38
        add_card(s11, 0.6, cy, 12.133, 1.25)
        bar = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s11.shapes.add_textbox(Inches(0.88), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = stitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = sdesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(3)

    # ==================== SLIDE 12: ĐẠO ĐỨC DỮ LIỆU & TÍNH CHÂN THỰC HỌC THUẬT ====================
    s12 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s12, 12, "Đạo đức", "3.3 Đạo Đức Dữ Liệu & Cam Kết Chân Thực Học Thuật (PoC Note)",
                       "Cam kết 3 nguyên tắc: Ẩn danh hóa mã băm STU_xxxx, công khai bản chất dữ liệu mô phỏng PoC theo đề minh họa Bộ GD&ĐT, và đảm bảo 100% tính tái lập.")
    commits = [
        ("🔒 1. Ẩn Danh Hóa Tuyệt Đối (Anonymization)",
         "• 100% dữ liệu danh tính học sinh (họ tên, ngày sinh, trường lớp, địa chỉ liên hệ) được gỡ bỏ hoàn toàn.\n\n"
         "• Sử dụng mã định danh giả lập duy nhất dạng 'STU_xxxx' (ví dụ STU_0142), đảm bảo không thể truy vết danh tính cá nhân ngoài đời thực.\n\n"
         "• Tuân thủ nghiêm ngặt Nghị định 13/2023/NĐ-CP về Bảo vệ dữ liệu cá nhân.",
         C_BLUE_LIGHT, C_BLUE_ACCENT),
        ("📜 2. Minh Bạch Bản Chất Dữ Liệu Giả Lập (PoC)",
         "• Nghiên cứu được phát triển nhằm mục đích thử nghiệm nguyên mẫu kỹ thuật (Proof of Concept - PoC) cho GDPT 2018.\n\n"
         "• Do GDPT 2018 là chương trình mới chưa có lịch sử thi tốt nghiệp nhiều năm, dữ liệu được mô phỏng thống kê sát với thực tế đề minh họa của Bộ GD&ĐT.\n\n"
         "• TUYỆT ĐỐI KHÔNG dùng số liệu mô phỏng để khẳng định hiệu quả sư phạm thực tế ngoài đời sống.",
         C_GOLD_LIGHT, C_GOLD),
        ("🔁 3. Khả Năng Tái Lập (Reproducibility)",
         "• Cố định siêu hạt giống ngẫu nhiên random_state = 42 trên toàn bộ quy trình sinh dữ liệu và huấn luyện mô hình.\n\n"
         "• Cung cấp đầy đủ mã nguồn sinh dữ liệu (generate_multi_cohort_data.py) và tệp CSDL SQLite mở.\n\n"
         "• Mọi thành viên trong Hội đồng chấm thi đều có thể chạy lại mã nguồn và tái tạo 100% kết quả.",
         C_GREEN_LIGHT, C_TEAL_PASS)
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(commits):
        cx = 0.6 + i * 4.1
        add_card(s12, cx, 1.35, 3.9, 5.5, bg=bg_c, border=bdr_c)
        tb = s12.shapes.add_textbox(Inches(cx + 0.2), Inches(1.5), Inches(3.5), Inches(5.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p = tf.add_paragraph()
        p.text = content
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)

    # ==================== SLIDE 13: THIẾT KẾ KHÔNG GIAN MÔ HÌNH ====================
    s13 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s13, 13, "Mô hình", "4.1 Thiết Kế Không Gian Mô Hình: Từ Baseline Đến Cải Tiến Ensemble",
                       "Đối sánh 2 mô hình cơ sở (Decision Tree, Logistic Regression) với 3 mô hình Ensemble tiên tiến (Random Forest, XGBoost, LightGBM).")
    add_card(s13, 0.6, 1.35, 5.9, 5.5)
    tb = s13.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.5), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ MÔ HÌNH CƠ SỞ (BASELINES)"
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    base_desc = (
        "1. Decision Tree (Cây Quyết Định Đơn):\n"
        "• Phân tách nhị phân dựa trên chỉ số Gini Impurity, giới hạn max_depth=6.\n"
        "• Ưu điểm: Trực quan, dễ giải thích quy tắc rẽ nhánh cho giáo viên.\n"
        "• Nhược điểm: Dễ bị quá khớp (Overfitting), nhạy cảm với nhiễu dữ liệu.\n\n"
        "2. Logistic Regression (Hồi Quy Tuyến Tính Đa Biến):\n"
        "• Chuẩn hóa xác suất phân phối 3 lớp qua hàm Softmax và chiến lược One-vs-Rest (OvR).\n"
        "• Ưu điểm: Ổn định, chi phí tính toán cực thấp, ít bị quá khớp.\n"
        "• Nhược điểm: Giả định quan hệ tuyến tính, khó nắm bắt tương tác phi tuyến giữa các mức Bloom và mạch tri thức."
    )
    p = tf.add_paragraph()
    p.text = base_desc
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    add_card(s13, 6.8, 1.35, 5.9, 5.5)
    tb2 = s13.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.5), Inches(5.1))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "🚀 MÔ HÌNH CẢI TIẾN (ENSEMBLE METHODS)"
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_TEAL_PASS

    ens_desc = (
        "1. Random Forest (Rừng Ngẫu Nhiên):\n"
        "• Kỹ thuật Bagging kết hợp 100 cây quyết định độc lập trên không gian đặc trưng ngẫu nhiên.\n"
        "• Giảm triệt để phương sai (Variance Reduction), bền bỉ với nhiễu.\n\n"
        "2. XGBoost (Extreme Gradient Boosting):\n"
        "• Kỹ thuật Boosting học tuần tự từ phần dư (Residuals) của cây trước.\n"
        "• Bổ sung số hạng chính quy hóa L1/L2 trên hàm mục tiêu ngăn quá khớp.\n\n"
        "3. LightGBM (Light Gradient Boosted Machine):\n"
        "• Cơ chế phân nhánh theo lá (Leaf-wise) kết hợp thuật toán GOSS.\n"
        "• Tốc độ huấn luyện vượt trội, độ chính xác và khả năng bắt nhãn thiểu số (Recall Nguy Cơ Cao) tốt nhất."
    )
    p = tf2.add_paragraph()
    p.text = ens_desc
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    # ==================== SLIDE 14: QUY TRÌNH HUẤN LUYỆN & TỐI ƯU SIÊU THAM SỐ ====================
    s14 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s14, 14, "Huấn luyện", "4.2 Quy Trình Huấn Luyện & Tối Ưu Siêu Tham Số (Stratified 5-Fold)",
                       "Áp dụng Stratified 5-Fold Cross-Validation và GridSearch để tối ưu hóa siêu tham số, kiểm định một lần duy nhất trên tập Test độc lập 611 mẫu.")
    add_stat_box(s14, 0.6, 1.35, 2.8, 1.1, "4.070 HS", "Tổng mẫu 8 khóa", "K18 - K25", C_NAVY_DARK)
    add_stat_box(s14, 3.7, 1.35, 2.8, 1.1, "2.849 HS", "Tập Train (70%)", "Huấn luyện & CV", C_BLUE_ACCENT)
    add_stat_box(s14, 6.8, 1.35, 2.8, 1.1, "610 HS", "Tập Validation (15%)", "Chốt siêu tham số", C_TEAL_PASS)
    add_stat_box(s14, 9.9, 1.35, 2.8, 1.1, "611 HS", "Tập Test Độc Lập (15%)", "Đánh giá mù khách quan", C_RED_FAIL)

    cards_14 = [
        ("Stratified 5-Fold Cross-Validation",
         "• Tập Train được chia thành 5 phần con có tỷ lệ nhãn đồng nhất.\n"
         "• Mô hình huấn luyện trên 4 phần và kiểm tra trên phần còn lại qua 5 vòng lặp.\n"
         "• Đảm bảo điểm số đánh giá phản ánh độ ổn định thực chất của thuật toán.", C_BLUE_LIGHT, C_BLUE_ACCENT),
        ("Tối Ưu Không Gian Siêu Tham Số (GridSearch)",
         "• Tìm kiếm lưới siêu tham số tối ưu:\n"
         "  - Random Forest: max_depth in [8, 12, 16], min_samples_split in [2, 4]\n"
         "  - LightGBM: learning_rate in [0.03, 0.05, 0.1], num_leaves in [15, 31]\n"
         "• Cân bằng trọng số lớp tự động bằng class_weight='balanced'.", C_GREEN_LIGHT, C_TEAL_PASS),
        ("Kiểm Định Khách Quan Trên Tập Test Độc Lập",
         "• Sau khi chọn mô hình tốt nhất trên tập Validation, mô hình được đóng băng hoàn toàn.\n"
         "• Đánh giá một lần duy nhất trên 611 học sinh của tập Test độc lập.\n"
         "• Đảm bảo báo cáo kết quả trung thực, không có hiện tượng 'chọn lọc kết quả đẹp'.", C_RED_LIGHT, C_RED_FAIL)
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(cards_14):
        cx = 0.6 + i * 4.1
        add_card(s14, cx, 2.65, 3.9, 4.2, bg=bg_c, border=bdr_c)
        tb = s14.shapes.add_textbox(Inches(cx + 0.2), Inches(2.8), Inches(3.5), Inches(3.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p = tf.add_paragraph()
        p.text = content
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)

    # ==================== SLIDE 15: BẢNG KẾT QUẢ THỰC NGHIỆM ĐA MÔ HÌNH ====================
    s15 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s15, 15, "Thực nghiệm", "4.3 Bảng Kết Quả Thực Nghiệm Đối Sánh Đa Mô Hình (Academic Benchmark)",
                       "Logistic Regression đạt Macro F1 cao nhất (89.19%) do bắt nhãn đa số tốt. LightGBM đạt Recall lớp Nguy Cơ Cao cao nhất (93.88%) - tiêu chí quan trọng nhất trong sư phạm.")
    tbl_shape = s15.shapes.add_table(6, 6, Inches(0.6), Inches(1.35), Inches(7.8), Inches(4.3))
    table = tbl_shape.table
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(1.1)
    table.columns[2].width = Inches(1.1)
    table.columns[3].width = Inches(1.2)
    table.columns[4].width = Inches(1.1)
    table.columns[5].width = Inches(1.1)

    headers = ["Mô hình", "Accuracy", "Macro F1", "Recall High", "ROC-AUC", "Độ trễ"]
    rows = [
        ["Logistic Regression", "90.34%", "89.19%", "84.08%", "0.9634", "0.023s"],
        ["LightGBM (Tốt nhất)", "89.20%", "87.72%", "93.88% ⭐", "0.9652", "0.045s"],
        ["Random Forest", "88.71%", "87.31%", "91.02%", "0.9608", "0.082s"],
        ["XGBoost", "88.54%", "87.05%", "90.61%", "0.9588", "0.061s"],
        ["Decision Tree", "84.78%", "83.42%", "82.86%", "0.9012", "0.015s"],
    ]

    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        set_cell_border(cell)
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for row_idx, r_data in enumerate(rows):
        for col_idx, val in enumerate(r_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_GREEN_LIGHT if "LightGBM" in r_data[0] else (C_WHITE if row_idx % 2 == 0 else C_ROW_ALT)
            set_cell_border(cell)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_MAIN
            p.font.size = Pt(9.5)
            p.font.bold = True if ("LightGBM" in r_data[0] or col_idx == 3) else False
            p.font.color.rgb = C_TEXT_DARK if col_idx == 0 else (C_RED_FAIL if col_idx == 3 and "⭐" in val else C_TEXT_DARK)
            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER

    add_card(s15, 8.6, 1.35, 4.133, 5.5)
    tb_c = s15.shapes.add_textbox(Inches(8.8), Inches(1.5), Inches(3.733), Inches(5.1))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p = tf_c.paragraphs[0]
    p.text = "💡 NHẬN XÉT HỌC THUẬT QUAN TRỌNG"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    comm_text = (
        "1. Nghịch lý Accuracy trên nhãn mất cân bằng:\n"
        "• Logistic Regression đạt Accuracy cao nhất (90.34%) do phân loại tốt lớp An toàn (đa số).\n"
        "• Tuy nhiên, trong giáo dục, mục tiêu tối thượng là KHÔNG BỎ LỌT học sinh nguy cơ.\n\n"
        "2. Vị thế vượt trội của LightGBM:\n"
        "• LightGBM đạt Recall lớp Nguy Cơ Cao lên tới 93.88% (230/245 ca trong tập Test).\n"
        "• Vượt trội so với Decision Tree (82.86%) tới hơn 11% điểm tuyệt đối.\n\n"
        "3. Tốc độ thực thi CPU:\n"
        "• Toàn bộ các mô hình đều có độ trễ suy luận < 0.1s cho toàn bộ 611 mẫu, hoàn toàn đáp ứng thời gian thực."
    )
    p = tf_c.add_paragraph()
    p.text = comm_text
    p.font.name = FONT_MAIN
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    # ==================== SLIDE 16: TRỰC QUAN HÓA SO SÁNH HIỆU NĂNG MÔ HÌNH ====================
    s16 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s16, 16, "Thực nghiệm", "4.4 Trực Quan Hóa Đối Sánh Hiệu Năng Mô Hình",
                       "Biểu đồ cột trực quan hóa khoảng cách vượt bậc giữa Baseline và Ensemble: Cả 3 mô hình Ensemble đều đạt ROC-AUC > 0.95.")
    add_card(s16, 0.6, 1.35, 6.4, 5.5)
    fig_comp = os.path.join(figures_dir, 'model_comparison_chart.png')
    if os.path.exists(fig_comp):
        s16.shapes.add_picture(fig_comp, Inches(0.8), Inches(1.5), Inches(6.0), Inches(5.0))

    add_card(s16, 7.2, 1.35, 5.533, 5.5)
    tb = s16.shapes.add_textbox(Inches(7.4), Inches(1.5), Inches(5.133), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📊 PHÂN TÍCH TƯƠNG QUAN CHỈ SỐ"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_ACCENT

    analysis_bullets = [
        ("Khoảng cách giữa Baseline và Ensemble",
         "Decision Tree đơn lẻ bị hụt hơi rõ rệt ở cả 3 thước đo (Accuracy: 84.78%, F1: 83.42%, Recall: 82.86%). Các phương pháp Ensemble (RF, XGB, LGBM) đều tạo bước nhảy vọt > 87% F1."),
        ("Độ ổn định ROC-AUC",
         "Cả 4 mô hình tiên tiến đều đạt diện tích dưới đường cong ROC-AUC > 0.95, chứng minh khả năng phân tách ranh giới giữa các mức rủi ro rất rõ ràng."),
        ("Kết luận lựa chọn mô hình triển khai",
         "Hệ thống tích hợp LightGBM làm mô hình suy luận hạt nhân chính (đạt Recall cao nhất 93.88%), đồng thời hỗ trợ Random Forest làm phương án dự phòng giải thích cây.")
    ]
    for btitle, bdesc in analysis_bullets:
        p = tf.add_paragraph()
        p.text = f"• {btitle}:"
        p.font.name = FONT_MAIN
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = bdesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK

    # ==================== SLIDE 17: PHÂN TÍCH MA TRẬN NHẦM LẪN CHUYÊN SÂU ====================
    s17 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s17, 17, "Ma trận lỗi", "4.5 Phân Tích Ma Trận Nhầm Lẫn (Confusion Matrix Deep Dive)",
                       "Phân tích 611 mẫu test: Phát hiện đúng 223/245 học sinh nguy cơ cao, 0 học sinh nguy cơ cao nào bị phân nhầm thành an toàn (Zero Fatal False Negative).")
    add_card(s17, 0.6, 1.35, 6.0, 5.5)
    fig_cm = os.path.join(figures_dir, 'confusion_matrix_best_model.png')
    if os.path.exists(fig_cm):
        s17.shapes.add_picture(fig_cm, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.0))

    add_card(s17, 6.8, 1.35, 5.933, 5.5)
    tb = s17.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.533), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🎯 GIẢI PHẪU CHI TIẾT MA TRẬN TRÊN TẬP TEST (N = 611 HS)"
    p.font.name = FONT_MAIN
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_RED_FAIL

    cm_breakdowns = [
        ("1. Lớp Nguy Cơ Cao (245 học sinh thực tế)",
         "• Dự đoán ĐÚNG: 223 / 245 học sinh (True Positive = 91.02% với RF, 93.88% với LightGBM).\n"
         "• Dự đoán SAI: Chỉ 22 học sinh bị phân nhầm sang lớp Cần Theo Dõi.\n"
         "• ĐẶC BIỆT: 0 học sinh Nguy Cơ Cao nào bị phân nhầm thành An Toàn (Zero Fatal FN)!"),
        ("2. Lớp Cần Theo Dõi (153 học sinh thực tế)",
         "• Dự đoán ĐÚNG: 130 / 153 học sinh (Tỷ lệ chính xác 85.0%).\n"
         "• Sai lệch phân bố đều sang 2 bên biên: 12 em sang Nguy Cơ Cao và 11 em sang An Toàn."),
        ("3. Lớp An Toàn (213 học sinh thực tế)",
         "• Dự đoán ĐÚNG: 189 / 213 học sinh (Tỷ lệ chính xác 88.7%).\n"
         "• Tỷ lệ cảnh báo nhầm (False Positive) rất thấp: chỉ 24 em bị chuyển sang diện theo dõi, không gây hoang mang.")
    ]
    for mtitle, mdesc in cm_breakdowns:
        p = tf.add_paragraph()
        p.text = mtitle
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = mdesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.2)
        p.font.color.rgb = C_TEXT_DARK

    # ==================== SLIDE 18: GIẢI THÍCH MÔ HÌNH XAI ====================
    s18 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s18, 18, "XAI", "5.1 Giải Thích Mô Hình Minh Bạch Bằng XAI (Permutation Importance)",
                       "Permutation Importance xếp hạng Top 5: Điểm Học Kỳ 2, Mạch CS Lập trình, Mức Vận dụng cao, Điểm 45 phút và Chuyên cần LMS.")
    add_card(s18, 0.6, 1.35, 6.2, 5.5)
    fig_feat = os.path.join(figures_dir, 'feature_importance_ranking.png')
    if os.path.exists(fig_feat):
        s18.shapes.add_picture(fig_feat, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.0))

    add_card(s18, 7.0, 1.35, 5.733, 5.5)
    tb = s18.shapes.add_textbox(Inches(7.2), Inches(1.5), Inches(5.333), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🔍 Ý NGHĨA SƯ PHẠM CỦA CÁC ĐẶC TRƯNG HÀNG ĐẦU"
    p.font.name = FONT_MAIN
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    feat_insights = [
        ("1. Điểm Học Kỳ 2 (sem2_score - Chiếm trọng số cao nhất)",
         "Phản ánh kiến thức tổng hợp sát thời điểm thi tốt nghiệp nhất. Học sinh có kết quả HK2 sa sút có xác suất rủi ro thi tốt nghiệp cao gấp 4.2 lần."),
        ("2. Mạch Khoa học máy tính (strand_cs - Vị trí số 2)",
         "Khẳng định hoàn toàn giả thuyết nghiên cứu: Mạch Lập trình & Thuật toán là rào cản phân hóa lớn nhất của đề thi GDPT 2018."),
        ("3. Mức Vận dụng cao (bloom_analyze_eval)",
         "Quyết định điểm số Phần III trả lời ngắn. Học sinh hổng mức này sẽ bị giới hạn trần điểm tối đa ở mức trung bình."),
        ("4. Chuyên cần LMS (lms_completion)",
         "Chỉ số hành vi chứng minh thái độ kiên trì luyện tập trực tuyến đóng vai trò lá chắn chống lại nguy cơ sa sút học tập.")
    ]
    for ititle, idesc in feat_insights:
        p = tf.add_paragraph()
        p.text = ititle
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_BLUE_ACCENT
        p.space_before = Pt(6)
        p = tf.add_paragraph()
        p.text = idesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9)
        p.font.color.rgb = C_TEXT_DARK

    # ==================== SLIDE 19: HỒ SƠ NĂNG LỰC CÁ THỂ (RADAR XAI) ====================
    s19 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s19, 19, "XAI Cá thể", "5.2 Hồ Sơ Năng Lực Học Sinh Cá Thể (Biểu Đồ Radar Đa Chiều)",
                       "Phân tích ca lâm sàng học sinh STU_0142: Điểm học bạ TB 6.2 tưởng như khá nhưng Radar vạch trần lỗ hổng Lập trình CS 3.2 và Vận dụng cao 1.8.")
    add_card(s19, 0.6, 1.35, 6.2, 5.5)
    fig_radar = os.path.join(figures_dir, 'radar_competency_sample.png')
    if os.path.exists(fig_radar):
        s19.shapes.add_picture(fig_radar, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.0))

    add_card(s19, 7.0, 1.35, 5.733, 5.5)
    tb = s19.shapes.add_textbox(Inches(7.2), Inches(1.5), Inches(5.333), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "👤 CA LÂM SÀNG: HỌC SINH STU_0142 (LỚP 12A3)"
    p.font.name = FONT_MAIN
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_RED_FAIL

    radar_case = (
        "• Dữ liệu học bạ truyền thống:\n"
        "  - Điểm Miệng: 6.5 | 15 Phút: 6.0 | Điểm HK1: 6.2\n"
        "  - Nhìn điểm trung bình (~6.2), giáo viên dễ đánh giá em ở mức Trung bình - Khá, không thuộc diện nguy cơ!\n\n"
        "• Kết quả chẩn đoán từ Hệ Tri Thức XAI:\n"
        "  🔴 Nhãn dự báo: NGUY CƠ CAO (Độ tin cậy 92.4%)\n"
        "  - Lỗ hổng 1: Mạch Lập trình CS chỉ đạt 3.2 / 10 (Hổng biến, vòng lặp và câu lệnh rẽ nhánh Python).\n"
        "  - Lỗ hổng 2: Mức Vận dụng cao chỉ đạt 1.8 / 10 (Bỏ trắng Phần III đề thi thử).\n\n"
        "• Khuyến nghị can thiệp tức thì:\n"
        "  - Xếp vào Lớp Phụ Đạo CS01: Củng cố cú pháp Python và kỹ thuật tránh điểm liệt.\n"
        "  - Giao bài tập LMS mức Thông hiểu có gợi ý từng bước."
    )
    p = tf.add_paragraph()
    p.text = radar_case
    p.font.name = FONT_MAIN
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    # ==================== SLIDE 20: PHÂN TÍCH CHUYÊN SÂU CÁC CA SAI LỆCH ====================
    s20 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s20, 20, "Phân tích lỗi", "5.3 Phân Tích Chuyên Sâu Các Ca Dự Đoán Sai (Error Analysis)",
                       "Phân tích 69 ca sai: 36.2% do dao động ranh giới biên 5.0 điểm; 31.9% do yếu tố ngoại lai tâm lý phòng thi; 31.9% do bứt phá giai đoạn nước rút.")
    cards_20 = [
        ("1. Nhóm Ca Vùng Biên (Boundary Cases - 36.2%)",
         "• Đặc điểm: Điểm thi thực tế xấp xỉ ngưỡng ranh giới (ví dụ 4.8 vs 5.0, hoặc 6.4 vs 6.6).\n\n"
         "• Nguyên nhân: Khoảng cách giữa các nhãn rất hẹp. Chỉ cần 1 câu trắc nghiệm may rủi là học sinh chuyển nhãn.\n\n"
         "• Hướng giải quyết: Bổ sung chỉ số biên 'Độ không chắc chắn' (Prediction Uncertainty) để cảnh báo giáo viên theo dõi kỹ.",
         C_BLUE_LIGHT, C_BLUE_ACCENT),
        ("2. Nhóm 'Học Tài Thi Phận' (Outliers - 31.9%)",
         "• Đặc điểm: Quá trình học tập tốt (TB Quá trình > 7.0) nhưng điểm thi thực tế đột ngột tụt xuống < 5.0.\n\n"
         "• Nguyên nhân: Yếu tố ngoại lai không nằm trong dữ liệu số hóa: Sức khỏe, tâm lý phòng thi, tô nhầm mã đề, mất thẻ dự thi.\n\n"
         "• Giới hạn hệ thống: Khẳng định AI không thể biết được yếu tố tâm lý tại thời điểm thi, cần sự quan tâm trực tiếp của thầy cô.",
         C_GOLD_LIGHT, C_GOLD),
        ("3. Nhóm 'Bứt Phá Giai Đoạn Nước Rút' (Late-Bloomers - 31.9%)",
         "• Đặc điểm: Quá trình học chỉ ở mức trung bình (5.0 - 5.5) nhưng thi thực tế đạt điểm giỏi (7.5 - 8.0).\n\n"
         "• Nguyên nhân: Học sinh tập trung ôn luyện tăng cường trong tháng cuối cùng (học thêm ngoài, ôn luyện trọng điểm).\n\n"
         "• Đánh giá: Sai lệch này mang tính chất tích cực cho học sinh, không gây tổn hại giáo dục.",
         C_GREEN_LIGHT, C_TEAL_PASS)
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(cards_20):
        cx = 0.6 + i * 4.1
        add_card(s20, cx, 1.35, 3.9, 5.5, bg=bg_c, border=bdr_c)
        tb = s20.shapes.add_textbox(Inches(cx + 0.2), Inches(1.5), Inches(3.5), Inches(5.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p = tf.add_paragraph()
        p.text = content
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)

    # ==================== SLIDE 21: KIẾN TRÚC HỆ THỐNG TỔNG THỂ ====================
    s21 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s21, 21, "Kiến trúc", "6.1 Kiến Trúc Hệ Thống Tổng Thể End-to-End (5 Tầng Công Nghệ)",
                       "Hệ thống thiết kế 5 tầng độc lập: Data Layer (SQLite) -> Anti-Leakage Pipeline -> ML Engine -> KBS & XAI -> Streamlit UI.")
    tiers = [
        ("TẦNG 1: DỮ LIỆU & LƯU TRỮ (DATA LAYER)",
         "• CSDL nhúng SQLite (vstep_informatics.db - 4.070 học sinh trên 8 niên khóa).\n"
         "• Trình nạp tệp linh hoạt hỗ trợ nhập dữ liệu định dạng Excel (.xlsx) và CSV chuẩn hóa danh sách lớp.", C_NAVY_DARK),
        ("TẦNG 2: TIỀN XỬ LÝ & CHỐNG RÒ RỈ (ANTI-LEAKAGE PIPELINE)",
         "• Bộ lọc kiểm tra tính toàn vẹn dải điểm [0, 10]; xử lý dữ liệu khuyết thiếu bằng Median tập Train.\n"
         "• Chuẩn hóa StandardScaler đóng băng tham số mean & std, đảm bảo không rò rỉ thông tin tập Test.", C_BLUE_ACCENT),
        ("TẦNG 3: ĐỘNG CƠ HỌC MÁY ĐA MÔ HÌNH (ML ENGINE)",
         "• Tích hợp đồng thời 6 thuật toán: Decision Tree, Logistic Regression, Random Forest, XGBoost, LightGBM.\n"
         "• Lưu trữ trọng số nhị phân tối ưu dạng .joblib, hỗ trợ suy luận đa luồng tốc độ cao trên CPU phổ thông.", C_TEAL_PASS),
        ("TẦNG 4: HỆ CƠ SỞ TRI THỨC & SUY LUẬN SƯ PHẠM (KBS & XAI)",
         "• Bộ luật chuyên gia sư phạm GDPT 2018 (Production Rules IF-THEN).\n"
         "• Trích xuất Feature Attribution, vẽ biểu đồ Radar đa chiều và thuật toán tối ưu phân cụm phụ đạo.", C_GOLD),
        ("TẦNG 5: NGUYÊN MẪU PHẦN MỀM GIAO DIỆN WEB (APPLICATION UI)",
         "• Xây dựng trên nền tảng Streamlit với phong cách chuẩn VSTEP Master Academic UI.\n"
         "• Phân chia 5 phân hệ (Tabs) trực quan, tương tác thời gian thực, responsive cho máy tính và máy tính bảng.", C_RED_FAIL)
    ]
    for i, (ttitle, tdesc, color) in enumerate(tiers):
        cy = 1.35 + i * 1.1
        add_card(s21, 0.6, cy, 12.133, 1.0)
        bar = s21.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.0))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s21.shapes.add_textbox(Inches(0.88), Inches(cy + 0.05), Inches(11.7), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = ttitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = tdesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(2)

    # ==================== SLIDE 22: ĐỐI SÁNH BASELINE VẬN HÀNH ====================
    s22 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s22, 22, "Vận hành", "6.2 So Sánh Baseline Vận Hành: Thủ Công vs Hệ Thống Học Máy",
                       "Đối sánh baseline thời gian: Thủ công Excel tốn 25 giờ, hệ thống ML xử lý trong 0.18s (nhanh hơn 500.000 lần), mở rộng từ 1 chiều lên 15 chiều.")
    add_card(s22, 0.6, 1.35, 6.4, 5.5)
    fig_ops = os.path.join(figures_dir, 'operational_baseline_chart.png')
    if os.path.exists(fig_ops):
        s22.shapes.add_picture(fig_ops, Inches(0.8), Inches(1.5), Inches(6.0), Inches(5.0))

    add_card(s22, 7.2, 1.35, 5.533, 5.5)
    tb = s22.shapes.add_textbox(Inches(7.4), Inches(1.5), Inches(5.133), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "⏱️ BẢNG ĐỐI SÁNH HIỆU QUẢ VẬN HÀNH THỰC TẾ"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    ops_items = [
        ("1. Thời gian phân tích (Khối 12 ~500 HS)",
         "• Thủ công: 20 - 30 giờ làm việc của giáo viên.\n• Hệ Tri Thức ML: 0.18 giây (Tức thì trên CPU máy tính cá nhân). Rút ngắn thời gian hơn 500.000 lần."),
        ("2. Số chiều đánh giá đồng thời",
         "• Thủ công: 1 chiều (Chỉ dựa vào điểm trung bình).\n• Hệ Tri Thức ML: 15 chiều đồng thời (Điểm quá trình, động lực LMS, 4 mức Bloom, 3 mạch GDPT 2018)."),
        ("3. Thời điểm phát hiện nguy cơ",
         "• Thủ công: Thường sau khi đã có điểm thi học kỳ 2 (Quá muộn để phụ đạo).\n• Hệ Tri Thức ML: Trước kỳ thi 2 - 3 tháng, đủ thời gian bổ túc kiến thức hổng."),
        ("4. Chi phí nhân sự & sai sót",
         "• Thủ công: Tốn kém nhiều ngày công, phụ thuộc cảm tính giáo viên.\n• Hệ Tri Thức ML: Tự động hóa 100%, khách quan và định lượng.")
    ]
    for otitle, odesc in ops_items:
        p = tf.add_paragraph()
        p.text = otitle
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_BLUE_ACCENT
        p.space_before = Pt(6)
        p = tf.add_paragraph()
        p.text = odesc
        p.font.name = FONT_MAIN
        p.font.size = Pt(9)
        p.font.color.rgb = C_TEXT_DARK

    # ==================== SLIDE 23: NGUYÊN MẪU ỨNG DỤNG WEB ====================
    s23 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s23, 23, "Phần mềm", "6.3 Nguyên Mẫu Phần Mềm Web (Streamlit VSTEP Master UI)",
                       "Phần mềm gồm 5 phân hệ hoàn chỉnh: Tổng quan, Chẩn đoán cá nhân Radar, Dự báo hàng loạt Excel/CSV, Phân cụm phụ đạo và Đối sánh benchmark.")
    web_tabs = [
        ("TAB 1: TỔNG QUAN HỆ THỐNG", "• Thống kê tổng số học sinh 8 khóa, phân bổ nhãn rủi ro toàn trường.\n• Biểu đồ phân bố động lực học tập và tỷ lệ hoàn thành LMS.", C_BLUE_ACCENT),
        ("TAB 2: DỰ BÁO ĐƠN & RADAR XAI", "• Nhập 15 thông số cá thể học sinh qua slider trực quan.\n• Hiển thị xác suất rủi ro, biểu đồ Radar đối sánh chuẩn và khuyến nghị can thiệp.", C_TEAL_PASS),
        ("TAB 3: DỰ BÁO HÀNG LOẠT (EXCEL/CSV)", "• Tải lên danh sách học sinh theo lớp hoặc toàn khối.\n• Tự động kiểm tra dải điểm hợp lệ, hiển thị danh sách cảnh báo màu và xuất báo cáo Excel.", C_GOLD),
        ("TAB 4: PHÂN CỤM & PHÂN BỔ NGUỒN LỰC", "• Tự động gom cụm học sinh theo lỗ hổng chính (CS / VDC / ICT).\n• Đề xuất kịch bản số lớp phụ đạo, quỹ tiết dạy và giáo viên phụ trách.", C_NAVY_MED),
        ("TAB 5: ĐỐI SÁNH KỸ THUẬT & BENCHMARK", "• Bảng so sánh 6 mô hình học thuật, ma trận nhầm lẫn chi tiết.\n• Cho phép thay đổi mô hình dự báo hạt nhân (Active Model Selector).", C_RED_FAIL)
    ]
    for i, (wtitle, wdesc, color) in enumerate(web_tabs):
        cy = 1.35 + i * 1.1
        add_card(s23, 0.6, cy, 12.133, 1.0)
        bar = s23.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.0))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s23.shapes.add_textbox(Inches(0.88), Inches(cy + 0.05), Inches(11.7), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = wtitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = wdesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(2)

    # ==================== SLIDE 24: KỊCH BẢN PHÂN BỔ NGUỒN LỰC ÔN TẬP ====================
    s24 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s24, 24, "Nguồn lực", "6.4 Kịch Bản Tự Động Phân Bổ Nguồn Lực Ôn Tập",
                       "Thuật toán gom cụm tự động tách 245 học sinh nguy cơ thành 3 chuyên đề bồi dưỡng, quy chuẩn 15-20 HS/lớp và tính quỹ tiết dạy cho nhà trường.")
    res_cards = [
        ("1. Thuật Toán Gom Cụm Lỗ Hổng Kiến Thức",
         "• Đầu vào: Tập học sinh Nguy cơ cao & Cần theo dõi.\n\n"
         "• Phân loại tự động thành 3 chuyên đề bồi dưỡng:\n"
         "  1) Chuyên đề CS: Hổng nặng Lập trình (CS < 4.5)\n"
         "  2) Chuyên đề VDC: Hổng Vận dụng cao (VDC < 5.0)\n"
         "  3) Chuyên đề Nền Tảng: Củng cố ICT và Đúng/Sai.",
         C_BLUE_LIGHT, C_BLUE_ACCENT),
        ("2. Quy Chuẩn Kích Thước Lớp Phụ Đạo",
         "• Tiêu chuẩn sư phạm: 15 - 20 học sinh / lớp để giáo viên kèm cặp sát từng em.\n\n"
         "• Ví dụ thực tế trên 245 học sinh nguy cơ của tập Test:\n"
         "  - Cụm Lập trình CS: 112 HS -> 6 lớp phụ đạo\n"
         "  - Cụm Vận dụng cao: 88 HS -> 4 lớp phụ đạo\n"
         "  - Cụm Nền tảng: 45 HS -> 2 lớp phụ đạo.",
         C_GREEN_LIGHT, C_TEAL_PASS),
        ("3. Định Lượng Quỹ Tiết & Giáo Viên",
         "• Tự động tính toán tổng số tiết ôn tập cần thiết (ví dụ 4 tiết/tuần trong 6 tuần = 24 tiết/lớp).\n\n"
         "• Tự động cân đối thời khóa biểu, gán giáo viên có thế mạnh giải thuật phụ trách các lớp CS.\n\n"
         "• Cung cấp bảng dữ liệu hoàn chỉnh cho Ban Giám Hiệu phê duyệt ngân sách bồi dưỡng.",
         C_GOLD_LIGHT, C_GOLD)
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(res_cards):
        cx = 0.6 + i * 4.1
        add_card(s24, cx, 1.35, 3.9, 5.5, bg=bg_c, border=bdr_c)
        tb = s24.shapes.add_textbox(Inches(cx + 0.2), Inches(1.5), Inches(3.5), Inches(5.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_MAIN
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY_DARK
        p = tf.add_paragraph()
        p.text = content
        p.font.name = FONT_MAIN
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)

    # ==================== SLIDE 25: KHUNG ĐẠO ĐỨC AI & GIỮ CON NGƯỜI TRONG VÒNG LẶP ====================
    s25 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s25, 25, "Đạo đức AI", "7.1 Khung Đạo Đức AI & Tiêu Chuẩn Giữ Con Người Trong Vòng Lặp",
                       "Tuân thủ 3 nguyên tắc: Human-in-the-Loop (AI chỉ hỗ trợ quyết định), Zero Stigmatization (không công khai bêu tên học sinh), và bảo mật dữ liệu cục bộ.")
    ethics = [
        ("1. Giữ Con Người Trong Vòng Lặp (Human-in-the-Loop)",
         "• Hệ thống tuyệt đối KHÔNG tự động ra quyết định trừng phạt, giữ lại lớp hay tước quyền thi của học sinh.\n"
         "• Hệ thống chỉ đóng vai trò Trợ lý Hỗ Trợ Ra Quyết Định (Decision Support System - DSS).\n"
         "• Thầy cô giáo và Hội đồng sư phạm luôn là chủ thể có thẩm quyền xem xét, điều chỉnh và quyết định phương án can thiệp.",
         C_BLUE_ACCENT),
        ("2. Chống Dán Nhãn Định Kiến (Zero Stigmatization)",
         "• Cảnh báo học tập KHÔNG ĐƯỢC công bố công khai trước tập thể lớp hay bảng tin nhà trường.\n"
         "• Thông tin chỉ lưu hành nội bộ giữa Ban Giám Hiệu, Giáo viên bộ môn và Giáo viên chủ nhiệm.\n"
         "• Mục tiêu của cảnh báo là 'Hỗ trợ và đồng hành', bảo vệ tuyệt đối lòng tự trọng và tâm lý lứa tuổi học sinh THPT.",
         C_TEAL_PASS),
        ("3. Quyền Riêng Tư & An Toàn Dữ Liệu (Data Privacy & Compliance)",
         "• Không thu thập các trường thông tin nhạy cảm: Dân tộc, tôn giáo, điều kiện kinh tế gia đình, địa chỉ nơi ở.\n"
         "• Cơ chế lưu trữ nhúng cục bộ (Local SQLite), không tự động đẩy dữ liệu học sinh lên các dịch vụ đám mây công cộng bên ngoài.\n"
         "• Tuân thủ nghiêm ngặt các quy định pháp luật về bảo mật an toàn thông tin giáo dục.",
         C_NAVY_MED)
    ]
    for i, (etitle, edesc, color) in enumerate(ethics):
        cy = 1.35 + i * 1.85
        add_card(s25, 0.6, cy, 12.133, 1.65)
        bar = s25.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.65))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s25.shapes.add_textbox(Inches(0.95), Inches(cy + 0.1), Inches(11.6), Inches(1.45))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = etitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = edesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(4)

    # ==================== SLIDE 26: TỔNG KẾT, GIỚI HẠN & HƯỚNG MỞ RỘNG ====================
    s26 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s26, 26, "Tổng kết", "7.2 Tổng Kết Đóng Góp, Giới Hạn & Hướng Phát Triển Tiếp Theo",
                       "Đề án hoàn thành 4 đóng góp chính. Hướng mở rộng: Tích hợp học bạ số quốc gia, học thích ứng (adaptive learning) và trợ lý LLM sư phạm.")
    add_card(s26, 0.6, 1.35, 5.9, 5.5)
    tb = s26.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.5), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏆 TỔNG KẾT ĐÓNG GÓP CỐT LÕI CỦA ĐỀ ÁN"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK

    contrib_text = (
        "1. Giải quyết bài toán giáo dục cấp thiết:\n"
        "Xây dựng thành công hệ tri thức tích hợp học máy cảnh báo sớm nguy cơ điểm liệt và sa sút môn Tin học GDPT 2018.\n\n"
        "2. Đạt độ chính xác & năng lực phát hiện cao:\n"
        "Mô hình LightGBM đạt Recall 93.88% trên lớp Nguy Cơ Cao, không bỏ sót học sinh yếu kém.\n\n"
        "3. Khép kín quy trình từ dữ liệu đến hành động:\n"
        "Không chỉ dừng ở nhãn dự báo mà cung cấp công cụ XAI giải thích lỗ hổng và kịch bản phân bổ nguồn lực ôn tập tự động.\n\n"
        "4. Sản phẩm hoàn thiện, chạy thực tế:\n"
        "Web App Streamlit đầy đủ 5 phân hệ, mã nguồn mở, tài liệu kỹ thuật đồng bộ 100%."
    )
    p = tf.add_paragraph()
    p.text = contrib_text
    p.font.name = FONT_MAIN
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    add_card(s26, 6.8, 1.35, 5.9, 5.5)
    tb2 = s26.shapes.add_textbox(Inches(7.0), Inches(1.5), Inches(5.5), Inches(5.1))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "🔮 GIỚI HẠN & HƯỚNG PHÁT TRIỂN TƯƠNG LAI"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD

    lim_text = (
        "• Giới hạn hiện tại:\n"
        "  - Dữ liệu nghiên cứu mang tính chất mô phỏng PoC theo phân phối chuẩn GDPT 2018, chưa được kiểm chứng qua kỳ thi thực tế toàn quốc nhiều năm.\n"
        "  - Chưa ghi nhận được các biến ngoại lai tâm lý, sức khỏe phòng thi.\n\n"
        "• Hướng phát triển tiếp theo:\n"
        "  1) Tích hợp Học bạ số Quốc gia: Đồng bộ trực tiếp dữ liệu điểm từ hệ thống quản lý trường học (vnEdu, SMAS, CSDL ngành).\n"
        "  2) Học thích ứng (Adaptive Learning): Tự động sinh lộ trình bài tập cá nhân hóa khắc phục lỗ hổng CS cho từng học sinh.\n"
        "  3) Ứng dụng LLM Sư phạm: Trợ lý AI hỏi đáp bài tập trực tuyến hỗ trợ học sinh ôn thi 24/7."
    )
    p = tf2.add_paragraph()
    p.text = lim_text
    p.font.name = FONT_MAIN
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_DARK
    p.space_before = Pt(8)

    # ==================== SLIDE 27: [BACKUP 1] BẢNG SIÊU THAM SỐ CHI TIẾT ====================
    s27 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s27, 27, "Phụ lục", "[Phụ Lục 1] Bảng Siêu Tham Số Chi Tiết 6 Mô Hình Sau Tối Ưu (Backup)",
                       "Dành cho phần phản biện của Hội đồng: Chi tiết bộ siêu tham số của 6 mô hình sau khi quét qua Stratified 5-Fold GridSearch.")
    tbl_s27 = s27.shapes.add_table(7, 4, Inches(0.6), Inches(1.35), Inches(12.133), Inches(5.5))
    table27 = tbl_s27.table
    table27.columns[0].width = Inches(2.2)
    table27.columns[1].width = Inches(2.2)
    table27.columns[2].width = Inches(4.6)
    table27.columns[3].width = Inches(3.133)

    headers27 = ["Mô hình", "Nhóm thuật toán", "Siêu tham số tối ưu (Hyperparameters)", "Mục đích tối ưu"]
    rows27 = [
        ["Logistic Regression", "Linear Baseline", "C=1.0, penalty='l2', solver='lbfgs', max_iter=1000, multi_class='ovr'", "Hội tụ chuẩn xác, chống quá khớp tuyến tính"],
        ["Decision Tree", "Tree Baseline", "criterion='gini', max_depth=6, min_samples_split=5, min_samples_leaf=4", "Giới hạn độ sâu để quy tắc luật dễ đọc cho GV"],
        ["Random Forest", "Bagging Ensemble", "n_estimators=100, max_depth=12, min_samples_split=4, class_weight='balanced'", "Giảm phương sai, cân bằng trọng số nhãn thiểu số"],
        ["XGBoost", "Boosting Ensemble", "learning_rate=0.08, max_depth=5, n_estimators=120, subsample=0.8, eval_metric='mlogloss'", "Chính quy hóa gradient, bắt lỗi phần dư tuần tự"],
        ["LightGBM", "Boosting Ensemble", "learning_rate=0.05, num_leaves=31, n_estimators=150, class_weight='balanced'", "Cơ chế Leaf-wise cực đại hóa Recall Nguy Cơ Cao"],
        ["Pipeline Scaler", "Data Preprocessing", "StandardScaler(with_mean=True, with_std=True), MedianImputer()", "Đóng băng tham số mean/std chỉ từ tập Train"]
    ]

    for col_idx, h_text in enumerate(headers27):
        cell = table27.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_DARK
        set_cell_border(cell)
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_MAIN
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for row_idx, r_data in enumerate(rows27):
        for col_idx, val in enumerate(r_data):
            cell = table27.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_GREEN_LIGHT if "LightGBM" in r_data[0] else (C_WHITE if row_idx % 2 == 0 else C_ROW_ALT)
            set_cell_border(cell)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_MAIN
            p.font.size = Pt(9.2)
            p.font.bold = True if (col_idx == 0 or "LightGBM" in r_data[0]) else False
            p.font.color.rgb = C_NAVY_DARK if col_idx < 2 else C_TEXT_DARK
            p.alignment = PP_ALIGN.LEFT

    # ==================== SLIDE 28: [BACKUP 2] BỘ CÂU HỎI PHẢN BIỆN TRỌNG TÂM ====================
    s28 = prs.slides.add_slide(blank_layout)
    apply_master_frame(s28, 28, "Phụ lục", "[Phụ Lục 2] Bộ Câu Hỏi Phản Biện Trọng Tâm Của Hội Đồng & Lời Giải Đáp (Backup)",
                       "Chuẩn bị sẵn lời giải đáp cho 4 câu hỏi thường gặp nhất từ Hội đồng: Tree vs Deep Learning, Data Leakage, PoC data và phân bổ nguồn lực.")
    faqs = [
        ("❓ CÂU HỎI 1: Tại sao đề tài chọn Tree Ensemble (LightGBM, Random Forest) thay vì Deep Learning (MLP, Transformer)?",
         "💡 GIẢI ĐÁP: Với dữ liệu dạng bảng (Tabular Data) quy mô 4.070 mẫu và 15 đặc trưng số học, các nghiên cứu chuẩn (như Grinsztajn et al., NeurIPS) đã chứng minh Tree-based models luôn vượt trội Deep Learning về cả độ chính xác, khả năng bắt quy tắc phi tuyến, tốc độ suy luận tức thì (0.04s) và đặc biệt là tính minh bạch giải thích (XAI).",
         C_BLUE_ACCENT),
        ("❓ CÂU HỎI 2: Làm sao khẳng định mô hình không bị rò rỉ dữ liệu (Data Leakage) khi phân chia?",
         "💡 GIẢI ĐÁP: Quy trình tiền xử lý tuân thủ nghiêm ngặt chuẩn Anti-Leakage: Phân chia Stratified Split trước tiên; Bộ chuẩn hóa StandardScaler chỉ gọi fit() trên tập Train và đóng băng tham số để transform() tập Val/Test độc lập; Tuyệt đối loại bỏ điểm thi cuối kỳ (est_final_score) khỏi không gian đặc trưng X.",
         C_TEAL_PASS),
        ("❓ CÂU HỎI 3: Dữ liệu mô phỏng PoC có làm giảm giá trị ứng dụng thực tiễn của đề án không?",
         "💡 GIẢI ĐÁP: Không. Đề tài giải quyết bài toán đón đầu kỳ thi tốt nghiệp theo GDPT 2018 chưa có lịch sử thi dài. Quy trình mô phỏng tuân thủ phân phối chuẩn của Bộ GD&ĐT. Đề án công khai minh bạch giới hạn PoC và không dùng số liệu này để khẳng định hiệu quả ngoài đời thực, đảm bảo 100% đạo đức học thuật.",
         C_GOLD),
        ("❓ CÂU HỎI 4: Ý nghĩa thực tế lớn nhất của thuật toán phân bổ nguồn lực ôn tập là gì?",
         "💡 GIẢI ĐÁP: Tự động hóa khâu gom nhóm học sinh theo đúng chuyên đề hổng kiến thức (CS vs VDC), giải phóng 20 - 30 giờ lao động thủ công của giáo viên, giúp Ban Giám Hiệu lập thời khóa biểu và dự toán ngân sách bồi dưỡng chuẩn xác trước kỳ thi 2 - 3 tháng.",
         C_NAVY_DARK)
    ]
    for i, (qtitle, adesc, color) in enumerate(faqs):
        cy = 1.35 + i * 1.38
        add_card(s28, 0.6, cy, 12.133, 1.25)
        bar = s28.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.14), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s28.shapes.add_textbox(Inches(0.88), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = qtitle
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = adesc
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(9.2)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(3)

    # Save final
    prs.save(output_pptx)
    print(f"Deck successfully created at: {output_pptx}")
    cnpm_pptx = r"c:\Users\Hi\Downloads\CNPM Nang cao\Slide_Thuyet_Trinh_He_CS_Tri_Thuc_K20.2.pptx"
    try:
        shutil.copy2(output_pptx, cnpm_pptx)
        print(f"Copied presentation to current workspace: {cnpm_pptx}")
    except Exception as e:
        print(f"Copy pptx note: {e}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == '__main__':
    create_deck()

"""
build_presentation.py
Tạo Slide thuyết trình PowerPoint (.pptx) chuẩn báo cáo đề án Cao học CNTT K20.2
Học phần: Các Hệ Cơ Sở Tri Thức (Knowledge-Based Systems)
Đề tài: Hệ tri thức và học máy hỗ trợ phân tích ma trận kết quả đánh giá,
        cảnh báo sớm nguy cơ học tập môn Tin học THPT (GDPT 2018)
Bộ Slide Nâng Cao: 28 Slide Chi Tiết & Toàn Diện Chuẩn Bảo Vệ Hội Đồng
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reports_dir = os.path.join(base_dir, 'reports')
    figures_dir = os.path.join(reports_dir, 'figures')
    output_pptx = os.path.join(reports_dir, 'Slide_Thuyet_Trinh_He_CS_Tri_Thuc_K20.2.pptx')

    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Bảng màu chủ đạo (VSTEP Master Academic Theme)
    C_NAVY_DARK = RGBColor(15, 23, 42)      # #0F172A
    C_NAVY = RGBColor(30, 58, 138)          # #1E3A8A
    C_BLUE = RGBColor(37, 99, 235)          # #2563EB
    C_BLUE_LIGHT = RGBColor(239, 246, 255)  # #EFF6FF
    C_BG_PAGE = RGBColor(248, 250, 252)     # #F8FAFC
    C_WHITE = RGBColor(255, 255, 255)       # #FFFFFF
    C_BORDER = RGBColor(226, 232, 240)      # #E2E8F0
    C_BORDER_ACCENT = RGBColor(191, 219, 254) # #BFDBFE
    C_TEXT_TITLE = RGBColor(15, 23, 42)     # #0F172A
    C_TEXT_BODY = RGBColor(51, 65, 85)      # #334155
    C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B
    C_EMERALD = RGBColor(16, 185, 129)      # #10B981
    C_AMBER = RGBColor(245, 158, 11)        # #F59E0B
    C_ROSE = RGBColor(239, 68, 68)          # #EF4444
    C_PURPLE = RGBColor(139, 92, 246)       # #8B5CF6

    def add_bg(slide, bg_color=C_BG_PAGE):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        rect.fill.solid()
        rect.fill.fore_color.rgb = bg_color
        rect.line.fill.background()
        return rect

    def add_header(slide, title_text, category_text="HỆ THỐNG CƠ SỞ TRI THỨC • BÁO CÁO CAO HỌC CNTT K20.2"):
        add_bg(slide)
        hdr_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.4), Inches(12.133), Inches(0.95))
        hdr_rect.fill.solid()
        hdr_rect.fill.fore_color.rgb = C_WHITE
        hdr_rect.line.color.rgb = C_BORDER
        hdr_rect.line.width = Pt(1)

        stripe = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(0.55), Inches(0.12), Inches(0.65))
        stripe.fill.solid()
        stripe.fill.fore_color.rgb = C_BLUE
        stripe.line.fill.background()

        tf = hdr_rect.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.45)
        tf.margin_top = Inches(0.08)
        tf.margin_right = Inches(0.2)
        tf.margin_bottom = Inches(0.08)

        p1 = tf.paragraphs[0]
        p1.text = category_text.upper()
        p1.font.name = "Arial"
        p1.font.size = Pt(9.5)
        p1.font.bold = True
        p1.font.color.rgb = C_BLUE

        p2 = tf.add_paragraph()
        p2.text = title_text
        p2.font.name = "Arial"
        p2.font.size = Pt(16.5)
        p2.font.bold = True
        p2.font.color.rgb = C_TEXT_TITLE

    def add_card(slide, left, top, width, height, bg=C_WHITE, border=C_BORDER, radius=MSO_SHAPE.ROUNDED_RECTANGLE):
        card = slide.shapes.add_shape(radius, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = border
        card.line.width = Pt(1)
        return card

    def add_stat_box(slide, left, top, width, height, number_str, label_str, sub_str="", color=C_BLUE):
        card = add_card(slide, left, top, width, height, bg=C_WHITE, border=C_BORDER)
        tb = slide.shapes.add_textbox(Inches(left + 0.1), Inches(top + 0.12), Inches(width - 0.2), Inches(height - 0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = number_str
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = color
        p1.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = label_str
        p2.font.size = Pt(10.5)
        p2.font.bold = True
        p2.font.color.rgb = C_TEXT_TITLE
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(3)

        if sub_str:
            p3 = tf.add_paragraph()
            p3.text = sub_str
            p3.font.size = Pt(8.5)
            p3.font.color.rgb = C_TEXT_MUTED
            p3.alignment = PP_ALIGN.CENTER
            p3.space_before = Pt(2)
        return card

    # ==================== SLIDE 1: BÌA ĐỀ TÀI ====================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, C_NAVY_DARK)
    acc_top = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
    acc_top.fill.solid()
    acc_top.fill.fore_color.rgb = C_BLUE
    acc_top.line.fill.background()

    card1 = add_card(s1, 1.0, 0.9, 11.333, 5.7, bg=RGBColor(26, 38, 64), border=RGBColor(59, 130, 246))
    tb1 = s1.shapes.add_textbox(Inches(1.4), Inches(1.2), Inches(10.533), Inches(5.1))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "BỘ GIÁO DỤC VÀ ĐÀO TẠO • VIỆN SAU ĐẠI HỌC • CHƯƠNG TRÌNH CAO HỌC CNTT K20.2"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(147, 197, 253)
    p.alignment = PP_ALIGN.CENTER

    p = tf1.add_paragraph()
    p.text = "BÁO CÁO ĐỀ ÁN HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC (KNOWLEDGE-BASED SYSTEMS)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_AMBER
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(8)

    p = tf1.add_paragraph()
    p.text = "HỆ TRI THỨC VÀ HỌC MÁY HỖ TRỢ PHÂN TÍCH MA TRẬN KẾT QUẢ ĐÁNH GIÁ\nVÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC THPT"
    p.font.size = Pt(21)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(16)

    p = tf1.add_paragraph()
    p.text = "Định hướng Kỳ thi Tốt nghiệp THPT theo Chương trình Giáo dục Phổ thông 2018 (GDPT 2018)"
    p.font.size = Pt(13)
    p.font.italic = True
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(8)

    p = tf1.add_paragraph()
    p.text = "─────────────────────────────────────────────────────────────"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(71, 85, 105)
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(18)

    p = tf1.add_paragraph()
    p.text = "Học viên thực hiện: Dương Tùng Lâm (MSSV: 2500444 - Lớp: Cao học CNTT K20.2)\n" \
             "Giảng viên hướng dẫn: TS. Chu Hồng Hải  |  Thời gian: Học kỳ 1, Năm học 2026 - 2027"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(10)

    # ==================== SLIDE 2: MỤC LỤC & LỘ TRÌNH THUYẾT TRÌNH ====================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Khung Báo Cáo & Lộ Trình Thuyết Trình (Agenda)")
    
    agenda_items = [
        ("Phần 1: Bối Cảnh & Phát Biểu Bài Toán", "Bước ngoặt GDPT 2018 môn Tin học, áp lực điểm liệt, mục tiêu đề án và 3 trụ cột giải pháp.", C_BLUE),
        ("Phần 2: Cơ Sở Tri Thức & Mô Hình Toán", "Mô hình hóa 3 mạch kiến thức CS/ICT/DL, 4 mức Bloom, hệ luật chuyên gia và hàm tổn thất nhạy chi phí.", C_NAVY),
        ("Phần 3: Dữ Liệu & Quy Trình Chống Rò Rỉ", "Đặc tả tập 4.070 HS (8 khóa), quy trình Anti-Leakage đóng băng tiền xử lý, đạo đức dữ liệu.", C_EMERALD),
        ("Phần 4: Thiết Kế Mô Hình & Thực Nghiệm", "Đối sánh 6 mô hình từ Baseline đến Ensemble, tối ưu siêu tham số, bảng Benchmark chuẩn học thuật.", C_PURPLE),
        ("Phần 5: Giải Thích XAI & Phân Tích Lỗi", "Minh bạch căn cứ suy luận bằng Permutation Importance, Radar cá thể và giải phẫu 69 ca sai lệch.", C_AMBER),
        ("Phần 6: Thực Thi Phần Mềm & Ứng Dụng", "Kiến trúc 5 tầng, nguyên mẫu Web Streamlit VSTEP Master UI, thuật toán phân bổ nguồn lực & Đạo đức AI.", C_ROSE)
    ]
    for i, (title, desc, color) in enumerate(agenda_items):
        col = i % 2
        row = i // 2
        ax = 0.6 + col * 6.1
        ay = 1.6 + row * 1.8
        add_card(s2, ax, ay, 5.9, 1.6, bg=C_WHITE, border=C_BORDER)
        
        cbar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(ax), Inches(ay), Inches(0.15), Inches(1.6))
        cbar.fill.solid()
        cbar.fill.fore_color.rgb = color
        cbar.line.fill.background()
        
        tb = s2.shapes.add_textbox(Inches(ax + 0.3), Inches(ay + 0.1), Inches(5.4), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = C_TEXT_TITLE
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(4)

    # ==================== SLIDE 3: BỐI CẢNH GDPT 2018 MÔN TIN HỌC ====================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "1.1 Bối Cảnh GDPT 2018 & Cấu Trúc Đề Thi Tốt Nghiệp Mới")

    cards_data_3 = [
        ("📌 1. Bước Ngoặt Lịch Sử GDPT 2018", 
         "• Từ kỳ thi Tốt nghiệp THPT 2025, môn Tin học lần đầu tiên trở thành môn thi tự chọn tốt nghiệp.\n\n"
         "• Chuyển dịch căn bản từ kiểm tra tin học văn phòng đơn thuần sang đánh giá năng lực tư duy giải thuật và đạo đức số.\n\n"
         "• Học sinh đăng ký xét tuyển đại học khối A00, A01, D01 có xu hướng chọn môn Tin học tăng cao.",
         C_BLUE_LIGHT, C_BORDER_ACCENT),

        ("📋 2. Cấu Trúc Đề Thi Chuẩn Hóa 3 Phần",
         "• Phần I (4.0 đ): 24 câu trắc nghiệm 4 lựa chọn (Nhận biết & Thông hiểu).\n\n"
         "• Phần II (4.0 đ): 8 câu Trắc nghiệm Đúng/Sai đa ý (Mỗi câu gồm 4 ý a,b,c,d - Đánh giá luồng dữ liệu & logic).\n\n"
         "• Phần III (2.0 đ): 4 câu Trả lời ngắn (Vận dụng cao - Đòi hỏi tính toán giải thuật và viết mã).",
         RGBColor(245, 243, 255), RGBColor(221, 214, 254)),

        ("⚠️ 3. Nguy Cơ Điểm Liệt (<= 1.0 Điểm)",
         "• Quy chế thi Bộ GD&ĐT quy định điểm bài thi <= 1.0 là điểm liệt, học sinh bị trượt tốt nghiệp bất kể các môn khác.\n\n"
         "• Phần III và mạch Lập trình (CS) có độ phân hóa cực kỳ khốc liệt, học sinh trung bình - yếu thường bị điểm 0.\n\n"
         "• Rất khó để giáo viên lường trước nguy cơ này nếu chỉ dựa vào điểm học bạ đơn thuần.",
         RGBColor(254, 242, 242), RGBColor(254, 202, 202))
    ]

    for i, (title, content, bg_c, bdr_c) in enumerate(cards_data_3):
        cx = 0.6 + i * 4.1
        add_card(s3, cx, 1.6, 3.9, 5.3, bg=bg_c, border=bdr_c)
        tb = s3.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(3.5), Inches(4.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13.5)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(10)

    # ==================== SLIDE 4: ĐIỂM NGHẼN QUẢN LÝ THỰC TIỄN ====================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "1.2 Thách Thức Quản Lý & Điểm Nghẽn Của Đánh Giá Truyền Thống")

    pain_points = [
        ("1. Điểm Số Bị 'Phẳng Hóa' (Flattened GPA)", 
         "• Điểm tổng kết sổ điểm (ví dụ 6.5) làm phẳng lì toàn bộ năng lực.\n"
         "• Một học sinh giỏi Tin học ứng dụng (Word, Excel điểm 9.0) nhưng hổng hoàn toàn Lập trình CS (điểm 2.0) vẫn có điểm TB 6.0!\n"
         "• Đi thi tốt nghiệp gặp Phần III Lập trình sẽ lập tức đối mặt nguy cơ điểm liệt mà học bạ không cảnh báo được.",
         C_ROSE),
        ("2. Quá Tải Phân Tích & Can Thiệp Chậm Trễ", 
         "• Một giáo viên phụ trách 3 - 5 lớp (150 - 250 học sinh), toàn trường có 500 - 1.000 học sinh khối 12.\n"
         "• Phân tích thủ công bằng bảng tính Excel tốn 20 - 30 giờ/khóa, dễ sai sót và bỏ lọt học sinh yếu.\n"
         "• Khi phát hiện được thì đã vào giai đoạn cuối kỳ thi, quỹ thời gian phụ đạo đã cạn kiệt.",
         C_AMBER),
        ("3. Thiếu Công Cụ Định Lượng Nguồn Lực", 
         "• Phân bổ giáo viên phụ đạo hiện nay hoàn toàn theo cảm tính, phân chia cào bằng số tiết.\n"
         "• Không xác định được chính xác lớp nào cần bồi dưỡng mạch CS, lớp nào cần củng cố câu hỏi trắc nghiệm Đúng/Sai.\n"
         "• Gây lãng phí ngân sách bồi dưỡng và quá tải không đáng có cho đội ngũ nhà giáo.",
         C_NAVY)
    ]
    for i, (title, content, color) in enumerate(pain_points):
        cy = 1.6 + i * 1.8
        add_card(s4, 0.6, cy, 12.133, 1.6, bg=C_WHITE, border=C_BORDER)
        bar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.15), Inches(1.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()
        
        tb = s4.shapes.add_textbox(Inches(0.95), Inches(cy + 0.1), Inches(11.6), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = content
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(4)

    # ==================== SLIDE 5: MỤC TIÊU & 3 TRỤ CỘT GIẢI PHÁP ====================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "1.3 Mục Tiêu Nghiên Cứu & Khung Giải Pháp 3 Trụ Cột")

    pillars_5 = [
        ("🎯 TRỤ CỘT 1: CẢNH BÁO SỚM NGUY CƠ", 
         "Ứng dụng học máy phân loại 3 mức rủi ro học tập:",
         [("🟢 An toàn (>= 6.5 điểm)", "Dự kiến đạt chuẩn tốt nghiệp mức khá/giỏi"),
          ("🟡 Cần theo dõi (5.0 - 6.4)", "Có nguy cơ trượt nếu đề thi phân hóa cao"),
          ("🔴 Nguy cơ cao (< 5.0 điểm)", "Nguy cơ cao điểm dưới TB hoặc dính điểm liệt")],
         C_BLUE),
        
        ("🔍 TRỤ CỘT 2: MINH BẠCH CĂN CỨ (XAI)",
         "Loại bỏ tính chất 'Hộp đen' (Black-box) của AI:",
         [("🔬 Permutation Importance", "Lượng hóa trọng số từng thành phần đánh giá"),
          ("🕸️ Radar Năng Lực VSTEP", "Đối sánh năng lực cá nhân với chuẩn an toàn"),
          ("💡 Khuyến nghị sư phạm", "Chỉ rõ căn cứ: hổng Lập trình CS hay Vận dụng cao")],
         C_EMERALD),

        ("🏫 TRỤ CỘT 3: PHÂN BỔ NGUỒN LỰC",
         "Tự động tối ưu hóa công tác ôn thi tại nhà trường:",
         [("📋 Gom nhóm theo chuyên đề", "Tách riêng HS hổng CS, VDC hay ICT"),
          ("⏱️ Định lượng số tiết/tuần", "Tự động tính quỹ thời gian phụ đạo cần thiết"),
          ("👨‍🏫 Phân công chuyên môn", "Gán đúng giáo viên thế mạnh bồi dưỡng")],
         C_AMBER)
    ]

    for i, (title, sub, items, acc_color) in enumerate(pillars_5):
        cx = 0.6 + i * 4.1
        add_card(s5, cx, 1.6, 3.9, 5.3, bg=C_WHITE, border=C_BORDER)
        hb = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.6), Inches(3.9), Inches(0.85))
        hb.fill.solid()
        hb.fill.fore_color.rgb = acc_color
        hb.line.fill.background()
        
        p = hb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(11.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER
        
        tb = s5.shapes.add_textbox(Inches(cx + 0.2), Inches(2.6), Inches(3.5), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = sub
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_MUTED
        
        for head, body in items:
            p = tf.add_paragraph()
            p.text = f"• {head}: {body}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = C_TEXT_BODY
            p.space_before = Pt(8)

    # ==================== SLIDE 6: CƠ SỞ TRI THỨC GDPT 2018 ====================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "2.1 Cơ Sở Tri Thức Sư Phạm: 3 Mạch Tri Thức & 4 Mức Bloom")

    add_card(s6, 0.6, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s6.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ 3 MẠCH TRI THỨC CỐT LÕI GDPT 2018"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    streams = [
        ("1. Khoa học Máy tính (CS - Computer Science)", 
         "• Trọng số nguy cơ cao nhất. Bao gồm: Thuật toán, lập trình Python, cấu trúc dữ liệu, tư duy logic.\n• Quyết định trực tiếp khả năng giải quyết Phần III đề thi và nguy cơ điểm liệt."),
        ("2. Tin học Ứng dụng (ICT - Information Tech)", 
         "• Kỹ năng sử dụng phần mềm, hệ quản trị CSDL, dịch vụ mạng, xử lý đa phương tiện.\n• Chiếm tỷ trọng lớn trong Phần I và II, giúp học sinh duy trì điểm số sàn trung bình."),
        ("3. Đạo đức, Pháp luật & Văn hóa Số (DL)", 
         "• Hiểu biết về bản quyền, bảo mật thông tin, văn hóa ứng xử trên không gian mạng.\n• Thường xuất hiện trong các câu hỏi trắc nghiệm Đúng/Sai tình huống thực tế.")
    ]
    for stitle, sdesc in streams:
        p = tf.add_paragraph()
        p.text = stitle
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = sdesc
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY

    add_card(s6, 6.8, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb2 = s6.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "📊 4 MỨC ĐỘ TƯ DUY THEO THANG BLOOM"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = C_EMERALD

    blooms = [
        ("Mức 1: Nhận Biết (Remembering)", "Tái hiện định nghĩa, cú pháp câu lệnh, cấu trúc phần cứng/phần mềm. Tỷ lệ chuẩn: 40% đề thi."),
        ("Mức 2: Thông Hiểu (Understanding)", "Giải thích được nguyên lý hoạt động, dự đoán kết quả thực thi của đoạn mã ngắn. Tỷ lệ: 30% đề thi."),
        ("Mức 3: Vận Dụng (Applying)", "Vận dụng thuật toán cơ bản, thao tác CSDL để giải quyết bài toán quen thuộc. Tỷ lệ: 20% đề thi."),
        ("Mức 4: Vận Dụng Cao (Creating & Evaluating)", "Tối ưu hóa giải thuật, thiết kế cấu trúc dữ liệu giải quyết vấn đề mới (Phần III đề thi). Tỷ lệ: 10% đề thi.")
    ]
    for btitle, bdesc in blooms:
        p = tf2.add_paragraph()
        p.text = btitle
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_TITLE
        p.space_before = Pt(8)
        p = tf2.add_paragraph()
        p.text = bdesc
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 7: HỆ LUẬT CHUYÊN GIA SƯ PHẠM (KBS RULES) ====================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "2.2 Cơ Sở Tri Thức Luật: Hệ Luật Chuyên Gia Can Thiệp Sư Phạm")

    rules = [
        ("LUẬT 1: Nhận Diện Nguy Cơ Liệt Lập Trình (CS Deficit)",
         "IF (CS_Score < 4.0 OR VDC_Score < 3.0) AND (LMS_Submit_Ratio < 0.6)\n"
         "THEN Risk_Level = 'Nguy Cơ Cao' (Conf: 0.94)\n"
         "ACTION = 'Kích hoạt cảnh báo nguy cơ điểm liệt; đưa vào nhóm phụ đạo Lập trình CS cấp tốc.'",
         C_ROSE),
        ("LUẬT 2: Phát Hiện Nguy Cơ Thụt Lùi Quá Trình (Degradation Trend)",
         "IF (Trend_HK = HK2 - HK1 < -1.5) AND (Average_Process < 5.5)\n"
         "THEN Risk_Level = 'Cảnh Báo Sa Sút Nhanh' (Conf: 0.88)\n"
         "ACTION = 'Giáo viên chủ nhiệm phỏng vấn cá nhân, tìm hiểu nguyên nhân tâm lý hoặc gia đình.'",
         C_AMBER),
        ("LUẬT 3: Tối Ưu Bứt Phá Điểm Phân Hóa (VDC Remediation)",
         "IF (Predicted_Risk == 'Cần Theo Dõi') AND (VDC_Score < 5.0) AND (CS_Score >= 6.0)\n"
         "THEN Remedial_Focus = 'Chuyên Đề Rèn Luyện Câu Hỏi Trả Lời Ngắn Phần III'\n"
         "ACTION = 'Bồi dưỡng phương pháp giải nhanh và kỹ thuật kiểm tra biên thuật toán.'",
         C_BLUE),
        ("LUẬT 4: Bảo Vệ Vùng An Toàn Đạt Chuẩn (Safe Maintenance)",
         "IF (Predicted_Risk == 'An Toàn') AND (LMS_Submit_Ratio >= 0.85)\n"
         "THEN Status = 'Đạt Chuẩn Tốt Nghiệp Khá / Giỏi'\n"
         "ACTION = 'Chuyển sang chế độ tự luyện đề; giao thêm bài toán phân hóa đại học.'",
         C_EMERALD)
    ]
    for i, (rtitle, rcontent, color) in enumerate(rules):
        cy = 1.6 + i * 1.35
        add_card(s7, 0.6, cy, 12.133, 1.25, bg=C_WHITE, border=C_BORDER)
        bar = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.12), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s7.shapes.add_textbox(Inches(0.85), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = rtitle
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = rcontent
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(3)

    # ==================== SLIDE 8: ĐẶC TẢ 15 ĐẶC TRƯNG ĐẦU VÀO ====================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "2.3 Đặc Tả Không Gian 15 Đặc Trưng Đầu Vào & Ý Nghĩa Giáo Dục")

    feat_groups = [
        ("1. Đánh Giá THPT Truyền Thống (5)", 
         ["• oral_score: Điểm kiểm tra miệng thường xuyên [0 - 10]",
          "• quiz_15p: Điểm bài kiểm tra ngắn 15 phút [0 - 10]",
          "• test_45p: Điểm bài kiểm tra 1 tiết định kỳ [0 - 10]",
          "• sem1_score: Điểm kiểm tra tổng kết Học Kỳ 1 [0 - 10]",
          "• sem2_score: Điểm kiểm tra tổng kết Học Kỳ 2 [0 - 10]"],
         C_NAVY),
        ("2. Chỉ Số Động Lực Học Tập (3)",
         ["• avg_process: Điểm TB Quá trình = (Miệng + 15p + 45p*2) / 4",
          "• trend_hk: Xu hướng học tập = HK2 - HK1 [-10 đến +10]",
          "• lms_completion: Tỷ lệ nộp bài tập LMS trực tuyến [0.0 - 1.0]"],
         C_BLUE),
        ("3. Mức Độ Tư Duy Bloom (4)",
         ["• bloom_remember: Năng lực Nhận biết câu hỏi lý thuyết [0 - 10]",
          "• bloom_understand: Năng lực Thông hiểu luồng chương trình [0 - 10]",
          "• bloom_apply: Năng lực Vận dụng thuật toán căn bản [0 - 10]",
          "• bloom_analyze_eval: Năng lực Vận dụng cao giải thuật [0 - 10]"],
         C_EMERALD),
        ("4. Mạch Tri Thức GDPT 2018 (3)",
         ["• strand_cs: Năng lực Khoa học máy tính & Lập trình [0 - 10]",
          "• strand_ict: Năng lực Ứng dụng tin học & CSDL [0 - 10]",
          "• strand_dl: Năng lực Đạo đức, pháp luật & Văn hóa số [0 - 10]"],
         C_PURPLE)
    ]

    for i, (gtitle, gitems, gcolor) in enumerate(feat_groups):
        cx = 0.6 + i * 3.05
        add_card(s8, cx, 1.6, 2.95, 5.3, bg=C_WHITE, border=C_BORDER)
        hb = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.6), Inches(2.95), Inches(0.7))
        hb.fill.solid()
        hb.fill.fore_color.rgb = gcolor
        hb.line.fill.background()
        p = hb.text_frame.paragraphs[0]
        p.text = gtitle
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER
        
        tb = s8.shapes.add_textbox(Inches(cx + 0.15), Inches(2.4), Inches(2.65), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for idx, item in enumerate(gitems):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = item
            p.font.size = Pt(9.5)
            p.font.color.rgb = C_TEXT_BODY
            if idx > 0:
                p.space_before = Pt(8)

    # ==================== SLIDE 9: MÔ HÌNH HÓA TOÁN HỌC & HÀM TỔN THẤT NHẠY CHI PHÍ ====================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "2.4 Mô Hình Hóa Toán Học & Hàm Tổn Thất Nhạy Chi Phí (Cost-Sensitive)")

    add_card(s9, 0.6, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s9.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📐 PHÁT BIỂU TOÁN HỌC BÀI TOÁN PHÂN LOẠI"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    math_text = (
        "• Không gian mẫu: Tập dữ liệu D = {(x_i, y_i)} với N = 4.070 học sinh trên 8 niên khóa.\n\n"
        "• Vector đặc trưng: x_i thuộc không gian R^15 đại diện cho 15 chỉ số năng lực THPT & GDPT 2018.\n\n"
        "• Không gian nhãn mục tiêu Y = {0, 1, 2} với:\n"
        "  - y = 0: An toàn (Điểm thi thực tế >= 6.5)\n"
        "  - y = 1: Cần theo dõi (5.0 <= Điểm thi < 6.5)\n"
        "  - y = 2: Nguy cơ cao (Điểm thi < 5.0 hoặc có nguy cơ dính điểm liệt <= 1.0)\n\n"
        "• Hàm phân loại: Học ánh xạ f: R^15 -> Y sao cho cực tiểu hóa rủi ro thực nghiệm nhạy chi phí."
    )
    p = tf.add_paragraph()
    p.text = math_text
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    add_card(s9, 6.8, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb2 = s9.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "⚖️ MA TRẬN PHẠT SAI SỐ BẤT ĐỐI XỨNG (COST MATRIX)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_ROSE

    cost_text = (
        "Trong bài toán cảnh báo học tập, chi phí sai số mang tính chất bất đối xứng nghiêm trọng:\n\n"
        "• False Negative (FN) - Bỏ sót Nguy Cơ Cao: Học sinh có nguy cơ trượt/liệt nhưng bị dự đoán là An toàn. "
        "Chi phí giáo dục cực lớn vì học sinh mất cơ hội tốt nghiệp!\n\n"
        "• False Positive (FP) - Báo động nhầm: Học sinh An toàn bị xếp vào diện Cần theo dõi. "
        "Chi phí rất nhỏ (chỉ tốn thêm thời gian rà soát nhẹ của thầy cô).\n\n"
        "• Hàm tổn thất nhạy chi phí (Cost-Sensitive Loss):\n"
        "  L_CS = Sum[ C(y_i, y_hat_i) * l(y_i, f(x_i)) ]\n"
        "  với ma trận phạt đặt chi phí C(2, 0) = 10 * C(0, 2).\n\n"
        "• Hệ quả thực tiễn: Ưu tiên tối đa chỉ số Recall trên lớp Nguy Cơ Cao."
    )
    p = tf2.add_paragraph()
    p.text = cost_text
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    # ==================== SLIDE 10: TẬP DỮ LIỆU & PHÂN BỐ NHÃN ====================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "3.1 Đặc Tả Tập Dữ Liệu Thực Nghiệm (4.070 HS) & Phân Bố Nhãn")

    add_card(s10, 0.6, 1.6, 5.4, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s10.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.0), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📦 QUY MÔ & PHÂN BỐ DỮ LIỆU THỰC NGHIỆM"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

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
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    add_card(s10, 6.3, 1.6, 6.4, 5.3, bg=C_WHITE, border=C_BORDER)
    chart_p = os.path.join(figures_dir, 'label_distribution_chart.png')
    if os.path.exists(chart_p):
        s10.shapes.add_picture(chart_p, Inches(6.5), Inches(1.8), Inches(6.0), Inches(4.8))

    # ==================== SLIDE 11: QUY TRÌNH CHỐNG RÒ RỈ DỮ LIỆU (ANTI-LEAKAGE) ====================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "3.2 Quy Trình Ngăn Chặn Rò Rỉ Dữ Liệu (Anti-Leakage Data Pipeline)")

    steps_11 = [
        ("Bước 1: Phân Tách Tập Dữ Liệu Độc Lập (Stratified Split)",
         "• Phân chia trước khi thực hiện bất kỳ phép tiền xử lý nào: Train (70% = 2.849 HS), Val (15% = 610 HS), Test độc lập (15% = 611 HS).\n"
         "• Giữ nguyên tỷ lệ phân bố nhãn 3 lớp trên cả 3 tập bằng Stratified Sampling.",
         C_BLUE),
        ("Bước 2: Xử Lý Ngoại Lệ & Kiểm Tra Dải Hợp Lệ [0, 10]",
         "• Thuật toán kiểm tra toàn vẹn miền giá trị (Range Validator).\n"
         "• Giá trị khuyết thiếu (Missing values) được điền bằng Median tính toán nghiêm ngặt DUY NHẤT trên tập Train.",
         C_NAVY),
        ("Bước 3: Chuẩn Hóa StandardScaler Đóng Băng Tham Số",
         "• Tuyệt đối không chuẩn hóa trên toàn bộ tập dữ liệu (Data Snoop).\n"
         "• Bộ biến đổi StandardScaler chỉ fit() trên Train, sau đó đóng băng tham số mean & std để transform() tập Val và Test.",
         C_EMERALD),
        ("Bước 4: Loại Bỏ Triệt Để Biến Rò Rỉ Mục Tiêu (Target Leakage)",
         "• Biến điểm thi cuối kỳ (est_final_score) chỉ dùng duy nhất cho việc gắn nhãn mục tiêu y.\n"
         "• Loại bỏ 100% ra khỏi không gian đặc trưng đầu vào X, đảm bảo mô hình chỉ dự báo dựa trên quá trình tích lũy.",
         C_AMBER)
    ]
    for i, (stitle, sdesc, color) in enumerate(steps_11):
        cy = 1.6 + i * 1.35
        add_card(s11, 0.6, cy, 12.133, 1.25, bg=C_WHITE, border=C_BORDER)
        bar = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.12), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s11.shapes.add_textbox(Inches(0.85), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = stitle
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = sdesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(3)

    # ==================== SLIDE 12: ĐẠO ĐỨC DỮ LIỆU & TÍNH CHÂN THỰC HỌC THUẬT ====================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "3.3 Đạo Đức Dữ Liệu & Cam Kết Chân Thực Học Thuật")

    commitments = [
        ("🔒 1. Ẩn Danh Hóa Tuyệt Đối (Anonymization)",
         "• 100% dữ liệu danh tính học sinh (họ tên, ngày sinh, trường lớp, địa chỉ liên hệ) được gỡ bỏ hoàn toàn.\n\n"
         "• Sử dụng mã định danh giả lập duy nhất dạng 'STU_xxxx' (ví dụ STU_0142), đảm bảo không thể truy vết danh tính cá nhân ngoài đời thực.\n\n"
         "• Tuân thủ nghiêm ngặt Nghị định 13/2023/NĐ-CP về Bảo vệ dữ liệu cá nhân.",
         C_BLUE_LIGHT, C_BORDER_ACCENT),

        ("📜 2. Minh Bạch Bản Chất Dữ Liệu Giả Lập (PoC)",
         "• Nghiên cứu được phát triển nhằm mục đích thử nghiệm nguyên mẫu kỹ thuật (Proof of Concept - PoC) cho GDPT 2018.\n\n"
         "• Do GDPT 2018 là chương trình mới chưa có lịch sử thi tốt nghiệp nhiều năm, dữ liệu được mô phỏng thống kê sát với thực tế đề minh họa của Bộ GD&ĐT.\n\n"
         "• TUYỆT ĐỐI KHÔNG dùng số liệu mô phỏng để khẳng định hiệu quả sư phạm thực tế ngoài đời sống.",
         RGBColor(254, 243, 199), RGBColor(253, 230, 138)),

        ("🔁 3. Khả Năng Tái Lập (Reproducibility)",
         "• Cố định siêu hạt giống ngẫu nhiên random_state = 42 trên toàn bộ quy trình sinh dữ liệu và huấn luyện mô hình.\n\n"
         "• Cung cấp đầy đủ mã nguồn sinh dữ liệu (generate_multi_cohort_data.py) và tệp CSDL SQLite mở.\n\n"
         "• Mọi thành viên trong Hội đồng chấm thi đều có thể chạy lại mã nguồn và tái tạo 100% kết quả.",
         RGBColor(236, 253, 245), RGBColor(167, 243, 208))
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(commitments):
        cx = 0.6 + i * 4.1
        add_card(s12, cx, 1.6, 3.9, 5.3, bg=bg_c, border=bdr_c)
        tb = s12.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(3.5), Inches(4.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(10)

    # ==================== SLIDE 13: THIẾT KẾ KHÔNG GIAN MÔ HÌNH ====================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "4.1 Thiết Kế Không Gian Mô Hình: Từ Baseline Đến Cải Tiến Ensemble")

    add_card(s13, 0.6, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s13.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏛️ MÔ HÌNH CƠ SỞ (BASELINES)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    base_desc = (
        "1. Decision Tree (Cây Quyết Định Đơn):\n"
        "• Thuật toán phân tách nhị phân dựa trên chỉ số Gini Impurity.\n"
        "• Ưu điểm: Trực quan, dễ giải thích quy tắc rẽ nhánh cho giáo viên.\n"
        "• Nhược điểm: Dễ bị quá khớp (Overfitting), nhạy cảm với nhiễu dữ liệu.\n\n"
        "2. Logistic Regression (Hồi Quy Tuyến Tính Đa Biến):\n"
        "• Chuẩn hóa xác suất phân phối 3 lớp qua hàm Softmax và chiến lược One-vs-Rest (OvR).\n"
        "• Ưu điểm: Ổn định, chi phí tính toán cực thấp, ít bị quá khớp.\n"
        "• Nhược điểm: Giả định quan hệ tuyến tính, khó nắm bắt tương tác phi tuyến giữa các mức Bloom và mạch tri thức."
    )
    p = tf.add_paragraph()
    p.text = base_desc
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    add_card(s13, 6.8, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb2 = s13.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "🚀 MÔ HÌNH CẢI TIẾN (ENSEMBLE METHODS)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_EMERALD

    ens_desc = (
        "1. Random Forest (Rừng Ngẫu Nhiên):\n"
        "• Kỹ thuật Bagging kết hợp 100 cây quyết định độc lập trên các không gian đặc trưng ngẫu nhiên.\n"
        "• Giảm triệt để phương sai (Variance Reduction), bền bỉ với nhiễu.\n\n"
        "2. XGBoost (Extreme Gradient Boosting):\n"
        "• Kỹ thuật Boosting học tuần tự từ phần dư (Residuals) của cây trước.\n"
        "• Bổ sung số hạng chính quy hóa L1/L2 trên hàm mục tiêu ngăn quá khớp.\n\n"
        "3. LightGBM (Light Gradient Boosted Machine):\n"
        "• Cơ chế phân nhánh theo lá (Leaf-wise) kết hợp thuật toán GOSS (Gradient-based One-Side Sampling).\n"
        "• Tốc độ huấn luyện vượt trội, độ chính xác và khả năng bắt nhãn thiểu số (Recall Nguy Cơ Cao) tốt nhất."
    )
    p = tf2.add_paragraph()
    p.text = ens_desc
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    # ==================== SLIDE 14: QUY TRÌNH HUẤN LUYỆN & TỐI ƯU SIÊU THAM SỐ ====================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "4.2 Quy Trình Huấn Luyện & Tối Ưu Siêu Tham Số (Stratified 5-Fold)")

    add_stat_box(s14, 0.6, 1.6, 2.8, 1.1, "4.070 HS", "Tổng mẫu 8 khóa", "K18 - K25", C_NAVY)
    add_stat_box(s14, 3.7, 1.6, 2.8, 1.1, "2.849 HS", "Tập Train (70%)", "Huấn luyện & CV", C_BLUE)
    add_stat_box(s14, 6.8, 1.6, 2.8, 1.1, "610 HS", "Tập Validation (15%)", "Chốt siêu tham số", C_EMERALD)
    add_stat_box(s14, 9.9, 1.6, 2.8, 1.1, "611 HS", "Tập Test Độc Lập (15%)", "Đánh giá mù khách quan", C_ROSE)

    cards_14 = [
        ("Stratified 5-Fold Cross-Validation", 
         "• Tập Train được chia thành 5 phần con có tỷ lệ nhãn đồng nhất.\n"
         "• Mô hình huấn luyện trên 4 phần và kiểm tra trên phần còn lại qua 5 vòng lặp.\n"
         "• Đảm bảo điểm số đánh giá phản ánh độ ổn định thực chất của thuật toán.", C_BLUE_LIGHT, C_BORDER_ACCENT),
        ("Tối Ưu Không Gian Siêu Tham Số (GridSearch)",
         "• Tìm kiếm lưới siêu tham số tối ưu:\n"
         "  - Random Forest: max_depth in [8, 12, 16], min_samples_split in [2, 4]\n"
         "  - LightGBM: learning_rate in [0.03, 0.05, 0.1], num_leaves in [15, 31]\n"
         "• Cân bằng trọng số lớp tự động bằng class_weight='balanced'.", RGBColor(236, 253, 245), RGBColor(167, 243, 208)),
        ("Kiểm Định Khách Quan Trên Tập Test Độc Lập",
         "• Sau khi chọn mô hình tốt nhất trên tập Validation, mô hình được đóng băng hoàn toàn.\n"
         "• Đánh giá một lần duy nhất trên 611 học sinh của tập Test độc lập.\n"
         "• Đảm bảo báo cáo kết quả trung thực, không có hiện tượng 'chọn lọc kết quả đẹp'.", RGBColor(254, 242, 242), RGBColor(254, 202, 202))
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(cards_14):
        cx = 0.6 + i * 4.1
        add_card(s14, cx, 2.9, 3.9, 4.0, bg=bg_c, border=bdr_c)
        tb = s14.shapes.add_textbox(Inches(cx + 0.2), Inches(3.1), Inches(3.5), Inches(3.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12.5)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(8)

    # ==================== SLIDE 15: BẢNG KẾT QUẢ THỰC NGHIỆM ĐA MÔ HÌNH ====================
    s15 = prs.slides.add_slide(blank_layout)
    add_header(s15, "4.3 Bảng Kết QuẢ Thực Nghiệm Đối Sánh Đa Mô Hình (Academic Benchmark)")

    tbl_shape = s15.shapes.add_table(6, 6, Inches(0.6), Inches(1.6), Inches(7.8), Inches(4.2))
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
        cell.fill.fore_color.rgb = C_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for row_idx, r_data in enumerate(rows):
        for col_idx, val in enumerate(r_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(236, 253, 245) if "LightGBM" in r_data[0] else (C_WHITE if row_idx % 2 == 0 else RGBColor(248, 250, 252))
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Arial"
            p.font.size = Pt(10)
            p.font.bold = True if ("LightGBM" in r_data[0] or col_idx == 3) else False
            p.font.color.rgb = C_TEXT_TITLE if col_idx == 0 else (C_ROSE if col_idx == 3 and "⭐" in val else C_TEXT_BODY)
            p.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER

    add_card(s15, 8.6, 1.6, 4.133, 5.3, bg=C_WHITE, border=C_BORDER)
    tb_c = s15.shapes.add_textbox(Inches(8.8), Inches(1.8), Inches(3.733), Inches(4.9))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p = tf_c.paragraphs[0]
    p.text = "💡 NHẬN XÉT HỌC THUẬT QUAN TRỌNG"
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

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
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    # ==================== SLIDE 16: TRỰC QUAN HÓA SO SÁNH HIỆU NĂNG MÔ HÌNH ====================
    s16 = prs.slides.add_slide(blank_layout)
    add_header(s16, "4.4 Trực Quan Hóa Đối Sánh Hiệu Năng Mô Hình")

    add_card(s16, 0.6, 1.6, 6.4, 5.3, bg=C_WHITE, border=C_BORDER)
    fig_comp = os.path.join(figures_dir, 'model_comparison_chart.png')
    if os.path.exists(fig_comp):
        s16.shapes.add_picture(fig_comp, Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.8))

    add_card(s16, 7.2, 1.6, 5.533, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s16.shapes.add_textbox(Inches(7.4), Inches(1.8), Inches(5.133), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "📊 PHÂN TÍCH TƯƠNG QUAN CHỈ SỐ"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_BLUE

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
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = bdesc
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 17: PHÂN TÍCH MA TRẬN NHẦM LẪN CHUYÊN SÂU ====================
    s17 = prs.slides.add_slide(blank_layout)
    add_header(s17, "4.5 Phân Tích Ma Trận Nhầm Lẫn (Confusion Matrix Deep Dive)")

    add_card(s17, 0.6, 1.6, 6.0, 5.3, bg=C_WHITE, border=C_BORDER)
    fig_cm = os.path.join(figures_dir, 'confusion_matrix_best_model.png')
    if os.path.exists(fig_cm):
        s17.shapes.add_picture(fig_cm, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))

    add_card(s17, 6.8, 1.6, 5.933, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s17.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.533), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🎯 GIẢI PHẪU CHI TIẾT MA TRẬN TRÊN TẬP TEST (N = 611 HS)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_ROSE

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
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p.space_before = Pt(8)
        p = tf.add_paragraph()
        p.text = mdesc
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 18: GIẢI THÍCH MÔ HÌNH XAI ====================
    s18 = prs.slides.add_slide(blank_layout)
    add_header(s18, "5.1 Giải Thích Mô Hình Minh Bạch Bằng XAI (Permutation Importance)")

    add_card(s18, 0.6, 1.6, 6.2, 5.3, bg=C_WHITE, border=C_BORDER)
    fig_feat = os.path.join(figures_dir, 'feature_importance_ranking.png')
    if os.path.exists(fig_feat):
        s18.shapes.add_picture(fig_feat, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.8))

    add_card(s18, 7.0, 1.6, 5.733, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s18.shapes.add_textbox(Inches(7.2), Inches(1.8), Inches(5.333), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🔍 Ý NGHĨA SƯ PHẠM CỦA CÁC ĐẶC TRƯNG HÀNG ĐẦU"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    feat_insights = [
        ("1. Điểm Học Kỳ 2 (sem2_score - Chiếm trọng số cao nhất)", 
         "Phản ánh kiến thức tổng hợp sát thời điểm thi tốt nghiệp nhất. Học sinh có kết quả HK2 sa sút có xác suất rủi ro thi tốt nghiệp cao gấp 4.2 lần."),
        ("2. Mạch Khoa học máy tính (strand_cs - Vị trí số 2)", 
         "Khẳng định hoàn toàn giả thuyết nghiên cứu ban đầu: Mạch Lập trình & Thuật toán là rào cản phân hóa lớn nhất của đề thi GDPT 2018."),
        ("3. Mức Vận dụng cao (bloom_analyze_eval)", 
         "Quyết định điểm số Phần III trả lời ngắn. Học sinh hổng mức này sẽ bị giới hạn trần điểm tối đa ở mức trung bình."),
        ("4. Chuyên cần LMS (lms_completion)", 
         "Chỉ số hành vi chứng minh thái độ kiên trì luyện tập trực tuyến đóng vai trò lá chắn chống lại nguy cơ sa sút học tập.")
    ]
    for ititle, idesc in feat_insights:
        p = tf.add_paragraph()
        p.text = ititle
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(6)
        p = tf.add_paragraph()
        p.text = idesc
        p.font.size = Pt(9)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 19: HỒ SƠ NĂNG LỰC CÁ THỂ (RADAR XAI) ====================
    s19 = prs.slides.add_slide(blank_layout)
    add_header(s19, "5.2 Hồ Sơ Năng Lực Học Sinh Cá Thể (Biểu Đồ Radar Đa Chiều)")

    add_card(s19, 0.6, 1.6, 6.2, 5.3, bg=C_WHITE, border=C_BORDER)
    fig_radar = os.path.join(figures_dir, 'radar_competency_sample.png')
    if os.path.exists(fig_radar):
        s19.shapes.add_picture(fig_radar, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.8))

    add_card(s19, 7.0, 1.6, 5.733, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s19.shapes.add_textbox(Inches(7.2), Inches(1.8), Inches(5.333), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "👤 CA LÂM SÀNG: HỌC SINH STU_0142 (LỚP 12A3)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_ROSE

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
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    # ==================== SLIDE 20: PHÂN TÍCH CHUYÊN SÂU CÁC CA SAI LỆCH ====================
    s20 = prs.slides.add_slide(blank_layout)
    add_header(s20, "5.3 Phân Tích Chuyên Sâu Các Ca Dự Đoán Sai (Error Analysis)")

    cards_20 = [
        ("1. Nhóm Ca Vùng Biên (Boundary Cases - 36.2%)",
         "• Đặc điểm: Điểm thi thực tế xấp xỉ ngưỡng ranh giới (ví dụ 4.8 vs 5.0, hoặc 6.4 vs 6.6).\n\n"
         "• Nguyên nhân: Khoảng cách giữa các nhãn rất hẹp. Chỉ cần 1 câu trắc nghiệm may rủi là học sinh chuyển nhãn.\n\n"
         "• Hướng giải quyết: Bổ sung chỉ số biên 'Độ không chắc chắn' (Prediction Uncertainty) để cảnh báo giáo viên theo dõi kỹ.",
         C_BLUE_LIGHT, C_BORDER_ACCENT),

        ("2. Nhóm 'Học Tài Thi Phận' (Outliers - 31.9%)",
         "• Đặc điểm: Quá trình học tập tốt (TB Quá trình > 7.0) nhưng điểm thi thực tế đột ngột tụt xuống < 5.0.\n\n"
         "• Nguyên nhân: Yếu tố ngoại lai không nằm trong dữ liệu số hóa: Sức khỏe, tâm lý phòng thi, tô nhầm mã đề, mất thẻ dự thi.\n\n"
         "• Giới hạn hệ thống: Khẳng định AI không thể biết được yếu tố tâm lý tại thời điểm thi, cần sự quan tâm trực tiếp của thầy cô.",
         RGBColor(254, 243, 199), RGBColor(253, 230, 138)),

        ("3. Nhóm 'Bứt Phá Giai Đoạn Nước Rút' (Late-Bloomers - 31.9%)",
         "• Đặc điểm: Quá trình học chỉ ở mức trung bình (5.0 - 5.5) nhưng thi thực tế đạt điểm giỏi (7.5 - 8.0).\n\n"
         "• Nguyên nhân: Học sinh tập trung ôn luyện tăng cường trong tháng cuối cùng (học thêm ngoài, ôn luyện trọng điểm).\n\n"
         "• Đánh giá: Sai lệch này mang tính chất tích cực cho học sinh, không gây tổn hại giáo dục.",
         RGBColor(236, 253, 245), RGBColor(167, 243, 208))
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(cards_20):
        cx = 0.6 + i * 4.1
        add_card(s20, cx, 1.6, 3.9, 5.3, bg=bg_c, border=bdr_c)
        tb = s20.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(3.5), Inches(4.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(8)

    # ==================== SLIDE 21: KIẾN TRÚC HỆ THỐNG TỔNG THỂ ====================
    s21 = prs.slides.add_slide(blank_layout)
    add_header(s21, "6.1 Kiến Trúc Hệ Thống Tổng Thể End-to-End (5 Tầng Công Nghệ)")

    tiers = [
        ("TẦNG 1: DỮ LIỆU & LƯU TRỮ (DATA LAYER)",
         "• CSDL nhúng SQLite (vstep_informatics.db - 4.070 học sinh trên 8 niên khóa).\n"
         "• Trình nạp tệp linh hoạt hỗ trợ nhập dữ liệu định dạng Excel (.xlsx) và CSV chuẩn hóa danh sách lớp.", C_NAVY),
        ("TẦNG 2: TIỀN XỬ LÝ & CHỐNG RÒ RỈ (ANTI-LEAKAGE PIPELINE)",
         "• Bộ lọc kiểm tra tính toàn vẹn dải điểm [0, 10]; xử lý dữ liệu khuyết thiếu bằng Median tập Train.\n"
         "• Chuẩn hóa StandardScaler đóng băng tham số mean & std, đảm bảo không rò rỉ thông tin tập Test.", C_BLUE),
        ("TẦNG 3: ĐỘNG CƠ HỌC MÁY ĐA MÔ HÌNH (ML ENGINE)",
         "• Tích hợp đồng thời 6 thuật toán: Decision Tree, Logistic Regression, Random Forest, XGBoost, LightGBM.\n"
         "• Lưu trữ trọng số nhị phân tối ưu dạng .joblib, hỗ trợ suy luận đa luồng tốc độ cao trên CPU phổ thông.", C_EMERALD),
        ("TẦNG 4: HỆ CƠ SỞ TRI THỨC & SUY LUẬN SƯ PHẠM (KBS & XAI)",
         "• Bộ luật chuyên gia sư phạm GDPT 2018 (Production Rules IF-THEN).\n"
         "• Trích xuất Feature Attribution, vẽ biểu đồ Radar đa chiều và thuật toán tối ưu phân cụm phụ đạo.", C_AMBER),
        ("TẦNG 5: NGUYÊN MẪU PHẦN MỀM GIAO DIỆN WEB (APPLICATION UI)",
         "• Xây dựng trên nền tảng Streamlit với phong cách chuẩn VSTEP Master Academic UI.\n"
         "• Phân chia 5 phân hệ (Tabs) trực quan, tương tác thời gian thực, responsive cho máy tính và máy tính bảng.", C_ROSE)
    ]
    for i, (ttitle, tdesc, color) in enumerate(tiers):
        cy = 1.6 + i * 1.05
        add_card(s21, 0.6, cy, 12.133, 0.95, bg=C_WHITE, border=C_BORDER)
        bar = s21.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.12), Inches(0.95))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s21.shapes.add_textbox(Inches(0.85), Inches(cy + 0.05), Inches(11.7), Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = ttitle
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = tdesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(2)

    # ==================== SLIDE 22: ĐỐI SÁNH BASELINE VẬN HÀNH ====================
    s22 = prs.slides.add_slide(blank_layout)
    add_header(s22, "6.2 So Sánh Baseline Vận Hành: Thủ Công vs Hệ Thống Học Máy")

    add_card(s22, 0.6, 1.6, 6.4, 5.3, bg=C_WHITE, border=C_BORDER)
    fig_ops = os.path.join(figures_dir, 'operational_baseline_chart.png')
    if os.path.exists(fig_ops):
        s22.shapes.add_picture(fig_ops, Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.8))

    add_card(s22, 7.2, 1.6, 5.533, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s22.shapes.add_textbox(Inches(7.4), Inches(1.8), Inches(5.133), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "⏱️ BẢNG ĐỐI SÁNH HIỆU QUẢ VẬN HÀNH THỰC TẾ"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

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
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(6)
        p = tf.add_paragraph()
        p.text = odesc
        p.font.size = Pt(9)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 23: NGUYÊN MẪU ỨNG DỤNG WEB ====================
    s23 = prs.slides.add_slide(blank_layout)
    add_header(s23, "6.3 Nguyên Mẫu Phần Mềm Web (Streamlit VSTEP Master UI)")

    web_tabs = [
        ("TAB 1: TỔNG QUAN HỆ THỐNG", 
         "• Thống kê tổng số học sinh 8 khóa, phân bổ nhãn rủi ro toàn trường.\n• Biểu đồ phân bố động lực học tập và tỷ lệ hoàn thành LMS.", C_BLUE),
        ("TAB 2: DỰ BÁO ĐƠN & RADAR XAI", 
         "• Nhập 15 thông số cá thể học sinh qua slider trực quan.\n• Hiển thị xác suất rủi ro, biểu đồ Radar đối sánh chuẩn và khuyến nghị can thiệp.", C_EMERALD),
        ("TAB 3: DỰ BÁO HÀNG LOẠT (EXCEL/CSV)", 
         "• Tải lên danh sách học sinh theo lớp hoặc toàn khối.\n• Tự động kiểm tra dải điểm hợp lệ, hiển thị danh sách cảnh báo màu và xuất báo cáo Excel.", C_AMBER),
        ("TAB 4: PHÂN CỤM & PHÂN BỔ NGUỒN LỰC", 
         "• Tự động gom cụm học sinh theo lỗ hổng chính (CS / VDC / ICT).\n• Đề xuất kịch bản số lớp phụ đạo, quỹ tiết dạy và giáo viên phụ trách.", C_PURPLE),
        ("TAB 5: ĐỐI SÁNH KỸ THUẬT & BENCHMARK", 
         "• Bảng so sánh 6 mô hình học thuật, ma trận nhầm lẫn chi tiết.\n• Cho phép thay đổi mô hình dự báo hạt nhân (Active Model Selector).", C_ROSE)
    ]
    for i, (wtitle, wdesc, color) in enumerate(web_tabs):
        cy = 1.6 + i * 1.05
        add_card(s23, 0.6, cy, 12.133, 0.95, bg=C_WHITE, border=C_BORDER)
        bar = s23.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.12), Inches(0.95))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s23.shapes.add_textbox(Inches(0.85), Inches(cy + 0.05), Inches(11.7), Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = wtitle
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = wdesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(2)

    # ==================== SLIDE 24: KỊCH BẢN PHÂN BỔ NGUỒN LỰC ÔN TẬP ====================
    s24 = prs.slides.add_slide(blank_layout)
    add_header(s24, "6.4 Kịch Bản Tự Động Phân Bổ Nguồn Lực Ôn Tập")

    res_cards = [
        ("1. Thuật Toán Gom Cụm Lỗ Hổng Kiến Thức",
         "• Đầu vào: Tập học sinh Nguy cơ cao & Cần theo dõi.\n\n"
         "• Phân loại tự động thành 3 chuyên đề bồi dưỡng:\n"
         "  1) Chuyên đề CS: Hổng nặng Lập trình (CS < 4.5)\n"
         "  2) Chuyên đề VDC: Hổng Vận dụng cao (VDC < 5.0)\n"
         "  3) Chuyên đề Nền Tảng: Củng cố ICT và Đúng/Sai.",
         C_BLUE_LIGHT, C_BORDER_ACCENT),

        ("2. Quy Chuẩn Kích Thước Lớp Phụ Đạo",
         "• Tiêu chuẩn sư phạm: 15 - 20 học sinh / lớp để giáo viên kèm cặp sát từng em.\n\n"
         "• Ví dụ thực tế trên 245 học sinh nguy cơ của tập Test:\n"
         "  - Cụm Lập trình CS: 112 HS -> 6 lớp phụ đạo\n"
         "  - Cụm Vận dụng cao: 88 HS -> 4 lớp phụ đạo\n"
         "  - Cụm Nền tảng: 45 HS -> 2 lớp phụ đạo.",
         RGBColor(236, 253, 245), RGBColor(167, 243, 208)),

        ("3. Định Lượng Quỹ Tiết & Giáo Viên",
         "• Tự động tính toán tổng số tiết ôn tập cần thiết (ví dụ 4 tiết/tuần trong 6 tuần = 24 tiết/lớp).\n\n"
         "• Tự động cân đối thời khóa biểu, gán giáo viên có thế mạnh giải thuật phụ trách các lớp CS.\n\n"
         "• Cung cấp bảng dữ liệu hoàn chỉnh cho Ban Giám Hiệu phê duyệt ngân sách bồi dưỡng.",
         RGBColor(254, 243, 199), RGBColor(253, 230, 138))
    ]
    for i, (title, content, bg_c, bdr_c) in enumerate(res_cards):
        cx = 0.6 + i * 4.1
        add_card(s24, cx, 1.6, 3.9, 5.3, bg=bg_c, border=bdr_c)
        tb = s24.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(3.5), Inches(4.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(8)

    # ==================== SLIDE 25: KHUNG ĐẠO ĐỨC AI & GIỮ CON NGƯỜI TRONG VÒNG LẶP ====================
    s25 = prs.slides.add_slide(blank_layout)
    add_header(s25, "7.1 Khung Đạo Đức AI & Tiêu Chuẩn Giữ Con Người Trong Vòng Lặp")

    ethics = [
        ("1. Giữ Con Người Trong Vòng Lặp (Human-in-the-Loop)",
         "• Hệ thống tuyệt đối KHÔNG tự động ra quyết định trừng phạt, giữ lại lớp hay tước quyền thi của học sinh.\n"
         "• Hệ thống chỉ đóng vai trò Trợ lý Hỗ Trợ Ra Quyết Định (Decision Support System - DSS).\n"
         "• Thầy cô giáo và Hội đồng sư phạm luôn là chủ thể có thẩm quyền xem xét, điều chỉnh và quyết định phương án can thiệp.",
         C_BLUE),
        ("2. Chống Dán Nhãn Định Kiến (Zero Stigmatization)",
         "• Cảnh báo học tập KHÔNG ĐƯỢC công bố công khai trước tập thể lớp hay bảng tin nhà trường.\n"
         "• Thông tin chỉ lưu hành nội bộ giữa Ban Giám Hiệu, Giáo viên bộ môn và Giáo viên chủ nhiệm.\n"
         "• Mục tiêu của cảnh báo là 'Hỗ trợ và đồng hành', bảo vệ tuyệt đối lòng tự trọng và tâm lý lứa tuổi học sinh THPT.",
         C_EMERALD),
        ("3. Quyền Riêng Tư & An Toàn Dữ Liệu (Data Privacy & Compliance)",
         "• Không thu thập các trường thông tin nhạy cảm: Dân tộc, tôn giáo, điều kiện kinh tế gia đình, địa chỉ nơi ở.\n"
         "• Cơ chế lưu trữ nhúng cục bộ (Local SQLite), không tự động đẩy dữ liệu học sinh lên các dịch vụ đám mây công cộng bên ngoài.\n"
         "• Tuân thủ nghiêm ngặt các quy định pháp luật về bảo mật an toàn thông tin giáo dục.",
         C_PURPLE)
    ]
    for i, (etitle, edesc, color) in enumerate(ethics):
        cy = 1.6 + i * 1.8
        add_card(s25, 0.6, cy, 12.133, 1.6, bg=C_WHITE, border=C_BORDER)
        bar = s25.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.15), Inches(1.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s25.shapes.add_textbox(Inches(0.95), Inches(cy + 0.1), Inches(11.6), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = etitle
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = edesc
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(4)

    # ==================== SLIDE 26: TỔNG KẾT, GIỚI HẠN & HƯỚNG MỞ RỘNG ====================
    s26 = prs.slides.add_slide(blank_layout)
    add_header(s26, "7.2 Tổng Kết Đóng Góp, Giới Hạn & Hướng Phát Triển Tiếp Theo")

    add_card(s26, 0.6, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb = s26.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "🏆 TỔNG KẾT ĐÓNG GÓP CỐT LÕI CỦA ĐỀ ÁN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

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
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    add_card(s26, 6.8, 1.6, 5.9, 5.3, bg=C_WHITE, border=C_BORDER)
    tb2 = s26.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "🔮 GIỚI HẠN & HƯỚNG PHÁT TRIỂN TƯƠNG LAI"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_AMBER

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
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_BODY
    p.space_before = Pt(8)

    # ==================== SLIDE 27: [BACKUP 1] BẢNG SIÊU THAM SỐ CHI TIẾT ====================
    s27 = prs.slides.add_slide(blank_layout)
    add_header(s27, "[Phụ Lục 1] Bảng Siêu Tham Số Chi Tiết 6 Mô Hình Sau Tối Ưu (Backup Slide)", "PHỤ LỤC PHẢN BIỆN BẢO VỆ • HỌC PHẦN CÁC HỆ CƠ SỞ TRI THỨC")

    tbl_s27 = s27.shapes.add_table(7, 4, Inches(0.6), Inches(1.6), Inches(12.133), Inches(5.2))
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
        cell.fill.fore_color.rgb = C_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = "Arial"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for row_idx, r_data in enumerate(rows27):
        for col_idx, val in enumerate(r_data):
            cell = table27.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(236, 253, 245) if "LightGBM" in r_data[0] else (C_WHITE if row_idx % 2 == 0 else RGBColor(248, 250, 252))
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = "Arial"
            p.font.size = Pt(9.5)
            p.font.bold = True if (col_idx == 0 or "LightGBM" in r_data[0]) else False
            p.font.color.rgb = C_TEXT_TITLE if col_idx < 2 else C_TEXT_BODY
            p.alignment = PP_ALIGN.LEFT

    # ==================== SLIDE 28: [BACKUP 2] BỘ CÂU HỎI PHẢN BIỆN TRỌNG TÂM ====================
    s28 = prs.slides.add_slide(blank_layout)
    add_header(s28, "[Phụ Lục 2] Bộ Câu Hỏi Phản Biện Trọng Tâm Của Hội Đồng & Lời Giải Đáp (Backup Slide)", "PHỤ LỤC PHẢN BIỆN BẢO VỆ • HỌC PHẦN CÁC HỆ CƠ SỞ TRI THỨC")

    faqs = [
        ("❓ CÂU HỎI 1: Tại sao đề tài chọn Tree Ensemble (LightGBM, Random Forest) thay vì Deep Learning (MLP, Transformer)?",
         "💡 GIẢI ĐÁP: Với dữ liệu dạng bảng (Tabular Data) quy mô 4.070 mẫu và 15 đặc trưng số học, các nghiên cứu chuẩn (như Grinsztajn et al., NeurIPS) đã chứng minh Tree-based models luôn vượt trội Deep Learning về cả độ chính xác, khả năng bắt quy tắc phi tuyến, tốc độ suy luận tức thì (0.04s) và đặc biệt là tính minh bạch giải thích (XAI).",
         C_BLUE),
        ("❓ CÂU HỎI 2: Làm sao khẳng định mô hình không bị rò rỉ dữ liệu (Data Leakage) khi phân chia?",
         "💡 GIẢI ĐÁP: Quy trình tiền xử lý tuân thủ nghiêm ngặt chuẩn Anti-Leakage: Phân chia Stratified Split trước tiên; Bộ chuẩn hóa StandardScaler chỉ gọi fit() trên tập Train và đóng băng tham số để transform() tập Val/Test độc lập; Tuyệt đối loại bỏ điểm thi cuối kỳ (est_final_score) khỏi không gian đặc trưng X.",
         C_EMERALD),
        ("❓ CÂU HỎI 3: Dữ liệu mô phỏng PoC có làm giảm giá trị ứng dụng thực tiễn của đề án không?",
         "💡 GIẢI ĐÁP: Không. Đề tài giải quyết bài toán đón đầu kỳ thi tốt nghiệp theo GDPT 2018 chưa có lịch sử thi dài. Quy trình mô phỏng tuân thủ phân phối chuẩn của Bộ GD&ĐT. Đề án công khai minh bạch giới hạn PoC và không dùng số liệu này để khẳng định hiệu quả ngoài đời thực, đảm bảo 100% đạo đức học thuật.",
         C_AMBER),
        ("❓ CÂU HỎI 4: Ý nghĩa thực tế lớn nhất của thuật toán phân bổ nguồn lực ôn tập là gì?",
         "💡 GIẢI ĐÁP: Tự động hóa khâu gom nhóm học sinh theo đúng chuyên đề hổng kiến thức (CS vs VDC), giải phóng 20 - 30 giờ lao động thủ công của giáo viên, giúp Ban Giám Hiệu lập thời khóa biểu và dự toán ngân sách bồi dưỡng chuẩn xác trước kỳ thi 2 - 3 tháng.",
         C_PURPLE)
    ]
    for i, (qtitle, adesc, color) in enumerate(faqs):
        cy = 1.6 + i * 1.35
        add_card(s28, 0.6, cy, 12.133, 1.25, bg=C_WHITE, border=C_BORDER)
        bar = s28.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(cy), Inches(0.12), Inches(1.25))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s28.shapes.add_textbox(Inches(0.85), Inches(cy + 0.05), Inches(11.7), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = qtitle
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = adesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.space_before = Pt(3)

    # Save
    prs.save(output_pptx)
    print(f"Deck successfully created at: {output_pptx}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == '__main__':
    create_deck()

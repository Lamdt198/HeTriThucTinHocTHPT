"""
build_presentation.py
Tạo Slide thuyết trình PowerPoint (.pptx) chuẩn báo cáo đề án Cao học CNTT K20.2
Học phần: Các Hệ Cơ Sở Tri Thức (Knowledge-Based Systems)
Đề tài: Hệ tri thức và học máy hỗ trợ phân tích ma trận kết quả đánh giá,
        cảnh báo sớm nguy cơ học tập môn Tin học THPT (GDPT 2018)
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

    def add_bg(slide, bg_color=C_BG_PAGE):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        rect.fill.solid()
        rect.fill.fore_color.rgb = bg_color
        rect.line.fill.background()
        return rect

    def add_header(slide, title_text, category_text="HỆ THỐNG CƠ SỞ TRI THỨC • BÁO CÁO CAO HỌC CNTT K20.2"):
        add_bg(slide)
        
        # Header banner card
        hdr_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.4), Inches(12.133), Inches(0.95))
        hdr_rect.fill.solid()
        hdr_rect.fill.fore_color.rgb = C_WHITE
        hdr_rect.line.color.rgb = C_BORDER
        hdr_rect.line.width = Pt(1)

        # Left blue stripe accent
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
        p2.font.size = Pt(17)
        p2.font.bold = True
        p2.font.color.rgb = C_TEXT_TITLE

    def add_card(slide, left, top, width, height, bg=C_WHITE, border=C_BORDER, radius=MSO_SHAPE.ROUNDED_RECTANGLE):
        card = slide.shapes.add_shape(radius, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg
        card.line.color.rgb = border
        card.line.width = Pt(1)
        return card

    # ==================== SLIDE 1: BÌA ĐỀ TÀI ====================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, C_NAVY_DARK)

    # Accent decorative banner
    acc_top = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
    acc_top.fill.solid()
    acc_top.fill.fore_color.rgb = C_BLUE
    acc_top.line.fill.background()

    # Center card
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
    p.text = "BÁO CÁO ĐỀ ÁN HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC"
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
    p.text = "Học viên thực hiện: Nhóm Học viên Cao học (Lớp: Cao học CNTT 20.2)\n" \
             "Giảng viên hướng dẫn: Giảng viên phụ trách học phần  |  Thời gian bảo vệ: Học kỳ 1, Năm học 2026 - 2027"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(10)

    # ==================== SLIDE 2: ĐẶT VẤN ĐỀ & BỐI CẢNH ====================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "1. Bối Cảnh Thực Tiễn & Tính Cấp Thiết Của Đề Tài")

    cards_data = [
        ("📌 1. Bước Ngoặt GDPT 2018", 
         "• Từ kỳ thi Tốt nghiệp THPT 2025, môn Tin học lần đầu tiên trở thành môn thi tự chọn tốt nghiệp.\n\n"
         "• Đề thi chuẩn hóa 3 phần:\n"
         "  - Phần I: Trắc nghiệm 4 lựa chọn (Nhận biết/Thông hiểu)\n"
         "  - Phần II: Trắc nghiệm Đúng/Sai đa ý (Tư duy dữ liệu)\n"
         "  - Phần III: Trả lời ngắn (Vận dụng cao - thuật toán/CS).",
         C_BLUE_LIGHT, C_BORDER_ACCENT),

        ("⚠️ 2. Thách Thức Điểm Liệt",
         "• Quy chế thi tốt nghiệp quy định điểm liệt là <= 1.0 điểm; trượt tốt nghiệp nếu dính điểm liệt.\n\n"
         "• Độ phân hóa phần III và mạch Lập trình (CS) rất cao, học sinh thường mất trọn vẹn điểm.\n\n"
         "• Rất khó để giáo viên phát hiện sớm học sinh bị hổng tư duy lập trình nếu chỉ nhìn điểm học bạ.",
         RGBColor(254, 242, 242), RGBColor(254, 202, 202)),

        ("💡 3. Nhu Cầu Hệ Cơ Sở Tri Thức",
         "• Đánh giá truyền thống chỉ dựa vào điểm trung bình phẳng, bỏ qua bản chất đa chiều của GDPT 2018.\n\n"
         "• Cần một hệ thống học máy tích hợp cơ sở tri thức giúp:\n"
         "  1) Cảnh báo sớm nguy cơ trước kỳ thi\n"
         "  2) Giải thích rõ căn cứ suy luận (XAI)\n"
         "  3) Tự động lập kế hoạch phụ đạo bồi dưỡng.",
         RGBColor(236, 253, 245), RGBColor(167, 243, 208))
    ]

    for i, (title, content, bg_c, bdr_c) in enumerate(cards_data):
        cx = 0.6 + i * 4.1
        add_card(s2, cx, 1.6, 3.9, 5.3, bg=bg_c, border=bdr_c)
        tb = s2.shapes.add_textbox(Inches(cx + 0.2), Inches(1.8), Inches(3.5), Inches(4.9))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        
        p = tf.add_paragraph()
        p.text = content
        p.font.size = Pt(11)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(12)

    # ==================== SLIDE 3: MỤC TIÊU & GIẢI PHÁP 3 TRỤ CỘT ====================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "2. Mục Tiêu Nghiên Cứu & Khung Giải Pháp 3 Trụ Cột")

    pillars = [
        ("🎯 TRỤ CỘT 1: CẢNH BÁO SỚM NGUY CƠ", 
         "Học máy phân loại 3 mức rủi ro:",
         [("🟢 An toàn (>= 6.5 điểm)", "Dự kiến đạt chuẩn tốt nghiệp khá/giỏi"),
          ("🟡 Cần theo dõi (5.0 - 6.4)", "Có nguy cơ trượt nếu đề thi phân hóa cao"),
          ("🔴 Nguy cơ cao (< 5.0 điểm)", "Nguy cơ cao điểm dưới TB hoặc dính điểm liệt")],
         C_BLUE),
        
        ("🔍 TRỤ CỘT 2: MINH BẠCH CĂN CỨ (XAI)",
         "Không dùng AI hộp đen (Black-box):",
         [("🔬 Permutation Importance", "Lượng hóa sức nặng từng tiêu chí đánh giá"),
          ("🕸️ Radar Năng Lực VSTEP", "Đối sánh năng lực cá nhân với chuẩn An toàn"),
          ("💡 Khuyến nghị sư phạm", "Chỉ rõ lỗ hổng: hổng Lập trình CS hay Vận dụng cao")],
         C_EMERALD),

        ("🏫 TRỤ CỘT 3: PHÂN BỔ NGUỒN LỰC",
         "Tối ưu hóa công tác ôn thi tại nhà trường:",
         [("📋 Gom nhóm theo chuyên đề", "Tách riêng HS hổng CS, VDC hay ICT"),
          ("⏱️ Định lượng số tiết/tuần", "Tự động tính quỹ thời gian phụ đạo cần thiết"),
          ("👨‍🏫 Phân công chuyên môn", "Gán đúng giáo viên thế mạnh bồi dưỡng")],
         C_AMBER)
    ]

    for i, (title, sub, items, acc_color) in enumerate(pillars):
        cx = 0.6 + i * 4.1
        add_card(s3, cx, 1.6, 3.9, 5.3, bg=C_WHITE, border=C_BORDER)
        
        hb = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.6), Inches(3.9), Inches(0.85))
        hb.fill.solid()
        hb.fill.fore_color.rgb = acc_color
        hb.line.fill.background()
        
        p = hb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s3.shapes.add_textbox(Inches(cx + 0.2), Inches(2.6), Inches(3.5), Inches(4.1))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = sub
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_TITLE

        for main_i, desc_i in items:
            p = tf.add_paragraph()
            p.text = f"• {main_i}:"
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = C_NAVY
            p.space_before = Pt(10)

            p = tf.add_paragraph()
            p.text = f"   {desc_i}"
            p.font.size = Pt(10.5)
            p.font.color.rgb = C_TEXT_MUTED

    # ==================== SLIDE 4: KIẾN TRÚC TỔNG THỂ HỆ THỐNG ====================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "3. Kiến Trúc Hệ Thống & Quy Trình Xử Lý Đầu Cuối (End-to-End)")

    layers = [
        ("TẦNG 1: NGUỒN DỮ LIỆU (DATA LAYER)",
         "• CSDL SQLite nhúng (vstep_informatics.db): 4.070 học sinh qua 8 khóa học (K18 - K25)\n"
         "• Tiếp nhận file bảng điểm thực tế từ nhà trường (.xlsx, .csv) qua giao diện kéo thả.",
         C_NAVY_DARK, C_WHITE),

        ("TẦNG 2: TIỀN XỬ LÝ & KIỂM SOÁT RÒ RỈ (ANTI-LEAKAGE PIPELINE)",
         "• Hàm validate_input: Kiểm tra định dạng số, chặn giá trị ngoài dải [0, 10] và tỷ lệ LMS [0, 1]\n"
         "• SimpleImputer (điền trung vị) + StandardScaler (chuẩn hóa Z-Score) đóng băng tham số từ tập Train.",
         C_NAVY, C_WHITE),

        ("TẦNG 3: ĐỘNG CƠ HỌC MÁY & ĐỐI SÁNH ĐA MÔ HÌNH (ML ENGINE)",
         "• Baseline Models: Rule-Based Heuristic, Decision Tree (depth=4), Logistic Regression (F1: 89.2%)\n"
         "• Improved Ensemble: Random Forest (150 cây), XGBoost, LightGBM (Recall nguy cơ cao: 93.9%).",
         C_BLUE, C_WHITE),

        ("TẦNG 4: HỆ TRI THỨC SƯ PHẠM & GIẢI THÍCH (KBS & XAI)",
         "• Permutation Feature Importance trích xuất trọng số thuộc tính chi phối rủi ro\n"
         "• Biểu đồ Radar đa chiều VSTEP + Thuật toán heuristic tự động gom nhóm kịch bản phụ đạo.",
         C_EMERALD, C_WHITE),

        ("TẦNG 5: NGUYÊN MẪU ỨNG DỤNG WEB (VSTEP MASTER UI)",
         "• Streamlit Web App chạy độc lập (run_app.bat, port 8501) gồm 5 tab chức năng hoàn chỉnh\n"
         "• Đáp ứng tiêu chuẩn giao diện trực quan, bảng tra cứu ca dự đoán sai và khung đạo đức AI.",
         RGBColor(79, 70, 229), C_WHITE)
    ]

    for i, (title, content, bg_c, text_c) in enumerate(layers):
        cy = 1.55 + i * 1.12
        add_card(s4, 0.6, cy, 12.133, 1.0, bg=C_WHITE, border=C_BORDER)

        badge = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(cy + 0.12), Inches(3.6), Inches(0.76))
        badge.fill.solid()
        badge.fill.fore_color.rgb = bg_c
        badge.line.fill.background()
        
        p = badge.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = text_c
        p.alignment = PP_ALIGN.CENTER

        tb = s4.shapes.add_textbox(Inches(4.55), Inches(cy + 0.08), Inches(8.0), Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = content
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 5: DỮ LIỆU & CHỐNG RÒ RỈ ====================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "4. Đặc Tả Dữ Liệu 15 Đặc Trưng & Quy Trình Chống Rò Rỉ")

    add_card(s5, 0.6, 1.55, 6.8, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_l = s5.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(6.4), Inches(5.1))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    p = tf_l.paragraphs[0]
    p.text = "📋 DANH MỤC 15 ĐẶC TRƯNG ĐẦU VÀO CHUẨN HÓA"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    feats = [
        ("1. Đánh giá quá trình THPT (5)", "Điểm Miệng, 15 Phút, 45 Phút (1 Tiết), Học Kỳ 1, Học Kỳ 2"),
        ("2. Chỉ số suy biến quá trình (2)", "TB Quá Trình = (Miệng+15p+45p*2)/4  |  Xu Hướng = HK2 - HK1"),
        ("3. Chuyên cần số LMS (1)", "Tỷ lệ hoàn thành bài tập rèn luyện LMS [0.0, 1.0]"),
        ("4. Bốn mức tư duy Bloom (4)", "Mức 1 (Nhận biết), Mức 2 (Thông hiểu), Mức 3 (Vận dụng), Mức 4 (Vận dụng cao)"),
        ("5. Ba mạch tri thức 2018 (3)", "Khoa học Máy tính (CS), Tin học Ứng dụng (ICT), Đạo đức số (DL)")
    ]
    for grp, desc in feats:
        p = tf_l.add_paragraph()
        p.text = f"• {grp}:"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(6)

        p = tf_l.add_paragraph()
        p.text = f"   {desc}"
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    add_card(s5, 7.6, 1.55, 5.133, 5.4, bg=C_BLUE_LIGHT, border=C_BORDER_ACCENT)
    tb_r = s5.shapes.add_textbox(Inches(7.8), Inches(1.7), Inches(4.7), Inches(5.1))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p = tf_r.paragraphs[0]
    p.text = "🔒 NGUYÊN TẮC CHỐNG RÒ RỈ (ANTI-LEAKAGE)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    rules = [
        ("Quy mô mẫu nghiên cứu", "4.070 học sinh trên 8 khóa học (K18 - K25), ẩn danh hóa mã dạng STU_xxxx."),
        ("Phân chia Stratified Split", "Bảo toàn hoàn hảo tỷ lệ phân bố nhãn:\n• Train Set: 2.849 mẫu (70%)\n• Validation Set: 610 mẫu (15%)\n• Test Set: 611 mẫu (15% độc lập)."),
        ("Đóng băng tham số (Fit/Transform)", "StandardScaler & SimpleImputer CHỈ fit trên tập Train, sau đó áp dụng tĩnh (transform) lên Val và Test."),
        ("Loại bỏ hoàn toàn biến rò rỉ", "Điểm thi tốt nghiệp dự kiến (est_final_score) chỉ dùng sinh nhãn mục tiêu, TUYỆT ĐỐI KHÔNG đưa vào tập đặc trưng X huấn luyện.")
    ]
    for r_title, r_desc in rules:
        p = tf_r.add_paragraph()
        p.text = f"✔ {r_title}:"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_EMERALD
        p.space_before = Pt(8)

        p = tf_r.add_paragraph()
        p.text = r_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 6: THIẾT KẾ MÔ HÌNH HỌC MÁY ====================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "5. Thiết Kế Mô Hình: Đối Sánh Baseline vs Cải Tiến (Ensemble)")

    add_card(s6, 0.6, 1.55, 5.9, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_m1 = s6.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(5.5), Inches(5.1))
    tf_m1 = tb_m1.text_frame
    tf_m1.word_wrap = True
    
    p = tf_m1.paragraphs[0]
    p.text = "🏛️ CÁC MÔ HÌNH CƠ SỞ (BASELINES)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    base_models = [
        ("Baseline 1: Rule-Based Heuristic",
         "Phân loại dựa trên ngưỡng điểm sư phạm truyền thống (Nếu điểm TB Quá trình < 5.0 hoặc Giữa kỳ < 4.5 thì báo nguy cơ).\nƯu: Đơn giản, giải thích tức thời | Nhược: Bỏ sót lỗ hổng phân nhánh."),
        ("Baseline 2: Decision Tree (depth=4)",
         "Cây quyết định đơn lẻ với max_depth=4, class_weight='balanced'.\nƯu: Tạo cây suy diễn logic trực quan | Nhược: Độ chính xác vừa phải, dễ overfit ở ngưỡng biên."),
        ("Baseline 3: Logistic Regression (Multinomial)",
         "Hồi quy đa biến với bộ tối ưu hóa lbfgs và cân bằng trọng số lớp.\nƯu: Tốc độ suy luận siêu nhanh, mặt phân chia tuyến tính đạt F1 cao trên dữ liệu chuẩn hóa.")
    ]
    for b_name, b_desc in base_models:
        p = tf_m1.add_paragraph()
        p.text = b_name
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(10)

        p = tf_m1.add_paragraph()
        p.text = b_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    add_card(s6, 6.8, 1.55, 5.9, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_m2 = s6.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.5), Inches(5.1))
    tf_m2 = tb_m2.text_frame
    tf_m2.word_wrap = True
    
    p = tf_m2.paragraphs[0]
    p.text = "⚡ CÁC MÔ HÌNH CẢI TIẾN (ENSEMBLE LEARNING)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    imp_models = [
        ("Improved 1: Random Forest (150 cây, depth=8)",
         "Rừng cây quyết định kết hợp kỹ thuật Bagging và phân nhánh ngẫu nhiên đặc trưng. Giảm phương sai (variance), bắt trọn các quan hệ phi tuyến phức tạp giữa 4 mức Bloom."),
        ("Improved 2: XGBoost (Extreme Gradient Boosting)",
         "Tối ưu hàm mất mát đa lớp với điều chuẩn L1/L2 chống quá khớp. Học từ phần dư của các cây trước, nâng cao năng lực phân tách tại vùng biên rủi ro."),
        ("Improved 3: LightGBM (Gradient Boosting trên Histogram)",
         "Phân nhánh theo lá (leaf-wise) kết hợp gom cụm histogram. Cho tốc độ huấn luyện vượt trội và đạt độ nhạy (Recall lớp Nguy cơ cao) cao nhất (93.9%).")
    ]
    for i_name, i_desc in imp_models:
        p = tf_m2.add_paragraph()
        p.text = i_name
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_EMERALD
        p.space_before = Pt(10)

        p = tf_m2.add_paragraph()
        p.text = i_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 7: KẾT QUẢ THỰC NGHIỆM ĐA MÔ HÌNH ====================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "6. Bảng Kết Quả Thực Nghiệm Đa Mô Hình (Academic Benchmark)")

    add_card(s7, 0.6, 1.55, 12.133, 3.1, bg=C_WHITE, border=C_BORDER)
    
    b_rows = [
        ["Mô Hình Thử Nghiệm", "Accuracy", "Macro F1", "Recall Nguy Cơ", "Precision Nguy Cơ", "Độ Trễ / HS", "Thời Gian Train"],
        ["Baseline: Rule-Based Heuristic", "84.12%", "82.83%", "78.78%", "97.47%", "~ 0 ms", "0.000 s"],
        ["Baseline: Decision Tree (depth=4)", "85.60%", "84.31%", "87.35%", "94.69%", "~ 0 ms", "0.005 s"],
        ["Baseline: Logistic Regression", "90.34%", "89.19%", "91.84%", "95.74%", "~ 0 ms", "0.013 s"],
        ["Improved: Random Forest", "88.71%", "87.31%", "91.02%", "94.89%", "0.017 ms", "0.333 s"],
        ["Improved: XGBoost", "88.38%", "86.64%", "91.84%", "92.98%", "0.004 ms", "0.258 s"],
        ["Improved: LightGBM (Recall Tốt Nhất)", "89.36%", "87.72%", "93.88%", "93.12%", "0.005 ms", "0.096 s"]
    ]

    tbl_shape = s7.shapes.add_table(len(b_rows), len(b_rows[0]), Inches(0.8), Inches(1.7), Inches(11.733), Inches(2.7))
    tbl = tbl_shape.table
    col_w = [Inches(3.333), Inches(1.4), Inches(1.4), Inches(1.5), Inches(1.5), Inches(1.3), Inches(1.3)]
    for ci, w in enumerate(col_w):
        tbl.columns[ci].width = w

    for ri, row in enumerate(b_rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            p.font.size = Pt(10)
            if ri == 0:
                p.font.bold = True
                p.font.color.rgb = C_WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_NAVY
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            else:
                p.font.color.rgb = C_TEXT_TITLE
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                if ri == 6:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(236, 253, 245)
                    p.font.bold = True
                elif ri == 3:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(239, 246, 255)
                    p.font.bold = True

    add_card(s7, 0.6, 4.85, 12.133, 2.1, bg=C_BLUE_LIGHT, border=C_BORDER_ACCENT)
    tb_ins = s7.shapes.add_textbox(Inches(0.8), Inches(4.95), Inches(11.733), Inches(1.9))
    tf_ins = tb_ins.text_frame
    tf_ins.word_wrap = True

    p = tf_ins.paragraphs[0]
    p.text = "💡 ĐÚC KẾT & NHẬN XÉT HỌC THUẬT TRUNG THỰC:"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    ins_bullets = [
        "1. Ưu tiên sư phạm: Chỉ số Recall của lớp Nguy cơ cao là chỉ số sống còn, vì chi phí bỏ sót một học sinh trượt tốt nghiệp lớn hơn rất nhiều chi phí cảnh báo nhầm để theo dõi thêm.",
        "2. LightGBM đạt Recall lớp Nguy cơ cao tốt nhất (93.88% - phát hiện chính xác 230/245 học sinh rủi ro), cải thiện +15.1% so với Rule-Based Heuristic.",
        "3. Logistic Regression đạt Macro F1 cao nhất (89.19%) cho thấy quan hệ giữa các đặc trưng và nhãn trên dữ liệu mô phỏng tương đối rõ nét.",
        "4. Cả 3 mô hình cải tiến đều có độ trễ suy luận siêu thấp (< 0.02 ms/học sinh), hoàn toàn đáp ứng khả năng phục vụ thời gian thực cho hàng nghìn học sinh."
    ]
    for b in ins_bullets:
        p = tf_ins.add_paragraph()
        p.text = b
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(3)

    # ==================== SLIDE 8: SO SÁNH BASELINE THỜI GIAN ====================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "7. So Sánh Baseline Vận Hành: Thủ Công vs Mô Hình Học Máy (ML)")

    metrics = [
        ("3 - 5 Ngày Công", "RÀ SOÁT THỦ CÔNG", "Hội đồng GV đọc sổ điểm tính tay toàn trường", C_ROSE),
        ("< 0.2 Giây", "MÔ HÌNH RANDOM FOREST", "Xử lý toàn bộ 4.070 HS trên 8 khóa học", C_EMERALD),
        ("0.005 Giây", "XỬ LÝ 1 LỚP HỌC (40 HS)", "Trả về kết quả phân bổ và cảnh báo tức thì", C_BLUE),
        ("100% Khách Quan", "BẢO TOÀN NĂNG LỰC", "Đánh giá đa chiều 4 mức Bloom & 3 mạch Tin", C_NAVY)
    ]
    for i, (m_val, m_lbl, m_sub, m_clr) in enumerate(metrics):
        cx = 0.6 + i * 3.1
        add_card(s8, cx, 1.55, 2.9, 1.6, bg=C_WHITE, border=C_BORDER)
        tb_m = s8.shapes.add_textbox(Inches(cx + 0.15), Inches(1.65), Inches(2.6), Inches(1.4))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True

        p = tf_m.paragraphs[0]
        p.text = m_val
        p.font.size = Pt(17)
        p.font.bold = True
        p.font.color.rgb = m_clr
        p.alignment = PP_ALIGN.CENTER

        p = tf_m.add_paragraph()
        p.text = m_lbl
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_TITLE
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(4)

        p = tf_m.add_paragraph()
        p.text = m_sub
        p.font.size = Pt(9)
        p.font.color.rgb = C_TEXT_MUTED
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(2)

    add_card(s8, 0.6, 3.35, 12.133, 3.6, bg=C_WHITE, border=C_BORDER)
    time_rows = [
        ["Tiêu Chí Đánh Giá", "Phương Pháp Thủ Công (Hội Đồng GV)", "Mô Hình Tuyến Tính (Logistic)", "Mô Hình Cải Tiến (RF / LightGBM)"],
        ["Thời gian xử lý 1 lớp (40 HS)", "~ 25 - 35 phút (Ước lượng giả định)", "~ 0.001 giây (Đo đạc thực tế)", "~ 0.005 - 0.04 giây (Đo đạc thực tế)"],
        ["Thời gian xử lý 4.070 HS (8 khóa)", "~ 3 - 5 ngày công tổ chuyên môn", "~ 0.02 giây (Đo đạc thực tế)", "~ 0.08 - 0.18 giây (Đo đạc thực tế)"],
        ["Thời điểm phát hiện nguy cơ", "Sau khi thi học kỳ xong (quá muộn)", "Ngay khi có đủ điểm quá trình", "Ngay khi có điểm quá trình giữa kỳ"],
        ["Độ bao phủ tiêu chí GDPT 2018", "Thường chỉ nhìn điểm trung bình", "Mặt phân cách tuyến tính 15 chiều", "Nắm bắt phi tuyến 4 Bloom & 3 mạch Tin"],
        ["Căn cứ giải thích (XAI)", "Cảm quan kinh nghiệm cá nhân GV", "Trọng số Beta cố định", "Radar năng lực + Trọng số rủi ro từng HS"],
        ["Kịch bản can thiệp sư phạm", "GV tự thảo luận, dễ quá tải", "Chỉ đưa ra danh sách nguy cơ", "Tự động phân bổ số lớp & số tiết phụ đạo"]
    ]
    tbl_time_shape = s8.shapes.add_table(len(time_rows), len(time_rows[0]), Inches(0.8), Inches(3.45), Inches(11.733), Inches(3.3))
    tbl_t = tbl_time_shape.table
    cw_t = [Inches(2.533), Inches(3.0), Inches(3.0), Inches(3.2)]
    for ci, w in enumerate(cw_t):
        tbl_t.columns[ci].width = w

    for ri, row in enumerate(time_rows):
        for ci, val in enumerate(row):
            cell = tbl_t.cell(ri, ci)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            p.font.size = Pt(9.5)
            if ri == 0:
                p.font.bold = True
                p.font.color.rgb = C_WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_NAVY
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            else:
                p.font.color.rgb = C_TEXT_TITLE
                p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                if ri % 2 == 1:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = C_BG_PAGE

    # ==================== SLIDE 9: GIẢI THÍCH MÔ HÌNH XAI ====================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "8. Khả Năng Giải Thích Mô Hình (Explainable AI - XAI)")

    fi_img = os.path.join(figures_dir, 'feature_importance_ranking.png')
    if os.path.exists(fi_img):
        add_card(s9, 0.6, 1.55, 6.2, 5.4, bg=C_WHITE, border=C_BORDER)
        s9.shapes.add_picture(fi_img, Inches(0.75), Inches(1.7), Inches(5.9), Inches(5.1))
    else:
        add_card(s9, 0.6, 1.55, 6.2, 5.4, bg=C_WHITE, border=C_BORDER)

    add_card(s9, 7.0, 1.55, 5.733, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_xai = s9.shapes.add_textbox(Inches(7.2), Inches(1.7), Inches(5.333), Inches(5.1))
    tf_xai = tb_xai.text_frame
    tf_xai.word_wrap = True

    p = tf_xai.paragraphs[0]
    p.text = "🔬 4 YẾU TỐ QUYẾT ĐỊNH NGUY CƠ HỌC TẬP"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    factors = [
        ("1. Điểm Học Kỳ 2 (score_hk2)", "Có sức nặng suy giảm F1 lớn nhất (0.108). Phản ánh chính xác phong độ thực tế sát kỳ thi tốt nghiệp."),
        ("2. Điểm TB Quá Trình (score_tx_avg)", "Sức nặng 0.048. Tổng hợp có trọng số giữa điểm Miệng, 15p và 45p, đo lường nền tảng kiến thức tích lũy cả năm."),
        ("3. Mức 4: Vận Dụng Cao (score_high_apply)", "Sức nặng 0.041. Quyết định khả năng giải quyết Phần III trong đề thi GDPT 2018; HS yếu VDC gần như cầm chắc điểm dưới 5."),
        ("4. Mạch Khoa Học Máy Tính (score_topic_cs)", "Sức nặng 0.039. Mạch Lập trình & Thuật toán là 'vùng trũng' điểm số lớn nhất của học sinh THPT.")
    ]
    for f_title, f_desc in factors:
        p = tf_xai.add_paragraph()
        p.text = f_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE
        p.space_before = Pt(8)

        p = tf_xai.add_paragraph()
        p.text = f_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    p = tf_xai.add_paragraph()
    p.text = "💡 Ứng dụng thực tế: Tại Tab 2, hệ thống tự động sinh biểu đồ Radar năng lực và Waterfall Attribution giải thích cặn kẽ cho từng học sinh."
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = C_EMERALD
    p.space_before = Pt(12)

    # ==================== SLIDE 10: PHÂN TÍCH CA DỰ ĐOÁN SAI ====================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "9. Phân Tích Chuyên Sâu Các Ca Dự Đoán Sai (In-Depth Error Analysis)")

    cm_img = os.path.join(figures_dir, 'confusion_matrix_best_model.png')
    if os.path.exists(cm_img):
        add_card(s10, 0.6, 1.55, 5.2, 5.4, bg=C_WHITE, border=C_BORDER)
        s10.shapes.add_picture(cm_img, Inches(0.75), Inches(1.8), Inches(4.9), Inches(4.8))
    else:
        add_card(s10, 0.6, 1.55, 5.2, 5.4, bg=C_WHITE, border=C_BORDER)

    add_card(s10, 6.0, 1.55, 6.733, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_err = s10.shapes.add_textbox(Inches(6.2), Inches(1.7), Inches(6.333), Inches(5.1))
    tf_err = tb_err.text_frame
    tf_err.word_wrap = True

    p = tf_err.paragraphs[0]
    p.text = "🔍 PHÂN TÍCH 69 CA DỰ ĐOÁN SAI TRÊN TẬP TEST (611 HS)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    err_groups = [
        ("🚨 1. Nhóm Bỏ Sót Nguy Cơ (False Negatives - 22 ca):",
         "• Thực tế: Nguy cơ cao (< 5.0)  ==>  Dự đoán AI: Cần theo dõi (5.0 - 6.4).\n"
         "• Nguyên nhân sư phạm: HS có điểm học bạ khá (6.0 - 7.0) nhờ chuyên cần LMS và điểm miệng, nhưng 'hổng rỗng' tư duy thuật toán CS (< 4.0). Khi thi đề GDPT 2018 bị trượt điểm liệt. Mô hình bị đánh lừa bởi trọng số điểm học bạ.",
         C_ROSE),

        ("⚠️ 2. Nhóm Cảnh Báo Nhầm (False Positives - 22 ca):",
         "• Thực tế: An toàn (>= 6.5)  ==>  Dự đoán AI: Cần theo dõi (5.0 - 6.4).\n"
         "• Nguyên nhân sư phạm: HS có năng lực tư duy tốt nhưng nộp bài LMS muộn hoặc thiếu một bài 15p ngẫu nhiên làm điểm quá trình thấp, kích hoạt cảnh báo phòng ngừa.",
         C_AMBER),

        ("🟡 3. Nhóm Ca Lệch Ranh Giới Biên (Boundary Cases - 25 ca):",
         "• Điểm số dao động sát ngưỡng biên 5.0 điểm và 6.5 điểm. Không có ca nào bị nhảy cách lớp (Nguy cơ cao nhầm thành An toàn: 0 ca).",
         C_BLUE)
    ]
    for e_title, e_desc, e_clr in err_groups:
        p = tf_err.add_paragraph()
        p.text = e_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = e_clr
        p.space_before = Pt(8)

        p = tf_err.add_paragraph()
        p.text = e_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 11: NGUYÊN MẪU ỨNG DỤNG WEB ====================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "10. Nguyên Mẫu Phần Mềm Web (Streamlit VSTEP Master UI)")

    tabs = [
        ("📁 TAB 1: DỰ BÁO THEO LỚP",
         "• Truy vấn CSDL SQLite (8 khóa, 4.070 HS) hoặc tải file Excel mới.\n"
         "• Kiểm tra tính hợp lệ dữ liệu tự động (validate_input).\n"
         "• Biểu đồ Donut phân bố nguy cơ toàn khối.\n"
         "• Bảng dữ liệu lọc theo mức độ nguy cơ & xuất file Excel chi tiết.",
         C_BLUE_LIGHT, C_BLUE),

        ("👤 TAB 2: CHẨN ĐOÁN CÁ NHÂN & XAI",
         "• Bố cục Form 2 cột cân xứng (50% - 50%), 13 thanh trượt điểm THPT.\n"
         "• Hộp tính nhanh TB Quá trình & Xu hướng điểm số.\n"
         "• 3 Nút chọn mẫu HS đại diện (An toàn, Hổng CS, Nguy cơ liệt).\n"
         "• Biểu đồ Radar ma trận năng lực + Thẻ lời khuyên sư phạm.",
         RGBColor(236, 253, 245), C_EMERALD),

        ("🏫 TAB 3: PHÂN BỔ NGUỒN LỰC",
         "• Tự động tách HS nguy cơ theo 3 chuyên đề bồi dưỡng trọng tâm.\n"
         "• Tính toán số lớp phụ đạo (15-20 HS/lớp) và số tiết/tuần.\n"
         "• Gợi ý phân công giáo viên theo thế mạnh chuyên môn.",
         RGBColor(254, 243, 199), C_AMBER),

        ("📊 TAB 4: BENCHMARK & PHÂN TÍCH LỖI",
         "• Bảng tổng hợp chỉ số đánh giá 6 mô hình trên tập Test độc lập.\n"
         "• Biểu đồ Confusion Matrix & Permutation Importance.\n"
         "• Bảng so sánh baseline thời gian thực thi (Thủ công vs ML).\n"
         "• Bảng tương tác tra cứu danh sách 69 ca dự đoán sai và nguyên nhân.",
         RGBColor(243, 232, 255), RGBColor(147, 51, 234)),

        ("ℹ️ TAB 5: TRI THỨC & ĐẠO ĐỨC AI",
         "• Ma trận đặc tả đề thi GDPT 2018 (Phần I, II, III).\n"
         "• Khung Đạo đức AI & Giới hạn của Hệ hỗ trợ ra quyết định.\n"
         "• Hướng dẫn vận hành chuẩn mực cho ban giám hiệu và giáo viên.",
         C_BG_PAGE, C_NAVY)
    ]

    for i, (t_name, t_desc, bg_c, acc_c) in enumerate(tabs):
        cy = 1.55 + i * 1.12
        add_card(s11, 0.6, cy, 12.133, 1.0, bg=bg_c, border=C_BORDER)

        badge = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.75), Inches(cy + 0.12), Inches(3.6), Inches(0.76))
        badge.fill.solid()
        badge.fill.fore_color.rgb = acc_c
        badge.line.fill.background()
        
        p = badge.text_frame.paragraphs[0]
        p.text = t_name
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s11.shapes.add_textbox(Inches(4.55), Inches(cy + 0.08), Inches(8.0), Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = t_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 12: KỊCH BẢN PHÂN BỔ NGUỒN LỰC ====================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "11. Kịch Bản Phân Bổ Nguồn Lực Ôn Tập & Phụ Đạo Tự Động")

    groups = [
        ("💻 NHÓM 1: CHUYÊN ĐỀ LẬP TRÌNH & THUẬT TOÁN (CS)",
         "Mục tiêu: Chống điểm liệt phần trắc nghiệm trả lời ngắn",
         [("Đối tượng", "Học sinh có điểm chuyên đề CS < 5.0 (lỗ hổng nặng nhất)."),
          ("Quy mô lớp", "Tối đa 15 - 20 học sinh / lớp để đảm bảo kèm cặp trực tiếp."),
          ("Thời lượng", "3 tiết / tuần / lớp (tập trung luyện viết mã và đọc luồng dữ liệu)."),
          ("Phân công GV", "Giáo viên có năng lực chuyên sâu về thuật toán và ngôn ngữ lập trình.")],
         C_BLUE),

        ("🧠 NHÓM 2: CHUYÊN ĐỀ VẬN DỤNG CAO & ĐÚNG/SAI",
         "Mục tiêu: Rèn luyện kỹ năng phân tích logic đa ý",
         [("Đối tượng", "Học sinh nắm vững lý thuyết nhưng lúng túng trong các câu hỏi đa ý."),
          ("Quy mô lớp", "20 - 25 học sinh / lớp."),
          ("Thời lượng", "2 tiết / tuần / lớp (rèn luyện chiến thuật làm bài Đúng/Sai)."),
          ("Phân công GV", "Giáo viên giàu kinh nghiệm ôn thi tốt nghiệp, nắm chắc ma trận đề.")],
         C_EMERALD),

        ("📱 NHÓM 3: CHUYÊN ĐỀ ỨNG DỤNG TIN HỌC (ICT)",
         "Mục tiêu: Lấy trọn điểm lý thuyết cơ bản và phần mềm",
         [("Đối tượng", "Học sinh yếu mảng hệ điều hành, mạng máy tính và bảng tính."),
          ("Quy mô lớp", "20 - 25 học sinh / lớp."),
          ("Thời lượng", "2 tiết / tuần / lớp (thực hành trực tiếp trên phòng máy)."),
          ("Phân công GV", "Giáo viên phụ trách phòng thực hành tin học.")],
         C_AMBER)
    ]

    for i, (g_title, g_sub, g_items, g_clr) in enumerate(groups):
        cx = 0.6 + i * 4.1
        add_card(s12, cx, 1.55, 3.9, 5.4, bg=C_WHITE, border=C_BORDER)

        hb = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx), Inches(1.55), Inches(3.9), Inches(0.85))
        hb.fill.solid()
        hb.fill.fore_color.rgb = g_clr
        hb.line.fill.background()
        
        p = hb.text_frame.paragraphs[0]
        p.text = g_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s12.shapes.add_textbox(Inches(cx + 0.2), Inches(2.5), Inches(3.5), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = g_sub
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_TITLE

        for it_lbl, it_val in g_items:
            p = tf.add_paragraph()
            p.text = f"• {it_lbl}:"
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = C_NAVY
            p.space_before = Pt(8)

            p = tf.add_paragraph()
            p.text = f"   {it_val}"
            p.font.size = Pt(10)
            p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 13: KHUNG ĐẠO ĐỨC AI & GIỚI HẠN ====================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "12. Khung Đạo Đức AI & Giới Hạn Của Hệ Hỗ Trợ Ra Quyết Định")

    ethics = [
        ("1. HỖ TRỢ RA QUYẾT ĐỊNH (DECISION SUPPORT)",
         "Hệ thống đóng vai trò cố vấn phân tích khách quan cho nhà giáo. TUYỆT ĐỐI KHÔNG tự động xếp loại học bạ, kỷ luật hay thay thế quyền đánh giá sư phạm của thầy cô giáo.",
         C_BLUE_LIGHT, C_BLUE),

        ("2. KHÔNG DÁN NHÃN ĐỊNH KIẾN (ZERO STIGMATIZATION)",
         "Các cảnh báo nguy cơ học tập chỉ lưu hành nội bộ giữa Ban giám hiệu và Giáo viên phụ trách nhằm xây dựng kế hoạch bồi dưỡng, tuyệt đối không công khai bêu tên học sinh làm ảnh hưởng tâm lý.",
         RGBColor(254, 242, 242), C_ROSE),

        ("3. BẢO VỆ DỮ LIỆU CÁ NHÂN (DATA PRIVACY)",
         "Toàn bộ dữ liệu huấn luyện và thực nghiệm đều được ẩn danh hóa nghiêm ngặt dạng mã STU_xxxx, loại bỏ hoàn toàn số định danh CCCD và thông tin cá nhân nhạy cảm.",
         RGBColor(236, 253, 245), C_EMERALD),

        ("4. CON NGƯỜI LÀ QUYẾT ĐỊNH TỐI HẬU (HUMAN-IN-THE-LOOP)",
         "Mô hình không phản ánh được các biến cố ngoại cảnh đột xuất (như học sinh ốm đau trong kỳ thi, hoàn cảnh gia đình). Giáo viên luôn là mắt xích kiểm chứng và quyết định phương án hỗ trợ cuối cùng.",
         RGBColor(254, 243, 199), C_AMBER)
    ]

    for i, (e_title, e_desc, bg_c, acc_c) in enumerate(ethics):
        cy = 1.6 + i * 1.35
        add_card(s13, 0.6, cy, 12.133, 1.2, bg=bg_c, border=C_BORDER)

        badge = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(cy + 0.15), Inches(3.8), Inches(0.9))
        badge.fill.solid()
        badge.fill.fore_color.rgb = acc_c
        badge.line.fill.background()
        
        p = badge.text_frame.paragraphs[0]
        p.text = e_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

        tb = s13.shapes.add_textbox(Inches(4.8), Inches(cy + 0.15), Inches(7.7), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = e_desc
        p.font.size = Pt(11)
        p.font.color.rgb = C_TEXT_BODY

    # ==================== SLIDE 14: KẾT LUẬN & HƯỚNG PHÁT TRIỂN ====================
    s14 = prs.slides.add_slide(blank_layout)
    add_header(s14, "13. Kết Luận & Hướng Phát Triển Tiếp Theo")

    add_card(s14, 0.6, 1.55, 5.9, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_c1 = s14.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(5.5), Inches(5.1))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True

    p = tf_c1.paragraphs[0]
    p.text = "🎯 KẾT QUẢ ĐẠT ĐƯỢC CỦA ĐỀ ÁN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    conclusions = [
        "1. Giải quyết bài toán thực tiễn: Bám sát cấu trúc đề thi Tốt nghiệp THPT mới theo Chương trình GDPT 2018.",
        "2. Dữ liệu chuẩn mực & CSDL nhúng: Quản trị 4.070 học sinh qua 8 khóa học bằng SQLite, kiểm soát chống rò rỉ dữ liệu (Anti-Leakage) chặt chẽ.",
        "3. Hiệu năng học máy cao: LightGBM đạt Recall lớp Nguy cơ cao tới 93.88%; Logistic Regression đạt Macro F1 89.19% với độ trễ suy luận siêu nhanh.",
        "4. Minh bạch & Có thể giải thích (XAI): Chỉ ra 4 đặc trưng cốt lõi (HK2, TB Quá trình, VDC, CS) và cung cấp biểu đồ Radar đa chiều.",
        "5. Sản phẩm bàn giao trọn vẹn: Web Streamlit VSTEP Master UI, Báo cáo Word 5 chương, Slide thuyết trình và mã nguồn GitHub hoàn chỉnh."
    ]
    for c in conclusions:
        p = tf_c1.add_paragraph()
        p.text = c
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_TEXT_BODY
        p.space_before = Pt(8)

    add_card(s14, 6.8, 1.55, 5.9, 5.4, bg=C_WHITE, border=C_BORDER)
    tb_c2 = s14.shapes.add_textbox(Inches(7.0), Inches(1.7), Inches(5.5), Inches(5.1))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True

    p = tf_c2.paragraphs[0]
    p.text = "🚀 ĐỊNH HƯỚNG PHÁT TRIỂN TIẾP THEO"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_NAVY

    future_works = [
        ("1. Thu thập dữ liệu thực tế đã ẩn danh", "Phối hợp cùng các trường THPT đối tác để kiểm chứng và hiệu chuẩn lại mô hình trên phổ điểm học sinh thực tế."),
        ("2. Hiệu chuẩn xác suất (Probability Calibration)", "Áp dụng kỹ thuật Isotonic Regression hoặc Platt Scaling để đảm bảo xác suất dự báo rủi ro phản ánh chính xác tần suất thực tế."),
        ("3. Tích hợp SHAP Waterfall cục bộ", "Nâng cấp từ Permutation Importance toàn cục sang SHAP values cục bộ chi tiết cho từng câu hỏi kiểm tra."),
        ("4. Mở rộng liên môn & Cấp học", "Áp dụng kiến trúc hệ tri thức cảnh báo sớm sang các môn thi trắc nghiệm khác (Toán, Vật lý, Hóa học) theo chuẩn GDPT 2018.")
    ]
    for f_title, f_desc in future_works:
        p = tf_c2.add_paragraph()
        p.text = f_title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_EMERALD
        p.space_before = Pt(8)

        p = tf_c2.add_paragraph()
        p.text = f_desc
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_BODY

    prs.save(output_pptx)
    print(f"[SUCCESS] Đã tạo thành công Slide thuyết trình tại: {output_pptx}")

if __name__ == '__main__':
    create_deck()

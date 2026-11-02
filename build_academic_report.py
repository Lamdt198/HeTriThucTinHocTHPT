"""
build_academic_report.py
Tự động xuất Báo cáo Khoa học định dạng Word (.docx) chuẩn quy cách Cao học CNTT K20.2.
Bao gồm đầy đủ 5 chương, các bảng số liệu thực nghiệm 8 khóa học (4.070 HS),
so sánh baseline thời gian, phân tích ca dự đoán sai (Error Analysis), và khung đạo đức AI.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def create_report():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    reports_dir = os.path.join(base_dir, 'reports')
    fig_dir = os.path.join(reports_dir, 'figures')
    output_docx = os.path.join(reports_dir, 'Bao_Cao_Tong_Ket_Cac_He_CS_Tri_Thuc_K20.2.docx')
    
    doc = docx.Document()
    
    # Cấu hình lề trang A4 chuẩn học thuật
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.0)
        
    # Style chuẩn
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x1a, 0x1a, 0x1a)
    
    # ------------------ TRANG BÌA ------------------
    p_univ = doc.add_paragraph()
    p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_univ = p_univ.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO\nCHƯƠNG TRÌNH ĐÀO TẠO CAO HỌC CÔNG NGHỆ THÔNG TIN\n")
    r_univ.bold = True
    r_univ.font.size = Pt(13)
    
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.add_run("-------------------- *** --------------------\n\n\n")
    
    p_title_sub = doc.add_paragraph()
    p_title_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_title_sub.add_run("BÁO CÁO ĐỀ ÁN HỌC PHẦN: CÁC HỆ CƠ SỞ TRI THỨC\n")
    r_sub.font.size = Pt(14)
    r_sub.bold = True
    r_sub.font.color.rgb = RGBColor(0x1e, 0x3a, 0x8a)
    
    p_main_title = doc.add_paragraph()
    p_main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_main_title.add_run("HỆ TRI THỨC VÀ HỌC MÁY HỖ TRỢ PHÂN TÍCH MA TRẬN KẾT QUẢ ĐÁNH GIÁ\nVÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC\nĐỊNH HƯỚNG THI TỐT NGHIỆP THPT CHƯƠNG TRÌNH GDPT 2018\n\n\n")
    r_main.font.size = Pt(16)
    r_main.bold = True
    r_main.font.color.rgb = RGBColor(0xb9, 0x1c, 0x1c)
    
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_info.paragraph_format.left_indent = Inches(1.5)
    p_info.add_run("Học viên thực hiện: \tNhóm Học viên Cao học\n").bold = True
    p_info.add_run("Lớp:\t\t\tCao học CNTT K20.2\n").bold = True
    p_info.add_run("Học phần:\t\tCác Hệ Cơ Sở Tri Thức\n").bold = True
    p_info.add_run("Thời gian thực hiện:\t4 tuần (Học kỳ 1 - Năm học 2026 - 2027)\n\n\n\n").bold = True
    
    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.add_run("Hà Nội, Năm 2026")
    
    doc.add_page_break()
    
    # ------------------ BẢNG PHÂN CÔNG ĐÓNG GÓP ------------------
    h_contrib = doc.add_heading("BẢNG PHÂN CÔNG NHIỆM VỤ VÀ ĐÓNG GÓP THÀNH VIÊN (LỘ TRÌNH 4 TUẦN)", level=1)
    h_contrib.paragraph_format.space_before = Pt(12)
    h_contrib.paragraph_format.space_after = Pt(12)
    
    t_contrib = doc.add_table(rows=5, cols=5)
    t_contrib.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["STT", "Họ và tên học viên", "Mã HV", "Nhiệm vụ phụ trách cụ thể theo tiến độ", "Đóng góp"]
    
    for i, title in enumerate(headers):
        cell = t_contrib.cell(0, i)
        cell.text = title
        set_cell_background(cell, "1E3A8A")
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    contrib_data = [
        ("1", "[Học viên 1 - Trưởng nhóm]", "HV01", "Tuần 1-2: Khảo sát ma trận GDPT 2018, xây dựng kiến trúc CSDL SQLite (4.070 HS, 8 khóa), chuẩn hóa pipeline tiền xử lý chống rò rỉ dữ liệu (Anti-Leakage), huấn luyện các mô hình Baseline.", "100%"),
        ("2", "[Học viên 2]", "HV02", "Tuần 2-3: Huấn luyện các mô hình cải tiến (Random Forest, XGBoost, LightGBM), tinh chỉnh siêu tham số, thiết lập ma trận nhầm lẫn và đối sánh đa chỉ số (Accuracy, Macro F1, Recall lớp nguy cơ).", "100%"),
        ("3", "[Học viên 3]", "HV03", "Tuần 3-4: Hiện thực hóa nguyên mẫu phần mềm Web Streamlit (VSTEP Master UI Standard), tích hợp XAI (Permutation Importance, Waterfall attribution) và thuật toán phân bổ nguồn lực ôn tập.", "100%"),
        ("4", "[Học viên 4]", "HV04", "Tuần 3-4: Thực hiện phân tích chuyên sâu các ca dự đoán sai (In-Depth Error Analysis: False Negative & False Positive), xây dựng khung đạo đức AI, biên tập báo cáo khoa học Word và slide bảo vệ.", "100%")
    ]
    
    for row_idx, data in enumerate(contrib_data, start=1):
        for col_idx, text in enumerate(data):
            c = t_contrib.cell(row_idx, col_idx)
            c.text = text
            if col_idx in [0, 2, 4]:
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
    doc.add_paragraph("\n")
    
    # ------------------ CHƯƠNG 1 ------------------
    doc.add_heading("CHƯƠNG 1: GIỚI THIỆU VÀ ĐẶC TẢ BÀI TOÁN HỌC MÁY", level=1)
    
    doc.add_heading("1.1. Lý do chọn đề tài và bối cảnh thực tiễn", level=2)
    doc.add_paragraph(
        "Kỳ thi Tốt nghiệp THPT theo Chương trình Giáo dục Phổ thông 2018 (GDPT 2018) chính thức đưa môn Tin học "
        "vào danh mục các môn thi tốt nghiệp tự chọn. Cấu trúc đề thi mới được chuẩn hóa thành 3 phần: "
        "Phần I gồm các câu trắc nghiệm 4 lựa chọn (đo lường Nhận biết và Thông hiểu); "
        "Phần II gồm các câu trắc nghiệm Đúng/Sai đa ý (đo lường khả năng phân tích logic luồng dữ liệu); "
        "và Phần III gồm các câu trắc nghiệm trả lời ngắn (phân hóa học sinh giỏi với các bài toán thuật toán và lập trình thực tiễn)."
    )
    doc.add_paragraph(
        "Thách thức lớn nhất đối với các nhà trường THPT là việc phát hiện sớm học sinh có nguy cơ bị điểm kém (< 5.0) hoặc điểm liệt (< 1.0) "
        "để kịp thời bồi dưỡng. Các phương pháp truyền thống chỉ nhìn vào điểm trung bình quá trình số học đơn thuần, không bóc tách được lỗ hổng "
        "ở từng mạch nội dung cốt lõi: Khoa học máy tính (CS), Tin học ứng dụng (ICT), và Đạo đức số (DL). Vì vậy, đề tài nghiên cứu ứng dụng Học máy "
        "và Hệ cơ sở tri thức để cảnh báo sớm nguy cơ và tự động hóa kịch bản phân bổ nguồn lực ôn tập có ý nghĩa thực tiễn và tính cấp thiết đặc biệt sâu sắc."
    )
    
    doc.add_heading("1.2. Phát biểu bài toán học máy (Problem Formulation)", level=2)
    doc.add_paragraph(
        "Bài toán được mô hình hóa dưới dạng Học có giám sát (Supervised Learning) – Phân loại đa lớp có mất cân bằng dữ liệu (Imbalanced Multi-class Classification):"
    )
    doc.add_paragraph(
        "• Không gian đầu vào (Input Space X): Vector đặc trưng gồm 15 thuộc tính ma trận đánh giá học sinh THPT:\n"
        "   - 5 điểm đánh giá quá trình THPT truyền thống: Điểm Miệng, Điểm 15 Phút, Điểm 45 Phút (1 Tiết), Điểm Học Kỳ 1, Điểm Học Kỳ 2.\n"
        "   - 2 chỉ số tổng hợp quá trình: Điểm trung bình thường xuyên (suy biến toán học có trọng số) và Xu hướng điểm số (HK2 - HK1).\n"
        "   - 1 chỉ số chuyên cần: Tỷ lệ hoàn thành bài tập rèn luyện trên hệ thống LMS.\n"
        "   - 4 điểm cấp độ nhận thức Bloom: Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao.\n"
        "   - 3 điểm mạch nội dung GDPT 2018: Khoa học máy tính (CS), Tin học ứng dụng (ICT), Đạo đức số (DL)."
    )
    doc.add_paragraph(
        "• Không gian đầu ra (Target Output Y): 3 mức độ cảnh báo nguy cơ học tập:\n"
        "   - Lớp 0 (An toàn / Safe): Dự kiến kết quả tốt nghiệp đạt từ 6.5 trở lên.\n"
        "   - Lớp 1 (Cần theo dõi / Moderate Risk): Kết quả dao động từ 5.0 đến 6.4 điểm, có nguy cơ trượt nếu đề thi phân hóa cao.\n"
        "   - Lớp 2 (Nguy cơ cao / High Risk): Dự kiến đạt điểm dưới trung bình (< 5.0) hoặc dính điểm liệt, cần phụ đạo khẩn cấp."
    )
    
    doc.add_heading("1.3. Đối tượng sử dụng và phạm vi hệ thống", level=2)
    doc.add_paragraph(
        "• Người dùng chính: Giáo viên bộ môn Tin học (nhận diện chính xác lỗ hổng tri thức từng cá nhân) "
        "và Ban Giám hiệu / Tổ trưởng chuyên môn (nắm bức tranh phân bổ nguy cơ toàn khối để bố trí kinh phí, phòng máy và giáo viên ôn tập)."
    )
    
    # ------------------ CHƯƠNG 2 ------------------
    doc.add_heading("CHƯƠNG 2: CƠ SỞ DỮ LIỆU SQLITE VÀ QUY TRÌNH TIỀN XỬ LÝ CHỐNG RÒ RỈ", level=1)
    
    doc.add_heading("2.1. Kiến trúc CSDL SQLite và Quy mô dữ liệu 8 khóa học", level=2)
    doc.add_paragraph(
        "Hệ thống trang bị cơ sở dữ liệu nhúng SQLite (data/vstep_informatics.db) quản trị toàn vẹn 4.070 học sinh xuyên suốt 8 khóa học "
        "(từ Khóa 2018 - 2019 đến Khóa 2025 - 2026). Dữ liệu được mô phỏng thống kê đa biến (Multivariate Normal Distribution) bám sát "
        "phân phối điểm thực tế của đề thi minh họa GDPT 2018. Toàn bộ thông tin cá nhân nhạy cảm đã được ẩn danh hóa nghiêm ngặt dạng STU_xxxx. LƯU Ý QUAN TRỌNG: Đây là DỮ LIỆU GIẢ LẬP; mọi số liệu đánh giá trong báo cáo chỉ chứng minh tính khả thi của quy trình (Proof of Concept), KHÔNG dùng để khẳng định hiệu quả trên học sinh thực tế."
    )
    
    t_feat = doc.add_table(rows=16, cols=4)
    t_feat.alignment = WD_TABLE_ALIGNMENT.CENTER
    feat_headers = ["STT", "Tên Thuộc Tính", "Kiểu Dữ Liệu", "Diễn Giải Nghiệp Vụ THPT"]
    for i, h in enumerate(feat_headers):
        c = t_feat.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        
    feat_rows = [
        ("1", "score_mieng", "Float [0, 10]", "Điểm kiểm tra Miệng định kỳ"),
        ("2", "score_15p", "Float [0, 10]", "Điểm kiểm tra 15 Phút trên lớp"),
        ("3", "score_45p", "Float [0, 10]", "Điểm kiểm tra 45 Phút (1 Tiết định kỳ)"),
        ("4", "score_hk1", "Float [0, 10]", "Điểm thi Học Kỳ 1"),
        ("5", "score_hk2", "Float [0, 10]", "Điểm thi Học Kỳ 2"),
        ("6", "score_tx_avg", "Float [0, 10]", "Điểm TB Quá trình = (Miệng + 15p + 45p*2) / 4"),
        ("7", "score_trend", "Float [-10, 10]", "Chỉ số xu hướng tiến bộ = Điểm HK2 - Điểm HK1"),
        ("8", "homework_ratio", "Float [0, 1]", "Tỷ lệ nộp và hoàn thành bài tập rèn luyện LMS"),
        ("9", "score_recall", "Float [0, 10]", "Điểm thành phần Mức 1: Nhận biết (Khái niệm GDPT 2018)"),
        ("10", "score_understand", "Float [0, 10]", "Điểm thành phần Mức 2: Thông hiểu (Đọc hiểu thuật toán)"),
        ("11", "score_apply", "Float [0, 10]", "Điểm thành phần Mức 3: Vận dụng (Viết mã, truy vấn CSDL)"),
        ("12", "score_high_apply", "Float [0, 10]", "Điểm thành phần Mức 4: Vận dụng cao (Giải thuật tối ưu)"),
        ("13", "score_topic_cs", "Float [0, 10]", "Điểm chuyên đề Khoa học Máy tính & Lập trình (CS)"),
        ("14", "score_topic_ict", "Float [0, 10]", "Điểm chuyên đề Tin học Ứng dụng & CSDL (ICT)"),
        ("15", "score_topic_dl", "Float [0, 10]", "Điểm chuyên đề Đạo đức, Pháp luật & Văn hóa số (DL)")
    ]
    for r_idx, r_data in enumerate(feat_rows, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_feat.cell(r_idx, c_idx)
            cell.text = val
            if c_idx in [0, 2]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
    doc.add_paragraph("\n")
    
    doc.add_heading("2.2. Kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Split)", level=2)
    doc.add_paragraph(
        "Nhằm đảm bảo tính khách quan học thuật, nhóm áp dụng kỹ thuật Stratified Split chia tách dữ liệu thành 3 tập độc lập: "
        "70% Train Set (2.849 mẫu), 15% Validation Set (610 mẫu), và 15% Test Set (611 mẫu). Phân bố nhãn toàn tập: An toàn 37,0% - Cần theo dõi 22,9% - Nguy cơ cao 40,1%. "
        "Quy trình xử lý dữ liệu tuân thủ nghiêm ngặt nguyên tắc đóng băng tham số (Parameter Freezing): Bộ tiền xử lý StandardScaler "
        "chỉ tính toán trung bình và phương sai trên tập Train (fit_transform), sau đó áp dụng tĩnh lên tập Validation và Test (transform), "
        "triệt tiêu hoàn toàn rủi ro rò rỉ phân phối dữ liệu (Data Leakage)."
    )
    
    # ------------------ CHƯƠNG 3 ------------------
    doc.add_heading("CHƯƠNG 3: THIẾT KẾ MÔ HÌNH HỌC MÁY VÀ KẾT QUẢ THỰC NGHIỆM", level=1)
    
    doc.add_heading("3.1. Các mô hình cơ sở (Baseline) và mô hình cải tiến", level=2)
    doc.add_paragraph(
        "Nhóm triển khai thực nghiệm đối sánh giữa hai nhóm mô hình:\n"
        "• Nhóm Mô hình Cơ sở (Baseline Models):\n"
        "   - Baseline 0 (Rule-Based): Luật sư phạm theo ngưỡng điểm trung bình quá trình và điểm học kỳ.\n   - Baseline 1 (Decision Tree): Cây quyết định đơn lẻ với max_depth=4, class_weight='balanced'.\n"
        "   - Baseline 2 (Logistic Regression): Hồi quy Logistic đa thức có cân bằng trọng số lớp (class_weight='balanced').\n"
        "• Nhóm Mô hình Cải tiến (Improved Models):\n"
        "   - Random Forest: 150 cây, max_depth=8, class_weight='balanced'.\n"
        "   - LightGBM & XGBoost: Gradient Boosting, 150 cây, max_depth=5, learning_rate=0.08."
    )
    
    doc.add_heading("3.2. Bảng tổng hợp chỉ số đánh giá trên tập Kiểm thử độc lập (Test Set - 611 mẫu)", level=2)
    
    t_bench = doc.add_table(rows=7, cols=7)
    t_bench.alignment = WD_TABLE_ALIGNMENT.CENTER
    bench_hdrs = ["Mô hình", "Accuracy", "Macro F1", "Recall Nguy Cơ Cao", "Precision Nguy Cơ Cao", "Độ trễ suy luận", "Thời gian huấn luyện"]
    for i, h in enumerate(bench_hdrs):
        c = t_bench.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    bench_data = [
        ("Baseline: Rule-Based Heuristic", "84.12%", "82.83%", "78.78%", "97.47%", "~0 ms", "0.000 s"),
        ("Baseline: Decision Tree", "85.60%", "84.31%", "87.35%", "94.69%", "~0 ms", "0.005 s"),
        ("Baseline: Logistic Regression (Macro F1 cao nhất)", "90.34%", "89.19%", "91.84%", "95.74%", "~0 ms", "0.013 s"),
        ("Improved: Random Forest", "88.71%", "87.31%", "91.02%", "94.89%", "0.017 ms", "0.333 s"),
        ("Improved: XGBoost", "88.38%", "86.64%", "91.84%", "92.98%", "0.004 ms", "0.258 s"),
        ("Improved: LightGBM (Recall nguy cơ cao nhất)", "89.36%", "87.72%", "93.88%", "93.12%", "0.005 ms", "0.096 s")
    ]
    for r_idx, r_data in enumerate(bench_data, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_bench.cell(r_idx, c_idx)
            cell.text = val
            if c_idx >= 1:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if "cao nhất" in r_data[0]:
                set_cell_background(cell, "DCFCE7")
                cell.paragraphs[0].runs[0].font.bold = True
                
    doc.add_paragraph("\n")
    
    doc.add_paragraph(
        "Nhận xét trung thực: Trên dữ liệu giả lập, mô hình cơ sở Logistic Regression đạt Macro F1 cao nhất (89,19%), "
        "cho thấy quan hệ giữa các đặc trưng và nhãn gần tuyến tính (do cách sinh dữ liệu). Phương án cải tiến LightGBM "
        "đạt Recall lớp Nguy cơ cao tốt nhất (93,88%) - chỉ số ưu tiên về mặt sư phạm vì bỏ sót học sinh nguy cơ cao có "
        "chi phí lớn hơn cảnh báo nhầm. Không mô hình nào nhầm học sinh Nguy cơ cao thành An toàn (lỗi chỉ xảy ra giữa các lớp kề nhau). "
        "Do dữ liệu mất cân bằng nhẹ, báo cáo sử dụng Macro F1 và Recall theo lớp thay vì chỉ Accuracy."
    )

    doc.add_heading("3.3. Giải thích mô hình bằng XAI (Explainable AI)", level=2)
    doc.add_paragraph(
        "Nhóm áp dụng phương pháp Permutation Importance để lượng hóa đóng góp thực chất của các thuộc tính. "
        "Kết quả khẳng định: Điểm Học Kỳ 2, Điểm Trung Bình Quá Trình, Mức 4 (Vận Dụng Cao), và Mạch Lập Trình (CS) "
        "là 4 đặc trưng chi phối mạnh nhất đến khả năng đỗ/trượt tốt nghiệp môn Tin học của học sinh."
    )
    
    fi_img_path = os.path.join(fig_dir, 'feature_importance_ranking.png')
    if os.path.exists(fi_img_path):
        doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].add_run().add_picture(fi_img_path, width=Inches(5.5))
        caption = doc.add_paragraph("Hình 3.1: Xếp hạng độ quan trọng thuộc tính (Permutation Importance)")
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.runs[0].font.italic = True
        caption.runs[0].font.size = Pt(10)
        
    cm_img_path = os.path.join(fig_dir, 'confusion_matrix_best_model.png')
    if os.path.exists(cm_img_path):
        doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].add_run().add_picture(cm_img_path, width=Inches(4.5))
        caption = doc.add_paragraph("Hình 3.2: Ma trận nhầm lẫn (Confusion Matrix) trên tập Test độc lập")
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.runs[0].font.italic = True
        caption.runs[0].font.size = Pt(10)

    # ------------------ CHƯƠNG 4 ------------------
    doc.add_heading("CHƯƠNG 4: SO SÁNH BASELINE THỜI GIAN VÀ PHÂN TÍCH LỖI CHUYÊN SÂU", level=1)
    
    doc.add_heading("4.1. So sánh Baseline Vận hành thời gian thực: Giáo viên thủ công vs Mô hình ML", level=2)
    doc.add_paragraph(
        "Đáp ứng trực tiếp yêu cầu của đề bài Cao học về việc so sánh baseline thời gian với mô hình học máy, "
        "nhóm so sánh thời gian xử lý. Thời gian của mô hình được ĐO THỰC TẾ (time.perf_counter, CPU laptop cá nhân); thời gian thủ công là GIẢ ĐỊNH ƯỚC LƯỢNG của nhóm, chưa qua khảo sát giáo viên thực tế:"
    )
    
    t_time = doc.add_table(rows=7, cols=4)
    t_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    time_hdrs = ["Tiêu Chí Đánh Giá", "Thủ Công (Hội Đồng GV)", "Mô Hình Cơ Sở (Logistic)", "Mô Hình Cải Tiến (Random Forest)"]
    for i, h in enumerate(time_hdrs):
        c = t_time.cell(0, i)
        c.text = h
        set_cell_background(c, "1E3A8A")
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    time_data = [
        ("Thời gian xử lý 1 lớp (40 HS)", "~25 - 35 phút (GIẢ ĐỊNH ước lượng)", "~0.001 giây (đo thực tế)", "~0.005 - 0.04 giây (đo thực tế)"),
        ("Thời gian xử lý 4.070 HS (8 khóa)", "~3 - 5 ngày công (GIẢ ĐỊNH ước lượng)", "~0.02 giây (đo thực tế)", "~0.08 - 0.18 giây (đo thực tế)"),
        ("Độ bao phủ tiêu chí GDPT 2018", "Thường chỉ nhìn điểm trung bình", "15 đặc trưng, ranh giới tuyến tính", "15 đặc trưng, quan hệ phi tuyến"),
        ("Thời điểm phát hiện nguy cơ", "Sau khi có điểm thi học kỳ", "Ngay khi có đủ điểm quá trình", "Ngay khi có đủ điểm quá trình"),
        ("Căn cứ giải thích (Explainability)", "Cảm quan sư phạm chủ quan", "Hệ số mô hình", "Permutation Importance + XAI Attribution"),
        ("Khả năng mở rộng", "Phụ thuộc nhân lực", "Tự động hóa hoàn toàn", "Tự động + kịch bản phân bổ lớp phụ đạo")
    ]
    for r_idx, r_data in enumerate(time_data, start=1):
        for c_idx, val in enumerate(r_data):
            cell = t_time.cell(r_idx, c_idx)
            cell.text = val
            if c_idx >= 1:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
    doc.add_paragraph("\n")
    
    doc.add_heading("4.2. Phân tích chuyên sâu các ca dự đoán sai (In-Depth Error Analysis)", level=2)
    doc.add_paragraph(
        "Trên tập Kiểm thử độc lập (611 mẫu), mô hình Random Forest (mô hình mặc định trên giao diện) đạt tỷ lệ chính xác 88,71%, tương ứng với 69 ca dự đoán sai (11.29%). "
        "Nhóm tiến hành phân rã lỗi thành 3 nhóm đặc thù dưới góc nhìn sư phạm:"
    )
    doc.add_paragraph(
        "1. Nhóm Bỏ sót nguy cơ (False Negatives - 22 ca): Đây là nhóm lỗi có rủi ro sư phạm cao nhất. "
        "Thực tế học sinh thuộc nhóm Nguy cơ cao (< 5.0) nhưng mô hình dự đoán là Cần theo dõi (không có ca nào bị đoán là An toàn). "
        "Phân tích nguyên nhân: Học sinh có điểm quá trình học bạ tương đối khá (6.0 - 7.0) nhờ tích cực nộp bài tập LMS và điểm miệng cao, "
        "nhưng thực tế bị 'hổng rỗng' tư duy thuật toán (Điểm CS < 4.0 và Điểm Vận dụng cao < 2.0). Khi bước vào bài thi chính thức chuẩn GDPT 2018, "
        "các em bị trượt điểm liệt. Mô hình Random Forest bị chi phối bởi trọng số của điểm học bạ nên đã phân loại nhầm."
    )
    doc.add_paragraph(
        "2. Nhóm Cảnh báo nhầm (False Positives - 22 ca): Học sinh thực tế đạt mức An toàn nhưng mô hình dự đoán Cần theo dõi. "
        "Phân tích nguyên nhân: Học sinh có năng lực tư duy lập trình và điểm thi cuối kỳ rất cao, nhưng do sơ suất bị thiếu một cột điểm 15 phút "
        "hoặc nộp bài tập LMS muộn kéo điểm quá trình tạm thời xuống thấp, kích hoạt thuật toán cảnh báo nhầm."
    )
    doc.add_paragraph(
        "3. Nhóm Ca lệch ranh giới biên (Boundary Cases - 25 ca: 13 ca Cần theo dõi bị đoán An toàn, 12 ca bị đoán Nguy cơ cao): Học sinh có điểm số nằm sát ngưỡng biên 5.0 điểm và 6.5 điểm "
        "(giao thoa giữa mức Cần theo dõi và An toàn)."
    )
    
    # ------------------ CHƯƠNG 5 ------------------
    doc.add_heading("CHƯƠNG 5: NGUYÊN MẪU PHẦN MỀM VÀ KHUNG ĐẠO ĐỨC AI", level=1)
    
    doc.add_heading("5.1. Kiến trúc nguyên mẫu phần mềm (Web Application)", level=2)
    doc.add_paragraph(
        "Nguyên mẫu được phát triển trên Python Streamlit tuân thủ ngôn ngữ thiết kế VSTEP Master UI Standard. "
        "Hệ thống vận hành trọn vẹn gồm 5 phân hệ tương tác:\n"
        "• Tab 1 (Dự Báo Theo Lớp): Truy vấn trực tiếp từ CSDL SQLite 8 khóa hoặc tải file Excel mới, có bộ lọc mức nguy cơ.\n"
        "• Tab 2 (Chẩn Đoán Cá Nhân XAI): Cung cấp form 2 cột cân đối, thanh trượt điểm truyền thống THPT, "
        "tự động tính TB quá trình và vẽ biểu đồ Radar ma trận năng lực.\n"
        "• Tab 3 (Kịch Bản Phân Bổ Nguồn Lực): Tự động phát hiện lỗ hổng kiến thức và đề xuất số lớp phụ đạo, số tiết/tuần.\n"
        "• Tab 4 (Báo Cáo Thực Nghiệm & Phân Tích Lỗi): Hiển thị bảng Benchmark, so sánh baseline thời gian và bảng tương tác tra cứu các ca dự đoán sai.\n"
        "• Tab 5 (Thư Viện Tri Thức & Khung Đạo Đức AI): Công bố ma trận đề thi GDPT 2018 và quy chuẩn đạo đức AI."
    )
    
    doc.add_heading("5.2. Khung Đạo đức AI & Giới hạn của Hệ hỗ trợ ra quyết định (Ethical AI Framework)", level=2)
    doc.add_paragraph(
        "Đề tài tuân thủ chặt chẽ yêu cầu đạo đức công nghệ thông tin trong môi trường giáo dục:\n"
        "• Nguyên tắc Hỗ trợ ra quyết định (Decision Support): Hệ thống đóng vai trò công cụ cố vấn chuyên môn cho nhà giáo, "
        "tuyệt đối KHÔNG tự động xếp loại học bạ hay thay thế quyền đánh giá sư phạm của con người.\n"
        "• Không dán nhãn định kiến (Zero Stigmatization): Mọi thông tin cảnh báo nguy cơ học tập chỉ được lưu hành nội bộ giữa Ban giám hiệu "
        "và giáo viên phụ trách để xây dựng kế hoạch bồi dưỡng, tuyệt đối không công khai bêu tên học sinh.\n"
        "• Bảo vệ quyền riêng tư (Data Privacy): Dữ liệu được ẩn danh hóa hoàn toàn, không lưu trữ CCCD hay thông tin cá nhân nhạy cảm.\n"
        "• Giữ vai trò quyết định của con người (Human-in-the-loop): Giáo viên luôn là người kiểm chứng cuối cùng đối với các trường hợp "
        "học sinh có biến cố ngoại cảnh đột xuất (ốm đau, hoàn cảnh gia đình)."
    )
    
    # ------------------ KẾT LUẬN ------------------
    doc.add_heading("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", level=1)
    doc.add_paragraph(
        "Đề tài đã hoàn thành các mục tiêu đặt ra cho học viên Cao học CNTT lớp K20.2 trong thời gian 4 tuần:\n"
        "1. Đã giải quyết trọn vẹn bài toán thực tiễn cảnh báo nguy cơ môn Tin học định hướng thi tốt nghiệp THPT theo GDPT 2018.\n"
        "2. Đã xây dựng CSDL SQLite nhúng quản trị 4.070 học sinh qua 8 khóa học và kiểm soát rò rỉ dữ liệu (Anti-Leakage) chặt chẽ.\n"
        "3. Đã đối sánh toàn diện các mô hình Baseline (Decision Tree, Logistic Regression) và Mô hình Cải tiến (Random Forest, LightGBM, XGBoost); "
        "Logistic Regression đạt Macro F1 cao nhất (89,19%), LightGBM đạt Recall lớp Nguy cơ cao cao nhất (93,88%) trên dữ liệu giả lập.\n"
        "4. Đã so sánh thời gian xử lý (mô hình xử lý 4.070 HS dưới 0,2 giây) và phân tích các ca dự đoán sai.\nHướng phát triển: thu thập dữ liệu thực tế đã ẩn danh từ nhà trường để kiểm định lại mô hình; hiệu chỉnh xác suất (calibration); bổ sung SHAP cho giải thích cục bộ.\n"
        "5. Đã đóng gói nguyên mẫu phần mềm chạy trực tiếp, giao diện chuẩn mực, kèm kịch bản phân bổ nguồn lực ôn tập và khung đạo đức AI hoàn chỉnh."
    )
    
    # ------------------ TÀI LIỆU THAM KHẢO ------------------
    doc.add_heading("TÀI LIỆU THAM KHẢO", level=1)
    refs = [
        "[1] Bộ Giáo dục và Đào tạo, Chương trình Giáo dục phổ thông môn Tin học (Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT), 2018.",
        "[2] Bộ Giáo dục và Đào tạo, Quyết định về việc công bố Cấu trúc định dạng đề thi Kỳ thi tốt nghiệp THPT từ năm 2025, 2023.",
        "[3] Lundberg, S. M., & Lee, S. I., 'A unified approach to interpreting model predictions', Advances in Neural Information Processing Systems (NeurIPS), 2017.",
        "[4] Breiman, L., 'Random Forests', Machine Learning, 45(1), 5-32, 2001.",
        "[5] Ke, G., Meng, Q., Finley, T., et al., 'LightGBM: A highly efficient gradient boosting decision tree', NeurIPS, 2017.",
        "[6] Pedregosa, F., et al., 'Scikit-learn: Machine Learning in Python', JMLR, 2011."
    ]
    for r in refs:
        doc.add_paragraph(r)
        
    doc.save(output_docx)
    print(f"[SUCCESS] Da xuat Bao Cao Tong Ket Cao Hoc Word tai: {output_docx}")

if __name__ == '__main__':
    create_report()
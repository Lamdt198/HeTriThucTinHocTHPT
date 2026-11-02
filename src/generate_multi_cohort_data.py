"""
generate_multi_cohort_data.py
Sinh bộ dữ liệu 8 khóa học sinh THPT (mỗi khóa từ 400 - 600 học sinh)
theo định hướng ma trận năng lực Tin học GDPT 2018.
Các cột điểm số áp dụng theo cấu trúc truyền thống trường học Việt Nam:
Điểm Miệng, Điểm 15 Phút, Điểm 45 Phút, Điểm Học Kỳ 1, Điểm Học Kỳ 2.
Đồng thời nạp trực tiếp toàn bộ dữ liệu vào SQLite Database.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src_dir = os.path.join(BASE_DIR, 'src')
if src_dir not in sys.path:
    sys.path.append(src_dir)

from database import import_cohort_dataframe, init_database

HO_LIST = [
    'Nguyễn', 'Trần', 'Lê', 'Phạm', 'Hoàng', 'Huỳnh', 'Phan', 'Vũ', 
    'Võ', 'Đặng', 'Bùi', 'Đỗ', 'Hồ', 'Ngô', 'Dương', 'Lý', 'Đinh', 'Đoàn'
]

DEM_LIST = [
    'Văn', 'Thị', 'Minh', 'Hoàng', 'Đức', 'Hữu', 'Tuấn', 'Ngọc', 
    'Thu', 'Mai', 'Thanh', 'Quốc', 'Hải', 'Bảo', 'Gia', 'Phương'
]

TEN_LIST = [
    'An', 'Bình', 'Cường', 'Dũng', 'Đạt', 'Giang', 'Hà', 'Hải', 'Hiếu', 
    'Hoàng', 'Hùng', 'Huy', 'Khánh', 'Khoa', 'Lâm', 'Linh', 'Long', 'Mai', 
    'Nam', 'Nghĩa', 'Nhung', 'Phát', 'Phúc', 'Quân', 'Sơn', 'Tâm', 'Thảo', 
    'Thắng', 'Thịnh', 'Trang', 'Trung', 'Tuấn', 'Tùng', 'Vy', 'Yến'
]

COHORTS_CONFIG = [
    {'cohort_name': 'Khóa 2018 - 2019', 'cohort_code': 'K18', 'num_students': 485, 'seed': 101},
    {'cohort_name': 'Khóa 2019 - 2020', 'cohort_code': 'K19', 'num_students': 512, 'seed': 102},
    {'cohort_name': 'Khóa 2020 - 2021', 'cohort_code': 'K20', 'num_students': 468, 'seed': 103},
    {'cohort_name': 'Khóa 2021 - 2022', 'cohort_code': 'K21', 'num_students': 530, 'seed': 104},
    {'cohort_name': 'Khóa 2022 - 2023', 'cohort_code': 'K22', 'num_students': 495, 'seed': 105},
    {'cohort_name': 'Khóa 2023 - 2024', 'cohort_code': 'K23', 'num_students': 525, 'seed': 106},
    {'cohort_name': 'Khóa 2024 - 2025', 'cohort_code': 'K24', 'num_students': 510, 'seed': 107},
    {'cohort_name': 'Khóa 2025 - 2026', 'cohort_code': 'K25', 'num_students': 545, 'seed': 108},
]

def generate_cohort_students(cohort_info: dict) -> pd.DataFrame:
    n = cohort_info['num_students']
    seed = cohort_info['seed']
    c_code = cohort_info['cohort_code']
    c_name = cohort_info['cohort_name']
    
    np.random.seed(seed)
    
    # 1. Mã học sinh & Họ tên
    student_ids = [f"STU_{c_code}_{i+1:04d}" for i in range(n)]
    full_names = [
        f"{np.random.choice(HO_LIST)} {np.random.choice(DEM_LIST)} {np.random.choice(TEN_LIST)}"
        for _ in range(n)
    ]
    
    # Lớp
    num_classes = max(10, int(np.ceil(n / 38)))
    classes_pool = [f"12A{i+1}" for i in range(num_classes)]
    assigned_classes = np.random.choice(classes_pool, size=n)
    
    # 2. Năng lực nội tại tiềm ẩn (Latent Ability theta ~ N(0, 1))
    year_shift = (int(c_code.replace('K', '')) - 18) * 0.04
    theta = np.random.normal(0, 1, n)
    
    # 3. Điểm 4 mức độ nhận thức GDPT 2018 (Thang điểm 10)
    score_recall = np.clip(7.2 + year_shift + 1.8 * theta + np.random.normal(0, 0.7, n), 1.0, 10.0)
    score_understand = np.clip(6.4 + year_shift + 2.0 * theta + np.random.normal(0, 0.8, n), 0.5, 10.0)
    score_apply = np.clip(5.3 + year_shift + 2.3 * theta + np.random.normal(0, 1.0, n), 0.0, 10.0)
    score_high_apply = np.clip(3.9 + year_shift + 2.4 * theta + np.random.normal(0, 1.2, n), 0.0, 10.0)
    
    # 4. Điểm theo 3 mạch kiến thức Tin học 2018
    score_cs = np.clip(5.1 + 2.2 * theta + np.random.normal(0, 1.1, n), 0.0, 10.0)
    score_ict = np.clip(6.6 + 1.9 * theta + np.random.normal(0, 0.9, n), 1.0, 10.0)
    score_dl = np.clip(7.4 + 1.5 * theta + np.random.normal(0, 0.8, n), 2.0, 10.0)
    
    # 5. Điểm đánh giá truyền thống: Miệng, 15 Phút, 45 Phút, Học Kỳ 1, Học Kỳ 2
    # Điểm Miệng: Khảo sát nhận biết và thông hiểu lý thuyết
    score_mieng = np.clip(0.6 * score_recall + 0.4 * score_understand + np.random.normal(0, 0.6, n), 0.0, 10.0)
    # Điểm 15 Phút: Vận dụng nhanh, cú pháp code và trắc nghiệm
    score_15p = np.clip(0.4 * score_understand + 0.4 * score_apply + 0.2 * score_cs + np.random.normal(0, 0.6, n), 0.0, 10.0)
    # Điểm 45 Phút (1 tiết): Bài kiểm tra toàn diện giữa kỳ
    score_45p = np.clip(
        0.25 * score_recall + 0.35 * score_understand + 0.25 * score_apply + 0.15 * score_high_apply + np.random.normal(0, 0.5, n),
        0.0, 10.0
    )
    # Điểm Học Kỳ 1 (Tổng kết HK1)
    score_hk1 = np.clip(
        0.20 * score_recall + 0.30 * score_understand + 0.30 * score_apply + 0.20 * score_high_apply + np.random.normal(0, 0.5, n),
        0.0, 10.0
    )
    # Điểm Học Kỳ 2 (Đề thi học kỳ 2 phân hóa mạnh hơn theo định hướng tốt nghiệp)
    score_hk2 = np.clip(
        0.15 * score_recall + 0.25 * score_understand + 0.35 * score_apply + 0.25 * score_high_apply + np.random.normal(0, 0.5, n),
        0.0, 10.0
    )
    
    # 6. Điểm Trung Bình Quá Trình: (Miệng + 15p + 45p*2) / 4
    score_tx_avg = (score_mieng + score_15p + score_45p * 2.0) / 4.0
    
    # 7. Tỷ lệ hoàn thành bài tập trên LMS
    homework_ratio = np.clip(0.70 + 0.18 * theta + np.random.normal(0, 0.12, n), 0.05, 1.0)
    
    # 8. Xu hướng điểm số: Chênh lệch tiến bộ giữa Học Kỳ 2 và Học Kỳ 1
    score_trend = score_hk2 - score_hk1
    
    # 9. Điểm thi tốt nghiệp mô phỏng (Ground Truth)
    predicted_graduation_score = (
        0.15 * score_recall +
        0.25 * score_understand +
        0.35 * score_apply +
        0.25 * score_high_apply +
        0.15 * (homework_ratio * 10 - 7.0) +
        0.10 * score_trend +
        np.random.normal(0, 0.40, n)
    )
    predicted_graduation_score = np.clip(predicted_graduation_score, 0.5, 10.0)
    
    # Gán nhãn nguy cơ
    risk_labels = []
    risk_codes = []
    for s in predicted_graduation_score:
        if s >= 6.5:
            risk_labels.append("An toàn (>= 6.5)")
            risk_codes.append(0)
        elif s >= 5.0:
            risk_labels.append("Cần theo dõi (5.0 - 6.4)")
            risk_codes.append(1)
        else:
            risk_labels.append("Nguy cơ cao (< 5.0)")
            risk_codes.append(2)
            
    df = pd.DataFrame({
        'Mã Học Sinh': student_ids,
        'Họ Và Tên': full_names,
        'Khóa Học': [c_name] * n,
        'Lớp': assigned_classes,
        'Điểm Miệng': np.round(score_mieng, 1),
        'Điểm 15 Phút': np.round(score_15p, 1),
        'Điểm 45 Phút': np.round(score_45p, 1),
        'Điểm Học Kỳ 1': np.round(score_hk1, 1),
        'Điểm Học Kỳ 2': np.round(score_hk2, 1),
        'Điểm Trung Bình Quá Trình': np.round(score_tx_avg, 1),
        'Tỷ Lệ Hoàn Thành BTVN': np.round(homework_ratio, 2),
        'Điểm Nhận Biết': np.round(score_recall, 1),
        'Điểm Thông Hiểu': np.round(score_understand, 1),
        'Điểm Vận Dụng': np.round(score_apply, 1),
        'Điểm Vận Dụng Cao': np.round(score_high_apply, 1),
        'Điểm Khoa Học Máy Tính (CS)': np.round(score_cs, 1),
        'Điểm Tin Học Ứng Dụng (ICT)': np.round(score_ict, 1),
        'Điểm Đạo Đức Số (DL)': np.round(score_dl, 1),
        'Xu Hướng Điểm Số': np.round(score_trend, 2),
        'Điểm Thi Tốt Nghiệp Dự Kiến': np.round(predicted_graduation_score, 1),
        'Mức Độ Nguy Cơ': risk_labels,
        'Mã Mức Nguy Cơ': risk_codes
    })
    
    return df

def generate_and_save_all_datasets():
    raw_dir = os.path.join(BASE_DIR, 'data', 'raw')
    tung_khoa_dir = os.path.join(raw_dir, 'tung_khoa')
    mau_upload_dir = os.path.join(raw_dir, 'mau_upload_theo_khoa')
    
    os.makedirs(raw_dir, exist_ok=True)
    os.makedirs(tung_khoa_dir, exist_ok=True)
    os.makedirs(mau_upload_dir, exist_ok=True)
    
    # Xóa file demo 60 mẫu nếu có
    demo_file = os.path.join(raw_dir, 'sample_upload_test.xlsx')
    if os.path.exists(demo_file):
        try:
            os.remove(demo_file)
            print(f"[OK] Đã loại bỏ file mẫu cũ: {demo_file}")
        except Exception:
            pass
            
    all_cohort_dfs = []
    cohort_dict = {}
    
    print("=" * 70)
    print("KHỞI TẠO BỘ DỮ LIỆU 8 KHÓA (ĐIỂM MIỆNG, 15P, 45P, HK1, HK2)")
    print("=" * 70)
    
    for c_info in COHORTS_CONFIG:
        df_c = generate_cohort_students(c_info)
        all_cohort_dfs.append(df_c)
        c_code = c_info['cohort_code']
        c_name = c_info['cohort_name']
        cohort_dict[c_code] = df_c
        
        clean_name = c_name.replace(' ', '_').replace('-', 'den')
        xlsx_c_path = os.path.join(tung_khoa_dir, f"Bang_Diem_{clean_name}.xlsx")
        csv_c_path = os.path.join(tung_khoa_dir, f"Bang_Diem_{clean_name}.csv")
        df_c.to_excel(xlsx_c_path, index=False)
        df_c.to_csv(csv_c_path, index=False, encoding='utf-8-sig')
        
        # Mẫu upload kiểm thử (không có cột kết quả)
        input_cols = [c for c in df_c.columns if c not in ['Điểm Thi Tốt Nghiệp Dự Kiến', 'Mức Độ Nguy Cơ', 'Mã Mức Nguy Cơ']]
        df_upload = df_c[input_cols].copy()
        upload_path = os.path.join(mau_upload_dir, f"Mau_Upload_{clean_name}.xlsx")
        df_upload.to_excel(upload_path, index=False)
        
        print(f" [+] {c_name} ({c_code}): {len(df_c)} học sinh | Lưu tại: Bang_Diem_{clean_name}.xlsx")
        
    df_all = pd.concat(all_cohort_dfs, ignore_index=True)
    print("-" * 70)
    print(f" => Tổng số học sinh cả 8 khóa: {len(df_all)} học sinh.")
    print("-" * 70)
    
    # Lưu file Excel tổng hợp
    all_excel_path = os.path.join(raw_dir, 'du_lieu_8_khoa_tin_hoc_thpt.xlsx')
    with pd.ExcelWriter(all_excel_path, engine='openpyxl') as writer:
        df_all.to_excel(writer, sheet_name='Tat_Ca_8_Khoa', index=False)
        for c_code, df_c in cohort_dict.items():
            sheet_title = f"{c_code}_{len(df_c)}HS"
            df_c.to_excel(writer, sheet_name=sheet_title, index=False)
            
    print(f"[OK] Đã lưu Excel tổng hợp: {all_excel_path}")
    
    # Lưu CSV tổng hợp
    all_csv_path = os.path.join(raw_dir, 'du_lieu_8_khoa_tin_hoc_thpt.csv')
    df_all.to_csv(all_csv_path, index=False, encoding='utf-8-sig')
    print(f"[OK] Đã lưu CSV tổng hợp: {all_csv_path}")
    
    # Nạp trực tiếp vào SQLite Database
    print("\n[INFO] Đang nạp dữ liệu vào Cơ sở Dữ liệu SQLite (vstep_informatics.db)...")
    db_path = os.path.join(BASE_DIR, 'data', 'vstep_informatics.db')
    inserted = import_cohort_dataframe(df_all, db_path)
    print(f"[OK] Đã nạp thành công {inserted} học sinh vào CSDL SQLite tại {db_path}!")
    
    # Cập nhật tập raw huấn luyện tương thích
    train_raw_path = os.path.join(raw_dir, 'informatics_assessment_raw.csv')
    df_all.to_csv(train_raw_path, index=False, encoding='utf-8-sig')

if __name__ == '__main__':
    generate_and_save_all_datasets()

"""
data_generator.py
Sinh bộ dữ liệu mô phỏng chuẩn hóa (Realistic Synthetic Dataset) 
bám sát ma trận kết quả đánh giá môn Tin học THPT theo chương trình GDPT 2018.
"""

import os
import numpy as np
import pandas as pd

def generate_informatics_assessment_data(
    num_students: int = 1200,
    random_state: int = 42
) -> pd.DataFrame:
    """
    Sinh tập dữ liệu học sinh với các thuộc tính ma trận đánh giá GDPT 2018.
    
    Phân bổ nhãn nguy cơ (target_risk_level):
    - 0: An toàn (Safe / Low Risk) (~65%)
    - 1: Cần theo dõi (Moderate Risk) (~22%)
    - 2: Nguy cơ cao (High Risk / At Risk of Failing) (~13%)
    """
    np.random.seed(random_state)
    
    classes = ['12A1', '12A2', '12A3', '12A4', '12A5', '12A6', '12A7', '12A8', '12A9', '12A10']
    student_ids = [f"STU_{i+1:04d}" for i in range(num_students)]
    assigned_classes = np.random.choice(classes, size=num_students)
    
    # Năng lực nội tại tiềm ẩn (Latent Ability Factor theta ~ N(0, 1))
    theta = np.random.normal(0, 1, num_students)
    
    # 1. Điểm các mức độ nhận thức (Thang điểm 10)
    score_recall = np.clip(7.2 + 1.8 * theta + np.random.normal(0, 0.7, num_students), 1.0, 10.0)
    score_understand = np.clip(6.4 + 2.0 * theta + np.random.normal(0, 0.8, num_students), 0.5, 10.0)
    score_apply = np.clip(5.3 + 2.3 * theta + np.random.normal(0, 1.0, num_students), 0.0, 10.0)
    score_high_apply = np.clip(3.9 + 2.4 * theta + np.random.normal(0, 1.2, num_students), 0.0, 10.0)
    
    # 2. Điểm theo 3 mạch kiến thức Tin học GDPT 2018
    # Khoa học máy tính (CS - Lập trình, Thuật toán, Cấu trúc dữ liệu)
    score_topic_cs = np.clip(5.1 + 2.2 * theta + np.random.normal(0, 1.1, num_students), 0.0, 10.0)
    # Ứng dụng tin học (ICT - Mạng máy tính, CSDL, Phần mềm bảng tính/đồ họa)
    score_topic_ict = np.clip(6.6 + 1.9 * theta + np.random.normal(0, 0.9, num_students), 1.0, 10.0)
    # Đạo đức, pháp luật & Văn hóa số (DL - An toàn thông tin, Sở hữu trí tuệ)
    score_topic_dl = np.clip(7.4 + 1.5 * theta + np.random.normal(0, 0.8, num_students), 2.0, 10.0)
    
    # 3. Điểm kiểm tra thường xuyên và giữa kỳ
    score_tx_1 = np.clip(0.4 * score_recall + 0.6 * score_understand + np.random.normal(0, 0.5, num_students), 0.0, 10.0)
    score_tx_2 = np.clip(0.3 * score_understand + 0.4 * score_apply + 0.3 * score_topic_cs + np.random.normal(0, 0.6, num_students), 0.0, 10.0)
    score_tx_avg = (score_tx_1 + score_tx_2) / 2.0
    
    score_midterm = np.clip(
        0.25 * score_recall + 0.30 * score_understand + 0.30 * score_apply + 0.15 * score_high_apply + np.random.normal(0, 0.5, num_students),
        0.0, 10.0
    )
    
    # 4. Tỷ lệ hoàn thành bài tập trên LMS
    homework_ratio = np.clip(0.70 + 0.18 * theta + np.random.normal(0, 0.12, num_students), 0.05, 1.0)
    
    # 5. Xu hướng tiến bộ: giữa kỳ so với thường xuyên
    score_trend = score_midterm - score_tx_avg
    
    # 6. Dự báo điểm thi tốt nghiệp mô phỏng (Ground Truth) để gán nhãn rủi ro
    predicted_graduation_score = (
        0.20 * score_recall +
        0.25 * score_understand +
        0.35 * score_apply +
        0.20 * score_high_apply +
        0.15 * (homework_ratio * 10 - 7.0) +
        np.random.normal(0, 0.45, num_students)
    )
    predicted_graduation_score = np.clip(predicted_graduation_score, 0.5, 10.0)
    
    # Gán nhãn mục tiêu 3 mức:
    # 0: An toàn (score >= 6.5)
    # 1: Cần theo dõi (5.0 <= score < 6.5)
    # 2: Nguy cơ cao (score < 5.0)
    target_risk_level = np.zeros(num_students, dtype=int)
    target_risk_level[(predicted_graduation_score >= 5.0) & (predicted_graduation_score < 6.5)] = 1
    target_risk_level[predicted_graduation_score < 5.0] = 2
    
    df = pd.DataFrame({
        'student_id': student_ids,
        'class_name': assigned_classes,
        'score_tx_1': np.round(score_tx_1, 1),
        'score_tx_2': np.round(score_tx_2, 1),
        'score_tx_avg': np.round(score_tx_avg, 1),
        'score_midterm': np.round(score_midterm, 1),
        'homework_ratio': np.round(homework_ratio, 2),
        'score_recall': np.round(score_recall, 1),
        'score_understand': np.round(score_understand, 1),
        'score_apply': np.round(score_apply, 1),
        'score_high_apply': np.round(score_high_apply, 1),
        'score_topic_cs': np.round(score_topic_cs, 1),
        'score_topic_ict': np.round(score_topic_ict, 1),
        'score_topic_dl': np.round(score_topic_dl, 1),
        'score_trend': np.round(score_trend, 2),
        'est_final_score': np.round(predicted_graduation_score, 1),
        'target_risk_level': target_risk_level
    })
    
    return df

def save_initial_datasets():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    raw_dir = os.path.join(base_dir, 'data', 'raw')
    os.makedirs(raw_dir, exist_ok=True)
    
    # 1. Sinh tập dữ liệu huấn luyện đầy đủ 1200 học sinh
    df = generate_informatics_assessment_data(num_students=1200, random_state=42)
    csv_path = os.path.join(raw_dir, 'informatics_assessment_raw.csv')
    xlsx_path = os.path.join(raw_dir, 'informatics_assessment_raw.xlsx')
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    print(f"[OK] Saved raw dataset: {csv_path} ({df.shape[0]} rows)")
    
    # 2. Sinh tập mẫu test tải lên Web (không có cột est_final_score và target_risk_level)
    sample_test_df = df.sample(60, random_state=123).copy()
    sample_upload = sample_test_df.drop(columns=['est_final_score', 'target_risk_level'])
    upload_sample_path = os.path.join(raw_dir, 'sample_upload_test.xlsx')
    sample_upload.to_excel(upload_sample_path, index=False)
    print(f"[OK] Saved sample upload file for demo: {upload_sample_path} ({sample_upload.shape[0]} rows)")

if __name__ == '__main__':
    save_initial_datasets()

"""
database.py
Module quản lý Cơ sở Dữ liệu Nhúng SQLite cho Hệ Tri thức Môn Tin học THPT.
Lưu trữ hồ sơ học sinh 8 khóa, ma trận điểm đánh giá, lịch sử dự báo và cơ sở tri thức luật.
"""

import os
import sys
import sqlite3
from datetime import datetime
from typing import Dict, List, Tuple, Any, Optional
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DB_PATH = os.path.join(BASE_DIR, 'data', 'vstep_informatics.db')

def get_connection(db_path: Optional[str] = None) -> sqlite3.Connection:
    path = db_path or DEFAULT_DB_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn

def init_database(db_path: Optional[str] = None):
    """
    Khởi tạo cấu trúc các bảng trong SQLite Database.
    """
    conn = get_connection(db_path)
    cur = conn.cursor()
    
    # 1. Bảng học sinh
    cur.execute("""
    CREATE TABLE IF NOT EXISTS students (
        student_id TEXT PRIMARY KEY,
        full_name TEXT NOT NULL,
        cohort TEXT NOT NULL,
        class_name TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    
    # 2. Bảng ma trận kết quả đánh giá năng lực
    cur.execute("""
    CREATE TABLE IF NOT EXISTS assessments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id TEXT NOT NULL,
        score_mieng REAL NOT NULL,
        score_15p REAL NOT NULL,
        score_45p REAL NOT NULL,
        score_hk1 REAL NOT NULL,
        score_hk2 REAL NOT NULL,
        score_tx_avg REAL NOT NULL,
        homework_ratio REAL NOT NULL,
        score_recall REAL NOT NULL,
        score_understand REAL NOT NULL,
        score_apply REAL NOT NULL,
        score_high_apply REAL NOT NULL,
        score_topic_cs REAL NOT NULL,
        score_topic_ict REAL NOT NULL,
        score_topic_dl REAL NOT NULL,
        score_trend REAL NOT NULL,
        est_final_score REAL,
        target_risk_level INTEGER,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(student_id)
    );
    """)
    
    # 3. Bảng nhật ký dự báo và can thiệp (Audit Trail & Tracking)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS prediction_logs (
        log_id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id TEXT NOT NULL,
        model_used TEXT NOT NULL,
        predicted_risk INTEGER NOT NULL,
        risk_label TEXT NOT NULL,
        high_risk_prob REAL NOT NULL,
        primary_gap TEXT,
        recommendation TEXT,
        predicted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(student_id) REFERENCES students(student_id)
    );
    """)
    
    # 4. Bảng Cơ sở tri thức (Knowledge Rules Base)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS knowledge_rules (
        rule_code TEXT PRIMARY KEY,
        domain TEXT NOT NULL,
        benchmark_threshold REAL NOT NULL,
        pedagogical_advice TEXT NOT NULL
    );
    """)
    
    # Nạp sẵn các luật tri thức chuẩn GDPT 2018
    default_rules = [
        ('RULE_CS_GAP', 'Lập trình & Thuật toán (CS)', 5.5, 
         'Hổng tư duy lập trình và cấu trúc dữ liệu. Đề xuất kèm cặp thực hành viết code Python 2 tiết/tuần.'),
        ('RULE_ICT_GAP', 'Ứng dụng Tin học (ICT)', 6.0, 
         'Chưa thành thạo CSDL quan hệ và phần mềm văn phòng. Bổ sung bài tập thực hành truy vấn SQL và Excel.'),
        ('RULE_VDC_GAP', 'Vận dụng cao & Trả lời ngắn', 5.0, 
         'Gặp khó khăn ở các câu hỏi phân hóa thực tiễn. Rèn luyện kỹ năng phân tích bài toán và tối ưu độ phức tạp.'),
        ('RULE_DL_GAP', 'Đạo đức, Pháp luật & Văn hóa số (DL)', 6.5, 
         'Cần củng cố kiến thức an toàn thông tin, bảo vệ dữ liệu cá nhân và sở hữu trí tuệ số.')
    ]
    
    cur.executemany("""
    INSERT OR REPLACE INTO knowledge_rules (rule_code, domain, benchmark_threshold, pedagogical_advice)
    VALUES (?, ?, ?, ?);
    """, default_rules)
    
    conn.commit()
    conn.close()

def import_cohort_dataframe(df: pd.DataFrame, db_path: Optional[str] = None) -> int:
    """
    Nạp dữ liệu học sinh và điểm số từ DataFrame vào SQLite Database.
    Hỗ trợ cả tiêu đề tiếng Việt lẫn tiếng Anh.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    cur = conn.cursor()
    
    # Chuẩn hóa tên cột
    from preprocessor import standardize_dataframe
    df_std = standardize_dataframe(df)
    
    inserted_count = 0
    for _, row in df_std.iterrows():
        s_id = str(row.get('student_id', ''))
        if not s_id:
            continue
            
        name = str(row.get('full_name', row.get('Họ Và Tên', 'Học Sinh')))
        cohort = str(row.get('cohort', row.get('Khóa Học', 'Khóa 2025 - 2026')))
        c_name = str(row.get('class_name', row.get('Lớp', '12A1')))
        
        # 1. Thêm hoặc cập nhật bảng students
        cur.execute("""
        INSERT OR REPLACE INTO students (student_id, full_name, cohort, class_name)
        VALUES (?, ?, ?, ?);
        """, (s_id, name, cohort, c_name))
        
        # 2. Thêm hoặc cập nhật bảng assessments
        score_mieng = float(row.get('score_mieng', 7.0))
        score_15p = float(row.get('score_15p', 7.0))
        score_45p = float(row.get('score_45p', 6.5))
        score_hk1 = float(row.get('score_hk1', 6.5))
        score_hk2 = float(row.get('score_hk2', 7.0))
        score_tx_avg = float(row.get('score_tx_avg', round((score_mieng + score_15p + score_45p * 2) / 4.0, 1)))
        hw = float(row.get('homework_ratio', 0.8))
        rec = float(row.get('score_recall', 7.0))
        und = float(row.get('score_understand', 6.5))
        app = float(row.get('score_apply', 5.5))
        happ = float(row.get('score_high_apply', 4.0))
        cs = float(row.get('score_topic_cs', 5.5))
        ict = float(row.get('score_topic_ict', 6.5))
        dl = float(row.get('score_topic_dl', 7.5))
        trend = float(row.get('score_trend', round(score_hk2 - score_hk1, 2)))
        val_est = row.get('est_final_score', None)
        if isinstance(val_est, (pd.Series, np.ndarray, list)):
            val_est = val_est.iloc[0] if hasattr(val_est, 'iloc') else val_est[0]
        est_score = float(val_est) if pd.notnull(val_est) else None
        
        val_risk = row.get('target_risk_level', None)
        if isinstance(val_risk, (pd.Series, np.ndarray, list)):
            val_risk = val_risk.iloc[0] if hasattr(val_risk, 'iloc') else val_risk[0]
        target_risk = int(val_risk) if pd.notnull(val_risk) else None
        
        # Xóa đánh giá cũ nếu có để nạp mới
        cur.execute("DELETE FROM assessments WHERE student_id = ?;", (s_id,))
        cur.execute("""
        INSERT INTO assessments (
            student_id, score_mieng, score_15p, score_45p, score_hk1, score_hk2,
            score_tx_avg, homework_ratio, score_recall, score_understand,
            score_apply, score_high_apply, score_topic_cs, score_topic_ict,
            score_topic_dl, score_trend, est_final_score, target_risk_level
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            s_id, score_mieng, score_15p, score_45p, score_hk1, score_hk2,
            score_tx_avg, hw, rec, und, app, happ, cs, ict, dl, trend,
            est_score, target_risk
        ))
        inserted_count += 1
        
    conn.commit()
    conn.close()
    return inserted_count

def get_cohort_list(db_path: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Lấy danh sách các khóa học hiện có trong CSDL và số lượng học sinh mỗi khóa.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("""
    SELECT cohort, COUNT(student_id) as total_students
    FROM students
    GROUP BY cohort
    ORDER BY cohort ASC;
    """)
    rows = cur.fetchall()
    conn.close()
    return [{'cohort': r['cohort'], 'total_students': r['total_students']} for r in rows]

def get_cohort_dataframe(cohort_name: str, db_path: Optional[str] = None, in_vietnamese: bool = True) -> pd.DataFrame:
    """
    Truy vấn toàn bộ học sinh và điểm số của một khóa cụ thể từ CSDL SQLite.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    search_param = cohort_name
    if cohort_name.upper().startswith('K') and cohort_name[1:].isdigit():
        k_num = int(cohort_name[1:])
        if k_num < 100:
            search_param = f"20{k_num}"
            
    query = """
    SELECT 
        s.student_id, s.full_name, s.cohort, s.class_name,
        a.score_mieng, a.score_15p, a.score_45p, a.score_hk1, a.score_hk2,
        a.score_tx_avg, a.homework_ratio,
        a.score_recall, a.score_understand, a.score_apply, a.score_high_apply,
        a.score_topic_cs, a.score_topic_ict, a.score_topic_dl, a.score_trend,
        a.est_final_score, a.target_risk_level
    FROM students s
    JOIN assessments a ON s.student_id = a.student_id
    WHERE s.cohort = ? OR s.cohort LIKE ?
    ORDER BY s.student_id ASC;
    """
    df = pd.read_sql_query(query, conn, params=(cohort_name, f"%{search_param}%"))
    conn.close()
    
    if in_vietnamese:
        vn_rename = {
            'student_id': 'Mã Học Sinh',
            'full_name': 'Họ Và Tên',
            'cohort': 'Khóa Học',
            'class_name': 'Lớp',
            'score_mieng': 'Điểm Miệng',
            'score_15p': 'Điểm 15 Phút',
            'score_45p': 'Điểm 45 Phút',
            'score_hk1': 'Điểm Học Kỳ 1',
            'score_hk2': 'Điểm Học Kỳ 2',
            'score_tx_avg': 'Điểm Trung Bình Quá Trình',
            'homework_ratio': 'Tỷ Lệ Hoàn Thành BTVN',
            'score_recall': 'Điểm Nhận Biết',
            'score_understand': 'Điểm Thông Hiểu',
            'score_apply': 'Điểm Vận Dụng',
            'score_high_apply': 'Điểm Vận Dụng Cao',
            'score_topic_cs': 'Điểm Khoa Học Máy Tính (CS)',
            'score_topic_ict': 'Điểm Tin Học Ứng Dụng (ICT)',
            'score_topic_dl': 'Điểm Đạo Đức Số (DL)',
            'score_trend': 'Xu Hướng Điểm Số',
            'est_final_score': 'Điểm Thi Tốt Nghiệp Dự Kiến',
            'target_risk_level': 'Mã Mức Nguy Cơ'
        }
        df = df.rename(columns=vn_rename)
        risk_map = {0: 'An toàn (>= 6.5)', 1: 'Cần theo dõi (5.0 - 6.4)', 2: 'Nguy cơ cao (< 5.0)'}
        df['Mức Độ Nguy Cơ'] = df['Mã Mức Nguy Cơ'].map(risk_map)
        
    return df

def get_all_students_dataframe(db_path: Optional[str] = None, in_vietnamese: bool = True) -> pd.DataFrame:
    """
    Truy vấn toàn bộ học sinh và điểm số của tất cả các khóa từ CSDL SQLite.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    query = """
    SELECT 
        s.student_id, s.full_name, s.cohort, s.class_name,
        a.score_mieng, a.score_15p, a.score_45p, a.score_hk1, a.score_hk2,
        a.score_tx_avg, a.homework_ratio,
        a.score_recall, a.score_understand, a.score_apply, a.score_high_apply,
        a.score_topic_cs, a.score_topic_ict, a.score_topic_dl, a.score_trend,
        a.est_final_score, a.target_risk_level
    FROM students s
    JOIN assessments a ON s.student_id = a.student_id
    ORDER BY s.student_id ASC;
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if in_vietnamese:
        vn_rename = {
            'student_id': 'Mã Học Sinh',
            'full_name': 'Họ Và Tên',
            'cohort': 'Khóa Học',
            'class_name': 'Lớp',
            'score_mieng': 'Điểm Miệng',
            'score_15p': 'Điểm 15 Phút',
            'score_45p': 'Điểm 45 Phút',
            'score_hk1': 'Điểm Học Kỳ 1',
            'score_hk2': 'Điểm Học Kỳ 2',
            'score_tx_avg': 'Điểm Trung Bình Quá Trình',
            'homework_ratio': 'Tỷ Lệ Hoàn Thành BTVN',
            'score_recall': 'Điểm Nhận Biết',
            'score_understand': 'Điểm Thông Hiểu',
            'score_apply': 'Điểm Vận Dụng',
            'score_high_apply': 'Điểm Vận Dụng Cao',
            'score_topic_cs': 'Điểm Khoa Học Máy Tính (CS)',
            'score_topic_ict': 'Điểm Tin Học Ứng Dụng (ICT)',
            'score_topic_dl': 'Điểm Đạo Đức Số (DL)',
            'score_trend': 'Xu Hướng Điểm Số',
            'est_final_score': 'Điểm Thi Tốt Nghiệp Dự Kiến',
            'target_risk_level': 'Mã Mức Nguy Cơ'
        }
        df = df.rename(columns=vn_rename)
        risk_map = {0: 'An toàn (>= 6.5)', 1: 'Cần theo dõi (5.0 - 6.4)', 2: 'Nguy cơ cao (< 5.0)'}
        df['Mức Độ Nguy Cơ'] = df['Mã Mức Nguy Cơ'].map(risk_map)
        
    return df

def log_prediction_result(
    student_id: str,
    model_used: str,
    predicted_risk: int,
    risk_label: str,
    high_risk_prob: float,
    primary_gap: str,
    recommendation: str,
    db_path: Optional[str] = None
):
    """
    Ghi nhật ký kết quả dự báo của học sinh vào CSDL để lưu vết và theo dõi tiến độ.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO prediction_logs (
        student_id, model_used, predicted_risk, risk_label,
        high_risk_prob, primary_gap, recommendation
    ) VALUES (?, ?, ?, ?, ?, ?, ?);
    """, (
        student_id, model_used, predicted_risk, risk_label,
        high_risk_prob, primary_gap, recommendation
    ))
    conn.commit()
    conn.close()

def get_database_statistics(db_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Tổng hợp các chỉ số hoạt động của CSDL SQLite.
    """
    init_database(db_path)
    conn = get_connection(db_path)
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM students;")
    total_students = cur.fetchone()[0]
    
    cur.execute("SELECT COUNT(DISTINCT cohort) FROM students;")
    total_cohorts = cur.fetchone()[0]
    
    cur.execute("SELECT COUNT(*) FROM assessments;")
    total_assessments = cur.fetchone()[0]
    
    cur.execute("SELECT COUNT(*) FROM prediction_logs;")
    total_logs = cur.fetchone()[0]
    
    conn.close()
    return {
        'total_students': total_students,
        'total_cohorts': total_cohorts,
        'total_assessments': total_assessments,
        'total_logs': total_logs
    }

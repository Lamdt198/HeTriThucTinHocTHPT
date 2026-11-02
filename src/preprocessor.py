"""
preprocessor.py
Module tiền xử lý dữ liệu và kiểm soát rò rỉ dữ liệu (Anti-Leakage Data Pipeline)
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import joblib
import pandas as pd
import numpy as np
from typing import Tuple, List, Dict, Any
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer

FEATURE_COLUMNS = [
    'score_mieng',
    'score_15p',
    'score_45p',
    'score_hk1',
    'score_hk2',
    'score_tx_avg',
    'homework_ratio',
    'score_recall',
    'score_understand',
    'score_apply',
    'score_high_apply',
    'score_topic_cs',
    'score_topic_ict',
    'score_topic_dl',
    'score_trend'
]

TARGET_COLUMN = 'target_risk_level'

CLASS_NAMES = {
    0: 'An toàn (Safe)',
    1: 'Cần theo dõi (Moderate Risk)',
    2: 'Nguy cơ cao (High Risk)'
}

VIETNAMESE_COLUMN_MAPPING = {
    'Mã Học Sinh': 'student_id',
    'Mã học sinh': 'student_id',
    'Họ Và Tên': 'full_name',
    'Họ và tên': 'full_name',
    'Khóa': 'cohort',
    'Khóa Học': 'cohort',
    'Niên Khóa': 'cohort',
    'Lớp': 'class_name',
    'Điểm Miệng': 'score_mieng',
    'Điểm 15 Phút': 'score_15p',
    'Điểm 15p': 'score_15p',
    'Điểm 45 Phút': 'score_45p',
    'Điểm 45p': 'score_45p',
    'Điểm 1 Tiết': 'score_45p',
    'Điểm Học Kỳ 1': 'score_hk1',
    'Điểm HK1': 'score_hk1',
    'Điểm Học Kỳ 2': 'score_hk2',
    'Điểm HK2': 'score_hk2',
    'Điểm Trung Bình Quá Trình': 'score_tx_avg',
    'Điểm Trung Bình Thường Xuyên': 'score_tx_avg',
    'Điểm Thường Xuyên 1': 'score_mieng',
    'Điểm Thường Xuyên 2': 'score_15p',
    'Điểm Giữa Kỳ': 'score_45p',
    'Tỷ Lệ Hoàn Thành BTVN': 'homework_ratio',
    'Tỷ Lệ BTVN': 'homework_ratio',
    'Tỷ Lệ Hoàn Thành Bài Tập': 'homework_ratio',
    'Điểm Nhận Biết': 'score_recall',
    'Điểm Thông Hiểu': 'score_understand',
    'Điểm Vận Dụng': 'score_apply',
    'Điểm Vận Dụng Cao': 'score_high_apply',
    'Điểm Khoa Học Máy Tính (CS)': 'score_topic_cs',
    'Điểm Khoa Học Máy Tính': 'score_topic_cs',
    'Điểm Tin Học Ứng Dụng (ICT)': 'score_topic_ict',
    'Điểm Tin Học Ứng Dụng': 'score_topic_ict',
    'Điểm Đạo Đức Số (DL)': 'score_topic_dl',
    'Điểm Đạo Đức Số': 'score_topic_dl',
    'Xu Hướng Điểm Số': 'score_trend',
    'Xu Hướng Điểm': 'score_trend',
    'Điểm Thi Tốt Nghiệp Dự Kiến': 'est_final_score',
    'Mức Độ Nguy Cơ': 'risk_label',
    'Mã Mức Nguy Cơ': 'target_risk_level'
}

def standardize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hóa tiêu đề cột nếu đầu vào là tiếng Việt sang tên thuộc tính nội bộ tiếng Anh.
    """
    rename_dict = {col: VIETNAMESE_COLUMN_MAPPING[col] for col in df.columns if col in VIETNAMESE_COLUMN_MAPPING}
    df_std = df.rename(columns=rename_dict) if rename_dict else df.copy()
    
    # Tự động tính toán các cột phái sinh nếu thiếu
    if 'score_tx_avg' not in df_std.columns and all(c in df_std.columns for c in ['score_mieng', 'score_15p', 'score_45p']):
        df_std['score_tx_avg'] = np.round((df_std['score_mieng'] + df_std['score_15p'] + df_std['score_45p'] * 2) / 4.0, 1)
    if 'score_trend' not in df_std.columns and all(c in df_std.columns for c in ['score_hk1', 'score_hk2']):
        df_std['score_trend'] = np.round(df_std['score_hk2'] - df_std['score_hk1'], 2)
        
    return df_std

def validate_input(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """
    Kiểm tra tính hợp lệ của dữ liệu đầu vào.
    Hỗ trợ cả tiêu đề tiếng Anh lẫn tiếng Việt viết hoa.
    """
    df_std = standardize_dataframe(df)
    errors = []
    
    # 1. Kiểm tra thiếu cột
    missing_cols = [col for col in FEATURE_COLUMNS if col not in df_std.columns]
    if missing_cols:
        rev_map = {
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
            'score_trend': 'Xu Hướng Điểm Số'
        }
        missing_vn = [rev_map.get(c, c) for c in missing_cols]
        errors.append(f"Dữ liệu thiếu các cột bắt buộc: {', '.join(missing_vn)}")
        return False, errors
        
    # 2. Kiểm tra dải điểm hợp lệ [0, 10]
    score_cols = [c for c in FEATURE_COLUMNS if c not in ['homework_ratio', 'score_trend']]
    for col in score_cols:
        invalid_vals = df_std[df_std[col].notnull() & ((df_std[col] < 0.0) | (df_std[col] > 10.0))]
        if len(invalid_vals) > 0:
            errors.append(f"Cột '{col}' chứa {len(invalid_vals)} giá trị điểm ngoài khoảng [0, 10].")
            
    # 3. Kiểm tra homework_ratio [0, 1]
    invalid_hw = df_std[df_std['homework_ratio'].notnull() & ((df_std['homework_ratio'] < 0.0) | (df_std['homework_ratio'] > 1.0))]
    if len(invalid_hw) > 0:
        errors.append(f"Cột tỷ lệ hoàn thành BTVN chứa {len(invalid_hw)} giá trị ngoài khoảng [0.0, 1.0].")
        
    return len(errors) == 0, errors

class AssessmentPreprocessor:
    """
    Wrapper tiền xử lý sử dụng dict chuẩn hóa để tránh lỗi pickle.
    """
    def __init__(self):
        self.feature_columns = FEATURE_COLUMNS
        self.imputer = SimpleImputer(strategy='median')
        self.scaler = StandardScaler()
        self.is_fitted = False

    def fit_transform(self, X: pd.DataFrame) -> np.ndarray:
        X_std = standardize_dataframe(X)
        X_feats = X_std[self.feature_columns].copy()
        X_imputed = self.imputer.fit_transform(X_feats)
        X_scaled = self.scaler.fit_transform(X_imputed)
        self.is_fitted = True
        return X_scaled

    def transform(self, X: pd.DataFrame) -> np.ndarray:
        if not self.is_fitted:
            raise ValueError("Preprocessor chưa được fit.")
        X_std = standardize_dataframe(X)
        X_feats = X_std[self.feature_columns].copy()
        X_imputed = self.imputer.transform(X_feats)
        X_scaled = self.scaler.transform(X_imputed)
        return X_scaled

    def save(self, filepath: str):
        payload = {
            'imputer': self.imputer,
            'scaler': self.scaler,
            'feature_columns': self.feature_columns,
            'is_fitted': self.is_fitted
        }
        joblib.dump(payload, filepath)

    @classmethod
    def load(cls, filepath: str):
        payload = joblib.load(filepath)
        obj = cls()
        obj.imputer = payload['imputer']
        obj.scaler = payload['scaler']
        obj.feature_columns = payload['feature_columns']
        obj.is_fitted = payload['is_fitted']
        return obj

def process_and_split_data(raw_csv_path: str, output_dir: str, save_preprocessor_path: str):
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.dirname(save_preprocessor_path), exist_ok=True)
    
    df = standardize_dataframe(pd.read_csv(raw_csv_path))
    print(f"[INFO] Doc du lieu: {df.shape[0]} dong, {df.shape[1]} cot")
    
    train_df, temp_df = train_test_split(
        df, test_size=0.30, random_state=42, stratify=df[TARGET_COLUMN]
    )
    val_df, test_df = train_test_split(
        temp_df, test_size=0.50, random_state=42, stratify=temp_df[TARGET_COLUMN]
    )
    
    print(f"-> Tap Train: {train_df.shape[0]} mau ({train_df.shape[0]/len(df):.1%})")
    print(f"-> Tap Val:   {val_df.shape[0]} mau ({val_df.shape[0]/len(df):.1%})")
    print(f"-> Tap Test:  {test_df.shape[0]} mau ({test_df.shape[0]/len(df):.1%})")
    
    preprocessor = AssessmentPreprocessor()
    preprocessor.fit_transform(train_df)
    preprocessor.save(save_preprocessor_path)
    
    train_df.to_csv(os.path.join(output_dir, 'train.csv'), index=False)
    val_df.to_csv(os.path.join(output_dir, 'val.csv'), index=False)
    test_df.to_csv(os.path.join(output_dir, 'test.csv'), index=False)
    print(f"[OK] Luu du lieu va preprocessor thanh cong.")

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    raw_path = os.path.join(base_dir, 'data', 'raw', 'informatics_assessment_raw.csv')
    out_dir = os.path.join(base_dir, 'data', 'processed')
    prep_path = os.path.join(base_dir, 'models_saved', 'preprocessor.joblib')
    process_and_split_data(raw_path, out_dir, prep_path)

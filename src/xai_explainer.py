"""
xai_explainer.py
Module Giải thích Mô hình Học máy (Explainable AI - XAI).
Kết hợp Permutation Importance (Toàn cục) và Feature Contribution Attribution (Cục bộ từng cá nhân).
Phù hợp 100% với môi trường học thuật, độc lập và chạy nhanh, trực quan.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
try:
    from sklearn.inspection._permutation_importance import permutation_importance
except Exception:
    permutation_importance = None

from preprocessor import FEATURE_COLUMNS, TARGET_COLUMN, AssessmentPreprocessor

FEATURE_DISPLAY_NAMES = {
    'score_mieng': 'Điểm Miệng',
    'score_15p': 'Điểm 15 Phút',
    'score_45p': 'Điểm 45 Phút (1 Tiết)',
    'score_hk1': 'Điểm Học Kỳ 1',
    'score_hk2': 'Điểm Học Kỳ 2',
    'score_tx_avg': 'Điểm TB Quá Trình',
    'homework_ratio': 'Tỷ Lệ Nộp BTVN (LMS)',
    'score_recall': 'Mức 1: Nhận Biết',
    'score_understand': 'Mức 2: Thông Hiểu',
    'score_apply': 'Mức 3: Vận Dụng',
    'score_high_apply': 'Mức 4: Vận Dụng Cao',
    'score_topic_cs': 'Mạch CS (Khoa Học MT)',
    'score_topic_ict': 'Mạch ICT (Ứng Dụng Tin)',
    'score_topic_dl': 'Mạch DL (Đạo Đức Số)',
    'score_trend': 'Xu Hướng Điểm Số (HK2 - HK1)'
}

def compute_and_save_global_xai():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    models_dir = os.path.join(base_dir, 'models_saved')
    fig_dir = os.path.join(base_dir, 'reports', 'figures')
    processed_dir = os.path.join(base_dir, 'data', 'processed')
    
    os.makedirs(fig_dir, exist_ok=True)
    
    # 1. Load model va test data
    preprocessor = AssessmentPreprocessor.load(os.path.join(models_dir, 'preprocessor.joblib'))
    rf_model = joblib.load(os.path.join(models_dir, 'improved_rf.joblib'))
    test_df = pd.read_csv(os.path.join(processed_dir, 'test.csv'))
    train_df = pd.read_csv(os.path.join(processed_dir, 'train.csv'))
    
    X_test_scaled = preprocessor.transform(test_df)
    y_test = test_df[TARGET_COLUMN].values
    
    # 2. Tinh toan Feature Importance tu Random Forest
    rf_importances = rf_model.feature_importances_
    
    # 3. Tinh toan Permutation Importance tren Tap Test
    print("[INFO] Dang tinh toan Permutation Importance tren tap Test...")
    perm_result = permutation_importance(
        rf_model, X_test_scaled, y_test,
        n_repeats=10, random_state=42, scoring='f1_macro'
    )
    perm_importances = perm_result.importances_mean
    
    # Tao DataFrame tong hop
    importance_df = pd.DataFrame({
        'Feature_Code': FEATURE_COLUMNS,
        'Feature_Name': [FEATURE_DISPLAY_NAMES.get(c, c) for c in FEATURE_COLUMNS],
        'RF_Gini_Importance': rf_importances,
        'Permutation_Importance': perm_importances
    }).sort_values(by='Permutation_Importance', ascending=False)
    
    importance_path = os.path.join(models_dir, 'feature_importance.csv')
    importance_df.to_csv(importance_path, index=False)
    print(f"[OK] Da luu Feature Importance vao: {importance_path}")
    
    # 4. Tinh nguong trung binh va do lech chuan cua nhom An Toan (Class 0) de lam moc Benchmark giai thich ca nhan
    safe_students = train_df[train_df[TARGET_COLUMN] == 0]
    safe_stats = {
        'mean': safe_students[FEATURE_COLUMNS].mean().to_dict(),
        'std': safe_students[FEATURE_COLUMNS].std().replace(0, 1.0).to_dict()
    }
    joblib.dump(safe_stats, os.path.join(models_dir, 'safe_benchmark_stats.joblib'))
    print(f"[OK] Da luu chi so Benchmark nhom An Toan phuc vu XAI ca nhan.")
    
    # 5. Ve bieu do Permutation Feature Importance
    plt.figure(figsize=(10, 6.5))
    sns.barplot(
        x='Permutation_Importance', y='Feature_Name',
        data=importance_df, palette='Blues_r'
    )
    plt.title('Độ Quan Trọng Thuộc Tính (Permutation Feature Importance - F1 Macro trên Test)', fontsize=12, pad=12)
    plt.xlabel('Mức Giảm Điểm F1-Macro Khi Xáo Trộn Thuộc Tính', fontsize=11)
    plt.ylabel('Thuộc Tính Đánh Giá', fontsize=11)
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    chart_path = os.path.join(fig_dir, 'feature_importance_ranking.png')
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"[OK] Da luu bieu do Feature Importance tai: {chart_path}")

def explain_student_risk(student_df_or_dict) -> dict:
    """
    Giai thich nguyen nhan du bao nguy co cho mot hoc sinh cu the.
    Tra ve:
    - risk_class: 0, 1, hoac 2
    - risk_label: chuoi mo ta
    - probabilities: xac suat 3 lop
    - risk_factors: DataFrame cac yeu to lam tang/giam nguy co
    - recommendation: loi khuyen su pham can thiep
    """
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    models_dir = os.path.join(base_dir, 'models_saved')
    
    preprocessor = AssessmentPreprocessor.load(os.path.join(models_dir, 'preprocessor.joblib'))
    rf_model = joblib.load(os.path.join(models_dir, 'improved_rf.joblib'))
    safe_stats = joblib.load(os.path.join(models_dir, 'safe_benchmark_stats.joblib'))
    
    if isinstance(student_df_or_dict, dict):
        df_single = pd.DataFrame([student_df_or_dict])
    else:
        df_single = student_df_or_dict.copy()
        
    X_scaled = preprocessor.transform(df_single)
    probs = rf_model.predict_proba(X_scaled)[0]
    pred_class = int(np.argmax(probs))
    
    # Tinh toan do dong gop cua tung thuoc tinh:
    # Attribution = Importance_weight * (mean_safe - actual_val) / std_safe
    # Neu Attribution > 0: Dac trung nay thap hon chuan an toan => DAY NGUY CO CAO LEN
    # Neu Attribution < 0: Dac trung nay cao hon chuan an toan => YEU TO BAO VE
    weights = rf_model.feature_importances_
    contributions = []
    
    for i, col in enumerate(FEATURE_COLUMNS):
        val = df_single[col].iloc[0]
        mean_safe = safe_stats['mean'][col]
        std_safe = safe_stats['std'][col]
        z_gap = (mean_safe - val) / std_safe
        attr_score = float(weights[i] * z_gap)
        
        contributions.append({
            'Feature_Code': col,
            'Feature_Name': FEATURE_DISPLAY_NAMES.get(col, col),
            'Student_Value': round(float(val), 2),
            'Safe_Benchmark_Avg': round(float(mean_safe), 2),
            'Risk_Attribution_Score': round(attr_score, 4),
            'Impact_Type': 'Làm tăng nguy cơ' if attr_score > 0 else 'Yếu tố an toàn/bảo vệ'
        })
        
    contrib_df = pd.DataFrame(contributions).sort_values(by='Risk_Attribution_Score', ascending=False)
    
    # Sinh loi khuyen su pham
    top_risk_drivers = contrib_df[contrib_df['Risk_Attribution_Score'] > 0].head(3)
    rec_lines = []
    if pred_class == 2:
        rec_lines.append("[CẢNH BÁO ĐỎ] Học sinh thuộc nhóm NGUY CƠ CAO (< 5.0 điểm tốt nghiệp).")
        for _, row in top_risk_drivers.iterrows():
            rec_lines.append(f"- Điểm yếu cốt lõi: {row['Feature_Name']} đạt {row['Student_Value']} (thấp hơn mức an toàn {row['Safe_Benchmark_Avg']}).")
        rec_lines.append("-> Hành động can thiệp: Đưa vào danh sách kèm cặp trọng điểm; bổ túc gấp phần Vận dụng cao hoặc Chuyên đề Lập trình (CS).")
    elif pred_class == 1:
        rec_lines.append("[CẢNH BÁO VÀNG] Học sinh thuộc nhóm CẦN THEO DÕI (5.0 - 6.4 điểm).")
        for _, row in top_risk_drivers.iterrows():
            rec_lines.append(f"- Điểm cần cải thiện: {row['Feature_Name']} hiện đạt {row['Student_Value']}.")
        rec_lines.append("-> Hành động can thiệp: Giao thêm bài tập rèn kỹ năng giải đề trắc nghiệm Đúng/Sai trên hệ thống LMS.")
    else:
        rec_lines.append("[AN TOÀN] Học sinh có nền tảng vững vàng (Dự kiến >= 6.5 điểm).")
        rec_lines.append("-> Khuyến khích tham gia nhóm ôn luyện nâng cao để chinh phục mức điểm 8.0 - 10.0.")
        
    return {
        'risk_class': pred_class,
        'risk_label': ['An toàn (>= 6.5)', 'Cần theo dõi (5.0 - 6.4)', 'Nguy cơ cao (< 5.0)'][pred_class],
        'probabilities': {
            'An toàn': round(float(probs[0]), 3),
            'Cần theo dõi': round(float(probs[1]), 3),
            'Nguy cơ cao': round(float(probs[2]), 3)
        },
        'contribution_df': contrib_df,
        'recommendation': "\n".join(rec_lines)
    }

if __name__ == '__main__':
    compute_and_save_global_xai()
    
    # Test thu 1 hoc sinh gia dinh
    sample_student = {
        'score_mieng': 4.0,
        'score_15p': 3.5,
        'score_45p': 4.0,
        'score_hk1': 4.5,
        'score_hk2': 4.0,
        'score_tx_avg': 4.0,
        'homework_ratio': 0.40,
        'score_recall': 6.0,
        'score_understand': 4.5,
        'score_apply': 3.0,
        'score_high_apply': 1.0,
        'score_topic_cs': 2.5,
        'score_topic_ict': 4.5,
        'score_topic_dl': 5.5,
        'score_trend': -0.5
    }
    result = explain_student_risk(sample_student)
    print("\n--- TEST EXPLAIN SINGLE STUDENT ---")
    print(f"Ket qua: {result['risk_label']}")
    print(f"Xac suat: {result['probabilities']}")
    print("Top 3 nguyen nhan rui ro:")
    print(result['contribution_df'][['Feature_Name', 'Student_Value', 'Safe_Benchmark_Avg', 'Risk_Attribution_Score']].head(3))
    print("\nKhuyen nghi su pham:")
    print(result['recommendation'])

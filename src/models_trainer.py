"""
models_trainer.py
Huấn luyện và so sánh Baseline vs Improved Models (XGBoost, LightGBM, Random Forest).
Đo lường thời gian suy luận, tính toán toàn diện các metrics và lưu trữ mô hình.
"""

import os
import sys
import time
import json
import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from xgboost import XGBClassifier
from lightgbm import LGBMClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    confusion_matrix, classification_report
)

from preprocessor import AssessmentPreprocessor, FEATURE_COLUMNS, TARGET_COLUMN

class RuleBasedBaseline:
    """
    Mo hinh co so dua tren quy tac su pham truyen thong:
    - Neu diem trung binh thuong xuyen < 5.0 hoac giua ky < 4.5 -> Nguy co cao (2)
    - Neu giua ky tu 4.5 den 6.0 -> Can theo doi (1)
    - Con lai -> An toan (0)
    """
    def __init__(self):
        pass
        
    def fit(self, X, y=None):
        return self
        
    def predict(self, df_or_X):
        # df_or_X co the la DataFrame raw hoac array
        # Neu la array: cot 2 la score_tx_avg, cot 3 la score_midterm
        if isinstance(df_or_X, pd.DataFrame):
            tx_avg = df_or_X.get('score_tx_avg', df_or_X.get('Điểm Trung Bình Quá Trình', 6.0)).values
            hk2 = df_or_X.get('score_hk2', df_or_X.get('Điểm Học Kỳ 2', 6.0)).values
        else:
            tx_avg = df_or_X[:, 5] if df_or_X.shape[1] > 5 else df_or_X[:, 0]
            hk2 = df_or_X[:, 4] if df_or_X.shape[1] > 4 else df_or_X[:, 0]
            
        preds = np.zeros(len(tx_avg), dtype=int)
        preds[(hk2 >= 4.5) & (hk2 < 6.5)] = 1
        preds[(tx_avg < 5.0) | (hk2 < 4.5)] = 2
        return preds
        
    def predict_proba(self, df_or_X):
        preds = self.predict(df_or_X)
        probas = np.zeros((len(preds), 3))
        for i, p in enumerate(preds):
            probas[i, p] = 0.8
            other_indices = [c for c in [0, 1, 2] if c != p]
            probas[i, other_indices[0]] = 0.1
            probas[i, other_indices[1]] = 0.1
        return probas

def train_and_benchmark_all():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    processed_dir = os.path.join(base_dir, 'data', 'processed')
    models_dir = os.path.join(base_dir, 'models_saved')
    fig_dir = os.path.join(base_dir, 'reports', 'figures')
    
    os.makedirs(models_dir, exist_ok=True)
    os.makedirs(fig_dir, exist_ok=True)
    
    # 1. Load data
    train_df = pd.read_csv(os.path.join(processed_dir, 'train.csv'))
    val_df = pd.read_csv(os.path.join(processed_dir, 'val.csv'))
    test_df = pd.read_csv(os.path.join(processed_dir, 'test.csv'))
    
    preprocessor = AssessmentPreprocessor.load(os.path.join(models_dir, 'preprocessor.joblib'))
    
    X_train_scaled = preprocessor.transform(train_df)
    y_train = train_df[TARGET_COLUMN].values
    
    X_val_scaled = preprocessor.transform(val_df)
    y_val = val_df[TARGET_COLUMN].values
    
    X_test_scaled = preprocessor.transform(test_df)
    y_test = test_df[TARGET_COLUMN].values
    
    # 2. Dinh nghia cac mo hinh
    models = {
        'Baseline: Rule-Based Heuristic': RuleBasedBaseline(),
        'Baseline: Logistic Regression': LogisticRegression(
            class_weight='balanced', max_iter=1000, random_state=42
        ),
        'Baseline: Decision Tree': DecisionTreeClassifier(
            max_depth=4, class_weight='balanced', random_state=42
        ),
        'Improved: Random Forest': RandomForestClassifier(
            n_estimators=150, max_depth=8, class_weight='balanced', random_state=42
        ),
        'Improved: XGBoost': XGBClassifier(
            n_estimators=150, max_depth=5, learning_rate=0.08,
            objective='multi:softprob', num_class=3, random_state=42,
            eval_metric='mlogloss'
        ),
        'Improved: LightGBM': LGBMClassifier(
            n_estimators=150, max_depth=5, learning_rate=0.08,
            class_weight='balanced', random_state=42, verbose=-1
        )
    }
    
    benchmark_results = []
    trained_models = {}
    
    print("[INFO] Bat dau Huan luyen va Danh gia tren Tap Test doc lap (15% - 611 mau)...")
    print("-" * 85)
    print(f"{'Ten Mo Hinh':<32} | {'Acc':<6} | {'Macro F1':<8} | {'Recall High':<11} | {'Time (ms)':<9}")
    print("-" * 85)
    
    for name, model in models.items():
        # Do thoi gian huan luyen
        t0 = time.time()
        if 'Rule-Based' in name:
            # Rule based khong can fit
            pass
        else:
            model.fit(X_train_scaled, y_train)
        train_time = time.time() - t0
        
        # Do thoi gian suy luan tren tap Test
        t_inf_0 = time.time()
        if 'Rule-Based' in name:
            y_pred = model.predict(test_df)
        else:
            y_pred = model.predict(X_test_scaled)
        inf_time_total = time.time() - t_inf_0
        latency_ms_per_sample = (inf_time_total / len(y_test)) * 1000.0
        
        # Tinh cac metrics
        acc = accuracy_score(y_test, y_pred)
        macro_prec = precision_score(y_test, y_pred, average='macro', zero_division=0)
        macro_rec = recall_score(y_test, y_pred, average='macro', zero_division=0)
        macro_f1 = f1_score(y_test, y_pred, average='macro', zero_division=0)
        weighted_f1 = f1_score(y_test, y_pred, average='weighted', zero_division=0)
        
        # Chi so quan trong nhat: Recall lop Nguy co cao (Class 2)
        rec_high = recall_score(y_test, y_pred, labels=[2], average=None, zero_division=0)[0]
        prec_high = precision_score(y_test, y_pred, labels=[2], average=None, zero_division=0)[0]
        f1_high = f1_score(y_test, y_pred, labels=[2], average=None, zero_division=0)[0]
        
        benchmark_results.append({
            'Model': name,
            'Accuracy': round(acc, 4),
            'Macro_Precision': round(macro_prec, 4),
            'Macro_Recall': round(macro_rec, 4),
            'Macro_F1': round(macro_f1, 4),
            'Weighted_F1': round(weighted_f1, 4),
            'Recall_High_Risk': round(rec_high, 4),
            'Precision_High_Risk': round(prec_high, 4),
            'F1_High_Risk': round(f1_high, 4),
            'Train_Time_Sec': round(train_time, 4),
            'Inference_Latency_ms': round(latency_ms_per_sample, 4)
        })
        
        trained_models[name] = model
        print(f"{name:<32} | {acc:.4f} | {macro_f1:.4f}   | {rec_high:.4f}      | {latency_ms_per_sample:.3f} ms")
        
    print("-" * 85)
    
    # 3. Luu bang ket qua
    benchmark_df = pd.DataFrame(benchmark_results)
    benchmark_csv = os.path.join(models_dir, 'benchmark_results.csv')
    benchmark_json = os.path.join(models_dir, 'benchmark_results.json')
    benchmark_df.to_csv(benchmark_csv, index=False)
    benchmark_df.to_json(benchmark_json, orient='records', indent=2)
    print(f"[OK] Da luu bang so sanh tai: {benchmark_csv}")
    
    # 4. Luu cac mo hinh
    joblib.dump(trained_models['Baseline: Logistic Regression'], os.path.join(models_dir, 'baseline_logistic.joblib'))
    joblib.dump(trained_models['Baseline: Decision Tree'], os.path.join(models_dir, 'baseline_tree.joblib'))
    joblib.dump(trained_models['Improved: XGBoost'], os.path.join(models_dir, 'improved_xgboost.joblib'))
    joblib.dump(trained_models['Improved: LightGBM'], os.path.join(models_dir, 'improved_lightgbm.joblib'))
    joblib.dump(trained_models['Improved: Random Forest'], os.path.join(models_dir, 'improved_rf.joblib'))
    print(f"[OK] Da luu tat ca model weights vao: {models_dir}")
    
    # 5. Ve va luu Confusion Matrix cua Best Model (LightGBM hoac XGBoost)
    best_model = trained_models['Improved: LightGBM']
    y_pred_best = best_model.predict(X_test_scaled)
    cm = confusion_matrix(y_test, y_pred_best)
    
    plt.figure(figsize=(7, 5.5))
    class_labels = ['An toàn (0)', 'Cần theo dõi (1)', 'Nguy cơ cao (2)']
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=class_labels, yticklabels=class_labels)
    plt.title('Ma Trận Nhầm Lẫn (Confusion Matrix) - LightGBM (Tập Test)', fontsize=12, pad=12)
    plt.xlabel('Nhãn Dự Đoán (Predicted)', fontsize=11)
    plt.ylabel('Nhãn Thực Tế (Ground Truth)', fontsize=11)
    plt.tight_layout()
    cm_path = os.path.join(fig_dir, 'confusion_matrix_best_model.png')
    plt.savefig(cm_path, dpi=300)
    plt.close()
    print(f"[OK] Da luu bieu do Confusion Matrix tai: {cm_path}")

    # 6. Ve bieu do so sanh Macro F1 & Recall High Risk giua cac mo hinh
    fig, ax = plt.subplots(figsize=(10, 5))
    x = np.arange(len(benchmark_df))
    width = 0.35
    ax.bar(x - width/2, benchmark_df['Macro_F1'], width, label='Macro F1-Score', color='#2563eb')
    ax.bar(x + width/2, benchmark_df['Recall_High_Risk'], width, label='Recall (Lớp Nguy cơ cao)', color='#dc2626')
    ax.set_ylabel('Điểm số (0.0 - 1.0)', fontsize=11)
    ax.set_title('So sánh Hiệu Năng: Baseline vs Mô hình Cải Tiến', fontsize=13, pad=12)
    ax.set_xticks(x)
    ax.set_xticklabels([m.replace(': ', '\n') for m in benchmark_df['Model']], rotation=0, fontsize=9)
    ax.set_ylim(0.0, 1.05)
    ax.legend(loc='lower right')
    ax.grid(axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    comp_path = os.path.join(fig_dir, 'model_comparison_chart.png')
    plt.savefig(comp_path, dpi=300)
    plt.close()
    print(f"[OK] Da luu bieu do so sanh tai: {comp_path}")

if __name__ == '__main__':
    train_and_benchmark_all()

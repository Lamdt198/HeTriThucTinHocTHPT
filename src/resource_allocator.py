"""
resource_allocator.py
Thuật toán phân bổ nguồn lực phụ đạo và can thiệp sư phạm.
Tự động gom nhóm học sinh có nguy cơ học tập theo lỗ hổng kiến thức chính
và đề xuất kịch bản phân công giáo viên, số tiết phụ đạo hợp lý.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import pandas as pd
import numpy as np
from typing import Dict, List, Any
from preprocessor import standardize_dataframe

def generate_remediation_plan(df_with_preds: pd.DataFrame) -> Dict[str, Any]:
    """
    Tạo kịch bản phân bổ nguồn lực phụ đạo dựa trên kết quả dự báo.
    
    Tham số:
    - df_with_preds: DataFrame chứa dữ liệu học sinh kèm cột 'predicted_risk_level' (0, 1, 2)
    """
    df = standardize_dataframe(df_with_preds.copy())
    
    if 'predicted_risk_level' not in df.columns:
        raise ValueError("DataFrame cần có cột 'predicted_risk_level' để lập kịch bản phân bổ.")
        
    total_students = len(df)
    safe_count = int((df['predicted_risk_level'] == 0).sum())
    moderate_count = int((df['predicted_risk_level'] == 1).sum())
    high_risk_count = int((df['predicted_risk_level'] == 2).sum())
    
    # Chỉ can thiệp vào học sinh Nguy cơ cao (2) và Cần theo dõi (1)
    at_risk_df = df[df['predicted_risk_level'].isin([1, 2])].copy()
    
    if len(at_risk_df) == 0:
        return {
            'total_students': total_students,
            'at_risk_count': 0,
            'summary': "Toàn bộ học sinh đều ở ngưỡng An toàn. Không cần mở lớp phụ đạo.",
            'groups': []
        }
        
    # Xác định lỗ hổng kiến thức lớn nhất của từng học sinh:
    # So sánh điểm CS, ICT, và Vận dụng cao
    def identify_primary_gap(row):
        scores = {
            'Lập trình & Thuật toán (CS)': row.get('score_topic_cs', 5.0),
            'Ứng dụng Tin học & CSDL (ICT)': row.get('score_topic_ict', 5.0),
            'Kỹ năng Vận dụng cao & Trả lời ngắn': row.get('score_high_apply', 4.0),
            'Nền tảng Lý thuyết & Đạo đức số (DL)': row.get('score_topic_dl', 6.0)
        }
        # Tìm mảng có điểm thấp nhất
        min_domain = min(scores, key=scores.get)
        return min_domain
        
    at_risk_df['primary_weakness'] = at_risk_df.apply(identify_primary_gap, axis=1)
    
    # Gom nhóm theo lỗ hổng kiến thức
    groups = []
    group_configs = [
        {
            'name': 'Nhóm Phụ đạo 1: Lập trình & Thuật toán (CS)',
            'weakness_key': 'Lập trình & Thuật toán (CS)',
            'teacher_role': 'Giáo viên phụ trách mảng Lập trình / Khoa học Máy tính',
            'recommended_hours_per_week': 3,
            'priority': 'Ưu tiên Cao (Cốt lõi chống điểm liệt)'
        },
        {
            'name': 'Nhóm Phụ đạo 2: Chinh phục Vận dụng cao & Trả lời ngắn',
            'weakness_key': 'Kỹ năng Vận dụng cao & Trả lời ngắn',
            'teacher_role': 'Giáo viên giàu kinh nghiệm luyện thi tốt nghiệp',
            'recommended_hours_per_week': 2,
            'priority': 'Ưu tiên Trung bình - Khá (Nâng điểm 6-8)'
        },
        {
            'name': 'Nhóm Phụ đạo 3: Ứng dụng Tin học, CSDL & Mạng (ICT)',
            'weakness_key': 'Ứng dụng Tin học & CSDL (ICT)',
            'teacher_role': 'Giáo viên thực hành Tin học ứng dụng',
            'recommended_hours_per_week': 2,
            'priority': 'Ưu tiên Trung bình'
        },
        {
            'name': 'Nhóm Phụ đạo 4: Củng cố Khái niệm Cốt lõi & Lý thuyết số (DL)',
            'weakness_key': 'Nền tảng Lý thuyết & Đạo đức số (DL)',
            'teacher_role': 'Giáo viên bộ môn Tin học',
            'recommended_hours_per_week': 1.5,
            'priority': 'Ưu tiên Tiêu chuẩn'
        }
    ]
    
    for cfg in group_configs:
        sub_df = at_risk_df[at_risk_df['primary_weakness'] == cfg['weakness_key']]
        count = len(sub_df)
        if count == 0:
            continue
            
        high_in_group = int((sub_df['predicted_risk_level'] == 2).sum())
        mod_in_group = int((sub_df['predicted_risk_level'] == 1).sum())
        
        # Đề xuất số lớp (mỗi lớp tối đa 15-20 học sinh)
        suggested_classes = max(1, int(np.ceil(count / 18)))
        total_hours = suggested_classes * cfg['recommended_hours_per_week']
        
        student_list = []
        for _, s in sub_df.head(25).iterrows():
            student_list.append({
                'Mã HS': s.get('student_id', 'N/A'),
                'Lớp': s.get('class_name', 'N/A'),
                'Mức nguy cơ': 'Nguy cơ cao' if s.get('predicted_risk_level') == 2 else 'Cần theo dõi',
                'Điểm HK2': s.get('score_hk2', s.get('Điểm Học Kỳ 2', 'N/A')),
                'Điểm yếu nhất': round(float(s.get('score_topic_cs' if 'CS' in cfg['weakness_key'] else 'score_high_apply', 0)), 1)
            })
            
        groups.append({
            'group_name': cfg['name'],
            'total_students': count,
            'high_risk_students': high_in_group,
            'moderate_students': mod_in_group,
            'suggested_classes': suggested_classes,
            'hours_per_week_per_class': cfg['recommended_hours_per_week'],
            'total_teaching_hours': total_hours,
            'assigned_teacher_role': cfg['teacher_role'],
            'priority': cfg['priority'],
            'sample_students': pd.DataFrame(student_list)
        })
        
    total_remediation_hours = sum(g['total_teaching_hours'] for g in groups)
    total_classes = sum(g['suggested_classes'] for g in groups)
    
    plan_summary = {
        'total_students': total_students,
        'safe_count': safe_count,
        'safe_pct': round(safe_count / total_students * 100, 1),
        'moderate_count': moderate_count,
        'moderate_pct': round(moderate_count / total_students * 100, 1),
        'high_risk_count': high_risk_count,
        'high_risk_pct': round(high_risk_count / total_students * 100, 1),
        'total_at_risk': len(at_risk_df),
        'total_suggested_classes': total_classes,
        'total_teaching_hours_per_week': total_remediation_hours,
        'groups': groups
    }
    
    return plan_summary

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    test_csv = os.path.join(base_dir, 'data', 'processed', 'test.csv')
    df = pd.read_csv(test_csv)
    # Giả định predicted_risk_level giống target_risk_level để test thuật toán
    df['predicted_risk_level'] = df['target_risk_level']
    
    plan = generate_remediation_plan(df)
    print("=== KỊCH BẢN PHÂN BỔ NGUỒN LỰC PHỤ ĐẠO MÔN TIN HỌC ===")
    print(f"Tổng số học sinh đánh giá: {plan['total_students']}")
    print(f"- An toàn: {plan['safe_count']} ({plan['safe_pct']}%)")
    print(f"- Cần theo dõi: {plan['moderate_count']} ({plan['moderate_pct']}%)")
    print(f"- Nguy cơ cao: {plan['high_risk_count']} ({plan['high_risk_pct']}%)")
    print(f"Tổng số lớp phụ đạo đề xuất: {plan['total_suggested_classes']} lớp")
    print(f"Tổng số tiết phụ đạo/tuần: {plan['total_teaching_hours_per_week']} tiết\n")
    
    for g in plan['groups']:
        print(f"[{g['group_name']}] - {g['total_students']} học sinh ({g['high_risk_students']} nguy cơ cao)")
        print(f"  -> {g['suggested_classes']} lớp, {g['hours_per_week_per_class']} tiết/tuần | {g['assigned_teacher_role']}")

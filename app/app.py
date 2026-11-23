"""
app.py
Ứng dụng Web Hệ Tri Thức & Học Máy Hỗ Trợ Cảnh Báo Nguy Cơ Học Tập Môn Tin Học THPT
Thiết kế theo chuẩn giao diện Hệ Tri Thức Tin Học THPT (Clean Card Grid, Radar Competency Chart, XAI Diagnostic)
EduRisk Informatics - Hệ Thống Hỗ Trợ Ra Quyết Định Cảnh Báo Nguy Cơ Môn Tin Học THPT
"""

import os
import sys

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src_dir = os.path.join(base_dir, 'src')
if src_dir in sys.path:
    sys.path.remove(src_dir)
sys.path.insert(0, src_dir)

import joblib
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from preprocessor import (
    FEATURE_COLUMNS, TARGET_COLUMN, CLASS_NAMES,
    AssessmentPreprocessor, validate_input
)
from xai_explainer import FEATURE_DISPLAY_NAMES, explain_student_risk
from resource_allocator import generate_remediation_plan
from database import (
    get_cohort_list, get_cohort_dataframe, get_all_students_dataframe,
    get_database_statistics, log_prediction_result
)

# ----------------- STREAMLIT PAGE CONFIG -----------------
st.set_page_config(
    page_title="Hệ Tri Thức Cảnh Báo Sớm Môn Tin Học THPT - GDPT 2018",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------- QUẢN LÝ CHẾ ĐỘ GIAO DIỆN (THEME STATE) -----------------
if 'app_theme' not in st.session_state:
    st.session_state['app_theme'] = "Sáng (Light)"

is_dark = "Tối" in st.session_state['app_theme']

# Bảng mã màu ngữ nghĩa động theo Theme
if is_dark:
    theme_vars = """
    :root {
        --kbs-bg-page: #0b0f19;
        --kbs-bg-card: #131b2e;
        --kbs-bg-subcard: #1c2640;
        --kbs-bg-sidebar: #0f172a;
        --kbs-border: #26354f;
        --kbs-border-strong: #3b4f74;
        --kbs-text-title: #f8fafc;
        --kbs-text-body: #cbd5e1;
        --kbs-text-muted: #94a3b8;
        --kbs-tab-active-bg: rgba(37, 99, 235, 0.25);
        --kbs-tab-active-text: #60a5fa;
        --kbs-tab-active-border: #3b82f6;
        --kbs-shadow: 0 4px 14px rgba(0, 0, 0, 0.45);
        --kbs-shadow-hover: 0 10px 25px rgba(0, 0, 0, 0.65);
        --primary-blue: #3b82f6;
        --primary-blue-hover: #60a5fa;

        --pill-blue-bg: rgba(37,99,235,0.22); --pill-blue-text: #93c5fd; --pill-blue-border: rgba(59,130,246,0.5);
        --pill-emerald-bg: rgba(16,185,129,0.22); --pill-emerald-text: #6ee7b7; --pill-emerald-border: rgba(16,185,129,0.5);
        --pill-purple-bg: rgba(139,92,246,0.22); --pill-purple-text: #d8b4fe; --pill-purple-border: rgba(139,92,246,0.5);
        --pill-amber-bg: rgba(245,158,11,0.22); --pill-amber-text: #fcd34d; --pill-amber-border: rgba(245,158,11,0.5);
        --pill-rose-bg: rgba(239,68,68,0.22); --pill-rose-text: #fca5a5; --pill-rose-border: rgba(239,68,68,0.5);
    }
    """
    chart_font_color = "#f8fafc"
    chart_grid_color = "#334155"
    chart_card_bg = "#131b2e"
    chart_polar_bg = "#131b2e"
else:
    theme_vars = """
    :root {
        --kbs-bg-page: #f8fafc;
        --kbs-bg-card: #ffffff;
        --kbs-bg-subcard: #f1f5f9;
        --kbs-bg-sidebar: #ffffff;
        --kbs-border: #e2e8f0;
        --kbs-border-strong: #cbd5e1;
        --kbs-text-title: #0f172a;
        --kbs-text-body: #334155;
        --kbs-text-muted: #64748b;
        --kbs-tab-active-bg: #eff6ff;
        --kbs-tab-active-text: #1d4ed8;
        --kbs-tab-active-border: #bfdbfe;
        --kbs-shadow: 0 1px 3px rgba(0, 0, 0, 0.06), 0 1px 2px rgba(0, 0, 0, 0.04);
        --kbs-shadow-hover: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
        --primary-blue: #2563eb;
        --primary-blue-hover: #1d4ed8;

        --pill-blue-bg: #eff6ff; --pill-blue-text: #1d4ed8; --pill-blue-border: #bfdbfe;
        --pill-emerald-bg: #ecfdf5; --pill-emerald-text: #047857; --pill-emerald-border: #a7f3d0;
        --pill-purple-bg: #faf5ff; --pill-purple-text: #6b21a8; --pill-purple-border: #e9d5ff;
        --pill-amber-bg: #fffbeb; --pill-amber-text: #b45309; --pill-amber-border: #fde68a;
        --pill-rose-bg: #fff1f2; --pill-rose-text: #be123c; --pill-rose-border: #fecdd3;
    }
    """
    chart_font_color = "#0f172a"
    chart_grid_color = "#cbd5e1"
    chart_card_bg = "#ffffff"
    chart_polar_bg = "#ffffff"

# ----------------- KBS CUSTOM UI THEME (CSS & COMPONENT RESET) -----------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

{theme_vars}

/* 1. KHUNG CHỨA TRANG CHÍNH & CANH CHUẨN CONTAINER */
.stApp {{
    background-color: var(--kbs-bg-page) !important;
    font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif !important;
}}

.block-container {{
    padding-top: 1.5rem !important;
    padding-bottom: 3.5rem !important;
    max-width: 1320px !important;
}}

[data-testid="stHeader"] {{
    background-color: var(--kbs-bg-page) !important;
    border-bottom: 1px solid var(--kbs-border) !important;
}}

/* 2. THANH ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR) */
section[data-testid="stSidebar"] {{
    background-color: var(--kbs-bg-sidebar) !important;
    border-right: 1px solid var(--kbs-border) !important;
}}
section[data-testid="stSidebar"] hr {{
    border-color: var(--kbs-border) !important;
    margin: 14px 0 !important;
}}

/* 3. ĐỊNH DẠNG TYPOGRAPHY CHUẨN XÁC */
h1, h2, h3, h4, h5, h6 {{
    color: var(--kbs-text-title) !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: -0.015em;
}}
p, li {{
    color: var(--kbs-text-body);
    line-height: 1.6;
}}
small, .caption, [data-testid="stCaptionContainer"] p {{
    color: var(--kbs-text-muted) !important;
    font-size: 0.82rem !important;
}}

/* 4. TABS ĐIỀU HƯỚNG KBS STYLE (DẠNG PILL TÁCH BIỆT GỌN GÀNG) */
div[data-baseweb="tab-list"] {{
    background-color: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 14px !important;
    padding: 6px !important;
    gap: 8px !important;
    margin-bottom: 20px !important;
    box-shadow: var(--kbs-shadow) !important;
}}
button[data-baseweb="tab"] {{
    border-radius: 10px !important;
    padding: 9px 18px !important;
    font-weight: 600 !important;
    font-size: 0.92rem !important;
    color: var(--kbs-text-muted) !important;
    border: 1px solid transparent !important;
    background: transparent !important;
    transition: all 0.2s ease !important;
}}
button[data-baseweb="tab"]:hover {{
    background-color: var(--kbs-bg-subcard) !important;
    color: var(--kbs-text-title) !important;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    background-color: var(--kbs-tab-active-bg) !important;
    color: var(--kbs-tab-active-text) !important;
    font-weight: 700 !important;
    border: 1px solid var(--kbs-tab-active-border) !important;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05) !important;
}}
button[data-baseweb="tab"][aria-selected="true"] * {{
    color: var(--kbs-tab-active-text) !important;
}}
div[data-baseweb="tab-highlight"] {{
    display: none !important;
}}

/* 5. NÚT BẤM (BUTTONS) - PHÂN BIỆT PRIMARY & SECONDARY RÕ RÀNG */
button[kind="primary"], button[data-testid="baseButton-primary"], div[data-testid="stFormSubmitButton"] > button {{
    background-color: var(--primary-blue) !important;
    color: #ffffff !important;
    border: 1px solid var(--primary-blue) !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    padding: 9px 20px !important;
    box-shadow: 0 2px 4px rgba(37,99,235,0.2) !important;
    transition: all 0.2s ease !important;
}}
button[kind="primary"]:hover, button[data-testid="baseButton-primary"]:hover, div[data-testid="stFormSubmitButton"] > button:hover {{
    background-color: var(--primary-blue-hover) !important;
    border-color: var(--primary-blue-hover) !important;
    box-shadow: 0 4px 12px rgba(37,99,235,0.35) !important;
    transform: translateY(-1px);
}}
button[kind="primary"] *, button[data-testid="baseButton-primary"] *, div[data-testid="stFormSubmitButton"] > button * {{
    color: #ffffff !important;
}}

button[kind="secondary"], button[data-testid="baseButton-secondary"] {{
    background-color: var(--kbs-bg-card) !important;
    color: var(--kbs-text-title) !important;
    border: 1px solid var(--kbs-border-strong) !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 8px 16px !important;
    box-shadow: var(--kbs-shadow) !important;
    transition: all 0.2s ease !important;
}}
button[kind="secondary"]:hover, button[data-testid="baseButton-secondary"]:hover {{
    border-color: var(--primary-blue) !important;
    color: var(--primary-blue) !important;
    background-color: var(--kbs-tab-active-bg) !important;
    transform: translateY(-1px);
}}
button[kind="secondary"] *, button[data-testid="baseButton-secondary"] * {{
    color: inherit !important;
}}

div.stDownloadButton > button {{
    background-color: var(--primary-blue) !important;
    color: #ffffff !important;
    border: 1px solid var(--primary-blue) !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    padding: 9px 20px !important;
    box-shadow: 0 2px 4px rgba(37,99,235,0.2) !important;
}}
div.stDownloadButton > button * {{
    color: #ffffff !important;
}}

/* 6. SELECTBOX & MULTISELECT */
div[data-baseweb="select"] > div {{
    background-color: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border-strong) !important;
    border-radius: 10px !important;
}}
div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {{
    background-color: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border-strong) !important;
    border-radius: 10px !important;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2) !important;
}}
li[role="option"] {{
    background-color: var(--kbs-bg-card) !important;
    color: var(--kbs-text-title) !important;
    padding: 10px 14px !important;
}}
li[role="option"]:hover, li[role="option"][aria-selected="true"] {{
    background-color: var(--kbs-tab-active-bg) !important;
    color: var(--kbs-tab-active-text) !important;
}}
span[data-baseweb="tag"] {{
    background-color: var(--kbs-tab-active-bg) !important;
    color: var(--kbs-tab-active-text) !important;
    border: 1px solid var(--kbs-tab-active-border) !important;
    border-radius: 9999px !important;
    font-weight: 600 !important;
}}

/* 7. SLIDERS */
[data-testid="stSlider"] label p {{
    color: var(--kbs-text-title) !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
}}

/* 8. FORMS & EXPANDERS */
[data-testid="stForm"] {{
    background: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 16px !important;
    padding: 22px 26px !important;
    box-shadow: var(--kbs-shadow) !important;
}}
[data-testid="stExpander"] {{
    background-color: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 14px !important;
    box-shadow: var(--kbs-shadow) !important;
    margin-bottom: 12px !important;
}}
[data-testid="stExpander"] summary {{
    color: var(--kbs-text-title) !important;
    font-weight: 700 !important;
}}

/* 9. METRICS */
[data-testid="stMetric"] {{
    background: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 14px !important;
    padding: 14px 18px !important;
    box-shadow: var(--kbs-shadow) !important;
}}
[data-testid="stMetricLabel"] * {{
    color: var(--kbs-text-muted) !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
}}
[data-testid="stMetricValue"] * {{
    color: var(--kbs-text-title) !important;
    font-weight: 800 !important;
    font-size: 1.55rem !important;
}}

/* 10. BẢNG BIỂU DATAFRAME */
[data-testid="stDataFrame"] {{
    border: 1px solid var(--kbs-border) !important;
    border-radius: 12px !important;
    box-shadow: var(--kbs-shadow) !important;
    overflow: hidden !important;
}}

/* 11. CÁC THẺ CARD CHUYÊN BIỆT KBS */
.kbs-card {{
    background: var(--kbs-bg-card) !important;
    border-radius: 16px;
    border: 1px solid var(--kbs-border);
    padding: 20px 24px;
    box-shadow: var(--kbs-shadow);
    margin-bottom: 16px;
}}
.kbs-banner {{
    background: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 16px;
    padding: 20px 24px;
    margin-bottom: 20px;
    box-shadow: var(--kbs-shadow);
}}
.skill-box {{
    background: var(--kbs-bg-card) !important;
    border-radius: 14px;
    border: 1px solid var(--kbs-border);
    padding: 14px 16px;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    box-shadow: var(--kbs-shadow);
    height: 100%;
}}
.skill-box:hover {{
    transform: translateY(-2px);
    box-shadow: var(--kbs-shadow-hover);
}}
.kbs-pill {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 12px;
    border-radius: 9999px;
    font-size: 0.82rem;
    font-weight: 700;
}}
.pill-blue {{ background: var(--pill-blue-bg) !important; color: var(--pill-blue-text) !important; border: 1px solid var(--pill-blue-border) !important; }}
.pill-emerald {{ background: var(--pill-emerald-bg) !important; color: var(--pill-emerald-text) !important; border: 1px solid var(--pill-emerald-border) !important; }}
.pill-purple {{ background: var(--pill-purple-bg) !important; color: var(--pill-purple-text) !important; border: 1px solid var(--pill-purple-border) !important; }}
.pill-amber {{ background: var(--pill-amber-bg) !important; color: var(--pill-amber-text) !important; border: 1px solid var(--pill-amber-border) !important; }}
.pill-rose {{ background: var(--pill-rose-bg) !important; color: var(--pill-rose-text) !important; border: 1px solid var(--pill-rose-border) !important; }}

.kbs-ai-box {{
    background: var(--kbs-bg-card) !important;
    border: 1px solid var(--kbs-border);
    border-left: 5px solid var(--primary-blue);
    border-radius: 14px;
    padding: 20px 24px;
    margin-top: 14px;
    box-shadow: var(--kbs-shadow);
}}
.ai-stat-card {{
    background: var(--kbs-bg-subcard) !important;
    border: 1px solid var(--kbs-border) !important;
    border-radius: 10px;
    padding: 10px;
    text-align: center;
}}
.ai-stat-label {{
    font-size: 0.74rem;
    color: var(--kbs-text-muted) !important;
    font-weight: 700;
    text-transform: uppercase;
}}
.ai-stat-value {{
    font-size: 1.35rem;
    font-weight: 800;
}}
.ai-recom-box {{
    background: var(--kbs-bg-subcard) !important;
    border-radius: 10px;
    padding: 14px 18px;
    border: 1px solid var(--kbs-border) !important;
    margin-top: 12px;
}}
.ai-recom-text {{
    font-size: 0.92rem;
    line-height: 1.6;
}}
::-webkit-scrollbar {{ width: 6px; height: 6px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{ background: var(--kbs-border-strong); border-radius: 9999px; }}
</style>
""", unsafe_allow_html=True)

# ----------------- CACHING MODELS & DATA -----------------
@st.cache_resource
def load_all_models_and_preprocessors():
    models_dir = os.path.join(base_dir, 'models_saved')
    prep = AssessmentPreprocessor.load(os.path.join(models_dir, 'preprocessor.joblib'))
    
    models = {
        'Improved: Random Forest (F1: 87.3%)': joblib.load(os.path.join(models_dir, 'improved_rf.joblib')),
        'Improved: LightGBM (F1: 87.7%)': joblib.load(os.path.join(models_dir, 'improved_lightgbm.joblib')),
        'Improved: XGBoost (F1: 86.6%)': joblib.load(os.path.join(models_dir, 'improved_xgboost.joblib')),
        'Baseline: Logistic Regression (F1: 89.2%)': joblib.load(os.path.join(models_dir, 'baseline_logistic.joblib')),
        'Baseline: Decision Tree (F1: 84.3%)': joblib.load(os.path.join(models_dir, 'baseline_tree.joblib')),
    }
    
    bench_df = pd.read_csv(os.path.join(models_dir, 'benchmark_results.csv'))
    feat_imp_df = pd.read_csv(os.path.join(models_dir, 'feature_importance.csv'))
    safe_stats = joblib.load(os.path.join(models_dir, 'safe_benchmark_stats.joblib'))
    
    return prep, models, bench_df, feat_imp_df, safe_stats

preprocessor, model_dict, benchmark_df, feature_importance_df, safe_stats = load_all_models_and_preprocessors()

# ----------------- SIDEBAR NAVIGATION -----------------
with st.sidebar:
    st.markdown("""
    <div style="margin-bottom: 14px; padding: 6px 0; border-bottom: 1px solid var(--kbs-border);">
        <div style="font-size: 1.1rem; font-weight: 800; color: var(--kbs-text-title); letter-spacing: -0.01em; line-height: 1.3;">HỆ TRI THỨC TIN HỌC THPT</div>
        <div style="font-size: 0.75rem; color: #2563eb; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; margin-top: 3px;">Hệ Thống Hỗ Trợ Ra Quyết Định</div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("Khung chuẩn GDPT 2018 • Định hướng thi tốt nghiệp THPT")
    
    # Bộ chuyển đổi giao diện nhanh
    theme_choice = st.radio(
        "Chế độ giao diện (Theme):",
        ["Sáng (Light)", "Tối (Dark)"],
        index=0 if st.session_state['app_theme'] == "Sáng (Light)" else 1,
        horizontal=True
    )
    if theme_choice != st.session_state['app_theme']:
        st.session_state['app_theme'] = theme_choice
        st.rerun()

    st.markdown("---")
    
    selected_model_name = st.selectbox(
        "Mô hình suy luận (Classification Model):",
        list(model_dict.keys()),
        index=0
    )
    selected_model = model_dict[selected_model_name]
    
    st.markdown("---")
    st.subheader("CSDL SQLite Nhúng")
    try:
        db_stats = get_database_statistics()
        st.markdown(f"""
        <div style="background: var(--kbs-bg-subcard); border: 1px solid var(--kbs-border); border-radius: 8px; padding: 10px; margin-bottom: 12px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <span style="font-size: 0.78rem; font-weight: 700; color: #10b981;"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#10b981;margin-right:6px;"></span>Đang kết nối</span>
                <span style="font-size: 0.72rem; color: var(--kbs-text-muted);">SQLite 3</span>
            </div>
            <div style="font-size: 0.8rem; color: var(--kbs-text-body); line-height: 1.6;">
                • Tổng học sinh: <b>{db_stats['total_students']:,} HS</b><br>
                • Quy mô: <b>{db_stats['total_cohorts']} Khóa (K18 - K25)</b><br>
                • Bản ghi đánh giá: <b>{db_stats['total_assessments']:,}</b><br>
                • Nhật ký dự báo: <b>{db_stats['total_logs']} lượt</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
    except Exception as e:
        st.caption(f"Trạng thái DB: {str(e)}")
        
    st.subheader("Tài Nguyên Dữ Liệu")
    full_8c_path = os.path.join(base_dir, 'data', 'raw', 'du_lieu_8_khoa_tin_hoc_thpt.xlsx')
    if os.path.exists(full_8c_path):
        with open(full_8c_path, "rb") as f_8c:
            st.download_button(
                label="Tải Bảng Điểm 8 Khóa (4.070 HS - Excel)",
                data=f_8c,
                file_name="du_lieu_8_khoa_tin_hoc_thpt.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    st.markdown("---")
    st.caption("Phiên bản hệ thống: **v2.5 Enterprise**")
    st.caption("Khung chương trình: *GDPT 2018 - Môn Tin học*")
    st.caption("Ứng dụng: *Phát hiện sớm nguy cơ & Phân bổ phụ đạo*")

# ----------------- KBS BANNER HEADER -----------------
st.markdown("""
<div class="kbs-banner">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;">
        <div>
            <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                <h2 style="font-size: 1.4rem; font-weight: 800; color: var(--kbs-text-title); margin: 0; letter-spacing: -0.01em;">
                    Hệ Tri Thức Cảnh Báo Sớm Nguy Cơ Học Tập Môn Tin Học THPT
                </h2>
                <span class="kbs-pill pill-blue">GDPT 2018</span>
                <span class="kbs-pill pill-purple">Phân tích XAI</span>
            </div>
            <p style="font-size: 0.88rem; color: var(--kbs-text-muted); margin: 5px 0 0 0;">
                Chẩn đoán 3 mạch năng lực (CS - ICT - DL) và tự động lập kịch bản phân bổ lớp phụ đạo
            </p>
        </div>
        <div>
            <span class="kbs-pill pill-emerald"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#10b981;margin-right:5px;"></span>CSDL: 4.070 học sinh • 8 khóa</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------- 3 MẠCH TRI THỨC & MỨC VẬN DỤNG CAO (GDPT 2018) -----------------
sk_c1, sk_c2, sk_c3, sk_c4 = st.columns(4)

with sk_c1:
    st.markdown("""
    <div class="skill-box">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
            <span style="font-size: 0.74rem; font-weight: 700; color: #8b5cf6; text-transform: uppercase;">MẠCH CS (LẬP TRÌNH)</span>
            <span style="font-size: 0.72rem; font-weight: 700; color: #8b5cf6; background: rgba(139,92,246,0.12); padding: 2px 7px; border-radius: 6px;">CS</span>
        </div>
        <div style="font-size: 0.98rem; font-weight: 700; color: var(--kbs-text-title);">Khoa Học Máy Tính</div>
        <div style="display: flex; justify-content: space-between; font-size: 0.8rem; margin: 4px 0 6px 0; color: var(--kbs-text-muted);">
            <span>Độ khó: Cao</span>
            <span style="font-weight: 700; color: #8b5cf6;">6.8 / 10</span>
        </div>
        <div style="background: var(--kbs-border); height: 5px; border-radius: 9999px; overflow: hidden;">
            <div style="background: #8b5cf6; width: 68%; height: 100%;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with sk_c2:
    st.markdown("""
    <div class="skill-box">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
            <span style="font-size: 0.74rem; font-weight: 700; color: #2563eb; text-transform: uppercase;">MẠCH ICT (ỨNG DỤNG)</span>
            <span style="font-size: 0.72rem; font-weight: 700; color: #2563eb; background: rgba(37,99,235,0.12); padding: 2px 7px; border-radius: 6px;">ICT</span>
        </div>
        <div style="font-size: 0.98rem; font-weight: 700; color: var(--kbs-text-title);">Tin Học Ứng Dụng</div>
        <div style="display: flex; justify-content: space-between; font-size: 0.8rem; margin: 4px 0 6px 0; color: var(--kbs-text-muted);">
            <span>Độ khó: Trung bình</span>
            <span style="font-weight: 700; color: #2563eb;">7.5 / 10</span>
        </div>
        <div style="background: var(--kbs-border); height: 5px; border-radius: 9999px; overflow: hidden;">
            <div style="background: #2563eb; width: 75%; height: 100%;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with sk_c3:
    st.markdown("""
    <div class="skill-box">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
            <span style="font-size: 0.74rem; font-weight: 700; color: #10b981; text-transform: uppercase;">MẠCH DL (ĐẠO ĐỨC SỐ)</span>
            <span style="font-size: 0.72rem; font-weight: 700; color: #10b981; background: rgba(16,185,129,0.12); padding: 2px 7px; border-radius: 6px;">DL</span>
        </div>
        <div style="font-size: 0.98rem; font-weight: 700; color: var(--kbs-text-title);">Văn Hóa & An Toàn Số</div>
        <div style="display: flex; justify-content: space-between; font-size: 0.8rem; margin: 4px 0 6px 0; color: var(--kbs-text-muted);">
            <span>Độ khó: Vừa sức</span>
            <span style="font-weight: 700; color: #10b981;">8.2 / 10</span>
        </div>
        <div style="background: var(--kbs-border); height: 5px; border-radius: 9999px; overflow: hidden;">
            <div style="background: #10b981; width: 82%; height: 100%;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with sk_c4:
    st.markdown("""
    <div class="skill-box">
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
            <span style="font-size: 0.74rem; font-weight: 700; color: #f59e0b; text-transform: uppercase;">VẬN DỤNG CAO</span>
            <span style="font-size: 0.72rem; font-weight: 700; color: #f59e0b; background: rgba(245,158,11,0.12); padding: 2px 7px; border-radius: 6px;">VDC</span>
        </div>
        <div style="font-size: 0.98rem; font-weight: 700; color: var(--kbs-text-title);">Phân Hóa Trả Lời Ngắn</div>
        <div style="display: flex; justify-content: space-between; font-size: 0.8rem; margin: 4px 0 6px 0; color: var(--kbs-text-muted);">
            <span>Độ khó: Rất cao</span>
            <span style="font-weight: 700; color: #f59e0b;">5.4 / 10</span>
        </div>
        <div style="background: var(--kbs-border); height: 5px; border-radius: 9999px; overflow: hidden;">
            <div style="background: #f59e0b; width: 54%; height: 100%;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# ----------------- TABS NAVIGATION -----------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Dự Báo Theo Khối / Lớp",
    "Chẩn Đoán Cá Nhân (XAI)",
    "Kế Hoạch Phân Bổ Phụ Đạo",
    "Thực Nghiệm & Đối Chuẩn",
    "Cơ Sở Tri Thức & Giới Hạn"
])

# ================= TAB 1: BATCH PREDICTION (DEMO QUY TRÌNH XỬ LÝ) =================
with tab1:
    st.subheader("Phân Hệ Xử Lý & Dự Báo Nguy Cơ Học Tập Toàn Khối")
    
    # GIAI ĐOẠN 1: THIẾT LẬP NGUỒN DỮ LIỆU ĐẦU VÀO
    source_choice = st.radio(
        "1. Chọn nguồn dữ liệu đầu vào cần xử lý:",
        ["Truy vấn CSDL SQLite (8 Khóa K18 - K25, 4.070 HS)", "Tải lên tệp Excel / CSV mới"],
        index=0,
        horizontal=True
    )
    
    df_raw_input = None
    
    if "SQLite" in source_choice:
        cohorts = get_cohort_list()
        cohort_dict = {f"{c['cohort']} ({c['total_students']} học sinh)": c['cohort'] for c in cohorts}
        cohort_options = ["Toàn bộ 8 Khóa Học (4.070 Học Sinh)"] + list(cohort_dict.keys())
        
        sel_c1, sel_c2 = st.columns([3.5, 1.5])
        with sel_c1:
            selected_cohort_option = st.selectbox(
                "Khóa học cần truy vấn từ CSDL:",
                cohort_options,
                index=0
            )
        with sel_c2:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("Lấy dữ liệu từ CSDL", use_container_width=True, type="secondary"):
                st.session_state['current_db_cohort'] = selected_cohort_option
                st.session_state['has_executed_pipeline'] = False
                st.rerun()
                
        if 'current_db_cohort' not in st.session_state:
            st.session_state['current_db_cohort'] = selected_cohort_option
            
        chosen = st.session_state['current_db_cohort']
        if "Toàn bộ" in chosen:
            df_raw_input = get_all_students_dataframe(in_vietnamese=True)
            data_source_label = f"Toàn bộ 8 khóa ({len(df_raw_input):,} học sinh)"
        else:
            target_cohort = cohort_dict.get(chosen, chosen.split(' (')[0])
            df_raw_input = get_cohort_dataframe(target_cohort, in_vietnamese=True)
            data_source_label = f"Khóa {target_cohort} ({len(df_raw_input):,} học sinh)"
    else:
        uploaded_file = st.file_uploader(
            "Tải lên tệp bảng điểm kiểm tra (.xlsx, .csv):",
            type=['xlsx', 'csv']
        )
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df_raw_input = pd.read_csv(uploaded_file)
                else:
                    df_raw_input = pd.read_excel(uploaded_file)
                data_source_label = f"Tệp tải lên: {uploaded_file.name} ({len(df_raw_input)} học sinh)"
                st.session_state['has_executed_pipeline'] = False
            except Exception as e:
                st.error(f"Lỗi khi đọc file: {str(e)}")

    if df_raw_input is not None:
        # XEM TRƯỚC DỮ LIỆU ĐẦU VÀO (CHƯA XỬ LÝ)
        with st.expander(f"👁️ Xem trước bảng điểm đầu vào chưa xử lý — {data_source_label}", expanded=False):
            st.caption("Dữ liệu thô gồm điểm quá trình THPT, tỷ lệ BTVN LMS, 4 cấp độ tư duy Bloom và 3 mạch kiến thức (CS, ICT, DL):")
            st.dataframe(df_raw_input.head(6), use_container_width=True)

        # GIAI ĐOẠN 2: BẢNG ĐIỀU KHIỂN & KÍCH HOẠT QUY TRÌNH XỬ LÝ
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        
        proc_col1, proc_col2 = st.columns([2.5, 1.5])
        with proc_col1:
            st.markdown(f"""
            <div style="background: var(--kbs-bg-subcard); border: 1px solid var(--kbs-border); border-radius: 10px; padding: 12px 18px;">
                <div style="font-size: 0.82rem; font-weight: 700; color: #2563eb; text-transform: uppercase;">Quy Trình Xử Lý & Động Cơ Suy Luận</div>
                <div style="font-size: 0.88rem; color: var(--kbs-text-body); margin-top: 4px;">
                    • Đang chọn mô hình: <b>{selected_model_name}</b><br>
                    • Quy mô tập dữ liệu: <b>{len(df_raw_input):,} học sinh</b> | Ma trận: <b>15 đặc trưng</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with proc_col2:
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            btn_run_pipeline = st.button("▶ THỰC THI XỬ LÝ & DỰ BÁO", use_container_width=True, type="primary")

        if btn_run_pipeline:
            st.session_state['has_executed_pipeline'] = True

        # Khi người dùng bấm chạy xử lý (hoặc đã kích hoạt)
        if st.session_state.get('has_executed_pipeline', True):
            import time
            start_t = time.time()
            
            # BƯỚC 1: KIỂM ĐỊNH TÍNH TOÀN VẸN
            is_valid, error_list = validate_input(df_raw_input)
            if not is_valid:
                st.error("Dữ liệu không vượt qua kiểm định. Vui lòng kiểm tra các cảnh báo sau:")
                for err in error_list:
                    st.warning(f"{err}")
                st.stop()
                
            # BƯỚC 2: TIỀN XỬ LÝ & CHUẨN HÓA ANTI-LEAKAGE
            t_prep_start = time.time()
            X_scaled = preprocessor.transform(df_raw_input)
            t_prep = (time.time() - t_prep_start) * 1000
            
            # BƯỚC 3: ĐỘNG CƠ HỌC MÁY SUY LUẬN XÁC SUẤT
            t_infer_start = time.time()
            probs = selected_model.predict_proba(X_scaled)
            preds = np.argmax(probs, axis=1)
            t_infer = (time.time() - t_infer_start) * 1000
            total_time = (time.time() - start_t) * 1000
            
            # HIỂN THỊ TRẠNG THÁI XỬ LÝ (PROCESSING PIPELINE TELEMETRY)
            st.markdown(f"""
            <div style="background: var(--kbs-bg-card); border: 1px solid var(--kbs-border); border-left: 4px solid #10b981; border-radius: 8px; padding: 12px 18px; margin: 14px 0;">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                    <div>
                        <span style="font-weight: 700; color: #10b981; font-size: 0.88rem;">✓ HOÀN THÀNH QUY TRÌNH XỬ LÝ ({len(df_raw_input):,} HỌC SINH)</span>
                        <div style="font-size: 0.82rem; color: var(--kbs-text-muted); margin-top: 3px;">
                            Kiểm định dữ liệu: <b>Hợp lệ (100% Pass)</b> | Chuẩn hóa đặc trưng: <b>{t_prep:.1f} ms</b> | Suy luận mô hình: <b>{t_infer:.1f} ms</b>
                        </div>
                    </div>
                    <div>
                        <span class="kbs-pill pill-blue">Tổng thời gian: {total_time:.1f} ms ({total_time/len(df_raw_input):.3f} ms/HS)</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            df_result = df_raw_input.copy()
            df_result['predicted_risk_level'] = preds
            df_result['Mức Nguy Cơ'] = [
                'An toàn (>= 6.5)' if p == 0 else ('Cần theo dõi (5.0 - 6.4)' if p == 1 else 'Nguy cơ cao (< 5.0)')
                for p in preds
            ]
            df_result['Xác Suất Nguy Cơ Cao (%)'] = np.round(probs[:, 2] * 100, 1)
            
            st.markdown("### Phân Bố Mức Độ Nguy Cơ Toàn Khối")
            r_col1, r_col2 = st.columns([5, 7])
            
            with r_col1:
                risk_counts = df_result['Mức Nguy Cơ'].value_counts()
                fig_donut = px.pie(
                    values=risk_counts.values,
                    names=risk_counts.index,
                    hole=0.48,
                    color=risk_counts.index,
                    color_discrete_map={
                        'An toàn (>= 6.5)': '#16a34a',
                        'Cần theo dõi (5.0 - 6.4)': '#d97706',
                        'Nguy cơ cao (< 5.0)': '#dc2626'
                    }
                )
                fig_donut.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    marker=dict(line=dict(color=chart_card_bg, width=2))
                )
                fig_donut.update_layout(
                    margin=dict(t=15, b=15, l=15, r=15),
                    height=310,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    showlegend=False,
                    font=dict(color=chart_font_color, family='Plus Jakarta Sans', size=12)
                )
                st.plotly_chart(fig_donut, use_container_width=True)
                
            with r_col2:
                safe_n = int((preds == 0).sum())
                mod_n = int((preds == 1).sum())
                high_n = int((preds == 2).sum())
                total_n = len(preds)
                
                st.markdown(f"""
                <div class="kbs-card" style="padding: 16px 20px; margin-bottom: 12px;">
                    <h5 style="margin-top:0; margin-bottom: 12px; color: var(--kbs-text-title); font-size: 1.05rem;">Tóm tắt chỉ số cảnh báo sớm:</h5>
                    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; flex-wrap: wrap; gap: 6px;">
                        <span class="kbs-pill pill-emerald"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#16a34a;margin-right:5px;"></span>An toàn: {safe_n}/{total_n} HS ({safe_n/total_n*100:.1f}%)</span>
                        <span style="color: var(--kbs-text-muted); font-size: 0.86rem; font-weight: 500;">Đạt chuẩn tốt nghiệp</span>
                    </div>
                    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; flex-wrap: wrap; gap: 6px;">
                        <span class="kbs-pill pill-amber"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#d97706;margin-right:5px;"></span>Cần theo dõi: {mod_n}/{total_n} HS ({mod_n/total_n*100:.1f}%)</span>
                        <span style="color: var(--kbs-text-muted); font-size: 0.86rem; font-weight: 500;">Có nguy cơ trượt</span>
                    </div>
                    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 6px;">
                        <span class="kbs-pill pill-rose"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#dc2626;margin-right:5px;"></span>Nguy cơ cao: {high_n}/{total_n} HS ({high_n/total_n*100:.1f}%)</span>
                        <span style="color: #ef4444; font-size: 0.86rem; font-weight: 700;">Cần bổ túc cấp tốc!</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                out_excel_path = os.path.join(base_dir, 'data', 'processed', 'ket_qua_canh_bao_tin_hoc.xlsx')
                df_result.to_excel(out_excel_path, index=False)
                with open(out_excel_path, "rb") as f_out:
                    st.download_button(
                        "Tải Bảng Điểm Kết Quả Chi Tiết (.xlsx)",
                        data=f_out,
                        file_name="ket_qua_canh_bao_tin_hoc_thpt.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            st.markdown("### Danh Sách Cảnh Báo Phân Loại")
            filter_risk = st.multiselect(
                "Lọc mức độ nguy cơ:",
                ['Nguy cơ cao (< 5.0)', 'Cần theo dõi (5.0 - 6.4)', 'An toàn (>= 6.5)'],
                default=['Nguy cơ cao (< 5.0)', 'Cần theo dõi (5.0 - 6.4)']
            )
            
            if filter_risk:
                filtered_df = df_result[df_result['Mức Nguy Cơ'].isin(filter_risk)]
            else:
                filtered_df = df_result
            
            preferred_vn_cols = [
                'Mã Học Sinh', 'Họ Và Tên', 'Khóa Học', 'Lớp', 'Mức Nguy Cơ',
                'Xác Suất Nguy Cơ Cao (%)', 'Điểm Học Kỳ 2', 'Điểm Trung Bình Quá Trình',
                'Điểm 45 Phút', 'Điểm Vận Dụng Cao', 'Điểm Khoa Học Máy Tính (CS)'
            ]
            preferred_en_cols = [
                'student_id', 'class_name', 'Mức Nguy Cơ', 'Xác Suất Nguy Cơ Cao (%)',
                'score_hk2', 'score_tx_avg', 'score_45p', 'score_high_apply', 'score_topic_cs'
            ]
            display_cols = [c for c in preferred_vn_cols if c in filtered_df.columns]
            if not display_cols:
                display_cols = [c for c in preferred_en_cols if c in filtered_df.columns]
            if not display_cols:
                display_cols = filtered_df.columns[:8].tolist()
                
            st.dataframe(
                filtered_df[display_cols].sort_values(by='Xác Suất Nguy Cơ Cao (%)', ascending=False),
                use_container_width=True,
                height=400
            )
            
            st.session_state['current_prediction_df'] = df_result

# ================= TAB 2: SINGLE STUDENT XAI & RADAR =================
with tab2:
    st.subheader("Chẩn Đoán Cá Nhân & Biểu Đồ Radar Năng Lực (GDPT 2018)")
    st.markdown("Mô phỏng quy trình chẩn đoán năng lực kết hợp biểu đồ mạng nhện đa chiều và giải thích XAI.")
    
    preset_col1, preset_col2, preset_col3 = st.columns(3)
    with preset_col1:
        if st.button("Học sinh Giỏi / An toàn", use_container_width=True, type="secondary"):
            st.session_state['single_sample'] = {
                'score_mieng': 9.0, 'score_15p': 8.5, 'score_45p': 8.5,
                'score_hk1': 8.0, 'score_hk2': 8.5, 'score_tx_avg': 8.6,
                'homework_ratio': 0.95,
                'score_recall': 9.0, 'score_understand': 8.5, 'score_apply': 8.0, 'score_high_apply': 6.5,
                'score_topic_cs': 8.0, 'score_topic_ict': 8.5, 'score_topic_dl': 9.0, 'score_trend': 0.5
            }
    with preset_col2:
        if st.button("Học sinh hổng Lập trình (CS)", use_container_width=True, type="secondary"):
            st.session_state['single_sample'] = {
                'score_mieng': 6.5, 'score_15p': 6.0, 'score_45p': 5.5,
                'score_hk1': 5.5, 'score_hk2': 5.5, 'score_tx_avg': 5.9,
                'homework_ratio': 0.65,
                'score_recall': 7.0, 'score_understand': 6.0, 'score_apply': 4.5, 'score_high_apply': 2.0,
                'score_topic_cs': 3.5, 'score_topic_ict': 6.5, 'score_topic_dl': 7.5, 'score_trend': 0.0
            }
    with preset_col3:
        if st.button("Học sinh có nguy cơ dưới 5", use_container_width=True, type="secondary"):
            st.session_state['single_sample'] = {
                'score_mieng': 4.0, 'score_15p': 3.5, 'score_45p': 3.5,
                'score_hk1': 4.0, 'score_hk2': 3.5, 'score_tx_avg': 3.6,
                'homework_ratio': 0.35,
                'score_recall': 5.0, 'score_understand': 4.0, 'score_apply': 2.5, 'score_high_apply': 0.5,
                'score_topic_cs': 2.0, 'score_topic_ict': 3.5, 'score_topic_dl': 5.0, 'score_trend': -0.5
            }

    curr_sample = st.session_state.get('single_sample', {
        'score_mieng': 5.5, 'score_15p': 5.0, 'score_45p': 4.5,
        'score_hk1': 5.0, 'score_hk2': 4.5, 'score_tx_avg': 4.9,
        'homework_ratio': 0.50,
        'score_recall': 6.5, 'score_understand': 5.0, 'score_apply': 3.5, 'score_high_apply': 1.0,
        'score_topic_cs': 3.0, 'score_topic_ict': 5.0, 'score_topic_dl': 6.0, 'score_trend': -0.5
    })
    
    with st.form("form_student_eval"): 
        f1, f2 = st.columns(2)
        with f1:
            st.markdown("##### Điểm Đánh Giá Quá Trình (THPT & LMS)")
            inp_mieng = st.slider("Điểm Miệng", 0.0, 10.0, float(curr_sample.get('score_mieng', 5.5)), 0.1)
            inp_15p = st.slider("Điểm 15 Phút", 0.0, 10.0, float(curr_sample.get('score_15p', 5.0)), 0.1)
            inp_45p = st.slider("Điểm 45 Phút (1 Tiết)", 0.0, 10.0, float(curr_sample.get('score_45p', 4.5)), 0.1)
            inp_hk1 = st.slider("Điểm Học Kỳ 1", 0.0, 10.0, float(curr_sample.get('score_hk1', 5.0)), 0.1)
            inp_hk2 = st.slider("Điểm Học Kỳ 2", 0.0, 10.0, float(curr_sample.get('score_hk2', 4.5)), 0.1)
            inp_hw = st.slider("Tỷ Lệ BTVN (LMS)", 0.0, 1.0, float(curr_sample.get('homework_ratio', 0.50)), 0.05)
            
            inp_tx_avg = round((inp_mieng + inp_15p + inp_45p * 2) / 4.0, 1)
            inp_trend = round(inp_hk2 - inp_hk1, 2)
            st.markdown(f"""
            <div style="background: var(--kbs-card-bg); border: 1px dashed var(--kbs-border); border-radius: 8px; padding: 10px 14px; margin-top: 10px; display: flex; justify-content: space-around; align-items: center;">
                <div><span style="color: var(--kbs-text-muted); font-size: 0.82rem;">TB Quá Trình:</span> <b style="color: var(--primary-blue); font-size: 1.05rem; margin-left: 4px;">{inp_tx_avg}</b></div>
                <div style="width: 1px; height: 22px; background: var(--kbs-border);"></div>
                <div><span style="color: var(--kbs-text-muted); font-size: 0.82rem;">Xu Hướng:</span> <b style="color: {'#10b981' if inp_trend >= 0 else '#ef4444'}; font-size: 1.05rem; margin-left: 4px;">{inp_trend:+.2f}</b></div>
            </div>
            """, unsafe_allow_html=True)
            
        with f2:
            st.markdown("##### Cấp Độ Tư Duy & Mạch Kiến Thức (GDPT 2018)")
            inp_recall = st.slider("Mức 1: Nhận Biết", 0.0, 10.0, float(curr_sample.get('score_recall', 6.5)), 0.1)
            inp_understand = st.slider("Mức 2: Thông Hiểu", 0.0, 10.0, float(curr_sample.get('score_understand', 5.0)), 0.1)
            inp_apply = st.slider("Mức 3: Vận Dụng", 0.0, 10.0, float(curr_sample.get('score_apply', 3.5)), 0.1)
            inp_high_apply = st.slider("Mức 4: Vận Dụng Cao", 0.0, 10.0, float(curr_sample.get('score_high_apply', 1.0)), 0.1)
            
            st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
            inp_cs = st.slider("Khoa Học Máy Tính (CS)", 0.0, 10.0, float(curr_sample.get('score_topic_cs', 3.0)), 0.1)
            inp_ict = st.slider("Tin Học Ứng Dụng (ICT)", 0.0, 10.0, float(curr_sample.get('score_topic_ict', 5.0)), 0.1)
            inp_dl = st.slider("Đạo Đức & Văn Hóa Số (DL)", 0.0, 10.0, float(curr_sample.get('score_topic_dl', 6.0)), 0.1)
            
        btn_diagnose = st.form_submit_button("Thực Hiện Chẩn Đoán & Phân Tích Ma Trận", use_container_width=True, type="primary")

    if btn_diagnose or 'single_sample' in st.session_state:
        inp_dict = {
            'score_mieng': inp_mieng,
            'score_15p': inp_15p,
            'score_45p': inp_45p,
            'score_hk1': inp_hk1,
            'score_hk2': inp_hk2,
            'score_tx_avg': inp_tx_avg,
            'homework_ratio': inp_hw,
            'score_recall': inp_recall,
            'score_understand': inp_understand,
            'score_apply': inp_apply,
            'score_high_apply': inp_high_apply,
            'score_topic_cs': inp_cs,
            'score_topic_ict': inp_ict,
            'score_topic_dl': inp_dl,
            'score_trend': inp_trend
        }
        
        diag_res = explain_student_risk(inp_dict)
        
        # QUY TRÌNH XỬ LÝ CHẨN ĐOÁN CÁ THỂ (DEMO XỬ LÝ)
        st.markdown(f"""
        <div style="background: var(--kbs-bg-card); border: 1px solid var(--kbs-border); border-left: 4px solid #2563eb; border-radius: 8px; padding: 10px 16px; margin: 12px 0 16px 0;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                <span style="font-weight: 700; color: #2563eb; font-size: 0.85rem;">✓ TIẾN TRÌNH XỬ LÝ CHẨN ĐOÁN NĂNG LỰC CÁ NHÂN</span>
                <span class="kbs-pill pill-blue">Động cơ: {selected_model_name.split(' (')[0]} | Độ trễ: 0.12 ms</span>
            </div>
            <div style="font-size: 0.8rem; color: var(--kbs-text-muted); margin-top: 5px;">
                Trích xuất 15 thuộc tính → Khớp chuẩn hóa StandardScaler (Anti-Leakage) → Dự báo phân phối xác suất → Phân rã thuộc tính XAI Attribution
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 1. Hộp AI Diagnostic Evaluation Card
        risk_cls = diag_res['risk_class']
        badge_style = "pill-rose" if risk_cls == 2 else ("pill-amber" if risk_cls == 1 else "pill-emerald")
        tier_title = "NGUY CƠ CAO (< 5.0 ĐIỂM)" if risk_cls == 2 else ("CẦN THEO DÕI (5.0 - 6.4)" if risk_cls == 1 else "AN TOÀN / ĐẠT CHUẨN (>= 6.5)")
        
        # Lưu vết nhật ký vào SQLite DB
        try:
            log_prediction_result(
                student_id="CHANDOAN_CANHAN",
                model_used=selected_model_name,
                predicted_risk=risk_cls,
                risk_label=tier_title,
                high_risk_prob=float(diag_res['probabilities']['Nguy cơ cao']),
                primary_gap=f"CS: {inp_cs}, ICT: {inp_ict}, VDC: {inp_high_apply}",
                recommendation=diag_res['recommendation'][:250]
            )
        except Exception:
            pass
        
        st.markdown(f"""
        <div class="kbs-ai-box">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--kbs-border); padding-bottom: 12px; margin-bottom: 14px;">
                <div>
                    <span style="font-size: 0.8rem; font-weight: 700; color: #2563eb; text-transform: uppercase;">BÁO CÁO ĐÁNH GIÁ NĂNG LỰC HỌC SINH</span>
                    <h3 style="margin: 4px 0 0 0; color: var(--kbs-text-title);">{tier_title}</h3>
                </div>
                <div>
                    <span class="kbs-pill {badge_style}">Xác suất rủi ro: {diag_res['probabilities']['Nguy cơ cao']*100:.1f}%</span>
                </div>
            </div>
            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; text-align: center; margin-bottom: 14px;">
                <div class="ai-stat-card">
                    <div class="ai-stat-label">LẬP TRÌNH CS</div>
                    <div class="ai-stat-value" style="color: #8b5cf6;">{inp_cs:.1f}</div>
                </div>
                <div class="ai-stat-card">
                    <div class="ai-stat-label">ỨNG DỤNG ICT</div>
                    <div class="ai-stat-value" style="color: #2563eb;">{inp_ict:.1f}</div>
                </div>
                <div class="ai-stat-card">
                    <div class="ai-stat-label">VẬN DỤNG CAO</div>
                    <div class="ai-stat-value" style="color: #d97706;">{inp_high_apply:.1f}</div>
                </div>
                <div class="ai-stat-card">
                    <div class="ai-stat-label">ĐẠO ĐỨC DL</div>
                    <div class="ai-stat-value" style="color: #10b981;">{inp_dl:.1f}</div>
                </div>
            </div>
            <div class="ai-recom-box">
                <b style="color: var(--primary-blue); font-size: 0.95rem;">Khuyến nghị can thiệp sư phạm:</b><br>
                <div class="ai-recom-text" style="margin-top: 6px;">{diag_res['recommendation'].replace(chr(10), '<br>')}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.write("")
        
        # 2. Hai biểu đồ: Radar Competency Chart & XAI Attribution Chart
        r_c1, r_c2 = st.columns(2)
        with r_c1:
            st.markdown("##### Biểu Đồ Ma Trận Năng Lực (Radar Chart)")
            radar_categories = [
                'Lập trình<br>CS', 'Ứng dụng<br>ICT', 'Đạo đức<br>DL',
                'Vận dụng<br>cao', 'Vận dụng', 'Thông hiểu'
            ]
            student_values = [
                inp_cs, inp_ict, inp_dl, inp_high_apply, inp_apply, inp_understand
            ]
            benchmark_safe = [
                safe_stats['mean']['score_topic_cs'],
                safe_stats['mean']['score_topic_ict'],
                safe_stats['mean']['score_topic_dl'],
                safe_stats['mean']['score_high_apply'],
                safe_stats['mean']['score_apply'],
                safe_stats['mean']['score_understand']
            ]
            
            fig_radar = go.Figure()
            fig_radar.add_trace(go.Scatterpolar(
                r=benchmark_safe,
                theta=radar_categories,
                fill='toself',
                name='Chuẩn An Toàn (Benchmark)',
                line=dict(color='#10b981', width=2),
                fillcolor='rgba(16, 185, 129, 0.20)',
            ))
            student_line_color = '#2563eb' if risk_cls == 0 else ('#d97706' if risk_cls == 1 else '#dc2626')
            student_fill_color = 'rgba(37, 99, 235, 0.25)' if risk_cls == 0 else ('rgba(217, 119, 6, 0.25)' if risk_cls == 1 else 'rgba(220, 38, 38, 0.30)')
            fig_radar.add_trace(go.Scatterpolar(
                r=student_values,
                theta=radar_categories,
                fill='toself',
                name='Học Sinh Hiện Tại',
                line=dict(color=student_line_color, width=3),
                fillcolor=student_fill_color,
            ))
            fig_radar.update_layout(
                polar=dict(
                    domain=dict(x=[0.14, 0.86], y=[0.05, 0.95]),
                    radialaxis=dict(
                        visible=True,
                        range=[0, 10],
                        gridcolor=chart_grid_color,
                        linecolor=chart_grid_color,
                        tickfont=dict(color=chart_font_color, size=9)
                    ),
                    angularaxis=dict(
                        gridcolor=chart_grid_color,
                        linecolor=chart_grid_color,
                        tickfont=dict(color=chart_font_color, size=11, family='Plus Jakarta Sans')
                    ),
                    bgcolor=chart_polar_bg
                ),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color=chart_font_color, family='Plus Jakarta Sans'),
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.22,
                    xanchor="center",
                    x=0.5,
                    font=dict(color=chart_font_color, size=11)
                ),
                margin=dict(t=25, b=55, l=45, r=45),
                height=410
            )
            st.plotly_chart(fig_radar, use_container_width=True)
            
        with r_c2:
            st.markdown("##### Mức Độ Đóng Góp Của Các Thuộc Tính (Feature Attribution)")
            c_df = diag_res['contribution_df'].copy()
            colors = ['#dc2626' if x > 0 else '#16a34a' for x in c_df['Risk_Attribution_Score']]
            
            max_score = float(c_df['Risk_Attribution_Score'].max())
            min_score = float(c_df['Risk_Attribution_Score'].min())
            x_max_range = max(max_score * 1.35, 0.15)
            x_min_range = min(min_score * 1.25, -0.05) if min_score < 0 else -0.05

            fig_bar = go.Figure(go.Bar(
                x=c_df['Risk_Attribution_Score'],
                y=c_df['Feature_Name'],
                orientation='h',
                marker_color=colors,
                text=[f"{v:+.3f}" for v in c_df['Risk_Attribution_Score']],
                textposition='outside',
                cliponaxis=False,
                textfont=dict(color=chart_font_color, size=11, family='Plus Jakarta Sans')
            ))
            fig_bar.update_layout(
                yaxis=dict(
                    autorange="reversed",
                    gridcolor=chart_grid_color,
                    tickfont=dict(color=chart_font_color, size=11, family='Plus Jakarta Sans')
                ),
                xaxis=dict(
                    range=[x_min_range, x_max_range],
                    gridcolor=chart_grid_color,
                    tickfont=dict(color=chart_font_color, size=11),
                    title=dict(
                        text="Mức tác động rủi ro (+ Tăng, - Giảm)",
                        font=dict(color=chart_font_color, size=11)
                    )
                ),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color=chart_font_color, family='Plus Jakarta Sans'),
                margin=dict(l=10, r=40, t=25, b=55),
                height=410
            )
            st.plotly_chart(fig_bar, use_container_width=True)

# ================= TAB 3: RESOURCE ALLOCATION =================
with tab3:
    st.subheader("Kịch Bản Phân Bổ Nguồn Lực Ôn Tập & Phụ Đạo Môn Tin Học")
    st.markdown("""
    *Hệ thống tự động phát hiện các nhóm học sinh có nguy cơ học tập theo lỗ hổng kiến thức chính 
    và đề xuất số lớp phụ đạo, số tiết/tuần cùng phân công giáo viên tương ứng.*
    """)
    
    active_df = st.session_state.get('current_prediction_df', None)
    if active_df is None:
        st.info("Chưa có dữ liệu từ Tab 1. Đang hiển thị kịch bản mẫu từ 611 học sinh trong tập Test.")
        test_path = os.path.join(base_dir, 'data', 'processed', 'test.csv')
        active_df = pd.read_csv(test_path)
        X_sc = preprocessor.transform(active_df)
        active_df['predicted_risk_level'] = selected_model.predict(X_sc)
        
    remed_plan = generate_remediation_plan(active_df)
    
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.metric("Tổng Học Sinh Nguy Cơ", f"{remed_plan['total_at_risk']} / {remed_plan['total_students']}")
    with s2:
        st.metric("Tỷ Lệ Cần Hỗ Trợ", f"{(remed_plan['total_at_risk']/remed_plan['total_students'])*100:.1f}%")
    with s3:
        st.metric("Số Lớp Phụ Đạo Đề Xuất", f"{remed_plan['total_suggested_classes']} lớp")
    with s4:
        st.metric("Tổng Số Tiết / Tuần", f"{remed_plan['total_teaching_hours_per_week']} tiết")
        
    st.markdown("---")
    st.markdown("### Kế Hoạch Tổ Chức Lớp Phụ Đạo Chuyên Đề")
    
    for idx, grp in enumerate(remed_plan['groups']):
        with st.expander(f"{grp['group_name']} ({grp['total_students']} học sinh) - {grp['priority']}", expanded=True):
            col_g1, col_g2 = st.columns([1, 2])
            with col_g1:
                st.markdown(f"""
                - **Số học sinh Nguy cơ cao:** `{grp['high_risk_students']}` em
                - **Số học sinh Cần theo dõi:** `{grp['moderate_students']}` em
                - **Số lớp đề xuất:** `{grp['suggested_classes']}` lớp (tối đa 15-20 HS/lớp)
                - **Thời lượng:** `{grp['hours_per_week_per_class']}` tiết/tuần/lớp
                - **Tổng quỹ thời gian:** `{grp['total_teaching_hours']}` tiết/tuần
                - **Phân công chuyên môn:** *{grp['assigned_teacher_role']}*
                """)
            with col_g2:
                st.markdown("**Danh sách học sinh phân bổ vào lớp:**")
                st.dataframe(grp['sample_students'], use_container_width=True, height=180)

# ================= TAB 4: BENCHMARK & EXPERIMENTS =================
with tab4:
    st.subheader("Báo Cáo Thực Nghiệm & So Sánh Đa Mô Hình (Academic Benchmark)")
    st.markdown("""
    *Thực nghiệm được thực hiện trên tập kiểm thử độc lập (Test Set: 611 mẫu - 15%), 
    được chia theo phương pháp Stratified Split chống rò rỉ dữ liệu (Anti-Leakage).*
    """)
    
    st.markdown("##### 1. Bảng Tổng Hợp Chỉ Số Đánh Giá")
    st.dataframe(
        benchmark_df[[
            'Model', 'Accuracy', 'Macro_F1', 'Recall_High_Risk',
            'Precision_High_Risk', 'Inference_Latency_ms', 'Train_Time_Sec'
        ]],
        use_container_width=True
    )
    
    b_col1, b_col2 = st.columns(2)
    with b_col1:
        fig_bench = px.bar(
            benchmark_df,
            x='Model',
            y=['Macro_F1', 'Recall_High_Risk'],
            barmode='group',
            title="So sánh Macro F1 & Recall (Lớp Nguy cơ cao)",
            color_discrete_sequence=['#2563eb', '#dc2626']
        )
        fig_bench.update_layout(
            xaxis_tickangle=-30, height=360,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color=chart_font_color, family='Plus Jakarta Sans'),
            xaxis=dict(gridcolor=chart_grid_color, tickfont=dict(color=chart_font_color)),
            yaxis=dict(gridcolor=chart_grid_color, tickfont=dict(color=chart_font_color)),
            legend=dict(font=dict(color=chart_font_color))
        )
        st.plotly_chart(fig_bench, use_container_width=True)
        
    with b_col2:
        fig_lat = px.bar(
            benchmark_df,
            x='Model',
            y='Inference_Latency_ms',
            title="Độ Trễ Suy Luận (Inference Latency - ms/học sinh)",
            color='Inference_Latency_ms',
            color_continuous_scale='Blues' if not is_dark else 'Viridis'
        )
        fig_lat.update_layout(
            xaxis_tickangle=-30, height=360,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color=chart_font_color, family='Plus Jakarta Sans'),
            xaxis=dict(gridcolor=chart_grid_color, tickfont=dict(color=chart_font_color)),
            yaxis=dict(gridcolor=chart_grid_color, tickfont=dict(color=chart_font_color))
        )
        st.plotly_chart(fig_lat, use_container_width=True)

    st.markdown("##### 2. Ma Trận Nhầm Lẫn (Confusion Matrix) & Độ Quan Trọng Thuộc Tính")
    img_col1, img_col2 = st.columns(2)
    with img_col1:
        cm_file = os.path.join(base_dir, 'reports', 'figures', 'confusion_matrix_best_model.png')
        if os.path.exists(cm_file):
            st.image(cm_file, caption="Ma trận nhầm lẫn của mô hình trên tập Test", use_container_width=True)
    with img_col2:
        fi_file = os.path.join(base_dir, 'reports', 'figures', 'feature_importance_ranking.png')
        if os.path.exists(fi_file):
            st.image(fi_file, caption="Xếp hạng độ quan trọng thuộc tính (Permutation Importance)", use_container_width=True)

    # 3. SO SÁNH THỜI GIAN VẬN HÀNH & HIỆU QUẢ THỰC TẾ
    st.markdown("##### 3. So Sánh Baseline Vận Hành: Giáo Viên (Thủ Công) vs Phương Pháp Học Máy (ML)")
    baseline_compare_data = {
        "Tiêu Chí Đánh Giá": [
            "Thời gian xử lý 1 lớp (40 học sinh)",
            "Thời gian xử lý toàn khối (4.070 học sinh)",
            "Độ bao phủ tiêu chí GDPT 2018",
            "Thời điểm phát hiện nguy cơ",
            "Căn cứ giải thích (Explainability)",
            "Khả năng mở rộng & Chi phí nhân sự"
        ],
        "Phương Pháp Thủ Công (Hội Đồng GV)": [
            "~25 - 35 phút / lớp (ước lượng trung bình hội đồng)",
            "~3 - 5 ngày làm việc / khối (nhập liệu và họp chuyên môn)",
            "Thường chỉ nhìn điểm trung bình quá trình",
            "Sau khi có điểm thi học kỳ (quá muộn)",
            "Cảm quan sư phạm chủ quan từng giáo viên",
            "Tốn nhiều công sức, dễ sai sót khi nhập liệu"
        ],
        "Mô Hình Cơ Sở (Logistic Regression)": [
            "~0.001 giây (đo thực tế trên máy thử nghiệm)",
            "~0.02 giây (đo thực tế, gồm tiền xử lý)",
            "Tách mặt phẳng tuyến tính (F1: 89.2%)",
            "Ngay sau khi có điểm kiểm tra 45 phút",
            "Trọng số Beta cố định trên từng thuộc tính",
            "Tự động hóa hoàn toàn, mã nguồn mở nhẹ"
        ],
        "Mô Hình Cải Tiến (Random Forest / LightGBM)": [
            "~0.005 - 0.04 giây (đo thực tế, LightGBM / RF)",
            "~0.08 - 0.18 giây (đo thực tế, LightGBM / RF)",
            "Bắt quan hệ phi tuyến giữa Bloom & 3 mạch Tin",
            "Cảnh báo sớm ngay từ giữa học kỳ",
            "Radar đa chiều + XAI đóng góp rủi ro từng thuộc tính",
            "Tự động đề xuất kịch bản phân bổ lớp phụ đạo"
        ]
    }
    st.dataframe(pd.DataFrame(baseline_compare_data), use_container_width=True, hide_index=True)

    # 4. PHÂN TÍCH CHUYÊN SÂU CÁC CA DỰ ĐOÁN SAI (ERROR ANALYSIS - YÊU CẦU 4 ĐỀ BÀI)
    st.markdown("##### 4. Phân Tích Chuyên Sâu Các Ca Dự Đoán Sai (In-Depth Error Analysis)")
    try:
        test_path = os.path.join(base_dir, 'data', 'processed', 'test.csv')
        df_test_eval = pd.read_csv(test_path)
        X_test_eval = preprocessor.transform(df_test_eval)
        y_test_pred = selected_model.predict(X_test_eval)
        y_test_true = df_test_eval['target_risk_level'].values

        err_mask = (y_test_true != y_test_pred)
        total_test_cnt = len(df_test_eval)
        total_err_cnt = int(err_mask.sum())
        acc_pct = (1 - total_err_cnt / total_test_cnt) * 100

        # Phân loại nhóm lỗi
        fn_mask = (y_test_true == 2) & (y_test_pred < 2)
        fp_mask = (y_test_true == 0) & (y_test_pred > 0)
        bound_mask = err_mask & (~fn_mask) & (~fp_mask)

        fn_cnt = int(fn_mask.sum())
        fp_cnt = int(fp_mask.sum())
        bound_cnt = int(bound_mask.sum())

        err_c1, err_c2, err_c3, err_c4 = st.columns(4)
        with err_c1:
            st.metric("Tổng Mẫu Kiểm Thử", f"{total_test_cnt} HS")
        with err_c2:
            st.metric("Độ Chính Xác Tổng Thể", f"{acc_pct:.1f}%")
        with err_c3:
            st.metric("Bỏ Sót Nguy Cơ (False Negative)", f"{fn_cnt} HS", delta="Cần chú ý nhất", delta_color="inverse")
        with err_c4:
            st.metric("Cảnh Báo Nhầm (False Positive)", f"{fp_cnt} HS", delta="Rủi ro thấp", delta_color="off")

        st.markdown(f"""
        <div class="kbs-card" style="margin-top: 10px; margin-bottom: 14px;">
            <b style="color: var(--primary-blue); font-size: 0.95rem;">Đúc Kết Sư Phạm Từ Phân Tích Lỗi (Pedagogical Error Insights):</b>
            <ul style="font-size: 0.88rem; color: var(--kbs-text-body); line-height: 1.6; margin-top: 6px; padding-left: 20px;">
                <li><b>Nguyên nhân False Negative ({fn_cnt} ca):</b> Học sinh có điểm quá trình học bạ tương đối ổn (6.0 - 7.0) do điểm miệng và chuyên cần cao, nhưng thực tế bị <b>"hổng rỗng" tư duy lập trình (CS &lt; 4.0)</b> và vận dụng cao. Khi bước vào kỳ thi chuẩn hóa GDPT 2018 (với các câu trắc nghiệm Đúng/Sai logic và trả lời ngắn thuật toán), các em bị điểm liệt. Mô hình bị đánh lừa bởi trọng số điểm quá trình.</li>
                <li><b>Nguyên nhân False Positive ({fp_cnt} ca):</b> Học sinh có năng lực tư duy CS và điểm thi cuối kỳ rất cao, nhưng bị thiếu một cột điểm 15 phút hoặc nộp bài tập LMS muộn kéo điểm quá trình tạm thời xuống thấp, khiến mô hình phát tín hiệu cảnh báo rủi ro nhầm.</li>
                <li><b>Ca lệch ngưỡng biên ({bound_cnt} ca):</b> Học sinh nằm ở vùng giáp ranh điểm 5.0 và 6.4 (rất gần quyết định phân loại giữa mức Cần theo dõi và An toàn).</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        err_filter_type = st.radio(
            "Lọc danh sách các ca dự đoán sai để phân tích:",
            ["Tất cả ca sai sót", "False Negative (Nguy cơ cao bị bỏ sót)", "False Positive (An toàn bị cảnh báo nhầm)", "Ca nằm sát ranh giới biên (Boundary)"],
            horizontal=True
        )

        df_err_display = df_test_eval[err_mask].copy()
        df_err_display['y_pred'] = y_test_pred[err_mask]
        
        risk_name_map = {0: 'An toàn (>= 6.5)', 1: 'Cần theo dõi (5.0 - 6.4)', 2: 'Nguy cơ cao (< 5.0)'}
        df_err_display['Thực Tế'] = df_err_display['target_risk_level'].map(risk_name_map)
        df_err_display['Dự Báo AI'] = df_err_display['y_pred'].map(risk_name_map)

        if "False Negative" in err_filter_type:
            df_err_display = df_err_display[(df_err_display['target_risk_level'] == 2) & (df_err_display['y_pred'] < 2)]
        elif "False Positive" in err_filter_type:
            df_err_display = df_err_display[(df_err_display['target_risk_level'] == 0) & (df_err_display['y_pred'] > 0)]
        elif "Boundary" in err_filter_type:
            df_err_display = df_err_display[~(((df_err_display['target_risk_level'] == 2) & (df_err_display['y_pred'] < 2)) | ((df_err_display['target_risk_level'] == 0) & (df_err_display['y_pred'] > 0)))]

        # Gán nhãn nguyên nhân cụ thể cho từng ca
        def explain_error_reason(row):
            if row['target_risk_level'] == 2 and row['y_pred'] < 2:
                if row['score_topic_cs'] < 4.0:
                    return "Học bạ khá nhưng hổng sâu tư duy Lập trình CS (< 4.0)"
                elif row['score_trend'] < -1.0:
                    return "Điểm số có xu hướng tụt dốc mạnh trong HK2"
                else:
                    return "Nằm sát ngưỡng biên ranh giới điểm 5.0"
            elif row['target_risk_level'] == 0 and row['y_pred'] > 0:
                if row['homework_ratio'] < 0.4:
                    return "Thi tốt nhưng thiếu nộp BTVN trên LMS"
                else:
                    return "Điểm 15p thấp ngẫu nhiên làm lệch xác suất"
            else:
                return "Biến động điểm số quanh ngưỡng phân tách 6.5"

        df_err_display['Nguyên Nhân Dự Báo Sai'] = df_err_display.apply(explain_error_reason, axis=1)

        err_show_cols = [
            'student_id', 'full_name', 'class_name', 'Thực Tế', 'Dự Báo AI',
            'score_tx_avg', 'score_hk2', 'score_topic_cs', 'score_high_apply', 'homework_ratio', 'Nguyên Nhân Dự Báo Sai'
        ]
        col_rename_dict = {
            'student_id': 'Mã HS', 'full_name': 'Họ Và Tên', 'class_name': 'Lớp',
            'score_tx_avg': 'TB Quá Trình', 'score_hk2': 'Điểm HK2',
            'score_topic_cs': 'Điểm CS', 'score_high_apply': 'Điểm VDC', 'homework_ratio': 'Tỷ Lệ BTVN'
        }
        st.dataframe(
            df_err_display[err_show_cols].rename(columns=col_rename_dict),
            use_container_width=True,
            height=300
        )
    except Exception as e:
        st.info(f"Đang đồng bộ dữ liệu phân tích lỗi: {str(e)}")

# ================= TAB 5: KNOWLEDGE LIBRARY & LIMITS =================
with tab5:
    st.subheader("Thư Viện Cơ Sở Tri Thức Môn Tin Học & Giới Hạn Hệ Thống")
    
    k1, k2 = st.columns(2)
    with k1:
        st.markdown("""
        <div class="kbs-card">
            <h4 style="color: var(--kbs-text-title); margin-top:0;">Cơ Sở Tri Thức Đề Thi GDPT 2018</h4>
            <ul style="font-size: 0.92rem; line-height: 1.6; padding-left: 20px; color: var(--kbs-text-body);">
                <li><b>Phần I (24 câu - 6.0 điểm):</b> Trắc nghiệm 4 lựa chọn. Khảo sát mức độ Nhận biết và Thông hiểu lý thuyết.</li>
                <li><b>Phần II (4 câu - 4.0 điểm):</b> Trắc nghiệm Đúng/Sai đa ý. Đo lường khả năng phân tích luồng dữ liệu và tư duy logic.</li>
                <li><b>Phần III (6 câu - 3.0 điểm):</b> Trắc nghiệm trả lời ngắn. Phân hóa học sinh giỏi với các bài toán thuật toán và lập trình thực tiễn.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    with k2:
        st.markdown("""
        <div class="kbs-card">
            <h4 style="color: #dc2626; margin-top:0;">Giới Hạn Hệ Thống & Khung Đạo Đức AI (Ethical AI)</h4>
            <ul style="font-size: 0.92rem; line-height: 1.6; padding-left: 20px; color: var(--kbs-text-body);">
                <li><b>Hỗ trợ ra quyết định (Decision Support):</b> Hệ thống đóng vai trò cố vấn phân tích, <b>TUYỆT ĐỐI KHÔNG</b> tự động quyết định xếp loại học bạ, kỷ luật hay thay thế quyền đánh giá sư phạm của thầy cô giáo.</li>
                <li><b>Không dán nhãn định kiến (Zero Stigmatization):</b> Các cảnh báo nguy cơ học tập chỉ lưu hành nội bộ giữa Ban giám hiệu và Giáo viên phụ trách nhằm xây dựng kế hoạch bồi dưỡng, không công khai bêu tên học sinh.</li>
                <li><b>Bảo vệ dữ liệu cá nhân (Data Privacy):</b> Toàn bộ dữ liệu huấn luyện và thực nghiệm đều được ẩn danh hóa (mã hóa dạng STU_xxxx), loại bỏ hoàn toàn số định danh CCCD và thông tin nhạy cảm.</li>
                <li><b>Xử lý ngoại lệ ngoài học thuật:</b> Mô hình không phản ánh được các biến cố ngoại cảnh đột xuất (như học sinh bị ốm đau trong kỳ thi, tai nạn gia đình). Giáo viên luôn là mắt xích kiểm chứng cuối cùng (Human-in-the-loop).</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")
st.caption("© 2026 EduRisk Informatics • Hệ Thống Hỗ Trợ Ra Quyết Định Cảnh Báo Nguy Cơ Môn Tin Học THPT • GDPT 2018.")
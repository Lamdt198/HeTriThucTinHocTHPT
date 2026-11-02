# Hệ tri thức & học máy cảnh báo sớm nguy cơ học tập môn Tin học THPT (GDPT 2018)

> **Chương trình:** Cao học Công nghệ Thông tin – Lớp K20.2
> **Học phần:** Các Hệ Cơ Sở Tri Thức (Knowledge-Based Systems)
> **Hình thức:** Đề tài nhóm (2 – 4 học viên), thời gian 4 tuần

Ứng dụng web (Streamlit) hỗ trợ giáo viên và Ban giám hiệu **phát hiện sớm học sinh có nguy cơ** đạt kết quả thấp ở môn Tin học thi tốt nghiệp THPT, **giải thích căn cứ** của từng cảnh báo (XAI) và **đề xuất kịch bản phân bổ lớp phụ đạo**.

> [!WARNING]
> **Dữ liệu là GIẢ LẬP.** Bộ dữ liệu 4.070 học sinh / 8 khóa được sinh bằng mô phỏng thống kê bám theo cấu trúc đánh giá THPT và ma trận GDPT 2018. Mọi số liệu đánh giá bên dưới chỉ chứng minh tính khả thi của quy trình (Proof of Concept), **không dùng để khẳng định hiệu quả trên học sinh thực tế**.

---

## 1. Bài toán

| Mục | Mô tả |
|---|---|
| **Người dùng** | Giáo viên bộ môn Tin học; Tổ trưởng chuyên môn / Ban giám hiệu |
| **Loại bài toán** | Phân loại có giám sát, đa lớp (3 lớp), mất cân bằng nhẹ |
| **Đầu vào (15 đặc trưng)** | 5 cột điểm THPT (Miệng, 15 phút, 45 phút, Học kỳ 1, Học kỳ 2) · TB quá trình · Xu hướng (HK2 − HK1) · Tỷ lệ BTVN trên LMS · 4 mức Bloom (Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao) · 3 mạch GDPT 2018 (CS, ICT, DL) |
| **Đầu ra** | `0` An toàn (≥ 6.5) · `1` Cần theo dõi (5.0 – 6.4) · `2` Nguy cơ cao (< 5.0) + xác suất từng lớp |
| **Phạm vi** | Hệ **hỗ trợ ra quyết định** – không tự động xếp loại, kỷ luật hay thay thế đánh giá của giáo viên |

---

## 2. Dữ liệu

- **Quy mô:** 4.070 học sinh, 8 khóa (2018–2019 → 2025–2026), lưu trong SQLite nhúng `data/vstep_informatics.db` và Excel `data/raw/du_lieu_8_khoa_tin_hoc_thpt.xlsx`.
- **Ẩn danh:** mã học sinh dạng `STU_Kxx_xxxx`; họ tên được sinh ngẫu nhiên, không phải người thật.
- **Phân bố nhãn:** An toàn 37,0% · Cần theo dõi 22,9% · Nguy cơ cao 40,1%.
- **Chia tập (Stratified Split):** Train 2.849 (70%) · Validation 610 (15%) · Test 611 (15%).
- **Chống rò rỉ dữ liệu:** `SimpleImputer` + `StandardScaler` chỉ `fit` trên tập Train; Validation/Test chỉ `transform`. Điểm thi cuối (`est_final_score`) dùng để tạo nhãn **không** được đưa vào đặc trưng.
- **Kiểm tra đầu vào:** file tải lên được kiểm tra cột bắt buộc, kiểu số, giá trị trong [0, 10] (tỷ lệ BTVN trong [0, 1]); lỗi được hiển thị cảnh báo trên giao diện.

---

## 3. Kết quả trên tập Test độc lập (611 mẫu)

Nguồn: [`models_saved/benchmark_results.csv`](models_saved/benchmark_results.csv)

| Mô hình | Accuracy | Macro F1 | Recall Nguy cơ cao | Precision Nguy cơ cao | Huấn luyện |
|---|:---:|:---:|:---:|:---:|:---:|
| Baseline: Rule-Based Heuristic | 84.12% | 82.83% | 78.78% | 97.47% | – |
| Baseline: Decision Tree (depth 4) | 85.60% | 84.31% | 87.35% | 94.69% | 0.005 s |
| **Baseline: Logistic Regression** | **90.34%** | **89.19%** | 91.84% | 95.74% | 0.013 s |
| Improved: Random Forest | 88.71% | 87.31% | 91.02% | 94.89% | 0.333 s |
| Improved: XGBoost | 88.38% | 86.64% | 91.84% | 92.98% | 0.258 s |
| **Improved: LightGBM** | 89.36% | 87.72% | **93.88%** | 93.12% | 0.096 s |

**Nhận xét trung thực**
- Logistic Regression (baseline) có Macro F1 cao nhất → quan hệ đặc trưng – nhãn trong dữ liệu giả lập gần tuyến tính.
- LightGBM có **Recall lớp Nguy cơ cao** tốt nhất (93,88%) – chỉ số được ưu tiên vì bỏ sót học sinh nguy cơ tốn kém hơn cảnh báo nhầm.
- Không mô hình nào nhầm *Nguy cơ cao ↔ An toàn*; lỗi chỉ xảy ra giữa các lớp kề nhau.

**Phân tích lỗi (Random Forest, 69/611 ca sai)**
- 22 ca *Nguy cơ cao → Cần theo dõi* (bỏ sót): điểm quá trình khá nhưng điểm mạch CS / Vận dụng cao thấp.
- 22 ca *An toàn → Cần theo dõi* (cảnh báo nhầm): thiếu điểm 15 phút hoặc tỷ lệ BTVN thấp.
- 25 ca lệch biên giữa các ngưỡng 5.0 / 6.5.

**Thời gian (so sánh baseline)**
- Mô hình xử lý 4.070 học sinh trong ~0,02 s (Logistic) đến ~0,18 s (Random Forest), **đo thực tế** trên laptop cá nhân.
- Thời gian rà soát thủ công (~25–35 phút/lớp) là **giả định ước lượng** của nhóm, chưa khảo sát.

---

## 4. Ứng dụng (5 tab)

| Tab | Chức năng |
|---|---|
| 📁 Dự báo theo lớp | Nạp khóa học từ SQLite hoặc tải Excel/CSV mới → kiểm tra hợp lệ → dự báo, biểu đồ phân bố, lọc theo mức nguy cơ, tải kết quả `.xlsx` |
| 👤 Chẩn đoán & Radar (XAI) | Nhập điểm 1 học sinh (hoặc chọn 3 mẫu có sẵn) → mức nguy cơ, xác suất, radar năng lực so với chuẩn An toàn, biểu đồ đóng góp rủi ro, khuyến nghị sư phạm |
| 🏫 Phân bổ nguồn lực | Gom nhóm học sinh theo lỗ hổng chính → số lớp phụ đạo, số tiết/tuần, phân công giáo viên |
| 📊 Benchmark | Bảng chỉ số, biểu đồ so sánh, ma trận nhầm lẫn, độ quan trọng đặc trưng, so sánh thời gian, bảng tra cứu ca dự đoán sai |
| ℹ️ Tri thức & Giới hạn | Cấu trúc đề thi GDPT 2018, giới hạn hệ thống, nguyên tắc đạo đức AI |

---

## 5. Cài đặt & chạy

Yêu cầu: Python 3.10+ (đã thử trên 3.12).

```bash
pip install -r requirements.txt
streamlit run app/app.py        # hoặc nhấp đúp run_app.bat (Windows)
```

Mở trình duyệt tại `http://localhost:8501`.

**Tái lập toàn bộ pipeline (tùy chọn)**

```bash
python src/generate_multi_cohort_data.py   # sinh dữ liệu 8 khóa + nạp SQLite
python src/preprocessor.py                 # tiền xử lý, chia Train/Val/Test
python src/models_trainer.py               # huấn luyện + benchmark
python src/xai_explainer.py                # độ quan trọng đặc trưng, biểu đồ
python build_academic_report.py            # xuất báo cáo Word
```

---

## 6. Cấu trúc thư mục

```text
HeTriThucTinHocTHPT/
├── .streamlit/config.toml        # Theme giao diện
├── app/app.py                    # Ứng dụng Streamlit (5 tab)
├── src/
│   ├── database.py               # CSDL SQLite: truy vấn khóa học, ghi nhật ký dự báo
│   ├── generate_multi_cohort_data.py  # Sinh dữ liệu giả lập 8 khóa
│   ├── data_generator.py         # Bộ sinh dữ liệu phiên bản đầu
│   ├── preprocessor.py           # Kiểm tra đầu vào, chuẩn hóa, chia tập chống rò rỉ
│   ├── models_trainer.py         # Huấn luyện baseline + cải tiến, đo thời gian
│   ├── xai_explainer.py          # Permutation Importance, giải thích từng học sinh
│   └── resource_allocator.py     # Kịch bản phân bổ lớp phụ đạo
├── data/
│   ├── vstep_informatics.db      # SQLite 4.070 học sinh
│   ├── raw/                      # Excel/CSV gốc, mẫu upload theo từng khóa
│   └── processed/                # train.csv / val.csv / test.csv
├── models_saved/                 # Trọng số mô hình .joblib, benchmark, feature importance
├── reports/                      # Báo cáo Word + biểu đồ
├── build_academic_report.py      # Sinh báo cáo Word
├── de_cuong_hoc_phan_k20_2.md    # Đề cương & phân công 4 tuần
├── requirements.txt
└── run_app.bat
```

---

## 7. Phân công thành viên

Xem chi tiết trong [`de_cuong_hoc_phan_k20_2.md`](de_cuong_hoc_phan_k20_2.md) và báo cáo Word. *(Điền họ tên, mã học viên thật trước khi nộp.)*

| Thành viên | Phụ trách chính |
|---|---|
| Học viên 1 (Trưởng nhóm) | Đặc tả bài toán, CSDL SQLite, tiền xử lý & chống rò rỉ, mô hình baseline |
| Học viên 2 | Mô hình cải tiến (RF, XGBoost, LightGBM), đánh giá đa chỉ số |
| Học viên 3 | Ứng dụng Streamlit, XAI, kịch bản phân bổ nguồn lực |
| Học viên 4 | Phân tích lỗi, so sánh thời gian, đạo đức AI, báo cáo & slide |

---

## 8. Giới hạn

- Dữ liệu giả lập → chưa kiểm định trên học sinh thực; cần dữ liệu thật đã ẩn danh trước khi dùng thực tế.
- Ngưỡng nhãn (5.0 / 6.5) là quy ước của nhóm.
- Mô hình không phản ánh yếu tố ngoài học thuật (sức khỏe, hoàn cảnh gia đình); giáo viên luôn là người quyết định cuối cùng.
- Phần giải thích XAI là xấp xỉ mức đóng góp, không phải quan hệ nhân quả.

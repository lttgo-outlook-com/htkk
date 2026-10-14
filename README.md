# HTKK — Trích xuất Biểu mẫu, Templates, XML & Công thức

## 1. Mục tiêu cốt lõi
Kho dữ liệu này tập trung **trích xuất, chuẩn hóa, tra cứu và theo dõi thay đổi (track changes) các biểu mẫu thuế** từ bộ cài HTKK của Tổng cục Thuế:
- **Biểu mẫu & Templates**: `.xls`, `.xls1`, `.xlsx` (giao diện bảng tính, tiêu đề cột, hàng, định dạng, Data Validation).
- **Ánh xạ & Schema XML**: `.xml` mẫu ánh xạ ô (`CellID` ⇄ `Path` chỉ tiêu XML), `.xsd` (schema kiểm tra cấu trúc hồ sơ khai thuế), `menu.xml` (danh mục tờ khai, kỳ tính thuế, phụ lục).
- **Danh mục & Thông báo**: `catalogue_*`, `dm_*`, `message.xml`.
- **Công thức**: `cong_thuc.txt` trong từng thư mục biểu mẫu và `_spec/formulas/`.

## 2. Nguyên tắc làm việc (Tránh làm việc dư thừa)
- **KHÔNG tập trung decompile để rebuild app**: Không decompile lại `.exe`/`.dll` (`ilspycmd`) và không chạy lại các bộ phân tích AST C# nặng (`RoslynDump`, `RulesExtract`, `RoslynValidations`) khi cập nhật phiên bản HTKK mới hoặc khi làm việc với mẫu biểu, trừ khi có yêu cầu cụ thể từ người dùng.
- **Chỉ xử lý dữ liệu phục vụ biểu mẫu**: Khi có bản HTKK mới, bỏ qua các tệp thực thi (`.exe`, `.dll`, `.pdb`, `.config`) và chỉ tập trung vào những tệp biểu mẫu/cấu hình thực sự thay đổi trong `Data1.cab`.

## 3. Cấu trúc thư mục chính
- `_forms/`: Toàn bộ biểu mẫu đã phân nhóm theo nghiệp vụ (`01_GTGT`, `02_TNDN`, `06_BCTC`, `11_TNCN`, ...), kèm `_manifest.csv` và `cong_thuc.txt` cho từng tờ khai.
- `_extracted/`: Các tệp gốc giải nén phẳng từ `Data1.cab` của bộ cài HTKK.
- `TT58_BCTC/`: Bộ tài liệu, biểu mẫu, công thức và template trích xuất riêng cho Báo cáo tài chính Thông tư 58/2026/TT-BTC.
- `_mau_tt89/`: Các mẫu biểu theo Thông tư 89/2026/TT-BTC.
- `_spec/`: Dữ liệu đặc tả JSON/CSV đã trích xuất (`formulas/`, `xsd/`, `dv/`, `catalog/`, `config/`...).
- `_re/`: Script ánh xạ biểu mẫu (`reorg_forms.py`, `check_mapping.py`) và tài liệu tra cứu cấu trúc (`NOTES.md`, `CONTROLTYPE_MAP.md`, `DOC_ENGINE.md`).

## 4. Quy trình chuẩn khi cập nhật bản HTKK mới (ví dụ 5.7.9 → 5.8.0)
Khi có thư mục bộ cài mới (ví dụ `HTKK_v5.8.0_DLL_signed/Data1.cab`):

1. **Giải nén & So sánh file biểu mẫu (Track Changes)**:
   - Giải nén `Data1.cab` ra thư mục tạm (`/tmp/...`).
   - So sánh SHA-256 với `_extracted/` **chỉ trên các tệp dữ liệu/biểu mẫu** (`.xls`, `.xls1`, `.xml`, `.xsd`, `menu.xml`, `catalogue_*`, `dm_*`, `message.xml`). **Bỏ qua** `.exe`, `.dll`, `.pdb`, `.config`.
2. **Cập nhật `_extracted/` và `_forms/`**:
   - Copy các tệp biểu mẫu có thay đổi vào `_extracted/`.
   - Chạy `python3 -I _re/reorg_forms.py --write` để đồng bộ sang `_forms/` và cập nhật `_forms/_manifest.csv`.
3. **Cập nhật công thức Excel (`_spec/formulas/` & `cong_thuc.txt`)**:
   - Chạy trích xuất công thức BIFF:
     ```sh
     /tmp/audit_venv/bin/python -I _spec/tools/biff_formulas.py --src _extracted --menu _extracted/menu.xml --out _spec/formulas
     ```
   - Sinh lại `cong_thuc.txt`:
     ```sh
     /tmp/audit_venv/bin/python -I _spec/tools/render_cong_thuc.py --repo .
     ```
   - Kiểm tra tính toàn vẹn:
     ```sh
     python3 -I _spec/tools/verify_cong_thuc.py --excel <số_dòng_excel_mới>
     ```
4. **Báo cáo Track Changes chi tiết**:
   - Liệt kê rõ tờ khai/phụ lục nào thay đổi, thêm/bớt/dịch cột nào trong `.xls`, thay đổi `CellID`/`Path`/`Controltype` nào trong `.xml`, và công thức nào thay đổi trong `cong_thuc.txt`.

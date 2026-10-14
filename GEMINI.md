# Quy tắc làm việc cho AI Agent (GEMINI.md / CLAUDE.md)

## Trọng tâm công việc
1. **Tập trung vào trích xuất dữ liệu biểu mẫu**:
   - Mục tiêu chính của repo là trích xuất và theo dõi **forms, templates (`.xls`, `.xls1`, `.xlsx`), `.xml` ánh xạ chỉ tiêu, `.xsd` schema, danh mục (`menu.xml`, `catalogue_*`, `dm_*`), và công thức (`cong_thuc.txt`, `_spec/formulas/`)**.
2. **KHÔNG làm việc dư thừa (Không decompile để rebuild app)**:
   - **KHÔNG** tự ý decompile lại `.dll`/`.exe` (`ilspycmd`) hoặc cập nhật `_decompiled/`.
   - **KHÔNG** tự ý chạy các công cụ phân tích C#/Roslyn nặng (`RoslynDump`, `RulesExtract`, `RoslynValidations`) hay cập nhật các bảng rule C# (`_spec/csharp/`, `_spec/rules/`, `_spec/validations/`, `_spec/app/`) khi người dùng chỉ yêu cầu cập nhật/tra cứu biểu mẫu, templates, công thức.
   - **KHÔNG** copy đè các file `.exe`, `.dll`, `.pdb`, `.config` vào `_extracted/` khi cập nhật phiên bản mới trừ khi được yêu cầu rõ.
3. **Quy trình cập nhật phiên bản HTKK mới**:
   - Giải nén `Data1.cab` ra `/tmp`, đối chiếu SHA-256 chỉ các tệp biểu mẫu/cấu hình (`.xls`, `.xls1`, `.xml`, `.xsd`, `menu.xml`, `catalogue_*`, `dm_*`, `message.xml`).
   - Copy các tệp biểu mẫu thay đổi vào `_extracted/`.
   - Chạy `python3 -I _re/reorg_forms.py --write` để cập nhật `_forms/` và `_forms/_manifest.csv`.
   - Chạy `biff_formulas.py` và `render_cong_thuc.py` để cập nhật `_spec/formulas/` và `_forms/**/cong_thuc.txt`.
   - Báo cáo **Track Changes** trực diện vào nội dung biểu mẫu (thay đổi hàng/cột trên sheet `.xls`, thay đổi `CellID`/`Path` trong `.xml`, thay đổi công thức trong `cong_thuc.txt`).

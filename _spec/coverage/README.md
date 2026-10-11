# _spec/coverage

Sinh bởi `nice -n 15 python3 -I _spec/tools/coverage.py` (chạy từ gốc worktree). Hai lần chạy cho kết quả giống từng byte.

Quy ước:
- Form = (MaFile, VersionXML) theo menu.xml (242 form). Tệp `<MaFile>_<VersionXML>.json`.
- Ô dữ liệu = ô có Path trong tệp Sections (XML mẫu) và đường dẫn đầy đủ có trong XSD (T5). Ô có Path nhưng không vào XSD nằm ở `path_ngoai_xsd` (ô trợ giúp như path_5, null_1).
- CellID = ô trong .xls (mẫu giao diện, TemplateRowStart); CellID2 = ô trong .xls1 (mẫu báo cáo, ReportRowStart). C# chỉnh bảng tính giao diện (frmWinToKhai.cs:417 nạp InterfaceTemplateExcel), nên toạ độ C# và Validate được khớp theo CellID (xls). Công thức xls1 và DV khớp theo CellID2.
- Ô bảng động (Dynamic=1, RowInfo) là một mục theo cột (`COT:<cột>`); toạ độ literal của C#/Validate chỉ khớp ở hàng bắt đầu dữ liệu. Ô Header/Footer của bảng động dùng toạ độ tĩnh trong mẫu (có thể lệch khi bảng dài ra), gắn `khu_vuc`.
- StatusID: không có mã C# nào đọc StatusID; chỉ có hằng HTKKConstant.cs:49-51 (0001 = LOAD_DT_YES, 0002 = LOAD_DT_NO). Script dùng 0001 làm "ô nhập" (heuristic); 35 nghìn ô không có StatusID.
- Nhãn: nhap_tay = không có nguồn tính (xls, xls1, rowreplace, setFormula khác rỗng của C#); tinh_tu_dong = có nguồn tính, không phải ô nhập; ca_hai = có nguồn tính và (StatusID 0001 hoặc có unlock).
- setValue của C# chỉ ghi `csharp_setvalue` (số lần), không tính là nguồn tính. setFormula rỗng (reset) không tính là công thức.
- Không có form cho 152 lớp C#/T3 ngoài menu (phiên bản cũ) và lớp dùng chung: đếm ở `conflicts.json > unmapped`.
- Sheet KHBS: không có tệp Sections cho phiên bản của form, nên không có ô; lý do từng sheet ở `sheet_log.json` (nhom_khbs).
- Lệch tên sheet xls so với menu: `phan_loai` = chi_khac_ten (đúng chỉ số) hoặc lech_chi_so.

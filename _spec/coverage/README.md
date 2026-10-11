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

Cập nhật vòng sửa 1:
- XSD: sheet chính dùng XSD chính (origin "menu.xml"), không dùng XSD KHBS (origin "menu.xml (VersionKHBS)"). Trước sửa 140 form dùng nhầm XSD KHBS cho sheet chính; sau sửa 0. 7 form không có XSD chính: 4 form _VIEW dùng XSD của form gốc không _VIEW (header `xsd_thay_the_form_goc`), 3 form còn lại không có XSD (ô vào path_ngoai_xsd).
- Đường dẫn ô tĩnh lặp đoạn (PL_x/PL_x/...) đã sửa: nếu Path bắt đầu bằng tên nút cuối của XMLnode thì ghép với nút cha.
- excel_khac_csharp: chỉ so lần ghi cuối của mỗi (hook, tệp, guards, when); C# trỏ ô phụ (bare ref) có công thức T1 bằng công thức Excel của ô đích thì coi là tương đương. 14 giảm còn 8.
- khoa_nhung_la_o_nhap tách khoa_co_dieu_kien (mọi lần khoá có guards hoặc trigger.when; ghi guard_khoa). Ô có setValue hằng ghi `co_dinh_gia_tri`.
- KHBS: sheet KHBS không có Sections được phủ theo toạ độ (sheet, ô) trong `khbs_cells_theo_toa_do` (từ T1 xls/xls1, và T2/T3 nếu có literal). Nguồn mẫu: sheet KHBS trong .xls của chính form. Quy tắc thay sheet (WinFnProcess.cs:69-80): KieuKHBS=="3" và LoaiToKhai=="B" thì 2 sheet cuối trước Header bị thay bằng sheet 0,1 của mẫu chung KHBS_TT80(_DVT)/khbs_tt80_<VersionKHBS>.xls; ô của mẫu chung CHƯA phủ. Không có Path/XSD cho ô KHBS.
- Effect có sheet `activeSheet()` và guard duy nhất `activeSheet() == N` được ánh xạ về sheet N.
- Ô dòng động: kiểm cột hợp lệ và có hàng mẫu cho mọi ô; khớp T1 ở hàng mẫu ghi trong stdout (dyn_*). C# literal chỉ khớp đúng hàng mẫu; effect literal ở hàng sau hàng mẫu của bảng động vẫn không ánh xạ (không xác định được độ dài bảng).

Cập nhật vòng sửa 2:
- Tệp T1: đuôi tên nguồn bị cắt cụt (.xl, .xl1, .x1) được phân loại đúng (xls/xl = mẫu .xls; xls1/xl1/x1 = mẫu .xls1; xlsx bảng kê bỏ). Trước sửa 5 khoá (mafile, version) bị ghi đè cùng khe: 209_B01_DTDP_GND_BCTC 2.9.0 và 334_B01_CTCK_BCTC_GND 2.3.4 (.xl bị .xl1 đè), 01_CCTT_DLVC 3.0.5 (.xlsx đè .xls), 01_TTDB 2.6.1 (tệp view), 02_PTHU_VSP_TT80 2.5.1 (tệp tên có "_"); sau sửa 0 (tệp trùng khoá giữ ưu tiên không-view, không "_", tên nhỏ nhất).
- Path bắt đầu từ gốc XSD khác HSoThueDTu (ví dụ TKhai/DLTKhai/...) được khớp trực tiếp; đoạn lặp liền kề bị gộp nếu không khớp XSD. 0 path còn đoạn lặp liền kề.
- XSD KHBS: phép thử `_KHBS_` không phân biệt hoa thường.
- Mẫu chung KHBS (WinFnProcess.cs:74): KieuKHBS=="3" và LoaiToKhai=="B" thì sheet Count-3 và Count-2 của .xls form được thay bằng sheet 0 và 1 của khbs_tt80_<VersionKHBS> (MenuID thuộc chuỗi "I7,I4,C7,I8,C8" theo string.Contains: KHBS_TT80_DVT; thuộc "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5": KHBS_DIA_BAN; còn lại KHBS_TT80). Ô theo toạ độ ghi ở `mau_chung_khbs_cells` (nguon "mau_chung_khbs"), tách khỏi `khbs_cells_theo_toa_do` (sheet KHBS riêng của form). Header `mau_chung_khbs` ghi mẫu, phiên bản, tệp T1; nếu không tìm thấy phiên bản thì ghi `ung_vien`. LoaiToKhai là giá trị lúc chạy nên mẫu chung chỉ áp dụng khi nộp bổ sung.

# Data validation (DV) trong template .xls của HTKK

Kết luận cổng (cập nhật T9): CÓ MỘT PHẦN, theo tài liệu API FarPoint, chưa xác nhận ở mức mã. Thư viện nạp DV dạng danh sách (list) thành combo của ô khi OpenExcel đọc BIFF; chưa xác định với DV số, độ dài, thông báo nhập/lỗi.

Bằng chứng (đường dẫn gốc /Users/tuanle/codes/HTKK/_decompiled):
- `grep -rlI DataValidation _decompiled` cho 0 tệp (brief nêu 3 tệp: không tái hiện được). Chỉ `htkk/HTKK/frmWinToKhai.cs:46` khớp "datavalid" không phân biệt hoa thường: biến cờ `isDataValid`, không liên quan DV của Excel. Không có `Validator`, `DataValidator` nào trong mã.
- `htkk/HTKK/WinFnProcess.cs:69`: `winFpSpread.OpenExcel(TemplateFile, ExcelOpenFlags.DoNotRecalculateAfterLoad)`; không có mã đọc, sửa hay xóa DV sau khi nạp. Các lệnh nạp khác: `frmWinToKhai.cs:1539`, `frmWinHTKKOpenFile.cs:139,366,1308` (DataOnly).
- Cơ chế "validate" của app là tự viết: `tax.htkk.utilities/TAX.HTKK.Utilities.Core/HTKKFile.cs:458-519` (`RemoveValidate*` chỉ đổi BackColor, NoteStyle, ResetNote của ô); hàng trăm lớp nghiệp vụ gọi `RemoveValidate*` (vd `DkTain01Nckt224.cs:306-442`).
- Sửa khẳng định cũ ("DLL đóng, không có trong máy" là sai): `_extracted/` có `farpoint.excel.dll`, `farpoint.win.spread.dll`, `farpoint.excel.xml` (tài liệu API), `farpoint.win.spread.xml`. Không có ilspycmd, monodis, ikdasm offline (`which`: không thấy; `~/.dotnet/tools` không tồn tại) nên không decompile; chỉ đọc XML tài liệu và dò chuỗi byte trong DLL (không nạp, không chạy).
- Bằng chứng thư viện: `_extracted/farpoint.excel.xml:35022-35032` giao diện `IExcelBiffRead.ExcelSetValidationData(valSheet, inputSheet, inputRgn, valItems, valRgnList, dwDVFlags)` ("validation combo"): bộ đọc BIFF gọi vào host để gắn validation. `_extracted/farpoint.win.spread.xml:31996-32006` Spread triển khai `Excel.ExcelSetValidationData(..., isEditable)`; `:32007-32016` `ExcelGetValidationData` trả `m_dvStringArr/RowArr/ColumnArr/IsEditable` ("cell validation combo"). `farpoint.excel.dll` chứa chuỗi "DVAL". `farpoint.excel.xml:37728-37752`: `ExcelOpenFlags.DataOnly` = 1 chỉ nạp dữ liệu; `DoNotRecalculateAfterLoad` = 1024 chỉ tắt tính lại, không loại DV.
- Giới hạn: API chỉ mô tả danh sách giá trị và cờ combo (ô soạn được hay không). Không có API nào cho DV kiểu whole/decimal/textLength, toán tử, thông báo nhập/lỗi; nên (suy luận, chưa kiểm) các kiểu này bị bỏ khi nạp. 1.260 bản ghi gồm chỉ 141 list; 1.105 bản ghi còn lại (số, độ dài, any) khả năng không có tác dụng trong HTKK. Cần decompile hoặc chạy thử để chốt.
- `_re/DOC_ENGINE.md` mục 3.4 ("không có DV") chỉ dựa 5 tệp mẫu; không đúng với toàn bộ template.

Số bản ghi DV tồn tại trong template (gốc `_forms`):
- 392 tệp .xls/.xls1 (trên 942 tệp quét) có DV; 1.260 bản ghi DV (0x01BE); 876 bản ghi DVAL (0x01B2) (sinh lại sau reorg `_forms`; trước reorg 388/932/1.245/867).
- Đếm độc lập (quét thô mọi record trong stream, khác script giải mã) làm trên tập `_forms` trước reorg (388 tệp / 1.245 bản ghi): khớp với qc. Tập hiện tại: `_spec/tools/check_counts.py dv` đối chiếu tổng với từng tệp JSON, `rebuild_all.sh` chạy kiểm này.
- Loại: decimal 819, whole 208, list 141, textLength 60, any 32. Chỉ 3 bản ghi list là chuỗi literal (fStrLookup); phần còn lại trỏ NAME hoặc vùng ô. Bản ghi list literal trong `15_THUE_KHOAN_THUE_TAI_SAN/01_TTS_TT40/bangke_01_2_tts_tt40.xls` thực chất là chuỗi công thức INDEX/MATCH dùng NUL thay dấu phẩy (do nguồn tạo; giải mã trung thực).
- `_extracted` có 394 tệp / 1.264 bản ghi (2 tệp `_3_tain.xls`, `template.xls` không phải OLE2); chênh +6 tệp, +19 bản ghi so với `_forms`; không xuất JSON cho `_extracted`.
- Cấp độ trùng: theo hash nội dung, `_forms` có 354 tệp DV duy nhất, 1.160 bản ghi duy nhất.

Mã: `_spec/tools/biff_dv.py` (chạy: `python3 -I biff_dv.py <out_dir> <root>`). Giải ptg: chỉ các token thường gặp trong DV (Str, Int, Num, Ref, Area, RefN, AreaN, Name, NameX, Ref3d, MemFunc...); token lạ ghi vào `warnings` (hiện 0).
Công thức NAME#n là chỉ số tên trong bảng Name của workbook (chưa ánh xạ sang tên/vùng).

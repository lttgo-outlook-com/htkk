# TT58_BCTC: bộ báo cáo tài chính năm (mã 58_BCTC, phiên bản 3.0.5)

## Thư mục này là gì

Bản sao các biểu mẫu, công thức tính và template của bộ báo cáo tài chính năm mã 58_BCTC, lấy từ phần mềm kê khai thuế HTKK 5.7.9. Thư mục có 21 tệp: 19 tệp là bản sao từ tệp gốc (tệp gốc ở nơi khác trong kho không bị sửa), cộng 2 tệp tạo thêm cho thư mục này là 00_DOC_TRUOC.md (tệp bạn đang đọc) và 02_Cong_thuc/cong_thuc_doc_duoc.csv. Thư mục này chỉ để tra cứu, không phải phần mềm kê khai.

## Căn cứ pháp lý

Theo chú thích trong phần mềm HTKK 5.7.9: Thông tư 58/2026/TT-BTC ngày 25/5/2026. CHƯA XÁC MINH với văn bản gốc.

## Từng thư mục con

| Thư mục | Chứa gì | Mở bằng gì |
|---|---|---|
| 01_Bieu_mau | Biểu mẫu báo cáo ở dạng Excel (2 tệp) | Excel hoặc LibreOffice Calc |
| 02_Cong_thuc | Danh sách công thức tính trong biểu mẫu | Excel (tệp .csv), Notepad hoặc TextEdit (tệp .txt) |
| 03_Template_goc | Template gốc lấy từ bộ cài HTKK (2 tệp Excel và 1 tệp cấu trúc dữ liệu) | Excel (.xls, .xls1); Notepad (.xsd) |
| 04_Ky_thuat | Tài liệu cho lập trình viên | Không cần mở nếu bạn không phải lập trình viên |

## Chi tiết các tệp

### 01_Bieu_mau

- 58_BCTC_v3.0.5_mau.xls: biểu mẫu chính. Các sheet trong tệp: 01, 01-1, KHBS, Header.
- 58_BCTC_v3.0.5_mau.xls1: tệp thứ hai cùng cấu trúc sheet (01, 01-1, KHBS, Header). Phần mở rộng .xls1 là cách đặt tên của HTKK; nếu Excel không mở được, đổi tên bản sao thành .xls rồi mở. Tệp này có 33 ô công thức, khác với tệp .xls (44 ô công thức). Vai trò chính xác của tệp chưa được xác định trong tài liệu của kho.

### 02_Cong_thuc

- cong_thuc_doc_duoc.csv: mỗi dòng là một ô có công thức. Các cột: tệp mẫu, sheet, ô, công thức, ý nghĩa (nhãn chữ trong biểu mẫu), loại công thức. Mở bằng Excel; nếu chữ tiếng Việt bị lỗi, dùng Data > From Text/CSV và chọn bảng mã Unicode.
- cong_thuc.txt: bản công thức gốc đi kèm dữ liệu của HTKK, mỗi ô một dòng: ô | công thức | nhãn | nguồn.
- Cách đọc công thức: "Header!C3" là ô C3 trong sheet Header (nơi nhập thông tin chung như tên đơn vị, mã số thuế, kỳ báo cáo). "R40+R41" là cộng ô R40 với ô R41 trong cùng sheet. "SUM(F5:F7)" là cộng các ô từ F5 đến F7.
- Số dòng của tệp CSV: 77 công thức (44 của tệp .xls và 33 của tệp .xls1), khớp với số công thức trong hai tệp đặc tả công thức ở 04_Ky_thuat/spec (tên bắt đầu bằng 58_BCTC@). Trong 77 dòng có 64 dòng có nhãn; 13 dòng nhãn trống vì tệp cong_thuc.txt không có nhãn cho ô đó.
- Cột "Loại" do chương trình tự phân loại theo hình thức công thức (30 dòng lấy dữ liệu từ sheet Header, 6 dòng cộng dải ô, 14 dòng cộng trừ ô, 27 dòng công thức khác); chỉ mô tả hình thức, không phải diễn giải pháp lý.

### 03_Template_goc

- _8_bctc_305.xls, _8_bctc_305.xls1: template gốc trích từ bộ cài HTKK. Nội dung trùng hoàn toàn (so bằng mã băm sha256) với 58_BCTC_v3.0.5_mau.xls và 58_BCTC_v3.0.5_mau.xls1 trong 01_Bieu_mau; chỉ khác tên. Giữ cả hai bản để mỗi bản nằm đúng nơi dùng.
- _8_bctc_305.xsd: khai báo cấu trúc dữ liệu (tên trường, kiểu dữ liệu) của báo cáo. Nội dung trùng hoàn toàn (sha256) với 04_Ky_thuat/58_BCTC_v3.0.5_xsd.xsd; chỉ khác tên, giữ cả hai vì cùng lý do.

### 04_Ky_thuat (dành cho lập trình viên)

- 58_BCTC_v3.0.5_xsd.xsd: khai báo cấu trúc dữ liệu XML của báo cáo.
- xml_cau_truc/: 4 tệp XML mẫu cấu trúc các bảng (8_b01_01_thtc_305.xml, 8_b01_01_thtc_xml_305.xml, 8_b01_02_sxkd_305.xml, 8_b01_02_sxkd_xml_305.xml).
- spec/: 7 tệp đặc tả (định dạng JSON, tệp văn bản có cấu trúc) do công cụ phân tích của kho tạo ra. Tên tệp bắt đầu bằng nhóm: rules__ (quy tắc), validations__ (kiểm tra hợp lệ), xsd__ (cấu trúc XSD), coverage__ (độ phủ); tệp 58_BCTC@...json là công thức; tệp _forms__...json là danh sách ô có ràng buộc dữ liệu. Bốn tệp có tiền tố được đặt thêm tiền tố vì tên gốc trùng nhau theo hai cặp: rules__ và validations__ (cùng tên gốc BctcTt58Nckt305.json), xsd__ và coverage__ (cùng tên gốc 58_BCTC_3.0.5.json).
- ma_nguon_C#/BctcTt58Nckt305.cs: mã nguồn lớp xử lý của biểu mẫu. Tệp dài 306 dòng, có 13 khối try (khối bắt lỗi của lập trình). 7 khối rỗng, không có lệnh bên trong: ValidatePeriod_ALL, ValidatePeriod_axfpsPhuLuc, AutoBindData_0, AutoBindData_1, CellChange_0, CellChange_1, Validate_1. 6 khối có lệnh: ButtonClick_0 và ButtonClick_1 (khi bấm ô D_17 thì gọi hàm lấy số kỳ trước), isValidDuLieuTK (gọi Validate_0, Validate_1 và đổi màu tab sheet thành hồng, vàng hoặc trắng theo mức cảnh báo), Validate_0 (kiểm tra sheet đầu, hiện thông báo mã 0484 khi sheet thứ hai bị ẩn), getSoKyTruoc (lấy số kỳ trước), getPathDataFileOld (tạo đường dẫn tệp dữ liệu kỳ cũ).

## Lưu ý

- Không sửa trực tiếp các tệp trong thư mục này nếu cần giữ làm bản đối chiếu; làm việc trên bản sao khác.
- Mọi kết luận về cách áp dụng cho một hồ sơ cụ thể cần người có chuyên môn xác nhận.

# 05_Cau_hinh_trich_tu_HTKK: các đoạn cấu hình TT58 trích từ tệp dùng chung

Phần mềm HTKK 5.7.9 gom cấu hình của rất nhiều tờ khai vào các tệp chung. Thư mục này giữ riêng các dòng nói về bộ báo cáo tài chính TT58 (mã 58_BCTC, mục menu X4, mã tờ khai 1280). Mỗi tệp là bản chép nguyên văn từ tệp nguồn, không diễn giải, không sửa. Dòng đầu mỗi đoạn (bắt đầu bằng "@@@ nguồn:") ghi đường dẫn nguồn tính từ gốc kho HTKK và số dòng. Đuôi .txt được thêm vào để mở bằng Notepad hoặc TextEdit; nội dung bên dưới dòng "@@@" là nội dung gốc.

Giải thích chữ dùng ở đây:

- JSON là định dạng văn bản liệt kê dữ liệu theo cặp "tên": giá trị. Các tệp JSON trong thư mục _spec của kho là bản chuyển đổi do công cụ phân tích của dự án tạo ra từ tệp gốc, không phải tệp của HTKK. Ở tệp 18, nguồn là tệp JSON nén chỉ có một dòng, nên thay cho số dòng, dòng đầu ghi đường dẫn khóa JSON (vị trí của đối tượng trong cây dữ liệu).
- "Đoạn" là các dòng liền nhau có liên quan đến TT58, không nhất thiết là một khối XML hay JSON đóng kín (có đủ thẻ mở và thẻ đóng, hoặc đủ dấu `{` và `}`). Ví dụ tệp 01 dừng ở dòng 1113 và không có thẻ đóng `</Menu>` ở dòng 1114; tệp 03, 04, 20 thì có đủ một đối tượng JSON.
- Ba tên trường trong mục menu X4 (tệp 01 và 03):
  - PopID="101": chưa rõ nghĩa, đây là tên trường trong HTKK. Giá trị 101 trùng mã của mục menu gốc "Kê khai" (_extracted/menu.xml dòng 11).
  - ParentID="101_6": mã của mục menu cha, tức nhóm chứa mục X4 trong cây menu. 101_6 là nhóm "Báo cáo tài chính" (_extracted/menu.xml dòng 762), nằm dưới "Kê khai". Mã giải mã frmHTKKSystems.cs dòng 181 dùng ParentID khi thêm từng mục vào cây menu.
  - RowTTDLT="13_17": một khoảng dòng viết dạng "dòng đầu_dòng cuối". Theo mã giải mã SpreadPrint2.cs dòng 318-336: khi in, nếu mã số thuế đại lý thuế để trống, HTKK ẩn các dòng từ 13 đến 17 của sheet đó trên bản in; giá trị "_" (ở sheet KHBS) nghĩa là không ẩn dòng nào. Chữ viết tắt TTDLT có thể là "thông tin đại lý thuế", nhưng chưa có tài liệu HTKK nào xác nhận.

| Tệp | Nội dung | Ý nghĩa đơn giản |
|---|---|---|
| 01 | _extracted/menu.xml dòng 1108-1113 | Mục menu X4: tên tờ khai, thư mục 58_BCTC, hai trang tính có dữ liệu (BCTHTC và BCKQHĐKD) và trang KHBS không dùng (Active=0). Đoạn dừng ở thẻ `</Validity>` dòng 1113; thẻ đóng `</Menu>` của mục X4 nằm ở dòng 1114 và không có trong đoạn. Bản đầy đủ của cùng mục ở dạng JSON là tệp 03 |
| 02 | _extracted/map.xml dòng 1342-1344 | Khai báo kỳ kê khai theo năm và ba tên tệp dữ liệu: 58_B01_01_THTC, 58_B01_02_SXKD, KHBS_58_B01_01_THTC |
| 03 | _spec/config/menu.xml.json dòng 10533-10598 | Toàn bộ đối tượng menu X4 ở dạng JSON, từ dấu `{` mở (dòng 10533) đến dấu `}` đóng tương ứng (dòng 10598): ID X4, PopID 101, ParentID 101_6, tên tờ khai, MaFile và XsdSchema 58_BCTC, MaToKhai 1280, tên xử lý (Class) BctcTt58, phần Validity với FolderName 58_BCTC và ba sheet 01 (BCTHTC), 01-1 (BCKQHĐKD), KHBS (Active=0). Tương ứng tệp 01 nhưng có cả dấu đóng của đối tượng |
| 04 | _spec/config/map.xml.json dòng 9271-9294 | Toàn bộ đối tượng map X4 ở dạng JSON, tương ứng tệp 02 |
| 05 | _forms/_manifest.csv dòng 1387-1393 | Bảng liệt kê 7 tệp của TT58 trong bộ cài, kèm mã băm sha256 (chuỗi 64 ký tự tính từ nội dung tệp, như dấu vân tay) |
| 06 | _re/moved_files.csv dòng 601-604 | Ghi nhận bốn tệp XML được chuyển từ thư mục 8_B01_01 và 8_B01_02 sang thư mục 58_BCTC |
| 07 | _spec/coverage/summary.csv dòng 1 và 221 | Dòng tiêu đề cột và dòng thống kê của 58_BCTC (số trang tính, số ô, số ô tính tự động, số ô nhập tay, v.v.) |
| 08 | _re/NOTES.md dòng 252 và 370 | Hai dòng ghi chú của người làm dự án nhắc đến TT58 |
| 09 | _spec/coverage/sheet_log.json dòng 4856-4866 | Ghi nhận không tìm thấy tệp dữ liệu cho trang KHBS của 58_BCTC |
| 10 | _spec/csharp/classes.json dòng 504-513 | Khai báo lớp xử lý (đoạn chương trình điều khiển biểu mẫu) BctcTt58Nckt305 gắn với menu X4 |
| 11 | _spec/csharp/index.json dòng 28972-29141 | Danh sách 13 hàm trong tệp mã nguồn BctcTt58Nckt305.cs: tên, dòng bắt đầu, dòng kết thúc |
| 12 | _spec/formulas/index.json dòng 15999-16041 | Mục lục công thức: 44 công thức ở tệp .xls, 33 ở tệp .xls1 |
| 13 | _spec/rules/index.json dòng 731-741 | Mục lục quy tắc xử lý của lớp BctcTt58Nckt305: 6 hàm, 6 bản ghi |
| 14 | _spec/validations/index.json dòng 975-992 | Mục lục kiểm tra hợp lệ của lớp BctcTt58Nckt305 |
| 15 | _spec/xsd/index.json dòng 3885-3901 | Mục lục tệp xsd (tệp khai báo cấu trúc dữ liệu hợp lệ) của 58_BCTC: 187 mục |
| 16 | _spec/dv/index.json dòng 2598-2606 | Mục lục ràng buộc nhập liệu: tệp .xls1 có 1 quy tắc kiểu số thập phân |
| 17 | _spec/config/index.json dòng 2146-2149 | Một tham chiếu gián tiếp từ cấu hình đến dòng 287 của tệp BctcTt58Nckt305.cs |
| 18 | _spec/catalog/map.xml.json (khóa JSON) | Hai đối tượng của X4 trong danh mục map, tương ứng tệp 02 |
| 19 | _extracted/button.xml dòng 278 | Cấu hình nút bấm trên màn hình kê khai của mục X4. Giá trị "1" là hiện nút, "0" là ẩn nút: hiện Thêm phụ lục, Xóa, Nhập lại, Ghi, In, Kết xuất, Nhập từ XML, Đóng; ẩn Tải tờ khai và Kiểm tra |
| 20 | _spec/config/button.xml.json dòng 4026-4041 | Cùng dòng nút bấm của tệp 19 ở dạng JSON, đủ một đối tượng từ `{` đến `}` |
| 21 | _spec/app/by_menu.json dòng 41148-41219 | Chỉ mục do công cụ phân tích của dự án tạo ra: 8 điểm trong mã nguồn chung của HTKK (in, kiểm tra, xuất XML, nạp dữ liệu, hiển thị) có điều kiện nhắc đến mã menu X4; mỗi điểm ghi tệp mã nguồn, số dòng và tên hàm |
| 22 | _decompiled/tax.htkk.utilities/TAX.HTKK.Utilities.Core/HTKKConstant.cs dòng 763, 767 và 1209 | Ba hằng số (giá trị đặt tên cố định trong mã) có danh sách mã menu chứa X4. Tên hằng số cho biết mục đích: ID_BCTC_PL_LCTTTT_LCTTGT (nhóm bộ BCTC có phụ lục lưu chuyển tiền tệ trực tiếp và gián tiếp), LIST_ID_TK_SU_DUNG_NGAY_BD_NAM_TC (tờ khai dùng ngày bắt đầu năm tài chính), DS_ID_TO_KHAI_NO_PRINT_BARCODE (tờ khai không in mã vạch). Đã so bằng script với 8 điểm của tệp 21: chuỗi của DS_ID_TO_KHAI_NO_PRINT_BARCODE có nguyên văn ở SpreadPrint2.cs:200; của ID_BCTC_PL_LCTTTT_LCTTGT ở WinDoWorkEvent.cs:2426, WinSpreadProcess.cs:2502 và 3817; của LIST_ID_TK_SU_DUNG_NGAY_BD_NAM_TC ở frmHTKKPeriod.cs:3492 và 11607; hai điểm WinDoWorkEvent.cs:8339 và 10015 dùng danh sách khác, không trùng hằng số nào. Tên hằng số do HTKK đặt; phần giải thích trong ngoặc là cách đọc tên, chưa đối chiếu tài liệu HTKK |

Lý do lấy tệp 19, 20, 21 và 22. Bốn đoạn này không chứa chữ 1280, 58_BCTC hay TT58. Chúng được nhận vì chính mã menu X4 là khóa của đoạn: button.xml đặt nút bấm theo ID menu; by_menu.json là chỉ mục theo mã menu, ghi "in_menu_xml": true (mã X4 lấy từ menu.xml); ba hằng số trong HTKKConstant.cs là danh sách mã menu. Trong _extracted/menu.xml chỉ có một mục ID="X4", và mục đó là TT58 (MaToKhai 1280, FolderName 58_BCTC).

Quan hệ giữa tệp 21 và _spec/app/app_branches.json. Hai tệp do cùng công cụ phân tích của dự án sinh ra và mô tả cùng 8 điểm rẽ nhánh. app_branches.json liệt kê theo từng điểm: 24 dòng có X4, thuộc đúng 8 điểm (tệp mã nguồn và số dòng) trùng 8 điểm của tệp 21 (đã so bằng script). Mỗi điểm là một điều kiện dùng chung cho nhiều mã menu, X4 chỉ là một mã trong danh sách. by_menu.json gom các điểm đó theo mã menu, và phần của X4 nằm liền trong một đối tượng. Vì vậy thư mục này chỉ giữ tệp 21 làm bản đại diện; trong CSV, các dòng của app_branches.json, _spec/rules/_app.json và bốn tệp mã giải mã trong _decompiled/htkk/HTKK có quyết định "ĐÃ CÓ (cùng điểm với tệp 21)".

## Cách đọc tệp 03, 04, 20, 21 bằng chương trình JSON

Dòng cuối của bốn tệp này là `},`. Dấu phẩy là dấu ngăn cách giữa các phần tử của danh sách lớn hơn trong tệp nguồn; nó được giữ nguyên vì đoạn trích chép nguyên văn cả dòng. Muốn đọc bằng chương trình JSON: bỏ dòng "@@@" và bỏ dấu phẩy cuối. Riêng tệp 21 là một cặp `"X4": {...}`, không phải một đối tượng đứng riêng, nên phải bọc thêm một cặp `{` `}` bên ngoài. Làm như vậy thì cả bốn tệp đều đọc được thành JSON hợp lệ (đã kiểm bằng script).

## Danh sách kết quả tìm theo khóa: 00_KET_QUA_TIM_THEO_KHOA.csv

Tệp CSV này liệt kê mọi tệp trong _extracted/, _spec/, _re/, _forms/_manifest.csv và _decompiled/ có chứa một trong các khóa định danh của TT58, kèm quyết định lấy hay loại và lý do. Mở bằng Excel qua Data > From Text/CSV, chọn bảng mã "65001: Unicode (UTF-8)".

Các khóa đã tìm (không phân biệt dạng thuộc tính XML hay JSON, không phụ thuộc khoảng trắng):

- X4_token: chữ X4 đứng riêng, không dính chữ hoặc số khác. Khóa này bao gồm `ID="X4"`, `"ID": "X4"`, `"ID":"X4"` và X4 không có dấu nháy (trong CSV, trong danh sách mã).
- num_1280: số 1280 đứng riêng. Khóa này bao gồm `MaToKhai="1280"` và `"MaToKhai": "1280"`.
- FolderName_58_BCTC: thuộc tính FolderName có giá trị 58_BCTC.
- BCTHTC và BCKQHDKD (chữ BCKQHĐKD, tìm cả hai dạng mã hóa Unicode NFC và NFD): tên hai sheet của X4.
- KHBS_58_B01_01_THTC: tên tệp dữ liệu KHBS trong khai báo map của X4.
- BctcTt58: chuỗi BctcTt58, không phân biệt chữ hoa và chữ thường (tên xử lý của X4 và tên lớp BctcTt58Nckt305).
- TT58: chuỗi TT58, không phân biệt chữ hoa và chữ thường, không có chữ cái đứng ngay trước và không có chữ số đứng ngay sau. Vì có điều kiện "không có chữ cái đứng trước", chữ Tt58 nằm trong BctcTt58 không tính vào khóa này (đã tính ở khóa BctcTt58). "TT58/2026" có tính.
- HelpContextID_X41_X42: mã trợ giúp dạng X41 kèm 1 đến 2 chữ số, hoặc X42 kèm 1 chữ số (ví dụ X412, X4110, X426), là giá trị HelpContextID của các ô trong hai tệp cấu trúc 8_b01_01_thtc_305.xml và 8_b01_02_sxkd_305.xml.

Các cột: khoa_tim, tep_nguon, so_lan_khop, cac_dong (tối đa 20 số dòng đầu), ngu_canh_dong_dau (đoạn chữ quanh lần khớp đầu tiên), quyet_dinh, noi_luu_trong_TT58_BCTC, ly_do. Mỗi dòng của CSV là một cặp khóa và tệp, nên một tệp có thể xuất hiện ở nhiều dòng. Tệp nguồn mã hóa UTF-16 (ví dụ các tệp XML cấu trúc) được giải mã đúng trước khi trích ngữ cảnh; CSV không chứa byte rỗng (byte 0) nào.

Kết quả: 174 dòng, 134 tệp khác nhau. Quyết định:

- 1 dòng LẤY MỚI: HTKKConstant.cs, thành tệp 22.
- 35 dòng ĐÃ CÓ: mọi dòng khớp nằm trong khoảng dòng của một đoạn trích đã có (kiểm bằng script), gồm cả tệp 19, 20, 21.
- 9 dòng ĐÃ CÓ (toàn tệp): tệp riêng của TT58 đã có bản sao toàn tệp trong thư mục TT58_BCTC, sha256 trùng.
- 6 dòng ĐÃ CÓ (cùng điểm với tệp 21): app_branches.json, _spec/rules/_app.json và bốn tệp mã giải mã trong _decompiled/htkk/HTKK (xem đoạn "Quan hệ giữa tệp 21 và app_branches.json").
- 4 dòng ĐÃ CÓ MỘT PHẦN: tên sheet BCTHTC và BCKQHĐKD trong menu.xml và menu.xml.json; dòng thuộc mục X4 đã có, các dòng còn lại thuộc mục BCTC khác.
- 2 dòng ĐÃ CÓ (trùng tệp 08): bản sao lưu _re/NOTES.md.bak-20261008-153710, hai dòng khớp giống hệt hai dòng của NOTES.md ở tệp 08.
- 117 dòng LOẠI. Lý do loại chính: X4 là địa chỉ ô Excel (cột X, dòng 4) của tờ khai khác, kể cả trong công thức do lớp Tndn03 gán; X4 nằm giữa dữ liệu base64 (dạng chữ của dữ liệu nhị phân) trong tệp tài nguyên .resx; 1280 là mã kho bạc, kích thước màn hình 1280x1024, số dòng trong mã nguồn hoặc tệp xsd của tờ khai khác, hoặc chuỗi con của một mã sha256; mã trợ giúp X41x xuất hiện ở tệp cấu trúc của 41_BCTC (menu X3) và 99_BCTC_GND_TL (menu X2); chuỗi dạng X410 trở lên trong _spec/formulas là địa chỉ ô.

Kết quả riêng cho hai khóa mới. BctcTt58: 12 dòng, không dòng nào bị loại (8 ĐÃ CÓ, 3 ĐÃ CÓ toàn tệp, 1 trùng tệp 08); trong _decompiled, chuỗi này chỉ có ở chính tệp BctcTt58Nckt305.cs (dòng 15, khai báo lớp). TT58: 7 dòng, không dòng nào bị loại (6 ĐÃ CÓ, 1 trùng tệp 08); _decompiled không có chuỗi TT58. Hai khóa này không tìm ra tệp mới nào ngoài các đoạn đã có.

Phạm vi chưa quét. Lượt tìm đọc 7641 tệp văn bản và bỏ qua 1235 tệp: 1188 tệp bỏ theo đuôi (.xls 569, .xls1 420, .png 64, .bmp 40, .gif 27, .dll 25, .jpg 17, .ico 9, .xlsx 6, .pdb 5, .exe 4, .chm 1, .pyc 1) và 47 tệp nhị phân không đọc được thành văn bản. Không có tệp .msi trong năm nguồn này. Chưa quét X4 trong tệp nhị phân (.dll, .pdb, .exe) vì X4 là mã nội bộ chỉ hai ký tự, khớp trong nhị phân không cho biết thuộc mục nào; phần mã của bốn tệp .dll và .exe chính đã được quét qua bản giải mã trong _decompiled. Chưa quét trong tệp .xls và .xls1 vì ở đó X4 là địa chỉ ô; các mẫu .xls và .xls1 của TT58 đã có bản sao toàn tệp trong 01_Bieu_mau và 03_Template_goc. Tệp trợ giúp htkk.chm là tệp nén nên được giải nén và tìm riêng; kết quả ghi ở mục "Những điều đã kiểm và chưa kiểm được" của 00_DOC_TRUOC.md.

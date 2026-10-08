# HTKK 5.7.9 — Map Controltype → widget/hành vi (bản đầy đủ)

> Reference cho engine web (kèm theo `_re/DOC_ENGINE.md` §7). Source: 3 switch render trong `htkk.exe` + helpers, tất cả đã verify 07/10/2026 bằng 6 agent điều tra song song. Đường dẫn rút gọn: `WSP` = `_decompiled/htkk/HTKK/WinSpreadProcess.cs`, `WDE` = `WinDoWorkEvent.cs`, `FTK` = `frmWinToKhai.cs`, `FOF` = `frmWinHTKKOpenFile.cs`, `GF` = `_decompiled/tax.htkk.business/GeneralFunctions.cs`.

## 3 dispatch render chính

1. **`fpSpFormatCellTemplate`** — ô **static** (ngoài bảng): `WSP:4186`, switch cases `4211–4396`, fallback `DC_` → combo / `DC_*_M_*` → checkbox-combo tại `4410–4420`, còn lại → `DefaultCellFomat` (`5370`).
2. **`fpSpFormatTableCellTemplate`** — ô **mẫu trong bảng**: `WSP:4432`, cases `4438–4607`, fallback DC_ `4609–4616`, còn lại → `DefaultTableCellFomat` (`5298–5330`: XLS format numeric → NumberCellType, có DateTimeFormat → date, khác → TextCellType).
3. **`fpSpFormatTableCell`** — render/bind **hàng dữ liệu lớn** (bảng lớn): `WSP:4819`, block date `4690–4742`, cases combo `4839–5011`.

Helpers bổ sung:
- **`FormatCelltypeFromSinfo`** (mask nhập): `GF:10187` — `"14"→"00/00/0000"`, `"96"→"00:00:00"`, `"94"→"00/00/0000 00:00:00"`, `"16"→number/currency`, `"2"→CheckBoxCellType`.
- **`GetValueCustom`** (đọc giá trị khi compute/export): `GF:258` — `2→"True"/"False"`, `14→dd/mm/yyyy` (từ 8 hoặc 10 ký tự), `96→hh:mm:ss`, `94→dd/mm/yyyy hh:mm:ss`, `0→text`, `16→raw number`, default: number nếu cell có numeric format, button → `"TaiBangKe"+caption`.
- **`BindDataLarge`** (hiển thị bulk): `WSP:2049–2078` — `2→"x"`, `14→toDateDDMMYYYY`, `96→toTimeHHMMSS`, `94→toDateAndTimeDDMMYYYYHHMMSS`, `19→toDateMMYYYY`.
- Nhóm date gia nhập key input: `FTK:517` — `"14"|"19"|"104"|"109"|"96"|"94"` → `CompareKeycode(e,1)`.
- Mapping giá trị hiển thị → mã `SelectedValue`/`indexColumnSelectedValue`: static `FTK:12828–13346`, table `13358–14530`.
- Validate mã–tên `isValidMa`: `FOF:1071–1230`.
- Export XML: `WDE:10444–10452` (rỗng: `16→"0"`, `2→"0"/"1"`, `26` 13 ký tự → chèn `-` tại vị trí 10); `setNilAttributeByControlType` `10463–10491`: các type **14, 26, 33, 96, 94** thêm `xsi:nil="true"` khi rỗng.
- Danh sách ô "free value" (không có mã SelectedValue): `GF:9856` — `{0, 15, 1, 2, 14, 96, 94, 16, 17, 26, 27, 52}`.
- `ConvertHelpers.cs:128–158` (utilities): `16→cell.Value`, `14→cell.Text`; import Excel: `16` (strip dấu), `14`, `2` (x/1/true→True).

## Bảng hoàn chỉnh

### Nhóm widget cơ bản

| Type | Nghĩa / widget | Evidence |
|---|---|---|
| 0 | Text thường (Path chứa `LYDOKHAC` → InsertTextCellKHBS) | `WSP:4396–4408`; free-value `GF:9856` |
| 1 | **Default text** — không có case trong bất kỳ switch nào | chỉ xuất hiện trong danh sách free-value `GF:9856`; template `_1_ntnn_tt80_251.xml` (Path=BM31) |
| 2 | Checkbox (hiển thị "x"; export "True"/"False"/"0"/"1") | `WSP:4223, 4571, 5060 (InsertCheckbox)`; `WSP:2057–2067`; `WDE:10448` |
| 15 | Nút (button) — skip khi export | `WSP:4247, 4598`; `InsertButton` `5042`; `WDE:2580` |
| 16 | Số (thousand separator tuỳ DefaultValue có dấu ","; rỗng→"0") | `WSP:4241, 4589 (InsertNumberCell)`; `FOF:717–721`; `WDE:10444` |
| 17 | Text (CellTextFomat, word-wrap) | `WSP:4601–4602`; static fall-through case "26"→"0" `4394–4396` |
| 26 | Text dài; export 13 ký tự → chèn `-` (định dạng mã 10+3) | `WSP:4601 (CellTextFomat)`; `WDE:10452–10454` |
| 27 | Số (InsertNumberCell) | `WSP:4595`; free-value `9856` |
| 33 | Số "quy/năm" (QYYYY), không separator | `WSP:4244, 4592 → InsertQYear` `5587–5612`; export nil `10472` |
| 52 | Free value (không có case; không xuất hiện trong template đã đo) | `GF:9856` |

### Nhóm ngày/giờ

| Type | Nghĩa | Evidence |
|---|---|---|
| 14 | Ngày dd/mm/yyyy (mask 00/00/0000) | `WSP:4574 InsertDateCell`; mask `GF:10224`; key `FTK:517` |
| 19 | Tháng/năm mm/yyyy | `WSP:4583 InsertMonthYearCell`; `WSP:2074–2076 (toDateMMYYYY)`; key `FTK:517` |
| 96 | Giờ hh:mm:ss | `WSP:4577 InsertTimeCell`; mask `GF:10251`; `WSP:2072` |
| 94 | Ngày+giờ dd/mm/yyyy hh:mm:ss | `WSP:4580 InsertDateTimeCell`; mask `GF:10257`; `WSP:2090` |
| 104 | dd/mm | `WSP:4605 InsertDayMonthCell`; key `FTK:517` |
| 109 | yyyy | `WSP:4586 InsertYearCell`; key `FTK:517` |
| 18 | Ngày đầu quý `Catalogue_FirstDayQuarter.xml` (key) — chỉ có trong code, không có trong template đã đo | `WSP:4229` |

### Nhóm lookup combo (ô hiển thị tên, mã lưu vào cột `SelectedValue`/`indexColumnSelectedValue`)

Cột evidence = case trong switch render; cột "mapping" = dòng giá trị→mã trong `FTK`.

| Type | Catalog (DM) | Render (WSP) | Mapping (FTK) |
|---|---|---|---|
| 3 | Quan hệ (QuaHe) | 4468 | 13358 (table: `QuanHe[ten]→ma`) |
| 4 | Quốc tịch (VN/Khác) — static: `GetDMQuocGiaVietNamKhac`; table: dict `QuocTich`; giá trị "02" (nước ngoài) → trigger combo tỉnh/huyện/xã "99999" | 4296 / 4447 | 13074 (static); 13363–13386 (table) |
| 5 | Tỉnh (cũ) `dicTinhTP` + `ChildCell` = combo huyện | 4250 | 13228–13243 (`WinInsertCombox` huyện `dicQuanHuyenByTinhTP`) |
| 6 | Huyện `getMaQuanHuyen` + `ChildCell` = combo phường/xã | 4266 | 13292–13303 |
| 7 | Phường/xã (cũ) `getMaPhuongXa` | 4269 | 13305 |
| 8 | CQT cấp Cục (static: `getAllTextCqtCucMoi` + `ChildCell`=CQT quản lý; table: combo động `GetDMFromCombobox`) | 4214 | 12828–12843 (static); 13523–13543 (table) |
| 9 | CQT chi cục (CQT quản lý, `ParentCell`=cục) | 4217 | 12845–12865 (static); 13545–13561 (table) |
| 10 | Lý do gia hạn (static: `WinInsertComboxKey` Catalogue_DM_Lydo_GH_TT16; table: `LyDoGiaHan`) | 4235 / 4471 | 12905–12914 |
| 11 | Ngành nghề có doanh thu cao nhất | 4211 | 12872–12881 |
| 20 | Khoáng sản (mã tách "_" lấy phần [2]; `ChildValue`=đơn vị) | 4486 | 13563–13593 |
| 21 | Đơn vị (khoáng sản / bảng lớn: `GetDMDonVi`) | 4525 / 4930 | — (ô con) |
| 22 | Phí (không lệ phí) `DMPhiLephi` + `ChildValue`=tiêu mục | 4532 | 13896–13915 |
| 23 | Nguồn vốn `DMTaiNguyen` | 4220 / 4522 | 12894 / 13880 |
| 25 | Loại tem rượu `DMLoaiTemRuou` (index=mã, `ChildValue`=tên) | 4238 / 4900 | 13960–13970 |
| 28 | Ghi chú tem rượu `DMGhiChuTemRuou` | 4903 | 13972–13987 |
| 29 | Lý do tạm nghỉ (Catalogue_DM_Lydo_GH_TT16, key) | 4226 | 12867–12870 |
| 32 | Loại hóa đơn `GetDMHoaDon` | 4936 | — (mã qua indexColumn) |
| 34 | Hình thức hóa đơn `GetHTHoaDon` | 4942 | — |
| 35 | Thuế BVMT `GetDMThueBVMT` | 4945 | 14125 |
| 36 | Than `GetDMThan` | 4308 | 14142 |
| 37 | Ngành nghề KD `GetDMNganhNgheKD` | 4281 | 12942 / 14352 |
| 38 | Nhóm ngành nghề KD `GetDMNhomNganhNgheKD` | 4284 | — |
| 39 | Địa bạ thuế tiêu thụ `GetDMThueTieuThuDB` | 4960 | 14424 (if rỗng — không có hành vi thêm) |
| 40 | Biên lai phí `GetDMBienLaiPhi` | 4966 | — |
| 41 | Đơn vị TTDB `GetDMTTDBDonVi` | 4933 | — |
| 42 | Dịch vụ thuế TTDB `GetDichVuThueTieuThuDB` + **unique check** (nhắc lỗi "0261", xóa ô) | 4963 | 14514–14540 |
| 43 | Loại thông tin thay đổi `GetDMLoaiThongTinThayDoi` + **unique check** (lỗi "0445", xóa ô + cột AA) | 4990 | 14427–14449 |
| 44 | Mục dịch vụ `GetDMMucDich` | 4984 | validate `FOF:1086` |
| 45 | Trạng thái `GetDMTrangThai` | 4987 | validate `1094` |
| 46 | Hợp đồng `GetDMHopDong` | 4969 | — |
| 47 | Ghi chú `GetDMGhiChu` | 5002 | — |
| 48 | Loại hoàn thuế `LoaiHoanThue` | 4365 | 13195 |
| 49 | Lý do hoàn thuế `LyDoHoanThue` | 4389 | 13184 |
| 50 | Khoản nộp NS hoàn thuế `KhoanNopNSSHoanThue` (sort theo key) | 4568 | 14244 |
| 51 | Kỳ phát sinh hoàn thuế `KyPhatSinhHoanThue` | 4541 | 14262 |
| 53 | Ký kiểm toán `GetDMYKienKiemToan` | 4287 / 4492 | 12997 / 13617 |
| 54 | Giới tính `GioiTinh` | 4278 | 12931 / 14176 |
| 55 | Loại giấy tờ tự thân `LoaiGiayToTuyThan` | 4477 | 12953 / 14193 |
| 56 | Quốc gia `GetDMQuocGia` (table 4: `QuocTich` + ChildValue=provincia) | 4290 / 4495 | 13063 / 13388–13407; validate `1150` |
| 57 | PPLK `GetDMPPLK` | 4347 / 4501 | 13096 / 13784 |
| 58 | Loại TT thay đổi (bảng liệt kê) `GetDMLoaiThongTinThayDoiBLP` + **unique check** ("0445") | 4993 | 14455–14477 |
| 59 | Mục lệ phí môn bài CT22 `GetDMMucLePhiMonBaiCT22` | 4232 | 12916 / 14480 |
| 60 | Mục lệ phí môn bài CT23 `GetDMMucLePhiMonBaiCT23` | 4999 | 14497 |
| 61 | Nơi cấp `GetDMNoiCap` | 4299 / 4483 | 12975 / 14227 |
| 62 | Khoáng sản BVMT `GetDMKSBVMT` + đơn vị | 4489 / 4927 | 13595 |
| 64 | Lĩnh vực hoạt động `LinhVucHD` | 4480 | 12964 / 14210 |
| 65 | PPLK ND132 `GetDMPPLKND132` | 4350 / 4504 | 13107 / 13800 |
| 66 | Tiêu mục `GetDMTM` | 4353 / 4507 | 13118 / 13816 |
| 67 | Quốc tịch full `GetDMQuocTichFull` | 4293 / 4498 | 13085 / 13768; validate `1212` |
| 68 | Hợp đồng (1) `GetDMHopDong1` | 4972 | validate `FOF:1078` |
| 69 | Phường/xã theo CQT `GetPhuongXaTheoCQT` | 4975 | — |
| 70 | Quốc tịch/làng giềng `GetDMCacNuocLangGieng` | 4302 / 4450 | 12986 |
| 71 | Xăng dầu `GetDMXangDau` | 4311 | 14159 |
| 72 | Tháng `DMMo` | 4314 / 4978 | 13650 |
| 73 | Loại tiền tệ `DMLoaiTienTe` | 4906 | 13989; validate `1138` |
| 74 | Loại địa điểm `GetDMLoaiDiaDiem` | 4332 / 4453 | 13019 / 14334 |
| 75 | Miễn lệ phí môn bài `GetDMMienPhiMonBai` | 4335 / 4456 | 13030 / 14298 |
| 76 | Loại hóa đơn doanh thu `GetDMHoaDonDT` | 4939 | validate `1126` |
| 77 | Tờ khai phụ lục KHBS `GetDMToKhaiPhuLucKHBS` | 4341 / 4462 | 13041 / 14370 |
| 78 | Tiêu mục KHBS `GetDMTieuMucKHBS` | 4344 / 4465 | 13052 / 14388 |
| 79 | Mã chi tiết KHBS `GetDMMaChiTieuKHBS` | 4338 / 4459 | 14316 |
| 80 | Thuế suất `DMThueSuat` | 4305 / 4909 | 13633 / 14006 |
| 81 | Loại đất `DMLoaiDat` | 4320 / 4912 | 13667 / 14040 |
| 82 | Hạng đất `DMHangDat` | 4323 / 4915 | 13684 / 14057 |
| 83 | Hàng hóa GTGT `GetDMHangHoaGTGT` | 4356 / 4510 | 13129 / 13832 |
| 84 | Thuế suất GTGT (form 71) `GetDMThueSuatGTGT("71")` | 4359 / 4513 | 13140 / 13848 |
| 85 | Thuế suất GTGT (form 01) `GetDMThueSuatGTGT("01")` | 4362 / 4516 | 13151 / 13864 |
| 86 | PHLPNG `DMPHLPNG` (mã + ChildValue=tiêu mục) | 4317 / 4538 | 13938 |
| 87 | NN cá thể `DMNnkdNtnn` | 4326 / 4918 | 13701 / 14074; validate `1118` |
| 88 | Lệ phí TT63 `DMPhiLephiTt63` + tiêu mục | 4535 | 13917 |
| 89 | Loại tiền đăng ký `DMLoaiTienDk` | 4329 | 13162 / 13718 |
| 90 | Mã tiêu mục `GetMaTieuMuc` | 4544 | 14280 |
| 91 | Ngành có thu nhập cấp `GetDMNganhCoThuNhapCap` | 4368 | 13206 |
| 92 | Chương `GetDMChuong` | 4371 | 13217 |
| 93 | Ngân hàng `GetDMNganHang` | 4374 / 4553 | load 4236–4256; validate `1162` |
| 98 | Ngân hàng KBNN (combo **editable** `WinInsertComboxEdit`) | 4386 / 4565 | 12883; load 4336; validate `1174` |
| 100 | Tỉnh (mới) `dicTinhTPMoi` + ChildCell=phường/xã mới; menu N2/Q4: trigger combo `DC 107` (DM_Kho_Bac) | 4263 | 13245–13261; table 13442–13460 |
| 101 | Phường/xã (mới) `getMaPhuongXaMoi` + set CQT cục theo mã xã | 4272 | 13313–13332 (static); 13499 (table) |
| 102 | CQT `coQuanThus` (theo tỉnh, MaCQT→Ma) | 4275 | 13334 (static); 13511 (table) |
| 103 | Thuế tỉnh `dicThueTinh` + ChildCell=CQT (menu Y9) | 4260 | 13262–13290 |
| 105 | Thuế suất hoàn thuế `ThueSuatHoanThue` | 4392 / 4519 | — |
| 106 | Ngân hàng (mới) `DM_NganHang_Moi.xml` (combo editable, **auto-combobox khi gõ Enter**) | 4377 / 4556 | load 4261–4284; `FTK:545–548` (Enter→`autoCombobox`); validate `1186` |
| 107 | Kho bạc `DM_Kho_Bac.xml` | 4380 / 4559 | load 4286–4309 |
| 108 | Cơ quan hành chính `DM_CQHQ.xml` | 4383 / 4562 | load 4311–4334; validate `1198` |

### Nhóm DC_ (catalog dùng chung) — 327 cell trong template, 66 mã khác nhau

- `DC_N` → combo từ file ngoài **`DM_DUNG_CHUNG.xml`** (thẻ `<Map Controltype="DC_N">`, item `mã###tên`): `GetDMDungChung` `WSP:7848–7889`; render fallback `4410–4420` (static) / `4609–4616` (table); mapping mã `FTK:13173, 13734, 14107, 14406`.
- `DC_*_M_*` → **checkbox-combo** (combo nhiều lựa chọn, mỗi item là checkbox): `WinInsertComboxCheckboxEdit` `5715`, widget `HTKK.Class/CheckedComboBoxCellType.cs:110,191`; render `4416–4419 / 4612–4615`. Template có đúng 1 cell `DC_M_2`.
- Riêng `DC_08`/`DC_13` ở menu N8: đọc từ combo động của sheet thay vì file: `FTK:4364`; `DC_128`/`DC_129` dùng trong Ttdb01: `tax.htkk.business/Ttdb01Nckt207.cs:1077–1078`.
- File catalog đã extract sẵn ở `_extracted/`: `dm_dung_chung.xml`, `dm_nganhang.xml`, `dm_kho_bac.xml`, `dm_nganhang_moi.xml`, `dm_cqhq.xml`.

### Header động (fix template)

`frmFormStart.cs` gán lại Controltype khi sửa HEADER.xml: `tenTinhNNT→"5"` + `ChildCell` (793), `tenHuyenNNT→"6"` + `ParentCell` (801), `tenTinhDLyThue→"5"` (812), `tenHuyenDLyThue→"6"` (820), `maTinhNNT (mới)→"100"` (952, 1004), `maXaNNT (mới)→"101"` (1035, 1047).

## Default (không có xử lý đặc biệt — render text thường)

Đã grep toàn bộ `_decompiled/` với các pattern: `Controltype ==`, `Controltype !=`, `x/y.Controltype ==`, `controlType.Equals(`, `controlType.IndexOf(`, switch trên controltype (2 switch render static `4211–4396`, table-template `4438–4607`, table-data `4839–5011`), `FormatCelltypeFromSinfo`, `GetValueCustom`, `isValidMa`. Kết quả:

- **Type 1** — default text (2 cell trong template; chỉ được liệt kê trong danh sách free-value `GF:9856`).
- **Type 52** — default (free-value theo `9856`; không xuất hiện trong template đã đo).
- Mọi type còn lại trong tần suất quan sát đều có case hoặc mapping trong code (bảng trên). Type **không tồn tại** ở cả code lẫn template: 13, 24, 30, 31, 63, 95, 97, 99.

Lưu ý cho engine web: "default" trong code = `DefaultCellFomat`/`DefaultTableCellFomat` (`5370`, `5298`) — nó vẫn chuyển cell có XLS-format numeric thành NumberCellType và có DateTimeFormat thành date; với web app nên render text + validate theo attribute `MinValue/MaxValue/MaxLen` (type 12 mặc dù có combo `PhuongPhapXacDinhGia` ở bảng lớn (`4897`) thì trong template các cell 12 đều mang Min/Max/Default=0 — xem combo là chính).

## Tổng kết

- **Đã map: 100% các type trong tần suất quan sát** (78 type số + họ DC_ 66 mã + DC_M). Không còn type nào "chưa rõ".
- Nhóm widget: text {0, 1, 17, 26}, số {16, 27, 33, 12*}, checkbox {2}, nút {15}, ngày/giờ {14, 19, 94, 96, 104, 109, 18}, ~60 type lookup combo (bảng "Nhóm lookup"), DC_N (combo file ngoài) / DC_N_M (checkbox-combo), unique-check {42, 43, 58}, auto-combo-Enter {106}.
- Điểm cần engine web lưu ý khi export: nil-attr cho {14, 26, 33, 96, 94} khi rỗng (`WDE:10463`); `2→"True"/"False"`, `16` rỗng→"0", `26` 13 ký tự→chèn "-" (`10444–10454`); mã của lookup phải nằm ở cột `SelectedValue`/`indexColumnSelectedValue`, không phải ô hiển thị; combo phụ (ChildValue/ChildCell) tự populate khi mã cha đổi (provincia→huyện→xã: 5/6/7, 100/101/102; CQT: 8/9).

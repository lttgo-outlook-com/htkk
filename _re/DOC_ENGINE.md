# HTKK Web — Kiến trúc engine render form từ CMC template

> Doc thiết kế (CHƯA build app). Người đọc cần đọc trước: `NOTES.md` §2 (data format), §6.3 (kỳ kê khai), §6.4a (luồng 01/GTGT).
> Mọi số đếm/sự thật trong doc này được đo trực tiếp từ `_extracted/` và `_decompiled/` trong phiên 2026-10-07; chỗ chưa verify được ghi rõ **[CẦN VERIFY]**, không bịa.

## 1. Mục tiêu & phạm vi

Mục tiêu: thay vì hard-code từng form khai thuế, web app **render form động** từ 3 loại file template sẵn có của HTKK (CMC form, skeleton XML GDT, XSD) + catalog `menu.xml` + bảng message — đúng cách HTKK gốc làm, để (a) trung thành với mẫu cục thuế, (b) import/export file XML ra đúng format GDT, (c) khi TCT đổi mẫu chỉ cần nạp template mới thay vì viết lại code.

Phạm vi doc này: (1) inventory asset nguồn, (2) schema 3 loại CMC file, (3) catalog menu, (4) data runtime, (5) pipeline render của HTKK gốc (reference model, kèm evidence), (6) kiến trúc engine web đề xuất, (7) mapping Controltype, (8) roadmap + rủi ro + câu hỏi mở. Không trong phạm vi: port logic nghiệp vụ (AutoBindData/validate) — đó là phase sau, spec từng form theo NOTES §3/§6.

## 2. Inventory asset nguồn (đo 2026-10-07 trong `_extracted/`)

| Asset | Số lượng | Encoding | Dùng cho |
|---|---|---|---|
| CMC **form** `{MaFile}_{VersionXML bỏ chấm}.xml` (vd `_0_th_dk_tct_101.xml`) | **1.235** | UTF-16LE (đa số; NOTES §2 ghi 2.125 CMC tổng hỗn hợp — 1.235 form + 761 skeleton + ~129 KHBS/đặc biệt ≈ 2.125) | Bản đồ cell form ↔ Path XML GDT |
| CMC **skeleton XML** `{MaFile}_xml_{v}.xml` (vd `_0_th_dk_tct_xml_101.xml`) | **761** | UTF-8 | Skeleton file GDT khi export |
| **XSD** GDT (vd `_0_th_dk_tct_101.xsd`) | **573** | UTF-8 | Validate file XML trước khi ký/nộp |
| **Spreadsheet layout** `{...}.xls` (+`.xls1`) | 569 (+420) | OLE2/BIFF (magic `D0CF11E0`) — FarPoint đọc | Bố cục ô, merge, **công thức** (NOTES §3.9: công thức nằm đây, không nằm trong DLL) |
| `menu.xml` | 1 file, 348 KB | UTF-8 | Catalog 314 Menu / 242 MaFile form-level / 1.011 Sheet |
| `message.xml` | 1 file, 71.479 B | UTF-8 | 691 dòng `<Infor ID Msg>` — bảng tra text lỗi (NOTES §3.9) |
| `catalogue_*` + `dm_*` + `khbs_*` | ~190 file | hỗn hợp | Danh mục tra cứu (địa bàn, CQT, ngành nghề, HS code…) — feed cho cell lookup (Controltype 55/56, `SelectedValue`) |
| `validcombobox.xml`, `ttinchung.xml`, `config_header.xml`, `cell_dynamic_format.xml`, `button.xml`, `mapcongthucchosua.xml`, `convertversion.xml`, `replacedatafile.xml`, `maptentk.xml`, `maptaidulieulon.xml` | 1 each | — | Cấu hình hành vi app (dropdown validate, cấu trúc Header, format cell động, map công thức sửa…) |

Runtime (không nằm trong cab, web app phải tự tạo): thư mục `DataFiles\<MST>\Header.xml` + file data mỗi form kỳ `<MaFile>_<kyKKhai>_<VersionXML bỏ chấm>.xml` (cấu trúc `<Sections>` — NOTES §2).

## 3. Schema 3 loại CMC file

### 3.1 CMC form (UTF-16LE) — "bản đồ" form

Root `<Sections Version="463">` chứa nhiều `<Section>`:

- **Section tĩnh** (`Dynamic="0"`): một tập `<Cell>` mô tả các ô cố định của sheet. Mỗi cell:
  - `CellID` — tọa độ ô trên **bản form nhập liệu** (vd `K_4`, hàng 4 cột K; trỏ file `InterfaceTemplates\Excel\{MaFile}_{v}.xls`),
  - `CellID2` — tọa độ trên **bản in/khai thác** (vd `J_7`; trỏ file `ReportTemplates\Excel\{MaFile}_{v}.xls` — xem §3.4),
  - `CellID3` (chỉ cell bảng động, parser `HTKKCMCSheetInfo.cs:251-254` → property `CellBangKe`) — **chữ cột trên file Excel "bảng kê" người dùng tải lên** (row đầu = `DataLocationTaiBangKeRowStart`). Mọi caller đều thuộc flow nút "Tải bảng kê" (`btnTaiBangKe_Click` `frmWinToKhai.cs:2592-2598` → `frmWinHTKKOpenFile.cs:591-627` → `WinSpreadProcess.GetDataFromExcelFile` :533/739/938/9783/9816). HTKK **chỉ đọc** file bảng kê theo mẫu TCT, không render/ghi layout thứ 3. 238/2.125 template có CellID3.
  - `Encode` — **KHÔNG phải mã hóa DataFile** (verify Q1 07/10): pipeline đọc/viết DataFile không branch theo `Encode` ở bất kỳ đâu (DataFile = clone template, chỉ ghi đè `Value`; bảng động nối bằng `###` — `HTKKCMCTableInfo.cs:331,353`). Nghĩa thật: `Encode="1"` = **cell được đưa vào chuỗi mã KHBS/barcode** — 4 XPath `Cell[@Encode='1']` duy nhất ở `WinDoWorkEvent.cs:7331,7382,7402,7425` (nối `~`, checkbox → "x"/rỗng, `:7435`); giá trị khác (0: 5.926, "25": 78, 2/3/4: hiếm) = không đưa vào; 1 ngoại lệ: checkbox `Encode="0"` trên menu 42/44 không đếm vào cờ `CoDuLieu` (`WinDoWorkEvent.cs:2560`).
  - `Controltype` — loại control (bảng §7),
  - `Value` — giá trị label/đặt trước (vd tên nút "Mẫu tải bảng kê"),
  - `DefaultValue`, `MaxLen`, `MinValue`, `MaxValue`, `SelectedValue` (đối với lookup),
  - `Path` — **đường dẫn node trong file XML GDT** (vd `HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/dangKyThue`; cell bảng động dùng path ngắn `recordID`, `ct2`, `ct4`, `CT6/ct6_ten`); cell không map ra XML (nút, label) có Path là tên control (`btnTBK`).
- **Section động** (`Dynamic="1"`, `id`, `TableName`, `TablePath` — vd `TablePath="HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/DSachNPTDaCoMST"`, `iColumnSTT`, `TableSTTValue`): mô tả **bảng có dòng lặp** (bảng kê). Chứa:
  - `<Location>` — 5 con số `…RowStart` (TableTemplateRowStart, TableReportRowStart, DataLocationTemplateRowStart, DataLocationReportRowStart, DataLocationTaiBangKeRowStart) — vị trí dòng đầu bảng trên 5 layout,
  - `<RowInfo>` — tập `<Cell>` theo **chữ cái cột** (`B`, `C`, `D`, `G`…) kèm `CellID/CellID2/CellID3` — 3 hệ tọa độ (form / khai thác / bảng kê), cùng bộ thuộc tính như cell tĩnh (Encode, Controltype, Path, MaxLen…).

Ví dụ thật (từ `_0_th_dk_tct_101.xml`, form đăng ký thuế TT86): dòng bảng `DSachNPTDaCoMST` có cell `D` = `ct4` (MST, MaxLen 14), `G` = `ct3` (Controltype 14 = ngày), `H` = `ct5` (Controltype 56 + `SelectedValue="AB"` = lookup mã địa bàn), `K` = `CT6/ct6_ten` (Controltype 55 + `SelectedValue="AC"`).

### 3.2 CMC skeleton XML (UTF-8) — "hình rỗng" file GDT

`_0_th_dk_tct_xml_101.xml` là **file XML GDT hoàn chỉnh với mọi element rỗng**: `HSoThueDTu/HSoKhaiThue/TTinChung/{TTinDVu (maDVu, tenDVu, pbanDVu, ttinNhaCCapDVu), TTinTKhaiThue (maTKhai, tenTKhai, moTaBMau, pbanTKhaiXML, mst, tenNNT, maCQTQuanLy, …, DLyThue…)} / CTieuTKhaiChinh / <các ct, bảng động (mỗi bản ghi có `recordID`), pattern `ctN_ma` + `ctN` cho cell lookup (mã + tên hiển thị)>`.

Vai trò trong app (evidence `WinDoWorkEvent.cs:842, 5700-5734, 10370, 10754`): engine nạp skeleton, **đổ giá trị cell DataFile vào đúng Path**, với naming đặc biệt cho KHBS (khai bổ sung) — `KHBS_01_TT80_xml.xml` / `KHBS_01_1_TT80_xml.xml`, biến thể `…_DVT_xml.xml` cho MenuID thuộc danh sách `"I7,I4,C7,I8,C8"` (đối chiếu địa bàn), `"KHBS_01_DIA_BAN_xml.xml"` cho danh sách `"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5"` — tức engine web phải copy nguyên mapping này khi hỗ trợ KHBS.

### 3.3 XSD GDT — validate trước export

Mỗi XSD khai `targetNamespace="http://kekhaithue.gdt.gov.vn/TKhaiThue"` và dùng **`<xs:redefine schemaLocation="..\TKhaiDkyThue.xsd">`** để mở rộng `CTieuTKhaiChinhType` base thêm các element của form đó (vd `dangKyThue` boolean, `thayDoiTTin` boolean, `DSachNPTDaCoMST` maxOccurs unbounded). System: base XSD nằm trong InterfaceTemplates (`TKhaiDkyThue.xsd`), engine web cần resolve relative reference này khi validate (hoặc preprocess gộp base vào từng form XSD).

### 3.4 .xls/.fxp — layout + công thức

File BIFF8 (OLE2, `biff_version: 80` theo xlrd — BIFF8 chuẩn Excel 97-2003, không phải dialect custom của FarPoint). Mỗi form có **2 template Excel cùng base name trong 2 folder**, app chỉ build path kết thúc `.xls` (`WinFnProcess.cs:857-898` `GetInterfaceTemplateExcel`/`GetReportTemplatesExcel`; grep `xls1` trong cả `_decompiled/` + `_extracted/*.xml` = 0 hit):

| Folder | Vai trò | Ai đọc | CMC trỏ bằng |
|---|---|---|---|
| `InterfaceTemplates\Excel\{MaFile}_{v}.xls` | Layout nhập liệu (form on-screen) | `frmWinToKhai.cs:417` khi mở form | `CellID` |
| `ReportTemplates\Excel\{MaFile}_{v}.xls` | Layout in / bản khai-thác | `frmHTKKReport.cs:24`, `frmHTKKPrintOption.cs:97,220` | `CellID2` |

**`.xls1` = file report trong cặp, bị đổi tên khi đóng gói cab phẳng**: 420/420 file `.xls1` đều có anh em `.xls` cùng base; verify bằng marker (cặp `_0_th_dk_tct_101`: file input chứa 4 ô marker `'chk'` trùng đúng CellID `K_4`/`O_4` trong CMC; file report chứa quốc hiệu, 0 ô `chk`). Quét 420 cặp: 184 cặp input/report phân biệt rõ, 4 cặp ngược, 127 cặp không marker (bảng kê/KHBS thuần — fallback so label CMC với cả 2 file). 569 − 420 = 149 form chỉ có 1 template. File "bảng kê" (74 file `Bang_Ke_*`/`PL_*`) là `.xls` thường riêng, không liên quan `.xls1`.

**Fidelity đọc trên web (verify Q2 07/10, SheetJS CE 0.18.5 + xlrd 2.0.2, 5 file đại diện: GTGT/TNCN/TNDN/TTS/KHBS — script `/tmp/htkk-biff-test/`)**:
- **ĐỦ phần cấu trúc**: 5/5 parse OK; merges (GTGT: 1.040), formulas kèm cached value 100% (GTGT: 303), sheet names, hidden rows/cols, row height, col width — SheetJS và xlrd cùng số ở mọi file. Hàm trong mẫu chỉ `IF/TRIM/SUM/COUNTIF/CONCATENATE/MOD/AND/OR/LOWER/MID` + tham chiếu cross-sheet → evaluator nhỏ viết được. Không có OBJECT/data-validation trong stream.
- **CHƯA ĐỦ phần visual style**: SheetJS CE **mất font (tên/đậm/cỡ), border, alignment, number format** (`z` luôn "General") — trong khi xlrd thấy thật (GTGT: 47 font, 465/839 XF có border; sheet 01-KHBS: 742 ô border, 2.647 ô bold). Form HTKK là form có khung (hàng trăm ô border, nhãn bold) → **phải pre-compute style bằng xlrd offline (chạy 1 lần khi build assets, ra JSON style per-form)**; web engine đọc JSON + SheetJS cho dữ liệu. Bắt buộc đọc SheetJS với `cellStyles:true` (khi `false`, `!cols` — chiều rộng cột + cột ẩn — biến mất im lặng).
- Mới test 5/569 file — nên chạy probe batch qua đủ 569 làm bước verify 1-lần trước khi chốt engine.

## 4. Catalog `menu.xml` — nguồn sự thật của "app có gì"

Cấu trúc (verify 07/10):

```
Root FirstTimeRunID="100_1" Version="100"
  Menu ID PopID ParentID Caption FormName Class Year FinanceYear Title Note Icon
  Menu(ID form): TenToKhai, MaHoSo, MoTaBieuMau, KieuKy, KieuKHBS, SignNode,
                 MaFile, VersionXML, VersionKHBS, MaToKhai, Class, XsdSchema
    Validity StartDate EndDate FolderName CatalogueID
      Sheet ID RowBreakPage RowTTDLT S?heetName MaHoSo Active DataFile XMLnode
```

Mỗi Menu form = 1 `MaFile` (tên DataFile gốc) + `VersionXML` + `Class` (business class, nối hậu tố `Nckt` theo NOTES §2) + `XsdSchema` + N `Sheet` (từng sheet có `DataFile`, `XMLnode` — prefix node trong XML GDT). **Chọn version theo ngày**: engine web phải replicate `frmHTKKPeriod.UpdatePeriodByVerSion` + `GeneralFunctions.checkVersionDateEffect` (NOTES §6.3) — so `ky` với `Validity StartDate/EndDate` của từng menu để chọn đúng mẫu hiệu lực tại kỳ (vd 01/XSBHDC-HKD: 01/01/2026–30/06/2026; từ 01/07/2026 chuyển mẫu MG-CNKD).

## 5. Data runtime (DataFiles)

- Layout: `DataFiles\<MST>\Header.xml` (38 field — NOTES §5.1) + mỗi form/kỳ 1 file XML cấu trúc `<Sections Version><Section Dynamic><Cells><Cell Path=… value…>` (NOTES §2).
- App lưu/đọc qua class `HTKKCMCXML` (evidence `WinDoWorkEvent.cs:7158-7277` nạp form CMC, `:7314-7316` nạp data CMC; `WinFnProcess.cs:244-245` cũng dùng cùng class).
- Web app: giữ nguyên format này làm storage (file-per-MST) để tương thích 100% import/export với HTKK gốc; hoặc dùng storage riêng + 2 bộ converter (rủi ro drift — khuyến nghị phase 1 dùng thẳng format gốc).

## 6. Pipeline render của HTKK gốc (reference model)

Thứ tự khi người dùng mở 1 tờ khai (reconstruct từ `_decompiled/htkk`, evidence kèm):

1. **Chọn menu + kỳ** — `frmHTKKPeriod` (`LoadDefaultInfor` :6106–6340 theo PT; `UpdatePeriodByVerSion` chọn version theo Validity; `SetPeriodCustom` cho menu đặc biệt).
2. **Nạp form CMC** vào `ToKhaiInfo.SheetInfos[]` (mỗi sheet: `StaticCells[]`, `Tables[]` = `HTKKCMCTableInfo` có `Header.StaticCells` + `RowsInfo`) — class `HTKKCMCXML`, engine `WinDoWorkEvent.cs:842` (naming skeleton xml theo MenuID/KHBS).
3. **Render spreadsheet** — FarPoint grid theo layout `.xls`/`.fxp` (InterfaceTemplates); cell CMC gắn vào ô (CellID) — control type quyết định widget (checkbox, date, nút, lookup…). `validcombobox.xml` feed dropdown.
4. **Kéo số** — business class `AutoBindData_N` (per version, NOTES §3) chạy khi mở; kéo dữ liệu kỳ trước/PL/header theo cell Path.
5. **Validate** — `isValidDuLieuTK` → `checkTKChinh` + `checkPLx_y` (7.925 điểm, NOTES §5.2); kết quả = `Dictionary<int,int> levelWarning` (0/1/2) → tô ô hồng/vàng; text lỗi = cell ẩn `E_xx` trong template hoặc ID `message.xml`.
6. **Export** — `WinDoWorkEvent.cs:7158-7277`: build từ skeleton XML (§3.2), đổ cell, `SetAttribute("Version", …)`, save → tmpExport → **validate XSD** → hash `ttinNhaCCapDVu` (3DES/MD5, NOTES §4.1) → lưu DataFile + file nộp.
7. **Print** — layout khai thác (CellID2) → FarPoint → PDF (`PdfWriteTo.File`, NOTES §5.1) + barcode nội dung (`GetHeaderStringBarcode` `WinFnProcess.cs:943-965`).

## 7. Mapping Controltype (đã verify + chưa verify)

Đã verify bằng code app (`WinDoWorkEvent.cs:2533-2580` switch + NOTES §3.2) và mẫu template:

| Type | Count (toàn 1.235 form, đo 07/10) | Nghĩa | Evidence |
|---|---|---|---|
| 0 | 19.410 | text (MaxLen) | mẫu + switch mặc định |
| 2 | 2.260 | checkbox (True/False) | `:2558` |
| 14 | 1.257 | ngày dd/mm/yyyy | `CheckFormatDdmmyyyy` :2534 |
| 15 | 165 | nút (không export) | `:2580` skip |
| 16 | 34.471 | số (round/export theo NOTES §3.2) | `:2551` |
| 26 | 602 | text dài (vd MST, MaxLen 14–30) | mẫu `ct4` |
| 55/56 | 4/171 | lookup (combo tra danh mục, `SelectedValue` = mã catalogue) | mẫu + `CheckedComboBoxCellType.cs` |
| 104 | 6 | ngày dd/mm | `CheckFormatDdmm` :2539 |
| 109 | 25 | năm yyyy | `CheckFormatYyyy` :2544 |

**Bảng đầy đủ (verify Q6 07/10 — 100% type quan sát đã map, không còn type "chưa rõ")**: xem **`_re/CONTROLTYPE_MAP.md`** — 3 dispatch render (`fpSpFormatCellTemplate` `WinSpreadProcess.cs:4186`, `fpSpFormatTableCellTemplate` :4432, `fpSpFormatTableCell` :4819) + ~60 type lookup combo (từng type → tên DM + dòng mapping mã trong `frmWinToKhai.cs:12828-14530`) + nhóm `DC_N` (combo từ `dm_dung_chung.xml`, 327 cell, 66 mã) + `DC_*_M_*` (checkbox-combo). Type default-text chỉ có **1** và **52**; type không tồn tại ở cả code lẫn template: 13, 24, 30, 31, 63, 95, 97, 99. Điểm bắt buộc khi export (mất là sai file): nil-attr `xsi:nil="true"` cho {14, 26, 33, 96, 94} khi rỗng (`WinDoWorkEvent.cs:10463-10491`); `2` rỗng → "0"/"1", `16` rỗng → "0", `26` 13 ký tự → chèn "-" (`:10444-10454`); mã lookup nằm ở cột `SelectedValue`/`indexColumnSelectedValue`, không phải ô hiển thị; combo cha→con tự populate (tỉnh→huyện→xã: 5/6/7, 100/101/102; CQT cục→chi cục: 8/9).

## 8. Kiến trúc engine web đề xuất

```
assets/  (từ _extracted/, sạch encoding — build script normalize về UTF-8)
  menu.xml · message.xml · validcombobox.xml · catalogue_* (190)
  cmc/{MaFile}_{v}.form.xml (1235) · cmc/{MaFile}_xml_{v}.skeleton.xml (761)
  xsd/*.xsd (573, đã gộp base) · layout/*.xls (569, đọc bằng SheetJS)
  khbs_naming.map (danh sách MenuID từ WinDoWorkEvent.cs:842/5717)

engine/
  FormSpec.ts      — parse menu.xml + CMC form → FormSpec { menuId, version, period:PT,
                      sheets: [ { dataFile, xmlNode, staticCells[], tables[] } ] }
  render/          — grid (AG Grid/Univer) theo layout .xls; widget per Controltype (§7);
                      lookup cell → catalogue_*; date cell → mask dd/mm/yyyy
  state/           — CellState per (sheet, cellId, row?): value/lock/warning(0-2)/source
                      ("input" | "autobind" | "formula" | "carryover"); storage = DataFiles\<MST>\*.xml (format gốc)
  rules/           — hooks: onOpen(sheet) → AutoBindData_N; onCellChange(cell) → CellChange_N;
                      validate(sheet) → checkTKChinh/checkPL (port từng class business — phase 2)
  export/          — skeleton GDT + đổ value theo Path (Encode không dùng ở bước này — §3.1) + rounding (NOTES §3.2)
                      → validate XSD (XSD-processor web) → hash ttinNhaCCapDVu (algo NOTES §4.1)
                      → file nộp + DataFile
  period/          — PT mapping (NOTES §6.3) + chọn version theo Validity (NOTES §6.3)
  i18n/            — message.xml (691 Infor) + text E_xx đọc từ template
```

Quyết định kiến trúc chốt (từ scope 07/10): (1) **engine trước, rules sau** — phase 1 render + import/export trung tính (không rule nghiệp vụ), phase 2 port rules từng form; (2) **giữ format DataFile/XML GDT gốc 100%** để tương thích HTKK (điểm bán chính của web app); (3) **KHBS = naming mapping cứng** (§3.2) không tự phát; (4) spreadsheet grid là lõi UI — cell CMC là lớp mapping đè lên grid (đúng model HTKK: CMC không phải layout, CMC chỉ map ô→Path); (5) **style BIFF pre-compute offline bằng xlrd** (verify Q2: SheetJS CE giữ đủ merges/formulas/geometry nhưng mất font/border/alignment — JSON style per-form sinh 1 lần khi build assets); (6) **pairing .xls/.xls1 theo base name + heuristic marker** (`chk` = input, quốc hiệu = report; 4 cặp ngược, 127 cặp không marker — fallback so label CMC, §3.4).

Rủi ro chính (sau 6 câu hỏi verify 07/10): (a) **style visual** — đã chốt hướng mitigate (xlrd offline, §3.4); còn rủi ro probe mới chạy 5/569 file → batch probe trước khi chốt; (b) formula evaluator — tập hàm nhỏ (IF/TRIM/SUM/COUNTIF/…) nhưng form tương tác cần tính lại khi gõ → phải test trên đủ 569 file; (c) XSD `xs:redefine` cross-file — mitigate: preprocess gộp base XSD lúc build assets; (d) version hiệu lực theo ngày — replicate `checkVersionDateEffect` đúng, test với cặp menu 01/XSBHDC-HKD ↔ MG-CNKD (đổi mẫu 01/07/2026). Controltype đã map 100% (CONTROLTYPE_MAP.md) — không còn rủi ro "widget sai loại".

## 9. Roadmap đề xuất (mỗi phase = 1 deliverable kiểm được)

1. **P0 — Asset pipeline**: script normalize 2.125 CMC + 573 XSD + menu.xml + message.xml + catalogue về UTF-8, gộp base XSD, sinh `khbs_naming.map`; output = thư mục `assets/` sạch + file kiểm `assets/manifest.json` (số file, hash). Accept: mở lại 100% file qua parser web không lỗi encoding.
2. **P1 — FormSpec + render 01/GTGT**: parse menu (MenuID D1) + CMC form 2.8.3 + layout .xls input (pairing §3.4) + JSON style xlrd → grid web đúng bố cục; cell editable theo Controltype (CONTROLTYPE_MAP.md); lưu/đọc DataFile format gốc (UTF-8, `Value` nguyên dạng, bảng động separator `###`). Accept: nhập số vào 01/GTGT web → đọc lại được bằng parser DataFile (roundtrip giá trị cho mọi Controltype, kể cả checkbox "True"/"False" và ngày dd/mm/yyyy).
3. **P2 — Export GDT đúng chuẩn**: skeleton + Path + Encode + rounding → file XML; validate XSD 01_GTGT_TT80; hash ttinNhaCCapDVu khớp HTKKConstant key. Accept: file export qua XSD + diff structure so file mẫu HTKK (cùng input) = cùng node.
4. **P3 — Rules 01/GTGT**: port `Gtgt01TT80Nckt283` (checkTKChinh, V1–V8 NOTES §3.3/§3.7, KHBS ct4/ct5/ct6) + AutoBindData (kéo CT43 kỳ trước). Accept: 22 case test từ NOTES §3.3 (tol min, 5%/10% row 33/34, ct21 lock 12 cell).
5. **P4 — Mở rộng**: engine chung chạy được 242 form (render-only) → chọn form thứ 2 theo nghiệp vụ (05/KK-TNCN hoặc 03/TNDN — quyết định khi P3 xong).

## 10. Câu hỏi mở — ĐÃ VERIFY (6/6, phiên 07/10, 6 agent song song + spot-check)

1. **Encode (ĐÃ VERIFY — SAI giả định gốc)**: không phải mã hóa DataFile. DataFile đọc/ghi `Value` nguyên dạng, branch kiểu theo `Controltype` (không branch `Encode`); nghĩa thật `Encode="1"` = cell đưa vào **chuỗi mã KHBS/barcode** (`WinDoWorkEvent.cs:7331,7382,7402,7425`, nối `~`, checkbox → "x"/rỗng). Ngoại lệ duy nhất: checkbox `Encode="0"` menu 42/44 không đếm cờ `CoDuLieu` (`:2560`). → Engine: parse DataFile theo Controltype; chỉ cần quan tâm Encode khi sinh chuỗi barcode KHBS.
2. **BIFF/SheetJS (ĐÃ VERIFY, 5/569 file)**: BIFF8 chuẩn (xlrd `biff_version: 80`); SheetJS CE giữ đủ merges/formulas(+cached)/names/hidden/row-height/col-width (bắt buộc `cellStyles:true`) nhưng **mất font/border/alignment/number-format** → pre-compute style bằng xlrd offline (JSON per-form). Hàm formula chỉ `IF/TRIM/SUM/COUNTIF/CONCATENATE/MOD/AND/OR/LOWER/MID` + cross-sheet. Chưa test đủ 569 → batch probe trước khi chốt.
3. **`.xls1` (ĐÃ VERIFY)**: không phải extension app đọc (grep `xls1` code+template = 0) — là **file report** (`ReportTemplates\Excel`) của cùng form bị đổi tên khi đóng gói cab phẳng; app build path luôn `.xls` (`WinFnProcess.cs:857-898`). Pairing: 420 cặp (184 marker rõ, 4 ngược, 127 không marker) + 149 form đơn. Phase 1 chỉ cần file input.
4. **`CellID3` (ĐÃ VERIFY)**: không phải layout thứ 3 — là **chữ cột trên file Excel bảng kê người dùng tải lên** (`CellBangKe`, parser `HTKKCMCSheetInfo.cs:251-254`; mọi caller thuộc flow "Tải bảng kê" `WinSpreadProcess.GetDataFromExcelFile` :533 v.v.). HTKK không render layout bảng kê. Phase 1 chỉ parse + lưu trường; implement khi làm tính năng tải bảng kê (238/2.125 template có).
5. **SignNode vs CKyDTu (ĐÃ VERIFY — không mâu thuẫn)**: node ký thật = **`HSoKhaiThue`** (hardcode `WinDoWorkEvent.cs:7505`, `WinDoWorkSignXml` def :7454, caller duy nhất `frmHTKKSystems.cs:905` menu 104_2) — `id="_NODE_TO_SIGN"` + `Reference "#_NODE_TO_SIGN"` (`CA/XmlDsigEnveloping.cs:46-49`), `XmlDsigEnvelopedSignatureTransform(true)`, C14N mặc định, KeyInfo RSAKeyValue+X509Data+SubjectName, **`CKyDTu` là bao bì chứa `ds:Signature`, chèn vào root `HSoThueDTu` làm em của `HSoKhaiThue`** (`:7521`). `SignNode` trong menu.xml (240 entry, 100% = `HSoKhaiThue`) = **attribute chết** (parse `WinFnProcess.cs:357`, 0 reader). Guard "đã ký" = presence `CKyDTu` (`:7508-7509`); `_signed.xml` chỉ là default tên SaveFileDialog. Gate trước ký: cert từ X509Store, MST cert == MST đăng nhập, cert còn hạn. Web app: target `HSoKhaiThue` + wrapper `CKyDTu` vào root; thêm bước verify (HTKK gốc không verify sau ký).
6. **78 Controltype (ĐÃ VERIFY — 100% map)**: bảng đầy đủ trong **`_re/CONTROLTYPE_MAP.md`** — 3 switch render + helpers, ~60 type lookup combo (mỗi type → DM + dòng mapping mã), `DC_N`/`DC_*_M_*`, default-text chỉ {1, 52}. Export invariants: nil-attr {14,26,33,96,94}, `2`→"True"/"False", `16` rỗng→"0", `26` 13 ký tự→chèn "-", mã lookup ở cột `SelectedValue`.

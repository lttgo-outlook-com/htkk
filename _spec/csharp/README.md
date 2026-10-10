# _spec/csharp

Sinh bởi `_spec/tools/roslyn_dump` (Roslyn chỉ ParseText, không biên dịch). Chạy lại: xem Program.cs.

## in_menu (classes.json)
- Class nghiệp vụ: lớp cấp cao, namespace TAX.HTKK.Business, assembly tax.htkk.business (394).
- Tên class = menu.xml `Menu@Class` + "Nckt" + `VersionXML` bỏ dấu chấm. Ví dụ Gtgt01 + 2.6.2 = Gtgt01Nckt262.
- Khớp phiên bản chính xác: 242 class in_menu=true (kèm MenuID, MaFile, VersionXML). 152 class còn lại in_menu=false, trong đó 131 class cùng tiền tố Class nhưng khác phiên bản.
- Menu có 243 mục có Class; mục GeneralInfo (100_1) không thuộc business.

## hook (index.json)
- hook_kind = CellChange, AutoBindData, Refresh, UpdateFormula, ClearData, CellEditchange, ButtonClick; tên có thể có hậu tố `_N` (sheet_index = N) hoặc không.
- Validate: tên bắt đầu bằng "Validate" (gồm Validate_N, ValidatePeriod_*, hàm tiện ích). Constructor ValidateInfo là .ctor, không tính.
- return_type: kiểu trả về của method (void hoặc bool; null với constructor). UpdateFormula có cả bản void và bool.

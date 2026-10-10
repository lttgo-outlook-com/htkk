# _spec

Đặc tả máy đọc được của HTKK 5.7.9, sinh từ `_extracted` (tệp gốc đã giải nén) và `_decompiled` (mã C# decompile). Đường dẫn nguồn tuyệt đối: `/Users/tuanle/codes/HTKK/_extracted`, `/Users/tuanle/codes/HTKK/_decompiled`. Mọi lệnh chạy từ gốc repo, dùng `python3 -I` (cần olefile thì dùng `/tmp/audit_venv/bin/python -I`).

## Thư mục con

| Thư mục | Nội dung | Nguồn | Công cụ sinh | Lệnh chạy lại | QC |
|---|---|---|---|---|---|
| `csharp` | Lớp, method, hook (CellChange, AutoBindData, Validate...) của mã C# | `_decompiled` + `_extracted/menu.xml` | `tools/roslyn_dump` (Roslyn ParseText) | `dotnet run --project _spec/tools/roslyn_dump -- /Users/tuanle/codes/HTKK/_decompiled /Users/tuanle/codes/HTKK/_extracted/menu.xml _spec/csharp` (xem `csharp/README.md`) | T0 đạt sau sửa |
| `formulas` | Công thức BIFF của từng template .xls | `_extracted/*.xls`, `menu.xml` | `tools/biff_formulas.py` (+ `sheetjs_dump.cjs` để đối chiếu) | `/tmp/audit_venv/bin/python -I _spec/tools/biff_formulas.py --src /Users/tuanle/codes/HTKK/_extracted --menu /Users/tuanle/codes/HTKK/_extracted/menu.xml --out _spec/formulas`; kiểm: `verify_formulas.py` | T1 đạt có điều kiện; đã sửa lost_refs, CR, khóa |
| `config` | Tệp cấu hình XML/DTD (ttinchung, validcombobox, versiontokhai, schema.dtd...) | `_extracted` + `_decompiled` | `tools/config_xml.py` | `python3 -I _spec/tools/config_xml.py [--src DIR] [--cs DIR] [--out DIR]` | T4 đạt sau sửa |
| `xsd` | XSD gộp và chuẩn hóa theo tờ khai + phiên bản (608 tệp) | `_extracted` (XSD), `_forms`, `menu.xml` | `tools/xsd_merge.py` | `python3 -I _spec/tools/xsd_merge.py [--src DIR] [--out DIR] [--forms DIR] [--menu FILE]` | T5 đạt |
| `catalog` | Danh mục XML (maptentk, replacedatafile, validcombobox, versiontokhai) sang JSON | `_extracted`, `_decompiled` | `tools/catalog.py` | `python3 -I _spec/tools/catalog.py [--src DIR] [--dec DIR] [--out DIR]` | T6 đạt |
| `app` | Các điểm rẽ nhánh theo menuID/MaFile/MaToKhai trong ứng dụng | `_decompiled`, `menu.xml` + bản menu cũ | `tools/app_branches.py` rồi `tools/menu_source.py` | `python3 -I _spec/tools/app_branches.py` rồi `python3 -I _spec/tools/menu_source.py`; độ đúng: `app_branches.py --accuracy` | T7 đạt; phân loại tác động chỉ là gợi ý (13/20) |
| `dv` | Data validation (DV/DVAL) trong template .xls | `_forms` | `tools/biff_dv.py` | `python3 -I _spec/tools/biff_dv.py _spec/dv /Users/tuanle/codes/HTKK/_forms` (xem `dv/README.md`) | T9: DV kiểu list có tác dụng theo tài liệu FarPoint; loại khác chưa xác định |
| `_forms/**/cong_thuc.txt` (view) | Công thức từng tệp bảng tính trong `_forms`, mỗi ô một dòng `REF | công thức | nhãn | nguồn` | `formulas`, `csharp/classes.json`, `config/mapcongthucchosua.xml.json`, `_forms/_manifest.csv`, T2 `rules` (tùy chọn) | `tools/render_cong_thuc.py` | `/tmp/audit_venv/bin/python -I _spec/tools/render_cong_thuc.py --repo . [--rules _spec/rules]`; kiểm: `verify_formulas.py`, `_re/check_mapping.py` | Sinh lại toàn bộ, không sửa tay; render hai lần giống từng byte |
| `tools` | Mã sinh và kiểm cho các thư mục trên | n/a | n/a | n/a | n/a |

## Trạng thái gộp

- Nhánh T5 gộp theo commit a162367; trạng thái worktree T5 chưa kiểm do guard.
- T2, T3, T8 chưa gộp (đang sửa).
- Phân loại `impact` trong `app` chỉ là gợi ý, cần đọc tay trước khi dùng.

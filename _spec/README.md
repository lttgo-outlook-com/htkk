# _spec

## Mục tiêu

Decompile và lập danh mục toàn bộ công cụ HTKK 5.7.9 của Tổng cục Thuế (logic, thuật toán, dữ liệu, hành vi tĩnh). Không nhằm dựng web app. `_spec` là đặc tả máy đọc được (JSON, CSV) sinh từ hai nguồn:

- `/Users/tuanle/codes/HTKK/_extracted`: 4.526 tệp giải nén từ bộ cài (template .xls/.xls1, XSD, XML cấu hình, DLL).
- `/Users/tuanle/codes/HTKK/_decompiled`: mã C# decompile của 4 assembly lõi.

Ba thư mục cùng cấp: `_forms` (template tổ chức theo form, kèm `_manifest.csv` và `cong_thuc.txt` từng form), `_re` (ghi chú reverse engineering `NOTES.md`, `reorg_forms.py`, `check_mapping.py`) và `_spec`. Chi tiết nghiệp vụ và kiến trúc: `_re/NOTES.md`.

Quy ước chạy: từ gốc repo; Python dùng `-I`; cần olefile hoặc openpyxl thì dùng `/tmp/audit_venv/bin/python -I`. Mọi công cụ không chạy mã trong `_extracted`, `_decompiled`: Roslyn chỉ `ParseText`, BIFF đọc byte, XML parse bằng bộ cấm DTD.

## Thư mục con

Cột QC: trạng thái sau lần chạy `sh _spec/tools/rebuild_all.sh` gần nhất (xem mục "rebuild_all.sh"). T0 đến T11 là mã việc trong job lập danh mục.

| Thư mục | Nguồn | Công cụ sinh | Lệnh chạy lại | Số liệu chính | QC |
|---|---|---|---|---|---|
| `csharp` (T0) | `_decompiled`, `_extracted/menu.xml` | `tools/roslyn_dump` (Roslyn) | dotnet, xem mục "Lệnh dotnet" | 394 class nghiệp vụ (242 in_menu), 586 tệp, 9.113 method, 0 lỗi parse | Đạt sau sửa; không chạy lại trong `rebuild_all.sh` (cần dotnet) |
| `formulas` (T1) | `_extracted/*.xls`, `menu.xml` | `tools/biff_formulas.py` (+ `sheetjs_dump.cjs` đối chiếu) | `/tmp/audit_venv/bin/python -I _spec/tools/biff_formulas.py --src /Users/tuanle/codes/HTKK/_extracted --menu /Users/tuanle/codes/HTKK/_extracted/menu.xml --out _spec/formulas` | 1.028 tệp, 452.819 công thức, 0 lỗi giải mã | `verify_formulas.py` ĐẠT (mục 3 đã chuẩn hóa cả dạng `'#REF'!E26`) |
| `rules` (T2) | `_decompiled`, `csharp/classes.json`, `tools/rules_extract/lock_config.json` | `tools/rules_extract` (Roslyn) | dotnet, xem mục "Lệnh dotnet" | 405 tệp JSON (theo class, `_app`, `_shared`) | `rules_extract/verify.py` ĐẠT; nhánh khác còn sửa thêm, gộp xong thì chạy lại `rebuild_all.sh` |
| `validations` (T3) | `_decompiled`, `csharp/classes.json` | `tools/roslyn_validations` + `tools/validations/{build,postprocess}.py` | dotnet qua `tools/validations/run.sh`; phần Python: `postprocess.py` | 8.000 check, 394 class, 6 class `_app` | `validations/verify.py`, `scan_idents.py` ĐẠT |
| `config` (T4) | `_extracted` (XML cấu hình), `_decompiled` | `tools/config_xml.py` | `python3 -I _spec/tools/config_xml.py --src /Users/tuanle/codes/HTKK/_extracted --cs /Users/tuanle/codes/HTKK/_decompiled --out _spec/config` | 40 tệp cấu hình + `index.json` | ĐẠT: kiểm vòng XML-JSON-XML lệch 0; tái tạo giống từng byte |
| `xsd` (T5) | `_extracted` (XSD), `_forms`, `menu.xml` | `tools/xsd_merge.py` | `python3 -I _spec/tools/xsd_merge.py --src /Users/tuanle/codes/HTKK/_extracted --out _spec/xsd --forms "$PWD/_forms" --menu /Users/tuanle/codes/HTKK/_extracted/menu.xml` | 607 schema gốc, 608 JSON (+ `index.json`), 205.320 entry; menu 242 MaFile+VersionXML, 3 không có XSD | ĐẠT: số đếm nguồn = JSON + không đến được; 0 tham chiếu không giải được |
| `catalog` (T6) | `_extracted` (catalogue_*, dm_*, convert*), `_decompiled` | `tools/catalog.py` | `/tmp/audit_venv/bin/python -I _spec/tools/catalog.py --src /Users/tuanle/codes/HTKK/_extracted --dec /Users/tuanle/codes/HTKK/_decompiled --out _spec/catalog` | 530 tệp kiểm kê, 530 chuyển được, 0 không chuyển | ĐẠT: số ô BIFF và số phần tử XML khớp đếm độc lập |
| `app` (T7) | `_decompiled`, `menu.xml` + bản menu cũ | `tools/app_branches.py`, `tools/menu_source.py` | `python3 -I _spec/tools/app_branches.py` rồi `python3 -I _spec/tools/menu_source.py` | 1.172 điểm rẽ nhánh, 20 điểm dispatch động | `app_branches.py --accuracy` ĐẠT; phân loại tác động chỉ là gợi ý (13/20) |
| `helpers` (T8) | `validations`, `rules`, `csharp` | `tools/helpers.py` (+ overlay tay `manual_overlay*.json`) | `python3 -I _spec/tools/helpers.py` | 86 hàm helper phủ, tỷ lệ phủ 95,079% (ngưỡng 95%) | ĐẠT: cả hai định nghĩa mẫu số >= 95%, 0 helper thiếu overlay |
| `dv` (T9) | `_forms` | `tools/biff_dv.py` | `/tmp/audit_venv/bin/python -I _spec/tools/biff_dv.py _spec/dv _forms` | 942 tệp quét, 392 tệp có DV, 1.260 bản ghi DV, 876 DVAL, 0 lỗi đọc | ĐẠT: `check_counts.py dv` + tái tạo giống từng byte; DV kiểu list có tác dụng theo tài liệu FarPoint, loại khác chưa xác định |
| `_forms/**/cong_thuc.txt` (T10) | `formulas`, `csharp/classes.json`, `config/mapcongthucchosua.xml.json`, `_forms/_manifest.csv`, `rules` | `tools/render_cong_thuc.py` | `/tmp/audit_venv/bin/python -I _spec/tools/render_cong_thuc.py --repo .` | 268 tệp `cong_thuc.txt`; mỗi ô một dòng `REF \| công thức \| nhãn \| nguồn` (nguồn `xls`/`xls1`/`xlsx`, `ROWREPLACE`, `C#:Class.Hook:Lnn`) | `verify_cong_thuc.py`, `check_mapping.py` ĐẠT; render hai lần giống từng byte |
| `coverage` (T11) | `xsd`, `dv`, `formulas`, `rules`, `validations`, `csharp`, `_forms` | `tools/coverage.py` | `python3 -I _spec/tools/coverage.py --out _spec/coverage` | 242 form (MaFile, VersionXML), mỗi form một JSON, 246 tệp gồm `conflicts.json`, `sheet_log.json`, README | ĐẠT: chạy lần hai ra thư mục tạm giống từng byte |
| `tools` | n/a | mã sinh và kiểm cho các thư mục trên, `rebuild_all.sh`, `check_counts.py` | n/a | n/a | n/a |

Quy tắc C# trong `cong_thuc.txt`: chỉ effect `setFormula` của bản ghi không `unreachable`; `setValue` và `lock` chỉ đếm ở dòng `# C# tóm tắt sheet:` cuối mỗi sheet; target có placeholder hoặc sheet tham số ghi nguyên ký hiệu; class gắn vào form theo (MaFile, VersionXML) qua `classes.json` hoặc suy từ tên `<MenuClass>Nckt<số phiên bản>`, một lần vào tệp đầu tiên của form (xls trước xls1). Tệp chỉ là view, không sửa tay.

## Lệnh dotnet sinh lại T0, T2, T3

T0, T2, T3 cần dotnet (SDK có Roslyn qua NuGet) nên `rebuild_all.sh` không chạy. Dùng giới hạn tài nguyên sau để không làm treo máy; chạy tuần tự, từng lệnh một, và chờ load 1 phút dưới 10 trước khi chạy:

```sh
export DOTNET_gcServer=0 DOTNET_GCHeapHardLimit=0x100000000 DOTNET_TieredPGO=0 MSBUILDDISABLENODEREUSE=1 \
       DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_NOLOGO=1
D=/Users/tuanle/codes/HTKK
T=$(mktemp -d /tmp/htkk_dn.XXXXXX)

# T0: _spec/csharp
nice -n 15 dotnet build _spec/tools/roslyn_dump/RoslynDump.csproj -c Release -maxcpucount:1 \
  -p:UseSharedCompilation=false --disable-build-servers -o "$T/dump"
dotnet build-server shutdown
nice -n 15 dotnet "$T/dump/RoslynDump.dll" $D/_decompiled $D/_extracted/menu.xml _spec/csharp

# T2: _spec/rules (đối số cuối là cấu hình khóa/mở ô sinh bởi lock_config.py)
nice -n 15 dotnet build _spec/tools/rules_extract/RulesExtract.csproj -c Release -maxcpucount:1 \
  -p:UseSharedCompilation=false --disable-build-servers -o "$T/rules"
dotnet build-server shutdown
nice -n 15 dotnet "$T/rules/RulesExtract.dll" $D/_decompiled _spec/csharp _spec/rules _spec/tools/rules_extract/lock_config.json

# T3: _spec/validations (run.sh tự đặt các biến trên, tự build với -maxcpucount:1, tự shutdown build server)
nice -n 15 sh _spec/tools/validations/run.sh $D _spec
rm -rf "$T"
```

Lệnh trên theo `Program.cs` và `run.sh` của từng công cụ; không chạy lại trong lần hoàn tất này (brief cấm dotnet). Sau khi sinh lại, chạy `sh _spec/tools/rebuild_all.sh` để kiểm. `rules_extract/lock_config.json` sinh bởi `rules_extract/lock_config.py` (tham số ở docstring của tệp).

## rebuild_all.sh

`sh _spec/tools/rebuild_all.sh` (từ bất kỳ đâu; biến môi trường `PY`, `VPY`, `HTKK_DATA`). Không gọi dotnet. Chạy tuần tự với `nice -n 15`, chờ khi load 1 phút > 10.

1. Bước sinh: xóa JSON cũ rồi sinh lại `xsd` (T5) và `dv` (T9); chạy `validations/postprocess.py` (T3), `helpers.py` (T8), `render_cong_thuc.py` (T10), `coverage.py` (T11).
2. Bộ kiểm, mỗi bộ in một dòng `KIỂM <tên>: ĐẠT` hoặc `KIỂM <tên>: KHÔNG ĐẠT`: `check_mapping`, `verify_formulas`, `verify_cong_thuc`, T2 verify, T3 verify, T3 `scan_idents` (thoát mã 1 khi lệch), T4 (kiểm vòng + tái tạo ra thư mục tạm, `diff -rq`), T5 (số đếm + tái tạo), T6 (số ô + tái tạo), T9 (đếm DV + tái tạo), T8 tỷ lệ phủ, T11 tất định, T7 `--accuracy`.
3. Dòng cuối `TỔNG KẾT`; thoát mã 0 khi mọi bộ kiểm ĐẠT, 1 khi có KHÔNG ĐẠT, 2 khi thiếu công cụ, 3 khi máy quá tải. Chạy hai lần liên tiếp cho kết quả giống từng byte (`git status` sạch sau lần hai).

Lưu ý: `_spec/xsd` và `_spec/dv` lưu từ trước reorg `_forms` không còn khớp (tên tệp và số đếm lệch); lần sinh lại này đã cập nhật. `_spec/dv/index.json` ghi root tương đối `_forms`.

## Giới hạn đã biết

- DV loại khác `list` (whole, decimal, textLength, any) chưa xác định có tác dụng khi nạp hay không: tài liệu API FarPoint chỉ mô tả danh sách giá trị và cờ combo (`_spec/dv/README.md`).
- `LoaiToKhai` chỉ biết lúc chạy, nên nhánh KHBS mẫu chung (`WinFnProcess.cs:74`, KieuKHBS == "3" và LoaiToKhai == "B") chỉ là ứng viên; `coverage` không xác nhận được nhánh nào áp dụng.
- 86 effect cột bảng động không gắn được vào ô cụ thể (độ dài bảng động chỉ biết lúc chạy).
- 3 form không có XSD (`xsd_merge.py` in "khong tim thay XSD: 3").
- 29 class không map được vào form trong `cong_thuc.txt` (không suy ra được form từ tên class); `render_cong_thuc.py` in danh sách khi render, các class này không có dòng.
- Phân loại tác động (`impact`) của T7 là gợi ý (13/20 điểm dispatch được xác nhận), cần đọc tay trước khi dùng.
- StatusID `0001` dùng làm "ô nhập" trong `coverage` là heuristic: không có mã C# nào đọc StatusID, chỉ có hằng `HTKKConstant.cs:49-51`.
- Mã HTKK ghi Formula vào cột BB thay vì BA trong `checkLockUnLockCells` (lỗi gốc của HTKK, không phải lỗi của bộ trích); đặc tả giữ nguyên hành vi của mã gốc.
- 40 menuID chỉ có trong mã, không rõ nguồn (không có trong `menu.xml` hay bản menu cũ).
- 15 tệp pending trong `_re/pending_mapping.csv` chưa xác định được form (stem không nhận dạng được, hoặc tên cab bị cắt khớp nhiều tên gọi); không nằm trong `_forms` theo form.
- Kết luận từ tệp này là mô tả tĩnh của mã và template; áp dụng vào nghiệp vụ thuế thực tế cần người có chuyên môn kiểm tra với văn bản pháp luật hiện hành.

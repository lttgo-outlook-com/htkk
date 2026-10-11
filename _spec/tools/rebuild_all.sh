#!/bin/sh
# rebuild_all.sh: chạy lại tuần tự các bước Python tái lập được từ JSON đã có, rồi chạy mọi bộ kiểm.
# Không gọi dotnet (T0 csharp, T2 rules, T3 validations raw được sinh bằng dotnet; lệnh ở _spec/README.md).
# Dùng (từ bất kỳ đâu):  sh _spec/tools/rebuild_all.sh
# Biến môi trường: PY (mặc định python3), VPY (python có olefile/openpyxl, mặc định /tmp/audit_venv/bin/python),
#                  HTKK_DATA (thư mục chứa _decompiled và _extracted, mặc định /Users/tuanle/codes/HTKK).
# Mọi bước chạy với nice -n 15, nối tiếp nhau (không song song); trước mỗi bước nặng chờ khi load 1 phút > 10.
# Mỗi bộ kiểm in đúng một dòng 'KIỂM <tên>: ĐẠT' hoặc 'KIỂM <tên>: KHÔNG ĐẠT'.
# Thoát mã 0 khi mọi bộ kiểm ĐẠT, 1 khi có KHÔNG ĐẠT, 2 khi thiếu công cụ, 3 khi máy quá tải.
set -u
REPO=$(cd "$(dirname "$0")/../.." && pwd)
cd "$REPO" || exit 2
PY=${PY:-python3}
VPY=${VPY:-/tmp/audit_venv/bin/python}
DATA=${HTKK_DATA:-/Users/tuanle/codes/HTKK}
FAIL=""
TMPD=$(mktemp -d /tmp/htkk_rebuild.XXXXXX)
export REPO PY VPY DATA TMPD
trap 'rm -rf "$TMPD"' EXIT

"$VPY" -c 'import olefile, openpyxl' 2>/dev/null || { echo "Thiếu $VPY có olefile + openpyxl"; exit 2; }

waitload() {
  i=0
  while :; do
    L=$(sysctl -n vm.loadavg 2>/dev/null | awk '{print $2}')
    [ -z "$L" ] && return 0
    if awk "BEGIN{exit !($L > 10)}"; then
      i=$((i + 1))
      [ $i -ge 10 ] && { echo "load $L > 10 sau 10 lần đo, dừng"; exit 3; }
      echo "load $L > 10, chờ 60 giây"; sleep 60
    else
      return 0
    fi
  done
}

step() {   # step <tên> <lệnh...>: chạy bước sinh; lỗi thì dừng
  name=$1; shift
  waitload
  echo "== BƯỚC $name"
  nice -n 15 "$@" || { echo "BƯỚC $name lỗi (mã $?)"; exit 1; }
}

check() {  # check <tên> <mẫu lỗi trong output hoặc -> <lệnh...>: chạy bộ kiểm; in tail khi lỗi; một dòng kết quả
  name=$1; bad=$2; shift 2
  waitload
  echo "== KIỂM $name"
  nice -n 15 "$@" > "$TMPD/out.txt" 2>&1
  rc=$?
  if [ $rc -ne 0 ] || { [ "$bad" != "-" ] && grep -E -q "$bad" "$TMPD/out.txt"; }; then
    tail -n 8 "$TMPD/out.txt"
    echo "KIỂM $name: KHÔNG ĐẠT (mã $rc)"; FAIL="$FAIL $name"
  else
    echo "KIỂM $name: ĐẠT"
  fi
}

# ---------------------------------------------------------------- bước sinh (từ dữ liệu nguồn và JSON đã có)
# T5 và T9 sinh lại vào repo, xóa JSON cũ trước để không còn tệp mồ côi sau reorg _forms.
# T4 và T6 không cần bước sinh trong repo: bộ kiểm bên dưới tái tạo ra thư mục tạm và so từng byte.
find _spec/xsd -name '*.json' -delete
step "T5 xsd_merge" "$PY" -I _spec/tools/xsd_merge.py --src "$DATA/_extracted" --out _spec/xsd \
  --forms "$REPO/_forms" --menu "$DATA/_extracted/menu.xml"
find _spec/dv -name '*.json' -delete
step "T9 biff_dv" "$VPY" -I _spec/tools/biff_dv.py _spec/dv _forms
step "T3 postprocess" "$PY" -I _spec/tools/validations/postprocess.py _spec/validations
step "T8 helpers" "$PY" -I _spec/tools/helpers.py
step "T10 render_cong_thuc" "$VPY" -I _spec/tools/render_cong_thuc.py --repo "$REPO"
step "T11 coverage" "$PY" -I _spec/tools/coverage.py --out _spec/coverage

# ---------------------------------------------------------------- bộ kiểm
check "check_mapping" "RESULT: FAIL" "$PY" -I _re/check_mapping.py
check "verify_formulas" "KHÔNG ĐẠT|LỆCH|FAIL" "$VPY" -I _spec/tools/verify_formulas.py --formulas _spec/formulas \
  --src "$DATA/_extracted" --forms _forms
check "verify_cong_thuc" "KẾT QUẢ: LỆCH" "$PY" -I _spec/tools/verify_cong_thuc.py --repo "$REPO"
check "T2 rules verify" "KHÔNG ĐẠT|FAIL" "$PY" -I _spec/tools/rules_extract/verify.py "$DATA/_decompiled" \
  _spec/rules _spec/csharp/classes.json
check "T3 validations verify" "KHÔNG ĐẠT|FAIL" "$PY" -I _spec/tools/validations/verify.py "$DATA" "$REPO/_spec"
# scan_idents.py thoát mã 1 khi có mẫu cấm, binding thiếu khởi tạo hoặc biến tự do khác 0
check "T3 scan_idents" "-" "$PY" -I _spec/tools/validations/scan_idents.py _spec/validations

# T4: config_xml.py tự kiểm vòng XML -> JSON -> XML (in "lệch kiểm vòng: N"); tái tạo ra thư mục tạm, so từng byte.
check "T4 config_xml (kiểm vòng + tái tạo)" "lệch kiểm vòng: [1-9]" sh -c '
  "$PY" -I _spec/tools/config_xml.py --src "$DATA/_extracted" --cs "$DATA/_decompiled" --out "$TMPD/config" &&
  diff -rq _spec/config "$TMPD/config"'

# T5: xsd_merge.py in số đếm nguồn/JSON từng khối; check_counts.py đối chiếu, rồi so từng byte với _spec/xsd.
check "T5 xsd_merge (số đếm + tái tạo)" "-" sh -c '
  "$PY" -I _spec/tools/xsd_merge.py --src "$DATA/_extracted" --out "$TMPD/xsd" --forms "$REPO/_forms" \
    --menu "$DATA/_extracted/menu.xml" > "$TMPD/xsd.log" &&
  "$PY" -I _spec/tools/check_counts.py xsd "$TMPD/xsd.log" &&
  diff -rq _spec/xsd "$TMPD/xsd"'

# T6: catalog.py thoát mã khác 0 khi số ô BIFF hoặc số phần tử XML lệch số đếm độc lập; rồi so từng byte với _spec/catalog.
check "T6 catalog (số ô + tái tạo)" "-" sh -c '
  "$VPY" -I _spec/tools/catalog.py --src "$DATA/_extracted" --dec "$DATA/_decompiled" --out "$TMPD/catalog" &&
  diff -rq _spec/catalog "$TMPD/catalog"'

# T9: biff_dv.py đếm bản ghi DV/DVAL; check_counts.py đối chiếu tổng với từng tệp; rồi so từng byte với _spec/dv.
check "T9 biff_dv (đếm DV + tái tạo)" "-" sh -c '
  "$VPY" -I _spec/tools/biff_dv.py "$TMPD/dv" _forms &&
  "$PY" -I _spec/tools/check_counts.py dv "$TMPD/dv" &&
  diff -rq -x README.md _spec/dv "$TMPD/dv"'

# T8: tỷ lệ phủ cả hai định nghĩa >= 95% và không còn helper thiếu overlay (đọc coverage.json vừa sinh).
check "T8 tỷ lệ phủ" "-" "$PY" -I -c '
import json, sys
c = json.load(open("_spec/helpers/coverage.json"))
new, old = c["coverage_pct"], c["old_definition"]["coverage_pct"]
h = json.load(open("_spec/helpers/helpers.json"))["helpers"]
low = [x["name"] for x in h if x["confidence"] == "thấp"]
print("định nghĩa mới %.4f%%, định nghĩa cũ %.4f%%, tập phủ %d hàm, chưa có overlay %d" % (new, old, len(h), len(low)))
sys.exit(0 if new >= 95 and old >= 95 and not low else 1)'
# T11: chạy lần hai ra thư mục tạm, so từng byte với _spec/coverage.
check "T11 tất định" "-" sh -c '
  "$PY" -I _spec/tools/coverage.py --out "$TMPD/cov2" > /dev/null &&
  diff -rq -x README.md _spec/coverage "$TMPD/cov2"'
check "T7 app_branches --accuracy" "-" "$PY" -I _spec/tools/app_branches.py --accuracy

echo "=================================================="
if [ -z "$FAIL" ]; then echo "TỔNG KẾT: mọi bộ kiểm ĐẠT"; exit 0; fi
echo "TỔNG KẾT: KHÔNG ĐẠT ở:$FAIL"; exit 1

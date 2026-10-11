#!/bin/sh
# rebuild_all.sh: chạy lại tuần tự các bước Python tái lập được từ JSON đã có, rồi chạy mọi bộ kiểm.
# Không gọi dotnet (T0 csharp, T2 rules, T3 validations raw được sinh bằng dotnet; lệnh ở _spec/README.md).
# Dùng (từ bất kỳ đâu):  sh _spec/tools/rebuild_all.sh
# Biến môi trường: PY (mặc định python3), VPY (python có olefile/openpyxl, mặc định /tmp/audit_venv/bin/python),
#                  HTKK_DATA (thư mục chứa _decompiled và _extracted, mặc định /Users/tuanle/codes/HTKK).
# Mọi bước chạy với nice -n 15, nối tiếp nhau (không song song); trước mỗi bước nặng chờ khi load 1 phút > 10.
# Thoát mã 0 khi mọi bộ kiểm ĐẠT, 1 khi có bộ kiểm lệch, 2 khi thiếu công cụ.
set -u
REPO=$(cd "$(dirname "$0")/../.." && pwd)
cd "$REPO" || exit 2
PY=${PY:-python3}
VPY=${VPY:-/tmp/audit_venv/bin/python}
DATA=${HTKK_DATA:-/Users/tuanle/codes/HTKK}
FAIL=""
TMPD=$(mktemp -d /tmp/htkk_rebuild.XXXXXX)
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

check() {  # check <tên> <mẫu lỗi trong output hoặc -> <lệnh...>: chạy bộ kiểm, ghi KẾT QUẢ
  name=$1; bad=$2; shift 2
  waitload
  echo "== KIỂM $name"
  nice -n 15 "$@" > "$TMPD/out.txt" 2>&1
  rc=$?
  tail -n 6 "$TMPD/out.txt"
  if [ $rc -ne 0 ] || { [ "$bad" != "-" ] && grep -E -q "$bad" "$TMPD/out.txt"; }; then
    echo "KIỂM $name: LỆCH (mã $rc)"; FAIL="$FAIL $name"
  else
    echo "KIỂM $name: ĐẠT"
  fi
}

# ---------------------------------------------------------------- bước sinh (từ JSON đã có)
step "T3 postprocess" "$PY" -I _spec/tools/validations/postprocess.py _spec/validations
step "T8 helpers" "$PY" -I _spec/tools/helpers.py
step "T10 render_cong_thuc" "$VPY" -I _spec/tools/render_cong_thuc.py --repo "$REPO"
step "T11 coverage" "$PY" -I _spec/tools/coverage.py --out _spec/coverage

# ---------------------------------------------------------------- bộ kiểm
check "check_mapping" "RESULT: FAIL" "$PY" -I _re/check_mapping.py
check "verify_formulas" "KHÔNG ĐẠT|LỆCH|FAIL" "$VPY" -I _spec/tools/verify_formulas.py --formulas _spec/formulas \
  --src "$DATA/_extracted" --forms _forms
check "verify_cong_thuc" "KẾT QUẢ: LỆCH" "$PY" -I _spec/tools/verify_cong_thuc.py --repo "$REPO"
check "T2 rules_extract/verify" "KHÔNG ĐẠT|FAIL" "$PY" -I _spec/tools/rules_extract/verify.py "$DATA/_decompiled" \
  _spec/rules _spec/csharp/classes.json
check "T3 validations/verify" "KHÔNG ĐẠT|FAIL" "$PY" -I _spec/tools/validations/verify.py "$DATA" "$REPO/_spec"
# scan_idents: mọi dòng đếm phải bằng 0
waitload
echo "== KIỂM T3 scan_idents"
nice -n 15 "$PY" -I _spec/tools/validations/scan_idents.py _spec/validations > "$TMPD/scan.txt" 2>&1
NZ=$(awk '!/^tổng/ && $NF ~ /^[0-9]+$/ && $NF > 0' "$TMPD/scan.txt" | wc -l | tr -d ' ')
if [ "$NZ" = "0" ]; then echo "KIỂM T3 scan_idents: ĐẠT (mọi mẫu 0)"; else
  awk '!/^tổng/ && $NF ~ /^[0-9]+$/ && $NF > 0' "$TMPD/scan.txt"; echo "KIỂM T3 scan_idents: LỆCH"; FAIL="$FAIL scan_idents"; fi
# T8: tỷ lệ phủ cả hai định nghĩa >= 95% (đọc lại coverage.json vừa sinh)
echo "== KIỂM T8 tỷ lệ phủ"
"$PY" -I -c '
import json, sys
c = json.load(open("_spec/helpers/coverage.json"))
new, old = c["coverage_pct"], c["old_definition"]["coverage_pct"]
h = json.load(open("_spec/helpers/helpers.json"))["helpers"]
low = [x["name"] for x in h if x["confidence"] == "thấp"]
print("định nghĩa mới %.4f%%, định nghĩa cũ %.4f%%, tập phủ %d hàm, chưa có overlay %d" % (new, old, len(h), len(low)))
sys.exit(0 if new >= 95 and old >= 95 and not low else 1)' && echo "KIỂM T8: ĐẠT" || { echo "KIỂM T8: LỆCH"; FAIL="$FAIL T8"; }
# T11: chạy lần hai ra thư mục tạm, so từng byte
waitload
echo "== KIỂM T11 tất định"
nice -n 15 "$PY" -I _spec/tools/coverage.py --out "$TMPD/cov2" > /dev/null 2>&1
if diff -rq -x README.md _spec/coverage "$TMPD/cov2" > "$TMPD/diff.txt"; then echo "KIỂM T11 tất định: ĐẠT (diff -rq rỗng)"; else
  head -5 "$TMPD/diff.txt"; echo "KIỂM T11 tất định: LỆCH"; FAIL="$FAIL T11_det"; fi
check "T7 app_branches --accuracy" "-" "$PY" -I _spec/tools/app_branches.py --accuracy

echo "=================================================="
if [ -z "$FAIL" ]; then echo "TỔNG KẾT: mọi bộ kiểm ĐẠT"; exit 0; fi
echo "TỔNG KẾT: LỆCH ở:$FAIL"; exit 1

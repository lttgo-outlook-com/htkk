#!/usr/bin/env bash
# Đo coordinator chọn đúng agent bao nhiêu phần trăm, không tốn token của agent thực thi.
# Chạy từ thư mục nào cũng được: script tự chuyển về gốc dự án (cha của .claude/tests). Cần: claude, python3, uuidgen. Mỗi ca chạy một phiên không tương tác ở chế độ chạy thử:
# hook delegation_guard.py ghi lựa chọn rồi chặn lệnh.
# Mỗi ca có session id cố định (claude --session-id); manifest .claude/routing-manifest-$i.jsonl ghi cặp ca và session id
# để routing_report.py tìm transcript thật ~/.claude/projects/*/<session_id>.jsonl.
# RUNS: số lần lặp toàn bộ bộ ca (mặc định 3, số nguyên từ 1 đến 20); mỗi lần ghi .claude/routing-run-$i.jsonl.
# Mỗi dòng manifest có thêm trường "exit" là mã thoát của claude; routing_report.py chấm SAI ca có exit khác 0.
# Khi ROUTING_ONLY không đặt hoặc rỗng, mọi lần gọi routing_report.py nhận cờ --expect-all: ca vắng trong manifest bị chấm SAI (vắng_trong_manifest).
# ROUTING_ONLY: danh sách mã ca cách nhau dấu phẩy (ví dụ c01,c09) để chạy tập con (mặc định chạy tất cả); mã không tồn tại thì thoát mã 2.
# Kiểm tra khói dùng ca đầu tiên (theo thứ tự trong routing_cases.json) của tập chọn có kỳ vọng lệnh Agent; không có ca nào thì bỏ qua kiểm tra khói.
# Mã thoát của script là mã của routing_report.py --runs (0 khi mọi ca đã chấm đúng, 1 khi có ca SAI); lỗi cấu hình thoát mã 2.
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
if [ ! -f "$ROOT/.claude/agents/coordinator.md" ]; then
  echo "Lỗi: không thấy $ROOT/.claude/agents/coordinator.md; script phải nằm ở <gốc dự án>/.claude/tests/." >&2
  exit 2
fi
cd "$ROOT"
RUNS="${RUNS:-3}"
if ! [[ "$RUNS" =~ ^([1-9]|1[0-9]|20)$ ]]; then
  echo "Lỗi: RUNS phải là số nguyên từ 1 đến 20 (nhận được: '$RUNS')." >&2
  exit 2
fi
# macOS không có timeout; dùng gtimeout (brew coreutils) nếu có, không thì chạy không giới hạn thời gian.
TO=$(command -v timeout || command -v gtimeout || true)
TSV=$(mktemp)
trap 'rm -f "$TSV" "$TSV.rest" "$TSV.smoke"' EXIT
SMOKE_ID=""
# Cột TSV: mã ca, prompt. Dòng đầu của $TSV.smoke: mã ca chạy kiểm tra khói (rỗng nếu bỏ qua).
if ! python3 -I - "$TSV" "${ROUTING_ONLY:-}" <<'PY'
import json, sys
tsv, only = sys.argv[1], sys.argv[2]
cases = json.load(open(".claude/tests/routing_cases.json", encoding="utf-8"))
if only.strip():
    want = [x.strip() for x in only.split(",")]
    known = {c["id"] for c in cases}
    bad = [x for x in want if x not in known]
    if bad:
        print("Lỗi: ROUTING_ONLY có mã ca không tồn tại: " + ", ".join(repr(x) for x in bad)
              + " (có: " + ", ".join(sorted(known)) + ").", file=sys.stderr)
        sys.exit(2)
    cases = [c for c in cases if c["id"] in want]
with open(tsv, "w", encoding="utf-8") as f:
    for c in cases:
        f.write(c["id"] + "\t" + c["prompt"].replace("\t", " ").replace("\n", " ") + "\n")
smoke = next((c["id"] for c in cases if c["expect"]), "")
with open(tsv + ".smoke", "w", encoding="utf-8") as f:
    f.write(smoke + "\n")
PY
then
  exit 2
fi
SMOKE_ID=$(head -n 1 "$TSV.smoke")
if [ -z "$SMOKE_ID" ]; then
  echo "Bỏ qua kiểm tra khói: tập ca chọn không có ca nào kỳ vọng lệnh Agent." >&2
fi
LAST_SID=""
run_case() { # $1 = nhật ký, $2 = mã ca, $3 = prompt, $4 = manifest; stdin lấy từ /dev/null để claude không đọc phần còn lại của TSV
  LAST_SID=$(uuidgen | tr 'A-Z' 'a-z')
  local rc=0
  # prompt đứng trước các cờ để cờ nhận danh sách không nuốt prompt; giữ mã thoát, không dừng script
  DELEGATION_DRYRUN=1 DELEGATION_LOG="$1" DELEGATION_CASE="$2" \
    ${TO:+$TO 300} claude -p "$3" --agent coordinator --session-id "$LAST_SID" > /dev/null 2>&1 < /dev/null || rc=$?
  printf '{"case": "%s", "session_id": "%s", "ts": "%s", "exit": %d}\n' "$2" "$LAST_SID" "$(date -u +%Y-%m-%dT%H:%M:%SZ)" "$rc" >> "$4"
}
# Chạy đủ bộ ca (ROUTING_ONLY không đặt hoặc chỉ có khoảng trắng): truyền --expect-all cho routing_report.py.
EXPECT=()
if [ -z "${ROUTING_ONLY//[[:space:]]/}" ]; then
  EXPECT=(--expect-all)
fi
LOGS=()
for ((i = 1; i <= RUNS; i++)); do
  LOG=".claude/routing-run-$i.jsonl"
  MAN=".claude/routing-manifest-$i.jsonl"
  : > "$LOG"
  : > "$MAN"
  LOGS+=("$LOG")
  if [ "$i" -eq 1 ] && [ -n "$SMOKE_ID" ]; then
    # Kiểm tra khói: chạy ca đầu tiên có kỳ vọng lệnh Agent trước, phân biệt claude không chạy, hook không chạy, định tuyến sai.
    IFS=$'\t' read -r id prompt < <(awk -F '\t' -v id="$SMOKE_ID" '$1 == id' "$TSV")
    run_case "$LOG" "$id" "$prompt" "$MAN"
    TR=$(compgen -G "$HOME/.claude/projects/*/$LAST_SID.jsonl" | head -n 1 || true)
    if [ -z "$TR" ]; then
      echo "Kiểm tra khói thất bại: không thấy transcript của $id (session $LAST_SID); claude không chạy được (kiểm đăng nhập, quyền)." >&2
      exit 1
    fi
    if ! python3 -I - "$TR" <<'PY'
import json, sys
for line in open(sys.argv[1], encoding="utf-8", errors="replace"):
    try:
        o = json.loads(line)
    except ValueError:
        continue
    if not isinstance(o, dict) or o.get("isSidechain") or o.get("type") != "assistant":
        continue
    m = o.get("message")
    c = m.get("content") if isinstance(m, dict) else None
    if isinstance(c, list) and any(isinstance(b, dict) and b.get("type") == "tool_use" and b.get("name") in ("Agent", "Task") for b in c):
        sys.exit(0)
sys.exit(1)
PY
    then
      echo "Kiểm tra khói thất bại: $id không có lệnh Agent: coordinator định tuyến sai." >&2
      exit 1
    fi
    if ! python3 -I - "$LOG" "$id" <<'PY'
import json, sys
for line in open(sys.argv[1], encoding="utf-8", errors="replace"):
    try:
        o = json.loads(line)
    except ValueError:
        continue
    if isinstance(o, dict) and o.get("event") == "dryrun" and o.get("case") == sys.argv[2]:
        sys.exit(0)
sys.exit(1)
PY
    then
      echo "Kiểm tra khói thất bại: $id có lệnh Agent nhưng hook không chạy (hook frontmatter cần workspace trust: mở claude tương tác một lần trong thư mục này và chấp nhận trust; claude -p không tính)." >&2
      exit 1
    fi
    awk -F '\t' -v id="$SMOKE_ID" '$1 != id' "$TSV" > "$TSV.rest"
    SRC="$TSV.rest"
  else
    SRC="$TSV"
  fi
  while IFS=$'\t' read -r id prompt; do
    run_case "$LOG" "$id" "$prompt" "$MAN"
  done < "$SRC"
  rm -f "$TSV.rest"
done
RC=0
python3 -I .claude/tests/routing_report.py --runs "${LOGS[@]+"${LOGS[@]}"}" "${EXPECT[@]+"${EXPECT[@]}"}" || RC=$?
exit "$RC"

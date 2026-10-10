#!/bin/sh
# Trích logic kiểm tra dữ liệu HTKK 5.7.9 -> _spec/validations/
# Dùng: sh _spec/tools/validations/run.sh <repo_root_chứa__decompiled_forms_extracted> <spec_root> [out_dir] [python]
# Dựng CSharpCompilation trong bộ nhớ theo lô (một class + tệp utilities nó tham chiếu, tối đa 50 tệp) để lấy ControlFlowGraph;
# không Emit, không build/run mã trong _decompiled. Giới hạn tài nguyên: GC workstation, heap 4 GB, 1 CPU khi build, nice 15.
set -eu
DATA="$1"
SPEC="$2"
OUT="${3:-$SPEC/validations}"
PY="${4:-python3}"
export DOTNET_gcServer=0 DOTNET_GCHeapHardLimit=0x100000000 DOTNET_TieredPGO=0 MSBUILDDISABLENODEREUSE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_NOLOGO=1
# chờ máy rảnh: load 1 phút > 10 thì chờ 60 giây, tối đa 10 lần
i=0
while :; do
  L=$(sysctl -n vm.loadavg | awk '{print $2}')
  if awk "BEGIN{exit !($L > 10)}"; then
    i=$((i+1)); [ $i -ge 10 ] && { echo "load $L > 10 sau 10 lần đo, dừng"; exit 3; }; sleep 60
  else break; fi
done
TMP="$(mktemp -d /tmp/htkk_val.XXXXXX)"
nice -n 15 dotnet build "$SPEC/tools/roslyn_validations/RoslynValidations.csproj" -c Release --no-restore -maxcpucount:1 \
  -p:UseSharedCompilation=false --disable-build-servers -o "$TMP/bin" >/dev/null
dotnet build-server shutdown >/dev/null 2>&1 || true
nice -n 15 /usr/bin/time -l dotnet "$TMP/bin/RoslynValidations.dll" "$DATA/_decompiled" "$TMP/raw.json" --log "$TMP/batches.tsv"
nice -n 15 "$PY" -I "$SPEC/tools/validations/build.py" "$TMP/raw.json" "$DATA" "$OUT" "$SPEC/csharp/classes.json"
nice -n 15 "$PY" -I "$SPEC/tools/validations/postprocess.py" "$OUT"
rm -rf "$TMP"

#!/bin/sh
# Trích logic kiểm tra dữ liệu HTKK 5.7.9 -> _spec/validations/
# Dùng: sh _spec/tools/validations/run.sh <repo_root_chứa__decompiled_forms_extracted> <spec_root> [out_dir]
# Chỉ parse mã decompile (Roslyn ParseText); không build/run mã trong _decompiled. Python chạy với -I.
set -eu
DATA="$1"
SPEC="$2"
OUT="${3:-$SPEC/validations}"
TMP="$(mktemp -d /tmp/htkk_val.XXXXXX)"
dotnet build -c Release -o "$TMP/bin" "$SPEC/tools/roslyn_validations/RoslynValidations.csproj" >/dev/null
dotnet "$TMP/bin/RoslynValidations.dll" "$DATA/_decompiled" "$TMP/raw.json"
/tmp/audit_venv/bin/python -I "$SPEC/tools/validations/build.py" "$TMP/raw.json" "$DATA" "$OUT" "$SPEC/csharp/classes.json"
rm -rf "$TMP"

#!/bin/bash
# Decompile HTKK managed assemblies with ILSpy CLI (ilspycmd) in .NET SDK container.
# /src  = extracted cabinet contents (read-only)
# /out  = decompiled C# output
set -uo pipefail

echo "=== install ilspycmd (pinned) ==="
dotnet tool install -g ilspycmd --version 11.1.0.9782 2>&1 | tail -2
export PATH="$PATH:/root/.dotnet/tools"
if ! command -v ilspycmd >/dev/null; then
  echo "11.x failed, retry 9.1.0.7988"
  dotnet tool install -g ilspycmd --version 9.1.0.7988 2>&1 | tail -2
fi
ilspycmd --version

for a in htkk.exe tax.htkk.business.dll tax.htkk.utilities.dll tax.htkk.computehash.dll; do
  base="${a%.*}"
  out="/out/$base"
  echo "=== $a -> $out ==="
  mkdir -p "$out"
  if ilspycmd "/src/$a" -p -o "$out" >"$out/.ilspy.log" 2>&1; then
    ncs=$(find "$out" -name '*.cs' | wc -l | tr -d ' ')
    echo "OK $a: $ncs .cs files"
  else
    echo "FAIL $a:"
    tail -3 "$out/.ilspy.log"
  fi
done

echo "=== third-party spot check (expect all MANAGED too, we just skip them) ==="
for a in farpoint.win.dll newtonsoft.json.dll sharpsvn.dll vnconvert.dll; do
  echo -n "$a: "
  ilspycmd "/src/$a" -t > /dev/null 2>&1 && echo "decompiles (managed)" || echo "FAILS (native?) $(ilspycmd "/src/$a" -t 2>&1 | tail -1)"
done

#!/usr/bin/env python3
"""Ghi thân hàm (bỏ dòng trống, comment IL, khối Log) của các hàm confidence thấp hoặc một phần ra /tmp để đọc tay."""
import json
import os
import re
import sys
WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
H = json.load(open(WT + "/_spec/helpers/helpers.json"))["helpers"]
extra = {"GeneralFunctions.InsertNumberCell", "GeneralFunctions.SetLaiDongCon", "GeneralFunctions.GetHanNopToKhai",
         "GeneralFunctions.ThayDoiTinhTheoCQT2Cap", "GeneralFunctions.checkLockUnLockCells",
         "GeneralFunctions.CopyTableToTable2"}
sel = [h for h in H if h["confidence"] == "thấp" or h["name"] in extra]
sel.sort(key=lambda h: -h["call_count"]["total"])
print(len(sel), file=sys.stderr)
out = open("/tmp/low_bodies.txt", "w")
for h in sel:
    L = open("/Users/tuanle/codes/HTKK/" + h["source"]["file"], encoding="utf-8", errors="replace").read().split("\n")
    out.write("=" * 8 + " %s (%d)\n" % (h["name"], h["call_count"]["total"]))
    for s, e in h["source"]["lines"]:
        out.write("-- %d-%d\n" % (s, e))
        for i in range(s - 1, e):
            t = L[i].strip()
            if not t or t.startswith("//IL") or t in ("{", "}", "try", "catch (Exception ex)") or "HTKKFile.Log(" in t:
                continue
            out.write("%d %s\n" % (i + 1, t[:230]))
out.close()

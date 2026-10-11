#!/usr/bin/env python3
"""In thân hàm (đã lọc dòng trống/comment/Log) theo tên: show_names.py Class.Method ..."""
import json
import os
import sys
WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
H = {h["name"]: h for h in json.load(open(WT + "/_spec/helpers/helpers.json"))["helpers"]}
for n in sys.argv[1:]:
    h = H[n]
    print("=====", n, h["overloads"][0]["params"], "->", h["overloads"][0]["return_type"])
    L = open("/Users/tuanle/codes/HTKK/" + h["source"]["file"], encoding="utf-8", errors="replace").read().split("\n")
    for s, e in h["source"]["lines"]:
        print("--", s, e)
        for i in range(s - 1, e):
            t = L[i].strip()
            if t and t not in ("{", "}", "try") and not t.startswith("//") and "Log(" not in t:
                print(i + 1, t[:200])

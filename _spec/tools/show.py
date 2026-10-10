#!/usr/bin/env python3
"""In thân hàm từ index để đọc tay. Dùng: show.py <từ> <đến> (chỉ số trong xếp hạng phủ) hoặc show.py name1 name2"""
import json
import sys
import os
WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
H = json.load(open(WT + "/_spec/helpers/helpers.json"))["helpers"]
H.sort(key=lambda h: -h["call_count"]["total"])
a = sys.argv[1:]
if len(a) == 2 and all(x.isdigit() for x in a):
    sel = H[int(a[0]):int(a[1])]
else:
    sel = [h for h in H if h["name"] in a]
for h in sel:
    print("=" * 20, h["name"], h["call_count"], h["overloads"][0]["params"], "->", h["overloads"][0]["return_type"],
          "overloads:", len(h["overloads"]))
    L = open("/Users/tuanle/codes/HTKK/" + h["source"]["file"], encoding="utf-8", errors="replace").read().split("\n")
    for s, e in h["source"]["lines"][:3]:
        print("--", s, e)
        for i in range(s - 1, min(e, s + 40)):
            print(i + 1, L[i].rstrip()[:150])

#!/usr/bin/env python3
"""Chọn 15 hàm ngẫu nhiên trong tập phủ (seed cố định) để đọc tay."""
import json
import os
import random
SEED = 20261011
WT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
H = json.load(open(WT + "/_spec/helpers/helpers.json"))["helpers"]
names = sorted(h["name"] for h in H)
pick = random.Random(SEED).sample(names, 15)
by = {h["name"]: h for h in H}
for n in pick:
    h = by[n]
    print("=" * 10, n, "| conf:", h["confidence"], "| count:", h["call_count"]["total"], "| lines:", h["source"]["lines"])
    print("PURPOSE:", h["purpose"])
    print("EFFECTS:", h["semantics"]["effects"][:300], "| RET:", h["semantics"]["returns"][:200])
    print("CALLS:", h["calls"], "| SE:", h["side_effects"])
    L = open("/Users/tuanle/codes/HTKK/" + h["source"]["file"], encoding="utf-8", errors="replace").read().split("\n")
    for s, e in h["source"]["lines"][:2]:
        for i in range(s - 1, min(e, s + 30)):
            t = L[i].strip()
            if t and not t.startswith("//") and t not in ("{", "}", "try", "catch (Exception ex)"):
                print("  ", i + 1, t[:130])

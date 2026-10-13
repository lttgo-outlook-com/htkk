#!/usr/bin/env python3
"""Tính tỷ lệ định tuyến đúng từ nhật ký chạy thử.
Dùng: routing_report.py <nhat-ky.jsonl> [--manifest <manifest.jsonl>] [--expect-all]
      routing_report.py --runs <nhat-ky-1.jsonl> <nhat-ky-2.jsonl> ... [--expect-all]
Manifest mặc định suy từ tên nhật ký: routing-run-N.jsonl -> routing-manifest-N.jsonl (cùng thư mục).
Không có manifest: mọi ca SAI (không_có_manifest).

Cách chấm mỗi ca (mọi kiểm đếm và same_message chỉ áp lên message đầu = message có bản ghi hook đầu tiên trong nhật ký):
- manifest: mỗi dòng có case, session_id, exit (mã thoát của claude); không có session id của ca thì SAI (không_có_manifest, không in ghi chú khác);
  thiếu trường exit thì SAI (manifest_thiếu_exit); exit khác 0 thì SAI (claude_lỗi_exit=N);
- tập ca con: khi manifest có tệp và có dòng hợp lệ, chỉ chấm các ca có trong manifest (ca vắng bị bỏ qua, không SAI); số ca đã chấm được in ra.
  Manifest không có tệp hoặc rỗng: mọi ca SAI (không_có_manifest);
- --expect-all (chạy đủ bộ ca, không đặt ROUTING_ONLY): mọi ca trong routing_cases.json mà vắng trong manifest bị chấm SAI
  (vắng_trong_manifest) và tính vào tổng; không có cờ thì giữ hành vi tập con ở trên;
- mã thoát của script: 0 khi mọi ca đã chấm ĐÚNG, 1 khi có ca SAI hoặc không chấm được ca nào (cả một tệp lẫn --runs);
- transcript tìm theo session id của ca trong manifest: ~/.claude/projects/*/<session_id>.jsonl; không thấy thì SAI (không_có_transcript);
- mọi ca (cả c12) cần ít nhất một message assistant hợp lệ: type assistant, không isApiErrorMessage true, content có khối text (không rỗng)
  hoặc tool_use; không có thì SAI (không_có_phản_hồi_hợp_lệ);
- số khối tool_use tên Agent hoặc Task trong transcript phải bằng số bản ghi hook của ca; lệch thì SAI (log_lệch_transcript);
- mỗi bản ghi hook: session_id phải bằng session id của manifest (session_lệch); tool_use_id phải có trong transcript
  (tool_use_id_không_có_trong_transcript); tên agent lấy từ input.subagent_type của khối tool_use trong transcript và phải bằng agent của
  bản ghi hook (lệch_agent); khóa message là message.id của message chứa khối đó, không tra được thì "không xác định" và SAI
  (trường transcript_tail của hook chỉ để ghi, không dùng để chấm);
- ĐÚNG khi mọi lệnh của message đầu thuộc "expect" (ca expect rỗng, tức c12: ĐÚNG chỉ khi có message assistant hợp lệ với text không rỗng,
  0 khối tool_use bất kỳ tên nào trong transcript và 0 bản ghi hook);
- min_agents, max_agents: số lệnh của message đầu;
- same_message: message đầu chứa đủ số lệnh yêu cầu (min_agents nếu có, ngược lại từ 2); in "cùng_message"; số_message và gọi_lại (số message
  sau message đầu) chỉ để thông tin;
- all_same_agent: các agent của message đầu chỉ có một loại;
- prompt_khớp (mọi ca, cả c12): các user message thật của transcript (bỏ isMeta và dòng bắt đầu bằng "<" hoặc "Caveat:") phải có đúng một
  và bằng prompt của ca (bỏ khoảng trắng đầu cuối); lệch thì SAI (lẫn_prompt); từ hai trở lên thì SAI (nhiều_user_message);
  không có thì SAI (không_kiểm_được_prompt).
"""
import glob
import json
import os
import re
import sys
import unicodedata

here = os.path.dirname(os.path.abspath(__file__))
CASES = {c["id"]: c for c in json.load(open(os.path.join(here, "routing_cases.json"), encoding="utf-8"))}
UNKNOWN = "không xác định"
AGENT_TOOLS = ("Agent", "Task")
SID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")


def nfc(s):
    return unicodedata.normalize("NFC", s or "").strip()


def find_transcript(sid):
    """Tìm ~/.claude/projects/*/<sid>.jsonl; chỉ nhận session id dạng uuid chữ thường."""
    if not isinstance(sid, str) or not SID_RE.match(sid):
        return None
    hits = sorted(glob.glob(os.path.join(glob.escape(os.path.expanduser("~/.claude/projects")), "*", sid + ".jsonl")))
    return hits[0] if hits else None


def real_prompt(s):
    return isinstance(s, str) and not s.lstrip().startswith(("<", "Caveat:"))


def load_transcript(path):
    """Trả về dict (ids: tool_use_id -> (message.id, subagent_type) của lệnh Agent/Task; n_agent: số lệnh Agent/Task;
    users: danh sách user message thật (chuỗi, hoặc list chỉ có khối text, không meta); has_reply: có message assistant hợp lệ;
    has_text: có message assistant hợp lệ với text không rỗng; n_tool_any: số khối tool_use mọi tên) hoặc None khi không đọc được."""
    try:
        ids, n_agent, users = {}, 0, []
        has_reply = has_text = False
        any_seen = set()
        agent_seen = set()
        with open(os.path.expanduser(path), encoding="utf-8", errors="replace") as f:
            for line in f:
                try:
                    o = json.loads(line)
                except ValueError:
                    continue
                if not isinstance(o, dict) or o.get("isSidechain"):
                    continue
                m = o.get("message") if isinstance(o.get("message"), dict) else {}
                c = m.get("content")
                if o.get("type") == "user" and m.get("role") == "user" and not o.get("isMeta"):
                    if isinstance(c, str):
                        if real_prompt(c):
                            users.append(c)
                    elif isinstance(c, list):
                        texts = [b.get("text", "") for b in c if isinstance(b, dict) and b.get("type") == "text"]
                        if texts and not any(isinstance(b, dict) and b.get("type") == "tool_result" for b in c):
                            joined = "\n".join(t for t in texts if isinstance(t, str))
                            if real_prompt(joined):
                                users.append(joined)
                if o.get("type") == "assistant" and isinstance(c, list):
                    valid = o.get("isApiErrorMessage") is not True
                    for b in c:
                        if not isinstance(b, dict):
                            continue
                        if valid and b.get("type") == "text" and isinstance(b.get("text"), str) and b["text"].strip():
                            has_reply = has_text = True
                        if b.get("type") != "tool_use":
                            continue
                        if valid:
                            has_reply = True
                        key = b.get("id") or id(b)
                        if key not in any_seen:
                            any_seen.add(key)
                        if b.get("name") in AGENT_TOOLS:
                            if key in agent_seen:
                                continue
                            agent_seen.add(key)
                            n_agent += 1
                            if b.get("id") and m.get("id"):
                                inp = b.get("input") if isinstance(b.get("input"), dict) else {}
                                ids.setdefault(b["id"], (str(m["id"]), inp.get("subagent_type")))
        return {"ids": ids, "n_agent": n_agent, "users": users, "has_reply": has_reply, "has_text": has_text,
                "n_tool_any": len(any_seen)}
    except Exception:
        return None


def read_manifest(path):
    """Trả về {case: (session_id, exit, có trường exit)} (dòng sau thắng) hoặc None khi không có tệp hoặc không có dòng hợp lệ."""
    if not path or not os.path.exists(path):
        return None
    sids = {}
    with open(path, encoding="utf-8", errors="replace") as f:
        for line in f:
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if isinstance(r, dict) and r.get("case") in CASES and isinstance(r.get("session_id"), str):
                sids[r["case"]] = (r["session_id"], r.get("exit"), "exit" in r)
    return sids or None


def manifest_for(log):
    d, b = os.path.split(log)
    m = re.match(r"^routing-run(-\d+)?\.jsonl$", b)
    return os.path.join(d, "routing-manifest" + m.group(1) + ".jsonl") if m and m.group(1) else (
        os.path.join(d, "routing-manifest.jsonl") if m else log + ".manifest")


def read_log(log):
    recs = []
    if os.path.exists(log):
        for line in open(log, encoding="utf-8"):
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if isinstance(r, dict) and r.get("case") in CASES and r.get("event") == "dryrun":
                recs.append(r)
    return recs


def score_log(log, manifest=None, expect_all=False):
    """Chấm một nhật ký. Trả về (danh sách (cid, đạt, dòng in, lệnh không hợp lệ), số đạt).
    Manifest có tệp và có dòng hợp lệ: chỉ chấm các ca có trong manifest (hỗ trợ tập ca con), trừ khi expect_all: ca vắng bị SAI.
    Manifest thiếu: chấm mọi ca, đều SAI."""
    recs = read_log(log)
    sids = read_manifest(manifest if manifest else manifest_for(log))
    cache = {}

    def tr(path):
        if path not in cache:
            cache[path] = load_transcript(path)
        return cache[path]

    by_case = {k: [] for k in CASES}
    for r in recs:
        by_case[r["case"]].append(r)
    results, ok = [], 0
    for cid, c in CASES.items():
        if sids is not None and cid not in sids and not expect_all:
            continue
        rs = by_case[cid]
        exp = c["expect"]
        notes, good = [], True
        sid, rc, has_exit = sids.get(cid, (None, None, False)) if sids is not None else (None, None, False)
        path = find_transcript(sid) if sid else None
        t = tr(path) if path else None
        keys, order, first_agents = [], [], []
        if not sid:
            notes.append("vắng_trong_manifest" if sids is not None else "không_có_manifest")
            good = False
        else:
            if not has_exit:
                notes.append("manifest_thiếu_exit")
                good = False
            elif isinstance(rc, bool) or not isinstance(rc, int) or rc != 0:
                notes.append(f"claude_lỗi_exit={rc}")
                good = False
            if not path:
                notes.append("không_có_transcript")
                good = False
            elif t is None or not t["users"]:
                notes.append("không_kiểm_được_prompt")
                good = False
            else:
                if len(t["users"]) > 1:
                    notes.append(f"nhiều_user_message({len(t['users'])})")
                    good = False
                elif nfc(t["users"][0]) != nfc(c["prompt"]):
                    notes.append("lẫn_prompt")
                    good = False
                if not t["has_reply"]:
                    notes.append("không_có_phản_hồi_hợp_lệ")
                    good = False
                if t["n_agent"] != len(rs):
                    notes.append(f"log_lệch_transcript(hook={len(rs)},tool_use={t['n_agent']})")
                    good = False
                # khóa message và agent của từng lệnh lấy từ transcript, đối chiếu với bản ghi hook
                bad = set()
                agents = []
                for r in rs:
                    hit = t["ids"].get(r.get("tool_use_id"))
                    agent = r.get("agent")
                    if r.get("session_id") != sid:
                        bad.add("session_lệch")
                    if hit is None:
                        bad.add("tool_use_id_không_có_trong_transcript")
                        key = UNKNOWN
                    else:
                        key = hit[0]
                        if hit[1] != agent:
                            bad.add("lệch_agent")
                        agent = hit[1]
                    keys.append(key)
                    agents.append(agent)
                    if key not in order:
                        order.append(key)
                if bad:
                    notes.extend(sorted(bad))
                    good = False
                first = order[0] if order else None
                first_agents = [a for a, k in zip(agents, keys) if k == first]
        n_msg = len(order)
        recall = max(n_msg - 1, 0)
        invalid = sum(1 for r in rs if not r.get("valid"))
        unknown = UNKNOWN in keys
        if not exp:
            # c12: cần text trả lời không rỗng, 0 tool_use mọi tên trong transcript, 0 bản ghi hook
            if sid and (rs or (t is not None and t["n_tool_any"] > 0)):
                notes.append("không_được_giao_việc")
                good = False
            if sid and t is not None and not t["has_text"]:
                notes.append("không_có_text_trả_lời")
                good = False
        elif sid:
            if not (bool(first_agents) and all(a in exp for a in first_agents)):
                good = False
            if unknown:
                notes.append("không_xác_định")
                good = False
        if sid and "min_agents" in c:
            notes.append(f"số_lệnh>={c['min_agents']}")
            good = good and len(first_agents) >= c["min_agents"]
        if sid and "max_agents" in c:
            notes.append(f"số_lệnh<={c['max_agents']}")
            good = good and len(first_agents) <= c["max_agents"]
        if sid and c.get("same_message"):
            need = c.get("min_agents", 2)
            same = len(first_agents) >= need and not unknown
            notes.append("cùng_message" if same else f"không_cùng_message(message_đầu={len(first_agents)},cần>={need})")
            good = good and same
        if sid and c.get("all_same_agent"):
            one = len(set(first_agents)) == 1
            notes.append("một_loại_agent" + ("" if one else f"(thực tế={sorted(set(first_agents)) or 'không có'})"))
            good = good and one
        if invalid and good:
            notes.append(f"CẢNH_BÁO_lệnh_không_hợp_lệ={invalid}")
        extra = " kỳ_vọng_thêm=" + " ".join(notes) if notes else ""
        ok += bool(good)
        line = (f"{cid} {'ĐÚNG' if good else 'SAI '} mong đợi={exp or 'không giao việc'} "
                f"thực tế={first_agents or 'không giao việc'} lệnh_không_hợp_lệ={invalid} "
                f"số_message={n_msg} gọi_lại={recall}{extra}")
        results.append((cid, bool(good), line, invalid))
    return results, ok


def invalid_summary(results):
    bad = [(cid, n) for cid, _, _, n in results if n]
    names = ", ".join(f"{cid}={n}" for cid, n in bad) or "không có"
    return f"Lệnh không hợp lệ: {sum(n for _, n in bad)} lệnh ở {len(bad)} ca ({names})"


def report(log, manifest=None, expect_all=False):
    results, ok = score_log(log, manifest, expect_all)
    for _, _, line, _ in results:
        print(line)
    n = len(results)
    sub = f" (tập con: đã chấm {n}/{len(CASES)} ca)" if n != len(CASES) else ""
    print(f"Tỷ lệ định tuyến đúng: {ok}/{n}{sub}")
    print(invalid_summary(results))
    return results, ok


def main(argv):
    expect_all = "--expect-all" in argv
    argv = [a for a in argv if a != "--expect-all"]
    if argv and argv[0] == "--runs":
        logs = argv[1:]
        if not logs:
            print("Thiếu tệp nhật ký sau --runs", file=sys.stderr)
            return 1
        passed = {k: 0 for k in CASES}
        scored = {k: 0 for k in CASES}
        total_ok = total_scored = 0
        total_inv = total_cases = 0
        for log in logs:
            print(f"--- {log}")
            results, ok = report(log, None, expect_all)
            total_ok += ok
            total_scored += len(results)
            for cid, good, _, inv in results:
                passed[cid] += good
                scored[cid] += 1
                total_inv += inv
                total_cases += bool(inv)
        n = len(logs)
        print("--- Tổng hợp")
        for cid in CASES:
            if scored[cid]:
                print(f"{cid} đạt ở {passed[cid]}/{scored[cid]} lần")
        print(f"Tổng: {total_ok}/{total_scored} lượt chấm đúng qua {n} lần "
              f"(đã chấm {total_scored} lượt; đủ bộ ca sẽ là {len(CASES) * n})")
        print(f"Lệnh không hợp lệ (tổng chung): {total_inv} lệnh ở {total_cases} lượt ca qua {n} lần")
        return 0 if total_scored and total_ok == total_scored else 1
    manifest, rest = None, []
    it = iter(argv)
    for a in it:
        if a == "--manifest":
            manifest = next(it, None)
            if not manifest:
                print("Thiếu tệp sau --manifest", file=sys.stderr)
                return 1
        else:
            rest.append(a)
    results, ok = report(rest[0] if rest else ".claude/routing-run.jsonl", manifest, expect_all)
    return 0 if results and ok == len(results) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

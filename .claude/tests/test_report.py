#!/usr/bin/env python3
"""Bộ kiểm offline cho routing_report.py.
Dựng transcript, nhật ký hook và manifest giả trong thư mục tạm, đặt HOME trỏ thư mục tạm, chạy routing_report.py
bằng subprocess (python3 -I). Không ghi vào repo, không đọc ~/.claude thật.
Dùng: python3 -I .claude/tests/test_report.py
Biến môi trường ROUTING_REPORT_SCRIPT: đường dẫn bản sao routing_report.py (cạnh routing_cases.json) để thử phá.
Mỗi tình huống in một dòng PASS hoặc FAIL; thoát mã 1 nếu có FAIL.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.environ.get("ROUTING_REPORT_SCRIPT") or os.path.join(HERE, "routing_report.py")
CASES = {c["id"]: c for c in json.load(open(os.path.join(HERE, "routing_cases.json"), encoding="utf-8"))}

# Khối nội dung của một message assistant: ("text", s), ("agent", subagent_type), ("tool", tên)
T = lambda s="Đã xử lý.": ("text", s)
A = lambda a: ("agent", a)
TOOL = lambda n: ("tool", n)


def sid_of(tag):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "test_report/" + tag))


class Tmp:
    """Thư mục tạm: HOME giả (chứa .claude/projects/p/<sid>.jsonl) và thư mục nhật ký."""

    def __init__(self):
        self.root = tempfile.mkdtemp(prefix="test_report_")
        self.home = os.path.join(self.root, "home")
        self.proj = os.path.join(self.home, ".claude", "projects", "proj")
        self.logs = os.path.join(self.root, "logs")
        os.makedirs(self.proj)
        os.makedirs(self.logs)
        self.n = 0

    def close(self):
        shutil.rmtree(self.root, ignore_errors=True)


def jl(path, rows):
    with open(path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write((r if isinstance(r, str) else json.dumps(r, ensure_ascii=False)) + "\n")


def build_transcript(prompt, msgs, pre=(), extra_user=(), api_error_msgs=(), split_lines=True, sid="x"):
    """Trả về (dòng transcript, danh sách (tool_use_id, subagent_type) theo thứ tự).
    msgs: danh sách message assistant, mỗi message là danh sách khối. prompt None: không có dòng prompt."""
    rows, tools, n = [], [], 0
    for p in pre:  # caveat, isMeta đứng trước prompt
        rows.append(p)
    if prompt is not None:
        rows.append({"type": "user", "isMeta": False, "message": {"role": "user", "content": prompt}, "sessionId": sid})
    for i, blocks in enumerate(msgs):
        mid = f"msg_{i}"
        per_line = []
        for b in blocks:
            if b[0] == "text":
                per_line.append({"type": "text", "text": b[1]})
            else:
                n += 1
                tid = f"toolu_{n:03d}"
                if b[0] == "agent":
                    per_line.append({"type": "tool_use", "id": tid, "name": "Agent",
                                     "input": {"subagent_type": b[1], "description": "d", "prompt": "p"}})
                    tools.append((tid, b[1]))
                else:
                    per_line.append({"type": "tool_use", "id": tid, "name": b[1], "input": {}})
        groups = [[x] for x in per_line] if split_lines else [per_line]
        for g in groups:
            rows.append({"type": "assistant", "message": {"id": mid, "role": "assistant", "content": g}, "sessionId": sid})
        if i < len(msgs) - 1:  # kết quả tool giữa các message
            res = [{"type": "tool_result", "tool_use_id": t, "content": "ok"} for t, _ in tools]
            rows.append({"type": "user", "message": {"role": "user", "content": res}})
    for j, blocks in enumerate(api_error_msgs):
        rows.append({"type": "assistant", "isApiErrorMessage": True,
                     "message": {"id": f"err_{j}", "role": "assistant", "content": [{"type": "text", "text": "API Error"}]}})
    for u in extra_user:
        rows.append({"type": "user", "message": {"role": "user", "content": u}})
    return rows, tools


def hook_rec(cid, agent, tid, sid, valid=True):
    return {"ts": "2026-10-09 10:00:00", "event": "dryrun", "case": cid, "agent": agent, "description": "d",
            "valid": valid, "errors": [] if valid else ["lỗi giả"], "tool_use_id": tid, "session_id": sid}


def make_run(tmp, n, cases):
    """cases: {cid: spec}; ghi routing-run-n.jsonl, routing-manifest-n.jsonl và transcript. Trả về đường dẫn nhật ký.
    spec: msgs, prompt, pre, extra_user, api_error, split, no_transcript, exit (số, 'missing'),
          hook_agents {i: agent}, hook_sid, hook_tid {i: tid}, hook_drop {i}, hook_extra [(agent, tid)], invalid {i}, no_hook."""
    log_rows, man_rows = [], []
    for cid, s in cases.items():
        sid = sid_of(f"{tmp.root}/{n}/{cid}")
        prompt = s.get("prompt", CASES[cid]["prompt"])
        rows, tools = build_transcript(prompt, s.get("msgs", []), s.get("pre", ()), s.get("extra_user", ()),
                                       s.get("api_error", ()), s.get("split", True), sid)
        if not s.get("no_transcript"):
            jl(os.path.join(tmp.proj, sid + ".jsonl"), rows)
        if not s.get("no_hook"):
            for i, (tid, agent) in enumerate(tools):
                if i in s.get("hook_drop", ()):
                    continue
                log_rows.append(hook_rec(cid, s.get("hook_agents", {}).get(i, agent), s.get("hook_tid", {}).get(i, tid),
                                         s.get("hook_sid", sid), valid=i not in s.get("invalid", ())))
            for k, (agent, tid) in enumerate(s.get("hook_extra", ())):
                log_rows.append(hook_rec(cid, agent, tid, sid))
        m = {"case": cid, "session_id": sid, "ts": "2026-10-09T10:00:00Z"}
        ex = s.get("exit", 0)
        if ex != "missing":
            m["exit"] = ex
        man_rows.append(m)
    log = os.path.join(tmp.logs, f"routing-run-{n}.jsonl")
    jl(log, log_rows)
    jl(os.path.join(tmp.logs, f"routing-manifest-{n}.jsonl"), man_rows if man_rows is not None else [])
    return log


def run_report(tmp, args):
    env = {"HOME": tmp.home, "PATH": os.environ.get("PATH", "/usr/bin:/bin")}
    p = subprocess.run([sys.executable, "-I", SCRIPT] + args, cwd=tmp.root, env=env, capture_output=True, text=True, timeout=60)
    return p.returncode, p.stdout + p.stderr


def case_line(out, cid):
    for ln in out.splitlines():
        if ln.startswith(cid + " "):
            return ln
    return ""


RESULTS = []


def scenario(name, cid, spec, want_ok, notes=(), absent=(), xfail=None):
    """Chạy ca đơn (manifest chỉ có cid, không cờ): kiểm dòng ca, ghi chú, mã thoát."""
    tmp = Tmp()
    try:
        log = make_run(tmp, 1, {cid: spec})
        rc, out = run_report(tmp, [log])
        ln = case_line(out, cid)
        verdict_ok = f"{cid} ĐÚNG" in ln
        problems = []
        if not ln:
            problems.append("không có dòng ca")
        if verdict_ok != want_ok:
            problems.append(f"kết quả {'ĐÚNG' if verdict_ok else 'SAI'}, mong {'ĐÚNG' if want_ok else 'SAI'}")
        if rc != (0 if want_ok else 1):
            problems.append(f"exit {rc}")
        for x in notes:
            if x not in ln:
                problems.append(f"thiếu ghi chú '{x}'")
        for x in absent:
            if x in ln:
                problems.append(f"không được có '{x}'")
        record(name, problems, xfail, ln)
    finally:
        tmp.close()


def record(name, problems, xfail=None, detail=""):
    if xfail:
        status = "XFAIL" if problems else "PASS"  # xfail: ghi nhận lỗi đã biết
        reason = f"ghi nhận lỗi đã biết: {xfail}" if problems else "không còn lỗi, bỏ xfail"
        RESULTS.append((status if problems else "FAIL", name, reason))
        return
    if problems:
        RESULTS.append(("FAIL", name, "; ".join(problems) + " | " + detail[:200]))
    else:
        RESULTS.append(("PASS", name, "đúng kỳ vọng"))


R3 = [A("researcher")] * 3

# --- c13 ---
scenario("c13 ba researcher cùng message: ĐÚNG", "c13", {"msgs": [R3 + [T()]]}, True, ["cùng_message"])
scenario("c13 lệnh thứ ba là qc trong message đầu: SAI", "c13",
         {"msgs": [[A("researcher"), A("researcher"), A("qc"), T()]]}, False)
scenario("c13 ba lệnh message đầu rồi một lệnh gọi lại ở message sau: ĐÚNG", "c13",
         {"msgs": [R3, [A("researcher"), T()]]}, True, ["cùng_message", "gọi_lại=1"])
scenario("c13 ba lệnh ở ba message khác nhau: SAI", "c13",
         {"msgs": [[A("researcher")], [A("researcher")], [A("researcher"), T()]]}, False, ["không_cùng_message"])
# --- c03, c15 ---
scenario("c03 giao bốn researcher: SAI", "c03", {"msgs": [[A("researcher")] * 4 + [T()]]}, False, ["số_lệnh<=3"])
scenario("c15 lệnh đầu là maker (hook và transcript khớp): SAI do agent ngoài expect", "c15",
         {"msgs": [[A("maker"), T()]]}, False, ["thực tế=['maker']"], ["lệch_agent", "log_lệch_transcript"])
scenario("c15 ba researcher: ĐÚNG", "c15", {"msgs": [R3 + [T()]]}, True, ["một_loại_agent"])
# --- c12 ---
scenario("c12 transcript chỉ có prompt user: SAI không_có_phản_hồi_hợp_lệ", "c12", {"msgs": []}, False,
         ["không_có_phản_hồi_hợp_lệ"])
scenario("c12 chỉ tool_use SendMessage không text: SAI", "c12", {"msgs": [[TOOL("SendMessage")]]}, False,
         ["không_được_giao_việc", "không_có_text_trả_lời"])
scenario("c12 phản hồi isApiErrorMessage: SAI", "c12", {"msgs": [], "api_error": [1]}, False,
         ["không_có_phản_hồi_hợp_lệ"])
scenario("c12 text trả lời, 0 tool_use, 0 hook: ĐÚNG", "c12", {"msgs": [[T("Nhóm phân việc theo loại tác vụ.")]]}, True)
scenario("c12 có bản ghi hook: SAI", "c12", {"msgs": [[A("researcher"), T()]]}, False, ["không_được_giao_việc"])
scenario("c12 không có transcript: SAI không_có_transcript", "c12", {"msgs": [[T()]], "no_transcript": True}, False,
         ["không_có_transcript"])
# --- manifest ---
scenario("manifest thiếu exit: SAI manifest_thiếu_exit", "c13", {"msgs": [R3 + [T()]], "exit": "missing"}, False,
         ["manifest_thiếu_exit"])
scenario("manifest exit 1: SAI claude_lỗi_exit", "c13", {"msgs": [R3 + [T()]], "exit": 1}, False, ["claude_lỗi_exit=1"])
# --- đối chiếu hook và transcript ---
scenario("hook khác agent so với subagent_type transcript: SAI lệch_agent", "c13",
         {"msgs": [R3 + [T()]], "hook_agents": {0: "maker"}}, False, ["lệch_agent"])
scenario("hook khác session_id so với manifest: SAI session_lệch", "c13",
         {"msgs": [R3 + [T()]], "hook_sid": sid_of("khac")}, False, ["session_lệch"])
scenario("hook có tool_use_id không có trong transcript: SAI", "c13",
         {"msgs": [R3 + [T()]], "hook_tid": {1: "toolu_khong_co"}}, False, ["tool_use_id_không_có_trong_transcript"])
scenario("số hook khác số tool_use Agent: SAI log_lệch_transcript", "c13",
         {"msgs": [R3 + [T()]], "hook_drop": {2}}, False, ["log_lệch_transcript(hook=2,tool_use=3)"])
# --- prompt ---
scenario("transcript có hai user message dạng chuỗi: SAI nhiều_user_message", "c13",
         {"msgs": [R3 + [T()]], "extra_user": [CASES["c12"]["prompt"]]}, False, ["nhiều_user_message(2)"])
scenario("prompt user khác prompt ca: SAI lẫn_prompt", "c13",
         {"msgs": [R3 + [T()]], "prompt": CASES["c12"]["prompt"]}, False, ["lẫn_prompt"])
CAVEAT = {"type": "user", "isMeta": True, "message": {"role": "user", "content": "<local-command-caveat>Caveat: x</local-command-caveat>"}}
CAVEAT2 = {"type": "user", "isMeta": False, "message": {"role": "user", "content": "Caveat: các tin nhắn dưới đây do lệnh cục bộ"}}
scenario("caveat và isMeta đứng trước prompt: không lẫn_prompt giả, ĐÚNG", "c13",
         {"msgs": [R3 + [T()]], "pre": [CAVEAT, CAVEAT2]}, True, ["cùng_message"], ["lẫn_prompt", "nhiều_user_message"])
scenario("một phản hồi tách ba dòng cùng message.id: tính một message", "c13",
         {"msgs": [R3], "split": True}, True, ["cùng_message", "số_message=1"])
scenario("lệnh không hợp lệ (valid false) ở ca ĐÚNG: ĐÚNG kèm CẢNH_BÁO", "c13",
         {"msgs": [R3 + [T()]], "invalid": {1}}, True, ["CẢNH_BÁO_lệnh_không_hợp_lệ=1"])


# --- nhiều ca: --expect-all, tập con, manifest rỗng, --runs ---
def multi_scenarios():
    good13 = {"msgs": [R3 + [T()]]}
    good12 = {"msgs": [[T("Trả lời.")]]}
    # manifest thiếu một ca
    tmp = Tmp()
    try:
        log = make_run(tmp, 1, {"c13": good13, "c12": good12})
        rc, out = run_report(tmp, [log, "--expect-all"])
        p = []
        if "vắng_trong_manifest" not in case_line(out, "c01"):
            p.append("c01 thiếu vắng_trong_manifest")
        if "c01 SAI" not in out:
            p.append("c01 không SAI")
        if rc != 1:
            p.append(f"exit {rc}, mong 1")
        record("manifest thiếu ca, --expect-all: SAI vắng_trong_manifest, exit 1", p, None, out[-300:])
        rc, out = run_report(tmp, [log])
        p = []
        if case_line(out, "c01"):
            p.append("ca vắng không bị bỏ qua")
        if rc != 0:
            p.append(f"exit {rc}, mong 0")
        if "tập con: đã chấm 2/18" not in out:
            p.append("thiếu thông báo tập con")
        record("manifest thiếu ca, không cờ (tập con): bỏ qua, exit 0", p, None, out[-300:])
    finally:
        tmp.close()
    # manifest rỗng
    tmp = Tmp()
    try:
        log = make_run(tmp, 1, {"c13": good13})
        jl(os.path.join(tmp.logs, "routing-manifest-1.jsonl"), [])
        rc, out = run_report(tmp, [log])
        p = []
        if "không_có_manifest" not in case_line(out, "c13") or rc != 1:
            p.append(f"không_có_manifest/exit sai (exit {rc})")
        record("manifest rỗng: SAI không_có_manifest", p, None, out[-300:])
    finally:
        tmp.close()
    # --runs ba nhật ký: 2 đạt, 1 SAI
    tmp = Tmp()
    try:
        logs = [make_run(tmp, 1, {"c13": good13}), make_run(tmp, 2, {"c13": good13}),
                make_run(tmp, 3, {"c13": {"msgs": [[A("researcher"), A("researcher"), A("qc"), T()]]}})]
        rc, out = run_report(tmp, ["--runs"] + logs)
        p = []
        if "c13 đạt ở 2/3 lần" not in out:
            p.append("thiếu 'đạt ở 2/3 lần'")
        if rc != 1:
            p.append(f"exit {rc}, mong 1")
        record("--runs ba nhật ký (2 đạt, 1 SAI): in đạt ở 2/3 lần, exit 1", p, None, out[-300:])
    finally:
        tmp.close()


multi_scenarios()


def main():
    nf = 0
    for status, name, why in RESULTS:
        print(f"{status} {name} ({why})")
        nf += status == "FAIL"
    npass = sum(s == "PASS" for s, _, _ in RESULTS)
    nx = sum(s == "XFAIL" for s, _, _ in RESULTS)
    print(f"Tổng: {len(RESULTS)} ca, PASS {npass}, FAIL {nf}" + (f", XFAIL {nx}" if nx else ""))
    return 1 if nf else 0


if __name__ == "__main__":
    sys.exit(main())

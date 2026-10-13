#!/usr/bin/env python3
"""Bộ kiểm hook offline: delegation_guard, haiku_context_guard, subagent_log.

Chạy: python3 -I .claude/tests/test_hooks.py
Không gọi claude, không đụng nhật ký thật: mọi ca chạy trong thư mục tạm,
CLAUDE_PROJECT_DIR và DELEGATION_LOG trỏ vào thư mục tạm. In PASS/FAIL từng ca; thoát 1 nếu có FAIL.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
HOOKS = os.path.join(os.path.dirname(HERE), "hooks")
GUARD = os.path.join(HOOKS, "delegation_guard.py")
HAIKU = os.path.join(HOOKS, "haiku_context_guard.py")
SLOG = os.path.join(HOOKS, "subagent_log.py")

ROOT = tempfile.mkdtemp(prefix="hooks_test_")
RESULTS = []
_n = [0]

BASE = {
    "MỤC TIÊU": "Kiểm tra sản phẩm",
    "NGỮ CẢNH": "Bối cảnh thật của nhiệm vụ",
    "PHẠM VI": "Chỉ tệp a.md",
    "TIÊU CHÍ ĐẠT": "Đúng số liệu",
    "GIỚI HẠN HÀNH ĐỘNG": "Không ghi tệp",
    "ĐỊNH DẠNG TRẢ VỀ": "Tiếng Việt, 5 dòng",
}
QC = {"MỨC QC": "1", "YÊU CẦU GỐC": "Làm báo cáo", "SẢN PHẨM": "/tmp/a.md"}
EXPERT = {"LÝ DO DÙNG EXPERT": "Cần ý kiến chuyên sâu"}


def workdir():
    _n[0] += 1
    d = os.path.join(ROOT, f"case{_n[0]}")
    os.makedirs(d)
    return d


def run_hook(hook, stdin, env_extra=None, cwd=None, proj=None):
    """Chạy hook; trả (rc, stderr, log_path, proj)."""
    proj = proj or workdir()
    log = os.path.join(proj, "log.jsonl")
    env = {k: v for k, v in os.environ.items()
           if not k.startswith(("DELEGATION_", "HAIKU_", "CLAUDE_"))}
    env["CLAUDE_PROJECT_DIR"] = proj
    env["DELEGATION_LOG"] = log
    env.update(env_extra or {})
    if isinstance(stdin, (dict, list)):
        stdin = json.dumps(stdin, ensure_ascii=False)
    p = subprocess.run([sys.executable, "-I", hook], input=stdin.encode("utf-8"),
                       capture_output=True, env=env, cwd=cwd or proj, timeout=30)
    return p.returncode, p.stderr.decode("utf-8", "replace"), log, proj


def read_log(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return [json.loads(x) for x in f if x.strip()]


def report(name, ok, why=""):
    RESULTS.append(ok)
    print(("PASS " if ok else "FAIL ") + name + ("" if ok else f" | {why}"))


def expect_rc(name, rc, want, stderr="", must_have=None):
    ok = rc == want
    why = f"rc={rc}, mong đợi {want}; stderr={stderr.strip()[:200]!r}"
    if ok and must_have and must_have not in stderr:
        ok, why = False, f"stderr thiếu {must_have!r}: {stderr.strip()[:200]!r}"
    report(name, ok, why)


def prompt_of(d):
    return "\n".join(f"{k}: {v}" for k, v in d.items())


def agent_input(agent, d, **extra):
    ti = {"subagent_type": agent, "description": "x", "prompt": prompt_of(d)}
    ti.update(extra)
    return {"tool_name": "Agent", "tool_input": ti}


# ---------- delegation_guard ----------
def test_delegation():
    p = "delegation_guard: "
    rc, err, _, _ = run_hook(GUARD, agent_input("maker", BASE))
    expect_rc(p + "đủ mục rc=0", rc, 0, err)

    for h in BASE:
        d = {k: v for k, v in BASE.items() if k != h}
        rc, err, _, _ = run_hook(GUARD, agent_input("maker", d))
        expect_rc(p + f"thiếu '{h}' rc=2", rc, 2, err, f"Thiếu mục '{h}:'")

    for h in QC:
        d = dict(BASE)
        d.update({k: v for k, v in QC.items() if k != h})
        rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
        expect_rc(p + f"qc thiếu '{h}' rc=2", rc, 2, err, f"Thiếu mục '{h}:'")

    d = dict(BASE)
    d.update(QC)
    rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "qc đủ mục rc=0", rc, 0, err)

    rc, err, _, _ = run_hook(GUARD, agent_input("expert", BASE))
    expect_rc(p + "expert thiếu 'LÝ DO DÙNG EXPERT' rc=2", rc, 2, err, "LÝ DO DÙNG EXPERT")
    d = dict(BASE)
    d.update(EXPERT)
    rc, err, _, _ = run_hook(GUARD, agent_input("expert", d))
    expect_rc(p + "expert đủ mục rc=0", rc, 0, err)

    rc, err, _, _ = run_hook(GUARD, agent_input("hacker", BASE))
    expect_rc(p + "agent ngoài đội hình rc=2", rc, 2, err, "không thuộc đội hình")

    # Hồi quy: NGỮ CẢNH mở đầu bằng tiêu đề khác ("Yêu cầu gốc:") không bị coi là rỗng.
    d = dict(BASE)
    d.update(QC)
    d["NGỮ CẢNH"] = "Yêu cầu gốc: /Users/x/brief.md; sản phẩm tại /Users/x/out.md"
    rc, err, log, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "NGỮ CẢNH mở đầu 'Yêu cầu gốc:' rc=0 (hồi quy)", rc, 0, err)
    recs = read_log(log)
    report(p + "NGỮ CẢNH mở đầu 'Yêu cầu gốc:' ghi valid=true", bool(recs) and recs[-1].get("valid") is True,
           f"log={recs[-1:] }")

    # Tiêu đề thật ở đầu dòng kế tiếp vẫn cắt mục: NGỮ CẢNH rỗng, dòng sau là PHẠM VI.
    d = dict(BASE)
    d["NGỮ CẢNH"] = ""
    rc, err, _, _ = run_hook(GUARD, agent_input("maker", d))
    expect_rc(p + "NGỮ CẢNH thật sự trống rc=2", rc, 2, err, "để trống")

    d = dict(BASE)
    d["NGỮ CẢNH"] = "dòng 1\ndòng 2\n- dòng 3"
    rc, err, _, _ = run_hook(GUARD, agent_input("maker", d))
    expect_rc(p + "mục nhiều dòng rc=0", rc, 0, err)

    # Hồi quy: nhãn con viết thường ở đầu dòng trong giá trị NGỮ CẢNH không phải tiêu đề mới.
    for agent in ("maker", "qc"):
        for sub in ("Yêu cầu gốc: x", "- Yêu cầu gốc: x", "- Phạm vi: x", "Sản phẩm: x"):
            d = dict(BASE)
            if agent == "qc":
                d.update(QC)
            d["NGỮ CẢNH"] = "\n" + sub
            rc, err, _, _ = run_hook(GUARD, agent_input(agent, d))
            expect_rc(p + f"{agent}: NGỮ CẢNH xuống dòng '{sub}' rc=0 (hồi quy)", rc, 0, err)
    d = dict(BASE)
    d.update(QC)
    d["NGỮ CẢNH"] = "Bối cảnh thật\nMức QC: 2"
    rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "'Mức QC: 2' trong NGỮ CẢNH, MỨC QC 1 rc=0 (hồi quy)", rc, 0, err)

    # Nhãn viết thường không thay được mục bắt buộc (so khớp phân biệt hoa thường).
    d = {k: v for k, v in dict(BASE, **QC).items() if k != "YÊU CẦU GỐC"}
    d["NGỮ CẢNH"] = "Bối cảnh thật\n- Yêu cầu gốc: làm báo cáo"
    rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "qc chỉ có '- Yêu cầu gốc:' trong NGỮ CẢNH rc=2", rc, 2, err, "Thiếu mục 'YÊU CẦU GỐC:'")
    d = {k: v for k, v in dict(BASE, **QC).items() if k != "YÊU CẦU GỐC"}
    rc, err, _, _ = run_hook(GUARD, {"tool_name": "Agent", "tool_input": {
        "subagent_type": "qc", "description": "x", "prompt": prompt_of(d) + "\nYêu cầu gốc: làm báo cáo"}})
    expect_rc(p + "qc chỉ có 'Yêu cầu gốc:' viết thường rc=2", rc, 2, err, "Thiếu mục 'YÊU CẦU GỐC:'")
    low = {k.lower(): v for k, v in BASE.items()}
    rc, err, _, _ = run_hook(GUARD, agent_input("maker", low))
    expect_rc(p + "tất cả tiêu đề viết thường rc=2", rc, 2, err, "Thiếu mục 'MỤC TIÊU:'")

    # Dạng tiêu đề được phép: số thứ tự, NBSP đầu dòng, dấu hai chấm toàn khổ.
    for name, fmt in (("số thứ tự '1.'", lambda i, k, v: f"{i}. {k}: {v}"),
                      ("số thứ tự '1)'", lambda i, k, v: f"{i}) {k}: {v}"),
                      ("NBSP đầu dòng", lambda i, k, v: f"\u00a0\u00a0{k}: {v}"),
                      ("dấu hai chấm toàn khổ", lambda i, k, v: f"{k}\uff1a{v}"),
                      ("markdown '## **...**'", lambda i, k, v: f"## **{k}**: {v}")):
        pr = "\n".join(fmt(i, k, v) for i, (k, v) in enumerate(BASE.items(), 1))
        rc, err, _, _ = run_hook(GUARD, {"tool_name": "Agent", "tool_input": {
            "subagent_type": "maker", "description": "x", "prompt": pr}})
        expect_rc(p + f"tiêu đề dạng {name} rc=0", rc, 0, err)

    # Hai tiêu đề cùng một dòng: tiêu đề sau không tính, mục sau coi là thiếu.
    pr = prompt_of(BASE).replace("MỤC TIÊU: Kiểm tra sản phẩm\nNGỮ CẢNH:", "MỤC TIÊU: Kiểm tra sản phẩm NGỮ CẢNH:")
    rc, err, _, _ = run_hook(GUARD, {"tool_name": "Agent", "tool_input": {
        "subagent_type": "maker", "description": "x", "prompt": pr}})
    expect_rc(p + "'MỤC TIÊU: NGỮ CẢNH: ...' cùng dòng rc=2", rc, 2, err, "Thiếu mục 'NGỮ CẢNH:'")

    rc, err, log, _ = run_hook(GUARD, agent_input("maker", BASE), {"DELEGATION_DRYRUN": "1"})
    recs = read_log(log)
    ok = rc == 2 and bool(recs) and recs[-1].get("event") == "dryrun"
    report(p + "DRYRUN=1 ghi event dryrun và rc=2", ok, f"rc={rc}, log={recs[-1:]}")

    bad = {
        "'[1]'": "[1]", "'null'": "null", "rỗng": "", "rác": "}{ không phải json",
        "tool_input là chuỗi": {"tool_name": "Agent", "tool_input": "abc"},
        "prompt là số": {"tool_name": "Agent", "tool_input": {"subagent_type": "maker", "prompt": 123}},
    }
    for name, val in bad.items():
        rc, err, _, _ = run_hook(GUARD, val)
        expect_rc(p + f"input hỏng {name} rc=0", rc, 0, err)
    clerk = agent_input("clerk", BASE)
    rc, err, _, _ = run_hook(GUARD, clerk, {"DELEGATION_HAIKU_MAX_TOKENS": "abc"})
    expect_rc(p + "DELEGATION_HAIKU_MAX_TOKENS=abc rc=0", rc, 0, err)

    # Giới hạn đã biết: brief gốc dán vào YÊU CẦU GỐC có dòng tiêu đề HOA thì hook coi là mục mới.
    # Cách xử lý: thêm "> " đầu mỗi dòng đó (hoặc ghi đường dẫn thay vì dán).
    pasted = "\nMỤC TIÊU: Làm báo cáo cho khách\nPHẠM VI: Chỉ tệp a.md"
    d = dict(BASE, **QC)
    d["YÊU CẦU GỐC"] = pasted
    rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "qc: brief dán có tiêu đề HOA không '> ' rc=2 (giới hạn đã biết)", rc, 2, err, "YÊU CẦU GỐC")
    d["YÊU CẦU GỐC"] = pasted.replace("\nMỤC TIÊU", "\n> MỤC TIÊU").replace("\nPHẠM VI", "\n> PHẠM VI")
    rc, err, _, _ = run_hook(GUARD, agent_input("qc", d))
    expect_rc(p + "qc: brief dán có '> ' đầu dòng tiêu đề HOA rc=0", rc, 0, err)


# ---------- transcript giả ----------
def make_transcripts(proj, agent_id, model, usage, session="sess"):
    """Tạo <proj>/<session>.jsonl và <proj>/<session>/subagents/agent-<id>.jsonl; trả (main, sub)."""
    main = os.path.join(proj, session + ".jsonl")
    subdir = os.path.join(proj, session, "subagents")
    os.makedirs(subdir)
    sub = os.path.join(subdir, f"agent-{agent_id}.jsonl")
    with open(main, "w", encoding="utf-8") as f:
        f.write(json.dumps({"type": "user", "message": {"role": "user", "content": "hi"}}) + "\n")
    lines = [{"type": "user", "message": {"role": "user", "content": "việc"}}]
    msg = {"role": "assistant", "content": [{"type": "text", "text": "ok"}]}
    if model is not None:
        msg["model"] = model
    if usage is not None:
        msg["usage"] = usage
    lines.append({"type": "assistant", "message": msg})
    with open(sub, "w", encoding="utf-8") as f:
        for o in lines:
            f.write(json.dumps(o, ensure_ascii=False) + "\n")
    return main, sub


def usage_of(n):
    """Tách n token thành ba trường usage, tổng đúng n."""
    return {"input_tokens": 1000, "cache_read_input_tokens": n - 1500, "cache_creation_input_tokens": 500,
            "output_tokens": 50}


HAIKU_MODEL = "claude-haiku-4-5"
OPUS_MODEL = "claude-opus-4-5"


# ---------- haiku_context_guard ----------
HAIKU_EXTRA = {}  # biến môi trường thêm cho test_haiku (chế độ cờ)


def haiku_case(n, tool="Read", model=HAIKU_MODEL, env=None, agent_id="A1"):
    env = dict(HAIKU_EXTRA, **(env or {}))
    proj = workdir()
    main, _ = make_transcripts(proj, "A1", model, usage_of(n))
    inp = {"tool_name": tool, "transcript_path": main}
    if agent_id:
        inp["agent_id"] = agent_id
    return run_hook(HAIKU, inp, env, proj=proj)


def test_haiku(label="", extra=None):
    """Ca ngưỡng token; chạy ở cả hai chế độ cờ (tắt mặc định và HAIKU_ONE_TOOL_PER_TURN=1)."""
    HAIKU_EXTRA.clear()
    HAIKU_EXTRA.update(extra or {})
    p = "haiku_context_guard" + label + ": "
    rc, err, _, _ = haiku_case(64999)
    expect_rc(p + "prompt 64.999 rc=0", rc, 0, err)
    rc, err, _, _ = haiku_case(65000)
    expect_rc(p + "prompt 65.000 Read rc=2", rc, 2, err, "65000")
    rc, err, _, _ = haiku_case(70000)
    expect_rc(p + "prompt 70.000 Read rc=2", rc, 2, err, "70000")
    for n in (64999, 65000, 70000, 150000):
        rc, err, _, _ = haiku_case(n, tool="SubagentHandback")
        expect_rc(p + f"SubagentHandback ở {n} rc=0", rc, 0, err)
    rc, err, _, _ = haiku_case(70000, agent_id=None)
    expect_rc(p + "không agent_id rc=0", rc, 0, err)
    rc, err, _, _ = haiku_case(70000, agent_id="KHAC")
    expect_rc(p + "không có transcript subagent rc=0", rc, 0, err)
    rc, err, _, _ = haiku_case(70000, model=OPUS_MODEL)
    expect_rc(p + "model không phải Haiku rc=0", rc, 0, err)
    rc, err, _, _ = haiku_case(65000, env={"HAIKU_CONTEXT_BLOCK_TOKENS": "abc"})
    expect_rc(p + "ngưỡng 'abc' dùng 65000 (65.000 chặn)", rc, 2, err)
    rc, err, _, _ = haiku_case(64999, env={"HAIKU_CONTEXT_BLOCK_TOKENS": "abc"})
    expect_rc(p + "ngưỡng 'abc' dùng 65000 (64.999 qua)", rc, 0, err)
    rc, err, _, _ = haiku_case(70000, env={"HAIKU_CONTEXT_BLOCK_TOKENS": "80000"})
    expect_rc(p + "ngưỡng 80000: 70.000 qua", rc, 0, err)
    for name, val in {"'[1]'": "[1]", "'null'": "null", "rỗng": "", "rác": "}{",
                      "agent_id số, không transcript_path": {"agent_id": 5},
                      "transcript_path là list": {"agent_id": "A1", "transcript_path": ["x"]}}.items():
        rc, err, _, _ = run_hook(HAIKU, val, dict(HAIKU_EXTRA))
        expect_rc(p + f"input hỏng {name} rc=0", rc, 0, err)

    # Nhật ký tương đối ghi vào CLAUDE_PROJECT_DIR/.claude dù cwd là thư mục con.
    proj = workdir()
    main, _ = make_transcripts(proj, "A1", HAIKU_MODEL, usage_of(70000))
    sub = os.path.join(proj, "sub")
    os.makedirs(sub)
    os.makedirs(os.path.join(proj, ".claude"))  # hook không tự tạo thư mục; dự án thật đã có .claude
    rc, err, _, _ = run_hook(HAIKU, {"tool_name": "Read", "transcript_path": main, "agent_id": "A1"},
                             dict(HAIKU_EXTRA, HAIKU_CONTEXT_GUARD_LOG=".claude/guard.log"), cwd=sub, proj=proj)
    want = os.path.join(proj, ".claude", "guard.log")
    wrong = os.path.join(sub, ".claude", "guard.log")
    ok = rc == 2 and os.path.exists(want) and "CHAN" in open(want, encoding="utf-8").read() \
        and not os.path.exists(wrong)
    report(p + "log tương đối ghi vào CLAUDE_PROJECT_DIR/.claude (cwd thư mục con)", ok,
           f"rc={rc}, có {want}={os.path.exists(want)}, có {wrong}={os.path.exists(wrong)}")
    HAIKU_EXTRA.clear()


# ---------- subagent_log ----------
def slog_input(sub, **extra):
    d = {"agent_transcript_path": sub, "agent_type": "qc", "agent_id": "A1",
         "last_assistant_message": "xong", "stop_hook_active": False}
    d.update(extra)
    return d


def test_slog():
    p = "subagent_log: "
    for label, model, over in (("Haiku", HAIKU_MODEL, True), ("Opus", OPUS_MODEL, False)):
        proj = workdir()
        _, sub = make_transcripts(proj, "A1", model, usage_of(120000))
        rc, err, log, _ = run_hook(SLOG, slog_input(sub), proj=proj)
        recs = read_log(log)
        r = recs[-1] if len(recs) == 1 else {}
        want = {"event": "stop", "agent": "qc", "agent_id": "A1", "model": model,
                "last_prompt_tokens": 120000, "over_100k": over, "transcript_exists": True,
                "transcript_dong": 2, "model_nguon": "transcript", "agent_transcript_path": sub,
                "last_assistant_message": "xong", "stop_hook_active": False}
        diff = {k: (r.get(k), v) for k, v in want.items() if r.get(k) != v}
        ok = rc == 0 and bool(r) and not diff and "stop_reason" not in r
        report(p + f"có usage {label} 120.000 đủ trường, over_100k={over}, không stop_reason", ok,
               f"rc={rc}, bản ghi={len(recs)}, sai lệch (thực, mong đợi)={diff}, stop_reason={'stop_reason' in r}")

    proj = workdir()
    _, sub = make_transcripts(proj, "A1", None, None)
    rc, err, log, _ = run_hook(SLOG, slog_input(sub), proj=proj)
    recs = read_log(log)
    r = recs[-1] if recs else {}
    ok = rc == 0 and r.get("model") == "" and r.get("model_nguon") == "không_có_usage" \
        and r.get("last_prompt_tokens") == 0 and r.get("over_100k") is False
    report(p + "không usage: model rỗng, model_nguon không_có_usage", ok, f"rc={rc}, bản ghi={r}")

    proj = workdir()
    rc, err, log, _ = run_hook(SLOG, slog_input(os.path.join(proj, "khong_ton_tai.jsonl")), proj=proj)
    report(p + "transcript không tồn tại: rc=0, không ghi", rc == 0 and not read_log(log),
           f"rc={rc}, log={read_log(log)}")

    for name, val in {"stdin hỏng (rác)": "}{", "stdin rỗng": "", "'null'": "null"}.items():
        rc, err, log, _ = run_hook(SLOG, val)
        report(p + f"{name} rc=0", rc == 0 and not read_log(log), f"rc={rc}, log={read_log(log)}")
    rc, err, log, _ = run_hook(SLOG, {"agent_transcript_path": ["a", "b"], "agent_type": "qc"})
    report(p + "transcript_path là list rc=0, không ghi", rc == 0 and not read_log(log),
           f"rc={rc}, log={read_log(log)}")


# ---------- haiku_context_guard: mỗi lượt một công cụ (khối riêng) ----------
PAR_MSG = "Mỗi lượt chỉ gọi một công cụ. Gọi lại từng công cụ một ở lượt sau."


def par_append(sub, msg_id, tool_ids, n=20000, model=HAIKU_MODEL, usage=True):
    """Ghi vào transcript subagent một message assistant: mỗi tool_use một dòng, cùng message.id."""
    with open(sub, "a", encoding="utf-8") as f:
        for tid in tool_ids:
            msg = {"id": msg_id, "role": "assistant", "model": model,
                   "content": [{"type": "tool_use", "id": tid, "name": "Read", "input": {}}]}
            if usage:
                msg["usage"] = usage_of(n)
            f.write(json.dumps({"type": "assistant", "message": msg}) + "\n")


def par_run(proj, main, tool_id=None, tool="Read", agent_id="A1", env=None, flag="1"):
    """Chạy hook với cờ HAIKU_ONE_TOOL_PER_TURN=`flag` (None: không đặt biến)."""
    inp = {"tool_name": tool, "transcript_path": main, "agent_id": agent_id}
    if tool_id:
        inp["tool_use_id"] = tool_id
    e = {"HAIKU_CONTEXT_GUARD_LOG": ".claude/guard.log"}
    if flag is not None:
        e["HAIKU_ONE_TOOL_PER_TURN"] = flag
    e.update(env or {})
    return run_hook(HAIKU, inp, e, proj=proj)


def par_proj(agent_id="A1"):
    proj = workdir()
    os.makedirs(os.path.join(proj, ".claude"))
    main, sub = make_transcripts(proj, agent_id, HAIKU_MODEL, usage_of(20000))
    return proj, main, sub


def par_log(proj):
    fp = os.path.join(proj, ".claude", "guard.log")
    return open(fp, encoding="utf-8").read() if os.path.exists(fp) else ""


def test_haiku_parallel():
    """Cơ chế một công cụ mỗi lượt: mọi ca chạy với HAIKU_ONE_TOOL_PER_TURN=1."""
    p = "haiku_context_guard song song (cờ bật): "
    # 1 lệnh trong lượt: qua
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1"])
    rc, err, _, _ = par_run(proj, main, "t1")
    expect_rc(p + "lượt 1 lệnh rc=0", rc, 0, err)

    # 3 lệnh cùng message: lệnh 1 qua, 2 và 3 bị chặn; lượt kế tiếp qua lại
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1", "t2", "t3"])
    rcs = []
    for t in ("t1", "t2", "t3"):
        rc, err, _, _ = par_run(proj, main, t)
        rcs.append((rc, err))
    expect_rc(p + "lượt 3 lệnh: lệnh 1 rc=0", rcs[0][0], 0, rcs[0][1])
    expect_rc(p + "lượt 3 lệnh: lệnh 2 rc=2, đúng thông báo", rcs[1][0], 2, rcs[1][1], PAR_MSG)
    expect_rc(p + "lượt 3 lệnh: lệnh 3 rc=2, đúng thông báo", rcs[2][0], 2, rcs[2][1], PAR_MSG)
    n_log = par_log(proj).count("CHAN_SONG_SONG agent=A1")
    report(p + "log có 2 dòng CHAN_SONG_SONG agent=A1", n_log == 2, f"số dòng={n_log}")
    par_append(sub, "msg_2", ["t4"])
    rc, err, _, _ = par_run(proj, main, "t4")
    expect_rc(p + "lượt kế tiếp lệnh 1 rc=0 (đặt lại đếm)", rc, 0, err)
    par_append(sub, "msg_2", ["t5"])
    rc, err, _, _ = par_run(proj, main, "t5")
    expect_rc(p + "lượt kế tiếp lệnh 2 rc=2", rc, 2, err, PAR_MSG)

    # message đầu của subagent chưa có usage: model lấy từ chính message chứa lệnh
    for mdl, want in ((HAIKU_MODEL, [0, 2, 2]), (OPUS_MODEL, [0, 0, 0])):
        proj = workdir()
        os.makedirs(os.path.join(proj, ".claude"))
        main, sub = make_transcripts(proj, "A1", None, None)
        par_append(sub, "msg_1", ["t1", "t2", "t3"], model=mdl, usage=False)
        rcs = [par_run(proj, main, t)[0] for t in ("t1", "t2", "t3")]
        report(p + f"message đầu không usage, model {mdl}: rc={want}", rcs == want, f"rcs={rcs}")

    # SubagentHandback không tính vào đếm và không bị chặn
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1", "h1", "t2"])
    rc, err, _, _ = par_run(proj, main, "h1", tool="SubagentHandback")
    expect_rc(p + "SubagentHandback đầu lượt rc=0", rc, 0, err)
    rc, err, _, _ = par_run(proj, main, "t1")
    expect_rc(p + "Read sau Handback vẫn là lệnh đầu rc=0", rc, 0, err)
    rc, err, _, _ = par_run(proj, main, "h1", tool="SubagentHandback")
    expect_rc(p + "SubagentHandback sau Read rc=0", rc, 0, err)
    rc, err, _, _ = par_run(proj, main, "t2")
    expect_rc(p + "Read thứ hai sau Handback rc=2", rc, 2, err, PAR_MSG)

    # agent_id khác nhau không ảnh hưởng nhau (cùng dự án)
    proj, main, sub = par_proj("A1")
    main2, sub2 = make_transcripts(proj, "A2", HAIKU_MODEL, usage_of(20000), session="sess2")
    par_append(sub, "msg_1", ["t1", "t2"])
    par_append(sub2, "msg_1", ["u1"])
    rc1, e1, _, _ = par_run(proj, main, "t1", agent_id="A1")
    rc2, e2, _, _ = par_run(proj, main2, "u1", agent_id="A2")
    rc3, e3, _, _ = par_run(proj, main, "t2", agent_id="A1")
    expect_rc(p + "A1 lệnh 1 rc=0", rc1, 0, e1)
    expect_rc(p + "A2 cùng message.id vẫn lệnh đầu rc=0", rc2, 0, e2)
    expect_rc(p + "A1 lệnh 2 rc=2", rc3, 2, e3, PAR_MSG)

    # lỗi ghi trạng thái: guard-state là tệp thường, rc vẫn 0
    proj, main, sub = par_proj()
    open(os.path.join(proj, ".claude", "guard-state"), "w").close()
    par_append(sub, "msg_1", ["t1", "t2"])
    rcs = [par_run(proj, main, t)[0] for t in ("t1", "t2")]
    report(p + "lỗi ghi trạng thái rc=0", rcs == [0, 0], f"rcs={rcs}")
    # thư mục trạng thái chỉ đọc
    proj, main, sub = par_proj()
    sd = os.path.join(proj, ".claude", "guard-state")
    os.makedirs(sd)
    os.chmod(sd, 0o500)
    par_append(sub, "msg_1", ["t1", "t2"])
    rcs = [par_run(proj, main, t)[0] for t in ("t1", "t2")]
    os.chmod(sd, 0o700)
    report(p + "thư mục trạng thái chỉ đọc rc=0", rcs == [0, 0], f"rcs={rcs}")
    # tệp trạng thái hỏng: coi như lượt mới, rc=0
    proj, main, sub = par_proj()
    sd = os.path.join(proj, ".claude", "guard-state")
    os.makedirs(sd)
    with open(os.path.join(sd, "A1.json"), "w") as f:
        f.write("}{")
    par_append(sub, "msg_1", ["t1"])
    rc, err, _, _ = par_run(proj, main, "t1")
    expect_rc(p + "tệp trạng thái hỏng rc=0", rc, 0, err)

    # không có tool_use_id: cho qua (không đoán dấu từ message cũ, tránh chặn oan lượt sau)
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1", "t2"])
    rcs = [par_run(proj, main)[0] for _ in range(3)]
    report(p + "không tool_use_id rc=0,0,0", rcs == [0, 0, 0], f"rcs={rcs}")
    # dòng tool_use được ghi trễ sau khi hook đã chạy (thực tế Claude Code ghi sau hook): hook chờ và vẫn đếm đúng
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1"])
    par_run(proj, main, "t1")
    inp = {"tool_name": "Read", "transcript_path": main, "agent_id": "A1", "tool_use_id": "t2"}
    env = {k: v for k, v in os.environ.items() if not k.startswith(("DELEGATION_", "HAIKU_", "CLAUDE_"))}
    env.update({"CLAUDE_PROJECT_DIR": proj, "HAIKU_CONTEXT_GUARD_LOG": ".claude/guard.log",
                "HAIKU_ONE_TOOL_PER_TURN": "1"})
    proc = subprocess.Popen([sys.executable, "-I", HAIKU], stdin=subprocess.PIPE, stderr=subprocess.PIPE,
                            env=env, cwd=proj)
    proc.stdin.write(json.dumps(inp).encode())
    proc.stdin.close()
    time.sleep(0.4)
    par_append(sub, "msg_1", ["t2"])
    rc = proc.wait(timeout=10)
    err = proc.stderr.read().decode("utf-8", "replace")
    expect_rc(p + "dòng tool_use ghi trễ 0,4 giây vẫn chặn lệnh 2", rc, 2, err, PAR_MSG)
    # tool_use_id chưa có trong transcript: cho qua, không đếm
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1"])
    rcs = [par_run(proj, main, "chua_co")[0] for _ in range(3)]
    report(p + "tool_use_id không có trong transcript rc=0", rcs == [0, 0, 0], f"rcs={rcs}")
    # agent_id có ký tự lạ (đường dẫn): cho qua, không ghi tệp ngoài guard-state
    proj, main, sub = par_proj()
    par_append(sub, "msg_1", ["t1", "t2"])
    rcs = [par_run(proj, main, t, agent_id="../x")[0] for t in ("t1", "t2")]
    report(p + "agent_id chứa '../' rc=0", rcs == [0, 0], f"rcs={rcs}")
    # ngưỡng token vẫn ưu tiên và giữ thông báo cũ
    proj = workdir()
    os.makedirs(os.path.join(proj, ".claude"))
    main, sub = make_transcripts(proj, "A1", HAIKU_MODEL, usage_of(70000))
    par_append(sub, "msg_1", ["t1"], n=70000)
    rc, err, _, _ = par_run(proj, main, "t1")
    expect_rc(p + "70.000 vẫn chặn theo ngưỡng", rc, 2, err, "70000")

    # dọn tệp cũ hơn 24 giờ, chỉ trong guard-state; giới hạn 20 tệp mỗi lần
    proj, main, sub = par_proj()
    sd = os.path.join(proj, ".claude", "guard-state")
    os.makedirs(sd)
    old = time.time() - 25 * 3600
    for i in range(25):
        fp = os.path.join(sd, f"old{i}.json")
        open(fp, "w").close()
        os.utime(fp, (old, old))
    fresh = os.path.join(sd, "fresh.json")
    open(fresh, "w").close()
    other = os.path.join(proj, ".claude", "giu_lai.json")
    open(other, "w").close()
    os.utime(other, (old, old))
    par_append(sub, "msg_1", ["t1"])
    par_run(proj, main, "t1")
    left_old = [x for x in os.listdir(sd) if x.startswith("old")]
    report(p + "dọn tối đa 20 tệp cũ mỗi lần, giữ tệp mới và tệp ngoài guard-state",
           len(left_old) == 5 and os.path.exists(fresh) and os.path.exists(other),
           f"còn old={len(left_old)}, fresh={os.path.exists(fresh)}, other={os.path.exists(other)}")


def test_haiku_off():
    """Cờ tắt (không đặt biến hoặc "0"): hook không cưỡng chế, không chờ, không tạo trạng thái, không ghi LUOT."""
    for flag in (None, "0"):
        p = "haiku_context_guard cờ tắt " + ("(không đặt biến)" if flag is None else f"(={flag!r})") + ": "
        env = {"HAIKU_CONTEXT_GUARD_DEBUG": "1"}
        for first_no_usage in (False, True):
            proj = workdir()
            os.makedirs(os.path.join(proj, ".claude"))
            if first_no_usage:
                main, sub = make_transcripts(proj, "A1", None, None)
                par_append(sub, "msg_1", ["t1", "t2", "t3"], usage=False)
            else:
                main, sub = make_transcripts(proj, "A1", HAIKU_MODEL, usage_of(20000))
                par_append(sub, "msg_1", ["t1", "t2", "t3"])
            rcs, times = [], []
            for t in ("t1", "t2", "t3"):
                t0 = time.time()
                rc, err, _, _ = par_run(proj, main, t, env=env, flag=flag)
                times.append(time.time() - t0)
                rcs.append(rc)
            tag = "message đầu không usage" if first_no_usage else "3 lệnh cùng lượt"
            report(p + f"{tag}: rc=0,0,0", rcs == [0, 0, 0], f"rcs={rcs}")
            report(p + f"{tag}: mỗi lệnh dưới 0,5 giây", max(times) < 0.5, f"times={times}")
            report(p + f"{tag}: không tạo guard-state",
                   not os.path.exists(os.path.join(proj, ".claude", "guard-state")), "thư mục tồn tại")
            lg = par_log(proj)
            report(p + f"{tag}: log không có CHAN_SONG_SONG hay LUOT",
                   "CHAN_SONG_SONG" not in lg and "LUOT" not in lg, f"log={lg!r}")
        # ngưỡng vẫn hoạt động khi cờ tắt
        proj = workdir()
        os.makedirs(os.path.join(proj, ".claude"))
        main, sub = make_transcripts(proj, "A1", HAIKU_MODEL, usage_of(70000))
        par_append(sub, "msg_1", ["t1"], n=70000)
        rc, err, _, _ = par_run(proj, main, "t1", flag=flag)
        expect_rc(p + "70.000 vẫn chặn theo ngưỡng", rc, 2, err, "70000")


def main():
    try:
        test_delegation()
        test_haiku(" (cờ tắt, mặc định)", {})
        test_haiku(" (cờ bật)", {"HAIKU_ONE_TOOL_PER_TURN": "1"})
        test_haiku_off()
        test_slog()
        test_haiku_parallel()
    finally:
        shutil.rmtree(ROOT, ignore_errors=True)
    fails = RESULTS.count(False)
    print(f"\nTổng: {len(RESULTS)} ca, PASS {RESULTS.count(True)}, FAIL {fails}")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())

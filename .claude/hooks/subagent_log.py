#!/usr/bin/env python3
"""Ghi nhật ký tự động mỗi khi một subagent kết thúc (SubagentStop).

Ghi một dòng JSON vào nhật ký (mặc định .claude/team-log.jsonl trong thư mục dự án; biến DELEGATION_LOG ghi đè):
thời điểm, loại agent, model, kích thước prompt của yêu cầu cuối. Coordinator không cần (và không có công cụ) để tự ghi nhật ký.
"""
import json
import os
import sys
import time


def last_assistant(path):
    if not isinstance(path, str) or not path or not os.path.exists(path):
        return "", {}
    with open(path, "rb") as f:
        f.seek(0, 2)
        size = f.tell()
        f.seek(max(0, size - 1_500_000))
        tail = f.read().decode("utf-8", "ignore")
    for line in reversed(tail.splitlines()):
        try:
            o = json.loads(line)
        except ValueError:
            continue
        m = o.get("message") if isinstance(o, dict) else None
        if isinstance(m, dict) and m.get("role") == "assistant" and isinstance(m.get("usage"), dict):
            return str(m.get("model", "")), m["usage"]
    return "", {}


def main():
    try:
        d = json.load(sys.stdin)
    except ValueError:
        return 0
    if not isinstance(d, dict):
        return 0
    tpath = d.get("agent_transcript_path")
    if not isinstance(tpath, str):
        return 0
    tpath = os.path.expanduser(tpath)
    if not tpath or not os.path.isfile(tpath):
        return 0
    try:
        model, usage = last_assistant(tpath)
    except OSError:
        model, usage = "", {}
    exists = True
    dong = None
    if exists:
        try:
            with open(tpath, "rb") as f:
                dong = sum(1 for _ in f)
        except OSError:
            dong = None
    nguon = "transcript" if usage else "không_có_usage"
    prompt = 0
    for k in ("input_tokens", "cache_read_input_tokens", "cache_creation_input_tokens"):
        try:
            prompt += int(usage.get(k) or 0)
        except (TypeError, ValueError, OverflowError):
            continue
    lam = d.get("last_assistant_message")
    sha = d.get("stop_hook_active")
    rec = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event": "stop",
        "agent": d.get("agent_type"),
        "agent_id": d.get("agent_id"),
        "model": model,
        "last_prompt_tokens": prompt,
        "over_100k": "haiku" in model.lower() and prompt > 100000,
        "agent_transcript_path": tpath or None,
        "transcript_exists": exists,
        "transcript_dong": dong,
        "model_nguon": nguon,
        "last_assistant_message": lam[:120] if isinstance(lam, str) else None,
        "stop_hook_active": sha if isinstance(sha, bool) else None,
    }
    path = os.path.expanduser(os.environ.get("DELEGATION_LOG") or os.path.join(
        os.environ.get("CLAUDE_PROJECT_DIR", os.getcwd()), ".claude", "team-log.jsonl"))
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except (OSError, ValueError):
        pass
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        code = 0
    sys.exit(code)

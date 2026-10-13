#!/usr/bin/env python3
"""Kiểm tra mỗi lệnh giao việc của coordinator trước khi chạy (PreToolUse, công cụ Agent).

Chặn (mã thoát 2, stderr trả về cho coordinator) khi:
- loại agent không thuộc đội hình;
- thiếu một trong sáu mục bắt buộc của mẫu giao việc;
- giao cho Haiku (Explore, clerk) mục tiêu dài hơn ngưỡng ước lượng token;
- dùng expert mà không nêu lý do; giao qc mà thiếu mức QC, yêu cầu gốc, sản phẩm, hoặc thiếu LƯỢT khi mức QC là 2;
- đổi model theo lần gọi ngoài trường hợp qc chuyển sang Sonnet.

Giới hạn đã biết: dòng bắt đầu bằng tên trong ALL_HEADERS (11 tên mục viết HOA) ở trong giá trị một mục
(ví dụ brief gốc dán vào YÊU CẦU GỐC) tính là mục mới và cắt giá trị của mục đó; tiêu đề HOA khác không cắt. Xử lý: thêm "> " đầu mỗi dòng đó hoặc ghi đường dẫn thay vì dán.

Chế độ chạy thử định tuyến: đặt DELEGATION_DRYRUN=1. Hook ghi lựa chọn của coordinator
vào nhật ký rồi chặn lệnh, để đo coordinator chọn agent nào mà không tốn token của agent.
Biến môi trường: DELEGATION_LOG (đường dẫn nhật ký), DELEGATION_CASE (mã ca thử),
DELEGATION_HAIKU_MAX_TOKENS (mặc định 3000).
"""
import json
import math
import os
import hashlib
import re
import sys
import time
import unicodedata

ROSTER = {"Explore", "clerk", "researcher", "analyst", "maker", "senior-maker", "planner", "qc", "expert"}
HAIKU = {"Explore", "clerk"}
REQUIRED = ["MỤC TIÊU", "NGỮ CẢNH", "PHẠM VI", "TIÊU CHÍ ĐẠT", "GIỚI HẠN HÀNH ĐỘNG", "ĐỊNH DẠNG TRẢ VỀ"]
EXTRA = {"expert": ["LÝ DO DÙNG EXPERT"], "qc": ["MỨC QC", "YÊU CẦU GỐC", "SẢN PHẨM"]}
ALL_HEADERS = REQUIRED + ["LÝ DO DÙNG EXPERT", "MỨC QC", "LƯỢT", "YÊU CẦU GỐC", "SẢN PHẨM"]
CHARS_PER_TOKEN = 2.5
MIN_LEN = {"MỨC QC": 1, "LƯỢT": 1}
DEFAULT_HAIKU_MAX_TOKENS = 3000


def env_positive_int(name, default):
    """Đọc biến môi trường là số nguyên dương; sai kiểu hoặc không dương thì dùng mặc định."""
    try:
        v = int(os.environ.get(name, "").strip())
        return v if v > 0 else default
    except (ValueError, TypeError, AttributeError):
        return default


MAX_HAIKU_TOKENS = env_positive_int("DELEGATION_HAIKU_MAX_TOKENS", DEFAULT_HAIKU_MAX_TOKENS)


def norm(s):
    return unicodedata.normalize("NFC", s or "")


def project_dir():
    return os.environ.get("CLAUDE_PROJECT_DIR", os.getcwd())


def log_path():
    return os.environ.get("DELEGATION_LOG") or os.path.join(project_dir(), ".claude", "team-log.jsonl")


def write_log(rec):
    try:
        os.makedirs(os.path.dirname(log_path()), exist_ok=True)
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError:
        pass


def find_message_id(transcript_path, tool_use_id):
    """Tìm id của message assistant chứa lệnh có tool_use_id trong transcript JSONL.
    Trả về None khi thiếu dữ kiện, tệp không đọc được hoặc chưa có dòng khớp; không bao giờ ném lỗi."""
    try:
        if not transcript_path or not tool_use_id:
            return None
        fallback = None
        with open(os.path.expanduser(transcript_path), encoding="utf-8", errors="replace") as f:
            for line in f:
                if tool_use_id not in line:
                    continue
                try:
                    obj = json.loads(line)
                except ValueError:
                    continue
                if not isinstance(obj, dict):
                    continue
                msg = obj.get("message") if isinstance(obj.get("message"), dict) else {}
                content = msg.get("content")
                is_use = isinstance(content, list) and any(
                    isinstance(b, dict) and b.get("type") == "tool_use" and b.get("id") == tool_use_id
                    for b in content)
                mid = msg.get("id") or obj.get("requestId") or obj.get("message_id")
                if is_use and mid:
                    return str(mid)
                if mid and fallback is None and obj.get("type") != "user":
                    fallback = str(mid)
        return fallback
    except Exception:
        return None


def transcript_tail(transcript_path):
    """Trả về (transcript_tail, transcript_lines) của transcript lúc hook chạy.
    Dấu = uuid của dòng cuối nếu có, không thì 'sha1:<băm dòng cuối>:<số dòng>'. Chỉ đọc 64 KB cuối để lấy dòng cuối,
    đếm số dòng theo khối. Lỗi hoặc thiếu dữ kiện trả (None, None); không bao giờ ném lỗi."""
    try:
        if not transcript_path:
            return None, None
        path = os.path.expanduser(transcript_path)
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            lines = 0
            last = b""
            while True:
                chunk = f.read(1 << 20)
                if not chunk:
                    break
                lines += chunk.count(b"\n")
                last = chunk[-1:]
            if size and last != b"\n":
                lines += 1
            start = max(0, size - 65536)
            f.seek(start)
            data = f.read()
        parts = data.split(b"\n")
        if start > 0 and len(parts) > 1:
            parts = parts[1:]
        parts = [p for p in parts if p.strip()]
        if not parts:
            return None, None
        line = parts[-1]
        try:
            obj = json.loads(line.decode("utf-8", errors="replace"))
            if isinstance(obj, dict) and obj.get("uuid"):
                return str(obj["uuid"]), lines
        except ValueError:
            pass
        return "sha1:" + hashlib.sha1(line).hexdigest() + ":" + str(lines), lines
    except Exception:
        return None, None


# Tiêu đề mục: tên HOA đúng, đứng đầu dòng (cho phép trước nó khoảng trắng gồm NBSP, #, *, -, số thứ tự "1." hoặc "1)"),
# theo sau là dấu hai chấm (":" hoặc "：" toàn khổ; cho phép khoảng trắng hoặc * trước dấu hai chấm).
# Phân biệt hoa thường có chủ ý: nhãn con viết thường trong giá trị một mục ("Yêu cầu gốc:") không phải tiêu đề.
_PRE = r"^[ \t\u00a0#*\-]*(?:\d+[.)][ \t\u00a0*]*)?"
_COLON = r"[ \t\u00a0*]*[:\uff1a]"
_TAIL = r"[ \t\u00a0*]*"


def header_value(prompt, name):
    """Nội dung sau 'TÊN:' đến tiêu đề kế tiếp đứng đầu dòng hoặc hết văn bản; None nếu không có tiêu đề."""
    text = norm(prompt)
    m = re.search("(?m)" + _PRE + re.escape(norm(name)) + _COLON + _TAIL, text)
    if not m:
        return None
    # Tìm tiêu đề kế tiếp bằng pos (không cắt chuỗi): "^" không khớp tại vị trí bắt đầu tìm,
    # nên tiêu đề nằm ngay sau dấu hai chấm của chính mục này không bị tính là mục mới.
    nxt = re.compile(
        "(?m)" + _PRE + "(?:" + "|".join(re.escape(norm(h)) for h in ALL_HEADERS) + ")" + _COLON
    ).search(text, m.end())
    return (text[m.end(): nxt.start()] if nxt else text[m.end():]).strip()


def check(tool_input):
    errors = []
    agent = tool_input.get("subagent_type")
    prompt = norm(tool_input.get("prompt", ""))
    model = str(tool_input.get("model") or "").lower()
    if agent not in ROSTER:
        errors.append(f"Loại agent '{agent}' không thuộc đội hình. Chọn một trong: {', '.join(sorted(ROSTER))}.")
        return agent, errors
    needed = REQUIRED + EXTRA.get(agent, [])
    for h in needed:
        v = header_value(prompt, h)
        if v is None:
            errors.append(f"Thiếu mục '{h}:' trong lệnh giao việc.")
        elif len(v) < MIN_LEN.get(h, 3):
            errors.append(f"Mục '{h}:' để trống.")
    if agent == "qc":
        lvl = header_value(prompt, "MỨC QC")
        if lvl is not None and not re.match(r"^[12]\b", lvl):
            errors.append("Mục 'MỨC QC:' phải bắt đầu bằng 1 hoặc 2.")
        if lvl is not None and lvl.startswith("2"):
            turn = header_value(prompt, "LƯỢT")
            if turn is None or not re.match(r"^[12]\b", turn):
                errors.append("QC mức 2 cần mục 'LƯỢT:' bắt đầu bằng 1 hoặc 2.")
    if model:
        if not (agent == "qc" and "sonnet" in model):
            errors.append("Không đổi model theo lần gọi, trừ qc chuyển sang Sonnet. Bỏ tham số model.")
    haiku_bound = agent in HAIKU or "haiku" in model
    if haiku_bound:
        est = math.ceil(len(prompt) / CHARS_PER_TOKEN)
        if est > MAX_HAIKU_TOKENS:
            errors.append(
                f"Lệnh giao cho Haiku dài khoảng {est} token, vượt giới hạn {MAX_HAIKU_TOKENS}. "
                "Chỉ đưa đường dẫn và khoảng dòng, không dán nội dung; hoặc giao cho agent Sonnet.")
    return agent, errors


def clean_input(data):
    """Trả về (data, tool_input) khi input dùng được, ngược lại None (hook trả 0, không chặn).
    Không dùng được: data không phải dict; tool_input có mà không phải dict; prompt, subagent_type, description có mà không phải chuỗi.
    Thiếu hoặc null thì giữ hành vi gốc (coi như rỗng)."""
    if not isinstance(data, dict):
        return None
    ti = data.get("tool_input")
    if ti is None:
        ti = {}
    if not isinstance(ti, dict):
        return None
    for k in ("prompt", "subagent_type", "description"):
        if ti.get(k) is not None and not isinstance(ti[k], str):
            return None
    return data, ti


def run():
    try:
        data = json.load(sys.stdin)
    except Exception:
        return 0
    cleaned = clean_input(data)
    if cleaned is None:
        return 0
    data, ti = cleaned
    agent, errors = check(ti)
    tail, nlines = transcript_tail(data.get("transcript_path"))
    rec = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event": "delegation",
        "case": os.environ.get("DELEGATION_CASE", ""),
        "agent": agent,
        "description": ti.get("description", ""),
        "prompt_chars": len(ti.get("prompt", "") or ""),
        "valid": not errors,
        "errors": errors,
        "tool_use_id": data.get("tool_use_id"),
        "session_id": data.get("session_id"),
        "message_id": find_message_id(data.get("transcript_path"), data.get("tool_use_id")),
        "transcript_path": data.get("transcript_path"),
        "transcript_tail": tail,
        "transcript_lines": nlines,
    }
    if os.environ.get("DELEGATION_DRYRUN") == "1":
        rec["event"] = "dryrun"
        write_log(rec)
        sys.stderr.write("CHẾ ĐỘ CHẠY THỬ: đã ghi nhận lựa chọn agent của bạn. Dừng lại, không gọi lại, báo ngắn lý do chọn.\n")
        return 2
    write_log(rec)
    if errors:
        sys.stderr.write("Lệnh giao việc bị chặn:\n- " + "\n- ".join(errors) + "\nSửa rồi giao lại.\n")
        return 2
    return 0


def main():
    """Chỉ trả 2 khi chặn có chủ ý; mọi input hỏng hoặc lỗi bất ngờ trả 0 (Claude Code coi mã khác 0 và khác 2 là lỗi không chặn)."""
    try:
        return run()
    except Exception as e:
        try:
            write_log({"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "event": "guard_warning",
                       "error": type(e).__name__})
        except Exception:
            pass
        return 0


if __name__ == "__main__":
    sys.exit(main())

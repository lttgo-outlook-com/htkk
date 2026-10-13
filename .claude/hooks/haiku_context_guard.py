#!/usr/bin/env python3
"""Chặn lệnh gọi công cụ của agent chạy trên Haiku khi kích thước prompt gần ngưỡng 100.000 token.

Haiku 5.5 tính giá bậc cao (gấp 5 lần) cho cả yêu cầu khi prompt vượt 100.000 token.
Giới hạn áp cho prompt của chính subagent Haiku, không áp cho phiên chính. Vì vậy hook chỉ đo
bản ghi của subagent (thư mục <phiên>/subagents/agent-<agent_id>.jsonl cạnh bản ghi chính).
Input không có agent_id (công cụ chạy trong phiên chính) thì cho qua. Chưa có bản ghi của
subagent (thường ở lệnh gọi công cụ đầu tiên) thì cho qua và ghi nhật ký, không đo bản ghi
chính vì phiên chính dài sẽ chặn oan subagent mới nhận việc nhỏ.
Hook này đọc kích thước prompt của yêu cầu gần nhất từ bản ghi hội thoại (trường usage
của tin nhắn trợ lý) và trả mã thoát 2 để chặn lệnh gọi công cụ kế tiếp khi kích thước
đạt ngưỡng chặn. Agent khi đó phải trả kết quả hiện có thay vì đọc thêm.
Công cụ SubagentHandback (kênh báo cáo cuối) luôn được cho qua để agent báo cáo được.

Cưỡng chế "mỗi lượt một công cụ" (mặc định tắt): bật bằng biến môi trường HAIKU_ONE_TOOL_PER_TURN=1
(chỉ giá trị "1" bật; không đặt hoặc giá trị khác thì tắt). Khi tắt, hook không đọc dấu lượt, không chờ
bản ghi, không tạo thư mục guard-state và chỉ chạy phần chặn theo ngưỡng token. Quy tắc một công cụ mỗi lượt
khi tắt chỉ còn bằng văn bản trong Explore.md và clerk.md.
Khi bật: mọi lệnh gọi trong cùng một message assistant mang cùng dấu lượt
(message.id của message chứa tool_use_id của lệnh, tra trong bản ghi subagent; hook chờ tối đa 1 giây
cho dòng đó xuất hiện). Claude Code ghi mỗi khối của message thành một dòng riêng cùng message.id,
nên các lệnh song song cùng dấu. Không có tool_use_id hoặc không thấy dòng thì cho qua. Hook lưu trạng thái
(dấu lượt, số lệnh đã thấy) theo agent_id tại <CLAUDE_PROJECT_DIR>/.claude/guard-state/<agent_id>.json.
Lệnh đầu của lượt cho qua; hook chặn lệnh thứ hai trở đi (exit 2). Hook không chặn SubagentHandback và không tính nó vào đếm.
Chưa có usage (message đầu của subagent): khi tắt thì cho qua; khi bật thì lấy model từ message chứa lệnh
và vẫn cưỡng chế một công cụ mỗi lượt với Haiku.

Biến môi trường: HAIKU_CONTEXT_BLOCK_TOKENS (mặc định 65000); HAIKU_ONE_TOOL_PER_TURN (mặc định tắt);
đặt HAIKU_CONTEXT_GUARD_DEBUG thành bất kỳ giá trị khác rỗng (kể cả 0) để ghi dòng LUOT (dấu lượt, ms chờ, số đếm)
để chẩn đoán; chỉ có tác dụng khi bật HAIKU_ONE_TOOL_PER_TURN=1, và chỉ ở các lệnh tới bước kiểm lượt. Hook trả mã 0 hoặc 2 trước bước đó thì không ghi LUOT: không có agent_id,
không có bản ghi subagent, lỗi OSError khi đọc, model không phải Haiku, prompt đạt ngưỡng (chặn hoặc cho
SubagentHandback qua), SubagentHandback, lỗi khi tra dấu lượt.
Hook chỉ tác động tới tin nhắn có model chứa "haiku". Không có agent_id, lỗi OSError khi đọc bản ghi hoặc
thiếu tool_use_id thì cho qua mà không ghi. Không tìm thấy bản ghi subagent thì cho qua và ghi nhật ký;
SubagentHandback khi prompt đạt ngưỡng cũng cho qua và ghi CHO_QUA.
"""
import json
import os
import re
import sys
import time

DEFAULT_LIMIT = 65000
TAIL_BYTES = 1_500_000
POLL_MAX = 1.0
POLL_STEP = 0.03
POLL_TAIL_BYTES = 400_000
STATE_DIRNAME = "guard-state"
STATE_MAX_AGE = 24 * 3600
STATE_CLEAN_MAX = 20
ONE_TOOL = os.environ.get("HAIKU_ONE_TOOL_PER_TURN") == "1"
PARALLEL_MSG = "Mỗi lượt chỉ gọi một công cụ. Gọi lại từng công cụ một ở lượt sau.\n"


def _resolve_log(raw):
    if not raw:
        return ""
    raw = os.path.expanduser(raw)
    if os.path.isabs(raw):
        return raw
    return os.path.join(os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd(), raw)


LOG = _resolve_log(os.environ.get("HAIKU_CONTEXT_GUARD_LOG", ""))


def log(msg):
    if LOG:
        try:
            with open(LOG, "a", encoding="utf-8") as f:
                f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n")
        except (OSError, ValueError):
            pass


def _load_limit():
    raw = os.environ.get("HAIKU_CONTEXT_BLOCK_TOKENS")
    if raw is None:
        return DEFAULT_LIMIT
    try:
        v = int(raw.strip())
        if v > 0:
            return v
    except ValueError:
        pass
    log(f"CANH_BAO HAIKU_CONTEXT_BLOCK_TOKENS={raw!r} khong hop le, dung mac dinh {DEFAULT_LIMIT}")
    return DEFAULT_LIMIT


LIMIT = _load_limit()


def transcript_for(data):
    tp = data.get("transcript_path")
    aid = data.get("agent_id")
    if not isinstance(tp, str) or not tp or not aid:
        return None
    cand = os.path.join(os.path.splitext(tp)[0], "subagents", f"agent-{aid}.jsonl")
    return cand if os.path.exists(cand) else None


def last_assistant_usage(path):
    with open(path, "rb") as f:
        f.seek(0, 2)
        size = f.tell()
        f.seek(max(0, size - TAIL_BYTES))
        tail = f.read().decode("utf-8", "ignore")
    for line in reversed(tail.splitlines()):
        try:
            obj = json.loads(line)
        except ValueError:
            continue
        msg = obj.get("message") if isinstance(obj, dict) else None
        if isinstance(msg, dict) and msg.get("role") == "assistant" and isinstance(msg.get("usage"), dict):
            return str(msg.get("model", "")), msg["usage"]
    return None, None


def _find_tool_use(path, tool_use_id, tail_bytes):
    """Trả (message.id, model) của message assistant chứa tool_use có id `tool_use_id`, hoặc (None, "")."""
    with open(path, "rb") as f:
        f.seek(0, 2)
        size = f.tell()
        f.seek(max(0, size - tail_bytes))
        tail = f.read().decode("utf-8", "ignore")
    if tool_use_id not in tail:
        return None, ""
    for line in reversed(tail.splitlines()):
        if tool_use_id not in line:
            continue
        try:
            obj = json.loads(line)
        except ValueError:
            continue
        msg = obj.get("message") if isinstance(obj, dict) else None
        if not isinstance(msg, dict) or msg.get("role") != "assistant":
            continue
        mid = msg.get("id")
        content = msg.get("content")
        if isinstance(mid, str) and mid and isinstance(content, list) and any(
                isinstance(c, dict) and c.get("type") == "tool_use" and c.get("id") == tool_use_id
                for c in content):
            return mid, str(msg.get("model", ""))
    return None, ""


def turn_mark(path, tool_use_id):
    """Trả (dấu lượt, model của message đó, mili giây chờ). Dấu là message.id của message chứa lệnh gọi này.

    Hook chạy trước khi Claude Code ghi dòng tool_use của chính lệnh vào bản ghi, nên chờ ngắn
    (tối đa POLL_MAX giây) cho tới khi dòng đó xuất hiện. Không thấy thì trả (None, ms): cho qua,
    không đoán dấu từ message cũ (message cũ gây chặn oan lệnh đầu của lượt sau).
    """
    if not isinstance(tool_use_id, str) or not tool_use_id:
        return None, "", 0
    t0 = time.time()
    while True:
        mid, mdl = _find_tool_use(path, tool_use_id, POLL_TAIL_BYTES)
        ms = int((time.time() - t0) * 1000)
        if mid:
            return mid, mdl, ms
        if time.time() - t0 >= POLL_MAX:
            return None, "", ms
        time.sleep(POLL_STEP)


def state_dir():
    base = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    return os.path.join(base, ".claude", STATE_DIRNAME)


def clean_state(d):
    """Xóa tối đa STATE_CLEAN_MAX tệp .json, .lock hoặc .tmp cũ hơn 24 giờ (theo mtime), chỉ trong guard-state.

    Tệp .lock không cập nhật mtime nên bị xóa sau 24 giờ dù agent còn chạy; chấp nhận vì agent chạy ngắn hơn 24 giờ.
    """
    now = time.time()
    removed = 0
    try:
        names = os.listdir(d)
    except OSError:
        return
    for name in names:
        if removed >= STATE_CLEAN_MAX:
            break
        if not (name.endswith((".json", ".lock")) or ".tmp" in name):
            continue
        fp = os.path.join(d, name)
        try:
            if os.path.islink(fp) or not os.path.isfile(fp):
                continue
            if now - os.path.getmtime(fp) > STATE_MAX_AGE:
                os.remove(fp)
                removed += 1
        except OSError:
            continue


def parallel_blocked(aid, mark):
    """Ghi nhận lệnh trong lượt `mark`; trả (True nếu là lệnh thứ hai trở đi của lượt, số đếm). Lỗi thì (False, 0)."""
    lock = None
    try:
        if not isinstance(aid, str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", aid):
            return False, 0
        d = state_dir()
        os.makedirs(d, exist_ok=True)
        clean_state(d)
        fp = os.path.join(d, aid + ".json")
        try:
            import fcntl
            lock = open(os.path.join(d, aid + ".lock"), "a")
            t0 = time.time()
            while True:
                try:
                    fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                    break
                except OSError:
                    if time.time() - t0 > 1.0:
                        return False, 0
                    time.sleep(0.01)
        except (ImportError, OSError):
            lock = None  # không khóa được: chấp nhận tranh chấp
        count = 0
        try:
            with open(fp, encoding="utf-8") as f:
                st = json.load(f)
            if isinstance(st, dict) and st.get("mark") == mark and isinstance(st.get("count"), int):
                count = st["count"]
        except (OSError, ValueError):
            count = 0
        count += 1
        tmp = f"{fp}.tmp{os.getpid()}"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"mark": mark, "count": count, "ts": time.time()}, f)
            os.replace(tmp, fp)
        except OSError:
            try:
                os.remove(tmp)
            except OSError:
                pass
        return count >= 2, count
    except Exception:
        return False, 0
    finally:
        if lock is not None:
            try:
                lock.close()
            except OSError:
                pass


def prompt_tokens(usage):
    keys = ("input_tokens", "cache_read_input_tokens", "cache_creation_input_tokens")
    total = 0
    for k in keys:
        try:
            total += int(usage.get(k) or 0)
        except (TypeError, ValueError, OverflowError):
            continue
    return total


def main():
    try:
        data = json.load(sys.stdin)
    except ValueError:
        return 0
    if not isinstance(data, dict):
        return 0
    aid = data.get("agent_id")
    if not aid:
        return 0
    try:
        path = transcript_for(data)
    except (TypeError, ValueError, OSError):
        path = None
    if not path:
        log(f"KHONG_TIM_THAY_TRANSCRIPT_SUBAGENT cho_qua agent={aid}")
        return 0
    try:
        model, usage = last_assistant_usage(path)
    except OSError:
        return 0
    if usage and "haiku" not in model.lower():
        return 0
    if not usage and not ONE_TOOL:
        return 0
    n = prompt_tokens(usage) if usage else 0
    if usage and n >= LIMIT:
        if data.get("tool_name") == "SubagentHandback":
            log(f"CHO_QUA tool=SubagentHandback agent={data.get('agent_id')} prompt={n} gioi_han={LIMIT}")
            return 0
        log(f"CHAN agent={data.get('agent_id')} prompt={n} gioi_han={LIMIT}")
        sys.stderr.write(
            f"Ngân sách ngữ cảnh đã đạt {n} token (giới hạn {LIMIT}). "
            "Không gọi thêm công cụ. Trả ngay kết quả tốt nhất đang có, "
            "và ghi một dòng bắt đầu bằng NÂNG MỨC nếu chưa đủ để hoàn thành.\n"
        )
        return 2
    if data.get("tool_name") == "SubagentHandback":
        return 0
    if not ONE_TOOL:
        return 0
    dbg = bool(os.environ.get("HAIKU_CONTEXT_GUARD_DEBUG"))
    try:
        mark, mark_model, ms = turn_mark(path, data.get("tool_use_id"))
    except (OSError, TypeError, ValueError):
        return 0
    if not usage and "haiku" not in mark_model.lower():
        return 0  # chưa có usage (message đầu của subagent): lấy model từ chính message chứa lệnh
    if not mark:
        if dbg:
            log(f"LUOT agent={aid} khong_co_dau tool={data.get('tool_name')} cho_ms={ms} "
                f"tool_use_id={data.get('tool_use_id')!r} keys={sorted(data)}")
        return 0
    blocked, count = parallel_blocked(aid, mark)
    if dbg:
        log(f"LUOT agent={aid} dau={mark} tool={data.get('tool_name')} cho_ms={ms} dem={count}")
    if blocked:
        log(f"CHAN_SONG_SONG agent={aid}")
        sys.stderr.write(PARALLEL_MSG)
        return 2
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        code = 0
    sys.exit(code)

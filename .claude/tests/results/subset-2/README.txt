Ngày: 2026-10-09
Tập ca: c02,c04,c05,c06,c07,c08,c10,c13,c14,c15,c18 (ROUTING_ONLY; 11 ca bù của bộ 18 ca, cùng subset-1 hợp thành 18), 3 lần.
Kết quả: 31/33 đúng; 10 ca đạt 3/3; c15 đạt 1/3 (run 2 và 3: Explore thay vì researcher). Lệnh không hợp lệ: 0. Exit code 1 (do c15).
Chạy trên coordinator.md cuối và delegation_guard.py đã sửa header_value.
Chấm lại: python3 -I .claude/tests/routing_report.py --runs .claude/tests/results/subset-2/routing-run-1.jsonl .claude/tests/results/subset-2/routing-run-2.jsonl .claude/tests/results/subset-2/routing-run-3.jsonl
(cần transcript trong ~/.claude/projects, có thể đã bị dọn; thêm --expect-all không áp dụng cho tập con)

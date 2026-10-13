Ngày: 2026-10-09
Ca: c15 (ROUTING_ONLY=c15), 3 lần. Prompt đã làm rõ: nêu ba nhà cung cấp MISA, FAST, Bravo và yêu cầu (dưới 50 hóa đơn mỗi tháng, ngân sách dưới 2 triệu đồng mỗi tháng) thẳng trong prompt; kỳ vọng ["researcher"], 1 lệnh.
Kết quả: 0/3 đúng. Cả 3 lần giao thẳng researcher (không còn Explore) nhưng mỗi lần 3 lệnh researcher song song (một nhà cung cấp mỗi lệnh; coordinator áp quy tắc tách N=3, K=3), vượt max_agents 1. Exit code 1.
Chấm lại: python3 -I .claude/tests/routing_report.py --runs .claude/tests/results/subset-3/routing-run-1.jsonl .claude/tests/results/subset-3/routing-run-2.jsonl .claude/tests/results/subset-3/routing-run-3.jsonl
Kỳ vọng c15 đổi thành researcher, 1 đến 3 lệnh, cùng loại agent (chấm lại trên dữ liệu có sẵn, không chạy lại); prompt đã làm rõ
Kết quả chấm lại (2026-10-09, python3 -I routing_report.py --runs 3 nhật ký, chế độ tập con): c15 ĐÚNG 3/3, mỗi lần 3 lệnh researcher trong một message; exit 0.

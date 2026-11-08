# exam-pdf-to-excel-quiz

Agent Skill (dùng được với Claude Code, Codex CLI, Gemini CLI, GitHub Copilot, Cursor, OpenCode…) để **tách phần trắc nghiệm từ file PDF đề thi** (đề in, đề có khoanh đáp án, trang review Moodle, ảnh chụp bài làm…) và gom vào **một file Excel chung**. Câu có hình được chèn ảnh ngay trong ô, kèm đáp án đúng, nguồn đáp án và ghi chú. Skill cũng lọc được các câu trùng giữa nhiều đề.

Mỗi lần bạn gửi một đề, agent thêm đề đó vào cùng một file Excel, rồi báo số câu vừa thêm, tổng số câu và những câu nên xem lại.

## Kết quả đầu ra

Một file `.xlsx`, sheet `Trắc nghiệm`, gồm các cột:

| Đề | Câu | Câu hỏi | A | B | C | D | E | Đáp án đúng | Nguồn đáp án | Ghi chú |
|----|-----|---------|---|---|---|---|---|-------------|--------------|---------|

- **Hình trong đề bài** được chèn trong ô `Câu hỏi`, nằm dưới phần chữ.
- **Phương án là hình**: mỗi phương án được cắt riêng và chèn vào đúng cột A/B/C/D.
- **Vết khoanh bút đỏ** được xoá khỏi hình để hình không lộ đáp án.
- **Câu nhiều đáp án** ghi dạng `A, D`. **Câu kéo thả / chọn từ danh sách / nhập số** ghi ngắn ở cột đáp án, chi tiết để ở `Ghi chú`.
- **Câu cần xem lại** có ghi chú bắt đầu bằng `Nên xem lại:` và ô Ghi chú được tô vàng.
- Phần **tự luận** bị bỏ qua.
- Khi lọc trùng, các câu bị bỏ được liệt kê ở sheet `Câu trùng đã xóa`.

Cột `Nguồn đáp án` nhận một trong các giá trị sau:

| Giá trị | Ý nghĩa |
|---------|---------|
| `Đáp án tham chiếu` | Đề có bảng đáp án, hoặc trang review Moodle ghi "The correct answer is…" |
| `Khoanh sẵn` | Đáp án được khoanh bút đỏ, tô vàng hoặc tick trong đề |
| `Bài làm đạt 100%` | Lấy từ bài làm được chấm điểm tối đa |
| `Theo đáp án <đề khác>` | Câu trùng với một đề khác đã có đáp án |
| `AI chọn` | Các bản đánh dấu mâu thuẫn nhau, agent chọn đáp án đúng |
| `AI giải` | Không có đáp án, hoặc bài làm chưa chắc đúng, nên agent tự giải |

## Cài đặt

Skill theo chuẩn mở [Agent Skills](https://agentskills.io) (một thư mục có `SKILL.md`), nên dùng được với mọi agent hỗ trợ skill.

### 1. Thư viện Python

Agent cần các thư viện này để đọc PDF và ghi Excel:

```bash
pip install -r requirements.txt
```

(Gồm `pymupdf`, `openpyxl`, `numpy` và `pillow`. Bản sao nằm ở `exam-pdf-to-excel-quiz/requirements.txt` để agent tự cài khi cần.) Không bắt buộc cài [LibreOffice](https://www.libreoffice.org/); nếu có, agent sẽ render file Excel ra PDF để tự kiểm tra ảnh có lệch hay không.

### 2. Cài skill

**Cách nhanh (mọi agent)**: dùng CLI [`skills`](https://github.com/vercel-labs/skills) (cần Node.js). CLI sẽ hỏi bạn cài cho agent nào.

```bash
npx skills add minhpham1512/exam-pdf-to-excel-quiz          # cài vào dự án hiện tại
npx skills add minhpham1512/exam-pdf-to-excel-quiz -g       # cài cho cả máy (thư mục người dùng)
npx skills add minhpham1512/exam-pdf-to-excel-quiz -g -a codex -a gemini-cli   # chọn sẵn agent
```

**Cài thủ công**: chép thư mục `exam-pdf-to-excel-quiz/` (thư mục có `SKILL.md`) vào thư mục skills của agent.

| Agent | Cho cả máy | Chỉ trong một dự án |
|-------|------------|---------------------|
| Claude Code | `~/.claude/skills/` | `.claude/skills/` |
| Codex CLI | `~/.agents/skills/` (bản cũ: `~/.codex/skills/`) | `.agents/skills/` |
| Gemini CLI | `~/.gemini/skills/` | `.gemini/skills/` |
| GitHub Copilot | `~/.copilot/skills/` | `.github/skills/` |
| Cursor | `~/.cursor/skills/` | `.cursor/skills/` |
| OpenCode | `~/.config/opencode/skills/` | `.opencode/skills/` |

Đường dẫn có thể thay đổi theo phiên bản agent. Nếu agent không nhận skill, xem tài liệu của agent đó hoặc dùng `npx skills`.

```bash
git clone https://github.com/minhpham1512/exam-pdf-to-excel-quiz.git
# macOS / Linux, ví dụ cho Codex CLI
mkdir -p ~/.agents/skills
cp -r exam-pdf-to-excel-quiz/exam-pdf-to-excel-quiz ~/.agents/skills/
```

```powershell
# Windows (PowerShell), ví dụ cho Claude Code
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
Copy-Item -Recurse exam-pdf-to-excel-quiz\exam-pdf-to-excel-quiz "$env:USERPROFILE\.claude\skills\"
```

**Agent chưa hỗ trợ skill**: clone repo rồi mở agent trong thư mục repo; file [AGENTS.md](AGENTS.md) chỉ agent đọc `SKILL.md`. Hoặc dán nội dung `SKILL.md` vào phần hướng dẫn (custom instructions / rules) của agent.

**Ứng dụng chat cho phép tải skill lên (ví dụ claude.ai)**: nén thư mục `exam-pdf-to-excel-quiz/` (có `SKILL.md` ở gốc) thành `.zip` rồi tải lên phần Skills trong cài đặt.

```bash
cd exam-pdf-to-excel-quiz && zip -r ../exam-pdf-to-excel-quiz.zip exam-pdf-to-excel-quiz
```

## Cách dùng

1. Gửi file kèm một câu yêu cầu (hoặc gọi skill theo tên nếu agent có lệnh gọi skill, ví dụ `/exam-pdf-to-excel-quiz`):
   > Tách phần trắc nghiệm trong `C:\Users\ban\Downloads\de_giua_ky.pdf` sang Excel.
2. Agent khảo sát file, nhận dạng loại đề (khoanh sẵn, review Moodle, nhiều mã đề…), cắt hình, ghi dữ liệu rồi chạy `scripts/build.py` để dựng file Excel.
3. Gửi tiếp đề sau, agent sẽ thêm vào **cùng file Excel**:
   > Thêm đề này: `de_cuoi_ky_20231.pdf`
4. Khi đã gửi đủ các đề, yêu cầu lọc trùng:
   > Lọc các câu trùng giữa các đề.
5. Mở Excel, lọc cột `Ghi chú` theo `Nên xem lại:` để rà các câu agent đánh dấu.

Nên nói cho agent biết bạn muốn lưu file Excel ở đâu; nếu không nói, file nằm ở `work/trac_nghiem.xlsx`.

Agent cần xem được ảnh (đa số agent hiện nay đều được) để thấy hình, vết khoanh và chọn toạ độ cắt. Nếu agent không xem được ảnh, nó sẽ dựa vào phần chữ của PDF và hỏi bạn về các câu có hình.

Mẹo:
- Đề có nhiều mã (Lẻ/Chẵn, Kíp 3/4/5) được tách thành các `Đề` riêng, ví dụ `Final 20221 (Lẻ)`.
- Nếu cùng một đề được chụp nhiều lần với cách đánh dấu khác nhau, agent sẽ gộp lại và nêu rõ chỗ mâu thuẫn.
- Agent không lấy lựa chọn của sinh viên làm đáp án nếu bài làm chưa đạt 100%.

## Cấu trúc repo

```
├── exam-pdf-to-excel-quiz/      # thư mục skill, chép thư mục này để cài
│   ├── SKILL.md                 # hướng dẫn cho agent (quy trình, quy ước, định dạng dữ liệu)
│   ├── requirements.txt         # thư viện Python skill cần
│   └── scripts/build.py         # gom data/*.json thành file Excel
├── AGENTS.md                    # chỉ dẫn cho agent không hỗ trợ skill
├── examples/
│   ├── data/vi-du.json          # ví dụ định dạng dữ liệu một đề
│   └── dups.json                # ví dụ danh sách câu trùng
├── requirements.txt
└── LICENSE
```

## Chạy `build.py` thủ công

Agent ghi dữ liệu mỗi đề vào `work/data/<slug>.json`, rồi chạy `build.py` để gom tất cả thành Excel. Bạn cũng có thể tự sửa JSON và chạy lại:

```bash
python exam-pdf-to-excel-quiz/scripts/build.py --work work --out trac_nghiem.xlsx
```

- `--work`: thư mục chứa `data/*.json` và `dups.json` (tuỳ chọn). Ảnh cắt tạm được lưu vào `<work>/crops/`.
- `--out`: file Excel đầu ra, mặc định là `<work>/trac_nghiem.xlsx`.

### Định dạng `data/<slug>.json`

```json
{"de": "Final 20232", "slug": "final20232", "pdf": "C:/duong/dan/de.pdf",
 "qs": [[1, "câu hỏi", "A", "B", "C", "D", "đáp án", "ghi chú", null, "Khoanh sẵn", "E (tuỳ chọn)"]]}
```

Mỗi câu là một mảng: `[số câu, câu hỏi, A, B, C, D, đáp án, ghi chú, IMG, nguồn đáp án, E]`. Nếu bỏ trống nguồn đáp án thì mặc định là `Khoanh sẵn`.

`IMG` là `null` hoặc một trong ba dạng sau:

| Dạng | Dùng khi | Ý nghĩa |
|------|----------|---------|
| `{"fig": [trang, xref, k], "crops": {"Q": [x0,y0,x1,y1], "A": [...]}}` | Đề in có khoanh đỏ | Render ảnh nhúng thứ `k` ở 200 dpi, xoá màu đỏ, rồi cắt theo toạ độ pixel |
| `{"page": trang, "rects": {"Q": [x0,y0,x1,y1], ...}}` | Trang web in ra (Moodle), hình vẽ vector | Render trực tiếp vùng trang theo toạ độ point của PDF |
| `{"xrefs": {"Q": xref, "A": xref, ...}}` | Câu kéo thả, phương án là ảnh riêng | Lấy thẳng ảnh nhúng theo xref; nền trong suốt được dán lên nền trắng |

Khoá `Q` là hình của đề bài; `A`–`E` là hình của phương án (ô chữ tương ứng sẽ để trống). Số trang tính từ 0. Xem ví dụ đầy đủ trong [examples/data/vi-du.json](examples/data/vi-du.json).

### Định dạng `dups.json`

```json
{"<Đề>": {"<số câu>": "<Đề giữ lại> câu <n>"}}
```

`build.py` sẽ bỏ các câu này khỏi sheet chính và liệt kê chúng ở sheet `Câu trùng đã xóa`, giữ nguyên số câu gốc để tra ngược. Xem [examples/dups.json](examples/dups.json).

## Xử lý sự cố

- **Agent không nhận skill**: kiểm tra thư mục skill nằm đúng chỗ (xem bảng ở trên) và trong đó có `SKILL.md`, rồi khởi động lại agent.
- **Ảnh lệch hàng trong Excel**: `build.py` đã neo ảnh bằng `OneCellAnchor` và đặt chiều cao hàng tường minh. Nếu bạn sửa tay, đừng để Excel tự co giãn chiều cao hàng.
- **Ảnh cắt bị cụt hộp hoặc mũi tên**: chỉnh lại toạ độ `crops`/`rects` trong JSON rồi chạy lại `build.py`.
- **Sai xref với ảnh lặp lại**: dùng chỉ số `k` trong `fig`, hoặc chuyển sang `rects` theo toạ độ trang.
- **PDF dài**: agent đọc từng nhóm trang; quá trình sẽ lâu hơn nhưng vẫn ra một file Excel.

## Skill liên quan

`excel-quiz-to-ms-forms` dùng file Excel do skill này tạo ra để làm bài quiz Microsoft Forms có chấm điểm. Skill đó không nằm trong repo này.

## Giấy phép

[MIT](LICENSE)

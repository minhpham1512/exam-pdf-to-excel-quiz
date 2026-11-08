---
name: exam-pdf-to-excel-quiz
description: Tách câu hỏi trắc nghiệm từ file PDF đề thi (kể cả câu có hình) vào một file Excel có ảnh trong ô, cột đáp án đúng, ghi chú và lọc câu trùng. Use when the user wants to extract multiple-choice questions from exam PDFs into an Excel (.xlsx) quiz bank with embedded images, answer keys and duplicate removal.
---

# Tách trắc nghiệm từ PDF đề thi sang Excel

Dùng khi người dùng gửi (từng) file PDF đề thi / đề ôn / bài làm và muốn gom **chỉ phần trắc nghiệm** vào **một file Excel chung**, mỗi đề thêm vào cùng file. Trao đổi bằng tiếng Việt, trình bày gọn.

Skill gồm:
- `SKILL.md` (file này): quy trình và quy ước.
- `scripts/build.py`: gom tất cả dữ liệu JSON thành file Excel. Đường dẫn trong file này tính từ thư mục chứa `SKILL.md` (gọi là `<skill>`).
- `requirements.txt`: thư viện Python cần cài.

## Quy ước đầu ra (người dùng đã chốt)
- 1 file `.xlsx` duy nhất, sheet `Trắc nghiệm`, cột: `Đề | Câu | Câu hỏi | A | B | C | D | E | Đáp án đúng | Nguồn đáp án | Ghi chú`.
- **Hình trong đề bài** → chèn ảnh ngay trong ô `Câu hỏi`, dưới phần chữ.
- **Phương án là hình** → cắt riêng từng phương án, chèn vào đúng cột A/B/C/D (ô để trống chữ). Nếu có phương án chữ lẫn hình thì giữ chữ.
- **Không** tạo cột đường dẫn / thư mục ảnh rời.
- Xóa vết khoanh bút đỏ trên hình để hình không lộ đáp án.
- Bỏ qua phần tự luận.
- Câu nhiều đáp án: ghi `A, D`. Câu kéo thả / chọn từ danh sách thả xuống / nhập số: cột đáp án ghi ngắn (hoặc `Xem ghi chú`), cách đặt/chọn chi tiết ghi ở `Ghi chú`; với kéo thả thì ô Câu hỏi là sơ đồ gốc còn trống, các cột A–D là các mảnh kéo thả.
- `Nguồn đáp án` dùng một trong: `Đáp án tham chiếu` (đề có bảng đáp án / trang review Moodle có "The correct answer is"), `Khoanh sẵn` (bút đỏ / tô vàng / tick), `Bài làm đạt 100%`, `Theo đáp án <đề khác>` (câu trùng đề đã có đáp án), `AI chọn` (các bản đánh dấu mâu thuẫn, agent chọn đáp án đúng), `AI giải` (không có đáp án, hoặc bài làm của sinh viên chưa chắc đúng — KHÔNG lấy lựa chọn của sinh viên làm đáp án).
- Câu nghi ngờ (đáp án khoanh có vẻ sai, 2 bản đánh dấu khác nhau, đề lỗi): ghi chú bắt đầu bằng `Nên xem lại:` → ô Ghi chú tự tô vàng. Ghi rõ lý do và đáp án đề xuất.
- Sửa lỗi đề gốc hiển nhiên (nhãn phương án trùng, thiếu hình) và nói trong Ghi chú.
- Sau mỗi file: đưa file Excel cho người dùng (đính kèm nếu môi trường hỗ trợ gửi file, nếu không thì báo đường dẫn đầy đủ), báo ngắn: số câu thêm, tổng số câu, các câu cần xem lại, hỏi file tiếp theo.

## Thư mục làm việc
- `work/` (gọi là `<work>`): đặt cạnh file PDF hoặc trong thư mục hiện tại; giữ nguyên qua các lần thêm đề để mọi đề nằm chung một Excel.
  - `<work>/data/<slug>.json`: dữ liệu từng đề.
  - `<work>/dups.json`: (tuỳ chọn) danh sách câu trùng.
  - `<work>/crops/`: ảnh cắt tạm, `build.py` tự tạo.
- File Excel: hỏi người dùng muốn lưu ở đâu nếu chưa rõ; mặc định `<work>/trac_nghiem.xlsx`.

## Quy trình cho mỗi PDF
1. Lần đầu: `pip install -r <skill>/requirements.txt` (pymupdf, openpyxl, numpy, pillow). Nếu pip báo môi trường được quản lý (externally-managed) thì dùng virtualenv hoặc thêm `--break-system-packages`.
2. Khảo sát bằng pymupdf: số trang, `page.get_text()`, `get_image_info(xrefs=True)` (bbox, xref), số drawings (vết khoanh đỏ là drawing màu ~(0.9,0.1,0.1)). Render trang ra PNG (`get_pixmap(dpi=110)`) và **mở ảnh xem** bằng công cụ đọc ảnh của agent để thấy hình, đáp án khoanh/tô/tick, lựa chọn đã chọn. Nếu agent không xem được ảnh: dựa vào text layer, và hỏi người dùng đáp án/toạ độ cho các câu có hình thay vì đoán.
3. Nhận dạng loại file:
   - Đề in (Word/PDF) có khoanh đỏ → `Khoanh sẵn`.
   - Có nhiều mã đề (Lẻ/Chẵn, Kíp 3/4/5) → mỗi mã một `Đề` riêng, ví dụ `Final 20221 (Lẻ)`.
   - Trang review Moodle/daotao.ai ("The correct answer is…", "Grade 100%") → đáp án tham chiếu; text layer chứa đáp án. Trang chụp nhiều lần cùng một đề → gộp, đối chiếu các bản đánh dấu.
   - Ảnh chụp đang làm bài (chưa chấm) → tự giải (`AI giải`).
4. Với hình: dựng lưới (vẽ đường kẻ mỗi 50px lên ảnh render 200 dpi) rồi xem để xác định toạ độ cắt từng phương án / phần hình của đề bài (bỏ nhãn "A." và phần chữ phương án nếu đã chép thành text). Kiểm tra lại crop sau khi build, đặc biệt crop bị cụt hộp/mũi tên.
5. Ghi dữ liệu vào `<work>/data/<slug>.json` (định dạng dưới), rồi chạy:
   ```bash
   python <skill>/scripts/build.py --work <work> --out <file.xlsx>
   ```
   Script in ra số câu và số câu trùng đã xoá. Nếu có LibreOffice, render kiểm tra bằng `soffice --headless --convert-to pdf <file.xlsx>` rồi xem các trang có đề mới (ảnh có lệch hàng / bị cụt không).
6. Tự kiểm tra đáp án: với đáp án khoanh sẵn/tham chiếu vẫn soát nhanh; câu thấy sai thì giữ đáp án nguồn nhưng ghi `Nên xem lại:`; khi 2 nguồn mâu thuẫn thì chọn đáp án đúng (`AI chọn`) và ghi rõ 2 nguồn chọn gì.

## Định dạng dữ liệu `<work>/data/<slug>.json`
```json
{"de":"Final 20232","slug":"final20232","pdf":"/đường/dẫn/file.pdf",
 "qs":[[số_câu, "câu hỏi", "A","B","C","D", "đáp án", "ghi chú", IMG, "nguồn đáp án", "E (tùy chọn)"]]}
```
- Ghi file UTF-8. Đường dẫn `pdf` là tuyệt đối (trên Windows dùng `/` hoặc `\\`).
- Bỏ trống nguồn đáp án thì mặc định `Khoanh sẵn`. Số trang tính từ 0.

`IMG` là `null` hoặc một trong:
- `{"fig":[trang, xref, k], "crops":{"Q":[x0,y0,x1,y1],"A":[...],...}}` — render vùng ảnh nhúng thứ k (mặc định 0) ở 200 dpi, xoá màu đỏ, cắt theo toạ độ pixel. Dùng cho đề in có khoanh đỏ.
- `{"page":trang, "rects":{"Q":[x0,y0,x1,y1],...}}` — render trực tiếp vùng trang (toạ độ point PDF). Dùng cho trang web in ra (Moodle), hình vẽ bằng vector.
- `{"xrefs":{"Q":xref,"A":xref,...}}` — lấy thẳng ảnh nhúng theo xref (nền trong suốt → dán lên nền trắng). Dùng cho câu kéo thả (sơ đồ gốc + các mảnh) và ảnh từng phương án.

Khoá `Q` = hình của đề bài; `A`–`E` = hình phương án (ô chữ tương ứng tự để trống). Nếu `"crops"` giống nhau cho nhiều câu có thể dùng lại.

## build.py làm gì
- Đọc mọi `<work>/data/*.json` (theo thứ tự tên file), bỏ các câu có trong `dups.json`, dựng sheet `Trắc nghiệm` và (nếu có câu trùng) sheet `Câu trùng đã xóa`.
- Ảnh neo bằng `OneCellAnchor` + chiều cao hàng đặt tường minh (không đặt thì LibreOffice/Excel hiển thị ảnh lệch hàng). Cột đáp án `wrap_text=True`, chữ đỏ đậm.
- Ô Ghi chú chứa "xem lại" được tô vàng.
- Không cần sửa script cho từng đề; chỉ sửa JSON rồi chạy lại.

## Lọc câu trùng (khi người dùng yêu cầu)
- Quét: chuẩn hoá câu hỏi (bỏ dấu câu, bỏ "(chọn tất cả)"), so khớp chính xác + so gần đúng (`difflib` ratio > 0.85), so bộ phương án (không phân biệt thứ tự). Hash ảnh (dHash) chỉ để gợi ý — ảnh nhỏ dễ dương tính giả, luôn kiểm tra bằng mắt.
- Là trùng: cùng nội dung + cùng bộ phương án (kể cả đảo thứ tự, diễn đạt hơi khác như "tiến trình/quy trình hợp nhất", "hình khối biểu diễn CSD" ≈ "ký hiệu biểu diễn ca sử dụng").
- KHÔNG trùng: câu hỏi ngược nghĩa (có/không, chức năng/phi chức năng), khác chi tiết (vị trí x/y, ngăn thuộc tính/phương thức), cùng lời nhưng hình khác (sơ đồ trạng thái đổi sự kiện), bộ phương án khác.
- Giữ bản có nguồn đáp án tin cậy nhất: Đáp án tham chiếu > Khoanh sẵn/Bài làm 100% > Theo đáp án… > AI chọn/AI giải; nếu bản khoanh sai thì giữ bản khoanh đúng.
- Ghi danh sách vào `<work>/dups.json` dạng `{"<Đề>": {"<số câu>": "<Đề giữ lại> câu <n>"}}` — `build.py` tự bỏ các câu này và tạo sheet `Câu trùng đã xóa`. Giữ nguyên số câu gốc để tra ngược. Khi thêm đề mới sau này, quét trùng lại.

## Lưu ý
- PDF dài: đọc / render theo từng nhóm trang (ví dụ 5–10 trang một lần) thay vì cả file.
- Ghép nhiều crop thành một ảnh (montage) để xem nhanh một lần.
- `get_image_info` có thể báo sai xref cho ảnh lặp → dùng chỉ số k (lần xuất hiện) hoặc `rects` theo toạ độ trang.
- Không bịa đáp án cho câu thiếu hình/thiếu phương án; ghi chú rõ thiếu gì.

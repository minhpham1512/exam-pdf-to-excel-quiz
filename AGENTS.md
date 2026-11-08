# AGENTS.md

Repo này chứa một Agent Skill: tách phần trắc nghiệm từ file PDF đề thi vào một file Excel có ảnh trong ô, cột đáp án đúng, ghi chú và lọc câu trùng.

Khi người dùng yêu cầu tách / thêm câu trắc nghiệm từ PDF đề thi sang Excel, hoặc lọc câu trùng giữa các đề:

1. Đọc và làm theo [exam-pdf-to-excel-quiz/SKILL.md](exam-pdf-to-excel-quiz/SKILL.md). Trong đó `<skill>` là thư mục `exam-pdf-to-excel-quiz/`.
2. Cài thư viện: `pip install -r exam-pdf-to-excel-quiz/requirements.txt`.
3. Dựng Excel bằng `python exam-pdf-to-excel-quiz/scripts/build.py --work work --out <file.xlsx>`.

Không commit file PDF đề thi, thư mục `work/` hay file `.xlsx` (đã có trong `.gitignore`).

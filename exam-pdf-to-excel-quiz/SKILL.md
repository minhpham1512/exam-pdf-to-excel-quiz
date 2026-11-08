---
name: "exam-pdf-to-excel-quiz"
description: "Tách câu hỏi trắc nghiệm từ file PDF đề thi (kể cả câu có hình) vào một file Excel có ảnh trong ô, cột đáp án đúng, ghi chú và lọc câu trùng."
---

# Tách trắc nghiệm từ PDF đề thi sang Excel

Dùng khi người dùng gửi (từng) file PDF đề thi / đề ôn / bài làm và muốn gom **chỉ phần trắc nghiệm** vào **một file Excel chung**, mỗi đề thêm vào cùng file. Trao đổi bằng tiếng Việt, trình bày gọn.

## Quy ước đầu ra (người dùng đã chốt)
- 1 file `.xlsx` duy nhất, sheet `Trắc nghiệm`, cột: `Đề | Câu | Câu hỏi | A | B | C | D | E | Đáp án đúng | Nguồn đáp án | Ghi chú`.
- **Hình trong đề bài** → chèn ảnh ngay trong ô `Câu hỏi`, dưới phần chữ.
- **Phương án là hình** → cắt riêng từng phương án, chèn vào đúng cột A/B/C/D (ô để trống chữ). Nếu có phương án chữ lẫn hình thì giữ chữ.
- **Không** tạo cột đường dẫn / thư mục ảnh rời.
- Xóa vết khoanh bút đỏ trên hình để hình không lộ đáp án.
- Bỏ qua phần tự luận.
- Câu nhiều đáp án: ghi `A, D`. Câu kéo thả / chọn từ danh sách thả xuống / nhập số: cột đáp án ghi ngắn (hoặc `Xem ghi chú`), cách đặt/chọn chi tiết ghi ở `Ghi chú`; với kéo thả thì ô Câu hỏi là sơ đồ gốc còn trống, các cột A–D là các mảnh kéo thả.
- `Nguồn đáp án` dùng một trong: `Đáp án tham chiếu` (đề có bảng đáp án / trang review Moodle có "The correct answer is"), `Khoanh sẵn` (bút đỏ / tô vàng / tick), `Bài làm đạt 100%`, `Theo đáp án <đề khác>` (câu trùng đề đã có đáp án), `Claude chọn` (các bản đánh dấu mâu thuẫn), `Claude giải` (không có đáp án, hoặc bài làm của sinh viên chưa chắc đúng — KHÔNG lấy lựa chọn của sinh viên làm đáp án).
- Câu nghi ngờ (đáp án khoanh có vẻ sai, 2 bản đánh dấu khác nhau, đề lỗi): ghi chú bắt đầu bằng `Nên xem lại:` → ô Ghi chú tự tô vàng. Ghi rõ lý do và đáp án đề xuất.
- Sửa lỗi đề gốc hiển nhiên (nhãn phương án trùng, thiếu hình) và nói trong Ghi chú.
- Sau mỗi file: gửi lại Excel (SendUserFile), báo ngắn: số câu thêm, tổng số câu, các câu cần xem lại, hỏi file tiếp theo.

## Quy trình cho mỗi PDF
1. `pip install pymupdf openpyxl --break-system-packages` (lần đầu).
2. Khảo sát: số trang, `page.get_text()`, `get_image_info(xrefs=True)` (bbox, xref), số drawings (vết khoanh đỏ là drawing màu ~(0.9,0.1,0.1)). Render trang (`get_pixmap(dpi=110)`) và **xem ảnh** để thấy hình, đáp án khoanh/tô/tick, lựa chọn đã chọn.
3. Nhận dạng loại file:
   - Đề in (Word/PDF) có khoanh đỏ → `Khoanh sẵn`.
   - Có nhiều mã đề (Lẻ/Chẵn, Kíp 3/4/5) → mỗi mã một `Đề` riêng, ví dụ `Final 20221 (Lẻ)`.
   - Trang review Moodle/daotao.ai ("The correct answer is…", "Grade 100%") → đáp án tham chiếu; text layer chứa đáp án. Trang chụp nhiều lần cùng một đề → gộp, đối chiếu các bản đánh dấu.
   - Ảnh chụp đang làm bài (chưa chấm) → tự giải.
4. Với hình: dựng lưới (vẽ đường kẻ mỗi 50px lên ảnh render 200 dpi) rồi xem để xác định toạ độ cắt từng phương án / phần hình của đề bài (bỏ nhãn "A." và phần chữ phương án nếu đã chép thành text). Kiểm tra lại crop sau khi build, đặc biệt crop bị cụt hộp/mũi tên.
5. Ghi dữ liệu vào `work/data/<slug>.json` (xem định dạng dưới), chạy `build.py`, render kiểm tra bằng `soffice --headless --convert-to pdf` rồi xem trang có đề mới.
6. Tự kiểm tra đáp án: với đáp án khoanh sẵn/tham chiếu vẫn soát nhanh; câu thấy sai thì giữ đáp án nguồn nhưng ghi `Nên xem lại:`; khi 2 nguồn mâu thuẫn thì chọn đáp án đúng và ghi rõ 2 nguồn chọn gì.

## Định dạng dữ liệu `work/data/<slug>.json`
```json
{"de":"Final 20232","slug":"final20232","pdf":"/đường/dẫn/file.pdf",
 "qs":[[số_câu, "câu hỏi", "A","B","C","D", "đáp án", "ghi chú", IMG, "nguồn đáp án", "E (tùy chọn)"]]}
```
`IMG` là `null` hoặc một trong:
- `{"fig":[trang, xref, k], "crops":{"Q":[x0,y0,x1,y1],"A":[...],...}}` — render vùng ảnh nhúng thứ k (mặc định 0) ở 200 dpi, xoá màu đỏ, cắt theo toạ độ pixel. Dùng cho đề in có khoanh đỏ.
- `{"page":trang, "rects":{"Q":[x0,y0,x1,y1],...}}` — render trực tiếp vùng trang (toạ độ point PDF). Dùng cho trang web in ra (Moodle), hình vẽ bằng vector.
- `{"xrefs":{"Q":xref,"A":xref,...}}` — lấy thẳng ảnh nhúng theo xref (nền trong suốt → dán lên nền trắng). Dùng cho câu kéo thả (sơ đồ gốc + các mảnh) và ảnh từng phương án.
Khoá `Q` = hình của đề bài; `A`–`E` = hình phương án (ô chữ tương ứng tự để trống). Nếu `"crops"` giống nhau cho nhiều câu có thể dùng lại.

## build.py (gom tất cả data → Excel)
Đặt ở `work/build.py`; sửa `OUT` cho phù hợp (người dùng tải về từ `/mnt/user-data/outputs/`).
```python
import json,glob,os,math,numpy as np,pymupdf as f
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XI
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor,AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment,Font,PatternFill,Border,Side
from PIL import Image
WORK=os.path.dirname(os.path.abspath(__file__))
OUT='/mnt/user-data/outputs/trac_nghiem.xlsx'; TMP=WORK+'/crops'; os.makedirs(TMP,exist_ok=True)
wb=Workbook(); ws=wb.active; ws.title='Trắc nghiệm'
H=['Đề','Câu','Câu hỏi','A','B','C','D','E','Đáp án đúng','Nguồn đáp án','Ghi chú']
W=[14,6,58,34,34,34,34,22,9,13,34]
ws.append(H)
for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='305496'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
ws.row_dimensions[1].height=32
for i,w in enumerate(W): ws.column_dimensions[chr(65+i)].width=w
colpx=lambda w:int(w*7+5)
thin=Side(style='thin',color='BFBFBF')
def fig(doc,pg,x,k=0):
  bb=[i['bbox'] for i in doc[pg].get_image_info(xrefs=True) if i['xref']==x][k]
  p=doc[pg].get_pixmap(dpi=200,clip=f.Rect(bb)); A=np.frombuffer(p.samples,np.uint8).reshape(p.h,p.w,p.n)[:,:,:3].astype(int).copy()
  red=(A[:,:,0]>150)&(A[:,:,0]-A[:,:,1]>60)&(A[:,:,0]-A[:,:,2]>60); A[red]=255  # xoá vết khoanh đỏ
  return Image.fromarray(A.astype('uint8'))
def cr(F,b): return F.crop((max(0,b[0]),max(0,b[1]),min(F.width,b[2]),min(F.height,b[3])))
SC=[0.55]
def put(im,name,col,row,maxw,top=4):
  path=f'{TMP}/{name}.png'
  w=min(maxw,int(im.width*SC[0])); h=int(im.height*w/im.width); im.save(path)
  xi=XI(path); xi.width,xi.height=w,h
  xi.anchor=OneCellAnchor(_from=AnchorMarker(col=col,colOff=pixels_to_EMU(4),row=row-1,rowOff=pixels_to_EMU(top)),ext=XDRPositiveSize2D(pixels_to_EMU(w),pixels_to_EMU(h)))
  ws.add_image(xi); return top+h+6
DUP=json.load(open(WORK+'/dups.json')) if os.path.exists(WORK+'/dups.json') else {}; REMOVED=[]
r=1
for fn in sorted(glob.glob(WORK+'/data/*.json')):
  D=json.load(open(fn)); doc=f.open(D['pdf'])
  for q in D['qs']:
    if str(q[0]) in DUP.get(D['de'],{}): REMOVED.append((D['de'],q[0],q[1],DUP[D['de']][str(q[0])])); continue
    n,qt,a,b,c,d_,ans,note,img=q[:9]; src=q[9] if len(q)>9 and q[9] else 'Khoanh sẵn'
    r+=1; opts=[a,b,c,d_,q[10] if len(q)>10 else '']
    crops=(img or {}).get('crops') or {k:1 for k in ((img or {}).get('rects') or (img or {}).get('xrefs') or {})}
    for i,k in enumerate('ABCDE'):
      if k in crops: opts[i]=''
    ws.append([D['de'],n,qt,*opts,ans,src,note])
    need=[]
    if crops:
      tag=f"{D['slug']}_{n:02d}"; SC[0]=0.55; F=None; FF=None
      if 'xrefs' in img:
        FF={}; SC[0]=1.2
        for k,x in img['xrefs'].items():
          p=f.Pixmap(doc,x)
          if p.colorspace and p.colorspace.n!=3: p=f.Pixmap(f.csRGB,p)
          im=Image.frombytes('RGBA' if p.alpha else 'RGB',(p.w,p.h),p.samples)
          if p.alpha: bg=Image.new('RGB',im.size,'white'); bg.paste(im,mask=im.split()[3]); im=bg
          FF[k]=im
        crops={k:[0,0,99999,99999] for k in FF}
      elif 'page' in img:
        FF={}
        for k,rc in img['rects'].items():
          p=doc[img['page']].get_pixmap(dpi=200,clip=f.Rect(rc)); FF[k]=Image.frombytes('RGB',(p.w,p.h),p.samples) if p.n==3 else Image.frombytes('RGBA',(p.w,p.h),p.samples).convert('RGB')
        crops={k:[0,0,99999,99999] for k in FF}
      else: F=fig(doc,*img['fig'])
      if 'Q' in crops:
        lines=math.ceil(len(qt)/(W[2]*1.05)); need.append(put(cr(FF['Q'] if FF else F,crops['Q']),tag+'Q',2,r,colpx(W[2])-10,top=lines*20+8))
      for i,k in enumerate('ABCDE'):
        if k in crops: need.append(put(cr(FF[k] if FF else F,crops[k]),tag+k,3+i,r,colpx(W[3+i])-10))
    tl=max(math.ceil(len(str(v))/(W[j]*1.1)) for j,v in enumerate([D['de'],n,qt,*opts,ans,src,note]) if v)
    ws.row_dimensions[r].height=max([tl*15+4]+[x*0.75+4 for x in need])  # đặt chiều cao hàng rõ ràng để ảnh không lệch
    for cc in ws[r]: cc.alignment=Alignment(wrap_text=True,vertical='top'); cc.border=Border(top=thin,bottom=thin,left=thin,right=thin)
    ws.cell(r,9).alignment=Alignment(horizontal='center',vertical='top',wrap_text=True); ws.cell(r,9).font=Font(bold=True,color='C00000')
    if 'xem lại' in note.lower(): ws.cell(r,11).fill=PatternFill('solid',fgColor='FFF2CC')
ws.freeze_panes='C2'; ws.auto_filter.ref=f'A1:K{r}'
ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
if REMOVED:
  w2=wb.create_sheet('Câu trùng đã xóa'); w2.append(['Đề','Câu','Câu hỏi','Trùng với câu được giữ lại'])
  for c in w2[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='305496')
  for x in REMOVED: w2.append(list(x))
  for col,wd in zip('ABCD',[18,6,80,30]): w2.column_dimensions[col].width=wd
  for row in w2.iter_rows(min_row=2):
    for c in row: c.alignment=Alignment(wrap_text=True,vertical='top')
wb.save(OUT); print(r-1,'câu,',len(REMOVED),'câu trùng đã xóa')
```
Mẹo: ảnh neo bằng `OneCellAnchor` + chiều cao hàng đặt tường minh (không đặt thì LibreOffice/Excel hiển thị ảnh lệch hàng). Cột đáp án phải `wrap_text=True`.

## Lọc câu trùng (khi người dùng yêu cầu)
- Quét: chuẩn hoá câu hỏi (bỏ dấu câu, bỏ "(chọn tất cả)"), so khớp chính xác + so gần đúng (`difflib` ratio > 0.85), so bộ phương án (không phân biệt thứ tự). Hash ảnh (dHash) chỉ để gợi ý — ảnh nhỏ dễ dương tính giả, luôn kiểm tra bằng mắt.
- Là trùng: cùng nội dung + cùng bộ phương án (kể cả đảo thứ tự, diễn đạt hơi khác như "tiến trình/quy trình hợp nhất", "hình khối biểu diễn CSD" ≈ "ký hiệu biểu diễn ca sử dụng").
- KHÔNG trùng: câu hỏi ngược nghĩa (có/không, chức năng/phi chức năng), khác chi tiết (vị trí x/y, ngăn thuộc tính/phương thức), cùng lời nhưng hình khác (sơ đồ trạng thái đổi sự kiện), bộ phương án khác.
- Giữ bản có nguồn đáp án tin cậy nhất: Đáp án tham chiếu > Khoanh sẵn/Bài làm 100% > Theo đáp án… > Claude chọn/giải; nếu bản khoanh sai thì giữ bản khoanh đúng.
- Ghi danh sách vào `work/dups.json` dạng `{"<Đề>": {"<số câu>": "<Đề giữ lại> câu <n>"}}` — `build.py` tự bỏ các câu này và tạo sheet `Câu trùng đã xóa`. Giữ nguyên số câu gốc để tra ngược. Khi thêm đề mới sau này, quét trùng lại.

## Lưu ý
- File PDF > 10 trang: dùng Read với `pages` hoặc render bằng pymupdf rồi xem ảnh.
- Ảnh ghép nhiều hình (montage) để xem nhanh nhiều crop một lần.
- `get_image_info` có thể báo sai xref cho ảnh lặp → dùng chỉ số k (lần xuất hiện) hoặc `rects` theo toạ độ trang.
- Không bịa đáp án cho câu thiếu hình/thiếu phương án; ghi chú rõ thiếu gì.
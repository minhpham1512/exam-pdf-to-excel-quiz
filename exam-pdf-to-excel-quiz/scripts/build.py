"""Gom tất cả file data/*.json thành một file Excel trắc nghiệm có ảnh trong ô.

Không phụ thuộc agent nào: agent (hoặc người dùng) chỉ cần ghi JSON rồi chạy script này.

Cách dùng:
    python build.py --work <thư_mục_work> [--out <file.xlsx>]

    <thư_mục_work>/data/*.json   dữ liệu từng đề (xem README / SKILL.md)
    <thư_mục_work>/dups.json     (tuỳ chọn) danh sách câu trùng cần bỏ
    <thư_mục_work>/crops/        ảnh cắt tạm, script tự tạo
"""
import argparse, json, glob, os, math
import numpy as np
import pymupdf as f
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XI
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from PIL import Image

ap = argparse.ArgumentParser(description='Gom data/*.json thành file Excel trắc nghiệm.')
ap.add_argument('--work', default='.', help='Thư mục chứa data/ và dups.json (mặc định: thư mục hiện tại)')
ap.add_argument('--out', default=None, help='File Excel đầu ra (mặc định: <work>/trac_nghiem.xlsx)')
args = ap.parse_args()

WORK = os.path.abspath(args.work)
OUT = os.path.abspath(args.out) if args.out else os.path.join(WORK, 'trac_nghiem.xlsx')
TMP = os.path.join(WORK, 'crops'); os.makedirs(TMP, exist_ok=True)
os.makedirs(os.path.dirname(OUT), exist_ok=True)

def load(path):
    with open(path, encoding='utf-8') as fh:
        return json.load(fh)

wb = Workbook(); ws = wb.active; ws.title = 'Trắc nghiệm'
H = ['Đề', 'Câu', 'Câu hỏi', 'A', 'B', 'C', 'D', 'E', 'Đáp án đúng', 'Nguồn đáp án', 'Ghi chú']
W = [14, 6, 58, 34, 34, 34, 34, 22, 9, 13, 34]
ws.append(H)
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws.row_dimensions[1].height = 32
for i, w in enumerate(W): ws.column_dimensions[chr(65 + i)].width = w
colpx = lambda w: int(w * 7 + 5)
thin = Side(style='thin', color='BFBFBF')

def fig(doc, pg, x, k=0):
    """Render ảnh nhúng thứ k có xref x ở trang pg (200 dpi) và xoá vết khoanh đỏ."""
    bb = [i['bbox'] for i in doc[pg].get_image_info(xrefs=True) if i['xref'] == x][k]
    p = doc[pg].get_pixmap(dpi=200, clip=f.Rect(bb))
    A = np.frombuffer(p.samples, np.uint8).reshape(p.h, p.w, p.n)[:, :, :3].astype(int).copy()
    red = (A[:, :, 0] > 150) & (A[:, :, 0] - A[:, :, 1] > 60) & (A[:, :, 0] - A[:, :, 2] > 60)
    A[red] = 255
    return Image.fromarray(A.astype('uint8'))

def cr(F, b): return F.crop((max(0, b[0]), max(0, b[1]), min(F.width, b[2]), min(F.height, b[3])))

SC = [0.55]
def put(im, name, col, row, maxw, top=4):
    path = os.path.join(TMP, f'{name}.png')
    w = min(maxw, int(im.width * SC[0])); h = int(im.height * w / im.width); im.save(path)
    xi = XI(path); xi.width, xi.height = w, h
    xi.anchor = OneCellAnchor(_from=AnchorMarker(col=col, colOff=pixels_to_EMU(4), row=row - 1, rowOff=pixels_to_EMU(top)),
                              ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
    ws.add_image(xi); return top + h + 6

DUP = load(os.path.join(WORK, 'dups.json')) if os.path.exists(os.path.join(WORK, 'dups.json')) else {}
REMOVED = []
r = 1
for fn in sorted(glob.glob(os.path.join(WORK, 'data', '*.json'))):
    D = load(fn); doc = f.open(D['pdf'])
    for q in D['qs']:
        if str(q[0]) in DUP.get(D['de'], {}):
            REMOVED.append((D['de'], q[0], q[1], DUP[D['de']][str(q[0])])); continue
        n, qt, a, b, c, d_, ans, note, img = q[:9]
        src = q[9] if len(q) > 9 and q[9] else 'Khoanh sẵn'
        r += 1; opts = [a, b, c, d_, q[10] if len(q) > 10 else '']
        crops = (img or {}).get('crops') or {k: 1 for k in ((img or {}).get('rects') or (img or {}).get('xrefs') or {})}
        for i, k in enumerate('ABCDE'):
            if k in crops: opts[i] = ''
        ws.append([D['de'], n, qt, *opts, ans, src, note])
        need = []
        if crops:
            tag = f"{D['slug']}_{n:02d}"; SC[0] = 0.55; F = None; FF = None
            if 'xrefs' in img:
                FF = {}; SC[0] = 1.2
                for k, x in img['xrefs'].items():
                    p = f.Pixmap(doc, x)
                    if p.colorspace and p.colorspace.n != 3: p = f.Pixmap(f.csRGB, p)
                    im = Image.frombytes('RGBA' if p.alpha else 'RGB', (p.w, p.h), p.samples)
                    if p.alpha:
                        bg = Image.new('RGB', im.size, 'white'); bg.paste(im, mask=im.split()[3]); im = bg
                    FF[k] = im
                crops = {k: [0, 0, 99999, 99999] for k in FF}
            elif 'page' in img:
                FF = {}
                for k, rc in img['rects'].items():
                    p = doc[img['page']].get_pixmap(dpi=200, clip=f.Rect(rc))
                    FF[k] = Image.frombytes('RGB', (p.w, p.h), p.samples) if p.n == 3 else \
                        Image.frombytes('RGBA', (p.w, p.h), p.samples).convert('RGB')
                crops = {k: [0, 0, 99999, 99999] for k in FF}
            else:
                F = fig(doc, *img['fig'])
            if 'Q' in crops:
                lines = math.ceil(len(qt) / (W[2] * 1.05))
                need.append(put(cr(FF['Q'] if FF else F, crops['Q']), tag + 'Q', 2, r, colpx(W[2]) - 10, top=lines * 20 + 8))
            for i, k in enumerate('ABCDE'):
                if k in crops:
                    need.append(put(cr(FF[k] if FF else F, crops[k]), tag + k, 3 + i, r, colpx(W[3 + i]) - 10))
        tl = max(math.ceil(len(str(v)) / (W[j] * 1.1)) for j, v in enumerate([D['de'], n, qt, *opts, ans, src, note]) if v)
        ws.row_dimensions[r].height = max([tl * 15 + 4] + [x * 0.75 + 4 for x in need])  # chiều cao hàng tường minh để ảnh không lệch
        for cc in ws[r]:
            cc.alignment = Alignment(wrap_text=True, vertical='top')
            cc.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        ws.cell(r, 9).alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
        ws.cell(r, 9).font = Font(bold=True, color='C00000')
        if 'xem lại' in note.lower(): ws.cell(r, 11).fill = PatternFill('solid', fgColor='FFF2CC')

ws.freeze_panes = 'C2'; ws.auto_filter.ref = f'A1:K{r}'
ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
if REMOVED:
    w2 = wb.create_sheet('Câu trùng đã xóa'); w2.append(['Đề', 'Câu', 'Câu hỏi', 'Trùng với câu được giữ lại'])
    for c in w2[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    for x in REMOVED: w2.append(list(x))
    for col, wd in zip('ABCD', [18, 6, 80, 30]): w2.column_dimensions[col].width = wd
    for row in w2.iter_rows(min_row=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical='top')
wb.save(OUT); print(r - 1, 'câu,', len(REMOVED), 'câu trùng đã xóa ->', OUT)

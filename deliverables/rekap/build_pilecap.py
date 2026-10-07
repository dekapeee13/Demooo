"""Pile-head forces per pile and per combination (GROUP +Icap) for pile-cap reinforcement design.
Usage: python3 build_pilecap.py <source Rekap_Gaya_Dalam_Tiang_GROUP_SPColumn.xlsx> <out.xlsx>"""
import sys, collections
import openpyxl
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

src, out = sys.argv[1], sys.argv[2]
rows = [r for r in openpyxl.load_workbook(src, data_only=True)['Data'].iter_rows(min_row=2, values_only=True) if r[0]]
head = {(r[0], r[1], r[2], int(r[3])): r for r in rows if r[4] == 0}        # z = 0 -> pile head
COND = {'M1': 'Non-liquefied', 'M3': 'Liquefied'}
ORDER = ['ASD1', 'ASD2(max)', 'ASD2(min)', 'ASD3', 'ASD4', 'LRFD1', 'LRFD2(max)', 'LRFD2(min)', 'LRFD3', 'LRFD4']
lckey = lambda lc: (lc.split(' ')[1] if ' ' in lc else '', ORDER.index(lc.split(' ')[0]))
S_ASSUMED = 3.0          # pile spacing in a row (3D) - ASSUMED, confirm with .gp11d
Z_ROW = {1: -4.75, 2: 4.75}   # rows under rig node 1 / node 2 (GROUP Z of the rig loads)

NAVY, GREY, YEL = '1F3864', 'F2F2F2', 'FFF2CC'
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
HF, HFILL = Font(name='Calibri', size=10, bold=True, color='FFFFFF'), PatternFill('solid', fgColor=NAVY)
F, FB = Font(name='Calibri', size=10), Font(name='Calibri', size=10, bold=True)
RED, BLUE = Font(name='Calibri', size=10, bold=True, color='C00000'), Font(name='Calibri', size=10, bold=True, color='0070C0')
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
H1 = Font(name='Calibri', size=13, bold=True, color=NAVY); H2 = Font(name='Calibri', size=11, bold=True, color=NAVY)

def cell(ws, r, c, v, font=F, fill=None, fmt=None, al=C):
    x = ws.cell(r, c, v); x.font, x.alignment, x.border = font, al, B
    if fill: x.fill = PatternFill('solid', fgColor=fill)
    if fmt: x.number_format = fmt
    return x

def hdr(ws, r, c0, labels, h=30):
    for j, t in enumerate(labels):
        x = ws.cell(r, c0 + j, t); x.font, x.fill, x.alignment, x.border = HF, HFILL, C, B
    ws.row_dimensions[r].height = h

def setup(ws, land=True):
    ws.page_setup.orientation = 'landscape' if land else 'portrait'; ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False

row_of = lambda p: 1 if p <= 4 else 2
pos_of = lambda p: (p - 1) % 4 + 1

wb = openpyxl.Workbook()
# ------------------------------------------------------------------ 1. Layout
ws = wb.active; ws.title = 'Layout'
ws['A1'] = 'PILE IDENTIFICATION AND LAYOUT - GROUP MODEL (1 cellar, 8 piles, 4x1 under each rig node)'; ws['A1'].font = H1
ws['A2'] = ('Rows identified from the GROUP axial forces: piles 1-4 = row 1, piles 5-8 = row 2 (axial force changes linearly '
            'along each row). Yellow cells are ASSUMED (no .gp11d with pile coordinates available) - confirm and edit; y/z update.')
ws['A2'].font = Font(name='Calibri', size=9, italic=True, color='C00000')
cell(ws, 4, 1, 'Pile spacing in a row s (m)', FB, al=Alignment(horizontal='left'))
cell(ws, 4, 3, S_ASSUMED, FB, YEL, '0.00')
cell(ws, 5, 1, 'Row 1 at GROUP Z (m)', FB, al=Alignment(horizontal='left')); cell(ws, 5, 3, Z_ROW[1], FB, YEL, '0.00')
cell(ws, 6, 1, 'Row 2 at GROUP Z (m)', FB, al=Alignment(horizontal='left')); cell(ws, 6, 3, Z_ROW[2], FB, YEL, '0.00')
ws.merge_cells('A4:B4'); ws.merge_cells('A5:B5'); ws.merge_cells('A6:B6')
hdr(ws, 8, 1, ['GROUP pile no.', 'Row', 'Position in row', 'GROUP Y (m)', 'GROUP Z (m)',
               'Rig node above row (GROUP load N)', 'WPD cellar A: rig node', 'WPD cellar B: rig node'])
for p in range(1, 9):
    r = 8 + p; rw = row_of(p)
    for j, v in enumerate([p, rw, pos_of(p)], 1): cell(ws, r, j, v)
    cell(ws, r, 4, f'=(C{r}-2.5)*$C$4', fill=YEL, fmt='0.00')
    cell(ws, r, 5, f'=$C${4 + rw}', fill=YEL, fmt='0.00')
    cell(ws, r, 6, f'node {rw} (Z = {Z_ROW[rw]:+.2f} m)')
    cell(ws, r, 7, '3241' if rw == 1 else '3233'); cell(ws, r, 8, '3226' if rw == 1 else '3234')
for j, w in enumerate([12, 8, 10, 11, 11, 22, 16, 16], 1): ws.column_dimensions[L(j)].width = w
notes = ['Notes:',
         'GROUP axes: X = vertical (axial), Y and Z = horizontal. Rig loads: node 1 at Z = -4.75 m, node 2 at Z = +4.75 m; cap weight/inertia at origin.',
         'Real cap 25 x 12 m has 2 cellars (4 rig nodes, 16 piles). It was modelled as 1 cellar (8 piles) with load-case variation:',
         '   WPD: LC "1C-A" = cellar A loads (node 1 = rig 3241, node 2 = rig 3233); LC "1C-B" = cellar B loads (node 1 = rig 3226, node 2 = rig 3234).',
         '   -> for one combination, piles under cellar A take the "1C-A" results and piles under cellar B take the "1C-B" results.',
         '   WPA: one load set per combination (no 1C-A/1C-B split in the GROUP run).',
         'Which row lies under node 1 and the numbering direction along a row must be checked against the GROUP input (.gp11d).']
for i, t in enumerate(notes): ws.cell(19 + i, 1, t).font = FB if i == 0 else F
# layout sketch
fig, ax = plt.subplots(figsize=(6.4, 4.2))
for p in range(1, 9):
    y, z = (pos_of(p) - 2.5) * S_ASSUMED, Z_ROW[row_of(p)]
    ax.add_patch(plt.Circle((y, z), 0.5, fc='#D9E2F3', ec='#1F3864', lw=1.2)); ax.text(y, z, str(p), ha='center', va='center', fontsize=11, weight='bold', color='#1F3864')
for n, z in Z_ROW.items():
    ax.plot(0, z, marker='s', color='#C00000', ms=8); ax.text(0.3, z + (0.75 if z > 0 else -1.0), f'rig node {n}', color='#C00000', fontsize=9)
ax.plot(0, 0, marker='+', color='k', ms=10); ax.text(0.2, 0.25, 'origin (cap load)', fontsize=8)
ax.set_xlabel('GROUP Y (m)'); ax.set_ylabel('GROUP Z (m)'); ax.set_aspect('equal'); ax.grid(color='0.9')
ax.set_xlim(-7, 7); ax.set_ylim(-7, 7)
ax.set_title(f'GROUP model plan - spacing s = {S_ASSUMED:g} m ASSUMED', fontsize=10, color='#C00000')
fig.tight_layout(); png = out.rsplit('.', 1)[0] + '_layout.png'; fig.savefig(png, dpi=130); plt.close(fig)
img = XLImage(png); img.anchor = 'J4'; ws.add_image(img)
setup(ws)

# ------------------------------------------------------------------ 2. matrices per wellpad / model
for wp in ('WPA', 'WPD'):
    for mdl in ('M1', 'M3'):
        ws = wb.create_sheet(f'{wp}_{mdl}')
        ws['A1'] = f'PILE-HEAD FORCES - {wp} {mdl} ({COND[mdl]}), GROUP +Icap'; ws['A1'].font = H1
        ws['A2'] = 'P = axial (+ compression, - tension); M, V = resultant moment / shear at the pile head (z = 0, fixed head into cap); y = head deflection.'
        ws['A2'].font = Font(name='Calibri', size=9, italic=True, color='595959')
        lcs = sorted({k[2] for k in head if k[:2] == (wp, mdl)}, key=lckey)
        r0 = 4
        for qi, (q, lab, fmt) in enumerate(((5, 'Axial force P (kN)', '#,##0.0'), (6, 'Head moment M (kNm)', '#,##0.0'),
                                             (7, 'Head shear V (kN)', '#,##0.0'), (8, 'Head deflection y (mm)', '0.00'))):
            ws.cell(r0, 1, lab).font = H2
            hdr(ws, r0 + 1, 1, ['Load combination', 'Design'] + [f'Pile {p}' for p in range(1, 9)] + (['Sum P (kN)'] if q == 5 else []), 22)
            for i, lc in enumerate(lcs):
                r = r0 + 2 + i
                cell(ws, r, 1, lc, FB, GREY if i % 2 else None, al=Alignment(horizontal='left'))
                cell(ws, r, 2, 'LRFD' if lc.startswith('LRFD') else 'ASD', fill=GREY if i % 2 else None)
                vals = [head[(wp, mdl, lc, p)][q] for p in range(1, 9)]
                for p, v in enumerate(vals, 1):
                    f = RED if (q == 5 and v < 0) else F
                    cell(ws, r, 2 + p, v, f, GREY if i % 2 else None, fmt)
                if q == 5:
                    cell(ws, r, 11, f'=SUM(C{r}:J{r})', FB, GREY if i % 2 else None, '#,##0.0')
            r = r0 + 2 + len(lcs)
            for k, (nm, fn) in enumerate((('max', 'MAX'), ('min', 'MIN'))):
                cell(ws, r + k, 1, f'{nm} (all)', BLUE, 'DDEBF7', al=Alignment(horizontal='left')); cell(ws, r + k, 2, '', fill='DDEBF7')
                for p in range(1, 9):
                    cl = L(2 + p); cell(ws, r + k, 2 + p, f'={fn}({cl}{r0 + 2}:{cl}{r - 1})', BLUE, 'DDEBF7', fmt)
            r0 = r + 4
        ws.column_dimensions['A'].width = 18; ws.column_dimensions['B'].width = 7
        for j in range(3, 12): ws.column_dimensions[L(j)].width = 10
        ws.freeze_panes = 'C4'; setup(ws)

# ------------------------------------------------------------------ 3. governing per pile (LRFD for cap design)
ws = wb.create_sheet('Governing', 1)
ws['A1'] = 'GOVERNING PILE-HEAD FORCES PER PILE (envelope M1 + M3) - use LRFD for cap reinforcement, ASD for checks'; ws['A1'].font = H1
r0 = 3
for wp in ('WPA', 'WPD'):
    ws.cell(r0, 1, wp).font = H2
    hdr(ws, r0 + 1, 1, ['Pile', 'Row', 'Pu max (kN)', 'LC', 'Pu min (kN)', 'LC', 'Mu max (kNm)', 'LC', 'Vu max (kN)', 'LC', 'P max ASD (kN)', 'LC'])
    for p in range(1, 9):
        r = r0 + 1 + p
        ks = [k for k in head if k[0] == wp and k[3] == p]
        lr = [k for k in ks if k[2].startswith('LRFD')]; asd = [k for k in ks if k[2].startswith('ASD')]
        lab = lambda k: f'{k[1]} {k[2]}'
        picks = [max(lr, key=lambda k: head[k][5]), min(lr, key=lambda k: head[k][5]), max(lr, key=lambda k: head[k][6]),
                 max(lr, key=lambda k: head[k][7]), max(asd, key=lambda k: head[k][5])]
        qs = [5, 5, 6, 7, 5]
        cell(ws, r, 1, p, FB); cell(ws, r, 2, row_of(p))
        for j, (k, q) in enumerate(zip(picks, qs)):
            v = head[k][q]
            cell(ws, r, 3 + 2 * j, v, RED if v < 0 else F, fmt='#,##0.0'); cell(ws, r, 4 + 2 * j, lab(k))
    r0 += 12
for j, w in enumerate([6, 6, 11, 20, 11, 20, 11, 20, 11, 20, 11, 20], 1): ws.column_dimensions[L(j)].width = w
ws.cell(r0, 1, 'Red = tension (uplift): cap top reinforcement and pile-to-cap anchorage must carry it.').font = F
setup(ws)


# ------------------------------------------------------------------ 5. equilibrium check: sum of pile-head P vs applied vertical load
# applied = 2 rig nodes x FZ (Rev_Reaction, per node, max of the 2 nodes is equal) + cap weight x factor (GROUP input load N3)
RIG = {'ASD1': 1618.3, 'ASD2(max)': 7679.6, 'ASD2(min)': 804.9, 'ASD3': 1812.5, 'ASD4': 1812.5,
       'LRFD1': 2265.6, 'LRFD2(max)': 11640.0, 'LRFD2(min)': 640.4, 'LRFD3': 2217.1, 'LRFD4': 2217.1}
CAPW = {'ASD1': 5575.5, 'ASD2(max)': 5575.5, 'ASD2(min)': 5575.5, 'ASD3': 6235.6, 'ASD4': 6235.6,
        'LRFD1': 7805.7, 'LRFD2(max)': 6690.6, 'LRFD2(min)': 6690.6, 'LRFD3': 7633.5, 'LRFD4': 7633.5}
ws = wb.create_sheet('Check_SumP', 1)
ws['A1'] = 'EQUILIBRIUM CHECK - sum of pile-head axial forces vs applied vertical load (current rig reactions)'; ws['A1'].font = H1
ws['A2'] = ('Applied = 2 x rig FZ per node (Rev_Reaction, 2 nodes per cellar) + cap weight incl. load factor. '
            'Ratio outside 0.98-1.02 = the GROUP run used a different load -> results of that combination must be re-run.')
ws['A2'].font = Font(name='Calibri', size=9, italic=True, color='595959')
hdr(ws, 4, 1, ['Wellpad', 'Model', 'Load combination', 'Sum P GROUP (kN)', 'Rig FZ per node (kN)', 'Cap weight (kN)',
               'Applied (kN)', 'Ratio', 'Status', 'Rig FZ per node back-calc. (kN)'])
r = 5
for k in sorted({k[:3] for k in head}, key=lambda k: (k[0], k[1], lckey(k[2]))):
    lc = k[2].split(' ')[0]
    sp = sum(head[k + (p,)][5] for p in range(1, 9))
    cell(ws, r, 1, k[0]); cell(ws, r, 2, k[1]); cell(ws, r, 3, k[2], al=Alignment(horizontal='left'))
    cell(ws, r, 4, round(sp, 1), fmt='#,##0.0'); cell(ws, r, 5, RIG[lc], fmt='#,##0.0'); cell(ws, r, 6, CAPW[lc], fmt='#,##0.0')
    cell(ws, r, 7, f'=2*E{r}+F{r}', fmt='#,##0.0'); cell(ws, r, 8, f'=D{r}/G{r}', fmt='0.000')
    cell(ws, r, 9, f'=IF(ABS(H{r}-1)<=0.02,"OK","RE-RUN")', FB)
    cell(ws, r, 10, f'=(D{r}-F{r})/2', fmt='#,##0.0')
    r += 1
from openpyxl.formatting.rule import CellIsRule
ws.conditional_formatting.add(f'I5:I{r - 1}', CellIsRule(operator='equal', formula=['"RE-RUN"'], font=RED, fill=PatternFill('solid', fgColor='FCE4D6')))
for j, w in enumerate([8, 7, 18, 13, 13, 12, 12, 8, 9, 15], 1): ws.column_dimensions[L(j)].width = w
ws.freeze_panes = 'A5'; setup(ws)

# ------------------------------------------------------------------ 4. long list (filterable)
ws = wb.create_sheet('All_Heads')
hdr(ws, 1, 1, ['Wellpad', 'Model', 'Condition', 'Load combination', 'Design', 'Pile', 'Row', 'P (kN)', 'M (kNm)', 'V (kN)', 'y (mm)'])
for i, k in enumerate(sorted(head, key=lambda k: (k[0], k[1], lckey(k[2]), k[3])), 2):
    r = head[k]
    for j, v in enumerate([k[0], k[1], COND[k[1]], k[2], 'LRFD' if k[2].startswith('LRFD') else 'ASD', k[3], row_of(k[3]),
                           r[5], r[6], r[7], r[8]], 1):
        cell(ws, i, j, v, fmt='#,##0.0' if j >= 8 else None)
for j, w in enumerate([8, 7, 13, 18, 7, 6, 6, 10, 10, 10, 9], 1): ws.column_dimensions[L(j)].width = w
ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:K{len(head) + 1}'
wb.save(out)

if __name__ == '__main__':
    # check: pile-head P equals the source envelope (WPA M1 ASD2(max) pile 1 = 7043.2 kN)
    assert head[('WPA', 'M1', 'ASD2(max)', 1)][5] == 7043.2
    assert len(head) == 2 * (5 + 10) * 2 * 8 * 2 // 2, len(head)
    print('ok', out, len(head), 'pile-head rows')

"""Membuat workbook Excel: data uji, analisa BS 8081 (rumus), grafik native & perbandingan pola standar."""
import sys
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

media_dir = sys.argv[1] if len(sys.argv) > 1 else None
out = __file__.replace("buat_excel.py", "Analisa_Uji_Angkur_GA22-P2_BS8081.xlsx")

F = "Arial"
f_in = Font(name=F, color="0000FF")
f_b = Font(name=F, bold=True)
f_n = Font(name=F)
f_h = Font(name=F, bold=True, color="FFFFFF")
f_t = Font(name=F, bold=True, size=14)
fill_h = PatternFill("solid", fgColor="1F4E78")
fill_y = PatternFill("solid", fgColor="FFFF00")
fill_ok = PatternFill("solid", fgColor="C6EFCE")
fill_g = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="A6A6A6")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
C1, C2, C3 = "2A78D6", "EB6834", "1BAF7A"
GREY = "7F7F7F"

wb = Workbook()
ws = wb.active
ws.title = "Data & Analisa"
ws.sheet_properties.tabColor = "1F4E78"

def put(cell, v, font=f_n, fmt=None, fill=None, al=None):
    c = ws[cell]; c.value = v; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if al: c.alignment = al
    return c

put("A1", "Analisa Uji Angkur GA 22. P 2 terhadap BS 8081 (On-site acceptance test)", f_t)
put("A2", "Teks biru = input (dapat diubah); teks hitam = rumus. Semua hasil dihitung ulang otomatis.", Font(name=F, italic=True, color=GREY))

# ---- Input ----
put("A4", "DATA ANGKUR", f_b)
inputs = [
    ("Testing date", "02.10.2026", None, "Dari Anchor_Extension_Record.xlsx"),
    ("Anchor ID", "GA 22. P 2", None, "Dari record"),
    ("Anchor type", "PGA2-R2", None, "Dari record"),
    ("Working load, Tw (kN)", 460, "0.00", "Dari record"),
    ("Free length, Lf (m)", 8, "0.00", "Dari record"),
    ("Fixed / bond length, Lb (m)", 18, "0.00", "Dari record"),
    ("Jumlah strand", 5, "0", "Dari record: 5 nos dia 0.5\""),
    ("Luas 1 strand (mm²)", 98.7, "0.0", "Asumsi: strand 12.7 mm (0.5\") standar, 98.7 mm²"),
    ("Modulus elastisitas Es (kN/mm²)", 195, "0", "Asumsi: nilai tipikal strand 7-wire; ganti dengan sertifikat pabrik"),
    ("Panjang tendon di dalam jack (m)", 0, "0.00", "Asumsi 0 — isi bila diketahui (menambah panjang bebas efektif)"),
    ("Theoretical ext'n lower (mm)", 55, "0", "Dari record"),
    ("Theoretical ext'n upper (mm)", 68, "0", "Dari record"),
]
r0 = 5
for i, (lab, val, fmt, note) in enumerate(inputs):
    r = r0 + i
    put(f"A{r}", lab)
    c = put(f"B{r}", val, f_in, fmt)
    c.comment = Comment(note, "Analisa")
    put(f"C{r}", note, Font(name=F, size=8, color=GREY))
# named rows
TW, LF, LB, NS, AS, ES, LJ = "$B$8", "$B$9", "$B$10", "$B$11", "$B$12", "$B$13", "$B$14"

put("A18", "Luas tendon At (mm²)"); put("B18", f"={NS}*{AS}", fmt="0.0")
put("A19", "Beban datum T0 (kN)"); put("B19", "=B26", fmt="0.00")
put("A20", "Beban uji Tp (kN)"); put("B20", "=B29", fmt="0.00")
put("A21", "Tp / Tw"); put("B21", f"=B20/{TW}", fmt="0.0%")
put("A22", "Perpanjangan teoritis per m (mm/m) = (Tp-T0)·1000/(At·Es)"); put("B22", f"=(B20-B19)*1000/(B18*{ES})", fmt="0.000")

# ---- Raw data table ----
hdr = ["Siklus", "Beban (kN)", "% Tw", "Gauge (psi)", "Tahan (menit)", "Ram ext. (mm)",
       "Perpindahan relatif thd awal (mm)", "Perpindahan dlm siklus (mm)"]
HR = 24
for j, h in enumerate(hdr):
    c = ws.cell(HR, 1 + j, h); c.font = f_h; c.fill = fill_h; c.border = box
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.row_dimensions[HR].height = 42
ws.cell(HR + 1, 1, "Data uji (sumber: Anchor_Extension_Record.xlsx)").font = Font(name=F, italic=True, size=8, color=GREY)
loads = [57.2, 240.22, 469, 697.78, 469, 240.22, 57.2]
psi = [250, 1050, 2050, 3050, 2050, 1050, 250]
holds = ["-", "-", "-", 15, "-", "-", "-"], [1, 1, 1, 15, 1, 1, 1], [1, 1, 1, 15, 1, 1, 1]
ext = [25, 40, 54, 70, 64, 53, 37], [37, 44, 53, 75, 67, 50, 19], [19, 45, 60, 75, 67, 53, 16]
names = ["1st Cycle", "2nd Cycle", "3rd Cycle"]
DR = HR + 2  # first data row = 26
rows_of = []
for k in range(3):
    rows = []
    for i in range(7):
        r = DR + k * 7 + i
        rows.append(r)
        vals = [names[k], loads[i], None, psi[i], holds[k][i], ext[k][i]]
        for j, v in enumerate(vals):
            c = ws.cell(r, 1 + j, v); c.border = box; c.font = f_in if j in (1, 3, 4, 5) else f_n
        ws.cell(r, 3, f"=B{r}/{TW}").number_format = "0.0%"
        ws.cell(r, 7, f"=F{r}-$F${DR}")
        ws.cell(r, 8, f"=F{r}-$F${rows[0] if i else r}")
        for j in (3, 7, 8):
            ws.cell(r, j).font = f_n; ws.cell(r, j).border = box
        ws.cell(r, 2).number_format = "0.00"
        if k % 2:
            for j in range(1, 9): ws.cell(r, j).fill = fill_g
    rows_of.append(rows)
# fix B19/B20 references to actual rows
ws["B19"] = f"=B{DR}"; ws["B20"] = f"=B{DR+3}"
LAST = DR + 20  # 46

# ---- Evaluation table ----
ER = LAST + 3  # 49
put(f"A{ER-1}", "EVALUASI BS 8081 — PANJANG TENDON BEBAS SEMU (apparent free tendon length)", f_b)
eh = ["Siklus", "Ext. di Tp (mm)", "Ext. setelah kembali ke datum (mm)", "Elastik δe (mm)",
      "Permanen kumulatif (mm)", "L_app (m)", "Batas bawah 0.9·Lf (m)", "Batas atas Lf+0.5·Lb (m)", "Status"]
for j, h in enumerate(eh):
    c = ws.cell(ER, 1 + j, h); c.font = f_h; c.fill = fill_h; c.border = box
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.row_dimensions[ER].height = 42
for k in range(3):
    r = ER + 1 + k; rr = rows_of[k]
    ws.cell(r, 1, f"Siklus {k+1}")
    ws.cell(r, 2, f"=F{rr[3]}")
    ws.cell(r, 3, f"=F{rr[6]}")
    ws.cell(r, 4, f"=B{r}-C{r}")
    ws.cell(r, 5, f"=C{r}-$F${DR}")
    ws.cell(r, 6, f"=D{r}/$B$22").number_format = "0.00"
    ws.cell(r, 7, f"=0.9*{LF}+{LJ}").number_format = "0.00"
    ws.cell(r, 8, f"={LF}+0.5*{LB}+{LJ}").number_format = "0.00"
    ws.cell(r, 9, f'=IF(AND(F{r}>=G{r},F{r}<=H{r}),"OK","TIDAK OK")')
    for j in range(1, 10):
        ws.cell(r, j).font = f_n; ws.cell(r, j).border = box
ws.cell(ER + 1, 9).comment = Comment("Siklus 1 biasanya masih mengandung bedding/seating wedge; evaluasi utama memakai siklus terakhir (siklus 3).", "Analisa")

NR = ER + 5
notes = [
    ("Kriteria", "BS 8081 (acceptance test): 0.9·Lf ≤ L_app ≤ Lf + 0.5·Lb, dgn L_app = At·Es·δe/(Tp−T0). Evaluasi utama: siklus terakhir."),
    ("Creep di Tp", "TIDAK DAPAT DIEVALUASI — record hanya 1 bacaan pada akhir tahan 15 menit; standar meminta bacaan berkala (mis. 1, 5, 15 menit) dan creep ≤ 5% ext. elastik."),
    ("Anomali", "Bacaan datum turun antar siklus (37 → 19 → 16 mm) sehingga perpindahan permanen negatif; tidak sesuai pola standar (lihat sheet 'Perbandingan Pola')."),
    ("Kesimpulan", "Kriteria panjang bebas semu: OK (siklus 2 & 3). Kelulusan penuh belum dapat dinyatakan sebelum data creep dilengkapi & anomali dikonfirmasi."),
]
put(f"A{NR-1}", "CATATAN", f_b)
for i, (a, b) in enumerate(notes):
    put(f"A{NR+i}", a, f_b); put(f"B{NR+i}", b, al=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells(f"B{NR+i}:I{NR+i}"); ws.row_dimensions[NR + i].height = 30
ws[f"A{NR+3}"].fill = fill_y; ws[f"B{NR+3}"].fill = fill_y

# ---- Chart helper data (elastic vs load, limit lines) ----
CR = NR + 6
put(f"A{CR-1}", "DATA BANTU GRAFIK — Elastik vs beban (siklus 3) & garis batas", f_b)
for j, h in enumerate(["T−T0 (kN)", "Elastik siklus 3 (mm)", "Batas bawah 0.9Lf (mm)", "Lf teoritis (mm)", "Batas atas Lf+0.5Lb (mm)",
                       "Permanen kumulatif (mm)", "Elastik di Tp (mm)"]):
    c = ws.cell(CR, 1 + j, h); c.font = f_h; c.fill = fill_h; c.border = box
    c.alignment = Alignment(wrap_text=True, horizontal="center")
ws.row_dimensions[CR].height = 30
r3 = rows_of[2]
for i in range(4):
    r = CR + 1 + i
    ws.cell(r, 1, f"=B{r3[i]}-$B$19").number_format = "0.0"
    ws.cell(r, 2, f"=F{r3[i]}-F{r3[6]}")
# garis batas: δ = (T−T0)·L·1000/(At·Es)
for i in range(4):
    r = CR + 1 + i
    ws.cell(r, 3, f"=A{r}*$G${ER+1}*1000/($B$18*{ES})").number_format = "0.0"
    ws.cell(r, 4, f"=A{r}*({LF}+{LJ})*1000/($B$18*{ES})").number_format = "0.0"
    ws.cell(r, 5, f"=A{r}*$H${ER+1}*1000/($B$18*{ES})").number_format = "0.0"
for k in range(3):
    r = CR + 1 + k
    ws.cell(r, 6, f"=E{ER+1+k}"); ws.cell(r, 7, f"=D{ER+1+k}")
    ws.cell(r, 8, f"=$A${CR+4}")  # x for Tp points
ws.cell(CR, 8, "x @Tp (kN)").font = f_h; ws.cell(CR, 8).fill = fill_h
for r in range(CR + 1, CR + 5):
    for j in range(1, 9):
        ws.cell(r, j).font = f_n; ws.cell(r, j).border = box

widths = [30, 13, 15, 13, 13, 13, 16, 16, 12]
for j, w in enumerate(widths):
    ws.column_dimensions[chr(65 + j)].width = w

# ======================= Sheet: Perbandingan Pola =========================
wp = wb.create_sheet("Perbandingan Pola")
wp.sheet_properties.tabColor = "EB6834"
wp["A1"] = "Pola tipikal BS 8081 (angkur permanen, 3 siklus) vs data uji — dinormalisasi"; wp["A1"].font = f_t
wp["A2"] = ("Pola standar didigitasi secara skematik dari gambar BS 8081 'Load increments and minimum periods of observation "
            "for on-site acceptance tests — (b) Permanent anchorage' (gambar yang ada di record). Sumbu perpindahan pada standar "
            "tidak berskala, maka kedua pola dinormalisasi: 0 = bacaan awal, 1 = perpindahan maksimum.")
wp["A2"].font = Font(name=F, italic=True, size=9, color=GREY); wp["A2"].alignment = Alignment(wrap_text=True, vertical="top")
wp.merge_cells("A2:L2"); wp.row_dimensions[2].height = 42
pct = [10, 50, 100, 150, 100, 50, 10]
# x pixel positions digitized from figure (b); origin x=73 px
std_px = [[73, 87, 115, 198, 190, 165, 120],
          [120, 127, 185, 238, 230, 200, 155],
          [155, 157, 200, 275, 268, 240, 215]]
h2 = ["Beban (% Tw)", "Std S1", "Std S2", "Std S3", "Uji S1", "Uji S2", "Uji S3"]
PR = 4
for j, h in enumerate(h2):
    c = wp.cell(PR, 1 + j, h); c.font = f_h; c.fill = fill_h; c.border = box; c.alignment = Alignment(horizontal="center")
wp.cell(PR - 1, 2, "Pola standar (normalisasi, input digitasi)").font = f_b
wp.cell(PR - 1, 5, "Data uji (normalisasi, rumus)").font = f_b
DS = "'Data & Analisa'"
for i in range(7):
    r = PR + 1 + i
    wp.cell(r, 1, pct[i])
    for k in range(3):
        c = wp.cell(r, 2 + k, round((std_px[k][i] - 73) / (275 - 73), 3)); c.font = f_in; c.number_format = "0.00"
        rr = rows_of[k][i]
        c = wp.cell(r, 5 + k, f"=({DS}!F{rr}-MIN({DS}!$F${DR}:$F${LAST}))/(MAX({DS}!$F${DR}:$F${LAST})-MIN({DS}!$F${DR}:$F${LAST}))")
        c.number_format = "0.00"; c.font = f_n
    for j in range(1, 8): wp.cell(r, j).border = box
    wp.cell(r, 1).font = f_in
# Comparison table
TR = PR + 10
wp.cell(TR - 1, 1, "PERBANDINGAN KARAKTERISTIK").font = f_b
th = ["Aspek", "Pola umum BS 8081", "Data uji GA 22. P 2", "Sesuai?"]
for j, h in enumerate(th):
    c = wp.cell(TR, 1 + j, h); c.font = f_h; c.fill = fill_h; c.border = box
ES_ = f"{DS}!"
comp = [
    ("Bentuk kurva pembebanan", "Cekung ke bawah (kekakuan menurun saat beban naik), kurva bongkar lebih curam dari kurva muat (histeresis).",
     "Kurva muat & bongkar membentuk loop histeresis; kurva bongkar lebih curam di siklus 1–3.", "Ya"),
    ("Pergeseran loop antar siklus", "Tiap siklus bergeser ke kanan: perpindahan permanen di datum bertambah (S1 < S2 < S3).",
     f'="Bacaan datum TURUN: "&{ES_}F{rows_of[0][6]}&" → "&{ES_}F{rows_of[1][6]}&" → "&{ES_}F{rows_of[2][6]}&" mm (loop bergeser ke kiri)."', "Tidak"),
    ("Perpindahan di Tp tiap siklus", "Meningkat sedikit tiap siklus (akumulasi permanen + creep).",
     f'={ES_}F{rows_of[0][3]}&" → "&{ES_}F{rows_of[1][3]}&" → "&{ES_}F{rows_of[2][3]}&" mm (naik lalu tetap)."', "Sebagian"),
    ("Perpindahan elastik tiap siklus", "Relatif konsisten antar siklus dan dalam batas 0.9Lf … Lf+0.5Lb.",
     f'={ES_}D{ER+1}&" / "&{ES_}D{ER+2}&" / "&{ES_}D{ER+3}&" mm — siklus 1 jauh lebih kecil; siklus 2 & 3 konsisten dan dalam batas."', "Sebagian"),
    ("Tahanan 15 menit di Tp", "Segmen horizontal kecil (creep) yang tercatat pada beberapa bacaan waktu.",
     "Hanya satu bacaan per tahanan — creep tidak dapat dinilai.", "Tidak dapat dinilai"),
    ("Perpindahan permanen akhir", "Positif dan bertambah; sisa setelah siklus 3 > siklus 1.",
     f'={ES_}E{ER+3}&" mm (negatif) — tidak wajar secara fisik; cek re-stroke ram / slip wedge / pencatatan."', "Tidak"),
]
for i, row in enumerate(comp):
    r = TR + 1 + i
    for j, v in enumerate(row):
        c = wp.cell(r, 1 + j, v); c.font = f_n; c.border = box
        c.alignment = Alignment(wrap_text=True, vertical="top")
    s = wp.cell(r, 4)
    s.fill = fill_ok if row[3] == "Ya" else (fill_y if row[3] in ("Sebagian",) else PatternFill("solid", fgColor="FFC7CE"))
    wp.row_dimensions[r].height = 45
for col, w in zip("ABCDEFGHIJKL", [28, 45, 50, 16, 10, 10, 10, 10, 10, 10, 10, 10]):
    wp.column_dimensions[col].width = w

# ======================= Charts =====================================
def scatter(title, xt, yt, xmin=None, xmax=None, ymin=None, ymax=None, w=17, h=11):
    ch = ScatterChart(); ch.title = title; ch.style = 13
    ch.x_axis.title = xt; ch.y_axis.title = yt
    ch.width, ch.height = w, h
    for ax, lo, hi in ((ch.x_axis, xmin, xmax), (ch.y_axis, ymin, ymax)):
        if lo is not None: ax.scaling.min = lo
        if hi is not None: ax.scaling.max = hi
        ax.delete = False
    ch.legend.position = "b"
    return ch

def add(ch, sheet, xcol, ycol, r1, r2, name, color, dash=None, marker="circle", line=True, size=6):
    x = Reference(sheet, min_col=xcol, min_row=r1, max_row=r2)
    y = Reference(sheet, min_col=ycol, min_row=r1, max_row=r2)
    s = Series(y, x, title=name)
    s.marker = Marker(symbol=marker if marker else "none", size=size)
    if marker:
        s.marker.graphicalProperties.solidFill = color
        s.marker.graphicalProperties.line.solidFill = color
    if line:
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 22000
        if dash: s.graphicalProperties.line.dashStyle = dash
    else:
        s.graphicalProperties.line.noFill = True
    s.smooth = False
    ch.series.append(s)

gs = wb.create_sheet("Grafik")
gs.sheet_properties.tabColor = "1BAF7A"
gs["A1"] = "Grafik uji angkur GA 22. P 2 — format sesuai rekomendasi BS 8081"; gs["A1"].font = f_t

# (a) Load-displacement
ch = scatter("(a) Beban – perpindahan, 3 siklus", "Ram extension (mm)", "Beban (% Tw)", 0, 85, 0, 1.6)
ch.y_axis.number_format = "0%"; ch.y_axis.majorUnit = 0.1
for k, col in enumerate((C1, C2, C3)):
    add(ch, ws, 6, 3, rows_of[k][0], rows_of[k][6], f"Siklus {k+1}", col)
gs.add_chart(ch, "A3")

# (b) elastic vs limits
ch = scatter("(b) Perpindahan elastik & permanen vs batas BS 8081", "Beban di atas datum, T − T0 (kN)", "Perpindahan (mm)", 0, 700, -20, 120)
ch.y_axis.majorUnit = 20; ch.x_axis.majorUnit = 100; ch.x_axis.number_format = "0"
add(ch, ws, 1, 5, CR + 1, CR + 4, "Batas atas Lf+0.5Lb", GREY, dash="dash", marker=None)
add(ch, ws, 1, 4, CR + 1, CR + 4, "Lf teoritis", "BFBFBF", dash="sysDot", marker=None)
add(ch, ws, 1, 3, CR + 1, CR + 4, "Batas bawah 0.9Lf", GREY, dash="dash", marker=None)
add(ch, ws, 1, 2, CR + 1, CR + 4, "Elastik siklus 3", C3)
add(ch, ws, 8, 7, CR + 1, CR + 3, "Elastik di Tp (S1, S2, S3)", C2, marker="diamond", line=False, size=9)
add(ch, ws, 8, 6, CR + 1, CR + 3, "Permanen kumulatif (S1, S2, S3)", "404040", marker="square", line=False, size=8)
gs.add_chart(ch, "L3")

# (c) standard pattern, (d) test normalized
ch = scatter("(c) Pola umum BS 8081 — angkur permanen (skematik)", "Perpindahan (ternormalisasi)", "Beban (% Tw)", 0, 1.05, 0, 160)
ch.y_axis.majorUnit = 10
for k, col in enumerate((C1, C2, C3)):
    add(ch, wp, 2 + k, 1, PR + 1, PR + 7, f"Std siklus {k+1}", col)
gs.add_chart(ch, "A26")
ch = scatter("(d) Data uji — dinormalisasi (skala sama dgn c)", "Perpindahan (ternormalisasi)", "Beban (% Tw)", 0, 1.05, 0, 160)
ch.y_axis.majorUnit = 10
for k, col in enumerate((C1, C2, C3)):
    add(ch, wp, 5 + k, 1, PR + 1, PR + 7, f"Uji siklus {k+1}", col)
gs.add_chart(ch, "L26")

# (e) overlay
ch = scatter("(e) Overlay: pola standar (putus-putus) vs data uji (garis penuh)", "Perpindahan (ternormalisasi)", "Beban (% Tw)", 0, 1.05, 0, 160, w=34, h=12)
ch.y_axis.majorUnit = 10
for k, col in enumerate((C1, C2, C3)):
    add(ch, wp, 2 + k, 1, PR + 1, PR + 7, f"Std S{k+1}", col, dash="dash", marker=None)
    add(ch, wp, 5 + k, 1, PR + 1, PR + 7, f"Uji S{k+1}", col)
gs.add_chart(ch, "A49")

gs["A73"] = "Catatan: grafik (c) adalah digitasi skematik dari gambar standar (sumbu perpindahan pada standar tidak berskala); perbandingan bersifat kualitatif (bentuk & urutan loop)."
gs["A73"].font = Font(name=F, italic=True, size=9, color=GREY)

# Original figures from record
if media_dir:
    gs["A75"] = "Gambar acuan dari record (BS 8081): (a) temporary, (b) permanent anchorage"; gs["A75"].font = f_b
    for p, anchor in (("image1.png", "A76"), ("image2.png", "H76")):
        img = Image(f"{media_dir}/{p}"); gs.add_image(img, anchor)

for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.value is not None and (c.font is None or c.font.name != F):
                c.font = Font(name=F, bold=c.font.b, italic=c.font.i, color=c.font.color, size=c.font.sz)
from openpyxl.worksheet.pagebreak import Break
for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.page_setup.paperSize = sh.PAPERSIZE_A4
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0 if sh.title == "Grafik" else 1
gs.print_area = "A1:W95"
gs.row_breaks.append(Break(id=47))
wb.save(out)
print(out)

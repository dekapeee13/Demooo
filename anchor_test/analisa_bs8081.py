"""Evaluasi uji angkur GA 22. P 2 terhadap BS 8081 (uji penerimaan / acceptance test).

Kriteria panjang tendon bebas semu (apparent free tendon length):
    0.9 * L_free  <=  L_app  <=  L_free + 0.5 * L_bond
    L_app = A_t * E_s * delta_e / (T_p - T_datum)
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# --- Data proyek -------------------------------------------------------------
Tw = 460.0                 # kN
L_free, L_bond = 8.0, 18.0 # m
n_strand, A_strand = 5, 98.7   # mm2 (strand 0.5", 12.7 mm)
Es = 195.0                 # kN/mm2
At = n_strand * A_strand
loads = [57.2, 240.22, 469.0, 697.78, 469.0, 240.22, 57.2]
ext = {
    "Siklus 1": [25, 40, 54, 70, 64, 53, 37],
    "Siklus 2": [37, 44, 53, 75, 67, 50, 19],
    "Siklus 3": [19, 45, 60, 75, 67, 53, 16],
}
T0, Tp = loads[0], loads[3]
dT = Tp - T0
mm_per_m = dT * 1000 / (At * Es)  # mm perpanjangan per meter panjang bebas

L_lo, L_hi = 0.9 * L_free, L_free + 0.5 * L_bond
d_lo, d_hi, d_th = L_lo * mm_per_m, L_hi * mm_per_m, L_free * mm_per_m

print(f"A_t = {At:.1f} mm2, Tp = {Tp:.2f} kN ({Tp/Tw*100:.1f}% Tw), datum = {T0:.1f} kN ({T0/Tw*100:.1f}% Tw)")
print(f"Elastik teoritis: 0.9Lf={d_lo:.1f} mm | Lf={d_th:.1f} mm | Lf+0.5Lb={d_hi:.1f} mm")
rows = []
for k, e in ext.items():
    de = e[3] - e[6]          # perpindahan elastik (di Tp - residu setelah kembali ke datum)
    dp = e[6] - ext["Siklus 1"][0]  # perpindahan permanen kumulatif
    Lapp = de / mm_per_m
    ok = L_lo <= Lapp <= L_hi
    rows.append((k, de, dp, Lapp, ok))
    print(f"{k}: elastik={de} mm, permanen={dp} mm, L_app={Lapp:.2f} m -> {'OK' if ok else 'TIDAK OK'}")

# --- Grafik (format seperti Gambar BS 8081) -----------------------------------
col = {"Siklus 1": "#2a78d6", "Siklus 2": "#eb6834", "Siklus 3": "#1baf7a"}
ink, muted, grid = "#1f1f1e", "#6b6a63", "#e4e3dc"
plt.rcParams.update({"font.size": 10, "axes.edgecolor": muted, "axes.labelcolor": ink,
                     "xtick.color": muted, "ytick.color": muted})
fig, (a1, a2) = plt.subplots(1, 2, figsize=(13, 6.2))

pct = [l / Tw * 100 for l in loads]
holds = ["", "1", "1", "15", "1", "1", "1"]
for k, e in ext.items():
    a1.plot(e, pct, "-o", color=col[k], lw=2, ms=6, mec="white", mew=1.5, label=k)
    for x, y, h in zip(e[1:], pct[1:], holds[1:]):
        if h == "15":
            a1.annotate("15", (x, y), xytext=(0, 7), textcoords="offset points",
                        ha="center", color=ink, fontsize=9)
a1.set_xlabel("Perpindahan / ram extension (mm)")
a1.set_ylabel("Beban (% Tw)")
a1.set_title("(a) Beban – perpindahan, 3 siklus (uji penerimaan)", loc="left", color=ink)
a1.set_yticks(range(0, 161, 10)); a1.set_ylim(0, 160); a1.set_xlim(0, 85)
a1.grid(color=grid, lw=0.8); a1.legend(frameon=False, loc="lower right")

# Elastik & permanen vs beban (siklus 3) + garis batas BS 8081
L = loads[:4]
e3 = ext["Siklus 3"]
elastic = [e3[i] - e3[6] for i in range(4)]
perm = [ext[k][6] - ext["Siklus 1"][0] for k in ext]
a2.plot([0, Tp], [0, d_hi], "--", color=muted, lw=1.2)
a2.plot([0, Tp], [0, d_lo], "--", color=muted, lw=1.2)
a2.plot([0, Tp], [0, d_th], ":", color=muted, lw=1.2)
a2.text(Tp + 8, d_hi, f"Lf + 0,5Lb = {L_hi:.1f} m\n(batas atas, {d_hi:.0f} mm)", va="center", fontsize=8.5, color=muted)
a2.text(Tp + 8, d_th + 3, f"Lf = {L_free:.1f} m ({d_th:.0f} mm)", va="center", fontsize=8.5, color=muted)
a2.text(Tp + 8, d_lo - 3, f"0,9Lf = {L_lo:.1f} m\n(batas bawah, {d_lo:.0f} mm)", va="center", fontsize=8.5, color=muted)
# garis acuan dimulai dari datum
a2.plot([l - T0 for l in L], elastic, "-o", color=col["Siklus 3"], lw=2, ms=7, mec="white", mew=1.5,
        label="Perpindahan elastik (siklus 3)")
for k, (_, de, *_r) in zip(ext, rows):
    a2.plot(dT, de, "D", color=col[k], ms=8, mec="white", mew=1.5, label=f"Elastik di Tp – {k} ({de} mm)")
a2.plot([dT] * 3, perm, "s", mfc="white", mec=ink, ms=7, mew=1.5, label="Perpindahan permanen (kumulatif)")
for v, k in zip(perm, ext):
    a2.annotate(f"{v} mm", (dT, v), xytext=(-8, 0), textcoords="offset points", ha="right", va="center", fontsize=8.5, color=ink)
a2.axhline(0, color=muted, lw=0.8)
a2.set_xlabel("Beban di atas datum, T − T₀ (kN)")
a2.set_ylabel("Perpindahan (mm)")
a2.set_title("(b) Perpindahan elastik & permanen vs batas panjang bebas semu", loc="left", color=ink)
a2.set_xlim(0, 860); a2.set_ylim(-15, 125)
a2.grid(color=grid, lw=0.8); a2.legend(frameon=False, loc="upper left", fontsize=8.5)

fig.suptitle("Uji Angkur GA 22. P 2 (PGA2-R2) – 02.10.2026 – evaluasi BS 8081", x=0.01, ha="left",
             fontsize=13, color=ink, fontweight="bold")
fig.tight_layout()
fig.savefig(__file__.replace("analisa_bs8081.py", "grafik_uji_angkur_GA22-P2.png"), dpi=160)

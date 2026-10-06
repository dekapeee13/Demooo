"""Engineering line illustrations (black/white, template style) for the bearing capacity & settlement sheets."""
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle, FancyArrowPatch

plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'hatch.linewidth': 0.5})
K = 'black'; G = '0.55'
OUT = 'bc/fig'
import os
os.makedirs(OUT, exist_ok=True)


def arrow(ax, x0, y0, x1, y1, lw=0.8, ms=7, style='-|>'):
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle=style, mutation_scale=ms, lw=lw, color=K))


def dim(ax, x0, x1, y, txt, off=0.25, vertical=False):
    if vertical:
        arrow(ax, y, x0, y, x1, style='<|-|>', ms=6, lw=0.6)
        ax.text(y - off, (x0 + x1) / 2, txt, ha='right', va='center')
    else:
        arrow(ax, x0, y, x1, y, style='<|-|>', ms=6, lw=0.6)
        ax.text((x0 + x1) / 2, y + off, txt, ha='center', va='bottom')


def save(fig, name):
    fig.savefig(f'{OUT}/{name}.png', dpi=150, bbox_inches='tight', pad_inches=0.05, facecolor='white')
    plt.close(fig)


# ---------- Figure 1: cross-section of embankment ----------
def fig_section():
    fig, ax = plt.subplots(figsize=(11.5, 3.6))
    ax.set_xlim(-16, 16); ax.set_ylim(-9.5, 4.2); ax.axis('off')
    Hf, s = 2.0, 2.0                       # fill height, slope H:V
    top = 9.0; bot = top + s * Hf
    ax.add_patch(Polygon([(-bot, 0), (-top, Hf), (top, Hf), (bot, 0)], closed=True, fc='0.93', ec=K, lw=1.2, hatch='..'))
    for x in np.linspace(-top + 0.6, top - 0.6, 15):
        arrow(ax, x, Hf + 1.1, x, Hf + 0.05, ms=6)
    ax.plot([-top + 0.3, top - 0.3], [Hf + 1.1, Hf + 1.1], color=K, lw=0.8)
    ax.text(0, Hf + 1.3, 'Uniform surcharge  qs = 100 kPa', ha='center', va='bottom', fontsize=9.5)
    ax.text(0, Hf / 2, 'Compacted embankment fill   Hf, γf', ha='center', va='center', fontsize=9.5, bbox=dict(fc='white', ec='none', pad=1))
    # ground layers
    ax.plot([-16, 16], [0, 0], color=K, lw=1.2)
    ax.add_patch(Rectangle((-bot, -2.0), 2 * bot, 2.0, fc='white', ec=K, lw=0.8, hatch='xx'))
    ax.text(0, -1.0, 'Replacement with compacted granular fill (0 - Dr)', ha='center', va='center', fontsize=9, bbox=dict(fc='white', ec='none', pad=1))
    ax.add_patch(Rectangle((-16, -3.2), 32, 1.2, fc='0.85', ec=K, lw=0.8, hatch='---'))
    ax.text(13.0, -2.6, "Soft silt / clay  (cu, H')", ha='center', va='center', fontsize=8.5, bbox=dict(fc='white', ec='none', pad=0.5))
    ax.add_patch(Rectangle((-16, -9.4), 32, 6.2, fc='white', ec=K, lw=0.8, hatch='...'))
    ax.text(10.5, -6.5, "Medium dense - dense sand\n(N-SPT, φ', E)", ha='center', va='center', fontsize=8.5, bbox=dict(fc='white', ec='none', pad=0.5))
    # GWL
    ax.plot([-8, 2], [-5.2, -5.2], color=K, lw=0.8, ls='--')
    ax.add_patch(Polygon([(-6.2, -5.2), (-5.8, -5.2), (-6.0, -4.85)], closed=True, fc='white', ec=K, lw=0.8))
    ax.text(-5.4, -5.0, 'GWL', fontsize=8.5, bbox=dict(fc='white', ec='none', pad=0.3))
    # dims
    dim(ax, -bot, bot, -9.0, 'B (base width of loaded area)', off=0.15)
    arrow(ax, bot + 0.9, 0, bot + 0.9, Hf, style='<|-|>', ms=6, lw=0.6); ax.text(bot + 1.2, Hf / 2, 'Hf', va='center')
    arrow(ax, -bot - 0.9, 0, -bot - 0.9, -2.0, style='<|-|>', ms=6, lw=0.6); ax.text(-bot - 1.2, -1.0, 'Dr', va='center', ha='right')
    ax.text(bot - 1.0, Hf + 0.25, 'slope 1V : 2H', fontsize=8.5, ha='left')
    save(fig, 'fig1_section')


# ---------- Figure 2: bearing capacity mechanisms ----------
def fig_bc():
    fig, axs = plt.subplots(1, 2, figsize=(11.5, 3.4))
    # (a) general shear in sand (drained)
    ax = axs[0]; ax.set_xlim(-19, 19); ax.set_ylim(-9.0, 2.2); ax.axis('off'); ax.set_aspect('equal')
    ax.plot([-19, 19], [0, 0], color=K, lw=1)
    ax.add_patch(Rectangle((-1.5, 0), 3, 0.35, fc='0.8', ec=K, lw=1))
    for x in np.linspace(-1.3, 1.3, 7):
        arrow(ax, x, 1.2, x, 0.38, ms=6)
    ax.text(0, 1.3, 'q', ha='center', va='bottom')
    phi = np.radians(35); a = np.pi / 4 + phi / 2
    apex = (0, -1.5 * np.tan(a))
    ax.add_patch(Polygon([(-1.5, 0), (1.5, 0), apex], closed=True, fc='0.92', ec=K, lw=0.8))
    ax.text(0, -0.6, 'I', ha='center', fontsize=10)
    b = 1.5; r0 = np.hypot(b, apex[1])
    for sgn in (1, -1):
        th = np.linspace(a - np.pi, -(np.pi / 4 - phi / 2), 40)
        rr = r0 * np.exp((th - th[0]) * np.tan(phi))
        xs = b + rr * np.cos(th); ys = rr * np.sin(th)
        xs = sgn * xs
        ax.plot(xs, ys, color=K, lw=0.9)
        xe, ye = xs[-1], ys[-1]
        Lp = abs(ye) / np.tan(np.pi / 4 - phi / 2)
        ax.plot([xe, xe + sgn * Lp], [ye, 0], color=K, lw=0.9)
        ax.plot([sgn * b, xe], [0, ye], color=G, lw=0.6)
        for t in np.linspace(th[0], th[-1], 4)[1:-1]:
            r_ = r0 * np.exp((t - th[0]) * np.tan(phi)); ax.plot([sgn * b, sgn * (b + r_ * np.cos(t))], [0, r_ * np.sin(t)], color=G, lw=0.4)
        ax.text(sgn * (b + 3.2), -3.2, 'II', ha='center'); ax.text(xe + sgn * Lp * 0.45, ye * 0.3, 'III', ha='center')
    ax.text(0, -9.0, "(a) General shear - granular soil (drained)\nqu = 0.5·γ'·B·Nγ·sγ·rγ", ha='center', va='bottom', fontsize=8.5)
    # (b) squeezing of thin soft layer over firm base
    ax = axs[1]; ax.set_xlim(-6, 6); ax.set_ylim(-4.3, 1.7); ax.axis('off'); ax.set_aspect('equal')
    ax.plot([-6, 6], [0, 0], color=K, lw=1)
    ax.add_patch(Rectangle((-4.5, 0), 9, 0.35, fc='0.8', ec=K, lw=1))
    for x in np.linspace(-4.2, 4.2, 13):
        arrow(ax, x, 1.2, x, 0.38, ms=6)
    ax.text(0, 1.3, 'q  (B >> H\')', ha='center', va='bottom')
    ax.add_patch(Rectangle((-6, -1.3), 12, 1.3, fc='0.88', ec=K, lw=0.8, hatch='---'))
    ax.add_patch(Rectangle((-6, -3.0), 12, 1.7, fc='white', ec=K, lw=0.8, hatch='...'))
    ax.text(5.2, -2.2, 'firm sand', ha='center', fontsize=8, bbox=dict(fc='white', ec='none', pad=0.5))
    for sgn in (1, -1):
        for y in (-0.45, -0.85):
            arrow(ax, sgn * 3.6, y, sgn * 5.6, y, ms=7, lw=0.9)
    arrow(ax, -5.9, 0, -5.9, -1.3, style='<|-|>', ms=5, lw=0.6); ax.text(-5.75, -0.65, "H'", va='center', fontsize=8.5)
    ax.text(0, -4.2, "(b) Squeezing of a thin soft layer (undrained)\nqu = cu·Nc*·sc ,  Nc* = 4.14 + 0.5·B/H'  ≥ 5.14  (Meyerhof 1974)", ha='center', va='bottom', fontsize=8.5)
    save(fig, 'fig2_bearing')


# ---------- Figure 3: lateral squeeze ----------
def fig_squeeze():
    fig, ax = plt.subplots(figsize=(11.5, 3.0))
    ax.set_xlim(-2, 22); ax.set_ylim(-4.6, 3.3); ax.axis('off'); ax.set_aspect('equal')
    ax.add_patch(Polygon([(0, 0), (0, 2), (8, 2), (12, 0)], closed=True, fc='0.93', ec=K, lw=1.2, hatch='..'))
    for x in np.linspace(0.4, 7.6, 9):
        arrow(ax, x, 3.0, x, 2.05, ms=6)
    ax.text(3.6, 3.05, 'qs', ha='center', va='bottom')
    ax.plot([-2, 22], [0, 0], color=K, lw=1.2)
    ax.add_patch(Rectangle((-2, -2.6), 24, 2.6, fc='0.88', ec=K, lw=0.8, hatch='---'))
    ax.add_patch(Rectangle((-2, -4.4), 24, 1.8, fc='white', ec=K, lw=0.8, hatch='...'))
    ax.text(19.5, -3.5, 'firm layer', ha='center', fontsize=8.5, bbox=dict(fc='white', ec='none', pad=0.5))
    ax.text(19.5, -1.3, 'soft layer, cu', ha='center', fontsize=8.5, bbox=dict(fc='white', ec='none', pad=0.5))
    for y in (-0.7, -1.3, -1.9):
        arrow(ax, 7.5, y, 13.5, y, ms=8, lw=1.0)
    ax.plot(np.linspace(12, 16, 30), 0.35 * np.sin(np.linspace(0, np.pi, 30)), color=K, lw=1.0)
    ax.text(14, 0.45, 'heave at toe', ha='center', va='bottom', fontsize=8)
    arrow(ax, 12.8, 0, 12.8, -2.6, style='<|-|>', ms=5, lw=0.6); ax.text(13.0, -2.35, 'Ds', fontsize=8.5)
    arrow(ax, -0.8, 0, -0.8, 2, style='<|-|>', ms=5, lw=0.6); ax.text(-1.0, 1, 'Hf', ha='right', va='center')
    th = np.linspace(np.pi, np.pi - np.arctan(0.5), 20)
    ax.plot(12 + 1.8 * np.cos(th), 1.8 * np.sin(th), color=K, lw=0.6); ax.text(10.0, 0.25, 'θ', fontsize=9)
    ax.text(10, -4.55, "FS = 2·cu / (γf·Ds·tanθ) + 4.14·cu / q_edge   (Silvestri 1983; FHWA-NHI-06-088),   q_edge = γf·Hf + qs",
            ha='center', va='top', fontsize=8.5)
    save(fig, 'fig3_squeeze')


# ---------- Figure 4: stress distribution & layered elastic settlement ----------
def fig_stress():
    fig, axs = plt.subplots(1, 3, figsize=(11.5, 3.6), gridspec_kw=dict(width_ratios=[1.0, 1.3, 1.2]))
    ax = axs[0]; ax.set_xlim(-0.2, 3.4); ax.set_ylim(-0.6, 2.6); ax.axis('off'); ax.set_aspect('equal')
    ax.add_patch(Rectangle((0, 0), 3, 2, fc='0.93', ec=K, lw=1))
    ax.plot([1.5, 1.5], [0, 2], color=K, lw=0.6, ls='--'); ax.plot([0, 3], [1, 1], color=K, lw=0.6, ls='--')
    ax.plot(1.5, 1, 'o', color=K, ms=4)
    ax.text(1.55, 1.08, 'C', fontsize=8)
    for (x, y) in ((0.75, 0.5), (2.25, 0.5), (0.75, 1.5), (2.25, 1.5)):
        ax.text(x, y, 'B/2 × L/2', ha='center', va='center', fontsize=7)
    dim(ax, 0, 3, -0.35, 'L', off=-0.25)
    ax.text(-0.15, 1, 'B', ha='right', va='center')
    ax.text(1.5, 2.15, 'Plan: centre C = 4 × corner\nΔσz = 4·q·I(m,n), m = B/2z, n = L/2z', ha='center', va='bottom', fontsize=7.5)
    # stress vs depth
    ax = axs[1]
    z = np.linspace(0.3, 60, 200); B, L, q = 74, 202, 136
    m, n = B / 2 / z, L / 2 / z; A = m ** 2 + n ** 2 + 1
    I = 1 / (4 * np.pi) * (2 * m * n * np.sqrt(A) / (A + m ** 2 * n ** 2) * (A + 1) / A + np.arctan2(2 * m * n * np.sqrt(A), A - m ** 2 * n ** 2))
    ds = 4 * q * I; s0 = 18.5 * np.minimum(z, 8) + 8.5 * np.maximum(z - 8, 0)
    ax.plot(ds, z, color=K, lw=1.4, label="Δσz (centre)")
    ax.plot(0.2 * s0, z, color=K, lw=1.0, ls='--', label="0.2·σ'v0")
    zc = z[np.argmax(ds < 0.2 * s0)] if np.any(ds < 0.2 * s0) else 60
    ax.axhline(zc, color=G, lw=0.8, ls=':'); ax.text(5, zc - 1.0, 'influence depth', fontsize=7.5, va='bottom')
    ax.set_ylim(60, 0); ax.set_xlim(0, 150); ax.set_xlabel('Stress (kPa)', fontsize=8); ax.set_ylabel('Depth z (m)', fontsize=8)
    ax.tick_params(labelsize=7); ax.legend(fontsize=7, frameon=False, loc='lower right')
    ax.set_title('Boussinesq stress (example B×L = 74×202 m)', fontsize=8)
    for sp_ in ('top', 'right'):
        ax.spines[sp_].set_visible(False)
    # layered elastic
    ax = axs[2]; ax.set_xlim(0, 6); ax.set_ylim(-6.6, 1.2); ax.axis('off')
    ax.add_patch(Rectangle((1.0, 0), 3.6, 0.3, fc='0.8', ec=K))
    for x in np.linspace(1.2, 4.4, 7):
        arrow(ax, x, 1.0, x, 0.33, ms=5)
    ys = [0, -1.4, -2.6, -4.2, -5.4]
    names = ['layer 1 : E1, ν1', 'layer 2 : E2, ν2', 'layer i : Ei, νi', '...']
    for k in range(4):
        ax.add_patch(Rectangle((0.2, ys[k + 1]), 5.2, ys[k] - ys[k + 1], fc='white', ec=K, lw=0.7, hatch=['...', '---', '...', ''][k]))
        ax.text(2.8, (ys[k] + ys[k + 1]) / 2, names[k], ha='center', va='center', fontsize=7.5, bbox=dict(fc='white', ec='none', pad=0.3))
    ax.text(2.8, -6.55, "Si = 4·q·(B/2)·(1-ν²)/E·[ΔF1 + (1-2ν)/(1-ν)·ΔF2]\nSteinbrenner (1934); S = Σ Si", ha='center', va='bottom', fontsize=7.5)
    save(fig, 'fig4_stress')


# ---------- Figure 5: e - log σ' ----------
def fig_elog():
    fig, ax = plt.subplots(figsize=(11.5, 3.1))
    s = np.logspace(1, 3.3, 200); sp = 200.0; e0 = 0.9; s0 = 60.0
    Cr, Cc = 0.05, 0.25
    e = np.where(s <= sp, e0 - Cr * np.log10(s / s0), e0 - Cr * np.log10(sp / s0) - Cc * np.log10(s / sp))
    ax.semilogx(s, e, color=K, lw=1.4)
    sf = 450.0
    for x, lab in ((s0, "σ'v0"), (sp, "σ'p"), (sf, "σ'v0 + Δσz")):
        ax.axvline(x, color=G, lw=0.7, ls='--'); ax.text(x * 1.03, 0.98, lab, fontsize=8, va='top')
    ax.text(100, e0 - Cr * np.log10(100 / s0) + 0.012, 'Cr (recompression)', fontsize=8)
    ax.text(600, e0 - Cr * np.log10(sp / s0) - Cc * np.log10(600 / sp) + 0.015, 'Cc (virgin)', fontsize=8)
    ax.set_xlabel("Effective vertical stress σ'v (kPa, log scale)", fontsize=8); ax.set_ylabel('Void ratio e', fontsize=8)
    ax.set_ylim(0.55, 1.0); ax.tick_params(labelsize=7)
    for sp_ in ('top', 'right'):
        ax.spines[sp_].set_visible(False)
    ax.set_title("Sc = h/(1+e0)·[Cr·log(σ'p/σ'v0) + Cc·log((σ'v0+Δσz)/σ'p)] ;  σ'p = max(cu/0.22 ; σ'v0) ;  t90 = Tv·Hdr²/Cv",
                 fontsize=8.5)
    save(fig, 'fig5_elog')


if __name__ == '__main__':
    fig_section(); fig_bc(); fig_squeeze(); fig_stress(); fig_elog()
    print(sorted(os.listdir(OUT)))

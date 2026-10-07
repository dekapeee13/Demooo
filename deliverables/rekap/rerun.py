"""Load the GROUP rerun outputs (16 piles, 2 cellars) into flat records + back-calculated pile coordinates."""
import glob, os, math
import numpy as np
import gp11o

FILES = {('WPA', 'M1'): 'M1_Wellpad*A*.gp11o*', ('WPA', 'M3'): 'M3_Wellpad*A*.gp11o*',
         ('WPD', 'M1'): 'M1_Wellpad*D*.gp11o*', ('WPD', 'M3'): 'M3_Wellpad*D*.gp11o*'}
lcname = lambda s: s.split(' ')[0]          # 'ASD2(max) M1 WPD 2C +Icap' -> 'ASD2(max)'

def load(srcdir):
    out = {}
    for key, pat in FILES.items():
        f = glob.glob(os.path.join(srcdir, pat))
        assert len(f) == 1, (key, f)
        out[key] = dict(file=os.path.basename(f[0]), cases=gp11o.parse(f[0]))
    return out

def coords(cases):
    """Pile-head (Y, Z) from rigid-cap kinematics: ux_i = ux0 + ry*z_i - rz*y_i (u = theta x r)."""
    G, rhs = [], {}
    for c in cases:
        ux, uy, uz, rx, ry, rz = c['disp']
        if abs(ry) + abs(rz) < 1e-12: continue
        G.append([ry, -rz])
        for p, d in c['piles'].items(): rhs.setdefault(p, []).append(d['gdisp'][0] - ux)
    xy = {}
    for p, v in rhs.items():
        (z, y), res, *_ = np.linalg.lstsq(np.array(G), np.array(v), rcond=None)
        xy[p] = (round(y * 4) / 4, round(z * 4) / 4)       # snap to 0.25 m grid
    return xy

def rows(data):
    """(wellpad, model, lc, pile, z, P, M, V, y_mm, None, None) like the old 'Data' sheet; M,V,y = resultants."""
    out = []
    for (wp, mdl), d in data.items():
        for c in d['cases']:
            for p, pl in sorted(c['piles'].items()):
                P = pl['gforce'][0]
                for r in pl['depth']:
                    out.append((wp, mdl, lcname(c['name']), p, r['x'], P, math.hypot(r['Mz'], r['My']),
                                math.hypot(r['Vy'], r['Vz']), math.hypot(r['dy'], r['dz']) * 1000, None, None))
    return out

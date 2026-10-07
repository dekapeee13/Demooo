"""Minimal parser for GROUP 2019 .gp11o output (3D): cap results, pile-top forces, depth tables."""
import re

def _nums(s):
    return [float(x) for x in re.findall(r'[-+]?\d*\.\d+(?:[Ee][-+]?\d+)?|[-+]?\d+(?:[Ee][-+]?\d+)?', s)]

def _after(lines, i, key):
    """numbers on the first line after the header line containing key (searched from line i)."""
    j = i
    while key not in lines[j]:
        j += 1
    return _nums(lines[j + 1])

def parse(path):
    lines = open(path, encoding='latin-1').read().splitlines()
    cases, cur, pile, i = [], None, None, 0
    while i < len(lines):
        t = lines[i]
        if t.strip().startswith('LOAD CASE :'):
            cur = dict(n=int(_nums(t)[0]), piles={}); cases.append(cur)
        elif t.strip().startswith('CASE NAME :') and cur is not None:
            cur['name'] = t.split(':', 1)[1].strip()
        elif 'EQUIVALENT CONCENTRATED LOAD AT ORIGIN' in t:
            cur['load'] = _nums(lines[i + 4]) + _nums(lines[i + 7])            # V, Hy, Hz, Mx, My, Mz
        elif 'DISPLACEMENT OF GROUPED PILE FOUNDATION AT ORIGIN' in t:
            cur['disp'] = _nums(lines[i + 4]) + _nums(lines[i + 7])            # ux, uy, uz, rx, ry, rz
        elif '* PILE GROUP *' in t:
            pile = dict(depth=[]); cur['piles'][int(_nums(t)[0])] = pile
        elif 'THE GLOBAL STRUCTURAL COORDINATE SYSTEM' in t and pile is not None:
            pile['gdisp'] = _after(lines, i, 'DISP. X'); pile['gforce'] = _after(lines, i, 'FOR. X')   # FX FY FZ MX MY MZ
        elif 'THE PILE COORDINATE SYSTEM (LOCAL AXES)' in t and pile is not None:
            pile['lforce'] = _after(lines, i, 'AXIAL')                                               # axial, Vy, Vz, Mx, My, Mz
        elif re.match(r'\s+\*{10}', t) and pile is not None:
            j = i + 1
            while j < len(lines) and len(_nums(lines[j])) >= 10:
                v = _nums(lines[j]); pile['depth'].append(dict(x=v[0], dy=v[1], dz=v[2], Mz=v[3], My=v[4], Vy=v[5], Vz=v[6]))
                j += 1
            i = j; continue
        i += 1
    return cases

if __name__ == '__main__':
    import sys
    c = parse(sys.argv[1])
    print(len(c), c[0]['name'], len(c[0]['piles']), c[0]['piles'][1]['gforce'], len(c[0]['piles'][1]['depth']), c[0]['piles'][1]['depth'][-1]['x'])

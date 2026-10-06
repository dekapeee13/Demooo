"""PLAXIS 2D Hardening Soil material commands, Wellpad A (BH-02) & Wellpad D (BH-04/BH-01).
Layering/parameters from Midas GTS NX HSs input (rev). pRef = 100, OCR = 1, POP = 0."""
# name, top, bot, type(S/C), gUnsat, gSat, E50ref, m, c, phi, psi, e0, usda, colour
WPA = [
    ('L01', 0, 7.5, 'S', 18.5, 18.4, 60000, 0.387, 1, 41.0, 11.0, 0.884, 'Sand', 0xC0E0F0),
    ('L02', 7.5, 12, 'S', 18.5, 18.4, 33994, 0.523, 1, 32.7, 2.7, 0.884, 'Sand', 0xA0C8E8),
    ('L03', 12, 15, 'C', 18.5, 18.4, 35275, 0.8, 10, 30.0, 0, 0.941, 'Clay', 0x5080B0),
    ('L04', 15, 28.5, 'S', 18.5, 18.4, 32908, 0.529, 1, 32.3, 2.3, 0.884, 'Sand', 0x90D0C0),
    ('L05', 28.5, 48, 'S', 18.5, 18.4, 38877, 0.498, 1, 34.3, 4.3, 0.884, 'Sand', 0x60B0A0),
    ('L06', 48, 54, 'S', 18.5, 18.4, 28927, 0.549, 1, 31.2, 1.2, 0.884, 'Sand', 0x80C0E0),
    ('L07', 54, 60, 'S', 18.5, 18.4, 35562, 0.515, 1, 33.2, 3.2, 0.884, 'Sand', 0x4090C0),
]
WPD = [
    ('L01', 0, 3, 'C', 19.6, 20.2, 17432, 0.8, 2, 24.0, 0, 0.605, 'Silty clay', 0x6080C0),
    ('L02', 3, 7.5, 'S', 19.6, 20.2, 51946, 0.429, 1, 39.5, 9.5, 0.558, 'Sand', 0xC0E0F0),
    ('L03', 7.5, 10.5, 'C', 19.6, 20.2, 13090, 0.8, 4, 26.0, 0, 0.605, 'Clay', 0x7090D0),
    ('L04', 10.5, 19.5, 'C', 19.6, 20.2, 15860, 0.8, 6, 28.0, 0, 0.605, 'Clay', 0x5070B0),
    ('L05', 19.5, 24, 'C', 19.6, 20.2, 27521, 0.8, 10, 30.0, 0, 0.605, 'Clay', 0x406090),
    ('L06', 24, 28.5, 'C', 17.8, 19.2, 13803, 0.8, 6, 28.0, 0, 0.776, 'Clay', 0x5878A8),
    ('L07', 28.5, 49.5, 'C', 17.8, 19.2, 24921, 0.8, 10, 30.0, 0, 0.776, 'Clay', 0x304878),
    ('L08', 49.5, 60, 'S', 17.8, 19.2, 29093, 0.548, 1, 31.2, 1.2, 0.724, 'Sand', 0x60B0A0),
]
# ASSUMED (no lab data): compacted embankment fill (gamma 18 as in bearing calc) and 0.5 m replacement fill
FILLS = [
    ('Embankment fill', 'S', 18.0, 20.0, 20000, 0.5, 5, 30.0, 0, 0.6, 'Sand', 0x8C8CB4),
    ('Replacement fill 0.5m', 'S', 19.0, 20.0, 30000, 0.5, 1, 25.0, 0, 0.6, 'Sand', 0x80A0A0),
]

def cmd(ident, typ, gU, gS, E50, m, c, phi, psi, e0, usda, col):
    clay = typ == 'C'
    return (f'_soilmat "Identification" "{ident}" "SoilModel" "Hardening Soil" "Colour" {col} '
            f'"DrainageType" "{"Undrained (A)" if clay else "Drained"}" '
            f'"gammaUnsat" {gU} "gammaSat" {gS} "E50Ref" {E50} "EOedRef" {E50} "EURRef" {3 * E50} '
            f'"PowerM" {m} "nu" 0.2 "pRef" 100 "Rf" 0.9 "cRef" {c} "phi" {phi} "psi" {psi} '
            f'"OCR" 1 "POP" 0 '
            f'"GroundwaterClassificationType" "USDA" "GroundwaterSoilClassUSDA" "{usda}" "GwUseDefaults" True '
            f'"nInit" {e0 / (1 + e0):.3f} "InterfaceStrengthDetermination" "Manual" "Rinter" {0.7 if clay else 0.67}')

def block(area, rows):
    out = []
    for n, t, b, typ, *p in rows:
        kind = 'Cohesive' if typ == 'C' else 'Granular'
        out.append(cmd(f'{area}-{n}-{kind}-{t}-{b}m', typ, *p))
    return out

lines = ['# Wellpad A (BH-02), GWL -7.0 m'] + block('WPA', WPA) + \
        ['', '# Wellpad D (BH-04, below 25 m BH-01), GWL -9.0 m'] + block('WPD', WPD) + \
        ['', '# Fill materials (assumed)'] + [cmd(*f) for f in FILLS]
open('PLAXIS_HS_materials.txt', 'w').write('\n'.join(lines) + '\n')

if __name__ == '__main__':
    assert len([l for l in lines if l.startswith('_soilmat')]) == len(WPA) + len(WPD) + len(FILLS)
    print('\n'.join(lines))

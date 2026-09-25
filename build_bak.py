# -*- coding: utf-8 -*-
"""Tollhúsið – Bakreikningur v1.0: tilboðsverð (hlutfall af fasteignamati) → nauðsynleg leiga → nauðsynlegt herbergisverð.
Allt formúlur. Sama keðja og „Rétt verð", snúin við."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
OUT = sys.argv[1]
F = 'Arial'; BLUE = Font(name=F, size=10, color='0000FF'); BLK = Font(name=F, size=10); BOLD = Font(name=F, size=10, bold=True)
H1 = Font(name=F, size=14, bold=True); H2 = Font(name=F, size=11, bold=True, color='1F3864'); SRC = Font(name=F, size=9, color='808080')
YEL = PatternFill('solid', fgColor='FFFF00'); GREY = PatternFill('solid', fgColor='D9E2F3'); LIGHT = PatternFill('solid', fgColor='EEF3F8')
NUM = '#,##0;(#,##0);-'; NUM1 = '#,##0.0;(#,##0.0);-'; PCT = '0.0%'; PCT2 = '0.00%'
wb = Workbook(); ws = wb.active; ws.title = 'Bakreikningur'
for i, w in enumerate([58, 16, 12, 80], 1): ws.column_dimensions[L(i)].width = w
ws['A1'] = 'Tollhúsið – Bakreikningur: hvað þarf herbergið að kosta á nótt fyrir gefið tilboðsverð?'; ws['A1'].font = H1
ws['A2'] = 'Gular reitir eru sleðar. m.kr án VSK á verðlagi 2026. Keðjan: verð → heildarkostnaður → verðmæti sem fjárfestar þurfa (× (1+D)) → NOI (× krafa) → leiga → tekjur → verð á nótt.'; ws['A2'].font = SRC
R = {}; r = 4
def sec(t):
    global r; ws.cell(r, 1, t).font = H2; r += 1
def inp(k, label, v, unit='', src='', fmt=NUM1, key=False):
    global r; ws.cell(r, 1, label); c = ws.cell(r, 2, v); c.font = BLUE; c.number_format = fmt
    if key: c.fill = YEL
    ws.cell(r, 3, unit).font = SRC; ws.cell(r, 4, src).font = SRC; R[k] = f'$B${r}'; r += 1
def fml(k, label, f, unit='', src='', fmt=NUM1, bold=False):
    global r; ws.cell(r, 1, label).font = BOLD if bold else BLK; c = ws.cell(r, 2, f); c.number_format = fmt; c.font = BOLD if bold else BLK
    if bold: c.fill = LIGHT
    ws.cell(r, 3, unit).font = SRC; ws.cell(r, 4, src).font = SRC; R[k] = f'$B${r}'; r += 1

sec('1. Sleðar')
inp('fmat', 'Fasteignamat 2026', 4725.5, 'm.kr', 'Fasteignaskrá F2000241 (2027: 4.914,9)')
inp('pct', 'Tilboðsverð sem hlutfall af fasteignamati', 0.42, '%', '0–100%. 42% ≈ 2.000 m.kr', PCT, True)
fml('price', 'Tilboðsverð', f'={R["fmat"]}*{R["pct"]}', 'm.kr', '', NUM, True)
inp('keys', 'Herbergi', 123, 'stk', 'Talning af grunnmyndum: 106 (4,5 m) / 123 (3,9 m)', '0', True)
inp('occ', 'Nýting á ársgrundvelli', 0.75, '%', 'Höfuðborgarsvæðið 74,7% 2025 (Hagstofa); Íslandshótel 68% 2023', PCT, True)
inp('fb', 'Veitingar, bar og aðrar tekjur, hlutfall af herbergistekjum', 0.40, '%', 'B-flokkur 35% + 5%; Íslandshótel: veitingatekjur ~30% af gistitekjum', PCT, True)
inp('C', 'C · leiga hótels sem hlutfall af heildartekjum', 0.25, '%', 'ODDSSON-samningur ÍF 2020: max(30% herb., 25% alls); USALI 20–24%', PCT, True)
inp('D', 'D · arðsemiskrafa fjárfesta, álag á heildarkostnað', 0.20, '%', '20% ≈ 12–15% IRR þróunaraðila; 10% = langtímaeigandi', PCT, True)
inp('y', 'Ávöxtunarkrafa kaupanda á NOI (endurfjármögnun/sala)', 0.0675, '%', 'Hótelleiga 6,75%; langtímaeigandi 6,0%; ríkisleiga 5,7%', PCT2, True)
inp('kola_m2', 'Kolaportssalur', 2409.6, 'm²', 'Fasteignaskrá 01-0105 + 01-0106', NUM1)
inp('kola_now', 'Núverandi leiga Kolaports, kr/m²/mán', 1000, 'kr', 'GT: ~1.000 kr/m²/mán; auglýsing RVK 2025 til rekstraraðila 1.570 kr/m²/mán (3,78 m.kr/mán). Staðfesta með samningi ríkis og borgar frá TORG.', NUM, True)
inp('kola_x', 'Margfaldari á Kolaportsleigu (×1 = óbreytt … ×4)', 1.0, 'x', 'Matarhöll/markaður á markaðsleigu ≈ ×3,5–5,5', '0.0', True)
fml('kola', 'Leiga jarðhæðar á ári', f'={R["kola_m2"]}*{R["kola_now"]}*{R["kola_x"]}*12/1000000', 'm.kr', '', NUM1)
inp('A0', 'A · hús fullbúið án lauss búnaðar (áætlun v1.0, 123 herb.)', 5483, 'm.kr', 'Kostnaðaráætlun fullbúið hótel v1.0 23.9.2026: framkvæmd 4.198 + ófyrirséð 12% + hönnun/stjórnun 781', NUM)
inp('ffe', 'Laus búnaður (FF&E) – 0 = rekstraraðili greiðir, 1 = eigandi', 0, '0/1', 'Áætlun v1.0: 739 m.kr. Hamranes: laus búnaður ekki hluti tilboðs', '0', True)
inp('ffe0', 'Laus búnaður alls (123 herb.)', 739, 'm.kr', 'Áætlun v1.0 kafli 9', NUM)
inp('cf', 'Kostnaðarstig (1,0 = áætlun v1.0)', 1.0, 'x', '', '0.00', True)
fml('A', 'A notað: (65% húsbundið + 35% herbergjabundið × herb/123) × kostnaðarstig + FF&E', f'=({R["A0"]}*0.65+{R["A0"]}*0.35*{R["keys"]}/123+{R["ffe"]}*{R["ffe0"]}*{R["keys"]}/123)*{R["cf"]}', 'm.kr', '', NUM, True)
inp('eur', 'Gengi EUR/ISK', 145, 'kr', 'Seðlabankinn sept 2026 ~145', NUM)
inp('vsk_g', 'VSK á gistingu', 0.11, '%', '', PCT)

sec('2. Fjármagns- og biðtímakostnaður (Hamranes-kjör)')
inp('ltc', 'Byggingarlán, hlutfall', 0.70, '%', 'GT: 70/30', PCT)
inp('r', 'Vextir byggingarláns, óverðtryggt, PIK', 0.102, '%', 'ÍSB kjörvextir + 0,15%', PCT2)
inp('t_a', 'Fjármögnunartími framkvæmdar, ár × meðalstaða', 1.0, 'ár', '2 ár × 50%', '0.0')
inp('t_p', 'Fjármögnunartími kaupverðs fram að endurfjármögnun', 4.5, 'ár', 'Kaup 2027 → stöðugur rekstur 2032', '0.0')
inp('bid', 'Biðtími fram að opnun', 3, 'ár', 'DSK, hönnun, framkvæmd', '0')
inp('fgj', 'Fasteignagjöld + vatn/fráveita í dag', 93.31, 'm.kr/ár', 'Söluyfirlit TORG: 85,48 + 7,83', NUM1)
inp('fgj_up', 'Hækkun fasteignagjalda eftir endurmat', 1.25, 'x', '', '0.00')
inp('ins', 'Tryggingar og annað á biðtíma', 15.5, 'm.kr/ár', 'Brunatrygging 5,5 + annað 10', NUM1)
inp('K', 'Stimpilgjald 1,6% af fasteignamati + kaupkostnaður', f'={R["fmat"]}*0.016+20', 'm.kr', '', NUM1)
fml('F_A', 'Fjármagnskostnaður á framkvæmd', f'={R["A"]}*{R["ltc"]}*{R["r"]}*{R["t_a"]}', 'm.kr')
fml('F_P', 'Fjármagnskostnaður á kaupverð', f'={R["price"]}*{R["ltc"]}*{R["r"]}*{R["t_p"]}', 'm.kr')
fml('H', 'Biðtímakostnaður að frádreginni Kolaportsleigu á biðtíma (óbreytt leiga, hálf á framkvæmdatíma)', f'=({R["fgj"]}+{R["ins"]})*{R["bid"]}-{R["kola_m2"]}*{R["kola_now"]}*12/1000000*({R["bid"]}-1)', 'm.kr')
fml('T', 'Heildarkostnaður verkefnis', f'={R["price"]}+{R["K"]}+{R["A"]}+{R["F_A"]}+{R["F_P"]}+{R["H"]}', 'm.kr', '', NUM, True)

sec('3. Það sem tekjurnar þurfa að bera')
fml('V', 'Verðmæti sem fjárfestar þurfa = heildarkostnaður × (1 + D)', f'={R["T"]}*(1+{R["D"]})', 'm.kr', '', NUM, True)
fml('noi', 'Nauðsynlegt NOI = verðmæti × krafa', f'={R["V"]}*{R["y"]}', 'm.kr/ár', '', NUM1, True)
fml('fixed', 'Fastur eigandakostnaður (fasteignagjöld eftir endurmat, tryggingar 8,25, viðhaldssjóður 0,5% af A)', f'={R["fgj"]}*{R["fgj_up"]}+8.25+{R["A"]}*0.005', 'm.kr/ár')
fml('leiga', 'Nauðsynlegar leigutekjur alls (NOI + fastur kostnaður, umsýsla 1%)', f'=({R["noi"]}+{R["fixed"]})/0.99', 'm.kr/ár', '', NUM1, True)
fml('hrent', 'Þar af leiga hótels (alls − jarðhæð)', f'={R["leiga"]}-{R["kola"]}', 'm.kr/ár', '', NUM1, True)
fml('rev', 'Nauðsynlegar heildartekjur hótels = leiga / C', f'={R["hrent"]}/{R["C"]}', 'm.kr/ár', '', NUM)
fml('rooms', 'Þar af herbergistekjur', f'={R["rev"]}/(1+{R["fb"]})', 'm.kr/ár', '', NUM)
fml('nights', 'Seldar nætur á ári', f'={R["keys"]}*365*{R["occ"]}', 'nætur', '', NUM)
fml('adr', 'NAUÐSYNLEGT MEÐALVERÐ Á SELDA NÓTT, án VSK', f'={R["rooms"]}*1000000/{R["nights"]}', 'kr', '', NUM, True)
ws.cell(r - 1, 2).fill = YEL
fml('adr_vsk', 'Sama með 11% VSK (verðið sem gesturinn sér)', f'={R["adr"]}*(1+{R["vsk_g"]})', 'kr', '', NUM, True)
fml('adr_eur', 'Í evrum með VSK', f'={R["adr_vsk"]}/{R["eur"]}', '€', '', NUM)
fml('revpar', 'RevPAR án VSK (verð × nýting)', f'={R["adr"]}*{R["occ"]}', 'kr', '', NUM)
fml('sumar', 'Sumarverð með VSK (meðaltal × 1,35)', f'={R["adr_vsk"]}*1.35', 'kr', 'Reykjavík: sumar 40–50% yfir vetri; JHB verðskrá 55/42/32', NUM)
fml('vetur', 'Vetrarverð með VSK (meðaltal × 0,78)', f'={R["adr_vsk"]}*0.78', 'kr', '', NUM)
fml('rev_key', 'Tekjur á herbergi á ári', f'={R["rev"]}/{R["keys"]}', 'm.kr', 'Íslandshótel 8,5 (2024, öll keðjan)', NUM1)
fml('op', 'Rekstraraðili heldur eftir leigu (USALI: EBITDAR ≈ 37% af tekjum við þessa tekjusamsetningu, − FF&E 4% − leiga)', f'=0.37-0.04-{R["C"]}', '% tekna', 'Neikvætt eða < 8% = enginn rekstraraðili skrifar undir', PCT)

sec('4. Viðmið')
for name, v, src in [('Íslandshótel, öll keðjan 2023, án VSK', 25400, 'Útgefandalýsing maí 2024'), ('B-flokkur, upper upscale, án VSK', 48000, 'Konsulat/Parliament-flokkur, €330'), ('C-flokkur, lúxus, án VSK', 65000, 'EDITION-flokkur, €450'), ('Black Dunes Þorlákshöfn áætlun 2027–32, án VSK', 30000, 'Flóra rekstraráætlun 17.3.2025'), ('Næsland-áætlun ár 3, m/VSK ≈ 40 þ.kr', 36000, 'Rekstrar- og söluáætlun draft 1')]:
    ws.cell(r, 1, name); c = ws.cell(r, 2, v); c.font = BLUE; c.number_format = NUM; ws.cell(r, 3, 'kr').font = SRC; ws.cell(r, 4, src).font = SRC; r += 1
r += 1

sec('5. Tafla: nauðsynlegt verð á nótt án VSK (kr) eftir tilboðsverði og nýtingu – annað eins og í sleðum')
ws.cell(r, 1, 'Tilboðsverð % af fmat ↓ · nýting →').font = BOLD; ws.cell(r, 1).fill = GREY
occs = [0.65, 0.70, 0.75, 0.80]
for j, o in enumerate(occs): c = ws.cell(r, 2 + j, o); c.number_format = PCT; c.font = BOLD; c.fill = GREY
hr = r; r += 1
def adr_formula(pct_cell, occ_cell, kola_x_cell):
    price = f'({R["fmat"]}*{pct_cell})'
    T = f'({price}+{R["K"]}+{R["A"]}+{R["F_A"]}+{price}*{R["ltc"]}*{R["r"]}*{R["t_p"]}+{R["H"]})'
    leiga = f'(({T}*(1+{R["D"]})*{R["y"]}+{R["fixed"]})/0.99)'
    kola = f'({R["kola_m2"]}*{R["kola_now"]}*{kola_x_cell}*12/1000000)'
    return f'=({leiga}-{kola})/{R["C"]}/(1+{R["fb"]})*1000000/({R["keys"]}*365*{occ_cell})'
for pc in [0.1, 0.2, 0.3, 0.42, 0.5, 0.6, 0.75, 0.9, 1.0]:
    c = ws.cell(r, 1, pc); c.number_format = PCT; c.font = BOLD
    for j in range(len(occs)):
        c = ws.cell(r, 2 + j, adr_formula(f'$A{r}', f'{L(2+j)}${hr}', R['kola_x'])); c.number_format = NUM
    r += 1
r += 1
sec('6. Tafla: nauðsynlegt verð á nótt án VSK (kr) eftir tilboðsverði og Kolaportsleigu (nýting úr sleða)')
ws.cell(r, 1, 'Tilboðsverð % af fmat ↓ · margfaldari Kolaports →').font = BOLD; ws.cell(r, 1).fill = GREY
xs = [1, 2, 3, 4]
for j, x in enumerate(xs): c = ws.cell(r, 2 + j, x); c.number_format = '"×"0'; c.font = BOLD; c.fill = GREY
hr2 = r; r += 1
for pc in [0.1, 0.2, 0.3, 0.42, 0.5, 0.6, 0.75, 0.9, 1.0]:
    c = ws.cell(r, 1, pc); c.number_format = PCT; c.font = BOLD
    for j in range(len(xs)):
        c = ws.cell(r, 2 + j, adr_formula(f'$A{r}', R['occ'], f'{L(2+j)}${hr2}')); c.number_format = NUM
    r += 1
r += 1
ws.cell(r, 1, 'Lestur: hver reitur er ársmeðaltal á selda nótt án VSK. Margfaldið með 1,11 fyrir verðið sem gesturinn sér, deilið með 145 fyrir evrur. Sumarverð er um 35% yfir meðaltali og vetrarverð um 22% undir.').font = SRC
ws.freeze_panes = 'A4'
wb.save(OUT); print('saved', OUT)

# ---- python cross-check
def bak(pct, keys=123, occ=0.75, fb=0.40, C=0.25, D=0.20, y=0.0675, kola_now=1000, kx=1.0, A0=5483, ffe=0, cf=1.0):
    price = 4725.5 * pct; A = (A0 * 0.65 + A0 * 0.35 * keys / 123 + ffe * 739 * keys / 123) * cf
    K = 4725.5 * 0.016 + 20; FA = A * 0.7 * 0.102; FP = price * 0.7 * 0.102 * 4.5; H = (93.31 + 15.5) * 3 - 2409.6 * kola_now * 12 / 1e6 * 2
    T = price + K + A + FA + FP + H; V = T * (1 + D); noi = V * y; fixed = 93.31 * 1.25 + 8.25 + A * 0.005
    leiga = (noi + fixed) / 0.99; kola = 2409.6 * kola_now * kx * 12 / 1e6; hrent = leiga - kola; rev = hrent / C; rooms = rev / (1 + fb)
    return rooms * 1e6 / (keys * 365 * occ), leiga, hrent
if __name__ == '__main__':
    for args in [dict(pct=0.42), dict(pct=0.42, D=0.10, y=0.06), dict(pct=0.42, D=0.10, y=0.06, kx=3), dict(pct=1.0), dict(pct=0.2, D=0.10, y=0.06, kx=3)]:
        a, l, h = bak(**args); print(args, f'ADR án VSK {a:,.0f}  leiga {l:,.0f}  hótel {h:,.0f}')

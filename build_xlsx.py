# Byggir "Tollhúsið – Kostnaðar-, rekstrar- og verðgetulíkan v0.1.xlsx" úr model.py (sömu forsendur, allt formúlur)
import sys, json
sys.path.insert(0, '/Users/villithor/repos/tollhusid')
import model as M
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
wb = Workbook()
F = 'Arial'
BLUE = Font(name=F, size=10, color='0000FF'); BLK = Font(name=F, size=10); GRN = Font(name=F, size=10, color='008000')
BOLD = Font(name=F, size=10, bold=True); H1 = Font(name=F, size=14, bold=True); H2 = Font(name=F, size=11, bold=True, color='1F3864')
YEL = PatternFill('solid', fgColor='FFFF00'); GREY = PatternFill('solid', fgColor='D9E2F3'); LIGHT = PatternFill('solid', fgColor='F2F2F2')
NUM = '#,##0;(#,##0);-'; NUM1 = '#,##0.0;(#,##0.0);-'; PCT = '0.0%'; PCT2 = '0.00%'
thin = Side(style='thin', color='BFBFBF')

def setw(ws, widths):
    for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w

# ============================================================== FORSENDUR
ws = wb.active; ws.title = 'Forsendur'
setw(ws, [46, 14, 10, 14, 14, 14, 70])
ws['A1'] = 'Tollhúsið – Næsland: kostnaðar-, rekstrar- og verðgetulíkan v0.1'; ws['A1'].font = H1
ws['A2'] = 'Allar fjárhæðir í m.kr án VSK á verðlagi 2026 nema annað sé tekið fram. Bláar tölur = forsendur sem má breyta. Gular = lykilstýringar. Svartar = formúlur. Vinnuskjal GT/ÍF 22.9.2026, trúnaðarmál.'; ws['A2'].font = Font(name=F, size=9, italic=True)
ws['A3'] = 'Stýringar: veldu gæðaflokk (A/B/C), viðbótarbyggingarmagn (none/top/full), kostnaðarstig, leiguhlutfall og kaupverð – allt annað reiknast.'; ws['A3'].font = Font(name=F, size=9, italic=True)

ref = {}
row = [5]
def sec(title):
    r = row[0]; ws.cell(r, 1, title).font = H2; row[0] += 1
def inp(key, label, val, unit='', src='', fmt=NUM1, key_cell=False):
    r = row[0]; ws.cell(r, 1, label).font = BLK
    c = ws.cell(r, 2, val); c.font = BLUE; c.number_format = fmt
    if key_cell: c.fill = YEL
    ws.cell(r, 3, unit).font = Font(name=F, size=9, color='808080'); ws.cell(r, 7, src).font = Font(name=F, size=9, color='808080')
    ref[key] = f"Forsendur!$B${r}"; row[0] += 1
def fml(key, label, formula, unit='', src='', fmt=NUM1):
    r = row[0]; ws.cell(r, 1, label).font = BLK
    c = ws.cell(r, 2, formula); c.font = BLK; c.number_format = fmt
    ws.cell(r, 3, unit).font = Font(name=F, size=9, color='808080'); ws.cell(r, 7, src).font = Font(name=F, size=9, color='808080')
    ref[key] = f"Forsendur!$B${r}"; row[0] += 1

B = M.BUILDING; P = M.BASE
sec('1. Stýringar')
inp('tier', 'Gæðaflokkur (A / B / C)', P['tier'], '', 'A = Næsland eins og kynnt; B = upper upscale lifestyle; C = lúxus/boutique', '@', True)
inp('extra', 'Viðbótarbyggingarmagn (none / top / full)', P['extra_floor'], '', 'top = 6. hæð ofan á 5. hæð (1.560 m²); full = auk hæðar yfir plötu 3. hæðar / Kolaportsþaki (alls 4.460 m²). Háð breytingu á deiliskipulagi.', '@', True)
inp('cost_factor', 'Kostnaðarstig (1,0 = grunnur, 0,8 = bjartsýnt, 1,2 = varfærið)', P['cost_factor'], 'x', 'Margfaldari á heildarframkvæmdakostnað', '0.00', True)
inp('rent_override', 'Leiguhlutfall af heildartekjum (autt = reiknað úr EBITDA)', '', '%', 'JHB-áætlun jafngildir ~0,30 af tekjum með F&B; markaður 0,22–0,30', PCT, True)
inp('price', 'Kaupverð (prófað í sjóðstreymi)', 1000.0, 'm.kr', 'Breyta og lesa IRR í Sjóðstreymi!B34', NUM, True)

sec('2. Eignin (fasteignaskrá, söluyfirlit TORG 31.8.2026)')
inp('A_total', 'Heildarstærð brúttó', B['A_total'], 'm²', 'Fasteignaskrá F2000241')
inp('A_kola', 'Kolaportið (01-0105 + 01-0106)', B['A_kolaport'], 'm²', 'Rekstrareiningar í fasteignayfirliti: 1.417,1 + 992,5')
fml('A_hotel_base', 'Hótelhluti í núverandi húsi', f"={ref['A_total']}-{ref['A_kola']}", 'm²')
inp('fmat', 'Fasteignamat 2026', B['fmat'], 'm.kr', 'Fasteignaskrá; 2027: 4.914,9')
inp('stimpil', 'Stimpilgjald lögaðila', B['stimpil'], '%', 'Söluyfirlit', PCT)
inp('fgj', 'Álögð fasteignagjöld 2026', B['fasteignagjold'], 'm.kr/ár', 'Söluyfirlit TORG')
inp('vatn', 'Vatns- og fráveitugjöld', B['vatn_fraveita'], 'm.kr/ár', 'Söluyfirlit TORG')
inp('brunatr', 'Brunatrygging', B['brunatr'], 'm.kr/ár', 'Söluyfirlit TORG')
inp('kaupkostn', 'Annar kaupkostnaður (þinglýsing, ráðgjöf, DD)', 20.0, 'm.kr', 'Mat')

sec('3. Gæðaflokkar – rekstrar- og innréttingaforsendur')
hdr = row[0]
for j, t in enumerate(['A', 'B', 'C']):
    ws.cell(hdr, 2+j, t).font = BOLD; ws.cell(hdr, 2+j).fill = GREY
ws.cell(hdr, 1, 'Forsenda').font = BOLD; ws.cell(hdr, 1).fill = GREY; ws.cell(hdr, 7, 'Heimild / rök').font = BOLD; ws.cell(hdr, 7).fill = GREY
row[0] += 1
tier_rows = {}
tier_items = [
    ('keys', 'Fjöldi herbergja í núverandi húsi', NUM, 'A: JHB 100 herb. á 2.–5. hæð; B/C: stærri herbergi'),
    ('adr', 'Meðalverð á selda nótt (ADR) án 11% VSK, kr', NUM, 'A ≈ JHB verðskrá (€245); B ≈ €330 (Konsulat/Parliament-flokkur); C ≈ €450 (EDITION-flokkur). Staðfesta með rekstraraðila.'),
    ('occ', 'Nýting við stöðugan rekstur', PCT, 'Höfuðborgarsvæðið 74,7% 2025 (Hagstofa); lúxus lægri'),
    ('fb', 'Veitingar og bar, hlutfall af herbergistekjum', PCT, 'Lifestyle 20–35%, lúxus m/þakgarði 45%'),
    ('other', 'Aðrar tekjur (spa, meðlimir, verslun), hlutfall', PCT, ''),
    ('ebitda', 'EBITDA fyrir leigu og FF&E-sjóð, hlutfall af heildartekjum', PCT, 'Íslensk hótel 28–36%; JHB 33–44% (án F&B)'),
    ('fit_hard', 'Innréttingar herbergja, hard fit-out, m.kr/herb', NUM1, '≈ 52–66 m² × 170–330 þ.kr/m²'),
    ('ffe', 'Laus búnaður herbergja (FF&E), m.kr/herb', NUM1, 'JHB 2,4–3,2 m.kr'),
    ('public', 'Almenn rými: lobby, veitingar, spa/gym, fundir, þakgarðsbygging, m.kr', NUM, 'JHB: eldhús 17, bar 4, gym 3,5, þakgarður 220 (m/VSK)'),
    ('key_m2', 'Brúttó m² á herbergi í viðbótarhæð (nettó, ×1,2 fyrir ganga)', NUM, ''),
]
for key, label, fmt, src in tier_items:
    r = row[0]; ws.cell(r, 1, label)
    for j, t in enumerate(['A', 'B', 'C']):
        c = ws.cell(r, 2+j, M.TIERS[t][key]); c.font = BLUE; c.number_format = fmt
    ws.cell(r, 7, src).font = Font(name=F, size=9, color='808080')
    tier_rows[key] = r; row[0] += 1
# chosen values
sec('4. Valinn flokkur (reiknað)')
fml('tier_col', 'Dálkur valins flokks', f'=MATCH({ref["tier"]},$B${hdr}:$D${hdr},0)', '', '', '0')
for key, label, fmt, src in tier_items:
    r = tier_rows[key]
    fml('t_'+key, label, f'=INDEX($B${r}:$D${r},1,{ref["tier_col"]})', '', '', fmt)

sec('5. Viðbótarbyggingarmagn')
r0 = row[0]
for j, (k, v) in enumerate([('none', 0.0), ('top', 1560.0), ('full', 4460.0)]):
    ws.cell(r0+j, 1, f'  {k}').font = BLK; c = ws.cell(r0+j, 2, v); c.font = BLUE; c.number_format = NUM
ws.cell(r0, 7, 'top: grunnflötur 4./5. hæðar; full: + þakflötur 3. hæðar (bílastæði + hraðbraut) ~2.900 m². COWI 6.9.2024: ein hæð í lagi burðarþolslega.').font = Font(name=F, size=9, color='808080')
row[0] += 3
ws.cell(r0-1, 1).font = H2
fml('extra_m2', 'Viðbótarbyggingarmagn valið', f'=IF({ref["extra"]}="none",$B${r0},IF({ref["extra"]}="top",$B${r0+1},$B${r0+2}))', 'm²', '', NUM)
inp('extra_rate', 'Nýbygging ofan á hús: burður, klæðning, lagnir, án innréttinga herbergja', P['extra_floor_rate'], 'þ.kr/m²', 'Hamranes EAC 780 þ/m² fullbúið; hér skel + kerfi')
fml('extra_keys', 'Viðbótarherbergi', f'=ROUND({ref["extra_m2"]}/({ref["t_key_m2"]}*1.2),0)', 'herb', '', NUM)
fml('keys', 'Herbergi alls', f'={ref["t_keys"]}+{ref["extra_keys"]}', 'herb', '', NUM)
fml('A_hotel', 'Hótelhluti alls', f'={ref["A_hotel_base"]}+{ref["extra_m2"]}', 'm²', '', NUM)

sec('6. Framkvæmd – einingarverð byggingarhluta (þ.kr/m² af hótelhluta eða m.kr heild)')
inp('demo_rate', 'Niðurrif innanhúss, hreinsun, förgun', P['demo_rate'], 'þ.kr/m²', 'Mat ÍF; COWI 2025')
inp('hazmat', 'Asbest, PCB, þungmálmar, mygla – afmengun', P['hazmat'], 'm.kr', 'COWI 2024 viðauki B (550–600 m² asbest), COWI 2025 sýnataka')
inp('sewer', 'Frárennslislagnir í grunni', P['sewer'], 'm.kr', 'COWI 2025 forgangsatriði 1')
inp('windows_m2', 'Gluggaveggjaeiningar, magn', P['windows_m2'], 'm²', 'COWI 2024 magnskrá')
inp('windows_rate', 'Gluggaveggjaeiningar, einingarverð', P['windows_rate'], 'þ.kr/m²', 'Mat ÍF (álkerfi, gler); JHB 10 þ/m² er ekki nothæft')
inp('single_windows', 'Stakir gluggar og hurðir', P['single_windows'], 'm.kr', 'COWI 2024')
inp('facade', 'Ytra byrði: klæðning, múr, handrið, mósaík/gabbró', P['facade'], 'm.kr', 'COWI 2024: klæðning 190 (án glugga) + listaverk')
inp('roof_garden', 'Þak og þakgarður 3. hæðar (án sundlaugar/padel)', P['roof_garden'], 'm.kr', 'COWI minnisblað 6.9.2024; JHB þakgarður 220 m/VSK')
inp('structure', 'Burðarvirki: op, lyftuop, stigar, styrkingar', P['structure'], 'm.kr', 'Mat ÍF')
inp('lifts', 'Lyftur, 3 stk', P['lifts'], 'm.kr', 'JHB 2 × 24 m/VSK')
inp('plumb_rate', 'Pípulagnir innanhúss', P['plumb_rate'], 'þ.kr/m²', 'Hamranes kafli 3: 24,8 þ/m² nýbygging')
inp('hvac_rate', 'Loftræsting', P['hvac_rate'], 'þ.kr/m²', 'COWI 2025; JHB 4 m.kr er ekki nothæft')
inp('elec_rate', 'Raflagnir, lýsing, öryggis- og hússtjórnarkerfi', P['elec_rate'], 'þ.kr/m²', 'Hamranes kafli 4: 49,5 þ/m²')
inp('sprinkler_rate', 'Vatnsúðakerfi', P['sprinkler_rate'], 'þ.kr/m²', 'JHB 12,8 þ/m²')
inp('kolaport_fix', 'Kolaportið: leki, loft, salerni, hringhurð, lýsing', P['kolaport_fix'], 'm.kr', 'COWI 2025 kafli 2.7')
inp('unc_building', 'Ófyrirséð, byggingarhluti', P['unc_building'], '%', 'COWI: skrár ekki tæmandi, úrtaksskoðun', PCT)
inp('unc_fitout', 'Ófyrirséð, innréttingar', P['unc_fitout'], '%', '', PCT)
inp('design', 'Hönnun, ráðgjöf, eftirlit (af hard cost)', P['design'], '%', 'Hamranes ~5% nýbygging; umbreyting hærra', PCT)
inp('permits', 'Byggingarleyfi, úttektir, gjöld', P['permits'], '%', '', PCT)
inp('dev_fee', 'Þróunar- og verkefnastjórnunarþóknun ÍF', P['dev_fee'], '%', 'Þ113-líkan 6%', PCT)

sec('7. Rekstur eiganda og leiga')
inp('kola_rent', 'Leiga Kolaportsins án VSK', P['kolaport_rent'], 'm.kr/ár', 'RVK auglýsing júní 2025: lágmark 3,777 m.kr/mán; samningur ríkis og borgar óséður')
inp('kola_works', 'Hlutfall Kolaportsleigu á framkvæmdatíma', P['kolaport_during_works'], '%', 'Truflun vegna frárennslis/gafla/þaks', PCT)
inp('ffe_reserve', 'FF&E-sjóður rekstraraðila, hlutfall af tekjum', P['ffe_reserve'], '%', 'Alþjóðlegt viðmið 3–5%', PCT)
inp('op_min', 'Lágmarksframlegð rekstraraðila eftir leigu, hlutfall af tekjum', P['operator_min'], '%', 'JHB miðaði við 10% af herbergistekjum', PCT)
inp('fgj_uplift', 'Hækkun fasteignagjalda eftir endurmat', 1.25, 'x', 'Fasteignamat hækkar með endurbótum', '0.00')
inp('ins', 'Tryggingar eftir endurbætur', P['ins']*1.5, 'm.kr/ár', 'Brunatrygging 5,5 í dag')
inp('mgmt', 'Umsýsla eiganda, hlutfall af leigu', P['mgmt'], '%', 'Hamranes-líkan 1%', PCT)
inp('maint', 'Viðhaldssjóður eiganda, hlutfall af framkvæmdakostnaði', P['maint'], '%', 'Hamranes-líkan 0,1–0,3% af brunabótamati', PCT2)
inp('hold_other', 'Annar kostnaður á biðtíma (öryggi, hiti, umsjón)', 10.0, 'm.kr/ár', 'Mat')

sec('8. Tímalína og fjármögnun')
inp('buy_year', 'Kaupár (afhending eftir fjárlög 2027)', P['buy_year'], 'ár', 'Söluheimild í fjárlagafrv. 2027', '0')
inp('works1', 'Framkvæmdaár 1 (eftir DSK-breytingu og hönnun)', P['works_years'][0], 'ár', '', '0')
inp('works2', 'Framkvæmdaár 2', P['works_years'][1], 'ár', '', '0')
inp('split1', 'Hlutfall framkvæmdar á ári 1', 0.40, '%', '', PCT)
inp('open_year', 'Opnun', P['open_year'], 'ár', '', '0')
inp('ramp1', 'Leiga ár 1 eftir opnun, hlutfall af stöðugri', P['ramp'][0], '%', 'JHB: 20% undir markaði ár 1', PCT)
inp('ramp2', 'Leiga ár 2', P['ramp'][1], '%', '', PCT)
inp('infl', 'Verðbólga / vísitala', P['infl'], '%', 'Hamranes-líkan 4%; hér 3,5%', PCT)
inp('ltc', 'Byggingarlán, hlutfall af kaupverði + framkvæmd', P['ltc'], '%', 'GT: 70/30', PCT)
inp('cl_rate', 'Vextir byggingarláns, óverðtryggt, PIK', P['cl_rate'], '%', 'Hamranes ÍSB: kjörvextir + 0,15% ≈ 10,2–10,55%', PCT2)
inp('cl_fee', 'Lántökugjald', P['cl_fee'], '%', 'ÍSB 0,20%', PCT2)
inp('yld', 'Ávöxtunarkrafa á NOI við endurfjármögnun og sölu', P['yld'], '%', 'Reitir nýkaup 7,8%, matskrafa 6,4–6,7%, ríkisleiga 5,7%; hótelleiga hér 6,75%', PCT2)
inp('ltv_refi', 'LTV verðtryggðs skuldabréfs við endurfjármögnun', P['ltv_refi'], '%', 'Hamranes 67,7%', PCT)
inp('bond_real', 'Raunvextir bréfs (RIKS + álag)', P['bond_real'], '%', 'Hamranes-líkan 2,86% + 1,0%', PCT2)
inp('bond_years', 'Lánstími bréfs', P['bond_years'], 'ár', '', '0')
inp('bond_cost', 'Útgáfukostnaður', P['bond_cost'], '%', '', PCT2)
inp('sale_cost', 'Sölukostnaður við útgöngu eftir 10 ár', P['sale_cost'], '%', '', PCT)

# ============================================================== FRAMKVÆMD
wf = wb.create_sheet('Framkvæmd'); setw(wf, [70, 12, 8, 12, 12, 80])
wf['A1'] = 'Framkvæmdakostnaður – m.kr án VSK, verðlag 2026'; wf['A1'].font = H1
hdrs = ['Verkþáttur', 'Magn', 'Ein.', 'Ein.verð þ.kr', 'Upphæð m.kr', 'Heimild / rök']
for j, h in enumerate(hdrs, 1): c = wf.cell(3, j, h); c.font = BOLD; c.fill = GREY
r = 4; fr = {}
def line(key, name, qty_f, unit, rate_f, src):
    global r
    wf.cell(r, 1, name); wf.cell(r, 2, qty_f).number_format = NUM; wf.cell(r, 3, unit); wf.cell(r, 4, rate_f).number_format = NUM
    wf.cell(r, 5, f'=IF(C{r}="heild",D{r},B{r}*D{r}/1000)').number_format = NUM1
    wf.cell(r, 6, src).font = Font(name=F, size=9, color='808080')
    for c in (2, 4): wf.cell(r, c).font = GRN
    fr[key] = r; r += 1
wf.cell(r, 1, 'Hús – byggingarhluti').font = H2; r += 1; b0 = r
A = f"={ref['A_hotel_base']}"
line('demo', 'Niðurrif innanhúss, hreinsun og förgun (hótelhluti)', A, 'm²', f"={ref['demo_rate']}", 'COWI 2025: gólfefni, klæðningar, milliveggir að mestu upprunaleg')
line('hazmat', 'Asbest, PCB, þungmálmar og mygla – afmengun með leyfishöfum', 1, 'heild', f"={ref['hazmat']}", 'COWI 2024 viðauki B; COWI 2025 sýnataka (8 af 33 sýnum menguð)')
line('sewer', 'Frárennslislagnir í grunni undir gólfplötu', 1, 'heild', f"={ref['sewer']}", 'COWI 2025 forgangsatriði 1')
line('win', 'Gluggaveggjaeiningar, ný kerfi', f"={ref['windows_m2']}", 'm²', f"={ref['windows_rate']}", 'COWI 2024 magnskrá 1.600 m²')
line('swin', 'Stakir gluggar og hurðir í steyptum veggjum', 1, 'heild', f"={ref['single_windows']}", 'COWI 2024')
line('facade', 'Ytra byrði: klæðning steyptra flata, múr, handrið, mósaík/gabbró', 1, 'heild', f"={ref['facade']}", 'COWI 2024: 84 m.kr múr+málun / 190 m.kr klæðning')
line('roof', 'Þak og þakgarður á 3. hæð', 1, 'heild', f'=IF({ref["extra"]}="full",60,{ref["roof_garden"]})', 'Fellur að mestu niður ef byggt er yfir plötu 3. hæðar')
line('struct', 'Burðarvirki: op, lyftuop, nýir stigar, styrkingar', 1, 'heild', f"={ref['structure']}", 'Mat ÍF')
line('lifts', 'Lyftur, 3 stk', 1, 'heild', f"={ref['lifts']}", 'Mat ÍF')
line('plumb', 'Pípulagnir innanhúss', A, 'm²', f"={ref['plumb_rate']}", 'Hamranes kafli 3 sem viðmið')
line('hvac', 'Loftræsting, nýtt kerfi', A, 'm²', f"={ref['hvac_rate']}", 'COWI 2025')
line('elec', 'Raflagnir, lýsing, öryggis- og hússtjórnarkerfi', A, 'm²', f"={ref['elec_rate']}", 'Hamranes kafli 4')
line('spr', 'Vatnsúðakerfi', A, 'm²', f"={ref['sprinkler_rate']}", 'JHB')
line('kola', 'Kolaportið: leki, loftaklæðning, salerni, hringhurð, lýsing', 1, 'heild', f"={ref['kolaport_fix']}", 'COWI 2025 kafli 2.7')
line('extra', 'Viðbótarbyggingarmagn ofan á hús (skel, klæðning, kerfi)', f"={ref['extra_m2']}", 'm²', f"={ref['extra_rate']}", 'Háð DSK; COWI: ein hæð í lagi burðarþolslega')
b1 = r-1
wf.cell(r, 1, 'Samtals byggingarhluti').font = BOLD; wf.cell(r, 5, f'=SUM(E{b0}:E{b1})').number_format = NUM1; wf.cell(r, 5).font = BOLD; fr['building'] = r; r += 2
wf.cell(r, 1, 'Herbergi og almenn rými').font = H2; r += 1; f0 = r
line('fit', 'Innréttingar herbergja, hard fit-out', f"={ref['keys']}", 'herb', f"={ref['t_fit_hard']}*1000", 'Flokkur valinn á Forsendur')
line('ffe', 'Laus búnaður herbergja (FF&E)', f"={ref['keys']}", 'herb', f"={ref['t_ffe']}*1000", '')
line('public', 'Almenn rými: lobby, veitingar, bar, eldhús, spa/gym, fundir, þakgarðsbygging', 1, 'heild', f"={ref['t_public']}", '')
f1 = r-1
wf.cell(r, 1, 'Samtals innréttingar').font = BOLD; wf.cell(r, 5, f'=SUM(E{f0}:E{f1})').number_format = NUM1; wf.cell(r, 5).font = BOLD; fr['fitout'] = r; r += 2
def tot(key, label, formula, bold=True):
    global r
    wf.cell(r, 1, label).font = BOLD if bold else BLK; c = wf.cell(r, 5, formula); c.number_format = NUM1; c.font = BOLD if bold else BLK; fr[key] = r; r += 1
tot('unc', 'Ófyrirséð', f"=E{fr['building']}*{ref['unc_building']}+E{fr['fitout']}*{ref['unc_fitout']}", False)
tot('hard', 'Hard cost alls', f"=E{fr['building']}+E{fr['fitout']}+E{fr['unc']}")
tot('design', 'Hönnun, ráðgjöf, eftirlit', f"=E{fr['hard']}*{ref['design']}", False)
tot('permits', 'Leyfi og gjöld', f"=E{fr['hard']}*{ref['permits']}", False)
tot('fee', 'Þróunarþóknun ÍF', f"=E{fr['hard']}*{ref['dev_fee']}", False)
tot('total', 'FRAMKVÆMD ALLS án VSK og án fjármagnskostnaðar (× kostnaðarstig)', f"=(E{fr['hard']}+E{fr['design']}+E{fr['permits']}+E{fr['fee']})*{ref['cost_factor']}")
wf.cell(fr['total'], 5).fill = YEL
tot('per_key', 'á herbergi, m.kr', f"=E{fr['total']}/{ref['keys']}", False)
tot('per_m2', 'á m² hótelhluta, þ.kr', f"=E{fr['total']}/{ref['A_hotel']}*1000", False)
r += 1
wf.cell(r, 1, 'Til samanburðar: JHB vinnuskjal 1.697 m.kr m/VSK = 1.368 án VSK (17 m.kr/herb) – vantar frárennsli, glugga, asbest, ytra byrði, loftræstingu, Kolaport, ófyrirséð. KEF-tilboð ÍF 2026: 30–36 m.kr/herb nýbygging án lóðar. Hyatt Centric L176 (Reitir): ~55–60 m.kr/herb með lóð. LHÍ-umbreyting ríkisins: 12,8 → 17 ma.kr fyrir 15.500 m².').font = Font(name=F, size=9, italic=True)
CAP = f"Framkvæmd!$E${fr['total']}"

# ============================================================== REKSTUR
wr = wb.create_sheet('Rekstur'); setw(wr, [60, 14, 60])
wr['A1'] = 'Rekstur hótelsins við stöðugan rekstur (ár 3 eftir opnun) og leiguþol – m.kr á verðlagi 2026'; wr['A1'].font = H1
rr = 3; rref = {}
def rl(key, label, formula, fmt=NUM1, note='', bold=False):
    global rr
    wr.cell(rr, 1, label).font = BOLD if bold else BLK; c = wr.cell(rr, 2, formula); c.number_format = fmt; c.font = BOLD if bold else BLK
    wr.cell(rr, 3, note).font = Font(name=F, size=9, color='808080'); rref[key] = f"Rekstur!$B${rr}"; rr += 1
rl('keys', 'Herbergi', f"={ref['keys']}", NUM)
rl('adr', 'ADR án VSK, kr', f"={ref['t_adr']}", NUM)
rl('occ', 'Nýting', f"={ref['t_occ']}", PCT)
rl('revpar', 'RevPAR án VSK, kr', f"=B4*B5", NUM)
rl('rooms', 'Herbergistekjur', f"=B3*B4*B5*365/1000000", NUM1, 'herbergi × ADR × nýting × 365')
rl('fb', 'Veitingar og bar', f"=B7*{ref['t_fb']}", NUM1)
rl('other', 'Aðrar tekjur', f"=B7*{ref['t_other']}", NUM1)
rl('rev', 'Heildartekjur', f"=B7+B8+B9", NUM1, '', True)
rl('rev_key', 'Tekjur á herbergi', f"=B10/B3", NUM1)
rl('ebitda', 'EBITDA fyrir leigu og FF&E-sjóð', f"=B10*{ref['t_ebitda']}", NUM1)
rl('reserve', 'FF&E-sjóður', f"=B10*{ref['ffe_reserve']}", NUM1)
rl('opkeep', 'Lágmarksframlegð rekstraraðila eftir leigu', f"=B10*{ref['op_min']}", NUM1)
rl('rent_calc', 'Leiguþol reiknað (EBITDA – FF&E – framlegð rekstraraðila)', f"=B12-B13-B14", NUM1)
rl('rent', 'Leiga hótels notuð í líkani', f'=IF({ref["rent_override"]}="",B15,B10*{ref["rent_override"]})', NUM1, 'Yfirskrifað með leiguhlutfalli á Forsendur ef sett', True)
rl('rent_share', 'Leiga sem hlutfall af heildartekjum', f"=B16/B10", PCT)
rl('rent_rooms_share', 'Leiga sem hlutfall af herbergistekjum (JHB: 34%)', f"=B16/B7", PCT)
rl('kola', 'Leiga Kolaportsins', f"={ref['kola_rent']}", NUM1)
rl('stab_rent', 'Leigutekjur eiganda alls', f"=B16+B19", NUM1, '', True)
rr += 1
rl('fixed', 'Fastur eigandakostnaður (fasteignagjöld × uplift, tryggingar, viðhaldssjóður)', f"=({ref['fgj']}+{ref['vatn']})*{ref['fgj_uplift']}+{ref['ins']}+{CAP}*{ref['maint']}", NUM1)
rl('mgmtc', 'Umsýsla', f"=B20*{ref['mgmt']}", NUM1)
rl('noi', 'NOI eiganda við stöðugan rekstur (verðlag 2026)', f"=B20-B22-B23", NUM1, '', True)
rl('noi_key', 'NOI á herbergi', f"=B24/B3", NUM1)
rl('value26', 'Verðmæti við ávöxtunarkröfu (verðlag 2026)', f"=B24/{ref['yld']}", NUM, 'NOI / krafa', True)
rl('capex', 'Framkvæmd alls', f"={CAP}", NUM)
rl('residual', 'Verðmæti – framkvæmd (áður en fjármagnskostnaður, biðtími og ávöxtunarkrafa fjárfesta koma til)', f"=B26-B27", NUM, 'Grófasta vísbending um verðgetu', True)
wr.cell(rr+1, 1, 'Til samanburðar JHB (ár 3): herbergistekjur 1.066 án VSK, EBITDA fyrir leigu 473 (44%), leiga 367 (34% af herbergistekjum), Kolaport 48. Þeirra rekstraráætlun sleppir F&B, þakgarði og meðlimum og gerir ráð fyrir 30% launahlutfalli.').font = Font(name=F, size=9, italic=True)

# ============================================================== SJÓÐSTREYMI
wc = wb.create_sheet('Sjóðstreymi'); setw(wc, [52] + [11]*12)
wc['A1'] = 'Sjóðstreymi eiginfjár – nafnverð, m.kr. Kaup 2027, DSK/hönnun 2027, framkvæmd 2028–29, opnun 2030, endurfjármögnun við stöðugan rekstur 2032, sala 2037.'; wc['A1'].font = H1
years = list(range(2027, 2038)); cols = {y: 2+i for i, y in enumerate(years)}
wc.cell(3, 1, 'Ár').font = BOLD
for y in years: c = wc.cell(3, cols[y], y); c.font = BOLD; c.fill = GREY; c.number_format = '0'
def yc(y): return L(cols[y])
rows = {}
def crow(key, label, fn, fmt=NUM, bold=False):
    r = rows.setdefault('_next', 4); rows[key] = r
    wc.cell(r, 1, label).font = BOLD if bold else BLK
    for y in years:
        c = wc.cell(r, cols[y], fn(y)); c.number_format = fmt; c.font = BOLD if bold else BLK
    rows['_next'] = r+1
def Y(y): return f"{yc(y)}$3"
crow('idx', 'Vísitala', lambda y: f"=(1+{ref['infl']})^({Y(y)}-{ref['buy_year']})", '0.000')
crow('purchase', 'Kaupverð + stimpilgjald + kaupkostnaður', lambda y: f"=IF({Y(y)}={ref['buy_year']},{ref['price']}+{ref['fmat']}*{ref['stimpil']}+{ref['kaupkostn']},0)")
crow('capex', 'Framkvæmd (verðbætt)', lambda y: f"=IF({Y(y)}={ref['works1']},{CAP}*{ref['split1']},IF({Y(y)}={ref['works2']},{CAP}*(1-{ref['split1']}),0))*{yc(y)}{rows['idx']}")
crow('kola', 'Leiga Kolaportsins', lambda y: f"={ref['kola_rent']}*{yc(y)}{rows['idx']}*IF(OR({Y(y)}={ref['works1']},{Y(y)}={ref['works2']}),{ref['kola_works']},1)")
crow('hold', 'Biðtímakostnaður (fasteignagjöld, trygging, annað)', lambda y: f"=IF({Y(y)}<{ref['open_year']},-({ref['fgj']}+{ref['vatn']}+{ref['brunatr']}+{ref['hold_other']})*{yc(y)}{rows['idx']},0)")
crow('ramp', 'Hlutfall stöðugrar leigu', lambda y: f"=IF({Y(y)}<{ref['open_year']},0,IF({Y(y)}={ref['open_year']},{ref['ramp1']},IF({Y(y)}={ref['open_year']}+1,{ref['ramp2']},1)))", PCT)
crow('rent', 'Leiga hótels', lambda y: f"={rref['rent']}*{yc(y)}{rows['ramp']}*{yc(y)}{rows['idx']}")
crow('fixed', 'Fastur eigandakostnaður', lambda y: f"=IF({Y(y)}<{ref['open_year']},0,-{rref['fixed']}*{yc(y)}{rows['idx']})")
crow('mgmt', 'Umsýsla', lambda y: f"=IF({Y(y)}<{ref['open_year']},0,-({yc(y)}{rows['rent']}+{yc(y)}{rows['kola']})*{ref['mgmt']})")
crow('noi', 'NOI', lambda y: f"={yc(y)}{rows['kola']}+{yc(y)}{rows['hold']}+{yc(y)}{rows['rent']}+{yc(y)}{rows['fixed']}+{yc(y)}{rows['mgmt']}", NUM, True)
crow('need', 'Fjárþörf á kaup- og framkvæmdatíma', lambda y: f"=IF({Y(y)}<={ref['works2']},{yc(y)}{rows['purchase']}+{yc(y)}{rows['capex']}-{yc(y)}{rows['noi']},0)")
crow('draw', 'Ádráttur byggingarláns', lambda y: f"=MAX({yc(y)}{rows['need']},0)*{ref['ltc']}")
def loan_f(y):
    prev = f"{L(cols[y]-1)}{rows['_next']}" if y != years[0] else "0"
    return f"=IF({Y(y)}<={ref['works2']},{prev}*(1+{ref['cl_rate']})+{yc(y)}{rows['draw']}*(1+{ref['cl_fee']}),IF({Y(y)}<{ref['open_year']}+2,{prev}*(1+{ref['cl_rate']}),0))"
crow('loan', 'Staða byggingarláns í árslok (PIK), greitt upp við endurfjármögnun', loan_f)
crow('equity', 'Eigið fé inn', lambda y: f"=IF({Y(y)}<={ref['works2']},-(MAX({yc(y)}{rows['need']},0)-{yc(y)}{rows['draw']}),0)")
# refinancing block
r = rows['_next']+1
wc.cell(r, 1, 'Endurfjármögnun við stöðugan rekstur').font = H2; r += 1
wc.cell(r, 1, 'Ár stöðugs rekstrar'); wc.cell(r, 2, f"={ref['open_year']}+2").number_format = '0'; STAB = f"$B${r}"; r += 1
wc.cell(r, 1, 'NOI á stöðugu ári'); wc.cell(r, 2, f"=INDEX($B${rows['noi']}:$L${rows['noi']},1,MATCH({STAB},$B$3:$L$3,0))").number_format = NUM; NOIS = f"$B${r}"; r += 1
wc.cell(r, 1, 'Verðmæti = NOI / krafa'); wc.cell(r, 2, f"={NOIS}/{ref['yld']}").number_format = NUM; VAL = f"$B${r}"; r += 1
wc.cell(r, 1, 'Skuldabréf = verðmæti × LTV'); wc.cell(r, 2, f"={VAL}*{ref['ltv_refi']}").number_format = NUM; BOND = f"$B${r}"; r += 1
wc.cell(r, 1, 'Nafnvextir bréfs'); wc.cell(r, 2, f"=(1+{ref['bond_real']})*(1+{ref['infl']})-1").number_format = PCT2; RN = f"$B${r}"; r += 1
wc.cell(r, 1, 'Árgreiðsla (jafngreiðslulán)'); wc.cell(r, 2, f"=-PMT({RN},{ref['bond_years']},{BOND})").number_format = NUM; ANN = f"$B${r}"; r += 1
wc.cell(r, 1, 'Byggingarlán greitt upp (staða árið á undan)'); wc.cell(r, 2, f"=INDEX($B${rows['loan']}:$L${rows['loan']},1,MATCH({STAB},$B$3:$L$3,0)-1)").number_format = NUM; LOANREP = f"$B${r}"; r += 2
rows['_next'] = r
def bal_f(y):
    prev = f"{L(cols[y]-1)}{rows['_next']}" if y != years[0] else "0"
    return f"=IF({Y(y)}<{STAB},0,IF({Y(y)}={STAB},{BOND}-({ANN}-{BOND}*{RN}),{prev}-({ANN}-{prev}*{RN})))"
crow('bal', 'Staða skuldabréfs í árslok', bal_f)
crow('ds', 'Greiðslubyrði bréfs', lambda y: f"=IF({Y(y)}>={STAB},-{ANN},0)")
crow('refi', 'Endurfjármögnun: bréf – kostnaður – uppgreiðsla byggingarláns', lambda y: f"=IF({Y(y)}={STAB},{BOND}*(1-{ref['bond_cost']})-{LOANREP},0)")
last = years[-1]
crow('exit', 'Sala eftir 10 ár (NOI næsta árs / krafa – sölukostnaður) – uppgreiðsla bréfs', lambda y: f"=IF({Y(y)}={last},({rref['stab_rent']}*{yc(y)}{rows['idx']}*(1+{ref['infl']})-{rref['fixed']}*{yc(y)}{rows['idx']}*(1+{ref['infl']})-{rref['stab_rent']}*{yc(y)}{rows['idx']}*(1+{ref['infl']})*{ref['mgmt']})/{ref['yld']}*(1-{ref['sale_cost']})-{yc(y)}{rows['bal']},0)")
crow('cf', 'Sjóðstreymi til eiginfjár', lambda y: f"=IF({Y(y)}<={ref['works2']},{yc(y)}{rows['equity']},{yc(y)}{rows['noi']})+{yc(y)}{rows['refi']}+{yc(y)}{rows['ds']}+{yc(y)}{rows['exit']}", NUM, True)
r = rows['_next']+1
wc.cell(r, 1, 'IRR eiginfjár (nafnverð, fyrir skatta)').font = BOLD; c = wc.cell(r, 2, f"=IFERROR(IRR($B${rows['cf']}:$L${rows['cf']},0.1),IFERROR(IRR($B${rows['cf']}:$L${rows['cf']},-0.15),\"engin\"))"); c.number_format = PCT; c.font = BOLD; c.fill = YEL; IRRC = r; r += 1
wc.cell(r, 1, 'IRR eiginfjár (raunvirði)'); wc.cell(r, 2, f"=(1+B{IRRC})/(1+{ref['infl']})-1").number_format = PCT; r += 1
wc.cell(r, 1, 'Eigið fé alls inn'); wc.cell(r, 2, f"=-SUMIF($B${rows['equity']}:$L${rows['equity']},\"<0\")").number_format = NUM; r += 1
wc.cell(r, 1, 'Hámarksstaða byggingarláns'); wc.cell(r, 2, f"=MAX($B${rows['loan']}:$L${rows['loan']})").number_format = NUM; r += 1
wc.cell(r, 1, 'Verðmæti við stöðugan rekstur (nafnverð)'); wc.cell(r, 2, f"={VAL}").number_format = NUM; r += 1
wc.cell(r, 1, 'DSCR fyrsta heila ár eftir endurfjármögnun'); wc.cell(r, 2, f"=INDEX($B${rows['noi']}:$L${rows['noi']},1,MATCH({STAB},$B$3:$L$3,0)+1)/{ANN}").number_format = '0.00'; r += 1
wc.cell(r+1, 1, 'Skattar eru ekki reiknaðir (afskriftir og vaxtagjöld skýla afkomu fyrstu árin; sjá Hamranes-líkan). Byggingarlán er PIK eins og hjá ÍSB. Kaupverð er stillt á Forsendur!B9.').font = Font(name=F, size=9, italic=True)
wc.freeze_panes = 'B4'
IRR_CELL = f"Sjóðstreymi!B{IRRC}"

# ============================================================== VERÐGETA (python grid)
wv = wb.create_sheet('Verðgeta'); setw(wv, [6, 8, 10, 10, 8, 10, 10, 10, 10, 10, 10, 10, 10])
wv['A1'] = 'Verðgeta: hæsta kaupverð (m.kr) sem skilar eiginfé tilgreindri IRR – reiknað með model.py v0.1 (sömu forsendur og þetta skjal). Taflan uppfærist EKKI sjálfkrafa.'; wv['A1'].font = H1
wv['A2'] = 'Neikvæð tala = verkefnið stendur ekki undir kaupverði við þessar forsendur (það vantar upp á jafnvel þótt húsið væri gefið).'; wv['A2'].font = Font(name=F, size=9, italic=True)
hd = ['Flokkur', 'Viðbót', 'Kostn.stig', 'Leiguhlutf.', 'Krafa', 'Herb.', 'Framkv.', 'm.kr/herb', 'Leiga+Kola', 'NOI stöðugt', 'Verðmæti', 'P @ 12%', 'P @ 15%']
for j, h in enumerate(hd, 1): c = wv.cell(4, j, h); c.font = BOLD; c.fill = GREY
grid = []
rr = 5
for tier in 'ABC':
    for ef in ('none', 'top', 'full'):
        for cf in (1.0, 0.8):
            for rs in (None, 0.30):
                for y in (0.0675, 0.06):
                    p = dict(M.BASE); p.update(tier=tier, extra_floor=ef, cost_factor=cf, rent_share_override=rs, yld=y)
                    C = M.capex(p); O = M.operations(p); d = M.dcf(p, 0)
                    p12 = M.price_for_irr(p, 0.12); p15 = M.price_for_irr(p, 0.15)
                    p12 = None if p12 < -3900 else p12; p15 = None if p15 < -3900 else p15
                    vals = [tier, ef, cf, (rs if rs else O['rent_share']), y, O['keys'], C['total'], C['per_key'], O['rent']+p['kolaport_rent'], d['noi_stab']/((1+p['infl'])**(d['stab_y']-2027)), d['value_stab']/((1+p['infl'])**(d['stab_y']-2027)), p12, p15]
                    grid.append(dict(tier=tier, extra=ef, cf=cf, rs=vals[3], yld=y, keys=O['keys'], capex=C['total'], per_key=C['per_key'], rent=vals[8], noi=vals[9], value=vals[10], p12=p12, p15=p15))
                    for j, v in enumerate(vals, 1):
                        c = wv.cell(rr, j, v if v is not None else '< −3.900')
                        c.font = BLK; c.number_format = PCT if j in (4, 5) else (NUM if j in (6, 7, 9, 10, 11, 12, 13) else ('0.0' if j == 8 else '0.00'))
                    rr += 1
wv.freeze_panes = 'A5'
json.dump(grid, open('/Users/villithor/repos/tollhusid/grid.json', 'w'), ensure_ascii=False, indent=1)

# ============================================================== SAMKEPPNI
wk = wb.create_sheet('Samkeppni'); setw(wk, [62, 14, 70])
wk['A1'] = 'Hvað gætu aðrir boðið? Tvær einfaldar linsur – m.kr'; wk['A1'].font = H1
kr = 3
def kl(label, val, fmt=NUM, inp_=False, note=''):
    global kr
    wk.cell(kr, 1, label); c = wk.cell(kr, 2, val); c.number_format = fmt; c.font = BLUE if inp_ else BLK
    wk.cell(kr, 3, note).font = Font(name=F, size=9, color='808080'); kr += 1; return f"B{kr-1}"
wk.cell(kr, 1, 'A. Fasteignafélag (Reitir/Heimar/Regin) – skrifstofur/blönduð notkun eftir endurbætur').font = H2; kr += 1
o_m2 = kl('Útleigjanlegt skrifstofu-/þjónusturými (án Kolaports)', 7300, NUM, True, 'Hótelhluti 7.741 m² brúttó, ~94% útleigjanlegt')
o_rent = kl('Leiga nýuppgerðs skrifstofuhúsnæðis í Kvosinni, kr/m²/mán án VSK', 4800, NUM, True, 'Hafnartorg/Austurhöfn 5.500–7.000; eldri uppgerð 3.800–4.800')
o_kola = kl('Leiga Kolaportsins', f"={ref['kola_rent']}", NUM1)
o_gross = kl('Leigutekjur alls', f"={o_m2}*{o_rent}*12/1000000+{o_kola}", NUM1)
o_noi = kl('NOI (85% af leigu)', f"={o_gross}*0.85", NUM1)
o_y = kl('Ávöxtunarkrafa félags', 0.0625, PCT2, True, 'Reitir nýkaup 7,8%; matskrafa 6,4–6,7%; hér 6,25% fyrir miðborgareign')
o_val = kl('Verðmæti fullbúið', f"={o_noi}/{o_y}", NUM)
o_cap = kl('Endurbótakostnaður til skrifstofustaðals (byggingarhluti + 200 þ/m² innanhúss + 20% soft/ófyrirséð)', f"=(Framkvæmd!E{fr['building']}-Framkvæmd!E{fr['extra']}+{o_m2}*200/1000)*1.2", NUM, False, 'Byggingarhlutinn er sá sami og hjá okkur')
o_marg = kl('Þróunarálag félags (hlutfall af verðmæti)', 0.10, PCT, True)
o_res = kl('Verð sem félag gæti boðið', f"={o_val}-{o_cap}-{o_val}*{o_marg}", NUM)
wk.cell(kr-1, 2).fill = YEL; kr += 1
wk.cell(kr, 1, 'B. Hótelfélag sem á og rekur sjálft (Íslandshótel, Berjaya, Keahótel)').font = H2; kr += 1
h_ebitda = kl('EBITDA hótels fyrir leigu (Rekstur)', f"={rref['ebitda']}", NUM1)
h_res = kl('FF&E-sjóður', f"={rref['reserve']}", NUM1)
h_noi = kl('Sjóðstreymi eiganda-rekstraraðila + Kolaport – fastur kostnaður', f"={h_ebitda}-{h_res}+{rref['kola']}-{rref['fixed']}", NUM1)
h_y = kl('Ávöxtunarkrafa á rekstrarsjóðstreymi hótels', 0.085, PCT2, True, 'Hærri en á leigusamning: rekstraráhætta hjá eiganda')
h_val = kl('Verðmæti', f"={h_noi}/{h_y}", NUM)
h_cap = kl('Framkvæmd (sama og okkar)', f"={CAP}", NUM)
h_marg = kl('Krafa um þróunarálag', 0.15, PCT, True)
h_res2 = kl('Verð sem hótelfélag gæti boðið', f"={h_val}-{h_cap}-{h_val}*{h_marg}", NUM)
wk.cell(kr-1, 2).fill = YEL; kr += 1
wk.cell(kr, 1, 'Viðmið seljanda: fasteignamat 4.725 (2026) / 4.915 (2027); mat ríkisins sjálfs 2022 „a.m.k. 2 ma.kr“; brunabótamat 6.000. Ríkið selur eftir fjárlagaheimild og velur að öllu jöfnu hæsta gilda tilboð, en má hafna öllum.').font = Font(name=F, size=9, italic=True)

# ============================================================== HEIMILDIR
wh = wb.create_sheet('Heimildir'); setw(wh, [120])
srcs = [
    'Söluyfirlit TORG 31.8.2026 (Tryggvagata 19, F2000241) – stærðir, mat, gjöld, kvaðir, lóðarleigusamningur útrunninn.',
    'Fasteigna- og veðbandayfirlit HMS 8.7./21.7.2026 – rekstrareiningar, lóðarleigusamningur 411-F-18055 (50 ár frá 1.1.1967), kvaðir hafnarstjórnar.',
    'COWI: Ástandsskýrsla 3210316-000-CRP-0003, 16.10.2024 (ytra byrði, burðarþol, asbest) + minnisblað um burðarvirki 6.9.2024.',
    'COWI: Innivistarskoðun A290951-001-HMO-20.08.25, 20.8.2025 (mygla, asbest, PCB, frárennsli, forgangslisti 1–14).',
    'Jón Haukur Baldvinsson: Vinnuskjal stofnkostnaðar og rekstrar-/söluáætlun (Numbers-útflutningur), Næsland-kynning, samantekt rauðra flagga 8.9.2026 – póstur 22.9.2026.',
    'Umsögn skipulagsfulltrúa Reykjavíkur 4.4.2024 (Tryggvagata 19, LHÍ) – AR2040 M1a, deiliskipulag Kvosar 1988/2015/2020, landnotkun skrifstofur/vörugeymsla, ein hæð til viðbótar möguleg, DSK-breyting nauðsynleg.',
    'Minnisblað vinnuhóps ríkis og borgar 20.4.2022 (MSS22040217) – virði Tollhúss a.m.k. 2 ma.kr, LHÍ-umbreyting 12,8 ma.kr; Stjórnarráðið 11.9.2024: 17 ma.kr, hætt við.',
    'Reykjavíkurborg: auglýsing eftir rekstraraðila Kolaportsins (frestur 8.7.2025) – lágmarksleiga 3.777.457 kr/mán; Vísir 12.1. og 16.6.2025.',
    'Hamranes fjárhagslíkan v2.2 / FORSENDUSKJAL.md – ÍSB byggingarlán kjörvextir + 0,15% PIK, LTV 67,7%, verðtryggt bréf 2,86% + 1,0%, krafa 5,7%.',
    'Ávöxtunarkrafa fasteignafélaga – greining 19.8.2026: Reitir/Heimar/Eik/Kaldalón matskröfur 6,4–6,7%, Reitir nýkaup 7,8% (H1 2026), ríkisleiga 5,5–6,0%.',
    'Þ113 Hamrar model.py v1.5 – Hamranes EAC einingarverð (kafli 3 24,8; kafli 4 49,5; kafli 5 38,0; kafli 7 40,8 þ.kr/m²).',
    'KEF Airport Hotel PME-svar ÍF júní 2026 – 30–36 m.kr/herb, 600–750 þ.kr/m² nýbygging án lóðar.',
    'Hagstofa Íslands: herbergjanýting höfuðborgarsvæðis 74,7% 2025; 58 hótel / 5.555 herbergi. Reitir–Íslandshótel leigusamningur 29.4.2026 (470 herb., 26.500 m², NOI +620/720 m.kr).',
    'DV 7.9.2026 / mbl 7.9.2026: söluheimild fyrir Tollhúsið í fjárlagafrumvarpi 2027.',
]
wh['A1'] = 'Heimildir'; wh['A1'].font = H1
for i, s_ in enumerate(srcs, 3): wh.cell(i, 1, f'{i-2}. {s_}').alignment = Alignment(wrap_text=True)

for w in wb.worksheets:
    for row_ in w.iter_rows():
        for c in row_:
            if c.font is None or c.font.name != F: c.font = Font(name=F, size=c.font.size if c.font and c.font.size else 10, bold=c.font.bold if c.font else False, italic=c.font.italic if c.font else False, color=c.font.color if c.font else None)
wb.save(OUT); print('saved', OUT, 'IRR cell', IRR_CELL)

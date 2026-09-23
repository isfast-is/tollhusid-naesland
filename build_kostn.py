# -*- coding: utf-8 -*-
"""Tollhúsið – Næsland: ítarleg kostnaðaráætlun fullbúins hótels v1.0 í ÍF-sniði (kaflar 0–8 + hönnun + búnaður).
Allt formúlur. Verðlag 2026, án VSK nema annað sé tekið fram. Heimildir við hverja línu (innanhúss-útgáfa)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
BVT_MAI25 = float(sys.argv[2]) if len(sys.argv) > 2 else 200.0
wb = Workbook()
F = 'Arial'
BLUE = Font(name=F, size=10, color='0000FF'); BLK = Font(name=F, size=10); GRN = Font(name=F, size=10, color='008000')
BOLD = Font(name=F, size=10, bold=True); H1 = Font(name=F, size=14, bold=True); H2 = Font(name=F, size=11, bold=True, color='1F3864')
SRC = Font(name=F, size=9, color='808080'); YEL = PatternFill('solid', fgColor='FFFF00'); GREY = PatternFill('solid', fgColor='D9E2F3'); LIGHT = PatternFill('solid', fgColor='EEF3F8')
NUM = '#,##0;(#,##0);-'; NUM1 = '#,##0.0;(#,##0.0);-'; PCT = '0.0%'
def setw(ws, widths):
    for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w

# ============================================================ STÆRÐIR
wsS = wb.active; wsS.title = 'Stærðir'
setw(wsS, [44, 14, 14, 14, 14, 70])
wsS['A1'] = 'Tollhúsið, Tryggvagötu 19 – flatarmál og skipting í hótelhluta, tæknirými og Kolaport'; wsS['A1'].font = H1
wsS['A2'] = 'Brúttóflatarmál skv. fasteignaskrá (rekstrareiningar 01-0001…01-0601) og aðaluppdráttum BN045618/BN048223. Bláar tölur eru forsendur.'; wsS['A2'].font = SRC
hdr = ['Hæð / eining', 'Brúttó m²', 'Hótel, breytt', 'Tæknirými / BOH', 'Kolaport (utan)', 'Heimild / rök']
for j, h in enumerate(hdr, 1): c = wsS.cell(4, j, h); c.font = BOLD; c.fill = GREY
rows = [
    ('Kjallari (01-0001 skjalageymsla 684 + færiband 26,4)', 710.4, 0, 710.4, 0, 'Fasteignaskrá. Verður þvottahús, geymslur, tæknirými: gert nothæft, ekki innréttað sem gestarými.'),
    ('1. hæð: stigagangar og lyftur (01-0100, 01-0102, 01-0107)', 113.5, 0, 113.5, 0, 'Fasteignaskrá 90,3 + 12 + 11,2'),
    ('1. hæð: afgreiðsla tollstjóra (01-0101)', 730.0, 730.0, 0, 0, 'Austurblokk við Pósthússtræti/Steinbryggju: lobby, móttaka, veitingastaður, bar. Grunnmynd 1. hæðar.'),
    ('1. hæð: uppboðssalur (01-0103) + lögregla (01-0104)', 351.9, 351.9, 0, 0, '218,4 + 133,5: eldhús, fundarrými, starfsmannaaðstaða. Gamla lögreglustöðin þarfnast algjörrar endurnýjunar (COWI 2025 kafli 2.3).'),
    ('1. hæð: vörugeymsla (01-0105)', 1417.1, 0, 0, 1417.1, 'Kolaportssalur. UTAN áætlunar nema COWI-viðgerðir (kafli 7).'),
    ('1. hæð: „hugsað fyrir Kolaport" (01-0106)', 992.5, 0, 0, 992.5, 'Kolaportssalur. UTAN áætlunar.'),
    ('2. hæð (01-0200…01-0206)', 1591.4, 1591.4, 0, 0, 'Norðurhlið álmunnar liggur inn í Kolaportssalinn: einhliða gangur, 14–17 herbergi + gym/spa.'),
    ('3. hæð, inndregin (01-0300…01-0304)', 943.3, 943.3, 0, 0, 'Lokað 1993, léttir útveggir. 24–28 herbergi; norðurhlið út á þakflöt.'),
    ('4. hæð (01-0400…01-0404)', 1620.0, 1620.0, 0, 0, '34–39 herbergi, tvíhliða gangur.'),
    ('5. hæð (01-0500…01-0504)', 1620.0, 1620.0, 0, 0, '34–39 herbergi.'),
    ('Lyftuhús 6. hæð (01-0601)', 61.2, 0, 61.2, 0, 'Tæknirými lyfta/loftræsting.'),
]
r = 5
for name, tot, hot, boh, kola, src in rows:
    wsS.cell(r, 1, name); c = wsS.cell(r, 2, tot); c.font = BLUE; c.number_format = NUM1
    for j, v in ((3, hot), (4, boh), (5, kola)):
        c = wsS.cell(r, j, v); c.font = BLUE; c.number_format = NUM1
    wsS.cell(r, 6, src).font = SRC; r += 1
last = r - 1
wsS.cell(r, 1, 'SAMTALS').font = BOLD
for j in range(2, 6):
    c = wsS.cell(r, j, f'=SUM({L(j)}5:{L(j)}{last})'); c.font = BOLD; c.number_format = NUM1
SUMROW = r; r += 2
wsS.cell(r, 1, 'Hótelhluti alls, byggður (breytt rými + tæknirými) – GRUNNUR fyrir kr/m²').font = BOLD
c = wsS.cell(r, 2, f'=C{SUMROW}+D{SUMROW}'); c.font = BOLD; c.number_format = NUM1; c.fill = YEL; A_HOTEL = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Þar af breytt í gestarými (2.–5. hæð + jarðhæð austur/vestur)'); c = wsS.cell(r, 2, f'=C{SUMROW}'); c.number_format = NUM1; A_CONV = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Þar af tæknirými / BOH (kjallari, stigar, lyftuhús)'); c = wsS.cell(r, 2, f'=D{SUMROW}'); c.number_format = NUM1; A_BOH = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Kolaportssalur – utan áætlunar'); c = wsS.cell(r, 2, f'=E{SUMROW}'); c.number_format = NUM1; A_KOLA = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Herbergjahæðir 2.–5. hæð'); c = wsS.cell(r, 2, f'=C11+C12+C13+C14'); c.number_format = NUM1; A_FLOORS = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Jarðhæð hótels (lobby, veitingar, eldhús, fundir)'); c = wsS.cell(r, 2, f'=C7+C8'); c.number_format = NUM1; A_GROUND = f"Stærðir!$B${r}"; r += 2
wsS.cell(r, 1, 'Byggingarvísitala mars 2018 (grunnur 2009)'); c = wsS.cell(r, 2, 137.0); c.font = BLUE; c.number_format = NUM1; BVT18 = f"Stærðir!$B${r}"; wsS.cell(r, 6, 'Hagstofa, VIS13001: 2018M03 = 137,0 (Grensásvegs-áætlun dags. mars 2018)').font = SRC; r += 1
wsS.cell(r, 1, 'Byggingarvísitala maí 2025 (grunnur 2009)'); c = wsS.cell(r, 2, BVT_MAI25); c.font = BLUE; c.number_format = NUM1; BVT25 = f"Stærðir!$B${r}"; wsS.cell(r, 6, 'Hagstofa, VIS13001: 2025M05 (Hamranes EAC á verðlagi maí 2025)').font = SRC; r += 1
wsS.cell(r, 1, 'Byggingarvísitala september 2026 (grunnur 2009)'); c = wsS.cell(r, 2, 206.9); c.font = BLUE; c.number_format = NUM1; BVT26 = f"Stærðir!$B${r}"; wsS.cell(r, 6, 'Hagstofa, VIS13001: 2026M09 = 206,9 (2026M08 = 206,7; á 2021-grunni 129,6). Hyatt EAC júlí 2026 er á sama verðlagi.').font = SRC; r += 1
wsS.cell(r, 1, 'Uppreikningsstuðull Grensásvegur 2018 → 2026').font = BOLD; c = wsS.cell(r, 2, f'={BVT26}/{BVT18}'); c.number_format = '0.000'; c.font = BOLD; F18 = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Uppreikningsstuðull Hamranes 2025 → 2026'); c = wsS.cell(r, 2, f'={BVT26}/{BVT25}'); c.number_format = '0.000'; F25 = f"Stærðir!$B${r}"; r += 2
wsS.cell(r, 1, 'Herbergi (forsenda, sjá talningu af grunnmyndum)').font = BOLD; c = wsS.cell(r, 2, 123); c.font = BLUE; c.fill = YEL; c.number_format = '0'; KEYS = f"Stærðir!$B${r}"
wsS.cell(r, 6, 'Talning: 5. hæð 39, 4. hæð 39, 3. hæð 28, 2. hæð 17 við 3,9 m breið herbergi (27–30 m² nettó); 106 við 4,5 m; Næsland-skipan 89.').font = SRC; r += 1
wsS.cell(r, 1, 'Brúttó m² á herbergi (hótelhluti / herbergi)'); c = wsS.cell(r, 2, f'={A_HOTEL}/{KEYS}'); c.number_format = NUM1; r += 1
wsS.cell(r, 1, 'Meðalstærð herbergis nettó, m²'); c = wsS.cell(r, 2, 28); c.font = BLUE; c.number_format = NUM1; ROOM_M2 = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Herbergjaflötur alls, nettó m² (herbergi × meðalstærð)'); c = wsS.cell(r, 2, f'={KEYS}*{ROOM_M2}'); c.number_format = NUM1; A_ROOMS = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Gangar, kjarnar og þjónusturými á herbergjahæðum (afgangur)'); c = wsS.cell(r, 2, f'={A_FLOORS}-{A_ROOMS}'); c.number_format = NUM1; A_CORR = f"Stærðir!$B${r}"; r += 1
wsS.cell(r, 1, 'Baðherbergi, m² hvert (inni í meðalstærð)'); c = wsS.cell(r, 2, 4.5); c.font = BLUE; c.number_format = NUM1; BATH_M2 = f"Stærðir!$B${r}"; r += 2
wsS.cell(r, 1, 'Svar við spurningunni um Kolaportið: Kolaportssalurinn (2.410 m²) er hvergi inni í kostnaði hótelsins. Hann kemur aðeins fyrir sem sérmerkt COWI-viðgerðarlína í kafla 7 (þak, gaflar, hringhurð), sem má strika út ef borgin/rekstraraðili ber hana. Kr/m² eru reiknuð á hótelhlutann einan, 7.741 m² með tæknirýmum.').font = Font(name=F, size=10, italic=True)

# ============================================================ kaflablöð
CH = {}   # chapter key -> (sheet, total cell ref)
def chapter(key, title, lines, note=''):
    ws = wb.create_sheet(key); setw(ws, [6, 58, 8, 12, 14, 16, 78])
    ws['A1'] = title; ws['A1'].font = H1
    ws['A2'] = 'Tölur án VSK, verðlag 2026. Magn og einingarverð blá (forsendur); upphæð = magn × einingarverð. Heimild/rök í G.'; ws['A2'].font = SRC
    for j, h in enumerate(['Nr', 'Verkþáttur', 'Ein.', 'Magn', 'Ein.verð kr', 'Upphæð kr', 'Heimild / rök'], 1):
        c = ws.cell(4, j, h); c.font = BOLD; c.fill = GREY
    r = 5; n = 0; sub_start = None
    for ln in lines:
        if isinstance(ln, str):   # subheading
            ws.cell(r, 2, ln).font = H2; r += 1; continue
        name, unit, qty, rate, src = ln
        n += 1
        ws.cell(r, 1, f'{key.split()[0]}.{n}').font = SRC
        ws.cell(r, 2, name)
        ws.cell(r, 3, unit)
        q = ws.cell(r, 4, qty); q.font = GRN if isinstance(qty, str) and qty.startswith('=') else BLUE; q.number_format = NUM1 if not (isinstance(qty, (int, float)) and float(qty).is_integer()) else NUM
        p = ws.cell(r, 5, rate); p.font = GRN if isinstance(rate, str) and rate.startswith('=') else BLUE; p.number_format = NUM
        u = ws.cell(r, 6, f'=D{r}*E{r}'); u.number_format = NUM
        ws.cell(r, 7, src).font = SRC; ws.cell(r, 7).alignment = Alignment(wrap_text=False)
        r += 1
    ws.cell(r, 2, 'SAMTALS KAFLI').font = BOLD
    t = ws.cell(r, 6, f'=SUM(F5:F{r-1})'); t.font = BOLD; t.number_format = NUM; t.fill = LIGHT
    ws.cell(r + 1, 2, 'kr/m² hótelhluta'); c = ws.cell(r + 1, 6, f'=F{r}/{A_HOTEL}'); c.number_format = NUM
    ws.cell(r + 2, 2, 'kr á herbergi'); c = ws.cell(r + 2, 6, f'=F{r}/{KEYS}'); c.number_format = NUM
    if note: ws.cell(r + 4, 2, note).font = Font(name=F, size=9, italic=True)
    ws.freeze_panes = 'A5'
    CH[key] = f"'{key}'!$F${r}"
    return ws

HY = 'Hyatt L176 EAC júlí 2026 (trúnaðarmál ÍF/Reitir)'
GR = 'ÍF kostnaðaráætlun Grensásvegur 16A 2018 (78 herb.), uppreiknað með byggingarvísitölu 137,0 → 206,9 (×1,51)'
HAM = 'Hamranes EAC 2025 (nýbygging, einingar)'
NL = 'Næsland vinnuskjal stofnkostnaðar (m/VSK, 100 herb.)'
CO24 = 'COWI ástandsskýrsla 16.10.2024'
CO25 = 'COWI innivistarskoðun 20.8.2025'

chapter('0 Aðstaða', '0. Aðstaða, umsjón og rekstur vinnusvæðis', [
    ('Aðstöðusköpun: girðingar við torg og Naustin, skúrar, hlið, merkingar', 'heild', 1, 15000000, f'{GR}: 8 m.kr 2018 fyrir lítið verk; hér miðborgartorg og tvær aðkomur'),
    ('Rekstur vinnustaðar (verkstjórn verktaka, þrif, sorp, öryggi)', 'mán', 20, 8000000, f'{HY}: 427 m.kr á ~44 mán ≈ 9,7 m.kr/mán með aukaverkum; {GR}: 3,0 m.kr/mán 2018'),
    ('Byggingarkrani / efnislyfta / vörulyfta á verktíma', 'mán', 18, 1400000, 'Mat ÍF; ein krani á þakfleti 3. hæðar + efnislyfta í stigakjarna'),
    ('Upphitun og rafmagn á verktíma', 'mán', 20, 1200000, f'{GR}: 0,5 m.kr/mán 2018 fyrir 3.400 m²; hér 7.700 m²'),
    ('Öryggisgæsla og lokanir gagnvart Kolaporti og torgi', 'mán', 20, 600000, 'Kolaportið í rekstri um helgar á framkvæmdatíma'),
    ('Lokaþrif og afhending', 'm²', f'={A_HOTEL}', f'=2000000/3400*{F18}*1.6', f'{GR}: 2 m.kr fyrir 3.400 m² 2018'),
    ('Framkvæmdatrygging og verktrygging (byggingarstjóri)', 'heild', 1, 12000000, 'Sjóvá/VÍS viðmið ~0,3% af verksamningi'),
], 'Hyatt-viðmið: 50 þ.kr/m² (með 239 m.kr aukaverkum). Hamranes: 30 þ.kr/m². Hér ≈ 30–35 þ.kr/m².')

chapter('1 Niðurrif', '1. Niðurrif innanhúss, afmengun og förgun', [
    'Niðurrif',
    ('Niðurrif innanhúss í gestarýmum: milliveggir, loft, gólfefni, lagnir, innréttingar', 'm²', f'={A_CONV}', f'=5000*{F18}*1.2', f'{GR}: rif 1,8–2,0 m.kr á 400 m² hæð 2018 = 5 þ/m² → 7 þ/m² 2026; hér þyngra (dökk viðarþil, skjalageymslur, mötuneyti). {CO25} kafli 2.3, 2.7'),
    ('Niðurrif og hreinsun í tæknirýmum / kjallara', 'm²', f'={A_BOH}', 4000, 'Léttara: geymsluinnréttingar, gamlar lagnir'),
    ('Förgun og flutningur úrgangs (gámar, urðun)', 'm²', f'={A_CONV}', 3000, 'Mat ÍF; miðborg, engin lóð fyrir gáma nema á þakfleti/torgi'),
    ('Kjarnaborun og op í plötur fyrir nýjar lagnir (stokkar, loftræsing)', 'stk', 80, 220000, f'{GR}: kjarnaborun 1,5 m.kr heild 2018 fyrir 5 hæðir'),
    'Afmengun – COWI',
    ('Asbestplötur í gluggaveggjaeiningum: niðurtekt af leyfishöfum, innpökkun, förgun', 'm²', 600, 45000, f'{CO24} viðauki B: 550–600 m² asbest; vinna leyfishafa í samráði við Vinnueftirlit og Heilbrigðiseftirlit'),
    ('Asbest í gólfflísum (matsalur 2. hæð, geymslur, salerni, gangar gömlu lögreglustöðvar)', 'm²', 450, 22000, f'{CO25} kafli 2.2–2.3, sýni E-14; magn áætlað'),
    ('PCB og þungmálmar í gluggamálningu: meðhöndlun sem mengað byggingarefni við frárif', 'heild', 1, 25000000, f'{CO25} samantekt: „meðhöndla þarf alla glugga sem mengað byggingarefni"'),
    ('Mygla og rakaskemmdir: opnanir, afmengun, fjarlæging plötuklæðninga og gólfdúka, þurrkun, lokaprófanir', 'heild', 1, 60000000, f'{CO25}: mygla í 8 af 33 efnissýnum, örveruvöxtur á öllum hæðum nema 5.; matsalur, Brúarsel, lagnastokkur 3. hæð, 4. hæð'),
    ('Lyktarmengun 4. hæð: gólfdúkar fjarlægðir og undirlag hreinsað', 'm²', 1620, 2500, f'{CO25} forgangsatriði 7'),
], 'Næsland: niðurrif 34 m.kr m/VSK, engin afmengun. Þessi kafli er að mestu COWI-drifinn.')

chapter('2 Burðarvirki', '2. Burðarvirki, op, stigar og frárennsli í grunni', [
    'Breytingar á burðarvirki',
    ('Ný op í plötur fyrir lyftur/stiga og styrkingar með stálbitum', 'stk', 6, 9000000, 'Mat ÍF; lyftuop 3 × 5 hæðir, nýir flóttastigar. Steypa >35 MPa, járnabinding skv. teikningum (COWI 2024 viðauki A)'),
    ('Nýir steyptir/stál flóttastigar í kjörnum þar sem þarf', 'stk', 2, 16000000, f'{HAM}/Þ113: utanáliggjandi stálstigar 3 hæðir ≈ 11 m.kr; hér innanhúss 5 hæðir'),
    ('Styrking þakplötu 3. hæðar fyrir þakgarð (jöfnun yfirborðs, ílögn, álag)', 'm²', 1500, 15000, 'COWI minnisblað 6.9.2024: plata þolir veitingaálag 2–3 kN/m²; jafna þarf halla að niðurföllum og huga að þyngd ílagnar'),
    ('Steypuviðgerðir innanhúss við op og lagnastokka', 'heild', 1, 15000000, 'Mat ÍF'),
    'Frárennsli í grunni – COWI forgangsatriði 1',
    ('Uppbrot gólfplötu 1. hæðar og kjallara yfir lagnaleiðum', 'm²', 600, 25000, f'{CO25} viðauki C: pottlagnir mjög ryðgaðar og bólgnar, myndavél komst ekki áfram; magn áætlað þar til myndun og kortlagning liggur fyrir'),
    ('Nýjar frárennslislagnir í grunni með tengingu við götulögn', 'lm', 260, 120000, 'Mat ÍF; Akureyrar-einingarverð lagnaskurða 8–11 þ.kr/lm + rör + tengingar innanhúss'),
    ('Brunnar og niðurföll', 'stk', 10, 800000, 'Mat ÍF'),
    ('Endursteypa og frágangur gólfplötu eftir lagnavinnu', 'm²', 600, 20000, f'Akureyri: plötumót 14.500 + steypa 55.000/m³ + járn'),
    ('Álag vegna vinnu undir Kolaporti í rekstri (áfangaskipting, helgarlokanir)', 'heild', 1, 20000000, 'Kolaportið opið um helgar; borgin leigutaki'),
], 'Næsland: „frárennslislagnir í grunni" óverðlagt. COWI: „skynsamlegt að framkvæma meðan engin starfsemi er í húsinu".')

chapter('3 Lagnir', '3. Pípulagnir, hitakerfi, loftræsing og vatnsúðakerfi', [
    'Pípulagnir',
    ('Baðherbergi: neysluvatn, frárennsli, tengingar, hreinlætistæki', 'herb', f'={KEYS}', 950000, f'Næsland: salerni 208 + handlaug 160 + baðkar 240 þ.kr m/VSK = 490 án VSK aðeins tæki; {GR}: lagnir í bað 2018; {HY} 7. hæð FOH lagnir og loftræsing 95 þ.kr/m²'),
    ('Stofnlagnir, stammar og lagnastokkar milli hæða (neysluvatn, frárennsli, hiti)', 'm²', f'={A_CONV}', 11000, f'{HAM} kafli 3 (lagnir + loftræsing) 24,8 þ.kr/m² nýbygging með einingum; hér allt nýtt utan eininga'),
    ('Hitakerfi: ofnar/gólfhiti í herbergjum og handklæðaofnar', 'herb', f'={KEYS}', 260000, 'Næsland: ofnakerfi 200 + handklæðaofn 35 þ.kr m/VSK á herbergi'),
    ('Hitakerfi í almennum rýmum, göngum og BOH', 'm²', f'={A_GROUND}+{A_CORR}+{A_BOH}', 9000, 'Mat ÍF'),
    ('Ofnalagnir í gólfstokk 4. hæðar austur endurnýjaðar (ryð vegna leka)', 'heild', 1, 8000000, f'{CO25} forgangsatriði 6'),
    ('Inntök hitaveitu og vatns, varmaskiptar, dælur, tæknirými', 'heild', 1, 28000000, 'Mat ÍF; núverandi inntök fyrir skrifstofur, hótel þarf mun meira heitt vatn'),
    ('Eldhús: lagnir, fituskilja, gufugleypir tengingar', 'heild', 1, 18000000, 'Mat ÍF; Hyatt eldhústæki 7. hæð BOH 10 m.kr aðeins tæki'),
    'Loftræsing',
    ('Loftræsing herbergja (VAV/fan-coil eða útsog + innblástur um ganga)', 'herb', f'={KEYS}', 380000, f'{HY}: fan coil viðbót 7. hæð; Næsland: útsog 120 stöðum 10,4 m.kr + loftræsting 4 m.kr m/VSK er ekki nothæft'),
    ('Loftræsisamstæður, stokkar og stýringar fyrir almenn rými, eldhús, spa/gym', 'm²', f'={A_GROUND}', 55000, f'{CO25}: keypt samstæða óuppsett í miðrými 2. hæðar dugar ekki fyrir hótel; stýringar úreltar'),
    ('Loftræsing ganga, kjarna og BOH', 'm²', f'={A_CORR}+{A_BOH}', 14000, 'Mat ÍF'),
    'Brunavarnir',
    ('Vatnsúðakerfi í öllu húsi (hótelhluti)', 'm²', f'={A_HOTEL}', 13000, 'Næsland 12,8 þ.kr/m² m/VSK; mat ÍF svipað án VSK með dælustöð'),
    ('Reyklosun og brunalokur', 'heild', 1, 12000000, 'Brunahönnun hótels; stigakjarnar'),
], f'Viðmið: {HY} kafli 3 = 78 þ.kr/m² (öll bygging), 7. hæð FOH 95 þ.kr/m². Hamranes 24,8 (einingar með lögnum).')

chapter('4 Raflagnir', '4. Raflagnir, lýsing, öryggis- og hússtjórnarkerfi, lyftur', [
    ('Raflagnir og lýsing í herbergjum (tenglar, rofar, lýsing, hótelkortakerfi, sjónvarpstengi)', 'herb', f'={KEYS}', 780000, f'{HY} 7. hæð FOH rafmagn 119 þ.kr/m² × ~30 m² ≈ 3,6 m.kr á lúxusherbergi með AV; hér upper upscale án AV'),
    ('Raflagnir og lýsing í almennum rýmum (lobby, veitingar, eldhús, fundir, spa)', 'm²', f'={A_GROUND}', 60000, f'{HY} FOH 119 þ.kr/m² lúxus; hér 60'),
    ('Raflagnir og lýsing á göngum, kjörnum og BOH', 'm²', f'={A_CORR}+{A_BOH}', 28000, 'Mat ÍF'),
    ('Aðaltafla, stofnlagnir, dreifitöflur á hæðum, rafmagnsinntak (aukin heimild)', 'heild', 1, 60000000, f'{CO25}: töflur í ágætu eftirliti en þarf tiltekt; hótel þarf mun meira afl (eldhús, loftræsing)'),
    ('Brunaviðvörunarkerfi og neyðarlýsing', 'm²', f'={A_HOTEL}', 5000, 'Mat ÍF; krafa í hóteli'),
    ('Aðgangsstýring, læsingakerfi og öryggismyndavélar', 'herb', f'={KEYS}', 130000, f'{HY}: læsingarkerfi 28 m.kr fyrir 169 herb. = 167 þ.kr/herb (lúxus)'),
    ('Netkerfi, WiFi, sjónvarpsdreifing, hótelkerfi (PMS-tengingar)', 'm²', f'={A_HOTEL}', 4500, f'{HY}: netkerfi 30 m.kr fyrir 9.923 m²'),
    ('Hússtjórnarkerfi (BMS) og orkumælingar', 'heild', 1, 25000000, 'Mat ÍF'),
    ('Lyftur, 3 stk (2 gestalyftur + 1 þjónustulyfta), 6 stopp', 'stk', 3, 32000000, f'Næsland 24 m.kr m/VSK; {HY} lyftur kostnaðaraukning 22 m.kr; {CO25}: núverandi lyfta gömul og bilanagjörn'),
    ('Rafmagn á þakgarði, útilýsing, skilti', 'heild', 1, 15000000, 'Mat ÍF'),
], f'Viðmið: {HY} kafli 4 = 113 þ.kr/m² (með lúxus AV og lýsingu). Hamranes 49,5.')

chapter('5 Innanhúss', '5. Frágangur innanhúss', [
    'Herbergi (magn = herbergjafjöldi; einingarverð á herbergi)',
    ('Hljóðeinangraðir gipsveggir milli herbergja og að gangi (EI60, 55 dB)', 'herb', f'={KEYS}', f'=30*14200*{F18}', f'{GR}: gifsveggir hljóðkrafa 14.200 kr/m² 2018 → 20.600 2026; ~30 m² á herbergi'),
    ('Baðherbergi: flísar veggir og gólf, sturtugler, innrétting, spegill, fylgihlutir', 'herb', f'={KEYS}', f'=25*14000*{F18}+900000', f'Næsland: flísar 30 þ/m² m/VSK, fastur búnaður bað 200 þ; {GR}: flísar 14.000 kr/m² 2018; 4,5 m² bað → 25 m² flísar × 24 þ = 0,6 + innrétting/gler 0,6 + fylgihlutir 0,3'),
    ('Herbergishurð EI30 með hótellæsingu og baðhurð', 'herb', f'={KEYS}', 520000, 'Næsland: 412,5 + 180 þ.kr m/VSK = 478 án VSK'),
    ('Gólfefni herbergis (parket/teppi) með undirlagi og flotun', 'herb', f'={KEYS}', f'=24*(9500+1800)*{F18}', f'Næsland: parket 35 þ/m² m/VSK; {GR}: parket 9.500 kr/m² 2018 → 13.800; 24 m² × 14 þ + flotun'),
    ('Loft: gifs/hljóðloft með lýsingaropum', 'herb', f'={KEYS}', f'=26*7100*{F18}', f'{GR}: föst gifsloft 7.100 kr/m² 2018 → 10.300; Næsland 25 þ/m² m/VSK'),
    ('Málun og spörtlun veggja og lofta', 'herb', f'={KEYS}', f'=55*2800*{F18}', f'{GR}: 2.800 kr/m² 2018 → 4.100; ~55 m² fletir'),
    ('Fastar innréttingar: fataskápur, höfuðgafl, skrifborð/hilla, minibar-skápur, gluggabekkur', 'herb', f'={KEYS}', 1100000, 'Næsland: fastur búnaður herbergi 400 þ.kr m/VSK er lágt; upper upscale sérsmíði'),
    ('Frágangur við nýja gluggaveggi að innan (kistur, sólarvarnir, gluggabekkir)', 'herb', f'={KEYS}', 180000, 'Mat ÍF'),
    'Gangar, kjarnar og stigahús',
    ('Gangar: teppi, loft, veggklæðning, hurðir að stigahúsum, lýsingarop, skilti', 'm²', f'={A_CORR}', f'=(11000+9100+2*2800)*{F18}+25000', f'{GR}: teppi 11.000 kr/m² 2018, kerfisloft 9.100; {HY} FOH 300 þ.kr/m² lúxus; hér millistig'),
    ('Stigahús og lyftuforrými: steinflísar, handrið, málun, hurðir', 'hæð', 10, 4500000, 'Tveir kjarnar × 5 hæðir; Næsland: steinflísar 20 þ/m² m/VSK, hurðir 375 þ'),
    'Almenn rými á jarðhæð',
    ('Lobby, móttaka, setustofa, bókasafn/búð: gólf, loft, veggir, sérsmíði, lýsingarfrágangur', 'm²', 500, 260000, f'{HY} 7. hæð FOH innanhúss 300 þ.kr/m² (lúxus, 2026); Næsland: gólfefni jarðhæð 48 þ/m² m/VSK + lýsing 7,2 + bókasafn 2,7 m.kr'),
    ('Veitingastaður og bar innanhúss (án lausra húsgagna)', 'm²', 400, 240000, f'{HY} FOH 300; Næsland bar 4 m.kr m/VSK'),
    ('Eldhús: gólf/veggir/loft í eldhússtaðli, frárennslisrennur, innréttingar (án tækja)', 'm²', 180, 320000, 'Mat ÍF; heilbrigðiskröfur; Næsland eldhús 16,8 m.kr m/VSK með tækjum'),
    ('Fundarrými og vinnuaðstaða', 'm²', 130, 160000, 'Mat ÍF'),
    ('Salerni gesta á jarðhæð og starfsmannaaðstaða', 'heild', 1, 30000000, f'{CO25} forgangsatriði 9: uppfæra stærstan hluta salernisaðstöðu'),
    'Spa, gym og þakgarðsálma innanhúss',
    ('Gym og spa (búningsklefar, gufa, meðferðarherbergi) innanhúss', 'm²', 350, 230000, f'{HY}: infrasaunur 4 stk 10,5 m.kr tilboð; Næsland gym 3,5 m.kr m/VSK'),
    ('Innanhússfrágangur yfirbyggðrar veitingaálmu á þakgarði', 'm²', 150, 220000, 'Mat ÍF; byggingin sjálf í kafla 7'),
    'Tæknirými og BOH',
    ('Kjallari og tæknirými: gólfmálun, brunahólfun, hurðir, lýsingarfrágangur, þvottahús', 'm²', f'={A_BOH}', f'=7000*{F18}+40000', f'{GR}: epoxi 7.000 kr/m² 2018; brunaþéttingar'),
    ('Brunaþéttingar, eldvarnarhurðir og brunahólfun (utan herbergja)', 'heild', 1, 35000000, f'{GR}: brunaþéttingar 0,85 m.kr 2018 fyrir 3.400 m²; hótel mun strangara'),
    ('Merkingar, skilti innanhúss, leiðakerfi', 'heild', 1, 8000000, 'Næsland skilti 3,2 m.kr m/VSK'),
], f'Viðmið: {HY} kafli 5 = 254 þ.kr/m² (öll bygging, lúxus, með 633 m.kr aukaverkum), 7. hæð FOH 300 þ.kr/m². {GR}: economy-hótel 2018.')

chapter('7 Utanhúss', '7. Frágangur utanhúss, gluggar, þak og þakgarður', [
    'Gluggar og gluggaveggir – COWI',
    ('Nýjar gluggaveggjaeiningar (álkerfi, tvöfalt/þrefalt gler, opnanleg fög eða loftristar) í stað eininga með asbesti', 'm²', 1600, 210000, f'{CO24} magnskrá: 1.600 m² gluggaveggjaeiningar; einingarverð mat ÍF fyrir uppsett álkerfi; asbestförgun í kafla 1'),
    ('Stakir gluggar og útihurðir í steyptum göflum 1.–2. og 4.–5. hæðar', 'm²', 300, 150000, f'{CO24} kaflar 3.2–3.4'),
    ('Léttir útveggir 3. hæðar (300 m²): endurnýjun eða þétting', 'm²', 300, 120000, f'{CO24}: 300 m² í þokkalegu ástandi; „gæti verið skynsamlegt að endurnýja" ef gengið er langt í endurnýjun'),
    ('Inngangar, anddyri og hringhurð Kolaports', 'heild', 1, 25000000, f'{CO25}: hringhurð ónýt; nýir hótelinngangar'),
    'Steyptir fletir, múr og sprungur – COWI magnskrá',
    ('Klæðning steyptra flata (val COWI í stað múrviðgerða + málun á 5 ára fresti)', 'heild', 1, 190000000, f'{CO24} kafli 5: klæðning 190 m.kr; múr+málun 84 m.kr + 200 m.kr viðhald í 20 ár'),
    ('Inndælingar í plötuskil', 'lm', 200, 25000, f'{CO24} magnskrá 200 lm'),
    ('Láréttir fletir ofan á handriðum 3. hæðar: brjóta lausan múr, múra með vatnshalla/áfella', 'lm', 270, 40000, f'{CO24} magnskrá 270 lm'),
    ('Vatnsbretti undir gluggum', 'lm', 100, 30000, f'{CO24} magnskrá 100 lm'),
    ('Minniháttar múrviðgerðir', 'm²', 150, 35000, f'{CO24} magnskrá 150 m²'),
    ('Stærri múr- og steypuviðgerðir (þ.m.t. brjóta inn að járnagrind við þakbrún norðurhliðar)', 'm²', 40, 120000, f'{CO24} magnskrá 40 m²'),
    ('Gabbróflísaveggur við mósaíkverk: taka niður, endurfesta/endurgera', 'm²', 100, 250000, f'{CO24} kafli 3.1.2: mikið los, flísar læstar með járnpinnum; endurnýta má flísar. Kvöð um varðveislu mósaíks'),
    ('Þensluskil yfirfarin og þétt', 'heild', 1, 6000000, f'{CO24} kafli 3.1.2'),
    'Þök',
    ('Þak 5. hæðar: nýr dúkur, einangrun, niðurföll', 'm²', 1560, 14000, f'{CO25}: þakdúkur virðist í nokkuð góðu ástandi; endurnýjun þó eðlileg í 40 ára rekstri'),
    ('Þakflötur 3. hæðar: nýr dúkur, ílögn með halla, niðurföll (undir þakgarði)', 'm²', 2900, 26000, f'{HY}: þakdúkur undir yfirborðsfrágang 15.000 kr/m²; hér með ílögn'),
    'Þakgarður 3. hæðar',
    ('Garður: pallar, gróðurker, jarðvegur, lýsing, handrið (án sundlaugar og padelvallar)', 'm²', 1500, 48000, 'Næsland: garður/landslag 30 m.kr m/VSK + padel 50 + sundlaug 40 + sauna 20 (m/VSK); sundlaug/padel EKKI inni hér, sjá valkostalínu'),
    ('Yfirbyggð veitingaálma á þakgarði (glervirki, gólf, hiti) – bygging', 'm²', 150, 520000, 'Næsland: 60 + 12 + 8 = 80 m.kr m/VSK; hér 78 án VSK'),
    ('Sauna og heitir pottar á þakgarði', 'heild', 1, 22000000, 'Næsland: sauna 20 m.kr m/VSK'),
    ('VALKOSTUR (0 = ekki inni): sundlaug og padelvöllur á þaki með burðarstyrkingu', 'heild', 0, 120000000, 'Næsland: 90 m.kr m/VSK án burðarstyrkingar; COWI minnisblað: þyngd ílagnar/aukið álag þarf ráðstafanir. Setja magn 1 til að taka inn'),
    'Kolaportið – aðeins COWI-viðgerðir (má fella á borgina/rekstraraðila)',
    ('Kolaportsþak yfir vesturhluta: leki við lofttúður/stokk, einangrun inntaksloftrörs, loftaklæðning', 'heild', 1, 35000000, f'{CO25} kafli 2.7.5 og forgangsatriði 3'),
    ('Gaflveggir brúarbyggingar við Kolaportið: sprungur og lekar', 'heild', 1, 25000000, f'{CO24} kaflar 3.2.2 og 3.4.2: „þarfnast tafarlausra viðgerða"'),
], f'Viðmið: {HY} kafli 7 = 118 þ.kr/m² (ný klæðning og gluggar). COWI verðleggur aðeins múr 84 / klæðningu 190 / 20 ára viðhald 200.')

chapter('8 Lóð', '8. Lóð, aðkoma og torg', [
    ('Aðkoma hótels frá Tryggvagötu/Pósthússtræti: hellulögn, lýsing, skilti, hjólastæði innan lóðar', 'heild', 1, 30000000, 'Umsögn skipulagsfulltrúa 2024: hjólastæði skulu leyst innan lóðar; torgið er borgarinnar'),
    ('Aðkoma þjónustu og sorps frá Naustinni/Geirsgötu', 'heild', 1, 12000000, 'Mat ÍF'),
], f'Viðmið: {HY} 15 þ.kr/m² (stór lóð); Hamranes 15,7.')

chapter('9 Búnaður', '9. Laus búnaður (FF&E) og rekstrarbúnaður', [
    ('Herbergi: rúm, dýnur, húsgögn, textíll, gardínur, sjónvarp, minibar, öryggisskápur, lampar', 'herb', f'={KEYS}', 3600000, f'{NL}: standard 2,44 m.kr m/VSK = 1,97 án VSK, superior 3,25; Hamranes laus búnaður 209,7 m.kr / 99 rými = 2,1 m.kr á hjúkrunarrými; upper upscale hótel 3,5–4,5'),
    ('Lobby, setustofa, bókasafn/búð: laus húsgögn, listaverk, gólfteppi', 'heild', 1, 60000000, f'{NL}: lobby 12,2 + listaverk 45 m.kr m/VSK'),
    ('Veitingastaður og bar: húsgögn, borðbúnaður, glös, línar, kælar', 'heild', 1, 45000000, f'{NL}: matsalur 23,3 m.kr m/VSK fyrir 60 sæti; hér ~120 sæti með þakgarði'),
    ('Eldhústæki og eldhúsbúnaður (stór- og smátæki)', 'heild', 1, 60000000, f'{HY}: eldhústæki BOH 7. hæð 10 m.kr + uppþvottavélar 1,2 m.kr stk; aðaleldhús hótels 45–70'),
    ('Fundarrými: húsgögn, skjáir, hljóðkerfi', 'heild', 1, 15000000, f'{NL}: 12,1 m.kr m/VSK'),
    ('Gym og spa tæki', 'heild', 1, 25000000, f'{NL}: 4,1 m.kr m/VSK er of lágt fyrir upper upscale'),
    ('Þvottahús, ræsting, húsvarðarbúnaður', 'heild', 1, 18000000, f'{NL}: þvottavél 5,1 + þurrkari 2,1 + ræstibúnaður 1,65 m.kr m/VSK'),
    ('Hótelkerfi (PMS), símkerfi, tölvur, prentarar, sjónvarpskerfi', 'heild', 1, 30000000, f'{NL}: símkerfi 1,6 + símtæki 2,85 + tölvur 2,3 m.kr m/VSK'),
    ('Rekstrarbúnaður við opnun (OS&E: sængur, handklæði, sloppar, vagnar, smávara)', 'herb', f'={KEYS}', 350000, f'{NL}: sængur/koddar 150 + handklæði 45 + sloppar 45 þ.kr m/VSK á herbergi'),
], 'Í leigulíkani greiðir rekstraraðili oft FF&E og OS&E sjálfur (Hamranes: laus búnaður ekki hluti tilboðs). Hér inni til að sýna fullbúið hótel; má taka út ef leigutaki leggur til.')

# ============================================================ HÖNNUN OG STJÓRNUN
wsB = wb.create_sheet('B Hönnun'); setw(wsB, [6, 58, 8, 12, 14, 16, 78])
wsB['A1'] = 'B. Hönnun, ráðgjöf, byggingarstjórn, eftirlit, leyfi og umsýsla'; wsB['A1'].font = H1
wsB['A2'] = 'Hlutföll af framkvæmdakostnaði (kaflar 0–8 + ófyrirséð) eða heildartölur. Bláar tölur eru forsendur.'; wsB['A2'].font = SRC
for j, h in enumerate(['Nr', 'Liður', 'Ein.', 'Hlutfall / magn', 'Grunnur kr', 'Upphæð kr', 'Heimild / rök'], 1):
    c = wsB.cell(4, j, h); c.font = BOLD; c.fill = GREY
FRAMKV_REF = "Samantekt!$C$22"   # framkvæmd + ófyrirséð (röð 22)
blines = [
    ('Arkitekt (aðaluppdrættir, deiliskipulagsbreyting, innanhúss, þakgarður)', '%', 0.040, f'{HY}: arkitekt 288 m.kr / 7.540 = 3,8%; {GR}: 16 m.kr 2018'),
    ('Verkfræði: burðarþol, lagnir, loftræsing, rafmagn, bruni, hljóð', '%', 0.035, f'{HY}: verkfræði 373 m.kr = 4,9% (nýbygging); hér lægra'),
    ('Aðrir ráðgjafar: asbest/mengun, innivist, umhverfismat/vottun, Minjastofnun, umferð/skuggavarp fyrir DSK', '%', 0.012, f'{HY}: aðrir ráðgjafar 138 m.kr = 1,8%'),
    ('Byggingarstjórn og eftirlit', '%', 0.035, f'{HY}: 311 m.kr / 48 mán = 4,1%; {HAM}: umsýsla 3,3 þ.kr/m²'),
    ('Umsýsla, þinglýsingar, bókhald, tryggingar á byggingartíma', '%', 0.005, f'{HY}: 20 m.kr'),
]
r = 5; n = 0
for name, unit, val, src in blines:
    n += 1; wsB.cell(r, 1, f'B.{n}').font = SRC; wsB.cell(r, 2, name); wsB.cell(r, 3, unit)
    c = wsB.cell(r, 4, val); c.font = BLUE; c.number_format = PCT
    c = wsB.cell(r, 5, f'={FRAMKV_REF}'); c.font = GRN; c.number_format = NUM
    c = wsB.cell(r, 6, f'=D{r}*E{r}'); c.number_format = NUM
    wsB.cell(r, 7, src).font = SRC; r += 1
for name, unit, qty, rate, src in [
    ('Byggingarleyfisgjald, úttektir, skráning', 'heild', 1, 6000000, f'{HY}: byggingarleyfisgjald 3,2 m.kr + skoðanir'),
    ('Skipulagsgjald 0,3% af brunabótamati nýrra/endurbættra hluta', 'heild', 1, 12000000, 'Lög um skipulagsgjald; mat'),
    ('Módelherbergi (mock-up) og sýnishorn', 'heild', 1, 25000000, f'{HY}: módelherbergi 40 m² 49 m.kr; Hamranes mock-up'),
    ('Þróunar- og verkefnastjórnunarþóknun ÍF', '%', 0.03, None, 'Þ113-líkan 6% (nýbygging); Hamranes; hér 3% á framkvæmd + ófyrirséð'),
]:
    n += 1; wsB.cell(r, 1, f'B.{n}').font = SRC; wsB.cell(r, 2, name); wsB.cell(r, 3, unit)
    if rate is None:
        c = wsB.cell(r, 4, qty); c.font = BLUE; c.number_format = PCT
        c = wsB.cell(r, 5, f'={FRAMKV_REF}'); c.font = GRN; c.number_format = NUM
        c = wsB.cell(r, 6, f'=D{r}*E{r}'); c.number_format = NUM
    else:
        c = wsB.cell(r, 4, qty); c.font = BLUE; c.number_format = NUM
        c = wsB.cell(r, 5, rate); c.font = BLUE; c.number_format = NUM
        c = wsB.cell(r, 6, f'=D{r}*E{r}'); c.number_format = NUM
    wsB.cell(r, 7, src).font = SRC; r += 1
wsB.cell(r, 2, 'SAMTALS HÖNNUN, STJÓRNUN OG UMSÝSLA').font = BOLD
t = wsB.cell(r, 6, f'=SUM(F5:F{r-1})'); t.font = BOLD; t.number_format = NUM; t.fill = LIGHT
B_TOTAL = f"'B Hönnun'!$F${r}"
wsB.cell(r + 1, 2, 'kr/m² hótelhluta'); c = wsB.cell(r + 1, 6, f'=F{r}/{A_HOTEL}'); c.number_format = NUM
wsB.cell(r + 3, 2, f'Viðmið: {HY} hönnun+stjórnun+umsýsla 1.178 m.kr = 15,6% af framkvæmd (118 þ.kr/m²). Hamranes: hönnun ~5% + umsýsla.').font = Font(name=F, size=9, italic=True)

# ============================================================ SAMANTEKT
ws = wb.create_sheet('Samantekt', 0); setw(ws, [52, 12, 16, 16, 14, 14, 14, 16, 16, 60])
ws['A1'] = 'Tollhúsið – Næsland: kostnaðaráætlun fullbúins hótels, v1.0 (23.9.2026)'; ws['A1'].font = H1
ws['A2'] = 'Innanhúss-útgáfa með heimildum. Kr á byggðan m² hótelhluta (7.741 m² með tæknirýmum, Kolaportssalur utan). Án VSK nema í dálki D. Verðlag september 2026.'; ws['A2'].font = SRC
ws['A3'] = 'Helstu stærðir'; ws['A3'].font = H2
ws['A4'] = 'Hótelhluti byggður, m²'; c = ws['B4']; c.value = f'={A_HOTEL}'; c.font = GRN; c.number_format = NUM1
ws['A5'] = 'Herbergi'; c = ws['B5']; c.value = f'={KEYS}'; c.font = GRN; c.number_format = '0'
ws['A6'] = 'Brúttó m² á herbergi'; c = ws['B6']; c.value = '=B4/B5'; c.number_format = NUM1
ws['A7'] = 'Kolaportssalur utan áætlunar, m²'; c = ws['B7']; c.value = f'={A_KOLA}'; c.font = GRN; c.number_format = NUM1
ws['A8'] = 'Ófyrirséð á framkvæmd'; c = ws['B8']; c.value = 0.12; c.font = BLUE; c.fill = YEL; c.number_format = PCT
ws['C8'] = 'ÍF-venja: samningar 5%, tilboð 12%, áætlað 15%. Allt hér er áætlað; 12% valið, 15% varfærið.'; ws['C8'].font = SRC
ws['A9'] = 'VSK-hlutfall til upplýsingar'; c = ws['B9']; c.value = 0.24; c.font = BLUE; c.number_format = PCT
ws['C9'] = 'Endurheimtist með frjálsri skráningu ef leigt með VSK; hönnun og hluti búnaðar ber VSK líka.'; ws['C9'].font = SRC

hdr = ['Kafli', 'kr/m² (þ.kr)', 'Alls án VSK', 'Alls m/VSK', 'kr/herb (m.kr)', 'Næsland án VSK', 'Hyatt L176 EAC þ.kr/m² (07/2026)', 'Hamranes þ.kr/m² (uppr. 09/2026)', '', 'Athugasemd']
for j, h in enumerate(hdr, 1):
    c = ws.cell(11, j, h); c.font = BOLD; c.fill = GREY
CHAPS = [
    ('0. Aðstaða, umsjón og rekstur vinnusvæðis', CH['0 Aðstaða'], 0, 50.1, 29.8, 28, 'Næsland: 0'),
    ('1. Niðurrif, afmengun og förgun', CH['1 Niðurrif'], 34.0/1.24, 29.6, 11.3, 9, 'Næsland: niðurrif 34 m/VSK, engin afmengun'),
    ('2. Burðarvirki, op, stigar, frárennsli í grunni', CH['2 Burðarvirki'], 0, 90.1, 68.7, None, 'Næsland: frárennsli í grunni óverðlagt; Hyatt byggði nýjar hæðir'),
    ('3. Pípulagnir, hiti, loftræsing, vatnsúði', CH['3 Lagnir'], (3.5+20+3.5+32+4.3+20.8+16+24+1+38.4+10.4+4)/1.24, 78.2, 24.8, None, 'Næsland: loftræsting 4 m.kr, útsog 10,4'),
    ('4. Raflagnir, lýsing, öryggiskerfi, lyftur', CH['4 Raflagnir'], (168+48)/1.24, 112.7, 49.5, None, 'Næsland: raflagnir 168 + lyftur 48 m/VSK'),
    ('5. Frágangur innanhúss', CH['5 Innanhúss'], (12.5+5+12.5+137.5+60+41.25+18+1.5+5.6+63+14.4+9.6+1.35+3+57+11.4+20+40+1.5+9.6+16.8+37.9+13.6+4+2.7+7.2+2.7+3.5+3.2)/1.24, 253.9, 38.0, None, 'Næsland: veggir, gólf, loft, hurðir, bað, fastur búnaður, jarðhæð'),
    ('7. Frágangur utanhúss, gluggar, þak, þakgarður', CH['7 Utanhúss'], (8+20+60+12+8+40+50+30)/1.24, 118.5, 40.8, None, 'Næsland: gluggar 8 m.kr, þakgarður 220 m/VSK'),
    ('8. Lóð og aðkoma', CH['8 Lóð'], 0, 14.9, 15.7, None, ''),
]
r = 12; first = r
for name, ref, nl, hy, ham, gr, note in CHAPS:
    ws.cell(r, 1, name)
    c = ws.cell(r, 3, f'={ref}'); c.font = GRN; c.number_format = NUM
    c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1
    c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM
    c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1
    c = ws.cell(r, 6, nl * 1e6 if nl else 0); c.font = BLUE; c.number_format = NUM
    c = ws.cell(r, 7, hy); c.font = BLUE; c.number_format = NUM1
    c = ws.cell(r, 8, f'={ham}*{F25}'); c.font = BLUE; c.number_format = NUM1
    ws.cell(r, 10, note).font = SRC; r += 1
last = r - 1
ws.cell(r, 1, 'FRAMKVÆMD, kaflar 0–8').font = BOLD
for j, f in ((2, f'=C{r}/$B$4/1000'), (3, f'=SUM(C{first}:C{last})'), (4, f'=C{r}*(1+$B$9)'), (5, f'=C{r}/$B$5/1000000'), (6, f'=SUM(F{first}:F{last})'), (7, f'=SUM(G{first}:G{last})'), (8, f'=SUM(H{first}:H{last})')):
    c = ws.cell(r, j, f); c.font = BOLD; c.number_format = NUM1 if j in (2, 5, 7, 8) else NUM
FR = r; r += 1
ws.cell(r, 1, 'Ófyrirséð'); c = ws.cell(r, 3, f'=C{FR}*$B$8'); c.number_format = NUM
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1
ws.cell(r, 6, 0).font = BLUE; ws.cell(r, 7, 11.9).font = BLUE; ws.cell(r, 10, 'Næsland: 7% skráð en ekki reiknað. Hyatt: 12% á seint stigi með aukaverkum inni.').font = SRC
UNC = r; r += 1
ws.cell(r, 1, 'Framkvæmd með ófyrirséðu').font = BOLD
c = ws.cell(r, 3, f'=C{FR}+C{UNC}'); c.font = BOLD; c.number_format = NUM
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1
assert r == 22, r   # FRAMKV_REF = Samantekt!C22
FWU = r; r += 1
ws.cell(r, 1, 'B. Hönnun, stjórnun, eftirlit, leyfi, þóknun ÍF'); c = ws.cell(r, 3, f'={B_TOTAL}'); c.font = GRN; c.number_format = NUM
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1
ws.cell(r, 6, (48+1.2)/1.24*1e6).font = BLUE; ws.cell(r, 6).number_format = NUM; ws.cell(r, 7, 118.75).font = BLUE; ws.cell(r, 8, f'=30*{F25}').font = BLUE; ws.cell(r, 10, 'Næsland: hönnun 48 + leyfi 1,2 m/VSK = 2,8%. Hyatt 15,6%.').font = SRC
BR = r; r += 1
ws.cell(r, 1, 'HÚS FULLBÚIÐ án lauss búnaðar').font = BOLD
c = ws.cell(r, 3, f'=C{FWU}+C{BR}'); c.font = BOLD; c.number_format = NUM; c.fill = LIGHT
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c.font = BOLD; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1; c.font = BOLD
ws.cell(r, 7, 894.3).font = BLUE; ws.cell(r, 8, f'=780*{F25}').font = BLUE; ws.cell(r, 10, 'Hyatt EAC alls 894 þ.kr/m² = 52,5 m.kr/herb (59 m²/herb), verðlag júlí 2026. Hamranes ~780 þ.kr/m² fullbúið á verðlagi maí 2025, uppreiknað í dálki H.').font = SRC
HUS = r; r += 1
ws.cell(r, 1, '9. Laus búnaður (FF&E) og rekstrarbúnaður'); c = ws.cell(r, 3, f"={CH['9 Búnaður']}"); c.font = GRN; c.number_format = NUM
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1
ws.cell(r, 6, (271.3+62.9+74.2)/1.24*1e6).font = BLUE; ws.cell(r, 6).number_format = NUM; ws.cell(r, 10, 'Næsland: innbú herbergja 271 + sameiginleg 63 + annað 74 m/VSK. Oft á rekstraraðila í leigulíkani.').font = SRC
FFE = r; r += 1
ws.cell(r, 1, 'HÓTEL FULLBÚIÐ MEÐ BÚNAÐI').font = BOLD
c = ws.cell(r, 3, f'=C{HUS}+C{FFE}'); c.font = BOLD; c.number_format = NUM; c.fill = YEL
c = ws.cell(r, 2, f'=C{r}/$B$4/1000'); c.number_format = NUM1; c.font = BOLD; c = ws.cell(r, 4, f'=C{r}*(1+$B$9)'); c.number_format = NUM; c.font = BOLD; c = ws.cell(r, 5, f'=C{r}/$B$5/1000000'); c.number_format = NUM1; c.font = BOLD; c.fill = YEL
c = ws.cell(r, 6, f'=SUM(F{first}:F{FFE})'); c.number_format = NUM; c.font = BOLD
ws.cell(r, 10, 'Næsland alls: 1.697 m.kr m/VSK = 1.368 án VSK = 13,7 m.kr/herb við 100 herb.').font = SRC
TOT = r; r += 2
ws.cell(r, 1, 'Næmni: kostnaður á herbergi eftir herbergjafjölda (sama hús)').font = H2; r += 1
for j, h in enumerate(['Herbergi', 'Hús fullbúið m.kr/herb', 'Með búnaði m.kr/herb'], 1):
    c = ws.cell(r, j, h); c.font = BOLD; c.fill = GREY
r += 1
for k in (89, 106, 123, 130):
    c = ws.cell(r, 1, k); c.number_format = '0'
    # herbergjatengdir liðir breytast með fjölda: nota nálgun = hús − herbergjalínur + herbergjalínur × k/keys (einfaldað: kr/herb við k = C{HUS}/k þar sem herbergislínur eru ~35% af framkvæmd)
    c = ws.cell(r, 2, f'=(C${HUS}*(1-0.35)+C${HUS}*0.35*{k}/$B$5)/{k}/1000000'); c.number_format = NUM1
    c = ws.cell(r, 3, f'=(C${HUS}*(1-0.35)+C${HUS}*0.35*{k}/$B$5+C${FFE}*{k}/$B$5)/{k}/1000000'); c.number_format = NUM1
    r += 1
ws.cell(r, 1, 'Nálgun: um 35% framkvæmdar er bundið herbergjafjölda (bað, veggir, hurðir, lagnir og raflagnir í herbergjum), afgangurinn er húsið sjálft. Nákvæm tala fæst með því að breyta herbergjafjölda á Stærðir-flipa.').font = SRC
r += 2
ws.cell(r, 1, 'Lestur').font = H2; r += 1
for t in [
    'Kolaportssalurinn (2.410 m²) er utan áætlunar að öllu leyti nema tveimur COWI-viðgerðarlínum í kafla 7 (þak, gaflar, hringhurð), sem eru sérmerktar og má strika út.',
    'Kr/m² eru reiknuð á hótelhlutann einan, 7.741 m² (6.880 m² gestarými + 861 m² tæknirými). Það er sá grunnur sem ÍF notar venjulega.',
    'Herbergjatengdir liðir eru með magn = herbergjafjöldi og einingarverð á herbergi, svo talan á herbergi sést beint. Húsbundnir liðir (gluggar, klæðning, frárennsli, kerfi, aðstaða) dreifast á herbergin og það er þar sem óhagstæð nýting (63 m² brúttó á herbergi við 123 herbergi) kemur fram.',
    'Heimildir eru við hverja línu í kaflablöðunum: COWI 2024/2025 fyrir magntölur og forgangsatriði, Hyatt L176 EAC 07/2026 fyrir raunverð á fullbúinni hótelhæð, Grensásvegur 2018 fyrir okkar eigin einingarverð (verðbætt ×1,45), Hamranes fyrir nýbyggingarviðmið, Næsland-skjalið til samanburðar. Þessar tilvísanir eru hreinsaðar út áður en nokkuð fer út úr húsi.',
]:
    ws.cell(r, 1, t).alignment = Alignment(wrap_text=True); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10); ws.row_dimensions[r].height = 30; r += 1
ws.freeze_panes = 'A12'

# ============================================================ HEIMILDIR
wh = wb.create_sheet('Heimildir'); setw(wh, [130])
wh['A1'] = 'Heimildir (innanhúss-útgáfa – hreinsa út fyrir dreifingu)'; wh['A1'].font = H1
for i, s in enumerate([
    'COWI: Ástandsskýrsla Tryggvagata 19, 3210316-000-CRP-0003, 16.10.2024 – magnskrá bls. 7 (gluggaveggir 1.600 m², asbestplötur 600 m², vatnsbretti 100 lm, handrið 270 lm, inndælingar 200 lm, múr 150 m² + 40 m², gabbró 100 m², léttir veggir 300 m²); kostnaðarmat múr 84 / klæðning 190 / viðhald 200 m.kr; viðauki A burðarþol, B asbest; minnisblað 6.9.2024 um ofanábyggingu.',
    'COWI: Innivistarskoðun Tryggvagata 19, A290951-001-HMO-20.08.25, 20.8.2025 – sýnataka (34 efnissýni, 14 DNA), forgangslisti 1–14, viðauki C lagnir í grunni.',
    'Hyatt Centric Laugavegi 176 – EAC júlí 2026 (isfast Drive: Hyatt - Reitir/Framkvæmd/Kostnaðaráætlun/Hyatt - EAC.xlsx). 9.922,9 m², 169 herb. Kaflar án VSK þ.kr/m²: 0 50,1 · 1 29,6 · 2 90,1 · 3 78,2 · 4 112,7 · 5 253,9 · 7 118,5 · 8 14,9 · óvissa 11,9 = 760; hönnun/stjórnun 118,8; alls 894,3. 7. hæð FOH áætlun: lagnir 95, rafmagn 119, innanhúss 300 þ.kr/m². TRÚNAÐARMÁL Reita/ÍF.',
    'ÍF: D - Kostnaðaráætlun - Samningur 78 herbergja hótel, Grensásvegur 16A, mars 2018 (isfast Drive: Grensásvegur 16a/Kaupsamningur final og fylgigögn). Einingarverð 2018: gifsveggir 11.000 / hljóð 14.200 kr/m², kerfisloft 9.100, gifsloft 7.100, parket 9.500, teppi 11.000, flísar 14.000, málning 2.800, rekstur vinnustaðar 3,0 m.kr/mán, arkitekt 16 m.kr. Verðbætt ×1,45 (BVT 2018→2026).',
    'Hamranes EAC 2025 (Hjúkrunarheimilið Hamranesi ehf.): 5.591,9 m.kr án VSK / 6.730–6.899 m² ≈ 780 þ.kr/m² fullbúið með einingum; kaflar 0 29,8 · 1 11,3 · 3 24,8 · 4 49,5 · 5 38,0 · 7 40,8 · 8 15,7 þ.kr/m²; laus búnaður 209,7 m.kr.',
    'Þ113 Hamrar model.py v1.5: Akureyrar-einingarverð 2026 (plötumót 14.500, steypa 55.000/m³, járn 600/kg, lagnaskurðir 8–11 þ.kr/lm), utanáliggjandi stálstigar 22 m.kr/2 stk × 3 hæðir.',
    'Næsland: Vinnuskjal_Naesland_Tollhusid_Stofnkostnadur_EXC.xlsx (22.9.2026) – 1.697 m.kr m/VSK, 100 herb.; línur notaðar til samanburðar í dálki F á Samantekt.',
    'Aðaluppdrættir BN045618 (2013, 1:200) og BN048223 (2014), skjalasafn Reykjavíkur – flatarmál hæða og herbergjatalning. Fasteignaskrá F2000241 – rekstrareiningar.',
], 3):
    wh.cell(i, 1, f'{i-2}. {s}').alignment = Alignment(wrap_text=True)

for w in wb.worksheets:
    for row_ in w.iter_rows():
        for c in row_:
            if c.value is not None and (c.font is None or c.font.name != F):
                c.font = Font(name=F, size=10)
wb.save(OUT); print('saved', OUT)

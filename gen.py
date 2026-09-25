#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tollhúsið – Næsland: mælaborð með glærum. python3 model3.py && python3 gen.py && python3 encrypt.py"""
import base64, json, os, html as H
from openpyxl import load_workbook
_KX = '/Users/villithor/Library/CloudStorage/GoogleDrive-gunnar@lgt.is/My Drive/Claude Projects/Tollhúsið Næsland/Tollhúsið - Kostnaðaráætlun fullbúið hótel v1.0 - 23.09.2026.xlsx'
_ks = load_workbook(_KX, data_only=True)['Samantekt']
KOST = []   # (nafn, þ.kr/m², m.kr, m.kr/herb, Næsland m.kr)
for _r in range(12, 27):
    _a = _ks.cell(_r, 1).value
    if _a: KOST.append((_a, _ks.cell(_r, 2).value, _ks.cell(_r, 3).value, _ks.cell(_r, 5).value, _ks.cell(_r, 6).value))
A_HUS = [x for x in KOST if x[0].startswith('HÚS FULLBÚIÐ')][0][2] / 1e6
A_FFE = [x for x in KOST if x[0].startswith('9.')][0][2] / 1e6
HERE = os.path.dirname(os.path.abspath(__file__))
R = json.load(open(os.path.join(HERE, 'results3.json')))
def b64(fn):
    with open(os.path.join(HERE, 'assets', fn), 'rb') as f: return 'data:image/jpeg;base64,' + base64.b64encode(f.read()).decode()
def n0(x): return f"{x:,.0f}".replace(",", ".")
def n1(x): return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
def pct(x, d=0): return f"{x*100:.{d}f}".replace(".", ",") + "%"
def neg(x): return f'<span class="neg">{n0(x)}</span>' if x < 0 else n0(x)

PL = R['pl']; KEYS = R['keys']; CAP = R['capex']; BENCH = R['bench']; SCEN = R['scen']
tiers = R['tiers']

# ---------- tables
def tbl(head, rows, cls=''):
    h = ''.join(f'<th>{c}</th>' for c in head)
    b = ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>' for r in rows)
    return f'<table class="t {cls}"><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table>'

keys_rows = [[k, v['h5'], v['h4'], v['h3'], v['h2'], f"<b>{v['total']}</b>"] for k, v in KEYS.items()]
capex_rows = []
for t in 'ABC':
    for k, v in KEYS.items():
        X = CAP[t][k]
        capex_rows.append([t, k, n0(v['total']), n0(X['conv']), n0(X['boh'] + X['lumps']), n0(X['unc']), n0(X['soft'] + X['fee']), n0(X['ffe']), f"<b>{n0(X['hotel_only'])}</b>", f"<b>{n0(X['per_key'])}</b>"])
pl_rows = []
for t in 'ABC':
    P = PL[t]
    pl_rows.append([t, n0(P['adr']), pct(P['occ']), n0(P['revpar']), n0(P['rooms']), n0(P['fb'] + P['other']), f"<b>{n0(P['rev'])}</b>", n1(P['rev_key'])])
usali_rows = []
for t in 'ABC':
    P = PL[t]
    usali_rows.append([t, n0(P['rev']), n0(P['dept']), n0(P['gop']) + f" ({pct(P['gop_pct'])})", n0(P['ffe']),
                       f"<b>{n0(P['rent_usali'])}</b> ({pct(P['c_usali'], 1)})", f"<b>{n0(P['rent_oddsson'])}</b> ({pct(P['c_oddsson'], 1)})", f"{n0(P['rent_jhb'])} ({pct(P['c_jhb'], 1)})",
                       pct(P['op_after_usali'], 1) + ' / ' + pct(P['op_after_oddsson'], 1), n1(P['noi_key_usali']) + ' / ' + n1(P['noi_key_oddsson'])])
bench_rows = [[b['name'], b['value'], f'<span class="src">{b["src"]}</span>'] for b in BENCH]

def scen_rows(filter_fn):
    rows = []
    for s in SCEN:
        if not filter_fn(s): continue
        rows.append([s['tier'], s['keys'], pct(s['cf']), s['cname'], pct(s['C'], 1), 'Kolaport 45' if s['ground'] == 'kolaport' else 'Matarhöll 159', pct(s['D']) + ' / ' + pct(s['yld'], 2),
                     n0(s['A']), n0(s['per_key']), n0(s['leiga']), n0(s['noi']), n0(s['V']), f"<b>{neg(s['price'])}</b>"])
    return rows
sel = scen_rows(lambda s: s['keys'] == 123 and s['cname'] in ('USALI 65%', 'ODDSSON') and s['tier'] in 'BC')

head_scen = ['Fl.', 'Herb.', 'Kostn.', 'C-viðmið', 'C', 'Jarðhæð', 'D / krafa', 'A', 'A/herb', 'Leiga', 'NOI', 'Verðmæti', 'Rétt verð']

HTML = f'''<!doctype html><html lang="is"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow">
<title>Tollhúsið – Næsland · greining fyrir tilboð</title>
<style>
:root{{--navy:#12263A;--navy2:#1D3A57;--gold:#C9A227;--ink:#1b2733;--muted:#5A6B7A;--line:#DCE3EA;--bg:#F4F6F8;--red:#B23A3A;--green:#2a7d4f}}
*{{box-sizing:border-box}}html{{scroll-behavior:smooth}}
body{{margin:0;font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif;color:var(--ink);background:var(--bg)}}
nav{{position:fixed;left:0;top:0;bottom:0;width:214px;background:var(--navy);color:#EAF1F8;padding:22px 14px;overflow:auto;z-index:5}}
nav h1{{font-size:14px;letter-spacing:3px;margin:0 0 4px;color:#fff}}nav .sub{{font-size:11px;color:var(--gold);margin-bottom:18px}}
nav a{{display:block;color:#C9D6E2;text-decoration:none;font-size:12.5px;padding:6px 8px;border-radius:6px;margin:1px 0}}nav a:hover,nav a.on{{background:var(--navy2);color:#fff}}
nav .foot{{position:absolute;bottom:14px;left:14px;right:14px;font-size:10.5px;color:#7F91A3}}nav .foot button{{margin-top:6px;background:none;border:1px solid #3A5573;color:#C9D6E2;border-radius:6px;padding:4px 8px;font-size:11px;cursor:pointer}}
main{{margin-left:214px}}
section{{min-height:100vh;padding:48px 56px 56px;border-bottom:1px solid var(--line);background:#fff}}section:nth-child(even){{background:var(--bg)}}
section h2{{font-size:26px;margin:0 0 4px;color:var(--navy)}}section .lead{{font-size:15px;color:var(--muted);margin:0 0 22px;max-width:980px}}
.k{{font-size:11px;letter-spacing:2px;color:var(--gold);font-weight:700;margin-bottom:6px}}
.grid{{display:grid;gap:18px}}.g2{{grid-template-columns:1fr 1fr}}.g3{{grid-template-columns:repeat(3,1fr)}}.g4{{grid-template-columns:repeat(4,1fr)}}
.card{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:16px 18px}}.card h3{{margin:0 0 6px;font-size:14px;color:var(--navy)}}.card .big{{font-size:26px;font-weight:700;color:var(--navy)}}.card .big small{{font-size:13px;color:var(--muted);font-weight:400}}.card p{{margin:6px 0 0;font-size:13px;color:var(--muted);line-height:1.45}}
.t{{width:100%;border-collapse:collapse;font-size:12.5px;background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden}}.t th{{background:var(--navy);color:#fff;text-align:left;padding:8px 9px;font-weight:600;font-size:11.5px}}.t td{{padding:7px 9px;border-top:1px solid var(--line);vertical-align:top}}.t td:not(:first-child){{white-space:nowrap}}.t.wrap td{{white-space:normal}}.t tr:hover td{{background:#F7F9FB}}
.neg{{color:var(--red)}}.src{{color:var(--muted);font-size:11px}}
.callout{{border-left:4px solid var(--gold);background:#FFF9E8;padding:12px 16px;border-radius:8px;font-size:14px;margin:16px 0;line-height:1.5}}
.warn{{border-left-color:var(--red);background:#FBEDED}}
.hero{{background:var(--navy);color:#fff;min-height:100vh;display:grid;grid-template-columns:55% 45%;padding:0;align-items:stretch}}.hero h2{{color:#fff;font-size:44px;margin:0 0 10px}}.hero .sub{{color:var(--gold);font-size:18px;margin-bottom:26px}}.hero p{{max-width:720px;font-size:15px;line-height:1.6;color:#C9D6E2}}
.hero img{{width:100%;height:100%;min-height:100vh;object-fit:cover;opacity:.85;display:block}}.hero .txt{{padding:70px 56px 60px 64px;align-self:center}}
.plans img{{width:100%;border:1px solid var(--line);border-radius:8px;background:#fff}}.plans figcaption{{font-size:12px;color:var(--muted);margin:6px 0 14px}}
ul.tight{{margin:6px 0 0 18px;padding:0;font-size:13.5px;line-height:1.55}}ul.tight li{{margin-bottom:5px}}
.calc{{display:grid;grid-template-columns:330px 1fr;gap:26px}}.calc label{{display:block;font-size:12.5px;margin:10px 0 2px;color:var(--muted)}}.calc select,.calc input[type=range]{{width:100%}}.calc .v{{float:right;color:var(--navy);font-weight:600}}
.wf td:last-child{{text-align:right;font-variant-numeric:tabular-nums}}.res{{font-size:34px;font-weight:800;color:var(--navy)}}.res.neg{{color:var(--red)}}
@media(max-width:1000px){{nav{{display:none}}main{{margin:0}}section{{padding:28px 18px}}.g2,.g3,.g4,.calc{{grid-template-columns:1fr}}.hero{{grid-template-columns:1fr}}.hero img{{min-height:40vh;height:40vh}}}}
@media print{{nav{{display:none}}main{{margin:0}}section{{page-break-after:always;min-height:auto}}}}
</style></head><body>
<nav><h1>TOLLHÚSIÐ</h1><div class="sub">Næsland · greining fyrir tilboð · trúnaðarmál</div>
<a href="#s0">Forsíða</a><a href="#s1">Í hnotskurn</a><a href="#s2">Eignin og áhætturnar</a><a href="#s3">Teikningar og herbergi</a><a href="#s4">A · Umbreyting</a><a href="#s5">B · Tekjur</a><a href="#s6">C · Leiga</a><a href="#s7">D · Fjármögnun og arðsemi</a><a href="#s8">Rétt verð · reiknivél</a><a href="#s8b">Bakreikningur · verð á nótt</a><a href="#s9">Sviðsmyndir</a><a href="#s10">Samkeppni og tilboðsform</a><a href="#s11">Næstu skref og heimildir</a>
<div class="foot">Drög 2 · 25.9.2026 · líkan v0.4<br>Örvatakkar fletta glærum<br><button onclick="try{{localStorage.removeItem('tollhus_pass')}}catch(e){{}};location.reload()">Læsa síðu</button></div></nav>
<main>
<section id="s0" class="hero"><div class="txt"><div class="k">ÍSLENSKAR FASTEIGNIR · NÆSLAND · 22. SEPTEMBER 2026</div><h2>Tollhúsið, Tryggvagötu 19</h2><div class="sub">Hvað stendur hótelverkefni undir háu kaupverði? Afleiðsla í fjórum skrefum: A framkvæmd · B tekjur · C leiga · D arðsemi</div>
<p>Ríkissjóður selur Tollhúsið, 10.150,8 m² á 4.862 m² leigulóð, í gegnum Fasteignasöluna TORG. Tilboðsfrestur er föstudaginn 9. október 2026. Næsland-hópurinn, ásamt Íslandshótelum, leitar til ÍF sem þróunaraðila. Þessi síða leiðir kaupverðið út sem afleidda stærð, með teikningum hússins, kostnaðaráætlun ÍF, íslenskum leiguviðmiðum og fjármögnunarkjörum sem ÍF þekkir. Engin tala er föst fyrirfram.</p>
<p style="color:#7F91A3;font-size:12.5px">Vinnuskjal GT/ÍF. Gögn frá Næslandi, COWI, TORG, skjalasafni Reykjavíkur og áætlunum ÍF. Ekkert hefur farið út úr húsi.</p></div><img src="{b64('hero.jpg')}" alt=""></section>

<section id="s1"><div class="k">1</div><h2>Í hnotskurn</h2><p class="lead">Fimm staðreyndir sem ráða málinu, og aðferðin sem síðan fylgir.</p>
<div class="grid g4">
<div class="card"><h3>Húsið ber</h3><div class="big">106–123 <small>herbergi</small></div><p>Lesið af grunnmyndum 1:200 (BN045618). Næsland skipaði 89–100 með lobby á 3. hæð. Jarðhæðin, afgreiðsla tollstjóra 730 m² bak við mósaíkið, er lobbyið.</p></div>
<div class="card"><h3>Umbreyting kostar</h3><div class="big">{n1(A_HUS/1000)} <small>ma.kr án VSK og án lauss búnaðar</small></div><p>Kostnaðaráætlun ÍF v1.0, kaflar 0–8 með magni og einingarverðum, rýnd 25.9. Það eru {n0(A_HUS/123)} m.kr á herbergi við 123 herbergi, {n0((A_HUS+A_FFE)/123)} með lausum búnaði. Næsland gerði ráð fyrir 1,4 ma.kr, 14 m.kr á herbergi án VSK.</p></div>
<div class="card"><h3>Leiga sem hótel þolir</h3><div class="big">20–25% <small>af heildartekjum</small></div><p>USALI-uppbygging 20–24%, ODDSSON-samningur ÍF 25% (30% af herbergistekjum). Næsland miðaði við 34% af herbergistekjum án F&B, sem jafngildir 22–28%.</p></div>
<div class="card"><h3>Verð á nótt sem kaupverðið krefst</h3><div class="big">60–80 <small>þ.kr með VSK</small></div><p>Við 2 ma.kr kaupverð, 123 herbergi og 75% nýtingu þarf meðalverð á selda nótt að vera 66 þ.kr með VSK hjá langtímaeiganda og 80 þ.kr hjá þróunaraðila. Sjá Bakreikning.</p></div>
</div>
<div class="callout"><b>Aðferð.</b> Verðið er afleidd stærð, og Bakreikningurinn snýr keðjunni við og sýnir hvaða verð á nótt tiltekið kaupverð krefst. Verðið er afleidd stærð: verðmæti fullbúins hótels (NOI eiganda / ávöxtunarkrafa) að frádreginni arðsemiskröfu fjárfesta (D), umbreytingu (A), fjármagnskostnaði og biðtíma, gefur það sem hægt er að greiða fyrir húsið eins og það er. Hver liður er reitur í reiknivélinni á glæru 8 og í Excel-líkaninu. Tilboðið verður eitt verð með walk-away fyrirvara um deiliskipulag (hótel + tilgreint viðbótarbyggingarmagn), fjármögnun og lóðarleigusamning.</div>
<div class="grid g3">
<div class="card"><h3>Það sem er sterkt</h3><ul class="tight"><li>Burðarvirki og steypa í góðu lagi, ein hæð til viðbótar möguleg burðarþolslega (COWI 6.9.2024)</li><li>Staðsetning, kennileiti, mósaík Gerðar Helgadóttur, torgið</li><li>Kolaportið og þakflötur 3. hæðar sem sölutæki gagnvart ríki og borg</li><li>Hæðarhæð 3,2–3,45 m og 15,5 m djúp álma hentar tvíhliða hótelgangi</li></ul></div>
<div class="card"><h3>Það sem er veikt</h3><ul class="tight"><li>Hótel ekki heimilt í deiliskipulagi Kvosar; lóðarleigusamningur útrunninn 2017</li><li>Asbest 550–600 m², PCB, mygla á öllum hæðum nema 5., frárennsli í grunni ónýtt, 1.600 m² gluggaveggir</li><li>2. hæð er einhliða (norðurhlið inni í Kolaportssal); 63 m² brúttó á herbergi við 123 herbergi, nýbygging er nær 55–60</li><li>Fasteignagjöld 93 m.kr á ári frá kaupdegi; forkaupsréttur og samþykki hafnarstjórnar</li></ul></div>
<div class="card"><h3>Það sem ræður úrslitum</h3><ul class="tight"><li><b>Kostnaðarprófið:</b> áætlun v1.0 er rýnd, en gluggar, klæðning, frárennsli og afmengun eru mat þar til tilboð liggja fyrir.</li><li><b>Jarðhæðarprófið:</b> Kolaportið greiðir 45 m.kr fyrir 2.410 m²; markaðsleiga væri 145–200.</li><li><b>Eigendaprófið:</b> D 20% á 6,75% eða langtímaeigandi á D 10% og 6,0% munar 1–2 ma.kr.</li><li><b>Rekstraraðilinn:</b> hver skrifar undir 25 ára leigu og með hvaða ábyrgð.</li></ul></div>
</div></section>

<section id="s2"><div class="k">2</div><h2>Eignin og áhætturnar</h2><p class="lead">Söluyfirlit TORG 31.8.2026, fasteigna- og veðbandayfirlit HMS, umsögn skipulagsfulltrúa 4.4.2024, COWI 2024 og 2025.</p>
<div class="grid g2">
<div>{tbl(['Atriði', 'Staðreynd'], [
 ['Stærð', '10.150,8 m² brúttó: kjallari 710, 1. hæð 3.515 (þ.a. Kolaport 2.410), 2. hæð 1.591, 3. hæð 943, 4. hæð 1.620, 5. hæð 1.620, lyftuhús 61. Byggt 1967–71, Gísli Halldórsson.'],
 ['Lóð', '4.862 m² leigulóð Reykjavíkurborgar (deiliskipulag 2020: 4.616 m²). 50 ára samningur frá 1.1.1967, samkomulag 1993, <b>útrunninn 2017</b>; kaupandi semur við borgina.'],
 ['Fasteignamat', '4.725,5 m.kr 2026 (hús 3.738 + lóð 987), 4.914,9 m.kr 2027. Brunabótamat 6.000. Ríkið mat húsið sjálft á „a.m.k. 2 ma.kr“ 2022.'],
 ['Gjöld', 'Fasteignagjöld 85,5 + vatn/fráveita 7,8 = 93 m.kr/ár. Stimpilgjald 1,6% = 76 m.kr. Brunatrygging 5,5.'],
 ['Kvaðir', 'Forkaupsréttur hafnarstjórnar, sala og uppdrættir háð samþykki hafnarstjórnar (411-F-18055); varðveisla mósaíkverks.'],
 ['Skipulag', 'AR2040 M1a, hverfisvernd. Deiliskipulag Kvosar 1988/2015/2020: landnotkun <b>skrifstofur, bílageymsla, vörugeymsla</b>. Hótel krefst DSK-breytingar með samgöngumati og skuggavarpi; skipulagsfulltrúi útilokaði ekki eina hæð til viðbótar (LHÍ-fyrirspurn 2024).'],
 ['Kolaportið', 'Ríkið leigir borginni, borgin framleigir. Rekstrarfélag þrot nóv 2024; borgin auglýsti rekstraraðila júní 2025, lágmarksleiga 3,78 m.kr/mán (45 m.kr/ár). Samningur ríkis og borgar óséður.'],
 ['Sala', 'TORG (Þorlákur Ómar Einarsson / Sigurður Gunnlaugsson), ásett verð „tilboð“, söluheimild í fjárlagafrumvarpi 2027. <b>Tilboðsfrestur 9.10.2026.</b>'],
], 'wrap')}</div>
<div><h3 style="margin:0 0 8px;color:var(--navy)">Rauð flögg COWI 2024 og 2025</h3>{tbl(['Atriði', 'Niðurstaða COWI', 'Vægi'], [
 ['Frárennsli í grunni', 'Pottlagnir mjög ryðgaðar og bólgnar, myndavél komst ekki áfram; forgangsatriði 1, gera meðan húsið er tómt', 'Mjög mikið'],
 ['Gluggaveggir', '1.600 m² einingar með asbestplötum; skipta út öllu, leyfishafar', 'Mjög mikið'],
 ['Asbest, PCB, þungmálmar', '550–600 m² asbest í gluggaeiningum, asbest í gólfflísum; PCB og þungmálmar í gluggamálningu', 'Mjög mikið'],
 ['Mygla og innivist', 'Örveruvöxtur á öllum hæðum nema 5.; skrifstofuhluti ekki nothæfur fyrr en innivist er bætt', 'Mjög mikið'],
 ['Ytra byrði', 'Múr og steypa 84 m.kr, klæðning 190 m.kr (án glugga); gaflveggir brúar lekir', 'Mikið'],
 ['Kolaportið', 'Virkur leki, mygluð loftaklæðning, hringhurð ónýt, salerni bágborin', 'Mikið'],
 ['Burðarvirki', 'Steypa >35 MPa, járnabinding skv. teikningum; ein hæð ofan á 3. og 5. hæð í lagi', 'Jákvætt'],
], 'wrap')}
<div class="callout warn">Ríkið hætti við að setja Listaháskólann í húsið 2024 þegar áætlun fór úr 12,8 í 17 ma.kr fyrir 15.500 m². Sú tala er fyrir aðra og stærri umbreytingu, en hún segir hvað þetta hús kostar þegar farið er í það af alvöru.</div></div>
</div></section>

<section id="s3" class="plans"><div class="k">3</div><h2>Teikningar og herbergjatalning</h2><p class="lead">Aðaluppdrættir 1:200 úr skjalasafni Reykjavíkur (BN045618, 2013). Suðurálman við Tryggvagötu er um 78 m × 15,5 m með tveimur kjörnum; norðurálman á 4. og 5. hæð um 18 m. Hæðarhæð 3,2–3,45 m, Kolaportssalur 7,1 m.</p>
<div class="grid g2">
<figure><img src="{b64('plan5.jpg')}" alt="5. hæð"><figcaption><b>5. hæð, 1.620 m².</b> Tvíhliða gangur: 34 herbergi á 4,5 m breidd (33–36 m²), 39 á 3,9 m (27–30 m²). Sama á 4. hæð.</figcaption></figure>
<figure><img src="{b64('plan3.jpg')}" alt="3. hæð"><figcaption><b>3. hæð, 943 m², inndregin.</b> Norðurhliðin opnast út á þakflötinn (bílastæði og „hraðbraut“): garðherbergi og spa. 24–28 herbergi ef lobby fer á jarðhæð.</figcaption></figure>
<figure><img src="{b64('plan2.jpg')}" alt="2. hæð"><figcaption><b>2. hæð, 1.591 m².</b> Norðurhlið álmunnar liggur inn í Kolaportssalinn: aðeins herbergi sunnan megin, 14–17. Næsland setti þar 11 og gym.</figcaption></figure>
<figure><img src="{b64('plan1.jpg')}" alt="1. hæð"><figcaption><b>1. hæð, 3.515 m².</b> Kolaportssalur 2.410 m². Austurblokkin, afgreiðsla tollstjóra 730 m² við Pósthússtræti og Steinbryggju, er lobby og veitingastaður; vesturblokkin starfsmannarými.</figcaption></figure>
</div>
<div class="grid g2" style="margin-top:8px"><div>{tbl(['Skipan', '5. hæð', '4. hæð', '3. hæð', '2. hæð', 'Alls'], keys_rows)}<p class="src">Talning: 30% í kjarna og ganga af brúttó; herbergi 3,9 m eða 4,5 m breið og 6,5 m djúp. Nákvæm skipan bíður arkitekts, en bilið 106–123 er það sem húsið ber.</p></div>
<figure><img src="{b64('snid.jpg')}" alt="Snið"><figcaption><b>Snið A-A og B-B.</b> Gólfkótar 3,90 / 7,10 / 10,40 / 13,60 / 17,05. Þakplata 3. hæðar og 5. hæðar þola eina hæð til viðbótar (COWI 6.9.2024).</figcaption></figure></div></section>

<section id="s4"><div class="k">4 · A</div><h2>Umbreytingarkostnaður – áætlun ÍF v1.0</h2><p class="lead">Kostnaðaráætlun fullbúins hótels í sniði ÍF, kaflar 0–9 með magni, einingarverðum og heimild við hverja línu. Rýnd af GT og Sveini 25.9.2026. Allt án VSK á verðlagi september 2026. Kolaportssalurinn, 2.410 m², er utan áætlunar.</p>
<div class="grid g2">
<div>{tbl(['Kafli', 'þ.kr/m²', 'm.kr', 'm.kr/herb', 'Næsland m.kr'], [[k, n1(a) if a else '', n0(b/1e6) if b else '', n1(c) if c else '', (n0(d/1e6) if d else '–')] for k, a, b, c, d in KOST])}
<p class="src">Grunnur kr/m²: hótelhluti 7.742 m² (6.857 m² gestarými + 885 m² tæknirými), 123 herbergi, 63 m² brúttó á herbergi. Næsland-dálkur: vinnuskjal Næslands án VSK, 100 herbergi.</p></div>
<div>
<div class="callout"><b>Herbergið sjálft er ekki vandamálið.</b> Veggir, bað, hurðir, gólf, loft, málun og fastar innréttingar eru um 4,8 m.kr á herbergi. Hitt, um 40 m.kr á herbergi, er húsið og kerfin: gluggaveggir og klæðning, lagnir, loftræsing, raflagnir og lyftur, aðstaða, niðurrif og afmengun, frárennsli í grunni, gangar, lobby, veitingar og spa, ófyrirséð og hönnun. Deilt með 123 herbergjum.</div>
<div class="callout"><b>Það sem COWI kallar á</b> er um 1,0 ma.kr af heildinni, eða 8 m.kr á herbergi: gluggaveggir 1.600 m², klæðning steyptra flata, frárennsli í grunni, asbest, PCB og mygla, Kolaportsþak og gaflar. COWI verðleggur aðeins múr, klæðningu og 20 ára viðhald; hitt er mat ÍF þar til tilboð liggja fyrir.</div>
<div class="callout"><b>Það sem þarf að rýna sérstaklega:</b> gluggaveggir 210 þ.kr/m² uppsettir, klæðning 190 m.kr (COWI-tala óbreytt), frárennsli í grunni 600 m² uppbrot og 260 lengdarmetrar, afmengun 85 m.kr, og hvort 12% ófyrirséð dugi þegar allt er áætlað en ekki tilboð (ÍF-venja 15% á áætlaða liði).</div>
<p class="src">Heimildir í áætluninni: COWI 2024 og 2025 fyrir magntölur og forgangsatriði; reynslutölur ÍF af sambærilegri umbreytingu 2026; eldri áætlun ÍF um 78 herbergja hótel (2018) uppreiknuð með byggingarvísitölu 137,0 → 206,9; áætlun ÍF um nýbyggingu (2025) uppreiknuð með byggingarvísitölu 200,4 → 206,9.</p></div>
</div></section>
<section id="s5"><div class="k">5 · B</div><h2>Tekjur hótelsins</h2><p class="lead">Þrír gæðaflokkar. ADR er meðalverð á selda nótt yfir árið án 11% VSK. Tekjur miðaðar við 123 herbergi.</p>
<div class="grid g2">
<div>{tbl(['Fl.', 'ADR án VSK', 'Nýting', 'RevPAR', 'Herbergistekjur', 'F&B og annað', 'Tekjur alls', 'm.kr/herb'], pl_rows)}
{tbl(['Flokkur', 'Lýsing', 'Viðmið'], [
 ['A', tiers['A']['name'], 'Næsland-verðskrá 55/42/32 þ.kr m/VSK → ~36 þ.kr án VSK að meðaltali; €245'],
 ['B', tiers['B']['name'], '€330; Konsulat, Iceland Parliament, Sand'],
 ['C', tiers['C']['name'], '€450; EDITION-flokkur, krefst vörumerkis og þjónustu'],
], 'wrap')}</div>
<div>{tbl(['Markaðsviðmið', 'Gildi', 'Heimild'], [
 ['Íslandshótel, öll keðjan 2023', 'ADR 25,4 þ.kr · nýting 68% · RevPAR 17,2 þ.kr · 8,5 m.kr tekjur á herbergi (2024)', 'Útgefandalýsing maí 2024, ársreikningur 2024'],
 ['Höfuðborgarsvæðið 2025', 'Nýting 74,7%; 58 hótel, 5.555 herbergi', 'Hagstofa Íslands'],
 ['Black Dunes Þorlákshöfn (Flóra, 120 herb.)', 'ADR 29–32 þ.kr · nýting 67–74% (2027–32)', 'Rekstraráætlun 17.3.2025, isfast Drive'],
 ['ODDSSON 2021 (ÍF, 77 herb.)', 'Brúttó ADR €60–80 · nýting 70–85% eftir mánuðum', 'Áætlun 06/2020, isfast Drive'],
 ['Næsland ár 3', 'Herbergistekjur 1.066 án VSK, nýting 74%, engar F&B-tekjur', 'Rekstrar- og söluáætlun draft 1'],
 ['Markaðurinn 2025', 'Nýting féll, RevPAR staðið í stað í krónum í tvö ár; sumar 2026 gott', 'Arion greining 2025, mbl 2.9.2026'],
], 'wrap')}
<div class="callout">Tekjur á herbergi eru 11,6 / 17,7 / 24,7 m.kr eftir flokki. Íslandshótel eru á 8,5 að meðaltali yfir allt landið. Flokkur C tvöfaldar tekjurnar á herbergi en fækkar herbergjum um fimmtung og kallar á vörumerki, F&B-rekstur og þjónustustig sem þarf að staðfesta með rekstraraðila.</div></div>
</div></section>

<section id="s6"><div class="k">6 · C</div><h2>Leiga sem hótelið þolir</h2><p class="lead">Byggt upp eftir USALI-uppgjörsstaðli hótela og borið saman við íslenska samninga. 123 herbergi, m.kr á ári.</p>
{tbl(['Fl.', 'Tekjur', 'Deildarframlegð', 'GOP', 'FF&E 4%', 'Leiga USALI (65% af EBITDAR)', 'Leiga ODDSSON (max 30% herb. / 25% alls)', 'Leiga Næslands (34% herb.)', 'Rekstraraðili eftir leigu (USALI / ODDSSON)', 'Leiga á herbergi, m.kr (USALI / ODDSSON)'], usali_rows)}
<p class="src">Forsendur USALI: herbergjadeild kostar 32% (laun 18%, sölukostnaður og OTA 8%, þvottur og vörur 6%); veitingadeild 75%; annað 50%; ódeilt 19% (stjórnun 7, markaðsmál 4, UT 1,5, viðhald 3,5, orka 3); enginn stjórnunarsamningur (eigin rekstur); FF&E-sjóður 4%. Fastur leigusamningur = 65% af EBITDAR eftir FF&E til leigusala (alþjóðleg venja 60–70%).</p>
<div class="grid g2" style="margin-top:14px">
<div>{tbl(['Viðmið', 'Gildi', 'Heimild'], bench_rows, 'wrap')}</div>
<div><div class="callout"><b>Niðurstaða um C.</b> Fyrir þetta hús er leiga á bilinu <b>20–25% af heildartekjum</b> raunhæf: neðri mörkin úr USALI-uppbyggingu með 65% til leigusala, efri mörkin úr samningi ÍF sjálfs við RR hótel um ODDSSON (30% af herbergistekjum eða 25% af heildartekjum). Flóra gerði ráð fyrir 30% í Þorlákshöfn og skildi rekstraraðilann eftir með 2–4% EBITDA, sem er ekki sjálfbært. Tala Næslands, 34% af herbergistekjum, lendir á 22–28% af heildartekjum eftir flokki og er því ekki fráleit, en hún var ekki rökstudd.</div>
<div class="callout"><b>Hvað Íslandshótel greiða í raun.</b> Nýi samningurinn við Reiti gefur 1,53 m.kr NOI á herbergi á ári fyrir uppgerð fjögurra stjörnu hótel. Okkar B-flokkur þarf 3,9–4,4 og C-flokkur 5,0–6,2 á herbergi. Munurinn er verðflokkurinn: 25 þ.kr ADR á móti 48–65 þ.kr. Það er fyrsta spurningin til Íslandshótela: trúa þau á þessa verðflokka í þessu húsi?</div>
<div class="callout"><b>Form leigusamnings.</b> Blandaður samningur með grunnleigu (t.d. 70% af væntri leigu, vísitölubundinni) og 8–12% af brúttósölu, eins og í Fosshótel Austfjörðum, er bankahæfur og deilir uppsveiflunni. Hreinn veltusamningur eins og ODDSSON er ekki bankahæfur fyrir 65% LTV endurfjármögnun.</div></div>
</div></section>

<section id="s7"><div class="k">7 · D</div><h2>Fjármögnun og arðsemiskrafa</h2><p class="lead">Sömu kjör og ÍF hefur í sambærilegu verkefni í dag. D er sett fram sem álag á heildarkostnað, sem jafngildir 12–15% IRR eiginfjár yfir 4–5 ára þróunartíma.</p>
<div class="grid g3">
<div class="card"><h3>Byggingartími</h3><div class="big">70/30</div><p>Byggingarlán 70% af kaupverði og framkvæmd á 10,2% óverðtryggðum PIK-vöxtum (kjörvextir + álag), lántökugjald 0,2%. Eigið fé 30%.</p></div>
<div class="card"><h3>Tímalína</h3><div class="big">2027 → 2032</div><p>Kaup 2027 eftir fjárlagaheimild, deiliskipulag og hönnun 2027, framkvæmd 2028–29, opnun 2030, leiga 80/92/100%, stöðugur rekstur 2032. Fasteignagjöld 93 m.kr á ári allan tímann.</p></div>
<div class="card"><h3>Endurfjármögnun og sala</h3><div class="big">6,0–6,75%</div><p>Verðtryggt bréf, 65% af verðmæti, 3,9% raunvextir, 25 ár. Krafa á NOI: 6,75% fyrir hótelleigu til þróunaraðila sem selur, 6,0% fyrir langtímaeiganda (Reitir matskrafa 6,4–6,7%, nýkaup 7,8%, ríkisleiga 5,7%).</p></div>
</div>
<div class="grid g2" style="margin-top:18px">
<div class="card"><h3>D = 20%, krafa 6,75%</h3><p>Þróunaraðili sem selur fullbúna eign eftir stöðugan rekstur og þarf 12–15% IRR á eigið fé. Þetta er sjónarhorn ÍF og Næsland-hópsins án langtímaeiganda.</p></div>
<div class="card"><h3>D = 10%, krafa 6,0%</h3><p>Fasteignafélag eða lífeyrissjóður sem skuldbindur sig fyrirfram til að kaupa fullbúna eign og reiknar á eigin kröfu. ÍF fær þróunarþóknun og árangurshlut. Munurinn á sjónarhornunum er 1–2 ma.kr í verði.</p></div>
</div>
<div class="callout" style="margin-top:18px"><b>Keðjan.</b> Rétt verð = [ (NOI / krafa) / (1 + D) − A − fjármagnskostnaður á A (A × 70% × 10,2% × 1 ár) − biðtími (3 ár fasteignagjöld o.fl. að frádreginni Kolaportsleigu) − stimpilgjald og kaupkostnaður ] / (1 + 70% × 10,2% × 4,5 ár). NOI = leiga − fasteignagjöld eftir endurmat (×1,25) − tryggingar − viðhaldssjóður 0,5% af A − umsýsla 1%. Skattar ekki reiknaðir. Sjóðstreymislíkan með IRR er í Excel-skjalinu.</div></section>

<section id="s8"><div class="k">8</div><h2>Rétt verð · reiknivél</h2><p class="lead">Stilltu A, B, C, D og jarðhæðina. Verðið er afleidd stærð. A er úr áætlun ÍF v1.0 (hús fullbúið án lauss búnaðar, rekstraraðili leggur til búnað); herbergjafjöldi færir 35% af A, gæðaflokkur ±7–12% á innanhússfrágang.</p>
<div class="calc"><div class="card">
<label>Gæðaflokkur <span class="v" id="v_tier"></span></label><select id="tier"><option value="A">A · upper midscale</option><option value="B" selected>B · upper upscale</option><option value="C">C · lúxus</option></select>
<label>Herbergi <span class="v" id="v_keys"></span></label><input type="range" id="keys" min="85" max="130" value="123" step="1">
<label>Kostnaðarstig gagnvart áætlun ÍF v1.0 <span class="v" id="v_cf"></span></label><input type="range" id="cf" min="0.6" max="1.2" value="1.0" step="0.05">
<label>C · leiguhlutfall af heildartekjum <span class="v" id="v_C"></span></label><input type="range" id="C" min="0.15" max="0.34" value="0.25" step="0.005">
<label>Jarðhæð <span class="v" id="v_ground"></span></label><select id="ground"><option value="kolaport">Kolaportið á núverandi leigu (45 m.kr/ár)</option><option value="matarholl">Matarhöll/markaður á markaðsleigu (5.500 kr/m²/mán = 159 m.kr/ár)</option></select>
<label>D · arðsemiskrafa fjárfesta, álag á kostnað <span class="v" id="v_D"></span></label><input type="range" id="D" min="0" max="0.30" value="0.20" step="0.01">
<label>Ávöxtunarkrafa kaupanda á NOI <span class="v" id="v_y"></span></label><input type="range" id="y" min="0.05" max="0.08" value="0.0675" step="0.0025">
<p class="src" style="margin-top:14px">Sjálfgefið: B-flokkur, 123 herbergi, áætlun ÍF v1.0, ODDSSON-leiga 25%, Kolaportið óbreytt, D 20%, krafa 6,75%.</p>
</div>
<div><table class="t wf" id="wf"></table><div style="margin-top:14px">Rétt verð fyrir húsið eins og það er: <span class="res" id="res"></span> <span class="src" id="res2"></span></div>
<div class="callout" id="verdict"></div></div></div></section>

<section id="s8b"><div class="k">8b</div><h2>Bakreikningur · hvað þarf nóttin að kosta?</h2><p class="lead">Sama keðja snúin við: tilboðsverð sem hlutfall af fasteignamati → heildarkostnaður → verðmæti sem fjárfestar þurfa → NOI → leiga → tekjur → meðalverð á selda nótt. Excel: „Tollhúsið - Bakreikningur v1.0".</p>
<div class="calc"><div class="card">
<label>Tilboðsverð, hlutfall af fasteignamati 4.725 m.kr <span class="v" id="b_v_pct"></span></label><input type="range" id="b_pct" min="0" max="1" value="0.42" step="0.01">
<label>Herbergi <span class="v" id="b_v_keys"></span></label><input type="range" id="b_keys" min="85" max="130" value="123" step="1">
<label>Nýting á ársgrundvelli <span class="v" id="b_v_occ"></span></label><input type="range" id="b_occ" min="0.6" max="0.85" value="0.75" step="0.01">
<label>Veitingar og annað, hlutfall af herbergistekjum <span class="v" id="b_v_fb"></span></label><input type="range" id="b_fb" min="0.1" max="0.6" value="0.40" step="0.05">
<label>C · leiga sem hlutfall af heildartekjum <span class="v" id="b_v_C"></span></label><input type="range" id="b_C" min="0.18" max="0.32" value="0.25" step="0.005">
<label>Kolaportsleiga: núverandi kr/m²/mán × margfaldari <span class="v" id="b_v_kx"></span></label><input type="range" id="b_kx" min="1" max="4" value="1" step="0.5">
<label>D · arðsemiskrafa fjárfesta, álag á kostnað <span class="v" id="b_v_D"></span></label><input type="range" id="b_D" min="0" max="0.3" value="0.20" step="0.01">
<label>Ávöxtunarkrafa kaupanda á NOI <span class="v" id="b_v_y"></span></label><input type="range" id="b_y" min="0.05" max="0.08" value="0.0675" step="0.0025">
<label>Kostnaðarstig gagnvart áætlun ÍF v1.0 <span class="v" id="b_v_cf"></span></label><input type="range" id="b_cf" min="0.7" max="1.2" value="1.0" step="0.05">
<label><input type="checkbox" id="b_ffe"> Eigandi leggur til lausan búnað (739 m.kr) í stað rekstraraðila</label>
<p class="src" style="margin-top:12px">Kolaportsleiga í dag ≈ 1.000 kr/m²/mán á 2.410 m² (29 m.kr/ár, staðfesta með samningi ríkis og borgar). Fjármögnun: 70% byggingarlán á 10,2% PIK, kaupverð fjármagnað 4,5 ár, framkvæmd 1 ár að meðaltali, biðtími 3 ár með fasteignagjöldum 93 m.kr/ár.</p>
</div>
<div><table class="t wf" id="b_wf"></table>
<div style="margin-top:14px">Nauðsynlegt meðalverð á selda nótt: <span class="res" id="b_res"></span><div class="src" id="b_res2" style="margin-top:4px"></div></div>
<div class="callout" id="b_verdict"></div>
<div class="grid g3" style="margin-top:10px"><div class="card"><h3>Íslandshótel 2023</h3><div class="big">25,4 <small>þ.kr án VSK</small></div><p>Öll keðjan, nýting 68%</p></div><div class="card"><h3>Upper upscale, B</h3><div class="big">48 <small>þ.kr án VSK</small></div><p>Konsulat/Parliament-flokkur, €330</p></div><div class="card"><h3>Lúxus, C</h3><div class="big">65 <small>þ.kr án VSK</small></div><p>EDITION-flokkur, €450</p></div></div>
</div></div></section>
<section id="s9"><div class="k">9</div><h2>Sviðsmyndir</h2><p class="lead">123 herbergi, B- og C-flokkur, leiga eftir USALI (65%) eða ODDSSON (25%), kostnaður skv. áætlun ÍF v1.0 eða 20% undir, Kolaport eða matarhöll, þróunaraðili (D 20%, 6,75%) eða langtímaeigandi (D 10%, 6,0%). m.kr án lauss búnaðar.</p>
{tbl(head_scen, sel)}
<div class="callout"><b>Lesturinn.</b> Sviðsmyndirnar eru reiknaðar með áætlun ÍF v1.0 sem grunn (án lauss búnaðar). Kostnaðarstig 100% er áætlunin; 80% er 20% undir henni. Bakreikningurinn á glæru 8b er þægilegri leið að sömu niðurstöðu: hann sýnir hvaða verð á nótt hvert kaupverð krefst.</div></section>

<section id="s10"><div class="k">10</div><h2>Samkeppni og tilboðsform</h2>
<div class="grid g2">
<div>{tbl(['Bjóðandi', 'Verð sem hann gæti boðið', 'Forsendur'], [
 ['Fasteignafélag (Reitir, Heimar, Regin, Kaldalón) – skrifstofur eftir endurbætur', '~1.000 m.kr', '7.300 m² á 4.800 kr/m²/mán, Kolaport 45, NOI 85%, krafa 6,25%, endurbætur á sama byggingarhluta, 10% álag'],
 ['Hótelfélag sem á og rekur sjálft (Íslandshótel, Berjaya, Keahótel)', 'neikvætt til ~500', 'Sama umbreyting, EBITDA − FF&E, krafa 8,5%, 15% álag'],
 ['Þróunarfélag sem verðleggur byggingarrétt', '1.500–2.500', 'Kaupir tíma og skipulagsleið; hærra ef það sér 5.000–9.000 m² viðbót'],
 ['Viðmið seljanda', '2.000–4.700', 'Eigið mat 2022 ≥2 ma.kr; fasteignamat 4.725; brunabótamat 6.000'],
], 'wrap')}
<div class="callout">Ríkið velur að öllu jöfnu hæsta gilda tilboð en má hafna öllum. Faxaflóahafnir eiga forkaupsrétt samkvæmt kvöð. Söluheimildin er í fjárlagafrumvarpi 2027 og afhending því vart fyrr en 2027.</div></div>
<div><h3 style="margin:0 0 8px;color:var(--navy)">Tilboðsform sem GT hefur ákveðið</h3><ul class="tight">
<li>Eitt verð, afleitt af A–D, ekki bundið neinni fyrirframtölu.</li><li>Walk-away fyrirvari um deiliskipulag: hótel auk tilgreinds viðbótarbyggingarmagns; gangi breytingin ekki í gegn fellur samningurinn niður.</li><li>Fyrirvari um fjármögnun og um nýjan lóðarleigusamning við Reykjavíkurborg með skilgreindum lágmarksskilmálum.</li><li>Þróunarfélag ÍF og Næsland-hópsins, ÍF þróunaraðili og verkefnastjóri með þóknun sem eignarhlut; Íslandshótel hugsanlega í félaginu.</li><li>Yfirlýsing um Kolaportið og opinn þakgarð sem sölutæki gagnvart ríki og borg.</li></ul>
<h3 style="margin:16px 0 8px;color:var(--navy)">Hvað þarf að vera satt fyrir jákvætt verð</h3><ul class="tight">
<li>Umbreyting á eða undir áætlun ÍF v1.0, staðfest með verktakatilboðum í COWI-liðina.</li><li>Rekstraraðili sem skrifar undir 25 ára samning á 25% af tekjum með grunnleigu og móðurfélagsábyrgð.</li><li>Jarðhæðin á markaðsleigu, með borginni, ekki gegn henni.</li><li>Langtímaeigandi sem kaupir fullbúna eign á 6,0% kröfu, samið fyrirfram.</li></ul></div>
</div></section>

<section id="s11"><div class="k">11</div><h2>Næstu skref og heimildir</h2>
<div class="grid g2"><div><ul class="tight">
<li><b>Fundur með Næslandi:</b> sýna bilið milli þess sem hótelið stendur undir og þess sem ríkið ætlast til; fá hlutverk Íslandshótela skjalfest.</li>
<li><b>Íslandshótel:</b> spyrja beint um verðflokk, leiguhlutfall og form samnings.</li>
<li><b>Kostnaðarpróf:</b> verktakatilboð í frárennsli, glugga, klæðningu og asbest til að festa stærstu matsliðina í áætlun v1.0.</li>
<li><b>Skipulag:</b> Kvosardeiliskipulag með 2015-breytingunni (hlutfall gististaða) og afstaða borgarinnar til viðbótar ofan á húsið.</li>
<li><b>TORG:</b> tilboðsreglur, Kolaportssamningur ríkis og borgar, lóðarleigusamningur 1967/1993, hvort fjárlagaheimild sé í gildi.</li>
<li><b>Langtímaeigandi:</b> þreifa á fasteignafélagi eða lífeyrissjóði um framvirk kaup.</li>
<li><b>Tilboðsbréf og skilmálar:</b> drög þegar A og C hafa verið fest.</li></ul></div>
<div><ul class="tight" style="font-size:12px">
<li>Söluyfirlit TORG 31.8.2026; fasteigna- og veðbandayfirlit HMS júlí 2026.</li><li>COWI: ástandsskýrsla 3210316-000-CRP-0003 (16.10.2024) með burðarþolsminnisblaði 6.9.2024; innivistarskoðun A290951-001 (20.8.2025).</li><li>Aðaluppdrættir BN045618 (2013, 1:200) og BN048223 (2014), skjalasafn Reykjavíkur.</li><li>Umsögn skipulagsfulltrúa 4.4.2024; minnisblað ríkis og borgar um LHÍ 20.4.2022; Stjórnarráðið 11.9.2024.</li><li>Kostnaðaráætlun ÍF fullbúið hótel v1.0 (23.9.2026) og reynslutölur ÍF af sambærilegum verkefnum (trúnaðarmál, ekki birtar).</li><li>Íslandshótel hf.: útgefandalýsing 13.5.2024, fjárfestakynning maí 2024, ársreikningur 2024; Reitir tilkynning 29.4.2026.</li><li>ÍF/RR hótel: leigusamningur ODDSSON 2020 með viðauka 2021; áætlun 06/2020.</li><li>Flóra/ÍF: Black Dunes Þorlákshöfn rekstraráætlun 17.3.2025, BBR módel v1 2026.</li><li>KPMG: Hótelgeirinn á Íslandi 2014 (Ferðamálastofa). Hagstofa Íslands: gistinætur 2025–26. HVS: hotel lease structures.</li><li>Fjármögnunar- og ávöxtunarforsendur ÍF (2026) og greining á ávöxtunarkröfu fasteignafélaga 19.8.2026.</li><li>Gögn Næslands 22.9.2026: kynning, stofnkostnaðar- og rekstraráætlun, samantekt COWI-flagga.</li></ul></div></div>
<p class="src" style="margin-top:20px">Líkan: model.py / model2.py / model3.py og Excel v0.1 í Claude Projects/Tollhúsið Næsland/. Þessi síða er drög 1, 22.9.2026, og verður uppfærð þegar A og C hafa verið fest.</p></section>
</main>
<script>
"use strict";
const TIER={{A:{{adr:36000,occ:.72,fb:.20,other:.03,innan:210,ffe:3}},B:{{adr:48000,occ:.72,fb:.35,other:.05,innan:254,ffe:5}},C:{{adr:65000,occ:.68,fb:.45,other:.08,innan:320,ffe:8}}}};
const A_CONV=6879.7,A_BOH=861.5,KOLA=2409.6;
const A_V1={A_HUS:.1f},FFE_V1={A_FFE:.1f};
function capex(t,keys,cf,ground){{const tierAdj={{A:0.93,B:1.0,C:1.12}}[t];const base=(A_V1*0.65+A_V1*0.35*keys/123)*tierAdj;const g=ground==='matarholl'?KOLA*220/1000*1.3:0;return {{total:(base+g)*cf}};}}
function rev(t,keys){{const T=TIER[t];const rooms=keys*T.adr*T.occ*365/1e6;return {{rooms,total:rooms*(1+T.fb+T.other)}};}}
function calc(){{const t=tier.value,k=+keys.value,c=+cf.value,C=+Cc.value,g=ground.value,D=+Dd.value,y=+yy.value;
 const X=capex(t,k,c,g),Rv=rev(t,k),gr=g==='kolaport'?45.3:KOLA*5500*12/1e6;const leiga=C*Rv.total+gr;const fixed=(85.48+7.83)*1.25+5.5*1.5+X.total*.005;const noi=leiga-fixed-leiga*.01;const V=noi/y;const FA=X.total*.7*.102*1.0;const Hh=(85.48+7.83+5.53+10)*3-45.3*2;const K=4725.5*.016+20;const price=(V/(1+D)-X.total-FA-Hh-K)/(1+.7*.102*4.5);
 const f=x=>x.toLocaleString('de-DE',{{maximumFractionDigits:0}});
 v_tier.textContent=t;v_keys.textContent=k;v_cf.textContent=Math.round(c*100)+'% af áætlun';v_C.textContent=(C*100).toFixed(1).replace('.',',')+'%';v_ground.textContent=g==='kolaport'?'45 m.kr':'159 m.kr';v_D.textContent=Math.round(D*100)+'%';v_y.textContent=(y*100).toFixed(2).replace('.',',')+'%';
 const rows=[['A · framkvæmd án VSK (m.kr/herb '+f(X.total/k)+')',-X.total],['B · tekjur hótels (herbergi '+f(Rv.rooms)+')',Rv.total],['Leiga hótels + jarðhæð ('+f(gr)+')',leiga],['Fastur eigandakostnaður og umsýsla',-(fixed+leiga*.01)],['NOI eiganda',noi],['Verðmæti = NOI / '+(y*100).toFixed(2).replace('.',',')+'%',V],['÷ (1 + D) → verðmæti að frádreginni arðsemiskröfu',V/(1+D)],['− A',-X.total],['− fjármagn á A',-FA],['− biðtími 3 ár (fasteignagjöld o.fl. − Kolaport)',-Hh],['− stimpilgjald og kaupkostnaður',-K],['÷ (1 + fjármagn á kaupverð 4,5 ár)','']];
 wf.innerHTML=rows.map(r=>'<tr><td>'+r[0]+'</td><td>'+(r[1]===''?'':(r[1]<0?'<span class="neg">':'')+f(r[1])+(r[1]<0?'</span>':''))+'</td></tr>').join('');
 res.textContent=f(price)+' m.kr';res.className='res'+(price<0?' neg':'');res2.textContent=' · '+f(price/k)+' m.kr á herbergi · '+f(price/10150.8*1000)+' þ.kr/m² núverandi húss';
 verdict.innerHTML=price<0?'<b>Neikvætt:</b> við þessar forsendur ber verkefnið ekki kaupverð; það vantar '+f(-price)+' m.kr upp á þótt húsið væri gefið. Færðu C upp, kostnaðarstigið niður eða D niður til að sjá hvað þarf að vera satt.':'<b>Jákvætt:</b> þetta er hæsta verð fyrir húsið eins og það er sem skilar fjárfestum D-álaginu að gefnum A, B og C. Berðu saman við viðmið seljanda: eigið mat 2022 ≥2.000, fasteignamat 4.725.';}}
const tier=document.getElementById('tier'),keys=document.getElementById('keys'),cf=document.getElementById('cf'),Cc=document.getElementById('C'),ground=document.getElementById('ground'),Dd=document.getElementById('D'),yy=document.getElementById('y');
[tier,keys,cf,Cc,ground,Dd,yy].forEach(e=>e.addEventListener('input',calc));calc();
function bak(){{const pct=+b_pct.value,keys=+b_keys.value,occ=+b_occ.value,fb=+b_fb.value,C=+b_C.value,kx=+b_kx.value,D=+b_D.value,y=+b_y.value,c=+b_cf.value,ffe=b_ffe.checked?1:0;
 const price=4725.5*pct;const A=(A_V1*0.65+A_V1*0.35*keys/123+ffe*FFE_V1*keys/123)*c;const K=4725.5*.016+20;const FA=A*.7*.102;const FP=price*.7*.102*4.5;const kolaNow=2409.6*1000*12/1e6;const Hh=(93.31+15.5)*3-kolaNow*2;
 const T=price+K+A+FA+FP+Hh;const V=T*(1+D);const noi=V*y;const fixed=93.31*1.25+8.25+A*.005;const leiga=(noi+fixed)/.99;const kola=kolaNow*kx;const hrent=leiga-kola;const rev=hrent/C;const rooms=rev/(1+fb);const nights=keys*365*occ;const adr=rooms*1e6/nights;const adrv=adr*1.11;
 const f=x=>x.toLocaleString('de-DE',{{maximumFractionDigits:0}});
 b_v_pct.textContent=Math.round(pct*100)+'% = '+f(price)+' m.kr';b_v_keys.textContent=keys;b_v_occ.textContent=Math.round(occ*100)+'%';b_v_fb.textContent=Math.round(fb*100)+'%';b_v_C.textContent=(C*100).toFixed(1).replace('.',',')+'%';b_v_kx.textContent='×'+kx+' = '+f(kola)+' m.kr/ár';b_v_D.textContent=Math.round(D*100)+'%';b_v_y.textContent=(y*100).toFixed(2).replace('.',',')+'%';b_v_cf.textContent=Math.round(c*100)+'%';
 const rows=[['Tilboðsverð',price],['Stimpilgjald og kaupkostnaður',K],['A · framkvæmd ('+f(A/keys)+' m.kr/herb)',A],['Fjármagnskostnaður á framkvæmd og kaupverð',FA+FP],['Biðtími 3 ár að frádreginni Kolaportsleigu',Hh],['Heildarkostnaður',T],['× (1 + D) = verðmæti sem fjárfestar þurfa',V],['× krafa = nauðsynlegt NOI á ári',noi],['+ fastur eigandakostnaður = leigutekjur alls',leiga],['− jarðhæð ('+f(kola)+') = leiga hótels',hrent],['÷ C = heildartekjur hótels',rev],['÷ (1 + F&B) = herbergistekjur',rooms],['÷ '+f(nights)+' seldar nætur',adr]];
 b_wf.innerHTML=rows.map(r=>'<tr><td>'+r[0]+'</td><td>'+f(r[1])+'</td></tr>').join('');
 b_res.textContent=f(adr)+' kr án VSK';b_res.className='res'+(adr>70000?' neg':'');
 b_res2.textContent=f(adrv)+' kr með VSK · €'+f(adrv/145)+' · sumar ≈ '+f(adrv*1.35)+' · vetur ≈ '+f(adrv*0.78)+' · RevPAR '+f(adr*occ)+' · tekjur á herbergi '+(rev/keys).toFixed(1).replace('.',',')+' m.kr · rekstraraðili heldur ≈ '+Math.round((0.37-0.04-C)*100)+'% eftir leigu';
 let v='';if(adr<30000)v='<b>Millistig.</b> Þetta er verð sem íslensk hótelkeðja nær í dag; kaupverðið er raunhæft á þessum forsendum.';else if(adr<50000)v='<b>Efra millistig / upper upscale.</b> Konsulat- og Parliament-verð; krefst vörumerkis, F&B og markaðssetningar, en er til í Reykjavík.';else if(adr<68000)v='<b>Lúxus.</b> EDITION-verðflokkur á ársgrundvelli. Aðeins fáein hótel á Íslandi ná þessu; rekstraraðili þarf að staðfesta.';else v='<b>Yfir markaðnum.</b> Ekkert hótel í Reykjavík nær þessu meðalverði yfir árið. Kaupverðið, D eða krafan þurfa að lækka, eða Kolaportsleigan að hækka.';
 b_verdict.innerHTML=v;}}
['b_pct','b_keys','b_occ','b_fb','b_C','b_kx','b_D','b_y','b_cf','b_ffe'].forEach(id=>document.getElementById(id).addEventListener('input',bak));bak();
const secs=[...document.querySelectorAll('section')],links=[...document.querySelectorAll('nav a')];
document.addEventListener('keydown',e=>{{if(e.target.tagName==='INPUT'||e.target.tagName==='SELECT')return;const y=window.scrollY+10;let i=secs.findIndex(s=>s.offsetTop>y);if(i<0)i=secs.length;if(e.key==='ArrowRight'||e.key==='PageDown'||e.key===' '){{e.preventDefault();(secs[Math.min(i,secs.length-1)]||secs[0]).scrollIntoView();}}if(e.key==='ArrowLeft'||e.key==='PageUp'){{e.preventDefault();(secs[Math.max(i-2,0)]).scrollIntoView();}}}});
const io=new IntersectionObserver(es=>es.forEach(en=>{{if(en.isIntersecting){{links.forEach(l=>l.classList.toggle('on',l.getAttribute('href')==='#'+en.target.id));}}}}),{{threshold:.4}});secs.forEach(s=>io.observe(s));
</script></body></html>'''
open(os.path.join(HERE, 'app.html'), 'w', encoding='utf-8').write(HTML)
print('app.html', len(HTML.encode()) // 1000, 'kB')

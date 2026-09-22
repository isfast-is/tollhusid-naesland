#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tollhúsið – Næsland: mælaborð með glærum. python3 model3.py && python3 gen.py && python3 encrypt.py"""
import base64, json, os, html as H
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
.hero{{background:var(--navy);color:#fff;min-height:100vh;display:flex;align-items:center;padding:60px 70px}}.hero h2{{color:#fff;font-size:44px;margin:0 0 10px}}.hero .sub{{color:var(--gold);font-size:18px;margin-bottom:26px}}.hero p{{max-width:720px;font-size:15px;line-height:1.6;color:#C9D6E2}}
.hero img{{position:absolute;right:0;top:0;bottom:0;height:100vh;width:46%;object-fit:cover;opacity:.55}}.hero .txt{{position:relative;z-index:2;max-width:54%}}
.plans img{{width:100%;border:1px solid var(--line);border-radius:8px;background:#fff}}.plans figcaption{{font-size:12px;color:var(--muted);margin:6px 0 14px}}
ul.tight{{margin:6px 0 0 18px;padding:0;font-size:13.5px;line-height:1.55}}ul.tight li{{margin-bottom:5px}}
.calc{{display:grid;grid-template-columns:330px 1fr;gap:26px}}.calc label{{display:block;font-size:12.5px;margin:10px 0 2px;color:var(--muted)}}.calc select,.calc input[type=range]{{width:100%}}.calc .v{{float:right;color:var(--navy);font-weight:600}}
.wf td:last-child{{text-align:right;font-variant-numeric:tabular-nums}}.res{{font-size:34px;font-weight:800;color:var(--navy)}}.res.neg{{color:var(--red)}}
@media(max-width:1000px){{nav{{display:none}}main{{margin:0}}section{{padding:28px 18px}}.g2,.g3,.g4,.calc{{grid-template-columns:1fr}}.hero img{{display:none}}.hero .txt{{max-width:100%}}}}
@media print{{nav{{display:none}}main{{margin:0}}section{{page-break-after:always;min-height:auto}}}}
</style></head><body>
<nav><h1>TOLLHÚSIÐ</h1><div class="sub">Næsland · greining fyrir tilboð · trúnaðarmál</div>
<a href="#s0">Forsíða</a><a href="#s1">Í hnotskurn</a><a href="#s2">Eignin og áhætturnar</a><a href="#s3">Teikningar og herbergi</a><a href="#s4">A · Umbreyting</a><a href="#s5">B · Tekjur</a><a href="#s6">C · Leiga</a><a href="#s7">D · Fjármögnun og arðsemi</a><a href="#s8">Rétt verð · reiknivél</a><a href="#s9">Sviðsmyndir</a><a href="#s10">Samkeppni og tilboðsform</a><a href="#s11">Næstu skref og heimildir</a>
<div class="foot">Drög 1 · 22.9.2026 · líkan v0.3<br>Örvatakkar fletta glærum<br><button onclick="try{{localStorage.removeItem('tollhus_pass')}}catch(e){{}};location.reload()">Læsa síðu</button></div></nav>
<main>
<section id="s0" class="hero"><img src="{b64('hero.jpg')}" alt=""><div class="txt"><div class="k">ÍSLENSKAR FASTEIGNIR · NÆSLAND · 22. SEPTEMBER 2026</div><h2>Tollhúsið, Tryggvagötu 19</h2><div class="sub">Hvað stendur hótelverkefni undir háu kaupverði? Afleiðsla í fjórum skrefum: A framkvæmd · B tekjur · C leiga · D arðsemi</div>
<p>Ríkissjóður selur Tollhúsið, 10.150,8 m² á 4.862 m² leigulóð, í gegnum Fasteignasöluna TORG. Tilboðsfrestur er föstudaginn 9. október 2026. Næsland-hópurinn (Jón Haukur Baldvinsson, Daníel Freyr Atlason, ásamt Íslandshótelum) leitar til ÍF sem þróunaraðila. Þessi síða leiðir kaupverðið út sem afleidda stærð, með teikningum hússins, raunkostnaði Hyatt Centric, íslenskum leiguviðmiðum og Hamranes-fjármögnun. Engin tala er föst fyrirfram.</p>
<p style="color:#7F91A3;font-size:12.5px">Vinnuskjal GT/ÍF. Öll gögn frá Jóni Hauki, COWI, TORG, skjalasafni Reykjavíkur og isfast-drifi. Ekkert hefur farið út úr húsi.</p></div></section>

<section id="s1"><div class="k">1</div><h2>Í hnotskurn</h2><p class="lead">Fimm staðreyndir sem ráða málinu, og aðferðin sem síðan fylgir.</p>
<div class="grid g4">
<div class="card"><h3>Húsið ber</h3><div class="big">106–123 <small>herbergi</small></div><p>Lesið af grunnmyndum 1:200 (BN045618). Jón og Daníel skipuðu 89–100 með lobby á 3. hæð. Jarðhæðin, afgreiðsla tollstjóra 730 m² bak við mósaíkið, er lobbyið.</p></div>
<div class="card"><h3>Umbreyting kostar</h3><div class="big">6,5–8,1 <small>ma.kr án VSK</small></div><p>Á raunköflum Hyatt Centric L176 (894 þ.kr/m², EAC júlí 2026). Það eru 53–66 m.kr á herbergi við 123 herbergi. Jón og Daníel gerðu ráð fyrir 1,4 ma.kr, 17 m.kr á herbergi.</p></div>
<div class="card"><h3>Leiga sem hótel þolir</h3><div class="big">20–25% <small>af heildartekjum</small></div><p>USALI-uppbygging 20–24%, ODDSSON-samningur ÍF 25% (30% af herbergistekjum). JHB miðaði við 34% af herbergistekjum án F&B, sem jafngildir 22–28%.</p></div>
<div class="card"><h3>Rétt verð fyrir húsið</h3><div class="big">−2,7 til +2,6 <small>ma.kr</small></div><p>Eftir því hvaða A, C, D og jarðhæðarnýting er valin. Grunnsviðsmyndir eru neikvæðar; jákvæðar tölur krefjast lúxusflokks, kostnaðar 20% undir Hyatt og langtímaeiganda.</p></div>
</div>
<div class="callout"><b>Aðferð.</b> Verðið er afleidd stærð: verðmæti fullbúins hótels (NOI eiganda / ávöxtunarkrafa) að frádreginni arðsemiskröfu fjárfesta (D), umbreytingu (A), fjármagnskostnaði og biðtíma, gefur það sem hægt er að greiða fyrir húsið eins og það er. Hver liður er reitur í reiknivélinni á glæru 8 og í Excel-líkaninu. Tilboðið verður eitt verð með walk-away fyrirvara um deiliskipulag (hótel + tilgreint viðbótarbyggingarmagn), fjármögnun og lóðarleigusamning.</div>
<div class="grid g3">
<div class="card"><h3>Það sem er sterkt</h3><ul class="tight"><li>Burðarvirki og steypa í góðu lagi, ein hæð til viðbótar möguleg burðarþolslega (COWI 6.9.2024)</li><li>Staðsetning, kennileiti, mósaík Gerðar Helgadóttur, torgið</li><li>Kolaportið og þakflötur 3. hæðar sem sölutæki gagnvart ríki og borg</li><li>Hæðarhæð 3,2–3,45 m og 15,5 m djúp álma hentar tvíhliða hótelgangi</li></ul></div>
<div class="card"><h3>Það sem er veikt</h3><ul class="tight"><li>Hótel ekki heimilt í deiliskipulagi Kvosar; lóðarleigusamningur útrunninn 2017</li><li>Asbest 550–600 m², PCB, mygla á öllum hæðum nema 5., frárennsli í grunni ónýtt, 1.600 m² gluggaveggir</li><li>2. hæð er einhliða (norðurhlið inni í Kolaportssal); 84 m² brúttó á herbergi í stað 59 hjá Hyatt</li><li>Fasteignagjöld 93 m.kr á ári frá kaupdegi; forkaupsréttur og samþykki hafnarstjórnar</li></ul></div>
<div class="card"><h3>Það sem ræður úrslitum</h3><ul class="tight"><li><b>Kostnaðarprófið:</b> getum við byggt marktækt undir Hyatt? Hyatt fór 75% fram úr fyrstu áætlun.</li><li><b>Jarðhæðarprófið:</b> Kolaportið greiðir 45 m.kr fyrir 2.410 m²; markaðsleiga væri 145–200.</li><li><b>Eigendaprófið:</b> D 20% á 6,75% eða langtímaeigandi á D 10% og 6,0% munar 1–2 ma.kr.</li><li><b>Rekstraraðilinn:</b> hver skrifar undir 25 ára leigu og með hvaða ábyrgð.</li></ul></div>
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
<figure><img src="{b64('plan2.jpg')}" alt="2. hæð"><figcaption><b>2. hæð, 1.591 m².</b> Norðurhlið álmunnar liggur inn í Kolaportssalinn: aðeins herbergi sunnan megin, 14–17. Jón og Daníel settu þar 11 og gym.</figcaption></figure>
<figure><img src="{b64('plan1.jpg')}" alt="1. hæð"><figcaption><b>1. hæð, 3.515 m².</b> Kolaportssalur 2.410 m². Austurblokkin, afgreiðsla tollstjóra 730 m² við Pósthússtræti og Steinbryggju, er lobby og veitingastaður; vesturblokkin starfsmannarými.</figcaption></figure>
</div>
<div class="grid g2" style="margin-top:8px"><div>{tbl(['Skipan', '5. hæð', '4. hæð', '3. hæð', '2. hæð', 'Alls'], keys_rows)}<p class="src">Talning: 30% í kjarna og ganga af brúttó; herbergi 3,9 m eða 4,5 m breið og 6,5 m djúp. Nákvæm skipan bíður arkitekts, en bilið 106–123 er það sem húsið ber.</p></div>
<figure><img src="{b64('snid.jpg')}" alt="Snið"><figcaption><b>Snið A-A og B-B.</b> Gólfkótar 3,90 / 7,10 / 10,40 / 13,60 / 17,05. Þakplata 3. hæðar og 5. hæðar þola eina hæð til viðbótar (COWI 6.9.2024).</figcaption></figure></div></section>

<section id="s4"><div class="k">4 · A</div><h2>Umbreytingarkostnaður</h2><p class="lead">Kaflar úr raunkostnaði Hyatt Centric Laugavegi 176 (EAC júlí 2026, ÍF byggingarstjóri), lagaðir að Tollhúsinu. Allt án VSK á verðlagi 2026.</p>
<div class="grid g2">
<div>{tbl(['Kafli', 'Hyatt þ.kr/m²', 'Tollhúsið þ.kr/m²', 'Rök'], [
 ['Aðstaða og umsjón', '50', '45', 'Styttri verktími'],
 ['Jarðvinna og niðurrif', '30', '25', 'Niðurrif innanhúss, engin jarðvinna'],
 ['Burðarvirki', '90', '30', 'Hyatt byggði nýjar hæðir; hér op, lyftuop, stigar, styrkingar'],
 ['Lagnir og loftræsing', '78', '78', 'Nýtt kerfi; COWI: miðrými án loftræstingar'],
 ['Raflagnir og kerfi', '113', '100', ''],
 ['Frágangur innanhúss', '254', 'A 210 / B 254 / C 320', 'Gæðaflokkur ræður'],
 ['Frágangur utanhúss', '118', '118', 'Gluggaveggir 1.600 m² + klæðning skv. COWI'],
 ['Lóð', '15', '8', ''],
 ['<b>Framkvæmd</b>', '<b>760</b>', '<b>614 / 658 / 724</b>', 'á 6.880 m² sem breytt er í hótel'],
 ['Hönnun, stjórnun, umsýsla', '118 (15,6%)', '13% + 3% þóknun ÍF', 'Hyatt: arkitekt 288, verkfræði 373, eftirlit 311 m.kr'],
 ['<b>Alls</b>', '<b>894 = 52,5 m.kr/herb</b>', 'sjá töflu', 'Hyatt: 9.923 m², 169 herb., 59 m²/herb'],
], 'wrap')}
<div class="callout">Hyatt var kynnt opinberlega á „rúmlega 5 ma.kr“ 2023. EAC er 8,9 ma.kr án VSK. Aðstaða fór úr 168 m.kr samningi í 427 með aukaverkum. Ófyrirséð í þessu líkani er 10% og kostnaðarstigið er sleði í reiknivélinni.</div></div>
<div>{tbl(['Fl.', 'Skipan', 'Herb.', 'Breytt rými', 'Kjallari + COWI-liðir', 'Ófyrirséð', 'Hönnun + þóknun', 'FF&E', 'A alls', 'm.kr/herb'], capex_rows)}
<p class="src">COWI-liðir: asbest/PCB/mygla 180, frárennsli í grunni 120, Kolaportið 120. Kjallari og tæknirými 861 m² á 200 þ.kr/m². FF&E á herbergi: A 3, B 5, C 8 m.kr. Til samanburðar: JHB 1.368 án VSK (17 m.kr/herb); KEF-tilboð ÍF 30–36 m.kr/herb nýbygging án lóðar.</p>
<div class="callout">A á herbergi ræðst af herbergjafjöldanum, ekki einingarverðum: sama hús kostar 6,5–8 ma.kr hvort sem herbergin verða 89 eða 123. Þess vegna er skipan hæðanna fyrsta hönnunarverkefnið, ekki verðskrá.</div></div>
</div></section>

<section id="s5"><div class="k">5 · B</div><h2>Tekjur hótelsins</h2><p class="lead">Þrír gæðaflokkar. ADR er meðalverð á selda nótt yfir árið án 11% VSK. Tekjur miðaðar við 123 herbergi.</p>
<div class="grid g2">
<div>{tbl(['Fl.', 'ADR án VSK', 'Nýting', 'RevPAR', 'Herbergistekjur', 'F&B og annað', 'Tekjur alls', 'm.kr/herb'], pl_rows)}
{tbl(['Flokkur', 'Lýsing', 'Viðmið'], [
 ['A', tiers['A']['name'], 'JHB-verðskrá 55/42/32 þ.kr m/VSK → ~36 þ.kr án VSK að meðaltali; €245'],
 ['B', tiers['B']['name'], '€330; Konsulat, Iceland Parliament, Sand'],
 ['C', tiers['C']['name'], '€450; EDITION-flokkur, krefst vörumerkis og þjónustu'],
], 'wrap')}</div>
<div>{tbl(['Markaðsviðmið', 'Gildi', 'Heimild'], [
 ['Íslandshótel, öll keðjan 2023', 'ADR 25,4 þ.kr · nýting 68% · RevPAR 17,2 þ.kr · 8,5 m.kr tekjur á herbergi (2024)', 'Útgefandalýsing maí 2024, ársreikningur 2024'],
 ['Höfuðborgarsvæðið 2025', 'Nýting 74,7%; 58 hótel, 5.555 herbergi', 'Hagstofa Íslands'],
 ['Black Dunes Þorlákshöfn (Flóra, 120 herb.)', 'ADR 29–32 þ.kr · nýting 67–74% (2027–32)', 'Rekstraráætlun 17.3.2025, isfast Drive'],
 ['ODDSSON 2021 (ÍF, 77 herb.)', 'Brúttó ADR €60–80 · nýting 70–85% eftir mánuðum', 'Áætlun 06/2020, isfast Drive'],
 ['JHB Næsland ár 3', 'Herbergistekjur 1.066 án VSK, nýting 74%, engar F&B-tekjur', 'Rekstrar- og söluáætlun draft 1'],
 ['Markaðurinn 2025', 'Nýting féll, RevPAR staðið í stað í krónum í tvö ár; sumar 2026 gott', 'Arion greining 2025, mbl 2.9.2026'],
], 'wrap')}
<div class="callout">Tekjur á herbergi eru 11,6 / 17,7 / 24,7 m.kr eftir flokki. Íslandshótel eru á 8,5 að meðaltali yfir allt landið. Flokkur C tvöfaldar tekjurnar á herbergi en fækkar herbergjum um fimmtung og kallar á vörumerki, F&B-rekstur og þjónustustig sem þarf að staðfesta með rekstraraðila.</div></div>
</div></section>

<section id="s6"><div class="k">6 · C</div><h2>Leiga sem hótelið þolir</h2><p class="lead">Byggt upp eftir USALI-uppgjörsstaðli hótela og borið saman við íslenska samninga. 123 herbergi, m.kr á ári.</p>
{tbl(['Fl.', 'Tekjur', 'Deildarframlegð', 'GOP', 'FF&E 4%', 'Leiga USALI (65% af EBITDAR)', 'Leiga ODDSSON (max 30% herb. / 25% alls)', 'Leiga JHB (34% herb.)', 'Rekstraraðili eftir leigu (USALI / ODDSSON)', 'Leiga á herbergi, m.kr (USALI / ODDSSON)'], usali_rows)}
<p class="src">Forsendur USALI: herbergjadeild kostar 32% (laun 18%, sölukostnaður og OTA 8%, þvottur og vörur 6%); veitingadeild 75%; annað 50%; ódeilt 19% (stjórnun 7, markaðsmál 4, UT 1,5, viðhald 3,5, orka 3); enginn stjórnunarsamningur (eigin rekstur); FF&E-sjóður 4%. Fastur leigusamningur = 65% af EBITDAR eftir FF&E til leigusala (alþjóðleg venja 60–70%).</p>
<div class="grid g2" style="margin-top:14px">
<div>{tbl(['Viðmið', 'Gildi', 'Heimild'], bench_rows, 'wrap')}</div>
<div><div class="callout"><b>Niðurstaða um C.</b> Fyrir þetta hús er leiga á bilinu <b>20–25% af heildartekjum</b> raunhæf: neðri mörkin úr USALI-uppbyggingu með 65% til leigusala, efri mörkin úr samningi ÍF sjálfs við RR hótel um ODDSSON (30% af herbergistekjum eða 25% af heildartekjum). Flóra gerði ráð fyrir 30% í Þorlákshöfn og skildi rekstraraðilann eftir með 2–4% EBITDA, sem er ekki sjálfbært. Tala Jóns og Daníels, 34% af herbergistekjum, lendir á 22–28% af heildartekjum eftir flokki og er því ekki fráleit, en hún var ekki rökstudd.</div>
<div class="callout"><b>Hvað Íslandshótel greiða í raun.</b> Nýi samningurinn við Reiti gefur 1,53 m.kr NOI á herbergi á ári fyrir uppgerð fjögurra stjörnu hótel. Okkar B-flokkur þarf 3,9–4,4 og C-flokkur 5,0–6,2 á herbergi. Munurinn er verðflokkurinn: 25 þ.kr ADR á móti 48–65 þ.kr. Það er fyrsta spurningin til Íslandshótela: trúa þau á þessa verðflokka í þessu húsi?</div>
<div class="callout"><b>Form leigusamnings.</b> Blandaður samningur með grunnleigu (t.d. 70% af væntri leigu, vísitölubundinni) og 8–12% af brúttósölu, eins og í Fosshótel Austfjörðum, er bankahæfur og deilir uppsveiflunni. Hreinn veltusamningur eins og ODDSSON er ekki bankahæfur fyrir 65% LTV endurfjármögnun.</div></div>
</div></section>

<section id="s7"><div class="k">7 · D</div><h2>Fjármögnun og arðsemiskrafa</h2><p class="lead">Sömu kjör og í Hamranes-líkaninu. D er sett fram sem álag á heildarkostnað, sem jafngildir 12–15% IRR eiginfjár yfir 4–5 ára þróunartíma.</p>
<div class="grid g3">
<div class="card"><h3>Byggingartími</h3><div class="big">70/30</div><p>Byggingarlán 70% af kaupverði og framkvæmd á 10,2% óverðtryggðum PIK-vöxtum (ÍSB kjörvextir + 0,15%), lántökugjald 0,2%. Eigið fé 30%.</p></div>
<div class="card"><h3>Tímalína</h3><div class="big">2027 → 2032</div><p>Kaup 2027 eftir fjárlagaheimild, deiliskipulag og hönnun 2027, framkvæmd 2028–29, opnun 2030, leiga 80/92/100%, stöðugur rekstur 2032. Fasteignagjöld 93 m.kr á ári allan tímann.</p></div>
<div class="card"><h3>Endurfjármögnun og sala</h3><div class="big">6,0–6,75%</div><p>Verðtryggt bréf, 65% af verðmæti, 3,9% raunvextir, 25 ár. Krafa á NOI: 6,75% fyrir hótelleigu til þróunaraðila sem selur, 6,0% fyrir langtímaeiganda (Reitir matskrafa 6,4–6,7%, nýkaup 7,8%, ríkisleiga 5,7%).</p></div>
</div>
<div class="grid g2" style="margin-top:18px">
<div class="card"><h3>D = 20%, krafa 6,75%</h3><p>Þróunaraðili sem selur fullbúna eign eftir stöðugan rekstur og þarf 12–15% IRR á eigið fé. Þetta er sjónarhorn ÍF og Næsland-hópsins án langtímaeiganda.</p></div>
<div class="card"><h3>D = 10%, krafa 6,0%</h3><p>Fasteignafélag eða lífeyrissjóður sem skuldbindur sig fyrirfram til að kaupa fullbúna eign og reiknar á eigin kröfu. ÍF fær þróunarþóknun og árangurshlut. Munurinn á sjónarhornunum er 1–2 ma.kr í verði.</p></div>
</div>
<div class="callout" style="margin-top:18px"><b>Keðjan.</b> Rétt verð = [ (NOI / krafa) / (1 + D) − A − fjármagnskostnaður á A (A × 70% × 10,2% × 1 ár) − biðtími (3 ár fasteignagjöld o.fl. að frádreginni Kolaportsleigu) − stimpilgjald og kaupkostnaður ] / (1 + 70% × 10,2% × 4,5 ár). NOI = leiga − fasteignagjöld eftir endurmat (×1,25) − tryggingar − viðhaldssjóður 0,5% af A − umsýsla 1%. Skattar ekki reiknaðir. Sjóðstreymislíkan með IRR er í Excel-skjalinu.</div></section>

<section id="s8"><div class="k">8</div><h2>Rétt verð · reiknivél</h2><p class="lead">Stilltu A, B, C, D og jarðhæðina. Verðið er afleidd stærð. Sama reikniregla og í Python-líkaninu v0.3 og flipanum „Rétt verð“ í Excel.</p>
<div class="calc"><div class="card">
<label>Gæðaflokkur <span class="v" id="v_tier"></span></label><select id="tier"><option value="A">A · upper midscale</option><option value="B" selected>B · upper upscale</option><option value="C">C · lúxus</option></select>
<label>Herbergi <span class="v" id="v_keys"></span></label><input type="range" id="keys" min="85" max="130" value="123" step="1">
<label>Kostnaðarstig gagnvart Hyatt <span class="v" id="v_cf"></span></label><input type="range" id="cf" min="0.6" max="1.2" value="1.0" step="0.05">
<label>C · leiguhlutfall af heildartekjum <span class="v" id="v_C"></span></label><input type="range" id="C" min="0.15" max="0.34" value="0.25" step="0.005">
<label>Jarðhæð <span class="v" id="v_ground"></span></label><select id="ground"><option value="kolaport">Kolaportið á núverandi leigu (45 m.kr/ár)</option><option value="matarholl">Matarhöll/markaður á markaðsleigu (5.500 kr/m²/mán = 159 m.kr/ár)</option></select>
<label>D · arðsemiskrafa fjárfesta, álag á kostnað <span class="v" id="v_D"></span></label><input type="range" id="D" min="0" max="0.30" value="0.20" step="0.01">
<label>Ávöxtunarkrafa kaupanda á NOI <span class="v" id="v_y"></span></label><input type="range" id="y" min="0.05" max="0.08" value="0.0675" step="0.0025">
<p class="src" style="margin-top:14px">Sjálfgefið: B-flokkur, 123 herbergi, Hyatt-kostnaður, ODDSSON-leiga 25%, Kolaportið óbreytt, D 20%, krafa 6,75%.</p>
</div>
<div><table class="t wf" id="wf"></table><div style="margin-top:14px">Rétt verð fyrir húsið eins og það er: <span class="res" id="res"></span> <span class="src" id="res2"></span></div>
<div class="callout" id="verdict"></div></div></div></section>

<section id="s9"><div class="k">9</div><h2>Sviðsmyndir</h2><p class="lead">123 herbergi, B- og C-flokkur, leiga eftir USALI (65%) eða ODDSSON (25%), kostnaður á Hyatt-stigi eða 20% undir, Kolaport eða matarhöll, þróunaraðili (D 20%, 6,75%) eða langtímaeigandi (D 10%, 6,0%). m.kr.</p>
{tbl(head_scen, sel)}
<div class="callout"><b>Lesturinn.</b> Á Hyatt-kostnaði er verðið neikvætt í öllum B-sviðsmyndum og jákvætt í C aðeins með langtímaeiganda. Með kostnað 20% undir Hyatt og ODDSSON-leigu verður B rétt yfir núlli með langtímaeiganda og matarhöll; C nær 1,1–2,6 ma.kr. Íbúðir ofan á húsið og hótel ofan á húsið bæta ekki verðið á núverandi forsendum (VSK af íbúðabyggingu fæst ekki endurgreiddur, nýbygging ofan á 894 þ.kr/m²). Full tafla með 89 og 106 herbergjum og A-flokki er í líkaninu.</div></section>

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
<li>Umbreyting 20% undir Hyatt-raunkostnaði, staðfest með verktakatilboðum í COWI-liðina.</li><li>Rekstraraðili sem skrifar undir 25 ára samning á 25% af tekjum með grunnleigu og móðurfélagsábyrgð.</li><li>Jarðhæðin á markaðsleigu, með borginni, ekki gegn henni.</li><li>Langtímaeigandi sem kaupir fullbúna eign á 6,0% kröfu, samið fyrirfram.</li></ul></div>
</div></section>

<section id="s11"><div class="k">11</div><h2>Næstu skref og heimildir</h2>
<div class="grid g2"><div><ul class="tight">
<li><b>Fundur með Jóni og Daníel:</b> sýna bilið milli þess sem hótelið stendur undir og þess sem ríkið ætlast til; fá hlutverk Íslandshótela skjalfest.</li>
<li><b>Íslandshótel:</b> spyrja beint um verðflokk, leiguhlutfall og form samnings.</li>
<li><b>Kostnaðarpróf:</b> GT, Sveinn og Ellert meta hvort 20% undir Hyatt er raunhæft; verktakatilboð í frárennsli, glugga og asbest.</li>
<li><b>Skipulag:</b> Kvosardeiliskipulag með 2015-breytingunni (hlutfall gististaða) og afstaða borgarinnar til viðbótar ofan á húsið.</li>
<li><b>TORG:</b> tilboðsreglur, Kolaportssamningur ríkis og borgar, lóðarleigusamningur 1967/1993, hvort fjárlagaheimild sé í gildi.</li>
<li><b>Langtímaeigandi:</b> þreifa á fasteignafélagi eða lífeyrissjóði um framvirk kaup.</li>
<li><b>Tilboðsbréf og skilmálar:</b> drög þegar A og C hafa verið fest.</li></ul></div>
<div><ul class="tight" style="font-size:12px">
<li>Söluyfirlit TORG 31.8.2026; fasteigna- og veðbandayfirlit HMS júlí 2026.</li><li>COWI: ástandsskýrsla 3210316-000-CRP-0003 (16.10.2024) með burðarþolsminnisblaði 6.9.2024; innivistarskoðun A290951-001 (20.8.2025).</li><li>Aðaluppdrættir BN045618 (2013, 1:200) og BN048223 (2014), skjalasafn Reykjavíkur.</li><li>Umsögn skipulagsfulltrúa 4.4.2024; minnisblað ríkis og borgar um LHÍ 20.4.2022; Stjórnarráðið 11.9.2024.</li><li>Hyatt Centric L176: EAC júlí 2026 (isfast Drive, Hyatt - Reitir/Framkvæmd/Kostnaðaráætlun).</li><li>Íslandshótel hf.: útgefandalýsing 13.5.2024, fjárfestakynning maí 2024, ársreikningur 2024; Reitir tilkynning 29.4.2026.</li><li>ÍF/RR hótel: leigusamningur ODDSSON 2020 með viðauka 2021; áætlun 06/2020.</li><li>Flóra/ÍF: Black Dunes Þorlákshöfn rekstraráætlun 17.3.2025, BBR módel v1 2026.</li><li>KPMG: Hótelgeirinn á Íslandi 2014 (Ferðamálastofa). Hagstofa Íslands: gistinætur 2025–26. HVS: hotel lease structures.</li><li>Hamranes fjárhagslíkan v2.2 og greining á ávöxtunarkröfu fasteignafélaga 19.8.2026.</li><li>Gögn Jóns Hauks 22.9.2026: Næsland-kynning, stofnkostnaðar- og rekstraráætlun, samantekt COWI-flagga.</li></ul></div></div>
<p class="src" style="margin-top:20px">Líkan: model.py / model2.py / model3.py og Excel v0.1 í Claude Projects/Tollhúsið Næsland/. Þessi síða er drög 1, 22.9.2026, og verður uppfærð þegar A og C hafa verið fest.</p></section>
</main>
<script>
"use strict";
const TIER={{A:{{adr:36000,occ:.72,fb:.20,other:.03,innan:210,ffe:3}},B:{{adr:48000,occ:.72,fb:.35,other:.05,innan:254,ffe:5}},C:{{adr:65000,occ:.68,fb:.45,other:.08,innan:320,ffe:8}}}};
const A_CONV=6879.7,A_BOH=861.5,KOLA=2409.6;
function capex(t,keys,cf,ground){{const r=45+25+30+78+100+118+8+TIER[t].innan;const conv=A_CONV*r/1000,boh=A_BOH*200/1000,lumps=420,g=ground==='matarholl'?KOLA*220/1000:0;const c=conv+boh+lumps+g,unc=.10*c,soft=.13*(c+unc),fee=.03*(c+unc),ffe=keys*TIER[t].ffe;return {{total:(c+unc+soft+fee+ffe)*cf,conv,boh,lumps,g,unc,soft,fee,ffe}};}}
function rev(t,keys){{const T=TIER[t];const rooms=keys*T.adr*T.occ*365/1e6;return {{rooms,total:rooms*(1+T.fb+T.other)}};}}
function calc(){{const t=tier.value,k=+keys.value,c=+cf.value,C=+Cc.value,g=ground.value,D=+Dd.value,y=+yy.value;
 const X=capex(t,k,c,g),Rv=rev(t,k),gr=g==='kolaport'?45.3:KOLA*5500*12/1e6;const leiga=C*Rv.total+gr;const fixed=(85.48+7.83)*1.25+5.5*1.5+X.total*.005;const noi=leiga-fixed-leiga*.01;const V=noi/y;const FA=X.total*.7*.102*1.0;const Hh=(85.48+7.83+5.53+10)*3-45.3*2;const K=4725.5*.016+20;const price=(V/(1+D)-X.total-FA-Hh-K)/(1+.7*.102*4.5);
 const f=x=>x.toLocaleString('de-DE',{{maximumFractionDigits:0}});
 v_tier.textContent=t;v_keys.textContent=k;v_cf.textContent=Math.round(c*100)+'% af Hyatt';v_C.textContent=(C*100).toFixed(1).replace('.',',')+'%';v_ground.textContent=g==='kolaport'?'45 m.kr':'159 m.kr';v_D.textContent=Math.round(D*100)+'%';v_y.textContent=(y*100).toFixed(2).replace('.',',')+'%';
 const rows=[['A · framkvæmd án VSK (m.kr/herb '+f(X.total/k)+')',-X.total],['B · tekjur hótels (herbergi '+f(Rv.rooms)+')',Rv.total],['Leiga hótels + jarðhæð ('+f(gr)+')',leiga],['Fastur eigandakostnaður og umsýsla',-(fixed+leiga*.01)],['NOI eiganda',noi],['Verðmæti = NOI / '+(y*100).toFixed(2).replace('.',',')+'%',V],['÷ (1 + D) → verðmæti að frádreginni arðsemiskröfu',V/(1+D)],['− A',-X.total],['− fjármagn á A',-FA],['− biðtími 3 ár (fasteignagjöld o.fl. − Kolaport)',-Hh],['− stimpilgjald og kaupkostnaður',-K],['÷ (1 + fjármagn á kaupverð 4,5 ár)','']];
 wf.innerHTML=rows.map(r=>'<tr><td>'+r[0]+'</td><td>'+(r[1]===''?'':(r[1]<0?'<span class="neg">':'')+f(r[1])+(r[1]<0?'</span>':''))+'</td></tr>').join('');
 res.textContent=f(price)+' m.kr';res.className='res'+(price<0?' neg':'');res2.textContent=' · '+f(price/k)+' m.kr á herbergi · '+f(price/10150.8*1000)+' þ.kr/m² núverandi húss';
 verdict.innerHTML=price<0?'<b>Neikvætt:</b> við þessar forsendur ber verkefnið ekki kaupverð; það vantar '+f(-price)+' m.kr upp á þótt húsið væri gefið. Færðu C upp, kostnaðarstigið niður eða D niður til að sjá hvað þarf að vera satt.':'<b>Jákvætt:</b> þetta er hæsta verð fyrir húsið eins og það er sem skilar fjárfestum D-álaginu að gefnum A, B og C. Berðu saman við viðmið seljanda: eigið mat 2022 ≥2.000, fasteignamat 4.725.';}}
const tier=document.getElementById('tier'),keys=document.getElementById('keys'),cf=document.getElementById('cf'),Cc=document.getElementById('C'),ground=document.getElementById('ground'),Dd=document.getElementById('D'),yy=document.getElementById('y');
[tier,keys,cf,Cc,ground,Dd,yy].forEach(e=>e.addEventListener('input',calc));calc();
const secs=[...document.querySelectorAll('section')],links=[...document.querySelectorAll('nav a')];
document.addEventListener('keydown',e=>{{if(e.target.tagName==='INPUT'||e.target.tagName==='SELECT')return;const y=window.scrollY+10;let i=secs.findIndex(s=>s.offsetTop>y);if(i<0)i=secs.length;if(e.key==='ArrowRight'||e.key==='PageDown'||e.key===' '){{e.preventDefault();(secs[Math.min(i,secs.length-1)]||secs[0]).scrollIntoView();}}if(e.key==='ArrowLeft'||e.key==='PageUp'){{e.preventDefault();(secs[Math.max(i-2,0)]).scrollIntoView();}}}});
const io=new IntersectionObserver(es=>es.forEach(en=>{{if(en.isIntersecting){{links.forEach(l=>l.classList.toggle('on',l.getAttribute('href')==='#'+en.target.id));}}}}),{{threshold:.4}});secs.forEach(s=>io.observe(s));
</script></body></html>'''
open(os.path.join(HERE, 'app.html'), 'w', encoding='utf-8').write(HTML)
print('app.html', len(HTML.encode()) // 1000, 'kB')

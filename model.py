# Tollhúsið – Næsland: kostnaðar-, rekstrar- og verðgetulíkan v0.1 (22.9.2026)
# Allar fjárhæðir í m.kr án VSK nema annað sé tekið fram. Verðlag 2026.
import math, copy

# ---------------------------------------------------------------- Eignin
BUILDING = dict(
    A_total=10150.8,        # m² brúttó skv. fasteignaskrá
    A_kolaport=2410.0,      # 01-0105 vörugeymsla 1.417,1 + 01-0106 "hugsað fyrir Kolaport" 992,5
    A_kjallari=710.4,
    fmat=4725.5,            # fasteignamat 2026
    fmat_2027=4914.9,
    lodmat=987.35,
    fasteignagjold=85.48,   # álögð 2026 skv. söluyfirliti
    vatn_fraveita=7.83,
    brunatr=5.53,
    stimpil=0.016,          # lögaðili
)

# ---------------------------------------------------------------- Sviðsmyndir (gæðaflokkur)
# ADR án VSK (11% VSK á gistingu), stöðug nýting, F&B og annað sem hlutfall af herbergistekjum,
# EBITDA fyrir leigu/FF&E-sjóð sem hlutfall af heildartekjum, innréttingar á herbergi (hard + FF&E).
TIERS = {
    'A': dict(name='A – Næsland eins og kynnt (upper midscale / lifestyle)', keys=100, adr=36000, occ=0.72,
              fb=0.20, other=0.03, ebitda=0.32, fit_hard=9.0, ffe=3.0, public=200.0, key_m2=52),
    'B': dict(name='B – Upper upscale lifestyle (Konsulat / Parliament-flokkur)', keys=92, adr=48000, occ=0.72,
              fb=0.35, other=0.05, ebitda=0.33, fit_hard=14.0, ffe=5.0, public=400.0, key_m2=58),
    'C': dict(name='C – Lúxus / boutique (EDITION-flokkur, minna hús)', keys=80, adr=65000, occ=0.68,
              fb=0.45, other=0.08, ebitda=0.34, fit_hard=22.0, ffe=8.0, public=750.0, key_m2=66),
}

# ---------------------------------------------------------------- Almennar forsendur
BASE = dict(
    tier='B',
    extra_floor='none',     # 'none' | 'top' (6. hæð ofan á 5. hæð, 1.560 m²) | 'full' (auk hæðar yfir plötu 3. hæðar/Kolaportsþaki, +2.900 m²) – háð DSK
    extra_m2=dict(none=0.0, top=1560.0, full=4460.0),
    extra_floor_rate=520.0, # þ.kr/m² nýbygging ofan á, fullbúið án innréttinga herbergja
    extra_keys_per_m2=1/58.0,
    kolaport_rent=45.3,     # 3,777 m.kr/mán lágmark í auglýsingu RVK júní 2025 (án VSK)
    kolaport_during_works=0.5,
    # Framkvæmd – byggingarhluti, þ.kr/m² af hótelhluta nema "heild" (m.kr)
    demo_rate=18.0, hazmat=180.0, sewer=120.0, windows_m2=1600.0, windows_rate=220.0, single_windows=45.0,
    facade=230.0, roof_garden=180.0, structure=150.0, lifts=90.0,
    plumb_rate=28.0, hvac_rate=30.0, elec_rate=40.0, sprinkler_rate=14.0, kolaport_fix=120.0,
    unc_building=0.12, unc_fitout=0.08,
    design=0.08, permits=0.01, dev_fee=0.05,       # hlutföll af hard cost
    # Tímalína
    buy_year=2027, works_years=(2028, 2029), open_year=2030, ramp=(0.80, 0.92, 1.0),
    # Fjármögnun
    ltc=0.70, cl_rate=0.102, cl_fee=0.002, infl=0.035,
    yld=0.0675, ltv_refi=0.65, bond_real=0.0286+0.010, bond_years=25, bond_cost=0.005,
    hold_years=10, sale_cost=0.01,
    cost_factor=1.0,        # 0,8 = bjartsýnt, 1,0 = grunnur, 1,2 = varfærið
    rent_share_override=None,  # ef sett: leiga = hlutfall af heildartekjum (JHB ≈ 0,30 af tekjum m. F&B)
    # Rekstur eiganda
    ffe_reserve=0.04, operator_min=0.08,   # rekstraraðili heldur ≥8% af tekjum eftir leigu
    ins=5.5, mgmt=0.01, maint=0.005,       # maint = hlutfall af brunabótamati-ígildi ≈ capex
)

def hotel_area(p):
    A = BUILDING['A_total'] - BUILDING['A_kolaport']
    return A + p['extra_m2'][p['extra_floor']]

def keys(p):
    t = TIERS[p['tier']]; k = t['keys']
    return k + round(p['extra_m2'][p['extra_floor']] / (t['key_m2']*1.2))

def capex(p):
    t = TIERS[p['tier']]; A = BUILDING['A_total'] - BUILDING['A_kolaport']
    L = []
    def add(g, name, qty, unit, rate, src):
        amt = qty*rate/1000.0 if unit != 'heild' else rate
        L.append(dict(g=g, name=name, qty=qty, unit=unit, rate=rate, amt=amt, src=src))
    add('Hús', 'Niðurrif innanhúss, hreinsun og förgun (hótelhluti)', A, 'm²', p['demo_rate'], 'Mat ÍF; COWI 2025: gólfefni, klæðningar, milliveggir og lagnir að mestu upprunaleg')
    add('Hús', 'Asbest, PCB, þungmálmar og mygla – afmengun með leyfishöfum', 1, 'heild', p['hazmat'], 'COWI 2024 viðauki B: 550–600 m² asbestplötur; COWI 2025: asbest í gólfflísum, mygla á öllum hæðum nema 5.')
    add('Hús', 'Frárennslislagnir í grunni undir gólfplötu + tengingar', 1, 'heild', p['sewer'], 'COWI 2025 forgangsatriði 1: pottlagnir mjög ryðgaðar/bólgnar, myndavél komst ekki áfram')
    add('Hús', 'Gluggaveggjaeiningar 4.–5. hæð og langhliðar, ný kerfi', p['windows_m2'], 'm²', p['windows_rate'], 'COWI 2024 magnskrá 1.600 m²; einingarverð mat ÍF (álkerfi, gler, förgun asbestplatna sér)')
    add('Hús', 'Stakir gluggar og hurðir í steyptum veggjum', 1, 'heild', p['single_windows'], 'COWI 2024: gaflar 1.–2. hæð og 4.–5. hæð')
    add('Hús', 'Ytra byrði: klæðning steyptra flata, múr, handrið, mósaík/gabbró', 1, 'heild', p['facade'], 'COWI 2024: klæðning 190 m.kr (án glugga) + endurgerð gabbróflísa við listaverk')
    add('Hús', 'Þak og þakgarður á 3. hæð (dúkur, ílögn, burðarstyrking, garður)', 1, 'heild', p['roof_garden'] if p['extra_floor']!='full' else 60.0, 'COWI minnisblað 6.9.2024: plata þolir eina hæð/veitingaálag; sundlaug/padel ekki inni')
    add('Hús', 'Burðarvirki: op, lyftuop, nýir stigar, styrkingar', 1, 'heild', p['structure'], 'Mat ÍF; steypa >35 MPa skv. COWI')
    add('Hús', 'Lyftur, 3 stk', 1, 'heild', p['lifts'], 'Mat ÍF; núverandi lyfta gömul og bilanagjörn skv. COWI 2025')
    add('Hús', 'Pípulagnir (hita-, neyslu- og frárennslislagnir innanhúss)', A, 'm²', p['plumb_rate'], 'Hamranes EAC kafli 3 24,8 þ/m² nýbygging; hærra í umbreytingu')
    add('Hús', 'Loftræsting (nýtt kerfi, uppsetning keyptrar samstæðu ekki nóg fyrir hótel)', A, 'm²', p['hvac_rate'], 'COWI 2025: miðrými án loftræstingar, stýringar úreltar')
    add('Hús', 'Raflagnir, lýsing, öryggis- og hússtjórnarkerfi', A, 'm²', p['elec_rate'], 'Hamranes EAC kafli 4 49,5 þ/m²')
    add('Hús', 'Vatnsúðakerfi', A, 'm²', p['sprinkler_rate'], 'JHB 12,8 þ/m²; mat ÍF')
    add('Hús', 'Kolaportið: leki, loftaklæðning, salerni, hringhurð, lýsing', 1, 'heild', p['kolaport_fix'], 'COWI 2025 kafli 2.7')
    if p['extra_floor'] != 'none':
        add('Hús', 'Viðbótarbyggingarmagn ofan á hús (burður, klæðning, lagnir, án innréttinga herbergja)', p['extra_m2'][p['extra_floor']], 'm²', p['extra_floor_rate'], 'COWI 6.9.2024: ein hæð í lagi; umsögn skipulagsfulltrúa 2024 útilokaði ekki eina hæð')
    K = keys(p)
    add('Herbergi', f'Innréttingar herbergja, hard fit-out ({K} herb.)', K, 'herb', t['fit_hard']*1000, f'Flokkur {p["tier"]}: {t["fit_hard"]} m.kr/herb (veggir, böð, hurðir, yfirborð, sérlagnir)')
    add('Herbergi', 'Laus búnaður herbergja (FF&E)', K, 'herb', t['ffe']*1000, f'Flokkur {p["tier"]}: {t["ffe"]} m.kr/herb; JHB 2,4–3,2')
    add('Almenn rými', 'Lobby, veitingastaður, bar, eldhús, spa/gym, fundarrými, þakgarðsbygging', 1, 'heild', t['public'], f'Flokkur {p["tier"]}')
    building = sum(l['amt'] for l in L if l['g']=='Hús')
    fitout = sum(l['amt'] for l in L if l['g']!='Hús')
    unc = building*p['unc_building'] + fitout*p['unc_fitout']
    hard = building + fitout + unc
    design = hard*p['design']; permits = hard*p['permits']; fee = hard*p['dev_fee']
    total = (hard + design + permits + fee)*p['cost_factor']
    return dict(lines=L, building=building, fitout=fitout, unc=unc, hard=hard, design=design, permits=permits, fee=fee, total=total, keys=K, per_key=total/K, per_m2=total/hotel_area(p)*1000)

def operations(p):
    t = TIERS[p['tier']]; K = keys(p)
    rooms = K*t['adr']*t['occ']*365/1e6
    fb = rooms*t['fb']; other = rooms*t['other']; rev = rooms+fb+other
    ebitda = rev*t['ebitda']; reserve = rev*p['ffe_reserve']; op_keep = rev*p['operator_min']
    rent = ebitda - reserve - op_keep
    if p.get('rent_share_override'): rent = rev*p['rent_share_override']
    rent_share = rent/rev
    return dict(keys=K, rooms=rooms, fb=fb, other=other, rev=rev, ebitda=ebitda, reserve=reserve, op_keep=op_keep, rent=rent, rent_share=rent_share, revpar=t['adr']*t['occ'], rev_per_key=rev/K)

def owner_fixed(p, C):
    # fastur eigandakostnaður á verðlagi 2026: fasteignagjöld (+25% eftir endurmat), tryggingar, viðhaldssjóður
    return (BUILDING['fasteignagjold']+BUILDING['vatn_fraveita'])*1.25 + p['ins']*1.5 + C['total']*p['maint']

def owner_noi(p, rent_total, C, ix=1.0):
    return rent_total - owner_fixed(p, C)*ix - rent_total*p['mgmt']

def dcf(p, price):
    C = capex(p); O = operations(p)
    infl = p['infl']; y0 = p['buy_year']; N = p['hold_years']
    years = list(range(y0, y0+N+1)); idx = {y: (1+infl)**(y-y0) for y in years}
    open_y = p['open_year']; w = p['works_years']; split = (0.40, 0.60)
    stab_rent = O['rent'] + p['kolaport_rent']
    purchase = {y: 0.0 for y in years}; purchase[y0] = price + BUILDING['fmat']*BUILDING['stimpil'] + 20.0
    cap = {y: 0.0 for y in years}
    for i, y in enumerate(w): cap[y] = C['total']*split[i]*idx[y]
    noi = {}; rent = {}; kola = {}
    for y in years:
        if y < open_y:
            hold = ((BUILDING['fasteignagjold']+BUILDING['vatn_fraveita']) + BUILDING['brunatr'] + 10.0)*idx[y]
            kola[y] = p['kolaport_rent']*(p['kolaport_during_works'] if y in w else 1.0)*idx[y]
            rent[y] = 0.0; noi[y] = kola[y] - hold
        else:
            r = p['ramp'][min(y-open_y, len(p['ramp'])-1)]
            rent[y] = O['rent']*r*idx[y]; kola[y] = p['kolaport_rent']*idx[y]
            noi[y] = owner_noi(p, rent[y]+kola[y], C, idx[y])
    loan = 0.0; equity = {y: 0.0 for y in years}; cl_bal = {}
    for y in years:
        if y <= w[-1]:
            need = purchase[y] + cap[y] - noi[y]
            draw = max(need, 0)*p['ltc']; loan = loan*(1+p['cl_rate']) + draw*(1+p['cl_fee'])
            equity[y] = -(max(need, 0) - draw)
        cl_bal[y] = loan
    stab_y = open_y + len(p['ramp'])-1
    noi_stab = owner_noi(p, stab_rent, C, idx[stab_y])*idx[stab_y]/idx[stab_y]
    noi_stab = stab_rent*idx[stab_y] - owner_fixed(p, C)*idx[stab_y] - stab_rent*idx[stab_y]*p['mgmt']
    value_stab = noi_stab/p['yld']; bond = value_stab*p['ltv_refi']
    r_nom = (1+p['bond_real'])*(1+infl)-1
    ann = bond*r_nom/(1-(1+r_nom)**-p['bond_years'])
    cf = {y: 0.0 for y in years}; bal = bond
    for y in years:
        if y <= w[-1]:
            cf[y] = equity[y]
        elif y < stab_y:
            loan = loan*(1+p['cl_rate']); cf[y] = noi[y]
        elif y == stab_y:
            cf[y] = noi[y] + bond*(1-p['bond_cost']) - loan; loan = 0.0
            interest = bal*r_nom; bal -= (ann-interest); cf[y] -= ann
        else:
            interest = bal*r_nom; bal -= (ann-interest); cf[y] = noi[y] - ann
    ye = years[-1]
    noi_next = (stab_rent*idx[ye]*(1+infl)) - owner_fixed(p, C)*idx[ye]*(1+infl) - stab_rent*idx[ye]*(1+infl)*p['mgmt']
    exit_v = noi_next/p['yld']*(1-p['sale_cost'])
    cf[ye] += exit_v - bal
    flows = [cf[y] for y in years]
    return dict(years=years, flows=flows, irr=irr(flows), C=C, O=O, noi=noi, value_stab=value_stab, noi_stab=noi_stab, bond=bond,
                cl_peak=max(cl_bal.values()), equity_in=-sum(v for v in equity.values() if v < 0), exit_v=exit_v, stab_rent=stab_rent,
                purchase=purchase, cap=cap, ann=ann, r_nom=r_nom, stab_y=stab_y)

def irr(cfs):
    f = lambda r: sum(cf/(1+r)**t for t, cf in enumerate(cfs))
    a, b = -0.95, 2.0
    if f(a)*f(b) > 0: return float('nan')
    for _ in range(200):
        m = (a+b)/2
        if f(a)*f(m) <= 0: b = m
        else: a = m
    return (a+b)/2

def price_for_irr(p, target):
    lo, hi = -4000.0, 8000.0
    for _ in range(80):
        m = (lo+hi)/2
        r = dcf(p, m)['irr']
        if r != r: hi = m; continue
        if r > target: lo = m
        else: hi = m
    return (lo+hi)/2

def scenarios():
    out = []
    for tier in 'ABC':
        for ef in ('none','top','full'):
            p = dict(BASE); p['tier'] = tier; p['extra_floor'] = ef
            C = capex(p); O = operations(p)
            row = dict(tier=tier, extra=ef, keys=O['keys'], capex=C['total'], per_key=C['per_key'], rev=O['rev'], rent=O['rent'], rent_share=O['rent_share'],
                       stab_rent=O['rent']+p['kolaport_rent'])
            for tgt in (0.10, 0.12, 0.15):
                row[f'p{int(tgt*100)}'] = price_for_irr(p, tgt)
            d = dcf(p, max(row['p12'], 0))
            row['value_stab'] = d['value_stab']; row['noi_stab'] = d['noi_stab']
            out.append(row)
    return out

if __name__ == '__main__':
    for r in scenarios():
        print(f"{r['tier']} extra={r['extra']:4s} keys={r['keys']:3d} capex={r['capex']:6.0f} ({r['per_key']:.0f}/key) rev={r['rev']:5.0f} rent={r['rent']:4.0f} ({r['rent_share']*100:.0f}%) "
              f"NOIstab={r['noi_stab']:4.0f} value={r['value_stab']:5.0f}  P@10%={r['p10']:6.0f} P@12%={r['p12']:6.0f} P@15%={r['p15']:6.0f}")
    p = dict(BASE); C = capex(p)
    for l in C['lines']: print(f"  {l['g']:12s} {l['name'][:70]:70s} {l['amt']:8.1f}")
    print({k: round(v,1) for k, v in C.items() if isinstance(v, float)})
    print(operations(p))

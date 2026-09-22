# Tollhúsið v0.2 – A á Hyatt-köflum, herbergi skv. teikningum, valkostir fyrir jarðhæð og ofanábyggingu.
# m.kr án VSK, verðlag 2026. Heimild fyrir kafla: Hyatt L176 EAC júlí 2026 (isfast Drive), 9.922,9 m², 169 herb.
import sys
sys.path.insert(0, '/Users/villithor/repos/tollhusid')
import model as M1

# ---- Flatarmál skv. fasteignaskrá og aðaluppdráttum BN045618/BN048223 (m² brúttó)
AREA = dict(kjallari=710.4, h1_kolaport=2409.6, h1_hotel=1105.0,   # 1. hæð: afgreiðsla tollstjóra 730 + uppboðssalur 218 + lögregla 133,5 + lagnir/stigar
            h2=1591.4, h3=943.3, h4=1620.0, h5=1620.0, lyftuhus=61.2)
A_TOTAL = 10150.8
A_HOTEL_CONV = AREA['h1_hotel'] + AREA['h2'] + AREA['h3'] + AREA['h4'] + AREA['h5']    # 6.879,7 – breytt í hótel
A_BOH = A_TOTAL - AREA['h1_kolaport'] - A_HOTEL_CONV                                   # kjallari + lyftuhús + afgangur ≈ 861

# ---- Herbergi eftir hæðum, lesið af grunnmyndum 1:200 (BN045618). Suðurálma ~78 m × 15,5 m með tveimur kjörnum;
#      norðurálma 4./5. hæðar ~18 m; 2. hæð: norðurhlið álmunnar liggur inn í Kolaportssalinn → einhliða gangur;
#      3. hæð: inndregin, norðurhlið opnast út á þakflöt (garðherbergi). Breidd herbergja ræður: 3,9 m (27–30 m² nettó) / 4,5 m (33–36 m²).
KEYS = {
    'þétt (3,9 m)':   dict(h5=39, h4=39, h3=28, h2=17, total=123),
    'JHB (4,5 m)':    dict(h5=34, h4=34, h3=24, h2=14, total=106),
    'JHB skipan':     dict(h5=34, h4=34, h3=10, h2=11, total=89),   # lobby á 3. hæð, gym á 2. hæð eins og í kynningu
}

# ---- Hyatt-kaflar án VSK, þ.kr/m² (EAC júlí 2026) og aðlögun fyrir Tollhúsið
HYATT = dict(adstada=50.1, jardvinna=29.6, burdarvirki=90.1, lagnir=78.2, raflagnir=112.7, innanhuss=253.9, utanhuss=118.5, lod=14.9, ovissa=11.9,
             soft=118.75, construction=759.9, total=894.3, keys=169, m2=9922.9)
ADJ = dict(adstada=45.0, jardvinna=25.0, burdarvirki=30.0, lagnir=78.0, raflagnir=100.0, utanhuss=118.0, lod=8.0)   # innanhúss eftir flokki
TIER_INNAN = {'A': 210.0, 'B': 254.0, 'C': 320.0}
TIER_FFE = {'A': 3.0, 'B': 5.0, 'C': 8.0}
COWI_LUMPS = dict(hazmat=180.0, sewer=120.0, kolaport_fix=120.0)
BOH_RATE = 200.0          # þ.kr/m² – kjallari/tæknirými gerð nothæf, ekki innréttuð
UNC = 0.10                # ófyrirséð á framkvæmd (Hyatt endaði í 12% ofan á samninga með aukaverkum)
SOFT = 0.13               # hönnun, byggingarstjórn, eftirlit, umsýsla (Hyatt 15,6% af framkvæmd)
IF_FEE = 0.03
NEWBUILD_RATE = HYATT['total']   # ofanábygging: Hyatt allt-í-allt sem viðmið fyrir nýbyggingu ofan á hús

def capex(tier='B', keys=106, cost_factor=1.0, extra_m2=0.0, extra_use='none', ground='kolaport', detail=False):
    rate = sum(ADJ.values()) + TIER_INNAN[tier]
    conv = A_HOTEL_CONV * rate / 1000.0
    boh = A_BOH * BOH_RATE / 1000.0
    lumps = sum(COWI_LUMPS.values())
    ground_capex = 0.0
    if ground == 'matarholl': ground_capex = AREA['h1_kolaport'] * 220.0 / 1000.0   # skel + lagnir + salerni fyrir matarhöll, rekstraraðili innréttar bása
    construction = conv + boh + lumps + ground_capex
    unc = construction * UNC
    soft = (construction + unc) * SOFT
    fee = (construction + unc) * IF_FEE
    ffe = keys * TIER_FFE[tier]
    extra = 0.0; extra_keys = 0; extra_ffe = 0.0
    if extra_m2 > 0:
        if extra_use == 'hotel':
            extra = extra_m2 * NEWBUILD_RATE / 1000.0
            extra_keys = round(extra_m2 / 70.0)
            extra_ffe = extra_keys * TIER_FFE[tier]
        elif extra_use == 'ibudir':
            extra = extra_m2 * 950.0 / 1000.0 * 1.18   # íbúðir: VSK af byggingarkostnaði fæst ekki endurgreiddur
    total = (construction + unc + soft + fee + ffe) * cost_factor + extra + extra_ffe
    out = dict(rate=rate, conv=conv, boh=boh, lumps=lumps, ground_capex=ground_capex, construction=construction, unc=unc, soft=soft, fee=fee, ffe=ffe,
               extra=extra, extra_keys=extra_keys, extra_ffe=extra_ffe, total=total, keys_total=keys + extra_keys,
               hotel_only=(construction + unc + soft + fee + ffe) * cost_factor - ground_capex * (1 + UNC) * (1 + SOFT + IF_FEE) * cost_factor)
    out['per_key'] = out['hotel_only'] / keys
    return out

def revenue(tier='B', keys=106, occ=None, adr=None):
    t = M1.TIERS[tier]
    adr = adr or t['adr']; occ = occ or t['occ']
    rooms = keys * adr * occ * 365 / 1e6
    rev = rooms * (1 + t['fb'] + t['other'])
    return dict(rooms=rooms, rev=rev, ebitda=rev * t['ebitda'])

def ground_rent(ground='kolaport', rent_m2_man=5500.0):
    if ground == 'kolaport': return 45.3
    return AREA['h1_kolaport'] * rent_m2_man * 12 / 1e6

def rett_verd(tier='B', keys=106, C=0.30, D=0.20, yld=0.0675, cost_factor=1.0, extra_m2=0.0, extra_use='none', ground='kolaport',
              rent_m2_man=5500.0, sale_m2=1350.0, saleable=0.85):
    X = capex(tier, keys, cost_factor, extra_m2, extra_use, ground)
    R = revenue(tier, X['keys_total'] if extra_use == 'hotel' else keys)
    leiga = C * R['rev'] + ground_rent(ground, rent_m2_man)
    fixed = (85.48 + 7.83) * 1.25 + 5.5 * 1.5 + X['total'] * 0.005
    noi = leiga - fixed - leiga * 0.01
    V = noi / yld
    apt = 0.0
    if extra_use == 'ibudir' and extra_m2 > 0:
        apt = extra_m2 * saleable * sale_m2 / 1000.0 * 0.98     # nettó söluverðmæti, 2% sölukostnaður
    A = X['total']
    F_A = A * 0.70 * 0.102 * 1.0
    H = (85.48 + 7.83 + 5.53 + 10) * 3 - 45.3 * 2.0
    K = 4725.5 * 0.016 + 20
    price = ((V + apt) / (1 + D) - A - F_A - H - K) / (1 + 0.70 * 0.102 * 4.5)
    return dict(A=A, per_key=X['per_key'], keys=X['keys_total'], rev=R['rev'], leiga=leiga, noi=noi, V=V, apt=apt, price=price, X=X)

if __name__ == '__main__':
    print(f"Flatarmál: hótelhluti breyttur {A_HOTEL_CONV:.0f} m², BOH {A_BOH:.0f} m², Kolaport {AREA['h1_kolaport']:.0f} m²")
    for tier in 'ABC':
        for name, k in KEYS.items():
            X = capex(tier, k['total'])
            print(f"{tier} {name:14s} herb {k['total']:3d}: A hótel {X['hotel_only']:6.0f} m.kr = {X['per_key']:4.0f} m.kr/herb  (rate {X['rate']:.0f} þ/m², framkv {X['construction']:.0f}, ófyrirs {X['unc']:.0f}, soft {X['soft']:.0f}, þóknun {X['fee']:.0f}, FF&E {X['ffe']:.0f})")
    print()
    print("Rétt verð (m.kr), D=20%, krafa 6,75%:")
    print("tier keys   C    cf   jarðhæð    viðbót           A     leiga  NOI    V     íbúðir  verð")
    for tier in 'BC':
        for keys in (89, 106, 123):
            for C in (0.22, 0.30):
                for cf in (1.0, 0.8):
                    for ground in ('kolaport', 'matarholl'):
                        for (em, eu) in ((0, 'none'), (4460, 'hotel'), (4460, 'ibudir')):
                            r = rett_verd(tier, keys, C, 0.20, 0.0675, cf, em, eu, ground)
                            print(f"{tier}  {keys:3d}  {C:.2f}  {cf:.1f}  {ground:9s}  {em:5.0f} {eu:7s}  {r['A']:6.0f}  {r['leiga']:5.0f}  {r['noi']:5.0f}  {r['V']:6.0f}  {r['apt']:6.0f}  {r['price']:6.0f}")

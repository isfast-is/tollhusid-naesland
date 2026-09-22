# Tollhúsið v0.3 – C (leiga) byggt upp eftir USALI með íslenskum viðmiðum; sviðsmyndir og JSON fyrir mælaborð.
import sys, json
sys.path.insert(0, '/Users/villithor/repos/tollhusid')
import model as M1, model2 as M2

# ---------------- B: tekjuforsendur eftir gæðaflokki (ADR án 11% VSK, stöðug nýting)
TIERS = {
    'A': dict(name='Upper midscale / lifestyle (Næsland eins og kynnt)', adr=36000, occ=0.72, fb=0.20, other=0.03, key_m2=52),
    'B': dict(name='Upper upscale lifestyle (Konsulat / Parliament-flokkur)', adr=48000, occ=0.72, fb=0.35, other=0.05, key_m2=58),
    'C': dict(name='Lúxus / boutique (EDITION-flokkur)', adr=65000, occ=0.68, fb=0.45, other=0.08, key_m2=66),
}
# ---------------- USALI kostnaðaruppbygging (hlutfall af tekjum viðkomandi deildar / heildartekjum)
USALI = dict(
    rooms_cost=0.32,      # laun 18%, sölukostnaður/OTA 8%, þvottur/vörur 6%  (Íslandshótel: laun 42% af heildartekjum 2023)
    fb_cost=0.75,         # veitingadeild: hráefni 30%, laun 38%, annað 7%
    other_cost=0.50,
    undistributed=0.19,   # stjórnun 7%, markaðsmál 4%, upplýsingatækni 1,5%, viðhald/POM 3,5%, orka 3%
    mgmt_fee=0.0,         # 0 = eigin rekstur (Íslandshótel/Næsland); 4% ef alþjóðlegt vörumerki
    ffe=0.04,             # FF&E-sjóður
    landlord_share=0.65,  # fastur leigusamningur: 60–70% af EBITDAR eftir FF&E til leigusala (HVS/alþjóðleg venja)
)
def pl(tier, keys):
    t = TIERS[tier]
    rooms = keys * t['adr'] * t['occ'] * 365 / 1e6
    fb = rooms * t['fb']; other = rooms * t['other']; rev = rooms + fb + other
    dept = rooms * (1 - USALI['rooms_cost']) + fb * (1 - USALI['fb_cost']) + other * (1 - USALI['other_cost'])
    gop = dept - rev * USALI['undistributed']
    ebitdar = gop - rev * USALI['mgmt_fee']
    ffe = rev * USALI['ffe']
    rent_usali = (ebitdar - ffe) * USALI['landlord_share']
    rent_oddsson = max(0.30 * rooms, 0.25 * rev)          # ÍF/RR hótel leigusamningur 2020 (Grensásvegur 16A, 77 herb.)
    rent_jhb = 0.34 * rooms                               # Jón Haukur: 367 af 1.066 herbergistekjum
    return dict(keys=keys, adr=t['adr'], occ=t['occ'], revpar=t['adr'] * t['occ'], rooms=rooms, fb=fb, other=other, rev=rev, rev_key=rev / keys,
                dept=dept, gop=gop, gop_pct=gop / rev, ebitdar=ebitdar, ffe=ffe,
                rent_usali=rent_usali, c_usali=rent_usali / rev, rent_oddsson=rent_oddsson, c_oddsson=rent_oddsson / rev,
                rent_jhb=rent_jhb, c_jhb=rent_jhb / rev, op_after_usali=(ebitdar - ffe - rent_usali) / rev, op_after_oddsson=(ebitdar - ffe - rent_oddsson) / rev,
                noi_key_usali=rent_usali / keys, noi_key_oddsson=rent_oddsson / keys)

# ---------------- Viðmið (heimildir í mælaborði)
BENCH = [
    dict(name='ODDSSON – leigusamningur ÍF/RR hótel 2020 (77 herb., Grensásvegur 16A)', value='30% af herbergistekjum eða 25% af heildartekjum, hvort hærra er; leigusali greiðir fasteignagjöld og tryggingar', src='isfast Drive: Grensásvegur 16a/Leigusamningar/RR hótel'),
    dict(name='Flóra – Black Dunes Þorlákshöfn rekstraráætlun 17.3.2025 (120 herb.)', value='Föst leiga 3,6 m.kr/herb/ár = 29–30% af tekjum; ADR 29–32 þ.kr, nýting 67–74%; EBITDA eftir leigu aðeins 2–4,5%', src='isfast Drive: Ölfus/BBR/Black Beach Resort Collaboration'),
    dict(name='BBR módel v1 (ÍF, apríl 2026)', value='Leiga = 12% af byggingarkostnaði; 25,3 m.kr/herb capex; rekstraraðili heldur 8,7%', src='isfast Drive: Ölfus/BBR/2026 Verkefnavinna'),
    dict(name='Íslandshótel hf. – útgefandalýsing maí 2024', value='ADR 25,4 þ.kr, nýting 68%, RevPAR 17,2 þ.kr (2023, öll keðjan); laun 42% af tekjum; EBITDA fyrir leigu ~30%; Fosshótel Austfirðir: grunnleiga + 10% af brúttósölu', src='cdn.islandsbanki.is IH-Utgefandalysing'),
    dict(name='Íslandshótel hf. – ársreikningur 2024', value='Tekjur 16.669 m.kr, EBITDA 4.416 (26,5%, IFRS 16 = fyrir leigu); 1.966 herbergi → 8,5 m.kr tekjur á herbergi', src='GlobeNewswire 4.4.2025'),
    dict(name='Reitir – nýr leigusamningur við Íslandshótel 1.9.2026', value='470 herbergi, 26.500 m², NOI +720 m.kr/ár = 1,53 m.kr á herbergi = 27 þ.kr/m²/ár', src='GlobeNewswire 29.4.2026'),
    dict(name='KPMG – Hótelgeirinn á Íslandi 2014', value='78,6% hótelrekstraraðila á höfuðborgarsvæðinu leigja húsnæðið; afkoma á höfuðborgarsvæði lakari en á landsbyggð, m.a. vegna húsnæðiskostnaðar', src='ferdamalastofa.is'),
    dict(name='Hagstofa Íslands 2025', value='Herbergjanýting höfuðborgarsvæðis 74,7% (2025); 58 hótel / 5.555 herbergi', src='hagstofa.is'),
    dict(name='Alþjóðleg venja (HVS, hótelleigusamningar)', value='Fastur leigusamningur = 60–70% af EBITDAR til leigusala; blandaður samningur = grunnleiga + 8–12% af tekjum', src='HVS: Hotel contracts – to lease or not to lease'),
]

def scenarios():
    out = []
    for tier in 'ABC':
        for keys in (89, 106, 123):
            P = pl(tier, keys)
            for cf in (1.0, 0.8):
                for cname, C in (('USALI 65%', P['c_usali']), ('ODDSSON', P['c_oddsson']), ('JHB 34% herb.', P['c_jhb'])):
                    for ground in ('kolaport', 'matarholl'):
                        for D, y in ((0.20, 0.0675), (0.10, 0.06)):
                            r = M2.rett_verd(tier, keys, C, D, y, cf, 0.0, 'none', ground)
                            out.append(dict(tier=tier, keys=keys, cf=cf, cname=cname, C=C, ground=ground, D=D, yld=y, A=r['A'], per_key=r['A'] / keys,
                                            rev=r['rev'], leiga=r['leiga'], noi=r['noi'], V=r['V'], price=r['price']))
    return out

if __name__ == '__main__':
    for tier in 'ABC':
        P = pl(tier, 123)
        print(f"{tier}: tekjur {P['rev']:.0f} (herb {P['rooms']:.0f}, F&B {P['fb']:.0f}), GOP {P['gop']:.0f} ({P['gop_pct']*100:.0f}%), leiga USALI {P['rent_usali']:.0f} = {P['c_usali']*100:.1f}% | ODDSSON {P['rent_oddsson']:.0f} = {P['c_oddsson']*100:.1f}% | JHB {P['rent_jhb']:.0f} = {P['c_jhb']*100:.1f}%; rekstraraðili eftir leigu: {P['op_after_usali']*100:.1f}% / {P['op_after_oddsson']*100:.1f}%; leiga/herb {P['noi_key_usali']:.2f}/{P['noi_key_oddsson']:.2f}")
    S = scenarios()
    json.dump(dict(scen=S, bench=BENCH, usali=USALI, tiers=TIERS, pl={t: pl(t, 123) for t in 'ABC'}, keys=M2.KEYS,
                   capex={t: {k: M2.capex(t, v['total']) for k, v in M2.KEYS.items()} for t in 'ABC'}), open('/Users/villithor/repos/tollhusid/results3.json', 'w'), ensure_ascii=False, indent=1, default=float)
    print()
    print("tier keys cf   C-viðmið      C     jarðhæð   D    y     A     leiga  NOI   V      verð")
    for r in S:
        if r['keys'] == 123 and r['tier'] in 'BC':
            print(f"{r['tier']}  {r['keys']}  {r['cf']:.1f}  {r['cname']:13s} {r['C']*100:4.1f}%  {r['ground']:9s} {r['D']:.2f} {r['yld']:.4f} {r['A']:6.0f} {r['leiga']:5.0f} {r['noi']:5.0f} {r['V']:6.0f} {r['price']:6.0f}")

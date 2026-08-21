"""
CAS LOSS RESERVE DATABASE - indirme ve donusturme

Neden bu veri onemli: CAS'in yayinladigi Schedule P ucgenlerinde ALT
UCGEN DE VAR. Yani her kaza yilinin gercekte ne kadar gelistigi biliniyor.

Bunun anlami su: aracin ust ucgenden yaptigi tahmini, gercekte olanla
karsilastirabiliyoruz. Yuzlerce sirket icin. Sentetik veri "arac dogru
hesapliyor mu" sorusunu cevapliyor; bu veri "tahmin ne kadar isabetli"
sorusunu cevapliyor. Ikisi farkli sorular ve ikisi de gerekli.

Kaynak
------
Casualty Actuarial Society, "Loss Reserving Data Pulled from NAIC
Schedule P with Additional Accident Years".
https://www.casact.org/publications-research/research/research-resources/loss-reserving-data-pulled-naic-schedule-p

CAS'in sarti: bu veriyi kullanan calismalar dosyalara CAS'in web
sayfasindan erisildigini ve verinin CAS adina S&P Global Market
Intelligence tarafindan derlendigini belirtmeli. README'de belirtiliyor.

Veri yapisi
-----------
10 kaza yili (1998-2007) x 10 gelisim yili, alti brans, birkac yuz sirket.
Kolonlar brans son ekiyle geliyor: IncurLoss_B, CumPaidLoss_B ...

  B   ozel oto sorumluluk        C   ticari oto
  D   isci tazminat              F2  tibbi malpraktis (claims made)
  H1  diger sorumluluk           R1  urun sorumlulugu

DONEM ESLEMESI - dikkat
-----------------------
Arac ceyreklik calisiyor, CAS verisi YILLIK. Kaza yili i, ardisik donem
koduna esleniyor (1998 -> 2000Q1, 1999 -> 2000Q2 ...). Etiketler ceyrek
gibi gorunuyor ama bir donem = bir YIL. Aracin aritmetigi yalnizca donem
KODUNU kullandigi icin hesap dogru cikar; sadece rapordaki etiketler
"ceyrek" yazar. Alternatifi ucgende bosluk birakmakti ve o, link
ratio'lari sessizce bozardi.

Calistirma:
    python3 data/cas_indir.py --brans ppauto
    python3 data/cas_indir.py --brans wkcomp --sirket 43
"""

import argparse
import os
import urllib.request

import pandas as pd

CAS_KOK = "https://www.casact.org/sites/default/files/2026-03/"

BRANSLAR = {
    "ppauto":   ("ppauto_pos98-07%20%281%29.csv", "B",  "Ozel oto sorumluluk"),
    "wkcomp":   ("wkcomp_pos_98-07.csv",          "D",  "Isci tazminat"),
    "comauto":  ("comauto_pos_98-07.csv",         "C",  "Ticari oto"),
    "medmal":   ("medmal_pos_98-07.csv",          "F2", "Tibbi malpraktis"),
    "prodliab": ("prodliab_pos_98-07.csv",        "R1", "Urun sorumlulugu"),
    "othliab":  ("othliab_pos_98-07.csv",         "H1", "Diger sorumluluk"),
}

ILK_KAZA_YILI = 1998
GELISIM_YILI = 10
TABAN_DONEM = "2000Q1"        # eslemede kullanilan yapay baslangic


def indir(brans: str, klasor: str) -> str:
    dosya, _, _ = BRANSLAR[brans]
    hedef = os.path.join(klasor, f"cas_{brans}.csv")
    if os.path.isfile(hedef):
        print(f"zaten var: {hedef}")
        return hedef
    url = CAS_KOK + dosya
    print(f"indiriliyor: {url}")
    urllib.request.urlretrieve(url, hedef)
    print(f"kaydedildi: {hedef}")
    return hedef


def donem_etiketi(kaydir: int, taban: str = TABAN_DONEM) -> str:
    yil, ceyrek = int(taban[:4]), int(taban[-1])
    toplam = yil * 4 + (ceyrek - 1) + kaydir
    return f"{toplam // 4}Q{toplam % 4 + 1}"


def donustur(ham: pd.DataFrame, sonek: str, sadece_ust: bool = True) -> pd.DataFrame:
    """CAS ucgenini aracin bekledigi hasar kaydi tablosuna cevirir.

    Aracin semasi:  Odenen satirlari ARTIS, Muallak satiri o gelisim
    noktasindaki ACIK tutar. CAS'te CumPaidLoss kumulatif odenen,
    IncurLoss ise incurred; ikisinin farki acik tutar (vaka rezervi +
    bulk/IBNR). Esleme dogrudan.

    sadece_ust=True ise gozlem penceresi (ust ucgen) verilir; alt ucgen
    dogrulama icin ayrilir.
    """
    ham = ham.rename(columns=lambda c: c.strip())
    kum_odenen = f"CumPaidLoss_{sonek}"
    incurred = f"IncurLoss_{sonek}"
    for k in (kum_odenen, incurred, "AccidentYear", "DevelopmentLag", "GRCODE"):
        if k not in ham.columns:
            raise ValueError(f"Beklenen kolon yok: {k}")

    d = ham.copy()
    d["i"] = d["AccidentYear"] - ILK_KAZA_YILI          # kaza donemi indeksi
    d["d"] = d["DevelopmentLag"].astype(int) - 1        # gelisim indeksi (0 tabanli)
    if sadece_ust:
        d = d[d["i"] + d["d"] <= GELISIM_YILI - 1]

    d = d.sort_values(["GRCODE", "i", "d"])
    d["onceki_odenen"] = d.groupby(["GRCODE", "i"])[kum_odenen].shift(1).fillna(0.0)
    d["odenen_artis"] = d[kum_odenen] - d["onceki_odenen"]
    d["acik"] = d[incurred] - d[kum_odenen]

    ortak = dict(HazAltBrnKod=715, Bedeni=0, Havuz=0)
    satirlar = []
    for _, r in d.iterrows():
        temel = dict(
            HasarDonem=donem_etiketi(int(r["i"])),
            GelDonemHT=int(r["d"]),
            GRCODE=int(r["GRCODE"]),
            **ortak,
        )
        if r["odenen_artis"] != 0:
            satirlar.append({**temel, "HasarDurum": "Odenen",
                             "HasarTutar": float(r["odenen_artis"])})
        if r["acik"] != 0:
            satirlar.append({**temel, "HasarDurum": "Muallak",
                             "HasarTutar": float(r["acik"])})
    return pd.DataFrame(satirlar)


def gercek_ultimate(ham: pd.DataFrame, sonek: str) -> pd.DataFrame:
    """Her (sirket, kaza yili) icin GERCEKLESMIS nihai hasar.

    Alt ucgen bilindigi icin 10. gelisim yilindaki incurred, o kaza
    yilinin fiili ultimate'i sayilir. Tahminin karsisina konacak sayi bu.
    """
    ham = ham.rename(columns=lambda c: c.strip())
    son = ham[ham["DevelopmentLag"] == GELISIM_YILI]
    return (son[["GRCODE", "AccidentYear", f"IncurLoss_{sonek}"]]
            .rename(columns={f"IncurLoss_{sonek}": "GercekUltimate"})
            .reset_index(drop=True))


def main() -> None:
    ap = argparse.ArgumentParser(description="CAS loss reserve verisini indir ve cevir")
    ap.add_argument("--brans", default="ppauto", choices=sorted(BRANSLAR))
    ap.add_argument("--sirket", type=int, default=None,
                    help="Tek bir GRCODE; verilmezse en cok primli sirket")
    args = ap.parse_args()

    klasor = os.path.dirname(os.path.abspath(__file__))
    yol = indir(args.brans, klasor)
    ham = pd.read_csv(yol)
    ham = ham.rename(columns=lambda c: c.strip())
    _, sonek, ad = BRANSLAR[args.brans]

    print(f"\n{ad} · {len(ham):,} satir · {ham['GRCODE'].nunique()} sirket")

    kod = args.sirket
    if kod is None:
        prim = f"EarnedPremNet_{sonek}"
        kod = int(ham.groupby("GRCODE")[prim].sum().idxmax())
    sirket = ham[ham["GRCODE"] == kod]
    print(f"secilen sirket: {kod} - {sirket['GRNAME'].iloc[0]}")

    kayit = donustur(sirket, sonek, sadece_ust=True)
    hedef = os.path.join(klasor, f"cas_{args.brans}_{kod}.xlsx")
    kayit.drop(columns=["GRCODE"]).to_excel(hedef, index=False)
    print(f"{len(kayit):,} kayit -> {hedef}")

    ger = gercek_ultimate(sirket, sonek)
    print("\nGercek ultimate (alt ucgenden):")
    print(ger.to_string(index=False))
    print("\nAraci calistir:")
    print(f"  set IBNR_GIRIS_DOSYASI={hedef}")
    print("  python3 ibnr.py")


if __name__ == "__main__":
    main()

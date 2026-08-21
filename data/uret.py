"""
SENTETIK HASAR VERISI URETECI

Iki isi birden yapar ve ikincisi daha degerli:

  1. Depoya gercek portfoy verisi koymadan aracin calisabilmesini saglar.
  2. Gelisim deseni BILINDIGI icin aracin dogru sonuc uretip uretmedigi
     olculebilir hale gelir. "Bekledigimiz faktorler sunlardi, arac
     sunlari buldu" cumlesi kurulabiliyorsa arac test edilebilir demektir.

Kurgu: her kaza doneminin nihai hasari (ultimate) rastgele secilir, ama
gelisim yuzdeleri TUM SATIRLAR ICIN AYNI kalir. Boylece incurred ucgenin
link ratio'lari tam olarak asagidaki KUMULATIF_PAY dizisinden cikar ve
arac bunlari kusursuz geri bulmalidir.

Calistirma:
    python3 data/uret.py
"""

import os
import numpy as np
import pandas as pd

TOHUM = 20260821
BRANS = 715          # sektor standardi kod; herhangi bir kuruma ait degil
HAVUZ = 0
ILK_DONEM = "2020Q1"
KAZA_DONEMI = 24     # ceyrek
GELISIM = 12         # ceyrek

# Kumulatif gelisim payi: d. ceyrek sonunda nihai hasarin yuzde kaci
# gerceklesmis olur. Son eleman 1.0 DEGIL — kalan kisim kuyruk faktoru.
KUMULATIF_PAY = np.array([
    0.42, 0.61, 0.72, 0.795, 0.85, 0.888,
    0.916, 0.937, 0.953, 0.965, 0.974, 0.981,
])

# Odenen/muallak ayrimi: odenen kismin incurred icindeki payi. Basta
# dusuk (hasar aciliyor, henuz odenmiyor), sonra 1'e yaklasiyor.
ODENEN_PAY = np.array([
    0.18, 0.35, 0.50, 0.62, 0.72, 0.80,
    0.86, 0.90, 0.935, 0.96, 0.98, 0.995,
])

SEGMENTLER = {
    0: dict(ad="maddi", ultimate=(3.0e6, 5.5e6)),   # daha hizli kapanir
    1: dict(ad="bedeni", ultimate=(1.2e6, 2.8e6)),  # daha uzun kuyruk
}


def donem_etiketi(baslangic: str, kaydir: int) -> str:
    yil, ceyrek = int(baslangic[:4]), int(baslangic[-1])
    toplam = yil * 4 + (ceyrek - 1) + kaydir
    return f"{toplam // 4}Q{toplam % 4 + 1}"


def beklenen_faktorler(bedeni: int) -> list[float]:
    """Aracin geri bulmasi GEREKEN link ratio'lar."""
    pay = KUMULATIF_PAY.copy()
    if bedeni:
        # Bedeni hasarlar daha yavas gelisir — deseni hafifce yay
        pay = pay ** 1.15
    return [float(pay[j + 1] / pay[j]) for j in range(len(pay) - 1)]


def uret(sovtaj: bool = True) -> pd.DataFrame:
    rng = np.random.default_rng(TOHUM)
    satirlar = []
    diyagonal = KAZA_DONEMI - 1     # en son gozlemlenen kosegen

    for bedeni, cfg in SEGMENTLER.items():
        pay = KUMULATIF_PAY.copy()
        if bedeni:
            pay = pay ** 1.15
        odenen_pay = ODENEN_PAY.copy()

        for i in range(KAZA_DONEMI):
            alt, ust = cfg["ultimate"]
            ultimate = rng.uniform(alt, ust)

            onceki_odenen = 0.0
            for d in range(GELISIM):
                # Yalnizca GOZLEMLENMIS hucreler — gercek veri de boyledir
                if i + d > diyagonal:
                    break

                incurred = ultimate * pay[d]
                odenen_kum = incurred * odenen_pay[d]
                artis = odenen_kum - onceki_odenen
                onceki_odenen = odenen_kum
                muallak = incurred - odenen_kum

                ortak = dict(
                    HasarDonem=donem_etiketi(ILK_DONEM, i),
                    GelDonemHT=d,
                    HazAltBrnKod=BRANS,
                    Bedeni=bedeni,
                    Havuz=HAVUZ,
                )
                if artis > 0:
                    satirlar.append({**ortak, "HasarDurum": "Odenen",
                                     "HasarTutar": round(artis, 2)})
                if muallak > 0:
                    satirlar.append({**ortak, "HasarDurum": "Muallak",
                                     "HasarTutar": round(muallak, 2)})

            # Az sayida sovtaj/rucu kaydi — isaret mantigini sinamak icin.
            # Son kaza donemlerinde gozlem penceresi 1 ceyregi gecmiyor;
            # orada kayit uretilmez (yoksa kosegen disina tasar).
            ust = min(6, diyagonal - i + 1)
            if sovtaj and rng.random() < 0.30 and ust > 1:
                d = int(rng.integers(1, ust))
                satirlar.append(dict(
                    HasarDonem=donem_etiketi(ILK_DONEM, i), GelDonemHT=d,
                    HazAltBrnKod=BRANS, Bedeni=bedeni, Havuz=HAVUZ,
                    HasarDurum="SovtajRucu",
                    HasarTutar=round(-ultimate * rng.uniform(0.002, 0.008), 2),
                ))

    return pd.DataFrame(satirlar)


if __name__ == "__main__":
    df = uret()
    hedef = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "ornek_hasar.xlsx")
    df.to_excel(hedef, index=False)

    print(f"{len(df):,} satir yazildi -> {hedef}")
    print(f"Kaza donemi: {KAZA_DONEMI} ceyrek · gelisim: {GELISIM} ceyrek")
    print()
    print("Aracin geri bulmasi GEREKEN link ratio'lar (maddi):")
    for j, f in enumerate(beklenen_faktorler(0)):
        print(f"  {j}->{j+1}: {f:.4f}")
    print()
    print("Gercek kuyruk faktoru (maddi): "
          f"{1.0 / KUMULATIF_PAY[-1]:.4f}")

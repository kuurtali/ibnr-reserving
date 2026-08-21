"""CAS veri donusumunun testleri.

Indirme testi yok - ag erisimi gerektiriyor. Test edilen sey DONUSUM:
CAS'in kumulatif ucgeninden aracin bekledigi kayit tablosuna gecis
dogru mu. Sahte ama CAS SEKLINDE bir cerceve kuruluyor.
"""

import os
import sys

import pandas as pd
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(KOK, "data"))

from cas_indir import (                                    # noqa: E402
    GELISIM_YILI, ILK_KAZA_YILI, donem_etiketi, donustur, gercek_ultimate,
)


def cas_ornegi() -> pd.DataFrame:
    """Iki sirket x 10 kaza yili x 10 gelisim - tam kare."""
    satirlar = []
    for kod in (100, 200):
        for i in range(10):
            ultimate = 1_000_000 * (kod / 100) * (1 + i * 0.05)
            kum_odenen = 0.0
            for d in range(GELISIM_YILI):
                pay = min(1.0, 0.30 + 0.08 * d)
                incurred = ultimate * min(1.0, 0.85 + 0.017 * d)
                kum_odenen = incurred * pay
                satirlar.append(dict(
                    GRCODE=kod, GRNAME=f"Sirket {kod}",
                    AccidentYear=ILK_KAZA_YILI + i,
                    DevelopmentLag=d + 1,
                    IncurLoss_B=incurred,
                    CumPaidLoss_B=kum_odenen,
                    EarnedPremNet_B=ultimate * 1.2,
                ))
    return pd.DataFrame(satirlar)


@pytest.fixture(scope="module")
def ham():
    return cas_ornegi()


def test_donem_etiketleri_ardisik():
    etiketler = [donem_etiketi(i) for i in range(5)]
    assert etiketler == ["2000Q1", "2000Q2", "2000Q3", "2000Q4", "2001Q1"]


def test_ust_ucgen_disi_hucre_gelmez(ham):
    kayit = donustur(ham, "B", sadece_ust=True)
    kayit["i"] = [int(donem_etiketi_ters(x)) for x in kayit["HasarDonem"]]
    assert (kayit["i"] + kayit["GelDonemHT"] <= GELISIM_YILI - 1).all()


def donem_etiketi_ters(etiket: str) -> int:
    yil, ceyrek = int(etiket[:4]), int(etiket[-1])
    taban = 2000 * 4 + 0
    return yil * 4 + (ceyrek - 1) - taban


def test_odenen_satirlari_ARTIS_kumulatif_degil(ham):
    kayit = donustur(ham, "B", sadece_ust=True)
    bir = kayit[(kayit["GRCODE"] == 100) & (kayit["HasarDonem"] == "2000Q1")
                & (kayit["HasarDurum"] == "Odenen")].sort_values("GelDonemHT")
    # Artislarin toplami, o satirin son kumulatif odenenine esit olmali
    ham100 = ham[(ham["GRCODE"] == 100) & (ham["AccidentYear"] == ILK_KAZA_YILI)]
    son_kum = ham100.sort_values("DevelopmentLag")["CumPaidLoss_B"].iloc[
        len(bir) - 1]
    assert bir["HasarTutar"].sum() == pytest.approx(son_kum)


def test_muallak_acik_tutari(ham):
    kayit = donustur(ham, "B", sadece_ust=True)
    m = kayit[(kayit["GRCODE"] == 100) & (kayit["HasarDonem"] == "2000Q1")
              & (kayit["GelDonemHT"] == 0) & (kayit["HasarDurum"] == "Muallak")]
    h = ham[(ham["GRCODE"] == 100) & (ham["AccidentYear"] == ILK_KAZA_YILI)
            & (ham["DevelopmentLag"] == 1)]
    beklenen = float(h["IncurLoss_B"].iloc[0] - h["CumPaidLoss_B"].iloc[0])
    assert float(m["HasarTutar"].iloc[0]) == pytest.approx(beklenen)


def test_incurred_ucgen_yeniden_kurulabiliyor(ham):
    """Odenen kumulatifi + muallak = incurred. Aracin yaptigi islem bu."""
    kayit = donustur(ham, "B", sadece_ust=True)
    k = kayit[(kayit["GRCODE"] == 100) & (kayit["HasarDonem"] == "2000Q1")]
    for d in sorted(k["GelDonemHT"].unique()):
        odenen_kum = k[(k["HasarDurum"] == "Odenen")
                       & (k["GelDonemHT"] <= d)]["HasarTutar"].sum()
        muallak = k[(k["HasarDurum"] == "Muallak")
                    & (k["GelDonemHT"] == d)]["HasarTutar"].sum()
        h = ham[(ham["GRCODE"] == 100) & (ham["AccidentYear"] == ILK_KAZA_YILI)
                & (ham["DevelopmentLag"] == d + 1)]
        assert odenen_kum + muallak == pytest.approx(float(h["IncurLoss_B"].iloc[0]))


def test_gercek_ultimate_son_gelisimden_gelir(ham):
    ger = gercek_ultimate(ham, "B")
    assert len(ger) == 20                       # 2 sirket x 10 kaza yili
    son = ham[(ham["GRCODE"] == 100) & (ham["AccidentYear"] == ILK_KAZA_YILI)
              & (ham["DevelopmentLag"] == GELISIM_YILI)]
    satir = ger[(ger["GRCODE"] == 100) & (ger["AccidentYear"] == ILK_KAZA_YILI)]
    assert float(satir["GercekUltimate"].iloc[0]) == pytest.approx(
        float(son["IncurLoss_B"].iloc[0]))


def test_eksik_kolon_reddedilir(ham):
    with pytest.raises(ValueError):
        donustur(ham.drop(columns=["IncurLoss_B"]), "B")

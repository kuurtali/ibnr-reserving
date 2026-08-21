"""Chain ladder cekirdeginin testleri.

Sentetik veri BILINEN bir gelisim deseninden uretildigi icin (data/uret.py)
aracin dogru sonuc verip vermedigi olculebiliyor: veriyi kurarken
kullandigimiz link ratio'lari geri bulmasi gerekiyor.

Calistirma:  pytest -q
"""

import os
import sys

import numpy as np
import pandas as pd
import pytest

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, KOK)
sys.path.insert(0, os.path.join(KOK, "data"))

import ibnr                                      # noqa: E402
from uret import beklenen_faktorler, uret        # noqa: E402


@pytest.fixture(scope="module")
def veri():
    return uret()


@pytest.fixture(scope="module")
def veri_sovtajsiz():
    """Sovtaj/rucu kayitlari ucgeni hafifce oynatir; faktorlerin TAM
    olarak geri gelmesini sinayan test bu gurultuyu istemiyor."""
    return uret(sovtaj=False)


@pytest.fixture(scope="module")
def meta(veri, tmp_path_factory):
    yol = tmp_path_factory.mktemp("veri") / "hasar.xlsx"
    veri.to_excel(yol, index=False)
    _, m = ibnr.veri_oku(str(yol))
    return m


def ayar(**ek):
    temel = dict(ibnr.VARSAYILAN_AYAR)
    temel.update(brans=715, havuz=0, ceyrek=24, pencere="all", tail=1.0,
                 tail_otomatik=False, sovtaj=1, senaryolar=False,
                 geriye_test=False)
    temel.update(ek)
    return temel


# ───────────────────────────────────────────────── donem kodlamasi

def test_donem_kodu_gidis_donus():
    for q in ["2020Q1", "2023Q4", "2026Q2"]:
        assert ibnr.kod_donem(ibnr.donem_kodu(q)) == q


def test_gecersiz_donem_reddedilir():
    for kotu in ["2026", "2026Q5", "26Q1", "abc"]:
        with pytest.raises(ValueError):
            ibnr.donem_kodu(kotu)


# ─────────────────────────────────────── cekirdek: faktorler geri gelmeli

def test_link_ratiolar_bilinen_desene_esit(veri_sovtajsiz, meta):
    a = ayar(degerleme=ibnr.kod_donem(meta["veri_son"]))
    a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
    sonuc = ibnr.segment_hesapla(veri_sovtajsiz, 0, a, meta)

    bekl = beklenen_faktorler(0)
    bulunan = sonuc["hesaplanan_faktorler"]

    # Veri kurulusu 2 ondaliga yuvarlandigi icin tam esitlik degil,
    # ama sapma parasal yuvarlama mertebesinde kalmali.
    assert len(bulunan) >= len(bekl)
    for j, b in enumerate(bekl):
        assert bulunan[j] == pytest.approx(b, abs=1e-6), f"{j}. faktor sapti"


def test_tail_bir_ise_ultimate_son_sutun(veri, meta):
    a = ayar(degerleme=ibnr.kod_donem(meta["veri_son"]), tail=1.0)
    a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
    s = ibnr.segment_hesapla(veri, 0, a, meta)
    son = s["projekte"][s["son_indeks"]]
    assert s["detay"]["Nihai_Hasar_Ultimate"].values == pytest.approx(son.values)


def test_kuyruk_ultimatei_buyutur(veri, meta):
    ortak = dict(degerleme=ibnr.kod_donem(meta["veri_son"]))
    a1 = ayar(tail=1.0, **ortak); a1["degerleme_kodu"] = ibnr.donem_kodu(a1["degerleme"])
    a2 = ayar(tail=1.1, **ortak); a2["degerleme_kodu"] = a1["degerleme_kodu"]
    u1 = ibnr.segment_hesapla(veri, 0, a1, meta)["ultimate"]
    u2 = ibnr.segment_hesapla(veri, 0, a2, meta)["ultimate"]
    assert u2 > u1


def test_ibnr_pozitif_ve_ultimate_kumulatiften_buyuk(veri, meta):
    a = ayar(degerleme=ibnr.kod_donem(meta["veri_son"]))
    a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
    s = ibnr.segment_hesapla(veri, 0, a, meta)
    assert s["ibnr"] > 0
    assert s["ultimate"] > s["kumulatif"]


def test_diyagonal_disi_hucre_projeksiyona_girmez(veri, meta):
    a = ayar(degerleme=ibnr.kod_donem(meta["veri_son"]))
    a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
    s = ibnr.segment_hesapla(veri, 0, a, meta)
    p, d = s["periyot"], s["diyagonal"]
    maske = np.add.outer(np.arange(p), np.arange(p)) > d
    assert (s["ucgen"].values[maske] == 0).all()


# ─────────────────────────────────────────────────── kuyruk tahmini

def test_tail_tahmini_artmayan_faktorde_bir_doner():
    assert ibnr.tail_tahmin([1.0, 1.0, 1.0])[0] == 1.0
    assert ibnr.tail_tahmin([1.1, 1.2, 1.3])[0] == 1.0


def test_tail_tahmini_ust_sinirla_kisitli():
    # Cok yavas sonen faktorler sinirsiz kuyruk uretmemeli
    faktorler = [1.0 + 0.5 * (0.99 ** j) for j in range(10)]
    tail, _ = ibnr.tail_tahmin(faktorler)
    assert 1.0 <= tail <= 1.5


def test_tail_tahmini_sonen_faktorde_birden_buyuk():
    faktorler = [1.0 + 0.4 * (0.5 ** j) for j in range(8)]
    tail, not_ = ibnr.tail_tahmin(faktorler)
    assert tail > 1.0
    assert "uyum" in not_


# ─────────────────────────────────────────────── faktor penceresi

def test_pencere_secimi_faktorleri_degistirebilir(veri, meta):
    ortak = dict(degerleme=ibnr.kod_donem(meta["veri_son"]))
    sonuclar = {}
    for pencere in ibnr.PENCERE_SECENEKLERI:
        a = ayar(pencere=pencere, **ortak)
        a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
        sonuclar[pencere] = ibnr.segment_hesapla(veri, 0, a, meta)["ultimate"]
    # Desen sabit uretildigi icin uc pencere de ayni sonuca yakin cikmali;
    # bu, pencere mantiginin veriyi bozmadigini gosterir.
    degerler = list(sonuclar.values())
    assert max(degerler) / min(degerler) < 1.02


def test_ucgen_iki_ceyrekten_kisa_olamaz(veri, meta):
    a = ayar(ceyrek=1, degerleme=ibnr.kod_donem(meta["veri_son"]))
    a["degerleme_kodu"] = ibnr.donem_kodu(a["degerleme"])
    with pytest.raises(ValueError):
        ibnr.segment_hesapla(veri, 0, a, meta)


# ───────────────────────────────────────────────────── girdi dogrulama

def test_dogrulayicilar_kotu_girdiyi_reddeder():
    with pytest.raises(ValueError):
        ibnr.tail_dogrula("0")
    with pytest.raises(ValueError):
        ibnr.isaret_dogrula("2")
    with pytest.raises(ValueError):
        ibnr.pencere_dogrula("last_10y")


def test_zorunlu_kolonlar_eksikse_hata(tmp_path):
    eksik = pd.DataFrame({"HasarDonem": ["2020Q1"], "HasarTutar": [100.0]})
    yol = tmp_path / "eksik.xlsx"
    eksik.to_excel(yol, index=False)
    # NOT: veri_oku bir veri hatasinda SystemExit firlatiyor. Kutuphane
    # koduna gore yanlis secim (cagiran taraf yakalayamaz, test edilemez)
    # ama davranis su an bu; test onu belgeliyor.
    with pytest.raises(SystemExit):
        ibnr.veri_oku(str(yol))

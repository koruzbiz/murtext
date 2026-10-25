# -*- coding: utf-8 -*-
"""
KoruzBiz_MurText NVDA eklentisinin tr() kullanımlarını dosyalardan toplayıp .po ile karşılaştırır.

Bu dosya eklenti paketinin dizinine (KoruzBiz_MurText.py, settings.py, __init__.py ile aynı yere) konur.
PO dosyaları: locales/<dil>/LC_MESSAGES/KoruzBiz_MurText.po

Parametre: --rapor  → soru sormadan yalnızca po-*.txt raporlarını üretir.

Üretilen dosyalar (script dizininde):
- po-eklenecekler.txt    : Kodda var, PO'da yok (eklenmesi gerekenler)
- po-kaldirilacak.txt    : PO'da var, kodda yok (artık kullanılmayanlar)
- po-eslesenler.txt      : Hem kodda hem PO'da olanlar
- po-ozet.txt            : Tüm diller için özet (eklenecek/kaldırılacak/eşleşen sayıları)

Ek özellik - Adım 1:
- İşlem bitince sorar: "Kaldırılması gerekenler yorum satırı yapılsın mı?"
  - Hayır: hiçbir şey yapmaz
  - Evet : tüm dillerdeki po-kaldirilacak.txt içindeki msgid'lerin .po girdilerini "#-" ile yorum satırı yapar
- Yorumlama yapıldıysa 2 saniye sonra ikinci soru sorar:
  "#- ile işaretlenenler .po dosyasından silinsin mi?"
  - Hayır: hiçbir şey yapmaz
  - Evet : tüm dillerdeki PO'larda "#-" ile başlayan satırların tamamını siler

Ek özellik - Adım 2:
- Adım 1 sonrası sorar: "Eksikleri yüklemek istiyor musun?"
  - Hayır: çıkar
  - Evet : locales/locale dizininde "ekle.txt" arar
    - ekle.txt içinde "#+dil_kodu" işaretçileriyle dillere ayrılmış msgid/msgstr blokları var
    - Her dile ait blok ilgili .po dosyasına eklenir
    - Zaten olan msgid'ler tekrar eklenmez

ekle.txt örnek yapısı:
  #+en
  msgid "Örnek metin"
  msgstr "Example text"
  #+de
  msgid "Örnek metin"
  msgstr "Beispieltext"

Notlar:
- Tarama regex ile değil, AST ile yapıldığı için girinti, satır bölmeleri,
  sarmalanmış kullanımlar (label(tr("..."))) vb. stabil şekilde yakalanır.
- tr(...) içindeki birleştirilmiş string literal'lar ("a""b") tek metin olur.
"""

from __future__ import annotations

import ast
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set

import logging
import sys

logger = logging.getLogger("po_murtext_eklenti")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _fh = logging.FileHandler(Path(__file__).resolve().parent / "po-log.txt", mode="w", encoding="utf-8")
    _fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(_fh)


# Taranacak dosyalar
taranacak = {
    "KoruzBiz_MurText.py": True,
    "settings.py": True,
    "__init__.py": True,
}

# Çeviri fonksiyon adları (projede tr() / self.tr() / self._tr() görülebiliyor)
ceviri_fonksiyon_adlari = {"tr", "_tr", "N_"}

# PO dosya adı (her dil dizininin LC_MESSAGES altında aranır)
po_dosya_adi = "KoruzBiz_MurText.po"

# Dil dizinlerinin aranacağı klasör adları (script dizinine göre)
locales_klasor_adlari = ["locales", "locale"]


def po_dizinlerini_bul(kok: Path) -> List[Path]:
    """
    locales veya locale dizinindeki alt dizinleri tarar,
    LC_MESSAGES/KoruzBiz_MurText.po içeren dizinleri döndürür.
    Sıralama: dizin adına göre artan (01, 02, ...).
    """
    bulunan: List[Path] = []
    for klasor_adi in locales_klasor_adlari:
        klasor = kok / klasor_adi
        if not klasor.is_dir():
            continue
        for alt in sorted(klasor.iterdir()):
            if not alt.is_dir():
                continue
            po_yolu = alt / "LC_MESSAGES" / po_dosya_adi
            if po_yolu.exists():
                bulunan.append(po_yolu)
    return bulunan


def ekle_txt_bul(kok: Path) -> Optional[Path]:
    """
    locales veya locale dizininde ekle.txt arar.
    Bulunursa Path döndürür, bulunamazsa None.
    """
    for klasor_adi in locales_klasor_adlari:
        klasor = kok / klasor_adi
        aday = klasor / "ekle.txt"
        if aday.exists():
            return aday
    return None


def ekle_txt_parse(ekle_yolu: Path) -> Dict[str, List[str]]:
    """
    ekle.txt dosyasını parse eder.
    "#+dil_kodu" satırlarını bölüm başlangıcı olarak kullanır.
    Her dil kodu için o dile ait satır listesini döndürür.

    Dönüş: {"en": ["msgid ...", "msgstr ...", ...], "de": [...], ...}
    """
    try:
        icerik = _dosya_oku_utf8(ekle_yolu)
    except Exception as e:
        logger.error("ekle.txt okunamadı: %s | Hata: %s", ekle_yolu, e)
        return {}

    sonuc: Dict[str, List[str]] = {}
    aktif_dil: Optional[str] = None

    for satir in icerik.splitlines():
        temiz = satir.strip()
        if temiz.startswith("#+"):
            dil_kodu = temiz[2:].strip().lower()
            if dil_kodu:
                aktif_dil = dil_kodu
                if aktif_dil not in sonuc:
                    sonuc[aktif_dil] = []
            continue
        if aktif_dil is not None:
            sonuc[aktif_dil].append(satir)

    # Her dil için baştaki ve sondaki boş satırları temizle
    for dil in sonuc:
        satirlar = sonuc[dil]
        while satirlar and not satirlar[0].strip():
            satirlar.pop(0)
        while satirlar and not satirlar[-1].strip():
            satirlar.pop()
        sonuc[dil] = satirlar

    logger.info("ekle.txt parse edildi. Dil sayısı: %d | Diller: %s", len(sonuc), list(sonuc.keys()))
    return sonuc

def po_ya_ekle(po_yolu: Path, eklenecek_satirlar: List[str], mevcut_msgidler: Set[str]) -> int:
    """
    eklenecek_satirlar listesindeki msgid/msgstr bloklarını po_yolu'na ekler.
    mevcut_msgidler kümesinde zaten olan msgid'ler atlanır.
    Tek satırlı ve çok satırlı msgid formatlarını destekler.
    Dönüş: eklenen blok sayısı.
    """
    if not eklenecek_satirlar:
        return 0

    # Blokları ayır: msgid ile başlar, msgstr içerince ve boş satır görününce kapanır
    bloklar: List[List[str]] = []
    aktif_blok: List[str] = []

    for satir in eklenecek_satirlar:
        temiz = satir.strip()
        if temiz.startswith("msgid "):
            if aktif_blok:
                bloklar.append(aktif_blok)
            aktif_blok = [satir]
        elif not temiz and aktif_blok:
            mevcut_metin = "\n".join(aktif_blok)
            if "msgstr" in mevcut_metin:
                bloklar.append(aktif_blok)
                aktif_blok = []
        else:
            if aktif_blok:
                aktif_blok.append(satir)

    if aktif_blok:
        bloklar.append(aktif_blok)

    if not bloklar:
        logger.info("ekle: blok bulunamadı. po_yolu=%s", po_yolu)
        return 0

    eklenen = 0
    eklenecek_metin_parcalari: List[str] = []

    for blok in bloklar:
        msgid_satiri = next((s for s in blok if s.strip().startswith("msgid ")), None)
        if not msgid_satiri:
            continue

        parca = msgid_satiri.strip()[len("msgid "):].strip()

        if parca == '""':
            # Çok satırlı msgid: msgid "" satırından sonraki tırnak satırlarını birleştir
            msgid_parcalari = []
            msgid_idx = blok.index(msgid_satiri)
            for devam_satiri in blok[msgid_idx + 1:]:
                temiz_d = devam_satiri.strip()
                if temiz_d.startswith("msgstr"):
                    break
                if temiz_d.startswith('"') and temiz_d.endswith('"'):
                    try:
                        parca_deger = ast.literal_eval(temiz_d)
                    except Exception:
                        parca_deger = temiz_d[1:-1]
                    msgid_parcalari.append(parca_deger)
            msgid_degeri = "".join(msgid_parcalari)
        elif parca.startswith('"') and parca.endswith('"'):
            try:
                msgid_degeri = ast.literal_eval(parca)
            except Exception:
                msgid_degeri = parca[1:-1]
        else:
            msgid_degeri = ""

        if not msgid_degeri:
            logger.info("ekle: msgid cozulemedi, atlandı. blok=%s", blok[:2])
            continue

        if msgid_degeri in mevcut_msgidler:
            logger.info("ekle: zaten var, atlandı: %r | %s", msgid_degeri[:50], po_yolu.parent.parent.name)
            continue

        # Bloğun sonundaki boş satırları temizle
        temiz_blok = list(blok)
        while temiz_blok and not temiz_blok[-1].strip():
            temiz_blok.pop()

        eklenecek_metin_parcalari.append("")
        eklenecek_metin_parcalari.extend(temiz_blok)
        eklenen += 1

    if not eklenecek_metin_parcalari:
        logger.info("ekle: eklenecek yeni blok yok. po_yolu=%s", po_yolu)
        return 0

    _dosya_yedekle(po_yolu, ".pre_ekle.bak")

    try:
        mevcut = _dosya_oku_utf8(po_yolu)
        if not mevcut.endswith("\n"):
            mevcut += "\n"
        yeni_icerik = mevcut + "\n".join(eklenecek_metin_parcalari) + "\n"
        po_yolu.write_text(yeni_icerik, encoding="utf-8")
        logger.info("ekle: %d blok eklendi. po_yolu=%s", eklenen, po_yolu)
    except Exception as e:
        logger.error("ekle: PO yazılamadı: %s | Hata: %s", po_yolu, e)
        return 0

    return eklenen

def po_isaretli_satir_sayisini_al(po_yolu: Path) -> int:
    try:
        kaynak = _dosya_oku_utf8(po_yolu)
    except Exception as e:
        logger.error("PO okunamadı (işaretli satır sayımı): %s | Hata: %s", po_yolu, e)
        return 0
    sayi = 0
    for s in kaynak.splitlines():
        if s.lstrip().startswith("#-"):
            sayi += 1
    return sayi


def _dosya_oku_utf8(dosya_yolu: Path) -> str:
    try:
        return dosya_yolu.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        logger.error("Kodlama hatası: %s UTF-8 olarak okunamadı.", dosya_yolu)
        raise
    except Exception as e:
        logger.error("Dosya okunamadı: %s | Hata: %s", dosya_yolu, e)
        raise


def _dosya_yedekle(dosya_yolu: Path, ek: str) -> None:
    try:
        icerik = _dosya_oku_utf8(dosya_yolu)
    except Exception as e:
        logger.error("Yedek için dosya okunamadı: %s | Hata: %s", dosya_yolu, e)
        return
    yedek = dosya_yolu.with_suffix(dosya_yolu.suffix + ek)
    try:
        if not yedek.exists():
            yedek.write_text(icerik + ("\n" if not icerik.endswith("\n") else ""), encoding="utf-8")
            logger.info("Yedek oluşturuldu: %s", yedek)
        else:
            logger.info("Yedek zaten var, dokunmadım: %s", yedek)
    except Exception as e:
        logger.error("Yedek oluşturulamadı: %s | Hata: %s", yedek, e)


def _string_dugumu_yazisini_al(node: ast.AST) -> Optional[str]:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.Str):
        return node.s
    return None


def tr_metnilerini_topla_koddan(dosya_yolu: Path) -> Set[str]:
    """Dosyadaki tr(...) metinlerini döndürür."""
    kaynak = _dosya_oku_utf8(dosya_yolu)
    try:
        agac = ast.parse(kaynak, filename=str(dosya_yolu))
    except SyntaxError as e:
        logger.error("SyntaxError: %s | Satır: %s | Detay: %s", dosya_yolu, getattr(e, "lineno", "?"), e)
        return set()

    bulunanlar: Set[str] = set()
    for node in ast.walk(agac):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        tr_mi = (
            (isinstance(func, ast.Name) and func.id in ceviri_fonksiyon_adlari)
            or (isinstance(func, ast.Attribute) and func.attr in ceviri_fonksiyon_adlari)
        )
        if not tr_mi:
            continue
        if not node.args:
            continue
        metin = _string_dugumu_yazisini_al(node.args[0])
        if metin is None:
            logger.info(
                "Dinamik/çözülemeyen tr() atlandı: %s | Dosya: %s",
                ast.get_source_segment(kaynak, node) or "tr(...)",
                dosya_yolu,
            )
            continue
        if metin.strip() == "":
            continue
        bulunanlar.add(metin)
    return bulunanlar


def _po_kacislarini_coz(icerik: str) -> str:
    """
    icerik, .po dosyasındaki bir satırın tırnaklar arasındaki ham metnidir
    ve PO kuralına göre içindeki her tırnak zaten \\" olarak escape edilmiştir.
    Bu yüzden burada ekstra bir escape adımı YAPILMAZ; doğrudan literal_eval
    ile çözülür. (Önceden burada yapılan icerik.replace('"', r'\\"') adımı,
    zaten escape edilmiş \\" dizilerini \\\\" haline getirip bozuyordu.)
    """
    try:
        return ast.literal_eval('"' + icerik + '"')
    except Exception:
        logger.error("PO kaçış çözümü başarısız: %r", icerik)
        return icerik


def msgidleri_topla_podan(po_yolu: Path) -> Set[str]:
    kaynak = _dosya_oku_utf8(po_yolu)
    satirlar = kaynak.splitlines()
    msgidler: Set[str] = set()
    i = 0
    while i < len(satirlar):
        s = satirlar[i].lstrip()
        if s.startswith("msgid "):
            parca = s[len("msgid "):].strip()
            metin_parcalari: List[str] = []
            if parca.startswith('"') and parca.endswith('"'):
                metin_parcalari.append(_po_kacislarini_coz(parca[1:-1]))
            else:
                logger.info("Beklenmeyen msgid formatı: %r | Satır: %s", satirlar[i], i + 1)
            i += 1
            while i < len(satirlar):
                s2 = satirlar[i].lstrip()
                if s2.startswith('"') and s2.endswith('"'):
                    metin_parcalari.append(_po_kacislarini_coz(s2[1:-1]))
                    i += 1
                    continue
                break
            msgid = "".join(metin_parcalari)
            if msgid.strip() != "":
                msgidler.add(msgid)
            continue
        i += 1
    return msgidler

def _dosyaya_yaz(hedef: Path, satirlar: Iterable[str], baslik: str) -> None:
    icerik = []
    icerik.append(baslik)
    icerik.append("-" * len(baslik))
    for s in satirlar:
        icerik.append(s)
    icerik.append("")
    hedef.write_text("\n".join(icerik), encoding="utf-8")


def _po_metni_bicimlendir(metin: str) -> str:
    """
    Tek bir msgid metnini PO formatında biçimlendirir.
    Satır içi \\n varsa çok satırlı PO formatı kullanır.
    """
    def _po_kacis(s: str) -> str:
        return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")

    if "\n" in metin:
        # metin gerçek bir \n ile bitiyorsa, split("\n") sondaki \n yüzünden
        # fazladan boş bir eleman üretir (ör: "abc\n".split("\n") -> ["abc", ""]).
        # Bu boş elemanı ayrı bir tırnaklı satır olarak yazmamak için işaretleyip atıyoruz.
        newline_ile_bitiyor = metin.endswith("\n")
        satirlar = metin.split("\n")
        if newline_ile_bitiyor:
            satirlar = satirlar[:-1]

        parcalar = ['msgid ""']
        for satir in satirlar:
            parcalar.append('"' + _po_kacis(satir) + '\\n"')

        if not newline_ile_bitiyor:
            # Metin \n ile bitmiyorsa son satırdaki fazladan \n kaldırılır
            if parcalar[-1].endswith('\\n"'):
                parcalar[-1] = parcalar[-1][:-3] + '"'

        parcalar.append('msgstr ""')
        return "\n".join(parcalar)
    else:
        return 'msgid "' + _po_kacis(metin) + '"\nmsgstr ""'


def _eklenecekler_yaz(hedef: Path, msgidler: Iterable[str], baslik: str) -> None:
    """
    po-eklenecekler.txt için PO formatında yazar.
    Her msgid msgstr "" çifti olarak, boş satırla ayrılmış şekilde çıkar.
    """
    satirlar = [baslik, "-" * len(baslik), ""]
    for metin in msgidler:
        satirlar.append(_po_metni_bicimlendir(metin))
        satirlar.append("")
    hedef.write_text("\n".join(satirlar), encoding="utf-8")

def _wx_evet_hayir_sor(soru: str, baslik: str) -> bool:
    import wx
    app = wx.GetApp()
    kendi_appim = False
    if app is None:
        app = wx.App(False)
        kendi_appim = True
    try:
        dlg = wx.MessageDialog(
            None,
            soru,
            baslik,
            style=wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION | wx.STAY_ON_TOP,
        )
        try:
            return dlg.ShowModal() == wx.ID_YES
        finally:
            dlg.Destroy()
    finally:
        if kendi_appim:
            try:
                app.Destroy()
            except Exception:
                pass


def kaldirilacaklari_txtden_oku(txt_yolu: Path) -> Set[str]:
    try:
        icerik = _dosya_oku_utf8(txt_yolu)
    except Exception as e:
        logger.error("Kaldırılacaklar txt okunamadı: %s | Hata: %s", txt_yolu, e)
        return set()
    satirlar = [s.rstrip("\n") for s in icerik.splitlines()]
    msgidler: Set[str] = set()
    for s in satirlar:
        s2 = (s or "").strip()
        if not s2:
            continue
        if s2.startswith("PO'dan kaldırılabilecek"):
            continue
        if set(s2) == {"-"}:
            continue
        msgidler.add(s2)
    return msgidler


def _po_msgid_coz_satirlarindan(msgid_satirlari: List[str]) -> str:
    if not msgid_satirlari:
        return ""
    ilk = msgid_satirlari[0].lstrip()
    if not ilk.startswith("msgid "):
        return ""
    parca = ilk[len("msgid "):].strip()
    metin_parcalari: List[str] = []
    if parca.startswith('"') and parca.endswith('"'):
        metin_parcalari.append(_po_kacislarini_coz(parca[1:-1]))
    for s in msgid_satirlari[1:]:
        s2 = s.lstrip()
        if s2.startswith('"') and s2.endswith('"'):
            metin_parcalari.append(_po_kacislarini_coz(s2[1:-1]))
    return "".join(metin_parcalari)


def po_kaldirilacaklari_yorum_satiri_yap(po_yolu: Path, kaldirilacaklar: Set[str]) -> int:
    kaynak = _dosya_oku_utf8(po_yolu)
    satirlar = kaynak.splitlines()
    _dosya_yedekle(po_yolu, ".bak")
    yorumlanan_entry_sayisi = 0
    yeni: List[str] = []
    i = 0
    while i < len(satirlar):
        s = satirlar[i]
        ls = s.lstrip()
        if ls.startswith("msgid "):
            msgid_satirlari: List[str] = [s]
            j = i + 1
            while j < len(satirlar):
                ls2 = satirlar[j].lstrip()
                if ls2.startswith('"') and ls2.endswith('"'):
                    msgid_satirlari.append(satirlar[j])
                    j += 1
                    continue
                break
            msgid = _po_msgid_coz_satirlarindan(msgid_satirlari)
            k = j
            while k < len(satirlar):
                lsk = satirlar[k].lstrip()
                if lsk.startswith("msgid "):
                    break
                k += 1
            entry_satirlari = satirlar[i:k]
            if msgid and (msgid in kaldirilacaklar):
                yorumlanan_entry_sayisi += 1
                for es in entry_satirlari:
                    if (es or "").strip() == "":
                        yeni.append(es)
                        continue
                    if es.lstrip().startswith("#"):
                        yeni.append(es)
                        continue
                    yeni.append("#- " + es)
            else:
                yeni.extend(entry_satirlari)
            i = k
            continue
        yeni.append(s)
        i += 1
    try:
        po_yolu.write_text("\n".join(yeni) + "\n", encoding="utf-8")
        logger.info("PO güncellendi. Yorumlanan entry sayısı: %d | %s", yorumlanan_entry_sayisi, po_yolu)
    except Exception as e:
        logger.error("PO yazılamadı: %s | Hata: %s", po_yolu, e)
        return 0
    return yorumlanan_entry_sayisi


def _bos_satirlari_normallestir(satirlar: List[str]) -> List[str]:
    yeni: List[str] = []
    onceki_bos = False
    for s in satirlar:
        bos = ((s or "").strip() == "")
        if bos:
            if onceki_bos:
                continue
            yeni.append("")
            onceki_bos = True
            continue
        yeni.append(s)
        onceki_bos = False
    while yeni and yeni[0].strip() == "":
        yeni.pop(0)
    return yeni


def po_isaretlenenleri_sil(po_yolu: Path, bos_satir_normallestir: bool = True) -> int:
    kaynak = _dosya_oku_utf8(po_yolu)
    satirlar = kaynak.splitlines()
    _dosya_yedekle(po_yolu, ".predelete.bak")
    yeni: List[str] = []
    silinen = 0
    for s in satirlar:
        if s.lstrip().startswith("#-"):
            silinen += 1
            continue
        yeni.append(s)
    if bos_satir_normallestir:
        yeni = _bos_satirlari_normallestir(yeni)
    try:
        po_yolu.write_text("\n".join(yeni) + "\n", encoding="utf-8")
        logger.info("PO güncellendi. Silinen '#-' satır sayısı: %d | %s", silinen, po_yolu)
    except Exception as e:
        logger.error("PO yazılamadı: %s | Hata: %s", po_yolu, e)
        return 0
    return silinen


def _bos_msgstr_bul(po_yolu: Path) -> List[str]:
    """
    PO dosyasında msgstr boş olan kayıtların msgid değerlerini döndürür.
    Header bloğu (msgid "") hariç tutulur.
    """
    try:
        kaynak = _dosya_oku_utf8(po_yolu)
    except Exception as e:
        logger.error("_bos_msgstr_bul: PO okunamadı: %s | %s", po_yolu, e)
        return []

    satirlar = kaynak.splitlines()
    bos_msgstrler: List[str] = []
    i = 0

    while i < len(satirlar):
        temiz = satirlar[i].strip()

        if temiz.startswith("msgid "):
            parca = temiz[len("msgid "):].strip()

            if parca == '""':
                # Çok satırlı msgid - parcaları topla
                msgid_parcalari = []
                j = i + 1
                while j < len(satirlar):
                    s = satirlar[j].strip()
                    if s.startswith("msgstr"):
                        # msgstr boş mu kontrol et
                        if s == 'msgstr ""':
                            # Sonraki satır da tırnak değilse gerçekten boş
                            if j + 1 >= len(satirlar) or not satirlar[j + 1].strip().startswith('"'):
                                msgid_val = "".join(msgid_parcalari)
                                if msgid_val:
                                    bos_msgstrler.append(msgid_val)
                        break
                    if s.startswith('"') and s.endswith('"'):
                        try:
                            msgid_parcalari.append(ast.literal_eval(s))
                        except Exception:
                            msgid_parcalari.append(s[1:-1])
                    j += 1
                i = j + 1
                continue

            elif parca.startswith('"') and parca.endswith('"'):
                # Tek satırlı msgid
                try:
                    msgid_val = ast.literal_eval(parca)
                except Exception:
                    msgid_val = parca[1:-1]

                if msgid_val:
                    # Sonraki satırlarda msgstr bul
                    j = i + 1
                    while j < len(satirlar):
                        s = satirlar[j].strip()
                        if s == 'msgstr ""':
                            if j + 1 >= len(satirlar) or not satirlar[j + 1].strip().startswith('"'):
                                bos_msgstrler.append(msgid_val)
                            break
                        if s.startswith("msgid "):
                            break
                        j += 1

        i += 1

    return bos_msgstrler

def po_basligi_gecerli_mi(po_yolu: Path) -> bool:
    """PO başlığındaki (ilk msgid "") msgstr bloğunda charset var mı bakar; başlık boş satırla bölünmüşse False döner."""
    satirlar = _dosya_oku_utf8(po_yolu).splitlines()
    i = 0
    while i < len(satirlar) and not satirlar[i].lstrip().startswith("msgid"):
        i += 1
    if i >= len(satirlar) or satirlar[i].strip() != 'msgid ""':
        return False
    i += 1
    while i < len(satirlar) and not satirlar[i].lstrip().startswith("msgstr"):
        i += 1
    blok = []
    i += 1
    while i < len(satirlar) and satirlar[i].lstrip().startswith('"'):
        blok.append(satirlar[i])
        i += 1
    return "charset=" in "".join(blok).lower()


def tek_po_isle(po_yolu: Path, kod_metnileri: Set[str], kok: Path) -> Dict:
    """
    Tek bir PO dosyasını işler. Karşılaştırma sonuçlarını döndürür.
    po-rapor.txt yazmaz, özet po-ozet.txt ile script dizininde tutulur.
    """
    dil_kodu = po_yolu.parent.parent.name

    try:
        po_metnileri = msgidleri_topla_podan(po_yolu)
        logger.info("[%s] PO içinde %d adet msgid bulundu.", dil_kodu, len(po_metnileri))
    except Exception as e:
        logger.error("[%s] PO okuma hatası: %s", dil_kodu, e)
        return {"dil": dil_kodu, "hata": str(e), "po_yolu": po_yolu}

    baslik_gecerli = po_basligi_gecerli_mi(po_yolu)
    if not baslik_gecerli:
        logger.warning("[%s] PO başlığında charset bulunamadı.", dil_kodu)

    # Boş msgstr kontrolü: msgstr "" olan kayıtları bul
    bos_msgstr_msgidler = _bos_msgstr_bul(po_yolu)
    if bos_msgstr_msgidler:
        logger.warning("[%s] Boş msgstr olan %d kayıt bulundu.", dil_kodu, len(bos_msgstr_msgidler))

    eklenecekler = sorted(kod_metnileri - po_metnileri)
    kaldirilacaklar = sorted(po_metnileri - kod_metnileri)
    eslesenler = sorted(kod_metnileri & po_metnileri)

    logger.info(
        "[%s] Eklenecek: %d | Kaldırılacak: %d | Eşleşen: %d",
        dil_kodu, len(eklenecekler), len(kaldirilacaklar), len(eslesenler),
    )

    return {
        "dil": dil_kodu,
        "po_yolu": po_yolu,
        "po_metnileri": po_metnileri,
        "eklenecekler": eklenecekler,
        "kaldirilacaklar": kaldirilacaklar,
        "eslesenler": eslesenler,
        "bos_msgstr": bos_msgstr_msgidler,
        "baslik_gecerli": baslik_gecerli,
        "hata": None,
    }

def po_alfabetik_sirala(po_yolu: Path) -> bool:
    """
    PO dosyasındaki msgid/msgstr bloklarını msgid'e göre alfabetik sıralar.
    Gerçek PO header bloğu (Content-Type içeren msgstr) korunur, her zaman başta kalır.
    Yorum satırları kendi msgid bloklarıyla birlikte taşınır.
    Dönüş: True başarılı, False hata.
    """
    try:
        kaynak = _dosya_oku_utf8(po_yolu)
    except Exception as e:
        logger.error("Alfabetik: PO okunamadı: %s | Hata: %s", po_yolu, e)
        return False

    satirlar = kaynak.splitlines()

    # --- BLOK AYIRMA ---
    # Her blok: sıfır veya daha fazla # yorum satırı + msgid + msgstr satırları
    # Yeni blok: msgid ile başlayan satır gelince önceki blok kapanır
    bloklar: List[List[str]] = []
    aktif_blok: List[str] = []
    dosya_basi_satirlar: List[str] = []  # msgid öncesi dosya başı boş/yorum satırları

    ilk_msgid_geldi = False

    for satir in satirlar:
        temiz = satir.strip()

        if not ilk_msgid_geldi:
            if temiz.startswith("msgid "):
                ilk_msgid_geldi = True
                aktif_blok = [satir]
            else:
                dosya_basi_satirlar.append(satir)
            continue

        if temiz.startswith("msgid "):
            if aktif_blok:
                bloklar.append(aktif_blok)
            aktif_blok = [satir]
        else:
            aktif_blok.append(satir)

    if aktif_blok:
        bloklar.append(aktif_blok)

    if not bloklar:
        logger.info("Alfabetik: blok bulunamadı. po_yolu=%s", po_yolu)
        return True

    # --- HEADER TESPİTİ ---
    # Gerçek PO header: msgid "" olan VE msgstr içinde Content-Type veya
    # MIME-Version veya charset içeren blok
    def _blok_header_mi(blok: List[str]) -> bool:
        blok_metni = "\n".join(blok)
        msgid_satiri = next((s for s in blok if s.strip().startswith("msgid ")), None)
        if not msgid_satiri:
            return False
        parca = msgid_satiri.strip()[len("msgid "):].strip()
        if parca not in ('""', "''"):
            return False
        # msgstr içeriğinde meta bilgi var mı
        kontrol = blok_metni.lower()
        return any(
            anahtar in kontrol
            for anahtar in ("content-type", "mime-version", "charset", "po-revision-date", "pot-creation-date")
        )

    header_blok: Optional[List[str]] = None
    diger_bloklar: List[List[str]] = []

    for blok in bloklar:
        if _blok_header_mi(blok):
            if header_blok is None:
                header_blok = blok
            else:
                # İkinci bir header benzeri blok varsa normal bloğa gönder
                diger_bloklar.append(blok)
        else:
            diger_bloklar.append(blok)

    # --- ALFABETİK SIRALAMA ---
    def blok_msgid_getir(blok: List[str]) -> str:
        msgid_satiri = next(
            (s for s in blok if s.strip().startswith("msgid ")), None
        )
        if not msgid_satiri:
            return ""
        parca = msgid_satiri.strip()[len("msgid "):].strip()

        if parca == '""':
            # Çok satırlı: sonraki tırnak satırlarından ilkini al (sıralama için yeterli)
            msgid_idx = blok.index(msgid_satiri)
            for devam in blok[msgid_idx + 1:]:
                temiz_d = devam.strip()
                if temiz_d.startswith("msgstr"):
                    break
                if temiz_d.startswith('"') and temiz_d.endswith('"'):
                    try:
                        return ast.literal_eval(temiz_d).lower()
                    except Exception:
                        return temiz_d[1:-1].lower()
            return ""

        if parca.startswith('"') and parca.endswith('"'):
            try:
                return ast.literal_eval(parca).lower()
            except Exception:
                return parca[1:-1].lower()

        return parca.lower()

    diger_bloklar.sort(key=blok_msgid_getir)

    # --- YENİDEN BİRLEŞTİR ---
    yeni_satirlar: List[str] = []

    # Dosya başı satırlar (encoding bildirimi vs.)
    if dosya_basi_satirlar:
        yeni_satirlar.extend(dosya_basi_satirlar)

    # Header
    if header_blok:
        # Bloğun baş ve sonundaki boş satırları temizle
        while header_blok and not header_blok[0].strip():
            header_blok.pop(0)
        while header_blok and not header_blok[-1].strip():
            header_blok.pop()
        yeni_satirlar.extend(header_blok)
        yeni_satirlar.append("")

    for blok in diger_bloklar:
        while blok and not blok[0].strip():
            blok.pop(0)
        while blok and not blok[-1].strip():
            blok.pop()
        if blok:
            yeni_satirlar.extend(blok)
            yeni_satirlar.append("")

    # Son fazla boş satırları temizle
    while yeni_satirlar and not yeni_satirlar[-1].strip():
        yeni_satirlar.pop()

    _dosya_yedekle(po_yolu, ".pre_alfa.bak")

    try:
        po_yolu.write_text("\n".join(yeni_satirlar) + "\n", encoding="utf-8")
        logger.info("Alfabetik sıralama tamamlandı: %s", po_yolu)
        return True
    except Exception as e:
        logger.error("Alfabetik: PO yazılamadı: %s | Hata: %s", po_yolu, e)
        return False

def po_alfabetik_rapor_yaz(po_yolu: Path, kok: Path, kok_gercek: Path) -> List[str]:
    """
    PO dosyasındaki msgid'leri alfabetik sırayla okur, liste döndürür.
    po_yolu'nun kok_gercek altında olduğunu doğrular.
    """
    try:
        po_yolu.resolve().relative_to(kok_gercek)
    except ValueError:
        logger.error("GÜVENLİK: PO dosyası kok dışında! Atlandı. %s", po_yolu)
        return []

    try:
        mevcut = msgidleri_topla_podan(po_yolu)
        return sorted(mevcut)
    except Exception as e:
        logger.error("Alfabetik rapor: PO okunamadı: %s | Hata: %s", po_yolu, e)
        return []


def calistir() -> int:
    kok = Path(__file__).resolve().parent
    logger.info("Çalışma dizini (script dizini): %s", kok)

    # --- KODDAN MSGID TOPLA ---
    kod_metnileri: Set[str] = set()
    dosya_sayilari: List[str] = []
    eksik_kod_dosyalari: List[str] = []

    for ad in taranacak.keys():
        yol = kok / ad
        if not yol.exists():
            eksik_kod_dosyalari.append(ad)
            continue
        try:
            bulunan = tr_metnilerini_topla_koddan(yol)
            logger.info("%s içinde %d adet tr() metni bulundu.", ad, len(bulunan))
            kod_metnileri.update(bulunan)
            dosya_sayilari.append(f"{ad}: {len(bulunan)} tr() metni")
        except Exception as e:
            logger.error("Tarama hatası: %s | Hata: %s", ad, e)

    if eksik_kod_dosyalari:
        logger.error("Bulunamayan kod dosyaları: %s", ", ".join(eksik_kod_dosyalari))

    logger.info("Koddan toplam %d benzersiz tr() metni toplandı.", len(kod_metnileri))

    # --- PO DİZİNLERİNİ BUL ---
    po_dosyalari = po_dizinlerini_bul(kok)
    if not po_dosyalari:
        logger.error("Hiç PO dosyası bulunamadı. locales/locale dizinleri kontrol edildi.")
        return 2

    logger.info("%d adet PO dosyası bulundu.", len(po_dosyalari))

    # --- TÜM PO'LARI İŞLE ---
    sonuclar: List[Dict] = []
    for po_yolu in po_dosyalari:
        sonuc = tek_po_isle(po_yolu, kod_metnileri, kok)
        sonuclar.append(sonuc)

    # --- ÖZET TXT ---
    ozet_satirlar = [f"Koddan toplam benzersiz tr() metni: {len(kod_metnileri)}"]
    ozet_satirlar.extend("  " + d for d in dosya_sayilari)
    ozet_satirlar.extend(f"  BULUNAMADI: {d}" for d in eksik_kod_dosyalari)
    ozet_satirlar.append("")
    toplam_eklenecek = 0
    toplam_kaldirilacak = 0
    basarili_sonuclar = [s for s in sonuclar if not s.get("hata")]

    for s in basarili_sonuclar:
        ekl = len(s.get("eklenecekler", []))
        kal = len(s.get("kaldirilacaklar", []))
        esl = len(s.get("eslesenler", []))
        toplam_eklenecek += ekl
        toplam_kaldirilacak += kal
        bos = len(s.get("bos_msgstr", []))
        bos_str = f" | Boş msgstr: {bos}" if bos > 0 else ""
        baslik_str = "" if s.get("baslik_gecerli", True) else " | UYARI: PO başlığında charset yok (başlık bölünmüş olabilir)"
        ozet_satirlar.append(
            f"[{s['dil']}] Eklenecek: {ekl} | Kaldırılacak: {kal} | Eşleşen: {esl}{bos_str}{baslik_str}"
        )

    try:
        _dosyaya_yaz(kok / "po-ozet.txt", ozet_satirlar, "Tüm diller için PO özeti")
    except Exception as e:
        logger.error("Özet dosyası yazılamadı: %s", e)

    # Referans dil olarak 01, yoksa ilk bulunan
    ref = next((s for s in basarili_sonuclar if s["dil"] == "01"), None)
    if ref is None and basarili_sonuclar:
        ref = basarili_sonuclar[0]

    if ref:
        try:
            tum_eklenecekler = list(ref["eklenecekler"])
            bos_msgstr_ref = ref.get("bos_msgstr", [])
            if bos_msgstr_ref:
                logger.warning("po-eklenecekler.txt'e %d bos msgstr kaydi ekleniyor.", len(bos_msgstr_ref))
                for msgid in bos_msgstr_ref:
                    if msgid not in tum_eklenecekler:
                        tum_eklenecekler.append(msgid)
            _eklenecekler_yaz(
                kok / "po-eklenecekler.txt",
                tum_eklenecekler,
                "PO'ya eklenecek msgid listesi (kodda var PO'da yok + bos msgstr olanlar)"
            )
            _dosyaya_yaz(
                kok / "po-kaldirilacak.txt",
                ref["kaldirilacaklar"],
                "PO'dan kaldırılabilecek msgid listesi (PO'da var, kodda yok)"
            )
            _dosyaya_yaz(
                kok / "po-eslesenler.txt",
                ref["eslesenler"],
                "Eşleşen msgid listesi (kodda ve PO'da var)"
            )
        except Exception as e:
            logger.error("Çıktı dosyaları yazılamadı: %s", e)

    logger.info(
        "Tüm diller işlendi. Toplam eklenecek: %d | Toplam kaldırılacak: %d",
        toplam_eklenecek, toplam_kaldirilacak,
    )

    if "--rapor" in sys.argv:
        logger.info("--rapor: yalnızca raporlar üretildi.")
        return 0

    # === ADIM 1: KALDIRILACAKLAR ===
    herhangi_kaldirilacak = any(
        len(s.get("kaldirilacaklar", [])) > 0 for s in basarili_sonuclar
    )

    if herhangi_kaldirilacak:
        dil_sayisi = sum(
            1 for s in basarili_sonuclar if len(s.get("kaldirilacaklar", [])) > 0
        )
        soru1 = (
            f"Kaldırılması gerekenler yorum satırı yapılsın mı?\n\n"
            f"({dil_sayisi} dil dizinindeki PO dosyalarına uygulanacak.)"
        )
        if _wx_evet_hayir_sor(soru1, "PO Temizliği"):
            toplam_yorumlanan = 0
            for s in basarili_sonuclar:
                kal = s.get("kaldirilacaklar", [])
                if not kal:
                    continue
                kaldirilacak_set = set(kal)
                po_yolu = s["po_yolu"]
                yorum_sayisi = po_kaldirilacaklari_yorum_satiri_yap(po_yolu, kaldirilacak_set)
                toplam_yorumlanan += yorum_sayisi
                logger.info("[%s] Yorumlanan kayıt: %d", s["dil"], yorum_sayisi)

            logger.info("Toplam yorumlanan kayıt sayısı: %d", toplam_yorumlanan)

            if toplam_yorumlanan > 0:
                time.sleep(2)
                toplam_isaretli = sum(
                    po_isaretli_satir_sayisini_al(s["po_yolu"])
                    for s in basarili_sonuclar
                )
                soru2 = (
                    f"#- ile işaretlenenler .po dosyalarından silinsin mi?\n\n"
                    f"(Toplam {toplam_isaretli} işaretli satır, "
                    f"{dil_sayisi} dil dizininde silinecek.)"
                )
                baslik2 = f"[{toplam_isaretli} satır] PO Temizliği"

                if _wx_evet_hayir_sor(soru2, baslik2):
                    toplam_silinen = 0
                    for s in basarili_sonuclar:
                        silinen = po_isaretlenenleri_sil(s["po_yolu"], bos_satir_normallestir=True)
                        toplam_silinen += silinen
                        logger.info("[%s] Silinen '#-' satır: %d", s["dil"], silinen)
                    logger.info("Toplam silinen '#-' satır sayısı: %d", toplam_silinen)
                else:
                    logger.info("Kullanıcı silme için 'Hayır' dedi.")
            else:
                logger.info("Yorumlanan kayıt yok. Silme sorusu sorulmadı.")
        else:
            logger.info("Kullanıcı 'Hayır' dedi. PO'larda işlem yapılmadı.")
    else:
        logger.info("Kaldırılacak bir kayıt yok. Soru sorulmadı.")

    # === ADIM 2: EKLE.TXT'DEN EKLEME ===
    soru_ekle = "Eksikleri yüklemek istiyor musun?\n\n(ekle.txt dosyasındaki kayıtlar ilgili PO dosyalarına eklenecek.)"
    if not _wx_evet_hayir_sor(soru_ekle, "PO Ekleme"):
        logger.info("Kullanıcı ekleme için 'Hayır' dedi. Çıkılıyor.")
        return 0

    ekle_yolu = ekle_txt_bul(kok)
    if ekle_yolu is None:
        logger.error("ekle.txt bulunamadı. locales/locale dizinleri kontrol edildi.")
        _wx_evet_hayir_sor(
            "ekle.txt bulunamadı.\n\nlocales veya locale dizini altına ekle.txt dosyasını oluşturun.",
            "Dosya Bulunamadı"
        )
        return 1

    logger.info("ekle.txt bulundu: %s", ekle_yolu)
    dil_bloklari = ekle_txt_parse(ekle_yolu)

    if not dil_bloklari:
        logger.info("ekle.txt boş veya dil bloğu yok, ekleme atlanıyor.")

    # PO dosyalarını dil koduna göre eşleştir
    po_harita: Dict[str, Path] = {
        s["dil"]: s["po_yolu"] for s in basarili_sonuclar if not s.get("hata")
    }

    toplam_eklenen = 0
    eslesmeyenler: List[str] = []

    for dil_kodu, satirlar in dil_bloklari.items():
        if dil_kodu not in po_harita:
            logger.warning("ekle.txt'de dil kodu var ama PO dizini yok: %s", dil_kodu)
            eslesmeyenler.append(dil_kodu)
            continue

        po_yolu = po_harita[dil_kodu]
        # Güncel mevcut msgid'leri al (adım 1 sonrası değişmiş olabilir)
        try:
            mevcut_msgidler = msgidleri_topla_podan(po_yolu)
        except Exception as e:
            logger.error("[%s] PO tekrar okunamadı: %s", dil_kodu, e)
            continue

        eklenen = po_ya_ekle(po_yolu, satirlar, mevcut_msgidler)
        toplam_eklenen += eklenen
        logger.info("[%s] Eklenen blok: %d", dil_kodu, eklenen)

    if eslesmeyenler:
        logger.warning("PO dizini bulunamayan dil kodları: %s", ", ".join(eslesmeyenler))

    logger.info("Adım 2 tamamlandı. Toplam eklenen blok: %d", toplam_eklenen)

    # === ADIM 3: YEDEK VE .MO TEMİZLİĞİ ===
    soru_temizlik = (
        "Yedek dosyalar (.bak) ve .mo dosyaları silinsin mi?\n\n"
        "Yalnızca bu script'in altındaki locale/locales dizinlerinde işlem yapılır."
    )
    if _wx_evet_hayir_sor(soru_temizlik, "Temizlik"):
        _yedek_ve_mo_sil(kok)

    # === ADIM 4: ALFABETİK SIRALAMA ===
    soru_alfa = (
        "Tüm PO dosyaları alfabetik sıraya göre yeniden düzenlensin mi?\n\n"
        "Header bloğu korunur, msgid/msgstr blokları msgid'e göre sıralanır.\n"
        "İşlem sonrası po-alfabetik.txt oluşturulur (sadece msgid listesi)."
    )
    if _wx_evet_hayir_sor(soru_alfa, "Alfabetik Sıralama"):
        kok_gercek = kok.resolve()
        tum_msgidler: Set[str] = set()
        basarili_alfa = 0
        hatali_alfa = 0

        for s in basarili_sonuclar:
            po_yolu = s["po_yolu"]
            tamam = po_alfabetik_sirala(po_yolu)
            if tamam:
                basarili_alfa += 1
                # Sıralama sonrası güncel msgid'leri topla (referans dil olarak ilk başarılı)
                if not tum_msgidler:
                    msgidler = po_alfabetik_rapor_yaz(po_yolu, kok, kok_gercek)
                    tum_msgidler.update(msgidler)
            else:
                hatali_alfa += 1
                logger.error("[%s] Alfabetik sıralama başarısız.", s["dil"])

        if tum_msgidler:
            try:
                _dosyaya_yaz(
                    kok / "po-alfabetik.txt",
                    sorted(tum_msgidler),
                    "Tüm msgid listesi (alfabetik sıralı)"
                )
                logger.info("po-alfabetik.txt oluşturuldu. Toplam msgid: %d", len(tum_msgidler))
            except Exception as e:
                logger.error("po-alfabetik.txt yazılamadı: %s", e)

        logger.info(
            "Alfabetik sıralama tamamlandı. Başarılı: %d | Hatalı: %d",
            basarili_alfa, hatali_alfa
        )

    return 0


def _yedek_ve_mo_sil(kok: Path) -> None:
    """
    kok / locales ve kok / locale altındaki .bak ve .mo dosyalarını siler.
    Güvenlik: silinen her dosyanın resolve() ile gerçek yolu kok altında olduğu doğrulanır.
    Üst dizine çıkan veya kok dışında kalan hiçbir dosyaya dokunulmaz.
    """
    kok_gercek = kok.resolve()
    silinen_bak = 0
    silinen_mo = 0
    atlanan = 0

    for klasor_adi in locales_klasor_adlari:
        klasor = kok / klasor_adi
        if not klasor.is_dir():
            continue

        klasor_gercek = klasor.resolve()

        # Güvenlik: klasör kok altında mı?
        try:
            klasor_gercek.relative_to(kok_gercek)
        except ValueError:
            logger.error(
                "GÜVENLİK: locales dizini kok dışında! Atlandı. klasor=%s kok=%s",
                klasor_gercek, kok_gercek
            )
            atlanan += 1
            continue

        for dosya in klasor.rglob("*"):
            if not dosya.is_file():
                continue

            # Güvenlik: her dosyanın gerçek yolunu kontrol et
            try:
                dosya_gercek = dosya.resolve()
                dosya_gercek.relative_to(kok_gercek)
            except ValueError:
                logger.error(
                    "GÜVENLİK: Dosya kok dışında! Atlandı. dosya=%s",
                    dosya
                )
                atlanan += 1
                continue
            except Exception as e:
                logger.error("Dosya yolu çözümlenemedi: %s | Hata: %s", dosya, e)
                atlanan += 1
                continue

            uzanti = dosya.suffix.lower()

            if uzanti == ".mo":
                try:
                    dosya.unlink()
                    silinen_mo += 1
                    logger.info("Silindi (.mo): %s", dosya)
                except Exception as e:
                    logger.error("Silinemedi (.mo): %s | Hata: %s", dosya, e)
                    atlanan += 1

            elif uzanti == ".bak":
                try:
                    dosya.unlink()
                    silinen_bak += 1
                    logger.info("Silindi (.bak): %s", dosya)
                except Exception as e:
                    logger.error("Silinemedi (.bak): %s | Hata: %s", dosya, e)
                    atlanan += 1

    logger.info(
        "Temizlik tamamlandı. Silinen .bak: %d | Silinen .mo: %d | Atlanan: %d",
        silinen_bak, silinen_mo, atlanan
    )


if __name__ == "__main__":
    raise SystemExit(calistir())

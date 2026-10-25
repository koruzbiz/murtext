# -*- coding: utf-8 -*-
# Plug-in version: 3.0.3
# MurText compatible version: 3.0.3

# Standart modüller / Standard Modules
import os
import time
import subprocess
import webbrowser
import winreg
import shutil
import wx

# Log kaydı yapmak isterseniz 'logger_pz' değerini 'True' yapın. / If you want to log, set the 'logger_pz' value to 'True'.
logger_pz = False

from ._log import baslat_loglama
LOG_DIZINI = os.path.dirname(os.path.abspath(__file__))
logger = baslat_loglama(
    appdata_dir=LOG_DIZINI,
    eklenti_adi="MurText",
    stdout_yonlendir=False,
    aktif=logger_pz,
    excepthook_kur=logger_pz
)

# NVDA modules
# import gui
import speech
import api
import ui
from scriptHandler import script
from config import conf
from keyboardHandler import KeyboardInputGesture as KIG

# Gettext
from . import tr
from .settings import SECTION, KEY_AV_TASK, KEY_AV_TASK_USER, KEY_OUTPUT_DIR, KEY_COPY_KEY, KEY_DOWNLOAD_KEY, _get_documents_dir, _find_copy, _find_download

# GlobalPlugin alias 
import globalPluginHandler
_BaseGlobalPlugin = getattr(globalPluginHandler, "GlobalPlugin", None)

# NVDA rol sabitleri / NVDA ROLE STABILITIES
try:
    from controlTypes import Role
    ROLE_POPUPMENU = Role.POPUPMENU
    ROLE_MENU = Role.MENU
except Exception:
    Role = None
    ROLE_POPUPMENU = ROLE_MENU = None

# Proje sabitleri / Project strings
ALLOWED_EXTS = (".opus", ".mp3", ".mp4", ".m4a", ".mpeg", ".aac", ".flac", ".ogg", ".wav", ".wma", ".dat", ".waptt", ".avi", ".flv", ".m4v", ".mkv", ".mov", ".mpg", ".ts")
MurText_path = os.path.join(os.environ.get("LOCALAPPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Local")), "Koruz_Biz", "MurText", "MurText.exe")
MurText_INSTALLED = False
APP_WhatsApp = "WhatsApp"
APP_DESKTOP  = "desktop"
APP_EXPLORER = "explorer"
APP_UNKNOWN  = "unknown"
APP_BROWSER = "browser"
TARAYICILAR = {
    "chrome": ("LOCALAPPDATA", ("Google", "Chrome", "User Data")),
    "msedge": ("LOCALAPPDATA", ("Microsoft", "Edge", "User Data")),
    "brave": ("LOCALAPPDATA", ("BraveSoftware", "Brave-Browser", "User Data")),
    "vivaldi": ("LOCALAPPDATA", ("Vivaldi", "User Data")),
    "opera": ("APPDATA", ("Opera Software", "Opera Stable")),
}
INDIRME_BEKLEME_SN = 60
INDIRME_YOKLAMA_MS = 500
INDIRME_PENCERE_ONCE_SN = 2
YARIM_UZANTILAR = (".crdownload", ".tmp", ".part", ".opdownload", ".download")
_indirme_jetonu = 0

def get_output_dir() -> str:
    """Ayarlar > MurText'te seçilen klasörü döndürür; yoksa Windows 'Belgelerim' klasörünü verir."""
    try:
        p = conf.get(SECTION, {}).get(KEY_OUTPUT_DIR)
    except Exception:
        p = None
    return p or _get_documents_dir()

# g=2 Masaüstü / Desktop 
def MurText_is_desktop_context():
    """Masaüstü (Explorer'ın Desktop yüzü) mü?"""
    try:
        obj = api.getForegroundObject()
        app_name = str(getattr(getattr(obj, "appModule", None), "appName", "")).lower()
        window_class = str(getattr(obj, "windowClassName", "")).lower()
        name = str(getattr(obj, "name", "")).lower()

        logger.info(f"[Ctx/Desktop] app={app_name}, class={window_class}, name={name}")

        # Explorer uygulaması ve Desktop göstergeleri
        if app_name == "explorer" and (
            "desktop" in name or "masaüstü" in name or window_class in ("progman", "folderview")
        ):
            return True

        return False
    except Exception as e:
        logger.error(f"[Ctx/Desktop] f: MurText_is_desktop_context. {e}")
        return False

def _MurText_get_real_desktop():
    """Masaüstü taşınmış olsa bile gerçek yolunu döndür."""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            val, _ = winreg.QueryValueEx(key, "Desktop")
            path = os.path.expandvars(val)
            if os.path.isdir(path):
                #! logger.info(f"[Desktop] Reg Desktop: {path}")
                return path
    except Exception as e:
        logger.error(f"[Desktop] _MurText_get_real_desktop Reg okunamadı: {e}")

    # Fallback'lar
    home = os.path.expanduser("~")
    cand = os.path.join(home, "Desktop")
    if os.path.isdir(cand):
        logger.info(f"[Desktop] Fallback: {cand}")
        return cand
    od = os.path.join(home, "OneDrive", "Desktop")
    if os.path.isdir(od):
        logger.info(f"[Desktop] OneDrive Fallback: {od}")
        return od

    logger.info("[Desktop] F: _MurText_get_real_desktop Masaüstü bulunamadı")
    return None

def _MurText_try_append_allowed_exts(base_without_ext):
    """Uzantı gizlenmişse izinli uzantıları deneyip var olanı döndür."""
    for ext in ALLOWED_EXTS:
        cand = base_without_ext + ext
        if os.path.isfile(cand):
            logger.info(f"[Desktop] Uzantı tahmini tuttu: {cand}")
            return cand
    return None

def _MurText_resolve_shortcut_if_needed(path):
    """'.lnk' ise gerçek hedefi döndür; değilse olduğu gibi ver."""
    try:
        if path and path.lower().endswith(".lnk"):
            import win32com.client  # pywin32
            shell = win32com.client.Dispatch("WScript.Shell")
            target = shell.CreateShortcut(path).Targetpath
            if target and os.path.exists(target):
                logger.info(f"[Desktop] Kısayol hedefi: {target}")
                return target
    except Exception as e:
        logger.error(f"[Desktop] f: _MurText_resolve_shortcut_if_needed Kısayol çözülemedi: {e}")
    return path

def _MurText_get_selected_file_desktop():
    """
    Masaüstünde seçili dosyanın yolunu tahmin eder.
    NVDA'nın 'navigator object' adını kullanır.
    """
    try:
        obj = api.getNavigatorObject()
        name = (getattr(obj, "name", None) or "").strip()
        desktop = _MurText_get_real_desktop()
        logger.info(f"[Desktop] navigator.name='{name}', desktop='{desktop}'")

        if not name or not desktop:
            return None

        # 1) Tam isimle dene
        cand = os.path.join(desktop, name)
        if os.path.isfile(cand):
            return _MurText_resolve_shortcut_if_needed(cand)

        # 2) Uzantı gizli olabilir: 'Dosya' -> Dosya.mp3/.wav... gibi
        no_ext = os.path.join(desktop, os.path.splitext(name)[0])
        guessed = _MurText_try_append_allowed_exts(no_ext)
        if guessed:
            return _MurText_resolve_shortcut_if_needed(guessed)

        # 3) Olmazsa None
        logger.info("[Desktop] f: _MurText_get_selected_file_desktop Seçili dosya bulunamadı.")
        return None
    except Exception as e:
        logger.error(f"[Desktop] f:_MurText_get_selected_file_desktop {e}")
        return None

# g=3 Dosya Gezgini / Explorer
def MurText_is_explorer_context():
    """Dosya Gezgini (klasör penceresi) mi? (Masaüstü hariç)"""
    try:
        # Masaüstünü dışla
        if MurText_is_desktop_context():
            return False

        obj = api.getForegroundObject()
        app_name = str(getattr(getattr(obj, "appModule", None), "appName", "")).lower()
        window_class = str(getattr(obj, "windowClassName", "")).lower()
        name = str(getattr(obj, "name", "")).lower()

        logger.info(f"[Ctx/Explorer] app={app_name}, class={window_class}, name={name}")

        if app_name == "explorer":
            return True
        if window_class in ("cabinetwclass", "explorer"):
            return True
        # Yerelleştirilmiş başlıklar
        if "dosya gezgini" in name or "file explorer" in name:
            return True

        return False
    except Exception as e:
        logger.error(f"[Ctx/Explorer] f: MurText_is_explorer_context {e}")
        return False

def MurText_get_selected_file_explorer():
    """Sadece ÖN PLANDAKİ (aktif) Explorer penceresinden seçili dosyanın tam yolunu alır.
    Seçim yoksa klasör yolunu döndürür; bulunamazsa None.
    """
    try:
        import comtypes.client
        # Ön plan pencere HWNDi
        try:
            from winUser import getForegroundWindow  # NVDA'nın kendi modülü
            fg_hwnd = int(getForegroundWindow())
        except Exception:
            import ctypes
            fg_hwnd = int(ctypes.windll.user32.GetForegroundWindow())
        logger.info(f"[Explorer] FG HWND: {fg_hwnd}")

        shell = comtypes.client.CreateObject("Shell.Application")

        # Sadece FG (ön plandaki) Explorer penceresi
        for w in shell.Windows():
            try:
                w_hwnd = int(getattr(w, "HWND", 0))
                w_name = str(getattr(w, "Name", ""))
                logger.info(f"[Explorer] window: hwnd={w_hwnd} name={w_name!r}")
                if w_hwnd != fg_hwnd:
                    continue
                doc = getattr(w, "Document", None)
                if not doc:
                    logger.info("[Explorer] FG: Document yok")
                    break
                # Seçim var mı?
                try:
                    sel = doc.SelectedItems()
                    if sel and getattr(sel, "Count", 0) > 0:
                        p = sel.Item(0).Path
                        logger.info(f"[Explorer] Seçili (FG): {p}")
                        return p
                except Exception as e_sel:
                    logger.error(f"[Explorer] FG: SelectedItems hatası: {e_sel}")
                # Seçim yoksa klasör yolu
                try:
                    folderPath = doc.Folder.Self.Path
                    logger.info(f"[Explorer] Seçim yok, klasör yolu (FG): {folderPath}")
                    return folderPath
                except Exception as e_fold:
                    logger.error(f"[Explorer] FG: Folder.Path hatası: {e_fold}")
                break  # FG bulundu; daha ötesine bakmaya gerek yok
            except Exception as e_loop:
                logger.error(f"[Explorer] FG döngü hatası: {e_loop}")

        logger.info("[Explorer] Başarısız: Seçili dosya bulunamadı (FG).")
        return None

    except Exception as e:
        logger.error(f"[Explorer] COM API hatası. f:MurText_get_selected_file_explorer {e}")
        # COM başarısızsa PowerShell fallback (son çare)
        try:
            ps_cmd = r'''powershell -command "& { $sel = (New-Object -ComObject Shell.Application).Windows() | Where-Object { $_.Document.SelectedItems().Count -gt 0 } | ForEach-Object { $_.Document.SelectedItems().Item(0).Path }; Write-Output $sel }"'''
            result = subprocess.check_output(ps_cmd, shell=True, universal_newlines=True).strip()
            logger.info(f"[Explorer] PowerShell sonucu: {result}")
            return result if result else None
        except Exception as e2:
            logger.error(f"[Explorer] PowerShell hatası. F: MurText_get_selected_file_explorer {e2}")
            return None

def MurText_get_selected_file():
    """Bağlama göre dosya yolunu alır (Explorer için)."""
    try:
        ctx = MurText_which_app()
        if ctx == APP_EXPLORER:
            return MurText_get_selected_file_explorer()
        logger.info(f"[get_selected_file] Hata: Bağlam desteklenmiyor. f: MurText_get_selected_file {ctx}")
        return None
    except Exception as e:
        logger.error(f"[get_selected_file] {e}")
        return None

# g=4 WhatsApp 
def MurText_is_WhatsApp_context():
    """Ön plandaki pencere WhatsApp mı? (Microsoft Store/Desktop sürümleriyle uyumlu)"""
    try:
        obj = api.getForegroundObject()
        app_name = str(getattr(getattr(obj, "appModule", None), "appName", "")).lower()
        window_class = str(getattr(obj, "windowClassName", "")).lower()
        role = str(getattr(obj, "role", "")).lower()
        name = str(getattr(obj, "name", "")).lower()

        logger.info(f"[Ctx/WA] app={app_name}, class={window_class}, role={role}, name={name}")

        # En basit ve en sağlam eşleşmeler:
        if "whatsapp" in app_name or "whatsapp" in window_class or "whatsapp" in name:
            return True

        return False
    except Exception as e:
        logger.error(f"[Ctx/WA] f: MurText_is_WhatsApp_context {e}")
        return False

# WhatsApp Yardımcıları
def _MurText_n(metin):
    """Karşılaştırma için metni normalize eder (boşluk, büyük/küçük harf ve Türkçe i/İ/ı farkı yok sayılır)."""
    try:
        if metin is None:
            return ""
        t = " ".join(str(metin).split()).strip()
        t = t.replace("İ", "i").replace("I", "i").replace("ı", "i")
        return t.casefold().replace("\u0307", "")
    except Exception:
        return ""

def _MurText_tarama_kipi_acik_mi():
    try:
        ti = getattr(api.getFocusObject(), "treeInterceptor", None)
        return ti is not None and not bool(getattr(ti, "passThrough", False))
    except Exception as e:
        logger.error(f"[Kip] Kip okunamadı: {e}")
        return False

_sessiz_onceki_mod = None
_sessiz_zamanlayici = None
SESSIZ_EN_FAZLA_MS = 3500
BEKLE_MESAJI_MS = 100

# Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
def _MurText_sustur(*_):
    try:
        speech.cancelSpeech()
    except Exception:
        pass

# Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
def _MurText_sessiz_baslat():
    """Menü gezintisi boyunca NVDA konuşmasını kapatır (hece/duyuru kalabalığını önler); en geç SESSIZ_EN_FAZLA_MS sonra konuşma geri yüklenir."""
    global _sessiz_onceki_mod, _sessiz_zamanlayici
    try:
        if _sessiz_onceki_mod is None:
            _sessiz_onceki_mod = speech.getState().speechMode
            speech.setSpeechMode(speech.SpeechMode.off)
            logger.info("[Sessiz] Konuşma kapatıldı.")
    except Exception as e:
        _sessiz_onceki_mod = None
        logger.error(f"[Sessiz] Konuşma kapatılamadı, yalnızca susturma kullanılacak: {e}")
    if _sessiz_zamanlayici is not None:
        try:
            _sessiz_zamanlayici.Stop()
        except Exception:
            pass
    _sessiz_zamanlayici = wx.CallLater(SESSIZ_EN_FAZLA_MS, _MurText_sessiz_bitir)

# Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
def _MurText_sessiz_bitir(*_):
    """Konuşmayı eski haline döndürür; kapatılmadıysa hiçbir şey yapmaz."""
    global _sessiz_onceki_mod, _sessiz_zamanlayici
    try:
        if _sessiz_zamanlayici is not None:
            try:
                _sessiz_zamanlayici.Stop()
            except Exception:
                pass
            _sessiz_zamanlayici = None
        if _sessiz_onceki_mod is not None:
            onceki, _sessiz_onceki_mod = _sessiz_onceki_mod, None
            speech.setSpeechMode(onceki)
            logger.info("[Sessiz] Konuşma geri açıldı.")
    except Exception as e:
        _sessiz_onceki_mod = None
        logger.error(f"[Sessiz] Konuşma geri açılamadı, varsayılan moda dönülüyor: {e}")
        try:
            speech.setSpeechMode(speech.SpeechMode.talk)
        except Exception as e2:
            logger.error(f"[Sessiz] Varsayılan konuşma moduna dönülemedi: {e2}")

def _MurText_mesaj(metin, gecikme_ms=450):
    """Mesajı gecikmeli okutur; konuşma kapalıysa önce açar, WhatsApp/NVDA'nın bastırdığı konuşmayı keser."""
    def _calistir():
        try:
            _MurText_sessiz_bitir()
            _MurText_sustur()
            ui.message(metin)
        except Exception:
            pass
    try:
        wx.CallLater(int(gecikme_ms), _calistir)
    except Exception:
        _calistir()

def _MurText_nearest_menu_root(obj):
    node = obj
    while node:
        role = getattr(node, "role", None)
        # Role sabitleri yoksa da çalışsın
        if (ROLE_POPUPMENU and role == ROLE_POPUPMENU) or (ROLE_MENU and role == ROLE_MENU):
            return node
        node = getattr(node, "parent", None)
    return None

def MurText_WhatsApp():
    """Panodaki dosya yolunu alır ve doğrudan MurText ile açar."""
    try:
        logger.info("MurText_WhatsApp tetiklendi")

        ps_script = (
            "[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false); "
            "Get-Clipboard -Format FileDropList | ForEach-Object { $_.FullName }"
        )
        # Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
        speech.cancelSpeech()

        # DİKKAT: shell=False ve argüman listesi
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            shell=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            creationflags=0x08000000,
        )

        if result.returncode != 0:
            logger.info(f"PS hata: rc={result.returncode} err={result.stderr!r}")
            #! "Panodan dosya alınamadı."
            _MurText_sessiz_bitir()
            ui.message(tr("Failed to retrieve file from clipboard."))
            return

        output = (result.stdout or "").strip()
        logger.info(f"PowerShell clipboard sonucu (raw utf8): {output!r}")

        candidate = ""
        if output:
            for line in (l.strip() for l in output.splitlines() if l.strip()):
                p = os.path.normpath(line)
                # Uzun yol emniyeti (nadir ama dursun)
                if len(p) >= 240 and not p.startswith("\\\\?\\"):
                    p_long = "\\\\?\\" + p
                else:
                    p_long = p

                logger.info(f"Kontrol edilen yol: {p!r}")

                if os.path.isfile(p) or os.path.isfile(p_long):
                    candidate = p
                    break

        if not candidate:
            #! "Panodan dosya alınamadı."
            _MurText_sessiz_bitir()
            ui.message(tr("Failed to retrieve file from clipboard."))
            return

        try:
            import ctypes
            ctypes.windll.user32.OpenClipboard(0)
            ctypes.windll.user32.EmptyClipboard()
            ctypes.windll.user32.CloseClipboard()
            logger.info("[WA] Pano temizlendi.")
        except Exception as e:
            logger.error(f"[WA] Pano temizlenemedi: {e}")

        try:
            time.sleep(0.1)
        except Exception:
            pass
        MurText_open(file_path=candidate, source=APP_WhatsApp)

    except Exception as e:
        _MurText_sessiz_bitir()
        #! "Bir hata oluştu."
        ui.message(tr("An error occurred."))
        logger.info(f"MurText_WhatsApp {e}")

# g=5 genel
def get_murtext_exe_path():
    """
    Sadece HKCU App Paths MurText.exe altındaki Default değeri varsa onu döndürür.
    Yoksa varsayılan MurText_path'ı döndürür.
    (Basit ve tek amaçlı: for/loop/scan yok.)
    """
    subkey = r"Software\Microsoft\Windows\CurrentVersion\App Paths\MurText.exe"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey, 0, winreg.KEY_READ) as k:
            try:
                val = winreg.QueryValue(k, None)  # Default değeri oku
            except OSError:
                val = None
            if val and isinstance(val, str):
                return os.path.normpath(val.strip('"'))
    except OSError:
        pass
    return os.path.normpath(MurText_path)

def MurText_which_app():
    """Ön plandaki uygulamayı ayıklar."""
    try:
        if MurText_is_WhatsApp_context():
            logger.info("[Ctx] Tespit: WhatsApp")
            return APP_WhatsApp
        if MurText_is_desktop_context():
            logger.info("[Ctx] Tespit: Masaüstü")
            return APP_DESKTOP
        if MurText_is_explorer_context():
            logger.info("[Ctx] Tespit: Gezgini")
            return APP_EXPLORER

        logger.info("[Ctx] Tespit: Unknown")
        return APP_UNKNOWN
    except Exception as e:
        logger.error(f"[Ctx] f: MurText_which_app {e}")
        return APP_UNKNOWN

def MurText_get_selected_file_smart():
    """
    Masaüstü/Explorer bağlamına göre seçili dosyayı döndürür.
    - Masaüstündeysek: _MurText_get_selected_file_desktop()
    - Değilsek (Explorer ise): MurText_get_selected_file_explorer()
    """
    try:
        if MurText_is_desktop_context():
            logger.info("[Smart] Bağlam: Masaüstü")
            return _MurText_get_selected_file_desktop()

        if MurText_is_explorer_context():
            logger.info("[Smart] Bağlam: Gezgini")
            return MurText_get_selected_file_explorer()

        logger.info("[Smart] Bağlam desteklenmiyor")
        return None
    except Exception as e:
        logger.error(f"[Smart] f: MurText_get_selected_file_smart {e}")
        return None

def file_control(file_path):
    """
    Dosyanın MurText tarafından işlenebilir olup olmadığını kontrol eder.
    Geri dönüş, ileriye dönük genişletmeye uygun, yapılandırılmış bir sonuçtur.

    Returns:
        dict: {
          "ok": bool,                # True: destekleniyor ve mevcut
          "file_path": str|None,     # Tam yol
          "ext": str|None,           # '.mp3' gibi (küçük harf)
          "reason": str|None         # 'missing', 'not_exists', 'unsupported' vb.
        }
    """

    if not file_path:
        return {"ok": False, "file_path": None, "ext": None, "reason": "missing"}

    # Normalize
    file_path = os.path.abspath(file_path)
    _, ext = os.path.splitext(file_path.lower())

    if not os.path.exists(file_path):
        return {"ok": False, "file_path": file_path, "ext": ext or None, "reason": "not_exists"}

    if ext not in ALLOWED_EXTS:
        return {"ok": False, "file_path": file_path, "ext": ext or None, "reason": "unsupported"}

    return {"ok": True, "file_path": file_path, "ext": ext, "reason": None}

def MurText_dosyayi_kaydet(file_path, tasi=False):
    """Dosyayı ayarlardaki kayıt klasörüne kopyalar (tasi=True ise taşır) ve kullanıcıyı bilgilendirir."""
    try:
        os.makedirs(get_output_dir(), exist_ok=True)
        dest_file = os.path.join(get_output_dir(), os.path.basename(file_path))
        if tasi:
            shutil.move(file_path, dest_file)
        else:
            shutil.copy2(file_path, dest_file)
        #! "Dosya MurText ile kaydedildi."
        _MurText_mesaj(tr("The file was saved with MurText."))
        logger.info(f"Dosya kaydedildi | src={file_path} -> dest={dest_file}")
    except Exception as e:
        #! "Dosya kaydedilemedi."
        ui.message(tr("The file could not be saved."))
        logger.error(f"Dosya kaydedilemedi | src={file_path} | hata={e}")

def Unputable_File(source, file_path, ext):
    """Desteklenmeyen uzantı: WhatsApp'tan geldiyse klasöre kaydeder, değilse bilgi verir."""
    if source in (APP_WhatsApp, APP_BROWSER):
        MurText_dosyayi_kaydet(file_path, tasi=(source == APP_BROWSER))
        return
    #! "Seçilen öğe MurText tarafından desteklenmiyor."
    ui.message(tr("The selected item is not supported by MurText."))
    logger.info(f"Desteklenmeyen uzantı | ext={ext} | path={file_path} | source={source}")

def _MurText_indirilenler_dizini():
    """Windows'un (taşınmış olsa bile) gerçek İndirilenler klasörünü verir."""
    try:
        import ctypes
        import uuid

        class _GUID(ctypes.Structure):
            _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort), ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

        u = uuid.UUID("374DE290-123F-4565-9164-39C4925E467B")
        g = _GUID(u.time_low, u.time_mid, u.time_hi_version, (ctypes.c_ubyte * 8)(*u.bytes[8:]))
        yol = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(yol)) == 0:
            sonuc = yol.value
            ctypes.windll.ole32.CoTaskMemFree(yol)
            if sonuc and os.path.isdir(sonuc):
                return sonuc
    except Exception as e:
        logger.error(f"[İndirme] Bilinen klasör okunamadı: {e}")
    varsayilan = os.path.join(os.path.expanduser("~"), "Downloads")
    return varsayilan if os.path.isdir(varsayilan) else None

def _MurText_tarayici_indirme_dizini(uygulama):
    """Tarayıcının ayarlarında seçilmiş indirme klasörünü verir; yoksa None."""
    try:
        import json
        degisken, parcalar = TARAYICILAR[uygulama]
        kok = os.path.join(os.environ.get(degisken, ""), *parcalar)
        profil = "Default"
        try:
            with open(os.path.join(kok, "Local State"), encoding="utf-8") as f:
                profil = json.load(f).get("profile", {}).get("last_used") or "Default"
        except Exception:
            pass
        for aday in (os.path.join(kok, profil, "Preferences"), os.path.join(kok, "Preferences")):
            if os.path.isfile(aday):
                with open(aday, encoding="utf-8") as f:
                    klasor = json.load(f).get("download", {}).get("default_directory")
                if klasor and os.path.isdir(klasor):
                    return klasor
                return None
    except Exception as e:
        logger.error(f"[İndirme] Tarayıcı ayarı okunamadı ({uygulama}): {e}")
    return None

def _MurText_indirmeyi_bekle(uygulama, t0):
    """Tıklama anına yakın oluşturulmuş yeni dosyayı tarayıcının indirme klasörlerinde bekler."""
    global _indirme_jetonu
    _indirme_jetonu += 1
    jeton = _indirme_jetonu
    dizinler = []
    for d in (_MurText_tarayici_indirme_dizini(uygulama), _MurText_indirilenler_dizini()):
        if d and d not in dizinler:
            dizinler.append(d)
    logger.info(f"[İndirme] İzlenen klasörler: {dizinler}")
    durum = {"onceki": None}

    def _yokla():
        if jeton != _indirme_jetonu:
            return
        gecen = time.time() - t0
        yarim = False
        hazir = []
        for d in dizinler:
            try:
                for ad in os.listdir(d):
                    yol = os.path.join(d, ad)
                    if not os.path.isfile(yol):
                        continue
                    olusma = os.path.getctime(yol)
                    if olusma < t0 - INDIRME_PENCERE_ONCE_SN:
                        continue
                    if ad.lower().endswith(YARIM_UZANTILAR):
                        yarim = True
                    else:
                        hazir.append((olusma, yol))
            except Exception as e:
                logger.error(f"[İndirme] Klasör okunamadı: {d} | {e}")
        if hazir and not yarim:
            hazir.sort()
            yol = hazir[0][1]
            try:
                boyut = os.path.getsize(yol)
            except Exception:
                boyut = -1
            if boyut > 0 and durum["onceki"] == (yol, boyut):
                logger.info(f"[İndirme] Dosya bulundu: {yol} ({boyut} bayt, {gecen:.1f} sn)")
                MurText_open(yol, APP_BROWSER)
                return
            durum["onceki"] = (yol, boyut)
        if gecen > INDIRME_BEKLEME_SN:
            logger.error(f"[İndirme] Zaman aşımı: {INDIRME_BEKLEME_SN} sn içinde dosya bulunamadı.")
            #! "İndirilen dosya bulunamadı. Tarayıcının indirme konumunu kontrol edin."
            ui.message(tr("The downloaded file was not found. Check the browser's download location."))
            return
        wx.CallLater(INDIRME_YOKLAMA_MS, _yokla)

    wx.CallLater(INDIRME_YOKLAMA_MS, _yokla)

def _MurText_pencereyi_one_cikar(pencere):
    """Pencereyi ön plana getirir (ön plan kilidini aşmak için iş parçacığı bağlama kullanılır)."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        hwnd = pencere.GetHandle()
        pencere.Raise()
        on_plan = user32.GetForegroundWindow()
        on_plan_is = user32.GetWindowThreadProcessId(on_plan, None)
        benim_is = kernel32.GetCurrentThreadId()
        baglandi = bool(on_plan_is and on_plan_is != benim_is and user32.AttachThreadInput(on_plan_is, benim_is, True))
        try:
            user32.ShowWindow(hwnd, 5)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        finally:
            if baglandi:
                user32.AttachThreadInput(on_plan_is, benim_is, False)
        logger.info("[Dialog] Pencere ön plana getirildi.")
    except Exception as e:
        logger.error(f"[Dialog] Pencere ön plana getirilemedi: {e}")

def MurText_gorev_sor():
    """Küçük görev seçim penceresi. (secim, hatirla) döndürür; iptalde None."""
    import gui
    sonuc = {"deger": None}
    dlg = wx.Dialog(gui.mainFrame, title=tr("MurText - Select task type"), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)
    try:
        ana = wx.BoxSizer(wx.VERTICAL)
        satir = wx.BoxSizer(wx.HORIZONTAL)
        satir.Add(wx.StaticText(dlg, label=tr("What would you like to do?")), flag=wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=8)
        secim = wx.Choice(dlg, choices=[tr("Send the file to MurText"), tr("Save the file")])
        secim.SetSelection(0)
        satir.Add(secim)
        ana.Add(satir, flag=wx.ALL, border=12)
        hatirla = wx.CheckBox(dlg, label=tr("Remember my choice"))
        ana.Add(hatirla, flag=wx.LEFT | wx.RIGHT | wx.BOTTOM, border=12)
        dugmeler = dlg.CreateButtonSizer(wx.OK | wx.CANCEL)
        ana.Add(dugmeler, flag=wx.ALL | wx.EXPAND, border=12)
        dlg.SetSizerAndFit(ana)
        dlg.CentreOnScreen()
        secim.SetFocus()
        gui.mainFrame.prePopup()
        wx.CallLater(100, _MurText_pencereyi_one_cikar, dlg)
        try:
            if dlg.ShowModal() == wx.ID_OK:
                sonuc["deger"] = (secim.GetSelection(), hatirla.GetValue())
        finally:
            gui.mainFrame.postPopup()
    finally:
        dlg.Destroy()
    return sonuc["deger"]

def MurText_gorev_belirle(source):
    """Ses/video dosyası WhatsApp'tan geldiyse ayara göre 0 (MurText'e gönder), 1 (kaydet) veya None (iptal) döndürür."""
    if source not in (APP_WhatsApp, APP_BROWSER):
        return 0
    try:
        ayar = int(conf[SECTION].get(KEY_AV_TASK, 2)) if conf[SECTION].get(KEY_AV_TASK_USER) else 2
    except Exception:
        ayar = 2
    logger.info(f"Görev tipi ayarı: {ayar}")
    if ayar in (0, 1):
        return ayar
    cevap = MurText_gorev_sor()
    if cevap is None:
        return None
    secim, hatirla = cevap
    if hatirla:
        try:
            conf[SECTION][KEY_AV_TASK] = secim
            conf[SECTION][KEY_AV_TASK_USER] = 1
            conf.save()
        except Exception as e:
            logger.error(f"Görev tipi kaydedilemedi: {e}")
    return secim

def MurText_open(file_path=None, source=None):
    _MurText_sessiz_bitir()
    try:
        logger.info(f"MurText_open tetiklendi | source: {source}")

        # Dosya yolu belirlenmemişse, kaynak üzerinden alınır
        if file_path is None:
            logger.info(f"Dosya yolu belirtilmedi. Kaynak: {source}")
            if source == APP_DESKTOP:
                file_path = MurText_get_selected_file_smart()
            elif source == APP_EXPLORER:
                file_path = MurText_get_selected_file()

        logger.info(f"Alınan dosya yolu (ham): {file_path}")

        # Merkezî kontrol (varlık + uzantı desteği)
        fc = file_control(file_path)

        if not fc["ok"]:
            reason = fc.get("reason")
            full_path = fc.get("file_path")
            ext = fc.get("ext")

            if reason in ("missing", "not_exists"):
                #! "Geçersiz yordam veya dosya yolu."
                ui.message(tr("Invalid procedure or file path."))
                logger.info(f"Başarısız: Dosya yolu alınamadı veya mevcut değil. path={full_path}")
                return

            if reason == "unsupported":
                # Desteklenmeyen dosyayı Unputable_File'a devret
                _ = Unputable_File(source=source, file_path=full_path, ext=ext)
                return

            # Beklenmeyen durum
            #! "Bir hata oluştu."
            ui.message(tr("An error occurred."))
            logger.info(f"Başarısız: Bilinmeyen kontrol sonucu: reason={reason} | path={full_path} | ext={ext}")
            return

        # Buraya gelindiyse dosya mevcut ve uzantı destekleniyor
        file_path = fc["file_path"]
        file_name = os.path.basename(file_path)
        logger.info(f"Dosya adı: {file_name} | Uzantı: {fc['ext']}")

        gorev = MurText_gorev_belirle(source)
        if gorev is None:
            logger.info("Görev seçimi iptal edildi.")
            return
        if gorev == 1:
            MurText_dosyayi_kaydet(file_path, tasi=(source == APP_BROWSER))
            return

        if not MurText_INSTALLED and not MurText_probe_installation_on_load():
            return

        #! "MurText ile açılıyor. Uygulama hazırlanıyor."
        if source not in (APP_WhatsApp, APP_BROWSER):
            ui.message(tr("Opening with MurText. Preparing the application."))
        subprocess.Popen([get_murtext_exe_path(), file_path])
        logger.info(f"MurText çalıştırıldı: {file_path} -> {get_murtext_exe_path()}")

    except Exception as e:
        #! "Bir hata oluştu."
        ui.message(tr("An error occurred."))
        logger.error(f"MurText_open istisnası: {e}")

def MurText_probe_installation_on_load():
    """
    MurText'e dosya gönderilmeden hemen önce çağrılır.
    Gerçek MurText.exe yolunun (kayıt defteri veya varsayılan) varlığına bakar; MurText_INSTALLED'ı ayarlar.
    Yüklü değilse kurulum diyaloğunu tetikler.
    True/False döndürür.
    """
    global MurText_INSTALLED
    try:
        exe = get_murtext_exe_path()
        MurText_INSTALLED = os.path.isfile(exe)
        logger.info(f"[Probe] MurText var mı? {MurText_INSTALLED} {exe}")
        if not MurText_INSTALLED:
            logger.info("MurText kurulu değil")
            MurText_prompt_to_install_if_missing()
        return MurText_INSTALLED
    except Exception as e:
        MurText_INSTALLED = False
        logger.error("MurText kurulu değil")
        logger.error(f"[Probe] f:MurText_probe_installation_on_load {e}")
        MurText_prompt_to_install_if_missing()
        return False

def MurText_prompt_to_install_if_missing():
    """
    MurText_INSTALLED True değilse çağrılır.
    Basit Yes/No diyalog: Evet -> indirme sayfası, Hayır -> sadece log.
    """
    def _show():
        try:
            # 100 ms sonra metni seslendirt: ekranda pencere varken okumayı tetikler
            t = wx.Timer()
            def _onTimer(evt):
                try:
                     #! "Ücretsiz bir uygulama olan MurText olmadan devam edemezsiniz. İndirmek ister misiniz?"
                    ui.message(tr("You cannot proceed without MurText, a free application. Would you like to download it?"))
                finally:
                    t.Stop()
            t.Bind(wx.EVT_TIMER, _onTimer)
            t.Start(100)

            dlg = wx.MessageDialog(
                None, 
                #! "Ücretsiz bir uygulama olan MurText olmadan devam edemezsiniz. İndirmek ister misiniz?"
                tr("You cannot proceed without MurText, a free application. Would you like to download it?"),
                #! "MurText bulunamadı"
                tr("MurText not found"),
                style=wx.YES_NO | wx.ICON_WARNING
            )
            res = dlg.ShowModal()
            dlg.Destroy()

            logger.info(f"[Prompt] Sonuç id: {res}")

            if res == wx.ID_YES:
                try:
                    webbrowser.open("https://MurText.org?page=download&source=nvda", new=1)
                except Exception as e:
                    logger.error(f"[Prompt] URL açılamadı: {e}")
            elif res == wx.ID_NO:
                logger.info("[Prompt] HAYIR: Kullanıcı reddetti.")
            else:
                logger.info("[Prompt] Kapatıldı / iptal edildi.")
        except Exception as e:
            logger.error(f"[Prompt] pop up : {e}")

    wx.CallAfter(_show)

class GlobalPlugin(_BaseGlobalPlugin):
    def __init__(self):
        super().__init__()
        logger.info("Yüklendi")

    def terminate(self):
        _MurText_sessiz_bitir()
        super().terminate()

    # Girdi hareketleri kategori ata
    scriptCategory = tr("MurText")

    # Varsayılan kısayollar (kullanıcı burada değiştirebilir)
    __gestures = {
        "kb:NVDA+alt+q": "MurText_master",
    }

    @script(
        description="MurText kısayol tuşu",
    )
    def script_MurText_master(self, gesture):
        logger.info("\n#! Tetiklendi !#")

        try:
            ctx = MurText_which_app()
            logger.info(f"[Master] Bağlam: {ctx}")

            if ctx == APP_WhatsApp:
                try:
                    logger.info("[Master] WhatsApp algılandı, klasik menü açma akışı başlatılıyor")
                    self._MurText_open_context_menu()
                    return
                except Exception as e:
                    logger.error(f"[Master] WhatsApp menü açma hatası: {e}")
                    ui.message(tr("Failed to open the menu."))
                return

            if ctx == APP_DESKTOP:
                #! logger.info("[Master] Masaüstü algılandı, MurText_open çağrılıyor")
                MurText_open(source=APP_DESKTOP)
                return

            if ctx == APP_EXPLORER:
                #! logger.info("[Master] Gezginde tetiklendi, MurText_open çağrılıyor")
                MurText_open(source=APP_EXPLORER)
                return

            #! "MurText eklentisi bu uygulama için yapılandırılmamış."
            ui.message(tr("The MurText add-on is not configured for this application."))
            logger.info("[Master] Başarısız: Bağlam desteklenmiyor")

        except Exception as e:
            #! "Uygulama belirlenirken bir hata oluştu."
            ui.message(tr("An error occurred while identifying the application."))
            logger.error(f"[Master] HATA: {e}")

    def _MurText_find_buttons(self, obj, derinlik=0):
        dugmeler = []
        if derinlik > 6:
            return dugmeler
        try:
            from controlTypes import Role
            if getattr(obj, "role", None) == Role.BUTTON:
                dugmeler.append(obj)
        except Exception:
            pass
        try:
            for cocuk in getattr(obj, "children", []) or []:
                dugmeler.extend(self._MurText_find_buttons(cocuk, derinlik + 1))
        except Exception:
            pass
        return dugmeler

    @staticmethod
    def _MurText_collapsed_mi(dugme):
        try:
            return 512 in (getattr(dugme, "states", None) or set())
        except Exception:
            return False

    def _MurText_context_menu_dogrudan_ac(self):
        """Mesajın 'bağlam menüsü' (COLLAPSED) düğmesini bulup çalıştırır.
        Odağın kardeşleri arasında arar."""
        try:
            odak = api.getFocusObject()
            ebeveyn = getattr(odak, "parent", None)
            if not ebeveyn:
                logger.info("[Kopyala] Doğrudan menü açma: ebeveyn yok")
                return False

            kardesler = getattr(ebeveyn, "children", []) or []
            logger.info(f"[Kopyala] Doğrudan menü açma: odak={str(getattr(odak, 'name', None))[:60]!r} ebeveyn={str(getattr(ebeveyn, 'name', None))[:60]!r} kardeş sayısı: {len(kardesler)}")

            for kardes in kardesler:
                for dugme in self._MurText_find_buttons(kardes):
                    if self._MurText_collapsed_mi(dugme):
                        logger.info(f"[Kopyala] Doğrudan menü açma: collapsed düğme bulundu: {getattr(dugme, 'name', None)!r}")
                        dugme.doAction()
                        return True

            logger.info("[Kopyala] Doğrudan menü açma: uygun düğme bulunamadı")
            return False
        except Exception as e:
            logger.error(f"[Kopyala] Doğrudan menü açma hatası: {e}")
            return False

    def _MurText_appmodule_context_menu_cagir(self):
        try:
            odak = api.getFocusObject()
            app_modul = getattr(odak, "appModule", None)
            if not app_modul:
                logger.info("[Kopyala] AppModule bulunamadı")
                return False

            komut = getattr(app_modul, "script_contextMenu", None)
            if not komut:
                logger.info("[Kopyala] AppModule script_contextMenu bulunamadı")
                return False

            logger.info("[Kopyala] AppModule script_contextMenu doğrudan çağrılıyor")
            komut(None)
            return True
        except Exception as e:
            logger.error(f"[Kopyala] AppModule script_contextMenu çağrı hatası: {e}")
            return False

    def _MurText_open_context_menu(self):
        """Tarama kipi açıksa uyarır; değilse 'MurText ile açılıyor' der, ardından konuşmayı kapatıp menü açma işlemini başlatır."""
        if _MurText_tarama_kipi_acik_mi():
            logger.info("[Kip] Tarama kipi açık, kullanıcı uyarıldı.")
            ui.message(tr("Please switch modes first with NVDA+SPACE."))
            return
        try:
            ui.message(tr("Opening with MurText. Preparing the application."))
        except Exception:
            pass
        wx.CallLater(BEKLE_MESAJI_MS, self._MurText_open_context_menu_ic)

    def _MurText_open_context_menu_ic(self):
        """WhatsApp üzerinde bağlam menüsünü nesne tabanlı yollarla açmayı dener."""
        try:
            _MurText_sessiz_baslat()
            self._murtext_tab_sayisi = 0

            def _menu_acildi_mi_ve_kopyala_aransın():
                try:
                    odak = api.getFocusObject()
                    logger.info(f"[Kopyala] Menü açma sonrası odak: name={getattr(odak, 'name', None)!r}, role={getattr(odak, 'role', None)!r}")
                    try:
                        from controlTypes import Role
                        if getattr(odak, "role", None) == Role.TOGGLEBUTTON:
                            logger.info("[Kopyala] Emoji paneli tespit edildi (TOGGLEBUTTON). Tab gönderiliyor.")
                            self._murtext_tab_sayisi = 1
                            KIG.fromName("tab").send()
                            wx.CallLater(200, self._MurText_try_invoke_copy, True, 1)
                            return
                    except Exception as e:
                        logger.error(f"[Kopyala] Emoji panel kontrolü hatası: {e}")
                except Exception as e:
                    logger.error(f"[Kopyala] Menü açma sonrası odak okunamadı: {e}")
                wx.CallLater(120, self._MurText_try_invoke_copy, True, 1)

            try:
                odak = api.getFocusObject()
                logger.info(f"[Kopyala] Menü açma öncesi odak: name={getattr(odak, 'name', None)!r}, role={getattr(odak, 'role', None)!r}")
            except Exception as e:
                logger.error(f"[Kopyala] Odak çözümleme hatası: {e}")

            def _ileri_asama():
                if self._MurText_appmodule_context_menu_cagir():
                    logger.info("[Kopyala] Menü AppModule üzerinden açılmaya çalışıldı")
                    wx.CallLater(220, _menu_acildi_mi_ve_kopyala_aransın)
                    return

                if self._MurText_context_menu_dogrudan_ac():
                    logger.info("[Kopyala] Menü doğrudan doAction ile açılmaya çalışıldı")
                    wx.CallLater(220, _menu_acildi_mi_ve_kopyala_aransın)
                    return

                logger.info("[Kopyala] Menü düğmesi bulunamadı, kullanıcı bilgilendirildi.")
                _MurText_sessiz_bitir()
                _MurText_mesaj(tr("The message's menu button was not found. Move focus to the first line of the message and try again."), 100)

            wx.CallLater(180, _ileri_asama)

        except Exception as e:
            _MurText_sessiz_bitir()
            logger.error(f"[Kopyala] Sağ menü açma hata: {e}")

    def _MurText_etiketler(self, anahtar):
        """Ayarlardaki etiketi (birden fazlaysa ';' ile ayrılmış) normalize edilmiş küme olarak verir."""
        try:
            ham = str(conf[SECTION].get(anahtar, "") or "")
        except Exception:
            ham = ""
        if not ham.strip():
            try:
                ham = str((_find_copy if anahtar == KEY_COPY_KEY else _find_download)() or "")
            except Exception:
                ham = ""
        return set(x for x in (_MurText_n(p) for p in ham.split(";")) if x)

    def _MurText_menu_ogeleri(self, focus, pid, kok, kap=250):
        """Verilen kökten başlayarak (en çok 250 düğüm) menü öğelerini belge sırasıyla toplar."""
        ogeler = []
        yigin = [kok]
        while yigin and len(ogeler) < kap:
            node = yigin.pop()
            if not node:
                continue
            ogeler.append(node)
            try:
                for cocuk in reversed(getattr(node, "children", None) or []):
                    if pid is not None and getattr(cocuk, "processID", None) != pid:
                        continue
                    yigin.append(cocuk)
            except Exception:
                pass
        return ogeler

    def _MurText_menu_yuklendi_mi(self, ogeler):
        """Menü tamamen açıldıysa (emoji paneli hariç en az 4 adlı MENUITEM) True verir."""
        try:
            from controlTypes import Role as _R
            adlar = set(_MurText_n(getattr(o, "name", "")) for o in ogeler if getattr(o, "role", None) == _R.MENUITEM)
            adlar.discard("")
            return len(adlar) >= 4
        except Exception:
            return False

    def _MurText_emoji_panelinde_mi(self, focus, ogeler):
        """Odak emoji (ifade) panelindeyse True verir: odak nesnesi menünün kendisi olabilir, bu yüzden odaklı TOGGLEBUTTON da aranır."""
        try:
            from controlTypes import Role as _R, State as _S
            if getattr(focus, "role", None) in (_R.TOGGLEBUTTON, _R.POPUPMENU) and not _MurText_n(getattr(focus, "name", "")):
                return True
            for o in ogeler:
                if getattr(o, "role", None) == _R.TOGGLEBUTTON and _S.FOCUSED in (getattr(o, "states", None) or ()):
                    return True
        except Exception as e:
            logger.error(f"[Kopyala] Emoji panel kontrolü hatası: {e}")
        return False

    def _MurText_ogeleri_logla(self, ogeler, baslik):
        """Emoji paneli hariç, adı olan menü öğelerini tek satırda loglar."""
        try:
            from controlTypes import Role as _R
            adlar = []
            for o in ogeler:
                if getattr(o, "role", None) in (_R.GRAPHIC, _R.TOGGLEBUTTON):
                    continue
                ad = " ".join(str(getattr(o, "name", "") or "").split())
                if ad:
                    adlar.append(ad[:40])
                if len(adlar) >= 30:
                    break
            logger.info(f"[Kopyala] {baslik}: {len(ogeler)} düğüm | " + " | ".join(dict.fromkeys(adlar)))
        except Exception as e:
            logger.error(f"[Kopyala] Öğe dökümü hatası: {e}")

    def _MurText_try_invoke_copy(self, afterMenu: bool = False, deneme_no: int = 0):
        """Açık menüde 'Kopyala'yı (yoksa 'İndir'i) sıradan bağımsız, ada göre bulur ve çalıştırır."""
        try:
            focus = api.getFocusObject()
            pid = getattr(focus, "processID", None)

            if not afterMenu:
                self._MurText_open_context_menu()
                return

            max_deneme = 4
            kopyala_kume = self._MurText_etiketler(KEY_COPY_KEY)
            indir_kume = self._MurText_etiketler(KEY_DOWNLOAD_KEY)
            logger.info(f"[Kopyala] Deneme {deneme_no}/{max_deneme} | odak: name={getattr(focus, 'name', None)!r} role={getattr(focus, 'role', None)!r} | kopyala={sorted(kopyala_kume)} indir={sorted(indir_kume)}")

            kokler = []
            kok = _MurText_nearest_menu_root(focus)
            if kok:
                kokler.append(("menü kökü", kok))
            ebeveyn = getattr(focus, "parent", None)
            if ebeveyn:
                kokler.append(("ebeveyn", ebeveyn))

            if deneme_no >= 2:
                ust = ebeveyn
                for n in range(2, 5):
                    ust = getattr(ust, "parent", None) if ust else None
                    if not ust:
                        break
                    kokler.append((f"üst{n}", ust))

            kopyala_hedef = indir_hedef = None
            tum_ogeler = []
            for etiket, k in kokler:
                ogeler = self._MurText_menu_ogeleri(focus, pid, k, 400 if etiket.startswith("üst") else 250)
                tum_ogeler.extend(ogeler)
                for o in ogeler:
                    ad = _MurText_n(getattr(o, "name", ""))
                    if not ad:
                        continue
                    if kopyala_hedef is None and ad in kopyala_kume:
                        kopyala_hedef = o
                    elif indir_hedef is None and ad in indir_kume:
                        indir_hedef = o
                if kopyala_hedef is not None or indir_hedef is not None:
                    break

            if kopyala_hedef is not None or indir_hedef is not None or deneme_no >= max_deneme:
                self._MurText_ogeleri_logla(tum_ogeler, "Menü öğeleri")

            if kopyala_hedef is None and indir_hedef is None and self._MurText_emoji_panelinde_mi(focus, tum_ogeler) and (getattr(self, "_murtext_tab_sayisi", 0) == 0 or (getattr(self, "_murtext_tab_sayisi", 0) == 1 and deneme_no >= 2)):
                self._murtext_tab_sayisi = getattr(self, "_murtext_tab_sayisi", 0) + 1
                logger.info(f"[Kopyala] Emoji paneli tespit edildi. Tab gönderiliyor ({self._murtext_tab_sayisi}. kez).")
                try:
                    KIG.fromName("tab").send()
                except Exception as e:
                    logger.error(f"[Kopyala] Tab gönderilemedi: {e}")
                # Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
                _MurText_sustur()
                wx.CallLater(80, _MurText_sustur)
                wx.CallLater(200, _MurText_sustur)
                wx.CallLater(300, self._MurText_try_invoke_copy, True, deneme_no)
                return

            if kopyala_hedef is not None:
                logger.info("[Kopyala] 'Kopyala' bulundu, çalıştırılıyor.")
                if self._MurText_hedefi_calistir(kopyala_hedef):
                    _MurText_sessiz_baslat()
                    wx.CallLater(999, MurText_WhatsApp)
                return

            if indir_hedef is not None:
                logger.info("[Kopyala] 'Kopyala' yok, 'İndir' bulundu, çalıştırılıyor.")
                uygulama = str(getattr(getattr(focus, "appModule", None), "appName", "") or "").lower()
                t0 = time.time()
                if self._MurText_hedefi_calistir(indir_hedef):
                    if uygulama in TARAYICILAR:
                        #! "MurText, indirme işleminin bitmesini bekliyor."
                        _MurText_mesaj(tr("MurText is waiting for the download to finish."), 600)
                        _MurText_indirmeyi_bekle(uygulama, t0)
                    else:
                        _MurText_mesaj(tr("The file is downloading. Please try again in a few seconds."), 600)
                return

            if self._MurText_menu_yuklendi_mi(tum_ogeler):
                logger.info("[Kopyala] Menü yüklü ama Kopyala ve İndir yok. Dosya indiriliyor olabilir; menü kapatılıyor.")
                self._MurText_menuyu_kapat(0)
                _MurText_mesaj(tr("Copy or Download was not found in the menu. The file may not be downloaded yet, or the download may still be in progress; try again in a few seconds. If the problem persists, check the Copy and Download labels in the settings."), 300)
                return

            if deneme_no < max_deneme:
                gecikmeler = {0: 250, 1: 350, 2: 500, 3: 700}
                wx.CallLater(gecikmeler.get(deneme_no, 700), self._MurText_try_invoke_copy, True, deneme_no + 1)
                return

            logger.info("[Kopyala] Ne 'Kopyala' ne 'İndir' bulundu.")
            self._MurText_menuyu_kapat_gerekirse(focus)
            _MurText_sessiz_bitir()
            ui.message(tr("The file has not been downloaded yet or the Copy option is unavailable. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))

        except Exception as e:
            logger.error(f"[Kopyala] Genel hata: {e}")
            _MurText_sessiz_bitir()
            ui.message(tr("Could not click the Copy option. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))

    def _MurText_menuyu_kapat_gerekirse(self, focus):
        """Odak hâlâ bir menünün içindeyse Escape ile kapatır."""
        try:
            from controlTypes import Role as _R
            odak = api.getFocusObject()
            if getattr(odak, "role", None) in (_R.MENUITEM, _R.POPUPMENU, _R.MENU, _R.TOGGLEBUTTON) or _MurText_nearest_menu_root(odak) is not None:
                self._MurText_menuyu_kapat(0)
        except Exception as e:
            logger.error(f"[Kopyala] Menü kapatma kontrolü hatası: {e}")

    def _MurText_menuyu_kapat(self, deneme=0):
        """Menüyü Escape ile kapatır; hâlâ açıksa (en çok 3 kez) yeniden dener."""
        try:
            KIG.fromName("escape").send()
        except Exception as e:
            logger.error(f"[Kopyala] Escape gönderilemedi: {e}")
            return
        if deneme < 2:
            wx.CallLater(300, self._MurText_menu_acik_mi, deneme + 1)

    def _MurText_menu_acik_mi(self, deneme):
        try:
            from controlTypes import Role as _R
            odak = api.getFocusObject()
            if getattr(odak, "role", None) in (_R.MENUITEM, _R.POPUPMENU, _R.MENU, _R.TOGGLEBUTTON) or _MurText_nearest_menu_root(odak) is not None:
                logger.info(f"[Kopyala] Menü hâlâ açık, Escape yeniden gönderiliyor ({deneme + 1}. kez).")
                self._MurText_menuyu_kapat(deneme)
        except Exception as e:
            logger.error(f"[Kopyala] Menü durumu okunamadı: {e}")

    def _MurText_hedefi_calistir(self, hedef):
        try:
            logger.info(f"[Kopyala] Hedef çalıştırılıyor: name={getattr(hedef, 'name', None)!r} role={getattr(hedef, 'role', None)!r}")
            # Odak değişiminde ezilen seslendirmeler rahatsız ediyor, geçici olarak sustur ve sonra daima güvenli şekilde aç / Spoken feedback that gets overridden on focus changes is disruptive; mute it temporarily and always restore it safely
            _MurText_sustur()
            hedef.doAction()
            wx.CallLater(100, _MurText_sustur)
            return True
        except Exception as e:
            logger.error(f"[Kopyala] doAction hatası: {e}")
            try:
                KIG.fromName("enter").send()
                logger.info("[Kopyala] Enter ile ikinci deneme yapıldı")
                return True
            except Exception as e2:
                logger.error(f"[Kopyala] Enter gönderilemedi: {e2}")
                return False

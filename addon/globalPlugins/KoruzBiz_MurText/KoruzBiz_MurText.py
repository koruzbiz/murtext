# -*- coding: utf-8 -*-
# Plug-in version: 2.1
# Murdext compatible version: 2.1.2

# Standart modüller / Standard Modules
import datetime
import os
import time
import subprocess
import winreg as reg
import webbrowser
import winreg
import shutil
import wx
import winUser

# Pasif modüller / Passive modules
# import unicodedata
# import re
# import keyboardHandler
# import controlTypes

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
from ui import message
from scriptHandler import script
from config import conf
from keyboardHandler import KeyboardInputGesture as KIG

# Gettex 
from . import tr

# GlobalPlugin alias 
import globalPluginHandler
_BaseGlobalPlugin = getattr(globalPluginHandler, "GlobalPlugin", None)

# NVDA rol sabitleri / NVDA ROLE STABILITIES
try:
    from controlTypes import Role
    ROLE_POPUPMENU = Role.POPUPMENU
    ROLE_MENU = Role.MENU
    ROLE_MENUITEM = Role.MENUITEM
except Exception:
    Role = None
    ROLE_POPUPMENU = ROLE_MENU = ROLE_MENUITEM = None

# Proje sabitleri / Project strings
ALLOWED_EXTS = (".opus", ".mp3", ".mp4", ".m4a", ".mpeg", ".aac", ".flac", ".ogg", ".wav", ".dat", ".waptt")
MurText_path = os.path.join(os.environ.get("LOCALAPPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Local")), "Koruz_Biz", "MurText", "MurText.exe")
MurText_INSTALLED = False
APP_WhatsApp = "WhatsApp"
APP_DESKTOP  = "desktop"
APP_EXPLORER = "explorer"
APP_UNKNOWN  = "unknown"

def get_output_dir() -> str:
    """Ayarlar > MurText’te seçilen klasörü döndürür.
    Metin girişi kapalı olduğu için burada ~/%ENV% genişletmesi veya klasör oluşturma yapılmaz.
    Değer yoksa Windows 'Belgelerim' (settings._get_documents_dir) döner.
    """
    try:
        p = conf.get(SECTION, {}).get(KEY_OUTPUT_DIR)
    except Exception:
        p = None

    if p:
        return p

    # Ayarlara hiç girilmediyse veya değer yoksa, settings.py'deki belge klasörünü kullan
    try:
        from .settings import _get_documents_dir
        return _get_documents_dir()
    except Exception:
        # En basit güvenli geri dönüş
        return os.path.join(os.path.expanduser("~"), "Documents")

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
        with reg.OpenKey(
            reg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            val, _ = reg.QueryValueEx(key, "Desktop")
            path = os.path.expandvars(val)
            if os.path.isdir(path):
                #- logger.info(f"[Desktop] Reg Desktop: {path}")
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
                    logger.v(f"[Explorer] FG: Folder.Path hatası: {e_fold}")
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
def _MurText_safe(s):
    try:
        return str(s).strip()
    except Exception:
        return ""

def _MurText_is_WhatsApp_obj(obj, target_pid=None):
    try:
        app_name = _MurText_safe(getattr(getattr(obj, "appModule", None), "appName", ""))
        if app_name.lower() == "WhatsApp":
            return True
    except Exception:
        pass
    try:
        if target_pid is not None and getattr(obj, "processID", None) == target_pid:
            return True
    except Exception:
        pass
    return False

def _MurText_nearest_menu_root(obj):
    node, prev = obj, None
    while node and node != prev:
        role = getattr(node, "role", None)
        # Role sabitleri yoksa da çalışsın
        if (ROLE_POPUPMENU and role == ROLE_POPUPMENU) or (ROLE_MENU and role == ROLE_MENU):
            return node
        prev = node
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
        speech.cancelSpeech()                 

        # DİKKAT: shell=False ve argüman listesi
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            shell=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

        if result.returncode != 0:
            logger.info(f"PS hata: rc={result.returncode} err={result.stderr!r}")
            #! "Panodan dosya alınamadı."
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
            ui.message(tr("Failed to retrieve file from clipboard."))
            return

        try:
            time.sleep(0.1)
        except Exception:
            pass
        MurText_open(file_path=candidate, source=APP_WhatsApp)

    except Exception as e:
        #! "Bir hata oluştu."
        ui.message(tr("An error occurred."))
        logger.info(f"MurText_WhatsApp {e}")

# g=5 genel
def get_murtext_exe_path():
    """
    Sadece HKCU\App Paths\MurText.exe altındaki Default değeri varsa onu döndürür.
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
        obj = api.getForegroundObject()
        app_name = str(getattr(getattr(obj, "appModule", None), "appName", "")).lower()
        window_class = str(getattr(obj, "windowClassName", "")).lower()
        role = str(getattr(obj, "role", "")).lower()
        name = str(getattr(obj, "name", "")).lower()
        #- logger.info(f"[Ctx] app={app_name}, class={window_class}, role={role}, name={name}")

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

def Unputable_File(source, file_path, ext):
    """
    Desteklenmeyen dosya senaryosunu ele alır.
    - WhatsApp'tan geldiyse: save_path içine kopyalar ve kullanıcıya bildirir.
    - Explorer/Desktop ise: sadece kullanıcıya desteklenmediğini söyler.
    Genişletmeye uygun, yapılandırılmış bir sonuç döndürür.

    Returns:
        dict: {
          "handled": bool,         # Akış başarıyla işlendi mi
          "saved": bool|None,      # WhatsApp senaryosunda kopyalama yapıldıysa True/False, diğerlerinde None
          "dest": str|None,        # Kopyalanan hedef yol (varsa)
          "source": str,
          "file_path": str,
          "ext": str
        }
    """
    result = {
        "handled": False,
        "saved": None,
        "dest": None,
        "source": source,
        "file_path": file_path,
        "ext": ext,
    }

    try:
        if source == "WhatsApp":
            try:
                os.makedirs(get_output_dir(), exist_ok=True)
                dest_file = os.path.join(get_output_dir(), os.path.basename(file_path))
                shutil.copy2(file_path, dest_file)

                #! "Dosya MurText ile kaydedildi."
                ui.message(tr("The file was saved with MurText."))
                logger.info(f"Unputable_File: WhatsApp kaydı başarılı | src={file_path} -> dest={dest_file} | ext={ext}")
                result.update({"handled": True, "saved": True, "dest": dest_file})
            except Exception as copy_err:
                #! "Dosya kaydedilemedi."
                ui.message(tr("The file could not be saved."))
                logger.error(f"Unputable_File: WhatsApp kaydı HATASI | src={file_path} | hata={copy_err}")
                result.update({"handled": True, "saved": False})
        else:
            # Explorer/Desktop vb.
            #! "Seçilen öğe MurText tarafından desteklenmiyor."
            ui.message(tr("The selected item is not supported by MurText."))
            logger.info(f"Unputable_File: Desteklenmeyen uzantı | ext={ext} | path={file_path} | source={source}")
            result.update({"handled": True})
    except Exception as e:
        logger.error(f"Unputable_File: İstisna: {e}")

    return result

def MurText_open(file_path=None, source=None):
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

        #! "MurText ile açılıyor. Uygulama hazırlanıyor."
        ui.message(tr("Opening with MurText. Preparing the application."))
        subprocess.Popen([get_murtext_exe_path(), file_path])
        logger.info(f"MurText çalıştırıldı: {file_path} -> {get_murtext_exe_path()}")

    except Exception as e:
        #! "Bir hata oluştu."
        ui.message(tr("An error occurred."))
        logger.error(f"MurText_open istisnası: {e}")

def MurText_probe_installation_on_load():
    """
    NVDA eklentisi yüklenirken veya ilk tetikte çağrılır.
    MurText_path var mı diye bakar; sonuca göre MurText_INSTALLED ayarlanır.
    Eğer yüklü değilse (False) debug log yazar ve kurulum diyaloğunu tetikler.
    True/False döndürür.
    """
    global MurText_INSTALLED
    try:
        exists = os.path.isfile(MurText_path)
        MurText_INSTALLED = bool(exists)
        logger.info(f"[Probe] MurText var mı? {MurText_INSTALLED} {MurText_path}")
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
        logger.info(f"Yüklendi")

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

        # Sadece tutucu false ise 
        if not MurText_INSTALLED:
            #- logger.info("Varlık kontrol ediliyor...")
            if not MurText_probe_installation_on_load():
                # Kurulu değil -> pencere açıldı, işi kesiyoruz
                return
        try:
            ctx = MurText_which_app()
            logger.info(f"[Master] Bağlam: {ctx}")

            if ctx == APP_WhatsApp:
                #- logger.info("[Master] WhatsApp algılandı, menü açılacak ve tetikleme yapılacak")
                try:
                    # Bağlam menüsü: Shift+F10'u 'kb' jesti olarak gönder
                    try:
                        KIG.fromName("shift+f10").send()
                    except Exception as e:
                        logger.error(f"[Master] Shift+F10 gönderilemedi: {e}")
                        try:
                            KIG.fromName("applications").send()
                            logger.info("[Master] Uygulama tusu (applications) gonderildi")
                        except Exception as e2:
                            logger.error(f"[Master] Uygulama tusu da gonderilemedi: {e2}")

                    # Menü render olsun; sonra 'Kopyala' taraması
                    wx.CallLater(1000, self._MurText_try_invoke_copy)
                    return

                    #- logger.info("[Master] Menü açıldı ve Insert+Shift+K gönderildi")
                except Exception as e:
                    logger.error(f"[Master] Tuş gönderimi hatası: {e}")
                    #! "Menü açma işlemi başarısız."
                    ui.message(tr("Failed to open the menu."))
                return

            if ctx == APP_DESKTOP:
                #- logger.info("[Master] Masaüstü algılandı, MurText_open çağrılıyor")
                MurText_open(source=APP_DESKTOP)
                return

            if ctx == APP_EXPLORER:
                #- logger.info("[Master] Gezginde tetiklendi, MurText_open çağrılıyor")
                MurText_open(source=APP_EXPLORER)
                return

            #! "MurText eklentisi bu uygulama için yapılandırılmamış."
            ui.message(tr("The MurText add-on is not configured for this application."))
            logger.info("[Master] Başarısız: Bağlam desteklenmiyor")

        except Exception as e:
            #! "Uygulama belirlenirken bir hata oluştu."
            ui.message(tr("An error occurred while identifying the application."))
            logger.error(f"[Master] HATA: {e}")

    def _MurText_kopyala_icin_menu_ac_ve_dene(self):
        try:
            import winUser
            VK_APPS = 0x5D
    
            # Sağ menü tuşu
            winUser.keybd_event(VK_APPS, 0, 0, 0)
            winUser.keybd_event(VK_APPS, 0, winUser.KEYEVENTF_KEYUP, 0)
            logger.info("[Kopyala] Sağ menü tuşu gönderildi. Kısa deneme döngüsü başlıyor...")
    
            # Yavaş cihazlar için 4 kısa deneme: toplam ~1.5 sn içinde biter
            gecikmeler = [150, 250, 400, 700]
            durum = {"i": 0}
    
            def _deneme():
                i = durum["i"]
                bulundu = False
                try:
                    bulundu = bool(self._MurText_try_invoke_copy(afterMenu=True))
                except Exception as e:
                    logger.error(f"[Kopyala] Deneme sırasında hata: {e}")
    
                if bulundu:
                    logger.info(f"[Kopyala] Başarılı. Deneme sayısı: {i+1}")
                    return
    
                if i >= len(gecikmeler) - 1:
                    logger.info("[Kopyala] Deneme bitti. Kopyala bulunamadı.")
                    return
    
                durum["i"] += 1
                wx.CallLater(gecikmeler[durum['i']], _deneme)
    
            wx.CallLater(gecikmeler[0], _deneme)
    
        except Exception as e:
            logger.error(f"[Kopyala] Menü açma/deneme döngüsü hata: {e}")

    def _MurText_open_context_menu(self):
        """WhatsApp üzerinde doğru öğeye odak alıp bağlam menüsünü açar, sonra denemeli aramayı başlatır."""
        try:
            import winUser
    
            VK_APPS = 0x5D
    
            # 1) NVDA navigator nesnesini odağa çekmeyi dene
            try:
                nav = api.getNavigatorObject()
                if nav and getattr(nav, "setFocus", None):
                    nav.setFocus()
                    logger.info("[Kopyala] Navigator nesnesine setFocus denendi.")
            except Exception as e:
                logger.info(f"[Kopyala] Navigator setFocus denenemedi: {e}")
    
            # 2) Bağlam menüsü tuşu
            try:
                winUser.keybd_event(VK_APPS, 0, 0, 0)
                winUser.keybd_event(VK_APPS, 0, winUser.KEYEVENTF_KEYUP, 0)
                logger.info("[Kopyala] Sağ menü tuşu gönderildi, denemeli arama başlatılıyor...")
            except Exception as e:
                logger.error(f"[Kopyala] Sağ menü tuşu gönderilemedi: {e}")
                return
    
            # 3) Menü açılışını bekleyip 4 deneme yap
            wx.CallLater(200, self._MurText_try_invoke_copy, True, 1)
    
        except Exception as e:
            logger.error(f"[Kopyala] Sağ menü açma hata: {e}")
    
    
    def _MurText_try_invoke_copy(self, afterMenu: bool = False, deneme_no: int = 0):
        try:
            focus = api.getFocusObject()
            pid = getattr(focus, "processID", None)
            hwnd = getattr(focus, "windowHandle", None)
    
            def _n(s: str) -> str:
                try:
                    return " ".join(str(s).split()).strip().lower()
                except Exception:
                    return ""
    
            def _obj_ozet(o):
                try:
                    return (
                        f"name={getattr(o,'name',None)!r}, "
                        f"role={getattr(o,'role',None)!r}, "
                        f"class={getattr(o,'windowClassName',None)!r}, "
                        f"hwnd={getattr(o,'windowHandle',None)!r}, "
                        f"pid={getattr(o,'processID',None)!r}"
                    )
                except Exception:
                    return "ozet_alinamadi"
    
            if not _MurText_is_WhatsApp_obj(focus, target_pid=pid):
                ui.message(tr("WhatsApp is not focused."))
                logger.info("[Kopyala] Odak WhatsApp değil")
                logger.info(f"[Kopyala] Odak: {_obj_ozet(focus)}")
                return
    
            # Menü açma ilk çağrı
            if not afterMenu:
                self._MurText_open_context_menu()
                return
    
            # Denemeli arama başlığı ui mesajı tekralamadan
            max_deneme = 4
            logger.info(f"[Kopyala] Deneme: {deneme_no}/{max_deneme}")
    
            # Kopyala etiketi ayarlardan gelsin
            copy_val = conf["KoruzBiz_MurText"].get("copy_key_val", "Kopyala")
            copy_anahtar = _n(copy_val)
    
            logger.info(f"[Kopyala] Odak: {_obj_ozet(focus)}")
            logger.info(f"[Kopyala] Aranan metin: {copy_val!r} (normalize={copy_anahtar!r})")
    
            # 0) Odak zaten hedef mi?
            focus_name_n = _n(getattr(focus, "name", ""))
            if copy_anahtar and focus_name_n and (copy_anahtar in focus_name_n):
                logger.info("[Kopyala] Odak zaten hedef; doAction deneniyor")
                focus.doAction()
                wx.CallLater(300, MurText_WhatsApp)
                return
    
            # 1) Menü kökünü daralt: DIALOG/LIST yakalarsan oradan tara
            root = focus
            prev = None
            zincir_limit = 30
    
            try:
                from controlTypes import Role
                hedef_menu_root_role = {Role.DIALOG, Role.LIST}
            except Exception:
                hedef_menu_root_role = set()
    
            menu_root = None
            for _ in range(zincir_limit):
                if not root or root == prev:
                    break
                prev = root
    
                # Menü konteynerine denk geldiysek dur
                try:
                    r = getattr(root, "role", None)
                    if hedef_menu_root_role and r in hedef_menu_root_role:
                        menu_root = root
                        break
                except Exception:
                    pass
    
                parent = getattr(root, "parent", None)
                if not parent:
                    break
                if pid is not None and getattr(parent, "processID", None) != pid:
                    break
    
                # hwnd koparsa (Chrome katmanları) en fazla burada bırak
                if hwnd is not None and getattr(parent, "windowHandle", None) != hwnd:
                    break
    
                root = parent
    
            if menu_root:
                root = menu_root
                logger.info(f"[Kopyala] Menü kökü daraltıldı: {_obj_ozet(root)}")
            else:
                logger.info(f"[Kopyala] Tarama kökü: {_obj_ozet(root)}")
    
            # 2) Root altında ara
            ziyaret = 0
            max_dugum = 900
            stack = [root]
            seen = set()
    
            while stack and ziyaret < max_dugum:
                node = stack.pop()
                if not node:
                    continue
    
                node_id = id(node)
                if node_id in seen:
                    continue
                seen.add(node_id)
                ziyaret += 1
    
                try:
                    nname_n = _n(getattr(node, "name", ""))
                except Exception:
                    nname_n = ""
    
                if copy_anahtar and nname_n and (copy_anahtar in nname_n):
                    logger.info(f"[Kopyala] Hedef bulundu: {_obj_ozet(node)}")
                    node.doAction()
                    wx.CallLater(300, MurText_WhatsApp)
                    return
    
                try:
                    kids = getattr(node, "children", None) or []
                    for k in reversed(kids):
                        if pid is not None and getattr(k, "processID", None) != pid:
                            continue
                        stack.append(k)
                except Exception:
                    pass
    
            logger.info(f"[Kopyala] Walk bitti. Gezilen düğüm sayısı: {ziyaret}")
    
            # 3) Bulunamadıysa: denemeyi sürdür (UI mesajını sadece en sonda ver)
            if deneme_no < max_deneme:
                # Yavaş cihazlar için küçük artan gecikmeler
                gecikmeler = {1: 350, 2: 550, 3: 800}
                wx.CallLater(gecikmeler.get(deneme_no, 800), self._MurText_try_invoke_copy, True, deneme_no + 1)
                return
    
            # 4) Tüm denemeler bitti: tek UI mesajı
            ui.message(tr("The file has not been downloaded yet or the Copy option is unavailable. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))
    
        except Exception as e:
            logger.error(f"[Kopyala] Genel hata: {e}")
            ui.message(tr("Could not click the Copy option. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))
    
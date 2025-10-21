# -*- coding: utf-8 -*-
# Plug-in version: 1.9
# Murdext compatible version: 2.0.2

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

# NVDA modules
import gui
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

# Hata ayıklama / Debug
def MurText_log_debug(message: str, g: int = 0, t: int = 1):
    """
    1 class
    2 masaüstü
    3 gezgin
    4 WhatsApp
    5 genel
    6 settings
    """
    g_ = 0 # 0: Tüm gruplar 1+ ilgili grup
    t_ = 0 # 0: Hata ve bilgi 1: Sadece bilgi

    try:
        # Filtreleme mantığı
        if g_ == 0 or g == g_:
            if t_ == 0 or (t_ == 1 and t == 1):
                xerr = "Hata: " if t == 0 else ""
                # Yerel loglamayı açmak için aşağıdaki 3 "#" işaretini kaldırın.
                # Remove the following 3 "#" sign to open local loging.
#                log_path = os.path.join(os.path.dirname(__file__), "debug.txt")
#                with open(log_path, "a", encoding="utf-8") as f:
#                    f.write(f"{xerr}{message} [g={g} t={t}] \n")
    except Exception as e:
        pass

# ayarlar / settings
SECTION = "KoruzBiz_MurText"
KEY_OUTPUT_DIR = "outputDir"

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

        MurText_log_debug(f"[Ctx/Desktop] app={app_name}, class={window_class}, name={name}", g=2, t=1)

        # Explorer uygulaması ve Desktop göstergeleri
        if app_name == "explorer" and (
            "desktop" in name or "masaüstü" in name or window_class in ("progman", "folderview")
        ):
            return True

        return False
    except Exception as e:
        MurText_log_debug(f"[Ctx/Desktop] f: MurText_is_desktop_context. {e}", g=2, t=0)
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
                #- MurText_log_debug(f"[Desktop] Reg Desktop: {path}", g=2, t=1)
                return path
    except Exception as e:
        MurText_log_debug(f"[Desktop] _MurText_get_real_desktop Reg okunamadı: {e}", g=2, t=0)

    # Fallback'lar
    home = os.path.expanduser("~")
    cand = os.path.join(home, "Desktop")
    if os.path.isdir(cand):
        MurText_log_debug(f"[Desktop] Fallback: {cand}", g=2, t=1)
        return cand
    od = os.path.join(home, "OneDrive", "Desktop")
    if os.path.isdir(od):
        MurText_log_debug(f"[Desktop] OneDrive Fallback: {od}", g=2, t=1)
        return od

    MurText_log_debug("[Desktop] F: _MurText_get_real_desktop Masaüstü bulunamadı", g=2, t=0)
    return None

def _MurText_try_append_allowed_exts(base_without_ext):
    """Uzantı gizlenmişse izinli uzantıları deneyip var olanı döndür."""
    for ext in ALLOWED_EXTS:
        cand = base_without_ext + ext
        if os.path.isfile(cand):
            MurText_log_debug(f"[Desktop] Uzantı tahmini tuttu: {cand}", g=2, t=1)
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
                MurText_log_debug(f"[Desktop] Kısayol hedefi: {target}", g=2, t=1)
                return target
    except Exception as e:
        MurText_log_debug(f"[Desktop] f: _MurText_resolve_shortcut_if_needed Kısayol çözülemedi: {e}", g=2, t=0)
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
        MurText_log_debug(f"[Desktop] navigator.name='{name}', desktop='{desktop}'", g=2, t=1)

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
        MurText_log_debug("[Desktop] f: _MurText_get_selected_file_desktop Seçili dosya bulunamadı.", g=2, t=0)
        return None
    except Exception as e:
        MurText_log_debug(f"[Desktop] f:_MurText_get_selected_file_desktop {e}", g=2, t=0)
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

        MurText_log_debug(f"[Ctx/Explorer] app={app_name}, class={window_class}, name={name}", g=3, t=1)

        if app_name == "explorer":
            return True
        if window_class in ("cabinetwclass", "explorer"):
            return True
        # Yerelleştirilmiş başlıklar
        if "dosya gezgini" in name or "file explorer" in name:
            return True

        return False
    except Exception as e:
        MurText_log_debug(f"[Ctx/Explorer] f: MurText_is_explorer_context {e}", g=3, t=0)
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
        MurText_log_debug(f"[Explorer] FG HWND: {fg_hwnd}", g=3, t=1)

        shell = comtypes.client.CreateObject("Shell.Application")

        # Sadece FG (ön plandaki) Explorer penceresi
        for w in shell.Windows():
            try:
                w_hwnd = int(getattr(w, "HWND", 0))
                w_name = str(getattr(w, "Name", ""))
                MurText_log_debug(f"[Explorer] window: hwnd={w_hwnd} name={w_name!r}", g=3, t=1)
                if w_hwnd != fg_hwnd:
                    continue
                doc = getattr(w, "Document", None)
                if not doc:
                    MurText_log_debug("[Explorer] FG: Document yok", g=3, t=0)
                    break
                # Seçim var mı?
                try:
                    sel = doc.SelectedItems()
                    if sel and getattr(sel, "Count", 0) > 0:
                        p = sel.Item(0).Path
                        MurText_log_debug(f"[Explorer] Seçili (FG): {p}", g=3, t=1)
                        return p
                except Exception as e_sel:
                    MurText_log_debug(f"[Explorer] FG: SelectedItems hatası: {e_sel}", g=3, t=0)
                # Seçim yoksa klasör yolu
                try:
                    folderPath = doc.Folder.Self.Path
                    MurText_log_debug(f"[Explorer] Seçim yok, klasör yolu (FG): {folderPath}", g=3, t=1)
                    return folderPath
                except Exception as e_fold:
                    MurText_log_debug(f"[Explorer] FG: Folder.Path hatası: {e_fold}", g=3, t=0)
                break  # FG bulundu; daha ötesine bakmaya gerek yok
            except Exception as e_loop:
                MurText_log_debug(f"[Explorer] FG döngü hatası: {e_loop}", g=3, t=0)

        MurText_log_debug("[Explorer] Başarısız: Seçili dosya bulunamadı (FG).", g=3, t=1)
        return None

    except Exception as e:
        MurText_log_debug(f"[Explorer] COM API hatası. f:MurText_get_selected_file_explorer {e}", g=3, t=0)
        # COM başarısızsa PowerShell fallback (son çare)
        try:
            ps_cmd = r'''powershell -command "& { $sel = (New-Object -ComObject Shell.Application).Windows() | Where-Object { $_.Document.SelectedItems().Count -gt 0 } | ForEach-Object { $_.Document.SelectedItems().Item(0).Path }; Write-Output $sel }"'''
            result = subprocess.check_output(ps_cmd, shell=True, universal_newlines=True).strip()
            MurText_log_debug(f"[Explorer] PowerShell sonucu: {result}", g=3, t=1)
            return result if result else None
        except Exception as e2:
            MurText_log_debug(f"[Explorer] PowerShell hatası. F: MurText_get_selected_file_explorer {e2}", g=3, t=0)
            return None

def MurText_get_selected_file():
    """Bağlama göre dosya yolunu alır (Explorer için)."""
    try:
        ctx = MurText_which_app()
        if ctx == APP_EXPLORER:
            return MurText_get_selected_file_explorer()
        MurText_log_debug(f"[get_selected_file] Hata: Bağlam desteklenmiyor. f: MurText_get_selected_file {ctx}", g=3, t=0)
        return None
    except Exception as e:
        MurText_log_debug(f"[get_selected_file] {e}", g=3, t=0)
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

        MurText_log_debug(f"[Ctx/WA] app={app_name}, class={window_class}, role={role}, name={name}", g=4, t=1)

        # En basit ve en sağlam eşleşmeler:
        if "whatsapp" in app_name or "whatsapp" in window_class or "whatsapp" in name:
            return True

        return False
    except Exception as e:
        MurText_log_debug(f"[Ctx/WA] f: MurText_is_WhatsApp_context {e}", g=4, t=0)
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
        MurText_log_debug("MurText_WhatsApp tetiklendi", g=4, t=1)

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
            MurText_log_debug(f"PS hata: rc={result.returncode} err={result.stderr!r}", g=4, t=0)
            #! "Panodan dosya alınamadı."
            ui.message(tr("Failed to retrieve file from clipboard."))
            return

        output = (result.stdout or "").strip()
        MurText_log_debug(f"PowerShell clipboard sonucu (raw utf8): {output!r}", g=4, t=1)

        candidate = ""
        if output:
            for line in (l.strip() for l in output.splitlines() if l.strip()):
                p = os.path.normpath(line)
                # Uzun yol emniyeti (nadir ama dursun)
                if len(p) >= 240 and not p.startswith("\\\\?\\"):
                    p_long = "\\\\?\\" + p
                else:
                    p_long = p

                MurText_log_debug(f"Kontrol edilen yol: {p!r}", g=4, t=1)

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
        MurText_log_debug(f"MurText_WhatsApp {e}", g=4, t=0)

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
        #- MurText_log_debug(f"[Ctx] app={app_name}, class={window_class}, role={role}, name={name}", g=5, t=1)

        if MurText_is_WhatsApp_context():
            MurText_log_debug("[Ctx] Tespit: WhatsApp", g=5, t=1)
            return APP_WhatsApp
        if MurText_is_desktop_context():
            MurText_log_debug("[Ctx] Tespit: Masaüstü", g=5, t=1)
            return APP_DESKTOP
        if MurText_is_explorer_context():
            MurText_log_debug("[Ctx] Tespit: Gezgini", g=5, t=1)
            return APP_EXPLORER

        MurText_log_debug("[Ctx] Tespit: Unknown", g=5, t=1)
        return APP_UNKNOWN
    except Exception as e:
        MurText_log_debug(f"[Ctx] f: MurText_which_app {e}", g=5, t=0)
        return APP_UNKNOWN

def MurText_get_selected_file_smart():
    """
    Masaüstü/Explorer bağlamına göre seçili dosyayı döndürür.
    - Masaüstündeysek: _MurText_get_selected_file_desktop()
    - Değilsek (Explorer ise): MurText_get_selected_file_explorer()
    """
    try:
        if MurText_is_desktop_context():
            MurText_log_debug("[Smart] Bağlam: Masaüstü", g=5, t=1)
            return _MurText_get_selected_file_desktop()

        if MurText_is_explorer_context():
            MurText_log_debug("[Smart] Bağlam: Gezgini", g=5, t=1)
            return MurText_get_selected_file_explorer()

        MurText_log_debug("[Smart] Bağlam desteklenmiyor", g=5, t=0)
        return None
    except Exception as e:
        MurText_log_debug(f"[Smart] f: MurText_get_selected_file_smart {e}", g=5, t=0)
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
                MurText_log_debug(f"Unputable_File: WhatsApp kaydı başarılı | src={file_path} -> dest={dest_file} | ext={ext}", g=5, t=1)
                result.update({"handled": True, "saved": True, "dest": dest_file})
            except Exception as copy_err:
                #! "Dosya kaydedilemedi."
                ui.message(tr("The file could not be saved."))
                MurText_log_debug(f"Unputable_File: WhatsApp kaydı HATASI | src={file_path} | hata={copy_err}", g=5, t=0)
                result.update({"handled": True, "saved": False})
        else:
            # Explorer/Desktop vb.
            #! "Seçilen öğe MurText tarafından desteklenmiyor."
            ui.message(tr("The selected item is not supported by MurText."))
            MurText_log_debug(f"Unputable_File: Desteklenmeyen uzantı | ext={ext} | path={file_path} | source={source}", g=5, t=0)
            result.update({"handled": True})
    except Exception as e:
        MurText_log_debug(f"Unputable_File: İstisna: {e}", g=5, t=0)

    return result

def MurText_open(file_path=None, source=None):
    try:
        MurText_log_debug(f"MurText_open tetiklendi | source: {source}", g=5, t=1)

        # Dosya yolu belirlenmemişse, kaynak üzerinden alınır
        if file_path is None:
            MurText_log_debug(f"Dosya yolu belirtilmedi. Kaynak: {source}", g=5, t=0)
            if source == APP_DESKTOP:
                file_path = MurText_get_selected_file_smart()
            elif source == APP_EXPLORER:
                file_path = MurText_get_selected_file()

        MurText_log_debug(f"Alınan dosya yolu (ham): {file_path}", g=5, t=1)

        # Merkezî kontrol (varlık + uzantı desteği)
        fc = file_control(file_path)

        if not fc["ok"]:
            reason = fc.get("reason")
            full_path = fc.get("file_path")
            ext = fc.get("ext")

            if reason in ("missing", "not_exists"):
                #! "Geçersiz yordam veya dosya yolu."
                ui.message(tr("Invalid procedure or file path."))
                MurText_log_debug(f"Başarısız: Dosya yolu alınamadı veya mevcut değil. path={full_path}, g=5, t=1")
                return

            if reason == "unsupported":
                # Desteklenmeyen dosyayı Unputable_File'a devret
                _ = Unputable_File(source=source, file_path=full_path, ext=ext)
                return

            # Beklenmeyen durum
            #! "Bir hata oluştu."
            ui.message(tr("An error occurred."))
            MurText_log_debug(f"Başarısız: Bilinmeyen kontrol sonucu: reason={reason} | path={full_path} | ext={ext}")
            return

        # Buraya gelindiyse dosya mevcut ve uzantı destekleniyor
        file_path = fc["file_path"]
        file_name = os.path.basename(file_path)
        MurText_log_debug(f"Dosya adı: {file_name} | Uzantı: {fc['ext']}", g=5, t=1)

        #! "MurText ile açılıyor. Uygulama hazırlanıyor."
        ui.message(tr("Opening with MurText. Preparing the application."))
        subprocess.Popen([get_murtext_exe_path(), file_path])
        MurText_log_debug(f"MurText çalıştırıldı: {file_path} -> {get_murtext_exe_path()}", g=5, t=1)

    except Exception as e:
        #! "Bir hata oluştu."
        ui.message(tr("An error occurred."))
        MurText_log_debug(f"MurText_open istisnası: {e}", g=5, t=0)

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
        MurText_log_debug(f"[Probe] MurText var mı? {MurText_INSTALLED} {MurText_path}", g=5, t=1)
        if not MurText_INSTALLED:
            MurText_log_debug("MurText kurulu değil", g=5, t=1)
            MurText_prompt_to_install_if_missing()
        return MurText_INSTALLED
    except Exception as e:
        MurText_INSTALLED = False
        MurText_log_debug("MurText kurulu değil", g=5, t=1)
        MurText_log_debug(f"[Probe] f:MurText_probe_installation_on_load {e}", g=5, t=0)
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

            MurText_log_debug(f"[Prompt] Sonuç id: {res}", g=5, t=1)

            if res == wx.ID_YES:
                try:
                    webbrowser.open("https://MurText.org?page=download&source=nvda", new=1)
                except Exception as e:
                    MurText_log_debug(f"[Prompt] URL açılamadı: {e}", g=5, t=0)
            elif res == wx.ID_NO:
                MurText_log_debug("[Prompt] HAYIR: Kullanıcı reddetti.", g=5, t=1)
            else:
                MurText_log_debug("[Prompt] Kapatıldı / iptal edildi.", g=5, t=1)
        except Exception as e:
            MurText_log_debug(f"[Prompt] pop up : {e}", g=5, t=0)

    wx.CallAfter(_show)

class GlobalPlugin(_BaseGlobalPlugin):
    def __init__(self):
        super().__init__()
        MurText_log_debug(f"Yüklendi",  g=1, t=1)

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
        MurText_log_debug("\n#! Tetiklendi !#", g=1, t=1)

        # Sadece tutucu false ise 
        if not MurText_INSTALLED:
            #- MurText_log_debug("Varlık kontrol ediliyor...")
            if not MurText_probe_installation_on_load():
                # Kurulu değil -> pencere açıldı, işi kesiyoruz
                return
        try:
            ctx = MurText_which_app()
            MurText_log_debug(f"[Master] Bağlam: {ctx}", g=1, t=1)

            if ctx == APP_WhatsApp:
                #- MurText_log_debug("[Master] WhatsApp algılandı, menü açılacak ve tetikleme yapılacak", g=1, t=1)
                try:
                    # Bağlam menüsü: Shift+F10'u 'kb' jesti olarak gönder
                    try:
                        KIG.fromName("shift+f10").send()
                    except Exception as e:
                        MurText_log_debug(f"[Master] Shift+F10 gönderilemedi: {e}", g=1, t=0)

                    # Menü render olsun; sonra 'Kopyala' taraması
                    wx.CallLater(250, self._MurText_try_invoke_copy)
                    return

                    #- MurText_log_debug("[Master] Menü açıldı ve Insert+Shift+K gönderildi", g=1, t=1)
                except Exception as e:
                    MurText_log_debug(f"[Master] Tuş gönderimi hatası: {e}", g=1, t=0)
                    #! "Menü açma işlemi başarısız."
                    ui.message(tr("Failed to open the menu."))
                return

            if ctx == APP_DESKTOP:
                #- MurText_log_debug("[Master] Masaüstü algılandı, MurText_open çağrılıyor", g=1, t=1)
                MurText_open(source=APP_DESKTOP)
                return

            if ctx == APP_EXPLORER:
                #- MurText_log_debug("[Master] Gezginde tetiklendi, MurText_open çağrılıyor", g=1, t=1)
                MurText_open(source=APP_EXPLORER)
                return

            #! "MurText eklentisi bu uygulama için yapılandırılmamış."
            ui.message(tr("The MurText add-on is not configured for this application."))
            MurText_log_debug("[Master] Başarısız: Bağlam desteklenmiyor", g=1, t=0)

        except Exception as e:
            #! "Uygulama belirlenirken bir hata oluştu."
            ui.message(tr("An error occurred while identifying the application."))
            MurText_log_debug(f"[Master] HATA: {e}", g=1, t=0)




    def _MurText_try_invoke_copy(self):
        """WhatsApp bağlam menüsünde 'Kopyala' öğesini bulup tıklar; başarıysa MurText_WhatsApp() çağırır."""
        try:
            focus = api.getFocusObject()
            pid = getattr(focus, "processID", None)
    
            if not _MurText_is_WhatsApp_obj(focus, target_pid=pid):
                #! "WhatsApp odağı yok."
                ui.message(tr("WhatsApp is not focused."))
                MurText_log_debug("[Kopyala] Odak WhatsApp değil", g=1, t=0)
                return
    
            menu_root = _MurText_nearest_menu_root(focus) or _MurText_nearest_menu_root(api.getFocusObject())
            if not menu_root:
                #! "Bağlam menüsü bulunamadı."
                ui.message(tr("Context menu not found."))
                MurText_log_debug("[Kopyala] Menü kökü bulunamadı", g=1, t=0)
                return
    
            # Role sabitlemesi — NAMEERROR riskini kaldır
            try:
                from controlTypes import Role
                ROLE_MENUITEM = Role.MENUITEM
            except Exception:
                ROLE_MENUITEM = getattr(menu_root, "role", None).__class__.MENUITEM if hasattr(menu_root, "role") else None
    
            for child in (getattr(menu_root, "children", None) or []):
                if not _MurText_is_WhatsApp_obj(child, target_pid=pid):
                    continue
                if ROLE_MENUITEM and getattr(child, "role", None) != ROLE_MENUITEM:
                    continue
                name = _MurText_safe(getattr(child, "name", "")).lower()
                copy_val = conf["KoruzBiz_MurText"].get("copy_key_val", "copy")
                if name and copy_val.lower() in name.lower():
                    try:
                        child.doAction()
                        try:
                            wx.CallLater(250, MurText_WhatsApp)  # UI thread’i bloklamadan
                        except Exception:
                            MurText_WhatsApp()
                        return
                    except Exception as e:
                        MurText_log_debug(f"[Kopyala] doAction hata: {e}", g=1, t=0)
                        #! Kopyala seçeneğine tıklanamadı. Lütfen ayarlar iletişim kutusundan kopyala etiketini tanımlayın.
                        ui.message(tr("Could not click the Copy option. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))
                        return
    
            #! "Dosya henüz bilgisayara indirilmemiş veya kopyala seçeneği yok. Lütfen ayarlar iletişim kutusundan kopyala etiketini tanımlayın.
            ui.message(tr("The file has not been downloaded yet or the Copy option is unavailable. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))
        except Exception as e:
            MurText_log_debug(f"[Kopyala] Genel hata: {e}", g=1, t=0)
            #! "Kopyala seçeneğine tıklanamadı. Lütfen ayarlar iletişim kutusundan kopyala etiketini tanımlayın."
            ui.message(tr("Could not click the Copy option. Please open the Settings dialog and save the Copy label you see in WhatsApp’s context menu."))
    
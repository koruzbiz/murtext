import os
import wx
import speech
from config import conf
import gettext
import locale

# Taban sınıfı ve kayıt API'sini doğrudan al:
try:
    from gui.settingsDialogs import SettingsPanel, registerSettingsPanel, NVDASettingsDialog
except Exception:
    from gui import settingsDialogs
    SettingsPanel = settingsDialogs.SettingsPanel
    registerSettingsPanel = None
    NVDASettingsDialog = settingsDialogs.NVDASettingsDialog

# Paket içindeki tr'yi kullan (fallback: düz metin)
try:
    from . import tr as _pkg_tr
    def tr(msg): return _pkg_tr(msg)
except Exception:
    def tr(msg): return msg

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


# ayarlar / keys
SECTION = "KoruzBiz_MurText"
KEY_OUTPUT_DIR = "outputDir"
KEY_COPY_KEY = "copy_key_val"
KEY_COPY_SOURCE = "copy_key_source"  # 'manual_map' veya 'gettext_fallback' veya 'manual_user'

# Mini sözlük
_MANUAL_COPY_MAP= {
    'en': 'Copy',
    'tr': 'Kopyala',
    'es': 'Copiar',
    'fr': 'Copier',
    'de': 'Kopieren',
    'it': 'Copia',
    'pt': 'Copiar',
    'nl': 'Kopiëren',
    'ru': 'Копировать',
    'bg': 'Копиране',
    'pl': 'Kopiuj',
    'sv': 'Kopiera',
    'no': 'Kopier',
    'da': 'Kopier',
    'fi': 'Kopioi',
    'cs': 'Kopírovat',
    'sk': 'Kopírovať',
    'hu': 'Másol',
    'ro': 'Copiază',
    'el': 'Αντιγραφή',
    'ja': 'コピー',
    'ko': '복사',
    'zh_CN': '复制',
    'zh_TW': '複製',
    'vi': 'Sao chép',
    'id': 'Salin',
    'ms': 'Salin',
    'th': 'คัดลอก',
    'he': 'העתק',
    'ar': 'نسخ',
    'hi': 'कॉपी',
    'bn': 'কপি',
    'uk': 'Копіювати',
    'sr': 'Копирај',
    'hr': 'Kopiraj',
    'sl': 'Kopiraj',
    'et': 'Kopeeri',
    'lv': 'Kopēt',
    'lt': 'Kopijuoti',
    'mk': 'Копирај',
    'fa': 'کپی',
    'sw': 'Nakili',
    'ha': 'Kwafi',
    'ta': 'நகலெடு',
    'ml': 'പകർത്തുക',
    'gu': 'નકલ કરો',
    'kn': 'ನಕಲಿಸಿ', 
    'mr': 'प्रत करा',
    'te': 'కాపీ చేయండి',
    'zh_HK': '複製',
    'az': 'Köçür',  
    'am': 'ቅጂ',
    'fil': 'Kopyahin',
    'ur': 'کاپی',
}
def _get_documents_dir() -> str:
    try:
        import ctypes, ctypes.wintypes as wt
        CSIDL_PERSONAL = 5
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(wt.MAX_PATH)
        if ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, buf) == 0 and buf.value:
            return buf.value
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Documents")

def _ensure_defaults():
    if SECTION not in conf:
        conf[SECTION] = {}
    if not conf[SECTION].get(KEY_OUTPUT_DIR):
        conf[SECTION][KEY_OUTPUT_DIR] = _get_documents_dir()
        conf.save()

def _find_copy():
    # Hazırla
    if SECTION not in conf:
        conf[SECTION] = {}

    # Eğer zaten varsa çık
    if conf[SECTION].get(KEY_COPY_KEY):
        logger.info("Zaten copy anahtarı mevcut; atlandı.")
        return conf[SECTION][KEY_COPY_KEY]

    # 1) OS dilini al
    lang_code = None
    try:
        loc = locale.getdefaultlocale()
        if loc and loc[0]:
            lang_code = loc[0]  # örn 'tr_TR' veya 'zh_TW'
    except Exception:
        lang_code = None

    # Normalize: küçük harf ve alt çizgi (örn 'tr', 'zh_TW')
    normalized = None
    lang_short = None
    if lang_code:
        normalized = lang_code.replace('-', '_')
        lang_short = normalized.split('_')[0]

    # 2) Mini sözlük kontrolü (öncelik: bölgesel zh varyantı, sonra kısa kod)
    # Eğer zh ise bölgeyi netleştir
    if lang_short == 'zh':
        # Eğer 'zh_TW' veya 'zh_HK' gibi TW benzeri varsa TW, aksi halde CN
        if normalized and ('TW' in normalized.upper() or 'HK' in normalized.upper() or 'MO' in normalized.upper()):
            key = 'zh_TW'
        else:
            key = 'zh_CN'
        val = _MANUAL_COPY_MAP.get(key)
        if val:
            conf[SECTION][KEY_COPY_KEY] = val
            conf[SECTION][KEY_COPY_SOURCE] = 'manual_map'
            conf.save()
            logger.info(f"manual_map ile bulundu: {key} => {val}")
            return val

    # Diğer diller için kısa kodu kontrol et 
    if lang_short:
        val = _MANUAL_COPY_MAP.get(lang_short)
        if val:
            conf[SECTION][KEY_COPY_KEY] = val
            conf[SECTION][KEY_COPY_SOURCE] = 'manual_map'
            conf.save()
            logger.info(f"manual_map ile bulundu: {lang_short} => {val}")
            return val

    # 3) Mini sözlükte yoksa fallback: gettext ile 'Copy' çevirisini al 
    try:
        msg = tr('Copy')
    except Exception:
        msg = 'Copy'

    # Kaydet ve çık
    conf[SECTION][KEY_COPY_KEY] = msg
    conf[SECTION][KEY_COPY_SOURCE] = 'gettext_fallback'
    conf.save()
    logger.info(f"gettext fallback kaydedildi => {msg}")
    return msg

class MurTextSettingsPanel(SettingsPanel):
    title = tr("Koruz.biz MurText")

    def makeSettings(self, sizer):
        _ensure_defaults()

        grid = wx.FlexGridSizer(rows=3, cols=2, vgap=6, hgap=6)
        grid.AddGrowableCol(1, 1)

        # Varsayılan dosya kayıt yeri
        labelText = tr("Default file save location")
        label = wx.StaticText(self, label=labelText + ":")
        grid.Add(label, flag=wx.ALIGN_CENTER_VERTICAL)

        startPath = conf[SECTION].get(KEY_OUTPUT_DIR, _get_documents_dir())
        self.dirPicker = wx.DirPickerCtrl(
            self,
            path=startPath,
            message=tr("Select the save folder"),
            style=wx.DIRP_DIR_MUST_EXIST | wx.DIRP_USE_TEXTCTRL
        )
        grid.Add(self.dirPicker, flag=wx.EXPAND)

        # WP Desktop "Kopyala" karşılığı alanı
        #! WhatsApp Desktop Kopyala etiketi
        label_copy = wx.StaticText(self, label=tr("Define the Copy label in WhatsApp Desktop.") + ":")
        grid.Add(label_copy, flag=wx.ALIGN_CENTER_VERTICAL)
        copy_val = conf[SECTION].get(KEY_COPY_KEY, '')
        self.copyText = wx.TextCtrl(self, value=copy_val if copy_val else '', style=wx.TE_PROCESS_ENTER)
        try:
            self.copyText.SetHelpText(tr('Enter the context-menu "Copy" text as shown in WhatsApp Desktop (e.g. "Copy", "Kopyala", "複製").'))
        except Exception:
            pass
        grid.Add(self.copyText, flag=wx.EXPAND)

        # A11y: minimal müdahale
        try:
            import ui  # NVDA konuşma API'sı
            btn = self.dirPicker.GetPickerCtrl() if hasattr(self.dirPicker, "GetPickerCtrl") else None
            if btn:
                try:
                    btn.SetName(tr("Browse"))
                except Exception:
                    pass
                def _announce_after_browse_focus(evt):
                    wx.CallLater(80, lambda: ui.message(labelText))
                    evt.Skip()
                btn.Bind(wx.EVT_SET_FOCUS, _announce_after_browse_focus)
        except Exception:
            pass

        sizer.Add(grid, flag=wx.ALL | wx.EXPAND, border=12)

    def onSave(self):
        try:
            # Kayıt klasörü
            if hasattr(self, "dirPicker"):
                path = self.dirPicker.GetPath()
                if path and os.path.isdir(path):
                    conf[SECTION][KEY_OUTPUT_DIR] = path
                    conf.save()
    
            if hasattr(self, 'copyText'):
                val = self.copyText.GetValue().strip()
                if val:
                    # Normal durumda kullanıcı bir değer girmiş
                    conf[SECTION][KEY_COPY_KEY] = val
                    conf[SECTION][KEY_COPY_SOURCE] = 'manual_user'
                else:
                    # Kullanıcı kutuyu boş bıraktıysa: ayarı temizle
                    if KEY_COPY_KEY in conf[SECTION]:
                        try:
                            # conf[SECTION].pop(KEY_COPY_KEY, None)  # tamamen kaldır
                            conf[SECTION][KEY_COPY_KEY] = ''   # boş string kaydet
                            conf[SECTION].pop(KEY_COPY_SOURCE, None)
                        except Exception:
                            # Basitçe boş string ata fallback
                            conf[SECTION][KEY_COPY_KEY] = ''
                            conf[SECTION].pop(KEY_COPY_SOURCE, None)
                conf.save()
        except Exception:
            pass

    def save(self):
        return self.onSave()

# Panel kaydı
_MurText_SETTINGS_REGISTERED = False
def _register_settings_panel_once():
    global _MurText_SETTINGS_REGISTERED
    if _MurText_SETTINGS_REGISTERED:
        return
    try:
        if registerSettingsPanel:
            registerSettingsPanel(MurTextSettingsPanel)
        else:
            if MurTextSettingsPanel not in NVDASettingsDialog.categoryClasses:
                NVDASettingsDialog.categoryClasses.append(MurTextSettingsPanel)
    except Exception:
        pass
    else:
        _MurText_SETTINGS_REGISTERED = True

_register_settings_panel_once()

# Eklenti başlarken: sadece eğer key yoksa otomatik tespit çalışsın
try:
    if SECTION not in conf:
        conf[SECTION] = {}
    if not conf[SECTION].get(KEY_COPY_KEY):
        _res = _find_copy()
        logger.info(f"Otomatik tespit sonucu: {_res}")
except Exception:
    # Sessiz kal
    pass
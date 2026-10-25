import os
import wx
from config import conf
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
KEY_AV_TASK = "av_task_type"
KEY_AV_TASK_USER = "av_task_user_set"
KEY_DOWNLOAD_KEY = "download_key_val"
KEY_DOWNLOAD_SOURCE = "download_key_source"  # 'manual_map' veya 'default_en' veya 'manual_user'

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
# İndir sözlüğü: birden fazla karşılık ';' ile ayrılır. Listede olmayan diller için kullanıcı ayarlardan kendi etiketini girer.
_MANUAL_DOWNLOAD_MAP = {
    'en': 'Download',
    'tr': 'İndir',
    'es': 'Descargar',
    'fr': 'Télécharger',
    'de': 'Herunterladen;Download',
    'it': 'Scarica',
    'pt': 'Baixar;Transferir',
    'nl': 'Downloaden',
    'ru': 'Скачать;Загрузить',
    'bg': 'Изтегляне;Изтегли',
    'pl': 'Pobierz',
    'sv': 'Ladda ned',
    'no': 'Last ned',
    'da': 'Download;Hent',
    'fi': 'Lataa',
    'cs': 'Stáhnout',
    'sk': 'Stiahnuť',
    'hu': 'Letöltés',
    'ro': 'Descarcă',
    'el': 'Λήψη',
    'ja': 'ダウンロード',
    'ko': '다운로드',
    'zh_CN': '下载',
    'zh_TW': '下載',
    'zh_HK': '下載',
    'vi': 'Tải xuống',
    'id': 'Unduh',
    'ms': 'Muat turun',
    'th': 'ดาวน์โหลด',
    'he': 'הורדה;הורד',
    'ar': 'تنزيل',
    'hi': 'डाउनलोड करें;डाउनलोड',
    'uk': 'Завантажити',
    'fa': 'دانلود;بارگیری',
    'hr': 'Preuzmi',
    'sr': 'Преузми;Preuzmi',
    'sl': 'Prenesi',
    'et': 'Laadi alla',
    'lv': 'Lejupielādēt',
    'lt': 'Atsisiųsti',
    'fil': 'I-download',
    'ur': 'ڈاؤن لوڈ',
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
    if conf[SECTION].get(KEY_AV_TASK) is None:
        conf[SECTION][KEY_AV_TASK] = 2

def _dil_kodlari():
    """OS dilinden (örn. 'tr_TR', 'zh_TW') sözlük anahtarı adaylarını sırayla döndürür."""
    try:
        loc = locale.getdefaultlocale()
        lang_code = loc[0] if loc and loc[0] else None
    except Exception:
        lang_code = None
    if not lang_code:
        return []
    normalized = lang_code.replace('-', '_')
    kisa = normalized.split('_')[0]
    if kisa == 'zh':
        bolge = normalized.upper()
        return ['zh_TW' if ('TW' in bolge or 'HK' in bolge or 'MO' in bolge) else 'zh_CN']
    return [kisa]

def _etiket_bul(anahtar, kaynak_anahtar, sozluk, varsayilan):
    if SECTION not in conf:
        conf[SECTION] = {}
    if conf[SECTION].get(anahtar):
        return conf[SECTION][anahtar]
    for kod in _dil_kodlari():
        val = sozluk.get(kod)
        if val:
            conf[SECTION][anahtar] = val
            conf[SECTION][kaynak_anahtar] = 'manual_map'
            conf.save()
            logger.info(f"manual_map ile bulundu: {anahtar} {kod} => {val}")
            return val
    conf[SECTION][anahtar] = varsayilan
    conf[SECTION][kaynak_anahtar] = 'gettext_fallback' if anahtar == KEY_COPY_KEY else 'default_en'
    conf.save()
    logger.info(f"varsayılan kaydedildi: {anahtar} => {varsayilan}")
    return varsayilan

def _find_copy():
    try:
        varsayilan = tr('Copy')
    except Exception:
        varsayilan = 'Copy'
    return _etiket_bul(KEY_COPY_KEY, KEY_COPY_SOURCE, _MANUAL_COPY_MAP, varsayilan)

def _find_download():
    return _etiket_bul(KEY_DOWNLOAD_KEY, KEY_DOWNLOAD_SOURCE, _MANUAL_DOWNLOAD_MAP, 'Download')

class MurTextSettingsPanel(SettingsPanel):
    title = tr("Koruz.biz MurText")

    def makeSettings(self, sizer):
        _ensure_defaults()

        grid = wx.FlexGridSizer(rows=5, cols=2, vgap=6, hgap=6)
        grid.AddGrowableCol(1, 1)

        label_task = wx.StaticText(self, label=tr("Specify the task type") + ":")
        grid.Add(label_task, flag=wx.ALIGN_CENTER_VERTICAL)
        self.taskChoice = wx.Choice(self, choices=[
            tr("Send audio / video files to MurText"),
            tr("Save audio / video files to the folder"),
            tr("Ask me every time"),
        ])
        try:
            secim = int(conf[SECTION].get(KEY_AV_TASK, 2)) if conf[SECTION].get(KEY_AV_TASK_USER) else 0
        except Exception:
            secim = 0
        self.taskChoice.SetSelection(secim if secim in (0, 1, 2) else 0)
        self._taskTouched = False
        self.taskChoice.Bind(wx.EVT_CHOICE, self._onTaskChoice)
        grid.Add(self.taskChoice, flag=wx.EXPAND)

        label = wx.StaticText(self, label=tr("Default file save location") + ":")
        grid.Add(label, flag=wx.ALIGN_CENTER_VERTICAL)

        self.dirText = wx.TextCtrl(self, value=conf[SECTION].get(KEY_OUTPUT_DIR, _get_documents_dir()))
        self.dirButton = wx.Button(self, label=tr("Select folder"))
        self.dirButton.Bind(wx.EVT_BUTTON, self._onSelectDir)
        dirSizer = wx.BoxSizer(wx.HORIZONTAL)
        dirSizer.Add(self.dirText, proportion=1, flag=wx.EXPAND | wx.RIGHT, border=6)
        dirSizer.Add(self.dirButton)
        grid.Add(dirSizer, flag=wx.EXPAND)

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

        label_dl = wx.StaticText(self, label=tr("Define the Download label in WhatsApp Desktop.") + ":")
        grid.Add(label_dl, flag=wx.ALIGN_CENTER_VERTICAL)
        dl_val = conf[SECTION].get(KEY_DOWNLOAD_KEY, '')
        self.downloadText = wx.TextCtrl(self, value=dl_val if dl_val else '', style=wx.TE_PROCESS_ENTER)
        try:
            self.downloadText.SetHelpText(tr('Enter the context-menu "Download" text as shown in WhatsApp Desktop (e.g. "Download", "İndir", "下載"). Use ";" to separate multiple equivalents.'))
        except Exception:
            pass
        grid.Add(self.downloadText, flag=wx.EXPAND)

        sizer.Add(grid, flag=wx.ALL | wx.EXPAND, border=12)

    def _onTaskChoice(self, evt):
        self._taskTouched = True
        evt.Skip()

    def _onSelectDir(self, evt):
        dlg = wx.DirDialog(self, message=tr("Select the save folder"), defaultPath=self.dirText.GetValue(), style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        try:
            if dlg.ShowModal() == wx.ID_OK:
                self.dirText.SetValue(dlg.GetPath())
        finally:
            dlg.Destroy()

    def onSave(self):
        try:
            if hasattr(self, "taskChoice") and (self._taskTouched or conf[SECTION].get(KEY_AV_TASK_USER)):
                conf[SECTION][KEY_AV_TASK] = self.taskChoice.GetSelection()
                conf[SECTION][KEY_AV_TASK_USER] = 1
            # Kayıt klasörü
            if hasattr(self, "dirText"):
                path = self.dirText.GetValue().strip()
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

            if hasattr(self, 'downloadText'):
                dval = self.downloadText.GetValue().strip()
                if dval:
                    conf[SECTION][KEY_DOWNLOAD_KEY] = dval
                    conf[SECTION][KEY_DOWNLOAD_SOURCE] = 'manual_user'
                else:
                    conf[SECTION][KEY_DOWNLOAD_KEY] = ''
                    conf[SECTION].pop(KEY_DOWNLOAD_SOURCE, None)
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

# Eklenti başlarken: anahtar yoksa otomatik tespit çalışsın
try:
    if SECTION not in conf:
        conf[SECTION] = {}
    if not conf[SECTION].get(KEY_COPY_KEY):
        logger.info(f"Otomatik tespit (Kopyala): {_find_copy()}")
    if not conf[SECTION].get(KEY_DOWNLOAD_KEY):
        logger.info(f"Otomatik tespit (İndir): {_find_download()}")
except Exception:
    pass

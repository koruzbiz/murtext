import os
import wx
import speech
from config import conf
# Taban sınıfı ve kayıt API'sini doğrudan al:
try:
	from gui.settingsDialogs import SettingsPanel, registerSettingsPanel, NVDASettingsDialog
except Exception:
	# Çok eski NVDA'larda registerSettingsPanel olmayabilir
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
#                log_path = os.path.join(os.path.dirname(__file__), "cikti.txt")
#                with open(log_path, "a", encoding="utf-8") as f:
#                    f.write(f"{xerr}{message} [g={g} t={t}] \n")
    except Exception as e:
        pass

# ayarlar / settings
SECTION = "KoruzBiz_MurText"
KEY_OUTPUT_DIR = "outputDir"

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
		MurText_log_debug(f"Defaults set: {KEY_OUTPUT_DIR}={conf[SECTION][KEY_OUTPUT_DIR]}")

class MurTextSettingsPanel(SettingsPanel):
	title = tr("Koruz.biz MurText")

	def makeSettings(self, sizer):
		MurText_log_debug("makeSettings: start")
		_ensure_defaults()
	
		grid = wx.FlexGridSizer(rows=2, cols=2, vgap=6, hgap=6)
		grid.AddGrowableCol(1, 1)
	
	#! "Varsayılan dosya kayıt yeri"
		labelText = tr("Default file save location")
		label = wx.StaticText(self, label=labelText + ":")
		grid.Add(label, flag=wx.ALIGN_CENTER_VERTICAL)
	
		startPath = conf[SECTION].get(KEY_OUTPUT_DIR, _get_documents_dir())
		self.dirPicker = wx.DirPickerCtrl(
			self,
			path=startPath,
			#! "Kayıt klasörünü seçin"
			message=tr("Select the save folder"),
			style=wx.DIRP_DIR_MUST_EXIST | wx.DIRP_USE_TEXTCTRL
		)
		grid.Add(self.dirPicker, flag=wx.EXPAND)
	
		# A11y ve davranış: minimum müdahale ---
		try:
			import ui  # NVDA konuşma API'sı
	
			# İçteki text ctrl'ü bul ve yazmayı kapat (manuel giriş olmasın)
			txt = self.dirPicker.GetTextCtrl() if hasattr(self.dirPicker, "GetTextCtrl") else None
			if not txt:
				for c in self.dirPicker.GetChildren():
					if isinstance(c, wx.TextCtrl):
						txt = c
						break
			if txt:
				txt.SetEditable(False)
				try:
					txt.SetName(labelText)
					#! "MurText için varsayılan kayıt klasörü"
					txt.SetHelpText(tr("Default save folder for MurText"))
				except Exception:
					pass
				# DİKKAT: txt için EVT_SET_FOCUS bağlanmadı
	
			# Sadece "göz at" odak alınca etiketi hemen ardından duyur
			btn = self.dirPicker.GetPickerCtrl() if hasattr(self.dirPicker, "GetPickerCtrl") else None
			if btn:
				try:
					btn.SetName(tr("Browse"))
				except Exception:
					pass
				def _announce_after_browse_focus(evt):
					# Önce NVDA butonu okur; minik gecikmeyle etiketi ekle
					wx.CallLater(80, lambda: ui.message(labelText))
					evt.Skip()
				btn.Bind(wx.EVT_SET_FOCUS, _announce_after_browse_focus)
		except Exception:
			pass
		# A11y son ---
	
		sizer.Add(grid, flag=wx.ALL | wx.EXPAND, border=12)
		MurText_log_debug("makeSettings: done")
	
	# NVDA bu ismi (abstract) bekliyor
	def onSave(self):
		try:
			if hasattr(self, "dirPicker"):
				path = self.dirPicker.GetPath()
				if path and os.path.isdir(path):
					conf[SECTION][KEY_OUTPUT_DIR] = path
					conf.save()
					MurText_log_debug(f"onSave: {KEY_OUTPUT_DIR}={path}")
		except Exception as e:
			MurText_log_debug(f"onSave ERROR: {e!r}")

	# Geriye dönük: bazı örnekler save() çağırabilir
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
			MurText_log_debug("registered via registerSettingsPanel")
		else:
			# Eski yol
			if MurTextSettingsPanel not in NVDASettingsDialog.categoryClasses:
				NVDASettingsDialog.categoryClasses.append(MurTextSettingsPanel)
				MurText_log_debug("appended to NVDASettingsDialog.categoryClasses")
	except Exception as e:
		MurText_log_debug(f"register ERROR: {e!r}")
	else:
		_MurText_SETTINGS_REGISTERED = True

# Bazı kurulumlarda kayıt timing'i kritik — GUI ayağa kalkınca kaydet
_register_settings_panel_once()


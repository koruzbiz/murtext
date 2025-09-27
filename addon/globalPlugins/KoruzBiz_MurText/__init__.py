# globalPlugins/KoruzBiz_MurText/__init__.py
import os
import gettext
import languageHandler  # NVDA'nın dil bilgisi modülü

_pkg_dir = os.path.dirname(__file__)
_locales = os.path.join(_pkg_dir, "locales")

# NVDA'nın aktif dili 
_lang_code = languageHandler.getLanguage()

# .mo adı: KoruzBiz_MurText 
_lang = gettext.translation(
	domain="KoruzBiz_MurText",
	localedir=_locales,
	languages=[_lang_code],
	fallback=True,
)

# Eklenti içi çeviri fonksiyonu
tr = _lang.gettext

# GlobalPlugin'i dışa aktar 
from .KoruzBiz_MurText import GlobalPlugin  # noqa: E402

# Ayarlar panelini yükle/registre etsin diye import 
from . import settings  # noqa: F401

__all__ = ["GlobalPlugin", "tr"]

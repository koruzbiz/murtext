# Derleme özelleştirmeleri
# Mümkün olduğunca sconstruct veya manifest dosyaları yerine bu dosyayı değiştirin.

from site_scons.site_tools.NVDATool.typings import AddonInfo, BrailleTables, SymbolDictionaries

# `addon_info` içindeki bazı dizeler çevrilebilir olduğundan,  bu dizeleri .po dosyalarına dahil etmemiz gerekir.  Gettext yalnızca `_` fonksiyonuna parametre olarak verilen dizeleri tanır.  Bu modülde çevirileri başlatmaktan kaçınmak için, argümanı olduğu gibi döndüren "sahte" bir `_` fonksiyonunu yalnızca içe aktarıyoruz.

from site_scons.site_tools.NVDATool.utils import _


# Eklenti bilgi değişkenleri
addon_info = AddonInfo(
	# Eklentinin adı/kimliği, NVDA için dahili
	addon_name="KoruzBiz_MurText",
	# Eklenti özeti/başlığı, genelde kullanıcıya görünen eklenti adı
	# Çevirmenler: Eklenti yüklemesinde ve eklenti mağazasında gösterilecek
	# bu eklenti için özet/başlık
	#! "Koruz.biz MurText"
	addon_summary=_("Koruz.biz MurText"),
	# Eklenti açıklaması
	# Çevirmenler: Eklenti mağazasındaki eklenti bilgi ekranında gösterilecek uzun açıklama
	#! "Bilgisayardaki ses/video dosyalarını ve WhatsApp ses dosyalarını 100 dilde yazıya dönüştüren MurText eklentisi."
	addon_description=_("""An NVDA plug-in for the MurText application, which transcribes audio/video files and WhatsApp voice messages in 100 languages."""),
	# sürüm
	addon_version="2.0.2",
	# Bu sürüm için kısa değişiklik günlüğü
	# Çevirmenler: Eklenti mağazasında gösterilecek bu eklenti sürümü için "yenilikler" içeriği
	#! "Tek kısayolla bilgisayarınızdaki ve WhatsApp Masaüstü'ndeki ses/video dosyaları sınırsız sürede transkribe edin. WhatsApp'taki .doc, .zip vb. dosyaları aynı kısayolla kaydedin."
	addon_changelog=_("""Free and unlimited transcription! With a single shortcut, transcribe audio/video files and WhatsApp voice messages instantly. Use the same shortcut to quickly save other files from the WhatsApp desktop app."""),
	# Yazar(lar)
	addon_author="Murat Kefeli <bilgi@koruz.biz>",
	# Eklenti dokümantasyon desteği için URL
	addon_url="https://MurText.org",
	# Kaynak kodunun bulunabileceği eklenti deposu URL’si
	addon_sourceURL=None,
	# Dokümantasyon dosya adı
	addon_docFileName="readme.html",
	# Desteklenen minimum NVDA sürümü (örn. "2019.3.0", minor sürüm isteğe bağlı)
	addon_minimumNVDAVersion="2022.1",
	# Desteklenen/test edilmiş son NVDA sürümü (örn. "2025.2.0", ideal olarak minimumdan daha yeni)
	addon_lastTestedNVDAVersion="2025.3",
	# Eklenti güncelleme kanalı (varsayılan None: kararlı sürümler,
	# geliştirme sürümleri için "dev" kullanın.)
	# Ne yaptığınızdan emin değilseniz değiştirmeyin!
	addon_updateChannel=None,
	# Eklenti lisansı (örn. GPL 2)
	addon_license=None,
	# Eklentinin lisanslandığı lisans belgesi için URL
	addon_licenseURL=None,
)

# Eklentinizin kaynaklarını oluşturan Python dosyalarını tanımlayın.
# Her dosyayı tek tek listeleyebilir (yol ayırıcı olarak "/" kullanarak)
# veya glob desenleri kullanabilirsiniz.
# Örneğin eklentinizin "globalPlugins" klasöründeki tüm ".py" dosyalarını dahil etmek için
# listeyi şu şekilde yazabilirsiniz:
# pythonSources = ["addon/globalPlugins/KoruzBiz_MurText/*.py"]
# SCons Glob ifadeleri hakkında daha fazla bilgi için:
# https://scons.org/doc/production/HTML/scons-user/apd.html
pythonSources: list[str] = ["addon/globalPlugins/*.py"]

# Çeviri için dize içeren dosyalar. Genellikle Python kaynak dosyalarınız
i18nSources: list[str] = pythonSources + ["buildVars.py"]

# nvda-addon dosyası oluşturulurken yoksayılacak dosyalar
# Yollar, eklenti kaynaklarınızın kök dizinine değil eklenti dizinine göredir.
# Her dosyayı tek tek listeleyebilir (yol ayırıcı olarak "/")
# veya glob desenleri kullanabilirsiniz.
excludedFiles: list[str] = []

# NVDA eklentisi için temel dil  Eklentiniz İngilizce dışında bir dille yazılmışsa bu değişkeni düzenleyin.  Örneğin eklentiniz ağırlıklı olarak İspanyolca ise baseLanguage değerini "es" yapın.  Ayrıca yoksayılacak temel dil dosyalarını belirtmek için .gitignore dosyasını da düzenlemelisiniz. 
baseLanguage: str = "en"

# Eklenti dokümantasyonu için Markdown eklentileri  Çoğu eklenti ek Markdown eklentilerine ihtiyaç duymaz.  Tablolar gibi biçemlere destek eklemeniz gerekiyorsa aşağıdaki listeyi doldurun.  Uzantı dizeleri "markdown.extensions.uzantiAdi" biçiminde olmalıdır  örn. tablolar eklemek için "markdown.extensions.tables".
markdownExtensions: list[str] = []

# Özel braille çeviri tabloları
# Eklentiniz özel braille tabloları içeriyorsa (çoğu içermez) bu sözlüğü doldurun.
# Her anahtar braille tablo dosya adına göre adlandırılmış bir sözlüktür,
# içindeki anahtarlar şu öznitelikleri belirtir:
# displayName (kullanıcılara gösterilen ve çevrilebilir tablo adı),
# contracted (kısaltmalı (True) ya da kısaltmasız (False) braille kodu),
# output (çıktı tablo listesinde gösterimi),
# input (girdi tablo listesinde gösterimi).
brailleTables: BrailleTables = {}

# Özel konuşma sembol sözlükleri
# Sembol sözlüğü dosyaları locale klasöründe bulunur, örn. `locale\en`, ve `symbols-<ad>.dic` şeklinde adlandırılır.
# Eklentiniz özel konuşma sembol sözlükleri içeriyorsa (çoğu içermez) bu sözlüğü doldurun.
# Her anahtar sözlüğün adıdır,
# içindeki anahtarlar şu öznitelikleri belirtir:
# displayName (kullanıcılara gösterilen ve çevrilebilir sözlük adı),
# mandatory (Her zaman etkinse True, değilse False).
symbolDictionaries: SymbolDictionaries = {}

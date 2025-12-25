# _log.py
import os
import sys
import time
import traceback
import threading
import re

_MB = 1024 * 1024

def _default_log_dir(alt_klasor="Koruz_Biz\\murtext"):
    appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, alt_klasor)

def _ensure_dir(p):
    try:
        os.makedirs(p, exist_ok=True)
    except Exception:
        pass

def _rotate_if_needed(dosya_yolu, max_bytes=5 * _MB):
    try:
        if os.path.exists(dosya_yolu) and os.path.getsize(dosya_yolu) >= max_bytes:
            yedek = dosya_yolu + ".1"
            try:
                if os.path.exists(yedek):
                    os.remove(yedek)
            except Exception:
                pass
            try:
                os.replace(dosya_yolu, yedek)
            except Exception:
                pass
    except Exception:
        pass

class _DosyaYazici:
    def __init__(self, dosya_yolu, max_bytes=5 * _MB, encoding="utf-8"):
        self.dosya_yolu = dosya_yolu
        self.max_bytes = max_bytes
        self.encoding = encoding
        self._kilit = threading.RLock()

    def yaz(self, satir):
        with self._kilit:
            _ensure_dir(os.path.dirname(self.dosya_yolu))
            _rotate_if_needed(self.dosya_yolu, self.max_bytes)
            try:
                with open(self.dosya_yolu, "a", encoding=self.encoding, errors="replace") as f:
                    f.write(satir + "\n")
            except Exception:
                # Eklenti asla log yüzünden patlamasın
                pass

def _mesaj_temizle(msg, logger_adi):
    """
    Kullanıcı yanlışlıkla 'INFO MurText: ...' gibi prefix gönderirse tekrar oluşmasın.
    Ayrıca baş/son boşlukları da düzelt.
    """
    try:
        s = str(msg)
    except Exception:
        return ""

    s = s.strip()

    # Örnekler:
    # "INFO MurText: [Kopyala] ..." -> "[Kopyala] ..."
    # "ERROR MurText: ..." -> "..."
    # "INFO Koruzbiz_MurText: ..." -> "..."
    # Genel kural: "SEVİYE <ad>: " prefixini kırp
    try:
        # SEVİYE kısmını yakala (DEBUG/INFO/WARNING/ERROR/EXCEPTION)
        # ad kısmında boşluk olabilir (ör: "MurText") olmasın diye esnek tutuyoruz
        prefix_re = re.compile(r"^(DEBUG|INFO|WARNING|ERROR|EXCEPTION)\s+.+?:\s+", re.IGNORECASE)
        s = prefix_re.sub("", s, count=1).strip()
    except Exception:
        pass

    return s

class BasitLogger:
    def __init__(self, ad, logs_dir):
        self.ad = ad
        self.logs_dir = logs_dir
        _ensure_dir(logs_dir)

        self._all = _DosyaYazici(os.path.join(logs_dir, "logs.txt"))
        self._err = _DosyaYazici(os.path.join(logs_dir, "errors.txt"))
        self._dbg = _DosyaYazici(os.path.join(logs_dir, "debug.txt"))
        self._prt = _DosyaYazici(os.path.join(logs_dir, "print.txt"))

    def _ts(self):
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

    def _fmt(self, seviye, msg):
        # İstenen format:
        # DEBUG MurText: başla - 2025-12-24 19:56:57
        return f"{seviye} {self.ad}: {msg} - {self._ts()}"

    def debug(self, msg):
        msg = _mesaj_temizle(msg, self.ad)
        satir = self._fmt("DEBUG", msg)
        self._dbg.yaz(satir)
        self._all.yaz(satir)

    def info(self, msg):
        msg = _mesaj_temizle(msg, self.ad)
        satir = self._fmt("INFO", msg)
        self._prt.yaz(satir)
        self._all.yaz(satir)

    def warning(self, msg):
        msg = _mesaj_temizle(msg, self.ad)
        satir = self._fmt("WARNING", msg)
        self._err.yaz(satir)
        self._all.yaz(satir)

    def error(self, msg):
        msg = _mesaj_temizle(msg, self.ad)
        satir = self._fmt("ERROR", msg)
        self._err.yaz(satir)
        self._all.yaz(satir)

    def exception(self, msg, exc_info=None):
        if exc_info is None:
            exc_info = sys.exc_info()

        msg = _mesaj_temizle(msg, self.ad)
        satir = self._fmt("EXCEPTION", msg)
        self._err.yaz(satir)
        self._all.yaz(satir)

        try:
            tb = "".join(traceback.format_exception(*exc_info))
        except Exception:
            tb = "Traceback oluşturulamadı."

        self._err.yaz(tb)
        self._all.yaz(tb)

class _NoOpLogger:
    """
    Log kapalıyken dönen logger.
    Metotlar var ama hiçbir şey yazmaz.
    """
    def __init__(self, ad="MurText"):
        self.ad = ad

    def debug(self, msg):    return
    def info(self, msg):     return
    def warning(self, msg):  return
    def error(self, msg):    return
    def exception(self, msg, exc_info=None): return

class _StreamToLogger:
    def __init__(self, logger, level="INFO"):
        self.logger = logger
        self.level = level
        self._buf = ""

    def write(self, msg):
        self._buf += msg
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.rstrip("\r")
            if line:
                if self.level == "ERROR":
                    self.logger.error(line)
                else:
                    self.logger.info(line)

    def flush(self):
        if self._buf:
            if self.level == "ERROR":
                self.logger.error(self._buf.rstrip("\r"))
            else:
                self.logger.info(self._buf.rstrip("\r"))
            self._buf = ""

def baslat_loglama(
    appdata_dir=None,
    eklenti_adi="MurText",
    stdout_yonlendir=False,
    aktif=True,
    excepthook_kur=True
):
    """
    NVDA eklentisi için log başlat.

    aktif=False ise:
      - dosyaya yazmaz
      - stdout/stderr yönlendirme yapmaz
      - excepthook kurmaz
      - ama logger.info/error çağrıları patlamaz

    Dosyalar:
      - logs.txt: hepsi
      - errors.txt: warning/error/exception
      - debug.txt: debug
      - print.txt: info/print gibi
    """
    if not aktif:
        return _NoOpLogger(eklenti_adi)

    logs_dir = appdata_dir or _default_log_dir()
    logger = BasitLogger(eklenti_adi, logs_dir)

    if stdout_yonlendir:
        try:
            sys.stdout = _StreamToLogger(logger, "INFO")
            sys.stderr = _StreamToLogger(logger, "ERROR")
        except Exception:
            pass

    if excepthook_kur:
        try:
            def _excepthook(exc_type, exc, tb):
                logger.exception("Uncaught exception", exc_info=(exc_type, exc, tb))
            sys.excepthook = _excepthook
        except Exception:
            pass

    return logger

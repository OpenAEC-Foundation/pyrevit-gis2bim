# -*- coding: utf-8 -*-
"""File-logger voor Kozijnstaat — schrijft naar %TEMP%\\3bm_exchange.

Wordt gebruikt voor zelf-diagnose vanuit Claude Code. Elke run kan
aan begin reset() roepen om een schone log te krijgen.

IronPython 2.7 compatible — geen exotische strftime-codes,
alle calls zijn try/except gewrapt zodat falen van de logger
nooit het hoofdscript onderuit haalt.
"""

import os
import datetime
import traceback

try:
    LOG_DIR = os.path.join(
        os.environ.get("TEMP", os.path.expanduser("~")),
        "3bm_exchange",
    )
except Exception:
    LOG_DIR = "."

LOG_FILE = os.path.join(LOG_DIR, "kozijnstaat_debug.log")


def get_log_dir():
    return LOG_DIR


def _ensure_dir(directory=None):
    target = directory or LOG_DIR
    try:
        if not os.path.isdir(target):
            os.makedirs(target)
    except Exception:
        pass


def _timestamp():
    try:
        now = datetime.datetime.now()
        return "{0:04d}-{1:02d}-{2:02d} {3:02d}:{4:02d}:{5:02d}".format(
            now.year, now.month, now.day,
            now.hour, now.minute, now.second,
        )
    except Exception:
        return "????-??-?? ??:??:??"


def _safe_str(value):
    try:
        if isinstance(value, bytes):
            return value.decode("utf-8", "replace")
        if isinstance(value, str):
            return value
        try:
            return unicode(value)
        except Exception:
            return str(value)
    except Exception:
        return "<unprintable>"


def _write(level, message):
    try:
        _ensure_dir()
        line = u"[{0}] {1}: {2}\n".format(
            _timestamp(), level, _safe_str(message),
        )
        try:
            f = open(LOG_FILE, "a")
        except Exception:
            return
        try:
            try:
                f.write(line.encode("utf-8"))
            except Exception:
                try:
                    f.write(line)
                except Exception:
                    pass
            try:
                f.flush()
            except Exception:
                pass
        finally:
            try:
                f.close()
            except Exception:
                pass
    except Exception:
        pass


def info(message):
    _write("INFO", message)


def warn(message):
    _write("WARN", message)


def error(message):
    _write("ERROR", message)


def exc(message):
    """Log een message + huidige exception traceback."""
    try:
        tb = traceback.format_exc()
    except Exception:
        tb = "<traceback unavailable>"
    _write("ERROR", u"{0}\n{1}".format(_safe_str(message), _safe_str(tb)))


def reset(header=None):
    """Schoon de log + zet optioneel een header-regel."""
    try:
        _ensure_dir()
        try:
            f = open(LOG_FILE, "w")
        except Exception:
            return
        try:
            if header:
                line = u"=== {0} | {1} ===\n".format(
                    _safe_str(header), _timestamp(),
                )
                try:
                    f.write(line.encode("utf-8"))
                except Exception:
                    try:
                        f.write(line)
                    except Exception:
                        pass
        finally:
            try:
                f.close()
            except Exception:
                pass
    except Exception:
        pass


def get_log_path():
    return LOG_FILE


# --- Per-tool logger factory ----------------------------------------------
#
# Voor tools die hun eigen log-file willen (i.p.v. de gedeelde
# kozijnstaat_debug.log), bv. zodat reset() van andere tools de output
# niet wist. Returnt een dun object met dezelfde interface als de
# module-level functies: info/warn/error/exc/reset/get_log_path/session.

class _ToolLogger(object):
    def __init__(self, filename):
        try:
            self.path = os.path.join(LOG_DIR, filename)
        except Exception:
            self.path = filename

    def _write(self, level, message):
        try:
            _ensure_dir()
            line = u"[{0}] {1}: {2}\n".format(
                _timestamp(), level, _safe_str(message),
            )
            try:
                f = open(self.path, "a")
            except Exception:
                return
            try:
                try:
                    f.write(line.encode("utf-8"))
                except Exception:
                    try:
                        f.write(line)
                    except Exception:
                        pass
                try:
                    f.flush()
                except Exception:
                    pass
            finally:
                try:
                    f.close()
                except Exception:
                    pass
        except Exception:
            pass

    def info(self, message):
        self._write("INFO", message)

    def warn(self, message):
        self._write("WARN", message)

    def error(self, message):
        self._write("ERROR", message)

    def exc(self, message):
        try:
            tb = traceback.format_exc()
        except Exception:
            tb = "<traceback unavailable>"
        self._write(
            "ERROR", u"{0}\n{1}".format(_safe_str(message), _safe_str(tb)),
        )

    def session(self, header):
        """Append een session-separator (geen reset — log blijft bewaard)."""
        try:
            _ensure_dir()
            try:
                f = open(self.path, "a")
            except Exception:
                return
            try:
                line = u"\n===== {0} | {1} =====\n".format(
                    _safe_str(header), _timestamp(),
                )
                try:
                    f.write(line.encode("utf-8"))
                except Exception:
                    try:
                        f.write(line)
                    except Exception:
                        pass
            finally:
                try:
                    f.close()
                except Exception:
                    pass
        except Exception:
            pass

    def reset(self, header=None):
        try:
            _ensure_dir()
            try:
                f = open(self.path, "w")
            except Exception:
                return
            try:
                if header:
                    line = u"=== {0} | {1} ===\n".format(
                        _safe_str(header), _timestamp(),
                    )
                    try:
                        f.write(line.encode("utf-8"))
                    except Exception:
                        try:
                            f.write(line)
                        except Exception:
                            pass
            finally:
                try:
                    f.close()
                except Exception:
                    pass
        except Exception:
            pass

    def get_log_path(self):
        return self.path


def make_logger(filename):
    """Geef een dedicated logger voor een specifieke tool.

    Voorbeeld:
        from kozijnstaat import logger as klog
        log = klog.make_logger("kozijnstaat_legend.log")
        log.session("KozijnstaatLegend start")
        log.info("config={0}".format(cfg))
    """
    return _ToolLogger(filename)

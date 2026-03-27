# -*- coding: utf-8 -*-
"""3BM Bouwkunde - Centrale Logging Module

Gedeelde logging voor alle pyRevit tools.
Logs worden opgeslagen in een centrale folder voor MCP access.

Gebruik in scripts:
    from lib.bm_logger import get_logger
    log = get_logger("AutoDim")
    log.info("Tool gestart")
"""

import os
import sys
import datetime
import traceback
import codecs

# =============================================================================
# CONFIGURATIE
# =============================================================================

LOG_PATHS = [
    r"Z:\50_projecten\7_3BM_bouwkunde\_AI_logs\pyrevit",
    r"X:\10_3BM_bouwkunde\50_Claude-Code-Projects\pyrevit\logs",
    r"C:\DATA\3BM_projecten\50_projecten\7_3BM_bouwkunde\_AI_logs\pyrevit",
    os.path.join(os.environ.get('APPDATA', ''), 'bouwkunde', 'logs'),
]

FALLBACK_LOG_PATH = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), '3BM_pyrevit_logs')

MAX_LOGS_PER_TOOL = 10
MAX_LOG_SIZE_MB = 5
LOG_LEVEL = "DEBUG"
PRINT_TO_CONSOLE = False  # Zet op False om print errors te voorkomen


# =============================================================================
# LOG LEVELS
# =============================================================================
class LogLevel:
    DEBUG = 10
    INFO = 20
    WARNING = 30
    ERROR = 40
    
    _names = {10: "DEBUG", 20: "INFO", 30: "WARNING", 40: "ERROR"}
    _values = {"DEBUG": 10, "INFO": 20, "WARNING": 30, "ERROR": 40}
    
    @classmethod
    def get_name(cls, level):
        return cls._names.get(level, "UNKNOWN")
    
    @classmethod
    def get_value(cls, name):
        return cls._values.get(name.upper(), 20)


# =============================================================================
# LOGGER CLASS
# =============================================================================
class BMLogger:
    """Logger voor 3BM pyRevit tools."""
    
    def __init__(self, tool_name):
        self.tool_name = tool_name
        self.log_level = LogLevel.get_value(LOG_LEVEL)
        self.log_dir = self._get_log_directory()
        self._log_file_path = self._create_log_file()
        self._cleanup_old_logs()
        self._write_header()

    def __call__(self, message, **kwargs):
        """Maakt log("msg") mogelijk als alias voor log.info("msg")."""
        self.info(message, **kwargs)

    @property
    def log_file(self):
        """Pad naar het huidige log bestand."""
        return self._log_file_path

    @log_file.setter
    def log_file(self, value):
        self._log_file_path = value
    
    def _get_log_directory(self):
        """Vind of maak de log directory."""
        for path in LOG_PATHS:
            try:
                if os.path.exists(os.path.dirname(path)) or os.path.exists(path):
                    if not os.path.exists(path):
                        os.makedirs(path)
                    test_file = os.path.join(path, '.write_test')
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                    return path
            except:
                continue
        
        if not os.path.exists(FALLBACK_LOG_PATH):
            os.makedirs(FALLBACK_LOG_PATH)
        return FALLBACK_LOG_PATH
    
    def _create_log_file(self):
        """Maak nieuw log bestand."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = "{}_{}.log".format(self.tool_name, timestamp)
        return os.path.join(self.log_dir, filename)
    
    def _cleanup_old_logs(self):
        """Verwijder oude logs."""
        try:
            pattern = "{}_".format(self.tool_name)
            logs = []
            
            for f in os.listdir(self.log_dir):
                if f.startswith(pattern) and f.endswith('.log'):
                    full_path = os.path.join(self.log_dir, f)
                    logs.append((full_path, os.path.getmtime(full_path)))
            
            logs.sort(key=lambda x: x[1])
            
            while len(logs) >= MAX_LOGS_PER_TOOL:
                old_log = logs.pop(0)[0]
                try:
                    os.remove(old_log)
                except:
                    pass
        except:
            pass
    
    def _write_header(self):
        """Schrijf log header."""
        header = [
            "=" * 70,
            "3BM BOUWKUNDE - {} LOG".format(self.tool_name.upper()),
            "=" * 70,
            "Timestamp: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            "Log file: {}".format(self.log_file),
            "Log level: {}".format(LOG_LEVEL),
            "-" * 70,
            ""
        ]
        self._write_lines(header)
    
    def _write_lines(self, lines):
        """Schrijf regels naar log file."""
        try:
            with codecs.open(self.log_file, 'a', encoding='utf-8') as f:
                for line in lines:
                    f.write(line + "\n")
        except:
            pass  # Silently fail
    
    def _safe_print(self, line):
        """Veilige print die niet crasht."""
        if not PRINT_TO_CONSOLE:
            return
        try:
            print(line)
        except:
            pass  # Silently fail als output window problemen heeft
    
    def _format_message(self, level, message, **kwargs):
        """Formatteer log message."""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        level_name = LogLevel.get_name(level)
        
        lines = ["[{}] [{}] {}".format(timestamp, level_name, message)]
        
        if kwargs.get('data'):
            lines.append("  DATA: {}".format(kwargs['data']))
        
        if kwargs.get('exc_info'):
            lines.append("  EXCEPTION:")
            for line in traceback.format_exc().split('\n'):
                if line.strip():
                    lines.append("    {}".format(line))
        
        return lines
    
    def _log(self, level, message, **kwargs):
        """Interne log methode."""
        if level >= self.log_level:
            lines = self._format_message(level, message, **kwargs)
            self._write_lines(lines)
            
            for line in lines:
                self._safe_print(line)
    
    def debug(self, message, **kwargs):
        self._log(LogLevel.DEBUG, message, **kwargs)
    
    def info(self, message, **kwargs):
        self._log(LogLevel.INFO, message, **kwargs)
    
    def warning(self, message, **kwargs):
        self._log(LogLevel.WARNING, message, **kwargs)
    
    def warn(self, message, **kwargs):
        self.warning(message, **kwargs)
    
    def error(self, message, **kwargs):
        self._log(LogLevel.ERROR, message, **kwargs)
    
    def exception(self, message):
        self._log(LogLevel.ERROR, message, exc_info=True)
    
    def separator(self, char="-", length=50):
        self._write_lines([char * length])
    
    def section(self, title):
        self._write_lines([
            "",
            ">>> {} <<<".format(title.upper()),
            ""
        ])
    
    def log_revit_info(self):
        """Log Revit document info."""
        try:
            from pyrevit import revit
            doc = revit.doc
            
            self.section("Revit Info")
            self.info("Document: {}".format(doc.Title if doc else "None"))
            self.info("Path: {}".format(doc.PathName if doc and doc.PathName else "Not saved"))
            
            view = revit.active_view
            if view:
                self.info("Active View: {} ({})".format(view.Name, view.ViewType))
        except:
            pass
    
    def log_selection(self, elements, label="Selection"):
        self.info("{}: {} elements".format(label, len(elements) if elements else 0))
    
    def log_options(self, options_dict):
        self.section("Options")
        for key, value in options_dict.items():
            self.info("  {}: {}".format(key, value))
    
    def finalize(self, success=True, message=None):
        """Sluit log af."""
        self._write_lines([
            "",
            "-" * 70,
            "RESULT: {}".format("SUCCESS" if success else "FAILED"),
        ])
        
        if message:
            self._write_lines(["MESSAGE: {}".format(message)])
        
        self._write_lines([
            "End time: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            "Log file: {}".format(self.log_file),
            "=" * 70
        ])
        
        return self.log_file


# =============================================================================
# FACTORY
# =============================================================================
_loggers = {}

def get_logger(tool_name):
    """Verkrijg logger instance."""
    if tool_name not in _loggers:
        _loggers[tool_name] = BMLogger(tool_name)
    return _loggers[tool_name]


def get_log_directory():
    """Geef log directory pad."""
    temp_logger = BMLogger("_temp")
    log_dir = temp_logger.log_dir
    try:
        os.remove(temp_logger.log_file)
    except:
        pass
    return log_dir

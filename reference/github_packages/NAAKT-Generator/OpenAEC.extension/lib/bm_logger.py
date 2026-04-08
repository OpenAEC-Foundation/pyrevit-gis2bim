# -*- coding: utf-8 -*-
"""OpenAEC - Logging Module

Shared logging for all pyRevit tools.
Logs are saved to a central folder.

Usage:
    from bm_logger import get_logger
    log = get_logger("ToolName")
    log.info("Tool started")
"""

import os
import sys
import datetime
import traceback
import codecs

# =============================================================================
# CONFIGURATION
# =============================================================================

# Log directory: %APPDATA%\OpenAEC\logs
LOG_DIR = os.path.join(os.environ.get('APPDATA', ''), 'OpenAEC', 'logs')
FALLBACK_LOG_DIR = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), 'OpenAEC_logs')

MAX_LOGS_PER_TOOL = 10
MAX_LOG_SIZE_MB = 5
LOG_LEVEL = "DEBUG"
PRINT_TO_CONSOLE = False


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
    """Logger for OpenAEC pyRevit tools."""
    
    def __init__(self, tool_name):
        self.tool_name = tool_name
        self.log_level = LogLevel.get_value(LOG_LEVEL)
        self.log_dir = self._get_log_directory()
        self.log_file = self._create_log_file()
        self._cleanup_old_logs()
        self._write_header()
    
    def _get_log_directory(self):
        """Find or create the log directory."""
        # Try primary location
        try:
            if not os.path.exists(LOG_DIR):
                os.makedirs(LOG_DIR)
            # Test write access
            test_file = os.path.join(LOG_DIR, '.write_test')
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            return LOG_DIR
        except:
            pass
        
        # Fallback to temp
        try:
            if not os.path.exists(FALLBACK_LOG_DIR):
                os.makedirs(FALLBACK_LOG_DIR)
            return FALLBACK_LOG_DIR
        except:
            return os.environ.get('TEMP', 'C:\\Temp')
    
    def _create_log_file(self):
        """Create new log file."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = "{}_{}.log".format(self.tool_name, timestamp)
        return os.path.join(self.log_dir, filename)
    
    def _cleanup_old_logs(self):
        """Remove old logs."""
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
        """Write log header."""
        header = [
            "=" * 70,
            "OPENAEC - {} LOG".format(self.tool_name.upper()),
            "=" * 70,
            "Timestamp: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            "Log file: {}".format(self.log_file),
            "Log level: {}".format(LOG_LEVEL),
            "-" * 70,
            ""
        ]
        self._write_lines(header)
    
    def _write_lines(self, lines):
        """Write lines to log file."""
        try:
            with codecs.open(self.log_file, 'a', encoding='utf-8') as f:
                for line in lines:
                    f.write(line + "\n")
        except:
            pass
    
    def _safe_print(self, line):
        """Safe print that doesn't crash."""
        if not PRINT_TO_CONSOLE:
            return
        try:
            print(line)
        except:
            pass
    
    def _format_message(self, level, message, **kwargs):
        """Format log message."""
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
        """Internal log method."""
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
        """Close log."""
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
    """Get logger instance."""
    if tool_name not in _loggers:
        _loggers[tool_name] = BMLogger(tool_name)
    return _loggers[tool_name]


def get_log_directory():
    """Get log directory path."""
    temp_logger = BMLogger("_temp")
    log_dir = temp_logger.log_dir
    try:
        os.remove(temp_logger.log_file)
    except:
        pass
    return log_dir

# -*- coding: utf-8 -*-
"""
MCP Status Tool - Controleer Revit MCP Server status
WPF versie met 3BM huisstijl
"""
__title__ = "MCP"
__author__ = "3BM Bouwkunde"
__doc__ = "Controleer de status van de Revit MCP Server"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color

import sys
import os
import socket
import datetime

# Lib path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib'))
from bm_logger import get_logger

log = get_logger("MCPStatus")

# Configuratie
MCP_HOST = "localhost"
MCP_PORT = 48884  # Standaard Revit MCP port


# ==============================================================================
# KLEUREN
# ==============================================================================
def get_color(hex_color):
    """Maak Color van hex string"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return Color.FromRgb(r, g, b)

TEAL = get_color("#45B6A8")
RED = get_color("#DB4C40")
GRAY = get_color("#808080")
LIGHT_TEAL_BG = get_color("#E8F6F4")
LIGHT_RED_BG = get_color("#FDEDEC")


# ==============================================================================
# MCP STATUS CHECK
# ==============================================================================
def check_mcp_status(host, port):
    """Check of MCP server bereikbaar is via socket"""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        result = sock.connect_ex((host, port))
        sock.close()
        return result == 0
    except Exception as e:
        log.warning("Socket check failed: {}".format(e))
        return False


# ==============================================================================
# MAIN WINDOW
# ==============================================================================
class MCPStatusWindow(Window):
    """MCP Server Status Checker - WPF versie"""
    
    def __init__(self):
        Window.__init__(self)
        self._load_xaml()
        self._setup_config()
        self._bind_events()
        self._check_status()
    
    def _load_xaml(self):
        """Laad XAML layout"""
        xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        
        try:
            with open(xaml_path, 'r') as f:
                xaml_content = f.read()
            
            reader = StringReader(xaml_content)
            loaded = XamlReader.Load(SysXmlReader.Create(reader))
            
            # Copy window properties
            self.Title = loaded.Title
            self.Width = loaded.Width
            self.SizeToContent = loaded.SizeToContent
            self.WindowStartupLocation = loaded.WindowStartupLocation
            self.ResizeMode = loaded.ResizeMode
            self.Background = loaded.Background
            self.Content = loaded.Content
            
            # Bind named elements
            self.txt_host = loaded.FindName('txt_host')
            self.txt_port = loaded.FindName('txt_port')
            self.status_border = loaded.FindName('status_border')
            self.txt_status_icon = loaded.FindName('txt_status_icon')
            self.txt_status = loaded.FindName('txt_status')
            self.txt_status_desc = loaded.FindName('txt_status_desc')
            self.txt_last_check = loaded.FindName('txt_last_check')
            self.btn_refresh = loaded.FindName('btn_refresh')
            self.btn_close = loaded.FindName('btn_close')
            
            log.info("XAML loaded successfully")
            
        except Exception as e:
            log.error("XAML load error: {}".format(e), exc_info=True)
            raise
    
    def _setup_config(self):
        """Toon server configuratie"""
        if self.txt_host:
            self.txt_host.Text = MCP_HOST
        if self.txt_port:
            self.txt_port.Text = str(MCP_PORT)
    
    def _bind_events(self):
        """Bind event handlers"""
        if self.btn_refresh:
            self.btn_refresh.Click += self._on_refresh
        if self.btn_close:
            self.btn_close.Click += self._on_close
    
    def _check_status(self):
        """Check MCP server status en update UI"""
        log.info("Checking MCP server status on {}:{}".format(MCP_HOST, MCP_PORT))
        
        is_online = check_mcp_status(MCP_HOST, MCP_PORT)
        now = datetime.datetime.now().strftime("%H:%M:%S")
        
        if is_online:
            # Online status
            self.txt_status_icon.Text = "✓"
            self.txt_status_icon.Foreground = SolidColorBrush(TEAL)
            self.txt_status.Text = "ONLINE"
            self.txt_status.Foreground = SolidColorBrush(TEAL)
            self.txt_status_desc.Text = "MCP Server is actief en bereikbaar"
            self.status_border.Background = SolidColorBrush(LIGHT_TEAL_BG)
            log.info("MCP server is ONLINE")
        else:
            # Offline status
            self.txt_status_icon.Text = "✗"
            self.txt_status_icon.Foreground = SolidColorBrush(RED)
            self.txt_status.Text = "OFFLINE"
            self.txt_status.Foreground = SolidColorBrush(RED)
            self.txt_status_desc.Text = "MCP Server is niet bereikbaar"
            self.status_border.Background = SolidColorBrush(LIGHT_RED_BG)
            log.warning("MCP server is OFFLINE")
        
        self.txt_last_check.Text = "Laatste check: {}".format(now)
    
    def _on_refresh(self, sender, args):
        """Ververs status"""
        log.info("Refresh clicked")
        self._check_status()
    
    def _on_close(self, sender, args):
        """Sluit venster"""
        self.Close()


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    """Entry point"""
    log.info("=== MCP Status Tool Started (WPF) ===")
    
    try:
        window = MCPStatusWindow()
        window.ShowDialog()
    except Exception as e:
        log.error("Error: {}".format(e), exc_info=True)
        from pyrevit import forms
        forms.alert("Fout bij laden MCP Status:\n\n{}".format(e), title="Error")
    
    log.info("=== MCP Status Tool Closed ===")


if __name__ == "__main__":
    main()

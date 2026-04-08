# -*- coding: utf-8 -*-
"""
XAML Window Helper - GIS2BIM
============================

Gedeelde functies voor het laden van WPF XAML windows
in pyRevit IronPython tools.

Gebruik:
    from gis2bim.ui.xaml_helper import load_xaml_window, bind_ui_elements

    class MijnWindow(Window):
        def __init__(self):
            Window.__init__(self)
            load_xaml_window(self, xaml_path)
            bind_ui_elements(self, self, ['btn_ok', 'txt_name'])
"""

from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows.Markup import XamlReader


def load_xaml_window(window, xaml_path):
    """Laad een XAML bestand en pas de properties toe op een Window.

    Leest het XAML bestand, parst het, en kopieert de relevante
    window properties (Title, Width, Content, etc.) naar het
    gegeven Window object.

    Args:
        window: Het Window object (self in de tool)
        xaml_path: Absoluut pad naar het UI.xaml bestand

    Returns:
        Het geladen XAML root element (voor bind_ui_elements)
    """
    with open(xaml_path, 'r') as f:
        xaml_content = f.read()

    reader = StringReader(xaml_content)
    xml_reader = SysXmlReader.Create(reader)
    loaded = XamlReader.Load(xml_reader)

    window.Title = loaded.Title
    window.Width = loaded.Width
    window.SizeToContent = loaded.SizeToContent
    window.WindowStartupLocation = loaded.WindowStartupLocation
    window.ResizeMode = loaded.ResizeMode
    window.Background = loaded.Background
    window.Content = loaded.Content

    return loaded


def bind_ui_elements(target, root, element_names):
    """Bind XAML elementen aan attributen op het target object.

    Zoekt elk element op naam in de XAML root en stelt het in
    als attribuut op target.

    Args:
        target: Object waarop attributen worden gezet (meestal self)
        root: XAML root element (resultaat van load_xaml_window)
        element_names: Lijst van element namen (x:Name in XAML)
    """
    for name in element_names:
        element = root.FindName(name)
        if element:
            setattr(target, name, element)

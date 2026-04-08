# -*- coding: utf-8 -*-
"""
WPF Template voor 3BM Bouwkunde pyRevit tools
Vervangt Windows Forms met moderne WPF UI

Gebruik:
    from wpf_template import WPFWindow, Huisstijl
    
    class MijnToolWindow(WPFWindow):
        def __init__(self):
            xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
            super(MijnToolWindow, self).__init__(xaml_file, "Tool Titel")
            
            # Bind events
            self.btn_execute.Click += self.on_execute
        
        def on_execute(self, sender, args):
            # Doe iets
            self.close_ok()
"""

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
clr.AddReference('System.Xml')

import System
from System.IO import StringReader
from System.Xml import XmlReader as SysXmlReader
from System.Windows import Window, Application, ResourceDictionary, Thickness, SizeToContent
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color, BrushConverter
from System.Windows.Controls import Button, TextBox, ComboBox, CheckBox, Label

import os


# ==============================================================================
# HUISSTIJL KLEUREN
# ==============================================================================
class Huisstijl:
    """3BM Bouwkunde huisstijl kleuren - WPF versie"""
    
    # Primaire kleuren (hex)
    VIOLET_HEX = "#350E35"
    TEAL_HEX = "#45B6A8"
    YELLOW_HEX = "#EFBD75"
    MAGENTA_HEX = "#A01C48"
    PEACH_HEX = "#DB4C40"
    
    # Neutrale kleuren
    WHITE_HEX = "#FFFFFF"
    LIGHT_GRAY_HEX = "#F5F5F5"
    MEDIUM_GRAY_HEX = "#E0E0E0"
    TEXT_PRIMARY_HEX = "#323232"
    TEXT_SECONDARY_HEX = "#808080"
    
    @staticmethod
    def get_brush(hex_color):
        """Maak SolidColorBrush van hex kleur"""
        converter = BrushConverter()
        return converter.ConvertFromString(hex_color)
    
    @staticmethod
    def get_color(hex_color):
        """Maak Color van hex kleur"""
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return Color.FromRgb(r, g, b)
    
    # Pre-made brushes (lazy loading)
    _brushes = {}
    
    @classmethod
    def brush(cls, name):
        """Haal brush op naam: 'violet', 'teal', 'yellow', etc."""
        if name not in cls._brushes:
            hex_map = {
                'violet': cls.VIOLET_HEX,
                'teal': cls.TEAL_HEX,
                'yellow': cls.YELLOW_HEX,
                'magenta': cls.MAGENTA_HEX,
                'peach': cls.PEACH_HEX,
                'white': cls.WHITE_HEX,
                'light_gray': cls.LIGHT_GRAY_HEX,
                'medium_gray': cls.MEDIUM_GRAY_HEX,
                'text_primary': cls.TEXT_PRIMARY_HEX,
                'text_secondary': cls.TEXT_SECONDARY_HEX,
            }
            cls._brushes[name] = cls.get_brush(hex_map.get(name, cls.TEXT_PRIMARY_HEX))
        return cls._brushes[name]


# ==============================================================================
# BASE XAML STYLE
# ==============================================================================
BASE_STYLE_XAML = '''
<ResourceDictionary 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    
    <!-- Kleuren -->
    <Color x:Key="VioletColor">#350E35</Color>
    <Color x:Key="TealColor">#45B6A8</Color>
    <Color x:Key="YellowColor">#EFBD75</Color>
    <Color x:Key="MagentaColor">#A01C48</Color>
    <Color x:Key="PeachColor">#DB4C40</Color>
    <Color x:Key="LightGrayColor">#F5F5F5</Color>
    <Color x:Key="MediumGrayColor">#E0E0E0</Color>
    <Color x:Key="TextPrimaryColor">#323232</Color>
    <Color x:Key="TextSecondaryColor">#808080</Color>
    
    <SolidColorBrush x:Key="VioletBrush" Color="{StaticResource VioletColor}"/>
    <SolidColorBrush x:Key="TealBrush" Color="{StaticResource TealColor}"/>
    <SolidColorBrush x:Key="YellowBrush" Color="{StaticResource YellowColor}"/>
    <SolidColorBrush x:Key="MagentaBrush" Color="{StaticResource MagentaColor}"/>
    <SolidColorBrush x:Key="PeachBrush" Color="{StaticResource PeachColor}"/>
    <SolidColorBrush x:Key="LightGrayBrush" Color="{StaticResource LightGrayColor}"/>
    <SolidColorBrush x:Key="MediumGrayBrush" Color="{StaticResource MediumGrayColor}"/>
    <SolidColorBrush x:Key="TextPrimaryBrush" Color="{StaticResource TextPrimaryColor}"/>
    <SolidColorBrush x:Key="TextSecondaryBrush" Color="{StaticResource TextSecondaryColor}"/>
    
    <!-- Button Styles -->
    <Style x:Key="PrimaryButton" TargetType="Button">
        <Setter Property="Background" Value="{StaticResource TealBrush}"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Padding" Value="20,8"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            CornerRadius="3" 
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#3DA396"/>
                        </Trigger>
                        <Trigger Property="IsPressed" Value="True">
                            <Setter Property="Background" Value="#359384"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
    <Style x:Key="SecondaryButton" TargetType="Button">
        <Setter Property="Background" Value="White"/>
        <Setter Property="Foreground" Value="{StaticResource VioletBrush}"/>
        <Setter Property="FontWeight" Value="Medium"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Padding" Value="20,8"/>
        <Setter Property="BorderBrush" Value="{StaticResource VioletBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="3" 
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="{StaticResource LightGrayBrush}"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
    <Style x:Key="DangerButton" TargetType="Button">
        <Setter Property="Background" Value="{StaticResource PeachBrush}"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="FontWeight" Value="Medium"/>
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Padding" Value="20,8"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" 
                            CornerRadius="3" 
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#C94438"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    
    <!-- TextBox Style -->
    <Style x:Key="StandardTextBox" TargetType="TextBox">
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Padding" Value="8,6"/>
        <Setter Property="BorderBrush" Value="{StaticResource MediumGrayBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Background" Value="White"/>
    </Style>
    
    <!-- ComboBox Style -->
    <Style x:Key="StandardComboBox" TargetType="ComboBox">
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Padding" Value="8,6"/>
        <Setter Property="BorderBrush" Value="{StaticResource MediumGrayBrush}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Background" Value="White"/>
    </Style>
    
    <!-- Label Styles -->
    <Style x:Key="SectionHeader" TargetType="TextBlock">
        <Setter Property="FontSize" Value="12"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Foreground" Value="{StaticResource VioletBrush}"/>
        <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    
    <Style x:Key="FieldLabel" TargetType="TextBlock">
        <Setter Property="FontSize" Value="11"/>
        <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
        <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>
    
    <!-- CheckBox Style -->
    <Style x:Key="StandardCheckBox" TargetType="CheckBox">
        <Setter Property="FontSize" Value="13"/>
        <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    
</ResourceDictionary>
'''


# ==============================================================================
# WPF WINDOW BASE CLASS
# ==============================================================================
class WPFWindow(Window):
    """
    Base class voor WPF windows in pyRevit.
    
    Gebruik:
        class MijnWindow(WPFWindow):
            def __init__(self):
                xaml_path = os.path.join(os.path.dirname(__file__), 'UI.xaml')
                super(MijnWindow, self).__init__(xaml_path, "Mijn Tool")
    """
    
    def __init__(self, xaml_file=None, title="3BM Tool", width=500, height=400):
        """
        Initialize WPF Window.
        
        Args:
            xaml_file: Pad naar XAML bestand (optioneel)
            title: Window titel
            width: Breedte in pixels
            height: Hoogte in pixels (of None voor SizeToContent)
        """
        Window.__init__(self)
        
        # Load base styles
        self._load_base_styles()
        
        # Window properties
        self.Title = title
        self.Width = width
        if height:
            self.Height = height
        else:
            self.SizeToContent = SizeToContent.Height
        
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip
        
        # Result tracking
        self._result = None
        self._dialog_result = False
        
        # Load XAML if provided
        if xaml_file and os.path.exists(xaml_file):
            self._load_xaml(xaml_file)
    
    def _load_base_styles(self):
        """Laad de basis 3BM styles"""
        try:
            reader = StringReader(BASE_STYLE_XAML)
            xml_reader = SysXmlReader.Create(reader)
            resource_dict = XamlReader.Load(xml_reader)
            self.Resources.MergedDictionaries.Add(resource_dict)
        except Exception as e:
            print("Kon base styles niet laden: {}".format(e))
    
    def _load_xaml(self, xaml_file):
        """Laad XAML bestand en bind aan window"""
        try:
            with open(xaml_file, 'r') as f:
                xaml_content = f.read()
            
            reader = StringReader(xaml_content)
            xml_reader = SysXmlReader.Create(reader)
            content = XamlReader.Load(xml_reader)
            self.Content = content
            
            # Auto-bind named elements
            self._bind_elements(content)
            
        except Exception as e:
            print("XAML load error: {}".format(e))
            raise
    
    def _bind_elements(self, element, prefix=""):
        """Recursief bind named elements als attributen"""
        try:
            name = element.Name
            if name:
                setattr(self, name, element)
        except:
            pass
        
        # Recurse children
        try:
            if hasattr(element, 'Children'):
                for child in element.Children:
                    self._bind_elements(child)
            elif hasattr(element, 'Child') and element.Child:
                self._bind_elements(element.Child)
            elif hasattr(element, 'Content') and element.Content:
                if hasattr(element.Content, 'Children') or hasattr(element.Content, 'Name'):
                    self._bind_elements(element.Content)
        except:
            pass
    
    def find_element(self, name):
        """Zoek element op naam in de visual tree"""
        return self.FindName(name)
    
    @property
    def result(self):
        """Haal resultaat op na dialog close"""
        return self._result
    
    @result.setter
    def result(self, value):
        self._result = value
    
    def close_ok(self):
        """Sluit window met OK resultaat"""
        self._dialog_result = True
        self.DialogResult = True
        self.Close()
    
    def close_cancel(self):
        """Sluit window met Cancel resultaat"""
        self._dialog_result = False
        self.DialogResult = False
        self.Close()
    
    def show_dialog(self):
        """Toon als modal dialog, return True als OK"""
        self.ShowDialog()
        return self._dialog_result
    
    # Convenience methods voor messages
    def show_info(self, message, title="Info"):
        """Toon info message box"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information)
    
    def show_warning(self, message, title="Waarschuwing"):
        """Toon warning message box"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning)
    
    def show_error(self, message, title="Fout"):
        """Toon error message box"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error)
    
    def ask_confirm(self, message, title="Bevestig"):
        """Vraag bevestiging, return True/False"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
        result = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question)
        return result == MessageBoxResult.Yes


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def load_xaml_file(xaml_path):
    """Load XAML file and return root element"""
    with open(xaml_path, 'r') as f:
        xaml_content = f.read()
    reader = StringReader(xaml_content)
    xml_reader = SysXmlReader.Create(reader)
    return XamlReader.Load(xml_reader)


def create_simple_window(title, width=400, height=300):
    """Create a simple window without XAML"""
    window = WPFWindow(title=title, width=width, height=height)
    return window

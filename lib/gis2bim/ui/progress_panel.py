# -*- coding: utf-8 -*-
"""
Progress Panel - GIS2BIM
========================

Gedeelde progress/status panel functies voor GIS2BIM pyRevit tools.

Vereist dat het Window de volgende attributen heeft:
    - pnl_progress (StackPanel)
    - txt_progress (TextBlock)
    - txt_status (TextBlock)

Gebruik:
    from gis2bim.ui.progress_panel import show_progress, hide_progress, update_ui

    show_progress(self, "Downloaden...")
    hide_progress(self)
"""

import System
from System.Windows import Visibility


def show_progress(window, message):
    """Toon progress panel met bericht.

    Args:
        window: Het Window object met pnl_progress en txt_progress
        message: Tekst om te tonen
    """
    window.pnl_progress.Visibility = Visibility.Visible
    window.txt_progress.Text = message
    window.txt_status.Text = ""
    update_ui()


def hide_progress(window):
    """Verberg progress panel.

    Args:
        window: Het Window object met pnl_progress
    """
    window.pnl_progress.Visibility = Visibility.Collapsed


def update_ui():
    """Forceer een WPF UI update (Dispatcher pump).

    Nodig zodat progress tekst zichtbaar wordt tijdens
    langlopende operaties in de UI thread.
    """
    try:
        from System.Windows.Threading import Dispatcher, DispatcherPriority
        Dispatcher.CurrentDispatcher.Invoke(
            System.Action(lambda: None),
            DispatcherPriority.Render
        )
    except Exception:
        pass

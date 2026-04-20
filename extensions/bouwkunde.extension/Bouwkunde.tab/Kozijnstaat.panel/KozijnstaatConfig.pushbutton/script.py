# -*- coding: utf-8 -*-
"""Kozijnstaat - Config.

WinForms dialoog om de Kozijnstaat-instellingen te bewerken. Gebruikt
het gedeelde BaseForm/UIFactory uit lib/ui_template.py voor huisstijl.

IronPython 2.7.
"""

__title__ = "Config"
__author__ = "3BM Bouwkunde"
__doc__ = "Bewerk Kozijnstaat config (family namen, grid, offsets, refs)"

import os
import sys

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
)
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    TabControl, TabPage, DockStyle, Padding, AnchorStyles, ScrollBars,
)
from System.Drawing import Size

from pyrevit import forms, script

from ui_template import BaseForm, UIFactory, DPIScaler
from kozijnstaat.config import load_config, save_config, reset_config


def _refs_to_text(refs):
    """Lijst van strings naar een nieuwe-regel-getekste TextBox waarde."""
    if not refs:
        return ""
    return "\r\n".join(refs)


def _text_to_refs(text):
    """Scheid textbox-inhoud op regels/kommas, strip lege items."""
    if not text:
        return []
    out = []
    for chunk in text.replace(",", "\n").splitlines():
        s = chunk.strip()
        if s:
            out.append(s)
    return out


class KozijnstaatConfigForm(BaseForm):
    """Config dialoog met tabs voor alle instellingen."""

    def __init__(self, cfg):
        self.cfg = dict(cfg)
        super(KozijnstaatConfigForm, self).__init__(
            "Kozijnstaat - Config", 760, 640
        )
        self.set_subtitle("Instellingen per project")

        self.add_footer_button(
            "Opslaan", "primary", self._on_save, 140
        )
        self.add_footer_button(
            "Reset defaults", "secondary", self._on_reset, 140
        )

        self._build_ui()

    def _build_ui(self):
        self.tabs = TabControl()
        self.tabs.Dock = DockStyle.Fill
        self.pnl_content.Controls.Add(self.tabs)

        self._build_tab_families()
        self._build_tab_layout()
        self._build_tab_references()

    # ---- Tab 1: Families ----
    def _build_tab_families(self):
        tab = TabPage("Families & Canvas")
        tab.Padding = Padding(DPIScaler.scale(15))
        self.tabs.TabPages.Add(tab)

        y = 10
        lw = 210

        y = self._add_textbox(tab, y, lw, "Kozijn family-naam",
                              "kozijn_family", 300)
        y = self._add_textbox(tab, y, lw, "Glas tag family-naam",
                              "glas_tag_family", 300)
        y = self._add_textbox(tab, y, lw, "Naam-filter (bevat)",
                              "name_filter_contains", 200)
        y = self._add_textbox(tab, y, lw, "Canvas-wall Mark",
                              "canvas_wall_name", 300)
        y = self._add_textbox(tab, y, lw, "Canvas-wall type",
                              "canvas_wall_type", 300)
        y = self._add_textbox(tab, y, lw, "Canvas-wall level (leeg = eerste)",
                              "canvas_wall_level", 250)

        y = self._add_textbox(tab, y, lw, "Parameter 'aantal'",
                              "param_aantal", 200)
        y = self._add_textbox(tab, y, lw,
                              "Parameter 'aantal_gespiegeld'",
                              "param_aantal_gespiegeld", 200)

    # ---- Tab 2: Layout ----
    def _build_tab_layout(self):
        tab = TabPage("Grid & Offsets")
        tab.Padding = Padding(DPIScaler.scale(15))
        self.tabs.TabPages.Add(tab)

        y = 10
        lw = 210

        y = self._add_textbox(tab, y, lw, "Grid rijen",
                              "grid_rows", 80)
        y = self._add_textbox(tab, y, lw, "Grid kolommen",
                              "grid_cols", 80)
        y = self._add_textbox(tab, y, lw, "Tag-offset Z (mm)",
                              "tag_offset_mm", 100)
        y = self._add_textbox(tab, y, lw, "Glas-tag offset X (mm)",
                              "glas_tag_offset_x_mm", 100)
        y = self._add_textbox(tab, y, lw, "Glas-tag offset Y (mm)",
                              "glas_tag_offset_y_mm", 100)

    # ---- Tab 3: References ----
    def _build_tab_references(self):
        tab = TabPage("Maatvoering References")
        tab.Padding = Padding(DPIScaler.scale(15))
        self.tabs.TabPages.Add(tab)

        y = 10
        lw = 230

        y = self._add_multiline(
            tab, y, lw,
            "Detail horizontaal",
            "detail_h_refs", 380, 120,
        )
        y = self._add_multiline(
            tab, y, lw,
            "Detail verticaal",
            "detail_v_refs", 380, 80,
        )
        y = self._add_multiline(
            tab, y, lw,
            "Hoofdmaat horizontaal",
            "main_h_refs", 380, 50,
        )
        y = self._add_multiline(
            tab, y, lw,
            "Hoofdmaat verticaal",
            "main_v_refs", 380, 50,
        )

    # ---- Helpers ----
    def _add_textbox(self, tab, y, lw, label, key, width):
        lbl = UIFactory.create_label(label)
        lbl.Location = DPIScaler.scale_point(10, y + 4)
        tab.Controls.Add(lbl)
        tb = UIFactory.create_textbox(width)
        tb.Text = _safe_str(self.cfg.get(key, ""))
        tb.Location = DPIScaler.scale_point(lw, y)
        tb.Name = "txt_" + key
        tab.Controls.Add(tb)
        setattr(self, "txt_" + key, tb)
        return y + 38

    def _add_multiline(self, tab, y, lw, label, key, width, height):
        lbl = UIFactory.create_label(label)
        lbl.Location = DPIScaler.scale_point(10, y + 4)
        tab.Controls.Add(lbl)
        tb = UIFactory.create_textbox(width)
        tb.Multiline = True
        tb.ScrollBars = ScrollBars.Vertical
        tb.Size = DPIScaler.scale_size(width, height)
        tb.Text = _refs_to_text(self.cfg.get(key, []))
        tb.Location = DPIScaler.scale_point(lw, y)
        tb.Anchor = (AnchorStyles.Top | AnchorStyles.Left
                     | AnchorStyles.Right)
        tb.Name = "txt_" + key
        tab.Controls.Add(tb)
        setattr(self, "txt_" + key, tb)
        return y + height + 20

    # ---- Actions ----
    def _on_save(self, sender, args):
        try:
            new_cfg = self._collect()
            save_config(new_cfg)
            self.cfg = new_cfg
            forms.alert(
                "Opgeslagen.",
                title="Kozijnstaat Config",
            )
            self.Close()
        except Exception as ex:
            self.show_error("Fout: {0}".format(ex))

    def _on_reset(self, sender, args):
        if forms.alert(
            "Alle user-overrides verwijderen en defaults herstellen?",
            yes=True, no=True,
        ):
            reset_config()
            forms.alert("Defaults hersteld. Open Config opnieuw.",
                        title="Kozijnstaat Config")
            self.Close()

    def _collect(self):
        c = dict(self.cfg)

        # Strings
        for key in ("kozijn_family", "glas_tag_family",
                    "name_filter_contains", "canvas_wall_name",
                    "canvas_wall_type", "param_aantal",
                    "param_aantal_gespiegeld"):
            tb = getattr(self, "txt_" + key, None)
            if tb is not None:
                c[key] = tb.Text.strip()

        # Optional string (level - leeg = None)
        lvl_tb = getattr(self, "txt_canvas_wall_level", None)
        if lvl_tb is not None:
            v = lvl_tb.Text.strip()
            c["canvas_wall_level"] = v if v else None

        # Ints
        for key in ("grid_rows", "grid_cols"):
            tb = getattr(self, "txt_" + key, None)
            if tb is not None:
                c[key] = _safe_int(tb.Text, c.get(key, 0))

        # Floats
        for key in ("tag_offset_mm", "glas_tag_offset_x_mm",
                    "glas_tag_offset_y_mm"):
            tb = getattr(self, "txt_" + key, None)
            if tb is not None:
                c[key] = _safe_float(tb.Text, c.get(key, 0.0))

        # Reference-lijsten
        for key in ("detail_h_refs", "detail_v_refs",
                    "main_h_refs", "main_v_refs"):
            tb = getattr(self, "txt_" + key, None)
            if tb is not None:
                c[key] = _text_to_refs(tb.Text)

        return c


def _safe_str(v):
    return "" if v is None else str(v)


def _safe_int(text, default):
    try:
        return int(text)
    except (ValueError, TypeError):
        return default


def _safe_float(text, default):
    try:
        return float(text.replace(",", "."))
    except (ValueError, TypeError, AttributeError):
        return default


def run():
    cfg = load_config()
    dlg = KozijnstaatConfigForm(cfg)
    dlg.ShowDialog()


if __name__ == "__main__":
    run()

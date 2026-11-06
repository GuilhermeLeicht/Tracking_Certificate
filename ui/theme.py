# -*- coding: utf-8 -*-
"""Tema visual azul moderno (folha de estilos QSS) do Tracking Certificate - PySide6."""

DARK_BLUE = "#0d3b66"
MEDIUM_BLUE = "#1f6fb2"
ACCENT_BLUE = "#3d8bd4"
LIGHT_BG = "#f2f6fb"
CARD_BG = "#ffffff"
TEXT_DARK = "#1c2733"
TEXT_MUTED = "#6b7785"
GREEN = "#2e8b57"
AMBER = "#e08e0b"
RED = "#c0392b"
BLUE = "#2f6fed"
PURPLE = "#8e5ad1"
GRAY_BORDER = "#dfe6ee"

FONT_FAMILY = "Segoe UI"

STYLESHEET = f"""
QWidget {{
    background-color: {LIGHT_BG};
    color: {TEXT_DARK};
    font-family: "{FONT_FAMILY}";
    font-size: 10.5pt;
}}

QMainWindow {{
    background-color: {LIGHT_BG};
}}

#HeaderBar {{
    background-color: {DARK_BLUE};
}}
#HeaderTitle {{
    color: white;
    font-size: 23px;
    font-weight: 700;
    background: transparent;
}}
#HeaderSubtitle {{
    color: #a9c6e8;
    font-size: 10.5pt;
    background: transparent;
}}

QTabWidget::pane {{
    border: none;
    background: transparent;
    top: 2px;
}}
QTabBar::tab {{
    background: #dbe7f5;
    color: {DARK_BLUE};
    padding: 10px 22px;
    margin-right: 4px;
    border-top-left-radius: 9px;
    border-top-right-radius: 9px;
    font-weight: 600;
}}
QTabBar::tab:selected {{
    background: {DARK_BLUE};
    color: white;
}}
QTabBar::tab:hover:!selected {{
    background: #c7dbf0;
}}

QTabWidget#SubTabs::pane {{
    border: 1px solid {GRAY_BORDER};
    border-radius: 12px;
    background: {CARD_BG};
    top: -1px;
}}
QTabWidget#SubTabs QTabBar::tab {{
    background: #eaf1fb;
    padding: 8px 18px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}}
QTabWidget#SubTabs QTabBar::tab:selected {{
    background: {MEDIUM_BLUE};
    color: white;
}}

#Sidebar {{
    background-color: {DARK_BLUE};
}}
QPushButton#SidebarButton {{
    background: transparent;
    color: #cfe0f2;
    text-align: left;
    padding: 12px 20px;
    border-radius: 0px;
    border-left: 4px solid transparent;
    font-weight: 600;
}}
QPushButton#SidebarButton:hover {{
    background: #14497d;
    color: white;
}}
QPushButton#SidebarButton:checked {{
    background: {MEDIUM_BLUE};
    color: white;
    border-left: 4px solid white;
}}
QPushButton#SidebarToggleOld {{
    background: transparent;
    color: white;
    font-size: 15pt;
    padding: 6px 14px;
    border-radius: 6px;
    text-align: left;
}}
QPushButton#SidebarToggle:hover {{
    background: #14497d;
}}
QPushButton#SidebarButton[collapsed="true"] {{
    padding: 12px 0px;
    text-align: center;
    border-left: 0px;
}}
QLabel#TimeSavedLabel {{
    color: white;
    background-color: {GREEN};
    border-radius: 9px;
    padding: 7px 16px;
    font-weight: 800;
    font-size: 11pt;
}}

QPushButton {{
    background-color: {MEDIUM_BLUE};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 20px;
    font-weight: 600;
}}
QPushButton:hover {{
    background-color: {DARK_BLUE};
}}
QPushButton:pressed {{
    background-color: #082746;
}}
QPushButton:disabled {{
    background-color: #b9cbde;
    color: #eef2f7;
}}
QPushButton#SecondaryButton {{
    background-color: white;
    color: {DARK_BLUE};
    border: 1px solid {GRAY_BORDER};
}}
QPushButton#SecondaryButton:hover {{
    background-color: #eaf1fb;
}}
QPushButton#SmallButton {{
    padding: 7px 16px;
    font-size: 9.5pt;
    border-radius: 7px;
}}
QPushButton#BrowseButton {{
    padding: 6px 10px;
    border-radius: 6px;
    font-weight: 700;
}}

QLineEdit, QComboBox, QSpinBox {{
    background: white;
    border: 1px solid {GRAY_BORDER};
    border-radius: 7px;
    padding: 7px 9px;
    selection-background-color: {ACCENT_BLUE};
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
    border: 1px solid {ACCENT_BLUE};
}}

QTableWidget {{
    background: white;
    border: 1px solid {GRAY_BORDER};
    border-radius: 0px;
    gridline-color: #eef1f5;
    alternate-background-color: #f8fafc;
}}
QHeaderView::section {{
    background-color: {DARK_BLUE};
    color: white;
    padding: 9px 8px;
    border: none;
    font-weight: 600;
}}
QHeaderView::section:hover {{
    background-color: {MEDIUM_BLUE};
}}
QTableWidget::item {{
    padding: 5px 6px;
}}
QTableWidget::item:selected {{
    background-color: {ACCENT_BLUE};
    color: white;
}}
QTableCornerButton::section {{
    background-color: {DARK_BLUE};
    border: none;
}}

QProgressBar {{
    border: none;
    border-radius: 6px;
    background-color: #dbe7f5;
    max-height: 10px;
    min-height: 10px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    background-color: {MEDIUM_BLUE};
    border-radius: 6px;
}}

QTextEdit {{
    background: #10233a;
    color: #cfe3f7;
    border: 1px solid {GRAY_BORDER};
    border-radius: 10px;
    padding: 10px;
    font-family: Consolas, "Courier New", monospace;
    font-size: 9.5pt;
}}

QGroupBox {{
    background: white;
    border: 1px solid {GRAY_BORDER};
    border-radius: 12px;
    margin-top: 16px;
    font-weight: 700;
    padding-top: 14px;
    color: {DARK_BLUE};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 8px;
}}

QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollArea > QWidget > QWidget {{
    background: transparent;
}}

QLabel#StatValue {{
    font-size: 28px;
    font-weight: 800;
    background: transparent;
}}
QLabel#StatTitle {{
    color: {TEXT_MUTED};
    font-size: 9pt;
    font-weight: 700;
    background: transparent;
}}
QLabel#PendingCounter {{
    color: {TEXT_MUTED};
    font-size: 10pt;
}}
QLabel#FileLabel {{
    color: {TEXT_MUTED};
    font-size: 9.5pt;
}}
QLabel#FieldLabel {{
    color: {TEXT_DARK};
    font-weight: 600;
    font-size: 9.5pt;
}}
QLabel#SectionHint {{
    color: {TEXT_MUTED};
    font-size: 9pt;
}}
QFrame#StatCard {{
    background: white;
    border: 1px solid {GRAY_BORDER};
    border-radius: 12px;
}}

QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 0px;
}}
QScrollBar::handle:vertical {{
    background: #c3d3e6;
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {ACCENT_BLUE};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
"""

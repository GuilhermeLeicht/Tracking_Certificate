# -*- coding: utf-8 -*-
"""Janela principal do Tracking Certificate - PySide6, tema azul moderno."""
import os
import re
import threading
from datetime import date, datetime, timedelta

from PySide6.QtCore import Qt, QObject, Signal, QPointF, QRectF, QSize
from PySide6.QtGui import QColor, QBrush, QIcon, QPixmap, QPainter, QPen
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QPushButton, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressBar, QLineEdit, QSpinBox, QGroupBox, QScrollArea, QTextEdit,
    QFileDialog, QMessageBox, QGraphicsDropShadowEffect, QSizePolicy,
    QDialog, QComboBox, QTextBrowser, QStackedWidget, QButtonGroup,
)

import config as cfgmod
import contacts_db
import excel_reader
import reminder_engine
import templates
import notifications
import supplier_rows
from ui.common import SortItem, numeric_or_text_key
from ui.pages import RiskAssessmentPage, ReportsPage, ReceivedCertDialog
from received_db import ReceivedDB
from risk_db import RiskDB
from history_db import HistoryDB
from reminder_engine import EXPIRED, NO_DATE
from supplier_rows import (
    CERT_TYPES, CERT_TYPE_LABELS_PT, CERT_ISO_TAG, COMBINED_FIELD, CERT_LABEL_EN,
)
from ui import theme

try:
    import email_sender
    OUTLOOK_AVAILABLE = True
except Exception:
    OUTLOOK_AVAILABLE = False


LOG_FILE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tracking_app_log.txt")

TABLE_HEADERS = ["Supplier", "RBSNO", "PMD", "Expiry Date", "Status", "Reminder", "E-mail"]
CENTERED_COLUMNS = (1, 2, 3, 4, 5)  # RBSNO, PMD, Expiry Date, Status, Reminder

ROW_COLORS = {
    "active": QColor("#d9f2df"),
    "expiring": QColor("#fff3cd"),
    "expired": QColor("#ffb8b8"),
    "no_date": QColor("#e8eaed"),
    "blue": QColor("#cfe0ff"),
    "risk": QColor("#e3d4f7"),
}


# --------------------------------------------------------------------------- funções auxiliares

_numeric_or_text_key = numeric_or_text_key


def make_nav_icon(kind, color="#ffffff", size=40):
    """Desenha um ícone simples (traço branco) para o menu lateral.
    kind: panel | history | log | settings. Tudo em código, sem arquivos."""
    import math
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(color), size * 0.075, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)
    u = size / 24.0
    if kind == "panel":  # grade 2x2 (dashboard)
        for x, y in ((3, 3), (13, 3), (3, 13), (13, 13)):
            p.drawRoundedRect(QRectF(x * u, y * u, 8 * u, 8 * u), 1.5 * u, 1.5 * u)
    elif kind == "history":  # relógio
        p.drawEllipse(QPointF(12 * u, 12 * u), 9 * u, 9 * u)
        p.drawLine(QPointF(12 * u, 6.5 * u), QPointF(12 * u, 12 * u))
        p.drawLine(QPointF(12 * u, 12 * u), QPointF(16 * u, 14.5 * u))
    elif kind == "log":  # documento com linhas
        p.drawRoundedRect(QRectF(5 * u, 2.5 * u, 14 * u, 19 * u), 2 * u, 2 * u)
        for y in (8, 12, 16):
            p.drawLine(QPointF(8.5 * u, y * u), QPointF(15.5 * u, y * u))
    elif kind == "menu":  # hambúrguer
        for y in (6, 12, 18):
            p.drawLine(QPointF(4 * u, y * u), QPointF(20 * u, y * u))
    elif kind == "risk":  # triângulo de alerta
        p.drawPolygon([QPointF(12 * u, 3 * u), QPointF(22 * u, 20 * u), QPointF(2 * u, 20 * u)])
        p.drawLine(QPointF(12 * u, 9.5 * u), QPointF(12 * u, 14 * u))
        p.drawPoint(QPointF(12 * u, 17 * u))
    elif kind == "reports":  # gráfico de barras
        p.drawLine(QPointF(3 * u, 21 * u), QPointF(21 * u, 21 * u))
        for x, top in ((5, 12), (10.5, 5), (16, 9)):
            p.drawRoundedRect(QRectF(x * u, top * u, 3.5 * u, (20 - top) * u), 0.8 * u, 0.8 * u)
    elif kind == "settings":  # engrenagem
        c = QPointF(12 * u, 12 * u)
        p.drawEllipse(c, 3.2 * u, 3.2 * u)
        p.drawEllipse(c, 6.6 * u, 6.6 * u)
        for i in range(8):
            a = math.radians(i * 45)
            p.drawLine(QPointF(12 * u + 6.6 * u * math.cos(a), 12 * u + 6.6 * u * math.sin(a)),
                       QPointF(12 * u + 9.6 * u * math.cos(a), 12 * u + 9.6 * u * math.sin(a)))
    p.end()
    return QIcon(pm)


class StatCard(QFrame):
    """Cartão pequeno e arredondado, com uma faixa colorida lateral, um
    título e um valor grande."""

    def __init__(self, title, accent_color, parent=None):
        super().__init__(parent)
        self.setObjectName("StatCard")
        self.setMinimumHeight(86)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        strip = QFrame()
        strip.setFixedWidth(6)
        strip.setStyleSheet(
            f"background-color: {accent_color}; "
            "border-top-left-radius: 12px; border-bottom-left-radius: 12px;"
        )
        outer.addWidget(strip)

        inner = QVBoxLayout()
        inner.setContentsMargins(18, 14, 18, 14)
        inner.setSpacing(6)
        self.title_label = QLabel(title.upper())
        self.title_label.setObjectName("StatTitle")
        self.value_label = QLabel("0")
        self.value_label.setObjectName("StatValue")
        inner.addWidget(self.title_label)
        inner.addWidget(self.value_label)
        outer.addLayout(inner)
        outer.addStretch()

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(18)
        shadow.setXOffset(0)
        shadow.setYOffset(3)
        shadow.setColor(QColor(13, 59, 102, 40))
        self.setGraphicsEffect(shadow)

    def set_value(self, value):
        self.value_label.setText(str(value))


class Worker(QObject):
    """Central de sinais usada para levar os resultados calculados em
    segunda thread até a thread da interface gráfica (o Qt enfileira
    automaticamente sinais emitidos entre threads diferentes)."""
    analysis_finished = Signal(list, list)
    analysis_error = Signal(str)
    log_message = Signal(str)
    send_progress = Signal(int)
    send_finished = Signal(int, int)


class EmailPreviewDialog(QDialog):
    """Pop-up que mostra como o e-mail fica, combinando tipo de certificado
    e prazo (90/60/30 dias ou vencido). Usa os valores dos campos de
    Settings em tempo real (mesmo que ainda não tenham sido salvos), para o
    usuário poder ajustar o texto e já ver o resultado. O app só envia
    e-mails para fornecedores nacionais, então a prévia é sempre em
    português."""

    TIER_OPTIONS = [
        ("90 days before expiry", "90", 90),
        ("60 days before expiry", "60", 60),
        ("30 days before expiry", "30", 30),
        ("Expired", EXPIRED, -12),
        ("Missing certificate (ISO 9001 and IATF 16949 blank)", NO_DATE, 0),
    ]

    def __init__(self, parent, settings_fields):
        super().__init__(parent)
        self.setWindowTitle("Preview E-mail Content")
        self.resize(680, 660)
        self.settings_fields = settings_fields

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)

        controls = QHBoxLayout()
        controls.setSpacing(10)

        self.combo_cert = QComboBox()
        self.combo_cert.addItems([label for _field, label in CERT_TYPES]
                                 + [CERT_LABEL_EN[COMBINED_FIELD] + " (neither valid)"])
        self.combo_tier = QComboBox()
        self.combo_tier.addItems([label for label, _tier, _days in self.TIER_OPTIONS])

        for combo in (self.combo_cert, self.combo_tier):
            combo.currentIndexChanged.connect(self._update_preview)
            controls.addWidget(combo, stretch=1)
        layout.addLayout(controls)

        self.subject_label = QLabel()
        self.subject_label.setObjectName("FieldLabel")
        self.subject_label.setWordWrap(True)
        layout.addWidget(self.subject_label)

        self.browser = QTextBrowser()
        layout.addWidget(self.browser, stretch=1)

        close_row = QHBoxLayout()
        close_row.addStretch()
        btn_close = QPushButton("Close")
        btn_close.setObjectName("SecondaryButton")
        btn_close.clicked.connect(self.accept)
        close_row.addWidget(btn_close)
        layout.addLayout(close_row)

        self._update_preview()

    def _update_preview(self):
        idx = self.combo_cert.currentIndex()
        field = CERT_TYPES[idx][0] if idx < len(CERT_TYPES) else COMBINED_FIELD
        _label, tier, days_left = self.TIER_OPTIONS[self.combo_tier.currentIndex()]

        cert_label = CERT_TYPE_LABELS_PT[field]
        sample_supplier = "Fornecedor Exemplo Ltda"

        company_name = self.settings_fields["company_name"].text().strip() or "Your Company"
        subject_prefix = self.settings_fields["default_subject_pt"].text().strip()
        if tier == NO_DATE:
            subject_prefix = self.settings_fields["default_subject_missing_pt"].text().strip()
        iso_tag = CERT_ISO_TAG[field]

        expiry_str = (date.today() + timedelta(days=days_left)).strftime("%Y-%m-%d")

        body = templates.build_email_body(
            sample_supplier, cert_label, expiry_str, days_left, tier, company_name,
            language="pt-br")

        subject = f"{subject_prefix} - {iso_tag} - {sample_supplier}"
        self.subject_label.setText(f"Subject: {subject}")
        self.browser.setHtml(body)


# --------------------------------------------------------------------------- janela principal

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tracking Certificate - Supplier Certificate Renewal")
        self.resize(1220, 760)
        self.setMinimumSize(1000, 620)
        self.setStyleSheet(theme.STYLESHEET)

        self.cfg = cfgmod.load_config()
        self.suppliers = []
        self.status_rows = []
        self.processing = False
        self.settings_fields = {}
        # totais por tipo de certificado, usados para os cartões dinâmicos
        # que mudam conforme a sub-aba (ISO 9001 / Ambiental) selecionada
        self.totals_by_field = {}

        self.worker = Worker()
        self.worker.analysis_finished.connect(self._on_analysis_finished)
        self.worker.analysis_error.connect(self._on_analysis_error)
        self.worker.log_message.connect(self._log)
        self.worker.send_progress.connect(self._on_send_progress)
        self.worker.send_finished.connect(self._on_send_finished)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_header())

        self.panel_tab = self._build_panel_tab()
        self.history_tab = self._build_history_tab()
        self.log_tab = self._build_log_tab()
        self.settings_tab = self._build_settings_tab()

        self.risk_tab = RiskAssessmentPage(
            lambda: self.cfg["history_db_path"], lambda: self.suppliers, self._log,
            folder_getter=lambda: self.settings_fields["certificates_folder"].text().strip()
            if "certificates_folder" in self.settings_fields else self.cfg.get("certificates_folder", ""),
            on_change=self._refresh_follow)
        self.reports_tab = ReportsPage()

        self.stack = QStackedWidget()
        self.nav_pages = ["panel", "risk", "reports", "history", "log", "settings"]
        for page in (self.panel_tab, self.risk_tab, self.reports_tab,
                     self.history_tab, self.log_tab, self.settings_tab):
            self.stack.addWidget(page)

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        body.addWidget(self._build_sidebar())
        body.addWidget(self.stack, stretch=1)
        root.addLayout(body, stretch=1)
        self._navigate("panel")

        # conecta a troca de sub-aba (ISO 9001 / Ambiental) à atualização
        # dinâmica dos 4 cartões, feita depois de todas as abas montadas
        self.sub_tabs.currentChanged.connect(self._on_subtab_changed)

        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready.")

        self._restore_last_excel()
        self._load_previous_log()
        self._log("Application started.")
        if not OUTLOOK_AVAILABLE:
            self._log("Warning: Outlook automation (pywin32) is not available in this "
                       "environment. Sending only works on Windows with Outlook installed.")

    # ------------------------------------------------------------------ barra lateral

    SIDEBAR_EXPANDED = 215
    SIDEBAR_COLLAPSED = 64

    def _build_sidebar(self):
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(self.SIDEBAR_EXPANDED)
        self._sidebar_collapsed = False
        lay = QVBoxLayout(self.sidebar)
        lay.setContentsMargins(0, 10, 0, 14)
        lay.setSpacing(2)

        self.btn_toggle_sidebar = QPushButton("  Menu")
        self.btn_toggle_sidebar.setObjectName("SidebarButton")
        self.btn_toggle_sidebar.setIcon(make_nav_icon("menu"))
        self.btn_toggle_sidebar.setIconSize(QSize(22, 22))
        self.btn_toggle_sidebar.setToolTip("Collapse / expand menu")
        self.btn_toggle_sidebar.clicked.connect(self._toggle_sidebar)
        lay.addWidget(self.btn_toggle_sidebar)
        lay.addSpacing(10)

        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons = {}
        self.nav_labels = {"panel": "Panel", "risk": "Risk Assessment", "reports": "Reports",
                           "history": "Send History", "log": "Log", "settings": "Settings"}
        for key in ("panel", "risk", "reports"):
            lay.addWidget(self._make_nav_button(key))
        lay.addStretch()  # de baixo para cima: Settings, Log, Send History
        for key in ("history", "log", "settings"):
            lay.addWidget(self._make_nav_button(key))
        return self.sidebar

    def _make_nav_button(self, key):
        btn = QPushButton("  " + self.nav_labels[key])
        btn.setObjectName("SidebarButton")
        btn.setCheckable(True)
        btn.setIcon(make_nav_icon(key))
        btn.setIconSize(QSize(22, 22))
        btn.setToolTip(self.nav_labels[key])
        btn.clicked.connect(lambda _checked=False, k=key: self._navigate(k))
        self.nav_group.addButton(btn)
        self.nav_buttons[key] = btn
        return btn

    def _navigate(self, key):
        self.stack.setCurrentIndex(self.nav_pages.index(key))
        self.nav_buttons[key].setChecked(True)

    def _toggle_sidebar(self):
        self._sidebar_collapsed = not self._sidebar_collapsed
        self.sidebar.setFixedWidth(
            self.SIDEBAR_COLLAPSED if self._sidebar_collapsed else self.SIDEBAR_EXPANDED)
        buttons = [(self.btn_toggle_sidebar, "Menu")] + [
            (b, self.nav_labels[k]) for k, b in self.nav_buttons.items()]
        for btn, full in buttons:
            btn.setText("" if self._sidebar_collapsed else "  " + full)
            # propriedade usada pelo QSS para centralizar o ícone quando recolhido
            btn.setProperty("collapsed", self._sidebar_collapsed)
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    # ------------------------------------------------------------------ cabeçalho

    def _build_header(self):
        header = QFrame()
        header.setObjectName("HeaderBar")
        layout = QVBoxLayout(header)
        layout.setContentsMargins(28, 18, 28, 18)
        layout.setSpacing(2)

        title = QLabel("Tracking Certificate")
        title.setObjectName("HeaderTitle")
        subtitle = QLabel("Automatic certificate renewal monitoring and alerts (ISO 9001 / ISO 14001)")
        subtitle.setObjectName("HeaderSubtitle")

        layout.addWidget(title)
        layout.addWidget(subtitle)
        return header

    # ------------------------------------------------------------------ aba Panel

    def _build_panel_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)
        self.btn_select_excel = QPushButton("1. Select SRM Excel File")
        self.btn_select_excel.clicked.connect(self._select_excel)
        self.btn_analyze = QPushButton("2. Load && Analyze")
        self.btn_analyze.clicked.connect(self._load_and_analyze)
        toolbar.addWidget(self.btn_select_excel)
        toolbar.addWidget(self.btn_analyze)
        toolbar.addStretch()
        self.btn_received = QPushButton("Register Received Certificate")
        self.btn_received.setObjectName("SecondaryButton")
        self.btn_received.setToolTip("Supplier already sent the file (not uploaded to the datalake yet): pause reminders")
        self.btn_received.clicked.connect(self._register_received)
        toolbar.addWidget(self.btn_received)
        layout.addLayout(toolbar)

        self.file_label = QLabel("No Excel file selected.")
        self.file_label.setObjectName("FileLabel")
        layout.addWidget(self.file_label)

        # os 4 cartões refletem apenas o tipo de certificado da sub-aba
        # atualmente selecionada (ver _on_subtab_changed)
        stats_row = QHBoxLayout()
        stats_row.setSpacing(14)
        self.card_suppliers = StatCard("Suppliers", theme.MEDIUM_BLUE)
        self.card_active = StatCard("Active", theme.GREEN)
        self.card_expiring = StatCard("Expiring", theme.AMBER)
        self.card_expired = StatCard("Expired", theme.RED)
        for card in (self.card_suppliers, self.card_active, self.card_expiring, self.card_expired):
            stats_row.addWidget(card)
        layout.addLayout(stats_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.sub_tabs = QTabWidget()
        self.sub_tabs.setObjectName("SubTabs")
        self.tables = {}
        for field, label in CERT_TYPES:
            table = self._build_cert_table()
            self.sub_tabs.addTab(table, label)
            self.tables[field] = table
        layout.addWidget(self.sub_tabs, stretch=1)

        bottom_bar = QHBoxLayout()
        bottom_bar.addStretch()
        self.pending_label = QLabel("0 pending")
        self.pending_label.setObjectName("PendingCounter")
        self.btn_send = QPushButton("Send Pending Reminders")
        self.btn_send.setObjectName("SmallButton")
        self.btn_send.clicked.connect(self._send_pending)
        bottom_bar.addWidget(self.pending_label)
        bottom_bar.addSpacing(10)
        bottom_bar.addWidget(self.btn_send)
        layout.addLayout(bottom_bar)

        return widget

    def _build_cert_table(self):
        table = QTableWidget(0, len(TABLE_HEADERS))
        table.setHorizontalHeaderLabels(TABLE_HEADERS)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setSortingEnabled(True)
        table.verticalHeader().setVisible(False)
        table.setShowGrid(True)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)          # Supplier
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # RBSNO
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # PMD
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Expiry Date
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Status
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Reminder
        header.setSectionResizeMode(6, QHeaderView.Stretch)          # E-mail
        return table

    # ------------------------------------------------------------------ aba Send History

    def _build_history_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)

        top = QHBoxLayout()
        btn_refresh = QPushButton("Refresh")
        btn_refresh.setObjectName("SecondaryButton")
        btn_refresh.clicked.connect(self._refresh_history)
        top.addWidget(btn_refresh)
        top.addStretch()
        self.history_total_label = QLabel("0 total sends")
        self.history_total_label.setObjectName("PendingCounter")
        top.addWidget(self.history_total_label)
        self.time_saved_label = QLabel("Time Saved: 0.00 hours")
        self.time_saved_label.setObjectName("TimeSavedLabel")
        top.addSpacing(14)
        top.addWidget(self.time_saved_label)
        layout.addLayout(top)

        headers = ["RBSNO", "PMD", "Supplier", "Certificate", "Expiry Date", "Tier", "Sent Date", "E-mail"]
        self.history_table = QTableWidget(0, len(headers))
        self.history_table.setHorizontalHeaderLabels(headers)
        self.history_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.setSortingEnabled(True)
        self.history_table.verticalHeader().setVisible(False)
        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.history_table, stretch=1)

        return widget

    def _refresh_history(self):
        self.history_table.setSortingEnabled(False)
        self.history_table.setRowCount(0)
        try:
            hdb = HistoryDB(self.cfg["history_db_path"])
            for row in hdb.history():
                rbsno, pmd, name, cert_type, expiry, tier, sent_date, recipient = row
                cert_label = CERT_LABEL_EN.get(cert_type, cert_type)
                r = self.history_table.rowCount()
                self.history_table.insertRow(r)
                values = [rbsno, pmd, name, cert_label, expiry,
                          reminder_engine.tier_label(tier), sent_date, recipient]
                for c, val in enumerate(values):
                    self.history_table.setItem(r, c, SortItem(str(val or "")))
            total_sends = hdb.count()
            self.history_total_label.setText(f"{total_sends} total sends")
            hours = (total_sends * 5) / 60
            self.time_saved_label.setText(f"Time Saved: {hours:.2f} hours")
            hdb.close()
        except Exception as exc:
            self._log(f"ERROR reading the send history: {exc}")
            QMessageBox.critical(self, "History", f"Could not read the send history:\n{exc}")
        self.history_table.setSortingEnabled(True)

    # ------------------------------------------------------------------ aba Settings

    def _build_settings_tab(self):
        outer = QWidget()
        outer_layout = QVBoxLayout(outer)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        outer_layout.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(16)

        # --- Colunas do Excel do SRM ---------------------------------------------------
        group_excel = QGroupBox("SRM Excel Columns")
        v = QVBoxLayout(group_excel)
        v.setSpacing(10)
        hint = QLabel("Copy and paste the exact column header text from your Excel file below.")
        hint.setObjectName("SectionHint")
        hint.setWordWrap(True)
        v.addWidget(hint)

        self._add_field(v, "excel_sheet_name", "Sheet name with the data (leave blank to auto-detect):")
        self._add_field(v, "excel_col_rbsno", "Column: Supplier code (RBSNO):")
        self._add_field(v, "excel_col_pmd", "Column: Supplier code (PMD/SAP):")
        self._add_field(v, "excel_col_name", "Column: Supplier name:")
        self._add_field(v, "excel_col_country", "Column: Supplier country code (only \"BR\" suppliers are processed):")
        self._add_field(v, "excel_col_iso", "Column: ISO 9001 certificate expiry date:")
        self._add_field(v, "excel_col_iatf", "Column: IATF 16949 certificate expiry date:")
        self._add_field(v, "excel_col_env", "Column: Environmental certificate expiry date:")
        layout.addWidget(group_excel)

        # --- Diretório de e-mails no Access ----------------------------------------------
        group_access = QGroupBox("Supplier E-mail Directory (Access)")
        v2 = QVBoxLayout(group_access)
        v2.setSpacing(10)
        self._add_file_field(v2, "access_db_path", "Access database (.accdb):",
                              "Access Database (*.accdb *.mdb)")
        self._add_field(v2, "access_table", "Table:")
        self._add_field(v2, "access_col_rbsno", "RBSNO column:")
        self._add_field(v2, "access_col_pmd", "PMD column:")
        self._add_field(v2, "access_col_email", "E-mail column:")
        layout.addWidget(group_access)

        # --- Arquivos de certificados -------------------------------------------------
        group_files = QGroupBox("Certificate Files")
        v2b = QVBoxLayout(group_files)
        v2b.setSpacing(10)
        self._add_file_field(v2b, "certificates_folder",
                             "Folder for certificate files (a sub-folder \"Supplier - RBSNO\" is created automatically):",
                             None)
        layout.addWidget(group_files)

        # --- Regras de notificação ----------------------------------------------------
        group_rules = QGroupBox("Notification Rules")
        v3 = QVBoxLayout(group_rules)
        v3.setSpacing(10)

        row = QHBoxLayout()
        label = QLabel("Alert at (days before expiry):")
        label.setObjectName("FieldLabel")
        row.addWidget(label)
        days = sorted(self.cfg["reminder_days"], reverse=True)
        self.spin_days = []
        for d in days:
            spin = QSpinBox()
            spin.setRange(1, 365)
            spin.setValue(d)
            spin.setFixedWidth(80)
            row.addWidget(spin)
            self.spin_days.append(spin)
        row.addStretch()
        v3.addLayout(row)

        row2 = QHBoxLayout()
        label2 = QLabel("Resend expired-certificate notice every (days):")
        label2.setObjectName("FieldLabel")
        self.spin_overdue = QSpinBox()
        self.spin_overdue.setRange(1, 365)
        self.spin_overdue.setValue(self.cfg["overdue_resend_interval_days"])
        self.spin_overdue.setFixedWidth(80)
        row2.addWidget(label2)
        row2.addWidget(self.spin_overdue)
        row2.addStretch()
        v3.addLayout(row2)

        layout.addWidget(group_rules)

        # --- Conteúdo do e-mail ----------------------------------------------------------
        group_email = QGroupBox("E-mail Content")
        v4 = QVBoxLayout(group_email)
        v4.setSpacing(10)
        hint_lang = QLabel(
            "Only suppliers with Country Code \"BR\" in the SRM Excel are processed by this "
            "app, so reminder e-mails are always sent in Portuguese.")
        hint_lang.setObjectName("SectionHint")
        hint_lang.setWordWrap(True)
        v4.addWidget(hint_lang)
        self._add_field(v4, "company_name", "Company name (shown in the e-mail):")
        self._add_field(v4, "bcc_email", "BCC address (blind copy - the supplier does not see it; optional):")
        self._add_field(v4, "default_subject_pt", "Default subject (renewal / expired):")
        self._add_field(v4, "default_subject_missing_pt",
                        "Subject when ISO 9001 and IATF 16949 are blank:")
        self._add_field(v4, "sender_email",
                        "Sender address (send on behalf of, e.g. a shared mailbox; blank = your own Outlook account):")

        preview_row = QHBoxLayout()
        btn_preview = QPushButton("Preview E-mail Content...")
        btn_preview.setObjectName("SecondaryButton")
        btn_preview.clicked.connect(self._preview_email)
        preview_row.addWidget(btn_preview)
        preview_row.addStretch()
        v4.addLayout(preview_row)

        layout.addWidget(group_email)

        if not OUTLOOK_AVAILABLE:
            warn = QLabel("Note: Outlook automation (pywin32) is not available in this "
                           "environment. Sending only works on Windows with Outlook installed.")
            warn.setWordWrap(True)
            warn.setStyleSheet(f"color: {theme.RED};")
            layout.addWidget(warn)

        btn_save = QPushButton("Save Settings")
        btn_save.clicked.connect(self._save_settings)
        save_row = QHBoxLayout()
        save_row.addWidget(btn_save)
        save_row.addStretch()
        layout.addLayout(save_row)

        layout.addStretch()
        return outer

    def _add_field(self, parent_layout, key, label_text):
        container = QVBoxLayout()
        container.setSpacing(4)
        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        label.setWordWrap(True)
        edit = QLineEdit(str(self.cfg.get(key, "")))
        container.addWidget(label)
        container.addWidget(edit)
        parent_layout.addLayout(container)
        self.settings_fields[key] = edit

    def _add_file_field(self, parent_layout, key, label_text, file_filter):
        container = QVBoxLayout()
        container.setSpacing(4)
        label = QLabel(label_text)
        label.setObjectName("FieldLabel")
        label.setWordWrap(True)
        row = QHBoxLayout()
        edit = QLineEdit(str(self.cfg.get(key, "")))
        btn = QPushButton("...")
        btn.setObjectName("BrowseButton")
        btn.setFixedWidth(36)
        btn.clicked.connect(lambda: self._browse_file(edit, file_filter))
        row.addWidget(edit)
        row.addWidget(btn)
        container.addWidget(label)
        container.addLayout(row)
        parent_layout.addLayout(container)
        self.settings_fields[key] = edit

    def _browse_file(self, line_edit, file_filter):
        if file_filter is None:  # modo pasta
            path = QFileDialog.getExistingDirectory(self, "Select folder")
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Select file", "", file_filter)
        if path:
            line_edit.setText(path)

    def _refresh_follow(self):
        """Gráfico 'Follow-up': Risk Assessment x arquivos recebidos x sem data."""
        try:
            rdb = RiskDB(self.cfg["history_db_path"])
            risk = len({r[1] or r[3] for r in rdb.all()})
            rdb.close()
            fdb = ReceivedDB(self.cfg["history_db_path"])
            files = fdb.supplier_count()
            fdb.close()
        except Exception as exc:
            self._log(f"ERROR reading follow-up counts: {exc}")
            return
        # "Missing date" só conta quem ainda está sem tratamento (fora do
        # Risk Assessment e sem arquivo recebido)
        missing_total = sum(1 for r in self.status_rows
                            if r["tab"] == "iso_expiry" and r["field"] == COMBINED_FIELD and r["tier"] == NO_DATE)
        missing = max(0, missing_total - risk - files)
        self.reports_tab.set_follow({"risk": risk, "files": files, "missing": missing})

    def _register_received(self):
        if not self.suppliers:
            QMessageBox.information(self, "Register Received Certificate",
                                    "Load and analyze the SRM Excel file first.")
            return
        folder = self.settings_fields["certificates_folder"].text().strip()
        dlg = ReceivedCertDialog(self, self.suppliers, folder, self.cfg["history_db_path"], self._log)
        if dlg.exec():
            self._log("Certificate registered. Re-analyzing to pause reminders for this supplier...")
            self._load_and_analyze()

    def _preview_email(self):
        dialog = EmailPreviewDialog(self, self.settings_fields)
        dialog.exec()

    def _save_settings(self):
        try:
            days = sorted([spin.value() for spin in self.spin_days], reverse=True)
        except Exception:
            QMessageBox.critical(self, "Settings", "Invalid values in the day fields.")
            return

        for key, edit in self.settings_fields.items():
            self.cfg[key] = edit.text().strip()

        self.cfg["reminder_days"] = days
        self.cfg["overdue_resend_interval_days"] = self.spin_overdue.value()

        cfgmod.save_config(self.cfg)
        self._log("Settings saved.")
        QMessageBox.information(self, "Settings", "Settings saved successfully.")

    # ------------------------------------------------------------------ aba Log

    def _build_log_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(22, 20, 22, 20)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        return widget

    def _log(self, message):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{stamp}] {message}"
        self.log_text.append(line)
        try:  # grava em arquivo também, para o log sobreviver ao fechar o app
            with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    def _load_previous_log(self):
        try:
            with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()[-500:]
            if lines:
                self.log_text.append("\n".join(lines))
                self.log_text.append("---------- previous sessions above ----------")
        except Exception:
            pass

    # ------------------------------------------------------------------ ações

    def _restore_last_excel(self):
        """Reabre já com o último arquivo Excel escolhido (salvo na config)."""
        path = self.cfg.get("excel_path", "")
        if not path:
            return
        if os.path.exists(path):
            self.file_label.setText(f"Selected file: {path}")
        else:
            self.file_label.setText(f"Last file not found: {path}")

    def _select_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select the SRM Excel export", "", "Excel Files (*.xlsx)")
        if path:
            self.cfg["excel_path"] = path
            cfgmod.save_config(self.cfg)
            self.file_label.setText(f"Selected file: {path}")
            self._log(f"Excel file selected: {path}")

    def _load_and_analyze(self):
        if self.processing:
            return
        excel_path = self.cfg.get("excel_path", "")
        if not excel_path or not os.path.exists(excel_path):
            self._log("Analysis not started: no SRM Excel file selected (or file not found).")
            QMessageBox.warning(self, "Notice", "Please select the SRM Excel export file first.")
            return
        self._log(f"Starting analysis of: {excel_path}")

        self.processing = True
        self.status_bar.showMessage("Reading the Excel file and computing certificate status...")
        self.progress.setRange(0, 0)  # modo indeterminado / ocupado
        threading.Thread(target=self._analyze_worker, args=(excel_path,), daemon=True).start()

    def _analyze_worker(self, excel_path):
        try:
            reminder_engine.WARNING_THRESHOLDS = tuple(self.cfg["reminder_days"])
            excel_columns = {
                "rbsno": self.cfg["excel_col_rbsno"],
                "pmd": self.cfg["excel_col_pmd"],
                "name": self.cfg["excel_col_name"],
                "iso_expiry": self.cfg["excel_col_iso"],
                "env_expiry": self.cfg["excel_col_env"],
                "iatf_expiry": self.cfg["excel_col_iatf"],
                "country_code": self.cfg["excel_col_country"],
            }
            suppliers = excel_reader.read_suppliers(
                excel_path, columns=excel_columns,
                sheet_name=self.cfg.get("excel_sheet_name", "").strip() or None,
                country_filter="BR")

            emails = {}
            if self.cfg.get("access_db_path") and os.path.exists(self.cfg["access_db_path"]):
                try:
                    emails = contacts_db.load_emails(
                        self.cfg["access_db_path"], self.cfg["access_table"],
                        self.cfg["access_col_rbsno"], self.cfg["access_col_pmd"],
                        self.cfg["access_col_email"])
                except Exception as exc:
                    self.worker.log_message.emit(f"ERROR: could not read the e-mail directory ({exc}).")
            else:
                self.worker.log_message.emit(
                    "Warning: Access e-mail directory not configured or file not found - "
                    "no supplier e-mails available.")

            hdb = HistoryDB(self.cfg["history_db_path"])
            rdb = ReceivedDB(self.cfg["history_db_path"])
            received = supplier_rows.load_received(rdb, suppliers)
            rdb.close()
            kdb = RiskDB(self.cfg["history_db_path"])
            risk = kdb.keys()
            kdb.close()
            rows = []
            for sup in suppliers:
                email = contacts_db.resolve_email(emails, sup.rbsno, sup.pmd)
                rows.extend(supplier_rows.build_rows(
                    sup, email, hdb, self.cfg["overdue_resend_interval_days"], received, risk))
            hdb.close()

            self.worker.analysis_finished.emit(suppliers, rows)
        except Exception as exc:
            self.worker.analysis_error.emit(str(exc))
            self.worker.log_message.emit(f"ERROR traceback summary: {type(exc).__name__}")

    def _on_analysis_error(self, message):
        self.processing = False
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.status_bar.showMessage("Error while processing.")
        self._log(f"ERROR during analysis: {message}")
        QMessageBox.critical(self, "Error", message)

    def _on_analysis_finished(self, suppliers, rows):
        self.processing = False
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.suppliers = suppliers
        self.status_rows = rows

        # totais gerais (combinando os dois tipos de certificado) - usados
        # só para o contador de pendentes, que continua somando os dois
        total_pending = 0
        # totais separados por tipo de certificado - usados pelos 4 cartões,
        # que mostram apenas o tipo da sub-aba selecionada no momento
        self.totals_by_field = {
            field: {"suppliers": 0, "active": 0, "expiring": 0, "expired": 0}
            for field, _label in CERT_TYPES
        }

        for field, _label in CERT_TYPES:
            self.tables[field].setSortingEnabled(False)
            self.tables[field].setRowCount(0)

        for item in rows:
            table = self.tables[item["tab"]]
            field_totals = self.totals_by_field[item["tab"]]
            field_totals["suppliers"] += 1
            kind = item.get("kind", "normal")
            is_blue = kind in ("blue_iso", "blue_iatf")

            if is_blue:
                # fornecedor em dia graças ao OUTRO certificado (ISO <-> IATF)
                field_totals["active"] += 1
                color = ROW_COLORS["blue"]
                status_text = "Supplier ISO" if kind == "blue_iso" else "Supplier IATF"
                expiry_text = item["expiry"].strftime("%Y-%m-%d") if item["expiry"] else "-"
                sort_rank = 2
                sort_days = item["days_left"] or 0
            elif item["tier"] == NO_DATE:
                color = ROW_COLORS["no_date"]
                status_text = "Missing date"
                expiry_text = "-"
                sort_rank = 3
                sort_days = 0
            elif item["tier"] is None:
                field_totals["active"] += 1
                color = ROW_COLORS["active"]
                status_text = "Active"
                expiry_text = item["expiry"].strftime("%Y-%m-%d")
                sort_rank = 2
                sort_days = item["days_left"]
            elif item["tier"] == EXPIRED:
                field_totals["expired"] += 1
                color = ROW_COLORS["expired"]
                status_text = f"Expired {abs(item['days_left'])} day(s) ago"
                expiry_text = item["expiry"].strftime("%Y-%m-%d")
                sort_rank = 0
                sort_days = item["days_left"]
            else:
                field_totals["expiring"] += 1
                color = ROW_COLORS["expiring"]
                status_text = f"Expires in {item['days_left']} day(s)"
                expiry_text = item["expiry"].strftime("%Y-%m-%d")
                sort_rank = 1
                sort_days = item["days_left"]

            notify = item.get("notify", item["tier"] not in (None, NO_DATE))
            if item.get("risk"):
                color = ROW_COLORS["risk"]
                status_text = "Risk Assessment"
                reminder_text = "-"
                reminder_color = theme.PURPLE
                reminder_sort = 2
            elif item.get("received"):
                reminder_text = "Received"
                reminder_color = theme.BLUE
                reminder_sort = 3
            elif is_blue or not notify:
                reminder_text = "-"
                reminder_color = theme.TEXT_MUTED
                reminder_sort = 2
            elif item.get("shown_pending", item["pending"]):
                if item["pending"]:
                    total_pending += 1
                reminder_text = "Pending"
                reminder_color = theme.RED
                reminder_sort = 0
            else:
                reminder_text = "Sent"
                reminder_color = theme.GREEN
                reminder_sort = 1

            r = table.rowCount()
            table.insertRow(r)

            cells = [
                SortItem(item["name"] or "", (item["name"] or "").lower()),
                SortItem(item["rbsno"] or "", _numeric_or_text_key(item["rbsno"])),
                SortItem(item["pmd"] or "", _numeric_or_text_key(item["pmd"])),
                SortItem(expiry_text, item["expiry"] or date.min),
                SortItem(status_text, (re.sub(r"\d+", "", status_text).lower(),
                                       sort_days if sort_days is not None else 0)),  # ordem alfabética do status
                SortItem(reminder_text, reminder_sort),
                SortItem(item["email"] or "(no e-mail on file)", (item["email"] or "").lower()),
            ]
            for c, cell in enumerate(cells):
                cell.setBackground(QBrush(color))
                if c in CENTERED_COLUMNS:
                    cell.setTextAlignment(Qt.AlignCenter)
                if c == 5:
                    cell.setForeground(QBrush(QColor(reminder_color)))
                table.setItem(r, c, cell)

        for field, _label in CERT_TYPES:
            self.tables[field].setSortingEnabled(True)

        self.pending_label.setText(f"{total_pending} pending")

        # atualiza os 4 cartões com os números do tipo de certificado
        # atualmente selecionado na sub-aba
        self._on_subtab_changed(self.sub_tabs.currentIndex())

        counts = {f: {} for f, _l in CERT_TYPES}
        for item in rows:
            kind = item.get("kind", "normal")
            if kind in ("blue_iso", "blue_iatf"):
                cat = "active"  # em dia pelo outro certificado = Ativo
            elif item["tier"] == NO_DATE:
                cat = "missing"
            elif item["tier"] is None:
                cat = "active"
            elif item["tier"] == EXPIRED:
                cat = "expired"
            else:
                cat = "expiring"
            d = counts[item["tab"]]
            d[cat] = d.get(cat, 0) + 1
        self.reports_tab.set_data(counts)
        self.risk_tab.refresh()
        self._refresh_follow()
        self.status_bar.showMessage(f"Analysis complete: {len(suppliers)} supplier(s) read from the Excel file.")
        self._log(f"Analysis complete: {len(suppliers)} supplier(s), {total_pending} reminder(s) pending.")

    def _on_subtab_changed(self, index):
        """Atualiza os 4 cartões do topo para refletir apenas o tipo de
        certificado da sub-aba selecionada (ISO 9001 ou Ambiental)."""
        if index < 0 or index >= len(CERT_TYPES):
            return
        field = CERT_TYPES[index][0]
        totals = self.totals_by_field.get(field, {"suppliers": 0, "active": 0, "expiring": 0, "expired": 0})
        self.card_suppliers.set_value(totals["suppliers"])
        self.card_active.set_value(totals["active"])
        self.card_expiring.set_value(totals["expiring"])
        self.card_expired.set_value(totals["expired"])

    def _send_pending(self):
        if self.processing:
            return
        pending = [row for row in self.status_rows if row["pending"]]
        if not pending:
            QMessageBox.information(self, "Send", "There are no pending reminders to send right now.")
            return

        no_email = [row for row in pending if not row["email"]]
        if no_email:
            answer = QMessageBox.question(
                self, "Suppliers without an e-mail on file",
                f"{len(no_email)} pending supplier(s) have no e-mail on file in Access "
                "and will be skipped. Continue anyway?")
            if answer != QMessageBox.Yes:
                return

        if not OUTLOOK_AVAILABLE:
            self._log("ERROR: send attempted but Outlook automation (pywin32) is not available.")
            QMessageBox.critical(self, "Send", "Outlook automation is not available in this environment.")
            return

        answer = QMessageBox.question(
            self, "Confirm sending", f"{len(pending)} reminder e-mail(s) will be sent now. Continue?")
        if answer != QMessageBox.Yes:
            self._log("Sending cancelled by the user.")
            return
        self._log(f"Starting to send {len(pending)} pending reminder(s)...")

        self.processing = True
        self.status_bar.showMessage("Sending e-mails...")
        self.progress.setRange(0, len(pending))
        self.progress.setValue(0)
        threading.Thread(target=self._send_worker, args=(pending,), daemon=True).start()

    def _send_worker(self, pending):
        hdb = HistoryDB(self.cfg["history_db_path"])
        sent, failed = 0, 0
        for i, item in enumerate(pending, start=1):
            if not item["email"]:
                self.worker.log_message.emit(f"Skipped (no e-mail on file): {item['name']} ({item['rbsno']})")
                self.worker.send_progress.emit(i)
                continue
            try:
                cert_label = CERT_TYPE_LABELS_PT[item["field"]]
                iso_tag = CERT_ISO_TAG[item["field"]]
                subject_prefix = (self.cfg["default_subject_missing_pt"]
                                  if item["tier"] == NO_DATE else self.cfg["default_subject_pt"])

                expiry_str = item["expiry"].strftime("%Y-%m-%d") if item["expiry"] else "-"
                body = templates.build_email_body(
                    item["name"], cert_label, expiry_str, item["days_left"],
                    item["tier"], self.cfg["company_name"], language="pt-br")
                subject = f"{subject_prefix} - {iso_tag} - {item['name']}"
                email_sender.send_email(item["email"], self.cfg.get("bcc_email", ""), subject, body,
                                        sender=self.cfg.get("sender_email", ""))
                hdb.log_send(item["rbsno"], item["pmd"], item["name"], item["field"],
                             item["expiry"], item["tier"], item["email"])
                sent += 1
                self.worker.log_message.emit(
                    f"Sent: {item['name']} ({item['label']}, "
                    f"{reminder_engine.tier_label(item['tier'])}) -> {item['email']}")
            except Exception as exc:
                failed += 1
                self.worker.log_message.emit(f"Failed to send to {item['name']}: {exc}")
            self.worker.send_progress.emit(i)
        hdb.close()
        self.worker.send_finished.emit(sent, failed)

    def _on_send_progress(self, value):
        self.progress.setValue(value)

    def _on_send_finished(self, sent, failed):
        self.processing = False
        self.status_bar.showMessage(f"Sending complete: {sent} sent, {failed} failed.")
        self._log(f"Sending complete: {sent} sent, {failed} failed.")
        notifications.notify("Tracking Certificate", f"{sent} reminder(s) sent, {failed} failed.", log=self._log)
        self._refresh_history()
        self._load_and_analyze()

# -*- coding: utf-8 -*-
"""Telas "Risk Assessment" e "Reports" (gráficos desenhados com QPainter,
sem depender do módulo QtCharts)."""
import math

import os

from PySide6.QtCore import Qt, QRectF, QPointF, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtGui import QColor, QBrush, QPainter, QPen, QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QFrame, QPushButton,
    QTableWidget, QHeaderView, QLineEdit, QComboBox, QDialog, QFormLayout,
    QMessageBox, QScrollArea, QSizePolicy, QFileDialog, QListWidget, QCompleter,
)

import cert_files
from received_db import ReceivedDB
from risk_db import RiskDB
from reminder_engine import compute_status, EXPIRED, NO_DATE
from supplier_rows import CERT_TYPES

# IATF 16949 é redundante com ISO 9001 nos relatórios: um único gráfico "Quality"
REPORT_TYPES = [("iso_expiry", "Quality"), ("env_expiry", "Environmental (ISO 14001)")]
from ui import theme
from ui.common import SortItem, numeric_or_text_key

RISK_CERTS = ["ISO 9001", "IATF 16949", "ISO 9001 or IATF 16949"]
RISK_STATUSES = ["Open", "In progress", "Completed"]
RISK_HEADERS = ["Supplier", "RBSNO", "PMD", "Certificate", "Status", "Added", "Files", "Notes"]
STATUS_COLORS = {"Open": "#fff3cd", "In progress": "#d6e4ff", "Completed": "#d9f2df"}


# --------------------------------------------------------------------------- Risk Assessment

class RiskDialog(QDialog):
    def __init__(self, parent, suppliers, record=None):
        super().__init__(parent)
        self.setWindowTitle("Edit supplier" if record else "Add supplier")
        self.resize(520, 430)
        self.existing = []   # caminhos já salvos
        self.new_files = []  # arquivos escolhidos agora (copiados ao salvar)
        self.suppliers = {s.name: s for s in suppliers}

        form = QFormLayout(self)
        form.setContentsMargins(22, 20, 22, 20)
        form.setSpacing(10)

        self.name = QComboBox()
        self.name.setEditable(True)
        self.name.addItems(sorted(self.suppliers))
        self.name.setCurrentText("")
        self.name.currentTextChanged.connect(self._autofill)
        self.rbsno = QLineEdit()
        self.pmd = QLineEdit()
        self.cert = QComboBox()
        self.cert.addItems(RISK_CERTS)
        self.status = QComboBox()
        self.status.addItems(RISK_STATUSES)
        self.notes = QLineEdit()

        form.addRow("Supplier name:", self.name)
        form.addRow("RBSNO:", self.rbsno)
        form.addRow("PMD:", self.pmd)
        form.addRow("Missing certificate:", self.cert)
        form.addRow("Status:", self.status)
        form.addRow("Notes:", self.notes)

        self.file_list = QListWidget()
        self.file_list.setFixedHeight(80)
        form.addRow("Attachments:", self.file_list)
        frow = QHBoxLayout()
        b_add = QPushButton("Attach files...")
        b_add.setObjectName("SecondaryButton")
        b_add.clicked.connect(self._attach)
        b_rm = QPushButton("Remove selected")
        b_rm.setObjectName("SecondaryButton")
        b_rm.clicked.connect(self._detach)
        frow.addWidget(b_add)
        frow.addWidget(b_rm)
        frow.addStretch()
        form.addRow("", frow)

        row = QHBoxLayout()
        row.addStretch()
        cancel = QPushButton("Cancel")
        cancel.setObjectName("SecondaryButton")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("Save")
        ok.clicked.connect(self._accept)
        row.addWidget(cancel)
        row.addWidget(ok)
        form.addRow(row)

        if record:
            _id, rbsno, pmd, name, cert, status, _added, notes, files = record
            self.existing = [f for f in (files or "").split("|") if f]
            self.name.setCurrentText(name or "")
            self.rbsno.setText(rbsno or "")
            self.pmd.setText(pmd or "")
            self.cert.setCurrentText(cert or RISK_CERTS[0])
            self.status.setCurrentText(status or RISK_STATUSES[0])
            self.notes.setText(notes or "")
        self._refresh_files()

    def _refresh_files(self):
        self.file_list.clear()
        for f in self.existing:
            self.file_list.addItem(os.path.basename(f))
        for f in self.new_files:
            self.file_list.addItem(os.path.basename(f) + "  (new)")

    def _attach(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Attach files")
        self.new_files.extend(paths)
        self._refresh_files()

    def _detach(self):
        i = self.file_list.currentRow()
        if i < 0:
            return
        if i < len(self.existing):
            del self.existing[i]  # só remove a referência; o arquivo continua na pasta
        else:
            del self.new_files[i - len(self.existing)]
        self._refresh_files()

    def _autofill(self, text):
        sup = self.suppliers.get(text)
        if sup:
            self.rbsno.setText(sup.rbsno)
            self.pmd.setText(sup.pmd)

    def _accept(self):
        if not self.name.currentText().strip() and not self.rbsno.text().strip():
            QMessageBox.warning(self, "Risk Assessment", "Please fill in the supplier name or RBSNO.")
            return
        self.accept()

    def values(self):
        return (self.rbsno.text().strip(), self.pmd.text().strip(), self.name.currentText().strip(),
                self.cert.currentText(), self.status.currentText(), self.notes.text().strip())

    def file_state(self):
        return self.existing, self.new_files


class RiskAssessmentPage(QWidget):
    """Suppliers that need a field action (Risk Assessment): add / edit /
    remove, search by name or RBSNO and filter by column."""

    def __init__(self, db_path_getter, suppliers_getter, log, folder_getter=lambda: "", on_change=lambda: None):
        super().__init__()
        self._folder = folder_getter
        self._on_change = on_change
        self._db_path = db_path_getter
        self._suppliers = suppliers_getter
        self._log = log

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)

        hint = QLabel("Suppliers that do not hold the requested certificate and need a field action "
                      "(Risk Assessment).")
        hint.setObjectName("SectionHint")
        layout.addWidget(hint)

        bar = QHBoxLayout()
        bar.setSpacing(10)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search by supplier name or RBSNO...")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.refresh)
        self.filter_cert = QComboBox()
        self.filter_cert.addItems(["Certificate: All"] + RISK_CERTS)
        self.filter_status = QComboBox()
        self.filter_status.addItems(["Status: All"] + RISK_STATUSES)
        for c in (self.filter_cert, self.filter_status):
            c.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.search, stretch=1)
        bar.addWidget(self.filter_cert)
        bar.addWidget(self.filter_status)
        layout.addLayout(bar)

        self.table = QTableWidget(0, len(RISK_HEADERS))
        self.table.setHorizontalHeaderLabels(RISK_HEADERS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(lambda _i: self.edit())
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Stretch)
        h.setSectionResizeMode(7, QHeaderView.Stretch)
        for c in (1, 2, 3, 4, 5, 6):
            h.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        layout.addWidget(self.table, stretch=1)

        bottom = QHBoxLayout()
        self.count_label = QLabel("0 supplier(s)")
        self.count_label.setObjectName("PendingCounter")
        bottom.addWidget(self.count_label)
        bottom.addStretch()
        for text, slot, obj in (("Open folder", self.open_folder, "SecondaryButton"), ("Add", self.add, None), ("Edit", self.edit, "SecondaryButton"),
                                ("Remove", self.remove, "SecondaryButton")):
            b = QPushButton(text)
            if obj:
                b.setObjectName(obj)
            b.clicked.connect(slot)
            bottom.addWidget(b)
        layout.addLayout(bottom)
        self.refresh()

    def _db(self):
        return RiskDB(self._db_path())

    def refresh(self, *_):
        try:
            db = self._db()
            records = db.all()
            db.close()
        except Exception as exc:
            self._log(f"ERROR reading Risk Assessment list: {exc}")
            return
        q = self.search.text().strip().lower()
        cert = self.filter_cert.currentText()
        status = self.filter_status.currentText()
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        for rec in records:
            rec_id, rbsno, pmd, name, certificate, st, added, notes, files = rec
            if q and q not in (name or "").lower() and q not in (rbsno or "").lower():
                continue
            if not cert.endswith("All") and certificate != cert:
                continue
            if not status.endswith("All") and st != status:
                continue
            r = self.table.rowCount()
            self.table.insertRow(r)
            cells = [
                SortItem(name or "", (name or "").lower()),
                SortItem(rbsno or "", numeric_or_text_key(rbsno)),
                SortItem(pmd or "", numeric_or_text_key(pmd)),
                SortItem(certificate or ""), SortItem(st or ""),
                SortItem(added or ""),
                SortItem(str(len([f for f in (files or "").split("|") if f])),
                         len([f for f in (files or "").split("|") if f])),
                SortItem(notes or ""),
            ]
            cells[0].setData(Qt.UserRole, rec_id)
            for c, cell in enumerate(cells):
                if c in (1, 2, 3, 4, 5, 6):
                    cell.setTextAlignment(Qt.AlignCenter)
                if c == 4:
                    cell.setBackground(QBrush(QColor(STATUS_COLORS.get(st, "#ffffff"))))
                self.table.setItem(r, c, cell)
        self.table.setSortingEnabled(True)
        self.count_label.setText(f"{self.table.rowCount()} supplier(s)")

    def _selected_id(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "Risk Assessment", "Select a supplier in the table first.")
            return None
        return self.table.item(rows[0].row(), 0).data(Qt.UserRole)

    def _store_files(self, values, existing, new_files):
        """Copia os anexos novos para a pasta do fornecedor e devolve a lista final."""
        paths = list(existing)
        if new_files:
            base = self._folder()
            if not base:
                QMessageBox.warning(self, "Risk Assessment",
                                    "Set the certificates folder in Settings to save attachments.")
                return None
            for src in new_files:
                paths.append(cert_files.save_file(base, values[2], values[0], src))
        return "|".join(paths)

    def add(self):
        dlg = RiskDialog(self, self._suppliers())
        if dlg.exec():
            v = dlg.values()
            files = self._store_files(v, *dlg.file_state())
            if files is None:
                return
            db = self._db()
            db.add(*v, files=files)
            db.close()
            self._log(f"Risk Assessment: added {v[2] or v[0]}.")
            self.refresh()
            self._on_change()

    def edit(self):
        rec_id = self._selected_id()
        if rec_id is None:
            return
        db = self._db()
        record = next((r for r in db.all() if r[0] == rec_id), None)
        db.close()
        if not record:
            return
        dlg = RiskDialog(self, self._suppliers(), record)
        if dlg.exec():
            v = dlg.values()
            files = self._store_files(v, *dlg.file_state())
            if files is None:
                return
            db = self._db()
            db.update(rec_id, *v, files=files)
            db.close()
            self._log(f"Risk Assessment: edited {v[2] or v[0]}.")
            self.refresh()
            self._on_change()

    def open_folder(self):
        rec_id = self._selected_id()
        base = self._folder()
        if rec_id is None:
            return
        if not base:
            QMessageBox.information(self, "Risk Assessment", "Set the certificates folder in Settings first.")
            return
        row = self.table.selectionModel().selectedRows()[0].row()
        name, rbsno = self.table.item(row, 0).text(), self.table.item(row, 1).text()
        folder = cert_files.supplier_folder(base, name, rbsno, create=False)
        if os.path.isdir(folder):
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
        else:
            QMessageBox.information(self, "Risk Assessment", "No folder for this supplier yet (no files attached).")

    def remove(self):
        rec_id = self._selected_id()
        if rec_id is None:
            return
        if QMessageBox.question(self, "Remove", "Remove the selected supplier from the list?") != QMessageBox.Yes:
            return
        db = self._db()
        db.delete(rec_id)
        db.close()
        self._log("Risk Assessment: supplier removed.")
        self.refresh()
        self._on_change()


# --------------------------------------------------------------------------- Certificado recebido

RECEIVED_CERTS = ["ISO 9001", "IATF 16949", "Environmental (ISO 14001)"]


class ReceivedCertDialog(QDialog):
    """Registra que o fornecedor já enviou o certificado (ainda não subido
    no datalake) para suspender a cobrança por e-mail."""

    def __init__(self, parent, suppliers, folder, db_path, log):
        super().__init__(parent)
        self.setWindowTitle("Register Received Certificate")
        self.resize(560, 520)
        self.suppliers = {s.name: s for s in suppliers}
        self.folder, self.db_path, self._log = folder, db_path, log
        self.src = ""

        form = QFormLayout(self)
        form.setContentsMargins(22, 20, 22, 20)
        form.setSpacing(10)

        self.name = QComboBox()
        self.name.setEditable(True)
        self.name.setInsertPolicy(QComboBox.NoInsert)
        self.name.addItems(sorted(self.suppliers))
        comp = QCompleter(sorted(self.suppliers), self.name)
        comp.setFilterMode(Qt.MatchContains)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        self.name.setCompleter(comp)
        self.name.setCurrentText("")
        self.name.currentTextChanged.connect(self._fill)
        self.rbsno = QLineEdit()
        self.rbsno.setReadOnly(True)
        self.info = QLabel("Select a supplier to see the certificate status.")
        self.info.setObjectName("SectionHint")
        self.info.setTextFormat(Qt.RichText)
        self.cert = QComboBox()
        self.cert.addItems(RECEIVED_CERTS)
        self.file_label = QLabel("No file attached (required)")
        self.file_label.setObjectName("SectionHint")
        b_file = QPushButton("Attach certificate file...")
        b_file.setObjectName("SecondaryButton")
        b_file.clicked.connect(self._pick)

        form.addRow("Supplier:", self.name)
        form.addRow("RBSNO:", self.rbsno)
        form.addRow("Current status:", self.info)
        form.addRow("Certificate attached:", self.cert)
        form.addRow("File (required):", b_file)
        form.addRow("", self.file_label)

        row = QHBoxLayout()
        row.addStretch()
        cancel = QPushButton("Cancel")
        cancel.setObjectName("SecondaryButton")
        cancel.clicked.connect(self.reject)
        ok = QPushButton("Save")
        ok.clicked.connect(self._save)
        row.addWidget(cancel)
        row.addWidget(ok)
        form.addRow(row)

    @staticmethod
    def _fmt(label, expiry):
        if expiry is None:
            return f"{label}: <b>Missing date</b>"
        tier, days = compute_status(expiry)
        state = f"Expired {abs(days)} day(s) ago" if tier == EXPIRED else (
            f"Valid, expires in {days} day(s)")
        return f"{label}: <b>{state}</b> ({expiry:%Y-%m-%d})"

    def _fill(self, text):
        sup = self.suppliers.get(text)
        if not sup:
            self.rbsno.clear()
            self.info.setText("Select a supplier to see the certificate status.")
            return
        self.rbsno.setText(sup.rbsno)
        self.info.setText("<br>".join([
            self._fmt("ISO 9001", sup.iso_expiry), self._fmt("IATF 16949", sup.iatf_expiry),
            self._fmt("ISO 14001", sup.env_expiry)]))

    def _pick(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select certificate file")
        if path:
            self.src = path
            self.file_label.setText(os.path.basename(path))

    def _save(self):
        sup = self.suppliers.get(self.name.currentText())
        if not sup:
            QMessageBox.warning(self, "Register", "Select a supplier from the list.")
            return
        if not self.src:
            QMessageBox.warning(self, "Register", "Attaching the certificate file is required.")
            return
        if not self.folder:
            QMessageBox.warning(self, "Register",
                                "Set the certificates folder in Settings (Certificate Files) first.")
            return
        try:
            dest = cert_files.save_file(self.folder, sup.name, sup.rbsno, self.src)
            db = ReceivedDB(self.db_path)
            db.add(sup.rbsno, sup.pmd, sup.name, self.cert.currentText(), os.path.basename(dest), dest)
            db.close()
        except Exception as exc:
            self._log(f"ERROR registering certificate for {sup.name}: {exc}")
            QMessageBox.critical(self, "Register", f"Could not save the certificate:\n{exc}")
            return
        self._log(f"Certificate received: {sup.name} ({sup.rbsno}) - {self.cert.currentText()} -> {dest}")
        self.accept()


# --------------------------------------------------------------------------- Reports

CATEGORIES = [
    ("active", "Active", theme.GREEN),
    ("expiring", "Expiring (90/60/30)", theme.AMBER),
    ("expired", "Expired", theme.RED),
    ("missing", "Missing date", "#9aa5b1"),
]
FOLLOW_CATEGORIES = [
    ("risk", "In Risk Assessment", theme.PURPLE),
    ("files", "Files added manually", theme.BLUE),
    ("missing", "Missing date (Quality)", "#9aa5b1"),
]


class DonutChart(QWidget):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(170, 170)
        self.values = []  # [(valor, cor)]

    def set_values(self, values):
        self.values = values
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("white"))
        side = min(self.width(), self.height()) - 24
        thick = side * 0.17
        rect = QRectF((self.width() - side) / 2 + thick / 2, (self.height() - side) / 2 + thick / 2,
                      side - thick, side - thick)
        total = sum(v for v, _c in self.values)
        pen = QPen(QColor("#e6ebf1"), thick, Qt.SolidLine, Qt.FlatCap)
        p.setPen(pen)
        if total <= 0:
            p.drawArc(rect, 0, 360 * 16)
        else:
            angle = 90 * 16
            for value, color in self.values:
                if value <= 0:
                    continue
                span = -int(round(value / total * 360 * 16))
                p.setPen(QPen(QColor(color), thick, Qt.SolidLine, Qt.FlatCap))
                p.drawArc(rect, angle, span)
                angle += span
        p.setPen(QColor(theme.TEXT_DARK))
        f = QFont(self.font())
        f.setPointSize(20)
        f.setBold(True)
        p.setFont(f)
        p.drawText(self.rect().adjusted(0, -8, 0, 0), Qt.AlignCenter, str(int(total)))
        f.setPointSize(8)
        f.setBold(False)
        p.setFont(f)
        p.setPen(QColor(theme.TEXT_MUTED))
        p.drawText(self.rect().adjusted(0, 34, 0, 0), Qt.AlignCenter, "suppliers")


class BarChart(QWidget):
    """Barras agrupadas: um grupo por certificado, uma barra por categoria."""

    def __init__(self, categories=None):
        super().__init__()
        self.categories = categories or CATEGORIES
        self.setMinimumHeight(300)
        self.groups = []  # [(rótulo, {cat: valor})]

    def set_groups(self, groups):
        self.groups = groups
        self.update()

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("white"))
        left, right, top, bottom = 44, 16, 16, 40
        area = QRectF(left, top, self.width() - left - right, self.height() - top - bottom)
        peak = max([v for _l, d in self.groups for v in d.values()] + [1])
        step = 1 if peak <= 5 else 5 * 10 ** max(0, len(str(peak)) - 2)
        top_val = int(math.ceil(peak / step) * step)
        small = QFont(self.font())
        small.setPointSize(8)
        p.setFont(small)
        for i in range(0, top_val + 1, step):
            y = area.bottom() - area.height() * i / top_val
            p.setPen(QPen(QColor("#eef1f5"), 1))
            p.drawLine(QPointF(area.left(), y), QPointF(area.right(), y))
            p.setPen(QColor(theme.TEXT_MUTED))
            p.drawText(QRectF(0, y - 8, left - 6, 16), Qt.AlignRight | Qt.AlignVCenter, str(i))
        n = max(len(self.groups), 1)
        gw = area.width() / n
        bw = min(34.0, gw * 0.7 / len(self.categories))
        for gi, (label, data) in enumerate(self.groups):
            gx = area.left() + gi * gw + (gw - bw * len(self.categories)) / 2
            for ci, (key, _name, color) in enumerate(self.categories):
                v = data.get(key, 0)
                h = area.height() * v / top_val
                bar = QRectF(gx + ci * bw + 1, area.bottom() - h, bw - 2, h)
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(color))
                p.drawRect(bar)
                if v:
                    p.setPen(QColor(theme.TEXT_DARK))
                    p.drawText(QRectF(bar.left() - 6, bar.top() - 16, bar.width() + 12, 14),
                               Qt.AlignCenter, str(v))
            p.setPen(QColor(theme.TEXT_DARK))
            p.drawText(QRectF(area.left() + gi * gw, area.bottom() + 8, gw, 18), Qt.AlignCenter, label)


class ReportsPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(scroll)
        content = QWidget()
        scroll.setWidget(content)
        lay = QVBoxLayout(content)
        lay.setContentsMargins(22, 20, 22, 20)
        lay.setSpacing(16)

        self.note = QLabel("Load and analyze the SRM Excel file in the Panel to see the charts.")
        self.note.setObjectName("SectionHint")
        self.note.setWordWrap(True)
        lay.addWidget(self.note)

        row = QHBoxLayout()
        row.setSpacing(14)
        self.donuts = {}
        for field, label in REPORT_TYPES:
            card = QFrame()
            card.setObjectName("StatCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(16, 14, 16, 14)
            title = QLabel(label.upper())
            title.setObjectName("StatTitle")
            donut = DonutChart()
            legend = QLabel()
            legend.setStyleSheet("background: transparent;")
            legend.setTextFormat(Qt.RichText)
            cl.addWidget(title)
            cl.addWidget(donut, stretch=1)
            legend.setAlignment(Qt.AlignHCenter)
            cl.addWidget(legend)
            row.addWidget(card)
            self.donuts[field] = (donut, legend)
        lay.addLayout(row)

        bottom = QHBoxLayout()
        bottom.setSpacing(14)
        self.bar = BarChart()
        bottom.addWidget(self._chart_card("COMPARISON BY CERTIFICATE", self.bar, CATEGORIES), stretch=3)
        self.follow = BarChart(FOLLOW_CATEGORIES)
        bottom.addWidget(self._chart_card("FOLLOW-UP", self.follow, FOLLOW_CATEGORIES), stretch=2)
        lay.addLayout(bottom)
        lay.addStretch()

    @staticmethod
    def _chart_card(title, chart, categories):
        card = QFrame()
        card.setObjectName("StatCard")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(16, 14, 16, 14)
        t = QLabel(title)
        t.setObjectName("StatTitle")
        cl.addWidget(t)
        cl.addWidget(chart)
        legend = " &nbsp; ".join(f'<span style="color:{c}">&#9632;</span> {n}' for _k, n, c in categories)
        lg = QLabel(legend)
        lg.setStyleSheet("background: transparent;")
        lg.setTextFormat(Qt.RichText)
        lg.setWordWrap(True)
        lg.setAlignment(Qt.AlignCenter)
        cl.addWidget(lg)
        return card

    def set_follow(self, follow):
        """follow: {risk, files, missing} (nº de fornecedores)"""
        self.follow.set_groups([("Suppliers", follow)])

    def set_data(self, counts):
        """counts: {campo: {categoria: n}}"""
        self.note.setText("Active = valid certificate (more than 90 days left). Expiring = within the "
                          "90/60/30-day alerts. A supplier holding ISO 9001 or IATF 16949 counts as Active "
                          "in both tabs.")
        groups = []
        for field, label in REPORT_TYPES:
            data = counts.get(field, {})
            donut, legend = self.donuts[field]
            donut.set_values([(data.get(k, 0), c) for k, _n, c in CATEGORIES])
            total = sum(data.values()) or 1
            legend.setText('<table align="center" cellspacing="2">' + "".join(
                f'<tr><td style="color:{c}">&#9632;</td><td>{n}:&nbsp;<b>{data.get(k, 0)}</b> '
                f'({data.get(k, 0) * 100 / total:.0f}%)</td></tr>'
                for k, n, c in CATEGORIES) + "</table>")
            groups.append((label.split(" (")[0], data))
        self.bar.set_groups(groups)

# -*- coding: utf-8 -*-
"""
Lista de fornecedores que precisam de Risk Assessment (ação em campo).
Guardada no mesmo arquivo SQLite do histórico de envios.
"""
import sqlite3
from datetime import date

SCHEMA = """
CREATE TABLE IF NOT EXISTS risk_assessment (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rbsno TEXT,
    pmd TEXT,
    supplier_name TEXT,
    certificate TEXT,
    status TEXT,
    added_date TEXT,
    notes TEXT,
    files TEXT
);
"""

COLUMNS = "id, rbsno, pmd, supplier_name, certificate, status, added_date, notes, files"


class RiskDB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.execute(SCHEMA)
        cols = [r[1] for r in self.conn.execute("PRAGMA table_info(risk_assessment)")]
        if "files" not in cols:  # migração de bancos criados antes dos anexos
            self.conn.execute("ALTER TABLE risk_assessment ADD COLUMN files TEXT")
        self.conn.commit()

    def close(self):
        self.conn.close()

    def all(self):
        return self.conn.execute(f"SELECT {COLUMNS} FROM risk_assessment ORDER BY id DESC").fetchall()

    def keys(self):
        """RBSNOs e nomes (minúsculos) dos fornecedores em Risk Assessment."""
        out = set()
        for rbsno, name in self.conn.execute("SELECT rbsno, supplier_name FROM risk_assessment"):
            if rbsno:
                out.add(rbsno.strip())
            if name:
                out.add(name.strip().lower())
        return out

    def add(self, rbsno, pmd, name, certificate, status, notes, files=""):
        self.conn.execute(
            "INSERT INTO risk_assessment (rbsno, pmd, supplier_name, certificate, status, added_date, notes, files) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (rbsno, pmd, name, certificate, status, str(date.today()), notes, files))
        self.conn.commit()

    def update(self, rec_id, rbsno, pmd, name, certificate, status, notes, files=""):
        self.conn.execute(
            "UPDATE risk_assessment SET rbsno=?, pmd=?, supplier_name=?, certificate=?, status=?, notes=?, files=? "
            "WHERE id=?", (rbsno, pmd, name, certificate, status, notes, files, rec_id))
        self.conn.commit()

    def delete(self, rec_id):
        self.conn.execute("DELETE FROM risk_assessment WHERE id=?", (rec_id,))
        self.conn.commit()

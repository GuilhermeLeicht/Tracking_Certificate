# -*- coding: utf-8 -*-
"""
Certificados recebidos do fornecedor mas ainda não enviados ao datalake.
Enquanto o registro existir, os lembretes daquele grupo de certificado
("quality" = ISO 9001/IATF 16949, "env" = Ambiental) ficam suspensos.
Guardado no mesmo SQLite do histórico.
"""
import sqlite3
from datetime import date

SCHEMA = """
CREATE TABLE IF NOT EXISTS received_certs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rbsno TEXT, pmd TEXT, supplier_name TEXT,
    certificate TEXT, grp TEXT, file_name TEXT, saved_path TEXT, received_date TEXT
);
"""


def group_of(certificate):
    return "env" if "14001" in certificate or "Environ" in certificate else "quality"


class ReceivedDB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.execute(SCHEMA)
        self.conn.commit()

    def close(self):
        self.conn.close()

    def add(self, rbsno, pmd, name, certificate, file_name, saved_path):
        # mesmo arquivo (mesmo caminho/nome) registrado de novo: substitui o registro
        self.conn.execute("DELETE FROM received_certs WHERE saved_path=?", (saved_path,))
        self.conn.execute(
            "INSERT INTO received_certs (rbsno, pmd, supplier_name, certificate, grp, file_name, "
            "saved_path, received_date) VALUES (?,?,?,?,?,?,?,?)",
            (rbsno, pmd, name, certificate, group_of(certificate), file_name, saved_path, str(date.today())))
        self.conn.commit()

    def prune_missing_files(self):
        """Remove os registros cujo arquivo não existe mais na pasta."""
        import os
        removed = 0
        for rid, path in self.conn.execute("SELECT id, saved_path FROM received_certs").fetchall():
            if not path or not os.path.isfile(path):
                self.conn.execute("DELETE FROM received_certs WHERE id=?", (rid,))
                removed += 1
        self.conn.commit()
        return removed

    def suppressed_groups(self):
        return {(r[0], r[1]) for r in self.conn.execute("SELECT DISTINCT rbsno, grp FROM received_certs")}

    def all_certs(self):
        return self.conn.execute("SELECT DISTINCT rbsno, certificate FROM received_certs").fetchall()

    def delete_cert(self, rbsno, certificate):
        self.conn.execute("DELETE FROM received_certs WHERE rbsno=? AND certificate=?", (rbsno, certificate))
        self.conn.commit()

    def delete_group(self, rbsno, grp):
        self.conn.execute("DELETE FROM received_certs WHERE rbsno=? AND grp=?", (rbsno, grp))
        self.conn.commit()

    def supplier_count(self):
        return self.conn.execute("SELECT COUNT(DISTINCT rbsno) FROM received_certs").fetchone()[0]

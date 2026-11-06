# -*- coding: utf-8 -*-
"""
Histórico de envios, guardado localmente em SQLite (nada a configurar).
Usado para saber o que já foi enviado, para nunca disparar o mesmo aviso
duas vezes na mesma faixa.
"""
import sqlite3
from datetime import date, datetime

SCHEMA = """
CREATE TABLE IF NOT EXISTS notifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rbsno TEXT NOT NULL,
    pmd TEXT,
    supplier_name TEXT,
    cert_type TEXT NOT NULL,
    expiry_date TEXT NOT NULL,
    tier TEXT NOT NULL,
    sent_date TEXT NOT NULL,
    recipient_email TEXT
);
"""


class HistoryDB:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.execute(SCHEMA)
        self.conn.commit()

    def close(self):
        self.conn.close()

    def get_last_send(self, rbsno, cert_type, expiry_date):
        cur = self.conn.execute(
            "SELECT tier, sent_date FROM notifications "
            "WHERE rbsno=? AND cert_type=? AND expiry_date=? "
            "ORDER BY id DESC LIMIT 1",
            (rbsno, cert_type, str(expiry_date)),
        )
        return cur.fetchone()

    def log_send(self, rbsno, pmd, supplier_name, cert_type, expiry_date, tier, recipient_email):
        self.conn.execute(
            "INSERT INTO notifications "
            "(rbsno, pmd, supplier_name, cert_type, expiry_date, tier, sent_date, recipient_email) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (rbsno, pmd, supplier_name, cert_type, str(expiry_date), tier, str(date.today()), recipient_email),
        )
        self.conn.commit()

    def history(self, limit=1000):
        cur = self.conn.execute(
            "SELECT rbsno, pmd, supplier_name, cert_type, expiry_date, tier, sent_date, recipient_email "
            "FROM notifications ORDER BY id DESC LIMIT ?",
            (limit,),
        )
        return cur.fetchall()

    def count(self):
        """Total de envios já registrados (independente do limite usado em history())."""
        cur = self.conn.execute("SELECT COUNT(*) FROM notifications")
        return cur.fetchone()[0]

# -*- coding: utf-8 -*-
"""
Executa o Tracking Certificate automaticamente, sem abrir a interface gráfica.

Feito para ser chamado pelo Agendador de Tarefas do Windows (ou similar),
rodando sozinho em um horário fixo, sem ninguém precisar abrir o app.

Uso:
    python main.py --auto

Grava um log em 'tracking_auto_log.txt' (na mesma pasta do app) e também
imprime no console. As mensagens de log ficam em inglês, igual ao restante
do app (só os comentários deste arquivo estão em português).

Requisitos para funcionar sem ninguém sentado no computador:
    - O computador precisa estar ligado (não em suspensão) no horário agendado.
    - O Outlook precisa estar instalado e com uma conta já configurada.
    - As configurações (Excel, Access, colunas, regras) já devem estar salvas
      em 'tracking_config.json' - ou seja, rode o app normalmente pelo menos
      uma vez, configure tudo na aba Settings e selecione o Excel.
"""
import os
from datetime import datetime

import config as cfgmod
import contacts_db
import excel_reader
import reminder_engine
from reminder_engine import NO_DATE
import templates
import notifications
from received_db import ReceivedDB
from supplier_rows import load_received, build_rows, CERT_TYPE_LABELS_PT, CERT_ISO_TAG
from history_db import HistoryDB

LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracking_auto_log.txt")

def _log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def run_automatic():
    _log("========== Automatic run started ==========")
    cfg = cfgmod.load_config()

    excel_path = cfg.get("excel_path", "")
    if not excel_path or not os.path.exists(excel_path):
        _log("ERROR: no Excel file configured, or the file no longer exists at the saved path. "
             "Open the app normally, select the Excel file via '1. Select SRM Excel File' "
             "to save the path, then try again.")
        return 1

    excel_columns = {
        "rbsno": cfg["excel_col_rbsno"],
        "pmd": cfg["excel_col_pmd"],
        "name": cfg["excel_col_name"],
        "iso_expiry": cfg["excel_col_iso"],
        "env_expiry": cfg["excel_col_env"],
        "iatf_expiry": cfg["excel_col_iatf"],
        "country_code": cfg["excel_col_country"],
    }

    try:
        suppliers = excel_reader.read_suppliers(
            excel_path, columns=excel_columns,
            sheet_name=cfg.get("excel_sheet_name", "").strip() or None,
            country_filter="BR")
    except Exception as exc:
        _log(f"ERROR reading the Excel file: {exc}")
        return 1

    emails = {}
    if cfg.get("access_db_path") and os.path.exists(cfg["access_db_path"]):
        try:
            emails = contacts_db.load_emails(
                cfg["access_db_path"], cfg["access_table"],
                cfg["access_col_rbsno"], cfg["access_col_pmd"], cfg["access_col_email"])
        except Exception as exc:
            _log(f"Warning: could not read the e-mail directory: {exc}")
    else:
        _log("Warning: Access e-mail directory not configured or not found.")

    try:
        import email_sender
    except Exception as exc:
        _log(f"ERROR: Outlook automation is unavailable in this environment ({exc}). "
             "Automatic sending only works on Windows with Outlook installed.")
        return 1

    reminder_engine.WARNING_THRESHOLDS = tuple(cfg["reminder_days"])
    hdb = HistoryDB(cfg["history_db_path"])
    subject_prefix = cfg["default_subject_pt"]

    from risk_db import RiskDB
    kdb = RiskDB(cfg["history_db_path"])
    risk = kdb.keys()
    kdb.close()
    rdb = ReceivedDB(cfg["history_db_path"])
    received = load_received(rdb, suppliers)
    rdb.close()
    sent, failed, skipped = 0, 0, 0
    for sup in suppliers:
        email = contacts_db.resolve_email(emails, sup.rbsno, sup.pmd)
        for item in build_rows(sup, email, hdb, cfg["overdue_resend_interval_days"], received, risk):
            if not item["pending"]:
                continue
            label = item["label"]
            if not email:
                skipped += 1
                _log(f"Skipped (no e-mail on file): {sup.name} ({sup.rbsno}) - {label}")
                continue
            try:
                field = item["field"]
                expiry = item["expiry"]
                expiry_str = expiry.strftime("%Y-%m-%d") if expiry else "-"
                body = templates.build_email_body(
                    sup.name, CERT_TYPE_LABELS_PT[field], expiry_str,
                    item["days_left"], item["tier"], cfg["company_name"], language="pt-br")
                prefix = cfg["default_subject_missing_pt"] if item["tier"] == NO_DATE else subject_prefix
                subject = f"{prefix} - {CERT_ISO_TAG[field]} - {sup.name}"
                email_sender.send_email(email, cfg.get("bcc_email", ""), subject, body,
                                        sender=cfg.get("sender_email", ""))
                hdb.log_send(sup.rbsno, sup.pmd, sup.name, field, expiry, item["tier"], email)
                sent += 1
                _log(f"Sent: {sup.name} ({label}, {reminder_engine.tier_label(item['tier'])}) -> {email}")
            except Exception as exc:
                failed += 1
                _log(f"Failed to send to {sup.name} ({label}): {exc}")

    hdb.close()
    _log(f"Automatic run finished: {sent} sent, {failed} failed, "
         f"{skipped} skipped for missing e-mail.")
    _log("====================================================")
    notifications.notify("Tracking Certificate", f"{sent} sent, {failed} failed, {skipped} skipped.", log=_log)
    return 0

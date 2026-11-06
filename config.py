# -*- coding: utf-8 -*-
"""
Configuração persistente do Tracking Certificate.
As configurações são salvas em 'tracking_config.json', ao lado do app.
"""
import json
import os

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracking_config.json")

_BASE = os.path.dirname(os.path.abspath(__file__))


def _migrate_legacy_files():
    """Renomeia os arquivos antigos (renovacert_*) para os novos nomes
    (tracking_*), uma única vez, preservando configurações e histórico."""
    for old, new in (("renovacert_config.json", "tracking_config.json"),
                     ("renovacert_history.db", "tracking.db"),
                     ("renovacert_app_log.txt", "tracking_app_log.txt"),
                     ("renovacert_auto_log.txt", "tracking_auto_log.txt")):
        old_p, new_p = os.path.join(_BASE, old), os.path.join(_BASE, new)
        if os.path.exists(old_p) and not os.path.exists(new_p):
            try:
                os.rename(old_p, new_p)
            except OSError:
                pass


_migrate_legacy_files()

DEFAULT_CONFIG = {
    # Arquivo Excel exportado do SRM
    "excel_path": "",

    # Nome da aba dentro do Excel que contém os dados dos fornecedores.
    # Deixe em branco para o app procurar automaticamente em todas as abas.
    "excel_sheet_name": "",

    # Nomes das colunas no Excel do SRM (editável na aba Settings)
    "excel_col_rbsno": "Basic Data: RBSNO",
    "excel_col_pmd": "Athena Master",
    "excel_col_name": "Supplier Name",
    "excel_col_iso": "Certificates: Quality ISO 9001 and others: Expiry date",
    "excel_col_env": "Certificates: Environment: Expiry date",
    "excel_col_iatf": "Certificates: Quality IATF 16949: Expiry date",
    # Usada só para filtrar: o app processa apenas fornecedores nacionais
    # (Country Code = "BR"), fornecedores de outros países são descartados
    # já na leitura do Excel.
    "excel_col_country": "Country Code",

    # Banco de dados Access com o cadastro de e-mails dos fornecedores
    "access_db_path": "",
    "access_table": "tbl_Fornecedores",
    "access_col_rbsno": "RBSNO",
    "access_col_pmd": "PMD Master",
    "access_col_email": "Emails Concatenados",

    # Histórico de envios interno (SQLite, nada a configurar)
    "history_db_path": os.path.join(os.path.dirname(os.path.abspath(__file__)), "tracking.db"),

    # Regras de notificação
    "reminder_days": [90, 60, 30],
    "overdue_resend_interval_days": 15,  # a cada quantos dias reenviar aviso de certificado vencido

    # Conteúdo usado no corpo do e-mail (o app agora só envia para
    # fornecedores nacionais, então o conteúdo é sempre em português)
    "company_name": "Bosch Rexroth LTDA.",
    # cópia oculta (BCC): o fornecedor não vê esse endereço
    "bcc_email": "Digital.Solutions@boschrexroth.com.br",
    "default_subject_pt": "Renovação de Certificado - Ação Necessária",
    # assunto do e-mail enviado quando ISO 9001 E IATF 16949 estão em branco
    "default_subject_missing_pt": "Confirmação de Certificado - Ação Necessária",

    # Remetente: endereço (ex.: caixa compartilhada "Digital Solutions") em
    # nome do qual os e-mails são enviados. Em branco = conta padrão do Outlook.
    "sender_email": "",

    # pasta onde os certificados recebidos/anexos são salvos (uma subpasta por fornecedor)
    "certificates_folder": "",
}


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = DEFAULT_CONFIG.copy()
            merged.update(cfg)
            # configs antigas guardam o caminho do banco com o nome antigo
            if os.path.basename(str(merged.get("history_db_path", ""))).lower() == "renovacert_history.db":
                merged["history_db_path"] = DEFAULT_CONFIG["history_db_path"]
            return merged
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

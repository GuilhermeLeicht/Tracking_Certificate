# -*- coding: utf-8 -*-
"""
Regras das faixas de notificação:

    - 90, 60 e 30 dias antes do vencimento -> um aviso por faixa.
    - Depois de vencido -> continua notificando periodicamente (padrão: a
      cada 7 dias, configurável) até o certificado ser renovado (ou seja,
      até a próxima leitura do Excel trazer uma data de validade futura
      para aquele fornecedor).
"""
from datetime import date, datetime

WARNING_THRESHOLDS = (90, 60, 30)
EXPIRED = "EXPIRED"
NO_DATE = "NO_DATE"


def compute_status(expiry_date, today=None):
    """
    Retorna (faixa, dias_restantes).
    faixa é None quando ainda não é hora de avisar (mais de 90 dias de validade).
    dias_restantes é negativo quando o certificado já venceu.
    """
    if today is None:
        today = date.today()
    if expiry_date is None:
        return None, None

    days_left = (expiry_date - today).days

    if days_left < 0:
        return EXPIRED, days_left

    tier = None
    for threshold in sorted(WARNING_THRESHOLDS):
        if days_left <= threshold:
            tier = str(threshold)
            break
    return tier, days_left


def needs_notification(history_db, rbsno, cert_type, expiry_date, tier, overdue_interval_days):
    """Decide, com base no histórico, se um novo e-mail deve ser disparado agora."""
    if tier is None:
        return False

    last = history_db.get_last_send(rbsno, cert_type, expiry_date)

    if tier in (EXPIRED, NO_DATE):
        if last is None:
            return True
        last_tier, last_date_str = last
        if last_tier != tier:
            return True
        last_date = datetime.strptime(last_date_str, "%Y-%m-%d").date()
        return (date.today() - last_date).days >= overdue_interval_days

    # faixas 90/60/30: envia uma vez por faixa
    if last is None:
        return True
    last_tier, _ = last
    return last_tier != tier


def tier_label(tier):
    if tier == EXPIRED:
        return "Expired"
    if tier == NO_DATE:
        return "Missing"
    return f"{tier} days"

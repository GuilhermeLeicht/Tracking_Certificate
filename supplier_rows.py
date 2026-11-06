# -*- coding: utf-8 -*-
"""
Monta, para cada fornecedor, as linhas de status dos 3 tipos de certificado
(ISO 9001, Ambiental/ISO 14001, IATF 16949), incluindo a lógica de
substituição cruzada entre ISO 9001 e IATF 16949: um fornecedor é considerado
em conformidade se tiver QUALQUER UM dos dois certificados válido (não
precisa ter os dois). A Ambiental (ISO 14001) é sempre avaliada de forma
independente, sem nenhuma relação com ISO 9001/IATF 16949.

Regras da substituição cruzada:
    - Se o fornecedor tem ISO 9001 válido (não vencido): na aba IATF 16949,
      a linha fica AZUL com status "Supplier ISO" (não precisa de IATF).
    - Se o fornecedor tem IATF 16949 válido (não vencido) mas NÃO tem ISO
      9001 válido: na aba ISO 9001, a linha fica AZUL com status
      "Supplier IATF".
    - Se NENHUM dos dois é válido: a linha fica VERMELHA (ou cinza, se faltar
      data nos dois) em ambas as abas, mas um ÚNICO e-mail combinado
      ("ISO 9001 or IATF 16949") é dividido - o envio/pendência fica
      marcado apenas na linha da aba ISO 9001, para nunca disparar dois
      e-mails separados para o mesmo problema.

Usado tanto pela interface gráfica (ui/main_window.py) quanto pela execução
automática sem interface (auto_run.py), para as duas ficarem sempre
consistentes entre si.
"""
from reminder_engine import compute_status, needs_notification, EXPIRED, NO_DATE

# rótulos em inglês (interface/tabelas) e tags/rótulos em português (conteúdo
# dos e-mails - o app só processa fornecedores nacionais, Country Code "BR")
CERT_TYPES = [
    ("iso_expiry", "ISO 9001 (Quality)"),
    ("env_expiry", "Environmental (ISO 14001)"),
    ("iatf_expiry", "IATF 16949 (Quality)"),
]
CERT_TYPE_LABELS_PT = {
    "iso_expiry": "ISO 9001 (Qualidade)",
    "env_expiry": "Ambiental (ISO 14001)",
    "iatf_expiry": "IATF 16949 (Qualidade)",
}
CERT_ISO_TAG = {
    "iso_expiry": "ISO 9001",
    "env_expiry": "ISO 14001",
    "iatf_expiry": "IATF 16949",
}

# "campo" artificial usado no histórico/envio quando nenhum dos dois
# certificados (ISO 9001 / IATF 16949) é válido - gera um único e-mail
# cobrindo as duas opções, em vez de dois e-mails separados.
COMBINED_FIELD = "combined_cert"
CERT_TYPE_LABELS_PT[COMBINED_FIELD] = "ISO 9001 ou IATF 16949"
CERT_ISO_TAG[COMBINED_FIELD] = "ISO 9001 / IATF 16949"

CERT_LABEL_EN = {
    "iso_expiry": "ISO 9001",
    "env_expiry": "Environmental (ISO 14001)",
    "iatf_expiry": "IATF 16949",
    COMBINED_FIELD: "ISO 9001 or IATF 16949",
}

# marca local (não gravada no histórico) usada na tabela para indicar que o
# fornecedor está em dia graças ao OUTRO certificado (ISO 9001 <-> IATF 16949)
BLUE = "BLUE"


def _status(expiry):
    if expiry is None:
        return NO_DATE, None
    return compute_status(expiry)


def _combined(iso_tier, iso_days, iso_expiry, iatf_tier, iatf_days, iatf_expiry):
    """Calcula a situação 'combinada' (pior caso) entre ISO 9001 e IATF
    16949 quando NENHUM dos dois é válido - usada para a linha/e-mail
    combinados ("ISO 9001 or IATF 16949")."""
    if iso_tier == EXPIRED or iatf_tier == EXPIRED:
        candidates = [(d, e) for t, d, e in
                      ((iso_tier, iso_days, iso_expiry), (iatf_tier, iatf_days, iatf_expiry))
                      if t == EXPIRED]
        days, expiry = min(candidates, key=lambda c: c[0])
        return EXPIRED, days, expiry
    return NO_DATE, None, None


def quality_valid(sup):
    """ISO 9001 ou IATF 16949 válido (não vencido, com data)."""
    for exp in (sup.iso_expiry, sup.iatf_expiry):
        if exp is not None and compute_status(exp)[0] != EXPIRED:
            return True
    return False


def env_valid(sup):
    return sup.env_expiry is not None and compute_status(sup.env_expiry)[0] != EXPIRED


def load_received(rdb, suppliers):
    """Lê os certificados 'recebidos, ainda não subidos'. Quando o Excel passa
    a mostrar o certificado como válido, o registro é encerrado (o dado já
    chegou ao datalake) e os lembretes voltam a valer normalmente."""
    received = rdb.suppressed_groups()
    for sup in suppliers:
        if (sup.rbsno, "quality") in received and quality_valid(sup):
            rdb.delete_group(sup.rbsno, "quality")
            received.discard((sup.rbsno, "quality"))
        if (sup.rbsno, "env") in received and env_valid(sup):
            rdb.delete_group(sup.rbsno, "env")
            received.discard((sup.rbsno, "env"))
    return received


def build_rows(sup, email, hdb, overdue_days, received=frozenset(), risk=frozenset()):
    """Retorna uma lista com 3 dicionários de linha (aba "iso_expiry",
    "env_expiry" e "iatf_expiry"), já com a lógica de substituição cruzada e
    de pendência de envio aplicada.

    Cada dicionário tem as chaves:
        tab        - em qual aba da tabela a linha aparece
        field      - "campo" usado para histórico/envio (pode ser
                     "combined_cert" quando nenhum dos dois certs é válido)
        kind       - "normal", "blue_iso" (ISO válido cobrindo a aba IATF) ou
                     "blue_iatf" (IATF válido cobrindo a aba ISO)
        expiry, tier, days_left, pending, label, rbsno, pmd, name, email
    """
    rows = []

    # --- Ambiental (ISO 14001): totalmente independente --------------------
    env_tier, env_days = _status(sup.env_expiry)
    env_pending = False
    if sup.env_expiry is not None and env_tier is not None:
        env_pending = needs_notification(hdb, sup.rbsno, "env_expiry", sup.env_expiry, env_tier, overdue_days)
    rows.append({
        "tab": "env_expiry", "field": "env_expiry", "kind": "normal",
        "expiry": sup.env_expiry, "tier": env_tier, "days_left": env_days,
        "pending": env_pending, "label": CERT_LABEL_EN["env_expiry"],
    })

    # --- ISO 9001 x IATF 16949: substituição cruzada ------------------------
    iso_tier, iso_days = _status(sup.iso_expiry)
    iatf_tier, iatf_days = _status(sup.iatf_expiry)
    iso_valid = sup.iso_expiry is not None and iso_tier != EXPIRED
    iatf_valid = sup.iatf_expiry is not None and iatf_tier != EXPIRED

    # um certificado ATIVO (mais de 90 dias) repõe o outro: se o outro está
    # expirando/vencido/em branco, a linha fica azul e conta como Ativo
    if iso_tier is None or (iso_valid and iatf_tier is not None):
        iso_pending = False
        if iso_tier is not None:
            iso_pending = needs_notification(hdb, sup.rbsno, "iso_expiry", sup.iso_expiry, iso_tier, overdue_days)
        row_iso = {
            "tab": "iso_expiry", "field": "iso_expiry", "kind": "normal",
            "expiry": sup.iso_expiry, "tier": iso_tier, "days_left": iso_days,
            "pending": iso_pending, "label": CERT_LABEL_EN["iso_expiry"],
        }
    elif iatf_valid:
        row_iso = {
            "tab": "iso_expiry", "field": "iso_expiry", "kind": "blue_iatf",
            "expiry": sup.iatf_expiry, "tier": BLUE, "days_left": iatf_days,
            "pending": False, "label": CERT_LABEL_EN["iso_expiry"],
        }
    else:
        c_tier, c_days, c_expiry = _combined(iso_tier, iso_days, sup.iso_expiry,
                                              iatf_tier, iatf_days, sup.iatf_expiry)
        c_pending = False
        if c_tier in (EXPIRED, NO_DATE):  # NO_DATE: ISO 9001 e IATF em branco
            c_pending = needs_notification(hdb, sup.rbsno, COMBINED_FIELD, c_expiry, c_tier, overdue_days)
        row_iso = {
            "tab": "iso_expiry", "field": COMBINED_FIELD, "kind": "normal",
            "expiry": c_expiry, "tier": c_tier, "days_left": c_days,
            "pending": c_pending, "label": CERT_LABEL_EN[COMBINED_FIELD],
        }

    if iatf_tier is None or (iatf_valid and iso_tier is not None):
        iatf_pending = False
        if iatf_tier is not None:
            iatf_pending = needs_notification(hdb, sup.rbsno, "iatf_expiry", sup.iatf_expiry, iatf_tier, overdue_days)
        row_iatf = {
            "tab": "iatf_expiry", "field": "iatf_expiry", "kind": "normal",
            "expiry": sup.iatf_expiry, "tier": iatf_tier, "days_left": iatf_days,
            "pending": iatf_pending, "label": CERT_LABEL_EN["iatf_expiry"],
        }
    elif iso_valid:
        row_iatf = {
            "tab": "iatf_expiry", "field": "iatf_expiry", "kind": "blue_iso",
            "expiry": sup.iso_expiry, "tier": BLUE, "days_left": iso_days,
            "pending": False, "label": CERT_LABEL_EN["iatf_expiry"],
        }
    else:
        c_tier, c_days, c_expiry = _combined(iso_tier, iso_days, sup.iso_expiry,
                                              iatf_tier, iatf_days, sup.iatf_expiry)
        # mesma situação "combinada" calculada acima, mas NUNCA pendente
        # aqui: o único e-mail (cobrindo os dois certificados) já é
        # disparado a partir da linha da aba ISO 9001, acima.
        row_iatf = {
            "tab": "iatf_expiry", "field": COMBINED_FIELD, "kind": "normal",
            "expiry": c_expiry, "tier": c_tier, "days_left": c_days,
            "pending": False, "label": CERT_LABEL_EN[COMBINED_FIELD],
            "shown_pending": row_iso["pending"],  # só exibição; não conta p/ envio
        }

    rows.append(row_iso)
    rows.append(row_iatf)
    for r in rows:
        # ISO e IATF ambos em branco: o e-mail pergunta se o fornecedor possui
        # o certificado (a Ambiental em branco NUNCA gera e-mail)
        r["notify"] = (r["tier"] in (EXPIRED, NO_DATE)) if r["field"] == COMBINED_FIELD else (r["tier"] not in (None, NO_DATE, BLUE))
        grp = "env" if r["tab"] == "env_expiry" else "quality"
        if r["notify"] and (sup.rbsno, grp) in received:
            # certificado já recebido (ainda não subido): suspende a cobrança
            r["received"] = True
            r["pending"] = False
            r["shown_pending"] = False
        elif r["notify"] and grp == "quality" and (sup.rbsno in risk or (sup.name or "").lower() in risk):
            # fornecedor já em Risk Assessment: não cobrar mais por e-mail
            r["risk"] = True
            r["pending"] = False
            r["shown_pending"] = False
    for row in rows:
        row.update({"rbsno": sup.rbsno, "pmd": sup.pmd, "name": sup.name, "email": email})
    return rows

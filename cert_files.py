# -*- coding: utf-8 -*-
"""Gravação dos arquivos de certificados em pastas por fornecedor:
<pasta base>/<Nome do fornecedor> - <RBSNO>/"""
import os
import re
import shutil
from datetime import datetime


def _clean(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (name or "").strip()).rstrip(". ") or "Supplier"


def supplier_folder(base, name, rbsno, create=True):
    path = os.path.join(base, f"{_clean(name)} - {_clean(rbsno)}")
    if create:
        os.makedirs(path, exist_ok=True)
    return path


# prefixo do nome do arquivo (e grupo) por tipo de certificado
def file_label(certificate):
    c = certificate or ""
    if "IATF" in c:
        return "IATF 16949"
    if "14001" in c or "Environ" in c:
        return "ISO 14001"
    return "ISO 9001"


def scan_keys(base, name, rbsno):
    """Lê a pasta do fornecedor e devolve os certificados ("iso", "iatf",
    "env") que têm arquivo cujo nome começa com 'ISO 9001 - ',
    'IATF 16949 - ' ou 'ISO 14001 - '."""
    keys = set()
    try:
        folder = supplier_folder(base, name, rbsno, create=False)
        for fn in os.listdir(folder):
            if fn.startswith("ISO 9001 - "):
                keys.add("iso")
            elif fn.startswith("IATF 16949 - "):
                keys.add("iatf")
            elif fn.startswith("ISO 14001 - "):
                keys.add("env")
    except OSError:
        pass
    return keys


def save_file(base, name, rbsno, src, certificate=""):
    """Copia `src` para a pasta do fornecedor (criada se não existir),
    renomeando para '<Certificado> - <Fornecedor> - <RBSNO>.<ext>', e devolve
    o caminho final. Se já existir um arquivo com o mesmo nome, ele é substituído."""
    folder = supplier_folder(base, name, rbsno)
    ext = os.path.splitext(src)[1]
    stem = f"{file_label(certificate)} - {_clean(name)} - {_clean(rbsno)}"
    dest = os.path.join(folder, stem + ext)
    # mesmo nome = substitui o arquivo anterior (copy2 sobrescreve)
    shutil.copy2(src, dest)
    return dest


def save_extra(base, name, rbsno, src, subfolder=""):
    """Copia um arquivo (certificado do Risk Assessment ou evidência) mantendo
    o nome original, sem nunca sobrescrever (acrescenta data/hora se já
    existir), para preservar o histórico. `subfolder` ex.: "Evidence"."""
    folder = supplier_folder(base, name, rbsno)
    if subfolder:
        folder = os.path.join(folder, subfolder)
        os.makedirs(folder, exist_ok=True)
    dest = os.path.join(folder, os.path.basename(src))
    if os.path.exists(dest):
        stem, ext = os.path.splitext(os.path.basename(src))
        dest = os.path.join(folder, f"{stem}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
    shutil.copy2(src, dest)
    return dest


def standardize_existing(path, name, rbsno, certificate):
    """Renomeia um arquivo já existente na pasta do fornecedor para o padrão
    '<Certificado> - <Fornecedor> - <RBSNO>.<ext>' (sobrescreve o de mesmo
    nome). Devolve o novo caminho (ou o original se não for preciso/possível)."""
    if not os.path.isfile(path):
        return path
    label = file_label(certificate)
    stem = f"{label} - {_clean(name)} - {_clean(rbsno)}"
    dest = os.path.join(os.path.dirname(path), stem + os.path.splitext(path)[1])
    if os.path.abspath(dest) == os.path.abspath(path):
        return path
    os.replace(path, dest)
    return dest

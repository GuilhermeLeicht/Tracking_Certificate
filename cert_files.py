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


def save_file(base, name, rbsno, src):
    """Copia `src` para a pasta do fornecedor (criada se não existir) e
    devolve o caminho final. Nunca sobrescreve: se já existir, acrescenta data/hora."""
    folder = supplier_folder(base, name, rbsno)
    dest = os.path.join(folder, os.path.basename(src))
    if os.path.exists(dest):
        stem, ext = os.path.splitext(os.path.basename(src))
        dest = os.path.join(folder, f"{stem}_{datetime.now():%Y%m%d_%H%M%S}{ext}")
    shutil.copy2(src, dest)
    return dest

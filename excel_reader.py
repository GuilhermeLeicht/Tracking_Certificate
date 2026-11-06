# -*- coding: utf-8 -*-
"""
Leitura do arquivo .xlsx exportado do SRM.

Os nomes das colunas são configuráveis (aba Settings). Os valores abaixo são
apenas os padrões iniciais.

O arquivo pode ter várias abas (ex.: uma aba de capa/metadados e outra com os
dados dos fornecedores) - o app procura automaticamente em todas as abas até
encontrar a que contém as colunas configuradas.
"""
from datetime import datetime, date
import openpyxl

DEFAULT_COLUMNS = {
    "rbsno": "Basic Data: RBSNO",
    "pmd": "PMD Master",
    "name": "Supplier Name",
    "iso_expiry": "Certificates: Quality ISO 9001 and others: Expiry date",
    "env_expiry": "Certificates: Environment: Expiry date",
    "iatf_expiry": "Certificates: Quality IATF 16949: Expiry date",
    "country_code": "Country Code",
}

# nomes amigáveis usados nas mensagens de erro (mensagens ficam em inglês,
# igual ao restante do app - só os comentários/docstrings estão em português)
FIELD_LABELS = {
    "rbsno": "Supplier code (RBSNO)",
    "pmd": "Supplier code (PMD/SAP)",
    "name": "Supplier name",
    "iso_expiry": "ISO 9001 certificate expiry date",
    "env_expiry": "Environmental certificate expiry date",
    "iatf_expiry": "IATF 16949 certificate expiry date",
    "country_code": "Supplier country code (Country Code)",
}


class SupplierRecord:
    __slots__ = ("rbsno", "pmd", "name", "iso_expiry", "env_expiry", "iatf_expiry", "country_code")

    def __init__(self, rbsno, pmd, name, iso_expiry, env_expiry, iatf_expiry, country_code):
        self.rbsno = rbsno
        self.pmd = pmd
        self.name = name
        self.iso_expiry = iso_expiry
        self.env_expiry = env_expiry
        self.iatf_expiry = iatf_expiry
        self.country_code = country_code


def _parse_date(value):
    """Aceita datetime/date já convertido pelo Excel, ou texto dd/mm/aaaa."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _norm(s):
    """Normaliza um texto de cabeçalho para comparação: sem espaços
    duplicados/nas pontas, sem ':' no final, sem diferenciar maiúsc/minúsc."""
    text = " ".join(str(s).split()).strip().rstrip(":").strip()
    return text.lower()


def _find_header_row(ws, columns, max_scan_rows=40):
    """Procura, nas primeiras linhas de UMA aba, a linha que mais se parece
    com o cabeçalho configurado. Retorna (numero_da_linha, mapa_campo->coluna,
    lista_de_cabecalhos_brutos_dessa_linha)."""
    best_row = None
    best_map = {}
    best_raw = []

    # linha com mais células de texto preenchidas, usada como preview mesmo
    # quando nenhuma coluna bate com o que está configurado
    most_filled_row = None
    most_filled_raw = []

    for row in ws.iter_rows(min_row=1, max_row=max_scan_rows):
        row_map = {}
        raw = []
        for cell in row:
            if cell.value is None:
                continue
            raw.append(str(cell.value).strip())
            key = _norm(cell.value)
            for field, col_name in columns.items():
                if key == _norm(col_name):
                    row_map[field] = cell.column

        if len(raw) > len(most_filled_raw):
            most_filled_raw = raw
            most_filled_row = row[0].row

        if len(row_map) > len(best_map):
            best_map = row_map
            best_row = row[0].row
            best_raw = raw
        if len(row_map) == len(columns):
            return row[0].row, row_map, raw

    if not best_raw:
        best_row = most_filled_row
        best_raw = most_filled_raw

    return best_row, best_map, best_raw


def _extract_suppliers(ws, header_row_idx, col_map):
    suppliers = []
    max_col = max(col_map.values())
    for row in ws.iter_rows(min_row=header_row_idx + 1):
        if max_col > len(row):
            continue
        rbsno_val = row[col_map["rbsno"] - 1].value
        if rbsno_val is None or str(rbsno_val).strip() == "":
            continue

        pmd_val = row[col_map["pmd"] - 1].value
        name_val = row[col_map["name"] - 1].value
        iso_val = row[col_map["iso_expiry"] - 1].value
        env_val = row[col_map["env_expiry"] - 1].value
        iatf_val = row[col_map["iatf_expiry"] - 1].value
        country_val = row[col_map["country_code"] - 1].value

        suppliers.append(SupplierRecord(
            rbsno=str(rbsno_val).strip(),
            pmd=str(pmd_val).strip() if pmd_val is not None else "",
            name=str(name_val).strip() if name_val is not None else "",
            iso_expiry=_parse_date(iso_val),
            env_expiry=_parse_date(env_val),
            iatf_expiry=_parse_date(iatf_val),
            country_code=str(country_val).strip() if country_val is not None else "",
        ))
    return suppliers


def read_suppliers(path, columns=None, sheet_name=None, country_filter=None):
    """Lê o Excel e devolve uma lista de SupplierRecord.

    `columns` é um dicionário {campo: nome_da_coluna_no_excel}; se omitido,
    usa DEFAULT_COLUMNS. Os campos são fixos: rbsno, pmd, name, iso_expiry,
    env_expiry, country_code.

    `sheet_name`, se informado, restringe a busca a essa aba específica.
    Se omitido/vazio, procura em TODAS as abas do arquivo (algumas
    exportações do SRM trazem uma aba de capa/metadados antes da aba com os
    dados dos fornecedores).

    `country_filter`, se informado (ex.: "BR"), mantém apenas os fornecedores
    cujo Country Code bate exatamente com esse valor (sem diferenciar
    maiúsc/minúsc). Usado para restringir o app a fornecedores nacionais.
    """
    columns = columns or DEFAULT_COLUMNS

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet_names = list(wb.sheetnames)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            wb.close()
            raise ValueError(
                f'The sheet "{sheet_name}" (configured in Settings > "SRM Excel Columns") '
                f"does not exist in this file.\nAvailable sheets: {', '.join(sheet_names)}"
            )
        sheets_to_search = [wb[sheet_name]]
    else:
        if not wb.worksheets:
            wb.close()
            raise ValueError("The Excel file has no sheets.")
        sheets_to_search = wb.worksheets

    best_sheet = None  # (sheet_name, row, col_map, match_count, raw)

    for ws in sheets_to_search:
        header_row_idx, col_map, raw = _find_header_row(ws, columns)
        if header_row_idx is None:
            continue

        if len(col_map) == len(columns):
            suppliers = _extract_suppliers(ws, header_row_idx, col_map)
            wb.close()
            if country_filter:
                target = country_filter.strip().upper()
                suppliers = [s for s in suppliers if (s.country_code or "").strip().upper() == target]
            return suppliers

        if best_sheet is None or len(col_map) > best_sheet[3]:
            best_sheet = (ws.title, header_row_idx, col_map, len(col_map), raw)

    wb.close()

    if best_sheet is None:
        raise ValueError(
            "No recognizable header was found in the file.\n"
            f"Sheets checked: {', '.join(sheet_names)}\n\n"
            "Adjust the column names in Settings > \"SRM Excel Columns\" "
            "to match your file exactly."
        )

    sheet_title, header_row_idx, col_map, _count, raw = best_sheet
    missing = [FIELD_LABELS[f] + f'  (configured as: "{columns[f]}")'
               for f in columns if f not in col_map]
    preview = ", ".join(raw[:60]) if raw else "(no header identified)"

    raise ValueError(
        f'On the closest matching sheet ("{sheet_title}", row {header_row_idx}), '
        "these columns were not found:\n"
        "- " + "\n- ".join(missing) +
        "\n\nHeaders found on that row:\n" + preview +
        f"\n\nSheets checked: {', '.join(sheet_names)}" +
        "\n\nAdjust the column names in Settings > \"SRM Excel Columns\" "
        "to match your file exactly."
    )

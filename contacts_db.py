# -*- coding: utf-8 -*-
"""
Busca o e-mail de cada fornecedor num banco Access (.accdb/.mdb), o mesmo
usado pelo EvoSend. Requer o driver "Microsoft Access Driver" (normalmente já
vem com o Office no Windows) e o pacote pyodbc.
"""
import pyodbc


def _list_table_columns(cursor, table):
    """Lista os nomes reais das colunas de uma tabela/consulta do Access,
    usado para dar uma mensagem de erro mais clara quando um nome de coluna
    configurado não existe."""
    try:
        return [row.column_name for row in cursor.columns(table=table)]
    except Exception:
        return []


def load_emails(access_db_path, table, col_rbsno, col_pmd, col_email):
    """
    Retorna um dicionário {("rbsno", codigo): email, ("pmd", codigo): email}
    para permitir localizar o e-mail tanto por RBSNO quanto por PMD.
    """
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={access_db_path};"
    )
    emails = {}
    conn = pyodbc.connect(conn_str)
    try:
        cur = conn.cursor()
        try:
            cur.execute(f"SELECT [{col_rbsno}], [{col_pmd}], [{col_email}] FROM [{table}]")
        except pyodbc.Error as exc:
            # O erro cru do driver ODBC do Access ("Too few parameters...")
            # normalmente significa que um dos nomes de coluna configurados
            # não existe de fato na tabela/consulta - o Access interpreta o
            # nome desconhecido como um parâmetro esperado. Damos uma
            # mensagem mais útil, listando as colunas reais encontradas.
            available = _list_table_columns(cur, table)
            hint = ""
            if available:
                hint = f'\nColumns actually found in "{table}": {", ".join(available)}'
            else:
                hint = (f'\nCould not list the columns of "{table}" either - '
                        "double-check the table/query name itself.")
            raise ValueError(
                f'Could not query columns "{col_rbsno}", "{col_pmd}", "{col_email}" '
                f'in "{table}". This usually means one of those names does not match '
                f"exactly what exists in the Access database.{hint}\n\n"
                'Check Settings > "Supplier E-mail Directory (Access)" for typos in '
                "the table or column names."
            ) from exc

        for row in cur.fetchall():
            rbsno = str(row[0]).strip() if row[0] is not None else ""
            pmd = str(row[1]).strip() if row[1] is not None else ""
            email = str(row[2]).strip() if row[2] is not None else ""
            if not email:
                continue
            if rbsno:
                emails[("rbsno", rbsno)] = email
            if pmd:
                emails[("pmd", pmd)] = email
    finally:
        conn.close()
    return emails


def resolve_email(emails, rbsno, pmd):
    if rbsno and ("rbsno", rbsno) in emails:
        return emails[("rbsno", rbsno)]
    if pmd and ("pmd", pmd) in emails:
        return emails[("pmd", pmd)]
    return None

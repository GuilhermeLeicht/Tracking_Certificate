# -*- coding: utf-8 -*-
"""
Envio de e-mail via automação do Outlook (Windows).
Requer o Outlook instalado e configurado na máquina, e o pacote pywin32.
"""


def send_email(recipient, bcc, subject, html_body, sender=""):
    import win32com.client as win32  # import tardio: só existe no Windows

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = recipient
    if bcc:
        mail.BCC = bcc  # cópia oculta: o fornecedor não vê
    if sender:
        # 1) se o endereço é uma conta configurada neste Outlook, envia por ela
        account = None
        try:
            for acc in outlook.Session.Accounts:
                if (acc.SmtpAddress or "").strip().lower() == sender.strip().lower():
                    account = acc
                    break
        except Exception:
            account = None
        if account is not None:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # SendUsingAccount
        else:
            # 2) caixa compartilhada / "enviar como" / "enviar em nome de":
            # exige a permissão correspondente no Exchange para o seu usuário
            mail.SentOnBehalfOfName = sender
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()

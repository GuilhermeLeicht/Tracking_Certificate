# -*- coding: utf-8 -*-
"""
Notificações nativas do sistema (via Plyer). Plyer é opcional: se não estiver
instalado ou a plataforma não tiver backend, nada quebra - o motivo exato da
falha é gravado no Log (parâmetro `log`).
"""


def notify(title, message, log=None):
    try:
        from plyer import notification
        notification.notify(title=title, message=message, app_name="Tracking Certificate", timeout=8)
        if log:
            log(f"Desktop notification sent: {message}")
    except Exception as exc:
        if log:
            log(f"Desktop notification NOT shown ({type(exc).__name__}: {exc}). "
                "Check that 'plyer' is installed (pip install plyer).")

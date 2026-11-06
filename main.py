# -*- coding: utf-8 -*-
"""
Tracking Certificate - Ponto de entrada.

Lê a exportação do Excel do SRM, calcula quais fornecedores estão a 90/60/30
dias do vencimento dos certificados (ISO 9001 e Ambiental) ou já vencidos, e
envia e-mails de aviso automaticamente via Outlook (em português para
fornecedores brasileiros, em inglês para os demais).

Execução normal (abre a interface gráfica):
    python main.py

Execução automática (sem interface, feita para rodar sozinha via Agendador
de Tarefas do Windows - veja README.md):
    python main.py --auto
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    if "--auto" in sys.argv or "--auto-send" in sys.argv:
        from auto_run import run_automatic
        sys.exit(run_automatic())
    else:
        from PySide6.QtWidgets import QApplication
        from ui.main_window import MainWindow

        app = QApplication(sys.argv)
        app.setApplicationName("Tracking Certificate")
        window = MainWindow()
        window.show()
        sys.exit(app.exec())


if __name__ == "__main__":
    main()

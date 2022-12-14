# -*- coding: utf-8 -*-
"""
Created on Tue Sep 27 13:30:56 2022.

@author: Ítalo Ferreira Fernandes

Aplicativo com UI que pega as informações passada pelo usuário atravez da interface
e executa o arquivo Python download_files_outlook com essas informações.

"""

from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QDateTime
from download_files_outlook import main as download_files_outlook
import sys
import os


class Window(QtWidgets.QMainWindow):
    """Janela principal."""

    files_format = {
        'Excel Workbook (.xlsx)': r'.*\.xlsx',
        'Excel 97- Excel 2003 Workbook (.xls)': r'.*\.xls',
        'CSV File (.csv)': r'.*\.csv',
        'Text File (.txt)': r'.*\.txt',
        'PDF (.pdf)': r'.*\.pdf'}

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        uic.loadUi("baixar_anexo_outlook.ui", self)

        # Variáveis auxiliares
        agora = QDateTime.currentDateTime()
        ontem = agora.addDays(-1)
        folders = get_folders_outlook()

        # Valores fixos
        self.dt_fim.setDateTime(agora)  # Antes de
        self.dt_ini.setDateTime(ontem)  # Depois de
        self.box_format.addItems(list(self.files_format.keys()))  # Formato do Arquivo
        self.box_folders.addItems(['Caixa de Entrada'] + folders)  # Pasta do Outlook
        self.text_path.setText(f"{os.environ['USERPROFILE']}\\Downloads")  # Pasta para Salvar

        # Ativar e Desativar filtros
        self.check_from.clicked.connect(lambda: self.change_checked(self.check_from, self.text_from))
        self.check_subject.clicked.connect(lambda: self.change_checked(self.check_subject, self.text_subject))
        self.check_body.clicked.connect(lambda: self.change_checked(self.check_body, self.text_body))
        self.check_dt_ini.clicked.connect(lambda: self.change_checked(self.check_dt_ini, self.dt_ini))
        self.check_dt_fim.clicked.connect(lambda: self.change_checked(self.check_dt_fim, self.dt_fim))

        # Botão para procurar pasta
        self.btn_path.clicked.connect(self.browse_path)

        # Botão para baixar arquivos
        self.btn_main.clicked.connect(self.concluir)

    def change_checked(self, check_obj, text_obj):
        """Se tiver marcado deixa como ativo o campo. Se desmarcar, desativa o filtro."""
        if check_obj.isChecked():
            text_obj.setEnabled(True)
        else:
            text_obj.setEnabled(False)

    def browse_path(self):
        """Abre um Dialog para procurar uma pasta onde salvara os arquivos."""
        filepath = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Folder')
        if filepath:
            path = filepath.replace('/', '\\')
            self.text_path.setText(path)

    def concluir(self):
        """Conclui o aplicativo: baixa os arquivos do Outlook."""
        dados = {
            'subject': self.return_text(self.text_subject),
            'body': self.return_text(self.text_body),
            'dt_ini': self.return_text(self.dt_ini),
            'dt_fim': self.return_text(self.dt_fim),
            'format_info': self.return_text(self.box_format, box=True),
            'path': self.return_text(self.text_path),
            'folder': self.return_text(self.box_folders, box=True),
            'from': self.return_text(self.text_from)
        }
        dados['format'] = self.files_format[dados['format_info']]

        try:
            text, info = download_files_outlook(dados)
            showdialog(text, info, icon=QMessageBox.Information)
        except Exception as e:
            showdialog(type(e).__name__, '\n'.join(e.args), icon=QMessageBox.Critical)

    def return_text(self, text_obj, box=False):
        """Retorna o texto do Input se ele estiver Ativo, se não retorna False."""
        if box:
            return text_obj.currentText()
        elif text_obj.isEnabled():
            return text_obj.text()
        else:
            return False


def showdialog(text, info, icon=QMessageBox.Information):
    """Mostra uma Message Box como o Texto e a Info."""
    msg = QMessageBox()
    msg.setIcon(icon)
    msg.setText(text)
    msg.setWindowTitle("Mensagem")
    msg.setDetailedText(info)
    msg.setStandardButtons(QMessageBox.Ok)
    # msg.setStyleSheet("QLabel{min-width: 170px;}")
    msg.exec_()


def get_folders_outlook():
    """Retorna as pastas que tem dentro do Inbox no Outlook."""
    import win32com.client
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6)
    folders = [folder.Name for folder in inbox.Folders]
    return folders


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = Window()
    w.show()
    sys.exit(app.exec_())

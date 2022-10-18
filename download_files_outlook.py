# -*- coding: utf-8 -*-
"""
Created on Tue Sep 27 17:25:11 2022.

@author: Ítalo Ferreira Fernandes

Código para baixar anexos de email do outlook.

Recebe um dicionario que tem as seguintes informações:
    'from',
    'body',
    'subject',
    'dt_ini',
    'dt_fim',
    'format_info',
    'path',
    'format',
    'folder'

"""

import win32com.client
import os
import re
from datetime import datetime

def verificacoes(dados):
    """Verifica se todos os dados estão no formato correto."""
    # Dado em branco
    assert all([str(s).strip() != '' for s in dados.values()]), "Não pode ter nenhum campo em branco ativo."
    # Data mal formatada
    try:
        if dados['dt_ini']:
            datetime.strptime(dados['dt_ini'], '%d/%m/%Y %H:%M')
        if dados['dt_fim']:
            datetime.strptime(dados['dt_fim'], '%d/%m/%Y %H:%M')
    except Exception as e:
        raise eval(f'{type(e).__name__}("Não foi possivel converter a data passada em data.")')
    # Caminho existe
    assert os.path.exists(dados['path']), 'O Caminho passado não existe.'
    # Keys do Dicionario
    assert set(dados.keys()) == set(['from', 'body', 'subject', 'dt_ini', 'dt_fim', 'format_info', 'path', 'format', 'folder']),\
        "Está faltando dados no dicionário."


def get_messages(dados):
    """Pegas os emails do outlook e filtra de acordo com os dados."""
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    if dados['folder'] == 'Caixa de Entrada':
        inbox = mapi.GetDefaultFolder(6)
    else:
        inbox = mapi.GetDefaultFolder(6).Folders[dados['folder']]  # Tentar colocar como variavel
    messages = inbox.Items

    # Filtros de Email
    if dados['subject']:
        messages = messages.Restrict(f"@SQL=\"urn:schemas:httpmail:subject\" like '%{dados['subject']}%'")
    if dados['body']:
        messages = messages.Restrict(f"@SQL=\"urn:schemas:httpmail:textdescription\" like '%{dados['body']}%'")
    if dados['dt_ini']:
        messages = messages.Restrict(f"[ReceivedTime] >= '{dados['dt_ini']}'")
    if dados['dt_fim']:
        messages = messages.Restrict(f"[ReceivedTime] <= '{dados['dt_fim']}'")
    if dados['from']:
        messages = messages.Restrict(f"@SQL=\"urn:schemas:httpmail:fromemail\" like '%{dados['from']}%'")
    return messages


def baixar_anexos(messages, dados):
    """Baixa os anexos e retorna a quantidade de anexos baixados."""
    n_baixados = 0
    arquivos_nome = ""
    n_error = 0
    try:
        for message in list(messages):
            try:
                for attachment in message.Attachments:
                    filename = attachment.FileName
                    if re.match(dados['format'], filename):
                        attachment.SaveASFile(os.path.join(dados['path'], attachment.FileName))
                        n_baixados += 1
                        arquivos_nome += f"{filename}\n"
            except Exception:
                n_error += 1
    except Exception as e:
        raise eval(f'{type(e).__name__}("Erro ao processar os emails: {str(e)}.")')

    return (f"{n_baixados} arquivos baixados com sucesso"+" "*20,
            f"Tiveram {n_error} erros ao baixar arquivo.\n\n"
            f"Arquivos baixados em {dados['path']}:\n\n{arquivos_nome}")


def main(dados):
    """Verifica os dados, pegas os emails e baixa os anexos."""
    verificacoes(dados)
    messages = get_messages(dados)
    out = baixar_anexos(messages, dados)
    return out

# dados = {'subject': 'Base Implantação', 'body': False, 'dt_ini': '26/09/2022 17:21', 'dt_fim': '27/09/2022 17:21',
#           'format_info': 'Excel Workbook (.xlsx)', 'path': 'C:\\Users\\ab1416392\\Downloads',
#           'format': '.*\\.xlsx', 'folder': 'Caixa de Entrada'}

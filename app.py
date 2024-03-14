"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES GOSTARIA DE SABER VALORES, E GOSTARIA QUE ENTRASSEM EM CONTATO COMIGO P/ EXPLICAR MELHOR, QUERO PODER MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM VENCIMENTO DIFERENTE
"""

import openpyxl
from urllib.parse import quote
import webbrowser

# Ler planilha e ler nome, telefone e data de vencimento

workbook = openpyxl.load_workbook("clientes.xlsx")
pagina_clientes = workbook["Sheet1"]


for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone e data de vencimento

    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    pix = '81996257747'

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar via pix {pix}'

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'


    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha

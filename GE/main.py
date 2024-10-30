import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# 1. Configuração do Servidor SMTP (Outlook)
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
EMAIL_USER = "gabriel.dias@ilclog.com.br"
EMAIL_PASS = "W1nd@ws2011"

TW_EMAILS = ['nicoli.oliveira@twl.com.br', 'thalia.drum@twtransportes.com.br']

ARMANI_EMAILS = ['marcello.almeida@armanitransportes.com.br']

JSL_EMAILS = ['rayna.lima@jsl.com.br', 'wellington.siqueira@jsl.com.br']

CC = ['plannersILC@ilclog.com.br',
      'regiane.zanetti@ilclog.com.br', 'coletas@ilclog.com.br']

VF_EMAILS = [
    'bruno.dias@vfexpress.com.br',
    'operacionalcpq@vfexpress.com.br',
    'rafael.almeida@vfexpress.com.br',
    'relacionamento@vfexpress.com.br',
    'OPERACIONAL@VFEXPRESS.COM.BR'
]

MIRASSOL_LOUVERIA_UBERABA = [
    'jussara.silva@expressomirassol.com.br',
    'adilson.marino@expressomirassol.com.br',
    'liliane.batista@expressomirassol.com.br',
    'daliane.reis@grupomirassol.com.br',
    'suellen.nicola@expressomirassol.com.br'
]

MIRASSOL_CATALAO_GRAVATAI_UBERABA_FTL = [
    'magnus.mello@expressomirassol.com.br',
    'tamires.carvalho@expressomirassol.com.br',
    'marcelo.melo@expressomirassol.com.br'
]

# Substitui todas a janelas com 00:00 para 23:59


def substituir_hora(x):
    if x.hour == 0:
        return x.replace(hour=23, minute=59)
    else:
        return x

# 2. Função para selecionar o arquivo Excel


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    return caminho_arquivo


# 3. Leitura do Excel e filtragem das colunas desejadas


def ler_dados_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)

        # Criando a nova coluna STATUS GE
        status_ge = []

        # ignorando todas as linhas que tem o status "Finalizado"
        # lembrando que ~ é para inverter a seleção de TRUE ou FALSE
        df = df[~df['MENSAGEM GE'].str.contains(
            "Finalizado", case=False, na=False)]

        df = df[~df['TIPO_TRANSP_FORN'].str.contains(
            "CIF", case=False, na=False)]

        for index, row in df.iterrows():
            mensagem_ge = row['MENSAGEM GE']
            qtd_nf = row['QTDE NF PEDIDO']

            # Verifico se o campo não esta vazio e contém "Iniciado"
            if pd.notna(mensagem_ge) and "Iniciado" in mensagem_ge:

                # Verifico se o campo não está vazio e contém NF
                if not pd.notna(qtd_nf):
                    status_ge.append("Sem inclusão de NF")

                # Esta iniciado e tem NF porém ainda nao foi finalizado
                # incluir pendente finalizacao quando tem outra coisa escrito tbm alem de iniciado
                else:
                    status_ge.append("Pendente Finalização")
            else:
                status_ge.append("Confirmar Coleta")

        # Adicionando a nova coluna ao DataFrame
        df['STATUS GE'] = status_ge

        # Verifica se existe alguma entrada não nula na coluna 'CODIGO COLETA TRANSMISSAO'
        if df['CODIGO COLETA TRANSMISSAO'].notnull().any():
            # Se existir, preenche a coluna normalmente
            df['CODIGO COLETA TRANSMISSAO'] = df['CODIGO COLETA TRANSMISSAO'].fillna(
                'Pedido não Transmitido')
        else:
            # Se não existir nenhuma entrada, preenche com 'Pedido não Transmitido'
            df['CODIGO COLETA TRANSMISSAO'] = 'Pedido não Transmitido'

        agora = datetime.now()

        df['JANELA'] = pd.to_datetime(df['JANELA'], errors='coerce')

        df['JANELA'] = df['JANELA'].apply(substituir_hora)

        # Filtrando apenas linhas com janelas futuras
        df = df[df['JANELA'] > agora]

        # Reordenando as colunas para colocar STATUS GE no início
        colunas_desejadas = [
            "STATUS GE", "PEDIDO", "JANELA", "FORNECEDOR",
            "PLANTA", "TIPO", "AGLUTINADOR", "VEICULO AGLUTINADO",
            "TRANSPORTADOR", "CODIGO COLETA TRANSMISSAO", 'MANIFESTO'
        ]
        return df[colunas_desejadas]
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

# 4. Gerar tabela HTML apenas com as colunas selecionadas


def gerar_tabela_html(df):
    html = """
    <table style="border-collapse: collapse; width: 100%;">
      <thead>
        <tr>
    """
    # Cabeçalhos das colunas
    for coluna in df.columns:
        html += f'<th style="border: 1.2px solid #ddd; padding: 8px; text-align: center; font-size: 12px; background-color: #c9c7c7;">{coluna}</th>'
    html += "</tr></thead><tbody>"

    # Dados das linhas
    for i, linha in enumerate(df.iterrows()):
        html += "<tr>"
        for valor in linha[1]:
            # Aplica cor de fundo alternada para linhas pares
            background_color = '#ffffff' if i % 2 == 0 else '#f2f2f2'
            html += f'<td style="border: 1.2px solid #ddd; padding: 8px; text-align: center; font-size: 9px; background-color: {background_color};">{valor}</td>'
        html += "</tr>"
    html += "</tbody></table>"
    return html


# 5. Envio do e-mail com a tabela para cada transportador


def enviar_email_por_transportador(df):
    hoje = datetime.today().strftime('%d/%m')

    def extrair_transportador(nome):
        if 'MIRASSOL' in nome:
            return nome.strip()

        else:
            return nome.split('(')[0].strip()

    # Extraindo apenas o nome do transportador antes dos parênteses
    df['TRANSPORTADOR_GRUPO'] = df['TRANSPORTADOR'].apply(
        extrair_transportador)

    # Identifica os transportadores únicos
    transportadores = df['TRANSPORTADOR_GRUPO'].unique()

    for transportador in transportadores:
        # Filtra o DataFrame para o transportador atual
        df_transportador = df[df['TRANSPORTADOR_GRUPO'] == transportador]

        # Gera a tabela HTML apenas para o transportador atual
        tabela_html = gerar_tabela_html(df_transportador)

        # Mensagem padrão a ser incluída no corpo do e-mail
        mensagem_padrao = f"""
        <p>Prezado(a),</p>
        <p>Por gentileza informar uma atualização sobre o status das coletas pendentes de inicialização de GE e,
        confirmar se os atendimentos ocorrerão normalmente.</p>
        <p>Confirme se há algum problema no atendimento ou se está tudo ok iniciem as GEs das Coletas ainda pendentes.</p>

        <p>Atenciosamente, Time ILC.</p>
        """

        # Monta o corpo do e-mail
        corpo_email = mensagem_padrao + tabela_html

        destinatarios = ""

        if "TW" in transportador:
            destinatarios = ''

        if "ARMANI" in transportador:
            destinatarios = ''

        if "JSL" in transportador:
            destinatarios = ''

        if "VF" in transportador:
            destinatarios = ''

        if "MIRASSOL (LOUVEIRA)" in transportador:
            destinatarios = MIRASSOL_LOUVERIA_UBERABA

        if "MIRASSOL (TUBARAO)" in transportador:
            destinatarios = ''

        if "MIRASSOL (CATALAO)" in transportador:
            destinatarios = ''

        if "MIRASSOL (UBERABA)" in transportador:
            destinatarios = ''

        if "MIRASSOL (SJ PINHAIS)" in transportador:
            destinatarios = ''

        # Ajusta destinatários e une com cópias
        todos_destinatarios = list(filter(None, destinatarios)) + CC

        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_USER
            # Substitua pelo e-mail real do transportador
            msg['To'] = ", ".join(todos_destinatarios)
            msg['Subject'] = f"Pendencias GE {hoje} - Teste Automatização {transportador}"

            msg.attach(MIMEText(corpo_email, 'html'))

            servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
            servidor.starttls()
            servidor.login(EMAIL_USER, EMAIL_PASS)
            servidor.sendmail(EMAIL_USER, msg['To'], msg.as_string())
            servidor.quit()

            print(f"E-mail enviado com sucesso para {transportador}!")
        except Exception as e:
            print(f"Erro ao enviar e-mail para {transportador}: {e}")


# 6. Execução principal
if __name__ == "__main__":
    caminho_arquivo = selecionar_arquivo()  # Abre janela para seleção do arquivo
    if caminho_arquivo:
        df = ler_dados_excel(caminho_arquivo)
        if df is not None:
            enviar_email_por_transportador(df)
        else:
            print("Erro: Não foi possível gerar a tabela HTML.")
    else:
        print("Erro: Nenhum arquivo selecionado.")

import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
import re  # Importa o módulo de expressões regulares
import shutil

# Definir as variáveis
network_path = r'\\192.168.0.239\ExportFiles\Nectar'
drive_letter = 'O:'

# Verificar se a unidade já está mapeada
if not os.path.exists(drive_letter):
    command = f'net use {drive_letter} {network_path} /persistent:yes'
    os.system(command)
    print(f'Unidade de rede {drive_letter} mapeada para {network_path}')
else:
    print(f'Unidade de rede {drive_letter} já está mapeada.')

# Configurações do email
smtp_server = "smtp.gmail.com"
port = 587
user_name = "aylaatilio@gmail.com"
password = "tcyw lyth ymns dtux"

def enviar_email_erro(mensagem_erro):
    """Envia um e-mail de erro."""
    subject = "ERRO EM AUDIO METADADOS"
    body = f"Erro ocorrido: {mensagem_erro}"
    
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = user_name
    msg['To'] = "aylaatilio@gmail.com"
    
    try:
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(user_name, password)
            server.send_message(msg)
            print("E-mail de erro enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def extrair_digitos(celula):
    """Extrai dígitos antes do caractere '_' de uma célula"""
    if not isinstance(celula, str):
        celula = str(celula)
    match = re.search(r'(\d+)_', celula)
    if match:
        return match.group(1)
    else:
        mensagem_erro = f"Nenhum número encontrado antes de '_' na célula: {celula}"
        enviar_email_erro(mensagem_erro)
        return "ERRO"

def processar_arquivo_csv(entrada, coluna):
    """Processa o arquivo CSV e extrai dígitos de uma coluna específica"""
    codificacoes = ['utf-8', 'latin1', 'ISO-8859-1', 'cp1252']
    delimitadores = [',', ';', '\t']
    df = None

    for codificacao in codificacoes:
        for delimitador in delimitadores:
            try:
                df = pd.read_csv(entrada, encoding=codificacao, sep=delimitador, on_bad_lines='skip')
                df.columns = df.columns.str.strip()
                if coluna in df.columns:
                    break
            except Exception:
                continue
        if df is not None and coluna in df.columns:
            break

    if df is None or coluna not in df.columns:
        return None

    df['CallIdMaster'] = df[coluna].apply(extrair_digitos)
    return df[['CallIdMaster']]

def mover_arquivos(diretorio_origem, diretorio_destino):
    """Move todos os arquivos do diretório de origem para o diretório de destino"""
    if not os.path.exists(diretorio_destino):
        os.makedirs(diretorio_destino)
    
    for arquivo in os.listdir(diretorio_origem):
        caminho_origem = os.path.join(diretorio_origem, arquivo)
        caminho_destino = os.path.join(diretorio_destino, arquivo)
        if os.path.isfile(caminho_origem):
            shutil.move(caminho_origem, caminho_destino)
            print(f"Arquivo movido para: {caminho_destino}")

def processar_varios_arquivos(arquivos, coluna, saida_dir):
    """Processa vários arquivos CSV e salva a saída em um único arquivo"""
    dfs = []
    for arquivo in arquivos:
        df_processado = processar_arquivo_csv(arquivo, coluna)
        if df_processado is not None:
            dfs.append(df_processado)
    
    if not dfs:
        print("Nenhum arquivo foi processado com sucesso.")
        return

    df_concatenado = pd.concat(dfs, ignore_index=True)
    df_concatenado['Data'] = datetime.now().strftime('%d/%m/%Y')

    if not os.path.exists(saida_dir):
        os.makedirs(saida_dir)

    nome_arquivo = f'Nectar_{datetime.now().strftime("%d_%m_%Y")}.csv'
    caminho_saida = os.path.join(saida_dir, nome_arquivo)
    df_concatenado.to_csv(caminho_saida, index=False, sep=';', header=['CallIdMaster', 'Data'])
    print(f'Arquivo concatenado salvo com sucesso em: {caminho_saida}')
    
    mover_arquivos(os.path.dirname(arquivos[0]), r'\\192.168.0.249\NectarServices\Nectar\Agendamentos\Monitoria')

# Lista de arquivos a serem processados
arquivos = [
    r'\\192.168.0.249\NectarServices\Nectar\Agendamentos\metadados_18.csv',
    r'\\192.168.0.249\NectarServices\Nectar\Agendamentos\metadados_55.csv'
]
coluna = 'nome_audio'
saida_dir = r'\\192.168.0.249\NectarServices\Nectar\Agendamentos\Monitoria'
#saida_dir = r'\\192.168.0.239\ExportFiles\Nectar'

processar_varios_arquivos(arquivos, coluna, saida_dir)
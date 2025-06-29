import os
import re
import sys
import socket
import gspread
import logging
import platform
import subprocess
import unicodedata
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
#   Pacote para testar o uso do ambiente virtual. 
#   No 'requeriments' esse pacote não está relacionado. 
#   Portanto se for utilizado irá apresentar falhar.
# from tabulate import tabulate 
from gspread.auth import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from openpyxl.utils.cell import range_boundaries
from google.oauth2.service_account import Credentials
from openpyxl.utils import get_column_letter, column_index_from_string

# Variáveis constantes (utilizadas de forma global no programa)
GOOGLE_DRIVE_SCOPE       = 'https://www.googleapis.com/auth/drive'
GOOGLE_SHEETS_SCOPE      = 'https://www.googleapis.com/auth/spreadsheets'
GOOGLE_DNS_SERVER        = "8.8.8.8"
GOOGLE_DNS_PORT          = 53
CONNECTION_TIMEOUT       = 5
DEFAULT_SHEET_ID         = 0  # Geralmente o ID da primeira aba
DEFAULT_BACKGROUND_COLOR = {'red': 0.85, 'green': 0.85, 'blue': 0.85}

def agendar_tarefa(tarefa_nome, tarefa_executar, tarefa_intervalo='hora', tarefa_tempo=1, hidden_bool=False):
    sistema = platform.system()
    tarefa_executar_path = sys.argv[0]
    tarefa_pasta = os.path.dirname(os.path.realpath(__file__))
    # produção
    tarefa_python = sys.executable # Caminho do Python do ambiente virtual atual
    # homologação (teste: simulando executável gerado pelo PyInstaller)
    #tarefa_python =  r'D:\capacita-brasil_equipe-8_5\indicadores3\CapacitaBrasilEquipe8-5Indicadores.exe'
    executavel = 0
    # homologação (teste: simulando executável gerado pelo PyInstaller)
    #if tarefa_python.lower().endswith('.exe'):
    if getattr(sys, 'frozen', False): # Quando o script é executado como um executável PyInstaller
        executavel = 1
        tarefa_executar_path = ""             # nenhum argumento necessário
        #tarefa_pasta = os.path.dirname(sys.argv[0]) # Pegamos o caminho do executável usando sys.argv[0]
        tarefa_pasta = os.path.dirname(tarefa_python) # Pegamos o caminho do executável usando sys.argv[0]

    if not os.path.exists(tarefa_python):
        mensagem(1, f"O caminho para execução não existe: {tarefa_python}", 'c')
        return

    if hidden_bool == True:
       if executavel == 0:
          python_dir = os.path.dirname(tarefa_python)
          tarefa_python = os.path.join(python_dir, 'pythonw.exe')        
            
    xml_template_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'template_tarefa.xml')
    xml_modified_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), f'{tarefa_nome}_modificado.xml')

    NS = {'t': 'http://schemas.microsoft.com/windows/2004/02/mit/task'}

    if sistema == "Windows":
        if executavel == 0:
           if not os.path.exists(tarefa_executar_path):
              mensagem(1,f"O arquivo '{tarefa_executar_path}' não foi encontrado.",'c')
              return
        else:
           if not os.path.exists(tarefa_python):
              mensagem(1,f"O arquivo '{tarefa_executar_path}' não foi encontrado.",'c')
              return

        if not os.path.exists(xml_template_path):
           mensagem(1,f"O arquivo XML template '{xml_template_path}' não foi encontrado. Certifique-se de que ele existe.",'c')
           return

        # Verifica se a tarefa já existe
        comando_query = ['schtasks', '/query', '/tn', tarefa_nome]
        resultado = subprocess.run(comando_query, capture_output=True, text=True)

        if resultado.returncode == 0:
            mensagem(1,f"Tarefa '{tarefa_nome}' já está agendada no Windows.")
            comando_status = ['schtasks', '/query', '/tn', tarefa_nome, '/fo', 'LIST']
            resultado_status = subprocess.run(comando_status, capture_output=True, text=True)

            if resultado_status.returncode == 0:
                status = resultado_status.stdout.replace(" ","")
                if "Status:Emexec" in status or "Status:Pronto" in status:
                    mensagem(1,f"Tarefa '{tarefa_nome}' já está ativa.",'i')
                else:
                    mensagem(1,f"Tarefa '{tarefa_nome}' não está ativa. Ativando...",'i')
                    # Com mensagem no console
                    #subprocess.run(['schtasks', '/run', '/tn', tarefa_nome], check=True)
                    # Sem mensagem no console
                    subprocess.run(['schtasks', '/run', '/tn', tarefa_nome], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            mensagem(1,f"Criando tarefa '{tarefa_nome}' usando XML...",'i')

            try:
                tree = ET.parse(xml_template_path)
                root = tree.getroot()

                # Modificar 'tags': Command, Arguments e WorkingDirectory
                # Encontra o elemento <Exec> dentro de <Actions>
                exec_element = root.find("./t:Actions/t:Exec", NS) # Caminho direto para o Exec dentro de Actions
                
                if exec_element is not None:
                    # Modifica a tag <Command> (caminho do Python)
                    command_element = exec_element.find("t:Command", NS)
                    if command_element is not None:
                        command_element.text = tarefa_python

                    # Modifica a tag <Arguments> (caminho do script Python)
                    arguments_element = exec_element.find("t:Arguments", NS)
                    if arguments_element is not None:
                       arguments_element.text = tarefa_executar_path

                    # Modifica a tag <WorkingDirectory> (diretório de trabalho)
                    working_directory_element = exec_element.find("t:WorkingDirectory", NS)
                    if working_directory_element is not None:
                        working_directory_element.text = tarefa_pasta # Não precisa de aspas aqui
                else:
                    mensagem(1,"Erro: A tag <Exec> dentro de <Actions> não foi encontrada no XML template. Verifique a estrutura do seu XML.",'c')
                    return # Sai da função se não encontrar a tag crucial

                # --- Opcional: Modificar o URI da tarefa ---
                uri_element = root.find("./t:RegistrationInfo/URI", NS)
                if uri_element is not None:
                    uri_element.text = f"\\{tarefa_nome}"
                
                # --- Opcional: Modificar o Trigger Intervalo dinamicamente ---
                # A sua tag <Interval> é PT2M. Se quiser mudar para 1 minuto (PT1M) ou 1 hora (PT1H)
                repetition_interval_element = root.find("./t:Triggers/t:TimeTrigger/t:Repetition/t:Interval", NS)
                if repetition_interval_element is not None:
                    if tarefa_intervalo.lower() == 'hora':
                        repetition_interval_element.text = f"PT{str(tarefa_tempo)}H" # Define para hora
                    elif tarefa_intervalo.lower() == 'minuto':
                        repetition_interval_element.text = f"PT{str(tarefa_tempo)}M" # Define para minuto (era PT2M no XML)
                
                # Opcional: Ajustar StartBoundary para ser a data/hora atual ou futura
                # Para evitar que a tarefa tente rodar com base em um StartBoundary passado no XML
                from datetime import datetime
                start_boundary_element = root.find("./t:Triggers/t:TimeTrigger/t:StartBoundary", NS)
                if start_boundary_element is not None:
                    # Define StartBoundary para a data/hora atual para iniciar imediatamente ou no próximo gatilho
                    start_boundary_element.text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                
                # ---Modificar a tag <Hidden> ---
                hidden_element = root.find("./t:Settings/t:Hidden", NS)
                if hidden_element is not None:
                    hidden_element.text = str(hidden_bool).lower() # Converte True/False para 'true'/'false'
                else:
                    mensagem(1,"Atenção: A tag <Hidden> não foi encontrada em <Settings> no XML template. A visibilidade não será definida.",'w')

                # Salva o XML modificado em um novo arquivo temporário
                # Usar 'xml_declaration=True' e 'encoding="UTF-16"' para manter o formato original do seu XML
                tree.write(xml_modified_path, encoding='UTF-16', xml_declaration=True)

                comando_args = [
                    'schtasks',
                    '/create',
                    '/tn', tarefa_nome,
                    '/xml', xml_modified_path
                ]

                subprocess.run(comando_args, check=True, capture_output=True, text=True)
                mensagem(1,f"Tarefa '{tarefa_nome}' agendada no Windows.",'i')

                # Opcional: Remover o XML temporário após o uso
                os.remove(xml_modified_path)

            except ET.ParseError as e:
                mensagem(1,f"Erro ao analisar o XML template: {e}",'c')
                mensagem(1,"Verifique se o arquivo XML está bem formatado e se não há caracteres inválidos.",'c')
            except subprocess.CalledProcessError as e:
                mensagem(1,f"Tarefa '{tarefa_nome}' não foi criada.",'c')
                mensagem(1,f"Erro: {e.stderr}",'c')
            except Exception as e:
                mensagem(1,f"Ocorreu um erro inesperado: {e}",'c')

    elif sistema == "Linux":
        cron_tag = f'# TASK_NAME:{tarefa_nome}'
        comentario = f'# Sérgio Sousa https://wa.me/+5585985265541. Agendamento: a cada {tarefa_tempo} {tarefa_intervalo}(s) - gerado automaticamente por script'

        comando_listar_cron = 'crontab -l'
        resultado = subprocess.run(comando_listar_cron, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if cron_tag in resultado.stdout.decode():
            mensagem(1, f"Tarefa '{tarefa_nome}' já está agendada no Linux.", 'i')

            comando_status = 'systemctl is-active cron'
            resultado_status = subprocess.run(comando_status, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            if resultado_status.stdout.decode().strip() == "inactive":
                mensagem(1, "Serviço cron não está ativo. Ativando o cron...", 'i')
                subprocess.run('systemctl start cron', shell=True)
            else:
                mensagem(1, "Serviço cron está ativo e funcionando.", 'i')
        else:
            # Define o cron_schedule com base em tarefa_intervalo e tarefa_tempo
            if tarefa_intervalo.lower() == 'minuto':
                cron_schedule = f"*/{tarefa_tempo} * * * *"
            elif tarefa_intervalo.lower() == 'hora':
                cron_schedule = f"0 */{tarefa_tempo} * * *"
            elif tarefa_intervalo.lower() == 'dia':
                cron_schedule = f"0 0 */{tarefa_tempo} * *"
            else:
                mensagem(1, f"Intervalo '{tarefa_intervalo}' não suportado no Linux. Usando padrão de 1 hora.", 'w')
                cron_schedule = "0 * * * *"

            exec_command = f'cd "{tarefa_pasta}" && "{sys.executable}" "{tarefa_executar_path}"'

            # Adiciona duas linhas ao crontab: o comentário + o agendamento
            comando = f'(crontab -l 2>/dev/null; echo "{comentario}"; echo "{cron_schedule} {exec_command} {cron_tag}") | crontab -'

            try:
                subprocess.run(comando, shell=True, check=True)
                mensagem(1, f"Tarefa '{tarefa_nome}' agendada no Linux.", 'i')
            except subprocess.CalledProcessError as e:
                mensagem(1, f"Erro ao agendar tarefa no Linux: {e.stderr.decode()}", 'w')
            except Exception as e:
                mensagem(1, f"Ocorreu um erro inesperado ao agendar no Linux: {e}", 'w')

    else:
        mensagem(1,f"Sistema {sistema} não suportado para agendamento de tarefas.",'w')

# (A função remover_tarefa permanece inalterada)
def remover_tarefa(tarefa_nome):
    if not tarefa_nome:
        mensagem(1,"A função 'remover_tarefa' precisa do nome da tarefa como parâmetro.",'w')
        return None

    sistema = platform.system()

    if sistema == "Windows":
        comando_query = ['schtasks', '/query', '/tn', tarefa_nome]
        resultado = subprocess.run(comando_query, capture_output=True, text=True)

        if resultado.returncode == 0:
            comando_remover = ['schtasks', '/delete', '/tn', tarefa_nome, '/f']
            subprocess.run(comando_remover, check=True)
            mensagem(1,f"Tarefa '{tarefa_nome}' removida do Windows.")
        else:
            mensagem(1,f"Tarefa '{tarefa_nome}' não encontrada no Windows.",'w')

    elif sistema == "Linux":
        cron_tag = f'# TASK_NAME:{tarefa_nome}'

        comando_listar_cron = 'crontab -l'
        resultado = subprocess.run(comando_listar_cron, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if cron_tag in resultado.stdout.decode():
            comando_remover = f'crontab -l | grep -v "{cron_tag}" | crontab -'
            try:
                subprocess.run(comando_remover, shell=True, check=True)
                mensagem(1,f"Tarefa '{tarefa_nome}' removida do crontab do Linux.")
            except subprocess.CalledProcessError as e:
                mensagem(1,f"Erro ao remover tarefa do Linux: {e.stderr.decode()}",'w')
            except Exception as e:
                mensagem(1,f"Ocorreu um erro inesperado ao remover no Linux: {e}",'w')
        else:
            mensagem(1,f"Tarefa '{tarefa_nome}' não encontrada no crontab do Linux.",'w')

    else:
        mensagem(1,f"Sistema {sistema} não suportado para remoção de tarefas.",'w')

# Configuração do sistema de log
def configurar_log():
    log_dir = 'logs'  # Diretório onde os logs serão armazenados
    if not os.path.exists(log_dir):
       os.makedirs(log_dir)

    # Obtém o timestamp atual e formata como dd-mm-aaaa_hh-mm-ss
    data_formatada = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")		

    log_filename = os.path.join(log_dir, f'execucao_{data_formatada}.log')
		
    logging.basicConfig(
			level=logging.DEBUG,  # Define o nível mínimo de log
			format='%(asctime)s - %(levelname)s - %(message)s',
			handlers=[
				logging.FileHandler(log_filename, encoding='utf-8'),  # Salva o log no arquivo com codificação de caracteres específica.
				#logging.StreamHandler()  # Exibe o log no console
			]
	)
    # Ajusta o formato da data para dd/mm/aaaa
    for handler in logging.root.handlers:
        if isinstance(handler, logging.FileHandler):
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%d/%m/%Y às %H:%M:%S'))
    
    #mensagem(0,"")
    #mensagem(0,"Sistema de Log iniciado.")

# Verificar conexão com a internet
def verificar_conexao(host=GOOGLE_DNS_SERVER, port=GOOGLE_DNS_PORT, timeout=CONNECTION_TIMEOUT):
    try:
        socket.create_connection((host, port), timeout=timeout)
        mensagem(0,f"")
        mensagem(0,"Conectado a internet.")
        return True
    except OSError as e:
        mensagem(0,f"")
        mensagem(0,f"Erro de conexão com a internet: {e}", 'c')
        return False

# Inicia os serviços da API do Google Drive, Google Sheets e autentica o cliente gspread.
# Retorna uma tupla contendo as instâncias de serviço do Drive, Sheets e o cliente gspread.
def conectar_google_apis():
    SCOPES = [GOOGLE_DRIVE_SCOPE, GOOGLE_SHEETS_SCOPE, "https://spreadsheets.google.com/feeds"]
    creds_path = os.getenv('GOOGLE_CREDS_JSON_PATH')
    if not creds_path:
        mensagem(0,"Arquivo de credenciais não especificado ou não encontrado na variável de ambiente 'GOOGLE_CREDS_JSON_PATH'.", 'c')
        orientacoes()
        sys.exit(1)

    try:
        # Autenticação para googleapiclient (Drive e Sheets)
        creds          = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        service_drive  = build('drive', 'v3', credentials=creds)
        service_sheets = build('sheets', 'v4', credentials=creds)
        # Autenticação para gspread
        gs_client = gspread.authorize(creds)
        mensagem(1,"Serviços do Google Drive, Sheets e cliente gspread ativos.")
        return service_drive, service_sheets, gs_client
    except Exception as e:
        mensagem(1,f"Falha ao carregar credenciais ou construir serviços: {e}", 'c')
        sys.exit(1)

# Extrai o ID de um link de arquivo ou pasta do Google Drive.
def link_id(service_drive, url: str):
    url_id = re.search(r"(?:folders|file)/([a-zA-Z0-9_-]+)", url)
    id_retorno = ''
    if url_id:
        id = url_id.group(1)
        informacoes = informacoes_driver(service_drive, id)
        if informacoes:
           mensagem(0,f"")
           mensagem(0,f"Pasta do compartilhamento: {informacoes[0]}")
           mensagem(1,f"link: {url}")
           id_retorno = id
    else:
        mensagem(0,"Pasta do compartilhamento não encontrada:", 'e')
        mensagem(1,f"ID não encontrado na URL: {url}")
    return id_retorno

# Função para tratar datas de 2 e 4 dígitos
def ajustar_data(data):
    # Tentar primeiro converter com 4 dígitos no ano
    try:
        return pd.to_datetime(data, format='%d/%m/%Y', errors='raise', dayfirst=True)
    except:
        # Se falhar, tentar com 2 dígitos no ano
        try:
            return pd.to_datetime(data, format='%d/%m/%y', errors='raise', dayfirst=True)
        except:
            return pd.NaT  # Se não conseguir converter, retorna NaT

def remover_acentos(texto):
    if isinstance(texto, str):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto or ''

# Carrega dados de uma aba específica de uma planilha, realiza transformações e validações.
# Retorna um DataFrame com os dados transformados e a lista de colunas finais.
def carregar_dados_planilha(planilha, aba_nome):
    aba = planilha.worksheet(aba_nome)
    # A condição "IF" é utiliza para acrescentar outras abas e seus processamentos.
    # A condição para 'Dados Seleção' é do 'Banco 2', e serve de referência para quando for juntar os códigos.
    if 'banco 2' in planilha.title.lower() and aba_nome.lower() == 'dados seleção':
       dados = pd.DataFrame(aba.get_all_records())
       colunas_necessarias = ['Ano', 'Nome', 'Contrato', 'Cidade', 'Estado', 'Área']
       if not all(col in dados.columns for col in colunas_necessarias):
          missing_cols = [col for col in colunas_necessarias if col not in dados.columns]
          mensagem(1,f"Colunas obrigatórias faltando na aba '{aba}': {', '.join(missing_cols)}", 'c')
          mensagem(1,'Programa interrompido.', 'c')
          sys.exit(1)
        
       dados = dados[colunas_necessarias].copy()
       # Acrescentar uma coluna fixa 'País' com conteúdo "Brasil"
       dados['País'] = 'Brasil'
       # Transforma o conteúdo da coluna 'Contrato' de 'Sim' para 'Aprovado' e 'Não' para 'Inscrito'.
       dados['Contrato'] = dados['Contrato'].replace({'Sim': 'Aprovado', 'Não': 'Inscrito'})
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Nome': 'identificação',
            'Contrato': 'status',
            'Cidade': 'cidade',
            'Estado': 'UF',
            'Área': 'area'
        })
       # Define os títulos das colunas com os novos nomes
       colunas_final = ['Ano', 'identificação', 'status', 'cidade', 'UF', 'País', 'area']
       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       # Contabiliza registro duplicados 
       repetidos = dados.duplicated().sum()
       if repetidos > 0:
          mensagem(1,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
       else:
          mensagem(1,"Nenhum registro repetido encontrado.")
    elif 'banco 4' in planilha.title.lower() and aba_nome == "Banco de Consultorias":
       mensagem(1,"")
       mensagem(1,f"Analisando aba '{aba_nome}'...")
       valores = aba.get_all_values()
       dados = pd.DataFrame(valores[1:], columns=valores[0])
       colunas_necessarias = ['Ano', 'Nome do Consultor', 'Quantidade de horas', 'Área']
       if not all(col in dados.columns for col in colunas_necessarias):
          missing_cols = [col for col in colunas_necessarias if col not in dados.columns]
          mensagem(2,f"Colunas obrigatórias faltando: {', '.join(missing_cols)}", 'c')
          mensagem(2,"Programa interrompido.")
          sys.exit(1)

       # Remove linhas totalmente vazias
       dados = dados.dropna(how='all')  
       # Remove linhas onde qualquer uma das colunas obrigatórias está vazia
       dados = dados.dropna(subset=['Ano', 'Nome do Consultor', 'Quantidade de horas', 'Área'])
       # Remove linhas onde 'Nome do Consultor' está vazio ou contém apenas espaços
       dados = dados[dados['Nome do Consultor'].str.strip() != '']       
       # Contabiliza registro duplicados 
       colunas_duplicadas = ['Ano', 'Nome do Consultor', 'Nome Startup', 'Quantidade de horas', 'Área']
       repetidos = dados.duplicated(subset=colunas_duplicadas).sum()
       if repetidos > 0:
          mensagem(2,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
          mensagem(2,"Registros duplicados (removido):", 'w')
          mensagem(2,dados[dados.duplicated()], 'w')
          # Remove os duplicados
          dados = dados.drop_duplicates(subset=colunas_duplicadas)          
       else:
          mensagem(2,"Nenhum registro repetido encontrado.")
       dados = dados[colunas_necessarias].copy()
       # Acrescentar uma coluna fixa 'atividade' com conteúdo "Consultoria"
       dados['atividade'] = 'Consultoria'
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Ano': 'ano',
            'Nome do Consultor': 'profissional',
            'Quantidade de horas': 'horas',
            'Área': 'area'
        })
       # Converte a coluna 'horas' para tipo numérico (float)
       dados['horas'] = pd.to_numeric(dados['horas'], errors='coerce')
       # Remove acentuação de colunas do tipo texto
       for col in ['profissional', 'atividade']:
           dados[col] = dados[col].apply(remover_acentos)       
       # Remover espaços em branco para impedir repetições por diferenças:
       # Exemplo: "Paulo", "Paulo " e " Paulo" são conteúdos diferentes.
       for col in ['ano', 'profissional', 'atividade', 'area']:
           dados[col] = dados[col].astype(str).str.strip()
       # Capitaliza a primeira letra de cada palavra no nome do profissional
       dados['profissional'] = dados['profissional'].str.title()
       # Corrige nomes específicos (ex: remove "Dos" de "Moises Dos Santos")
       dados['profissional'] = dados['profissional'].replace({
             'Moises Dos Santos': 'Moises Santos'
       })       
       # Define os títulos das colunas com os novos nomes
       colunas_final = ['ano', 'atividade', 'profissional', 'horas', 'area']
       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       mensagem(2,f"Quantidade de registros: {len(dados)}")

    elif 'banco 4' in planilha.title.lower() and aba_nome == "Banco de Mentorias":
       mensagem(1,"")
       mensagem(1,f"Analisando aba '{aba_nome}'...")
       valores = aba.get_all_values()
       dados = pd.DataFrame(valores[1:], columns=valores[0])
       colunas_necessarias = ['Data', 'Nome do mentor', 'Horas de Mentorias', 'Mentoria']
       if not all(col in dados.columns for col in colunas_necessarias):
          missing_cols = [col for col in colunas_necessarias if col not in dados.columns]
          mensagem(2,f"Colunas obrigatórias faltando: {', '.join(missing_cols)}", 'c')
          mensagem(2,"Programa interrompido.")
          sys.exit(1)
       # Remove linhas totalmente vazias
       dados = dados.dropna(how='all')  
       # Remove linhas onde qualquer uma das colunas obrigatórias esteja vazia
       dados = dados.dropna(subset=['Data', 'Nome do mentor', 'Horas de Mentorias', 'Mentoria'])
       # Remove linhas onde 'Nome do mentor' está vazio ou contém apenas espaços
       dados = dados[dados['Nome do mentor'].str.strip() != '']       
       # Convertendo a coluna 'Data' para o formato de data, caso não esteja
       dados['Data'] = pd.to_datetime(dados['Data'], errors='coerce', dayfirst=True)
       # Remove linhas com datas inválidas após conversão
       dados = dados.dropna(subset=['Data'])  
       # Contabiliza registro duplicados 
       colunas_duplicadas = ['Data', 'Mentoria', 'Horas de Mentorias', 'Nome do mentor', 'Mentoria']
       repetidos = dados.duplicated(subset=colunas_duplicadas).sum()
       if repetidos > 0:
          mensagem(2,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
          mensagem(2,"Registros duplicados (removido):", 'w')
          mensagem(2,dados[dados.duplicated()], 'w')
          # Remove os duplicados
          dados = dados.drop_duplicates(subset=colunas_duplicadas)          
       else:
           mensagem(2,"Nenhum registro repetido encontrado.")
       # Acrescenta a coluna 'Ano' através da coluna 'Data' 
       dados['Ano'] = dados['Data'].dt.year
       colunas_necessarias = ['Ano', 'Nome do mentor', 'Horas de Mentorias', 'Mentoria']
       dados = dados[colunas_necessarias].copy()
       # Acrescentar uma coluna fixa 'atividade' com conteúdo "Mentoria"
       dados['atividade'] = 'Mentoria'
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={            
            'Ano': 'ano',
            'Nome do mentor': 'profissional',
            'Horas de Mentorias': 'horas',
            'Mentoria': 'area'
        })
       # Converte a coluna 'horas' para tipo numérico (float)
       dados['horas'] = pd.to_numeric(dados['horas'], errors='coerce')
       # Remove acentuação de colunas do tipo texto
       for col in ['profissional', 'atividade']:
           dados[col] = dados[col].apply(remover_acentos)       
       # Remover espaços em branco para impedir repetições por diferenças:
       # Exemplo: "Paulo", "Paulo " e " Paulo" são conteúdos diferentes.
       for col in ['ano', 'profissional', 'atividade']:
           dados[col] = dados[col].astype(str).str.strip()
       # Capitaliza a primeira letra de cada palavra no nome do profissional
       dados['profissional'] = dados['profissional'].str.title()
       # Corrige nomes específicos (ex: remove "Dos" de "Moises Dos Santos")
       dados['profissional'] = dados['profissional'].replace({
             'Moises Dos Santos': 'Moises Santos'
       })       
       # Define os títulos das colunas com os novos nomes
       colunas_final = ['ano', 'atividade', 'profissional', 'horas', 'area']
       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       mensagem(2,f"Quantidade de registros: {len(dados)}")
    elif 'banco 5' in planilha.title.lower() and aba_nome == "Versão resumida":
       valores = aba.get_all_values()
       dados = pd.DataFrame(valores[1:], columns=valores[0])
       # Remove linhas totalmente vazias
       dados = dados.dropna(how='all')  
       colunas_obrigatorias = ['Nome da Incubada Graduada:', 'Ano de graduação:']
       if not all(col in dados.columns for col in colunas_obrigatorias):
          missing_cols = [col for col in colunas_obrigatorias if col not in dados.columns]
          mensagem(1,f"Colunas obrigatórias faltando: {', '.join(missing_cols)}", 'c')
          mensagem(1,"Programa interrompido.", 'c')
          sys.exit(1)

       # Remove linhas onde qualquer uma das colunas obrigatórias está vazia
       dados = dados.dropna(subset=colunas_obrigatorias)
       # Remove linhas onde 'Nome do Consultor' está vazio ou contém apenas espaços
       dados = dados[dados['Nome da Incubada Graduada:'].str.strip() != '']       
       # Contabiliza registro duplicados 
       repetidos = dados.duplicated(subset=colunas_obrigatorias).sum()
       if repetidos > 0:
          mensagem(1,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
          mensagem(1,"Registros duplicados (removido):", 'w')
          mensagem(1,dados[dados.duplicated()], 'w')
          # Remove os duplicados
          dados = dados.drop_duplicates(subset=colunas_obrigatorias)
       else:
           mensagem(1,"Nenhum registro repetido encontrado.")
       
       dados = dados[colunas_obrigatorias].copy()
       # Remover espaços em branco para impedir repetições por diferenças:
       # Exemplo: "Paulo", "Paulo " e " Paulo" são conteúdos diferentes.
       for col in colunas_obrigatorias:
           dados[col] = dados[col].astype(str).str.strip()

       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Nome da Incubada Graduada:': 'nome',
            'Ano de graduação:': 'ano'
        })
       # Define os títulos das colunas com os novos nomes
       colunas_final = ['nome','ano']
       # Atribui esses novos nomes 
       dados = dados[['nome','ano']]
       mensagem(1,f"Quantidade de registros: {len(dados)}")
    elif 'banco 6' in planilha.title.lower() and aba_nome == "Banco de Parceiro":
       mensagem(1,"")
       mensagem(1,f"Analisando aba '{aba_nome}'...")
       valores = aba.get_all_values()
       dados = pd.DataFrame(valores[1:], columns=valores[0])
       # Verifica se as colunas de processamento existem
       colunas_obrigatorias = ['Ano', 'Parceiro']
       if not all(col in dados.columns for col in colunas_obrigatorias):
          missing_cols = [col for col in colunas_obrigatorias if col not in dados.columns]
          mensagem(2,f"Colunas obrigatórias faltando: {', '.join(missing_cols)}", 'c')
          mensagem(2,"Programa interrompido.", 'c')
          sys.exit(1)
       # Remove linhas onde qualquer uma das colunas obrigatórias está vazia
       dados = dados.dropna(subset=colunas_obrigatorias)
       # Remove linhas onde 'Nome do Consultor' está vazio ou contém apenas espaços
       dados = dados[dados['Parceiro'].str.strip() != '']       
       # Contabiliza registro duplicados 
       repetidos = dados.duplicated(subset=colunas_obrigatorias).sum()
       if repetidos > 0:
          mensagem(2,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
          mensagem(2,"Registros duplicados (removido):", 'w')
          mensagem(2,dados[dados.duplicated()], 'w')
          # Remove os duplicados
          dados = dados.drop_duplicates(subset=colunas_obrigatorias)
       else:
           mensagem(2,"Nenhum registro repetido encontrado.")
       dados = dados[colunas_obrigatorias].copy()
       # Acrescentar uma coluna fixa 'atividade' com conteúdo "Parceria"
       dados['quantidade'] = 1
       dados['atividade'] = 'Parceria'
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Ano': 'Ano',
            'Parceiro': 'Descrição',
            'quantidade': 'Quantidade',
            'atividade': 'Atividade'
        })
       # Garantir que 'Ano' seja do tipo integer, tratando 'NaN' corretamente
       dados['Ano'] = pd.to_numeric(dados['Ano'], errors='coerce', downcast='integer')
       # Colunas de retorno para gerar a planilha
       colunas_final = ['Ano', 'Descrição', 'Quantidade', 'Atividade']
       # Remover espaços em branco para impedir repetições por diferenças:
       # Exemplo: "Paulo", "Paulo " e " Paulo" são conteúdos diferentes.
       for col in colunas_final:
           if dados[col].dtype == 'O':  # 'O' indica o tipo 'object', geralmente para strings
              dados[col] = dados[col].astype(str).str.strip()

       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       mensagem(2,f"Quantidade de registros: {len(dados)}")
    elif 'banco 6' in planilha.title.lower() and aba_nome == "Impactos de Marketing":
       mensagem(1,"")
       mensagem(1,f"Analisando aba '{aba_nome}'...")
       valores = aba.get_all_values()
       dados = pd.DataFrame(valores[1:], columns=valores[0])
       # Verifica se as colunas de processamento existem
       colunas_obrigatorias = ['Data', 'Postagem','Impacto']
       if not all(col in dados.columns for col in colunas_obrigatorias):
          missing_cols = [col for col in colunas_obrigatorias if col not in dados.columns]
          mensagem(2,f"Colunas obrigatórias faltando: {', '.join(missing_cols)}", 'c')
          mensagem(2,"Programa interrompido.", 'c')
          sys.exit(1)
       # Convertendo a coluna 'Data do Evento' para o formato de data, caso não esteja
       # Utilizando a função 'ajustar_data' para verificar digitação com 2 ou 4 digitos no ano.
       dados['Data'] = dados['Data'].apply(ajustar_data)
       # Acrescenta a coluna 'Ano' através da coluna 'Data' 
       dados['Ano'] = dados['Data'].dt.year
       # Remove linhas totalmente vazias
       dados = dados.dropna(how='all')  
       # Remove linhas onde qualquer uma das colunas obrigatórias está vazia
       dados = dados.dropna(subset=colunas_obrigatorias)
       # Remove linhas onde esteja vazio ou contém apenas espaços
       dados = dados[dados['Postagem'].str.strip() != '']
       # Contabiliza registro duplicados 
       repetidos = dados.duplicated(subset=colunas_obrigatorias).sum()
       if repetidos > 0:
          mensagem(2,f"Quantidade de registros repetidos encontrados: {repetidos}", 'w')
          mensagem(2,"Registros duplicados (removido):", 'w')
          mensagem(2,dados[dados.duplicated()], 'w')
          # Remove os duplicados
          dados = dados.drop_duplicates(subset=colunas_obrigatorias)
       else:
           mensagem(2,"Nenhum registro repetido encontrado.")
       colunas_obrigatorias = ['Ano', 'Postagem','Impacto']
       dados = dados[colunas_obrigatorias].copy()
       # Acrescentar uma coluna fixa 'atividade' com conteúdo "Marketing"
       dados['atividade'] = 'Marketing'
       # Altera o nome das colunas para ficar compatível com o dashboard criado.
       dados = dados.rename(columns={
            'Ano': 'Ano',
            'Postagem': 'Descrição',
            'Impacto': 'Quantidade',
            'atividade': 'Atividade'
        })
       # Garantir que 'Ano' e 'Quantidade' sejam do tipo integer, tratando 'NaN' corretamente
       dados['Ano'] = pd.to_numeric(dados['Ano'], errors='coerce', downcast='integer')
       dados['Quantidade'] = pd.to_numeric(dados['Quantidade'], errors='coerce', downcast='integer')
       # Colunas de retorno para gerar a planilha
       colunas_final = ['Ano', 'Descrição', 'Quantidade', 'Atividade']
       # Remover espaços em branco para impedir repetições por diferenças:
       # Exemplo: "Paulo", "Paulo " e " Paulo" são conteúdos diferentes.
       for col in colunas_final:
           if dados[col].dtype == 'O':  # 'O' indica o tipo 'object', geralmente para strings
              dados[col] = dados[col].str.strip()  # Remove espaços em branco

       # Define os títulos das colunas com os novos nomes
       # Atribui esses novos nomes 
       dados = dados[colunas_final]
       mensagem(2,f"Quantidade de registros: {len(dados)}")

    return dados, colunas_final

# Prepara os intervalos de células para o cabeçalho e os dados na planilha.
# Retorna uma tupla contendo as strings dos intervalos do cabeçalho e dos dados.
def preparar_intervalos(cabecalho: list, dados: pd.DataFrame):
    col_ini   = os.getenv('planilha_coluna_inicial', 'A').upper()
    linha_ini = os.getenv('planilha_linha_inicial', '1')

    try:
        linha_ini = int(linha_ini)
    except ValueError:
        mensagem(1,f"função 'preparar_intervalos' falhou 'linha_ini'={linha_ini} inválida. Será usado '1'.", 'w')
        linha_ini = 1

    # Para o cabeçalho, a linha inicial é a definida
    intervalo_cabecalho = planilha_celulas_intervalo(col_ini, linha_ini, [cabecalho], 'c')
    
    # Para os dados, a linha inicial é a linha do cabeçalho + 1 (onde os dados efetivamente começam)
    linha_dados_ini = linha_ini + 1
    intervalo_dados = planilha_celulas_intervalo(col_ini, linha_dados_ini, dados.values.tolist(), 'd')

    return intervalo_cabecalho, intervalo_dados

# Verifica se uma pasta existe no Google Drive pelo nome (pasta_nome) e local onde pesquisar (parent_id: pasta pai)
def pasta_existe(service_drive, pasta_nome: str, parent_id: str = None):
    query = f"name = '{pasta_nome}' and mimeType = 'application/vnd.google-apps.folder'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    
    try:
        results = service_drive.files().list(q=query, fields="files(id)").execute()
        files = results.get('files', [])
        if files:
            return files[0]['id']
        return None
    except HttpError as e:
        return None

# Verifica se uma planilha existe no Google Drive pelo nome (planilha_nome) e local onde pesquisar (parent_id: pasta pai)
def planilha_existe(service_drive, planilha_nome: str, parent_id: str):
    query = f"name = '{planilha_nome}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    try:
        results = service_drive.files().list(q=query, fields="files(id)").execute()
        files = results.get('files', [])
        if files:
           return files[0]['id']
        return None
    except HttpError as e:
        mensagem(1,f"Erro ao verificar a existência da planilha '{planilha_nome}': {e}", 'e')
        mensagem(1,f"falha: {e}", 'e')
        return None

# Cria uma pasta no Google Drive, com opção de definir uma pasta pai.
# Se a pasta já existir, retorna o ID da pasta existente.
# Atribui permissões 'anyone' e 'writer' nas duas situações.
def criar_pasta(service_drive, pasta_nome: str, parent_id: str = None):
    pasta_id = pasta_existe(service_drive, pasta_nome, parent_id)
    mensagem1 = mensagem2 = mensagem3 = ''
    log = 'i'
    if pasta_id:
        mensagem1 = f"Sub-pasta encontrada: {pasta_nome}"
        mensagem2 = f'link: https://drive.google.com/drive/folders/{pasta_id}'
        mensagem3 = permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'writer')
    else:
        file_metadata = {
            'name': pasta_nome,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_id:
            file_metadata['parents'] = [parent_id]

        try:
            pasta = service_drive.files().create(body=file_metadata, fields='id').execute()
            pasta_id = pasta['id']
            mensagem1 = f"Sub-pasta criada: {pasta_nome}"
            mensagem2 = f'link: https://drive.google.com/drive/folders/{pasta_id}'
            mensagem3 = permissoes_pasta_arquivo(service_drive, pasta_id, 'anyone', 'writer')
        except HttpError as e:
            log = 'e'
            mensagem1 = f"Erro ao criar a pasta '{pasta_nome}'"
            mensagem2 = f"falha: {e}"
            raise
    mensagem(0,mensagem1, log)
    mensagem(1,mensagem2, log)
    mensagem(1,mensagem3, log)
    return pasta_id

# Cria uma planilha no Google Sheets, com opção de definir uma pasta pai.
# Se a planilha já existir, retorna o ID da planilha existente.
def criar_planilha(service_drive, service_sheets, planilha_nome: str, parent_id: str = None):
    mensagem1 = mensagem2 = mensagem3 = ''
    planilha_id = planilha_existe(service_drive, planilha_nome, parent_id)
    log = 'i'
    if planilha_id:
        mensagem1 = f"Planilha encontrada: {planilha_nome}"
        mensagem2 = f"link: https://docs.google.com/spreadsheets/d/{planilha_id}"
    else:
        body = {
            'properties': {'title': planilha_nome}
        }
        try:
            # Criando planilha
            planilha = service_sheets.spreadsheets().create(body=body, fields='spreadsheetId').execute()
            planilha_id = planilha['spreadsheetId']
            # Transferindo planilha para a pasta do compartilhamento (pasta pai)
            if parent_id:
                service_drive.files().update(
                    fileId=planilha_id,
                    addParents=parent_id,
                    fields='parents'
                ).execute()
            mensagem1 = f"Planilha criada: {planilha_nome} "
            mensagem2 = f"link: https://docs.google.com/spreadsheets/d/{planilha_id}"
        except HttpError as e:
            log = 'e'
            mensagem1 = f"Erro ao criar a planilha '{planilha_nome}'"
            mensagem2 = f"falha: {e}"
            raise
    mensagem(0,mensagem1, log)
    mensagem(1,mensagem2, log)
    return planilha_id

# Atribui permissões por tipo e função a uma pasta ou arquivo no Google Drive.
def permissoes_pasta_arquivo(service_drive, item_id: str, tipo: str, funcao: str):
    informacoes = informacoes_driver(service_drive, item_id)
    mensagem = ''
    permissao = {
        'type': tipo,
        'role': funcao
    }
    try:
        service_drive.permissions().create(
            fileId=item_id,
            body=permissao
        ).execute()
        mensagem = f"Permissões atribuídas: acesso '{tipo}' e permissão '{funcao}'."
    except HttpError as e:
        mensagem = f"Falha ao atribuir permissões: {e}"
    return mensagem

# Retorna o nome e o tipo MIME de um item (arquivo ou pasta) no Google Drive.
def informacoes_driver(service_drive, item_id: str):
    try:
        conteudo = service_drive.files().get(fileId=item_id, fields='name,mimeType').execute()
        conteudo_nome = conteudo.get('name', '')
        conteudo_tipo = conteudo.get('mimeType', '')
        return conteudo_nome, conteudo_tipo
    except HttpError as e:
        mensagem(0,f"", 'e')
        mensagem(0,f"informacoes_driver: Erro ao obter informações para o item com ID '{item_id}'", 'e')
        mensagem(1,f"falha: {e}", 'e')
        return ""
    
# Limpeza da aba da planilha antes da gravação
def planilha_aba_limpeza(service_sheets, planilha_id: str, aba_nome):
    try:
        mensagem(1,f"Preparando aba: limpeza.")
        # Limpa a planilha inteira (intervalo arbitrariamente grande)
        clear_request = service_sheets.spreadsheets().values().clear(
        spreadsheetId=planilha_id,
                range=f"'{aba_nome}'!A1:Z1000"  # Limpa a aba especificada
        )
        clear_request.execute()
    except Exception as e:
        mensagem(1,f"Função 'planilha_aba_limpeza': Exception: {e}", 'e')
    except HttpError as e:
        mensagem(1,f"Função 'planilha_aba_limpeza': HttpError: {e}", 'e')
        return None

# Grava informações em um intervalo específico de uma planilha.
def planilha_dados(service_sheets, planilha_id: str, aba_nome, intervalo: str, dados_para_gravar: list):
    try:
        # Monta o intervalo completo usando o nome da aba
        intervalo_completo = f"'{aba_nome}'!{intervalo}"

        request = service_sheets.spreadsheets().values().update(
            spreadsheetId=planilha_id,
            range=intervalo_completo,
            valueInputOption='RAW',
            body={'values': dados_para_gravar}
        )
        response = request.execute()
        mensagem(1,f"Dados inseridos no intervalo: {intervalo}")
        return response
    except Exception as e:
        mensagem(1,f"Função 'planilha_dados': Exception: {e}", 'e')
    except HttpError as e:
        mensagem(1,f"Função 'planilha_dados': HttpError: {e}", 'e')
        return None

# Apaga uma pasta ou arquivo pelo ID ou nome fornecido.
def apagar_pasta_arquivo(service_drive, item_id: str = None, item_nome: str = None, parent_id: str = None):
    if not item_id and not item_nome:
        mensagem(1,"função 'apagar_pasta_arquivo' falhou, sem parâmetro 'item_id' ou 'item_nome'.", 'e')
        return

    if item_nome:
        found_id = pasta_existe(service_drive, item_nome, parent_id)
        if not found_id:
            found_id = planilha_existe(service_drive, item_nome, parent_id)
        
        if not found_id:
            mensagem(1,"função 'apagar_pasta_arquivo' não encontrou {item_nome}.", 'e')
            return
        item_id = found_id

    informacoes = informacoes_driver(service_drive, item_id)
    item_nome_real = informacoes[0]
    item_tipo_mime = informacoes[1]
    # Verifica se o item é uma pasta
    if 'folder' in item_tipo_mime:
        query = f"'{item_id}' in parents"
        try:
            results = service_drive.files().list(q=query, fields="files(id, name)").execute()
            files_in_folder = results.get('files', [])

            if not files_in_folder:
                service_drive.files().delete(fileId=item_id).execute()
                mensagem(1,f"Pasta '{item_nome_real}' removida.")
            else:
                mensagem(1,f"A pasta '{item_nome_real}' contém {len(files_in_folder)} arquivo(s).")
                resposta = input("Digite 'sim' para apagar os arquivos dentro dela e a pasta: ").strip().lower()
                if resposta.lower() == 'sim':
                    for arquivo in files_in_folder:
                        try:
                            service_drive.files().delete(fileId=arquivo['id']).execute()
                            mensagem(2,f"Arquivo '{arquivo['name']}' apagado.")
                        except HttpError as error:
                            mensagem(2,f"Erro ao apagar o arquivo '{arquivo['name']}': {error}", 'e')
                    
                    try:
                        service_drive.files().delete(fileId=item_id).execute()
                        mensagem(1,f"Pasta '{item_nome_real}' foi removida com todos os arquivos.")
                    except HttpError as error:
                        mensagem(1,f"Erro ao remover a pasta '{item_nome_real}': {error}", 'e')
                else:
                    mensagem(1,"A exclusão da pasta foi cancelada.", 'w')
        except HttpError as e:
            mensagem(1,f"Erro ao verificar ou apagar a pasta '{item_nome_real}': {e}", 'e')
    # Se for uma arquivo (planilha)
    else:
        try:
            service_drive.files().delete(fileId=item_id).execute()
            mensagem(1,f"Arquivo '{item_nome_real}' apagado!")
        except HttpError as e:
            mensagem(1,f"Erro ao remover o arquivo '{item_nome_real}': {e}", 'e')

# Apagar aba (sheet) de uma planilha pelo nome.
def apagar_aba(service_sheets, planilha_id, aba_nome):
    try:
        # Obter as informações da planilha
        planilha = service_sheets.spreadsheets().get(spreadsheetId=planilha_id).execute()

        # Verificar se a aba existe
        abas = planilha['sheets']
        aba_existe = False
        aba_id = None

        for aba in abas:
            if aba['properties']['title'] == aba_nome:
                aba_existe = True
                aba_id = aba['properties']['sheetId']
                break

        # Se a aba existe, verifique o número de abas restantes
        if aba_existe:
            if len(abas) > 1:  # Verifica se há mais de uma aba na planilha
                # Fazer a requisição para excluir a aba
                requests = [{
                    'deleteSheet': {
                        'sheetId': aba_id
                    }
                }]
                
                # Executar a requisição para deletar a aba
                response = service_sheets.spreadsheets().batchUpdate(
                    spreadsheetId=planilha_id,
                    body={'requests': requests}
                ).execute()
                mensagem(1,f"Aba excluída: {aba_nome}")
        else:
            mensagem(1,f"Aba não encontrada: {aba_nome}", 'e')
    except Exception as e:
        mensagem(1,f"Erro ao tentar excluir a aba: {e}", 'e')

# Obter o Id da aba da planilha pelo nome.
def id_aba_planilha_por_nome(service_sheets, planilha_id, aba_nome, mensagem_mostrar=False):
    # Obtém todas as abas da planilha
    result = service_sheets.spreadsheets().get(spreadsheetId=planilha_id).execute()
    sheets = result.get('sheets', [])
    aba_padrao = 'Sheet1'

    # Verifica se a aba já existe
    for sheet in sheets:
        if sheet['properties']['title'] == aba_padrao:
           apagar_aba(service_sheets, planilha_id, aba_padrao)
        if sheet['properties']['title'] == aba_nome:
            if mensagem_mostrar:
               mensagem(1,f"Aba encontrada: {aba_nome}")
            return sheet['properties']['sheetId']  # Retorna o ID da aba existente

    # Se não encontrar, cria a aba
    request = {
        'requests': [
            {
                'addSheet': {
                    'properties': {
                        'title': aba_nome
                    }
                }
            }
        ]
    }
    response = service_sheets.spreadsheets().batchUpdate(spreadsheetId=planilha_id, body=request).execute()
    if mensagem_mostrar:
       mensagem(1,f"Aba criada: {aba_nome}")
    return response['replies'][0]['addSheet']['properties']['sheetId']  # Retorna o ID da nova aba

# Converte uma string de intervalo (ex: 'B3:H9') para (min_row, max_row, min_col, max_col), utilizando o pacote 'openpyxl'.
def celula_intervalo_para_linhas_colunas(intervalo_str: str):
    min_col, min_row, max_col, max_row = range_boundaries(intervalo_str)
    return min_row -1, max_row, min_col -1, max_col # Ajustar para 0-index

# Planilha: formatação: remover linhas de grade
def formatar_remover_linhas_grade(sheet_id: int):
    return {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'gridProperties': {
                    'hideGridlines': True
                }
            },
            'fields': 'gridProperties.hideGridlines'
        }
    }

# Planilha: formatação: cor de fundo do cabeçalho
def formatar_fundo_cabecalho(sheet_id: int, cabecalho_intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(cabecalho_intervalo)
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row, 
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo conforme especificação da API Google Sheets
            },
            'fields': 'userEnteredFormat.backgroundColor',
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': DEFAULT_BACKGROUND_COLOR
                            }
                        }
                    ] * (end_col - start_col) # Usar end_col - start_col para quantidade de colunas
                }
            ]
        }
    }

# Planilha: formatação: bordas num intervalo
def formatar_bordas(sheet_id: int, intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(intervalo)
    return {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo
            },
            'top': {'style': 'SOLID', 'width': 1},
            'bottom': {'style': 'SOLID', 'width': 1},
            'left': {'style': 'SOLID', 'width': 1},
            'right': {'style': 'SOLID', 'width': 1},
            'innerHorizontal': {'style': 'SOLID', 'width': 1},
            'innerVertical': {'style': 'SOLID', 'width': 1}
        }
    }

# Planilha: formatação: renomear aba
def formatar_renomear_aba(sheet_id: int, nova_aba_nome: str):
    return {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'title': nova_aba_nome
            },
            'fields': 'title'
        }
    }

# Planilha: formatação: centralizar conteúdo de um intervalo
def formatar_centralizar_conteudo(sheet_id: int, intervalo: str):
    start_row, end_row, start_col, end_col = celula_intervalo_para_linhas_colunas(intervalo)
    return {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col # endColumnIndex é exclusivo
            },
            'fields': 'userEnteredFormat.horizontalAlignment',
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'horizontalAlignment': 'CENTER'
                            }
                        }
                    ] * (end_col - start_col)
                }
            ]
        }
    }

# Planilha: formatação: autoajuste de colunas de um intervalo
def formatar_auto_ajustar_colunas(sheet_id: int, start_col: int, end_col: int):
    return {
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': start_col,
                'endIndex': end_col # endIndex é exclusivo
            }
        }
    }

# Aplica formatações desejadas a uma planilha do Google Sheets.
def aplicar_formatacoes_planilha(service_sheets, planilha_id: str, aba_nome: str, cabecalho_intervalo: str, dados_intervalo: str):
    # Variável para acrescentar as requisições de formatações desejadas
    requisicao = []
    sheet_id = id_aba_planilha_por_nome(service_sheets, planilha_id, aba_nome, False)

    _, _, start_col_cabecalho, _ = celula_intervalo_para_linhas_colunas(cabecalho_intervalo)
    _, _, _, end_col_dados       = celula_intervalo_para_linhas_colunas(dados_intervalo)

    requisicao.append(formatar_remover_linhas_grade(sheet_id))
    requisicao.append(formatar_fundo_cabecalho(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_bordas(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_bordas(sheet_id, dados_intervalo))
    requisicao.append(formatar_centralizar_conteudo(sheet_id, cabecalho_intervalo))
    requisicao.append(formatar_auto_ajustar_colunas(sheet_id, start_col_cabecalho, end_col_dados))

    try:
        body = {'requests': requisicao}
        response = service_sheets.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body=body
        ).execute()
        mensagem(1,"Formatação aba: linhas de grades, bordas, títulos das colunas.")
        return response
    except HttpError as e:
        mensagem(1,f"Formatação aba: HttpError: {e}", 'e')
        return None
    
# Constrói uma string de intervalo de células do Google Sheets.
# Retorna a string do intervalo de células (ex: 'A1:C1').
def planilha_celulas_intervalo(letra_inicial: str, linha_inicial: int, conteudo: list, tipo: str):
    if not conteudo or not isinstance(conteudo, list) or not conteudo[0]:
        mensagem(1,"Função 'planilha_celulas_intervalo': parâmetro 'conteudo' incorreto ou vazio.", 'e')
        raise ValueError("Parâmetro 'conteudo' inválido.")

    num_colunas = len(conteudo[0])
    num_linhas  = len(conteudo)

    coluna_inicial_numero = column_index_from_string(letra_inicial.upper())
    ultima_coluna_numero  = coluna_inicial_numero + num_colunas - 1
    ultima_coluna_letra   = get_column_letter(ultima_coluna_numero)

    if tipo.upper() == 'C':
        # Intervalo para o cabeçalho (apenas uma linha)
        intervalo = f'{letra_inicial.upper()}{linha_inicial}:{ultima_coluna_letra}{linha_inicial}'
    elif tipo.upper() == 'D':
        # Intervalo para os dados (começa na linha seguinte ao cabeçalho ou linha_inicial se sem cabeçalho)
        linha_final = linha_inicial + num_linhas - 1
        intervalo = f'{letra_inicial.upper()}{linha_inicial}:{ultima_coluna_letra}{linha_final}'
    else:
        mensagem(1,f"Função 'planilha_celulas_intervalo': parâmetro 'tipo' inválido '{tipo}'. Use 'C' para Cabeçalho ou 'D' para Dados.", 'e')
        raise ValueError("Parâmetro 'tipo' inválido.")
    
    return intervalo

def processar_eventos_e_pessoas(sheets):
    #=============================================================================================
    # variáveis para montagem da planilha
    #=============================================================================================
    eventos_sensibilizacao = {}
    eventos_prospeccao = {}
    eventos_qualificacao = {}
    pessoas_sensibilizadas = {}
    pessoas_prospectadas = {}
    pessoas_qualificadas = {}
    #=============================================================================================
    # Eventos
    #=============================================================================================
    for sheet in sheets:
        # aba conteúdo total
        aba_total = sheet.get_all_values()
        # Cabeçalhos: primeira linha [0]
        cabecalho = aba_total[0]
        # Dados: segunda linha em diante [1:]
        linhas = aba_total[1:]
        # DataFrama definido com 'cabecalho' e 'linhas'
        dados = pd.DataFrame(linhas, columns=cabecalho)
        if sheet.title == 'Dados de Inscrições em Eventos':
           # Convertendo a coluna 'Data do Evento' para o formato de data, caso não esteja
           dados['Data do Evento'] = pd.to_datetime(dados['Data do Evento'], errors='coerce', dayfirst=True)
           # Removendo itens duplicados
           resultado1 = dados[['Data do Evento', 'Evento']].drop_duplicates()
           # Contabilizando os eventos por data (mantendo a descrição do evento e data) e definindo a coluna 'Ano'
           resultado2 = resultado1.groupby([resultado1['Data do Evento'].dt.year.rename('Ano'), 'Evento']).size().reset_index(name='Total de Eventos')
           # Agrupando por ano e somando a quantidade total de eventos por ano
           resultado3 = resultado2.groupby('Ano')['Total de Eventos'].sum().reset_index(name='Total de Eventos por Ano')
           for index, row in resultado3.iterrows():
               ano = row['Ano']
               eventos_sensibilizacao[ano] = eventos_sensibilizacao.get(ano, 0) + row['Total de Eventos por Ano']
        if sheet.title == 'Dados de Prospecção e Qualificação':
           # Convertendo a coluna 'Data do Evento' para o formato de data, caso não esteja
           dados['Data do Evento:'] = pd.to_datetime(dados['Data do Evento:'], errors='coerce', dayfirst=True)

           # Extraindo o ano da coluna 'Data do Evento:'
           dados['Ano do Evento'] = dados['Data do Evento:'].dt.year

           # Filtrando para manter apenas os eventos do tipo 'Prospecção' (exclusivos)
           df_prospec = dados[dados['Tipo do Evento:'] == 'Prospecção'].drop_duplicates(subset=['Nome do Evento:', 'Data do Evento:'])

           # Contabilizando o número de eventos 'Prospecção' por ano
           resultado = df_prospec.groupby(['Ano do Evento']).size().reset_index(name='Quantidade')
           for index, row in resultado.iterrows():
               ano = row['Ano do Evento']
               eventos_prospeccao[ano] = eventos_prospeccao.get(ano, 0) + row['Quantidade']

           # Filtrando para manter apenas os eventos do tipo 'Qualificação' (exclusivos)
           df_prospec = dados[dados['Tipo do Evento:'] == 'Qualificação'].drop_duplicates(subset=['Nome do Evento:', 'Data do Evento:'])

           # Contabilizando o número de eventos 'Prospecção' por ano
           resultado = df_prospec.groupby(['Ano do Evento']).size().reset_index(name='Quantidade')
           for index, row in resultado.iterrows():
               ano = row['Ano do Evento']
               eventos_qualificacao[ano] = eventos_qualificacao.get(ano, 0) + row['Quantidade']

    #=============================================================================================
    # Pessoas
    #=============================================================================================
    dados_satisfacao = []
    dados_prospeccao = []
    for sheet in sheets:
        # conteúdo da aba
        aba_total = sheet.get_all_values()
        # Cabeçalhos: primeira linha [0]
        cabecalho = aba_total[0]
        # Dados: segunda linha em diante [1:]
        linhas = aba_total[1:]
        # DataFrama definido com 'cabecalho' e 'linhas'
        dados = pd.DataFrame(linhas, columns=cabecalho)
        if sheet.title == 'Dados de Inscrições em Eventos':
           # Obter as combinações únicas de 'Data do Evento' e 'Pessoas'
           dados['Data do Evento'] = pd.to_datetime(dados['Data do Evento'], errors='coerce', dayfirst=True)
           # Verificar se há valores nulos em 'Data do Evento' após a conversão
           if dados['Data do Evento'].isnull().any():
              mensagem(1,"Atenção: Existem valores inválidos em 'Data do Evento'", 'e')
           # Remover linhas onde 'Data do Evento' ou 'Pessoas' são vazias
           dados = dados.dropna(subset=['Data do Evento', 'Pessoas'])
           # Remover entradas onde 'Pessoas' são strings vazias (se necessário)
           dados = dados[dados['Pessoas'].str.strip() != '']
           resultado1 = dados[['Data do Evento', 'Pessoas']].drop_duplicates(subset=['Data do Evento', 'Pessoas'])
           resultado1['Ano'] = resultado1['Data do Evento'].dt.year
           # Remover entradas onde 'Ano' são strings vazias (se necessário)
           resultado1 = resultado1.dropna(subset=['Ano'])
           resultado1 = resultado1[resultado1['Ano'] != 0]
           if resultado1['Ano'].isnull().any():
              mensagem(1,"Atenção: Existem valores inválidos em 'Ano'", 'e')
           resultado2 = resultado1.groupby('Ano')['Pessoas'].count().to_dict()
           # Converter o dicionário resultado3 para uma lista de listas no formato [ano, quantidade]
           resultado3 = [[ano, quantidade] for ano, quantidade in resultado2.items()]
           for ano, quantidade in resultado3:
               ano = int(ano)
               quantidade = int(quantidade)
               pessoas_sensibilizadas[ano] = pessoas_sensibilizadas.get(ano, 0) + quantidade

        if sheet.title == 'Dados de Satisfação em Eventos':
           dados_satisfacao = dados
        if sheet.title == 'Dados de Prospecção e Qualificação':
           dados_prospeccao = dados

    if not dados_satisfacao.empty and not dados_prospeccao.empty:
        #dicionario_eventos = []
        dicionario_eventos = {}
        for index, linha in dados_prospeccao.iterrows():
            data_evento = linha.get('Data do Evento:', '').strip() # Usando .get() para evitar KeyError
            nome_evento = linha.get('Nome do Evento:', '').strip()
            tipo_evento = linha.get('Tipo do Evento:', '').strip()
            if data_evento and nome_evento:
               dicionario_eventos[(data_evento, nome_evento, tipo_evento)] = True 
        #=========================================================================================================
        # Pessoas: prospecção
        #=========================================================================================================
        pessoas_prospeccao_eventos_encontrados = []
        pessoas_prospeccao_resultado_anual = {}  # Dicionário para armazenar a contagem de eventos por ano
        pessoas_prospeccao_ano_contagem = {}
        for index, linha in dados_satisfacao.iterrows():
            data_satisfacao = linha.get('Data do evento', '').strip()
            evento_satisfacao = linha.get('Evento', '').strip()
            email = linha.get('E-mail', '').strip()
            if (data_satisfacao, evento_satisfacao, "Prospecção") in dicionario_eventos:
               pessoas_prospeccao_eventos_encontrados.append((email, data_satisfacao, evento_satisfacao))
               try:
                  # Converte a string da data para um objeto datetime e extrai o ano
                  ano = pd.to_datetime(data_satisfacao, dayfirst=True).year
                  # Incrementa a contagem para aquele ano
                  pessoas_prospeccao_ano_contagem[ano] = pessoas_prospeccao_ano_contagem.get(ano, 0) + 1
               except ValueError:
                  mensagem(1,f"Aviso: Não foi possível extrair o ano da data: {data_satisfacao}", 'e')
        pessoas_prospeccao_resultado_anual = [[ano, count] for ano, count in sorted(pessoas_prospeccao_ano_contagem.items())]
        for ano, quantidade in pessoas_prospeccao_resultado_anual:
            ano = int(ano)
            quantidade = int(quantidade)
            pessoas_prospectadas[ano] = pessoas_prospectadas.get(ano, 0) + quantidade
        #=========================================================================================================
        # Pessoas: qualificação
        #=========================================================================================================
        pessoas_qualificacao_eventos_encontrados = []
        pessoas_qualificacao_resultado_anual = {}  # Dicionário para armazenar a contagem de eventos por ano
        pessoas_qualificacao_ano_contagem = {}
        for index, linha in dados_satisfacao.iterrows():
            data_satisfacao = linha.get('Data do evento', '').strip()
            evento_satisfacao = linha.get('Evento', '').strip()
            email = linha.get('E-mail', '').strip()
            if (data_satisfacao, evento_satisfacao, "Qualificação") in dicionario_eventos:
               pessoas_qualificacao_eventos_encontrados.append((email, data_satisfacao, evento_satisfacao))
               try:
                  # Converte a string da data para um objeto datetime e extrai o ano
                  ano = pd.to_datetime(data_satisfacao, dayfirst=True).year
                  # Incrementa a contagem para aquele ano
                  pessoas_qualificacao_ano_contagem[ano] = pessoas_qualificacao_ano_contagem.get(ano, 0) + 1
               except ValueError:
                  mensagem(1,f"Aviso: Não foi possível extrair o ano da data: {data_satisfacao}", 'e')
        pessoas_qualificacao_resultado_anual = [[ano, count] for ano, count in sorted(pessoas_qualificacao_ano_contagem.items())]
        for ano, quantidade in pessoas_qualificacao_resultado_anual:
            ano = int(ano)
            quantidade = int(quantidade)
            pessoas_qualificadas[ano] = pessoas_qualificadas.get(ano, 0) + quantidade

    #=============================================================================================
    # Planilha com resultados
    #=============================================================================================
    # Cabeçalho da tabela
    cabecalho_planilha = ['Ano', 'Eventos: Sensibilização', 'Eventos: Prospecção', 'Eventos: Qualificação',
        'Pessoas: Sensibilizadas', 'Pessoas: Prospectadas', 'Pessoas: Qualificadas']
    
    # Lista com as contagens por ano:
    anos = sorted(set(eventos_sensibilizacao.keys()).union(eventos_prospeccao.keys(), eventos_qualificacao.keys()))
    dados_planilha = []
    for ano in anos:
        eventos_sens = eventos_sensibilizacao.get(ano, 0)
        eventos_prosp = eventos_prospeccao.get(ano, 0)
        eventos_qual = eventos_qualificacao.get(ano, 0)
        pessoas_sens = pessoas_sensibilizadas.get(ano, 0)
        pessoas_prosp = pessoas_prospectadas.get(ano, 0)
        pessoas_qual = pessoas_qualificadas.get(ano, 0)
        dados_planilha.append([int(ano), int(eventos_sens), int(eventos_prosp), int(eventos_qual),
                            int(pessoas_sens), int(pessoas_prosp), int(pessoas_qual)])
    # Criando o DataFrame com os resultados
    dados_planilha = pd.DataFrame(dados_planilha, columns=cabecalho_planilha)
    #dados_planilha = pd.DataFrame(dados_planilha)
    return dados_planilha, cabecalho_planilha

def criar_subpasta_planilha(service_drive, service_sheets):
    pasta_id = link_id(service_drive, os.getenv('PASTA_COMPARTILHADA'))
    if not pasta_id:
       mensagem(0,f"")
       mensagem(0,'Programa interrompido.', 'c')
       mensagem(0,f"")
       sys.exit() 

    subpasta_id = ''
    if os.getenv('SUB_PASTA'):
       subpasta_id = criar_pasta(service_drive, os.getenv('SUB_PASTA'), pasta_id)
    
    # Valida existência da variável PLANILHA no .env, necessária para nomear a nova planilha
    if not os.getenv('PLANILHA'):
        mensagem(1,"Erro: Variável PLANILHA não encontrada no .env", 'c')
        sys.exit(1)
    
    planilha_nome = os.getenv('PLANILHA')
    planilha_id   = criar_planilha(service_drive, service_sheets, planilha_nome, subpasta_id)
    return planilha_id

def preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho, dados, aba):

    intervalo_cabecalho, intervalo_dados = preparar_intervalos(cabecalho, dados)

    try:
        id_aba_planilha_por_nome(service_sheets, planilha_id, aba, True)
        planilha_aba_limpeza(service_sheets, planilha_id, aba)
        # Insere os dados na nova planilha e aplica formatações (borda, cabeçalho, colunas)
        planilha_dados(service_sheets, planilha_id, aba, intervalo_cabecalho, [cabecalho])
        planilha_dados(service_sheets, planilha_id, aba, intervalo_dados, dados.values.tolist())
        aplicar_formatacoes_planilha(service_sheets, planilha_id, aba, intervalo_cabecalho, intervalo_dados)
    except HttpError as error:
        mensagem(1,f"Ocorreu um erro: {error}", 'c')

def mensagem(nivel: int, mensagem: str, log_tipo='i'):
    # log_tipo: d - debug, i - informação, w - aviso, e - erro, c - erro crítico
    # A função gera uma string com `nivel` tabulações
    tabulacao = '\t' * nivel
    # Imprime a mensagem com a tabulação correta
    print(f"{tabulacao}{mensagem}")
    if log_tipo.lower() == 'd':
       logging.debug(f"{tabulacao}{mensagem}") 
    elif log_tipo.lower() == 'i':
       #print(f"conferindo: {log_tipo} = {tabulacao}{mensagem}")
       logging.info(f"{tabulacao}{mensagem}") 
    elif log_tipo.lower() == 'w':
       logging.warning(f"{tabulacao}{mensagem}") 
    elif log_tipo.lower() == 'e':
       logging.error(f"{tabulacao}{mensagem}") 
    elif log_tipo.lower() == 'c':
       logging.critical(f"{tabulacao}{mensagem}") 

def orientacoes():
    mensagem(1, "Utilize o link abaixo para utilizar orientações de uso:", 'c')
    mensagem(1, "https://drive.google.com/drive/folders/1H3tHrMOHu-J4iQAXnf69AHi2RWOjW5WW", 'c')
#==============================================================================================
# Bibliotecas
#==============================================================================================
import os
import sys
# chamada de arquivo com métodos para o processamento
import funcoes
import pandas as pd
from dotenv import load_dotenv
# capturar uma exceção HttpError
from googleapiclient.errors import HttpError 

#==============================================================================================
# Início da execução: agendar tarefa, configurar log, acessar serviços (drive, sheets, cliente)
#                     percorrer planilhas, processar dados, gravar em nova planilha e formatar.
#==============================================================================================
# Carrega variáveis de ambiente do arquivo .env (URLs, IDs, nomes de planilhas, etc.)

# Ativar o log
funcoes.configurar_log()

funcoes.mensagem(0, "")
funcoes.mensagem(0, "Capacita Brasil: Ciência de Dados: Equipe 8_5: Processamento de indicadores.")

dotenv_path = os.path.join(os.getcwd(), '.env')
# Verifica se o arquivo '.env' existe antes de tentar carregá-lo
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
    funcoes.mensagem(0, "")
    funcoes.mensagem(0, "Arquivo '.env' carregado com sucesso.")
else:
    # Caso o arquivo .env não seja encontrado, você pode logar um erro crítico
    # ou até mesmo encerrar o programa se ele não puder funcionar sem as variáveis de ambiente.
    funcoes.mensagem(1, "Erro: O arquivo '.env' não foi encontrado no diretório atual.", 'c')
    funcoes.mensagem(1, "Por favor, certifique-se de que o '.env' está na mesma pasta do executável.", 'c')
    funcoes.orientacoes()
    sys.exit(1) # Importe 'sys' se for usar sys.exit()

# agendar tarefa
# nome do aplicativo a ser executado
tarefa_executar = os.getenv('tarefa_executavel')

# nome da tarefa no agendador de tarefas
tarefa_nome = os.getenv('tarefa_nome')

# opções 'Minuto' ou 'Hora'
tarefa_intervalo = os.getenv('tarefa_tipo')

# tempo para execução: 2 horas ou 2 minutos, conforme opção em 'tarefa_intervalo'
tarefa_tempo = os.getenv('tarefa_tempo')

# ocultar = False, True ou 0, 1
tarefa_ocultar = False
tarefa_ocultar_str = os.getenv('tarefa_ocultar')
if tarefa_ocultar_str is not None and tarefa_ocultar_str.lower() in ['true', '1']:
   tarefa_ocultar = True 

#funcoes.remover_tarefa(tarefa_nome)
# Verificando os parâmetros de agendamento (opcional)
if not tarefa_executar or not tarefa_nome or not tarefa_intervalo or not tarefa_tempo:
   funcoes.mensagem(0,"")
   funcoes.mensagem(0,"Execução manual.",'i')
else:    
   funcoes.mensagem(0,"")
   funcoes.mensagem(0,"Verificando tarefa agendada...",'i')
   funcoes.agendar_tarefa(tarefa_nome, tarefa_executar, tarefa_intervalo, tarefa_tempo, tarefa_ocultar)

# Verifica conexão com a internet. Encerra o programa se estiver offline.
if not funcoes.verificar_conexao():
   funcoes.mensagem(1,"")
   funcoes.mensagem(1,f"Sem conexão com a internet.", 'c')
   funcoes.mensagem(1,f"Programa interrompido.", 'c')
   sys.exit()

# Conecta às APIs do Google Drive, Sheets e gspread usando credenciais do serviço
try:
   service_drive, service_sheets, cliente = funcoes.conectar_google_apis()
except SystemExit:
   funcoes.mensagem(1,"Falha na autenticação.", 'c')
   funcoes.mensagem(1,"Programa interrompido.", 'c')
   sys.exit(1)

#sys.exit(1)

# Dicionário contendo os nomes e URLs dos bancos de dados
bancos = {
'Banco 1': os.getenv('BANCO_1_URL'),
'Banco 2': os.getenv('BANCO_2_URL'),
'Banco 3': os.getenv('BANCO_3_URL'),
'Banco 4': os.getenv('BANCO_4_URL'),
'Banco 5': os.getenv('BANCO_5_URL'),
'Banco 6': os.getenv('BANCO_6_URL')
}

try:
    planilha_id = funcoes.criar_subpasta_planilha(service_drive, service_sheets)
    for banco, url in bancos.items():
        try:
            if not url:  # Verifica se o conteúdo de 'url' está vazio
               raise ValueError("URL está vazia.")  # Levanta uma exceção personalizada se estiver vazio
            planilha = cliente.open_by_url(url)
        except ValueError as e:  # Exceção personalizada para conteúdo vazio
            funcoes.mensagem(0,"")
            funcoes.mensagem(0,f"Falha na abertura do {banco}: {e}",'c')
            continue           
        except HttpError as e:
           funcoes.mensagem(0,"")
           funcoes.mensagem(0,f"Falha na abertura do {banco} 'HttpError': {e}",'c')
           continue
        except Exception as e:
           funcoes.mensagem(0,"")
           funcoes.mensagem(0,f"Falha na abertura do {banco} 'Exception': {e}",'c')
           continue
        
        sheets   = planilha.worksheets()
        funcoes.mensagem(0,f"")
        funcoes.mensagem(0,f"{banco}: {planilha.title}")
        funcoes.mensagem(1,f"link: {url}")
        if 'banco 1' in planilha.title.lower():
            funcoes.mensagem(1,f"Processando informações...")
            dados, cabecalho = funcoes.processar_eventos_e_pessoas(sheets)
            funcoes.preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho, dados, banco)
        if 'banco 2' in planilha.title.lower():
            dados, cabecalho = funcoes.carregar_dados_planilha(planilha, 'Dados Seleção')
            funcoes.preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho, dados, banco)
        if 'banco 4' in planilha.title.lower():
            dados_cons, cabecalho_cons = funcoes.carregar_dados_planilha(planilha, "Banco de Consultorias")
            dados_ment, cabecalho_ment = funcoes.carregar_dados_planilha(planilha, "Banco de Mentorias")
            if dados_cons.empty and dados_ment.empty:
               funcoes.mensagem(1,"Nenhum dado carregado das abas 'Consultoria' e 'Mentoria'. Encerrando.", 'w')
               continue
            dados = pd.concat([dados_cons, dados_ment], ignore_index=True)
            # Remove linhas totalmente vazias
            dados = dados.dropna(how='all')  
            funcoes.mensagem(1,f"Quantidade de registros (mentoria+consultoria): {len(dados)}")
            funcoes.preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho_cons, dados, banco)
        if 'banco 5' in planilha.title.lower():
            dados, cabecalho = funcoes.carregar_dados_planilha(planilha, "Versão resumida")
            funcoes.preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho, dados, banco)
        if 'banco 6' in planilha.title.lower():
            # Lê os dados e cabeçalhos das abas "Banco de Parceiro" e "Impactos de Marketing"
            dados_parc, cabecalho_parc = funcoes.carregar_dados_planilha(planilha, "Banco de Parceiro")
            dados_mark, cabecalho_mark = funcoes.carregar_dados_planilha(planilha, "Impactos de Marketing")
            if dados_parc.empty and dados_mark.empty:
                funcoes.mensagem(1,"Nenhum dado encontrado. Encerrando.")
                continue
            # Junta os dados das duas abas
            dados = pd.concat([dados_parc, dados_mark], ignore_index=True)
            # Remove linhas totalmente vazias
            dados = dados.dropna(how='all')  
            funcoes.mensagem(1,f"Quantidade de registros (Parcerias+Marketing): {len(dados)}")
            funcoes.preencher_formatar_planilha(service_drive, service_sheets, planilha_id, cabecalho_parc, dados, banco)
    funcoes.mensagem(0,f"")
    funcoes.mensagem(0,f"Planilha criada: https://docs.google.com/spreadsheets/d/{planilha_id}")
    funcoes.mensagem(0,"Fim.")
    funcoes.mensagem(0,f"")
except HttpError as error:
   funcoes.mensagem(0,f"")
   funcoes.mensagem(0,f"Falha: HttpError: {error}", 'e')
   funcoes.mensagem(0,f"")
except Exception as e:
    funcoes.mensagem(0,f"")
    funcoes.mensagem(0,f"Falha: Exception: {e}", 'e')
    funcoes.mensagem(0,f"")
    sys.exit()

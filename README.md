# MCTI FUTURO — Ministério da Ciência, Tecnologia e Inovação  
## Capacita Brasil - Residência em TIC 20
### UECE 2025.1 - Ciência de dados - Equipe 8_5
### Processamento de Indicadores

Este projeto, desenvolvido como parte da Residência em TIC 20 do Ministério da Ciência, Tecnologia e Inovação (MCTI Futuro) - Ciência de Dados, tem como objetivo principal processar dados de diversas fontes (Google Sheets) e gerar uma planilha com dados consolidados e formatados, servindo como banco de dados otimizado para visualização e análise num dashboard no Looker Studio.

## 1. Funcionalidades

O aplicativo realiza as seguintes operações:

- **Coleta de Dados**: Lê informações de múltiplas planilhas do Google Sheets (Bancos de Dados de Leitura).
- **Processamento de Indicadores**: Processa informações, consolidando e contabilizando indicadores que serão utilizados como fonte de dados no dashboard no Looker Studio.
- **Geração da Planilha**: Cria uma nova planilha no Google Sheets (Banco de Dados do dashboard no Looker Studio) com os indicadores processados.
- **Gerenciamento de Pastas no Drive**:
    - Verifica a existência de pastas e subpastas no Google Drive.
    - Cria novas subpastas, se necessário.
    - Atribui permissões de compartilhamento automaticamente para facilitar o acesso.
- **Formatação de Planilhas**:
    - Limpa abas antes de gravar novos dados.
    - Aplica diversas formatações (remoção de linhas de grade, cor de fundo em cabeçalhos, bordas, centralização de conteúdo e autoajuste de colunas) para apresentar os dados de forma clara e profissional.
- **Automatização (Opcional)**: Prepara o ambiente para agendamento da execução do aplicativo via Agendador de Tarefas do Windows ou Cron no Linux.

## 2. Como Usar

Para configurar e executar este programa, siga os passos abaixo:

### 2.1. Pré-requisitos

Certifique-se de ter o Python instalado (versão 3.x recomendada) e as seguintes bibliotecas:

- `pandas`
- `gspread`
- `google-api-python-client`
- `google-auth-httplib2`
- `google-auth-oauthlib`
- `python-dotenv`
- `openpyxl`

Você pode instalá-las via pip:

```bash
pip install pandas gspread google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv openpyxl
```
### 2.2. Configuração das Credenciais do Google API
- 1. Habilite as APIs: No Google Cloud Console, certifique-se de que as APIs Google Drive API e Google Sheets API estejam habilitadas para o seu projeto.
- 2. Crie as Credenciais: Gere um arquivo de credenciais JSON do tipo "Conta de Serviço".
- 3. Salve o Arquivo: Transfira o arquivo JSON baixado para a pasta raiz do seu projeto ou ajuste o caminho em .env.
    
### 2.3. Configuração do Arquivo .env
Crie um arquivo chamado .env na raiz do seu projeto e preencha-o com as suas configurações. Este arquivo é crucial para personalizar o comportamento do programa sem alterar o código-fonte e está comentado para orientação de preenchimento.

### 2.4. Execução do Programa
Com todas as configurações feitas, você pode executar o programa:
```bash
python capacita-brasil_bancos-final_indicadores.py
```
O programa irá gerar logs detalhados de sua execução para que você possa acompanhar o processamento e identificar quaisquer problemas.

## 3. Criando um Executável com PyInstaller

Caso queira distribuir o programa como um executável, você pode criar um arquivo executável para Windows, Linux ou Mac usando o PyInstaller.
### 3.1. Instalando o PyInstaller

Primeiro, instale o PyInstaller com o seguinte comando:
```bash
pip install pyinstaller
```
### 3.2. Criando o Executável

Depois de instalar o PyInstaller, basta executar o seguinte comando no terminal dentro do diretório onde está o arquivo capacita-brasil_bancos-final_indicadores.py:
```bash
pyinstaller capacita-brasil_bancos-final_indicadores.py --onefile --noconsole --name "CapacitaBrasilEquipe8-5Indicadores" --icon=capacita-brasil.ico --add-data "funcoes.py;." --add-data "template_tarefa.xml;."
```
 Onde:<br>
 - --onefile, faz com que o PyInstaller crie um único arquivo executável.
 - --icon=capacita-brasil.ico, acrescenta um ícone ao arquivo executável.
### 3.3. Localização do Executável

Após a execução do comando, o PyInstaller criará uma pasta chamada dist. Dentro dessa pasta, você encontrará o arquivo executável (capacita-brasil_bancos-final_indicadores.exe no caso do Windows).
3.4. Executando o Executável

Agora você pode executar o programa diretamente, sem precisar de um ambiente Python configurado. Basta rodar o arquivo executável gerado.

### 4. Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues para reportar bugs, sugerir melhorias ou enviar pull requests com novas funcionalidades.

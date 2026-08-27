# Processamento de Indicadores — Capacita Brasil

Pipeline em Python desenvolvido na Residência em TIC 20 do MCTI Futuro / Capacita Brasil (UECE), para consolidar dados de Google Sheets e alimentar dashboards no Looker Studio.

## Objetivo

Transformar dados distribuídos em planilhas em uma fonte única, organizada e confiável para análise de indicadores organizacionais.

## Funcionalidades

- Leitura de múltiplas planilhas do Google Sheets.
- Consolidação e cálculo de indicadores com Python e Pandas.
- Criação e formatação da planilha de saída para o Looker Studio.
- Criação e organização de pastas no Google Drive.
- Atribuição automatizada de permissões de compartilhamento.
- Geração de logs para acompanhamento da execução.
- Preparação para agendamento no Windows Task Scheduler ou Cron.

## Tecnologias

- Python 3
- Pandas e OpenPyXL
- Google Sheets API e Google Drive API
- gspread
- python-dotenv
- Looker Studio

## Como executar

1. Instale as dependências:

```bash
pip install -r requirements.txt
```

2. Configure as credenciais da conta de serviço e as variáveis de ambiente.
3. Execute:

```bash
python capacita-brasil_bancos-final_indicadores.py
``

> **Segurança:** credenciais e arquivos `.env` não devem ser versionados ou compartilhados publicamente.

## Qualidade e validações

O pipeline deve ser validado com dados ausentes, duplicados, formatos incompatíveis, falhas de autenticação e reconciliação dos indicadores gerados com as fontes de origem.

## Contexto acadêmico

Projeto desenvolvido pela Equipe 8_5 da trilha de Ciência de Dados — UECE 2025.1.

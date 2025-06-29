# MCTI FUTURO — Ministério da Ciência, Tecnologia e Inovação  
## Capacita Brasil - Residência em TIC 20  
### Ciência de Dados – Equipe 8_5  
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
    - Aplica diversas formatações (remoção de linhas de grade,

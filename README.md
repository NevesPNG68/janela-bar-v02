# Janela Bar v02 — Dashboard Seguro

Este repositório será usado para a nova versão do dashboard do Janela Bar.

## Objetivo

- Manter o HTML público.
- Não expor o link da planilha.
- Não deixar a planilha original acessível pelo código-fonte.
- Gerar um arquivo de dados tratado para o dashboard.

## Modelo de segurança

```text
Planilha privada no Dropbox
↓
GitHub Actions acessa com segredo
↓
Script gera dados/dashboard-data.json
↓
index.html lê apenas os dados tratados
```

## Segredos necessários no GitHub

Em `Settings > Secrets and variables > Actions > New repository secret`, cadastrar:

```text
DROPBOX_TOKEN
DROPBOX_FILE_PATH
```

Exemplo de `DROPBOX_FILE_PATH`:

```text
/Janela/Por venda_2026.xlsm
```

Nunca colocar token ou link da planilha dentro do HTML.

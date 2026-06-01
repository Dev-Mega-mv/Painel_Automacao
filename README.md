# Ativador Pro - Megavale

Aplicacao web em Flask para automatizar ativacao, desativacao, verificacao e coleta de dados de credenciados nas tecnologias CIELO e CARDSE.

O sistema usa Selenium para acessar o portal operacional, processar contratos/produtos e exibir o andamento em tempo real pela tela de status.

## Funcionalidades

- Ativacao/desativacao individual por codigo ou CNPJ.
- Processamento em lote por planilha `.xlsx` ou `.xls`.
- Selecao de tecnologia: `CIELO`, `CARDSE` ou `AMBOS`.
- Acao independente para CIELO e CARDSE: `ATIVAR` ou `DESATIVAR`.
- Modo verificacao, sem alterar produtos, com geracao de relatorio.
- Coleta de codigo de estabelecimento por origem de transacao.
- Status em tempo real com progresso, item atual e logs.
- Download de relatorios gerados ao final do processamento.

## Estrutura

```text
.
|-- app.py                    # Aplicacao Flask, rotas e leitura de planilhas
|-- automacao_ativador.py      # Fluxo Selenium e regras de automacao
|-- templates/                # Telas HTML
|-- static/css/               # Estilos e logo
|-- uploads/                  # Planilhas enviadas pela interface
|-- selenium-profile/         # Perfil local usado pelo Chrome/Selenium
|-- ativador_log.txt          # Log operacional
`-- debug_ativador.txt        # Log tecnico de erros
```

## Requisitos

- Windows.
- Python 3.11 ou compativel.
- Google Chrome instalado.
- Acesso ao portal `https://online.fwcard.com.br/fwcard/f`.
- Credenciais validas do Infox/portal.

Dependencias Python usadas pelo projeto:

```text
flask
selenium
webdriver-manager
pandas
openpyxl
pyautogui
keyboard
werkzeug
```

## Instalacao

Crie e ative um ambiente virtual:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Instale as dependencias:

```powershell
pip install flask selenium webdriver-manager pandas openpyxl pyautogui keyboard werkzeug
```

Observacao: o arquivo `automacao_ativador.py` define um caminho fixo para o Chrome:

```python
opts.binary_location = r"C:\BACKUP_EMANUEL\Program Files\Google\Chrome\Application\chrome.exe"
```

Se o Chrome estiver em outro local, ajuste esse caminho antes de executar.

## Como executar

Com o ambiente virtual ativo:

```powershell
python app.py
```

Acesse no navegador:

```text
http://localhost:8081
```

O Flask sobe em `0.0.0.0:8081`, entao tambem pode ser acessado por outro computador da rede, se firewall e rede permitirem.

## Como executar a automacao

1. Inicie a aplicacao com `python app.py`.
2. Abra `http://localhost:8081`.
3. Escolha uma das opcoes da tela inicial:
   - `Individual`: para processar um unico credenciado.
   - `Lote`: para processar uma planilha.
   - `Coleta`: para coletar codigo de estabelecimento.
4. Informe login e senha do portal.
5. Selecione o tipo de pesquisa:
   - `Codigo`: usa o codigo do credenciado.
   - `CNPJ`: usa CNPJ com preenchimento para 14 digitos quando necessario.
6. Configure tecnologia e acao:
   - Tecnologia: `CIELO`, `CARDSE` ou `AMBOS`.
   - Acao: `ATIVAR` ou `DESATIVAR`.
7. Se estiver usando lote, envie uma planilha com a coluna obrigatoria `COD`.
8. Clique no botao de inicio da tela escolhida.
9. A aplicacao redireciona automaticamente para `Status`.
10. Aguarde o navegador Chrome automatizado abrir e realizar o login.
11. Acompanhe o progresso, item atual e logs em tempo real na tela `Status`.
12. Ao final, baixe o relatorio se o modo usado gerar arquivo.

Durante a execucao, mantenha a janela do Chrome aberta e evite usar o navegador automatizado manualmente. Para parar uma execucao em andamento, use o botao de interrupcao na tela `Status`.

## Telas

- `Inicio`: atalhos para os principais fluxos.
- `Individual`: processa um credenciado por vez.
- `Lote`: processa uma planilha com varios credenciados.
- `Coleta`: coleta codigo de estabelecimento por origem de transacao.
- `Status`: acompanha progresso, logs e download de relatorio.

## Uso individual

1. Informe login e senha.
2. Escolha o tipo de pesquisa: codigo ou CNPJ.
3. Informe o codigo/CNPJ.
4. Escolha tecnologia: `CIELO`, `CARDSE` ou `AMBOS`.
5. Escolha a acao desejada.
6. Opcionalmente, habilite modo verificacao.
7. Clique para iniciar e acompanhe pela tela de status.

## Uso em lote

A planilha deve conter obrigatoriamente a coluna:

```text
COD
```

Colunas opcionais:

```text
TECNOLOGIA
ACAO_CIELO
ACAO_CARDSE
```

Valores aceitos:

```text
TECNOLOGIA: CIELO, CARDSE, AMBOS
ACAO_CIELO: ATIVAR, DESATIVAR
ACAO_CARDSE: ATIVAR, DESATIVAR
```

Exemplo:

| COD | TECNOLOGIA | ACAO_CIELO | ACAO_CARDSE |
| --- | --- | --- | --- |
| 12345 | AMBOS | ATIVAR | ATIVAR |
| 67890 | CARDSE |  | DESATIVAR |
| 11223 | CIELO | ATIVAR |  |

Se uma coluna opcional nao existir ou uma celula estiver vazia, o sistema usa o padrao selecionado na tela.

## Modo verificacao

No modo verificacao, a automacao percorre os produtos e registra os status de CIELO e CARDSE sem fazer alteracoes.

Ao final, um relatorio `.xlsx` fica disponivel para download na tela de status.

## Coleta de codigo de estabelecimento

Use a tela `Coleta` quando precisar buscar codigo de estabelecimento por origem de transacao.

A planilha tambem deve conter a coluna `COD`. O resultado e salvo em um relatorio com nome semelhante a:

```text
coleta_codigo_estabelecimento_YYYYMMDD_HHMMSS.xlsx
```

## Logs e relatorios

- `ativador_log.txt`: historico das operacoes exibidas tambem na tela de status.
- `debug_ativador.txt`: detalhes tecnicos de erros e excecoes.
- `verificacao_*.xlsx`: relatorios do modo verificacao.
- `coleta_codigo_estabelecimento_*.xlsx`: relatorios de coleta.

## Observacoes operacionais

- Nao feche a janela do Chrome enquanto a automacao estiver rodando.
- Evite interagir manualmente com o navegador automatizado durante o processamento.
- Use a tela `Status` para acompanhar o progresso e solicitar interrupcao.
- Em caso de erro persistente, consulte `debug_ativador.txt`.
- O diretorio `selenium-profile/` guarda dados do perfil do Chrome usado pela automacao.

## Rotas principais

| Rota | Descricao |
| --- | --- |
| `/` | Tela inicial |
| `/individual` | Ativacao/desativacao individual |
| `/lote` | Processamento em lote |
| `/coleta-codigo-estabelecimento` | Coleta de codigo de estabelecimento |
| `/status` | Status visual da automacao |
| `/api/status` | Status em JSON |
| `/download_relatorio` | Download do relatorio gerado |
| `/stop` | Solicita interrupcao da automacao |

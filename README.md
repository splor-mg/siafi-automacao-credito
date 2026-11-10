# siafi-automacao-credito

Robô que registra no **SIAFI** (terminal TN3270) os remanejamentos de crédito
orçamentário pedidos em planilha. Roda em Ubuntu/WSL2, controla o emulador
`x3270`/`s3270` pela biblioteca `py3270` e pode ser acionado pelo duplo-clique
no `robo_credito.bat` ou pelo comando `/credito` do bot do Telegram.

## O que ele faz

1. **Consolida** (`consolida.py`) — lê os `.xlsm` das pastas `Remanejamentos` e
   `Remanejamentos (SEGOV_DCAF_FONTE 95_ FONTE 80)` do OneDrive, junta tudo na
   aba `PREENCHER AQUI` de `data/copia.xlsm` (cópia do `data/modelo.xlsm`) e
   move os originais para `Realizados`.
2. **Recalcula** (`excel.py`) — o Excel abre o `copia.xlsm` via COM para
   calcular a aba `ROBO`, que é toda derivada por fórmula.
3. **Valida** (`linhas.py`) — converte todas as linhas da aba `ROBO` e para
   se houver coluna vazia, `TIPO` não preenchido, orientação desconhecida ou
   anulação nos tipos que só suplementam (2 – Saldo Financeiro e 3 – Excesso
   de arrecadação).
4. **Login** no SIAFI (`login.py`, que também orquestra todo o fluxo).
5. **Análise de saldo** (`analise_saldo.py`) — consulta a dotação de cada
   linha de anulação e para se alguma não tiver saldo.
6. **Solicitações** (`fluxo_tipo_1.py` … `fluxo_tipo_4.py` e
   `utils_siafi.py`) — digita os documentos, um por UO, e lê o desfecho.
7. **Resultado** (`resultado.py`) — grava `ok`/`erro` e o número do documento
   nas colunas U e V da aba `ROBO` e salva a cópia em
   `Conferencia arquivo robo/Conferencia arquivo robo DD.MM.xlsm`.

### Quando as planilhas voltam para a fila

As planilhas voltam sozinhas de `Realizados` para a pasta de origem
(`arquivos.py`) sempre que é certo que nada foi registrado no SIAFI:

- o robô parou **antes** do envio (etapas 2 a 5); ou
- o envio quebrou no meio **sem que nenhum documento chegasse à confirmação**
  (o PF5 do `finalizar_documento`). Linha digitada e não confirmada não vira
  documento.

A partir da primeira confirmação — documento registrado, recusado ou sem
resposta lida — elas ficam em `Realizados`: devolver faria o próximo
acionamento reenviar o que já está no SIAFI. A mensagem do grupo sempre diz
onde as planilhas ficaram.

### Print da tela do SIAFI

Quando o envio para no meio, um documento é recusado ou a análise de saldo
não consegue ler o saldo, o robô copia a tela inteira do SIAFI para o log
(`tela.py`): as 24 linhas numeradas, a régua de colunas, o estado do teclado e
a posição do cursor — as mesmas coordenadas do `fill_field()`. A linha 1 da
tela vai também para o grupo; o print completo sai com `/log`. Para ver todos:

```bash
grep -A 28 "TELA DO SIAFI" data/logs/robo-*.log
```

## Estrutura

```
siafi_automacao/
├── login.py          # orquestrador (ponto de entrada)
├── consolida.py      # junta as planilhas de origem no copia.xlsm
├── excel.py          # recálculo das fórmulas do copia.xlsm pelo Excel (COM)
├── linhas.py         # conversão e validação das linhas da aba ROBO
├── analise_saldo.py  # consulta de saldo das dotações a anular
├── fluxo_tipo_N.py   # digitação do documento, por tipo de alteração (1 a 4)
├── utils_siafi.py    # finalização do documento e leitura do desfecho
├── tela.py           # print da tela do SIAFI no log, para diagnóstico
├── resultado.py      # desfecho de cada documento, gravado na aba ROBO
├── arquivos.py       # devolução das planilhas à pasta de origem
└── relato.py         # eventos (.jsonl) lidos pelo bot do Telegram
data/
├── modelo.xlsm       # versionado: macros + aba ROBO (nunca é escrito)
├── copia.xlsm        # gerado a cada execução (fora do Git)
└── logs/             # robo-*.log e relato-*.jsonl de cada execução
rodar.sh              # sequência única de execução (bat e bot usam a mesma)
robo_credito.bat      # duplo-clique no Windows -> robo.ps1 -> rodar.sh
setup.sh              # prepara uma máquina nova (WSL/Ubuntu)
```

## Configuração

Em máquina nova, rode o `setup.sh` dentro do WSL: ele instala as dependências,
clona o repositório, cria a venv e pergunta as variáveis do `.env`.

Manualmente: copie `.env.example` para `.env`, preencha, e

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Execução

```bash
bash rodar.sh
```

O `rodar.sh` pega o lock compartilhado com o robô de cota (os dois usam o
mesmo usuário do SIAFI e nunca rodam juntos), faz `git pull` na `main` e
executa o `login.py`. Códigos de saída:

| Código | Significado |
|---|---|
| 0 | Foi até o fim (pode haver documento recusado — veja o resumo) |
| 1 | Falha |
| 2 | Interrompido antes do envio (validação ou análise de saldo); nada foi enviado |
| 3 | Senha do SIAFI expirada; nada foi enviado |
| 10 | Já existe uma execução (cota ou crédito) em andamento |

## Se o Excel não recalcular

O recálculo usa uma instância escondida do Excel na máquina do robô. Se a
mensagem for *"o copia.xlsm já está aberto em outra instância do Excel"* (ou
*"A aba ROBO ficou vazia depois do recálculo"*), há um Excel segurando o
arquivo — em geral uma instância escondida que sobrou de uma execução
interrompida. Nesse estado o Excel abre o arquivo como somente leitura e não
grava o recálculo, e **toda** execução falha até o processo ser encerrado:

```powershell
# No Windows: lista as instâncias do Excel e como foram iniciadas
Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" |
    Select-Object ProcessId, CreationDate, CommandLine
```

A instância do robô aparece com `/automation -Embedding` e sem janela.
Encerre-a pelo Gerenciador de Tarefas (ou `Stop-Process -Id <PID>`) e acione
de novo.

## Telegram

O bot vive no repositório
[siafi-automacao-cota](https://github.com/splor-mg/siafi-automacao-cota)
(`siafi_automacao/bot_telegram.py`). No `/credito` ele executa o `rodar.sh`
deste repositório e monta as mensagens do grupo a partir do `.jsonl` gravado
pelo `relato()`. Os tipos de evento e seus campos são um contrato entre os
dois repositórios, conferido lá em `tests/test_contrato_relato.py` — ao
criar ou renomear um `relato(...)` aqui, rode a suíte do robô de cota.

## Testes

```bash
pip install -r requirements-dev.txt
pytest
```

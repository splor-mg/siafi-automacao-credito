# siafi-automacao-credito

Automação SIAFI para remanejamento de crédito orçamentário.

## Estrutura
siafi_automacao/
├── login.py                        # Login e conexão com o SIAFI
├── credito.py                      # Orquestrador principal
└── fluxo_remanejamento_credito.py  # Fluxo de remanejamento de crédito
data/
└── (planilha de trabalho - não versionada)

## Configuração

1. Crie o arquivo `.env` na raiz com as variáveis:
SISTEMA=
USUARIO=
SENHA=
UNIDADE_EXECUTORA=

2. Crie o ambiente virtual e instale as dependências:
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Execução

```bash
python3 siafi_automacao/credito.py
```

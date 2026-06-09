import os
from dotenv import load_dotenv
from py3270 import Emulator
from datetime import datetime
import pandas as pd
import time


load_dotenv()
sistema = os.getenv('SISTEMA')
usuario = os.getenv('USUARIO')
senha = os.getenv('SENHA')
unidade_executora = os.getenv('UNIDADE_EXECUTORA')

day = datetime.today().strftime("%d")
month = datetime.today().strftime("%m")
year = datetime.today().strftime("%Y")

# Caminho da planilha a ser executada
CAMINHO_LOCAL = '/home/guilhermemelof/code/splor-mg/siafi-automacao-credito/data/copia.xlsm'

# Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'ROBO'

em = Emulator(visible=True)  # caso queira que a tela apareça utilize visible=True
em.connect('bhmvsb.prodemge.gov.br')
em.wait_for_field()

# Preenche os dados de login
em.fill_field(19, 13, sistema, 7)
em.fill_field(20, 13, usuario, 7)
em.fill_field(21, 13, senha, 7)
em.send_enter()

# Loop: navega pelas telas até encontrar a mensagem de sucesso
max_tentativas = 10
tentativas = 0

while tentativas < max_tentativas:
    time.sleep(1)

    try:
        em.send_enter()

        if em.string_found(1, 13, 'Logon executado com sucesso'):
            print("Login realizado com sucesso!")
            break

        else:
            print(f"Tentativa {tentativas + 1} - tela intermediária, avançando...")
            em.send_enter()

    except:
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

em.fill_field(1, 2, sistema, 4)
em.send_enter()

# nova tela buscando login...
max_tentativas = 10
tentativas = 0

while tentativas < max_tentativas:
    time.sleep(1)

    try:
        em.send_enter()

        if em.string_found(22, 11, 'Unidade Executora'):
            print("Texto encontrado")
            break

        else:
            print(f"Tentativa {tentativas + 1} - tela intermediária, avançando...")
            em.send_enter()

    except:
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

# Entrar com a Unidade Executora
em.fill_field(22, 30, unidade_executora, 7)
em.send_enter()
em.wait_for_field()

# Entrar em consultas -> movimentação orçamentária ->
em.fill_field(21, 19, '09', 2)
em.send_enter()
em.wait_for_field()

em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()

em.fill_field(21, 19, '07', 2)
em.fill_field(21, 41, '1', 1)
em.send_enter()
em.wait_for_field()

# Leitura da Planilha
df = pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME)
df = df.dropna(how='all')  # remove linhas completamente vazias
df = df.reset_index(drop=False)

# Loop para processar cada linha da planilha
for idx, row in df.iterrows():

    # Pula linhas onde UO_COD está vazio (fim da planilha ou linha vazia)
    if pd.isna(row['UO_COD']):
        continue

    data_row = {}
    data_row['orientacao']  = str(row['ORIENTACAO']).strip()
    data_row['uo']          = str(int(row['UO_COD']))
    data_row['acao']        = str(int(row['ACAO_COD']))
    data_row['funcao']    = str(int(row['FUNCAO_COD'])).zfill(2)   # 4  → "04"
    data_row['subfuncao'] = str(int(row['SUBFUNCAO_COD'])).zfill(3) # 12 → "012"
    data_row['programa']  = str(int(row['PROGRAMA_COD'])).zfill(3)  # 7  → "007"
    data_row['subprojeto']  = str(int(row['SUBPROJETO_COD']))
    data_row['categoria']   = str(int(row['CATEGORIA_COD']))
    data_row['grupo']       = str(int(row['GRUPO_COD']))
    data_row['modalidade']  = str(int(row['MODALIDADE_COD']))
    data_row['elemento']    = str(int(row['ELEMENTO_COD']))
    data_row['iag']         = str(int(row['IPG_COD']))
    data_row['fonte']       = str(int(row['FONTE_COD']))
    data_row['procedencia'] = str(int(row['IPU_COD']))
    data_row['valor']       = float(row['VALOR'] * -1 / 100)

    if data_row['uo'] == 0:
        break

    if data_row['orientacao'] == 'Anular':
        em.fill_field(9, 75, data_row['uo'], 4)
        em.send_enter()
        em.wait_for_field()

        # verifica dotação
        em.fill_field(12, 67, data_row['funcao'], 2)
        em.fill_field(13, 67, data_row['subfuncao'], 3)
        em.fill_field(14, 67, data_row['programa'], 3)
        em.fill_field(15, 67, data_row['acao'], 4)
        em.fill_field(16, 67, "0001", 4)
        em.fill_field(17, 67, data_row['grupo'], 1)
        em.fill_field(17, 69, data_row['modalidade'], 2)
        em.fill_field(17, 72, data_row['iag'], 1)
        em.fill_field(18, 67, data_row['fonte'], 2)
        em.fill_field(18, 70, data_row['procedencia'], 1)
        em.send_enter()
        em.wait_for_field()
        time.sleep(0.5)

        retorno = em.string_get(1, 1, 80).strip()

        if retorno == '0101-CREDITOS AUTORIZ./COTA DESCENTRALIZADA INEXISTENTE(S).':
            print(f" UO: {data_row['uo']}; Ação: {data_row['acao']}; Grupo: {data_row['grupo']}; Modalidade: {data_row['modalidade']}; Fonte: {data_row['fonte']}; Procedência: {data_row['procedencia']}; Valor: {data_row['valor']}:")
            print("Dotação INEXISTENTE")
            print()
            em.send_pf(3)
            em.wait_for_field()
            time.sleep(1)
            em.fill_field(21, 19, '07', 2)
            em.fill_field(21, 41, '1', 1)
            em.send_enter()
            em.wait_for_field()
        else:
            # Número alinhado à direita terminando na coluna 78.
            # Lê 29 caracteres a partir da coluna 50 e pega o último token,
            # cobrindo qualquer tamanho de valor (de "1,00" a "1.000.000.000.000,00").
            saldo_raw = em.string_get(13, 50, 29)

            try:
                saldo_str = saldo_raw.split()[-1]
                saldo = float(saldo_str.replace('.', '').replace(',', '.'))
            except (ValueError, IndexError):
                print(f"Erro ao converter saldo lido: {repr(saldo_raw)}")
                saldo = 0.0

            linha_log = (
                f" UO: {data_row['uo']}; Ação: {data_row['acao']}; Grupo: {data_row['grupo']}; "
                f"Modalidade: {data_row['modalidade']}; Fonte: {data_row['fonte']}; Procedência: {data_row['procedencia']}; "
                f"Valor: {data_row['valor']}; Saldo: {saldo}"
            )

            if saldo >= data_row['valor']:
                print(linha_log)
                print("Ok")
            else:
                print(linha_log)
                print("Saldo INSUFICIENTE")
            print()

            em.send_pf(3)
            em.wait_for_field()
            
            em.fill_field(21, 19, '07', 2)
            em.fill_field(21, 41, '1', 1)
            em.send_enter()
            em.wait_for_field()
    else:
        print("Linha Suplementar OK")
        print()


print()
print("Fluxo concluído.")

em.terminate()
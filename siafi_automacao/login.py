import os
import shutil
from dotenv import load_dotenv
from py3270 import Emulator
from datetime import datetime
import pandas as pd
import openpyxl
import time
from fluxo_tipo_1 import tipo_1
from fluxo_tipo_2 import tipo_2
from fluxo_tipo_3 import tipo_3
from fluxo_tipo_4 import tipo_4
from utils_siafi import finalizar_documento


load_dotenv()
sistema = os.getenv('SISTEMA')
usuario = os.getenv('USUARIO')
senha = os.getenv('SENHA')
unidade_executora = os.getenv('UNIDADE_EXECUTORA')

day = datetime.today().strftime("%d")
month = datetime.today().strftime("%m")
year = datetime.today().strftime("%Y")

# Caminho da planilha a ser executada
CAMINHO_LOCAL     = '/home/guilhermemelof/code/splor-mg/siafi-automacao-credito/data/copia.xlsm'

#Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'ROBO'

em = Emulator(visible=True) ##caso queira que a tela apareça utilize visible=True
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

        # Tela COM campo editável — verifica se é a tela de sucesso
        if em.string_found(1, 13, 'Logon executado com sucesso'):
            print("Login realizado com sucesso!")
            break

        else:
            # Tela com campo editável, mas ainda não é a de sucesso
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

##nova tela buscando login...
max_tentativas = 10
tentativas = 0

while tentativas < max_tentativas:
    time.sleep(1)

    try:
        em.send_enter()

        # Tela COM campo editável — verifica se é a tela de sucesso
        if em.string_found(22, 11, 'Unidade Executora'):
            print("Texto encontrado")
            break

        else:
            # Tela com campo editável, mas ainda não é a de sucesso
            print(f"Tentativa {tentativas + 1} - tela intermediária, avançando...")
            em.send_enter()

    except:
        # Tela SEM campo editável — é a tela de aviso, só dá Enter e segue
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

#Entrar com a Unidade Executora
em.fill_field(22, 30, unidade_executora, 7)
em.send_enter()
em.wait_for_field()
# Fim do login

# Leitura da Planilha
df = pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME)
df = df.dropna(how='all')  # remove linhas completamente vazias
df = df.sort_values(by=['TIPO', 'UO_COD', 'ORIENTACAO'], ascending=[True, True, False]) # ordena por anulação e depois por UO
df = df.reset_index(drop=False)

# Definição de variáveis para controle do fluxo
verifica_tipo = 0
conclusao = 0

# Loop para processar cada linha da planilha
for idx, row in df.iterrows():

        # Pula linhas onde UO_COD está vazio (fim da planilha ou linha vazia)
    if pd.isna(row['UO_COD']):
        continue

        
    data_row = {}
    data_row['month']   = month
    data_row['day']     = day
    data_row['year']    = year
    data_row['orientacao']    = str(row['ORIENTACAO']).strip()
    data_row['uo']            = str(int(row['UO_COD']))
    data_row['acao']          = str(int(row['ACAO_COD']))
    data_row['funcao']        = str(int(row['FUNCAO_COD']))
    data_row['subfuncao']     = str(int(row['SUBFUNCAO_COD']))
    data_row['programa']      = str(int(row['PROGRAMA_COD']))
    data_row['subprojeto']    = str(int(row['SUBPROJETO_COD']))
    data_row['categoria']     = str(int(row['CATEGORIA_COD']))
    data_row['grupo']         = str(int(row['GRUPO_COD']))
    data_row['modalidade']    = str(int(row['MODALIDADE_COD']))
    data_row['elemento']      = str(int(row['ELEMENTO_COD']))
    data_row['iag']           = str(int(row['IPG_COD']))
    data_row['fonte']         = str(int(row['FONTE_COD']))
    data_row['procedencia']   = str(int(row['IPU_COD']))

    if data_row['orientacao'] == 'Anular': # se for anulação, o valor deve ser multiplicado por -1 para ficar negativo
        data_row['valor']      = str(-int(row['VALOR']))
    else:
         data_row['valor']      = str(int(row['VALOR']))

    data_row['uo_suplementada'] = str(int(row['UO_SUPLEMENTADA'])) if pd.notna(row['UO_SUPLEMENTADA']) else '0'
    data_row['tipo']          = str(int(row['TIPO']))

    retorno = ''

    ## Definição de variável para controle do fluxo.
    if verifica_tipo != data_row['tipo']:
        ## finaliza o processo anterior, aguardando mensagem de sucesso e pegando o número do documento
        if verifica_tipo != 0:
            retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior)

        uo_anterior = 0
        linha = 11
        orientacao_anterior = "Suplementar"
        conclusao = 0

    if data_row['tipo'] == '1':
        retorno, linha, conclusao = tipo_1(em, data_row, uo_anterior, orientacao_anterior, linha, conclusao)
    elif data_row['tipo'] == '2':
        retorno, linha, conclusao = tipo_2(em, data_row, uo_anterior, orientacao_anterior, linha, conclusao)
    elif data_row['tipo'] == '3':
        retorno, linha, conclusao = tipo_3(em, data_row, uo_anterior, orientacao_anterior, linha, conclusao)
    elif data_row['tipo'] == '4':
        retorno, linha, conclusao = tipo_4(em, data_row, uo_anterior, orientacao_anterior, linha, conclusao)

    uo_anterior = data_row['uo']  # armazena a UO da linha atual para comparação na próxima iteração
    orientacao_anterior = data_row['orientacao']  # armazena a orientação da linha atual para comparação na próxima iteração
    verifica_tipo = data_row['tipo']  # armazena a orientação da linha atual para comparação na próxima iteração

##    if data_row['tipo'] == '999':
##        break

### Conclui última linha processada, aguardando mensagem de sucesso e pegando o número do documento
if linha == 21:
    em.send_pf(8)  # envia F8 para ir para a próxima página
    em.wait_for_field()

retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior)

print()
print(f"Fluxo concluído.")

em.terminate()
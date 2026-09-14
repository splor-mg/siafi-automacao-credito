import os
import sys
import datetime


import shutil
import subprocess
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
import resultado
import analise_saldo
from arquivos import devolver_planilhas_de_origem
from relato import relato

agora = datetime.now()

hora_atual = datetime.now().strftime("%H:%M:%S")
print(f'Inicio do processo: {hora_atual}')

sistema = os.getenv('SISTEMA')
usuario = os.getenv('USUARIO')
senha = os.getenv('SENHA')
unidade_executora = os.getenv('UNIDADE_EXECUTORA')

day = datetime.today().strftime("%d")
month = datetime.today().strftime("%m")
year = datetime.today().strftime("%Y")

em = Emulator(visible=True)
em.connect('bhmvsb.prodemge.gov.br')
em.wait_for_field()


# Preenche os dados de login
em.fill_field(19, 13, sistema, 8)
em.fill_field(20, 13, usuario, 8)
em.fill_field(21, 13, senha, 8)
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
            relato('login', 'Login no SIAFI realizado')
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

em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()

em.fill_field(21, 19, '01', 2)
em.send_enter()
em.wait_for_field()

#Fazer loop para verificar "ok" e montar a solicitação

em.fill_field(21, 19, '02', 2)
em.fill_field(21, 41, '1', 1)
em.send_enter()
em.wait_for_field()

em.fill_field(9, 51, 'x', 1)
em.send_enter()
em.wait_for_field()

em.fill_field(9, 64, '01', 2)
em.fill_field(11, 64, '1261', 4) # Utilizar UO
em.send_enter()
em.wait_for_field()

#

x==11
y==20

#acrescentar 1 no x e 26 no y caso não encontrar a solicitação


breakpoint()

em.terminate()
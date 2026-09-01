import os
import sys
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
from relato import relato

load_dotenv()
ONEDRIVE_BASE = os.getenv('ONEDRIVE_BASE')

# Janela grafica do x3270 depende do WSLg da sessao interativa. Como servico do
# systemd nao ha DISPLAY, e o emulador morre com 'Can't open display'.
SIAFI_VISIVEL = os.getenv('SIAFI_VISIVEL', 'true').lower() == 'true'

# Antes o laco de conexao era infinito: com a VPN fora do ar o robo ficava
# preso para sempre, segurando o lock e sem devolver codigo de erro a ninguem.
CONEXAO_TENTATIVAS = int(os.getenv('CONEXAO_TENTATIVAS', '10'))

DIR_SCRIPTS = os.path.dirname(os.path.abspath(__file__))

#Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'ROBO'


# Como servico do systemd o PATH nao traz os diretorios do Windows, entao
# 'powershell.exe' nao e encontrado pelo nome — mesmo com o interop do WSL
# funcionando normalmente. Pelo caminho absoluto funciona.
_PS_PADRAO = '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
POWERSHELL = (_PS_PADRAO if os.path.exists(_PS_PADRAO)
              else (shutil.which('powershell.exe') or 'powershell.exe'))


def recalcular_no_excel(caminho):
    """Faz o proprio Excel abrir, recalcular e salvar a planilha.

    A aba ROBO nao guarda dados: ela e derivada por formula da aba
    'PREENCHER AQUI' — VLOOKUP contra a tabela qdd, concatenacao do codigo
    SIAFI completo, derivacao condicional de elemento e categoria. O openpyxl
    escreve os dados de origem mas NAO tem motor de formulas, entao os valores
    em cache da aba ROBO ficam vazios ate alguem abrir o arquivo no Excel.

    Era exatamente isso que a antiga confirmacao manual fazia: clicar em
    'Habilitar Edicao' (sair do Modo de Exibicao Protegido) e salvar. Aqui o
    Excel faz o servico sozinho, via COM — mesma engine, mesmos resultados.
    """
    win = subprocess.check_output(['wslpath', '-w', caminho]).decode().strip()
    ps = (
        "$ErrorActionPreference='Stop';"
        "try {"
        "  $xl = New-Object -ComObject Excel.Application;"
        "  $xl.Visible = $false; $xl.DisplayAlerts = $false;"
        "  $xl.AutomationSecurity = 3;"
        f"  $wb = $xl.Workbooks.Open('{win}');"
        "  $xl.CalculateFullRebuild(); $wb.Save(); $wb.Close($true); $xl.Quit();"
        "  Write-Output 'OK'"
        "} catch { Write-Output ('ERRO: ' + $_.Exception.Message) }"
    )
    try:
        saida = subprocess.run([POWERSHELL, '-NoProfile', '-Command', ps],
                               capture_output=True, text=True).stdout.strip()
    except OSError as e:
        # Sem isto o FileNotFoundError subia como traceback e o grupo recebia
        # 'FALHOU' com rodape vazio, sem nenhuma pista.
        print(f"[erro] nao consegui executar o PowerShell ({POWERSHELL}): {e}")
        return False

    if not saida.endswith('OK'):
        print(f"[erro] Excel nao recalculou: {saida}")
        return False
    return True


def linhas_prontas(caminho):
    """Quantas linhas a aba ROBO entrega de fato.

    E a mesma leitura que o robo faz depois. Zero aqui significa que o
    recalculo nao surtiu efeito: sem esta checagem o laco de processamento nao
    rodaria nenhuma vez e o robo estouraria adiante, em finalizar_documento(),
    com data_row indefinido.
    """
    df = pd.read_excel(caminho, sheet_name=SHEET_NAME)
    return int(df['UO_COD'].notna().sum())

# ---------------------------------------------------------------------------
# Etapa 1 — Consolidação das planilhas
# ---------------------------------------------------------------------------
try:
    subprocess.run(
        ['python3', os.path.join(DIR_SCRIPTS, 'consolida.py')],
        check=True
    )
except subprocess.CalledProcessError:
    print()
    print("=" * 70)
    print("  PROCESSO INTERROMPIDO NA CONSOLIDAÇÃO")
    print("=" * 70)
    print("  Corrija os erros apontados acima nas planilhas de origem")
    print("  e execute o robô novamente.")
    print("=" * 70)
    print()
    relato('erro', 'A consolidação não passou. O motivo está acima. '
                   'Use /log para o relatório completo.')
    sys.exit(1)

relato('planilha', 'Planilhas consolidadas, validação OK.')

CAMINHO_LOCAL = os.path.realpath(os.path.join(DIR_SCRIPTS, '..', 'data', 'copia.xlsm'))

# ---------------------------------------------------------------------------
# Etapa 2 — Recálculo das fórmulas
#
# A aba ROBO e derivada por formula da aba PREENCHER AQUI, e o openpyxl nao
# calcula nada. Sem este passo ela chega vazia no laco de solicitacoes.
#
# Nao ha mais parada para revisao: a conferencia da planilha e feita ANTES de
# ela ser posta na pasta de origem, igual ao robo de cota.
# ---------------------------------------------------------------------------
print()
print("Recalculando as fórmulas da planilha no Excel...")
if not recalcular_no_excel(CAMINHO_LOCAL):
    relato('erro', 'Não consegui fazer o Excel recalcular a planilha. A aba '
                   'ROBO é toda derivada por fórmula e ficaria vazia, então o '
                   'robô parou antes de tocar no SIAFI.')
    sys.exit(1)

_prontas = linhas_prontas(CAMINHO_LOCAL)
if _prontas == 0:
    relato('erro', 'A aba ROBO ficou vazia depois do recálculo. Confira se a '
                   'aba PREENCHER AQUI foi realmente preenchida. Nada foi '
                   'enviado ao SIAFI.')
    sys.exit(1)
print(f"{_prontas} linha(s) prontas na aba ROBO.")
# Sem o campo 'linhas': aqui o robo sabe a contagem, nao quais linhas
# sao — diferente do robo de cota, que lista os numeros.
relato('pendentes', f'{_prontas} linha(s) a processar')

# ---------------------------------------------------------------------------
# Etapa 3 — Login no SIAFI (uma sessão só, usada tambem pela analise)
# ---------------------------------------------------------------------------
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

em = None
for _tentativa in range(1, CONEXAO_TENTATIVAS + 1):
    em = Emulator(visible=SIAFI_VISIVEL)
    em.connect('bhmvsb.prodemge.gov.br')
    em.wait_for_field()

    if not em.string_found(1, 2, 'UNABLE TO ESTABLISH SESSION'):
        break

    print(f"Tentativa {_tentativa}/{CONEXAO_TENTATIVAS}: o servidor recusou a "
          "sessao (UNABLE TO ESTABLISH SESSION). Tentando novamente...")
    em.terminate()
    em = None
    time.sleep(1)
else:
    print()
    print(f"Nao foi possivel conectar ao SIAFI apos {CONEXAO_TENTATIVAS} tentativas.")
    relato('erro',
           f'Nao foi possivel conectar ao SIAFI apos {CONEXAO_TENTATIVAS} tentativas.\n'
           'Verifique se a VPN esta conectada e se nao ha sessao anterior aberta.')
    sys.exit(1)

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

# ---------------------------------------------------------------------------
# Etapa 4 — Análise de saldo, no MESMO emulador ja logado
#
# Antes a analise rodava como subprocesso e abria a sua propria sessao no
# SIAFI; o login.py abria outra logo depois, com o mesmo usuario. Enquanto o
# mainframe nao liberava a primeira, a segunda podia ser recusada com
# 'UNABLE TO ESTABLISH SESSION'. Com um emulador so, essa corrida some.
# ---------------------------------------------------------------------------
print()
print("Iniciando a análise de saldo das dotações...")
print()

if analise_saldo.analisar(em) != 0:
    print()
    print("=" * 70)
    print("  PROCESSO INTERROMPIDO NA ANÁLISE DE SALDO")
    print("=" * 70)
    print("  Nenhuma solicitação foi enviada ao SIAFI.")
    print("=" * 70)
    print()
    relato('aviso', 'Análise de saldo reprovada: NENHUMA solicitação foi enviada '
                    'ao SIAFI. Corrija as dotações apontadas e acione de novo. '
                    'Use /log para ver quais.')
    em.terminate()
    sys.exit(2)

# A analise deixou o terminal nas telas de consulta. Voltar ao ponto de partida
# do laco de solicitacoes e obrigatorio: digitar numa tela errada do SIAFI e o
# pior desfecho possivel, entao aqui e abortar, nao seguir na duvida.
if not analise_saldo.voltar_ao_menu(em):
    relato('erro', 'Não consegui voltar ao menu do SIAFI depois da análise de '
                   'saldo. O robô parou por segurança, antes de enviar '
                   'qualquer solicitação.')
    em.terminate()
    sys.exit(1)

print("Análise de saldo aprovada. Iniciando as solicitações no SIAFI...")
print()

# ---------------------------------------------------------------------------
# Etapa 5 — Solicitações
# ---------------------------------------------------------------------------
df = pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME)
df = df.dropna(how='all')
df = df.reset_index(drop=False)

# Definição de variáveis para controle do fluxo
verifica_tipo = 0
conclusao = 0
linha = 11

try:
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
        data_row['funcao']        = str(int(row['FUNCAO_COD'])).zfill(2) 
        data_row['subfuncao']     = str(int(row['SUBFUNCAO_COD'])).zfill(3) 
        data_row['programa']      = str(int(row['PROGRAMA_COD'])).zfill(3) 
        data_row['subprojeto']    = str(int(row['SUBPROJETO_COD']))
        data_row['categoria']     = str(int(row['CATEGORIA_COD']))
        data_row['grupo']         = str(int(row['GRUPO_COD']))
        data_row['modalidade']    = str(int(row['MODALIDADE_COD']))
        data_row['elemento']      = str(int(row['ELEMENTO_COD'])).zfill(2) 
        data_row['iag']           = str(int(row['IPG_COD']))
        data_row['fonte']         = str(int(row['FONTE_COD'])).zfill(2) 
        data_row['procedencia']   = str(int(row['IPU_COD']))

        if data_row['orientacao'] == 'Anular': # se for anulação, o valor deve ser multiplicado por -1 para ficar negativo
            data_row['valor']      = str(-int(round(row['VALOR'])))
        else:
            data_row['valor']      = str(int(round(row['VALOR'])))

        data_row['uo_suplementada'] = str(int(row['UO_SUPLEMENTADA'])) if pd.notna(row['UO_SUPLEMENTADA']) else '0'
        data_row['tipo']          = str(int(row['TIPO']))

        retorno = ''

        ## Definição de variável para controle do fluxo.
        if verifica_tipo != data_row['tipo']:
            ## finaliza o processo anterior, aguardando mensagem de sucesso e pegando o número do documento
            if verifica_tipo != 0:
                retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior, data_row)

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
        resultado.marcar_linha(int(row['index']) + 2)  # última linha que entrou no documento em aberto

    ##    if data_row['tipo'] == '999':
    ##        break

    ### Conclui última linha processada, aguardando mensagem de sucesso e pegando o número do documento
    if linha == 21:
        em.send_pf(8)  # envia F8 para ir para a próxima página
        em.wait_for_field()

    retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior, data_row)


    print()
    relato('fim', 'Fluxo concluído.')

finally:
    hora_atual = datetime.now().strftime("%H:%M:%S")
    print(f'Fim do processo: {hora_atual}')

    em.terminate()

    # Grava o desfecho de cada documento na aba ROBO (colunas U e V)
    resultado.gravar(CAMINHO_LOCAL)
    resultado.imprimir_resumo()

    # Salva cópia de conferência do arquivo lido
    conferencia_folder = os.path.join(ONEDRIVE_BASE, 'Conferencia arquivo robo')
    realizados_automacao = os.path.join(ONEDRIVE_BASE, 'Realizados', 'Automação Python')

    os.makedirs(conferencia_folder, exist_ok=True)
    os.makedirs(realizados_automacao, exist_ok=True)

    for arquivo_existente in os.listdir(conferencia_folder):
        origem = os.path.join(conferencia_folder, arquivo_existente)
        if os.path.isfile(origem):
            destino = os.path.join(realizados_automacao, arquivo_existente)
            if os.path.exists(destino):
                nome, ext = os.path.splitext(arquivo_existente)
                contador = 1
                while os.path.exists(destino):
                    destino = os.path.join(realizados_automacao, f"{nome} ({contador}){ext}")
                    contador += 1
            shutil.move(origem, destino)
            print(f"Arquivo anterior movido para: {destino}")

    hoje = datetime.today().strftime("%d.%m")
    novo_nome = f'Conferencia arquivo robo {hoje}.xlsm'
    destino_copia = os.path.join(conferencia_folder, novo_nome)
    shutil.copyfile(CAMINHO_LOCAL, destino_copia)
    relato('planilha_final',
           f'Cópia de conferência salva em: {destino_copia}', arquivo=novo_nome)
    print(f"Cópia de conferência salva em: {destino_copia}")
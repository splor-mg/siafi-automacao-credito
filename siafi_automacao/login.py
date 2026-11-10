import os
import sys
import shutil
import subprocess
from dotenv import load_dotenv
from py3270 import Emulator
from datetime import datetime
import pandas as pd
import time
from fluxo_tipo_1 import tipo_1
from fluxo_tipo_2 import tipo_2
from fluxo_tipo_3 import tipo_3
from fluxo_tipo_4 import tipo_4
from utils_siafi import finalizar_documento, TelaDesconhecida
import resultado
import analise_saldo
from arquivos import devolver_planilhas_de_origem
from excel import recalcular_no_excel
from linhas import ler_linhas
from relato import relato
from tela import imprimir_tela

load_dotenv()
ONEDRIVE_BASE = os.getenv('ONEDRIVE_BASE')

# Janela grafica do x3270 depende do WSLg da sessao interativa. Como servico do
# systemd nao ha DISPLAY, e o emulador morre com 'Can't open display'.
SIAFI_VISIVEL = os.getenv('SIAFI_VISIVEL', 'true').lower() == 'true'

# Antes o laco de conexao era infinito: com a VPN fora do ar o robo ficava
# preso para sempre, segurando o lock e sem devolver codigo de erro a ninguem.
CONEXAO_TENTATIVAS = int(os.getenv('CONEXAO_TENTATIVAS', '10'))

DIR_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.realpath(os.path.join(DIR_SCRIPTS, '..', 'data'))
CAMINHO_LOCAL = os.path.join(DATA_DIR, 'copia.xlsm')
# Manifesto gravado pelo consolida.py: de onde saiu cada planilha de origem.
MOVIDOS_PATH = os.path.join(DATA_DIR, '.movidos')

SIAFI_HOST = 'bhmvsb.prodemge.gov.br'

#Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'ROBO'


def senha_expirada(em):
    """Procura o aviso de senha expirada em qualquer linha da tela.

    Varre a tela inteira em vez de apostar numa linha fixa: errar a linha faria
    o robo seguir adiante e morrer depois em 'Nao foi possivel fazer login',
    escondendo o motivo real. Em caso de falha de leitura devolve False, para
    nao trocar um diagnostico ruim por uma excecao no meio do login.
    """
    try:
        tela = '\n'.join(em.string_get(l, 1, 80) for l in range(1, 25))
    except Exception:
        return False
    return 'SENHA EXPIRADA' in tela.upper()


def encerrar_emulador(em):
    """Encerra o emulador ignorando falhas.

    Se a sessao ja caiu o terminate() estoura, e isso nao pode impedir o que
    vem depois dele — devolver as planilhas, gravar o resultado na aba ROBO.
    """
    if em is None:
        return
    try:
        em.terminate()
    except Exception as e:
        print(f"[aviso] falha ao encerrar o emulador (ignorando): {e}")


def encerrar_sem_envio(codigo, motivo, em=None):
    """Sai do robo ANTES de qualquer envio ao SIAFI, devolvendo as planilhas.

    O consolida.py ja tirou as planilhas da pasta de origem. Como nada foi
    enviado, elas precisam voltar para a fila: sem isso a pessoa resolve o
    problema (VPN, senha, dotacao, coluna vazia), aciona de novo e recebe
    'NENHUMA PLANILHA ENCONTRADA PARA PROCESSAR'. Antes so a reprovacao na
    analise de saldo devolvia; as outras saidas deixavam tudo em Realizados.
    """
    if devolver_planilhas_de_origem(MOVIDOS_PATH):
        aviso = (f'{motivo}: NENHUMA solicitação foi enviada ao SIAFI. As '
                 'planilhas voltaram para a pasta de origem — resolva o '
                 'problema apontado e acione de novo.')
    else:
        aviso = (f'{motivo}: NENHUMA solicitação foi enviada ao SIAFI. Não '
                 'consegui devolver as planilhas para a pasta de origem; veja '
                 'em Realizados antes de acionar de novo.')
    relato('aviso', aviso)
    encerrar_emulador(em)
    sys.exit(codigo)


def abortar_senha_expirada(em):
    """Encerra o robo explicando o que fazer. Codigo 3, tratado no robo.ps1."""
    print()
    print("=" * 70)
    print("Senha expirada. Abra o SIAFI manualmente e atualize sua senha.")
    print("Apos isso, salve a nova senha no arquivo .env e execute o script novamente.")
    print("=" * 70)
    relato('erro', 'Senha do SIAFI expirada. Atualize a senha no SIAFI '
                   'manualmente e grave a nova no .env.')
    encerrar_sem_envio(3, 'Senha expirada', em)


def abortar_login(em):
    """O SIAFI nao chegou a tela esperada depois de varias tentativas.

    Antes o robo so imprimia o aviso, encerrava o emulador e SEGUIA: a linha
    seguinte digitava num emulador morto e o grupo recebia um traceback.
    """
    print("Não foi possível fazer login após várias tentativas.")
    relato('erro', 'Não foi possível fazer login no SIAFI após várias '
                   'tentativas. Confira usuário e senha no .env e se o SIAFI '
                   'está no ar.')
    encerrar_sem_envio(1, 'Login não concluído', em)


def relatar_parada(em, erro, linha_planilha, uo):
    """Explica ao grupo ONDE e POR QUE o envio parou no meio.

    Das 83 execucoes ate 10/11/2026, 14 morreram com 'Keyboard locked' e nada
    mais: o robo tentou digitar numa posicao protegida porque o SIAFI nao foi
    para a tela esperada, e a mensagem do SIAFI — que diz o motivo — nunca
    era lida. Aqui a tela inteira vai para o log e a linha 1 para o grupo.
    """
    tela = imprimir_tela(em, f'envio interrompido na linha {linha_planilha} '
                             f'da planilha, UO {uo}')
    print(f"  {type(erro).__name__}: {erro}")
    mensagem = tela[0].strip() if tela else ''

    if 'Keyboard locked' in str(erro):
        causa = ('o SIAFI não foi para a tela esperada (em geral porque '
                 'recusou os dados digitados)')
    else:
        causa = f'{type(erro).__name__}: {erro}'

    texto = (f'O envio parou ao processar a linha {linha_planilha} da '
             f'planilha (UO {uo}): {causa}.')
    if mensagem:
        texto += f' O SIAFI mostrava: "{mensagem}".'

    sem_resposta = resultado.confirmacao_sem_resposta()
    if sem_resposta:
        texto += (f' A confirmação do documento da UO {sem_resposta} já tinha '
                  'sido enviada e o robô não leu a resposta: confira no SIAFI '
                  'se ele foi registrado.')

    if tela:
        texto += ' A tela completa do SIAFI está no log (/log).'
    relato('erro', texto)


def planilhas_apos_parada():
    """O envio parou no meio: as planilhas de origem voltam para a fila?

    So quando e CERTO que nada foi registrado, isto e, quando nenhuma
    confirmacao (o PF5 do finalizar_documento) chegou a ser enviada nesta
    execucao. Linha digitada e nao confirmada nao vira documento no SIAFI.

    Com alguma confirmacao ja enviada — documento registrado, recusado ou sem
    resposta lida — elas ficam em Realizados: devolver faria o proximo
    acionamento reenviar tudo, inclusive o que ja esta no SIAFI.
    """
    if resultado.houve_confirmacao():
        relato('aviso', 'As planilhas de origem ficaram em Realizados, porque '
                        'já houve documento confirmado ou recusado nesta '
                        'execução. Confira no SIAFI e reenvie só o que faltou.')
    elif devolver_planilhas_de_origem(MOVIDOS_PATH):
        relato('aviso', 'Nenhum documento chegou a ser confirmado no SIAFI, '
                        'então as planilhas voltaram para a pasta de origem — '
                        'resolva o problema apontado e acione de novo.')
    else:
        relato('aviso', 'Nenhum documento chegou a ser confirmado no SIAFI, '
                        'mas não consegui devolver as planilhas para a pasta '
                        'de origem; veja em Realizados antes de acionar de '
                        'novo.')


# ---------------------------------------------------------------------------
# Etapa 1 — Consolidação das planilhas
# ---------------------------------------------------------------------------
try:
    subprocess.run(
        [sys.executable, os.path.join(DIR_SCRIPTS, 'consolida.py')],
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

relato('planilha', 'Planilhas consolidadas.')

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
_erro_excel = recalcular_no_excel(CAMINHO_LOCAL)
if _erro_excel:
    # A aba ROBO e toda derivada por formula: sem o recalculo ela fica vazia.
    relato('erro', 'Não consegui fazer o Excel recalcular a planilha, então o '
                   'robô parou antes de tocar no SIAFI. Motivo: '
                   f'{_erro_excel[:400].rstrip(".")}.')
    encerrar_sem_envio(1, 'Planilha não recalculada')

hoje = datetime.today()
LINHAS, _problemas = ler_linhas(
    pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME), hoje)
_prontas = len(LINHAS)
if _prontas == 0:
    relato('erro', 'A aba ROBO ficou vazia depois do recálculo. Confira se a '
                   'aba PREENCHER AQUI foi realmente preenchida — e, na '
                   'máquina do robô, se não há um Excel aberto segurando o '
                   'copia.xlsm.')
    encerrar_sem_envio(1, 'Planilha sem linhas')

# Validacao de TODAS as linhas antes do login: linha com defeito descoberta no
# meio do laco deixava documento aberto pela metade no SIAFI.
if _problemas:
    print()
    print("=" * 70)
    print("  LINHAS COM PROBLEMA NA PLANILHA — PROCESSO NÃO PODE CONTINUAR")
    print("=" * 70)
    for p in _problemas:
        print(f"  - {p}")
    print("=" * 70)
    print()
    _LIMITE = 8
    _detalhes = [f'- {p}' for p in _problemas[:_LIMITE]]
    if len(_problemas) > _LIMITE:
        _detalhes.append(f'- ... e mais {len(_problemas) - _LIMITE}, veja /log')
    relato('erro', 'Linhas com problema na planilha:\n' + '\n'.join(_detalhes))
    encerrar_sem_envio(2, 'Planilha reprovada na validação')

print(f"{_prontas} linha(s) prontas na aba ROBO.")
# Sem o campo 'linhas': aqui o robo sabe a contagem, nao quais linhas
# sao — diferente do robo de cota, que lista os numeros.
relato('pendentes', f'{_prontas} linha(s) a processar')

# ---------------------------------------------------------------------------
# Etapa 3 — Login no SIAFI (uma sessão só, usada tambem pela analise)
# ---------------------------------------------------------------------------
hora_atual = datetime.now().strftime("%H:%M:%S")
print(f'Inicio do processo: {hora_atual}')

sistema = os.getenv('SISTEMA')
usuario = os.getenv('USUARIO')
senha = os.getenv('SENHA')
unidade_executora = os.getenv('UNIDADE_EXECUTORA')

em = None
_ultimo_motivo = None
for _tentativa in range(1, CONEXAO_TENTATIVAS + 1):
    # O try cobre o Emulator()/connect() e nao so a tela de recusa: com a VPN
    # fora do ar a excecao estoura antes de haver tela para ler, e subia como
    # traceback cru, sem motivo nenhum para o grupo.
    try:
        em = Emulator(visible=SIAFI_VISIVEL)
        em.connect(SIAFI_HOST)
        em.wait_for_field()

        if not em.string_found(1, 2, 'UNABLE TO ESTABLISH SESSION'):
            break
        _ultimo_motivo = 'o servidor recusou a sessao (UNABLE TO ESTABLISH SESSION)'
    except Exception as e:
        _ultimo_motivo = f'{type(e).__name__}: {e}'

    print(f"Tentativa {_tentativa}/{CONEXAO_TENTATIVAS} de conectar falhou: "
          f"{_ultimo_motivo}")
    encerrar_emulador(em)
    em = None
    time.sleep(1)
else:
    print()
    relato('erro',
           f'Nao foi possivel conectar ao SIAFI apos {CONEXAO_TENTATIVAS} tentativas.\n'
           f'Ultimo motivo: {_ultimo_motivo}\n'
           'Verifique se a VPN esta conectada e se nao ha sessao anterior aberta.')
    encerrar_sem_envio(1, 'Sem conexão com o SIAFI')

# Preenche os dados de login
em.fill_field(19, 13, sistema, 8)
em.fill_field(20, 13, usuario, 8)
em.fill_field(21, 13, senha, 8)
em.send_enter()

time.sleep(1)
if senha_expirada(em):
    abortar_senha_expirada(em)

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

    # Fora do try de proposito: o 'except:' acima captura BaseException e
    # engoliria o SystemExit de abortar_senha_expirada(). O aviso pode aparecer
    # so depois de um Enter, numa tela intermediaria.
    if senha_expirada(em):
        abortar_senha_expirada(em)

if tentativas == max_tentativas:
    abortar_login(em)

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
    abortar_login(em)

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

try:
    _saldo_reprovado = analise_saldo.analisar(em) != 0
except Exception as e:
    # A analise so consulta: se ela quebrar, nada foi enviado ainda.
    print(f"[erro] a análise de saldo quebrou: {type(e).__name__}: {e}")
    imprimir_tela(em, 'análise de saldo interrompida')
    relato('erro', f'A análise de saldo não terminou '
                   f'({type(e).__name__}: {e}).')
    encerrar_sem_envio(1, 'Análise de saldo interrompida', em)

if _saldo_reprovado:
    print()
    print("=" * 70)
    print("  PROCESSO INTERROMPIDO NA ANÁLISE DE SALDO")
    print("=" * 70)
    print("  Nenhuma solicitação foi enviada ao SIAFI.")
    print("=" * 70)
    print()
    encerrar_sem_envio(2, 'Análise de saldo reprovada', em)

# A analise deixou o terminal nas telas de consulta. Voltar ao ponto de partida
# do laco de solicitacoes e obrigatorio: digitar numa tela errada do SIAFI e o
# pior desfecho possivel, entao aqui e abortar, nao seguir na duvida.
if not analise_saldo.voltar_ao_menu(em):
    relato('erro', 'Não consegui voltar ao menu do SIAFI depois da análise de '
                   'saldo. O robô parou por segurança, antes de enviar '
                   'qualquer solicitação.')
    encerrar_sem_envio(1, 'Menu do SIAFI não encontrado', em)

print("Análise de saldo aprovada. Iniciando as solicitações no SIAFI...")
print()

# ---------------------------------------------------------------------------
# Etapa 5 — Solicitações
# ---------------------------------------------------------------------------
# Definição de variáveis para controle do fluxo
verifica_tipo = 0
conclusao = 0
linha = 11
linha_planilha = None   # linha da planilha em processamento, para o diagnostico

try:
    # As linhas ja chegam convertidas e validadas (ver linhas.py)
    for linha_planilha, data_row in LINHAS:

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
        resultado.marcar_linha(linha_planilha)  # última linha que entrou no documento em aberto

    ### Conclui última linha processada, aguardando mensagem de sucesso e pegando o número do documento
    if linha == 21:
        em.send_pf(8)  # envia F8 para ir para a próxima página
        em.wait_for_field()

    retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior, data_row)


    print()
    relato('fim', 'Fluxo concluído.')

except Exception as e:
    # TelaDesconhecida: o finalizar_documento ja explicou o motivo ao grupo
    # (e ja tirou o print da tela) antes de levantar.
    if not isinstance(e, TelaDesconhecida):
        relatar_parada(em, e, linha_planilha, data_row['uo'])
    planilhas_apos_parada()
    raise

finally:
    hora_atual = datetime.now().strftime("%H:%M:%S")
    print(f'Fim do processo: {hora_atual}')

    encerrar_emulador(em)

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

    novo_nome = f'Conferencia arquivo robo {datetime.today().strftime("%d.%m")}.xlsm'
    destino_copia = os.path.join(conferencia_folder, novo_nome)
    shutil.copyfile(CAMINHO_LOCAL, destino_copia)
    relato('planilha_final',
           f'Cópia de conferência salva em: {destino_copia}', arquivo=novo_nome)

"""
Análise de saldo das dotações do arquivo consolidado (copia.xlsm).

Executado pelo login.py logo após a confirmação do usuário no Excel.

Regra de saída:
    exit 0  -> todas as dotações a anular têm saldo suficiente
    exit 1  -> pelo menos uma dotação está inexistente, com saldo insuficiente
               ou não pôde ser lida

O login.py trata o exit 1 abortando o processo antes de enviar qualquer
solicitação ao SIAFI.
"""

import os
import sys
import time
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from py3270 import Emulator

load_dotenv()

SISTEMA = os.getenv('SISTEMA')
USUARIO = os.getenv('USUARIO')
SENHA = os.getenv('SENHA')
UNIDADE_EXECUTORA = os.getenv('UNIDADE_EXECUTORA')

# Caminho da planilha consolidada
CAMINHO_LOCAL = os.path.realpath(
    os.path.join(os.path.dirname(__file__), '..', 'data', 'copia.xlsm')
)

# Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'ROBO'

MSG_INEXISTENTE = 'INEXISTENTE'


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------

def conectar_e_logar():
    """Abre o emulador, faz login no SIAFI e informa a Unidade Executora."""
    em = Emulator(visible=True)  # use visible=False para rodar sem janela gráfica
    em.connect('bhmvsb.prodemge.gov.br')
    em.wait_for_field()

    # Preenche os dados de login
    em.fill_field(19, 13, SISTEMA, 7)
    em.fill_field(20, 13, USUARIO, 7)
    em.fill_field(21, 13, SENHA, 7)
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

        except Exception:
            print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
            em.send_enter()

        tentativas += 1

    if tentativas == max_tentativas:
        em.terminate()
        raise RuntimeError("Não foi possível fazer login após várias tentativas.")

    em.fill_field(1, 2, SISTEMA, 4)
    em.send_enter()

    # Nova tela buscando login...
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

        except Exception:
            print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
            em.send_enter()

        tentativas += 1

    if tentativas == max_tentativas:
        em.terminate()
        raise RuntimeError("Não foi possível fazer login após várias tentativas.")

    # Entrar com a Unidade Executora
    em.fill_field(22, 30, UNIDADE_EXECUTORA, 7)
    em.send_enter()
    em.wait_for_field()

    return em


# ---------------------------------------------------------------------------
# Navegação
# ---------------------------------------------------------------------------

def entrar_na_consulta(em):
    """Consultas -> Movimentação orçamentária -> Consulta de dotação."""
    em.fill_field(21, 19, '09', 2)
    em.send_enter()
    em.wait_for_field()

    em.fill_field(21, 19, '03', 2)
    em.send_enter()
    em.wait_for_field()

    abrir_tela_dotacao(em)


def abrir_tela_dotacao(em):
    """Abre (ou reabre) a tela de consulta de dotação."""
    em.fill_field(21, 19, '07', 2)
    em.fill_field(21, 41, '1', 1)
    em.send_enter()
    em.wait_for_field()


def voltar_para_dotacao(em):
    """Sai da consulta atual (PF3) e reabre a tela de consulta de dotação."""
    em.send_pf(3)
    em.wait_for_field()
    time.sleep(0.5)
    abrir_tela_dotacao(em)


# ---------------------------------------------------------------------------
# Consulta de uma dotação
# ---------------------------------------------------------------------------

def consultar_dotacao(em, data_row):
    """
    Consulta a dotação no SIAFI e devolve (status, saldo, retorno_tela).

    status: 'OK', 'INEXISTENTE', 'SALDO INSUFICIENTE' ou 'ERRO DE LEITURA'
    """
    em.fill_field(9, 75, data_row['uo'], 4)
    em.send_enter()
    em.wait_for_field()

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

    if MSG_INEXISTENTE in retorno.upper():
        voltar_para_dotacao(em)
        return 'INEXISTENTE', None, retorno

    # Número alinhado à direita terminando na coluna 78.
    # Lê 29 caracteres a partir da coluna 50 e pega o último token,
    # cobrindo qualquer tamanho de valor (de "1,00" a "1.000.000.000.000,00").
    saldo_raw = em.string_get(13, 50, 29)

    try:
        saldo_str = saldo_raw.split()[-1]
        saldo = float(saldo_str.replace('.', '').replace(',', '.'))
    except (ValueError, IndexError):
        voltar_para_dotacao(em)
        return 'ERRO DE LEITURA', None, f"saldo lido: {saldo_raw!r}"

    voltar_para_dotacao(em)

    if saldo >= data_row['valor']:
        return 'OK', saldo, retorno
    else:
        return 'SALDO INSUFICIENTE', saldo, retorno


# ---------------------------------------------------------------------------
# Montagem da linha
# ---------------------------------------------------------------------------

def montar_data_row(row):
    """Extrai e normaliza os campos necessários para a consulta de dotação."""
    data_row = {}
    data_row['orientacao']  = str(row['ORIENTACAO']).strip()
    data_row['uo']          = str(int(row['UO_COD'])).zfill(4)
    data_row['acao']        = str(int(row['ACAO_COD'])).zfill(4)
    data_row['funcao']      = str(int(row['FUNCAO_COD'])).zfill(2)
    data_row['subfuncao']   = str(int(row['SUBFUNCAO_COD'])).zfill(3)
    data_row['programa']    = str(int(row['PROGRAMA_COD'])).zfill(3)
    data_row['subprojeto']  = str(int(row['SUBPROJETO_COD']))
    data_row['categoria']   = str(int(row['CATEGORIA_COD']))
    data_row['grupo']       = str(int(row['GRUPO_COD']))
    data_row['modalidade']  = str(int(row['MODALIDADE_COD'])).zfill(2)
    data_row['elemento']    = str(int(row['ELEMENTO_COD'])).zfill(2)
    data_row['iag']         = str(int(row['IPG_COD']))
    data_row['fonte']       = str(int(row['FONTE_COD'])).zfill(2)
    data_row['procedencia'] = str(int(row['IPU_COD']))

    # VALOR vem em centavos na planilha; a tela do SIAFI mostra o saldo em reais.
    # Usa-se o valor absoluto porque a comparação é sempre "há saldo para anular?".
    data_row['valor'] = abs(float(row['VALOR'])) / 100

    return data_row


def descrever(data_row, linha_excel):
    return (
        f"linha {linha_excel} | UO: {data_row['uo']}; Ação: {data_row['acao']}; "
        f"Grupo: {data_row['grupo']}; Modalidade: {data_row['modalidade']}; "
        f"Fonte: {data_row['fonte']}; Procedência: {data_row['procedencia']}; "
        f"Valor: {data_row['valor']:.2f}"
    )


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

def main():
    print()
    print("=" * 70)
    print("  ANÁLISE DE SALDO DAS DOTAÇÕES")
    print("=" * 70)
    print(f"Início: {datetime.now().strftime('%H:%M:%S')}")
    print()

    df = pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME)
    df = df.dropna(how='all')          # remove linhas completamente vazias
    df = df.reset_index(drop=False)    # mantém o índice original para o nº da linha no Excel

    em = conectar_e_logar()

    resultados = []

    try:
        entrar_na_consulta(em)

        for _, row in df.iterrows():

            # Pula linhas onde UO_COD está vazio (fim da planilha ou linha vazia)
            if pd.isna(row['UO_COD']):
                continue

            linha_excel = int(row['index']) + 2  # +2: cabeçalho na linha 1

            data_row = montar_data_row(row)

            # Somente linhas de anulação consomem saldo; suplementações não.
            if data_row['orientacao'].lower() != 'anular':
                print(f"{descrever(data_row, linha_excel)} -> Suplementar (sem verificação)")
                continue

            status, saldo, retorno = consultar_dotacao(em, data_row)

            saldo_txt = f"{saldo:.2f}" if saldo is not None else "-"
            print(f"{descrever(data_row, linha_excel)}; Saldo: {saldo_txt} -> {status}")

            resultados.append({
                'linha': linha_excel,
                'descricao': descrever(data_row, linha_excel),
                'status': status,
                'saldo': saldo,
                'retorno': retorno,
            })

    finally:
        em.terminate()

    # -----------------------------------------------------------------------
    # Relatório final
    # -----------------------------------------------------------------------
    problemas = [r for r in resultados if r['status'] != 'OK']

    print()
    print("-" * 70)
    print(f"Dotações verificadas: {len(resultados)} | "
          f"OK: {len(resultados) - len(problemas)} | Com problema: {len(problemas)}")
    print(f"Fim: {datetime.now().strftime('%H:%M:%S')}")
    print("-" * 70)

    if problemas:
        print()
        print("=" * 70)
        print("  DOTAÇÕES COM PROBLEMA — PROCESSO NÃO PODE CONTINUAR")
        print("=" * 70)
        for r in problemas:
            print(f"  [{r['status']}] {r['descricao']}")
            if r['saldo'] is not None:
                print(f"      saldo disponível: {r['saldo']:.2f}")
            if r['retorno']:
                print(f"      mensagem SIAFI: {r['retorno']}")
        print("=" * 70)
        print()
        return 1

    print()
    print("Todas as dotações a anular possuem saldo suficiente.")
    print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
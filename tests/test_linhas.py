"""A aba ROBO e conferida inteira ANTES de o robo tocar no SIAFI.

O caso que motivou este arquivo e a execucao de 10/11/2026 as 10:13: tres
linhas chegaram com a coluna TIPO vazia (a formula da aba ROBO devolve '999'),
foram puladas em silencio pelo laco de solicitacoes e o robo quebrou em
seguida com 'Keyboard locked', tentando finalizar um documento que nunca
abriu. O grupo recebeu 'FALHOU (codigo 1)' sem nenhum motivo.
"""
from datetime import datetime

import pandas as pd

from linhas import ler_linhas, montar_data_row

HOJE = datetime(2026, 11, 10)


def linha_robo(**trocas):
    """Uma linha da aba ROBO como o pandas a entrega depois do recalculo."""
    base = {
        'ORIENTACAO': 'Suplementar', 'UO_COD': 1521, 'ACAO_COD': 2500,
        'FUNCAO_COD': 4, 'SUBFUNCAO_COD': 122, 'PROGRAMA_COD': 705,
        'SUBPROJETO_COD': 1, 'CATEGORIA_COD': 3, 'GRUPO_COD': 3,
        'MODALIDADE_COD': 90, 'ELEMENTO_COD': 40, 'IPG_COD': '0',
        'FONTE_COD': 10, 'IPU_COD': 1, 'VALOR': 25041.0,
        'UO_SUPLEMENTADA': 4031, 'TIPO': '1',
    }
    base.update(trocas)
    return base


def test_linha_valida_sai_no_formato_que_os_fluxos_digitam():
    data_row, problemas = montar_data_row(linha_robo(), HOJE)

    assert problemas == []
    assert data_row == {
        'day': '10', 'month': '11', 'year': '2026',
        'orientacao': 'Suplementar', 'uo': '1521', 'acao': '2500',
        'funcao': '04', 'subfuncao': '122', 'programa': '705',
        'subprojeto': '1', 'categoria': '3', 'grupo': '3',
        'modalidade': '90', 'elemento': '40', 'iag': '0', 'fonte': '10',
        'procedencia': '1', 'valor': '25041', 'uo_suplementada': '4031',
        'tipo': '1',
    }


def test_anulacao_vai_com_valor_positivo():
    """Na aba ROBO a anulacao vem negativa; o SIAFI recebe o valor positivo."""
    data_row, problemas = montar_data_row(
        linha_robo(ORIENTACAO='Anular', VALOR=-25041.0), HOJE)

    assert problemas == []
    assert data_row['valor'] == '25041'


def test_tipo_nao_preenchido_e_apontado():
    _data_row, problemas = montar_data_row(linha_robo(TIPO='999'), HOJE)

    assert problemas == ['coluna TIPO não preenchida na planilha de origem']


def test_tipo_que_o_robo_nao_trata_e_apontado():
    _data_row, problemas = montar_data_row(linha_robo(TIPO='7'), HOJE)

    assert len(problemas) == 1 and 'TIPO 7 desconhecido' in problemas[0]


def test_anulacao_nos_tipos_que_so_suplementam_e_apontada():
    """Saldo financeiro (2) e excesso de arrecadacao (3) nao tem dotacao a
    anular. Antes a linha passava pela analise de saldo e era pulada em
    silencio pelo fluxo, sem ninguem saber que nao tinha sido enviada."""
    _d, do_tipo_2 = montar_data_row(
        linha_robo(TIPO='2', ORIENTACAO='Anular', VALOR=-25041.0), HOJE)
    _d, do_tipo_3 = montar_data_row(
        linha_robo(TIPO='3', ORIENTACAO='Anular', VALOR=-25041.0), HOJE)

    assert do_tipo_2 == ['TIPO 2 (Saldo Financeiro) só aceita suplementação, '
                         'e a linha é de anulação']
    assert do_tipo_3 == ['TIPO 3 (Excesso de arrecadação) só aceita '
                         'suplementação, e a linha é de anulação']


def test_suplementacao_nos_tipos_2_e_3_passa():
    for tipo in ('2', '3'):
        _d, problemas = montar_data_row(linha_robo(TIPO=tipo), HOJE)
        assert problemas == []


def test_anulacao_nos_tipos_1_e_4_continua_valendo():
    for tipo in ('1', '4'):
        _d, problemas = montar_data_row(
            linha_robo(TIPO=tipo, ORIENTACAO='Anular', VALOR=-25041.0), HOJE)
        assert problemas == []


def test_celula_vazia_e_apontada_pelo_nome_da_coluna():
    """ELEMENTO_COD vazio e o caso real (modalidades de transferencia): antes
    estourava 'cannot convert float NaN to integer' no meio do envio."""
    _data_row, problemas = montar_data_row(
        linha_robo(ELEMENTO_COD=float('nan'), FUNCAO_COD=''), HOJE)

    assert problemas == ['FUNCAO_COD vazio', 'ELEMENTO_COD vazio']


def test_orientacao_desconhecida_e_apontada():
    """Os fluxos so digitam 'Suplementar' e 'Anular': qualquer outra coisa
    passava sem ser enviada e sem aviso."""
    _data_row, problemas = montar_data_row(
        linha_robo(ORIENTACAO='Suplementação'), HOJE)

    assert len(problemas) == 1 and 'ORIENTACAO' in problemas[0]


def test_uo_suplementada_vazia_vira_zero():
    data_row, problemas = montar_data_row(
        linha_robo(UO_SUPLEMENTADA=float('nan')), HOJE)

    assert problemas == []
    assert data_row['uo_suplementada'] == '0'


def test_planilha_de_10_11_e_reprovada_antes_do_siafi():
    """As cinco linhas daquela execucao: duas boas e tres sem TIPO."""
    df = pd.DataFrame([
        linha_robo(),
        linha_robo(ORIENTACAO='Anular', ACAO_COD=4304, VALOR=-25041.0),
        linha_robo(ORIENTACAO='Anular', UO_COD=1991, ACAO_COD=9999,
                   VALOR=-49100000.0, TIPO='999'),
        linha_robo(UO_COD=2171, ACAO_COD=2417, VALOR=3159412.0, TIPO='999'),
        linha_robo(UO_COD=2171, ACAO_COD=2500, VALOR=45940588.0, TIPO='999'),
    ])

    linhas, problemas = ler_linhas(df, HOJE)

    assert [n for n, _ in linhas] == [2, 3, 4, 5, 6]
    assert problemas == [
        'linha 4 (UO 1991): coluna TIPO não preenchida na planilha de origem',
        'linha 5 (UO 2171): coluna TIPO não preenchida na planilha de origem',
        'linha 6 (UO 2171): coluna TIPO não preenchida na planilha de origem',
    ]


def test_linhas_sem_uo_sao_o_fim_dos_dados_e_nao_um_problema():
    """A aba ROBO tem formulas ate a linha 1000; as vazias nao contam."""
    df = pd.DataFrame([
        linha_robo(),
        linha_robo(UO_COD=float('nan'), TIPO='999', VALOR=float('nan')),
        linha_robo(UO_COD=2171),
    ])

    linhas, problemas = ler_linhas(df, HOJE)

    assert [n for n, _ in linhas] == [2, 4]
    assert problemas == []

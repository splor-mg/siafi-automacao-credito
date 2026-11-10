"""Leitura e validacao das linhas da aba ROBO, ANTES de tocar no SIAFI.

Antes a conversao de cada linha acontecia dentro do laco de solicitacoes: uma
celula vazia na linha 30 so estourava depois de 29 linhas ja digitadas, com um
documento aberto pela metade no SIAFI. E linha com a coluna TIPO vazia (a aba
ROBO devolve '999' nesse caso) era pulada em silencio — em 10/11/2026 tres
linhas ficaram sem envio e o robo ainda quebrou adiante com 'Keyboard locked',
tentando finalizar um documento que nunca foi aberto.

Aqui tudo e convertido e conferido de uma vez. Funcoes puras, sem SIAFI e sem
arquivo, para serem testaveis.
"""

import pandas as pd

TIPOS_VALIDOS = ('1', '2', '3', '4')
ORIENTACOES_VALIDAS = ('Suplementar', 'Anular')

# Tipos de alteracao que so suplementam: a origem do recurso nao e outra
# dotacao, entao nao existe linha de anulacao. Os fluxos desses tipos nem
# digitam anulacao — a linha era pulada em silencio.
TIPOS_SO_SUPLEMENTACAO = {
    '2': 'Saldo Financeiro',
    '3': 'Excesso de arrecadação',
}

# O que a formula da aba ROBO devolve quando a coluna TIPO da planilha de
# origem esta vazia.
TIPO_NAO_PREENCHIDO = '999'


def _vazio(valor):
    if isinstance(valor, str):
        return valor.strip() == ''
    return pd.isna(valor)


def montar_data_row(row, hoje):
    """Converte uma linha da aba ROBO no dicionario que os fluxos digitam.

    Devolve (data_row, problemas). Com 'problemas' preenchido a linha NAO pode
    ser enviada: os campos com defeito ficam vazios no data_row.
    """
    problemas = []

    def cod(coluna, tamanho=0):
        try:
            valor = row[coluna]
        except KeyError:
            problemas.append(f'coluna {coluna} não existe na aba ROBO')
            return ''
        if _vazio(valor):
            problemas.append(f'{coluna} vazio')
            return ''
        try:
            return str(int(valor)).zfill(tamanho)
        except (ValueError, TypeError):
            problemas.append(f'{coluna} inválido ({valor})')
            return ''

    data_row = {}
    data_row['month'] = hoje.strftime("%m")
    data_row['day'] = hoje.strftime("%d")
    data_row['year'] = hoje.strftime("%Y")

    data_row['orientacao'] = str(row['ORIENTACAO']).strip()
    if data_row['orientacao'] not in ORIENTACOES_VALIDAS:
        problemas.append(f"ORIENTACAO '{data_row['orientacao']}' desconhecida "
                         f"(esperado Suplementar ou Anular)")

    data_row['uo']          = cod('UO_COD')
    data_row['acao']        = cod('ACAO_COD')
    data_row['funcao']      = cod('FUNCAO_COD', 2)
    data_row['subfuncao']   = cod('SUBFUNCAO_COD', 3)
    data_row['programa']    = cod('PROGRAMA_COD', 3)
    data_row['subprojeto']  = cod('SUBPROJETO_COD')
    data_row['categoria']   = cod('CATEGORIA_COD')
    data_row['grupo']       = cod('GRUPO_COD')
    data_row['modalidade']  = cod('MODALIDADE_COD')
    data_row['elemento']    = cod('ELEMENTO_COD', 2)
    data_row['iag']         = cod('IPG_COD')
    data_row['fonte']       = cod('FONTE_COD', 2)
    data_row['procedencia'] = cod('IPU_COD')

    # Na aba ROBO o VALOR vem em centavos e, nas anulacoes, negativo. O SIAFI
    # recebe sempre o valor positivo: o sinal e dado pela tela em que a linha
    # e digitada.
    try:
        valor = int(round(row['VALOR']))
        data_row['valor'] = str(-valor if data_row['orientacao'] == 'Anular'
                                else valor)
    except (ValueError, TypeError, KeyError):
        problemas.append('VALOR vazio ou inválido')
        data_row['valor'] = ''

    try:
        suplementada = row['UO_SUPLEMENTADA']
        data_row['uo_suplementada'] = ('0' if _vazio(suplementada)
                                       else str(int(suplementada)))
    except (ValueError, TypeError, KeyError):
        problemas.append('UO_SUPLEMENTADA inválida')
        data_row['uo_suplementada'] = '0'

    data_row['tipo'] = cod('TIPO')
    if data_row['tipo'] == TIPO_NAO_PREENCHIDO:
        problemas.append('coluna TIPO não preenchida na planilha de origem')
    elif data_row['tipo'] and data_row['tipo'] not in TIPOS_VALIDOS:
        problemas.append(f"TIPO {data_row['tipo']} desconhecido "
                         f"(o robô trata os tipos 1 a 4)")
    elif (data_row['tipo'] in TIPOS_SO_SUPLEMENTACAO
          and data_row['orientacao'] == 'Anular'):
        problemas.append(
            f"TIPO {data_row['tipo']} "
            f"({TIPOS_SO_SUPLEMENTACAO[data_row['tipo']]}) só aceita "
            f"suplementação, e a linha é de anulação")

    return data_row, problemas


def ler_linhas(df, hoje):
    """Converte a aba ROBO inteira.

    'df' e a aba lida pelo pandas, sem tratamento. Devolve (linhas, problemas):
    'linhas' e a lista de (linha_excel, data_row) na ordem da planilha, e
    'problemas' uma frase por linha com defeito — vazia quando esta tudo certo.
    """
    df = df.dropna(how='all').reset_index(drop=False)

    linhas = []
    problemas = []

    for _, row in df.iterrows():
        # UO_COD vazio: fim dos dados (a aba ROBO tem formulas ate a linha 1000)
        if _vazio(row['UO_COD']):
            continue

        linha_excel = int(row['index']) + 2  # +2: cabecalho na linha 1
        data_row, defeitos = montar_data_row(row, hoje)
        linhas.append((linha_excel, data_row))

        if defeitos:
            uo = data_row['uo'] or '?'
            problemas.append(f"linha {linha_excel} (UO {uo}): "
                             + '; '.join(defeitos))

    return linhas, problemas

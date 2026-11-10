"""Quais arquivos da pasta de origem entram na consolidacao."""
from consolida import e_planilha_de_origem


def test_aceita_xlsm_com_qualquer_caixa_na_extensao():
    assert e_planilha_de_origem('Alteracoes Orcamentarias LEMG.xlsm')
    assert e_planilha_de_origem('SEJUSP.XLSM')


def test_ignora_o_arquivo_de_trava_do_excel():
    """Planilha aberta no Excel deixa um '~$nome.xlsm' ao lado, que nao e
    planilha: a leitura dele derrubava a consolidacao inteira."""
    assert not e_planilha_de_origem('~$Alteracoes Orcamentarias LEMG.xlsm')


def test_ignora_o_que_nao_e_xlsm():
    assert not e_planilha_de_origem('copia.xlsx')
    assert not e_planilha_de_origem('leia-me.txt')

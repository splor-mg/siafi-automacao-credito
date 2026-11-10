"""O desfecho de cada documento no SIAFI e lido de uma tela de texto.

Quando o SIAFI recusa o documento, a tela nao tem numero de documento para ler
e tem mais de uma mensagem de erro possivel. Se o finalizar_documento tratar a
recusa como sucesso, ele quebra no meio do laco de solicitacoes e as linhas
seguintes da planilha (outras UOs, ja conferidas) nunca sao enviadas.

As telas abaixo vem dos logs de 05 e 06/11/2026. A recusa de anulacao e a
mensagem verbatim da execucao de 06/11 as 08:32 (UO 1101), que o log de 05/11
so mostrou cortada em '-S' porque o robo lia 46 caracteres da linha 1.
"""
import pytest

import resultado
import utils_siafi
from utils_siafi import finalizar_documento


class EmuladorFalso:
    """Emulador 3270 de mentira: devolve as telas combinadas no teste.

    string_get(linha, coluna, tamanho) e a unica leitura que o
    finalizar_documento faz, entao e so isso que precisa ser combinado.
    """

    def __init__(self, telas, menu_depois_de=0):
        self.telas = telas
        self.menu_depois_de = menu_depois_de
        self.pf_enviados = []
        self.campos = []

    def string_get(self, linha, coluna, tamanho):
        return self.telas.get((linha, coluna, tamanho), '')

    def string_found(self, linha, coluna, texto):
        # voltar_ao_menu procura a tela da Unidade Executora
        return len([p for p in self.pf_enviados if p == 3]) >= self.menu_depois_de

    def send_pf(self, n):
        self.pf_enviados.append(n)

    def send_enter(self):
        pass

    def wait_for_field(self):
        pass

    def fill_field(self, linha, coluna, valor, tamanho):
        self.campos.append((linha, coluna, valor))


@pytest.fixture(autouse=True)
def sem_espera_e_sem_historico(monkeypatch):
    """Zera o estado global do resultado e os sleeps de espera do mainframe."""
    monkeypatch.setattr(utils_siafi.time, 'sleep', lambda _s: None)
    monkeypatch.setattr(resultado, '_resultados', [])
    monkeypatch.setattr(resultado, '_confirmacao_enviada', False)
    monkeypatch.setattr(resultado, '_sem_resposta', None)
    resultado.marcar_linha(3)
    yield
    resultado.marcar_linha(None)


TELA_PENDENTE_COM_INCONSISTENCIA = {
    (1, 1, 46): '0017-TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANU',
    (1, 1, 80): '0017-TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANULAR',
    (4, 19, 46): 'Inconsistencia no Registro da Contabilizacao',
    (6, 1, 80): 'Nr. Documento:',
}

TELA_ERRO_DE_ANULACAO = {
    (1, 1, 46): 'ERRO- ANUL. GIFP 30101 VL.       17.159,00 -S',
    (1, 1, 80): 'ERRO- ANUL. GIFP 30101 VL.       17.159,00 -SEM SALDO COTA-COMUNICAR SUCOR.',
    (4, 19, 46): '',
    (6, 1, 80): 'Nr. Documento:',
}

TELA_REGISTRO_EFETUADO = {
    (1, 1, 46): '0011-REGISTRO EFETUADO.',
    (1, 1, 80): '0011-REGISTRO EFETUADO.',
    (4, 19, 46): '',
    (6, 1, 80): 'UO: 1631       Nr. Documento: 2026NL00123   Data: 05/11/2026',
}


def test_documento_recusado_e_registrado_uma_unica_vez():
    """A tela do dia 05/11 casava com as DUAS verificacoes de erro, e o
    documento da UO 1631 apareceu duplicado no resumo."""
    em = EmuladorFalso(TELA_PENDENTE_COM_INCONSISTENCIA)

    finalizar_documento(em, '2061', '1631', {})

    assert resultado._resultados == [(3, '1631', 'erro', '')]


def test_documento_recusado_nao_tenta_ler_numero_de_documento():
    """Sem numero na tela, o caminho de sucesso estourava IndexError e
    interrompia o laco: as UOs seguintes da planilha nao eram enviadas."""
    em = EmuladorFalso(TELA_ERRO_DE_ANULACAO)

    retorno, nr_doc = finalizar_documento(em, '1451', '2061', {})

    assert nr_doc == ''
    assert 'Erro' in retorno
    assert resultado._resultados == [(3, '2061', 'erro', '')]


def test_documento_aceito_registra_o_numero():
    em = EmuladorFalso(TELA_REGISTRO_EFETUADO)

    retorno, nr_doc = finalizar_documento(em, '1451', '1631', {})

    assert nr_doc == '2026NL00123'
    assert retorno == '0011-REGISTRO EFETUADO.'
    assert resultado._resultados == [(3, '1631', 'ok', '2026NL00123')]


def test_depois_da_recusa_confere_que_voltou_ao_menu(monkeypatch):
    """Tres PF3 cegos bastam depois de um registro efetuado, mas nao depois de
    uma recusa: a tela de partida pode ser outra."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    em = EmuladorFalso(TELA_ERRO_DE_ANULACAO, menu_depois_de=4)

    finalizar_documento(em, '1451', '2061', {})

    # Chegou na tela da Unidade Executora e reinformou a UO
    assert em.pf_enviados.count(3) >= 4
    assert any(linha == 22 and coluna == 30 for linha, coluna, _v in em.campos)


def test_sem_voltar_ao_menu_o_robo_para_em_vez_de_digitar_na_tela_errada(monkeypatch):
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    em = EmuladorFalso(TELA_ERRO_DE_ANULACAO, menu_depois_de=999)

    with pytest.raises(RuntimeError):
        finalizar_documento(em, '1451', '2061', {})

    # A recusa foi registrada antes de parar: o resumo e a planilha nao perdem
    # o desfecho desse documento.
    assert resultado._resultados == [(3, '2061', 'erro', '')]


def test_recusa_de_anulacao_explica_a_cota_aprovada(capsys, monkeypatch):
    """A analise de saldo le o saldo de dotacao, que nao desconta cota aprovada
    e nao empenhada: ela aprova anulacoes que o SIAFI recusa. O log precisa
    dizer onde esta a correcao, porque nao e na planilha."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_ERRO_DE_ANULACAO), '1451', '2061', {})

    saida = capsys.readouterr().out
    assert 'preso em cota aprovada e ainda não empenhada' in saida
    assert 'Anule a cota para liberar o crédito' in saida


def test_recusa_sem_anulacao_nao_sugere_cota(capsys, monkeypatch):
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_PENDENTE_COM_INCONSISTENCIA), '2061', '1631', {})

    assert 'cota aprovada' not in capsys.readouterr().out


def test_relato_da_recusa_de_anulacao_diz_o_motivo(capsys, monkeypatch):
    """O grupo do Telegram le o texto do relato, nao o console: ele precisa
    dizer por que o SIAFI recusou."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_ERRO_DE_ANULACAO), '1451', '2061', {})

    linhas = capsys.readouterr().out.splitlines()
    relato_doc = [l for l in linhas if l.startswith('UO 2061 | linha 3')]
    assert relato_doc == ['UO 2061 | linha 3 | ERRO | SIAFI recusou a anulação '
                          'por falta de saldo de cota — anule a cota para '
                          'liberar o crédito']


def test_relato_da_inconsistencia_contabil_diz_o_motivo(capsys, monkeypatch):
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_PENDENTE_COM_INCONSISTENCIA), '2061', '1631', {})

    linhas = capsys.readouterr().out.splitlines()
    relato_doc = [l for l in linhas if l.startswith('UO 1631 | linha 3')]
    assert relato_doc == ['UO 1631 | linha 3 | ERRO | inconsistência na '
                          'contabilização']


def test_relato_de_recusa_desconhecida_repassa_a_tela(capsys, monkeypatch):
    """Sem motivo conhecido, o grupo recebe a tela do SIAFI em vez de um
    'nao registrado' que nao ajuda ninguem."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    tela = {(1, 1, 46): '0099-MENSAGEM NOVA DO SIAFI',
            (1, 1, 80): '0099-MENSAGEM NOVA DO SIAFI',
            (4, 19, 46): '',
            (6, 1, 80): 'Nr. Documento:'}

    finalizar_documento(EmuladorFalso(tela), '1451', '2061', {})

    linhas = capsys.readouterr().out.splitlines()
    relato_doc = [l for l in linhas if l.startswith('UO 2061 | linha 3')]
    assert relato_doc == ['UO 2061 | linha 3 | ERRO | SIAFI respondeu '
                          '"0099-MENSAGEM NOVA DO SIAFI"']


def test_outra_recusa_de_anulacao_repassa_a_tela_sem_chutar_a_causa(capsys, monkeypatch):
    """Falta de saldo de COTA e so uma das recusas possiveis de anulacao.
    Nas outras o grupo recebe a mensagem do SIAFI, nao um palpite."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    tela = {(1, 1, 46): 'ERRO- ANUL. GIFP 30101 VL.       17.159,00 -D',
            (1, 1, 80): 'ERRO- ANUL. GIFP 30101 VL.       17.159,00 -DOTACAO BLOQUEADA',
            (4, 19, 46): '',
            (6, 1, 80): 'Nr. Documento:'}

    finalizar_documento(EmuladorFalso(tela), '1451', '2061', {})

    saida = capsys.readouterr().out
    assert 'cota' not in saida.lower()
    assert 'DOTACAO BLOQUEADA' in saida


# ---------------------------------------------------------------------------
# A fronteira da confirmacao
#
# Quando o envio para no meio, o login.py so devolve as planilhas para a fila
# se nenhuma confirmacao (PF5) chegou a ser enviada: ate ali e certo que nada
# foi registrado no SIAFI. Errar para menos deixa a planilha presa em
# Realizados; errar para mais faz o proximo acionamento duplicar documento.
# ---------------------------------------------------------------------------

class TecladoTravado(Exception):
    """Faz as vezes do py3270.CommandError('Keyboard locked')."""


class EmuladorQueNaoAvanca(EmuladorFalso):
    """O SIAFI recusou as linhas digitadas e nao foi para a tela da
    justificativa: a posicao (11, 11) e protegida e o emulador trava."""

    def fill_field(self, linha, coluna, valor, tamanho):
        raise TecladoTravado('Keyboard locked')


class EmuladorQueCaiDepoisDoPF5(EmuladorFalso):
    """A sessao cai entre a confirmacao e a leitura da resposta — o que houve
    em 22/10/2026 as 13:57 (BrokenPipeError lendo a linha 6)."""

    def string_get(self, linha, coluna, tamanho):
        if 5 in self.pf_enviados:
            raise BrokenPipeError('emulador caiu')
        return super().string_get(linha, coluna, tamanho)


def test_parada_antes_do_pf5_nao_conta_como_confirmacao():
    em = EmuladorQueNaoAvanca(TELA_REGISTRO_EFETUADO)

    with pytest.raises(TecladoTravado):
        finalizar_documento(em, '1451', '1631', {})

    assert em.pf_enviados == []
    assert not resultado.houve_confirmacao()
    assert resultado.confirmacao_sem_resposta() is None


def test_queda_depois_do_pf5_deixa_a_confirmacao_sem_resposta():
    """O documento pode ter sido registrado: o robo nao sabe, e precisa dizer
    qual UO conferir no SIAFI."""
    em = EmuladorQueCaiDepoisDoPF5(TELA_REGISTRO_EFETUADO)

    with pytest.raises(BrokenPipeError):
        finalizar_documento(em, '1451', '1631', {})

    assert resultado.houve_confirmacao()
    assert resultado.confirmacao_sem_resposta() == '1631'
    assert resultado._resultados == []


def test_documento_registrado_conta_como_confirmacao_com_resposta():
    finalizar_documento(EmuladorFalso(TELA_REGISTRO_EFETUADO), '1451', '1631', {})

    assert resultado.houve_confirmacao()
    assert resultado.confirmacao_sem_resposta() is None


def test_documento_recusado_tambem_conta_como_confirmacao(monkeypatch):
    """Recusa nao devolve planilha sozinha: o resumo manda conferir a UO no
    SIAFI antes de reprocessar."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_ERRO_DE_ANULACAO), '1451', '2061', {})

    assert resultado.houve_confirmacao()
    assert resultado.confirmacao_sem_resposta() is None


def test_recusa_copia_a_tela_inteira_para_o_log(capsys, monkeypatch):
    """A linha 1 nem sempre conta a historia toda: na recusa de 05/11 a
    inconsistencia contabil estava na linha 4."""
    import analise_saldo
    monkeypatch.setattr(analise_saldo.time, 'sleep', lambda _s: None)

    finalizar_documento(EmuladorFalso(TELA_PENDENTE_COM_INCONSISTENCIA), '2061', '1631', {})

    saida = capsys.readouterr().out
    assert 'TELA DO SIAFI: documento da UO 1631 recusado' in saida
    assert '  01| 0017-TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANULAR' in saida
    assert '  06| Nr. Documento:' in saida


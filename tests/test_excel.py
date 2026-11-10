"""O recalculo no Excel e o unico passo que depende do Windows, e quando ele
falha o motivo precisa chegar inteiro ao grupo.

Em 10/11/2026 as 10:40 o Excel devolveu um erro, o PowerShell respondeu na
pagina de codigo do console (cp850) e o robo morreu num UnicodeDecodeError
sem dizer nada. Pior: a instancia escondida do Excel ficou aberta segurando o
copia.xlsm, e as tres execucoes seguintes abriram o arquivo como somente
leitura, 'recalcularam' sem gravar e pararam em 'aba ROBO vazia'.
"""
import subprocess

import excel
from excel import _comando, _interpretar, recalcular_no_excel


def test_ok_e_sucesso():
    assert _interpretar(b'OK\r\n') is None


def test_ok_com_marca_de_utf8_no_inicio_tambem_e_sucesso():
    assert _interpretar(b'\xef\xbb\xbfOK\r\n') is None


def test_erro_do_excel_chega_legivel_e_numa_linha_so():
    resposta = ('ERRO: O Microsoft Excel não pode acessar o arquivo.\r\n\r\n'
                '• O arquivo está sendo usado por outro programa.\r\n')

    assert _interpretar(resposta.encode('utf-8')) == (
        'O Microsoft Excel não pode acessar o arquivo. • O arquivo está '
        'sendo usado por outro programa.')


def test_resposta_fora_de_utf8_nao_derruba_o_robo():
    """Os bytes reais de 10:40: 'ERRO: Exce' e o 0x87, que e o 'ç' em cp850."""
    motivo = _interpretar('ERRO: Exceção de HRESULT: 0x800A03EC'.encode('cp850'))

    assert motivo.startswith('Exce')
    assert 'HRESULT: 0x800A03EC' in motivo


def test_arquivo_aberto_em_outro_excel_vira_instrucao_para_quem_opera():
    motivo = _interpretar(b'ERRO: ABERTO_EM_OUTRO_EXCEL\r\n')

    assert 'já está aberto em outra instância do Excel' in motivo
    assert 'Gerenciador de Tarefas' in motivo


def test_sem_resposta_usa_o_que_o_powershell_disse_no_erro_padrao():
    assert _interpretar(b'', b'Falha ao iniciar o script\r\n') == (
        'Falha ao iniciar o script')
    assert _interpretar(b'', b'') == (
        'o PowerShell não devolveu nenhuma resposta')


def test_comando_fecha_o_excel_mesmo_quando_da_erro():
    """Sem o finally a instancia escondida fica viva segurando o arquivo, e
    toda execucao seguinte falha ate alguem matar o processo."""
    comando = _comando(r'\\wsl.localhost\Ubuntu\home\robo\data\copia.xlsm')

    assert 'finally { if ($xl) { try { $xl.Quit() } catch {} } }' in comando
    assert "if ($wb.ReadOnly) { throw 'ABERTO_EM_OUTRO_EXCEL' }" in comando
    assert 'UTF8Encoding' in comando
    # So ASCII: acento em comando do PowerShell 5.1 depende da pagina de codigo
    assert comando.isascii()


def test_comando_escapa_aspa_simples_do_caminho():
    assert "Open('C:\\pasta d''agua\\copia.xlsm')" in _comando(
        r"C:\pasta d'agua\copia.xlsm")


class _Resposta:
    def __init__(self, stdout, stderr=b''):
        self.stdout, self.stderr = stdout, stderr


def test_recalcular_devolve_none_quando_o_excel_responde_ok(monkeypatch):
    monkeypatch.setattr(subprocess, 'check_output', lambda *a, **k: b'C:\\copia.xlsm\n')
    monkeypatch.setattr(subprocess, 'run', lambda *a, **k: _Resposta(b'OK\r\n'))

    assert recalcular_no_excel('/repo/data/copia.xlsm') is None


def test_recalcular_devolve_o_motivo_quando_o_excel_falha(monkeypatch):
    monkeypatch.setattr(subprocess, 'check_output', lambda *a, **k: b'C:\\copia.xlsm\n')
    monkeypatch.setattr(subprocess, 'run',
                        lambda *a, **k: _Resposta('ERRO: Exceção X'.encode('cp850')))

    assert 'Exce' in recalcular_no_excel('/repo/data/copia.xlsm')


def test_recalcular_sem_powershell_devolve_o_motivo_em_vez_de_quebrar(monkeypatch):
    def sem_powershell(*a, **k):
        raise FileNotFoundError('powershell.exe')

    monkeypatch.setattr(subprocess, 'check_output', lambda *a, **k: b'C:\\copia.xlsm\n')
    monkeypatch.setattr(subprocess, 'run', sem_powershell)

    motivo = recalcular_no_excel('/repo/data/copia.xlsm')

    assert 'não consegui executar o PowerShell' in motivo
    assert excel.POWERSHELL in motivo

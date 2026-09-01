"""O resgate das planilhas mexe em arquivos de producao no OneDrive.

Se ele falhar em silencio, a pessoa corrige a dotacao, aciona de novo e recebe
'NENHUMA PLANILHA ENCONTRADA PARA PROCESSAR' sem entender por que.
"""
import os

from arquivos import devolver_planilhas_de_origem


def montar_cenario(tmp_path, nomes=('planilha A.xlsm', 'planilha B.xlsm')):
    """Simula o estado deixado pelo consolida.py: arquivos em Realizados e um
    manifesto dizendo de onde cada um veio."""
    origem = tmp_path / 'Remanejamentos'
    realizados = tmp_path / 'Realizados'
    origem.mkdir()
    realizados.mkdir()

    manifesto = tmp_path / '.movidos'
    linhas = []
    for nome in nomes:
        (realizados / nome).write_text('conteudo')
        linhas.append(f'{realizados / nome}\t{origem / nome}')
    manifesto.write_text('\n'.join(linhas) + '\n', encoding='utf-8')

    return origem, realizados, manifesto


def test_devolve_os_arquivos_para_a_pasta_de_origem(tmp_path):
    origem, realizados, manifesto = montar_cenario(tmp_path)

    devolvidos = devolver_planilhas_de_origem(str(manifesto))

    assert sorted(devolvidos) == ['planilha A.xlsm', 'planilha B.xlsm']
    assert sorted(p.name for p in origem.iterdir()) == ['planilha A.xlsm',
                                                        'planilha B.xlsm']
    assert list(realizados.iterdir()) == []


def test_apaga_o_manifesto_depois_de_devolver(tmp_path):
    """Manifesto que sobra faria a execucao seguinte tentar devolver de novo."""
    _origem, _realizados, manifesto = montar_cenario(tmp_path)

    devolver_planilhas_de_origem(str(manifesto))

    assert not manifesto.exists()


def test_sem_manifesto_nao_faz_nada_nem_quebra(tmp_path):
    """Falha antes da consolidacao mover qualquer coisa: nao ha o que desfazer."""
    assert devolver_planilhas_de_origem(str(tmp_path / 'nao-existe')) == []


def test_ignora_arquivo_que_ja_nao_esta_la(tmp_path):
    """Alguem pode ter mexido no Realizados na mao entre uma coisa e outra."""
    origem, realizados, manifesto = montar_cenario(tmp_path)
    (realizados / 'planilha A.xlsm').unlink()

    devolvidos = devolver_planilhas_de_origem(str(manifesto))

    assert devolvidos == ['planilha B.xlsm']
    assert [p.name for p in origem.iterdir()] == ['planilha B.xlsm']


def test_recria_a_pasta_de_origem_se_ela_sumiu(tmp_path):
    origem, _realizados, manifesto = montar_cenario(tmp_path)
    for p in origem.iterdir():
        p.unlink()
    origem.rmdir()

    devolvidos = devolver_planilhas_de_origem(str(manifesto))

    assert len(devolvidos) == 2
    assert origem.is_dir()

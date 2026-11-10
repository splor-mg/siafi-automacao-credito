"""Movimentacao de arquivos entre as pastas do robo."""
import os
import shutil


def devolver_planilhas_de_origem(manifesto):
    """Devolve os arquivos que a consolidacao tirou das pastas de origem.

    O consolida.py move os .xlsm para 'Realizados' logo no inicio, antes de a
    analise de saldo rodar — protecao deliberada contra reprocessar arquivos e
    duplicar documentos se o robo cair no meio das solicitacoes. Mas quando o
    robo para ANTES de enviar qualquer coisa (validacao da planilha, falha de
    conexao ou de login, analise de saldo reprovada), os originais precisam
    voltar para a fila: sem isso a pessoa corrige o problema, aciona de novo e
    recebe 'NENHUMA PLANILHA ENCONTRADA PARA PROCESSAR'.

    Le o manifesto gravado pelo consolida.py (uma linha por arquivo, no
    formato 'destino<TAB>origem') e desfaz cada movimentacao.

    Nunca levanta excecao: roda durante o tratamento de outra falha, e mascarar
    o erro original so atrapalharia o diagnostico. Devolve a lista de nomes que
    conseguiu devolver.
    """
    devolvidos = []
    try:
        with open(manifesto, encoding='utf-8') as f:
            pares = [l.rstrip('\n').split('\t') for l in f if '\t' in l]
    except OSError:
        return devolvidos

    for destino, origem in pares:
        try:
            if os.path.exists(origem):
                # Alguem ja pos outro arquivo com esse nome na fila (a versao
                # corrigida, provavelmente). Sobrescrever apagaria o trabalho
                # da pessoa sem aviso; o original fica em Realizados.
                print(f"[aviso] ja existe {origem}; mantive {destino} onde esta.")
                continue
            if os.path.exists(destino):
                os.makedirs(os.path.dirname(origem), exist_ok=True)
                shutil.move(destino, origem)
                devolvidos.append(os.path.basename(origem))
                print(f"Devolvido para a pasta de origem: {origem}")
        except OSError as e:
            print(f"[aviso] nao consegui devolver {destino}: {e}")

    try:
        os.remove(manifesto)
    except OSError:
        pass
    return devolvidos

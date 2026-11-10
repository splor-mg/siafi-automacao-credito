import re
import time

import resultado
from relato import relato
from tela import imprimir_tela

MSG_REGISTRO_EFETUADO = "0011-REGISTRO EFETUADO."
MSG_INCONSISTENCIA = 'Inconsistencia no Registro da Contabilizacao'

MOTIVO_SALDO_CONTABIL = "Erro de saldo contábil"
MOTIVO_GENERICO = "Erro"


class TelaDesconhecida(RuntimeError):
    """O terminal nao voltou ao menu depois de uma recusa.

    Classe propria para o login.py distinguir esta parada — que ja foi
    explicada ao grupo — de uma falha inesperada no meio do envio.
    """


def recusa_de_anulacao(tela):
    """A tela e a recusa de uma anulacao pelo SIAFI?

    Casa o prefixo do erro ('ERRO- ANUL. GIFP ...'), nao a palavra solta: a
    tela de documento pendente diz 'PF2 PARA ANULAR' e cairia aqui por engano.
    """
    return re.search(r'ERRO-\s*ANUL', tela.upper()) is not None


def sem_saldo_de_cota(tela):
    """A recusa e a do credito preso em cota.

    Mensagem real do SIAFI, de 06/11/2026:
        'ERRO- ANUL. GIFP 30101 VL. 17.159,00 -SEM SALDO COTA-COMUNICAR SUCOR.'

    Haver saldo de dotacao nao significa que o credito esta livre: parte dele
    pode ja ter sido aprovada em cota e ainda nao empenhada. A analise de saldo
    nao enxerga isso, entao a recusa so aparece aqui — e a correcao e anular a
    cota, nao mexer na planilha.
    """
    return recusa_de_anulacao(tela) and 'SEM SALDO COTA' in tela.upper()


def motivo_para_o_grupo(motivo, tela):
    """Frase curta e autoexplicativa para quem acompanha pelo Telegram.

    O relato mandava sempre 'nao registrado no SIAFI', e quem esta no grupo
    nao tem o console do robo para descobrir o porque.
    """
    if motivo == MOTIVO_SALDO_CONTABIL:
        return 'inconsistência na contabilização'
    if sem_saldo_de_cota(tela):
        return ('SIAFI recusou a anulação por falta de saldo de cota — '
                'anule a cota para liberar o crédito')
    if tela:
        return f'SIAFI respondeu "{tela}"'
    return 'nao registrado no SIAFI'


def numero_do_documento(linha_6):
    """Le o numero do documento na linha 6 da tela de confirmacao.

    Quando o SIAFI recusa o documento o rotulo 'Nr. Documento:' continua na
    tela, mas sem numero depois dele. Por isso aqui nao se assume que ha algo
    para ler: o split() de uma string vazia devolve lista vazia, e o [0] disso
    derrubava o robo no meio do laco de solicitacoes.
    """
    marcador = "Nr. Documento:"
    idx = linha_6.find(marcador)
    if idx == -1:
        return ""
    depois = linha_6[idx + len(marcador):].split()
    return depois[0] if depois else ""


def finalizar_documento(em, uo, uo_anterior, data_row):
    em.send_enter()
    em.wait_for_field()
    time.sleep(3)

    em.fill_field(11, 11, 'Remanejamento realizado conforme solicitado', 60)
    em.send_enter()
    em.wait_for_field()

    # Ate aqui nada foi registrado. Do PF5 em diante o SIAFI pode ter
    # registrado o documento, mesmo que o robo quebre antes de ler a tela.
    resultado.confirmacao_enviada(uo_anterior)
    em.send_pf(5)
    em.wait_for_field()
    em.send_pf(5)
    em.wait_for_field()
    time.sleep(3)

    em.wait_for_field()
    time.sleep(3)

    # Duas leituras, UMA decisao.
    #
    # As duas mensagens de recusa podem aparecer na MESMA tela — foi o que
    # houve com a UO 1631 em 05/11/2026: linha 1 em '0017-TECLE PF5 PARA
    # CONFIRMAR' e linha 4 com a inconsistencia contabil. Tratando cada uma no
    # seu proprio 'if', o mesmo documento era registrado duas vezes no resumo
    # e, pior, quando so a primeira casava o fluxo seguia para o caminho de
    # sucesso e tentava ler um numero de documento que a tela recusada nao tem.
    mensagem = em.string_get(1, 1, 46).strip()
    saldo_contabil = em.string_get(4, 19, 46).strip()

    if saldo_contabil == MSG_INCONSISTENCIA:
        motivo = MOTIVO_SALDO_CONTABIL
        aviso = f"Erro de saldo contábil na solicitação da UO {uo_anterior}"
    elif mensagem != MSG_REGISTRO_EFETUADO:
        motivo = MOTIVO_GENERICO
        aviso = f"Erro ao finalizar documento da UO {uo_anterior}"
    else:
        motivo = None

    if motivo is not None:
        print(aviso)
        # A linha 1 inteira, nao os 46 caracteres da comparacao: foi assim que
        # 'ERRO- ANUL. GIFP 30601 VL. 30.000,00 -S' chegou cortado no log, sem
        # dizer qual era o problema.
        tela = em.string_get(1, 1, 80).strip()
        print(f"  tela do SIAFI: {tela}")
        # A tela inteira, antes de sair dela: a linha 1 nem sempre conta a
        # historia toda (a inconsistencia contabil, por exemplo, vem na 4).
        imprimir_tela(em, f'documento da UO {uo_anterior} recusado')
        if sem_saldo_de_cota(tela):
            print("  O crédito existe, mas está preso em cota aprovada e "
                  "ainda não empenhada.")
            print("  Anule a cota para liberar o crédito e reprocesse esta UO.")
        retorno = motivo
        nr_doc = ""
        resultado.registrar_erro(uo_anterior,
                                 motivo_para_o_grupo(motivo, tela))
    else:
        retorno = em.string_get(1, 1, 80).strip()
        time.sleep(1)
        nr_doc = numero_do_documento(em.string_get(6, 1, 80).strip())
        if not nr_doc:
            # O SIAFI confirmou o registro: tratar como erro faria a pessoa
            # refazer um remanejamento que ja existe. Avisa e segue.
            print(f"Atenção: o SIAFI registrou o documento da UO "
                  f"{uo_anterior}, mas não foi possível ler o número na tela.")
        print(f"SIAFI retornou: {retorno}: UO {uo_anterior} - "
              f"Nº do documento: {nr_doc}")
        resultado.registrar_ok(uo_anterior, nr_doc)

    if motivo is None:
        em.send_pf(3)
        em.wait_for_field()
        em.send_pf(3)
        em.wait_for_field()
        em.send_pf(3)
        em.wait_for_field()
    else:
        _voltar_ao_menu_depois_da_recusa(em, uo_anterior)

    return retorno, nr_doc


def _voltar_ao_menu_depois_da_recusa(em, uo_anterior):
    """Confere que o terminal voltou ao menu antes do proximo documento.

    No caminho de sucesso tres PF3 bastam, porque a tela de partida e sempre a
    mesma. Depois de uma recusa ela pode ser outra (a UO 1631 ficou em
    '0017-TECLE PF5 PARA CONFIRMAR OU PF2 PARA ANULAR'), e digitar o documento
    seguinte numa tela desconhecida do SIAFI e o pior desfecho possivel. Entao
    aqui se CONFERE a chegada, em vez de contar PF3 — mesma estrategia que o
    login.py usa depois da analise de saldo.
    """
    # Import local: o utils_siafi e importado pelos quatro fluxos, e o
    # analise_saldo carrega pandas e py3270 no import.
    from analise_saldo import voltar_ao_menu

    if voltar_ao_menu(em):
        return

    imprimir_tela(em, f'sem voltar ao menu depois da recusa da UO {uo_anterior}')
    relato('erro',
           f'O SIAFI recusou o documento da UO {uo_anterior} e o robô não '
           f'conseguiu voltar ao menu depois disso. Parei antes de enviar os '
           f'documentos seguintes, para não digitar numa tela errada do SIAFI. '
           f'Confira essa UO no SIAFI antes de acionar de novo.')
    raise TelaDesconhecida(
        f'Terminal em tela desconhecida depois da recusa da UO {uo_anterior}')

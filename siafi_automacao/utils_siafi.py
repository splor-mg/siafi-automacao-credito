import time


def finalizar_documento(em, uo, uo_anterior, data_row):
    em.send_enter()
    em.wait_for_field()
    time.sleep(3)

    em.fill_field(11, 11, 'Remanejamento realizado conforme solicitado', 60)
    em.send_enter()
    em.wait_for_field()

    em.send_pf(5)
    em.wait_for_field()
    em.send_pf(5)
    em.wait_for_field()
    time.sleep(1)

    saldo_contabil = em.string_get(4, 19, 46).strip()
    if saldo_contabil == 'Inconsistencia no Registro da Contabilizacao':
        print(f"Erro de saldo contábil na solicitação da UO {uo_anterior}")
    else:
        retorno = em.string_get(1, 1, 80).strip()

        linha_6 = em.string_get(6, 1, 80).strip()

        marcador = "Nr. Documento:"
        idx = linha_6.find(marcador)
        if idx != -1:
            nr_doc = linha_6[idx + len(marcador):].strip().split()[0]
        else:
            nr_doc = ""

        print(f"SIAFI retornou: {retorno}: UO {uo_anterior} - Nº do documento: {nr_doc}")

    em.send_pf(3)
    em.wait_for_field()
    em.send_pf(3)
    em.wait_for_field()
    em.send_pf(3)
    em.wait_for_field()

    return retorno, nr_doc
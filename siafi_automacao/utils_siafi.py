import time


def finalizar_documento(em, uo, uo_anterior):
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

    retorno = em.string_get(1, 1, 80).strip()
    nr_doc = em.string_get(6, 39, 7).strip()
    print(f"SIAFI retornou: {retorno}: UO {uo_anterior} - Nº do documento: {nr_doc}")

    em.send_pf(3)
    em.wait_for_field()
    em.send_pf(3)
    em.wait_for_field()
    em.send_pf(3)
    em.wait_for_field()

    return retorno, nr_doc
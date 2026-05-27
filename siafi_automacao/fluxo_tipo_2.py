import time
from utils_siafi import finalizar_documento

def tipo_2(em, data_row, uo_anterior, orientacao_anterior, linha, conclusao):
    ## Verifica se é anulação ou aprovação e preencche 03-1 para aprovação e 04-1 para anulação
    retorno = ''

    while True:

        if uo_anterior != data_row['uo']:
            linha = 11
            print("Realizando procedimento Tipo 2")

            if uo_anterior != 0:
                conclusao = 1

            if conclusao == 1:
                retorno, nr_doc = finalizar_documento(em, data_row['uo'], uo_anterior)


            ### Abrir UO
            em.fill_field(21, 19, '10', 2)
            em.send_enter()
            em.wait_for_field()

            em.fill_field(21, 19, '05', 2)
            em.send_enter()
            em.wait_for_field()

            em.fill_field(21, 19, '25', 2)
            em.fill_field(21, 41, '1', 1)
            em.send_enter()
            em.wait_for_field()
        

            em.fill_field(9, 66, 'n', 1)
            em.fill_field(10, 59, data_row['uo'], 4)
            em.fill_field(11, 54, data_row['day'], 2)
            em.fill_field(11, 59, data_row['month'], 2)
            em.fill_field(11, 64, data_row['year'], 4)
            em.send_enter()
            em.wait_for_field()
            
            em.send_pf(5)  # envia F5 para incluir a autorização de abrir a UO
            em.send_pf(3)  # envia F3 para voltar para a tela de movimentação orçamentária
            em.wait_for_field()
            em.send_pf(3)  # envia F3 para voltar para a tela de movimentação orçamentária
            em.wait_for_field()
            em.send_pf(3)  # envia F3 para voltar para a tela de movimentação orçamentária
            em.wait_for_field()
            
            ### Início do processo de realização do remanejamento de crédito ###
            # Entrar em 03 - Movimentacao Orcamentaria
            em.fill_field(21, 19, '03', 2)
            em.send_enter()
            em.wait_for_field()

            #Entrar em 01 - Alteração orçamentária
            em.fill_field(21, 19, '01', 2)
            em.send_enter()
            em.wait_for_field()
            
            # Movimentação de tela
            em.fill_field(22, 19, '01', 2)
            em.fill_field(22, 41, '1', 1)
            em.send_enter()
            em.wait_for_field()

            em.fill_field(11, 8, 'x', 1)
            em.send_enter()
            em.wait_for_field()

            em.fill_field(9, 51, 'x', 1)
            em.send_enter()
            em.wait_for_field()

            em.fill_field(10, 68, data_row['uo'], 4)
            em.send_enter()
            em.wait_for_field()

        if linha == 21:
            em.send_pf(8)  # envia F8 para ir para a próxima página
            em.wait_for_field()
            retorno = em.string_get(1, 1, 80).strip()
            linha = 11
            ## fazer leitura de possiveis erros e usar if

        if orientacao_anterior != data_row['orientacao']:
            em.send_enter()
            em.wait_for_field()
            linha = 11

        
        if data_row['orientacao'] == 'Suplementar':
            em.fill_field(linha, 3, data_row['funcao'], 2) # função
            em.fill_field(linha, 6, data_row['subfuncao'], 3) # subfunção
            em.fill_field(linha, 10, data_row['programa'], 3) # programa
            em.fill_field(linha, 14, data_row['acao'], 4) # ação
            em.fill_field(linha, 19, '0001', 4)
            em.fill_field(linha, 25, f"{data_row['categoria']}{data_row['grupo']}{data_row['modalidade']}{data_row['elemento']}", 6) # natureza de despesa
            em.fill_field(linha, 32, data_row['iag'], 1) # procedência
            em.fill_field(linha, 35, data_row['fonte'], 2) # fonte
            em.fill_field(linha, 40, data_row['procedencia'], 1) # procedência
            em.fill_field(linha, 47, '4', 1) #indica que é suplementação por saldo financeiro

            if data_row['procedencia'] == '2':
                em.fill_field(linha, 56, data_row['uo_suplementada'], 4) # uo_financiadora
            em.fill_field(linha, 64, data_row['valor'], 15) # valor

        elif data_row['orientacao'] == 'Anular':
            break


        linha += 1

        if linha == 21:
            em.send_pf(8)  # envia F8 para ir para a próxima página
            em.wait_for_field()
        
        break

    return retorno, linha, conclusao
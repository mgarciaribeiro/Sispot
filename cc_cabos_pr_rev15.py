"""
Arquivo para cálculo de curto-circuito em cabos para-raios
revisão 00 - revisão inicial considerando circuitos mais complexos
revisão 01 - escreve vãos
revisão 02 - escreve vãos considerando diversas infos
revisão 03 - escreve em sistema com 2 barras
revisão 04 - escreve sistema com seccionamentos
revisão 05 - corrige problema relativo a T_1 ser torre de seccionamento
revisão 06 - Insere ramais de seccionamento
revisão 07 - Insere aterramento dos para-raios na SE
revisão 08 - Insere controle do OpenDSS com funções de estudo de para-raios
revisão 09 - Insere controle na execução do arquivo (só montar modelo, só rodar estudo, etc)
revisão 10 - Permite opção de sair do pgraoama após tentativa de validação
revisão 11 - Corrige forma de pegar resultados
revisão 12 - Corrige conexão no seccionamento
revisão 13 - Melhora conexão na torre de seccionamento com o ramal
revisão 14 - Melhora representação no seccionamento
revisão 15 - Corrige bug no ramal de seccionamento
Última revisão elaborado por Matheus Garcia Ribeiro - Planejamento da Expansão
contato: mgribeiro@isacteep.com.br
"""

import math
import cmath
import numpy as np
import pandas as pd
import win32com.client
from collections import defaultdict
import matplotlib.pyplot as plt


def caminho(de, para, linecodes, matrix, subestacoes, derivacoes, torres, arquivo):
    """
    Função para escrever no OpenDSS trechos de linha.
    :param de: Barra de origem
    :param para: Barra de destino
    :param linecodes: Lista de linecodes disponíveis
    :param matrix: Qualquer matriz de parâmetros (recebe xmatrix)
    :param subestacoes: Lista de barras restantes para modelagem que são subestações
    :param derivacoes: Lista de barras restantes para modelagem que são torres de derivação (não utilizado, recebe vazio)
    :param torres: Vetor contendo número de torres e torres de seccionamento
    :param: arquivo: Nome do arquivo do OpenDSS que será escrito
    :return: Retorna lista contendo na posição 0 o número de torres utilizadas e nas posições restantes dicionários contendo
             numeração das torres de seccionamento e circuito
    """
    if de in subestacoes or de in derivacoes:
        if de in subestacoes:
            codigo_linha = 0
            # Escreve trecho entre SE e sua torre de entrada. Houve um equivo inicial na interpretação do modelo por parte
            # do Matheus e esta primeira torre, que costuma ser mais perto da SE está sendo chamada de pórtico
            l_trecho = (10 ** -3) * float(input(f'Digite a distância entre a SE {de} a torre de entrada em [m]: '))
            r_at_se = float(input(f'Digite a resistência de aterramento a ser considerada para a SE {de} em [ohm]: '))
            r_at = float(input(f'Digite a resistência de aterramento a ser considerada para a primeira torre torre em [ohm]: '))

            while codigo_linha not in linecodes:
                print('\nLinecodes disponíveis: ')
                for i in range(len(linecodes)):
                    print(f'{linecodes[i]}')
                codigo_linha = input('Digite o linecode do trecho: ')

            portico = 'T_port_SE_' + str(de)
            arquivo.write(f'new line.{str(de) + portico} bus1={de} bus2={portico} ' 
                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                          f' length= {l_trecho} \n')
            # Escreve pórticos (torres de entrada) com respectivos aterramentos
            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4:
                node_1 = 4
                arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{de} bus1={de}.{node_1}.0 r={r_at_se} \n')
            elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
                node_1 = 4
                node_2 = 5
                arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                              f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{de} bus1={de}.{node_1} bus2={de}.{node_2} r=0 \n'
                              f'new Fault.RSE{de} bus1={de}.{node_2}.0 r={r_at_se} \n')
            elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                node_1 = 7
                arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{de} bus1={de}.{node_1}.0 r={r_at_se} \n')
                # Faz conexão dos equivalentes com demais fases
                arquivo.write(f'new Fault.Eq_{de}_A bus1={de}.1 bus2={de}.4 \n'
                              f'new Fault.Eq_{de}_B bus1={de}.2 bus2={de}.5 \n'
                              f'new Fault.Eq_{de}_C bus1={de}.3 bus2={de}.6 \n')
            else:
                node_1 = 7
                node_2 = 8
                arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                              f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{de} bus1={de}.{node_1} bus2={de}.{node_2} r=0 \n'
                              f'new Fault.RSE{de} bus1={de}.{node_2}.0 r={r_at_se} \n')
                # Faz conexão dos equivalentes com demais fases
                arquivo.write(f'new Fault.Eq_{de}_A bus1={de}.1 bus2={de}.4 \n'
                              f'new Fault.Eq_{de}_B bus1={de}.2 bus2={de}.5 \n'
                              f'new Fault.Eq_{de}_C bus1={de}.3 bus2={de}.6 \n')
            subestacoes.remove(de)
            torre_de = portico

        else:
            derivacoes.remove(de)
            torre_de = 'der_' + str(de)

    if para in subestacoes or para in derivacoes:
        if para in subestacoes:
            codigo_linha = 0
            # Escreve trecho entre SE e seu pórtico
            l_trecho = (10 ** -3) * float(input(f'Digite a distância entre a SE {para} e a torre de entrada em [m]: '))
            r_at_se = float(input(f'Digite a resistência de aterramento a ser considerada para a SE {para} em [ohm]: '))
            r_at = float(input(f'Digite a resistência de aterramento a ser considerada para a primeira torre em [ohm]: '))

            while codigo_linha not in linecodes:
                print('\nLinecodes disponíveis: ')
                for i in range(len(linecodes)):
                    print(linecodes[i])
                codigo_linha = input('Digite o linecode do trecho: ')

            portico = 'T_port_SE_' + str(para)
            arquivo.write(f'new line.{str(para) + portico} bus1={para} bus2={portico} ' 
                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                          f' length= {l_trecho} \n')
            # Escreve pórticos com respectivos aterramentos
            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4:
                node_1 = 4
                arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{para} bus1={para}.{node_1}.0 r={r_at_se} \n')
            elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
                node_1 = 4
                node_2 = 5
                arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                              f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{para} bus1={para}.{node_1} bus2={para}.{node_2} r=0 \n'
                              f'new Fault.RSE{para} bus1={para}.{node_2}.0 r={r_at_se} \n')
            elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                node_1 = 7
                arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{para} bus1={para}.{node_1}.0 r={r_at_se} \n')
                # Faz conexão dos equivalentes com demais fases
                arquivo.write(f'new Fault.Eq_{para}_A bus1={para}.1 bus2={para}.4 \n'
                              f'new Fault.Eq_{para}_B bus1={para}.2 bus2={para}.5 \n'
                              f'new Fault.Eq_{para}_C bus1={para}.3 bus2={para}.6 \n')
            else:
                node_1 = 7
                node_2 = 8
                arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                              f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
                # Escreve aterramento da SE
                arquivo.write(f'new Fault.R_SE_{para} bus1={para}.{node_1} bus2={para}.{node_2} r=0 \n'
                              f'new Fault.RSE{para} bus1={para}.{node_2}.0 r={r_at_se} \n')
                # Faz conexão dos equivalentes com demais fases
                arquivo.write(f'new Fault.Eq_{para}_A bus1={para}.1 bus2={para}.4 \n'
                              f'new Fault.Eq_{para}_B bus1={para}.2 bus2={para}.5 \n'
                              f'new Fault.Eq_{para}_C bus1={para}.3 bus2={para}.6 \n')
            subestacoes.remove(para)
            torre_para = portico

        else:
            derivacoes.remove(para)
            torre_para = 'der_' + str(para)

    l = float(input(f'\nDigite a distância em [km] entre a torre {torre_de} e a torre {torre_para}: '))
    l_rest = l
    trecho = 0
    trecho_lt = range(0, 0)
    torre_sec = {'torre': 0, 'circuito': 0}

    # Tolerancia de 10 metros
    while l_rest > 0.01:
        l_trecho = float(input(f'\nDigite a distância em [km] entre a torre {torre_de} e a torre a ser modelada com configuração (linecode) específica. \n'
                               f'Caso exista apenas uma configuração entre a torre {torre_de} e torre {torre_para}, será necessário entrar uma única vez com '
                               f'essa distância e seu valor será igual a {l}. \n'
                               f'Caso existam múltiplas configurações neste trecho, a soma dos trechos deverá ser igual a {l}. \n'
                               f'Ainda faltam {l_rest} km \n'
                               f'Caso não seja o primeiro trecho, digite o comprimento em [km] do trecho {trecho + 1} \n'
                               f'Trecho sendo escrito: {trecho + 1} \n'))
        l_rest = l_rest - l_trecho

        if l_rest < 0:
            print(f'\nHá uma inconsistência na entrada de dados, os comprimentos inseridos excedem o trecho total em {abs(l_rest)} km. \n'
                      f'Caso não queira conviver com esse erro, reinicie o programa.')
            # Vão médio

        vao_m = (10 ** -3) * round(float(input(f'\nDigite o comprimento do vão médio em [m] do trecho {trecho + 1}: ')), 2)
        trecho = trecho + 1

        torres_aux = int(torres[0])
        torres[0] = int(torres[0] + l_trecho / vao_m + 1 - 1)

        codigo_linha = 0
        # Escreve trecho

        r_at = float(input(f'\nDigite a resistência de aterramento a ser considerada para as torres em [ohm]: '))

        while codigo_linha not in linecodes:
            print('Linecodes disponíveis: \n')
            for i in range(len(linecodes)):
                print(linecodes[i])

            # Do início até ponto de troca de linecode (ponto intermediário)
            if torres_aux == 0 and l_rest != 0:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre {torre_de} e a torre T_{str(torres[0])}: \n')
                if codigo_linha in linecodes:
                    trecho_lt = range(torres_aux + 1, torres[0])
                    trecho_at = range(torres_aux + 1, torres[0] + 1)

                    # Corrige problema relativo a torre T_1 poder ser torre de seccionamento
                    bug_t_1 = input(f'A torre T_1 é torre de seccionamento? \n'
                                    f'Vale lembrar que a primeira torre após o pórtico recebe o nome de "T_port" (pois é). \n'
                                    f'Então, na prática, a torre T_1 é a segunda torre do modelo. \n'
                                    f'1 - Não (Default) \n'
                                    f'2 - Sim \n')
                    if bug_t_1 == "2":
                        print('\nAtualmente, existe uma limitação da ferramenta. Você deverá corrigir posteriormente a conexão \n'
                              'na torre de seccionamento T_1 de forma manual. \n'
                              f'Não há conexão entre o pórtico {torre_de} e T_1.')

                    else:
                        arquivo.write(f'new line.{str(torre_de) + "_T" + str(torres_aux + 1)} bus1={torre_de} bus2=T_{torres_aux + 1} '
                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                      f' length= {vao_m} \n')

                    #Constatação de seccionamento no trecho
                    sec = input(f'\nHá seccionamento no trecho que está sendo descrito (trecho {trecho})? \n'
                                f'0 - Não (Default) \n'
                                f'1 - Sim \n')
                    if sec == "1":
                        n_sec = int(input(f'\nDigite o número de torres de seccionamento no trecho {trecho}: '))
                        for i in range(n_sec):
                            l_sec = round(float(input(f'Digite a distância entre a torre {torre_de} e a torre de seccionamento {i + 1} em [km]: ')), 2)

                            # Verifica entrada desta distância
                            while l_sec > l_trecho:
                                l_sec = round(float(input(f'\nA distância até o ponto de seccionamento é superior à distância do trecho completo. \n'
                                                          f'Trecho: {l_trecho} km vs Distância ao seccionamento: {l_sec} km \n'
                                                          f'Entre novamente com a distância ao ponto de seccionamento em [km]: ')), 2)

                            torre_sec['torre'] = int(l_sec / vao_m)

                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7 or len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:

                                circuito_sec = input('\nDigite o circuito a ser seccionado: \n'
                                                     '1 - Circuito 1 (Default) \n'
                                                     '2 - Circuito 2 \n')

                                if circuito_sec == "2":
                                    circuito_sec = int(circuito_sec)
                                else:
                                    circuito_sec = 1

                                torre_sec['circuito'] = circuito_sec

                            else:
                                torre_sec['circuito'] = 1

                            torres.append(torre_sec.copy())

                            t_compara = []
                            for m in torres[1: len(torres)]:
                                t_compara.append(m['torre'])

                            print(f"\nTorre de seccionamento número {i + 1} definida no trecho entre torre {torre_de} e torre T_{torres[0]}: T_{torre_sec['torre']}")

                    for j in trecho_lt:
                        if sec != "1":
                            arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                          f' length= {vao_m} \n')
                        else:
                            imprime = 0
                            for k in torres[1: len(torres)]:
                                if j == k['torre'] - 1:
                                    # Insere torre de seccionamento
                                    seccionamento(j, k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    imprime = 1

                                elif k['torre'] == 1 and j == 1:
                                    # Insere torre de seccionamento
                                    seccionamento(torre_de[2:], k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                  f' length= {vao_m} \n')
                                    imprime = 1

                                else:
                                    while imprime == 0 and j + 1 not in t_compara:
                                        arquivo.write(f'new line.T_{j}_T{j+1} bus1=T_{j} bus2=T_{j + 1} '
                                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                      f' length= {vao_m} \n')
                                        imprime = 1


            # Do início até o final especificado
            elif torres_aux == 0 and l_rest == 0:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre {torre_de} e a torre {torre_para}: \n')
                if codigo_linha in linecodes:
                    torres[0] = torres[0] - 1
                    trecho_lt = range(torres_aux + 1, torres[0])
                    trecho_at = range(torres_aux + 1, torres[0] + 1)

                    # Corrige problema relativo a torre T_1 poder ser torre de seccionamento
                    bug_t_1 = input(f'A torre T_1 é torre de seccionamento? \n'
                                    f'Vale lembrar que a primeira torre após o pórtico recebe o nome de "T_port" (pois é). \n'
                                    f'Então, na prática, a torre T_1 é a segunda torre do modelo. \n'
                                    f'1 - Não (Default) \n'
                                    f'2 - Sim \n')
                    if bug_t_1 == "2":
                        print(
                            '\nAtualmente, existe uma limitação da ferramenta. Você deverá corrigir posteriormente a conexão \n'
                            'na torre de seccionamento T_1 de forma manual. \n'
                            f'Não há conexão entre o pórtico {torre_de} e T_1.')

                    else:
                        arquivo.write(f'new line.{str(torre_de) + "_T" + str(torres_aux + 1)} bus1={torre_de} bus2=T_{torres_aux + 1} '
                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                      f' length= {vao_m} \n')

                    # Constatação de seccionamento no trecho
                    sec = input(f'\nHá seccionamento no trecho que está sendo descrito (trecho {trecho}?) \n'
                                f'0 - Não (Default) \n'
                                f'1 - Sim \n')
                    if sec == "1":
                        n_sec = int(input(f'\nDigite o número de torres de seccionamento no trecho {trecho}: '))
                        for i in range(n_sec):
                            l_sec = round(float(input(
                                f'Digite a distância entre a torre {torre_de} e a torre de seccionamento {i + 1} em [km]: ')),
                                                       2)

                            # Verifica entrada desta distância
                            while l_sec >= l_trecho:
                                l_sec = round(float(input(
                                    f'\nA distância até o ponto de seccionamento é superior à distância do trecho completo. \n'
                                    f'Trecho: {l_trecho} km vs Distância ao seccionamento: {l_sec} km \n'
                                    f'Entre novamente com a distância ao ponto de seccionamento em [km]: ')), 2)

                            torre_sec['torre'] = int(l_sec / vao_m)
                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7 or len(
                                    rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:

                                circuito_sec = input('\nDigite o circuito a ser seccionado: \n'
                                                     '1 - Circuito 1 (Default) \n'
                                                     '2 - Circuito 2 \n')

                                if circuito_sec == "2":
                                    circuito_sec = int(circuito_sec)
                                else:
                                    circuito_sec = 1

                                torre_sec['circuito'] = circuito_sec

                            else:
                                torre_sec['circuito'] = 1

                            torres.append(torre_sec.copy())

                            t_compara = []
                            for m in torres[1: len(torres)]:
                                t_compara.append(m['torre'])

                            print(
                                f"\nTorre de seccionamento número {i + 1} definida no trecho entre torre {torre_de} e torre {torre_para}: T_{torre_sec['torre']}")

                    for j in trecho_lt:
                        if sec != "1":
                            arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                          f' length= {vao_m} \n')

                        else:
                            imprime = 0
                            for k in torres[1: len(torres)]:
                                if j == k['torre'] - 1:
                                    # Insere torre de seccionamento
                                    seccionamento(j, k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    imprime = 1

                                elif k['torre'] == 1 and j == 1:
                                    # Insere torre de seccionamento
                                    seccionamento(torre_de[2:], k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                  f' length= {vao_m} \n')
                                    imprime = 1

                                else:
                                    while imprime == 0 and j + 1 not in t_compara:
                                        arquivo.write(f'new line.T_{j}_T{j+1} bus1=T_{j} bus2=T_{j + 1} '
                                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                      f' length= {vao_m} \n')
                                        imprime = 1

                    arquivo.write(f'new line.T_{torres[0]}_{torre_para} bus1=T_{torres[0]} bus2={torre_para} '
                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                  f' length= {vao_m} \n')

            # De algum ponto intermediário até outro ponto intermediário
            elif torres_aux != 0 and l_rest > 0.01:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre T_{torres_aux} e a torre T_{str(torres[0])}: \n')
                if codigo_linha in linecodes:
                    trecho_lt = range(torres_aux, torres[0])
                    trecho_at = range(torres_aux + 1, torres[0] + 1)

                    # Constatação de seccionamento no trecho
                    sec = input(f'\nHá seccionamento no trecho que está sendo descrito (trecho {trecho}?) \n'
                                f'0 - Não (Default) \n'
                                f'1 - Sim \n')
                    if sec == "1":
                        n_sec = int(input(f'\nDigite o número de torres de seccionamento no trecho {trecho}: '))
                        for i in range(n_sec):
                            l_sec = round(float(input(
                                f'Digite a distância entre a torre T_{torres_aux} e a torre de seccionamento {i + 1} em [km]: ')),
                                                       2)

                            # Verifica entrada desta distância
                            while l_sec > l_trecho:
                                l_sec = round(float(input(
                                    f'\nA distância até o ponto de seccionamento é superior à distância do trecho completo. \n'
                                    f'Trecho: {l_trecho} km vs Distância ao seccionamento: {l_sec} km \n'
                                    f'Entre novamente com a distância ao ponto de seccionamento em [km]: ')), 2)


                            torre_sec['torre'] = torres_aux + int(l_sec / vao_m)

                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7 or len(
                                    rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:

                                circuito_sec = input('\nDigite o circuito a ser seccionado: \n'
                                                     '1 - Circuito 1 (Default) \n'
                                                     '2 - Circuito 2 \n')

                                if circuito_sec == "2":
                                    circuito_sec = int(circuito_sec)
                                else:
                                    circuito_sec = 1

                                torre_sec['circuito'] = circuito_sec

                            else:
                                torre_sec['circuito'] = 1

                            torres.append(torre_sec.copy())

                            t_compara = []
                            for m in torres[1: len(torres)]:
                                t_compara.append(m['torre'])

                            print(
                                f"\nTorre de seccionamento número {i + 1} definida no trecho entre torre T_{torres_aux} e torre T_{torres[0]}: T_{torre_sec['torre']} \n")

                    for j in trecho_lt:
                        if sec != "1":
                            arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                          f' length= {vao_m} \n')
                        else:
                            imprime = 0
                            for k in torres[1: len(torres)]:
                                if j == k['torre'] - 1:
                                    # Insere torre de seccionamento
                                    seccionamento(j, k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    imprime = 1

                                else:
                                    while imprime == 0 and j + 1 not in t_compara:
                                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                      f' length= {vao_m} \n')
                                        imprime = 1

            # De algum ponto intermediário até o final especificado
            else:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre T_{torres_aux} e a torre {torre_para}: \n')
                if codigo_linha in linecodes:
                    torres[0] = torres[0] - 1
                    trecho_lt = range(torres_aux, torres[0])
                    print(trecho_lt)
                    trecho_at = range(torres_aux + 1, torres[0] + 1)

                    # Constatação de seccionamento no trecho
                    sec = input(f'\nHá seccionamento no trecho que está sendo descrito (trecho {trecho})? \n'
                                f'0 - Não (Default) \n'
                                f'1 - Sim \n')
                    if sec == "1":
                        n_sec = int(input(f'\nDigite o número de torres de seccionamento no trecho {trecho}: '))
                        for i in range(n_sec):
                            l_sec = round(float(input(
                                f'Digite a distância entre a torre T_{torres_aux} e a torre de seccionamento {i + 1} em [km]: ')),
                                2)

                            # Verifica entrada desta distância
                            while l_sec >= l_trecho:
                                l_sec = round(float(input(
                                    f'\nA distância até o ponto de seccionamento é superior à distância do trecho completo. \n'
                                    f'Trecho: {l_trecho} km vs Distância ao seccionamento: {l_sec} km \n'
                                    f'Entre novamente com a distância ao ponto de seccionamento em [km]: ')), 2)

                            torre_sec['torre'] = torres_aux + int(l_sec / vao_m)

                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7 or len(
                                    rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:

                                circuito_sec = input('\nDigite o circuito a ser seccionado: \n'
                                                     '1 - Circuito 1 (Default) \n'
                                                     '2 - Circuito 2 \n')

                                if circuito_sec == "2":
                                    circuito_sec = int(circuito_sec)
                                else:
                                    circuito_sec = 1

                                torre_sec['circuito'] = circuito_sec

                            else:
                                torre_sec['circuito'] = 1

                            torres.append(torre_sec.copy())
                            t_compara = []
                            for m in torres[1: len(torres)]:
                                t_compara.append(m['torre'])

                            print(
                                f"\nTorre de seccionamento número {i + 1} definida no trecho entre torre T_{torres_aux} e torre {torre_para}: T_{torre_sec['torre']} \n")

                    for j in trecho_lt:
                        if sec != "1":
                            arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                          f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                          f' length= {vao_m} \n')
                        else:
                            imprime = 0
                            for k in torres[1: len(torres)]:
                                if j == k['torre'] - 1:
                                    # Insere torre de seccionamento
                                    seccionamento(j, k, linecodes, codigo_linha, matrix, vao_m, arquivo)
                                    imprime = 1

                                else:
                                    while imprime == 0 and j + 1 not in t_compara:
                                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                      f' length= {vao_m} \n')
                                        imprime = 1

                    arquivo.write(f'new line.T_{torres[0]}_{torre_para} bus1=T_{torres[0]} bus2={torre_para} '
                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                  f' length= {vao_m} \n')

        # Escreve torres com respectivos aterramentos
        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4:
            node_1 = 4
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1}.0 r={r_at} \n')
            if torre_para[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_para} bus1=T{torre_para}.{node_1}.0 r={r_at} \n')

        elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
            node_1 = 4
            node_2 = 5
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1} bus2=T_{j}.{node_2} r=0 \n'
                              f'new Fault.RT{j} bus1=T_{j}.{node_2}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1} bus2={torre_de}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_de} bus1=T{torre_de}.{node_2}.0 r={r_at} \n')
            if torre_para[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{j} bus1={torre_para}.{node_1} bus2={torre_para}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_para} bus1={torre_para}.{node_2}.0 r={r_at} \n')

        elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
            node_1 = 7
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1}.0 r={r_at} \n')
            if torre_para[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_para} bus1={torre_para}.{node_1}.0 r={r_at} \n')

        else:
            node_1 = 7
            node_2 = 8
            for j in trecho_at:
                arquivo.write(f'new Fault.R{j} bus1=T_{j}.{node_1} bus2=T_{j}.{node_2} r=0 \n'
                              f'new Fault.RT{j} bus1=T_{j}.{node_2}.0 r={r_at} \n')

            if torre_de[0:4] == 'der':
                arquivo.write(f'new Fault.R{torre_de} bus1={torre_de}.{node_1} bus2={torre_de}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_de} bus1={torre_de}.{node_2}.0 r={r_at} \n')

            if torre_para[0:4] == 'der':
                arquivo.write(f'new Fault.R{torre_para} bus1={torre_para}.{node_1} bus2={torre_para}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_para} bus1={torre_para}.{node_2}.0 r={r_at} \n')
#    print(torres)
    return torres

# -------------------------------------------------------------------------------------------------- #


def seccionamento(torre_de, torre_sec, linecodes, codigo_linha, matrix, vao_m, arquivo):
    """
    Função que insere torre de seccionamento em uma LT
    :param torre_de: Torre de origem até a torre que é uma torre de seccionamento
    :param torre_sec: Dicionário indicando número da torre de seccionamento e circuito
    :linecodes: todos linecodes possíveis
    :param linecode: Configuração (linecode) do trecho de torre_de para torre_sec
    :param matrix: Matriz do linecode para entender nós da torre
    :param vao_m: Vão médio para escrever trecho de torre_de até torre_sec
    :param arquivo: Arquivo .dss que será alterado
    :return: Não retorna nada
    """
    # Insere trecho torre_de para torre_sec

    arquivo.write(f"new line.T_{torre_de}_T{torre_sec['torre']} bus1=T_{torre_de} bus2=T_{torre_sec['torre']}_X "
                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                  f' length= {vao_m} \n')

    # Circuito seccionado é o 1
    if torre_sec['circuito'] == 1:

        # Caso seja seccionamento do circuito 1 em torre de circuito simples
        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4 or len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
            arquivo.write(
                f"new line.T_{torre_sec['torre']}_aux_aberto bus1=T_{torre_sec['torre']}_X.1.2.3 bus2=T_{torre_sec['torre']}.1.2.3 "
                f'phases=3 r1=999999 x1=99999 r0=999999 x0=99999 \n')

            arquivo.write(
                f"new line.T_{torre_sec['torre']}_4 bus1=T_{torre_sec['torre']}_X.4 bus2=T_{torre_sec['torre']}.4 "
                f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
                arquivo.write(
                    f"new line.T_{torre_sec['torre']}_5 bus1=T_{torre_sec['torre']}_X.5 bus2=T_{torre_sec['torre']}.5 "
                    f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

        # Caso seja seccionamento do circuito 1 em torre de circuito duplo
        elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7 or len(
                rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:
            arquivo.write(
                f"new line.T_{torre_sec['torre']}_aux_aberto bus1=T_{torre_sec['torre']}_X.1.2.3 bus2=T_{torre_sec['torre']}.1.2.3 "
                f'phases=3 r1=999999 x1=99999 r0=999999 x0=99999 \n')

            arquivo.write(
                f"new line.T_{torre_sec['torre']}_aux_jumper bus1=T_{torre_sec['torre']}_X.4.5.6 bus2=T_{torre_sec['torre']}.4.5.6 "
                f'phases=3 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

            arquivo.write(
                f"new line.T_{torre_sec['torre']}_7 bus1=T_{torre_sec['torre']}_X.7 bus2=T_{torre_sec['torre']}.7 "
                f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:
                arquivo.write(
                    f"new line.T_{torre_sec['torre']}_8 bus1=T_{torre_sec['torre']}_X.8 bus2=T_{torre_sec['torre']}.8 "
                    f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

    # Circuito seccionado é o 2
    elif torre_sec['circuito'] == 2:
        arquivo.write(
            f"new line.T_{torre_sec['torre']}_aux_aberto bus1=T_{torre_sec['torre']}_X.4.5.6 bus2=T_{torre_sec['torre']}.4.5.6 "
            f'phases=3 r1=999999 x1=99999 r0=999999 x0=99999 \n')

        arquivo.write(
            f"new line.T_{torre_sec['torre']}_aux_jumper bus1=T_{torre_sec['torre']}_X.1.2.3 bus2=T_{torre_sec['torre']}.1.2.3 "
            f'phases=3 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

        arquivo.write(
            f"new line.T_{torre_sec['torre']}_7 bus1=T_{torre_sec['torre']}_X.7 bus2=T_{torre_sec['torre']}.7 "
            f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')

        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 8:
            arquivo.write(
                f"new line.T_{torre_sec['torre']}_8 bus1=T_{torre_sec['torre']}_X.8 bus2=T_{torre_sec['torre']}.8 "
                f'phases=1 r1=0.00001 x1=0.00001 r0=0.00001 x0=0.00001 \n')


# -------------------------------------------------------------------------------------------------- #


def ramal_sec(torres, torre_sec, subestacoes, linecodes, matrix, arquivo):
    """
    Função para inserir ramais de seccionamento
    :param torres: Número de torres já utilizadas no trecho "tronco"
    :param torre_sec: Dicionário contendo número da torre de seccionamento e circuito que foi seccionado
    :param subestacoes: Subestacçoes que restam ser inseridas no modelo (só subestações de seccionamento)
    :param linecodes: Todos linecodes disponíveis
    :param matrix: Todas matrizes de parâmetros lidos dos linecoes
    :param arquivo: Arquivo .dss que será escrito
    :return: Retorna número de torres atualizado
    """
    se_sec = "XXX"
    print(f"\nInserção do ramal de seccionamento que se conecta à torre {torre_sec['torre']} no circuito "
          f"{torre_sec['circuito']} da LT tronco. \n")

    while se_sec not in subestacao:
        se_sec = int(input('\nDigite a SE de seccionamento em questão. Subestações disponíveis: \n'
                           f'{subestacao} \n'))

    codigo_linha = 0
    # Escreve trecho entre SE e sua torre de entrada (definida aqui como portico)
    l_trecho = (10 ** -3) * float(input(f'Digite a distância entre a SE {se_sec} e sua torre de entrada em [m]: '))
    r_at_se = float(input(f'Digite a resistência de aterramento a ser considerada para a SE {se_sec} em [ohm]: '))
    r_at = float(input(f'Digite a resistência de aterramento a ser considerada para a primeira torre em [ohm]: '))

    while codigo_linha not in linecodes:
        print('\nLinecodes disponíveis: ')
        for i in range(len(linecodes)):
            print(f'{linecodes[i]}')
        codigo_linha = input('Digite o linecode do trecho: ')

    portico = 'T_port_SE_' + str(se_sec)

    arquivo.write(f'new line.{str(se_sec) + portico} bus1={se_sec} bus2={portico} '
                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                  f' length= {l_trecho} \n')

    # Escreve pórtico com respectivo aterramento
    if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4:
        node_1 = 4
        arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
        # Escreve aterramento da SE
        arquivo.write(f'new Fault.R_SE{se_sec} bus1={se_sec}.{node_1}.0 r={r_at_se} \n')

    elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
        node_1 = 4
        node_2 = 5
        arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                      f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
        # Escreve aterramento da SE
        arquivo.write(f'new Fault.R_SE{se_sec} bus1={se_sec}.{node_1} bus2={se_sec}.{node_2} r=0 \n'
                      f'new Fault.RSE{se_sec} bus1={se_sec}.{node_2}.0 r={r_at_se} \n')

    elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
        node_1 = 7
        arquivo.write(f'new Fault.RT{portico} bus1={portico}.{node_1}.0 r={r_at} \n')
        # Escreve aterramento da SE
        arquivo.write(f'new Fault.R_SE{se_sec} bus1={se_sec}.{node_1}.0 r={r_at_se} \n')
        # Faz conexão dos equivalentes com demais fases
        arquivo.write(f'new Fault.Eq_{se_sec}_A bus1={se_sec}.1 bus2={se_sec}.4 \n'
                      f'new Fault.Eq_{se_sec}_B bus1={se_sec}.2 bus2={se_sec}.5 \n'
                      f'new Fault.Eq_{se_sec}_C bus1={se_sec}.3 bus2={se_sec}.6 \n')
    else:
        node_1 = 7
        node_2 = 8
        arquivo.write(f'new Fault.R{portico} bus1={portico}.{node_1} bus2={portico}.{node_2} r=0 \n'
                      f'new Fault.RT{portico} bus1={portico}.{node_2}.0 r={r_at} \n')
        # Escreve aterramento da SE
        arquivo.write(f'new Fault.R_SE{se_sec} bus1={se_sec}.{node_1} bus2={se_sec}.{node_2} r=0 \n'
                      f'new Fault.RSE{se_sec} bus1={se_sec}.{node_2}.0 r={r_at_se} \n')
        # Faz conexão dos equivalentes com demais fases
        arquivo.write(f'new Fault.Eq_{se_sec}_A bus1={se_sec}.1 bus2={se_sec}.4 \n'
                      f'new Fault.Eq_{se_sec}_B bus1={se_sec}.2 bus2={se_sec}.5 \n'
                      f'new Fault.Eq_{se_sec}_C bus1={se_sec}.3 bus2={se_sec}.6 \n')

    subestacoes.remove(se_sec)
    torre_de = portico

    # Faz isso porque estava conectando primeira torre de seccionamento ao pórtico da última SE tronco
    torres = torres + 1

    l = float(input(f"\nDigite a distância em [km] entre a torre T_{torre_sec['torre']} e a torre {torre_de}: "))
    l_rest = l
    trecho = 0
    trecho_lt = range(0, 0)

    # Tolerancia de 10 metros
    while l_rest > 0.01:
        l_trecho = float(input(f"\nDigite a distância em [km] entre a torre T_{torre_sec['torre']} e a torre a ser modelada com configuração (linecode) específica. \n"
                                   f"Caso exista apenas uma configuração entre a torre T_{torre_sec['torre']} e torre {torre_de}, será necessário entrar uma única vez com "
                                   f'essa distância e seu valor será igual a {l}. \n'
                                   f'Caso existam múltiplas configurações neste trecho, a soma dos trechos deverá ser igual a {l}. \n'
                                   f'Ainda faltam {l_rest} km \n'
                                   f'Trecho sendo escrito: {trecho + 1} \n'))
        l_rest = l_rest - l_trecho

        if l_rest < 0:
            print(f'\nHá uma inconsistência na entrada de dados, os comprimentos inseridos excedem o trecho total em {abs(l_rest)} km. \n'
                      f'Caso não queira conviver com esse erro, reinicie o programa. (Pois é, nesta altura da execução....)')
            # Vão médio

        vao_m = (10 ** -3) * round(float(input(f'\nDigite o comprimento do vão médio em [m] do trecho {trecho + 1}: ')), 2)
        trecho = trecho + 1

        torres_aux = torres
        # Considera caso em que a torre adjacente à de seccionamento ja é o pórtico
        if int(l_trecho / vao_m) == 1 and trecho == 1:
            torres = int(torres + l_trecho / vao_m - 1)

        else:
            torres = int(torres + l_trecho / vao_m - 1 + 1)

        codigo_linha = 0
        # Escreve trecho

        r_at = float(input(f'\nDigite a resistência de aterramento a ser considerada para as torres em [ohm]: '))

        while codigo_linha not in linecodes:
            print('Linecodes disponíveis: \n')
            for i in range(len(linecodes)):
                print(linecodes[i])

            # Da torre de seccionamento até ponto intermediário
            if trecho == 1 and l_rest != 0:
                codigo_linha = input(f"\nDigite o linecode do trecho compreendido entre a torre T_{torre_sec['torre']} e a torre T_{torres}: \n")
                if codigo_linha in linecodes:
                    # Atualizadas proximas duas linhas em 24/07/2020
                    trecho_lt = range(torres_aux, torres)
                    trecho_at = range(torres_aux, torres + 1)

                    # Conexão com a torre de seccionamento
                    if torre_sec['circuito'] == 2:
                        arquivo.write(f"new line.T_{torre_sec['torre']}_X_{torrex_aux} bus1=T_{torre_sec['torre']}_X.4.5.6  "
                                      f"bus2=T_{torre_aux}.1.2.3 length={vao_m} units=km \n")
                        arquivo.write(f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                                      f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                                      f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                                      f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                                     f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                                      f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n")


                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_{torres_aux} length={vao_m} units=km ")

                        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                            arquivo.write(f"bus1=T_{torre_sec['torre']}.4.5.6.7 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                            arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]}) |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]}) |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]}) |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                        else:
                            arquivo.write(f"bus1=T_{torre_sec['torre']}.4.5.6.7.8 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                            arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")

                    else:
                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.1.2.3 "
                            f"bus2=T_{torres_aux}.1.2.3 length={vao_m} units=km \n")
                        arquivo.write(
                            f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n")


                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_{torres_aux} length={vao_m} units=km ")

                        circuito_sec = int(input('Circuito que foi seccionado é simples ou duplo? \n'
                                                 '1 - Simples (Default) \n'
                                                 '2 - Duplo \n'))

                        if circuito_sec == 1:
                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.4 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                            else:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.4.5 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")

                        else:
                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.7 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                            else:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.7.8 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")

#                    tronco = input('\nLT tronco é circuito simples ou duplo? \n'
#                                   '1 - Simples (Default) \n'
#                                   '2 - Duplo \n')

#                    if tronco == "2":
#                        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")

#                        else:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_8_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.8 linecode={codigo_linha} length = {vao_m} \n")

                    for j in trecho_lt:

                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                                  f' length= {vao_m} \n')

            # Da torre de seccionamento à SE especificada
            elif trecho == 1 and l_rest == 0:
                codigo_linha = input(f"\nDigite o linecode do trecho compreendido entre a torre {torre_sec['torre']} e a torre {torre_de}: \n")
                if codigo_linha in linecodes:
#                    torres_aux = torres - 1
                    trecho_lt = range(torres_aux, torres)
                    trecho_at = range(torres_aux, torres + 1)

                    # Conexão com a torre de seccionamento
                    if torre_sec['circuito'] == 2:
                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.4.5.6 "
                            f"bus2=T_{torres_aux}.1.2.3 length={vao_m} units=km \n")
                        arquivo.write(
                            f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n")

                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_{torres_aux} length={vao_m} units=km ")

                        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                            arquivo.write(f"bus1=T_{torre_sec['torre']}.4.5.6.7 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                            arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                        else:
                            arquivo.write(f"bus1=T_{torre_sec['torre']}.4.5.6.7.8 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                            arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")

                    else:
                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.1.2.3 "
                            f"bus2=T_{torres_aux}.1.2.3 length={vao_m} units=km \n")
                        arquivo.write(
                            f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n"
                            f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][0][0]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][1][1]} |"
                            f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][0]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][1]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][2][2]}) \n")


                        arquivo.write(
                            f"new line.T_{torre_sec['torre']}_{torres_aux} length={vao_m} units=km ")

                        circuito_sec = int(input('Circuito que foi seccionado é simples ou duplo? \n'
                                                 '1 - Simples (Default) \n'
                                                 '2 - Duplo \n'))

                        if circuito_sec == 1:

                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.4 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                            else:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.4.5 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")

                        else:

                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.7 bus2=T_{torres_aux}.4.5.6.7 phases=4 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]}) \n")

                            else:
                                arquivo.write(f"bus1=T_{torre_sec['torre']}.1.2.3.7.8 bus2=T_{torres_aux}.4.5.6.7.8 phases=5 \n")
                                arquivo.write(
                                f"~ rmatrix=({rmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {rmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ xmatrix=({xmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {xmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n"
                                f"~ cmatrix=({cmatrix[linecodes.index(codigo_linha)][codigo_linha][3][3]} |{cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][4][4]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][5][5]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][6][6]} |"
                                f"{cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][3]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][4]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][5]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][6]} {cmatrix[linecodes.index(codigo_linha)][codigo_linha][7][7]})\n")




#                        tronco = input('\nLT tronco é circuito simples ou duplo? \n'
#                                       '1 - Simples (Default) \n'
#                                       '2 - Duplo \n')

#                        if tronco == "2":
#                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")

#                            else:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_8_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux}.8 linecode={codigo_linha} length = {vao_m} \n")

                    for j in trecho_lt:
                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                      f' length= {vao_m} \n')

                    # Conexão com a SE
                    arquivo.write(f'new line.T_{torres}_{torre_de} bus1=T_{torres} bus2={torre_de} '
                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                  f' length= {vao_m} \n')


            # De algum ponto intermediário até outro ponto intermediário
            elif l_rest > 0.01:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre T_{torres_aux} e a torre T_{torres}: \n')
                if codigo_linha in linecodes:
                    trecho_lt = range(torres_aux, torres)
                    trecho_at = range(torres_aux + 1, torres + 1)

                    # Conexão com a torre de seccionamento
#                    if torre_sec['circuito'] == 2:
#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.4.5.6 "
#                            f"bus2=T_{torres_aux}.1.2.3 linecode={codigo_linha} length = {vao_m} \n")

#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_{torres_aux} bus1=T_{torre_sec['torre']}.4.5.6 "
#                            f"bus2=T_{torres_aux}.4.5.6 linecode={codigo_linha} length = {vao_m} \n")

#                        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")

#                        else:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux}.8 linecode={codigo_linha} length = {vao_m} \n")

#                    else:
#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.1.2.3 "
#                            f"bus2=T_{torres_aux}.1.2.3 linecode={codigo_linha} length = {vao_m} \n")

#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_{torres_aux} bus1=T_{torre_sec['torre']}.1.2.3 "
#                            f"bus2=T_{torres_aux}.4.5.6 linecode={codigo_linha} length = {vao_m} \n")

#                        tronco = input('\nLT tronco é circuito simples ou duplo? \n'
#                                       '1 - Simples (Default) \n'
#                                       '2 - Duplo \n')

#                        if tronco == "2":
#                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1={torre_sec['torre']}.7 "
#                                    f"bus2={torres_aux + 1}.7 linecode={codigo_linha} length = {vao_m} \n")

#                            else:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux}.7 linecode={codigo_linha} length = {vao_m} \n")
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_8_{torres_aux} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux}.8 linecode={codigo_linha} length = {vao_m} \n")

                    for j in trecho_lt:
                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                      f' length= {vao_m} \n')

            # De algum ponto intermediário até a SE especificada
            else:
                codigo_linha = input(f'\nDigite o linecode do trecho compreendido entre a torre T_{torres_aux} e a torre {torre_de}: \n')
                if codigo_linha in linecodes:
                    torres = torres - 1
                    trecho_lt = range(torres_aux, torres)
                    trecho_at = range(torres_aux + 1, torres + 1)

#                    # Conexão com a torre de seccionamento
#                    if torre_sec['circuito'] == 2:
#                        arquivo.write(
#                           f"new line.T_{torre_sec['torre']}_X_{torres_aux} bus1=T_{torre_sec['torre']}_X.4.5.6 "
#                            f"bus2=T_{torres_aux}.1.2.3 linecode={codigo_linha} length = {vao_m} \n")

#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_{torres_aux + 1} bus1=T_{torre_sec['torre']}.4.5.6 "
#                            f"bus2=T_{torres_aux + 1}.4.5.6 linecode={codigo_linha} length = {vao_m} \n")

#                        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux + 1}.7 linecode={codigo_linha} length = {vao_m} \n")

#                        else:
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux + 1}.7 linecode={codigo_linha} length = {vao_m} \n")
#                            arquivo.write(
#                                f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                f"bus2=T_{torres_aux + 1}.8 linecode={codigo_linha} length = {vao_m} \n")

#                    else:
#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_X_{torres_aux + 1} bus1=T_{torre_sec['torre']}_X.1.2.3 "
#                            f"bus2=T_{torres_aux + 1}.1.2.3 linecode={codigo_linha} length = {vao_m} \n")

#                        arquivo.write(
#                            f"new line.T_{torre_sec['torre']}_{torres_aux + 1} bus1=T_{torre_sec['torre']}.1.2.3 "
#                            f"bus2=T_{torres_aux + 1}.4.5.6 linecode={codigo_linha} length = {vao_m} \n")

#                        tronco = input('\nLT tronco é circuito simples ou duplo? \n'
#                                       '1 - Simples (Default) \n'
#                                       '2 - Duplo \n')

#                        if tronco == "2":
#                            if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux + 1}.7 linecode={codigo_linha} length = {vao_m} \n")

#                            else:
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_7_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux + 1}.7 linecode={codigo_linha} length = {vao_m} \n")
#                                arquivo.write(
#                                    f"new line.T_{torre_sec['torre']}_8_{torres_aux + 1} bus1=T_{torre_sec['torre']}.7 "
#                                    f"bus2=T_{torres_aux + 1}.8 linecode={codigo_linha} length = {vao_m} \n")

                    for j in trecho_lt:
                        arquivo.write(f'new line.T_{j}_T{j + 1} bus1=T_{j} bus2=T_{j + 1} '
                                      f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                      f' length= {vao_m} \n')

                    # Conexão com a SE
                    arquivo.write(f'new line.T_{torres}_{torre_de} bus1=T_{torres} bus2={torre_de} '
                                  f'phases={len(rmatrix[linecodes.index(codigo_linha)][codigo_linha])} linecode={codigo_linha}'
                                  f' length= {vao_m} \n')


        # Escreve torres com respectivos aterramentos
        if len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 4:
            node_1 = 4
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1}.0 r={r_at} \n')
#            if torre_para[0:4] == 'der_':
#                arquivo.write(f'new Fault.R_{torre_para} bus1=T{torre_para}.{node_1}.0 r={r_at} \n')

        elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 5:
            node_1 = 4
            node_2 = 5
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1} bus2=T_{j}.{node_2} r=0 \n'
                              f'new Fault.RT{j} bus1=T_{j}.{node_2}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1} bus2={torre_de}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_de} bus1=T{torre_de}.{node_2}.0 r={r_at} \n')
#            if torre_para[0:4] == 'der_':
#                arquivo.write(f'new Fault.R_{j} bus1={torre_para}.{node_1} bus2={torre_para}.{node_2} r=0 \n'
#                              f'new Fault.RT{torre_para} bus1={torre_para}.{node_2}.0 r={r_at} \n')

        elif len(rmatrix[linecodes.index(codigo_linha)][codigo_linha]) == 7:
            node_1 = 7
            for j in trecho_at:
                arquivo.write(f'new Fault.R_{j} bus1=T_{j}.{node_1}.0 r={r_at} \n')

            if torre_de[0:4] == 'der_':
                arquivo.write(f'new Fault.R_{torre_de} bus1={torre_de}.{node_1}.0 r={r_at} \n')
#           if torre_para[0:4] == 'der_':
#                arquivo.write(f'new Fault.R_{torre_para} bus1={torre_para}.{node_1}.0 r={r_at} \n')

        else:
            node_1 = 7
            node_2 = 8
            for j in trecho_at:
                arquivo.write(f'new Fault.R{j} bus1=T_{j}.{node_1} bus2=T_{j}.{node_2} r=0 \n'
                              f'new Fault.RT{j} bus1=T_{j}.{node_2}.0 r={r_at} \n')

            if torre_de[0:4] == 'der':
                arquivo.write(f'new Fault.R{torre_de} bus1={torre_de}.{node_1} bus2={torre_de}.{node_2} r=0 \n'
                              f'new Fault.RT{torre_de} bus1={torre_de}.{node_2}.0 r={r_at} \n')

    print(torres)
    return torres

# ------------------------------------------------------------------------------------#

class DSS():

    def __init__(self, arquivo):

        self.arquivo = arquivo

        # Criar a conexão entre Python e OpenDSS
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Iniciar o Objeto DSS
        if self.dssObj.Start(0) == False:
            print('Problemas em iniciar o OpenDSS')

        else:
            # Criar variávels para as principais interfaces
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssLines = self.dssCircuit.Lines
            self.dssVsources = self.dssCircuit.Vsources

    def versao_DSS(self):

        return self.dssObj.Version

    def compila_DSS(self):

        # limpar informações da última simulação
        self.dssObj.ClearAll()

        self.dssText.Command = "compile " + self.arquivo

    def CC_DSS_snapshot(self, nome_barra, tipo_falta):
        """
        Aplica CC no OpenDSS
        :param nome_barra: Nome da barra que terá falta aplicada
        :param tipo_falta: TIpo da falta (3F ou FT)
        :return: Resolve sistema
        """

        # Curto-circuito na barra
        self.dssText.Command = 'Set Mode=Snapshot'

        if tipo_falta == '3F':
            self.dssText.Command = 'new Fault.CC_' + tipo_falta + ' bus1=' + nome_barra + '.1.2.3 phases=3'

        else:
            self.dssText.Command = 'new Fault.CC_' + tipo_falta + ' bus1=' + nome_barra + '.1.0 phases=1'

        # Resolver fluxo de potência (falta)
        self.dssSolution.Solve()

    def estudo_pr_DSS(self, barras_modelo, linhas_modelo):
        """
        Recebe vetores contendo nomes das barras do modelo e nomes das lts do modelo
        Realiza todas as combinações possíveis de curto-circuito entra para-raios e cabo de fase
        e só retorna valores máximos
        :param barras_modelo: Vetor com todos as barras do modelo
        :param linhas_modelo: Vetor com todas as lts do modelo
        :return: Retorna vetor contendo [barra em falta, PR em falta, fase em falta, nome do vão, corrente no pr_1,
        corrente no pr_2 (se aplicável), corrente de falta fluindo pela resistência de aterramento e linecode do trecho
        """
        saida = []

        for barra in barras_modelo:
            # Ativa barra e pega número de nós
            print(barra)
            num_nos = self.ativa_barra(barra)

            if 't_' in barra:
                # Caso com 1 para-raios - circuito simples
                if num_nos == 4:
                    for fase in range(1, 4):
                        # Simula curto-circuito
                        self.dssObj.ClearAll()
                        self.dssText.Command = "compile " + self.arquivo
                        self.dssText.Command = 'Set Mode=Snapshot'
                        self.dssText.Command = 'new Fault.CC_' + barra + ' bus1=' + barra + '.' + str(fase) + ' bus2=' + barra + '.4'
                        self.dssSolution.Solve()
                        # Pega valores de correntes nas linhas do modelo
                        for linha in linhas_modelo:
                            self.ativa_elemento('line.' + linha)
                            if self.fases_linha_opendss() > 3:
                                linecode = self.linecodes_opendss()

                                # 1 Cabo para-raios com circuito simples
                                if self.fases_linha_opendss() == 4:
                                    pr_1 = self.get_resultados_corrente()[6]
                                    pr_2 = 0
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 4, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 2 cabos para-raios com circuito simples
                                elif self.fases_linha_opendss() == 5:
                                    pr_1 = self.get_resultados_corrente()[6]
                                    pr_2 = self.get_resultados_corrente()[8]
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 4, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 1 cabo para-raios com circuito duplo
                                elif self.fases_linha_opendss() == 7:
                                    pr_1 = self.get_resultados_corrente()[12]
                                    pr_2 = 0
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 4, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 2 cabos para-raios com circuito duplo
                                else:
                                    pr_1 = self.get_resultados_corrente()[12]
                                    pr_2 = self.get_resultados_corrente()[14]
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 4, fase, linha, pr_1, pr_2, r_t, linecode])

                # Caso com 2 para-raios - circuito simples
                elif num_nos == 5:
                    for pr in ['4', '5']:
                        for fase in range(1, 4):
                            # Simula curto-circuito
                            self.dssObj.ClearAll()
                            self.dssText.Command = "compile " + self.arquivo
                            self.dssText.Command = 'Set Mode=Snapshot'
                            self.dssText.Command = 'new Fault.CC_' + barra + ' bus1=' + barra + '.' + str(fase) + ' bus2=' + barra + '.' + pr
                            self.dssSolution.Solve()
                            # Pega valores de correntes nas linhas do modelo
                            for linha in linhas_modelo:
                                self.ativa_elemento('line.' + linha)
                                if self.fases_linha_opendss() > 3:
                                    linecode = self.linecodes_opendss()

                                    # 1 Cabo para-raios com circuito simples
                                    if self.fases_linha_opendss() == 4:
                                        pr_1 = self.get_resultados_corrente()[6]
                                        pr_2 = 0
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 2 cabos para-raios com circuito simples
                                    elif self.fases_linha_opendss() == 5:
                                        pr_1 = self.get_resultados_corrente()[6]
                                        pr_2 = self.get_resultados_corrente()[8]
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 1 cabo para-raios com circuito duplo
                                    elif self.fases_linha_opendss() == 7:
                                        pr_1 = self.get_resultados_corrente()[12]
                                        pr_2 = 0
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 2 cabos para-raios com circuito duplo
                                    else:
                                        pr_1 = self.get_resultados_corrente()[12]
                                        pr_2 = self.get_resultados_corrente()[14]
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                # Caso com 1 para-raios - circuito duplo
                elif num_nos == 7:
                    for fase in range(1, 7):
                        # Simula curto-circuito
                        self.dssObj.ClearAll()
                        self.dssText.Command = "compile " + self.arquivo
                        self.dssText.Command = 'Set Mode=Snapshot'
                        self.dssText.Command = 'new Fault.CC_' + barra + ' bus1=' + barra + '.' + str(
                                            fase) + ' bus2=' + barra + '.7'
                        self.dssSolution.Solve()
                        # Pega valores de correntes nas linhas do modelo
                        for linha in linhas_modelo:
                            self.ativa_elemento('line.' + linha)
                            if self.fases_linha_opendss() > 3:
                                linecode = self.linecodes_opendss()

                                # 1 Cabo para-raios com circuito simples
                                if self.fases_linha_opendss() == 4:
                                    pr_1 = self.get_resultados_corrente()[6]
                                    pr_2 = 0
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 7, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 2 cabos para-raios com circuito simples
                                elif self.fases_linha_opendss() == 5:
                                    pr_1 = self.get_resultados_corrente()[6]
                                    pr_2 = self.get_resultados_corrente()[8]
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 7, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 1 cabo para-raios com circuito duplo
                                elif self.fases_linha_opendss() == 7:
                                    pr_1 = self.get_resultados_corrente()[12]
                                    pr_2 = 0
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 7, fase, linha, pr_1, pr_2, r_t, linecode])

                                # 2 cabos para-raios com circuito duplo
                                else:
                                    pr_1 = self.get_resultados_corrente()[12]
                                    pr_2 = self.get_resultados_corrente()[14]
                                    self.ativa_elemento('Fault.RT' + barra)
                                    r_t = self.get_resultados_corrente()[0]
                                    saida.append([barra, 7, fase, linha, pr_1, pr_2, r_t, linecode])

                # Caso com 2 para-raios - circuito duplo
                elif num_nos == 8:
                    for pr in ['7', '8']:
                        for fase in range(1, 7):
                            # Simula curto-circuito
                            self.dssObj.ClearAll()
                            self.dssText.Command = "compile " + self.arquivo
                            self.dssText.Command = 'Set Mode=Snapshot'
                            self.dssText.Command = 'new Fault.CC_' + barra + ' bus1=' + barra + '.' + str(
                                     fase) + ' bus2=' + barra + '.' + pr
                            self.dssSolution.Solve()
                            # Pega valores de correntes nas linhas do modelo
                            for linha in linhas_modelo:
                                self.ativa_elemento('line.' + linha)
                                if self.fases_linha_opendss() > 3:
                                    linecode = self.linecodes_opendss()

                                    # 1 Cabo para-raios com circuito simples
                                    if self.fases_linha_opendss() == 4:
                                        pr_1 = self.get_resultados_corrente()[6]
                                        pr_2 = 0
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 2 cabos para-raios com circuito simples
                                    elif self.fases_linha_opendss() == 5:
                                        pr_1 = self.get_resultados_corrente()[6]
                                        pr_2 = self.get_resultados_corrente()[8]
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 1 cabo para-raios com circuito duplo
                                    elif self.fases_linha_opendss() == 7:
                                        pr_1 = self.get_resultados_corrente()[12]
                                        pr_2 = 0
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

                                    # 2 cabos para-raios com circuito duplo
                                    else:
                                        pr_1 = self.get_resultados_corrente()[12]
                                        pr_2 = self.get_resultados_corrente()[14]
                                        self.ativa_elemento('Fault.RT' + barra)
                                        r_t = self.get_resultados_corrente()[0]
                                        saida.append([barra, pr, fase, linha, pr_1, pr_2, r_t, linecode])

        return saida

    def get_resultados_corrente(self):

        # Pega vetor de correntes
        #self.dssText.Command = 'Show Current Elements'
        return self.dssCktElement.CurrentsMagAng

    def barras_opendss(self):

        # Pega todas as barras do modelo
        return self.dssCircuit.AllBusNames

    def linhas_opendss(self):

        # Pega todas as linhas do modelo
        return self.dssLines.AllNames

    def fases_linha_opendss(self):

        # Pega a fase da LT ativa
        return self.dssLines.Phases

    def linecodes_opendss(self):

        # Pega todos os linecodes do modelo
        return self.dssLines.LineCode


    def ativa_barra(self, nome_barra):

        # Ativa a barra pelo nome e retorna o número de nós da mesma
        self.dssCircuit.SetActivebus(nome_barra)
        return self.dssBus.NumNodes

    def ativa_elemento(self, nome_elemento):

        # Ativa elemento pelo nome completo: Tipo.Nome
        self.dssCircuit.SetActiveElement(nome_elemento)
        return self.dssCktElement.Name

    def vsource_opendss(self):

        # Recebe todas fontes de tensão do modelo
        return self.dssVsources.AllNames

# ------------------------------------------------------------------------------------#


print('Ferramenta para elaboração de estudo de superação/dimensionamento de cabos para raios. \n'
      'O estudo é dividido em três etapas: \n'
      '1 - Leitura de arquivo .ANA contendo equivalente e arquivos .lis contendo matrizes de impedâncias e capacitâncias \n'
      'da LT em estudo. Em seguida é montado o arquivo em formato .dss para execução do estudo. \n'
      '2 - Leitura do arquivo .dss e validação no modelo. Essa validação é feita por meio de curtos-circuitos aplicados \n'
      'em pontos específicos do modelo e monitoramento dos níveis de falta e contribuições das subestações. \n'
      '3 - Execução do estudo de curto-circuito e distribuição de correntes nos cabos para-raios. Exportação dos resultados \n'
      'em planilha do Excel, contendo apenas as máximas correntes nos para-raios de cada vão modelado. \n'
      '\nO estudo pode ser executado de forma sequencial ou não, ou seja, caso o usuário já tenha o arquivo .dss, é possível \n'
      'ir direto para a etapa 2 ou etapa 3.\n')

entrada_usuario = "X"

while entrada_usuario != "1" and entrada_usuario != "2":
    entrada_usuario = input('Digite o que deseja fazer: \n'
                            '1 - Iniciar na Etapa 1 (Montagem do arquivo .dss com base em arquivos .ANA e .lis. \n'
                            '2 - Pular etapa 1 e ir direto para a etapa 2. \n'
                            'Caso queira ir direto para a etapa 3, digite "2" e em seguida escolha a opção de iniciar estudo \n'
                            'de superação de cabos para-raios. \n')

diretorio = 0

if entrada_usuario == "1":

    #diretorio = 0

    # Conjunto de barras no sistema

    barras_num = []
    barras_nome = []
    barras_kv = []
    # Tipo de representação das barras, se for derivação, o ANAFAS indica como 2
    barras_tipo = []
    # Conjunto de circuitos no sistema
    circuitos = []

    # Dados de LTs
    lt_de = []
    lt_para = []
    # Dados de equivalentes de transferência
    lt_eq_de = []
    lt_eq_para = []
    r1_lt_eq = []
    x1_lt_eq = []
    r0_lt_eq = []
    x0_lt_eq = []
    # Dados de equivalenets de curto-circuito
    barra_eq_cc = []
    r1_eq_cc = []
    x1_eq_cc = []
    r0_eq_cc = []
    x0_eq_cc = []
    # Dados de shunts equivalentes
    barra_sh_eq = []
    r1_sh_eq = []
    x1_sh_eq = []
    r0_sh_eq = []
    x0_sh_eq = []
    # Dados de mútuas
    de_mutua_de = []
    para_mutua_de = []
    de_mutua_para = []
    para_mutua_para = []
    i_mutua_de = []
    f_mutua_de = []
    i_mutua_para = []
    f_mutua_para = []

    while diretorio == 0:
        try:
            leitura = input('Digite o diretório do arquivo com seu nome e sua extensão \n'
                            'por exemplo: C:/Users/Matheus/Desktop/meuarquivo.ana: \n')
            arquivo = open(leitura, 'r', encoding='utf8')
        except:
            print('Há algo errado com o diretorio ou arquivo informados, verifique.')
        else:
            arquivo.close
            diretorio = 1

    # Interpreta o arquivo e descreve o modelo

    with open(leitura, 'r', encoding='ISO-8859-2') as arquivo:
        texto = arquivo.readlines()
        for linha in texto:

            # Leitura dos dados de barras
            if 'DBAR' in linha:
                i = 1
                while texto[texto.index(linha) + i][0:5] != '99999':

                    if texto[texto.index(linha) + i][0] == '(':
                        i = i + 1

                    else:
                        barras_num.append(int(texto[texto.index(linha) + i][0:5]))
                        barras_nome.append(texto[texto.index(linha) + i][9:21])
                        barras_kv.append(int(texto[texto.index(linha) + i][31:35]))
                        barras_tipo.append(texto[texto.index(linha) + i][7])

                        i = i + 1

            # Leitura dos dados de circuito
            if 'DCIR' in linha:
                i = 1
                while texto[texto.index(linha) + i][0:5] != '99999':

                    if texto[texto.index(linha) + i][0] == '(':
                        i = i + 1

                    # Identificação de equivalentes
                    elif texto[texto.index(linha) + i][41:47] == 'EQUIV.':
                        # LT Equivalente
                        if int(texto[texto.index(linha) + i][7:12]) != 0:
                            lt_eq_de.append(int(texto[texto.index(linha) + i][0:5]))
                            lt_eq_para.append(int(texto[texto.index(linha) + i][7:12]))
                            if "." in texto[texto.index(linha) + i][17:23]:
                                r1_lt_eq.append(float(texto[texto.index(linha) + i][17:23]))
                            else:
                                r1_lt_eq.append(float(texto[texto.index(linha) + i][17:23]) / 100)

                            if "." in texto[texto.index(linha) + i][23:29]:
                                x1_lt_eq.append(float(texto[texto.index(linha) + i][23:29]))
                            else:
                                x1_lt_eq.append(float(texto[texto.index(linha) + i][23:29]) / 100)

                            if "." in texto[texto.index(linha) + i][29:35]:
                                r0_lt_eq.append(float(texto[texto.index(linha) + i][29:35]))
                            else:
                                r0_lt_eq.append(float(texto[texto.index(linha) + i][29:35]) / 100)

                            if "." in texto[texto.index(linha) + i][35:41]:
                                x0_lt_eq.append(float(texto[texto.index(linha) + i][35:41]))
                            else:
                                x0_lt_eq.append(float(texto[texto.index(linha) + i][35:41]) / 100)

                            i = i + 1

                        # Shunt Equivalente
                        elif int(texto[texto.index(linha) + i][7:12]) == 0:
                            if texto[texto.index(linha) + i][17:23] == '999999' and texto[texto.index(linha) + i][23:29] == '999999':
                                barra_sh_eq.append(int(texto[texto.index(linha) + i][0:5]))
                                r1_sh_eq.append(int(texto[texto.index(linha) + i][17:23]))
                                x1_sh_eq.append(int(texto[texto.index(linha) + i][23:29]))

                                if "." in texto[texto.index(linha) + i][29:35]:
                                    r0_sh_eq.append(float(texto[texto.index(linha) + i][29:35]))

                                else:
                                    r0_sh_eq.append(float(texto[texto.index(linha) + i][29:35]) / 100)

                                if "." in texto[texto.index(linha) + i][35:41]:
                                    x0_sh_eq.append(float(texto[texto.index(linha) + i][35:41]))

                                else:
                                    x0_sh_eq.append(float(texto[texto.index(linha) + i][35:41]) / 100)

                                i = i + 1

                            # Gerador Equivalente
                            else:
                                barra_eq_cc.append(int(texto[texto.index(linha) + i][0:5]))

                                if "." in texto[texto.index(linha) + i][17:23]:
                                    r1_eq_cc.append(float(texto[texto.index(linha) + i][17:23]))

                                else:
                                    r1_eq_cc.append(float(texto[texto.index(linha) + i][17:23]) / 100)

                                if "." in texto[texto.index(linha) + i][23:29]:
                                    x1_eq_cc.append(float(texto[texto.index(linha) + i][23:29]))

                                else:
                                    x1_eq_cc.append(float(texto[texto.index(linha) + i][23:29]) / 100)

                                if "." in texto[texto.index(linha) + i][29:35]:
                                    r0_eq_cc.append(float(texto[texto.index(linha) + i][29:35]))

                                else:
                                    r0_eq_cc.append(float(texto[texto.index(linha) + i][29:35]) / 100)

                                if "." in texto[texto.index(linha) + i][35:41]:
                                    x0_eq_cc.append(float(texto[texto.index(linha) + i][35:41]))

                                else:
                                    x0_eq_cc.append(float(texto[texto.index(linha) + i][35:41]) / 100)

                                i = i + 1

                    # Identificação da LT
                    else:
                        lt_de.append(int(texto[texto.index(linha) + i][0:5]))
                        lt_para.append(int(texto[texto.index(linha) + i][7:12]))
                        i = i + 1

            # Leitura de dados de acoplamentos de sequência zero
            if 'DMUT' in linha:
                i = 1
                while texto[texto.index(linha) + i][0:5] != '99999':

                    if texto[texto.index(linha) + i][0] == '(':
                        i = i + 1

                    # Identificação de trechos
                    else:
                        # Salvando trechos como string
                        de_mutua_de.append(texto[texto.index(linha) + i][0:5])
                        para_mutua_de.append(texto[texto.index(linha) + i][7:12])
                        de_mutua_para.append(texto[texto.index(linha) + i][16:21])
                        para_mutua_para.append(texto[texto.index(linha) + i][23:28])

                        # Salvando porcentagens
                        if texto[texto.index(linha) + i][45:51] == '      ':
                            i_mutua_de.append(0)
                        else:
                            i_mutua_de.append(float(texto[texto.index(linha) + i][45:51]))

                        if texto[texto.index(linha) + i][51:57] == '      ':
                            f_mutua_de.append(100)
                        else:
                            f_mutua_de.append(float(texto[texto.index(linha) + i][51:57]))

                        if texto[texto.index(linha) + i][57:63] == '      ':
                            i_mutua_para.append(0)
                        else:
                            i_mutua_para.append(float(texto[texto.index(linha) + i][57:63]))

                        if texto[texto.index(linha) + i][63:69] == '      ':
                            f_mutua_para.append(100)
                        else:
                            f_mutua_para.append(float(texto[texto.index(linha) + i][63:69]))

                        i = i + 1

    leitura = '1'

    leitura = input('Gostaria de imprimir os dados lidos do arquivo .ANA ? \n'
                    'Qualquer dígito = Sim \n'
                    '2 = Não \n')

    if leitura != '2':

        # Dados de Barras
        print(f'O sistema lido possui {len(barras_num)} barras com os seguintes dados: \n')
        for i in range(len(barras_num)):
            print(f'Barra {i+1}: \n'
                  f'Número: {barras_num[i]} \n'
                f'Nome: {barras_nome[i]} \n'
                f'Tensão de base {barras_kv[i]} kV')

            if barras_tipo[i] == '2':
                print('Barra de derivação (Tipo 2) no ANAFAS.')
                input('Aperte "Enter" para continuar. \n')

            else:
                print('Barra é uma SE.')
                input('Aperte "Enter" para continuar. \n')

        # LT em estudo
        print('LT em estudo:')
        print(f'LT descrita no modelo composta pelas conexões: ')
        for i in range(len(lt_de)):
            print(f'DE: {lt_de[i]} PARA: {lt_para[i]}')
        input('Aperte "Enter" para continuar. \n')

        # Equivalentes de transferência
        print('Equivalentes de Transferência')
        for i in range(len(lt_eq_de)):
            print(f'DE: {lt_eq_de[i]} PARA: {lt_eq_para[i]}')
            print(f'Impedância em % na base 100 MVA e {barras_kv[0]} kV: \n'
                  f'z0 = {r0_lt_eq[i]} + j * {x0_lt_eq[i]} \n'
                f'z1 = {r1_lt_eq[i]} + j * {x1_lt_eq[i]}')
            input('Aperte "Enter" para continuar. \n')

        # Equivalentes shunt
        print('Equivalentes Shunt (reatores)')
        for i in range(len(barra_sh_eq)):
            print(f'Barra: {barra_sh_eq[i]}')
            print(f'Impedância em % na base 100 MVA e {barras_kv[0]} kV: \n'
                 f'z0 = {r0_sh_eq[i]} + j * {x0_sh_eq[i]} \n'
                 f'z1 = {r1_sh_eq[i]} + j * {x1_sh_eq[i]}')
            input('Aperte "Enter" para continuar. \n')

        # Equivalentes de curto-circuito
        print('Equivalentes de curto-circuito (fontes)')
        for i in range(len(barra_eq_cc)):
            print(f'Barra: {barra_eq_cc[i]}')
            print(f'Impedância em % na base 100 MVA e {barras_kv[0]} kV: \n'
                  f'z0 = {r0_eq_cc[i]} + j * {x0_eq_cc[i]} \n'
                  f'z1 = {r1_eq_cc[i]} + j * {x1_eq_cc[i]}')
            input('Aperte "Enter" para continuar. \n ')

        # Acoplamentos mútuos
        print('Mútuas de sequência zero')
        for i in range(len(de_mutua_de)):
            print(f'LT DE   : {de_mutua_de[i]} - {para_mutua_de[i]} no trecho de {i_mutua_de[i]} % a {f_mutua_de[i]} % \n'
                  f'LT PARA : {de_mutua_para[i]} - {para_mutua_para[i]} no trecho de {i_mutua_para[i]} % a {f_mutua_para[i]} % \n')
            input('Aperte "Enter" para continuar. \n')

    print('\nLeitura dos arquivos .lis contendo configurações das torres. \n'
          'No OpenDSS chamamamos de linecode. No ATP é a saída do Line Constants.\n'
          'Importante que o arquivo contenha as matrizes Zeq e Yeq sem eliminação dos cabos para-raios. \n'
          'Importante que as matrizes .lis estejam em [ohm/km] e [mho/km] respectivamente. \n'
          'O programa converte conforme entrada no OpenDSS que é em [ohm/km] para R e X e [nF/km] para capacitância.\n')


    rmatrix = []
    xmatrix = []
    cmatrix = []
    linecode = []

    fim = '1'

    # Permissão para leitura sequencial de arquivos .lis (diversos line constants)


    while fim != '0':
        if len(xmatrix) == 0:
            fim = input('Início da leitura dos .lis do atp para montagem dos linecodes. \n'
                        'Pressione qualquer tecla para iniciar ou 0 para finalizar: \n')

        else:
            fim = input(f'Leitura do {len(xmatrix) + 1}º arquivo .lis do atp para montagem dos linecodes. \n'
                        f'Pressione qualquer tecla para iniciar ou 0 para finalizar: \n')

        if fim == '0':
            break

        # Número de cabos equivalentes
        n_ca = 0
        # Número de fases (3 ou 6)
        n_fa = 0
        # Número de cabos para-raios (1 ou 2)
        n_pr = 0
        diretorio = 0

        while diretorio == 0:
            try:
                leitura = input('Digite o diretório do arquivo com seu nome e sua extensão \n'
                                'por exemplo: C:/Users/Matheus/Desktop/meuarquivo.lis: \n')
                arquivo = open(leitura, 'r', encoding='utf8')
            except:
                print('Há algo errado com o diretorio ou arquivo informados, verifique.')
            else:
                arquivo.close
                diretorio = 1

        # Verifica número de cabos de fase e cabos para-raios

        with open(leitura, 'r', encoding='utf8') as arquivo:
            texto = arquivo.readlines()
            for linha in texto:

                if 'Line conductor table after sorting and initial processing.' in linha:
                    j = 0
                    while texto[texto.index(linha) + 3 + j][0:22] != 'Matrices are for earth':
                        # Regra para guardar número de fases e cabos para-raios
                        try:
                            if int(texto[texto.index(linha) + 3 + j][7:14]) > n_ca:
                                # Guarda número de cabos equivalentes
                                n_ca = int(texto[texto.index(linha) + 3 + j][7:14])
                                j = j + 1
                            else:
                                j = j + 1
                        except ValueError:
                            j = j + 1
                    # Caso com 1 cabo para-raio (muito raro)
                    if n_ca == 4 or n_ca == 7:
                        n_pr = 1
                        n_fa = n_ca - n_pr
                    # Caso com 2 cabos para-raios (comum)
                    else:
                        n_pr = 2
                        n_fa = n_ca - n_pr

            print(f'No arquivo indicado, há {n_fa} condutores reduzidos (fases) e {n_pr} para-raios.')
            r = defaultdict(
                lambda: [[[numero for numero in range(1, 1)] for linha in range(0, n_ca)] for coluna in range(0, n_ca)])
            x = defaultdict(
                lambda: [[[numero for numero in range(1, 1)] for linha in range(0, n_ca)] for coluna in range(0, n_ca)])
            c = defaultdict(
                lambda: [[[numero for numero in range(1, 1)] for linha in range(0, n_ca)] for coluna in range(0, n_ca)])

            for linha in texto:

                # Guarda matriz de capacitâncias
                if 'Susceptance matrix,  in units of  [mhos/kmeter ]  for the system of equivalent phase conductors.' in linha:
                    nome_linecode = input('Digite o nome que deseja dar à essa configuração (Não repetir!): ')
                    linecode.append(nome_linecode)
                    for i in range(n_ca):
                        for k in range(i + 1):
                            c[nome_linecode][i][k] = float(
                                texto[texto.index(linha) + 3 + i * 2][(5 + 14 * k):(18 + 14 * k)]) * 1 / (2 * math.pi * 60) * 10 ** 9

                # Guarda matriz de resistências e reatâncias
                if 'Impedance matrix,  in units of Ohms/kmeter  for the system of equivalent phase conductors.' in linha:
                    for i in range(n_ca):
                        for k in range(i + 1):
                            r[nome_linecode][i][k] = float(texto[texto.index(linha) + 3 + i*3][(6 + 14*k):(18 + 14*k)])
                            x[nome_linecode][i][k] = float(texto[texto.index(linha) + 3 + i*3 + 1][(6 + 14*k):(18 + 14*k)])

            rmatrix.append(r)
            xmatrix.append(x)
            cmatrix.append(c)

            print('Leitura do .lis efetuada. \n')

    # Apresentação dos linecodes lidos:

    leitura = '1'

    leitura = input('Gostaria de imprimir os dados lidos do(s) arquivo(s) .lis ? \n'
                    'Qualquer dígito = Sim \n'
                    '2 = Não \n')

    if leitura != '2':

        for i in range(len(linecode)):
            print(f'Identificação do linecode: {linecode[i]} \n'
                  f'Matriz de Resistências em ohm/km: {rmatrix[i][linecode[i]]} \n'
                f'Matriz de Reatâncias em ohm/km: {xmatrix[i][linecode[i]]} \n'
                f'Matriz de Capacitâncias em nF/km: {cmatrix[i][linecode[i]]} \n')

    # Montagem do modelo no OpenDSS

    print('Montagem do modelo no OpenDSS. \n')

    subestacao = []
    derivacao = []

    for i in range(len(barras_nome)):
        if barras_tipo[i] == '2':
            derivacao.append(barras_num[i])
        else:
            subestacao.append(barras_num[i])

    # Definição do nome do arquivo .dss a ser gerado

    lt_estudo = input('Digite o nome da LT em estudo: ')

    nome_arq = input('Digite o nome do arquivo .dss a ser gerado para o modelo em estudo \n'
                     'O OpenDSS não reconhece espaços em branco, assim, utilize o caractere "_" para isso \n'
                     'Exemplo: "LT 230 kV APA - SCA"\n'
                     ':Utilize: "LT_230_kV_APA-SCA"\n')

    # Testa cabeçalho

    with open(nome_arq + '.dss', 'w') as arquivo:
        arquivo.write(f'// Modelo para estudo de superacao de cabos para-raios da {lt_estudo} // \n'
                      f'\n'
                      f'// Equivalentes de Curto-Circuito // \n'
                      f'\n'
                    f'Clear \n'
                    f'\n')

        # Escrevendo equivalentes de curto-circuito

        for i in range(len(barra_eq_cc)):

            if i == 0:
                arquivo.write(f'New Circuit.{barra_eq_cc[i]} bus1={barra_eq_cc[i]} basekv={barras_kv[0]} phases=3 '
                              f'pu=1.0 '
                              f'R0={round(r0_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                              f'X0={round(x0_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                            f'R1={round(r1_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                            f'X1={round(x1_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} \n')
            else:
                arquivo.write(f'New Vsource.{barra_eq_cc[i]} bus1={barra_eq_cc[i]} basekv={barras_kv[0]} phases=3 '
                              f'pu=1.0 '
                              f'R0={round(r0_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                              f'X0={round(x0_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                            f'R1={round(r1_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                            f'X1={round(x1_eq_cc[i] / 100 * barras_kv[0] ** 2 / 100, 4)} \n')

            arquivo.write('\n')

        arquivo.write("// Equivalentes de Transferência // \n")
        for i in range(len(lt_eq_de)):
            arquivo.write(f'new line.eq_tr{i+1} bus1={lt_eq_de[i]} bus2={lt_eq_para[i]} phases=3 '
                          f'r0={round(r0_lt_eq[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                          f'x0={round(x0_lt_eq[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                          f'r1={round(r1_lt_eq[i] / 100 * barras_kv[0] ** 2 / 100, 4)} '
                          f'x1={round(x1_lt_eq[i] / 100 * barras_kv[0] ** 2 / 100, 4)} \n')

        arquivo.write('\n \n')

        arquivo.write("// Reator Equivalente // \n")
        for i in range(len(barra_sh_eq)):
            arquivo.write(f'new Reactor.eq_sh{i+1} bus1={barra_sh_eq[i]} phases=3 '
                          f'Z1=[9999, 9999] '
                          f'Z0=[{round(r0_sh_eq[i] / 100 * barras_kv[0] ** 2 / 100 , 4)}, '
                          f'{round(x0_sh_eq[i] / 100 * barras_kv[0] ** 2 / 100 ,4)}] \n')

        arquivo.write('\n \n')

        # Escrevendo linecodes
        arquivo.write("// LineCode - Modelo do vão // \n")
        for i in range(len(linecode)):
            rmatrix_str = []
            xmatrix_str = []
            cmatrix_str = []

            rmatrix_str = str(rmatrix[i][linecode[i]])
            rmatrix_str = rmatrix_str.replace('[]], [', '|')
            for char in '[] ':
                rmatrix_str = rmatrix_str.replace(char, '')
            rmatrix_str = rmatrix_str.replace(',', ' ')

            xmatrix_str = str(xmatrix[i][linecode[i]])
            xmatrix_str = xmatrix_str.replace('[]], [', '|')
            for char in '[] ':
                xmatrix_str = xmatrix_str.replace(char, '')
            xmatrix_str = xmatrix_str.replace(',', ' ')

            cmatrix_str = str(cmatrix[i][linecode[i]])
            cmatrix_str = cmatrix_str.replace('[]], [', '|')
            for char in '[] ':
                cmatrix_str = cmatrix_str.replace(char, '')
            cmatrix_str = cmatrix_str.replace(',', ' ')

            arquivo.write(f'\nNew linecode.{linecode[i]} nphases={len(rmatrix[i][linecode[i]])} basefreq=60 units=km \n'
                          f'~ rmatrix=({rmatrix_str}) \n'
                          f'~ xmatrix=({xmatrix_str}) \n'
                          f'~ cmatrix=({cmatrix_str}) \n \n')

        # Escrevendo trechos de LT

    #    porticos = []
    #    torre = 1
        arquivo.write(f"\n// Trechos de LT e Torres //\n")

    #    # Decisão sobre representação das torres de derivação
    #    decide_der = input('\nDeseja representar as torres de derivação no modelo? \n'
    #                       '1 - Sim (Default) \n'
    #                       '2 - Não \n')

    #    if decide_der == "2":
    #        derivacao = []

        # Vetor torres contendo em sua posição 0 o número de torres já modeladas e nas demais posições
        #  receberá os números das torres de seccionamento que serão utilizadas na segunda etapa (função seccionamento)

        torres = [0]
        de = 0
        para = 0
        ses = subestacao.copy()

        print('\nPrimeira parte do modelo: \n'
              'Definir uma LT tida como "tronco" e respectivas torres de seccionamento, porém,'
              'os trechos de linha originários de um seccionamento são definidos na segunda etapa. \n')

        while subestacao != []:

            while de not in subestacao:
                de = int(input('\nDigite a SE de origem. Subestações disponíveis: \n'
                               f'{subestacao} \n'))

            while para not in subestacao or para == de:
                para = int(input('Digite a SE de destino: \n'))

            # Lista torres recebe valores atualizados de número de torres utilizadas e torres de seccionamento
            torres = caminho(de, para, linecode, xmatrix, subestacao, derivacao, torres, arquivo)

            if len(torres) != 1:
                print(f'\nNúmero de torres e torres de Seccionamento definidas na LT tronco: \n'
                      f'{torres}\n')

                print('\nSegunda parte do modelo: \n'
                      'Inserção de SEs que são conectadas por meio de ramal de seccionamento. \n')

                arquivo.write('\n// Ramais de Seccionamento //\n')

                for i in range(1, len(torres)):
                    torres[0] = ramal_sec(torres[0], torres[i], subestacao, linecode, xmatrix, arquivo)
#                    torres = ramal_sec(torres[0], torres[i], subestacao, linecode, xmatrix, arquivo)

            else:
                print(f'\nNúmero de torres definidas na LT tronco: {torres}. \n')

            arquivo.write(f'\nSet voltagebases=[{barras_kv[0]}] \n'
                          f'calcvoltagebases \n'
                          f'calcv \n')

    # Modelo finalizado, entramos agora no estudo em si

    # coding ISO-8859-2

#if __name__ == "__main__":


#    print(type(arquivo))
    #arquivo = "C:/Users/mgribeiro/PycharmProjects/Sispot/cabos_pr/2_barras_lt_simples.dss"

    # Criar objeto da classe DSS
    #objeto = DSS(nome_arq)
    #print(nome_arq)
    #print(type(nome_arq))

elif entrada_usuario == "2":

    # Usuário escolheu ir para leitura direta de um arquivo .dss
#    while diretorio == 0:
#        try:
#            nome_arq = input('Digite o diretório do arquivo com seu nome e SEM sua extensão .dss \n'
#                            'por exemplo: C:/Users/Matheus/Desktop/meuarquivo: \n')
#            arquivo = open(leitura, 'r', encoding='ISO-8859-2')
#        except:
#            print('Há algo errado com o diretorio ou arquivo informados, verifique.')
#        else:
#            arquivo.close
#            diretorio = 1

    nome_arq = input('Digite o diretório do arquivo com seu nome. \n'
                     'por exemplo: C:/Users/Matheus/Desktop/meuarquivo.dss: \n')

    ses = []


objeto = DSS(nome_arq)
objeto.compila_DSS()

barras_modelo = objeto.barras_opendss()
#print(barras_modelo)

num_nos = []
for barra in barras_modelo:
    num_nos.append(objeto.ativa_barra(barra))

#print(num_nos)

linhas_modelo = objeto.linhas_opendss()
print(linhas_modelo)

# Caso o usuário tenha entrado com Opção de ir direto para o estudo
if ses == []:
    for barra in barras_modelo:
        if 't_port_se_' in barra:
            ses.append(barra.replace('t_port_se_', ''))

#print(ses)
#ses = ['1', '2']
faltas = ['3F', 'FT']
equivalentes = objeto.vsource_opendss()

porticos = []
for portico in ses:
    porticos.append('t_port_se_' + str(portico))

valida_cc = 99

while valida_cc != "0":
    valida_cc = input('\nValidação do modelo no OpenDSS aplicando faltas no modelo. Escolha a opção: \n'
                      '\n1 - Faltas 3F e FT nas SEs imprimindo contribuições das SEs'
                      '\n2 - Faltas 3F e FT nas primeiras torres (T_port) imprimindo contribuições das SEs'
                      '\n3 - Faltas 3F e FT em torre selecionada imprimindo contribuições das SEs'
                      '\n4 - Prosseguir para o estudo dos cabos para-raios'
                      '\nQualquer outra tecla - Encerrar programa \n')

    if valida_cc == "1":
        for i in ses:
            for j in faltas:
                objeto.compila_DSS()
                objeto.CC_DSS_snapshot(str(i), j)
                objeto.ativa_elemento('Fault.CC_' + j)
                correntes = objeto.get_resultados_corrente()

                print(f'\nCorrentes de curto-circuito {j} na SE {i}: \n'
                      f'Correntes de falta: '
                      f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                if j == '3F':
                    print(
                      f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                      f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

                for k in equivalentes:
                    objeto.ativa_elemento('Vsource.' + k)
                    correntes = objeto.get_resultados_corrente()
                    print(f'Contribuição da SE {k}: '
                          f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                    if j == '3F':
                        print(
                          f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                          f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

    elif valida_cc == "2":
        for i in porticos:
            for j in faltas:
                objeto.compila_DSS()
                objeto.CC_DSS_snapshot(i, j)
                objeto.ativa_elemento('Fault.CC_' + j)
                correntes = objeto.get_resultados_corrente()

                print(f'\nCorrentes de curto-circuito {j} na torre (pórtico) {i}: \n'
                      f'Correntes de falta: '
                      f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                if j == '3F':
                    print(
                      f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                      f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

                for k in equivalentes:
                    objeto.ativa_elemento('Vsource.' + k)
                    correntes = objeto.get_resultados_corrente()
                    print(f'Contribuição da SE {k}: '
                          f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                    if j == '3F':
                        print(
                          f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                          f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

    elif valida_cc == "3":
        # Exclui SEs e pórticos da contagem
        num_torres = 0
        for torre in barras_modelo:
            if 'port_se' not in torre and 't_' in torre:
                num_torres = num_torres + 1
        try:
            falta_torre = int(input(f'\nHá {num_torres} torres disponíveis. Digite o número da torre em que se deseja aplicar falta. \n'
                                    f'Primeira torre = 1 e última torre = {num_torres}: \n'))
        except ValueError:
            print('\nEntre com um número dentro do intervalo apresentado acima.')
        else:
            for j in faltas:
                objeto.compila_DSS()
                objeto.CC_DSS_snapshot('T_' + str(falta_torre), j)
                objeto.ativa_elemento('Fault.CC_' + j)
                correntes = objeto.get_resultados_corrente()

                print(f'\nCorrentes de curto-circuito {j} na torre T_{falta_torre}: \n'
                      f'Correntes de falta: '
                      f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                if j == '3F':
                    print(
                        f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                        f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

                for k in equivalentes:
                    objeto.ativa_elemento('Vsource.' + k)
                    correntes = objeto.get_resultados_corrente()
                    print(f'Contribuição da SE {k}: '
                          f'\nFase A: {round(correntes[0] / 1000, 2)} kA')
                    if j == '3F':
                        print(
                            f'Fase B: {round(correntes[2] / 1000, 2)} kA'
                            f'\nFase C: {round(correntes[4] / 1000, 2)} kA')

    elif valida_cc == "4":

        print('\nEstudo em execução.....')
        # Executa estudo de superação de cabos para_raios - Recebe matriz gigante contendo muitas saídas
        resultado = objeto.estudo_pr_DSS(barras_modelo, linhas_modelo)

        dados = pd.DataFrame(resultado, columns=['Barra_Modelo', 'PR', 'Fase', 'Vão', 'PR_1', 'PR_2', 'R_torre', 'Linecode'])
        filtro_max = pd.DataFrame()

        # Filtragem dos dados, apresentando apenas maiores correntes

        for linha in linhas_modelo:
            filtro = dados.loc[dados['Vão'] == linha]

            # Pega máxima corrente no PR_1 para o vão específico
            filtro_max = filtro_max.append(filtro.loc[(filtro['PR_1']) == filtro['PR_1'].max()])
            # Pega máxima corrente no PR_2 para o vão específico

            objeto.ativa_elemento('line.' + linha)
            if objeto.fases_linha_opendss() == 5 or objeto.fases_linha_opendss() == 8:
                filtro_max = filtro_max.append(filtro.loc[(filtro['PR_2']) == filtro['PR_2'].max()])

        print('\nOs resultados serão salvos em uma planilha chamada Analises_ISA_CTEEP_aux.xlsx.')
        writer = pd.ExcelWriter('Analises_ISA_CTEEP_aux.xlsx', engine='xlsxwriter')

        filtro_max.to_excel(writer, 'PR', header=False, index=False, startcol=1, startrow=2)
        #dados.to_excel(writer, 'PR', header=False, index=False, startcol=1, startrow=2)
        writer.save()

        #print(filtro_max.head())

        plotar = input('Deseja plotar as distribuições máximas de curto-circuito nos para-raios ? \n'
                        'Eixo y: Máximas correntes calculadas. \n'
                        'Eixo x: Torre em curto-circuito: \n'
                        '1 - Sim \n'
                        'Qualquer tecla - Não \n')

        if plotar == "1":
            max_pr1 = pd.DataFrame()
            max_pr2 = pd.DataFrame()
            for linha in linhas_modelo:
                filtro = dados.loc[dados['Vão'] == linha]
                max_pr1 = max_pr1.append(filtro.loc[(filtro['PR_1']) == filtro['PR_1'].max()])
                max_pr2 = max_pr2.append(filtro.loc[(filtro['PR_2']) == filtro['PR_2'].max()])

            plt.figure()
            plt.plot(max_pr1['Vão'], max_pr1['PR_1'], label='PR_1')
            plt.plot(max_pr2['Vão'], max_pr2['PR_2'], label='PR_2')
            plt.xlabel('Torre')
            plt.ylabel('Corrente [A]')
            plt.grid(True)
            plt.legend(loc=0)
            plt.title('Máximas correntes nos para-raios em função da torre em falta')
            plt.show()

    else:
        valida_cc = "0"

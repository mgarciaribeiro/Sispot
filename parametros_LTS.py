"""
Calculando parâmetros LTS
Versão segmentada por funções para integração com Excel
"""

import numpy as np
from scipy import special
from scipy import linalg
import pandas as pd
import xlwings as xw

import math
import cmath
import random

import matplotlib.pyplot as plt

wb = xw.Book('parametros_LTS.xlsm') # o arquivo tem que estar na mesma pasta
# ab = wb.sheets['aba'] # exemplo de aba
dados = wb.sheets['DADOS'] # aba com os dados da LTS

# todos os parâmetros de entrada serão computados aqui a partir do que foi preenchido pelo usuário
# ------------------------------------------------------------------------------------------------
w_r0 = float(dados.range('U4').value) / 10                                      # Entre com o diâmetro interno do núcleo em [cm]
w_r1 = float(dados.range('U5').value) / 10                                      # Entre com o diâmetro externo (nominal) do núcleo em [cm]
w_pc = float(dados.range('U20').value)                                          # Entre com a resistividade elétrica do material do núcleo em 20 graus Celsius
w_tp = float(dados.range('U23').value)                                          # Temperatura a ser considerada no núcleo em graus Celsius
w_rd = float(dados.range('U22').value)                                          # Resistência DC no núcleo em [ohm/km] em 20 graus Celsius
w_mt = '1' if dados.range('U19').value == 'Alumínio' else '2'                   # Material do núcleo
w_ks = float(dados.range('U45').value)                                          # Digite o valor do fator ks para consideração acerca do efeito pelicular
w_kp = float(dados.range('U46').value)                                          # Digite o valor do fator kp para consideração acerca do efeito de proximidade
w_fq = float(dados.range('U24').value)                                          # Digite a frequência de cálculo de regime em [Hz]
w_Nc = 3                                                                        # Entre com o número de cabos da instalação (vou ter que editar, sem jeito)

# print('Diâmetro interno (cm): ', w_r0)
# print('Diâmetro externo (cm): ', w_r1)
# print('Resistividade núcleo (ohm.m): ', w_pc)
# print('Temperatura (Celsius): ', w_tp)
# print('Resistência DC: ', w_rd)
# print('Material núcleo: ', w_mt)
# print('ks: ', w_ks)
# print('kp: ', w_kp)
# print('Frequência (Hz): ', w_fq)

w__X = dados.range('AE3:AE5').value                                             # Posições em X
w__Y = dados.range('AA3:AA5').value                                             # Posições em Y

print(w__X, '\n', w__Y)

w_uc = float(dados.range('U25').value)                                          # Entre com a permeabilidade magnética relativa do núcleo
w_er = 0.0                                                                      # Perturbação
w_s1 = float(dados.range('U7').value) / 10                                      # Entre com a espessura da camada semicondutora entre núcleo e isolação [cm]
w_xl = float(dados.range('U9').value) / 10                                      # Entre com a espessura da primeira camada isolante [cm]
w_s2 = float(dados.range('U11').value) / 10                                     # Entre com a espessura da camada semicondutora entre isolação e blindagem [cm]

# print('Dados da Blindagem:')

w_bl = '2' if 'Fios' in dados.range('U28').value else '1'                       # Blindagem metálica de fios?
w_bd = 2 if 'Fita' in dados.range('U28').value else 1                           # Blindagem composta por fios + fitas metálicas?
w_eb = dados.range('U13').value / 2 if w_bd == 2 else 0.0                       # Digite a espessura da camada de fitas metálicas em [cm]
w_eb = w_eb / 10 # cm
w_fi = int(dados.range('U31').value)
w_ef = dados.range('U13').value / 2 if w_bd == 2 else dados.range('U13').value  # Digite a espessura (diâmetro) dos fios em [cm]
w_ef = w_ef / 10 # cm
w_mb = 2
w_tb = w_tp

w_pb = float(dados.range('U29').value)                                          # Entre com a resistividade elétrica dos fios da blindagem em 20 graus Celsius
w_ml = '1' if dados.range('U27').value == 'Cobre' else '2'                      # Especifique o material da blindagem de fios
w_pf = float(dados.range('U29').value)                                          # Entre com a resistividade elétrica das fitas da blindagem em 20 graus Celsius
w_mf = '1' if dados.range('U27').value == 'Cobre' else '2'                      # Especifique o material da blindagem de fitas
w_et = float(dados.range('U13').value) / 10                                     # Entre com a espessura da blindagem metálica (considerada puramente tubular) em [cm]
w_ee = float(dados.range('U15').value) / 10                                     # Entre com a espessura da segunda camada isolante em [cm]
w_ps = float(dados.range('U29').value)                                          # Entre com a resistividade elétrica da blindagem em 20 graus Celsius
w_gs = w_tp
w_ms = '2' if dados.range('U27').value == 'Cobre' else '1'                      # Especifique o material da blindagem
w_us = float(dados.range('U30').value)                                          # Entre com a permeabilidade magnética relativa da blindagem

# print('Blindagem metálica de fios: ', w_bl)
# print('Blindagem metálica de fios e fita: ', w_bd)
# print('Espessura fita: ', w_eb)
# print('Espessura fios: ', w_ef)
# print('Quantidade de fios: ', w_fi)
# print('Resistividade fios blindagem: ', w_pb)
# print('Resistividade fita blindagem: ', w_pf)
# print('Material fios blindagem: ', w_ml)
# print('Material fios blindagem: ', w_mf)
# print('-----------------------------------------------------------------------')

# print('Resistividade blindagem: ', w_ps)
# print('Espessura da segunda camada isolante (cm): ', w_ee)
# print('Permeabilidade da blindagem: ', w_us)

w_u1 = float(dados.range('U33').value)                                          # Entre com a permeabilidade magnética relativa da primeira camada isolante
w_ci = '1'                                                                      # Deseja corrigir a permeabilidade magnética da primeira camada isolante para levar em conta o efeito da blindagem metálica de fios
w_u2 = float(dados.range('U36').value)                                          # Entre com a permeabilidade magnética relativa da segunda camada isolante
w_e1 = float(dados.range('U34').value)                                          # Entre com a permissividade elétrica relativa da primeira camada isolante
w_ce = '1'                                                                      # Deseja corrigir a permissividade elétrica da primeira camada isolante para levar em conta o efeito das camadas semicondutoras
w_e2 = float(dados.range('U37').value)                                          # Entre com a permissividade elétrica relativa da segunda camada isolante
w_gn = float(dados.range('U39').value)                                          # Entre com a resistividade elétrica do solo [ohm.m]
w_gu = float(dados.range('U40').value)                                          # Entre com a permeabilidade magnética relativa do solo

# print('Permeabilidade da primeira camada isolante: ', w_u1)
# print('Permeabilidade da segunda camada isolante: ', w_u2)
# print('Permissividade da primeira camada isolante: ', w_e1)
# print('Permissividade da segunda camada isolante: ', w_e2)
# print('Resistividade solo: ', w_gn)
# print('Permeabilidade solo: ', w_gu)

def solve_LTS(method = 2):

    # Entrada de dados

    r_0 = 1 / 2 * (10 ** -2) * w_r0

    r_1 = 1 / 2 * (10 ** -2) * w_r1

    p_c = w_pc

    graus_c = w_tp

    r_dc = (10 ** -3) * w_rd

    material_c = w_mt

    if material_c == '2':
        r_dc_0 = r_dc * (1 + 0.00393 * (graus_c - 20))

    else:
        r_dc_0 = r_dc * (1 + 0.00403 * (graus_c - 20))

    # Efeito Pelicular

    if material_c == "1":
        if r_0 != 0:
            ks = ((r_1 - r_0) / (r_1 + r_0)) * ((r_1 + 2 * r_0) / (r_1 + r_0)) ** 2
        else:
            ks = w_ks

    else:
        if r_0 != 0:
            ks = ((r_1 - r_0) / (r_1 + r_0)) * ((r_1 + 2 * r_0) / (r_1 + r_0)) ** 2
        else:
            ks = w_ks

    # Efeito de proximidade

    if r_0 != 0:
        kp = w_kp

    else:
        kp = w_kp

    f_rps = w_fq

    xs = math.sqrt(8 * math.pi * f_rps / r_dc_0 * 10 ** -7 * ks)

    if xs > 0 and xs <= 2.8:
        ys = xs ** 4 / (192 + 0.8 * xs ** 4)

    elif xs > 2.8 and xs <= 3.8:
        ys = -0.136 - 0.0177 * xs + 0.0563 * xs ** 2

    else:
        ys = 0.354 * xs - 0.733

    xp = math.sqrt(8 * math.pi * f_rps / r_dc_0 * 10 ** -7 * kp)

    N = w_Nc
    X = w__X
    Y = w__Y

    # for i in range(N):
    #     X.append(float(input(f'\nDigite a posição no eixo x do cabo {i+1} em metros: ')))
    #     Y.append(float(input(f'Digite a profundidade do cabo {i + 1} em metros (valor positivo): ')))

    # X = [-0.35, -0.35, -0.35, 0.35, 0.35, 0.35] # Posição dos cabos no eixo x [m]
    # Y = [1.25, 1.6, 1.95, 1.25, 1.6, 1.95] # Posição dos cabos no eixo y (valor positivo) [m]


    yp = xp ** 4 / (192 + 0.8 * xp ** 4) * (2 * r_1 / math.sqrt((X[0] - X[1]) ** 2 + (Y[0] - Y[1]) ** 2)) ** 2 * (0.312 * (2 * r_1 / math.sqrt((X[0] - X[1]) ** 2 + (Y[0] - Y[1]) ** 2)) ** 2 +
                                                                                                                  1.18 / (xp ** 4 / (192 + 0.8 * xp ** 4) + 0.27))

    r_ac_iec = r_dc_0*(1 + ys + yp)

    print(f'\nResistência AC em [ohm/km] conforme IEC em {graus_c} graus Celsius: {round(1000 * r_ac_iec, 5)} \n')

    eps = 0.1  # Tolerancia entre R calculado por Wedepohl e R calculado pela IEC em %

    w = 2 * math.pi * f_rps

    u_c = w_uc
    u_c = u_c * (4 * math.pi * 10 ** -7)  # Permeabilidade magnética do núcleo [H/m]

    m_c = cmath.sqrt(w * u_c / p_c * 1j)

    Z1 = p_c * m_c / (2 * math.pi * r_1) * 1/(np.tanh(0.777 * m_c * r_1)) + 0.356 * p_c / (math.pi * r_1 ** 2)

    n_iter = 100000  # Número máximo de iterações
    iter = 0
    p_c_orig = p_c

    p = 1 / 100 * w_er

    while np.abs((np.real(Z1) - r_ac_iec) / r_ac_iec) * 100 > eps and iter <= n_iter:
        p_c = p_c_orig * (1 + random.uniform(-p, p))  # Perturba p_c em +/- p%
        m_c = cmath.sqrt(w * u_c / p_c * 1j)

        if r_0 == 0:
            Z1 = p_c * m_c / (2 * math.pi * r_1) * (special.iv(0, m_c * r_1) / special.iv(1, m_c * r_1))
        else:
            Z1 = p_c * m_c / (2 * math.pi * r_1) * ((special.iv(0, m_c * r_1) * special.kv(1, m_c * r_0) +
                                                     special.kv(0, m_c * r_1) * special.iv(1, m_c * r_0)) /
                                                    (special.iv(1, m_c * r_1) * special.kv(1, m_c * r_0) -
                                                     special.iv(1, m_c * r_0) * special.kv(1, m_c * r_1)))

        iter = iter + 1
        if iter == n_iter:
            iter = 0
            # p = 1 / 100 * float(input(f'Iteração {n_iter} atingida sem convergência, digite a nova porcentagem de '
                                      # f'perturbação em p_c (iteração anterior = {100*p}%). \n'))
            # terminando o programa
            # print('Convergência não atingida para resistência em ohm/km, terminando o programa. Tente outra perturbação...')
            # exit()
            # incrementa a perturbação até conseguir
            p = p + 1


    print(f'Resistência em ohm/km calculada em {iter} iterações: {round(1000 * np.real(Z1), 5)}. \n'
          f'Erro: {round(((np.real(Z1) - r_ac_iec) / r_ac_iec) * 100, 4)} % \n'
          f'Resistividade estimada para o condutor em [ohm.m]: {p_c}. \n')

    e_sc_in = (10 ** -2) * w_s1

    e_isol_1 = (10 ** -2) * w_xl

    e_sc_out = (10 ** -2) * w_s2

    r_2 = r_1 + e_sc_in + e_isol_1 + e_sc_out

    a = r_1 + e_sc_in

    b = a + e_isol_1

    blindagem = w_bl

    if blindagem == "2":

        blindagem_dupla = w_bd

        if blindagem_dupla == 2:

            e_fita = (10 ** -2) * w_eb

        else:

            e_fita = 0.0

        n_fios = w_fi

        d_f = (10 ** -2) * w_ef

        metodo_blindagem = w_mb

        if metodo_blindagem == 1:
            r_3 = r_2 + d_f + e_fita

        else:

            if blindagem_dupla == 2:

                area_cu = math.pi * n_fios * (d_f / 2) ** 2
                area_al = math.pi * ((r_2 + d_f + e_fita) ** 2 - (r_2 + d_f) ** 2)
                area_bl = area_cu + area_al

                graus_s = w_tb

                p_cu = w_pb

                material_cu = w_ml

                if material_cu == '2':

                    p_cu = p_cu * (1 + 0.00403 * (graus_s - 20))

                else:

                    p_cu = p_cu * (1 + 0.00393 * (graus_s - 20))

                p_al = w_pf

                material_al = w_mf

                if material_al == '2':

                    p_al = p_al * (1 + 0.00403 * (graus_s - 20))

                else:

                    p_al = p_al * (1 + 0.00393 * (graus_s - 20))

    #            print('\nTendo em vista que há blindagem composta por fios e fitas, assume-se como resistividade a 20º Celsius: \n'
    #                  'Referência TB 531: \n'
    #                  'Cobre - 1.7241E-8 ohm.m \n'
    #                  'Alumínio - 2.8264E-8 ohm.m \n')

                R_bl = (p_al / area_al * p_cu / area_cu) / (p_al / area_al + p_cu / area_cu)

                p_s = area_bl * R_bl
                r_3 = math.sqrt(area_bl / math.pi + r_2 ** 2)

            else:

                area_s = math.pi * n_fios * (d_f / 2) ** 2
                r_3 = math.sqrt(area_s / math.pi + r_2 ** 2)

    else:
        blindagem_dupla = 1
        metodo_blindagem = 2
        e_s = (10 ** -2) * w_et
        r_3 = r_2 + e_s

    e_isol_2 = (10 ** -2) * w_ee

    r_4 = r_3 + e_isol_2

    if blindagem_dupla != 2:

        # Valores de referência para resistividade, tirados da Brochura 531 do Cigré:
        # Cobre - p_c = 1.7241E-8 ohm.m
        # Alumínio - pc = 2.8264E-8 ohm.m

        # p_c = 4.4969E-8 # Resistividade do núcleo [ohm.m]
        # p_s = 1.68E-8 # Resistividade da blindagem [ohm.m]

        p_s = w_ps

        graus_s = w_gs

        material_s = w_ms

        if material_s == '2':
            p_s = p_s * (1 + 0.00393 * (graus_s - 20))

        else:
            p_s = p_s * (1 + 0.00403 * (graus_s - 20))

        # Correção da resistividade da blindagem em funcao do parâmetro metodo_blindagem
        if metodo_blindagem == 1:
            p_s = p_s * (math.pi * (r_3 ** 2 - r_2 ** 2)) / (math.pi * n_fios * (d_f / 2) ** 2)

    u_s = w_us
    u_1 = w_u1

    if blindagem == "2":
        corrige_ind = w_ci

        # if corrige_ind == "2":
        #     length_of_lay = (10 ** -2) * float(input('Digite o comprimento em [cm] necessário para a blindagem '
        #                                              'completar uma volta em torno da primeira camada isolante: '))
        #     u_1 = u_1*(1 + (2 * (1 / length_of_lay) ** 2 * math.pi ** 2 * (r_2 ** 2 - r_1 ** 2)) / math.log(r_2 / r_1))
        #     print(f'Permabilidade magnética relativa após a correção: {u_1} \n')

    u_2 = w_u2
    e_1 = w_e1

    corrige_e_1 = w_ce

    # if corrige_e_1 == "2":
    #     e_1 = e_1 * math.log(r_2 / r_1) / math.log(b / a)
    #     print(f'Permissividade elétrica relativa após a correção: {e_1} \n')

    e_2 = w_e2
    p_solo = w_gn
    u_solo = w_gu

    u_s = u_s * (4 * math.pi * 10 ** -7)  # Permeabilidade magnética da blindagem [H/m]
    u_1 = u_1 * (4 * math.pi * 10 ** -7)  # Permeabilidade magnética da primeira camada isolante [H/m]
    u_2 = u_2 * (4 * math.pi * 10 ** -7)  # Permeabilidade magnética da segunda camada isolante [H/m]
    # Teste e_1 = 2.88
    e_1 = e_1 * 8.85 * 10 ** -12  # Permissividade dielétrica da primeira camada isolante [F/m]
    # Teste e_2 = 2.89
    e_2 = e_2 * 8.85 * 10 ** -12  # Permissividade dielétrica da segunda camada isolante [F/m]
    # p_solo = 100 # Resistividade do solo [ohm.m]
    u_solo = u_solo * (4 * math.pi * 10 ** -7)  # permeabilidade magnética do solo [H/m]

    L = 1000  # Comprimento do trecho [m]

    # Entrada da instalação

    # Fórmulas utilizadas conforme Wedepohl

    # Z1 - Impedancia interna do condutor central [ohm/m]

    f = w_fq

    w = 2 * math.pi * f

    m_c = cmath.sqrt(w * u_c / p_c * 1j)

    # método de cálculo
    mtd = '2' if method == 2 else '1'

    print('\nConsiderações sobre o cálculo de Z1 (Impedância interna do núcleo)')
    metodo_z1 = mtd

    if metodo_z1 == '1':
        Z1 = p_c * m_c / (2 * math.pi * r_1) * 1/(np.tanh(0.777 * m_c * r_1)) + 0.356 * p_c / (math.pi * r_1 ** 2)

    else:
        # Aplicação das funções de Bessel modificadas. Utilização da biblioteca special do scipy
        # iv - Função de Bessel modificada de primeira espécie e ordem real
        # iv(1, m) - ordem 1 e argumento complexo m
        # iv(0, m) - ordem 0 e argumento complexo m
        # kv - Função de Bessel modificada de segunda espécie e ordem real
        if r_0 == 0:
            Z1 = p_c * m_c / (2 * math.pi * r_1) * (special.iv(0, m_c * r_1) / special.iv(1, m_c * r_1))
        else:
            Z1 = p_c * m_c / (2 * math.pi * r_1) * ((special.iv(0, m_c * r_1) * special.kv(1, m_c * r_0) +
                                                     special.kv(0, m_c * r_1) * special.iv(1, m_c * r_0)) /
                                                    (special.iv(1, m_c * r_1) * special.kv(1, m_c * r_0) -
                                                     special.iv(1, m_c * r_0) * special.kv(1, m_c * r_1)))


    # Z2 - Impedancia devido a variação do campo magnético na isolação principal

    Z2 = w * u_1 / (2 * math.pi) * math.log(r_2 / r_1) * 1j

    # Z6 - Impedancia devido a variação do campo magnético na isolação externa

    Z6 = w * u_2 / (2 * math.pi) * math.log(r_4 / r_3) * 1j

    D = r_3 - r_2

    m_s = cmath.sqrt(w * u_s / p_s * 1j)

    print('\nConsiderações sobre o cálculo de Z3, Z4 e Z5 (Impedâncias da blindagem)')
    metodo_blindagem = mtd

    if metodo_blindagem == '1':
        # Z3 - Impedancia dada pela queda de tensão na superfície interna da blindagem devido a corrente no condutor
        Z3 = p_s * m_s / (2 * math.pi * r_2) * 1 / (np.tanh(m_s * D) - p_s / (2 * math.pi * r_2 * (r_2 + r_3)))
        # Z4
        Z4 = p_s * m_s / (math.pi * (r_2 + r_3)) * 1 / np.sinh(m_s * D)
        # Z5 - Impedancia dada pela queda de tensão na superfície externa da blindagem devido a corrente pelo solo
        Z5 = p_s * m_s / (2 * math.pi * r_3) * 1 / np.tanh(m_s * D) + p_s / (2 * math.pi * r_3 * (r_2 + r_3))

    else:
        # Aplicação das funções de Bessel modificadas. Utilização da biblioteca special do scipy
        # iv - Função de Bessel modificada de primeira espécie e ordem real
        # iv(1, m) - ordem 1 e argumento complexo m
        # iv(0, m) - ordem 0 e argumento complexo m
        # iv - Função de Bessel modificada de segunda espécie e ordem real
        Z3 = p_s * m_s / (2 * math.pi * r_2) * \
             (special.iv(0, m_s * r_2) * special.kv(1, m_s * r_3) + special.kv(0, m_s * r_2) * special.iv(1, m_s * r_3)) / \
             (special.iv(1, m_s * r_3) * special.kv(1, m_s * r_2) - special.iv(1, m_s * r_2) * special.kv(1, m_s * r_3))

        Z4 = p_s / (2 * math.pi * r_2 * r_3) * 1 / (special.iv(1, m_s * r_3) * special.kv(1, m_s * r_2) - special.iv(1, m_s * r_2) * special.kv(1, m_s * r_3))

        Z5 = p_s * m_s / (2 * math.pi * r_3) * \
             (special.iv(0, m_s * r_3) * special.kv(1, m_s * r_2) + special.kv(0, m_s * r_3) * special.iv(1, m_s * r_2)) / \
             (special.iv(1, m_s * r_3) * special.kv(1, m_s * r_2) - special.iv(1, m_s * r_2) * special.kv(1, m_s * r_3))

    # Z7 - Impedancia do solo

    m_solo = cmath.sqrt(w * u_solo / p_solo * 1j)

    gama = 0.5772156649

    Z_int = np.zeros((2 * len(X), 2 * len(X)), dtype=complex)
    Z_earth = np.zeros((2 * len(X), 2 * len(X)), dtype=complex)

    for i in range(len(X)):
        Z_int[i][i] = Z1 + Z2 + Z3 + Z5 + Z6 - 2 * Z4
        Z_int[i][i + len(X)] = Z5 + Z6 - Z4
        Z_int[i + len(X)][i] = Z_int[i][i + len(X)]
        Z_int[i + len(X)][i + len(X)] = Z5 + Z6


    # Montagem da matriz de retorno pela terra

    X_aux = X.copy()
    Y_aux = Y.copy()

    X.extend(X)
    Y.extend(Y)

    # Permite métodos de cálculo para impedâncias do solo
    # Wedepohl e Saad

    print('\nConsiderações sobre o cálculo de Z7 (Impedâncias de terra)')
    metodo_earth = mtd

    for i in range(len(X)):
        for k in range(i, len(X)):
            if k == i:
                if metodo_earth == '1':
                    Z_earth[i][k] = 1j * w * u_solo / (2 * math.pi) * (-cmath.log(gama * m_solo * r_4 / 2) + 1 / 2 - 4 * m_solo * Y[i] / 3)
                else:
                    Z_earth[i][k] = p_solo * m_solo ** 2 / (2 * math.pi) * (special.kv(0, m_solo * r_4) + 2 / (4 + m_solo ** 2 * r_4 ** 2) * cmath.exp(-2 * Y[i] * m_solo))
            elif k == i + len(X) / 2:
                if metodo_earth == '1':
                    Z_earth[i][k] = 1j * w * u_solo / (2 * math.pi) * (-cmath.log(gama * m_solo * r_4 / 2) + 1 / 2 - 4 * m_solo * Y[i] / 3)
                else:
                    Z_earth[i][k] = p_solo * m_solo ** 2 / (2 * math.pi) * (special.kv(0, m_solo * r_4) + 2 / (4 + m_solo ** 2 * r_4 ** 2) * cmath.exp(-2 * Y[i] * m_solo))
            else:
                if metodo_earth == '1':
                    Z_earth[i][k] = 1j * w * u_solo / (2 * math.pi) * (-cmath.log(gama * m_solo * math.sqrt((X[i] - X[k]) ** 2 + (Y[i] - Y[k]) ** 2) / 2) + 1 / 2 - 2 / 3 * m_solo * (Y[i] + Y[k]))

                else:
                    Z_earth[i][k] = p_solo * m_solo ** 2 / (2 * math.pi) * (special.kv(0, m_solo * math.sqrt((X[i] - X[k])**2 + (Y[i] - Y[k])**2)) + 2 / (4 + m_solo ** 2 * (np.abs(X[i] - X[k])) ** 2) * cmath.exp(-1 * (Y[i] + Y[k]) * m_solo))

            Z_earth[k][i] = Z_earth[i][k]

    # Montagem da matriz de impedancias serie

    Z_serie = Z_int + Z_earth

    # Montagem das matrizes relativas à admitância da LT

    Y_shunt = np.zeros([len(X), len(X)], dtype=complex)

    for i in range(int(len(X) / 2)):
        Y_shunt[i][i] = w * 2 * math.pi * e_1 / math.log(r_2 / r_1) * 1j

        Y_shunt[i][i + int(len(X) / 2)] = -Y_shunt[i][i]
        Y_shunt[i + int(len(X) / 2), i] = Y_shunt[i][i + int(len(X) / 2)]
        Y_shunt[i + int(len(X) / 2), i + int(len(X) / 2)] = Y_shunt[i][i] + w * 2 * math.pi * e_2 / math.log(r_4 / r_3) * 1j

    # Calculo dos parametros de sequencia da matriz
    # Matriz R

    # Construção da matriz R conforme transposições

    if N == 3:
        transp = '1'

        if transp == "2":
            R = np.array([[1, 0, 0, 0, 0, 0],
                          [0, 1, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0],
                          [0, 0, 0, 0, 1, 0],
                          [0, 0, 0, 0, 0, 1],
                          [0, 0, 0, 1, 0, 0]])
        else:
            R = np.array([[0, 1, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0],
                          [1, 0, 0, 0, 0, 0],
                          [0, 0, 0, 1, 0, 0],
                          [0, 0, 0, 0, 1, 0],
                          [0, 0, 0, 0, 0, 1]])

    else:
        transp = '1'

        if transp == "2":
            R = np.array([[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]])

        elif transp == "4":
            R = np.array([[1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0]])

        elif transp == "3":
            R = np.array([[1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0]])

        else:
            R = np.array([[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0],
                          [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]])

    Zf = 1 / 3 * (Z_serie + (np.linalg.inv(R).dot(Z_serie)).dot(R) + R.dot(Z_serie).dot(np.linalg.inv(R)));

    Yf = 1 / 3 * (Y_shunt + np.linalg.inv(R).dot(Y_shunt).dot(R) + R.dot(Y_shunt).dot(np.linalg.inv(R)));

    # Elimina blindagens

    Z_cc = Zf[0:N, 0:N]
    Z_cs = Zf[0:N, N:2 * N]
    Z_sc = Zf[N:2 * N, 0:N]
    Z_ss = Zf[N:2 * N, N:2 * N]

    Zeq = Z_cc - (Z_cs.dot(np.linalg.inv(Z_ss))).dot(Z_sc)

    # Matriz dos coeficientes de potenciais

    P_1 = 1 / (1j * w) * (Yf)

    Pf = np.linalg.inv(P_1)

    P_cc = Pf[0:N, 0:N]
    P_cs = Pf[0:N, N:2 * N]
    P_sc = Pf[N:2 * N, 0:N]
    P_ss = Pf[N:2 * N, N:2 * N]

    Peq = P_cc - (P_cs.dot(np.linalg.inv(P_ss))).dot(P_sc)

    alfa = math.cos(120 * math.pi / 180) + 1j * math.sin(120 * math.pi / 180);

    T = np.array([[1, 1, 1], [1, alfa ** 2, alfa], [1, alfa, alfa ** 2]])
    T_1 = np.linalg.inv(T)

    if N == 6:
        T_aux_1 = np.concatenate((T, np.zeros([3, 3])), axis=1)
        T_aux_2 = np.concatenate((np.zeros([3, 3]), T), axis=1)
        T = np.concatenate((T_aux_1, T_aux_2), axis=0)

        T_aux_1 = np.concatenate((T_1, np.zeros([3, 3])), axis=1)
        T_aux_2 = np.concatenate((np.zeros([3, 3]), T_1), axis=1)
        T_1 = np.concatenate((T_aux_1, T_aux_2), axis=0)


    Z_012 = (T_1.dot(Zeq)).dot(T)

    P_012 = (T_1.dot(Peq)).dot(T)

    # Capacitancias de sequencia

    C_012 = np.linalg.inv(P_012)

    Y_012 = 1j * w * C_012

    # Apresenta parametros de sequencia

    print('Impedancia de Sequencia zero da LT em ohm/km')
    print(f'Z_0 = {np.round(1000 * Z_012[0][0], 4)}')

    dados.range('B21').value = str(1000 * Z_012[0][0]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('C21').value = str(1000 * Z_012[0][0]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    print('Impedancia de sequencia positiva da LT em ohm/km')
    print(f'Z_1 = {np.round(1000 * Z_012[1][1], 4)}')

    dados.range('D21').value = str(1000 * Z_012[1][1]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('E21').value = str(1000 * Z_012[1][1]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    print('Impedancia de sequencia negativa da LT em ohm/km')
    print(f'Z_2 = {np.round(1000 * Z_012[2][2], 4)}')

    dados.range('F21').value = str(1000 * Z_012[2][2]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('G21').value = str(1000 * Z_012[2][2]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    print('Admitancia de Sequencia zero da LT em micro mho/km')
    print(f'Y_0 = {np.round(10** 6 * 1000 * Y_012[0][0], 4)}')

    dados.range('B24').value = str(1E9 * Y_012[0][0]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('C24').value = str(1E9 * Y_012[0][0]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    print('Admitancia de Sequencia positiva da LT em micro mho/km')
    print(f'Y_1 = {np.round(10** 6 * 1000 * Y_012[1][1], 4)}')

    dados.range('D24').value = str(1E9 * Y_012[1][1]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('E24').value = str(1E9 * Y_012[1][1]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    print('Admitancia de Sequencia negativa da LT em micro mho/km')
    print(f'Y_2 = {np.round(10** 6 * 1000 * Y_012[2][2], 4)}')

    dados.range('F24').value = str(1E9 * Y_012[2][2]).replace('(', '').replace(')', '').replace('j', '').split('+')[0]
    dados.range('G24').value = str(1E9 * Y_012[2][2]).replace('(', '').replace(')', '').replace('j', '').split('+')[1]

    matriz_z_seq = pd.DataFrame(np.round(Z_012 * 1000, 5))
    matriz_y_seq = pd.DataFrame(np.round(Y_012 * 1000 * 10 ** 6, 5))

    # writer = pd.ExcelWriter('LSCable_1400mm_metodos_aux.xlsx', engine='xlsxwriter')
    #
    # matriz_z_seq.to_excel(writer, 'Z1=' + metodo_z1 + 'Zs=' + metodo_blindagem + 'Z7=' + metodo_earth +
    #                       'Transp=' + transp, header=False, index=False, startcol=2, startrow=2)
    # matriz_y_seq.to_excel(writer, 'Z1=' + metodo_z1 + 'Zs=' + metodo_blindagem + 'Z7=' + metodo_earth +
    #                       'Transp=' + transp, header=False, index=False, startcol=2, startrow=4 + N)
    #
    # writer.save()

    saida_atp = '2'

    if saida_atp == "2":
        print('Dados para entrada na rotina Cable Constants do ATP: \n'
              'Dados comuns relativos à geometria do cabo: \n'
              'CORE: \n'
              f'Rin [m]: {round(r_0, 5)} \n'
              f'Rout [m]: {round(r_1, 5)} \n'
              f'Rho [ohm*m]: {p_c} \n'
              f'mu : {round(u_c / (4 * math.pi * 10 ** -7), 2)} \n'
              f'mu (ins): {round(u_1 / (4 * math.pi * 10 ** -7), 4)} \n'
              f'eps(ins): {round(e_1 / (8.85 * 10 ** -12), 2)} \n'
              f'\nSHEATH: \n'
              f'Rin [m]: {round(r_2, 5)} \n'
              f'Rout [m]: {round(r_3, 5)} \n'
              f'Rho [ohm*m]: {p_s} \n'
              f'mu : {round(u_s / (4 * math.pi * 10 ** -7), 2)} \n'
              f'mu (ins): {round(u_2 / (4 * math.pi * 10 ** -7), 2)} \n'
              f'eps(ins): {round(e_2 / (8.85 * 10 ** -12), 2)} \n'
              f'\nR5 [m]: {round(r_4, 5)} \n'
              f'\nDados relativos às posições dos cabos: \n'
              f'Posição Horizontal [m]: {X[0:N]} \n'
              f'Posição Vertical [m]: {Y[0:N]} \n'
              f'\nResistividade do solo Rho [ohm*m]: {p_solo} \n')
        dados.range('D28').value = round(r_0, 5)
        dados.range('D29').value = round(r_1, 5)
        dados.range('D30').value = p_c
        dados.range('D31').value = round(u_c / (4 * math.pi * 10 ** -7), 2)
        dados.range('D32').value = round(u_1 / (4 * math.pi * 10 ** -7), 4)
        dados.range('D33').value = round(e_1 / (8.85 * 10 ** -12), 2)

        dados.range('G28').value = round(r_2, 5)
        dados.range('G29').value = round(r_3, 5)
        dados.range('G30').value = p_s
        dados.range('G31').value = round(u_s / (4 * math.pi * 10 ** -7), 2)
        dados.range('G32').value = round(u_2 / (4 * math.pi * 10 ** -7), 4)
        dados.range('G33').value = round(e_2 / (8.85 * 10 ** -12), 2)

        dados.range('G26').value = round(r_4, 5)

    matriz_z = '2'
    if matriz_z == "2":
        print('Matriz Z de fase em ohm/km: \n'
              f'{Z_serie * 1000}')

    matriz_y = '2'
    if matriz_y == "2":
        print('Matriz Y de fase em mho/km: \n'
              f'{Y_shunt * 1000}')

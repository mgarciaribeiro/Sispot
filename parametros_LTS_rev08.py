"""
Programa para cálculo de parâmetros de LT subterrânea
rev 00 - Recebe parâmetros conforme entrada de dados no ATP
rev 01 - Torna entrada de parâmetros interativa
rev 02 - Possibilida escolha de modos de cálculo de alguns parâmetros
rev 03 - Considerações sobre camadas semicondutoras da isolação e blindagem composta por fios
rev 04 - Permite a correção da indutância devido ai efeito solenóide das blindagens
rev 05 - Correção da resistividade do núcleo: Estima diâmetro de modo a deixar a resistência AC o mais próximo
         possível do cálculo da norma 602871-1-1
rev 07 - Insere possibilidade de cálculo de cabo isolado a óleo
rev 08 - Possibilita duas formas para o cálculo dos parâmetros da blindagem:
        a) Mantém resistividade e altera raio (paper GUstavsen)
        b) Altera resistividade e mantém espessura da blindagem
Referências:
Wedepohl - Transient Analysis of underground power-transmission systems: System-model and wave-propagation characteristics
Ametani - A General Formulation of Impedance and Admittance of Cables
Gustavsen - Panel Session on Data for Modeling System Transients Insulated Cables
Cigré TB 531
Saad - A Closed-Form Approximation for Ground Return Impedance of Underground Cables
"""
import math
import cmath
import numpy as np
from scipy import special
import pandas as pd
import random

# Calcula parametros de sequencia de LTs isoladas

# Entrada de dados

r_0 = 1 / 2 * (10 ** -2) * float(input('Entre com o diâmetro interno do núcleo em [cm]. Lembrando que este parâmetro '
                                       'só é diferente de zero caso o condutor seja tubular: \n'))

r_1 = 1 / 2 * (10 ** -2) * float(input('Entre com o diâmetro externo (nominal) do núcleo em [cm]: \n'))

p_c = float(input('\nEntre com a resistividade elétrica do material do núcleo em 20 graus Celsius: \n'
                  'Referência TB 531: \n'
                  'Cobre - 1.7241E-8 ohm.m \n'
                  'Alumínio - 2.8264E-8 ohm.m \n'))

graus_c = float(input('\nTemperatura a ser considerada no núcleo em graus Celsius: '))

r_dc = (10 ** -3) * float(input('\nResistência DC no núcleo em [ohm/km] em 20 graus Celsius: \n'
                   'Valores máximos conforme IEC 60228 para condutores do tipo encordoados ou Milliken: \n'
                   'Bitola [mm2]  Material  R_dc [ohm/km] \n'
                   '800             Cobre         0.0224 \n'
                   '1000            Cobre         0.0177 \n'
                   '1200            Cobre         0.0151 \n'
                   '1400            Cobre         0.0129 \n'
                   '1600            Cobre         0.0113 \n'
                   '1800            Cobre         0.0101 \n'
                   '2000            Cobre         0.0090 \n'
                   '2500            Cobre         0.0072 \n'
                   '800          Alumínio         0.0367 \n'
                   '1000         Alumínio         0.0291 \n'
                   '1200         Alumínio         0.0247 \n'
                   '1400         Alumínio         0.0212 \n'
                   '1600         Alumínio         0.0186 \n'
                   '1800         Alumínio         0.0165 \n'
                   '2000         Alumínio         0.0149 \n'
                   '2500         Alumínio         0.0127 \n'))

material_c = input('\nMaterial do núcleo: \n'
                   '1 - Alumínio (Default) \n'
                   '2 - Cobre. \n')

if material_c == '2':
    r_dc_0 = r_dc * (1 + 0.00393 * (graus_c - 20))

else:
    r_dc_0 = r_dc * (1 + 0.00403 * (graus_c - 20))

# Efeito Pelicular

if material_c == "1":
    if r_0 != 0:
        ks = ((r_1 - r_0) / (r_1 + r_0)) * ((r_1 + 2 * r_0) / (r_1 + r_0)) ** 2
    else:
        ks = float(input('Digite o valor do fator ks para consideração acerca do efeito pelicular. \n'
                         'Valores de referência da norma IEC 60287-1-1:2006 para cabos de alumínio: \n'
                         'Tipo do condutor      ks      kp \n'
                         '  Sólido              1       1  \n'
                         '  Encordoado          1      0.8 \n'
                        '  Milliken            0.25    0.15 \n'))

else:
    if r_0 != 0:
        ks = ((r_1 - r_0) / (r_1 + r_0)) * ((r_1 + 2 * r_0) / (r_1 + r_0)) ** 2
    else:
        ks = float(input('Digite o valor do fator ks para consideração acerca do efeito pelicular. \n'
                         'Valores de referência da norma IEC 60287-1-1:2006 para cabos de cobre: \n'
                         'Tipo do condutor   Isolação  ks      kp \n'
                         '  Sólido             Todas    1      1 \n'
                         ' Encordoado        Extrudada  1      1 \n'
                         ' Milliken          Extrudada  0.35  0.20 \n'))

# Efeito de proximidade

if r_0 != 0:
    kp = 0.8

else:
    kp = float(input('Digite o valor do fator kp para consideração acerca do efeito de proximidade \n'
                     'Ver Tabela anterior \n'))

f_rps = float(input('Digite a frequência de cálculo de regime em [Hz]: '))

xs = math.sqrt(8 * math.pi * f_rps / r_dc_0 * 10 ** -7 * ks)

if xs > 0 and xs <= 2.8:
    ys = xs ** 4 / (192 + 0.8 * xs ** 4)

elif xs > 2.8 and xs <= 3.8:
    ys = -0.136 - 0.0177 * xs + 0.0563 * xs ** 2

else:
    ys = 0.354 * xs - 0.733

xp = math.sqrt(8 * math.pi * f_rps / r_dc_0 * 10 ** -7 * kp)

N = int(input('\nEntre com o número de cabos da instalação: '))
X = []
Y = []

for i in range(N):
    X.append(float(input(f'\nDigite a posição no eixo x do cabo {i+1} em metros: ')))
    Y.append(float(input(f'Digite a profundidade do cabo {i + 1} em metros (valor positivo): ')))

# X = [-0.35, -0.35, -0.35, 0.35, 0.35, 0.35] # Posição dos cabos no eixo x [m]
# Y = [1.25, 1.6, 1.95, 1.25, 1.6, 1.95] # Posição dos cabos no eixo y (valor positivo) [m]


yp = xp ** 4 / (192 + 0.8 * xp ** 4) * (2 * r_1 / math.sqrt((X[0] - X[1]) ** 2 + (Y[0] - Y[1]) ** 2)) ** 2 * (0.312 * (2 * r_1 / math.sqrt((X[0] - X[1]) ** 2 + (Y[0] - Y[1]) ** 2)) ** 2 +
                                                                                                              1.18 / (xp ** 4 / (192 + 0.8 * xp ** 4) + 0.27))

r_ac_iec = r_dc_0*(1 + ys + yp)

print(f'\nResistência AC em [ohm/km] conforme IEC em {graus_c} graus Celsius: {round(1000 * r_ac_iec, 5)} \n')

eps = 0.1  # Tolerancia entre R calculado por Wedepohl e R calculado pela IEC em %

w = 2 * math.pi * f_rps

u_c = float(input('Entre com a permeabilidade magnética relativa do núcleo: '))
u_c = u_c * (4 * math.pi * 10 ** -7)  # Permeabilidade magnética do núcleo [H/m]

m_c = cmath.sqrt(w * u_c / p_c * 1j)

Z1 = p_c * m_c / (2 * math.pi * r_1) * 1/(np.tanh(0.777 * m_c * r_1)) + 0.356 * p_c / (math.pi * r_1 ** 2)

n_iter = 100000  # Número máximo de iterações
iter = 0
p_c_orig = p_c

p = 1 / 100 * float(input(f'Digite a porcentagem de perturbação em p_c. \n'))

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
        p = 1 / 100 * float(input(f'Iteração {n_iter} atingida sem convergência, digite a nova porcentagem de '
                                  f'perturbação em p_c (iteração anterior = {100*p}%). \n'))


print(f'Resistência em ohm/km calculada em {iter} iterações: {round(1000 * np.real(Z1), 5)}. \n'
      f'Erro: {round(((np.real(Z1) - r_ac_iec) / r_ac_iec) * 100, 4)} % \n'
      f'Resistividade estimada para o condutor em [ohm.m]: {p_c}. \n')

e_sc_in = (10 ** -2) * float(input('Entre com a espessura da camada semicondutora entre núcleo e isolação [cm]: \n'))

e_isol_1 = (10 ** -2) * float(input('Entre com a espessura da primeira camada isolante [cm]: \n'))

e_sc_out = (10 ** -2) * float(input('Entre com a espessura da camada semicondutora entre isolação e blindagem [cm]: \n'))

r_2 = r_1 + e_sc_in + e_isol_1 + e_sc_out

a = r_1 + e_sc_in

b = a + e_isol_1

blindagem = input('\nBlindagem metálica de fios? \n'
                  '1 - Não (Default) \n'
                  '2 - Sim \n')

if blindagem == "2":
    n_fios = int(input('\nDigite o número de fios da blindagem metálica: '))
    d_f = (10 ** -2) * float(input('Digite a espessura (diâmetro) dos fios em [cm]: '))
    metodo_blindagem = int(input('Digite o método para modelagem da blindagem: \n'
                                 '1 - Calcula resistividade mantendo espessura da blindagem \n'
                                 '2 - Calcula raio para mesma área da blindagem (Default) \n'))

    if metodo_blindagem == 1:
        r_3 = r_2 + d_f / 2

    else:
        area_s = math.pi * n_fios * (d_f / 2) ** 2
        r_3 = math.sqrt(area_s / math.pi + r_2 ** 2)

else:
    e_s = (10 ** -2) * float(input('\nEntre com a espessura da blindagem metálica (considerada puramente tubular) em [cm]: '))
    r_3 = r_2 + e_s

e_isol_2 = (10 ** -2) * float(input('\nEntre com a espessura da segunda camada isolante em [cm]: '))

r_4 = r_3 + e_isol_2

# Valores de referência para resistividade, tirados da Brochura 531 do Cigré:
# Cobre - p_c = 1.7241E-8 ohm.m
# Alumínio - pc = 2.8264E-8 ohm.m

# p_c = 4.4969E-8 # Resistividade do núcleo [ohm.m]
# p_s = 1.68E-8 # Resistividade da blindagem [ohm.m]

p_s = float(input('\nEntre com a resistividade elétrica da blindagem em 20 graus Celsius: \n'
                  'Referência TB 531: \n'
                  'Cobre - 1.7241E-8 ohm.m \n'
                  'Alumínio - 2.8264E-8 ohm.m \n'))

graus_s = float(input('\nTemperatura a ser considerada na blindagem em graus Celsius: '))

material_s = input('\nMaterial da blindagem: \n'
                    '1 - Alumínio (Default) \n'
                    '2 - Cobre. \n')

if material_s == '2':
    p_s = p_s * (1 + 0.00393 * (graus_s - 20))

else:
    p_s = p_s * (1 + 0.00403 * (graus_s - 20))

# Correção da resistividade da blindagem em funcao do parâmetro metodo_blindagem
if metodo_blindagem == 1:
    p_s = p_s * (math.pi * (r_3 ** 2 - r_2 ** 2)) / (math.pi * n_fios * (d_f / 2) ** 2)

u_s = float(input('Entre com a permeabilidade magnética relativa da blindagem: '))
u_1 = float(input('Entre com a permeabilidade magnética relativa da primeira camada isolante: '))

if blindagem == "2":
    corrige_ind = input('\nDeseja corrigir a permeabilidade magnética da primeira camada isolante para \n'
                        'levar em conta o efeito da blindagem metálica de fios? \n'
                        '1 - Não (Default) \n'
                        '2 - Sim \n')

    if corrige_ind == "2":
        length_of_lay = (10 ** -2) * float(input('Digite o comprimento em [cm] necessário para a blindagem '
                                                 'completar uma volta em torno da primeira camada isolante: '))
        u_1 = u_1*(1 + (2 * (1 / length_of_lay) ** 2 * math.pi ** 2 * (r_2 ** 2 - r_1 ** 2)) / math.log(r_2 / r_1))
        print(f'Permabilidade magnética relativa após a correção: {u_1} \n')

u_2 = float(input('Entre com a permeabilidade magnética relativa da segunda camada isolante: '))
e_1 = float(input('Entre com a permissividade elétrica relativa da primeira camada isolante: '))

corrige_e_1 = input('Deseja corrigir a permissividade elétrica da primeira camada isolante para \n'
                    'levar em conta o efeito das camadas semicondutoras? \n'
                    '1 - Não (Default) \n'
                    '2 - Sim \n')

if corrige_e_1 == "2":
    e_1 = e_1 * math.log(r_2 / r_1) / math.log(b / a)
    print(f'Permissividade elétrica relativa após a correção: {e_1} \n')

e_2 = float(input('Entre com a permissividade elétrica relativa da segunda camada isolante: '))
p_solo = float(input('Entre com a resistividade elétrica do solo [ohm.m]: '))
u_solo = float(input('Entre com a permeabilidade magnética relativa do solo: '))

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

f = float(input('\nFrequência de cálculo em [Hz]: '))

w = 2 * math.pi * f

m_c = cmath.sqrt(w * u_c / p_c * 1j)

print('\nConsiderações sobre o cálculo de Z1 (Impedância interna do núcleo)')
metodo_z1 = input('1 - Método Simplificado apresentado em Wedepohl.\n'
                  '2 - Método Completo (Funções de Bessel) de Ametani (Default). \n')

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
metodo_blindagem = input('1 - Método Simplificado apresentado em Wedepohl.\n'
                         'Bons resultados para a relação (r3 - r2) / (r3 + r2) < 1/8 (0.125). \n'
                         f'Relação calculada para o caso atual: {round((r_3 - r_2) / (r_3 + r_2), 4)}. \n'
                         '2 - Método Completo (Funções de Bessel) de Wedepohl (Default). \n')

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

X.extend(X)
Y.extend(Y)

# Permite métodos de cálculo para impedâncias do solo
# Wedepohl e Saad



print('\nConsiderações sobre o cálculo de Z7 (Impedâncias de terra)')
metodo_earth = input('1 - Método Simplificado apresentado em Wedepohl.\n'
                     'Bons resultados para a relação | m_solo * dist | < 0.25. \n'
                     f'Relação calculada para o caso atual: {math.sqrt((min(X) - max(X)) ** 2 + (min(Y) - max(Y)) ** 2)}.\n'
                     '2 - Método aproximado completo (Funções de Bessel) de Saad (Default). \n')

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
                Z_earth[i][k] = p_solo * m_solo ** 2 / (2 * math.pi) * (special.kv(0, m_solo * math.sqrt((X[i] - X[k])**2 + (Y[i] - Y[k])**2)) + 2 / (4 + m_solo ** 2 * (np.abs(X[i] - X[k])) ** 2) * cmath.exp(-2 * (Y[i] + Y[k]) * m_solo))

        Z_earth[k][i] = Z_earth[i][k]

# Montagem da matriz de impedancias serie

Z_serie = Z_int + Z_earth

# Montagem das matrizes relativas à admitância da LT

Y_shunt = np.zeros([len(X), len(X)], dtype=complex)

for i in range(int(len(X) / 2)):
    Y_shunt[i][i] = w * 2 * math.pi * e_1 / math.log(r_2 / r_1) * 1j

    Y_shunt[i][i + int(len(X) / 2)] = -Y_shunt[i][i]
    Y_shunt[i + int(len(X) / 2), i] = Y_shunt[i][i + int(len(X) / 2)]
    Y_shunt[i + int(len(X) / 2), i + int(len(X) / 2)] = Y_shunt[i][i] + w * 2 * math.pi * e_2 / math.log(r_4 / r_3)

# Calculo dos parametros de sequencia da matriz
# Matriz R

# Construção da matriz R conforme transposições

if N == 3:
    transp = input('\nConsideração sobre transposição: \n'
                         '1 - Apenas núcleos (Default) \n'
                         '2 - Apenas blindagens. \n')

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
    transp = input('\nConsideração sobre transposição: \n'
                         '1 - Apenas núcleos com sentidos inversos nos circuitos (Default) \n'
                         '2 - Apenas núcleos com sentidos iguais nos circuitos \n'
                         '3 - Apenas blindagens com sentidos inversos nos circuitos \n'
                         '4 - Apenas blindagens com sentidos iguais nos circuitos \n')

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

print('Impedancia de sequencia positiva da LT em ohm/km')
print(f'Z_1 = {np.round(1000 * Z_012[1][1], 4)}')

print('Impedancia de sequencia negativa da LT em ohm/km')
print(f'Z_2 = {np.round(1000 * Z_012[2][2], 4)}')


print('Admitancia de Sequencia zero da LT em micro mho/km')
print(f'Y_0 = {np.round(10** 6 * 1000 * Y_012[0][0], 4)}')

print('Admitancia de Sequencia positiva da LT em micro mho/km')
print(f'Y_1 = {np.round(10** 6 * 1000 * Y_012[1][1], 4)}')

print('Admitancia de Sequencia negativa da LT em micro mho/km')
print(f'Y_2 = {np.round(10** 6 * 1000 * Y_012[2][2], 4)}')

matriz_z_seq = pd.DataFrame(np.round(Z_012 * 1000, 5))
matriz_y_seq = pd.DataFrame(np.round(Y_012 * 1000 * 10 ** 6, 5))

writer = pd.ExcelWriter('LSCable_1400mm_metodos_aux.xlsx', engine='xlsxwriter')

matriz_z_seq.to_excel(writer, 'Z1=' + metodo_z1 + 'Zs=' + metodo_blindagem + 'Z7=' + metodo_earth +
                      'Transp=' + transp, header=False, index=False, startcol=2, startrow=2)
matriz_y_seq.to_excel(writer, 'Z1=' + metodo_z1 + 'Zs=' + metodo_blindagem + 'Z7=' + metodo_earth +
                      'Transp=' + transp, header=False, index=False, startcol=2, startrow=4 + N)

writer.save()

saida_atp = input('\nDeseja imprimir os dados para entrada no ATP? \n'
                  '1 - Não (Default) \n'
                  '2 - Sim \n')

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

matriz_z = input('Deseja imprimir a matriz Zserie? \n'
                 '1 - Não (Default) \n'
                 '2 - Sim \n')
if matriz_z == "2":
    print('Matriz Z de fase em ohm/km: \n'
          f'{Z_serie * 1000}')

matriz_y = input('Deseja imprimir a matriz Yshunt? \n'
                 '1 - Não (Default) \n'
                 '2 - Sim \n')
if matriz_y == "2":
    print('Matriz Y de fase em mho/km: \n'
          f'{Y_shunt * 1000}')

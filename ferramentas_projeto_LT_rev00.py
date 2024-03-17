# Ferramenta para cálculo de projeto de LT (PECE)
# Matheus G. Ribeiro
# rev00 - Assume que todos os cabos de fase são idênticos e os para-raios também.
# Utiliza correções de Deri ao invés de Carson

import numpy as np
import math
from scipy import special
from scipy import linalg
import matplotlib.pyplot as plt


def campo_critico(m, temp, h, raio):
    """
    Parâmetros
    :param r: Raio do condutor [cm]
    :param m: fator de irregularidade, entre 0,82 e 0,92
    :param temp: temperatura ambiente [Graus Celsius]
    :param h: Altura da instalação em relação ao nível do mar [m]
    :return: Retorna gradiente crítico de corona [kV/cm]
    """
    # Cálculo da pressão atmosférica
    p = 760 - 0.086 * h

    # Cálculo da densidade relativa do ar
    delta = 0.386 * p / (273 + temp)
    print(f'Densidade relativa do ar calculada (del): {np.round(delta, 2)} p.u.')

    # Cálculo do gradiente crítico conforme Peek
    ec = 30 * m * delta * (1 + 0.301 / np.sqrt(delta * raio))

    return np.round(ec, 2)


def gradiente(matriz_p, tensao, n_sub, r_1_fase, s):
    """
    Calcula o gradiente máximo na superfície do condutor
    :param matriz_p: Matriz dos potenciais de Maxwell
    :param tensao: Tensão consdieradas nos cálculos (nominal ou superior)
    :param n_sub: Número de subcondutores do bundle
    :param r_1_fase: Raio externo do condutor
    :param s: espaçamento entre subcondutores
    :return: Gradiente máximo em kVp/cm
    """

    alfa = np.cos(120 * np.pi / 180) + 1j * np.sin(120 * np.pi / 180)
    e_0 = 8.85 * 10 ** -12  # F/m

    # 3 fases + 1 pr
    if len(matriz_p) == 4:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0])

    # 3 fases + 2 pr
    elif len(matriz_p) == 5:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0, 0])
    # 6 fases + 1 pr
    elif len(matriz_p) == 7:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0])

    else:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0, 0])

    q = np.linalg.inv(matriz_p).dot(v)
    q_max = np.abs(np.amax(q))

    # Fator de 10 ^-9 por cauda das unidades da matriz P (km/uF)
    e_med = 1 / (2 * np.pi * e_0) * q_max / (n_sub * r_1_fase) * 10 ** -9

    if n_sub == 1:
        raio = r_1_fase

    elif n_sub == 2:
        raio = s / 2

    elif n_sub == 3:
        raio = s / np.sqrt(3)

    elif n_sub == 4:
        raio = s / np.sqrt(2)

    e_max = e_med * (1 + (r_1_fase / raio) * (n_sub - 1))

    return e_max


def calcula_req(n, rmg, s):
    """
    Calcula raio equivalente de um feixe de subcondutores
    :param n: Número de subcondutores
    :param rmg: Raio médio geométrico do condutor [m]
    :param s: Espaçamento entre subcondutores [cm]
    :return: req - Raio equivalente do feixe [m]
    """

    # rmg = r_cond * np.exp(-1 / 4)
    s = s / 100

    if n == 1:
        raio = rmg

    elif n == 2:
        raio = s / 2

    elif n == 3:
        raio = s / np.sqrt(3)

    elif n == 4:
        raio = s / np.sqrt(2)

    req = (n * raio ** (n - 1) * rmg) ** (1 / n)

    return np.round(req, 6)


def calcula_reatancia(x_cond, y_cond, x_pr, y_pr, flag_pr, req_cond, req_pr, freq):
    """
    Função para calcular matriz de reatâncias dos condutores
    :param x_cond: Posição das fases no eixo x [m]
    :param y_cond: Posição das fases no eixo y [m]
    :param x_pr: Posição dos cabos para-raios no eixo x [m]
    :param y_pr: Posição dos cabos para-raios no eixo y [m]
    :param flag_pr: Flag para indicar se os para-raios devem ser considerados ou não (1 = aterrados, 0 = isolados)
    :param req_cond: Raio equivalente dos condutores de fase [m]
    :param req_pr: Raio equivalente dos para-raios [m]
    :param freq: Frequência do cálculo [Hz]
    :return: Matriz com reatâncias próprias e mútuas em ohm/km
    """

    w = 2 * np.pi * freq
    n_cond = len(x_cond)
    n_pr = len(x_pr) * flag_pr

    x = x_cond
    y = y_cond

    if n_pr != 0:
        x = np.append(x, x_pr)
        y = np.append(y, y_pr)

    matriz_x = np.zeros((n_cond + n_pr, n_cond + n_pr), dtype=complex)

    for i in range(n_cond + n_pr):
        for j in range(n_cond + n_pr):

            if i == j:
                if i <= n_cond - 1:
                    matriz_x[i][j] = 1j * w * 2 * (10 ** (-4)) * np.log((2 * y[i]) / req_cond)

                else:
                    matriz_x[i][j] = 1j * w * 2 * (10 ** (-4)) * np.log((2 * y[i]) / req_pr)

            else:
                matriz_x[i][j] = 1j * w * 2 * (10 ** (-4)) * np.log((np.sqrt((x[i] - x[j]) ** 2 + (y[i] + y[j]) ** 2)) /
                                                                     np.sqrt((x[i] - x[j]) ** 2 + (y[i] - y[j]) ** 2))

    return matriz_x


def calcula_rac(n_fase, n_sub, n_pr, r_0_fase, r_1_fase, rcc_fase, r_0_pr, r_1_pr, rcc_pr, flag_pr, freq):
    """
    Cálculo da resistência CA dos condutores da LT. Equações utilizadas são as mesmas do programa de cálculo
    de parâmetros de LT subterrânea.
    :param n_fase: Número de condutores de fases
    :param n_sub: Número de subcondutores do bundle
    :param n_pr: Número de cabos para-raios
    :param r_0_fase: Raio interno dos condutores de fase [cm]
    :param r_1_fase: Raio externo dos condutores de fase [cm]
    :param rcc_fase: Resistência CC na temperatura desejada dos condutores de fase [ohm/km]
    :param r_0_pr: Raio interno dos para-raios [cm]
    :param r_1_pr: Raio externo dos para-raios [cm]
    :param rcc_pr: Resistência CC na temperatura desejada dos para-raios [ohm/km]
    :param flag_pr: Flag para indicar se os para-raios devem ser considerados ou não
    :param freq: Frequência de cálculo [Hz]
    :return: Retorna matriz diagonal contendo as resistência em ohm/km
    """

    u = 4 * np.pi * 10 ** -7  # H/m
    w = 2 * np.pi * freq
    matriz_r = np.zeros((n_fase + n_pr * flag_pr, n_fase + n_pr * flag_pr), dtype=complex)

    # Cálculo da área dos condutores de fase [m²]
    s_fase = np.pi * ((r_1_fase / 100) ** 2 - (r_0_fase / 100) ** 2)
    s_pr = np.pi * ((r_1_pr / 100) ** 2 - (r_0_pr / 100) ** 2)

    # Cálculo da resistividade dos condutores (ohm . m)
    p_fase = rcc_fase / 1000 * s_fase
    p_pr = rcc_pr / 1000 * s_pr

    # Cálculo da profundidade de penetração
    m_fase = np.sqrt(w * u / p_fase * 1j)
    m_pr = np.sqrt(w * u / p_pr * 1j)

    # Cálculo das resistências em corrente alternada

    # Conversão de cm para m para os raios
    r_0_fase = r_0_fase / 100
    r_1_fase = r_1_fase / 100
    r_0_pr = r_0_pr / 100
    r_1_pr = r_1_pr / 100

    for i in range(len(matriz_r)):
        if i <= n_fase - 1:
            if r_0_fase == 0:
                matriz_r[i][i] = 1/n_sub * p_fase * m_fase / (2 * np.pi * r_1_fase) * \
                                 ((special.iv(0, m_fase * r_1_fase)) / (special.iv(1, m_fase * r_1_fase)))

            else:
                matriz_r[i][i] = 1/n_sub * p_fase * m_fase / (2 * np.pi * r_1_fase) * \
                ((special.iv(0, m_fase * r_1_fase) * special.kv(1, m_fase * r_0_fase) +
                  special.kv(0, m_fase * r_1_fase) * special.iv(1, m_fase * r_0_fase)) /
                 (special.iv(1, m_fase * r_1_fase) * special.kv(1, m_fase * r_0_fase) -
                  special.iv(1, m_fase * r_0_fase) * special.kv(1, m_fase * r_1_fase)))

        else:
            if r_0_pr == 0:
                matriz_r[i][i] = p_pr * m_pr / (2 * np.pi * r_1_pr) * \
                                 ((special.iv(0, m_pr * r_1_pr)) / (special.iv(1, m_pr * r_1_pr)))

            else:
                matriz_r[i][i] = p_pr * m_pr / (2 * np.pi * r_1_pr) * \
                ((special.iv(0, m_pr * r_1_pr) * special.kv(1, m_pr * r_0_pr) +
                  special.kv(0, m_pr * r_1_pr) * special.iv(1, m_pr * r_0_pr)) /
                 (special.iv(1, m_pr * r_1_pr) * special.kv(1, m_pr * r_0_pr) -
                  special.iv(1, m_pr * r_0_pr) * special.kv(1, m_pr * r_1_pr)))

    return np.real(matriz_r * 1000)


def calcula_deri(x_cond, y_cond, x_pr, y_pr, flag_pr, freq, p_solo):
    """
    Rotina para cálculo das correções de Deri (não Carson).
    :param x_cond: Posição no eixo x dos condutores de fase
    :param y_cond: Altura dos condutores de fase
    :param x_pr: Posição no eixo x dos para-raios
    :param y_pr: Altura dos cabos para-raios
    :param flag_pr: Flag para identificar se para-raios serão considerados ou não
    :param freq: Frequência de cálculo (60 Hz)
    :param p_solo: Resistividade do solo (ohm.m)
    :return: Retorna matriz com as correções de carson
    """
    u = 4 * np.pi * 10 ** -7
    w = 2 * np.pi * freq
    n_cond = len(x_cond)
    n_pr = len(x_pr) * flag_pr
    delta = 1 / np.sqrt(1j * w * u * (1 / p_solo))

    x = x_cond
    y = y_cond

    if n_pr != 0:
        x = np.append(x, x_pr)
        y = np.append(y, y_pr)

    matriz_deri = np.zeros((n_cond + n_pr, n_cond + n_pr), dtype=complex)

    for i in range(n_cond + n_pr):
        for j in range(n_cond + n_pr):

            if i == j:
                matriz_deri[i][j] = (1j * w * u) / (2 * np.pi) * np.log(1 + delta / y[i])

            else:
                matriz_deri[i][j] = (1j * w * u) / (4 * np.pi) * np.log(((y[i] + y[j] + 2 * delta) ** 2 +
                                                                         (x[i] - x[j]) ** 2) /
                                                                        ((y[i] + y[j]) ** 2 + (x[i] - x[j]) ** 2))

    return matriz_deri * 1000


def calcula_p_maxwell(x_cond, y_cond, r_1_cond, x_pr, y_pr, r_1_pr, flag_pr):
    """
    Cálculo dos coeficientes de Maxwell da LT
    :param x_cond: Posição dos condutores de fase no eixo x
    :param y_cond: Posição dos condutores de fase no eixo y
    :param r_1_cond: Raio externo dos condutores de fase
    :param x_pr: Posição dos para-raios no eixo x
    :param y_pr: Posição dos para-raios no eixo y
    :param r_1_pr: Raio externo dos para-raios
    :param flag_pr: Indicação para uso dos para-raios
    :return: Matriz dos potenciais
    """

    n_cond = len(x_cond)
    n_pr = len(x_pr) * flag_pr

    x = x_cond
    y = y_cond

    if n_pr != 0:
        x = np.append(x, x_pr)
        y = np.append(y, y_pr)

    matriz_p = np.zeros((n_cond + n_pr, n_cond + n_pr), dtype=complex)

    for i in range(n_cond + n_pr):
        for j in range(n_cond + n_pr):

            if i == j:
                if i <= n_cond - 1:
                    matriz_p[i][j] = 17.98 * np.log((2 * y[i]) / r_1_cond)
                else:
                    matriz_p[i][j] = 17.98 * np.log((2 * y[i]) / r_1_pr)

            else:
                matriz_p[i][j] = 17.98 * np.log((np.sqrt((x[i] - x[j]) ** 2 + (y[i] + y[j]) ** 2))/
                                                 np.sqrt((x[i] - x[j]) ** 2 + (y[i] - y[j]) ** 2))

    return matriz_p


def parametros_seq(matriz_z, matriz_p, x_fase, x_pr):
    """
    Calcula parâmetros de sequencia da LT
    :param matriz_z: Matriz de impedâncias série
    :param matriz_p: Matriz de potenciais de Maxwell
    :param x_fase: Posições no eixo x dos condutores de fase
    :param x_pr: Posições no eixo x dos para-raios
    :return: Matriz z_seq e p_seq
    """
    n_cond = len(x_fase)
    n_pr = len(x_pr)

    # Eliminação dos cabos para-raios

    # Matriz de impedancias
    z_cc = matriz_z[0:n_cond, 0:n_cond]
    z_cs = matriz_z[0:n_cond, n_cond:n_cond + n_pr]
    z_sc = matriz_z[n_cond:n_cond + n_pr, 0:n_cond]
    z_ss = matriz_z[n_cond:n_cond + n_pr, n_cond:n_cond + n_pr]

    z_eq = z_cc - (z_cs.dot(np.linalg.inv(z_ss))).dot(z_sc)

    # Matriz de potenciais
    p_cc = matriz_p[0:n_cond, 0:n_cond]
    p_cs = matriz_p[0:n_cond, n_cond:n_cond + n_pr]
    p_sc = matriz_p[n_cond:n_cond + n_pr, 0:n_cond]
    p_ss = matriz_p[n_cond:n_cond + n_pr, n_cond:n_cond + n_pr]

    p_eq = p_cc - (p_cs.dot(np.linalg.inv(p_ss))).dot(p_sc)

    alfa = np.cos(120 * np.pi / 180) + 1j * np.sin(120 * np.pi / 180)

    T = np.array([[1, 1, 1], [1, alfa ** 2, alfa], [1, alfa, alfa ** 2]])
    T_1 = np.linalg.inv(T)

    if n_cond == 3:
        R = np.array([[1, 0, 0],
                     [0, 1, 0],
                     [0, 0, 1]])

        z_eq = 1 / 3 * (z_eq + (np.linalg.inv(R).dot(z_eq)).dot(R) + R.dot(z_eq).dot(np.linalg.inv(R)))
        p_eq = 1 / 3 * (p_eq + (np.linalg.inv(R).dot(p_eq)).dot(R) + R.dot(p_eq).dot(np.linalg.inv(R)))

        z_012 = (T_1.dot(z_eq)).dot(T)
        p_012 = (T_1.dot(p_eq)).dot(T)

    elif n_cond == 6:
        T_aux_1 = np.concatenate((T, np.zeros([3, 3])), axis=1)
        T_aux_2 = np.concatenate((np.zeros([3, 3]), T), axis=1)
        T = np.concatenate((T_aux_1, T_aux_2), axis=0)

        T_aux_1 = np.concatenate((T_1, np.zeros([3, 3])), axis=1)
        T_aux_2 = np.concatenate((np.zeros([3, 3]), T_1), axis=1)
        T_1 = np.concatenate((T_aux_1, T_aux_2), axis=0)

        R = np.array([[0, 1, 0, 0, 0, 0],
                     [0, 0, 1, 0, 0, 0],
                     [1, 0, 0, 0, 0, 0],
                     [0, 0, 0, 0, 0, 1],
                     [0, 0, 0, 1, 0, 0],
                     [0, 0, 0, 0, 1, 0]])

        z_eq = 1 / 3 * (z_eq + (np.linalg.inv(R).dot(z_eq)).dot(R) + R.dot(z_eq).dot(np.linalg.inv(R)))
        p_eq = 1 / 3 * (p_eq + (np.linalg.inv(R).dot(p_eq)).dot(R) + R.dot(p_eq).dot(np.linalg.inv(R)))

        z_012 = (T_1.dot(z_eq)).dot(T)
        p_012 = (T_1.dot(p_eq)).dot(T)

    return z_012, p_012


def calcula_campo_eletrico(x_cond, y_cond, x_pr, y_pr, tensao, matriz_p, h_fixa):
    """
    Cálculo do campo elétrico em um plano ortogonal a LT
    :param x_cond: Posição no eixo x dos condutores de fase
    :param y_cond: Posição no eixo y dos condutores de fase
    :param x_pr: Posição no eixo x dos para-raios
    :param y_pr: Posição no eixo y dos para-raios
    :param tensao: Tensão de cálculo
    :param matriz_p: Matriz dos potenciais de Maxwell
    :param h_fixa: Altura fixa para cálculo dos campo (geralmente 1,5 m)
    :return: Plota o campo elétrico
    """

    n_cond = len(x_cond)
    n_pr = len(x_pr)
    alfa = np.cos(120 * np.pi / 180) + 1j * np.sin(120 * np.pi / 180)
    e_0 = 8.85 * 10 ** -12  # F/m

    matriz_p = matriz_p * 10 ** 9  # F/m

    # Cálculo da matriz Q

    # 3 fases + 1 pr
    if len(matriz_p) == 4:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0])

    # 3 fases + 2 pr
    elif len(matriz_p) == 5:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0, 0])
    # 6 fases + 1 pr
    elif len(matriz_p) == 7:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0])

    else:
        v = np.array([tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      tensao * np.sqrt(2) / np.sqrt(3),
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      tensao * np.sqrt(2) / np.sqrt(3) * alfa,
                      0, 0])

    q = np.linalg.inv(matriz_p).dot(v)

    # Cria vetor do espaço
    pos_x = []
    for i in np.arange(-60, 60, 0.5):
        pos_x.append(i)

    campo_x = np.zeros(len(pos_x), dtype=complex)
    campo_y = np.zeros(len(pos_x), dtype=complex)
    campo_r = np.zeros(len(pos_x), dtype=complex)

    for ponto in range(len(pos_x)):

        for i in range(len(q)):

            if i <= n_cond - 1:
                dist_i = np.sqrt((x_cond[i] - pos_x[ponto]) ** 2 + (y_cond[i] - h_fixa) ** 2)
                dist_im_i = np.sqrt((x_cond[i] - pos_x[ponto]) ** 2 + (-y_cond[i] - h_fixa) ** 2)
                teta_i = np.arcsin(np.abs(x_cond[i] - pos_x[ponto]) / dist_i)
                teta_im_i = np.arcsin(np.abs(x_cond[i] - pos_x[ponto]) / dist_im_i)

            else:
                dist_i = np.sqrt((x_pr[i - n_cond] - pos_x[ponto]) ** 2 + (y_pr[i - n_cond] - h_fixa) ** 2)
                dist_im_i = np.sqrt((x_pr[i - n_cond] - pos_x[ponto]) ** 2 + (-y_pr[i - n_cond] - h_fixa) ** 2)
                teta_i = np.arcsin(np.abs(x_pr[i - n_cond] - pos_x[ponto]) / dist_i)
                teta_im_i = np.arcsin(np.abs(x_pr[i - n_cond] - pos_x[ponto]) / dist_im_i)

            campo_x[ponto] = campo_x[ponto] + 1 / (2 * np.pi * e_0) * q[i] * \
                (np.sin(teta_i) / dist_i - np.sin(teta_im_i) / dist_im_i)

            campo_y[ponto] = campo_y[ponto] - 1 / (2 * np.pi * e_0) * q[i] * \
                (np.cos(teta_i) / dist_i + np.cos(teta_im_i) / dist_im_i)

        campo_r[ponto] = (campo_x[ponto] + campo_y[ponto])

    plt.plot(pos_x, np.abs(campo_r), label='Campo Elétrico')
    plt.legend(loc=0)
    plt.ylabel('kV/m')
    plt.xlabel('m')
    plt.grid(True)
    plt.title(f'Campo elétrico calculado a {h_fixa} m do solo')
    plt.show()
    #print(np.abs(campo_r))


def calcula_campo_magnetico(x_cond, y_cond, x_pr, y_pr, matriz_z, corrente, h_fixa):
    """
    Calcula campo magnético em um plano ortogonal a linha
    :param x_cond: Posição dos condutores de fase no eixo x [m]
    :param y_cond: Posição dos condutores de fase no eixo y [m]
    :param x_pr: Posição dos para-raios no eixo x [m]
    :param y_pr: Posição dos para-raios no eixo y [m]
    :param matriz_z: Matriz de impedâncias [ohm/km]
    :param corrente: Corrente utilizada para cálculo do campo [A]
    :param h_fixa: Altura de referência para cálculo do campo [m]
    :return: Plot o campo magnético
    """

    n_cond = len(x_cond)
    n_pr = len(x_pr)
    alfa = np.cos(120 * np.pi / 180) + 1j * np.sin(120 * np.pi / 180)
    u_0 = 4 * np.pi * 10 ** -7  # H/m

    # Cálculo de corrente nos cabos para-raios

    # Matriz de impedancias
    z_cc = matriz_z[0:n_cond, 0:n_cond]
    z_cs = matriz_z[0:n_cond, n_cond:n_cond + n_pr]
    z_sc = matriz_z[n_cond:n_cond + n_pr, 0:n_cond]
    z_ss = matriz_z[n_cond:n_cond + n_pr, n_cond:n_cond + n_pr]

    # Circuito Simples
    if n_cond == 3:
        i_c = np.array([corrente * np.sqrt(2) / np.sqrt(3),
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa])

    # Circuito duplo
    elif n_cond == 6:
        i_c = np.array([corrente * np.sqrt(2) / np.sqrt(3),
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa,
                      corrente * np.sqrt(2) / np.sqrt(3),
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa ** 2,
                      corrente * np.sqrt(2) / np.sqrt(3) * alfa])

    i_s = (np.linalg.inv(z_ss).dot(-z_sc)).dot(i_c)

    i_total = np.concatenate((i_c, i_s))

    # Cria vetor do espaço
    pos_x = []
    for i in np.arange(-60, 60, 0.5):
        pos_x.append(i)

    campo_x = np.zeros(len(pos_x), dtype=complex)
    campo_y = np.zeros(len(pos_x), dtype=complex)
    campo_r = np.zeros(len(pos_x), dtype=complex)

    for ponto in range(len(pos_x)):

        for i in range(len(i_total)):

            if i <= n_cond - 1:
                num_x = (h_fixa - y_cond[i])
                den_x = (pos_x[ponto] - x_cond[i]) ** 2 + (h_fixa - y_cond[i]) ** 2
                num_y = (pos_x[ponto] - x_cond[i])
                den_y = (pos_x[ponto] - x_cond[i]) ** 2 + (h_fixa - y_cond[i]) ** 2

            else:
                num_x = (h_fixa - y_pr[i - n_cond])
                den_x = (pos_x[ponto] - x_pr[i - n_cond]) ** 2 + (h_fixa - y_pr[i - n_cond]) ** 2
                num_y = (pos_x[ponto] - x_pr[i - n_cond])
                den_y = (pos_x[ponto] - x_pr[i - n_cond]) ** 2 + (h_fixa - y_pr[i - n_cond]) ** 2

            campo_x[ponto] = campo_x[ponto] + u_0 / (2 * np.pi) * i_total[i] * num_x / den_x

            campo_y[ponto] = campo_y[ponto] + u_0 / (2 * np.pi) * i_total[i] * num_y / den_y

        campo_r[ponto] = (campo_x[ponto] + campo_y[ponto])

    plt.plot(pos_x, 1 / np.sqrt(2) * np.abs(campo_r) * 10 ** 6, label='Campo Magnético')
    plt.legend(loc=0)
    plt.ylabel('uT')
    plt.xlabel('m')
    plt.grid(True)
    plt.title(f'Campo magnético calculado a {h_fixa} m do solo')
    plt.show()


# Dados da LT

# Nível de tensão [kV]
Vn = 500

# Frequência de cálculo [Hz]
freq = 60

# Disposição dos condutores na torre
pos_cond = {'x': [-10, 0, 10], 'y': [30, 30, 30]}
flecha_cond = 21
pos_cond = {'x': [-10, 0, 10],
            'y': [30 - 2 / 3 * flecha_cond, 30 - 2 / 3 * flecha_cond, 30 - 2 / 3 * flecha_cond]}

# Informações sobre os condutores
# Raio interno do condutor [cm] (se não tiver alma é zero)
r_0_cond = 0.927 / 2
# Raio externo do condutor [cm]
r_1_cond = 2.515 / 2
# Raio medio geométrico (rmg) do condutor [m]
rmg_cond = 0.01021
# Resistencia CC do condutor (ohm/km)
rcc_cond = 0.09
# Número de subcondutores
nsub = 4
# Espaçamento do bundle [cm]
s_cond = 45.7

# Disposição dos cabos para-raios na torre
pos_pr = {'x': [-8, 8], 'y': [37, 37]}
flecha_pr = 0.9 * flecha_cond
pos_pr = {'x': [-8, 8], 'y': [37 - 2 / 3 * flecha_pr, 37 - 2 / 3 * flecha_pr]}
# Indicação se os para-raios estão aterrados (1) ou isolados (0)
flag_pr = 1
# Informações sobre os para-raios
# Raio interno do condutor [cm]
r_0_pr = 0
# Raio externo do condutor [cm]
r_1_pr = 0.952 / 2
# Raio medio geométrico (rmg) do condutor [m]
rmg_pr = 0.0003
# Resistencia CC do condutor (ohm/km)
rcc_pr = 3.9075

# Resistividade do solo
p_solo = 1000

# Dados para cálculo do gradiente superficial
m = 0.92
temp = 40
h = 350

# Dados para cálculo do campo elétrico
h_campo = 1.5  # [m]
pu = 1  # [p.u.]

# Dados para cálculo do campo magnético
corrente = 4000  # [A]

print('---------------------------------------------------------------- \n'
      'Descrição dos dados de entrada: \n'
      f'Tensão nominal da linha: {Vn} kV \n'
      f'Frequência de operação/cálculos: {freq} Hz \n'
      f'Abscissas dos condutores de fase: {pos_cond["x"]} m \n'
      f'Oordenadas dos condutores de fase {pos_cond["y"]} m \n'
      f'Flecha dos condutores de fase: {flecha_cond} m \n'
      f'Raio interno dos condutores de fase: {r_0_cond} cm \n'
      f'Raio externo dos condutores de fase: {r_1_cond} cm \n'
      f'Raio médio geométrico (rmg) dos condutores de fase: {rmg_cond} m \n'
      f'Resistência em corrente contínua dos condutores de fase: {rcc_cond} ohm/km \n'
      f'Número de subcondutores do feixe dos condutores de fase: {nsub} \n'
      f'Espaçamento entre subcondutores do feixe: {s_cond} cm \n'
      f'Abscissas dos para-raios: {pos_pr["x"]} m \n'
      f'Oordenadas dos para-raios: {pos_pr["y"]} m \n'
      f'Flecha dos para-raios: {flecha_pr} m \n'
      f'Raio interno dos para-raios: {r_0_pr} cm \n'
      f'Raio externo dos para-raios: {r_1_pr} cm \n'
      f'Raio médio geométrico (rmg) dos para-raios: {rmg_pr} m \n'
      f'Resistência em corrente contínua dos para-raios: {rcc_pr} ohm/km \n'
      f'Resistividade do solo: {p_solo} ohm.m')

if flag_pr == 1:
    print('Para-raios aterrados')
else:
    print('Para-raios isolados')

print('----------------------------------------------------------------')

matriz_x = calcula_reatancia(pos_cond['x'], pos_cond['y'], pos_pr['x'], pos_pr['y'], 1,
                             calcula_req(nsub, rmg_cond, s_cond), calcula_req(1, rmg_pr, 0),
                             freq)

matriz_r = calcula_rac(len(pos_cond['x']), nsub, len(pos_pr['x']), r_0_cond, r_1_cond, rcc_cond, r_0_pr, r_1_pr,
                       rcc_pr, flag_pr, freq)

matriz_deri = calcula_deri(pos_cond['x'], pos_cond['y'], pos_pr['x'], pos_pr['y'], flag_pr, freq, p_solo)

matriz_z = matriz_r + matriz_x + matriz_deri

matriz_p = calcula_p_maxwell(pos_cond['x'], pos_cond['y'], calcula_req(nsub, r_1_cond / 100, s_cond),
                                           pos_pr['x'], pos_pr['y'], calcula_req(1, r_1_pr / 100, 0), flag_pr)

matriz_y = np.linalg.inv(matriz_p) * 1j * (2 * np.pi * freq)

parametros_seq(matriz_z, matriz_p, pos_cond['x'], pos_pr['x'])

z_012, p_012 = parametros_seq(matriz_z, matriz_p, pos_cond['x'], pos_pr['x'])
y_012 = 1j * 2 * np.pi * freq * np.linalg.inv(p_012)



print('\nCálculo dos parâmetros de sequência da LT: \n')

print(f'Impedância de sequência zero: {np.round(z_012[0][0], 4)} ohm/km')
print(f'Impedância de sequência positiva: {np.round(z_012[1][1], 4)} ohm/km')
print(f'Impedância de sequência negativa: {np.round(z_012[2][2], 4)} ohm/km')

print(f'Susceptância de sequência zero: {np.round(y_012[0][0], 4)} uS/km')
print(f'Susceptância de sequência positiva: {np.round(y_012[1][1], 4)} uS/km')
print(f'Susceptância de sequência negativa: {np.round(y_012[2][2], 4)} uS/km')

if len(pos_cond['x']) == 6:
    print(f'Impedância mútua de sequência zero: {np.round(z_012[0][3], 4)} ohm/km')

print('----------------------------------------------------------------')

print('\nCálculo do gradiente superficial \n')

print('Dados de entrada considerados: \n'
      f'Fator de superfície: {m} \n'
      f'Temperatura média: {temp} graus Celsius \n'
      f'Altitude: {h} m')

c_crit = campo_critico(m, temp, h, r_1_cond)
print(f'Gradiente crítico de corona: {c_crit} kVp/cm')

gradiente_max = gradiente(matriz_p, Vn, nsub, r_1_cond, s_cond)
print(f'Gradiente superficial calculado: {np.round(gradiente_max, 3)} kVp/cm')
print('----------------------------------------------------------------')

print('\nCálculo do campo elétrico \n')

print('Dados de entrada considerados: \n'
      f'Tensão de operação: {pu} p.u. ({pu * Vn}) kV \n'
      f'Altura de cálculo: {h_campo} m')
print('----------------------------------------------------------------')

calcula_campo_eletrico(pos_cond['x'], pos_cond['y'], pos_pr['x'], pos_pr['y'], pu * Vn, matriz_p, h_campo)

print('\nCálculo do campo magnético \n')

print('Dados de entrada considerados: \n'
      f'Corrente de operação: {corrente} Aef \n'
      f'Altura de cálculo: {h_campo} m')
print('----------------------------------------------------------------')


calcula_campo_magnetico(pos_cond['x'], pos_cond['y'], pos_pr['x'], pos_pr['y'], matriz_z, corrente, h_campo)

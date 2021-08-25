# Código para leitura dos dados estatísticos do aquivo .lis do ATP
# Leitura das frequências, médias, variâncias e desvios padrão

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

def histograma(vet_1, vet_2):
    """"
    Função que recebe vet_1 contendo o valor e vet_2 contendo a quantidade de vezes que o valor aparece
    Retorna vet_3, contendo somente dados do vet_1 replicado o número de vezes que consta em vet_2
    """
    vet_3 = []
    for i in range(len(vet_1)):
        cont = vet_2[i]
        for j in range(cont):
            vet_3.append(vet_1[i])

    return vet_3

n = 0
diretorio = 0

#atp_estat = {'Tipo': [], 'Variavel': [], 'Media': [], 'D_Padrao': [], 'Valores': [], 'Frequencia': []}
atp_estat = {'Tipo': [], 'Variavel': [], 'Media': [], 'D_Padrao': [], 'Valores': []}

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

linhas = 0

#with open(leitura, 'r', encoding='utf-8') as arquivo:
with open(leitura, 'r') as arquivo:

    texto = arquivo.readlines()

    for linha in texto:

        if '  SUMMARY   SUMMARY   SUMMARY   SUMMARY' in linha:

            # Pega tipo de variável e nome da variável

            if '0' in texto[linhas - 3][51:53]:

                atp_estat['Tipo'].append('Tensao FT')
                atp_estat['Variavel'].append(texto[linhas - 3][65:70])

            elif '-1' in texto[linhas - 3][51:53]:

                atp_estat['Tipo'].append('Tensão FF')
                atp_estat['Variavel'].append(texto[linhas - 3][65:70])

            else:

                atp_estat['Tipo'].append('C. Inrush')
                atp_estat['Variavel'].append(texto[linhas - 3][65:70] + ' - ' + texto[linhas - 3][71:76])

#            atp_estat['Tipo'].append(texto[linhas - 3][51:53])
#            atp_estat['Variavel'].append(texto[linhas - 3][65:100])

            j = 0
            val_aux = []
            freq_aux = []

            while 'Summary of preceding table follows:' not in texto[linhas + j + 6]:

                val_aux.append(float(texto[linhas + j + 6][19:30]))
                freq_aux.append(int(texto[linhas + j + 6][56:64]))
                j = j + 1

            atp_estat['Valores'].append(histograma(val_aux, freq_aux))
#            atp_estat['Frequencia'].append(freq_aux)
            atp_estat['Media'].append(float(texto[linhas + j + 7][39:55]))
            atp_estat['D_Padrao'].append(float(texto[linhas + j + 9][39:55]))

        linhas = linhas + 1

# Cria Dataframe

atp_estat = pd.DataFrame(atp_estat)

#sns.distplot(atp_estat[0]['Valores'])

# Plotagem dos gráficos

for i in range(len(atp_estat)):

    if i == 0 or i % 8 == 0:

        plt.subplots(2, 4)

#        plt.figure()
        plt.subplot(2, 4, 1)

    else:

        plt.subplot(2, 4, (i % 8) + 1)

    sns.histplot(atp_estat.iloc[i]['Valores'], kde=True, stat = 'density')
#    plt.plot(atp_estat.iloc[i]['Valores'], stats.norm(atp_estat.iloc[i]['Valores'].mean(), atp_estat.iloc[i]['Valores'].std()).pdf(atp_estat.iloc[i]['Valores']),
#             label='Normal Padrão')

    plt.plot(atp_estat.iloc[i]['Valores'], stats.norm(np.mean(atp_estat.iloc[i]['Valores']), np.std(atp_estat.iloc[i]['Valores'])).pdf(atp_estat.iloc[i]['Valores']),
             label='Normal Padrão', color = 'red')

    plt.legend(loc=0)
    plt.title(f'{atp_estat.iloc[i]["Tipo"]} {atp_estat.iloc[i]["Variavel"]}')

plt.show()

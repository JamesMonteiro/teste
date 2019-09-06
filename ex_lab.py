# importação de bibliotecas
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import sys
import xlrd

def valores_realizado(arquivo,sheet,mes,realizado):    #vare valores da aba 'realizado' e atribui a um novo DataFrame 12X2
    df = pd.read_excel(arquivo, sheet_name = sheet)
    df_copy = df.copy()
    
    colunas = df_copy.columns                               #captura todos o 'título' de cada coluna
    realizado_m = df_copy.drop(columns = ['Unnamed: 0'])    #retira a coluna selecionada
    
    tamanho = realizado_m.shape     #qtd de linha x colunas
    tam = tamanho[1]                #numero de colunas

    realizado_m = realizado_m.transpose()
 
    index_ = list(range(tam))                   #cria o index
    realizado_m.index = index_                  #atribui os valores do index ao dataframe
    realizado_m.columns = [mes,realizado]   #renomeia as colunas

    return realizado_m     

def valores_orcado(arquivo,sheet):                      #vare valores da aba 'orcado' 
    df2 = pd.read_excel(arquivo, sheet_name = sheet)    
    df2_copy = df2.copy()
    return df2_copy

def dif_valores(orcado,realizado,p_orcado,p_realizado,p_diferenca):       
    orcado_valores = orcado[p_orcado]                   #seleciona a coluna orcado
    realizado_valores = realizado[p_realizado]          #seleciona a coluna realizado
    dif_valores = orcado_valores - realizado_valores    #faz a diferença entre os valores das duas colunas

    diferenca = pd.Series(dif_valores, name=p_diferenca)       #atribui a diferença a uma série
    return diferenca

def uniao_das_matrizes(orcado,realizado,mes,p_orcado,p_realizado):
    
    meses = orcado[mes]                                                       #coleta a coluna dos meses
    orcado_valores = orcado[p_orcado]                                           #coleta os valores orçados
    realizado_valores = realizado[p_realizado]                                  #coleta os valores que foram realizados
    result = pd.concat([meses, orcado_valores, realizado_valores], axis=1)      #cria a matriz correspondente, com todos os valores desejáveis
    return result

def concatena(A,B):
    result = pd.concat([A,B], axis=1)                                           #concatena as duas matrizes, uma ao lado da outra, ou seja, acrescenta B em A                 
    return result

def ajuste(matriz,diferenca):
    copia = matriz.copy()
    t = copia.shape
    for i in range (t[0]):
        if (copia.loc[i,diferenca]< 0 ):
            copia.loc[i,diferenca] = 0
    return copia

def plot_grafico(matriz,mes,realizado,orcado,diff):       
    meses = matriz[mes]
    realizado_valores = matriz[realizado]
    orcado_valores = matriz[orcado]
    diferenca_valores = matriz[diff]

    x_pos = np.arange(len(realizado_valores))
    fig = plt.figure()
    first_bar = plt.bar(x_pos, realizado_valores, 0.8, color='g',label = 'Realizado')
    second_bar = plt.bar(x_pos, diferenca_valores, 0.8, color='y', label = 'Orçado',bottom=realizado_valores)
    
    # Definir posição e labels no eixo X
    plt.xlabel('Mês') 
    plt.ylabel('$') 
    plt.title('Orçamento') 
    plt.xticks(x_pos, (meses))
    plt.legend()
        
    #maximiza a tela
    mng = plt.get_current_fig_manager()
    mng.window.state('zoomed')
    
    #exibe o gráfico
    plt.show()
    
    #sava o gráfico
    fig.savefig('orçamento.png')


#declaração das sheets que serão utilizadas
arquivo = 'dados.xlsx'
nome_sheet_1 = 'realizado'
nome_sheet_2 = "orcado"

#parêmetros
p_realizado = 'realizado'
p_orcado = 'orcado'
p_diferenca = 'diff'
p_mes_realizado = 'mes'
p_mes_orcado ='mês'


#inicio

nova_matriz_realizado = valores_realizado(arquivo,nome_sheet_1,p_mes_realizado,p_realizado)                             #faz a leitura da 'tabela realizado' e transforma numa matriz 12x2
nova_matriz_orcado = valores_orcado(arquivo,nome_sheet_2)                                                               #faz a leitura da 'tabela orcado'

m_realizado_orcado = uniao_das_matrizes(nova_matriz_orcado,nova_matriz_realizado,p_mes_orcado,p_orcado,p_realizado)     #une as duas tabelas anteriores em: Mes, Orçado e Realizado

matriz_diferenca = dif_valores(nova_matriz_orcado,nova_matriz_realizado,p_orcado,p_realizado,p_diferenca)               #faz a diferenca 
matriz_final = concatena(m_realizado_orcado,matriz_diferenca)
#print(matriz_final)

matriz_graf = ajuste(matriz_final,p_diferenca)                                                                          #ajusta a matriz para ser exibida no gráfico
plot_grafico(matriz_graf,p_mes_orcado,p_realizado,p_mes_orcado,p_diferenca)                                             #constroi o gráfico, exibe e salva o mesmo

matriz_final.to_csv('dados_saida.csv', encoding='utf-8', index=False)                                                   #cria um arquivo .csv com os valores mes, valores orçados e realizados e a diferença

#Importa biblioteca pandas (para ler o arquivo xlsx), numpy para fazer os calculos com os dados extraídos e matplotlib para fazer os gráficos de comparação
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#caminho arquivo da planilha
planilha = "Dados Projeto Expansionista.xlsx"

#cria o dicionario dos retornos dos ativos
retornos_publicos = {}
retornos_privados = {}
for aba in ['IPCA+ 2030', 'IPCA+ 2045','IPCA+ 2028', 'SELIC 2026', 'Prefixado 2026', 'PETR27', 'AEGP23', 'AERI11', 'VALE38', 'PETR36']:
    if 'PETR' in aba or 'AEG' in aba or 'AERI' in aba or 'VALE' in aba:  # Verifica se a aba corresponde a PETR, AEG, AERI, VALE
        coluna = 'H'
    else:
        coluna = 'B'
    # Ler apenas a coluna 'PU' de cada Título, pulando as três primeiras linhas (que estão vazias) e removendo valores NaN
    dados_ativos = pd.read_excel(planilha, sheet_name=aba, usecols=coluna, skiprows=3).dropna()['PU']

    if coluna == 'H':
        # Adicionar ao dicionario os retornos diarios de titulos privados em forma de lista
        retornos_privados[aba] = dados_ativos.tolist()
    else:
        # Adicionar ao dicionario os retornos diarios de titulos privados em forma de lista
        retornos_publicos[aba] = dados_ativos.tolist()


#Calcular os Coeficientes de Variação para cada título publico e privado
def calcular_coeficiente(retornos):
    #Cria um dicionário para adicionar os coeficientes
    resultados_cv = {}

    #Para cada título, fazer os cálculos baseados em seus retornos diários. No caso retornos é uma variável, para diferenciar os públicos dos privados.
    for titulo in retornos:
        # Acessar a lista de retornos diários para o ativo 'PETR27'
        retorno = retornos[titulo]

        # Calculo das variacoes diarias
        variacao = np.diff(retorno) / retorno[:-1]

        #Calcular a média
        m = np.mean(variacao)

        #Calculo do desvido padrão
        dp = np.std(variacao)

        #Faz calculo do Coeficiente de Variação Pearson
        cv = (dp/m)

        #Verificar se o coeficiente é positivo, se sim, adiciona ao dicionario
        if cv > 0:
            resultados_cv[titulo] = {"M": m, "DP": dp, "CV": cv}

        #Retorna o dicionario com os dados do Coeficiente de Variação
    return resultados_cv


#cria dicionarios para os coeficientes positivos de titulos publicos e privados
cv_publicos = calcular_coeficiente(retornos_publicos)
cv_privados = calcular_coeficiente(retornos_privados)

#Monta dicionários de duas carteiras com base nos menores Coeficientes de Variação de Pearson, uma para públicos e outra para privados.
# Ordena os títulos com base no coeficiente de variação
publicos_ordenados = sorted(cv_publicos.items(), key=lambda x: x[1]['CV'])
privados_ordenados = sorted(cv_privados.items(), key=lambda x: x[1]['CV'])
#Seleciona os titulos com os dois menores coeficientes de variação
carteira_publicos = dict(publicos_ordenados[:2])
carteira_privados = dict(privados_ordenados[:2])


#Cálculo do indice sharpe e do peso de cada título selecionado na carteira privada.
def sharpe_privado():
    # Taxa de retorno livre de risco (CDI a.d.)
    rf = 0.000481440437333386
    
    # Inicializando variáveis para armazenar a melhor combinação de pesos e maior Sharpe
    global peso_privado
    maior_sharpe = float('-inf')

# Testando diferentes combinações de pesos
    for peso_petr36 in range(0, 101, 2):  # Variação de 2% em 2%
        peso_vale38 = 100 - peso_petr36  # O restante é o peso para VALE38

        # Convertendo os pesos para a escala de 0 a 1.0
        peso_petr36 = peso_petr36 / 100
        peso_vale38 = peso_vale38 / 100

        # Calculando o retorno médio da carteira
        m_carteira = (carteira_privados['PETR36']['M'] * peso_petr36) + (carteira_privados['VALE38']['M'] * peso_vale38)

        # Calculando o desvio padrão da carteira
        dp_carteira = ((carteira_privados['PETR36']['DP'] ** 2) * (peso_petr36 ** 2) + (carteira_privados['VALE38']['DP'] ** 2) * (peso_vale38 ** 2)) ** 0.5

        # Calculando o índice de Sharpe
        sharpe = (m_carteira - rf) / dp_carteira

        # Atualizando a melhor combinação de pesos para cada título e maior Sharpe
        if sharpe > maior_sharpe:
            peso_privado = {'PETR36': peso_petr36, 'VALE38': peso_vale38}
            maior_sharpe = sharpe
    return peso_privado


#Cálculo do indice sharpe e do peso de cada título selecionado na carteira pública.
def sharpe_publico():
    # Taxa de retorno livre de risco (CDI a.d.)
    rf = 0.000481440437333386
    
    # Inicializando variáveis para armazenar a melhor combinação de pesos e maior Sharpe
    global peso_publico
    maior_sharpe = float('-inf')

# Testando diferentes combinações de pesos
    for peso_selic in range(0, 101, 2):  # Variação de 2% em 2%
        peso_prefixado = 100 - peso_selic  # O restante é o peso para PREFIXADO

        # Convertendo os pesos para a escala de 0 a 1.0
        peso_selic = peso_selic / 100
        peso_prefixado = peso_prefixado / 100

        # Calculando o retorno médio da carteira
        m_carteira = (carteira_publicos['SELIC 2026']['M'] * peso_selic) + (carteira_publicos['Prefixado 2026']['M'] * peso_prefixado)

        # Calculando o desvio padrão da carteira
        dp_carteira = ((carteira_publicos['SELIC 2026']['DP'] ** 2) * (peso_selic ** 2) + (carteira_publicos['Prefixado 2026']['DP'] ** 2) * (peso_prefixado ** 2)) ** 0.5

        # Calculando o índice de Sharpe
        sharpe = (m_carteira - rf) / dp_carteira

        # Atualizando a melhor combinação de pesos e maior Sharpe
        if sharpe > maior_sharpe:
            peso_publico = {'SELIC 2026': peso_selic, 'Prefixado 2026': peso_prefixado}
            maior_sharpe = sharpe

    return peso_publico


#Funções para cálculo de retornos das carteiras e faz a criação dos gráficos de comparação
def r_carteira_publica(investimento):
    #Cálculo de retornos anteriores dos títulos da carteira no período de 31/12/2021 a 23/10/2023
    r_selic = round(((retornos_publicos['SELIC 2026'][-1] - retornos_publicos['SELIC 2026'][0]) / retornos_publicos['SELIC 2026'][0])*100, 2)
    r_prefixado = round(((retornos_publicos['Prefixado 2026'][-1] - retornos_publicos['Prefixado 2026'][0]) / retornos_publicos['Prefixado 2026'][0])*100, 2)

    #Chamando a função que faz os cálculos de sharpe e peso para calcular os retornos anteriores da carteira
    sharpe_publico()
    #Cálculo para retornos anteriores da carteira baseado no peso de cada título público
    r_total_anterior = peso_publico['SELIC 2026']*r_selic + peso_publico['Prefixado 2026']*r_prefixado/2
    r_esperado = 10.07
    r_investido = round(investimento*(r_esperado/100 + 1), 2)
    print(f"\nCARTEIRA PÚBLICA - Selic 2026 Prefixado 2026\n-------------------\nInvestimento: R${investimento}\nRetorno esperado: {r_esperado}%\nRetorno Anterior(12/2021 a 10/2023): {r_total_anterior}%\nMontante final: R${r_investido}\n")

    # Valores esperados dos índices de comparação
    cdi_esperado = 9.15
    selic_esperada = 10.07
    ipca_esperado = 3.90

    # Valores anteriores dos índices de comparação
    cdi_anterior = 24.42
    ibov_anterior = 8.53
    ipca_anterior = 8.38

    #definindo rótulos do gráfico
    valores = ['CARTEIRA PÚBLICA', 'SELIC', 'CDI', 'IPCA']
    valores_anteriores = ['CARTEIRA PÚBLICA', 'IBOV', 'CDI', 'IPCA']
    #Definindo valores do gráfico
    retornos_esperados = [r_esperado, selic_esperada, cdi_esperado, ipca_esperado]
    retornos_anteriores = [r_total_anterior, ibov_anterior, cdi_anterior, ipca_anterior]

    plt.subplot(1, 2, 1)
    #definindo a largura das barras do gráfico
    largura = 0.35

    #definindo os índices dos elementos para posicionar as barras no gráfico
    r1 = range(len(valores))
    #Configurando a barra do gráfico(posição), valores, cor, largura e legenda
    plt.bar(r1, retornos_esperados, color='yellow', width=largura, label='Retorno Esperado')
    #Título do gráfico
    plt.title('Comparação de Retornos Esperados - Carteira Pública')
    #Título do eixo x e peso da fonte
    plt.xlabel('Indicadores', fontweight='bold')
    #Título do eixo y e peso da fonte
    plt.ylabel('Retorno (%)', fontweight='bold')
    #Ajusta os rótulos de cada barra em suas devidas posiçoes
    plt.xticks(r1, valores)
    #Aciona as legendas
    plt.legend()
    
    #segundo gráfico(retornos anteriores)
    plt.subplot(1, 2, 2)
    plt.bar(r1, retornos_anteriores, color='red', width=largura, label='Retorno Anterior')
    plt.title('Comparação de Retornos Anteriores - Carteira Pública')
    plt.xlabel('Indicadores', fontweight='bold')
    plt.ylabel('Retorno (%)', fontweight='bold')
    plt.xticks(r1, valores_anteriores)
    plt.legend()

    #inicia os gráficos
    plt.show()

def r_carteira_privada(investimento):
    #Cálculo de retornos anteriores dos títulos da carteira no período de 31/12/2021 a 23/10/2023
    r_petr = round(((retornos_privados['PETR36'][-1] - retornos_privados['PETR36'][0]) / retornos_privados['PETR36'][0])*100, 2)
    r_vale = round(((retornos_privados['VALE38'][-1] - retornos_privados['VALE38'][0]) / retornos_privados['VALE38'][0])*100, 2)

    #Chamando a função que faz os cálculos de sharpe e peso para calcular os retornos anteriores da carteira
    sharpe_privado()

    #Cálculo para retornos anteriores e esperados da carteira baseado no peso de cada título privado
    r_total_anterior = peso_privado['PETR36']*r_petr + peso_privado['VALE38']*r_vale/2
    r_esperado = (round(1.0625*9.97, 2))
    r_investido = round(investimento*(r_esperado/100 + 1), 2)
    print(f"\nCARTEIRA PRIVADA - Vale38 Petr36\n-------------------\nInvestimento: R${investimento}\nRetorno esperado: {r_esperado}%\nRetorno Anterior(12/2021 a 10/2023): {r_total_anterior}%\nMontante final: R${r_investido}\n")

    # Valores esperados dos índices de comparação
    cdi_esperado = 9.97
    selic_esperada = 10.07
    ibov_esperada = 0
    ipca_esperado = 3.90

    # Valores anteriores dos índices de comparação
    cdi_anterior = 24.42
    selic_anterior = 0
    ibov_anterior = 8.53
    ipca_anterior = 8.38

    #definindo rótulos do gráfico
    valores = ['CARTEIRA PRIVADA', 'SELIC', 'CDI', 'IPCA']
    valores_anteriores = ['CARTEIRA PRIVADA', 'IBOV', 'CDI', 'IPCA']
    #Definindo valores do gráfico
    retornos_esperados = [r_esperado, selic_esperada, cdi_esperado, ipca_esperado]
    retornos_anteriores = [r_total_anterior, ibov_anterior, cdi_anterior, ipca_anterior]

    plt.subplot(1, 2, 1)
    #definindo a largura das barras do gráfico
    largura = 0.35

    #definindo os índices dos elementos para posicionar as barras no gráfico
    r1 = range(len(valores))
    #Configurando a barra do gráfico(posição), valores, cor, largura e legenda
    plt.bar(r1, retornos_esperados, color='green', width=largura, label='Retorno Esperado')
    #Título do gráfico
    plt.title('Comparação de Retornos Esperados - Carteira Privada')
    #Título do eixo x e peso da fonte
    plt.xlabel('Indicadores', fontweight='bold')
    #Título do eixo y e peso da fonte
    plt.ylabel('Retorno (%)', fontweight='bold')
    #Ajusta os rótulos de cada barra em suas devidas posiçoes
    plt.xticks(r1, valores)
    #Aciona as legendas
    plt.legend()
    
    #segundo gráfico (retornos anteriores)
    plt.subplot(1, 2, 2)
    plt.bar(r1, retornos_anteriores, color='blue', width=largura, label='Retorno Anterior')
    plt.title('Comparação de Retornos Anteriores - Carteira Privada')
    plt.xlabel('Indicadores', fontweight='bold')
    plt.ylabel('Retorno (%)', fontweight='bold')
    plt.xticks(r1, valores_anteriores)
    plt.legend()

    #inicia os gráficos
    plt.show()

investimento = float(input("Digite o valor do investimento: "))
r_carteira_publica(investimento)
r_carteira_privada(investimento)




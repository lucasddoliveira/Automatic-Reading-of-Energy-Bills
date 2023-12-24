from PyPDF2 import PdfReader
import pandas as pd
import os
import pygsheets
import locale
import tabula
import numpy as np
import math
import datetime
import time
import matplotlib.pyplot as plt

# get the start datetime
st = datetime.datetime.now()

base = pd.read_excel('Infos Adicionais.xlsx')

base['CDC'] = base['CDC'].astype(str)

#print(base.head())

# Configura o locale para o formato brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def write_to_gsheet(service_file_path, spreadsheet_id, sheet_name, data_df):
    """
    this function takes data_df and writes it under spreadsheet_id
    and sheet_name using your credentials under service_file_path
    """
    gc = pygsheets.authorize(service_file=service_file_path)
    sh = gc.open_by_key(spreadsheet_id)
    try:
        sh.add_worksheet(sheet_name)
    except:
        pass
    wks_write = sh.worksheet_by_title(sheet_name)
    wks_write.clear('A1',None,'*')
    wks_write.set_dataframe(data_df, (1,1), encoding='utf-8', fit=True)
    wks_write.frozen_rows = 1

def somar(str1,str2):
    
    # Remover pontos e vírgulas e substituir vírgulas por pontos
    num1 = float(str1.replace(".", "").replace(",", "."))
    num2 = float(str2.replace(".", "").replace(",", "."))

    # Somar os números
    total = num1 + num2
    total = locale.currency(total, grouping=True, symbol=False)
    #print(total)
    return total

tabelaA = {
    'ARQUIVO': [],
    'MATRÍCULA': [],#
    'REGIONAL': [],
    'CIDADE': [],
    'ENDEREÇO': [],
    'FORNECEDORA': [],
    'FINALIDADE': [],
    'SITUAÇÃO': [],
    'TARIFA ATUAL': [],
    'TARIFA SUGERIDA': [],
    'DEMANDA ATUAL PONTA (KW)': [],
    'DEMANDA ATUAL FORA PONTA (KW)': [],
    'DEMANDA SUGERIDA PONTA (KW)': [],
    'DEMANDA SUGERIDA FORA PONTA (KW)': [],
    'MAIOR ECONOMIA ANUAL COM MUDANÇA DE TARIFA E MUDANÇA DE DEMANDA (R$)': [],
    'MELHOR TÉCNICA PARA OBTER DEMANDA': []
}

tabelaB = {
    'ARQUIVO': [],
    'MATRÍCULA': [],#
    'REGIONAL': [],
    'CIDADE': [],
    'ENDEREÇO': [],
    'FORNECEDORA': [],
    'FINALIDADE': [],
    'SITUAÇÃO': [],
    'TARIFA ATUAL': [],
    'MÉDIA - CONSUMO (KWH)': [],
    'MÁXIMO - CONSUMO (KWH)': [],
    'MÍNIMO - CONSUMO (KWH)': [],
    'DESVIO PADRÃO - CONSUMO (KWH)': [],
    'QNTD. DE MESES COM CONSUMO NULO': [],
    'VALOR GASTO ANUALMENTE COM MESES DE CONSUMO NULO': []
}

historicoB = {
    'MATRÍCULA': [],#
    'REGIONAL': [],
    'CIDADE': [],
    'ENDEREÇO': [],
    'FORNECEDORA': [],
    'FINALIDADE': [],
    'SITUAÇÃO': [],
    'TARIFA ATUAL': [],
    'MÉDIA - CONSUMO (KWH)': [],
    'MÁXIMO - CONSUMO (KWH)': [],
    'MÍNIMO - CONSUMO (KWH)': [],
    'DESVIO PADRÃO - CONSUMO (KWH)': [],
    'QNTD. DE MESES COM CONSUMO NULO': [],
}

tarifas = {
    'AZUL A3 - CONSUMO KWH PONTA': [0.40786],
    'AZUL A3 - CONSUMO KWH FORA PONTA': [0.26927],
    'AZUL A3 - DEMANDA KW PONTA': [13.14],
    'AZUL A3 - DEMANDA KW FORA PONTA': [8.51],
    'AZUL A3 - ULTRAPASSAGEM DEMANDA KW PONTA': [26.28],
    'AZUL A3 - ULTRAPASSAGEM DEMANDA KW FORA PONTA': [17.02],
    'AZUL A4 - CONSUMO KWH PONTA': [0.43221],
    'AZUL A4 - CONSUMO KWH FORA PONTA': [0.29361],
    'AZUL A4 - DEMANDA KW PONTA': [44.11],
    'AZUL A4 - DEMANDA KW FORA PONTA': [22.76],
    'AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA': [88.22],
    'AZUL A4 - ULTRAPASSAGEM DEMANDA KW FORA PONTA': [45.52],

    'VERDE - CONSUMO KWH PONTA': [1.50179],
    'VERDE - CONSUMO KWH FORA PONTA': [0.29361],
    'VERDE - DEMANDA KW FORA PONTA': [22.76],
    'VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA': [45.52],

    'B3 - CONSUMO KWH': [0.60219]

}
def calcular_ultrapassagem(demanda_contratada, demanda_medida):
    demanda_ultrapassada = 0
    
    for i in range(len(demanda_medida)):
        if demanda_medida[i] > 1.05 * demanda_contratada:
            demanda_ultrapassada += demanda_medida[i]-demanda_contratada

    return demanda_ultrapassada

def demanda_otimizada(demanda_medida_p, demanda_medida_fp, tarifa):
    #print('cheguei na funcao')
    #print(demanda_medida_p[-1])
    if (tarifa == 'AZUL A3'):
        tarifa_demanda_p = tarifas['AZUL A3 - DEMANDA KW PONTA'][0]
        tarifa_demanda_fp = tarifas['AZUL A3 - DEMANDA KW FORA PONTA'][0]
        tarifa_demanda_up = tarifas['AZUL A3 - ULTRAPASSAGEM DEMANDA KW PONTA'][0]
        tarifa_demanda_ufp = tarifas['AZUL A3 - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]
    elif (tarifa == 'AZUL A4'):
        tarifa_demanda_p = tarifas['AZUL A4 - DEMANDA KW PONTA'][0]
        tarifa_demanda_fp = tarifas['AZUL A4 - DEMANDA KW FORA PONTA'][0]
        tarifa_demanda_up = tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA'][0]
        tarifa_demanda_ufp = tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]
    else:
        tarifa_demanda_p = 0
        tarifa_demanda_fp = tarifas['VERDE - DEMANDA KW FORA PONTA'][0]
        tarifa_demanda_up = 0
        tarifa_demanda_ufp = tarifas['VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]

    
    ### MELHOR DEMANDA NA PONTA
    demanda_otimizada = []

    valor_melhor_conta = 1000000000
    melhor_demanda_p = 0
    demandas_testadas_p = []
    valores_achados_p = []
    
    if(max(demanda_medida_fp)>max(demanda_medida_p)):
        teto = math.ceil(max(demanda_medida_fp))
    else: 
        teto = math.ceil(max(demanda_medida_p))



    for i in range(0, teto, 1):
            #print('cheguei na ')
            valor_na_conta = len(demanda_medida_p) * i * tarifa_demanda_p + calcular_ultrapassagem(i, demanda_medida_p) * tarifa_demanda_up

            if(valor_na_conta<valor_melhor_conta):
                valor_melhor_conta = valor_na_conta
                melhor_demanda_p = i 
            
            demandas_testadas_p.append(i)
            valores_achados_p.append(valor_na_conta)

    demanda_otimizada.append(melhor_demanda_p)
    
    ### MELHOR DEMANDA FORA PONTA

    valor_melhor_conta = 1000000000
    melhor_demanda_fp = 0
    valores_achados_fp = []

    

    for i in range(0, teto, 1):
            valor_na_conta = len(demanda_medida_fp) * i * tarifa_demanda_fp + calcular_ultrapassagem(i, demanda_medida_fp) * tarifa_demanda_ufp
            
            if(valor_na_conta<valor_melhor_conta):
                valor_melhor_conta = valor_na_conta
                melhor_demanda_fp = i


            
            valores_achados_fp.append(valor_na_conta) 

    demanda_otimizada.append(melhor_demanda_fp)
    #print(demanda_medida_p)
    if(tarifa=='VERDE A4'):
        demanda_otimizada[0] = demanda_otimizada[1]

    
    plt.plot(demandas_testadas_p, valores_achados_p, linestyle='-',  label='Ponta')
    
    plt.plot(demandas_testadas_p, valores_achados_fp, linestyle='-', label='Fora Ponta')

    
    # Adding labels and title
    plt.xlabel('Demanda Contratada')
    plt.ylabel('Valor Final da Conta de Demanda')
    plt.title('Análise de Demanda Contratada')
    plt.legend()
    
    # Display the plot
    plt.grid(True)
    plt.show()
    
    
    return demanda_otimizada

#demanda_otimizada()
dfA = pd.DataFrame(tabelaA)
dfB = pd.DataFrame(tabelaB)

diretorio = 'C:\\Users\\55839\\Desktop\\ESTÁGIO - CAGEPA\\Extrair_Info_ENERGISA\\Faturas_ENERGISA'
#diretorio = 'C:\\Users\\55839\\Desktop\\ESTÁGIO - CAGEPA\\Extrair_Info_ENERGISA\\Faturas_ENERGISA\\TESTE'

j = 1
#and nome_arquivo == '0001382659072023.pdf' VERDE
#0001558890072023.pdf VERDE
#0001804830072023.pdf VERDE CARO
#0009980556072023.pdf VERDE
#0009981243072023.pdf VERDE CARO ESTRANHO

for nome_arquivo in os.listdir(diretorio):
    if (nome_arquivo.endswith('.pdf') or nome_arquivo.endswith('.PDF')) and nome_arquivo=='9998067.pdf':
        if j>0:
            
                j+=1

                nome_fatura = nome_arquivo
                nome_arquivo = diretorio +"\\"+ nome_arquivo
                #print("************  LENDO O ARQUIVO: " + nome_fatura)

                primeiro_indice = next((i for i, c in enumerate(nome_fatura) if c.isdigit() and c != '0'), None)
                posicao_substring = -2
                if nome_fatura.find('2023')>0:
                    posicao_substring = nome_fatura.find('2023')
                

                if primeiro_indice is not None:
                    matricula = nome_fatura[primeiro_indice:posicao_substring-2]
                    #print(matricula)
                ucs = base['CDC'].values
                #print(ucs)
                indice_encontrado1 = -1

                    # Iterar pela lista e procurar a substring
                for indice, elemento in enumerate(ucs):
                    if matricula in elemento or matricula == elemento:
                        #print('SIM')
                        indice_encontrado1 = indice
                        break  # Parar assim que encontrar o primeiro índice

                    #print(indice)
                if indice_encontrado1!=-1:
                    indice_encontrado1 = int(indice_encontrado1)
                    resultado = base.iloc[indice_encontrado1]
                    #print(resultado)
                    matricula = resultado['CDC']
                    fornec1 = resultado['Fornecedora']
                    situacao1 = resultado['Situação']
                    regional1 = resultado['Unid. Negócio']
                    finalidade1 = resultado['Finalidade']
                    cidade1 = resultado['Cidade']
                    endereco1 = resultado['Endereço EnergiaWEB']
                else:
                    try:
                        matricula = nome_fatura.split('_')[1]
                    except:
                        matricula = 'ERRO'
                    for indice, elemento in enumerate(ucs):
                        if matricula in elemento or matricula == elemento:
                            #print('SIM')
                            indice_encontrado1 = indice
                            break  # Parar assim que encontrar o primeiro índice

                if indice_encontrado1!=-1:
                    indice_encontrado1 = int(indice_encontrado1)
                    resultado = base.iloc[indice_encontrado1]
                    #print(resultado)
                    matricula = resultado['CDC']
                    fornec1 = resultado['Fornecedora']
                    situacao1 = resultado['Situação']
                    regional1 = resultado['Unid. Negócio']
                    finalidade1 = resultado['Finalidade']
                    cidade1 = resultado['Cidade']
                    endereco1 = resultado['Endereço EnergiaWEB']        
                else:
                    matricula = '-'
                    fornec1 = '-'
                    situacao1 = '-'
                    regional1 = '-'
                    finalidade1 = '-'
                    cidade1 = '-'
                    endereco1 = '-'

                dfs = tabula.read_pdf(nome_arquivo, pages='2')
                
                try:
                    try:
                        dfs__ = tabula.read_pdf(nome_arquivo, pages='2', guess = False)
                        dfs_ = dfs__[0].iloc[:, 1:6]
                        dfs_.columns = ['Mês','Consumo Faturado Ponta', 'Demanda Medida Ponta', 'Consumo Faturado Fora Ponta', 'Demanda Medida Fora Ponta']
                        
                        
                        #print(dfs__[0])
                        #print(dfs_)

                        padrao_mes_ano = r'\b[A-Z]{3}/\d{2}\b'
                        #dfs_ = dfs_.dropna()
                        #print(dfs_)
                        fatia = dfs_.iloc[:, 0:1]
                        fatia = fatia.dropna()
                        
                        linhas_filtradas_ = fatia[fatia['Mês'].str.contains(padrao_mes_ano, regex=True)]
                        #print(linhas_filtradas_)
                        dfs_ = dfs_.iloc[linhas_filtradas_.index[0]:linhas_filtradas_.index[-1]+1, :]

                        dataframe = dfs_
                        dataframe = dataframe.reset_index()
                        dataframe.drop('index', inplace=True, axis=1)

                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace('.','')
                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace(',','.')
                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace('*','')
                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace(' ','')
                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].astype(float)
                        
                        #print('AAA')
                    except Exception as error:
                    #print(dfs_)
                        
                        #print('BBB')
                        dataframe = dfs[0].iloc[:11, 5:10]
                        
                        #print(dataframe)
                        first_row = dataframe.columns

                        dataframe.loc[-1] = first_row
                        dataframe.index = dataframe.index + 1  # shifting index
                        dataframe = dataframe.sort_index()  # sorting by index
                        
                        
                        #print(dataframe)
                        dataframe.columns = ['Mês','Consumo Faturado Ponta', 'Demanda Medida Ponta', 'Consumo Faturado Fora Ponta', 'Demanda Medida Fora Ponta']
                        
                    
                        #dfs2 = tabula.read_pdf(nome_arquivo, pages='2')
                        #print(dfs2)

                        
                    reader = PdfReader(nome_arquivo)
                    
                    page = reader.pages[1]
                    
                    # extracting text from page
                    text = page.extract_text()
                    #print(text2)
                    lines = text.split('\n')
                    #print(lines)
                    
                    contratadas = lines[-2].split(' ')
                    contratadas2 = lines[-1].split(' ')
                    
                    valores_numericos = [float(item) for item in contratadas if item.isnumeric()]
                    valores_numericos2 = [float(item) for item in contratadas2 if item.isnumeric()]
                    
                    if valores_numericos2:
                        valores_numericos = valores_numericos2

                    #dataframe = dfs_

                

                    #print(dataframe)
                    #print('LUPA')
                    
                    dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace('.','')
                    dataframe['Demanda Medida Ponta'] = dataframe['Demanda Medida Ponta'].str.replace('.','')
                    dataframe['Consumo Faturado Fora Ponta'] = dataframe['Consumo Faturado Fora Ponta'].str.replace('.','')
                    dataframe['Demanda Medida Fora Ponta'] = dataframe['Demanda Medida Fora Ponta'].str.replace('.','')

                    dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace(',','.')
                    dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace('*','')
                    dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].str.replace(' ','')
                    dataframe['Demanda Medida Ponta'] = dataframe['Demanda Medida Ponta'].str.replace(',','.')
                    dataframe['Consumo Faturado Fora Ponta'] = dataframe['Consumo Faturado Fora Ponta'].str.replace(',','.')
                    dataframe['Demanda Medida Fora Ponta'] = dataframe['Demanda Medida Fora Ponta'].str.replace(',','.')
                    #print(dataframe)
                    #print(valores_numericos)
                    #print('LUPA')
                    u=1
                    economiamax = []
                    registro = []
                    for z in range(1,5):    

                        if len(valores_numericos) == 2:
                            if u==1:
                                contratada_ponta = valores_numericos[0]
                                contratada_fora_ponta = valores_numericos[1]

                                contr_pont_orig = contratada_ponta
                                contr_fpont_orig = contratada_fora_ponta
                            else:
                                #print('---------------RETROALIMENTAÇÃO DE DEMANDA IDEAL---------------')
                                dataframe['Demanda Contratada Ponta'] = demanda_ideal_ponta
                                dataframe['Demanda Contratada Fora Ponta'] = demanda_ideal_fora_ponta
                                contratada_fora_ponta = demanda_ideal_fora_ponta
                                contratada_ponta = demanda_ideal_ponta
                            tarifacao = 'AZUL A4'
                        else:
                            if u==1:
                                contratada_fora_ponta = valores_numericos[0]
                                contratada_ponta = valores_numericos[0]

                                contr_pont_orig = contratada_fora_ponta
                                contr_fpont_orig = contratada_fora_ponta
                            else:
                                #print('---------------RETROALIMENTAÇÃO DE DEMANDA IDEAL---------------')
                                
                                if demanda_ideal_ponta>demanda_ideal_fora_ponta:
                                    dataframe['Demanda Contratada Ponta'] = demanda_ideal_ponta
                                    dataframe['Demanda Contratada Fora Ponta'] = demanda_ideal_ponta
                                    contratada_fora_ponta = demanda_ideal_ponta
                                    contratada_ponta = demanda_ideal_ponta
                                    
                                else:
                                    dataframe['Demanda Contratada Ponta'] = demanda_ideal_fora_ponta
                                    dataframe['Demanda Contratada Fora Ponta'] = demanda_ideal_fora_ponta
                                    contratada_fora_ponta = demanda_ideal_fora_ponta
                                    contratada_ponta = demanda_ideal_fora_ponta
                                    

                            tarifacao = 'VERDE A4'
                        
                        #print("@@@@  TARIFACAO: " + tarifacao)

                        if u==1:
                            dataframe.insert(5, 'Demanda Contratada Ponta', contratada_ponta)
                            dataframe.insert(6, 'Demanda Contratada Fora Ponta', contratada_fora_ponta)
                        #print(dataframe['Demanda Contratada Ponta'])
                        #print(dataframe['Demanda Contratada Fora Ponta'])
                        dataframe = dataframe.replace(to_replace='.*Unnamed.*', value=np.nan, regex=True)
                        
                        #print(dataframe)
                        dataframe = dataframe.replace('', np.nan)
                        dataframe['Consumo Faturado Ponta'] = dataframe['Consumo Faturado Ponta'].astype(float)
                        dataframe['Demanda Medida Ponta'] = dataframe['Demanda Medida Ponta'].astype(float)
                        dataframe['Consumo Faturado Fora Ponta'] = dataframe['Consumo Faturado Fora Ponta'].astype(float)
                        dataframe['Demanda Medida Fora Ponta'] = dataframe['Demanda Medida Fora Ponta'].astype(float)
                        #print(dataframe)
                        #print('AA')

                        media_coluna_A = dataframe['Consumo Faturado Ponta'].mean()
                        media_coluna_B = dataframe['Demanda Medida Ponta'].mean()
                        media_coluna_C = dataframe['Consumo Faturado Fora Ponta'].mean()
                        media_coluna_D = dataframe['Demanda Medida Fora Ponta'].mean()
                        
                        # Substitua os valores NaN pela média da coluna correspondente
                        dataframe['Consumo Faturado Ponta'].fillna(media_coluna_A, inplace=True)
                        dataframe['Demanda Medida Ponta'].fillna(media_coluna_B, inplace=True)
                        dataframe['Consumo Faturado Fora Ponta'].fillna(media_coluna_C, inplace=True)
                        dataframe['Demanda Medida Fora Ponta'].fillna(media_coluna_D, inplace=True)
                        #print(dataframe['Demanda Medida Fora Ponta'])
                        
                        dataframe['Demanda Faturada Ponta'] = np.where(dataframe['Demanda Medida Ponta'] >dataframe['Demanda Contratada Ponta'],
                                                            dataframe['Demanda Medida Ponta'],
                                                            dataframe['Demanda Contratada Ponta'])
                        dataframe['Demanda Faturada Fora Ponta'] = np.where(dataframe['Demanda Medida Fora Ponta'] > dataframe['Demanda Contratada Fora Ponta'],
                                                            dataframe['Demanda Medida Fora Ponta'],
                                                            dataframe['Demanda Contratada Fora Ponta'])
                        
                        dataframe['Demanda Excedida Ponta'] = np.where(dataframe['Demanda Faturada Ponta']-dataframe['Demanda Contratada Ponta']*1.05 > 0,
                                                            dataframe['Demanda Faturada Ponta'] - dataframe['Demanda Contratada Ponta'],
                                                            0)
                        dataframe['Demanda Excedida Fora Ponta'] = np.where(dataframe['Demanda Faturada Fora Ponta']-dataframe['Demanda Contratada Fora Ponta']*1.05 > 0,
                                                            dataframe['Demanda Faturada Fora Ponta'] - dataframe['Demanda Contratada Fora Ponta'],
                                                            0)
                        

                        ## SIMULAÇÃO AZUL A4 -> AZUL A4
                        
                        #print(dataframe['Demanda Faturada Ponta'])
                        #print(dataframe['Demanda Faturada Fora Ponta'])
                        ''''
                        simulacaoAzulA3 = dataframe.copy()
                        
                        simulacaoAzulA3['Consumo Faturado Ponta'] = simulacaoAzulA3['Consumo Faturado Ponta']*tarifas['AZUL A3 - CONSUMO KWH PONTA'][0]
                        simulacaoAzulA3['Consumo Faturado Fora Ponta'] = simulacaoAzulA3['Consumo Faturado Fora Ponta']*tarifas['AZUL A3 - CONSUMO KWH FORA PONTA'][0]



                        simulacaoAzulA3['Demanda Faturada Ponta'] = simulacaoAzulA3['Demanda Faturada Ponta']*tarifas['AZUL A3 - DEMANDA KW PONTA'][0]
                        simulacaoAzulA3['Demanda Faturada Fora Ponta'] = simulacaoAzulA3['Demanda Faturada Fora Ponta']*tarifas['AZUL A3 - DEMANDA KW FORA PONTA'][0]

                        

                        simulacaoAzulA3['Demanda Excedida Ponta'] = simulacaoAzulA3['Demanda Excedida Ponta']*tarifas['AZUL A3 - ULTRAPASSAGEM DEMANDA KW PONTA'][0]
                        simulacaoAzulA3['Demanda Excedida Fora Ponta'] = simulacaoAzulA3['Demanda Excedida Fora Ponta']*tarifas['AZUL A3 - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]
                        #print(simulacaoAzul)
                        teste = simulacaoAzulA3['Consumo Faturado Ponta'][0]+simulacaoAzulA3['Consumo Faturado Fora Ponta'][0]+simulacaoAzulA3['Demanda Faturada Ponta'][0]+simulacaoAzulA3['Demanda Faturada Fora Ponta'][0]+simulacaoAzulA3['Demanda Excedida Ponta'][0]+simulacaoAzulA3['Demanda Excedida Fora Ponta'][0] 
                        teste = teste+((simulacaoAzulA3['Consumo Faturado Ponta'][0]+simulacaoAzulA3['Consumo Faturado Fora Ponta'][0])*0.0108)+((simulacaoAzulA3['Consumo Faturado Ponta'][0]+simulacaoAzulA3['Consumo Faturado Fora Ponta'][0])*0.0499)
                        print(teste)

                        '''
                        
                        simulacaoAzul = dataframe.copy()
                        
                        simulacaoAzul['Consumo Faturado Ponta'] = simulacaoAzul['Consumo Faturado Ponta']*tarifas['AZUL A4 - CONSUMO KWH PONTA'][0]
                        simulacaoAzul['Consumo Faturado Fora Ponta'] = simulacaoAzul['Consumo Faturado Fora Ponta']*tarifas['AZUL A4 - CONSUMO KWH FORA PONTA'][0]


                        

                        simulacaoAzul['Demanda Faturada Ponta'] = simulacaoAzul['Demanda Faturada Ponta']*tarifas['AZUL A4 - DEMANDA KW PONTA'][0]
                        simulacaoAzul['Demanda Faturada Fora Ponta'] = simulacaoAzul['Demanda Faturada Fora Ponta']*tarifas['AZUL A4 - DEMANDA KW FORA PONTA'][0]

                        

                        simulacaoAzul['Demanda Excedida Ponta'] = simulacaoAzul['Demanda Excedida Ponta']*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA'][0]
                        simulacaoAzul['Demanda Excedida Fora Ponta'] = simulacaoAzul['Demanda Excedida Fora Ponta']*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]
                        #print(simulacaoAzul)

                        

                        teste = simulacaoAzul['Consumo Faturado Ponta'][0]+simulacaoAzul['Consumo Faturado Fora Ponta'][0]+simulacaoAzul['Demanda Faturada Ponta'][0]+simulacaoAzul['Demanda Faturada Fora Ponta'][0]+simulacaoAzul['Demanda Excedida Ponta'][0]+simulacaoAzul['Demanda Excedida Fora Ponta'][0] 
                        #print(simulacaoAzul['Consumo Faturado Ponta'][0])
                        
                        consumo_azul_azul_total = simulacaoAzul['Consumo Faturado Ponta'].sum()+simulacaoAzul['Consumo Faturado Fora Ponta'].sum()+simulacaoAzul['Demanda Faturada Ponta'].sum()+simulacaoAzul['Demanda Faturada Fora Ponta'].sum()+simulacaoAzul['Demanda Excedida Ponta'].sum()+simulacaoAzul['Demanda Excedida Fora Ponta'].sum()
                        
                        #if tarifacao=='AZUL A4':
                        #    print("----  AZUL PARA AZUL: " + str(consumo_azul_azul_total))
                        
                        ## SIMULAÇÃO AZUL A4 -> VERDE A4
                        
                        simulacaoAzulVerde = dataframe.copy()

                        simulacaoAzulVerde['Consumo Faturado Ponta'] = simulacaoAzulVerde['Consumo Faturado Ponta']*tarifas['VERDE - CONSUMO KWH PONTA'][0]

                        simulacaoAzulVerde['Consumo Faturado Fora Ponta'] = simulacaoAzulVerde['Consumo Faturado Fora Ponta']*tarifas['VERDE - CONSUMO KWH FORA PONTA'][0]

                        #print(simulacaoAzulVerde)
                        simulacaoAzulVerde['Demanda Faturada Max'] = np.where(simulacaoAzulVerde['Demanda Faturada Ponta']>simulacaoAzulVerde['Demanda Faturada Fora Ponta'],
                                                            simulacaoAzulVerde['Demanda Faturada Ponta'],
                                                            simulacaoAzulVerde['Demanda Faturada Fora Ponta'])
                        
                        simulacaoAzulVerde['Demanda Excedida Max'] = np.where(simulacaoAzulVerde['Demanda Excedida Ponta']>simulacaoAzulVerde['Demanda Excedida Fora Ponta'],
                                                            simulacaoAzulVerde['Demanda Excedida Ponta'],
                                                            simulacaoAzulVerde['Demanda Excedida Fora Ponta'])
                        
                        
                        simulacaoAzulVerde['Demanda Faturada Max'] = simulacaoAzulVerde['Demanda Faturada Max'] * tarifas['VERDE - DEMANDA KW FORA PONTA'][0]
                        simulacaoAzulVerde['Demanda Excedida Max'] = simulacaoAzulVerde['Demanda Excedida Max'] * tarifas['VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]

                        #print(simulacaoAzulVerde)

                        consumo_azul_verde_total = simulacaoAzulVerde['Consumo Faturado Ponta'].sum()+simulacaoAzulVerde['Consumo Faturado Fora Ponta'].sum()+simulacaoAzulVerde['Demanda Faturada Max'].sum()+simulacaoAzulVerde['Demanda Excedida Max'].sum()
                        j = j + 1

                        #if tarifacao=='AZUL A4':
                        #    print("----  AZUL PARA VERDE: " + str(consumo_azul_verde_total))

                        
                        
                        ## SIMULAÇÃO AZUL A4 -> B3
                        
                        simulacaoB3 = dataframe.copy()

                        consumo_azul_b3_total = (simulacaoB3['Consumo Faturado Ponta'].sum()+simulacaoB3['Consumo Faturado Fora Ponta'].sum())* tarifas['B3 - CONSUMO KWH'][0]
                        if tarifacao=='AZUL A4':
                            #print("----  AZUL PARA B3: " + str(consumo_azul_b3_total))

                            if consumo_azul_azul_total-consumo_azul_verde_total>0:
                                #print("--- ECONOMIA ANUAL MÁXIMA DE AZUL PARA VERDE: "+str(consumo_azul_azul_total-consumo_azul_verde_total))
                                tarifa_sugerida = 'VERDE A4'
                                
                                if u==1:
                                    despesa_original = consumo_azul_azul_total
                                    economia1 = 0
                                else:
                                    economia1 = despesa_original-consumo_azul_verde_total

                                

                            if consumo_azul_azul_total-consumo_azul_b3_total>0:
                                #print("--- ECONOMIA ANUAL MÁXIMA DE AZUL PARA B3: "+str(consumo_azul_azul_total-consumo_azul_b3_total))
                                tarifa_sugerida = 'B3'
                                
                                if u==1:
                                    despesa_original = consumo_azul_azul_total
                                    economia1 = 0
                                else:
                                    economia1 = despesa_original-consumo_azul_b3_total

                            if consumo_azul_azul_total<consumo_azul_b3_total and consumo_azul_azul_total<consumo_azul_verde_total:
                                #print("--- TARIFA MAIS VANTAJOSA É A AZUL: "+str(consumo_azul_azul_total))
                                tarifa_sugerida = 'AZUL A4'
                                if u==1:
                                    despesa_original = consumo_azul_azul_total
                                    economia1 = 0
                                else:
                                    economia1 = despesa_original-consumo_azul_azul_total


                        ## SIMULAÇÃO VERDE -> VERDE

                        simulacaoVerde = dataframe.copy()
                        
                        simulacaoVerde['Consumo Faturado Ponta'] = simulacaoVerde['Consumo Faturado Ponta']*tarifas['VERDE - CONSUMO KWH PONTA'][0]
                        simulacaoVerde['Consumo Faturado Fora Ponta'] = simulacaoVerde['Consumo Faturado Fora Ponta']*tarifas['VERDE - CONSUMO KWH FORA PONTA'][0]

                        simulacaoVerde['Demanda Faturada Max'] = np.where(simulacaoVerde['Demanda Faturada Ponta']>simulacaoVerde['Demanda Faturada Fora Ponta'],
                                                simulacaoVerde['Demanda Faturada Ponta'],
                                                simulacaoVerde['Demanda Faturada Fora Ponta'])

                        simulacaoVerde['Demanda Excedida Fora Ponta'] = np.where(simulacaoVerde['Demanda Faturada Max']>contratada_fora_ponta*1.05,
                                                (simulacaoVerde['Demanda Faturada Max']-contratada_fora_ponta)*tarifas['VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0],
                                                0)
                        
                        simulacaoVerde['Demanda Faturada Max'] = simulacaoVerde['Demanda Faturada Max']*tarifas['VERDE - DEMANDA KW FORA PONTA'][0]   

                        consumo_verde_verde_total = simulacaoVerde['Consumo Faturado Ponta'].sum()+simulacaoVerde['Consumo Faturado Fora Ponta'].sum()+simulacaoVerde['Demanda Faturada Max'].sum()+simulacaoAzul['Demanda Excedida Fora Ponta'].sum()
                        #if tarifacao=='VERDE A4':
                        #    print("----  VERDE PARA VERDE: " + str(consumo_verde_verde_total))

                        ## SIMULAÇÃO VERDE A4 -> AZUL A4

                        simulacaoVerdeAzul = dataframe.copy()
                        
                        simulacaoVerdeAzul['Consumo Faturado Ponta'] = simulacaoVerdeAzul['Consumo Faturado Ponta']*tarifas['AZUL A4 - CONSUMO KWH PONTA'][0]
                        simulacaoVerdeAzul['Consumo Faturado Fora Ponta'] = simulacaoVerdeAzul['Consumo Faturado Fora Ponta']*tarifas['AZUL A4 - CONSUMO KWH FORA PONTA'][0]

                        simulacaoVerdeAzul['Demanda Faturada Ponta'] = simulacaoVerdeAzul['Demanda Faturada Ponta']*tarifas['AZUL A4 - DEMANDA KW PONTA'][0]
                        simulacaoVerdeAzul['Demanda Faturada Fora Ponta'] = simulacaoVerdeAzul['Demanda Faturada Fora Ponta']*tarifas['AZUL A4 - DEMANDA KW FORA PONTA'][0]



                        simulacaoVerdeAzul['Demanda Excedida Ponta'] =  simulacaoVerdeAzul['Demanda Excedida Ponta']*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA'][0]
                        simulacaoVerdeAzul['Demanda Excedida Fora Ponta'] =  simulacaoVerdeAzul['Demanda Excedida Fora Ponta']*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0]
                        
                        consumo_verde_azul_total = simulacaoVerdeAzul['Consumo Faturado Ponta'].sum()+simulacaoVerdeAzul['Consumo Faturado Fora Ponta'].sum()+simulacaoVerdeAzul['Demanda Faturada Ponta'].sum()+simulacaoVerdeAzul['Demanda Faturada Fora Ponta'].sum()+simulacaoVerdeAzul['Demanda Excedida Ponta'].sum()+simulacaoVerdeAzul['Demanda Excedida Fora Ponta'].sum()
                        #if tarifacao=='VERDE A4':
                        #    print("----  VERDE PARA AZUL: " + str(consumo_verde_azul_total))

                        ## SIMULAÇÃO VERDE A4 -> B3
                        
                        simulacaoVerdeB3 = dataframe.copy()

                        consumo__verde_b3_total = (simulacaoVerdeB3['Consumo Faturado Ponta'].sum()+simulacaoVerdeB3['Consumo Faturado Fora Ponta'].sum())* tarifas['B3 - CONSUMO KWH'][0]
                        if tarifacao=='VERDE A4':
                            #print("----  VERDE PARA B3: " + str(consumo__verde_b3_total))

                            if consumo_verde_verde_total-consumo_verde_azul_total>0:
                                #print("--- ECONOMIA ANUAL MÁXIMA DE VERDE PARA AZUL: "+str(consumo_verde_verde_total-consumo_verde_azul_total))
                                tarifa_sugerida = 'AZUL A4'

                                if u==1:
                                    despesa_original = consumo_verde_verde_total
                                    economia1 = 0
                                else:
                                    economia1 = +despesa_original-consumo_verde_azul_total
                                    
                            if consumo_verde_verde_total-consumo__verde_b3_total>0:
                                #print("--- ECONOMIA ANUAL MÁXIMA DE VERDE PARA B3: "+str(consumo_verde_verde_total-consumo__verde_b3_total))
                                tarifa_sugerida = 'B3'

                                if u==1:
                                    despesa_original = consumo_verde_verde_total
                                    economia1 = 0
                                else:
                                    economia1 = despesa_original-consumo__verde_b3_total
                            if consumo_verde_verde_total<consumo__verde_b3_total and consumo_verde_verde_total<consumo_verde_azul_total:
                                #print("--- TARIFA MAIS VANTAJOSA É A VERDE: "+str(consumo_verde_verde_total))
                                tarifa_sugerida = 'VERDE A4'
                                if u==1:
                                    despesa_original = consumo_verde_verde_total
                                    economia1 = 0
                                else:
                                    economia1 = despesa_original-consumo_verde_verde_total
                        
                        
                        ###########################################################################################################################
                        
                        # DEMANDA ######
                        
                        demanda = dataframe.copy()
                        
                        if tarifacao=='AZUL A4':
                            demanda['Demanda Não Utilizada Ponta'] = np.where(contratada_ponta-demanda['Demanda Medida Ponta']<=0,
                                                    0,
                                                    (contratada_ponta-demanda['Demanda Medida Ponta'])*tarifas['AZUL A4 - DEMANDA KW PONTA'][0])
                            
                            demanda['Demanda Não Utilizada Fora Ponta'] = np.where(contratada_fora_ponta-demanda['Demanda Medida Fora Ponta']<=0,
                                                    0,
                                                    (contratada_fora_ponta-demanda['Demanda Medida Fora Ponta'])*tarifas['AZUL A4 - DEMANDA KW FORA PONTA'][0])
                            
                            dataframe['Demanda Excedida Real Ponta'] = np.where(dataframe['Demanda Faturada Ponta']-dataframe['Demanda Contratada Ponta'] > 0,
                                                    (dataframe['Demanda Faturada Ponta'] - dataframe['Demanda Contratada Ponta'])*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA'][0],
                                                    0)
                            
                            dataframe['Demanda Excedida Real Fora Ponta'] = np.where(dataframe['Demanda Faturada Fora Ponta']-dataframe['Demanda Contratada Fora Ponta'] > 0,
                                                    (dataframe['Demanda Faturada Fora Ponta'] - dataframe['Demanda Contratada Fora Ponta'])*tarifas['AZUL A4 - ULTRAPASSAGEM DEMANDA KW PONTA'][0],
                                                    0)
                            
                            demanda_nao_utilizada_max = demanda['Demanda Não Utilizada Ponta'].sum()+ demanda['Demanda Não Utilizada Fora Ponta'].sum()
                            demanda_ultrapassada_max = dataframe['Demanda Excedida Real Ponta'].sum() + dataframe['Demanda Excedida Real Fora Ponta'].sum()

                        
                        elif tarifacao=='VERDE A4': 
                            demanda['Demanda Não Utilizada Ponta'] = np.where(contratada_ponta-demanda['Demanda Medida Ponta']<=0,
                                                    0,
                                                    (contratada_ponta-demanda['Demanda Medida Ponta'])*tarifas['VERDE - DEMANDA KW FORA PONTA'][0])
                            
                            demanda['Demanda Não Utilizada Fora Ponta'] = np.where(contratada_fora_ponta-demanda['Demanda Medida Fora Ponta']<=0,
                                                    0,
                                                    
                                                    (contratada_fora_ponta-demanda['Demanda Medida Fora Ponta'])*tarifas['VERDE - DEMANDA KW FORA PONTA'][0])
                            
                            dataframe['Demanda Excedida Real Ponta'] = np.where(dataframe['Demanda Faturada Ponta']-dataframe['Demanda Contratada Ponta'] > 0,
                                                                (dataframe['Demanda Faturada Ponta'] - dataframe['Demanda Contratada Ponta'])*tarifas['VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0],
                                                                0)
                            
                            dataframe['Demanda Excedida Real Fora Ponta'] = np.where(dataframe['Demanda Faturada Fora Ponta']-dataframe['Demanda Contratada Fora Ponta'] > 0,
                                                                (dataframe['Demanda Faturada Fora Ponta'] - dataframe['Demanda Contratada Fora Ponta'])*tarifas['VERDE - ULTRAPASSAGEM DEMANDA KW FORA PONTA'][0],
                                                                0)
                            demanda_nao_utilizada_max = demanda['Demanda Não Utilizada Ponta'].sum()+ demanda['Demanda Não Utilizada Fora Ponta'].sum()
                            demanda_ultrapassada_max = dataframe['Demanda Excedida Real Ponta'].sum() + dataframe['Demanda Excedida Real Fora Ponta'].sum()
                        
                        #print("----  DEMANDA NAO UTILIZADA PONTA: " + str(demanda['Demanda Não Utilizada Ponta'].sum()))
                        #print("----  DEMANDA NAO UTILIZADA FORA PONTA: " + str(demanda['Demanda Não Utilizada Fora Ponta'].sum()))
                        #print("----  DEMANDA ULTRAPASSADA PONTA: " + str(dataframe['Demanda Excedida Real Ponta'].sum()))
                        #print("----  DEMANDA ULTRAPASSADA FORA PONTA: " + str(dataframe['Demanda Excedida Real Fora Ponta'].sum()))
                        if u==1: 
                            
                            demandas_ideais = demanda_otimizada(dataframe['Demanda Medida Ponta'].tolist(), dataframe['Demanda Medida Fora Ponta'].tolist(), tarifacao)
                            
                            demanda_ideal_ponta = demandas_ideais[0]
                            demanda_ideal_fora_ponta = demandas_ideais[1]
                            #print(demanda_ideal_ponta)
                            #print(demanda_ideal_fora_ponta)
                        if(u==2):
                            
                            demandas_ideais = demanda_otimizada(dataframe['Demanda Medida Ponta'].tolist(), dataframe['Demanda Medida Fora Ponta'].tolist(), tarifacao)
                            
                            demanda_ideal_ponta = demandas_ideais[0]
                            demanda_ideal_fora_ponta = demandas_ideais[1]
                            #demanda_ideal_ponta = math.ceil(dataframe['Demanda Medida Ponta'].mean())
                            #demanda_ideal_fora_ponta = math.ceil(dataframe['Demanda Medida Fora Ponta'].mean())
                            tecnica = 'PONTO IDEAL'
                            #print(demanda_ideal_ponta)
                            #print(demanda_ideal_fora_ponta)
                        if(u==3):
                            
                            demandas_ideais = demanda_otimizada(dataframe['Demanda Medida Ponta'].tolist(), dataframe['Demanda Medida Fora Ponta'].tolist(), tarifacao)
                            
                            demanda_ideal_ponta = demandas_ideais[0]
                            demanda_ideal_fora_ponta = demandas_ideais[1]
                            #demanda_ideal_ponta = math.ceil(dataframe['Demanda Medida Ponta'].median())
                            #demanda_ideal_fora_ponta = math.ceil(dataframe['Demanda Medida Fora Ponta'].median())
                            #tecnica = 'VALOR MÉDIO'
                            #print(demanda_ideal_ponta)
                            #print(demanda_ideal_fora_ponta)
                        if(u==4):
                            
                            demandas_ideais = demanda_otimizada(dataframe['Demanda Medida Ponta'].tolist(), dataframe['Demanda Medida Fora Ponta'].tolist(), tarifacao)
                            
                            demanda_ideal_ponta = demandas_ideais[0]
                            demanda_ideal_fora_ponta = demandas_ideais[1]
                            #tecnica = 'VALOR MEDIANO'
                            #print(demanda_ideal_ponta)
                            #print(demanda_ideal_fora_ponta)
                        
                            

                        #print("----  DEMANDA CONTRATADA PONTA: " + str(contratada_ponta))
                        #print("----  DEMANDA CONTRATADA FORA PONTA: " + str(contratada_fora_ponta))

                        #print("----  DEMANDA IDEAL PONTA: " + str(demanda_ideal_ponta))
                        #print("----  DEMANDA IDEAL FORA PONTA: " + str(demanda_ideal_fora_ponta))
                        
                        
                        f = 0
                        if(u>1):
                            economiamax.append(economia1)
                            registro.append([nome_fatura,
                                matricula, 
                                regional1, 
                                cidade1, 
                                endereco1, 
                                fornec1,
                                finalidade1, 
                                situacao1,
                                tarifacao,
                                tarifa_sugerida,
                                contr_pont_orig,
                                contr_fpont_orig,
                                demanda_ideal_ponta,
                                demanda_ideal_fora_ponta,
                                economia1,
                                tecnica])
                        u += 1
                except Exception as error: ##B3
                    
                    #print('CHEGUEEI')
                    try:
                        try: 
                            
                            #print('AAA')
                            dfs = tabula.read_pdf(nome_arquivo, pages='2', guess = False)

                            a = tabula.read_pdf(nome_arquivo, pages='1', guess = False)
                            
                            conta = a[0].iloc[:, 1:2]
                            conta = conta.dropna()
                            conta.columns = ['Coluna']
                            padraoconta = r'\b\d{2}/\d{2}/\d{4} R\b'

                            filtro = conta[conta['Coluna'].str.contains(padraoconta, regex=True)]
                            #print(filtro)

                            for i in filtro['Coluna']:
                                real = i.split(' ')[-1]

                            

                            real = real.replace('*','0')
                            real = real.replace(' ','') 
                            real = real.replace('.','')
                            real = real.replace(',','.')
                            real = float(real)

                            


                            novo_df1 = dfs[0].iloc[:, 0:1]
                            novo_df1.columns = ['Consumo']
                            novo_df1 = novo_df1.dropna()
                            #print(novo_df)
                            padrao_mes_ano = r'\b[A-Z]{3}/\d{2}\b'

                            flag = False
                            linhas_filtradas1 = novo_df1[novo_df1['Consumo'].str.contains(padrao_mes_ano, regex=True)]
                            
                            #print(linhas_filtradas1)
                            if len(linhas_filtradas1)>2:
                                
                                
                                for linha in linhas_filtradas1['Consumo']:                                
                                    parte = linha.split(' ')
                                    flag = True
                                    resultado = dfs[0].iloc[linhas_filtradas1.index[0]:linhas_filtradas1.index[-1]+1, 1:2]
                                    if '/2' not in parte[-1]:
                                        flag = False
                                        break
                            
                            #f parte_central[-1].contains('/'):
                            #print(flag)
                            
                            if(flag == False):
                                
                                teste = dfs[0]
                                #novo = dfs[0].iloc[2:15, 6:7]
                                novo = dfs[0].iloc[1:14, 2:3]
                                item = novo.iloc[0, 0]
                                #print(novo)
                                resultado = teste.loc[(teste['Unnamed: 0'].str.len() == 6) & teste['Unnamed: 0'].str.contains('/'), ['Unnamed: 1']]
                                #print(teste)
                                
                                #print(resultado)
                                
                                if len(resultado) == 0:
                                    novo_df = dfs[0].iloc[:, 0:1]
                                    novo_df.columns = ['Consumo']
                                    novo_df = novo_df.dropna()
                                    
                                    
                                    padrao_mes_ano = r'\b[A-Z]{3}/\d{2}\b'
                                    
                                    linhas_filtradas = novo_df[novo_df['Consumo'].str.contains(padrao_mes_ano, regex=True)]

                                    #print('AAAAAA')
                                    #print(linhas_filtradas)

                                    faturas_consumo = []   
                                    meses = []   
                                    
                                    for linha in linhas_filtradas['Consumo']:
                                        partes = linha.split()
                                        
                                        

                                        if len(partes)==1:
                                            
                                            if '/2' in partes[0]: 
                                                fat = '0'
                                                mes = partes[0]
                                        else:        
                                            if len(partes)==2:
                                                fat = partes[1]
                                                mes = partes[0]
                                            else:
                                                #print(partes)
                                                fat = partes[-1]
                                                mes = partes[-2]
                                        
                                        if '/2' in partes[-1]:
                                            fat = '0'
                                            mes = partes[-1]
                                        
                                        
                                        
                                        if(partes[-1]=='*'):
                                            #print(partes)
                                            fat = partes[-2]
                                            #mes = partes[-3]
                                            
                                            
                                            
                                            if len(partes)==2:
                                                if '/2' in partes[0]:
                                                    
                                                    fat = '0'
                                                    mes = partes[-2]
                                            else:
                                                mes = partes[-3]
                                                if '/2' in partes[-2]:
                                                    fat = '0'
                                                    mes = partes[-2]
                                        
                                            
                                        
                                        faturas_consumo.append(fat)
                                        meses.append(mes)
                                        
                                        #resultado = pd.DataFrame(fat, columns=['Consumo'], index=index)
                                        #print(faturas_consumo)
                                        #print(resultado)
                                    #print(faturas_consumo)
                                    resultado = pd.DataFrame({'Consumo': faturas_consumo})
                                    
                                    # Exibe as linhas filtradas
                                    
                                    ##print('AAA')

                            
                                

                            dataframe = resultado
                            dataframe.columns = ['Consumo']

                            dataframe = dataframe.reset_index()
                            dataframe.drop('index', inplace=True, axis=1)

                            dataframe['Consumo'] = dataframe['Consumo'].str.replace('*','0')
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace(' ','') 
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace('.','')
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace(',','.')
                            dataframe['Consumo'] = dataframe['Consumo'].astype(float)
                            dataframe['Consumo'].fillna(0, inplace=True)

                            contagem = dataframe['Consumo'].value_counts().get(0, 0)

                            primeiro_item = dataframe.loc[0, 'Consumo']

                            if primeiro_item == 0:
                                desperdicio = contagem * real
                            else:
                                desperdicio = '-'

                        except Exception as error:
                            #print(error)

                            #print('CASO 2')
                            novo = dfs[0].iloc[2:15, 2:3]
                            #print(novo)

                            dataframe = novo

                            dataframe.columns = ['Consumo']

                            dataframe = dataframe.reset_index()
                            dataframe.drop('index', inplace=True, axis=1)

                            dataframe['Consumo'] = dataframe['Consumo'].str.replace('*','0')
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace(' ','') 
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace('.','')
                            dataframe['Consumo'] = dataframe['Consumo'].str.replace(',','.')
                            dataframe['Consumo'] = dataframe['Consumo'].astype(float)
                            dataframe['Consumo'].fillna(0, inplace=True)
                            #print(dataframe)

                            contagem = dataframe['Consumo'].value_counts().get(0, 0)
                            
                            primeiro_item = dataframe.loc[0, 'Consumo']

                            if primeiro_item == 0:
                                desperdicio = contagem * real
                            else:
                                desperdicio = '-'
                        
                    #print(dataframe)
                        primeiro_indice = next((i for i, c in enumerate(nome_fatura) if c.isdigit() and c != '0'), None)
                        posicao_substring = -3
                        if nome_fatura.find('2023')>0:
                            posicao_substring = nome_fatura.find('2023')
                        

                        if primeiro_indice is not None:
                            matriculaB = nome_fatura[primeiro_indice:posicao_substring-2]
                            #print(matriculaB)

                        indice_encontrado1 = -1
                
                            # Iterar pela lista e procurar a substring
                        for indice, elemento in enumerate(ucs):
                            #print(type(elemento))
                            if matriculaB == elemento:
                                #print('AA')
                                #print(type(elemento))
                                #print('AA')
                                #print('SIM')
                                indice_encontrado1 = indice
                                break  # Parar assim que encontrar o primeiro índice

                        #print(indice)
                        
                        if(indice_encontrado1!=-1):
                            indice_encontrado1 = int(indice_encontrado1)
                            resultado2 = base.iloc[indice_encontrado1]
                            #print(resultado2)
                            fornec = resultado2['Fornecedora']
                            situacao = resultado2['Situação']
                            regional = resultado2['Unid. Negócio']
                            finalidade = resultado2['Finalidade']
                            cidade = resultado2['Cidade']
                            endereco = resultado2['Endereço EnergiaWEB']
                        else:
                            fornec = '-'
                            situacao = '-'
                            regional = '-'
                            finalidade = '-'
                            cidade = '-'
                            endereco = '-'
                        #print('AAA')
                        #print(dfs)
                        #print(dfs[0].iloc[:11, 6:7])
                        
                        
            
                        #print('.... UNIDADE B3 ....')

                        num_rows = dataframe.shape[0]
                        #if(num_rows!=13):
                            #print('OPAAAAAAA: '+num_rows)

                        registroB3 = [nome_fatura, matriculaB, regional, cidade, endereco, fornec, finalidade, situacao, 'B3', dataframe['Consumo'].mean(), dataframe['Consumo'].max(), dataframe['Consumo'].min(), dataframe['Consumo'].std(), contagem, desperdicio]

                        #print(dataframe)
                        
                        dfB.loc[len(dfB)] = registroB3
                    
                        f = 1
                    except Exception as error:
                    
                        print("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAATENÇÃO!!!!! ERRO DE LEITURA NO ARQUIVO: " + nome_arquivo)
                        print(error)
                        
                        print("ATENÇÃO!!!!! ERRO DE LEITURA NO ARQUIVO: " + nome_arquivo)
                        print(error)
                        print('OI')

              
                    
                   

                if f==0:       
                    max_index = economiamax.index(max(economiamax))
                    #print(registro)
                    dfA.loc[len(dfA)] = registro[max_index]
        
#if len(dfA)>0:
    #write_to_gsheet("leitura-de-pdfs-7632167dd370.json","1FGBQ_srO0pyK2qwb-i5RvEM0FypAixXQN0ULmcE6fLk","Sheet16",dfA)
#if len(dfB)>0:
    #write_to_gsheet("leitura-de-pdfs-7632167dd370.json","1FGBQ_srO0pyK2qwb-i5RvEM0FypAixXQN0ULmcE6fLk","Sheet15",dfB)
et = datetime.datetime.now()
elapsed_time = et - st
print('Execution time:', elapsed_time, 'seconds')

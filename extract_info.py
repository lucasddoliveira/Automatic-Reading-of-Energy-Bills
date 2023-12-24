from PyPDF2 import PdfReader
import pandas as pd
import os
import pygsheets

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
    
    #print(wks_write.get_col(1).index(''))
    linha_start = wks_write.get_col(1).index('') + 1
    print(linha_start)
    #wks_write.clear('A3',None,'*')
    linha_final = linha_start + data_df.shape[0]
    print(linha_final)
    
    for lin in range(linha_start, linha_final):

        lista = [data_df.iloc[lin-linha_start].tolist()]
        wks_write.update_values('A'+str(lin), lista)
    """
    wks_write = sh.worksheet_by_title(sheet_name)
    wks_write.clear('A1',None,'*')
    wks_write.set_dataframe(data_df, (1,1), encoding='utf-8', fit=True)
    wks_write.frozen_rows = 1

    """

tabela = {
    'Unidade': [],
    'Contrato': [],
    'Data de Emissão': [],
    'UNID': [],
    'Quantidade': [],
    'Valor Unitário': [],
    'Valor Total': [],
    'Produto': []
}

df = pd.DataFrame(tabela)

diretorio = 'C:\\Users\\55839\\Desktop\\ESTÁGIO - CAGEPA\\Extrair_Info_PDF\\Faturas_do_Mes'

for nome_arquivo in os.listdir(diretorio):
    if nome_arquivo.endswith('.pdf'):
# creating a pdf reader object
        print(nome_arquivo)
        nome_arquivo = diretorio +"\\"+ nome_arquivo
        
        reader = PdfReader(nome_arquivo)
        
        # getting a specific page from the pdf file
        page = reader.pages[0]
        
        # extracting text from page
        text = page.extract_text()
        lines = text.split('\n')

        inicioData = lines[5].find(": ") + 2
        fimData = lines[5].find("20", inicioData)

        data = lines[5][inicioData:fimData+4]

        palavras = lines[74].split()
        palavrasDados = lines[78].split()

        inicio = lines[78].find(": ") + 2
        fim = lines[78].find(" -", inicio)

        # Extraindo a substring entre as posições encontradas
        unidade = lines[78][inicio:fim]

        unid = palavras[2][-3:]
        qtde = palavras[3]
        vlr_unit = palavras[4]
        valor_total = palavras[5]
        contrato = palavrasDados[-1]
        produto = lines[74].split(' ')[-3:]
        produto = ' '.join(produto)
        

        #produto = 

        novo_registro = [unidade, contrato, data, unid, qtde, vlr_unit, valor_total, produto]

        df.loc[len(df)] = novo_registro

        #print(data)
        #print(unid)
        #print(qtde)
        #print(vlr_unit)
        #print(valor_total)
        #print(contrato)
        #print(unidade)

        print(df)

write_to_gsheet("leitura-de-pdfs-7632167dd370.json","1pWpF9BxNhaQK_2v4BHaay_Mh1eBKpbPtN5NSW0RkSb0","Leitura",df)
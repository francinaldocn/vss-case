# Importando as bibliotecas necessárias
import os
import re
import pandas as pd
from rarfile import RarFile

def extrair_arquivo_rar(arquivo, extensao):
    """
    Esta função extrai os arquivos em um diretório especificado.
    
    Parâmetros:
    arquivo (str): O caminho para o arquivo compactado
    extensao (str): A extensão dos arquivos a serem extraídos
    """
    with RarFile(arquivo) as rf:
        file_names = rf.namelist()
        for file_name in file_names:
            if file_name.endswith(extensao):
                rf.extract(file_name)


def listar_arquivos(diretorio):
    """
    Esta função lista os arquivos em um diretório especificado,
    e salva o resultado em uma lista.
    
    Parâmetros:
    diretorio (str): O caminho para o diretório.
    
    Retorna:
    list: Uma lista com os nomes dos arquivos no diretório.
    """
    # Criar uma lista vazia para armazenar os nomes dos arquivos.
    lista_arquivos = []
    
    # Verificar se o diretório existe.
    if not os.path.isdir(diretorio):
        raise ValueError(f"O diretório {diretorio} não existe.")
    
    # Iterar sobre os arquivos e adicioná-los à lista.
    for arquivo in os.listdir(diretorio):
        lista_arquivos.append(arquivo)
    
    # Retornar a lista de arquivos.
    return lista_arquivos


def extrair_relatorio(arquivo_entrada, inicio, final, termo):
    """
    Funcao para extracao dos textos
    """
    # Definir variaveis
    texto_completo = ""
    with open(arquivo_entrada, "r") as arq:
        # Percorrer o arquivo linha por linha
        for linha in arq:
            # Verificar a linha atual
            if inicio in linha:
                # Se achar, então começar a extrair o subtexto
                subtexto = linha
                # Continuar percorrendo o arquivo
                for x in arq:
                    if final in x:
                        # Se achar, então finalizar a extração
                        subtexto += x
                        # Verificar se contém a string esperada
                        if termo in subtexto:
                            # Se achar, então salvar o resultado
                            texto_completo += subtexto
                        break
                    else:
                        # Se não achar, então continuar a extração
                        subtexto += x        
        
    return texto_completo

def salvar_relatorio(texto_extraido, arquivo_saida):
    """
    Funcao para salvar o resultado em um arquivo
    """
    with open(arquivo_saida, "w") as arq:
        arq.write(texto_extraido)


def convert_currency_format(value_str):
    """
    Funcao para converter o formato dos numeros
    """
    # Adiciona o sinal negativo caso seja "DB"
    if "DB" in value_str:
        value_str = value_str.replace("DB", "")
        value_str = '-'+value_str
    elif "CR" in value_str:
        value_str = value_str.replace("CR", "")
  
    # Remove a virgula para inteiro e adiciona para decimal
    value_tmp = value_str.replace(",", "")
    value_tmp = value_tmp.replace(".", ",")
    
    # Divide a string a partir do último caractere antes da vírgula
    value_int, value_dec = value_tmp.split(",")
    
    # Insere um ponto a cada três caracteres a partir do final da string
    if value_int[:1] == '-':
        value_int = value_int[1:len(value_int)]
        value_int = ".".join(re.findall(r'.{1,3}', value_int[::-1]))[::-1]
        value_int = '-'+value_int
    else:
        value_int = ".".join(re.findall(r'.{1,3}', value_int[::-1]))[::-1]
    
    # Junta as duas partes da string
    value = value_int + "," + value_dec
    
    return value


def convert_vss110_to_csv(arquivo_entrada, arquivo_saida):
    """
    Funcao para converter o relatorio VSS-110 em CSV
    """
    # Variaveis
    corpo = False
    dic_dados_base = {'REPORT DATE':'', 'SETTLEMENT CURRENCY':'', 'SOURCE':'', 'COUNT':'', 
                      'CREDIT AMOUNT':'', 'DEBIT AMOUNT':'', 'TOTAL AMOUNT':'','ROW': ''}
    tp = ''
    tipo_dados = ''
    list_rel = []
    page = ''
    report_date = ''

    with open(arquivo_entrada) as file:

        while True:

            linha = file.readline()

            if len(linha.strip()) > 0:

                if corpo:

                    if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-110 REPORT ***' in linha):
                        corpo = False    

                    else:   
                        if len(linha[2:34].strip()) > 0 and  len(linha[35:134].strip()) == 0:
                            tipo_dados = linha[0:34].strip()

                        elif len(linha[2:8].strip()) > 0 and len(linha[44:134].strip()) > 0:
                            dic_dados_base = {'COUNT':'', 'CREDIT AMOUNT':'', 'DEBIT AMOUNT':'', 'TOTAL AMOUNT':'', 'ROW': ''}
                            dic_dados_base['COUNT'] = linha[40:60].strip()
                            dic_dados_base['CREDIT AMOUNT'] = linha[61:90].strip()
                            dic_dados_base['DEBIT AMOUNT'] = linha[91:115].strip()
                            dic_dados_base['TOTAL AMOUNT'] = linha[116:134].strip()
                            dic_dados_base['ROW'] = linha[0:30].strip()  
                            dic_dados_base['REPORT DATE'] = report_date 
                            dic_dados_base['SETTLEMENT CURRENCY'] = clearing_currency 
                            dic_dados_base['SOURCE'] = tipo_dados
                            list_rel.append(dic_dados_base)

                if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-110 REPORT ***' in linha):

                    if (len(page) > 0) or ('*** END OF VSS-110 REPORT ***' in linha):
                        list_rel.append(dic_dados_base)
                        tp = ''
                        page= linha[117:134].strip()

                    else:
                        page = linha[117:134].strip()

                if 'REPORT DATE:' in linha[111:134]:
                    report_date  = linha[124:134].strip()

                if 'SETTLEMENT CURRENCY:' in linha[0:22]:
                    clearing_currency = linha[21:134].strip()
                    corpo = True

            if not linha:
                break
                
    # Converter o dicionario em um pandas dataframe
    vss_110 = pd.DataFrame(list_rel)

    # Reordenando as colunas
    ord_col_vss110 = ['SOURCE', 'ROW', 'COUNT', 'CREDIT AMOUNT', 'DEBIT AMOUNT', 'TOTAL AMOUNT']
    vss_110 = vss_110[ord_col_vss110]

    # Convertendo o formato das moedas
    vss_110['CREDIT AMOUNT'] = vss_110['CREDIT AMOUNT'].apply(convert_currency_format)
    vss_110['DEBIT AMOUNT'] = vss_110['DEBIT AMOUNT'].apply(convert_currency_format)
    vss_110['TOTAL AMOUNT'] = vss_110['TOTAL AMOUNT'].apply(convert_currency_format)

    # Salvando o datraframe com CSV, separado os campos com ";"
    vss_110.to_csv(arquivo_saida, index=False, sep=';')


def convert_vss130_to_csv(arquivo_entrada, arquivo_saida):
    """
    Funcao para converter o relatorio VSS-130 em CSV
    """
    corpo = False
    dic_dados_base = { 'COUNT':'', 'INTERCHANGE AMOUNT':'', 'REIMBURSEMENT FEE CREDITS':'', 'REIMBURSEMENT FEE DEBITS':'', 'SOURCE': ''}
    tp = ''
    tipo_dados = ''
    list_rel = []
    page = ''
    report_date = ''
    settlement_currency = ''
    list_source = ['', '', '', '', '', '']
    posicao_list = -1

    with open(arquivo_entrada) as file:
        while True:
            linha = file.readline()
            if len(linha.strip()) > 0:
                if corpo:
                    if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-130 REPORT ***' in linha):
                        corpo = False    
                    else: 
                        if (len(linha[1:2].strip()) > 0): 
                            posicao_list = 0
                            list_source[posicao_list] = linha[0:45].strip()
                        elif (len(linha[4:5].strip()) > 0): 
                            posicao_list = 1
                            list_source[posicao_list] = linha[0:45].strip()
                        elif (len(linha[7:8].strip()) > 0): 
                            posicao_list = 2
                            list_source[posicao_list] = linha[0:45].strip()
                        elif (len(linha[10:11].strip()) > 0): 
                            posicao_list = 3
                            list_source[posicao_list] = linha[0:45].strip()
                        elif (len(linha[13:14].strip()) > 0): 
                            posicao_list = 4
                            list_source[posicao_list] = linha[0:45].strip()
                        elif (len(linha[16:17].strip()) > 0): 
                            posicao_list = 5
                            list_source[posicao_list] = linha[0:45].strip()
                        if posicao_list >= 0 and len(linha[44:134].strip()) > 0:
                            dic_dados_base = { 'COUNT':'', 'INTERCHANGE AMOUNT':'', 'REIMBURSEMENT FEE CREDITS':'', 'REIMBURSEMENT FEE DEBITS':'', 'SOURCE': ''}
                            dic_dados_base['COUNT'] = linha[56:73].strip()
                            dic_dados_base['INTERCHANGE AMOUNT'] = linha[74:94].strip()
                            dic_dados_base['REIMBURSEMENT FEE CREDITS'] = linha[95:112].strip()
                            dic_dados_base['REIMBURSEMENT FEE DEBITS'] = linha[113:134].strip()
                            dic_dados_base['SOURCE'] = list_source[:(posicao_list + 1)]
                            dic_dados_base['REPORT DATE'] = report_date 
                            dic_dados_base['SETTLEMENT CURRENCY'] = settlement_currency
                            dic_dados_base['PAGE'] = page

                            list_rel.append(dic_dados_base)

                if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-130 REPORT ***' in linha):

                    if (len(page) > 0) or ('*** END OF VSS-130 REPORT ***' in linha):
                        page = linha[117:134].strip()
                    else:
                        page = linha[117:134].strip()

                if 'REPORT DATE:' in linha[111:134]:
                    report_date  = linha[124:134].strip()

                if 'SETTLEMENT CURRENCY:' in linha[0:40]:
                    settlement_currency = linha[21:134].strip()
                    file.readline()
                    file.readline()
                    file.readline()
                    corpo = True

            if not linha:
                break
                
    # Converter o dicionario em um pandas dataframe
    vss_130 = pd.DataFrame(list_rel)

    # Reordenando as colunas do datraframe
    ord_col_vss130 = ['SOURCE', 'COUNT', 'INTERCHANGE AMOUNT', 'REIMBURSEMENT FEE CREDITS', 
                      'REIMBURSEMENT FEE DEBITS', 'REPORT DATE', 'PAGE']    
    vss_130 = vss_130[ord_col_vss130]

    # Convertendo o formato das moedas
    for i in range(len(vss_130)):
        if len(vss_130['INTERCHANGE AMOUNT'][i]) > 0:
            vss_130['INTERCHANGE AMOUNT'][i] = convert_currency_format(vss_130['INTERCHANGE AMOUNT'][i])
        if len(vss_130['REIMBURSEMENT FEE CREDITS'][i]) > 0:
            vss_130['REIMBURSEMENT FEE CREDITS'][i] = convert_currency_format(vss_130['REIMBURSEMENT FEE CREDITS'][i])
        if len(vss_130['REIMBURSEMENT FEE DEBITS'][i]) > 0:
            vss_130['REIMBURSEMENT FEE DEBITS'][i] = convert_currency_format(vss_130['REIMBURSEMENT FEE DEBITS'][i])  
    
    # Salvando o datraframe com CSV, separado os campos com ";"        
    vss_130.to_csv(arquivo_saida, index=False, sep=';')


def convert_vss900_to_csv(arquivo_entrada, arquivo_saida):
    """
    Funcao para converter o relatorio VSS-900 em CSV
    """
    # Variaveis
    corpo = False
    dic_dados_base = {'PAGE':'', 'REPORT DATE':'', 'CLEARING CURRENCY':'', 'BUSINESS MODE':'', 'SOURCE':'', 'CRS DATE': '', 
                      'COUNT':'', 'CLEARING AMOUNT':'', 'TOTAL COUNT':'', 'TOTAL CLEARING AMOUNT':'', 'TP': '','SUB': ''}
    tp = ''
    tipo_dados = ''
    list_rel = []
    page = ''
    report_date = ''
    clearing_currency = ''
    business_mode = ''

    with open(arquivo_entrada) as file:

        while True:
            linha = file.readline()
            if len(linha.strip()) > 0:
                if corpo:
                    if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-900 REPORT ***' in linha):
                        corpo = False    
                    else:   
                        if len(linha[2:5].strip()) > 0:
                            tipo_dados = linha[0:134].strip()
                            tp = ''
                        elif len(linha[2:8].strip()) > 0 and len(linha[44:134].strip()) == 0:
                            tp = linha[0:41].strip()
                        elif len(linha[2:8].strip()) == 0 and len(linha[2:10].strip()) > 0 and len(linha[44:134].strip()) > 0:
                            dic_dados_base = {'CRS DATE': '', 'COUNT':'', 'CLEARING AMOUNT':'', 'TOTAL COUNT':'', 'TOTAL CLEARING AMOUNT':'', 'TP': '','SUB': ''}
                            dic_dados_base['CRS DATE'] = linha[42:55].strip()
                            dic_dados_base['COUNT'] = linha[56:73].strip()
                            dic_dados_base['CLEARING AMOUNT'] = linha[74:94].strip()
                            dic_dados_base['TOTAL COUNT'] = linha[95:112].strip()
                            dic_dados_base['TOTAL CLEARING AMOUNT'] = linha[113:134].strip()
                            dic_dados_base['TP'] = tp
                            dic_dados_base['SUB'] = linha[0:41].strip()
                            dic_dados_base['PAGE'] = page
                            dic_dados_base['REPORT DATE'] = report_date 
                            dic_dados_base['CLEARING CURRENCY'] = clearing_currency 
                            dic_dados_base['BUSINESS MODE'] = business_mode
                            dic_dados_base['SOURCE'] = tipo_dados
                            list_rel.append(dic_dados_base)
                        elif len(linha[2:8].strip()) > 0 and len(linha[44:134].strip()) > 0:
                            dic_dados_base = {'CRS DATE': '', 'COUNT':'', 'CLEARING AMOUNT':'', 'TOTAL COUNT':'', 'TOTAL CLEARING AMOUNT':'', 'TP': '','SUB': ''}
                            dic_dados_base['CRS DATE'] = linha[42:55].strip()
                            dic_dados_base['COUNT'] = linha[56:73].strip()
                            dic_dados_base['CLEARING AMOUNT'] = linha[74:94].strip()
                            dic_dados_base['TOTAL COUNT'] = linha[95:112].strip()
                            dic_dados_base['TOTAL CLEARING AMOUNT'] = linha[113:134].strip()
                            dic_dados_base['TP'] = linha[0:41].strip()
                            dic_dados_base['PAGE'] = page
                            dic_dados_base['REPORT DATE'] = report_date 
                            dic_dados_base['CLEARING CURRENCY'] = clearing_currency 
                            dic_dados_base['BUSINESS MODE'] = business_mode
                            dic_dados_base['SOURCE'] = tipo_dados
                            list_rel.append(dic_dados_base)

                if ('PAGE: ' in linha[111:134]) or ('*** END OF VSS-900 REPORT ***' in linha):
                    if (len(page) > 0) or ('*** END OF VSS-900 REPORT ***' in linha):
                        list_rel.append(dic_dados_base)
                        tp = ''
                        page= linha[117:134].strip()
                    else:
                        page = linha[117:134].strip()

                if 'REPORT DATE:' in linha[111:134]:
                    report_date  = linha[124:134].strip()

                if 'CLEARING CURRENCY:' in linha[0:20]:
                    clearing_currency = linha[21:134].strip()

                if 'BUSINESS MODE:' in linha[0:134]:
                    business_mode = linha[21:134].strip()
                    corpo = True

            if not linha:
                break

    # Converter o dicionario em um pandas dataframe
    vss_900 = pd.DataFrame(list_rel)

    # Concatenando as colunas TP e SUB na coluna ROW  
    vss_900['ROW'] = vss_900[['TP','SUB']].apply(' '.join, axis=1)

    # Removendo as colunas TP e SUB
    vss_900 = vss_900.drop(['TP','SUB'], axis=1)

    # Reordenando as colunas do datraframe
    ord_col_vss900 = ['BUSINESS MODE', 'SOURCE', 'ROW', 'CRS DATE', 'COUNT', 'CLEARING AMOUNT', 'TOTAL COUNT', 
                      'TOTAL CLEARING AMOUNT', 'REPORT DATE', 'PAGE']
    vss_900 = vss_900[ord_col_vss900]

    # Convertendo o formato das moedas
    for i in range(len(vss_900)):
        if len(vss_900['CLEARING AMOUNT'][i]) > 0:
            vss_900['CLEARING AMOUNT'][i] = convert_currency_format(vss_900['CLEARING AMOUNT'][i])
        if len(vss_900['TOTAL CLEARING AMOUNT'][i]) > 0:
            vss_900['TOTAL CLEARING AMOUNT'][i] = convert_currency_format(vss_900['TOTAL CLEARING AMOUNT'][i])  
        if len(vss_900['COUNT'][i]) > 0:
            vss_900['COUNT'][i] = vss_900['COUNT'][i].replace(",", ".")
        if len(vss_900['TOTAL COUNT'][i]) > 0:
            vss_900['TOTAL COUNT'][i] = vss_900['TOTAL COUNT'][i].replace(",", ".")

    # Salvando o datraframe com CSV, separado os campos com ";"        
    vss_900.to_csv(arquivo_saida, index=False, sep=';')
# Importando as bibliotecas necessárias
import os
from utils import *


# Extrair os arquivos compactados
extrair_arquivo_rar('Case.rar','.TXT')

# Lista dos arquivos descompactados
lista_arquivos = sorted(listar_arquivos('./Case'))

# Criando diretorios para armazenamento dos dados
# Diretorio dos relatorios extraidos em txt
relatorios_txt = "relatorios/txt"
if not os.path.exists(relatorios_txt):
    os.makedirs(relatorios_txt)

# Diretorio dos relatorios extraidos em txt
relatorios_csv = "relatorios/csv"
if not os.path.exists(relatorios_csv):
    os.makedirs(relatorios_csv)

# Extrair os REPORT ID: VSS-110, FOR: 9000392692 BANCO CBSS e converte-los em formato CSV
num_relatorio = 0
for arquivo in lista_arquivos:
    relatorio = extrair_relatorio(f"Case/{arquivo}", "REPORT ID:  VSS-110 ", "*** END OF VSS-110 REPORT ***", "REPORTING FOR:      9000392692 BANCO CBSS")
    num_relatorio += 1
    salvar_relatorio(relatorio, f'relatorios/txt/relatorio_vss_110_{num_relatorio}.txt')
    convert_vss110_to_csv(f'relatorios/txt/relatorio_vss_110_{num_relatorio}.txt', f'relatorios/csv/relatorio_vss_110_{num_relatorio}.csv') 
    
# Extrair REPORT ID: VSS-130, FOR: 9000392692 BANCO CBSS
num_relatorio = 0
for arquivo in lista_arquivos:
    relatorio = extrair_relatorio(f"Case/{arquivo}", "REPORT ID:  VSS-130 ", "*** END OF VSS-130 REPORT ***", "REPORTING FOR:      9000392692 BANCO CBSS")
    num_relatorio += 1
    salvar_relatorio(relatorio, f'relatorios/txt/relatorio_vss_130_{num_relatorio}.txt')
    convert_vss130_to_csv(f'relatorios/txt/relatorio_vss_130_{num_relatorio}.txt', f'relatorios/csv/relatorio_vss_130_{num_relatorio}.csv')
        
# Extrair REPORT ID: VSS-900, FOR: 9000392692 BANCO CBSS
num_relatorio = 0
for arquivo in lista_arquivos:
    relatorio = extrair_relatorio(f"Case/{arquivo}", "REPORT ID:  VSS-900 ", "*** END OF VSS-900 REPORT ***", "REPORTING FOR:      9000392692 BANCO CBSS")
    num_relatorio += 1
    salvar_relatorio(relatorio, f'relatorios/txt/relatorio_vss_900_{num_relatorio}.txt')
    convert_vss900_to_csv(f'relatorios/txt/relatorio_vss_900_{num_relatorio}.txt', f'relatorios/csv/relatorio_vss_900_{num_relatorio}.csv') 
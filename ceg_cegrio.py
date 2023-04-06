import tabula
import pandas as pd
import numpy as np
import getpass
from openpyxl import load_workbook
pd.options.mode.chained_assignment = None

### INCIANDO O SCRIPT ###
# pegando ad do usuario
userName = getpass.getuser()

# caminho dos pfds
ceg_xlsx = 'C:/Users/'+userName+'/Downloads/ceg&cegrio.xlsx'

# leitura dos pdfs
ceg_pdf = tabula.read_pdf('C:/Users/'+userName +
                          '/Downloads/DELIBERACAO4502.pdf', pages='all')
cegrio_pdf = tabula.read_pdf(
    'C:/Users/'+userName+'/Downloads/DELIBERACAO4503.pdf', pages='all')

#### TARIFAS CEG #####

# data do documento CEG
data_vigencia = ceg_pdf[0][:1].iloc[:, [1]].\
    rename(columns={'TARIFAS CEG': 'Data da vigencia'}).\
    reset_index(drop=True)

# tarifa ceg rio
tarifa_ceg = ceg_pdf[3].iloc[1:7, [0, 1]].\
    rename(columns={'TARIFAS CEG': 'descrição', 'Unnamed: 0': 'valores'}).\
    reset_index(drop=True)

tarifa_ceg['data da vigencia'] = data_vigencia
tarifa_ceg['data da vigencia'] = tarifa_ceg['data da vigencia'].fillna(
    method="ffill")

# tarifa ceg rio gás natural - residencial
tarifa_ceg_gs = ceg_pdf[0][21:26].\
    rename(columns={'TARIFAS CEG': 'valores', 'Unnamed: 0': 'descrição'}).\
    reset_index(drop=True)
tarifa_ceg_gs = tarifa_ceg_gs.drop(labels=[2], axis=0)
tarifa_ceg_gs['data da vigencia'] = data_vigencia
tarifa_ceg_gs['data da vigencia'] = tarifa_ceg_gs['data da vigencia'].fillna(
    method="ffill")
tarifa_ceg_gs['descrição'] = tarifa_ceg_gs['descrição'].replace(
    np.nan, 'Residencial')

tarifa_ceg_gs['faixa de consumo: m³'] = 0
tarifa_ceg_gs['faixa de consumo: mês'] = 0
tarifa_ceg_gs['tarifa limite R$ / m³'] = 0

tarifa_ceg_gs['faixa de consumo: m³'].iloc[0] = tarifa_ceg_gs.loc[:,
                                                                  'valores'].iloc[0][:1]
tarifa_ceg_gs['faixa de consumo: m³'].iloc[1] = tarifa_ceg_gs.loc[:,
                                                                  'valores'].iloc[1][:1]
tarifa_ceg_gs['faixa de consumo: m³'].iloc[2] = tarifa_ceg_gs.loc[:,
                                                                  'valores'].iloc[2][:2]
tarifa_ceg_gs['faixa de consumo: m³'].iloc[3] = tarifa_ceg_gs.loc[:,
                                                                  'valores'].iloc[3][:12]
tarifa_ceg_gs['faixa de consumo: mês'].iloc[0] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[0][4:6]
tarifa_ceg_gs['faixa de consumo: mês'].iloc[1] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[1][4:7]
tarifa_ceg_gs['faixa de consumo: mês'].iloc[2] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[2][4:7]
tarifa_ceg_gs['faixa de consumo: mês'].iloc[3] = '-'
tarifa_ceg_gs['tarifa limite R$ / m³'].iloc[0] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[0][5:]
tarifa_ceg_gs['tarifa limite R$ / m³'].iloc[1] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[1][6:]
tarifa_ceg_gs['tarifa limite R$ / m³'].iloc[2] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[2][7:]
tarifa_ceg_gs['tarifa limite R$ / m³'].iloc[3] = tarifa_ceg_gs.loc[:,
                                                                   'valores'].iloc[3][13:]
tarifa_ceg_gs = tarifa_ceg_gs.drop(columns='valores')

# tarifa ceg rio gás natural - residencial mcmv, comercial, industrial e vidreiro
tarifaceg = ceg_pdf[8][1:30].iloc[:, [1, 3, 4]].\
    rename(columns={'Unnamed: 1': 'descrição',
                    '24 - 83': 'faixa de consumo: m³ / mês',
                    '14.1919': 'tarifa limite R$ / m³'
                    }).\
    reset_index(drop=True).\
    fillna(method="ffill")
new = tarifaceg["faixa de consumo: m³ / mês"].str.split("-", n=1, expand=True)
tarifaceg["faixa de consumo: m³"] = new[0]
tarifaceg["faixa de consumo: mês"] = new[1]
tarifaceg.drop(columns=["faixa de consumo: m³ / mês"], inplace=True)
tarifaceg['data da vigencia'] = data_vigencia
tarifaceg['data da vigencia'] = tarifaceg['data da vigencia'].fillna(
    method="ffill")
tarifaceg = tarifaceg.replace({None: '-'})

# tarifa ceg rio gás natural - residenciais, comercial, industrial e vidreiro
tarifa_gasNAT_ceg = pd.concat([tarifa_ceg_gs, tarifaceg])

# tarifa glp
tarifa_glp = ceg_pdf[2][1:3].\
    reset_index(drop=True).\
    rename(columns={'Unnamed: 0': 'descrição',
                    'Unnamed: 1': 'faixa unica',
                    'Unnamed: 2': 'tarifa'
                    })

tarifa_glp['data da vigencia'] = data_vigencia
tarifa_glp['data da vigencia'] = tarifa_glp['data da vigencia'].fillna(
    method="ffill")

# tarifa glp gás natural - industrial
tarifa_glp_ind = ceg_pdf[2][22:27].\
    rename(columns={'Unnamed: 0': 'descrição',
                    'Unnamed: 1': 'faixa de consumo: m³ / mês',
                    'Unnamed: 2': 'tarifa limite R$ / m³'}).\
    reset_index(drop=True).\
    fillna(method="ffill")
new = tarifa_glp_ind["faixa de consumo: m³ / mês"].str.split(
    "-", n=1, expand=True)
tarifa_glp_ind["faixa de consumo: m³"] = new[0]
tarifa_glp_ind["faixa de consumo: mês"] = new[1]
tarifa_glp_ind.drop(columns=["faixa de consumo: m³ / mês"], inplace=True)
tarifa_glp_ind['data da vigencia'] = data_vigencia
tarifa_glp_ind['data da vigencia'] = tarifa_glp_ind['data da vigencia'].fillna(
    method="ffill")
tarifa_glp_ind = tarifa_glp_ind.replace({None: '-'})
tarifa_glp_ind['descrição'] = tarifa_glp_ind['descrição'].replace(
    np.nan, 'Industrial')

#### TARIFAS CEG RIO ###

# data do documento
data_vigencia2 = cegrio_pdf[0][:1].iloc[:, [2]].\
    rename(columns={'Unnamed: 1': 'Data da vigencia'}).\
    reset_index(drop=True)

# tarifa ceg rio
tarifa_ceg_rio = cegrio_pdf[8][1:7][['Unnamed: 0', 'Unnamed: 3']].\
    rename(columns={'Unnamed: 0': 'descrição', 'Unnamed: 3': 'valores'}).\
    reset_index(drop=True)

tarifa_ceg_rio['data da vigencia'] = data_vigencia2
tarifa_ceg_rio['data da vigencia'] = tarifa_ceg_rio['data da vigencia'].fillna(
    method="ffill")

# tarifa ceg rio gás natural - residencial, comercial, industrial e vidreiro
tarifa_rio = cegrio_pdf[8][14:47].iloc[:, [0, 2, 4]].\
    rename(columns={'Unnamed: 0': 'descrição',
                    'Unnamed: 1': 'faixa de consumo: m³ / mês',
                    'Unnamed: 3': 'tarifa limite R$ / m³'
                    }).\
    reset_index(drop=True).\
    fillna(method="ffill")
new = tarifa_rio["faixa de consumo: m³ / mês"].str.split("-", n=1, expand=True)
tarifa_rio["faixa de consumo: m³"] = new[0]
tarifa_rio["faixa de consumo: mês"] = new[1]
tarifa_rio.drop(columns=["faixa de consumo: m³ / mês"], inplace=True)
tarifa_rio['data da vigencia'] = data_vigencia2
tarifa_rio['data da vigencia'] = tarifa_rio['data da vigencia'].fillna(
    method="ffill")
tarifa_rio = tarifa_rio.replace({None: '-'})

# tarifa glp
tarifa_glp_rio = cegrio_pdf[11].iloc[[4, 5], 0:3].\
    reset_index(drop=True).\
    rename(columns={'Unnamed: 0': 'descrição',
                    'R = Fator redutor cujo valor máximo é 1;': 'faixa unica',
                    'Unnamed: 1': 'tarifa'
                    })

tarifa_glp_rio['data da vigencia'] = data_vigencia2
tarifa_glp_rio['data da vigencia'] = tarifa_glp_rio['data da vigencia'].fillna(
    method="ffill")

# tarifa glp gás natural - industrial
tarifa_glp_ind_rio = cegrio_pdf[2][35:45].\
    rename(columns={'Barrilhista': 'descrição',
                    '0 - 200': 'faixa de consumo: m³ / mês',
                    '3,9604': 'tarifa limite R$ / m³'}).\
    reset_index(drop=True).\
    fillna(method="ffill")
new = tarifa_glp_ind_rio["faixa de consumo: m³ / mês"].str.split(
    "-", n=1, expand=True)
tarifa_glp_ind_rio["faixa de consumo: m³"] = new[0]
tarifa_glp_ind_rio["faixa de consumo: mês"] = new[1]
tarifa_glp_ind_rio.drop(columns=["faixa de consumo: m³ / mês"], inplace=True)
tarifa_glp_ind_rio['data da vigencia'] = data_vigencia2
tarifa_glp_ind_rio['data da vigencia'] = tarifa_glp_ind_rio['data da vigencia'].fillna(
    method="ffill")
tarifa_glp_ind_rio = tarifa_glp_ind_rio.replace({None: '-'})

# abrindo excel
book = load_workbook(ceg_xlsx)
writer = pd.ExcelWriter(ceg_xlsx, engine='openpyxl')

# inserindo os dados no excel
df_tarifa_ceg = tarifa_ceg.copy()
df_tarifa_ceg.to_excel(writer, sheet_name='CEG - Tarifas', index=False)

df_tarifa = tarifa_gasNAT_ceg.copy()
df_tarifa.to_excel(writer, sheet_name='CEG - Tarifa Gás Natural', index=False)

df_tarifa_glp = tarifa_glp.copy()
df_tarifa_glp.to_excel(writer, sheet_name='CEG - Tarifa GLP', index=False)

df_tarifa_glp_ind = tarifa_glp_ind.copy()
df_tarifa_glp_ind.to_excel(
    writer, sheet_name='CEG - Tarifa GLP Industrial', index=False)

df_tarifa_ceg_rio = tarifa_ceg_rio.copy()
df_tarifa_ceg_rio.to_excel(writer, sheet_name='CEG RIO - Tarifas', index=False)

df_tarifa_rio = tarifa_rio.copy()
df_tarifa_rio.to_excel(
    writer, sheet_name='CEG RIO - Tarifa Gás Natural', index=False)

df_tarifa_glp_rio = tarifa_glp_rio.copy()
df_tarifa_glp_rio.to_excel(
    writer, sheet_name='CEG RIO - Tarifa GLP', index=False)

df_tarifa_glp_ind_rio = tarifa_glp_ind_rio.copy()
df_tarifa_glp_ind_rio.to_excel(
    writer, sheet_name='CEG RIO - Tarifa GLP Industrial', index=False)

writer.close()

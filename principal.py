import pandas as pd
import requests as rq 
import pprint
import sqlite3
from localidades import nacional, estadual, municipal
import ssl
from Google import Create_Service
from googleapiclient.http import MediaFileUpload
import openpyxl
from openpyxl.styles import Font, Border, Side
from ajustar_planilha import ajustar_colunas, ajustar_bordas

api_nacional =  f'https://servicodados.ibge.gov.br/api/v3/agregados/5457/periodos/2013|2014|2015|2016|2017|2018|2019|2020|2021|2022/variaveis/8331|216|214|112?{nacional}&classificacao=782[40092,45982,40099,40101,40102,40136,40104,40137,40138,40139,40106,40143,40145,40112,40114,40149,40150,40151,40152,40261,40118,40119,40262,40263,40120,40121,40122,40266,40269,40124,40125,40271,40126,40127,40273,40274]'
api_estadual = f'https://servicodados.ibge.gov.br/api/v3/agregados/5457/periodos/2013|2014|2015|2016|2017|2018|2019|2020|2021|2022/variaveis/8331|216|214|112?{estadual}&classificacao=782[40092,45982,40099,40101,40102,40136,40104,40137,40138,40139,40106,40143,40145,40112,40114,40149,40150,40151,40152,40261,40118,40119,40262,40263,40120,40121,40122,40266,40269,40124,40125,40271,40126,40127,40273,40274]'

#api_municipal = f'https://servicodados.ibge.gov.br/api/v3/agregados/5457/periodos/2020|2021|2022/variaveis/8331|216|214|112|215?{municipal}&classificacao=782[40092,45982,40099,40101,40102,40136,40104,40137,40138,40139,40106,40143,40145,40112,40114,40149,40150,40151,40152,40261,40118,40119,40262,40263,40120,40121,40122,40266,40269,40124,40125,40271,40126,40127,40273,40274]'


class TLSAdapter(rq.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        ctx.options |= 0x4   # OP_LEGACY_SERVER_CONNECT
        kwargs["ssl_context"] = ctx
        return super(TLSAdapter, self).init_poolmanager(*args, **kwargs)

def requisitando_dados(api):
    with rq.session() as s:
        s.mount("https://", TLSAdapter())
        dados_brutos_api = s.get(api, verify=True)
    
    # Verificação se a solicitação foi bem-sucedida antes de continuar
    if dados_brutos_api.status_code != 200:
        raise Exception(f"A solicitação à API falhou com o código de status: {dados_brutos_api.status_code}")

    # Verificação se a resposta pode ser convertida para JSON
    try:
        dados_brutos = dados_brutos_api.json()
    except Exception as e:
        raise Exception(f"Erro ao analisar a resposta JSON da API: {str(e)}")

    # Verificação se a resposta contém os dados esperados
    if len(dados_brutos) < 4:
        raise Exception("A resposta da API não contém dados suficientes.")
    
    if dados_brutos_api.status_code == 500:
        raise Exception(f"Os dados passou de 100.0000 por isso o codigo de: {dados_brutos_api.status_code}")

    dados_brutos_8331 = dados_brutos[0]
    dados_brutos_216 = dados_brutos[1]
    dados_brutos_214 = dados_brutos[2]
    dados_brutos_112 = dados_brutos[3]

    return dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112



def tratando_dados(dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112 ):
    dados_limpos_8331 = []
    dados_limpos_216 = []
    dados_limpos_214 = []
    dados_limpos_112 = []

    variaveis = [dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112 ]

    for i in variaveis:
        id_tabela = i['id']
        variavel = i['variavel']
        unidade = i['unidade']
        dados = i['resultados']
        
        if unidade == '':
            unidade = 'Hectares'

        for ii in dados:
            dados_produto = ii['classificacoes']
            dados_producao = ii['series']
            
            for iii in dados_produto:
                dados_id_produto = iii['categoria']

                for id_produto, nome_produto in dados_id_produto.items():
                    nome_produto = nome_produto.replace('Abacaxi*', 'Abacaxi').replace(' (em caroço)', '').replace(' (em casca)', '').replace(' (em grão) Total', '')\
                    .replace('Coco-da-baía*', 'Coco-da-baía').replace(' (em grão)', '').replace(' (cacho)', '').replace(' (em amêndoa)', '').replace(' (fruto seco)', '')\
                    .replace('Borracha (látex coagulado)', 'Borracha látex coagulado').replace('Borracha (látex líquido)', 'Borracha látex líquido').replace(' (folha verde)', '')\
                    .replace(' (cacho de coco)', '').replace(' (em folha)', '').replace(' (semente)', '').replace(' (fibra)', '').replace(' (baga)', '')
                    
                    for iv in dados_producao:
                        id = iv['localidade']['id']
                        nome = iv['localidade']['nome'].replace(' (MT)', '')
                        dados_ano_producao = iv['serie'] 
                        
                        for ano, producao in dados_ano_producao.items():
                            producao = producao.replace('-', '0').replace('...', '0')
                            
                            dict = {

                                'id': id,
                                'nome': nome,
                                'id_produto': id_produto,
                                'produto': nome_produto,
                                variavel: producao,
                                'ano': f'01/01/{ano}'   
                            }
                           
                            if id_tabela == '8331':
                                dados_limpos_8331.append(dict)
                            elif id_tabela == '216':
                                dados_limpos_216.append(dict)
                            elif id_tabela == '214':
                                dados_limpos_214.append(dict)
                            elif id_tabela == '112':
                                dados_limpos_112.append(dict)


    return dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112

def gerando_dataframe(dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112):

    df8331 = pd.DataFrame(dados_limpos_8331)
    df216 = pd.DataFrame(dados_limpos_216)
    df214 = pd.DataFrame(dados_limpos_214)
    df112 = pd.DataFrame(dados_limpos_112)

    dataframe = pd.merge(df8331, df216, on=['id', 'nome', 'id_produto', 'produto', 'ano'], how='inner')
    dataframe = pd.merge(dataframe, df214, on=['id', 'nome', 'id_produto', 'produto', 'ano'], how='inner')
    dataframe = pd.merge(dataframe, df112, on=['id', 'nome', 'id_produto', 'produto', 'ano'], how='inner')
    dataframe['Área plantada ou destinada à colheita'] = dataframe['Área plantada ou destinada à colheita'].astype(float)
    dataframe['Área colhida'] = dataframe['Área colhida'].astype(float)
    dataframe['Quantidade produzida'] = dataframe['Quantidade produzida'].astype(float)
    dataframe['Rendimento médio da produção'] = dataframe['Rendimento médio da produção'].astype(float)
    return dataframe

def coluna_cultura(dataframe):
    
    culturas_temporarias = ['Abacaxi', 'Algodão herbáceo', 'Amendoim', 'Arroz', 'Batata-doce',
                            'Cana-de-açúcar', 'Feijão', 'Girassol', 'Mamona', 'Mandioca'
                            'Melancia', 'Melão', 'Milho', 'Soja', 'Sorgo', 'Tomate', 'Trigo']
    
    culturas_permanentes = ['Açaí', 'Algodão arbóreo', 'Banana', 'Borracha', 'Cacau', 'Café', 'Castanha de caju',
                            'Coco-da-baía', 'Goiaba', 'Guaraná', 'Laranja', 'Limão', 'Mamão', 'Manga',
                            'Maracujá', 'Palmito', 'Pimenta-do-reino', 'Tangerina', 'Urucum', 'Uva']
    
    
    for index, row in dataframe.iterrows():
        if row['produto'] in culturas_temporarias:
            dataframe.at[index, 'cultura'] = 'Temporária'
        elif row['produto'] in culturas_permanentes:
            dataframe.at[index, 'cultura'] = 'Permanente'
        else:
            dataframe.at[index, 'cultura'] = 'Outros'
    
    return dataframe


#PARTE NACIONAL
dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112 = requisitando_dados(api_nacional)
dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112 = tratando_dados(dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112)
dataframe = gerando_dataframe(dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112)
df5457_nacional = coluna_cultura(dataframe)
df5457_nacional.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_NACIONAL.xlsx', index=False)
df5457_nacional.to_html('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\CHATBOT\\Banco de dados Bot\\PAM 5457_NACIONAL.html', index=False)
#print(df5457_nacional)


#PARTE ESTADUAL
dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112 = requisitando_dados(api_estadual)
dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112 = tratando_dados(dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112)
dataframe = gerando_dataframe(dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112)
df5457_estadual = coluna_cultura(dataframe)
df5457_estadual.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_ESTADUAL.xlsx', index=False)
df5457_estadual.to_html('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\CHATBOT\\Banco de dados Bot\\PAM_5457_ESTADUAL.html', index=False)
#print(df5457_estadual)

'''
#PARTE MUNICIPAL
dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112 = requisitando_dados(api_municipal)
dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112 = tratando_dados(dados_brutos_8331, dados_brutos_216, dados_brutos_214, dados_brutos_112)
dataframe = gerando_dataframe(dados_limpos_8331, dados_limpos_216, dados_limpos_214, dados_limpos_112)
df5457_municipal = coluna_cultura(dataframe)
print(df5457_municipal)
df5457_estadual.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\PAM\\Tabelas em csv\\PAM_5457_MUNICIPAL.xlsx', index=False)

'''

# CARREGA A PLANILHA DO PAM 5457 E FAZ AS ALTERAÇÕES ESTRUTURAIS DA PLANILHA
wb_5457_nacional = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_NACIONAL.xlsx")  
wb_5457_estadual = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_ESTADUAL.xlsx")  

ws_5457_nacional = wb_5457_nacional.active
ws_5457_estadual = wb_5457_estadual.active

lista_ws = [ws_5457_nacional, ws_5457_estadual]
lista_wb = [wb_5457_nacional, wb_5457_estadual]
for ws, wb in zip(lista_ws, lista_wb):
    ajustar_colunas(ws)
    ajustar_bordas(wb)

wb_5457_nacional.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_NACIONAL.xlsx")
wb_5457_estadual.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_ESTADUAL.xlsx")


#TESTANDO JUNTAR TODAS PLANILHAS EM UMA SÓ
planilha_principal = openpyxl.Workbook()

wb_5457_nacional = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_NACIONAL.xlsx")  
wb_5457_estadual = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM_5457_ESTADUAL.xlsx")  

aba_5457_nacional = planilha_principal.create_sheet("PAM 5457 NACIONAL")
aba_5457_estadual = planilha_principal.create_sheet("PAM 5457 ESTADUAL")

# Copiar os dados da primeira planilha para a nova planilha
for linha in wb_5457_nacional.active.iter_rows(values_only=True):
    aba_5457_nacional.append(linha)

# Copiar os dados da segunda planilha para a nova planilha
for linha in wb_5457_estadual.active.iter_rows(values_only=True):
    aba_5457_estadual.append(linha)

for aba in planilha_principal.sheetnames:
    if aba not in ["PAM 5457 NACIONAL", "PAM 5457 ESTADUAL"]:
        del planilha_principal[aba]


colunas_para_ajustar = ['B', 'C', 'D', 'F', 'G', 'J']
largura_desejada = 22

for coluna in colunas_para_ajustar:
    aba_5457_nacional.column_dimensions[coluna].width = largura_desejada
    aba_5457_estadual.column_dimensions[coluna].width = largura_desejada

colunas_maiores = ['E', 'H', 'I']
largura_planejada = 35

for coluna in colunas_maiores:
    aba_5457_nacional.column_dimensions[coluna].width = largura_planejada
    aba_5457_estadual.column_dimensions[coluna].width = largura_planejada


planilha_principal.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM.xlsx")
worksheet = planilha_principal.active
df = pd.read_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM.xlsx')

for sheet_name in planilha_principal.sheetnames:
    worksheet = planilha_principal[sheet_name]
    
    for col_num in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

planilha_principal.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\PAM.xlsx")

#Faz autenticação do google drive para jogar os arquivos gerados no codigo python
CLIENT_SECRET_FILE = 'credencials.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ["https://www.googleapis.com/auth/drive"]

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

#PASSA O PAM PARA O DRIVE
file_id = "1KSR_XbhD-SEL9Wpu-f8SEZT0ETGhGeLb"
FILE_NAMES = ["PAM_5457_NACIONAL.xlsx", "PAM_5457_ESTADUAL.xlsx"]
MIME_TYPES = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]

#LISTA TODOS OS ARQUIVO DENTRO DA PASTA
def listar_arquivos():
    results = service.files().list(
        q=f"trashed=false and '{file_id}' in parents",
        spaces='drive',
        pageSize=10,  # Ajuste o valor conforme necessário
        fields="nextPageToken, files(id, name, createdTime)"
    ).execute()
    items = results.get('files', [])
    items_sorted = sorted(items, key=lambda x: x['createdTime']) 
    return items_sorted

def obter_id_do_arquivo(file_name):
    items = listar_arquivos()
    for item in items:
        if item['name'] == file_name:
            return item['id']
    return None 
    

#ADICIONA TODOS OS ARQUIVOS NA PASTA
for file_name, mime_type in zip(FILE_NAMES, MIME_TYPES):
    id_arquivo = obter_id_do_arquivo(file_name)

    if id_arquivo:
        # O arquivo já existe, então atualizamos
        media_replace = MediaFileUpload("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\{0}".format(file_name), mimetype=mime_type)
        service.files().update(
            fileId=id_arquivo,
            media_body=media_replace
        ).execute()
        print(f"Documento '{file_name}' atualizado")
    else:
        file_metadata = {
            "name": file_name,
            "parents": [file_id]
        }
        media = MediaFileUpload("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\TABELAS EM CSV\\{0}".format(file_name), mimetype=mime_type)

        service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()
        print(f"Arquivo '{file_name}' criado")


'''
#ABRE O ARQUIVO SQL.PY E EXECUTA TODOS OS COMANDOS DENTRO DELE
if __name__ == '__main__':
    from sql import executar_sql 
    executar_sql()

'''
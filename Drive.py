#LISTA TODOS OS ARQUIVO DENTRO DA PASTA
from googleapiclient.http import MediaFileUpload

def listar_arquivos(service,file_id):
    results = service.files().list(
        q=f"trashed=false and '{file_id}' in parents",
        spaces='drive',
        pageSize=10,  # Ajuste o valor conforme necessário
        fields="nextPageToken, files(id, name, createdTime)"
    ).execute()
    items = results.get('files', [])
    items_sorted = sorted(items, key=lambda x: x['createdTime']) 
    return items_sorted

def obter_id_do_arquivo(file_name, service,file_id):
    items = listar_arquivos(service, file_id)
    for item in items:
        if item['name'] == file_name:
            return item['id']
    return None 
    
def add_arquivos_a_pasta(FILE_NAMES, MIME_TYPES, service, file_id):
    for file_name, mime_type in zip(FILE_NAMES, MIME_TYPES):
        id_arquivo = obter_id_do_arquivo(file_name, service,file_id)

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

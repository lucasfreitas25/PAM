from principal import df5457_estadual
import psycopg2
from conexão import conexao
def executar_sql():
    
    cursor = conexao.cursor()
    cursor.execute('SET search_path TO pam, public')
    
    criando_tabela_estadual = f'''
        CREATE TABLE IF NOT EXISTS pam.PAM_5457_ESTADUAL (
            id_pam_5457_estadual SERIAL PRIMARY KEY,
            id INTEGER,
            nome TEXT,
            id_produto INTEGER,
            produto TEXT, 
            area_plantada INTEGER,
            area_colhida INTEGER,
            quantidade_produzida INTEGER, 
            rendimento_producao INTEGER, 
            ano DATE,
            cultura TEXT);
        '''
    cursor.execute(criando_tabela_estadual)
    
    verificando_existencia_estadual = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_type='BASE TABLE' AND table_name='pam_5457_estadual';
    '''
    
    cursor.execute(verificando_existencia_estadual)
    resultado_estadual = cursor.fetchone()
    # Verifique se as tabelas existem e exclua todos dados da tabela

    if resultado_estadual[0] == 1:
        dropando_tabela_estadual = '''
        TRUNCATE PAM_5457_ESTADUAL;
        '''
        cursor.execute(dropando_tabela_estadual)


    inserindo_dados_estadual = '''
    INSERT INTO pam.PAM_5457_ESTADUAL (id, nome, id_produto, produto, area_plantada, area_colhida, quantidade_produzida, rendimento_producao, ano)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    '''
    try:
        for idx, i in df5457_estadual.iterrows():
            dados = (
                i['id'],
                i['nome'],
                i['id_produto'],
                i['produto'],
                i['Área plantada ou destinada à colheita'],
                i['Área colhida'],
                i['Quantidade produzida'],
                i['Rendimento médio da produção'],
                i['ano']
            )
            cursor.execute(inserindo_dados_estadual, dados)
            conexao.commit()

    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")

    atualizando_cultura_estadual = '''
    UPDATE pam.PAM_5457_ESTADUAL
    SET cultura = %s
    WHERE produto = %s AND ano = %s;
    '''

    for idx, i in df5457_estadual.iterrows():
        cultura = i['cultura']  # Substitua 'cultura' pelo nome real da coluna em seu DataFrame
        produto = i['produto']
        ano = i['ano']
        cursor.execute(atualizando_cultura_estadual, (cultura, produto, ano))

    conexao.commit()
    conexao.close()

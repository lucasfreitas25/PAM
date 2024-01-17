import sqlite3
from principal import df5457_nacional, df5457_estadual

def executar_sql():
    conexao = sqlite3.connect('Z:\\Lucas Freitas Arquivos\\DATAHUB\\TABELAS\\DATAHUB_DATABASE.db')
    cursor = conexao.cursor()

    #FALTA A PARTE MUNICIPAL TEM QUE VER A API
    verificando_existencia_nacional = '''
    SELECT 1
    FROM sqlite_master
    WHERE type='table' AND name='SIDRA_PAM_5457_NACIONAL';
    '''
    verificando_existencia_estadual = '''
    SELECT 1
    FROM sqlite_master
    WHERE type='table' AND name='SIDRA_PAM_5457_ESTADUAL';
    '''

    # Execute as consultas de verificação
    cursor.execute(verificando_existencia_nacional)
    resultado_nacional = cursor.fetchone()
    cursor.execute(verificando_existencia_estadual)
    resultado_estadual = cursor.fetchone()

    # Verifique se as tabelas existem e exclua, se necessário
    if resultado_nacional and resultado_nacional[0] == 1:
        dropando_tabela_nacional = '''  
        DROP TABLE SIDRA_PAM_5457_NACIONAL;
        '''
        cursor.execute(dropando_tabela_nacional)

    if resultado_estadual and resultado_estadual[0] == 1:
        dropando_tabela_estadual = '''
        DROP TABLE SIDRA_PAM_5457_ESTADUAL;
        '''
        cursor.execute(dropando_tabela_estadual)

    criando_tabela_nacional = f'''
    CREATE TABLE IF NOT EXISTS SIDRA_PAM_5457_NACIONAL (
        id_pam_5457_nacional INTEGER PRIMARY KEY AUTOINCREMENT,
        id INTEGER,
        nome TEXT,
        id_produto INTEGER,
        produto TEXT, 
        area_plantada NUMERIC,
        area_colhida NUMERIC,
        quantidade_produzida NUMERIC, 
        rendimento_producao NUMERIC, 
        ano DATE,
        cultura TEXT);
    '''

    criando_tabela_estadual = f'''
    CREATE TABLE IF NOT EXISTS SIDRA_PAM_5457_ESTADUAL (
        id_pam_5457_estadual INTEGER PRIMARY KEY AUTOINCREMENT,
        id INTEGER,
        nome TEXT,
        id_produto INTEGER,
        produto TEXT, 
        area_plantada NUMERIC,
        area_colhida NUMERIC,
        quantidade_produzida NUMERIC, 
        rendimento_producao NUMERIC, 
        ano DATE,
        cultura TEXT);
    '''
    cursor.execute(criando_tabela_nacional)
    cursor.execute(criando_tabela_estadual)

    inserindo_dados_nacional = '''
    INSERT INTO SIDRA_PAM_5457_NACIONAL (id, nome, id_produto, produto, area_plantada, area_colhida, quantidade_produzida, rendimento_producao, ano)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    '''

    for idx, i in df5457_nacional.iterrows():
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
        cursor.execute(inserindo_dados_nacional, dados)

    inserindo_dados_estadual = '''
    INSERT INTO SIDRA_PAM_5457_ESTADUAL (id, nome, id_produto, produto, area_plantada, area_colhida, quantidade_produzida, rendimento_producao, ano)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    '''

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

    atualizando_cultura_nacional = '''
    UPDATE SIDRA_PAM_5457_NACIONAL
    SET cultura = ?
    WHERE produto = ? AND ano = ?;
    '''

    for idx, i in df5457_nacional.iterrows():
        cultura = i['cultura']  # Substitua 'cultura' pelo nome real da coluna em seu DataFrame
        produto = i['produto']
        ano = i['ano']
        cursor.execute(atualizando_cultura_nacional, (cultura, produto, ano))

    conexao.commit()
    conexao.close()

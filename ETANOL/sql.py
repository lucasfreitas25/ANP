import psycopg2
from ETL_ETANOL import df_capacidade, df_matprima, df_prod
from conexão import conexao

def executar_sql():
    cur = conexao.cursor()
    
    cur.execute('SET search_path TO anp, public')
    
    
    etanol_capacidade = \
    '''
    CREATE TABLE IF NOT EXISTS anp.etanol_capacidade (
        id_etanol_capacidade SERIAL PRIMARY KEY,
        Data DATE,
        Razao_Social VARCHAR(255),
        CNPJ CHAR(18),
        Região VARCHAR(30),
        Estado VARCHAR(30),
        Município VARCHAR(50),
        Capacidade_Produ_Etanol_Anidro INTEGER,
        Capacidade_Prod_Etanol_Hidratado INTEGER);
    '''
    etanol_matprima = \
    '''
    CREATE TABLE IF NOT EXISTS anp.etanol_matprima (
        id_etanol_matprima SERIAL PRIMARY KEY,
        Data DATE,
        Região VARCHAR(30),
        Estado VARCHAR(25),
        Produto VARCHAR(30),
        Quantidade_Processada NUMERIC(18, 2));
    '''
    etanol_producao = \
    '''
    CREATE TABLE IF NOT EXISTS anp.etanol_producao (
        id_etanol_producao SERIAL PRIMARY KEY,
        Data DATE,
        Região VARCHAR(100),
        Estado VARCHAR(100),
        Prod_Etanol_Hidratado INTEGER,
        Prod_Etanol_Anidro INTEGER);
    '''

    cur.execute(etanol_capacidade)
    cur.execute(etanol_matprima)
    cur.execute(etanol_producao)

    verificando_existencia_etanol_capacidade = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='etanol_capacidade';
    '''
    verificando_existencia_etanol_matprima = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='etanol_matprima';
    '''
    verificando_existencia_etanol_producao = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='etanol_producao';
    '''
    

    # Execute as consultas de verificação
    cur.execute(verificando_existencia_etanol_capacidade)
    resultado_etanol_capacidade= cur.fetchone()
    cur.execute(verificando_existencia_etanol_matprima)
    resultado_etanol_matprima= cur.fetchone()
    cur.execute(verificando_existencia_etanol_producao)
    resultado_etanol_producao= cur.fetchone()
    
    # Verifique se as tabelas existem e exclua, se necessário
    if resultado_etanol_capacidade[0] == 1:
        dropando_tabela_etanol_capacidade = '''
        TRUNCATE TABLE anp.etanol_capacidade;
        '''
        cur.execute(dropando_tabela_etanol_capacidade)

    if resultado_etanol_matprima[0] == 1:
        dropando_tabela_etanol_matprima = '''
        TRUNCATE TABLE anp.etanol_matprima;
        '''
        cur.execute(dropando_tabela_etanol_matprima)

    if resultado_etanol_producao[0] == 1:
        dropando_tabela_etanol_producao = '''
        TRUNCATE TABLE anp.etanol_producao;
        '''
        cur.execute(dropando_tabela_etanol_producao)


    #INSERINDO DADOS
    inserindo_etanol_capacidade= \
    '''
    INSERT INTO anp.etanol_capacidade (Data, Razao_Social, CNPJ, Região, Estado, Município, Capacidade_Produ_Etanol_Anidro, Capacidade_Prod_Etanol_Hidratado) VALUES (%s, %s, %s, %s, %s, %s, %s, %s);
    '''
    try:
        for idx, i in df_capacidade.iterrows():
            dados = (
                i['Data'],
                i['Razão Social'],
                i['CNPJ'],
                i['Região'],
                i['Estado'],
                i['Município'],
                i['Capacidade Produção Etanol Anidro (m³/d)'],
                i['Capacidade Produção Etanol Hidratado (m³/d)']
            )
            cur.execute(inserindo_etanol_capacidade, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")
        
    inserindo_etanol_matprima= \
    '''
    INSERT INTO anp.etanol_matprima (Data, Região, Estado, Produto, Quantidade_Processada) VALUES(%s,%s,%s,%s,%s) 
    '''
    try:
        for idx, i in df_matprima.iterrows():
            dados = (
                i['Data'],
                i['Região'],
                i['Estado'],
                i['Produto'],
                i['Quantidade Processada (t)']
            )
            cur.execute(inserindo_etanol_matprima, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")
        
    inserindo_etanol_producao= \
    '''
    INSERT INTO anp.etanol_producao(Data, Região, Estado, Prod_Etanol_Hidratado, Prod_Etanol_Anidro)
    VALUES(%s,%s,%s,%s,%s) 
    '''
    try:
        for idx, i in df_prod.iterrows():
            dados = (
                i['Data'],
                i['Região'],
                i['Estado'],
                i['Produção Etanol Hidratado(m³)'],
                i['Produção Etanol Anidro (m³)']
            )
            cur.execute(inserindo_etanol_producao, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")

    conexao.commit()
    conexao.close()
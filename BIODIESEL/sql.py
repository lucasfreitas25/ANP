import psycopg2
from ETL_Biodiesel import df_capacidade, df_matprima, df_prod, df_vendas
from conexão import conexao

def executar_sql():
    cur = conexao.cursor()
    
    cur.execute('SET search_path TO anp, public')
    
    
    biodiesel_capacidade = \
    '''
    CREATE TABLE IF NOT EXISTS anp.biodiesel_capacidade (
        Data DATE,
        Razão_Social VARCHAR(255),
        CNPJ CHAR(20),
        Região VARCHAR(25),
        Estado VARCHAR(25),
        Município VARCHAR(35),
        Capacidade_Produção_Biodiesel INTEGER,
        Tancagem_Biodiesel INTEGER
    );
    '''
    biodiesel_matprima = \
    '''
    CREATE TABLE IF NOT EXISTS anp.biodiesel_matprima (
        Data DATE,
        Região VARCHAR(25),
        Estado VARCHAR(25),
        Produto VARCHAR(50),
        Quantidade DECIMAL(10,2));
    '''
    biodiesel_producao = \
    '''
    CREATE TABLE IF NOT EXISTS anp.biodiesel_producao (
        id_biodiesel_producao SERIAL PRIMARY KEY,
        Data DATE,
        Regiao VARCHAR(100),
        Producao_Biodiesel DECIMAL(10,2));
    '''
    
    biodiesel_vendas = \
    '''
    CREATE TABLE IF NOT EXISTS anp.biodiesel_vendas (
        Data DATE,
        Regiao_Origem VARCHAR(25),
        Regiao_Destino VARCHAR(25),
        Vendas_Biodiesel DECIMAL(10,2));
    '''

    cur.execute(biodiesel_capacidade)
    cur.execute(biodiesel_matprima)
    cur.execute(biodiesel_producao)
    cur.execute(biodiesel_vendas)

    verificando_existencia_biodiesel_capacidade = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='biodiesel_capacidade';
    '''
    verificando_existencia_biodiesel_matprima = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='biodiesel_matprima';
    '''
    verificando_existencia_biodiesel_producao = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='biodiesel_producao';
    '''
    verificando_existencia_biodiesel_vendas = '''
    SELECT 1
    FROM information_schema.tables
    WHERE table_schema= 'anp' AND table_type='BASE TABLE' AND table_name='biodiesel_vendas';
    '''
    

    # Execute as consultas de verificação
    cur.execute(verificando_existencia_biodiesel_capacidade)
    resultado_biodiesel_capacidade = cur.fetchone()
    
    cur.execute(verificando_existencia_biodiesel_matprima)
    resultado_biodiesel_matprima = cur.fetchone()
    
    cur.execute(verificando_existencia_biodiesel_producao)
    resultado_biodiesel_producao = cur.fetchone()
    
    cur.execute(verificando_existencia_biodiesel_vendas)
    resultado_biodiesel_vendas = cur.fetchone()
    
    # Verifique se as tabelas existem e exclua, se necessário
    if resultado_biodiesel_capacidade[0] == 1:
        dropando_tabela_biodiesel_capacidade = '''
        TRUNCATE TABLE anp.biodiesel_capacidade;
        '''
        cur.execute(dropando_tabela_biodiesel_capacidade)

    if resultado_biodiesel_matprima[0] == 1:
        dropando_tabela_biodiesel_matprima = '''
        TRUNCATE TABLE anp.biodiesel_matprima;
        '''
        cur.execute(dropando_tabela_biodiesel_matprima)

    if resultado_biodiesel_producao[0] == 1:
        dropando_tabela_biodiesel_producao = '''
        TRUNCATE TABLE anp.biodiesel_producao;
        '''
        cur.execute(dropando_tabela_biodiesel_producao)
        
    if resultado_biodiesel_vendas[0] == 1:
        dropando_tabela_biodiesel_vendas = '''
        TRUNCATE TABLE anp.biodiesel_vendas;
        '''
        cur.execute(dropando_tabela_biodiesel_vendas)


    #INSERINDO DADOS
    inserindo_biodiesel_capacidade= \
    '''
    INSERT INTO anp.biodiesel_capacidade (Data, Razão_Social, CNPJ, Região, Estado, Município, Capacidade_Produção_Biodiesel, Tancagem_Biodiesel) VALUES (%s, %s, %s, %s, %s, %s, %s, %s);
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
                i['Capacidade Produção Biodiesel (m³/d)'],
                i['Tancagem Biodiesel (m³)']
            )
            cur.execute(inserindo_biodiesel_capacidade, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")
        
    inserindo_biodiesel_matprima= \
    '''
    INSERT INTO anp.biodiesel_matprima (Data, Região, Estado, Produto, Quantidade) VALUES(%s,%s,%s,%s,%s) 
    '''
    try:
        for idx, i in df_matprima.iterrows():
            dados = (
                i['Data'],
                i['Região'],
                i['Estado'],
                i['Produto'],
                i['Quantidade (m³)']
            )
            cur.execute(inserindo_biodiesel_matprima, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")
        
    inserindo_biodiesel_producao= \
    '''
    INSERT INTO anp.biodiesel_producao(Data, Regiao, Producao_Biodiesel)
    VALUES(%s,%s,%s) 
    '''
    try:
        for idx, i in df_prod.iterrows():
            dados = (
                i['Data'],
                i['Região'],
                i['Produção de Biodiesel']
            )
            cur.execute(inserindo_biodiesel_producao, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")
        
    inserindo_biodiesel_vendas= \
    '''
    INSERT INTO anp.biodiesel_vendas(Data, Regiao_Origem, Regiao_Destino, Vendas_Biodiesel)
    VALUES(%s,%s,%s,%s) 
    '''
    try:
        for idx, i in df_vendas.iterrows():
            dados = (
                i['Data'],
                i['Região Origem'],
                i['Região Destino'],
                i['Vendas de Biodiesel']
            )
            cur.execute(inserindo_biodiesel_vendas, dados)
    except psycopg2.Error as e:
        print(f"Erro ao inserir dados estaduais: {e}")

    conexao.commit()
    conexao.close()
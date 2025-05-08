import os
import shutil
from datetime import datetime, timedelta

import pyodbc
DB_CONFIG = {
    "server": "192.30.100.206,pip install pymssql1433",
    "database": "bdloyal",
    "username": "raul",
    "password": "123456",
    "driver": "{ODBC Driver 17 for SQL Server}"
}


def conectar_sql():
    try:
        conn = pyodbc.connect(
            f"DRIVER={DB_CONFIG['driver']};"
            f"SERVER={DB_CONFIG['server']};"
            f"DATABASE={DB_CONFIG['database']};"
            f"UID={DB_CONFIG['username']};"
            f"PWD={DB_CONFIG['password']}"
        )
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco: {e}")
        return None
    
def obter_mes_passado():
    hoje = datetime.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    mes_passado = primeiro_dia_mes_atual - timedelta(days=1)
    return mes_passado.strftime('%m%Y')

def copiar_arquivos(origem_base, destino_base):
    mes_passado = obter_mes_passado()
    
    for pasta_empresa in os.listdir(origem_base):
        caminho_empresa = os.path.join(origem_base, pasta_empresa)
        if os.path.isdir(caminho_empresa):
            caminho_mes = os.path.join(caminho_empresa, mes_passado)
            if os.path.exists(caminho_mes):
                caminho_nfse = os.path.join(caminho_mes, "NFSE PRESTADO")
                if os.path.exists(caminho_nfse):
                    codigo_empresa = pasta_empresa[:4]
                    codigo_empresa= codigo_empresa+'\\'+'servico'
                    destino_empresa = os.path.join(destino_base, codigo_empresa)
                    os.makedirs(destino_empresa, exist_ok=True)
                    
                    for arquivo in os.listdir(caminho_nfse):
                        if arquivo.endswith(".txt"):
                            origem_arquivo = os.path.join(caminho_nfse, arquivo)
                            destino_arquivo = os.path.join(destino_empresa, arquivo)
                            shutil.copy2(origem_arquivo, destino_arquivo)
                            print(f"Copiado: {origem_arquivo} -> {destino_arquivo}")
                            #UPDATE SuaTabela SET QuantidadeArquivo = QuantidadeArquivo + 5 WHERE CodInterno = 'ABC123';
                            conn = conectar_sql()
                            
                            cursor = conn.cursor()
                            # Comando SQL de INSERT
                            sql_update = '''
                                UPDATE RelatorioHubSieg
                                SET QuantidadeArquivo = QuantidadeArquivo + 1
                                WHERE CodInterno = ?
                            '''
                            codigo_empresaF = codigo_empresa[:4]
                            # Executando a atualização
                            cursor.execute(sql_update, (codigo_empresaF,))                          
                            
                            # Executar o UPDATE                            
                            conn.commit()

                            # Fechar a conexão
                            cursor.close()
                            conn.close()
                else:
                    print(f"A pasta 'NFSE PRESTADO' não foi encontrada em {caminho_mes}")
            else:
                print(f"A pasta do mês ({mes_passado}) não foi encontrada em {caminho_empresa}")

if __name__ == "__main__":
    origem = r"C:\Users\Eunice Ramalho\Meu Drive\PREFEITURA SÃO PAULO"
    destino = r"Z:\MODELO ALTERDATA"
    

    # Listando todas as pastas no diretório
    pastas = [f for f in os.listdir(origem) if os.path.isdir(os.path.join(origem, f))]

    # Inserindo os dados na tabela RelatorioHubSieg com apenas os 4 primeiros dígitos do nome da pasta
    for pasta in pastas:
        conn = conectar_sql()
        cursor = conn.cursor()
        # Pegando os primeiros 4 caracteres do nome da pasta
        cod_interno = pasta[:4]
    
        # Comando SQL para inserir o nome da pasta na coluna CodInterno e 0 na coluna QuantidadeArquivo
        sql_insert = '''
        INSERT INTO RelatorioHubSieg (CodInterno, QuantidadeArquivo)
        VALUES (?, 0)  -- Colocando 0 para QuantidadeArquivo
        '''
        cursor.execute(sql_insert, (cod_interno,))

        # Commit para garantir que as alterações sejam salvas
        conn.commit()

        # Fechando a conexão
        cursor.close()
        conn.close()

    copiar_arquivos(origem, destino)
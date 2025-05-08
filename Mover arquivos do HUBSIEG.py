import pyodbc
import os
import shutil
import re
import locale
from datetime import datetime, timedelta
import unicodedata
# Define o locale para exibir meses em português
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")  # Para Linux/macOS
except:
    locale.setlocale(locale.LC_TIME, "Portuguese_Brazil.1252")  # Para Windows

# Configuração do SQL Server
DB_CONFIG = {
    "server": "192.30.100.206,1433",
    "database": "bdloyal",
    "username": "raul",
    "password": "123456",
    "driver": "{ODBC Driver 17 for SQL Server}"
}

# Diretórios base
BASE_DIR = r"C:\Users\Eunice Ramalho\HUB SIEG\XML"
DESTINO_DIR = r"Z:\MODELO ALTERDATA"







# Obtém o nome do mês passado em português
hoje = datetime.today()
ano_atual = hoje.year
mes_passado = (hoje.replace(day=1) - timedelta(days=1)).strftime("%B").capitalize()
if mes_passado == 'Marã§o':
    mes_passado = 'Marco'
# Dicionário para converter nome do mês para português corretamente
meses_traduzidos = {
    "January": "Janeiro", "February": "Fevereiro", "March": "Marco",
    "April": "Abril", "May": "Maio", "June": "Junho",
    "July": "Julho", "August": "Agosto", "September": "Setembro",
    "October": "Outubro", "November": "Novembro", "December": "Dezembro"
}

mes_passado = meses_traduzidos.get(mes_passado, mes_passado)

# Conectar ao SQL Server
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

# Remove caracteres especiais do CNPJ
def limpar_cnpj(cnpj):
    return re.sub(r'\D', '', cnpj)  # Remove tudo que não for número

# Lê empresas do banco de dados
def obter_empresas():
    conn = conectar_sql()
    if not conn:
        return []

    cursor = conn.cursor()
    cursor.execute("SELECT CNPJ, CodSistema FROM Empresas")
    empresas = cursor.fetchall()
    conn.close()
    
    return [(limpar_cnpj(str(cnpj)), str(cod)) for cnpj, cod in empresas]

# Move arquivos da pasta CT-es para servico
def mover_ctes(cnpj, cod_interno):
    origem_pasta = os.path.join(BASE_DIR, "CT-es", cnpj, cnpj, str(ano_atual), mes_passado)
    destino_pasta = os.path.join(DESTINO_DIR, cod_interno, "entrada_saida")
    mover_arquivos(origem_pasta, destino_pasta,cod_interno)

# Move arquivos da pasta NFS-es para servico
def mover_nfses(cnpj, cod_interno):
    origem_pasta = os.path.join(BASE_DIR, "NFS-es", cnpj, "Servico Prestado", str(ano_atual), mes_passado)
    destino_pasta = os.path.join(DESTINO_DIR, cod_interno, "servico")
    mover_arquivos(origem_pasta, destino_pasta,cod_interno)

# Move arquivos da pasta SAT para entrada_saida
def mover_sat(cnpj, cod_interno):
    base_sat = os.path.join(BASE_DIR, "SAT", cnpj, cnpj, str(ano_atual), mes_passado)

    if not os.path.exists(base_sat):
        print(f"🚫 Pasta SAT não encontrada: {base_sat}")
        return

    pastas_internas = [p for p in os.listdir(base_sat) if os.path.isdir(os.path.join(base_sat, p))]
    
    if not pastas_internas:
        print(f"🚫 Nenhuma subpasta encontrada em: {base_sat}")
        return

    for subpasta in pastas_internas:
        origem_pasta = os.path.join(base_sat, subpasta)
        destino_pasta = os.path.join(DESTINO_DIR, cod_interno, "entrada_saida")
        mover_arquivos(origem_pasta, destino_pasta,cod_interno)

# Move arquivos das pastas Entrada e Saída para entrada_saida
def mover_entrada_saida(cnpj, cod_interno):
    origens = [
        os.path.join(BASE_DIR,"NF-es", cnpj, "Entrada", str(ano_atual), mes_passado),
        os.path.join(BASE_DIR,"NF-es", cnpj, "Saida", str(ano_atual), mes_passado)
    ]
    
    destino_pasta = os.path.join(DESTINO_DIR, cod_interno, "entrada_saida")
    for origem_pasta in origens:
        mover_arquivos(origem_pasta, destino_pasta,cod_interno)

# Função genérica para mover arquivos
def mover_arquivos(origem_pasta, destino_pasta, cod_interno):
    xml=0
    if not os.path.exists(origem_pasta):
        print(f"🚫 Pasta não encontrada: {origem_pasta}")
        return

    os.makedirs(destino_pasta, exist_ok=True)

    for arquivo in os.listdir(origem_pasta):
        origem_arquivo = os.path.join(origem_pasta, arquivo)
        destino_arquivo = os.path.join(destino_pasta, arquivo)

        try:
            shutil.move(origem_arquivo, destino_arquivo)
            print(f"✅ Arquivo movido: {arquivo} -> {destino_pasta}")
            xml = xml+1
            
        except Exception as e:
            print(f"❌ Erro ao mover {arquivo}: {e}")
    with open(r'Z:\Logs'+'\\'+'Log.txt', 'a') as arquivo:          
        informacoes = 'Cod:'+cod_interno+' Baixou: '+str(xml)+' XML'
        arquivo.write(informacoes+'\n\n')    
    conn = conectar_sql()
    if not conn:
        return []

    cursor = conn.cursor()
    # Comando SQL de INSERT
    sql = """
    INSERT INTO RelatorioHubSieg (CodInterno, QuantidadeArquivo)  
    VALUES (?, ?)
    """

    # Dados para inserir
    dados = (cod_interno,xml)
    # Executar o INSERT
    cursor.execute(sql, dados)
    conn.commit()

# Fechar a conexão
    cursor.close()
    conn.close()
# Executa o processo
def main():
    empresas = obter_empresas()
    if not empresas:
        print("Nenhuma empresa encontrada.")
        return
    
    conn = conectar_sql()
    if not conn:
        return []
    



    for cnpj, cod_interno in empresas:
        print(f"🔄 Processando CNPJ {cnpj} (CodInterno: {cod_interno})")
        mover_ctes(cnpj, cod_interno)
        mover_nfses(cnpj, cod_interno)
        mover_sat(cnpj, cod_interno)
        mover_entrada_saida(cnpj, cod_interno)

if __name__ == "__main__":
    main()

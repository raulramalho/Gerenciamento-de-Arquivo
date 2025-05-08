import os
import shutil
import pandas as pd

# Caminho para a pasta MODELO ALTERDATA
modelo_alterdata_path = r'Z:\MODELO ALTERDATA'
# Caminho para o diretório Desktop do usuário
desktop_path = r'Z:\ADICIONADOS'

# Função para mover e substituir arquivos e pastas
def mover_pasta(src, dst):
    try:
        # Verifica se o destino existe, caso não, cria a estrutura de diretórios
        if not os.path.exists(dst):
            os.makedirs(dst)

        # Move o conteúdo da pasta para o destino, substituindo arquivos se necessário
        for item in os.listdir(src):
            item_path = os.path.join(src, item)
            dest_path = os.path.join(dst, item)

            # Se já existir um arquivo com o mesmo nome, remove antes de mover
            if os.path.exists(dest_path):
                if os.path.isdir(dest_path):
                    shutil.rmtree(dest_path)  # Remove pasta existente
                else:
                    os.remove(dest_path)  # Remove arquivo existente

            shutil.move(item_path, dest_path)  # Move o novo arquivo/pasta

        # Após mover tudo, remover a pasta vazia
        shutil.rmtree(src)
        print(f"Pasta {src} foi movida para {dst} e removida com sucesso!")
    except Exception as e:
        print(f"Erro ao mover pasta {src} para {dst}: {e}")

# Função para verificar e mover a pasta 'Adicionadas automaticamente'
def verificar_e_mover_pasta(codigo_empresa):
    # Caminho da pasta da empresa
    empresa_path = os.path.join(modelo_alterdata_path, str(codigo_empresa))
    
    if os.path.isdir(empresa_path):
        # Caminhos das pastas entrada_saida e servico
        entrada_saida_path = os.path.join(empresa_path, 'entrada_saida')
        servico_path = os.path.join(empresa_path, 'servico')

        # Verifica e move a pasta 'Adicionadas automaticamente' em entrada_saida
        src_entrada_saida = os.path.join(entrada_saida_path, 'Adicionadas automaticamente')
        if os.path.isdir(src_entrada_saida):
            destino = os.path.join(desktop_path, str(codigo_empresa), 'entrada_saida', 'Adicionadas automaticamente')
            mover_pasta(src_entrada_saida, destino)

        # Verifica e move a pasta 'Adicionadas automaticamente' em servico
        src_servico = os.path.join(servico_path, 'Adicionadas automaticamente')
        if os.path.isdir(src_servico):
            destino = os.path.join(desktop_path, str(codigo_empresa), 'servico', 'Adicionadas automaticamente')
            mover_pasta(src_servico, destino)

# Função para ler os códigos da empresa a partir da linha 5 da planilha
def ler_codigos_empresa(planilha_path):
    # Lê a planilha, ignorando as primeiras 4 linhas (linhas 0 a 3)
    df = pd.read_excel(planilha_path, skiprows=3)
    
    # Supondo que a coluna com os códigos das empresas seja chamada 'CÓD.'
    return df['CÓD.'].dropna().astype(str).tolist()

# Caminho da planilha que contém os códigos das empresas
planilha_path = r'C:\Users\Eunice Ramalho\Desktop\Testando codigo'

# Cria a pasta no Desktop se não existir
if not os.path.exists(desktop_path):
    os.makedirs(desktop_path)

# Lê os códigos das empresas
codigos_empresas = ler_codigos_empresa(planilha_path)

# Para cada código de empresa, verifica e move a pasta 'Adicionadas automaticamente' se existir
for codigo in codigos_empresas:
    verificar_e_mover_pasta(codigo)

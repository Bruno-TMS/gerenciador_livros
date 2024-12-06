
import shutil
from pathlib import Path
import re
import unicodedata 

ORIGINAIS_PATH = Path('originais')
PRODUCAO_PATH = Path('producao')

def criar_pasta_originais():
    try:
        ORIGINAIS_PATH.mkdir(parents=True, exist_ok=True)
        print(f"Diretório '{ORIGINAIS_PATH.name}' criado com sucesso!")
    
    except Exception as e:
        print(f"Erro ao criar o diretório: {e}")


def criar_pasta_producao():
    try:
        PRODUCAO_PATH.mkdir(parents=True, exist_ok=True)
        print(f"Diretório '{PRODUCAO_PATH.name}' criado com sucesso!")
    
    except Exception as e:
        print(f"Erro ao criar o diretório: {e}")


def copiar_arquivos():
    # Definir os diretórios de origem e destino
    origem = ORIGINAIS_PATH
    destino = PRODUCAO_PATH
    
    # Certificar-se de que a pasta de destino existe
    destino.mkdir(parents=True, exist_ok=True)
    
    # Listar todos os arquivos .docx na pasta de origem
    for arquivo in origem.glob("*.docx"):
        # Renomear o arquivo conforme as regras
        nome_renomeado = renomear_arquivo(arquivo.name)
        
        # Caminho completo do arquivo de destino
        caminho_destino = destino / nome_renomeado
        
        # Copiar o arquivo para a pasta de produção
        shutil.copy(arquivo, caminho_destino)
        print(f"Arquivo {arquivo.name} copiado e renomeado para {nome_renomeado}")

def renomear_arquivo(nome):
    # Separa o nome do arquivo da sua extensão
    nome_base, extensao = nome.rsplit('.', 1)
    
    # Remove acentos e caracteres especiais (exceto letras, números e _)
    nome_base = ''.join(c for c in unicodedata.normalize('NFD', nome_base) 
                        if unicodedata.category(c) != 'Mn' and (c.isalnum() or c == ' '))
    
    # Converte para minúsculas
    nome_base = nome_base.lower()
    
    # Substitui múltiplos espaços por apenas 1 underscore
    nome_base = re.sub(r'\s+', '_', nome_base)  # Substitui múltiplos espaços por _
    
    # Substitui todos os pontos (exceto o último) por underscore (_)
    nome_base = nome_base.replace('.', '_')
    
    # Remove qualquer _ antes do ponto da extensão
    nome_base = nome_base.rstrip('_')  # Remove _ no final, se houver
    
    # Junta novamente o nome com a extensão
    return f"{nome_base}.{extensao}"

def listar_arquivos():
    for x in ORIGINAIS_PATH.glob('*.docx'):
        print (x.name)

if __name__ == "__main__":
    pass
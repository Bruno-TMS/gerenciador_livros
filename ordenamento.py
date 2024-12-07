
import shutil
from pathlib import Path
import re
import unicodedata
from docx import Document
import csv
from pprint import pprint as pp

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


def criar_preordenamento():
    'método que vai extrair path de origem, titulo, autor e inserir uma ordem em um csv'
    
    if not PRODUCAO_PATH.exists():
        raise FileNotFoundError('Pasta de produção não existe')
    
    arquivos = {k:{'path': v} for k, v in enumerate(sorted(PRODUCAO_PATH.glob('*.docx')))}
        
    nome_restricao_letras = 250
        
    for k, v in arquivos.items():
        wd = Document(v['path'])
        wd_prgs = wd.paragraphs
            
        prgs = [x for x in wd_prgs if x.text.strip()]
            
        v.setdefault('titulo', prgs[0].text.strip().rstrip(',.'))
            
        if len(prgs[1].text.strip()) <= nome_restricao_letras:
            v.setdefault('autor', prgs[1].text.strip().rstrip(',.'))
        else:
            v.setdefault('autor', '')
            
        campos = ['ordem', 'path', 'titulo', 'autor']
        
        with open(Path(PRODUCAO_PATH,'00_pre_ordenamento_automatico'), mode='w', newline='', encoding='utf-8') as arquivo_csv:
            escritor_csv = csv.DictWriter(arquivo_csv, fieldnames=campos, delimiter=';')
            
            for ordem, info in sorted(arquivos.items()):
                escritor_csv.writerow({
                    'ordem': ordem + 1
                    , 'path': str(Path(info.get('path', '')).name)
                    , 'titulo': info.get('titulo', '')
                    , 'autor': info.get('autor', '')
                    
            })
                
        pp(arquivos)
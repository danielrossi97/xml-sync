import os
import shutil
from datetime import datetime

def listar_arquivos_recentes(pasta, data_minima, arquivos_ja_copiados):
    arquivos_novos = []

    if not os.path.exists(pasta):
        return arquivos_novos

    for arquivo in os.listdir(pasta):
        if not arquivo.lower().endswith('.xml'):
            continue
        if not arquivo.startswith('CTe'):
            continue
        if arquivo.endswith('_Canc.xml'):
            continue  

        caminho_completo = os.path.join(pasta, arquivo)
        data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_completo))

        if data_mod >= data_minima and arquivo not in arquivos_ja_copiados:
            arquivos_novos.append((arquivo, caminho_completo, data_mod))

    return arquivos_novos

def copiar_arquivo(origem, destino_pasta, nome_arquivo):
    os.makedirs(destino_pasta, exist_ok=True)
    destino_completo = os.path.join(destino_pasta, nome_arquivo)
    shutil.copy2(origem, destino_completo)
    return datetime.now()
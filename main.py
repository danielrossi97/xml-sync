import time
import os
from config import CNPJS, PASTA_ORIGEM_BASE, PASTA_DESTINO, RELATORIO_EXCEL, DATA_CORTE
from excel_utils import carregar_registros_existentes, registrar_transferencias
from file_handler import listar_arquivos_recentes, copiar_arquivo

def executar_transferencia():
    print("Iniciando execução...\n")

    df_registros = carregar_registros_existentes(RELATORIO_EXCEL)
    arquivos_ja_copiados = set(df_registros["NomeArquivo"])

    registros_novos = []

    for cnpj in CNPJS:
        pasta_origem = os.path.join(PASTA_ORIGEM_BASE, cnpj)
        arquivos_para_copiar = listar_arquivos_recentes(pasta_origem, DATA_CORTE, arquivos_ja_copiados)

        for nome, caminho, data_mod in arquivos_para_copiar:
            data_transferencia = copiar_arquivo(caminho, PASTA_DESTINO, nome)
            registros_novos.append((cnpj, nome, data_mod, data_transferencia))
            print(f"[{cnpj}] Copiado: {nome} | Modificado: {data_mod} | Transferido: {data_transferencia}")

    if registros_novos:
        registrar_transferencias(RELATORIO_EXCEL, registros_novos)
        print(f"\n{len(registros_novos)} novos registros salvos no relatório.")
    else:
        print("Nenhum novo arquivo para transferir.")

    print("Execução finalizada.\n")

if __name__ == "__main__":
        executar_transferencia()
        print("Fim da execução.")

import os
import sys
from datetime import datetime
from config import CNPJS, PASTA_ORIGEM_BASE, PASTA_DESTINO, RELATORIO_EXCEL
from excel_utils import carregar_registros_existentes, registrar_transferencias
from file_handler import listar_arquivos_recentes, copiar_arquivo, enviar_para_ftp

def get_diretorio_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def carregar_data_corte():
    diretorio_base = get_diretorio_base()
    caminho = os.path.join(diretorio_base, "ultima_execucao.txt")
    if os.path.exists(caminho):
        with open(caminho, "r") as f:
            return datetime.fromisoformat(f.read().strip())
    else:
        return datetime(2025, 7, 27)  # data inicial padrão

def salvar_data_execucao_atual():
    diretorio_base = get_diretorio_base()
    caminho = os.path.join(diretorio_base, "ultima_execucao.txt")
    with open(caminho, "w") as f:
        f.write(datetime.now().isoformat())

def executar_transferencia():
    print("\nIniciando execução...\n")
    data_corte = carregar_data_corte()
    print(f"Usando data de corte: {data_corte}")

    df_registros = carregar_registros_existentes(RELATORIO_EXCEL)
    arquivos_ja_copiados = set(df_registros["NomeArquivo"])

    registros_novos = []
    arquivos_com_falha_ftp = []
    total_enviados = 0

    for cnpj in CNPJS:
        pasta_origem = os.path.join(PASTA_ORIGEM_BASE, cnpj)
        arquivos_para_copiar = listar_arquivos_recentes(pasta_origem, data_corte, arquivos_ja_copiados)

        for nome, caminho, data_mod in arquivos_para_copiar:
            data_transferencia = copiar_arquivo(caminho, PASTA_DESTINO, nome)

            sucesso = enviar_para_ftp(os.path.join(PASTA_DESTINO, nome), nome)
            if sucesso:
                total_enviados += 1
            else:
                arquivos_com_falha_ftp.append(nome)

            registros_novos.append((cnpj, nome, data_mod, data_transferencia))
            print(f"[{cnpj}] Copiado: {nome} | Modificado: {data_mod} | Transferido: {data_transferencia}")

    if registros_novos:
        registrar_transferencias(RELATORIO_EXCEL, registros_novos)

    print(f"\n{total_enviados} arquivos enviados com sucesso via FTP.")

    if arquivos_com_falha_ftp:
        print("\n❌ Arquivos que não foram enviados ao FTP:")
        for nome in arquivos_com_falha_ftp:
            print(f" - {nome}")

    # Log .txt
    diretorio_base = get_diretorio_base()
    caminho_log = os.path.join(diretorio_base, "log_execucao.txt")

    with open(caminho_log, "a", encoding="utf-8") as log:
        log.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ")
        log.write(f"{total_enviados} arquivos enviados com sucesso.\n")
        if arquivos_com_falha_ftp:
            log.write("Falha ao enviar os seguintes arquivos:\n")
            for nome in arquivos_com_falha_ftp:
                log.write(f" - {nome}\n")
        else:
            log.write("Nenhuma falha de envio.\n")

    salvar_data_execucao_atual()
    print("Execução finalizada.\n")

if __name__ == "__main__":
    inicio = datetime.now()
    print("Início:", inicio)

    executar_transferencia()

    fim = datetime.now()
    print("Fim:", fim)
    print("Duração:", fim - inicio)

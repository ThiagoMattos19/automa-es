import os
import shutil
import tkinter


def extrair_arquivos_txt():
    
    pasta_principal = r"c:\Users\Thiago_Mattos\Downloads\Assinados"
    pasta_destino = r"c:\Users\Thiago_Mattos\Documents\teste"
    os.makedirs(pasta_destino, exist_ok=True)

    for raiz, subpastas, arquivos in os.walk(pasta_principal):
        for arquivo in arquivos:
            if arquivo.endswith(".pdf"):
                caminho_arquivo = os.path.join(raiz, arquivo)
                destino_arquivo = os.path.join(pasta_destino, arquivo)

                shutil.move(caminho_arquivo, destino_arquivo)
                print(f"Arquivo movido: {caminho_arquivo} -> {destino_arquivo}")

    print("Todos os arquivos .txt foram consolidados na pasta destino.")


extrair_arquivos_txt()
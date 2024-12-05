import os
import pandas as pd
from docx import Document

caminho_planilha = r"C:\Users\Thiago_Mattos\Downloads\Colaboradores-factorial2.xlsx"  
caminho_modelo = r"C:\Users\Thiago_Mattos\Downloads\Termo de Responsabilidade1.docx"
pasta_saida = r"C:\Users\Thiago_Mattos\Downloads\termos_colaboradores"

df = pd.read_excel(caminho_planilha, header=1)

df.columns = df.columns.str.strip()

print("Colunas disponíveis:", df.columns.tolist())

def formatar_cpf(cpf):
    try:
        cpf = str(int(cpf))  
        cpf = cpf.zfill(11)  
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    except Exception as e:
        print(f"Erro ao formatar CPF: {e}")
        return "CPF inválido"  

os.makedirs(pasta_saida, exist_ok=True)

def substituir_placeholder_em_runs(paragrafo, placeholder, texto, negrito=False):
    if placeholder in paragrafo.text:
        for run in paragrafo.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, texto)
                if negrito:
                    run.bold = True

if not {"Unnamed: 5", "Número do documento de identidade"}.issubset(df.columns):
    print("Faltando alguma coluna no arquivo.")
    print("Colunas encontradas:", df.columns.tolist())  
else:
    
    for _, row in df.iterrows():
        nome_completo = row["Unnamed: 5"]
        cpf_formatado = formatar_cpf(row["Número do documento de identidade"])
        doc = Document(caminho_modelo)
        for paragrafo in doc.paragraphs:
            substituir_placeholder_em_runs(paragrafo, "{{NOME}}", nome_completo, negrito=True)
            substituir_placeholder_em_runs(paragrafo, "{{CPF}}", cpf_formatado, negrito=True)

        nome_arquivo = f"{pasta_saida}/termo_{nome_completo.replace(' ', '_')}.docx"
        
        doc.save(nome_arquivo)

    print("Documentos gerados com sucesso!")

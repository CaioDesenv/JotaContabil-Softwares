import pandas as pd
import pyodbc
from tkinter import filedialog, Tk
from datetime import datetime

# Função para solicitar datas do usuário
def solicitar_datas():
    data_inicial = input("Digite a data inicial (aaaa-mm-dd): ")
    data_final = input("Digite a data final (aaaa-mm-dd): ")
    return data_inicial, data_final

# Função para abrir o seletor de arquivos do Excel
def selecionar_arquivo_excel():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)  # Garante que a janela apareça na frente
    caminho_arquivo = r"C:\Users\Jota11\AppData\Local\anaconda3\envs\automacao_odbc_dominio_nfse\Planilhas\Empresas_Buscar.xlsx"
    root.destroy()
    return caminho_arquivo

# Função para carregar os códigos das empresas do Excel
def carregar_codigos_empresa(caminho_arquivo):
    df_excel = pd.read_excel(caminho_arquivo, header=0, dtype=str)
    codigos = df_excel["Codigo"].dropna().unique()
    return codigos

# Função para executar consulta no banco de dados
def consultar_dados_empresa(codigo_empresa, data_inicial, data_final):
    conn = pyodbc.connect("DSN=REMJOTA;UID=REMOTO;PWD=Jota@9960")
    cursor = conn.cursor()

    query = f"""
    SELECT
        e.codi_emp AS "Codigo Empresa",
        e.nume_ser AS "Cod Nota",
        e.vcon_ser AS "Valor Contabil",
        CASE 
            WHEN e.situacao_ser = 0 THEN 'N'
            WHEN e.situacao_ser = 2 THEN 'S'
            WHEN e.situacao_ser = 8 THEN 'NFS-e Numeracao inutilizada'
            ELSE CAST(e.situacao_ser AS VARCHAR)
        END AS "Situacao",
        e.codi_usu AS "Usuario"
    FROM bethadba.efservicos e
    WHERE e.codi_emp = '{codigo_empresa}'
    AND e.dser_ser >= '{data_inicial}'
    AND e.dser_ser <= '{data_final}'
    AND e.codi_esp = 39
    ORDER BY e.nume_ser
    """

    cursor.execute(query)
    rows = cursor.fetchall()
    colunas = [column[0] for column in cursor.description]
    df_resultado = pd.DataFrame.from_records(rows, columns=colunas)

    cursor.close()
    conn.close()

    return df_resultado

# Função para salvar os dados no CSV final
def salvar_em_csv(df, caminho_csv):
    df.to_csv(caminho_csv, sep=';', mode='a', index=False, header=False, encoding='utf-8')

# Execução do processo
def main():
    data_inicial, data_final = solicitar_datas()
    caminho_excel = selecionar_arquivo_excel()
    codigos_empresas = carregar_codigos_empresa(caminho_excel)

    caminho_csv_saida = f"relatorio-nfse-dominio.csv"

    for codigo in codigos_empresas:
        print(f"Consultando empresa {codigo}...")
        df_empresa = consultar_dados_empresa(codigo, data_inicial, data_final)
        salvar_em_csv(df_empresa, caminho_csv_saida)
        print(f"Dados da empresa {codigo} salvos com sucesso.\n")

if __name__ == "__main__":
    main()

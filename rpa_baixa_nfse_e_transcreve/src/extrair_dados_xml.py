import xml.etree.ElementTree as ET
import pandas as pd
import locale
import os
from datetime import datetime
import logging

# Configuração do log
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Configuração da localidade para formato brasileiro
locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")


# Função para formatar valores numéricos
def format_value(value):
    try:
        return locale.format_string("%.2f", float(value), grouping=True)
    except (ValueError, TypeError):
        return "0.00"

# Função para somar os valores de um único XML
def sum_xml_values(matrix):
    """
    Calcula os totais das colunas numéricas de uma matriz, abrangendo todas as colunas necessárias.
    """
    if not matrix:
        return [0.0] * 9  # Retorna uma lista de zeros para evitar erros

    # Inicializa os totais para as colunas numéricas: ValorPis até BaseCalculo (10 colunas)
def sum_xml_values(matrix):
    if not matrix:
        return [0.0] * 9

    sum_values = [0.0] * 7
    for row in matrix:
        for i in range(7):  # Somando 9 colunas, excluindo Codigo, Numero e DataEmissao
            try:
                value = row[i+3] if i+3 < len(row) else "0"
                value = value.replace(".", "").replace(",", ".") if isinstance(value, str) else value
                sum_values[i] += float(value)
            except (ValueError, TypeError, IndexError):
                logging.warning(f"Valor inválido ignorado na coluna {i+3}: {value}")

    return sum_values

def parse_xml(xml_file, folder_name, matriz, totals):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        xml_matrix = []  # Matriz temporária para armazenar os dados deste XML
        for nfse in root.findall(".//ListaNfse/tcNfCompNfse/Nfse/InfNfse"):
            nfse_data = {}
            nfse_data["Codigo"] = folder_name.replace("-", "")
            nfse_data["Numero"] = (
                nfse.find("Numero").text if nfse.find("Numero") is not None else "N/A"
            )
            nfse_data["DataEmissao"] = (
                nfse.find("DataEmissao").text
                if nfse.find("DataEmissao") is not None
                else None
            )
            
            # Extração de valores de impostos e liquido da NFSe
            valores = nfse.find("ValoresNfse")
            if valores is not None:
                nfse_data["PIS-RET"] = format_value(
                    valores.find("ValorPis").text
                    if valores.find("ValorPis") is not None
                    else 0
                )
                nfse_data["COFINS-R"] = format_value(
                    valores.find("ValorCofins").text
                    if valores.find("ValorCofins") is not None
                    else 0
                )
                nfse_data["INSS"] = format_value(
                    valores.find("ValorInss").text
                    if valores.find("ValorInss") is not None
                    else 0
                )
                nfse_data["IRRF"] = format_value(
                    valores.find("ValorIr").text
                    if valores.find("ValorIr") is not None
                    else 0
                )
                nfse_data["CSOC-RET"] = format_value(
                    valores.find("ValorCsll").text
                    if valores.find("ValorCsll") is not None
                    else 0
                )
                nfse_data["ISS"] = format_value(
                    valores.find("ValorIss").text
                    if valores.find("ValorIss") is not None
                    else 0
                )
            else:
                nfse_data["PIS-RET"] = nfse_data["COFINS-R"] = nfse_data["INSS"] = nfse_data["IRRF"] = nfse_data["CSOC-RET"] = nfse_data["ISS"] = "0.00"

            # Nova seção: Extração de valores da tag <DeclaracaoPrestacaoServico>
            declaracao = nfse.find(
                ".//DeclaracaoPrestacaoServico/InfDeclaracaoPrestacaoServico"
            )

            # Extração da tag <Status>
            status = (
                declaracao.find(".//Rps/Status") if declaracao is not None else None
            )
            if status is not None:
                status_value = status.text
                if status_value == "1":
                    status_value = "Ativa"
                elif status_value == "2":
                    status_value = "Cancelada"
            else:
                status_value = "N/A"

            nfse_data["Status"] = status_value

            if declaracao is not None:
                servico_valores = declaracao.findall(".//ListaServicos/Servico/Valores")
                maior_valores_nfse = 0.0  # Inicializa a variável para somar os valores
                
                for valores in servico_valores:
                    valor_servicos = float(valores.find("ValorServicos").text if valores.find("ValorServicos") is not None else 0)
                    maior_valores_nfse += valor_servicos  # Soma os valores de todas as tags <ValorServicos>

            else:
                nfse_data["ValorTotalNota"] = nfse_data["ValorDeducoes"] = "0,00"

            # Verifica se a NFSe está cancelada
            if status_value == "S":
                maior_valores_nfse = 0.0  # Zera o valor se a NFSe estiver cancelada
                nfse_data["PIS-RET"] = nfse_data["COFINS-R"] = nfse_data["INSS"] = nfse_data["IRRF"] = nfse_data["CSOC-RET"] = nfse_data["ISS"] = nfse_data["ISS"] = "0,00"
                
            nfse_data["ValorTotalNota"] = format_value(maior_valores_nfse)
            nfse_data["Aliquota"] = format_value(
                valores.find("Aliquota").text
                if valores.find("Aliquota") is not None
                else 0
            )
            # Segunda extração: captura do CNPJ do prestador
            prestador = nfse.find(
                ".//PrestadorServico/IdentificacaoPrestador/CpfCnpj/Cnpj"
            )
            nfse_data["CNPJ_Prestador"] = (
                prestador.text if prestador is not None else "N/A"
            )

            # Extração da Razão Social do prestador
            razao_social = nfse.find(".//PrestadorServico/RazaoSocial")
            nfse_data["RazaoSocial_Prestador"] = (
                razao_social.text if razao_social is not None else "N/A"
            )

            # Extração do CNPJ ou CPF do tomador
            cpf_cnpj_tomador = nfse.find(
                ".//TomadorServico/IdentificacaoTomador/CpfCnpj/Cnpj"
            )
            if cpf_cnpj_tomador is None:
                cpf_cnpj_tomador = nfse.find(
                    ".//TomadorServico/IdentificacaoTomador/CpfCnpj/Cpf"
                )
            nfse_data["CPF_CNPJ_Tomador"] = (
                cpf_cnpj_tomador.text if cpf_cnpj_tomador is not None else "N/A"
            )

            # Extração da Razão Social do tomador
            razao_social_tomador = nfse.find(".//TomadorServico/RazaoSocial")
            nfse_data["RazaoSocial_Tomador"] = (
                razao_social_tomador.text if razao_social_tomador is not None else "N/A"
            )

            # Extração da Razão Social do tomador
            razao_social_tomador = nfse.find(".//TomadorServico/RazaoSocial")
            nfse_data["RazaoSocial_Tomador"] = (
                razao_social_tomador.text if razao_social_tomador is not None else "N/A"
            )
            
            # Adiciona os dados na matriz temporária
            xml_matrix.append(
                [
                    nfse_data["Codigo"],
                    nfse_data["Numero"],
                    nfse_data["DataEmissao"],
                    nfse_data["PIS-RET"],
                    nfse_data["COFINS-R"],
                    nfse_data["INSS"],
                    nfse_data["ISS"],
                    nfse_data["IRRF"],
                    nfse_data["CSOC-RET"],
                    nfse_data["ValorTotalNota"],
                    nfse_data["Aliquota"],
                    nfse_data["Status"],
                    nfse_data["CNPJ_Prestador"],
                    nfse_data["RazaoSocial_Prestador"],
                    nfse_data["CPF_CNPJ_Tomador"],
                    nfse_data["RazaoSocial_Tomador"],
                    
                ]
            )

        # Calcula os totais para este XML e adiciona à lista de totais
        xml_totals = sum_xml_values(xml_matrix)
        formatted_totals = [format_value(value) for value in xml_totals]
        totals.append([folder_name] + formatted_totals)

        # Adiciona os dados do XML à matriz geral
        matriz.extend(xml_matrix)
    except Exception as e:
        logging.error(f"Erro ao processar o arquivo {xml_file}: {e}")


# Função para processar as pastas e XMLs
def process_folders(base_path):
    matriz = []
    totals = []
    for root, _, files in os.walk(base_path):
        for file in files:
            if file.endswith(".xml"):
                file_path = os.path.join(root, file)
                folder_name = os.path.basename(root)
                parse_xml(file_path, folder_name, matriz, totals)
    return matriz, totals


# Função principal
if __name__ == "__main__":
    base_path = r"C:\Users\Jota11\AppData\Local\anaconda3\envs\rpa_baixa_nfse_e_transcreve\src\download"  # Substituir pelo caminho correto
    logging.info("Iniciando processamento de pastas...")

    matriz, totals = process_folders(base_path)
    logging.info(f"Processamento concluído. Total de registros: {len(matriz)}")

    # Atualização do DataFrame com nova coluna
    df_data = pd.DataFrame(
    matriz,
    columns=[
        "Codigo",
        "Numero",
        "DataEmissao",
        "PIS-RET",
        "COFINS-R",
        "INSS",
        "ISS",
        "IRRF",
        "CSOC-RET",
        "ValorTotalNota",
        "Aliquota",
        "Status",
        "CNPJ_Prestador",
        "RazaoSocial_Prestador",
        "CPF_CNPJ_Tomador",
        "RazaoSocial_Tomador",
    ],
)

    df_totals = pd.DataFrame(
    totals,
    columns=[
        "Codigo",  # Esta coluna vem do folder_name
        "PIS-RET",
        "COFINS-R",
        "INSS",
        "ISS",
        "IRRF",
        "CSOC-RET",
        "ValorTotalNota",
    ],
)

    # Salvar resultados em Excel
    try:
        output_file = r"C:\Users\Jota11\AppData\Local\anaconda3\envs\rpa_baixa_nfse_e_transcreve\Planilhas\relatorio_nfse_prefeitura.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_data.to_excel(writer, sheet_name="PREFEITURA", index=False)
            df_totals.to_excel(writer, sheet_name="Totais", index=False)
        logging.info(f"Resultados salvos em {output_file}.")
    except ImportError:
        logging.error(
            "O módulo 'openpyxl' não está instalado. Instale-o usando 'pip install openpyxl'."
        )
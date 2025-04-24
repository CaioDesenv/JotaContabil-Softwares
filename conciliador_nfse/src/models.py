import pandas as pd
import re
from datetime import datetime


class FormatUtils:
    """Utilitários para formatação de CPF/CNPJ e valores."""
    
    @staticmethod
    def formatar_cpf_cnpj(valor):
        valor = re.sub(r'\D', '', str(valor))
        if len(valor) == 11:
            return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
        elif len(valor) == 14:
            return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
        return valor

    @staticmethod
    def limpar_valor(valor):
        if pd.isna(valor):
            return 0.0
        valor = str(valor).strip()
        if ',' in valor:
            # Assume formato brasileiro: remove separador de milhar e troca vírgula por ponto
            valor = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor)
        except ValueError:
            return 0.0

    @staticmethod
    def formatar_valor_brl(valor):
        try:
            # Formata o número com duas casas decimais, utilizando vírgula como separador decimal
            formatted = f"{valor:,.2f}"
            formatted = formatted.replace(',', 'X').replace('.', ',').replace('X', '.')
            return formatted
        except Exception:
            return str(valor)

class DataImporter:
    """Responsável pela importação dos dados das planilhas e arquivos CSV."""
    
    @staticmethod
    def importar_prefeitura(caminho):
        df = pd.read_excel(caminho, sheet_name='PREFEITURA')
        df.columns = df.columns.str.strip()
        df["Codigo"] = df["Codigo"].astype(str)
        df["Numero"] = df["Numero"].astype(str)
        df["Chave"] = df["Codigo"] + "-" + df["Numero"]

        colunas_numericas = ["ValorTotalNota", "Aliquota", "ISS", "IRRF", "INSS", "PIS-RET", "COFINS-R", "CSOC-RET"]
        for col in colunas_numericas:
            if col in df.columns:
                df[col] = df[col].apply(FormatUtils.limpar_valor)

        df["Valor"] = df["ValorTotalNota"]
        df["Situacao"] = df["Status"].map({"N": "Ativa", "S": "Cancelada"})
        df["Aliquota"] = df["Aliquota"] / 100
        df["DataEmissao"] = pd.to_datetime(df["DataEmissao"].str.split("T").str[0], format="%Y-%m-%d")
        df["CPF_CNPJ_Tomador"] = df["CPF_CNPJ_Tomador"].apply(FormatUtils.formatar_cpf_cnpj)
        df["CNPJ_Prestador"] = df["CNPJ_Prestador"].apply(FormatUtils.formatar_cpf_cnpj)

        colunas_pref = [
            "Chave", "Codigo", "Numero", "DataEmissao", "PIS-RET", "COFINS-R",
            "INSS", "ISS", "IRRF", "CSOC-RET", "Valor", "Aliquota", "Status",
            "CNPJ_Prestador", "RazaoSocial_Prestador", "CPF_CNPJ_Tomador", "RazaoSocial_Tomador"
        ]
        return df[colunas_pref].copy()

    @staticmethod
    def importar_dominio(caminho):
        df = pd.read_csv(caminho, sep=';', dtype=str)
        df.columns = df.columns.str.strip()
        df["Codigo"] = df["Codigo"].astype(str)
        df["Numero"] = df["Numero"].astype(str)
        df["Chave"] = df["Codigo"] + "-" + df["Numero"]
        df["Valor"] = df["Valor"].apply(FormatUtils.limpar_valor)
        df["Situacao"] = df["Situacao"].map({"N": "Ativa", "S": "Cancelada"})
        # Presume que o CSV possui a coluna "Usuario"
        return df

    @staticmethod
    def importar_responsaveis(caminho):
        df = pd.read_excel(caminho)
        df.columns = df.columns.str.strip()
        df["ID"] = df["ID"].astype(str)
        return df

class Reconciliator:
    """Contém as regras de negócio para realizar a conciliação e gerar o relatório."""
    
    @staticmethod
    def gerar_mescla(df_pref, df_dom):
        chaves_pref = set(df_pref["Chave"])
        chaves_dom = set(df_dom["Chave"])
        chaves_mescla = sorted(chaves_pref.union(chaves_dom))
        dados_mescla = []
        for chave in chaves_mescla:
            if chave in chaves_pref and chave in chaves_dom:
                origem = "AMBAS"
            elif chave in chaves_pref:
                origem = "PREFEITURA"
            else:
                origem = "DOMINIO"
            dados_mescla.append((chave, origem))
        return pd.DataFrame(dados_mescla, columns=["Chave", "Origem"])

    @staticmethod
    def verificar_divergencias(row):
        divergencias = []
        # Extrai "empresa" e "nota" a partir da "Chave"
        try:
            empresa, nota = row['Chave'].split('-', 1)
        except Exception:
            empresa, nota = row['Chave'], ''
        
        if row['Origem'] == 'AMBAS':
            status_pref = str(row['Status_Pref']).strip().lower()
            situacao_dom = str(row['Situacao_Dom']).strip().lower()
            if not (status_pref == "cancelada" and situacao_dom == "cancelada"):
                if round(row['Valor_Pref'], 2) != round(row['Valor_Dom'], 2):
                    divergencias.append(
                        f"Valor divergente (Pref: {FormatUtils.formatar_valor_brl(row['Valor_Pref'])} vs Dom: {FormatUtils.formatar_valor_brl(row['Valor_Dom'])})"
                    )
            if status_pref != situacao_dom:
                divergencias.append(
                    f"Status divergente (Pref: {row['Status_Pref']} vs Dom: {row['Situacao_Dom']})"
                )
        elif row['Origem'] == 'PREFEITURA':
            divergencias.append(
                f"Status divergente (Empresa: {empresa} Nota: {nota}\n           Pref: {row['Status_Pref']} vs Dom: Não importado\n           Valor {FormatUtils.formatar_valor_brl(row['Valor_Pref'])})"
            )
        elif row['Origem'] == 'DOMINIO':
            divergencias.append(
                f"Status divergente (Empresa: {empresa} Nota: {nota}\n           Pref: Não consta vs Dom: {row['Situacao_Dom']}\n           Valor {FormatUtils.formatar_valor_brl(row['Valor_Dom'])})"
            )
        return "; ".join(divergencias)

    @staticmethod
    def verificar_sequencia_notas(df_pref, mes_referencia=None):
        """
        Verifica a sequência de notas na Prefeitura.
        Se mes_referencia for informado (no formato MM/AAAA), filtra o DataFrame para este mês.
        """
        resultados = []
        if not pd.api.types.is_datetime64_any_dtype(df_pref["DataEmissao"]):
            df_pref["DataEmissao"] = pd.to_datetime(df_pref["DataEmissao"], errors='coerce')
        df_pref["Ano"] = df_pref["DataEmissao"].dt.year
        df_pref["Mes"] = df_pref["DataEmissao"].dt.month

        # Filtra para o mês de referência, se informado
        if mes_referencia is not None:
            try:
                referencia = pd.to_datetime("01/" + mes_referencia, format="%d/%m/%Y")
                ref_ano = referencia.year
                ref_mes = referencia.month
                df_pref = df_pref[(df_pref["Ano"] == ref_ano) & (df_pref["Mes"] == ref_mes)]
            except Exception:
                raise Exception("Formato de mês referência inválido. Use MM/AAAA.")

        grupos = df_pref.groupby(["Codigo", "Ano", "Mes"])
        for (codigo, ano, mes), grupo in grupos:
            try:
                notas_int = sorted([int(n) for n in grupo["Numero"].unique() if n.isdigit()])
            except Exception:
                continue
            if not notas_int:
                continue
            min_val = notas_int[0]
            max_val = notas_int[-1]
            esperado = set(range(min_val, max_val + 1))
            atual = set(notas_int)
            faltantes = sorted(esperado - atual)
            for nota_falta in faltantes:
                msg = f"Foi encontrado um pulo de notas da empresa: {codigo} no mês {mes}/{ano}. A nota faltante é: {nota_falta}"
                resultados.append({
                    "Empresa": codigo,
                    "Ano": ano,
                    "Mes": mes,
                    "Missing_Nota": nota_falta,
                    "Divergencia": msg
                })
        if resultados:
            return pd.DataFrame(resultados)
        else:
            return pd.DataFrame(columns=["Empresa", "Ano", "Mes", "Missing_Nota", "Divergencia"])

    @staticmethod
    def verificar_sequencia_entre_meses(df_anterior, df_atual):
        """
        Verifica a sequência entre dois meses, comparando a última nota do mês anterior
        com a primeira nota do mês atual para cada empresa.
        """
        # Garantindo que as colunas de data estejam no formato datetime e criando Ano e Mes
        if not pd.api.types.is_datetime64_any_dtype(df_anterior["DataEmissao"]):
            df_anterior["DataEmissao"] = pd.to_datetime(df_anterior["DataEmissao"], errors='coerce')
        if not pd.api.types.is_datetime64_any_dtype(df_atual["DataEmissao"]):
            df_atual["DataEmissao"] = pd.to_datetime(df_atual["DataEmissao"], errors='coerce')
        df_anterior["Ano"] = df_anterior["DataEmissao"].dt.year
        df_anterior["Mes"] = df_anterior["DataEmissao"].dt.month
        df_atual["Ano"] = df_atual["DataEmissao"].dt.year
        df_atual["Mes"] = df_atual["DataEmissao"].dt.month

        resultados = []
        # Verifica para cada empresa que esteja presente em ambos os arquivos
        empresas = set(df_anterior["Codigo"]).intersection(set(df_atual["Codigo"]))
        for empresa in empresas:
            df_ant = df_anterior[df_anterior["Codigo"] == empresa]
            df_atl = df_atual[df_atual["Codigo"] == empresa]
            try:
                notas_ant = [int(n) for n in df_ant["Numero"].unique() if n.isdigit()]
            except Exception:
                continue
            if not notas_ant:
                continue
            max_ant = max(notas_ant)
            try:
                notas_atl = [int(n) for n in df_atl["Numero"].unique() if n.isdigit()]
            except Exception:
                continue
            if not notas_atl:
                continue
            min_atl = min(notas_atl)
            if min_atl != max_ant + 1:
                msg = f"Sequência incorreta: última nota do mês anterior é {max_ant} e a primeira nota do mês atual é {min_atl}"
                mes_ant = df_ant["Mes"].iloc[0]
                mes_atl = df_atl["Mes"].iloc[0]
                resultados.append({
                    "Empresa": empresa,
                    "Mes_Anterior": mes_ant,
                    "Ultima_Nota": max_ant,
                    "Mes_Atual": mes_atl,
                    "Primeira_Nota": min_atl,
                    "Divergencia": msg
                })
        if resultados:
            return pd.DataFrame(resultados, columns=["Empresa", "Mes_Anterior", "Ultima_Nota", "Mes_Atual", "Primeira_Nota", "Divergencia"])
        else:
            return pd.DataFrame(columns=["Empresa", "Mes_Anterior", "Ultima_Nota", "Mes_Atual", "Primeira_Nota", "Divergencia"])

    @staticmethod
    def salvar_excel(df_pref, df_dom, df_mescla, df_sequencia, caminho_saida):
        with pd.ExcelWriter(caminho_saida, engine='xlsxwriter', datetime_format='dd/mm/yyyy',
                            engine_kwargs={'options': {'strings_to_numbers': True}}) as writer:
            df_pref.to_excel(writer, sheet_name='PREFEITURA', index=False)
            df_dom.to_excel(writer, sheet_name='DOMINIO', index=False)
            df_mescla.to_excel(writer, sheet_name='MESCLA', index=False)
            df_sequencia.to_excel(writer, sheet_name='SEQUENCIA-PREFEITURA', index=False)

            workbook = writer.book
            formato_moeda = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
            formato_porcentagem = workbook.add_format({'num_format': '0%', 'align': 'center'})
            formato_data = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center'})
            formato_cpf_cnpj = workbook.add_format({'align': 'center'})

            # Formatação da aba PREFEITURA
            ws_pref = writer.sheets['PREFEITURA']
            ws_pref.set_column('D:D', 12, formato_data)
            ws_pref.set_column('K:K', 12, formato_moeda)
            ws_pref.set_column('L:L', 10, formato_porcentagem)
            ws_pref.set_column('M:M', 10)
            ws_pref.set_column('N:N', 20, formato_cpf_cnpj)
            ws_pref.set_column('P:P', 20, formato_cpf_cnpj)

            # Formatação da aba DOMINIO
            ws_dom = writer.sheets['DOMINIO']
            ws_dom.set_column('C:C', 12, formato_moeda)
            ws_dom.set_column('D:D', 12)

            # Formatação da aba MESCLA
            ws_mescla = writer.sheets['MESCLA']
            ws_mescla.set_column('A:A', 20)
            ws_mescla.set_column('B:B', 20)
            ws_mescla.set_column('C:C', 20)
            ws_mescla.set_column('D:D', 20)
            ws_mescla.set_column('E:E', 20, formato_moeda)
            ws_mescla.set_column('F:F', 20, formato_moeda)
            ws_mescla.set_column('G:G', 20)
            ws_mescla.set_column('H:H', 20)
            ws_mescla.set_column('I:I', 40)

            # Formatação da aba SEQUENCIA-PREFEITURA
            ws_seq = writer.sheets['SEQUENCIA-PREFEITURA']
            ws_seq.set_column('A:A', 15)
            ws_seq.set_column('B:B', 15)
            ws_seq.set_column('C:C', 15)
            ws_seq.set_column('D:D', 15)
            ws_seq.set_column('E:E', 15)
            ws_seq.set_column('F:F', 60)

from models import DataImporter, Reconciliator
import pandas as pd
from datetime import datetime

class ReconciliationController:
    def executar_conciliacao(self, caminho_pref, caminho_dom, caminho_resp, mes_referencia,
                              caminho_mes_ant=None, caminho_mes_atual=None):
        try:
            # Importação dos dados para a conciliação padrão
            df_pref = DataImporter.importar_prefeitura(caminho_pref)
            df_dom  = DataImporter.importar_dominio(caminho_dom)
            df_resp = DataImporter.importar_responsaveis(caminho_resp)

            # Geração da mescla base
            df_mescla = Reconciliator.gerar_mescla(df_pref, df_dom)

            # Incorporação dos responsáveis (merge)
            df_mescla["Codigo"] = df_mescla["Chave"].str.split("-").str[0]
            df_mescla = pd.merge(df_mescla, df_resp, left_on="Codigo", right_on="ID", how="left")
            df_mescla.drop(["Codigo", "ID"], axis=1, inplace=True)

            # Incorporação da coluna "Usuario" do Domínio
            df_usuario = df_dom[["Chave", "Usuario"]].drop_duplicates()
            df_mescla = pd.merge(df_mescla, df_usuario, on="Chave", how="left")

            # Incorporação dos valores e situação do Domínio
            df_dom_extras = df_dom[["Chave", "Valor", "Situacao"]].drop_duplicates()
            df_dom_extras.rename(columns={"Valor": "Valor_Dom", "Situacao": "Situacao_Dom"}, inplace=True)
            df_mescla = pd.merge(df_mescla, df_dom_extras, on="Chave", how="left")

            # Incorporação dos valores e status da Prefeitura
            df_pref_extras = df_pref[["Chave", "Valor", "Status"]].drop_duplicates()
            df_pref_extras.rename(columns={"Valor": "Valor_Pref", "Status": "Status_Pref"}, inplace=True)
            df_mescla = pd.merge(df_mescla, df_pref_extras, on="Chave", how="left")

            # Reordenação e verificação de divergências internas (dentro do mesmo mês)
            df_mescla = df_mescla[["Chave", "Origem", "Responsavel", "Usuario", 
                                   "Valor_Pref", "Valor_Dom", "Status_Pref", "Situacao_Dom"]]
            df_mescla['Divergencia'] = df_mescla.apply(Reconciliator.verificar_divergencias, axis=1)
            df_mescla = df_mescla[["Chave", "Origem", "Responsavel", "Usuario", 
                                   "Valor_Pref", "Valor_Dom", "Status_Pref", "Situacao_Dom", "Divergencia"]]

            # Verificação da sequência de notas para o mês de referência
            df_sequencia = Reconciliator.verificar_sequencia_notas(df_pref, mes_referencia)

            # Verificação da sequência entre meses, se os arquivos forem informados
            if caminho_mes_ant and caminho_mes_atual:
                df_mes_ant = DataImporter.importar_prefeitura(caminho_mes_ant)
                df_mes_atual = DataImporter.importar_prefeitura(caminho_mes_atual)
                df_cross = Reconciliator.verificar_sequencia_entre_meses(df_mes_ant, df_mes_atual)
                # Concatenando os resultados, se houver divergências entre meses
                if not df_cross.empty:
                    df_sequencia = pd.concat([df_sequencia, df_cross], ignore_index=True)

            # Geração do nome do arquivo de saída com a data atual
            data_hoje = datetime.today().strftime('%Y-%m-%d')
            nome_saida = f"relatorio_conciliado_{data_hoje}.xlsx"
            Reconciliator.salvar_excel(df_pref, df_dom, df_mescla, df_sequencia, nome_saida)

            return nome_saida
        except Exception as e:
            raise Exception(f"Erro ao processar: {str(e)}")
        
    def executar_verificacao_sequencia_entre_meses(self, caminho_mes_ant, caminho_mes_atual):
        try:
            df_mes_ant = DataImporter.importar_prefeitura(caminho_mes_ant)
            df_mes_atual = DataImporter.importar_prefeitura(caminho_mes_atual)
            df_cross = Reconciliator.verificar_sequencia_entre_meses(df_mes_ant, df_mes_atual)

            if df_cross.empty:
                raise Exception("Nenhuma divergência de sequência foi encontrada entre os meses.")

            data_hoje = datetime.today().strftime('%Y-%m-%d')
            nome_saida = f"verificacao_entre_meses_{data_hoje}.xlsx"
            with pd.ExcelWriter(nome_saida, engine='xlsxwriter') as writer:
                df_cross.to_excel(writer, sheet_name='SEQUENCIA-PREFEITURA', index=False)

            return nome_saida
        except Exception as e:
            raise Exception(f"Erro na verificação de sequência entre meses: {str(e)}")

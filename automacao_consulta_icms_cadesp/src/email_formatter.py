# email_formatter.py
from datetime import datetime
from typing import Dict, Any, Tuple

class EmailFormatter:
    """
    Classe responsável por formatar e-mails profissionais com suporte a HTML e texto simples.
    """
    
    @staticmethod
    def formatar_email_ocorrencias_cadesp(dados: Dict[int, Tuple]) -> Dict[str, str]:
        """
        Formata um email de ocorrências do CADesp-SP em formatos HTML e texto simples.
        
        Args:
            dados: Dicionário com linha como chave e dados da empresa como valor
        
        Returns:
            Dicionário contendo assunto, versão HTML e versão texto do email
        """
        # Data atual formatada
        data_hoje = datetime.now().strftime("%d/%m/%Y")
        
        # Criar versão HTML do email
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333333;
                }}
                .container {{
                    max-width: 800px;
                    margin: 0 auto;
                    padding: 20px;
                }}
                .header {{
                    border-bottom: 2px solid #0066cc;
                    padding-bottom: 10px;
                    margin-bottom: 20px;
                }}
                .footer {{
                    margin-top: 30px;
                    padding-top: 15px;
                    border-top: 1px solid #dddddd;
                    font-size: 14px;
                    color: #666666;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 20px 0;
                }}
                th, td {{
                    padding: 12px 15px;
                    border: 1px solid #dddddd;
                    text-align: left;
                }}
                th {{
                    background-color: #f2f2f2;
                    font-weight: bold;
                }}
                tr:nth-child(even) {{
                    background-color: #f9f9f9;
                }}
                .highlight {{
                    background-color: #ffeeee;
                }}
                .status {{
                    font-weight: bold;
                    color: #cc0000;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h2>Relatório de Ocorrências - CADesp-SP</h2>
                    <p>Data: {data_hoje}</p>
                </div>
                
                <p>Prezados,</p>
                
                <p>Durante o processamento automático de verificação do status cadastral no sistema CADesp-SP, 
                foram identificadas as seguintes <strong>ocorrências</strong> que requerem atenção:</p>
                
                <table>
                    <tr>
                        <th>Linha</th>
                        <th>Empresa</th>
                        <th>CNPJ</th>
                        <th>Status</th>
                        <th>Ocorrência</th>
                    </tr>
        """
        
        # Adicionar cada empresa como uma linha na tabela
        for linha_numero, dados_linha in dados.items():
            nome_empresa = dados_linha[0] if dados_linha[0] and dados_linha[0] != "None" else "Empresa sem nome"
            numero_cadastro = dados_linha[1] if len(dados_linha) > 1 else "-"
            cnpj = dados_linha[2] if len(dados_linha) > 2 else "-"
            status = dados_linha[3] if len(dados_linha) > 3 else "-"
            ocorrencia = dados_linha[4] if len(dados_linha) > 4 else "-"
            
            html_content += f"""
                    <tr class="highlight">
                        <td>{linha_numero}</td>
                        <td><strong>{nome_empresa}</strong></td>
                        <td>{cnpj}</td>
                        <td class="status">{status}</td>
                        <td>{ocorrencia}</td>
                    </tr>
            """
        
        # Fechar a tabela e adicionar o rodapé
        html_content += """
                </table>
                
                <p>Recomendamos a verificação detalhada destas situações para evitar possíveis 
                penalidades ou restrições fiscais.</p>
                
                <p>Este é um email automático gerado pelo sistema de monitoramento. Em caso de dúvidas, 
                favor contatar o departamento fiscal.</p>
                
                <div class="footer">
                    <p>Atenciosamente,<br>
                    <strong>Departamento de Automação Fiscal</strong><br>
                    JotaContábil | Soluções Contábeis Inteligentes<br>
                    <a href="https://www.jotacontabil.com.br">www.jotacontabil.com.br</a>
                    </p>
                </div>
            </div>
        </body>
        </html>
        """
        
        # Criar versão em texto simples para clientes que não suportam HTML
        texto_simples = f"""
Relatório de Ocorrências - CADesp-SP
Data: {data_hoje}

Prezados,

Durante o processamento automático de verificação do status cadastral no sistema CADesp-SP, 
foram identificadas as seguintes ocorrências que requerem atenção:

"""
        
        for linha_numero, dados_linha in dados.items():
            nome_empresa = dados_linha[0] if dados_linha[0] and dados_linha[0] != "None" else "Empresa sem nome"
            cnpj = dados_linha[2] if len(dados_linha) > 2 else "-"
            status = dados_linha[3] if len(dados_linha) > 3 else "-"
            ocorrencia = dados_linha[4] if len(dados_linha) > 4 else "-"
            
            texto_simples += f"Linha {linha_numero} - Empresa: {nome_empresa} | CNPJ: {cnpj} | Status: {status} | Ocorrência: {ocorrencia}\n"
        
        texto_simples += """
Recomendamos a verificação detalhada destas situações para evitar possíveis 
penalidades ou restrições fiscais.

Este é um email automático gerado pelo sistema de monitoramento. Em caso de dúvidas, 
favor contatar o departamento fiscal.

Atenciosamente,
Departamento de Automação Fiscal
JotaContábil | Soluções Contábeis Inteligentes
www.jotacontabil.com.br
"""
        
        assunto = f"[ATENÇÃO] Ocorrências Identificadas no CADesp-SP ({data_hoje})"
        
        return {
            "assunto": assunto,
            "html": html_content,
            "texto": texto_simples
        }
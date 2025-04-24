import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.message import EmailMessage
from smtp_config import SMTP_CONFIG

EMAILS_RESPONSAVEIS = {
    "LIVIA": "fiscal@jotacontabil.com.br, info@jotacontabil.com.br",
    "CLAUDIO": "fiscal1@jotacontabil.com.br, info@jotacontabil.com.br",
    "ROSANA": "fiscal2@jotacontabil.com.br, info@jotacontabil.com.br",
    "CRISTINA": "fiscal3@jotacontabil.com.br, info@jotacontabil.com.br",
    "LUCIA": "fiscal4@jotacontabil.com.br, info@jotacontabil.com.br",
    "CELULA": "jota03@jotacontabil.com.br, info@jotacontabil.com.br",
}

def carregar_abas_relatorio(caminho):
    abas = pd.read_excel(caminho, sheet_name=["MESCLA", "SEQUENCIA-PREFEITURA"])
    return abas["MESCLA"], abas["SEQUENCIA-PREFEITURA"]

def montar_email_html(nome, lista_mescla, lista_seq):
    html = f"""
    <html>
    <body>
        <p>Olá <b>{nome.title()}</b>,</p>
    """
    if lista_mescla:
        html += f"<p>Identificamos <b>{len(lista_mescla)}</b> divergência(s) entre Prefeitura e Domínio:</p><ul>"
        for item in lista_mescla[:5]:
            html += f"<li>{item}</li>"
        html += "</ul>"
    if lista_seq:
        html += f"<p>Também foram detectados <b>{len(lista_seq)}</b> problema(s) de sequência de notas:</p><ul>"
        for item in lista_seq[:5]:
            html += f"<li>{item}</li>"
        html += "</ul>"
    html += """
        <p>Por favor, verifique essas inconsistências para tratativas necessárias.<br>
        <br>A notas que estão ativas na prefeitura mais não estão no sistema Dominio podem ser encontradas na pasta: T:\PMI<br>
        <br>O relatorio que apurou as inconsistencias está na pasta: T:\RELATORIOS-TI-ROTINAS\Servicos<br>
        <br>Atenciosamente,<br><b>TI Jota Contábil</b></p>
    </body>
    </html>
    """
    return html

def enviar_emails(df_mescla, df_seq):
    df_mescla.columns = ['Chave', 'Origem', 'Responsavel', 'Usuario',
                         'Valor_Pref', 'Valor_Dom', 'Status_Pref', 'Situacao_Dom', 'Divergencia']
    df_mescla["Empresa"] = df_mescla["Chave"].str.split("-").str[0]

    mescla_valid = df_mescla.dropna(subset=["Divergencia", "Responsavel"])
    divs_mescla = mescla_valid.groupby("Responsavel")["Divergencia"].apply(list).to_dict()

    df_seq["Empresa"] = df_seq["Empresa"].astype(str)
    responsaveis = df_mescla.dropna(subset=["Responsavel"]).groupby("Empresa")["Responsavel"].first().to_dict()

    divs_seq_por_resp = {}
    for _, row in df_seq.iterrows():
        empresa = row["Empresa"]
        resp = responsaveis.get(str(empresa))
        if isinstance(resp, str) and resp.strip():
            divs_seq_por_resp.setdefault(resp, []).append(row["Divergencia"])

    for resp, email in EMAILS_RESPONSAVEIS.items():
        lista_mescla = divs_mescla.get(resp, [])
        lista_seq = divs_seq_por_resp.get(resp, [])

        if not lista_mescla and not lista_seq:
            continue

        msg = EmailMessage()
        msg['Subject'] = "Inconsistências identificadas na conciliação de NFSe"
        msg['From'] = SMTP_CONFIG["remetente"]
        msg['To'] = email
        msg.set_content("Este e-mail requer um cliente compatível com HTML.")
        msg.add_alternative(montar_email_html(resp, lista_mescla, lista_seq), subtype='html')

        try:
            with smtplib.SMTP_SSL(SMTP_CONFIG["host"], SMTP_CONFIG["port"]) as smtp:
                smtp.login(SMTP_CONFIG["usuario"], SMTP_CONFIG["senha"])
                smtp.send_message(msg)
                print(f"E-mail enviado para {resp} ({email})")
        except Exception as e:
            print(f"Erro ao enviar e-mail para {resp}: {e}")

def abrir_interface_envio():
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(title="Selecione o arquivo Excel com as inconsistências",
                                         filetypes=[("Planilhas Excel", "*.xlsx")])
    if not caminho:
        return
    try:
        df_mescla, df_seq = carregar_abas_relatorio(caminho)
        enviar_emails(df_mescla, df_seq)
        messagebox.showinfo("Sucesso", "E-mails enviados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

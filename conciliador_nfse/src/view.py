import tkinter as tk
from tkinter import filedialog, messagebox
from controller import ReconciliationController
from email_inconsistencias_nfse import abrir_interface_envio


class ReconciliationView:
    def __init__(self, controller: ReconciliationController):
        self.controller = controller

    def selecionar_arquivo(self, titulo, filtros):
        return filedialog.askopenfilename(title=titulo, filetypes=filtros)

    def iniciar(self):
        root = tk.Tk()
        root.title("Conciliador de NFSe")
        root.geometry("500x400")
        
        frame = tk.Frame(root)
        frame.pack(pady=10)

        tk.Label(frame, text="Informe o Mês de Referência (MM/AAAA):").pack(pady=5)
        self.mes_entry = tk.Entry(frame, width=15, justify="center")
        self.mes_entry.pack(pady=5)

        tk.Label(frame, text="Selecione os arquivos para a conciliação").pack(pady=10)
        tk.Button(frame, text="Selecionar Arquivos e Iniciar Conciliação", command=self.on_iniciar_conciliacao).pack(pady=10)

        # NOVO BOTÃO
        tk.Label(frame, text="--- ou ---").pack(pady=5)
        tk.Button(frame, text="Verificar Sequência entre Meses", command=self.on_verificar_sequencia).pack(pady=10)

        tk.Button(frame, text="Sair", command=root.destroy).pack(pady=5)
        tk.Button(frame, text="Enviar Inconsistências por E-mail", command=abrir_interface_envio).pack(pady=10)

        root.mainloop()

    def on_iniciar_conciliacao(self):
        mes_referencia = self.mes_entry.get().strip()
        if not mes_referencia:
            messagebox.showwarning("Atenção", "Informe o mês de referência no formato MM/AAAA.")
            return

        caminho_pref = self.selecionar_arquivo("Selecione o arquivo da Prefeitura", [("Excel", "*.xlsx")])
        caminho_dom  = self.selecionar_arquivo("Selecione o arquivo do Domínio", [("CSV", "*.csv")])
        caminho_resp = self.selecionar_arquivo("Selecione o arquivo dos Funcionarios", [("Excel", "*.xlsx")])

        if not caminho_pref or not caminho_dom or not caminho_resp:
            messagebox.showwarning("Atenção", "Seleção de arquivos incompleta.")
            return

        try:
            nome_saida = self.controller.executar_conciliacao(
                caminho_pref, caminho_dom, caminho_resp, mes_referencia
            )
            messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso:\n{nome_saida}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def on_verificar_sequencia(self):
        caminho_mes_ant = self.selecionar_arquivo("Selecione o arquivo da Prefeitura (Mês Anterior)", [("Excel", "*.xlsx")])
        caminho_mes_atual = self.selecionar_arquivo("Selecione o arquivo da Prefeitura (Mês Seguinte)", [("Excel", "*.xlsx")])

        if not caminho_mes_ant or not caminho_mes_atual:
            messagebox.showwarning("Atenção", "Seleção de arquivos incompleta.")
            return

        try:
            nome_saida = self.controller.executar_verificacao_sequencia_entre_meses(caminho_mes_ant, caminho_mes_atual)
            messagebox.showinfo("Verificação Finalizada", f"Relatório gerado com sucesso:\n{nome_saida}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
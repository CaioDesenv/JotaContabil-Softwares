# gui_windows.py
# Módulo auxiliar para interagir com janelas do Windows usando pywin32 e pywinauto

import time
import win32gui
import win32con
import win32com.client
import pyautogui
from pywinauto.application import Application
import os
import tkinter as tk
from tkinter import messagebox, simpledialog

# ===========================
# Win32API - Comandos diretos
# ===========================
def listar_janelas_ativas():
    def callback(hwnd, janelas):
        if win32gui.IsWindowVisible(hwnd):
            janelas.append((hwnd, win32gui.GetWindowText(hwnd)))
    janelas = []
    win32gui.EnumWindows(callback, janelas)
    return janelas


def focar_janela_por_titulo(parte_do_titulo):
    janelas = listar_janelas_ativas()
    for hwnd, titulo in janelas:
        if parte_do_titulo.lower() in titulo.lower():
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            return True
    return False

# ===========================
# Pywinauto - Automacao de componentes
# ===========================
def selecionar_certificado(titulo_janela: str, texto_botao: str = "OK"):
    """
    Localiza a janela de seleção de certificado digital e confirma.
    """
    try:
        app = Application(backend="win32").connect(title_re=rf"{titulo_janela}")
        dlg = app.window(title_re=rf"{titulo_janela}")
        dlg.set_focus()
        time.sleep(1)
        dlg[texto_botao].click_input()
        return True
    except Exception as e:
        print(f"❌ Erro ao selecionar certificado: {e}")
        return False


def solicitar_input_usuario(titulo: str = "Entrada necessária", mensagem: str = "Digite qual linha do Excel\n vai iniciar a automação:") -> str:
    root = tk.Tk()
    root.withdraw()
    resposta = simpledialog.askstring(titulo, mensagem)
    root.destroy()
    return resposta

# ===========================
# Mensagem informativa
# ===========================
def exibir_mensagem_ok(titulo: str = "Informativo", mensagem: str = "Operação concluída com sucesso."):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(titulo, mensagem)
    root.destroy()

# ===========================
# Comandos extras
# ===========================
def fechar_janela_por_titulo(parte_do_titulo):
    janelas = listar_janelas_ativas()
    for hwnd, titulo in janelas:
        if parte_do_titulo.lower() in titulo.lower():
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True
    return False


def enviar_tecla_alt_f4():
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("%{F4}")


def minimizar_todas_as_janelas():
    """
    Minimiza todas as janelas visíveis da área de trabalho.
    """
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            except:
                pass

    win32gui.EnumWindows(callback, None)

def aguardar_janela_por_titulo(parte_do_titulo: str, timeout: int = 30, intervalo: float = 1.0) -> bool:
    """
    Aguarda até que uma janela com parte do título especificado esteja visível, dentro do tempo limite.
    """
    print(f"⏳ Aguardando janela: '{parte_do_titulo}' aparecer...")
    tempo_inicial = time.time()

    while (time.time() - tempo_inicial) < timeout:
        janelas = listar_janelas_ativas()
        for hwnd, titulo in janelas:
            if parte_do_titulo.lower() in titulo.lower():
                print(f"✅ Janela encontrada: {titulo}")
                return True
        time.sleep(intervalo)

    print(f"❌ Timeout: Janela '{parte_do_titulo}' não apareceu em {timeout} segundos.")
    return False
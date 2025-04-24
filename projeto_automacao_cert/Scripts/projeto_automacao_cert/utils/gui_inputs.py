# gui_inputs.py
# Módulo para digitação e entrada de dados usando pyautogui

import pyautogui
import time
import tkinter as tk
from tkinter import messagebox, simpledialog

# ===========================
# Utilitário de popup informativo
# ===========================
def exibir_info(titulo: str, mensagem: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(titulo, mensagem)
    root.destroy()

# ===========================
# Entrada de texto manual do usuário
# ===========================
def solicitar_input_usuario(titulo: str = "Entrada necessária", mensagem: str = "Digite qual linha do Excel\n vai iniciar a automação:") -> str:
    root = tk.Tk()
    root.withdraw()
    resposta = simpledialog.askstring(titulo, mensagem)
    root.destroy()
    return resposta

# ===========================
# Entrada de texto
# ===========================
def digitar_texto(texto: str, delay: float = 0.05):
    """Digita um texto simulando um usuário."""
    pyautogui.write(texto, interval=delay)


def colar_texto():
    """Simula CTRL+V para colar texto da área de transferência."""
    pyautogui.hotkey('ctrl', 'v')

# ===========================
# Teclas comuns com repeticao
# ===========================
def pressionar_enter(vezes: int = 1, ):
    for _ in range(vezes):
        pyautogui.press('enter', interval=0.05)


def pressionar_tab(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('tab', interval=1)


def pressionar_espaco(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('space', interval=0.05)


def pressionar_backspace(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('backspace', interval=0.05)


def pressionar_delete(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('delete', interval=0.05)


def pressionar_escape(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('esc', interval=0.05)

# ===========================
# Comandos combinados
# ===========================
def digitar_e_confirmar(texto: str, delay: float = 0.05):
    digitar_texto(texto, delay)
    time.sleep(0.2)
    pressionar_enter()


def preencher_com_tab(textos: list, delay: float = 0.05):
    """Digita uma lista de valores, separando com TAB entre os campos."""
    for valor in textos:
        digitar_texto(valor, delay)
        pyautogui.press('tab')

# ===========================
# Verificação assistida
# ===========================
def solicitar_entrada_manual(texto: str):
    exibir_info("Ação Manual Requerida", texto)
    time.sleep(1)

# Exemplo: solicitar_entrada_manual("Preencha o CAPTCHA e clique OK")
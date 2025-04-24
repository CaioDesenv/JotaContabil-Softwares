# gui_teclas.py
# Módulo auxiliar completo para comandos de teclado com pyautogui

import pyautogui
import time

# ===========================
# Funções Simples
# ===========================
def digitar_texto(texto: str, delay: float = 0.05):
    """Digita um texto caractere por caractere."""
    pyautogui.write(texto, interval=delay)


def pressionar_tecla(tecla: str, vezes: int = 1):
    """Pressiona uma tecla especial ou normal."""
    for _ in range(vezes):
        pyautogui.press(tecla)


# ===========================
# Navegação por setas
# ===========================
def seta_cima(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('up')


def seta_baixo(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('down')


def seta_esquerda(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('left')


def seta_direita(vezes: int = 1):
    for _ in range(vezes):
        pyautogui.press('right')

# ===========================
# Teclas de controle
# ===========================
def selecionar_tudo():
    pyautogui.hotkey('ctrl', 'a')


def copiar():
    pyautogui.hotkey('ctrl', 'c')


def colar():
    pyautogui.hotkey('ctrl', 'v')


def recortar():
    pyautogui.hotkey('ctrl', 'x')


def salvar():
    pyautogui.hotkey('ctrl', 's')


def fechar_janela():
    pyautogui.hotkey('alt', 'f4')


def atualizar_pagina():
    pyautogui.hotkey('f5')


def nova_aba():
    pyautogui.hotkey('ctrl', 't')


def fechar_aba():
    pyautogui.hotkey('ctrl', 'w')


def alternar_abas():
    pyautogui.hotkey('ctrl', 'tab')

# ===========================
# Teclas de função (F1 a F12)
# ===========================
def pressionar_f(tecla: int):
    """Pressiona uma tecla F1 a F12."""
    if 1 <= tecla <= 12:
        pyautogui.press(f'f{tecla}')

# ===========================
# Delay entre comandos
# ===========================
def esperar(segundos: float):
    time.sleep(segundos)

# ===========================
# Múltiplas teclas customizadas
# ===========================
def atalho(*teclas):
    """Executa um atalho com combinação de teclas."""
    pyautogui.hotkey(*teclas)

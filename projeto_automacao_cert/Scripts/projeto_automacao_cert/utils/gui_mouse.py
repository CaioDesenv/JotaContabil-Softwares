# gui_mouse.py
# Módulo auxiliar completo para controle de mouse com pyautogui + tkinter para feedback visual

import pyautogui
import time
import tkinter as tk
from tkinter import messagebox

# ===========================
# Utilitário de Janela
# ===========================
def exibir_mensagem(titulo: str, mensagem: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(titulo, mensagem)
    root.destroy()

# ===========================
# Movimento de Mouse
# ===========================
def mover_para(x: int, y: int, duracao: float = 0.2):
    pyautogui.moveTo(x, y, duration=duracao)


def mover_relativo(x_offset: int, y_offset: int, duracao: float = 0.2):
    pyautogui.moveRel(x_offset, y_offset, duration=duracao)

# ===========================
# Cliques
# ===========================
def clicar(botao: str = 'left'):
    pyautogui.click(button=botao)


def clicar_duplo(botao: str = 'left'):
    pyautogui.doubleClick(button=botao)


def clicar_direito():
    pyautogui.click(button='right')


def clicar_em_posicao(x: int, y: int, botao: str = 'left'):
    mover_para(x, y)
    clicar(botao=botao)

# ===========================
# Scroll
# ===========================
def rolar(pixels: int):
    pyautogui.scroll(pixels)

# ===========================
# Captura de posição atual do mouse
# ===========================
def obter_posicao():
    pos = pyautogui.position()
    exibir_mensagem("Posição do Mouse", f"X: {pos.x}, Y: {pos.y}")
    return pos


def imprimir_posicao_ao_mover():
    print("🖱️ Pressione Ctrl+C para sair.")
    try:
        while True:
            pos = pyautogui.position()
            print(f"Posição atual: {pos}", end='\r')
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\n🛑 Monitoramento encerrado.")

# ===========================
# Localização por imagem na tela
# ===========================
def localizar_elemento(imagem_path: str, confidence: float = 0.9):
    pos = pyautogui.locateCenterOnScreen(imagem_path, confidence=confidence)
    if pos:
        exibir_mensagem("Elemento Encontrado", f"Posição: X={pos.x}, Y={pos.y}")
    return pos


def clicar_em_elemento(imagem_path: str, tentativas: int = 10, intervalo: float = 1.0, confidence: float = 0.9):
    for tentativa in range(tentativas):
        posicao = localizar_elemento(imagem_path, confidence)
        if posicao:
            mover_para(posicao.x, posicao.y)
            clicar()
            return True
        time.sleep(intervalo)
    return False

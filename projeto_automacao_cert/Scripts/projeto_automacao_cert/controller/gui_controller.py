# controller/gui_controller.py

import pyautogui
import time
import subprocess
import os
import tkinter as tk
from tkinter import messagebox
from PIL import Image

# ==============================
# Utilitário: Popups
# ==============================
def exibir_popup_erro(titulo: str, mensagem: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(titulo, mensagem)
    root.destroy()

# ==============================
# Utilitário: Captura de tela
# ==============================
def salvar_screenshot_debug(nome: str = "log\erro.png"):
    pyautogui.screenshot(nome)
    print(f"📸 Screenshot salva como {nome} para análise posterior.")

# ==============================
# Abrir navegador Chrome com perfil
# ==============================
def abrir_chrome_com_perfil():
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    user_data_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    profile_directory = "Default"
    url = "https://fgtsdigital.sistema.gov.br"

    argumentos = [
        chrome_path,
        f"--user-data-dir={user_data_dir}",
        f"--profile-directory={profile_directory}",
        "--start-maximized",
        "--no-first-run",
        "--no-default-browser-check",
        url
    ]

    try:
        subprocess.Popen(argumentos)
        print("✅ Chrome iniciado com o perfil correto.")
    except Exception as e:
        print(f"❌ Erro ao iniciar o Chrome: {e}")

# ==============================
# Nova função: localizar imagem com variações
# ==============================
def localizar_imagem_com_variacoes(imagem_path: str, escalas: list, confiancas: list):
    for escala in escalas:
        try:
            imagem = Image.open(imagem_path)
            if escala != 1.0:
                largura, altura = imagem.size
                imagem = imagem.resize((int(largura * escala), int(altura * escala)))

            for conf in confiancas:
                posicao = pyautogui.locateCenterOnScreen(imagem, confidence=conf, grayscale=True)
                if posicao:
                    return posicao, escala, conf
        except Exception as e:
            print(f"⚠️ Erro durante a localização: {e}")
    return None, None, None

# ==============================
# Função principal: clicar botão certificado
# ==============================
def clicar_na_imagem(imagem_path: str, tentativas: int = 10):
    print(f"🔍 Procurando botão da imagem: {imagem_path} na tela...")

    escalas = [1.0, 0.95, 1.05]
    confiancas = [0.9, 0.8, 0.7]

    for tentativa in range(tentativas):
        print(f"⏳ Tentativa {tentativa + 1}/{tentativas}...")
        posicao, escala, conf = localizar_imagem_com_variacoes(imagem_path, escalas, confiancas)

        if posicao:
            pyautogui.moveTo(posicao, duration=0.2)
            pyautogui.click()
            print(f"✅ Botão clicado com sucesso! (escala={escala}, conf={conf})")
            return True

        time.sleep(1.5)

    print("❌ Botão não encontrado.")
    salvar_screenshot_debug()
    exibir_popup_erro("Erro", f"Botão da imagem: {imagem_path} não foi localizado na tela.")
    return False
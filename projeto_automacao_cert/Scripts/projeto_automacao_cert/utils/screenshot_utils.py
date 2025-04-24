# screenshot_utils.py
# Módulo exclusivo para capturas de tela com pyautogui

import pyautogui
import os
import time
from datetime import datetime

# ===========================
# Captura simples
# ===========================
def capturar_tela_simples(caminho_saida: str):
    """
    Captura a tela inteira e salva no caminho especificado.
    """
    try:
        pyautogui.screenshot(caminho_saida)
        print(f"📸 Screenshot salva em: {caminho_saida}")
    except Exception as e:
        print(f"❌ Erro ao capturar tela: {e}")

# ===========================
# Captura com timestamp
# ===========================
def capturar_com_timestamp(diretorio: str, prefixo: str = "screenshot") -> str:
    """
    Captura a tela e salva com timestamp no nome do arquivo.
    Retorna o caminho completo salvo.
    """
    try:
        if not os.path.exists(diretorio):
            os.makedirs(diretorio)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"{prefixo}_{timestamp}.png"
        caminho = os.path.join(diretorio, nome_arquivo)
        pyautogui.screenshot(caminho)
        print(f"✅ Screenshot salva: {caminho}")
        return caminho
    except Exception as e:
        print(f"❌ Erro ao salvar screenshot com timestamp: {e}")
        return ""

# ===========================
# Captura com delay
# ===========================
def capturar_com_delay(caminho_saida: str, segundos: float = 3.0):
    """
    Aguarda um tempo e então captura a tela.
    """
    print(f"⏳ Aguardando {segundos} segundos para captura...")
    time.sleep(segundos)
    capturar_tela_simples(caminho_saida)

# ===========================
# Captura retornando objeto imagem
# ===========================

def capturar_como_imagem():
    """
    Captura a tela e retorna um objeto PIL.Image para uso em OCR, etc.
    """
    try:
        imagem = pyautogui.screenshot()
        return imagem
    except Exception as e:
        print(f"❌ Erro ao capturar como imagem: {e}")
        return None
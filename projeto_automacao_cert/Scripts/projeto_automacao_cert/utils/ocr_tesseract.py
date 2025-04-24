# ocr_tesseract.py
# Módulo de funções genéricas para uso com Tesseract OCR

import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import os
import pyautogui
import win32gui
import win32con
import time

# ===========================
# Configuração inicial
# ===========================
# Caso o Tesseract não esteja no PATH, configurar manualmente abaixo:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ===========================
# Funções OCR genéricas
# ===========================
def extrair_texto_da_imagem(caminho_imagem: str, lang: str = 'por') -> str:
    """
    Realiza OCR em uma imagem e retorna o texto detectado.
    lang: Idioma de OCR (ex: 'por', 'eng')
    """
    try:
        imagem = Image.open(caminho_imagem)
        texto = pytesseract.image_to_string(imagem, lang=lang)
        return texto.strip()
    except Exception as e:
        print(f"❌ Erro ao extrair texto da imagem: {e}")
        return ""


def salvar_texto_ocr(caminho_imagem: str, caminho_saida_txt: str, lang: str = 'por'):
    """
    Extrai texto de uma imagem e salva em um arquivo .txt.
    """
    texto = extrair_texto_da_imagem(caminho_imagem, lang)
    try:
        with open(caminho_saida_txt, 'w', encoding='utf-8') as f:
            f.write(texto)
        print(f"✅ Texto OCR salvo em: {caminho_saida_txt}")
    except Exception as e:
        print(f"❌ Erro ao salvar arquivo de texto: {e}")


def extrair_texto_ajustado(caminho_imagem: str, lang: str = 'por') -> str:
    """
    Realiza OCR após aplicar melhorias na imagem (binarização, contraste, etc).
    """
    try:
        imagem = Image.open(caminho_imagem)
        imagem = imagem.convert('L')  # escala de cinza
        imagem = imagem.filter(ImageFilter.SHARPEN)
        imagem = ImageEnhance.Contrast(imagem).enhance(2)
        texto = pytesseract.image_to_string(imagem, lang=lang)
        return texto.strip()
    except Exception as e:
        print(f"❌ Erro ao processar OCR ajustado: {e}")
        return ""


def verificar_palavra_na_imagem(caminho_imagem: str, palavra: str, lang: str = 'por') -> bool:
    """
    Verifica se uma palavra específica aparece no OCR da imagem.
    """
    texto = extrair_texto_da_imagem(caminho_imagem, lang)
    return palavra.lower() in texto.lower()


# ===========================
# Função de debug (opcional)
# ===========================
def exibir_texto_ocr_terminal(caminho_imagem: str, lang: str = 'por'):
    """
    Imprime no terminal o texto detectado via OCR.
    """
    texto = extrair_texto_da_imagem(caminho_imagem, lang)
    print("\n📄 TEXTO DETECTADO VIA OCR:\n")
    print(texto)
    print("\n--- Fim do OCR ---\n")


# ===========================
# OCR da janela ativa
# ===========================
def extrair_texto_janela_ativa(lang: str = 'por') -> str:
    """
    Captura uma screenshot da janela ativa e retorna o texto extraído via OCR.
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            screenshot = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
            texto = pytesseract.image_to_string(screenshot, lang=lang)
            return texto.strip()
        else:
            print("❌ Nenhuma janela ativa identificada.")
            return ""
    except Exception as e:
        print(f"❌ Erro ao extrair texto da janela ativa: {e}")
        return ""
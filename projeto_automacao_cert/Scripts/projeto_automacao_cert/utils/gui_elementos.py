# gui_elementos.py
# Módulo utilitário para localizar e interagir com qualquer elemento visual da tela

import pyautogui
import time
from PIL import Image
from tkinter import messagebox, Tk


def exibir_erro(titulo: str, mensagem: str):
    root = Tk()
    root.withdraw()
    messagebox.showerror(titulo, mensagem)
    root.destroy()


def localizar_elemento(imagem_path: str, escalas: list = [1.0, 0.95, 1.05], confiancas: list = [0.9, 0.8, 0.7]):
    """
    Tenta localizar um elemento na tela com variações de escala e confiança.
    Retorna a posição central se encontrado, ou None.
    """
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
            print(f"⚠️ Erro ao tentar localizar imagem: {e}")

    return None, None, None


def clicar_em_elemento_img_tratada(imagem_path: str, tentativas: int = 10, intervalo: float = 1.5):
    """
    Procura e clica em um elemento da tela com base na imagem fornecida.
    """
    print(f"🔍 Procurando elemento: {imagem_path}")

    for tentativa in range(tentativas):
        print(f"⏳ Tentativa {tentativa + 1}/{tentativas}...")
        posicao, escala, conf = localizar_elemento(imagem_path)

        if posicao:
            pyautogui.moveTo(posicao, duration=0.2)
            pyautogui.click()
            print(f"✅ Elemento clicado! (escala={escala}, conf={conf})")
            return True

        time.sleep(intervalo)

    print("❌ Elemento não encontrado.")
    exibir_erro("Erro", f"Elemento '{imagem_path}' não foi localizado na tela.")
    return False


def obter_posicao_elemento(imagem_path: str):
    """
    Retorna a posição de um elemento visual sem clicar nele.
    """
    posicao, escala, conf = localizar_elemento(imagem_path)
    if posicao:
        print(f"📌 Elemento localizado em {posicao} (escala={escala}, conf={conf})")
        return posicao
    else:
        print("❌ Elemento não localizado.")
        return None

def aguardar_elemento(imagem_path: str, timeout: int = 30, intervalo: float = 1.5):
    """
    Aguarda até que uma imagem apareça na tela ou até o tempo limite (timeout).
    """
    print(f"⏳ Aguardando elemento aparecer: {imagem_path}")
    tempo_inicial = time.time()

    while (time.time() - tempo_inicial) < timeout:
        posicao, _, _ = localizar_elemento(imagem_path)
        if posicao:
            print(f"✅ Elemento detectado na tela: {posicao}")
            return True
        time.sleep(intervalo)

    print(f"❌ Timeout: elemento '{imagem_path}' não apareceu após {timeout} segundos.")
    return False
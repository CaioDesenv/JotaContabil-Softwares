#utils.py
import cv2
import pytesseract

def preprocess_image(image_path: str):
    """
    Pré-processa a imagem:
      - Carrega a imagem;
      - Converte para escala de cinza;
      - Aplica desfoque gaussiano para redução de ruídos;
      - Executa a binarização utilizando o método Otsu.
    """
    image = cv2.imread(image_path)
    if image is None:
        raise FileNotFoundError("Imagem não encontrada. Verifique o caminho informado.")
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    return thresh

def extract_text(image) -> str:
    """
    Extrai o texto da imagem pré-processada utilizando o pytesseract.
    """
    custom_config = r'--oem 3 --psm 6'
    text = pytesseract.image_to_string(image, config=custom_config, lang='eng')
    return text

def process_image(image_path: str) -> str:
    """
    Função de alto nível que recebe o caminho da imagem, processa-a e extrai o texto.
    """
    processed_image = preprocess_image(image_path)
    return extract_text(processed_image)

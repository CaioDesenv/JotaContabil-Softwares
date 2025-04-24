#controller.py
from utils import process_image
from model import ImageExtractionResult
from view import display_extracted_text

def process_image_from_path(image_path: str) -> ImageExtractionResult:
    """
    Realiza a integração dos módulos: processa a imagem e encapsula o resultado no modelo.
    """
    extracted_text = process_image(image_path)
    return ImageExtractionResult(image_path, extracted_text)

def main():
    try:
        # Neste cenário, o caminho da imagem pode ter sido obtido via extração de dados da página web.
        image_path = input("Informe o caminho da imagem extraída da página web: ").strip()
        
        # Executa o processamento da imagem
        result = process_image_from_path(image_path)
        
        # Exibe o resultado utilizando a camada de visão
        display_extracted_text(result)
    except Exception as e:
        print(f"Erro durante o processamento: {e}")

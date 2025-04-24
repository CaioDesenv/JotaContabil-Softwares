#model.py
class ImageExtractionResult:
    def __init__(self, image_path: str, extracted_text: str):
        self.image_path = image_path
        self.extracted_text = extracted_text

    def __str__(self):
        return f"Caminho: {self.image_path}\nTexto Extraído:\n{self.extracted_text}"

import fasttext
import os

class LanguageDetector:
    def __init__(self, model_path=None):
        if model_path is None:
            # Default to the downloaded model folder
            model_path = os.path.join(os.path.dirname(__file__), '../../../models/language_detection/fasttext_lid/lid.176.bin')
            
        if not os.path.exists(model_path):
            raise FileNotFoundError(f"Fasttext model not found at {model_path}. Please run download_lid_model.py")
            
        # load model
        # suppress the 'Warning : `load_model` does not return WordVectorModel'
        fasttext.FastText.eprint = lambda x: None
        self.model = fasttext.load_model(model_path)

    def detect_language(self, text: str) -> tuple[str, float]:
        """
        Returns a tuple of (language_code, confidence_score)
        Example: ('hi', 0.98)
        """
        # Fasttext expects text without newlines
        cleaned_text = text.replace('\n', ' ').strip()
        if not cleaned_text:
            return ("unknown", 0.0)
            
        predictions = self.model.predict(cleaned_text, k=1)
        label = predictions[0][0]
        confidence = float(predictions[1][0])
        
        # fasttext labels look like '__label__hi'
        lang_code = label.replace('__label__', '')
        return lang_code, confidence

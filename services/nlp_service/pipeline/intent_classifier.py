import fasttext
import os
import json

class IntentClassifier:
    def __init__(self, model_path=None):
        if model_path is None:
            model_path = os.path.join(os.path.dirname(__file__), '../../../models/intent_classification/fasttext/intent_model.bin')
            
        if not os.path.exists(model_path):
            raise FileNotFoundError(f"Intent model not found at {model_path}.")
            
        fasttext.FastText.eprint = lambda x: None
        self.model = fasttext.load_model(model_path)

    def predict_intent(self, text: str) -> tuple[str, float]:
        """
        Predicts intent from preprocessed Hindi text.
        """
        cleaned_text = text.replace('\n', ' ').strip()
        if not cleaned_text:
            return ("unknown", 0.0)
            
        predictions = self.model.predict(cleaned_text, k=1)
        label = predictions[0][0]
        confidence = float(predictions[1][0])
        
        # Strip __label__
        intent = label.replace('__label__', '')
        return intent, confidence

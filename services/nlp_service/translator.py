"""
Translation Service — Uses deep-translator (free Google Translate wrapper).
Falls back to Ollama if deep-translator fails (rate limit, etc.)
"""
from deep_translator import GoogleTranslator

SUPPORTED_LANGUAGES = {
    "en": "english",
    "hi": "hindi",
    "bn": "bengali",
    "mr": "marathi",
    "ta": "tamil",
    "te": "telugu",
    "kn": "kannada"
}

LANG_NAMES = {
    "en": "English",
    "hi": "Hindi / हिंदी",
    "bn": "Bengali / বাংলা",
    "mr": "Marathi / मराठी",
    "ta": "Tamil / தமிழ்",
    "te": "Telugu / తెలుగు",
    "kn": "Kannada / ಕನ್ನಡ"
}


class TranslationService:
    """Handles text translation between Indian languages."""

    def __init__(self, ollama_client=None):
        self.ollama = ollama_client

    def translate(self, text: str, source: str, target: str) -> str:
        """
        Translate text from source language to target language.
        Uses deep-translator (Google Translate) as primary, Ollama as fallback.
        """
        if source == target:
            return text

        if not text or not text.strip():
            return ""

        # Validate languages
        if source not in SUPPORTED_LANGUAGES or target not in SUPPORTED_LANGUAGES:
            return f"Unsupported language pair: {source} → {target}"

        try:
            result = GoogleTranslator(source=source, target=target).translate(text)
            return result or text
        except Exception as e:
            print(f"deep-translator error: {e}, falling back to Ollama")
            if self.ollama:
                return self.ollama.translate_text(text, source, target)
            return f"Translation failed: {e}"

    def detect_target_language(self, message: str) -> str:
        """Try to detect target language from a translation request message."""
        message_lower = message.lower()

        lang_keywords = {
            "hi": ["hindi", "हिंदी", "हिन्दी"],
            "bn": ["bengali", "bangla", "বাংলা", "বাংলায়"],
            "mr": ["marathi", "मराठी", "मराठीत"],
            "ta": ["tamil", "தமிழ்", "தமிழில்"],
            "te": ["telugu", "తెలుగు", "తెలుగులో"],
            "kn": ["kannada", "ಕನ್ನಡ", "ಕನ್ನಡದಲ್ಲಿ"],
            "en": ["english", "अंग्रेजी", "ইংরেজি", "इंग्रजी", "ஆங்கிலம்", "ఆంగ్లం", "ಆಂಗ್ಲ"]
        }

        for lang_code, keywords in lang_keywords.items():
            for keyword in keywords:
                if keyword in message_lower:
                    return lang_code

        return "en"  # Default: translate to English

    @staticmethod
    def get_supported_languages() -> dict:
        return LANG_NAMES

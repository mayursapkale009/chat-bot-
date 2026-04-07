from indicnlp.tokenize import indic_tokenize

def tokenize_indic_text(text: str, lang_code: str) -> list[str]:
    """
    Tokenizes text based on specific Indian language rules using indic-nlp-library.
    English/Code-mixed text can be tokenized similarly since space-separated rules generally apply,
    but indic_tokenize handles punctuation mapping properly.
    """
    # indic-nlp expects 'hi', 'bn', 'mr', 'ta', 'te', 'kn', etc.
    try:
        return indic_tokenize.trivial_tokenize(text, lang_code)
    except Exception:
        # Fallback to simple space split if lang_code is unknown
        return text.split()

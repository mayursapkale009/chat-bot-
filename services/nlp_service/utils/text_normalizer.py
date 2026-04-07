import unicodedata
import re

def normalize_text(text: str) -> str:
    """
    Standard Unicode normalization (NFKC) and basic cleanup suitable for Indic languages.
    """
    if not isinstance(text, str):
        return ""
        
    # NFKC normalizes composed & decomposed characters, handles ligatures correctly
    text = unicodedata.normalize('NFKC', text)
    
    # Lowercase text (mostly impacts English / code-mixed datasets)
    text = text.lower()
    
    # Remove excessive whitespaces
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

import os
import pytest
from services.nlp_service.pipeline.language_detector import LanguageDetector

# Setup global detector to speed up tests (loading 126MB model takes a bit)
# Note: pytest needs python path set, we will run `python -m pytest` from root
detector = None

def setup_module(module):
    global detector
    # Ensure the fasttext model path resolves correctly from root
    import sys
    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))
    # Need to handle imports after modifying path
    from services.nlp_service.pipeline.language_detector import LanguageDetector as LD
    try:
        module.detector = LD()
    except Exception as e:
        print("Skipping tests because LID model is absent.", e)
        module.detector = None

def test_hindi_language_detection():
    if detector is None:
        pytest.skip("LID model not found")
        
    text = "mera order kahan hai?" # code-mixed or transliterated might be hi or en
    text2 = "मेरा ऑर्डर कहाँ है?" # native script
    
    lang, conf = detector.detect_language(text2)
    assert lang == 'hi', f"Failed to detect Hindi, got {lang}"
    assert conf > 0.9, "Confidence score is suspiciously low!"

def test_english_language_detection():
    if detector is None:
        pytest.skip("LID model not found")
        
    text = "Where is my refund for the returned item?"
    lang, conf = detector.detect_language(text)
    assert lang == 'en', f"Expected english, got {lang}"

def test_marathi_language_detection():
    if detector is None:
        pytest.skip("LID model not found")
        
    text = "माझा ऑर्डर कुठे आहे?"
    lang, conf = detector.detect_language(text)
    assert lang == 'mr', f"Expected marathi, got {lang}"

def test_bengali_language_detection():
    if detector is None:
        pytest.skip("LID model not found")
        
    text = "আমার অর্ডার কোথায়?"
    lang, conf = detector.detect_language(text)
    assert lang == 'bn', f"Expected bengali, got {lang}"

import os
import sys

# Ensure parent directory is in path for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from services.nlp_service.utils.text_normalizer import normalize_text
from services.nlp_service.pipeline.language_detector import LanguageDetector
from services.nlp_service.pipeline.intent_classifier import IntentClassifier

app = FastAPI(title="Multilingual Chatbot Service")

# Load AI Pipelines
try:
    language_detector = LanguageDetector()
    intent_classifier = IntentClassifier()
except Exception as e:
    print("Warning: Failed to load NLP models. Ensure models are trained/downloaded.", e)
    language_detector = None
    intent_classifier = None

class MessageRequest(BaseModel):
    message: str
    user_id: str = "guest"

class MessageResponse(BaseModel):
    original_message: str
    detected_language: str
    intent: str
    confidence: float
    reply: str

@app.post("/chat/message", response_model=MessageResponse)
async def handle_message(req: MessageRequest):
    if not language_detector or not intent_classifier:
        raise HTTPException(status_code=500, detail="NLP Services offline.")
        
    # 1. Preprocessing
    clean_text = normalize_text(req.message)
    
    # 2. Language Detection
    lang, lang_conf = language_detector.detect_language(clean_text)
    
    # 3. Intent Classification
    intent, intent_conf = intent_classifier.predict_intent(clean_text)
    
    # 4. Formulate Reply (MVP hardcoded logic for Iteration 1)
    if lang != 'hi' and lang != 'en':
        reply = f"Sorry, I only understand Hindi and English right now. (Detected: {lang})"
    else:
        # Dummy Knowledge base responses based on intent
        if intent == 'track_order':
            reply = "आपका ऑर्डर रास्ते में है और जल्द ही दिया जाएगा।" if lang == 'hi' else "Your order is on the way."
        elif intent == 'cancel_order':
            reply = "क्या आप वाकयी इस ऑर्डर को रद्द करना चाहते हैं?" if lang == 'hi' else "Do you want to cancel?"
        else:
            reply = f"Detected Intent: {intent}. No exact answer configured yet."

    return MessageResponse(
        original_message=req.message,
        detected_language=lang,
        intent=intent,
        confidence=intent_conf,
        reply=reply
    )

@app.get("/health")
def health_check():
    return {"status": "ok"}

import os
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

folders = [
    "docs",
    "experiments/mlruns",
    "experiments/configs",
    "data/raw/ai4bharat",
    "data/raw/indicvoices",
    "data/raw/social_media",
    "data/processed/intent/hindi",
    "data/processed/intent/bengali",
    "data/processed/intent/marathi",
    "data/processed/intent/tamil",
    "data/processed/intent/telugu",
    "data/processed/intent/kannada",
    "data/processed/ner",
    "data/processed/language_detection",
    "data/processed/speech",
    "data/knowledge_base",
    "models/language_detection/fasttext_lid",
    "models/intent_classification/mbert_multilingual",
    "models/intent_classification/indicbert_hindi",
    "models/intent_classification/indicbert_bengali",
    "models/ner/mbert_ner",
    "models/asr/indicconformer",
    "models/asr/whisper_hindi_finetuned",
    "services/api-gateway/middleware",
    "services/nlp-service/pipeline",
    "services/nlp-service/models",
    "services/nlp-service/utils",
    "services/chat-service/session",
    "services/chat-service/db",
    "services/admin-service/routers",
    "services/integration-service/translation",
    "services/voice-service/asr",
    "services/voice-service/tts",
    "frontend/chat-widget/voice",
    "frontend/admin-dashboard/public",
    "frontend/admin-dashboard/src/pages",
    "frontend/admin-dashboard/src/components",
    "frontend/admin-dashboard/src/services",
    "infrastructure/docker",
    "infrastructure/kubernetes",
    "infrastructure/aws",
    "scripts/data_collection",
    "scripts/training",
    "scripts/testing",
    "tests/unit",
    "tests/integration/conversations",
    "tests/performance"
]

files = {
    "docker-compose.yml": "",
    "frontend/admin-dashboard/package.json": "{\n  \"name\": \"admin-dashboard\",\n  \"version\": \"1.0.0\"\n}",
    "tests/__init__.py": "",
    "tests/unit/__init__.py": "",
    "tests/integration/__init__.py": "",
    "tests/performance/__init__.py": ""
}

def create_structure():
    # Create Directories
    for folder in folders:
        os.makedirs(folder, exist_ok=True)
        logging.info(f"Created directory: {folder}")
        
    # Create placeholder files
    for filepath, content in files.items():
        if not os.path.exists(filepath):
            with open(filepath, 'w') as f:
                f.write(content)
            logging.info(f"Created file: {filepath}")

if __name__ == "__main__":
    create_structure()
    logging.info("Project file structure successfully created!")

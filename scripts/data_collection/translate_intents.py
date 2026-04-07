import csv
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from deep_translator import GoogleTranslator

INPUT_FILE = os.path.join(os.path.dirname(__file__), '../../data/raw/english_base_intents.csv')
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), '../../data/processed/intent/hindi/all.jsonl')

def translate_row(row):
    translator = GoogleTranslator(source='en', target='hi')
    try:
        hi_text = translator.translate(row['text'])
        # Add a tiny sleep to prevent rate limiting somewhat, but multithreading will still hit servers.
        time.sleep(0.1)
        return {
            "text": hi_text,
            "intent": row['intent'],
            "lang": "hi"
        }
    except Exception as e:
        print(f"Error translating '{row['text']}': {e}")
        # fallback to English text if failed
        return {
            "text": row['text'],
            "intent": row['intent'],
            "lang": "error"
        }

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"File {INPUT_FILE} not found!")
        return

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    print(f"Loaded {len(rows)} rows for translation.")

    # We will run this cautiously. You can adjust max_workers if you get blocked.
    # For a dataset of 4000, 10 workers take about ~2-5 mins.
    results = []
    
    # Optional limit for testing: rows = rows[:100] (remove chunking for full run)
    print("Starting translation. This may take a few minutes...")

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(translate_row, row): row for row in rows}
        
        count = 0
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as out_f:
            for future in as_completed(futures):
                res = future.result()
                if res['lang'] == 'hi':
                    out_f.write(json.dumps(res, ensure_ascii=False) + '\n')
                count += 1
                if count % 100 == 0:
                    print(f"Translated {count}/{len(rows)} examples...")

    print(f"Done! Translated examples saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()

"""
Ollama Client — OPTIMIZED for speed.
Single LLM call for intent + response. Supports streaming.
Uses Llama 3.2 3B for faster inference on CPU.
"""
import ollama
import json
import re

SYSTEM_PROMPT = """You are an intelligent multilingual chatbot for Indian languages.
You can converse fluently in: Hindi (hi), Bengali (bn), Marathi (mr), Tamil (ta), Telugu (te), Kannada (kn), and English (en).

RULES:
1. Always respond in the SAME language the user wrote in.
2. Be polite, helpful, and concise.
3. If you don't know something, say so honestly.
4. Provide detailed, helpful, and informative responses.
5. For translation requests, translate accurately."""

# Combined prompt: classify intent AND generate response in ONE call
COMBINED_PROMPT = """User message: "{message}"
Detected language: {language}

First identify the intent from: greeting, farewell, thanks, ask_name, ask_help, language_change, general_knowledge, weather, time_date, translation_request, joke, news, math_calculation, dictionary, sentiment, complaint, feedback, small_talk, ask_creator, fallback

Respond with EXACTLY this format (2 lines only):
INTENT: <intent_name>
RESPONSE: <your natural response in {language}>"""


class OllamaClient:
    """Optimized Ollama wrapper — single call for classify+respond."""

    def __init__(self, model_name: str = "llama3.2:3b"):
        self.model = model_name
        self._verify_model()

    def _verify_model(self):
        """Check if the model is available."""
        try:
            ollama.show(self.model)
            print(f"✅ Ollama model loaded: {self.model}")
        except Exception as e:
            print(f"⚠️ Model '{self.model}' not available: {e}")
            # Fallback to default tag
            try:
                self.model = "llama3.2"
                ollama.show(self.model)
                print(f"✅ Fallback model: {self.model}")
            except:
                print("❌ No Ollama model available!")

    def classify_and_respond(self, message: str, language: str,
                              history: list = None) -> dict:
        """
        SINGLE Ollama call: classify intent + generate response together.
        Returns: {"intent": str, "confidence": float, "response": str}
        """
        lang_names = {
            "hi": "Hindi", "bn": "Bengali", "mr": "Marathi",
            "ta": "Tamil", "te": "Telugu", "kn": "Kannada", "en": "English"
        }
        lang_name = lang_names.get(language, "English")

        prompt = COMBINED_PROMPT.format(
            message=message,
            language=lang_name
        )

        messages = [{"role": "system", "content": SYSTEM_PROMPT}]

        # Add last 3 turns for context (minimal to keep it fast)
        if history:
            for turn in history[-3:]:
                messages.append({
                    "role": turn.get("role", "user"),
                    "content": turn.get("text", "")
                })

        messages.append({"role": "user", "content": prompt})

        try:
            response = ollama.chat(
                model=self.model,
                messages=messages,
                options={
                    "temperature": 0.5,
                    "num_predict": 1024,    # Allow for very detailed responses
                    "num_ctx": 2048,        # Large context window to prevent cutoffs
                    "repeat_penalty": 1.1,
                    "top_k": 10,
                    "top_p": 0.7,
                }
            )
            raw = response["message"]["content"].strip()
            return self._parse_combined_response(raw)

        except Exception as e:
            print(f"Ollama error: {e}")
            return {
                "intent": "fallback",
                "confidence": 0.0,
                "response": "Sorry, I encountered an error. Please try again."
            }

    def _parse_combined_response(self, raw: str) -> dict:
        """Parse the combined INTENT + RESPONSE output."""
        intent = "fallback"
        response_text = raw
        confidence = 0.7

        lines = raw.strip().split('\n')

        for i, line in enumerate(lines):
            line_stripped = line.strip()

            # Extract intent
            if line_stripped.upper().startswith('INTENT:'):
                intent_val = line_stripped.split(':', 1)[1].strip().lower()
                # Clean up any extra text
                intent_val = intent_val.split()[0] if intent_val else "fallback"
                # Remove any punctuation
                intent_val = re.sub(r'[^a-z_]', '', intent_val)
                if intent_val:
                    intent = intent_val
                    confidence = 0.85

            # Extract response
            elif line_stripped.upper().startswith('RESPONSE:'):
                response_text = line_stripped.split(':', 1)[1].strip()
                # Include remaining lines too
                remaining = '\n'.join(l.strip() for l in lines[i+1:] if l.strip()
                                      and not l.strip().upper().startswith('INTENT:'))
                if remaining:
                    response_text += '\n' + remaining
                break

        # If we couldn't parse, use the whole raw text as response
        if response_text == raw and 'INTENT:' in raw:
            # Remove the INTENT line from response
            response_text = '\n'.join(
                l for l in lines
                if not l.strip().upper().startswith('INTENT:')
                and not l.strip().upper().startswith('RESPONSE:')
            ).strip()

        if not response_text:
            response_text = raw

        return {
            "intent": intent,
            "confidence": confidence,
            "response": response_text
        }

    def generate_response(self, message: str, intent: str, language: str,
                          history: list = None) -> str:
        """Generate a standalone response (for dynamic intents)."""
        lang_names = {
            "hi": "Hindi", "bn": "Bengali", "mr": "Marathi",
            "ta": "Tamil", "te": "Telugu", "kn": "Kannada", "en": "English"
        }
        lang_name = lang_names.get(language, "English")

        messages = [
            {"role": "system", "content": f"{SYSTEM_PROMPT}\nRespond in {lang_name}. Provide a detailed and informative answer."},
            {"role": "user", "content": message}
        ]

        try:
            response = ollama.chat(
                model=self.model,
                messages=messages,
                options={
                    "temperature": 0.6,
                    "num_predict": 1024,
                    "num_ctx": 2048,
                    "top_k": 20,
                }
            )
            return response["message"]["content"].strip()
        except Exception as e:
            return f"Error: {e}"

    def stream_response(self, message: str, language: str,
                        history: list = None):
        """
        Stream response tokens for real-time display.
        Includes conversation history for multi-turn context.
        Yields chunks of text.
        """
        lang_names = {
            "hi": "Hindi", "bn": "Bengali", "mr": "Marathi",
            "ta": "Tamil", "te": "Telugu", "kn": "Kannada", "en": "English"
        }
        lang_name = lang_names.get(language, "English")

        messages = [
            {"role": "system", "content": f"Reply in {lang_name}. Provide a detailed and helpful response."}
        ]

        # Add conversation history for multi-turn context
        if history:
            for turn in history[-6:]:  # Last 3 exchanges
                messages.append({
                    "role": turn.get("role", "user"),
                    "content": turn.get("text", "")
                })

        messages.append({"role": "user", "content": message})

        try:
            stream = ollama.chat(
                model=self.model,
                messages=messages,
                stream=True,
                options={
                    "temperature": 0.6,
                    "num_predict": 1024,
                    "num_ctx": 2048,
                    "top_k": 10,
                }
            )
            for chunk in stream:
                token = chunk["message"]["content"]
                if token:
                    yield token
        except Exception as e:
            yield f"Error: {e}"

    def translate_text(self, text: str, source_lang: str, target_lang: str) -> str:
        """Translate text using Ollama."""
        lang_names = {
            "hi": "Hindi", "bn": "Bengali", "mr": "Marathi",
            "ta": "Tamil", "te": "Telugu", "kn": "Kannada", "en": "English"
        }
        src = lang_names.get(source_lang, source_lang)
        tgt = lang_names.get(target_lang, target_lang)

        try:
            response = ollama.chat(
                model=self.model,
                messages=[
                    {"role": "system", "content": f"Translate from {src} to {tgt}. Return ONLY the translation."},
                    {"role": "user", "content": text}
                ],
                options={"temperature": 0.1, "num_predict": 100, "num_ctx": 512}
            )
            return response["message"]["content"].strip()
        except Exception as e:
            return f"Translation error: {e}"

    def summarize_context(self, history: list) -> str:
        """Summarize conversation history into a concise context string."""
        if not history:
            return ""
        
        text_to_summarize = ""
        for turn in history:
            role = turn.get("role", "unknown")
            text = turn.get("text", "")
            text_to_summarize += f"{role.capitalize()}: {text}\n"
            
        messages = [
            {"role": "system", "content": "You are a highly efficient memory summarizer. Summarize the following conversation in 2-3 concise sentences. Focus on the main topics discussed and any important facts the user shared. Do not include pleasantries."},
            {"role": "user", "content": text_to_summarize}
        ]
        
        try:
            response = ollama.chat(
                model=self.model,
                messages=messages,
                options={"temperature": 0.3, "num_predict": 150, "num_ctx": 2048}
            )
            return response["message"]["content"].strip()
        except Exception as e:
            print(f"Summarization error: {e}")
            return "Previous conversation context retained."

    def health_check(self) -> bool:
        """Check if Ollama is running."""
        try:
            ollama.show(self.model)
            return True
        except:
            return False

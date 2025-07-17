import speech_recognition as sr
import pyttsx3
import pymongo
import requests
import re
import time
import threading
import keyboard
import subprocess
import os
import webbrowser
import ctypes
import sys

# ------------------ CONFIG ------------------ #
MONGO_URI = "mongodb+srv://knpranay2806:Pranay2806@jarvis.sm2rte0.mongodb.net/"
DB_NAME = "sera"
COLLECTION_NAME = "knowledge"
SERP_API_KEY = "a0b5b44349fdef7b6bba60adcb5de44fa99607d1d3a15d905f1cf470ff54f265"

# ------------------ TEXT TO SPEECH ------------------ #
class Speaker:
    def say(self, text):
        try:
            print(f"Speaking: {text}")
            engine = pyttsx3.init()
            engine.setProperty("rate", 160)
            engine.setProperty("volume", 1.0)
            voices = engine.getProperty("voices")
            for voice in voices:
                if "female" in voice.name.lower() or "zira" in voice.name.lower():
                    engine.setProperty("voice", voice.id)
                    break
            engine.say(text)
            engine.runAndWait()
            engine.stop()
        except Exception as e:
            print(f"TTS error: {e}")

# ------------------ VOICE INPUT ------------------ #
class VoiceInput:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        try:
            self.microphone = sr.Microphone()
            print("Microphone detected.")
        except Exception:
            print("No microphone found. You will be prompted to type instead.")
            self.microphone = None

    def listen(self):
        if not self.microphone:
            return input("Type your query: ")

        with self.microphone as source:
            source.SAMPLE_RATE = 16000
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source)
            try:
                audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=10)
            except sr.WaitTimeoutError:
                print("Timeout: No speech detected.")
                return None

        try:
            return self.recognizer.recognize_google(audio)
        except sr.UnknownValueError:
            print("Sorry, I didn't catch that.")
            return None
        except sr.RequestError:
            print("Speech service unavailable.")
            return None

# ------------------ PHRASE PARSER ------------------ #
class PhraseParser:
    def split_phrases(self, text):
        return re.split(r'[.?!,;]| and | then ', text.lower())

# ------------------ DATABASE ------------------ #
class MongoHandler:
    def __init__(self):
        self.client = pymongo.MongoClient(MONGO_URI)
        self.collection = self.client[DB_NAME][COLLECTION_NAME]

    def search_phrase(self, phrase):
        result = self.collection.find_one({"question": phrase})
        return result["answer"] if result else None

    def insert_phrase(self, question, answer):
        self.collection.insert_one({"question": question, "answer": answer})

# ------------------ SERP API SCRAPER ------------------ #
class SerpAPIScraper:
    def fetch_answer(self, query):
        try:
            params = {
                "engine": "google",
                "q": query,
                "api_key": SERP_API_KEY
            }
            response = requests.get("https://serpapi.com/search", params=params)
            data = response.json()

            if "answer_box" in data and "answer" in data["answer_box"]:
                return data["answer_box"]["answer"]
            elif "answer_box" in data and "snippet" in data["answer_box"]:
                return data["answer_box"]["snippet"]
            elif "snippet" in data.get("organic_results", [{}])[0]:
                return data["organic_results"][0]["snippet"]
            else:
                return "Sorry, I couldn't find an answer."
        except Exception as e:
            return f"Error using SERP API: {e}"

# ------------------ WINDOWS COMMAND EXECUTOR ------------------ #
class WindowsCommandHandler:
    def __init__(self, speaker, assistant):
        self.speaker = speaker
        self.assistant = assistant

    def handle(self, phrase):
        phrase = phrase.lower()

        if "open chrome" in phrase:
            os.startfile("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
            self.speaker.say("Opening Chrome.")
            return True

        elif "open notepad" in phrase:
            subprocess.Popen(["notepad.exe"])
            self.speaker.say("Opening Notepad.")
            return True

        elif "open calculator" in phrase:
            subprocess.Popen(["calc.exe"])
            self.speaker.say("Opening Calculator.")
            return True

        elif "shutdown" in phrase:
            self.speaker.say("Shutting down.")
            os.system("shutdown /s /t 1")
            return True

        elif "restart" in phrase:
            self.speaker.say("Restarting.")
            os.system("shutdown /r /t 1")
            return True

        elif "lock system" in phrase:
            self.speaker.say("Locking system.")
            ctypes.windll.user32.LockWorkStation()
            return True

        elif "open youtube" in phrase:
            webbrowser.open("https://www.youtube.com")
            self.speaker.say("Opening YouTube.")
            return True

        elif "open github" in phrase:
            webbrowser.open("https://www.github.com")
            self.speaker.say("Opening GitHub.")
            return True

        elif "open folder" in phrase:
            os.startfile("C:\\Users\\%USERNAME%\\Documents")
            self.speaker.say("Opening your Documents folder.")
            return True

        elif "stop" in phrase or "exit" in phrase:
            self.speaker.say("Goodbye.")
            self.assistant.quit()
            return True

        return False

# ------------------ KEYBOARD SHORTCUT HANDLER ------------------ #
class KeyboardHandler:
    def __init__(self, assistant):
        self.assistant = assistant
        self.bind_shortcuts()

    def bind_shortcuts(self):
        keyboard.add_hotkey("ctrl+alt+p", self.assistant.pause)
        keyboard.add_hotkey("ctrl+alt+r", self.assistant.resume)
        keyboard.add_hotkey("ctrl+alt+q", self.assistant.quit)
        # print("Shortcuts: Pause (Ctrl+Alt+P), Resume (Ctrl+Alt+R), Quit (Ctrl+Alt+Q)")

# ------------------ SERA ASSISTANT ------------------ #
class SeraAssistant:
    def __init__(self):
        self.voice = VoiceInput()
        self.parser = PhraseParser()
        self.db = MongoHandler()
        self.scraper = SerpAPIScraper()
        self.speaker = Speaker()
        self.running = True
        self.paused = False
        self.commands = WindowsCommandHandler(self.speaker, self)

    def respond_to_input(self, raw_text):
        if not raw_text:
            print("Empty input.")
            return

        print(f"You said: {raw_text}")
        phrases = self.parser.split_phrases(raw_text)

        for phrase in phrases:
            phrase = phrase.strip()
            if not phrase:
                continue

            if self.commands.handle(phrase):
                continue

            answer = self.db.search_phrase(phrase)
            if answer:
                print(f"Sera: {answer}")
                self.speaker.say(answer)
            else:
                print("Searching online...")
                answer = self.scraper.fetch_answer(phrase)
                # print(f"Sera (web): {answer}")
                self.speaker.say(answer)
                self.db.insert_phrase(phrase, answer)

    def start(self):
        print("Sera is running. Use shortcuts: Ctrl+Alt+P to pause, Ctrl+Alt+Q to quit.")
        self.speaker.say("SERA activated. Awaiting your command")
        while self.running:
            try:
                if self.paused:
                    time.sleep(1)
                    continue
                print("Waiting for voice input...")
                raw_text = self.voice.listen()
                self.respond_to_input(raw_text)
                time.sleep(1)
            except Exception as e:
                print(f"Error in loop: {e}")
                continue

    def pause(self):
        print("Paused.")
        self.paused = True

    def resume(self):
        print("Resumed.")
        self.paused = False

    def quit(self):
        print("Exiting.")
        self.running = False

# ------------------ MAIN ------------------ #
if __name__ == "__main__":
    print("Sera assistant with SERP, Windows Commands, and TTS started.")
    assistant = SeraAssistant()
    threading.Thread(target=KeyboardHandler, args=(assistant,), daemon=True).start()
    assistant.start()

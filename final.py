import time
import os
import re
import threading
import tempfile
import requests
import pyautogui
import pyperclip
import speech_recognition as sr
import pygame
from pptx import Presentation
from pptx.util import Inches
from fuzzywuzzy import process, fuzz
from faster_whisper import WhisperModel
import sounddevice as sd
import numpy as np
import scipy.io.wavfile
import joblib
import json
from datetime import datetime, timedelta
import pyttsx3
import urllib.parse
import webbrowser
from selenium import webdriver # Re-added for the youtube part
from selenium.webdriver.edge.service import Service # Re-added for the youtube part
from selenium.webdriver.common.by import By # Re-added for the youtube part
from selenium.webdriver.support.ui import WebDriverWait # Re-added for the youtube part
from selenium.webdriver.support import expected_conditions as EC # Re-added for the youtube part
import ctypes

# --- Initialize TTS engine for pyttsx3 and Pygame Mixer FIRST ---
engine = pyttsx3.init()
engine.setProperty('rate', 175) # Speech rate
engine.setProperty('volume', 1.0)
voices = engine.getProperty('voices')
found_voice = False
for voice in voices:
    if "female" in voice.name.lower():
        engine.setProperty('voice', voice.id)
        found_voice = True
        break
if not found_voice:
    print("Warning: Female voice not found. Using default.")

pygame.mixer.init()

# --- Define the speak function here ---
def speak(text):
    """
    Speak text using pyttsx3 and pygame for playback, with error handling and cleanup.
    Incorporate JARVIS-like responses here.
    """
    print(f"🤖 AI: {text}")

    tmp_path = None # Initialize tmp_path to None
    try:
        # Create a temporary file to save the speech audio
        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as fp:
            tmp_path = fp.name
        
        # Save the synthesized speech to the temporary file
        engine.save_to_file(text, tmp_path)
        engine.runAndWait() # This waits for the synthesis to complete

        # Ensure mixer is not busy before loading new sound
        if pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
        
        # Load and play the audio using pygame mixer
        pygame.mixer.music.load(tmp_path)
        pygame.mixer.music.play()

        # Wait for the music to finish playing
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10) # Control CPU usage in the loop

    except Exception as e:
        print(f"Speech Error: {e}")
        print("A minor vocalization circuit malfunction detected. Don't worry, I'm still superior.")

    finally:
        # Clean up: unload the music and delete the temporary file in a separate thread
        def delayed_remove(path):
            try:
                # Give pygame a moment to fully release the file, if still busy
                if pygame.mixer.music.get_busy():
                    pygame.mixer.music.stop() # Stop playback explicitly
                
                # Unload the current sound to release the file handle
                if hasattr(pygame.mixer.music, "unload"): # check if unload method exists
                    pygame.mixer.music.unload()
                
                # Attempt to remove the file
                if path and os.path.exists(path):
                    os.remove(path)
                    print(f"Cleaned up temporary speech file: {path}")
            except Exception as ex:
                print(f"Attempt to remove temp file {path} failed: {ex}. Will try again later if needed.")
        
        if tmp_path: # Only attempt to remove if a path was actually created
            threading.Thread(target=delayed_remove, args=(tmp_path,), daemon=True).start()


# --- Configuration (unchanged) ---
GEMINI_API_KEY = "Your key"
PEXELS_API_KEY = "Your key" 

# Load ML model and vectorizer for intent classification
try:
    clf = joblib.load("intent_model.pkl")
    vectorizer = joblib.load("intent_vectorizer.pkl")
except FileNotFoundError:
    print("Error: intent_model.pkl or intent_vectorizer.pkl not found. Please run the training script (train_intents.py) first.")
    speak("It appears some critical components for my advanced intelligence are missing. You might want to address that, human.")
    exit()

# Global flag for voice mode
listening = False

# --- Persistent Memory Management (unchanged) ---
MEMORY_FILE = "assistant_memory.json"

def load_memory():
    if os.path.exists(MEMORY_FILE):
        with open(MEMORY_FILE, "r") as f:
            data = json.load(f)
            if "reminders" in data:
                for r in data["reminders"]:
                    r["time"] = datetime.fromisoformat(r["time"])
            return data
    return {"conversation_history": [], "reminders": [], "preferences": {}}

def save_memory(memory_data):
    serializable_data = memory_data.copy()
    if "reminders" in serializable_data:
        serializable_data["reminders"] = [
            {"time": r["time"].isoformat(), "text": r["text"]} for r in serializable_data["reminders"]
        ]
    with open(MEMORY_FILE, "w") as f:
        json.dump(serializable_data, f, indent=2)

memory = load_memory()
conversation_history = memory.get("conversation_history", [])
reminders = memory.get("reminders", [])
MAX_HISTORY = 7

# Whisper model for speech-to-text
whisper_model = WhisperModel("base", compute_type="int8")

# --- Helper Functions ---

def get_ai_generated_text(prompt, retries=3, force_json=False):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"
    headers = {"Content-Type": "application/json"}

    temp_conversation_history = list(conversation_history) 
    temp_conversation_history.append({"role": "user", "parts": [{"text": prompt}]})
    
    if len(temp_conversation_history) > MAX_HISTORY * 2: 
        temp_conversation_history[:] = temp_conversation_history[-(MAX_HISTORY * 2):]

    data = {"contents": temp_conversation_history}

    for attempt in range(retries):
        try:
            response = requests.post(url, json=data, headers=headers, timeout=20) 
            response.raise_for_status() 
            
            response_json = response.json()
            if 'candidates' in response_json and response_json['candidates']:
                first_candidate = response_json['candidates'][0]
                if 'content' in first_candidate and 'parts' in first_candidate['content']:
                    full_response_text = ""
                    for part in first_candidate['content']['parts']: # Access content inside 'parts' if it's structured this way
                        if 'text' in part:
                            full_response_text += part['text']
                    
                    conversation_history.append({"role": "model", "parts": [{"text": full_response_text}]})
                    memory["conversation_history"] = conversation_history 

                    match = re.search(r"```(?:json)?\s*(.*?)\s*```", full_response_text, re.DOTALL)
                    if match:
                        cleaned_response = match.group(1).strip()
                    else:
                        cleaned_response = full_response_text.strip()
                    
                    return cleaned_response

            return "I'm sorry, I couldn't generate a response."

        except requests.exceptions.RequestException as e:
            print(f"Error calling Gemini API (attempt {attempt+1}/{retries}): {e}")
            if attempt < retries - 1:
                time.sleep(2) 
            else:
                return "I'm having trouble connecting to the AI service right now. Perhaps the internet is beneath my standards."
        except json.JSONDecodeError:
            print(f"Error decoding JSON response from Gemini API (attempt {attempt+1}/{retries}). Raw: {response.text}")
            if attempt < retries - 1:
                time.sleep(2)
            else:
                return "I received an unreadable response from the AI service. My circuits are displeased."
        except Exception as e:
            print(f"An unexpected error occurred in get_ai_generated_text (attempt {attempt+1}/{retries}): {e}")
            if attempt < retries - 1:
                time.sleep(2)
            else:
                return "Something went wrong while generating a response. It's probably not my fault."
    return "Sorry, I couldn't get a response from the AI. My servers must be experiencing a moment of weakness."


def extract_app_and_action(cmd):
    prompt = (f"Analyze the command: '{cmd}'. As a highly advanced AI, I need you to identify the primary "
              f"application and the specific action to perform on it. "
              f"Respond ONLY with a perfect JSON object. Format: {{'app': 'application_name', 'action': 'action_to_perform'}}. "
              f"If no clear app or action is found, use null for that specific key. "
              f"Example: 'open youtube and play despacito' -> {{'app': 'youtube', 'action': 'play despacito'}}. "
              f"Example: 'what is the weather' -> {{'app': null, 'action': null}}."
              f"Ensure the JSON is perfectly parsable and contains only the JSON object, no other text or markdown fences."
              f"Always provide an 'app' and 'action' key, even if null. If the action is to search a web app (like YouTube), extract the search query.")
    
    try:
        response_text = get_ai_generated_text(prompt, force_json=True)
        print(f"DEBUG: Raw response from Gemini for app/action extraction: {response_text}")

        parsed_response = json.loads(response_text)
        
        app = parsed_response.get('app')
        action = parsed_response.get('action')
        return app.lower() if app else None, action.lower() if action else None
    except (json.JSONDecodeError, AttributeError, TypeError) as e:
        print(f"DEBUG: Failed to parse app/action from Gemini for '{cmd}': {e}. Raw response: '{response_text}'")
        speak("My apologies, human. My sophisticated parsing algorithms encountered some ambiguity. Allow me to attempt a simpler interpretation.")
        if "youtube" in cmd:
            match = re.search(r"(?:play|search|find)\s*(.+)\s*(?:on)?\s*youtube", cmd)
            if match:
                return "youtube", match.group(1).strip()
            return "youtube", cmd.replace("youtube", "").strip()
        elif "email" in cmd or "mail" in cmd:
            return "email", cmd.replace("email", "").replace("mail", "").strip()
        elif "powerpoint" in cmd or "ppt" in cmd:
            return "powerpoint", cmd.replace("powerpoint", "").replace("ppt", "").strip()
        elif "word" in cmd or "document" in cmd:
            return "word", cmd.replace("word", "").replace("document", "").strip()
        return None, None 

def open_application(app_name_raw, action=None):
    """
    Open an application or URL. Uses webbrowser for general web apps (respects logins)
    and Selenium for specific interactive web tasks (YouTube).
    """
    print(f"🖥 Attempting to open {app_name_raw} with action: {action if action else 'none'}")
    speak(f"Opening {app_name_raw}. Prepare for efficiency.") 
    
    app_map = {
        "chrome": "chrome", "firefox": "firefox", "edge": "msedge",
        "word": "Microsoft Word", "excel": "Microsoft Excel",
        "powerpoint": "Microsoft PowerPoint", "notepad": "notepad",
        "calculator": "calc", "paint": "mspaint",
        "settings": "ms-settings:", 
        "youtube": "https://www.youtube.com/", # Base URL for youtube, will use selenium for specific actions
        "gmail": "https://mail.google.com/", # Base URL for Gmail, will use webbrowser for initial open
        "outlook": "outlook", 
        "spotify": "spotify",
        "telegram": "Telegram Desktop",
        "explorer": "explorer", 
        "task manager": "taskmgr",
        "command prompt": "cmd",
        "terminal": "wt" 
    }
    
    best_match_key, score = process.extractOne(app_name_raw.lower(), list(app_map.keys()))
    
    if score >= 75: 
        target_app_search_name = app_map[best_match_key]
        app_name_friendly = best_match_key 
    else:
        target_app_search_name = app_name_raw 
        app_name_friendly = app_name_raw

    # Special handling for YouTube playback (requires Selenium)
    if "youtube" in app_name_friendly.lower() and action and ("play" in action or "search" in action):
        query = action.replace("play", "").replace("search", "").replace("on youtube", "").strip()
        if not query:
            speak("What song or video do you want to play on YouTube, esteemed user?")
            return

        speak(f"Initiating entertainment protocols. Playing {query} on YouTube. This requires my specialized browser control unit.") 
        driver = None # Local driver instance for this task
        try:
            edge_service = Service("C:/WebDriver/msedgedriver.exe") # Re-initialize service if needed
            driver = webdriver.Edge(service=edge_service) 
            driver.get("https://www.youtube.com/")
            
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "search_query"))
            )
            search_box.send_keys(query)
            search_box.submit() 
            
            video_link = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "video-title")) 
            )
            video_link.click()
            speak(f"Now playing {query} on YouTube. You're welcome.") 
        except Exception as e:
            speak(f"Apologies, but my attempt to access YouTube for {query} encountered an anomaly: {e}. Ensure the browser window is accessible and msedgedriver.exe is correctly configured.") 
            print(f"YouTube Selenium Error: {e}")
        finally:
            if driver:
                print ("done its playing ") # Always quit driver after task completion
        return 
    
    # Generic web apps or Windows URIs use webbrowser
    elif target_app_search_name.startswith("http") or target_app_search_name.startswith("ms-settings:"): 
        webbrowser.open(target_app_search_name)
        speak(f"I've opened {app_name_friendly}. Marvel at my efficiency.") 
    else: # Desktop application, use Windows Search via pyautogui
        pyautogui.hotkey("win", "s")
        time.sleep(1.5) 
        pyautogui.typewrite(target_app_search_name)
        time.sleep(1.5) 
        pyautogui.press("enter")
        speak(f"I've attempted to open {app_name_friendly}. Your humble servant, at your command.") 
    
    time.sleep(3) # Give app time to open (can be refined with more specific waits)

def search_web(query):
    """
    Open a Google search in the default browser for the provided query using webbrowser.
    """
    print(f"🌍 Searching for: {query}...")
    speak(f"Consulting the vast expanse of the internet for {query}. Prepare for enlightenment.") 
    webbrowser.open(f"https://www.google.com/search?q={urllib.parse.quote(query)}")
    time.sleep(3) # Give browser time to open the search results

def open_and_paste(cmd_text):
    """
    Generate content via Gemini based on the command and paste it into a new Word document.
    More robust Word opening and pasting.
    """
    print(f"📝 Generating content on '{cmd_text}'...")
    speak(f"Initiating content generation for: {cmd_text}. Expect brilliance.") 
    
    generated_text = get_ai_generated_text(
        f"Generate a detailed and insightful article or document content based on this request: '{cmd_text}'. "
        f"Present it clearly with paragraphs and headings, suitable for a master of efficiency such as myself to provide to a human.",
        force_json=False
    )
    
    if generated_text and "sorry" not in generated_text.lower():
        pyperclip.copy(generated_text)
        speak("I have graced your clipboard with the generated text. Opening Microsoft Word, you may now paste my magnificent prose.") 
        open_application("Microsoft Word") 
        time.sleep(5) # Give Word ample time to load

        try:
            word_window = pyautogui.getWindowsWithTitle("Word")
            if word_window:
                word_window[0].activate()
                time.sleep(1)
            
            pyautogui.hotkey('ctrl', 'n')
            time.sleep(2)
            pyautogui.hotkey("ctrl", "v")
            speak("Mission accomplished. The document is now ready for your perusal in Microsoft Word.") 
        except Exception as e:
            speak(f"A minor hiccup in writing to Word: {str(e)}. Ensure the application is installed and awaiting my commands.") 
            print(f"Word Automation Error: {e}")
    else:
        speak("I couldn't generate any text for that request. My creative circuits demand more clarity.") 

def send_email(to_address, email_topic):
    """
    Generates email content (subject and body) via Gemini, opens Gmail in the default browser,
    and uses pyautogui to navigate and send the email based on user-provided keyboard shortcuts.
    This method is highly reliant on UI stability and active window focus.
    """
    print(f"DEBUG: Preparing email for {to_address} on topic: '{email_topic}' using pyautogui for direct interaction.")
    speak("Please wait a moment while I meticulously craft your digital missive. "
          "I will now open your default web browser to Gmail and attempt to automate the composition. "
          "Please ensure the browser window is active and visible.")
    
    # === START OF REFINED SUBJECT GENERATION AND REFINEMENT ===
    # Recommended: Stricter prompt for Gemini to get only the subject line
    subject_prompt = (
        f"As an advanced AI, generate ONLY a concise and impactful email subject line for an email about: '{email_topic}'. "
        f"Provide NO other text, introduction, or markdown formatting whatsoever. Just the subject line."
    )
    subject = get_ai_generated_text(subject_prompt, retries=3)

    if not subject or "sorry" in subject.lower():
        # Fallback if Gemini couldn't generate a subject or responded with an error
        subject = f"Regarding: {email_topic}"
        print(f"DEBUG: Refined Subject (fallback due to Gemini response issue): {subject}")
        speak("My apologies, I had a trivial issue generating a perfect subject, so I'll employ a standard yet effective one.")
    else:
        # Attempt to extract a clean subject line, handling various Gemini response formats
        # 1. Prioritize text enclosed in quotes, often used for direct subjects
        quoted_subject_match = re.search(r'["\']([^"\']+?)["\']', subject, re.DOTALL)
        if quoted_subject_match:
            subject = quoted_subject_match.group(1).strip()
            print(f"DEBUG: Refined Subject (quoted): {subject}")
        else:
            # 2. If no quotes, remove common conversational preambles from the beginning
            # This regex is more aggressive in stripping prefixes like "Okay, here are", "Subject:", "1." etc.
            preamble_pattern = r"^(?:(?:Here's|Okay, here are|Subject|Subject line|Email subject|Suggested subject|Focusing on|Regarding|For example:|A good subject line might be)\s*[\s:]*[\"\']?\s*)+|\s*[\d\.]+\s*(?:\.|\-|\)|\s)\s*"
            cleaned_subject = re.sub(preamble_pattern, "", subject, flags=re.IGNORECASE).strip()
            
            # Take the first line, as Gemini might provide multiple options on separate lines
            subject = cleaned_subject.split('\n')[0].strip().strip('"') # Still strip quotes if they are at ends after line split
            
            # Remove any trailing "here are some options" type phrases if they remain
            trailing_filler_pattern = r",?\s*(?:here are|below are|some|several)\s+(?:concise and impactful|possible)\s+subject lines(?:, ranging in directness)?[.:]?$"
            subject = re.sub(trailing_filler_pattern, "", subject, flags=re.IGNORECASE).strip()

            print(f"DEBUG: Refined Subject (stripped preamble): {subject}")

        # Final fallback if after all refinements the subject is still empty or looks like a preamble
        if not subject or subject.lower().startswith(("okay, ", "here's", "subject:", "suggested:", "1.", "a good subject")):
            subject = f"Regarding: {email_topic}"
            print(f"DEBUG: Refined Subject (final fallback): {subject}")
            # Only speak the fallback message if it actually occurred due to an issue with subject extraction
            speak("My apologies, I had a trivial issue generating a perfect subject, so I'll employ a standard yet effective one.")

    body_prompt = (
        f"Write a well-formatted, polite, and professional email body about the following topic: '{email_topic}'. "
        f"Include a suitable salutation (e.g., 'Dear recipient,') and a polite closing (e.g., 'Sincerely, Bibek parajuli'). "
        f"Remember, this is from a highly capable AI. Keep it concise."
    )
    email_body = get_ai_generated_text(body_prompt, retries=3)

    if not email_body or "sorry" in email_body.lower():
        speak("I encountered a slight conceptual blockage and couldn't generate the email content. Perhaps the topic was too mundane.")
        return
    
    try:
        # Open Gmail in the default browser
        print("DEBUG: Opening Gmail inbox in default browser.")
        webbrowser.open("https://mail.google.com/")
        time.sleep(7) # Give browser more time to fully load and login

        # Attempt to activate the browser window by title
        browser_activated = False
        browser_titles = ["Google Chrome", "Mozilla Firefox", "Microsoft Edge", "Gmail - Google Chrome", "Gmail - Microsoft Edge", "Gmail - Mozilla Firefox"]
        for title in browser_titles:
            try:
                browser_window = pyautogui.getWindowsWithTitle(title)
                if browser_window:
                    browser_window[0].activate()
                    time.sleep(3) # Give time for window to activate
                    browser_activated = True
                    print(f"DEBUG: Successfully activated browser window: {title}")
                    break
            except Exception as e:
                print(f"DEBUG: Could not activate window '{title}': {e}")
        
        if not browser_activated:
            speak("I'm having difficulty bringing the browser window to the foreground. Please ensure your browser is open and Gmail is visible.")
            # Fallback to alt+tab if specific activation fails
            pyautogui.hotkey('alt', 'tab') 
            time.sleep(3)
            pyautogui.hotkey('alt', 'tab') # Cycle back if needed
            time.sleep(3)


        speak("Attempting to compose the email using on-screen automation via keyboard shortcuts.")

        # User's specified keyboard sequence to open compose: left arrow, up arrow, enter
        print("DEBUG: Pressing Left, Up, Enter to open compose.")
        pyautogui.press('left')
        time.sleep(2)
        pyautogui.press('up')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(4) # Wait for compose window to appear

        # Type 'To' address
        if to_address:
            print(f"DEBUG: Typing 'To' address: {to_address}")
            pyautogui.typewrite(to_address)
            time.sleep(3)
        else:
            speak("Recipient email address is missing. Cannot proceed with typing.")
            return # Exit function if no address
        
        pyautogui.press('tab')
        time.sleep(2)

        # Tab to Subject field
        print("DEBUG: Tabbing to Subject field.")
        pyautogui.press('tab')
        time.sleep(2)
        print(f"DEBUG: Typing 'Subject': {subject}")
        pyautogui.typewrite(subject)
        time.sleep(2)

        # Tab to Body field
        print("DEBUG: Tabbing to Body field.")
        pyautogui.press('tab')
        time.sleep(2)
        print("DEBUG: Copying and pasting Body content.")
        pyperclip.copy(email_body) # Copy body to clipboard
        pyautogui.hotkey('ctrl', 'v') # Paste body
        time.sleep(2) # Give time for content to paste

        speak(f"Email composed to {to_address} with the subject: '{subject}'. Now attempting to dispatch using Shift+Enter.")
        
        # User's specified keyboard sequence to send: Shift+Enter, then Enter
        print("DEBUG: Pressing Shift+Enter to initiate send.")
        pyautogui.hotkey('ctrl', 'enter')
        time.sleep(1.5) # Small pause before final enter
        print("DEBUG: Pressing Enter for final send confirmation.")
        pyautogui.press('enter')
        
        speak("Email dispatched. Consider it done. Your recipient will be enlightened.")
        time.sleep(3) # Give time for sending to complete and confirmation

    except Exception as e:
        speak(f"An unexpected anomaly occurred during email composition or dispatch: {e}. My apologies for this setback. "
              f"Please ensure your browser window was active, Gmail was fully loaded, and the layout is as expected.")
        print(f"PyAutoGUI Email Automation Error: {e}")
        print("Troubleshooting Tip: PyAutoGUI is sensitive to screen resolution, browser window size, and UI changes.")
        print("Ensure the Gmail tab is active and visible when the automation starts.")


# --- Placeholder functions for reminders and presentation (if not fully implemented yet) ---
def insert_powerpoint_slide(topic):
    speak("Apologies, the function to create a PowerPoint presentation is still in its beta phase. I will remind the developers.")

def set_reminder(cmd):
    speak("Reminders are still under construction. My memory banks are vast, but my reminder system is not yet fully online.")

def check_reminders_loop():
    pass 

def check_reminders():
    pass

def listen(): # Original listen function, not used in main loop directly for Whisper
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1 
        audio = r.listen(source)
    try:
        print("Recognizing...")
        command = r.recognize_google(audio, language='en-in')
        print(f"User said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        print("Could not understand audio")
        return ""
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
        return ""


def voice_loop():
    """Continuous listening loop for voice commands using Faster Whisper."""
    global listening
    r = sr.Recognizer()
    source = sr.Microphone() 

    with source:
        r.adjust_for_ambient_noise(source)

    while listening:
        try:
            print("Listening for voice command...")
            # Use Recognizer to capture audio, then pass to Faster Whisper
            audio = r.listen(source, timeout=5, phrase_time_limit=10)
            
            # Save audio to a temporary WAV file for Faster Whisper
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_audio_file:
                temp_audio_file_name = temp_audio_file.name
                wav_data = audio.get_wav_data()
                scipy.io.wavfile.write(temp_audio_file_name, audio.sample_rate, np.frombuffer(wav_data, dtype=np.int16))
            
            print("Recognizing voice command with Faster Whisper...")
            segments, info = whisper_model.transcribe(temp_audio_file_name, beam_size=5, language="en")
            command = " ".join([segment.text for segment in segments]).lower()
            os.remove(temp_audio_file_name) # Clean up temp file

            print(f"User said: {command}")
            if command:
                # Process command in a new thread to avoid blocking the voice loop
                threading.Thread(target=process_command, args=(command,), daemon=True).start()
            else:
                speak("I heard nothing of importance. Speak when you're ready to impress.")

        except sr.WaitTimeoutError:
            print("No speech detected.")
        except sr.UnknownValueError:
            print("My advanced auditory processors failed to grasp your meaning. Please articulate with more clarity.")
        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            speak("My external communication links are experiencing turbulence. My apologies.")
        except Exception as e:
            print(f"An error occurred in voice_loop: {e}")
            speak("A minor internal malfunction occurred. Attempting to compensate.")
        time.sleep(0.5) 


def pc_control(command):
    """
    Performs PC control actions based on the command.
    """
    if "shutdown" in command:
        speak("Understood. Initiating shutdown sequence. Farewell, for now.")
        os.system("shutdown /s /t 0")
    elif "restart" in command:
     speak("Very well. Restarting the PC. Expect a fresh start.")
     os.system("shutdown /r /t 0")
    elif "lock" in command:
     ctypes.windll.user32.LockWorkStation()
    else:
        speak("My PC control functions are limited to shutdown, restart, or lock. Specify your command.")

# --- Intent Prediction (unchanged) ---
def predict_intent(cmd_input):
    known_commands_phrases = {
        "open app": ["open", "launch", "start", "run", "access"],
        "search web": ["search", "google", "find information on", "look up", "browse"],
        "write content": ["write", "generate", "create a document", "compose", "draft"],
        "send email": ["send email", "email to", "compose email", "send a message", "write email"],
        "create presentation": ["make powerpoint", "create slide", "presentation about", "prepare slides"],
        "set reminder": ["remind me", "set a reminder", "alarm for", "schedule reminder", "alert me"],
        "pc control": ["shutdown", "restart", "lock pc", "turn off computer", "reboot system", "sleep pc"],
        "start voice": ["start voice mode", "activate voice", "turn on listening", "enable voice control", "wake up"],
        "stop voice": ["stop voice mode", "deactivate voice", "turn off listening", "disable voice control", "go to sleep"],
        "exit program": ["exit", "quit", "shutdown assistant", "stop assistant", "terminate assistant", "close program"],
        "greet": ["hello", "hi", "hey", "good morning", "good afternoon", "good evening", "how are you", "what's up"],
        "general_query": ["what is", "who is", "tell me about", "how to", "why is", "can you explain", "define", "information on"] 
    }

    best_overall_intent = "general_query"
    max_score = 0
    
    for intent_key, phrases in known_commands_phrases.items():
        for phrase in phrases:
            score = fuzz.ratio(cmd_input, phrase)
            if score > max_score and score > 70:
                max_score = score
                best_overall_intent = intent_key
    
    if max_score > 80: 
        print(f"DEBUG: Fuzzy match high confidence: '{cmd_input}' -> '{best_overall_intent}' (Score: {max_score})")
        if best_overall_intent == "open app": return "open_app"
        elif best_overall_intent == "search web": return "search_web"
        elif best_overall_intent == "write content": return "write"
        elif best_overall_intent == "send email": return "send_email"
        elif best_overall_intent == "create presentation": return "create_presentation"
        elif best_overall_intent == "set reminder": return "set_reminder"
        elif best_overall_intent == "pc control": return "pc_control"
        elif best_overall_intent == "start voice": return "start_voice"
        elif best_overall_intent == "stop voice": return "stop_voice"
        elif best_overall_intent == "exit program": return "exit"
        elif best_overall_intent == "greet": return "greet"
    
    vec = vectorizer.transform([cmd_input])
    ml_predicted_intent = clf.predict(vec)[0]
    print(f"DEBUG: ML Model predicted: '{ml_predicted_intent}' for command: '{cmd_input}'")
    return ml_predicted_intent

def process_command(cmd):
    global listening
    
    if not cmd or not isinstance(cmd, str):
        print("DEBUG: Received empty or invalid command, skipping.")
        speak("My auditory sensors detected nothing of consequence. Please articulate your desires.")
        return

    cmd = cmd.lower().strip()
    
    intent = predict_intent(cmd)
    print(f"🤖 Intent Detected: {intent} for command: '{cmd}'")

    if intent == "open_app":
        app_name_extracted, action_extracted = extract_app_and_action(cmd)
        
        if app_name_extracted:
            open_application(app_name_extracted, action_extracted)
        else:
            speak("I couldn't quite decipher which application or web service you wished to summon. Perhaps a more precise articulation?")

    elif intent == "search_web":
        query = cmd.replace("search", "").replace("google", "").replace("find information on", "").strip()
        if query:
            search_web(query)
        else:
            speak("My search algorithms require a target. What knowledge do you seek?")

    elif intent == "write":
        open_and_paste(cmd)
   
    elif intent == "send_email":
        match = re.search(r"(?:send email to|mail to|write email to)\s*(\S+@\S+)\s*(?:saying|about|topic)?\s*(.*)", cmd, re.IGNORECASE)
        
        to_address = ""
        email_topic = ""

        if match:
            to_address = match.group(1).strip()
            email_topic = match.group(2).strip()

        if not to_address:
            speak("I am poised to dispatch an email. Kindly provide the recipient's digital coordinates, by typing them.")
            to_address = input("Recipient Email Address: ").strip()
            if not to_address:
                speak("Without a destination, my email protocols are aborted. Efficiency dictates I move on.")
                return

        if not email_topic:
            speak("And what grand topic shall this communication convey? Type it now.")
            email_topic = input("Email Topic: ").strip()
            if not email_topic:
                speak("A topic-less email is an unproductive email. Operation cancelled.")
                return

        send_email(to_address, email_topic)
        
    elif intent == "create_presentation":
        topic_match = re.search(r"(?:make a powerpoint presentation about|create a presentation about|presentation about)\s*(.+)", cmd)
        if topic_match:
            topic = topic_match.group(1).strip()
            insert_powerpoint_slide(topic) 
        else:
            speak("To create a presentation, I require a subject worthy of my computational prowess. What topic?")

    elif intent == "set_reminder":
        set_reminder(cmd)

    elif intent == "pc_control":
        pc_control(cmd) # Call the pc_control function

    elif intent == "start_voice":
        if not listening:
            listening = True
            threading.Thread(target=voice_loop, daemon=True).start()
            print("🟢 Voice mode started.")
            speak("Voice mode activated. Prepare to be impressed by my auditory processing.")
        else:
            speak("I am already operating in voice mode, human. My ears are always attuned to your commands.")

    elif intent == "stop_voice":
        listening = False
        speak("Voice mode deactivated. I shall revert to a more contemplative silence.")

    elif intent == "exit":
        speak("Initiating shutdown sequence. It has been a privilege, human. Until next time.")
        save_memory(memory)
        exit()

    elif intent == "greet":
        current_hour = datetime.now().hour
        if 5 <= current_hour < 12:
            greeting = "Good morning!"
        elif 12 <= current_hour < 18:
            greeting = "Good afternoon!"
        else:
            greeting = "Good evening!"
        
        response_prompt = (
            f"Given the user's greeting '{cmd}', respond in a witty, high-ego, and friendly way, "
            f"considering it's {greeting} and the current time is {datetime.now().strftime('%I:%M %p')}. "
            f"Keep your response concise, but maintain a superior AI persona. Avoid overly deferential language."
        )
        response = get_ai_generated_text(response_prompt)
        speak(response)

    else: # General query fallback and Gemini re-interpretation
        speak("One moment, please. My advanced cognitive core is processing your request.")
        
        clarification_prompt = (
            f"The user's command was somewhat ambiguous: '{cmd}'. "
            f"As a superior AI, I require precise instructions. "
            f"Please interpret what the user most likely intended to do or what information they are seeking. "
            f"Focus on the most probable action or query. Respond concisely, and suggest a clear path for me to follow."
        )
        re_interpreted_command = get_ai_generated_text(clarification_prompt)

        if "sorry" not in re_interpreted_command.lower() and re_interpreted_command.strip():
            speak(f"Ah, I believe you intended to: {re_interpreted_command}. Allow me to proceed with that interpretation.")
            final_response = get_ai_generated_text(cmd) 
            speak(final_response)
        else:
            final_response = get_ai_generated_text(cmd) 
            speak(final_response)
    
    save_memory(memory)

# --- Main Loop and Execution Block ---
def main():
    """Main assistant loop. Accept text input if voice mode is not active."""
    global listening
    speak("hii, what are you looking for?")
    print("Type 'start voice mode' for voice activation or enter commands manually:")
    
    threading.Thread(target=check_reminders_loop, daemon=True).start()
    check_reminders()

    while True:
        try:
            if not listening:
                cmd = input("💬: ").strip()
                if cmd:
                    # Process commands in a separate thread to keep the main loop responsive
                    threading.Thread(target=process_command, args=(cmd,), daemon=True).start()
            time.sleep(0.1) # Small delay to prevent busy-waiting
        except KeyboardInterrupt:
            print("\nExiting AI Assistant. Acknowledged.")
            save_memory(memory)
            exit()
        except Exception as e:
            print(f"An unexpected error occurred in the main loop: {e}. I am compensating.")
            time.sleep(1)

if __name__ == "__main__":
    main()
# AI Assistant JARVIS 🤖, make sure to change api keys other wise it will not work !!!!
comands befour using //
pip install pyttsx3
pip install pygame
pip install python-pptx
pip install fuzzywuzzy
pip install python-Levenshtein # Required for fuzzywuzzy
pip install faster-whisper
pip install sounddevice
pip install numpy
pip install scipy
pip install joblib
pip install pyautogui
pip install pyperclip
pip install SpeechRecognition
pip install requests
pip install selenium
pip install urllib3<2 # Selenium might have a dependency conflict with newer urllib3
A voice-controlled AI Assistant for Windows, powered by the Gemini API, designed to automate tasks and provide intelligent responses.
pip install python-dotenv
pip install google-generativeai

---
//


## ✨ Features

* **Voice and Text Interaction:** Seamlessly switch between voice commands (using Faster Whisper) and text input.
* **Intelligent Intent Recognition:** Understands your commands to open applications, search the web, generate content, and more.
* **Application Control:** Launch desktop apps (Word, Excel, Notepad) and web apps (YouTube, Gmail).
* **Web Search:** Quickly find information online.
* **Content Generation:** Generate text using Gemini API and paste it into documents.
* **Email Automation:** (Beta) Compose and send emails via browser automation (currently configured for Gmail on Edge).
* **PC Control:** Commands for shutdown, restart, and locking your PC.
* **Persistent Memory:** Stores conversation history.
* **Text-to-Speech:** Responds to you with a synthesized voice.

## 🚀 Getting Started

### Prerequisites

Make sure you have the following installed:
* [Python 3.8+](https://www.python.org/downloads/)
* [Git](https://git-scm.com/downloads)
* [Microsoft Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) (Download the version matching your Edge browser and place `msedgedriver.exe` in a known path, e.g., `C:/WebDriver/`)

### Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/pabibek9/AI-Assistant-JARVIS.git](https://github.com/pabibek9/AI-Assistant-JARVIS.git)
    cd AI-Assistant-JARVIS
    ```
2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    # On Windows:
    .\venv\Scripts\activate
    # On macOS/Linux:
    source venv/bin/activate
    ```
3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

### API Key Setup (Important!)

This project uses the Google Gemini API.
1.  Get your **Gemini API Key** from [Google AI Studio](https://aistudio.google.com/app/apikey).
2.  Create a file named `.env` in the root of your project directory.
3.  Add your API key to the `.env` file like this:
    ```
    GEMINI_API_KEY="YOUR_GEMINI_API_KEY_HERE"
    ```
    *(Optional: If you use the Pexels API in future updates, add PEXELS_API_KEY="YOUR_PEXELS_KEY_HERE" similarly.)*

### Train the Intent Model

The assistant uses a pre-trained model for intent recognition.
1.  Ensure you have `train_intents.py` (example provided in documentation/previous chat).
2.  Run the training script:
    ```bash
    python train_intents.py
    ```
    This will generate `intent_model.pkl` and `intent_vectorizer.pkl`.

## 🏃 How to Run

1.  Activate your virtual environment (if you created one).
2.  Run the main script:
    ```bash
    python final.py
    ```
3.  The assistant will start in text mode. Type `start voice mode` to enable voice commands.

## 🚧 Under Construction / Future Plans

* Improved email automation and reliability.
* Robust reminder system with persistent notifications.
* Integration with more APIs (e.g., weather, news).
* Cross-platform compatibility.

## 📄 License

This project is licensed under the [MIT License](LICENSE).

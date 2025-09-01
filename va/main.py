import os
import subprocess
import pyautogui
import webbrowser
import pyttsx3
import google.generativeai as genai
from dotenv import load_dotenv
import speech_recognition as sr
import json
import pygetwindow as gw
import time

# ---------- LOAD ENV ----------
load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# ---------- SPEAK ----------
engine = pyttsx3.init()
def speak(text):
    engine.say(text)
    engine.runAndWait()

# ---------- GLOBAL CONTEXT ----------
context = {
    "last_app": None,       # e.g., "chrome", "notepad"
    "last_action": None,    # e.g., "open_app", "youtube_search"
    "last_folder": None     # e.g., "C:\\Users\\User\\Downloads"
}

# ---------- GEMINI PARSER ----------
def parse_command_online(command: str):
    prompt = f"""
    You are a command parser for a Windows voice assistant.
    Convert the following instruction into JSON.

    Allowed actions:
    - open_app → open an app or a system folder (params: app)
    - open_file → open a file in the last opened folder (params: file)
    - move_file → move a file (params: source, destination)
    - shutdown → shut down the system
    - restart → restart the system
    - screenshot → take a screenshot (params: optional path)
    - youtube_search → search and play something on YouTube (params: query)
    - youtube_control → control playback (params: command: play, pause, stop, next, previous, volume_up, volume_down)
    - web_search → search the internet on Google (params: query)
    - app_control → control open apps/windows (params: app, command: close, minimize, maximize, restore, focus)

    Instruction: "{command}"

    Respond ONLY in JSON.
    """

    try:
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(prompt)
        text = response.text.strip()

        if text.startswith("```"):
            text = text.replace("```json", "").replace("```", "").strip()

        return json.loads(text)
    except Exception as e:
        print("Parsing error:", e)
        return {"action": "unknown", "params": {}}

# ---------- EXECUTE ----------
def execute(action, params):
    global context

    folder_map = {
        "download": "downloads", "downloads": "downloads",
        "documents": "documents", "document": "documents",
        "desktop": "desktop",
        "pictures": "pictures", "photos": "pictures", "images": "pictures",
        "music": "music", "songs": "music", "audio": "music",
        "videos": "videos", "video": "videos", "movies": "videos"
    }

    # ---------- OPEN APP / FOLDER ----------
    if action == "open_app":
        app = params.get("app", "").lower()
        app = folder_map.get(app, app)

        folder_paths = {
            "downloads": os.path.expanduser("~\\Downloads"),
            "documents": os.path.expanduser("~\\Documents"),
            "desktop": os.path.expanduser("~\\Desktop"),
            "pictures": os.path.expanduser("~\\Pictures"),
            "music": os.path.expanduser("~\\Music"),
            "videos": os.path.expanduser("~\\Videos")
        }

        # If it's a known folder
        if app in folder_paths:
            try:
                folder_path = folder_paths[app]
                os.startfile(folder_path)
                speak(f"Opening {app}")
                context["last_app"] = app
                context["last_action"] = "open_app"
                context["last_folder"] = folder_path
            except Exception as e:
                speak(f"Failed to open {app}: {e}")
            return

        # Otherwise, try as executable
        try:
            ps_command = f"Start-Process '{app}'"
            subprocess.Popen(["powershell", "-Command", ps_command])
            speak(f"Opening {app}")
            context["last_app"] = app
            context["last_action"] = "open_app"
        except Exception:
            speak(f"Could not find or open {app}")

    # ---------- OPEN FILE IN LAST FOLDER ----------
    elif action == "open_file":
        file_name = params.get("file", "")
        folder_path = context.get("last_folder")

        if folder_path:
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                os.startfile(file_path)
                speak(f"Opening {file_name}")
            else:
                speak(f"Could not find {file_name} in {folder_path}")
        else:
            speak("No folder context found. Please open a folder first.")

    # ---------- MOVE FILE ----------
    elif action == "move_file":
        try:
            src = params.get("source")
            dst = params.get("destination")
            os.rename(src, dst)
            speak("File moved successfully")
        except Exception as e:
            speak(f"Failed to move file: {e}")

    # ---------- SHUTDOWN / RESTART ----------
    elif action == "shutdown":
        speak("Shutting down your computer")
        os.system("shutdown /s /t 1")
    elif action == "restart":
        speak("Restarting your computer")
        os.system("shutdown /r /t 1")

    # ---------- SCREENSHOT ----------
    elif action == "screenshot":
        path = params.get("path", "screenshot.png")
        pyautogui.screenshot(path)
        speak(f"Screenshot saved as {path}")

    # ---------- YOUTUBE SEARCH ----------
    elif action == "youtube_search":
        query = params.get("query", "")
        browser = context.get("last_app")
        if browser in ["chrome", "firefox", "edge"]:
            windows = gw.getWindowsWithTitle(browser)
            if windows:
                windows[0].activate()
                pyautogui.hotkey("ctrl", "t")
                time.sleep(0.2)
                pyautogui.typewrite(f"https://www.youtube.com/results?search_query={query}")
                pyautogui.press("enter")
                speak(f"Searching YouTube for {query} in {browser}")
            else:
                webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
                speak(f"Searching YouTube for {query}")
        else:
            webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
            speak(f"Searching YouTube for {query}")
        context["last_action"] = "youtube_search"

    # ---------- WEB SEARCH ----------
    elif action == "web_search":
        query = params.get("query", "")
        browser = context.get("last_app")
        if browser in ["chrome", "firefox", "edge"]:
            windows = gw.getWindowsWithTitle(browser)
            if windows:
                windows[0].activate()
                pyautogui.hotkey("ctrl", "t")
                time.sleep(0.2)
                pyautogui.typewrite(f"https://www.google.com/search?q={query}")
                pyautogui.press("enter")
                speak(f"Searching Google for {query} in {browser}")
            else:
                webbrowser.open(f"https://www.google.com/search?q={query}")
                speak(f"Searching Google for {query}")
        else:
            webbrowser.open(f"https://www.google.com/search?q={query}")
            speak(f"Searching Google for {query}")
        context["last_action"] = "web_search"

    # ---------- YOUTUBE CONTROL ----------
    elif action == "youtube_control":
        cmd = params.get("command", "").lower()
        if cmd in ["pause", "play", "stop"]:
            pyautogui.press("k")
        elif cmd == "next":
            pyautogui.hotkey("shift", "n")
        elif cmd == "previous":
            pyautogui.hotkey("shift", "p")
        elif cmd == "volume_up":
            pyautogui.press("volumeup")
        elif cmd == "volume_down":
            pyautogui.press("volumedown")
        speak(f"YouTube {cmd}")

    # ---------- APP CONTROL ----------
    elif action == "app_control":
        app_name = params.get("app", "").lower()
        cmd = params.get("command", "").lower()
        if not app_name and context.get("last_app"):
            app_name = context["last_app"]

        try:
            windows = gw.getWindowsWithTitle(app_name)
            if not windows:
                speak(f"No open windows found for {app_name}")
                return
            window = windows[0]

            if cmd == "close":
                window.close()
            elif cmd == "minimize":
                window.minimize()
            elif cmd == "maximize":
                window.maximize()
            elif cmd == "restore":
                window.restore()
            elif cmd == "focus":
                window.activate()
            else:
                speak(f"Unknown command {cmd} for {app_name}")
                return
            speak(f"{cmd.capitalize()} executed on {app_name}")
        except Exception as e:
            speak(f"Failed to control {app_name}: {e}")

        context["last_app"] = app_name
        context["last_action"] = "app_control"

    else:
        speak("Sorry, I don’t understand that command.")

# ---------- LISTEN & EXECUTE ----------
def listen_and_execute():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio)
        print("You said:", command)
        result = parse_command_online(command)
        print("Gemini output:", result)
        execute(result.get("action"), result.get("params"))
    except Exception as e:
        print("Error:", e)
        speak("Sorry, I did not understand that.")

# ---------- MAIN LOOP ----------
if __name__ == "__main__":
    while True:
        listen_and_execute()
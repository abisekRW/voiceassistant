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
import re
from difflib import SequenceMatcher
from send2trash import send2trash  # for deleting files safely

# ---------- LOAD ENV ----------
load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# ---------- SPEAK ----------
engine = pyttsx3.init()
def speak(text):
    print("Assistant:", text)
    engine.say(text)
    engine.runAndWait()

# ---------- GLOBAL CONTEXT ----------
context = {
    "last_app": None,
    "last_action": None,
    "last_folder": None,
    "last_file": None
}

# ---------- NORMALIZATION / MATCHING HELPERS ----------
EXT_WORDS = {
    "pdf","doc","docx","ppt","pptx","xls","xlsx","txt",
    "png","jpg","jpeg","csv","json","mp3","mp4","mkv","py","zip","rar","7z"
}

def normalize_text(s: str) -> str:
    s = s.lower()
    s = os.path.splitext(s)[0]
    s = re.sub(r'[\W_]+', '', s, flags=re.UNICODE)
    return s

def extract_query_and_ext(file_query: str):
    tokens = re.findall(r'\w+', file_query.lower())
    ext = None
    for t in tokens[::-1]:
        if t in EXT_WORDS:
            ext = t
            break
    core_tokens = [t for t in tokens if t not in EXT_WORDS]
    core_query = ''.join(core_tokens)
    return core_query, ext

def rank_matches(files, folder_path, raw_query):
    q_norm_core, ext_hint = extract_query_and_ext(raw_query)
    tokens = re.findall(r'\w+', raw_query.lower())
    tokens = [t for t in tokens if t not in EXT_WORDS]

    ranked = []
    for f in files:
        full = os.path.join(folder_path, f)
        if not os.path.isfile(full):
            continue
        f_lower = f.lower()
        if ext_hint and not f_lower.endswith("." + ext_hint):
            continue
        f_norm = normalize_text(f)
        substr = 1.0 if (q_norm_core and q_norm_core in f_norm) else 0.0
        ratio = SequenceMatcher(None, q_norm_core, f_norm).ratio() if q_norm_core else 0.0
        token_hits = sum(1 for t in tokens if t and t in f_norm) * 0.15
        score = ratio + substr + token_hits
        ranked.append((score, f))

    if not ranked:
        for f in files:
            full = os.path.join(folder_path, f)
            if not os.path.isfile(full):
                continue
            f_norm = normalize_text(f)
            substr = 1.0 if (q_norm_core and q_norm_core in f_norm) else 0.0
            ratio = SequenceMatcher(None, q_norm_core, f_norm).ratio() if q_norm_core else 0.0
            token_hits = sum(1 for t in tokens if t and t in f_norm) * 0.15
            score = ratio + substr + token_hits
            ranked.append((score, f))

    ranked.sort(key=lambda x: x[0], reverse=True)
    ranked = [f for score, f in ranked if score >= 0.55]
    return ranked

def find_file_in_folder(file_name, folder_path):
    """
    Search for a file matching `file_name` in `folder_path`.
    Returns the full path of the first match or None.
    """
    try:
        files = os.listdir(folder_path)
        matches = rank_matches(files, folder_path, file_name)
        if matches:
            return os.path.join(folder_path, matches[0])
    except Exception as e:
        print("Find file error:", e)
    return None

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

    folder_paths = {
        "downloads": os.path.expanduser("~\\Downloads"),
        "documents": os.path.expanduser("~\\Documents"),
        "desktop": os.path.expanduser("~\\Desktop"),
        "pictures": os.path.expanduser("~\\Pictures"),
        "music": os.path.expanduser("~\\Music"),
        "videos": os.path.expanduser("~\\Videos"),
        "recycle bin": None
    }

    # ---------- OPEN APP / FOLDER ----------
    if action == "open_app":
        app = params.get("app", "").lower()
        app = folder_map.get(app, app)

        if app in folder_paths and folder_paths[app]:
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

        try:
            ps_command = f"Start-Process '{app}'"
            subprocess.Popen(["powershell", "-Command", ps_command])
            speak(f"Opening {app}")
            context["last_app"] = app
            context["last_action"] = "open_app"
        except Exception:
            speak(f"Could not find or open {app}")

    # ---------- OPEN FILE ----------
    elif action == "open_file":
        raw_query = params.get("file", "")
        folder_path = context.get("last_folder")
        if folder_path:
            try:
                files = os.listdir(folder_path)
                matches = rank_matches(files, folder_path, raw_query)
                if not matches:
                    speak(f"No file matching {raw_query} found in {folder_path}")
                    return
                chosen_file = matches[0] if len(matches) == 1 else None

                if not chosen_file:
                    limited_matches = matches[:5]
                    speak(f"I found {len(matches)} matching files. Top {len(limited_matches)}:")
                    for i, f in enumerate(limited_matches, start=1):
                        speak(f"Option {i}: {f}")

                    speak("Say the number or part of the file name to open (5s to respond)...")
                    r = sr.Recognizer()
                    with sr.Microphone() as source:
                        r.adjust_for_ambient_noise(source)
                        try:
                            audio = r.listen(source, timeout=5, phrase_time_limit=5)
                            choice = r.recognize_google(audio).lower()
                            print("User choice:", choice)
                            if choice.isdigit():
                                idx = int(choice) - 1
                                if 0 <= idx < len(limited_matches):
                                    chosen_file = limited_matches[idx]
                            else:
                                choice_norm = normalize_text(choice)
                                for f in limited_matches:
                                    if choice_norm in normalize_text(f):
                                        chosen_file = f
                                        break
                        except sr.WaitTimeoutError:
                            speak("No response detected. Opening the first file.")
                            chosen_file = limited_matches[0]
                        except Exception as e:
                            print("Choice error:", e)
                            speak("Error. Opening first file.")
                            chosen_file = limited_matches[0]

                file_path = os.path.join(folder_path, chosen_file)
                os.startfile(file_path)
                speak(f"Opening {chosen_file}")
                context["last_file"] = file_path

            except Exception as e:
                speak(f"Error opening file: {e}")
        else:
            speak("No folder context found. Please open a folder first.")

    # ---------- MOVE FILE ----------
    elif action == "move_file":
        try:
            src = params.get("source")
            dst = params.get("destination")

            if not src and context.get("last_file"):
                src = context["last_file"]

            if not src or not dst:
                speak("Please specify both source and destination.")
                return

            # Search for source in last folder
            last_folder = context.get("last_folder", os.path.expanduser("~"))
            src_full = src if os.path.isabs(src) else find_file_in_folder(src, last_folder)

            if not src_full:
                speak(f"File not found: {src}")
                return

            # Map destination folder
            dst_lower = dst.lower()
            if dst_lower in folder_paths and folder_paths[dst_lower]:
                dst_full = os.path.join(folder_paths[dst_lower], os.path.basename(src_full))
            elif dst_lower == "recycle bin":
                send2trash(src_full)
                speak(f"{os.path.basename(src_full)} sent to Recycle Bin")
                context["last_file"] = None
                return
            else:
                dst_full = dst  # assume full path

            os.rename(src_full, dst_full)
            speak(f"File moved successfully to {dst_full}")
            context["last_file"] = dst_full

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
        audio = r.listen(source, timeout=5, phrase_time_limit=5)

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
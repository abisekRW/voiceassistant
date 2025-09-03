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

# ---------- FILE EXTENSIONS ----------
EXT_WORDS = {
    "pdf","doc","docx","ppt","pptx","xls","xlsx","txt",
    "png","jpg","jpeg","jfif","csv","json","mp3","mp4","mkv","py","zip","rar","7z"
}

IMAGE_EXT = {"png","jpg","jpeg","jfif"}

# ---------- NORMALIZATION / MATCHING HELPERS ----------
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
    query_tokens = re.findall(r'\w+', raw_query.lower())

    ranked = []
    for f in files:
        full = os.path.join(folder_path, f)
        if not os.path.isfile(full):
            continue

        # Filter by extension if mentioned or if "image" is in query
        if ext_hint:
            if not f.lower().endswith("." + ext_hint):
                continue
        elif "image" in raw_query.lower() or "photo" in raw_query.lower():
            if not f.lower().endswith(tuple(IMAGE_EXT)):
                continue

        f_tokens = re.findall(r'\w+', f.lower())
        token_hits = sum(1 for qt in query_tokens if qt in f_tokens)
        ratio = SequenceMatcher(None, q_norm_core, normalize_text(f)).ratio()
        score = token_hits + ratio
        ranked.append((score, f))

    ranked.sort(key=lambda x: x[0], reverse=True)
    ranked = [f for score, f in ranked if score > 0]
    return ranked

def find_file_in_folder(file_name, folder_path):
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

        # Check if folder
        if app in folder_paths and folder_paths[app]:
            try:
                folder_path = folder_paths[app]
                os.startfile(folder_path)
                speak(f"Opening {app}")
                context["last_app"] = app
                context["last_action"] = "open_app"
                context["last_folder"] = folder_path
                return
            except Exception as e:
                speak(f"Failed to open {app}: {e}")
                return

        # Otherwise, search file in known folders
        for folder in folder_paths.values():
            if not folder: 
                continue
            file_path = find_file_in_folder(app, folder)
            if file_path:
                os.startfile(file_path)
                speak(f"Opening {os.path.basename(file_path)}")
                context["last_file"] = file_path
                return

        # Else try executable
        try:
            ps_command = f"Start-Process '{app}'"
            subprocess.Popen(["powershell", "-Command", ps_command])
            speak(f"Opening {app}")
            context["last_app"] = app
            context["last_action"] = "open_app"
        except Exception:
            speak(f"Could not find or open {app}")
        return

    # ---------- OPEN FILE ----------
    elif action == "open_file":
        raw_query = params.get("file", "").replace(" ", "")
        folder_path = context.get("last_folder")
        if folder_path:
            file_path = find_file_in_folder(raw_query, folder_path)
            if file_path:
                os.startfile(file_path)
                speak(f"Opening {os.path.basename(file_path)}")
                context["last_file"] = file_path
            else:
                speak(f"No matching file found for {raw_query}")
        else:
            speak("No folder context found. Please open a folder first.")

    # ---------- MOVE FILE ----------
    elif action == "move_file":
        try:
            src = params.get("source", "").replace(" ", "")
            dst = params.get("destination", "")

            if not src and context.get("last_file"):
                src = context["last_file"]

            if not src or not dst:
                speak("Please specify both source and destination.")
                return

            last_folder = context.get("last_folder", os.path.expanduser("~"))
            src_full = src if os.path.isabs(src) else find_file_in_folder(src, last_folder)

            if not src_full:
                speak(f"File not found: {src}")
                return

            dst_lower = dst.lower()
            if dst_lower in folder_paths and folder_paths[dst_lower]:
                dst_full = os.path.join(folder_paths[dst_lower], os.path.basename(src_full))
            elif dst_lower == "recycle bin":
                send2trash(src_full)
                speak(f"{os.path.basename(src_full)} sent to Recycle Bin")
                context["last_file"] = None
                return
            else:
                dst_full = dst

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
        webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
        speak(f"Searching YouTube for {query}")
        context["last_action"] = "youtube_search"

    # ---------- WEB SEARCH ----------
    elif action == "web_search":
        query = params.get("query", "")
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
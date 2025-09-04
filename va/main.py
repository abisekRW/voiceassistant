import os
import re
import sys
import json
import webbrowser
import subprocess
import pyautogui
import pyttsx3
import google.generativeai as genai
from dotenv import load_dotenv
import speech_recognition as sr
import pygetwindow as gw
import numpy as np
import sounddevice as sd
import noisereduce as nr
from difflib import SequenceMatcher
from send2trash import send2trash

try:
    import winsound
    def beep_ok(): winsound.Beep(600, 180)
    def beep_err(): 
        winsound.Beep(420, 140); winsound.Beep(320, 140)
    def beep_ready(): winsound.Beep(750, 120)
except Exception:
    def beep_ok(): pass
    def beep_err(): pass
    def beep_ready(): pass

# ---------- ENV ----------
load_dotenv()
genai.configure(api_key=os.getenv("GOOGLE_API_KEY", ""))

# ---------- TTS ----------
engine = pyttsx3.init()
def speak(text: str):
    print("Assistant:", text)
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception:
        pass

# ---------- CONTEXT ----------
context = {
    "last_app": None,
    "last_action": None,
    "last_folder": None,
    "last_file": None
}

# ---------- CONSTANTS ----------
EXT_WORDS = {
    "pdf","doc","docx","ppt","pptx","xls","xlsx","txt",
    "png","jpg","jpeg","jfif","csv","json","mp3","mp4","mkv","py","zip","rar","7z"
}
IMAGE_EXT = {"png","jpg","jpeg","jfif"}

FOLDER_MAP = {
    "download": "downloads", "downloads": "downloads",
    "document": "documents", "documents": "documents",
    "desktop": "desktop",
    "picture": "pictures", "pictures": "pictures", "photos": "pictures", "images": "pictures",
    "music": "music", "songs": "music", "audio": "music",
    "video": "videos", "videos": "videos", "movies": "videos"
}
FOLDER_PATHS = {
    "downloads": os.path.expanduser("~\\Downloads"),
    "documents": os.path.expanduser("~\\Documents"),
    "desktop":   os.path.expanduser("~\\Desktop"),
    "pictures":  os.path.expanduser("~\\Pictures"),
    "music":     os.path.expanduser("~\\Music"),
    "videos":    os.path.expanduser("~\\Videos"),
}

# ---------- HELPERS: FILE MATCH ----------
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
        if score > 0:
            ranked.append((score, f))

    ranked.sort(key=lambda x: x[0], reverse=True)
    return [f for score, f in ranked]

def find_file_in_folder(file_name, folder_path):
    try:
        files = os.listdir(folder_path)
        matches = rank_matches(files, folder_path, file_name)
        if matches:
            return os.path.join(folder_path, matches[0])
    except Exception as e:
        print("Find file error:", e)
    return None

def search_known_folders_for_file(file_query: str):
    for folder_path in FOLDER_PATHS.values():
        try:
            files = os.listdir(folder_path)
        except Exception:
            continue
        matches = rank_matches(files, folder_path, file_query)
        if matches:
            return os.path.join(folder_path, matches[0])
    return None

# ---------- NOISE-REDUCED RECORD ----------
def record_with_noise_suppression(duration=6, samplerate=16000):
    try:
        print("🎚️ Calibrating background noise (1s)...")
        beep_ready()
        noise = sd.rec(int(1 * samplerate), samplerate=samplerate, channels=1, dtype="int16")
        sd.wait()

        print("🎤 Recording (speak now)...")
        audio = sd.rec(int(duration * samplerate), samplerate=samplerate, channels=1, dtype="int16")
        sd.wait()

        noise = noise.flatten()
        voice = audio.flatten()

        reduced = nr.reduce_noise(y=voice, y_noise=noise, sr=samplerate)
        return reduced.astype(np.int16), samplerate
    except Exception as e:
        print("Noise reduction failed, falling back:", e)
        return None, None

# ---------- GEMINI PARSER ----------
def parse_command_online(command: str):
    prompt = f"""
You are a command parser for a Windows voice assistant.
Turn the instruction into STRICT JSON (no backticks, no commentary).

Allowed actions:
- open_app → open an app or a system folder (params: app)
- open_file → open a file in the last opened folder or known folders (params: file)
- move_file → move a file (params: source, destination)
- delete_file → delete a file (params: file)
- shutdown → shut down the system
- restart → restart the system
- screenshot → take a screenshot (params: optional path)
- youtube_search → search on YouTube (params: query)
- youtube_control → control playback (params: command: play, pause, stop, next, previous, volume_up, volume_down)
- web_search → Google search (params: query)
- app_control → control windows (params: app, command: close, minimize, maximize, restore, focus)
- list_files → list all files in the last opened folder (no params)
- exit → close the assistant (no params)

Examples:
- "show me everything in downloads" -> {{"action":"open_app","params":{{"app":"downloads"}}}}
- "move budget sheet to desktop" -> {{"action":"move_file","params":{{"source":"budget sheet","destination":"desktop"}}}}
- "trash the picture selfie" -> {{"action":"delete_file","params":{{"file":"selfie"}}}}

Instruction: "{command}"

Output ONLY JSON.
"""
    try:
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(prompt)
        text = (response.text or "").strip()
        if text.startswith("```"):
            text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except Exception as e:
        print("Gemini parse failed:", e)
        return None

# ---------- FALLBACK PARSER (NL heuristics) ----------
def fallback_parser(command: str):
    c = command.lower().strip()

    # Exit
    if any(kw in c for kw in ["stop assistant", "exit assistant", "quit assistant", "terminate assistant", "close assistant"]):
        return {"action": "exit", "params": {}}

    # Open common folders
    for k in ["downloads","documents","desktop","pictures","photos","images","music","videos","movies"]:
        if f"open {k}" in c or f"show {k}" in c or f"go to {k}" in c or f"take me to {k}" in c:
            return {"action": "open_app", "params": {"app": k}}

    # List files
    if any(kw in c for kw in ["list files", "show files", "what's in this folder", "show me everything here", "show everything here"]):
        return {"action": "list_files", "params": {}}

    # Move file: "move <x> to <dest>"
    m = re.search(r"(?:move|send|put)\s+(.+?)\s+(?:to|into|onto)\s+(.+)", c)
    if m:
        src = m.group(1).strip()
        dst = m.group(2).strip()
        return {"action": "move_file", "params": {"source": src, "destination": dst}}

    # Delete file
    m = re.search(r"(?:delete|remove|trash|bin)\s+(.+)", c)
    if m:
        return {"action": "delete_file", "params": {"file": m.group(1).strip()}}

    # Open file by name inside current or known folders
    m = re.search(r"(?:open|launch)\s+(.+)", c)
    if m and "youtube" not in c and "google" not in c:
        return {"action": "open_file", "params": {"file": m.group(1).strip()}}

    # YouTube search
    m = re.search(r"(?:youtube\s+search|search\s+youtube\s+for|play\s+on\s+youtube)\s+(.+)", c)
    if m:
        return {"action": "youtube_search", "params": {"query": m.group(1).strip()}}

    # YouTube control
    if "pause" in c or "resume" in c or "play" in c or "next" in c or "previous" in c or "volume" in c or "stop" in c:
        cmd = "pause" if "pause" in c else "play" if "play" in c or "resume" in c else "next" if "next" in c else "previous" if "previous" in c or "prev" in c else "stop" if "stop" in c else "volume_up" if "up" in c else "volume_down"
        return {"action": "youtube_control", "params": {"command": cmd}}

    # Web search
    m = re.search(r"(?:google\s+|search\s+)(.+)", c)
    if m:
        return {"action": "web_search", "params": {"query": m.group(1).strip()}}

    # System
    if "shut down" in c or "power off" in c or "shutdown" in c:
        return {"action": "shutdown", "params": {}}
    if "restart" in c or "reboot" in c:
        return {"action": "restart", "params": {}}
    if "screenshot" in c or "take a screenshot" in c:
        return {"action": "screenshot", "params": {}}

    # App control
    m = re.search(r"(close|minimize|maximize|restore|focus)\s+(.+)", c)
    if m:
        return {"action": "app_control", "params": {"command": m.group(1), "app": m.group(2)}}

    return {"action": "unknown", "params": {}}

# ---------- EXECUTION ----------
def execute(action, params):
    global context

    def map_folder(name: str):
        if not name: return None
        key = FOLDER_MAP.get(name.lower(), name.lower())
        return FOLDER_PATHS.get(key, None), key

    try:
        # OPEN APP / FOLDER (also tries executable name)
        if action == "open_app":
            app = params.get("app", "").lower()
            folder_path, mapped = map_folder(app)
            if folder_path:
                os.startfile(folder_path)
                speak(f"Opening {mapped}")
                context["last_app"] = mapped
                context["last_action"] = "open_app"
                context["last_folder"] = folder_path
                return

            # search file in known folders (open file by name)
            file_hit = search_known_folders_for_file(app)
            if file_hit:
                os.startfile(file_hit)
                speak(f"Opening {os.path.basename(file_hit)}")
                context["last_file"] = file_hit
                context["last_folder"] = os.path.dirname(file_hit)
                return

            # try executable
            try:
                subprocess.Popen(["powershell", "-Command", f"Start-Process '{app}'"])
                speak(f"Opening {app}")
                context["last_app"] = app
                context["last_action"] = "open_app"
            except Exception:
                speak(f"Could not find or open {app}")
            return

        # LIST FILES
        elif action == "list_files":
            folder_path = context.get("last_folder")
            if not folder_path:
                speak("No folder is currently open. Please open a folder first.")
                return
            try:
                files = os.listdir(folder_path)
                if not files:
                    speak("This folder is empty.")
                    return
                print("\n📂 Files in", folder_path, ":")
                for f in files:
                    print("-", f)
                preview = ", ".join(files[:8])
                speak(f"I found {len(files)} items. For example: {preview}")
            except Exception as e:
                speak(f"Failed to list files: {e}")
            return

        # OPEN FILE (in last folder or known folders)
        elif action == "open_file":
            raw_query = params.get("file", "").strip()
            # Prefer last folder
            folder_path = context.get("last_folder")
            file_path = None
            if folder_path:
                file_path = find_file_in_folder(raw_query, folder_path)
            if not file_path:
                file_path = search_known_folders_for_file(raw_query)

            if file_path:
                os.startfile(file_path)
                speak(f"Opening {os.path.basename(file_path)}")
                context["last_file"] = file_path
                context["last_folder"] = os.path.dirname(file_path)
            else:
                speak(f"No matching file found for '{raw_query}'.")
            return

        # MOVE FILE
        elif action == "move_file":
            src = (params.get("source") or "").strip()
            dst = (params.get("destination") or "").strip()
            if not src and context.get("last_file"):
                src = context["last_file"]
            if not src or not dst:
                speak("Please specify both source and destination.")
                return

            # resolve src
            if os.path.isabs(src) and os.path.exists(src):
                src_full = src
            else:
                # try last folder first, then known folders
                src_full = None
                if context.get("last_folder"):
                    src_full = find_file_in_folder(src, context["last_folder"])
                if not src_full:
                    src_full = search_known_folders_for_file(src)
            if not src_full:
                speak(f"File not found: {src}")
                return

            dst_lower = dst.lower().strip()
            if dst_lower in ["recycle bin", "trash", "bin"]:
                send2trash(src_full)
                speak(f"{os.path.basename(src_full)} sent to Recycle Bin")
                if context.get("last_file") == src_full:
                    context["last_file"] = None
                return

            dst_folder_path, mapped = map_folder(dst_lower)
            if dst_folder_path:
                dst_full = os.path.join(dst_folder_path, os.path.basename(src_full))
            else:
                # if path provided, use it (create folder if needed)
                if not os.path.isabs(dst_lower):
                    # treat as relative to Desktop
                    base = FOLDER_PATHS.get("desktop")
                    dst_lower = os.path.join(base, dst_lower)
                os.makedirs(dst_lower, exist_ok=True)
                dst_full = os.path.join(dst_lower, os.path.basename(src_full))

            try:
                os.replace(src_full, dst_full)
                speak(f"Moved to {dst_full}")
                context["last_file"] = dst_full
                context["last_folder"] = os.path.dirname(dst_full)
            except Exception as e:
                speak(f"Failed to move file: {e}")
            return

        # DELETE FILE
        elif action == "delete_file":
            raw_query = (params.get("file") or "").strip()
            if not raw_query and context.get("last_file"):
                raw_query = context["last_file"]

            if os.path.isabs(raw_query) and os.path.exists(raw_query):
                target = raw_query
            else:
                target = None
                if context.get("last_folder"):
                    target = find_file_in_folder(raw_query, context["last_folder"])
                if not target:
                    target = search_known_folders_for_file(raw_query)

            if not target:
                speak(f"No matching file found for '{raw_query}'.")
                return

            try:
                send2trash(target)
                speak(f"{os.path.basename(target)} sent to Recycle Bin")
                if context.get("last_file") == target:
                    context["last_file"] = None
            except Exception as e:
                speak(f"Failed to delete file: {e}")
            return

        # SYSTEM
        elif action == "shutdown":
            speak("Shutting down now.")
            os.system("shutdown /s /t 1")
            return

        elif action == "restart":
            speak("Restarting now.")
            os.system("shutdown /r /t 1")
            return

        # SCREENSHOT
        elif action == "screenshot":
            path = params.get("path") or "screenshot.png"
            try:
                pyautogui.screenshot(path)
                speak(f"Screenshot saved as {path}")
            except Exception as e:
                speak(f"Failed to take screenshot: {e}")
            return

        # YOUTUBE SEARCH / CONTROL
        elif action == "youtube_search":
            q = params.get("query", "").strip()
            webbrowser.open(f"https://www.youtube.com/results?search_query={q}")
            speak(f"Searching YouTube for {q}")
            context["last_action"] = "youtube_search"
            return

        elif action == "youtube_control":
            cmd = (params.get("command") or "").lower()
            import pyautogui as _pg
            if cmd in ["pause", "play", "stop"]:
                _pg.press("k")
            elif cmd == "next":
                _pg.hotkey("shift", "n")
            elif cmd == "previous":
                _pg.hotkey("shift", "p")
            elif cmd == "volume_up":
                _pg.press("volumeup")
            elif cmd == "volume_down":
                _pg.press("volumedown")
            speak(f"YouTube {cmd}")
            return

        # WEB SEARCH
        elif action == "web_search":
            q = params.get("query", "").strip()
            webbrowser.open(f"https://www.google.com/search?q={q}")
            speak(f"Searching Google for {q}")
            context["last_action"] = "web_search"
            return

        # APP CONTROL
        elif action == "app_control":
            app_name = (params.get("app") or context.get("last_app") or "").lower()
            cmd = (params.get("command") or "").lower()
            if not app_name:
                speak("Specify the app to control.")
                return
            try:
                windows = gw.getWindowsWithTitle(app_name)
                if not windows:
                    speak(f"No open windows found for {app_name}")
                    return
                window = windows[0]
                if cmd == "close": window.close()
                elif cmd == "minimize": window.minimize()
                elif cmd == "maximize": window.maximize()
                elif cmd == "restore": window.restore()
                elif cmd == "focus": window.activate()
                else:
                    speak(f"Unknown command {cmd} for {app_name}")
                    return
                speak(f"{cmd.capitalize()} executed on {app_name}")
                context["last_app"] = app_name
                context["last_action"] = "app_control"
            except Exception as e:
                speak(f"Failed to control {app_name}: {e}")
            return

        # EXIT
        elif action == "exit":
            speak("Goodbye! Shutting down assistant.")
            beep_ok()
            sys.exit(0)

        # UNKNOWN
        else:
            beep_err()
            speak("Sorry, I don’t understand that.")
            return

    except Exception as e:
        beep_err()
        speak(f"Execution error: {e}")

# ---------- LISTEN LOOP ----------
def listen_and_execute_loop():
    r = sr.Recognizer()

    while True:
        try:
            # Prefer our noise-reduced recorder
            audio_arr, sr_hz = record_with_noise_suppression(duration=6, samplerate=16000)
            if audio_arr is not None:
                audio_clean = sr.AudioData(audio_arr.tobytes(), sample_rate=sr_hz, sample_width=2)
                beep_ok()
                command = r.recognize_google(audio_clean)
            else:
                # Fallback to SpeechRecognition mic if noise pipeline failed
                with sr.Microphone() as source:
                    print("🎤 Listening...")
                    r.adjust_for_ambient_noise(source, duration=0.6)
                    audio = r.listen(source, phrase_time_limit=6)
                beep_ok()
                command = r.recognize_google(audio)

            print("You said:", command)

            result = parse_command_online(command)
            if not result:
                result = fallback_parser(command)

            print("Parsed:", result)
            execute(result.get("action"), result.get("params"))

        except sr.UnknownValueError:
            beep_err()
            speak("Sorry, I didn’t catch that.")
        except Exception as e:
            beep_err()
            print("Error:", e)
            speak("Something went wrong.")

# ---------- MAIN ----------
if __name__ == "__main__":
    speak("Assistant is running. Speak naturally after the beep. Say 'stop assistant' to exit.")
    listen_and_execute_loop()
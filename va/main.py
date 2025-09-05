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
import winreg
import time
from pathlib import Path

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
    "last_file": None,
    "folder_contents": {}  # Cache folder contents
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

# ---------- DYNAMIC APPLICATION FINDER ----------
_app_cache = None
_cache_timestamp = 0

def find_installed_apps():
    """Scan Windows registry to find all installed applications"""
    apps = {}
    
    # Registry paths where installed programs are listed
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    ]
    
    for hkey, reg_path in reg_paths:
        try:
            reg_key = winreg.OpenKey(hkey, reg_path)
            for i in range(winreg.QueryInfoKey(reg_key)[0]):
                try:
                    subkey_name = winreg.EnumKey(reg_key, i)
                    subkey = winreg.OpenKey(reg_key, subkey_name)
                    try:
                        # Get application display name
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        
                        # Try to get executable path
                        exe_path = None
                        try:
                            # Method 1: Direct executable path
                            exe_path = winreg.QueryValueEx(subkey, "DisplayIcon")[0]
                            if not exe_path.endswith('.exe'):
                                exe_path = None
                        except FileNotFoundError:
                            pass
                        
                        if not exe_path:
                            try:
                                # Method 2: Install location + search for exe
                                install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                if install_location and os.path.exists(install_location):
                                    # Find main executable in install directory
                                    exe_files = list(Path(install_location).glob("**/*.exe"))
                                    if exe_files:
                                        # Prefer exe with similar name to app
                                        app_words = display_name.lower().split()
                                        best_exe = None
                                        best_score = 0
                                        
                                        for exe in exe_files:
                                            exe_name = exe.stem.lower()
                                            score = sum(1 for word in app_words if word in exe_name)
                                            if score > best_score:
                                                best_score = score
                                                best_exe = str(exe)
                                        
                                        if best_exe:
                                            exe_path = best_exe
                                        else:
                                            exe_path = str(exe_files[0])  # Fallback to first exe
                            except (FileNotFoundError, IndexError):
                                pass
                        
                        if exe_path and os.path.exists(exe_path):
                            # Store with normalized name for easy matching
                            normalized_name = normalize_app_name(display_name)
                            apps[normalized_name] = {
                                'display_name': display_name,
                                'exe_path': exe_path,
                                'keywords': generate_keywords(display_name)
                            }
                    
                    except FileNotFoundError:
                        pass
                    finally:
                        winreg.CloseKey(subkey)
                        
                except Exception:
                    continue
            winreg.CloseKey(reg_key)
        except Exception as e:
            print(f"Error reading registry {reg_path}: {e}")
            continue
    
    return apps

def normalize_app_name(name):
    """Normalize application name for consistent matching"""
    # Remove common suffixes and prefixes
    name = re.sub(r'\s*\(.*?\)', '', name)  # Remove parentheses content
    name = re.sub(r'\s*(64-bit|32-bit|x64|x86)', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*version\s+[\d.]+', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*v[\d.]+', '', name, flags=re.IGNORECASE)
    return name.strip().lower()

def generate_keywords(app_name):
    """Generate search keywords for an application"""
    keywords = set()
    
    # Original name variations
    keywords.add(app_name.lower())
    keywords.add(normalize_app_name(app_name))
    
    # Split words
    words = re.findall(r'\w+', app_name.lower())
    keywords.update(words)
    
    # Common abbreviations
    if len(words) > 1:
        # First letters of each word
        abbreviation = ''.join(word[0] for word in words)
        keywords.add(abbreviation)
        
        # First word only
        keywords.add(words[0])
    
    # Remove short/common words
    keywords = {k for k in keywords if len(k) > 2 and k not in ['the', 'and', 'for', 'app']}
    
    return list(keywords)

def find_best_app_match(query, installed_apps):
    """Find best matching application using fuzzy matching"""
    query_lower = query.lower()
    query_words = re.findall(r'\w+', query_lower)
    
    best_matches = []
    
    for app_key, app_info in installed_apps.items():
        score = 0
        
        # Exact name match (highest score)
        if query_lower == app_key:
            score = 100
        
        # Check if query is in display name
        elif query_lower in app_info['display_name'].lower():
            score = 80
        
        # Check keywords
        else:
            for keyword in app_info['keywords']:
                if keyword in query_lower:
                    score += 20
                elif query_lower in keyword:
                    score += 15
            
            # Word matching
            app_words = re.findall(r'\w+', app_info['display_name'].lower())
            word_matches = sum(1 for q_word in query_words for a_word in app_words if q_word in a_word or a_word in q_word)
            score += word_matches * 10
        
        if score > 0:
            best_matches.append((score, app_info))
    
    # Sort by score and return best match
    best_matches.sort(key=lambda x: x[0], reverse=True)
    return best_matches[0][1] if best_matches else None

def get_installed_apps(force_refresh=False):
    """Get installed apps with caching to improve performance"""
    global _app_cache, _cache_timestamp
    
    current_time = time.time()
    
    # Refresh cache if older than 1 hour or forced
    if force_refresh or _app_cache is None or (current_time - _cache_timestamp) > 3600:
        print("🔍 Scanning installed applications...")
        _app_cache = find_installed_apps()
        _cache_timestamp = current_time
        print(f"Found {len(_app_cache)} applications")
    
    return _app_cache

def launch_app_dynamically(app_name):
    """Launch application using dynamic discovery"""
    try:
        # Get installed apps
        installed_apps = get_installed_apps()
        
        # Find best match
        app_info = find_best_app_match(app_name, installed_apps)
        
        if app_info:
            print(f"Launching: {app_info['display_name']} ({app_info['exe_path']})")
            
            # Try to launch the executable
            try:
                subprocess.Popen([app_info['exe_path']], shell=True)
                return True, app_info['display_name']
            except Exception as e:
                print(f"Failed to launch {app_info['exe_path']}: {e}")
                
                # Fallback: try launching by display name
                try:
                    subprocess.Popen(app_info['display_name'], shell=True)
                    return True, app_info['display_name']
                except Exception as e2:
                    print(f"Fallback also failed: {e2}")
                    return False, f"Could not launch {app_info['display_name']}"
        
        else:
            return False, f"Application '{app_name}' not found"
            
    except Exception as e:
        print(f"Error in launch_app_dynamically: {e}")
        return False, f"Error launching {app_name}"

def launch_app_with_powershell(app_name):
    """Alternative method using PowerShell Get-StartApps"""
    try:
        # PowerShell command to find and launch app
        ps_command = f"""
        $app = Get-StartApps | Where-Object {{$_.Name -like '*{app_name}*'}} | Select-Object -First 1
        if ($app) {{
            Start-Process -FilePath $app.AppID
            Write-Output "Launched: $($app.Name)"
        }} else {{
            Write-Output "Not found: {app_name}"
        }}
        """
        
        result = subprocess.run(
            ['powershell', '-Command', ps_command], 
            capture_output=True, 
            text=True, 
            timeout=10
        )
        
        if result.returncode == 0 and "Launched:" in result.stdout:
            launched_app = result.stdout.split("Launched: ")[1].strip()
            return True, launched_app
        else:
            return False, f"Could not find '{app_name}' in Start Menu"
            
    except Exception as e:
        print(f"PowerShell launch failed: {e}")
        return False, f"PowerShell error: {str(e)[:50]}"

def execute_open_app_dynamic(app_name):
    """Updated open_app logic without hardcoding"""
    
    # Method 1: Try dynamic registry-based launch
    success, message = launch_app_dynamically(app_name)
    if success:
        speak(f"Opening {message}")
        return True
    
    # Method 2: Try PowerShell method
    success, message = launch_app_with_powershell(app_name)
    if success:
        speak(f"Opening {message}")
        return True
    
    # Method 3: Try simple subprocess (for system commands)
    try:
        subprocess.Popen(app_name, shell=True)
        speak(f"Opening {app_name}")
        return True
    except Exception:
        pass
    
    # Method 4: Try with common executable extensions
    for ext in ['.exe', '.bat', '.cmd']:
        try:
            subprocess.Popen(app_name + ext, shell=True)
            speak(f"Opening {app_name}")
            return True
        except Exception:
            continue
    
    # All methods failed
    speak(f"Could not find application '{app_name}'. Please check if it's installed.")
    return False

# ---------- OPTIMIZED HELPERS ----------
def normalize_text(s: str) -> str:
    """Optimized text normalization"""
    s = s.lower()
    s = os.path.splitext(s)[0]
    s = re.sub(r'[\W_]+', '', s, flags=re.UNICODE)
    return s

def extract_query_and_ext(file_query: str):
    """Extract file extension and core query"""
    tokens = re.findall(r'\w+', file_query.lower())
    ext = None
    for t in tokens[::-1]:
        if t in EXT_WORDS:
            ext = t
            break
    core_tokens = [t for t in tokens if t not in EXT_WORDS]
    core_query = ''.join(core_tokens)
    return core_query, ext

def get_folder_contents(folder_path):
    """Get and cache folder contents"""
    try:
        if folder_path in context["folder_contents"]:
            return context["folder_contents"][folder_path]
        
        items = os.listdir(folder_path)
        files = []
        folders = []
        
        for item in items:
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                files.append(item)
            elif os.path.isdir(item_path):
                folders.append(item)
        
        contents = {"files": files, "folders": folders}
        context["folder_contents"][folder_path] = contents
        return contents
    except Exception as e:
        print(f"Error getting folder contents: {e}")
        return {"files": [], "folders": []}

def rank_matches(items, raw_query, item_type="both"):
    """Enhanced matching with better scoring"""
    q_norm_core, ext_hint = extract_query_and_ext(raw_query)
    query_tokens = re.findall(r'\w+', raw_query.lower())

    ranked = []
    for item in items:
        # Skip if looking for specific type
        if item_type == "files" and item not in get_folder_contents(context.get("last_folder", "")).get("files", []):
            continue
        elif item_type == "folders" and item not in get_folder_contents(context.get("last_folder", "")).get("folders", []):
            continue

        # Extension filtering
        if ext_hint and not item.lower().endswith("." + ext_hint):
            continue
        elif ("image" in raw_query.lower() or "photo" in raw_query.lower()) and not item.lower().endswith(tuple(IMAGE_EXT)):
            continue

        # Scoring
        item_tokens = re.findall(r'\w+', item.lower())
        token_hits = sum(1 for qt in query_tokens if qt in item_tokens)
        ratio = SequenceMatcher(None, q_norm_core, normalize_text(item)).ratio()
        
        # Exact name match gets highest score
        if item.lower() == raw_query.lower():
            score = 10
        elif normalize_text(item) == q_norm_core:
            score = 8 + token_hits
        else:
            score = token_hits * 2 + ratio

        if score > 0:
            ranked.append((score, item))

    ranked.sort(key=lambda x: x[0], reverse=True)
    return [item for score, item in ranked]

def smart_find_item(query, folder_path=None, item_type="both"):
    """Smart item finder that lists contents first, then matches"""
    if not folder_path:
        folder_path = context.get("last_folder")
    
    if not folder_path or not os.path.isdir(folder_path):
        return None, []
    
    # Get folder contents
    contents = get_folder_contents(folder_path)
    
    if item_type == "files":
        items = contents["files"]
    elif item_type == "folders":
        items = contents["folders"]
    else:
        items = contents["files"] + contents["folders"]
    
    if not items:
        return None, []
    
    # Find matches
    matches = rank_matches(items, query)
    
    return matches[0] if matches else None, matches

def search_all_folders_for_item(query, item_type="both"):
    """Search all known folders for an item"""
    for folder_name, folder_path in FOLDER_PATHS.items():
        try:
            match, matches = smart_find_item(query, folder_path, item_type)
            if match:
                return os.path.join(folder_path, match), folder_path
        except Exception:
            continue
    return None, None

def clear_folder_cache():
    """Clear cached folder contents"""
    context["folder_contents"].clear()

# ---------- OPTIMIZED NOISE REDUCTION ----------
def record_with_noise_suppression(duration=5, samplerate=16000):
    """Optimized noise reduction with shorter duration"""
    try:
        print("🎚️ Calibrating noise (0.5s)...")
        beep_ready()
        noise = sd.rec(int(0.5 * samplerate), samplerate=samplerate, channels=1, dtype="int16")
        sd.wait()

        print("🎤 Recording...")
        audio = sd.rec(int(duration * samplerate), samplerate=samplerate, channels=1, dtype="int16")
        sd.wait()

        noise = noise.flatten()
        voice = audio.flatten()

        reduced = nr.reduce_noise(y=voice, y_noise=noise, sr=samplerate)
        return reduced.astype(np.int16), samplerate
    except Exception as e:
        print("Noise reduction failed:", e)
        return None, None

# ---------- OPTIMIZED GEMINI PARSER ----------
def parse_command_online(command: str):
    """Optimized Gemini parser with better prompts for friendly interaction."""
    prompt = f"""

Actions:
- open_app: Open an application or a system folder (e.g., ", open Chrome for me", "Chief, show me my desktop")
    (params: app [string, required])
- open_file: Open a specific file (e.g., ", could you open my report?", ", find that picture")
    (params: file [string, required], type_filter [string, optional - e.g., 'image', 'pdf', 'document'])
- open_folder: Open a user-created folder (e.g., ", open my project folder", "Commander, go to the recipes directory")
    (params: folder [string, required])
- open_subfolder: Open a subfolder explicitly mentioned in relation to a parent (e.g., ", open reports inside documents")
    (params: subfolder [string, required], parent [string, optional])
- go_back: Navigate to the parent folder (e.g., ", go back", ", take me up one level")
- move_file: Move a file from one place to another (e.g., "Chief, move this document to downloads")
    (params: source [string, required], destination [string, required])
- delete_file: Delete a file (e.g., ", delete that old photo", ", trash this file")
    (params: file [string, required])
- app_control: Control an active window or application (e.g., ", close Chrome", ", minimize this window")
    (params: command [string, required - close, minimize, maximize, restore, focus], app [string, optional - if omitted, assume last active app/folder])
- shutdown/restart/screenshot: System-level actions (e.g., ", shut down my computer", ", take a screenshot")
- youtube_search: Search for a video on YouTube (e.g., ", search YouTube for cat videos", ", play some music on YouTube")
    (params: query [string, required])
- youtube_control: Control YouTube playback (e.g., ", pause YouTube", "Chief, next song")
    (params: command [string, required - pause, play, next, previous, volume_up, volume_down])
- web_search: Perform a web search on Google (e.g., ", google how to bake a cake", ", search for weather forecast")
    (params: query [string, required])
- list_files: List contents of a folder, potentially with a filter (e.g., ", show me files in documents", ", list my pictures")
    (params: folder [string, optional], type_filter [string, optional - e.g., 'pdf', 'image', 'document', 'video', 'audio'])
- exit: Stop the assistant (e.g., ", exit assistant", ", stop listening")

System Folders for 'open_app' or 'list_files': downloads, documents, desktop, pictures, music, videos
Common File Type Filters: image, video, audio, document, pdf, docx, xlsx, pptx, txt, png, jpg, mp3, mp4

Examples of 's friendly/casual commands and their JSON parsing:
"Hey , could you open my downloads folder please, ?" → {{"action":"open_app","params":{{"app":"downloads"}}}}
", can you please find my latest report for me?" → {{"action":"open_file","params":{{"file":"latest report","type_filter":"document"}}}}
", I'd love to see that cat video I watched yesterday!" → {{"action":"youtube_search","params":{{"query":"cat video I watched yesterday"}}}}
"Commander, could you quickly take me to my 'Work Projects' folder?" → {{"action":"open_folder","params":{{"folder":"Work Projects"}}}}
", can you list all the images in my pictures for me, please?" → {{"action":"list_files","params":{{"folder":"pictures","type_filter":"image"}}}}
", I need to close this browser window, if you don't mind." → {{"action":"app_control","params":{{"command":"close","app":"browser"}}}}
"Chief, could you move this file straight to the trash?" → {{"action":"delete_file","params":{{"file":"this file"}}}}
"What exactly is in the 'recipes' folder, ?" → {{"action":"list_files","params":{{"folder":"recipes"}}}}
"Hey , could you open 'Important Notes.docx' please?" → {{"action":"open_file","params":{{"file":"Important Notes.docx"}}}}
"Just go back one step for me, thanks !" → {{"action":"go_back","params":{{}}}}
"Would you mind minimizing this window, ?" → {{"action":"app_control","params":{{"command":"minimize"}}}}
", could you search Google for 'best pizza near me' for me?" → {{"action":"web_search","params":{{"query":"best pizza near me"}}}}
"Alright , I'm all done for now. Thanks!" → {{"action":"exit","params":{{}}}}
", enaku andha file-ah open pannu." (Tamil example for "open that file") -> {{"action":"open_file","params":{{"file":"andha file"}}}}
", intha folder la enna iruku?" (Tamil example for "what's in this folder") -> {{"action":"list_files","params":{{"folder":"intha folder"}}}}


Command: "{command}"
JSON only:
"""
    try:
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(prompt)
        text = (response.text or "").strip()
        text = re.sub(r'```(?:json)?', '', text).strip()
        return json.loads(text)
    except Exception as e:
        print(f"Gemini parse failed: {e}")
        return {"action": "unknown", "params": {}} # Return 'unknown' on parsing failure # Return 'unknown' on parsing failure
    
# ---------- ENHANCED FALLBACK PARSER ----------
def fallback_parser(command: str):
    """Enhanced fallback parser with better patterns"""
    c = command.lower().strip()

    # Exit commands
    if any(kw in c for kw in ["stop assistant", "exit", "quit", "terminate", "close assistant"]):
        return {"action": "exit", "params": {}}

    # Navigation commands
    if any(kw in c for kw in ["go back", "previous folder", "back", "parent folder"]):
        return {"action": "go_back", "params": {}}

    # Subfolder patterns - enhanced
    patterns = [
        r"(?:open|show|go to)\s+(.+?)\s+(?:inside|in|within)\s+(.+)",
        r"(?:open|show)\s+(.+?)\s+folder(?:\s+in\s+(.+))?",
        r"(?:open|go to)\s+(.+?)\s+(?:directory|dir)(?:\s+in\s+(.+))?"
    ]
    
    for pattern in patterns:
        m = re.search(pattern, c)
        if m:
            subfolder = m.group(1).strip()
            parent = m.group(2).strip() if m.group(2) else None
            if parent:
                return {"action": "open_subfolder", "params": {"subfolder": subfolder, "parent": parent}}
            else:
                return {"action": "open_subfolder", "params": {"subfolder": subfolder}}

    # System folders (still use open_app for these)
    for k in ["downloads", "documents", "desktop", "pictures", "photos", "music", "videos"]:
        if any(phrase in c for phrase in [f"open {k}", f"show {k}", f"go to {k}"]):
            return {"action": "open_app", "params": {"app": k}}

    # *** NEW: Open specific folder action ***
    open_folder_match = re.search(r"(?:open|show|go to)\s+(.+?)\s+(?:folder|directory|dir)", c)
    if open_folder_match:
        folder_name = open_folder_match.group(1).strip()
        return {"action": "open_folder", "params": {"folder": folder_name}}

    # File operations - ENHANCED FOR LIST FILES WITH TYPE FILTER
    list_files_match = re.search(r"(?:list|show)(?: all)?(?: the)?\s+(.*?)\s+(?:files|documents|images|videos|audios|pdfs|docs|sheets)?(?:\s+in\s+(.+))?", c)
    if list_files_match:
        file_type_query = list_files_match.group(1).strip()
        folder_query = list_files_match.group(2)
        
        params = {}
        if folder_query:
            params["folder"] = folder_query.strip()
            
        # Try to infer type filter
        for ext_word in EXT_WORDS:
            if ext_word in file_type_query:
                params["type_filter"] = ext_word
                break
        
        if "image" in file_type_query or "photo" in file_type_query:
            params["type_filter"] = "image"
        elif "video" in file_type_query:
            params["type_filter"] = "video"
        elif "audio" in file_type_query or "song" in file_type_query or "music" in file_type_query:
            params["type_filter"] = "audio"
        elif "document" in file_type_query or "doc" in file_type_query:
             # Be careful not to override a more specific 'doc' or 'docx'
             if not params.get("type_filter") or params["type_filter"] not in ["doc", "docx"]:
                params["type_filter"] = "document" # Generic document
        elif "pdf" in file_type_query:
            params["type_filter"] = "pdf"
        elif "all" in file_type_query or "files" in file_type_query or not file_type_query: # No specific filter mentioned
            pass # No type_filter needed

        return {"action": "list_files", "params": params}

    # Move/delete operations
    move_match = re.search(r"(?:move|send)\s+(.+?)\s+(?:to|into)\s+(.+)", c)
    if move_match:
        return {"action": "move_file", "params": {"source": move_match.group(1), "destination": move_match.group(2)}}

    delete_match = re.search(r"(?:delete|remove|trash)\s+(.+)", c)
    if delete_match:
        return {"action": "delete_file", "params": {"file": delete_match.group(1)}}

    # Open file (generic - this should catch "open myfile.txt" not "open folder")
    open_match = re.search(r"(?:open|launch|start)\s+(.+)", c)
    if open_match and "youtube" not in c and "google" not in c:
        # Check if it *looks* like a folder, to prevent misdirection to open_file
        potential_name = open_match.group(1).strip()
        if not any(potential_name.endswith(f".{ext}") for ext in EXT_WORDS) and "folder" not in potential_name and "directory" not in potential_name:
            return {"action": "open_file", "params": {"file": potential_name}}


    # Web operations
    if "youtube" in c:
        search_match = re.search(r"(?:youtube|search youtube|play)\s+(.+)", c)
        if search_match:
            return {"action": "youtube_search", "params": {"query": search_match.group(1)}}

    if any(ctrl in c for ctrl in ["pause", "play", "next", "previous", "volume"]):
        cmd = "pause" if "pause" in c else "play" if "play" in c else "next" if "next" in c else "previous" if "prev" in c else "volume_up" if "up" in c else "volume_down"
        return {"action": "youtube_control", "params": {"command": cmd}}

    google_match = re.search(r"(?:google|search)\s+(.+)", c)
    if google_match:
        return {"action": "web_search", "params": {"query": google_match.group(1)}}

    # System operations
    if any(kw in c for kw in ["shut down", "shutdown", "power off"]):
        return {"action": "shutdown", "params": {}}
    if any(kw in c for kw in ["restart", "reboot"]):
        return {"action": "restart", "params": {}}
    if "screenshot" in c:
        return {"action": "screenshot", "params": {}}

    return {"action": "unknown", "params": {}}
# ---------- ENHANCED EXECUTION ----------
def execute(action, params):
    """Enhanced execution with better error handling and persistence"""
    global context

    def map_folder(name: str):
        if not name: return None, None
        key = FOLDER_MAP.get(name.lower(), name.lower())
        return FOLDER_PATHS.get(key), key

    try:
        # OPEN APP/FOLDER - DYNAMIC
        if action == "open_app":
            app = params.get("app", "").strip()
            
            # First, try system folders
            folder_path, mapped = map_folder(app.lower())
            if folder_path and os.path.isdir(folder_path):
                os.startfile(folder_path)
                speak(f"Opening {mapped}")
                context.update({
                    "last_app": mapped,
                    "last_action": "open_app", 
                    "last_folder": folder_path
                })
                clear_folder_cache()
                return

            # Use dynamic app launcher
            if execute_open_app_dynamic(app):
                context.update({"last_app": app, "last_action": "open_app"})
            
            return
        
        elif action == "open_folder":
            query = params.get("folder", "").strip()
            
            folder_path = None
            current_folder = context.get("last_folder")
            
            # 1. Try system folders first
            mapped_path, mapped_name = map_folder(query)
            if mapped_path and os.path.isdir(mapped_path):
                folder_path = mapped_path
            
            if not folder_path:
                # 2. Try current folder
                if current_folder:
                    match, _ = smart_find_item(query, current_folder, "folders")
                    if match:
                        folder_path = os.path.join(current_folder, match)
                
                # 3. Search all known folders
                if not folder_path:
                    for sys_folder_name, sys_folder_path in FOLDER_PATHS.items():
                        if sys_folder_path != current_folder: # Avoid re-searching current
                            match, _ = smart_find_item(query, sys_folder_path, "folders")
                            if match:
                                folder_path = os.path.join(sys_folder_path, match)
                                break # Found it, so stop searching

            if folder_path and os.path.isdir(folder_path):
                os.startfile(folder_path)
                speak(f"Opening folder {os.path.basename(folder_path)}")
                context.update({
                    "last_folder": folder_path,
                    "last_action": "open_folder"
                })
                clear_folder_cache()
            else:
                speak(f"Folder '{query}' not found.")
            return

        # ENHANCED OPEN SUBFOLDER
        elif action == "open_subfolder":
            subfolder = params.get("subfolder", "").strip()
            parent_hint = params.get("parent", "").strip()
            
            # Determine parent folder
            parent_folder = None
            if parent_hint:
                folder_path, mapped = map_folder(parent_hint)
                if folder_path and os.path.isdir(folder_path):
                    parent_folder = folder_path
            
            # Use last folder if no parent specified
            if not parent_folder:
                parent_folder = context.get("last_folder")
            
            if not parent_folder or not os.path.isdir(parent_folder):
                speak("Please specify a parent folder or open one first.")
                return

            # Smart find in parent folder
            match, all_matches = smart_find_item(subfolder, parent_folder, "folders")
            
            if match:
                target_path = os.path.join(parent_folder, match)
                os.startfile(target_path)
                speak(f"Opening {match}")
                context.update({
                    "last_folder": target_path,
                    "last_action": "open_subfolder"
                })
                clear_folder_cache()
            else:
                # Show available folders
                contents = get_folder_contents(parent_folder)
                available = contents["folders"][:5]
                if available:
                    speak(f"Subfolder '{subfolder}' not found. Available: {', '.join(available)}")
                else:
                    speak("No subfolders found in this location.")
            return

        # GO BACK - ENHANCED
        elif action == "go_back":
            current = context.get("last_folder")
            if not current or not os.path.isdir(current):
                speak("No current folder to go back from.")
                return
            
            parent = os.path.dirname(current)
            if parent and parent != current and os.path.isdir(parent):
                os.startfile(parent)
                speak(f"Going back to {os.path.basename(parent) or 'parent folder'}")
                context.update({
                    "last_folder": parent,
                    "last_action": "go_back"
                })
                clear_folder_cache()
            else:
                speak("Already at root level.")
            return

        # ENHANCED LIST FILES
        elif action == "list_files":
            # Check for a 'folder' parameter in the command's params
            requested_folder_name = params.get("folder") or params.get("parent")
            type_filter = params.get("type_filter", "").lower() # New: Get type filter
            
            folder_path = None
            if requested_folder_name:
                # Try to map the requested folder name to a known path
                mapped_path, _ = map_folder(requested_folder_name.lower())
                if mapped_path and os.path.isdir(mapped_path):
                    folder_path = mapped_path
                else:
                    speak(f"Could not find a system folder named '{requested_folder_name}'.")
                    return
            else:
                # If no specific folder is requested, use the last opened folder
                folder_path = context.get("last_folder")

            if not folder_path or not os.path.isdir(folder_path):
                speak("No folder is currently open or specified.")
                return
            
            contents = get_folder_contents(folder_path)
            folders = contents["folders"]
            all_files = contents["files"] # Renamed to avoid conflict

            filtered_files = []
            if type_filter:
                # Apply filter based on common extensions and keywords
                if type_filter == "image":
                    filtered_files = [f for f in all_files if os.path.splitext(f)[1][1:].lower() in IMAGE_EXT]
                elif type_filter == "pdf":
                    filtered_files = [f for f in all_files if f.lower().endswith(".pdf")]
                elif type_filter == "doc" or type_filter == "docx":
                    filtered_files = [f for f in all_files if f.lower().endswith((".doc", ".docx"))]
                elif type_filter == "xls" or type_filter == "xlsx":
                    filtered_files = [f for f in all_files if f.lower().endswith((".xls", ".xlsx"))]
                elif type_filter == "ppt" or type_filter == "pptx":
                    filtered_files = [f for f in all_files if f.lower().endswith((".ppt", ".pptx"))]
                elif type_filter == "txt":
                    filtered_files = [f for f in all_files if f.lower().endswith(".txt")]
                elif type_filter == "audio":
                    filtered_files = [f for f in all_files if os.path.splitext(f)[1][1:].lower() in ["mp3", "wav", "aac"]] # Add more audio types if needed
                elif type_filter == "video":
                    filtered_files = [f for f in all_files if os.path.splitext(f)[1][1:].lower() in ["mp4", "mkv", "avi"]] # Add more video types if needed
                elif type_filter in EXT_WORDS: # For other specific extensions
                    filtered_files = [f for f in all_files if f.lower().endswith(f".{type_filter}")]
                else: # Fallback if specific filter not matched
                    speak(f"Cannot filter by '{type_filter}'. Listing all files.")
                    filtered_files = all_files
            else:
                filtered_files = all_files

            files_to_display = filtered_files
            
            if not folders and not files_to_display:
                speak(f"The folder '{os.path.basename(folder_path)}' is empty or no items match your filter.")
                return
            
            print(f"\n📂 Contents of {os.path.basename(folder_path)}:")
            if folders:
                print("📁 Folders:")
                for f in folders[:10]:  # Limit display
                    print(f"  - {f}")
            if files_to_display:
                print(f"📄 {'Filtered ' if type_filter else ''}Files:")
                for f in files_to_display[:10]:  # Limit display
                    print(f"  - {f}")
            
            total_items = len(folders) + len(files_to_display)
            preview_items = (folders + files_to_display)[:3]
            
            if type_filter and files_to_display:
                speak(f"Found {len(files_to_display)} {type_filter} files in {os.path.basename(folder_path)}. Including: {', '.join(preview_items)}")
            elif type_filter and not files_to_display:
                speak(f"No {type_filter} files found in {os.path.basename(folder_path)}.")
            elif not type_filter and total_items > 0:
                speak(f"Found {total_items} items in {os.path.basename(folder_path)}. Including: {', '.join(preview_items)}")
            else:
                speak(f"The folder '{os.path.basename(folder_path)}' is empty.")
            
            # Update last_folder if a new one was explicitly listed
            context["last_folder"] = folder_path
            context["last_action"] = "list_files"
            return

        # ENHANCED OPEN FILE
        elif action == "open_file":
            query = params.get("file", "").strip()
            
            file_path = None
            current_folder = context.get("last_folder")
            
            # 1. Try current folder first
            possible_matches = []
            if current_folder:
                _, matches_in_current = smart_find_item(query, current_folder, "files")
                possible_matches.extend([os.path.join(current_folder, m) for m in matches_in_current])
            
            # 2. Search all folders if not enough matches or no current folder
            if not possible_matches or len(possible_matches) < 2: # Get more options if only one or none
                for folder_name, folder_path in FOLDER_PATHS.items():
                    if folder_path != current_folder: # Avoid re-searching current folder
                        _, matches_in_other = smart_find_item(query, folder_path, "files")
                        possible_matches.extend([os.path.join(folder_path, m) for m in matches_in_other])
            
            # Remove duplicates while preserving order for better ranking
            seen = set()
            unique_matches = []
            for item in possible_matches:
                if item not in seen:
                    unique_matches.append(item)
                    seen.add(item)
            
            # If no matches, inform the user
            if not unique_matches:
                speak(f"File '{query}' not found.")
                return
            
            # If multiple matches, ask user to choose
            if len(unique_matches) > 1:
                speak("I found multiple files. Please choose one by saying its number:")
                for i, match_path in enumerate(unique_matches[:5]): # Limit options to 5
                    speak(f"Option {i+1}: {os.path.basename(match_path)} in {os.path.basename(os.path.dirname(match_path))}")
                
                try:
                    r = sr.Recognizer()
                    with sr.Microphone() as source:
                        print("🎤 Listening for choice...")
                        r.adjust_for_ambient_noise(source, duration=0.3)
                        audio = r.listen(source, phrase_time_limit=3)
                    choice_str = r.recognize_google(audio).lower()
                    
                    choice_num = -1
                    if "one" in choice_str or "1" in choice_str: choice_num = 1
                    elif "two" in choice_str or "2" in choice_str: choice_num = 2
                    elif "three" in choice_str or "3" in choice_str: choice_num = 3
                    elif "four" in choice_str or "4" in choice_str: choice_num = 4
                    elif "five" in choice_str or "5" in choice_str: choice_num = 5

                    if 1 <= choice_num <= len(unique_matches[:5]):
                        file_path = unique_matches[choice_num - 1]
                    else:
                        speak("Invalid choice. Opening the first match.")
                        file_path = unique_matches[0] # Default to first
                except sr.UnknownValueError:
                    speak("Did not catch your choice. Opening the first match.")
                    file_path = unique_matches[0] # Default to first
                except Exception as e:
                    speak(f"Error processing choice: {e}. Opening the first match.")
                    file_path = unique_matches[0] # Default to first
            else:
                # Only one unique match
                file_path = unique_matches[0]

            if file_path and os.path.isfile(file_path):
                os.startfile(file_path)
                speak(f"Opening {os.path.basename(file_path)}")
                context.update({
                    "last_file": file_path,
                    "last_folder": os.path.dirname(file_path),
                    "last_action": "open_file"
                })
            else:
                speak(f"Could not open '{os.path.basename(file_path) if file_path else query}'.")
            return

        # MOVE FILE - ENHANCED
        elif action == "move_file":
            src = params.get("source", "").strip()
            dst = params.get("destination", "").strip()
            
            # Use last file if no source specified
            if not src:
                src = context.get("last_file", "")
            
            if not src or not dst:
                speak("Please specify both source and destination.")
                return
            
            # Find source file
            src_path = None
            if os.path.isabs(src) and os.path.exists(src):
                src_path = src
            else:
                current_folder = context.get("last_folder")
                if current_folder:
                    match, _ = smart_find_item(src, current_folder, "files")
                    if match:
                        src_path = os.path.join(current_folder, match)
                
                if not src_path:
                    src_path, _ = search_all_folders_for_item(src, "files")
            
            if not src_path:
                speak(f"Source file '{src}' not found.")
                return
            
            # Handle destination
            if dst.lower() in ["recycle bin", "trash", "bin"]:
                send2trash(src_path)
                speak(f"Moved {os.path.basename(src_path)} to recycle bin")
                return
            
            dst_folder, mapped = map_folder(dst.lower())
            if dst_folder:
                dst_path = os.path.join(dst_folder, os.path.basename(src_path))
            else:
                # Create custom destination
                if not os.path.isabs(dst):
                    dst = os.path.join(FOLDER_PATHS["desktop"], dst)
                os.makedirs(dst, exist_ok=True)
                dst_path = os.path.join(dst, os.path.basename(src_path))
            
            try:
                os.replace(src_path, dst_path)
                speak(f"Moved to {os.path.dirname(dst_path)}")
                context.update({
                    "last_file": dst_path,
                    "last_folder": os.path.dirname(dst_path)
                })
                clear_folder_cache()
            except Exception as e:
                speak(f"Failed to move file: {e}")
            return

        # DELETE FILE - ENHANCED WITH CONFIRMATION
                # DELETE FILE - ENHANCED (NO CONFIRMATION)
        elif action == "delete_file":
            query = params.get("file", "").strip()
            
            if not query:
                query = context.get("last_file", "")
            
            target_path = None
            
            # Find target file, similar logic to open_file
            possible_matches = []
            current_folder = context.get("last_folder")
            if current_folder:
                _, matches_in_current = smart_find_item(query, current_folder, "files")
                possible_matches.extend([os.path.join(current_folder, m) for m in matches_in_current])
            
            if not possible_matches or len(possible_matches) < 2:
                for folder_name, folder_path in FOLDER_PATHS.items():
                    if folder_path != current_folder:
                        _, matches_in_other = smart_find_item(query, folder_path, "files")
                        possible_matches.extend([os.path.join(folder_path, m) for m in matches_in_other])
            
            seen = set()
            unique_matches = []
            for item in possible_matches:
                if item not in seen:
                    unique_matches.append(item)
                    seen.add(item)

            if not unique_matches:
                speak(f"File '{query}' not found.")
                return

            if len(unique_matches) > 1:
                speak("I found multiple files. Which one would you like to delete? Please say its number:")
                for i, match_path in enumerate(unique_matches[:5]):
                    speak(f"Option {i+1}: {os.path.basename(match_path)} in {os.path.basename(os.path.dirname(match_path))}")
                
                try:
                    r = sr.Recognizer()
                    with sr.Microphone() as source:
                        print("🎤 Listening for choice...")
                        r.adjust_for_ambient_noise(source, duration=0.3)
                        audio = r.listen(source, phrase_time_limit=3)
                    choice_str = r.recognize_google(audio).lower()
                    
                    choice_num = -1
                    if "one" in choice_str or "1" in choice_str: choice_num = 1
                    elif "two" in choice_str or "2" in choice_str: choice_num = 2
                    elif "three" in choice_str or "3" in choice_str: choice_num = 3
                    elif "four" in choice_str or "4" in choice_str: choice_num = 4
                    elif "five" in choice_str or "5" in choice_str: choice_num = 5

                    if 1 <= choice_num <= len(unique_matches[:5]):
                        target_path = unique_matches[choice_num - 1]
                    else:
                        speak("Invalid choice. Aborting deletion.")
                        return # Abort if invalid choice
                except sr.UnknownValueError:
                    speak("Did not catch your choice. Aborting deletion.")
                    return # Abort if no choice
                except Exception as e:
                    speak(f"Error processing choice: {e}. Aborting deletion.")
                    return

            else: # Only one unique match
                target_path = unique_matches[0]

            if target_path and os.path.exists(target_path):
                file_to_delete_name = os.path.basename(target_path)
                try:
                    send2trash(target_path)
                    speak(f"Deleted {file_to_delete_name} and moved it to the recycle bin.")
                    if context.get("last_file") == target_path:
                        context["last_file"] = None
                    clear_folder_cache()
                except Exception as e:
                    speak(f"Failed to delete file: {e}")
            else:
                speak(f"File '{query}' not found for deletion.")
            return

        # SYSTEM OPERATIONS
        elif action == "shutdown":
            speak("Shutting down in 3 seconds.")
            os.system("shutdown /s /t 3")
        elif action == "restart":
            speak("Restarting in 3 seconds.")
            os.system("shutdown /r /t 3")
        elif action == "screenshot":
            try:
                path = params.get("path", f"screenshot_{int(time.time())}.png")
                pyautogui.screenshot(path)
                speak(f"Screenshot saved")
            except Exception as e:
                speak(f"Screenshot failed: {e}")
        
        # WEB OPERATIONS
        elif action == "youtube_search":
            query = params.get("query", "")
            webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
            speak(f"Searching YouTube for {query}")
        elif action == "youtube_control":
            cmd = params.get("command", "").lower()
            key_map = {
                "pause": "k", "play": "k", "stop": "k",
                "next": "shift+n", "previous": "shift+p",
                "volume_up": "volumeup", "volume_down": "volumedown"
            }
            if cmd in key_map:
                if "+" in key_map[cmd]:
                    keys = key_map[cmd].split("+")
                    pyautogui.hotkey(*keys)
                else:
                    pyautogui.press(key_map[cmd])
                speak(f"YouTube {cmd}")
        elif action == "web_search":
            query = params.get("query", "")
            webbrowser.open(f"https://www.google.com/search?q={query}")
            speak(f"Searching for {query}")
        
        # APP CONTROL
        elif action == "app_control":
            app_name = params.get("app", context.get("last_app", ""))
            command = params.get("command", "")
            try:
                windows = gw.getWindowsWithTitle(app_name)
                if windows:
                    window = windows[0]
                    getattr(window, command, lambda: None)()
                    speak(f"{command} executed on {app_name}")
                else:
                    speak(f"No windows found for {app_name}")
            except Exception as e:
                speak(f"Failed to control {app_name}")
        
        # EXIT
        elif action == "exit":
            speak("Goodbye!")
            beep_ok()
            sys.exit(0)
        
        else:
            beep_err()
            speak("Command not recognized.")

    except Exception as e:
        beep_err()
        speak(f"Error executing command: {str(e)[:50]}")
        print(f"Execution error: {e}")

# ---------- OPTIMIZED MAIN LOOP ----------
def listen_and_execute_loop():
    """Optimized main loop with better error recovery"""
    r = sr.Recognizer()
    r.energy_threshold = 4000  # Optimize for better recognition
    r.dynamic_energy_threshold = True
    
    consecutive_errors = 0
    max_errors = 3

    while True:
        try:
            # Try noise-reduced recording first
            audio_arr, sr_hz = record_with_noise_suppression(duration=5)
            
            if audio_arr is not None:
                audio_clean = sr.AudioData(audio_arr.tobytes(), sample_rate=sr_hz, sample_width=2)
                beep_ok()
                command = r.recognize_google(audio_clean)
            else:
                # Fallback to standard microphone
                with sr.Microphone() as source:
                    print("🎤 Listening (fallback)...")
                    r.adjust_for_ambient_noise(source, duration=0.3)
                    audio = r.listen(source, phrase_time_limit=5)
                beep_ok()
                command = r.recognize_google(audio)

            print(f"You said: {command}")
            consecutive_errors = 0  # Reset error counter

            # Parse command
            result = parse_command_online(command)
            if not result or result.get("action") == "unknown":
                result = fallback_parser(command)

            print(f"Parsed: {result}")
            
            # Execute command
            if result.get("action") != "unknown":
                execute(result.get("action"), result.get("params", {}))
            else:
                beep_err()
                speak("I didn't understand that command.")

        except sr.UnknownValueError:
            consecutive_errors += 1
            beep_err()
            if consecutive_errors < max_errors:
                speak("Sorry, please repeat that.")
            else:
                speak("Having trouble hearing you. Please check your microphone.")
                consecutive_errors = 0
                
        except Exception as e:
            consecutive_errors += 1
            beep_err()
            print(f"Loop error: {e}")
            if consecutive_errors < max_errors:
                speak("Something went wrong, please try again.")
            else:
                speak("Multiple errors detected. Restarting listener.")
                consecutive_errors = 0

# ---------- MAIN ----------
def run_assistant():
    print("🤖 Voice Assistant Starting...")
    speak("Voice assistant is ready")

    try:
        listen_and_execute_loop()
    except KeyboardInterrupt:
        speak("Assistant shutting down")
        print("Assistant stopped by user")
    except Exception as e:
        print(f"Fatal error: {e}")
        speak("Assistant encountered a fatal error and will restart")

if __name__ == "__main__":
    run_assistant()

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
    "folder_contents": {},
    "last_operation_details": None
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
                                            exe_path = str(exe_files[0])
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
    name = re.sub(r'\s*\(.*?\)', '', name)
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
    """Extract file extension and core query, handling 'PDF' as a type word."""
    tokens = re.findall(r'\w+', file_query.lower())
    ext = None
    
    # Check for explicit extension words first
    potential_ext_tokens = []
    for t in tokens:
        if t in EXT_WORDS:
            ext = t
        else:
            potential_ext_tokens.append(t)
            
    # Remove the found extension from core_tokens only if it was explicit
    core_tokens = [t for t in potential_ext_tokens if t != ext] if ext else tokens
    
    core_query = ' '.join(core_tokens).strip()
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
    """Enhanced matching with better scoring and explicit extension handling"""
    raw_query_lower = raw_query.lower().strip()
    
    # Extract core query and extension hint once
    q_norm_core, ext_hint = extract_query_and_ext(raw_query_lower)
    
    # If the query itself ends with an extension word, prioritize that in ext_hint
    if not ext_hint:
        for possible_ext in EXT_WORDS:
            if raw_query_lower.endswith(f" {possible_ext}"):
                ext_hint = possible_ext
                q_norm_core = raw_query_lower.replace(f" {possible_ext}", "").strip()
                break

    query_tokens = re.findall(r'\w+', q_norm_core) if q_norm_core else []

    ranked = []
    for item in items:
        # Skip if looking for specific type (existing logic)
        if item_type == "files" and os.path.isdir(os.path.join(context.get("last_folder", ""), item)):
            continue
        elif item_type == "folders" and os.path.isfile(os.path.join(context.get("last_folder", ""), item)):
            continue

        item_lower = item.lower()
        item_base_name, item_ext = os.path.splitext(item_lower)
        item_ext = item_ext[1:] # Remove the dot
        
        score = 0
        
        if item_lower == raw_query_lower:
            score = 1000
        elif item_base_name == q_norm_core and (not ext_hint or item_ext == ext_hint):
            score = 900
        elif normalize_text(item_base_name) == normalize_text(q_norm_core):
            score = 850

        # 2. Strong Containment / Starts With
        elif raw_query_lower in item_lower:
            score = 700 + (len(raw_query_lower) / len(item_lower)) * 100
        elif item_lower.startswith(raw_query_lower):
            score = 650

        # 3. Fuzzy matching with SequenceMatcher (ratio of normalized text)
        norm_item = normalize_text(item_base_name)
        norm_query = normalize_text(q_norm_core)
        if norm_item and norm_query:
            ratio = SequenceMatcher(None, norm_query, norm_item).ratio()
            score += ratio * 500 # Add significant score based on similarity

        # 4. Keyword/Token matching
        item_tokens = re.findall(r'\w+', item_base_name)
        token_hits = sum(1 for qt in query_tokens if qt in item_tokens or any(qt in it for it in item_tokens))
        score += token_hits * 50

        # 5. Extension filtering (penalty if mismatch, bonus if match)
        if ext_hint:
            if item_ext == ext_hint:
                score += 100 # Bonus for correct extension
            elif item_ext != ext_hint and item_ext in EXT_WORDS:
                score -= 200 # Penalty for explicit mismatch
        
        # Special handling for "image/photo" and "pdf" type_filter
        if "image" in raw_query_lower or "photo" in raw_query_lower:
            if item_ext in IMAGE_EXT:
                score += 50
            else:
                score -= 50
        elif "pdf" in raw_query_lower:
             if item_ext == "pdf":
                 score += 50
             else:
                 score -= 50
        
        # Ensure only relevant items are considered
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

# ---------- NEW/IMPROVED RESTORE FUNCTION ----------
def restore_file_from_recycle_bin(file_query):
    """Fixed restore function with proper PowerShell handling"""
    try:
        ps_command = f"""
        $ErrorActionPreference = 'Stop'
        try {{
            $shell = New-Object -ComObject Shell.Application
            $recycleBin = $shell.NameSpace(10)
            
            if ($recycleBin -eq $null) {{
                Write-Output "ERROR: Cannot access Recycle Bin"
                exit 1
            }}
            
            $items = $recycleBin.Items()
            $fileToRestore = $null
            
            # First, try for an exact match
            foreach ($item in $items) {{
                if ($item.Name -eq "{file_query}") {{
                    $fileToRestore = $item
                    break
                }}
            }}
            
            # If no exact match, try for a partial match
            if ($fileToRestore -eq $null) {{
                foreach ($item in $items) {{
                    if ($item.Name -like "*{file_query}*") {{
                        $fileToRestore = $item
                        break
                    }}
                }}
            }}
            
            if ($fileToRestore -ne $null) {{
                # The verb is "undelete" or "restore" depending on Windows version.
                # "undelete" is generally more reliable for scripting.
                $fileToRestore.InvokeVerb("undelete")
                Write-Output "SUCCESS: Restored $($fileToRestore.Name)"
            }} else {{
                Write-Output "ERROR: File not found in Recycle Bin"
            }}
        }} catch {{
            Write-Output "ERROR: $($_.Exception.Message)"
        }}
        """
        
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command],
            capture_output=True,
            text=True,
            timeout=20
        )
        
        if result.returncode == 0 and "SUCCESS:" in result.stdout:
            restored_name = result.stdout.split("SUCCESS: Restored ")[1].strip()
            return True, restored_name
        else:
            error_msg = "Unknown error"
            if "ERROR:" in result.stdout:
                error_msg = result.stdout.split("ERROR: ")[1].strip().split('\n')[0]
            elif result.stderr:
                 error_msg = result.stderr.strip()
            return False, error_msg
            
    except Exception as e:
        return False, str(e)


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
- restore_file: Restore a file from the recycle bin (e.g., ", restore my report from the recycle bin")
    (params: file [string, required])
- undo_last_operation: Undo the last move or delete operation (e.g., ", undo that operation", ", revert last action")

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
"Chief, if you could just restore that deleted document from the bin?" → {{"action":"restore_file","params":{{"file":"that deleted document"}}}}
"Hey , could you undo what I just did, please?" → {{"action":"undo_last_operation","params":{{}}}}


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
        return {"action": "unknown", "params": {}}
    

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
    
    list_rb_match = re.search(r"(?:list|show)(?: all)?(?: the)?\s+(.*?)\s+(?:in|of)\s+(?:the\s+)?recycle bin", c)
    if list_rb_match:
        file_type_query = list_rb_match.group(1).strip()
        params = {"folder": "Recycle Bin"}
        
        # Try to infer type filter (reuse existing logic for list_files)
        for ext_word in EXT_type_query: # type: ignore
            if ext_word in file_type_query:
                params["type_filter"] = ext_word
                break
        
        if "image" in file_type_query or "photo" in file_type_query:
            params["type_filter"] = "image"
        elif "pdf" in file_type_query:
            params["type_filter"] = "pdf"
        # ... (add other type filters if needed, similar to existing list_files in fallback) ...

        return {"action": "list_files", "params": params}

    # Move/delete operations
    move_match = re.search(r"(?:move|send)\s+(.+?)\s+(?:to|into)\s+(.+)", c)
    if move_match:
        return {"action": "move_file", "params": {"source": move_match.group(1), "destination": move_match.group(2)}}

    delete_match = re.search(r"(?:delete|remove|trash)\s+(.+)", c)
    if delete_match:
        return {"action": "delete_file", "params": {"file": delete_match.group(1)}}
    
    # Restore file from recycle bin
    restore_match = re.search(r"(?:restore|recover)\s+(.+?)(?:\s+from\s+(?:the\s+)?recycle bin|\s+from\s+(?:the\s+)?trash)?", c)
    if restore_match:
        return {"action": "restore_file", "params": {"file": restore_match.group(1).strip()}}
    
    # Undo last operation
    if any(kw in c for kw in ["undo", "revert last", "go back on that"]):
        return {"action": "undo_last_operation", "params": {}}

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

# ---------- HELPER FOR FOLDER MAPPING (moved outside execute) ----------
def _map_folder(name: str):
    """Maps a folder name to its path and normalized key."""
    if not name: return None, None
    key = FOLDER_MAP.get(name.lower(), name.lower())
    return FOLDER_PATHS.get(key), key

# ---------- ACTION HANDLER FUNCTIONS ----------

def _handle_open_app(params):
    """Handles the 'open_app' action."""
    app = params.get("app", "").strip()
    
    if app.lower() == "recycle bin":
        try:
            subprocess.Popen(['explorer', 'shell:RecycleBinFolder'])
            speak("Opening Recycle Bin.")
            context.update({
                "last_app": "Recycle Bin",
                "last_action": "open_app", 
                "last_folder": None
            })
            return
        except Exception as e:
            speak(f"Failed to open Recycle Bin: {e}")
            print(f"Error opening Recycle Bin: {e}")

    folder_path, mapped = _map_folder(app.lower())
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

    if execute_open_app_dynamic(app):
        context.update({"last_app": app, "last_action": "open_app"})

def _handle_open_folder(params):
    """Handles the 'open_folder' action."""
    query = params.get("folder", "").strip()
    
    folder_path = None
    current_folder = context.get("last_folder")
    
    mapped_path, mapped_name = _map_folder(query)
    if mapped_path and os.path.isdir(mapped_path):
        folder_path = mapped_path
    
    if not folder_path:
        if current_folder:
            match, _ = smart_find_item(query, current_folder, "folders")
            if match:
                folder_path = os.path.join(current_folder, match)
        
        if not folder_path:
            for sys_folder_name, sys_folder_path in FOLDER_PATHS.items():
                if sys_folder_path != current_folder:
                    match, _ = smart_find_item(query, sys_folder_path, "folders")
                    if match:
                        folder_path = os.path.join(sys_folder_path, match)
                        break

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

def _handle_open_subfolder(params):
    """Handles the 'open_subfolder' action."""
    subfolder = params.get("subfolder", "").strip()
    parent_hint = params.get("parent", "").strip()
    
    parent_folder = None
    if parent_hint:
        folder_path, mapped = _map_folder(parent_hint)
        if folder_path and os.path.isdir(folder_path):
            parent_folder = folder_path
    
    if not parent_folder:
        parent_folder = context.get("last_folder")
    
    if not parent_folder or not os.path.isdir(parent_folder):
        speak("Please specify a parent folder or open one first.")
        return

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
        contents = get_folder_contents(parent_folder)
        available = contents["folders"][:5]
        if available:
            speak(f"Subfolder '{subfolder}' not found. Available: {', '.join(available)}")
        else:
            speak("No subfolders found in this location.")

def _handle_go_back():
    """Handles the 'go_back' action."""
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

def _handle_list_files(params):
    """Handles the 'list_files' action."""
    requested_folder_name = params.get("folder", "").strip().lower()
    type_filter = params.get("type_filter", "").lower() 
    
    if requested_folder_name == "recycle bin":
        try:
            ps_command = f"""
            $shell = New-Object -ComObject Shell.Application
            $recycleBin = $shell.NameSpace(10)
            $items = $recycleBin.Items()
            
            if ($items.Count -eq 0) {{
                Write-Output "Recycle Bin is empty."
            }} else {{
                Write-Output "Recycle Bin Contents:"
                $items | ForEach-Object {{ Write-Output "$($_.Name)|$($_.Path)" }}
            }}
            """
            
            print("Executing PowerShell command to list Recycle Bin contents...")
            result = subprocess.run(
                ['powershell', '-Command', ps_command],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                output_lines = result.stdout.splitlines()
                rb_items = []
                if "Recycle Bin Contents:" in output_lines:
                    for line in output_lines:
                        if "|" in line and not line.startswith("DEBUG"):
                            name, original_path = line.split("|", 1)
                            rb_items.append({"name": name, "original_path": original_path})

                if not rb_items:
                    speak("The Recycle Bin is empty.")
                    context["last_folder"] = None
                    return

                filtered_rb_items = []
                if type_filter:
                    for item in rb_items:
                        item_ext = os.path.splitext(item['name'])[1:].lower().lstrip('.')
                        if type_filter == "image":
                            if item_ext in IMAGE_EXT: filtered_rb_items.append(item)
                        elif type_filter == "pdf":
                            if item_ext == "pdf": filtered_rb_items.append(item)
                        elif type_filter == "document":
                            if item_ext in ["doc", "docx", "txt", "rtf", "odt"]: filtered_rb_items.append(item)
                        elif type_filter == "video":
                            if item_ext in ["mp4", "mkv", "avi", "mov"]: filtered_rb_items.append(item)
                        elif type_filter == "audio":
                            if item_ext in ["mp3", "wav", "flac"]: filtered_rb_items.append(item)
                        elif item_ext == type_filter:
                            filtered_rb_items.append(item)
                else:
                    filtered_rb_items = rb_items

                if not filtered_rb_items:
                    speak(f"No {type_filter} items found in the Recycle Bin.")
                    context["last_folder"] = None
                    return

                print("\n🗑️ Recycle Bin Contents (Filtered):")
                for item in filtered_rb_items[:10]:
                    print(f"  - {item['name']} (from {os.path.basename(item['original_path'])})")
                
                preview_names = [item['name'] for item in filtered_rb_items[:3]]
                speak(f"Found {len(filtered_rb_items)} items in the Recycle Bin. Including: {', '.join(preview_names)}")
                
                context["last_folder"] = None
                context["last_action"] = "list_files"
                return

            else:
                speak("Failed to list Recycle Bin contents.")
                print("PowerShell output:", result.stdout)
                if result.stderr:
                    print("PowerShell error:", result.stderr)
                context["last_folder"] = None
                return

        except Exception as e:
            speak(f"Error listing Recycle Bin contents: {e}")
            print(f"Recycle Bin listing error: {e}")
            context["last_folder"] = None
            return

    folder_path = None
    if requested_folder_name:
        mapped_path, _ = _map_folder(requested_folder_name.lower())
        if mapped_path and os.path.isdir(mapped_path):
            folder_path = mapped_path
        else:
            speak(f"Could not find a system folder named '{requested_folder_name}'.")
            return
    else:
        folder_path = context.get("last_folder")

    if not folder_path or not os.path.isdir(folder_path):
        speak("No folder is currently open or specified.")
        return

    contents = get_folder_contents(folder_path)
    
    items_to_list = []
    if type_filter:
        for item in contents["files"] + contents["folders"]:
            item_ext = os.path.splitext(item).lower().lstrip('.')
            if type_filter == "image":
                if item_ext in IMAGE_EXT: items_to_list.append(item)
            elif type_filter == "pdf":
                if item_ext == "pdf": items_to_list.append(item)
            elif type_filter == "document":
                if item_ext in ["doc", "docx", "txt", "rtf", "odt"]: items_to_list.append(item)
            elif type_filter == "video":
                if item_ext in ["mp4", "mkv", "avi", "mov"]: items_to_list.append(item)
            elif type_filter == "audio":
                if item_ext in ["mp3", "wav", "flac"]: items_to_list.append(item)
            elif item_ext == type_filter:
                items_to_list.append(item)
    else:
        items_to_list = contents["files"] + contents["folders"]

    if items_to_list:
        speak(f"In {os.path.basename(folder_path)}, I found: {', '.join(items_to_list[:5])}. There are {len(items_to_list)} items in total.")
    else:
        speak(f"No items found in {os.path.basename(folder_path)} (or no matching type).")
    context.update({"last_action": "list_files", "last_folder": folder_path})


def _handle_open_file(params):
    """Handles the 'open_file' action."""
    query = params.get("file", "").strip()
    
    file_path = None
    current_folder = context.get("last_folder")
    
    possible_matches = []
    if current_folder:
        _, matches_in_current = smart_find_item(query, current_folder, "files")
        possible_matches.extend([os.path.join(current_folder, m) for m in matches_in_current])
    
    if not possible_matches or len(possible_matches) < 2:
        for folder_name, folder_path_item in FOLDER_PATHS.items():
            if folder_path_item != current_folder:
                _, matches_in_other = smart_find_item(query, folder_path_item, "files")
                possible_matches.extend([os.path.join(folder_path_item, m) for m in matches_in_other])
    
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
        speak("I found multiple files. Please choose one by saying its number:")
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
                file_path = unique_matches[choice_num - 1]
            else:
                speak("Invalid choice. Opening the first match.")
                file_path = unique_matches
        except sr.UnknownValueError:
            speak("Did not catch your choice. Opening the first match.")
            file_path = unique_matches
        except Exception as e:
            speak(f"Error processing choice: {e}. Opening the first match.")
            file_path = unique_matches
    else:
        file_path = unique_matches

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

def _handle_move_file(params):
    """Handles the 'move_file' action."""
    src_query = params.get("source", "").strip()
    dst_query = params.get("destination", "").strip()
    
    if not src_query:
        src_query = context.get("last_file", "")
    
    if not src_query or not dst_query:
        speak("Please specify both source and destination for moving.")
        return
    
    src_path = None
    if os.path.isabs(src_query) and os.path.exists(src_query):
        src_path = src_query
    else:
        current_folder = context.get("last_folder")
        if current_folder:
            match, _ = smart_find_item(src_query, current_folder, "files")
            if match:
                src_path = os.path.join(current_folder, match)
        
        if not src_path:
            src_path, _ = search_all_folders_for_item(src_query, "files")
    
    if not src_path:
        speak(f"Source file '{src_query}' not found.")
        return
    
    original_dst_path = None
    
    if dst_query.lower() in ["recycle bin", "trash", "bin"]:
        speak("Moving files to recycle bin is now handled by delete_file action.")
        try:
            send2trash(src_path)
            speak(f"Moved {os.path.basename(src_path)} to recycle bin.")
            clear_folder_cache()
        except Exception as e:
            speak(f"Failed to move {os.path.basename(src_path)} to recycle bin: {e}")
        return
    
    dst_folder_path = None
    mapped_path, _ = _map_folder(dst_query.lower())
    if mapped_path and os.path.isdir(mapped_path):
        dst_folder_path = mapped_path
    else:
        current_folder = context.get("last_folder")
        if current_folder:
            match, _ = smart_find_item(dst_query, current_folder, "folders")
            if match:
                dst_folder_path = os.path.join(current_folder, match)
        
        if not dst_folder_path:
            found_folder_path, _ = search_all_folders_for_item(dst_query, "folders")
            if found_folder_path and os.path.isdir(found_folder_path):
                dst_folder_path = found_folder_path
        
        if not dst_folder_path:
            speak(f"Could not find destination folder '{dst_query}'. Creating a new one on desktop.")
            dst_folder_path = os.path.join(FOLDER_PATHS["desktop"], dst_query)
            os.makedirs(dst_folder_path, exist_ok=True)

    if not dst_folder_path:
        speak(f"Could not determine destination folder for '{dst_query}'.")
        return

    original_dst_path = os.path.join(dst_folder_path, os.path.basename(src_path))
    
    try:
        context["last_operation_details"] = {
            "action": "move",
            "source_path": src_path,
            "destination_path": original_dst_path,
            "original_parent_folder": os.path.dirname(src_path),
            "new_parent_folder": dst_folder_path,
            "file_name": os.path.basename(src_path)
        }

        os.replace(src_path, original_dst_path)
        speak(f"Moved '{os.path.basename(src_path)}' to '{os.path.basename(dst_folder_path)}'.")
        context.update({
            "last_file": original_dst_path,
            "last_folder": dst_folder_path,
            "last_action": "move_file"
        })
        clear_folder_cache()
    except Exception as e:
        speak(f"Failed to move file: {e}")
        context["last_operation_details"] = None

def _handle_delete_file(params):
    """Handles the 'delete_file' action."""
    query = params.get("file", "").strip()
    
    if not query:
        query = context.get("last_file", "")
    
    target_path = None
    
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
                return
        except sr.UnknownValueError:
            speak("Did not catch your choice. Aborting deletion.")
            return
        except Exception as e:
            speak(f"Error processing choice: {e}. Aborting deletion.")
            return

    else:
        target_path = unique_matches

    if target_path and os.path.exists(target_path):
        file_to_delete_name = os.path.basename(target_path)
        try:
            context["last_operation_details"] = {
                "action": "delete",
                "deleted_path": target_path,
                "original_parent_folder": os.path.dirname(target_path),
                "file_name": file_to_delete_name
            }
            send2trash(target_path)
            speak(f"Deleted {file_to_delete_name} and moved it to the recycle bin.")
            if context.get("last_file") == target_path:
                context["last_file"] = None
            clear_folder_cache()
        except Exception as e:
            speak(f"Failed to delete file: {e}")
            context["last_operation_details"] = None
    else:
        speak(f"File '{query}' not found for deletion.")

def _handle_restore_file(params):
    """Handles the 'restore_file' action."""
    file_query = params.get("file", "").strip()
    if not file_query:
        speak("Please tell me which file to restore from the recycle bin.")
        return

    speak(f"Searching for '{file_query}' in the recycle bin...")
    success, result_msg = restore_file_from_recycle_bin(file_query)
    
    if success:
        speak(f"Successfully restored '{result_msg}'.")
        clear_folder_cache()
    else:
        speak(f"Failed to restore the file. Reason: {result_msg}")

def _handle_undo_last_operation():
    """Handles the 'undo_last_operation' action."""
    details = context.get("last_operation_details")
    if not details:
        speak("No recent operation to undo.")
        return
    
    op_type = details.get("action")
    
    if op_type == "move":
        source_path = details.get("source_path")
        destination_path = details.get("destination_path")
        original_parent_folder = details.get("original_parent_folder")
        file_name = details.get("file_name")

        if not os.path.exists(destination_path):
            speak(f"Cannot undo move: '{file_name}' not found in its new location. It might have been moved again.")
            context["last_operation_details"] = None
            return
        
        try:
            os.replace(destination_path, source_path)
            speak(f"Undo successful. '{file_name}' moved back to '{os.path.basename(original_parent_folder)}'.")
            context.update({
                "last_file": source_path,
                "last_folder": original_parent_folder,
                "last_action": "undo_move"
            })
            context["last_operation_details"] = None
            clear_folder_cache()
        except Exception as e:
            speak(f"Failed to undo move operation: {e}")
            print(f"Undo move error: {e}")

    elif op_type == "delete":
        file_name = details.get("file_name")
        
        speak(f"Attempting to restore '{file_name}' from the recycle bin...")
        success, result_msg = restore_file_from_recycle_bin(file_name)

        if success:
            speak(f"Undo successful. '{result_msg}' has been restored.")
            context.update({
                "last_file": details.get("deleted_path"),
                "last_folder": details.get("original_parent_folder"),
                "last_action": "undo_delete"
            })
            context["last_operation_details"] = None
            clear_folder_cache()
        else:
            speak(f"Could not undo the delete. Reason: {result_msg}")

    else:
        speak("Cannot undo this type of operation.")

def _handle_shutdown():
    """Handles the 'shutdown' action."""
    speak("Shutting down in 3 seconds.")
    os.system("shutdown /s /t 3")

def _handle_restart():
    """Handles the 'restart' action."""
    speak("Restarting in 3 seconds.")
    os.system("shutdown /r /t 3")

def _handle_screenshot(params):
    """Handles the 'screenshot' action."""
    try:
        path = params.get("path", f"screenshot_{int(time.time())}.png")
        pyautogui.screenshot(path)
        speak(f"Screenshot saved")
    except Exception as e:
        speak(f"Screenshot failed: {e}")

def _handle_youtube_search(params):
    """Handles the 'youtube_search' action."""
    query = params.get("query", "")
    webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
    speak(f"Searching YouTube for {query}")

def _handle_youtube_control(params):
    """Handles the 'youtube_control' action."""
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

def _handle_web_search(params):
    """Handles the 'web_search' action."""
    query = params.get("query", "")
    webbrowser.open(f"https://www.google.com/search?q={query}")
    speak(f"Searching for {query}")

def _handle_app_control(params):
    """Handles the 'app_control' action."""
    app_name_query = params.get("app", context.get("last_app", "")).strip().lower()
    command = params.get("command", "").strip().lower()

    if not app_name_query:
        speak("Please specify which application to control.")
        return
    if not command in ["close", "minimize", "maximize", "restore"]:
        speak(f"Command '{command}' is not supported.")
        return

    window_title_to_find = FOLDER_MAP.get(app_name_query, app_name_query).capitalize()

    try:
        windows = gw.getWindowsWithTitle(window_title_to_find)
        if windows:
            for window in windows:
                if command == "close":
                    window.close()
                elif command == "minimize":
                    window.minimize()
                elif command == "maximize":
                    window.maximize()
                elif command == "restore":
                    window.restore()
            speak(f"Action '{command}' executed on '{window_title_to_find}'.")
        else:
            speak(f"No windows found for '{window_title_to_find}'.")
    except Exception as e:
        speak(f"Failed to control '{window_title_to_find}': {e}")
        print(f"Error in app_control: {e}")

def _handle_exit():
    """Handles the 'exit' action."""
    speak("Goodbye!")
    beep_ok()
    sys.exit(0)

# ---------- ENHANCED EXECUTION (dispatcher) ----------
def execute(action, params):
    """Enhanced execution with better error handling and persistence"""
    
    try:
        if action == "open_app":
            _handle_open_app(params)
        elif action == "open_folder":
            _handle_open_folder(params)
        elif action == "open_subfolder":
            _handle_open_subfolder(params)
        elif action == "go_back":
            _handle_go_back()
        elif action == "list_files":
            _handle_list_files(params)
        elif action == "open_file":
            _handle_open_file(params)
        elif action == "move_file":
            _handle_move_file(params)
        elif action == "delete_file":
            _handle_delete_file(params)
        elif action == "restore_file":
            _handle_restore_file(params)
        elif action == "undo_last_operation":
            _handle_undo_last_operation()
        elif action == "shutdown":
            _handle_shutdown()
        elif action == "restart":
            _handle_restart()
        elif action == "screenshot":
            _handle_screenshot(params)
        elif action == "youtube_search":
            _handle_youtube_search(params)
        elif action == "youtube_control":
            _handle_youtube_control(params)
        elif action == "web_search":
            _handle_web_search(params)
        elif action == "app_control":
            _handle_app_control(params)
        elif action == "exit":
            _handle_exit()
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
    r.energy_threshold = 4000
    r.dynamic_energy_threshold = True
    
    consecutive_errors = 0
    max_errors = 3

    while True:
        try:
            audio_arr, sr_hz = record_with_noise_suppression(duration=5)
            
            if audio_arr is not None:
                audio_clean = sr.AudioData(audio_arr.tobytes(), sample_rate=sr_hz, sample_width=2)
                beep_ok()
                command = r.recognize_google(audio_clean)
            else:
                with sr.Microphone() as source:
                    print("🎤 Listening (fallback)...")
                    r.adjust_for_ambient_noise(source, duration=0.3)
                    audio = r.listen(source, phrase_time_limit=5)
                beep_ok()
                command = r.recognize_google(audio)

            print(f"You said: {command}")
            consecutive_errors = 0

            result = parse_command_online(command)
            if not result or result.get("action") == "unknown":
                result = fallback_parser(command)

            print(f"Parsed: {result}")
            
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
import webview
import psutil
import threading
import time
import pyttsx3
import speech_recognition as sr
import os
import json
import requests
import subprocess
import pyautogui
import re
import datetime
from pathlib import Path
from bs4 import BeautifulSoup
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# pocketsphinx for offline wake word recognition (optional)
try:
    import pocketsphinx  # noqa: F401
    POCKETSPHINX_AVAILABLE = True
except ImportError:
    POCKETSPHINX_AVAILABLE = False
    print("Optional: install pocketsphinx for wake-word detection: pip install pocketsphinx")

# Try to import keyboard for global hotkey support
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False
    print("Optional: Install 'keyboard' package for global hotkey support: pip install keyboard")

# ============= ENHANCED TEXT-TO-SPEECH =============
# Pre-initialize TTS engine once to avoid blocking
_tts_engine = None
_tts_lock = threading.Lock()

def init_tts():
    """Initialize TTS engine once"""
    global _tts_engine
    try:
        _tts_engine = pyttsx3.init()
        voices = _tts_engine.getProperty('voices')
        print(f"[TTS] voices available: {[v.name for v in voices]}")
        if len(voices) > 1:
            _tts_engine.setProperty('voice', voices[1].id)
        _tts_engine.setProperty('rate', 160)
        _tts_engine.setProperty('volume', 1.0)
    except Exception as e:
        print(f"TTS init failed: {e}")

def speak(text, priority=False):
    """Improved TTS with non-blocking speech"""
    def ensure_sound():
        """Unmute/set volume if system audio is silent"""
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            if volume.GetMute():
                volume.SetMute(0, None)
            if volume.GetMasterVolumeLevelScalar() < 0.1:
                volume.SetMasterVolumeLevelScalar(0.5, None)
        except Exception:
            pass

    def run_speak():
        try:
            global _tts_engine
            if not _tts_engine:
                init_tts()
                time.sleep(0.5)  # Wait for init
            
            # make sure system sound is audible
            ensure_sound()

            if _tts_engine:
                with _tts_lock:
                    print(f"[TTS] speaking: {text}")
                    _tts_engine.say(text)
                    _tts_engine.runAndWait()
        except Exception as e:
            print(f"Speak error: {e}")
    
    thread = threading.Thread(target=run_speak, daemon=True)
    thread.start()

def clean_html(text):
    """Remove HTML tags from text"""
    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# ============= TASK SCHEDULER =============
class TaskScheduler:
    def __init__(self):
        self.tasks = []
    
    def schedule_task(self, delay_seconds, task_name, action_func):
        """Schedule a task to run after delay"""
        def run_delayed():
            time.sleep(delay_seconds)
            action_func()
        
        thread = threading.Thread(target=run_delayed, daemon=True)
        self.tasks.append({'name': task_name, 'thread': thread})
        thread.start()
    
    def add_reminder(self, message, delay_minutes, send_ui_func):
        """Schedule a reminder"""
        delay_seconds = delay_minutes * 60
        def reminder_action():
            speak(f"Reminder: {message}")
            send_ui_func(f"<b>⏰ Reminder:</b> {message}")
        
        self.schedule_task(delay_seconds, f"Reminder: {message}", reminder_action)

# ============= CONVERSATION MEMORY =============
class ConversationMemory:
    def __init__(self):
        self.history = []
    
    def add(self, role, content):
        """Add to conversation history"""
        self.history.append({
            'timestamp': datetime.datetime.now(),
            'role': role,  # 'user' or 'assistant'
            'content': content
        })
    
    def get_context(self, limit=5):
        """Get recent conversation context"""
        return self.history[-limit:]

# ============= NATURAL LANGUAGE PROCESSING =============
class CommandResolver:
    """Maps natural language to system commands"""
    
    def __init__(self):
        self.commands = {
            # System Info
            'battery|power|juice|percent': ('battery', 'get_battery'),
            'cpu|processor|usage': ('system', 'get_cpu'),
            'memory|ram': ('system', 'get_memory'),
            'disk|storage': ('system', 'get_disk'),
            
            # Voice & Audio
            'mute|quiet': ('audio', 'mute'),
            'unmute|sound': ('audio', 'unmute'),
            'volume|loud': ('audio', 'set_volume'),
            
            # Display
            'brightness|screen|dark|dim': ('display', 'set_brightness'),
            'full.*screen|maximize': ('window', 'fullscreen'),
            
            # Network
            'wifi|internet|network': ('network', 'toggle_wifi'),
            'connect|disconnect': ('network', 'toggle_wifi'),
            
            # Apps
            # handle explicit folder/directory requests separately so we can
            # clean the query before opening.  file_ops will strip the
            # hint word and fall back to the smart launcher if needed.
            'open\s+(?:folder|directory)\b': ('files', 'file_ops'),
            'open|launch|start': ('app', 'open_app'),
            'close|kill|terminate|exit': ('app', 'close_app'),
            'list.*apps|running': ('app', 'list_apps'),
            
            # System Control
            'lock|secure': ('system', 'lock'),
            'sleep|hibernate|rest': ('system', 'sleep'),
            'restart|reboot': ('system', 'restart'),
            'shutdown|power.*off': ('system', 'shutdown'),
            
            # Screenshots & Vision
            'screenshot|snap|capture|screen': ('vision', 'screenshot'),
            
            # Help
            'help|commands|what.*can.*do|abilities': ('help', 'show_help'),
            
            # Conversation
            'hello|hi|hey|good.*morning|good.*afternoon': ('chat', 'greet'),
            'how are you|how.*doing|what.?s up': ('chat', 'status'),
            'thank|thanks|appreciate': ('chat', 'acknowledge'),
            'who.*are.*you|tell.*about.*yourself': ('chat', 'identify'),
            'good|fine|okay|idle|not bad': ('chat', 'reply_status'),
        }
    
    def resolve(self, query):
        """Resolve user query to command type"""
        query_lower = query.lower()
        for pattern, (category, action) in self.commands.items():
            if re.search(pattern, query_lower):
                return (category, action, query)
        return None

# ============= THE MAIN API =============
class NelaAPI:
    def __init__(self):
        self.window = None
        self.identity = "I am Nela, your Windows AI assistant. I can control your system, search the web, and help with tasks. Just speak commands naturally!"
        self.memory = ConversationMemory()
        self.scheduler = TaskScheduler()
        self.resolver = CommandResolver()
        self.listening = False
        self.confirm_pending = None
        
    def send_to_ui(self, text, timestamp=True):
        """Send message to webview"""
        if self.window:
            if timestamp:
                ts = datetime.datetime.now().strftime("%H:%M:%S")
                formatted = f"<span style='color:#888;font-size:12px'>[{ts}]</span> {text}"
            else:
                formatted = text
            
            safe_text = json.dumps(formatted)
            try:
                self.window.evaluate_js(f"nelaResponse({safe_text})")
            except:
                pass

    # ============= CONVERSATION & HELP =============
    def show_help(self):
        """Display available commands"""
        help_text = """
        <b>🎤 Nela Commands:</b><br>
        <b>System:</b> battery, cpu, memory, disk, lock, sleep, restart, shutdown<br>
        <b>Audio:</b> mute, unmute, volume UP/DOWN<br>
        <b>Display:</b> brightness UP/DOWN, fullscreen<br>
        <b>Network:</b> wifi on/off<br>
        <b>Apps:</b> open [app], close [app], list apps<br>
        <b>Files:</b> open [file/folder]<br>
        <b>Smart:</b> screenshot, search anything, set reminder<br>
        <b>Help:</b> say "help" anytime
        """
        self.send_to_ui(help_text)
        speak("Nela is ready. I can control your system, manage apps, search the web, and more. Say help for commands.")

    def greet(self, query):
        """Friendly greeting"""
        greetings = [
            "Hello! Ready to assist.",
            "Hey there! What do you need?",
            "Good to see you! How can I help?",
            "Hi! At your service."
        ]
        greeting = greetings[hash(query) % len(greetings)]
        self.send_to_ui(greeting)
        speak(greeting)

    def identify(self, query):
        """Tell user about Nela"""
        self.send_to_ui(self.identity)
        speak(self.identity)

    # ============= SYSTEM COMMANDS =============
    def get_battery(self):
        """Get battery percentage and status"""
        print("[DEBUG] get_battery called")
        try:
            bat = psutil.sensors_battery()
            print(f"[DEBUG] battery object: {bat}")
            if bat:
                status = "charging" if bat.power_plugged else "on battery"
                msg = f"⚡ Battery: {bat.percent}% | {status}"
                self.send_to_ui(msg)
                speak(f"Battery is at {bat.percent} percent, {status}.")
            else:
                self.send_to_ui("No battery detection")
                speak("Battery information not available.")
        except Exception as e:
            print(f"[DEBUG] get_battery exception: {e}")
            self.send_to_ui("Cannot read battery")
            speak("Unable to read battery status.")

    def get_cpu(self):
        """Get CPU usage"""
        try:
            cpu = psutil.cpu_percent(interval=1)
            msg = f"📊 CPU Usage: {cpu}%"
            self.send_to_ui(msg)
            speak(f"CPU is at {cpu} percent.")
        except:
            self.send_to_ui("Cannot read CPU")

    def get_memory(self):
        """Get RAM usage"""
        try:
            mem = psutil.virtual_memory()
            msg = f"💾 Memory: {mem.percent}% used ({mem.available // (1024**3)}GB free)"
            self.send_to_ui(msg)
            speak(f"Memory is {mem.percent} percent used.")
        except:
            self.send_to_ui("Cannot read memory")

    def get_disk(self):
        """Get disk usage"""
        try:
            disk = psutil.disk_usage('C:')
            msg = f"💿 Disk: {disk.percent}% used"
            self.send_to_ui(msg)
            speak(f"Disk is {disk.percent} percent used.")
        except:
            self.send_to_ui("Cannot read disk")

    def set_brightness(self, query):
        """Control screen brightness"""
        nums = re.findall(r'\d+', query)
        level = int(nums[0]) if nums else 50
        level = max(0, min(100, level))
        
        try:
            os.system(f'powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{level})')
            msg = f"🌞 Brightness set to {level}%"
            self.send_to_ui(msg)
            speak(f"Brightness adjusted to {level} percent.")
        except:
            self.send_to_ui("Cannot adjust brightness")

    def set_volume(self, query):
        """Control system volume"""
        try:
            nums = re.findall(r'\d+', query)
            vol_level = int(nums[0]) if nums else 50
            vol_level = max(0, min(100, vol_level))
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevelScalar(vol_level / 100, None)
            
            msg = f"🔊 Volume set to {vol_level}%"
            self.send_to_ui(msg)
            speak(f"Volume set to {vol_level} percent.")
        except:
            self.send_to_ui("Cannot control volume")

    def mute(self):
        """Mute system"""
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMute(1, None)
            self.send_to_ui("🔇 System muted")
            speak("Muted.")
        except:
            self.send_to_ui("Cannot mute")

    def unmute(self):
        """Unmute system"""
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMute(0, None)
            self.send_to_ui("🔊 System unmuted")
            speak("Unmuted.")
        except:
            self.send_to_ui("Cannot unmute")

    def lock(self):
        """Lock the system"""
        self.require_confirmation("Lock your computer?", lambda: os.system("rundll32.exe user32.dll,LockWorkStation"))
        self.send_to_ui("🔒 System locked.")
        speak("Locking the system.")

    def sleep(self):
        """Put system to sleep"""
        self.require_confirmation("Sleep now?", lambda: os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0"))
        self.send_to_ui("😴 Going to sleep...")
        speak("Entering sleep mode.")

    def restart(self):
        """Restart system"""
        self.require_confirmation("Restart your computer? (closes all apps)", lambda: os.system("shutdown /r /t 30"))
        self.send_to_ui("🔄 Restarting in 30 seconds...")
        speak("System will restart in 30 seconds.")

    def shutdown(self):
        """Shut down system"""
        self.require_confirmation("Shut down? (closes all apps)", lambda: os.system("shutdown /s /t 30"))
        self.send_to_ui("⏻️  Shutting down in 30 seconds...")
        speak("System will shut down in 30 seconds.")

    def require_confirmation(self, message, action_func):
        """Require user confirmation for destructive actions"""
        self.send_to_ui(f"<b style='color:#ff9800'>⚠️  {message}</b>")
        self.confirm_pending = action_func

    # ============= WINDOW OPERATIONS =============
    def fullscreen(self):
        """Toggle fullscreen"""
        try:
            pyautogui.press('f11')
            self.send_to_ui("⛶ Fullscreen toggled")
        except:
            self.send_to_ui("Cannot toggle fullscreen")

    # ============= NETWORK OPERATIONS =============
    def toggle_wifi(self, query):
        """Toggle WiFi on or off"""
        if "off" in query or "disable" in query:
            os.system('netsh interface set interface "Wi-Fi" disabled')
            self.send_to_ui("📡 WiFi disabled")
            speak("WiFi is off.")
        else:
            os.system('netsh interface set interface "Wi-Fi" enabled')
            self.send_to_ui("📡 WiFi enabled")
            speak("WiFi is on.")

    # ============= APP MANAGEMENT =============
    def open_app(self, query):
        """Smart folder/app launcher.  Scan common locations first, then
        try launching by name, finally fall back to start.  Strips extra
        words like "app" or "folder" so queries feel natural.
        """
        item = re.sub(r'(open|launch|start)', '', query, flags=re.IGNORECASE).strip()
        # drop hint words that are not part of the actual name
        item = re.sub(r"\b(app|application|folder|directory)\b", "", item, flags=re.IGNORECASE).strip()

        # look in a few default user folders
        zones = [f"~/{item}", f"~/Desktop/{item}", f"~/Documents/{item}", f"~/Downloads/{item}"]
        for z in zones:
            path = os.path.expanduser(z)
            if os.path.exists(path):
                try:
                    os.startfile(path)
                    self.send_to_ui(f"Path located. Opening {item}.")
                    speak(f"Opening {item}.")
                except Exception:
                    self.send_to_ui(f"Found {item} but failed to open")
                    speak(f"Found {item} but could not open it.")
                return

        # nothing in the zones; try to execute by name (may be on PATH)
        try:
            subprocess.Popen(item)
            self.send_to_ui(f"🚀 Launching {item}...")
            speak(f"Launching {item}.")
            return
        except Exception:
            pass

        # final fallback - ask Windows to start it (might open a query or folder)
        try:
            os.system(f'start "" "{item}"')
            self.send_to_ui(f"🚀 Launching {item}...")
            speak(f"Launching {item}.")
        except Exception:
            self.send_to_ui(f"Cannot find {item}")
            speak(f"Could not find {item}.")

    def close_app(self, query):
        """Close application"""
        app = re.sub(r'(close|kill|terminate|exit)', '', query).strip()
        try:
            os.system(f"taskkill /f /im {app}.exe")
            self.send_to_ui(f"✖️ Closing {app}...")
            speak(f"Closing {app}.")
        except:
            self.send_to_ui(f"Cannot close {app}")

    def list_apps(self):
        """List running applications"""
        try:
            processes = psutil.process_iter(['pid', 'name'])
            apps = [p.info['name'] for p in processes][:15]  # First 15
            app_list = ", ".join(apps)
            self.send_to_ui(f"🔍 Running: {app_list}")
            speak("List of running apps shown.")
        except:
            self.send_to_ui("Cannot list applications")

    # ============= VISION & SNAPSHOT =============
    def screenshot(self, query):
        """Take screenshot"""
        try:
            path = os.path.expanduser("~/Desktop/Nela_Snapshot.png")
            pyautogui.screenshot(path)
            self.send_to_ui(f"📸 Screenshot saved to Desktop")
            speak("Screenshot captured.")
        except:
            self.send_to_ui("Cannot take screenshot")

    # ============= WEB SCRAPING & SEARCH =============
    def web_scrape(self, query):
        """Search the web and extract answer"""
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            }
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            res = requests.get(url, headers=headers, timeout=7)
            soup = BeautifulSoup(res.text, 'html.parser')
            
            # Try to find quick answer first
            ans = soup.find('div', class_='VwiC3b') or soup.find('span', class_='hgKElc') or soup.find('div', class_='Z0LcW')
            if ans:
                result = clean_html(ans.get_text())[:200]  # First 200 chars
                self.send_to_ui(f"<b>Answer:</b> {result}")
                speak(result)
            else:
                # Try snippet
                snippet = soup.find('div', class_='BNeawe vvjwJb AP7Wnd')
                if snippet:
                    text = clean_html(snippet.get_text())[:200]
                    self.send_to_ui(f"<b>Found:</b> {text}")
                    speak("Here is what I found.")
                else:
                    # No extract; open Chrome
                    self.send_to_ui("Opening Chrome to search...")
                    speak("Opening browser.")
                    os.system(f"start {url}")
        except:
            self.send_to_ui("Network error. Try again.")
            speak("Network failure.")

    # ============= REMINDER & SCHEDULING =============
    def set_reminder(self, query):
        """Set a reminder"""
        match = re.search(r'remind.*?in\s+(\d+)\s+(minutes?|hours?)', query, re.IGNORECASE)
        if match:
            amount = int(match.group(1))
            unit = match.group(2).lower()
            delay = amount * (60 if 'min' in unit else 3600)
            
            # Extract message
            msg = re.sub(r'remind.*?in.*?minutes?|hours?', '', query).strip()
            if not msg:
                msg = "Your reminder"
            
            self.scheduler.schedule_task(delay, f"Reminder: {msg}", 
                lambda: self.send_to_ui(f"<b>🔔 Reminder:</b> {msg}"))
            self.send_to_ui(f"⏰ Reminder set: {msg} in {amount} {unit}")
            speak(f"Reminder set for {amount} {unit}.")
        else:
            self.send_to_ui("Say: 'remind me in 5 minutes to [task]'")

    # ============= FILE OPERATIONS =============
    def file_ops(self, query):
        """Handle explicit folder/directory requests.
        Strip the hint word and then attempt to open the path.  Falls
        back to the smart launcher if the cleaned path isn't a folder.
        """
        match = re.search(r'(open|show|find)\s+(.+)', query, re.IGNORECASE)
        if not match:
            return
        item = match.group(2).strip()
        # remove the word "folder" or "directory" if user said it
        item = re.sub(r"\b(folder|directory)\b", "", item, flags=re.IGNORECASE).strip()
        path = os.path.expanduser(item)
        if os.path.isdir(path):
            try:
                os.startfile(path)
                self.send_to_ui(f"📂 Opening folder {item}...")
                speak(f"Opening folder {item}.")
            except Exception:
                self.send_to_ui(f"Found {item} but could not open it.")
                speak(f"Found {item} but could not open it.")
            return
        # not a folder? hand off to general opener
        self.open_app(f"open {item}")

    # ============= COMMAND HANDLER =============
    def handle_command(self, query):
        """Main command router"""
        query = query.lower().strip()
        if not query:
            return
        
        # Strip wake words ("hey nela", "ok nela", "nela ")
        wake_words = ['hey nela ', 'ok nela ', 'nela ']
        for wake in wake_words:
            if query.startswith(wake):
                query = query[len(wake):].strip()
                break
        
        if not query:
            return
        
        self.send_to_ui(f"<b>You:</b> {query}", timestamp=False)
        self.memory.add('user', query)
        
        self.send_to_ui("<i>🤔 Thinking...</i>", timestamp=False)
        time.sleep(0.5)
        
        # Resolve command using NLP
        print(f"[DEBUG] handle_command input: '{query}'")
        resolved = self.resolver.resolve(query)
        print(f"[DEBUG] resolved result: {resolved}")
        
        if resolved:
            category, action, _ = resolved
            
            if category == 'help':
                self.show_help()
            elif category == 'chat':
                if action == 'greet':
                    self.greet(query)
                elif action == 'identify':
                    self.identify(query)
                elif action == 'status':
                    # small talk query
                    resp = "I'm doing great, thanks for asking!"
                    self.send_to_ui(resp)
                    speak(resp)
                elif action == 'reply_status':
                    # user replied to the bot's status question
                    resp = "Glad to hear that!"
                    self.send_to_ui(resp)
                    speak(resp)
                elif action == 'acknowledge':
                    speak("You're welcome.")
                    self.send_to_ui("You're welcome!")
            elif category == 'system':
                if action == 'get_battery':
                    self.get_battery()
                elif action == 'get_cpu':
                    self.get_cpu()
                elif action == 'get_memory':
                    self.get_memory()
                elif action == 'get_disk':
                    self.get_disk()
                elif action == 'lock':
                    self.lock()
                elif action == 'sleep':
                    self.sleep()
                elif action == 'restart':
                    self.restart()
                elif action == 'shutdown':
                    self.shutdown()
            elif category == 'audio':
                if action == 'mute':
                    self.mute()
                elif action == 'unmute':
                    self.unmute()
                elif action == 'set_volume':
                    self.set_volume(query)
            elif category == 'display':
                if action == 'set_brightness':
                    self.set_brightness(query)
            elif category == 'window':
                if action == 'fullscreen':
                    self.fullscreen()
            elif category == 'network':
                self.toggle_wifi(query)
            elif category == 'app':
                if action == 'open_app':
                    self.open_app(query)
                elif action == 'close_app':
                    self.close_app(query)
                elif action == 'list_apps':
                    self.list_apps()
            elif category == 'vision':
                self.screenshot(query)
        else:
            # Default: search the web
            self.send_to_ui("<i>Searching the web...</i>", timestamp=False)
            speak("Let me look that up for you.")
            threading.Thread(target=self.web_scrape, args=(query,), daemon=True).start()

    # ============= VOICE COMMAND =============
    def start_voice_cmd(self):
        """Start listening for voice input"""
        # when invoked (either by UI button, hotkey, or wake word) we
        # indicate listening and spawn recognition thread
        self.send_to_ui("<b style='color:#00bcd4;'>🎤 Listening...</b>", timestamp=False)
        threading.Thread(target=self.process_voice, daemon=True).start()

    def wake_listener(self):
        """Continuously listen for wake-word using pocketsphinx."""
        r = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                r.adjust_for_ambient_noise(source, duration=1)
                # keyword_entries allow simple keyword spotting
                keywords = [("hey nela", 1.0), ("hi nela", 1.0), ("hello nela", 1.0)]
                while True:
                    try:
                        audio = r.listen(source, timeout=None)
                        text = r.recognize_sphinx(audio, keyword_entries=keywords)
                        if text and any(w in text.lower() for w, _ in keywords):
                            # wake word detected
                            self.send_to_ui("<i>🔊 Wake word heard</i>", timestamp=False)
                            speak("Yes?")
                            # small delay to avoid immediate re-trigger
                            time.sleep(0.5)
                            self.start_voice_cmd()
                    except sr.UnknownValueError:
                        # nothing recognizable, just continue listening
                        continue
                    except Exception as e:
                        print(f"Wake listener error: {e}")
                        time.sleep(1)
        except Exception as e:
            print(f"Could not start wake listener: {e}")

    def process_voice(self):
        """Process voice input from microphone"""
        r = sr.Recognizer()
        try:
            with sr.Microphone() as source:
                r.adjust_for_ambient_noise(source, duration=0.5)
                audio = r.listen(source, timeout=10, phrase_time_limit=10)
                try:
                    text = r.recognize_google(audio)
                    self.handle_command(text)
                except sr.UnknownValueError:
                    self.send_to_ui("Sorry, I didn't catch that.")
                    speak("I didn't understand. Please repeat.")
                except sr.RequestError:
                    self.send_to_ui("Voice service unavailable.")
                    speak("Connection error. Check your internet.")
        except Exception as e:
            self.send_to_ui(f"Microphone error: {str(e)}")
            speak("Microphone not available.")

    # ============= HARDWARE MONITORING =============
    def monitor_hardware(self, window):
        """Continuously monitor hardware"""
        time.sleep(7)
        while True:
            try:
                cpu = psutil.cpu_percent(interval=0.5)
                bat_obj = psutil.sensors_battery()
                bat = bat_obj.percent if bat_obj else 100
                window.evaluate_js(f"updateStats({cpu}, {bat})")
            except:
                pass
            time.sleep(5)

# ============= GLOBAL HOTKEY HANDLER =============
def setup_global_hotkey(api):
    """Setup global hotkey to activate Nela (Win+Shift+V)"""
    if not KEYBOARD_AVAILABLE:
        print("Global hotkey not available. Use UI button instead.")
        return
    
    def on_hotkey():
        if api.window:
            api.start_voice_cmd()
    
    try:
        keyboard.add_hotkey('win+shift+v', on_hotkey)
        print("Hotkey ready: Press Win+Shift+V to activate Nela")
    except:
        print("Could not register global hotkey")

# ============= MAIN APPLICATION =============
if __name__ == '__main__':
    print("🤖 Starting Nela...")
    print("🔧 Pre-initializing TTS engine...")
    threading.Thread(target=init_tts, daemon=True).start()
    time.sleep(1)  # Brief wait for TTS init
    
    print("🌐 Creating webview window...")
    api = NelaAPI()
    
    window = webview.create_window(
        'NELA - Windows AI Assistant',
        'nela_ui_v2.html',
        js_api=api,
        width=1200,
        height=800,
        background_color='#0b0e14'
    )
    api.window = window
    print("✅ Window created!")
    # quick speech test to confirm TTS is working
    speak("Nela is ready and initialized.")
    
    # Start hardware monitor ONLY (don't block on startup)
    print("📊 Starting hardware monitor...")
    threading.Thread(target=api.monitor_hardware, args=(window,), daemon=True).start()
    
    # Setup global hotkey (optional)
    if KEYBOARD_AVAILABLE:
        print("🎤 Setting up global hotkey...")
        threading.Thread(target=setup_global_hotkey, args=(api,), daemon=True).start()
    
    # Start wake-word listener if pocketsphinx is available
    if POCKETSPHINX_AVAILABLE:
        print("🧠 Starting wake-word listener...")
        threading.Thread(target=api.wake_listener, daemon=True).start()
    
    # Don't show welcome on startup to avoid blocking
    # Just let the HTML show the default welcome message
    print("▶️ Starting webview...")
    
    # Start webview without initial speech
    webview.start(debug=False)

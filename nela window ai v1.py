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
import re  # Added for brightness/volume numbers
from bs4 import BeautifulSoup
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# 1. THE CEO VOICE
def speak(text):
    def run_speak():
        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 185)
            engine.say(text)
            engine.runAndWait()
        except: pass
    threading.Thread(target=run_speak, daemon=True).start()

def clean_html(text):
    import re
    return re.sub(r'<[^>]+>', '', text)

# 2. THE OMEGA BRAIN
class NelaAPI:
    def __init__(self):
        self.window = None
        self.identity = "I am Nela. your AI assistant; is they any thing i can do for you today."

    def handle_command(self, query):
        query = query.lower().strip()
        if not query: return
        
        # --- THE THINKING PHASE ---
        self.send_to_ui("<b>Nela:</b> Processing your command...")
        # user query is logged, but not spoken back to avoid repeating the question
        time.sleep(1.2)

        # --- 1. IDENTITY & BOSS VIBE ---
        if any(word in query for word in ["who are you", "who own you", "hi", "assistant"]):
            self.send_to_ui(self.identity); speak(self.identity)
            return

        # --- 2. THE VISUAL SNAPSHOT ---
        elif any(word in query for word in ["snap", "screenshot", "capture"]):
            self.send_to_ui("<i>Capturing system visuals...</i>")
            speak("Capturing system visuals")
            path = os.path.expanduser("~/Desktop/Nela_Snapshot.png")
            pyautogui.screenshot(path)
            self.send_to_ui("Snapshot saved to your Desktop, Boss."); speak("Screen captured.")
            return

        # --- 3. HARDWARE OVERLORD (WiFi, Brightness, Lock) ---
        elif "wifi off" in query:
            os.system('netsh interface set interface "Wi-Fi" disabled')
            self.send_to_ui("Cutting the grid. WiFi is OFF."); speak("WiFi disabled.")
            return
        
        elif "wifi on" in query:
            os.system('netsh interface set interface "Wi-Fi" enabled')
            self.send_to_ui("Restoring the grid. WiFi is ON."); speak("WiFi enabled.")
            return

        elif "brightness" in query:
            # Added percent logic to your existing brightness
            nums = re.findall(r'\d+', query)
            lvl = int(nums[0]) if nums else (100 if "up" in query or "high" in query else 30)
            os.system(f'powershell (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{lvl})')
            self.send_to_ui(f"Brightness set to {lvl} percent."); speak(f"Brightness {lvl} percent.")
            return

        # --- NEW: VOLUME CONTROL ---
        elif "volume" in query or "mute" in query:
            try:
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                if "mute" in query:
                    volume.SetMute(1, None); self.send_to_ui("Volume neutralized."); speak("Muted.")
                else:
                    nums = re.findall(r'\d+', query)
                    vol_lvl = int(nums[0]) if nums else 50
                    volume.SetMasterVolumeLevelScalar(vol_lvl / 100, None)
                    self.send_to_ui(f"Volume set to {vol_lvl}%."); speak(f"Volume {vol_lvl} percent.")
            except: pass
            return

        elif "lock" in query:
            self.send_to_ui("Locking the Dell Iron..."); speak("System locked.")
            os.system("rundll32.exe user32.dll,LockWorkStation")
            return

        # --- 4. EXECUTIONER MODE (Kill Apps) ---
        elif "kill" in query or "close" in query:
            app = query.replace("kill", "").replace("close", "").strip()
            self.send_to_ui(f"Terminating {app} process..."); os.system(f"taskkill /f /im {app}.exe")
            speak(f"Process {app} neutralized.")
            return

        # --- 5. SYSTEM POWER (Shutdown, Restart, Hibernate) ---
        elif any(word in query for word in ["shutdown", "restart", "hibernate", "sleep"]):
            if "hibernate" in query or "sleep" in query:
                self.send_to_ui("Initiating Hibernate Sequence. Saving state..."); speak("Hibernating now.")
                time.sleep(2)
                os.system("shutdown /h")
            elif "restart" in query:
                self.send_to_ui("Restarting the Dell Iron..."); speak("Rebooting.")
                os.system("shutdown /r /t 5")
            else:
                self.send_to_ui("Executing Power-Down Sequence..."); speak("Goodbye, Senior Man.")
                os.system("shutdown /s /t 5")
            return

        # --- 6. PERCENT / BATTERY CHECK ---
        elif any(word in query for word in ["percent", "battery", "juice", "power"]):
            bat = psutil.sensors_battery()
            if bat:
                msg = f"Senior Man, juice is at {bat.percent}% and we are {'charging' if bat.power_plugged else 'on battery'}."
                self.send_to_ui(msg); speak(f"Battery is at {bat.percent} percent.")
            else:
                self.send_to_ui("I can't see the sensor, but the system is powered.")
                speak("System is powered.")
            return

        # --- 7. SMART FOLDER/APP LAUNCHER (Kept exactly as you wrote) ---
        elif "open" in query or "launch" in query:
            item = query.replace("open", "").replace("launch", "").strip()
            zones = [f"~/{item}", f"~/Desktop/{item}", f"~/Documents/{item}", f"~/Downloads/{item}"]
            found = False
            for z in zones:
                path = os.path.expanduser(z)
                if os.path.exists(path):
                    os.startfile(path); self.send_to_ui(f"Path located. Opening {item}."); speak(f"Opening {item}"); found = True; break
            if not found:
                os.system(f'start {item}')
                self.send_to_ui(f"Searching system registry for {item}...")
                speak(f"Searching for {item}")
            return

        # --- 8. GLOBAL INTELLIGENCE (Scraper) ---
        else:
            self.send_to_ui("<i>Thinking...</i>")
            speak("Thinking")
            threading.Thread(target=self.web_scrape, args=(query,), daemon=True).start()

    def web_scrape(self, query):
        try:
            # Stealth Headers to bypass blocks
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"}
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}&hl=en"
            res = requests.get(url, headers=headers, timeout=7)
            soup = BeautifulSoup(res.text, 'html.parser')
            ans = soup.find('div', class_='VwiC3b') or soup.find('span', class_='hgKElc') or soup.find('div', class_='Z0LcW')
            if ans:
                result = clean_html(ans.get_text())
                self.send_to_ui(f"<b>Extracted Data:</b> {result}")
                speak(result)
            else:
                # No quick answer element; try to grab a snippet from the first search result
                snippet = soup.find('div', class_='BNeawe vvjwJb AP7Wnd') or soup.find('div', class_='kCrYT')
                if snippet:
                    text = clean_html(snippet.get_text())
                    self.send_to_ui(f"<b>Search snippet:</b> {text}")
                    speak("Here is what I found")
                else:
                    # No data found; open Chrome to let user browse
                    self.send_to_ui("Opening Chrome to search...")
                    speak("Opening chrome to search.")
                    os.system(f"start {url}")
        except Exception:
            self.send_to_ui("Network Failure.")
            speak("Network failure.")

    def send_to_ui(self, text):
        if self.window:
            safe_text = json.dumps(text)
            try: self.window.evaluate_js(f"nelaResponse({safe_text})")
            except: pass

    # --- UPDATED VOICE COMMAND (Added 'Listening' status) ---
    def start_voice_cmd(self):
        self.send_to_ui("<b style='color:#00bcd4;'>Listening...</b>")
        threading.Thread(target=self.process_voice, daemon=True).start()

    def process_voice(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            try:
                audio = r.listen(source, timeout=5)
                text = r.recognize_google(audio)
                self.handle_command(text)
            except: 
                if self.window: self.window.evaluate_js("nelaResponse('No signal.')")
                speak("No signal.")

# 3. HARDWARE MONITOR
def monitor_hardware(window):
    time.sleep(7) # Wait for UI
    while True:
        try:
            cpu = psutil.cpu_percent()
            bat_obj = psutil.sensors_battery()
            bat = bat_obj.percent if bat_obj else 100
            window.evaluate_js(f"updateStats({cpu}, {bat})")
        except: pass
        time.sleep(5)

# 4. LAUNCHER
api = NelaAPI()
window = webview.create_window('NELA OMEGA CEO', 'nela_ui.html', js_api=api, width=1100, height=750)
api.window = window

if __name__ == '__main__':
    threading.Thread(target=monitor_hardware, args=(window,), daemon=True).start()
    # start with debug output to diagnose responsiveness issues
    webview.start(debug=True)

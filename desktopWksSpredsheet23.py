import time
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import subprocess

# --- Dependency Handling ---
def check_and_import_dependencies():
    missing_packages = []
    critical_packages = []
    
    # Critical dependencies that are required
    try: import pyautogui
    except ImportError: critical_packages.append("pyautogui")
    try: from pywinauto.application import Application
    except ImportError: critical_packages.append("pywinauto")
    try: from pyexcel_ods3 import get_data
    except ImportError: critical_packages.append("pyexcel-ods3")
    
    # Optional dependencies with fallbacks
    try: import numpy
    except ImportError: missing_packages.append("numpy")
    try: import sounddevice
    except ImportError: missing_packages.append("sounddevice")
    try: import soundfile
    except ImportError: missing_packages.append("soundfile")
    try: import vlc
    except ImportError: missing_packages.append("python-vlc")

    if critical_packages:
        msg = f"Critical packages missing:\n\n{', '.join(critical_packages)}\n\nThese are required. Please install them by running:\npip install {' '.join(critical_packages)}"
        messagebox.showerror("Critical Dependencies Missing", msg)
        return False
    
    if missing_packages:
        msg = f"Optional packages missing:\n\n{', '.join(missing_packages)}\n\nSome features may be limited. Install them by running:\npip install {' '.join(missing_packages)}"
        messagebox.showwarning("Optional Dependencies Missing", msg)
    return True

# Now that checks are done, import them for real
import pyautogui
from pywinauto.application import Application
from pyexcel_ods3 import get_data
import math

try: import numpy as np
except ImportError: np = None

try: import sounddevice as sd
except ImportError: sd = None

try: import soundfile as sf
except ImportError: sf = None

try: import winsound
except ImportError: winsound = None

try: import vlc
except ImportError: vlc = None

# --- Version 14.0 (Logging and Enhanced Audio Debugging) ---

# Fix audio playback by ensuring we always have a valid playback method
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()

# Try loading VLC module
try:
    import vlc
    VLC_AVAILABLE = True
except ImportError:
    vlc = None
    VLC_AVAILABLE = False

def create_silent_wav(filepath, duration=1):
    """Create a silent WAV file as fallback"""
    try:
        import wave
        import numpy as np
        
        # 1 second of silence
        sampleRate = 44100
        nSamples = int(duration * sampleRate)
        
        wav_file = wave.open(filepath, 'w')
        wav_file.setnchannels(1) # mono
        wav_file.setsampwidth(2) # 16-bit
        wav_file.setframerate(sampleRate)
        
        # Generate silent samples
        samples = np.zeros(nSamples).astype(np.int16)
        wav_file.writeframes(samples.tobytes())
        wav_file.close()
        return True
    except Exception:
        return False

# --- Configuration ---
LOOP_DELAY_SECONDS = 4
AUDIO_PLAY_DELAY_SECONDS = 10
MAX_AUDIO_DURATION_SECONDS = 40
BROWSER_WINDOW_TITLE = "phone-calls"
ATTEMPT_WINDOW_FOCUS = True

# --- Default Values ---
DEFAULT_TARGET_X = 1876
DEFAULT_TARGET_Y = 721  # Changed from 621 to 721 as requested
DEFAULT_PHONE_NUMBER = "3013461112"
DEFAULT_AUDIO_PATH = os.path.join(SCRIPT_DIR, "sample-3.mp3")
DEFAULT_HANGUP_X = 1800
DEFAULT_HANGUP_Y = 700

# --- Audio Player Configuration ---
AUDIO_BACKENDS = ['sounddevice', 'vlc', 'winsound']  # Priority order for audio playback
try:
    import vlc
    VLC_AVAILABLE = True
except ImportError:
    VLC_AVAILABLE = False

class AutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Desktop Automation UI v14.0")
        self.geometry("550x900")
        self.minsize(500, 800)
        self.configure(bg="#2E2E2E")
        self.attributes('-topmost', True)
        
        self.setup_styles()

        # --- Variables ---
        self.timer_text = tk.StringVar(value=f"{LOOP_DELAY_SECONDS}")
        self.status_text = tk.StringVar(value="Ready. Please test audio before starting.")
        self.cursor_pos_text = tk.StringVar(value="X: ---, Y: ---")
        self.target_x_var = tk.StringVar(value=str(DEFAULT_TARGET_X))
        self.target_y_var = tk.StringVar(value=str(DEFAULT_TARGET_Y))
        self.hangup_x_var = tk.StringVar(value=str(DEFAULT_HANGUP_X))
        self.hangup_y_var = tk.StringVar(value=str(DEFAULT_HANGUP_Y))
        self.current_row_text = tk.StringVar(value="Row: Not started")
        self.data_status_text = tk.StringVar(value="No data loaded")
        self.volume_level_var = tk.DoubleVar(value=0)
        self.audio_api_var = tk.StringVar()
        self.audio_device_var = tk.StringVar()
        self.call_timeout = tk.IntVar(value=30)
        self.auto_hangup = tk.BooleanVar(value=True)
        self.spreadsheet_path = tk.StringVar(value=os.path.join(SCRIPT_DIR, "data.ods"))
        self.volume_var = tk.DoubleVar(value=75)
        
        # Audio monitoring
        self.audio_file_status = tk.StringVar(value="No audio file loaded")
        self.audio_position = tk.StringVar(value="Time: --:--/--:--")
        self.current_player = None  # For VLC player instance
        self.audio_monitor_thread = None

        # --- State ---
        self.is_running = False
        self.audio_test_passed = False
        self.automation_thread = None
        self.spreadsheet_data = []
        self.current_row_index = 0
        self.stop_event = threading.Event()

        self.create_ui()
        self.log("Application initialized.")
        if np is None: self.log("Warning: NumPy not found, volume meter disabled.")
        if sd is None: self.log("Error: sounddevice not found, audio will not work.")
        if sf is None: self.log("Error: soundfile not found, audio will not work.")
        if winsound is None: self.log("Warning: winsound not found, beep test disabled.")

        self.periodic_ui_update()
        self.load_spreadsheet()
        self.populate_audio_settings()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.log("Closing application...")
        self.stop_automation()
        if sd: sd.stop()
        self.destroy()

    def setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#2E2E2E")
        style.configure("Header.TFrame", background="#1E1E1E", relief="raised", borderwidth=2)
        style.configure("TLabel", background="#2E2E2E", foreground="#FFFFFF", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"), foreground="#00D4FF")
        style.configure("Data.TLabel", font=("Segoe UI", 11), foreground="#00FF7F")
        style.configure("Status.TLabel", font=("Segoe UI", 10), foreground="#FFD700")
        style.configure("TEntry", font=("Segoe UI", 10), fieldbackground="#404040", foreground="#FFFFFF")
        style.map('TCombobox', fieldbackground=[('readonly', '#404040')], selectbackground=[('readonly', '#404040')], selectforeground=[('readonly', '#FFFFFF')])

    def create_ui(self):
        main_container = ttk.Frame(self, padding="10")
        main_container.grid(row=0, column=0, sticky="nsew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        header_frame = ttk.Frame(main_container, style="Header.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        ttk.Label(header_frame, text="🤖 Desktop Automation Control", style="Header.TLabel").pack(pady=10)

        self.create_status_section(main_container, 1)
        self.create_audio_test_section(main_container, 2)
        self.create_config_section(main_container, 3)
        self.create_control_section(main_container, 4)
        self.create_info_section(main_container, 5)
        self.create_log_section(main_container, 6)
        main_container.grid_rowconfigure(6, weight=1) # Make log section expandable

    def create_status_section(self, parent, row):
        status_frame = ttk.LabelFrame(parent, text="📊 Status Monitor", padding="10")
        status_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        info_items = [
            ("⏱ Timer:", self.timer_text), ("📍 Cursor:", self.cursor_pos_text),
            ("📋 Current Row:", self.current_row_text), ("💾 Data Status:", self.data_status_text)
        ]
        for i, (label_text, var) in enumerate(info_items):
            ttk.Label(status_frame, text=label_text).grid(row=i//2, column=(i%2)*2, sticky="w", padx=5, pady=2)
            ttk.Label(status_frame, textvariable=var, style="Data.TLabel").grid(row=i//2, column=(i%2)*2+1, sticky="w", padx=5, pady=2)
        ttk.Label(status_frame, text="Status:").grid(row=2, column=0, sticky="w", pady=(10,0))
        ttk.Label(status_frame, textvariable=self.status_text, style="Status.TLabel").grid(row=2, column=1, columnspan=3, sticky="w", padx=5, pady=(10,0))

    def create_audio_test_section(self, parent, row):
        audio_frame = ttk.LabelFrame(parent, text="🔊 Audio Test", padding="10")
        audio_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        audio_frame.grid_columnconfigure(1, weight=1)
        
        # Volume controls
        ttk.Label(audio_frame, text="Volume:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Scale(audio_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.volume_var).grid(row=0, column=1, columnspan=3, sticky="ew", padx=5)
        self.volume_meter = ttk.Progressbar(audio_frame, orient="horizontal", length=150, mode="determinate", variable=self.volume_level_var, maximum=100)
        self.volume_meter.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)
        if np is None: self.volume_meter['state'] = 'disabled'
        
        # Audio file status
        status_frame = ttk.Frame(audio_frame)
        status_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=5)
        ttk.Label(status_frame, textvariable=self.audio_file_status, style="Data.TLabel").pack(side=tk.LEFT, padx=5)
        ttk.Label(status_frame, textvariable=self.audio_position, style="Data.TLabel").pack(side=tk.RIGHT, padx=5)
        
        # Control buttons
        btn_frame = ttk.Frame(audio_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, sticky="ew")
        ttk.Button(btn_frame, text="Beep", command=self.play_beep).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Test Next Audio", command=self.test_next_audio).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        ttk.Button(btn_frame, text="Reset Audio", command=self.reset_audio).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.confirm_audio_btn = ttk.Button(btn_frame, text="Confirm Audio Works", command=self.confirm_audio)
        self.confirm_audio_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

    def create_config_section(self, parent, row):
        config_frame = ttk.LabelFrame(parent, text="⚙️ Configuration", padding="10")
        config_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        config_frame.grid_columnconfigure(1, weight=1)
        ttk.Button(config_frame, text="Reload Spreadsheet", command=self.select_spreadsheet).grid(row=0, column=0, columnspan=4, pady=(0,10), sticky="ew", padx=5)
        ttk.Label(config_frame, text="Target X:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(config_frame, textvariable=self.target_x_var, width=10).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(config_frame, text="Target Y:").grid(row=1, column=2, sticky="w", padx=(10,0))
        ttk.Entry(config_frame, textvariable=self.target_y_var, width=10).grid(row=1, column=3, sticky="w", padx=5)
        ttk.Label(config_frame, text="Hangup X:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(config_frame, textvariable=self.hangup_x_var, width=10).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Label(config_frame, text="Hangup Y:").grid(row=2, column=2, sticky="w", padx=(10,0))
        ttk.Entry(config_frame, textvariable=self.hangup_y_var, width=10).grid(row=2, column=3, sticky="w", padx=5)
        ttk.Label(config_frame, text="Audio API:").grid(row=3, column=0, sticky="w", pady=(10,2))
        self.audio_api_dropdown = ttk.Combobox(config_frame, textvariable=self.audio_api_var, state="readonly")
        self.audio_api_dropdown.grid(row=3, column=1, columnspan=3, sticky="ew", pady=(10,2), padx=5)
        self.audio_api_dropdown.bind("<<ComboboxSelected>>", self.refresh_audio_devices)
        ttk.Label(config_frame, text="Audio Output:").grid(row=4, column=0, sticky="w", pady=2)
        self.audio_device_dropdown = ttk.Combobox(config_frame, textvariable=self.audio_device_var, state="readonly")
        self.audio_device_dropdown.grid(row=4, column=1, columnspan=3, sticky="ew", pady=2, padx=5)
        ttk.Button(config_frame, text="⟳ Refresh Audio Devices", command=self.populate_audio_settings).grid(row=5, column=0, columnspan=4, pady=(5,0), sticky="ew", padx=5)
        f = ttk.Frame(config_frame); f.grid(row=6, column=0, columnspan=4, sticky='ew', pady=(10,0))
        ttk.Label(f, text="Call Timeout (s):").pack(side=tk.LEFT)
        ttk.Entry(f, textvariable=self.call_timeout, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(f, text="Auto Hangup", variable=self.auto_hangup).pack(side=tk.LEFT, padx=5)

    def create_control_section(self, parent, row):
        control_frame = ttk.LabelFrame(parent, text="🎮 Controls", padding="10")
        control_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        control_frame.grid_columnconfigure(0, weight=1); control_frame.grid_columnconfigure(1, weight=1); control_frame.grid_columnconfigure(2, weight=1)
        self.start_button = ttk.Button(control_frame, text="▶ START", command=self.start_automation, style="Start.TButton", state=tk.DISABLED)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0,5))
        self.stop_button = ttk.Button(control_frame, text="■ STOP", command=self.stop_automation, style="Stop.TButton", state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, sticky="ew", padx=5)
        self.reset_button = ttk.Button(control_frame, text="↻ RESET", command=self.reset_automation, style="Reset.TButton", state=tk.DISABLED)
        self.reset_button.grid(row=0, column=2, sticky="ew", padx=(5,0))

    def create_info_section(self, parent, row):
        info_frame = ttk.LabelFrame(parent, text="ℹ️ Information", padding="10")
        info_frame.grid(row=row, column=0, sticky="ew")
        self.info_label = ttk.Label(info_frame, text="", justify=tk.LEFT)
        self.info_label.pack(anchor=tk.W)
        self.update_info_label()

    def create_log_section(self, parent, row):
        log_frame = ttk.LabelFrame(parent, text="📜 Log", padding="10")
        log_frame.grid(row=row, column=0, sticky="nsew", pady=(10, 0))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, state=tk.DISABLED, bg="#1C1C1C", fg="#E0E0E0", font=("Consolas", 9), wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky='nsew')
        self.log_text['yscrollcommand'] = scrollbar.set

    def log(self, message, level="INFO"):
        timestamp = time.strftime("%H:%M:%S")
        full_message = f"[{timestamp}] [{level}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, full_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        print(full_message.strip())
        if level == "STATUS": self.status_text.set(message)

    def update_info_label(self):
        info_text = f"Spreadsheet: {os.path.basename(self.spreadsheet_path.get())}\nLoop Delay: {LOOP_DELAY_SECONDS}s, Audio Delay: {AUDIO_PLAY_DELAY_SECONDS}s, Max Audio: {MAX_AUDIO_DURATION_SECONDS}s"
        self.info_label.config(text=info_text)

    def select_spreadsheet(self):
        self.log("Opening file dialog to select spreadsheet.")
        filepath = filedialog.askopenfilename(title="Select Spreadsheet", filetypes=(("ODS files", "*.ods"), ("All files", "*.*")), initialdir=SCRIPT_DIR)
        if filepath:
            self.log(f"Spreadsheet selected: {filepath}")
            self.spreadsheet_path.set(filepath)
            self.update_info_label()
            self.load_spreadsheet()

    def load_spreadsheet(self):
        path = self.spreadsheet_path.get()
        self.log(f"Loading spreadsheet from: {path}")
        try:
            if not os.path.exists(path):
                raise FileNotFoundError(f"Spreadsheet not found at specified path.")
            data = get_data(path)
            sheet = list(next(iter(data.values())))  # Convert generator to list
            if len(sheet) <= 1: raise ValueError("No data rows found in spreadsheet.")
            self.spreadsheet_data = [{"phone": str(r[1]).strip(), "audio": str(r[2]).strip(), "row": i} for i, r in enumerate(sheet[1:], 2) if len(r) >= 3 and r[1] and r[2]]
            if not self.spreadsheet_data: raise ValueError("No valid, complete entries found.")
            self.data_status_text.set(f"✓ Loaded {len(self.spreadsheet_data)} entries")
            self.log(f"Successfully loaded {len(self.spreadsheet_data)} entries from spreadsheet.")
        except Exception as e:
            self.data_status_text.set(f"✗ Load failed")
            self.log(f"Failed to load spreadsheet: {e}", "ERROR")
        self.current_row_index = 0

    def start_automation(self):
        self.log("Start button pressed.")
        if self.is_running: self.log("Automation is already running.", "WARN"); return
        if not self.audio_test_passed:
            self.log("Audio test not passed. Cannot start.", "WARN")
            messagebox.showwarning("Audio Test Required", "Please confirm the audio test is working before starting.")
            return

        self.is_running = True
        self.stop_event.clear()
        self.start_button.config(state=tk.DISABLED); self.stop_button.config(state=tk.NORMAL)
        self.log("🚀 Starting automation...", "STATUS")
        self.automation_thread = threading.Thread(target=self.run_automation_loop, daemon=True)
        self.automation_thread.start()

    def stop_automation(self):
        if not self.is_running: return
        self.log("Stopping automation...")
        self.is_running = False
        self.stop_event.set()
        if sd: sd.stop()
        self.start_button.config(state=tk.NORMAL if self.audio_test_passed else tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.log("⏹ Automation stopped", "STATUS")

    def reset_automation(self):
        self.log("Resetting automation state.")
        if self.is_running: self.stop_automation(); time.sleep(0.1)
        self.current_row_index = 0
        self.current_row_text.set("Row: Not started")
        self.log("↻ Reset complete", "STATUS")
        self.load_spreadsheet()

    def periodic_ui_update(self):
        try: self.cursor_pos_text.set(f"X: {pyautogui.position().x}, Y: {pyautogui.position().y}")
        except: pass
        self.after(100, self.periodic_ui_update)

    def run_automation_loop(self):
        while self.is_running:
            if self.stop_event.is_set(): break
            try:
                phone, audio, _ = self.get_current_data()
                tx, ty = int(self.target_x_var.get()), int(self.target_y_var.get())

                if ATTEMPT_WINDOW_FOCUS and not self.focus_target_window():
                    if not self.responsive_wait(3): break; continue
                
                self.perform_dialing_sequence(phone, tx, ty)
                if self.stop_event.is_set(): break

                self.log(f"Waiting {AUDIO_PLAY_DELAY_SECONDS}s before playing audio...", "STATUS")
                if not self.responsive_wait(AUDIO_PLAY_DELAY_SECONDS): break

                self.play_audio_file(audio)
                if self.stop_event.is_set(): break

                if self.auto_hangup.get():
                    self.log(f"Waiting {self.call_timeout.get()}s before auto-hangup...", "STATUS")
                    if not self.responsive_wait(self.call_timeout.get()): break
                    self.perform_hangup()

                self.current_row_index = (self.current_row_index + 1) % len(self.spreadsheet_data)
                self.log(f"Next cycle in {LOOP_DELAY_SECONDS}s...", "STATUS")
                if not self.responsive_wait(LOOP_DELAY_SECONDS): break

            except Exception as e:
                self.log(f"Error in automation loop: {e}", "ERROR")
                if not self.responsive_wait(3): break
        self.after(0, self.stop_automation)

    def get_current_data(self):
        if not self.spreadsheet_data: return DEFAULT_PHONE_NUMBER, DEFAULT_AUDIO_PATH, True
        entry = self.spreadsheet_data[self.current_row_index]
        self.current_row_text.set(f"Row {entry['row']}")
        return entry["phone"], entry["audio"], False

    def focus_target_window(self):
        self.log(f"Finding window: {BROWSER_WINDOW_TITLE}", "STATUS")
        try:
            Application(backend="win32").connect(title_re=f".*{BROWSER_WINDOW_TITLE}.*", timeout=5).top_window().set_focus()
            self.log("✓ Window focused")
            return True
        except Exception as e:
            self.log(f"Window focus failed: {e}", "ERROR"); return False

    def perform_dialing_sequence(self, phone, x, y):
        self.log(f"Dialing {phone} at ({x},{y})", "STATUS")
        pyautogui.moveTo(x, y, duration=0.2); pyautogui.click()
        pyautogui.hotkey('ctrl', 'a'); pyautogui.press('delete')
        pyautogui.write(phone, interval=0.05); pyautogui.press('enter')

    def perform_hangup(self):
        try: hx, hy = int(self.hangup_x_var.get()), int(self.hangup_y_var.get())
        except ValueError: self.log("Invalid hangup coordinates.", "ERROR"); return
        self.log(f"Hanging up at ({hx},{hy})", "STATUS")
        pyautogui.moveTo(hx, hy, duration=0.2); pyautogui.click()

    def update_audio_monitor(self, audio_path):
        """Update audio file status and start position monitoring if using VLC"""
        try:
            if os.path.exists(audio_path):
                filename = os.path.basename(audio_path)
                filesize = os.path.getsize(audio_path) / (1024*1024)  # MB
                self.audio_file_status.set(f"File: {filename} ({filesize:.1f}MB)")
                
                # If using VLC, start position monitoring
                if self.current_player and VLC_AVAILABLE:
                    if self.audio_monitor_thread and self.audio_monitor_thread.is_alive():
                        self.stop_event.set()
                        self.audio_monitor_thread.join(timeout=1)
                    
                    def monitor_position():
                        while not self.stop_event.is_set() and self.current_player:
                            try:
                                if self.current_player.is_playing():
                                    pos_ms = self.current_player.get_time()
                                    length_ms = self.current_player.get_length()
                                    pos_str = time.strftime('%M:%S', time.gmtime(pos_ms/1000))
                                    length_str = time.strftime('%M:%S', time.gmtime(length_ms/1000))
                                    self.audio_position.set(f"Time: {pos_str}/{length_str}")
                                time.sleep(0.1)
                            except Exception:
                                break
                        self.audio_position.set("Time: --:--/--:--")
                    
                    self.audio_monitor_thread = threading.Thread(target=monitor_position, daemon=True)
                    self.audio_monitor_thread.start()
            else:
                self.audio_file_status.set("File not found!")
                self.audio_position.set("Time: --:--/--:--")
        except Exception as e:
            self.log(f"Error updating audio monitor: {e}", "ERROR")

    def play_audio_file(self, audio_path):
        """Enhanced audio playback with multiple backends and monitoring"""
        self.log(f"Attempting to play audio: {os.path.basename(audio_path)}")
        if not os.path.exists(audio_path):
            self.log(f"Audio file not found: {audio_path}", "ERROR")
            return

        self.update_audio_monitor(audio_path)
        volume = self.volume_var.get() / 100.0

        # Try each backend in order
        if sd and sf:  # Try sounddevice first
            try:
                device = self.audio_device_var.get()
                self.log(f"Trying sounddevice backend with device: '{device}'")
                
                # Find exact device index
                device_index = None
                devices = sd.query_devices()
                for i, dev in enumerate(devices):
                    try:
                        if isinstance(dev, dict):
                            name = dev.get('name', '')
                        else:
                            name = str(dev)
                        if device and name and device.lower() in name.lower():
                            device_index = i
                            break
                    except Exception:
                        continue
                
                if device_index is not None:
                    with sf.SoundFile(audio_path, 'r') as f:
                        data = f.read(dtype='float32')
                        data *= volume
                        sd.play(data, samplerate=f.samplerate, device=device_index, blocking=True)
                        self.log("✓ Audio playback completed via sounddevice")
                        return True
            except Exception as e:
                self.log(f"sounddevice playback failed: {e}", "ERROR")

        if VLC_AVAILABLE and vlc:  # Try VLC
            try:
                self.log("Trying VLC backend")
                # Create a new instance for each playback to avoid state issues
                instance = vlc.Instance('--quiet')
                if instance:
                    player = instance.media_player_new()
                    media = instance.media_new(audio_path)
                    if player and media:
                        player.set_media(media)
                        player.audio_set_volume(int(volume * 100))
                        self.current_player = player
                        player.play()
                        time.sleep(0.1)  # Let it start
                        while player.is_playing():
                            time.sleep(0.1)
                        self.log("✓ Audio playback completed via VLC")
                        return True
            except Exception as e:
                self.log(f"VLC playback failed: {e}", "ERROR")

        # Try winsound for WAV files
        if winsound and audio_path.lower().endswith('.wav'):
            try:
                self.log("Trying winsound backend")
                winsound.PlaySound(audio_path, winsound.SND_FILENAME)
                self.log("✓ Audio playback completed via winsound")
                return True
            except Exception as e:
                self.log(f"winsound playback failed: {e}", "ERROR")

        # Last resort: system default player
        try:
            self.log("Trying system default player")
            if sys.platform == "win32":
                os.startfile(audio_path)
            else:
                import subprocess
                subprocess.Popen(['xdg-open', audio_path])
            self.log("✓ Launched in system default player")
            return True
        except Exception as e:
            self.log(f"System player failed: {e}", "ERROR")
            return False

    def test_next_audio(self):
        self.log("Audio test button pressed.")
        _, audio, is_default = self.get_current_data()
        if is_default: self.log("No spreadsheet loaded. Testing default audio file.")
        threading.Thread(target=self.play_audio_file, args=(audio,), daemon=True).start()

    def reset_audio(self):
        self.log("Audio system reset.")
        if sd: sd.stop()
        self.audio_test_passed = False
        self.confirm_audio_btn.config(state=tk.NORMAL)
        self.start_button.config(state=tk.DISABLED)
        self.reset_button.config(state=tk.DISABLED)
        self.log("Audio test un-confirmed. Main controls disabled.", "WARN")

    def confirm_audio(self):
        self.log("Audio test confirmed by user.")
        self.audio_test_passed = True
        self.log("✓ Audio test passed. Controls enabled.", "STATUS")
        self.start_button.config(state=tk.NORMAL)
        self.reset_button.config(state=tk.NORMAL)
        self.confirm_audio_btn.config(state=tk.DISABLED)

    def play_beep(self):
        self.log("Attempting to play system beep.")
        if winsound:
            try:
                winsound.Beep(440, 500) # Frequency 440Hz, Duration 500ms
                self.log("✓ System beep played successfully.")
            except Exception as e:
                self.log(f"Failed to play system beep: {e}", "ERROR")
                messagebox.showerror("Beep Error", f"Could not play beep sound.\nError: {e}")
        else:
            self.log("winsound module not available on this system.", "ERROR")
            messagebox.showwarning("Beep Unavailable", "The winsound module is not available on your system.")

    def responsive_wait(self, duration):
        end_time = time.time() + duration
        while time.time() < end_time:
            if self.stop_event.is_set(): return False
            time.sleep(0.1)
        return True

    def populate_audio_settings(self):
        self.log("Populating audio settings...")
        if not sd: self.log("sounddevice not found, cannot populate audio settings.", "ERROR"); return
        try:
            apis = sd.query_hostapis()
            # Convert any list/dict hostapis to a list of names
            api_names = []
            for api in apis:
                try:
                    if isinstance(api, dict):
                        name = api.get('name', '')
                    elif isinstance(api, (list, tuple)):
                        name = api[0] if api else ''
                    else:
                        name = str(api)
                    if name:
                        api_names.append(name)
                except Exception:
                    continue
            
            self.audio_api_dropdown['values'] = api_names
            if api_names:
                try:
                    default_api = sd.query_hostapis(sd.default.hostapi)
                    default_name = default_api.get('name') if isinstance(default_api, dict) else str(default_api)
                    self.audio_api_dropdown.set(default_name)
                except Exception:
                    self.audio_api_dropdown.set(api_names[0])
            self.log(f"Found {len(api_names)} audio APIs.")
            self.refresh_audio_devices()
        except Exception as e:
            self.log(f"Error populating audio APIs: {e}", "ERROR")

    def refresh_audio_devices(self, event=None):
        if not sd: return
        try:
            api_name = self.audio_api_var.get()
            self.log(f"Refreshing devices for API: {api_name}")
            
            # Find API index by name
            api_index = None
            for i, api in enumerate(sd.query_hostapis()):
                try:
                    name = api.get('name') if isinstance(api, dict) else str(api)
                    if name and api_name and api_name.lower() in name.lower():
                        api_index = i
                        break
                except Exception:
                    continue
            
            if api_index is None:
                self.log(f"Could not find API: {api_name}", "ERROR")
                return
                
            # Get devices for this API
            devices = []
            for d in sd.query_devices():
                try:
                    if isinstance(d, dict):
                        if (d.get('hostapi') == api_index and 
                            d.get('max_output_channels', 0) > 0 and
                            'name' in d):
                            devices.append(d['name'])
                    else:
                        devices.append(str(d))
                except Exception:
                    continue
            
            self.audio_device_dropdown['values'] = devices or ['No Devices']
            self.log(f"Found {len(devices)} output devices.")
            
            if devices:
                try:
                    default_dev = sd.query_devices(sd.default.device[1])
                    default_name = default_dev.get('name') if isinstance(default_dev, dict) else str(default_dev)
                    if default_name in devices:
                        self.audio_device_dropdown.set(default_name)
                    else:
                        self.audio_device_dropdown.set(devices[0])
                except Exception:
                    self.audio_device_dropdown.set(devices[0])
        except Exception as e:
            self.log(f"Error refreshing audio devices: {e}", "ERROR")

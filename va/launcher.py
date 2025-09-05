import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import queue
import time
import sys
import os
import contextlib
import io

# Import your assistant module
try:
    import main as assistant
    import speech_recognition as sr
except ImportError as e:
    print(f"Error: Could not import required modules: {e}")
    sys.exit(1)

class VoiceAssistantGUI:
    def __init__(self, root):
        self.root = root
        self.setup_gui()
        self.message_queue = queue.Queue()
        self.assistant_thread = None
        self.is_running = False
        
        # Create a custom speech recognition class that integrates with GUI
        self.setup_speech_recognition()
        
        # Start checking for messages
        self.check_messages()
        
    def setup_gui(self):
        self.root.title("🤖 Voice Desktop Assistant")
        self.root.geometry("600x700+100+100")
        self.root.resizable(True, True)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = tk.Label(main_frame, text="🎙️ Voice Desktop Assistant", 
                              font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(1, weight=1)
        
        # Status indicator
        self.status_label = tk.Label(status_frame, text="Status:", font=("Arial", 10))
        self.status_label.grid(row=0, column=0, padx=(0, 5))
        
        self.status_indicator = tk.Label(status_frame, text="🔴 Idle", 
                                        font=("Arial", 10, "bold"))
        self.status_indicator.grid(row=0, column=1, sticky=tk.W)
        
        # Listening indicator
        self.listening_label = tk.Label(status_frame, text="", 
                                       font=("Arial", 10), fg="blue")
        self.listening_label.grid(row=1, column=0, columnspan=2, sticky=tk.W)
        
        # Current input display
        input_frame = ttk.LabelFrame(main_frame, text="Current Voice Input", padding="5")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        self.current_input = tk.Label(input_frame, text="", font=("Arial", 11, "italic"),
                                     fg="blue", wraplength=500, justify=tk.LEFT)
        self.current_input.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Log display
        log_frame = ttk.LabelFrame(main_frame, text="Assistant Log", padding="5")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_display = scrolledtext.ScrolledText(log_frame, height=20, width=70,
                                                    font=("Consolas", 9),
                                                    wrap=tk.WORD)
        self.log_display.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Control buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(0, 5))
        
        self.start_button = ttk.Button(button_frame, text="🎤 Start Assistant", 
                                      command=self.start_assistant)
        self.start_button.grid(row=0, column=0, padx=(0, 5))
        
        self.stop_button = ttk.Button(button_frame, text="⏹️ Stop Assistant", 
                                     command=self.stop_assistant, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=5)
        
        self.clear_button = ttk.Button(button_frame, text="🗑️ Clear Log", 
                                      command=self.clear_log)
        self.clear_button.grid(row=0, column=2, padx=(5, 0))
        
        # Initial log message
        self.add_log("GUI initialized. Click 'Start Assistant' to begin.", "INFO")
        
    def add_log(self, message, msg_type="INFO"):
        """Add a message to the log display"""
        timestamp = time.strftime("%H:%M:%S")
        
        # Color coding for different message types
        colors = {
            "INFO": "black",
            "SUCCESS": "green",
            "ERROR": "red",
            "WARNING": "orange",
            "USER": "blue",
            "ASSISTANT": "purple",
            "SYSTEM": "gray"
        }
        
        # Insert message
        formatted_msg = f"[{timestamp}] [{msg_type}] {message}\n"
        self.log_display.insert(tk.END, formatted_msg)
        
        # Apply color to the last line
        last_line_start = self.log_display.index("end-2c linestart")
        last_line_end = self.log_display.index("end-1c")
        
        tag_name = f"{msg_type}_{int(time.time() * 1000)}"
        self.log_display.tag_add(tag_name, last_line_start, last_line_end)
        self.log_display.tag_configure(tag_name, foreground=colors.get(msg_type, "black"))
        
        # Auto-scroll to bottom
        self.log_display.see(tk.END)
        
    def update_status(self, status, color="black"):
        """Update the status indicator"""
        status_colors = {
            "idle": ("🔴 Idle", "red"),
            "running": ("🟢 Running", "green"),
            "listening": ("🎤 Listening...", "blue"),
            "processing": ("⚙️ Processing...", "orange"),
            "executing": ("🔧 Executing...", "purple"),
            "error": ("❌ Error", "red")
        }
        
        if status.lower() in status_colors:
            text, color = status_colors[status.lower()]
            self.status_indicator.config(text=text, fg=color)
        else:
            self.status_indicator.config(text=status, fg=color)
            
    def update_listening_status(self, message):
        """Update the listening status message"""
        self.listening_label.config(text=message)
        
    def update_current_input(self, text):
        """Update the current input display"""
        self.current_input.config(text=text)
        
    def start_assistant(self):
        """Start the voice assistant in a separate thread"""
        if self.is_running:
            return
            
        self.is_running = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        self.update_status("running")
        self.add_log("Starting voice assistant...", "INFO")
        
        # Start assistant thread
        self.assistant_thread = threading.Thread(target=self.run_assistant_with_gui_integration, daemon=True)
        self.assistant_thread.start()
        
    def stop_assistant(self):
        """Stop the voice assistant"""
        if not self.is_running:
            return
            
        self.is_running = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        self.update_status("idle")
        self.update_listening_status("")
        self.update_current_input("")
        self.add_log("Voice assistant stopped.", "WARNING")
        
    def clear_log(self):
        """Clear the log display"""
        self.log_display.delete(1.0, tk.END)
        self.add_log("Log cleared.", "INFO")
        
    def setup_speech_recognition(self):
        """Set up custom speech recognition that integrates with GUI"""
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 4000
        self.recognizer.dynamic_energy_threshold = True
        
    def gui_listen_and_recognize(self):
        """Custom listening function that provides GUI feedback"""
        try:
            # Update GUI status
            self.message_queue.put(("status", "listening"))
            self.message_queue.put(("listening", "🎤 Listening for voice input..."))
            self.message_queue.put(("log", "🎤 Listening for voice input...", "SYSTEM"))
            
            # Try noise-reduced recording first
            audio_arr, sr_hz = assistant.record_with_noise_suppression(duration=5)
            
            if audio_arr is not None:
                audio_clean = sr.AudioData(audio_arr.tobytes(), sample_rate=sr_hz, sample_width=2)
                self.message_queue.put(("listening", "🔄 Processing audio..."))
                self.message_queue.put(("status", "processing"))
                
                command = self.recognizer.recognize_google(audio_clean)
            else:
                # Fallback to standard microphone
                with sr.Microphone() as source:
                    self.message_queue.put(("listening", "🎤 Listening (fallback mode)..."))
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                    audio = self.recognizer.listen(source, phrase_time_limit=5)
                
                self.message_queue.put(("listening", "🔄 Processing audio..."))
                self.message_queue.put(("status", "processing"))
                command = self.recognizer.recognize_google(audio)

            # Update GUI with recognized command
            self.message_queue.put(("current_input", command))
            self.message_queue.put(("log", f"You said: {command}", "USER"))
            self.message_queue.put(("listening", "✅ Voice input received"))
            
            return command
            
        except sr.UnknownValueError:
            self.message_queue.put(("log", "Could not understand audio", "ERROR"))
            self.message_queue.put(("listening", "❌ Could not understand"))
            self.message_queue.put(("status", "error"))
            raise
        except Exception as e:
            self.message_queue.put(("log", f"Speech recognition error: {str(e)}", "ERROR"))
            self.message_queue.put(("listening", f"❌ Error: {str(e)[:30]}"))
            self.message_queue.put(("status", "error"))
            raise
            
    def run_assistant_with_gui_integration(self):
        """Modified version of the assistant main loop with GUI integration"""
        
        # Patch assistant functions for GUI feedback
        original_speak = assistant.speak
        
        def gui_speak(text):
            self.message_queue.put(("log", f"Assistant: {text}", "ASSISTANT"))
            original_speak(text)
            
        assistant.speak = gui_speak
        
        # Main loop similar to assistant's listen_and_execute_loop
        consecutive_errors = 0
        max_errors = 3
        
        self.message_queue.put(("log", "Voice assistant ready! Say a command...", "SUCCESS"))
        
        while self.is_running:
            try:
                # Use our custom listening function
                command = self.gui_listen_and_recognize()
                consecutive_errors = 0  # Reset error counter
                
                self.message_queue.put(("status", "processing"))
                self.message_queue.put(("listening", "🤖 Processing command..."))
                
                # Parse command using assistant's parser
                result = assistant.parse_command_online(command)
                if not result or result.get("action") == "unknown":
                    result = assistant.fallback_parser(command)
                
                self.message_queue.put(("log", f"Parsed: {result}", "SYSTEM"))
                
                # Execute command
                if result.get("action") != "unknown":
                    self.message_queue.put(("status", "executing"))
                    self.message_queue.put(("listening", "⚙️ Executing command..."))
                    
                    assistant.execute(result.get("action"), result.get("params", {}))
                    
                    self.message_queue.put(("status", "running"))
                    self.message_queue.put(("listening", "✅ Command completed"))
                    self.message_queue.put(("current_input", ""))  # Clear current input
                else:
                    self.message_queue.put(("log", "Command not recognized", "WARNING"))
                    self.message_queue.put(("status", "running"))
                    self.message_queue.put(("listening", "❓ Command not recognized"))
                    
                # Small delay before next listening cycle
                time.sleep(1)
                
            except sr.UnknownValueError:
                consecutive_errors += 1
                if consecutive_errors < max_errors:
                    self.message_queue.put(("log", "Please repeat that.", "WARNING"))
                else:
                    self.message_queue.put(("log", "Having trouble hearing. Check microphone.", "ERROR"))
                    consecutive_errors = 0
                    
            except Exception as e:
                consecutive_errors += 1
                error_msg = f"Error: {str(e)[:50]}"
                self.message_queue.put(("log", error_msg, "ERROR"))
                
                if consecutive_errors >= max_errors:
                    self.message_queue.put(("log", "Multiple errors. Restarting listener.", "ERROR"))
                    consecutive_errors = 0
                    time.sleep(2)  # Brief pause before retrying
                    
        # Assistant stopped
        self.message_queue.put(("status", "idle"))
        self.message_queue.put(("listening", ""))
        self.message_queue.put(("current_input", ""))
            
    def check_messages(self):
        """Check for messages from the assistant thread"""
        try:
            while True:
                message_type, *args = self.message_queue.get_nowait()
                
                if message_type == "log":
                    message, msg_type = args
                    self.add_log(message, msg_type)
                elif message_type == "status":
                    status = args[0]
                    self.update_status(status)
                elif message_type == "listening":
                    message = args[0]
                    self.update_listening_status(message)
                elif message_type == "current_input":
                    text = args[0]
                    self.update_current_input(text)
                    
        except queue.Empty:
            pass
        
        # Continue checking every 100ms
        self.root.after(100, self.check_messages)

def main():
    root = tk.Tk()
    app = VoiceAssistantGUI(root)
    
    def on_closing():
        if app.is_running:
            app.stop_assistant()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
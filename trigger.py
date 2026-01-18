import socket
import keyboard
import csv
import threading
import os
import time
from tkinter import (Tk, Label, Button, Frame, Listbox, Scrollbar, Entry,
                     StringVar, BooleanVar, messagebox, filedialog, Toplevel, END, HORIZONTAL)
from tkinter import ttk

try:
    import pygame
except ImportError:
    print("\n--- WARNING: Pygame library not found! ---")
    print("Audio playback will be disabled. To enable, run: pip install pygame\n")
    pygame = None


# Unicorn Configuration
UDP_IP = "127.0.0.1"
UDP_PORT = 1000

EVENT_TRIGGERS = {
    'n': ('1', 'Song Start'), 'm': ('2', 'Song or Playlist End'),
    'v': ('9', 'Transition Start'), 'b': ('10', 'Transition End'),
    'q': ('3', 'Visual Event Start'), 'w': ('4', 'Visual Event End'),
    'a': ('5', 'Auditory Event Start'), 's': ('6', 'Auditory Event End'),
    'y': ('7', 'Body Movement Start'), 'x': ('8', 'Body Movement End'),
}

class TriggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unicorn Trigger Control & Playlist Automation")
        self.root.minsize(650, 700)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.playlist = []
        self.audio_file_path = None
        self.is_automation_running = False
        self.is_paused = False
        self.pause_time_ms = 0
        self.scheduled_events = []
        self.all_automation_events = []
        self.total_mix_duration = 0
        self.playlist_editor_window = None
        self.csv_editor_data = []
        self.editing_song_index = None

        self.manual_timing_mode = BooleanVar(value=False)
        self.start_time = 0
        self.pause_start_time = 0
        self.total_pause_duration = 0

        # Networking & Audio
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.endPoint = (UDP_IP, UDP_PORT)
        if pygame: pygame.mixer.init()

        self.setup_ui()
        self.setup_keyboard_hooks()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.log_event("Application initialized. Ready for playlist files.")

    def format_time_long(self, seconds):
        """Converts seconds into HH:MM:SS string format for long durations."""
        s = int(seconds)
        h, s = divmod(s, 3600)
        m, s = divmod(s, 60)
        return f"{h:02}:{m:02}:{s:02}"

    def format_time_short(self, seconds):
        """Converts seconds into MM:SS string format for song durations."""
        s = int(seconds)
        m, s = divmod(s, 60)
        return f"{m:02}:{s:02}"

    def parse_time_to_seconds(self, time_str):
        """Converts MM:SS or SS string format into seconds."""
        parts = str(time_str).split(':')
        parts.reverse() 
        total_seconds = 0
        try:
            if len(parts) > 0: # Seconds
                total_seconds += int(parts[0])
            if len(parts) > 1: # Minutes
                total_seconds += int(parts[1]) * 60
            return total_seconds
        except (ValueError, IndexError):
            return None

    def setup_ui(self):
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TButton", padding=6)
        style.configure("TLabel", background="#f0f0f0", font=('Helvetica', 10))
        
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        top_frame = ttk.Frame(main_frame, padding="10")
        top_frame.grid(row=0, column=0, sticky='ew')
        top_frame.grid_columnconfigure(0, weight=1)
        
        middle_frame = ttk.Frame(main_frame, padding="10")
        middle_frame.grid(row=1, column=0, sticky='nsew')
        middle_frame.grid_columnconfigure(0, weight=1)
        middle_frame.grid_rowconfigure(0, weight=1)
        middle_frame.grid_rowconfigure(1, weight=1)

        bottom_frame = ttk.Frame(main_frame, padding="10")
        bottom_frame.grid(row=2, column=0, sticky='ew')
        bottom_frame.grid_columnconfigure(0, weight=1)

        automation_frame = ttk.LabelFrame(top_frame, text="Automation Setup", padding="10")
        automation_frame.pack(fill='x', expand=True)
        automation_frame.grid_columnconfigure(0, weight=1)

        self.btn_load_audio = ttk.Button(automation_frame, text="Load Audio File", command=self.browse_for_audio)
        self.btn_load_audio.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        self.audio_path_var = StringVar(value="Audio File (mp3): Not loaded")
        ttk.Label(automation_frame, textvariable=self.audio_path_var, font=('Helvetica', 8, 'italic')).grid(row=1, column=0, columnspan=2, sticky='w', padx=5)
        
        self.manual_mode_check = ttk.Checkbutton(
            automation_frame, 
            text="or use manual Synchronizing-Mode [no audio file]", 
            variable=self.manual_timing_mode,
            command=self.toggle_timing_mode
        )
        self.manual_mode_check.grid(row=2, column=0, columnspan=2, sticky='w', padx=5, pady=(0, 10))

        self.btn_load_playlist = ttk.Button(automation_frame, text="Load Playlist Information", command=self.browse_for_playlist)
        self.btn_load_playlist.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        self.btn_create_playlist = ttk.Button(automation_frame, text="...or Create/Edit Playlist Information File", command=self.open_playlist_editor)
        self.btn_create_playlist.grid(row=4, column=0, columnspan=2, padx=5, pady=2, sticky='ew')

        self.playlist_path_var = StringVar(value="Playlist (CSV): Not loaded")
        ttk.Label(automation_frame, textvariable=self.playlist_path_var, font=('Helvetica', 8, 'italic')).grid(row=5, column=0, columnspan=2, sticky='w', padx=5)
        
        transition_frame = ttk.Frame(automation_frame)
        transition_frame.grid(row=6, column=0, columnspan=2, sticky='w', padx=5, pady=(15, 5))
        ttk.Label(transition_frame, text="Transition Time (s):").pack(side='left')
        self.transition_time_var = StringVar(value="15")
        ttk.Entry(transition_frame, textvariable=self.transition_time_var, width=8).pack(side='left', padx=5)
        
        control_frame = ttk.Frame(automation_frame)
        control_frame.grid(row=7, column=0, columnspan=2, pady=10)
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_columnconfigure(1, weight=1)
        control_frame.grid_columnconfigure(2, weight=1)
        self.btn_start_automation = ttk.Button(control_frame, text="START", command=self.start_automation, state='disabled')
        self.btn_start_automation.grid(row=0, column=0, padx=5, sticky='ew')
        self.btn_pause_automation = ttk.Button(control_frame, text="PAUSE", command=self.pause_automation, state='disabled')
        self.btn_pause_automation.grid(row=0, column=1, padx=5, sticky='ew')
        self.btn_stop_automation = ttk.Button(control_frame, text="STOP", command=self.stop_automation, state='disabled')
        self.btn_stop_automation.grid(row=0, column=2, padx=5, sticky='ew')

        status_frame = ttk.LabelFrame(middle_frame, text="Status", padding="10")
        status_frame.grid(row=0, column=0, sticky='nsew', pady=(0,5))
        status_frame.grid_columnconfigure(0, weight=1)
        status_frame.grid_rowconfigure(4, weight=1)

        self.now_playing_var = StringVar(value="Now Playing: -")
        ttk.Label(status_frame, textvariable=self.now_playing_var, font=('Helvetica', 10, 'italic')).grid(row=0, column=0, sticky='w')
        self.time_display_var = StringVar(value="00:00:00 / 00:00:00")
        ttk.Label(status_frame, textvariable=self.time_display_var, font=('Consolas', 12, 'bold')).grid(row=1, column=0, sticky='w', pady=5)
        self.last_trigger_var = StringVar(value="Last Trigger: -")
        ttk.Label(status_frame, textvariable=self.last_trigger_var).grid(row=2, column=0, sticky='w', pady=(0, 10))
        
        ttk.Label(status_frame, text="Playlist Information:", font=('Helvetica', 10, 'bold')).grid(row=3, column=0, sticky='w', pady=(10, 2))

        playlist_scroll_frame = ttk.Frame(status_frame)
        playlist_scroll_frame.grid(row=4, column=0, sticky='nsew')
        playlist_scroll_frame.grid_columnconfigure(0, weight=1)
        playlist_scroll_frame.grid_rowconfigure(0, weight=1)
        
        self.playlist_list_box = Listbox(playlist_scroll_frame, font=('Consolas', 10), height=8)
        self.playlist_list_box.grid(row=0, column=0, sticky='nsew')
        
        scrollbar_playlist_y = Scrollbar(playlist_scroll_frame, orient="vertical", command=self.playlist_list_box.yview)
        scrollbar_playlist_y.grid(row=0, column=1, sticky='ns')
        scrollbar_playlist_x = Scrollbar(playlist_scroll_frame, orient=HORIZONTAL, command=self.playlist_list_box.xview)
        scrollbar_playlist_x.grid(row=1, column=0, sticky='ew')
        self.playlist_list_box.config(yscrollcommand=scrollbar_playlist_y.set, xscrollcommand=scrollbar_playlist_x.set)
        
        log_frame = ttk.LabelFrame(middle_frame, text="Live Trigger Log", padding="10")
        log_frame.grid(row=1, column=0, sticky='nsew', pady=(5, 0))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        log_scroll_frame = ttk.Frame(log_frame)
        log_scroll_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')
        log_scroll_frame.grid_columnconfigure(0, weight=1)
        log_scroll_frame.grid_rowconfigure(0, weight=1)

        self.log_box = Listbox(log_scroll_frame, font=('Consolas', 9), fg="#333333")
        self.log_box.grid(row=0, column=0, sticky='nsew')
        
        scrollbar_log_y = Scrollbar(log_scroll_frame, orient="vertical", command=self.log_box.yview)
        scrollbar_log_y.grid(row=0, column=1, sticky='ns')
        scrollbar_log_x = Scrollbar(log_scroll_frame, orient=HORIZONTAL, command=self.log_box.xview)
        scrollbar_log_x.grid(row=1, column=0, sticky='ew')
        self.log_box.config(yscrollcommand=scrollbar_log_y.set, xscrollcommand=scrollbar_log_x.set)

        ttk.Button(log_frame, text="Clear Log", command=self.clear_log).grid(row=1, column=1, sticky='e', pady=(5,0))

        manual_frame = ttk.LabelFrame(bottom_frame, text="Manual Triggers", padding="10")
        manual_frame.pack(fill='x')
        self.manual_triggers_var = BooleanVar(value=False)
        ttk.Checkbutton(manual_frame, text="Enable Hotkeys", variable=self.manual_triggers_var).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        row, col = 1, 0
        for key, (val, desc) in EVENT_TRIGGERS.items():
            ttk.Label(manual_frame, text=f"'{key}' -> {desc} ({val})").grid(row=row, column=col, sticky='w', padx=10, pady=2)
            col = 1 - col
            if col == 0: row += 1

    def toggle_timing_mode(self):
        """Handles the logic when the manual timing mode checkbox is clicked."""
        if self.manual_timing_mode.get():
            self.btn_load_audio.config(state='disabled')
            self.audio_file_path = None
            self.audio_path_var.set("Audio File: Disabled (Manual Timing)")
        else:
            self.btn_load_audio.config(state='normal')
            self.audio_path_var.set("Audio File (mp3): Not loaded")
        
        self.check_if_ready()

    def open_playlist_editor(self):
        if self.playlist_editor_window and self.playlist_editor_window.winfo_exists():
            self.playlist_editor_window.lift()
            return
        
        self.playlist_editor_window = Toplevel(self.root)
        self.playlist_editor_window.title("Playlist Editor")
        self.playlist_editor_window.geometry("500x600")
        
        creator_frame = ttk.Frame(self.playlist_editor_window, padding="10")
        creator_frame.pack(fill='both', expand=True)

        input_frame = ttk.LabelFrame(creator_frame, text="Song Details", padding="10")
        input_frame.pack(fill='x')
        input_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(input_frame, text="Song Name:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.csv_name_var = StringVar()
        ttk.Entry(input_frame, textvariable=self.csv_name_var).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        ttk.Label(input_frame, text="Order:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.csv_order_var = StringVar(value=str(len(self.csv_editor_data) + 1))
        ttk.Entry(input_frame, textvariable=self.csv_order_var, width=10).grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(input_frame, text="Duration (MM:SS):").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.csv_duration_var = StringVar()
        ttk.Entry(input_frame, textvariable=self.csv_duration_var, width=10).grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        self.add_update_button = ttk.Button(input_frame, text="Add Song", command=self.add_or_update_song_in_editor)
        self.add_update_button.grid(row=3, column=1, sticky='e', padx=5, pady=10)

        list_frame = ttk.LabelFrame(creator_frame, text="Song List", padding="10")
        list_frame.pack(fill='both', expand=True, pady=10)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.csv_editor_list_box = Listbox(list_frame, font=('Consolas', 10))
        self.csv_editor_list_box.grid(row=0, column=0, sticky='nsew')
        scrollbar_csv_y = Scrollbar(list_frame, orient="vertical", command=self.csv_editor_list_box.yview)
        scrollbar_csv_y.grid(row=0, column=1, sticky='ns')
        self.csv_editor_list_box.config(yscrollcommand=scrollbar_csv_y.set)

        list_action_frame = ttk.Frame(list_frame)
        list_action_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(5,0))
        ttk.Button(list_action_frame, text="Edit Selected", command=self.edit_selected_song).pack(side='left', padx=5)
        ttk.Button(list_action_frame, text="Remove Selected", command=self.remove_selected_song).pack(side='left')

        action_frame = ttk.Frame(creator_frame, padding="5")
        action_frame.pack(fill='x', side='bottom')
        ttk.Button(action_frame, text="Save to CSV File...", command=self.save_editor_to_csv).pack(side='right', padx=5)
        ttk.Button(action_frame, text="Load from CSV...", command=self.load_csv_into_editor).pack(side='right', padx=5)
        ttk.Button(action_frame, text="Clear List", command=self.clear_editor_list).pack(side='right')
        self.update_playlist_editor_listbox()

    def add_or_update_song_in_editor(self):
        name = self.csv_name_var.get()
        order = self.csv_order_var.get()
        duration_str = self.csv_duration_var.get()

        if not all([name, order, duration_str]):
            messagebox.showwarning("Input Error", "All fields must be filled.", parent=self.playlist_editor_window)
            return
            
        duration_int = self.parse_time_to_seconds(duration_str)
        if duration_int is None:
            messagebox.showwarning("Input Error", "Invalid duration format.\nPlease use MM:SS or just seconds.", parent=self.playlist_editor_window)
            return
            
        try:
            order_int = int(order)
        except ValueError:
            messagebox.showwarning("Input Error", "'Order' must be a number.", parent=self.playlist_editor_window)
            return
        
        song_data = {'name': name, 'order': order_int, 'duration': duration_int}
        
        if self.editing_song_index is not None:
            self.csv_editor_data[self.editing_song_index] = song_data
        else:
            self.csv_editor_data.append(song_data)

        self.csv_editor_data.sort(key=lambda x: x['order'])
        self.update_playlist_editor_listbox()
        self.reset_editor_inputs()

    def reset_editor_inputs(self):
        self.editing_song_index = None
        self.csv_name_var.set("")
        self.csv_duration_var.set("")
        self.csv_order_var.set(str(len(self.csv_editor_data) + 1))
        self.add_update_button.config(text="Add Song")

    def edit_selected_song(self):
        selected_indices = self.csv_editor_list_box.curselection()
        if not selected_indices:
            messagebox.showinfo("No Selection", "Please select a song from the list to edit.", parent=self.playlist_editor_window)
            return
        
        selected_index_in_listbox = selected_indices[0]
        selected_song_text = self.csv_editor_list_box.get(selected_index_in_listbox).strip()
        
        for i, song in enumerate(self.csv_editor_data):
            listbox_text = f"#{song['order']} | {song['name']} ({self.format_time_short(song['duration'])})"
            if listbox_text == selected_song_text:
                self.editing_song_index = i
                break
        else:
            return

        song_to_edit = self.csv_editor_data[self.editing_song_index]
        self.csv_name_var.set(song_to_edit['name'])
        self.csv_order_var.set(str(song_to_edit['order']))
        self.csv_duration_var.set(self.format_time_short(song_to_edit['duration']))
        self.add_update_button.config(text="Update Song")

    def remove_selected_song(self):
        selected_indices = self.csv_editor_list_box.curselection()
        if not selected_indices:
            messagebox.showinfo("No Selection", "Please select a song from the list to remove.", parent=self.playlist_editor_window)
            return
            
        if messagebox.askyesno("Confirm Removal", "Are you sure you want to remove the selected song?", parent=self.playlist_editor_window):
            selected_song_text = self.csv_editor_list_box.get(selected_indices[0]).strip()
            
            self.csv_editor_data = [song for song in self.csv_editor_data if f"#{song['order']} | {song['name']} ({self.format_time_short(song['duration'])})" != selected_song_text]

            self.update_playlist_editor_listbox()
            self.reset_editor_inputs()

    def update_playlist_editor_listbox(self):
        if not (self.playlist_editor_window and self.playlist_editor_window.winfo_exists()): return
        self.csv_editor_list_box.delete(0, 'end')
        for item in self.csv_editor_data:
            self.csv_editor_list_box.insert('end', f"  #{item['order']} | {item['name']} ({self.format_time_short(item['duration'])})")

    def clear_editor_list(self):
        if messagebox.askyesno("Confirm", "Are you sure you want to clear the entire song list?", parent=self.playlist_editor_window):
            self.csv_editor_data.clear()
            self.update_playlist_editor_listbox()
            self.reset_editor_inputs()
    
    def load_csv_into_editor(self):
        if self.csv_editor_data:
            if not messagebox.askyesno("Confirm", "This will overwrite the current list in the editor. Continue?", parent=self.playlist_editor_window):
                return

        filepath = filedialog.askopenfilename(
            title="Select Playlist File to Load",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
            parent=self.playlist_editor_window
        )
        if not filepath: return

        temp_songs = []
        try:
            with open(filepath, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        temp_songs.append({
                            'name': row['SongName'],
                            'order': int(row['Reihenfolge']),
                            'duration': int(row['DauerInSekunden'])
                        })
                    except (KeyError, ValueError):
                        continue 
            
            self.csv_editor_data = temp_songs
            self.csv_editor_data.sort(key=lambda x: x['order'])
            self.update_playlist_editor_listbox()
            self.reset_editor_inputs()
            self.log_event(f"Loaded {os.path.basename(filepath)} into the playlist editor.")

        except Exception as e:
            messagebox.showerror("File Error", f"Failed to read or parse file.\nError: {e}", parent=self.playlist_editor_window)
            self.log_event(f"ERROR: Failed to load CSV into editor. Reason: {e}")

    def save_editor_to_csv(self):
        if not self.csv_editor_data:
            messagebox.showinfo("Nothing to Save", "The song list is empty.", parent=self.playlist_editor_window)
            return
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            title="Save Playlist File As...",
            parent=self.playlist_editor_window
        )
        if not filepath: return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['SongName', 'Reihenfolge', 'DauerInSekunden'])
                
                for item in self.csv_editor_data:
                    writer.writerow([item['name'], item['order'], item['duration']])
            
            messagebox.showinfo("Success", f"Playlist file saved successfully to:\n{filepath}", parent=self.playlist_editor_window)
            self.log_event(f"Created/saved playlist file: {os.path.basename(filepath)}")
            self.load_playlist(filepath)
            self.playlist_editor_window.destroy()

        except Exception as e:
            messagebox.showerror("File Error", f"Failed to save file.\nError: {e}", parent=self.playlist_editor_window)
            self.log_event(f"ERROR: Failed to save playlist file. Reason: {e}")

    def log_event(self, message):
        time_str = time.strftime("%H:%M:%S")
        log_entry = f"[{time_str}] {message}"
        self.log_box.insert(0, log_entry)
        if self.log_box.size() > 200: self.log_box.delete(200)

    def clear_log(self):
        self.log_box.delete(0, END)

    def browse_for_audio(self):
        filepath = filedialog.askopenfilename(title="Select Master Audio File", filetypes=(("Audio Files", "*.mp3 *.wav *.ogg"), ("All files", "*.*")))
        if filepath:
            self.audio_file_path = filepath
            basename = os.path.basename(filepath)
            self.audio_path_var.set(f"Audio File: {basename}")
            self.log_event(f"Loaded audio file: {basename}")
            self.check_if_ready()

    def browse_for_playlist(self):
        filepath = filedialog.askopenfilename(title="Select Playlist File", filetypes=(("CSV files", "*.csv"),))
        if filepath:
            self.load_playlist(filepath)

    def load_playlist(self, filepath):
        self.playlist.clear(); self.playlist_list_box.delete(0, 'end'); temp_playlist = []
        try:
            with open(filepath, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        temp_playlist.append({
                            'name': row['SongName'],
                            'order': int(row['Reihenfolge']),
                            'duration': int(row['DauerInSekunden'])
                        })
                    except (KeyError, ValueError):
                        continue 
            
            temp_playlist.sort(key=lambda x: x['order'])
            current_start_time = 0
            for song in temp_playlist:
                self.playlist.append({'name': song['name'], 'duration': song['duration'], 'start_time': current_start_time})
                display_text = f"  {song['name']} (Start: {self.format_time_long(current_start_time)}, Dur: {self.format_time_short(song['duration'])})"
                self.playlist_list_box.insert('end', display_text)
                current_start_time += song['duration']
            
            self.total_mix_duration = current_start_time
            self.time_display_var.set(f"00:00:00 / {self.format_time_long(self.total_mix_duration)}")
            
            if self.playlist:
                basename = os.path.basename(filepath)
                self.playlist_path_var.set(f"Playlist (CSV): {basename}")
                self.log_event(f"Loaded and processed playlist: {basename}")
            self.check_if_ready()
        except Exception as e:
            messagebox.showerror("Playlist File Error", f"Failed to read or parse file.\nError: {e}")
            self.log_event(f"ERROR: Failed to load playlist file. Reason: {e}")
    
    def check_if_ready(self):
        if (self.audio_file_path or self.manual_timing_mode.get()) and self.playlist:
            self.btn_start_automation.config(state='normal')
            self.log_event("Audio/Timing Mode and Playlist loaded. Ready to start automation.")
        else:
            self.btn_start_automation.config(state='disabled')

    def start_automation(self):
        if self.is_automation_running: return
        self.is_automation_running = True
        self.btn_start_automation.config(state='disabled', text="START")
        self.btn_load_audio.config(state='disabled')
        self.btn_load_playlist.config(state='disabled')
        self.btn_create_playlist.config(state='disabled')
        self.manual_mode_check.config(state='disabled') # Checkbox ebenfalls sperren
        self.btn_pause_automation.config(state='normal')
        self.btn_stop_automation.config(state='normal')
        
        self.log_event("=== Automation Started ===")
        
        if not self.manual_timing_mode.get():
            if pygame and self.audio_file_path:
                pygame.mixer.music.load(self.audio_file_path)
                pygame.mixer.music.play()
            else:
                messagebox.showerror("Audio Error", "Audio file is required but not found or Pygame is missing.")
                self.stop_automation()
                return
        else:
            self.start_time = time.monotonic()
            self.total_pause_duration = 0
            self.pause_start_time = 0

        self.prepare_all_triggers()
        self.run_scheduled_events(0)
        self.update_time_display()

    def pause_automation(self):
        if not self.is_automation_running or self.is_paused: return
        self.is_paused = True
        
        if not self.manual_timing_mode.get() and pygame:
            pygame.mixer.music.pause()
            self.pause_time_ms = pygame.mixer.music.get_pos()
        else:
            self.pause_start_time = time.monotonic()
            elapsed_time = (self.pause_start_time - self.start_time) - self.total_pause_duration
            self.pause_time_ms = elapsed_time * 1000

        self.cancel_scheduled_events()
        self.btn_pause_automation.config(text="CONTINUE", command=self.continue_automation)
        self.log_event(f"--- Automation Paused at {self.format_time_long(self.pause_time_ms / 1000)} ---")

    def continue_automation(self):
        if not self.is_automation_running or not self.is_paused: return
        self.is_paused = False

        if not self.manual_timing_mode.get() and pygame:
            pygame.mixer.music.unpause()
        else:
            pause_duration = time.monotonic() - self.pause_start_time
            self.total_pause_duration += pause_duration

        self.btn_pause_automation.config(text="PAUSE", command=self.pause_automation)
        self.log_event(f"--- Automation Continued at {self.format_time_long(self.pause_time_ms / 1000)} ---")
        self.run_scheduled_events(self.pause_time_ms)
        self.update_time_display()

    def stop_automation(self, finished=False):
        if not self.is_automation_running and not finished: return
        if finished: self.log_event("=== Automation Finished ===")
        else: self.log_event("=== Automation Stopped by User ===")
        
        if not self.manual_timing_mode.get() and pygame:
            pygame.mixer.music.stop()

        self.cancel_scheduled_events()
        self.is_automation_running = False
        self.is_paused = False
        self.pause_time_ms = 0

        self.start_time = 0
        self.total_pause_duration = 0
        self.pause_start_time = 0
        
        self.now_playing_var.set("Now Playing: - (Stopped)")
        self.time_display_var.set(f"00:00:00 / {self.format_time_long(self.total_mix_duration)}")
        self.btn_start_automation.config(state='normal', text="RESTART")
        self.btn_load_audio.config(state='normal')
        self.btn_load_playlist.config(state='normal')
        self.btn_create_playlist.config(state='normal')
        self.manual_mode_check.config(state='normal') 
        self.toggle_timing_mode()
        self.btn_pause_automation.config(state='disabled', text="PAUSE", command=self.pause_automation)
        self.btn_stop_automation.config(state='disabled')
        self.playlist_list_box.selection_clear(0, 'end')

    def prepare_all_triggers(self):
            self.all_automation_events.clear()
            try:
                transition_s = int(self.transition_time_var.get())
            except ValueError:
                transition_s = 15; self.transition_time_var.set("15")

            # Anzahl der Songs für späteren Check speichern
            total_songs = len(self.playlist)

            for i, song in enumerate(self.playlist):
                start_s = song['start_time']; duration_s = song['duration']; end_s = start_s + duration_s
                
                # --- EVENTS PLANEN ---

                # 1. Update GUI & Song Start (Trigger 1)
                self.all_automation_events.append((start_s * 1000, self.update_now_playing, (song['name'], i)))
                self.all_automation_events.append((start_s * 1000, self.send_trigger, ('1', f"Start: {song['name']}")))
                
                # 2. Initial Transition End (nur beim allerersten Song)
                if i == 0: 
                    self.all_automation_events.append((transition_s * 1000, self.send_trigger, ('10', 'Initial Transition End')))
                
                # 3. Transition End (für alle Songs außer dem ersten)
                if transition_s > 0 and i > 0: 
                    self.all_automation_events.append(((start_s + transition_s) * 1000, self.send_trigger, ('10', f"Transition End for {song['name']}")))
                
                # 4. Transition Start (Trigger 9) 
                # "and i < total_songs - 1" verhindert den Trigger beim allerletzten Song.
                if duration_s > transition_s and i < total_songs - 1: 
                    self.all_automation_events.append(((end_s - transition_s) * 1000, self.send_trigger, ('9', f"Transition Start for next song")))
                
                # 5. Song End (Trigger 2)
                # Beim letzten Song ist es gleichzeitig Marker für Playlist-Ende
                desc = f"End: {song['name']}"
                if i == total_songs - 1:
                    desc = f"End: {song['name']} (Playlist End)"
                
                self.all_automation_events.append((end_s * 1000, self.send_trigger, ('2', desc)))
            
            # Stopp-Funktion muss aber bleiben (etwas verzögert, damit der letzte Trigger sicher rausgeht)
            self.all_automation_events.append((self.total_mix_duration * 1000 + 100, self.stop_automation, (True,)))
            
            self.log_event(f"All triggers for {len(self.playlist)} songs have been prepared.")
    
    def run_scheduled_events(self, start_offset_ms):
        self.cancel_scheduled_events()
        for delay_ms, callback, args in self.all_automation_events:
            if delay_ms >= start_offset_ms:
                new_delay = delay_ms - start_offset_ms
                event_id = self.root.after(int(new_delay), callback, *args)
                self.scheduled_events.append(event_id)

    def update_time_display(self):
        if self.is_automation_running and not self.is_paused:
            current_pos_s = 0
            if self.manual_timing_mode.get():
                current_pos_s = (time.monotonic() - self.start_time) - self.total_pause_duration
            elif pygame and pygame.mixer.music.get_busy():
                current_pos_s = pygame.mixer.music.get_pos() / 1000
            
            self.time_display_var.set(f"{self.format_time_long(current_pos_s)} / {self.format_time_long(self.total_mix_duration)}")
        
        self.root.after(500, self.update_time_display)

    def update_now_playing(self, text, index):
        self.now_playing_var.set(f"Now Playing: {text}")
        self.playlist_list_box.selection_clear(0, 'end')
        self.playlist_list_box.selection_set(index); self.playlist_list_box.activate(index)
        self.log_event(f"Song changed to: '{text}'")

    def cancel_scheduled_events(self):
        for event_id in self.scheduled_events: self.root.after_cancel(event_id)
        self.scheduled_events.clear()

    def on_closing(self):
        self.log_event("Application shutting down.")
        self.stop_automation()
        self.sock.close()
        self.root.destroy()
        
    def setup_keyboard_hooks(self):
        threading.Thread(target=lambda: keyboard.hook(self.handle_key_press), daemon=True).start()

    def handle_key_press(self, event):
        if not self.manual_triggers_var.get(): return
        if event.event_type == keyboard.KEY_DOWN and event.name in EVENT_TRIGGERS:
            trigger, desc = EVENT_TRIGGERS[event.name]
            self.root.after(0, self.send_trigger, trigger, f"<<Manual>> {desc}")

    def send_trigger(self, trigger_value, event_desc):
        try:
            message = str(trigger_value)
            sendBytes = message.encode('utf-8')
            self.sock.sendto(sendBytes, self.endPoint)
            
            status_text = f"Sent '{trigger_value}' for '{event_desc}'"
            self.last_trigger_var.set(f"Last Trigger: {status_text}")
            self.log_event(f"TRIGGER SENT -> Value: {trigger_value}, Desc: {event_desc}")
            
        except Exception as e:
            self.last_trigger_var.set(f"Error sending trigger: {e}")
            self.log_event(f"ERROR sending trigger '{trigger_value}'. Reason: {e}")

if __name__ == "__main__":
    root = Tk()
    app = TriggerApp(root)
    root.mainloop()
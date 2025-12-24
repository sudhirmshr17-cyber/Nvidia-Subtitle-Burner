import customtkinter as ctk
import os
import threading
import time
import shutil
import re
import subprocess
import psutil
import winsound
import glob
import sys  # Added at top level for safety
from tkinter import filedialog, messagebox

# --- DRAG & DROP CHECK ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    dnd_available = True
except ImportError:
    dnd_available = False

# ==========================================
#         DYNAMIC THEME PALETTE
# ==========================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# Theme Colors (Light, Dark)
COLOR_BG = ("#F3F3F3", "#121212")        # Main Window Background
COLOR_SURFACE = ("#FFFFFF", "#1E1E1E")   # Card/Bar Backgrounds
COLOR_QUEUE_BG = ("#E5E5E5", "#181818")  # Queue Scroll Area
COLOR_ACCENT = ("#007EA7", "#00ADB5")    # Teal/Blue
COLOR_SUCCESS = ("#009955", "#008f4c")   # Green
COLOR_WARNING = ("#D4AF37", "#D4AF37")   # Gold
COLOR_DANGER = ("#DC3545", "#CF6679")    # Red
COLOR_TEXT_MAIN = ("#101010", "#E0E0E0") # Text High Contrast
COLOR_TEXT_DIM = ("#505050", "#A0A0A0")  # Text Low Contrast
COLOR_BORDER = ("#CCCCCC", "#333333")    # Borders

# Mappings
PRESET_MAP = {"Fast (p1)": "p1", "Balanced (p4)": "p4", "Best (p7)": "p7"}
AUDIO_MAP = {"Copy": "copy", "AAC": "aac", "Normalize": "normalize"}
COLORS = {"White": "&HFFFFFF&", "Yellow": "&H00FFFF&", "Cyan": "&HFFFF00&", "Green": "&H00FF00&"}

BaseClass = TkinterDnD.Tk if dnd_available else ctk.CTk

# ==========================================
#           CUSTOM WIDGET: QUEUE CARD
# ==========================================
class QueueItem(ctk.CTkFrame):
    def __init__(self, master, filepath, remove_callback, move_callback):
        # NOTE: bg_color=COLOR_QUEUE_BG fixes the corners of the cards inside the list
        super().__init__(master, fg_color=COLOR_SURFACE, corner_radius=6, border_width=1, border_color=COLOR_BORDER, bg_color=COLOR_QUEUE_BG)
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        
        self.grid_columnconfigure(0, weight=1)
        
        # Filename
        display_name = self.filename if len(self.filename) < 55 else self.filename[:52] + "..."
        self.lbl_name = ctk.CTkLabel(self, text=display_name, text_color=COLOR_TEXT_MAIN, anchor="w", font=("Roboto", 12))
        self.lbl_name.grid(row=0, column=0, padx=10, pady=8, sticky="ew")
        
        # Controls Frame
        self.ctrl_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.ctrl_frame.grid(row=0, column=1, padx=5, pady=2)
        
        # Button Colors (Tuples)
        btn_bg = ("#E0E0E0", "#333333")
        btn_hover = ("#D0D0D0", "#444444")

        self.btn_up = ctk.CTkButton(self.ctrl_frame, text="▲", width=25, height=25, fg_color=btn_bg, text_color=COLOR_TEXT_MAIN, hover_color=btn_hover, command=lambda: move_callback(self, -1))
        self.btn_up.pack(side="left", padx=2)
        
        self.btn_down = ctk.CTkButton(self.ctrl_frame, text="▼", width=25, height=25, fg_color=btn_bg, text_color=COLOR_TEXT_MAIN, hover_color=btn_hover, command=lambda: move_callback(self, 1))
        self.btn_down.pack(side="left", padx=2)
        
        self.btn_del = ctk.CTkButton(self.ctrl_frame, text="✕", width=25, height=25, fg_color=COLOR_DANGER, text_color="white", command=lambda: remove_callback(self))
        self.btn_del.pack(side="left", padx=(10, 2))

    def set_active(self, active=True):
        if active:
            self.configure(border_color=COLOR_ACCENT, border_width=2)
            self.lbl_name.configure(text_color=COLOR_ACCENT)
        else:
            self.configure(border_color=COLOR_BORDER, border_width=1)
            self.lbl_name.configure(text_color=COLOR_TEXT_MAIN)

    def set_done(self):
        self.configure(border_color=COLOR_SUCCESS)
        self.lbl_name.configure(text_color=COLOR_SUCCESS)

    def update_theme(self):
        """Forces the card to redraw colors based on current mode"""
        self.configure(fg_color=COLOR_SURFACE, border_color=COLOR_BORDER)
        self.lbl_name.configure(text_color=COLOR_TEXT_MAIN)
        
        btn_bg = ("#E0E0E0", "#333333")
        btn_hover = ("#D0D0D0", "#444444")
        self.btn_up.configure(fg_color=btn_bg, text_color=COLOR_TEXT_MAIN, hover_color=btn_hover)
        self.btn_down.configure(fg_color=btn_bg, text_color=COLOR_TEXT_MAIN, hover_color=btn_hover)

# ==========================================
#               MAIN APPLICATION
# ==========================================
class subtitle_burner(BaseClass):
    def __init__(self):
        super().__init__()
        
        # --- VERSION 1.0 ---
        self.title("Nvidia Subtitle Burner [V1.0]")
        self.geometry("1200x750")
        
        # Initial BG Set
        if dnd_available:
            self.config(bg=COLOR_BG[1]) # Start in Dark Mode Hex
        else:
            self.configure(fg_color=COLOR_BG)
        
        # --- ICON FIX ---
        # Robustly find the icon whether running as .py or .exe
        try:
            if getattr(sys, 'frozen', False):  # Running as compiled .exe
                icon_path = os.path.join(sys._MEIPASS, "myicon.ico")
            else:  # Running as .py script
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "myicon.ico")
            
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass  # Fail silently if icon has issues

        if dnd_available:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.drop_event)

        # Logic State
        self.queue_items = []
        self.is_running = False
        self.is_paused = False
        self.stop_event = threading.Event()
        self.current_process = None
        self.ffmpeg_exe = None

        # Grid Layout
        self.grid_columnconfigure(0, weight=0) # Sidebar
        self.grid_columnconfigure(1, weight=1) # Main
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)    # Footer

        self.setup_sidebar()
        self.setup_dashboard()
        self.setup_statusbar()
        
        self.check_ffmpeg()
        
        # GPU Thread
        self.monitor_thread = threading.Thread(target=self.monitor_system, daemon=True)
        self.monitor_thread.start()

    # --- GUI SETUP ---
    def setup_sidebar(self):
        # NOTE: bg_color=COLOR_BG fixes Sidebar corners
        self.sidebar = ctk.CTkFrame(self, width=260, fg_color=COLOR_SURFACE, corner_radius=0, bg_color=COLOR_BG)
        self.sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar.grid_propagate(False)

        # Title - Updated to V1
        ctk.CTkLabel(self.sidebar, text="AUTO SUBTITLE\nBURNER V1", font=("Montserrat", 20, "bold"), text_color=COLOR_ACCENT).pack(pady=(30, 20))
        ctk.CTkFrame(self.sidebar, height=2, fg_color=COLOR_BORDER).pack(fill="x", padx=20, pady=10)

        # VISUALS
        ctk.CTkLabel(self.sidebar, text="VISUALS", text_color=COLOR_TEXT_DIM, font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(10,5))
        
        self.side_font = ctk.CTkOptionMenu(self.sidebar, values=["Arial", "Roboto", "Consolas"], fg_color=COLOR_BG, button_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.side_font.pack(fill="x", padx=20, pady=5)
        
        self.side_color = ctk.CTkOptionMenu(self.sidebar, values=list(COLORS.keys()), fg_color=COLOR_BG, button_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.side_color.pack(fill="x", padx=20, pady=5)
        
        # SIZE SLIDER
        self.lbl_fontsize = ctk.CTkLabel(self.sidebar, text="Size: 24px", text_color=COLOR_TEXT_DIM)
        self.lbl_fontsize.pack(anchor="w", padx=25, pady=(5,0))
        
        self.side_size = ctk.CTkSlider(self.sidebar, from_=12, to=64, number_of_steps=52, progress_color=COLOR_ACCENT, command=self.update_font_label)
        self.side_size.set(24)
        self.side_size.pack(fill="x", padx=20, pady=5)

        # ENCODING
        ctk.CTkLabel(self.sidebar, text="ENCODING", text_color=COLOR_TEXT_DIM, font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(20,5))
        
        self.side_preset = ctk.CTkOptionMenu(self.sidebar, values=list(PRESET_MAP.keys()), fg_color=COLOR_BG, button_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.side_preset.set("Balanced (p4)")
        self.side_preset.pack(fill="x", padx=20, pady=5)
        
        self.side_audio = ctk.CTkOptionMenu(self.sidebar, values=list(AUDIO_MAP.keys()), fg_color=COLOR_BG, button_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.side_audio.pack(fill="x", padx=20, pady=5)

        # ACTION
        ctk.CTkLabel(self.sidebar, text="FINISH ACTION", text_color=COLOR_TEXT_DIM, font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(20,5))
        self.side_finish = ctk.CTkOptionMenu(self.sidebar, values=["Do Nothing", "Play Sound", "Close App", "Shutdown PC"], fg_color=COLOR_BG, button_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.side_finish.pack(fill="x", padx=20, pady=5)

        # THEME SWITCH
        self.switch_theme = ctk.CTkSwitch(self.sidebar, text="Light Mode", command=self.toggle_theme, progress_color=COLOR_ACCENT, text_color=COLOR_TEXT_MAIN)
        self.switch_theme.pack(side="bottom", padx=20, pady=20)

    def setup_dashboard(self):
        # FIX: Replaced "transparent" with COLOR_BG to fix the 'line' artifact
        self.main_area = ctk.CTkFrame(self, fg_color=COLOR_BG)
        self.main_area.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_area.grid_columnconfigure(0, weight=1)
        self.main_area.grid_rowconfigure(1, weight=1)

        # Top Bar (NOTE: bg_color=COLOR_BG fixes black corners)
        self.top_bar = ctk.CTkFrame(self.main_area, fg_color=COLOR_SURFACE, height=60, corner_radius=8, bg_color=COLOR_BG)
        self.top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        # Import Buttons
        self.btn_browse = ctk.CTkButton(self.top_bar, text="Import Files", fg_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN, hover_color=COLOR_ACCENT, width=100, command=self.browse_files)
        self.btn_browse.pack(side="right", padx=(5, 15), pady=10)

        self.btn_folder = ctk.CTkButton(self.top_bar, text="Import Folder", fg_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN, hover_color=COLOR_ACCENT, width=100, command=self.browse_folder)
        self.btn_folder.pack(side="right", padx=(0, 5), pady=10)
        
        self.path_entry = ctk.CTkEntry(self.top_bar, placeholder_text="Drag & Drop Files or Folders here...", fg_color=COLOR_BG, border_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=15, pady=10)

        # Queue (NOTE: bg_color=COLOR_BG fixes list area corners)
        self.queue_container = ctk.CTkScrollableFrame(self.main_area, label_text="BATCH QUEUE", label_font=("Arial", 12, "bold"), label_text_color=COLOR_TEXT_MAIN, fg_color=COLOR_QUEUE_BG, bg_color=COLOR_BG)
        self.queue_container.grid(row=1, column=0, sticky="nsew")

        # Action Bar (NOTE: fg_color=COLOR_BG fixes black box behind buttons)
        self.action_bar = ctk.CTkFrame(self.main_area, fg_color=COLOR_BG, corner_radius=0)
        self.action_bar.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        
        self.btn_start = ctk.CTkButton(self.action_bar, text="START BATCH", width=150, height=45, font=("Arial", 13, "bold"), fg_color=COLOR_SUCCESS, hover_color="#006b3a", text_color="white", bg_color=COLOR_BG, command=self.start_thread)
        self.btn_start.pack(side="right", padx=5)
        
        self.btn_preview = ctk.CTkButton(self.action_bar, text="PREVIEW", width=120, height=45, fg_color=COLOR_ACCENT, hover_color="#008c93", text_color="white", bg_color=COLOR_BG, command=self.preview_video)
        self.btn_preview.pack(side="right", padx=5)
        
        self.btn_pause = ctk.CTkButton(self.action_bar, text="PAUSE", width=100, height=45, fg_color=COLOR_WARNING, state="disabled", text_color="white", bg_color=COLOR_BG, command=self.toggle_pause)
        self.btn_pause.pack(side="right", padx=5)

        self.btn_cancel = ctk.CTkButton(self.action_bar, text="CANCEL", width=100, height=45, fg_color=COLOR_DANGER, state="disabled", text_color="white", bg_color=COLOR_BG, command=self.cancel_process)
        self.btn_cancel.pack(side="right", padx=5)
        
        self.btn_clear = ctk.CTkButton(self.action_bar, text="CLEAR ALL", width=100, height=45, fg_color="transparent", border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT_MAIN, hover_color=("#e6e6e6", "#2f2f2f"), command=self.clear_queue)
        self.btn_clear.pack(side="left", padx=0)

    def setup_statusbar(self):
        bar_bg = ("#D0D0D0", "#0f0f0f")
        self.status_bar = ctk.CTkFrame(self, height=35, fg_color=bar_bg, corner_radius=0, bg_color=COLOR_BG)
        self.status_bar.grid(row=1, column=1, sticky="ew")
        
        self.status_text = ctk.CTkLabel(self.status_bar, text="Ready", font=("Consolas", 11), text_color=COLOR_TEXT_DIM)
        self.status_text.pack(side="left", padx=20)
        
        self.gpu_stat = ctk.CTkLabel(self.status_bar, text="GPU: --%", font=("Consolas", 11, "bold"), text_color=COLOR_ACCENT)
        self.gpu_stat.pack(side="right", padx=20)
        
        self.progress_bar = ctk.CTkProgressBar(self.status_bar, width=300, height=8, progress_color=COLOR_ACCENT)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=20)

    # --- THEME LOGIC ---
    def toggle_theme(self):
        # Determine strict Mode Color (0=Light, 1=Dark)
        is_light = self.switch_theme.get() == 1
        mode_idx = 0 if is_light else 1
        current_hex = COLOR_BG[mode_idx]

        if is_light:
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")
            
        # 1. Update Root Window Background (Fix for DND Tkinter window)
        if dnd_available:
            self.config(bg=current_hex)

        # 2. Update Main Container to hide "transparency artifacts"
        self.main_area.configure(fg_color=current_hex)

        # 3. Explicitly update background colors of containers
        self.top_bar.configure(bg_color=current_hex)
        self.queue_container.configure(bg_color=current_hex)
        self.status_bar.configure(bg_color=current_hex)
        self.action_bar.configure(fg_color=current_hex) # action_bar is fg, not bg because it isn't rounded

        # 4. Update Buttons to match background
        self.btn_start.configure(bg_color=current_hex)
        self.btn_preview.configure(bg_color=current_hex)
        self.btn_pause.configure(bg_color=current_hex)
        self.btn_cancel.configure(bg_color=current_hex)
        self.btn_clear.configure(fg_color=current_hex)

        # 5. Refresh Cards
        for item in self.queue_items:
            item.update_theme()

    def update_font_label(self, value):
        self.lbl_fontsize.configure(text=f"Size: {int(value)}px")

    # --- LOGIC: SYSTEM & QUEUE ---
    def check_ffmpeg(self):
        self.ffmpeg_exe = shutil.which("ffmpeg")
        if not self.ffmpeg_exe and os.path.exists("ffmpeg.exe"):
            self.ffmpeg_exe = os.path.abspath("ffmpeg.exe")
        
        if not self.ffmpeg_exe:
            self.status_text.configure(text="ERROR: FFmpeg not found", text_color=COLOR_DANGER)
            self.btn_start.configure(state="disabled")

    def drop_event(self, event):
        if self.is_running: return
        data = event.data
        if data.startswith('{') and data.endswith('}'):
            items = re.split(r'\} \{', data[1:-1])
        else:
            items = data.split()
        
        for item in items:
            if os.path.isdir(item):
                self.add_folder_to_queue(item)
            else:
                self.add_file_to_queue(item)

    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Video", "*.mp4 *.mkv *.avi")])
        if files:
            for f in files: self.add_file_to_queue(f)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.add_folder_to_queue(folder)

    def add_folder_to_queue(self, folder):
        files = glob.glob(os.path.join(folder, "*"))
        valid_exts = ('.mp4', '.mkv', '.avi')
        count = 0
        for f in files:
            if f.lower().endswith(valid_exts):
                self.add_file_to_queue(f)
                count += 1
        self.status_text.configure(text=f"Imported {count} files from folder.")

    def add_file_to_queue(self, filepath):
        if filepath not in [item.filepath for item in self.queue_items]:
            item = QueueItem(self.queue_container, filepath, self.remove_item, self.move_item)
            item.pack(fill="x", pady=2, padx=5)
            self.queue_items.append(item)

    def remove_item(self, item_widget):
        if self.is_running: return
        item_widget.destroy()
        if item_widget in self.queue_items: self.queue_items.remove(item_widget)

    def move_item(self, item_widget, direction):
        if self.is_running: return
        if item_widget not in self.queue_items: return
        curr_idx = self.queue_items.index(item_widget)
        new_idx = curr_idx + direction
        if 0 <= new_idx < len(self.queue_items):
            self.queue_items[curr_idx], self.queue_items[new_idx] = self.queue_items[new_idx], self.queue_items[curr_idx]
            for item in self.queue_items: item.pack_forget()
            for item in self.queue_items: item.pack(fill="x", pady=2, padx=5)

    def clear_queue(self):
        if self.is_running: return
        for item in self.queue_items: item.destroy()
        self.queue_items.clear()

    # --- LOGIC: GPU MONITOR ---
    def monitor_system(self):
        while True:
            try:
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=utilization.gpu', '--format=csv,noheader,nounits'],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=0x08000000
                )
                if result.returncode == 0:
                    usage = result.stdout.strip()
                    self.gpu_stat.configure(text=f"GPU: {usage}%")
                else:
                    self.gpu_stat.configure(text="GPU: N/A")
            except:
                self.gpu_stat.configure(text="GPU: N/A")
            time.sleep(2)

    # --- LOGIC: PROCESSING ---
    def get_settings(self):
        return {
            "font": self.side_font.get(),
            "size": int(self.side_size.get()),
            "color": COLORS[self.side_color.get()],
            "preset": PRESET_MAP[self.side_preset.get()],
            "audio": AUDIO_MAP[self.side_audio.get()],
            "finish": self.side_finish.get()
        }

    def set_ui_locked(self, locked):
        state = "disabled" if locked else "normal"
        # Sidebar
        self.side_font.configure(state=state)
        self.side_size.configure(state=state)
        self.side_color.configure(state=state)
        self.side_preset.configure(state=state)
        self.side_audio.configure(state=state)
        # Main
        self.btn_browse.configure(state=state)
        self.btn_folder.configure(state=state)
        self.path_entry.configure(state=state)
        self.btn_clear.configure(state=state)
        self.btn_start.configure(state=state)
        self.btn_preview.configure(state=state)
        # Controls
        self.btn_pause.configure(state="normal" if locked else "disabled")
        self.btn_cancel.configure(state="normal" if locked else "disabled")
        # Items
        for item in self.queue_items:
            item.btn_del.configure(state=state)
            item.btn_up.configure(state=state)
            item.btn_down.configure(state=state)

    def start_thread(self):
        if not self.queue_items: return
        self.is_running = True
        self.stop_event.clear()
        self.set_ui_locked(True)
        threading.Thread(target=self.process_queue, daemon=True).start()

    def process_queue(self):
        settings = self.get_settings()
        output_dir = os.path.join(os.path.dirname(self.queue_items[0].filepath), "Output_V3")
        if not os.path.exists(output_dir): os.makedirs(output_dir)

        style = f"FontName={settings['font']},Fontsize={settings['size']},PrimaryColour={settings['color']},Bold=1,Outline=2,Shadow=1,MarginV=25"
        files_completed = True

        for item in self.queue_items:
            if self.stop_event.is_set(): 
                files_completed = False
                break

            item.set_active(True)
            self.status_text.configure(text=f"Processing: {item.filename}")
            
            fpath = item.filepath
            out_path = os.path.join(output_dir, item.filename)
            duration = self.get_duration(fpath)

            srt_path = os.path.splitext(fpath)[0] + ".srt"
            if os.path.exists(srt_path):
                safe_srt = srt_path.replace("\\", "/").replace(":", "\\:")
                vf = f"subtitles='{safe_srt}':force_style='{style}'"
            else:
                safe_inp = fpath.replace("\\", "/").replace(":", "\\:")
                vf = f"subtitles='{safe_inp}':force_style='{style}'"

            acmd = ['-c:a', 'copy']
            if settings['audio'] == 'aac': acmd = ['-c:a', 'aac', '-b:a', '192k']
            elif settings['audio'] == 'normalize': acmd = ['-af', 'loudnorm=I=-16:TP=-1.5:LRA=11', '-c:a', 'aac']

            cmd = [
                self.ffmpeg_exe, '-y', '-hide_banner', '-hwaccel', 'cuda',
                '-i', fpath, '-vf', f"{vf},format=yuv420p",
                '-c:v', 'hevc_nvenc', '-preset', settings['preset'], '-cq', '22'
            ] + acmd + [out_path]

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            self.progress_bar.set(0)
            
            self.current_process = subprocess.Popen(
                cmd, stderr=subprocess.PIPE, universal_newlines=True, encoding='utf-8', errors='replace',
                startupinfo=startupinfo, creationflags=0x08000000
            )

            pattern = re.compile(r"time=(\d{2}:\d{2}:\d{2}\.\d{2})")
            while True:
                if self.is_paused: time.sleep(0.5); continue
                line = self.current_process.stderr.readline()
                if not line and self.current_process.poll() is not None: break
                if line:
                    match = pattern.search(line)
                    if match:
                        t_str = match.group(1)
                        h,m,s = map(float, t_str.split(':'))
                        curr = h*3600 + m*60 + s
                        if duration > 0: self.progress_bar.set(curr/duration)

            if self.stop_event.is_set():
                files_completed = False
                break
                
            item.set_active(False)
            item.set_done()

        if files_completed:
            self.finish_sequence(settings['finish'])
        else:
            self.status_text.configure(text="Batch Cancelled.")
            self.progress_bar.set(0)
            self.set_ui_locked(False)
            self.is_running = False

    def finish_sequence(self, action):
        self.status_text.configure(text="All Tasks Completed.")
        self.progress_bar.set(1)
        self.set_ui_locked(False)
        self.is_running = False
        
        if action == "Play Sound":
            winsound.MessageBeep(winsound.MB_OK)
            messagebox.showinfo("Done", "Complete!")
        elif action == "Close App":
            self.destroy()
        elif action == "Shutdown PC":
            os.system("shutdown /s /t 60")
        else:
            messagebox.showinfo("Done", "Complete!")

    def preview_video(self):
        if not self.queue_items: return
        target = self.queue_items[0].filepath
        settings = self.get_settings()
        
        self.status_text.configure(text="Generating Preview...")
        threading.Thread(target=self.run_preview, args=(target, settings), daemon=True).start()

    def run_preview(self, input_path, settings):
        output_path = os.path.join(os.path.dirname(input_path), "preview.mp4")
        style = f"FontName={settings['font']},Fontsize={settings['size']},PrimaryColour={settings['color']},Bold=1,Outline=2,Shadow=1,MarginV=25"
        
        srt_path = os.path.splitext(input_path)[0] + ".srt"
        if os.path.exists(srt_path):
            safe_srt = srt_path.replace("\\", "/").replace(":", "\\:")
            vf = f"subtitles='{safe_srt}':force_style='{style}'"
        else:
            safe_inp = input_path.replace("\\", "/").replace(":", "\\:")
            vf = f"subtitles='{safe_inp}':force_style='{style}'"

        cmd = [
            self.ffmpeg_exe, '-y', '-hide_banner', '-t', '30',
            '-i', input_path, '-vf', vf,
            '-c:v', 'hevc_nvenc', '-preset', 'p1', output_path
        ]
        
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, startupinfo=startupinfo, creationflags=0x08000000)
        
        if os.path.exists(output_path):
            os.startfile(output_path)
            self.status_text.configure(text="Preview Launched.")
        else:
            self.status_text.configure(text="Preview Failed.")

    def get_duration(self, fpath):
        try:
            cmd = [self.ffmpeg_exe, '-i', fpath]
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            r = subprocess.run(cmd, stderr=subprocess.PIPE, encoding='utf-8', startupinfo=startupinfo, creationflags=0x08000000)
            m = re.search(r"Duration: (\d{2}):(\d{2}):(\d{2}\.\d{2})", r.stderr)
            if m:
                h,m,s = map(float, m.groups())
                return h*3600 + m*60 + s
        except: pass
        return 100

    def toggle_pause(self):
        if not self.current_process: return
        try:
            p = psutil.Process(self.current_process.pid)
            if self.is_paused:
                p.resume()
                self.is_paused = False
                self.btn_pause.configure(text="PAUSE", fg_color=COLOR_WARNING)
            else:
                p.suspend()
                self.is_paused = True
                self.btn_pause.configure(text="RESUME", fg_color="#3B8ED0")
        except: pass

    def cancel_process(self):
        if self.current_process:
            self.stop_event.set()
            self.current_process.terminate()

if __name__ == "__main__":
    app = subtitle_burner()
    app.mainloop()
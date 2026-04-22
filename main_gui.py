import threading
from tkinter import filedialog, StringVar
import customtkinter as ctk
from datetime import datetime

from report_logic import generate_report 
from folder_generator import open_folder_popup

# -------------------------
# Theme Setup & GUI Setup
# -------------------------
GUI_BG, GUI_DARK, GUI_DARKER = "#36393F", "#2F3136", "#202225"
GUI_BLURPLE, GUI_BLURPLE_HOVER = "#5865F2", "#4752C4"
GUI_GREEN, GUI_TEXT, GUI_MUTED = "#57F287", "#DCDDDE", "#8E9297"
GUI_RED, GUI_RED_HOVER = "#ED4245", "#C9383B" 

ctk.set_appearance_mode("dark")  
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Auto Test Report Generator")
app.geometry("950x500") 
app.configure(fg_color=GUI_BG)

app.grid_columnconfigure(0, weight=0, minsize=350) 
app.grid_columnconfigure(1, weight=1)              
app.grid_rowconfigure(0, weight=1)

excel_path = StringVar()

def select_excel():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file: excel_path.set(file)

def clear_all():
    word_path.set("")
    excel_path.set("") # ล้างค่า Excel
    folder_path.set("")

word_path, folder_path = StringVar(), StringVar()
status_text = StringVar(value="Ready to generate...")

def select_word():
    file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if file: word_path.set(file)

def select_folder():
    folder = filedialog.askdirectory()
    if folder: folder_path.set(folder)

def clear_all():
    word_path.set("")
    folder_path.set("")
    status_text.set("Ready to generate...")
    status_label.configure(text_color=GUI_TEXT)
    progress.set(0)
    
    log_box.configure(state="normal")
    log_box.delete("1.0", "end")
    log_box.configure(state="disabled")

# -------------------------
# Main Process
# -------------------------
def run_process():
    # เพิ่มการเช็ค excel_path เข้าไป
    if not word_path.get() or not folder_path.get() or not excel_path.get():
        status_text.set("Please select Word, Excel, and Image folder.")
        status_label.configure(text_color="#FEE75C") 
        return
    
    log_box.configure(state="normal")
    log_box.delete("1.0", "end")
    log_box.configure(state="disabled")
    
    btn_run.configure(state="disabled", text="Running...", fg_color=GUI_MUTED)
    btn_clear.configure(state="disabled", fg_color=GUI_MUTED)
    progress.set(0)

    log_filename = "history_log.txt"
    try:
        with open(log_filename, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*50}\n")
            f.write(f"NEW RUN STARTED: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Word Template: {word_path.get()}\n")
            f.write(f"{'='*50}\n")
    except Exception:
        pass 

    def update_progress(val): progress.set(val)
    
    def update_status(text, state="normal"):
        status_text.set(text)
        if state == "error": status_label.configure(text_color="#ED4245")
        elif state == "success": status_label.configure(text_color=GUI_GREEN)
        else: status_label.configure(text_color=GUI_TEXT)

    def write_log(text):
        log_box.configure(state="normal")
        log_box.insert("end", text + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")
        try:
            with open(log_filename, "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except:
            pass

    def on_finish():
        btn_run.configure(state="normal", text="Run Generator", fg_color=GUI_BLURPLE)
        btn_clear.configure(state="normal", fg_color=GUI_RED)

    threading.Thread(
        target=generate_report, 
        args=(word_path.get(), excel_path.get(), folder_path.get(), update_progress, update_status, write_log, on_finish),
        daemon=True
    ).start()

# -------------------------
# UI Layout: ฝั่งซ้าย
# -------------------------
left_frame = ctk.CTkFrame(app, fg_color=GUI_DARK, corner_radius=5)
left_frame.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=20)

lbl_word = ctk.CTkLabel(left_frame, text="Word Template File", font=("Sarabun", 14, "bold"), text_color=GUI_MUTED)
lbl_word.pack(anchor="w", padx=20, pady=(15, 2))
frame_word = ctk.CTkFrame(left_frame, fg_color="transparent")
frame_word.pack(fill="x", padx=20)
entry_word = ctk.CTkEntry(frame_word, textvariable=word_path, fg_color=GUI_DARKER, border_width=0, state="disabled")
entry_word.pack(side="left", fill="x", expand=True, padx=(0, 10))
btn_word = ctk.CTkButton(frame_word, text="Browse", width=80, fg_color=GUI_DARKER, command=select_word)
btn_word.pack(side="right")

lbl_excel = ctk.CTkLabel(left_frame, text="Excel Test Script File", font=("Helvetica", 14, "bold"), text_color=GUI_MUTED)
lbl_excel.pack(anchor="w", padx=20, pady=(15, 2))
frame_excel = ctk.CTkFrame(left_frame, fg_color="transparent")
frame_excel.pack(fill="x", padx=20)
entry_excel = ctk.CTkEntry(frame_excel, textvariable=excel_path, fg_color=GUI_DARKER, border_width=0, state="disabled")
entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))
btn_excel = ctk.CTkButton(frame_excel, text="Browse", width=80, fg_color=GUI_DARKER, command=select_excel)
btn_excel.pack(side="right")

lbl_folder = ctk.CTkLabel(left_frame, text="Image Folder (Test Scripts)", font=("Helvetica", 14, "bold"), text_color=GUI_MUTED)
lbl_folder.pack(anchor="w", padx=20, pady=(15, 2))
frame_folder = ctk.CTkFrame(left_frame, fg_color="transparent")
frame_folder.pack(fill="x", padx=20)
entry_folder = ctk.CTkEntry(frame_folder, textvariable=folder_path, fg_color=GUI_DARKER, border_width=0, state="disabled")
entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))
btn_folder = ctk.CTkButton(frame_folder, text="Browse", width=80, fg_color=GUI_DARKER, command=select_folder)
btn_folder.pack(side="right")

spacer = ctk.CTkLabel(left_frame, text="")
spacer.pack(fill="y", expand=True)

progress = ctk.CTkProgressBar(left_frame, height=12, fg_color=GUI_DARKER, progress_color=GUI_BLURPLE)
progress.set(0)
progress.pack(fill="x", padx=20, pady=(10, 5))

status_label = ctk.CTkLabel(left_frame, textvariable=status_text, font=("Helvetica", 12), text_color=GUI_TEXT)
status_label.pack(pady=(0, 5))

btn_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
btn_frame.pack(pady=(5, 20), padx=20, fill="x")

btn_run = ctk.CTkButton(btn_frame, text="Run Generator", height=40, font=("Helvetica", 15, "bold"), fg_color=GUI_BLURPLE, hover_color=GUI_BLURPLE_HOVER, command=run_process)
btn_run.pack(side="left", fill="x", expand=True, padx=(0, 10))

btn_clear = ctk.CTkButton(btn_frame, text="Clear", width=80, height=40, font=("Helvetica", 14, "bold"), fg_color=GUI_RED, hover_color=GUI_RED_HOVER, command=clear_all)
btn_clear.pack(side="right")

btn_tools = ctk.CTkButton(left_frame, text="🛠️ Create Test Folders", height=35, font=("Helvetica", 13, "bold"), fg_color=GUI_DARKER, text_color=GUI_TEXT, hover_color=GUI_BG, command=lambda: open_folder_popup(app))
btn_tools.pack(fill="x", padx=20, pady=(0, 20))

# -------------------------
# UI Layout: ฝั่งขวา
# -------------------------
right_frame = ctk.CTkFrame(app, fg_color=GUI_DARK, corner_radius=10)
right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=20)

lbl_log = ctk.CTkLabel(right_frame, text="Execution Log", font=("Helvetica", 14, "bold"), text_color=GUI_MUTED)
lbl_log.pack(anchor="w", padx=20, pady=(15, 2))

log_box = ctk.CTkTextbox(right_frame, fg_color=GUI_DARKER, text_color=GUI_MUTED, font=("Consolas", 12))
log_box.pack(fill="both", expand=True, padx=20, pady=(0, 20))
log_box.configure(state="disabled")

app.mainloop()
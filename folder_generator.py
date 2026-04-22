import os
from tkinter import filedialog, StringVar, messagebox
import customtkinter as ctk

# นำเข้าสีธีมเดียวกับหน้าหลัก
GUI_BG, GUI_DARKER = "#36393F", "#202225"
GUI_GREEN, GUI_TEXT, GUI_MUTED = "#57F287", "#DCDDDE", "#8E9297"

def open_folder_popup(parent_app):
    popup = ctk.CTkToplevel(parent_app)
    popup.title("Auto Create Test Folders")
    popup.geometry("520x350")
    popup.transient(parent_app) # ลอยเหนือหน้าต่างหลัก
    popup.grab_set()            # บังคับคลิกหน้านี้ก่อน
    popup.configure(fg_color=GUI_BG)

    target_path = StringVar()

    def select_target():
        folder = filedialog.askdirectory()
        if folder: target_path.set(folder)

    def do_create():
        base = entry_base.get()
        try:
            start_num = int(entry_start.get())
            end_num = int(entry_end.get())
            pad_num = int(entry_pad.get())
        except ValueError:
            messagebox.showerror("Error", "ช่อง Start, End และ Zero Padding ต้องเป็นตัวเลขเท่านั้นครับ!")
            return

        path = target_path.get()
        if not path:
            messagebox.showerror("Error", "กรุณาเลือกตำแหน่งที่จะสร้างโฟลเดอร์ (Save Location) ด้วยครับ!")
            return

        try:
            created_count = 0
            for i in range(start_num, end_num + 1):
                folder_name = f"{base}{str(i).zfill(pad_num)}" 
                full_path = os.path.join(path, folder_name)
                os.makedirs(full_path, exist_ok=True)
                created_count += 1
            
            messagebox.showinfo("Success", f"สร้างโฟลเดอร์สำเร็จจำนวน {created_count} โฟลเดอร์!")
            popup.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด:\n{str(e)}")

    # --- UI Layout ---
    ctk.CTkLabel(popup, text="Save Location (ตำแหน่งที่เก็บโฟลเดอร์):", font=("Helvetica", 14, "bold"), text_color=GUI_MUTED).pack(anchor="w", padx=20, pady=(20, 5))
    frame_path = ctk.CTkFrame(popup, fg_color="transparent")
    frame_path.pack(fill="x", padx=20)
    ctk.CTkEntry(frame_path, textvariable=target_path, fg_color=GUI_DARKER, border_width=0, state="disabled").pack(side="left", fill="x", expand=True, padx=(0, 10))
    ctk.CTkButton(frame_path, text="Browse", width=80, fg_color=GUI_DARKER, command=select_target).pack(side="right")

    ctk.CTkLabel(popup, text="Base Name (เช่น TS_IB_AU_MAK-):", font=("Helvetica", 14, "bold"), text_color=GUI_MUTED).pack(anchor="w", padx=20, pady=(15, 5))
    entry_base = ctk.CTkEntry(popup, fg_color=GUI_DARKER, border_width=0)
    entry_base.insert(0, "TS_IB_AU_MAK-")
    entry_base.pack(fill="x", padx=20)

    frame_nums = ctk.CTkFrame(popup, fg_color="transparent")
    frame_nums.pack(fill="x", padx=20, pady=15)

    ctk.CTkLabel(frame_nums, text="Start No:", text_color=GUI_TEXT).grid(row=0, column=0, sticky="w", padx=(0, 5))
    entry_start = ctk.CTkEntry(frame_nums, width=60, fg_color=GUI_DARKER, border_width=0)
    entry_start.insert(0, "1")
    entry_start.grid(row=0, column=1, padx=(0, 15))

    ctk.CTkLabel(frame_nums, text="End No:", text_color=GUI_TEXT).grid(row=0, column=2, sticky="w", padx=(0, 5))
    entry_end = ctk.CTkEntry(frame_nums, width=60, fg_color=GUI_DARKER, border_width=0)
    entry_end.insert(0, "10")
    entry_end.grid(row=0, column=3, padx=(0, 15))

    ctk.CTkLabel(frame_nums, text="Zero Padding:", text_color=GUI_TEXT).grid(row=0, column=4, sticky="w", padx=(0, 5))
    entry_pad = ctk.CTkEntry(frame_nums, width=60, fg_color=GUI_DARKER, border_width=0)
    entry_pad.insert(0, "4")
    entry_pad.grid(row=0, column=5)

    ctk.CTkButton(popup, text="Generate Folders", height=40, font=("Helvetica", 14, "bold"), fg_color=GUI_GREEN, text_color=GUI_DARKER, hover_color="#3BA55D", command=do_create).pack(pady=(10, 20), fill="x", padx=20)
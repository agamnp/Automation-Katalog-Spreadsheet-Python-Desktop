import os
import importlib.util
import inspect
import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading

MODULE_FOLDER = "Module"

def gui_logger(msg):
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)

def run_in_thread(func):
    def wrapper():
        try:
            log_text.insert(tk.END, f"▶️ Menjalankan: {func.__name__}...\n")
            func(logger=gui_logger)
            log_text.insert(tk.END, f"✅ Selesai: {func.__name__}\n\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            log_text.insert(tk.END, f"❌ Error: {e}\n\n")
    threading.Thread(target=wrapper, daemon=True).start()

def load_main_functions():
    function_list = []
    for filename in os.listdir(MODULE_FOLDER):
        if filename.endswith(".py") and not filename.startswith("_"):
            filepath = os.path.join(MODULE_FOLDER, filename)
            module_name = filename[:-3]  # remove .py

            spec = importlib.util.spec_from_file_location(module_name, filepath)
            module = importlib.util.module_from_spec(spec)
            try:
                spec.loader.exec_module(module)
            except Exception as e:
                print(f"⚠️ Gagal load modul {module_name}: {e}")
                continue

            for name, func in inspect.getmembers(module, inspect.isfunction):
                if name.startswith("main_"):
                    display_name = name.replace("main_", "").replace("_", " ").title()
                    function_list.append((display_name, func))
    return function_list

# === GUI Layout ===
def center_window(root, width=900, height=600):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
root = tk.Tk()
root.title("📊 Automasi Google Sheet")

# Atur ukuran window yang lebih besar dan minimal
root.geometry("900x600")
root.minsize(800, 600)

# Label judul
title_label = tk.Label(root, text="🛠 Pilih Fungsi Automasi", font=("Helvetica", 16, "bold"))
title_label.pack(pady=15)

# Frame tombol dengan scrollbar jika banyak tombol
btn_frame = tk.Frame(root)
btn_frame.pack(fill='x', padx=20)

function_buttons = load_main_functions()

if not function_buttons:
    no_func_label = tk.Label(root, text="❌ Tidak ada fungsi 'main_' ditemukan di folder 'Module/'", font=("Helvetica", 12), fg="red")
    no_func_label.pack(pady=20)
else:
    for label, func in function_buttons:
        btn = tk.Button(btn_frame, text=label, width=30, command=lambda f=func: run_in_thread(f),
                        font=("Helvetica", 12), bg="#2E8B57", fg="white", relief='raised', bd=3)
        btn.pack(pady=6, anchor='w')

# Text log dengan scrollbar, font monospace untuk rapih
log_text = scrolledtext.ScrolledText(root, width=100, height=25, font=("Courier New", 11))
log_text.pack(padx=15, pady=15, fill='both', expand=True)

root.mainloop()

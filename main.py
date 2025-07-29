import os
import importlib.util
import inspect
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox

import globals as g
from globals import stop_requested

MODULE_FOLDER = "Module"

# ========================== UTILITIES ==========================

def gui_logger(msg):
    log_text.insert(tk.END, msg + "\n")
    log_text.see(tk.END)

def run_in_thread(func):
    def wrapper():
        try:
            g.stop_requested = False
            stop_btn.config(state=tk.NORMAL)
            gui_logger(f"▶️ Menjalankan: {func.__name__}...")
            func(logger=gui_logger)
            gui_logger(f"✅ Selesai: {func.__name__}\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            gui_logger(f"❌ Error: {e}\n")
        finally:
            stop_btn.config(state=tk.DISABLED)
    threading.Thread(target=wrapper, daemon=True).start()

def request_stop():
    g.stop_requested = True
    gui_logger("⏹️ Permintaan STOP diterima...")

def load_main_functions():
    function_list = []
    for filename in os.listdir(MODULE_FOLDER):
        if filename.endswith(".py") and not filename.startswith("_"):
            filepath = os.path.join(MODULE_FOLDER, filename)
            module_name = filename[:-3]

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

def center_window(root, width=900, height=600):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

# ========================== MAIN GUI ==========================

def main():
    global log_text, stop_btn

    root = tk.Tk()
    root.title("📊 Automasi Google Sheet")
    center_window(root)
    root.minsize(800, 600)

    # Title
    title_label = tk.Label(root, text="🛠 Pilih Fungsi Automasi", font=("Helvetica", 16, "bold"))
    title_label.pack(pady=15)

    # Frame untuk tombol-tombol
    btn_frame = tk.Frame(root)
    btn_frame.pack(fill='x', padx=20)

    # Load semua fungsi
    function_buttons = load_main_functions()
    if not function_buttons:
        no_func_label = tk.Label(
            root,
            text="❌ Tidak ada fungsi 'main_' ditemukan di folder 'Module/'",
            font=("Helvetica", 12),
            fg="red"
        )
        no_func_label.pack(pady=20)
    else:
        for label, func in function_buttons:
            btn = tk.Button(
                btn_frame, text=label, width=30, command=lambda f=func: run_in_thread(f),
                font=("Helvetica", 12), bg="#2E8B57", fg="white", relief='raised', bd=3
            )
            btn.pack(pady=6, anchor='w')

    # Tombol STOP
    stop_btn = tk.Button(
        btn_frame, text="⛔ STOP", width=30, command=request_stop,
        font=("Helvetica", 12, "bold"), bg="#B22222", fg="white", relief='raised', bd=3,
        state=tk.DISABLED
    )
    stop_btn.pack(pady=10, anchor='w')

    # Area log
    log_text = scrolledtext.ScrolledText(root, width=100, height=25, font=("Courier New", 11))
    log_text.pack(padx=15, pady=15, fill='both', expand=True)

    root.mainloop()

# ========================== ENTRY POINT ==========================

if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import ttk, scrolledtext
import threading

from globals import get_stop_requested, set_stop_requested
from Module import fungsi_tampilansheet  # pastikan path-nya sesuai

def center_window(root, width=600, height=400):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

# Fungsi pemanggil utama
def show_window(root):
    def run_main_function():
        center_window(root)
        set_stop_requested(False)
        log_text.configure(state="normal")
        log_text.delete(1.0, tk.END)
        log_text.configure(state="disabled")

        def logger(msg):
            log_text.configure(state="normal")
            log_text.insert(tk.END, msg + "\n")
            log_text.see(tk.END)
            log_text.configure(state="disabled")

        def target():
            try:
                fungsi_tampilansheet.main_tampilan_sheet(logger=logger)
            finally:
                run_button.config(state="normal")
                stop_button.config(state="disabled")

        thread = threading.Thread(target=target)
        thread.start()

        run_button.config(state="disabled")
        stop_button.config(state="normal")

    def stop_process():
        set_stop_requested(True)
        stop_button.config(state="disabled")

    window = tk.Toplevel()
    window.title("🧾 Tampilan Sheet")
    window.geometry("700x500")
    window.configure(bg="white")
    # Sembunyikan root utama
    root.withdraw()
     # Tampilkan root saat window ditutup
    def on_close():
        window.destroy()
        root.deiconify()
        
    window.protocol("WM_DELETE_WINDOW", on_close)
#==============================================Layout===============================
    title = tk.Label(window, text="🧾 Tampilan Sheet", font=("Helvetica", 16, "bold"), bg="white")
    title.pack(pady=10)

    button_frame = tk.Frame(window, bg="white")
    button_frame.pack(pady=5)

    global run_button, stop_button
    run_button = tk.Button(button_frame, text="▶ Jalankan", width=20, bg="#2E8B57", fg="white", command=run_main_function)
    run_button.grid(row=0, column=0, padx=5)

    stop_button = tk.Button(button_frame, text="⛔ STOP", width=20, bg="#B22222", fg="white", command=stop_process)
    stop_button.grid(row=0, column=1, padx=5)
    stop_button.config(state="disabled")

    log_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, height=20)
    log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    log_text.configure(state="disabled")

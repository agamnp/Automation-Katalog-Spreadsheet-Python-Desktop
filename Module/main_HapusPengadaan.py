import tkinter as tk
from tkinter import scrolledtext
import threading
from Module import fungsi_hapuspengadaan  # nama file kamu (tanpa .py), pastikan sesuai

def show_window(root):
    def run_hapus_pengadaan():
        text_log.configure(state="normal")
        text_log.delete(1.0, tk.END)
        text_log.configure(state="disabled")

        def logger(msg):
            text_log.configure(state="normal")
            text_log.insert(tk.END, msg + "\n")
            text_log.see(tk.END)
            text_log.configure(state="disabled")

        def target():
            try:
                fungsi_hapuspengadaan.main_hapus_pengadaan(logger=logger)
            finally:
                run_button.config(state="normal")

        threading.Thread(target=target).start()
        run_button.config(state="disabled")

    # Buat window baru
    window = tk.Toplevel()
    window.title("🗑️ Hapus Data Pengadaan")
    window.geometry("700x400")
    window.configure(bg="white")
    root.withdraw()

    def on_close():
        window.destroy()
        root.deiconify()

    window.protocol("WM_DELETE_WINDOW", on_close)

    title = tk.Label(window, text="🗑️ Hapus Data Pengadaan", font=("Helvetica", 16, "bold"), bg="white")
    title.pack(pady=10)

    global run_button
    run_button = tk.Button(window, text="🔍 Pilih File & Hapus", bg="#B22222", fg="white", font=("Helvetica", 12), command=run_hapus_pengadaan)
    run_button.pack(pady=10)

    text_log = scrolledtext.ScrolledText(window, wrap=tk.WORD, height=15)
    text_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    text_log.configure(state="disabled")

import os
import importlib.util
import inspect
import tkinter as tk

MODULE_FOLDER = "Module"

def load_function_windows():
    functions = []
    for filename in os.listdir(MODULE_FOLDER):
        if filename.endswith(".py") and not filename.startswith("_"):
            filepath = os.path.join(MODULE_FOLDER, filename)
            module_name = filename[:-3]

            spec = importlib.util.spec_from_file_location(module_name, filepath)
            module = importlib.util.module_from_spec(spec)
            try:
                spec.loader.exec_module(module)
            except Exception as e:
                print(f"❌ Gagal load modul {module_name}: {e}")
                continue

            for name, func in inspect.getmembers(module, inspect.isfunction):
                if name == "show_window":
                    display_name = module_name.replace("main_", "").replace("_", " ").title()
                    functions.append((display_name, func))
    return functions

def center_window(root, width=600, height=400):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

def main():
    root = tk.Tk()
    root.title("📊 Automasi Google Sheet")
    center_window(root)
    root.configure(bg="white")

    # Title
    title_label = tk.Label(root, text="🛠 Pilih Fungsi Automasi", font=("Helvetica", 16, "bold"), bg="white")
    title_label.pack(pady=20)

    # Frame tombol
    btn_frame = tk.Frame(root, bg="white")
    btn_frame.pack(pady=10)

    # Load fungsi yang punya show_window
    functions = load_function_windows()
    if not functions:
        tk.Label(root, text="❌ Tidak ada fungsi ditemukan.", fg="red", bg="white").pack(pady=10)
    else:
        for label, func in functions:
            # Gunakan lambda untuk kirim root ke fungsi show_window
            btn = tk.Button(
                btn_frame, text=label, width=30, font=("Helvetica", 12),
                command=lambda f=func: f(root), bg="#2E8B57", fg="white"
            )
            btn.pack(pady=6)

    root.mainloop()

if __name__ == "__main__":
    main()

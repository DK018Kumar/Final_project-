import tkinter as tk

from gui import SubstationApp


if __name__ == "__main__":
    root = tk.Tk()
    # Basic High DPI awareness for clearer text on Windows
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = SubstationApp(root)
    root.mainloop()

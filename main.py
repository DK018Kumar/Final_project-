import tkinter as tk

from gui import SubstationApp


if __name__ == "__main__":
    # Basic High DPI awareness for clearer text on Windows
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    app = SubstationApp(root)
    root.mainloop()

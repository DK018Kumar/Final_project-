import tkinter as tk

from gui import SubstationApp


def main() -> None:
    root = tk.Tk()

    # Basic High DPI awareness for clearer text on Windows
    try:
        from ctypes import windll  # type: ignore

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    SubstationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

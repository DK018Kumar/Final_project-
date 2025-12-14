import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import visio_engine


class SubstationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("System Architecture Generator")
        self.root.geometry("1000x750")

        self.engine = visio_engine.VisioEngine()
        # list of dicts {'name':..., 'props':[...] }
        self.loaded_masters = []
        # lookup name -> props list
        self.masters_by_name = {}
        # matrix structure holds rows of dicts {'combobox':..., 'props_frame':...}
        self.matrix_widgets = []

        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        # Notebook with two sections (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.tab_central = ttk.Frame(self.notebook)
        self.tab_other = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_central, text="Central Substation")
        self.notebook.add(self.tab_other, text="Other Substations")

        self._build_central_tab()
        self._build_other_substations_tab()

    # ------------------------
    # Central tab
    # ------------------------
    def _build_central_tab(self):
        top = tk.Frame(self.tab_central, padx=10, pady=10)
        top.pack(fill="x")

        tk.Button(
            top, text="Load Substation Database", command=self.on_select_stencil
        ).grid(row=0, column=0, sticky="w")

        self.lbl_stencil = tk.Label(top, text="No database loaded", fg="gray")
        self.lbl_stencil.grid(row=0, column=1, sticky="w", padx=10)

        tk.Label(top, text="Central Substation:").grid(row=1, column=0, sticky="w", pady=10)
        self.cb_central = ttk.Combobox(top, state="readonly", width=60)
        self.cb_central.grid(row=1, column=1, sticky="w", padx=6)

        hint = tk.Label(
            self.tab_central,
            text="Next: go to 'Other Substations' to build the matrix layout.",
            fg="gray",
            padx=10,
            pady=6,
            anchor="w",
        )
        hint.pack(fill="x")

    # ------------------------
    # Other substations tab (matrix builder)
    # ------------------------
    def _build_other_substations_tab(self):
        controls = tk.Frame(self.tab_other, padx=10, pady=10)
        controls.pack(fill="x")

        tk.Label(controls, text="Rows:").grid(row=0, column=0, sticky="w")
        self.entry_rows = tk.Entry(controls, width=6)
        self.entry_rows.insert(0, "2")
        self.entry_rows.grid(row=0, column=1, sticky="w")

        tk.Label(controls, text="Cols:").grid(row=0, column=2, sticky="w", padx=(18, 0))
        self.entry_cols = tk.Entry(controls, width=6)
        self.entry_cols.insert(0, "3")
        self.entry_cols.grid(row=0, column=3, sticky="w")

        tk.Button(controls, text="Build Matrix", command=self.build_matrix).grid(
            row=0, column=4, padx=12
        )

        tk.Button(
            controls,
            text="Generate Visio Layout",
            bg="green",
            fg="white",
            command=self.on_generate,
        ).grid(row=1, column=0, columnspan=5, pady=12, sticky="ew")

        # Scrollable area for matrix
        container = tk.Frame(self.tab_other)
        container.pack(fill="both", expand=True, padx=10, pady=6)

        # Canvas for scrolling
        self.canvas = tk.Canvas(container, borderwidth=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Scrollbars
        v_scroll = tk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        v_scroll.pack(side="right", fill="y")
        h_scroll = tk.Scrollbar(self.tab_other, orient="horizontal", command=self.canvas.xview)
        h_scroll.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        # Frame inside canvas
        self.matrix_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.matrix_frame, anchor="nw"
        )

        # Bind events for scrolling region updates and mousewheel
        self.matrix_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mousewheel support (Windows, Mac, Linux)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel_windows)  # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_unix)  # Linux scroll up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_unix)  # Linux scroll down

    def _on_frame_configure(self, event):
        # update scrollregion to include new size
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_canvas_configure(self, event):
        # match inner frame width optionally
        try:
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        except Exception:
            pass

    def _on_mousewheel_windows(self, event):
        # For vertical scrolling; event.delta is multiple of 120 on Windows
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _on_mousewheel_unix(self, event):
        # For some Linux setups using Button-4/5
        try:
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        except Exception:
            pass

    # ------------------------
    # Stencil loading
    # ------------------------
    def on_select_stencil(self):
        path = filedialog.askopenfilename(
            filetypes=[("Visio Stencil", "*.vssx *.vss *.vstx")]
        )
        if not path:
            return
        threading.Thread(target=self._thread_load_stencil, args=(path,), daemon=True).start()

    def _thread_load_stencil(self, path):
        try:
            masters = self.engine.load_stencil_masters(path)
            self.root.after(0, lambda: self._after_load_stencil(masters))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def _after_load_stencil(self, masters):
        # masters: list of dicts {'name':..., 'props': [...]}
        self.loaded_masters = masters
        self.masters_by_name = {m["name"]: m.get("props", []) for m in masters}
        names = list(self.masters_by_name.keys())

        self.lbl_stencil.config(text="Database Loaded", fg="black")
        self.cb_central["values"] = names
        if names:
            self.cb_central.current(0)

        # If matrix is already built, refresh its combobox values
        self._refresh_matrix_combobox_values()

        messagebox.showinfo("Success", f"Loaded {len(names)} masters.")

    def _refresh_matrix_combobox_values(self):
        values = ["None"] + list(self.masters_by_name.keys())
        for row in self.matrix_widgets:
            for cell in row:
                cb = cell.get("combobox")
                if not cb:
                    continue
                current = cb.get() or "None"
                cb["values"] = values
                if current in values:
                    cb.set(current)
                else:
                    cb.current(0)

    # ------------------------
    # Build matrix UI
    # ------------------------
    def build_matrix(self):
        # destroy previous widgets
        for w in self.matrix_frame.winfo_children():
            w.destroy()
        self.matrix_widgets = []

        try:
            rows = int(self.entry_rows.get())
            cols = int(self.entry_cols.get())
            if rows <= 0 or cols <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "Rows and Cols must be positive integers.")
            return

        values = ["None"] + list(self.masters_by_name.keys())

        # Build grid: each cell is a small frame with combobox + props below
        for r in range(rows):
            row_container = tk.Frame(self.matrix_frame)
            row_container.grid(row=r, column=0, sticky="w", pady=6)

            row_widgets = []
            for c in range(cols):
                cell_frame = tk.Frame(row_container, relief="ridge", bd=1, padx=6, pady=6)
                cell_frame.pack(side="left", padx=8, pady=4)

                tk.Label(cell_frame, text=f"({r},{c})", anchor="w").pack(anchor="w")

                cb = ttk.Combobox(cell_frame, state="readonly", width=36, values=values)
                cb.pack(anchor="w", pady=(4, 2))
                cb.current(0)

                props_frame = tk.Frame(cell_frame)
                props_frame.pack(anchor="w", fill="x", pady=(4, 0))

                # bind selection event
                cb.bind(
                    "<<ComboboxSelected>>",
                    lambda e, cb=cb, pf=props_frame: self._on_master_selected(cb, pf),
                )

                row_widgets.append({"combobox": cb, "props_frame": props_frame})
            self.matrix_widgets.append(row_widgets)

        # update scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_master_selected(self, combobox, props_frame):
        # clear existing children
        for w in props_frame.winfo_children():
            w.destroy()

        name = combobox.get()
        if not name or name == "None":
            return

        props = self.masters_by_name.get(name, [])
        if not props:
            tk.Label(props_frame, text="(no shape data found)", fg="gray").pack(anchor="w")
            return

        # Header row
        header = tk.Frame(props_frame)
        header.pack(fill="x", anchor="w")
        tk.Label(header, text="Type", font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        tk.Label(header, text="Value", font=("Arial", 9, "bold")).grid(
            row=0, column=1, sticky="w"
        )

        # Each property
        for p in props:
            t = p.get("type", "")
            v = p.get("value", "")
            row = tk.Frame(props_frame)
            row.pack(fill="x", anchor="w", pady=1)
            tk.Label(row, text=t, anchor="w", width=20).grid(
                row=0, column=0, sticky="w", padx=(0, 8)
            )
            tk.Label(row, text=v, anchor="w", wraplength=360, justify="left").grid(
                row=0, column=1, sticky="w"
            )

    # ------------------------
    # Generate layout in Visio
    # ------------------------
    def on_generate(self):
        if not self.engine.current_stencil_path:
            messagebox.showwarning("No database", "Load the substation database first.")
            # take user to central tab
            try:
                self.notebook.select(self.tab_central)
            except Exception:
                pass
            return

        central = self.cb_central.get()
        if not central:
            messagebox.showwarning("Select Central", "Choose a central substation.")
            try:
                self.notebook.select(self.tab_central)
            except Exception:
                pass
            return

        matrix = []
        for row in self.matrix_widgets:
            matrix.append([cell["combobox"].get() for cell in row])

        data = {"central": central, "matrix_subs": matrix}
        threading.Thread(target=self._thread_generate, args=(data,), daemon=True).start()

    def _thread_generate(self, data):
        try:
            self.engine.generate_layout(data)
            self.root.after(0, lambda: messagebox.showinfo("Success", "Visio layout generated!"))
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def on_close(self):
        try:
            self.root.destroy()
        finally:
            sys.exit()

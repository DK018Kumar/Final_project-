import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import copy

import visio_engine


class SubstationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Substation Layout Generator")
        self.root.geometry("1000x750")

        self.engine = visio_engine.VisioEngine()
        # list of dicts {'name':..., 'props':[...] }
        self.loaded_masters = []
        # lookup name -> props list
        self.masters_by_name = {}
        # matrix structure holds rows of dicts {'combobox':..., 'props_frame':...}
        self.matrix_widgets = []
        self.live_visio_preview_enabled = tk.BooleanVar(value=True)

        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        # Top controls frame
        top = tk.Frame(self.root, padx=10, pady=10)
        top.pack(fill="x")

        tk.Button(top, text="Select Stencil", command=self.on_select_stencil).grid(
            row=0, column=0, sticky="w"
        )
        self.lbl_stencil = tk.Label(top, text="No stencil selected", fg="gray")
        self.lbl_stencil.grid(row=0, column=1, sticky="w", padx=10)

        tk.Label(top, text="Central Substation:").grid(row=1, column=0, sticky="w", pady=8)
        self.cb_central = ttk.Combobox(top, state="readonly", width=60)
        self.cb_central.grid(row=1, column=1, sticky="w", padx=6)
        self.cb_central.bind("<<ComboboxSelected>>", self._on_central_selected)

        tk.Label(top, text="Central Label:").grid(row=2, column=0, sticky="w")
        self.entry_central_label = tk.Entry(top, width=40)
        self.entry_central_label.grid(row=2, column=1, sticky="w", padx=6, pady=(0, 8))

        tk.Label(top, text="Rows:").grid(row=3, column=0, sticky="w")
        self.entry_rows = tk.Entry(top, width=6)
        self.entry_rows.insert(0, "2")
        self.entry_rows.grid(row=3, column=1, sticky="w")

        tk.Label(top, text="Cols:").grid(row=3, column=1, sticky="e", padx=(0, 160))
        self.entry_cols = tk.Entry(top, width=6)
        self.entry_cols.insert(0, "3")
        self.entry_cols.grid(row=3, column=1, sticky="e", padx=(0, 60))

        tk.Button(top, text="Build Matrix", command=self.build_matrix).grid(row=3, column=2, padx=10)
        tk.Button(top, text="Generate Visio Layout", bg="green", fg="white", command=self.on_generate).grid(
            row=4, column=0, columnspan=3, pady=12, sticky="ew"
        )
        tk.Checkbutton(
            top,
            text="Live Visio Selection Preview",
            variable=self.live_visio_preview_enabled,
            onvalue=True,
            offvalue=False,
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # Replace selected shape controls (uses the loaded stencil masters)
        replace_frame = tk.LabelFrame(top, text="Replace Selected Shape in Visio", padx=8, pady=6)
        replace_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        tk.Label(replace_frame, text="New Master:").grid(row=0, column=0, sticky="w")
        self.cb_replace_master = ttk.Combobox(replace_frame, state="readonly", width=50)
        self.cb_replace_master.grid(row=0, column=1, sticky="w", padx=(6, 10))

        tk.Label(replace_frame, text="Label (optional):").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.entry_replace_label = tk.Entry(replace_frame, width=52)
        self.entry_replace_label.grid(row=1, column=1, sticky="w", padx=(6, 10), pady=(6, 0))

        tk.Button(replace_frame, text="Replace Selected", command=self.on_replace_selected).grid(
            row=0, column=2, rowspan=2, padx=(6, 0), sticky="ns"
        )

        # Shape data preview block
        preview_frame = tk.LabelFrame(self.root, text="Shape Data Preview", padx=10, pady=6)
        preview_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.shape_data_text = tk.Text(preview_frame, height=8, wrap="none", state="disabled")
        self.shape_data_text.pack(side="left", fill="both", expand=True)
        preview_scroll = tk.Scrollbar(preview_frame, orient="vertical", command=self.shape_data_text.yview)
        preview_scroll.pack(side="right", fill="y")
        self.shape_data_text.configure(yscrollcommand=preview_scroll.set)

        # Scrollable area for matrix
        container = tk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=10, pady=6)

        # Canvas for scrolling
        self.canvas = tk.Canvas(container, borderwidth=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Scrollbars
        v_scroll = tk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        v_scroll.pack(side="right", fill="y")
        h_scroll = tk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        h_scroll.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        # Frame inside canvas
        self.matrix_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.matrix_frame, anchor="nw")

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
        path = filedialog.askopenfilename(filetypes=[("Visio Stencil", "*.vssx *.vss *.vstx")])
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
        self.lbl_stencil.config(text="Stencil Loaded", fg="black")
        self.cb_central["values"] = names
        # replacement dropdown uses the same masters
        if hasattr(self, "cb_replace_master"):
            self.cb_replace_master["values"] = names
            if names:
                self.cb_replace_master.current(0)
        if names:
            self.cb_central.current(0)
            self._render_shape_data(f"Master: {names[0]}", self.masters_by_name.get(names[0], []))
        messagebox.showinfo("Success", f"Loaded {len(names)} masters.")
        if not names:
            self._render_shape_data("Master: (none)", [])

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

        # Build grid: each cell is a small frame with combobox + props below
        for r in range(rows):
            row_container = tk.Frame(self.matrix_frame)
            row_container.grid(row=r, column=0, sticky="w", pady=6)

            row_widgets = []
            for c in range(cols):
                cell_frame = tk.Frame(row_container, relief="ridge", bd=1, padx=6, pady=6)
                cell_frame.pack(side="left", padx=8, pady=4)

                tk.Label(cell_frame, text=f"({r},{c})", anchor="w").pack(anchor="w")

                cb = ttk.Combobox(
                    cell_frame,
                    state="readonly",
                    width=36,
                    values=["None"] + list(self.masters_by_name.keys()),
                )
                cb.pack(anchor="w", pady=(4, 2))
                cb.current(0)

                tk.Label(cell_frame, text="Custom Label:", anchor="w").pack(anchor="w")
                entry_label = tk.Entry(cell_frame, width=34)
                entry_label.pack(anchor="w", pady=(0, 4))

                props_frame = tk.Frame(cell_frame)
                props_frame.pack(anchor="w", fill="x", pady=(4, 0))

                cell_state = {
                    "combobox": cb,
                    "props_frame": props_frame,
                    "label_entry": entry_label,
                    "selected_master": None,
                    "selected_props": [],
                }

                # bind selection event (store props on the cell itself)
                cb.bind("<<ComboboxSelected>>", lambda e, cs=cell_state: self._on_master_selected(cs))

                row_widgets.append(cell_state)
            self.matrix_widgets.append(row_widgets)

        # update scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_master_selected(self, cell_state):
        combobox = cell_state["combobox"]
        props_frame = cell_state["props_frame"]
        label_entry = cell_state.get("label_entry")

        # clear existing children
        for w in props_frame.winfo_children():
            w.destroy()

        name = combobox.get()
        if not name or name == "None":
            cell_state["selected_master"] = None
            cell_state["selected_props"] = []
            return

        if label_entry is not None:
            current_text = label_entry.get().strip()
            if not current_text:
                label_entry.delete(0, tk.END)
                label_entry.insert(0, name)

        props = self.masters_by_name.get(name, [])
        # Store a snapshot so each cell keeps its own data (won't be overwritten)
        cell_state["selected_master"] = name
        cell_state["selected_props"] = copy.deepcopy(props)
        if not props:
            tk.Label(props_frame, text="(no shape data found)", fg="gray").pack(anchor="w")
            return

        # Header row
        header = tk.Frame(props_frame)
        header.pack(fill="x", anchor="w")
        tk.Label(header, text="Type", font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        tk.Label(header, text="Value", font=("Arial", 9, "bold")).grid(row=0, column=1, sticky="w")

        # Each property
        for p in props:
            t = p.get("type", "")
            v = p.get("value", "")
            row = tk.Frame(props_frame)
            row.pack(fill="x", anchor="w", pady=1)
            tk.Label(row, text=t, anchor="w", width=28).grid(row=0, column=0, sticky="w", padx=(0, 8))
            tk.Label(row, text=v, anchor="w", wraplength=360, justify="left").grid(
                row=0, column=1, sticky="w"
            )

    # ------------------------
    # Generate layout in Visio
    # ------------------------
    def on_generate(self):
        if not self.engine.current_stencil_path:
            messagebox.showwarning("No stencil", "Load a stencil first.")
            return

        central = self.cb_central.get()
        if not central:
            messagebox.showwarning("Select Central", "Choose a central substation.")
            return

        matrix = []
        for row in self.matrix_widgets:
            row_data = []
            for cell in row:
                row_data.append({"name": cell["combobox"].get(), "label": cell["label_entry"].get().strip()})
            matrix.append(row_data)

        central_label = self.entry_central_label.get().strip()
        data = {
            "central": central,
            "matrix_subs": matrix,
            "central_label": central_label,
            # maintain legacy key name for engine compatibility
            "substation_label": central_label,
        }
        threading.Thread(target=self._thread_generate, args=(data,), daemon=True).start()

    def _on_central_selected(self, event=None):
        name = self.cb_central.get()
        if name:
            self._render_shape_data(f"Master: {name}", self.masters_by_name.get(name, []))

    def _render_shape_data(self, title, props):
        if not hasattr(self, "shape_data_text") or self.shape_data_text is None:
            return
        lines = [title or "Shape Data", "-" * 40]
        if props:
            for prop in props:
                t = prop.get("type", "")
                v = prop.get("value", "")
                lines.append(f"{t or '(unnamed)'}: {v}")
        else:
            lines.append("(no shape data)")

        text_value = "\n".join(lines)
        self.shape_data_text.configure(state="normal")
        self.shape_data_text.delete("1.0", tk.END)
        self.shape_data_text.insert(tk.END, text_value)
        self.shape_data_text.configure(state="disabled")

    def _on_visio_selection_payload(self, payload):
        # Called from a COM event thread; marshal to Tk main thread.
        try:
            title = payload.get("title", "Selected: (unknown)")
            props = payload.get("props", [])
        except Exception:
            title = "Selected: (unknown)"
            props = []

        self.root.after(0, lambda: self._render_shape_data(title, props))

    def _thread_generate(self, data):
        try:
            self.engine.generate_layout(data)
            def _after():
                messagebox.showinfo("Success", "Visio layout generated!")
                if self.live_visio_preview_enabled.get():
                    # Start live preview: click shapes in Visio to see their data here.
                    self.engine.start_selection_watch(self._on_visio_selection_payload)

            self.root.after(0, _after)
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def on_replace_selected(self):
        # Replace whatever is currently selected in Visio
        if not self.engine.current_stencil_path:
            messagebox.showwarning("No stencil", "Load a stencil first.")
            return

        master = ""
        if hasattr(self, "cb_replace_master"):
            master = (self.cb_replace_master.get() or "").strip()
        if not master:
            messagebox.showwarning("Select master", "Choose a replacement master.")
            return

        label = ""
        if hasattr(self, "entry_replace_label"):
            label = (self.entry_replace_label.get() or "").strip()

        def _run():
            try:
                self.engine.replace_selected_shape(master, label_text=label or None)
                self.root.after(0, lambda: messagebox.showinfo("Success", "Selected shape replaced."))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

        threading.Thread(target=_run, daemon=True).start()

    def on_close(self):
        try:
            self.engine.stop_selection_watch()
        except Exception:
            pass
        self.root.destroy()
        sys.exit()

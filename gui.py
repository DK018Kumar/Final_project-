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
        self.root.geometry("1100x780")

        self.engine = visio_engine.VisioEngine()

        # list of dicts {'name':..., 'props':[...] }
        self.loaded_masters = []
        # lookup name -> props list
        self.masters_by_name = {}
        # matrix structure holds rows of dicts
        self.matrix_widgets = []
        # per-row apply settings var
        self._row_apply_vars = []

        # Preview state (shared across tabs)
        self._preview_cache = {}  # master_name -> png_path
        self._preview_photo = None
        self._preview_image_id = None
        self._preview_loading_for = None

        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        # Main split: left notebook, right preview (shared)
        self.main_pane = ttk.Panedwindow(self.root, orient="horizontal")
        self.main_pane.pack(fill="both", expand=True)

        self.left_frame = ttk.Frame(self.main_pane)
        self.right_frame = ttk.Frame(self.main_pane)

        self.main_pane.add(self.left_frame, weight=3)
        self.main_pane.add(self.right_frame, weight=2)

        # Notebook with two sections (tabs)
        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.pack(fill="both", expand=True)

        self.tab_central = ttk.Frame(self.notebook)
        self.tab_other = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_central, text="Central Substation")
        self.notebook.add(self.tab_other, text="Other Substations")

        self._build_central_tab()
        self._build_other_substations_tab()
        self._build_preview_pane()

    # ------------------------
    # Central tab
    # ------------------------
    def _build_central_tab(self):
        top = tk.Frame(self.tab_central, padx=10, pady=10)
        top.pack(fill="x")

        tk.Button(top, text="Load Substation Database", command=self.on_select_stencil).grid(
            row=0, column=0, sticky="w"
        )
        tk.Button(top, text="Inspect / Replace Shapes", command=self._open_inspector).grid(
            row=0, column=1, sticky="w", padx=10
        )

        self.lbl_stencil = tk.Label(top, text="No database loaded", fg="gray")
        self.lbl_stencil.grid(row=0, column=2, sticky="w", padx=10)

        tk.Label(top, text="Central Substation:").grid(row=1, column=0, sticky="w", pady=(10, 6))
        self.cb_central = ttk.Combobox(top, state="readonly", width=60)
        self.cb_central.grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=(10, 6))
        self.cb_central.bind(
            "<<ComboboxSelected>>",
            lambda e: self._request_preview(self.cb_central.get()),
        )

        # Clean label entry (user-provided) for Visio textbox above the central substation
        tk.Label(top, text="Enter name of substation:").grid(row=2, column=0, sticky="w")
        self.entry_central_label = tk.Entry(top, width=50)
        self.entry_central_label.grid(row=2, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 8))

        tk.Button(
            top,
            text="View element shape data",
            command=lambda: self._open_master_elements_window(self.cb_central.get()),
        ).grid(row=3, column=1, sticky="w", padx=6, pady=(0, 8))

        hint = tk.Label(
            self.tab_central,
            text="Tip: The name you enter is placed as a bold textbox above the dropped substation in Visio.",
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

        tk.Button(controls, text="Inspect / Replace Shapes", command=self._open_inspector).grid(
            row=0, column=5, padx=6
        )

        tk.Button(
            controls,
            text="Generate Visio Layout",
            bg="green",
            fg="white",
            command=self.on_generate,
        ).grid(row=1, column=0, columnspan=6, pady=12, sticky="ew")

        # Scrollable area for matrix
        container = tk.Frame(self.tab_other)
        container.pack(fill="both", expand=True, padx=10, pady=6)

        self.canvas = tk.Canvas(container, borderwidth=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        v_scroll = tk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        v_scroll.pack(side="right", fill="y")
        h_scroll = tk.Scrollbar(self.tab_other, orient="horizontal", command=self.canvas.xview)
        h_scroll.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.matrix_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.matrix_frame, anchor="nw")

        self.matrix_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mousewheel support (Windows, Mac, Linux)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_unix)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_unix)

    def _on_frame_configure(self, event):
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_canvas_configure(self, event):
        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_mousewheel_windows(self, event):
        try:
            delta_units = int(-1 * (event.delta / 120))
            if getattr(event, "state", 0) & 0x0001:
                self.canvas.xview_scroll(delta_units, "units")
            else:
                self.canvas.yview_scroll(delta_units, "units")
        except Exception:
            pass

    def _on_mousewheel_unix(self, event):
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
            filetypes=[("Visio Stencil", "*.vssx *.vss *.vstx *.vssm")]
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
        self.loaded_masters = masters
        self.masters_by_name = {m["name"]: m.get("props", []) for m in masters}
        names = list(self.masters_by_name.keys())

        # Clear preview cache on new load
        self._preview_cache = {}
        self._preview_loading_for = None
        try:
            self.preview_canvas.delete("all")
        except Exception:
            pass
        try:
            self.preview_status.config(text="Select a substation to preview.", fg="gray")
        except Exception:
            pass

        self.lbl_stencil.config(text="Database Loaded", fg="black")
        self.cb_central["values"] = names
        if names:
            self.cb_central.current(0)
            self._request_preview(names[0])

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
        for w in self.matrix_frame.winfo_children():
            w.destroy()
        self.matrix_widgets = []
        self._row_apply_vars = []

        try:
            rows = int(self.entry_rows.get())
            cols = int(self.entry_cols.get())
            if rows <= 0 or cols <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "Rows and Cols must be positive integers.")
            return

        values = ["None"] + list(self.masters_by_name.keys())

        for r in range(rows):
            row_container = tk.Frame(self.matrix_frame)
            row_container.grid(row=r, column=0, sticky="w", pady=8)

            # Row-level checkbox: apply settings across the row
            row_apply_var = tk.BooleanVar(value=False)
            self._row_apply_vars.append(row_apply_var)

            row_header = tk.Frame(row_container)
            row_header.pack(side="top", anchor="w", fill="x")
            tk.Checkbutton(
                row_header,
                text=f"Apply selections / names to whole row {r}",
                variable=row_apply_var,
            ).pack(side="left")

            cells_wrap = tk.Frame(row_container)
            cells_wrap.pack(side="top", anchor="w")

            row_widgets = []
            for c in range(cols):
                cell_frame = tk.Frame(cells_wrap, relief="ridge", bd=1, padx=6, pady=6)
                cell_frame.pack(side="left", padx=8, pady=4)

                tk.Label(cell_frame, text=f"({r},{c})", anchor="w").pack(anchor="w")

                cb = ttk.Combobox(cell_frame, state="readonly", width=36, values=values)
                cb.pack(anchor="w", pady=(4, 2))
                cb.current(0)

                tk.Label(cell_frame, text="Enter name of substation:").pack(anchor="w", pady=(4, 0))
                name_entry = tk.Entry(cell_frame, width=38)
                name_entry.pack(anchor="w", pady=(2, 4))

                props_frame = tk.Frame(cell_frame)
                props_frame.pack(anchor="w", fill="x", pady=(2, 0))

                # bind selection event
                cb.bind(
                    "<<ComboboxSelected>>",
                    lambda e, r=r, cb=cb, pf=props_frame: self._on_master_selected(r, cb, pf),
                )

                # propagate label across row if checkbox is ticked
                name_entry.bind(
                    "<FocusOut>",
                    lambda e, r=r, ent=name_entry: self._on_label_changed(r, ent),
                )

                tk.Button(
                    cell_frame,
                    text="View element shape data",
                    command=lambda cb=cb: self._open_master_elements_window(cb.get()),
                ).pack(anchor="w", pady=(6, 0))

                row_widgets.append({"combobox": cb, "props_frame": props_frame, "label_entry": name_entry})

            self.matrix_widgets.append(row_widgets)

        try:
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception:
            pass

    def _on_label_changed(self, row_index: int, src_entry: tk.Entry):
        try:
            if row_index < 0 or row_index >= len(self._row_apply_vars):
                return
            if not self._row_apply_vars[row_index].get():
                return
            val = src_entry.get()
            # Apply to all entries in the row
            for cell in self.matrix_widgets[row_index]:
                ent = cell.get("label_entry")
                if ent is src_entry:
                    continue
                try:
                    ent.delete(0, "end")
                    ent.insert(0, val)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_master_selected(self, row_index: int, combobox, props_frame, _propagating: bool = False):
        for w in props_frame.winfo_children():
            w.destroy()

        name = combobox.get()
        if not name or name == "None":
            return

        self._request_preview(name)

        props = self.masters_by_name.get(name, [])
        if not props:
            tk.Label(props_frame, text="(no shape data found)", fg="gray").pack(anchor="w")
        else:
            header = tk.Frame(props_frame)
            header.pack(fill="x", anchor="w")
            tk.Label(header, text="Type", font=("Arial", 9, "bold")).grid(
                row=0, column=0, sticky="w", padx=(0, 8)
            )
            tk.Label(header, text="Value", font=("Arial", 9, "bold")).grid(
                row=0, column=1, sticky="w"
            )

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

        # Row-level apply: propagate master selection across the row
        try:
            if _propagating:
                return
            if row_index < 0 or row_index >= len(self._row_apply_vars):
                return
            if not self._row_apply_vars[row_index].get():
                return

            target_name = combobox.get()
            for cell in self.matrix_widgets[row_index]:
                cb = cell.get("combobox")
                pf = cell.get("props_frame")
                if cb is combobox:
                    continue
                try:
                    cb.set(target_name)
                    self._on_master_selected(row_index, cb, pf, _propagating=True)
                except Exception:
                    pass
        except Exception:
            pass

    # ------------------------
    # Generate layout in Visio
    # ------------------------
    def on_generate(self):
        if not self.engine.current_stencil_path:
            messagebox.showwarning("No database", "Load the substation database first.")
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

        central_label = (self.entry_central_label.get() or "").strip()

        matrix = []
        for row in self.matrix_widgets:
            out_row = []
            for cell in row:
                master = (cell["combobox"].get() or "").strip()
                label = (cell["label_entry"].get() or "").strip()
                if not master or master == "None":
                    out_row.append(None)
                else:
                    out_row.append({"master": master, "label": label})
            matrix.append(out_row)

        data = {"central": central, "central_label": central_label, "matrix_subs": matrix}
        threading.Thread(target=self._thread_generate, args=(data,), daemon=True).start()

    def _thread_generate(self, data):
        try:
            self.engine.generate_layout(data)
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success",
                    "Visio layout generated!\n\nNote: Substations are automatically ungrouped after drop.\nUse 'Inspect / Replace Shapes' to read Shape Data and replace elements.",
                ),
            )
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def on_close(self):
        try:
            self.root.destroy()
        finally:
            sys.exit()

    # ------------------------
    # Shared preview pane
    # ------------------------
    def _build_preview_pane(self):
        header = tk.Frame(self.right_frame, padx=10, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="Preview", font=("Arial", 11, "bold")).pack(anchor="w")
        self.preview_status = tk.Label(
            header,
            text="Select a substation to preview.",
            fg="gray",
            anchor="w",
            justify="left",
        )
        self.preview_status.pack(fill="x", pady=(6, 0))

        body = tk.Frame(self.right_frame, padx=10, pady=10)
        body.pack(fill="both", expand=True)

        self.preview_canvas = tk.Canvas(body, borderwidth=1, relief="sunken", background="white")
        self.preview_canvas.pack(side="left", fill="both", expand=True)

        pv_scroll = tk.Scrollbar(body, orient="vertical", command=self.preview_canvas.yview)
        pv_scroll.pack(side="right", fill="y")
        ph_scroll = tk.Scrollbar(self.right_frame, orient="horizontal", command=self.preview_canvas.xview)
        ph_scroll.pack(side="bottom", fill="x")

        self.preview_canvas.configure(yscrollcommand=pv_scroll.set, xscrollcommand=ph_scroll.set)

        self.preview_canvas.bind("<ButtonPress-1>", self._preview_pan_start)
        self.preview_canvas.bind("<B1-Motion>", self._preview_pan_move)

    def _preview_pan_start(self, event):
        try:
            self.preview_canvas.scan_mark(event.x, event.y)
        except Exception:
            pass

    def _preview_pan_move(self, event):
        try:
            self.preview_canvas.scan_dragto(event.x, event.y, gain=1)
        except Exception:
            pass

    def _request_preview(self, master_name: str):
        name = (master_name or "").strip()
        if not name or name == "None":
            return

        cached_path = self._preview_cache.get(name)
        if cached_path:
            self._show_preview_image(name, cached_path)
            return

        if not self.engine.current_stencil_path:
            self.preview_status.config(text="Load the substation database to preview.", fg="gray")
            return

        if self._preview_loading_for == name:
            return
        self._preview_loading_for = name

        self.preview_status.config(text=f"Loading preview: {name}", fg="gray")
        threading.Thread(target=self._thread_preview, args=(name,), daemon=True).start()

    def _thread_preview(self, name: str):
        try:
            png_path = self.engine.render_master_preview(name)
            self.root.after(0, lambda: self._after_preview_ready(name, png_path))
        except Exception as e:
            self.root.after(0, lambda: self._after_preview_error(name, str(e)))

    def _after_preview_ready(self, name: str, png_path: str):
        self._preview_loading_for = None
        self._preview_cache[name] = png_path
        self._show_preview_image(name, png_path)

    def _after_preview_error(self, name: str, err: str):
        self._preview_loading_for = None
        self.preview_status.config(text=f"Preview failed for '{name}': {err}", fg="red")

    def _show_preview_image(self, name: str, png_path: str):
        try:
            img = tk.PhotoImage(file=png_path)
        except Exception as e:
            self.preview_status.config(text=f"Could not load preview image: {e}", fg="red")
            return

        self._preview_photo = img
        self.preview_canvas.delete("all")
        self._preview_image_id = self.preview_canvas.create_image(0, 0, anchor="nw", image=img)
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        self.preview_status.config(text=f"Preview: {name}", fg="black")

    # ------------------------
    # Master element Shape Data (easy view)
    # ------------------------
    def _open_master_elements_window(self, master_name: str):
        name = (master_name or "").strip()
        if not name or name == "None":
            return
        if not self.engine.current_stencil_path:
            messagebox.showwarning("No database", "Load the substation database first.")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Element Shape Data: {name}")
        win.geometry("720x520")

        tk.Label(
            win,
            text=(
                "This view shows the Shape Data inside the selected master (including sub-shapes).\n"
                "Use 'Inspect / Replace Shapes' to replace shapes in the active Visio document."
            ),
            fg="gray",
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=10, pady=(10, 6))

        text = tk.Text(win, wrap="none")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        text.insert("end", "Loading...\n")
        text.config(state="disabled")

        def worker():
            try:
                lines = self.engine.get_master_elements_shape_data(name)
                payload = "\n".join(lines) if lines else "(no shape data found)"
                self.root.after(0, lambda: _set_text(payload))
            except Exception as e:
                self.root.after(0, lambda: _set_text(f"Error: {e}"))

        def _set_text(s: str):
            text.config(state="normal")
            text.delete("1.0", "end")
            text.insert("end", s)
            text.config(state="disabled")

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------
    # Active Visio Inspector / Replacer
    # ------------------------
    def _open_inspector(self):
        win = tk.Toplevel(self.root)
        win.title("Inspect / Replace Shapes (Active Visio Document)")
        win.geometry("980x620")

        top = tk.Frame(win, padx=10, pady=10)
        top.pack(fill="x")

        status = tk.Label(top, text="Click 'Refresh' to read shapes from the active Visio document.", fg="gray")
        status.pack(side="left", fill="x", expand=True)

        tk.Button(top, text="Bring Visio to front", command=self._bring_visio_front).pack(side="right", padx=(6, 0))
        tk.Button(top, text="Refresh", command=lambda: self._refresh_shapes(status, tree, details)).pack(
            side="right"
        )

        body = tk.Frame(win, padx=10, pady=10)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body)
        left.pack(side="left", fill="both", expand=True)

        cols = ("page", "id", "name", "master")
        tree = ttk.Treeview(left, columns=cols, show="headings", height=16)
        tree.heading("page", text="Page")
        tree.heading("id", text="ShapeID")
        tree.heading("name", text="Shape")
        tree.heading("master", text="Master")
        tree.column("page", width=140)
        tree.column("id", width=70)
        tree.column("name", width=260)
        tree.column("master", width=260)
        tree.pack(fill="both", expand=True)

        tree.bind("<<TreeviewSelect>>", lambda e: self._load_shape_details(status, tree, details))

        right = tk.Frame(body)
        right.pack(side="right", fill="both", expand=True, padx=(10, 0))

        tk.Label(right, text="Shape Data (Label=Value)", font=("Arial", 10, "bold")).pack(anchor="w")

        details_frame = tk.Frame(right)
        details_frame.pack(fill="both", expand=True, pady=(6, 10))
        details_scroll = tk.Scrollbar(details_frame, orient="vertical")
        details = tk.Listbox(details_frame, activestyle="dotbox", yscrollcommand=details_scroll.set)
        details_scroll.config(command=details.yview)
        details.pack(side="left", fill="both", expand=True)
        details_scroll.pack(side="right", fill="y")
        details.insert("end", "Select a shape to view its Shape Data (Label=Value only).")

        repl = ttk.LabelFrame(right, text="Replace selected shape", padding=10)
        repl.pack(fill="x")

        tk.Label(repl, text="Stencil file:").grid(row=0, column=0, sticky="w")
        stencil_entry = tk.Entry(repl, width=48)
        stencil_entry.grid(row=0, column=1, sticky="w", padx=(6, 6))

        master_cb = ttk.Combobox(repl, state="readonly", width=44)
        master_cb.grid(row=1, column=1, sticky="w", padx=(6, 6), pady=(6, 0))

        def browse_stencil():
            p = filedialog.askopenfilename(filetypes=[("Visio Stencil", "*.vss *.vssx *.vssm *.vstx")])
            if not p:
                return
            stencil_entry.delete(0, "end")
            stencil_entry.insert(0, p)
            master_cb["values"] = []
            master_cb.set("")
            status.config(text="Loading masters from stencil...", fg="gray")

            def worker():
                try:
                    names = self.engine.list_masters_in_stencil(p)
                    self.root.after(0, lambda: _done(names))
                except Exception as e:
                    msg = str(e)
                    self.root.after(0, lambda msg=msg: _err(msg))

            def _done(names):
                master_cb["values"] = names
                if names:
                    master_cb.current(0)
                status.config(text=f"Loaded {len(names)} masters from stencil.", fg="gray")

            def _err(msg):
                status.config(text=f"Stencil load failed: {msg}", fg="red")

            threading.Thread(target=worker, daemon=True).start()

        tk.Button(repl, text="Browse...", command=browse_stencil).grid(row=0, column=2, sticky="w")

        tk.Label(repl, text="Master:").grid(row=1, column=0, sticky="w", pady=(6, 0))

        def do_replace():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Select shape", "Select a shape to replace.")
                return
            item = tree.item(sel[0])
            page = item["values"][0]
            sid = int(item["values"][1])
            stencil = (stencil_entry.get() or "").strip()
            master = (master_cb.get() or "").strip()
            if not stencil or not master:
                messagebox.showwarning("Missing", "Choose a stencil file and a master.")
                return

            status.config(text="Replacing shape...", fg="gray")

            def worker():
                try:
                    self.engine.replace_shape_in_active_document(page, sid, stencil, master)
                    self.root.after(0, lambda: _done())
                except Exception as e:
                    msg = str(e)
                    self.root.after(0, lambda msg=msg: _err(msg))

            def _done():
                status.config(text="Replacement complete. Refresh to see updates.", fg="gray")
                self._refresh_shapes(status, tree, details)

            def _err(msg):
                status.config(text=f"Replace failed: {msg}", fg="red")

            threading.Thread(target=worker, daemon=True).start()

        tk.Button(repl, text="Replace", command=do_replace).grid(row=2, column=1, sticky="w", pady=(10, 0))

        # Kick off initial refresh
        self._refresh_shapes(status, tree, details)

    def _bring_visio_front(self):
        threading.Thread(target=self.engine.bring_visio_to_front, daemon=True).start()

    def _refresh_shapes(self, status_lbl: tk.Label, tree: ttk.Treeview, details):
        status_lbl.config(text="Refreshing shapes from active Visio document...", fg="gray")

        # clear
        for i in tree.get_children():
            tree.delete(i)

        def worker():
            try:
                shapes = self.engine.list_shapes_with_shape_data_in_active_document()
                self.root.after(0, lambda: _done(shapes))
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda msg=msg: _err(msg))

        def _done(shapes):
            for s in shapes:
                depth = int(s.get("depth", 0) or 0)
                indent = ("    " * depth) if depth > 0 else ""
                tree.insert(
                    "",
                    "end",
                    values=(
                        s.get("page", ""),
                        s.get("shape_id", 0),
                        f"{indent}{s.get('shape_name', '')}",
                        s.get("master", ""),
                    ),
                )
            status_lbl.config(text=f"Loaded {len(shapes)} shapes.", fg="gray")
            try:
                details.delete(0, "end")
                details.insert("end", "Select a shape to view its Shape Data (Label=Value only).")
            except Exception:
                pass

        def _err(msg):
            status_lbl.config(
                text=(
                    "Could not read active Visio document. "
                    "Make sure Visio is open and a document is active. "
                    f"({msg})"
                ),
                fg="red",
            )

        threading.Thread(target=worker, daemon=True).start()

    def _load_shape_details(self, status_lbl: tk.Label, tree: ttk.Treeview, details):
        sel = tree.selection()
        if not sel:
            return
        item = tree.item(sel[0])
        page = item["values"][0]
        sid = int(item["values"][1])

        status_lbl.config(text="Loading shape data...", fg="gray")

        def worker():
            try:
                lines = self.engine.get_shape_data_from_active_document(page, sid)
                self.root.after(0, lambda: _done(lines))
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda msg=msg: _err(msg))

        def _done(lines):
            try:
                details.delete(0, "end")
                if lines:
                    for ln in lines:
                        details.insert("end", ln)
                else:
                    details.insert("end", "(no Label=Value Shape Data found)")
            except Exception:
                pass
            status_lbl.config(text="Shape data loaded.", fg="gray")

        def _err(msg):
            status_lbl.config(text=f"Failed to load shape data: {msg}", fg="red")

        threading.Thread(target=worker, daemon=True).start()



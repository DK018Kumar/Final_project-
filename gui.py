import sys
import threading
import traceback
import tkinter as tk
import time
from tkinter import filedialog, messagebox, ttk

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
        self._replace_target = None  # {"page_id":..., "shape_id":..., "row":..., "col":..., "kind":...}
        self._central_visio = None  # {"page_id":..., "shape_id":...} set after generation

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

        # Central Shape Data (from ShapeSheet Prop rows of the master)
        central_sd = tk.LabelFrame(self.root, text="Central Shape Data (ShapeSheet)", padx=10, pady=6)
        central_sd.pack(fill="x", padx=10, pady=(0, 6))
        self.central_sd_text = tk.Text(central_sd, height=6, wrap="none", state="disabled")
        self.central_sd_text.pack(side="left", fill="both", expand=True)
        central_sd_scroll = tk.Scrollbar(central_sd, orient="vertical", command=self.central_sd_text.yview)
        central_sd_scroll.pack(side="right", fill="y")
        self.central_sd_text.configure(yscrollcommand=central_sd_scroll.set)

        # Global preview (keeps original preview behavior)
        preview_frame = tk.LabelFrame(self.root, text="Shape Data Preview (ShapeSheet)", padx=10, pady=6)
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
            tb = traceback.format_exc()
            # Ensure the real COM/Visio error is visible
            try:
                print(tb, file=sys.stderr)
            except Exception:
                pass
            self.root.after(0, lambda: messagebox.showerror("Error", tb))

    def _after_load_stencil(self, masters):
        # masters: list of dicts {'name':..., 'shape_data': [...]}
        self.loaded_masters = masters
        # master_name -> list of {"label":..., "value":...}
        self.masters_by_name = {m["name"]: m.get("shape_data", []) for m in masters}
        names = list(self.masters_by_name.keys())
        self.lbl_stencil.config(text="Stencil Loaded", fg="black")
        self.cb_central["values"] = names
        if names:
            self.cb_central.current(0)
            self._render_central_shape_data(names[0])
            self._render_preview_shape_data(names[0])
        messagebox.showinfo("Success", f"Loaded {len(names)} masters.")
        # Matrix cell Shape Data is shown per-cell, based on selected master and/or generated shape.

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

        cell_state["selected_master"] = name
        cell_state["selected_props"] = []

        # Show ShapeSheet Shape Data for the *master* (drop-to-temp is done in engine load)
        self._render_cell_shape_data(
            cell_state,
            row_idx=None,
            col_idx=None,
            shape_data_rows=self.masters_by_name.get(name, []),
            show_replace_button=False,
        )
        self._render_preview_shape_data(name)

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
            self._render_central_shape_data(name)
            self._render_preview_shape_data(name)

    def _render_central_shape_data(self, master_name):
        rows = self.masters_by_name.get(master_name, [])
        lines = [f"Master: {master_name}", "-" * 40]
        for r in rows:
            lbl = (r.get("label") or "").strip()
            val = (r.get("value") or "").strip()
            if lbl:
                lines.append(f"{lbl}: {val}")
        if len(lines) == 2:
            lines.append("(no Shape Data)")
        text_value = "\n".join(lines)
        self.central_sd_text.configure(state="normal")
        self.central_sd_text.delete("1.0", tk.END)
        self.central_sd_text.insert(tk.END, text_value)
        self.central_sd_text.configure(state="disabled")

    def _render_preview_shape_data(self, master_name):
        # Global preview panel (same ShapeSheet label/value list)
        rows = self.masters_by_name.get(master_name, [])
        lines = [f"Master: {master_name}", "-" * 40]
        for r in rows:
            lbl = (r.get("label") or "").strip()
            val = (r.get("value") or "").strip()
            if lbl:
                lines.append(f"{lbl}: {val}")
        if len(lines) == 2:
            lines.append("(no Shape Data)")

        text_value = "\n".join(lines)
        try:
            self.shape_data_text.configure(state="normal")
            self.shape_data_text.delete("1.0", tk.END)
            self.shape_data_text.insert(tk.END, text_value)
            self.shape_data_text.configure(state="disabled")
        except Exception:
            pass

    def _set_replace_target(self, cell_state, row_idx=None, col_idx=None):
        page_id = cell_state.get("visio_page_id")
        shape_id = cell_state.get("visio_shape_id")
        if not page_id or not shape_id:
            return
        self._replace_target = {
            "page_id": page_id,
            "shape_id": shape_id,
            "row": row_idx,
            "col": col_idx,
            "kind": "matrix_cell",
        }

    def _render_cell_shape_data(self, cell_state, row_idx, col_idx, shape_data_rows, show_replace_button):
        pf = cell_state["props_frame"]
        for w in pf.winfo_children():
            w.destroy()

        # Display Shape Data ONLY (Label/Value) and actions for dropped shapes
        title = tk.Label(pf, text="Shape Data", font=("Arial", 9, "bold"))
        title.pack(anchor="w")

        def _select_target(_event=None):
            self._set_replace_target(cell_state, row_idx=row_idx, col_idx=col_idx)

        if show_replace_button:
            title.bind("<Button-1>", _select_target)

            actions = tk.Frame(pf)
            actions.pack(anchor="w", pady=(4, 4))
            tk.Button(actions, text="Replace…", command=lambda: self._replace_via_dialog(cell_state)).pack(side="left")
            tk.Button(actions, text="Delete", command=lambda: self._delete_shape(cell_state)).pack(side="left", padx=(8, 0))

        if not shape_data_rows:
            lbl = tk.Label(pf, text="(no ShapeSheet Shape Data rows found)", fg="gray")
            lbl.pack(anchor="w")
            return

        header = tk.Frame(pf)
        header.pack(fill="x", anchor="w")
        tk.Label(header, text="Label", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        tk.Label(header, text="Value", font=("Arial", 9, "bold")).grid(row=0, column=1, sticky="w")

        # shape_data_rows can be list[{"label","value"}] or dict[row_name]->{...}
        if isinstance(shape_data_rows, dict):
            iterable = list(shape_data_rows.values())
        else:
            iterable = shape_data_rows

        for r in iterable:
            lbl = (r.get("label") or "").strip()
            val = (r.get("value") or "").strip()
            row = tk.Frame(pf)
            row.pack(fill="x", anchor="w", pady=1)
            l1 = tk.Label(row, text=lbl, anchor="w", width=18)
            l1.grid(row=0, column=0, sticky="w", padx=(0, 8))
            l2 = tk.Label(row, text=val, anchor="w", wraplength=280, justify="left")
            l2.grid(row=0, column=1, sticky="w")

            if show_replace_button:
                for w in (row, l1, l2):
                    w.bind("<Button-1>", _select_target)

    def _delete_shape(self, cell_state):
        page_id = cell_state.get("visio_page_id")
        shape_id = cell_state.get("visio_shape_id")
        if not page_id or not shape_id:
            messagebox.showwarning("Not generated", "Generate the Visio layout first, then you can delete shapes.")
            return

        if not messagebox.askyesno("Confirm delete", "Delete this shape from Visio?"):
            return

        def _run():
            try:
                self.engine.delete_shape_by_id(page_id, shape_id)

                def _after():
                    cell_state["visio_page_id"] = None
                    cell_state["visio_shape_id"] = None
                    pf = cell_state["props_frame"]
                    for w in pf.winfo_children():
                        w.destroy()
                    tk.Label(pf, text="(deleted)", fg="gray").pack(anchor="w")
                    messagebox.showinfo("Deleted", "Shape deleted.")

                self.root.after(0, _after)
            except Exception:
                tb = traceback.format_exc()
                self.root.after(0, lambda: messagebox.showerror("Error", tb))

        threading.Thread(target=_run, daemon=True).start()

    def _thread_generate(self, data):
        try:
            meta = self.engine.generate_layout(data)
            def _after():
                messagebox.showinfo("Success", "Visio layout generated!")
                try:
                    self._apply_layout_meta(meta)
                except Exception:
                    pass

            self.root.after(0, _after)
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))

    def _apply_layout_meta(self, meta):
        matrix = (meta or {}).get("matrix", [])
        central = (meta or {}).get("central")
        if isinstance(central, dict):
            self._central_visio = {"page_id": central.get("page_id"), "shape_id": central.get("shape_id")}
        for r, row in enumerate(self.matrix_widgets):
            for c, cell_state in enumerate(row):
                info = None
                try:
                    info = matrix[r][c]
                except Exception:
                    info = None

                if info and isinstance(info, dict):
                    cell_state["visio_shape_id"] = info.get("shape_id")
                    cell_state["visio_page_id"] = info.get("page_id")
                    # show ShapeSheet data for the actual dropped shape and enable Replace button
                    self._render_cell_shape_data(cell_state, r, c, info.get("shape_data", []), show_replace_button=True)
                else:
                    cell_state["visio_shape_id"] = None
                    cell_state["visio_page_id"] = None
                    pf = cell_state["props_frame"]
                    for w in pf.winfo_children():
                        w.destroy()
                    tk.Label(pf, text="(empty)", fg="gray").pack(anchor="w")

    def _replace_via_dialog(self, cell_state):
        # Must have a generated Visio shape to replace
        page_id = cell_state.get("visio_page_id")
        shape_id = cell_state.get("visio_shape_id")
        if not page_id or not shape_id:
            messagebox.showwarning("Not generated", "Generate the Visio layout first, then you can replace shapes.")
            return

        stencil_path = filedialog.askopenfilename(filetypes=[("Visio Stencil", "*.vssx *.vss *.vstx")])
        if not stencil_path:
            return

        def _run():
            try:
                masters = self.engine.get_stencil_master_names(stencil_path)
                if not masters:
                    raise Exception("No masters found in selected stencil.")

                # If only one master, use it automatically. Otherwise ask user.
                if len(masters) == 1:
                    chosen = masters[0]
                else:
                    chosen = None

                    def _ask():
                        dlg = tk.Toplevel(self.root)
                        dlg.title("Choose replacement master")
                        tk.Label(dlg, text="Select master to replace with:").pack(anchor="w", padx=10, pady=(10, 4))
                        cb = ttk.Combobox(dlg, state="readonly", values=masters, width=60)
                        cb.pack(padx=10, pady=(0, 10))
                        cb.current(0)

                        result = {"val": None}

                        def _ok():
                            result["val"] = cb.get()
                            dlg.destroy()

                        tk.Button(dlg, text="OK", command=_ok).pack(pady=(0, 10))
                        dlg.transient(self.root)
                        dlg.grab_set()
                        self.root.wait_window(dlg)
                        return result["val"]

                    chosen = self.root.after(0, lambda: None)  # placeholder
                    # marshal dialog to main thread
                    chosen_holder = {"val": None}

                    def _run_dialog():
                        chosen_holder["val"] = _ask()

                    self.root.after(0, _run_dialog)
                    # wait a bit for dialog completion (best-effort)
                    # We can't block Tk thread here; just poll.
                    for _ in range(600):
                        if chosen_holder["val"] is not None:
                            break
                        time.sleep(0.05)
                    chosen = chosen_holder["val"]

                if not chosen:
                    return

                label = (cell_state.get("label_entry").get() or "").strip() if cell_state.get("label_entry") else ""
                result = self.engine.replace_shape_by_id(
                    page_id=page_id,
                    shape_id=shape_id,
                    stencil_path=stencil_path,
                    replacement_master_name=chosen,
                    label_text=label or None,
                )
                def _after():
                    if isinstance(result, dict):
                        cell_state["visio_page_id"] = result.get("page_id")
                        cell_state["visio_shape_id"] = result.get("shape_id")
                        self._render_cell_shape_data(
                            cell_state,
                            row_idx=None,
                            col_idx=None,
                            shape_data_rows=result.get("shape_data", {}),
                            show_replace_button=True,
                        )
                    messagebox.showinfo("Success", "Shape replaced.")

                self.root.after(0, _after)
            except Exception as e:
                tb = traceback.format_exc()
                try:
                    print(tb, file=sys.stderr)
                except Exception:
                    pass
                self.root.after(0, lambda: messagebox.showerror("Error", tb))

        threading.Thread(target=_run, daemon=True).start()

    def on_close(self):
        try:
            self.engine.stop_selection_watch()
        except Exception:
            pass
        self.root.destroy()
        sys.exit()

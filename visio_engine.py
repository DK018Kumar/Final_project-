import os
import shutil
import tempfile
import threading
import time

import pythoncom
import win32com.client
from win32com.client import constants


class VisioEngine:
    def __init__(self):
        self.current_stencil_path = None
        self._selection_thread = None
        self._selection_stop = None

    # Always return a stable Visio instance (visible = True)
    def _get_visio(self):
        try:
            app = win32com.client.GetActiveObject("Visio.Application")
        except Exception:
            app = win32com.client.Dispatch("Visio.Application")
        app.Visible = True
        docs = app.Documents
        return app, docs

    def _normalize(self, name):
        return " ".join(name.strip().split()) if name else ""

    # ----------------------------------------------------------
    # Load stencil + extract master names + shape data (best-effort)
    # Returns list of dicts: {'name': <str>, 'props': [{'type':..., 'value':...}, ...]}
    # ----------------------------------------------------------
    def load_stencil_masters(self, stencil_path):
        pythoncom.CoInitialize()
        try:
            if not os.path.exists(stencil_path):
                raise Exception("Stencil not found.")

            base = os.path.basename(stencil_path)
            temp_copy = os.path.join(tempfile.gettempdir(), f"stencil_{int(time.time() * 1000)}_{base}")
            shutil.copy2(stencil_path, temp_copy)
            self.current_stencil_path = temp_copy

            app, docs = self._get_visio()
            stencil = docs.OpenEx(temp_copy, 64)

            # Create a temporary document/page to drop masters.
            # This is the most reliable way to read ALL Shape Data from grouped/nested shapes.
            tmp_doc = docs.Add("")
            tmp_page = tmp_doc.Pages(1)

            masters_info = []
            for m in stencil.Masters:
                try:
                    name = self._normalize(m.Name)
                except Exception:
                    name = self._normalize(getattr(m, "Name", "Unknown"))

                props = []
                # Primary: drop and scrape the resulting shape tree
                try:
                    dropped = tmp_page.Drop(m, 0, 0)
                    props = self._collect_shape_properties_from_instance(dropped)
                    # Clean up the temporary drop to keep the page light
                    try:
                        dropped.Delete()
                    except Exception:
                        pass
                except Exception:
                    props = []

                # Fallbacks
                if not props:
                    props = self._collect_master_properties(m)
                if not props:
                    props = self._legacy_prop_scrape(m)

                masters_info.append({"name": name, "props": props})

            try:
                tmp_doc.Close(False)
            except Exception:
                pass

            stencil.Close()

            # Deduplicate by name (preserve first)
            seen = set()
            final = []
            for info in masters_info:
                n = info["name"]
                if n not in seen:
                    seen.add(n)
                    final.append(info)

            return final
        finally:
            pythoncom.CoUninitialize()

    # ----------------------------------------------------------
    # Live selection watcher: when user clicks shapes in Visio,
    # call the provided callback with shape + Shape Data.
    #
    # callback signature:
    #   callback({"title": <str>, "props": <list[dict]>})
    # ----------------------------------------------------------
    def start_selection_watch(self, callback):
        if self._selection_thread and self._selection_thread.is_alive():
            return

        self._selection_stop = threading.Event()

        def _run():
            pythoncom.CoInitialize()
            try:
                app, _docs = self._get_visio()
                engine = self

                class _AppEvents:
                    def __init__(self):
                        self._last_sig = None

                    def SelectionChanged(self, window):
                        try:
                            sel = getattr(window, "Selection", None)
                            if sel is None:
                                return
                            try:
                                cnt = int(sel.Count)
                            except Exception:
                                cnt = 0
                            if cnt <= 0:
                                return

                            try:
                                shp = sel.Item(1)
                            except Exception:
                                return

                            title = engine._format_shape_title(shp)
                            props = engine._collect_shape_properties_from_instance(shp)
                            shape_id = None
                            page_id = None
                            try:
                                shape_id = int(getattr(shp, "ID", 0))
                            except Exception:
                                shape_id = None
                            try:
                                pg = getattr(shp, "ContainingPage", None)
                                page_id = int(getattr(pg, "ID", 0)) if pg is not None else None
                            except Exception:
                                page_id = None

                            # Avoid spamming callback with identical payloads
                            sig = (title, len(props), shape_id, page_id)
                            if sig == self._last_sig:
                                return
                            self._last_sig = sig

                            try:
                                callback({"title": title, "props": props, "shape_id": shape_id, "page_id": page_id})
                            except Exception:
                                return
                        except Exception:
                            return

                handler = win32com.client.WithEvents(app, _AppEvents)

                while not self._selection_stop.is_set():
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                    time.sleep(0.1)

                _ = handler  # keep alive until exit
            finally:
                pythoncom.CoUninitialize()

        self._selection_thread = threading.Thread(target=_run, daemon=True)
        self._selection_thread.start()

    def stop_selection_watch(self):
        if self._selection_stop is not None:
            try:
                self._selection_stop.set()
            except Exception:
                pass
        self._selection_stop = None
        self._selection_thread = None

    # ----------------------------------------------------------
    # List master names from an arbitrary stencil file (does not
    # change self.current_stencil_path).
    # ----------------------------------------------------------
    def get_stencil_master_names(self, stencil_path):
        pythoncom.CoInitialize()
        try:
            if not stencil_path or not os.path.exists(stencil_path):
                raise Exception("Stencil not found.")

            app, docs = self._get_visio()

            base = os.path.basename(stencil_path)
            temp_copy = os.path.join(tempfile.gettempdir(), f"pick_{int(time.time() * 1000)}_{base}")
            shutil.copy2(stencil_path, temp_copy)

            stencil = docs.OpenEx(temp_copy, 64)
            try:
                names = []
                for m in stencil.Masters:
                    try:
                        nm = self._normalize(getattr(m, "Name", "")) or self._normalize(getattr(m, "NameU", ""))
                    except Exception:
                        nm = ""
                    if nm:
                        names.append(nm)

                # Dedup preserve order
                seen = set()
                out = []
                for n in names:
                    if n in seen:
                        continue
                    seen.add(n)
                    out.append(n)
                return out
            finally:
                try:
                    stencil.Close()
                except Exception:
                    pass
        finally:
            pythoncom.CoUninitialize()

    # ----------------------------------------------------------
    # Replace the currently selected Visio shape with another master
    # from the currently loaded stencil, at the same location.
    #
    # - replacement_master_name: master name to drop
    # - label_text: optional label to draw above the replacement
    # ----------------------------------------------------------
    def replace_selected_shape(self, replacement_master_name, label_text=None):
        pythoncom.CoInitialize()
        try:
            if not self.current_stencil_path:
                raise Exception("Stencil not loaded.")

            app, docs = self._get_visio()

            # Selection
            window = getattr(app, "ActiveWindow", None)
            if window is None:
                raise Exception("No active Visio window.")

            selection = getattr(window, "Selection", None)
            if selection is None:
                raise Exception("No selection found in Visio.")

            try:
                sel_count = int(selection.Count)
            except Exception:
                sel_count = 0
            if sel_count <= 0:
                raise Exception("Select a shape in Visio first.")

            try:
                old_shape = selection.Item(1)
            except Exception:
                raise Exception("Could not read selected shape.")

            # Read old shape placement (in internal units)
            try:
                page = old_shape.ContainingPage
            except Exception:
                page = None
            if page is None:
                raise Exception("Selected shape has no containing page.")

            def _cell_result_iu(shp, cell_name, default=None):
                try:
                    return float(shp.CellsU(cell_name).ResultIU)
                except Exception:
                    return default

            x = _cell_result_iu(old_shape, "PinX", 0.0)
            y = _cell_result_iu(old_shape, "PinY", 0.0)
            width = _cell_result_iu(old_shape, "Width", None)
            height = _cell_result_iu(old_shape, "Height", None)
            angle = _cell_result_iu(old_shape, "Angle", None)

            # Open stencil (copy for safety)
            base = os.path.basename(self.current_stencil_path)
            stencil_copy = os.path.join(tempfile.gettempdir(), f"replace_{int(time.time() * 1000)}_{base}")
            shutil.copy2(self.current_stencil_path, stencil_copy)
            stencil = docs.OpenEx(stencil_copy, 64)

            try:
                new_master = None
                try:
                    new_master = stencil.Masters.Item(replacement_master_name)
                except Exception:
                    for m in stencil.Masters:
                        try:
                            if self._normalize(m.Name).lower() == self._normalize(replacement_master_name).lower():
                                new_master = m
                                break
                        except Exception:
                            continue

                if new_master is None:
                    raise Exception(f"Replacement master '{replacement_master_name}' not found.")

                # Drop replacement at same location
                new_shape = page.Drop(new_master, x, y)

                # Preserve basic transforms (best-effort)
                if width is not None:
                    try:
                        new_shape.CellsU("Width").ResultIU = float(width)
                    except Exception:
                        pass
                if height is not None:
                    try:
                        new_shape.CellsU("Height").ResultIU = float(height)
                    except Exception:
                        pass
                if angle is not None:
                    try:
                        new_shape.CellsU("Angle").ResultIU = float(angle)
                    except Exception:
                        pass

                # Delete old shape
                try:
                    old_shape.Delete()
                except Exception:
                    pass

                # Add/refresh label textbox above the replacement
                try:
                    bbox = self._get_shape_bbox(new_shape)
                    if label_text is None:
                        label_text = replacement_master_name
                    label_text = (label_text or "").strip()
                    if bbox and label_text:
                        self._place_label_top_left(page, bbox, label_text)
                except Exception:
                    pass
            finally:
                try:
                    stencil.Close()
                except Exception:
                    pass
        finally:
            pythoncom.CoUninitialize()

    # ----------------------------------------------------------
    # Replace a specific shape (by IDs) with a master from a given
    # stencil file, preserving position/size/rotation.
    # ----------------------------------------------------------
    def replace_shape_by_id(self, page_id, shape_id, stencil_path, replacement_master_name, label_text=None):
        pythoncom.CoInitialize()
        try:
            if not stencil_path:
                raise Exception("Stencil path is required.")
            if not os.path.exists(stencil_path):
                raise Exception("Stencil not found.")
            if not replacement_master_name:
                raise Exception("Replacement master name is required.")
            if not page_id or not shape_id:
                raise Exception("No Visio shape selected.")

            app, docs = self._get_visio()

            doc = getattr(app, "ActiveDocument", None)
            if doc is None:
                raise Exception("No active Visio document.")

            page = None
            try:
                page = doc.Pages.ItemFromID(int(page_id))
            except Exception:
                # fallback search
                try:
                    for p in doc.Pages:
                        try:
                            if int(getattr(p, "ID", 0)) == int(page_id):
                                page = p
                                break
                        except Exception:
                            continue
                except Exception:
                    page = None
            if page is None:
                raise Exception("Could not find the selected shape's page.")

            old_shape = None
            try:
                old_shape = page.Shapes.ItemFromID(int(shape_id))
            except Exception:
                old_shape = None
            if old_shape is None:
                raise Exception("Could not find the selected shape (maybe it was deleted).")

            def _cell_result_iu(shp, cell_name, default=None):
                try:
                    return float(shp.CellsU(cell_name).ResultIU)
                except Exception:
                    return default

            x = _cell_result_iu(old_shape, "PinX", 0.0)
            y = _cell_result_iu(old_shape, "PinY", 0.0)
            width = _cell_result_iu(old_shape, "Width", None)
            height = _cell_result_iu(old_shape, "Height", None)
            angle = _cell_result_iu(old_shape, "Angle", None)

            # Open stencil (copy for safety)
            base = os.path.basename(stencil_path)
            stencil_copy = os.path.join(tempfile.gettempdir(), f"replacefile_{int(time.time() * 1000)}_{base}")
            shutil.copy2(stencil_path, stencil_copy)
            stencil = docs.OpenEx(stencil_copy, 64)
            try:
                new_master = None
                try:
                    new_master = stencil.Masters.Item(replacement_master_name)
                except Exception:
                    for m in stencil.Masters:
                        try:
                            if self._normalize(m.Name).lower() == self._normalize(replacement_master_name).lower():
                                new_master = m
                                break
                        except Exception:
                            continue

                if new_master is None:
                    raise Exception(f"Replacement master '{replacement_master_name}' not found in selected stencil.")

                new_shape = page.Drop(new_master, x, y)

                # Preserve basic transforms (best-effort)
                if width is not None:
                    try:
                        new_shape.CellsU("Width").ResultIU = float(width)
                    except Exception:
                        pass
                if height is not None:
                    try:
                        new_shape.CellsU("Height").ResultIU = float(height)
                    except Exception:
                        pass
                if angle is not None:
                    try:
                        new_shape.CellsU("Angle").ResultIU = float(angle)
                    except Exception:
                        pass

                try:
                    old_shape.Delete()
                except Exception:
                    pass

                # Add label textbox above the replacement (simple & reliable)
                try:
                    bbox = self._get_shape_bbox(new_shape)
                    txt = (label_text or replacement_master_name or "").strip()
                    if bbox and txt:
                        self._place_label_top_left(page, bbox, txt)
                except Exception:
                    pass
            finally:
                try:
                    stencil.Close()
                except Exception:
                    pass
        finally:
            pythoncom.CoUninitialize()

    # ----------------------------------------------------------
    # Generate the real Visio layout (NO PREVIEW)
    # Adds a small bold label placed top-left of each dropped substation (best-effort)
    # ----------------------------------------------------------
    def generate_layout(self, data):
        pythoncom.CoInitialize()
        try:
            if not self.current_stencil_path:
                raise Exception("Stencil not loaded.")

            app, docs = self._get_visio()

            # Per-operation stencil copy
            base = os.path.basename(self.current_stencil_path)
            stencil_copy = os.path.join(tempfile.gettempdir(), f"layout_{int(time.time() * 1000)}_{base}")
            shutil.copy2(self.current_stencil_path, stencil_copy)
            stencil = docs.OpenEx(stencil_copy, 64)

            # Create new drawing
            doc = docs.Add("")
            page = doc.Pages(1)

            # CONFIGURABLE SPACING
            CENTRAL_X = 10.0
            CENTRAL_Y = 15.0
            H_SPACING = 18.0
            V_SPACING = 12.0
            FIRST_ROW_OFFSET = 15.0

            custom_label = data.get("central_label")
            if custom_label is None:
                custom_label = data.get("substation_label", "")
            if isinstance(custom_label, str):
                custom_label = custom_label.strip()
            elif custom_label is not None:
                custom_label = str(custom_label).strip()
            else:
                custom_label = ""

            # Drop central
            central_name = data.get("central")
            try:
                central_master = stencil.Masters.Item(central_name)
            except Exception:
                central_master = None
                for m in stencil.Masters:
                    try:
                        if self._normalize(m.Name).lower() == self._normalize(central_name).lower():
                            central_master = m
                            break
                    except Exception:
                        continue

            if central_master is None:
                raise Exception(f"Central master '{central_name}' not found.")

            label_text = custom_label if custom_label else central_name
            self._drop_shape_with_label(page, central_master, CENTRAL_X, CENTRAL_Y, label_text)

            # Matrix
            matrix = data.get("matrix_subs", [])
            rows = len(matrix)
            cols = max((len(r) for r in matrix), default=0)

            if cols > 0:
                total_width = (cols - 1) * H_SPACING
                start_x = CENTRAL_X - total_width / 2
            else:
                start_x = CENTRAL_X

            for r in range(rows):
                y = CENTRAL_Y - FIRST_ROW_OFFSET - r * V_SPACING
                for c in range(cols):
                    cell = matrix[r][c] if c < len(matrix[r]) else None

                    cell_name = ""
                    cell_label = ""
                    if isinstance(cell, dict):
                        name_candidate = cell.get("name") or cell.get("master") or cell.get("value")
                        if isinstance(name_candidate, str):
                            cell_name = name_candidate.strip()
                        elif name_candidate is not None:
                            cell_name = str(name_candidate).strip()

                        lbl_candidate = cell.get("label", "")
                        if isinstance(lbl_candidate, str):
                            cell_label = lbl_candidate.strip()
                        elif lbl_candidate is not None:
                            cell_label = str(lbl_candidate).strip()
                    else:
                        if isinstance(cell, str):
                            cell_name = cell.strip()
                        elif cell is not None:
                            cell_name = str(cell).strip()

                    if not cell_name or cell_name.lower() == "none":
                        continue

                    # find master
                    try:
                        m_obj = stencil.Masters.Item(cell_name)
                    except Exception:
                        m_obj = None
                        for item in stencil.Masters:
                            try:
                                if self._normalize(item.Name).lower() == self._normalize(cell_name).lower():
                                    m_obj = item
                                    break
                            except Exception:
                                continue

                    if m_obj is None:
                        continue

                    x = start_x + c * H_SPACING
                    label = cell_label if cell_label else cell_name
                    self._drop_shape_with_label(page, m_obj, x, y, label)

            # Fit view
            try:
                app.ActiveWindow.Zoom = constants.visZoomFitPage
            except Exception:
                pass

            stencil.Close()
        finally:
            pythoncom.CoUninitialize()

    def _drop_shape_with_label(self, page, master, x, y, label_text):
        try:
            shp = page.Drop(master, x, y)
        except Exception:
            return None

        bbox_before = self._get_shape_bbox(shp)

        # Ungroup after drop (so child shapes become accessible), but keep label based on final bbox if possible.
        self._try_ungroup(shp)

        bbox = self._get_shape_bbox(shp) or bbox_before

        if label_text and bbox:
            try:
                self._place_label_top_left(page, bbox, label_text)
            except Exception:
                pass

        return shp

    def _get_shape_bbox(self, shp):
        if shp is None:
            return None
        try:
            return shp.BoundingBox(constants.visBBoxUpright)
        except Exception:
            return None

    # ------------------------
    # Shape Data extraction
    # ------------------------
    def _format_shape_title(self, shp):
        if shp is None:
            return "Selected: (none)"

        master_name = ""
        try:
            m = getattr(shp, "Master", None)
            if m is not None:
                master_name = self._normalize(getattr(m, "Name", "")) or self._normalize(getattr(m, "NameU", ""))
        except Exception:
            master_name = ""

        shape_id = ""
        try:
            shape_id = self._normalize(getattr(shp, "NameU", "")) or self._normalize(getattr(shp, "Name", ""))
        except Exception:
            shape_id = ""

        if master_name and shape_id:
            return f"Selected: {master_name} ({shape_id})"
        if master_name:
            return f"Selected: {master_name}"
        if shape_id:
            return f"Selected: {shape_id}"
        return "Selected: (unknown shape)"

    def _collect_shape_properties_from_instance(self, root_shape):
        props = []
        try:
            props.extend(self._collect_props_from_shape(root_shape, parent_name=""))
        except Exception:
            pass

        # Deduplicate (type, value) while preserving order
        seen = set()
        final = []
        for p in props:
            key = (p.get("type", ""), p.get("value", ""))
            if key in seen:
                continue
            seen.add(key)
            final.append(p)
        return final

    def _collect_master_properties(self, master):
        props = []
        try:
            if hasattr(master, "Shapes"):
                for shp in master.Shapes:
                    props.extend(self._collect_props_from_shape(shp))
        except Exception:
            pass
        return props

    def _collect_props_from_shape(self, shape, parent_name=""):
        results = []
        shape_name = self._shape_display_name(shape, parent_name)

        # Collect both "Shape Data" (Prop) and common User-defined fields
        results.extend(self._extract_prop_rows(shape, shape_name))
        results.extend(self._extract_user_rows(shape, shape_name))

        try:
            if hasattr(shape, "Shapes"):
                for child in shape.Shapes:
                    results.extend(self._collect_props_from_shape(child, shape_name))
        except Exception:
            pass
        return results

    def _shape_display_name(self, shape, fallback):
        # Prefer the shape's master name (more stable than Sheet.12)
        try:
            m = getattr(shape, "Master", None)
            if m is not None:
                nm = self._normalize(getattr(m, "Name", "")) or self._normalize(getattr(m, "NameU", ""))
                if nm:
                    return nm
        except Exception:
            pass

        try:
            nm = self._normalize(getattr(shape, "NameU", "")) or self._normalize(getattr(shape, "Name", ""))
            if nm:
                return nm
        except Exception:
            pass

        return fallback or "Shape"

    def _extract_prop_rows(self, shape, shape_name):
        props = []
        if shape is None:
            return props

        try:
            if shape.SectionExists(constants.visSectionProp, 0):
                row_count = shape.RowCount(constants.visSectionProp)
                for idx in range(row_count):
                    try:
                        row_name = shape.RowNameU(constants.visSectionProp, idx)
                    except Exception:
                        row_name = f"Row_{idx}"

                    label = ""
                    value = ""

                    # Label
                    try:
                        label = shape.CellsU(f"Prop.{row_name}.Label").ResultStr(0)
                    except Exception:
                        try:
                            label = shape.CellsU(f"Prop.{row_name}.Prompt").ResultStr(0)
                        except Exception:
                            label = row_name

                    label = label or row_name

                    # Value
                    try:
                        value = shape.CellsU(f"Prop.{row_name}.Value").ResultStr(0)
                    except Exception:
                        # Some shapes store a direct cell or formula
                        try:
                            value = shape.CellsU(f"Prop.{row_name}").ResultStr(0)
                        except Exception:
                            value = ""

                    combined_label = f"{shape_name or 'Shape'} - {label}".strip(" -")
                    props.append({"type": combined_label, "value": value})
        except Exception:
            pass

        return props

    def _extract_user_rows(self, shape, shape_name):
        props = []
        if shape is None:
            return props

        try:
            if shape.SectionExists(constants.visSectionUser, 0):
                row_count = shape.RowCount(constants.visSectionUser)
                for idx in range(row_count):
                    try:
                        row_name = shape.RowNameU(constants.visSectionUser, idx)
                    except Exception:
                        row_name = f"Row_{idx}"

                    value = ""
                    try:
                        value = shape.CellsU(f"User.{row_name}").ResultStr(0)
                    except Exception:
                        try:
                            value = shape.CellsU(f"User.{row_name}").FormulaU
                        except Exception:
                            value = ""

                    combined_label = f"{shape_name or 'Shape'} - User.{row_name}".strip(" -")
                    props.append({"type": combined_label, "value": value})
        except Exception:
            pass

        return props

    def _legacy_prop_scrape(self, master):
        props = []
        try:
            if hasattr(master, "Properties"):
                for p in master.Properties:
                    try:
                        p_name = getattr(p, "Name", None) or getattr(p, "Label", None) or str(p)
                        p_val = ""
                        try:
                            p_val = getattr(p, "ValueAsString", None)
                        except Exception:
                            p_val = None
                        if p_val is None:
                            try:
                                p_val = getattr(p, "Value", None)
                            except Exception:
                                p_val = None
                        p_val = "" if p_val is None else str(p_val)
                        props.append({"type": str(p_name), "value": p_val})
                    except Exception:
                        continue
        except Exception:
            pass

        if props:
            return props

        try:
            if hasattr(master, "Shapes"):
                for shp in master.Shapes:
                    for i in range(1, 51):
                        try:
                            lbl_cell_name = f"Prop.Row_{i}.Label"
                            val_cell_name = f"Prop.Row_{i}.Value"
                            if shp.CellExistsU(lbl_cell_name, 0):
                                try:
                                    lbl = shp.CellsU(lbl_cell_name).ResultStr(0)
                                except Exception:
                                    lbl = str(i)
                            elif shp.CellExistsU(f"Prop.Row_{i}", 0):
                                try:
                                    lbl = shp.CellsU(f"Prop.Row_{i}.Label").ResultStr(0)
                                except Exception:
                                    lbl = f"Row_{i}"
                            else:
                                continue

                            try:
                                val = shp.CellsU(val_cell_name).ResultStr(0)
                            except Exception:
                                try:
                                    val = shp.CellsU(f"Prop.Row_{i}.Value").ResultStr(0)
                                except Exception:
                                    val = ""
                            props.append({"type": lbl, "value": val})
                        except Exception:
                            continue
                    if props:
                        break
        except Exception:
            pass

        return props

    # ------------------------
    # Group/label handling
    # ------------------------
    def _try_ungroup(self, shp):
        try:
            if shp is None:
                return None

            if getattr(shp, "Type", None) != constants.visTypeGroup:
                return None

            app = getattr(shp, "Application", None)
            window = app.ActiveWindow if app else None

            prev_alert_enabled = None
            prev_alert_response = None
            resp_ok = getattr(constants, "visResponseOK", 1)

            if app:
                try:
                    prev_alert_enabled = getattr(app, "AlertEnabled", None)
                except Exception:
                    prev_alert_enabled = None
                try:
                    prev_alert_response = getattr(app, "AlertResponse", None)
                except Exception:
                    prev_alert_response = None

                try:
                    app.AlertResponse = resp_ok
                except Exception:
                    pass
                try:
                    app.AlertEnabled = False
                except Exception:
                    pass

            selection = None
            if window:
                try:
                    window.DeselectAll()
                    shp.Select(constants.visSelect)
                    selection = window.Selection
                except Exception:
                    selection = None

            try:
                if selection:
                    selection.Ungroup()
                else:
                    shp.Ungroup()
            except Exception:
                pass
            finally:
                if app:
                    if prev_alert_enabled is not None:
                        try:
                            app.AlertEnabled = prev_alert_enabled
                        except Exception:
                            pass
                    if prev_alert_response is not None:
                        try:
                            app.AlertResponse = prev_alert_response
                        except Exception:
                            pass
        except Exception:
            pass
        return None

    def _place_label_top_left(self, page, bbox, label_text):
        if not bbox:
            return

        left, bottom, right, top = bbox

        # Place a small label rectangle just above the top-left of the shape.
        # (Top-left alignment, bold text)
        height = 0.55
        gap = 0.20

        # Width based on text length (bounded)
        width = max(1.8, min(6.0, 0.16 * max(1, len(label_text))))

        rect_left = left
        rect_right = rect_left + width
        rect_bottom = top + gap
        rect_top = rect_bottom + height

        # Visio expects (x1, y1, x2, y2) as opposite corners;
        # use (left,bottom,right,top) ordering to avoid inverted boxes.
        t = page.DrawRectangle(rect_left, rect_bottom, rect_right, rect_top)
        try:
            t.Text = label_text
        except Exception:
            pass

        # Transparent box (text only)
        try:
            t.CellsU("LinePattern").FormulaU = "0"
        except Exception:
            pass
        try:
            t.CellsU("FillPattern").FormulaU = "0"
        except Exception:
            pass

        # Bold + top-left alignment
        try:
            t.CellsU("Char.Bold").FormulaU = "1"
        except Exception:
            pass
        try:
            # 10 pt is readable; feel free to change
            t.CellsU("Char.Size").FormulaU = "10 pt"
        except Exception:
            pass
        try:
            t.CellsU("Para.HorzAlign").FormulaU = "0"  # left
        except Exception:
            pass
        try:
            t.CellsU("TextBlock.VerticalAlign").FormulaU = "0"  # top
        except Exception:
            pass

        # Margins
        try:
            t.CellsU("TextBlock.MarginLeft").FormulaU = "0.03 in"
        except Exception:
            pass
        try:
            t.CellsU("TextBlock.MarginRight").FormulaU = "0.03 in"
        except Exception:
            pass
        try:
            t.CellsU("TextBlock.MarginTop").FormulaU = "0.02 in"
        except Exception:
            pass

        try:
            t.BringToFront()
        except Exception:
            pass

import os
import shutil
import tempfile
import pythoncom
import win32com.client
from win32com.client import constants
import time


class VisioEngine:
    def __init__(self):
        self.current_stencil_path = None

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
            temp_copy = os.path.join(
                tempfile.gettempdir(),
                f"stencil_{int(time.time() * 1000)}_{base}"
            )
            shutil.copy2(stencil_path, temp_copy)
            self.current_stencil_path = temp_copy

            app, docs = self._get_visio()
            stencil = docs.OpenEx(temp_copy, 64)

            masters_info = []
            for m in stencil.Masters:
                # Name normalization
                try:
                    name = self._normalize(m.Name)
                except Exception:
                    name = self._normalize(getattr(m, "Name", "Unknown"))

                props = []
                # 1) Try direct Properties collection (if present)
                try:
                    if hasattr(m, "Properties"):
                        for p in m.Properties:
                            try:
                                p_name = getattr(p, "Name", None) or getattr(p, "Label", None) or str(p)
                                # Prefer ValueAsString if available
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

                # 2) Fallback: Inspect first shape inside master and read Prop.Row_* cells
                if not props:
                    try:
                        # Some masters contain shapes collection
                        if getattr(m, "Shapes", None):
                            # iterate shapes in master
                            for shp in m.Shapes:
                                # Try up to 50 Prop rows heuristically
                                for i in range(1, 51):
                                    try:
                                        # Try the common cell names
                                        lbl_cell_name = f"Prop.Row_{i}.Label"
                                        val_cell_name = f"Prop.Row_{i}.Value"
                                        # If label exists, read it
                                        if shp.CellExistsU(lbl_cell_name, 0):
                                            try:
                                                lbl = shp.CellsU(lbl_cell_name).ResultStr(0)
                                            except Exception:
                                                lbl = str(i)
                                        elif shp.CellExistsU(f"Prop.Row_{i}", 0):
                                            # fallback: use row existance
                                            try:
                                                lbl = shp.CellsU(f"Prop.Row_{i}.Label").ResultStr(0)
                                            except Exception:
                                                lbl = f"Row_{i}"
                                        else:
                                            # Not present; continue
                                            continue

                                        # value read
                                        try:
                                            val = shp.CellsU(val_cell_name).ResultStr(0)
                                        except Exception:
                                            # try ValueAsString style or generic cell
                                            try:
                                                val = shp.CellsU(f"Prop.Row_{i}.Value").ResultStr(0)
                                            except Exception:
                                                val = ""
                                        props.append({"type": lbl, "value": val})
                                    except Exception:
                                        # row doesn't exist or access error, continue
                                        continue
                                # if we got props for this shape, stop checking other shapes
                                if props:
                                    break
                    except Exception:
                        pass

                masters_info.append({"name": name, "props": props})

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
            stencil_copy = os.path.join(
                tempfile.gettempdir(),
                f"layout_{int(time.time() * 1000)}_{base}"
            )
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

        bbox = self._get_shape_bbox(shp)
        if label_text and bbox:
            try:
                self._place_label_for_shape(page, None, label_text, bbox=bbox)
            except Exception:
                pass

        self._try_ungroup(shp)
        return shp

    def _get_shape_bbox(self, shp):
        if shp is None:
            return None
        try:
            return shp.BoundingBox(constants.visBBoxUpright)
        except Exception:
            return None

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

    # Helper to place small bold label at top-left of a given shape (best-effort)
    def _place_label_for_shape(self, page, shp, label_text, bbox=None):
        try:
            # get bounding box: returns (Left, Bottom, Right, Top)
            if bbox is None:
                bbox = self._get_shape_bbox(shp)
            if not bbox:
                return
            left = bbox[0]
            bottom = bbox[1]
            right = bbox[2]
            top = bbox[3]
            # small rectangle width/height for label (in inches)
            w = max(1.2, min(2.5, (right - left) * 0.6))
            h = 0.5

            # place rectangle with top-left anchored inside the shape (slightly inset)
            rect_left = left + 0.05
            rect_top = top - 0.05
            rect_right = rect_left + w
            rect_bottom = rect_top - h

            # create a small textbox rectangle; if it collides, it's OK — user can reposition in Visio
            t = page.DrawRectangle(rect_left, rect_top, rect_right, rect_bottom)
            # set text and style
            try:
                t.Text = label_text
            except Exception:
                pass
            try:
                t.CellsU("LinePattern").FormulaU = "0"
            except Exception:
                pass
            try:
                t.CellsU("FillPattern").FormulaU = "0"
            except Exception:
                pass
            # try to set bold for characters
            try:
                t.CellsU("Char.Bold").FormulaU = "1"
            except Exception:
                try:
                    chars = t.Characters
                    chars.set_CharProps(constants.visCharacterBold, 1)
                except Exception:
                    pass
            # try to put label above the shape (bring to front)
            try:
                t.BringToFront()
            except Exception:
                pass
        except Exception:
            # swallow label placement errors to avoid crashing layout generation
            pass

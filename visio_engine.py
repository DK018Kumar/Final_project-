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

            custom_label = data.get("substation_label", "")
            if isinstance(custom_label, str):
                custom_label = custom_label.strip()
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

            try:
                central_shape = page.Drop(central_master, CENTRAL_X, CENTRAL_Y)
            except Exception:
                central_shape = None

            # Add label for central (best-effort)
            try:
                if central_shape:
                    label_text = custom_label if custom_label else central_name
                    self._place_label_for_shape(page, central_shape, label_text)
            except Exception:
                pass

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
                    name = matrix[r][c] if c < len(matrix[r]) else "None"
                    if not name or name == "None":
                        continue

                    # find master
                    try:
                        m_obj = stencil.Masters.Item(name)
                    except Exception:
                        m_obj = None
                        for item in stencil.Masters:
                            try:
                                if self._normalize(item.Name).lower() == self._normalize(name).lower():
                                    m_obj = item
                                    break
                            except Exception:
                                continue

                    if m_obj is None:
                        continue

                    x = start_x + c * H_SPACING
                    try:
                        shp = page.Drop(m_obj, x, y)
                    except Exception:
                        shp = None

                    if shp is None:
                        continue

                    # place label in top-left of shp (best-effort)
                    try:
                        self._place_label_for_shape(page, shp, name)
                    except Exception:
                        pass

            # Fit view
            try:
                app.ActiveWindow.Zoom = constants.visZoomFitPage
            except Exception:
                pass

            stencil.Close()
        finally:
            pythoncom.CoUninitialize()

    # Helper to place small bold label at top-left of a given shape (best-effort)
    def _place_label_for_shape(self, page, shp, label_text):
        try:
            # get bounding box: returns (Left, Bottom, Right, Top)
            bbox = shp.BoundingBox(constants.visBBoxUpright)
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

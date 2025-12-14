import os
import shutil
import tempfile
import time
from typing import Any, Dict, List, Optional, Tuple


class VisioEngine:
    def __init__(self):
        self.current_stencil_path: Optional[str] = None

    # ------------------------
    # COM helpers (import lazily so module imports on non-Windows)
    # ------------------------
    def _com(self):
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
            from win32com.client import constants  # type: ignore

            return pythoncom, win32com.client, constants
        except Exception as e:
            raise RuntimeError(
                "Visio automation requires Windows + pywin32 (pythoncom/win32com). "
                f"Import failed: {e}"
            )

    def _co_init(self):
        pythoncom, _, _ = self._com()
        pythoncom.CoInitialize()
        return pythoncom

    # Always return a stable Visio instance (visible = True)
    def _get_visio(self):
        _, win32com, _ = self._com()
        try:
            app = win32com.GetActiveObject("Visio.Application")
        except Exception:
            app = win32com.Dispatch("Visio.Application")
        app.Visible = True
        docs = app.Documents
        return app, docs

    def _normalize(self, name: str) -> str:
        return " ".join(name.strip().split()) if name else ""

    def _find_master(self, stencil, name: str):
        if not name:
            return None
        try:
            return stencil.Masters.Item(name)
        except Exception:
            pass
        # Fallback: normalized, case-insensitive scan
        for m in stencil.Masters:
            try:
                if (
                    self._normalize(getattr(m, "Name", "")).lower()
                    == self._normalize(name).lower()
                ):
                    return m
            except Exception:
                continue
        return None

    # ------------------------
    # Shape Data read helpers
    # ------------------------
    def _clean_text(self, s: Any) -> str:
        if s is None:
            return ""
        s = str(s).strip()
        if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
            s = s[1:-1]
        return s.strip()

    def _snapshot_shape_data(self, shape, max_rows: int = 256) -> List[Dict[str, str]]:
        _, _, c = self._com()
        props: List[Dict[str, str]] = []
        for row_idx in range(max_rows):
            try:
                label_exists = shape.CellsSRCExists(
                    c.visSectionProp, row_idx, c.visCustPropsLabel, c.visExistsAnywhere
                )
                value_exists = shape.CellsSRCExists(
                    c.visSectionProp, row_idx, c.visCustPropsValue, c.visExistsAnywhere
                )
            except Exception:
                break

            if not label_exists and not value_exists:
                if props:
                    break
                continue

            label = ""
            value = ""
            if label_exists:
                try:
                    label = self._clean_text(
                        shape.CellsSRC(
                            c.visSectionProp, row_idx, c.visCustPropsLabel
                        ).ResultStr("")
                    )
                except Exception:
                    label = ""
            if value_exists:
                try:
                    value = self._clean_text(
                        shape.CellsSRC(
                            c.visSectionProp, row_idx, c.visCustPropsValue
                        ).ResultStr("")
                    )
                except Exception:
                    value = ""

            if label or value:
                props.append({"label": label, "value": value})
        return props

    def _apply_shape_data_by_label(self, target_shape, props: List[Dict[str, str]], max_rows: int = 256) -> None:
        _, _, c = self._com()

        # Map existing label -> row_idx
        target_labels: Dict[str, int] = {}
        for row_idx in range(max_rows):
            try:
                if target_shape.CellsSRCExists(
                    c.visSectionProp, row_idx, c.visCustPropsLabel, c.visExistsAnywhere
                ):
                    lbl = self._clean_text(
                        target_shape.CellsSRC(
                            c.visSectionProp, row_idx, c.visCustPropsLabel
                        ).ResultStr("")
                    )
                    if lbl:
                        target_labels[lbl] = row_idx
            except Exception:
                break

        for prop in props:
            lbl = prop.get("label", "")
            val = prop.get("value", "")
            if not lbl or lbl not in target_labels:
                continue
            row_idx = target_labels[lbl]
            try:
                cell = target_shape.CellsSRC(c.visSectionProp, row_idx, c.visCustPropsValue)
                safe_val = (val or "").replace('"', '""')
                cell.FormulaU = f'"{safe_val}"'
            except Exception:
                pass

    def _format_shape_data_recursive(self, shape, indent: int = 0) -> List[str]:
        lines: List[str] = []
        pad = "  " * indent

        try:
            nm = (getattr(shape, "Name", "") or "").strip()
        except Exception:
            nm = ""
        try:
            mid = ""
            try:
                if getattr(shape, "Master", None):
                    mid = (shape.Master.Name or "").strip()
            except Exception:
                mid = ""

            lines.append(f"{pad}- {nm or '(unnamed)'} (Master='{mid}')")
        except Exception:
            lines.append(f"{pad}- (shape)")

        props = self._snapshot_shape_data(shape)
        for p in props:
            lines.append(f"{pad}    → Label='{p.get('label','')}' | Value='{p.get('value','')}'")

        # recurse
        try:
            if shape.Shapes is not None and shape.Shapes.Count > 0:
                for sub in shape.Shapes:
                    lines.extend(self._format_shape_data_recursive(sub, indent + 1))
        except Exception:
            pass

        return lines

    # ------------------------
    # Render a single master to a PNG preview (best-effort)
    # Returns: absolute path to generated PNG
    # ------------------------
    def render_master_preview(self, master_name: str) -> str:
        pythoncom = self._co_init()
        stencil = None
        doc = None
        try:
            if not self.current_stencil_path:
                raise Exception("Stencil not loaded.")

            app, docs = self._get_visio()

            # Use a per-operation copy to reduce file-lock issues
            base = os.path.basename(self.current_stencil_path)
            stencil_copy = os.path.join(
                tempfile.gettempdir(),
                f"preview_{int(time.time() * 1000)}_{base}",
            )
            shutil.copy2(self.current_stencil_path, stencil_copy)
            stencil = docs.OpenEx(stencil_copy, 64)

            master = self._find_master(stencil, master_name)
            if master is None:
                raise Exception(f"Master '{master_name}' not found.")

            doc = docs.Add("")
            page = doc.Pages(1)

            # Drop near center of a default page
            try:
                page.Drop(master, 4.25, 5.5)
            except Exception:
                pass

            # Fit view if possible
            try:
                _, _, c = self._com()
                app.ActiveWindow.Zoom = c.visZoomFitPage
            except Exception:
                pass

            out_path = os.path.join(
                tempfile.gettempdir(),
                f"master_preview_{int(time.time() * 1000)}.png",
            )
            page.Export(out_path)

            # Avoid save prompts on close
            try:
                doc.Saved = True
            except Exception:
                pass

            return out_path
        finally:
            try:
                if stencil is not None:
                    stencil.Close()
            except Exception:
                pass
            try:
                if doc is not None:
                    doc.Close()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    # ------------------------
    # Load stencil + extract master names + shape data (best-effort)
    # Returns list of dicts: {'name': <str>, 'props': [{'type':..., 'value':...}, ...]}
    # ------------------------
    def load_stencil_masters(self, stencil_path: str) -> List[Dict[str, Any]]:
        pythoncom = self._co_init()
        try:
            if not os.path.exists(stencil_path):
                raise Exception("Stencil not found.")

            base = os.path.basename(stencil_path)
            temp_copy = os.path.join(
                tempfile.gettempdir(),
                f"stencil_{int(time.time() * 1000)}_{base}",
            )
            shutil.copy2(stencil_path, temp_copy)
            self.current_stencil_path = temp_copy

            app, docs = self._get_visio()
            stencil = docs.OpenEx(temp_copy, 64)

            masters_info: List[Dict[str, Any]] = []
            for m in stencil.Masters:
                # Name normalization
                try:
                    name = self._normalize(m.Name)
                except Exception:
                    name = self._normalize(getattr(m, "Name", "Unknown"))

                props: List[Dict[str, str]] = []
                # 1) Try direct Properties collection (if present)
                try:
                    if hasattr(m, "Properties"):
                        for p in m.Properties:
                            try:
                                p_name = (
                                    getattr(p, "Name", None)
                                    or getattr(p, "Label", None)
                                    or str(p)
                                )
                                p_val: Any = ""
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
                        if getattr(m, "Shapes", None):
                            for shp in m.Shapes:
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
                                        props.append({"type": str(lbl), "value": str(val)})
                                    except Exception:
                                        continue
                                if props:
                                    break
                    except Exception:
                        pass

                masters_info.append({"name": name, "props": props})

            stencil.Close()

            # Deduplicate by name (preserve first)
            seen = set()
            final: List[Dict[str, Any]] = []
            for info in masters_info:
                n = info["name"]
                if n not in seen:
                    seen.add(n)
                    final.append(info)

            return final
        finally:
            pythoncom.CoUninitialize()

    # ------------------------
    # Optional: read element (sub-shape) Shape Data from a master
    # Used for an easy, user-friendly "what's inside this substation" view.
    # ------------------------
    def get_master_elements_shape_data(self, master_name: str) -> List[str]:
        pythoncom = self._co_init()
        stencil = None
        doc = None
        try:
            if not self.current_stencil_path:
                raise Exception("Stencil not loaded.")

            app, docs = self._get_visio()

            base = os.path.basename(self.current_stencil_path)
            stencil_copy = os.path.join(
                tempfile.gettempdir(),
                f"elements_{int(time.time() * 1000)}_{base}",
            )
            shutil.copy2(self.current_stencil_path, stencil_copy)
            stencil = docs.OpenEx(stencil_copy, 64)

            master = self._find_master(stencil, master_name)
            if master is None:
                raise Exception(f"Master '{master_name}' not found.")

            # Drop to a scratch page so we can traverse real Shape objects
            doc = docs.Add("")
            page = doc.Pages(1)
            shp = page.Drop(master, 4.25, 5.5)

            # If master is a group, show data for the whole thing and children
            lines = self._format_shape_data_recursive(shp)

            try:
                doc.Saved = True
            except Exception:
                pass

            return lines
        finally:
            try:
                if stencil is not None:
                    stencil.Close()
            except Exception:
                pass
            try:
                if doc is not None:
                    doc.Close()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    # ------------------------
    # Generate the real Visio layout (NO PREVIEW)
    # - Drops shapes, immediately ungroups them (best-effort)
    # - Adds a bold centered textbox ABOVE each dropped substation if a label was provided
    # ------------------------
    def generate_layout(self, data: Dict[str, Any]) -> None:
        pythoncom = self._co_init()
        stencil = None
        doc = None
        try:
            if not self.current_stencil_path:
                raise Exception("Stencil not loaded.")

            app, docs = self._get_visio()

            # Per-operation stencil copy
            base = os.path.basename(self.current_stencil_path)
            stencil_copy = os.path.join(
                tempfile.gettempdir(),
                f"layout_{int(time.time() * 1000)}_{base}",
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

            # Drop central
            central_name = (data.get("central") or "").strip()
            central_label = (data.get("central_label") or "").strip()
            central_master = self._find_master(stencil, central_name)

            if central_master is None:
                raise Exception(f"Central master '{central_name}' not found.")

            bbox = None
            try:
                central_shape = page.Drop(central_master, CENTRAL_X, CENTRAL_Y)
                bbox = self._safe_bbox(central_shape)
                self._try_ungroup(central_shape)
            except Exception:
                central_shape = None

            # Add label above central (only if user provided)
            if central_label and bbox is not None:
                self._place_textbox_above_bbox(page, bbox, central_label)

            # Matrix
            matrix = data.get("matrix_subs", []) or []
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

                    if not cell:
                        continue

                    # Backwards compat: allow plain string
                    if isinstance(cell, str):
                        name = cell
                        label = ""
                    else:
                        name = (cell.get("master") or "").strip()
                        label = (cell.get("label") or "").strip()

                    if not name or name == "None":
                        continue

                    m_obj = self._find_master(stencil, name)
                    if m_obj is None:
                        continue

                    x = start_x + c * H_SPACING
                    try:
                        shp = page.Drop(m_obj, x, y)
                    except Exception:
                        shp = None

                    if shp is None:
                        continue

                    bbox2 = self._safe_bbox(shp)
                    self._try_ungroup(shp)

                    # place label above dropped substation if provided
                    if label and bbox2 is not None:
                        try:
                            self._place_textbox_above_bbox(page, bbox2, label)
                        except Exception:
                            pass

            # Fit view
            try:
                _, _, c = self._com()
                app.ActiveWindow.Zoom = c.visZoomFitPage
            except Exception:
                pass

            try:
                stencil.Close()
            except Exception:
                pass
        finally:
            pythoncom.CoUninitialize()

    def _safe_bbox(self, shp) -> Optional[Tuple[float, float, float, float]]:
        try:
            _, _, c = self._com()
            bbox = shp.BoundingBox(c.visBBoxUpright)
            # (Left, Bottom, Right, Top)
            return (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
        except Exception:
            return None

    def _try_ungroup(self, shp) -> None:
        """Ungroup a dropped substation (best-effort)."""
        try:
            # Only groups can be ungrouped; calling on non-groups throws.
            shp.Ungroup()
        except Exception:
            pass

    def _place_textbox_above_bbox(self, page, bbox: Tuple[float, float, float, float], label_text: str) -> None:
        """Place a bold centered textbox slightly above the given bbox."""
        _, _, c = self._com()
        left, bottom, right, top = bbox
        width = max(2.0, min(5.0, (right - left) * 0.9))
        height = 0.8

        cx = (left + right) / 2.0

        margin = 0.35
        x1 = cx - width / 2.0
        y1 = top + margin
        x2 = cx + width / 2.0
        y2 = y1 + height

        t = page.DrawRectangle(x1, y1, x2, y2)
        try:
            t.Text = label_text
        except Exception:
            pass

        # Style to match your reference snippet (bold + centered)
        try:
            t.CellsU("Char.Bold").FormulaU = "TRUE"
        except Exception:
            try:
                t.CellsU("Char.Bold").FormulaU = "1"
            except Exception:
                pass

        try:
            t.CellsU("Char.Size").ResultIU = 14
        except Exception:
            pass
        try:
            t.CellsU("Para.HorzAlign").FormulaU = "1"  # center
        except Exception:
            pass
        try:
            t.CellsU("VertAlign").FormulaU = "1"  # middle
        except Exception:
            pass

        # Make it look like a label (no border/fill)
        try:
            t.CellsU("LinePattern").FormulaU = "0"
        except Exception:
            pass
        try:
            t.CellsU("FillPattern").FormulaU = "0"
        except Exception:
            pass

        try:
            t.BringToFront()
        except Exception:
            pass

    # ------------------------
    # Active-document Inspector / Replace
    # ------------------------
    def list_shapes_in_active_document(self) -> List[Dict[str, Any]]:
        pythoncom = self._co_init()
        try:
            app, _ = self._get_visio()
            doc = app.ActiveDocument
            out: List[Dict[str, Any]] = []
            for page in doc.Pages:
                for shape in page.Shapes:
                    out.append(
                        {
                            "page": str(getattr(page, "Name", "")),
                            "shape_id": int(getattr(shape, "ID", 0)),
                            "shape_name": str(getattr(shape, "Name", "")),
                            "master": self._get_master_name(shape),
                        }
                    )
            return out
        finally:
            pythoncom.CoUninitialize()

    def get_shape_data_from_active_document(self, page_name: str, shape_id: int) -> List[str]:
        pythoncom = self._co_init()
        try:
            app, _ = self._get_visio()
            doc = app.ActiveDocument
            page = self._find_page_by_name(doc, page_name)
            if page is None:
                raise RuntimeError(f"Page not found: {page_name}")
            shape = self._find_shape_by_id(page, shape_id)
            if shape is None:
                raise RuntimeError(f"Shape not found (ID={shape_id}) on page '{page_name}'")
            return self._format_shape_data_recursive(shape)
        finally:
            pythoncom.CoUninitialize()

    def replace_shape_in_active_document(self, page_name: str, shape_id: int, stencil_path: str, master_name: str) -> None:
        pythoncom = self._co_init()
        stencil_doc = None
        try:
            if not os.path.isfile(stencil_path):
                raise RuntimeError(f"Stencil file not found: {stencil_path}")

            app, _ = self._get_visio()
            doc = app.ActiveDocument
            page = self._find_page_by_name(doc, page_name)
            if page is None:
                raise RuntimeError(f"Page not found: {page_name}")
            old_shape = self._find_shape_by_id(page, shape_id)
            if old_shape is None:
                raise RuntimeError(f"Shape not found (ID={shape_id}) on page '{page_name}'")

            # Open stencil read-only
            stencil_doc = app.Documents.OpenEx(stencil_path, 64)  # visOpenRO

            new_master = None
            try:
                new_master = stencil_doc.Masters.ItemU(master_name)
            except Exception:
                try:
                    new_master = stencil_doc.Masters.Item(master_name)
                except Exception:
                    new_master = None

            if new_master is None:
                raise RuntimeError(f"Master '{master_name}' not found in stencil.")

            self._replace_shape_with_master(page, old_shape, new_master)

        finally:
            try:
                if stencil_doc is not None:
                    stencil_doc.Close()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    def list_masters_in_stencil(self, stencil_path: str) -> List[str]:
        pythoncom = self._co_init()
        stencil_doc = None
        try:
            if not os.path.isfile(stencil_path):
                raise RuntimeError(f"Stencil file not found: {stencil_path}")
            app, _ = self._get_visio()
            stencil_doc = app.Documents.OpenEx(stencil_path, 64)  # visOpenRO
            names: List[str] = []
            try:
                for m in stencil_doc.Masters:
                    nm = (getattr(m, "Name", "") or "").strip()
                    if nm:
                        names.append(nm)
            except Exception:
                pass
            return names
        finally:
            try:
                if stencil_doc is not None:
                    stencil_doc.Close()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    def bring_visio_to_front(self) -> None:
        pythoncom = self._co_init()
        try:
            app, _ = self._get_visio()
            try:
                app.Visible = True
            except Exception:
                pass
            try:
                win = app.ActiveWindow
                if win:
                    # best-effort maximize
                    _, _, c = self._com()
                    try:
                        win.WindowState = c.visWindowStateMaximized
                    except Exception:
                        pass
            except Exception:
                pass
        finally:
            pythoncom.CoUninitialize()

    def _get_master_name(self, shape) -> str:
        try:
            if shape.Master:
                return (shape.Master.Name or "").strip()
        except Exception:
            pass
        return ""

    def _find_page_by_name(self, doc, page_name: str):
        try:
            return doc.Pages.Item(page_name)
        except Exception:
            pass
        try:
            for p in doc.Pages:
                if (getattr(p, "Name", "") or "") == page_name:
                    return p
        except Exception:
            pass
        return None

    def _find_shape_by_id(self, page, shape_id: int):
        try:
            return page.Shapes.ItemFromID(shape_id)
        except Exception:
            pass
        # fallback scan
        try:
            for s in page.Shapes:
                if int(getattr(s, "ID", -1)) == int(shape_id):
                    return s
        except Exception:
            pass
        return None

    def _replace_shape_with_master(self, page, old_shape, new_master) -> None:
        # Snapshot Shape Data and text
        props = self._snapshot_shape_data(old_shape)
        try:
            old_text = old_shape.Text or ""
        except Exception:
            old_text = ""

        # Snapshot geometry & position (internal units)
        def _iu(cell_name: str) -> float:
            try:
                return float(old_shape.CellsU(cell_name).ResultIU)
            except Exception:
                return 0.0

        pinx = _iu("PinX")
        piny = _iu("PinY")
        width = _iu("Width")
        height = _iu("Height")
        angle = _iu("Angle")
        locpinx = _iu("LocPinX")
        locpiny = _iu("LocPinY")

        new_shape = page.Drop(new_master, pinx, piny)

        # Preserve size/rotation/position
        try:
            new_shape.CellsU("Width").ResultIU = width
            new_shape.CellsU("Height").ResultIU = height
            new_shape.CellsU("Angle").ResultIU = angle
            new_shape.CellsU("LocPinX").ResultIU = locpinx
            new_shape.CellsU("LocPinY").ResultIU = locpiny
            new_shape.CellsU("PinX").ResultIU = pinx
            new_shape.CellsU("PinY").ResultIU = piny
        except Exception:
            pass

        # Preserve text
        try:
            if old_text:
                new_shape.Text = old_text
        except Exception:
            pass

        # Apply Shape Data values by matching labels (do NOT create new rows)
        self._apply_shape_data_by_label(new_shape, props)

        # Delete old shape
        try:
            old_shape.Delete()
        except Exception:
            pass

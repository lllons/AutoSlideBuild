"""Progressively rebuild a PowerPoint file as a Google Slides presentation.

Requires:
    python -m pip install python-pptx pillow google-api-python-client \
        google-auth-httplib2 google-auth-oauthlib pynput

Google setup is described in the UI and in the instructions accompanying this file.
"""

from __future__ import annotations

import io
import queue
import threading
import time
import webbrowser
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from PIL import Image

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError

try:
    from pynput import keyboard
except Exception:
    keyboard = None


SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive.file",
]
EMU_PER_POINT = 12700.0

# Common python-pptx auto-shape names mapped to Google Slides shape types.
SHAPE_TYPES = {
    "RECTANGLE": "RECTANGLE",
    "ROUNDED_RECTANGLE": "ROUND_RECTANGLE",
    "ROUND_RECTANGLE": "ROUND_RECTANGLE",
    "OVAL": "ELLIPSE",
    "ISOSCELES_TRIANGLE": "TRIANGLE",
    "RIGHT_TRIANGLE": "RIGHT_TRIANGLE",
    "PARALLELOGRAM": "PARALLELOGRAM",
    "TRAPEZOID": "TRAPEZOID",
    "HEXAGON": "HEXAGON",
    "PENTAGON": "PENTAGON",
    "OCTAGON": "OCTAGON",
    "CHEVRON": "CHEVRON",
    "RIGHT_ARROW": "RIGHT_ARROW",
    "LEFT_ARROW": "LEFT_ARROW",
    "UP_ARROW": "UP_ARROW",
    "DOWN_ARROW": "DOWN_ARROW",
    "STAR_5_POINT": "STAR_5",
    "STAR_6_POINT": "STAR_6",
    "STAR_8_POINT": "STAR_8",
    "HEART": "HEART",
    "LIGHTNING_BOLT": "LIGHTNING_BOLT",
    "CLOUD": "CLOUD",
    "WAVE": "WAVE",
    "ARC": "ARC",
    "PIE": "PIE",
    "DONUT": "DONUT",
}


def rgb_dict(color: Any) -> dict | None:
    """Convert a python-pptx RGB colour to the Slides API RGB structure."""
    try:
        if color is None:
            return None
        value = str(color)
        if len(value) != 6:
            return None
        return {
            "red": int(value[0:2], 16) / 255.0,
            "green": int(value[2:4], 16) / 255.0,
            "blue": int(value[4:6], 16) / 255.0,
        }
    except Exception:
        return None


def get_solid_fill_color(fill: Any) -> dict | None:
    try:
        if fill.type is None:
            return None
        # MSO_FILL.SOLID has integer value 1 in python-pptx.
        if int(fill.type) != 1:
            return None
        return rgb_dict(fill.fore_color.rgb)
    except Exception:
        return None


def get_line_color(line: Any) -> dict | None:
    try:
        if line.fill.type is None:
            return None
        if int(line.fill.type) != 1:
            return None
        return rgb_dict(line.color.rgb)
    except Exception:
        return None


def text_shape_type(shape: Any) -> str:
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            name = shape.auto_shape_type.name
            return SHAPE_TYPES.get(name, "RECTANGLE")
        return "TEXT_BOX"
    except Exception:
        return "TEXT_BOX"


def shape_geometry(shape: Any, sx: float, sy: float, offset_x: float = 0.0, offset_y: float = 0.0) -> dict:
    x = max(0.0, float(shape.left) / EMU_PER_POINT * sx + offset_x)
    y = max(0.0, float(shape.top) / EMU_PER_POINT * sy + offset_y)
    w = max(1.0, float(shape.width) / EMU_PER_POINT * sx)
    h = max(1.0, float(shape.height) / EMU_PER_POINT * sy)
    return {
        "size": {
            "width": {"magnitude": w, "unit": "PT"},
            "height": {"magnitude": h, "unit": "PT"},
        },
        "transform": {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": x,
            "translateY": y,
            "unit": "PT",
        },
    }


def make_shape_properties(shape: Any) -> tuple[dict, list[str]]:
    properties: dict = {}
    fields: list[str] = []

    try:
        fill_color = get_solid_fill_color(shape.fill)
        if fill_color:
            properties["shapeBackgroundFill"] = {
                "solidFill": {"color": {"rgbColor": fill_color}, "alpha": 1.0}
            }
            fields.append("shapeBackgroundFill.solidFill")
        else:
            properties["shapeBackgroundFill"] = {"propertyState": "NOT_RENDERED"}
            fields.append("shapeBackgroundFill")
    except Exception:
        pass

    try:
        outline_color = get_line_color(shape.line)
        if outline_color:
            line_props: dict = {
                "solidFill": {"color": {"rgbColor": outline_color}, "alpha": 1.0}
            }
            try:
                line_props["weight"] = {
                    "magnitude": max(0.25, float(shape.line.width) / EMU_PER_POINT),
                    "unit": "PT",
                }
            except Exception:
                pass
            properties["outline"] = line_props
            fields.append("outline.solidFill")
            if "weight" in line_props:
                fields.append("outline.weight")
        else:
            properties["outline"] = {"propertyState": "NOT_RENDERED"}
            fields.append("outline")
    except Exception:
        pass

    return properties, fields


def alignment_name(value: Any) -> str | None:
    return {
        PP_ALIGN.LEFT: "START",
        PP_ALIGN.CENTER: "CENTER",
        PP_ALIGN.RIGHT: "END",
        PP_ALIGN.JUSTIFY: "JUSTIFIED",
    }.get(value)


def paragraph_insert_requests(paragraph: Any, object_id: str, offset: int, add_newline: bool) -> tuple[list[dict], int]:
    """Insert one paragraph and apply its run formatting in the same API call."""
    runs = list(paragraph.runs)
    paragraph_text = "".join(run.text for run in runs) if runs else paragraph.text
    insert_text = paragraph_text + ("\n" if add_newline else "")
    requests: list[dict] = [{
        "insertText": {
            "objectId": object_id,
            "text": insert_text,
            "insertionIndex": offset,
        }
    }]

    if paragraph_text:
        alignment = alignment_name(paragraph.alignment)
        if alignment:
            requests.append({
                "updateParagraphStyle": {
                    "objectId": object_id,
                    "textRange": {
                        "type": "FIXED_RANGE",
                        "startIndex": offset,
                        "endIndex": offset + len(paragraph_text),
                    },
                    "style": {"alignment": alignment},
                    "fields": "alignment",
                }
            })

    run_cursor = offset
    for run in runs:
        run_text = run.text or ""
        run_end = run_cursor + len(run_text)
        if run_text:
            style: dict = {}
            fields: list[str] = []
            font = run.font
            try:
                if font.name:
                    style["fontFamily"] = font.name
                    fields.append("fontFamily")
            except Exception:
                pass
            try:
                if font.size:
                    style["fontSize"] = {"magnitude": float(font.size.pt), "unit": "PT"}
                    fields.append("fontSize")
            except Exception:
                pass
            for attr in ("bold", "italic", "underline"):
                try:
                    value = getattr(font, attr)
                    if value is not None:
                        style[attr] = bool(value)
                        fields.append(attr)
                except Exception:
                    pass
            try:
                color = rgb_dict(font.color.rgb)
                if color:
                    style["foregroundColor"] = {"opaqueColor": {"rgbColor": color}}
                    fields.append("foregroundColor")
            except Exception:
                pass
            if style:
                requests.append({
                    "updateTextStyle": {
                        "objectId": object_id,
                        "textRange": {
                            "type": "FIXED_RANGE",
                            "startIndex": run_cursor,
                            "endIndex": run_end,
                        },
                        "style": style,
                        "fields": ",".join(fields),
                    }
                })
        run_cursor = run_end

    return requests, len(insert_text)


class SlidesRebuilder:
    def __init__(self, app: "BuilderApp"):
        self.app = app
        self.stop_event = threading.Event()
        self.temp_drive_files: list[str] = []
        self.listener = None
        self._hotkey_progress = 0
        self._hotkey_lock = threading.Lock()

    def stop(self):
        self.stop_event.set()

    def _on_global_key(self, key):
        try:
            char = key.char
        except Exception:
            return
        self._accept_hotkey_char(char)

    def _accept_hotkey_char(self, char: str):
        sequence = ["[", "]", "\\"]
        with self._hotkey_lock:
            if char == sequence[self._hotkey_progress]:
                self._hotkey_progress += 1
                if self._hotkey_progress == len(sequence):
                    self.stop_event.set()
                    self.app.events.put(("status", "Emergency stop sequence detected."))
                    self._hotkey_progress = 0
            else:
                self._hotkey_progress = 1 if char == "[" else 0

    def authenticate(self, credentials_path: Path):
        token_path = credentials_path.with_name("token.json")
        creds = None
        if token_path.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
            except Exception:
                creds = None

        scopes_ok = bool(creds and creds.has_scopes(SCOPES))
        if not creds or not creds.valid or not scopes_ok:
            if creds and creds.expired and creds.refresh_token and scopes_ok:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
                creds = flow.run_local_server(port=0)
            token_path.write_text(creds.to_json(), encoding="utf-8")

        slides = build("slides", "v1", credentials=creds, cache_discovery=False)
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        return slides, drive

    @staticmethod
    def _point_value(page_size: dict, dimension: str) -> float:
        item = page_size[dimension]
        magnitude = float(item["magnitude"])
        unit = item.get("unit", "PT")
        if unit == "EMU":
            return magnitude / EMU_PER_POINT
        if unit == "INCHES":
            return magnitude * 72.0
        return magnitude

    def _upload_image_temporarily(self, drive, shape) -> tuple[str, str]:
        blob = shape.image.blob
        content_type = shape.image.content_type or "image/png"
        if content_type not in ("image/png", "image/jpeg", "image/gif"):
            with Image.open(io.BytesIO(blob)) as im:
                converted = io.BytesIO()
                im.convert("RGBA").save(converted, format="PNG")
                blob = converted.getvalue()
                content_type = "image/png"

        media = MediaIoBaseUpload(io.BytesIO(blob), mimetype=content_type, resumable=False)
        uploaded = drive.files().create(
            body={"name": f"slides-builder-temp-{int(time.time())}.img"},
            media_body=media,
            fields="id",
        ).execute()
        file_id = uploaded["id"]
        self.temp_drive_files.append(file_id)
        drive.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
            fields="id",
        ).execute()
        return f"https://drive.google.com/uc?export=download&id={file_id}", file_id

    def _delete_temp_drive_file(self, drive, file_id: str):
        try:
            drive.files().delete(fileId=file_id).execute()
            try:
                self.temp_drive_files.remove(file_id)
            except ValueError:
                pass
        except Exception as exc:
            # Keep the ID for the end-of-run cleanup attempt.
            self.app.events.put(("log", f"Temporary image cleanup will be retried later: {exc}"))

    def _element_base_requests(self, shape, page_id: str, object_id: str, sx: float, sy: float, offset_x: float, offset_y: float, drive) -> tuple[list[dict], list[tuple[str, Callable[[], None]]]]:
        """Create the element shell and return optional follow-up text/cell actions."""
        geom = shape_geometry(shape, sx, sy, offset_x, offset_y)
        followups: list[tuple[str, Callable[[], None]]] = []

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            def insert_picture(sh=shape, pid=page_id, oid=object_id, geometry=geom):
                image_url, temp_file_id = self._upload_image_temporarily(drive, sh)
                req = {
                    "createImage": {
                        "objectId": oid,
                        "url": image_url,
                        "elementProperties": {"pageObjectId": pid, **geometry},
                    }
                }
                self._slides_api.presentations().batchUpdate(
                    presentationId=self._presentation_id, body={"requests": [req]}
                ).execute()
                # Slides stores its own copy after a successful createImage request.
                self._delete_temp_drive_file(drive, temp_file_id)
            followups.append(("Insert image", insert_picture))
            return ([], followups)

        if getattr(shape, "has_table", False):
            table = shape.table
            reqs = [{
                "createTable": {
                    "objectId": object_id,
                    "rows": len(table.rows),
                    "columns": len(table.columns),
                    "elementProperties": {"pageObjectId": page_id, **geom},
                }
            }]
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    cell_text = cell.text or ""
                    if cell_text:
                        def add_cell_text(text_value=cell_text, oid=object_id, row_index=r_idx, col_index=c_idx):
                            req = {
                                "insertText": {
                                    "objectId": oid,
                                    "cellLocation": {"rowIndex": row_index, "columnIndex": col_index},
                                    "text": text_value,
                                    "insertionIndex": 0,
                                }
                            }
                            self._slides_api.presentations().batchUpdate(
                                presentationId=self._presentation_id, body={"requests": [req]}
                            ).execute()
                        followups.append((f"Write table cell {r_idx + 1},{c_idx + 1}", add_cell_text))
            return reqs, followups

        try:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                raise ValueError("Grouped PowerPoint objects are not supported by this converter")
            unsupported = {"CHART", "MEDIA"}
            if getattr(shape.shape_type, "name", "") in unsupported:
                raise ValueError(f"PowerPoint {shape.shape_type.name.lower()} objects are not supported")
        except AttributeError:
            pass

        shape_type = text_shape_type(shape)
        reqs = [{
            "createShape": {
                "objectId": object_id,
                "shapeType": shape_type,
                "elementProperties": {"pageObjectId": page_id, **geom},
            }
        }]
        properties, fields = make_shape_properties(shape)
        if properties and fields:
            reqs.append({
                "updateShapeProperties": {
                    "objectId": object_id,
                    "shapeProperties": properties,
                    "fields": ",".join(fields),
                }
            })

        if getattr(shape, "has_text_frame", False):
            paragraphs = list(shape.text_frame.paragraphs)
            offset = 0
            for p_index, paragraph in enumerate(paragraphs):
                runs = list(paragraph.runs)
                paragraph_text = "".join(run.text for run in runs) if runs else paragraph.text
                add_newline = p_index < len(paragraphs) - 1
                if not paragraph_text and not add_newline:
                    continue
                def insert_paragraph(p=paragraph, oid=object_id, start=offset, newline=add_newline):
                    requests, _ = paragraph_insert_requests(p, oid, start, newline)
                    self._slides_api.presentations().batchUpdate(
                        presentationId=self._presentation_id, body={"requests": requests}
                    ).execute()
                followups.append((f"Write slide {self._current_slide_number} text paragraph {p_index + 1}", insert_paragraph))
                offset += len(paragraph_text) + (1 if add_newline else 0)

        return reqs, followups

    def run(self, pptx_path: Path, credentials_path: Path, duration_seconds: int):
        try:
            self.app.events.put(("status", "Reading PowerPoint file…"))
            prs = Presentation(str(pptx_path))
            if not prs.slides:
                raise ValueError("The selected PowerPoint file contains no slides.")

            self.app.events.put(("status", "Signing in to Google…"))
            slides_api, drive_api = self.authenticate(credentials_path)
            if self.stop_event.is_set():
                raise InterruptedError("Stopped before presentation creation.")

            title = f"{pptx_path.stem} (rebuilt)"
            created = slides_api.presentations().create(body={"title": title}).execute()
            presentation_id = created["presentationId"]
            presentation = slides_api.presentations().get(presentationId=presentation_id).execute()
            page_size = presentation.get("pageSize", {})
            target_w = self._point_value(page_size, "width") if page_size.get("width") else 720.0
            target_h = self._point_value(page_size, "height") if page_size.get("height") else 405.0
            source_w = float(prs.slide_width) / EMU_PER_POINT
            source_h = float(prs.slide_height) / EMU_PER_POINT
            # Slides API presentations use a fixed page size. Fit the source canvas
            # into it without distorting objects; non-16:9 files receive side/top margins.
            uniform_scale = min(target_w / source_w, target_h / source_h) if source_w and source_h else 1.0
            sx = sy = uniform_scale
            offset_x = (target_w - source_w * uniform_scale) / 2.0
            offset_y = (target_h - source_h * uniform_scale) / 2.0

            initial_slides = presentation.get("slides", [])
            if not initial_slides:
                raise RuntimeError("Google returned a presentation without a starter slide.")
            page_ids = [initial_slides[0]["objectId"]]
            for i in range(1, len(prs.slides)):
                page_ids.append(f"rebuilt_page_{i + 1}")

            operations: list[tuple[str, Callable[[], None]]] = []
            self._slides_api = slides_api
            self._presentation_id = presentation_id

            for slide_index, source_slide in enumerate(prs.slides):
                page_id = page_ids[slide_index]

                if slide_index > 0:
                    def create_slide(pid=page_id, index=slide_index):
                        slides_api.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={"requests": [{"createSlide": {"objectId": pid, "insertionIndex": index}}]},
                        ).execute()
                    operations.append((f"Create slide {slide_index + 1}", create_slide))
                else:
                    # Keep an operation for a genuinely blank one-slide deck too.
                    operations.append(("Prepare slide 1", lambda: None))

                try:
                    bg_color = get_solid_fill_color(source_slide.background.fill)
                except Exception:
                    bg_color = None
                if bg_color:
                    def set_background(pid=page_id, rgb=bg_color):
                        req = {
                            "updatePageProperties": {
                                "objectId": pid,
                                "pageProperties": {
                                    "pageBackgroundFill": {
                                        "solidFill": {"color": {"rgbColor": rgb}, "alpha": 1.0}
                                    }
                                },
                                "fields": "pageBackgroundFill.solidFill",
                            }
                        }
                        slides_api.presentations().batchUpdate(
                            presentationId=presentation_id, body={"requests": [req]}
                        ).execute()
                    operations.append((f"Set background for slide {slide_index + 1}", set_background))

                for shape_index, shape in enumerate(source_slide.shapes):
                    obj_id = f"s{slide_index + 1}_e{shape_index + 1}"
                    self._current_slide_number = slide_index + 1
                    try:
                        base_requests, followups = self._element_base_requests(
                            shape, page_id, obj_id, sx, sy, offset_x, offset_y, drive_api
                        )
                        def add_element(requests=base_requests):
                            if requests:
                                slides_api.presentations().batchUpdate(
                                    presentationId=presentation_id, body={"requests": requests}
                                ).execute()
                        if base_requests:
                            operations.append((f"Create slide {slide_index + 1}, element {shape_index + 1}", add_element))
                        operations.extend(followups)
                    except Exception as exc:
                        self.app.events.put(("log", f"Skipped element {shape_index + 1} on slide {slide_index + 1}: {exc}"))

            if not operations:
                raise RuntimeError("No slide content was found to rebuild.")

            url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            webbrowser.open(url)
            self.app.events.put(("created", url))
            self.app.events.put(("status", "Rebuilding in Google Slides…"))
            self.app.events.put(("progress_total", len(operations)))

            # Start the duration clock after login and presentation setup.
            start_time = time.monotonic()
            total = len(operations)
            warning_count = 0

            if keyboard is not None:
                try:
                    self.listener = keyboard.Listener(on_press=self._on_global_key)
                    self.listener.start()
                except Exception as exc:
                    self.listener = None
                    self.app.events.put(("log", f"Global emergency hotkey could not start: {exc}. The Stop button remains available."))

            for index, (label, action) in enumerate(operations, start=1):
                if self.stop_event.is_set():
                    break
                try:
                    action()
                except Exception as exc:
                    warning_count += 1
                    self.app.events.put(("log", f"WARNING — {label}: {exc}"))

                self.app.events.put(("progress", (index, total, label)))
                target_time = start_time + (duration_seconds * index / total)
                while not self.stop_event.is_set():
                    remaining = target_time - time.monotonic()
                    if remaining <= 0:
                        break
                    time.sleep(min(0.15, remaining))

            stopped = self.stop_event.is_set()
            if not stopped:
                remaining = start_time + duration_seconds - time.monotonic()
                while remaining > 0 and not self.stop_event.is_set():
                    time.sleep(min(0.15, remaining))
                    remaining = start_time + duration_seconds - time.monotonic()

            self.cleanup_drive_files(drive_api)
            if stopped:
                self.app.events.put(("finished", ("Stopped. The partially rebuilt presentation is available in Google Slides.", url, warning_count)))
            else:
                self.app.events.put(("finished", ("Finished rebuilding the presentation.", url, warning_count)))

        except InterruptedError as exc:
            self.app.events.put(("finished", (str(exc), "", 0)))
        except HttpError as exc:
            self.cleanup_drive_files(locals().get("drive_api"))
            self.app.events.put(("error", f"Google API error: {exc}"))
        except Exception as exc:
            self.cleanup_drive_files(locals().get("drive_api"))
            self.app.events.put(("error", f"{type(exc).__name__}: {exc}"))
        finally:
            if self.listener is not None:
                try:
                    self.listener.stop()
                except Exception:
                    pass

    def cleanup_drive_files(self, drive_api):
        if drive_api is None:
            return
        for file_id in list(self.temp_drive_files):
            try:
                drive_api.files().delete(fileId=file_id).execute()
            except Exception:
                self.app.events.put(("log", f"Could not remove temporary image file {file_id} from Drive."))
            finally:
                try:
                    self.temp_drive_files.remove(file_id)
                except ValueError:
                    pass


class BuilderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PowerPoint → Google Slides Builder")
        self.root.geometry("690x535")
        self.root.minsize(620, 480)
        self.events: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.rebuilder: SlidesRebuilder | None = None
        self.file_var = tk.StringVar()
        default_credentials = Path(__file__).with_name("credentials.json")
        self.credentials_var = tk.StringVar(value=str(default_credentials) if default_credentials.exists() else "")
        self.minutes_var = tk.StringVar(value="60")
        self.last_url = ""

        self._build_ui()
        if keyboard is None:
            self.root.bind_all("<KeyPress>", self._on_tk_key)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_events)

    def _build_ui(self):
        pad = {"padx": 14, "pady": 7}
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Rebuild a PPTX in Google Slides", font=("TkDefaultFont", 16, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(outer, text="Choose a PowerPoint file, set the build duration, then press Start.").pack(anchor="w", pady=(0, 12))

        file_frame = ttk.LabelFrame(outer, text="PowerPoint file", padding=10)
        file_frame.pack(fill="x", **pad)
        ttk.Entry(file_frame, textvariable=self.file_var).pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="Browse…", command=self._choose_pptx).pack(side="left", padx=(8, 0))

        cred_frame = ttk.LabelFrame(outer, text="Google OAuth credentials (credentials.json)", padding=10)
        cred_frame.pack(fill="x", **pad)
        ttk.Entry(cred_frame, textvariable=self.credentials_var).pack(side="left", fill="x", expand=True)
        ttk.Button(cred_frame, text="Browse…", command=self._choose_credentials).pack(side="left", padx=(8, 0))

        settings = ttk.Frame(outer)
        settings.pack(fill="x", **pad)
        ttk.Label(settings, text="Build duration (minutes):").pack(side="left")
        self.duration_spin = ttk.Spinbox(settings, from_=1, to=10080, width=8, textvariable=self.minutes_var)
        self.duration_spin.pack(side="left", padx=(8, 18))
        ttk.Label(settings, text="Default: 60 minutes").pack(side="left")

        privacy = (
            "Image note: Google Slides requires image URLs that it can fetch. Embedded images are "
            "temporarily shared by link in your Drive during insertion, then the script deletes the temporary files. "
            "Do not use this mode for confidential images."
        )
        ttk.Label(outer, text=privacy, wraplength=635, foreground="#8a4b08").pack(anchor="w", **pad)

        controls = ttk.Frame(outer)
        controls.pack(fill="x", **pad)
        self.start_button = ttk.Button(controls, text="Start", command=self._start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="Stop", command=self._stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        self.open_button = ttk.Button(controls, text="Open Google Slides", command=self._open_last, state="disabled")
        self.open_button.pack(side="left", padx=8)

        self.progress = ttk.Progressbar(outer, mode="determinate", maximum=100)
        self.progress.pack(fill="x", **pad)
        self.status_label = ttk.Label(outer, text="Ready.")
        self.status_label.pack(anchor="w", **pad)

        log_frame = ttk.LabelFrame(outer, text="Activity log", padding=6)
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(log_frame, height=9, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True)

        footer = ttk.Label(outer, text="Emergency stop: press [ then ] then \\ (works globally while the build is running).", foreground="#555555")
        footer.pack(anchor="w", padx=14, pady=(4, 0))

    def _append_log(self, message: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _choose_pptx(self):
        path = filedialog.askopenfilename(
            title="Choose a PowerPoint file",
            filetypes=[("PowerPoint presentations", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self.file_var.set(path)

    def _choose_credentials(self):
        path = filedialog.askopenfilename(
            title="Choose Google OAuth credentials.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            self.credentials_var.set(path)

    def _start(self):
        pptx_path = Path(self.file_var.get().strip())
        credentials_path = Path(self.credentials_var.get().strip())
        try:
            minutes = int(self.minutes_var.get())
            if minutes < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid duration", "Enter a whole number of minutes (at least 1).")
            return
        if not pptx_path.is_file() or pptx_path.suffix.lower() != ".pptx":
            messagebox.showerror("Choose a PPTX", "Select an existing .pptx file.")
            return
        if not credentials_path.is_file():
            messagebox.showerror("Credentials required", "Select the Google OAuth Desktop client credentials JSON file.")
            return

        self.progress.configure(value=0, maximum=100)
        self.last_url = ""
        self.open_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_label.configure(text="Starting…")
        self._append_log(f"Source: {pptx_path.name}")
        self._append_log(f"Build time: {minutes} minute(s)")
        self.rebuilder = SlidesRebuilder(self)
        self.worker = threading.Thread(
            target=self.rebuilder.run,
            args=(pptx_path, credentials_path, minutes * 60),
            daemon=True,
        )
        self.worker.start()

    def _stop(self):
        if self.rebuilder:
            self.rebuilder.stop()
            self.status_label.configure(text="Stop requested — finishing the current API request…")
            self.stop_button.configure(state="disabled")

    def _open_last(self):
        if self.last_url:
            webbrowser.open(self.last_url)

    def _on_tk_key(self, event):
        if self.rebuilder and self.worker and self.worker.is_alive():
            self.rebuilder._accept_hotkey_char(event.char)

    def _poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "status":
                    self.status_label.configure(text=payload)
                elif kind == "log":
                    self._append_log(payload)
                elif kind == "created":
                    self.last_url = payload
                    self.open_button.configure(state="normal")
                    self._append_log("Created Google Slides presentation: " + payload)
                elif kind == "progress_total":
                    self._total_actions = max(1, int(payload))
                elif kind == "progress":
                    index, total, label = payload
                    self.progress.configure(value=(index / total) * 100)
                    self.status_label.configure(text=f"{index}/{total}: {label}")
                elif kind == "finished":
                    message, url, warnings = payload
                    if url:
                        self.last_url = url
                        self.open_button.configure(state="normal")
                    self.status_label.configure(text=message)
                    self._append_log(message)
                    if warnings:
                        self._append_log(f"Completed with {warnings} warning(s). See log entries above.")
                    self.start_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                elif kind == "error":
                    self.status_label.configure(text="Stopped because of an error.")
                    self._append_log(payload)
                    self.start_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    messagebox.showerror("Build failed", payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _on_close(self):
        if self.rebuilder and self.worker and self.worker.is_alive():
            self.rebuilder.stop()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    BuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

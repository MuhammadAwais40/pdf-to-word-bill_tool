import os
import sys
import socket
import threading
import time
import traceback
from pathlib import Path
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")
import webview
from waitress import serve
from app import app, build_export, cleanup_old, APP_DATA
APP_TITLE = "PDF to Word & Excel Converter"
def free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]
def run_server(port):
    serve(app, host="127.0.0.1", port=port, threads=4)
def wait_for_server(port, timeout=15):
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.1)
    return False
class Api:
    def __init__(self):
        self._window = None  
    def save_file(self, fid, fmt):
        try:
            buf, name, _ = build_export(fid, fmt)
        except (ValueError, FileNotFoundError) as e:
            return {"ok": False, "error": str(e)}
        except Exception:
            return {"ok": False, "error": "Could not create the file."}
        ext = Path(name).suffix
        label = "Word Document (*.docx)" if ext == ".docx" else "Excel Workbook (*.xlsx)"
        downloads = Path.home() / "Downloads"
        save_type = webview.FileDialog.SAVE if hasattr(webview, "FileDialog") else webview.SAVE_DIALOG
        try:
            result = self._window.create_file_dialog(
                save_type,
                directory=str(downloads) if downloads.exists() else "",
                save_filename=name,
                file_types=(label,),)
        except Exception:
            return {"ok": False, "error": "Could not open the Save dialog."}
        if not result:
            return {"ok": False, "cancelled": True}
        target = result if isinstance(result, str) else result[0]
        if not target.lower().endswith(ext):
            target += ext
        try:
            with open(target, "wb") as f:
                f.write(buf.getvalue())
        except OSError as e:
            return {"ok": False, "error": f"Could not save file: {e.strerror or e}"}
        return {"ok": True, "path": target}
def main():
    cleanup_old()
    port = free_port()
    threading.Thread(target=run_server, args=(port,), daemon=True).start()
    if not wait_for_server(port):
        raise RuntimeError("Local server failed to start.")
    api = Api()
    window = webview.create_window(
        APP_TITLE,
        f"http://127.0.0.1:{port}/",
        js_api=api,
        width=1080,
        height=820,
        min_size=(420, 640),)
    api._window = window
    webview.start()
if __name__ == "__main__":
    try:
        main()
    except Exception:
        try:
            APP_DATA.mkdir(parents=True, exist_ok=True)
            (APP_DATA / "error.log").write_text(traceback.format_exc(), encoding="utf-8")
        except Exception:
            pass
        raise

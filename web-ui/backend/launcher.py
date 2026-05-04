"""
PyInstaller entry point. Starts uvicorn, opens browser, blocks until closed.
"""
import os
import sys
import threading
import time
import webbrowser
from pathlib import Path


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def _setup_paths() -> None:
    """When frozen, _MEIPASS contains bundled deps. Add to sys.path."""
    if _is_frozen():
        mei = Path(getattr(sys, "_MEIPASS", ""))
        for p in [mei, mei / "iTech_Accrual_Updater"]:
            if p.exists() and str(p) not in sys.path:
                sys.path.insert(0, str(p))


def _open_browser_when_ready(port: int) -> None:
    """Wait briefly for server, then open default browser."""
    import urllib.request

    url = f"http://127.0.0.1:{port}"
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            req = urllib.request.Request(url + "/api/health", headers={"Host": "127.0.0.1"})
            urllib.request.urlopen(req, timeout=1)
            break
        except Exception:
            time.sleep(0.3)
    webbrowser.open(url)


def main() -> None:
    _setup_paths()

    port = int(os.environ.get("ACCRUAL_PORT", "8100"))
    os.environ.setdefault("ACCRUAL_ALLOW_MACROS", "1")

    threading.Thread(target=_open_browser_when_ready, args=(port,), daemon=True).start()

    # Import here so _setup_paths has run
    import uvicorn
    from main import app  # type: ignore

    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")


if __name__ == "__main__":
    main()

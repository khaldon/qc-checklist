#!/usr/bin/env python3
"""
QC Toast Checklist  (single-overlay version)
============================================
Every QC check item is its own small, semi-transparent "toast" floating above
every other application (e.g. your 3D dental-arch viewer).

How it is built (why it behaves well)
-------------------------------------
* ALL toasts live inside ONE transparent, frameless, always-on-top overlay window
  that spans the desktop. The OS therefore sees a single tool window, not one per
  toast -> nothing extra in Alt-Tab / Ctrl-Tab / taskbar.
* The empty parts of the overlay are masked out, so mouse clicks pass straight
  through to the application underneath. Only the toast cards catch the mouse.
* Transparency/fade is rendered by Qt itself (per-toast opacity effect), not by the
  OS window-opacity feature, so it works the same everywhere.
* Toasts never take keyboard focus, so your 3D software keeps working normally.
* The overlay re-asserts "topmost" every second so other apps can't bury it.

Interaction
-----------
* Click a toast ............ tick / untick (strikethrough when done)
* Drag a toast ............. move it anywhere (position remembered)
* Mouse wheel on a toast ... change opacity of all toasts
* Right-click a toast ...... Edit / Move up / Move down / Delete + global actions
* Shift + drag any toast ... move ALL toasts together (to peek at what's behind them)
* Tray icon (or right-click menu) -> Add item, Reset all, Re-stack, Opacity, Quit

Global hotkeys (Windows; work while any other app has focus; editable in the JSON file)
* Ctrl+Alt+Space ... tick the NEXT unchecked item and hide its toast
* Ctrl+Alt+Z ....... undo: bring back the last item you ticked this way
* Ctrl+Alt+H ....... show / hide all toasts
* Ctrl+Alt+R ....... reset: untick everything and show all toasts (next arch)

State is stored in ``qc_checklist_state.json`` next to this script / the executable.

Auto-update (Windows .exe built by CI, see .github/workflows/build.yml)
-----------------------------------------------------------------------
On start (and every 6 h) the app asks GitHub Releases for a newer version, downloads
QCChecklist.exe in the background, verifies its SHA-256, and applies it the next time
the app starts. "Check for updates" in the menu installs a newer version immediately
(download -> restart). Background checks never interrupt you; they apply on next start.
Set UPDATE_REPO below to "your-user/your-repo".

Requires: Python 3.10+, PyQt6  ->  pip install PyQt6
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import urllib.request
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QAbstractNativeEventFilter, QEasingCurve, QObject, QPoint, QPropertyAnimation, QRect, Qt, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QAction, QActionGroup, QColor, QGuiApplication, QIcon, QPainter, QPainterPath, QPen, QPixmap, QRegion
from PyQt6.QtWidgets import (
    QApplication,
    QFrame,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMenu,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

APP_TITLE = "QC Toast Checklist"
STATE_FILENAME = "qc_checklist_state.json"
DEFAULT_ITEMS = [
    "Third-molar trimming",
    "Undercut blocking",
    "Black triangles",
    "Gingival margin finish",
]
DEFAULT_OPACITY = 85
MIN_OPACITY = 30
TOAST_WIDTH = 300
TOAST_GAP = 8
DRAG_THRESHOLD = 6  # px of movement before a click becomes a drag

# Global hotkeys (Windows). Format: modifiers+key, e.g. "ctrl+alt+space", "ctrl+shift+f9".
# Override any of them in qc_checklist_state.json under "hotkeys".
DEFAULT_HOTKEYS = {
    "check_next": "ctrl+alt+space",   # tick next unchecked item + hide its toast
    "undo": "ctrl+alt+z",             # restore the last item ticked with the hotkey
    "toggle_visible": "ctrl+alt+h",   # show / hide all toasts
    "reset": "ctrl+alt+r",            # untick all + show all
}

# --- auto-update settings ----------------------------------------------------
APP_VERSION = "0.0.0"                         # stamped automatically by the CI build
UPDATE_REPO = "https://github.com/khaldon/qc-checklist.git"    # <-- set this (public GitHub repo)
ASSET_NAME = "QCChecklist.exe"                # release asset names (must match build.yml)
API_URL = "https://api.github.com/repos/{repo}/releases/latest"
REQUIRE_HTTPS = True
UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000  # re-check every 6 hours


# --------------------------------------------------------------------------- #
# Persistence
# --------------------------------------------------------------------------- #
def is_frozen() -> bool:
    """True when running as the PyInstaller-built .exe."""
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    """Folder of the .exe (or of this script when run from source)."""
    return Path(sys.executable).resolve().parent if is_frozen() else Path(__file__).resolve().parent


def state_file_path() -> Path:
    return app_dir() / STATE_FILENAME


def load_state(path: Path) -> dict[str, Any]:
    """Read state; on any problem return {} (corrupt files are set aside, not overwritten)."""
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            raise ValueError("top-level JSON must be an object")
        return data
    except (OSError, ValueError) as exc:
        print(f"[warn] cannot read {path}: {exc}", file=sys.stderr)
        try:
            path.replace(path.with_suffix(".json.corrupt"))
        except OSError:
            pass
        return {}


def save_state(path: Path, data: dict[str, Any]) -> bool:
    """Atomic write (temp file + replace) so a crash can never corrupt the file."""
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
        return True
    except OSError as exc:
        print(f"[warn] cannot save {path}: {exc}", file=sys.stderr)
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        return False


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def make_tray_icon() -> QIcon:
    """Draw a small green check icon so no external image file is needed."""
    pm = QPixmap(64, 64)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setBrush(QColor("#34a853"))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRoundedRect(4, 4, 56, 56, 12, 12)
    pen = QPen(QColor("white"), 7, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
    p.setPen(pen)
    path = QPainterPath()
    path.moveTo(18, 34)
    path.lineTo(28, 44)
    path.lineTo(46, 21)
    p.drawPath(path)
    p.end()
    return QIcon(pm)


def ask_text(title: str, label: str, text: str = "") -> str | None:
    """Small always-on-top text prompt (so it isn't hidden behind the 3D viewer)."""
    dlg = QInputDialog()
    dlg.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dlg.setWindowTitle(title)
    dlg.setLabelText(label)
    dlg.setTextValue(text)
    dlg.setMinimumWidth(360)
    if dlg.exec() and dlg.textValue().strip():
        return dlg.textValue().strip()
    return None


def repolish(w: QWidget) -> None:
    w.style().unpolish(w)
    w.style().polish(w)


# --------------------------------------------------------------------------- #
# The single overlay window that hosts every toast
# --------------------------------------------------------------------------- #
class Overlay(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool                        # no taskbar / Alt-Tab entry
            | Qt.WindowType.WindowDoesNotAcceptFocus    # never steal focus from the 3D viewer
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        self.sync_geometry()

    def sync_geometry(self) -> None:
        """Cover the whole (multi-monitor) desktop so toasts can be dragged anywhere."""
        screen = QApplication.primaryScreen()
        if screen is not None:
            self.setGeometry(screen.virtualGeometry())

    def refresh_mask(self) -> None:
        """Only the toast rectangles receive the mouse; everything else is click-through."""
        region = QRegion()
        for child in self.children():
            if isinstance(child, Toast) and not child.isHidden():
                region = region.united(QRegion(child.geometry()))
        if region.isEmpty():
            self.setMask(QRegion(QRect(0, 0, 1, 1)))   # 1px mask: practically nothing to hit
        else:
            self.setMask(region)

    # ------------------------------------------------- always-on-top handling #
    def showEvent(self, e) -> None:  # noqa: N802
        super().showEvent(e)
        QTimer.singleShot(0, self._native_tweaks)

    def _native_tweaks(self) -> None:
        """Windows: make it a no-activate tool window (hidden from Alt-Tab) and topmost."""
        if sys.platform != "win32":
            return
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            hwnd = wintypes.HWND(int(self.winId()))
            get_long = getattr(user32, "GetWindowLongPtrW", user32.GetWindowLongW)
            set_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
            get_long.restype = ctypes.c_ssize_t
            get_long.argtypes = [wintypes.HWND, ctypes.c_int]
            set_long.restype = ctypes.c_ssize_t
            set_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
            GWL_EXSTYLE, WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE, WS_EX_APPWINDOW = -20, 0x80, 0x08000000, 0x40000
            ex = get_long(hwnd, GWL_EXSTYLE)
            ex = (ex | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE) & ~WS_EX_APPWINDOW
            set_long(hwnd, GWL_EXSTYLE, ex)
            user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                            ctypes.c_int, ctypes.c_int, wintypes.UINT]
            # NOMOVE|NOSIZE|NOZORDER|NOACTIVATE|FRAMECHANGED -> apply the new style
            user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0004 | 0x0010 | 0x0020)
        except Exception as exc:  # pragma: no cover - platform specific
            print(f"[warn] native window tweaks failed: {exc}", file=sys.stderr)
        self.keep_on_top()

    def keep_on_top(self) -> None:
        """Called every second: other apps (or 'topmost' apps) can't bury the overlay."""
        if not self.isVisible():
            return
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes

                user32 = ctypes.windll.user32
                user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                                ctypes.c_int, ctypes.c_int, wintypes.UINT]
                HWND_TOPMOST = wintypes.HWND(-1)
                # NOMOVE | NOSIZE | NOACTIVATE
                user32.SetWindowPos(wintypes.HWND(int(self.winId())), HWND_TOPMOST, 0, 0, 0, 0,
                                    0x0002 | 0x0001 | 0x0010)
            except Exception:  # pragma: no cover
                pass
        else:
            self.raise_()


# --------------------------------------------------------------------------- #
# One toast == one checklist item  (a child widget of the overlay)
# --------------------------------------------------------------------------- #
class Toast(QFrame):
    def __init__(self, mgr: "Manager", text: str, checked: bool = False,
                 pos: tuple[int, int] | None = None, hint: bool = False) -> None:
        super().__init__(mgr.overlay)
        self.mgr = mgr
        self.hint = hint                      # placeholder toast shown when the list is empty
        self.checked = checked
        self.custom_pos = pos is not None     # True once the user has placed it manually
        self._user_dragging = False
        self._press: QPoint | None = None
        self._offset = QPoint()
        self._moved = False
        self._move_all = False               # Shift+drag moves every toast
        self._last = QPoint()
        self._anim: QPropertyAnimation | None = None
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # Qt-rendered transparency (reliable on every platform)
        self.effect = QGraphicsOpacityEffect(self)
        self.effect.setOpacity(mgr.opacity / 100)
        self.setGraphicsEffect(self.effect)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.card = QFrame()
        self.card.setObjectName("card")
        outer.addWidget(self.card)

        row = QHBoxLayout(self.card)
        row.setContentsMargins(10, 8, 12, 8)
        row.setSpacing(10)

        self.num = QLabel("")
        self.num.setObjectName("num")
        self.num.setFixedWidth(16)
        self.box = QLabel("")
        self.box.setObjectName("box")
        self.box.setFixedSize(22, 22)
        self.box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label = QLabel(text)
        self.label.setObjectName("text")
        self.label.setWordWrap(True)
        # Mouse events fall through to the toast so click/drag work anywhere on it.
        for w in (self.num, self.box, self.label, self.card):
            w.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        if hint:
            self.num.hide()
            self.box.hide()
        else:
            row.addWidget(self.num)
            row.addWidget(self.box)
        row.addWidget(self.label, 1)

        self.setStyleSheet(TOAST_STYLE)
        self._refresh_style()
        self.fit()
        if pos is not None:
            self.place(*pos)

    # ------------------------------------------------------------ appearance #
    @property
    def text(self) -> str:
        return self.label.text()

    def set_text(self, text: str) -> None:
        self.label.setText(text)
        self.fit()

    def set_index(self, n: int) -> None:
        self.num.setText(str(n))

    def set_checked(self, checked: bool) -> None:
        self.checked = checked
        self._refresh_style()

    def _refresh_style(self) -> None:
        done = self.checked and not self.hint
        for w in (self.card, self.box, self.label):
            w.setProperty("done", "true" if done else "false")
            repolish(w)
        self.box.setText("✓" if done else "")
        f = self.label.font()
        f.setStrikeOut(done)
        self.label.setFont(f)

    def fit(self) -> None:
        """Fixed width; height follows the word-wrapped text (computed explicitly for reliability)."""
        avail = TOAST_WIDTH - 10 - 12 - 6          # card margins + border
        min_h = 0
        if not self.hint:
            avail -= 16 + 22 + 2 * 10              # number + box + spacings
            min_h = 22
        text_h = self.label.heightForWidth(max(avail, 50))
        if text_h <= 0:
            text_h = self.label.sizeHint().height()
        self.setFixedSize(TOAST_WIDTH, max(text_h, min_h) + 16 + 2)

    # ------------------------------------------------------------- geometry #
    def global_pos(self) -> QPoint:
        return self.mgr.overlay.mapToGlobal(self.pos())

    def place(self, gx: int, gy: int) -> None:
        """Move to GLOBAL screen coordinates (programmatic move, not a user drag)."""
        self._user_dragging = False
        self.move(self.mgr.overlay.mapFromGlobal(QPoint(gx, gy)))

    def fade_in(self) -> None:
        """Toast-style fade-in using Qt's own opacity effect."""
        self.effect.setOpacity(0.0)
        self.show()
        self._anim = QPropertyAnimation(self.effect, b"opacity", self)
        self._anim.setDuration(220)
        self._anim.setStartValue(0.0)
        self._anim.setEndValue(self.mgr.opacity / 100)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._anim.start()

    def stop_anim(self) -> None:
        if self._anim is not None:
            self._anim.stop()

    # ---------------------------------------------------------------- events #
    def mousePressEvent(self, e) -> None:  # noqa: N802
        if e.button() == Qt.MouseButton.LeftButton:
            self._press = e.globalPosition().toPoint()
            self._last = self._press
            self._offset = self._press - self.global_pos()
            self._moved = False
            self._move_all = bool(e.modifiers() & Qt.KeyboardModifier.ShiftModifier)
            if self._move_all:
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            e.accept()

    def mouseMoveEvent(self, e) -> None:  # noqa: N802
        if self._press is None or not (e.buttons() & Qt.MouseButton.LeftButton):
            return
        gp = e.globalPosition().toPoint()
        if not self._moved:
            if (gp - self._press).manhattanLength() < DRAG_THRESHOLD:
                return
            self._moved = True
            self._user_dragging = not self._move_all
        if self._move_all:                                   # Shift+drag: move the whole group
            d = gp - self._last
            self._last = gp
            self.mgr.move_all(d.x(), d.y())
            return
        ov = self.mgr.overlay
        target = ov.mapFromGlobal(gp - self._offset)
        x = max(0, min(target.x(), ov.width() - self.width()))   # keep inside the desktop
        y = max(0, min(target.y(), ov.height() - self.height()))
        self.move(x, y)

    def mouseReleaseEvent(self, e) -> None:  # noqa: N802
        if e.button() == Qt.MouseButton.LeftButton and self._press is not None:
            if not self._moved and not self.hint:
                self.mgr.toggle(self)
            elif self._moved:
                if not self._move_all:
                    self.custom_pos = True
                self.mgr.schedule_save()
        self._press = None
        self._moved = False
        self._user_dragging = False
        if self._move_all:
            self._move_all = False
            self.setCursor(Qt.CursorShape.PointingHandCursor)

    def moveEvent(self, e) -> None:  # noqa: N802
        super().moveEvent(e)
        self.mgr.refresh_mask()

    def resizeEvent(self, e) -> None:  # noqa: N802
        super().resizeEvent(e)
        self.mgr.refresh_mask()

    def wheelEvent(self, e) -> None:  # noqa: N802
        step = 5 if e.angleDelta().y() > 0 else -5
        self.mgr.set_opacity(self.mgr.opacity + step)
        e.accept()

    def contextMenuEvent(self, e) -> None:  # noqa: N802
        self.mgr.build_menu(self).exec(e.globalPos())


# --------------------------------------------------------------------------- #
# Auto-updater (GitHub Releases)
# --------------------------------------------------------------------------- #
def parse_version(text: str) -> tuple[int, ...]:
    """'v1.0.12' -> (1, 0, 12). Unparseable -> (0,)."""
    nums = re.findall(r"\d+", text or "")
    return tuple(int(n) for n in nums[:4]) or (0,)


def _open_url(url: str, accept: str = "application/vnd.github+json"):
    if REQUIRE_HTTPS and not url.startswith("https://"):
        raise RuntimeError("refusing non-HTTPS update URL")
    req = urllib.request.Request(url, headers={"User-Agent": "QCChecklist-updater", "Accept": accept})
    return urllib.request.urlopen(req, timeout=30)


def find_pending() -> Path | None:
    """A previously downloaded + verified update newer than this build (or None).
    Also cleans up stale/partial download files."""
    best: Path | None = None
    best_v = parse_version(APP_VERSION)
    found = list(app_dir().glob("QCChecklist-*.update"))
    for f in found:
        v = parse_version(f.stem.replace("QCChecklist-", ""))
        if v > best_v:
            best, best_v = f, v
    for f in found + list(app_dir().glob("QCChecklist-*.part")):
        if f != best:
            try:
                f.unlink()
            except OSError:
                pass
    return best


# Replaces the running exe once it has exited, then starts the new one.
# (ping is used as a sleep because `timeout` fails without a console.)
BAT_TEMPLATE = """@echo off
chcp 65001 >nul
set /a tries=0
:wait
ping 127.0.0.1 -n 2 >nul
move /y "{new}" "{target}" >nul 2>&1
if not errorlevel 1 goto done
set /a tries+=1
if %tries% lss 60 goto wait
del "{new}" >nul 2>&1
:done
start "" "{target}"
(goto) 2>nul & del "%~f0"
"""


def launch_updater(new_file: Path) -> bool:
    """Spawn the helper script that swaps the exe after we exit. Windows .exe only."""
    if sys.platform != "win32" or not is_frozen():
        return False
    try:
        bat = app_dir() / "qc_update.bat"
        bat.write_text(BAT_TEMPLATE.format(new=new_file, target=Path(sys.executable).resolve()),
                       encoding="utf-8")
        env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")   # child must not reuse our temp dir
        subprocess.Popen(["cmd.exe", "/c", str(bat)], creationflags=0x08000000,  # CREATE_NO_WINDOW
                         close_fds=True, env=env, cwd=str(app_dir()))
        return True
    except OSError as exc:
        print(f"[warn] cannot launch updater: {exc}", file=sys.stderr)
        return False


class Updater(QObject):
    status = pyqtSignal(str)   # human-readable message for the tray balloon
    ready = pyqtSignal(str, bool)   # (version, was_manual) once an update is downloaded + verified

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._busy = False
        self.pending_file: Path | None = None

    @staticmethod
    def configured() -> bool:
        return "/" in UPDATE_REPO and "YOUR_" not in UPDATE_REPO

    def check(self, manual: bool = False) -> None:
        """Start a background check (never blocks the UI)."""
        if self._busy:
            return
        if not self.configured():
            if manual:
                self.status.emit("Updates are not configured (set UPDATE_REPO in app.py).")
            return
        self._busy = True
        threading.Thread(target=self._worker, args=(manual,), daemon=True).start()

    def _worker(self, manual: bool) -> None:
        try:
            self.run_check(manual)
        except Exception as exc:                 # network down, rate limit, bad release...
            print(f"[warn] update check failed: {exc}", file=sys.stderr)
            if manual:
                self.status.emit(f"Update check failed: {exc}")
        finally:
            self._busy = False

    def run_check(self, manual: bool) -> None:
        with _open_url(API_URL.format(repo=UPDATE_REPO)) as r:
            release = json.loads(r.read().decode("utf-8"))
        latest = parse_version(str(release.get("tag_name", "")))
        if latest <= parse_version(APP_VERSION):
            if manual:
                self.status.emit(f"You're up to date (v{APP_VERSION}).")
            return
        ver = ".".join(map(str, latest))

        if self.pending_file is not None and ver in self.pending_file.name:
            self.ready.emit(ver, manual)
            return
        if not is_frozen():
            self.status.emit(f"Version {ver} is available (running from source: not auto-installing).")
            return

        assets = {a.get("name"): a.get("browser_download_url") for a in release.get("assets", [])}
        exe_url, sha_url = assets.get(ASSET_NAME), assets.get(ASSET_NAME + ".sha256")
        if not exe_url or not sha_url:
            raise RuntimeError("release is missing the exe or its .sha256 checksum")
        if manual:
            self.status.emit(f"Downloading update v{ver}…")
        with _open_url(sha_url, "application/octet-stream") as r:
            expected = r.read().decode("ascii", "ignore").split()[0].strip().lower()

        dest = app_dir() / f"QCChecklist-{ver}.update"
        part = dest.with_suffix(".part")
        digest = hashlib.sha256()
        try:
            with _open_url(exe_url, "application/octet-stream") as r, part.open("wb") as fh:
                while chunk := r.read(1 << 20):
                    fh.write(chunk)
                    digest.update(chunk)
            if digest.hexdigest() != expected:
                raise RuntimeError("downloaded file failed its SHA-256 check")
            with part.open("rb") as fh:
                if fh.read(2) != b"MZ":
                    raise RuntimeError("downloaded file is not a Windows executable")
            os.replace(part, dest)
        finally:
            if part.exists():
                try:
                    part.unlink()
                except OSError:
                    pass
        self.pending_file = dest
        self.ready.emit(ver, manual)


# --------------------------------------------------------------------------- #
# Global hotkeys (Windows RegisterHotKey)
# --------------------------------------------------------------------------- #
_VK_NAMED = {
    "space": 0x20, "enter": 0x0D, "return": 0x0D, "tab": 0x09, "esc": 0x1B, "escape": 0x1B,
    "backspace": 0x08, "insert": 0x2D, "delete": 0x2E, "home": 0x24, "end": 0x23,
    "pageup": 0x21, "pagedown": 0x22, "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
}
_MODS = {"alt": 0x1, "ctrl": 0x2, "control": 0x2, "shift": 0x4, "win": 0x8, "meta": 0x8, "super": 0x8}


def parse_hotkey(spec: str) -> tuple[int, int] | None:
    """'ctrl+alt+space' -> (modifier flags, virtual-key code); None if invalid."""
    if not isinstance(spec, str):
        return None
    mods, vk = 0, None
    for part in (p.strip().lower() for p in spec.split("+")):
        if not part:
            return None
        if part in _MODS:
            mods |= _MODS[part]
        elif vk is not None:
            return None                                   # two non-modifier keys
        elif part in _VK_NAMED:
            vk = _VK_NAMED[part]
        elif re.fullmatch(r"f([1-9]|1[0-9]|2[0-4])", part):
            vk = 0x70 + int(part[1:]) - 1                 # F1..F24
        elif re.fullmatch(r"num[0-9]", part):
            vk = 0x60 + int(part[3:])                     # numpad digits
        elif re.fullmatch(r"[a-z0-9]", part):
            vk = ord(part.upper())
        else:
            return None
    return (mods, vk) if vk is not None else None


def pretty_hotkey(spec: str) -> str:
    return "+".join(p.strip().capitalize() for p in spec.split("+"))


class HotkeyFilter(QAbstractNativeEventFilter):
    """Registers system-wide hotkeys and turns WM_HOTKEY messages into callbacks."""
    WM_HOTKEY = 0x0312
    MOD_NOREPEAT = 0x4000

    def __init__(self) -> None:
        super().__init__()
        self._callbacks: dict[int, Any] = {}
        self._next_id = 1

    def register(self, spec: str, callback) -> bool:
        if sys.platform != "win32":
            return False
        parsed = parse_hotkey(spec)
        if parsed is None:
            return False
        try:
            import ctypes
            mods, vk = parsed
            hid = self._next_id
            if not ctypes.windll.user32.RegisterHotKey(None, hid, mods | self.MOD_NOREPEAT, vk):
                return False                              # usually: another app owns that combo
            self._callbacks[hid] = callback
            self._next_id += 1
            return True
        except Exception as exc:  # pragma: no cover - platform specific
            print(f"[warn] RegisterHotKey failed: {exc}", file=sys.stderr)
            return False

    def unregister_all(self) -> None:
        if sys.platform != "win32":
            return
        try:
            import ctypes
            for hid in list(self._callbacks):
                ctypes.windll.user32.UnregisterHotKey(None, hid)
        except Exception:  # pragma: no cover
            pass
        self._callbacks.clear()

    def nativeEventFilter(self, event_type, message):  # noqa: N802
        if sys.platform == "win32" and bytes(event_type) == b"windows_generic_MSG":
            import ctypes
            from ctypes import wintypes
            msg = wintypes.MSG.from_address(int(message))
            if msg.message == self.WM_HOTKEY:
                cb = self._callbacks.get(msg.wParam)
                if cb is not None:
                    QTimer.singleShot(0, cb)              # run on the next event-loop turn
                    return True, 0
        return False, 0


# --------------------------------------------------------------------------- #
# Manager: owns the overlay, the toasts, the tray icon and persistence
# --------------------------------------------------------------------------- #
class Manager(QObject):
    def __init__(self) -> None:
        super().__init__()
        self.path = state_file_path()
        state = load_state(self.path)
        self.opacity = self._clean_opacity(state.get("opacity", DEFAULT_OPACITY))
        self.overlay = Overlay()
        self.toasts: list[Toast] = []
        self.hint_toast: Toast | None = None
        self.visible_all = True
        self.update_ready: str | None = None
        self.check_history: list[Toast] = []       # toasts ticked via hotkey (for undo)
        self._mask_blocked = False
        self.origin: QPoint | None = None          # where the auto-stack starts (moved by Shift+drag)
        o = state.get("origin")
        if isinstance(o, list) and len(o) == 2:
            try:
                ox, oy = int(o[0]), int(o[1])
                if self._on_screen(ox, oy):
                    self.origin = QPoint(ox, oy)
            except (TypeError, ValueError):
                pass
        self.hotkeys = dict(DEFAULT_HOTKEYS)
        saved_hk = state.get("hotkeys")
        if isinstance(saved_hk, dict):
            for k in DEFAULT_HOTKEYS:
                v = saved_hk.get(k)
                if isinstance(v, str) and parse_hotkey(v):
                    self.hotkeys[k] = v

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(400)
        self._save_timer.timeout.connect(self.save)

        self._top_timer = QTimer(self)               # keep the overlay above everything
        self._top_timer.setInterval(1000)
        self._top_timer.timeout.connect(self.overlay.keep_on_top)
        self._top_timer.start()

        self.tray: QSystemTrayIcon | None = None
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray = QSystemTrayIcon(make_tray_icon())
            self.tray.activated.connect(self._on_tray_activated)
            self.tray_menu = QMenu()
            self.tray_menu.setStyleSheet(MENU_STYLE)
            self.tray_menu.aboutToShow.connect(lambda: self._fill_menu(self.tray_menu, None))
            self.tray.setContextMenu(self.tray_menu)
            self.tray.show()

        self.updater = Updater(self)
        self.updater.status.connect(self._on_update_status)
        self.updater.ready.connect(self._on_update_ready)
        QTimer.singleShot(8000, self.updater.check)                 # first check shortly after start
        self._update_timer = QTimer(self)
        self._update_timer.setInterval(UPDATE_CHECK_INTERVAL_MS)
        self._update_timer.timeout.connect(self.updater.check)
        self._update_timer.start()

        # global hotkeys
        self.hk = HotkeyFilter()
        QApplication.instance().installNativeEventFilter(self.hk)
        failed = []
        for name, fn in (("check_next", self.check_next), ("undo", self.undo_check),
                         ("toggle_visible", self.toggle_visible), ("reset", self.reset_all)):
            if not self.hk.register(self.hotkeys[name], fn):
                failed.append(pretty_hotkey(self.hotkeys[name]))
        if failed and sys.platform == "win32":
            QTimer.singleShot(2000, lambda: self._notify(
                "Hotkey in use by another app (or invalid): " + ", ".join(failed)
                + ". Change it in qc_checklist_state.json."))

        gui = QGuiApplication.instance()
        gui.screenAdded.connect(self.on_screens_changed)
        gui.screenRemoved.connect(self.on_screens_changed)

        self._load_items(state)
        self._update_hint()
        self.restack()
        self.overlay.show()
        for t in self._all_toasts():
            if not t.isHidden():
                t.fade_in()
        self.refresh_mask()
        self._update_status()

    # ------------------------------------------------------------------ data #
    def _all_toasts(self) -> list[Toast]:
        return self.toasts + ([self.hint_toast] if self.hint_toast else [])

    @staticmethod
    def _clean_opacity(v: Any) -> int:
        try:
            return max(MIN_OPACITY, min(100, int(v)))
        except (TypeError, ValueError):
            return DEFAULT_OPACITY

    @staticmethod
    def _on_screen(x: int, y: int) -> bool:
        return any(s.availableGeometry().contains(QPoint(x + 40, y + 20)) for s in QApplication.screens())

    def _load_items(self, state: dict[str, Any]) -> None:
        raw = state.get("items")
        entries: list[dict[str, Any]] = []
        if isinstance(raw, list):
            entries = [e for e in raw if isinstance(e, dict)
                       and isinstance(e.get("text"), str) and e["text"].strip()]
        if not isinstance(raw, list):               # first run: defaults (an empty saved list stays empty)
            entries = [{"text": t} for t in DEFAULT_ITEMS]
        for e in entries:
            pos = None
            p = e.get("pos")
            if isinstance(p, list) and len(p) == 2:
                try:
                    x, y = int(p[0]), int(p[1])
                    if self._on_screen(x, y):
                        pos = (x, y)
                except (TypeError, ValueError):
                    pass
            t = Toast(self, e["text"].strip(), bool(e.get("checked", False)), pos)
            if t.checked and e.get("hidden"):          # only ticked items may stay hidden
                t.hide()
            self.toasts.append(t)

    def save(self) -> None:
        o = self.origin
        data = {
            "version": 4,
            "opacity": self.opacity,
            "origin": [o.x(), o.y()] if o is not None else None,
            "hotkeys": self.hotkeys,
            "items": [
                {
                    "text": t.text,
                    "checked": t.checked,
                    "hidden": t.isHidden(),
                    "pos": [t.global_pos().x(), t.global_pos().y()] if t.custom_pos else None,
                }
                for t in self.toasts
            ],
        }
        save_state(self.path, data)

    def schedule_save(self) -> None:
        self._save_timer.start()

    # ---------------------------------------------------------------- layout #
    def refresh_mask(self) -> None:
        if hasattr(self, "overlay") and not self._mask_blocked:
            self.overlay.refresh_mask()

    def _origin(self) -> QPoint:
        """Top-left of the auto-stack (default: top-right of the primary screen)."""
        if self.origin is None:
            screen = QApplication.primaryScreen()
            area = screen.availableGeometry() if screen else None
            self.origin = QPoint((area.right() - TOAST_WIDTH - 20) if area else 40,
                                 (area.top() + 60) if area else 60)
        return self.origin

    def restack(self, reset_custom: bool = False, reset_origin: bool = False) -> None:
        """Stack visible, auto-positioned toasts in a column starting at the stack origin."""
        if reset_origin:
            self.origin = None
        o = self._origin()
        x, y = o.x(), o.y()
        n = 0
        for t in self._all_toasts():
            if reset_custom:
                t.custom_pos = False
            t.fit()
            if not t.hint:
                n += 1                       # numbering counts hidden items too (stable step numbers)
                t.set_index(n)
            if t.isHidden() or t.custom_pos:
                continue
            t.place(x, y)
            y += t.height() + TOAST_GAP
        self.refresh_mask()
        self.schedule_save()

    def move_all(self, dx: int, dy: int) -> None:
        """Move every toast (and the stack origin) by the same offset, keeping them on the desktop."""
        everything = self._all_toasts()
        vis = [t for t in everything if not t.isHidden()]
        if not vis:
            return
        ov = self.overlay
        left, top = min(t.x() for t in vis), min(t.y() for t in vis)
        right, bottom = max(t.x() + t.width() for t in vis), max(t.y() + t.height() for t in vis)
        dx = max(-left, min(dx, ov.width() - right))
        dy = max(-top, min(dy, ov.height() - bottom))
        if dx == 0 and dy == 0:
            return
        self._mask_blocked = True            # update the click-through mask once, not per toast
        try:
            for t in everything:
                t.move(t.x() + dx, t.y() + dy)
        finally:
            self._mask_blocked = False
        self._origin()
        self.origin = self.origin + QPoint(dx, dy)
        self.refresh_mask()
        self.schedule_save()

    def on_screens_changed(self, *_: Any) -> None:
        """Monitors plugged/unplugged: resize overlay but keep toasts where they were on screen."""
        keep = {t: t.global_pos() for t in self._all_toasts()}
        self.overlay.sync_geometry()
        for t, p in keep.items():
            t.place(p.x(), p.y())
        self.restack()

    def _update_hint(self) -> None:
        """Placeholder toast: shown when the list is empty, or when everything is ticked + hidden."""
        want = None
        if not self.toasts:
            want = "No QC items.\nRight-click here to add one."
        elif all(t.checked for t in self.toasts) and all(t.isHidden() for t in self.toasts):
            want = f"✓ All checks complete\nReset: {pretty_hotkey(self.hotkeys['reset'])} or right-click"
        if want is None:
            if self.hint_toast is not None:
                self.hint_toast.setParent(None)
                self.hint_toast.deleteLater()
                self.hint_toast = None
        elif self.hint_toast is None:
            self.hint_toast = Toast(self, want, hint=True)
            self.restack()
            if self.overlay.isVisible():
                self.hint_toast.fade_in()
        else:
            self.hint_toast.set_text(want)

    def _notify(self, msg: str) -> None:
        if self.tray:
            self.tray.showMessage(APP_TITLE, msg, QSystemTrayIcon.MessageIcon.Information, 5000)

    def _update_status(self) -> None:
        total = len(self.toasts)
        done = sum(1 for t in self.toasts if t.checked)
        if self.tray:
            self.tray.setToolTip(f"{APP_TITLE}: {done}/{total} done")

    # --------------------------------------------------------------- actions #
    def toggle(self, toast: Toast) -> None:
        toast.set_checked(not toast.checked)
        self._update_status()
        self.schedule_save()
        if self.toasts and all(t.checked for t in self.toasts):
            self._notify("All checks complete ✓")

    def add_item(self) -> None:
        text = ask_text("Add QC item", "Check item:")
        if not text:
            return
        t = Toast(self, text)
        self.toasts.append(t)
        self._update_hint()
        self.restack()
        if self.visible_all:
            t.fade_in()
        self._update_status()

    def edit_item(self, toast: Toast) -> None:
        text = ask_text("Edit QC item", "Check item:", toast.text)
        if text:
            toast.set_text(text)
            self.restack()

    def delete_item(self, toast: Toast) -> None:
        if toast in self.toasts:
            self.toasts.remove(toast)
            toast.setParent(None)        # detach immediately so it leaves the mask/overlay
            toast.deleteLater()
            self._update_hint()
            self.restack()
            self._update_status()

    def move_item(self, toast: Toast, delta: int) -> None:
        i = self.toasts.index(toast)
        j = i + delta
        if 0 <= j < len(self.toasts):
            self.toasts[i], self.toasts[j] = self.toasts[j], self.toasts[i]
            self.restack(reset_custom=True)   # order == stacking order, so re-stack everything

    def check_next(self) -> None:
        """Hotkey: tick the first unchecked item (in list order) and hide its toast."""
        nxt = next((t for t in self.toasts if not t.checked), None)
        if nxt is None:
            return
        nxt.set_checked(True)
        nxt.hide()
        self.check_history.append(nxt)
        self._after_hotkey_change()

    def undo_check(self) -> None:
        """Hotkey: bring back the most recently ticked item (unticked and visible again)."""
        target = None
        while self.check_history:
            cand = self.check_history.pop()
            if cand in self.toasts and cand.checked:
                target = cand
                break
        if target is None:                                   # fallback: last ticked item in the list
            target = next((t for t in reversed(self.toasts) if t.checked), None)
        if target is None:
            return
        target.set_checked(False)
        was_hidden = target.isHidden()
        target.show()
        self._after_hotkey_change()
        if was_hidden and self.overlay.isVisible():
            target.fade_in()

    def _after_hotkey_change(self) -> None:
        self._update_hint()
        self.restack()                       # visible toasts close the gap / re-open the slot
        self._update_status()
        if self.toasts and all(t.checked for t in self.toasts):
            self._notify("All checks complete ✓")

    def reset_all(self) -> None:
        """Untick everything and show every toast again (next arch)."""
        self.check_history.clear()
        for t in self.toasts:
            t.set_checked(False)
            if t.isHidden():
                t.show()
        self._update_hint()
        self.restack()
        self._update_status()
        self.schedule_save()

    def set_opacity(self, pct: int) -> None:
        self.opacity = self._clean_opacity(pct)
        for t in self._all_toasts():
            t.stop_anim()
            t.effect.setOpacity(self.opacity / 100)
        self.schedule_save()

    def toggle_visible(self) -> None:
        self.visible_all = not self.visible_all
        self.overlay.setVisible(self.visible_all)
        if self.visible_all:
            self.refresh_mask()

    @pyqtSlot(str)
    def _on_update_status(self, msg: str) -> None:
        if msg and self.tray:
            self.tray.showMessage(APP_TITLE, msg, QSystemTrayIcon.MessageIcon.Information, 4000)

    @pyqtSlot(str, bool)
    def _on_update_ready(self, ver: str, manual: bool) -> None:
        self.update_ready = ver
        if manual:                       # "Check for updates" clicked -> install right away
            self._notify(f"Installing v{ver} and restarting…")
            self.apply_update()
            return
        if self.tray:
            self.tray.showMessage(APP_TITLE, f"Update v{ver} downloaded. It installs on next start, "
                                  "or use the tray menu: Restart to update.",
                                  QSystemTrayIcon.MessageIcon.Information, 6000)

    def apply_update(self) -> None:
        """Quit and let the helper script swap in the downloaded exe, then restart it."""
        pending = self.updater.pending_file
        if pending is not None and launch_updater(pending):
            self.quit()
        elif self.tray:
            self.tray.showMessage(APP_TITLE, "Update can only be applied from the packaged Windows .exe.",
                                  QSystemTrayIcon.MessageIcon.Warning, 4000)

    def quit(self) -> None:
        self.save()
        self.hk.unregister_all()
        if self.tray:
            self.tray.hide()
        QApplication.quit()

    def _on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.Trigger:   # left click
            self.toggle_visible()

    # ------------------------------------------------------------------ menu #
    def build_menu(self, toast: Toast | None) -> QMenu:
        """Right-click menu of a toast."""
        m = QMenu()
        m.setStyleSheet(MENU_STYLE)
        self._fill_menu(m, toast)
        return m

    def _fill_menu(self, m: QMenu, toast: Toast | None) -> None:
        """Populate a menu; rebuilt every time it opens so its state is always fresh."""
        m.clear()
        if toast is not None and not toast.hint and toast in self.toasts:
            idx = self.toasts.index(toast)
            m.addAction("Edit…", lambda: self.edit_item(toast))
            up = m.addAction("Move up", lambda: self.move_item(toast, -1))
            up.setEnabled(idx > 0)
            down = m.addAction("Move down", lambda: self.move_item(toast, 1))
            down.setEnabled(idx < len(self.toasts) - 1)
            m.addAction("Delete", lambda: self.delete_item(toast))
            m.addSeparator()
        hk = {k: pretty_hotkey(v) for k, v in self.hotkeys.items()}
        m.addAction("➕ Add item…", self.add_item)
        m.addAction(f"✓ Check next + hide    ({hk['check_next']})", self.check_next)
        m.addAction(f"↶ Undo last check    ({hk['undo']})", self.undo_check)
        m.addAction(f"↺ Reset all (untick + show)    ({hk['reset']})", self.reset_all)
        m.addAction(f"Show / hide toasts    ({hk['toggle_visible']})", self.toggle_visible)
        m.addAction("Re-stack toasts (default place)", lambda: self.restack(reset_custom=True, reset_origin=True))
        m.addAction("Tip: Shift + drag = move all toasts").setEnabled(False)

        sub = m.addMenu("Opacity")
        group = QActionGroup(sub)
        group.setExclusive(True)
        for pct in (100, 85, 70, 55, 40):
            a = QAction(f"{pct}%", sub, checkable=True)
            a.setChecked(pct == self.opacity)
            a.triggered.connect(lambda _=False, p=pct: self.set_opacity(p))
            group.addAction(a)
            sub.addAction(a)

        m.addSeparator()
        if self.update_ready:
            m.addAction(f"⟳ Restart to update to v{self.update_ready}", self.apply_update)
        else:
            m.addAction("Check for updates", lambda: self.updater.check(manual=True))
        m.addAction(f"Version {APP_VERSION}").setEnabled(False)
        m.addAction("Quit", self.quit)


# --------------------------------------------------------------------------- #
TOAST_STYLE = """
QWidget { font-family: "Segoe UI", "Noto Sans", "DejaVu Sans", sans-serif; }
QFrame { background: transparent; border: none; }
#card { background-color: #1e2127; border: 1px solid #3a3f4b; border-left: 4px solid #4a90d9; border-radius: 8px; }
#card[done="true"] { border-left-color: #34a853; background-color: #1b2420; }
#num { color: #6b7280; font-size: 11px; font-weight: 600; background: transparent; border: none; }
#text { color: #f1f3f4; font-size: 14px; background: transparent; border: none; }
#text[done="true"] { color: #8d9399; }
#box { border: 1.5px solid #6b7280; border-radius: 5px; background: #14161a; color: white; font-size: 14px; font-weight: bold; }
#box[done="true"] { background: #34a853; border-color: #34a853; }
"""

MENU_STYLE = """
QMenu { background: #2a2f38; color: #e8eaed; border: 1px solid #4a90d9; padding: 4px; }
QMenu::item { padding: 5px 22px; }
QMenu::item:selected { background: #2f4f7a; }
QMenu::item:disabled { color: #6b7280; }
"""


def main() -> int:
    if sys.platform == "win32" and is_frozen():       # finish a downloaded update before showing UI
        pending = find_pending()
        if pending is not None and launch_updater(pending):
            return 0
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setQuitOnLastWindowClosed(False)   # overlay/tray keep the app alive
    try:
        manager = Manager()
    except Exception as exc:
        print(f"Failed to start: {exc}", file=sys.stderr)
        return 1
    app._manager = manager  # keep a reference so it isn't garbage-collected
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())

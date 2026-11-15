#!/usr/bin/env python3
"""
QC Toast Checklist for 3D dental-arch reviews
=============================================
Semi-transparent, always-on-top "toasts" floating over every other application
(e.g. your 3D viewer). One toast per QC check, each with a Pass (✓) and Fail (✗) button,
plus a header toast showing the OPERATOR and the CASE you are checking.

The 8 checks: case version number, mismatch, B grade note, attachment processing,
tooth numbering, axis, arch curve, extra material and undercut.

Workflow
--------
1. Create an operator (menu -> Operator -> New operator), then paste that operator's case numbers.
2. Judge the 8 checks with the hotkeys (or click the ✓ / ✗ buttons on the toasts).
3. When all 8 are judged, the case is saved as one row in your GOOGLE SHEET
   (date, time, operator, case, each check, passed/failed, PASS/FAIL). Set it up once:
   menu -> "Connect Google Sheet..." (see google_apps_script/Code.gs). Without it the
   rows go to a local qc_results.xlsx instead.
4. Press the Next-case hotkey: the operator's next unfinished case loads.
Each operator has their own cases and their own Pass/Fail status for the same checklist.

Mouse
-----
* Click ✓ / ✗ on a toast ... Pass / Fail that item (click again to clear)
* Click the header toast .. menu (operators, cases, ...)    * Right-click any toast ... same menu
* Drag a toast ............ move it          * Shift + drag ...... move ALL toasts together
* Mouse wheel on a toast .. opacity of all toasts

Global hotkeys (Windows; work while another app has focus; editable in qc_checklist_state.json)
* Ctrl+Alt+Space .. PASS the next item   * Ctrl+Alt+X .. FAIL the next item   * Ctrl+Alt+Z .. undo
* Ctrl+Alt+N ..... next case             * Ctrl+Alt+H .. show/hide toasts     * Ctrl+Alt+R .. reset case
(The next item to be judged has a yellow outline.)

Files next to the exe: qc_checklist_state.json (operators, cases, statuses, Google connection) and,
only when no Google Sheet is connected, qc_results.xlsx. If the sheet can't be reached (offline,
file open in Excel) rows are kept and re-sent automatically, even after a restart.

Auto-update (Windows .exe built by CI, see .github/workflows/build.yml)
-----------------------------------------------------------------------
On start (and every 6 h) the app asks GitHub Releases for a newer version, downloads
QCChecklist.exe in the background, verifies its SHA-256, and applies it the next time
the app starts. "Check for updates" in the menu installs a newer version immediately
(download -> restart). Background checks never interrupt you; they apply on next start.
Set UPDATE_REPO below to "your-user/your-repo".

Requires: Python 3.10+, PyQt6, openpyxl  ->  pip install -r requirements.txt
"""

from __future__ import annotations

import csv
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
from pathlib import Path
from datetime import datetime
from typing import Any

from PyQt6.QtCore import QAbstractNativeEventFilter, QEasingCurve, QObject, QPoint, QPropertyAnimation, QRect, Qt, QTimer, QUrl, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QAction, QActionGroup, QColor, QDesktopServices, QGuiApplication, QIcon, QPainter, QPainterPath, QPen, QPixmap, QRegion
from PyQt6.QtWidgets import (
    QApplication,
    QFrame,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMenu,
    QMessageBox,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

APP_TITLE = "QC Toast Checklist"
STATE_FILENAME = "qc_checklist_state.json"
DEFAULT_CHECKLIST = [
    "Case version number",
    "Mismatch",
    "B grade note",
    "Attachment processing",
    "Tooth numbering",
    "Axis",
    "Arch curve",
    "Extra material and undercut",
]
DEFAULT_OPACITY = 85
MIN_OPACITY = 30
TOAST_WIDTH = 340
BTN_SIZE = 26       # Pass / Fail button size (px)
TOAST_GAP = 8
DRAG_THRESHOLD = 6  # px of movement before a click becomes a drag

# Global hotkeys (Windows). Format: modifiers+key, e.g. "ctrl+alt+space", "ctrl+shift+f9".
# Override any of them in qc_checklist_state.json under "hotkeys".
DEFAULT_HOTKEYS = {
    "pass": "ctrl+alt+space",         # PASS the next unjudged item
    "fail": "ctrl+alt+x",             # FAIL the next unjudged item
    "undo": "ctrl+alt+z",             # clear the last judgment
    "next_case": "ctrl+alt+n",        # load the operator's next unfinished case
    "toggle_visible": "ctrl+alt+h",   # show / hide all toasts
    "reset": "ctrl+alt+r",            # clear all marks of the current case
}

# --- auto-update settings ----------------------------------------------------
APP_VERSION = "0.0.0"                         # stamped automatically by the CI build
UPDATE_REPO = "YOUR_GITHUB_USER/YOUR_REPO"    # <-- set this (public GitHub repo)
ASSET_NAME = "QCChecklist.exe"                # release asset names (must match build.yml)
API_URL = "https://api.github.com/repos/{repo}/releases/latest"
REQUIRE_HTTPS = True
UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000  # re-check every 6 hours

# --- Google Sheet -----------------------------------------------------------
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/159GgHct3sJ-hU3QAlL2hNHqwoB1Js6H3BoIp41J9c7E/edit"
GOOGLE_WEBAPP_PREFIX = "https://script.google.com/macros/s/"   # Apps Script web-app URLs start like this


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


MODAL = [0]          # > 0 while one of our dialogs is open (global hotkeys are ignored meanwhile)
TOPMOST_PAUSE = [0]  # > 0 while a menu is open: the overlay must not re-raise itself above it


def make_menu_topmost(menu: "QMenu") -> None:
    """Menus must be always-on-top too, otherwise the topmost toast overlay hides them."""
    menu.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    for a in menu.actions():
        if a.menu() is not None:
            make_menu_topmost(a.menu())


def _run_dialog(dlg) -> bool:
    MODAL[0] += 1
    try:
        return bool(dlg.exec())
    finally:
        MODAL[0] -= 1


def ask_text(title: str, label: str, text: str = "") -> str | None:
    """Small always-on-top text prompt (so it isn't hidden behind the 3D viewer)."""
    dlg = QInputDialog()
    dlg.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dlg.setWindowTitle(title)
    dlg.setLabelText(label)
    dlg.setTextValue(text)
    dlg.setMinimumWidth(360)
    if _run_dialog(dlg) and dlg.textValue().strip():
        return dlg.textValue().strip()
    return None


def ask_multiline(title: str, label: str) -> str | None:
    """Multi-line prompt (used to paste a list of case numbers)."""
    dlg = QInputDialog()
    dlg.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dlg.setOption(QInputDialog.InputDialogOption.UsePlainTextEditForTextInput, True)
    dlg.setWindowTitle(title)
    dlg.setLabelText(label)
    dlg.resize(420, 320)
    if _run_dialog(dlg) and dlg.textValue().strip():
        return dlg.textValue().strip()
    return None


def confirm(title: str, text: str) -> bool:
    """Always-on-top Yes/No question (default: No)."""
    box = QMessageBox()
    box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    box.setIcon(QMessageBox.Icon.Question)
    box.setWindowTitle(title)
    box.setText(text)
    yes = box.addButton("Yes", QMessageBox.ButtonRole.YesRole)
    no = box.addButton("No", QMessageBox.ButtonRole.NoRole)
    box.setDefaultButton(no)
    _run_dialog(box)
    return box.clickedButton() is yes


def info(title: str, text: str) -> None:
    """Always-on-top OK message."""
    box = QMessageBox()
    box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    box.setIcon(QMessageBox.Icon.Information)
    box.setWindowTitle(title)
    box.setText(text)
    box.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
    _run_dialog(box)


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
        if not self.isVisible() or MODAL[0] or TOPMOST_PAUSE[0]:
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
# One toast (a child widget of the overlay): the header toast or a checklist-item toast
# --------------------------------------------------------------------------- #
class Toast(QFrame):
    """kind == "header": operator + case + progress (click it for the menu).
    kind == "item":   one checklist item with a Pass (✓) and a Fail (✗) button."""

    def __init__(self, mgr: "Manager", kind: str, key: Any, text: str = "", index: int = 0,
                 pos: tuple[int, int] | None = None) -> None:
        super().__init__(mgr.overlay)
        self.mgr = mgr
        self.kind = kind
        self.key = key
        self.status: str | None = None        # None | "pass" | "fail"
        self.is_current = False               # the item the Pass/Fail hotkeys will judge next
        self.custom_pos = pos is not None     # True once the user placed it manually
        self._user_dragging = False
        self._press: QPoint | None = None
        self._press_local = QPoint()
        self._offset = QPoint()
        self._moved = False
        self._move_all = False                # Shift+drag moves every toast
        self._last = QPoint()
        self._anim: QPropertyAnimation | None = None
        self._base_cursor = (Qt.CursorShape.PointingHandCursor if kind == "header"
                             else Qt.CursorShape.ArrowCursor)
        self.setCursor(self._base_cursor)

        # Qt-rendered transparency (reliable on every platform)
        self.effect = QGraphicsOpacityEffect(self)
        self.effect.setOpacity(mgr.opacity / 100)
        self.setGraphicsEffect(self.effect)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.card = QFrame()
        self.card.setObjectName("card")
        self.card.setProperty("kind", kind)
        self.card.setProperty("state", "none")
        self.card.setProperty("current", "false")
        outer.addWidget(self.card)

        row = QHBoxLayout(self.card)
        row.setContentsMargins(10, 8, 12, 8)
        row.setSpacing(10)

        self.label = QLabel(text)
        self.label.setObjectName("text")
        self.label.setWordWrap(True)
        self.label.setTextFormat(Qt.TextFormat.RichText if kind == "header" else Qt.TextFormat.PlainText)
        self.num = self.pbtn = self.fbtn = None
        widgets = [self.card, self.label]

        if kind == "item":
            self.num = QLabel(str(index))
            self.num.setObjectName("num")
            self.num.setFixedWidth(18)
            self.pbtn, self.fbtn = QLabel("✓"), QLabel("✗")
            for b, name in ((self.pbtn, "pbtn"), (self.fbtn, "fbtn")):
                b.setObjectName(name)
                b.setFixedSize(BTN_SIZE, BTN_SIZE)
                b.setAlignment(Qt.AlignmentFlag.AlignCenter)
                b.setProperty("active", "false")
            row.addWidget(self.num)
            row.addWidget(self.label, 1)
            btns = QHBoxLayout()
            btns.setSpacing(4)
            btns.addWidget(self.pbtn)
            btns.addWidget(self.fbtn)
            row.addLayout(btns)
            widgets += [self.num, self.pbtn, self.fbtn]
        else:
            row.addWidget(self.label, 1)

        # Mouse events fall through to the toast so click/drag work anywhere on it.
        for w in widgets:
            w.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.setStyleSheet(TOAST_STYLE)
        self.fit()
        if pos is not None:
            self.place(*pos)

    # ------------------------------------------------------------ appearance #
    @property
    def text(self) -> str:
        return self.label.text()

    def set_status(self, status: str | None) -> None:
        self.status = status
        self._refresh_props()

    def set_current(self, flag: bool) -> None:
        self.is_current = flag
        self._refresh_props()

    def set_info(self, html_text: str, state: str) -> None:
        """Header toast: rich text + colour state (none / progress / pass / fail)."""
        self.label.setText(html_text)
        self.card.setProperty("state", state)
        repolish(self.card)
        self.fit()

    def _refresh_props(self) -> None:
        if self.kind != "item":
            return
        st = self.status or "none"
        self.card.setProperty("state", st)
        self.card.setProperty("current", "true" if (self.is_current and self.status is None) else "false")
        self.pbtn.setProperty("active", "true" if st == "pass" else "false")
        self.fbtn.setProperty("active", "true" if st == "fail" else "false")
        for w in (self.card, self.pbtn, self.fbtn):
            repolish(w)

    def fit(self) -> None:
        """Fixed width; height follows the word-wrapped text (computed explicitly for reliability)."""
        avail = TOAST_WIDTH - 10 - 12 - 6                  # card margins + border
        min_h = 0
        if self.kind == "item":
            avail -= 18 + (2 * BTN_SIZE + 4) + 2 * 10      # number + buttons + spacings
            min_h = BTN_SIZE
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
            self._press_local = e.position().toPoint()
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
            if not self._moved:
                self._clicked(self._press_local, e.globalPosition().toPoint())
            else:
                if not self._move_all:
                    self.custom_pos = True
                self.mgr.schedule_save()
        self._press = None
        self._moved = False
        self._user_dragging = False
        if self._move_all:
            self._move_all = False
            self.setCursor(self._base_cursor)

    def _clicked(self, local: QPoint, global_pt: QPoint) -> None:
        if self.kind == "header":
            self.mgr.popup_menu(self, global_pt)
            return
        pad = 5   # forgiving hit area around the small buttons
        if self.pbtn.geometry().adjusted(-pad, -pad, pad, pad).contains(local):
            self.mgr.toggle_status(self.key, "pass")
        elif self.fbtn.geometry().adjusted(-pad, -pad, pad, pad).contains(local):
            self.mgr.toggle_status(self.key, "fail")

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
        self.mgr.popup_menu(self, e.globalPos())


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
# Results sheet (Excel .xlsx; falls back to .csv if openpyxl is not installed)
# --------------------------------------------------------------------------- #
def today_str() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _norm(v: Any) -> str:
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return "" if v is None else str(v).strip()


class ResultsSheet:
    """One row per (date, operator, case). Re-saving the same case updates its row."""
    FIXED = ["Date", "Time", "Operator", "Case number"]
    TAIL = ["Passed", "Failed", "Result"]
    FILLS = {"Pass": "C6EFCE", "PASS": "C6EFCE", "Fail": "FFC7CE", "FAIL": "FFC7CE", "INCOMPLETE": "FFEB9C"}

    def __init__(self, directory: Path, item_texts: list[str]) -> None:
        self.headers = self.make_headers(item_texts)
        try:
            import openpyxl  # noqa: F401
            self.use_xlsx = True
        except ImportError:
            self.use_xlsx = False
        self.path = directory / ("qc_results.xlsx" if self.use_xlsx else "qc_results.csv")
        self.label = self.path.name
        self.is_google = False

    @classmethod
    def make_headers(cls, item_texts: list[str]) -> list[str]:
        return cls.FIXED + item_texts + cls.TAIL

    @staticmethod
    def key(row: list[Any]) -> tuple[str, str, str]:
        return (_norm(row[0]), _norm(row[2]), _norm(row[3]))

    def _rotate(self) -> None:
        """Checklist changed -> keep the old file and start a fresh sheet."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            self.path.replace(self.path.with_name(f"{self.path.stem}_old_{stamp}{self.path.suffix}"))
        except OSError:
            pass

    def upsert(self, rows: list[list[Any]]) -> None:
        (self._upsert_xlsx if self.use_xlsx else self._upsert_csv)(rows)

    def _upsert_xlsx(self, rows: list[list[Any]]) -> None:
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Alignment, Font, PatternFill

        wb = None
        if self.path.exists():
            wb = load_workbook(self.path)
            if [c.value for c in wb.active[1]] != self.headers:
                wb.close()
                self._rotate()
                wb = None
        if wb is None:
            wb = Workbook()
            ws = wb.active
            ws.title = "QC Results"
            ws.append(self.headers)
            for c in ws[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2"
            for i in range(1, len(self.headers) + 1):
                ws.column_dimensions[ws.cell(1, i).column_letter].width = 12 if i <= 2 else 18
        ws = wb.active
        index = {self.key([ws.cell(r, c).value for c in (1, 2, 3, 4)]): r for r in range(2, ws.max_row + 1)}
        for row in rows:
            r = index.get(self.key(row))
            if r is None:
                ws.append(row)
                r = ws.max_row
                index[self.key(row)] = r
            else:
                for c, v in enumerate(row, start=1):
                    ws.cell(r, c).value = v
            for c in range(5, len(self.headers) + 1):
                cell = ws.cell(r, c)
                color = self.FILLS.get(str(cell.value))
                cell.fill = PatternFill("solid", fgColor=color) if color else PatternFill()
        tmp = self.path.with_name(self.path.stem + ".tmp" + self.path.suffix)
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)          # raises PermissionError while Excel has the file open
        finally:
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass

    def _upsert_csv(self, rows: list[list[Any]]) -> None:
        data: list[list[str]] = []
        if self.path.exists():
            with self.path.open("r", encoding="utf-8-sig", newline="") as fh:
                data = list(csv.reader(fh))
            if not data or data[0] != self.headers:
                self._rotate()
                data = []
        if not data:
            data = [list(self.headers)]
        index = {self.key(r): i for i, r in enumerate(data) if i > 0 and len(r) >= 4}
        for row in rows:
            row = ["" if v is None else str(v) for v in row]
            i = index.get(self.key(row))
            if i is None:
                data.append(row)
                index[self.key(row)] = len(data) - 1
            else:
                data[i] = row
        tmp = self.path.with_name(self.path.stem + ".tmp" + self.path.suffix)
        try:
            with tmp.open("w", encoding="utf-8-sig", newline="") as fh:   # BOM so Excel reads UTF-8
                csv.writer(fh).writerows(data)
            os.replace(tmp, self.path)
        finally:
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass


class GoogleSheet:
    """Sends rows to an Apps Script web app that upserts them into the Google Sheet."""
    label = "Google Sheet"
    is_google = True

    def __init__(self, url: str, token: str, headers: list[str]) -> None:
        self.url, self.token, self.headers = url.strip(), token.strip(), headers

    @staticmethod
    def valid_url(url: str) -> bool:
        return (not REQUIRE_HTTPS) or (url.strip().startswith(GOOGLE_WEBAPP_PREFIX) and url.strip().endswith("/exec"))

    def _post(self, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.valid_url(self.url):
            raise RuntimeError("not an Apps Script web-app URL (must end with /exec)")
        body = json.dumps({"token": self.token, **payload}, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(self.url, data=body, method="POST",
                                     headers={"Content-Type": "text/plain; charset=utf-8"})
        # Apps Script answers a POST with a redirect; urllib follows it with a GET and gets the JSON reply.
        with urllib.request.urlopen(req, timeout=30) as r:
            text = r.read().decode("utf-8", "replace")
        try:
            res = json.loads(text)
        except ValueError:
            raise RuntimeError("unexpected reply (is the web app deployed with access 'Anyone'?)") from None
        if not res.get("ok"):
            raise RuntimeError(str(res.get("error") or "rejected by the script"))
        return res

    def ping(self) -> None:
        self._post({"ping": True})

    def upsert(self, rows: list[list[Any]]) -> None:
        self._post({"headers": self.headers, "rows": rows})


# --------------------------------------------------------------------------- #
# Manager: owns the overlay, toasts, operators/cases, tray icon, sheet and persistence
# --------------------------------------------------------------------------- #
class Manager(QObject):
    sheet_result = pyqtSignal(object, str)      # (rows, error) from the background upload thread

    def __init__(self) -> None:
        super().__init__()
        self.path = state_file_path()
        state = load_state(self.path)
        self.opacity = self._clean_opacity(state.get("opacity", DEFAULT_OPACITY))
        self.hide_judged = bool(state.get("hide_judged", False))
        self.overlay = Overlay()
        self.visible_all = True
        self.update_ready: str | None = None
        self.history: list[int] = []                 # item ids judged in this case (for undo)
        self.pending_rows: dict[tuple[str, str, str], list[Any]] = {}
        self._last_rows: dict[tuple[str, str, str], tuple] = {}
        self._last_sheet_warn = 0.0
        self._mask_blocked = False
        self.origin: QPoint | None = None            # where the auto-stack starts (moved by Shift+drag)
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

        self.checklist = self._load_checklist(state)
        self.operators = self._load_operators(state)
        cur = state.get("current") if isinstance(state.get("current"), dict) else {}
        self.cur_op: str | None = cur.get("operator") if cur.get("operator") in self.operators \
            else next(iter(self.operators), None)
        self.cur_case: str | None = None
        if self.cur_op:
            cases = self.operators[self.cur_op]["cases"]
            want = cur.get("case") if cur.get("operator") == self.cur_op else None
            self.cur_case = want if want in cases else self.operators[self.cur_op].get("last_case")
        gcfg = state.get("google_sheet") if isinstance(state.get("google_sheet"), dict) else {}
        self.sheet_cfg: dict[str, str] = {k: str(gcfg.get(k) or "").strip() for k in ("url", "token")}
        self.sheet = self._make_sheet()
        self._flushing = False
        self._flush_thread: threading.Thread | None = None
        self.sheet_result.connect(self._on_sheet_result)
        # rows that could not be sent last time (offline / file locked) are retried
        n_cols = len(ResultsSheet.make_headers([c["text"] for c in self.checklist]))
        for r in state.get("pending_rows") or []:
            if isinstance(r, list) and len(r) == n_cols:
                self.pending_rows[ResultsSheet.key(r)] = r

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(400)
        self._save_timer.timeout.connect(self.save)

        self._top_timer = QTimer(self)               # keep the overlay above everything
        self._top_timer.setInterval(1000)
        self._top_timer.timeout.connect(self.overlay.keep_on_top)
        self._top_timer.start()

        self._sheet_timer = QTimer(self)             # retry if the sheet was locked (open in Excel)
        self._sheet_timer.setInterval(15000)
        self._sheet_timer.timeout.connect(self.flush_sheet)
        self._sheet_timer.start()

        self.tray: QSystemTrayIcon | None = None
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray = QSystemTrayIcon(make_tray_icon())
            self.tray.activated.connect(self._on_tray_activated)
            self.tray_menu = QMenu()
            self.tray_menu.setStyleSheet(MENU_STYLE)
            self.tray_menu.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
            self._tray_open = False
            self.tray_menu.aboutToShow.connect(self._tray_menu_opening)
            self.tray_menu.aboutToHide.connect(self._tray_menu_closing)
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

        # global hotkeys (ignored while a dialog is open)
        self.hk = HotkeyFilter()
        QApplication.instance().installNativeEventFilter(self.hk)
        bindings = (
            ("pass", lambda: self.hotkey_judge("pass")),
            ("fail", lambda: self.hotkey_judge("fail")),
            ("undo", self.undo),
            ("next_case", self.next_case),
            ("toggle_visible", self.toggle_visible),
            ("reset", self.reset_case),
        )
        failed = []
        for name, fn in bindings:
            guarded = (lambda f=fn: None if MODAL[0] else f())
            if not self.hk.register(self.hotkeys[name], guarded):
                failed.append(pretty_hotkey(self.hotkeys[name]))
        if failed and sys.platform == "win32":
            QTimer.singleShot(2000, lambda: self._notify(
                "Hotkey in use by another app (or invalid): " + ", ".join(failed)
                + ". Change it in qc_checklist_state.json."))

        gui = QGuiApplication.instance()
        gui.screenAdded.connect(self.on_screens_changed)
        gui.screenRemoved.connect(self.on_screens_changed)

        positions = state.get("positions") if isinstance(state.get("positions"), dict) else {}
        self.header = Toast(self, "header", "header", pos=self._pos_from(positions, "header"))
        self.toasts = [
            Toast(self, "item", c["id"], c["text"], i + 1, self._pos_from(positions, str(c["id"])))
            for i, c in enumerate(self.checklist)
        ]
        self.refresh_all()
        self.overlay.show()
        for t in self._all_toasts():
            if not t.isHidden():
                t.fade_in()
        self.refresh_mask()
        QTimer.singleShot(700, self._startup_prompt)
        QTimer.singleShot(3000, self.flush_sheet)

    # ------------------------------------------------------------ loading #
    @staticmethod
    def _clean_opacity(v: Any) -> int:
        try:
            return max(MIN_OPACITY, min(100, int(v)))
        except (TypeError, ValueError):
            return DEFAULT_OPACITY

    @staticmethod
    def _on_screen(x: int, y: int) -> bool:
        return any(s.availableGeometry().contains(QPoint(x + 40, y + 20)) for s in QApplication.screens())

    def _pos_from(self, positions: dict, key: str) -> tuple[int, int] | None:
        p = positions.get(key)
        if isinstance(p, list) and len(p) == 2:
            try:
                x, y = int(p[0]), int(p[1])
                if self._on_screen(x, y):
                    return (x, y)
            except (TypeError, ValueError):
                pass
        return None

    @staticmethod
    def _load_checklist(state: dict[str, Any]) -> list[dict[str, Any]]:
        """Checklist = [{"id": n, "text": "..."}]. Plain strings in the JSON list are accepted too."""
        items: list[dict[str, Any]] = []
        raw = state.get("checklist")
        if isinstance(raw, list):
            used = {int(e["id"]) for e in raw if isinstance(e, dict) and str(e.get("id", "")).isdigit()}
            nxt = max(used, default=0) + 1
            for e in raw:
                if isinstance(e, dict) and str(e.get("text", "")).strip():
                    try:
                        cid = int(e.get("id"))
                    except (TypeError, ValueError):
                        cid, nxt = nxt, nxt + 1
                    if all(i["id"] != cid for i in items):
                        items.append({"id": cid, "text": str(e["text"]).strip()})
                elif isinstance(e, str) and e.strip():
                    items.append({"id": nxt, "text": e.strip()})
                    nxt += 1
        if not items:
            items = [{"id": i + 1, "text": t} for i, t in enumerate(DEFAULT_CHECKLIST)]
        return items

    @staticmethod
    def _load_operators(state: dict[str, Any]) -> dict[str, dict[str, Any]]:
        ops: dict[str, dict[str, Any]] = {}
        raw = state.get("operators")
        if not isinstance(raw, dict):
            return ops
        for name, od in raw.items():
            if not isinstance(od, dict) or not str(name).strip():
                continue
            cases: dict[str, dict[str, Any]] = {}
            rc = od.get("cases")
            if isinstance(rc, dict):
                for cn, cd in rc.items():
                    if not isinstance(cd, dict):
                        continue
                    st = cd.get("status") if isinstance(cd.get("status"), dict) else {}
                    cases[str(cn)] = {
                        "date": str(cd.get("date") or ""),
                        "status": {str(k): v for k, v in st.items() if v in ("pass", "fail")},
                        "updated": str(cd.get("updated") or ""),
                        "in_sheet": bool(cd.get("in_sheet", False)),
                    }
            last = od.get("last_case")
            ops[str(name)] = {"cases": cases, "last_case": last if last in cases else None}
        return ops

    def save(self) -> None:
        o = self.origin
        positions = {str(t.key): [t.global_pos().x(), t.global_pos().y()]
                     for t in self._all_toasts() if t.custom_pos}
        data = {
            "version": 5,
            "opacity": self.opacity,
            "hide_judged": self.hide_judged,
            "origin": [o.x(), o.y()] if o is not None else None,
            "hotkeys": self.hotkeys,
            "checklist": self.checklist,
            "operators": self.operators,
            "current": {"operator": self.cur_op, "case": self.cur_case},
            "positions": positions,
            "google_sheet": self.sheet_cfg,
            "pending_rows": list(self.pending_rows.values()),
        }
        save_state(self.path, data)

    def schedule_save(self) -> None:
        self._save_timer.start()

    def _all_toasts(self) -> list[Toast]:
        return [self.header] + self.toasts

    # ----------------------------------------------------------- case data #
    def current_case(self) -> dict[str, Any] | None:
        if self.cur_op and self.cur_case:
            return self.operators.get(self.cur_op, {}).get("cases", {}).get(self.cur_case)
        return None

    def case_counts(self, case: dict[str, Any]) -> tuple[int, int, int, int]:
        """(judged, passed, failed, total) for the current checklist."""
        ids = {str(c["id"]) for c in self.checklist}
        st = case["status"]
        passed = sum(1 for i in ids if st.get(i) == "pass")
        failed = sum(1 for i in ids if st.get(i) == "fail")
        return passed + failed, passed, failed, len(ids)

    def _is_complete(self, case: dict[str, Any]) -> bool:
        j, _, _, total = self.case_counts(case)
        return j == total

    def first_unjudged(self) -> Toast | None:
        case = self.current_case()
        if case is None:
            return None
        return next((t for t in self.toasts if str(t.key) not in case["status"]), None)

    def _case_icon(self, case: dict[str, Any]) -> str:
        j, _, f, total = self.case_counts(case)
        if j == total:
            return "✗" if f else "✓"
        return "◐" if j else "○"

    # ---------------------------------------------------------- refreshing #
    def _header_info(self) -> tuple[str, str]:
        dim = "<span style='color:#9aa0a6;font-size:11px'>"
        if self.cur_op is None:
            return "none", "<b>No operator yet</b><br>" + dim + "Click here to create one</span>"
        op = html.escape(self.cur_op)
        top = f"{dim}OPERATOR</span>&nbsp;&nbsp;<b>{op}</b>"
        case = self.current_case()
        if case is None:
            return "none", top + "<br>" + dim + "No case yet. Click here to add cases</span>"
        j, _, f, total = self.case_counts(case)
        line = f"{dim}CASE</span>&nbsp;&nbsp;<b style='color:#8ab4f8'>{html.escape(self.cur_case or '')}</b>"
        line += f"&nbsp;&nbsp;·&nbsp;&nbsp;{j}/{total}"
        if f:
            line += f"&nbsp;&nbsp;·&nbsp;&nbsp;<span style='color:#f28b82'>{f} failed</span>"
        state = "progress"
        if j == total:
            state = "fail" if f else "pass"
            line += ("&nbsp;&nbsp;·&nbsp;&nbsp;<b style='color:#f28b82'>FAIL ✗</b>" if f
                     else "&nbsp;&nbsp;·&nbsp;&nbsp;<b style='color:#81c995'>PASS ✓</b>")
        return state, top + "<br>" + line

    def refresh_all(self) -> None:
        """Re-sync every toast with the current operator/case, then re-stack."""
        case = self.current_case()
        first = self.first_unjudged()
        for t in self.toasts:
            st = case["status"].get(str(t.key)) if case else None
            t.set_status(st)
            t.set_current(case is not None and t is first)
            hide = self.hide_judged and st is not None
            if hide != t.isHidden():
                t.setVisible(not hide)
        state, text = self._header_info()
        self.header.set_info(text, state)
        if self.tray:
            if case:
                j, _, _, total = self.case_counts(case)
                self.tray.setToolTip(f"{APP_TITLE}: {self.cur_op} / {self.cur_case}  {j}/{total}")
            else:
                self.tray.setToolTip(APP_TITLE)
        self.restack()

    def _notify(self, msg: str) -> None:
        if self.tray:
            self.tray.showMessage(APP_TITLE, msg, QSystemTrayIcon.MessageIcon.Information, 4000)
        else:
            print(f"[info] {msg}")

    # ---------------------------------------------------------- judgments #
    def _ensure_case(self) -> dict[str, Any] | None:
        """Make sure there is a current operator + case (asks the user if not)."""
        if self.cur_op is None:
            self.new_operator()
        elif self.cur_case is None:
            self.add_cases()
        return self.current_case()

    def set_status(self, item_id: Any, value: str | None, record: bool = True) -> None:
        case = self.current_case() or self._ensure_case()
        if case is None:
            return
        was_complete = self._is_complete(case)
        key = str(item_id)
        if value is None:
            case["status"].pop(key, None)
        else:
            case["status"][key] = value
            if not case.get("date"):
                case["date"] = today_str()
            if record:
                self.history.append(int(item_id))
        case["updated"] = now_iso()
        j, _, f, total = self.case_counts(case)
        if j == total or case.get("in_sheet"):
            self._queue_row(self.cur_op, self.cur_case, case)
        self.refresh_all()
        self.schedule_save()
        if j == total and not was_complete:
            self._notify(f"Case {self.cur_case}: " + (f"FAIL ✗ ({f} failed)" if f else "PASS ✓")
                         + f" · saving to {self.sheet.label}…")

    def toggle_status(self, item_id: Any, value: str) -> None:
        """Clicking the active button clears it; clicking the other one switches."""
        case = self.current_case() or self._ensure_case()
        if case is None:
            return
        cur = case["status"].get(str(item_id))
        self.set_status(item_id, None if cur == value else value)

    def hotkey_judge(self, value: str) -> None:
        """Hotkey: Pass/Fail the next unjudged item."""
        if (self.current_case() or self._ensure_case()) is None:
            return
        nxt = self.first_unjudged()
        if nxt is None:
            self._notify("All items judged. Press " + pretty_hotkey(self.hotkeys["next_case"])
                         + " for the next case.")
            return
        self.set_status(nxt.key, value)

    def undo(self) -> None:
        """Hotkey: clear the most recent judgment of this case."""
        case = self.current_case()
        if case is None:
            return
        target = None
        while self.history:
            i = self.history.pop()
            if str(i) in case["status"]:
                target = i
                break
        if target is None:                                   # fallback: last judged item in the list
            target = next((t.key for t in reversed(self.toasts) if str(t.key) in case["status"]), None)
        if target is not None:
            self.set_status(target, None, record=False)

    def reset_case(self) -> None:
        """Clear every judgment of the current case (asks first if something was judged)."""
        case = self.current_case()
        if case is None:
            return
        if case["status"] and not confirm("Reset case", f"Clear all Pass/Fail marks of case {self.cur_case}?"):
            return
        case["status"].clear()
        case["updated"] = now_iso()
        self.history.clear()
        if case.get("in_sheet"):
            self._queue_row(self.cur_op, self.cur_case, case)
        self.refresh_all()
        self.schedule_save()

    # ----------------------------------------------- operators and cases #
    def _leave_case(self) -> None:
        """Before switching away / quitting: store a partial case in the sheet as INCOMPLETE."""
        case = self.current_case()
        if case is not None and (self.case_counts(case)[0] or case.get("in_sheet")):
            self._queue_row(self.cur_op, self.cur_case, case)

    def _startup_prompt(self) -> None:
        if not self.operators:
            self.new_operator()

    def new_operator(self) -> None:
        name = ask_text("New operator", "Operator name:")
        if not name:
            return
        existing = next((n for n in self.operators if n.lower() == name.lower()), None)
        if existing is None:
            self.operators[name] = {"cases": {}, "last_case": None}
            existing = name
        self.switch_operator(existing)
        if not self.operators[existing]["cases"]:
            self.add_cases()

    def switch_operator(self, name: str) -> None:
        if name not in self.operators:
            return
        self._leave_case()
        self.cur_op = name
        od = self.operators[name]
        pick = od.get("last_case")
        if pick not in od["cases"]:
            pick = next((n for n, c in od["cases"].items() if not self._is_complete(c)),
                        next(iter(od["cases"]), None))
        self.cur_case = pick
        self.history.clear()
        self.refresh_all()
        self.schedule_save()

    def set_case(self, case_no: str) -> None:
        if self.cur_op is None or case_no not in self.operators[self.cur_op]["cases"]:
            return
        self._leave_case()
        self.cur_case = case_no
        self.operators[self.cur_op]["last_case"] = case_no
        self.history.clear()
        self.refresh_all()
        self.schedule_save()

    def add_cases(self) -> None:
        """Add one or several case numbers to the current operator (paste a whole list if you like)."""
        if self.cur_op is None:
            self.new_operator()
            return
        text = ask_multiline("Add cases", f"Case numbers for {self.cur_op}\n"
                             "(one per line, or separated by commas / spaces):")
        nums = list(dict.fromkeys(n for n in re.split(r"[,;\s]+", text or "") if n))
        if not nums:
            return
        cases = self.operators[self.cur_op]["cases"]
        new = [n for n in nums if n not in cases]
        for n in new:
            cases[n] = {"date": "", "status": {}, "updated": "", "in_sheet": False}
        self.set_case(new[0] if new else nums[0])
        if len(new) > 1:
            self._notify(f"{len(new)} cases added for {self.cur_op}")

    def next_case(self) -> None:
        """Hotkey: go to the operator's next unfinished case (or ask for new case numbers)."""
        if self.cur_op is None:
            self.new_operator()
            return
        cases = self.operators[self.cur_op]["cases"]
        if not cases:
            self.add_cases()
            return
        names = list(cases)
        start = names.index(self.cur_case) if self.cur_case in cases else -1
        for n in names[start + 1:] + names[:start + 1]:
            if n != self.cur_case and not self._is_complete(cases[n]):
                self.set_case(n)
                self._notify(f"Next case: {n}")
                return
        self._notify(f"All cases of {self.cur_op} are done. Add more case numbers.")
        self.add_cases()

    def delete_case(self) -> None:
        if self.cur_op is None or self.cur_case is None:
            return
        if not confirm("Delete case", f"Delete case {self.cur_case} of {self.cur_op}?\n"
                       "(Rows already saved in the sheet stay there.)"):
            return
        od = self.operators[self.cur_op]
        names = list(od["cases"])
        i = names.index(self.cur_case)
        del od["cases"][self.cur_case]
        rest = list(od["cases"])
        self.cur_case = rest[min(i, len(rest) - 1)] if rest else None
        od["last_case"] = self.cur_case
        self.history.clear()
        self.refresh_all()
        self.schedule_save()

    def delete_operator(self) -> None:
        if self.cur_op is None:
            return
        if not confirm("Delete operator", f"Delete operator {self.cur_op} and all their cases?\n"
                       "(Rows already saved in the sheet stay there.)"):
            return
        del self.operators[self.cur_op]
        self.cur_op = next(iter(self.operators), None)
        self.cur_case = self.operators[self.cur_op].get("last_case") if self.cur_op else None
        self.history.clear()
        self.refresh_all()
        self.schedule_save()

    # --------------------------------------------------------------- sheet #
    def _build_row(self, op: str, case_no: str, case: dict[str, Any]) -> list[Any]:
        j, p, f, total = self.case_counts(case)
        result = "INCOMPLETE" if j < total else ("FAIL" if f else "PASS")
        stamp = case.get("updated") or now_iso()
        cells = [{"pass": "Pass", "fail": "Fail"}.get(case["status"].get(str(c["id"])), "")
                 for c in self.checklist]
        return [case.get("date") or today_str(), stamp[11:19] or datetime.now().strftime("%H:%M:%S"),
                op, case_no, *cells, p, f, result]

    def _queue_row(self, op: str | None, case_no: str | None, case: dict[str, Any]) -> None:
        if not op or not case_no:
            return
        row = self._build_row(op, case_no, case)
        key = ResultsSheet.key(row)
        case["in_sheet"] = True
        if self._last_rows.get(key) == tuple(row[2:]) and key not in self.pending_rows:
            return                                    # nothing changed since the last write
        self.pending_rows[key] = row
        self.flush_sheet()

    def _make_sheet(self):
        texts = [c["text"] for c in self.checklist]
        if self.sheet_cfg.get("url") and self.sheet_cfg.get("token"):
            return GoogleSheet(self.sheet_cfg["url"], self.sheet_cfg["token"], ResultsSheet.make_headers(texts))
        return ResultsSheet(app_dir(), texts)

    def flush_sheet(self) -> None:
        """Send pending rows in a background thread (never blocks the toasts / hotkeys)."""
        if not self.pending_rows or self._flushing:
            return
        self._flushing = True
        rows = list(self.pending_rows.values())
        self._flush_thread = threading.Thread(target=self._flush_worker, args=(self.sheet, rows), daemon=True)
        self._flush_thread.start()

    def _flush_worker(self, sheet, rows: list[list[Any]]) -> None:
        err = ""
        try:
            sheet.upsert(rows)
        except PermissionError:
            err = "the file is open in Excel. Close it"
        except Exception as exc:
            err = str(exc) or exc.__class__.__name__
        self.sheet_result.emit(rows, err)

    @pyqtSlot(object, str)
    def _on_sheet_result(self, rows: list[list[Any]], err: str) -> None:
        self._flushing = False
        if err:
            print(f"[warn] sheet upload failed: {err}", file=sys.stderr)
            if time.monotonic() - self._last_sheet_warn > 60:
                self._last_sheet_warn = time.monotonic()
                self._notify(f"Could not save to {self.sheet.label}: {err}. Will retry automatically.")
            self.schedule_save()                  # keep the unsent rows on disk
            return
        for r in rows:
            key = ResultsSheet.key(r)
            if self.pending_rows.get(key) == r:
                del self.pending_rows[key]
            self._last_rows[key] = tuple(r[2:])
        self.schedule_save()
        if self.pending_rows:
            QTimer.singleShot(0, self.flush_sheet)  # rows added while uploading

    def _final_flush(self) -> None:
        """On quit: wait for a running upload, then try once more; leftovers stay in the state file."""
        t = self._flush_thread
        if t is not None and t.is_alive():
            t.join(12)
        QApplication.processEvents()
        if self.pending_rows:
            rows = list(self.pending_rows.values())
            try:
                self.sheet.upsert(rows)
                self.pending_rows.clear()
            except Exception as exc:
                print(f"[warn] final upload failed: {exc}", file=sys.stderr)

    def open_sheet(self) -> None:
        if self.sheet.is_google:
            QDesktopServices.openUrl(QUrl(GOOGLE_SHEET_URL))
        elif self.sheet.path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.sheet.path)))
        else:
            self._notify("No results yet. A row is saved when a case is completed.")

    def connect_google(self) -> None:
        """Ask for the Apps Script web-app URL + secret, test them, then use the Google Sheet."""
        url = ask_text("Connect Google Sheet",
                       "Paste the Web app URL from Apps Script (it ends with /exec):", self.sheet_cfg.get("url", ""))
        if not url:
            return
        if not GoogleSheet.valid_url(url):
            info("Connect Google Sheet", "That doesn't look like a Web app URL.\n"
                 "It should start with https://script.google.com/macros/s/ and end with /exec")
            return
        token = ask_text("Connect Google Sheet", "Secret (the SECRET text you set in the script):",
                         self.sheet_cfg.get("token", ""))
        if not token:
            return
        candidate = GoogleSheet(url, token, ResultsSheet.make_headers([c["text"] for c in self.checklist]))
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            candidate.ping()
        except Exception as exc:
            QApplication.restoreOverrideCursor()
            info("Connect Google Sheet", f"Could not connect:\n{exc}")
            return
        QApplication.restoreOverrideCursor()
        self.sheet_cfg = {"url": url.strip(), "token": token.strip()}
        self.sheet = candidate
        self._last_rows.clear()
        self.schedule_save()
        self._notify("Google Sheet connected ✓")
        n = sum(1 for od in self.operators.values() for cd in od["cases"].values() if self.case_counts(cd)[0])
        if n and confirm("Upload existing results", f"Upload {n} existing case(s) to the Google Sheet now?"):
            self.upload_all()

    def disconnect_google(self) -> None:
        if confirm("Google Sheet", "Stop using the Google Sheet and save to a local Excel file instead?"):
            self.sheet_cfg = {"url": "", "token": ""}
            self.sheet = self._make_sheet()
            self._last_rows.clear()
            self.schedule_save()

    def upload_all(self) -> None:
        """Queue every case that has marks (used right after connecting a sheet)."""
        for op, od in self.operators.items():
            for cn, cd in od["cases"].items():
                if self.case_counts(cd)[0]:
                    self._queue_row(op, cn, cd)
        self.flush_sheet()

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
        for t in self._all_toasts():
            if reset_custom:
                t.custom_pos = False
            t.fit()
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

    # --------------------------------------------------------------- actions #
    def set_hide_judged(self, value: bool) -> None:
        self.hide_judged = bool(value)
        self.refresh_all()
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
        self._leave_case()
        self._final_flush()
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
        """Right-click / header-click menu."""
        m = QMenu()
        m.setStyleSheet(MENU_STYLE)
        m.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self._fill_menu(m, toast)
        return m

    def popup_menu(self, toast: Toast | None, global_pos: QPoint) -> None:
        """Show the menu above the toasts; the overlay stops re-raising itself while it is open."""
        m = self.build_menu(toast)
        TOPMOST_PAUSE[0] += 1
        try:
            m.exec(global_pos)
        finally:
            TOPMOST_PAUSE[0] -= 1

    def _tray_menu_opening(self) -> None:
        self._fill_menu(self.tray_menu, None)
        if not self._tray_open:
            self._tray_open = True
            TOPMOST_PAUSE[0] += 1

    def _tray_menu_closing(self) -> None:
        if self._tray_open:
            self._tray_open = False
            TOPMOST_PAUSE[0] -= 1

    def _fill_menu(self, m: QMenu, toast: Toast | None) -> None:
        """Populate a menu; rebuilt every time it opens so its state is always fresh."""
        m.clear()
        hk = {k: pretty_hotkey(v) for k, v in self.hotkeys.items()}
        if toast is not None and toast.kind == "item":
            m.addAction("✓ Pass this item", lambda: self.set_status(toast.key, "pass"))
            m.addAction("✗ Fail this item", lambda: self.set_status(toast.key, "fail"))
            m.addAction("Clear this item", lambda: self.set_status(toast.key, None))
            m.addSeparator()

        opm = m.addMenu(f"Operator: {self.cur_op or '(none)'}")
        for name in self.operators:
            a = opm.addAction(name)
            a.setCheckable(True)
            a.setChecked(name == self.cur_op)
            a.triggered.connect(lambda _=False, n=name: self.switch_operator(n))
        if self.operators:
            opm.addSeparator()
        opm.addAction("➕ New operator…", lambda: self.new_operator())
        if self.cur_op:
            opm.addAction(f"🗑 Delete operator “{self.cur_op}”…", lambda: self.delete_operator())

        cm = m.addMenu(f"Case: {self.cur_case or '(none)'}")
        if self.cur_op:
            for cn, cd in self.operators[self.cur_op]["cases"].items():
                a = cm.addAction(f"{self._case_icon(cd)}  {cn}")
                a.setCheckable(True)
                a.setChecked(cn == self.cur_case)
                a.triggered.connect(lambda _=False, n=cn: self.set_case(n))
            if self.operators[self.cur_op]["cases"]:
                cm.addSeparator()
            cm.addAction("➕ Add case(s)…", lambda: self.add_cases())
            if self.cur_case:
                cm.addAction(f"🗑 Delete case “{self.cur_case}”…", lambda: self.delete_case())
        else:
            cm.setEnabled(False)

        m.addAction(f"⏭ Next case    ({hk['next_case']})", lambda: self.next_case())
        m.addSeparator()
        m.addAction(f"✓ Pass next item    ({hk['pass']})", lambda: self.hotkey_judge("pass"))
        m.addAction(f"✗ Fail next item    ({hk['fail']})", lambda: self.hotkey_judge("fail"))
        m.addAction(f"↶ Undo last    ({hk['undo']})", lambda: self.undo())
        m.addAction(f"↺ Reset this case…    ({hk['reset']})", lambda: self.reset_case())
        m.addSeparator()
        m.addAction("📊 Open results sheet", lambda: self.open_sheet())
        if self.sheet.is_google:
            m.addAction("☁ Google Sheet: connected ✓   (change…)", lambda: self.connect_google())
            m.addAction("Stop using Google Sheet…", lambda: self.disconnect_google())
        else:
            m.addAction("☁ Connect Google Sheet…", lambda: self.connect_google())
        m.addAction(f"Show / hide toasts    ({hk['toggle_visible']})", lambda: self.toggle_visible())
        hj = m.addAction("Hide judged toasts")
        hj.setCheckable(True)
        hj.setChecked(self.hide_judged)
        hj.triggered.connect(lambda checked: self.set_hide_judged(checked))
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
            m.addAction(f"⟳ Restart to update to v{self.update_ready}", lambda: self.apply_update())
        else:
            m.addAction("Check for updates", lambda: self.updater.check(manual=True))
        m.addAction(f"Version {APP_VERSION}").setEnabled(False)
        m.addAction("Quit", lambda: self.quit())
        for a in m.actions():                    # submenus created above must be topmost as well
            if a.menu() is not None:
                make_menu_topmost(a.menu())


# --------------------------------------------------------------------------- #
TOAST_STYLE = """
QWidget { font-family: "Segoe UI", "Noto Sans", "DejaVu Sans", sans-serif; }
QFrame { background: transparent; border: none; }
#card { background-color: #1e2127; border: 1px solid #3a3f4b; border-left: 4px solid #4a90d9; border-radius: 8px; }
#card[kind="header"] { background-color: #202733; border-left-color: #8ab4f8; }
#card[kind="header"][state="pass"] { background-color: #1b2420; border-left-color: #34a853; }
#card[kind="header"][state="fail"] { background-color: #2a1b1d; border-left-color: #ea4335; }
#card[kind="item"][state="pass"] { background-color: #1b2420; border-left-color: #34a853; }
#card[kind="item"][state="fail"] { background-color: #2a1b1d; border-left-color: #ea4335; }
#card[kind="item"][current="true"] { border: 1px solid #fbbc04; border-left: 4px solid #fbbc04; }
#num { color: #6b7280; font-size: 11px; font-weight: 600; background: transparent; border: none; }
#text { color: #f1f3f4; font-size: 14px; background: transparent; border: none; }
#pbtn, #fbtn { border: 1.5px solid #5f6672; border-radius: 6px; background: #14161a; color: #7b828c;
               font-size: 15px; font-weight: bold; }
#pbtn[active="true"] { background: #34a853; border-color: #34a853; color: white; }
#fbtn[active="true"] { background: #ea4335; border-color: #ea4335; color: white; }
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

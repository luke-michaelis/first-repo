# version 1.166
# Latest version of the program with OneFile Build in mind

import os
import threading
import sys
import time
import json
import socket
import subprocess
import datetime
import svgwrite
import pyautogui
import serial  # for Arduino pulses
from serial.tools import list_ports

import logging   # for detailed crash tracing
import shutil    # still here in case

import ctypes
from ctypes import wintypes
from functools import partial

from PyQt6.QtCore    import Qt, QThread, pyqtSignal
from PyQt6.QtGui     import QPixmap, QGuiApplication, QFont, QFontMetrics
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QComboBox, QVBoxLayout, QHBoxLayout, QMessageBox, QFrame,
    QPlainTextEdit, QStackedWidget, QCheckBox, QSizePolicy,
    QListWidget, QTableWidget, QTableWidgetItem, QAbstractItemView,
    QGridLayout, QFileDialog, QButtonGroup
)

# ─── Layer / Color mapping for LightBurn ───────────────────────────────────────
# LightBurn typically maps imported object colors to layers.
# We'll use exact hex colors to keep it predictable.
COLOR_HEX = {
    "Silver":    "#000000",  # black
    "Brass":     "#0000FF",  # blue
    "Plastic":   "#FF0000",  # red
    "Stainless": "#00FF00",  # green
}
COLOR_NAMES = ["Silver", "Brass", "Plastic", "Stainless"]

def normalize_color_name(s: str) -> str:
    """
    Normalize an incoming color name to one of the canonical keys in COLOR_HEX.
    Falls back to Silver (black) if unknown.
    """
    s = (s or "").strip().lower()
    for name in COLOR_NAMES:
        if s == name.lower():
            return name
    return "Silver"

# ─── BASE DIR & LOG FOLDER ────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    exe_dir = os.path.dirname(sys.executable)
else:
    exe_dir = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(exe_dir, "Logs")
os.makedirs(LOG_DIR, exist_ok=True)

# ─── Logging configuration ────────────────────────────────────────────────────
log_path = os.path.join(LOG_DIR, "app.log")
logging.basicConfig(
    filename=log_path,
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)-8s %(message)s'
)
def excepthook(exc_type, exc_value, exc_tb):
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))
sys.excepthook = excepthook

# ─── SINGLE INSTANCE ENFORCEMENT (WINDOWS NAMED MUTEX) ────────────────────────
_single_instance_mutex_handle = None
_MUTEX_NAME = "Laser"

def enforce_single_instance():
    """
    Ensure only ONE instance of this program is running on Windows by using a
    system-wide named mutex.
    """
    global _single_instance_mutex_handle

    if os.name != "nt":
        logging.debug("Non-Windows OS detected — skipping mutex single-instance enforcement.")
        return True

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        CreateMutexW = kernel32.CreateMutexW
        CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        CreateMutexW.restype  = wintypes.HANDLE

        ERROR_ALREADY_EXISTS = 183

        handle = CreateMutexW(None, False, _MUTEX_NAME)
        if not handle:
            err = ctypes.get_last_error()
            logging.error(f"CreateMutexW failed (error {err}); failing open and allowing instance.")
            return True

        _single_instance_mutex_handle = handle
        last_err = ctypes.get_last_error()

        if last_err == ERROR_ALREADY_EXISTS:
            logging.warning("Named mutex already exists — another instance is running.")
            return False

        logging.debug("Named mutex created — this is the primary instance.")
        return True

    except Exception:
        logging.error("Exception during enforce_single_instance(); failing open.", exc_info=True)
        return True

def show_single_instance_warning_topmost():
    """
    Show a top-most 'Already running' warning.
    """
    msg = (
        "Another instance of this application is already running.\n\n"
        "This window will now close."
    )
    title = "Already running"

    if os.name == "nt":
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            MB_OK = 0x00000000
            MB_ICONWARNING = 0x00000030
            MB_TOPMOST = 0x00040000
            style = MB_OK | MB_ICONWARNING | MB_TOPMOST
            user32.MessageBoxW(None, msg, title, style)
            return
        except Exception:
            logging.error("Failed to show topmost Win32 MessageBox, falling back to Qt.", exc_info=True)

    QMessageBox.warning(None, title, msg)

# ─── AUTO-DISMISS “Save Project?” POPUP ────────────────────────────────────────
auto_dismiss_event = threading.Event()
def auto_dismiss_popups():
    while not auto_dismiss_event.is_set():
        try:
            for dlg in pyautogui.getWindowsWithTitle("Save Project?"):
                dlg.activate()
                pyautogui.press("right")
                pyautogui.press("enter")
        except Exception:
            pass
        time.sleep(0.2)
threading.Thread(target=auto_dismiss_popups, daemon=True).start()

# ─── CONFIG ───────────────────────────────────────────────────────────────────
GENERATED    = os.path.join(exe_dir, "generated_jobs")
PRESETS_F    = os.path.join(exe_dir, ".laser_presets.json")
STENCILS_F   = os.path.join(exe_dir, ".laser_stencils.json")
LOGO         = os.path.join(exe_dir, "453cf780-195e-4bdc-aba6-617024c9dd4d.png")
LB_EXE       = r"C:\Program Files\LightBurn\LightBurn.exe"
ARDUINO_BAUD = 115200
os.makedirs(GENERATED, exist_ok=True)

DEFAULT_PRESETS = {
    "Preset 1": {"x": 50.0, "y": 50.0, "font": 5.0, "offset": 26.0, "color": "Silver"},
    "Preset 2": {"x": 50.0, "y": 52.0, "font": 6.0, "offset": 26.0, "color": "Brass"},
    "Preset 3": {"x": 50.0, "y": 54.0, "font": 7.0, "offset": 26.0, "color": "Plastic"},
}

# ─── PRESETS I/O ───────────────────────────────────────────────────────────────
def load_presets():
    try:
        data = json.load(open(PRESETS_F))
        out = {}
        for k, v in data.items():
            out[k] = {
                "x": v["x"],
                "y": v["y"],
                "font": v["font"],
                "offset": v.get("offset", DEFAULT_PRESETS.get(k, {}).get("offset", 26.0)),
                "color": normalize_color_name(v.get("color", DEFAULT_PRESETS.get(k, {}).get("color", "Silver"))),
            }
        return out
    except:
        d = DEFAULT_PRESETS.copy()
        for k in d:
            d[k]["color"] = normalize_color_name(d[k].get("color", "Silver"))
        return d

def save_presets_file(presets_dict):
    with open(PRESETS_F, 'w') as f:
        json.dump(presets_dict, f, indent=2)

# ─── STENCILS I/O ───────────────────────────────────────────────────────────────
def load_stencils():
    try:
        data = json.load(open(STENCILS_F))
        if isinstance(data, dict):
            return data
        return {name: "Preset 1" for name in data}
    except:
        return {}

def save_stencils(stencils_map):
    with open(STENCILS_F, 'w') as f:
        json.dump(stencils_map, f, indent=2)

# ─── UDP HELPERS ───────────────────────────────────────────────────────────────
sock_rcv = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock_rcv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
sock_rcv.bind(("127.0.0.1", 19841))

def send_cmd(cmd, timeout=0.1):
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.sendto(cmd.encode(), ("127.0.0.1", 19840))
        sock_rcv.settimeout(timeout)
        data, _ = sock_rcv.recvfrom(1024)
        return data.decode().strip()
    except socket.timeout:
        return None
    finally:
        s.close()

def ensure_lightburn():
    subprocess.Popen([LB_EXE], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    start = time.time()
    while time.time() - start < 10:
        if send_cmd("PING") == "OK":
            return
        time.sleep(0.2)
    raise RuntimeError("LightBurn handshake failed")

# ─── SVG GENERATION ────────────────────────────────────────────────────────────
def generate_svg_layers(t1, p1, t2, p2, t3, p3, copies):
    created = []

    def draw_copies(dwg, txt, p, canvas_size, filename):
        x0, y0, fz, off = p["x"], p["y"], p["font"], p["offset"]
        color_name = normalize_color_name(p.get("color", "Silver"))
        hex_color  = COLOR_HEX[color_name]

        if copies == 1:
            positions = [(0,0)]
        elif copies == 2:
            positions = [(0,0),(0,off)]
        else:
            positions = [(0,0),(off,0),(0,off),(off,off)]

        for idx, (dx,dy) in enumerate(positions):
            xi, yi = x0+dx, y0+dy
            attrs = {
                "insert": (f"{xi}mm", f"{yi}mm"),
                "font_size": f"{fz}mm",
                "text_anchor": "middle",

                # Color metadata for LightBurn layer mapping:
                "fill": hex_color,
                "stroke": hex_color,
                "stroke_width": 0,
            }
            if copies==2 and idx==1:
                attrs["transform"] = f"rotate(180 {xi} {yi})"
            dwg.add(dwg.text(txt, **attrs))

    if t1:
        path00 = os.path.join(GENERATED,"laser_job_c00.svg")
        dwg1 = svgwrite.Drawing(path00,size=("100mm","100mm"))
        draw_copies(dwg1,t1,p1,(100,100),path00)
        dwg1.save(); created.append(path00)

    if t2 or t3:
        path01 = os.path.join(GENERATED,"laser_job_c01.svg")
        dwg2 = svgwrite.Drawing(path01,size=("150mm","150mm"))
        entries=[]
        if t2: entries.append((t2,p2))
        if t3:
            p3m=p3.copy()
            p3m["y"]=(p2["y"]+4.0) if t2 else p3["y"]
            entries.append((t3,p3m))
        for txt,p in entries:
            draw_copies(dwg2,txt,p,(150,150),path01)
        dwg2.save(); created.append(path01)

    return created

# ─── BACKGROUND THREAD ─────────────────────────────────────────────────────────
class LoopThread(QThread):
    log   = pyqtSignal(str)
    ready = pyqtSignal()

    def __init__(self, layers, logfile, com_port, parent=None):
        super().__init__(parent)
        self.layers      = layers
        self.running     = True
        self.logf        = open(logfile,"w",encoding="utf8")
        self.current_idx = 0
        self.port        = com_port

        ports = [p.device for p in list_ports.comports()]
        self.logmaj(f"[DEBUG] Available COM ports: {ports}")

        try:
            self.ser = serial.Serial(self.port, ARDUINO_BAUD, timeout=0.1)
            time.sleep(2)
            self.logmaj(f"[DEBUG] Serial port opened on {self.port}")
        except Exception as e:
            self.logmaj(f"[WARNING] Could not open port {self.port}: {e}")
            self.ser = None

    def ts(self):
        return datetime.datetime.now().isoformat()

    def logmaj(self, msg, to_gui=False):
        line = f"{self.ts()} {msg}"
        self.logf.write(line + "\n"); self.logf.flush()
        if to_gui:
            self.log.emit(line)

    def _force_load_index(self, idx):
        path = self.layers[idx]
        self.logmaj(f"→ FORCELOAD {os.path.basename(path)}")
        send_cmd(f"FORCELOAD:{path}")
        st = send_cmd("STATUS")
        while st != "OK" and self.running:
            time.sleep(0.05); st = send_cmd("STATUS")
        if not self.running: return False
        self.logmaj("✅ load confirmed", to_gui=True)
        time.sleep(0.2)
        return True

    def run(self):
        self.logmaj("[DEBUG] LoopThread.run: Starting")
        ensure_lightburn()
        time.sleep(0.5)
        if not self._force_load_index(self.current_idx): return
        self.logmaj("→ START command for initial layer", to_gui=True)
        send_cmd("START"); self.ready.emit()

        while self.running:
            if not self.ser:
                time.sleep(0.1); continue

            raw = self.ser.readline()
            self.logmaj(f"[DEBUG] RAW SERIAL: {raw!r}")
            try:
                line = raw.decode("utf-8", errors="ignore").strip()
                self.logmaj(f"[DEBUG] DECODED LINE: {line!r}")
            except:
                continue
            self.logmaj(f"[DEBUG] Arduino EVENT: {line!r}")

            if "🔻 FALLING" in line:
                self.logmaj("🔻 FALLING detected", to_gui=True)
                self.current_idx = (self.current_idx + 1) % len(self.layers)
                if not self._force_load_index(self.current_idx): break
                self.ready.emit()

            elif "⚡️ RISING" in line:
                self.logmaj("⚡️ RISING detected", to_gui=True)
                self.logmaj("→ START command sent", to_gui=True)
                send_cmd("START"); time.sleep(1)

        subprocess.Popen(["taskkill","/IM","LightBurn.exe","/F"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.logmaj("[DEBUG] LoopThread.run: Finished")
        self.logf.close()
        if getattr(self, "ser", None):
            try:    self.ser.close()
            except: pass

# ─── MAIN GUI ─────────────────────────────────────────────────────────────────
class LaserGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.presets      = load_presets()
        self.stencils_map = load_stencils()
        self.stencils     = list(self.stencils_map.keys())
        self.current_layers = []
        self.worker       = None
        self.selected_port = None

        self._preset_table_populating = False  # prevents feedback loops

        self.text1 = QLineEdit(); self.text2 = QLineEdit(); self.text3 = QLineEdit()
        default_pt = self.text1.font().pointSize()
        big_font = QFont(); big_font.setPointSize(default_pt * 2)
        for w in (self.text1, self.text2, self.text3):
            w.setFont(big_font)
            height = QFontMetrics(big_font).height() + 12
            w.setMinimumHeight(height)

        self.preset1 = QComboBox(); self.preset2 = QComboBox(); self.preset3 = QComboBox()

        # ── Line 2 stencil dropdown ─────────────────────────────────────────────
        self.stencil_combo = QComboBox()
        self.stencil_combo.currentTextChanged.connect(self._on_stencil_selected)
        self.stencil_combo.setMinimumWidth(120)
        self.stencil_combo.setStyleSheet("background-color: white;")

        # ── Line 2 color override checkboxes (exclusive) ───────────────────────
        self.line2_color = "Silver"  # default → black
        self.line2_color_group = QButtonGroup(self)
        self.line2_color_group.setExclusive(True)

        self.line2_color_checks = {}
        # Tiny labels so you can actually tell which is which without guessing like it's a loot box
        short = {"Silver":"S", "Brass":"B", "Plastic":"P", "Stainless":"SS"}
        tip   = {
            "Silver":    "Silver → Black",
            "Brass":     "Brass → Blue",
            "Plastic":   "Plastic → Red",
            "Stainless": "Stainless → Green"
        }
        for name in COLOR_NAMES:
            cb = QCheckBox(short[name])
            cb.setToolTip(tip[name])
            cb.setStyleSheet("background-color: transparent; color: white;")  # readable on dark bg
            cb.toggled.connect(partial(self.on_line2_color_toggled, name))
            self.line2_color_group.addButton(cb)
            self.line2_color_checks[name] = cb

        # Default selection: Silver (black)
        self.line2_color_checks["Silver"].setChecked(True)

        self.start_btn = QPushButton("Start")
        self.stop_btn  = QPushButton("Stop")
        self.close_btn = QPushButton("Close Program")
        for btn in (self.start_btn, self.stop_btn):
            h = btn.sizeHint(); btn.setFixedSize(h.width()*5, h.height()*5)
        self.start_btn.setStyleSheet("background-color: green; color: black;")
        self.stop_btn .setStyleSheet("background-color: red;   color: black;")
        self.close_btn.setStyleSheet("background-color: white; color: black;")

        self.layer_label = QLabel("Layer: –")
        Lf = self.layer_label.font(); Lf.setPointSize(18); Lf.setBold(True)
        self.layer_label.setFont(Lf)

        self.line_labels   = []
        self.line_edits    = [self.text1, self.text2, self.text3]
        self.preset_cbs    = [self.preset1, self.preset2, self.preset3]
        self.extra_widgets = []

        self.initUI()

    def on_line2_color_toggled(self, name, checked: bool):
        if checked:
            self.line2_color = normalize_color_name(name)
            logging.debug(f"Line 2 color override set to: {self.line2_color}")

    def _get_line2_color_override(self) -> str:
        # Always return a valid canonical name
        for name, cb in self.line2_color_checks.items():
            if cb.isChecked():
                return normalize_color_name(name)
        return "Silver"

    def initUI(self):
        self.setWindowTitle("Raycus Laser UI v1.166-dev")
        self.enter_borderless_fullscreen()

        self.stack    = QStackedWidget(self)
        main         = QWidget()
        hl_main      = QHBoxLayout(main); hl_main.setContentsMargins(20,10,20,10)
        vl_left      = QVBoxLayout()

        if os.path.exists(LOGO):
            logo = QLabel(); logo.setPixmap(QPixmap(LOGO).scaledToWidth(300))
            logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            vl_left.addWidget(logo)

        grid = QGridLayout(); grid.setVerticalSpacing(8); grid.setHorizontalSpacing(8)
        grid.setColumnStretch(1, 1); grid.setColumnMinimumWidth(2, 180); grid.setColumnMinimumWidth(3, 320)

        for row_idx, (label_text, line_edit, preset_cb) in enumerate([
            ("Line 1", self.text1, self.preset1),
            ("Line 2", self.text2, self.preset2),
            ("Line 3", self.text3, self.preset3),
        ]):
            lbl = QLabel(label_text); f = lbl.font(); f.setPointSize(18); lbl.setFont(f)
            grid.addWidget(lbl,       row_idx, 0)
            grid.addWidget(line_edit, row_idx, 1)
            grid.addWidget(preset_cb, row_idx, 2)

            if row_idx == 1:
                # Container: stencil dropdown + little color checkboxes
                right = QWidget()
                rh = QHBoxLayout(right)
                rh.setContentsMargins(0,0,0,0)
                rh.setSpacing(8)
                rh.addWidget(self.stencil_combo)
                rh.addSpacing(10)
                rh.addWidget(QLabel("L2:"))  # tiny label so you remember it's line-2 only
                for nm in COLOR_NAMES:
                    cbox = self.line2_color_checks[nm]
                    cbox.setFixedHeight(24)
                    rh.addWidget(cbox)
                rh.addStretch(1)
                grid.addWidget(right, row_idx, 3)
                self.extra_widgets.append(right)
            else:
                self.extra_widgets.append(None)

            self.line_labels.append(lbl)

        vl_left.addLayout(grid)
        vl_left.addWidget(self._sep())
        vl_left.addWidget(self.layer_label, alignment=Qt.AlignmentFlag.AlignHCenter)

        btns = QHBoxLayout()
        btns.addWidget(self.start_btn)
        btns.addWidget(self.stop_btn)
        vl_left.addLayout(btns)

        self.log = QPlainTextEdit(readOnly=True, placeholderText="Job status…")
        self.log.setFixedHeight(100); self.log.setMinimumWidth(600)
        vl_left.addWidget(self.log)

        hl_main.addLayout(vl_left, 1)
        vr = QVBoxLayout(); vr.addStretch()
        dev_btn = QPushButton("Developer Page")
        df      = dev_btn.font(); df.setPointSize(16); dev_btn.setFont(df)
        dev_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        vr.addWidget(dev_btn, alignment=Qt.AlignmentFlag.AlignCenter); vr.addStretch()
        hl_main.addLayout(vr)

        self.stack.addWidget(main)

        # ── Developer Page ──────────────────────────────────────────────────────
        dev = QWidget(); dv = QHBoxLayout(dev)
        nav = QVBoxLayout()
        nav_btn = QPushButton("Main Menu"); nav_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        nav.addWidget(nav_btn); nav.addStretch(); dv.addLayout(nav)

        content = QVBoxLayout()
        import_row = QHBoxLayout()
        self.import_config_btn = QPushButton("Import Config"); self.import_config_btn.clicked.connect(self.import_config)
        import_row.addWidget(self.import_config_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        content.addLayout(import_row)

        st_row = QHBoxLayout()
        self.new_stencil_input = QLineEdit(placeholderText="New stencil name")
        self.new_stencil_preset = QComboBox(); self.new_stencil_preset.addItems(list(self.presets.keys()))
        self.add_stencil_btn = QPushButton("Add Stencil"); self.add_stencil_btn.clicked.connect(self.add_stencil)
        st_row.addWidget(self.new_stencil_input); st_row.addWidget(self.new_stencil_preset); st_row.addWidget(self.add_stencil_btn)
        content.addLayout(st_row)

        self.stencil_list = QListWidget(); self.stencil_list.addItems(self.stencils)
        self.stencil_list.currentTextChanged.connect(self.on_stencil_list_selected)
        content.addWidget(self.stencil_list)

        self.remove_stencil_btn = QPushButton("Remove Selected"); self.remove_stencil_btn.clicked.connect(self.remove_stencil)
        content.addWidget(self.remove_stencil_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        edit_row = QHBoxLayout(); edit_row.addWidget(QLabel("Assign Preset:"))
        self.edit_stencil_preset = QComboBox(); self.edit_stencil_preset.addItems(list(self.presets.keys()))
        self.update_stencil_btn = QPushButton("Update Preset"); self.update_stencil_btn.clicked.connect(self.update_stencil_mapping)
        edit_row.addWidget(self.edit_stencil_preset); edit_row.addWidget(self.update_stencil_btn)
        content.addLayout(edit_row)

        content.addWidget(QLabel("<h2>Preset Management</h2>"), alignment=Qt.AlignmentFlag.AlignCenter)
        pm_row = QHBoxLayout(); self.new_preset_input = QLineEdit(placeholderText="New preset name")
        self.add_preset_btn = QPushButton("Add Preset"); self.add_preset_btn.clicked.connect(self.add_preset)
        pm_row.addWidget(self.new_preset_input); pm_row.addWidget(self.add_preset_btn)
        content.addLayout(pm_row)

        self.preset_table = QTableWidget(0,6)
        self.preset_table.setHorizontalHeaderLabels(["Preset","Color","X","Y","Font","Offset"])
        self.preset_table.verticalHeader().setVisible(False)
        self.preset_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.preset_table.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        self.preset_table.cellChanged.connect(self.on_preset_cell_changed)
        self.preset_table.itemSelectionChanged.connect(self.on_preset_table_selected)
        content.addWidget(self.preset_table)

        self.remove_preset_btn = QPushButton("Remove Selected Preset"); self.remove_preset_btn.clicked.connect(self.remove_preset)
        content.addWidget(self.remove_preset_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        editp_row = QHBoxLayout(); editp_row.addWidget(QLabel("Selected:"))
        self.selected_preset_label = QLabel(""); editp_row.addWidget(self.selected_preset_label)

        editp_row.addWidget(QLabel("color:"))
        self.preset_color_edit = QComboBox()
        self.preset_color_edit.addItems(COLOR_NAMES)
        editp_row.addWidget(self.preset_color_edit)

        self.preset_edit_fields = {}
        for param in ("x","y","font","offset"):
            editp_row.addWidget(QLabel(f"{param}:"))
            e = QLineEdit(); self.preset_edit_fields[param] = e; editp_row.addWidget(e)

        self.update_preset_btn = QPushButton("Update Preset"); self.update_preset_btn.clicked.connect(self.update_preset)
        editp_row.addWidget(self.update_preset_btn); content.addLayout(editp_row)

        next_btn = QPushButton("Next Layer"); next_btn.clicked.connect(self.on_next_layer)
        content.addWidget(next_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        self.dev_log = QPlainTextEdit(readOnly=True, placeholderText="Job status…")
        self.dev_log.setMinimumWidth(600)
        self.dev_log.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        content.addWidget(self.dev_log, 1)

        port_row = QHBoxLayout(); port_row.addWidget(QLabel("Serial Port:"))
        self.port_combo = QComboBox(); self.port_combo.setToolTip("Select your Arduino COM port")
        self.port_combo.currentTextChanged.connect(self.on_port_selected)
        port_row.addWidget(self.port_combo); content.addLayout(port_row)

        self.reboot_btn = QPushButton("Reboot ESP32-C3")
        self.reboot_btn.clicked.connect(self.on_reboot_c3)
        content.addWidget(self.reboot_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.windowed_chk = QCheckBox("Windowed Mode"); self.windowed_chk.toggled.connect(self.on_windowed_toggle)
        content.addWidget(self.windowed_chk, alignment=Qt.AlignmentFlag.AlignCenter)

        self.line_show_layout = QHBoxLayout(); self.line_checkboxes = []
        for i in range(3):
            cb = QCheckBox(f"Show Line {i+1}"); cb.setChecked(True)
            cb.toggled.connect(lambda checked, idx=i: self.toggle_line(idx, checked))
            self.line_show_layout.addWidget(cb); self.line_checkboxes.append(cb)
        content.addLayout(self.line_show_layout)

        copies_layout = QHBoxLayout(); copies_layout.addWidget(QLabel("Copies:"))
        self.copies_combo = QComboBox(); self.copies_combo.addItems(["1","2","4"])
        copies_layout.addWidget(self.copies_combo); content.addLayout(copies_layout)

        content.addSpacing(12)
        content.addWidget(self._sep())
        content.addSpacing(6)
        content.addWidget(self.close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        content.addStretch(); dv.addLayout(content); self.stack.addWidget(dev)

        layout = QVBoxLayout(self); layout.addWidget(self.stack); self.stack.setCurrentIndex(0)
        for btn in self.findChildren(QPushButton):
            if btn not in (self.start_btn, self.stop_btn):
                btn.setStyleSheet("background-color: white; color: black;")

        self.start_btn.clicked.connect(self.on_start)
        self.stop_btn.clicked.connect(self.on_stop)
        self.close_btn.clicked.connect(self.close)

        self.refresh_presets()
        self.refresh_stencils()
        self.refresh_ports()

        self._apply_launch_defaults()

    # ─── Launch defaults: uncheck Show Line 3, set copies=2 (EVERY start) ──────
    def _apply_launch_defaults(self):
        try:
            if len(self.line_checkboxes) >= 3:
                self.line_checkboxes[2].setChecked(False)
            i = self.copies_combo.findText("2")
            if i != -1:
                self.copies_combo.setCurrentIndex(i)
            logging.debug("Applied launch defaults: hide Line 3, Copies=2 (every launch)")
        except Exception as e:
            logging.error(f"Failed to apply launch defaults: {e}")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select JSON Config", "", "JSON Files (*.json)")
        if not path: return
        try:
            cfg = json.load(open(path))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load config: {e}")
            return

        for idx, cb in enumerate(self.line_checkboxes):
            if idx < len(cfg.get("show_lines", [])):
                cb.setChecked(bool(cfg["show_lines"][idx]))

        copies = str(cfg.get("copies", self.copies_combo.currentText()))
        if copies in [self.copies_combo.itemText(i) for i in range(self.copies_combo.count())]:
            self.copies_combo.setCurrentText(copies)

        self.windowed_chk.setChecked(bool(cfg.get("windowed_mode", self.windowed_chk.isChecked())))

        default_stencil = cfg.get("default_stencil")
        if default_stencil in [self.stencil_combo.itemText(i) for i in range(self.stencil_combo.count())]:
            self.stencil_combo.setCurrentText(default_stencil)

        for i, name in enumerate(cfg.get("line_presets", [])):
            combo = self.preset_cbs[i]
            idx = combo.findData(name)
            if idx != -1:
                combo.setCurrentIndex(idx)

        # Optional: import line2 color override if present
        l2c = cfg.get("line2_color")
        if l2c:
            l2c = normalize_color_name(l2c)
            if l2c in self.line2_color_checks:
                self.line2_color_checks[l2c].setChecked(True)

    def _on_stencil_selected(self, name):
        self.text2.setText("" if name=="-- none --" else name)

    def _sep(self):
        s = QFrame(); s.setFrameStyle(QFrame.Shape.HLine | QFrame.Shadow.Sunken); return s

    def toggle_line(self, idx, visible):
        self.line_labels[idx].setVisible(visible)
        self.line_edits[idx].setVisible(visible)
        self.preset_cbs[idx].setVisible(visible)
        if self.extra_widgets[idx]:
            self.extra_widgets[idx].setVisible(visible)

    def refresh_ports(self):
        ports = [p.device for p in list_ports.comports()]
        self.port_combo.clear()
        self.port_combo.addItems(ports or ["<no ports>"])
        self.selected_port = ports[0] if ports else None

    def on_port_selected(self, port):
        self.selected_port = port if port and port!="<no ports>" else None

    def on_reboot_c3(self):
        if self.worker and getattr(self.worker, "ser", None):
            try:
                self.worker.ser.write(b"REBOOT\n")
                self.worker.ser.flush()
                msg = f"{datetime.datetime.now().isoformat()} ↻ Sent REBOOT to ESP32-C3 via active worker serial"
                self.log.appendPlainText("↻ Sent REBOOT to ESP32-C3 via active worker serial")
                self.dev_log.appendPlainText(msg)
                return
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to send REBOOT on active port:\n{e}")

        if not self.selected_port:
            QMessageBox.warning(self, "Error", "No COM port selected")
            return

        try:
            tmp = serial.Serial(self.selected_port, ARDUINO_BAUD, timeout=1)
            time.sleep(0.1)
            tmp.write(b"REBOOT\n")
            tmp.flush()
            tmp.close()
            msg = f"{datetime.datetime.now().isoformat()} ↻ Sent REBOOT to ESP32-C3 on {self.selected_port}"
            self.dev_log.appendPlainText(msg)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open {self.selected_port} to send REBOOT:\n{e}")

    def refresh_stencils(self):
        self.stencils_map = load_stencils()
        self.stencils = list(self.stencils_map.keys())
        self.stencil_combo.clear()
        self.stencil_combo.addItem("-- none --")
        self.stencil_combo.addItems(self.stencils)
        self.stencil_list.clear()
        self.stencil_list.addItems(self.stencils)

    def on_stencil_list_selected(self, name):
        self.edit_stencil_preset.setCurrentText(self.stencils_map.get(name,"Preset 1"))

    def add_stencil(self):
        nm = self.new_stencil_input.text().strip()
        if not nm or nm in self.stencils_map:
            QMessageBox.warning(self,"Error",f"Stencil '{nm}' invalid or exists"); return
        self.stencils_map[nm] = self.new_stencil_preset.currentText()
        save_stencils(self.stencils_map); self.new_stencil_input.clear(); self.refresh_stencils()

    def remove_stencil(self):
        items = self.stencil_list.selectedItems()
        if items:
            nm = items[0].text()
            self.stencils_map.pop(nm,None)
            save_stencils(self.stencils_map); self.refresh_stencils()

    def update_stencil_mapping(self):
        nm = self.stencil_list.currentItem().text() if self.stencil_list.currentItem() else None
        if not nm: return
        self.stencils_map[nm] = self.edit_stencil_preset.currentText()
        save_stencils(self.stencils_map); self.refresh_stencils()

    # ─── Presets UI refresh (includes color) ───────────────────────────────
    def refresh_presets(self):
        sel = [cb.currentData() for cb in (self.preset1, self.preset2, self.preset3)]
        table_sel_name = self.selected_preset_label.text()

        self.presets = load_presets()
        mono = QFont("Courier New")

        for cb in (self.preset1,self.preset2,self.preset3):
            cb.blockSignals(True)
            cb.clear()
            cb.setFont(mono)
            for name,v in self.presets.items():
                c = normalize_color_name(v.get("color","Silver"))
                display = f"{name:<12}  C:{c:<9}  X:{v['x']:>5.1f}  Y:{v['y']:>5.1f}  F:{v['font']:>4.1f}  O:{v['offset']:>4.1f}"
                cb.addItem(display,name)
            cb.blockSignals(False)

        for cb, wanted in zip((self.preset1,self.preset2,self.preset3), sel):
            if wanted:
                idx = cb.findData(wanted)
                if idx != -1:
                    cb.setCurrentIndex(idx)

        self._preset_table_populating = True
        self.preset_table.blockSignals(True)
        self.preset_table.setRowCount(0)

        for i,(name,v) in enumerate(self.presets.items()):
            self.preset_table.insertRow(i)

            item = QTableWidgetItem(name)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.preset_table.setItem(i,0,item)

            cbox = QComboBox()
            cbox.addItems(COLOR_NAMES)
            cval = normalize_color_name(v.get("color","Silver"))
            cbox.setCurrentText(cval)
            cbox.currentTextChanged.connect(partial(self.on_preset_color_changed, i))
            self.preset_table.setCellWidget(i, 1, cbox)

            self.preset_table.setItem(i,2,QTableWidgetItem(f"{v['x']:.1f}"))
            self.preset_table.setItem(i,3,QTableWidgetItem(f"{v['y']:.1f}"))
            self.preset_table.setItem(i,4,QTableWidgetItem(f"{v['font']:.1f}"))
            self.preset_table.setItem(i,5,QTableWidgetItem(f"{v['offset']:.1f}"))

        self.preset_table.blockSignals(False)
        self._preset_table_populating = False

        if table_sel_name and table_sel_name in self.presets:
            for r in range(self.preset_table.rowCount()):
                if self.preset_table.item(r,0) and self.preset_table.item(r,0).text() == table_sel_name:
                    self.preset_table.selectRow(r)
                    self.on_preset_table_selected()
                    break

    def on_preset_color_changed(self, row, color_text):
        if self._preset_table_populating:
            return
        try:
            name_item = self.preset_table.item(row, 0)
            if not name_item:
                return
            name = name_item.text()
            c = normalize_color_name(color_text)
            self.presets[name]["color"] = c
            save_presets_file(self.presets)
            self.refresh_presets()
        except Exception as e:
            logging.error(f"Failed to change preset color: {e}", exc_info=True)

    def add_preset(self):
        name = self.new_preset_input.text().strip()
        if not name or name in self.presets:
            QMessageBox.warning(self,"Error",f"Preset '{name}' invalid or exists"); return
        self.presets[name] = {"x":50.0,"y":50.0,"font":5.0,"offset":26.0,"color":"Silver"}
        save_presets_file(self.presets); self.new_preset_input.clear(); self.refresh_presets()

    def on_preset_table_selected(self):
        items = self.preset_table.selectedItems()
        if not items: return
        name = items[0].text()
        self.selected_preset_label.setText(name)
        vals = self.presets[name]
        self.preset_color_edit.setCurrentText(normalize_color_name(vals.get("color","Silver")))
        for p,fld in self.preset_edit_fields.items():
            fld.setText(str(vals[p]))

    def remove_preset(self):
        items = self.preset_table.selectedItems()
        if not items: return
        name = self.preset_table.item(items[0].row(), 0).text()
        self.presets.pop(name,None)
        save_presets_file(self.presets); self.refresh_presets()

    def update_preset(self):
        name = self.selected_preset_label.text()
        if not name: return
        try:
            vals = {p: float(self.preset_edit_fields[p].text()) for p in self.preset_edit_fields}
        except ValueError:
            QMessageBox.warning(self,"Error","All numeric fields must be numeric"); return

        vals["color"] = normalize_color_name(self.preset_color_edit.currentText())
        self.presets[name] = vals
        save_presets_file(self.presets)
        QMessageBox.information(self,"Success",f"Preset '{name}' updated")
        self.refresh_presets()

    def on_preset_cell_changed(self, row, col):
        if self._preset_table_populating:
            return
        if col not in {2,3,4,5}:
            return
        param = ("x","y","font","offset")[col-2]
        name = self.preset_table.item(row,0).text()
        text = self.preset_table.item(row,col).text()
        try:
            v = float(text)
        except ValueError:
            QMessageBox.warning(self,"Error",f"'{text}' not numeric"); self.refresh_presets(); return
        self.presets[name][param] = v
        save_presets_file(self.presets)
        self.preset_table.selectRow(row)
        self.on_preset_table_selected()
        self.refresh_presets()

    def on_start(self):
        logging.debug("╸ ENTER on_start()")
        try:
            global sock_rcv
            try: sock_rcv.close()
            except: pass
            sock_rcv = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock_rcv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock_rcv.bind(("127.0.0.1",19841))

            logging.debug("╸ Validating inputs…")
            t1,t2,t3 = self.text1.text().strip(),self.text2.text().strip(),self.text3.text().strip()
            logging.debug(f"╸ Inputs: {t1!r}, {t2!r}, {t3!r}")
            if not (t1 or t2 or t3):
                QMessageBox.warning(self,"Error","At least one line must be filled out"); return
            if not self.selected_port:
                QMessageBox.critical(self,"Error","No COM port selected"); return

            copies = int(self.copies_combo.currentText())
            logging.debug(f"╸ Copies: {copies}")

            if auto_dismiss_event.is_set():
                auto_dismiss_event.clear()
                threading.Thread(target=auto_dismiss_popups, daemon=True).start()

            if t1:
                t1 = f"{t1} {datetime.datetime.now().strftime('%V%y')}"

            # ── Apply your rules:
            # Line 1 always black (Silver).
            p1 = self.presets[self.preset1.currentData()].copy()
            p1["color"] = "Silver"

            # Line 2 gets override from the new checkbox cluster (Silver/Brass/Plastic/Stainless).
            p2 = self.presets[self.preset2.currentData()].copy()
            p2["color"] = self._get_line2_color_override()

            # Line 3 stays whatever preset says.
            p3 = self.presets[self.preset3.currentData()].copy()

            logging.debug("╸ Generating SVG layers…")
            layers = generate_svg_layers(
                t1, p1,
                t2, p2,
                t3, p3,
                copies
            )
            logging.debug(f"╸ Layers: {layers}")
            if not layers:
                QMessageBox.critical(self,"Error","No SVG layers created"); return

            self.current_layers = layers
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            logfile = os.path.join(LOG_DIR,f"laser_job_{ts}.txt")
            logging.debug(f"╸ Launching LoopThread with {logfile}")
            self.worker = LoopThread(layers, logfile, self.selected_port)
            self.worker.log.connect(self.log.appendPlainText)
            self.worker.log.connect(self.dev_log.appendPlainText)
            self.worker.ready.connect(self.update_layer_label)

            self.worker.start()
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)

        except Exception:
            logging.error("❌ Crash in on_start()", exc_info=True)
            QMessageBox.critical(self,"Error",f"Crash—see {log_path}")

    def on_next_layer(self):
        if not(self.current_layers and self.worker):
            QMessageBox.warning(self,"Error","No layers loaded"); return
        self.worker.current_idx = (self.worker.current_idx+1)%len(self.current_layers)
        self.worker._force_load_index(self.worker.current_idx)
        self.update_layer_label()
        self.log.appendPlainText("🖱 Next Layer")

    def on_stop(self):
        auto_dismiss_event.set()
        if self.worker:
            self.worker.running = False
            self.worker.wait()
            if getattr(self.worker, "ser", None):
                try:    self.worker.ser.close()
                except: pass
        self.kill_lightburn()
        self.log.appendPlainText("⏹ Stopped by user")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def update_layer_label(self):
        if self.worker and self.current_layers:
            idx = self.worker.current_idx
            total = len(self.current_layers)
            self.layer_label.setText(f"Layer: {idx+1}/{total}")

    def enter_borderless_fullscreen(self):
        self._prev_geom  = self.geometry()
        self._prev_flags = int(self.windowFlags())
        screen = QGuiApplication.primaryScreen().geometry()
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.show(); self.setGeometry(screen)

    def exit_borderless_fullscreen(self):
        if hasattr(self,"_prev_flags"):
            wf = Qt.WindowType(self._prev_flags)
            wf &= ~Qt.WindowType.FramelessWindowHint
            wf &= ~Qt.WindowType.WindowStaysOnTopHint
            self.setWindowFlags(wf)
        self.showNormal(); self.setGeometry(self._prev_geom)

    def on_windowed_toggle(self, checked: bool):
        if checked: self.exit_borderless_fullscreen()
        else:       self.enter_borderless_fullscreen()

    def kill_lightburn(self):
        try:
            subprocess.Popen(["taskkill","/IM","LightBurn.exe","/F"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            logging.error(f"Failed to kill LightBurn: {e}")

    def closeEvent(self, event):
        self.kill_lightburn(); event.accept()

if __name__ == "__main__":
    logging.debug("╸ Application starting")

    app = QApplication(sys.argv)

    if not enforce_single_instance():
        show_single_instance_warning_topmost()
        sys.exit(0)

    gui = LaserGUI()
    logging.debug("╸ Showing GUI")
    sys.exit(app.exec())

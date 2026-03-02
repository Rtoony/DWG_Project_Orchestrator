# DWG_Project_Orchestrator_Merged.py
# A combination of the feature-rich UI (from the old version)
# and the stable automation engine (from the new version).

import csv
import json
import re
import sys
import os
import time
import subprocess
import traceback
import datetime
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple, Union

# Suppress font table warnings from PyQt6/system fonts
warnings.filterwarnings("ignore", message=".*'name' table stringOffset.*")
warnings.filterwarnings("ignore", category=UserWarning, module=".*font.*")

# --- Python Environment Check ---
try:
    from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, pyqtSlot
    from PyQt6.QtGui import QFont, QColor
    from PyQt6.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
        QLineEdit, QPushButton, QTabWidget, QComboBox, QSplitter, QGroupBox, QMessageBox,
        QTreeWidget, QTreeWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView,
        QAbstractItemView, QListWidget, QListWidgetItem, QPlainTextEdit, QTreeWidgetItemIterator,
        QFormLayout, QFileDialog, QProgressBar, QTextEdit, QScrollArea, QCheckBox
    )
    # Imports for the automation engine
    from win32com.client.gencache import EnsureDispatch
    from win32com.client import VARIANT
    from pythoncom import VT_ARRAY, VT_R8
    import pywintypes
    import win32wnet
    # Import our new configuration manager
    from config_manager import ConfigurationManager, Rule
    
    # Import DXF analyzer
    from dxf_analyzer import DXFAnalyzer
    
    # Database integration disabled for standalone operation
    DATABASE_AVAILABLE = False
        
except ImportError as e:
    print(f"FATAL: A required library is not installed: {e.name}")
    print("Please run the following command from your terminal:")
    print(r"pip install -r requirements.txt")
    print("Or manually install: pip install PyQt6 pywin32 sqlalchemy psycopg2-binary")
    sys.exit(1)

# Database functionality disabled for standalone operation
def get_database_connection():
    """Database connection disabled - returns None for standalone operation"""
    return None

# =========================================================================================
# SECTION 1: CONFIGURATION
# =========================================================================================
APP_DIR = Path(__file__).resolve().parent
RULES_DEFAULT = APP_DIR / "backup_json" / "dwg_filename_rules.json"
TEMPLATES_DEFAULT = APP_DIR / "backup_json" / "templates.json"
RECIPES_CONFIG = APP_DIR / "backup_json" / "automation_recipes.json"
PRESETS_CONFIG = APP_DIR / "backup_json" / "project_presets.json"
DEFAULT_ROOT = Path(r"J:\J")
ARCHIVE_ROOT = Path(r"R:\J")
# REQUIRED_DESC_CODES now handled in config_manager.py

# --- Filename Sanitization Utilities (Prevent File Creation Failures) ---
import re as _re
_ILLEGAL_FILENAME_CHARS = r'<>:"/\\|?*'

def sanitize_filename(name: str) -> str:
    """Remove illegal characters from filename and clean up whitespace"""
    if not name:
        return "untitled"
    # Remove illegal characters
    clean_name = "".join(ch for ch in name if ch not in _ILLEGAL_FILENAME_CHARS)
    # Clean up multiple spaces and trim
    clean_name = _re.sub(r"\s+", " ", clean_name).strip(". ").strip()
    return clean_name or "untitled"

def unique_path(path: Path) -> Path:
    """Generate unique filename if path already exists (adds -02, -03, etc.)"""
    if not path.exists():
        return path
    
    stem, suffix = path.stem, path.suffix
    for i in range(2, 999):
        candidate = path.with_name(f"{stem}-{i:02d}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Too many collisions for {path}")

# --- AutoCAD/Civil 3D Specific Configuration ---

# --- AutoCAD executable discovery (2026-first, env overrides) ---
def _find_autocad_exe(exe_name: str, env_key: str, prefer_year: str = "2026") -> Path:
    from pathlib import Path as _P
    import os as _os
    p = _os.environ.get(env_key, "").strip()
    if p and _P(p).exists(): return _P(p)
    roots = [ _P(r"C:\Program Files\Autodesk"), _P(r"C:\Program Files (x86)\Autodesk") ]
    prefer_hit, last_hit = None, None
    try:
        for root in roots:
            if root.exists():
                for exe in root.rglob(exe_name):
                    if exe.is_file():
                        last_hit = exe
                        if prefer_year in str(exe):
                            prefer_hit = exe
                            break
                if prefer_hit: break
    except Exception: pass
    if prefer_hit and prefer_hit.exists(): return prefer_hit
    if last_hit and last_hit.exists(): return last_hit
    return _P(fr"C:\Program Files\Autodesk\AutoCAD {prefer_year}\{exe_name}")

ACAD_EXE = _find_autocad_exe("acad.exe", "ACAD_EXE", "2026")
ACCORECONSOLE_EXE = _find_autocad_exe("accoreconsole.exe", "ACCORECONSOLE_EXE", "2026")


CHOICE_LISTS = {
    "project_setup_config": ["School_Small", "School_Large", "BR_Plan", "BR_PlanProfile", "SR_Plan", "SR_PlanProfile"],
    "project_setup_tb_size": ["11x17", "22x34", "24x36", "30x42"],
    "project_setup_tb_type": ["BR", "EXHIBIT", "DSA", "QKA", "SR"],
    "project_status": ["SD", "DD", "CD"]
}

# =========================================================================================
# SECTION 2: HELPER FUNCTIONS
# =========================================================================================
def _last_segment(p_str: str) -> str:
    parts = re.split(r"[\\/]+", (p_str or "").strip())
    return parts[-1].strip() if parts else ""

def _expand_filename_pattern(pattern: str, mapping: dict) -> str:
    s = pattern or ""
    def _replace_vars(chunk: str) -> str:
        chunk = re.sub(r"<([A-Za-z_]+)>", lambda m: str(mapping.get(m.group(1), "")), chunk)
        chunk = re.sub(r"\[([A-Za-z_]+)\]", lambda m: str(mapping.get(m.group(1), "")), chunk)
        bare = r"\b(ProjectNumber|Subnumber|File_Type_Code|description|Phase)\b"
        chunk = re.sub(bare, lambda m: str(mapping.get(m.group(1), "")), chunk)
        return chunk
    def _group_replace(m):
        replaced = _replace_vars(m.group(1))
        if re.sub(r"[ \-_.]", "", replaced): return replaced
        return ""
    s = re.sub(r"\[([^\[\]]+)\]", _group_replace, s)
    s = _replace_vars(s)
    s = re.sub(r"\s+", " ", s).strip().replace(" -", "-").replace("- ", "-").replace("--", "-")
    return s

def resolve_script_path(recipe_data: dict, app_dir: Path) -> Path:
    raw = (recipe_data.get("script_file") or "").strip()
    if not raw: raise ValueError("Recipe is missing 'script_file' entry.")
    raw_expanded = os.path.expandvars(raw)
    cand = Path(raw_expanded)
    if cand.is_absolute() and cand.exists(): return cand
    cand2 = (app_dir / raw_expanded).resolve()
    if cand2.exists(): return cand2
    cand3 = (app_dir / "recipes" / raw_expanded).resolve()
    if cand3.exists(): return cand3
    for d in recipe_data.get("script_search", []):
        d_exp = os.path.expandvars(d)
        p = Path(d_exp) / Path(raw_expanded).name
        if p.exists(): return p
    raise FileNotFoundError(f"Could not resolve script path for: {raw}")

# Rule dataclass and load_rules_json now centralized in config_manager.py

def list_dwg_counts(root: Path, project: str, sub: str, rules: Dict[str, Rule]) -> Dict[str, int]:
    counts = {code: 0 for code in rules.keys()}
    if not project or not sub: return counts
    sub_base_path = root / project / "dwg" / f"{project} {sub}"
    if not sub_base_path.exists(): return counts
    possible_folders = {sub_base_path}.union({sub_base_path / r.folder_short for r in rules.values() if r.folder_short})
    for folder in possible_folders:
        if not folder.exists(): continue
        for dwg_path in folder.glob("*.dwg"):
            for code in rules.keys():
                if f" {code}" in dwg_path.name or f"-{code}" in dwg_path.name: counts[code] += 1
    return counts

# --- Viewport Helper Functions ---
def normalize_tb_size(s: Optional[str]) -> Optional[str]:
    if not s: return None
    s = s.strip().replace(" ", "").replace("×", "x").replace("X", "x")
    m = re.match(r"^(\d{2,3})x(\d{2,3})$", s, flags=re.I)
    return f"{m.group(1)}x{m.group(2)}" if m else s.lower()

def parse_layout_name(name: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    s = name.strip(); su = s.upper()
    m = re.match(r"^([A-Z]+)\d*([A-Z]{2,4})(\d{2,3}[Xx]\d{2,3})$", su)
    if m: return m.group(1), m.group(2), normalize_tb_size(m.group(3))
    m = re.match(r"^([A-Z]+)([A-Z]{2,4})(\d{2,3}[Xx]\d{2,3})$", su)
    if m: return m.group(1), m.group(2), normalize_tb_size(m.group(3))
    m = re.match(r"^([A-Za-z]+)\d*$", s)
    if m: return m.group(1).upper(), None, None
    return None, None, None

def find_project_db_path(dwg_path: Path) -> Optional[Path]:
    try:
        dwg_dir, sub_folder = dwg_path.parent, dwg_path.parent.name
        proj_dir, proj_num = dwg_dir.parent.parent, dwg_dir.parent.parent.name
        sub_num = sub_folder.split()[-1]
        db_path = proj_dir/"dwg"/f"{proj_num} {sub_num}"/"DESIGN"/"data"/f"{proj_num}.{sub_num}_Project_DB.json"
        return db_path if db_path.exists() else None
    except Exception: return None

def get_tb_from_project_db(db_path: Path) -> Tuple[Optional[str], Optional[str]]:
    try:
        data = json.loads(db_path.read_text(encoding="utf-8"))
        tb_type = str(data.get("project_setup_tb_type") or "").strip().upper() or None
        tb_size = normalize_tb_size(str(data.get("project_setup_tb_size") or ""))
        return tb_type, tb_size
    except Exception: return None, None

def dict_get_ci(d: dict, key: str):
    if key in d: return True, d[key]
    k_up = str(key).upper()
    for k in d.keys():
        if str(k).upper() == k_up: return True, d[k]
    return False, None

def get_tb_node(presets: dict, tb_type: str):
    ok, node = dict_get_ci(presets, tb_type or "")
    return node if ok else None

def get_size_node(tb_node: dict, tb_size: str):
    norm = normalize_tb_size(tb_size or "")
    if norm in tb_node: return tb_node[norm]
    for k in tb_node.keys():
        if normalize_tb_size(str(k)) == norm: return tb_node[k]
    return None

def get_layout_preset(size_node: dict, layout_code: str):
    ok, node = dict_get_ci(size_node, layout_code)
    if ok: return node
    for alt in ("LAYOUT", "PLAN", "COVER"):
        ok, node = dict_get_ci(size_node, alt)
        if ok: return node
    return None

def parse_scale_to_ratio(scale_str: str) -> Optional[float]:
    s = (scale_str or "").strip().replace(" ", "")
    m = re.match(r'^1"?=([\d\.]+)\'$', s)
    if not m: return None
    feet = float(m.group(1))
    return 1.0 / (feet * 12.0) if feet > 0 else None

# =========================================================================================
# SECTION 3: STABLE AUTOMATION BACKEND (THE "ENGINE")
# =========================================================================================
class AutomationEngine:
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        self.app: Optional[Any] = None  # AutoCAD COM Application object

    def log(self, message):
        if self.progress_callback: self.progress_callback(message + "\n")
        else: print(message)

    def APoint(self, x, y, z=0):
        return VARIANT(VT_ARRAY | VT_R8, (float(x), float(y), float(z)))

    def resolve_unc(self, p: Path) -> Path:
        p_abs = p.resolve()
        if p_abs.exists(): return p_abs
        if len(p.drive) == 2 and p.drive[1] == ":":
            drive_root = p.drive + "\\"
            try:
                unc_info = win32wnet.WNetGetUniversalName(drive_root)
                unc_root = unc_info.get('lpUniversalName', drive_root)
                candidate = Path(str(p).replace(drive_root, unc_root, 1))
                if candidate.exists(): return candidate
            except Exception: pass
        return p

    def wait_quiet(self, timeout: float = 240.0):
        if not self.app:
            return  # Gracefully handle case where AutoCAD app is not connected
        t0 = time.time()
        st = self.app.GetAcadState()
        while not st.IsQuiescent:
            if time.time() - t0 > timeout: raise TimeoutError("AutoCAD stayed busy for too long")
            time.sleep(0.25)
            st = self.app.GetAcadState()

    def activate_doc(self, doc):
        if not self.app or not doc:
            return  # Gracefully handle case where app or doc is None
        try:
            doc.Activate()
            self.app.Update()
            time.sleep(0.2)
            self.wait_quiet()
        except Exception: pass

    def wait_cmds_idle(self, doc, timeout=30.0):
        """Wait for all AutoCAD commands to complete - prevents race conditions"""
        if not self.app or not doc:
            return
        t0 = time.time()
        while time.time() - t0 < timeout:
            self.app.Update()
            try:
                # Check if any commands are active
                if doc.GetVariable("CMDACTIVE") == 0 and not doc.GetVariable("CMDNAMES"):
                    self.wait_quiet(2.0)  # Extra quiet time to be sure
                    return
            except Exception:
                pass  # Variable access might fail, continue trying
            time.sleep(0.1)
        raise TimeoutError("Commands stayed active too long - possible hang")

    def send_cmd(self, doc, cmd: str, retries: int = 60):
        if not cmd.endswith("\n"): cmd += "\n"
        for _ in range(retries):
            try:
                self.wait_quiet(timeout=15.0)
                self.activate_doc(doc)
                doc.SendCommand(cmd)
                return
            except (pywintypes.com_error, Exception): time.sleep(0.25)
        raise RuntimeError("SendCommand failed after multiple retries.")

    def open_dwg_robust(self, path: Path):
        if not self.app:
            raise RuntimeError("AutoCAD application is not connected")
        path_unc = self.resolve_unc(path)
        if not path_unc.exists(): raise FileNotFoundError(f"Drawing not found: {path_unc}")
        for i in range(30):
            try:
                self.wait_quiet(timeout=15.0)
                for d in self.app.Documents:
                    if Path(d.FullName).resolve() == path_unc.resolve():
                        self.activate_doc(d)
                        return d
                doc = self.app.Documents.Open(str(path_unc))
                self.activate_doc(doc)
                time.sleep(0.5)
                self.wait_quiet(timeout=60.0)
                return doc
            except (pywintypes.com_error, Exception):
                if i == 29: raise RuntimeError("Failed to open DWG after multiple retries.")
                time.sleep(0.5)

    def process_lisp_task(self, doc, lisp_path: Path, command_name: str):
        lisp_path_unc = self.resolve_unc(lisp_path)
        if not lisp_path_unc.exists():
            raise FileNotFoundError(f"LISP file not found: {lisp_path_unc}")
        self.log(f"  - Loading and running LISP: {lisp_path.name}")
        
        # Load LISP and wait for completion (prevents race conditions)
        self.send_cmd(doc, f'(load "{lisp_path_unc.as_posix()}" nil)')
        self.wait_cmds_idle(doc)  # Ensure LISP is fully loaded
        
        if command_name:
            self.send_cmd(doc, command_name)
            self.wait_cmds_idle(doc)  # Ensure command completes

    def process_coreconsole_task(self, dwg_paths: list[Path], script_path: Path):
        if not ACCORECONSOLE_EXE.exists(): raise FileNotFoundError(f"Core Console not found at {ACCORECONSOLE_EXE}")
        script_path_unc = self.resolve_unc(script_path)
        if not script_path_unc.exists(): raise FileNotFoundError(f"Script file (.scr) not found: {script_path_unc}")
        for dwg in dwg_paths:
            dwg_path_unc = self.resolve_unc(dwg)
            if not dwg_path_unc.exists():
                self.log(f"SKIP (missing): {dwg_path_unc}"); continue
            self.log(f"\n--- Core Console on: {dwg.name} ---")
            args = [str(ACCORECONSOLE_EXE), "/i", str(dwg_path_unc), "/s", str(script_path_unc), "/l", "en-US"]
            proc = subprocess.run(args, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            if proc.returncode == 0: self.log(f"✅ OK")
            else: self.log(f"  - FAILED. Output:\n{proc.stdout}\n{proc.stderr}")
    
    def process_viewport_task(self, doc, presets: dict):
        proj_tb_type, proj_tb_size = (None, None)
        db_path = find_project_db_path(Path(doc.FullName))
        if db_path:
            proj_tb_type, proj_tb_size = get_tb_from_project_db(db_path)
        layouts = [doc.Layouts.Item(i).Name for i in range(doc.Layouts.Count)]
        created_total = 0
        for lname in layouts:
            if lname.upper() == "MODEL": continue
            code, tb_type, tb_size = parse_layout_name(lname)
            tb_type = (tb_type or proj_tb_type or "").upper() or None
            tb_size = normalize_tb_size(tb_size or proj_tb_size or "")
            if code: 
                match = re.match(r"^([A-Z]+)", code)
                if match:
                    code = match.group(1).upper()
                else:
                    code = None
            if not (tb_type and tb_size and code):
                self.log(f"    - Skip '{lname}' (cannot resolve code/type/size)")
                continue
            tb_node = get_tb_node(presets, tb_type)
            if not tb_node: continue
            size_node = get_size_node(tb_node, tb_size)
            if not size_node: continue
            preset = get_layout_preset(size_node, code)
            if not preset:
                self.log(f"    - Skip '{lname}' (no layoutCode '{code}' preset for {tb_type}/{tb_size})")
                continue
            self.log(f"    * Creating viewports on layout: {lname}")
            doc.SetVariable("CTAB", lname)
            self.wait_quiet(5)
            time.sleep(1)
            for vpdef in preset.get("viewports", []):
                try:
                    self.wait_quiet(15)
                    layout_obj = doc.ActiveLayout
                    units = 'in' if layout_obj.PaperUnits == 1 else 'mm'
                    mult = 1.0 if units == 'in' else 25.4
                    center = [float(c) * mult for c in vpdef.get("center")]
                    width, height = float(vpdef.get("width"))*mult, float(vpdef.get("height"))*mult
                    if width <= 0 or height <= 0: continue
                    pspace = doc.PaperSpace
                    vp = pspace.AddPViewport(self.APoint(center[0], center[1]), width, height)
                    if scale_str := vpdef.get("scale"):
                        if cscale := parse_scale_to_ratio(scale_str):
                            vp.CustomScale = cscale
                    vp.DisplayLocked = bool(vpdef.get("lock", True))
                    if layer_name := vpdef.get("layer"):
                        try: doc.Layers.Item(layer_name)
                        except Exception: doc.Layers.Add(layer_name)
                        vp.Layer = layer_name
                    created_total += 1
                except Exception as e:
                    self.log(f"      - ERROR creating viewport: {e}")
        self.log(f"  - Viewports created: {created_total}.")

# =========================================================================================
# SECTION 4: QT WORKER THREAD (THE "BRIDGE") - REWRITTEN FOR STABILITY
# =========================================================================================
class AutomationWorker(QObject):
    progress = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, recipes_to_run: List[str], selected_dwgs: List[Path], recipes_config: dict):
        super().__init__()
        self.recipes_to_run = recipes_to_run
        self.selected_dwgs = selected_dwgs
        self.recipes_config = recipes_config
        self._cancel_requested = False  # Graceful cancel flag

    def log_message(self, message): self.progress.emit(message)
    
    def request_cancel(self):
        """Request graceful cancellation of the automation process"""
        self._cancel_requested = True
        self.log_message("🛑 Stop requested - cancelling after current operation...")

    def check_cancel(self):
        """Check if cancellation was requested"""
        if self._cancel_requested:
            raise InterruptedError("Operation cancelled by user")

    @pyqtSlot()
    def run(self):
        # Initialize COM apartment for thread safety (prevents crashes)
        import pythoncom
        pythoncom.CoInitialize()
        
        try:
            engine = AutomationEngine(progress_callback=self.log_message)
            
            # Try to connect to running AutoCAD, with fallback launch
            try:
                self.log_message("Attempting to connect to running AutoCAD session...\n")
                engine.app = EnsureDispatch("AutoCAD.Application")
                engine.app.Visible = True
                self.log_message(f"✅ Connection successful: {engine.app.Name} {engine.app.Version}\n")
            except Exception as connect_error:
                self.log_message(f"No running AutoCAD found. Attempting to launch...\n")
                try:
                    # Fallback: Launch AutoCAD then attach
                    import subprocess
                    subprocess.Popen([str(ACAD_EXE)], close_fds=True)
                    self.log_message("Waiting for AutoCAD to start...\n")
                    time.sleep(5)  # Give AutoCAD time to initialize
                    engine.app = EnsureDispatch("AutoCAD.Application")
                    engine.app.Visible = True
                    self.log_message(f"✅ Launch successful: {engine.app.Name} {engine.app.Version}\n")
                except Exception as launch_error:
                    self.log_message(f"❌ FATAL: Could not connect or launch AutoCAD.\n")
                    self.log_message(f"Connect Error: {connect_error}\n")
                    self.log_message(f"Launch Error: {launch_error}\n")
                    return

            # 1. Separate recipes by runner type
            core_console_recipes = []
            live_session_recipes = []
            for r_name in self.recipes_to_run:
                if self.recipes_config.get(r_name, {}).get("runner") == "core_console":
                    core_console_recipes.append(r_name)
                else:
                    live_session_recipes.append(r_name)

            # 2. Run all Core Console recipes first (if any)
            if core_console_recipes:
                self.log_message("\n--- Starting Core Console Batch Processing ---\n")
                for recipe_name in core_console_recipes:
                    try:
                        self.check_cancel()  # Check for cancel between recipes
                    except InterruptedError:
                        break  # Exit the recipe loop
                    
                    recipe_data = self.recipes_config.get(recipe_name, {})
                    self.log_message(f"  * Executing recipe: {recipe_name}")
                    try:
                        script = resolve_script_path(recipe_data, APP_DIR)
                        engine.process_coreconsole_task(self.selected_dwgs, script)
                    except Exception as e:
                        self.log_message(f"    - ERROR in recipe {recipe_name}: {e}\n{traceback.format_exc()}\n")
                
                if not self._cancel_requested:
                    self.log_message("\n--- Core Console Batch Processing Complete ---\n")

            # 3. Run all "live" recipes in a single session, processing one DWG at a time
            if live_session_recipes:
                self.log_message("\n--- Starting Live AutoCAD Session Processing ---\n")
                for dwg_path in self.selected_dwgs:
                    try:
                        self.check_cancel()  # Check for cancel between drawings
                    except InterruptedError:
                        break  # Exit the drawing loop
                    
                    doc = None
                    self.log_message(f"\n\n--- Processing Drawing: {dwg_path.name} ---")
                    try:
                        doc = engine.open_dwg_robust(dwg_path)
                    
                        # Apply all live recipes to this one open drawing
                        for recipe_name in live_session_recipes:
                            try:
                                self.check_cancel()  # Check for cancel between recipes
                            except InterruptedError:
                                raise  # Re-raise to be caught by outer handler
                            
                            self.log_message(f"\n  --- Applying recipe: {recipe_name} ---")
                            try:
                                recipe_data = self.recipes_config.get(recipe_name, {})
                                runner = recipe_data.get("runner")
                                
                                if runner == "pyautocad":
                                    script = resolve_script_path(recipe_data, APP_DIR)
                                    command = recipe_data.get("command", "")
                                    engine.process_lisp_task(doc, script, command)
                                elif runner == "python_direct":
                                    presets_file = APP_DIR / recipe_data.get("presets_file")
                                    # Use the config manager's preset loading method for consistency
                                    config_mgr = ConfigurationManager(APP_DIR)
                                    presets_data, error = config_mgr.load_preset_file(presets_file)
                                    if error:
                                        raise Exception(f"Failed to load preset file: {error}")
                                    engine.process_viewport_task(doc, presets_data)
                                else:
                                    self.log_message(f"    - Unknown live runner '{runner}'. Skipping.")
                            
                            except Exception as recipe_error:
                                # Log error for the specific recipe but continue with others on the same drawing
                                self.log_message(f"    - ❗ ERROR in recipe '{recipe_name}': {recipe_error}")
                                traceback.print_exc()

                        self.log_message(f"\n--- Finished Processing {dwg_path.name}. Saving... ---")
                        if doc:
                            doc.Save()
                        time.sleep(0.5) # Give AutoCAD a moment to finish writing the file before closing

                    except InterruptedError:
                        # User cancelled - break out of processing loop
                        if doc:
                            try:
                                doc.Close(False)
                            except Exception:
                                pass
                        break  # Exit the drawing loop
                    except Exception as e:
                        # Log a fatal error for this specific drawing and move to the next one
                        self.log_message(f"--- ❗ FATAL ERROR processing {dwg_path.name} ---\n{e}\n")
                        traceback.print_exc()
                    finally:
                        if doc and not self._cancel_requested:
                            try:
                                doc.Close(False)
                            except Exception:
                                self.log_message(f"    - Warning: Could not close {dwg_path.name} cleanly.")
                
                if not self._cancel_requested:
                    self.log_message("\n--- Live AutoCAD Session Processing Complete ---\n")

            if not self._cancel_requested:
                self.log_message("\n\n✅ All operations complete.\n")
        
        except InterruptedError as cancel_error:
            self.log_message(f"\n🛑 {cancel_error}")
        except Exception as e:
            self.log_message(f"\n❌ FATAL ERROR: {e}")
            traceback.print_exc()
        finally:
            # Clean up COM apartment
            pythoncom.CoUninitialize()
            self.finished.emit()

# =========================================================================================
# SECTION 5: PYQT6 UI (THE "DASHBOARD")
# =========================================================================================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DWG Project Orchestrator")
        self.resize(1800, 950)
        
        # Initialize the configuration manager
        self.config_manager = ConfigurationManager(APP_DIR)
        
        # Keep the same data attributes for backward compatibility
        self.rules: Dict[str, Rule] = {}
        self.recipes: Dict[str, Any] = {}
        self.recipes_categorized: Dict[str, Any] = {}  # Store categorized structure for UI
        self.presets: Dict[str, Any] = {}
        self.project_number = ""
        self.sub_number = ""
        self.root_dir = DEFAULT_ROOT
        self.rules_path: Path = RULES_DEFAULT
        self.templates_path: Path = TEMPLATES_DEFAULT
        self.templates: Dict[str, Any] = {}
        
        main_layout = QVBoxLayout(self)
        self.create_tab = CreateDrawingsTab(self)

        project_group = QGroupBox("Global Project Context")
        project_layout = QGridLayout(project_group)
        project_layout.addWidget(QLabel("Project:"), 0, 0)
        self.proj_edit = QLineEdit(); self.proj_edit.setPlaceholderText("e.g., 8888")
        project_layout.addWidget(self.proj_edit, 0, 1)
        project_layout.addWidget(QLabel("Sub:"), 0, 2)
        self.sub_edit = QLineEdit(); self.sub_edit.setPlaceholderText("e.g., 09")
        project_layout.addWidget(self.sub_edit, 0, 3)
        self.load_proj_btn = QPushButton("Load Project")
        self.load_proj_btn.clicked.connect(self.load_project)
        project_layout.addWidget(self.load_proj_btn, 0, 4)
        project_layout.setColumnStretch(5, 1)
        main_layout.addWidget(project_group)

        preset_layout = QHBoxLayout()
        preset_layout.addWidget(QLabel("<b>Preset:</b>"))
        self.preset_combo = QComboBox()
        self.preset_combo.setMinimumWidth(250)
        preset_layout.addWidget(self.preset_combo)
        self.load_preset_btn = QPushButton("Load Preset")
        self.load_preset_btn.clicked.connect(self.create_tab.on_load_preset)
        preset_layout.addWidget(self.load_preset_btn)
        preset_layout.addStretch(1)
        main_layout.addLayout(preset_layout)

        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tabs = QTabWidget()
        self.automation_tab = AutomationHubTab(self)
        self.tabs.addTab(self.automation_tab, "🚀 Automation Hub")
        self.tabs.addTab(self.create_tab, "➕ Create Drawings")
        self.tabs.addTab(HealthCheckTab(), "🩺 Project Health Check")
        self.tabs.addTab(SheetSetTab(), "📑 Sheet Set Manager")
        self.tabs.addTab(DrawingAnalysisTab(), "📊 Drawing Analysis")
        self.tabs.addTab(DXFAnalysisTab(), "🔍 DXF Analysis")
        self.tabs.addTab(LayerManagerTab(), "🎨 Layer Manager")
        self.tabs.addTab(BlockLibraryTab(), "🔧 Block Library")
        self.tabs.addTab(StandardsDashboardTab(), "📋 Standards Dashboard")
        self.tabs.addTab(ProjectReportsTab(), "📈 Project Reports")
        self.tabs.addTab(BatchOperationsTab(), "⚡ Batch Operations")
        main_splitter.addWidget(self.tabs)
        self.db_panel = ProjectDatabasePanel(self)
        main_splitter.addWidget(self.db_panel)
        main_splitter.setSizes([1200, 600])
        main_layout.addWidget(main_splitter)
        main_layout.setStretchFactor(project_group, 0)
        main_layout.setStretchFactor(main_splitter, 1)

        self.load_recipes()
        self.load_presets()
        self.load_rules()
        self.load_templates()

    def load_recipes(self):
        """Load recipes using the configuration manager (same interface as before)."""
        self.recipes, self.recipes_categorized, error = self.config_manager.load_recipes()
        
        if error:
            if "not found" in error:
                QMessageBox.warning(self, "File Not Found", error)
            else:
                QMessageBox.critical(self, "Recipes Error", error)
        else:
            self.automation_tab.populate_recipes()
    
    def load_presets(self):
        """Load presets using the configuration manager (same interface as before)."""
        self.presets, error = self.config_manager.load_presets()
        
        if error:
            if "not found" in error:
                QMessageBox.warning(self, "File Not Found", error)
            else:
                QMessageBox.critical(self, "Presets Error", error)
        else:
            self.populate_presets()

    def populate_presets(self):
        self.preset_combo.clear()
        if not self.presets:
            self.preset_combo.addItem("presets.json not found or empty.")
            return
        self.preset_combo.addItem("-- Select a Preset --")
        for name, data in self.presets.items():
            self.preset_combo.addItem(name, userData=data)
        self.preset_combo.currentIndexChanged.connect(
            lambda index: self.preset_combo.setToolTip(
                self.preset_combo.itemData(index).get("description", "") if index > 0 else ""))

    def load_rules(self):
        """Load rules using the configuration manager (same interface as before)."""
        self.rules, error = self.config_manager.load_rules(self.rules_path)
        
        if error:
            if "not found" in error:
                QMessageBox.warning(self, "File Not Found", error)
            else:
                QMessageBox.critical(self, "Rules Error", error)
        else:
            self.create_tab.rebuild_tree()

    def load_templates(self):
        """Load templates using the configuration manager (same interface as before)."""
        self.templates, error = self.config_manager.load_templates(self.templates_path)
        
        if error:
            QMessageBox.warning(self, "Templates Error" if "not found" not in error else "File Not Found", error)
            if "not found" not in error:
                self.templates = {}

    def load_project(self):
        self.project_number = self.proj_edit.text().strip()
        self.sub_number = self.sub_edit.text().strip()
        if not self.project_number or not self.sub_number:
            QMessageBox.warning(self, "Input Required", "Project and Sub Number are required.")
            return

        # --- ARCHIVE CHECK LOGIC ---
        project_path = self.root_dir / self.project_number
        archive_path = ARCHIVE_ROOT / self.project_number
        if not project_path.exists() and archive_path.exists():
            QMessageBox.warning(self, "Archived Project",
                                f"Project {self.project_number} exists in the archive (R: drive).\n\n"
                                "Please restore the project from the archive before proceeding.")
            return # This stops the function from continuing
        # --- END OF ARCHIVE CHECK ---

        self.automation_tab.refresh_dwg_list()
        self.create_tab.rebuild_tree()
        self.db_panel.display_project_db_info()
        self.tabs.setCurrentWidget(self.automation_tab)

    def get_project_root_path(self) -> Optional[Path]:
        return self.root_dir / self.project_number if self.project_number else None
    
    def get_sub_path(self) -> Optional[Path]:
        proj_root = self.get_project_root_path()
        return proj_root / "dwg" / f"{self.project_number} {self.sub_number}" if proj_root and self.sub_number else None
    
    def get_target_dir(self, folder_short: str) -> Optional[Path]:
        sub_path = self.get_sub_path()
        return sub_path / folder_short if sub_path and folder_short else sub_path
    
    def get_project_db_path(self) -> Optional[Path]:
        sub_path = self.get_sub_path()
        return sub_path / "DESIGN" / "data" / f"{self.project_number}.{self.sub_number}_Project_DB.json" if sub_path else None
    
    def ensure_standard_folders(self) -> bool:
        if not self.project_number or not self.sub_number:
            return False
        base_proj = self.get_project_root_path()
        sub_path = self.get_sub_path()
        if not base_proj or not sub_path:
            return False

        project_was_new = not base_proj.exists()
        subfolder_was_new = not sub_path.exists()

        try:
            dwg_path = base_proj / "dwg"
            for p in (base_proj, dwg_path, sub_path):
                p.mkdir(parents=True, exist_ok=True)
            for name in ("ARCH", "BR", "DESIGN", "EMAIL", "ENG", "EXHIBIT", "OBJECT", "VEHICLE TRACKING"):
                (sub_path / name).mkdir(parents=True, exist_ok=True)

            sc_path = dwg_path / "DATA SHORTCUTS" / "_Shortcuts"
            sc_path.mkdir(parents=True, exist_ok=True)
            for name in ("Alignments", "Corridors", "PipeNetworks", "Profiles", "Surfaces",
                         "PressurePipeNetworks", "SampleLineGroups", "ViewframeGroups"):
                (sc_path / name).mkdir(parents=True, exist_ok=True)

            if project_was_new:
                for name in ("GIS", "map", "survey"):
                    (base_proj / name).mkdir(exist_ok=True)
                (base_proj / "survey" / "CS" / "Construction Control").mkdir(parents=True, exist_ok=True)
                (base_proj / "survey" / "field data").mkdir(parents=True, exist_ok=True)

            if subfolder_was_new:
                self.create_default_project_db()
            return True
        except Exception as e:
            QMessageBox.critical(self, "Folder Creation Error", f"Could not create standard folders:\n{e}")
            return False
    
    def create_default_project_db(self):
        db_path = self.get_project_db_path()
        if not db_path or db_path.exists(): return
        default_data = {
            "project_number": self.project_number, "project_subnumber": self.sub_number, "project_name_long": "",
            "project_date": "", "project_setup_config": "", "project_setup_tb_size": "", "project_setup_tb_type": "",
            "project_manager": "", "lead_designer": "", "client_name": "", "project_status": "",
            "coordinate_system": "", "vertical_datum": ""
        }
        try:
            db_path.parent.mkdir(parents=True, exist_ok=True)
            db_path.write_text(json.dumps(default_data, indent=2), encoding="utf-8")
        except Exception as e:
            QMessageBox.warning(self, "DB Creation Failed", f"Could not create the default Project DB JSON file:\n{e}")
    
    def resolve_dwt_for_folder(self, folder_short: str) -> Optional[Path]:
        per_folder = self.templates.get("per_folder", {}); path_str = per_folder.get(folder_short) or self.templates.get("default", "")
        if path_str and Path(path_str).exists():
            return Path(path_str)
        return None

class ProjectDatabasePanel(QGroupBox):
    def __init__(self, main_window: MainWindow):
        super().__init__("Project Database Information")
        self.mw = main_window
        layout = QVBoxLayout(self)
        self.db_table = QTableWidget(0, 2)
        self.db_table.setHorizontalHeaderLabels(["Property", "Value"])
        self.db_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.db_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.db_table)
        self.save_db_btn = QPushButton("Save DB Changes")
        self.save_db_btn.clicked.connect(self.save_project_db_info)
        layout.addWidget(self.save_db_btn)
        
    def display_project_db_info(self):
        self.db_table.setRowCount(0)
        db_path = self.mw.get_project_db_path()
        if db_path and db_path.exists():
            try:
                data = json.loads(db_path.read_text(encoding="utf-8"))
                self.db_table.setRowCount(len(data))
                for row, (key, value) in enumerate(data.items()):
                    key_item = QTableWidgetItem(str(key))
                    key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.db_table.setItem(row, 0, key_item)
                    if key in CHOICE_LISTS:
                        combo = QComboBox()
                        combo.addItems(CHOICE_LISTS[key])
                        try: combo.setCurrentIndex(CHOICE_LISTS[key].index(str(value)))
                        except ValueError:
                            if CHOICE_LISTS[key]: combo.setCurrentIndex(0)
                        self.db_table.setCellWidget(row, 1, combo)
                    else:
                        value_item = QTableWidgetItem(str(value))
                        if key in ("project_number", "project_subnumber"):
                            value_item.setFlags(value_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        self.db_table.setItem(row, 1, value_item)
            except Exception as e:
                self.db_table.setRowCount(1)
                self.db_table.setItem(0, 0, QTableWidgetItem(f"Error reading JSON file: {e}"))
                self.db_table.setSpan(0, 0, 1, 2)
        else:
            self.db_table.setRowCount(1)
            not_found_item = QTableWidgetItem("Project DB JSON not found. Create a drawing to generate one.")
            not_found_item.setFont(QFont("Segoe UI", 9, italic=True)); not_found_item.setForeground(QColor("gray"))
            self.db_table.setItem(0, 0, not_found_item)
            self.db_table.setSpan(0, 0, 1, 2)
            
    def save_project_db_info(self):
        db_path = self.mw.get_project_db_path()
        if not db_path or not db_path.exists(): return
        updated_data = {}
        for row in range(self.db_table.rowCount()):
            key_item = self.db_table.item(row, 0)
            if not key_item: continue
            key, value = key_item.text(), ""
            widget = self.db_table.cellWidget(row, 1)
            if isinstance(widget, QComboBox): value = widget.currentText()
            else:
                value_item = self.db_table.item(row, 1)
                if value_item: value = value_item.text()
            updated_data[key] = value
        try:
            db_path.write_text(json.dumps(updated_data, indent=2), encoding="utf-8")
            QMessageBox.information(self, "Success", "Project DB changes saved successfully.")
        except Exception as e: QMessageBox.critical(self, "Save Failed", f"Failed to save Project DB JSON file:\n{e}")

class AutomationHubTab(QWidget):
    def __init__(self, main_window: MainWindow):
        super().__init__()
        self.mw = main_window
        self.worker_thread = None
        self.worker = None
        layout = QHBoxLayout(self)
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        dwg_group = QGroupBox("1. Select Target Drawings")
        dwg_layout = QVBoxLayout(dwg_group)
        self.dwg_tree = QTreeWidget()
        self.dwg_tree.setHeaderLabel("DWG Files")
        self.dwg_tree.setRootIsDecorated(False)  # No expand/collapse arrows for single level
        dwg_layout.addWidget(self.dwg_tree)
        main_splitter.addWidget(dwg_group)
        recipes_group = QGroupBox("2. Available Recipes")
        recipes_layout = QVBoxLayout(recipes_group)
        self.recipe_tree = QTreeWidget()
        self.recipe_tree.setHeaderLabel("Recipes by Category")
        self.recipe_tree.setRootIsDecorated(True)
        # Add double-click functionality to add recipes to queue
        self.recipe_tree.itemDoubleClicked.connect(self.on_recipe_double_clicked)
        recipes_layout.addWidget(self.recipe_tree)
        self.add_to_queue_btn = QPushButton("Add to Execution Queue ➡")
        self.add_to_queue_btn.clicked.connect(self.add_selected_recipes_to_queue)
        recipes_layout.addWidget(self.add_to_queue_btn)
        main_splitter.addWidget(recipes_group)
        exec_group = QGroupBox("3. Execution Queue & Output")
        exec_layout = QVBoxLayout(exec_group)
        exec_layout.addWidget(QLabel("Execution Order:"))
        self.queue_list = QListWidget()
        exec_layout.addWidget(self.queue_list)
        self.remove_from_queue_btn = QPushButton("⬅ Remove Selected from Queue")
        self.remove_from_queue_btn.clicked.connect(self.remove_selected_recipe_from_queue)
        exec_layout.addWidget(self.remove_from_queue_btn)
        # Button layout for Run and Stop
        button_layout = QHBoxLayout()
        self.run_btn = QPushButton("RUN SEQUENCE 🚀")
        self.run_btn.setStyleSheet("font-size: 14px; padding: 8px;")
        self.run_btn.clicked.connect(self.on_run_sequence)
        button_layout.addWidget(self.run_btn)
        
        self.stop_btn = QPushButton("⏹ STOP")
        self.stop_btn.setStyleSheet("font-size: 14px; padding: 8px; background-color: #d32f2f; color: white;")
        self.stop_btn.clicked.connect(self.on_stop_sequence)
        self.stop_btn.setEnabled(False)  # Disabled until running
        button_layout.addWidget(self.stop_btn)
        
        exec_layout.addLayout(button_layout)
        self.output_log = QPlainTextEdit()
        self.output_log.setReadOnly(True)
        self.output_log.setFont(QFont("Courier New"))
        exec_layout.addWidget(self.output_log)
        main_splitter.addWidget(exec_group)
        main_splitter.setSizes([300, 300, 600])
        layout.addWidget(main_splitter)

    def populate_recipes(self):
        self.recipe_tree.clear()
        if not self.mw.recipes_categorized: return
        
        # Create category nodes and recipe children
        for category_name, category_data in self.mw.recipes_categorized.items():
            category_item = QTreeWidgetItem([category_name])
            category_item.setToolTip(0, category_data.get("description", ""))
            # Mark as category (not selectable for adding to queue)
            category_item.setData(0, Qt.ItemDataRole.UserRole, {"type": "category"})
            
            # Add recipe children
            if "recipes" in category_data:
                for recipe_name, recipe_data in category_data["recipes"].items():
                    recipe_item = QTreeWidgetItem([recipe_name])
                    recipe_item.setToolTip(0, recipe_data.get("description", "No description available."))
                    # Mark as recipe and store the recipe name
                    recipe_item.setData(0, Qt.ItemDataRole.UserRole, {"type": "recipe", "name": recipe_name})
                    category_item.addChild(recipe_item)
            
            self.recipe_tree.addTopLevelItem(category_item)
        
        # Expand all categories by default
        self.recipe_tree.expandAll()

    def refresh_dwg_list(self):
        self.dwg_tree.clear()
        sub_path = self.mw.get_sub_path()
        if not sub_path or not sub_path.exists(): return
        dwgs = sorted(sub_path.glob("**/*.dwg"))
        for dwg in dwgs:
            item = QTreeWidgetItem([str(dwg.relative_to(sub_path))])
            item.setCheckState(0, Qt.CheckState.Unchecked)  # Add checkbox
            item.setData(0, Qt.ItemDataRole.UserRole, dwg)  # Store the path
            self.dwg_tree.addTopLevelItem(item)

    def add_selected_recipes_to_queue(self):
        selected_items = self.recipe_tree.selectedItems()
        for item in selected_items:
            # Only add recipe items to queue, not categories
            item_data = item.data(0, Qt.ItemDataRole.UserRole)
            if item_data and item_data.get("type") == "recipe":
                recipe_name = item_data.get("name", item.text(0))
                self.queue_list.addItem(recipe_name)
    
    def on_recipe_double_clicked(self, item, column):
        """Handle double-click on recipe tree to add recipe to queue"""
        # Only add recipe items to queue, not categories
        item_data = item.data(0, Qt.ItemDataRole.UserRole)
        if item_data and item_data.get("type") == "recipe":
            recipe_name = item_data.get("name", item.text(0))
            self.queue_list.addItem(recipe_name)

    def remove_selected_recipe_from_queue(self):
        for item in self.queue_list.selectedItems():
            self.queue_list.takeItem(self.queue_list.row(item))

    def on_run_sequence(self):
        # Collect checked drawings instead of selected ones
        selected_dwgs = []
        for i in range(self.dwg_tree.topLevelItemCount()):
            item = self.dwg_tree.topLevelItem(i)
            if item.checkState(0) == Qt.CheckState.Checked:
                dwg_path = item.data(0, Qt.ItemDataRole.UserRole)
                if dwg_path:
                    selected_dwgs.append(dwg_path)
        
        recipes_to_run = [self.queue_list.item(i).text() for i in range(self.queue_list.count())]

        if not selected_dwgs or not recipes_to_run:
            QMessageBox.warning(self, "Input Required", "Please select drawings and add recipes to the queue.")
            return

        self.output_log.clear()
        self.run_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)  # Enable stop button when running

        self.worker_thread = QThread()
        self.worker = AutomationWorker(recipes_to_run, selected_dwgs, self.mw.recipes)
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker.progress.connect(self.output_log.insertPlainText)
        self.worker_thread.finished.connect(lambda: [self.run_btn.setEnabled(True), self.stop_btn.setEnabled(False)])
        
        self.worker_thread.start()

    def on_stop_sequence(self):
        """Handle the Stop button click - graceful cancellation"""
        if self.worker:
            self.worker.request_cancel()
            self.stop_btn.setEnabled(False)  # Prevent multiple stops

class CreateDrawingsTab(QWidget):
    def __init__(self, main_window: MainWindow):
        super().__init__()
        self.mw = main_window
        self.required_field_color = QColor(255, 240, 240)
        layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        code_group = QGroupBox("Select File Type Codes")
        code_layout = QVBoxLayout(code_group)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Code", "Existing"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.itemChanged.connect(self.on_selection_changed)
        code_layout.addWidget(self.tree)
        splitter.addWidget(code_group)
        options_group = QGroupBox("Set Options for New Drawings")
        options_layout = QVBoxLayout(options_group)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Code", "Phase", "Description", "Filename Preview"])
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.itemChanged.connect(self.update_previews)
        options_layout.addWidget(self.table)
        instance_button_layout = QHBoxLayout()
        self.add_instance_btn = QPushButton("Add Another Instance")
        self.add_instance_btn.clicked.connect(self.add_instance)
        self.remove_instance_btn = QPushButton("Remove Selected Instance")
        self.remove_instance_btn.clicked.connect(self.remove_instance)
        instance_button_layout.addWidget(self.add_instance_btn)
        instance_button_layout.addWidget(self.remove_instance_btn)
        options_layout.addLayout(instance_button_layout)
        self.create_btn = QPushButton("Create Selected Drawings")
        self.create_btn.clicked.connect(self.run_create)
        options_layout.addWidget(self.create_btn)
        splitter.addWidget(options_group)
        layout.addWidget(splitter)
    
    def on_load_preset(self):
        preset_combo = self.mw.preset_combo 
        current_index = preset_combo.currentIndex()
        if current_index <= 0: return
        preset_name = preset_combo.currentText()
        preset_data = self.mw.presets.get(preset_name)
        if not preset_data or "drawings" not in preset_data: return
        if self.table.rowCount() > 0:
            reply = QMessageBox.question(self, 'Confirm Overwrite',
                "This will clear your current selections. Are you sure?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No: return
        self.tree.blockSignals(True)
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            if iterator.value().parent():
                iterator.value().setCheckState(0, Qt.CheckState.Unchecked)
            iterator += 1
        for dwg_info in preset_data["drawings"]:
            code = dwg_info.get("code")
            if not code: continue
            found_item = self.find_tree_item_by_code(code)
            if found_item:
                found_item.setCheckState(0, Qt.CheckState.Checked)
                row_pos = self.table.rowCount()
                self.table.insertRow(row_pos)
                code_item = QTableWidgetItem(code)
                code_item.setFlags(code_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_pos, 0, code_item)
                self.table.setItem(row_pos, 1, QTableWidgetItem(""))
                desc_item = QTableWidgetItem(dwg_info.get("description", ""))
                self.table.setItem(row_pos, 2, desc_item)
                self.table.setItem(row_pos, 3, QTableWidgetItem(""))
                rule = self.mw.rules.get(code)
                if rule and rule.Description_Required:
                    desc_item.setBackground(self.required_field_color)
        self.tree.blockSignals(False)
        self.table.blockSignals(False)
        self.update_previews()

    def find_tree_item_by_code(self, code: str) -> Optional[QTreeWidgetItem]:
        iterator = QTreeWidgetItemIterator(self.tree)
        while iterator.value():
            item = iterator.value()
            if item.text(0) == code and item.parent(): return item
            iterator += 1
        return None

    def rebuild_tree(self):
        self.tree.clear(); self.table.setRowCount(0)
        if not self.mw.rules: return
        counts = list_dwg_counts(self.mw.root_dir, self.mw.project_number, self.mw.sub_number, self.mw.rules)
        grouped_rules: Dict[str, List[str]] = {}
        for code, rule in self.mw.rules.items():
            folder_key = rule.folder_short or "(Base Drawings)"
            if folder_key not in grouped_rules: grouped_rules[folder_key] = []
            grouped_rules[folder_key].append(code)
        for folder_name in sorted(grouped_rules.keys()):
            parent_item = QTreeWidgetItem([folder_name])
            font = parent_item.font(0); font.setBold(True); parent_item.setFont(0, font)
            parent_item.setFlags(parent_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable & ~Qt.ItemFlag.ItemIsSelectable)
            self.tree.addTopLevelItem(parent_item)
            for code in sorted(grouped_rules[folder_name]):
                child_item = QTreeWidgetItem([code, str(counts.get(code, 0))])
                child_item.setFlags(child_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                child_item.setCheckState(0, Qt.CheckState.Unchecked)
                parent_item.addChild(child_item)
            parent_item.setExpanded(True)
            
    def on_selection_changed(self, item: QTreeWidgetItem, column: int):
        if not item.parent() or column != 0: return
        code = item.text(0)
        self.table.blockSignals(True)
        if item.checkState(0) == Qt.CheckState.Checked:
            is_present = any(self.table.item(r, 0).text() == code for r in range(self.table.rowCount()))
            if not is_present:
                row_pos = self.table.rowCount()
                self.table.insertRow(row_pos)
                code_item = QTableWidgetItem(code); code_item.setFlags(code_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_pos, 0, code_item)
                self.table.setItem(row_pos, 1, QTableWidgetItem(""))
                desc_item = QTableWidgetItem(""); self.table.setItem(row_pos, 2, desc_item)
                self.table.setItem(row_pos, 3, QTableWidgetItem(""))
                rule = self.mw.rules.get(code)
                if rule and rule.Description_Required: desc_item.setBackground(self.required_field_color)
        else:
            for r in range(self.table.rowCount() - 1, -1, -1):
                if self.table.item(r, 0).text() == code: self.table.removeRow(r)
        self.table.blockSignals(False); self.update_previews()

    def add_instance(self):
        selected = self.table.selectedItems()
        if not selected: return
        source_row = selected[0].row(); code = self.table.item(source_row, 0).text(); rule = self.mw.rules.get(code)
        if not rule or not rule.Multi_Instance_Allowed: return
        self.table.blockSignals(True)
        new_row = source_row + 1; self.table.insertRow(new_row)
        code_item = QTableWidgetItem(code); code_item.setFlags(code_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.table.setItem(new_row, 0, code_item)
        self.table.setItem(new_row, 1, QTableWidgetItem(""))
        desc_item = QTableWidgetItem(""); self.table.setItem(new_row, 2, desc_item)
        self.table.setItem(new_row, 3, QTableWidgetItem(""))
        if rule.Description_Required: desc_item.setBackground(self.required_field_color)
        self.table.blockSignals(False); self.update_previews()
        
    def remove_instance(self):
        selected = self.table.selectedItems()
        if not selected: return
        row_to_remove = selected[0].row(); code = self.table.item(row_to_remove, 0).text()
        count = sum(1 for r in range(self.table.rowCount()) if self.table.item(r, 0).text() == code)
        if count <= 1: return
        self.table.removeRow(row_to_remove); self.update_previews()
        
    def update_previews(self, *_):
        for row in range(self.table.rowCount()):
            code = self.table.item(row, 0).text()
            phase = self.table.item(row, 1).text().strip() if self.table.item(row, 1) else ""
            desc = self.table.item(row, 2).text().strip() if self.table.item(row, 2) else ""
            rule = self.mw.rules.get(code)
            if not rule: continue
            mapping = {"ProjectNumber": self.mw.project_number, "Subnumber": self.mw.sub_number, "Phase": phase, "File_Type_Code": code, "description": desc}
            preview_item = self.table.item(row, 3)
            if preview_item: preview_item.setText(_expand_filename_pattern(rule.filename_pattern or "", mapping))
            else: self.table.setItem(row, 3, QTableWidgetItem(_expand_filename_pattern(rule.filename_pattern or "", mapping)))
            
    def run_create(self):
        if not self.mw.project_number or not self.mw.sub_number:
            QMessageBox.warning(self, "Input Required", "Please load a Project and Sub number first.")
            return
            
        errors = [f"Description required for '{self.table.item(r,0).text()}'" for r in range(self.table.rowCount())
                  if ((rule := self.mw.rules.get(self.table.item(r,0).text())) and rule.Description_Required) and not self.table.item(r,2).text().strip()]
        if errors:
            QMessageBox.warning(self, "Missing Information", "\n- ".join(["Please provide the following required information:"] + errors))
            return
        if not self.mw.ensure_standard_folders(): return
        created, error_list = 0, []
        for row in range(self.table.rowCount()):
            code = self.table.item(row, 0).text()
            filename = self.table.item(row, 3).text()
            rule = self.mw.rules.get(code)
            if not rule: continue
            target_dir = self.mw.get_target_dir(rule.folder_short)
            if not target_dir: continue
 
            if not filename.lower().endswith('.dwg'):
                filename += ".dwg"
            target_path = target_dir / filename
 
            dwt = self.mw.resolve_dwt_for_folder(rule.folder_short)
            if target_path.exists():
                error_list.append(f"{filename}: File already exists.")
                continue
            if not dwt:
                error_list.append(f"{filename}: Template (.dwt) not found for folder '{rule.folder_short}'.")
                continue
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
                target_path.write_bytes(dwt.read_bytes())
                created += 1
            except Exception as e: error_list.append(f"{filename}: Failed to create - {e}")
        if created > 0:
            self.mw.automation_tab.refresh_dwg_list()
            self.rebuild_tree()
            self.mw.tabs.setCurrentWidget(self.mw.automation_tab)
            self.mw.db_panel.display_project_db_info()
        msg = f"Successfully created {created} drawing(s)."
        if error_list:
            msg += "\n\nThe following errors occurred:\n- " + "\n- ".join(error_list)
        QMessageBox.information(self, "Creation Complete", msg)
                
# =========================================================================================
# SECTION 6: ADVANCED DEMO PLACEHOLDER TABS (FOR PROFESSIONAL PRESENTATION)
# =========================================================================================

class DrawingAnalysisTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        # Top control panel
        control_panel = QHBoxLayout()
        control_panel.addWidget(QLabel("<b>Project Analysis Dashboard:</b>"))
        self.analyze_btn = QPushButton("🔍 Analyze Selected Drawings")
        self.analyze_btn.clicked.connect(lambda: QMessageBox.information(self, "Analysis", "Drawing analysis would scan for file sizes, object counts, layer usage, block definitions, and performance metrics."))
        control_panel.addWidget(self.analyze_btn)
        control_panel.addStretch()
        layout.addLayout(control_panel)
        
        # Statistics dashboard
        stats_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: File statistics
        file_stats = QGroupBox("📁 File Statistics")
        file_layout = QVBoxLayout(file_stats)
        self.file_table = QTableWidget(5, 3)
        self.file_table.setHorizontalHeaderLabels(["Drawing", "Size (MB)", "Objects"])
        self.file_table.setItem(0, 0, QTableWidgetItem("Plan.dwg"))
        self.file_table.setItem(0, 1, QTableWidgetItem("2.4"))
        self.file_table.setItem(0, 2, QTableWidgetItem("1,247"))
        self.file_table.setItem(1, 0, QTableWidgetItem("Profile.dwg"))
        self.file_table.setItem(1, 1, QTableWidgetItem("1.8"))
        self.file_table.setItem(1, 2, QTableWidgetItem("892"))
        file_layout.addWidget(self.file_table)
        stats_splitter.addWidget(file_stats)
        
        # Right: Performance metrics
        perf_group = QGroupBox("⚡ Performance Metrics")
        perf_layout = QVBoxLayout(perf_group)
        perf_layout.addWidget(QLabel("Drawing Health Score: <b>87/100</b>"))
        perf_layout.addWidget(QLabel("Optimization Potential: <b>Medium</b>"))
        perf_layout.addWidget(QLabel("Largest Objects: Hatch patterns (34%)"))
        perf_layout.addWidget(QLabel("Load Time Estimate: <b>3.2 seconds</b>"))
        perf_layout.addWidget(QPushButton("📊 Generate Detailed Report"))
        perf_layout.addStretch()
        stats_splitter.addWidget(perf_group)
        
        layout.addWidget(stats_splitter)

class LayerManagerTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        self.layer_standards = []
        self.all_disciplines = set()
        self.all_statuses = set()
        self.all_categories = set()
        
        self.load_layer_standards()
        
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("<b>Filters:</b>"))
        
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search by layer name or description...")
        self.search_box.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.search_box, 2)
        
        filter_layout.addWidget(QLabel("Discipline:"))
        self.discipline_combo = QComboBox()
        self.discipline_combo.addItem("All")
        self.discipline_combo.addItems(sorted(self.all_disciplines))
        self.discipline_combo.currentTextChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.discipline_combo, 1)
        
        filter_layout.addWidget(QLabel("Status:"))
        self.status_combo = QComboBox()
        self.status_combo.addItem("All")
        self.status_combo.addItems(sorted(self.all_statuses))
        self.status_combo.currentTextChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.status_combo, 1)
        
        filter_layout.addWidget(QLabel("Category:"))
        self.category_combo = QComboBox()
        self.category_combo.addItem("All")
        self.category_combo.addItems(sorted(self.all_categories))
        self.category_combo.currentTextChanged.connect(self.apply_filters)
        filter_layout.addWidget(self.category_combo, 1)
        
        layout.addLayout(filter_layout)
        
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        table_widget = QWidget()
        table_layout = QVBoxLayout(table_widget)
        table_layout.setContentsMargins(0, 0, 0, 0)
        
        self.layer_table = QTableWidget()
        self.layer_table.setColumnCount(7)
        self.layer_table.setHorizontalHeaderLabels([
            "Layer Name", "Color", "Linetype", "Status", "Category", "Discipline", "Plottable"
        ])
        self.layer_table.setSortingEnabled(True)
        self.layer_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.layer_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.layer_table.horizontalHeader().setStretchLastSection(False)
        self.layer_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col in range(1, 7):
            self.layer_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        
        self.layer_table.itemSelectionChanged.connect(self.on_layer_selected)
        
        table_layout.addWidget(self.layer_table)
        main_splitter.addWidget(table_widget)
        
        details_group = QGroupBox("Layer Details")
        details_layout = QVBoxLayout(details_group)
        
        self.details_text = QTextEdit()
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumWidth(400)
        details_layout.addWidget(self.details_text)
        
        main_splitter.addWidget(details_group)
        main_splitter.setStretchFactor(0, 3)
        main_splitter.setStretchFactor(1, 1)
        
        layout.addWidget(main_splitter)
        
        self.populate_table()
    
    def load_layer_standards(self):
        json_path = APP_DIR / "backup_json" / "layer_standards.json"
        try:
            if not json_path.exists():
                QMessageBox.warning(
                    self, 
                    "Layer Standards Not Found", 
                    f"Could not find layer_standards.json at:\n{json_path}\n\nNo layer data will be displayed."
                )
                return
            
            with open(json_path, 'r', encoding='utf-8') as f:
                self.layer_standards = json.load(f)
            
            for layer in self.layer_standards:
                discipline = layer.get('discipline', 'Unknown')
                status = layer.get('status', 'Unknown')
                category = layer.get('category', 'Unknown')
                
                if discipline:
                    self.all_disciplines.add(discipline)
                if status:
                    self.all_statuses.add(status)
                if category:
                    self.all_categories.add(category)
                    
        except json.JSONDecodeError as e:
            QMessageBox.critical(
                self, 
                "JSON Parse Error", 
                f"Error parsing layer_standards.json:\n{str(e)}\n\nNo layer data will be displayed."
            )
            self.layer_standards = []
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Load Error", 
                f"Unexpected error loading layer standards:\n{str(e)}\n\nNo layer data will be displayed."
            )
            self.layer_standards = []
    
    def populate_table(self):
        self.layer_table.setSortingEnabled(False)
        self.layer_table.setRowCount(0)
        
        search_text = self.search_box.text().lower()
        discipline_filter = self.discipline_combo.currentText()
        status_filter = self.status_combo.currentText()
        category_filter = self.category_combo.currentText()
        
        for layer in self.layer_standards:
            name = layer.get('name', '')
            description = layer.get('description', '')
            discipline = layer.get('discipline', '')
            status = layer.get('status', '')
            category = layer.get('category', '')
            
            if search_text and search_text not in name.lower() and search_text not in description.lower():
                continue
            
            if discipline_filter != "All" and discipline != discipline_filter:
                continue
            
            if status_filter != "All" and status != status_filter:
                continue
            
            if category_filter != "All" and category != category_filter:
                continue
            
            row_position = self.layer_table.rowCount()
            self.layer_table.insertRow(row_position)
            
            color_code = layer.get('color_code', '')
            linetype = layer.get('linetype', '')
            is_plottable = layer.get('is_plottable', False)
            
            self.layer_table.setItem(row_position, 0, QTableWidgetItem(name))
            self.layer_table.setItem(row_position, 1, QTableWidgetItem(str(color_code)))
            self.layer_table.setItem(row_position, 2, QTableWidgetItem(linetype))
            self.layer_table.setItem(row_position, 3, QTableWidgetItem(status))
            self.layer_table.setItem(row_position, 4, QTableWidgetItem(category))
            self.layer_table.setItem(row_position, 5, QTableWidgetItem(discipline))
            self.layer_table.setItem(row_position, 6, QTableWidgetItem("Yes" if is_plottable else "No"))
        
        self.layer_table.setSortingEnabled(True)
    
    def apply_filters(self):
        self.populate_table()
    
    def on_layer_selected(self):
        selected_items = self.layer_table.selectedItems()
        if not selected_items:
            self.details_text.clear()
            return
        
        row = selected_items[0].row()
        layer_name = self.layer_table.item(row, 0).text()
        
        layer_data = None
        for layer in self.layer_standards:
            if layer.get('name') == layer_name:
                layer_data = layer
                break
        
        if not layer_data:
            self.details_text.clear()
            return
        
        description = layer_data.get('description', 'N/A')
        notes = layer_data.get('notes', 'N/A')
        typical_objects = layer_data.get('typical_object_types', [])
        plot_style = layer_data.get('plot_style_name', 'N/A')
        lineweight = layer_data.get('lineweight', 'N/A')
        
        typical_objects_str = ', '.join(typical_objects) if typical_objects else 'N/A'
        
        details_html = f"""
        <h3>{layer_name}</h3>
        <p><b>Description:</b><br>{description}</p>
        <p><b>Notes:</b><br>{notes if notes else 'None'}</p>
        <p><b>Typical Object Types:</b><br>{typical_objects_str}</p>
        <p><b>Plot Style:</b> {plot_style}</p>
        <p><b>Lineweight:</b> {lineweight}</p>
        """
        
        self.details_text.setHtml(details_html)

class BlockLibraryTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        # Controls
        controls = QHBoxLayout()
        controls.addWidget(QLabel("<b>Block Library Manager:</b>"))
        for btn_text in ["📚 Scan Library", "🔄 Update Definitions", "📥 Import Blocks", "📤 Export Catalog"]:
            btn = QPushButton(btn_text)
            btn.clicked.connect(lambda checked, text=btn_text: QMessageBox.information(self, "Block Operation", f"{text} would manage block definitions and library organization."))
            controls.addWidget(btn)
        controls.addStretch()
        layout.addLayout(controls)
        
        # Block catalog
        catalog_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: Block tree
        tree_group = QGroupBox("🔧 Block Categories")
        tree_layout = QVBoxLayout(tree_group)
        self.block_tree = QTreeWidget()
        self.block_tree.setHeaderLabel("Block Library")
        
        # Sample block hierarchy
        categories = {
            "Architectural": ["Door-36", "Window-48", "Toilet", "Sink"],
            "Electrical": ["Outlet", "Switch", "Light-Fixture", "Panel"],
            "Mechanical": ["Duct-6x10", "Diffuser", "Unit-5T", "Pipe-4in"],
            "Civil": ["Manhole", "Catch-Basin", "Tree-Deciduous", "Benchmark"]
        }
        
        for category, blocks in categories.items():
            cat_item = QTreeWidgetItem(self.block_tree, [category])
            for block in blocks:
                QTreeWidgetItem(cat_item, [block])
        
        tree_layout.addWidget(self.block_tree)
        catalog_splitter.addWidget(tree_group)
        
        # Right: Block details
        details_group = QGroupBox("📋 Block Details")
        details_layout = QVBoxLayout(details_group)
        details_layout.addWidget(QLabel("Block Name: <b>Door-36</b>"))
        details_layout.addWidget(QLabel("File: J:/LIB/BLOCKS/ARCH/Door-36.dwg"))
        details_layout.addWidget(QLabel("Used in: <b>8 drawings</b>"))
        details_layout.addWidget(QLabel("Total Insertions: <b>24</b>"))
        details_layout.addWidget(QLabel("Last Modified: 2024-09-15"))
        details_layout.addWidget(QPushButton("🔍 Show Usage Report"))
        details_layout.addStretch()
        catalog_splitter.addWidget(details_group)
        
        layout.addWidget(catalog_splitter)

class StandardsDashboardTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        # Controls
        controls = QHBoxLayout()
        controls.addWidget(QLabel("<b>CAD Standards Dashboard:</b>"))
        for btn_text in ["✅ Run Standards Check", "🔧 Auto-Fix Issues", "📋 Generate Report", "⚙️ Configure Rules"]:
            btn = QPushButton(btn_text)
            btn.clicked.connect(lambda checked, text=btn_text: QMessageBox.information(self, "Standards", f"{text} would validate and enforce CAD standards across all project drawings."))
            controls.addWidget(btn)
        controls.addStretch()
        layout.addLayout(controls)
        
        # Compliance overview
        overview_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: Compliance scores
        scores_group = QGroupBox("📊 Compliance Scores")
        scores_layout = QVBoxLayout(scores_group)
        
        compliance_data = [
            ("Layer Standards", "92%", "🟢"),
            ("Text Styles", "87%", "🟡"), 
            ("Dimension Styles", "95%", "🟢"),
            ("Block Standards", "78%", "🟡"),
            ("Drawing Setup", "100%", "🟢"),
            ("File Naming", "85%", "🟡")
        ]
        
        for standard, score, status in compliance_data:
            score_layout = QHBoxLayout()
            score_layout.addWidget(QLabel(f"{status} {standard}:"))
            score_layout.addWidget(QLabel(f"<b>{score}</b>"))
            score_layout.addStretch()
            scores_layout.addLayout(score_layout)
        
        scores_layout.addStretch()
        overview_splitter.addWidget(scores_group)
        
        # Right: Issues list
        issues_group = QGroupBox("⚠️ Standards Violations")
        issues_layout = QVBoxLayout(issues_group)
        
        self.issues_table = QTableWidget(6, 3)
        self.issues_table.setHorizontalHeaderLabels(["Issue", "Count", "Severity"])
        
        issues = [
            ("Non-standard layer names", "3", "Medium"),
            ("Incorrect text height", "12", "Low"),
            ("Missing dimension styles", "1", "High"),
            ("Improper block naming", "5", "Medium"),
            ("Wrong plot settings", "2", "Medium"),
            ("Outdated templates", "1", "Low")
        ]
        
        for i, (issue, count, severity) in enumerate(issues):
            self.issues_table.setItem(i, 0, QTableWidgetItem(issue))
            self.issues_table.setItem(i, 1, QTableWidgetItem(count))
            self.issues_table.setItem(i, 2, QTableWidgetItem(severity))
        
        issues_layout.addWidget(self.issues_table)
        overview_splitter.addWidget(issues_group)
        
        layout.addWidget(overview_splitter)

class ProjectReportsTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        # Report controls
        controls = QHBoxLayout()
        controls.addWidget(QLabel("<b>Project Reporting Center:</b>"))
        for btn_text in ["📋 Drawing Inventory", "📈 Progress Report", "🎯 Quality Metrics", "💾 Export Reports"]:
            btn = QPushButton(btn_text)
            btn.clicked.connect(lambda checked, text=btn_text: QMessageBox.information(self, "Reports", f"{text} would generate comprehensive project documentation and metrics."))
            controls.addWidget(btn)
        controls.addStretch()
        layout.addLayout(controls)
        
        # Report dashboard
        reports_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: Project overview
        overview_group = QGroupBox("📊 Project Overview")
        overview_layout = QVBoxLayout(overview_group)
        overview_layout.addWidget(QLabel("Project: <b>8888-09</b>"))
        overview_layout.addWidget(QLabel("Total Drawings: <b>47</b>"))
        overview_layout.addWidget(QLabel("Completed: <b>32 (68%)</b>"))
        overview_layout.addWidget(QLabel("In Progress: <b>12 (26%)</b>"))
        overview_layout.addWidget(QLabel("Not Started: <b>3 (6%)</b>"))
        overview_layout.addWidget(QLabel("Last Updated: <b>Today 14:30</b>"))
        overview_layout.addWidget(QPushButton("🔄 Refresh Status"))
        overview_layout.addStretch()
        reports_splitter.addWidget(overview_group)
        
        # Right: Activity log
        activity_group = QGroupBox("📝 Recent Activity")
        activity_layout = QVBoxLayout(activity_group)
        
        self.activity_table = QTableWidget(8, 3)
        self.activity_table.setHorizontalHeaderLabels(["Time", "User", "Action"])
        
        activities = [
            ("14:25", "J.Smith", "Modified Plan-Level1.dwg"),
            ("14:18", "M.Jones", "Created Section-A.dwg"),
            ("13:45", "K.Brown", "Updated title blocks (12 files)"),
            ("13:30", "J.Smith", "Ran layer cleanup batch"),
            ("12:15", "A.Wilson", "Exported PDF set"),
            ("11:45", "M.Jones", "Added electrical details"),
            ("10:30", "J.Smith", "Project sync completed"),
            ("09:15", "K.Brown", "Standards check passed")
        ]
        
        for i, (time, user, action) in enumerate(activities):
            self.activity_table.setItem(i, 0, QTableWidgetItem(time))
            self.activity_table.setItem(i, 1, QTableWidgetItem(user))
            self.activity_table.setItem(i, 2, QTableWidgetItem(action))
        
        activity_layout.addWidget(self.activity_table)
        reports_splitter.addWidget(activity_group)
        
        layout.addWidget(reports_splitter)

class DXFAnalysisTab(QWidget):
    """DXF Analysis tab for extracting CAD drawing data into structured JSON files."""
    
    def __init__(self):
        super().__init__()
        self.analyzer = DXFAnalyzer()
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Header section
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("<b>🔍 CAD File Analysis & Data Extraction</b>"))
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # Description
        description = QLabel(
            "Analyze structured data from DXF files or load existing JSON analysis files. "
            "Convert DWG files to DXF first using the Batch Operations tab, then analyze the results. "
            "View entities, layers, text, dimensions, metadata, and statistics in an organized format."
        )
        description.setWordWrap(True)
        description.setStyleSheet("color: gray; padding: 5px;")
        layout.addWidget(description)
        
        # Main splitter for file operations and results
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel: File operations
        file_ops_group = QGroupBox("📁 File Operations")
        file_ops_layout = QVBoxLayout(file_ops_group)
        
        # File analysis section
        analysis_group = QGroupBox("📊 Analysis Options")
        analysis_layout = QVBoxLayout(analysis_group)
        
        # File selection
        file_select_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Select a DXF file or JSON analysis file...")
        self.file_path_edit.setReadOnly(True)
        file_select_layout.addWidget(self.file_path_edit)
        
        self.browse_file_btn = QPushButton("📂 Browse")
        self.browse_file_btn.clicked.connect(self.browse_file)
        file_select_layout.addWidget(self.browse_file_btn)
        analysis_layout.addLayout(file_select_layout)
        
        # Analysis buttons
        button_layout = QHBoxLayout()
        
        self.analyze_dxf_btn = QPushButton("🔍 Analyze DXF File")
        self.analyze_dxf_btn.clicked.connect(self.analyze_dxf_file)
        self.analyze_dxf_btn.setEnabled(False)
        button_layout.addWidget(self.analyze_dxf_btn)
        
        self.load_json_btn = QPushButton("📋 Load JSON Analysis")
        self.load_json_btn.clicked.connect(self.load_json_analysis)
        self.load_json_btn.setEnabled(False)
        button_layout.addWidget(self.load_json_btn)
        
        analysis_layout.addLayout(button_layout)
        
        file_ops_layout.addWidget(analysis_group)
        
        # Batch processing section
        batch_group = QGroupBox("Batch Processing")
        batch_layout = QVBoxLayout(batch_group)
        
        # Input folder selection
        input_folder_layout = QHBoxLayout()
        input_folder_layout.addWidget(QLabel("Input Folder:"))
        self.input_folder_edit = QLineEdit()
        self.input_folder_edit.setPlaceholderText("Select folder containing DXF and DWG files...")
        self.input_folder_edit.setReadOnly(True)
        input_folder_layout.addWidget(self.input_folder_edit)
        
        self.browse_input_btn = QPushButton("📂 Browse")
        self.browse_input_btn.clicked.connect(self.browse_input_folder)
        input_folder_layout.addWidget(self.browse_input_btn)
        batch_layout.addLayout(input_folder_layout)
        
        # Output folder selection
        output_folder_layout = QHBoxLayout()
        output_folder_layout.addWidget(QLabel("Output Folder:"))
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("Select output folder for JSON files...")
        self.output_folder_edit.setReadOnly(True)
        output_folder_layout.addWidget(self.output_folder_edit)
        
        self.browse_output_btn = QPushButton("📂 Browse")
        self.browse_output_btn.clicked.connect(self.browse_output_folder)
        output_folder_layout.addWidget(self.browse_output_btn)
        batch_layout.addLayout(output_folder_layout)
        
        # Batch process button
        self.batch_process_btn = QPushButton("⚡ Batch Process CAD Files")
        self.batch_process_btn.clicked.connect(self.batch_process_files)
        self.batch_process_btn.setEnabled(False)
        batch_layout.addWidget(self.batch_process_btn)
        
        file_ops_layout.addWidget(batch_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        file_ops_layout.addWidget(self.progress_bar)
        
        file_ops_layout.addStretch()
        main_splitter.addWidget(file_ops_group)
        
        # Right panel: Results and analysis
        results_group = QGroupBox("📊 Analysis Results")
        results_layout = QVBoxLayout(results_group)
        
        # Results display tabs
        self.results_tabs = QTabWidget()
        
        # Summary tab
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setPlaceholderText("Analysis summary will appear here...")
        self.results_tabs.addTab(self.summary_text, "📋 Summary")
        
        # Raw JSON tab
        self.json_text = QTextEdit()
        self.json_text.setReadOnly(True)
        self.json_text.setPlaceholderText("Raw JSON data will appear here...")
        self.results_tabs.addTab(self.json_text, "📝 Raw JSON")
        
        # Statistics tab
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)
        self.stats_text.setPlaceholderText("Detailed statistics will appear here...")
        self.results_tabs.addTab(self.stats_text, "📈 Statistics")
        
        results_layout.addWidget(self.results_tabs)
        
        # Export controls
        export_layout = QHBoxLayout()
        self.export_json_btn = QPushButton("💾 Export JSON")
        self.export_json_btn.clicked.connect(self.export_current_analysis)
        self.export_json_btn.setEnabled(False)
        export_layout.addWidget(self.export_json_btn)
        
        self.clear_results_btn = QPushButton("🗑️ Clear Results")
        self.clear_results_btn.clicked.connect(self.clear_results)
        export_layout.addWidget(self.clear_results_btn)
        export_layout.addStretch()
        results_layout.addLayout(export_layout)
        
        main_splitter.addWidget(results_group)
        
        # Set splitter proportions
        main_splitter.setSizes([400, 800])
        layout.addWidget(main_splitter)
        
        # Status tracking
        self.current_analysis = None
        
    def browse_file(self):
        """Browse for a DXF file or JSON analysis file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File",
            "",
            "Analysis Files (*.json *.dxf);;JSON Files (*.json);;DXF Files (*.dxf);;All Files (*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            file_ext = Path(file_path).suffix.lower()
            
            # Enable appropriate buttons based on file type
            if file_ext == '.json':
                self.load_json_btn.setEnabled(True)
                self.analyze_dxf_btn.setEnabled(False)
            elif file_ext == '.dxf':
                self.analyze_dxf_btn.setEnabled(True)
                self.load_json_btn.setEnabled(False)
            else:
                self.analyze_dxf_btn.setEnabled(False)
                self.load_json_btn.setEnabled(False)
    
    def browse_input_folder(self):
        """Browse for input folder containing DXF and DWG files."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Input Folder (DXF/DWG files)",
            ""
        )
        
        if folder_path:
            self.input_folder_edit.setText(folder_path)
            self.update_batch_button_state()
    
    def browse_output_folder(self):
        """Browse for output folder for JSON results."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            ""
        )
        
        if folder_path:
            self.output_folder_edit.setText(folder_path)
            self.update_batch_button_state()
    
    def update_batch_button_state(self):
        """Enable batch process button when both folders are selected."""
        input_ok = bool(self.input_folder_edit.text().strip())
        output_ok = bool(self.output_folder_edit.text().strip())
        self.batch_process_btn.setEnabled(input_ok and output_ok)
    
    def analyze_dxf_file(self):
        """Analyze a single DXF file.""" 
        file_path = Path(self.file_path_edit.text())
        
        if not file_path.exists():
            QMessageBox.warning(self, "File Error", "Selected file does not exist.")
            return
        
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            
            # Perform analysis
            self.current_analysis = self.analyzer.analyze_file(file_path)
            
            # Check if analysis succeeded
            if "error" in self.current_analysis:
                QMessageBox.warning(self, "Analysis Error", f"Error analyzing file:\n{self.current_analysis['error']}")
                return
            
            # Display results
            self.display_analysis_results()
            
            # Enable export button
            self.export_json_btn.setEnabled(True)
            
            QMessageBox.information(self, "Analysis Complete", 
                                  f"Successfully analyzed {file_path.name}")
            
        except Exception as e:
            QMessageBox.critical(self, "Analysis Error", 
                               f"Error analyzing file: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def analyze_single_file(self):
        """Legacy method - redirect to analyze_dxf_file for compatibility."""
        self.analyze_dxf_file()
    
    def load_json_analysis(self):
        """Load and display existing JSON analysis results."""
        file_path = Path(self.file_path_edit.text())
        
        if not file_path.exists():
            QMessageBox.warning(self, "File Error", "Selected file does not exist.")
            return
        
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            
            # Load JSON file
            with open(file_path, 'r', encoding='utf-8') as f:
                result = json.load(f)
            
            # Validate that this is an analysis JSON file
            if not self.validate_analysis_json(result):
                QMessageBox.warning(self, "Invalid File", 
                                   "This JSON file does not appear to be a CAD analysis result.")
                return
            
            # Store and display results
            self.current_analysis = result
            self.display_analysis_results()
            
            # Enable export button
            self.export_json_btn.setEnabled(True)
            
            QMessageBox.information(self, "JSON Loaded", 
                                  f"Successfully loaded analysis from {file_path.name}")
            
        except json.JSONDecodeError as e:
            QMessageBox.critical(self, "JSON Error", f"Error parsing JSON file:\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Failed to load file:\n{str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def validate_analysis_json(self, data):
        """Validate that the JSON contains CAD analysis data."""
        if not isinstance(data, dict):
            return False
        
        # Check for key analysis sections
        expected_keys = ["file_info", "statistics", "extraction_timestamp"]
        has_required_keys = any(key in data for key in expected_keys)
        
        # Also check for typical analysis structure
        has_entities = "entities" in data
        has_layers = "layers" in data
        
        return has_required_keys or has_entities or has_layers
    
    def batch_process_files(self):
        """Process multiple DXF and DWG files in batch."""
        input_folder = Path(self.input_folder_edit.text())
        output_folder = Path(self.output_folder_edit.text())
        
        if not input_folder.exists():
            QMessageBox.warning(self, "Folder Error", "Input folder does not exist.")
            return
        
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            
            # Perform batch analysis on both DXF and DWG files
            results = self.analyzer.batch_analyze(input_folder, output_folder, "*.*")
            
            # Display batch results
            self.display_batch_results(results)
            
            QMessageBox.information(self, "Batch Processing Complete", 
                                  f"Processed {results['processed_files']} files successfully.\n"
                                  f"DWG files: {results.get('dwg_files_processed', 0)}\n"
                                  f"DXF files: {results.get('dxf_files_processed', 0)}\n"
                                  f"Failed: {results['failed_files']} files.")
            
        except Exception as e:
            QMessageBox.critical(self, "Batch Processing Error", 
                               f"Error during batch processing: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def display_analysis_results(self):
        """Display the analysis results in the UI."""
        if not self.current_analysis:
            return
        
        # Generate summary
        if "error" in self.current_analysis:
            summary = f"❌ Analysis Failed\n\nError: {self.current_analysis['error']}"
            json_data = json.dumps(self.current_analysis, indent=2)
            stats = "No statistics available due to error."
        else:
            summary = self.generate_summary_text(self.current_analysis)
            json_data = json.dumps(self.current_analysis, indent=2)
            stats = self.generate_statistics_text(self.current_analysis)
        
        # Update UI
        self.summary_text.setPlainText(summary)
        self.json_text.setPlainText(json_data)
        self.stats_text.setPlainText(stats)
    
    def display_batch_results(self, results):
        """Display batch processing results."""
        summary = f"""🔍 Batch Processing Results

Total Files: {results['total_files']}
Successfully Processed: {results['processed_files']}
  • DWG files processed: {results.get('dwg_files_processed', 0)}
  • DXF files processed: {results.get('dxf_files_processed', 0)}
Failed: {results['failed_files']}

Processing Details:
"""
        
        for filename, details in results['processing_summary'].items():
            status_icon = "✅" if details['status'] == 'success' else "❌"
            summary += f"{status_icon} {filename}: {details['status']}\n"
            
            if details['status'] == 'success':
                summary += f"   → Format: {details.get('original_format', 'Unknown')}\n"
                summary += f"   → Entities: {details.get('entity_count', 0)}\n"
                summary += f"   → Output: {Path(details['output_file']).name}\n"
                if details.get('was_converted', False):
                    summary += f"   → Converted from DWG to DXF\n"
            elif 'error' in details:
                summary += f"   → Error: {details['error']}\n"
        
        # Clear current analysis since this is batch results
        self.current_analysis = None
        self.export_json_btn.setEnabled(False)
        
        # Display in summary tab
        self.summary_text.setPlainText(summary)
        self.json_text.setPlainText(json.dumps(results, indent=2))
        self.stats_text.setPlainText("Batch processing statistics shown in summary.")
        
        # Switch to summary tab
        self.results_tabs.setCurrentIndex(0)
    
    def generate_summary_text(self, analysis):
        """Generate a human-readable summary of the analysis."""
        file_info = analysis.get('file_info', {})
        metadata = analysis.get('drawing_metadata', {})
        stats = analysis.get('statistics', {})
        
        # Check for conversion info
        conversion_info = analysis.get('conversion_info', {})
        was_converted = conversion_info.get('was_converted_from_dwg', False)
        original_format = conversion_info.get('original_format', 'DXF')
        
        summary = f"""🔍 CAD File Analysis Summary

File Information:
📁 File: {file_info.get('file_name', 'Unknown')}
📏 Size: {file_info.get('file_size_bytes', 0):,} bytes
🔢 DXF Version: {file_info.get('dxf_version', 'Unknown')}
📐 Units: {metadata.get('units', 'Unknown')}
📋 Original Format: {original_format}"""
        
        if was_converted:
            summary += f"\n🔄 Converted from DWG using {conversion_info.get('converter_used', 'AutoCAD Core Console')}"
        
        summary += f"""

Entity Statistics:
🎨 Layers: {stats.get('layer_count', 0)}
🧱 Blocks: {stats.get('block_count', 0)}
📄 Layouts: {stats.get('layout_count', 0)}
📝 Text Objects: {stats.get('text_object_count', 0)}
📏 Dimensions: {stats.get('dimension_count', 0)}

Entity Breakdown:
"""
        
        entity_counts = stats.get('entity_counts', {})
        total_entities = sum(entity_counts.values())
        summary += f"Total Entities: {total_entities}\n\n"
        
        for entity_type, count in sorted(entity_counts.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_entities * 100) if total_entities > 0 else 0
            summary += f"• {entity_type}: {count} ({percentage:.1f}%)\n"
        
        return summary
    
    def generate_statistics_text(self, analysis):
        """Generate detailed statistics."""
        stats = f"""📈 Detailed Statistics

Drawing Metadata:
{json.dumps(analysis.get('drawing_metadata', {}), indent=2)}

Layer Information:
Total Layers: {len(analysis.get('layers', []))}

"""
        
        # Layer breakdown
        layers = analysis.get('layers', [])
        if layers:
            stats += "Layer Details:\n"
            for layer in layers[:10]:  # Show first 10 layers
                stats += f"• {layer.get('name', 'Unknown')}: Color {layer.get('color', 'N/A')}, "
                stats += f"Linetype {layer.get('linetype', 'N/A')}\n"
            if len(layers) > 10:
                stats += f"... and {len(layers) - 10} more layers\n"
        
        stats += f"\nText Objects: {analysis.get('statistics', {}).get('text_object_count', 0)}\n"
        stats += f"Dimensions: {analysis.get('statistics', {}).get('dimension_count', 0)}\n"
        
        return stats
    
    def export_current_analysis(self):
        """Export the current analysis to a JSON file."""
        if not self.current_analysis:
            QMessageBox.warning(self, "No Analysis", "No analysis data to export.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Analysis Results",
            "dxf_analysis.json",
            "JSON Files (*.json);;All Files (*)"
        )
        
        if file_path:
            try:
                if self.analyzer.export_to_json(Path(file_path)):
                    QMessageBox.information(self, "Export Complete", 
                                          f"Analysis results saved to {Path(file_path).name}")
                else:
                    QMessageBox.warning(self, "Export Failed", "Failed to save analysis results.")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Error saving file: {str(e)}")
    
    def clear_results(self):
        """Clear all analysis results."""
        self.summary_text.clear()
        self.json_text.clear()
        self.stats_text.clear()
        self.current_analysis = None
        self.export_json_btn.setEnabled(False)
        self.file_path_edit.clear()
        self.analyze_btn.setEnabled(False)

class BatchOperationsTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        
        # Operation controls
        controls = QHBoxLayout()
        controls.addWidget(QLabel("<b>Batch Operations Center:</b>"))
        for btn_text in ["📁 File Operations", "🏷️ Bulk Rename", "📋 Title Block Update", "🗂️ Archive Project"]:
            btn = QPushButton(btn_text)
            btn.clicked.connect(lambda checked, text=btn_text: QMessageBox.information(self, "Batch Operations", f"{text} would perform bulk operations across multiple drawings efficiently."))
            controls.addWidget(btn)
        controls.addStretch()
        layout.addLayout(controls)
        
        # Operations dashboard
        ops_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left: File operations
        file_ops_group = QGroupBox("📁 File Operations")
        file_ops_layout = QVBoxLayout(file_ops_group)
        
        # DWG Conversion Section
        conversion_group = QGroupBox("🔄 DWG Conversion")
        conversion_layout = QVBoxLayout(conversion_group)
        
        # Input folder selection
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("Input Folder:"))
        self.input_folder_edit = QLineEdit()
        self.input_folder_edit.setPlaceholderText("Select folder containing DWG files...")
        input_layout.addWidget(self.input_folder_edit)
        self.browse_input_btn = QPushButton("📂 Browse")
        self.browse_input_btn.clicked.connect(self.select_input_folder)
        input_layout.addWidget(self.browse_input_btn)
        conversion_layout.addLayout(input_layout)
        
        # Output folder selection
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Output Folder:"))
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("Select folder for DXF output...")
        output_layout.addWidget(self.output_folder_edit)
        self.browse_output_btn = QPushButton("📂 Browse")
        self.browse_output_btn.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.browse_output_btn)
        conversion_layout.addLayout(output_layout)
        
        # Conversion options
        options_layout = QHBoxLayout()
        self.include_analysis = QCheckBox("Include DXF Analysis")
        self.include_analysis.setChecked(True)
        self.include_analysis.setToolTip("Also analyze converted DXF files and generate JSON reports")
        options_layout.addWidget(self.include_analysis)
        
        self.include_metadata = QCheckBox("Export Metadata")
        self.include_metadata.setToolTip("Export comprehensive metadata for each drawing")
        options_layout.addWidget(self.include_metadata)
        
        options_layout.addStretch()
        conversion_layout.addLayout(options_layout)
        
        # Conversion buttons
        conv_buttons_layout = QHBoxLayout()
        self.convert_btn = QPushButton("🔧 Convert DWG → DXF")
        self.convert_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 8px 16px; }")
        self.convert_btn.clicked.connect(self.start_dwg_conversion)
        conv_buttons_layout.addWidget(self.convert_btn)
        
        self.automation_convert_btn = QPushButton("⚡ Use Automation Recipes")
        self.automation_convert_btn.setToolTip("Use the automation system with DWG conversion recipes")
        self.automation_convert_btn.clicked.connect(self.open_automation_for_conversion)
        conv_buttons_layout.addWidget(self.automation_convert_btn)
        
        conv_buttons_layout.addStretch()
        conversion_layout.addLayout(conv_buttons_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        conversion_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("")
        conversion_layout.addWidget(self.status_label)
        
        file_ops_layout.addWidget(conversion_group)
        
        # Other file operation buttons
        other_ops_group = QGroupBox("📋 Other Operations")
        other_ops_layout = QVBoxLayout(other_ops_group)
        
        ops_buttons = [
            ("🔄 Sync Project Folders", "Synchronize between local and network"),
            ("🗑️ Clean Backup Files", "Remove old .bak and temp files"),
            ("📦 Create Archive", "Package project for distribution"),
            ("🔗 Update Xref Paths", "Fix broken external references"),
            ("📏 Convert Units", "Batch convert between Imperial/Metric")
        ]
        
        for btn_text, tooltip in ops_buttons:
            btn = QPushButton(btn_text)
            btn.setToolTip(tooltip)
            btn.clicked.connect(lambda checked, text=btn_text: QMessageBox.information(self, "Operation", f"{text} operation ready. Select target files and configure parameters."))
            other_ops_layout.addWidget(btn)
        
        other_ops_layout.addStretch()
        file_ops_layout.addWidget(other_ops_group)
        
        file_ops_layout.addStretch()
        ops_splitter.addWidget(file_ops_group)
        
        # Right: Operation history
        history_group = QGroupBox("📜 Operation History")
        history_layout = QVBoxLayout(history_group)
        
        self.history_table = QTableWidget(10, 4)
        self.history_table.setHorizontalHeaderLabels(["Time", "Operation", "Files", "Status"])
        
        # Initialize with some sample history
        self.operation_history = [
            ("14:20", "Title Block Update", "12", "✅ Complete"),
            ("13:55", "Layer Cleanup", "8", "✅ Complete"),
            ("13:30", "Batch Plot", "15", "✅ Complete"),
            ("12:45", "Standards Check", "47", "✅ Complete"),
            ("11:20", "File Rename", "6", "✅ Complete"),
            ("10:15", "Export PDF", "12", "✅ Complete")
        ]
        
        self.update_history_table()
        
        history_layout.addWidget(self.history_table)
        ops_splitter.addWidget(history_group)
        
        layout.addWidget(ops_splitter)
    
    def update_history_table(self):
        """Update the operation history table."""
        self.history_table.setRowCount(len(self.operation_history))
        for i, (time, operation, files, status) in enumerate(self.operation_history):
            self.history_table.setItem(i, 0, QTableWidgetItem(time))
            self.history_table.setItem(i, 1, QTableWidgetItem(operation))
            self.history_table.setItem(i, 2, QTableWidgetItem(files))
            self.history_table.setItem(i, 3, QTableWidgetItem(status))
    
    def add_to_history(self, operation: str, files_count: int, status: str):
        """Add a new operation to the history."""
        import datetime
        current_time = datetime.datetime.now().strftime("%H:%M")
        self.operation_history.insert(0, (current_time, operation, str(files_count), status))
        
        # Keep only the last 10 operations
        if len(self.operation_history) > 10:
            self.operation_history = self.operation_history[:10]
        
        self.update_history_table()
    
    def select_input_folder(self):
        """Select the input folder containing DWG files."""
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder with DWG Files")
        if folder:
            self.input_folder_edit.setText(folder)
    
    def select_output_folder(self):
        """Select the output folder for DXF files."""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder for DXF Files")
        if folder:
            self.output_folder_edit.setText(folder)
    
    def start_dwg_conversion(self):
        """Start the DWG to DXF conversion process."""
        input_folder = self.input_folder_edit.text().strip()
        output_folder = self.output_folder_edit.text().strip()
        
        if not input_folder or not output_folder:
            QMessageBox.warning(self, "Missing Folders", "Please select both input and output folders.")
            return
        
        input_path = Path(input_folder)
        output_path = Path(output_folder)
        
        if not input_path.exists():
            QMessageBox.warning(self, "Invalid Input", "Input folder does not exist.")
            return
        
        # Check for DWG files
        dwg_files = list(input_path.glob("*.dwg"))
        if not dwg_files:
            QMessageBox.warning(self, "No DWG Files", "No DWG files found in the input folder.")
            return
        
        # Create output folder if it doesn't exist
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Start the conversion process
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(dwg_files))
        self.progress_bar.setValue(0)
        self.status_label.setText(f"Starting conversion of {len(dwg_files)} DWG files...")
        
        # Use DXFAnalyzer for conversion if analysis is requested
        if self.include_analysis.isChecked():
            self.convert_with_analysis(input_path, output_path, dwg_files)
        else:
            self.convert_dwg_only(input_path, output_path, dwg_files)
    
    def convert_with_analysis(self, input_path: Path, output_path: Path, dwg_files: list):
        """Convert DWG files to DXF with analysis using DXFAnalyzer."""
        try:
            analyzer = DXFAnalyzer()
            self.status_label.setText("Converting and analyzing DWG files...")
            
            # Use batch analysis which handles DWG conversion automatically
            results = analyzer.batch_analyze(input_path, output_path, "*.dwg")
            
            self.progress_bar.setValue(len(dwg_files))
            
            # Show results
            message = f"Conversion Complete!\n\n"
            message += f"Total files: {results['total_files']}\n"
            message += f"Successfully processed: {results['processed_files']}\n"
            message += f"DWG files converted: {results.get('dwg_files_processed', 0)}\n"
            message += f"Failed: {results['failed_files']}\n"
            
            if results['failed_files'] > 0:
                message += f"\nFailed files: {', '.join(results['failed_file_names'])}"
            
            QMessageBox.information(self, "Conversion Results", message)
            
            # Add to history
            status = "✅ Complete" if results['failed_files'] == 0 else f"⚠️ {results['failed_files']} Failed"
            self.add_to_history("DWG→DXF Conversion", results['processed_files'], status)
            
        except Exception as e:
            QMessageBox.critical(self, "Conversion Error", f"Error during conversion:\n{str(e)}")
            self.add_to_history("DWG→DXF Conversion", 0, "❌ Failed")
        finally:
            self.progress_bar.setVisible(False)
            self.status_label.setText("Ready")
    
    def convert_dwg_only(self, input_path: Path, output_path: Path, dwg_files: list):
        """Convert DWG files to DXF only using AutoCAD Core Console."""
        try:
            converted_count = 0
            failed_files = []
            
            # Check if AutoCAD Core Console is available
            if not ACCORECONSOLE_EXE.exists():
                QMessageBox.warning(self, "AutoCAD Not Found", 
                                   "AutoCAD Core Console not found. Please ensure AutoCAD is installed.")
                return
            
            for i, dwg_file in enumerate(dwg_files):
                try:
                    self.progress_bar.setValue(i)
                    self.status_label.setText(f"Converting {dwg_file.name}...")
                    
                    # Create output DXF path
                    dxf_file = output_path / (dwg_file.stem + ".dxf")
                    
                    # Create a temporary script for this conversion
                    script_content = f"""DXFOUT
{dxf_file}


QUIT
"""
                    
                    # Write temporary script
                    temp_script = output_path / f"temp_convert_{dwg_file.stem}.scr"
                    with open(temp_script, 'w') as f:
                        f.write(script_content)
                    
                    # Run AutoCAD Core Console
                    cmd = [
                        str(ACCORECONSOLE_EXE),
                        "/i", str(dwg_file),
                        "/s", str(temp_script)
                    ]
                    
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                    
                    # Clean up temporary script
                    if temp_script.exists():
                        temp_script.unlink()
                    
                    # Check if conversion was successful
                    if dxf_file.exists():
                        converted_count += 1
                        
                        # Export metadata if requested
                        if self.include_metadata.isChecked():
                            self.export_metadata_for_dwg(dwg_file, output_path)
                    else:
                        failed_files.append(dwg_file.name)
                        
                except subprocess.TimeoutExpired:
                    failed_files.append(f"{dwg_file.name} (timeout)")
                except Exception as e:
                    failed_files.append(f"{dwg_file.name} ({str(e)})")
            
            self.progress_bar.setValue(len(dwg_files))
            
            # Show results
            message = f"DWG Conversion Complete!\n\n"
            message += f"Total files: {len(dwg_files)}\n"
            message += f"Successfully converted: {converted_count}\n"
            message += f"Failed: {len(failed_files)}\n"
            
            if failed_files:
                message += f"\nFailed files:\n" + "\n".join(failed_files[:5])
                if len(failed_files) > 5:
                    message += f"\n... and {len(failed_files) - 5} more"
            
            QMessageBox.information(self, "Conversion Results", message)
            
            # Add to history
            status = "✅ Complete" if len(failed_files) == 0 else f"⚠️ {len(failed_files)} Failed"
            self.add_to_history("DWG→DXF Direct", converted_count, status)
            
        except Exception as e:
            QMessageBox.critical(self, "Conversion Error", f"Error during conversion:\n{str(e)}")
            self.add_to_history("DWG→DXF Direct", 0, "❌ Failed")
        finally:
            self.progress_bar.setVisible(False)
            self.status_label.setText("Ready")
    
    def export_metadata_for_dwg(self, dwg_file: Path, output_path: Path):
        """Export metadata for a DWG file using AutoCAD Core Console."""
        try:
            metadata_file = output_path / (dwg_file.stem + "_metadata.txt")
            
            # Create metadata export script - using Windows-style path separators
            metadata_file_str = str(metadata_file).replace("\\", "\\\\")
            script_content = f"""(setq dwg_path (getvar "dwgprefix"))
(setq dwg_name (vl-filename-base (getvar "dwgname")))
(setq json_path "{metadata_file_str}")

(setq outfile (open json_path "w"))
(if outfile
  (progn
    (write-line (strcat "Drawing: " (getvar "dwgname")) outfile)
    (write-line (strcat "Path: " dwg_path) outfile)
    (write-line (strcat "Version: " (getvar "dwgver")) outfile)
    (write-line (strcat "Units: " (rtos (getvar "insunits") 2 0)) outfile)
    (write-line (strcat "Export Date: " (rtos (getvar "cdate") 2 8)) outfile)
    (close outfile)
    (princ "\\nMetadata exported")
  )
  (princ "\\nError: Could not create metadata file")
)
QUIT
"""
            
            # Write temporary script
            temp_script = output_path / f"temp_metadata_{dwg_file.stem}.scr"
            with open(temp_script, 'w') as f:
                f.write(script_content)
            
            # Run AutoCAD Core Console for metadata export
            cmd = [
                str(ACCORECONSOLE_EXE),
                "/i", str(dwg_file),
                "/s", str(temp_script)
            ]
            
            subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            # Clean up temporary script
            if temp_script.exists():
                temp_script.unlink()
                
        except Exception as e:
            print(f"Warning: Failed to export metadata for {dwg_file.name}: {e}")
    
    def open_automation_for_conversion(self):
        """Open the Automation Hub tab with DWG conversion recipes pre-selected."""
        QMessageBox.information(self, "Automation Recipes", 
                               "This will open the Automation Hub where you can:\n\n"
                               "1. Select your DWG files\n"
                               "2. Choose 'DWG Conversion' recipes\n"
                               "3. Run bulk conversion via AutoCAD\n\n"
                               "Navigate to the '🚀 Automation Hub' tab to get started.")

class HealthCheckTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        self.run_audit_btn = QPushButton("Run Full Project Audit...")
        self.run_audit_btn.clicked.connect(lambda: QMessageBox.information(self, "Not Implemented", "This feature is not yet connected."))
        layout.addWidget(self.run_audit_btn)
        summary_group = QGroupBox("Audit Summary")
        summary_layout = QHBoxLayout(summary_group)
        summary_layout.addWidget(QLabel("Drawings Scanned: [ 0 ]")); summary_layout.addWidget(QLabel("Issues Found: [ 0 ]"))
        layout.addWidget(summary_group)
        results_group = QGroupBox("Results")
        results_layout = QVBoxLayout(results_group)
        self.results_tree = QTreeWidget(); self.results_tree.setHeaderLabels(["Issue", "Details", "File"])
        results_layout.addWidget(self.results_tree)
        layout.addWidget(results_group)

class SheetSetTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        ss_tree_group = QGroupBox("Sheet Set")
        ss_tree_layout = QVBoxLayout(ss_tree_group)
        self.ss_tree = QTreeWidget(); self.ss_tree.setHeaderHidden(True)
        ss_tree_layout.addWidget(self.ss_tree)
        splitter.addWidget(ss_tree_group)
        props_group = QGroupBox("Sheet Properties")
        props_layout = QFormLayout(props_group)
        props_layout.addRow("Name:", QLineEdit()); props_layout.addRow("File:", QLineEdit()); props_layout.addRow("Layout:", QLineEdit())
        splitter.addWidget(props_group)
        layout.addWidget(splitter)
        button_layout = QHBoxLayout()
        for text in ["Add Sheet...", "Remove Sheet", "Update Title Block"]:
            btn = QPushButton(text)
            btn.clicked.connect(lambda chk, t=text: QMessageBox.information(self, "Not Implemented", f"The '{t}' feature is not yet connected."))
            button_layout.addWidget(btn)
        layout.addLayout(button_layout)

def global_exception_hook(exctype, value, tb):
    """Global exception handler to prevent silent crashes."""
    error_message = "".join(traceback.format_exception(exctype, value, tb))
    error_box = QMessageBox()
    error_box.setIcon(QMessageBox.Icon.Critical)
    error_box.setText("An unexpected error occurred!")
    error_box.setInformativeText("Please copy the details below and report the issue.")
    error_box.setDetailedText(error_message)
    error_box.setModal(True)
    error_box.exec()
    sys.exit(1)

if __name__ == "__main__":
    sys.excepthook = global_exception_hook
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

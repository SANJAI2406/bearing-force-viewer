"""
Bearing Force Bode Plot Viewer
A sophisticated, interactive visualization tool for Romax DOE simulation results.

Features:
- Clean light theme with elegant typography
- RIGHT-CLICK VALIDATION: Click any curve to open source CSV and image
- Interactive graph tracking with crosshairs and coordinates
- Synchronized cursors across all subplots
- Data point inspection with hover tooltips
- Source file traceability for curve verification
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import subprocess
import platform

# Modern UI Framework
try:
    import customtkinter as ctk
    ctk.set_appearance_mode("light")  # Light mode
    ctk.set_default_color_theme("blue")
    HAS_CTK = True
except ImportError:
    HAS_CTK = False
    print("CustomTkinter not found. Install with: pip install customtkinter")

# Try importing PIL for image display
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ═══════════════════════════════════════════════════════════════════════════════
# OCR ORDER CORRECTION CONFIG - Modify this if you have different orders!
# ═══════════════════════════════════════════════════════════════════════════════
# Set to None to disable correction (accept OCR as-is)
# Set to a dict to enable smart correction for common OCR errors
# Example: OCR reads "52" as "5" -> correct "5" to "52"
#
# If your orders are multiples of 26 (26, 52, 78), keep this enabled.
# If you have Order 5, Order 7, etc. as REAL orders, set this to None.
OCR_ORDER_CORRECTIONS = {
    5: 52,   # OCR misreads "52" as "5" (missing the "2")
    7: 78,   # OCR misreads "78" as "7" (missing the "8")
    2: 26,   # OCR misreads "26" as "2" (missing the "6")
}
# To disable: OCR_ORDER_CORRECTIONS = None

# ═══════════════════════════════════════════════════════════════════════════════
# OCR IMAGE CROP PERCENTAGE - How much of the image top to scan for title text
# ═══════════════════════════════════════════════════════════════════════════════
# Default: 0.15 (15% of image height)
# Increase if title text is being cut off
# Decrease if OCR picks up too much noise from the plot area
OCR_CROP_PERCENTAGE = 0.10  # 10% - optimal for capturing "Force Y Component" text

# ═══════════════════════════════════════════════════════════════════════════════
# OCR SETUP - Bundled models for offline/firewall environments
# ═══════════════════════════════════════════════════════════════════════════════
USE_EASYOCR = False
USE_PYTESSERACT = False
ocr_reader = None
OCR_INIT_ERROR = None

def get_bundled_model_path():
    """Get path to bundled EasyOCR models (works with PyInstaller)."""
    if getattr(sys, 'frozen', False):
        # Running as compiled exe - models are in _MEIPASS/easyocr_models
        base_path = sys._MEIPASS
        model_path = os.path.join(base_path, 'easyocr_models')
        if os.path.exists(model_path):
            return model_path
    # Running as script - use default EasyOCR location
    return None

try:
    import easyocr

    # Check for bundled models first (for exe distribution)
    bundled_path = get_bundled_model_path()

    if bundled_path:
        # Use bundled models - NO internet required
        print(f"[INFO] Using bundled OCR models from: {bundled_path}")
        ocr_reader = easyocr.Reader(
            ['en'],
            gpu=False,
            verbose=False,
            model_storage_directory=bundled_path,
            download_enabled=False
        )
        USE_EASYOCR = True
        print("[OK] EasyOCR initialized with bundled models (offline mode)")
    else:
        # Normal initialization - may download models
        ocr_reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        USE_EASYOCR = True
        print("[OK] EasyOCR initialized successfully")

except ImportError:
    print("[INFO] EasyOCR not installed")
    try:
        import pytesseract
        USE_PYTESSERACT = True
        print("[OK] Pytesseract available")
    except ImportError:
        print("[INFO] No OCR engine installed")
except Exception as e:
    # Catch network errors, timeout, firewall blocks, etc.
    error_msg = str(e)
    if "urlopen error" in error_msg or "WinError 10060" in error_msg or "timed out" in error_msg.lower():
        OCR_INIT_ERROR = "Network blocked (firewall) - OCR models cannot be downloaded"
    else:
        OCR_INIT_ERROR = f"OCR init failed: {error_msg[:100]}"
    print(f"[WARN] {OCR_INIT_ERROR}")
    print("[INFO] Continuing without OCR - bearing/direction from filename only")

# ═══════════════════════════════════════════════════════════════════════════════
# DEBUG MODE - Set to True for detailed console logging
# ═══════════════════════════════════════════════════════════════════════════════
DEBUG_MODE = True
DEBUG_LOG_FILE = None  # Will be set when loading data

def debug_print(msg, level="INFO"):
    """Print debug message to console and optionally write to file."""
    if DEBUG_MODE:
        prefix = {
            "INFO": "[INFO]",
            "WARN": "[WARN]",
            "ERROR": "[ERROR]",
            "SUCCESS": "[OK]",
            "OCR": "[OCR]",
            "FILE": "[FILE]"
        }.get(level, "[DEBUG]")
        line = f"{prefix} {msg}"
        print(line)
        # Also write to file if set
        if DEBUG_LOG_FILE:
            try:
                with open(DEBUG_LOG_FILE, 'a', encoding='utf-8') as logf:
                    logf.write(line + "\n")
            except:
                pass

def start_debug_log(folder):
    """Start a new debug log file in the data folder."""
    global DEBUG_LOG_FILE
    import datetime
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    DEBUG_LOG_FILE = os.path.join(folder, f'debug_log_{timestamp}.txt')
    # Clear/create file
    with open(DEBUG_LOG_FILE, 'w', encoding='utf-8') as logf:
        logf.write("Bearing Force Viewer - Debug Log\n")
        logf.write(f"Generated: {datetime.datetime.now()}\n")
        logf.write("=" * 70 + "\n\n")
    debug_print(f"Debug log file: {DEBUG_LOG_FILE}", "INFO")
    return DEBUG_LOG_FILE

# ═══════════════════════════════════════════════════════════════════════════════
# OCR METADATA CACHING - Speeds up subsequent loads dramatically
# ═══════════════════════════════════════════════════════════════════════════════
OCR_CACHE_FILENAME = ".bearing_force_ocr_cache.json"

def get_cache_path(folder):
    """Get path to OCR cache file in the data folder."""
    return os.path.join(folder, OCR_CACHE_FILENAME)

def load_ocr_cache(folder):
    """Load OCR metadata cache from JSON file. Returns dict or None."""
    import json
    cache_path = get_cache_path(folder)
    if not os.path.exists(cache_path):
        debug_print(f"No cache file found at {cache_path}", "INFO")
        return None
    try:
        with open(cache_path, 'r', encoding='utf-8') as f:
            cache = json.load(f)
        debug_print(f"Loaded OCR cache with {len(cache.get('metadata', {}))} entries", "SUCCESS")
        return cache
    except Exception as e:
        debug_print(f"Failed to load cache: {e}", "WARN")
        return None

def save_ocr_cache(folder, file_metadata):
    """Save OCR metadata to JSON cache file."""
    import json
    import datetime
    cache_path = get_cache_path(folder)
    cache = {
        'version': '1.0',
        'created': datetime.datetime.now().isoformat(),
        'folder': folder,
        'metadata': file_metadata
    }
    try:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cache, f, indent=2)
        debug_print(f"Saved OCR cache with {len(file_metadata)} entries to {cache_path}", "SUCCESS")
        return True
    except Exception as e:
        debug_print(f"Failed to save cache: {e}", "WARN")
        return False

def is_cache_valid(folder, cache):
    """Check if cache is still valid (same CSV files exist)."""
    if not cache or 'metadata' not in cache:
        return False
    # Quick check: count CSV files
    csv_files = list(Path(folder).glob("*.csv"))
    cached_count = len(cache['metadata'])
    if len(csv_files) != cached_count:
        debug_print(f"Cache invalid: {len(csv_files)} CSVs vs {cached_count} cached", "WARN")
        return False
    # Check if all cached files still exist
    for file_key in cache['metadata'].keys():
        csv_path = Path(folder) / (file_key + ".csv")
        if not csv_path.exists():
            debug_print(f"Cache invalid: {file_key}.csv no longer exists", "WARN")
            return False
    debug_print("Cache is valid", "SUCCESS")
    return True

# Print startup info
debug_print("=" * 60, "INFO")
debug_print("Bearing Force Viewer - DEBUG MODE ENABLED", "INFO")
debug_print(f"PIL available: {HAS_PIL}", "INFO")
debug_print(f"EasyOCR available: {USE_EASYOCR}", "INFO")
debug_print(f"Pytesseract available: {USE_PYTESSERACT}", "INFO")
debug_print("=" * 60, "INFO")


# ═══════════════════════════════════════════════════════════════════════════════
# APPLE TV INSPIRED LIGHT THEME
# ═══════════════════════════════════════════════════════════════════════════════

class Theme:
    """Clean light theme - clean, elegant, minimal"""
    
    # Background colors - soft whites and light grays
    BG_PRIMARY = "#FFFFFF"        # Pure white
    BG_SECONDARY = "#F5F5F7"      # Apple light gray
    BG_TERTIARY = "#FAFAFA"       # Slightly off-white
    BG_HOVER = "#E8E8ED"          # Hover state
    BG_CARD = "#FFFFFF"           # Card background
    
    # Accent colors - Apple's signature palette
    ACCENT_PRIMARY = "#007AFF"    # Apple blue
    ACCENT_SECONDARY = "#34C759"  # Apple green
    ACCENT_WARNING = "#FF9500"    # Apple orange
    ACCENT_ERROR = "#FF3B30"      # Apple red
    ACCENT_PURPLE = "#AF52DE"     # Apple purple
    ACCENT_TEAL = "#5AC8FA"       # Apple teal
    
    # Text colors
    TEXT_PRIMARY = "#1D1D1F"      # Apple dark text
    TEXT_SECONDARY = "#86868B"    # Apple secondary text
    TEXT_MUTED = "#AEAEB2"        # Muted text
    
    # Border colors
    BORDER_DEFAULT = "#D2D2D7"    # Light border
    BORDER_ACTIVE = "#007AFF"     # Active border (blue)
    
    # Shadow and depth
    SHADOW_COLOR = "rgba(0,0,0,0.04)"
    
    # Plot colors - vibrant but refined
    PLOT_COLORS = [
        "#007AFF",  # Blue
        "#34C759",  # Green
        "#FF9500",  # Orange
        "#AF52DE",  # Purple
        "#FF3B30",  # Red
        "#5AC8FA",  # Teal
        "#FFCC00",  # Yellow
        "#FF2D55",  # Pink
        "#00C7BE",  # Mint
        "#FF6482",  # Coral
        "#30B0C7",  # Cyan
        "#A2845E",  # Brown
        "#8E8E93",  # Gray
        "#64D2FF",  # Light Blue
        "#BF5AF2",  # Violet
        "#32D74B",  # Bright Green
        "#FF453A",  # Bright Red
        "#FFD60A",  # Bright Yellow
        "#0A84FF",  # Bright Blue
        "#FF6F61",  # Salmon
    ]
    
    # Matplotlib style for light theme
    MPL_STYLE = {
        'figure.facecolor': BG_SECONDARY,
        'axes.facecolor': BG_CARD,
        'axes.edgecolor': BORDER_DEFAULT,
        'axes.labelcolor': TEXT_PRIMARY,
        'axes.titlecolor': TEXT_PRIMARY,
        'xtick.color': TEXT_SECONDARY,
        'ytick.color': TEXT_SECONDARY,
        'grid.color': BORDER_DEFAULT,
        'grid.alpha': 0.5,
        'text.color': TEXT_PRIMARY,
        'legend.facecolor': BG_CARD,
        'legend.edgecolor': BORDER_DEFAULT,
    }
    
    # Font settings - San Francisco style
    FONT_FAMILY = "SF Pro Display" if platform.system() == "Darwin" else "Segoe UI"
    FONT_FAMILY_MONO = "SF Mono" if platform.system() == "Darwin" else "Consolas"


# ═══════════════════════════════════════════════════════════════════════════════
# SOURCE VALIDATOR - KEY FEATURE FOR CURVE VERIFICATION
# ═══════════════════════════════════════════════════════════════════════════════

class SourceValidator:
    """
    Handles right-click validation of curves by opening source files.
    
    When user right-clicks a curve:
    1. Identifies which candidate/bearing/direction the curve belongs to
    2. Opens the source CSV file in Excel with the relevant column highlighted
    3. Opens the associated PNG image for visual verification
    """
    
    def __init__(self, data_folder, file_metadata, csv_data):
        self.data_folder = data_folder
        self.file_metadata = file_metadata
        self.csv_data = csv_data
        self.curve_registry = {}  # Maps line objects to source info
    
    def register_curve(self, line_obj, source_info):
        """
        Register a plotted curve with its source information.
        
        source_info = {
            'file_number': int,
            'csv_path': Path,
            'image_path': Path,
            'candidate': int,
            'bearing': str,
            'direction': str,
            'order': str,
            'column_index': int,  # Column in CSV for highlighting
            'data_type': 'magnitude' or 'phase'
        }
        """
        self.curve_registry[id(line_obj)] = source_info
        return source_info
    
    def get_source_info(self, line_obj):
        """Get source info for a line object"""
        return self.curve_registry.get(id(line_obj))
    
    def find_curve_at_point(self, ax, x, y, lines):
        """Find which curve is closest to a click point"""
        min_dist = float('inf')
        closest_line = None
        
        for line in lines:
            if line not in [id(l) for l in self.curve_registry]:
                continue
            
            xdata = line.get_xdata()
            ydata = line.get_ydata()
            
            if len(xdata) == 0:
                continue
            
            # Find closest point on this line
            try:
                # Transform to display coordinates
                transform = ax.transData
                click_display = transform.transform((x, y))
                
                for px, py in zip(xdata, ydata):
                    point_display = transform.transform((px, py))
                    dist = np.sqrt((click_display[0] - point_display[0])**2 + 
                                  (click_display[1] - point_display[1])**2)
                    if dist < min_dist:
                        min_dist = dist
                        closest_line = line
            except:
                pass
        
        # Only return if within reasonable distance (30 pixels)
        if min_dist < 30:
            return closest_line
        return None
    
    def open_source_files(self, source_info):
        """
        Open BOTH the source CSV in Excel AND the associated image.
        Use open_csv_only() or open_image_only() for separate operations.
        """
        if not source_info:
            messagebox.showwarning("No Source", "Could not identify source for this curve")
            return
        
        # Open CSV first
        self.open_csv_only(source_info)
        
        # Then open image
        self.open_image_only(source_info)
    
    def open_csv_only(self, source_info):
        """Open ONLY the CSV file in Excel (no image)"""
        if not source_info:
            messagebox.showwarning("No Source", "Could not identify source for this curve")
            return
        
        csv_path = source_info.get('csv_path')
        if csv_path and csv_path.exists():
            self._open_csv_in_excel(csv_path, source_info)
        else:
            messagebox.showwarning("File Not Found", f"CSV file not found:\n{csv_path}")
    
    def open_image_only(self, source_info):
        """Open ONLY the image file for the CORRECT candidate"""
        if not source_info:
            messagebox.showwarning("No Source", "Could not identify source for this curve")
            return
        
        csv_path = source_info.get('csv_path')
        candidate = source_info.get('candidate', 1)
        
        # Build the correct image path for THIS candidate
        if self.data_folder and csv_path:
            stem = csv_path.stem
            # Use the ACTUAL candidate number, not defaulting to 1
            correct_image = Path(self.data_folder) / f"{stem}_Candidate{candidate:06d}.png"
            
            if correct_image.exists():
                self._open_file(correct_image)
            else:
                # Show error with the path we tried
                messagebox.showwarning(
                    "Image Not Found", 
                    f"Could not find image for Candidate {candidate}:\n{correct_image.name}\n\n"
                    f"Make sure the image file exists in:\n{self.data_folder}"
                )
    
    def _open_file(self, filepath):
        """Open a file with the default application"""
        try:
            filepath = str(filepath)
            if platform.system() == 'Windows':
                os.startfile(filepath)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', filepath], check=True)
            else:  # Linux
                subprocess.run(['xdg-open', filepath], check=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{filepath}\n\nError: {e}")
    
    def _open_csv_in_excel(self, csv_path, source_info):
        """
        Open CSV in Excel and try to highlight the relevant column.
        Uses VBScript on Windows for Excel automation.
        """
        candidate = source_info.get('candidate', 1)
        data_type = source_info.get('data_type', 'magnitude')

        # Calculate which row the data is in
        # CSV structure: Row 7 is frequency header, Row 8 is empty, candidates start at row 9
        # Each candidate has 5 rows: real, imaginary, magnitude, phase, empty separator
        # So Candidate 1 starts at row 9, Candidate 2 at row 14, etc.
        base_row = 9 + (candidate - 1) * 5

        if data_type == 'real':
            target_row = base_row
        elif data_type == 'imaginary':
            target_row = base_row + 1
        elif data_type == 'magnitude':
            target_row = base_row + 2
        elif data_type == 'phase':
            target_row = base_row + 3
        else:
            target_row = base_row + 2  # Default to magnitude
        
        if platform.system() == 'Windows':
            try:
                # Create VBScript to open Excel and select the row
                vbs_content = f'''
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("{str(csv_path).replace(chr(92), chr(92)+chr(92))}")
objExcel.Rows({target_row}).Select
objExcel.ActiveWindow.ScrollRow = {max(1, target_row - 5)}
'''
                vbs_path = Path(csv_path).parent / "_temp_open_excel.vbs"
                with open(vbs_path, 'w') as f:
                    f.write(vbs_content)
                
                subprocess.run(['wscript', str(vbs_path)], check=True)
                
                # Clean up VBS file after a delay
                try:
                    import time
                    time.sleep(2)
                    vbs_path.unlink()
                except:
                    pass
                    
            except Exception as e:
                # Fallback: just open the file
                self._open_file(csv_path)
        else:
            # On Mac/Linux, just open the file
            self._open_file(csv_path)
    
    def show_source_info_dialog(self, parent, source_info):
        """Show a dialog with source information"""
        if not source_info:
            return
        
        dialog = ctk.CTkToplevel(parent) if HAS_CTK else tk.Toplevel(parent)
        dialog.title("📋 Curve Source Information")
        dialog.geometry("500x400")
        dialog.transient(parent)
        
        if HAS_CTK:
            dialog.configure(fg_color=Theme.BG_PRIMARY)
        
        # Content frame
        content = ctk.CTkFrame(dialog, fg_color=Theme.BG_SECONDARY, corner_radius=12) if HAS_CTK else tk.Frame(dialog)
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title = ctk.CTkLabel(
            content, 
            text="🔍 Source Validation",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=18, weight="bold"),
            text_color=Theme.TEXT_PRIMARY
        ) if HAS_CTK else tk.Label(content, text="Source Validation", font=("Arial", 14, "bold"))
        title.pack(pady=(15, 20))
        
        # Info display
        info_items = [
            ("Bearing", source_info.get('bearing', 'N/A')),
            ("Direction", source_info.get('direction', 'N/A')),
            ("Order", source_info.get('order', 'N/A')),
            ("Candidate", f"#{source_info.get('candidate', 'N/A')}"),
            ("Data Type", source_info.get('data_type', 'N/A').capitalize()),
            ("CSV File", Path(source_info.get('csv_path', 'N/A')).name if source_info.get('csv_path') else 'N/A'),
        ]
        
        for label, value in info_items:
            row = ctk.CTkFrame(content, fg_color="transparent") if HAS_CTK else tk.Frame(content)
            row.pack(fill="x", padx=20, pady=4)
            
            lbl = ctk.CTkLabel(row, text=f"{label}:", font=ctk.CTkFont(size=12), 
                              text_color=Theme.TEXT_SECONDARY, width=100, anchor="w") if HAS_CTK else tk.Label(row, text=f"{label}:")
            lbl.pack(side="left")
            
            val = ctk.CTkLabel(row, text=str(value), font=ctk.CTkFont(size=12, weight="bold"),
                              text_color=Theme.TEXT_PRIMARY, anchor="w") if HAS_CTK else tk.Label(row, text=str(value))
            val.pack(side="left", padx=10)
        
        # Buttons
        btn_frame = ctk.CTkFrame(content, fg_color="transparent") if HAS_CTK else tk.Frame(content)
        btn_frame.pack(fill="x", padx=20, pady=20)
        
        open_csv_btn = ctk.CTkButton(
            btn_frame, text="📊 Open CSV in Excel",
            command=lambda: self.open_source_files(source_info),
            fg_color=Theme.ACCENT_PRIMARY,
            hover_color="#0056b3",
            height=36
        ) if HAS_CTK else tk.Button(btn_frame, text="Open CSV", command=lambda: self.open_source_files(source_info))
        open_csv_btn.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        open_img_btn = ctk.CTkButton(
            btn_frame, text="🖼️ Open Image",
            command=lambda: self._open_file(source_info.get('image_path')) if source_info.get('image_path') else None,
            fg_color=Theme.ACCENT_SECONDARY,
            hover_color="#28a745",
            height=36
        ) if HAS_CTK else tk.Button(btn_frame, text="Open Image")
        open_img_btn.pack(side="left", expand=True, fill="x", padx=(5, 0))
        
        # Close button
        close_btn = ctk.CTkButton(
            content, text="Close",
            command=dialog.destroy,
            fg_color=Theme.BG_HOVER,
            hover_color=Theme.BORDER_DEFAULT,
            text_color=Theme.TEXT_PRIMARY,
            height=32
        ) if HAS_CTK else tk.Button(content, text="Close", command=dialog.destroy)
        close_btn.pack(pady=(0, 15))

    def open_csv_with_band(self, source_info):
        """
        Open CSV in Excel and highlight rows within the frequency band range.
        Used for Scalar mode validation where we want to show which frequency
        rows contributed to the RMS/Peak calculation.
        """
        if not source_info:
            messagebox.showwarning("No Source", "Could not identify source for this bar")
            return

        csv_path = source_info.get('csv_path')
        if not csv_path or not csv_path.exists():
            messagebox.showwarning("File Not Found", f"CSV file not found:\n{csv_path}")
            return

        candidate = source_info.get('candidate', 1)
        freq_band_low = source_info.get('freq_band_low', 0)
        freq_band_high = source_info.get('freq_band_high', 1000)
        band_label = source_info.get('band_label', 'Unknown')

        # Calculate which row the candidate's magnitude data is in
        # CSV structure: Row 7 is frequency header, Row 8 is empty
        # Each candidate has 5 rows: real, imaginary, magnitude, phase, empty
        # Candidate 1 starts at row 9
        base_row = 9 + (candidate - 1) * 5
        magnitude_row = base_row + 2  # Magnitude row

        # Read the CSV to find which columns fall within the frequency band
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            # Frequency values are in row 7 (index 6), starting from column B
            if len(lines) >= 7:
                freq_line = lines[6]  # Row 7 (0-indexed as 6)
                freq_parts = freq_line.strip().split(',')

                # Find column indices where frequency is within band
                start_col = None
                end_col = None
                for col_idx, val in enumerate(freq_parts[1:], start=2):  # Skip label column, Excel columns start at 1
                    try:
                        freq_val = float(val)
                        if freq_val >= freq_band_low and start_col is None:
                            start_col = col_idx
                        if freq_val < freq_band_high:
                            end_col = col_idx
                    except:
                        continue

                if start_col is None:
                    start_col = 2
                if end_col is None:
                    end_col = start_col + 10

        except Exception as e:
            debug_print(f"Error reading CSV for band detection: {e}", "ERROR")
            start_col = 2
            end_col = 50  # Default range

        if platform.system() == 'Windows':
            try:
                # Convert column numbers to Excel letters
                def col_to_letter(col):
                    result = ""
                    while col > 0:
                        col, remainder = divmod(col - 1, 26)
                        result = chr(65 + remainder) + result
                    return result

                start_letter = col_to_letter(start_col)
                end_letter = col_to_letter(end_col)

                # Create VBScript to open Excel and highlight the frequency band range
                vbs_content = f'''
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("{str(csv_path).replace(chr(92), chr(92)+chr(92))}")
Set ws = objWorkbook.Sheets(1)

' First scroll to the magnitude row for this candidate
objExcel.ActiveWindow.ScrollRow = {max(1, magnitude_row - 3)}

' Select the frequency band range for magnitude row (row {magnitude_row})
ws.Range("{start_letter}{magnitude_row}:{end_letter}{magnitude_row}").Select

' Highlight the selected cells with yellow background
ws.Range("{start_letter}{magnitude_row}:{end_letter}{magnitude_row}").Interior.Color = RGB(255, 255, 150)

' Also highlight frequency row to show which frequencies
ws.Range("{start_letter}7:{end_letter}7").Interior.Color = RGB(200, 230, 255)
'''
                vbs_path = Path(csv_path).parent / "_temp_open_excel_band.vbs"
                with open(vbs_path, 'w') as f:
                    f.write(vbs_content)

                subprocess.run(['wscript', str(vbs_path)], check=True)

                # Clean up VBS file after a delay
                try:
                    import time
                    time.sleep(2)
                    vbs_path.unlink()
                except:
                    pass

            except Exception as e:
                debug_print(f"VBScript error: {e}", "ERROR")
                # Fallback: just open the file
                self._open_file(csv_path)
        else:
            # On Mac/Linux, just open the file
            self._open_file(csv_path)


# ═══════════════════════════════════════════════════════════════════════════════
# CUSTOM WIDGETS - APPLE TV STYLE
# ═══════════════════════════════════════════════════════════════════════════════

class CollapsiblePanel(ctk.CTkFrame if HAS_CTK else tk.Frame):
    """Elegant collapsible panel with modern styling"""
    
    def __init__(self, parent, title, expanded=True, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.expanded = expanded
        self.title = title
        
        if HAS_CTK:
            self.configure(fg_color=Theme.BG_CARD, corner_radius=10)
        
        # Header
        self.header = ctk.CTkFrame(self, fg_color="transparent") if HAS_CTK else tk.Frame(self)
        self.header.pack(fill="x", padx=12, pady=10)
        
        self.toggle_btn = ctk.CTkButton(
            self.header,
            text=f"{'▼' if expanded else '▶'}  {title}",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=13, weight="bold") if HAS_CTK else None,
            fg_color="transparent",
            hover_color=Theme.BG_HOVER,
            text_color=Theme.TEXT_PRIMARY,
            anchor="w",
            command=self.toggle,
            height=28
        ) if HAS_CTK else tk.Button(self.header, text=title, command=self.toggle)
        self.toggle_btn.pack(fill="x")
        
        # Content
        self.content = ctk.CTkFrame(self, fg_color="transparent") if HAS_CTK else tk.Frame(self)
        if expanded:
            self.content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    
    def toggle(self):
        self.expanded = not self.expanded
        arrow = "▼" if self.expanded else "▶"
        self.toggle_btn.configure(text=f"{arrow}  {self.title}")
        
        if self.expanded:
            self.content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        else:
            self.content.pack_forget()


class StatusBar(ctk.CTkFrame if HAS_CTK else tk.Frame):
    """Clean status bar with progress bar"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        if HAS_CTK:
            self.configure(fg_color=Theme.BG_SECONDARY, height=40, corner_radius=0)
        
        # Left - status
        self.status_label = ctk.CTkLabel(
            self, text="Ready", 
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=11),
            text_color=Theme.TEXT_SECONDARY
        ) if HAS_CTK else tk.Label(self, text="Ready")
        self.status_label.pack(side="left", padx=15)
        
        # Progress bar (hidden by default)
        self.progress_bar = ctk.CTkProgressBar(
            self, width=200, height=12,
            fg_color=Theme.BG_TERTIARY,
            progress_color=Theme.ACCENT_PRIMARY
        ) if HAS_CTK else ttk.Progressbar(self, length=200, mode='determinate')
        # Don't pack yet - will show when needed
        
        # Right - coordinates
        self.coord_label = ctk.CTkLabel(
            self, 
            text="X: ─── Hz  Y: ─── N",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY_MONO, size=11),
            text_color=Theme.TEXT_MUTED
        ) if HAS_CTK else tk.Label(self, text="X: --- Hz  Y: --- N")
        self.coord_label.pack(side="right", padx=15)
        
        # Center - hint
        self.hint_label = ctk.CTkLabel(
            self, text="💡 Right-click any curve to validate source",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=10),
            text_color=Theme.ACCENT_PRIMARY
        ) if HAS_CTK else tk.Label(self, text="Right-click curve to validate")
        self.hint_label.pack(side="right", padx=20)
    
    def set_status(self, text, color=None):
        self.status_label.configure(text=text)
        if color:
            self.status_label.configure(text_color=color)
    
    def show_progress(self, value=0):
        """Show progress bar with value 0-1"""
        self.hint_label.pack_forget()
        self.progress_bar.pack(side="left", padx=10)
        if HAS_CTK:
            self.progress_bar.set(value)
        else:
            self.progress_bar['value'] = value * 100
    
    def update_progress(self, value):
        """Update progress bar value 0-1"""
        if HAS_CTK:
            self.progress_bar.set(value)
        else:
            self.progress_bar['value'] = value * 100
    
    def hide_progress(self):
        """Hide progress bar"""
        self.progress_bar.pack_forget()
        self.hint_label.pack(side="right", padx=20)
    
    def set_coordinates(self, x, y, unit_y="N"):
        if x is not None and y is not None:
            self.coord_label.configure(text=f"X: {x:.1f} Hz  Y: {y:.4e} {unit_y}")
        else:
            self.coord_label.configure(text=f"X: ─── Hz  Y: ─── {unit_y}")


# ═══════════════════════════════════════════════════════════════════════════════
# INTERACTIVE GRAPH TRACKER
# ═══════════════════════════════════════════════════════════════════════════════

class GraphTracker:
    """Graph interaction with crosshairs and right-click validation"""
    
    def __init__(self, fig, canvas, status_bar, source_validator=None):
        self.fig = fig
        self.canvas = canvas
        self.status_bar = status_bar
        self.source_validator = source_validator
        self.axes_list = []
        
        # Crosshair elements
        self.vlines = {}
        self.hlines = {}
        self.annotations = {}
        
        # Data tracking
        self.current_data = {}
        self.line_to_source = {}  # Maps line id to source info
        
        # State
        self.tracking_enabled = True
        self.snap_to_data = True
        self.snap_radius = 50

        # Highlight state for showing which curve was clicked
        self.highlighted_line = None
        self.original_linewidth = None
        self.original_alpha = None

        # Connect events
        self.cid_motion = canvas.mpl_connect('motion_notify_event', self.on_motion)
        self.cid_click = canvas.mpl_connect('button_press_event', self.on_click)
        self.cid_leave = canvas.mpl_connect('axes_leave_event', self.on_leave)
    
    def setup_crosshairs(self, axes_list):
        """Initialize crosshairs for all subplots"""
        self.axes_list = axes_list
        self.vlines = {}
        self.hlines = {}
        self.annotations = {}
        
        for ax in axes_list:
            vline = ax.axvline(x=0, color=Theme.ACCENT_PRIMARY, 
                              linestyle='--', alpha=0.6, linewidth=1, visible=False)
            self.vlines[ax] = vline
            
            hline = ax.axhline(y=0, color=Theme.ACCENT_PURPLE,
                              linestyle=':', alpha=0.4, linewidth=1, visible=False)
            self.hlines[ax] = hline
            
            annot = ax.annotate('', xy=(0, 0), xytext=(10, 10),
                               textcoords='offset points', fontsize=9,
                               color=Theme.TEXT_PRIMARY,
                               bbox=dict(boxstyle='round,pad=0.4',
                                        facecolor=Theme.BG_CARD,
                                        edgecolor=Theme.ACCENT_PRIMARY,
                                        alpha=0.95),
                               visible=False)
            self.annotations[ax] = annot
    
    def register_line(self, ax, line_obj, freq, values, label, color, source_info):
        """Register a line with its data and source info"""
        if ax not in self.current_data:
            self.current_data[ax] = []
        
        self.current_data[ax].append({
            'line': line_obj,
            'freq': np.array(freq),
            'values': np.array(values),
            'label': label,
            'color': color,
            'source_info': source_info
        })
        
        self.line_to_source[id(line_obj)] = source_info
    
    def clear_data(self):
        self.current_data = {}
        self.line_to_source = {}
    
    def find_nearest_point(self, ax, x, y):
        """Find nearest data point"""
        if ax not in self.current_data:
            return None
        
        min_dist = float('inf')
        nearest = None
        
        transform = ax.transData
        
        for data in self.current_data[ax]:
            freq = data['freq']
            values = data['values']
            
            for i, (fx, fy) in enumerate(zip(freq, values)):
                try:
                    cursor_display = transform.transform((x, y))
                    point_display = transform.transform((fx, fy))
                    dist = np.sqrt((cursor_display[0] - point_display[0])**2 + 
                                  (cursor_display[1] - point_display[1])**2)
                    
                    if dist < min_dist and dist < self.snap_radius:
                        min_dist = dist
                        nearest = {
                            'x': fx, 'y': fy, 'index': i,
                            'label': data['label'],
                            'color': data['color'],
                            'source_info': data['source_info'],
                            'line': data['line']
                        }
                except:
                    pass
        
        return nearest
    
    def on_motion(self, event):
        """Handle mouse motion for crosshair tracking"""
        if not self.tracking_enabled or event.inaxes is None:
            self._hide_all_tracking()
            self.status_bar.set_coordinates(None, None)
            return
        
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        
        # Update crosshairs
        for ax in self.axes_list:
            if ax in self.vlines:
                self.vlines[ax].set_xdata([x, x])
                self.vlines[ax].set_visible(True)
                
                if ax == event.inaxes:
                    self.hlines[ax].set_ydata([y, y])
                    self.hlines[ax].set_visible(True)
                else:
                    self.hlines[ax].set_visible(False)
        
        # Check for nearby data point
        if self.snap_to_data:
            nearest = self.find_nearest_point(event.inaxes, x, y)
            if nearest:
                annot = self.annotations.get(event.inaxes)
                if annot:
                    annot.xy = (nearest['x'], nearest['y'])
                    annot.set_text(f"{nearest['label']}\n"
                                  f"F: {nearest['x']:.1f} Hz\n"
                                  f"V: {nearest['y']:.4e}")
                    annot.set_visible(True)
                
                self.status_bar.set_coordinates(nearest['x'], nearest['y'])
            else:
                for annot in self.annotations.values():
                    annot.set_visible(False)
                self.status_bar.set_coordinates(x, y)
        else:
            self.status_bar.set_coordinates(x, y)
        
        self.canvas.draw_idle()
    
    def _highlight_curve(self, line):
        """Highlight a curve by making it thicker and bringing to front"""
        # First, unhighlight any previously highlighted line
        self._unhighlight_curve()
        
        if line is not None:
            self.highlighted_line = line
            self.original_linewidth = line.get_linewidth()
            self.original_alpha = line.get_alpha()
            
            # Make the line stand out
            line.set_linewidth(4)
            line.set_alpha(1.0)
            line.set_zorder(1000)  # Bring to front
            
            self.canvas.draw_idle()
    
    def _unhighlight_curve(self):
        """Restore highlighted curve to original state"""
        if self.highlighted_line is not None:
            try:
                self.highlighted_line.set_linewidth(self.original_linewidth or 1.2)
                self.highlighted_line.set_alpha(self.original_alpha or 0.85)
                self.highlighted_line.set_zorder(1)
            except:
                pass
            self.highlighted_line = None
            self.canvas.draw_idle()
    
    def on_click(self, event):
        """Handle click - LEFT for info, RIGHT for source validation with highlighting"""
        if event.inaxes is None:
            self._unhighlight_curve()
            return
        
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        
        nearest = self.find_nearest_point(event.inaxes, x, y)
        
        # RIGHT CLICK - button 3 on most systems, button 2 on some Windows configurations
        if event.button == 3:  # RIGHT CLICK - VALIDATION
            if nearest and self.source_validator:
                source_info = nearest.get('source_info')
                line = nearest.get('line')
                if source_info:
                    # HIGHLIGHT the curve so user can see which one was selected
                    self._highlight_curve(line)
                    # Show context menu
                    self._show_context_menu(event, source_info, line)
            else:
                self._unhighlight_curve()
        elif event.button == 1:  # LEFT CLICK
            if nearest:
                line = nearest.get('line')
                self._highlight_curve(line)
                info = f"📍 {nearest['label']} | F: {nearest['x']:.1f} Hz | V: {nearest['y']:.6e}"
                self.status_bar.set_status(info, Theme.ACCENT_PRIMARY)
            else:
                self._unhighlight_curve()
    
    def _show_context_menu(self, event, source_info, line=None):
        """Show right-click context menu for source validation - curve is highlighted"""
        menu = Menu(self.canvas.get_tk_widget(), tearoff=0)
        
        # Get candidate number for display
        candidate = source_info.get('candidate', '?')
        bearing = source_info.get('bearing', '?')
        direction = source_info.get('direction', '?')
        
        # Header showing which curve is selected (the highlighted one)
        menu.add_command(
            label=f"▶ SELECTED: Candidate {candidate} ({bearing}-{direction})",
            state="disabled"
        )
        menu.add_separator()
        
        # Menu items - SEPARATE actions
        # IMPORTANT: Use default argument (si=source_info) to capture by VALUE, not reference
        # This fixes the bug where wrong candidate was being opened in Excel
        menu.add_command(
            label="📊 Open CSV in Excel (highlight row)",
            command=lambda si=source_info: self.source_validator.open_csv_only(si)
        )
        menu.add_command(
            label=f"🖼️ Open Image (Candidate {candidate})",
            command=lambda si=source_info: self.source_validator.open_image_only(si)
        )
        menu.add_separator()
        menu.add_command(
            label="📊 + 🖼️ Open Both CSV and Image",
            command=lambda si=source_info: self.source_validator.open_source_files(si)
        )
        menu.add_separator()
        menu.add_command(
            label="ℹ️ Show Source Details",
            command=lambda si=source_info: self.source_validator.show_source_info_dialog(
                self.canvas.get_tk_widget().winfo_toplevel(), si)
        )
        menu.add_separator()
        menu.add_command(
            label="✖ Clear Highlight",
            command=self._unhighlight_curve
        )
        
        # Show menu at mouse position (original simple method)
        try:
            menu.tk_popup(event.guiEvent.x_root, event.guiEvent.y_root)
        finally:
            menu.grab_release()
    
    def on_leave(self, event):
        self._hide_all_tracking()
        self.canvas.draw_idle()
    
    def _hide_all_tracking(self):
        for vline in self.vlines.values():
            vline.set_visible(False)
        for hline in self.hlines.values():
            hline.set_visible(False)
        for annot in self.annotations.values():
            annot.set_visible(False)
    
    def disconnect(self):
        self.canvas.mpl_disconnect(self.cid_motion)
        self.canvas.mpl_disconnect(self.cid_click)
        self.canvas.mpl_disconnect(self.cid_leave)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════

class BearingForceViewer:
    """
    Bearing Force Bode Plot Viewer
    
    Key Features:
    - RIGHT-CLICK VALIDATION: Click any curve to open source CSV and image
    - Clean light theme
    - Interactive crosshair tracking
    - Source file traceability
    """
    
    def __init__(self, root):
        self.root = root
        self.root.title("Bearing Force Bode Plot Viewer")
        self.root.geometry("1600x1000")
        self.root.minsize(1200, 800)
        
        # Configure light theme
        if HAS_CTK:
            self.root.configure(fg_color=Theme.BG_SECONDARY)
        else:
            self.root.configure(bg=Theme.BG_SECONDARY)
        
        # Apply matplotlib style
        plt.rcParams.update(Theme.MPL_STYLE)
        
        # Data storage
        self.data_folder = None
        self.file_metadata = {}
        self.csv_data = {}
        self.candidate_count = 0
        
        # Options
        self.bearings = []
        self.directions = []
        self.orders = []
        self.stages = []
        self.torques = []
        self.conditions = []
        
        # Source validation
        self.source_validator = None
        self.graph_tracker = None
        
        # Build UI
        self.setup_ui()
    
    def setup_ui(self):
        """Build the Clean UI with RESIZABLE sidebar"""
        
        # Main container using PanedWindow for resizable split
        # This allows dragging the divider between sidebar and plot area
        self.paned = tk.PanedWindow(
            self.root, 
            orient=tk.HORIZONTAL,
            sashwidth=8,
            sashrelief=tk.FLAT,
            bg=Theme.BORDER_DEFAULT,
            opaqueresize=True
        )
        self.paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        # ─── LEFT SIDEBAR (RESIZABLE) ───
        self.sidebar = ctk.CTkFrame(
            self.paned, 
            fg_color=Theme.BG_PRIMARY,
            corner_radius=12
        ) if HAS_CTK else tk.Frame(self.paned, bg=Theme.BG_PRIMARY)
        
        # Add sidebar to paned window with minimum size
        self.paned.add(self.sidebar, minsize=280, width=360)
        
        # App title
        header = ctk.CTkFrame(self.sidebar, fg_color="transparent") if HAS_CTK else tk.Frame(self.sidebar)
        header.pack(fill="x", padx=15, pady=15)
        
        title_label = ctk.CTkLabel(
            header,
            text="Bearing Force Viewer",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=18, weight="bold"),
            text_color=Theme.TEXT_PRIMARY
        ) if HAS_CTK else tk.Label(header, text="Bearing Force Viewer", font=("Arial", 16, "bold"))
        title_label.pack(anchor="w")
        
        subtitle = ctk.CTkLabel(
            header,
            text="Romax DOE Analysis  •  ↔ Drag edge to resize",
            font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=11),
            text_color=Theme.TEXT_SECONDARY
        ) if HAS_CTK else tk.Label(header, text="Romax DOE Analysis")
        subtitle.pack(anchor="w")
        
        # Scrollable content
        self.sidebar_scroll = ctk.CTkScrollableFrame(
            self.sidebar,
            fg_color="transparent"
        ) if HAS_CTK else tk.Frame(self.sidebar)
        self.sidebar_scroll.pack(fill="both", expand=True, padx=5)
        
        # ─── DATA SOURCE ───
        self.data_panel = CollapsiblePanel(self.sidebar_scroll, "Data Source", expanded=True)
        self.data_panel.pack(fill="x", pady=6)
        
        self.folder_var = ctk.StringVar() if HAS_CTK else tk.StringVar()
        self.folder_entry = ctk.CTkEntry(
            self.data_panel.content,
            textvariable=self.folder_var,
            placeholder_text="Select data folder...",
            height=38,
            fg_color=Theme.BG_SECONDARY,
            border_color=Theme.BORDER_DEFAULT,
            text_color=Theme.TEXT_PRIMARY,
            corner_radius=8
        ) if HAS_CTK else tk.Entry(self.data_panel.content, textvariable=self.folder_var)
        self.folder_entry.pack(fill="x", pady=(0, 10))
        
        btn_row = ctk.CTkFrame(self.data_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.data_panel.content)
        btn_row.pack(fill="x")
        
        self.browse_btn = ctk.CTkButton(
            btn_row, text="Browse",
            command=self.browse_folder,
            height=36,
            fg_color=Theme.BG_SECONDARY,
            hover_color=Theme.BG_HOVER,
            text_color=Theme.TEXT_PRIMARY,
            corner_radius=8
        ) if HAS_CTK else tk.Button(btn_row, text="Browse", command=self.browse_folder)
        self.browse_btn.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        self.load_btn = ctk.CTkButton(
            btn_row, text="Load Data",
            command=self.load_data,
            height=36,
            fg_color=Theme.ACCENT_PRIMARY,
            hover_color="#0056b3",
            text_color="#ffffff",
            corner_radius=8
        ) if HAS_CTK else tk.Button(btn_row, text="Load", command=self.load_data)
        self.load_btn.pack(side="left", expand=True, fill="x", padx=(5, 0))
        
        # ─── FILTERS ───
        self.filter_panel = CollapsiblePanel(self.sidebar_scroll, "Filters", expanded=True)
        self.filter_panel.pack(fill="x", pady=6)

        # Stage
        self._create_filter_combo(self.filter_panel.content, "Stage", "stage")
        # Condition
        self._create_filter_combo(self.filter_panel.content, "Condition", "condition")
        # Order
        self._create_filter_combo(self.filter_panel.content, "Order", "order")

        # ─── TORQUES (multi-select checkboxes) ───
        self.torque_panel = CollapsiblePanel(self.sidebar_scroll, "Torques", expanded=True)
        self.torque_panel.pack(fill="x", pady=6)

        self.torque_frame = ctk.CTkFrame(self.torque_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.torque_panel.content)
        self.torque_frame.pack(fill="x")
        self.torque_vars = {}  # Multi-select torques
        self.torque_checks = {}

        self.torque_placeholder = ctk.CTkLabel(
            self.torque_frame,
            text="Load data to see torques",
            font=ctk.CTkFont(size=11),
            text_color=Theme.TEXT_MUTED
        ) if HAS_CTK else tk.Label(self.torque_frame, text="Load data to see torques")
        self.torque_placeholder.pack(pady=10)

        # Keep torque_var for backward compatibility (set to first selected torque)
        self.torque_var = ctk.StringVar() if HAS_CTK else tk.StringVar()

        # ─── BEARINGS ───
        self.bearing_panel = CollapsiblePanel(self.sidebar_scroll, "Bearings", expanded=True)
        self.bearing_panel.pack(fill="x", pady=6)
        
        self.bearing_frame = ctk.CTkFrame(self.bearing_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.bearing_panel.content)
        self.bearing_frame.pack(fill="x")
        self.bearing_vars = {}
        self.bearing_checks = {}
        
        self.bearing_placeholder = ctk.CTkLabel(
            self.bearing_frame,
            text="Load data to see bearings",
            font=ctk.CTkFont(size=11),
            text_color=Theme.TEXT_MUTED
        ) if HAS_CTK else tk.Label(self.bearing_frame, text="Load data to see bearings")
        self.bearing_placeholder.pack(pady=10)
        
        # ─── DIRECTIONS ───
        self.direction_panel = CollapsiblePanel(self.sidebar_scroll, "Force Directions", expanded=True)
        self.direction_panel.pack(fill="x", pady=6)
        
        dir_container = ctk.CTkFrame(self.direction_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.direction_panel.content)
        dir_container.pack(fill="x")
        
        self.direction_vars = {}
        self.direction_checks = {}
        
        dir_row1 = ctk.CTkFrame(dir_container, fg_color="transparent") if HAS_CTK else tk.Frame(dir_container)
        dir_row1.pack(fill="x", pady=3)
        dir_row2 = ctk.CTkFrame(dir_container, fg_color="transparent") if HAS_CTK else tk.Frame(dir_container)
        dir_row2.pack(fill="x", pady=3)
        
        for i, d in enumerate(['X', 'Y', 'Z', 'Mx', 'My', 'Mz']):
            var = ctk.BooleanVar(value=False) if HAS_CTK else tk.BooleanVar(value=False)
            self.direction_vars[d] = var
            
            parent = dir_row1 if i < 3 else dir_row2
            
            cb = ctk.CTkCheckBox(
                parent, text=d, variable=var,
                font=ctk.CTkFont(size=12),
                fg_color=Theme.ACCENT_PRIMARY,
                hover_color=Theme.ACCENT_PRIMARY,
                border_color=Theme.BORDER_DEFAULT,
                text_color=Theme.TEXT_SECONDARY,
                corner_radius=4
            ) if HAS_CTK else tk.Checkbutton(parent, text=d, variable=var)
            cb.pack(side="left", padx=10, pady=2)
            self.direction_checks[d] = cb
        
        # ─── CANDIDATES ───
        self.candidate_panel = CollapsiblePanel(self.sidebar_scroll, "Candidates", expanded=True)
        self.candidate_panel.pack(fill="x", pady=6)
        
        self.candidate_mode = ctk.StringVar(value="all") if HAS_CTK else tk.StringVar(value="all")
        
        mode_frame = ctk.CTkFrame(self.candidate_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.candidate_panel.content)
        mode_frame.pack(fill="x", pady=5)
        
        all_radio = ctk.CTkRadioButton(
            mode_frame, text="All Candidates", variable=self.candidate_mode, value="all",
            command=self.on_candidate_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=Theme.ACCENT_PRIMARY,
            text_color=Theme.TEXT_PRIMARY
        ) if HAS_CTK else tk.Radiobutton(mode_frame, text="All", variable=self.candidate_mode, value="all")
        all_radio.pack(anchor="w")
        
        select_radio = ctk.CTkRadioButton(
            mode_frame, text="Select Specific", variable=self.candidate_mode, value="select",
            command=self.on_candidate_mode_change,
            font=ctk.CTkFont(size=12),
            fg_color=Theme.ACCENT_PRIMARY,
            text_color=Theme.TEXT_PRIMARY
        ) if HAS_CTK else tk.Radiobutton(mode_frame, text="Select", variable=self.candidate_mode, value="select")
        select_radio.pack(anchor="w")
        
        self.candidate_entry = ctk.CTkEntry(
            self.candidate_panel.content,
            placeholder_text="e.g., 1,2,3 or 1-10",
            height=34,
            fg_color=Theme.BG_SECONDARY,
            border_color=Theme.BORDER_DEFAULT,
            text_color=Theme.TEXT_PRIMARY,
            state="disabled",
            corner_radius=6
        ) if HAS_CTK else tk.Entry(self.candidate_panel.content)
        self.candidate_entry.pack(fill="x", pady=5)
        
        self.cand_count_label = ctk.CTkLabel(
            self.candidate_panel.content,
            text="(0 candidates)",
            font=ctk.CTkFont(size=11),
            text_color=Theme.TEXT_MUTED
        ) if HAS_CTK else tk.Label(self.candidate_panel.content, text="(0 candidates)")
        self.cand_count_label.pack(anchor="w")
        
        # ─── PLOT OPTIONS ───
        self.plot_panel = CollapsiblePanel(self.sidebar_scroll, "Plot Options", expanded=True)
        self.plot_panel.pack(fill="x", pady=6)

        # ═══════════════════════════════════════════════════════════════════════
        # Output Mode: Scalar vs Dynamic
        # ═══════════════════════════════════════════════════════════════════════
        mode_frame = ctk.CTkFrame(self.plot_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.plot_panel.content)
        mode_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(mode_frame, text="Output Mode", font=ctk.CTkFont(size=11, weight="bold"),
                    text_color=Theme.ACCENT_PRIMARY).pack(anchor="w") if HAS_CTK else None

        self.output_mode = ctk.StringVar(value="dynamic") if HAS_CTK else tk.StringVar(value="dynamic")

        mode_btns = ctk.CTkFrame(mode_frame, fg_color="transparent") if HAS_CTK else tk.Frame(mode_frame)
        mode_btns.pack(fill="x", pady=2)

        for text, value in [("Dynamic", "dynamic"), ("Scalar", "scalar")]:
            btn = ctk.CTkRadioButton(
                mode_btns, text=text, variable=self.output_mode, value=value,
                font=ctk.CTkFont(size=11),
                fg_color=Theme.ACCENT_PRIMARY,
                text_color=Theme.TEXT_PRIMARY,
                command=self.on_output_mode_change
            ) if HAS_CTK else tk.Radiobutton(mode_btns, text=text, variable=self.output_mode, value=value)
            btn.pack(side="left", padx=8)

        # Separator
        ctk.CTkFrame(self.plot_panel.content, height=1, fg_color=Theme.BORDER_DEFAULT).pack(fill="x", pady=8) if HAS_CTK else None

        # Plot type
        pt_frame = ctk.CTkFrame(self.plot_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.plot_panel.content)
        pt_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(pt_frame, text="Plot Type", font=ctk.CTkFont(size=11), 
                    text_color=Theme.TEXT_SECONDARY).pack(anchor="w") if HAS_CTK else None
        
        self.plot_type = ctk.StringVar(value="magnitude") if HAS_CTK else tk.StringVar(value="magnitude")
        
        pt_btns = ctk.CTkFrame(pt_frame, fg_color="transparent") if HAS_CTK else tk.Frame(pt_frame)
        pt_btns.pack(fill="x", pady=2)
        
        for text, value in [("Magnitude", "magnitude"), ("Phase", "phase"), ("Both", "both")]:
            btn = ctk.CTkRadioButton(
                pt_btns, text=text, variable=self.plot_type, value=value,
                font=ctk.CTkFont(size=11),
                fg_color=Theme.ACCENT_PRIMARY,
                text_color=Theme.TEXT_PRIMARY
            ) if HAS_CTK else tk.Radiobutton(pt_btns, text=text, variable=self.plot_type, value=value)
            btn.pack(side="left", padx=8)
        
        # Y-Scale
        ys_frame = ctk.CTkFrame(self.plot_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.plot_panel.content)
        ys_frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(ys_frame, text="Y-Scale", font=ctk.CTkFont(size=11),
                    text_color=Theme.TEXT_SECONDARY).pack(anchor="w") if HAS_CTK else None
        
        self.y_scale = ctk.StringVar(value="log") if HAS_CTK else tk.StringVar(value="log")
        
        ys_btns = ctk.CTkFrame(ys_frame, fg_color="transparent") if HAS_CTK else tk.Frame(ys_frame)
        ys_btns.pack(fill="x", pady=2)
        
        for text, value in [("Log", "log"), ("Linear", "linear")]:
            btn = ctk.CTkRadioButton(
                ys_btns, text=text, variable=self.y_scale, value=value,
                font=ctk.CTkFont(size=11),
                fg_color=Theme.ACCENT_PRIMARY,
                text_color=Theme.TEXT_PRIMARY
            ) if HAS_CTK else tk.Radiobutton(ys_btns, text=text, variable=self.y_scale, value=value)
            btn.pack(side="left", padx=8)
        
        # Tracking options
        track_frame = ctk.CTkFrame(self.plot_panel.content, fg_color="transparent") if HAS_CTK else tk.Frame(self.plot_panel.content)
        track_frame.pack(fill="x", pady=5)
        
        self.tracking_var = ctk.BooleanVar(value=True) if HAS_CTK else tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            track_frame, text="Crosshair Tracking", variable=self.tracking_var,
            command=self.toggle_tracking,
            font=ctk.CTkFont(size=11),
            fg_color=Theme.ACCENT_PRIMARY,
            text_color=Theme.TEXT_PRIMARY
        ).pack(anchor="w") if HAS_CTK else None
        
        self.snap_var = ctk.BooleanVar(value=True) if HAS_CTK else tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            track_frame, text="Snap to Data", variable=self.snap_var,
            command=self.toggle_snap,
            font=ctk.CTkFont(size=11),
            fg_color=Theme.ACCENT_PRIMARY,
            text_color=Theme.TEXT_PRIMARY
        ).pack(anchor="w") if HAS_CTK else None
        
        # ─── ACTION BUTTONS ───
        action_frame = ctk.CTkFrame(self.sidebar_scroll, fg_color="transparent") if HAS_CTK else tk.Frame(self.sidebar_scroll)
        action_frame.pack(fill="x", pady=15)
        
        self.plot_btn = ctk.CTkButton(
            action_frame, text="Generate Plot",
            command=self.plot_data,
            height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=Theme.ACCENT_SECONDARY,
            hover_color="#28a745",
            text_color="#ffffff",
            corner_radius=10
        ) if HAS_CTK else tk.Button(action_frame, text="Plot", command=self.plot_data)
        self.plot_btn.pack(fill="x", pady=5)
        
        btn_row2 = ctk.CTkFrame(action_frame, fg_color="transparent") if HAS_CTK else tk.Frame(action_frame)
        btn_row2.pack(fill="x", pady=5)
        
        self.export_btn = ctk.CTkButton(
            btn_row2, text="Export Excel",
            command=self.export_to_excel,
            height=36,
            fg_color=Theme.BG_SECONDARY,
            hover_color=Theme.BG_HOVER,
            text_color=Theme.TEXT_PRIMARY,
            corner_radius=8
        ) if HAS_CTK else tk.Button(btn_row2, text="Export", command=self.export_to_excel)
        self.export_btn.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        self.clear_btn = ctk.CTkButton(
            btn_row2, text="Clear",
            command=self.clear_plot,
            height=36,
            fg_color=Theme.BG_SECONDARY,
            hover_color=Theme.ACCENT_ERROR,
            text_color=Theme.TEXT_PRIMARY,
            corner_radius=8
        ) if HAS_CTK else tk.Button(btn_row2, text="Clear", command=self.clear_plot)
        self.clear_btn.pack(side="left", expand=True, fill="x", padx=(5, 0))

        # Debug button row
        btn_row3 = ctk.CTkFrame(action_frame, fg_color="transparent") if HAS_CTK else tk.Frame(action_frame)
        btn_row3.pack(fill="x", pady=5)

        self.debug_btn = ctk.CTkButton(
            btn_row3, text="🔍 Export Debug Info",
            command=self.export_debug_info,
            height=32,
            fg_color="#6c757d",
            hover_color="#5a6268",
            text_color="#ffffff",
            corner_radius=6,
            font=ctk.CTkFont(size=11)
        ) if HAS_CTK else tk.Button(btn_row3, text="Debug", command=self.export_debug_info)
        self.debug_btn.pack(fill="x")
        
        # ─── MAIN CONTENT AREA (added to PanedWindow) ───
        self.content_area = ctk.CTkFrame(
            self.paned,
            fg_color=Theme.BG_PRIMARY,
            corner_radius=12
        ) if HAS_CTK else tk.Frame(self.paned, bg=Theme.BG_PRIMARY)
        
        # Add content area to paned window
        self.paned.add(self.content_area, minsize=600)
        
        # Plot container
        self.plot_container = ctk.CTkFrame(self.content_area, fg_color=Theme.BG_PRIMARY) if HAS_CTK else tk.Frame(self.content_area)
        self.plot_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Matplotlib figure
        self.fig = Figure(figsize=(12, 8), dpi=100, facecolor=Theme.BG_SECONDARY)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_container)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        
        # Hidden toolbar
        self.toolbar_frame = ctk.CTkFrame(self.plot_container, fg_color="transparent", height=0) if HAS_CTK else tk.Frame(self.plot_container)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbar_frame)
        self.toolbar.update()
        
        # Status bar
        self.status_bar = StatusBar(self.content_area)
        self.status_bar.pack(fill="x", side="bottom")
        
        # Initialize source validator (will be set after data load)
        self.source_validator = SourceValidator(None, {}, {})
        
        # Initialize graph tracker
        self.graph_tracker = GraphTracker(self.fig, self.canvas, self.status_bar, self.source_validator)
        
        # Show welcome
        self._show_welcome_screen()
    
    def _create_filter_combo(self, parent, label, attr_name):
        """Create a filter combobox"""
        frame = ctk.CTkFrame(parent, fg_color="transparent") if HAS_CTK else tk.Frame(parent)
        frame.pack(fill="x", pady=4)
        
        ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=11),
                    text_color=Theme.TEXT_SECONDARY).pack(anchor="w") if HAS_CTK else tk.Label(frame, text=label).pack(anchor="w")
        
        var = ctk.StringVar() if HAS_CTK else tk.StringVar()
        setattr(self, f"{attr_name}_var", var)
        
        combo = ctk.CTkComboBox(
            frame, variable=var, values=[],
            height=34, fg_color=Theme.BG_SECONDARY,
            border_color=Theme.BORDER_DEFAULT,
            button_color=Theme.BG_HOVER,
            dropdown_fg_color=Theme.BG_CARD,
            corner_radius=6
        ) if HAS_CTK else tk.OptionMenu(frame, var, "")
        combo.pack(fill="x")
        setattr(self, f"{attr_name}_combo", combo)
    
    def _show_welcome_screen(self):
        """Display welcome screen"""
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        ax.set_facecolor(Theme.BG_PRIMARY)
        ax.axis('off')
        
        ax.text(0.5, 0.6, "Bearing Force Bode Plot Viewer", 
               fontsize=26, fontweight='bold', color=Theme.TEXT_PRIMARY,
               ha='center', va='center', transform=ax.transAxes,
               fontfamily=Theme.FONT_FAMILY)
        
        ax.text(0.5, 0.48, "Romax DOE Frequency Response Viewer",
               fontsize=14, color=Theme.TEXT_SECONDARY,
               ha='center', va='center', transform=ax.transAxes)
        
        instructions = [
            "1. Browse and select your Romax DOE data folder",
            "2. Click 'Load Data' to parse CSV files",
            "3. Configure filters and generate plot",
            "",
            "🎯  Hover for crosshair tracking",
            "🖱️  RIGHT-CLICK any curve to validate source",
            "📊  Opens CSV in Excel with row highlighted",
            "🖼️  Opens associated image for verification"
        ]
        
        for i, line in enumerate(instructions):
            color = Theme.ACCENT_PRIMARY if "RIGHT-CLICK" in line else Theme.TEXT_MUTED
            ax.text(0.5, 0.32 - i*0.045, line,
                   fontsize=11, color=color,
                   ha='center', va='center', transform=ax.transAxes)
        
        self.canvas.draw()
    
    def toggle_tracking(self):
        if self.graph_tracker:
            self.graph_tracker.tracking_enabled = self.tracking_var.get()
    
    def toggle_snap(self):
        if self.graph_tracker:
            self.graph_tracker.snap_to_data = self.snap_var.get()
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Bearing Force Data Folder")
        if folder:
            self.folder_var.set(folder)
            self.status_bar.set_status(f"Selected: {Path(folder).name}", Theme.ACCENT_PRIMARY)
    
    def on_candidate_mode_change(self):
        if self.candidate_mode.get() == "select":
            if HAS_CTK:
                self.candidate_entry.configure(state="normal")
        else:
            if HAS_CTK:
                self.candidate_entry.configure(state="disabled")
    
    def parse_candidate_selection(self):
        if self.candidate_mode.get() == "all":
            return list(range(1, self.candidate_count + 1))
        
        selection = self.candidate_entry.get()
        candidates = set()
        for part in selection.split(','):
            part = part.strip()
            if '-' in part:
                try:
                    start, end = part.split('-')
                    candidates.update(range(int(start), int(end) + 1))
                except:
                    pass
            elif part.isdigit():
                candidates.add(int(part))
        return sorted(c for c in candidates if 1 <= c <= self.candidate_count)
    
    # ─── DATA LOADING ───
    
    def parse_filename_info(self, filename):
        stage_match = re.search(r'(\d+)(?:st|nd|rd|th)_stage', filename, re.IGNORECASE)
        stage = stage_match.group(1) if stage_match else "1"

        torque_match = re.search(r'(\d+Nm)_(\w+)', filename)
        torque = torque_match.group(1) if torque_match else "Unknown"
        # Normalize condition to title case (Coast, Drive) - fixes "coast" vs "Coast" issue
        condition = torque_match.group(2).title() if torque_match else "Unknown"

        number_match = re.search(r'--(\d+)\.csv$', filename)
        file_number = int(number_match.group(1)) if number_match else 0

        # Detect force vs moment from filename
        # "1st_stage_forces - 25Nm_coast--000.csv" -> force_type = "force"
        # "1st_stage_moments - 25Nm_coast--041.csv" -> force_type = "moment"
        if re.search(r'_moments\s*-', filename, re.IGNORECASE):
            force_type = "moment"
        elif re.search(r'_forces\s*-', filename, re.IGNORECASE):
            force_type = "force"
        else:
            force_type = "unknown"

        return {'stage': stage, 'torque': torque, 'condition': condition,
                'file_number': file_number, 'filename': filename, 'force_type': force_type}
    
    def preprocess_image_for_ocr(self, img):
        """
        Preprocess image for better OCR accuracy (enhanced version):
        1. Scale up 3x (makes Y, Z, 2, 7 more distinct)
        2. Convert to grayscale (removes color confusion)
        3. Increase sharpness (makes letter edges clearer)
        4. Increase contrast (makes text pop)
        """
        from PIL import ImageEnhance, ImageOps, ImageFilter

        original_size = img.size

        # Step 1: Scale up 3x (increased from 2x) - makes letters more distinct
        scale_factor = 2
        new_width = img.width * scale_factor
        new_height = img.height * scale_factor
        img = img.resize((new_width, new_height), Image.LANCZOS)
        debug_print(f"    Scaled: {original_size} -> {img.size} ({scale_factor}x)", "OCR")

        # Step 2: Convert to grayscale (removes color noise from graph elements)
        img = img.convert('L')  # 'L' = grayscale
        debug_print(f"    Converted to grayscale", "OCR")

        # Step 3: Sharpen the image (helps distinguish Z from 2, Y from V)
        enhancer_sharp = ImageEnhance.Sharpness(img)
        img = enhancer_sharp.enhance(2.0)  # 2x sharpness boost
        debug_print(f"    Sharpness enhanced (2x)", "OCR")

        # Step 4: Increase contrast (makes text edges sharper)
        enhancer_contrast = ImageEnhance.Contrast(img)
        img = enhancer_contrast.enhance(2.0)  # 2.5x contrast boost (increased from 2x)
        debug_print(f"    Contrast enhanced (2.5x)", "OCR")

        # Step 5: Optional - apply unsharp mask for additional clarity
        try:
            img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=2))
            debug_print(f"    Unsharp mask applied", "OCR")
        except Exception:
            pass  # UnsharpMask not available in older PIL versions

        return img

    def extract_metadata_from_image_ocr(self, image_path):
        debug_print(f"OCR processing: {image_path.name}", "OCR")

        if not HAS_PIL:
            debug_print(f"  SKIP: PIL not installed", "WARN")
            return None
        if not USE_EASYOCR and not USE_PYTESSERACT:
            debug_print(f"  SKIP: No OCR engine available", "WARN")
            return None

        try:
            img = Image.open(image_path)
            width, height = img.size
            debug_print(f"  Image size: {width}x{height}", "OCR")

            # Step 1: Crop top 35% (increased from 25% to capture more title area)
            crop_percentage = 0.25  # 35% - increased for better OCR detection
            crop_height = int(height * crop_percentage)
            title_area = img.crop((0, 0, width, crop_height))
            debug_print(f"  Crop area: top 25% = {crop_height} pixels", "OCR")

            # Step 2: Apply image preprocessing (scale 2x, grayscale, contrast)
            debug_print(f"  Preprocessing image for OCR...", "OCR")
            title_area = self.preprocess_image_for_ocr(title_area)

            if USE_EASYOCR and ocr_reader:
                img_array = np.array(title_area)
                results = ocr_reader.readtext(img_array, detail=0)
                text = ' '.join(results)
                debug_print(f"  EasyOCR raw: '{text}'", "OCR")
            elif USE_PYTESSERACT:
                text = pytesseract.image_to_string(title_area)
                debug_print(f"  Tesseract raw: '{text}'", "OCR")
            else:
                return None

            result = self.parse_title_text(text)
            if result:
                debug_print(f"  Parsed: B={result.get('bearing')}, Dir={result.get('direction')}, Ord={result.get('order')}", "SUCCESS")
            else:
                debug_print(f"  FAILED: No metadata parsed from text", "WARN")
            return result
        except Exception as e:
            debug_print(f"  ERROR: {e}", "ERROR")
            return None
    
    def parse_title_text(self, text):
        result = {}
        original_text = text

        # Common OCR corrections
        text = text.replace('BI', 'B1').replace('Bl', 'B1')
        text = text.replace('z_', '2.').replace('Z_', '2.')
        text = text.replace('0rder', 'Order').replace('0R', 'Or')

        if text != original_text:
            debug_print(f"    Text corrected: '{original_text}' -> '{text}'", "OCR")

        # Try bearing with description: B1 [Ring Gear - Input Side]
        bearing_match = re.search(r'(B\d+)\s*\[([^\]]+)\]', text)
        if bearing_match:
            result['bearing'] = bearing_match.group(1)
            result['bearing_desc'] = bearing_match.group(2).strip()
            debug_print(f"    BEARING: {result['bearing']} [{result['bearing_desc']}]", "OCR")
        else:
            # Try simpler pattern: just B1, B2, etc
            simple_bearing = re.search(r'\b(B\d+)\b', text)
            if simple_bearing:
                result['bearing'] = simple_bearing.group(1)
                debug_print(f"    BEARING (simple): {result['bearing']}", "OCR")
            else:
                debug_print("    NO BEARING found. Check if text contains B1, B2, etc.", "WARN")

        # Direction detection: X/Y/Z Component
        # Also handle common OCR misreads: V->Y, 2->Z, 7->Z
        direction_found = None
        
        # Pattern 1: Standard "X/Y/Z Component" (most common)
        direction_match = re.search(r'(X|Y|Z)\s*Component', text, re.IGNORECASE)
        if direction_match:
            direction_found = direction_match.group(1).upper()
        
        # Pattern 2: OCR misreads Y as V
        if not direction_found:
            v_match = re.search(r'V\s*Component', text, re.IGNORECASE)
            if v_match:
                direction_found = 'Y'
        
        # Pattern 3: OCR misreads Z as 2 or 7
        if not direction_found:
            z_misread = re.search(r'[27]\s*Component', text, re.IGNORECASE)
            if z_misread:
                direction_found = 'Z'
        
        # Pattern 4: "Force X" or "Force Y" or "Force Z" (alternate format)
        if not direction_found:
            force_match = re.search(r'Force\s+(X|Y|Z)', text, re.IGNORECASE)
            if force_match:
                direction_found = force_match.group(1).upper()
        
        if direction_found:
            result['direction'] = direction_found
            debug_print(f"    DIRECTION: {direction_found}", "OCR")
        else:
            debug_print(f"    NO DIRECTION found in: '{text[:80]}...'", "WARN")

        # Try order: Order 52.0, Order 26.0, Order 78.0, etc
        # OCR may read "Order 52" with space/period between digits, e.g. "Order 5 2" or "Order 5.2"
        # Uses OCR_ORDER_CORRECTIONS config at top of file for common OCR errors

        # Strategy:
        # 1. First try to find "Order XX" where XX is 2+ digits (best case)
        # 2. If not found, look for split digits: "Order 5 2" or "Order 5.2" -> "52"
        # 3. If still not found, accept single digit as-is (Order 1, Order 2, etc. are valid)
        # 4. Apply OCR_ORDER_CORRECTIONS if enabled (to fix "5" -> "52", etc.)

        order_int = None
        order_dec = '0'
        order_pattern_used = None

        # Pattern 1: Full 2+ digit number: "Order 52", "Order52", "Order 52.0"
        order_match = re.search(r'Order\s*(\d{2,})[._]?(\d*)', text, re.IGNORECASE)
        if order_match:
            order_int = order_match.group(1)
            order_dec = order_match.group(2) if order_match.group(2) else '0'
            order_pattern_used = "full 2+ digits"
        else:
            # Pattern 2: Split digits with space/period/underscore: "Order 5 2", "Order 5.2", "Order 5_2"
            # This handles OCR reading "52" as "5 2"
            order_match2 = re.search(r'Order\s*(\d)[._\s]+(\d+)[._]?(\d*)', text, re.IGNORECASE)
            if order_match2:
                # Concatenate: "5" + "2" = "52"
                order_int = order_match2.group(1) + order_match2.group(2)
                order_dec = order_match2.group(3) if order_match2.group(3) else '0'
                order_pattern_used = "split digits joined"
            else:
                # Pattern 3: Single digit - accept as-is (could be Order 1, Order 2, etc.)
                # Also handles multi-digit that didn't match above patterns
                order_match3 = re.search(r'Order\s*(\d+)[._]?(\d*)', text, re.IGNORECASE)
                if order_match3:
                    order_int = order_match3.group(1)
                    order_dec = order_match3.group(2) if order_match3.group(2) else '0'
                    order_pattern_used = "as-is"

        if order_int is not None:
            # Apply OCR_ORDER_CORRECTIONS if enabled
            if OCR_ORDER_CORRECTIONS is not None:
                try:
                    order_val = int(order_int)
                    if order_val in OCR_ORDER_CORRECTIONS:
                        corrected_val = OCR_ORDER_CORRECTIONS[order_val]
                        debug_print(f"    ORDER CORRECTED: {order_val} -> {corrected_val} (using OCR_ORDER_CORRECTIONS config)", "OCR")
                        order_int = str(corrected_val)
                except ValueError:
                    pass  # Not a valid integer, skip correction

            result['order'] = f"{order_int}.{order_dec}"
            debug_print(f"    ORDER ({order_pattern_used}): {result['order']}", "OCR")
        else:
            debug_print("    NO ORDER found in text", "WARN")

        return result if result else None
    
    def load_csv_data(self, csv_path):
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            freq_line = lines[6]
            freq_parts = freq_line.split(',')
            frequencies = np.array([float(x) for x in freq_parts[2:] if x.strip()])

            candidates = []
            i = 7

            while i < len(lines) - 3:
                if not lines[i].strip():
                    i += 1
                    continue

                candidate_data = {'frequencies': frequencies}

                for j, data_type in enumerate(['real', 'imaginary', 'magnitude', 'phase']):
                    if i + j >= len(lines):
                        break
                    line = lines[i + j]
                    parts = line.split(',')

                    if len(parts) > 2:
                        if j == 0:
                            cand_match = re.search(r'Candidate\s*(\d+)', parts[0])
                            if cand_match:
                                candidate_data['candidate'] = int(cand_match.group(1))

                        values = []
                        for x in parts[2:]:
                            x = x.strip()
                            if x:
                                try:
                                    values.append(float(x))
                                except ValueError:
                                    values.append(0.0)
                        candidate_data[data_type] = np.array(values)

                if 'candidate' in candidate_data and 'magnitude' in candidate_data:
                    candidates.append(candidate_data)
                i += 4

            return candidates
        except Exception as e:
            return None
    
    def _process_single_file_ocr(self, csv_file, folder):
        debug_print(f"Processing: {csv_file.name}", "FILE")
        meta = self.parse_filename_info(csv_file.name)
        file_num = meta['file_number']
        # Use filename (without extension) as unique key to avoid collision
        # between force and moment files that have same file_number
        file_key = csv_file.stem  # e.g., "1st_stage_forces - 120Nm_Coast--000"
        force_type = meta.get('force_type', 'unknown')
        debug_print(f"  From filename: stage={meta.get('stage')}, torque={meta.get('torque')}, cond={meta.get('condition')}, num={file_num}, force_type={force_type}", "FILE")

        # Look for corresponding image
        image_path = Path(folder) / (csv_file.stem + "_Candidate000001.png")
        debug_print(f"  Looking for: {image_path.name}", "FILE")

        if not image_path.exists():
            debug_print(f"  IMAGE NOT FOUND!", "WARN")
            # List what files ARE in the folder with similar name
            similar = list(Path(folder).glob(csv_file.stem[:20] + "*"))
            if similar:
                debug_print(f"  Similar files found:", "FILE")
                for s in similar[:5]:
                    debug_print(f"    - {s.name}", "FILE")
            return (file_key, meta, False)  # Use file_key (filename) not file_num

        if not USE_EASYOCR and not USE_PYTESSERACT:
            debug_print(f"  NO OCR ENGINE - cannot read image", "WARN")
            return (file_key, meta, False)  # Use file_key (filename) not file_num

        img_meta = self.extract_metadata_from_image_ocr(image_path)
        if img_meta and 'bearing' in img_meta:
            meta.update(img_meta)

            # Apply force_type from filename to determine final direction
            # If filename says "moments", direction X -> Mx, Y -> My, Z -> Mz
            # If filename says "forces", direction stays X, Y, Z
            ocr_direction = meta.get('direction', '')
            if force_type == 'moment' and ocr_direction in ['X', 'Y', 'Z']:
                meta['direction'] = 'M' + ocr_direction.lower()
                debug_print(f"  Direction adjusted for MOMENT: {ocr_direction} -> {meta['direction']}", "FILE")
            elif force_type == 'force' and ocr_direction in ['Mx', 'My', 'Mz']:
                # OCR wrongly detected moment, but filename says force
                meta['direction'] = ocr_direction[1].upper()
                debug_print(f"  Direction adjusted for FORCE: {ocr_direction} -> {meta['direction']}", "FILE")

            meta['bearing_full'] = f"{img_meta['bearing']} [{img_meta.get('bearing_desc', '')}]" if img_meta.get('bearing_desc') else img_meta['bearing']
            debug_print(f"  SUCCESS: {meta['bearing_full']}, Dir={meta.get('direction')}, Ord={meta.get('order')}", "SUCCESS")
            return (file_key, meta, True)  # Use file_key (filename) not file_num
        else:
            debug_print(f"  FAILED: OCR did not extract bearing info", "WARN")

        return (file_key, meta, False)  # Use file_key (filename) not file_num
    
    def _load_single_csv(self, csv_file):
        meta = self.parse_filename_info(csv_file.name)
        file_key = csv_file.stem  # Use filename stem as unique key
        data = self.load_csv_data(csv_file)
        return (file_key, data, csv_file)
    
    def load_data(self):
        folder = self.folder_var.get()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Error", "Please select a valid folder")
            return

        self.data_folder = folder
        self.file_metadata = {}
        self.csv_data = {}
        
        # Start debug log file
        if DEBUG_MODE:
            log_file = start_debug_log(folder)
            debug_print(f"Data folder: {folder}", "INFO")

        # Warn user if OCR initialization failed (firewall, network, etc.)
        if OCR_INIT_ERROR:
            debug_print(f"OCR unavailable: {OCR_INIT_ERROR}", "WARN")
            messagebox.showwarning(
                "OCR Unavailable",
                f"{OCR_INIT_ERROR}\n\n"
                "The app will still work, but bearing/direction/order\n"
                "must be inferred from filename patterns only.\n\n"
                "For full OCR support, run on a network without firewall\n"
                "restrictions (first run downloads ~100MB of models)."
            )

        self.csv_paths = {}  # Store CSV paths for source validation

        csv_files = list(Path(folder).glob("*.csv"))
        if not csv_files:
            messagebox.showerror("Error", "No CSV files found")
            return

        total_files = len(csv_files)
        self.status_bar.set_status(f"Loading {total_files} files...", Theme.ACCENT_WARNING)
        self.status_bar.show_progress(0)
        self.root.update()

        # ═══════════════════════════════════════════════════════════════════════════
        # CACHING: Try to load OCR metadata from cache first (MUCH faster!)
        # ═══════════════════════════════════════════════════════════════════════════
        cache = load_ocr_cache(folder)
        cache_used = False

        if cache and is_cache_valid(folder, cache):
            # Use cached metadata - skip OCR completely!
            debug_print("=" * 70, "INFO")
            debug_print("USING CACHED METADATA - Skipping OCR (fast load!)", "SUCCESS")
            debug_print("=" * 70, "INFO")
            self.file_metadata = cache['metadata']
            cache_used = True
            ocr_success = total_files  # Assume all successful from cache
            self.status_bar.set_status(f"Loaded {total_files} files from cache", Theme.ACCENT_SECONDARY)
            self.status_bar.update_progress(0.5)
            self.root.update()
        else:
            # No valid cache - run OCR (slow, but saves cache for next time)
            debug_print("=" * 70, "INFO")
            debug_print("NO CACHE - Running OCR detection (first load will be slow...)", "INFO")
            debug_print("=" * 70, "INFO")

            # OCR processing (Phase 1: detect metadata from images)
            ocr_success = 0
            completed = 0

            max_workers = min(30, total_files)
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(self._process_single_file_ocr, csv_file, folder): csv_file
                          for csv_file in csv_files}

                for future in as_completed(futures):
                    completed += 1
                    progress = completed / (total_files * 2)  # OCR is first half
                    self.status_bar.set_status(f"Detecting metadata (first load): {completed}/{total_files}", Theme.ACCENT_WARNING)
                    self.status_bar.update_progress(progress)
                    self.root.update()

                    file_key, meta, success = future.result()  # file_key is now filename stem
                    self.file_metadata[file_key] = meta
                    if success:
                        ocr_success += 1

            # Save cache for next time (async to not block UI)
            if ocr_success > 0:
                self.status_bar.set_status(f"Saving cache...", Theme.ACCENT_WARNING)
                self.root.update()
                try:
                    save_ocr_cache(folder, self.file_metadata)
                    debug_print("Cache saved - next load will be MUCH faster!", "SUCCESS")
                except Exception as e:
                    debug_print(f"Cache save failed (non-fatal): {e}", "WARN")

            # Print OCR summary
            debug_print("=" * 70, "INFO")
            debug_print(f"OCR SUMMARY: {ocr_success}/{total_files} files successfully processed", "INFO")
            if ocr_success < total_files:
                debug_print(f"  WARNING: {total_files - ocr_success} files FAILED OCR detection", "WARN")
            debug_print("=" * 70, "INFO")

        # Update UI to show OCR is done
        self.status_bar.set_status(f"Processing metadata...", Theme.ACCENT_WARNING)
        self.root.update()

        # Print detected metadata summary
        bearings_found = set()
        directions_found = set()
        orders_found = set()
        failed_files = []

        for fn, fm in sorted(self.file_metadata.items()):
            bearing = fm.get('bearing_full', fm.get('bearing', None))
            direction = fm.get('direction', None)
            order = fm.get('order', None)

            if bearing:
                bearings_found.add(bearing)
            if direction:
                directions_found.add(direction)
            if order:
                orders_found.add(order)

            if not bearing or not direction:
                failed_files.append((fn, fm.get('filename', f'file_{fn}')))

        debug_print(f"UNIQUE BEARINGS FOUND: {len(bearings_found)}", "INFO")
        for b in sorted(bearings_found):
            debug_print(f"  - {b}", "INFO")

        debug_print(f"UNIQUE DIRECTIONS FOUND: {len(directions_found)}", "INFO")
        for d in sorted(directions_found):
            debug_print(f"  - {d}", "INFO")

        debug_print(f"UNIQUE ORDERS FOUND: {len(orders_found)}", "INFO")
        for o in sorted(orders_found):
            debug_print(f"  - {o}", "INFO")

        if failed_files:
            debug_print(f"FILES WITH MISSING METADATA ({len(failed_files)}):", "WARN")
            for fn, fname in failed_files[:20]:  # Show first 20
                fm = self.file_metadata.get(fn, {})
                debug_print(f"  File {fn}: {fname}", "WARN")
                debug_print(f"    bearing={fm.get('bearing', 'MISSING')}, dir={fm.get('direction', 'MISSING')}, order={fm.get('order', 'MISSING')}", "WARN")
            if len(failed_files) > 20:
                debug_print(f"  ... and {len(failed_files) - 20} more", "WARN")

        debug_print("=" * 70, "INFO")

        # Update status before CSV loading phase
        self.status_bar.set_status(f"Starting CSV data loading...", Theme.ACCENT_WARNING)
        self.root.update()

        # Always proceed to loading
        self.finish_loading(csv_files)
    
    def show_mapping_dialog(self, csv_files):
        """Show mapping dialog for CSV files"""
        mapper = ctk.CTkToplevel(self.root) if HAS_CTK else tk.Toplevel(self.root)
        mapper.title("Map CSV Files")
        mapper.geometry("950x650")
        mapper.transient(self.root)
        mapper.grab_set()
        
        if HAS_CTK:
            mapper.configure(fg_color=Theme.BG_SECONDARY)

        header = ctk.CTkFrame(mapper, fg_color="transparent") if HAS_CTK else tk.Frame(mapper)
        header.pack(fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(header, text="Map CSV Files to Metadata",
                    font=ctk.CTkFont(size=16, weight="bold"),
                    text_color=Theme.TEXT_PRIMARY).pack(anchor="w") if HAS_CTK else None

        scroll_frame = ctk.CTkScrollableFrame(mapper, fg_color=Theme.BG_CARD, corner_radius=10) if HAS_CTK else tk.Frame(mapper)
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.mapping_vars = {}
        directions = ['X', 'Y', 'Z', 'Mx', 'My', 'Mz']

        for csv_file in sorted(csv_files, key=lambda f: self.parse_filename_info(f.name)['file_number']):
            meta = self.parse_filename_info(csv_file.name)
            file_num = meta['file_number']

            row = ctk.CTkFrame(scroll_frame, fg_color="transparent") if HAS_CTK else tk.Frame(scroll_frame)
            row.pack(fill="x", pady=3)

            ctk.CTkLabel(row, text=f"--{file_num:03d}", width=60, text_color=Theme.TEXT_MUTED).pack(side="left", padx=5) if HAS_CTK else None

            bearing_var = ctk.StringVar(value='B1') if HAS_CTK else tk.StringVar(value='B1')
            ctk.CTkEntry(row, textvariable=bearing_var, width=60, height=30,
                        fg_color=Theme.BG_SECONDARY).pack(side="left", padx=5) if HAS_CTK else None

            desc_var = ctk.StringVar(value='') if HAS_CTK else tk.StringVar(value='')
            ctk.CTkEntry(row, textvariable=desc_var, width=140, height=30,
                        fg_color=Theme.BG_SECONDARY).pack(side="left", padx=5) if HAS_CTK else None

            dir_var = ctk.StringVar(value='X') if HAS_CTK else tk.StringVar(value='X')
            ctk.CTkComboBox(row, variable=dir_var, values=directions, width=70, height=30,
                          fg_color=Theme.BG_SECONDARY).pack(side="left", padx=5) if HAS_CTK else None

            order_var = ctk.StringVar(value='26.0') if HAS_CTK else tk.StringVar(value='26.0')
            ctk.CTkEntry(row, textvariable=order_var, width=60, height=30,
                        fg_color=Theme.BG_SECONDARY).pack(side="left", padx=5) if HAS_CTK else None

            img_path = Path(self.data_folder) / (csv_file.stem + "_Candidate000001.png")
            if img_path.exists():
                ctk.CTkButton(row, text="👁️", width=36, height=30,
                             fg_color=Theme.BG_SECONDARY,
                             command=lambda p=img_path: self._preview_image(p, mapper)).pack(side="left", padx=5) if HAS_CTK else None

            self.mapping_vars[file_num] = {
                'bearing': bearing_var, 'desc': desc_var, 'direction': dir_var, 'order': order_var
            }

        btn_frame = ctk.CTkFrame(mapper, fg_color="transparent") if HAS_CTK else tk.Frame(mapper)
        btn_frame.pack(fill="x", padx=20, pady=15)

        ctk.CTkButton(btn_frame, text="Cancel", command=mapper.destroy,
                     fg_color=Theme.BG_HOVER, text_color=Theme.TEXT_PRIMARY).pack(side="right", padx=5) if HAS_CTK else None
        ctk.CTkButton(btn_frame, text="Apply & Load",
                     command=lambda: self.apply_mapping(mapper, csv_files),
                     fg_color=Theme.ACCENT_PRIMARY).pack(side="right", padx=5) if HAS_CTK else None
    
    def _preview_image(self, path, parent):
        if not HAS_PIL:
            return
        try:
            preview = ctk.CTkToplevel(parent) if HAS_CTK else tk.Toplevel(parent)
            preview.title(path.name)
            preview.geometry("850x500")
            
            img = Image.open(path)
            img.thumbnail((800, 450), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            lbl = ctk.CTkLabel(preview, image=photo, text="") if HAS_CTK else tk.Label(preview, image=photo)
            lbl.image = photo
            lbl.pack(padx=20, pady=20)
        except:
            pass
    
    def apply_mapping(self, window, csv_files):
        for file_num, vars in self.mapping_vars.items():
            bearing = vars['bearing'].get().strip()
            desc = vars['desc'].get().strip()
            direction = vars['direction'].get().strip()
            order = vars['order'].get().strip()

            self.file_metadata[file_num].update({
                'bearing': bearing,
                'bearing_desc': desc,
                'bearing_full': f"{bearing} [{desc}]" if desc else bearing,
                'direction': direction,
                'order': order
            })

        window.destroy()
        self.finish_loading(csv_files)
    
    def finish_loading(self, csv_files):
        all_bearings = set()
        all_directions = set()
        all_orders = set()
        all_stages = set()
        all_torques = set()
        all_conditions = set()

        total_files = len(csv_files)

        # LAZY LOADING: Don't load all CSV data now - load on demand when plotting
        # This makes the app usable much faster, especially on slow network drives
        self.status_bar.set_status(f"Setting up ({total_files} files)...", Theme.ACCENT_WARNING)
        self.root.update()

        # Store csv_files list for lazy loading later
        self._csv_files_list = {f.stem: f for f in csv_files}
        self.csv_paths = {f.stem: f for f in csv_files}

        # Load just ONE CSV to get candidate count (needed for UI)
        if csv_files:
            first_csv = csv_files[0]
            first_data = self.load_csv_data(first_csv)
            if first_data:
                self.csv_data[first_csv.stem] = first_data
                self.candidate_count = len(first_data)

        self.status_bar.update_progress(0.8)
        self.root.update()

        for csv_file in csv_files:
            file_key = csv_file.stem  # Use filename stem as unique key
            fm = self.file_metadata.get(file_key, {})
            all_bearings.add(fm.get('bearing_full', fm.get('bearing', 'Unknown')))
            all_directions.add(fm.get('direction', 'Unknown'))
            all_orders.add(fm.get('order', 'Unknown'))
            all_stages.add(fm.get('stage', '1'))
            all_torques.add(fm.get('torque', 'Unknown'))
            # Normalize condition to title case (Coast, Drive) to avoid duplicates
            cond = fm.get('condition', 'Unknown')
            all_conditions.add(cond.title() if cond and cond != 'Unknown' else cond)

        self.bearings = sorted([b for b in all_bearings if b != 'Unknown'],
                              key=lambda x: int(re.search(r'B(\d+)', x).group(1)) if re.search(r'B(\d+)', x) else 0)
        self.directions = sorted([d for d in all_directions if d != 'Unknown'])
        self.orders = sorted([o for o in all_orders if o != 'Unknown'],
                            key=lambda x: float(x) if x.replace('.','').isdigit() else 0)
        self.stages = sorted(all_stages)
        self.torques = sorted(all_torques)
        self.conditions = sorted(all_conditions)

        # Update torque checkboxes
        for widget in self.torque_frame.winfo_children():
            widget.destroy()

        self.torque_vars = {}
        for torque in self.torques:
            var = ctk.BooleanVar(value=False) if HAS_CTK else tk.BooleanVar(value=False)
            self.torque_vars[torque] = var

            cb = ctk.CTkCheckBox(
                self.torque_frame, text=torque, variable=var,
                font=ctk.CTkFont(size=12),
                fg_color=Theme.ACCENT_PRIMARY,
                text_color=Theme.TEXT_PRIMARY
            ) if HAS_CTK else tk.Checkbutton(self.torque_frame, text=torque, variable=var)
            cb.pack(anchor="w", pady=2)
            self.torque_checks[torque] = cb

        if self.torques:
            self.torque_vars[self.torques[0]].set(True)
            self.torque_var.set(self.torques[0])  # For backward compatibility

        # Update bearing checkboxes
        for widget in self.bearing_frame.winfo_children():
            widget.destroy()

        self.bearing_vars = {}
        for bearing in self.bearings:
            var = ctk.BooleanVar(value=False) if HAS_CTK else tk.BooleanVar(value=False)
            self.bearing_vars[bearing] = var
            short_name = re.search(r'(B\d+)', bearing)
            display_name = short_name.group(1) if short_name else bearing

            cb = ctk.CTkCheckBox(
                self.bearing_frame, text=display_name, variable=var,
                font=ctk.CTkFont(size=12),
                fg_color=Theme.ACCENT_PRIMARY,
                text_color=Theme.TEXT_PRIMARY
            ) if HAS_CTK else tk.Checkbutton(self.bearing_frame, text=display_name, variable=var)
            cb.pack(anchor="w", pady=2)

        if self.bearings:
            self.bearing_vars[self.bearings[0]].set(True)

        # Update combos - add "All" option for order to allow multi-bearing/direction plots
        if HAS_CTK:
            self.order_combo.configure(values=["All"] + self.orders)
            self.stage_combo.configure(values=self.stages)
            self.condition_combo.configure(values=self.conditions)

        if self.orders:
            self.order_var.set(self.orders[0])
        if self.stages:
            self.stage_var.set(self.stages[0])
        if self.conditions:
            self.condition_var.set(self.conditions[0])

        for d in self.directions:
            if d in self.direction_checks and HAS_CTK:
                self.direction_checks[d].configure(text_color=Theme.TEXT_PRIMARY)
        if self.directions:
            self.direction_vars[self.directions[0]].set(True)

        self.cand_count_label.configure(text=f"({self.candidate_count} candidates)")

        # Update source validator
        self.source_validator = SourceValidator(self.data_folder, self.file_metadata, self.csv_data)
        self.graph_tracker.source_validator = self.source_validator

        # Hide progress bar and show success
        self.status_bar.hide_progress()
        self.status_bar.set_status(f"✓ Loaded {len(self.csv_data)} files • {self.candidate_count} candidates", Theme.ACCENT_SECONDARY)
    
    def get_filtered_data(self):
        """Get filtered data with source info for validation"""
        order = self.order_var.get()
        stage = self.stage_var.get()
        torque = self.torque_var.get()
        condition = self.condition_var.get()

        selected_bearings = []
        for bearing_full, var in self.bearing_vars.items():
            if var.get():
                bearing_match = re.search(r'(B\d+)', bearing_full)
                if bearing_match:
                    selected_bearings.append((bearing_match.group(1), bearing_full))
        if not selected_bearings:
            return {}

        selected_dirs = [d for d, v in self.direction_vars.items() if v.get()]
        if not selected_dirs:
            return {}

        candidates = self.parse_candidate_selection()
        if not candidates:
            return {}

        result = {}
        for file_num, meta in self.file_metadata.items():
            bearing = meta.get('bearing')
            bearing_full = meta.get('bearing_full', bearing)

            is_selected = any(b == bearing for b, _ in selected_bearings)

            # Check order - "All" means match any order
            order_match = (order == "All") or (meta.get('order') == order)

            if (is_selected and
                order_match and
                meta.get('stage') == stage and
                meta.get('torque') == torque and
                meta.get('condition', '').lower() == condition.lower() and
                meta.get('direction') in selected_dirs):

                # LAZY LOADING: Load CSV data on demand if not already loaded
                if file_num not in self.csv_data:
                    csv_path = self._csv_files_list.get(file_num) if hasattr(self, '_csv_files_list') else None
                    if csv_path:
                        data = self.load_csv_data(csv_path)
                        if data:
                            self.csv_data[file_num] = data

                if file_num in self.csv_data:
                    if bearing_full not in result:
                        result[bearing_full] = {}
                    direction = meta['direction']
                    if direction not in result[bearing_full]:
                        result[bearing_full][direction] = []
                    
                    # Get CSV path for source validation
                    csv_path = self.csv_paths.get(file_num) if hasattr(self, 'csv_paths') else None
                    
                    for cand_data in self.csv_data[file_num]:
                        if cand_data.get('candidate') in candidates:
                            # Add source info for validation
                            cand_num = cand_data.get('candidate', 1)
                            cand_data['_source_info'] = {
                                'file_number': file_num,
                                'csv_path': csv_path,
                                'image_path': Path(self.data_folder) / f"{csv_path.stem}_Candidate{cand_num:06d}.png" if csv_path else None,
                                'candidate': cand_num,
                                'bearing': bearing,
                                'bearing_full': bearing_full,
                                'direction': direction,
                                'order': order,
                                'torque': torque,
                                'condition': condition
                            }
                            result[bearing_full][direction].append(cand_data)
        return result

    # ═══════════════════════════════════════════════════════════════════════════════
    # SCALAR MODE - RMS and Peak calculation for frequency bands
    # ═══════════════════════════════════════════════════════════════════════════════

    # Frequency bands for scalar calculation (Hz)
    SCALAR_BANDS = [
        (0, 1000, "0-1kHz"),
        (1000, 3000, "1-3kHz"),
        (3000, 5000, "3-5kHz"),
        (5000, 10000, "5-10kHz")
    ]

    def on_output_mode_change(self):
        """Handle output mode toggle between Dynamic and Scalar"""
        mode = self.output_mode.get()
        debug_print(f"Output mode changed to: {mode}", "INFO")
        # Auto-replot if data is loaded
        if hasattr(self, 'data') and self.data:
            self.plot_data()

    def calculate_scalar_values(self, frequencies, magnitude):
        """Calculate RMS and Peak values for each frequency band.

        Args:
            frequencies: Array of frequency values (Hz)
            magnitude: Array of magnitude values

        Returns:
            dict with band_name -> {'rms': value, 'peak': value}
        """
        freq = np.array(frequencies)
        mag = np.array(magnitude)

        results = {}
        for low, high, label in self.SCALAR_BANDS:
            # Find indices within this band
            mask = (freq >= low) & (freq < high)
            band_mag = mag[mask]

            if len(band_mag) > 0:
                rms = np.sqrt(np.mean(band_mag ** 2))
                peak = np.max(band_mag)
            else:
                rms = 0.0
                peak = 0.0

            results[label] = {'rms': rms, 'peak': peak}

        return results

    def plot_scalar_data(self):
        """Plot bar charts for Scalar mode (RMS and Peak values per frequency band)"""
        self.fig.clear()
        if self.graph_tracker:
            self.graph_tracker.clear_data()

        filtered = self.get_filtered_data()
        if not filtered:
            messagebox.showwarning("No Data", "No data matches filters.\nCheck selections.")
            return

        order = self.order_var.get()
        torque = self.torque_var.get()
        condition = self.condition_var.get()

        bearings = sorted(filtered.keys(),
                         key=lambda x: int(re.search(r'B(\d+)', x).group(1)) if re.search(r'B(\d+)', x) else 0)
        all_dirs = set()
        for bearing_data in filtered.values():
            all_dirs.update(bearing_data.keys())
        directions = sorted(all_dirs)

        num_bearings = len(bearings)
        num_dirs = len(directions)

        if num_bearings == 0 or num_dirs == 0:
            messagebox.showwarning("No Data", "No data to plot")
            return

        # 2 rows per bearing (RMS and Peak)
        num_rows = num_bearings * 2
        axes = self.fig.subplots(num_rows, num_dirs, squeeze=False)

        num_cands = len(self.parse_candidate_selection())
        colors = Theme.PLOT_COLORS
        band_labels = [b[2] for b in self.SCALAR_BANDS]
        x_pos = np.arange(len(band_labels))
        bar_width = 0.8 / max(num_cands, 1)

        # Store bar info for right-click validation
        self._scalar_bar_info = {}

        for bearing_idx, bearing_full in enumerate(bearings):
            short_match = re.search(r'(B\d+)', bearing_full)
            bearing_short = short_match.group(1) if short_match else bearing_full

            bearing_data = filtered[bearing_full]

            for dir_idx, direction in enumerate(directions):
                cands = bearing_data.get(direction, [])

                ax_rms = axes[bearing_idx * 2, dir_idx]
                ax_peak = axes[bearing_idx * 2 + 1, dir_idx]

                for i, cd in enumerate(cands):
                    color = colors[i % len(colors)]
                    cand_num = cd.get('candidate', i+1)
                    label = f"C{cand_num}"
                    source_info = cd.get('_source_info', {})

                    # Calculate scalar values for this candidate
                    if 'frequencies' in cd and 'magnitude' in cd:
                        scalar_vals = self.calculate_scalar_values(cd['frequencies'], cd['magnitude'])

                        rms_values = [scalar_vals[bl]['rms'] for bl in band_labels]
                        peak_values = [scalar_vals[bl]['peak'] for bl in band_labels]

                        # Bar positions offset for each candidate
                        offset = (i - num_cands/2 + 0.5) * bar_width

                        bars_rms = ax_rms.bar(x_pos + offset, rms_values, bar_width,
                                  color=color, label=label, alpha=0.8, edgecolor='white', picker=True)
                        bars_peak = ax_peak.bar(x_pos + offset, peak_values, bar_width,
                                   color=color, label=label, alpha=0.8, edgecolor='white', picker=True)

                        # Store info for each bar for right-click
                        for band_idx, (bar_rms, bar_peak) in enumerate(zip(bars_rms, bars_peak)):
                            band_range = self.SCALAR_BANDS[band_idx]
                            bar_info = {
                                'candidate': cand_num,
                                'bearing': bearing_short,
                                'bearing_full': bearing_full,
                                'direction': direction,
                                'band_label': band_labels[band_idx],
                                'band_low': band_range[0],
                                'band_high': band_range[1],
                                'source_info': source_info
                            }
                            self._scalar_bar_info[id(bar_rms)] = bar_info
                            self._scalar_bar_info[id(bar_peak)] = bar_info

                # Style RMS axis
                ax_rms.set_title(f"{bearing_short} - {direction} - RMS",
                                fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
                ax_rms.set_xlabel("Frequency Band", fontsize=10, color=Theme.TEXT_SECONDARY)
                ax_rms.set_ylabel("RMS (N)", fontsize=10, color=Theme.TEXT_SECONDARY)
                ax_rms.set_xticks(x_pos)
                ax_rms.set_xticklabels(band_labels, fontsize=9)
                ax_rms.grid(True, axis='y', ls='-', alpha=0.2, color=Theme.BORDER_DEFAULT)
                ax_rms.set_facecolor(Theme.BG_CARD)
                for spine in ax_rms.spines.values():
                    spine.set_color(Theme.BORDER_DEFAULT)
                    spine.set_linewidth(0.5)

                # Style Peak axis
                ax_peak.set_title(f"{bearing_short} - {direction} - Peak",
                                 fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
                ax_peak.set_xlabel("Frequency Band", fontsize=10, color=Theme.TEXT_SECONDARY)
                ax_peak.set_ylabel("Peak (N)", fontsize=10, color=Theme.TEXT_SECONDARY)
                ax_peak.set_xticks(x_pos)
                ax_peak.set_xticklabels(band_labels, fontsize=9)
                ax_peak.grid(True, axis='y', ls='-', alpha=0.2, color=Theme.BORDER_DEFAULT)
                ax_peak.set_facecolor(Theme.BG_CARD)
                for spine in ax_peak.spines.values():
                    spine.set_color(Theme.BORDER_DEFAULT)
                    spine.set_linewidth(0.5)

        # Add legend to first subplot
        if num_cands <= 15 and axes.size > 0:
            axes[0, 0].legend(loc='upper right', fontsize=8, ncol=2,
                             facecolor=Theme.BG_CARD, edgecolor=Theme.BORDER_DEFAULT,
                             labelcolor=Theme.TEXT_PRIMARY, framealpha=0.95)

        title = f"Bearing Force (SCALAR) - {torque} {condition}"
        if num_cands > 1:
            title += f" ({num_cands} candidates)"
        self.fig.suptitle(title, fontsize=14, fontweight='bold', color=Theme.TEXT_PRIMARY)

        self.fig.tight_layout()
        self.fig.subplots_adjust(top=0.93)

        # Connect right-click event for bar validation (use button_press_event, not pick_event)
        self._scalar_cid = self.canvas.mpl_connect('button_press_event', self._on_scalar_click)

        self.canvas.draw()
        self.status_bar.set_status(f"✓ Scalar plot: RMS & Peak • Right-click bar to validate", Theme.ACCENT_SECONDARY)

    def _on_scalar_click(self, event):
        """Handle right-click on scalar bar chart to show source validation"""
        if event.button != 3:  # Right-click only
            return

        if event.inaxes is None:
            return

        # Find which bar was clicked by checking if click point is inside any bar
        clicked_bar = None
        clicked_bar_info = None

        for bar_id, bar_info in self._scalar_bar_info.items():
            # We need to find the actual bar object - iterate through all axes patches
            for ax in self.fig.axes:
                for patch in ax.patches:
                    if id(patch) == bar_id:
                        # Check if click is inside this bar
                        if patch.contains_point([event.x, event.y]):
                            clicked_bar = patch
                            clicked_bar_info = bar_info
                            break
                if clicked_bar:
                    break
            if clicked_bar:
                break

        if not clicked_bar_info:
            return

        source_info = clicked_bar_info.get('source_info', {})
        if not source_info:
            messagebox.showinfo("No Source", "No source info available for this bar")
            return

        # Add band info to source_info for highlighting
        source_info_with_band = source_info.copy()
        source_info_with_band['freq_band_low'] = clicked_bar_info['band_low']
        source_info_with_band['freq_band_high'] = clicked_bar_info['band_high']
        source_info_with_band['band_label'] = clicked_bar_info['band_label']

        # Show context menu
        self._show_scalar_context_menu(event, clicked_bar_info, source_info_with_band)

    def _show_scalar_context_menu(self, event, bar_info, source_info):
        """Show right-click context menu for scalar bar validation"""
        menu = Menu(self.canvas.get_tk_widget(), tearoff=0)

        candidate = bar_info.get('candidate', '?')
        bearing = bar_info.get('bearing', '?')
        direction = bar_info.get('direction', '?')
        band_label = bar_info.get('band_label', '?')

        # Header
        menu.add_command(
            label=f"▶ Candidate {candidate} ({bearing}-{direction})",
            state="disabled"
        )
        menu.add_command(
            label=f"   Band: {band_label}",
            state="disabled"
        )
        menu.add_separator()

        # Menu items
        menu.add_command(
            label="📊 Open CSV (highlight freq band rows)",
            command=lambda si=source_info: self.source_validator.open_csv_with_band(si)
        )
        menu.add_command(
            label=f"🖼️ Open Image (Candidate {candidate})",
            command=lambda si=source_info: self.source_validator.open_image_only(si)
        )
        menu.add_separator()
        menu.add_command(
            label="ℹ️ Show Source Details",
            command=lambda si=source_info: self.source_validator.show_source_info_dialog(
                self.canvas.get_tk_widget().winfo_toplevel(), si)
        )

        try:
            # Get screen coordinates from canvas widget
            widget = self.canvas.get_tk_widget()
            x_screen = widget.winfo_rootx() + int(event.x)
            y_screen = widget.winfo_rooty() + int(widget.winfo_height() - event.y)  # Flip y
            menu.tk_popup(x_screen, y_screen)
        finally:
            menu.grab_release()

    def clear_plot(self):
        self.fig.clear()
        if self.graph_tracker:
            self.graph_tracker.clear_data()
        self._show_welcome_screen()
        self.status_bar.set_status("Plot cleared", Theme.TEXT_MUTED)

    def export_debug_info(self):
        """Export detailed debug info to a text file for troubleshooting"""
        from datetime import datetime

        # Ask where to save
        filepath = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile=f"debug_info_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            title="Save Debug Info"
        )
        if not filepath:
            return

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("BEARING FORCE VIEWER - DEBUG REPORT\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")

                # Environment info
                f.write("─" * 40 + "\n")
                f.write("ENVIRONMENT\n")
                f.write("─" * 40 + "\n")
                f.write(f"OCR Engine Available: EasyOCR={USE_EASYOCR}, PyTesseract={USE_PYTESSERACT}\n")
                f.write(f"OCR Reader Initialized: {ocr_reader is not None}\n")
                f.write(f"OCR Init Error: {OCR_INIT_ERROR}\n")
                f.write(f"PIL Available: {HAS_PIL}\n")
                f.write(f"CustomTkinter: {HAS_CTK}\n\n")

                # Data folder
                f.write("─" * 40 + "\n")
                f.write("DATA FOLDER\n")
                f.write("─" * 40 + "\n")
                folder = getattr(self, 'data_folder', None)
                f.write(f"Folder Path: {folder}\n\n")

                if folder and os.path.exists(folder):
                    # List ALL files in folder
                    f.write("ALL FILES IN FOLDER:\n")
                    all_files = sorted(os.listdir(folder))
                    for fname in all_files:
                        fpath = os.path.join(folder, fname)
                        fsize = os.path.getsize(fpath) if os.path.isfile(fpath) else 0
                        f.write(f"  {fname} ({fsize:,} bytes)\n")
                    f.write("\n")

                    # CSV files specifically
                    csv_files = [fn for fn in all_files if fn.endswith('.csv')]
                    f.write(f"CSV FILES FOUND: {len(csv_files)}\n")

                    # Count forces vs moments
                    forces_count = sum(1 for fn in csv_files if re.search(r'_forces\s*-', fn, re.IGNORECASE))
                    moments_count = sum(1 for fn in csv_files if re.search(r'_moments\s*-', fn, re.IGNORECASE))
                    f.write(f"  -> FORCES files (regex '_forces\\s*-'): {forces_count}\n")
                    f.write(f"  -> MOMENTS files (regex '_moments\\s*-'): {moments_count}\n")
                    f.write(f"  -> UNKNOWN: {len(csv_files) - forces_count - moments_count}\n\n")

                    for csv_name in csv_files:
                        # Parse filename info
                        meta = self.parse_filename_info(csv_name)
                        f.write(f"\n  FILE: {csv_name}\n")
                        f.write(f"    file_number: {meta.get('file_number')}\n")

                        # Show EXACT regex match for debugging
                        moment_match = re.search(r'_moments\s*-', csv_name, re.IGNORECASE)
                        force_match = re.search(r'_forces\s*-', csv_name, re.IGNORECASE)
                        f.write(f"    REGEX TEST: _moments match={moment_match is not None}, _forces match={force_match is not None}\n")
                        f.write(f"    force_type:  {meta.get('force_type')} <-- FROM FILENAME\n")
                        f.write(f"    stage:       {meta.get('stage')}\n")
                        f.write(f"    torque:      {meta.get('torque')}\n")
                        f.write(f"    condition:   {meta.get('condition')}\n")
                    f.write("\n")

                # Loaded metadata
                f.write("─" * 40 + "\n")
                f.write("LOADED FILE METADATA (after OCR)\n")
                f.write("─" * 40 + "\n")
                if hasattr(self, 'file_metadata') and self.file_metadata:
                    for file_key in sorted(self.file_metadata.keys()):
                        fm = self.file_metadata[file_key]
                        f.write(f"\nFile: {file_key}\n")
                        f.write(f"  filename:     {fm.get('filename', 'N/A')}\n")
                        f.write(f"  force_type:   {fm.get('force_type', 'N/A')} <-- FROM FILENAME\n")
                        f.write(f"  bearing:      {fm.get('bearing', 'N/A')}\n")
                        f.write(f"  bearing_full: {fm.get('bearing_full', 'N/A')}\n")
                        f.write(f"  direction:    {fm.get('direction', 'N/A')} <-- FINAL (after force_type applied)\n")
                        f.write(f"  order:        {fm.get('order', 'N/A')}\n")
                        f.write(f"  stage:        {fm.get('stage', 'N/A')}\n")
                        f.write(f"  torque:       {fm.get('torque', 'N/A')}\n")
                        f.write(f"  condition:    {fm.get('condition', 'N/A')}\n")
                else:
                    f.write("No file metadata loaded yet. Click 'Load Data' first.\n")

                # Detected directions
                f.write("\n")
                f.write("─" * 40 + "\n")
                f.write("DETECTED UNIQUE VALUES\n")
                f.write("─" * 40 + "\n")
                f.write(f"Bearings:   {getattr(self, 'bearings', [])}\n")
                f.write(f"Directions: {getattr(self, 'directions', [])}\n")
                f.write(f"Orders:     {getattr(self, 'orders', [])}\n")
                f.write(f"Stages:     {getattr(self, 'stages', [])}\n")
                f.write(f"Torques:    {getattr(self, 'torques', [])}\n")
                f.write(f"Conditions: {getattr(self, 'conditions', [])}\n")

                # Current UI selections
                f.write("\n")
                f.write("─" * 40 + "\n")
                f.write("CURRENT UI SELECTIONS\n")
                f.write("─" * 40 + "\n")
                if hasattr(self, 'bearing_vars'):
                    selected_bearings = [b for b, v in self.bearing_vars.items() if v.get()]
                    f.write(f"Selected Bearings: {selected_bearings}\n")
                if hasattr(self, 'direction_vars'):
                    selected_dirs = [d for d, v in self.direction_vars.items() if v.get()]
                    f.write(f"Selected Directions: {selected_dirs}\n")
                f.write(f"Order: {getattr(self, 'order_var', tk.StringVar()).get()}\n")
                f.write(f"Torque: {getattr(self, 'torque_var', tk.StringVar()).get()}\n")
                f.write(f"Condition: {getattr(self, 'condition_var', tk.StringVar()).get()}\n")

                f.write("\n" + "=" * 80 + "\n")
                f.write("END OF DEBUG REPORT\n")
                f.write("=" * 80 + "\n")

            messagebox.showinfo("Debug Info Exported", f"Debug info saved to:\n{filepath}\n\nPlease share this file for troubleshooting.")

            # Also try to open the file
            try:
                os.startfile(filepath)
            except:
                pass

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export debug info: {e}")

    def plot_data(self):
        """Plot data with 3-level tabbed interface: Torque → Bearing → Order"""
        # Check output mode - Scalar vs Dynamic
        if self.output_mode.get() == "scalar":
            self.plot_scalar_data()
            return

        # Get selected torques
        selected_torques = [t for t, v in self.torque_vars.items() if v.get()]
        if not selected_torques:
            messagebox.showwarning("No Torque Selected", "Please select at least one torque.")
            return

        # Get selected bearings
        selected_bearings = [b for b, v in self.bearing_vars.items() if v.get()]
        if not selected_bearings:
            messagebox.showwarning("No Bearing Selected", "Please select at least one bearing.")
            return

        # Clear old plot container contents
        for widget in self.plot_container.winfo_children():
            widget.destroy()

        # Clear graph tracker
        if self.graph_tracker:
            self.graph_tracker.clear_data()

        # Store references to all figures/canvases for graph tracking
        self.tab_figures = {}
        self.tab_canvases = {}

        plot_type = self.plot_type.get()
        y_scale = self.y_scale.get()
        selected_order = self.order_var.get()  # Could be "All" or specific order
        condition = self.condition_var.get()
        num_cands = len(self.parse_candidate_selection())
        colors = Theme.PLOT_COLORS

        # Determine which orders to show
        if selected_order == "All":
            # Get all available orders from metadata
            available_orders = set()
            for meta in self.file_metadata.values():
                if meta.get('order'):
                    available_orders.add(meta['order'])
            orders_to_plot = sorted(available_orders, key=lambda x: float(x) if x.replace('.', '').isdigit() else 0)
        else:
            orders_to_plot = [selected_order]

        # Create outer notebook for Torque tabs
        torque_notebook = ttk.Notebook(self.plot_container)
        torque_notebook.pack(fill="both", expand=True)

        total_plots = 0

        for torque in selected_torques:
            # Set torque_var for data filtering
            self.torque_var.set(torque)

            # Create frame for this torque tab
            torque_frame = ttk.Frame(torque_notebook)
            torque_notebook.add(torque_frame, text=f" {torque} ")

            # Create inner notebook for Bearing tabs
            bearing_notebook = ttk.Notebook(torque_frame)
            bearing_notebook.pack(fill="both", expand=True)

            for bearing_full in selected_bearings:
                # Get short bearing name
                short_match = re.search(r'(B\d+)', bearing_full)
                bearing_short = short_match.group(1) if short_match else bearing_full

                # Create frame for this bearing tab
                bearing_frame = ttk.Frame(bearing_notebook)
                bearing_notebook.add(bearing_frame, text=f" {bearing_short} ")

                # Create innermost notebook for Order tabs
                order_notebook = ttk.Notebook(bearing_frame)
                order_notebook.pack(fill="both", expand=True)

                for order in orders_to_plot:
                    # Get data for this torque/bearing/order combination
                    filtered = self.get_filtered_data_for_bearing(torque, bearing_full, order, condition)
                    if not filtered:
                        continue

                    # Create frame for this order tab
                    order_frame = ttk.Frame(order_notebook)
                    order_notebook.add(order_frame, text=f" Order {order} ")

                    # Create figure for this order
                    fig = Figure(figsize=(10, 6), dpi=100, facecolor=Theme.BG_SECONDARY)
                    canvas = FigureCanvasTkAgg(fig, master=order_frame)
                    canvas_widget = canvas.get_tk_widget()
                    canvas_widget.pack(fill="both", expand=True)

                    # Store references
                    key = f"{torque}_{bearing_short}_{order}"
                    self.tab_figures[key] = fig
                    self.tab_canvases[key] = canvas

                    # Get directions from filtered data
                    all_dirs = set()
                    for bearing_data in filtered.values():
                        all_dirs.update(bearing_data.keys())
                    directions = sorted(all_dirs)
                    num_dirs = max(1, len(directions))

                    # Calculate grid layout - max 3 columns for better viewing
                    max_cols = 3
                    num_cols = min(max_cols, num_dirs)
                    num_dir_rows = (num_dirs + num_cols - 1) // num_cols  # ceiling division

                    if plot_type == "both":
                        # For mag+phase, each direction needs 2 rows
                        num_rows = num_dir_rows * 2
                    else:
                        num_rows = num_dir_rows

                    axes = fig.subplots(num_rows, num_cols, squeeze=False)
                    all_axes = []

                    bearing_data = filtered.get(bearing_full, {})

                    for dir_idx, direction in enumerate(directions):
                        cands = bearing_data.get(direction, [])

                        # Convert linear index to grid row/col
                        grid_row = dir_idx // num_cols
                        grid_col = dir_idx % num_cols

                        if plot_type == "both":
                            # Each direction gets 2 rows (mag + phase)
                            ax_mag = axes[grid_row * 2, grid_col]
                            ax_phase = axes[grid_row * 2 + 1, grid_col]
                            all_axes.extend([ax_mag, ax_phase])
                        elif plot_type == "magnitude":
                            ax_mag = axes[grid_row, grid_col]
                            ax_phase = None
                            all_axes.append(ax_mag)
                        else:
                            ax_mag = None
                            ax_phase = axes[grid_row, grid_col]
                            all_axes.append(ax_phase)

                        for i, cd in enumerate(cands):
                            color = colors[i % len(colors)]
                            freq = cd['frequencies']
                            label = f"C{cd.get('candidate', i+1)}"
                            source_info = cd.get('_source_info', {})

                            if ax_mag is not None and 'magnitude' in cd:
                                mag = np.maximum(np.array(cd['magnitude']), 1e-10)
                                if y_scale == 'log':
                                    line, = ax_mag.semilogy(freq, mag, color=color, label=label,
                                                           alpha=0.85, linewidth=1.2, picker=5)
                                else:
                                    line, = ax_mag.plot(freq, mag, color=color, label=label,
                                                       alpha=0.85, linewidth=1.2, picker=5)
                                mag_source = source_info.copy()
                                mag_source['data_type'] = 'magnitude'
                                if self.graph_tracker:
                                    self.graph_tracker.register_line(ax_mag, line, freq, mag, label, color, mag_source)

                            if ax_phase is not None and 'phase' in cd:
                                line, = ax_phase.plot(freq, cd['phase'], color=color, label=label,
                                                     alpha=0.85, linewidth=1.2, picker=5)
                                phase_source = source_info.copy()
                                phase_source['data_type'] = 'phase'
                                if self.graph_tracker:
                                    self.graph_tracker.register_line(ax_phase, line, freq, cd['phase'], label, color, phase_source)

                        # Style axes
                        if ax_mag:
                            ax_mag.set_title(f"{direction}", fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
                            ax_mag.set_xlabel("Frequency (Hz)", fontsize=10, color=Theme.TEXT_SECONDARY)
                            ax_mag.set_ylabel("Magnitude (N)", fontsize=10, color=Theme.TEXT_SECONDARY)
                            ax_mag.grid(True, which='both', ls='-', alpha=0.2, color=Theme.BORDER_DEFAULT)
                            ax_mag.set_xlim(left=0)
                            ax_mag.tick_params(labelsize=9, colors=Theme.TEXT_SECONDARY)
                            ax_mag.set_facecolor(Theme.BG_CARD)
                            for spine in ax_mag.spines.values():
                                spine.set_color(Theme.BORDER_DEFAULT)
                                spine.set_linewidth(0.5)

                        if ax_phase:
                            if plot_type != "both":
                                ax_phase.set_title(f"{direction}", fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
                            ax_phase.set_xlabel("Frequency (Hz)", fontsize=10, color=Theme.TEXT_SECONDARY)
                            ax_phase.set_ylabel("Phase (rad)", fontsize=10, color=Theme.TEXT_SECONDARY)
                            ax_phase.grid(True, ls='-', alpha=0.2, color=Theme.BORDER_DEFAULT)
                            ax_phase.set_xlim(left=0)
                            ax_phase.tick_params(labelsize=9, colors=Theme.TEXT_SECONDARY)
                            ax_phase.set_facecolor(Theme.BG_CARD)
                            for spine in ax_phase.spines.values():
                                spine.set_color(Theme.BORDER_DEFAULT)
                                spine.set_linewidth(0.5)

                    if num_cands <= 15 and len(all_axes) > 0:
                        all_axes[0].legend(loc='upper right', fontsize=8, ncol=2,
                                          facecolor=Theme.BG_CARD, edgecolor=Theme.BORDER_DEFAULT,
                                          labelcolor=Theme.TEXT_PRIMARY, framealpha=0.95)

                    # Hide unused axes (when num_dirs < num_rows * num_cols)
                    total_cells = num_dir_rows * num_cols
                    for empty_idx in range(num_dirs, total_cells):
                        empty_row = empty_idx // num_cols
                        empty_col = empty_idx % num_cols
                        if plot_type == "both":
                            axes[empty_row * 2, empty_col].set_visible(False)
                            axes[empty_row * 2 + 1, empty_col].set_visible(False)
                        else:
                            axes[empty_row, empty_col].set_visible(False)

                    fig.suptitle(f"{bearing_short} - {torque} {condition} - Order {order}", fontsize=12, fontweight='bold', color=Theme.TEXT_PRIMARY)
                    fig.tight_layout()
                    fig.subplots_adjust(top=0.90)
                    canvas.draw()

                    # Add double-click handler for maximize/restore
                    canvas_widget.bind("<Double-Button-1>", lambda e, k=key: self._toggle_maximize_plot(k))

                    total_plots += 1

        if total_plots == 0:
            messagebox.showwarning("No Data", "No data matches the selected filters.")
            return

        self.status_bar.set_status(f"✓ Plotted {len(selected_torques)} torque(s), {len(selected_bearings)} bearing(s), {len(orders_to_plot)} order(s) • Double-click to maximize • Right-click to validate", Theme.ACCENT_SECONDARY)

    def _toggle_maximize_plot(self, key):
        """Toggle maximize/restore for a plot (opens in new window)."""
        if key not in self.tab_figures:
            return

        fig = self.tab_figures[key]

        # Create a new maximized window with the plot
        popup = tk.Toplevel(self)
        popup.title(f"Maximized Plot - {key.replace('_', ' ')}")
        popup.geometry("1200x800")
        popup.transient(self)

        # Create a new figure in the popup with the same content
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

        # Clone the figure by recreating it
        popup_fig = Figure(figsize=(14, 10), dpi=100, facecolor=Theme.BG_SECONDARY)

        # Copy axes from original figure
        orig_axes = fig.get_axes()
        if orig_axes:
            # Determine grid dimensions from original
            positions = [ax.get_position() for ax in orig_axes]
            num_orig_axes = len(orig_axes)

            # Recreate subplot structure
            # Find unique rows and cols from positions
            rows = sorted(set(round(1 - p.y0 - p.height, 2) for p in positions))
            cols = sorted(set(round(p.x0, 2) for p in positions))
            num_rows = len(rows)
            num_cols = len(cols) if cols else 1

            if num_rows == 0:
                num_rows = 1
            if num_cols == 0:
                num_cols = 1

            # Create subplots
            popup_axes = popup_fig.subplots(num_rows, num_cols, squeeze=False)

            # Copy each axis content
            for idx, ax in enumerate(orig_axes):
                row = idx // num_cols
                col = idx % num_cols
                if row < num_rows and col < num_cols:
                    popup_ax = popup_axes[row, col]

                    # Copy lines
                    for line in ax.get_lines():
                        popup_ax.plot(line.get_xdata(), line.get_ydata(),
                                     color=line.get_color(), label=line.get_label(),
                                     alpha=line.get_alpha() or 1.0, linewidth=line.get_linewidth())

                    # Copy axis properties
                    popup_ax.set_title(ax.get_title(), fontsize=12, fontweight='bold', color=Theme.TEXT_PRIMARY)
                    popup_ax.set_xlabel(ax.get_xlabel(), fontsize=11, color=Theme.TEXT_SECONDARY)
                    popup_ax.set_ylabel(ax.get_ylabel(), fontsize=11, color=Theme.TEXT_SECONDARY)
                    popup_ax.set_xlim(ax.get_xlim())
                    popup_ax.set_ylim(ax.get_ylim())
                    popup_ax.grid(True, which='both', ls='-', alpha=0.2, color=Theme.BORDER_DEFAULT)
                    popup_ax.tick_params(labelsize=10, colors=Theme.TEXT_SECONDARY)
                    popup_ax.set_facecolor(Theme.BG_CARD)

                    # Copy scale
                    if ax.get_yscale() == 'log':
                        popup_ax.set_yscale('log')

                    for spine in popup_ax.spines.values():
                        spine.set_color(Theme.BORDER_DEFAULT)
                        spine.set_linewidth(0.5)

            # Hide unused axes
            for idx in range(len(orig_axes), num_rows * num_cols):
                row = idx // num_cols
                col = idx % num_cols
                popup_axes[row, col].set_visible(False)

        # Copy suptitle
        popup_fig.suptitle(fig._suptitle.get_text() if fig._suptitle else key.replace('_', ' '),
                          fontsize=14, fontweight='bold', color=Theme.TEXT_PRIMARY)
        popup_fig.tight_layout()
        popup_fig.subplots_adjust(top=0.92)

        # Add canvas and toolbar
        popup_canvas = FigureCanvasTkAgg(popup_fig, master=popup)
        popup_canvas.draw()
        popup_canvas.get_tk_widget().pack(fill="both", expand=True)

        # Add navigation toolbar for zoom/pan
        toolbar_frame = ttk.Frame(popup)
        toolbar_frame.pack(side="bottom", fill="x")
        toolbar = NavigationToolbar2Tk(popup_canvas, toolbar_frame)
        toolbar.update()

        # Add legend to first visible axis
        visible_axes = [ax for ax in popup_fig.get_axes() if ax.get_visible() and ax.get_lines()]
        if visible_axes:
            visible_axes[0].legend(loc='upper right', fontsize=9, ncol=2,
                                  facecolor=Theme.BG_CARD, edgecolor=Theme.BORDER_DEFAULT,
                                  labelcolor=Theme.TEXT_PRIMARY, framealpha=0.95)

        # Add close button
        close_btn = ctk.CTkButton(popup, text="Close (Esc)", command=popup.destroy,
                                 fg_color=Theme.ACCENT_PRIMARY, hover_color=Theme.ACCENT_SECONDARY,
                                 width=100, height=30) if HAS_CTK else tk.Button(popup, text="Close", command=popup.destroy)
        close_btn.pack(side="bottom", pady=5)

        # Bind Escape key to close
        popup.bind("<Escape>", lambda e: popup.destroy())

        # Focus the popup
        popup.focus_set()

    def get_filtered_data_for_bearing(self, torque, bearing_full, order, condition):
        """Get filtered data for a specific torque/bearing combination.

        Returns dict: {bearing_full: {direction: [cand_data, ...]}}
        """
        stage = self.stage_var.get()
        candidates = self.parse_candidate_selection()

        selected_directions = [d for d, v in self.direction_vars.items() if v.get()]
        if not selected_directions:
            return {}

        result = {bearing_full: {}}

        for file_num, meta in self.file_metadata.items():
            meta_bearing = meta.get('bearing_full', meta.get('bearing'))
            if meta_bearing != bearing_full:
                continue

            order_match = (order == "All") or (meta.get('order') == order)

            if (order_match and
                meta.get('stage') == stage and
                meta.get('torque') == torque and
                meta.get('condition', '').lower() == condition.lower() and
                meta.get('direction') in selected_directions):

                # LAZY LOADING: Load CSV data on demand
                if file_num not in self.csv_data:
                    csv_path = self._csv_files_list.get(file_num) if hasattr(self, '_csv_files_list') else None
                    if csv_path:
                        data = self.load_csv_data(csv_path)
                        if data:
                            self.csv_data[file_num] = data

                if file_num in self.csv_data:
                    direction = meta['direction']
                    if direction not in result[bearing_full]:
                        result[bearing_full][direction] = []

                    csv_data = self.csv_data[file_num]
                    csv_path = self._csv_files_list.get(file_num) if hasattr(self, '_csv_files_list') else None
                    for cand in candidates:
                        if cand <= len(csv_data):
                            cand_data = csv_data[cand - 1].copy()
                            cand_data['candidate'] = cand
                            cand_data['_source_info'] = {
                                'file_number': file_num,
                                'csv_path': csv_path,
                                'image_path': Path(self.data_folder) / f"{csv_path.stem}_Candidate{cand:06d}.png" if csv_path else None,
                                'candidate': cand,
                                'bearing': meta.get('bearing', ''),
                                'bearing_full': bearing_full,
                                'direction': direction,
                                'order': meta.get('order', ''),
                                'torque': torque,
                                'condition': condition,
                            }
                            result[bearing_full][direction].append(cand_data)

        # Remove empty directions
        result[bearing_full] = {k: v for k, v in result[bearing_full].items() if v}
        if not result[bearing_full]:
            return {}

        return result
    
    def get_data_for_export(self, torque, order, bearings, directions, candidates, condition=None):
        """Get filtered data for specific torque/order/condition combination for export.

        Returns dict: {bearing_full: {direction: [cand_data, ...]}}
        """
        stage = self.stage_var.get()
        # Use provided condition for export, or fall back to GUI selection
        if condition is None:
            condition = self.condition_var.get()

        result = {}
        for file_num, meta in self.file_metadata.items():
            bearing = meta.get('bearing')
            bearing_full = meta.get('bearing_full', bearing)

            # Check if this bearing is selected
            is_selected = any(b == bearing for b, _ in bearings)

            # Check order match
            order_match = (order == "All") or (meta.get('order') == order)

            if (is_selected and
                order_match and
                meta.get('stage') == stage and
                meta.get('torque') == torque and
                meta.get('condition', '').lower() == condition.lower() and
                meta.get('direction') in directions):

                # LAZY LOADING: Load CSV data on demand if not already loaded
                if file_num not in self.csv_data:
                    csv_path = self._csv_files_list.get(file_num) if hasattr(self, '_csv_files_list') else None
                    if csv_path:
                        data = self.load_csv_data(csv_path)
                        if data:
                            self.csv_data[file_num] = data

                if file_num in self.csv_data:
                    if bearing_full not in result:
                        result[bearing_full] = {}
                    direction = meta['direction']
                    if direction not in result[bearing_full]:
                        result[bearing_full][direction] = []

                    for cand_data in self.csv_data[file_num]:
                        if cand_data.get('candidate') in candidates:
                            result[bearing_full][direction].append(cand_data)

        return result

    def export_to_excel(self):
        """Show export options dialog with Torque, Order, Bearing, Direction selection.

        - Multiple torques = multiple xlsx files
        - Multiple orders = multiple sheets within each file
        - Bearings/Directions = columns (same logic as before)
        """
        # Check output mode - Scalar exports different data
        if self.output_mode.get() == "scalar":
            self.export_scalar_to_excel()
            return

        if not self.file_metadata:
            messagebox.showwarning("No Data", "No data to export. Load data first.")
            return

        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("Error", "Install pandas: pip install pandas openpyxl")
            return

        # Create export options dialog
        dialog = ctk.CTkToplevel(self.root) if HAS_CTK else tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.geometry("600x750")
        dialog.transient(self.root)
        dialog.grab_set()

        if HAS_CTK:
            dialog.configure(fg_color=Theme.BG_PRIMARY)

        # Main container with scroll
        main_frame = ctk.CTkScrollableFrame(dialog, fg_color=Theme.BG_SECONDARY) if HAS_CTK else tk.Frame(dialog)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_lbl = ctk.CTkLabel(main_frame, text="Export Options",
                                  font=ctk.CTkFont(size=16, weight="bold"),
                                  text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(main_frame, text="Export Options", font=("Arial", 14, "bold"))
        title_lbl.pack(pady=(10, 5))

        # Info label
        info_lbl = ctk.CTkLabel(main_frame, text="Torque × Condition → Separate Files | Orders → Separate Sheets",
                                 font=ctk.CTkFont(size=11),
                                 text_color=Theme.TEXT_MUTED) if HAS_CTK else tk.Label(main_frame, text="Torque × Condition = Files, Orders = Sheets")
        info_lbl.pack(pady=(0, 10))

        # Get available options from metadata
        all_bearings = set()
        all_directions = set()
        for meta in self.file_metadata.values():
            if meta.get('bearing'):
                all_bearings.add((meta.get('bearing'), meta.get('bearing_full', meta.get('bearing'))))
            if meta.get('direction'):
                all_directions.add(meta.get('direction'))

        bearings_list = sorted(all_bearings, key=lambda x: int(re.search(r'B(\d+)', x[0]).group(1)) if re.search(r'B(\d+)', x[0]) else 0)
        directions_list = sorted(all_directions)

        # === TORQUE SECTION ===
        torque_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Torque")
        torque_frame.pack(fill="x", padx=5, pady=5)

        torque_lbl = ctk.CTkLabel(torque_frame, text="Torques (each = separate file):",
                                   font=ctk.CTkFont(weight="bold"),
                                   text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(torque_frame, text="Torques:")
        torque_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        export_torque_vars = {}
        torque_check_frame = ctk.CTkFrame(torque_frame, fg_color="transparent") if HAS_CTK else tk.Frame(torque_frame)
        torque_check_frame.pack(fill="x", padx=10, pady=5)

        for i, t in enumerate(self.torques):
            var = tk.BooleanVar(value=(t == self.torque_var.get()))  # Default to current selection
            export_torque_vars[t] = var
            cb = ctk.CTkCheckBox(torque_check_frame, text=t, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(torque_check_frame, text=t, variable=var)
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=5, pady=2)

        # === CONDITION SECTION (Drive/Coast) ===
        condition_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Condition")
        condition_frame.pack(fill="x", padx=5, pady=5)

        condition_lbl = ctk.CTkLabel(condition_frame, text="Conditions (each = separate file per torque):",
                                      font=ctk.CTkFont(weight="bold"),
                                      text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(condition_frame, text="Conditions:")
        condition_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        export_condition_vars = {}
        condition_check_frame = ctk.CTkFrame(condition_frame, fg_color="transparent") if HAS_CTK else tk.Frame(condition_frame)
        condition_check_frame.pack(fill="x", padx=10, pady=5)

        for i, c in enumerate(self.conditions):
            var = tk.BooleanVar(value=(c == self.condition_var.get()))  # Default to current selection
            export_condition_vars[c] = var
            cb = ctk.CTkCheckBox(condition_check_frame, text=c, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(condition_check_frame, text=c, variable=var)
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=5, pady=2)

        # === ORDER SECTION ===
        order_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Order")
        order_frame.pack(fill="x", padx=5, pady=5)

        order_lbl = ctk.CTkLabel(order_frame, text="Orders (each = separate sheet in file):",
                                  font=ctk.CTkFont(weight="bold"),
                                  text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(order_frame, text="Orders:")
        order_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        export_order_vars = {}
        order_check_frame = ctk.CTkFrame(order_frame, fg_color="transparent") if HAS_CTK else tk.Frame(order_frame)
        order_check_frame.pack(fill="x", padx=10, pady=5)

        current_order = self.order_var.get()
        for i, o in enumerate(self.orders):
            var = tk.BooleanVar(value=(o == current_order or current_order == "All"))
            export_order_vars[o] = var
            cb = ctk.CTkCheckBox(order_check_frame, text=o, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(order_check_frame, text=o, variable=var)
            cb.grid(row=i // 6, column=i % 6, sticky="w", padx=5, pady=2)

        # === BEARINGS SECTION ===
        bearing_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Bearings")
        bearing_frame.pack(fill="x", padx=5, pady=5)

        bearing_lbl = ctk.CTkLabel(bearing_frame, text="Bearings to Export:",
                                    font=ctk.CTkFont(weight="bold"),
                                    text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(bearing_frame, text="Bearings:")
        bearing_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        export_bearing_vars = {}
        bearing_check_frame = ctk.CTkFrame(bearing_frame, fg_color="transparent") if HAS_CTK else tk.Frame(bearing_frame)
        bearing_check_frame.pack(fill="x", padx=10, pady=5)

        for i, (b_short, b_full) in enumerate(bearings_list):
            # Check if this bearing is currently selected in main GUI
            is_selected = self.bearing_vars.get(b_full, tk.BooleanVar(value=False)).get() if hasattr(self, 'bearing_vars') else True
            var = tk.BooleanVar(value=is_selected)
            export_bearing_vars[(b_short, b_full)] = var
            cb = ctk.CTkCheckBox(bearing_check_frame, text=b_short, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(bearing_check_frame, text=b_short, variable=var)
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=5, pady=2)

        # === DIRECTIONS SECTION ===
        dir_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Directions")
        dir_frame.pack(fill="x", padx=5, pady=5)

        dir_lbl = ctk.CTkLabel(dir_frame, text="Directions to Export:",
                                font=ctk.CTkFont(weight="bold"),
                                text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(dir_frame, text="Directions:")
        dir_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        export_dir_vars = {}
        dir_check_frame = ctk.CTkFrame(dir_frame, fg_color="transparent") if HAS_CTK else tk.Frame(dir_frame)
        dir_check_frame.pack(fill="x", padx=10, pady=5)

        for i, d in enumerate(directions_list):
            # Check if this direction is currently selected in main GUI
            is_selected = self.direction_vars.get(d, tk.BooleanVar(value=False)).get() if hasattr(self, 'direction_vars') else True
            var = tk.BooleanVar(value=is_selected)
            export_dir_vars[d] = var
            cb = ctk.CTkCheckBox(dir_check_frame, text=d, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(dir_check_frame, text=d, variable=var)
            cb.grid(row=i // 6, column=i % 6, sticky="w", padx=10, pady=2)

        # === DATA TYPE SECTION ===
        data_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Data")
        data_frame.pack(fill="x", padx=5, pady=5)

        data_lbl = ctk.CTkLabel(data_frame, text="Data to Export:",
                                 font=ctk.CTkFont(weight="bold"),
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(data_frame, text="Data:")
        data_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        data_check_frame = ctk.CTkFrame(data_frame, fg_color="transparent") if HAS_CTK else tk.Frame(data_frame)
        data_check_frame.pack(fill="x", padx=10, pady=5)

        export_mag_var = tk.BooleanVar(value=True)
        export_phase_var = tk.BooleanVar(value=True)
        export_real_var = tk.BooleanVar(value=False)
        export_imag_var = tk.BooleanVar(value=False)

        cb_mag = ctk.CTkCheckBox(data_check_frame, text="Magnitude", variable=export_mag_var,
                                  text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(data_check_frame, text="Magnitude", variable=export_mag_var)
        cb_mag.grid(row=0, column=0, sticky="w", padx=10, pady=2)

        cb_phase = ctk.CTkCheckBox(data_check_frame, text="Phase", variable=export_phase_var,
                                    text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(data_check_frame, text="Phase", variable=export_phase_var)
        cb_phase.grid(row=0, column=1, sticky="w", padx=10, pady=2)

        cb_real = ctk.CTkCheckBox(data_check_frame, text="Real", variable=export_real_var,
                                   text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(data_check_frame, text="Real", variable=export_real_var)
        cb_real.grid(row=0, column=2, sticky="w", padx=10, pady=2)

        cb_imag = ctk.CTkCheckBox(data_check_frame, text="Imaginary", variable=export_imag_var,
                                   text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(data_check_frame, text="Imaginary", variable=export_imag_var)
        cb_imag.grid(row=0, column=3, sticky="w", padx=10, pady=2)

        # === SCALE SECTION ===
        scale_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Scale")
        scale_frame.pack(fill="x", padx=5, pady=5)

        scale_lbl = ctk.CTkLabel(scale_frame, text="Magnitude Scale:",
                                  font=ctk.CTkFont(weight="bold"),
                                  text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Label(scale_frame, text="Scale:")
        scale_lbl.pack(anchor="w", padx=10, pady=(10, 5))

        scale_var = tk.StringVar(value="linear")
        scale_radio_frame = ctk.CTkFrame(scale_frame, fg_color="transparent") if HAS_CTK else tk.Frame(scale_frame)
        scale_radio_frame.pack(fill="x", padx=10, pady=5)

        rb_linear = ctk.CTkRadioButton(scale_radio_frame, text="Linear", variable=scale_var, value="linear",
                                        text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Radiobutton(scale_radio_frame, text="Linear", variable=scale_var, value="linear")
        rb_linear.pack(side="left", padx=10)

        rb_log = ctk.CTkRadioButton(scale_radio_frame, text="Log (dB)", variable=scale_var, value="log",
                                     text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Radiobutton(scale_radio_frame, text="Log (dB)", variable=scale_var, value="log")
        rb_log.pack(side="left", padx=10)

        # === CANDIDATES INFO ===
        cand_frame = ctk.CTkFrame(main_frame, fg_color=Theme.BG_CARD) if HAS_CTK else tk.LabelFrame(main_frame, text="Candidates")
        cand_frame.pack(fill="x", padx=5, pady=5)

        candidates = self.parse_candidate_selection()
        cand_info = ctk.CTkLabel(cand_frame, text=f"Candidates: {len(candidates)} selected\n({self.candidate_entry.get()})",
                                  text_color=Theme.TEXT_SECONDARY) if HAS_CTK else tk.Label(cand_frame, text=f"Candidates: {len(candidates)}")
        cand_info.pack(padx=10, pady=10)

        # === BUTTONS ===
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent") if HAS_CTK else tk.Frame(main_frame)
        btn_frame.pack(fill="x", padx=5, pady=15)

        export_result = {'proceed': False}

        def do_export():
            export_result['proceed'] = True
            export_result['torques'] = [t for t, v in export_torque_vars.items() if v.get()]
            export_result['conditions'] = [c for c, v in export_condition_vars.items() if v.get()]
            export_result['orders'] = [o for o, v in export_order_vars.items() if v.get()]
            export_result['bearings'] = [(b, bf) for (b, bf), v in export_bearing_vars.items() if v.get()]
            export_result['directions'] = [d for d, v in export_dir_vars.items() if v.get()]
            export_result['magnitude'] = export_mag_var.get()
            export_result['phase'] = export_phase_var.get()
            export_result['real'] = export_real_var.get()
            export_result['imaginary'] = export_imag_var.get()
            export_result['scale'] = scale_var.get()
            dialog.destroy()

        def cancel_export():
            dialog.destroy()

        btn_export = ctk.CTkButton(btn_frame, text="Export", command=do_export,
                                    fg_color=Theme.ACCENT_PRIMARY, hover_color="#0056b3",
                                    width=120, height=36) if HAS_CTK else tk.Button(btn_frame, text="Export", command=do_export)
        btn_export.pack(side="left", expand=True, padx=5)

        btn_cancel = ctk.CTkButton(btn_frame, text="Cancel", command=cancel_export,
                                    fg_color=Theme.TEXT_MUTED, hover_color=Theme.TEXT_SECONDARY,
                                    width=120, height=36) if HAS_CTK else tk.Button(btn_frame, text="Cancel", command=cancel_export)
        btn_cancel.pack(side="left", expand=True, padx=5)

        # Wait for dialog
        dialog.wait_window()

        if not export_result['proceed']:
            return

        # Validate selections
        if not export_result['torques']:
            messagebox.showwarning("No Selection", "Select at least one torque")
            return
        if not export_result['conditions']:
            messagebox.showwarning("No Selection", "Select at least one condition (Drive/Coast)")
            return
        if not export_result['orders']:
            messagebox.showwarning("No Selection", "Select at least one order")
            return
        if not export_result['bearings']:
            messagebox.showwarning("No Selection", "Select at least one bearing")
            return
        if not export_result['directions']:
            messagebox.showwarning("No Selection", "Select at least one direction")
            return
        if not any([export_result['magnitude'], export_result['phase'], export_result['real'], export_result['imaginary']]):
            messagebox.showwarning("No Selection", "Select at least one data type")
            return

        # Get output folder (for multiple files) or single file
        sel_torques = export_result['torques']
        sel_conditions = export_result['conditions']
        sel_orders = export_result['orders']
        sel_bearings = export_result['bearings']
        sel_directions = export_result['directions']

        # Calculate total files: torques × conditions
        total_files = len(sel_torques) * len(sel_conditions)

        if total_files > 1:
            # Multiple files - ask for folder
            output_folder = filedialog.askdirectory(title="Select Output Folder for Excel Files")
            if not output_folder:
                return
        else:
            # Single file - ask for file
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"BearingForce_{sel_torques[0]}_{sel_conditions[0]}.xlsx",
                title="Export Data"
            )
            if not filepath:
                return
            output_folder = str(Path(filepath).parent)

        # Do the export
        try:
            self.status_bar.set_status("Exporting... Loading CSV data on demand", Theme.ACCENT_WARNING)
            self.root.update()

            files_created = []
            files_skipped = []

            for torque in sel_torques:
                for condition in sel_conditions:
                    # Create filename with torque AND condition
                    if total_files > 1:
                        filepath = str(Path(output_folder) / f"BearingForce_{torque}_{condition}.xlsx")

                    # Collect all sheets data first (to avoid empty file issue)
                    sheets_data = {}

                    for order in sel_orders:
                        # Update status to show progress
                        self.status_bar.set_status(f"Loading data for {torque} {condition} Order {order}...", Theme.ACCENT_WARNING)
                        self.root.update()

                        # Get data for this torque/order/condition combination (will lazy-load CSVs)
                        filtered = self.get_data_for_export(torque, order, sel_bearings, sel_directions, candidates, condition)

                        if not filtered:
                            debug_print(f"No data for torque={torque}, condition={condition}, order={order}", "WARN")
                            continue

                        # Get frequency array
                        freq = None
                        for bearing_data in filtered.values():
                            for cands in bearing_data.values():
                                if cands:
                                    freq = cands[0]['frequencies']
                                    break
                            if freq is not None:
                                break

                        if freq is None:
                            continue

                        # Build rows
                        all_rows = []
                        for cand_num in candidates:
                            for freq_idx in range(len(freq)):
                                row = {'Candidate': cand_num, 'Frequency_Hz': freq[freq_idx]}

                                for b_short, b_full in sel_bearings:
                                    bearing_data = filtered.get(b_full, {})

                                    for direction in sel_directions:
                                        cands_list = bearing_data.get(direction, [])
                                        cand_data = None
                                        for cd in cands_list:
                                            if cd.get('candidate') == cand_num:
                                                cand_data = cd
                                                break

                                        # Add selected data types
                                        if export_result['magnitude']:
                                            col = f'{b_short}_{direction}_Mag'
                                            if cand_data and 'magnitude' in cand_data and freq_idx < len(cand_data['magnitude']):
                                                val = cand_data['magnitude'][freq_idx]
                                                if export_result['scale'] == 'log' and val > 0:
                                                    val = 20 * np.log10(val)
                                                row[col] = val
                                            else:
                                                row[col] = None

                                        if export_result['phase']:
                                            col = f'{b_short}_{direction}_Phase'
                                            if cand_data and 'phase' in cand_data and freq_idx < len(cand_data['phase']):
                                                row[col] = cand_data['phase'][freq_idx]
                                            else:
                                                row[col] = None

                                        if export_result['real']:
                                            col = f'{b_short}_{direction}_Real'
                                            if cand_data and 'real' in cand_data and freq_idx < len(cand_data['real']):
                                                row[col] = cand_data['real'][freq_idx]
                                            else:
                                                row[col] = None

                                        if export_result['imaginary']:
                                            col = f'{b_short}_{direction}_Imag'
                                            if cand_data and 'imaginary' in cand_data and freq_idx < len(cand_data['imaginary']):
                                                row[col] = cand_data['imaginary'][freq_idx]
                                            else:
                                                row[col] = None

                                all_rows.append(row)

                        if all_rows:
                            sheet_name = f"Order_{order}"[:31]
                            sheets_data[sheet_name] = pd.DataFrame(all_rows)

                    # Only create file if we have at least one sheet
                    if sheets_data:
                        self.status_bar.set_status(f"Writing {Path(filepath).name}...", Theme.ACCENT_WARNING)
                        self.root.update()

                        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                            for sheet_name, df in sheets_data.items():
                                df.to_excel(writer, sheet_name=sheet_name, index=False)

                        files_created.append(filepath)
                    else:
                        files_skipped.append(f"{torque}_{condition}")
                        debug_print(f"No data found for torque={torque}, condition={condition}, skipping file creation", "WARN")

            if files_created:
                scale_text = " (dB)" if export_result['scale'] == 'log' else ""
                skip_msg = ""
                if files_skipped:
                    skip_msg = f"\n\nNote: {len(files_skipped)} combination(s) had no matching data: {', '.join(files_skipped)}"

                if len(files_created) == 1:
                    self.status_bar.set_status(f"✓ Exported{scale_text} to {Path(files_created[0]).name}", Theme.ACCENT_SECONDARY)
                    messagebox.showinfo("Success", f"Exported to:\n{files_created[0]}{skip_msg}")
                else:
                    self.status_bar.set_status(f"✓ Exported{scale_text} {len(files_created)} files", Theme.ACCENT_SECONDARY)
                    messagebox.showinfo("Success", f"Exported {len(files_created)} files to:\n{output_folder}{skip_msg}")
            else:
                msg = "No data was exported.\n\nPossible reasons:\n"
                msg += "• Selected bearings/directions don't have data for chosen torques/conditions/orders\n"
                msg += "• The combination doesn't exist in your CSV files\n\n"
                msg += "Try selecting different options or check the debug log."
                messagebox.showwarning("No Data", msg)

        except Exception as e:
            self.status_bar.set_status("Export failed", Theme.ACCENT_ERROR)
            messagebox.showerror("Error", f"Export failed: {e}")

    def export_scalar_to_excel(self):
        """Export Scalar mode data (RMS and Peak values per frequency band) in wide format.

        Output format:
        Candidate | Bearing | Direction | Freq 0-1kHz RMS | Freq 1-3kHz RMS | ... | Freq 0-1kHz Peak | Freq 1-3kHz Peak | ...
        """
        filtered = self.get_filtered_data()
        if not filtered:
            messagebox.showwarning("No Data", "No data to export. Load data and plot first.")
            return

        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("Error", "Install pandas: pip install pandas openpyxl")
            return

        # Get file path
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Export Scalar Data (RMS & Peak)"
        )
        if not filepath:
            return

        try:
            self.status_bar.set_status("Exporting Scalar data...", Theme.ACCENT_WARNING)
            self.root.update()

            bearings = sorted(filtered.keys(),
                             key=lambda x: int(re.search(r'B(\d+)', x).group(1)) if re.search(r'B(\d+)', x) else 0)
            all_dirs = set()
            for bearing_data in filtered.values():
                all_dirs.update(bearing_data.keys())
            directions = sorted(all_dirs)

            candidates = self.parse_candidate_selection()
            band_labels = [b[2] for b in self.SCALAR_BANDS]

            all_rows = []

            for cand_num in candidates:
                for bearing_full in bearings:
                    short_match = re.search(r'(B\d+)', bearing_full)
                    bearing_short = short_match.group(1) if short_match else bearing_full
                    bearing_data = filtered.get(bearing_full, {})

                    for direction in directions:
                        cands_list = bearing_data.get(direction, [])
                        cand_data = None
                        for cd in cands_list:
                            if cd.get('candidate') == cand_num:
                                cand_data = cd
                                break

                        if cand_data and 'frequencies' in cand_data and 'magnitude' in cand_data:
                            scalar_vals = self.calculate_scalar_values(
                                cand_data['frequencies'], cand_data['magnitude'])

                            # Wide format: one row per Candidate/Bearing/Direction
                            row = {
                                'Candidate': cand_num,
                                'Bearing': bearing_short,
                                'Direction': direction
                            }

                            # Add RMS columns for each band
                            for band_label in band_labels:
                                col_name = f"Freq {band_label} RMS"
                                row[col_name] = scalar_vals[band_label]['rms']

                            # Add Peak columns for each band
                            for band_label in band_labels:
                                col_name = f"Freq {band_label} Peak"
                                row[col_name] = scalar_vals[band_label]['peak']

                            all_rows.append(row)

            df = pd.DataFrame(all_rows)
            df.to_excel(filepath, index=False)

            self.status_bar.set_status(f"✓ Exported Scalar data to {Path(filepath).name}", Theme.ACCENT_SECONDARY)
            messagebox.showinfo("Success", f"Exported Scalar data (RMS & Peak) to:\n{filepath}")

        except Exception as e:
            self.status_bar.set_status("Export failed", Theme.ACCENT_ERROR)
            messagebox.showerror("Error", f"Export failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    if HAS_CTK:
        root = ctk.CTk()
    else:
        root = tk.Tk()
        print("Install customtkinter for best experience: pip install customtkinter")
    
    # High DPI
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = BearingForceViewer(root)
    
    # Center window
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f'+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
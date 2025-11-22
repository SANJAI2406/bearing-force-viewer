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

# Try importing OCR
USE_EASYOCR = False
USE_PYTESSERACT = False
ocr_reader = None

try:
    import easyocr
    ocr_reader = easyocr.Reader(['en'], gpu=False, verbose=False)
    USE_EASYOCR = True
except ImportError:
    try:
        import pytesseract
        USE_PYTESSERACT = True
    except ImportError:
        pass

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
        
        # Show menu at mouse position
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
        # Torque
        self._create_filter_combo(self.filter_panel.content, "Torque", "torque")
        # Condition
        self._create_filter_combo(self.filter_panel.content, "Condition", "condition")
        # Order
        self._create_filter_combo(self.filter_panel.content, "Order", "order")
        
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
        condition = torque_match.group(2) if torque_match else "Unknown"

        number_match = re.search(r'--(\d+)\.csv$', filename)
        file_number = int(number_match.group(1)) if number_match else 0

        return {'stage': stage, 'torque': torque, 'condition': condition,
                'file_number': file_number, 'filename': filename}
    
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
            title_area = img.crop((0, 0, width, int(height * 0.07)))

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

        # Try direction: X Component, Y Component, Z Component
        direction_match = re.search(r'(X|Y|Z)\s*Component', text, re.IGNORECASE)
        if direction_match:
            result['direction'] = direction_match.group(1).upper()
            debug_print(f"    DIRECTION: {result['direction']}", "OCR")
        else:
            # Try Force X, Force Y, Force Z
            force_match = re.search(r'Force\s*(X|Y|Z)', text, re.IGNORECASE)
            if force_match:
                result['direction'] = force_match.group(1).upper()
                debug_print(f"    DIRECTION (Force): {result['direction']}", "OCR")
            else:
                debug_print(f"    NO DIRECTION found. Patterns tried: '(X|Y|Z) Component' and 'Force (X|Y|Z)'", "WARN")

        # Try order: Order 1, Order 2.0, Order 1_0, etc
        order_match = re.search(r'Order\s*(\d+)[._]?(\d*)', text, re.IGNORECASE)
        if order_match:
            order_int = order_match.group(1)
            order_dec = order_match.group(2) if order_match.group(2) else '0'
            result['order'] = f"{order_int}.{order_dec}"
            debug_print(f"    ORDER: {result['order']}", "OCR")
        else:
                debug_print("    NO ORDER found. Check if text contains Order 1, Order 2, etc.", "WARN")

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
        debug_print(f"  From filename: stage={meta.get('stage')}, torque={meta.get('torque')}, cond={meta.get('condition')}, num={file_num}", "FILE")

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
            return (file_num, meta, False)

        if not USE_EASYOCR and not USE_PYTESSERACT:
            debug_print(f"  NO OCR ENGINE - cannot read image", "WARN")
            return (file_num, meta, False)

        img_meta = self.extract_metadata_from_image_ocr(image_path)
        if img_meta and 'bearing' in img_meta:
            meta.update(img_meta)
            meta['bearing_full'] = f"{img_meta['bearing']} [{img_meta.get('bearing_desc', '')}]" if img_meta.get('bearing_desc') else img_meta['bearing']
            debug_print(f"  SUCCESS: {meta['bearing_full']}, Dir={meta.get('direction')}, Ord={meta.get('order')}", "SUCCESS")
            return (file_num, meta, True)
        else:
            debug_print(f"  FAILED: OCR did not extract bearing info", "WARN")

        return (file_num, meta, False)
    
    def _load_single_csv(self, csv_file):
        meta = self.parse_filename_info(csv_file.name)
        file_num = meta['file_number']
        data = self.load_csv_data(csv_file)
        return (file_num, data, csv_file)
    
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
        self.csv_paths = {}  # Store CSV paths for source validation

        csv_files = list(Path(folder).glob("*.csv"))
        if not csv_files:
            messagebox.showerror("Error", "No CSV files found")
            return

        total_files = len(csv_files)
        self.status_bar.set_status(f"Loading {total_files} files...", Theme.ACCENT_WARNING)
        self.status_bar.show_progress(0)
        self.root.update()

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
                self.status_bar.set_status(f"Detecting metadata: {completed}/{total_files}", Theme.ACCENT_WARNING)
                self.status_bar.update_progress(progress)
                self.root.update()

                file_num, meta, success = future.result()
                self.file_metadata[file_num] = meta
                if success:
                    ocr_success += 1

        # Print OCR summary
        debug_print("=" * 70, "INFO")
        debug_print(f"OCR SUMMARY: {ocr_success}/{total_files} files successfully processed", "INFO")
        if ocr_success < total_files:
            debug_print(f"  WARNING: {total_files - ocr_success} files FAILED OCR detection", "WARN")
        debug_print("=" * 70, "INFO")

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
                debug_print(f"  File {fn:03d}: {fname}", "WARN")
                debug_print(f"    bearing={fm.get('bearing', 'MISSING')}, dir={fm.get('direction', 'MISSING')}, order={fm.get('order', 'MISSING')}", "WARN")
            if len(failed_files) > 20:
                debug_print(f"  ... and {len(failed_files) - 20} more", "WARN")

        debug_print("=" * 70, "INFO")

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
        self.status_bar.set_status(f"Loading CSV data...", Theme.ACCENT_WARNING)
        self.root.update()

        # CSV loading (Phase 2: second half of progress)
        completed = 0
        max_workers = min(30, total_files)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(self._load_single_csv, csv_file): csv_file
                      for csv_file in csv_files}

            for future in as_completed(futures):
                completed += 1
                progress = 0.5 + (completed / (total_files * 2))  # Second half
                self.status_bar.set_status(f"Loading CSV: {completed}/{total_files}", Theme.ACCENT_WARNING)
                self.status_bar.update_progress(progress)
                self.root.update()

                file_num, data, csv_path = future.result()
                if data:
                    self.csv_data[file_num] = data
                    self.csv_paths = getattr(self, 'csv_paths', {})
                    self.csv_paths[file_num] = csv_path
                    if self.candidate_count == 0:
                        self.candidate_count = len(data)

        for csv_file in csv_files:
            meta = self.parse_filename_info(csv_file.name)
            file_num = meta['file_number']
            fm = self.file_metadata.get(file_num, {})
            all_bearings.add(fm.get('bearing_full', fm.get('bearing', 'Unknown')))
            all_directions.add(fm.get('direction', 'Unknown'))
            all_orders.add(fm.get('order', 'Unknown'))
            all_stages.add(fm.get('stage', '1'))
            all_torques.add(fm.get('torque', 'Unknown'))
            all_conditions.add(fm.get('condition', 'Unknown'))

        self.bearings = sorted([b for b in all_bearings if b != 'Unknown'],
                              key=lambda x: int(re.search(r'B(\d+)', x).group(1)) if re.search(r'B(\d+)', x) else 0)
        self.directions = sorted([d for d in all_directions if d != 'Unknown'])
        self.orders = sorted([o for o in all_orders if o != 'Unknown'],
                            key=lambda x: float(x) if x.replace('.','').isdigit() else 0)
        self.stages = sorted(all_stages)
        self.torques = sorted(all_torques)
        self.conditions = sorted(all_conditions)

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

        # Update combos
        if HAS_CTK:
            self.order_combo.configure(values=self.orders)
            self.stage_combo.configure(values=self.stages)
            self.torque_combo.configure(values=self.torques)
            self.condition_combo.configure(values=self.conditions)
        
        if self.orders:
            self.order_var.set(self.orders[0])
        if self.stages:
            self.stage_var.set(self.stages[0])
        if self.torques:
            self.torque_var.set(self.torques[0])
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

            if (is_selected and
                meta.get('order') == order and
                meta.get('stage') == stage and
                meta.get('torque') == torque and
                meta.get('condition') == condition and
                meta.get('direction') in selected_dirs):

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
    
    def clear_plot(self):
        self.fig.clear()
        if self.graph_tracker:
            self.graph_tracker.clear_data()
        self._show_welcome_screen()
        self.status_bar.set_status("Plot cleared", Theme.TEXT_MUTED)
    
    def plot_data(self):
        """Plot data with source validation support"""
        self.fig.clear()
        if self.graph_tracker:
            self.graph_tracker.clear_data()

        filtered = self.get_filtered_data()
        if not filtered:
            messagebox.showwarning("No Data", "No data matches filters.\nCheck selections.")
            return

        plot_type = self.plot_type.get()
        y_scale = self.y_scale.get()
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

        if plot_type == "both":
            num_rows = num_bearings * 2
        else:
            num_rows = num_bearings

        axes = self.fig.subplots(num_rows, num_dirs, squeeze=False)
        all_axes = []

        num_cands = len(self.parse_candidate_selection())
        colors = Theme.PLOT_COLORS

        for bearing_idx, bearing_full in enumerate(bearings):
            short_match = re.search(r'(B\d+)', bearing_full)
            bearing_short = short_match.group(1) if short_match else bearing_full

            bearing_data = filtered[bearing_full]

            for dir_idx, direction in enumerate(directions):
                cands = bearing_data.get(direction, [])

                if plot_type == "both":
                    ax_mag = axes[bearing_idx * 2, dir_idx]
                    ax_phase = axes[bearing_idx * 2 + 1, dir_idx]
                    all_axes.extend([ax_mag, ax_phase])
                elif plot_type == "magnitude":
                    ax_mag = axes[bearing_idx, dir_idx]
                    ax_phase = None
                    all_axes.append(ax_mag)
                else:
                    ax_mag = None
                    ax_phase = axes[bearing_idx, dir_idx]
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
                        
                        # Register for tracking with source info
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
                    ax_mag.set_title(f"{bearing_short} - {direction} - Order {order}", 
                                    fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
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
                        ax_phase.set_title(f"{bearing_short} - {direction} - Order {order}",
                                          fontsize=11, fontweight='bold', color=Theme.TEXT_PRIMARY)
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

        title = f"Bearing Force - {torque} {condition}"
        if num_cands > 1:
            title += f" ({num_cands} candidates)"
        self.fig.suptitle(title, fontsize=14, fontweight='bold', color=Theme.TEXT_PRIMARY)
        
        self.fig.tight_layout()
        self.fig.subplots_adjust(top=0.93)

        if self.graph_tracker:
            self.graph_tracker.setup_crosshairs(all_axes)

        self.canvas.draw()
        self.status_bar.set_status(f"✓ Plotted {num_cands} candidates • Right-click to validate", Theme.ACCENT_SECONDARY)
    
    def export_to_excel(self):
        """Show export options dialog and export data"""
        filtered = self.get_filtered_data()
        if not filtered:
            messagebox.showwarning("No Data", "No data to export. Load data and plot first.")
            return

        try:
            import pandas as pd
        except ImportError:
            messagebox.showerror("Error", "Install pandas: pip install pandas openpyxl")
            return

        # Create export options dialog
        dialog = ctk.CTkToplevel(self.root) if HAS_CTK else tk.Toplevel(self.root)
        dialog.title("Export Options")
        dialog.geometry("500x600")
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
        title_lbl.pack(pady=(10, 15))

        # Get available options from filtered data
        bearings = sorted(filtered.keys(),
                         key=lambda x: int(re.search(r'B(\d+)', x).group(1)) if re.search(r'B(\d+)', x) else 0)
        all_dirs = set()
        for bearing_data in filtered.values():
            all_dirs.update(bearing_data.keys())
        directions = sorted(all_dirs)

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

        for i, b in enumerate(bearings):
            var = tk.BooleanVar(value=True)
            export_bearing_vars[b] = var
            short_match = re.search(r'(B\d+)', b)
            display_name = short_match.group(1) if short_match else b
            cb = ctk.CTkCheckBox(bearing_check_frame, text=display_name, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(bearing_check_frame, text=display_name, variable=var)
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

        for i, d in enumerate(directions):
            var = tk.BooleanVar(value=True)
            export_dir_vars[d] = var
            cb = ctk.CTkCheckBox(dir_check_frame, text=d, variable=var,
                                 text_color=Theme.TEXT_PRIMARY) if HAS_CTK else tk.Checkbutton(dir_check_frame, text=d, variable=var)
            cb.grid(row=0, column=i, sticky="w", padx=10, pady=2)

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
            export_result['bearings'] = [b for b, v in export_bearing_vars.items() if v.get()]
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
        if not export_result['bearings']:
            messagebox.showwarning("No Selection", "Select at least one bearing")
            return
        if not export_result['directions']:
            messagebox.showwarning("No Selection", "Select at least one direction")
            return
        if not any([export_result['magnitude'], export_result['phase'], export_result['real'], export_result['imaginary']]):
            messagebox.showwarning("No Selection", "Select at least one data type")
            return

        # Get file path
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Export Data"
        )
        if not filepath:
            return

        # Do the export
        try:
            self.status_bar.set_status("Exporting...", Theme.ACCENT_WARNING)
            self.root.update()

            sel_bearings = export_result['bearings']
            sel_directions = export_result['directions']

            freq = None
            for bearing_data in filtered.values():
                for cands in bearing_data.values():
                    if cands:
                        freq = cands[0]['frequencies']
                        break
                if freq is not None:
                    break

            if freq is None:
                messagebox.showwarning("No Data", "No frequency data")
                return

            all_rows = []

            for cand_num in candidates:
                for freq_idx in range(len(freq)):
                    row = {'Candidate': cand_num, 'Frequency_Hz': freq[freq_idx]}

                    for bearing_full in sel_bearings:
                        short_match = re.search(r'(B\d+)', bearing_full)
                        bearing_short = short_match.group(1) if short_match else bearing_full
                        bearing_data = filtered.get(bearing_full, {})

                        for direction in sel_directions:
                            cands_list = bearing_data.get(direction, [])
                            cand_data = None
                            for cd in cands_list:
                                if cd.get('candidate') == cand_num:
                                    cand_data = cd
                                    break

                            # Add selected data types
                            if export_result['magnitude']:
                                col = f'{bearing_short}_{direction}_Mag'
                                if cand_data and 'magnitude' in cand_data and freq_idx < len(cand_data['magnitude']):
                                    val = cand_data['magnitude'][freq_idx]
                                    if export_result['scale'] == 'log' and val > 0:
                                        val = 20 * np.log10(val)  # Convert to dB
                                    row[col] = val
                                else:
                                    row[col] = None

                            if export_result['phase']:
                                col = f'{bearing_short}_{direction}_Phase'
                                if cand_data and 'phase' in cand_data and freq_idx < len(cand_data['phase']):
                                    row[col] = cand_data['phase'][freq_idx]
                                else:
                                    row[col] = None

                            if export_result['real']:
                                col = f'{bearing_short}_{direction}_Real'
                                if cand_data and 'real' in cand_data and freq_idx < len(cand_data['real']):
                                    row[col] = cand_data['real'][freq_idx]
                                else:
                                    row[col] = None

                            if export_result['imaginary']:
                                col = f'{bearing_short}_{direction}_Imag'
                                if cand_data and 'imaginary' in cand_data and freq_idx < len(cand_data['imaginary']):
                                    row[col] = cand_data['imaginary'][freq_idx]
                                else:
                                    row[col] = None

                    all_rows.append(row)

            df = pd.DataFrame(all_rows)
            df.to_excel(filepath, index=False)

            scale_text = " (dB)" if export_result['scale'] == 'log' else ""
            self.status_bar.set_status(f"✓ Exported{scale_text} to {Path(filepath).name}", Theme.ACCENT_SECONDARY)
            messagebox.showinfo("Success", f"Exported to:\n{filepath}")

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
# Bearing Force Bode Plot Viewer

A Python GUI application for visualizing bearing force data from Romax DOE (Design of Experiments) simulation results.

## Quick Start (Windows)

**Option 1: One-click install and run**
1. Download the repository
2. Double-click `install_and_run.bat`
3. Wait for dependencies to install (first time only)
4. Application will start automatically

**Option 2: Manual installation**
```bash
pip install -r requirements.txt
python bearing_force_viewer.py
```

## Features

- **Load and visualize** bearing force frequency response data from CSV files
- **Auto-detect metadata** from PNG images using OCR (bearing, direction, order)
- **Multi-bearing support** - view multiple bearings side-by-side in separate subplots
- **Interactive plots** with crosshair tracking and data point inspection
- **Right-click validation** - click any curve to open source CSV in Excel with row highlighted
- **Export options** - choose bearings, directions, data types (mag/phase/real/imag), and scale (linear/dB)
- **Parallel processing** for fast data loading (up to 30 workers)
- **Debug logging** - generates debug_log.txt in data folder for troubleshooting

## Requirements

- Python 3.8+
- Windows 10/11

Dependencies (installed automatically by `install_and_run.bat`):
```
numpy
matplotlib
customtkinter
pillow
easyocr
pandas
openpyxl
```

## Usage

1. Run the application:
   ```
   python bearing_force_viewer.py
   ```

2. Browse and select your Romax DOE data folder containing CSV and PNG files

3. Click "Load Data" to parse CSV files and detect metadata from images

4. Configure filters (bearing, direction, order, candidates) and click "Plot"

5. Right-click any curve to validate source data in Excel

## CSV File Format

The application expects CSV files with the following structure:
- Row 7: Frequency values
- Row 9+: Candidate data (5 rows per candidate: real, imaginary, magnitude, phase, empty)

## Export

Click "Export" to open the export options dialog where you can select:
- Which bearings to export
- Which directions (X, Y, Z)
- Data types (Magnitude, Phase, Real, Imaginary)
- Scale (Linear or Log/dB)

## Troubleshooting

If OCR detection fails:
1. Check the `debug_log_*.txt` file in your data folder
2. Ensure `easyocr` is installed: `pip install easyocr`
3. First OCR run downloads language models (~100MB)

## License

MIT License

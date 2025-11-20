# AVL Cameo Integration Guide

## Overview

Your OLHC application now supports **hybrid model management** - it can use either:
1. **AVL Cameo Pre-trained Models** (when DLL is available - office laptop)
2. **Custom RNN/GPR/XGBoost Models** (trained from your data - any laptop)

The application automatically detects which environment it's running in and adapts accordingly!

---

## Features Added

### 1. **Hybrid Model Manager**
- Automatically detects if AVL Cameo DLL is available
- Seamlessly switches between Cameo and custom models
- Same workflow regardless of which model type is used

### 2. **AVL Cameo Support**
- Direct integration with exported AVL Cameo models
- Pre-trained RNN models (no training needed)
- 10 input parameters, 19 output parameters
- Instant predictions

### 3. **Fallback to Custom Models**
- If Cameo DLL not found, uses standard RNN/GPR/XGBoost
- Trains from your data files
- Same prediction workflow

---

## How to Use - Step by Step

### **On Your Office Laptop (with AVL Cameo DLL)**

#### Step 1: Prepare AVL Cameo Files
Make sure these files exist in `D:\Python\OLHC\AVL_Cameo_model\`:
```
D:\Python\OLHC\AVL_Cameo_model\
├── Variant.py
├── Variant_demo.py  (optional)
└── Variant_x64.dll  (or Variant_x86.dll for 32-bit)
```

#### Step 2: Load Cameo Model in RNN Tab
1. Open OLHC application
2. Go to **RNN Tab**
3. Click **"Load Cameo Model"** button (green button)
4. You'll see a success message showing:
   - 10 input parameters loaded
   - 19 output parameters loaded
5. Status will show: "✓ AVL Cameo Model Loaded (19 outputs)"

#### Step 3: Activate Cameo Models
1. Click **"Load Cameo Models"** button (was "Build selected")
2. All 19 outputs will be marked as loaded (green indicators)
3. No training needed - models are pre-trained!

#### Step 4: Adjust Parameters (Design Reset)
1. Use the sliders in RNN tab to adjust input values:
   - B1_offset
   - B2_offset
   - B3_offset
   - B4_offset
   - B5_offset
   - Helix_Angle
   - Input_Stifness
   - Lead_Crown_Pinion
   - Lead_Slope_Pinion
   - Pressure_Angle

2. The sliders will use these default ranges:
   - B1_offset: 100 - 130
   - B2_offset: 15 - 30
   - B3_offset: -210 to -180
   - B4_offset: 120 - 150
   - B5_offset: 0 - 15
   - Helix_Angle: 15 - 25°
   - Input_Stifness: 0.5 - 1.5
   - Lead_Crown_Pinion: 5 - 15
   - Lead_Slope_Pinion: 0.04 - 0.09
   - Pressure_Angle: 18 - 22°

#### Step 5: Go to Generator Tab
1. Switch to **Generator Tab**
2. Click **"Manual Load Preset"** or use your existing preset
3. Make sure **"RNN"** checkbox is CHECKED
4. Click **"Run Failure Simulation"**

#### Step 6: View Results
- The simulation will use AVL Cameo models for predictions
- Results will show predicted outputs for all 19 parameters
- Much faster than training custom models!

---

### **On Your Personal Laptop (without AVL Cameo DLL)**

#### Automatic Fallback
1. When you click **"Load Cameo Model"**, you'll see:
   ```
   "AVL Cameo model DLL not found.

   Make sure the following files are in D:\Python\OLHC\AVL_Cameo_model\:
   - Variant.py
   - Variant_x64.dll (or Variant_x86.dll)"
   ```

2. Instead, use the normal workflow:
   - Click **"Import Data..."**
   - Load your CSV/Excel/TXT file
   - Select mode (Scalar/Dynamic)
   - Build models as usual with RNN/GPR/XGBoost

---

## AVL Cameo Model Details

### Input Parameters (10)
1. **B1_offset** - Bearing 1 offset position
2. **B2_offset** - Bearing 2 offset position
3. **B3_offset** - Bearing 3 offset position
4. **B4_offset** - Bearing 4 offset position
5. **B5_offset** - Bearing 5 offset position
6. **Helix_Angle** - Gear helix angle (degrees)
7. **Input_Stifness** - Input shaft stiffness
8. **Lead_Crown_Pinion** - Lead crown on pinion
9. **Lead_Slope_Pinion** - Lead slope on pinion
10. **Pressure_Angle** - Gear pressure angle (degrees)

### Output Parameters (19)
1. **Hull** - ConvexPlus model output
2. **B1_radialStiffnessX** - Bearing 1 radial stiffness X-direction
3. **B1_radialStiffnessY** - Bearing 1 radial stiffness Y-direction
4. **B2_radialStiffnessX** - Bearing 2 radial stiffness X-direction
5. **B2_radialStiffnessY** - Bearing 2 radial stiffness Y-direction
6. **B3_axialStiffness** - Bearing 3 axial stiffness
7. **B3_radialStiffnessX** - Bearing 3 radial stiffness X-direction
8. **B3_radialStiffnessY** - Bearing 3 radial stiffness Y-direction
9. **B4_radialStiffnessX** - Bearing 4 radial stiffness X-direction
10. **B4_radialStiffnessY** - Bearing 4 radial stiffness Y-direction
11. **B5_radialStiffnessX** - Bearing 5 radial stiffness X-direction
12. **B5_radialStiffnessY** - Bearing 5 radial stiffness Y-direction
13. **Fx** - Force in X-direction
14. **Fy** - Force in Y-direction
15. **Fz** - Force in Z-direction
16. **Linear_TE** - Linear transmission error
17. **Mx** - Moment about X-axis
18. **My** - Moment about Y-axis
19. **Tilt_TE** - Tilt transmission error

---

## Troubleshooting

### "AVL Cameo model DLL not found"
**Problem**: The DLL file is not accessible

**Solutions**:
1. Check that `Variant_x64.dll` exists in `D:\Python\OLHC\AVL_Cameo_model\`
2. Make sure the path is exactly: `D:\Python\OLHC\AVL_Cameo_model\`
3. If on 32-bit system, use `Variant_x86.dll` instead
4. If DLL is on office network drive, copy to local D: drive

### "Prediction failed for 'output_name'"
**Problem**: Cameo prediction error

**Solutions**:
1. Check that input values are within valid ranges
2. Verify all 10 input parameters have valid values
3. Check console output for detailed error message

### Design Reset Not Working
**Problem**: Sliders don't update parameters

**Solutions**:
1. Make sure you clicked "Load Cameo Models" first
2. Check that all sliders are visible in RNN tab
3. Try adjusting one slider manually to verify it works

---

## Files Modified

### New Files Created:
1. **AVL_Cameo_Wrapper.py** - Wrapper for AVL Cameo models
2. **CAMEO_INTEGRATION_GUIDE.md** - This guide

### Existing Files Modified:
1. **olhc.py** - Added Cameo support to RNN and Generator tabs

### Required Files (Office Laptop Only):
1. **D:\Python\OLHC\AVL_Cameo_model\Variant.py** - Cameo Python wrapper
2. **D:\Python\OLHC\AVL_Cameo_model\Variant_x64.dll** - Cameo DLL library

---

## Technical Details

### Hybrid Manager Architecture
```
HybridModelManager
├── AVLCameoModelWrapper (when DLL available)
│   ├── Detects DLL at import time
│   ├── Wraps Variant.py functions
│   └── Provides unified prediction API
└── DynamicModelTrainer (fallback)
    ├── GPR models
    ├── RNN models
    └── XGBoost models
```

### Prediction Flow
```
Generator Tab → RNN Checkbox Enabled
    ↓
Check if using Cameo (rnn_tab.using_cameo)
    ↓
Yes → hybrid_manager.predict(X, output_name)
    ↓
No → safe_predict(model, X, features)
```

---

## Benefits

### Using AVL Cameo Models:
✅ **No training time** - models are pre-trained
✅ **Instant predictions** - DLL is highly optimized
✅ **Industrial quality** - models validated by AVL
✅ **19 outputs** - comprehensive gear analysis
✅ **Consistent results** - same model everywhere

### Using Custom Models:
✅ **Works anywhere** - no DLL dependency
✅ **Flexible** - retrain with your own data
✅ **Multiple algorithms** - GPR, RNN, XGBoost
✅ **Customizable** - adjust PCA components, hyperparameters
✅ **Portable** - pure Python implementation

---

## Future Enhancements (Optional)

1. **Hybrid predictions**: Use both Cameo and custom models together
2. **Model comparison**: Compare Cameo vs custom model accuracy
3. **Automatic selection**: Let app choose best model based on data
4. **Model persistence**: Save/load Cameo configuration
5. **Batch predictions**: Predict multiple configurations at once

---

## Support

If you encounter issues:
1. Check this guide first
2. Verify file paths and DLL availability
3. Check console output for error messages
4. Try fallback to custom models

---

**Generated with Claude Code** 🤖
https://claude.com/claude-code

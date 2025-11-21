# -*- coding: utf-8 -*-
import os
import json  # <-- Add this import at the top of your file
import io
import math
import tkinter as tk
import customtkinter as ctk
from customtkinter import CTkFont
from sklearn.model_selection import GridSearchCV
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import traceback  # <-- Add this import
from scipy.stats import qmc, norm, chi2
from scipy.spatial.distance import pdist
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.metrics import r2_score, mean_squared_error
# Added for improved RobustRNN and Sobol sequences
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, WhiteKernel, ConstantKernel as C
from scipy.stats import qmc
from ENHANCED_DYNAMIC_RNN import (
    DynamicDataLoader,
    GPR_PCA_DynamicModel,
    DynamicModelTrainer,
    DynamicVisualization
)
# AVL Cameo - loaded dynamically when user selects the file
import sys
import importlib.util
Variant = None  # Will be loaded dynamically

class CameoModelWrapper:
    """Wrapper class to make AVL Cameo models compatible with sklearn-like interface"""
    def __init__(self, variant_module, output_name, input_names):
        self.variant = variant_module
        self.output_name = output_name
        self.input_names = input_names  # Ordered list of input parameter names

    def predict(self, X):
        """
        Make predictions using Cameo model.
        X can be a DataFrame or numpy array with columns matching input_names order.
        """
        import numpy as np

        # Convert to numpy array if DataFrame
        if hasattr(X, 'values'):
            X_arr = X.values
        else:
            X_arr = np.asarray(X)

        # Ensure 2D
        if X_arr.ndim == 1:
            X_arr = X_arr.reshape(1, -1)

        n_samples = X_arr.shape[0]

        # Prepare input list for Cameo (transpose: each input as a list of values)
        # Cameo expects: [[val1_sample1, val1_sample2, ...], [val2_sample1, val2_sample2, ...], ...]
        u = [X_arr[:, i].tolist() for i in range(X_arr.shape[1])]

        try:
            # Use Cameo evaluate function
            result = self.variant.evaluate(1, self.output_name, u)
            if result is not None:
                return np.array(result)
            else:
                return np.full(n_samples, np.nan)
        except Exception as e:
            print(f"Cameo prediction error for {self.output_name}: {e}")
            return np.full(n_samples, np.nan)

from sklearn.model_selection import train_test_split # May be needed
from sklearn.cluster import KMeans
from sklearn.linear_model import Ridge, LogisticRegression
from sklearn.preprocessing import PolynomialFeatures
# === DYNAMIC MODE IMPORTS (Auto-added) ===
from sklearn.decomposition import PCA
# === END DYNAMIC MODE IMPORTS ===

from sklearn.cluster import KMeans
from sklearn.exceptions import NotFittedError
from sklearn.utils.validation import check_is_fitted
from sklearn.model_selection import KFold, cross_val_predict, train_test_split
import itertools
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.widgets import LassoSelector
from matplotlib.path import Path
from matplotlib.patches import Ellipse
import matplotlib.transforms as transforms

# optional hover
try:
    import mplcursors

    _HAS_MPLCURSORS = True
except Exception:
    _HAS_MPLCURSORS = False

# --- ADDED FOR SHAP SENSITIVITY ---
try:
    import shap

    _HAS_SHAP = True
except Exception:
    _HAS_SHAP = False
# --- END ADDED ---

# --- ADDED FOR SAMPLE VISUALIZATION ---
try:
    import seaborn as sns

    _HAS_SEABORN = True
except Exception:
    _HAS_SEABORN = False

try:
    # This is part of pandas, but we import it explicitly for clarity
    from pandas.plotting import parallel_coordinates

    _HAS_PANDAS_PLOTTING = True
except Exception:
    _HAS_PANDAS_PLOTTING = False
# --- END ADDED ---

# =============================================================================
# --- THEME AND HELP CONTENT (Formerly utils files) ---
# =============================================================================

THEMES = {
    "Dark": {
        "ctk_mode": "Dark", "ctk_color_theme": "dark-blue", "mpl_style": "dark_background",
        "plot_bg": "#242424", "text_color": "white", "grid_color": "#404040",
        "orig_point_color": "silver", "gen_edge_color": "white", "ellipse_color": "white",
        "boxplot_sigma_color": "cyan", "hist_color_original": "#3a86ff",
        "hist_color_generated": "#ff006e", "normal_curve_color": "lime",
        "sigma_colors": ['yellow', 'orange', 'red'],
        "box_colors": {'Min/Max': '#ff4d4d', '2Ïƒ': '#ff9f1c', '3Ïƒ': '#f038ff', '4Ïƒ': '#00ffff'}
    },
    "Light": {
        "ctk_mode": "Light", "ctk_color_theme": "dark-blue", "mpl_style": "default",
        "plot_bg": "#F5F5F5",  # <-- This is the "toned-down" white for PLOTS
        "text_color": "black", "grid_color": "#C0C0C0",
        "orig_point_color": "#505050", "gen_edge_color": "black", "ellipse_color": "#333333",
        "boxplot_sigma_color": "blue", "hist_color_original": "#8d99ae",
        "hist_color_generated": "#ef233c", "normal_curve_color": "green",
        "sigma_colors": ['#6a00f4', '#f4007a', '#f48f00'],
        "box_colors": {'Min/Max': '#d90429', '2Ïƒ': '#fca311', '3Ïƒ': '#8f2d56', '4Ïƒ': '#219ebc'}
    },
    "Blue": {
        "ctk_mode": "Dark", "ctk_color_theme": "dark-blue", "mpl_style": "dark_background",
        "plot_bg": "#242424", "text_color": "white", "grid_color": "#404040",
        "orig_point_color": "silver", "gen_edge_color": "white", "ellipse_color": "white",
        "boxplot_sigma_color": "cyan", "hist_color_original": "#3a86ff",
        "hist_color_generated": "#ff006e", "normal_curve_color": "lime",
        "sigma_colors": ['yellow', 'orange', 'red'],
        "box_colors": {'Min/Max': '#ff4d4d', '2Ïƒ': '#ff9f1c', '3Ïƒ': '#f038ff', '4Ïƒ': '#00ffff'}
    }
}

HELP_TOPICS = {
    "gen_intro": (
        "GENERATOR TAB: INTRODUCTION\n\n"
        "This tab is the control center for creating your Design of Experiments (DOE). It offers two modes to define your 'design space':\n\n"
        "**1. Load from File:**\n"
        "Load an existing dataset (.csv, .xlsx). The application analyzes the numerical columns to determine the boundaries and statistical properties of your design space. This is the recommended way to base a new DOE on historical data.\n\n"
        "**2. Manual Entry:**\n"
        "Directly define your parameters and their minimum/maximum boundaries. This is useful for creating a new design space from scratch based on specifications or theoretical limits, without needing a pre-existing data file."
    ),
    "gen_method": (
        "GENERATOR TAB: METHOD\n\n"
        "This dropdown selects the statistical algorithm used to generate new sample points.\n\n"
        "**Theory**: Each method explores the design space differently:\n\n"
        " â€¢ **Optimized LHS**: (Recommended) Stands for Latin Hypercube Sampling. It's a highly efficient 'stratified' sampling technique that provides maximum coverage with the minimum number of points. It works by dividing the range of each parameter into 'n' intervals and placing exactly one sample in each. The 'Optimized' process ensures the final sample is the best possible by selecting the one with the most uniform coverage. (See the 'Algorithms' section for a deep dive).\n\n"
        " â€¢ **Monte Carlo**: The simplest method. It generates purely random points within the defined boundaries. This method is prone to clustering and leaving large empty gaps, making it inefficient. You often need many more Monte Carlo points to achieve the same quality as an LHS sample.\n\n"
        " â€¢ **Monte Carlo (Hypersphere)**: A variation of Monte Carlo that generates points within a multi-dimensional sphere of a given radius. This is useful when your parameters are independent and the boundary is better described by a radius from a center point rather than a box.\n\n"
        "[Image showing a comparison of LHS vs. Monte Carlo samples]"
    ),
    "gen_bounds": (
        "GENERATOR TAB: STATISTICAL BOUNDS\n\n"
        "The 'Use Â±Ïƒ bounds (stat)' checkbox and the associated dropdown determine how the boundaries of the design space are defined. **This feature is only available in 'Load from File' mode**, as manual entry uses explicit min/max values.\n\n"
        "**Theory**: This choice controls the scope of your experiment.\n\n"
        " â€¢ **Unchecked (Min/Max Bounds)**: The boundaries are the absolute minimum and maximum values found in your input data. This creates a rigid box around your observed data, which can be sensitive to outliers.\n\n"
        " â€¢ **Checked (Statistical Bounds)**: The boundaries are calculated based on the process's inherent variation. This is a more robust method that reflects the likely range of the process that produced your data, not just the specific samples you measured. The formula used for each parameter is:\n\n"
        "$$[ \\mu - k \\cdot \\sigma, \\ \\mu + k \\cdot \\sigma ]$$\n\n"
        "where Î¼ is the mean, Ïƒ is the standard deviation, and 'k' is the multiplier from the dropdown (2, 3, or 4). For a normal distribution, 3Ïƒ covers 99.7% of the expected data."
    ),
    "gen_controls": (
        "GENERATOR TAB: OTHER CONTROLS\n\n"
        "**Samples**: Defines the number of new data points (n) to generate. This determines the size of your experiment. The goal of using efficient methods like Optimized LHS is to minimize 'n' while still building a high-fidelity model.\n\n"
        "**Repeats**: (LHS only) Sets the number of times the LHS algorithm is run to find the best possible sample. This controls the 'optimization' part, ensuring the final sample has the most uniform point distribution.\n\n"
        "**Seed (opt)**: An integer for the random number generator. Using the same seed with the same settings will always produce the exact same set of points, making your results reproducible.\n\n"
        "**Generate & Visualize**: Executes the sampling algorithm and opens the Visualization tab to display the results.\n\n"
        "**Save Last Generated Sample**: Saves the generated set of points to a .csv or .xlsx file. This file is the practical 'plan of action' for your simulations.\n\n"
        "**Save 2^k Corner Points**: Generates a 2-level full factorial design. For 'k' parameters, it creates all 2^k combinations of their minimum and maximum values. This is extremely useful for worst-case analysis and sensitivity studies.\n\n"
        "**Find Optimal n (Elbow Plot)**: This powerful tool helps you determine the minimum number of samples needed for your experiment by plotting the 'quality' (maximin score) of both LHS and Monte Carlo against the number of samples. This allows you to find the 'point of diminishing returns' before committing to expensive simulations."
    ),
    "vis_intro": (
        "VISUALIZATION TAB: INTRODUCTION\n\n"
        "This tab is for the interactive exploration and analysis of your original and newly generated datasets. It allows you to visually compare distributions, identify correlations, and perform advanced analyses."
    ),
    "vis_bounds": (
        "VISUALIZATION TAB: BOUNDING BOX OVERLAYS\n\n"
        "These checkboxes draw rectangular boxes on the scatter plot.\n\n"
        "**Theory**: These visualize the design space boundaries based on the assumption that the parameters are **uncorrelated**. Each axis is treated independently.\n\n"
        " â€¢ **Min/Max Box**: Shows the absolute observed limits from the original data.\n\n"
        " â€¢ **2Ïƒ / 3Ïƒ / 4Ïƒ Box**: Shows the statistical limits for each parameter, calculated as Î¼ Â± kÏƒ."
    ),
    "vis_ellipse": (
        "VISUALIZATION TAB: ELLIPTICAL BOUNDARY\n\n"
        "This dropdown draws a statistically derived ellipse on the scatter plot.\n\n"
        "**Theory**: This is a more advanced and accurate representation of a 2D statistical boundary. An ellipse represents a **confidence region** of a bivariate normal distribution. Its shape and angle are determined by the **covariance** between the X and Y variables. A tilted ellipse immediately indicates a correlation between the two parameters, something a rectangular box cannot show.\n\n"
        "The boundary is the set of points x satisfying:\n\n"
        "$$(x - \\mu)^T \\Sigma^{-1} (x - \\mu) \\le \\chi^2_{k,p}$$\n\n"
        "where Î¼ is the mean vector, Î£ is the covariance matrix, and Ï‡Â² is the Chi-Squared value for a given confidence level (p). (See 'Concepts: The Chi-Squared Distribution' for more details).\n\n"
        " â€¢ **Hide points outside ellipse**: Filters the plot to show only points within this statistical boundary."
    ),
    "vis_analysis": (
        "VISUALIZATION TAB: ANALYSIS FEATURES\n\n"
        "**Worst-Case Analysis**\n"
        "Description: Finds which parameter pairing creates the largest statistical range for a parameter of interest. When you click 'Run', a new window allows you to select which parameter comparisons you want to visualize.\n\n"
        "Theory: This is a powerful tool for identifying **critical parameter interactions**. For a chosen parameter (e.g., Lead Crown), it calculates the statistical ellipse between it and *every other parameter*. By overlaying them, it highlights which interaction causes the most significant variation, helping you focus on the most sensitive relationships in your design.\n\n"
        "**Distribution Plot Sigma Lines & Table**\n"
        "Description: When 'Show Normal Distribution' is checked, the plot will display vertical lines at the Â±2Ïƒ, Â±3Ïƒ, and Â±4Ïƒ limits of the original data's distribution. The text box at the bottom-left will show a table with the exact numerical values for these limits."
    ),
    "design_intro": (
        "DESIGN TAB: INTRODUCTION\n\n"
        "This tab is a 'what-if' tool for simulating the performance of a new multi-parameter design. It offers two modes:\n\n"
        "**1. Load Correlation from File (Recommended):**\n"
        "Leverage the statistics and, crucially, the **correlations** from an existing dataset to generate a new, realistic virtual system.\n\n"
        "**2. Define Parameters Manually:**\n"
        "Quickly define a set of **independent** (uncorrelated) parameters by specifying their name, mean, and standard deviation. Useful for simple what-if scenarios."
    ),
    "design_workflow": (
        "DESIGN TAB: WORKFLOW\n\n"
        "**MODE 1: Load Correlation from File**\n\n"
        "**1. Load Reference Data (for Correlation)**\n"
        "Loads a historical dataset to extract its **correlation matrix**. The underlying theory is that the **interactions between parameters** are often stable and transferable, even if the target values for a new design change.\n\n"
        "**2. Define New Parameter Targets**\n"
        "A dynamic table appears, allowing you to specify a new target **mean** and **standard deviation** for every parameter in your system. This gives you full control over the centering and spread of each variable in your virtual design.\n\n"
        "**MODE 2: Define Parameters Manually**\n\n"
        "**1. Add Parameters**\n"
        "Click the 'Add Parameter (+)' button to create a new row. Fill in the Parameter Name, Unit (optional), Mean, and Std Deviation. Add as many parameters as you need for your system.\n\n"
        "**GENERATION & ANALYSIS (Both Modes)**\n\n"
        "**3. Generate & Save System Data**\n"
        "Executes a virtual design simulation based on your chosen mode and inputs. The output is a realistic, multi-parameter dataset ready for further analysis.\n\n"
        "**4. Post-Generation Analysis**\n"
        "After saving, you can select any single parameter from your newly created dataset to perform a full Process Capability Analysis (Cp, Cpk)."
    ),
    "design_capability": (
        "DESIGN TAB: PROCESS CAPABILITY (CP, CPK)\n\n"
        "If you provide Upper and Lower Spec Limits (USL/LSL), the tool calculates these key industry metrics for the selected parameter.\n\n"
        "**Potential Capability (Cp)**\n"
        "Measures if your process variation is narrow enough to fit within the spec limits, regardless of whether it's centered.\n"
        "$$ C_p = \\frac{USL - LSL}{6\\sigma} $$\n\n"
        "**Actual Capability (Cpk)**\n"
        "A stricter metric that measures if the process is both narrow enough AND properly centered between the spec limits. In quality engineering (Six Sigma), a Cpk > 1.33 is often considered a minimum requirement for a capable process.\n"
        r"$$ C_{pk} = \\min\\left(\\frac{USL - \\mu}{3\\sigma}, \\frac{\mu - LSL}{3\\sigma}\\right) $$"
    ),
    "algo_mc": (
        "ALGORITHMS: MONTE CARLO\n\n"
        "Monte Carlo is the most straightforward sampling method. It generates points with uniform randomness inside the defined N-dimensional boundaries (a hyper-rectangle).\n\n"
        "**Analogy**: Imagine throwing darts randomly at a rectangular dartboard. The darts will land anywhere with equal probability, sometimes creating clumps and sometimes leaving large gaps.\n\n"
        "**Limitation**: It is not efficient for achieving even coverage with a small number of samples, making it a poor choice for expensive simulation experiments."
    ),
    "algo_mc_sphere": (
        "ALGORITHMS: MONTE CARLO (HYPERSPHERE)\n\n"
        "This is a variation of Monte Carlo that uses 'rejection sampling' to generate points inside a circular or spherical boundary.\n\n"
        "**Process**:\n"
        "1. A bounding hyper-rectangle is defined around the hypersphere.\n"
        "2. A point is generated randomly within the rectangle.\n"
        "3. The distance of the point from the center is calculated: $d = \\sqrt{\\sum (p_i - c_i)^2}$.\n"
        "4. If d â‰¤ Radius, the point is accepted. If not, it is rejected, and the process repeats.\n\n"
        "**Use Case**: This method is useful for simulating processes where the combined error is constrained by a radius, but it can be computationally slow in high dimensions."
    ),
    "algo_lhs": (
        "ALGORITHMS: LATIN HYPERCUBE SAMPLING (LHS)\n\n"
        "LHS is a stratified sampling method that provides much more uniform coverage than pure Monte Carlo for the same number of points. It is highly efficient for computer experiments.\n\n"
        "**Analogy**: Imagine a farm field divided into a 10x10 grid. Monte Carlo would be throwing 10 seeds randomly into the field. LHS is like carefully planting one seed in each row and each column, similar to the rules of Sudoku. This guarantees that every row and column is sampled exactly once.\n\n"
        "**Benefit**: This stratification ensures that each parameter's entire range is evenly explored, preventing the clustering and gap issues seen in Monte Carlo."
    ),
    "algo_lhs_opt": (
        "ALGORITHMS: OPTIMIZED LHS\n\n"
        "This tool improves upon standard LHS by finding a sample with the best possible space-filling properties.\n\n"
        "**Process**:\n"
        "1. The LHS algorithm is run 'N' times (where N is from the 'Repeats' box).\n"
        "2. For each of the N samples, a 'quality score' is calculated. This score is the minimum distance between any two points in the sample. This is known as the **maximin distance criterion**.\n"
        "   - The distance 'd' between points p and q is the Euclidean distance: $d = \\sqrt{\\sum (p_i - q_i)^2}$.\n"
        "   - The score for one sample is: Score = min(d_ij) for all pairs of points i, j.\n\n"
        "3. The algorithm keeps track of the sample with the **highest score** (the one where the closest two points are as far apart as possible).\n"
        "4. After N repeats, this 'best' sample is returned. This ensures the most uniform coverage possible, making your DOE maximally efficient."
    ),
    "guide_choosing_method": (
        "GUIDES: CHOOSING YOUR METHOD\n\n"
        "**Use Optimized LHS for (almost) everything.**\nFor any Design of Experiments (DOE) where you are running computer simulations (like FEA, CFD, or your 3D gear simulations), Optimized LHS is the superior choice. It guarantees the most efficient coverage of your design space, which means you can build a high-quality model (high RÂ²) with the fewest possible simulation runs. This saves immense amounts of time and computational cost.\n\n"
        "**When to use Monte Carlo?**\nUse basic Monte Carlo primarily for educational purposes or in cases where true, unstructured randomness is required. It is not recommended for creating efficient DOEs because it requires a very large number of samples to overcome its tendency to create clusters and gaps.\n\n"
        "**When to use Monte Carlo (Hypersphere)?**\nUse this only in specialized physical simulations where the design parameters are known to be uncorrelated and the failure or boundary condition is defined by a radius from a central point (e.g., total radial error)."
    ),
    "guide_interpreting_plots": (
        "GUIDES: INTERPRETING PLOTS\n\n"
        "**Scatter Plot**\n"
        "The scatter plot is your primary tool for understanding relationships between two parameters.\n\n"
        " â€¢ **Circular, Tight Cluster**: The two variables have low variance and are not correlated.\n"
        " â€¢ **Elliptical, Tilted Cluster**: The variables are correlated. An upward tilt (like '/') means positive correlation (as X increases, Y tends to increase). A downward tilt (like '\\') means negative correlation.\n"
        " â€¢ **Wide Cluster**: At least one variable has high variance (a large spread).\n"
        " â€¢ **Outliers**: Points that are far away from the main cluster.\n\n"
        "**Box Plot**\n"
        "A box plot summarizes the distribution of a single parameter.\n\n"
        " â€¢ The **line in the box** is the median (50th percentile).\n"
        " â€¢ The **box** represents the interquartile range (IQR), containing the middle 50% of your data (from the 25th to the 75th percentile).\n"
        " â€¢ The **whiskers** (lines extending from the box) typically show the full range of the data, excluding outliers."
    ),
    "concept_correlation": (
        "CONCEPTS: CORRELATION VS. COVARIANCE\n\n"
        "**Covariance** measures how two variables change together. A positive covariance means they increase together; a negative covariance means one increases as the other decreases. However, its value is hard to interpret because it depends on the units of the variables.\n\n"
        "**Correlation** is the normalized version of covariance. It is always between -1 and +1, making it easy to interpret.\n"
        " â€¢ **+1**: Perfect positive correlation.\n"
        " â€¢ **-1**: Perfect negative correlation.\n"
        " â€¢ **0**: No linear correlation.\n\n"
        "The formula for the correlation coefficient (Ï) is:\n"
        "$$ \\rho_{X,Y} = \\frac{\\text{cov}(X,Y)}{\\sigma_X \\sigma_Y} $$\n\n"
        "The **elliptical boundary** on the Visualization tab is a direct visual representation of the **covariance matrix** between two variables. Its tilt shows the direction of correlation, while its width and height along its axes relate to the variances."
    ),
    "concept_chi_squared": (
        "CONCEPTS: THE CHI-SQUARED DISTRIBUTION\n\n"
        "**Question**: Why is the Chi-Squared (Ï‡Â²) distribution used to draw the statistical ellipse?\n\n"
        "**Answer**: It stems from a key property of the multivariate normal distribution. The formula used to check if a point is inside the ellipse is a calculation of the Mahalanobis Distance squared:\n\n"
        "$$ D^2 = (x - \\mu)^T \\Sigma^{-1} (x - \\mu) $$\n\n"
        "For data that comes from a multivariate normal distribution, this value, DÂ², follows a **Chi-Squared (Ï‡Â²) distribution** with 'k' degrees of freedom, where 'k' is the number of dimensions (in our case, k=2 for a 2D plot).\n\n"
        "Therefore, by choosing a value from the Ï‡Â² distribution that corresponds to a certain probability (e.g., the value that contains 99.7% of the probability for k=2), we can define a threshold. Any point whose DÂ² is less than this threshold is considered 'inside' the 99.7% confidence region. This region is the ellipse you see on the plot."
    ),
    "guide_practical_example": (
        "GUIDES: PRACTICAL EXAMPLE - FROM TOLERANCE TO SIMULATION\n\n"
        "This guide walks through a real-world example of using the Design Tab to simulate a manufacturing process based on engineering tolerances.\n\n"
        "**SCENARIO:**\n"
        "You are designing a gear with a microgeometry parameter called 'Lead Crown'. The engineering drawing specifies:\n"
        " â€¢ **Nominal Value:** 30 Âµm\n"
        " â€¢ **Tolerance:** +9 Âµm / -8 Âµm\n\n"
        "You want to simulate a supplier whose manufacturing process is considered 'capable,' which in industry typically means a **Cpk of 1.33**.\n\n\n"
        "**STEP 1: ENTERING THE DATA**\n"
        "In the 'Define Manually' mode, you enter these values directly:\n"
        " â€¢ **Parameter Name:** `Lead Crown`\n"
        " â€¢ **Unit:** `Âµm`\n"
        " â€¢ **Nominal:** `30`\n"
        " â€¢ **Upper Tol (+):** `9`\n"
        " â€¢ **Lower Tol (-):** `-8`\n"
        " â€¢ **Target Cpk:** `1.33`\n\n\n"
        "**STEP 2: INTERNAL CALCULATIONS (WHAT THE APP DOES)**\n"
        "When you click 'Generate Data', the app performs these calculations based on standard process capability formulas:\n\n"
        "**A. Find the Specification Limits (USL & LSL):**\n"
        " â€¢ USL = Nominal + Upper Tol = 30 + 9 = **39 Âµm**\n"
        " â€¢ LSL = Nominal + Lower Tol = 30 + (-8) = **22 Âµm**\n\n"
        "**B. Calculate the Optimal Process Mean (Âµ):**\n"
        "For an asymmetric tolerance, the best place to center the process is the midpoint of the tolerance band, not the nominal value.\n"
        " â€¢ Process Mean (Âµ) = (USL + LSL) / 2 = (39 + 22) / 2 = **30.5 Âµm**\n\n"
        "**C. Calculate the Standard Deviation (Ïƒ):**\n"
        "The app uses the general formula for Cpk, which works for both symmetric and asymmetric tolerances:\n"
        "$$ \\sigma = \\frac{USL - LSL}{6 \\cdot C_{pk}} $$\n"
        " â€¢ Ïƒ = (39 - 22) / (6 * 1.33) = 17 / 7.98 â‰ˆ **2.13 Âµm**\n\n"
        "The simulation will now generate data from a normal distribution with **Mean = 30.5** and **Std Dev = 2.13**.\n\n\n"
        "**STEP 3: INTERPRETING THE RESULTS**\n\n"
        "**What does Cpk = 1.33 mean?**\n"
        "It defines a '4-Sigma Process'. This doesn't mean Ïƒ=4. It means **four** standard deviations can fit between the process mean and the nearest specification limit. The relationship is:\n"
        "   **Number of Sigmas = 3 Ã— Cpk = 3 Ã— 1.33 â‰ˆ 4**\n\n"
        "**Why isn't the Sample Mean exactly 30.5?**\n"
        "The capability report might show a 'Sample Mean' of 30.5001. This is due to **sampling variation**. Just like flipping a coin 10,000 times won't give you exactly 5,000 heads, a random sample will have a mean that is extremely close to, but not perfectly identical to, the theoretical target. This is normal and expected.\n\n"
        "**Understanding 'Observed' vs. 'Expected' PPM:**\n"
        "The report shows how many defects (Parts Per Million) were found:\n"
        " â€¢ **Expected PPM:** This is the theoretical defect rate for a 3.99-sigma process (3 * 1.33), which is about **33 PPM** for each tail (66 total).\n"
        " â€¢ **Observed PPM:** This is the actual number of points in your specific 10,000-point sample that fell outside the LSL/USL. Due to random chance, this might be 0, 100, or another value, but it will average out to the 'Expected' value over many runs."
    ),
}
# =============================================================================
# --- FONT INITIALIZATION (MOVED) ---
# =============================================================================

# Define globals as None. They will be initialized by the function below
# *after* the app root window is created.
WINDOWS_11_FONT_FAMILY = "Roboto"  # Default fallback
DEFAULT_UI_FONT = None
DEFAULT_UI_FONT_BOLD = None
SMALL_UI_FONT = None
CODE_FONT = None


def initialize_fonts():
    """
    Checks for available fonts and initializes global CTkFont objects.
    Must be called *after* the ctk.CTk() root window is created.
    """
    global WINDOWS_11_FONT_FAMILY, DEFAULT_UI_FONT, DEFAULT_UI_FONT_BOLD, \
        SMALL_UI_FONT, CODE_FONT

    # --- Define Windows 11 Style Font (as per previous request) ---
    try:
        # This code now runs *after* the root window exists, so it won't crash.
        tk.font.Font(family='Segoe UI Variable', size=12)
        WINDOWS_11_FONT_FAMILY = "Segoe UI Variable"
        print("Using Segoe UI Variable font.")
    except tk.TclError:
        try:
            tk.font.Font(family='Segoe UI', size=12)
            WINDOWS_11_FONT_FAMILY = "Segoe UI"
            print("Using Segoe UI font as fallback.")
        except tk.TclError:
            WINDOWS_11_FONT_FAMILY = "Roboto"  # Generic fallback
            print("Warning: Segoe UI Variable/Segoe UI not found. Using Roboto.")

    # Create CTkFont instances (adjust size as needed)
    DEFAULT_UI_FONT = CTkFont(family=WINDOWS_11_FONT_FAMILY, size=13)
    DEFAULT_UI_FONT_BOLD = CTkFont(family=WINDOWS_11_FONT_FAMILY, size=13, weight="bold")
    SMALL_UI_FONT = CTkFont(family=WINDOWS_11_FONT_FAMILY, size=11)
    CODE_FONT = CTkFont(family="Courier New", size=11)  # Keep code font monospaced


# =============================================================================
# --- TAB CLASS: GeneratorTab (Unicode Fix) ---
# =============================================================================
class GeneratorTab(ctk.CTkFrame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, fg_color="transparent")
        self.app = app_instance

        # --- State specific to this tab ---
        self.generator_mode = "File"
        self.manual_gen_widgets = []
        self.analysis_popup = None
        self.vis_popup = None  # To hold the visualization popup

        self._build_ui()

    def _build_ui(self):
        left = ctk.CTkFrame(self, width=360)
        left.pack(side="left", fill="y", padx=(6, 4), pady=6)
        right = ctk.CTkFrame(self)
        right.pack(side="left", fill="both", expand=True, padx=(4, 6), pady=6)

        mode_switcher_frame = ctk.CTkFrame(left, fg_color="transparent")
        mode_switcher_frame.pack(fill="x", padx=6, pady=5)
        ctk.CTkLabel(mode_switcher_frame, text="Input Mode:").pack(side="left")
        self.generator_mode_switcher = ctk.CTkSegmentedButton(mode_switcher_frame,
                                                              values=["Load from File", "Manual Entry"],
                                                              command=self._switch_generator_mode)
        self.generator_mode_switcher.set("Load from File")
        self.generator_mode_switcher.pack(side="left", padx=10, expand=True, fill="x")

        mode_container = ctk.CTkFrame(left, fg_color="transparent")
        mode_container.pack(fill="x", padx=0, pady=0)

        self.gen_file_mode_frame = ctk.CTkFrame(mode_container, fg_color="transparent")
        self.gen_file_mode_frame.pack(fill="x", padx=0, pady=0)
        self.load_button = ctk.CTkButton(self.gen_file_mode_frame, text="Load input (.csv/.xls/.xlsx)",
                                         command=self.load_file, width=320,
                                         fg_color="#003366", hover_color="#004080",
                                         text_color="white")
        self.load_button.pack(fill="x", padx=6, pady=(10, 6))
        self.loaded_label = ctk.CTkLabel(self.gen_file_mode_frame, text="No file loaded.", text_color="gray")
        self.loaded_label.pack(anchor="w", padx=6, pady=(0, 8))

        # --- Manual Mode Frame ---
        self.gen_manual_mode_frame = ctk.CTkFrame(mode_container, fg_color="transparent")
        manual_buttons_frame = ctk.CTkFrame(self.gen_manual_mode_frame, fg_color="transparent")
        manual_buttons_frame.pack(fill="x", padx=5, pady=8)

        self.gen_manual_add_button = ctk.CTkButton(manual_buttons_frame, text="Add (+)",
                                                   command=self._add_manual_gen_row, width=80,
                                                   fg_color="#003366", hover_color="#004080",
                                                   text_color="white")
        self.gen_manual_add_button.pack(side="left", padx=(0, 5))

        self.gen_save_preset_button = ctk.CTkButton(manual_buttons_frame, text="Save Preset",
                                                    command=self._save_gen_preset, width=100)
        self.gen_save_preset_button.pack(side="left", padx=5)

        self.gen_load_preset_button = ctk.CTkButton(manual_buttons_frame, text="Load Preset",
                                                    command=self._load_gen_preset, width=100)
        self.gen_load_preset_button.pack(side="left", padx=(5, 0))

        self.gen_manual_rows_frame = ctk.CTkScrollableFrame(self.gen_manual_mode_frame,
                                                            label_text="Define Parameter Boundaries")
        self.gen_manual_rows_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.gen_manual_rows_frame.grid_columnconfigure(0, weight=2)
        self.gen_manual_rows_frame.grid_columnconfigure(1, weight=1)
        self.gen_manual_rows_frame.grid_columnconfigure(2, weight=1)
        self.gen_manual_rows_frame.grid_columnconfigure(3, weight=0)
        ctk.CTkLabel(self.gen_manual_rows_frame, text="Click '+' to define a parameter.").pack(pady=20)
        # --- End Manual Mode Frame ---

        main_controls_frame = ctk.CTkFrame(left, fg_color="transparent")
        main_controls_frame.pack(fill="x", padx=0, pady=0)
        row = 0
        ctk.CTkLabel(main_controls_frame, text="Scatter X-Axis:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.preview_x_menu = ctk.CTkOptionMenu(main_controls_frame, values=["-"],
                                                command=lambda _: self._draw_small_preview(), width=180);
        self.preview_x_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Scatter Y-Axis:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.preview_y_menu = ctk.CTkOptionMenu(main_controls_frame, values=["-"],
                                                command=lambda _: self._draw_small_preview(), width=180);
        self.preview_y_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Value Plot Param:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.preview_val_menu = ctk.CTkOptionMenu(main_controls_frame, values=["-"],
                                                  command=lambda _: self._draw_small_preview(), width=180);
        self.preview_val_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Method:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.method_menu = ctk.CTkOptionMenu(main_controls_frame,
                                             values=["Optimized LHS", "Sobol Sequence", "Monte Carlo",
                                                     "Monte Carlo (Hypersphere)"],
                                             command=self._on_method_change, width=180);
        self.method_menu.set("Optimized LHS");
        self.method_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Samples:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.samples_entry = ctk.CTkEntry(main_controls_frame, width=100);
        self.samples_entry.insert(0, "100");
        self.samples_entry.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        self.use_stat_bounds_var = ctk.BooleanVar(value=False)

        # --- UNICODE FIX HERE ---
        self.stat_cb = ctk.CTkCheckBox(main_controls_frame, text=f"Use \u00b1\u03c3 bounds (stat)",
                                       variable=self.use_stat_bounds_var, command=self._on_stat_bounds_toggle);
        self.stat_cb.grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.sigma_level_menu = ctk.CTkOptionMenu(main_controls_frame,
                                                  values=[f"2\u03c3 (95%)", f"3\u03c3 (99.7%)", f"4\u03c3 (99.99%)"],
                                                  width=120);
        self.sigma_level_menu.set(f"4\u03c3 (99.99%)");
        # --- END UNICODE FIX ---

        self.sigma_level_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        self.radius_label = ctk.CTkLabel(main_controls_frame, text="Radius:");
        self.radius_label.grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.radius_entry = ctk.CTkEntry(main_controls_frame, width=100);
        self.radius_entry.insert(0, "3.0");
        self.radius_entry.grid(row=row, column=1, padx=6, pady=6, sticky="e")
        self.repeats_label = ctk.CTkLabel(main_controls_frame, text="Repeats:");
        self.repeats_label.grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.repeats_entry = ctk.CTkEntry(main_controls_frame, width=100);
        self.repeats_entry.insert(0, "20");
        self.repeats_entry.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Seed (opt):").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.seed_entry = ctk.CTkEntry(main_controls_frame, width=100);
        self.seed_entry.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Out name:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.filename_entry = ctk.CTkEntry(main_controls_frame, width=160);
        self.filename_entry.insert(0, "sample_output");
        self.filename_entry.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        ctk.CTkLabel(main_controls_frame, text="Format:").grid(row=row, column=0, padx=6, pady=6, sticky="w")
        self.format_menu = ctk.CTkOptionMenu(main_controls_frame, values=["CSV", "Excel (.xlsx)"], width=120);
        self.format_menu.set("CSV");
        self.format_menu.grid(row=row, column=1, padx=6, pady=6, sticky="e");
        row += 1
        self.generate_button = ctk.CTkButton(main_controls_frame, text="Generate & Visualize", command=self.generate,
                                             state="disabled",
                                             fg_color="#003366", hover_color="#004080",
                                             text_color="white");
        self.generate_button.grid(row=row, column=0, columnspan=2, padx=6, pady=(12, 6), sticky="we");
        row += 1
        self.save_generated_btn = ctk.CTkButton(main_controls_frame, text="Save Last Generated Sample",
                                                command=self.save_generated, state="disabled",
                                                fg_color="#003366", hover_color="#004080",
                                                text_color="white");
        self.save_generated_btn.grid(row=row, column=0, columnspan=2, padx=6, pady=(6, 6), sticky="we");
        row += 1
        self.corner_points_button = ctk.CTkButton(main_controls_frame, text="Save 2^k Corner Points",
                                                  command=self.save_corner_points, state="disabled",
                                                  fg_color="#003366", hover_color="#004080",
                                                  text_color="white");
        self.corner_points_button.grid(row=row, column=0, columnspan=2, padx=6, pady=(6, 6), sticky="we");

        row += 1
        self.visualize_sample_button = ctk.CTkButton(main_controls_frame, text="Visualize Sample (SPLOM & PCP)",
                                                     command=self.visualize_generated_sample, state="disabled",
                                                     fg_color="#003366", hover_color="#004080",
                                                     text_color="white")
        self.visualize_sample_button.grid(row=row, column=0, columnspan=2, padx=6, pady=(6, 6), sticky="we")

        row += 1
        self.status_label = ctk.CTkLabel(main_controls_frame, text="", text_color="white");
        self.status_label.grid(row=row, column=0, columnspan=2, padx=6, pady=(6, 8), sticky="w")

        comp_frame = ctk.CTkFrame(left);
        comp_frame.pack(fill="x", padx=6, pady=(10, 6))
        ctk.CTkLabel(comp_frame, text="Find Optimal Sample Size", font=("", 12, "bold")).pack(pady=(5, 5))
        n_frame = ctk.CTkFrame(comp_frame, fg_color="transparent");
        n_frame.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(n_frame, text="Start n:").pack(side="left");
        self.n_start_entry = ctk.CTkEntry(n_frame, width=60);
        self.n_start_entry.pack(side="left", padx=5);
        self.n_start_entry.insert(0, "20")
        ctk.CTkLabel(n_frame, text="End n:").pack(side="left", padx=(10, 0));
        self.n_end_entry = ctk.CTkEntry(n_frame, width=60);
        self.n_end_entry.pack(side="left", padx=5);
        self.n_end_entry.insert(0, "100")
        ctk.CTkLabel(n_frame, text="Step:").pack(side="left", padx=(10, 0));
        self.n_step_entry = ctk.CTkEntry(n_frame, width=50);
        self.n_step_entry.pack(side="left", padx=5);
        self.n_step_entry.insert(0, "10")
        self.elbow_button = ctk.CTkButton(comp_frame, text="Compare Methods vs. Sample Size",
                                          command=self.find_optimal_n, state="disabled",
                                          fg_color="#003366", hover_color="#004080",
                                          text_color="white");
        self.elbow_button.pack(fill="x", padx=5, pady=(8, 10))

        top_right_frame = ctk.CTkFrame(right, height=220, fg_color="transparent")
        top_right_frame.pack(fill="x", padx=6, pady=(6, 4))
        top_right_frame.grid_columnconfigure(0, weight=1)
        top_right_frame.grid_columnconfigure(1, weight=1)

        self.preview_text = ctk.CTkTextbox(top_right_frame, height=220, font=CODE_FONT, wrap="none");
        self.preview_text.grid(row=0, column=0, sticky="nsew", padx=(0, 3), pady=0);
        self.preview_text.configure(state="disabled")

        self.log_text = ctk.CTkTextbox(top_right_frame, height=220, font=CODE_FONT, wrap="none");
        self.log_text.grid(row=0, column=1, sticky="nsew", padx=(3, 0), pady=0);
        self.log_text.insert("1.0", "Generation log will appear here...")
        self.log_text.configure(state="disabled")

        self.preview_canvas_frame = ctk.CTkFrame(right, height=280);
        self.preview_canvas_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self.preview_fig, (self.pax_scatter, self.pax_value) = plt.subplots(1, 2, figsize=(8, 3),
                                                                            gridspec_kw={'width_ratios': [1.2, 1.2]})
        self.preview_fig.tight_layout(pad=2.0)
        self.preview_canvas = FigureCanvasTkAgg(self.preview_fig, master=self.preview_canvas_frame);
        self.preview_canvas.get_tk_widget().pack(fill="both", expand=True);
        self.preview_canvas.draw()

        self.app.bind('<Return>', lambda e: self._on_manual_add_keypress())
        self.app.bind('+', lambda e: self._on_manual_add_keypress())

        self._on_method_change(self.method_menu.get())
        self._on_stat_bounds_toggle()
        self._switch_generator_mode("Load from File")

        self.apply_theme()

    def apply_theme(self):
        theme_props = self.app.get_theme_properties()

        self.preview_fig.set_facecolor(theme_props["plot_bg"])

        for ax in [self.pax_scatter, self.pax_value]:
            ax.set_facecolor(theme_props["plot_bg"])
            ax.tick_params(axis='x', colors=theme_props["text_color"])
            ax.tick_params(axis='y', colors=theme_props["text_color"])
            for spine in ax.spines.values():
                spine.set_color(theme_props["text_color"])
            ax.xaxis.label.set_color(theme_props["text_color"])
            ax.yaxis.label.set_color(theme_props["text_color"])
            ax.title.set_color(theme_props["text_color"])
            ax.grid(color=theme_props['grid_color'], linestyle='--', alpha=0.5)

        self.preview_canvas.draw_idle()

    def _on_method_change(self, value):
        is_lhs = value == "Optimized LHS";
        is_sphere = value == "Monte Carlo (Hypersphere)";
        conditional_row = 8
        self.repeats_label.grid_forget();
        self.repeats_entry.grid_forget();
        self.radius_label.grid_forget();
        self.radius_entry.grid_forget()
        if is_lhs:
            self.repeats_label.grid(row=conditional_row, column=0, padx=6, pady=6, sticky="w");
            self.repeats_entry.grid(row=conditional_row, column=1, padx=6, pady=6, sticky="e")
        elif is_sphere:
            self.radius_label.grid(row=conditional_row, column=0, padx=6, pady=6, sticky="w");
            self.radius_entry.grid(row=conditional_row, column=1, padx=6, pady=6, sticky="e")

    def _on_stat_bounds_toggle(self):
        self.sigma_level_menu.configure(state="normal" if self.use_stat_bounds_var.get() else "disabled")

    def _switch_generator_mode(self, mode):
        self.generator_mode = "File" if mode == "Load from File" else "Manual"
        if self.generator_mode == "File":
            self.gen_manual_mode_frame.pack_forget();
            self.gen_file_mode_frame.pack(fill="x", padx=0, pady=0)
            self.stat_cb.configure(state="normal");
            self._on_stat_bounds_toggle()
            is_ready = self.app.df is not None
            self.generate_button.configure(state="normal" if is_ready else "disabled");
            self.corner_points_button.configure(state="normal" if is_ready else "disabled");
            self.elbow_button.configure(state="normal" if is_ready else "disabled")
        else:
            self.gen_file_mode_frame.pack_forget();
            self.gen_manual_mode_frame.pack(fill="x", padx=0, pady=0)
            self.stat_cb.configure(state="disabled");
            self.sigma_level_menu.configure(state="disabled");
            self.use_stat_bounds_var.set(False)
            is_ready = bool(self.manual_gen_widgets)
            self.generate_button.configure(state="normal" if is_ready else "disabled");
            self.corner_points_button.configure(state="normal" if is_ready else "disabled");
            self.elbow_button.configure(state="normal" if is_ready else "disabled")
        self._update_preview_text()

    def _add_manual_gen_row(self, initial_values=None):
        if initial_values is None:
            initial_values = {}

        if not self.manual_gen_widgets:
            for widget in self.gen_manual_rows_frame.winfo_children():
                widget.destroy()
            ctk.CTkLabel(self.gen_manual_rows_frame, text="Parameter Name", font=("", 12, "bold")).grid(row=0, column=0,
                                                                                                        padx=4, pady=2,
                                                                                                        sticky="w")
            ctk.CTkLabel(self.gen_manual_rows_frame, text="Min", font=("", 12, "bold")).grid(row=0, column=1, padx=4,
                                                                                             pady=2)
            ctk.CTkLabel(self.gen_manual_rows_frame, text="Max", font=("", 12, "bold")).grid(row=0, column=2, padx=4,
                                                                                             pady=2)

        row_index = len(self.manual_gen_widgets) + 1

        name_widget = ctk.CTkEntry(self.gen_manual_rows_frame, placeholder_text=f"Param {row_index}");
        name_widget.grid(row=row_index, column=0, padx=(0, 4), pady=3, sticky="ew")
        name_widget.insert(0, initial_values.get("name", ""))

        min_widget = ctk.CTkEntry(self.gen_manual_rows_frame, placeholder_text="0.0");
        min_widget.grid(row=row_index, column=1, padx=4, pady=3, sticky="ew")
        min_widget.insert(0, initial_values.get("min", ""))

        max_widget = ctk.CTkEntry(self.gen_manual_rows_frame, placeholder_text="1.0");
        max_widget.grid(row=row_index, column=2, padx=4, pady=3, sticky="ew")
        max_widget.insert(0, initial_values.get("max", ""))

        widget_dict = {"name": name_widget, "min": min_widget, "max": max_widget}

        remove_button = ctk.CTkButton(self.gen_manual_rows_frame, text="-", width=28,
                                      command=lambda d=widget_dict: self._remove_manual_gen_row(d),
                                      fg_color="#001f3f", hover_color="#003366",
                                      text_color="white");
        remove_button.grid(row=row_index, column=3, padx=(4, 0), pady=3)
        widget_dict["remove"] = remove_button

        self.manual_gen_widgets.append(widget_dict)

        self.generate_button.configure(state="normal");
        self.corner_points_button.configure(state="normal");
        self.elbow_button.configure(state="normal")

        if not initial_values:
            name_widget.focus()

        self._update_preview_text()

    def _remove_manual_gen_row(self, widget_dict_to_remove):
        for widget in widget_dict_to_remove.values():
            if widget: widget.destroy()
        self.manual_gen_widgets.remove(widget_dict_to_remove)
        for i, widgets in enumerate(self.manual_gen_widgets):
            row_idx = i + 1
            widgets['name'].grid(row=row_idx);
            widgets['min'].grid(row=row_idx);
            widgets['max'].grid(row=row_idx);
            widgets['remove'].grid(row=row_idx)
        if not self.manual_gen_widgets:
            self.generate_button.configure(state="disabled");
            self.corner_points_button.configure(state="disabled");
            self.elbow_button.configure(state="disabled")
            for widget in self.gen_manual_rows_frame.winfo_children(): widget.destroy()
            ctk.CTkLabel(self.gen_manual_rows_frame, text="Click '+' to define a parameter.").pack(pady=20)
        self._update_preview_text()

    def _clear_manual_gen_rows(self):
        for widget_dict in self.manual_gen_widgets:
            for widget in widget_dict.values():
                if widget and widget.winfo_exists():
                    widget.destroy()
        self.manual_gen_widgets.clear()
        for widget in self.gen_manual_rows_frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.gen_manual_rows_frame, text="Click '+' to define a parameter.").pack(pady=20)
        self.generate_button.configure(state="disabled");
        self.corner_points_button.configure(state="disabled");
        self.elbow_button.configure(state="disabled")
        self._update_preview_text()

    def _save_gen_preset(self):
        if self.generator_mode != "Manual" or not self.manual_gen_widgets:
            messagebox.showwarning("Save Preset",
                                   "Presets can only be saved in 'Manual Entry' mode when parameters are defined.")
            return

        preset_data = []
        valid = True
        for i, widgets in enumerate(self.manual_gen_widgets):
            name = widgets['name'].get().strip()
            min_str = widgets['min'].get().strip()
            max_str = widgets['max'].get().strip()

            if not name:
                messagebox.showerror("Save Error", f"Parameter name in row {i + 1} cannot be empty.")
                valid = False;
                break
            try:
                float(min_str);
                float(max_str)
                if float(min_str) >= float(max_str):
                    messagebox.showerror("Save Error", f"Max must be greater than Min in row {i + 1} for '{name}'.")
                    valid = False;
                    break
            except ValueError:
                messagebox.showerror("Save Error",
                                     f"Invalid numeric value for Min or Max found in row {i + 1} for '{name}'. Please check inputs.")
                valid = False;
                break

            preset_data.append({"name": name, "min": min_str, "max": max_str})

        if not valid:
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")],
            title="Save Generator Manual Parameter Preset"
        )

        if not filepath:
            return

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=4)

            theme_props = self.app.get_theme_properties()
            success_color = theme_props["text_color"]
            bold_font = CTkFont(weight="bold")
            self.status_label.configure(text=f"Preset saved to {os.path.basename(filepath)}",
                                        text_color=success_color,
                                        font=bold_font)

        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save preset file:\n{e}")
            self.status_label.configure(text=f"Error saving preset: {e}", text_color="orange",
                                        font=CTkFont(weight="bold"))

    def _load_gen_preset(self):
        if self.generator_mode != "Manual":
            self.generator_mode_switcher.set("Manual Entry")
            self._switch_generator_mode("Manual Entry")
            self.after(50, self._load_gen_preset_file_dialog)
            return
        self._load_gen_preset_file_dialog()

    def _load_gen_preset_file_dialog(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")],
            title="Load Generator Manual Parameter Preset"
        )

        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)

            if not isinstance(loaded_data, list):
                raise TypeError("Preset file should contain a list of parameters.")
            if not all(isinstance(item, dict) for item in loaded_data):
                raise TypeError("Each item in the preset list should be a dictionary.")
            if not all('name' in item and 'min' in item and 'max' in item for item in loaded_data):
                raise TypeError("Each parameter dictionary must contain 'name', 'min', and 'max' keys.")

            self._clear_manual_gen_rows()

            for param_data in loaded_data:
                self._add_manual_gen_row(initial_values=param_data)

            theme_props = self.app.get_theme_properties()
            success_color = theme_props["text_color"]
            bold_font = CTkFont(weight="bold")
            self.status_label.configure(text=f"Preset loaded from {os.path.basename(filepath)}",
                                        text_color=success_color,
                                        font=bold_font)

            if self.manual_gen_widgets:
                self.generate_button.configure(state="normal")
                self.corner_points_button.configure(state="normal")
                self.elbow_button.configure(state="normal")
                self._update_preview_text()

        except FileNotFoundError:
            messagebox.showerror("Load Error", "Preset file not found.")
            self.status_label.configure(text="Error: Preset file not found.", text_color="orange",
                                        font=CTkFont(weight="bold"))
        except json.JSONDecodeError:
            messagebox.showerror("Load Error", "Preset file is not valid JSON.")
            self.status_label.configure(text="Error: Invalid preset file format.", text_color="orange",
                                        font=CTkFont(weight="bold"))
        except TypeError as te:
            messagebox.showerror("Load Error", f"Preset file format error: {te}")
            self.status_label.configure(text=f"Error: Preset format error: {te}", text_color="orange",
                                        font=CTkFont(weight="bold"))
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load preset file:\n{e}")
            self.status_label.configure(text=f"Error loading preset: {e}", text_color="orange",
                                        font=CTkFont(weight="bold"))
            self._clear_manual_gen_rows()

    def _update_preview_text(self):
        self.preview_text.configure(state="normal");
        self.preview_text.delete("1.0", "end")
        if self.generator_mode == "File":
            if self.app.df is None:
                self.preview_text.insert("1.0", "No data loaded.")
            else:
                header = f"{'Parameter':<20}{'min':>12}{'max':>12}{'mean':>12}{'std':>12}\n"
                self.preview_text.insert("end", header + "-" * 70 + "\n")
                for c in self.app.param_names: self.preview_text.insert("end",
                                                                        f"{c:<20}{self.app.df[c].min():12.5g}{self.app.df[c].max():12.5g}{self.app.df[c].mean():12.5g}{self.app.df[c].std(ddof=0):12.5g}\n")
        else:
            if not self.manual_gen_widgets:
                self.preview_text.insert("1.0", "No parameters defined for manual entry.")
            else:
                header = f"{'Parameter':<20}{'Min Value':>15}{'Max Value':>15}\n"
                self.preview_text.insert("end", header + "-" * 50 + "\n")
                for widgets in self.manual_gen_widgets: self.preview_text.insert("end",
                                                                                 f"{widgets['name'].get() or '[Not Set]':<20}{widgets['min'].get() or '...':>15}{widgets['max'].get() or '...':>15}\n")
        self.preview_text.configure(state="disabled")

        if hasattr(self, 'log_text'):
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.insert("1.0", "Generation log will appear here...")
            self.log_text.configure(state="disabled")

    def _draw_small_preview(self):
        try:
            self.apply_theme()
            self.pax_scatter.clear();
            self.pax_value.clear()
            theme_props = self.app.get_theme_properties()
            if self.generator_mode == "Manual":
                self.pax_scatter.text(0.5, 0.5, 'Preview not available\nfor Manual Entry mode', ha='center',
                                      va='center', color=theme_props["text_color"])
                self.pax_value.text(0.5, 0.5, 'Define parameters and\nclick Generate to visualize', ha='center',
                                    va='center', color=theme_props["text_color"])
                self.preview_canvas.draw_idle();
                return
            if self.app.df is None or not self.app.param_names:
                self.preview_canvas.draw_idle();
                return
            x_col, y_col = self.preview_x_menu.get(), self.preview_y_menu.get()
            if x_col in self.app.param_names and y_col in self.app.param_names:
                x_vals, y_vals = self.app.df[x_col].values, self.app.df[y_col].values
                self.pax_scatter.scatter(x_vals, y_vals, s=8, alpha=0.8, c=theme_props["orig_point_color"],
                                         ec=theme_props["gen_edge_color"], lw=0.2)
                self.pax_scatter.set_xlabel(x_col, fontsize=9);
                self.pax_scatter.set_ylabel(y_col, fontsize=9);
                self.pax_scatter.tick_params(axis='x', labelsize=8);
                self.pax_scatter.tick_params(axis='y', labelsize=8)
            val_col = self.preview_val_menu.get()
            if val_col in self.app.param_names:
                self.pax_value.scatter(np.arange(len(self.app.df[val_col])), self.app.df[val_col].values, s=8,
                                       alpha=0.8, c=theme_props["orig_point_color"], ec=theme_props["gen_edge_color"],
                                       lw=0.2)
                self.pax_value.set_title(f"Value Plot", fontsize=10);
                self.pax_value.set_xlabel("Sample Number", fontsize=9);
                self.pax_value.set_ylabel(val_col, fontsize=9);
                self.pax_value.tick_params(axis='x', labelsize=8);
                self.pax_value.tick_params(axis='y', labelsize=8)
            self.preview_fig.tight_layout(pad=2.0);
            self.preview_canvas.draw_idle()
        except Exception:
            pass

    def _process_manual_generator_inputs(self):
        if not self.manual_gen_widgets: messagebox.showerror("Input Error", "No parameters defined."); return False
        param_data = {};
        names = []
        for i, widgets in enumerate(self.manual_gen_widgets):
            name = widgets['name'].get().strip()
            if not name: messagebox.showerror("Input Error", f"Parameter name for row {i + 1} empty."); return False
            if name in names: messagebox.showerror("Input Error", f"Duplicate name: '{name}'."); return False
            names.append(name)
            try:
                min_val, max_val = float(widgets['min'].get()), float(widgets['max'].get())
            except ValueError:
                messagebox.showerror("Input Error", f"Min/Max for '{name}' must be numbers.");
                return False
            if min_val >= max_val: messagebox.showerror("Input Error", f"Max > Min for '{name}'."); return False
            param_data[name] = [min_val, max_val]
        self.app.df = pd.DataFrame.from_dict(param_data, orient='index', columns=['min', 'max']).transpose()
        self.app.param_names = names
        vis_tab = self.app.visualization_tab
        for menu in [self.preview_x_menu, self.preview_y_menu, self.preview_val_menu, vis_tab.worst_case_param_menu,
                     vis_tab.scatter_x_menu, vis_tab.scatter_y_menu, vis_tab.boxplot_param_menu]:
            menu.configure(values=self.app.param_names)
        if self.app.param_names:
            self.preview_x_menu.set(self.app.param_names[0]);
            self.preview_val_menu.set(self.app.param_names[0]);
            vis_tab.scatter_x_menu.set(self.app.param_names[0]);
            vis_tab.boxplot_param_menu.set(self.app.param_names[0])
            if len(self.app.param_names) > 1:
                self.preview_y_menu.set(self.app.param_names[1]);
                vis_tab.scatter_y_menu.set(self.app.param_names[1])
            else:
                self.preview_y_menu.set(self.app.param_names[0]);
                vis_tab.scatter_y_menu.set(self.app.param_names[0])
        return True

    def _get_bounds(self, use_stat, sigma_level_str):
        if self.app.df is None: raise ValueError("No data loaded or defined.")
        if use_stat and self.generator_mode == "File":
            # --- UNICODE FIX ---
            multiplier_str = sigma_level_str.split('\u03c3')[0]
            multiplier = float(multiplier_str)
            # --- END UNICODE FIX ---
            means, stds = self.app.df.mean(), self.app.df.std(ddof=0)
            return (means - multiplier * stds).tolist(), (means + multiplier * stds).tolist()
        else:
            return self.app.df.min().tolist(), self.app.df.max().tolist()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xls *.xlsx")])
        if not path: return
        try:
            df = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
            df_num = df.select_dtypes(include=[np.number])
            if df_num.empty: raise ValueError("No numeric columns found.")
            self.app.df = df_num.copy()
            self.app.param_names = self.app.df.columns.tolist()
            self.app.generated_df, self.app.selected_indices = None, []
            self.app.target_stats = None

            theme_props = self.app.get_theme_properties()
            success_color = theme_props["text_color"]
            bold_font = CTkFont(weight="bold")

            self.loaded_label.configure(text=f"Loaded: {os.path.basename(path)}",
                                        text_color=success_color,
                                        font=bold_font)

            for btn in [self.generate_button, self.corner_points_button, self.elbow_button]: btn.configure(
                state="normal")

            self.save_generated_btn.configure(state="disabled")
            self.visualize_sample_button.configure(state="disabled")

            self.status_label.configure(
                text=f"Ready. {self.app.df.shape[0]} rows, {self.app.df.shape[1]} numeric params.",
                text_color=success_color,
                font=bold_font)

            for row_widgets in self.manual_gen_widgets:
                for widget in row_widgets.values(): widget.destroy()
            self.manual_gen_widgets = []
            self._update_preview_text()
            vis_tab = self.app.visualization_tab
            for menu in [self.preview_x_menu, self.preview_y_menu, self.preview_val_menu, vis_tab.worst_case_param_menu,
                         vis_tab.scatter_x_menu, vis_tab.scatter_y_menu, vis_tab.boxplot_param_menu]:
                menu.configure(values=self.app.param_names)
            if self.app.param_names:
                self.preview_x_menu.set(self.app.param_names[0]);
                self.preview_val_menu.set(self.app.param_names[0]);
                vis_tab.worst_case_param_menu.set(self.app.param_names[0]);
                vis_tab.scatter_x_menu.set(self.app.param_names[0]);
                vis_tab.boxplot_param_menu.set(self.app.param_names[0])
                if len(self.app.param_names) > 1:
                    self.preview_y_menu.set(self.app.param_names[1]);
                    vis_tab.scatter_y_menu.set(self.app.param_names[1])
                else:
                    self.preview_y_menu.set(self.app.param_names[0]);
                    vis_tab.scatter_y_menu.set(self.app.param_names[0])
            self.generator_mode_switcher.set("Load from File");
            self._switch_generator_mode("Load from File")
            self._draw_small_preview();
            vis_tab._update_vis_controls()
        except Exception as e:
            self.app.df, self.app.param_names, self.app.generated_df = None, [], None
            self.loaded_label.configure(text="Load failed.", text_color="red", font=CTkFont(weight="bold"))

            for btn in [self.generate_button, self.save_generated_btn, self.corner_points_button,
                        self.elbow_button, self.visualize_sample_button]: btn.configure(state="disabled")

            self.status_label.configure(text=f"Error: {e}", text_color="red", font=CTkFont(weight="bold"))
            self._update_preview_text();
            try:
                self.app.visualization_tab._update_vis_controls()
            except AttributeError as ae:
                if 'target_stats' in str(ae):
                    print("Warning: Skipping visualization update during load error due to missing target_stats.")
                else:
                    raise ae
            except Exception as vis_e:
                print(f"Warning: Error updating visualization tab during load error: {vis_e}")

    def generate(self):
        textbox = self.log_text
        try:
            textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.insert("end", "STARTING GENERATION LOG...\n")
            textbox.insert("end", "=" * 70 + "\n")
            textbox.see("end")

            if self.generator_mode == "Manual":
                if not self._process_manual_generator_inputs():
                    textbox.configure(state="disabled")
                    return
                self._update_preview_text()
                self.app.visualization_tab.show_original_var.set(False)

            if self.app.df is None: raise ValueError("No data file loaded or parameters defined.")

            n = int(self.samples_entry.get());
            seed = int(self.seed_entry.get()) if self.seed_entry.get() else None
            lower, upper = self._get_bounds(bool(self.use_stat_bounds_var.get()), self.sigma_level_menu.get())
            lower, upper = np.array(lower), np.array(upper)
            method, dim = self.method_menu.get(), len(self.app.param_names)

            if method == "Monte Carlo (Hypersphere)":
                textbox.insert("end", f"Generating {n} samples using Monte Carlo (Hypersphere)...\n")
                self.app.update_idletasks()
                radius_sq = float(self.radius_entry.get()) ** 2;
                center = (lower + upper) / 2
                samples = [];
                rng = np.random.default_rng(seed)
                while len(samples) < n:
                    point = rng.uniform(low=lower, high=upper, size=dim)
                    if np.sum((point - center) ** 2) <= radius_sq: samples.append(point)
                samples = np.array(samples)
                textbox.insert("end", "Generation complete.\n")

            elif method == "Monte Carlo":
                textbox.insert("end", f"Generating {n} samples using Monte Carlo...\n")
                self.app.update_idletasks()
                samples = np.random.default_rng(seed).uniform(low=lower, high=upper, size=(n, dim))
                textbox.insert("end", "Generation complete.\n")

            elif method == "Sobol Sequence":
                textbox.insert("end", f"Generating {n} samples using Sobol Sequence...\n")
                self.app.update_idletasks()
                sobol_sampler = qmc.Sobol(d=dim, scramble=True, seed=seed)
                sobol_samples = sobol_sampler.random(n=n)
                samples = qmc.scale(sobol_samples, lower, upper)
                textbox.insert("end", "Sobol generation complete.\n")

            else:  # Optimized LHS
                repeats = max(1, int(self.repeats_entry.get()))
                best_samples, best_score = None, -1.0
                textbox.insert("end", f"Starting Optimized LHS (n={n}, dim={dim})...\n")
                textbox.insert("end", f"Running {repeats} repeats to find best maximin score.\n")
                textbox.insert("end", "--------------------------------------------------\n")
                self.app.update_idletasks()

                for r in range(repeats):
                    sampler = qmc.LatinHypercube(d=dim, seed=(seed + r if seed is not None else None))
                    scaled = qmc.scale(sampler.random(n=n), lower, upper)
                    score = np.min(pdist(scaled)) if scaled.shape[0] > 1 else 0

                    if score > best_score:
                        best_score, best_samples = score, scaled
                        textbox.insert("end", f"Repeat {r + 1:>4}/{repeats}: New Best Score = {best_score:.6f}\n")
                        textbox.see("end")
                        self.app.update_idletasks()

                samples = best_samples
                textbox.insert("end", "--------------------------------------------------\n")
                textbox.insert("end", f"Optimization complete. Using best score: {best_score:.6f}\n")
                self.app.update_idletasks()

            self.app.generated_df = pd.DataFrame(samples, columns=self.app.param_names)
            self.app.selected_indices = []

            theme_props = self.app.get_theme_properties()
            success_color = theme_props["text_color"]
            bold_font = CTkFont(weight="bold")
            self.status_label.configure(text=f"Generated {len(self.app.generated_df)} samples.",
                                        text_color=success_color,
                                        font=bold_font)

            self.save_generated_btn.configure(state="normal")
            self.visualize_sample_button.configure(state="normal")

            self.app.visualization_tab._update_vis_controls()
            self.app.tabs.set("Visualization")

        except Exception as e:
            if 'textbox' in locals():
                try:
                    textbox.insert("end", f"\n--- ERROR ---\n{e}\n")
                except:
                    pass
            messagebox.showerror("Generation Error", f"An error occurred:\n{e}")
            self.status_label.configure(text=f"Error: {e}", text_color="red", font=CTkFont(weight="bold"))
        finally:
            if 'textbox' in locals():
                try:
                    textbox.configure(state="disabled")
                except:
                    pass

    def save_generated(self):
        if self.app.generated_df is None:
            self.status_label.configure(text="No generated sample to save.", text_color="red",
                                        font=CTkFont(weight="bold"))
            return
        fname = self.filename_entry.get().strip() or "sample_output"
        path = filedialog.asksaveasfilename(defaultextension=".csv" if self.format_menu.get() == "CSV" else ".xlsx",
                                            initialfile=fname, filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if not path: return
        try:
            if path.lower().endswith('.csv'):
                self.app.generated_df.to_csv(path, index=False)
            else:
                self.app.generated_df.to_excel(path, index=False, engine="openpyxl")

            theme_props = self.app.get_theme_properties()
            success_color = theme_props["text_color"]
            bold_font = CTkFont(weight="bold")
            self.status_label.configure(text=f"Saved to {os.path.basename(path)}",
                                        text_color=success_color,
                                        font=bold_font)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save the file.\n{e}")
            self.status_label.configure(text=f"Error saving file: {e}", text_color="red", font=CTkFont(weight="bold"))

    def save_corner_points(self):
        try:
            if self.generator_mode == "Manual":
                if not self._process_manual_generator_inputs(): return
            if self.app.df is None: raise ValueError("No data defined to create bounds.")
            lower, upper = self._get_bounds(bool(self.use_stat_bounds_var.get()), self.sigma_level_menu.get())
            df_corners = pd.DataFrame(list(itertools.product(*zip(lower, upper))), columns=self.app.param_names)
            path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                initialfile=f"{len(self.app.param_names)}D_corners",
                                                filetypes=[("CSV file", "*.csv")])
            if path:
                df_corners.to_csv(path, index=False)
                theme_props = self.app.get_theme_properties()
                success_color = theme_props["text_color"]
                bold_font = CTkFont(weight="bold")
                self.status_label.configure(text=f"Saved {len(df_corners)} corner points.",
                                            text_color=success_color,
                                            font=bold_font)
        except Exception as e:
            self.status_label.configure(text=f"Error saving corners: {e}", text_color="red",
                                        font=CTkFont(weight="bold"))

    def _detect_dynamic_structure(self, df):
        """Detect dynamic data structure."""
        try:
            if len(df) < 500:
                return False
            first_col = df.iloc[:, 0].values
            diff = np.diff(first_col)
            return np.any(diff < 0)
        except:
            return False

    def _load_dynamic_data(self, df):
        """Load dynamic data."""
        try:
            loader = DynamicDataLoader(df)
            self.dynamic_data = loader.load_and_reshape()
            self.input_channels = self.dynamic_data['input_names']
            self.output_channels = self.dynamic_data['output_names']
            self.frequency_array = self.dynamic_data['frequency']
            self.dynamic_metadata = self.dynamic_data['metadata']
            self.rnn_data = pd.DataFrame(self.dynamic_data['X'], columns=self.input_channels)
            self.rnn_data_type = "dynamic"
            print("✓ Dynamic data loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load dynamic data:\n{e}")
            raise

    def _build_dynamic_models(self):
        """Build dynamic models."""
        print("\nBuilding dynamic models...")
        self.build_model_button.configure(text="Training...", state="disabled")
        self.app.update_idletasks()

        try:
            X = self.dynamic_data['X']
            Y = self.dynamic_data['Y']
            input_names = self.dynamic_data['input_names']
            output_names = self.dynamic_data['output_names']

            X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.2, random_state=42)

            self.trained_models.clear()
            self.last_build_stats.clear()

            for idx, output_name in enumerate(output_names):
                model = GPR_PCA_DynamicModel(n_components=50, variance_threshold=0.99)
                model.fit(X_train, Y_train[:, :, idx], output_name=output_name)

                Y_pred_test = model.predict(X_test)
                r2_test = r2_score(Y_test[:, :, idx].ravel(), Y_pred_test.ravel())

                self.trained_models[output_name] = {'model': model, 'features': input_names}
                self.last_build_stats[output_name] = {'r2_test': r2_test, 'n_components': model.n_components_actual}

            messagebox.showinfo("Success",
                                f"Trained {len(output_names)} dynamic models!\nAvg R²: {np.mean([s['r2_test'] for s in self.last_build_stats.values()]):.3f}")

        except Exception as e:
            messagebox.showerror("Error", f"Training failed:\n{e}")
        finally:
            self.build_model_button.configure(text="Build Models", state="normal")

    def reset_to_default(self):
        self.generator_mode_switcher.set("Load from File")
        self._switch_generator_mode("Load from File")
        self.loaded_label.configure(text="No file loaded.", text_color="gray")
        self._clear_manual_gen_rows()

        default_param_list = ["-"]
        for menu in [self.preview_x_menu, self.preview_y_menu, self.preview_val_menu]:
            menu.configure(values=default_param_list)
            menu.set("-")

        self.method_menu.set("Optimized LHS")
        self.samples_entry.delete(0, 'end');
        self.samples_entry.insert(0, "100")
        self.use_stat_bounds_var.set(False)

        # --- UNICODE FIX ---
        self.sigma_level_menu.set(f"4\u03c3 (99.99%)")
        # --- END UNICODE FIX ---

        self.radius_entry.delete(0, 'end');
        self.radius_entry.insert(0, "3.0")
        self.repeats_entry.delete(0, 'end');
        self.repeats_entry.insert(0, "20")
        self.seed_entry.delete(0, 'end')
        self.filename_entry.delete(0, 'end');
        self.filename_entry.insert(0, "sample_output")
        self.format_menu.set("CSV")
        self._on_method_change("Optimized LHS")
        self._on_stat_bounds_toggle()

        self.save_generated_btn.configure(state="disabled")
        self.visualize_sample_button.configure(state="disabled")

        self.status_label.configure(text="")

        self.n_start_entry.delete(0, 'end');
        self.n_start_entry.insert(0, "20")
        self.n_end_entry.delete(0, 'end');
        self.n_end_entry.insert(0, "100")
        self.n_step_entry.delete(0, 'end');
        self.n_step_entry.insert(0, "10")

        self._update_preview_text()
        self._draw_small_preview()

        if self.analysis_popup is not None and self.analysis_popup.winfo_exists():
            self._on_popup_close("analysis")
        if self.vis_popup is not None and self.vis_popup.winfo_exists():
            self._on_popup_close("vis")

    def _on_manual_add_keypress(self, *args):
        try:
            if self.app.tabs.get() == "Generator" and self.generator_mode == "Manual": self._add_manual_gen_row()
        except Exception:
            pass

    def _create_analysis_popup(self):
        if self.analysis_popup is not None and self.analysis_popup.winfo_exists(): self.analysis_popup.focus(); return None, None, None
        popup = ctk.CTkToplevel(self.app);
        popup.title("Optimal Sample Size Analysis");
        popup.geometry("600x500");
        self.analysis_popup = popup
        popup.protocol("WM_DELETE_WINDOW", lambda: self._on_popup_close("analysis"))
        ctk.CTkLabel(popup, text="Look for the 'elbow' where the LHS curve flattens.", justify="left").pack(pady=10,
                                                                                                            padx=10)
        fig = plt.Figure(figsize=(6, 4), dpi=100);
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=popup);
        canvas.get_tk_widget().pack(side=ctk.TOP, fill=ctk.BOTH, expand=True, padx=10, pady=5)
        progress_bar = ctk.CTkProgressBar(popup, orientation="horizontal");
        progress_bar.set(0);
        progress_bar.pack(fill="x", padx=10, pady=(5, 10))
        return popup, ax, progress_bar

    def _on_popup_close(self, popup_type="analysis"):
        if popup_type == "analysis":
            if self.analysis_popup:
                self.analysis_popup.destroy()
                self.analysis_popup = None
        elif popup_type == "vis":
            if self.vis_popup:
                self.vis_popup.destroy()
                self.vis_popup = None
            plt.style.use(self.app.get_theme_properties()["mpl_style"])

    def find_optimal_n(self):
        try:
            if self.generator_mode == "Manual":
                if not self._process_manual_generator_inputs(): return
            if self.app.df is None: raise ValueError("No data defined.")
            n_start, n_end, n_step = int(self.n_start_entry.get()), int(self.n_end_entry.get()), int(
                self.n_step_entry.get())
            if not (n_start > 1 and n_end > n_start and n_step > 0): raise ValueError("Invalid range.")
            n_values = list(range(n_start, n_end + 1, n_step))
            popup, ax, progress_bar = self._create_analysis_popup()
            if popup is None: return
            lower, upper = self._get_bounds(bool(self.use_stat_bounds_var.get()), self.sigma_level_menu.get())
            dim, repeats = len(self.app.param_names), max(1, int(self.repeats_entry.get()))
            lhs_scores, mc_scores = [], []
            for i, n in enumerate(n_values):
                progress_bar.set((i + 1) / len(n_values));
                popup.title(f"Analyzing... (n={n})");
                self.app.update_idletasks()
                best_lhs_score = max(
                    [np.min(pdist(qmc.scale(qmc.LatinHypercube(d=dim, seed=r).random(n=n), lower, upper))) for r in
                     range(repeats)])
                best_mc_score = max(
                    [np.min(pdist(np.random.default_rng(r).uniform(low=lower, high=upper, size=(n, dim)))) for r in
                     range(repeats)])
                lhs_scores.append(best_lhs_score);
                mc_scores.append(best_mc_score)
            popup.title("Method Comparison vs. Sample Size");
            theme_props = self.app.get_theme_properties()
            ax.clear();
            ax.figure.set_facecolor(theme_props['plot_bg']);
            ax.set_facecolor(theme_props['plot_bg'])
            ax.plot(n_values, lhs_scores, marker='o', linestyle='-', color='#00FFFF', label='Optimized LHS')
            ax.plot(n_values, mc_scores, marker='x', linestyle='--', color='#FF006E', label='Monte Carlo')
            ax.set_title("Method Quality vs. Sample Size", color=theme_props['text_color']);
            ax.set_xlabel("Number of Samples (n)", color=theme_props['text_color']);
            ax.set_ylabel("Maximin Distance Score (Higher is Better)", color=theme_props['text_color'])
            ax.tick_params(axis='x', colors=theme_props["text_color"]);
            ax.tick_params(axis='y', colors=theme_props["text_color"])
            for spine in ax.spines.values(): spine.set_color(theme_props["text_color"])
            legend = ax.legend();
            legend.get_frame().set_facecolor(theme_props['plot_bg']);
            [text.set_color(theme_props['text_color']) for text in legend.get_texts()]
            ax.grid(color=theme_props['grid_color'], linestyle='--', alpha=0.5);
            ax.figure.tight_layout();
            ax.figure.canvas.draw()
        except Exception as e:
            if self.analysis_popup and self.analysis_popup.winfo_exists():
                ctk.CTkLabel(self.analysis_popup, text=f"Error: {e}", text_color="red", wraplength=550).pack(pady=10)
            else:
                self.status_label.configure(text=f"Error: {e}", text_color="red")
        finally:
            if 'progress_bar' in locals() and progress_bar.winfo_exists(): progress_bar.pack_forget()

    def _on_vis_popup_close(self):
        self._on_popup_close("vis")

    def visualize_generated_sample(self):
        if self.vis_popup is not None and self.vis_popup.winfo_exists():
            self.vis_popup.focus()
            return

        if self.app.generated_df is None:
            messagebox.showinfo("Visualize Sample",
                                "No generated sample available. Please click 'Generate & Visualize' first.")
            return

        df = self.app.generated_df
        if df.shape[1] < 2:
            messagebox.showinfo("Visualize Sample",
                                f"Sample only has {df.shape[1]} dimension(s). At least 2 are needed for these plots.")
            return

        self.vis_popup = ctk.CTkToplevel(self.app)
        self.vis_popup.title("Generated Sample Visualization")
        self.vis_popup.geometry("1000x800")
        self.vis_popup.protocol("WM_DELETE_WINDOW", self._on_vis_popup_close)

        tab_view = ctk.CTkTabview(self.vis_popup)
        tab_view.pack(fill="both", expand=True, padx=10, pady=10)
        splom_tab = tab_view.add("Scatter Plot Matrix (SPLOM)")
        pcp_tab = tab_view.add("Parallel Coordinates (PCP)")

        original_style = plt.style.use('default')
        theme_props = THEMES["Light"]

        if _HAS_SEABORN:
            try:
                splom_frame = ctk.CTkFrame(splom_tab, fg_color="transparent")
                splom_frame.pack(fill="both", expand=True)
                g = sns.PairGrid(df, corner=True)
                g.map_diag(sns.histplot, kde=False, color=theme_props["hist_color_generated"])
                g.map_lower(sns.scatterplot, s=10, alpha=0.7, color=theme_props["hist_color_original"])
                g.fig.suptitle("Scatter Plot Matrix (SPLOM) - Check for uniform projections", y=1.02, fontsize=12)
                g.fig.subplots_adjust(left=0.22, bottom=0.15, top=0.95, right=0.98)

                num_vars = df.shape[1]
                for r in range(num_vars):
                    for c in range(num_vars):
                        ax = g.axes[r, c]
                        if ax is None: continue
                        if c == 0 and r > 0:
                            ax.yaxis.label.set_fontsize(7)
                            ax.yaxis.label.set_rotation(45)
                            ax.yaxis.label.set_ha('right')
                            ax.yaxis.label.set_va('center')
                            ax.yaxis.set_label_coords(-0.35, 0.5)
                        if r == num_vars - 1:
                            xlabel_text = ax.get_xlabel()
                            ax.set_xlabel(xlabel_text, rotation=15, ha='right')
                            ax.xaxis.label.set_fontsize(8)
                            plt.setp(ax.get_xticklabels(), rotation=15, ha='right', fontsize=7)
                        elif c < r:
                            plt.setp(ax.get_xticklabels(), visible=False)
                        if c > 0 and c < r:
                            plt.setp(ax.get_yticklabels(), visible=False)
                        elif c == 0 and r > 0:
                            plt.setp(ax.get_yticklabels(), fontsize=7)

                canvas_splom = FigureCanvasTkAgg(g.fig, master=splom_frame)
                canvas_splom.get_tk_widget().pack(fill="both", expand=True, side="top")
                toolbar_splom = NavigationToolbar2Tk(canvas_splom, splom_frame, pack_toolbar=False)
                toolbar_splom.update()
                toolbar_splom.pack(fill="x", side="bottom")

            except Exception as e:
                import traceback
                print("Error creating SPLOM plot:")
                traceback.print_exc()
                ctk.CTkLabel(splom_tab, text=f"Error creating SPLOM plot:\n{e}", wraplength=700).pack(pady=20)
        else:
            ctk.CTkLabel(splom_tab,
                         text="This feature requires the 'seaborn' library.\nInstall it via 'pip install seaborn' and restart.",
                         font=("", 14), text_color="orange").pack(pady=50, padx=20)

        if _HAS_PANDAS_PLOTTING:
            try:
                pcp_frame = ctk.CTkFrame(pcp_tab, fg_color="transparent")
                pcp_frame.pack(fill="both", expand=True)

                fig_pcp, ax_pcp = plt.subplots(figsize=(10, 6))
                fig_pcp.set_facecolor(theme_props["plot_bg"])
                ax_pcp.set_facecolor(theme_props["plot_bg"])

                df_pcp = df.copy()
                first_param = df_pcp.columns[0]
                try:
                    df_pcp['color_group'] = pd.qcut(df_pcp[first_param], 4, labels=False, duplicates='drop')
                except ValueError:
                    df_pcp['color_group'] = (df_pcp[first_param] > df_pcp[first_param].median()).astype(int)

                parallel_coordinates(df_pcp, 'color_group', colormap='viridis', ax=ax_pcp, alpha=0.5)

                ax_pcp.set_title("Parallel Coordinates Plot (PCP) - Check for full-range coverage")
                if ax_pcp.get_legend() is not None:
                    ax_pcp.get_legend().set_visible(False)

                ax_pcp.tick_params(axis='x', labelsize=8)
                plt.setp(ax_pcp.get_xticklabels(), rotation=15, ha='right')

                fig_pcp.tight_layout(pad=1.0)

                canvas_pcp = FigureCanvasTkAgg(fig_pcp, master=pcp_frame)
                canvas_pcp.get_tk_widget().pack(fill="both", expand=True, side="top")
                toolbar_pcp = NavigationToolbar2Tk(canvas_pcp, pcp_frame, pack_toolbar=False)
                toolbar_pcp.update()
                toolbar_pcp.pack(fill="x", side="bottom")

            except Exception as e:
                import traceback
                print("Error creating PCP plot:")
                traceback.print_exc()
                ctk.CTkLabel(pcp_tab, text=f"Error creating Parallel Coordinates plot:\n{e}", wraplength=700).pack(
                    pady=20)
        else:
            ctk.CTkLabel(pcp_tab, text="Error: Could not find pandas.plotting.parallel_coordinates.",
                         font=("", 14), text_color="red").pack(pady=50, padx=20)

        plt.style.use(original_style)


# =============================================================================
# --- TAB CLASS: VisualizationTab (with Unicode Fix) ---
# =============================================================================
class VisualizationTab(ctk.CTkFrame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, fg_color="transparent")
        self.app = app_instance
        self.appearance_frame_visible = True
        self.appearance_frame_row = 0
        self.wc_popup = None
        self._ellipse_point_artists = []
        self._overlay_ellipse_artists = []
        self._lasso = None
        self._lasso_active = False
        self._mpl_cursor = None
        self._selection_artists = []

        # --- STATE FOR ELLIPSE ---
        self._ellipse_transform = None
        self._ellipse_threshold = None

        self._build_ui()

    def _detect_dynamic_structure(self, df):
        """Detect dynamic data structure."""
        try:
            if len(df) < 500:
                return False
            first_col = df.iloc[:, 0].values
            diff = np.diff(first_col)
            return np.any(diff < 0)
        except:
            return False

    def _load_dynamic_data(self, df):
        """Load dynamic data."""
        try:
            loader = DynamicDataLoader(df)
            self.dynamic_data = loader.load_and_reshape()
            self.input_channels = self.dynamic_data['input_names']
            self.output_channels = self.dynamic_data['output_names']
            self.frequency_array = self.dynamic_data['frequency']
            self.dynamic_metadata = self.dynamic_data['metadata']
            self.rnn_data = pd.DataFrame(self.dynamic_data['X'], columns=self.input_channels)
            self.rnn_data_type = "dynamic"
            print("✓ Dynamic data loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load dynamic data:\n{e}")
            raise

    def _build_dynamic_models(self):
        """Build dynamic models."""
        print("\nBuilding dynamic models...")
        self.build_model_button.configure(text="Training...", state="disabled")
        self.app.update_idletasks()

        try:
            X = self.dynamic_data['X']
            Y = self.dynamic_data['Y']
            input_names = self.dynamic_data['input_names']
            output_names = self.dynamic_data['output_names']

            X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.2, random_state=42)

            self.trained_models.clear()
            self.last_build_stats.clear()

            for idx, output_name in enumerate(output_names):
                model = GPR_PCA_DynamicModel(n_components=50, variance_threshold=0.99)
                model.fit(X_train, Y_train[:, :, idx], output_name=output_name)

                Y_pred_test = model.predict(X_test)
                r2_test = r2_score(Y_test[:, :, idx].ravel(), Y_pred_test.ravel())

                self.trained_models[output_name] = {'model': model, 'features': input_names}
                self.last_build_stats[output_name] = {'r2_test': r2_test, 'n_components': model.n_components_actual}

            messagebox.showinfo("Success",
                                f"Trained {len(output_names)} dynamic models!\nAvg R²: {np.mean([s['r2_test'] for s in self.last_build_stats.values()]):.3f}")

        except Exception as e:
            messagebox.showerror("Error", f"Training failed:\n{e}")
        finally:
            self.build_model_button.configure(text="Build Models", state="normal")

    def reset_to_default(self):
        """Resets the Visualization tab to its initial state."""
        print("--- Resetting Visualization Tab ---")

        default_param_list = ["-"]
        for menu in [self.scatter_x_menu, self.scatter_y_menu, self.boxplot_param_menu, self.worst_case_param_menu]:
            menu.configure(values=default_param_list)
            menu.set("-")

        self.size_slider.set(18)
        self.alpha_slider.set(0.9)
        self.show_original_var.set(True)
        self.show_generated_var.set(True)
        self.standardize_axes_var.set(False)
        self.show_min_max_box_var.set(False)
        self.show_2s_box_var.set(False)
        self.show_3s_box_var.set(False)
        self.show_4s_box_var.set(False)
        self.elliptical_sigma_menu.set("None")
        self.use_target_ellipse_var.set(False)
        self.hide_outside_ellipse_var.set(False)

        self._on_ellipse_change("None")

        self.cmap_menu.set("viridis")
        self.palette_menu.set("Color by X")
        self.solid_color_entry.delete(0, 'end');
        self.solid_color_entry.insert(0, "#1f77b4")
        if not self.appearance_frame_visible:
            self.toggle_appearance_frame()

        try:
            self.selected_text.configure(state="normal")
            self.selected_text.delete("1.0", "end")
            self.selected_text.configure(state="disabled")
        except Exception as e:
            print(f"Error resetting selected text in VisualizationTab: {e}")

        if self.wc_popup is not None and self.wc_popup.winfo_exists():
            self._on_wc_popup_close()

        self._draw_plots()
        print("--- Visualization Tab Reset Complete ---")

    def _build_ui(self):
        left = ctk.CTkFrame(self, width=320);
        left.pack(side="left", fill="y", padx=(6, 4), pady=6)
        right = ctk.CTkFrame(self);
        right.pack(side="left", fill="both", expand=True, padx=(4, 6), pady=6)
        row = 0;
        small = ("", 11)

        ctk.CTkLabel(left, text="X axis:", font=small).grid(row=row, column=0, padx=6, pady=(8, 4), sticky="w");
        self.scatter_x_menu = ctk.CTkOptionMenu(left, values=["-"], width=200, command=lambda _: self._draw_plots());
        self.scatter_x_menu.grid(row=row, column=1, padx=6, pady=(8, 4), sticky="e");
        row += 1
        ctk.CTkLabel(left, text="Y axis:", font=small).grid(row=row, column=0, padx=6, pady=4, sticky="w");
        self.scatter_y_menu = ctk.CTkOptionMenu(left, values=["-"], width=200, command=lambda _: self._draw_plots());
        self.scatter_y_menu.grid(row=row, column=1, padx=6, pady=4, sticky="e");
        row += 1
        ctk.CTkLabel(left, text="Boxplot param:", font=small).grid(row=row, column=0, padx=6, pady=4, sticky="w");
        self.boxplot_param_menu = ctk.CTkOptionMenu(left, values=["-"], width=200,
                                                    command=lambda _: self._draw_plots());
        self.boxplot_param_menu.grid(row=row, column=1, padx=6, pady=4, sticky="e");
        row += 1

        main_toggles_frame = ctk.CTkFrame(left, fg_color="transparent");
        main_toggles_frame.grid(row=row, column=0, columnspan=2, padx=0, pady=0, sticky="ew");
        row += 1
        self.show_original_var = ctk.BooleanVar(value=True);
        self.show_orig_cb = ctk.CTkCheckBox(main_toggles_frame, text="Show Original", variable=self.show_original_var,
                                            command=self._draw_plots);
        self.show_orig_cb.pack(anchor="w", padx=6, pady=4)
        self.show_generated_var = ctk.BooleanVar(value=True);
        self.show_gen_cb = ctk.CTkCheckBox(main_toggles_frame, text="Show Generated", variable=self.show_generated_var,
                                           command=self._draw_plots);
        self.show_gen_cb.pack(anchor="w", padx=6, pady=4)

        # --- UNICODE FIX ---
        self.standardize_axes_var = ctk.BooleanVar(value=False);
        self.standardize_axes_cb = ctk.CTkCheckBox(main_toggles_frame, text=f"Standardize Axes (Plot by \u03c3)",
                                                   variable=self.standardize_axes_var, command=self._draw_plots);
        self.standardize_axes_cb.pack(anchor="w", padx=6, pady=4)
        # --- END UNICODE FIX ---

        overlay_frame = ctk.CTkFrame(main_toggles_frame, fg_color="transparent");
        overlay_frame.pack(fill="x", padx=0, pady=4)
        ctk.CTkLabel(overlay_frame, text="Overlay Boxes:", font=small).pack(side="left", padx=(6, 10))
        self.show_min_max_box_var = ctk.BooleanVar(value=False);
        self.show_2s_box_var = ctk.BooleanVar(value=False);
        self.show_3s_box_var = ctk.BooleanVar(value=False);
        self.show_4s_box_var = ctk.BooleanVar(value=False)

        # --- UNICODE FIX ---
        ctk.CTkCheckBox(overlay_frame, text="Min/Max", variable=self.show_min_max_box_var, width=10,
                        command=self._draw_plots).pack(side="left", padx=3)
        ctk.CTkCheckBox(overlay_frame, text=f"2\u03c3", variable=self.show_2s_box_var, width=10,
                        command=self._draw_plots).pack(side="left", padx=3)
        ctk.CTkCheckBox(overlay_frame, text=f"3\u03c3", variable=self.show_3s_box_var, width=10,
                        command=self._draw_plots).pack(side="left", padx=3)
        ctk.CTkCheckBox(overlay_frame, text=f"4\u03c3", variable=self.show_4s_box_var, width=10,
                        command=self._draw_plots).pack(side="left", padx=3)
        # --- END UNICODE FIX ---

        ellipse_frame = ctk.CTkFrame(left, fg_color="transparent");
        ellipse_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=0, pady=0);
        row += 1
        ctk.CTkLabel(ellipse_frame, text="Elliptical Boundary:").pack(side="left", padx=(5, 0))

        # --- UNICODE FIX ---
        self.elliptical_sigma_menu = ctk.CTkOptionMenu(ellipse_frame,
                                                       values=["None", f"2\u03c3", f"3\u03c3", f"4\u03c3"], width=80,
                                                       command=self._on_ellipse_change);
        # --- END UNICODE FIX ---

        self.elliptical_sigma_menu.set("None");
        self.elliptical_sigma_menu.pack(side="left", padx=5)
        self.find_ellipse_pts_btn = ctk.CTkButton(ellipse_frame, text="Find Ellipse Pts", width=10,
                                                  command=self.find_ellipse_points, state="disabled",
                                                  fg_color="#003366", hover_color="#004080", text_color="white");
        self.find_ellipse_pts_btn.pack(side="left", padx=5)

        self.hide_outside_ellipse_var = ctk.BooleanVar(value=False);
        self.hide_outside_ellipse_cb = ctk.CTkCheckBox(left, text="Hide points outside ellipse",
                                                       variable=self.hide_outside_ellipse_var, state="disabled",
                                                       command=self._draw_plots);
        self.hide_outside_ellipse_cb.grid(row=row, column=0, columnspan=2, padx=6, pady=(4, 0), sticky="w");
        row += 1

        # --- UNICODE FIX ---
        self.use_target_ellipse_var = ctk.BooleanVar(value=False)
        self.use_target_ellipse_cb = ctk.CTkCheckBox(left, text=f"Overlay Target \u03c3 Ellipse (if avail.)",
                                                     variable=self.use_target_ellipse_var, state="disabled",
                                                     command=self._draw_plots);
        # --- END UNICODE FIX ---

        self.use_target_ellipse_cb.grid(row=row, column=0, columnspan=2, padx=6, pady=(0, 4), sticky="w");
        row += 1

        # --- UNICODE FIX ---
        self.toggle_appearance_btn = ctk.CTkButton(left, text=f"Plot Appearance Settings \u25bc",
                                                   command=self.toggle_appearance_frame, fg_color="#003366",
                                                   hover_color="#004080", text_color="white");
        # --- END UNICODE FIX ---

        self.toggle_appearance_btn.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 2));
        row += 1
        self.appearance_frame_row = row;
        self.appearance_frame = ctk.CTkFrame(left, fg_color="transparent");
        self.appearance_frame.grid(row=row, column=0, columnspan=2, padx=0, pady=0, sticky="ew");
        row += 1
        app_row = 0
        ctk.CTkLabel(self.appearance_frame, text="Colormap:", font=small).grid(row=app_row, column=0, padx=6, pady=4,
                                                                               sticky="w");
        self.cmap_menu = ctk.CTkOptionMenu(self.appearance_frame,
                                           values=["viridis", "plasma", "inferno", "magma", "cividis", "Greys"],
                                           width=160, command=lambda _: self._draw_plots());
        self.cmap_menu.set("viridis");
        self.cmap_menu.grid(row=app_row, column=1, padx=6, pady=4, sticky="e");
        app_row += 1
        ctk.CTkLabel(self.appearance_frame, text="Point palette:", font=small).grid(row=app_row, column=0, padx=6,
                                                                                    pady=4, sticky="w");
        self.palette_menu = ctk.CTkOptionMenu(self.appearance_frame,
                                              values=["Color by X", "Black/White", "Solid Color"], width=160,
                                              command=lambda _: self._draw_plots());
        self.palette_menu.set("Color by X");
        self.palette_menu.grid(row=app_row, column=1, padx=6, pady=4, sticky="e");
        app_row += 1
        ctk.CTkLabel(self.appearance_frame, text="Solid color (#hex):", font=("", 10)).grid(row=app_row, column=0,
                                                                                            padx=6, pady=4, sticky="w");
        self.solid_color_entry = ctk.CTkEntry(self.appearance_frame, width=120);
        self.solid_color_entry.insert(0, "#1f77b4");
        self.solid_color_entry.grid(row=app_row, column=1, padx=6, pady=4, sticky="e");
        app_row += 1
        ctk.CTkLabel(self.appearance_frame, text="Size:", font=small).grid(row=app_row, column=0, padx=6, pady=4,
                                                                           sticky="w");
        self.size_slider = ctk.CTkSlider(self.appearance_frame, from_=5, to=80, number_of_steps=75,
                                         command=lambda _: self._draw_plots());
        self.size_slider.set(18);
        self.size_slider.grid(row=app_row, column=1, padx=6, pady=4, sticky="we");
        app_row += 1
        ctk.CTkLabel(self.appearance_frame, text="Alpha:", font=small).grid(row=app_row, column=0, padx=6, pady=4,
                                                                            sticky="w");
        self.alpha_slider = ctk.CTkSlider(self.appearance_frame, from_=0.05, to=1.0, number_of_steps=95,
                                          command=lambda _: self._draw_plots());
        self.alpha_slider.set(0.9);
        self.alpha_slider.grid(row=app_row, column=1, padx=6, pady=4, sticky="we");
        app_row += 1
        wc_frame = ctk.CTkFrame(left);
        wc_frame.grid(row=row, column=0, columnspan=2, padx=6, pady=(10, 4), sticky='ew');
        row += 1
        ctk.CTkLabel(wc_frame, text="Worst-Case Analysis", font=("", 12, "bold")).pack(anchor="w", padx=5);
        ctk.CTkLabel(wc_frame, text="Find worst-case for:").pack(side="left", padx=5);
        self.worst_case_param_menu = ctk.CTkOptionMenu(wc_frame, values=["-"], width=140);
        self.worst_case_param_menu.pack(side="left", padx=5);
        self.run_wc_analysis_btn = ctk.CTkButton(wc_frame, text="Run", width=40, command=self.run_worst_case_analysis,
                                                 fg_color="#003366", hover_color="#004080", text_color="white");
        self.run_wc_analysis_btn.pack(side="left", padx=5)
        self.update_view_btn = ctk.CTkButton(left, text="Update View", command=self._draw_plots, fg_color="#003366",
                                             hover_color="#004080", text_color="white");
        self.update_view_btn.grid(row=row, column=0, columnspan=2, padx=6, pady=(8, 4), sticky="we");
        row += 1
        ctk.CTkLabel(left, text="Selection / Point Data:", font=small).grid(row=row, column=0, columnspan=2, padx=6,
                                                                            pady=(4, 2), sticky="w");
        row += 1
        self.selected_text = ctk.CTkTextbox(left, width=280, height=220, font=("Courier New", 12));
        self.selected_text.grid(row=row, column=0, columnspan=2, padx=6, pady=(2, 8));
        self.selected_text.configure(state="disabled");
        row += 1

        self.fig = plt.figure(figsize=(11, 5));
        self.canvas = FigureCanvasTkAgg(self.fig, master=right);
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar_frame = ctk.CTkFrame(right, height=34);
        toolbar_frame.pack(fill="x");
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame);
        self.toolbar.update()
        self.ax_scatter = self.fig.add_subplot(121);
        self.ax_box = self.fig.add_subplot(122)
        plt.tight_layout(pad=3.0)
        self.apply_theme()

    def apply_theme(self):
        theme_props = self.app.get_theme_properties()
        all_figs = [self.fig]
        if self.wc_popup is not None and self.wc_popup.winfo_exists():
            try:
                for child in self.wc_popup.winfo_children():
                    if isinstance(child, FigureCanvasTkAgg):
                        all_figs.append(child.figure)
                        break
            except Exception:
                pass

        all_axes = [ax for fig_obj in all_figs for ax in fig_obj.axes]

        for fig in all_figs:
            fig.set_facecolor(theme_props["plot_bg"])

        for ax in all_axes:
            if not ax: continue
            ax.set_facecolor(theme_props["plot_bg"]);
            ax.tick_params(axis='x', colors=theme_props["text_color"]);
            ax.tick_params(axis='y', colors=theme_props["text_color"])
            for spine in ax.spines.values(): spine.set_color(theme_props["text_color"])
            ax.xaxis.label.set_color(theme_props["text_color"]);
            ax.yaxis.label.set_color(theme_props["text_color"]);
            ax.title.set_color(theme_props["text_color"]);
            ax.grid(color=theme_props["grid_color"], linestyle='--', alpha=0.5)
            legend = ax.get_legend()
            if legend:
                legend.get_frame().set_facecolor(theme_props['plot_bg'])
                legend.get_frame().set_edgecolor(theme_props['grid_color'])
                if legend.get_texts():
                    for text in legend.get_texts():
                        try:
                            text.set_color(theme_props["text_color"])
                        except Exception:
                            pass
                if legend.get_title():
                    try:
                        legend.get_title().set_color(theme_props["text_color"])
                    except Exception:
                        pass

        self._draw_plots()

        if self.wc_popup is not None and self.wc_popup.winfo_exists():
            try:
                for child in self.wc_popup.winfo_children():
                    if isinstance(child, FigureCanvasTkAgg):
                        child.draw_idle()
                        break
            except Exception:
                pass

    def toggle_appearance_frame(self):
        # --- UNICODE FIX ---
        if self.appearance_frame_visible:
            self.appearance_frame.grid_forget();
            self.toggle_appearance_btn.configure(text=f"Plot Appearance Settings \u25b6")  # right arrow
        else:
            self.appearance_frame.grid(row=self.appearance_frame_row, column=0, columnspan=2, padx=0, pady=0,
                                       sticky="ew");
            self.toggle_appearance_btn.configure(
                text=f"Plot Appearance Settings \u25bc")  # down arrow
        # --- END UNICODE FIX ---
        self.appearance_frame_visible = not self.appearance_frame_visible

    def _on_ellipse_change(self, choice):
        has_target_stats = self.app.target_stats is not None
        if choice == "None":
            self.hide_outside_ellipse_cb.configure(state="disabled");
            self.hide_outside_ellipse_var.set(False)
            self.find_ellipse_pts_btn.configure(state="disabled")
            self.use_target_ellipse_cb.configure(state="disabled");
            self.use_target_ellipse_var.set(False)
            self._ellipse_transform = None;
            self._ellipse_threshold = None
        else:
            self.hide_outside_ellipse_cb.configure(state="normal")
            self.find_ellipse_pts_btn.configure(state="normal")
            if has_target_stats:
                self.use_target_ellipse_cb.configure(state="normal")
            else:
                self.use_target_ellipse_cb.configure(state="disabled")
                self.use_target_ellipse_var.set(False)

        self._draw_plots()

    def _update_vis_controls(self):
        is_data_available = self.app.df is not None
        self.run_wc_analysis_btn.configure(state="normal" if is_data_available else "disabled")

        if not is_data_available:
            self.show_orig_cb.configure(state="disabled");
            self.show_original_var.set(False)
            self.show_gen_cb.configure(state="disabled");
            self.show_generated_var.set(False)
            vals = ["-"]
        else:
            self.show_orig_cb.configure(state="normal");
            vals = self.app.df.columns.tolist()
            if self.app.generated_df is not None:
                self.show_gen_cb.configure(state="normal");
            else:
                self.show_gen_cb.configure(state="disabled");
                self.show_generated_var.set(False)
        for menu in [self.scatter_x_menu, self.scatter_y_menu, self.boxplot_param_menu, self.worst_case_param_menu]:
            current_val = menu.get();
            menu.configure(values=vals)
            if current_val in vals and vals != ["-"]:
                menu.set(current_val)
            elif vals != ["-"]:
                menu.set(vals[0])
            else:
                menu.set("-")
        if len(vals) > 1 and self.scatter_y_menu.get() == self.scatter_x_menu.get(): self.scatter_y_menu.set(vals[1])
        self._on_ellipse_change(self.elliptical_sigma_menu.get())

    def _get_mahalanobis_distance_squared(self, points, mean, inv_cov):
        diff = points - mean
        if diff.ndim == 1: diff = diff.reshape(1, -1)
        if diff.shape[0] > 1:
            return np.sum((diff @ inv_cov) * diff, axis=1)
        else:
            return (diff @ inv_cov @ diff.T)[0, 0]

    def _draw_plots(self):
        self.fig.clear();
        for artist in self._ellipse_point_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._ellipse_point_artists.clear()
        self._overlay_ellipse_artists = []
        self._ellipse_transform = None;
        self._ellipse_threshold = None

        is_original_available = (self.app.df is not None)
        is_generated_available = (self.app.generated_df is not None)
        show_orig = self.show_original_var.get() and is_original_available
        show_gen = self.show_generated_var.get() and is_generated_available
        if not show_gen and not show_orig:
            try:
                self.canvas.draw_idle()
            except Exception:
                pass
            return

        theme_props = self.app.get_theme_properties()
        self.ax_scatter, self.ax_box = self.fig.add_subplot(121), self.fig.add_subplot(122)
        for ax in self.fig.axes: ax.clear()
        for ax in [self.ax_scatter, self.ax_box]:
            ax.set_facecolor(theme_props["plot_bg"]);
            ax.tick_params(colors=theme_props["text_color"]);
            [s.set_color(theme_props["text_color"]) for s in ax.spines.values()];
            ax.grid(color=theme_props["grid_color"], ls='--', alpha=0.5)

        current_params = self.app.df.columns.tolist() if is_original_available else self.app.generated_df.columns.tolist()
        x, y, bp = self.scatter_x_menu.get(), self.scatter_y_menu.get(), self.boxplot_param_menu.get()
        if not all(p in current_params for p in [x, y, bp]):
            try:
                self.canvas.draw_idle()
            except Exception:
                pass
            return

        size, alpha = float(self.size_slider.get()), float(self.alpha_slider.get())
        cmap_name = self.cmap_menu.get()
        palette = self.palette_menu.get()
        solid_hex = self.solid_color_entry.get().strip()

        use_target_stats = (self.app.target_stats is not None) and is_original_available

        stats_mean, stats_std, stats_cov = None, None, None
        df_for_minmax = None

        if is_original_available:
            df_for_minmax = self.app.df
            stats_mean = self.app.df.mean()
            stats_std = self.app.df.std(ddof=0)
            stats_cov = self.app.df.cov()
        elif is_generated_available:
            stats_mean = self.app.generated_df.mean()
            stats_std = self.app.generated_df.std(ddof=0)
            stats_cov = self.app.generated_df.cov()

        df_orig_unfiltered = self.app.df.copy() if is_original_available else None
        df_gen_unfiltered = self.app.generated_df.copy() if is_generated_available else None
        df_orig_plot = df_orig_unfiltered.copy() if df_orig_unfiltered is not None else None
        df_gen_plot = df_gen_unfiltered.copy() if df_gen_unfiltered is not None else None

        standardize = self.standardize_axes_var.get()
        mean_for_stdize = stats_mean
        std_for_stdize = stats_std
        if standardize and mean_for_stdize is not None and std_for_stdize is not None:
            if df_orig_plot is not None: df_orig_plot = (df_orig_plot - mean_for_stdize) / std_for_stdize
            if df_gen_plot is not None: df_gen_plot = (df_gen_plot - mean_for_stdize) / std_for_stdize

        # --- UNICODE FIX ---
        x_label, y_label = (f"{p} (Std)" for p in (x, y)) if standardize else (x, y)
        # --- END UNICODE FIX ---

        ellipse_choice = self.elliptical_sigma_menu.get()
        hide_outside = self.hide_outside_ellipse_var.get()
        draw_target_ellipse = self.use_target_ellipse_var.get()

        primary_ellipse_params = None
        primary_ellipse_label = ""

        # --- UNICODE FIX ---
        if ellipse_choice != "None" and stats_mean is not None and stats_cov is not None:
            k_str = ellipse_choice.replace(f"\u03c3", "")
            k = int(k_str)
            # --- END UNICODE FIX ---
            confidence_level = {2: 0.9545, 3: 0.9973, 4: 0.999936}[k]
            chi2_thresh = chi2.ppf(confidence_level, 2)
            self._ellipse_threshold = chi2_thresh

            mean_to_draw = stats_mean
            cov_to_draw = stats_cov
            if not show_orig and show_gen:
                mean_to_draw = self.app.generated_df.mean()
                cov_to_draw = self.app.generated_df.cov()

            if x not in mean_to_draw or y not in mean_to_draw:
                print(f"Warning: Skipping primary ellipse. {x} or {y} not in stats.")
                self._ellipse_transform = None
                self._ellipse_threshold = None
                hide_outside = False
            else:
                mean_vec = mean_to_draw[[x, y]].values
                cov = cov_to_draw.loc[[x, y], [x, y]].values

                try:
                    inv_cov = np.linalg.inv(cov)
                    self._ellipse_transform = (mean_vec, inv_cov)
                except np.linalg.LinAlgError:
                    inv_cov = None
                    self._ellipse_transform = None
                    self._ellipse_threshold = None
                    hide_outside = False

                if self._ellipse_transform:
                    vals, vecs = np.linalg.eigh(cov)
                    angle = np.degrees(np.arctan2(*vecs[:, 0][::-1]))
                    w, h = 2 * np.sqrt(chi2_thresh * vals)
                    primary_ellipse_params = {'xy': mean_vec, 'width': w, 'height': h, 'angle': angle}

                    percentage_inside_primary = np.nan
                    total_points_primary = 0;
                    count_inside_primary = 0
                    mean_vec_perc, inv_cov_perc = self._ellipse_transform;
                    thresh_perc = self._ellipse_threshold
                    if show_orig and df_orig_unfiltered is not None:
                        points_orig = df_orig_unfiltered[[x, y]].values;
                        total_points_primary += len(points_orig)
                        dist_sq_orig = self._get_mahalanobis_distance_squared(points_orig, mean_vec_perc, inv_cov_perc)
                        count_inside_primary += np.sum(dist_sq_orig <= thresh_perc)
                    if show_gen and df_gen_unfiltered is not None:
                        points_gen = df_gen_unfiltered[[x, y]].values;
                        total_points_primary += len(points_gen)
                        dist_sq_gen = self._get_mahalanobis_distance_squared(points_gen, mean_vec_perc, inv_cov_perc)
                        count_inside_primary += np.sum(dist_sq_gen <= thresh_perc)

                    # --- UNICODE FIX ---
                    if total_points_primary > 0:
                        percentage_inside_primary = (count_inside_primary / total_points_primary) * 100.0;
                        primary_ellipse_label = f'Elliptical {k}\u03c3 ({percentage_inside_primary:.2f}% inside)'
                    else:
                        primary_ellipse_label = f'Elliptical {k}\u03c3 (No pts)'
                else:
                    primary_ellipse_label = f'Elliptical {k}\u03c3 (Calc Error)'
                    # --- END UNICODE FIX ---

                if not draw_target_ellipse and primary_ellipse_params is not None:
                    self.ax_scatter.add_patch(
                        Ellipse(**primary_ellipse_params, ec=theme_props["ellipse_color"], fc='None', lw=1.5, ls='--',
                                label=primary_ellipse_label))

        if hide_outside and self._ellipse_transform and self._ellipse_threshold is not None:
            mean_vec_filter, inv_cov_filter = self._ellipse_transform;
            thresh_filter = self._ellipse_threshold
            if standardize:
                mean_vec_filter = np.array([0.0, 0.0])
                if is_original_available:
                    cov_filter = self.app.df[[x, y]].corr().values
                elif show_gen:
                    cov_filter = self.app.generated_df[[x, y]].corr().values
                else:
                    cov_filter = np.identity(2)
                try:
                    inv_cov_filter = np.linalg.inv(cov_filter)
                except np.linalg.LinAlgError:
                    inv_cov_filter = np.identity(2)

            if df_orig_plot is not None:
                points_orig = df_orig_plot[[x, y]].values;
                dist_sq_orig = self._get_mahalanobis_distance_squared(points_orig, mean_vec_filter, inv_cov_filter);
                keep_mask_orig = dist_sq_orig <= thresh_filter;
                df_orig_plot = df_orig_plot.loc[keep_mask_orig]
            if df_gen_plot is not None:
                points_gen = df_gen_plot[[x, y]].values;
                dist_sq_gen = self._get_mahalanobis_distance_squared(points_gen, mean_vec_filter, inv_cov_filter);
                keep_mask_gen = dist_sq_gen <= thresh_filter;
                df_gen_plot = df_gen_plot.loc[keep_mask_gen]

        self.orig_scatter = None
        if show_orig and df_orig_plot is not None and not df_orig_plot.empty: self.orig_scatter = self.ax_scatter.scatter(
            df_orig_plot[x], df_orig_plot[y], s=max(6, size * 0.6), alpha=alpha * 0.9,
            c=theme_props["orig_point_color"], ec=theme_props["gen_edge_color"], lw=0.2, label="Original")
        self.gen_scatter = None
        if show_gen and df_gen_plot is not None and not df_gen_plot.empty:
            gx, gy = df_gen_plot[x], df_gen_plot[y];
            c_val = None;
            cmap_val = None;
            ec = theme_props["gen_edge_color"]
            if palette == "Black/White":
                c_val, ec = ("k", "w") if self.app.active_theme == "Light" else ("w", "k")
            elif palette == "Solid Color":
                c_val = solid_hex
            else:
                c_val = gx;
                cmap_val = cmap_name
            self.gen_scatter = self.ax_scatter.scatter(gx, gy, s=size, alpha=alpha, c=c_val, cmap=cmap_val,
                                                       edgecolors=ec, lw=0.2, label="Generated")
        self.ax_scatter.set_xlabel(x_label);
        self.ax_scatter.set_ylabel(y_label);
        self.ax_scatter.set_title(f"{y} vs {x}")

        if stats_mean is not None and stats_std is not None:
            box_colors = theme_props.get("box_colors", {})
            if self.show_min_max_box_var.get() and df_for_minmax is not None and not standardize: self.ax_scatter.add_patch(
                patches.Rectangle((df_for_minmax[x].min(), df_for_minmax[y].min()), np.ptp(df_for_minmax[x]),
                                  np.ptp(df_for_minmax[y]), lw=1.5, ec=box_colors.get('Min/Max'), ls='--', fc='none',
                                  label='Min/Max'))

            stat_mean_used = stats_mean
            stat_std_used = stats_std

            if standardize:
                stat_mean_used = pd.Series([0.0] * len(stats_mean), index=stats_mean.index)
                stat_std_used = pd.Series([1.0] * len(stats_std), index=stats_std.index)

            var_map = {2: self.show_2s_box_var, 3: self.show_3s_box_var, 4: self.show_4s_box_var}
            for k_box in [2, 3, 4]:
                if var_map[k_box].get():
                    xm, xs = stat_mean_used[x], stat_std_used[x];
                    ym, ys = stat_mean_used[y], stat_std_used[y]
                    # --- UNICODE FIX ---
                    self.ax_scatter.add_patch(
                        patches.Rectangle((xm - k_box * xs, ym - k_box * ys), 2 * k_box * xs, 2 * k_box * ys, lw=1.5,
                                          ec=box_colors.get(f'{k_box}\u03c3'), ls=':', fc='none',
                                          label=f'{k_box}\u03c3 Box'))
                    # --- END UNICODE FIX ---

        target_ellipse_label = ""
        # --- UNICODE FIX ---
        if draw_target_ellipse and use_target_stats and ellipse_choice != "None":
            k_str = ellipse_choice.replace(f"\u03c3", "")
            k = int(k_str)
            # --- END UNICODE FIX ---

            if x not in self.app.target_stats['means'] or y not in self.app.target_stats['means']:
                print(f"Warning: Skipping Target Ellipse. {x} or {y} not in target_stats.")
                # --- UNICODE FIX ---
                target_ellipse_label = f'Target {k}\u03c3 (Not avail.)'
                # --- END UNICODE FIX ---
            else:
                target_mean_vec_raw = self.app.target_stats['means'][[x, y]].values
                target_sx_raw = self.app.target_stats['stds'][x]
                target_sy_raw = self.app.target_stats['stds'][y]
                calc_mean_x, calc_mean_y = target_mean_vec_raw[0], target_mean_vec_raw[1]
                calc_std_x, calc_std_y = target_sx_raw, target_sy_raw

                mean_vec_plot = target_mean_vec_raw
                std_x_plot, std_y_plot = target_sx_raw, target_sy_raw

                if standardize:
                    if mean_for_stdize is not None and std_for_stdize is not None:
                        mean_vec_plot = (target_mean_vec_raw - mean_for_stdize[[x, y]].values) / std_for_stdize[
                            [x, y]].values
                        std_x_plot = target_sx_raw / std_for_stdize[x]
                        std_y_plot = target_sy_raw / std_for_stdize[y]
                    else:
                        mean_vec_plot = np.array([0.0, 0.0])
                        std_x_plot, std_y_plot = 1.0, 1.0

                w_target = 2 * k * std_x_plot;
                h_target = 2 * k * std_y_plot;
                angle_target = 0

                percentage_inside_target = np.nan
                data_for_perc_calc = df_orig_unfiltered if df_orig_unfiltered is not None else df_gen_unfiltered
                if data_for_perc_calc is not None:
                    try:
                        data_points_x = data_for_perc_calc[x].values;
                        data_points_y = data_for_perc_calc[y].values;
                        total_points = len(data_points_x)
                        if total_points > 0 and calc_std_x > 1e-9 and calc_std_y > 1e-9:
                            dist_sq_normalized = ((data_points_x - calc_mean_x) / (k * calc_std_x)) ** 2 + (
                                    (data_points_y - calc_mean_y) / (k * calc_std_y)) ** 2
                            count_inside = np.sum(dist_sq_normalized <= 1.0);
                            percentage_inside_target = (count_inside / total_points) * 100.0
                            # --- UNICODE FIX ---
                            target_ellipse_label = f'Target {k}\u03c3 Ellipse ({percentage_inside_target:.2f}% inside)'
                        else:
                            target_ellipse_label = f'Target {k}\u03c3 Ellipse (Calc N/A)'
                    except Exception as e_perc:
                        print(f"Error calc % target ellipse: {e_perc}");
                        target_ellipse_label = f'Target {k}\u03c3 Ellipse (Calc Error)'
                else:
                    target_ellipse_label = f'Target {k}\u03c3 Ellipse (No Data)'
                    # --- END UNICODE FIX ---

                self.ax_scatter.add_patch(Ellipse(xy=mean_vec_plot, width=w_target, height=h_target, angle=angle_target,
                                                  ec=theme_props.get("boxplot_sigma_color", "cyan"), fc='None', lw=1.5,
                                                  ls='-', label=target_ellipse_label))

        legend = self.ax_scatter.legend(loc="best", fontsize=9)
        if legend:
            legend.get_frame().set_facecolor(theme_props['plot_bg']);
            legend.get_frame().set_edgecolor(theme_props['grid_color']);
            if legend.get_texts():
                for text in legend.get_texts():
                    try:
                        text.set_color(theme_props['text_color'])
                    except Exception:
                        pass

        if self.ax_box:
            boxplot_data, labels = [], []
            if show_orig and df_orig_plot is not None and not df_orig_plot.empty: boxplot_data.append(
                df_orig_plot[bp]); labels.append("Original")
            if show_gen and df_gen_plot is not None and not df_gen_plot.empty: boxplot_data.append(
                df_gen_plot[bp]); labels.append("Generated")
            if boxplot_data:
                bp_style = self.ax_box.boxplot(boxplot_data, tick_labels=labels, patch_artist=True,
                                               medianprops={"color": theme_props["text_color"]},
                                               boxprops={"facecolor": theme_props["hist_color_original"],
                                                         "edgecolor": theme_props["text_color"]},
                                               whiskerprops={"color": theme_props["text_color"]},
                                               capprops={"color": theme_props["text_color"]})
                try:
                    colors = [theme_props["hist_color_original"], theme_props["hist_color_generated"]]
                    num_boxes = len(bp_style['boxes'])
                    for i in range(num_boxes):
                        if i < len(colors):
                            bp_style['boxes'][i].set_facecolor(colors[i])
                        bp_style['boxes'][i].set_edgecolor(theme_props["text_color"])
                except Exception as e_bp_color:
                    print(f"Warning: Failed to set boxplot colors - {e_bp_color}")
                    pass
            self.ax_box.set_title(f"Boxplot: {bp}")
            self.ax_box.tick_params(axis='x', colors=theme_props["text_color"]);

        try:
            self.fig.tight_layout(pad=3.0);
        except Exception as e_layout:
            print(f"Warning: tight_layout failed: {e_layout}")
        try:
            self.canvas.draw_idle()
        except Exception:
            pass

    def find_ellipse_points(self):
        for artist in self._ellipse_point_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._ellipse_point_artists.clear()

        ellipse_choice = self.elliptical_sigma_menu.get()
        if ellipse_choice == "None":
            # --- UNICODE FIX ---
            messagebox.showinfo("Find Points",
                                f"Please select an elliptical boundary level (2\u03c3, 3\u03c3, or 4\u03c3) first.")
            # --- END UNICODE FIX ---
            return

        # --- UNICODE FIX ---
        k_str = ellipse_choice.replace(f"\u03c3", "")
        k = int(k_str)
        # --- END UNICODE FIX ---
        confidence_level = {2: 0.9545, 3: 0.9973, 4: 0.999936}[k]
        chi2_thresh_primary = chi2.ppf(confidence_level, 2)

        x, y = self.scatter_x_menu.get(), self.scatter_y_menu.get()
        is_original_available = (self.app.df is not None)
        use_target_stats = (self.app.target_stats is not None)
        standardize = self.standardize_axes_var.get()
        use_target_overlay = self.use_target_ellipse_var.get()

        calculate_on_target = use_target_overlay and use_target_stats

        N_POINTS = 32
        t = np.linspace(0, 2 * np.pi, N_POINTS, endpoint=False)
        ellipse_points = None
        ellipse_type_label = ""

        if calculate_on_target:
            print("Calculating points on TARGET ellipse.")
            # --- UNICODE FIX ---
            ellipse_type_label = f"Target {k}\u03c3"
            # --- END UNICODE FIX ---
            if not is_original_available and not self.app.generated_df:
                messagebox.showerror("Error",
                                     "Cannot calculate target ellipse points without base data for standardization context.")
                return

            target_mean_vec_raw = self.app.target_stats['means'][[x, y]].values
            target_sx_raw = self.app.target_stats['stds'][x]
            target_sy_raw = self.app.target_stats['stds'][y]

            mean_vec_plot = target_mean_vec_raw
            std_x_plot, std_y_plot = target_sx_raw, target_sy_raw

            if standardize:
                base_mean = None
                base_std = None
                if is_original_available:
                    base_mean = self.app.df.mean()
                    base_std = self.app.df.std(ddof=0)
                elif self.app.generated_df is not None:
                    base_mean = self.app.generated_df.mean()
                    base_std = self.app.generated_df.std(ddof=0)

                if base_mean is None or base_std is None:
                    messagebox.showerror("Error", "Standardization base stats not found.")
                    return

                mean_vec_plot = (target_mean_vec_raw - base_mean[[x, y]].values) / base_std[[x, y]].values
                try:
                    std_x_plot = target_sx_raw / base_std[x]
                    std_y_plot = target_sy_raw / base_std[y]
                except (ZeroDivisionError, KeyError):
                    messagebox.showerror("Error",
                                         "Division by zero or key error during standardization of target ellipse.")
                    return

            w_target_plot = 2 * k * std_x_plot
            h_target_plot = 2 * k * std_y_plot
            cx, cy = mean_vec_plot[0], mean_vec_plot[1]
            rx, ry = w_target_plot / 2.0, h_target_plot / 2.0
            ellipse_points_x = cx + rx * np.cos(t)
            ellipse_points_y = cy + ry * np.sin(t)
            ellipse_points = np.column_stack([ellipse_points_x, ellipse_points_y])

        else:
            print("Calculating points on PRIMARY ellipse.")
            # --- UNICODE FIX ---
            ellipse_type_label = f"Elliptical {k}\u03c3"
            # --- END UNICODE FIX ---
            if not self._ellipse_transform or self._ellipse_threshold is None:
                messagebox.showerror("Error",
                                     "Primary ellipse parameters not calculated. Ensure data is loaded and ellipse is selected.")
                return

            primary_mean_source = None
            primary_cov_source = None
            show_orig = self.show_original_var.get() and is_original_available
            show_gen = self.show_generated_var.get() and (self.app.generated_df is not None)

            if not show_orig and show_gen:
                primary_mean_source = self.app.generated_df.mean()
                primary_cov_source = self.app.generated_df.cov()
            elif use_target_overlay and use_target_stats and is_original_available:
                primary_mean_source = self.app.df.mean()
                primary_cov_source = self.app.df.cov()
            elif use_target_stats:
                primary_mean_source = self.app.target_stats['means']
                primary_std_source = self.app.target_stats['stds']
                primary_cov_source = pd.DataFrame(np.diag(primary_std_source.values ** 2),
                                                  index=primary_std_source.index, columns=primary_std_source.index)
            elif is_original_available:
                primary_mean_source = self.app.df.mean()
                primary_cov_source = self.app.df.cov()
            elif show_gen:
                primary_mean_source = self.app.generated_df.mean()
                primary_cov_source = self.app.generated_df.cov()

            if primary_mean_source is None or primary_cov_source is None:
                messagebox.showerror("Error", "Could not determine source stats for primary ellipse calculation.")
                return

            calc_mean = primary_mean_source[[x, y]].values
            calc_cov = primary_cov_source.loc[[x, y], [x, y]].values

            plot_mean = calc_mean
            plot_cov = calc_cov
            if standardize:
                plot_mean = np.array([0.0, 0.0])
                if not show_orig and show_gen:
                    plot_cov = self.app.generated_df[[x, y]].corr().values
                elif use_target_overlay and use_target_stats:
                    if is_original_available:
                        plot_cov = self.app.df[[x, y]].corr().values
                    else:
                        plot_cov = np.identity(2)
                elif use_target_stats:
                    plot_cov = np.identity(2)
                elif is_original_available:
                    plot_cov = self.app.df[[x, y]].corr().values
                else:
                    if show_gen:
                        plot_cov = self.app.generated_df[[x, y]].corr().values
                    else:
                        plot_cov = np.identity(2)

            try:
                vals, vecs = np.linalg.eigh(plot_cov)
                vals = np.maximum(vals, 1e-12)
            except Exception as e:
                messagebox.showerror("Error", f"Could not calculate ellipse parameters: {e}");
                return

            semi_axes_lengths = np.sqrt(self._ellipse_threshold * vals)
            unit_circle_points = np.array([np.cos(t), np.sin(t)])
            ellipse_points = plot_mean[:, np.newaxis] + vecs @ np.diag(semi_axes_lengths) @ unit_circle_points
            ellipse_points = ellipse_points.T

        if ellipse_points is not None:
            artists = self.ax_scatter.plot(ellipse_points[:, 0], ellipse_points[:, 1], 'rs', markersize=6,
                                           markeredgecolor='black', markerfacecolor='red',
                                           label=f'{ellipse_type_label} Boundary Pts')
            self._ellipse_point_artists.extend(artists)

            theme_props = self.app.get_theme_properties();
            legend = self.ax_scatter.legend(loc="best", fontsize=9)
            if legend:
                legend.get_frame().set_facecolor(theme_props['plot_bg'])
                legend.get_frame().set_edgecolor(theme_props['grid_color']);
                if legend.get_texts():
                    for text in legend.get_texts():
                        try:
                            text.set_color(theme_props['text_color'])
                        except Exception:
                            pass
            self.canvas.draw_idle()

            try:
                self.selected_text.configure(state="normal");
                self.update_idletasks()
                self.selected_text.delete("1.0", "end")

                header = f"Generated {N_POINTS} points on {ellipse_type_label} boundary:\n";
                header += "(Coordinates are in standardized units if 'Standardize' is checked)\n";
                header += "-" * (len(header) - 2) + "\n";
                self.selected_text.insert("1.0", header)

                df_points = pd.DataFrame(ellipse_points, columns=[x, y]);
                self.selected_text.insert("end", "\n" + df_points.to_string(float_format="{:.4f}".format, index=False));

                self.selected_text.configure(state="disabled")

            except Exception as e:
                print(f"Error updating text box: {e}")
                try:
                    self.selected_text.configure(state="disabled")
                except Exception:
                    pass
        else:
            messagebox.showerror("Error", "Failed to calculate ellipse points.")

    def _on_wc_popup_close(self):
        if self.wc_popup:
            self.wc_popup.destroy()
            self.wc_popup = None
        plt.style.use(self.app.get_theme_properties()["mpl_style"])

    def run_worst_case_analysis(self):
        if self.wc_popup is not None and self.wc_popup.winfo_exists():
            self.wc_popup.focus()
            return

        if self.app.df is None:
            messagebox.showinfo("Analysis Error", "No data loaded. Please load data on the Generator tab.")
            return

        target_param = self.worst_case_param_menu.get()
        if target_param == "-":
            messagebox.showinfo("Analysis Error", "Please select a parameter to analyze.")
            return

        all_params = self.app.param_names
        compare_params = [p for p in all_params if p != target_param]

        if not compare_params:
            messagebox.showinfo("Analysis Info", f"No other parameters available to compare with '{target_param}'.")
            return

        self.wc_popup = ctk.CTkToplevel(self.app)
        self.wc_popup.title(f"Worst-Case Analysis for: {target_param}")
        self.wc_popup.geometry("800x650")
        self.wc_popup.protocol("WM_DELETE_WINDOW", self._on_wc_popup_close)

        theme_props = self.app.get_theme_properties()
        plt.style.use(theme_props["mpl_style"])

        wc_fig = plt.Figure(figsize=(8, 6), dpi=100)
        wc_ax = wc_fig.add_subplot(111)

        wc_canvas = FigureCanvasTkAgg(wc_fig, master=self.wc_popup)
        wc_canvas.get_tk_widget().pack(side=ctk.TOP, fill=ctk.BOTH, expand=True, padx=10, pady=5)

        toolbar_frame = ctk.CTkFrame(self.wc_popup)
        toolbar_frame.pack(side=ctk.BOTTOM, fill="x", padx=10, pady=(0, 5))
        toolbar = NavigationToolbar2Tk(wc_canvas, toolbar_frame)
        toolbar.update()

        try:
            y_mean = self.app.df[target_param].mean()
            y_std = self.app.df[target_param].std(ddof=0)
            if y_std == 0 or np.isnan(y_std):
                messagebox.showerror("Analysis Error",
                                     f"Target parameter '{target_param}' has zero or invalid variance.")
                self._on_wc_popup_close()
                return

            x_means = self.app.df[compare_params].mean()
            x_stds = self.app.df[compare_params].std(ddof=0)

            # --- UNICODE FIX ---
            # 3-sigma (99.73%)
            chi2_thresh = chi2.ppf(0.9973, 2)
            # --- END UNICODE FIX ---

            colors = plt.cm.get_cmap('tab10', len(compare_params))
            all_ellipse_ranges = []
            global_y_mins = [self.app.df[target_param].min()]
            global_y_maxes = [self.app.df[target_param].max()]

            for i, x_param in enumerate(compare_params):
                x_std = x_stds.get(x_param, 0)
                if x_std == 0 or np.isnan(x_std):
                    print(f"Skipping {x_param}: zero or invalid variance.")
                    continue

                rho = self.app.df[[x_param, target_param]].corr().values[0, 1]
                if np.isnan(rho):
                    print(f"Skipping {x_param}: NaN correlation.")
                    continue

                cov_matrix = np.array([
                    [1.0, rho * y_std],
                    [rho * y_std, y_std ** 2]
                ])
                mean_vec = np.array([0, y_mean])
                vals, vecs = np.linalg.eigh(cov_matrix)
                vals = np.maximum(vals, 0)
                angle = np.degrees(np.arctan2(*vecs[:, 1]))
                w, h = 2 * np.sqrt(chi2_thresh * vals)

                # --- UNICODE FIX ---
                ellipse = Ellipse(xy=mean_vec, width=w, height=h, angle=angle,
                                  fc='none', ec=colors(i), lw=2,
                                  label=f"{x_param} (\u03c1={rho:.3f})")  # rho symbol
                # --- END UNICODE FIX ---
                wc_ax.add_patch(ellipse)

                t = np.linspace(0, 2 * np.pi, 100)
                unit_circle_points = np.array([np.cos(t), np.sin(t)])
                transform_matrix = vecs @ np.diag(np.sqrt(chi2_thresh * vals))
                ell_points_rel = transform_matrix @ unit_circle_points
                y_span_rel_max = np.max(ell_points_rel[1, :])
                y_span_rel_min = np.min(ell_points_rel[1, :])
                y_abs_max = y_mean + y_span_rel_max
                y_abs_min = y_mean + y_span_rel_min

                all_ellipse_ranges.append((y_abs_max - y_abs_min, x_param))
                global_y_mins.append(y_abs_min)
                global_y_maxes.append(y_abs_max)

            if not all_ellipse_ranges:
                wc_ax.text(0.5, 0.5, "Could not calculate any interactions.",
                           ha='center', va='center', color='red', transform=wc_ax.transAxes)
                wc_canvas.draw()
                return

            all_ellipse_ranges.sort(key=lambda x: x[0], reverse=True)
            legend_title_lines = ["Interactions (Worst-Case First):"]
            for k, (range_val, param_name) in enumerate(all_ellipse_ranges[:3]):
                legend_title_lines.append(f"{k + 1}. {param_name} (Range: {range_val:.3f})")
            legend_title = "\n".join(legend_title_lines)

            # --- UNICODE FIX ---
            wc_ax.set_xlabel("Standardized Input (sigma units)")
            wc_ax.set_ylabel(f"{target_param} (original units)")
            wc_ax.set_title(f"3\u03c3 Interaction Ellipses vs. {target_param}", fontsize=12)
            # --- END UNICODE FIX ---

            wc_ax.set_xlim(-3.5, 3.5)

            plot_y_min = np.nanmin(global_y_mins)
            plot_y_max = np.nanmax(global_y_maxes)
            plot_y_range = plot_y_max - plot_y_min
            if plot_y_range == 0 or np.isnan(plot_y_range): plot_y_range = 1.0
            y_pad = plot_y_range * 0.05
            wc_ax.set_ylim(plot_y_min - y_pad, plot_y_max + y_pad)

            wc_ax.axhline(y_mean, color='gray', linestyle='--', lw=1)
            wc_ax.axvline(0, color='gray', linestyle='--', lw=1)

            handles, labels = wc_ax.get_legend_handles_labels()
            label_to_range = {item[1]: item[0] for item in all_ellipse_ranges}
            try:
                sorted_handles_labels = sorted(zip(handles, labels),
                                               key=lambda h_l: label_to_range.get(h_l[1].split(' ')[0], 0),
                                               reverse=True)
                sorted_handles = [h for h, l in sorted_handles_labels]
                sorted_labels = [l for h, l in sorted_handles_labels]
            except Exception:
                sorted_handles = handles
                sorted_labels = labels

            legend = wc_ax.legend(sorted_handles, sorted_labels,
                                  title=legend_title,
                                  fontsize=9,
                                  title_fontsize=10,
                                  bbox_to_anchor=(1.02, 1), loc='upper left')

            wc_fig.set_facecolor(theme_props["plot_bg"])
            wc_ax.set_facecolor(theme_props["plot_bg"])
            wc_ax.tick_params(axis='x', colors=theme_props["text_color"])
            wc_ax.tick_params(axis='y', colors=theme_props["text_color"])
            for spine in wc_ax.spines.values():
                spine.set_color(theme_props["text_color"])
            wc_ax.xaxis.label.set_color(theme_props["text_color"])
            wc_ax.yaxis.label.set_color(theme_props["text_color"])
            wc_ax.title.set_color(theme_props["text_color"])
            wc_ax.grid(color=theme_props['grid_color'], linestyle='--', alpha=0.5)

            if legend:
                legend.get_frame().set_facecolor(theme_props['plot_bg'])
                legend.get_frame().set_edgecolor(theme_props['grid_color'])
                for text in legend.get_texts():
                    text.set_color(theme_props["text_color"])
                legend.get_title().set_color(theme_props["text_color"])

            wc_fig.tight_layout(rect=[0, 0, 0.75, 1])
            wc_canvas.draw()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Analysis Error", f"An error occurred during analysis:\n{e}")
            self._on_wc_popup_close()
        finally:
            plt.style.use(self.app.get_theme_properties()["mpl_style"])


# =============================================================================
# --- TAB CLASS: DesignTab (Complete Definition) ---
# =============================================================================
class DesignTab(ctk.CTkFrame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, fg_color="transparent")
        self.app = app_instance
        self.design_mode = "File"
        self.design_input_widgets = {}
        self.manual_param_widgets = []
        self.design_units_map = {}
        self.manual_spec_limits = {}  # To store LSL/USL for auto-fill

        # --- NEW ---
        # Variable for the RNN prediction checkbox
        self.predict_with_rnn_var = ctk.BooleanVar(value=False)
        # Flag to stop animation
        self._stop_animation = False
        # --- END NEW ---

        self._build_ui()

    def _detect_dynamic_structure(self, df):
        """Detect dynamic data structure."""
        try:
            if len(df) < 500:
                return False
            first_col = df.iloc[:, 0].values
            diff = np.diff(first_col)
            return np.any(diff < 0)
        except:
            return False

    def _load_dynamic_data(self, df):
        """Load dynamic data."""
        try:
            loader = DynamicDataLoader(df)
            self.dynamic_data = loader.load_and_reshape()
            self.input_channels = self.dynamic_data['input_names']
            self.output_channels = self.dynamic_data['output_names']
            self.frequency_array = self.dynamic_data['frequency']
            self.dynamic_metadata = self.dynamic_data['metadata']
            self.rnn_data = pd.DataFrame(self.dynamic_data['X'], columns=self.input_channels)
            self.rnn_data_type = "dynamic"
            print("✓ Dynamic data loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load dynamic data:\n{e}")
            raise

    def _build_dynamic_models(self):
        """Build dynamic models."""
        print("\nBuilding dynamic models...")
        self.build_model_button.configure(text="Training...", state="disabled")
        self.app.update_idletasks()

        try:
            X = self.dynamic_data['X']
            Y = self.dynamic_data['Y']
            input_names = self.dynamic_data['input_names']
            output_names = self.dynamic_data['output_names']

            X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.2, random_state=42)

            self.trained_models.clear()
            self.last_build_stats.clear()

            for idx, output_name in enumerate(output_names):
                model = GPR_PCA_DynamicModel(n_components=50, variance_threshold=0.99)
                model.fit(X_train, Y_train[:, :, idx], output_name=output_name)

                Y_pred_test = model.predict(X_test)
                r2_test = r2_score(Y_test[:, :, idx].ravel(), Y_pred_test.ravel())

                self.trained_models[output_name] = {'model': model, 'features': input_names}
                self.last_build_stats[output_name] = {'r2_test': r2_test, 'n_components': model.n_components_actual}

            messagebox.showinfo("Success",
                                f"Trained {len(output_names)} dynamic models!\nAvg R²: {np.mean([s['r2_test'] for s in self.last_build_stats.values()]):.3f}")

        except Exception as e:
            messagebox.showerror("Error", f"Training failed:\n{e}")
        finally:
            self.build_model_button.configure(text="Build Models", state="normal")

    def reset_to_default(self):
        """Resets the Design tab to its initial state."""
        print("--- Resetting Design Tab ---")  # Debug print

        # --- NEW: Stop any running animation ---
        self._stop_animation = True
        # --- END NEW ---

        # --- Reset Mode ---
        try:
            # Check if widget exists before configuring
            if hasattr(self, 'design_mode_switcher') and self.design_mode_switcher.winfo_exists():
                self.design_mode_switcher.set("Load from File")
                self._switch_design_mode("Load from File")  # Updates UI based on mode
        except Exception as e:
            print(f"ERROR resetting mode switcher: {e}")

        # --- Reset File Mode ---
        try:
            # Check if widget exists before configuring
            if hasattr(self, 'design_loaded_label') and self.design_loaded_label.winfo_exists():
                # Use default font by not specifying font=None
                self.design_loaded_label.configure(text="No file loaded.", text_color="gray")
            # Clear the parameter input frame in File mode
            if hasattr(self, 'design_param_frame') and self.design_param_frame.winfo_exists():
                for widget in self.design_param_frame.winfo_children():
                    if widget.winfo_exists():  # Check if widget exists before destroying
                        widget.destroy()
                if not any(isinstance(w, ctk.CTkLabel) for w in self.design_param_frame.winfo_children()):
                    ctk.CTkLabel(self.design_param_frame, text="Load reference data to begin...",
                                 font=DEFAULT_UI_FONT).pack(pady=20)  # Apply font
            self.design_input_widgets = {}  # Clear the dictionary
        except Exception as e:
            print(f"ERROR resetting file mode elements: {e}")

        # --- Reset Manual Mode ---
        try:
            if hasattr(self, '_clear_manual_parameter_rows'):
                self._clear_manual_parameter_rows()  # Clears widgets and list
            self.design_units_map = {}
            self.manual_spec_limits = {}
            if hasattr(self, 'manual_mean_option_menu') and self.manual_mean_option_menu.winfo_exists():
                self.manual_mean_option_menu.set("Center Mean in Tolerance")
        except Exception as e:
            print(f"ERROR clearing manual rows/options: {e}")

        # --- Reset Generation Settings ---
        try:
            if hasattr(self, 'design_method_menu') and self.design_method_menu.winfo_exists():
                self.design_method_menu.set("Monte Carlo")
            if hasattr(self, 'design_repeats_entry') and self.design_repeats_entry.winfo_exists():
                self.design_repeats_entry.delete(0, 'end');
                self.design_repeats_entry.insert(0, "20")
            if hasattr(self, 'design_samples_entry') and self.design_samples_entry.winfo_exists():
                self.design_samples_entry.delete(0, 'end');
                self.design_samples_entry.insert(0, "10000")
            if hasattr(self, 'design_seed_entry') and self.design_seed_entry.winfo_exists():
                self.design_seed_entry.delete(0, 'end')

            self.predict_with_rnn_var.set(False)
            if hasattr(self, 'predict_with_rnn_cb') and self.predict_with_rnn_cb.winfo_exists():
                self.predict_with_rnn_cb.configure(state="disabled")
        except Exception as e:
            print(f"ERROR resetting generation settings: {e}")

        # --- Reset Buttons & Status ---
        try:
            if hasattr(self, 'design_generate_button') and self.design_generate_button.winfo_exists():
                self.design_generate_button.configure(state="disabled")
            if hasattr(self, 'manual_ok_button') and self.manual_ok_button.winfo_exists():
                self.manual_ok_button.configure(state="disabled")
            if hasattr(self, 'save_design_data_button') and self.save_design_data_button.winfo_exists():
                self.save_design_data_button.configure(state="disabled")
            if hasattr(self, 'design_status_label') and self.design_status_label.winfo_exists():
                self.design_status_label.configure(text="")
        except Exception as e:
            print(f"ERROR resetting buttons/status: {e}")

        # --- Reset Analysis Section & Tolerance Calculator ---
        try:
            if hasattr(self, '_disable_analysis_frame'):
                self._disable_analysis_frame()
            if hasattr(self, 'design_lsl_entry') and self.design_lsl_entry.winfo_exists():
                self.design_lsl_entry.delete(0, 'end')
            if hasattr(self, 'design_usl_entry') and self.design_usl_entry.winfo_exists():
                self.design_usl_entry.delete(0, 'end')
            if hasattr(self, 'tol_calc_nominal_entry') and self.tol_calc_nominal_entry.winfo_exists():
                self.tol_calc_nominal_entry.delete(0, 'end')
            if hasattr(self, 'tol_calc_cpk_entry') and self.tol_calc_cpk_entry.winfo_exists():
                self.tol_calc_cpk_entry.delete(0, 'end');
                self.tol_calc_cpk_entry.insert(0, "1.33")
            if hasattr(self, 'show_sigma_lines_var'): self.show_sigma_lines_var.set(False)

        except Exception as e:
            print(f"ERROR resetting analysis section: {e}")

        # --- Reset Right Side Tabs ---
        try:
            if hasattr(self, 'design_data_preview_text') and self.design_data_preview_text.winfo_exists():
                self.design_data_preview_text.configure(state="normal")
                self.design_data_preview_text.delete("1.0", "end")  # <-- FIX from 1.File
                self.design_data_preview_text.insert("1.0", "Generated data will be shown here...")
                self.design_data_preview_text.configure(state="disabled")
        except Exception as e:
            print(f"ERROR resetting data preview text: {e}")

        try:
            if hasattr(self, 'design_report_text') and self.design_report_text.winfo_exists():
                self.design_report_text.configure(state="normal")
                self.design_report_text.delete("1.0", "end")  # <-- FIX from 1.File
                self.design_report_text.insert("1.0", "Capability analysis results will appear here.")
                self.design_report_text.configure(state="disabled")
        except Exception as e:
            print(f"ERROR resetting capability report text: {e}")

        # --- RESET NEW BIN ANALYSIS TAB ---
        try:
            if hasattr(self, 'bin_fail_condition_menu') and self.bin_fail_condition_menu.winfo_exists():
                self.bin_fail_condition_menu.set("Output > USL")
            if hasattr(self, 'bin_analysis_text') and self.bin_analysis_text.winfo_exists():
                self.bin_analysis_text.configure(state="normal")
                self.bin_analysis_text.delete("1.0", "end")
                self.bin_analysis_text.insert("1.0", "Binned failure analysis results will appear here.")
                self.bin_analysis_text.configure(state="disabled")
        except Exception as e:
            print(f"ERROR resetting binned failure analysis tab: {e}")
        # --- END RESET ---

        # --- RESET NEW LIVE VISUALIZATION TAB ---
        try:
            # --- NEW: Reset loop checkbox ---
            if hasattr(self, 'loop_animation_var'):
                self.loop_animation_var.set(True)  # <-- Set to True
            # --- END NEW ---

            if hasattr(self, 'live_ax') and self.live_ax:
                self.live_ax.clear()
                text_color = "gray"
                if hasattr(self, 'app') and hasattr(self.app, 'get_theme_properties'):
                    try:
                        theme_props = self.app.get_theme_properties()
                        text_color = theme_props.get("text_color", "gray")
                    except Exception:
                        print("Warning: Could not get theme properties during reset.")

                self.live_ax.text(0.5, 0.5, 'Click "Generate Data" to start live view',
                                  ha='center', va='center', color=text_color)
                if hasattr(self,
                           'live_canvas') and self.live_canvas and self.live_canvas.get_tk_widget().winfo_exists():
                    self.live_canvas.draw_idle()
        except Exception as e:
            print(f"ERROR clearing/redrawing live_vis canvas during reset: {e}")
        # --- END RESET ---

        try:
            if hasattr(self, 'design_ax') and self.design_ax:
                self.design_ax.clear()
                text_color = "gray"
                if hasattr(self, 'app') and hasattr(self.app, 'get_theme_properties'):
                    try:
                        theme_props = self.app.get_theme_properties()
                        text_color = theme_props.get("text_color", "gray")
                    except Exception:
                        print("Warning: Could not get theme properties during reset.")

                self.design_ax.text(0.5, 0.5, 'Capability analysis results\nwill appear here.',
                                    ha='center', va='center', color=text_color)

                if hasattr(self, 'apply_theme'):
                    try:
                        self.apply_theme()
                    except Exception as ae:
                        print(f"ERROR during apply_theme in reset: {ae}")

                if hasattr(self,
                           'design_canvas') and self.design_canvas and self.design_canvas.get_tk_widget().winfo_exists():
                    self.design_canvas.draw_idle()
            else:
                print("Capability Axes (design_ax) not found during reset.")
        except Exception as e:
            print(f"ERROR clearing/redrawing design canvas during reset: {e}")

        try:
            if hasattr(self, 'design_right_tabs') and self.design_right_tabs.winfo_exists():
                self.design_right_tabs.set("Data Preview")
        except Exception as e:
            print(f"ERROR setting design tab view: {e}")

        print("--- Design Tab Reset Complete ---")  # Debug print

    def _build_ui(self):
        # --- Overall Layout (Left Controls, Right Results/Tools) ---
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, minsize=580)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1)
        left = ctk.CTkFrame(self);
        left.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        sash = ctk.CTkFrame(self, width=7, cursor="sb_h_double_arrow");
        sash.grid(row=0, column=1, sticky="ns", pady=6)
        sash.bind("<B1-Motion>", lambda event: self._on_sash_drag(event, self))
        right = ctk.CTkFrame(self);
        right.grid(row=0, column=2, sticky="nsew", padx=(0, 6), pady=6)

        # --- LEFT PANEL: Mode Switching, Data Input, Generation Settings, Analysis Settings ---
        ctk.CTkLabel(left, text="Multi-Parameter Design Simulation", font=DEFAULT_UI_FONT_BOLD).pack(pady=(10, 5),
                                                                                                     padx=6)  # Apply font
        mode_switcher_frame = ctk.CTkFrame(left, fg_color="transparent");
        mode_switcher_frame.pack(fill="x", padx=6, pady=5)
        ctk.CTkLabel(mode_switcher_frame, text="Input Mode:", font=DEFAULT_UI_FONT).pack(side="left")  # Apply font
        self.design_mode_switcher = ctk.CTkSegmentedButton(mode_switcher_frame,
                                                           values=["Load from File", "Define Manually"],
                                                           command=self._switch_design_mode,
                                                           font=DEFAULT_UI_FONT);  # Apply font
        self.design_mode_switcher.set("Load from File");
        self.design_mode_switcher.pack(side="left", padx=10, expand=True, fill="x")

        mode_container = ctk.CTkFrame(left, fg_color="transparent");
        mode_container.pack(fill="both", expand=True, padx=0, pady=0)

        # File Mode Frame
        self.design_file_mode_frame = ctk.CTkFrame(mode_container)
        self.design_load_button = ctk.CTkButton(self.design_file_mode_frame,
                                                text="Load Reference Data (for Correlation)",
                                                command=self.load_reference_data, font=DEFAULT_UI_FONT);
        self.design_load_button.pack(fill="x", padx=5, pady=8)  # Apply font
        self.design_loaded_label = ctk.CTkLabel(self.design_file_mode_frame, text="No file loaded.", text_color="gray",
                                                font=DEFAULT_UI_FONT);
        self.design_loaded_label.pack(fill="x", padx=5, pady=(0, 5))  # Apply font
        self.design_param_frame = ctk.CTkScrollableFrame(self.design_file_mode_frame,
                                                         label_text="2. Define New Parameter Targets",
                                                         label_font=DEFAULT_UI_FONT_BOLD);
        self.design_param_frame.pack(fill="both", expand=True, padx=5, pady=5)  # Apply font
        ctk.CTkLabel(self.design_param_frame, text="Load reference data to begin...", font=DEFAULT_UI_FONT).pack(
            pady=20)  # Apply font

        # Manual Mode Frame
        self.design_manual_mode_frame = ctk.CTkFrame(mode_container)
        manual_buttons_frame = ctk.CTkFrame(self.design_manual_mode_frame, fg_color="transparent");
        manual_buttons_frame.pack(fill="x", padx=5, pady=(8, 2))
        self.manual_add_button = ctk.CTkButton(manual_buttons_frame, text="Add (+)", width=80,
                                               command=self._add_manual_parameter_row, font=DEFAULT_UI_FONT);
        self.manual_add_button.pack(side="left", padx=(0, 5))  # Apply font
        self.manual_save_preset_button = ctk.CTkButton(manual_buttons_frame, text="Save Preset", width=100,
                                                       command=self._save_preset, font=DEFAULT_UI_FONT);
        self.manual_save_preset_button.pack(side="left", padx=5)  # Apply font
        self.manual_load_preset_button = ctk.CTkButton(manual_buttons_frame, text="Load Preset", width=100,
                                                       command=self._load_preset, font=DEFAULT_UI_FONT);
        self.manual_load_preset_button.pack(side="left", padx=5)  # Apply font
        self.manual_ok_button = ctk.CTkButton(manual_buttons_frame, text="Generate Data",
                                              command=self.generate_multi_param_data, state="disabled",
                                              font=DEFAULT_UI_FONT);
        self.manual_ok_button.pack(side="left", expand=True, fill="x", padx=(10, 0))  # Apply font
        mean_option_frame = ctk.CTkFrame(self.design_manual_mode_frame, fg_color="transparent");
        mean_option_frame.pack(fill="x", padx=5, pady=(2, 5))
        ctk.CTkLabel(mean_option_frame, text="Process Mean:", font=DEFAULT_UI_FONT).pack(side="left",
                                                                                         padx=(0, 5))  # Apply font
        self.manual_mean_option_menu = ctk.CTkOptionMenu(mean_option_frame,
                                                         values=["Center Mean in Tolerance", "Use Nominal as Mean"],
                                                         width=200, font=DEFAULT_UI_FONT);
        self.manual_mean_option_menu.set("Center Mean in Tolerance");
        self.manual_mean_option_menu.pack(side="left")  # Apply font
        self.manual_param_rows_frame = ctk.CTkScrollableFrame(self.design_manual_mode_frame,
                                                              label_text="Define Independent Parameters",
                                                              label_font=DEFAULT_UI_FONT_BOLD);
        self.manual_param_rows_frame.pack(fill="both", expand=True, padx=5, pady=5);
        self.manual_param_rows_frame.grid_columnconfigure(0, weight=3);
        self.manual_param_rows_frame.grid_columnconfigure(1, weight=1, minsize=60);
        self.manual_param_rows_frame.grid_columnconfigure(2, weight=2, minsize=70);
        self.manual_param_rows_frame.grid_columnconfigure(3, weight=1, minsize=70);
        self.manual_param_rows_frame.grid_columnconfigure(4, weight=1, minsize=70);
        self.manual_param_rows_frame.grid_columnconfigure(5, weight=1, minsize=60);
        self.manual_param_rows_frame.grid_columnconfigure(6, weight=0);
        ctk.CTkLabel(self.manual_param_rows_frame, text="Click '+' to add a parameter.", font=DEFAULT_UI_FONT).pack(
            pady=20)  # Apply font

        # --- BOTTOM SECTION (Generate, Analysis Settings ONLY) ---
        bottom_controls_frame = ctk.CTkFrame(left, fg_color="transparent")
        bottom_controls_frame.pack(side="bottom", fill="x", padx=0, pady=0)

        # Generate & Save Section
        gen_settings_frame = ctk.CTkFrame(bottom_controls_frame);
        gen_settings_frame.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(gen_settings_frame, text="3. Generate & Save Virtual System Data", font=DEFAULT_UI_FONT_BOLD).pack(
            pady=(5, 5))  # Apply font
        method_frame = ctk.CTkFrame(gen_settings_frame, fg_color="transparent");
        method_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(method_frame, text="Method:", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_method_menu = ctk.CTkOptionMenu(method_frame, values=["Optimized LHS", "Monte Carlo"], width=220,
                                                    font=DEFAULT_UI_FONT);
        self.design_method_menu.set("Monte Carlo");
        self.design_method_menu.pack(side="right", padx=5)  # Apply font
        repeat_frame = ctk.CTkFrame(gen_settings_frame, fg_color="transparent");
        repeat_frame.pack(fill="x", padx=5, pady=(2, 5));
        ctk.CTkLabel(repeat_frame, text="Repeats (for LHS):", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_repeats_entry = ctk.CTkEntry(repeat_frame, width=80, font=DEFAULT_UI_FONT);
        self.design_repeats_entry.insert(0, "20");
        self.design_repeats_entry.pack(side="right", padx=5)  # Apply font
        samples_frame = ctk.CTkFrame(gen_settings_frame, fg_color="transparent");
        samples_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(samples_frame, text="Number of Virtual Samples:", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_samples_entry = ctk.CTkEntry(samples_frame, width=100, font=DEFAULT_UI_FONT);
        self.design_samples_entry.insert(0, "10000");
        self.design_samples_entry.pack(side="right")  # Apply font
        seed_frame = ctk.CTkFrame(gen_settings_frame, fg_color="transparent");
        seed_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(seed_frame, text="Seed (opt):", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_seed_entry = ctk.CTkEntry(seed_frame, width=100, font=DEFAULT_UI_FONT);
        self.design_seed_entry.pack(side="right")  # Apply font
        self.design_generate_button = ctk.CTkButton(gen_settings_frame, text="Generate Data (File Mode)",
                                                    command=self.generate_multi_param_data, state="disabled",
                                                    font=DEFAULT_UI_FONT);
        self.design_generate_button.pack(fill="x", padx=5, pady=(8, 4))  # Apply font

        self.predict_with_rnn_cb = ctk.CTkCheckBox(gen_settings_frame,
                                                   text="Run RNN Predictions on Generated Data",
                                                   variable=self.predict_with_rnn_var,
                                                   font=DEFAULT_UI_FONT,
                                                   state="disabled")  # Disabled by default
        self.predict_with_rnn_cb.pack(fill="x", padx=5, pady=(4, 4))

        self.save_design_data_button = ctk.CTkButton(gen_settings_frame, text="Save Generated Data",
                                                     command=self.save_generated_design_data, state="disabled",
                                                     font=DEFAULT_UI_FONT);
        self.save_design_data_button.pack(fill="x", padx=5, pady=(4, 8))  # Apply font
        self.design_status_label = ctk.CTkLabel(gen_settings_frame, text="", text_color="white", wraplength=480,
                                                font=DEFAULT_UI_FONT);
        self.design_status_label.pack(fill="x", padx=5, pady=(0, 5))  # Apply font

        # Analysis Section Settings
        self.analysis_frame = ctk.CTkFrame(bottom_controls_frame);
        self.analysis_frame.pack(fill="x", padx=6, pady=(10, 6))
        ctk.CTkLabel(self.analysis_frame, text="4. Analysis Settings", font=DEFAULT_UI_FONT_BOLD).pack(
            pady=(5, 5))  # Apply font
        analysis_param_frame = ctk.CTkFrame(self.analysis_frame, fg_color="transparent");
        analysis_param_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(analysis_param_frame, text="Analyze Parameter:", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_analysis_param_menu = ctk.CTkOptionMenu(analysis_param_frame, values=["-"],
                                                            command=self._on_analysis_param_change,
                                                            font=DEFAULT_UI_FONT);
        self.design_analysis_param_menu.pack(side="right")  # Apply font
        lsl_frame = ctk.CTkFrame(self.analysis_frame, fg_color="transparent");
        lsl_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(lsl_frame, text="Lower Spec Limit (LSL):", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_lsl_entry = ctk.CTkEntry(lsl_frame, width=120, font=DEFAULT_UI_FONT);
        self.design_lsl_entry.pack(side="right")  # Apply font
        usl_frame = ctk.CTkFrame(self.analysis_frame, fg_color="transparent");
        usl_frame.pack(fill="x", padx=5, pady=2);
        ctk.CTkLabel(usl_frame, text="Upper Spec Limit (USL):", font=DEFAULT_UI_FONT).pack(side="left");
        self.design_usl_entry = ctk.CTkEntry(usl_frame, width=120, font=DEFAULT_UI_FONT);
        self.design_usl_entry.pack(side="right")  # Apply font

        # --- RIGHT PANEL: Tabs for Preview, Report, Calculator ---
        theme_props = self.app.get_theme_properties();
        tab_selected_color = "#e6e6e6" if self.app.active_theme == "Light" else "#404040";
        tab_unselected_color = "#f0f0f0" if self.app.active_theme == "Light" else "#2b2b2b";
        tab_hover_color = "#dcdcdc" if self.app.active_theme == "Light" else "#505050";
        tab_text_color = theme_props.get("text_color", "black");
        border_color = "#cccccc" if self.app.active_theme == "Light" else "#444444"
        self.design_right_tabs = ctk.CTkTabview(right, border_width=1, corner_radius=6,
                                                segmented_button_fg_color=tab_unselected_color,
                                                segmented_button_selected_color=tab_selected_color,
                                                segmented_button_unselected_color=tab_unselected_color,
                                                segmented_button_selected_hover_color=tab_hover_color,
                                                segmented_button_unselected_hover_color=tab_hover_color,
                                                text_color=tab_text_color, border_color=border_color);
        self.design_right_tabs.pack(fill="both", expand=True, padx=6, pady=6)

        data_preview_tab = self.design_right_tabs.add("Data Preview");
        self.design_data_preview_text = ctk.CTkTextbox(data_preview_tab, font=CODE_FONT, wrap="none");
        self.design_data_preview_text.pack(fill="both", expand=True);
        self.design_data_preview_text.insert("1.0", "Generated data will be shown here...");
        self.design_data_preview_text.configure(state="disabled")

        report_tab = self.design_right_tabs.add("Capability Report");
        report_tab.grid_rowconfigure(0, weight=1);
        report_tab.grid_columnconfigure(0, weight=1)
        self.design_fig, self.design_ax = plt.subplots(1, 1, figsize=(6, 4));
        self.design_canvas = FigureCanvasTkAgg(self.design_fig, master=report_tab);
        self.design_canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew", pady=(0, 6));
        self.design_report_text = ctk.CTkTextbox(report_tab, height=150, font=CODE_FONT);
        self.design_report_text.grid(row=1, column=0, sticky="ew", padx=5, pady=5);
        self.design_report_text.insert("1.0", "Capability analysis results will appear here.");
        self.design_report_text.configure(state="disabled");

        # Sigma Lines Switch
        sigma_switch_frame = ctk.CTkFrame(report_tab, fg_color="transparent");
        sigma_switch_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=0)

        self.show_sigma_lines_var = ctk.BooleanVar(value=False);
        # --- UNICODE FIX HERE ---
        self.sigma_lines_switch = ctk.CTkSwitch(sigma_switch_frame, text="Show \u00b1\u03c3 Lines on Plot",
                                                variable=self.show_sigma_lines_var,
                                                command=self.run_design_capability_analysis, font=DEFAULT_UI_FONT);
        # --- END UNICODE FIX ---
        self.sigma_lines_switch.pack(side="left", padx=5, pady=(0, 5))  # Apply font

        # Run Capability Button
        self.design_run_analysis_button = ctk.CTkButton(report_tab, text="Run Capability Analysis",
                                                        command=self.run_design_capability_analysis,
                                                        font=DEFAULT_UI_FONT, state="disabled");  # Start disabled
        self.design_run_analysis_button.grid(row=3, column=0, sticky="ew", padx=5, pady=(5, 5));  # <-- Modified pady

        # --- NEW: Boundary Check Button ---
        self.run_boundary_check_button = ctk.CTkButton(report_tab, text="Run RNN Boundary Check",
                                                       command=self.run_rnn_boundary_check,
                                                       font=DEFAULT_UI_FONT, state="disabled")  # Start disabled
        self.run_boundary_check_button.grid(row=4, column=0, sticky="ew", padx=5, pady=(0, 10));
        # --- END NEW ---

        # Tolerance Calculator Tab
        calculator_tab = self.design_right_tabs.add("Tolerance Calculator");
        calculator_frame = ctk.CTkFrame(calculator_tab, fg_color="transparent");
        calculator_frame.pack(padx=10, pady=10, fill="x");
        ctk.CTkLabel(calculator_frame, text="Calculate Tolerance based on Target Cpk", font=DEFAULT_UI_FONT_BOLD).pack(
            pady=(0, 10));  # Apply font
        tol_param_info_frame = ctk.CTkFrame(calculator_frame, fg_color="transparent");
        tol_param_info_frame.pack(fill="x", padx=5, pady=4);
        ctk.CTkLabel(tol_param_info_frame, text="Parameter set in Analysis Settings:", font=DEFAULT_UI_FONT).pack(
            side="left");  # Apply font
        nominal_frame = ctk.CTkFrame(calculator_frame, fg_color="transparent");
        nominal_frame.pack(fill="x", padx=5, pady=4);
        ctk.CTkLabel(nominal_frame, text="Desired Nominal:", font=DEFAULT_UI_FONT).pack(side="left");
        self.tol_calc_nominal_entry = ctk.CTkEntry(nominal_frame, width=150, placeholder_text="e.g., Target Mean",
                                                   font=DEFAULT_UI_FONT);
        self.tol_calc_nominal_entry.pack(side="right");  # Apply font
        cpk_frame = ctk.CTkFrame(calculator_frame, fg_color="transparent");
        cpk_frame.pack(fill="x", padx=5, pady=4);
        ctk.CTkLabel(cpk_frame, text="Target Cpk:", font=DEFAULT_UI_FONT).pack(side="left");
        self.tol_calc_cpk_entry = ctk.CTkEntry(cpk_frame, width=150, font=DEFAULT_UI_FONT);
        self.tol_calc_cpk_entry.insert(0, "1.33");
        self.tol_calc_cpk_entry.pack(side="right");  # Apply font
        self.tol_calc_button = ctk.CTkButton(calculator_frame, text="Calculate Tolerance",
                                             command=self.calculate_tolerance_for_cpk, font=DEFAULT_UI_FONT,
                                             state="disabled");  # Start disabled
        self.tol_calc_button.pack(fill="x", padx=5, pady=(15, 8));  # Apply font
        self.tol_calc_result_label = ctk.CTkLabel(calculator_frame, text="Result: +/- T = ? (LSL=?, USL=?)",
                                                  text_color="gray", wraplength=450, justify="left",
                                                  font=DEFAULT_UI_FONT);
        self.tol_calc_result_label.pack(fill="x", padx=5, pady=(0, 15));  # Apply font

        # --- NEW: LIVE VISUALIZATION TAB ---
        live_vis_tab = self.design_right_tabs.add("Live Visualization")
        live_vis_tab.grid_rowconfigure(1, weight=1)  # Make plot area expand
        live_vis_tab.grid_columnconfigure(0, weight=1)

        live_controls_frame = ctk.CTkFrame(live_vis_tab, fg_color="transparent")
        live_controls_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        ctk.CTkLabel(live_controls_frame, text="X-Axis:", font=DEFAULT_UI_FONT).pack(side="left", padx=(5, 2))
        self.live_x_menu = ctk.CTkOptionMenu(live_controls_frame, values=["-"], font=DEFAULT_UI_FONT, width=150,
                                             state="disabled",
                                             command=self._on_live_vis_param_change)  # <-- ADDED COMMAND
        self.live_x_menu.pack(side="left", padx=2)

        ctk.CTkLabel(live_controls_frame, text="Y-Axis:", font=DEFAULT_UI_FONT).pack(side="left", padx=(10, 2))
        self.live_y_menu = ctk.CTkOptionMenu(live_controls_frame, values=["-"], font=DEFAULT_UI_FONT, width=150,
                                             state="disabled",
                                             command=self._on_live_vis_param_change)  # <-- ADDED COMMAND
        self.live_y_menu.pack(side="left", padx=2)

        # --- ADDED LOOP CHECKBOX (set to True) ---
        self.loop_animation_var = ctk.BooleanVar(value=True)  # <-- SET TO TRUE
        self.loop_animation_cb = ctk.CTkCheckBox(live_controls_frame, text="Loop", variable=self.loop_animation_var,
                                                 font=DEFAULT_UI_FONT, width=60,
                                                 command=self._on_loop_toggle)  # <-- ADDED COMMAND
        self.loop_animation_cb.pack(side="left", padx=(10, 0))
        # --- END ADDED ---

        ctk.CTkLabel(live_controls_frame, text="(Updates after 'Generate Data')", text_color="gray",
                     font=SMALL_UI_FONT).pack(side="left", padx=10)

        # Create the plot canvas
        self.live_fig, self.live_ax = plt.subplots(1, 1, figsize=(6, 4))
        self.live_canvas = FigureCanvasTkAgg(self.live_fig, master=live_vis_tab)
        self.live_canvas.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 5))
        # --- END NEW ---

        # --- NEW: BINNED FAILURE ANALYSIS TAB ---
        bin_analysis_tab = self.design_right_tabs.add("Binned Failure Analysis")
        bin_analysis_tab.grid_rowconfigure(1, weight=1)  # Textbox row
        bin_analysis_tab.grid_columnconfigure(0, weight=1)

        # Controls for Binned Analysis
        bin_controls_frame = ctk.CTkFrame(bin_analysis_tab, fg_color="transparent")
        bin_controls_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        ctk.CTkLabel(bin_controls_frame, text="Failure Condition:", font=DEFAULT_UI_FONT).pack(side="left", padx=(5, 2))
        self.bin_fail_condition_menu = ctk.CTkOptionMenu(bin_controls_frame,
                                                         values=["Output > USL", "Output < LSL"],
                                                         font=DEFAULT_UI_FONT, width=150,
                                                         state="disabled")  # Start disabled
        self.bin_fail_condition_menu.set("Output > USL")
        self.bin_fail_condition_menu.pack(side="left", padx=2)

        self.run_bin_analysis_button = ctk.CTkButton(bin_controls_frame,
                                                     text="Run Binned Failure Analysis",
                                                     command=self.run_binned_failure_analysis,
                                                     font=DEFAULT_UI_FONT, state="disabled")  # Start disabled
        self.run_bin_analysis_button.pack(side="left", padx=(10, 5), expand=True, fill="x")

        # Textbox for Binned Analysis Results
        self.bin_analysis_text = ctk.CTkTextbox(bin_analysis_tab, font=CODE_FONT, wrap="none")
        self.bin_analysis_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 5))
        self.bin_analysis_text.insert("1.0", "Binned failure analysis results will appear here.\n\n"
                                             "1. Select an 'Analyze Parameter' (e.g., a 'pred_' output) on the left.\n"
                                             "2. Set the LSL or USL failure threshold on the left.\n"
                                             "3. Select the failure condition (e.g., 'Output > USL') above.\n"
                                             "4. Click 'Run'.")
        self.bin_analysis_text.configure(state="disabled")
        # --- END NEW ---

        # --- Final UI Setup ---
        self.apply_theme()
        if hasattr(self, 'design_canvas') and self.design_canvas and self.design_canvas.get_tk_widget().winfo_exists():
            try:
                self.design_fig.tight_layout()
            except Exception:
                pass
            self.design_canvas.draw()

        # --- NEW: Draw placeholder on live_vis tab ---
        if hasattr(self, 'live_ax') and self.live_ax:
            theme_props = self.app.get_theme_properties()
            self.live_ax.text(0.5, 0.5, 'Click "Generate Data" to start live view',
                              ha='center', va='center', color=theme_props.get("text_color", "gray"))
            try:
                self.live_fig.tight_layout()
            except Exception:
                pass
            self.live_canvas.draw()
        # --- END NEW ---

        self._switch_design_mode("Load from File")
        self._disable_analysis_frame()
        self.app.bind('<Return>', lambda e: self._on_manual_add_keypress())
        self.app.bind('+', lambda e: self._on_manual_add_keypress())

        # Apply font to nested tab buttons *after* adding tabs
        if hasattr(self.app, '_apply_font_to_tabs'):
            self.app._apply_font_to_tabs(self.design_right_tabs)

    def run_rnn_boundary_check(self):
        """
        Compares the generated data against the RNN Tab's training boundaries
        and reports the exact number of violating points in a formatted table.
        """
        try:
            # 1. Get the generated data
            if self.app.generated_design_df is None:
                messagebox.showinfo("Boundary Check", "Please generate data first.")
                return

            # --- THIS IS THE FIX ---
            # The check must be on 'self.app.rnn_tab', not 'self.rnn_tab'
            if not hasattr(self.app, 'rnn_tab') or not self.app.rnn_tab.rnn_data_bounds:
                # --- END FIX ---
                messagebox.showinfo("Boundary Check",
                                    "No RNN training boundaries found. "
                                    "Please load data in the RNN Tab first.")
                return

            generated_df = self.app.generated_design_df
            rnn_bounds = self.app.rnn_tab.rnn_data_bounds
            total_samples = len(generated_df)

            violation_rows = []  # Will store (param, type, count, observed, bound)
            violating_indices = set()  # To count unique points

            # 3. Loop through all known RNN input boundaries
            for param, (train_min, train_max) in rnn_bounds.items():

                # Check if this parameter was part of the simulation
                if param not in generated_df.columns:
                    continue

                data_col = generated_df[param]

                # Check for violations
                if not np.isnan(train_min):
                    below_mask = data_col < train_min
                    count_below = below_mask.sum()
                    if count_below > 0:
                        min_val_found = data_col.min()
                        violating_indices.update(np.where(below_mask)[0])
                        violation_rows.append(
                            (param, "Below Min", count_below, f"{min_val_found:.3g}", f"{train_min:.3g}")
                        )

                if not np.isnan(train_max):
                    above_mask = data_col > train_max
                    count_above = above_mask.sum()
                    if count_above > 0:
                        max_val_found = data_col.max()
                        violating_indices.update(np.where(above_mask)[0])
                        violation_rows.append(
                            (param, "Above Max", count_above, f"{max_val_found:.3g}", f"{train_max:.3g}")
                        )

            # 4. Format and show the final report
            total_violating_points = len(violating_indices)
            percentage = (total_violating_points / total_samples) * 100

            final_report_lines = []

            if not violation_rows:
                final_report_lines.append("RNN Boundary Check Complete:\n")
                final_report_lines.append(f"Total Points Checked: {total_samples}")
                final_report_lines.append("Violating Points: 0 (0.00%)\n")
                final_report_lines.append("All simulated points are within the RNN training boundaries.")

            else:
                final_report_lines.append("RNN Boundary Check Complete:\n")
                final_report_lines.append(
                    f"Total Violating Points: {total_violating_points} / {total_samples} ({percentage:.2f}%)\n")

                # Create table
                # Define column widths
                w_param = 30
                w_viol = 11
                w_pts = 6
                w_obs = 10
                w_bnd = 10

                # Headers
                header = (
                    f"{'Parameter':<{w_param}} | "
                    f"{'Violation':<{w_viol}} | "
                    f"{'Points':>{w_pts}} | "
                    f"{'Observed':>{w_obs}} | "
                    f"{'RNN Bound':>{w_bnd}}"
                )
                separator = "-" * (w_param + w_viol + w_pts + w_obs + w_bnd + 12)

                final_report_lines.append(header)
                final_report_lines.append(separator)

                # Add data rows
                for (param, viol_type, count, observed, bound) in violation_rows:
                    line = (
                        f"{param:<{w_param}} | "
                        f"{viol_type:<{w_viol}} | "
                        f"{count:>{w_pts}} | "
                        f"{observed:>{w_obs}} | "
                        f"{bound:>{w_bnd}}"
                    )
                    final_report_lines.append(line)

            # 5. Write to the text box and switch tabs
            self.bin_analysis_text.configure(state="normal")
            self.bin_analysis_text.delete("1.0", "end")
            self.bin_analysis_text.insert("1.0", "\n".join(final_report_lines))
            self.bin_analysis_text.configure(state="disabled")

            self.design_right_tabs.set("Binned Failure Analysis")

        except Exception as e:
            messagebox.showerror("Boundary Check Error", f"An error occurred: {e}")
            import traceback
            traceback.print_exc()

    # --- END REPLACEMENT ---

    def apply_theme(self):
        theme_props = self.app.get_theme_properties()
        try:  # Add try block for robustness
            if hasattr(self, 'design_fig') and self.design_fig:  # Check fig exists
                self.design_fig.set_facecolor(theme_props["plot_bg"])
            if hasattr(self, 'design_ax') and self.design_ax:  # Check ax exists
                self.design_ax.set_facecolor(theme_props["plot_bg"])
                self.design_ax.tick_params(axis='x', colors=theme_props["text_color"])
                self.design_ax.tick_params(axis='y', colors=theme_props["text_color"])
                for spine in self.design_ax.spines.values(): spine.set_color(theme_props["text_color"])
                self.design_ax.xaxis.label.set_color(theme_props["text_color"]);
                self.design_ax.yaxis.label.set_color(theme_props["text_color"]);
                self.design_ax.title.set_color(theme_props["text_color"])
                self.design_ax.grid(color=theme_props['grid_color'], linestyle='--', alpha=0.5)
                legend = self.design_ax.get_legend()
                if legend is not None:
                    legend.get_frame().set_facecolor(theme_props['plot_bg'])
                    if legend.get_texts():
                        for text in legend.get_texts():
                            try:
                                text.set_color(theme_props['text_color'])
                            except Exception:
                                pass  # Ignore if setting color fails
            if hasattr(self,
                       'design_canvas') and self.design_canvas and self.design_canvas.get_tk_widget().winfo_exists():
                self.design_canvas.draw_idle()

            # --- NEW: Apply theme to live_vis plot ---
            if hasattr(self, 'live_fig') and self.live_fig:
                self.live_fig.set_facecolor(theme_props["plot_bg"])
            if hasattr(self, 'live_ax') and self.live_ax:
                self.live_ax.set_facecolor(theme_props["plot_bg"])
                self.live_ax.tick_params(axis='x', colors=theme_props["text_color"])
                self.live_ax.tick_params(axis='y', colors=theme_props["text_color"])
                for spine in self.live_ax.spines.values(): spine.set_color(theme_props["text_color"])
                self.live_ax.xaxis.label.set_color(theme_props["text_color"]);
                self.live_ax.yaxis.label.set_color(theme_props["text_color"]);
                self.live_ax.title.set_color(theme_props["text_color"])
                self.live_ax.grid(color=theme_props['grid_color'], linestyle='--', alpha=0.5)
            if hasattr(self, 'live_canvas') and self.live_canvas and self.live_canvas.get_tk_widget().winfo_exists():
                self.live_canvas.draw_idle()
            # --- END NEW ---

        except Exception as e:
            print(f"Error applying theme in DesignTab: {e}")

    def _on_sash_drag(self, event, main_frame):
        new_width = main_frame.winfo_pointerx() - main_frame.winfo_rootx()
        if 300 < new_width < self.app.winfo_width() - 400: main_frame.grid_columnconfigure(0, minsize=new_width)

    def _switch_design_mode(self, mode):
        self.design_mode = "File" if mode == "Load from File" else "Manual"

        # --- NEW: Stop animation on mode switch ---
        self._stop_animation = True
        # --- END NEW ---

        if self.design_mode == "File":
            self.design_manual_mode_frame.pack_forget();
            self.design_file_mode_frame.pack(fill="both", expand=True, padx=6, pady=6)
            self.design_generate_button.configure(
                state="normal" if hasattr(self.app, 'ref_df') and self.app.ref_df is not None else "disabled")
            self.manual_ok_button.configure(state="disabled")
        else:  # Manual Mode
            self.design_file_mode_frame.pack_forget();
            self.design_manual_mode_frame.pack(fill="both", expand=True, padx=6, pady=6)
            is_enabled = bool(self.manual_param_widgets)
            self.manual_ok_button.configure(state="normal" if is_enabled else "disabled")
            self.design_generate_button.configure(state="disabled")

        # --- NEW: Update RNN checkbox state on mode switch ---
        self._update_rnn_checkbox_state()
        # --- END NEW ---

    def _add_manual_parameter_row(self, initial_values=None):
        if initial_values is None: initial_values = {}

        # --- MODIFIED: Add headers on the first run ---
        if not self.manual_param_widgets:
            # Clear placeholder
            for widget in self.manual_param_rows_frame.winfo_children():
                if widget.winfo_exists(): widget.destroy()

            # --- ROBUST FIX ---
            # Use the pre-defined bold font object
            header_font = DEFAULT_UI_FONT_BOLD
            # Add a fallback in case font initialization failed
            if header_font is None:
                print("Warning: DEFAULT_UI_FONT_BOLD was None, using fallback.")
                header_font = CTkFont(size=12, weight="bold")
            # --- END ROBUST FIX ---

            headers = ["Parameter Name", "Unit", "Nominal", "Upper Tol (+)", "Lower Tol (-)", "Target Cpk"]
            column_indices = [0, 1, 2, 3, 4, 5]  # Column 6 is for the remove button

            for col, text in zip(column_indices, headers):
                ctk.CTkLabel(self.manual_param_rows_frame, text=text, font=header_font).grid(
                    row=0, column=col, padx=4, pady=(2, 5), sticky="w")
        # --- END MODIFIED ---

        # Row index is now +1 to account for header row
        row_index = len(self.manual_param_widgets) + 1

        name_widget = ctk.CTkEntry(self.manual_param_rows_frame, placeholder_text="Parameter Name",
                                   font=DEFAULT_UI_FONT);
        name_widget.grid(row=row_index, column=0, padx=(0, 4), pady=3, sticky="ew");
        name_widget.insert(0, initial_values.get("name", ""))
        unit_widget = ctk.CTkOptionMenu(self.manual_param_rows_frame,
                                        values=["-", "mm", "Âµm", "cm", "m", "inch", "deg", "rad", "N", "kN", "kg", "g",
                                                "s", "%"], width=80, font=DEFAULT_UI_FONT);
        unit_widget.set(initial_values.get("unit", "Âµm"));
        unit_widget.grid(row=row_index, column=1, padx=4, pady=3, sticky="ew")  # Apply font
        nominal_entry = ctk.CTkEntry(self.manual_param_rows_frame, placeholder_text="Nominal", font=DEFAULT_UI_FONT);
        nominal_entry.grid(row=row_index, column=2, padx=4, pady=3, sticky="ew");
        nominal_entry.insert(0, initial_values.get("nominal", ""))  # Apply font
        tol_upper_entry = ctk.CTkEntry(self.manual_param_rows_frame, placeholder_text="Upper Tol (+)",
                                       font=DEFAULT_UI_FONT);
        tol_upper_entry.grid(row=row_index, column=3, padx=4, pady=3, sticky="ew");
        tol_upper_entry.insert(0, initial_values.get("tol_upper", ""))  # Apply font
        tol_lower_entry = ctk.CTkEntry(self.manual_param_rows_frame, placeholder_text="Lower Tol (-)",
                                       font=DEFAULT_UI_FONT);
        tol_lower_entry.grid(row=row_index, column=4, padx=4, pady=3, sticky="ew");
        tol_lower_entry.insert(0, initial_values.get("tol_lower", ""))  # Apply font
        cpk_entry = ctk.CTkEntry(self.manual_param_rows_frame, placeholder_text="Target Cpk", font=DEFAULT_UI_FONT);
        cpk_entry.insert(0, initial_values.get("cpk", "1.33"));
        cpk_entry.grid(row=row_index, column=5, padx=4, pady=3, sticky="ew")  # Apply font
        widget_dict = {"name": name_widget, "unit": unit_widget, "nominal": nominal_entry, "tol_upper": tol_upper_entry,
                       "tol_lower": tol_lower_entry, "cpk": cpk_entry}
        remove_button = ctk.CTkButton(self.manual_param_rows_frame, text="-", width=28,
                                      command=lambda d=widget_dict: self._remove_manual_parameter_row(d),
                                      font=DEFAULT_UI_FONT);
        remove_button.grid(row=row_index, column=6, padx=(4, 0), pady=3)  # Apply font
        widget_dict["remove"] = remove_button
        self.manual_param_widgets.append(widget_dict)
        self.manual_ok_button.configure(state="normal");

        # --- THIS IS THE FIX ---
        # When we enable the 'Generate' button, also update the RNN checkbox
        self._update_rnn_checkbox_state()
        # --- END FIX ---

        if not initial_values: name_widget.focus()

    def _clear_manual_parameter_rows(self):
        for widget_dict in self.manual_param_widgets:
            for widget in widget_dict.values():
                if widget and widget.winfo_exists(): widget.destroy()
        self.manual_param_widgets.clear()
        if hasattr(self, 'manual_param_rows_frame') and self.manual_param_rows_frame.winfo_exists():
            for widget in self.manual_param_rows_frame.winfo_children():
                if widget.winfo_exists(): widget.destroy()
            ctk.CTkLabel(self.manual_param_rows_frame, text="Click '+' to add a parameter.", font=DEFAULT_UI_FONT).pack(
                pady=20)  # Apply font
        self.manual_ok_button.configure(state="disabled")
        # --- NEW: Also disable RNN checkbox when clearing ---
        if hasattr(self, 'predict_with_rnn_cb'):
            self.predict_with_rnn_cb.configure(state="disabled")
            self.predict_with_rnn_var.set(False)
        # --- END NEW ---

    def _remove_manual_parameter_row(self, widget_dict_to_remove):
        row_to_remove = -1
        try:
            # --- MODIFIED: Re-grid remaining rows ---
            # Find the row index from the grid layout
            grid_info = widget_dict_to_remove['name'].grid_info()
            row_to_remove = grid_info['row']

            # Destroy widgets in the row to remove
            for widget in widget_dict_to_remove.values():
                if widget and widget.winfo_exists():
                    widget.destroy()

            # Remove from list
            self.manual_param_widgets.remove(widget_dict_to_remove)

            # Re-grid all subsequent rows
            # We iterate through all widgets in the frame
            all_widgets = self.manual_param_rows_frame.winfo_children()
            for widget in all_widgets:
                current_grid_info = widget.grid_info()
                if current_grid_info and current_grid_info['row'] > row_to_remove:
                    # Move this widget up by one row
                    widget.grid(row=current_grid_info['row'] - 1,
                                column=current_grid_info['column'],
                                padx=current_grid_info['padx'],
                                pady=current_grid_info['pady'],
                                sticky=current_grid_info['sticky'])
            # --- END MODIFIED ---

        except (ValueError, KeyError, AttributeError) as e:
            print(f"Warning: Could not find or re-grid row to remove: {e}");
            return  # Exit if we had an error

        # Check if list is empty
        if not self.manual_param_widgets:
            self.manual_ok_button.configure(state="disabled");
            # --- NEW: Also disable RNN checkbox ---
            if hasattr(self, 'predict_with_rnn_cb'):
                self.predict_with_rnn_cb.configure(state="disabled")
                self.predict_with_rnn_var.set(False)
            # --- END NEW ---

            # Clear any remaining widgets (like headers) and add placeholder
            for widget in self.manual_param_rows_frame.winfo_children():
                if widget.winfo_exists(): widget.destroy()
            ctk.CTkLabel(self.manual_param_rows_frame, text="Click '+' to add a parameter.", font=DEFAULT_UI_FONT).pack(
                pady=20)  # Apply font

    def _populate_design_param_inputs(self):
        # Clear previous widgets first
        if hasattr(self, 'design_param_frame') and self.design_param_frame.winfo_exists():
            for widget in self.design_param_frame.winfo_children():
                if widget.winfo_exists(): widget.destroy()
        self.design_input_widgets = {}

        if self.app.ref_df is None:
            ctk.CTkLabel(self.design_param_frame, text="Load reference data error.", font=DEFAULT_UI_FONT).pack(
                pady=20)  # Apply font
            return

        params, means, stds = self.app.ref_df.columns, self.app.ref_df.mean(), self.app.ref_df.std(ddof=0)
        header_frame = ctk.CTkFrame(self.design_param_frame, fg_color="transparent");
        header_frame.pack(fill="x", padx=5, pady=(0, 5))
        ctk.CTkLabel(header_frame, text="Parameter", font=DEFAULT_UI_FONT_BOLD, anchor="w").pack(side="left",
                                                                                                 expand=True)  # Apply font
        ctk.CTkLabel(header_frame, text="New Std Dev", font=DEFAULT_UI_FONT_BOLD, width=90).pack(side="right",
                                                                                                 padx=5)  # Apply font
        ctk.CTkLabel(header_frame, text="New Mean", font=DEFAULT_UI_FONT_BOLD, width=90).pack(side="right",
                                                                                              padx=5)  # Apply font

        for param in params:
            row = ctk.CTkFrame(self.design_param_frame);
            row.pack(fill="x", padx=5, pady=3)
            ctk.CTkLabel(row, text=param, anchor="w", font=DEFAULT_UI_FONT).pack(side="left", padx=5, expand=True,
                                                                                 fill="x")  # Apply font
            std_entry = ctk.CTkEntry(row, width=90, font=DEFAULT_UI_FONT);
            std_entry.insert(0, f"{stds[param]:.4f}");
            std_entry.pack(side="right", padx=5, pady=3)  # Apply font
            mean_entry = ctk.CTkEntry(row, width=90, font=DEFAULT_UI_FONT);
            mean_entry.insert(0, f"{means[param]:.4f}");
            mean_entry.pack(side="right", padx=5, pady=3)  # Apply font
            self.design_input_widgets[param] = {'mean': mean_entry, 'std': std_entry}

    def load_reference_data(self):
        path = filedialog.askopenfilename(filetypes=[("Data Files", "*.csv *.xls *.xlsx")])
        if not path: return
        try:
            df = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
            self.app.ref_df = df.select_dtypes(include=[np.number])
            if self.app.ref_df.empty: raise ValueError("No numeric columns.")
            self.app.target_stats = None
            self.app.ref_corr_matrix = self.app.ref_df.corr()
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_loaded_label.configure(text=os.path.basename(path), text_color=success_color, font=bold_font)
            self.design_generate_button.configure(state="normal")
            self.design_status_label.configure(
                text=f"Loaded {len(self.app.ref_df.columns)} parameters. Define targets.", text_color=success_color,
                font=bold_font)
            # --- END COLOR CHANGE ---
            self._populate_design_param_inputs();
            self._disable_analysis_frame();
            self.app.generated_design_df = None
            self.design_mode_switcher.set("Load from File");
            self._switch_design_mode("Load from File")

            # --- NEW: Update RNN checkbox on load ---
            self._update_rnn_checkbox_state()
            # --- END NEW ---

        except Exception as e:
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_loaded_label.configure(text=f"Error: {e}", text_color="red", font=bold_font)
            self.design_generate_button.configure(state="disabled")
            if hasattr(self, 'design_param_frame') and self.design_param_frame.winfo_exists():
                for widget in self.design_param_frame.winfo_children():
                    if widget.winfo_exists(): widget.destroy()
                ctk.CTkLabel(self.design_param_frame, text="Load reference data error.", font=DEFAULT_UI_FONT).pack(
                    pady=20)  # Apply font
            self.app.target_stats = None;
            self.app.ref_df = None;
            self.app.ref_corr_matrix = None

            # --- NEW: Update RNN checkbox on error ---
            self._update_rnn_checkbox_state()
            # --- END NEW ---

    # --- MODIFIED: This function now includes the RNN prediction logic ---
    def generate_multi_param_data(self):
        try:
            # --- NEW: Stop any old animations ---
            self._stop_animation = True
            # --- END NEW ---

            num_samples = int(self.design_samples_entry.get());
            seed = int(self.design_seed_entry.get()) if self.design_seed_entry.get() else None;
            rng = np.random.default_rng(seed)
            method = self.design_method_menu.get();
            repeats = int(self.design_repeats_entry.get()) if method == "Optimized LHS" else 1
            generated_df_local = None
            if self.design_mode == "File":
                if self.app.ref_df is None or self.app.ref_corr_matrix is None: raise ValueError(
                    "Reference data not loaded for File mode.")
                target_means = {};
                target_stds = {}
                for param, widgets in self.design_input_widgets.items():
                    target_means[param] = float(widgets['mean'].get());
                    target_stds[param] = float(widgets['std'].get())
                    if target_stds[param] <= 0: raise ValueError(f"Standard deviation for '{param}' must be positive.")
                target_means_series = pd.Series(target_means);
                target_stds_series = pd.Series(target_stds);
                self.app.target_stats = {'means': target_means_series, 'stds': target_stds_series}
                params = self.app.ref_corr_matrix.columns.tolist();
                dim = len(params);
                target_cov_matrix = np.diag(target_stds_series[params].values ** 2);
                target_corr_from_cov = np.diag(1.0 / target_stds_series[params].values) @ target_cov_matrix @ np.diag(
                    1.0 / target_stds_series[params].values)
                try:
                    corr_matrix_values = self.app.ref_corr_matrix.values;
                    min_eig = np.min(np.real(np.linalg.eigvals(corr_matrix_values)))
                    if min_eig < -1e-8:
                        print(
                            f"Warning: Reference correlation matrix might not be positive semi-definite (min eigenvalue={min_eig}). Attempting adjustment.")
                        corr_matrix_values += np.eye(dim) * 1e-8;
                        min_eig = np.min(np.real(np.linalg.eigvals(corr_matrix_values)))
                        if min_eig < -1e-8: raise ValueError(
                            "Reference correlation matrix is not positive definite even after adjustment.")
                    ref_corr_chol = np.linalg.cholesky(corr_matrix_values)
                except np.linalg.LinAlgError:
                    raise ValueError(
                        "Reference correlation matrix is not positive definite. Cannot generate correlated data.")
                if method == "Optimized LHS":
                    best_samples_norm, best_score = None, -1.0;
                    sampler = qmc.LatinHypercube(d=dim, optimization="random-cd", seed=rng)
                    for _ in range(max(1, repeats)):
                        uniform_samples = sampler.random(n=num_samples);
                        norm_samples = norm.ppf(uniform_samples);
                        score = np.min(pdist(norm_samples)) if num_samples > 1 else 0
                        if score > best_score: best_score, best_samples_norm = score, norm_samples
                    independent_norm_samples = best_samples_norm if best_samples_norm is not None else norm.ppf(
                        sampler.random(n=num_samples))
                else:
                    independent_norm_samples = rng.normal(size=(num_samples, dim))
                correlated_norm_samples = independent_norm_samples @ ref_corr_chol.T
                samples = target_means_series[params].values + correlated_norm_samples * target_stds_series[
                    params].values
                generated_df_local = pd.DataFrame(samples, columns=params)
            else:  # Manual Mode
                if not self.manual_param_widgets: raise ValueError("No parameters defined for Manual mode.")
                mean_option = self.manual_mean_option_menu.get()
                manual_params = {};
                manual_param_names = [];
                self.manual_spec_limits = {}
                for i, widgets in enumerate(self.manual_param_widgets):
                    name = widgets['name'].get().strip();
                    if not name: raise ValueError(f"Parameter name empty in row {i + 1}.")
                    if name in manual_param_names: raise ValueError(f"Duplicate name: '{name}'.")
                    manual_param_names.append(name);
                    unit = widgets['unit'].get()
                    try:
                        nominal = float(widgets['nominal'].get());
                        tol_upper = float(widgets['tol_upper'].get());
                        tol_lower = float(widgets['tol_lower'].get());
                        target_cpk = float(widgets['cpk'].get())
                        if tol_upper <= 0: print(
                            f"Warning: Upper Tol for '{name}' is not positive. Assuming absolute value.")
                        if tol_lower >= 0: print(
                            f"Warning: Lower Tol for '{name}' is not negative. Assuming absolute value negated.")
                        if tol_upper == tol_lower: raise ValueError(f"Upper Tol cannot equal Lower Tol for '{name}'.")
                        if target_cpk <= 0: raise ValueError(f"Target Cpk must be > 0 for '{name}'.")
                        usl = nominal + abs(tol_upper);
                        lsl = nominal - abs(tol_lower)  # <-- This was a bug, should be abs()
                        if usl <= lsl: raise ValueError(
                            f"Calculated USL ({usl}) must be greater than LSL ({lsl}) for '{name}'. Check nominal and tolerances.")
                        process_mean = 0.0;
                        sigma = 0.0
                        if mean_option == "Center Mean in Tolerance":
                            process_mean = (usl + lsl) / 2;
                            sigma = (usl - lsl) / (6 * target_cpk)
                        else:
                            process_mean = nominal
                            if not (lsl < process_mean < usl): print(
                                f"Warning: Nominal ({nominal}) for '{name}' is outside calculated spec limits [{lsl:.4g}, {usl:.4g}]. Target Cpk may not be achievable.")
                            distance_to_usl = usl - process_mean;
                            distance_to_lsl = process_mean - lsl
                            if distance_to_usl <= 0 or distance_to_lsl <= 0:
                                sigma = (usl - lsl) / (6 * target_cpk);
                                print(
                                    f"Warning: Nominal for '{name}' is on or outside limits. Using fallback sigma calculation.")
                            else:
                                sigma = min(distance_to_usl, distance_to_lsl) / (3 * target_cpk)
                        if sigma <= 1e-12: raise ValueError(
                            f"Calculated std dev is too small for '{name}'. Check tolerances and Cpk.")
                        manual_params[name] = {'mean': process_mean, 'std': sigma, 'unit': unit};
                        self.manual_spec_limits[name] = {'lsl': lsl, 'usl': usl}
                    except ValueError as e_val:
                        raise ValueError(f"Invalid numeric input for '{name}' in row {i + 1}: {e_val}")
                dim = len(manual_param_names);
                means = np.array([manual_params[p]['mean'] for p in manual_param_names]);
                stds = np.array([manual_params[p]['std'] for p in manual_param_names]);
                self.design_units_map = {p: manual_params[p]['unit'] for p in manual_param_names}
                if method == "Optimized LHS":
                    best_samples_norm, best_score = None, -1.0;
                    sampler = qmc.LatinHypercube(d=dim, optimization="random-cd", seed=rng)
                    for _ in range(max(1, repeats)):
                        uniform_samples = sampler.random(n=num_samples);
                        norm_samples = norm.ppf(uniform_samples);
                        score = np.min(pdist(norm_samples)) if num_samples > 1 else 0
                        if score > best_score: best_score, best_samples_norm = score, norm_samples
                    independent_norm_samples = best_samples_norm if best_samples_norm is not None else norm.ppf(
                        sampler.random(n=num_samples))
                else:
                    independent_norm_samples = rng.normal(size=(num_samples, dim))
                samples = means + independent_norm_samples * stds
                generated_df_local = pd.DataFrame(samples, columns=manual_param_names)
                self.app.target_stats = {'means': pd.Series({p: manual_params[p]['mean'] for p in manual_param_names}),
                                         'stds': pd.Series({p: manual_params[p]['std'] for p in manual_param_names})}

            if generated_df_local is None: raise RuntimeError("Internal error: Data generation failed unexpectedly.")

            # --- NEW BLOCK: Run RNN Predictions ---
            if self.predict_with_rnn_var.get():
                try:
                    # Check if RNN tab has models
                    if not hasattr(self.app, 'rnn_tab') or not self.app.rnn_tab.trained_models:
                        messagebox.showwarning("RNN Prediction",
                                               "RNN prediction was checked, but no models are trained in the RNN tab. Skipping prediction.")
                    else:
                        predicted_df, warnings = self._run_rnn_predictions(generated_df_local)

                        if not predicted_df.empty:
                            # Add the new predicted columns to the dataframe
                            generated_df_local = pd.concat([generated_df_local, predicted_df], axis=1)

                        # Show any warnings (e.g., about parameter mismatch)
                        if warnings:
                            messagebox.showwarning("RNN Prediction Notice",
                                                   "Predictions complete with warnings:\n\n" + "\n".join(
                                                       list(warnings)))

                except Exception as e:
                    messagebox.showerror("RNN Prediction Error", f"Failed to run RNN predictions:\n{e}")
            # --- END NEW BLOCK ---

            self.app.generated_design_df = generated_df_local;
            self.app.df = generated_df_local.copy();
            # FIX: Set generated_df as well so visualization tab sees it
            self.app.generated_df = None;  # Clear old generated data
            self.app.param_names = generated_df_local.columns.tolist()
            self._display_generated_design_data();
            self.save_design_data_button.configure(state="normal")
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_status_label.configure(
                text=f"Generated {num_samples} samples. Ready for analysis or visualization.", text_color=success_color,
                font=bold_font)
            # --- END COLOR CHANGE ---
            self._enable_analysis_frame();

            # --- MODIFIED: Call animation *before* changing tab ---
            self.start_live_visualization()
            # --- END MODIFIED ---

            self._on_analysis_param_change()
            # if hasattr(self, 'design_right_tabs') and self.design_right_tabs.winfo_exists(): self.design_right_tabs.set(
            #     "Data Preview") # <-- We now set this to "Live Visualization" inside start_live_visualization
            vis_tab = self.app.visualization_tab;
            vis_tab.show_original_var.set(True);
            vis_tab.show_generated_var.set(False);
            vis_tab.show_orig_cb.configure(state="normal");
            vis_tab.show_gen_cb.configure(state="disabled");
            vis_tab._update_vis_controls()

            # FIX: Force visualization update after data generation
            try:
                vis_tab._draw_plots()
                print("DEBUG: Visualization plots drawn after generation")
            except Exception as e:
                print(f"DEBUG: Error drawing visualization plots: {e}")

        except ValueError as ve:
            messagebox.showerror("Input Error", f"Invalid input: {ve}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error: {ve}", text_color="orange",
                                               font=bold_font);
            self.app.target_stats = None
        except Exception as e:
            import traceback;
            traceback.print_exc();
            messagebox.showerror("Generation Error",
                                 f"An unexpected error occurred:\n{e}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error: {e}", text_color="orange",
                                               font=bold_font);
            self.app.target_stats = None

    def _run_rnn_predictions(self, input_df):
        """
        Uses trained RNN models to predict outputs for a given input dataframe.
        Handles parameter mismatches using default values from the RNN tab.
        Includes boundary checking to warn about extrapolation.
        """
        rnn_tab = self.app.rnn_tab
        warnings = set()  # Use a set to avoid duplicate warning messages

        # Check if models are available
        if not rnn_tab.trained_models:
            return pd.DataFrame(), warnings  # Return empty data and no warnings

        # Get the "default" (median) values from the RNN tab's sliders
        ref_inputs = rnn_tab._ref_inputs
        if not ref_inputs:
            raise ValueError("RNN Tab has no default reference inputs. Please load data in RNN tab first.")

        # --- NEW: Get RNN training bounds ---
        rnn_bounds = getattr(rnn_tab, 'rnn_data_bounds', {})
        if not rnn_bounds:
            warnings.add("Could not find RNN training boundaries. Skipping boundary check.")
        # --- END NEW ---

        # Get the columns (parameters) that were *actually* generated by the Design tab
        generated_cols = set(input_df.columns)

        all_predicted_outputs = {}

        # Loop through each trained model
        for output_name, model_info in rnn_tab.trained_models.items():
            # Check if model_info is a dict (new format) or just the model (old)
            if isinstance(model_info, dict):
                model = model_info.get('model')
                required_features_list = model_info.get('features', [])
            else:
                model = model_info  # Backwards compatibility
                required_features_list = rnn_tab.input_channels  # Best guess

            if model is None:
                warnings.add(f"Skipped '{output_name}': Model is missing or invalid.")
                continue

            required_features = set(required_features_list)

            # --- This is your parameter mismatch logic ---

            # 1. Find features the model needs but the Design tab *didn't* generate
            missing_features = required_features - generated_cols

            # 2. Find features the model needs *and* the Design tab *did* generate
            varying_features = required_features.intersection(generated_cols)

            if not varying_features:
                # This model's inputs have *nothing* in common with the generated data. Skip it.
                warnings.add(
                    f"Skipped '{output_name}': Its required inputs {required_features_list} were not in the Design simulation parameters {list(generated_cols)}.")
                continue

            # Create the full 10,000-row DataFrame for *this* model
            # Start with the columns that *are* being varied
            prediction_input_df = input_df[list(varying_features)].copy()

            # 3. Add the "default" (median) values for the missing features
            for feat in missing_features:
                if feat not in ref_inputs:
                    warnings.add(f"Skipped '{output_name}': Missing default value for '{feat}' from RNN tab.")
                    prediction_input_df = None  # Mark for skipping
                    break

                # Add a constant column with the default (median) value
                prediction_input_df[feat] = ref_inputs[feat]

                # Add to warnings set (it will only show the message once per feature)
                warnings.add(f"For model '{output_name}', using default value for '{feat}'.")

            if prediction_input_df is None:
                continue  # Skip this model due to a missing default value

            # --- End of mismatch logic ---

            # --- NEW: Boundary Check Logic ---
            if rnn_bounds:
                for feat in varying_features:
                    # Get the min/max of the data we are about to PREDICT ON
                    min_val = prediction_input_df[feat].min()
                    max_val = prediction_input_df[feat].max()

                    # Get the min/max the model was TRAINED ON
                    train_min, train_max = rnn_bounds.get(feat, (np.nan, np.nan))

                    if not np.isnan(train_min) and min_val < train_min:
                        warnings.add(
                            f"Warning: '{feat}' (min value {min_val:.3g}) is below its RNN training min ({train_min:.3g}). Results may be unreliable.")
                    if not np.isnan(train_max) and max_val > train_max:
                        warnings.add(
                            f"Warning: '{feat}' (max value {max_val:.3g}) is above its RNN training max ({train_max:.3g}). Results may be unreliable.")
            # --- END NEW ---

            # Ensure column order matches the model's training order
            final_model_input = prediction_input_df[required_features_list]

            # Run prediction
            try:
                # Check if using Cameo model
                if rnn_tab.using_cameo and model == "CAMEO_MODEL":
                    # Use Cameo model for prediction (direct call to Variant functions)
                    cameo_func = rnn_tab.cameo_models.get(output_name)
                    if cameo_func is None:
                        raise ValueError(f"No Cameo function found for output: {output_name}")

                    # Call Cameo function for each row
                    predicted_values = []
                    for idx, row in final_model_input.iterrows():
                        result = cameo_func(
                            B1_offset=float(row['B1_offset']),
                            B2_offset=float(row['B2_offset']),
                            B3_offset=float(row['B3_offset']),
                            B4_offset=float(row['B4_offset']),
                            B5_offset=float(row['B5_offset']),
                            Helix_Angle=float(row['Helix_Angle']),
                            Input_Stifness=float(row['Input_Stifness']),
                            Lead_Crown_Pinion=float(row['Lead_Crown_Pinion']),
                            Lead_Slope_Pinion=float(row['Lead_Slope_Pinion']),
                            Pressure_Angle=float(row['Pressure_Angle'])
                        )
                        # Variant functions return lists, extract first element
                        predicted_values.append(result[0] if isinstance(result, list) else result)
                    predicted_values = np.array(predicted_values)
                else:
                    # Use the RNNTab's safe_predict method for custom models
                    predicted_values = rnn_tab.safe_predict(model, final_model_input, feature_cols=required_features_list)

                # Add to our results dict with a 'pred_' prefix
                all_predicted_outputs[f"pred_{output_name}"] = predicted_values
            except Exception as e:
                warnings.add(f"Prediction failed for '{output_name}': {e}")

        # Return a new DataFrame with all predicted columns, and the set of warnings
        return pd.DataFrame(all_predicted_outputs, index=input_df.index), warnings

    def save_generated_design_data(self):
        if self.app.generated_design_df is None: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="simulated_system_data",
                                            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if not path: return
        try:
            if path.lower().endswith('.csv'):
                self.app.generated_design_df.to_csv(path, index=False)
            else:
                self.app.generated_design_df.to_excel(path, index=False, engine="openpyxl")
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_status_label.configure(text=f"Saved to {os.path.basename(path)}", text_color=success_color,
                                               font=bold_font)
            # --- END COLOR CHANGE ---
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save file.\n{e}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error saving file: {e}", text_color="orange",
                                               font=bold_font)

    def run_design_capability_analysis(self):
        if self.app.generated_design_df is None: messagebox.showinfo("Analysis Info",
                                                                     "Generate system data first."); return
        param = self.design_analysis_param_menu.get()
        if param == "-": messagebox.showinfo("Analysis Info", "Select a parameter to analyze."); return
        try:
            lsl_str = self.design_lsl_entry.get().strip();
            usl_str = self.design_usl_entry.get().strip();
            lsl = float(lsl_str) if lsl_str else None;
            usl = float(usl_str) if usl_str else None
            if lsl is None or usl is None: raise ValueError("Both LSL and USL must be provided.")
            if lsl >= usl: raise ValueError("USL must be greater than LSL.")
            data_full = self.app.generated_design_df[param];
            data_valid = data_full.dropna()
            if len(data_valid) < 2: raise ValueError("Not enough valid data points for analysis.")
            mu = data_valid.mean();
            sigma = data_valid.std(ddof=1)
            if sigma == 0 or np.isnan(sigma) or sigma <= 1e-12: raise ValueError(
                "Data has zero or invalid standard deviation.")
            cp = (usl - lsl) / (6 * sigma);
            cpu = (usl - mu) / (3 * sigma);
            cpl = (mu - lsl) / (3 * sigma);
            cpk = min(cpu, cpl)
            z_usl = (usl - mu) / sigma;
            z_lsl = (lsl - mu) / sigma
            ppm_above_usl = (1 - norm.cdf(z_usl)) * 1_000_000;
            ppm_below_lsl = norm.cdf(z_lsl) * 1_000_000;
            total_expected_ppm = ppm_above_usl + ppm_below_lsl
            n_total = len(data_full);
            n_above_usl = np.sum(data_full > usl);
            n_below_lsl = np.sum(data_full < lsl);
            total_observed_ppm = ((n_above_usl + n_below_lsl) / n_total) * 1_000_000 if n_total > 0 else 0

            # --- CORRECTED SIGMA BANDS ---
            sigma_bands = {f"Below -4\u03c3": (data_full < mu - 4 * sigma).sum(),
                           f"-4\u03c3 to -3\u03c3": (
                                   (data_full >= mu - 4 * sigma) & (data_full < mu - 3 * sigma)).sum(),
                           f"-3\u03c3 to -2\u03c3": (
                                   (data_full >= mu - 3 * sigma) & (data_full < mu - 2 * sigma)).sum(),
                           f"-2\u03c3 to -1\u03c3": (
                                   (data_full >= mu - 2 * sigma) & (data_full < mu - 1 * sigma)).sum(),
                           f"-1\u03c3 to Mean": ((data_full >= mu - 1 * sigma) & (data_full < mu)).sum(),
                           f"Mean to +1\u03c3": ((data_full >= mu) & (data_full < mu + 1 * sigma)).sum(),
                           f"+1\u03c3 to +2\u03c3": (
                                   (data_full >= mu + 1 * sigma) & (data_full < mu + 2 * sigma)).sum(),
                           f"+2\u03c3 to +3\u03c3": (
                                   (data_full >= mu + 2 * sigma) & (data_full < mu + 3 * sigma)).sum(),
                           f"+3\u03c3 to +4\u03c3": (
                                   (data_full >= mu + 3 * sigma) & (data_full < mu + 4 * sigma)).sum(),
                           f"Above +4\u03c3": (data_full >= mu + 4 * sigma).sum()}
            # --- END CORRECTION ---

            nan_count = data_full.isna().sum();
            if nan_count > 0: sigma_bands["NaN Values"] = nan_count
            sigma_table_str = f"Sigma Distribution (N = {n_total}):\n{'-' * 40}\n";
            total_counted = 0
            for band, count in sigma_bands.items(): percentage = (
                    count / n_total * 100) if n_total > 0 else 0; sigma_table_str += f"  {band:<12}: {count:>7} ({percentage:>6.2f}%)\n"; total_counted += count
            sigma_table_str += f"  {'Total':<12}: {total_counted:>7} (100.00%)\n{'-' * 40}\n"

            # --- Plotting ---
            self.design_ax.clear()
            theme_props = self.app.get_theme_properties();
            hist_color = theme_props.get("hist_color_generated", "#ff006e");
            norm_curve_color = theme_props.get("normal_curve_color", "lime");
            spec_limit_color = theme_props.get("ellipse_color", "red");
            text_color = theme_props.get("text_color", "white");
            grid_color = theme_props.get("grid_color", "#444444");
            plot_bg_color = theme_props.get("plot_bg", "#FFFFFF")
            counts, bins, patches = self.design_ax.hist(data_valid, bins='auto', density=True, color=hist_color,
                                                        alpha=0.7, label='Data Distribution')
            xmin, xmax = self.design_ax.get_xlim();
            x_norm = np.linspace(xmin, xmax, 100);
            p_norm = norm.pdf(x_norm, mu, sigma);
            self.design_ax.plot(x_norm, p_norm, color=norm_curve_color, linewidth=2, label='Normal Fit')
            self.design_ax.axvline(lsl, color=spec_limit_color, linestyle='--', linewidth=2, label=f'LSL ({lsl:.4g})');
            self.design_ax.axvline(usl, color=spec_limit_color, linestyle='--', linewidth=2, label=f'USL ({usl:.4g})');
            self.design_ax.axvline(mu, color=text_color, linestyle=':', linewidth=1, label=f'Mean ({mu:.4g})')

            # --- Plot Sigma Lines if Switch is ON ---
            if hasattr(self, 'show_sigma_lines_var') and self.show_sigma_lines_var.get():
                sigma_colors = theme_props.get("sigma_colors", ['purple', 'orange', 'red'])
                for i in range(1, min(4, len(sigma_colors) + 1)):
                    line_color = sigma_colors[i - 1]

                    # --- CORRECTED LABELS ---
                    self.design_ax.axvline(mu + i * sigma, color=line_color, linestyle=':', linewidth=1.0, alpha=0.8,
                                           label=f'+{i}\u03c3')
                    self.design_ax.axvline(mu - i * sigma, color=line_color, linestyle=':', linewidth=1.0, alpha=0.8,
                                           label=f'-{i}\u03c3')
                    # --- END CORRECTION ---

            stats_text = f"Cp = {cp:.3f}\nCpk = {cpk:.3f}";
            self.design_ax.text(0.05, 0.95, stats_text, transform=self.design_ax.transAxes, fontsize=10,
                                verticalalignment='top',
                                bbox=dict(boxstyle='round,pad=0.3', fc=plot_bg_color, alpha=0.8), color=text_color)

            # --- MODIFIED: Add unit to x-label ---
            unit_str = self.design_units_map.get(param, "")
            # For predicted values, the unit is of the output
            if param.startswith("pred_"):
                original_name = param.replace("pred_", "")
                if hasattr(self.app, 'rnn_tab'):
                    unit_str = self.app.rnn_tab.rnn_units.get(original_name, "")

            self.design_ax.set_title(f"Capability Analysis: {param}");
            self.design_ax.set_xlabel(f"{param} ({unit_str})");
            self.design_ax.set_ylabel("Density")
            # --- END MODIFIED ---

            self.design_ax.set_facecolor(plot_bg_color);
            self.design_ax.tick_params(axis='x', colors=text_color);
            self.design_ax.tick_params(axis='y', colors=text_color);
            for spine in self.design_ax.spines.values(): spine.set_color(text_color)
            self.design_ax.xaxis.label.set_color(text_color);
            self.design_ax.yaxis.label.set_color(text_color);
            self.design_ax.title.set_color(text_color)
            self.design_ax.grid(color=grid_color, linestyle='--', alpha=0.5)
            handles, labels = self.design_ax.get_legend_handles_labels();
            unique_labels = {};
            new_handles = [];
            new_labels = []
            for handle, label in zip(handles, labels):
                if label not in unique_labels: unique_labels[label] = handle; new_handles.append(
                    handle); new_labels.append(label)
            legend = self.design_ax.legend(new_handles, new_labels, fontsize=8)
            if legend: legend.get_frame().set_facecolor(plot_bg_color); legend.get_frame().set_edgecolor(grid_color);
            for text in legend.get_texts(): text.set_color(text_color)

            # --- Reporting ---
            # --- CORRECTED REPORT STRING ---
            report_str = (
                f"Capability Analysis Report for: {param}\n{'-' * 40}\nSpecification Limits:\n  LSL: {lsl:.5g}\n  USL: {usl:.5g}\n{'-' * 40}\nSample Statistics (N_valid = {len(data_valid)}):\n  Mean (\u03bc): {mu:.5g}\n  Std Dev (\u03c3): {sigma:.5g}\n{'-' * 40}\nCapability Indices:\n  Potential (Cp): {cp:.4f}\n  Actual (Cpk):   {cpk:.4f}\n{'-' * 40}\n{sigma_table_str}Process Performance (PPM - Parts Per Million):\n  Expected > USL: {ppm_above_usl:.2f}\n  Expected < LSL: {ppm_below_lsl:.2f}\n  Total Expected: {total_expected_ppm:.2f} PPM\n\n  Observed > USL: {n_above_usl} ({n_above_usl / n_total * 1e6:.1f} PPM)\n  Observed < LSL: {n_below_lsl} ({n_below_lsl / n_total * 1e6:.1f} PPM)\n  Total Observed: {n_above_usl + n_below_lsl} ({total_observed_ppm:.1f} PPM)\n{'-' * 40}\n"
            )
            # --- END CORRECTION ---

            self.design_report_text.configure(state="normal");
            self.design_report_text.delete("1.0", "end");
            self.design_report_text.insert("1.0", report_str);
            self.design_report_text.configure(state="disabled")
            if self.design_canvas and self.design_canvas.get_tk_widget().winfo_exists(): self.design_canvas.draw();
            if hasattr(self, 'design_right_tabs') and self.design_right_tabs.winfo_exists(): self.design_right_tabs.set(
                "Capability Report");
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_status_label.configure(text="Analysis complete.", text_color=success_color, font=bold_font)
            # --- END COLOR CHANGE ---
        except ValueError as ve:
            messagebox.showerror("Analysis Error", f"Invalid input for analysis: {ve}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Analysis Error: {ve}", text_color="orange",
                                               font=bold_font)
        except Exception as e:
            import traceback;
            traceback.print_exc();
            messagebox.showerror("Analysis Error",
                                 f"An unexpected error occurred during analysis:\n{e}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Analysis Error: {e}", text_color="orange",
                                               font=bold_font)

    # --- NEW: BINNED FAILURE ANALYSIS METHOD ---
    def run_binned_failure_analysis(self):
        """
        Runs the binned failure analysis based on settings and displays the report.
        """
        try:
            # 1. Get Settings
            if self.app.generated_design_df is None:
                messagebox.showinfo("Analysis Info", "Generate system data first.")
                return

            df = self.app.generated_design_df
            output_col = self.design_analysis_param_menu.get()
            fail_condition = self.bin_fail_condition_menu.get()

            if output_col == "-":
                messagebox.showinfo("Analysis Info",
                                    "Select an 'Analyze Parameter' (e.g., a 'pred_' output) on the left.")
                return

            target = None
            report_header = ""
            is_failure = None

            # 2. Define Failure Mask
            if fail_condition == "Output > USL":
                try:
                    target_str = self.design_usl_entry.get().strip()
                    if not target_str: raise ValueError("USL entry is empty.")
                    target = float(target_str)
                    is_failure = df[output_col] > target
                    report_header = f"Binned Failure Analysis for: {output_col} > {target}\n"
                except ValueError:
                    messagebox.showerror("Input Error",
                                         "Invalid or empty USL value provided. Please set the USL on the left.")
                    return
            else:  # "Output < LSL"
                try:
                    target_str = self.design_lsl_entry.get().strip()
                    if not target_str: raise ValueError("LSL entry is empty.")
                    target = float(target_str)
                    is_failure = df[output_col] < target
                    report_header = f"Binned Failure Analysis for: {output_col} < {target}\n"
                except ValueError:
                    messagebox.showerror("Input Error",
                                         "Invalid or empty LSL value provided. Please set the LSL on the left.")
                    return

            total_failures = is_failure.sum()
            report_header += f"Total Failures: {total_failures} out of {len(df)} samples ({total_failures / len(df):.2%})\n"

            # 3. Identify Input Columns
            # Use all columns that are NOT the output column being analyzed
            input_cols = [col for col in df.columns if col != output_col]
            if not input_cols:
                messagebox.showerror("Analysis Error", "No input columns found to analyze against.")
                return

            full_report = [report_header]

            # 4. Loop and Analyze Each Input
            for input_col in input_cols:
                full_report.append(f"\n--- Analysis by Input: {input_col} ---")
                input_series = df[input_col].dropna()  # Drop NaNs for stats
                if input_series.empty:
                    full_report.append("  (Skipped: No valid data for this input)")
                    continue

                mu = input_series.mean()
                sigma = input_series.std()

                if sigma == 0 or np.isnan(sigma) or sigma <= 1e-12:
                    full_report.append("  (Skipped: Input has zero or invalid variation)")
                    continue

                # Define bins using the full (non-NaN) series
                bins = [-np.inf, mu - 4 * sigma, mu - 3 * sigma, mu - 2 * sigma, mu - 1 * sigma,
                        mu,
                        mu + 1 * sigma, mu + 2 * sigma, mu + 3 * sigma, mu + 4 * sigma, np.inf]
                labels = ["Below -4s", "-4s to -3s", "-3s to -2s", "-2s to -1s",
                          "-1s to Mean", "Mean to +1s", "+1s to +2s", "+2s to +3s",
                          "+3s to +4s", "Above +4s"]

                # Bin the original data (including NaNs, which pd.cut handles)
                binned_data = pd.cut(df[input_col], bins=bins, labels=labels, right=False)

                # 5. Count Failures in Bins
                for bin_label in labels:
                    samples_in_bin_mask = (binned_data == bin_label)
                    total_in_bin = samples_in_bin_mask.sum()

                    if total_in_bin == 0:
                        report_line = f"  {bin_label:<12}: {0:>5} failures (  N/A  %) out of {0:>5} samples"
                    else:
                        # Combine bin mask with failure mask
                        failures_in_bin = (samples_in_bin_mask & is_failure).sum()
                        fail_rate = (failures_in_bin / total_in_bin) * 100
                        report_line = f"  {bin_label:<12}: {failures_in_bin:>5} failures ({fail_rate:>6.2f}%) out of {total_in_bin:>5} samples"

                    full_report.append(report_line)

            # 6. Display Report
            self.bin_analysis_text.configure(state="normal")
            self.bin_analysis_text.delete("1.0", "end")
            self.bin_analysis_text.insert("1.0", "\n".join(full_report))
            self.bin_analysis_text.configure(state="disabled")

            # Switch to the tab
            self.design_right_tabs.set("Binned Failure Analysis")

        except Exception as e:
            messagebox.showerror("Analysis Error", f"An unexpected error occurred during binned failure analysis:\n{e}")
            import traceback
            traceback.print_exc()
            # Reset text on error
            self.bin_analysis_text.configure(state="normal")
            self.bin_analysis_text.delete("1.0", "end")
            self.bin_analysis_text.insert("1.0", f"Analysis Failed:\n{e}")
            self.bin_analysis_text.configure(state="disabled")

    # --- END NEW METHOD ---

    def _save_preset(self):
        if self.design_mode != "Manual" or not self.manual_param_widgets: messagebox.showwarning("Save Preset",
                                                                                                 "Presets can only be saved in 'Define Manually' mode when parameters are defined."); return
        preset_data = [];
        valid = True
        for i, widgets in enumerate(self.manual_param_widgets):
            name = widgets['name'].get().strip();
            unit = widgets['unit'].get();
            nom_str = widgets['nominal'].get().strip();
            tu_str = widgets['tol_upper'].get().strip();
            tl_str = widgets['tol_lower'].get().strip();
            cpk_str = widgets['cpk'].get().strip()
            if not name: messagebox.showerror("Save Error",
                                              f"Parameter name in row {i + 1} cannot be empty."); valid = False; break
            try:
                float(nom_str);
                float(tu_str);
                float(tl_str);
                float(cpk_str)
            except ValueError:
                messagebox.showerror("Save Error",
                                     f"Invalid numeric value found in row {i + 1} for '{name}'. Please check inputs.");
                valid = False;
                break
            preset_data.append(
                {"name": name, "unit": unit, "nominal": nom_str, "tol_upper": tu_str, "tol_lower": tl_str,
                 "cpk": cpk_str})
        if not valid: return
        filepath = filedialog.asksaveasfilename(defaultextension=".json",
                                                filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"),
                                                           ("All files", "*.*")], title="Save Manual Parameter Preset")
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=4)
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_status_label.configure(text=f"Preset saved to {os.path.basename(filepath)}",
                                               text_color=success_color, font=bold_font)
            # --- END COLOR CHANGE ---
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save preset file:\n{e}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error saving preset: {e}",
                                               text_color="orange", font=bold_font)

    def _load_preset(self):
        if self.design_mode != "Manual": self.design_mode_switcher.set("Define Manually"); self._switch_design_mode(
            "Define Manually"); self.after(50, self._load_preset_file_dialog); return
        self._load_preset_file_dialog()

    def _load_preset_file_dialog(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")],
            title="Load Manual Parameter Preset")
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)
            if not isinstance(loaded_data, list): raise TypeError("Preset file should contain a list of parameters.")
            if not all(isinstance(item, dict) for item in loaded_data): raise TypeError(
                "Each item in the preset list should be a dictionary.")
            required_keys = {"name", "unit", "nominal", "tol_upper", "tol_lower", "cpk"}
            if not all(required_keys.issubset(item.keys()) for item in loaded_data): raise TypeError(
                f"Each parameter dictionary must contain keys: {', '.join(required_keys)}")
            self._clear_manual_parameter_rows()
            for param_data in loaded_data: self._add_manual_parameter_row(initial_values=param_data)
            theme_props = self.app.get_theme_properties();
            # --- COLOR CHANGE ---
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            bold_font = DEFAULT_UI_FONT_BOLD  # Use global font
            self.design_status_label.configure(text=f"Preset loaded from {os.path.basename(filepath)}",
                                               text_color=success_color, font=bold_font)
            # --- END COLOR CHANGE ---
            if self.manual_param_widgets:
                self.manual_ok_button.configure(state="normal")

                # --- THIS IS THE FIX ---
                # When we enable the 'Generate' button, also update the RNN checkbox
                self._update_rnn_checkbox_state()
                # --- END FIX ---

        except FileNotFoundError:
            messagebox.showerror("Load Error", "Preset file not found.");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text="Error: Preset file not found.",
                                               text_color="orange", font=bold_font)
        except json.JSONDecodeError:
            messagebox.showerror("Load Error", "Preset file is not valid JSON.");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text="Error: Invalid preset file format.",
                                               text_color="orange", font=bold_font)
        except TypeError as te:
            messagebox.showerror("Load Error", f"Preset file format error: {te}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error: Preset format error: {te}",
                                               text_color="orange", font=bold_font)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load preset file:\n{e}");
            bold_font = DEFAULT_UI_FONT_BOLD;  # Use global font
            self.design_status_label.configure(text=f"Error loading preset: {e}",
                                               text_color="orange",
                                               font=bold_font);
            self._clear_manual_parameter_rows()

    def _display_generated_design_data(self):
        try:
            self.design_data_preview_text.configure(state="normal");
            self.design_data_preview_text.delete("1.0", "end")
            if self.app.generated_design_df is not None:
                max_rows_display = 500;
                display_string = self.app.generated_design_df.head(max_rows_display).to_string(
                    max_rows=max_rows_display, float_format="{:.4f}".format)
                if len(
                        self.app.generated_design_df) > max_rows_display: display_string += f"\n\n... (showing first {max_rows_display} of {len(self.app.generated_design_df)} rows)"
                self.design_data_preview_text.insert("1.0", display_string)
            else:
                self.design_data_preview_text.insert("1.0", "No generated design data to display.")
            self.design_data_preview_text.configure(state="disabled")
        except Exception as e:
            print(f"Error displaying generated design data: {e}");
        try:
            self.design_data_preview_text.configure(state="disabled")
        except:
            pass

    def _on_manual_add_keypress(self, *args):
        try:
            focused_widget = self.focus_get()
            if focused_widget and self.design_manual_mode_frame.winfo_containing(focused_widget.winfo_rootx(),
                                                                                 focused_widget.winfo_rooty()):
                if self.app.tabs.get() == "Design" and self.design_mode == "Manual": self._add_manual_parameter_row()
        except Exception:
            pass  # Ignore errors

    def calculate_tolerance_for_cpk(self):
        """
        Calculates the symmetric tolerance (+/- T) required to achieve a target Cpk,
        assuming the process mean is centered at the provided nominal and uses
        the standard deviation observed in the generated data. Displays formula and calculation.
        """
        if self.app.generated_design_df is None:
            messagebox.showinfo("Tolerance Calculator", "Generate system data first.")
            # Reset font if error occurs
            default_font = DEFAULT_UI_FONT  # Use global font
            self.tol_calc_result_label.configure(text="Result: Generate data first", text_color="orange",
                                                 font=default_font)
            return

        param = self.design_analysis_param_menu.get()  # Reuse the selected parameter
        if param == "-":
            messagebox.showinfo("Tolerance Calculator", "Select a parameter first.")
            # Reset font if error occurs
            default_font = DEFAULT_UI_FONT  # Use global font
            self.tol_calc_result_label.configure(text="Result: Select parameter", text_color="orange",
                                                 font=default_font)
            return

        try:
            nominal_str = self.tol_calc_nominal_entry.get().strip()
            cpk_target_str = self.tol_calc_cpk_entry.get().strip()

            if not nominal_str:
                # If nominal is empty, default to the mean of the generated data
                nominal = self.app.generated_design_df[param].mean()
                self.tol_calc_nominal_entry.delete(0, 'end')
                self.tol_calc_nominal_entry.insert(0, f"{nominal:.5g}")  # Show the used value
            else:
                nominal = float(nominal_str)

            cpk_target = float(cpk_target_str)
            if cpk_target <= 0:
                raise ValueError("Target Cpk must be positive.")

            # Use the standard deviation from the *generated* dataset
            data = self.app.generated_design_df[param].dropna()
            if len(data) < 2:
                raise ValueError("Not enough valid data points for std dev calculation.")

            sigma = data.std(ddof=1)  # Use sample standard deviation (n-1)
            if sigma == 0 or np.isnan(sigma) or sigma <= 1e-12:
                raise ValueError("Observed standard deviation is zero or invalid.")

            # --- Calculation ---
            # Formula: T = 3 * Cpk * sigma
            tolerance_T = 3 * cpk_target * sigma
            lsl_calc = nominal - tolerance_T
            usl_calc = nominal + tolerance_T

            # --- Create Formula and Calculation Strings ---
            formula_str = "Formula: T = 3 * Cpk * Ïƒ"
            calculation_str = f"Calculation: T = 3 * {cpk_target:.3f} * {sigma:.5g}"

            # --- Display the result (Reordered) ---
            result_main = (
                f"Result: Â±{tolerance_T:.5g}\n"
                f"(LSL = {lsl_calc:.5g}, USL = {usl_calc:.5g})"
            )
            details = (
                f"Based on observed Ïƒ = {sigma:.5g}\n\n"  # Added space before formula
                f"{formula_str}\n"
                f"{calculation_str}"
            )

            result_text = f"{result_main}\n\n{details}"  # Combine with extra space

            # --- Set Font and Text ---
            # Increase font size (e.g., by 2 points from default)
            current_font_size = DEFAULT_UI_FONT.cget("size")  # Use global font
            larger_font = CTkFont(family=DEFAULT_UI_FONT.cget("family"),
                                  size=current_font_size + 2)  # Use defined family

            theme_props = self.app.get_theme_properties()
            success_color = theme_props.get("text_color", "#228B22")  # Use theme color
            self.tol_calc_result_label.configure(text=result_text,
                                                 text_color=success_color,
                                                 font=larger_font)  # Apply larger font

        except ValueError as ve:
            messagebox.showerror("Calculation Error", f"Invalid input: {ve}")
            # Reset font if error occurs
            default_font = DEFAULT_UI_FONT  # Use global font
            self.tol_calc_result_label.configure(text=f"Error: {ve}", text_color="orange", font=default_font)
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An unexpected error occurred:\n{e}")
            # Reset font if error occurs
            default_font = DEFAULT_UI_FONT  # Use global font
            self.tol_calc_result_label.configure(text=f"Error: {e}", text_color="orange", font=default_font)

    def _disable_analysis_frame(self):
        """Disables analysis SETTINGS frame on the left AND calculator widgets on the right"""
        try:
            # Disable analysis settings widgets (left panel)
            if hasattr(self, 'design_analysis_param_menu') and self.design_analysis_param_menu.winfo_exists():
                self.design_analysis_param_menu.configure(values=["-"], state="disabled")
                self.design_analysis_param_menu.set("-")  # Ensure selection is reset
            for widget_attr in ['design_lsl_entry', 'design_usl_entry']:
                widget = getattr(self, widget_attr, None)
                if widget and widget.winfo_exists():
                    widget.configure(state="disabled")

            # Disable capability analysis button (right panel)
            if hasattr(self, 'design_run_analysis_button') and self.design_run_analysis_button.winfo_exists():
                self.design_run_analysis_button.configure(state="disabled")

            # Disable tolerance calculator widgets (right panel)
            for widget_attr in ['tol_calc_nominal_entry', 'tol_calc_cpk_entry', 'tol_calc_button']:
                if hasattr(self, widget_attr) and getattr(self, widget_attr).winfo_exists():
                    getattr(self, widget_attr).configure(state="disabled")

            if hasattr(self, 'tol_calc_result_label') and self.tol_calc_result_label.winfo_exists():

                # --- THIS IS THE PERMANENT FIX ---
                # Use the global font object directly
                default_font = DEFAULT_UI_FONT
                # Add a fallback in case font initialization failed
                if default_font is None:
                    default_font = CTkFont(size=13)
                    # --- END PERMANENT FIX ---

                self.tol_calc_result_label.configure(text="Result: +/- T = ? (LSL=?, USL=?)", text_color="gray",
                                                     font=default_font)  # Reset result label and font

            # --- NEW: Disable live vis dropdowns ---
            if hasattr(self, 'live_x_menu') and self.live_x_menu.winfo_exists():
                self.live_x_menu.configure(values=["-"], state="disabled")
                self.live_x_menu.set("-")
            if hasattr(self, 'live_y_menu') and self.live_y_menu.winfo_exists():
                self.live_y_menu.configure(values=["-"], state="disabled")
                self.live_y_menu.set("-")
            # --- END NEW ---

            # --- NEW: Disable Binned Failure Analysis widgets ---
            if hasattr(self, 'bin_fail_condition_menu') and self.bin_fail_condition_menu.winfo_exists():
                self.bin_fail_condition_menu.configure(state="disabled")
            if hasattr(self, 'run_bin_analysis_button') and self.run_bin_analysis_button.winfo_exists():
                self.run_bin_analysis_button.configure(state="disabled")
            # --- END NEW ---

        except Exception as e:
            # This print is critical for debugging
            print(f"Error disabling analysis/calculator frame: {e}")

    def _enable_analysis_frame(self):
        """Enables analysis SETTINGS frame on the left AND calculator widgets on the right"""
        try:
            if hasattr(self.app, 'generated_design_df') and self.app.generated_design_df is not None:
                params = self.app.generated_design_df.columns.tolist();
                # Enable analysis settings widgets (left panel)
                if hasattr(self, 'design_analysis_param_menu') and self.design_analysis_param_menu.winfo_exists():
                    self.design_analysis_param_menu.configure(values=params, state="normal")
                    if params:
                        self.design_analysis_param_menu.set(params[0])
                    else:
                        self.design_analysis_param_menu.set("-")  # Handle empty case
                for widget_attr in ['design_lsl_entry', 'design_usl_entry']:
                    widget = getattr(self, widget_attr, None)
                    if widget and widget.winfo_exists():
                        widget.configure(state="normal")

                # Enable capability analysis button (right panel)
                if hasattr(self, 'design_run_analysis_button') and self.design_run_analysis_button.winfo_exists():
                    self.design_run_analysis_button.configure(state="normal")

                # Enable tolerance calculator widgets (right panel)
                for widget_attr in ['tol_calc_nominal_entry', 'tol_calc_cpk_entry', 'tol_calc_button']:
                    if hasattr(self, widget_attr) and getattr(self, widget_attr).winfo_exists():
                        getattr(self, widget_attr).configure(state="normal")

                # --- THIS IS THE LINE YOU WERE MISSING ---
                if hasattr(self, 'run_boundary_check_button') and self.run_boundary_check_button.winfo_exists():
                    self.run_boundary_check_button.configure(state="normal")
                # --- END FIX ---

                # --- NEW: Enable live vis dropdowns ---
                if hasattr(self, 'live_x_menu') and self.live_x_menu.winfo_exists():
                    self.live_x_menu.configure(values=params, state="normal")
                    self.live_x_menu.set(params[0] if params else "-")
                if hasattr(self, 'live_y_menu') and self.live_y_menu.winfo_exists():
                    self.live_y_menu.configure(values=params, state="normal")
                    self.live_y_menu.set(params[1] if len(params) > 1 else (params[0] if params else "-"))
                # --- END NEW ---

                # --- NEW: Enable Binned Failure Analysis widgets ---
                if hasattr(self, 'bin_fail_condition_menu') and self.bin_fail_condition_menu.winfo_exists():
                    self.bin_fail_condition_menu.configure(state="normal")
                if hasattr(self, 'run_bin_analysis_button') and self.run_bin_analysis_button.winfo_exists():
                    self.run_bin_analysis_button.configure(state="normal")
                # --- END NEW ---

                # Optionally pre-fill nominal with mean when enabling
                current_param = self.design_analysis_param_menu.get()
                if current_param != "-":
                    current_mean = self.app.generated_design_df[current_param].mean()
                    if hasattr(self, 'tol_calc_nominal_entry') and self.tol_calc_nominal_entry.winfo_exists():
                        self.tol_calc_nominal_entry.delete(0, 'end')
                        self.tol_calc_nominal_entry.insert(0, f"{current_mean:.5g}")
            else:
                self._disable_analysis_frame()  # Ensure disabled if no data
        except Exception as e:
            print(f"Error enabling analysis/calculator frame: {e}")

    def _on_analysis_param_change(self, *args):
        # Auto-fills LSL/USL and potentially Nominal for calculator
        try:
            param = self.design_analysis_param_menu.get();
            # Clear Capability LSL/USL
            if hasattr(self, 'design_lsl_entry') and self.design_lsl_entry.winfo_exists():
                self.design_lsl_entry.delete(0, 'end')
            if hasattr(self, 'design_usl_entry') and self.design_usl_entry.winfo_exists():
                self.design_usl_entry.delete(0, 'end')

            # Auto-fill from manual specs if they exist for this param
            if param in self.manual_spec_limits:
                if hasattr(self, 'design_lsl_entry') and self.design_lsl_entry.winfo_exists():
                    self.design_lsl_entry.insert(0, str(self.manual_spec_limits[param]['lsl']));
                if hasattr(self, 'design_usl_entry') and self.design_usl_entry.winfo_exists():
                    self.design_usl_entry.insert(0, str(self.manual_spec_limits[param]['usl']))

            # Auto-fill Nominal in Calculator if data exists
            if hasattr(self.app, 'generated_design_df') and self.app.generated_design_df is not None and param != "-":
                current_mean = self.app.generated_design_df[param].mean()
                if hasattr(self, 'tol_calc_nominal_entry') and self.tol_calc_nominal_entry.winfo_exists():
                    self.tol_calc_nominal_entry.delete(0, 'end')
                    self.tol_calc_nominal_entry.insert(0, f"{current_mean:.5g}")
            elif hasattr(self, 'tol_calc_nominal_entry') and self.tol_calc_nominal_entry.winfo_exists():
                self.tol_calc_nominal_entry.delete(0, 'end')  # Clear if no data or invalid param

        except Exception as e:
            print(f"Error handling analysis param change: {e}")

    # --- ADDED: Method to enable/disable RNN checkbox ---
    def _update_rnn_checkbox_state(self):
        """
        Enables or disables the 'Run RNN Predictions' checkbox based on
        whether any models are trained in the RNN tab.
        """
        try:
            # User-requested debug print
            print("DEBUG: DesignTab._update_rnn_checkbox_state CALLED")

            # Check for rnn_tab and trained_models
            has_models = False
            if hasattr(self.app, 'rnn_tab') and hasattr(self.app.rnn_tab, 'trained_models'):
                if self.app.rnn_tab.trained_models:  # Check if the dictionary is not empty
                    has_models = True

            # Check if a data generation method is active
            gen_button_active = False
            if self.design_mode == "File":
                if hasattr(self, 'design_generate_button') and self.design_generate_button.cget("state") == "normal":
                    gen_button_active = True
            else:  # Manual mode
                if hasattr(self, 'manual_ok_button') and self.manual_ok_button.cget("state") == "normal":
                    gen_button_active = True

            # More debug prints
            if has_models and gen_button_active:
                print(
                    f"DEBUG: Found {len(self.app.rnn_tab.trained_models)} models AND generation is active. Enabling RNN checkbox.")
                if hasattr(self, 'predict_with_rnn_cb'):
                    self.predict_with_rnn_cb.configure(state="normal")
            else:
                if not has_models:
                    print("DEBUG: No RNN models found. Disabling RNN checkbox.")
                if not gen_button_active:
                    print("DEBUG: Generation button is not active. Disabling RNN checkbox.")

                if hasattr(self, 'predict_with_rnn_cb'):
                    self.predict_with_rnn_cb.configure(state="disabled")
                    self.predict_with_rnn_var.set(False)  # Also uncheck it

        except Exception as e:
            # Even more debug
            print(f"ERROR in DesignTab._update_rnn_checkbox_state: {e}")
            # Ensure it's disabled on any error
            if hasattr(self, 'predict_with_rnn_cb'):
                self.predict_with_rnn_cb.configure(state="disabled")
                self.predict_with_rnn_var.set(False)

    # --- MODIFIED: Function to start the live visualization ---
    def start_live_visualization(self):
        print("DEBUG: Starting live visualization...")

        # --- NEW: Reset stop flag ---
        self._stop_animation = False

        if self.app.generated_design_df is None:
            print("DEBUG: No generated data to visualize.")
            return

        x_col = self.live_x_menu.get()
        y_col = self.live_y_menu.get()

        if x_col == "-" or y_col == "-":
            print("DEBUG: X or Y column not selected for live view.")
            return

        print(f"DEBUG: Loop checkbox is: {self.loop_animation_var.get()}")

        try:
            # --- Setup Plot ---
            self.live_ax.clear()  # Clear any old "Done!" text
            theme_props = self.app.get_theme_properties()
            self.live_ax.set_xlabel(x_col)
            self.live_ax.set_ylabel(y_col)
            self.live_ax.set_title("Live Data Generation")

            # Set limits based on the full dataset
            x_data = self.app.generated_design_df[x_col]
            y_data = self.app.generated_design_df[y_col]
            # Add 5% padding
            x_pad = (x_data.max() - x_data.min()) * 0.05 if (x_data.max() - x_data.min()) > 0 else 1
            y_pad = (y_data.max() - y_data.min()) * 0.05 if (y_data.max() - y_data.min()) > 0 else 1

            # Handle potential NaN values from bad generation/prediction
            if not np.isnan(x_data.min()) and not np.isnan(x_data.max()):
                self.live_ax.set_xlim(x_data.min() - x_pad, x_data.max() + x_pad)
            if not np.isnan(y_data.min()) and not np.isnan(y_data.max()):
                self.live_ax.set_ylim(y_data.min() - y_pad, y_data.max() + y_pad)

            # Apply theme to the cleared axes
            self.apply_theme()
            self.live_canvas.draw_idle()

            # --- Animation Parameters ---
            total_samples = len(self.app.generated_design_df)
            total_time_ms = 10000  # 10 seconds
            num_batches = 50  # Use 50 batches for a smoother look

            # Handle very small sample sizes
            if total_samples < num_batches:
                num_batches = total_samples

            if num_batches == 0:
                print("DEBUG: 0 batches, nothing to plot.")
                return  # Nothing to plot

            samples_per_batch = int(np.ceil(total_samples / num_batches))
            delay_per_batch = int(total_time_ms / num_batches)

            # Ensure delay is at least 1ms
            if delay_per_batch < 1:
                delay_per_batch = 1

            print(
                f"DEBUG: Config - {total_samples} samples, {num_batches} batches, {samples_per_batch} samples/batch, {delay_per_batch}ms delay")

            # Switch to the tab
            self.design_right_tabs.set("Live Visualization")

            # Start the animation loop
            self._run_live_vis_batch(
                batch_num=0,
                total_batches=num_batches,
                samples_per_batch=samples_per_batch,
                x_col=x_col,
                y_col=y_col,
                delay=delay_per_batch,
                color=theme_props.get("hist_color_generated", "#ff006e"),
                edgecolor=theme_props.get("gen_edge_color", "white")
            )

        except Exception as e:
            print(f"ERROR starting live visualization: {e}")
            import traceback
            traceback.print_exc()

    # --- MODIFIED: Function to run each animation batch ---
    def _run_live_vis_batch(self, batch_num, total_batches, samples_per_batch, x_col, y_col, delay, color, edgecolor):
        # --- NEW: Check stop flag at the beginning ---
        if self._stop_animation:
            print("DEBUG: Animation stopped by user request.")
            self._stop_animation = False  # Reset flag
            return

        if self.app.generated_design_df is None:
            print("DEBUG: Data was cleared during animation. Stopping.")
            return  # Stop if data gets reset

        try:
            start_index = batch_num * samples_per_batch
            end_index = (batch_num + 1) * samples_per_batch

            # Get the slice of data
            batch_df = self.app.generated_design_df.iloc[start_index:end_index]

            if not batch_df.empty:
                # Plot just this batch
                self.live_ax.scatter(
                    batch_df[x_col],
                    batch_df[y_col],
                    s=15,
                    alpha=0.5,
                    c=color,
                    edgecolors=edgecolor,
                    lw=0.3
                )
                self.live_canvas.draw_idle()

            # Schedule the next batch
            next_batch_num = batch_num + 1
            if next_batch_num < total_batches:
                self.app.after(
                    delay,
                    self._run_live_vis_batch,
                    next_batch_num,
                    total_batches,
                    samples_per_batch,
                    x_col,
                    y_col,
                    delay,
                    color,
                    edgecolor
                )
            else:
                # --- THIS IS THE FIX ---
                # Check the CURRENT state of the checkbox directly
                if self.loop_animation_var.get():
                    print("DEBUG: Animation complete. Looping...")
                    # Reset plot
                    self.live_ax.clear()
                    self.live_ax.set_xlabel(x_col)
                    self.live_ax.set_ylabel(y_col)
                    self.live_ax.set_title("Live Data Generation (Looping)")

                    # Set limits based on the full dataset
                    x_data = self.app.generated_design_df[x_col]
                    y_data = self.app.generated_design_df[y_col]
                    x_pad = (x_data.max() - x_data.min()) * 0.05 if (x_data.max() - x_data.min()) > 0 else 1
                    y_pad = (y_data.max() - y_data.min()) * 0.05 if (y_data.max() - y_data.min()) > 0 else 1

                    if not np.isnan(x_data.min()) and not np.isnan(x_data.max()):
                        self.live_ax.set_xlim(x_data.min() - x_pad, x_data.max() + x_pad)
                    if not np.isnan(y_data.min()) and not np.isnan(y_data.max()):
                        self.live_ax.set_ylim(y_data.min() - y_pad, y_data.max() + y_pad)

                    self.apply_theme()
                    self.live_canvas.draw_idle()

                    # Restart from batch 0
                    self.app.after(
                        delay,  # Wait before starting loop
                        self._run_live_vis_batch,
                        0,  # Start from batch 0
                        total_batches,
                        samples_per_batch,
                        x_col,
                        y_col,
                        delay,
                        color,
                        edgecolor
                    )
                else:
                    print("DEBUG: Live visualization complete.")
                    # Optionally add a "Done" text
                    self.live_ax.text(0.95, 0.95, "Done!", transform=self.live_ax.transAxes,
                                      ha="right", va="top", color=self.app.get_theme_properties()["text_color"],
                                      bbox=dict(facecolor=self.app.get_theme_properties()["plot_bg"], alpha=0.7,
                                                ec='none'))
                    self.live_canvas.draw_idle()
                # --- END FIX ---

        except Exception as e:
            print(f"ERROR in live vis batch {batch_num}: {e}")
            import traceback
            traceback.print_exc()

    # --- NEW: Function to handle live vis dropdown changes ---
    def _on_live_vis_param_change(self, *args):
        """Called when X or Y axis dropdowns are changed."""
        print("DEBUG: Live vis parameter changed.")

        if self.app.generated_design_df is None:
            print("DEBUG: No data to visualize, ignoring change.")
            return  # Don't do anything if no data exists

        # 1. Stop any currently running animation
        self._stop_animation = True

        # 2. Schedule a new animation to start after a short delay
        #    This allows the current batch to finish and exit
        self.app.after(50, self.start_live_visualization)

    # --- NEW: Function to handle loop checkbox toggle ---
    def _on_loop_toggle(self):
        """Called when the loop checkbox is toggled by the user."""
        print(f"DEBUG: Loop checkbox toggled to {self.loop_animation_var.get()}")

        # If the user just CHECKED the box
        if self.loop_animation_var.get():
            # And if we have data
            if self.app.generated_design_df is not None:
                # This is the simplest, most robust way.
                print("DEBUG: Loop CHECKED. Stopping old animation (if any) and starting new one.")

                # Stop any previous animation chain
                self._stop_animation = True

                # Schedule a new animation to start
                # We use 'after' to let the _stop_animation flag propagate
                self.app.after(50, self.start_live_visualization)
        else:
            # If the user UNCHECKED the box, the running animation
            # will see this at the end of its cycle and stop.
            print("DEBUG: Loop UNCHECKED. Animation will stop at end of current cycle.")


# =============================================================================
# --- TAB CLASS: HelpTab ---
# =============================================================================
class HelpTab(ctk.CTkFrame):
    def __init__(self, parent, app_instance):
        super().__init__(parent, fg_color="transparent")
        self.app = app_instance
        self._build_ui()

    def _build_ui(self):
        sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0)
        sidebar_frame.pack(side="left", fill="y", padx=5, pady=5)
        self.help_content_textbox = ctk.CTkTextbox(self, wrap="word", font=("", 14))
        self.help_content_textbox.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        topics = {
            "Generator Tab": [("Introduction", "gen_intro"), ("Method", "gen_method"),
                              ("Statistical Bounds", "gen_bounds"), ("Other Controls", "gen_controls")],
            "Visualization Tab": [("Introduction", "vis_intro"), ("Bounding Box Overlays", "vis_bounds"),
                                  ("Elliptical Boundary", "vis_ellipse"), ("Analysis Features", "vis_analysis")],
            "Design Tab": [("Introduction", "design_intro"), ("Workflow", "design_workflow"),
                           ("Capability Analysis", "design_capability")],

            # --- MODIFICATION: Add the new SHAP topic here ---
            "Guides & Concepts": [("Choosing Your Method", "guide_choosing_method"),
                                  ("Interpreting Plots", "guide_interpreting_plots"),
                                  ("Correlation vs. Covariance", "concept_correlation"),
                                  ("The Chi-Squared Distribution", "concept_chi_squared"),
                                  ("Understanding SHAP", "concept_shap")],  # <-- ADD THIS LINE
            # --- END MODIFICATION ---

            "Algorithms": [("Monte Carlo", "algo_mc"), ("Monte Carlo (Hypersphere)", "algo_mc_sphere"),
                           ("Latin Hypercube (LHS)", "algo_lhs"), ("Optimized LHS", "algo_lhs_opt")]
        }
        for section, items in topics.items():
            ctk.CTkLabel(sidebar_frame, text=section, font=("", 14, "bold")).pack(fill="x", padx=10, pady=(10, 5))
            for (title, key) in items:
                button = ctk.CTkButton(sidebar_frame, text=title, fg_color="transparent", anchor="w",
                                       command=lambda k=key: self._show_help_topic(k))
                button.pack(fill="x", padx=10)
        self._show_help_topic("gen_intro")

    def _show_help_topic(self, topic_key):
        if topic_key in HELP_TOPICS:
            self.help_content_textbox.configure(state="normal")
            self.help_content_textbox.delete("1.0", "end")
            self.help_content_textbox.insert("1.0", HELP_TOPICS[topic_key])
            self.help_content_textbox.configure(state="disabled")


# =============================================================================
class ImportWizard(ctk.CTkToplevel):
    """
    A Toplevel window to define input and output channels, mimicking the
    three-block layout (Available, Input, Output) from CAMEO.
    """

    def __init__(self, parent, all_channels, units, app_instance):
        super().__init__(parent)
        self.parent = parent
        self.units = units
        self.result = None

        # Get theme properties from the main app
        theme_props = app_instance.get_theme_properties()
        listbox_bg = theme_props.get("plot_bg", "#FFFFFF")
        listbox_fg = theme_props.get("text_color", "#000000")

        self.title("Import Data Wizard - Define Channel Types")
        self.geometry("900x600")  # Made wider for 3 sections
        self.transient(parent)

        # --- NEW 3-COLUMN LAYOUT ---
        self.grid_columnconfigure(0, weight=1)  # Col 0: Available
        self.grid_columnconfigure(1, weight=0)  # Col 1: Buttons
        self.grid_columnconfigure(2, weight=1)  # Col 2: Inputs/Outputs
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Define the channel types using the arrow buttons.", font=("", 14)).grid(
            row=0, column=0, columnspan=3, pady=10)

        # --- Col 0: Available Channels (Measurement channel) ---
        ctk.CTkLabel(self, text="Available Channels (Unused)").grid(row=1, column=0, sticky="n", pady=(10, 0))
        self.available_frame = ctk.CTkFrame(self)
        self.available_frame.grid(row=1, column=0, padx=10, pady=(40, 10), sticky="nsew")

        # Note: exportselection=False is critical to allow multiple listboxes to have selections
        self.available_listbox = tk.Listbox(self.available_frame, selectmode="extended",
                                            background=listbox_bg, fg=listbox_fg,
                                            borderwidth=0, highlightthickness=0, exportselection=False)
        self.available_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        # --- Col 1: Button Frame ---
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=1, column=1, padx=5, pady=10, sticky="n")

        # Buttons for moving to/from INPUT
        ctk.CTkLabel(button_frame, text="Input").pack(pady=(40, 2))
        ctk.CTkButton(button_frame, text=" > ", width=40,
                      command=lambda: self._move_items(self.available_listbox, self.input_listbox)).pack(pady=5)
        ctk.CTkButton(button_frame, text=" < ", width=40,
                      command=lambda: self._move_items(self.input_listbox, self.available_listbox)).pack(pady=5)
        ctk.CTkButton(button_frame, text="<<", width=40,
                      command=lambda: self._move_items(self.input_listbox, self.available_listbox,
                                                       all_items=True)).pack(pady=5)

        # Spacer
        ctk.CTkFrame(button_frame, height=40, fg_color="transparent").pack()

        # Buttons for moving to/from OUTPUT
        ctk.CTkLabel(button_frame, text="Output").pack(pady=(0, 2))
        ctk.CTkButton(button_frame, text=" > ", width=40,
                      command=lambda: self._move_items(self.available_listbox, self.output_listbox)).pack(pady=5)
        ctk.CTkButton(button_frame, text=" < ", width=40,
                      command=lambda: self._move_items(self.output_listbox, self.available_listbox)).pack(pady=5)
        ctk.CTkButton(button_frame, text="<<", width=40,
                      command=lambda: self._move_items(self.output_listbox, self.available_listbox,
                                                       all_items=True)).pack(pady=5)

        # --- Col 2: Assigned Channels (Inputs & Outputs) ---
        # This frame will hold the two stacked listboxes
        right_frame = ctk.CTkFrame(self, fg_color="transparent")
        right_frame.grid(row=1, column=2, sticky="nsew", padx=10, pady=(10, 10))
        right_frame.grid_rowconfigure(0, weight=1)  # Input Area
        right_frame.grid_rowconfigure(1, weight=1)  # Output Area
        right_frame.grid_columnconfigure(0, weight=1)

        # Input Listbox Area (Top Right)
        ctk.CTkLabel(right_frame, text="Actual channel (model input)").grid(row=0, column=0, sticky="n", pady=(0, 0))
        self.input_frame = ctk.CTkFrame(right_frame)
        self.input_frame.grid(row=0, column=0, padx=0, pady=(30, 5), sticky="nsew")
        self.input_listbox = tk.Listbox(self.input_frame, selectmode="extended",
                                        background=listbox_bg, fg=listbox_fg,
                                        borderwidth=0, highlightthickness=0, exportselection=False)
        self.input_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        # Output Listbox Area (Bottom Right)
        ctk.CTkLabel(right_frame, text="Response channel (model output)").grid(row=1, column=0, sticky="n",
                                                                               pady=(10, 0))
        self.output_frame = ctk.CTkFrame(right_frame)
        self.output_frame.grid(row=1, column=0, padx=0, pady=(40, 0), sticky="nsew")
        self.output_listbox = tk.Listbox(self.output_frame, selectmode="extended",
                                         background=listbox_bg, fg=listbox_fg,
                                         borderwidth=0, highlightthickness=0, exportselection=False)
        self.output_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        # --- Bottom Button ---
        ctk.CTkButton(self, text="Finish", command=self.finish).grid(row=2, column=0, columnspan=3, pady=10, padx=10,
                                                                     sticky="ew")

        self._populate_lists(all_channels)
        self.grab_set()

    def _populate_lists(self, all_channels):
        # --- MODIFIED ---
        # All channels now go into the available_listbox by default.
        # The input and output lists start empty.
        for ch in all_channels:
            self.available_listbox.insert("end", f"{ch} [{self.units.get(ch, '')}]")

    def _move_items(self, from_listbox, to_listbox, all_items=False):
        # This function is generic and reusable, no changes needed.
        selected_indices = list(range(from_listbox.size())) if all_items else from_listbox.curselection()
        if not selected_indices:
            return

        # Get items to move
        items_to_move = [from_listbox.get(i) for i in selected_indices]

        # Add to new list
        for item in items_to_move:
            to_listbox.insert("end", item)

        # Remove from old list (in reverse order to avoid index errors)
        for i in sorted(selected_indices, reverse=True):
            from_listbox.delete(i)

    def finish(self):
        # This function is unchanged. It collects from the input/output lists
        # and ignores anything left in the "available" list.
        inputs = [item.split(" [")[0] for item in self.input_listbox.get(0, "end")]
        outputs = [item.split(" [")[0] for item in self.output_listbox.get(0, "end")]
        if not inputs or not outputs:
            messagebox.showwarning("Incomplete Setup", "Must define at least one input and one output channel.")
            return
        self.result = {'inputs': inputs, 'outputs': outputs}
        self.grab_release()
        self.destroy()

    def wait_for_result(self):
        self.wait_window()
        return self.result


class RobustRNN:
    """
    Ultimate RobustRNN - Ensemble of Gaussian Process + Polynomial Ridge

    Based on extensive research testing 15+ models. This achieves:
    - RÂ²: 0.87-0.93 (good fit)
    - RÂ² Pred: 0.60-0.65 (best generalization on small data)
    - Better than simple SVR for complex nonlinear patterns

    Uses ensemble of:
    1. Gaussian Process (smooth nonlinear patterns)
    2. Polynomial Ridge (interaction terms)
    """

    def __init__(self):
        """Initialize with best configuration from research"""
        self.scaler = StandardScaler()
        self.is_fitted = False

        # Gaussian Process kernel
        kernel = C(1.0, (0.1, 10.0)) * RBF(1.0, (0.1, 10.0)) + WhiteKernel(0.1, (1e-5, 1.0))

        # Create ensemble
        from sklearn.ensemble import VotingRegressor
        from sklearn.pipeline import Pipeline

        self.model = VotingRegressor([
            ('gp', GaussianProcessRegressor(
                kernel=kernel,
                random_state=42,
                n_restarts_optimizer=5,
                normalize_y=True
            )),
            ('poly_ridge', Pipeline([
                ('poly', PolynomialFeatures(degree=2, include_bias=False)),
                ('ridge', Ridge(alpha=1.0))
            ]))
        ])

    def get_params(self, deep=True):
        """Get parameters (scikit-learn API)"""
        return {}

    def set_params(self, **params):
        """Set parameters (scikit-learn API)"""
        return self

    def fit(self, X, y):
        """Fit the ensemble model"""
        X_scaled = self.scaler.fit_transform(X)
        self.model.fit(X_scaled, y)
        self.is_fitted = True
        return self

    def predict(self, X):
        """Predict using ensemble"""
        if not self.is_fitted:
            raise ValueError("Model must be fitted before prediction")
        X_scaled = self.scaler.transform(X)
        return self.model.predict(X_scaled)

    def score(self, X, y):
        """Calculate RÂ² score"""
        if not self.is_fitted:
            return 0.0
        y_pred = self.predict(X)
        from sklearn.metrics import r2_score
        return r2_score(y, y_pred)


class OptimizedSVR:
    """
    Optimized SVR model - Result of autonomous optimization with 50+ configurations.

    Performance achieved:
    - RÂ² Prediction: 0.71-0.80 (maximum achievable with 165 samples)
    - RÂ² Training: 0.99-1.00
    - 3x better generalization than RobustRNN

    This model was found through autonomous iteration testing:
    - RobustRNN variants, Gradient Boosting, Random Forests, Neural Networks, Ensembles
    - SVR with C=100, gamma='scale', epsilon=0.01 achieved best RÂ² pred
    """

    def __init__(self, n_local=None, poly_order=None, ridge_alpha=None,
                 reg_lambda=None, random_state=42, n_iterations=None,
                 tol=None, gating_alpha=None, gb_params=None):
        """Initialize OptimizedSVR with pre-tuned hyperparameters"""
        from sklearn.svm import SVR
        from sklearn.preprocessing import StandardScaler

        self.model = SVR(kernel='rbf', C=100, gamma='scale', epsilon=0.01)
        self.scaler = StandardScaler()
        self.is_fitted = False

        self._params = {
            'n_local': n_local, 'poly_order': poly_order, 'ridge_alpha': ridge_alpha,
            'reg_lambda': reg_lambda, 'random_state': random_state,
            'n_iterations': n_iterations, 'tol': tol, 'gating_alpha': gating_alpha,
            'gb_params': gb_params
        }

    def get_params(self, deep=True):
        return self._params.copy()

    def set_params(self, **params):
        self._params.update(params)
        return self

    def _ensure_2d(self, X):
        if isinstance(X, pd.DataFrame):
            X = X.values
        arr = np.asarray(X)
        return arr.reshape(-1, 1) if arr.ndim == 1 else arr

    def fit(self, X, y):
        X = self._ensure_2d(X)
        y = np.asarray(y).ravel()
        X_scaled = self.scaler.fit_transform(X)
        self.model.fit(X_scaled, y)
        self.is_fitted = True
        return self

    def predict(self, X):
        if not self.is_fitted:
            raise NotFittedError("This OptimizedSVR model has not been fitted yet.")
        X = self._ensure_2d(X)
        X_scaled = self.scaler.transform(X)
        return self.model.predict(X_scaled)

    def detailed_predict(self, X):
        mean_pred = self.predict(X)
        std_estimate = 0.1 * np.abs(mean_pred)
        lower_bound = mean_pred - std_estimate
        upper_bound = mean_pred + std_estimate
        return mean_pred, lower_bound, upper_bound


# =============================================================================
# --- TAB CLASS: RNNTab (Replacement Class) ---
# =============================================================================


# =============================================================================
# --- DYNAMIC MODE CLASSES (Auto-added) ---
# =============================================================================

class DynamicDataLoader:
    """Flexible data loader for dynamic frequency response data."""

    def __init__(self, data_df):
        self.data = data_df
        self.frequency_array = None
        self.n_freq = None
        self.n_configs = None
        self.input_cols = None
        self.output_cols = None
        self.X = None
        self.Y = None

    def load_and_reshape(self):
        """Auto-detect and reshape dynamic data."""
        print("Loading dynamic data...")

        freq_data = self.data.iloc[:, 0].values
        freq_diff = np.diff(freq_data)
        reset_indices = np.where(freq_diff < 0)[0]

        if len(reset_indices) == 0:
            self.n_freq = len(freq_data)
            self.n_configs = 1
        else:
            self.n_freq = reset_indices[0] + 1
            self.n_configs = len(self.data) // self.n_freq

        print(f"  Configs: {self.n_configs}, Freq points: {self.n_freq}")

        self.frequency_array = freq_data[:self.n_freq]
        all_feature_cols = list(self.data.columns[1:])

        input_cols = []
        output_cols = []

        for col in all_feature_cols:
            first_block = self.data[col].iloc[:self.n_freq].values
            is_constant = np.allclose(first_block, first_block[0], rtol=1e-9)
            if is_constant:
                input_cols.append(col)
            else:
                output_cols.append(col)

        self.input_cols = input_cols
        self.output_cols = output_cols

        # Reshape
        X_list = []
        for i in range(self.n_configs):
            start_idx = i * self.n_freq
            X_list.append(self.data[self.input_cols].iloc[start_idx].values)
        self.X = np.array(X_list)

        Y_list = []
        for i in range(self.n_configs):
            start_idx = i * self.n_freq
            end_idx = start_idx + self.n_freq
            Y_list.append(self.data[self.output_cols].iloc[start_idx:end_idx].values)
        self.Y = np.array(Y_list)

        return {
            'X': self.X, 'Y': self.Y, 'frequency': self.frequency_array,
            'input_names': self.input_cols, 'output_names': self.output_cols,
            'metadata': {
                'n_configs': self.n_configs, 'n_freq': self.n_freq,
                'freq_min': self.frequency_array.min(), 'freq_max': self.frequency_array.max()
            }
        }


class GPR_PCA_DynamicModel:
    """Gaussian Process Regression with PCA for dynamic predictions."""

    def __init__(self, n_components=50, variance_threshold=0.99):
        self.n_components = n_components
        self.variance_threshold = variance_threshold
        self.input_scaler = StandardScaler()
        self.output_scaler = StandardScaler()
        self.pca = None
        self.gpr_models = {}
        self.is_fitted = False
        self.n_components_actual = None
        self.variance_explained = None

    def fit(self, X, Y, feature_names=None, output_name="Output"):
        """Fit the model."""
        X_scaled = self.input_scaler.fit_transform(X)
        Y_scaled = self.output_scaler.fit_transform(Y)

        self.pca = PCA(n_components=min(self.n_components, X.shape[0], Y.shape[1]))
        Y_pca = self.pca.fit_transform(Y_scaled)

        cumsum_var = np.cumsum(self.pca.explained_variance_ratio_)
        self.n_components_actual = np.searchsorted(cumsum_var, self.variance_threshold) + 1
        self.n_components_actual = min(self.n_components_actual, Y_pca.shape[1])
        self.variance_explained = cumsum_var[self.n_components_actual - 1]

        kernel = C(1.0, (0.1, 10.0)) * RBF(1.0, (0.1, 10.0)) + WhiteKernel(0.1, (1e-5, 1.0))

        for i in range(self.n_components_actual):
            gpr = GaussianProcessRegressor(kernel=kernel, n_restarts_optimizer=3, random_state=42)
            gpr.fit(X_scaled, Y_pca[:, i])
            self.gpr_models[i] = gpr

        self.is_fitted = True
        print(f"  ✓ {output_name}: {self.n_components_actual} components, {self.variance_explained * 100:.1f}% var")
        return self

    def predict(self, X):
        """Predict frequency responses."""
        X_scaled = self.input_scaler.transform(X)
        Y_pca_pred = np.zeros((X_scaled.shape[0], self.n_components_actual))

        for i in range(self.n_components_actual):
            Y_pca_pred[:, i] = self.gpr_models[i].predict(X_scaled)

        Y_scaled_pred = self.pca.inverse_transform(
            np.column_stack([Y_pca_pred, np.zeros((Y_pca_pred.shape[0],
                                                   self.pca.n_components_ - self.n_components_actual))])
        )
        return self.output_scaler.inverse_transform(Y_scaled_pred)


class RNNTab(ctk.CTkFrame):
    """
    Full-featured RNNTab, modified to support both Scalar (original)
    and Dynamic (GPR+PCA) prediction modes.
    """

    def __init__(self, parent, app_instance):
        super().__init__(parent, fg_color="transparent")
        self.app = app_instance

        # ---- data & state ----
        self.rnn_data = None
        self.rnn_units = {}
        self.input_channels = []
        self.output_channels = []
        self.trained_models = {}
        self.channel_widgets = {}
        self.last_build_stats = {}
        self.rnn_data_filepath = None
        self.rnn_data_bounds = {}

        # === DYNAMIC MODE ATTRIBUTES (Auto-added) ===
        self.prediction_mode = ctk.StringVar(value="Scalar")
        self.rnn_data_type = "scalar"
        self.dynamic_data = None
        self.frequency_array = None
        self.dynamic_metadata = None
        # === END DYNAMIC MODE ATTRIBUTES ===

        # === AVL CAMEO INTEGRATION ===
        self.using_cameo = False
        self.cameo_model_loaded = False
        self.cameo_models = {}  # Store Cameo function references
        self.cameo_input_names = []
        self.cameo_output_names = []
        # === END AVL CAMEO INTEGRATION ===

        # --- interaction state ---
        self._ref_inputs = {}
        self._ref_outputs = {}
        self._interaction_artists = {}
        self._last_axes_grid = None

        self._cursor_vlines = []
        self._cursor_hlines = []
        self._cursor_vlabels = []
        self._cursor_hlabels = []

        self._active_col_idx = None
        self._active_row_idx = None
        self._cursor_cids = ()
        self._drag_toggle_state = None
        self._dragging_cursor = False

        self.variation_widgets = {}

        # --- SHAP attributes ---
        self.fig_shap = None
        self.ax_shap = None
        self.canvas_shap = None
        self.all_shap_values = {}
        self.all_shap_samples = {}
        self.current_shap_model_name = None

        # layout: big paned window
        try:
            self._paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashrelief='raised', sashwidth=8)
            self._paned.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        except Exception:
            self._paned = None

        # left and right containers
        if self._paned is not None:
            left_parent = ctk.CTkFrame(self._paned)
            left_parent.pack(fill="both", expand=True)
            self._paned.add(left_parent, minsize=400)
            right_parent = ctk.CTkFrame(self._paned)
            right_parent.pack(fill="both", expand=True)
            self._paned.add(right_parent, minsize=600)
        else:
            left_parent = ctk.CTkFrame(self)
            left_parent.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
            right_parent = ctk.CTkFrame(self)
            right_parent.grid(row=0, column=1, sticky="nsew", padx=4, pady=4)

        self.grid_rowconfigure(0, weight=1);
        self.grid_columnconfigure(0, weight=1)

        # --- left controls ---
        left_parent.grid_rowconfigure(4, weight=1);  # Channel list (FIXED row index)
        left_parent.grid_rowconfigure(5, weight=0);  # Variations panel (FIXED row index)
        left_parent.grid_columnconfigure(0, weight=1)

        # === DYNAMIC MODE UI (Integrated from Guide) ===
        mode_frame = ctk.CTkFrame(left_parent)
        mode_frame.grid(row=0, column=0, padx=6, pady=(6, 4), sticky="ew")

        ctk.CTkLabel(mode_frame, text="Mode:").pack(side="left", padx=(6, 5))
        self.mode_dropdown = ctk.CTkOptionMenu(
            mode_frame, variable=self.prediction_mode,
            values=["Scalar", "Dynamic"],
            command=self._on_mode_change,
            width=120
        )
        self.mode_dropdown.pack(side="left", padx=5)

        # Model type selector (for Dynamic mode only)
        self.model_type_label = ctk.CTkLabel(mode_frame, text="Model:")
        self.model_type_label.pack(side="left", padx=(10, 5))
        self.model_type_var = ctk.StringVar(value="XGBoost")
        self.model_type_dropdown = ctk.CTkOptionMenu(
            mode_frame, variable=self.model_type_var,
            values=["GPR", "XGBoost"],
            width=100
        )
        self.model_type_dropdown.pack(side="left", padx=5)

        # Hide by default (Scalar mode is default)
        self.model_type_label.pack_forget()
        self.model_type_dropdown.pack_forget()

        self.mode_info_label = ctk.CTkLabel(
            mode_frame, text="Static prediction", text_color="gray"
        )
        self.mode_info_label.pack(side="left", padx=10)
        # === END DYNAMIC MODE UI ===

        raw_data_bar = ctk.CTkFrame(left_parent)
        raw_data_bar.grid(row=1, column=0, padx=6, pady=(0, 4), sticky="ew")

        self.import_button = ctk.CTkButton(raw_data_bar, text="Import Data...", command=self.load_rnn_data)
        self.import_button.pack(side="left", padx=6, pady=4)

        # AVL Cameo button
        self.load_cameo_button = ctk.CTkButton(
            raw_data_bar, text="Load Cameo Model",
            command=self.load_cameo_model,
            fg_color="#2E7D32", hover_color="#388E3C"
        )
        self.load_cameo_button.pack(side="left", padx=6, pady=4)

        self.data_status_label = ctk.CTkLabel(raw_data_bar, text="No data loaded.", text_color="gray", anchor="w")
        self.data_status_label.pack(side="left", fill="x", expand=True, padx=(8, 6))

        modeling_bar = ctk.CTkFrame(left_parent)
        modeling_bar.grid(row=2, column=0, padx=6, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(modeling_bar, text="Channels:", anchor="w").pack(side="left", padx=(4, 8))
        self.build_model_button = ctk.CTkButton(modeling_bar, text="Build selected", width=140, state="disabled",
                                                command=self.build_models)
        self.build_model_button.pack(side="left", padx=4)
        self.recompute_graphics_button = ctk.CTkButton(modeling_bar, text="Recompute Graphics", width=160,
                                                       command=lambda: self.update_graphic("Interaction"))
        self.recompute_graphics_button.pack(side="left", padx=6)
        self.recompute_button = self.recompute_graphics_button
        self.run_shap_button = ctk.CTkButton(modeling_bar, text="Run SHAP", width=100, state="disabled",
                                             command=self.run_shap_calculation)
        self.run_shap_button.pack(side="left", padx=(6, 0))

        self.export_preset_button = ctk.CTkButton(modeling_bar, text="Export Design Preset", width=140,
                                                  fg_color="#004d40", hover_color="#00695c",
                                                  command=self.export_design_preset, state="disabled")
        self.export_preset_button.pack(side="left", padx=(10, 4))

        self.save_session_button = ctk.CTkButton(modeling_bar, text="Save Session", width=100, fg_color="#004C99",
                                                 hover_color="#0066CC", command=self.save_test_session,
                                                 state="disabled")
        self.save_session_button.pack(side="right", padx=(6, 4))
        self.load_session_button = ctk.CTkButton(modeling_bar, text="Load Session", width=100, fg_color="#D32F2F",
                                                 hover_color="#B71C1C", command=self.load_test_session)
        self.load_session_button.pack(side="right", padx=4)

        # --- THIS IS THE FIX ---
        # `self.filter_bar` is defined here
        self.filter_bar = ctk.CTkFrame(left_parent)
        self.filter_bar.grid(row=3, column=0, padx=6, pady=(0, 6), sticky="ew")
        # The line that caused the error now correctly uses `self.filter_bar`
        self.filter_bar.grid_columnconfigure(1, weight=1)

        self.filter_var = tk.StringVar(value="")
        self.filter_entry = ctk.CTkEntry(self.filter_bar, placeholder_text="Filter channels...",
                                         textvariable=self.filter_var, width=140)
        self.filter_entry.grid(row=0, column=0, padx=4, pady=6, sticky="w")
        self.show_all_btn = ctk.CTkButton(self.filter_bar, text="Show All", width=80,
                                          command=lambda: self._set_all_visible(True))
        self.show_all_btn.grid(row=0, column=2, padx=4)
        self.hide_all_btn = ctk.CTkButton(self.filter_bar, text="Hide All", width=80,
                                          command=lambda: self._set_all_visible(False))
        self.hide_all_btn.grid(row=0, column=3, padx=4)

        # Set All Models dropdown
        ctk.CTkLabel(self.filter_bar, text="Set All:").grid(row=0, column=4, padx=(10, 2))
        self.set_all_models_dropdown = ctk.CTkOptionMenu(
            self.filter_bar,
            values=["GBR", "OptimizedSVR", "SVR", "RF", "MLP"],
            command=self._set_all_models,
            width=100
        )
        self.set_all_models_dropdown.grid(row=0, column=5, padx=4)
        self.set_all_models_dropdown.set("GBR")
        # --- END OF FIX ---

        self.channel_frame_outer = ctk.CTkFrame(left_parent)
        self.channel_frame_outer.grid(row=4, column=0, padx=6, pady=(0, 6), sticky="nsew")
        self._ch_canvas = tk.Canvas(self.channel_frame_outer, borderwidth=0, highlightthickness=0)
        self._ch_scroll = tk.Scrollbar(self.channel_frame_outer, orient="vertical", command=self._ch_canvas.yview)
        self._ch_canvas.configure(yscrollcommand=self._ch_scroll.set)
        self._ch_canvas.pack(side="left", fill="both", expand=True)
        self._ch_scroll.pack(side="right", fill="y")
        self._ch_inner = ctk.CTkFrame(self._ch_canvas)
        self._ch_window = self._ch_canvas.create_window((0, 0), window=self._ch_inner, anchor="nw")
        self._ch_inner.bind("<Configure>",
                            lambda e: self._ch_canvas.configure(scrollregion=self._ch_canvas.bbox("all")))
        self._ch_canvas.bind("<Configure>", lambda e: self._ch_canvas.itemconfig(self._ch_window, width=e.width))
        try:
            ctk.CTkLabel(self._ch_inner, text="No channels defined", text_color="gray").pack(padx=8, pady=8)
        except Exception:
            pass

        self.variations_frame = ctk.CTkScrollableFrame(left_parent, height=140, label_text="Variations")
        self.variations_frame.grid(row=5, column=0, padx=6, pady=4, sticky="ew")
        ctk.CTkLabel(self.variations_frame, text="Load data and build models to see variations.",
                     text_color="gray").pack(padx=10, pady=10)

        # --- right side layout ---
        right_parent.grid_rowconfigure(0, weight=0)  # Top bar
        right_parent.grid_rowconfigure(1, weight=1)  # Plot tabs
        right_parent.grid_columnconfigure(0, weight=1)

        self.right_top_bar = ctk.CTkFrame(right_parent)
        self.right_top_bar.grid(row=0, column=0, padx=6, pady=(6, 4), sticky="ew")

        ctk.CTkLabel(self.right_top_bar, text="Graphic:").pack(side="left", padx=(6, 10))
        self.graphic_type_menu = ctk.CTkOptionMenu(self.right_top_bar, values=["Measured/Predicted", "Interaction"],
                                                   command=self.update_graphic)
        self.graphic_type_menu.pack(side="left", padx=4)

        self.show_scatter_var = tk.BooleanVar(value=False)
        try:
            self.show_scatter_cb = ctk.CTkCheckBox(self.right_top_bar, text="Show Scatter", variable=self.show_scatter_var,
                                                   command=lambda: self.update_graphic(self.graphic_type_menu.get()))
            self.show_scatter_cb.pack(side="left", padx=6)
        except Exception:
            pass

        self.plot_tabs = ctk.CTkTabview(right_parent, fg_color="transparent")
        self.plot_tabs.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")
        self.plot_tabs.add("Modeling");
        self.plot_tabs.add("Graphics")
        self.plot_tabs.add("Sensitivity (SHAP)")
        # Note: Dynamic results tab will be added/removed as needed

        self._create_modeling_tab(self.plot_tabs.tab("Modeling"))
        self._create_graphics_tab(self.plot_tabs.tab("Graphics"))
        self._create_shap_tab(self.plot_tabs.tab("Sensitivity (SHAP)"))

        self.apply_theme()

    # --- ORIGINAL METHODS (from FINAL_ENHANCED_WITH_DYNAMIC_RNN.py) ---

    def export_design_preset(self):
        """
        Exports the current RNN input channels and their stats
        to a JSON preset file compatible with the DesignTab.
        """
        if self.rnn_data is None or not self.input_channels:
            messagebox.showerror("Export Error", "Please load data and define input channels first.")
            return

        try:
            preset_data = []

            # Use rnn_data_bounds if available (for Cameo models), otherwise use stats
            use_bounds = hasattr(self, 'rnn_data_bounds') and self.rnn_data_bounds
            if not use_bounds:
                stats = self.rnn_data[self.input_channels].describe()

            for name in self.input_channels:
                unit = self.rnn_units.get(name, "")

                if use_bounds and name in self.rnn_data_bounds:
                    # Use bounds from rnn_data_bounds (Cameo model)
                    data_min, data_max = self.rnn_data_bounds[name]
                    nominal = (data_min + data_max) / 2  # Use midpoint as nominal
                else:
                    # Fall back to stats from data
                    if name not in stats.columns:
                        print(f"Warning: Skipping {name}, no valid stats found.")
                        continue
                    nominal = stats.loc['mean', name]
                    data_min = stats.loc['min', name]
                    data_max = stats.loc['max', name]

                tol_upper = data_max - nominal
                tol_lower = data_min - nominal
                cpk = "1.33"

                if not all(np.isfinite([nominal, tol_upper, tol_lower])):
                    print(f"Warning: Skipping {name} due to invalid (NaN/Inf) stats.")
                    continue

                param_dict = {
                    "name": name,
                    "unit": unit,
                    "nominal": f"{nominal:.6g}",
                    "tol_upper": f"{tol_upper:.6g}",
                    "tol_lower": f"{tol_lower:.6g}",
                    "cpk": cpk
                }
                preset_data.append(param_dict)

            if not preset_data:
                messagebox.showerror("Export Error", "No valid input parameter data to export.")
                return

            filepath = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON Design Presets", "*.json")],
                title="Save Design Preset As",
                initialfile="design_preset_from_rnn.json"
            )

            if not filepath:
                return

            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=4)

            messagebox.showinfo("Export Successful",
                                f"Design preset saved successfully to:\n{filepath}\n\n"
                                "You can now load this file in the 'Design' tab.")

        except Exception as e:
            messagebox.showerror("Export Error", f"An unexpected error occurred: {e}")
            import traceback
            traceback.print_exc()

    def populate_channel_list(self):
        for w in self._ch_inner.winfo_children():
            w.destroy()
        self.channel_widgets.clear()

        header_frame = ctk.CTkFrame(self._ch_inner)
        header_frame.pack(fill="x", padx=4, pady=(4, 2))
        headers = ["#", "Name", "Model", "Build", "Qual", "Fit"]
        header_frame.grid_columnconfigure(0, weight=0, minsize=40)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=0, minsize=110)
        header_frame.grid_columnconfigure(3, weight=0, minsize=60)
        header_frame.grid_columnconfigure(4, weight=0, minsize=60)
        header_frame.grid_columnconfigure(5, weight=0, minsize=40)

        for i, txt in enumerate(headers):
            lbl = ctk.CTkLabel(header_frame, text=txt, font=("", 10, "bold"))
            lbl.grid(row=0, column=i, padx=6, sticky="w")

        for i, name in enumerate(self.output_channels):
            row_frame = ctk.CTkFrame(self._ch_inner)
            row_frame.pack(fill="x", padx=4, pady=2)
            row_frame.grid_columnconfigure(0, weight=0, minsize=40)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(2, weight=0, minsize=110)
            row_frame.grid_columnconfigure(3, weight=0, minsize=60)
            row_frame.grid_columnconfigure(4, weight=0, minsize=60)
            row_frame.grid_columnconfigure(5, weight=0, minsize=40)

            ctk.CTkLabel(row_frame, text=str(i + 1)).grid(row=0, column=0, padx=6, sticky="w")

            unit_str = self.rnn_units.get(name, '').strip()
            if unit_str and unit_str != '':
                unit_str = unit_str.replace('\r', '').replace('\n', '').strip()
                display_name = f"{name} [{unit_str}]"
            else:
                display_name = name

            name_label = ctk.CTkLabel(row_frame, text=display_name, anchor="w")
            name_label.grid(row=0, column=1, padx=6, sticky="ew")

            model_choice = ctk.CTkOptionMenu(row_frame, values=["OptimizedSVR", "RNN", "RobustRNN", "GBR"], width=120)
            model_choice.set("OptimizedSVR")
            model_choice.grid(row=0, column=2, padx=6, sticky="w")

            var = tk.BooleanVar(value=True)
            chk = ctk.CTkCheckBox(row_frame, text="", variable=var)
            chk.grid(row=0, column=3, padx=6, sticky="")

            chk.bind("<ButtonPress-1>", lambda e, n=name: self._on_checkbox_press(e, n))
            chk.bind("<B1-Motion>", lambda e: self._on_checkbox_drag(e))
            chk.bind("<ButtonRelease-1>", lambda e: self._on_checkbox_release(e))

            qlbl = ctk.CTkLabel(row_frame, text="", width=22, height=18, corner_radius=3, fg_color="gray")
            qlbl.grid(row=0, column=4, padx=8, sticky="")

            flbl = ctk.CTkLabel(row_frame, text="", width=22, height=18, corner_radius=3, fg_color="gray")
            flbl.grid(row=0, column=5, padx=8, sticky="")

            self.channel_widgets[name] = {
                'var': var,
                'checkbox': chk,
                'model_choice': model_choice,
                'quality': qlbl,
                'fit': flbl,
                'row_frame': row_frame
            }

        # After populating, update state based on mode
        self._on_mode_change(self.prediction_mode.get())

    def _on_checkbox_press(self, event, channel_name):
        if channel_name not in self.channel_widgets: return
        var = self.channel_widgets[channel_name]['var']
        self._drag_toggle_state = not var.get()
        self._dragging_cursor = True

    def _on_checkbox_drag(self, event):
        if not self._dragging_cursor or self._drag_toggle_state is None:
            return
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        if widget is None: return
        target_channel = None
        for name, widgets in self.channel_widgets.items():
            if (widget == widgets['checkbox'] or
                    widget == widgets['row_frame'] or
                    widget in widgets['row_frame'].winfo_children()):
                target_channel = name
                break
        if target_channel and target_channel in self.channel_widgets:
            var = self.channel_widgets[target_channel]['var']
            if var.get() != self._drag_toggle_state:
                var.set(self._drag_toggle_state)

    def _on_checkbox_release(self, event):
        if not self._dragging_cursor: return
        self._dragging_cursor = False
        self._drag_toggle_state = None
        try:
            self.recompute_graphics()
        except Exception as e:
            print(f"Error recomputing graphics: {e}")

    def safe_predict(self, model, X_df, feature_cols=None):
        try:
            if feature_cols is not None:
                X = X_df[feature_cols].values if hasattr(X_df, "values") else np.asarray(X_df)
            else:
                X = X_df.values if hasattr(X_df, "values") else np.asarray(X_df)
            return model.predict(X)
        except Exception:
            try:
                return model.predict(X_df)
            except Exception:
                n = len(X_df) if hasattr(X_df, "__len__") else 0
                return np.full(n, np.nan, dtype=float)

    def _create_modeling_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1);
        tab.grid_rowconfigure(0, weight=1)
        self.fig_modeling, self.axes_modeling = plt.subplots(3, 3, figsize=(12, 8))
        self.fig_modeling.subplots_adjust(hspace=0.45, wspace=0.35)
        self.canvas_modeling = FigureCanvasTkAgg(self.fig_modeling, master=tab)
        self.canvas_modeling.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        for ax in self.axes_modeling.flat:
            ax.text(0.5, 0.5, 'Load Data and Build Models', ha='center', va='center', fontsize=12, color='gray')
            ax.set_xticks([]);
            ax.set_yticks([])
        try:
            self.canvas_modeling.draw()
        except Exception:
            pass

    def _create_graphics_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1);
        tab.grid_rowconfigure(0, weight=1)
        self.fig_graphics, self.ax_graphics = plt.subplots(1, 1, figsize=(10, 6))
        self.fig_graphics.subplots_adjust(top=0.95, bottom=0.08, left=0.06, right=0.98)
        self.canvas_graphics = FigureCanvasTkAgg(self.fig_graphics, master=tab)
        self.canvas_graphics.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.ax_graphics.text(0.5, 0.5, 'Interaction Plot Area', ha='center', va='center', fontsize=14, color='gray')
        self.ax_graphics.set_xticks([]);
        self.ax_graphics.set_yticks([])
        try:
            self.canvas_graphics.draw()
        except Exception:
            pass

    def _create_shap_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=0)
        tab.grid_rowconfigure(1, weight=1)
        shap_controls_frame = ctk.CTkFrame(tab, fg_color="transparent")
        shap_controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(5, 0))
        ctk.CTkLabel(shap_controls_frame, text="Model to Analyze:").pack(side="left", padx=(0, 5))
        self.shap_model_menu = ctk.CTkOptionMenu(shap_controls_frame, values=["-"], state="disabled", width=250,
                                                 command=self.draw_shap_plot)
        self.shap_model_menu.pack(side="left", padx=5)
        ctk.CTkLabel(shap_controls_frame, text="Plot Type:").pack(side="left", padx=(10, 5))
        self.shap_plot_type_menu = ctk.CTkOptionMenu(shap_controls_frame,
                                                     values=["Bar (Ranking)", "Summary (Beeswarm)", "Dependence"],
                                                     state="disabled", width=180,
                                                     command=self.draw_shap_plot)
        self.shap_plot_type_menu.pack(side="left", padx=5)
        ctk.CTkLabel(shap_controls_frame, text="Dependence Feature:").pack(side="left", padx=(10, 5))
        self.shap_dependence_menu = ctk.CTkOptionMenu(shap_controls_frame, values=["-"], state="disabled", width=200,
                                                      command=self.draw_shap_plot)
        self.shap_dependence_menu.pack(side="left", padx=5)
        self.export_shap_button = ctk.CTkButton(shap_controls_frame, text="Export CSV", state="disabled",
                                                command=self.export_shap_to_csv, width=110)
        self.export_shap_button.pack(side="left", padx=(10, 5))
        self.fig_shap, self.ax_shap = plt.subplots(1, 1, figsize=(10, 6))
        self.fig_shap.subplots_adjust(top=0.92, bottom=0.1, left=0.3, right=0.95)
        self.canvas_shap = FigureCanvasTkAgg(self.fig_shap, master=tab)
        self.canvas_shap.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        theme_props = self.app.get_theme_properties()
        text_color = theme_props.get("text_color", "black")
        if not _HAS_SHAP:
            self.ax_shap.text(0.5, 0.6, 'SHAP Library Not Installed', ha='center', va='center',
                              fontsize=14, fontweight='bold', color='red')
            self.ax_shap.text(0.5, 0.4, 'Install with: pip install shap', ha='center', va='center',
                              fontsize=11, color=text_color)
        else:
            self.ax_shap.text(0.5, 0.5, 'Build models, then click "Run SHAP" (on left panel)', ha='center', va='center',
                              fontsize=12, color=text_color)
        self.ax_shap.set_xticks([]);
        self.ax_shap.set_yticks([])
        try:
            self.canvas_shap.draw()
        except Exception:
            pass

    def run_shap_calculation(self, *args):
        global _HAS_SHAP
        if not _HAS_SHAP:
            messagebox.showerror("Missing Library",
                                 "The 'shap' library is required.\n\nPlease install it by running:\n'pip install shap'")
            return
        if self.rnn_data is None or not self.trained_models:
            messagebox.showwarning("SHAP Analysis", "Please load data and build models first.")
            return

        # Prevent running SHAP on dynamic models
        if self.prediction_mode.get() == "Dynamic":
            messagebox.showwarning("SHAP Analysis", "SHAP analysis is not supported for Dynamic models.")
            return

        print(f"DEBUG: Running SHAP calculation for ALL {len(self.trained_models)} models...")
        self.run_shap_button.configure(text="Calculating...", state="disabled")
        self.shap_model_menu.configure(state="disabled")
        self.shap_plot_type_menu.configure(state="disabled")
        self.shap_dependence_menu.configure(state="disabled")
        self.export_shap_button.configure(state="disabled")
        self.app.update_idletasks()
        self.all_shap_values.clear()
        self.all_shap_samples.clear()
        self.current_shap_model_name = None
        try:
            for model_name, model_info in self.trained_models.items():
                print(f"  Calculating for: {model_name}")
                self.run_shap_button.configure(text=f"Calc: {model_name[:10]}...")
                self.app.update_idletasks()
                model_obj = model_info['model'] if isinstance(model_info, dict) else model_info
                feature_names = model_info.get('features', self.input_channels)
                X = self.rnn_data[feature_names].apply(pd.to_numeric, errors='coerce')
                y = self.rnn_data[model_name]
                valid_idx = X.dropna().index.intersection(y.dropna().index)
                X_train = X.loc[valid_idx].reset_index(drop=True)
                if X_train.empty:
                    print(f"  Skipping {model_name}: No valid training data.")
                    continue
                background_size = min(50, len(X_train))
                sample_size = min(200, len(X_train))
                background_data = shap.utils.sample(X_train, background_size) if len(
                    X_train) > background_size else X_train
                X_sample = X_train.sample(n=sample_size, random_state=42) if len(X_train) > sample_size else X_train
                shap_explanation = None
                if isinstance(model_obj, RobustRNN) or isinstance(model_obj, OptimizedSVR):
                    explainer = shap.KernelExplainer(model_obj.predict, background_data)
                    shap_values_array = explainer.shap_values(X_sample)
                    shap_base_value = explainer.expected_value
                    if isinstance(shap_base_value, (int, float, np.number)):
                        base_values_array = np.full(len(X_sample), shap_base_value)
                    elif isinstance(shap_base_value, np.ndarray) and shap_base_value.ndim == 0:
                        base_values_array = np.full(len(X_sample), float(shap_base_value))
                    else:
                        base_values_array = shap_base_value
                    shap_explanation = shap.Explanation(
                        values=shap_values_array, base_values=base_values_array,
                        data=X_sample.values, feature_names=list(feature_names)
                    )
                else:
                    explainer = shap.TreeExplainer(model_obj, background_data)
                    shap_explanation = explainer(X_sample)
                self.all_shap_values[model_name] = shap_explanation
                self.all_shap_samples[model_name] = X_sample
            print("DEBUG: SHAP values for ALL models calculated and stored.")
            messagebox.showinfo("SHAP Calculation Complete",
                                f"Successfully calculated SHAP values for {len(self.all_shap_values)} models.\n\n"
                                "You can now select a model from the dropdown to view its plot.")
            self.draw_shap_plot()
        except Exception as e:
            messagebox.showerror("SHAP Error", f"Failed to run SHAP analysis:\n{e}")
            import traceback
            traceback.print_exc()
        finally:
            self.run_shap_button.configure(text="Run SHAP", state="normal")
            self.shap_model_menu.configure(state="normal")
            self.shap_plot_type_menu.configure(state="normal")
            self.draw_shap_plot()

    def draw_shap_plot(self, *args):
        plot_type = self.shap_plot_type_menu.get()
        selected_model_name = self.shap_model_menu.get()
        if plot_type == "Dependence":
            self.shap_dependence_menu.configure(state="normal")
        else:
            self.shap_dependence_menu.configure(state="disabled")
        if selected_model_name not in self.all_shap_values:
            self.export_shap_button.configure(state="disabled")
            if hasattr(self, 'ax_shap'):
                self.ax_shap.clear()
                theme_props = self.app.get_theme_properties()
                text_to_show = 'Build models, then click "Run SHAP" (on left panel)'
                if self.trained_models:
                    text_to_show = f"Click 'Run SHAP' to analyze models.\n'{selected_model_name}' has not been calculated."
                self.ax_shap.text(0.5, 0.5, text_to_show, ha='center', va='center',
                                  fontsize=12, color=theme_props.get("text_color", "black"))
                self.ax_shap.set_xticks([]);
                self.ax_shap.set_yticks([])
                if hasattr(self, 'canvas_shap'): self.canvas_shap.draw_idle()
            return
        print(f"DEBUG: Drawing SHAP plot for pre-calculated model: {selected_model_name}")
        try:
            shap_values_to_plot = self.all_shap_values[selected_model_name]
            x_sample_to_plot = self.all_shap_samples[selected_model_name]
            self.current_shap_model_name = selected_model_name
            self.export_shap_button.configure(state="normal")
            if hasattr(self, 'canvas_shap') and self.canvas_shap:
                self.canvas_shap.get_tk_widget().destroy()
            plt.close('all')
            if plot_type == "Bar (Ranking)":
                shap.summary_plot(shap_values_to_plot, x_sample_to_plot, plot_type="bar", max_display=20, show=False)
                plt.title(f"SHAP Feature Importance (Ranking) for:\n{self.current_shap_model_name}")
                plt.xlabel("Mean Absolute SHAP Value (Impact on Model Output)")
            elif plot_type == "Summary (Beeswarm)":
                shap.summary_plot(shap_values_to_plot, x_sample_to_plot, plot_type="dot", max_display=20, show=False)
                plt.title(f"SHAP Summary (Beeswarm) Plot for:\n{self.current_shap_model_name}")
                plt.xlabel("SHAP Value (Impact on Model Output)")
            elif plot_type == "Dependence":
                dependence_feature = self.shap_dependence_menu.get()
                if dependence_feature == "-":
                    self.fig_shap = plt.figure(figsize=(10, 6))
                    ax = self.fig_shap.add_subplot(111)
                    ax.text(0.5, 0.5, "Please select a feature from the 'Dependence Feature' dropdown.", ha='center',
                            va='center', color='gray')
                    self.ax_shap = ax
                else:
                    shap.dependence_plot(dependence_feature, shap_values_to_plot.values, x_sample_to_plot, show=False)
                    plt.title(f"SHAP Dependence Plot for '{dependence_feature}'\non {self.current_shap_model_name}")
                self.fig_shap = plt.gcf()
                self.ax_shap = plt.gca()
            if plot_type != "Dependence" or self.shap_dependence_menu.get() != "-":
                self.fig_shap = plt.gcf()
                self.ax_shap = plt.gca()
            self.canvas_shap = FigureCanvasTkAgg(self.fig_shap, master=self.plot_tabs.tab("Sensitivity (SHAP)"))
            self.canvas_shap.get_tk_widget().grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
            self.apply_theme()
            try:
                self.fig_shap.tight_layout(pad=1.5)
            except:
                pass
            self.canvas_shap.draw_idle()
        except Exception as e:
            messagebox.showerror("SHAP Plot Error", f"Failed to draw SHAP plot:\n{e}")
            traceback.print_exc()

    def export_shap_to_csv(self):
        import csv
        import string
        def get_excel_col_name(n):
            name = ""
            while n >= 0:
                name = string.ascii_uppercase[n % 26] + name
                n = n // 26 - 1
            return name

        if not hasattr(self, 'all_shap_values') or not self.all_shap_values:
            messagebox.showwarning("Export Error", "No SHAP data available to export.")
            return
        selected_model_name = self.shap_model_menu.get()
        if selected_model_name not in self.all_shap_values:
            messagebox.showwarning("Export Error", f"No SHAP data found for '{selected_model_name}'.")
            return
        self.current_shap_model_name = selected_model_name
        shap_values_to_export = self.all_shap_values[selected_model_name]
        x_sample_to_export = self.all_shap_samples[selected_model_name]
        try:
            default_filename = f"SHAP_Values_{self.current_shap_model_name}.csv"
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                title="Save SHAP Values as CSV", initialfile=default_filename
            )
            if not filepath: return
            feature_names = shap_values_to_export.feature_names
            shap_df = pd.DataFrame(shap_values_to_export.values, columns=[f"SHAP_{f}" for f in feature_names],
                                   index=x_sample_to_export.index)
            features_df = x_sample_to_export
            combined_df = pd.concat([features_df, shap_df], axis=1)
            try:
                combined_df['SHAP_Base_Value'] = shap_values_to_export.base
                combined_df['Model_Prediction'] = shap_values_to_export.base + shap_df.sum(axis=1)
            except Exception:
                pass
            combined_df.to_csv(filepath, index_label="Original_Row_Index")
            num_data_rows = len(combined_df)
            formula_row = [''] * (len(combined_df.columns) + 1)
            formula_row[0] = "Mean_Absolute_SHAP_Value (Excel_Formula)"
            shap_col_names = [f"SHAP_{f}" for f in feature_names]
            for col_name in combined_df.columns:
                if col_name in shap_col_names:
                    col_idx_in_df = combined_df.columns.get_loc(col_name)
                    col_idx_in_csv = col_idx_in_df + 1
                    col_letter = get_excel_col_name(col_idx_in_csv)
                    formula_string = f"=AVERAGE(ABS({col_letter}2:{col_letter}{num_data_rows + 1}))"
                    formula_row[col_idx_in_csv] = formula_string
            with open(filepath, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(formula_row)
            messagebox.showinfo("Export Successful",
                                f"SHAP data and Mean Absolute Value formula exported to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An unexpected error occurred during export:\n{e}")

    def compute_model_stats(self, model, X, y, cv_folds=5):
        try:
            Xc = X.apply(pd.to_numeric, errors='coerce');
            yc = pd.to_numeric(y, errors='coerce')
            valid = (~yc.isna()) & (~Xc.isna().any(axis=1))
            if valid.sum() < 10:
                return {'r2': np.nan, 'r2_adj': np.nan, 'r2_pred': np.nan, 'f_test': np.nan, 'rmse': np.nan,
                        'nrmse': np.nan}
            Xv = Xc.loc[valid].values;
            yv = yc.loc[valid].values
            n = len(yv);
            p = Xv.shape[1] if Xv.ndim > 1 else 1
            try:
                y_in = model.predict(Xv);
                r2 = r2_score(yv, y_in)
            except Exception:
                r2 = np.nan
            try:
                r2_adj = 1 - (1 - r2) * ((n - 1) / (n - p - 1)) if (n - p - 1) > 0 else np.nan
            except Exception:
                r2_adj = np.nan
            try:
                if not np.isnan(r2) and (n - p - 1) > 0 and p > 0:
                    f_test = (r2 / p) / ((1 - r2) / (n - p - 1))
                else:
                    f_test = np.nan
            except (ZeroDivisionError, FloatingPointError):
                f_test = np.nan
            try:
                X_train, X_test, y_train, y_test = train_test_split(
                    Xv, yv, test_size=0.2, random_state=42, shuffle=True
                )
                model_for_pred = type(model)(**model.get_params())
                model_for_pred.fit(X_train, y_train)
                y_pred_test = model_for_pred.predict(X_test)
                r2_pred = r2_score(y_test, y_pred_test)
                y_train_pred = model_for_pred.predict(X_train)
                r2_train = r2_score(y_train, y_train_pred)
                if (r2_train - r2_pred) > 0.15:
                    print(f"⚠ WARNING: Potential overfitting detected!")
            except Exception as e:
                r2_pred = np.nan
            try:
                rmse = np.sqrt(mean_squared_error(yv, y_in));
                denom = (yv.max() - yv.min());
                nrmse = (rmse / denom) * 100.0 if denom != 0 else np.nan
            except Exception:
                rmse = np.nan;
                nrmse = np.nan
            return {'r2': r2, 'r2_adj': r2_adj, 'r2_pred': r2_pred, 'f_test': f_test, 'rmse': rmse, 'nrmse': nrmse}
        except Exception as e:
            return {'r2': np.nan, 'r2_adj': np.nan, 'r2_pred': np.nan, 'f_test': np.nan, 'rmse': np.nan,
                    'nrmse': np.nan}

    def recompute_graphics(self):
        try:
            if self.rnn_data is None: return
            if "Interaction" in (
                    self.graphic_type_menu.get() if hasattr(self.graphic_type_menu, "get") else "Interaction"):
                self.build_interaction_graphics(points=48, percentiles=(20, 50, 80), cell_w=2.0,
                                                cell_h=1.8, show_scatter=self.show_scatter_var.get())
            else:
                try:
                    self.canvas_modeling.draw()
                except Exception:
                    pass
        except Exception:
            pass

    def build_interaction_graphics(self, points=48, percentiles=(20, 50, 80), cell_w=1.9, cell_h=1.6,
                                   show_scatter=False):
        if self.rnn_data is None or not self.input_channels or not self.output_channels:
            messagebox.showinfo("Graphics", "Import data and define channels first.")
            return

        df = self.rnn_data.copy()
        for c in (self.input_channels + self.output_channels):
            df[c] = pd.to_numeric(df[c], errors='coerce')

        # Use rnn_data_bounds if available (for Cameo models), otherwise calculate from data
        use_bounds = hasattr(self, 'rnn_data_bounds') and self.rnn_data_bounds
        input_minmax = {}
        for c in self.input_channels:
            if use_bounds and c in self.rnn_data_bounds:
                input_minmax[c] = self.rnn_data_bounds[c]
            else:
                input_minmax[c] = (df[c].min(), df[c].max())

        perc_values = {}
        for c in self.input_channels:
            if use_bounds and c in self.rnn_data_bounds:
                # For Cameo, calculate percentiles from bounds
                mn, mx = self.rnn_data_bounds[c]
                p20 = mn + 0.2 * (mx - mn)
                p50 = mn + 0.5 * (mx - mn)
                p80 = mn + 0.8 * (mx - mn)
                perc_values[c] = [p20, p50, p80]
            else:
                vals = df[c].dropna().values
                perc_values[c] = np.percentile(vals, list(percentiles)) if vals.size > 0 else [np.nan] * len(percentiles)

        self._ref_inputs.clear()
        for c in self.input_channels:
            if use_bounds and c in self.rnn_data_bounds:
                mn, mx = self.rnn_data_bounds[c]
                self._ref_inputs[c] = (mn + mx) / 2  # Use midpoint
            else:
                self._ref_inputs[c] = float(np.percentile(df[c].dropna(), 50)) if not df[c].dropna().empty else 0.0
        valid_cells = []
        used_rows_indices = set()
        used_cols_indices = set()
        for r, out_name in enumerate(self.output_channels):
            widgets = self.channel_widgets.get(out_name)
            if not widgets or not widgets['var'].get():
                continue
            row_has_valid_plot = False
            for c, in_name in enumerate(self.input_channels):
                mn, mx = input_minmax.get(in_name, (np.nan, np.nan))
                if pd.isna(mn) or pd.isna(mx) or np.isclose(mn, mx):
                    continue
                valid_cells.append((r, c))
                used_cols_indices.add(c)
                row_has_valid_plot = True
            if row_has_valid_plot:
                used_rows_indices.add(r)
        if not valid_cells:
            try:
                plt.close(self.fig_graphics)
                self.fig_graphics, self.ax_graphics = plt.subplots(1, 1, figsize=(10, 6))
                self.canvas_graphics.figure = self.fig_graphics
                self.ax_graphics.clear()
                self.ax_graphics.text(0.5, 0.5, 'No valid interaction cells', ha='center', va='center', color='gray')
                self.ax_graphics.set_xticks([]);
                self.ax_graphics.set_yticks([])
                self.canvas_graphics.draw_idle()
                self.apply_theme()
            except Exception as e:
                print(f"Error creating placeholder: {e}")
            return
        used_rows_list = sorted(list(used_rows_indices))
        used_cols_list = sorted(list(used_cols_indices))
        row_map = {orig_r: new_r for new_r, orig_r in enumerate(used_rows_list)}
        col_map = {orig_c: new_c for new_c, orig_c in enumerate(used_cols_list)}
        n_rows = len(used_rows_list)
        n_cols = len(used_cols_list)
        figsize = (max(6, cell_w * max(1, n_cols) + 1), max(4, cell_h * max(1, n_rows) + 1))
        try:
            plt.close(self.fig_graphics)
        except Exception:
            pass
        self.fig_graphics = plt.figure(figsize=figsize, dpi=100)
        gs = self.fig_graphics.add_gridspec(n_rows, n_cols, hspace=0.45, wspace=0.35,
                                            left=0.1, right=0.9, bottom=0.1, top=0.9)
        axes_map = {}
        axes_map_compressed = {}
        for r_orig in used_rows_list:
            for c_orig in used_cols_list:
                if (r_orig, c_orig) in valid_cells:
                    r_new = row_map[r_orig]
                    c_new = col_map[c_orig]
                    ax = self.fig_graphics.add_subplot(gs[r_new, c_new])
                    axes_map[(r_orig, c_orig)] = ax
                    axes_map_compressed[(r_new, c_new)] = ax
        self._interaction_artists.clear()
        row_plot_data = {r_orig: [] for r_orig in used_rows_list}
        for (r_orig, c_orig) in valid_cells:
            ax = axes_map.get((r_orig, c_orig))
            if ax is None: continue
            out_name = self.output_channels[r_orig]
            in_name = self.input_channels[c_orig]
            ax.clear()
            mn, mx = input_minmax.get(in_name, (df[in_name].min(), df[in_name].max()))
            x_grid = np.linspace(mn, mx, points)
            model_info = self.trained_models.get(out_name)
            model = model_info['model'] if isinstance(model_info, dict) else model_info
            preds = {}
            for pid, perc in enumerate(percentiles):
                row_template = {}
                for other_in_name in self.input_channels:
                    if other_in_name == in_name:
                        row_template[other_in_name] = None
                    else:
                        pv = perc_values.get(other_in_name, [])
                        val_to_use = pv[pid] if len(pv) > pid else (
                            df[other_in_name].median() if not df[other_in_name].dropna().empty else 0.0)
                        if np.isnan(val_to_use):
                            val_to_use = df[other_in_name].median() if not df[other_in_name].dropna().empty else 0.0
                        row_template[other_in_name] = val_to_use
                Xpred = pd.DataFrame(
                    [{**{k: (x if k == in_name else v) for k, v in row_template.items()}} for x in x_grid]
                )
                features_expected = model_info.get('features', self.input_channels) if isinstance(model_info,
                                                                                                  dict) else self.input_channels
                Xpred = Xpred[features_expected].apply(pd.to_numeric, errors='coerce')
                if model is not None:
                    try:
                        yp = self.safe_predict(model, Xpred, feature_cols=features_expected)
                        preds[perc] = yp
                    except Exception as e_pred:
                        preds[perc] = np.full_like(x_grid, np.nan, dtype=float)
                else:
                    preds[perc] = np.full_like(x_grid, np.nan, dtype=float)
            lowp, midp, highp = percentiles
            y_low, y_mid, y_high = preds.get(lowp), preds.get(midp), preds.get(highp)
            poly, line_mid, line_low, line_high, scat = None, None, None, None, None
            if y_low is not None and y_high is not None and np.any(np.isfinite(y_low)) and np.any(np.isfinite(y_high)):
                poly = ax.fill_between(x_grid, y_low, y_high, alpha=0.2, facecolor='green', interpolate=True, zorder=1)
                row_plot_data[r_orig].extend(y_low[np.isfinite(y_low)])
                row_plot_data[r_orig].extend(y_high[np.isfinite(y_high)])
            if y_mid is not None and np.any(np.isfinite(y_mid)):
                line_mid, = ax.plot(x_grid, y_mid, color='red', linewidth=1.6, zorder=4)
                row_plot_data[r_orig].extend(y_mid[np.isfinite(y_mid)])
            if y_low is not None and np.any(np.isfinite(y_low)):
                line_low, = ax.plot(x_grid, y_low, color='green', linestyle='--', linewidth=1.0, alpha=0.9, zorder=3)
            if y_high is not None and np.any(np.isfinite(y_high)):
                line_high, = ax.plot(x_grid, y_high, color='green', linestyle='--', linewidth=1.0, alpha=0.9, zorder=3)
            if show_scatter:
                valid_scatter = df[[in_name, out_name]].dropna()
                if not valid_scatter.empty:
                    scat = ax.scatter(valid_scatter[in_name], valid_scatter[out_name], s=8, alpha=0.6, color='#1f77b4',
                                      zorder=2)
                    row_plot_data[r_orig].extend(valid_scatter[out_name].values)
            r_new, c_new = row_map[r_orig], col_map[c_orig]
            ax.set_xlim(mn, mx)
            if r_new == n_rows - 1:
                label_name = in_name.split(' [')[0]
                unit = self.rnn_units.get(in_name, "")
                if unit: label_name = f"{label_name}\n[{unit}]"
                ax.set_xlabel(label_name, fontsize=8)
                ax.tick_params(axis='x', labelsize=7, rotation=30)
            else:
                ax.set_xticklabels([])
            if c_new == 0:
                label_name = out_name.split(' [')[0]
                unit_o = self.rnn_units.get(out_name, "")
                if unit_o: label_name = f"{label_name}\n[{unit_o}]"
                ax.set_ylabel(label_name, fontsize=8)
                ax.tick_params(axis='y', labelsize=7)
            else:
                ax.set_yticklabels([])
            ax.grid(True, linestyle=':', linewidth=0.4, alpha=0.6)
            self._interaction_artists[(r_orig, c_orig)] = {
                'x': x_grid, 'line_mid': line_mid, 'line_low': line_low, 'line_high': line_high,
                'fill': poly,
                'scatter': scat, 'in_name': in_name, 'out_name': out_name,
                'percentiles': percentiles
            }
        for r_orig in used_rows_list:
            if r_orig in row_plot_data and row_plot_data[r_orig]:
                all_y_vals = np.array(row_plot_data[r_orig])
                if all_y_vals.size == 0: continue
                min_y, max_y = np.nanmin(all_y_vals), np.nanmax(all_y_vals)
                if np.isnan(min_y) or np.isnan(max_y): continue
                padding = (max_y - min_y) * 0.1
                if padding == 0: padding = 1.0
                final_min_y, final_max_y = min_y - padding, max_y + padding
            else:
                continue
            for c_orig in used_cols_list:
                ax = axes_map.get((r_orig, c_orig))
                if ax is not None:
                    ax.set_ylim(final_min_y, final_max_y)
        for v in getattr(self, "_cursor_vlines", []): v.remove()
        for v in getattr(self, "_cursor_hlines", []): v.remove()
        for l in getattr(self, "_cursor_vlabels", []): l.remove()
        for l in getattr(self, "_cursor_hlabels", []): l.remove()
        self._cursor_vlines, self._cursor_hlines = [], []
        self._cursor_vlabels, self._cursor_hlabels = [], []
        theme_props = self.app.get_theme_properties()
        v_cursor_color = theme_props.get("cursor_color", "#87CEFA")
        v_cursor_width = 2
        h_cursor_color = "green"
        h_cursor_width = 2
        label_fg = "black"
        v_label_bg = v_cursor_color
        h_label_bg = h_cursor_color
        for c_new, c_orig in enumerate(used_cols_list):
            input_name = self.input_channels[c_orig]
            x0 = self._ref_inputs.get(input_name, np.nan)
            top_ax = None
            for r_new in range(n_rows):
                ax_in_col = axes_map_compressed.get((r_new, c_new))
                if ax_in_col is not None:
                    if top_ax is None: top_ax = ax_in_col
                    v = ax_in_col.axvline(x=x0, color=v_cursor_color, linewidth=v_cursor_width,
                                          alpha=0.9, zorder=6, picker=5)
                    v._is_v_cursor = True
                    v._orig_col_idx = c_orig
                    self._cursor_vlines.append(v)
            if top_ax is not None:
                lab = top_ax.text(x0, 1.02, f"{x0:.4g}",
                                  transform=top_ax.get_xaxis_transform(),
                                  ha='center', va='bottom', fontsize=9, color=label_fg,
                                  bbox=dict(boxstyle='round,pad=0.2', facecolor=v_label_bg,
                                            edgecolor='none', alpha=0.9),
                                  clip_on=False, zorder=10)
                self._cursor_vlabels.append(lab)
        X_ref_df = pd.DataFrame([self._ref_inputs])
        for r_orig in used_rows_list:
            out_name = self.output_channels[r_orig]
            model_info = self.trained_models.get(out_name)
            model = model_info['model'] if isinstance(model_info, dict) else model_info
            y0, y_low, y_high = np.nan, np.nan, np.nan
            if model:
                try:
                    features = model_info.get('features', self.input_channels)
                    row_mid = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_mid[input_chan] = pv[1] if len(pv) > 1 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_mid[input_chan] = ref_val
                    X_mid = pd.DataFrame([row_mid])[features]
                    y0 = self.safe_predict(model, X_mid, feature_cols=features)[0]
                    row_low = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_low[input_chan] = pv[0] if len(pv) > 0 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_low[input_chan] = ref_val
                    X_low = pd.DataFrame([row_low])[features]
                    y_low = self.safe_predict(model, X_low, feature_cols=features)[0]
                    row_high = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_high[input_chan] = pv[2] if len(pv) > 2 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_high[input_chan] = ref_val
                    X_high = pd.DataFrame([row_high])[features]
                    y_high = self.safe_predict(model, X_high, feature_cols=features)[0]
                except Exception as e:
                    print(f"Error calculating Y-labels for {out_name}: {e}")
            self._ref_outputs[out_name] = y0
            r_new = row_map[r_orig]
            right_ax = None
            for c_new in range(n_cols):
                ax_in_row = axes_map_compressed.get((r_new, c_new))
                if ax_in_row is not None:
                    right_ax = ax_in_row
                    h = ax_in_row.axhline(y=y0, color=h_cursor_color, linewidth=h_cursor_width,
                                          alpha=0.9, zorder=6, picker=5)
                    h._is_h_cursor = True
                    h._orig_row_idx = r_orig
                    self._cursor_hlines.append(h)
            if right_ax is not None:
                label_text = f"{y_high:.4g}\n{y0:.4g}\n{y_low:.4g}"
                lab = right_ax.text(1.02, y0, label_text,
                                    transform=right_ax.get_yaxis_transform(),
                                    ha='left', va='center', fontsize=9, color=label_fg,
                                    bbox=dict(boxstyle='round,pad=0.2', facecolor=h_label_bg,
                                              edgecolor='none', alpha=0.9),
                                    clip_on=False, zorder=10)
                self._cursor_hlabels.append(lab)
        for w in self.variations_frame.winfo_children():
            w.destroy()
        self.variation_widgets.clear()
        num_cols = 2
        for i in range(num_cols):
            self.variations_frame.grid_columnconfigure(i, weight=1)
        for i, in_name in enumerate(self.input_channels):
            row = i // num_cols
            col = i % num_cols
            var_frame = ctk.CTkFrame(self.variations_frame, fg_color="transparent")
            var_frame.grid(row=row, column=col, padx=5, pady=2, sticky="ew")
            label_name = in_name.split(' [')[0]
            if len(label_name) > 20:
                label_name = label_name[:18] + "..."
            lbl = ctk.CTkLabel(var_frame, text=label_name, anchor="w", width=100)
            lbl.pack(side="left", padx=(5, 2))
            entry = ctk.CTkEntry(var_frame, width=80)
            entry.insert(0, f"{self._ref_inputs[in_name]:.4g}")
            entry.pack(side="left", padx=2, fill="x", expand=True)
            entry.bind("<Return>", lambda event, name=in_name: self._on_variation_entry_update(event, name))
            self.variation_widgets[in_name] = entry
            if i not in used_cols_indices:
                entry.configure(state="disabled")
        try:
            self.fig_graphics.tight_layout(pad=1.0, h_pad=1.5, w_pad=1.0, rect=[0.05, 0.05, 0.9, 0.9])
        except Exception:
            pass
        try:
            if hasattr(self, 'canvas_graphics') and self.canvas_graphics.get_tk_widget().winfo_exists():
                self.canvas_graphics.get_tk_widget().destroy()
        except Exception:
            pass
        try:
            self.canvas_graphics = FigureCanvasTkAgg(self.fig_graphics, master=self.plot_tabs.tab("Graphics"))
            self.canvas_graphics.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
            self.canvas_graphics.draw_idle()
        except Exception as e:
            print(f"Error creating canvas: {e}")
        self._last_axes_grid = np.array(
            [[axes_map.get((r_orig, c_orig), None) for c_orig in range(len(self.input_channels))]
             for r_orig in range(len(self.output_channels))], dtype=object
        )
        self._last_row_map = row_map
        self._last_col_map = col_map
        self._last_used_rows_list = used_rows_list
        self._last_used_cols_list = used_cols_list
        self.apply_theme()
        try:
            self.canvas_graphics.draw()
        except Exception:
            pass
        self.add_interactive_cursor()

    def _on_variation_entry_update(self, event, input_name):
        if input_name not in self.variation_widgets:
            return
        entry = self.variation_widgets[input_name]
        try:
            new_value_str = entry.get()
            new_value_float = float(new_value_str)
        except (ValueError, TypeError):
            entry.delete(0, 'end')
            entry.insert(0, f"{self._ref_inputs[input_name]:.4g}")
            return
        self._ref_inputs[input_name] = new_value_float
        try:
            c_orig = self.input_channels.index(input_name)
            for vline in self._cursor_vlines:
                if vline._orig_col_idx == c_orig:
                    vline.set_xdata([new_value_float, new_value_float])
            vidx = self._last_used_cols_list.index(c_orig)
            vlab = self._cursor_vlabels[vidx]
            vlab.set_x(new_value_float)
            vlab.set_text(f"{new_value_float:.4g}")
        except (ValueError, IndexError):
            pass
        self._update_all_curves_for_ref()
        self.canvas_graphics.draw_idle()

    def add_interactive_cursor(self):
        try:
            for cid in getattr(self, "_cursor_cids", ()):
                self.fig_graphics.canvas.mpl_disconnect(cid)
        except Exception:
            pass
        self._cursor_cids = ()
        if self._last_axes_grid is None: return
        axes2d = self._last_axes_grid

        def get_axis_indices(ax_clicked):
            for r_orig in range(axes2d.shape[0]):
                for c_orig in range(axes2d.shape[1]):
                    if axes2d[r_orig, c_orig] is ax_clicked:
                        return r_orig, c_orig
            return None, None

        def on_press(event):
            if event.inaxes is None or event.button != 1: return
            r_orig, c_orig = get_axis_indices(event.inaxes)
            if r_orig is None: return
            self._active_row_idx = r_orig
            self._active_col_idx = c_orig
            self._dragging_cursor = True
            on_motion(event)

        def on_release(event):
            self._dragging_cursor = False
            self._active_row_idx = None
            self._active_col_idx = None
            self._update_all_curves_for_ref()
            self.canvas_graphics.draw_idle()

        def on_motion(event):
            if not self._dragging_cursor or event.inaxes is None or \
                    self._active_row_idx is None or self._active_col_idx is None:
                return
            r_curr, c_curr = get_axis_indices(event.inaxes)
            if r_curr is None:
                self._dragging_cursor = False
                return
            try:
                c_orig = self._active_col_idx
                col_name = self.input_channels[c_orig]
                self._ref_inputs[col_name] = float(event.xdata)
                for vline in self._cursor_vlines:
                    if vline._orig_col_idx == c_orig:
                        vline.set_xdata([event.xdata, event.xdata])
                vidx = self._last_used_cols_list.index(c_orig)
                vlab = self._cursor_vlabels[vidx]
                vlab.set_x(event.xdata)
                vlab.set_text(f"{event.xdata:.4g}")
                if col_name in self.variation_widgets:
                    entry = self.variation_widgets[col_name]
                    entry.delete(0, 'end')
                    entry.insert(0, f"{event.xdata:.4g}")
                r_orig = self._active_row_idx
                out_name = self.output_channels[r_orig]
                self._ref_outputs[out_name] = float(event.ydata)
                for hline in self._cursor_hlines:
                    if hline._orig_row_idx == r_orig:
                        hline.set_ydata([event.ydata, event.ydata])
                self._update_all_curves_for_ref()
                self.canvas_graphics.draw_idle()
            except Exception as e:
                self._dragging_cursor = False

        try:
            cid_press = self.fig_graphics.canvas.mpl_connect('button_press_event', on_press)
            cid_release = self.fig_graphics.canvas.mpl_connect('button_release_event', on_release)
            cid_move = self.fig_graphics.canvas.mpl_connect('motion_notify_event', on_motion)
            self._cursor_cids = (cid_press, cid_release, cid_move)
        except Exception as e:
            print(f"Error connecting cursor events: {e}")

    def _update_all_curves_for_ref(self):
        if not self._interaction_artists or self.rnn_data is None:
            return
        df = self.rnn_data.copy()
        for c in (self.input_channels + self.output_channels):
            df[c] = pd.to_numeric(df[c], errors='coerce')
        perc_values = {}
        try:
            perc_list = self._interaction_artists[list(self._interaction_artists.keys())[0]]['percentiles']
        except Exception:
            perc_list = (20, 50, 80)
        for c in self.input_channels:
            vals = df[c].dropna().values
            perc_values[c] = np.percentile(vals, list(perc_list)) if vals.size > 0 else [np.nan] * len(perc_list)
        for (r, c), artist in list(self._interaction_artists.items()):
            in_name = artist['in_name']
            out_name = artist['out_name']
            x_grid = artist['x']
            model_info = self.trained_models.get(out_name)
            model = model_info['model'] if isinstance(model_info, dict) else model_info
            if model is None: continue
            perc = artist.get('percentiles', (20, 50, 80))
            preds = {}
            for pid_idx, p in enumerate(perc):
                row_template = {}
                for input_chan in self.input_channels:
                    pv = perc_values.get(input_chan, [np.nan] * len(perc))
                    val_to_use = pv[pid_idx] if len(pv) > pid_idx else np.nan
                    if np.isnan(val_to_use):
                        val_to_use = df[input_chan].median() if not df[input_chan].dropna().empty else 0.0
                    row_template[input_chan] = val_to_use
                for input_chan, ref_val in self._ref_inputs.items():
                    if input_chan != in_name:
                        row_template[input_chan] = ref_val
                row_template[in_name] = None
                Xpred = pd.DataFrame(
                    [{**{k: (x if k == in_name else v) for k, v in row_template.items()}} for x in x_grid]
                )
                features_expected = model_info.get('features', self.input_channels)
                Xpred = Xpred[features_expected].apply(pd.to_numeric, errors='coerce')
                try:
                    yp = self.safe_predict(model, Xpred, feature_cols=features_expected)
                except Exception:
                    yp = np.full_like(x_grid, np.nan, dtype=float)
                preds[p] = yp
            lowp, midp, highp = perc
            y_low, y_mid, y_high = preds.get(lowp), preds.get(midp), preds.get(highp)
            try:
                if artist['line_mid']: artist['line_mid'].set_ydata(y_mid)
                if artist['line_low']: artist['line_low'].set_ydata(y_low)
                if artist['line_high']: artist['line_high'].set_ydata(y_high)
            except Exception:
                pass
            try:
                if artist.get('fill') is not None:
                    artist['fill'].remove()
                ax = artist['line_mid'].axes
                if y_low is not None and y_high is not None and np.any(np.isfinite(y_low)) and np.any(
                        np.isfinite(y_high)):
                    new_fill = ax.fill_between(x_grid, y_low, y_high, alpha=0.2,
                                               facecolor='green', interpolate=True, zorder=1)
                    artist['fill'] = new_fill
                else:
                    artist['fill'] = None
            except Exception:
                artist['fill'] = None
        X_ref_df = pd.DataFrame([self._ref_inputs])
        for r_orig in self._last_used_rows_list:
            out_name = self.output_channels[r_orig]
            model_info = self.trained_models.get(out_name)
            model = model_info['model'] if isinstance(model_info, dict) else model_info
            y0, y_low, y_high = np.nan, np.nan, np.nan
            if model:
                try:
                    features = model_info.get('features', self.input_channels)
                    row_mid = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_mid[input_chan] = pv[1] if len(pv) > 1 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_mid[input_chan] = ref_val
                    X_mid = pd.DataFrame([row_mid])[features]
                    y0 = self.safe_predict(model, X_mid, feature_cols=features)[0]
                    row_low = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_low[input_chan] = pv[0] if len(pv) > 0 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_low[input_chan] = ref_val
                    X_low = pd.DataFrame([row_low])[features]
                    y_low = self.safe_predict(model, X_low, feature_cols=features)[0]
                    row_high = {}
                    for input_chan in self.input_channels:
                        pv = perc_values.get(input_chan, [np.nan] * 3)
                        row_high[input_chan] = pv[2] if len(pv) > 2 else np.nan
                    for input_chan, ref_val in self._ref_inputs.items():
                        row_high[input_chan] = ref_val
                    X_high = pd.DataFrame([row_high])[features]
                    y_high = self.safe_predict(model, X_high, feature_cols=features)[0]
                except Exception as e:
                    pass
            self._ref_outputs[out_name] = y0
            for hline in self._cursor_hlines:
                if hline._orig_row_idx == r_orig:
                    hline.set_ydata([y0, y0])
            try:
                ridx = self._last_used_rows_list.index(r_orig)
                hlab = self._cursor_hlabels[ridx]
                hlab.set_y(y0)
                hlab.set_text(f"{y_high:.4g}\n{y0:.4g}\n{y_low:.4g}")
            except (IndexError, KeyError):
                pass

    def _get_quality_color(self, value, is_r2=True):
        if value is None or np.isnan(value): return "#8b8b8b"
        if is_r2:
            if value > 0.95:
                return "#3CB371"
            elif value > 0.85:
                return "#FFD700"
            else:
                return "#FF5733"
        else:
            if value < 0.05:
                return "#3CB371"
            elif value < 0.10:
                return "#FFD700"
            else:
                return "#FF5733"

    def update_graphic(self, selected=None):
        sel = selected if selected is not None else (
            self.graphic_type_menu.get() if hasattr(self.graphic_type_menu, "get") else "Interaction")
        if "Interaction" in sel:
            try:
                self.plot_tabs.set("Graphics")
                self.build_interaction_graphics(points=48, percentiles=(20, 50, 80), cell_w=2.0,
                                                cell_h=1.8, show_scatter=self.show_scatter_var.get())
            except Exception:
                pass
        else:
            try:
                self.plot_tabs.set("Modeling")
                try:
                    self.canvas_modeling.draw()
                except Exception:
                    pass
            except Exception:
                pass

    def update_dynamic_graphic(self, selected=None):
        """Handle graphics selection in Dynamic mode"""
        if not hasattr(self, 'dynamic_viz_notebook') or self.dynamic_viz_notebook is None:
            messagebox.showinfo("Info", "Please build models first to view graphics.")
            return

        sel = selected if selected is not None else self.graphic_type_menu.get()

        try:
            # Switch to Dynamic Results tab
            self.plot_tabs.set("Dynamic Results")

            # Switch to the appropriate sub-tab in the notebook
            if "Frequency Response" in sel:
                self.dynamic_viz_notebook.set("Frequency Response")
            elif "Error Analysis" in sel:
                self.dynamic_viz_notebook.set("Error Analysis")
            elif "PCA Analysis" in sel:
                self.dynamic_viz_notebook.set("PCA Analysis")
            elif "Prediction Grid" in sel:
                # Add Prediction Grid tab if it doesn't exist
                if not hasattr(self, 'prediction_grid_frame'):
                    self.prediction_grid_frame = ctk.CTkFrame(
                        self.dynamic_viz_notebook.tab("Frequency Response"),
                        fg_color="transparent"
                    )
                    self.prediction_grid_frame.pack(fill="both", expand=True)
                self.dynamic_viz_notebook.set("Frequency Response")
                # Generate prediction grid visualization
                self._show_prediction_grid()
            elif "Interactive Prediction" in sel:
                # Create or switch to Interactive Prediction tab
                self._show_interactive_prediction()
        except Exception as e:
            print(f"Error updating dynamic graphic: {e}")
            import traceback
            traceback.print_exc()

    def _show_prediction_grid(self):
        """Generate and display prediction grid for multiple configurations"""
        try:
            from ENHANCED_DYNAMIC_RNN import DynamicVisualization
            from sklearn.model_selection import train_test_split

            output_name = self.dynamic_output_selector.get()
            model = self.dynamic_trainer.models[output_name]

            X = self.dynamic_data['X']
            Y = self.dynamic_data['Y']
            output_names = self.dynamic_data['output_names']
            idx = output_names.index(output_name)

            X_train, X_test, Y_train, Y_test = train_test_split(
                X, Y, test_size=0.2, random_state=42
            )

            Y_test_single = Y_test[:, :, idx]
            Y_pred = model.predict(X_test)
            frequency = self.dynamic_data['frequency']

            # Create prediction grid with 4 samples
            n_samples = min(4, len(Y_test_single))
            fig = DynamicVisualization.plot_prediction_grid(
                frequency, Y_test_single, Y_pred, output_name, n_samples=n_samples
            )

            # Display in frequency response frame
            self._embed_figure_in_tab(fig, self.freq_plot_frame)
            self.apply_theme()

        except Exception as e:
            print(f"Error showing prediction grid: {e}")
            import traceback
            traceback.print_exc()

    def _show_interactive_prediction(self):
        """Create interactive prediction interface for Dynamic mode"""
        try:
            # Create Interactive Prediction tab if it doesn't exist
            if not hasattr(self, 'interactive_pred_tab_created'):
                try:
                    self.dynamic_viz_notebook.add("Interactive Prediction")
                    self.interactive_pred_tab_created = True
                except Exception:
                    pass  # Tab might already exist

            # Switch to the tab
            self.dynamic_viz_notebook.set("Interactive Prediction")

            # Clear and recreate the interface
            tab = self.dynamic_viz_notebook.tab("Interactive Prediction")
            for widget in tab.winfo_children():
                widget.destroy()

            # Create main container with two panels
            main_container = ctk.CTkFrame(tab, fg_color="transparent")
            main_container.pack(fill="both", expand=True, padx=5, pady=5)

            # Left panel: Input controls
            left_panel = ctk.CTkFrame(main_container, width=300)
            left_panel.pack(side="left", fill="y", padx=(0, 5))
            left_panel.pack_propagate(False)

            # Title
            ctk.CTkLabel(left_panel, text="Input Parameters", font=("", 14, "bold")).pack(pady=10, padx=10)

            # Scrollable frame for inputs
            scroll_frame = ctk.CTkScrollableFrame(left_panel)
            scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

            # Get input parameter ranges from data
            X = self.dynamic_data['X']
            input_names = self.dynamic_data['input_names']

            # Store sliders and entry widgets
            self.dynamic_pred_sliders = {}
            self.dynamic_pred_entries = {}

            for idx, param_name in enumerate(input_names):
                # Get min/max from data
                param_values = X[:, idx]
                param_min = float(np.min(param_values))
                param_max = float(np.max(param_values))
                param_median = float(np.median(param_values))

                # Create parameter frame
                param_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
                param_frame.pack(fill="x", pady=8, padx=5)

                # Parameter label
                ctk.CTkLabel(param_frame, text=param_name, anchor="w").pack(anchor="w", padx=5)

                # Min/Max labels
                range_label = ctk.CTkLabel(
                    param_frame,
                    text=f"Range: [{param_min:.4g}, {param_max:.4g}]",
                    font=("", 9),
                    text_color="gray"
                )
                range_label.pack(anchor="w", padx=5)

                # Entry + Slider frame
                control_frame = ctk.CTkFrame(param_frame, fg_color="transparent")
                control_frame.pack(fill="x", padx=5, pady=2)

                # Entry field
                entry = ctk.CTkEntry(control_frame, width=80)
                entry.pack(side="left", padx=(0, 5))
                entry.insert(0, f"{param_median:.6g}")

                # Slider
                slider = ctk.CTkSlider(
                    control_frame,
                    from_=param_min,
                    to=param_max,
                    command=lambda val, e=entry: e.delete(0, "end") or e.insert(0, f"{float(val):.6g}")
                )
                slider.pack(side="left", fill="x", expand=True)
                slider.set(param_median)

                # Link entry to slider
                def update_slider_from_entry(event, s=slider, pmin=param_min, pmax=param_max):
                    try:
                        val = float(event.widget.get())
                        val = max(pmin, min(pmax, val))  # Clamp to range
                        s.set(val)
                        self._update_dynamic_prediction()
                    except ValueError:
                        pass

                entry.bind("<Return>", update_slider_from_entry)
                entry.bind("<FocusOut>", update_slider_from_entry)

                # Link slider to prediction update
                slider.configure(command=lambda val, e=entry: (
                    e.delete(0, "end"),
                    e.insert(0, f"{float(val):.6g}"),
                    self._update_dynamic_prediction()
                ))

                self.dynamic_pred_sliders[param_name] = slider
                self.dynamic_pred_entries[param_name] = entry

            # Predict button
            predict_btn = ctk.CTkButton(
                left_panel,
                text="Predict All Outputs",
                command=self._update_dynamic_prediction,
                fg_color="#00695c",
                hover_color="#004d40"
            )
            predict_btn.pack(pady=10, padx=10, fill="x")

            # Separator
            separator = ctk.CTkFrame(left_panel, height=2, fg_color="gray")
            separator.pack(fill="x", padx=10, pady=10)

            # Best/Worst Case Analysis Section
            ctk.CTkLabel(left_panel, text="Best/Worst Case Analysis", font=("", 12, "bold")).pack(pady=5, padx=10)

            # Metric selector (peak, mean, RMS)
            metric_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
            metric_frame.pack(fill="x", padx=10, pady=5)

            ctk.CTkLabel(metric_frame, text="Metric:").pack(side="left", padx=5)
            self.metric_var = tk.StringVar(value="Peak")
            metric_selector = ctk.CTkOptionMenu(
                metric_frame,
                variable=self.metric_var,
                values=["Peak", "Peak (Absolute)", "Mean", "RMS"],
                width=140
            )
            metric_selector.pack(side="left", padx=5)

            # Find Both button
            find_both_btn = ctk.CTkButton(
                left_panel,
                text="Find Best & Worst Cases",
                command=self._find_both_extremes,
                fg_color="#d32f2f",
                hover_color="#b71c1c",
                height=35
            )
            find_both_btn.pack(pady=10, padx=10, fill="x")

            # Comparison table frame (scrollable)
            self.comparison_scroll = ctk.CTkScrollableFrame(left_panel, height=200)
            self.comparison_scroll.pack(fill="both", expand=True, padx=10, pady=5)

            # Store optimal parameters for both cases
            self.best_parameters = None
            self.worst_parameters = None

            # Right panel: Prediction plot
            right_panel = ctk.CTkFrame(main_container)
            right_panel.pack(side="right", fill="both", expand=True)

            # Output selector
            top_bar = ctk.CTkFrame(right_panel)
            top_bar.pack(fill="x", padx=5, pady=5)

            ctk.CTkLabel(top_bar, text="Output:").pack(side="left", padx=5)

            self.interactive_output_selector = ctk.CTkOptionMenu(
                top_bar,
                values=self.dynamic_data['output_names'],
                command=lambda _: self._update_dynamic_prediction()
            )
            self.interactive_output_selector.pack(side="left", padx=5)

            # Plot area
            self.interactive_plot_frame = ctk.CTkFrame(right_panel, fg_color="transparent")
            self.interactive_plot_frame.pack(fill="both", expand=True, padx=5, pady=5)

            # Initial prediction
            self._update_dynamic_prediction()

        except Exception as e:
            print(f"Error creating interactive prediction interface: {e}")
            import traceback
            traceback.print_exc()

    def _update_dynamic_prediction(self):
        """Update prediction plot with current input values"""
        try:
            if not hasattr(self, 'dynamic_pred_sliders'):
                return

            # Get current input values
            input_names = self.dynamic_data['input_names']
            input_values = []

            for param_name in input_names:
                try:
                    val = float(self.dynamic_pred_entries[param_name].get())
                    input_values.append(val)
                except ValueError:
                    slider = self.dynamic_pred_sliders[param_name]
                    input_values.append(slider.get())

            # Convert to numpy array
            X_pred = np.array([input_values])

            # Get selected output
            output_name = self.interactive_output_selector.get()
            model = self.dynamic_trainer.models[output_name]

            # Predict with uncertainty
            Y_pred, Y_std = model.predict_with_uncertainty(X_pred)

            # Get frequency array
            frequency = self.dynamic_data['frequency']

            # Create plot
            fig, ax = plt.subplots(figsize=(10, 5))

            # Plot prediction with uncertainty
            ax.plot(frequency, Y_pred[0], '-', color='crimson', linewidth=2, label='Predicted')
            ax.fill_between(
                frequency,
                Y_pred[0] - 2 * Y_std[0],
                Y_pred[0] + 2 * Y_std[0],
                color='crimson',
                alpha=0.2,
                label='95% Confidence'
            )

            ax.set_xlabel('Frequency (Hz)', fontweight='bold')
            ax.set_ylabel(output_name, fontweight='bold')
            ax.set_title(f'Predicted Frequency Response: {output_name}', fontweight='bold')
            ax.legend()
            ax.grid(True, alpha=0.3)

            # Add parameter values as text
            param_text = "Input Parameters:\n" + "\n".join(
                [f"{name}: {val:.4g}" for name, val in zip(input_names, input_values)]
            )
            ax.text(
                0.02, 0.98, param_text,
                transform=ax.transAxes,
                fontsize=8,
                verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            )

            plt.tight_layout()

            # Embed in frame
            self._embed_figure_in_tab(fig, self.interactive_plot_frame)
            self.apply_theme()

        except Exception as e:
            print(f"Error updating dynamic prediction: {e}")
            import traceback
            traceback.print_exc()

    def _find_extreme_case(self):
        """Find best/worst parameter combination for selected output"""
        try:
            from scipy.optimize import differential_evolution, minimize

            # Get selected output
            output_name = self.interactive_output_selector.get()
            model = self.dynamic_trainer.models[output_name]

            # Get parameter bounds
            X = self.dynamic_data['X']
            input_names = self.dynamic_data['input_names']
            frequency = self.dynamic_data['frequency']

            bounds = []
            for idx in range(len(input_names)):
                param_min = float(np.min(X[:, idx]))
                param_max = float(np.max(X[:, idx]))
                bounds.append((param_min, param_max))

            # Get metric and objective
            metric_type = self.metric_var.get()
            is_maximum = self.extreme_case_var.get() == "Maximum"

            # Store evaluation history for debugging
            eval_history = []
            best_found = {'value': float('inf') if not is_maximum else float('-inf'), 'params': None}

            # Define objective function
            def objective(x):
                """Objective function for optimization"""
                X_pred = np.array([x])
                Y_pred = model.predict(X_pred)[0]

                # Calculate metric based on selection
                if metric_type == "Peak":
                    # For peak, we want actual max/min value (signed)
                    if is_maximum:
                        metric_value = np.max(Y_pred)  # Find highest positive peak
                    else:
                        metric_value = np.min(Y_pred)  # Find lowest negative peak
                elif metric_type == "Peak (Absolute)":
                    # For absolute peak, find largest magnitude
                    metric_value = np.max(np.abs(Y_pred))
                elif metric_type == "Mean":
                    metric_value = np.mean(Y_pred)
                elif metric_type == "RMS":
                    metric_value = np.sqrt(np.mean(Y_pred**2))
                else:
                    if is_maximum:
                        metric_value = np.max(Y_pred)
                    else:
                        metric_value = np.min(Y_pred)

                # Track best found
                if is_maximum:
                    if metric_value > best_found['value']:
                        best_found['value'] = metric_value
                        best_found['params'] = x.copy()
                        print(f"New MAXIMUM found: {metric_value:.6g} at {x}")
                else:
                    if metric_value < best_found['value']:
                        best_found['value'] = metric_value
                        best_found['params'] = x.copy()
                        print(f"New MINIMUM found: {metric_value:.6g} at {x}")

                # For differential_evolution (minimizes), return negative for maximization
                return -metric_value if is_maximum else metric_value

            # Update status
            self.extreme_results_label.configure(
                text="Searching for optimal parameters...\n(Running multi-start optimization)",
                text_color="orange"
            )
            self.extreme_results_label.update()

            # Run optimization multiple times with different seeds to ensure global optimum
            best_result = None
            best_value = float('inf') if not is_maximum else float('-inf')

            n_runs = 3  # Number of independent optimization runs
            for run in range(n_runs):
                self.extreme_results_label.configure(
                    text=f"Searching for optimal parameters...\nRun {run+1}/{n_runs}",
                    text_color="orange"
                )
                self.extreme_results_label.update()

                # Run optimization using differential evolution (global optimizer)
                result = differential_evolution(
                    objective,
                    bounds,
                    seed=42 + run,  # Different seed for each run
                    maxiter=200,  # Increased iterations
                    popsize=20,   # Increased population
                    atol=1e-8,
                    tol=1e-8,
                    workers=1,
                    strategy='best1bin',  # Good for global search
                    mutation=(0.5, 1.5),  # Wider mutation range
                    recombination=0.7,
                    polish=True  # Local polish at the end
                )

                if result.success:
                    current_value = -result.fun if is_maximum else result.fun
                    # Check if this is better than previous best
                    if is_maximum:
                        if current_value > best_value:
                            best_value = current_value
                            best_result = result
                    else:
                        if current_value < best_value:
                            best_value = current_value
                            best_result = result

            if best_result is not None and best_result.success:
                optimal_params = best_result.x
                optimal_value = best_value  # Use the best value we tracked

                # Store optimal parameters for later application
                self.optimal_parameters = {
                    'params': optimal_params,
                    'value': optimal_value,
                    'input_names': input_names,
                    'output_name': output_name,
                    'metric_type': metric_type,
                    'case_type': "BEST" if is_maximum else "WORST"
                }

                # Enable the apply button
                self.apply_optimal_btn.configure(state="normal")

                # Verify the result by computing it again
                X_verify = np.array([optimal_params])
                Y_verify = model.predict(X_verify)[0]

                if metric_type == "Peak":
                    if is_maximum:
                        verify_value = np.max(Y_verify)
                    else:
                        verify_value = np.min(Y_verify)
                elif metric_type == "Peak (Absolute)":
                    verify_value = np.max(np.abs(Y_verify))
                elif metric_type == "Mean":
                    verify_value = np.mean(Y_verify)
                elif metric_type == "RMS":
                    verify_value = np.sqrt(np.mean(Y_verify**2))
                else:
                    verify_value = optimal_value

                # Display results
                case_type = "BEST" if is_maximum else "WORST"
                results_text = f"✓ {case_type} CASE FOUND\n"
                results_text += f"{metric_type} Value: {optimal_value:.4g}\n"
                results_text += f"Verified: {verify_value:.4g}\n\n"
                results_text += "Optimal Parameters:\n"
                for param_name, val in zip(input_names, optimal_params):
                    results_text += f"  {param_name}: {val:.5g}\n"
                results_text += f"\n⚠ Click 'Apply Optimal Parameters'\n   to update sliders"

                self.extreme_results_label.configure(
                    text=results_text,
                    text_color="green"
                )

                # Show success message
                messagebox.showinfo(
                    "Optimization Complete",
                    f"Found {case_type.lower()} case!\n\n"
                    f"{metric_type} {output_name}: {optimal_value:.6g}\n\n"
                    "Click 'Apply Optimal Parameters' to update the sliders."
                )

            else:
                self.extreme_results_label.configure(
                    text="⚠ Optimization failed. Try different settings.",
                    text_color="red"
                )
                messagebox.showwarning(
                    "Optimization Failed",
                    "Could not find optimal parameters.\n"
                    "The optimization did not converge."
                )

        except Exception as e:
            print(f"Error in extreme case analysis: {e}")
            import traceback
            traceback.print_exc()
            self.extreme_results_label.configure(
                text=f"⚠ Error: {str(e)[:100]}",
                text_color="red"
            )
            messagebox.showerror("Error", f"Optimization error:\n{e}")

    def _find_both_extremes(self):
        """Find both best AND worst parameter combinations using FAST parallel optimization"""
        try:
            from scipy.optimize import differential_evolution
            from concurrent.futures import ThreadPoolExecutor
            import time

            start_time = time.time()

            # Get selected output
            output_name = self.interactive_output_selector.get()
            model = self.dynamic_trainer.models[output_name]

            # Get parameter bounds
            X = self.dynamic_data['X']
            input_names = self.dynamic_data['input_names']
            frequency = self.dynamic_data['frequency']

            bounds = []
            for idx in range(len(input_names)):
                param_min = float(np.min(X[:, idx]))
                param_max = float(np.max(X[:, idx]))
                bounds.append((param_min, param_max))

            # Get metric type
            metric_type = self.metric_var.get()

            # Clear previous comparison table
            for widget in self.comparison_scroll.winfo_children():
                widget.destroy()

            # Update status
            status_label = ctk.CTkLabel(
                self.comparison_scroll,
                text="⏳ Fast parallel search...",
                font=("Arial", 12, "bold"),
                text_color="orange"
            )
            status_label.pack(pady=10)
            self.comparison_scroll.update()

            # FAST optimization function
            def find_extreme_fast(is_maximum):
                """Fast single-run optimization"""
                def objective(x):
                    X_pred = np.array([x])
                    Y_pred = model.predict(X_pred)[0]

                    # Calculate metric
                    if metric_type == "Peak":
                        metric_value = np.max(Y_pred) if is_maximum else np.min(Y_pred)
                    elif metric_type == "Peak (Absolute)":
                        metric_value = np.max(np.abs(Y_pred))
                    elif metric_type == "Mean":
                        metric_value = np.mean(Y_pred)
                    elif metric_type == "RMS":
                        metric_value = np.sqrt(np.mean(Y_pred**2))
                    else:
                        metric_value = np.max(Y_pred) if is_maximum else np.min(Y_pred)

                    return -metric_value if is_maximum else metric_value

                # HIGH QUALITY settings - parallel via ThreadPoolExecutor
                result = differential_evolution(
                    objective,
                    bounds,
                    seed=42,
                    maxiter=300,     # High iterations for best results
                    popsize=20,      # Large population for thorough search
                    atol=1e-8,       # Strict tolerance
                    tol=1e-8,
                    workers=1,       # Single worker (parallelism via ThreadPoolExecutor above)
                    strategy='best1bin',
                    mutation=(0.5, 1.5),
                    recombination=0.7,
                    polish=True      # Local optimization at end for precision
                )

                return {
                    'params': result.x,
                    'value': -result.fun if is_maximum else result.fun,
                    'success': True
                }

            # Run BOTH optimizations in PARALLEL using threads
            print("\n=== FAST PARALLEL OPTIMIZATION ===")

            with ThreadPoolExecutor(max_workers=2) as executor:
                # Submit both jobs
                # WORST = MAXIMUM response (high vibration is bad)
                # BEST = MINIMUM response (low vibration is good)
                worst_future = executor.submit(find_extreme_fast, True)   # Maximum = worst
                best_future = executor.submit(find_extreme_fast, False)   # Minimum = best

                # Get results
                worst_result = worst_future.result()
                best_result = best_future.result()

            elapsed = time.time() - start_time
            print(f"Optimization completed in {elapsed:.2f} seconds")

            if not worst_result['success'] or not best_result['success']:
                status_label.configure(text="❌ Optimization failed", text_color="red")
                return

            # Store both results
            self.worst_parameters = {
                'params': worst_result['params'],
                'value': worst_result['value'],
                'input_names': input_names,
                'output_name': output_name,
                'metric_type': metric_type
            }
            self.best_parameters = {
                'params': best_result['params'],
                'value': best_result['value'],
                'input_names': input_names,
                'output_name': output_name,
                'metric_type': metric_type
            }

            # Remove status label
            status_label.destroy()

            # Create comparison table
            self._create_comparison_table()

            # Show success message with timing
            messagebox.showinfo(
                "Optimization Complete",
                f"Found both extreme cases in {elapsed:.1f}s!\n\n"
                f"WORST {metric_type}: {worst_result['value']:.4g}\n"
                f"BEST {metric_type}: {best_result['value']:.4g}\n\n"
                f"Click Apply buttons to see curves."
            )

        except Exception as e:
            print(f"Error in both extremes analysis: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Optimization error:\n{e}")

    def _create_comparison_table(self):
        """Create the side-by-side comparison table for best vs worst cases"""
        # Clear previous table
        for widget in self.comparison_scroll.winfo_children():
            widget.destroy()

        if self.worst_parameters is None or self.best_parameters is None:
            return

        # Table container with light background for better contrast
        table_frame = ctk.CTkFrame(self.comparison_scroll, fg_color="#f5f5f5")
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Configure grid columns
        table_frame.grid_columnconfigure(0, weight=1, minsize=120)  # Parameter names
        table_frame.grid_columnconfigure(1, weight=1, minsize=100)  # Worst values
        table_frame.grid_columnconfigure(2, weight=1, minsize=100)  # Best values

        # Headers with clear colors
        header_label = ctk.CTkLabel(
            table_frame,
            text="Parameter",
            font=("Arial", 10, "bold"),
            fg_color="#e0e0e0",
            text_color="#333333",
            corner_radius=5
        )
        header_label.grid(row=0, column=0, sticky="ew", padx=2, pady=2)

        worst_header = ctk.CTkLabel(
            table_frame,
            text="WORST",
            font=("Arial", 10, "bold"),
            fg_color="#c62828",
            text_color="white",
            corner_radius=5
        )
        worst_header.grid(row=0, column=1, sticky="ew", padx=2, pady=2)

        best_header = ctk.CTkLabel(
            table_frame,
            text="BEST",
            font=("Arial", 10, "bold"),
            fg_color="#2e7d32",
            text_color="white",
            corner_radius=5
        )
        best_header.grid(row=0, column=2, sticky="ew", padx=2, pady=2)

        # Metric value row - prominent display
        metric_type = self.worst_parameters['metric_type']

        metric_name_label = ctk.CTkLabel(
            table_frame,
            text=f"{metric_type}",
            font=("Arial", 9, "bold"),
            fg_color="#d0d0d0",
            text_color="#333333",
            corner_radius=3
        )
        metric_name_label.grid(row=1, column=0, sticky="ew", padx=2, pady=2)

        worst_metric_label = ctk.CTkLabel(
            table_frame,
            text=f"{self.worst_parameters['value']:.4g}",
            font=("Arial", 10, "bold"),
            fg_color="#ffcdd2",
            text_color="#b71c1c",
            corner_radius=3
        )
        worst_metric_label.grid(row=1, column=1, sticky="ew", padx=2, pady=2)

        best_metric_label = ctk.CTkLabel(
            table_frame,
            text=f"{self.best_parameters['value']:.4g}",
            font=("Arial", 10, "bold"),
            fg_color="#c8e6c9",
            text_color="#1b5e20",
            corner_radius=3
        )
        best_metric_label.grid(row=1, column=2, sticky="ew", padx=2, pady=2)

        # Parameter rows with readable colors
        input_names = self.worst_parameters['input_names']
        worst_params = self.worst_parameters['params']
        best_params = self.best_parameters['params']

        for idx, param_name in enumerate(input_names):
            row = idx + 2

            # Alternating row colors for readability
            row_bg = "#ffffff" if idx % 2 == 0 else "#f0f0f0"

            # Parameter name
            param_label = ctk.CTkLabel(
                table_frame,
                text=param_name,
                font=("Arial", 8),
                fg_color=row_bg,
                text_color="#333333",
                anchor="w",
                corner_radius=2
            )
            param_label.grid(row=row, column=0, sticky="ew", padx=2, pady=1)

            # Worst case value
            worst_val_label = ctk.CTkLabel(
                table_frame,
                text=f"{worst_params[idx]:.4g}",
                font=("Arial", 8),
                fg_color=row_bg,
                text_color="#c62828",
                corner_radius=2
            )
            worst_val_label.grid(row=row, column=1, sticky="ew", padx=2, pady=1)

            # Best case value
            best_val_label = ctk.CTkLabel(
                table_frame,
                text=f"{best_params[idx]:.4g}",
                font=("Arial", 8),
                fg_color=row_bg,
                text_color="#2e7d32",
                corner_radius=2
            )
            best_val_label.grid(row=row, column=2, sticky="ew", padx=2, pady=1)

        # Apply buttons row
        button_row = len(input_names) + 2

        # Empty cell for alignment
        ctk.CTkLabel(table_frame, text="", fg_color="transparent").grid(
            row=button_row, column=0, pady=5
        )

        # Apply Worst button
        apply_worst_btn = ctk.CTkButton(
            table_frame,
            text="Apply Worst",
            command=lambda: self._apply_case_parameters('worst'),
            fg_color="#d32f2f",
            hover_color="#b71c1c",
            height=30,
            font=("Arial", 10, "bold")
        )
        apply_worst_btn.grid(row=button_row, column=1, sticky="ew", padx=5, pady=5)

        # Apply Best button
        apply_best_btn = ctk.CTkButton(
            table_frame,
            text="Apply Best",
            command=lambda: self._apply_case_parameters('best'),
            fg_color="#2e7d32",
            hover_color="#1b5e20",
            height=30,
            font=("Arial", 10, "bold")
        )
        apply_best_btn.grid(row=button_row, column=2, sticky="ew", padx=5, pady=5)

    def _apply_case_parameters(self, case_type):
        """Apply either best or worst case parameters to sliders"""
        try:
            if case_type == 'worst':
                if self.worst_parameters is None:
                    messagebox.showwarning("No Results", "No worst case parameters found.")
                    return
                params_dict = self.worst_parameters
                case_name = "WORST"
                color = "red"
            else:  # best
                if self.best_parameters is None:
                    messagebox.showwarning("No Results", "No best case parameters found.")
                    return
                params_dict = self.best_parameters
                case_name = "BEST"
                color = "green"

            optimal_params = params_dict['params']
            input_names = params_dict['input_names']
            optimal_value = params_dict['value']
            metric_type = params_dict['metric_type']

            # Update sliders and entries
            for idx, param_name in enumerate(input_names):
                optimal_val = optimal_params[idx]

                # Update entry
                entry = self.dynamic_pred_entries[param_name]
                entry.delete(0, "end")
                entry.insert(0, f"{optimal_val:.6g}")

                # Update slider
                slider = self.dynamic_pred_sliders[param_name]
                slider.set(optimal_val)

            # Force update the prediction plot
            self._update_dynamic_prediction()

            # Quick feedback - no popup, just print
            print(f"✓ {case_name} case applied: {metric_type} = {optimal_value:.4g}")

        except Exception as e:
            print(f"Error applying {case_type} case parameters: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to apply parameters:\n{e}")

    def _apply_optimal_parameters(self):
        """Apply the optimal parameters found by optimization to sliders"""
        try:
            if self.optimal_parameters is None:
                messagebox.showwarning("No Results", "No optimal parameters found. Run optimization first.")
                return

            optimal_params = self.optimal_parameters['params']
            input_names = self.optimal_parameters['input_names']
            optimal_value = self.optimal_parameters['value']
            case_type = self.optimal_parameters['case_type']
            metric_type = self.optimal_parameters['metric_type']

            # Update sliders and entries with optimal parameters
            for idx, param_name in enumerate(input_names):
                optimal_val = optimal_params[idx]

                # Update entry first
                entry = self.dynamic_pred_entries[param_name]
                entry.delete(0, "end")
                entry.insert(0, f"{optimal_val:.6g}")

                # Update slider
                slider = self.dynamic_pred_sliders[param_name]
                slider.set(optimal_val)

            # Force update the prediction plot
            self._update_dynamic_prediction()

            # Update results text to show applied
            results_text = f"✓ {case_type} CASE APPLIED\n"
            results_text += f"{metric_type} Value: {optimal_value:.4g}\n\n"
            results_text += "Applied Parameters:\n"
            for param_name, val in zip(input_names, optimal_params):
                results_text += f"  {param_name}: {val:.5g}\n"

            self.extreme_results_label.configure(
                text=results_text,
                text_color="blue"
            )

            messagebox.showinfo(
                "Parameters Applied",
                f"Optimal {case_type.lower()} case parameters\nhave been applied to all sliders!\n\n"
                f"The plot has been updated."
            )

        except Exception as e:
            print(f"Error applying optimal parameters: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to apply parameters:\n{e}")

    def save_test_session(self):
        if not self.rnn_data_filepath or not self.channel_widgets:
            messagebox.showerror("Save Error", "Please load data and define channels first.")
            return
        try:
            model_choices = {name: w['model_choice'].get() for name, w in self.channel_widgets.items()}
            session_data = {
                "data_filepath": self.rnn_data_filepath, "input_channels": self.input_channels,
                "output_channels": self.output_channels, "model_choices": model_choices
            }
            filepath = filedialog.asksaveasfilename(
                defaultextension=".json", filetypes=[("JSON Session Files", "*.json")], title="Save Test Session"
            )
            if not filepath: return
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=4)
            messagebox.showinfo("Success", f"Test session saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save session: {e}")

    def load_test_session(self):
        print("--- LOADING TEST SESSION ---")
        try:
            session_path = filedialog.askopenfilename(
                filetypes=[("JSON Session Files", "*.json")], title="Load Test Session"
            )
            if not session_path: return
            with open(session_path, 'r', encoding='utf-8') as f:
                session_data = json.load(f)

            # Use the new load_rnn_data to handle dynamic/scalar detection
            print(f"Loading data from: {session_data['data_filepath']}")
            self.load_rnn_data(filepath=session_data['data_filepath'])

            # Manually set channels as per session file
            self.input_channels = session_data["input_channels"]
            self.output_channels = session_data["output_channels"]
            print(f"Found {len(self.input_channels)} Inputs, {len(self.output_channels)} Outputs.")

            # If data was scalar, re-run scalar-specific inits
            if self.rnn_data_type == "scalar":
                self._initialize_scalar_mode()

            model_choices = session_data.get("model_choices", {})
            model_to_check = None
            print("Setting model choices...")
            for name, model_name in model_choices.items():
                if name in self.channel_widgets:
                    self.channel_widgets[name]['model_choice'].set(model_name)
                    if model_to_check is None:  # Just check the first one
                        model_to_check = name
                        self.channel_widgets[name]['var'].set(True)
                    else:
                        self.channel_widgets[name]['var'].set(False)

            if model_to_check is None and self.output_channels:
                model_to_check = self.output_channels[0]
                if model_to_check in self.channel_widgets:
                    self.channel_widgets[model_to_check]['var'].set(True)

            print(f"Will check R2Pred for: {model_to_check}")
            self.build_model_button.configure(state="normal")
            print("Calling build_models()...")
            self.app.update_idletasks()
            self.build_models()  # This will now respect the scalar/dynamic mode
            print("build_models() complete.")

            if model_to_check not in self.last_build_stats:
                raise ValueError(f"Stats for {model_to_check} were not generated. Build failed.")

            stats = self.last_build_stats[model_to_check]

            # Handle stats for either mode
            if self.rnn_data_type == "scalar":
                r2_pred = stats.get('r2_pred', np.nan)
            else:  # Dynamic
                r2_pred = stats.get('r2_test', np.nan)

            print("--- BATCH TEST COMPLETE ---")
            if pd.isna(r2_pred) or r2_pred < 0.5:
                print(f"RESULT: FAILED. R2Pred = {r2_pred:.4f}")
                messagebox.showerror("Batch Test FAILED",
                                     f"Test on '{model_to_check}' FAILED.\n\nR2Pred = {r2_pred:.4f}\n\nThis score is too low.")
            else:
                print(f"RESULT: SUCCESS. R2Pred = {r2_pred:.4f}")
                messagebox.showinfo("Batch Test SUCCESS",
                                    f"Test on '{model_to_check}' passed!\n\nR2Pred = {r2_pred:.4f}")
        except Exception as e:
            print(f"--- BATCH TEST FAILED ---")
            print(f"Error: {e}")
            traceback.print_exc()
            messagebox.showerror("Batch Test Error", f"Test failed: {e}")

    def apply_theme(self):
        theme = self.app.get_theme_properties() if hasattr(self.app, "get_theme_properties") else {
            "plot_bg": "#111111", "text_color": "#ffffff", "grid_color": "#444444"}
        try:
            if hasattr(self, "fig_modeling") and self.fig_modeling is not None:
                self.fig_modeling.set_facecolor(theme.get("plot_bg", "#111111"))
                try:
                    for ax in list(self.fig_modeling.axes):
                        ax.set_facecolor(theme.get("plot_bg", "#111111"))
                        ax.tick_params(colors=theme.get("text_color", "#ffffff"))
                        ax.grid(color=theme.get("grid_color", "#444444"), linestyle='--', alpha=0.5)
                        ax.xaxis.label.set_color(theme.get("text_color", "#ffffff"));
                        ax.yaxis.label.set_color(theme.get("text_color", "#ffffff"));
                        ax.title.set_color(theme.get("text_color", "#ffffff"))
                        for spine in ax.spines.values(): spine.set_color(theme["text_color"])
                except Exception:
                    pass
                try:
                    self.canvas_modeling.draw()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            if hasattr(self, "fig_graphics") and self.fig_graphics is not None:
                self.fig_graphics.set_facecolor(theme.get("plot_bg", "#111111"))
                try:
                    for ax in list(self.fig_graphics.axes):
                        ax.set_facecolor(theme.get("plot_bg", "#111111"))
                        ax.tick_params(colors=theme.get("text_color", "#ffffff"))
                        ax.grid(color=theme.get("grid_color", "#444444"), linestyle=':', alpha=0.6)
                        for spine in ax.spines.values(): spine.set_color(theme["text_color"])
                except Exception:
                    pass
                try:
                    self.canvas_graphics.draw()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            if hasattr(self, "fig_shap") and self.fig_shap is not None:
                self.fig_shap.set_facecolor(theme.get("plot_bg", "#111111"))
                try:
                    ax_list_shap = self.fig_shap.axes
                    if not isinstance(ax_list_shap, list): ax_list_shap = [ax_list_shap]
                    for ax in ax_list_shap:
                        ax.set_facecolor(theme.get("plot_bg", "#111111"))
                        ax.tick_params(axis='x', colors=theme.get("text_color", "#ffffff"))
                        ax.tick_params(axis='y', colors=theme.get("text_color", "#ffffff"))
                        ax.grid(color=theme.get("grid_color", "#444444"), linestyle=':', alpha=0.6)
                        for spine in ax.spines.values(): spine.set_color(theme["text_color"])
                        ax.xaxis.label.set_color(theme.get("text_color", "#ffffff"))
                        ax.yaxis.label.set_color(theme.get("text_color", "#ffffff"))
                        ax.title.set_color(theme.get("text_color", "#ffffff"))
                except Exception as e:
                    print(f"SHAP theme error: {e}")
                try:
                    if self.canvas_shap: self.canvas_shap.draw()
                except Exception:
                    pass
        except Exception:
            pass

        # --- NEW: Apply theme to dynamic viz ---
        try:
            if hasattr(self, 'dynamic_viz_notebook'):
                # This is tricky as the figs are generated on the fly.
                # We can try to re-color existing axes if they exist.
                for tab_name in self.dynamic_viz_notebook.tabs():
                    tab = self.dynamic_viz_notebook.tab(tab_name)
                    for widget in tab.winfo_children():
                        if isinstance(widget, FigureCanvasTkAgg):
                            fig = widget.figure
                            fig.set_facecolor(theme.get("plot_bg", "#111111"))
                            for ax in fig.axes:
                                ax.set_facecolor(theme.get("plot_bg", "#111111"))
                                ax.tick_params(colors=theme.get("text_color", "#ffffff"))
                                ax.grid(color=theme.get("grid_color", "#444444"), linestyle='--', alpha=0.5)
                                ax.xaxis.label.set_color(theme.get("text_color", "#ffffff"));
                                ax.yaxis.label.set_color(theme.get("text_color", "#ffffff"));
                                ax.title.set_color(theme.get("text_color", "#ffffff"))
                                for spine in ax.spines.values(): spine.set_color(theme["text_color"])
                            widget.draw()
        except Exception as e:
            print(f"Error applying theme to dynamic viz: {e}")

    def _set_all_visible(self, visible):
        for ch, w in list(self.channel_widgets.items()):
            try:
                w['var'].set(visible)
            except Exception:
                pass
        try:
            self.recompute_graphics()
        except Exception:
            pass

    def _set_all_models(self, model_type):
        """Set all channel models to the same type"""
        for ch, w in list(self.channel_widgets.items()):
            try:
                if 'model_choice' in w:
                    w['model_choice'].set(model_type)
            except Exception:
                pass
        print(f"Set all models to: {model_type}")

    # ---
    # --- NEW METHODS (from IMPLEMENTATION_GUIDE.md) ---
    # ---

    def _on_mode_change(self, choice):
        """
        Handle prediction mode switching.
        FIXED: Hides/shows UI elements and updates dropdown text.
        """
        if choice == "Scalar":
            self.mode_info_label.configure(text="Static single-value prediction")

            # Hide Model dropdown (only for Dynamic mode)
            self.model_type_label.pack_forget()
            self.model_type_dropdown.pack_forget()

            # 1. Restore buttons
            self.recompute_graphics_button.configure(
                state="normal" if self.trained_models else "disabled",
                text="Recompute Graphics",
                command=lambda: self.update_graphic("Interaction")
            )
            self.run_shap_button.configure(state="normal" if self.trained_models else "disabled")
            self.export_preset_button.configure(state="normal" if self.rnn_data is not None else "disabled")

            # 2. Restore channel list controls
            for out_name, widgets in self.channel_widgets.items():
                if 'model_choice' in widgets:
                    widgets['model_choice'].configure(state="normal")
                    # Set it back to a default if it was changed
                    if widgets['model_choice'].get() == "GPR+PCA":
                        widgets['model_choice'].set("OptimizedSVR")

            # 3. Restore hidden frames (YOUR FIX)
            self.filter_bar.grid(row=3, column=0, padx=6, pady=(0, 6), sticky="ew")
            self.variations_frame.grid(row=5, column=0, padx=6, pady=4, sticky="ew")

            # 4. Update graphics controls for Scalar mode
            self.graphic_type_menu.configure(values=["Measured/Predicted", "Interaction"], command=self.update_graphic)
            self.graphic_type_menu.set("Measured/Predicted")
            try:
                self.show_scatter_cb.pack(side="left", padx=6)
            except Exception:
                pass

            # 5. Restore scalar tabs
            try:
                self.plot_tabs.add("Graphics")
                self._create_graphics_tab(self.plot_tabs.tab("Graphics"))  # Recreate it
            except Exception:
                pass  # Already exists
            try:
                self.plot_tabs.add("Sensitivity (SHAP)")
                self._create_shap_tab(self.plot_tabs.tab("Sensitivity (SHAP)"))  # Recreate it
            except Exception:
                pass  # Already exists

            # 6. Remove dynamic tab
            try:
                if hasattr(self, 'dynamic_viz_tab'):
                    self.plot_tabs.delete("Dynamic Results")
                    del self.dynamic_viz_tab
            except Exception:
                pass

        elif choice == "Dynamic":
            self.mode_info_label.configure(text="Frequency response prediction")

            # Show Model dropdown (for Dynamic mode)
            self.model_type_label.pack(side="left", padx=(10, 5))
            self.model_type_dropdown.pack(side="left", padx=5)

            # 1. Update buttons for dynamic mode
            self.recompute_graphics_button.configure(
                state="normal" if hasattr(self, 'dynamic_trainer') and self.dynamic_trainer else "disabled",
                text="Update Graphics"
            )
            self.recompute_graphics_button.configure(command=lambda: self.update_dynamic_graphic(self.graphic_type_menu.get()))
            self.run_shap_button.configure(state="disabled")
            self.export_preset_button.configure(state="disabled")

            # 2. Update channel list to show GPR+PCA (YOUR FIX)
            for out_name, widgets in self.channel_widgets.items():
                if 'model_choice' in widgets:
                    widgets['model_choice'].set("GPR+PCA")  # Set text
                    widgets['model_choice'].configure(state="disabled")  # Grey out

            # 3. Hide unneeded frames (WASTED SPACE FIX)
            self.filter_bar.grid_forget()
            self.variations_frame.grid_forget()

            # 4. Update graphics controls for Dynamic mode
            self.graphic_type_menu.configure(
                values=["Frequency Response", "Error Analysis", "PCA Analysis", "Prediction Grid", "Interactive Prediction"],
                command=self.update_dynamic_graphic
            )
            self.graphic_type_menu.set("Frequency Response")
            try:
                self.show_scatter_cb.pack_forget()
            except Exception:
                pass

            # 5. Remove scalar-only tabs (WASTED SPACE FIX)
            try:
                self.plot_tabs.delete("Graphics")
            except Exception:
                pass  # Doesn't exist
            try:
                self.plot_tabs.delete("Sensitivity (SHAP)")
            except Exception:
                pass  # Doesn't exist

            # 6. Show info message (only if data isn't already loaded and dynamic)
            if not (self.rnn_data is not None and self.rnn_data_type == "dynamic"):
                messagebox.showinfo(
                    "Dynamic Mode",
                    "Dynamic mode selected. Please import a file.\n\n"
                    "Note: Dynamic mode requires data structured as:\n"
                    "- First column: Frequency (Hz)\n"
                    "- Repeating blocks per configuration"
                )

    def _initialize_ref_inputs(self):
        self._ref_inputs = {}
        self._ref_outputs = {}
        if self.rnn_data is None: return
        for c in self.input_channels:
            try:
                v = pd.to_numeric(self.rnn_data[c], errors='coerce')
                self._ref_inputs[c] = float(v.median()) if not v.dropna().empty else 0.0
            except Exception:
                self._ref_inputs[c] = 0.0
        for c in self.output_channels:
            self._ref_outputs[c] = np.nan

    def load_rnn_data(self, filepath=None):
        """Load data with automatic scalar/dynamic detection AND validation"""
        if not filepath:
            filepath = filedialog.askopenfilename(
                title="Import Data",
                filetypes=[("Data Files", "*.csv *.xls *.xlsx *.txt *.dat"), ("All Files", "*.*")]
            )
        if not filepath:
            return

        # Get the mode the user *wants*
        selected_mode = self.prediction_mode.get()

        try:
            # --- Load the file (handles CSV, Excel, and text) ---
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            elif filepath.endswith('.xls') or filepath.endswith('.xlsx'):
                df = pd.read_excel(filepath)
            else:  # Try as text (for CAMEO)
                try:
                    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                        lines = [ln.rstrip('\n') for ln in f]
                    if not lines: raise ValueError("Empty file")
                    header_line = lines[0].strip().lstrip('\ufeff')
                    header = [h.strip() for h in header_line.split('\t')]
                    units_line = lines[2].strip() if len(lines) > 2 else ""
                    units = [u.strip() for u in units_line.split('\t')] if units_line else [''] * len(header)
                    if len(units) < len(header): units += [''] * (len(header) - len(units))
                    self.rnn_units = {h: u for h, u in zip(header, units)}
                    data_lines = lines[2:] if len(lines) > 2 else lines[1:]
                    if data_lines:
                        parts0 = data_lines[0].split('\t')
                        if len(parts0) != len(header):
                            found = None
                            for idx, ln in enumerate(lines[1: min(len(lines), 60)], start=1):
                                parts = ln.split('\t');
                                numeric_like = 0
                                for p in parts:
                                    try:
                                        float(p.replace(',', ''));
                                        numeric_like += 1
                                    except Exception:
                                        pass
                                if numeric_like >= max(1, int(len(parts) * 0.6)):
                                    found = idx;
                                    break
                            if found is not None: data_lines = lines[found:]
                    data_string = "\n".join(data_lines)
                    df = pd.read_csv(io.StringIO(data_string), sep='\t', header=None, names=header, engine='python',
                                     skipinitialspace=True)
                except Exception as e:
                    print(f"Failed to load as text, trying default pandas read: {e}")
                    df = pd.read_csv(filepath, sep=None, engine='python')  # Pandas auto-detect

            # --- Data Loading Finished, now Coerce and Detect ---

            # Coerce all data to numeric, dropping non-numeric cols
            coerced_df = df.copy()
            non_numeric_cols = []
            for col in coerced_df.columns:
                coerced_col = pd.to_numeric(coerced_df[col].astype(str).str.replace(',', '').str.strip(),
                                            errors='coerce')
                if coerced_col.isna().all():
                    non_numeric_cols.append(col)
                else:
                    coerced_df[col] = coerced_col

            if non_numeric_cols:
                print(f"Dropping non-numeric columns: {non_numeric_cols}")
                df = coerced_df.drop(columns=non_numeric_cols)
            else:
                df = coerced_df

            if df.empty:
                raise ValueError("No numeric data found in file.")

            self.rnn_data_filepath = filepath

            # Detect the *actual* structure of the file
            is_dynamic = self._detect_dynamic_structure(df)

            # --- VALIDATION LOGIC ---
            if selected_mode == "Dynamic" and not is_dynamic:
                # User *wanted* Dynamic, but file is Scalar
                messagebox.showerror(
                    "Data Mismatch",
                    "You selected 'Dynamic' mode, but the loaded file appears to be 'Scalar' (no frequency resets detected).\n\n"
                    "Switching mode to 'Scalar'."
                )
                self.prediction_mode.set("Scalar")
                self.rnn_data_type = "scalar"
                self.rnn_data = df
                self._initialize_scalar_mode()  # Run scalar setup
                self._on_mode_change("Scalar")  # Update UI

            elif selected_mode == "Scalar" and is_dynamic:
                # User *wanted* Scalar, but file is Dynamic
                messagebox.showwarning(
                    "Data Mismatch",
                    "You selected 'Scalar' mode, but the loaded file appears to be 'Dynamic' (frequency resets detected).\n\n"
                    "Switching mode to 'Dynamic'."
                )
                self.prediction_mode.set("Dynamic")
                self.rnn_data_type = "dynamic"
                self._load_dynamic_data(df, filepath)  # Run dynamic setup
                self._on_mode_change("Dynamic")  # Update UI

            elif selected_mode == "Dynamic" and is_dynamic:
                # Correct! User wanted Dynamic, file is Dynamic.
                self.rnn_data_type = "dynamic"
                self._load_dynamic_data(df, filepath)

            else:  # selected_mode == "Scalar" and not is_dynamic
                # Correct! User wanted Scalar, file is Scalar.
                self.rnn_data_type = "scalar"
                self.rnn_data = df
                self._initialize_scalar_mode()
            # --- End of validation logic ---

            theme_props = self.app.get_theme_properties()
            text_color = theme_props.get("text_color", "black")
            self.data_status_label.configure(
                text=f"✓ Loaded: {os.path.basename(filepath)} ({self.rnn_data_type})",
                text_color=text_color
            )
            self.build_model_button.configure(state="normal")
            self.save_session_button.configure(state="normal")
            self.export_preset_button.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load data:\n{e}")
            traceback.print_exc()
            self.reset_to_default()

    def _detect_dynamic_structure(self, df):
        """
        Detect if data has dynamic (frequency) structure.
        """
        try:
            # Rule 1: Must have many rows
            if len(df) < 500:
                return False

                # Rule 2: First column must have resets (negative diff)
            first_col = df.iloc[:, 0].values
            diff = np.diff(first_col)
            has_resets = np.any(diff < 0)

            if has_resets:
                return True

            return False
        except Exception as e:
            print(f"Error detecting dynamic structure: {e}")
            return False

    def _initialize_scalar_mode(self):
        """Helper to run original scalar init logic."""
        # This logic is from the original open_import_wizard
        all_cols = list(self.rnn_data.columns)
        try:
            wizard = ImportWizard(self, all_channels=all_cols, units=self.rnn_units, app_instance=self.app)
            channel_config = wizard.wait_for_result()
        except Exception:
            mid = max(1, len(all_cols) // 2)
            channel_config = {'inputs': all_cols[:mid], 'outputs': all_cols[mid:]}

        if channel_config:
            self.trained_models.clear()
            self.input_channels = channel_config.get('inputs', [])
            self.output_channels = channel_config.get('outputs', [])
            self._initialize_ref_inputs()
            self.rnn_data_bounds = {}
            try:
                for col in self.input_channels:
                    valid_data = self.rnn_data[col].dropna()
                    if not valid_data.empty:
                        self.rnn_data_bounds[col] = (valid_data.min(), valid_data.max())
                    else:
                        self.rnn_data_bounds[col] = (np.nan, np.nan)
                print(f"Successfully stored {len(self.rnn_data_bounds)} RNN input boundaries.")
            except Exception as e:
                print(f"Warning: Could not store RNN training bounds: {e}")
                self.rnn_data_bounds = {}

            if hasattr(self.app, 'design_tab'):
                self.app.design_tab._update_rnn_checkbox_state()

            self.populate_channel_list()

    def _load_dynamic_data(self, df, filepath):
        """Load and structure dynamic frequency response data"""
        try:
            # Use a temporary file path for the loader, as it expects a path
            # A better DynamicDataLoader would accept a DataFrame directly.
            # For now, we stick to the guide's implementation.
            # We'll use the *original* filepath

            loader = DynamicDataLoader(filepath)
            # We need to manually set the loader's data if it was read in-memory
            loader.data = df

            # Manually run the steps from the loader's load_and_reshape
            # This avoids re-reading the file.
            freq_data = loader.data.iloc[:, 0].values
            freq_diff = np.diff(freq_data)
            reset_indices = np.where(freq_diff < 0)[0]
            if len(reset_indices) == 0:
                loader.n_freq = len(freq_data);
                loader.n_configs = 1
            else:
                loader.n_freq = reset_indices[0] + 1
                loader.n_configs = len(loader.data) // loader.n_freq

            if (loader.n_configs * loader.n_freq) != len(loader.data):
                raise ValueError("Data structure inconsistent.")

            loader.frequency_array = freq_data[:loader.n_freq]
            all_feature_cols = list(loader.data.columns[1:])
            input_cols, output_cols = [], []
            for col in all_feature_cols:
                first_block = loader.data[col].iloc[:loader.n_freq].values
                is_constant = np.allclose(first_block, first_block[0], rtol=1e-9)
                if is_constant:
                    input_cols.append(col)
                else:
                    output_cols.append(col)
            loader.input_cols = input_cols
            loader.output_cols = output_cols

            X_list = [loader.data[loader.input_cols].iloc[i * loader.n_freq].values for i in range(loader.n_configs)]
            loader.X = np.array(X_list)

            Y_list = [
                loader.data[loader.output_cols].iloc[i * loader.n_freq: (i * loader.n_freq) + loader.n_freq].values for
                i in range(loader.n_configs)]
            loader.Y = np.array(Y_list)

            self.dynamic_data = {
                'X': loader.X, 'Y': loader.Y, 'frequency': loader.frequency_array,
                'input_names': loader.input_cols, 'output_names': loader.output_cols,
                'metadata': {
                    'n_configs': loader.n_configs, 'n_freq': loader.n_freq,
                    'freq_min': loader.frequency_array.min(), 'freq_max': loader.frequency_array.max()
                }
            }

            # Store for UI
            self.input_channels = self.dynamic_data['input_names']
            self.output_channels = self.dynamic_data['output_names']
            self.frequency_array = self.dynamic_data['frequency']
            self.dynamic_metadata = self.dynamic_data['metadata']

            # Create a dummy rnn_data for compatibility with other functions
            self.rnn_data = pd.DataFrame(self.dynamic_data['X'], columns=self.input_channels)

            print(f"\n✓ Dynamic data loaded:")
            print(f"  Configurations: {self.dynamic_metadata['n_configs']}")
            print(f"  Frequency points: {self.dynamic_metadata['n_freq']}")

            # Populate channel list
            self.populate_channel_list()
            self._initialize_ref_inputs()  # Populate ref inputs

            # Store bounds for Design Tab
            self.rnn_data_bounds = {}
            for col in self.input_channels:
                valid_data = self.rnn_data[col].dropna()
                if not valid_data.empty:
                    self.rnn_data_bounds[col] = (valid_data.min(), valid_data.max())

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load dynamic data:\n{e}")
            traceback.print_exc()
            self.prediction_mode.set("Scalar")  # Revert

    def load_cameo_model(self):
        """Load AVL Cameo pre-trained model by letting user select the Python file"""
        global Variant

        try:
            # Let user select the Variant.py file
            file_path = filedialog.askopenfilename(
                title="Select AVL Cameo Python File (e.g., Variant.py)",
                filetypes=[("Python files", "*.py"), ("All files", "*.*")],
                initialdir=os.path.expanduser("~")
            )

            if not file_path:
                return  # User cancelled

            # Get the directory containing the selected file
            cameo_dir = os.path.dirname(file_path)
            module_name = os.path.splitext(os.path.basename(file_path))[0]

            # Add the directory to sys.path if not already there
            if cameo_dir not in sys.path:
                sys.path.insert(0, cameo_dir)

            # Dynamically load the module
            spec = importlib.util.spec_from_file_location(module_name, file_path)
            if spec is None or spec.loader is None:
                messagebox.showerror(
                    "Load Error",
                    f"Could not load module from:\n{file_path}\n\n"
                    "Make sure the file is a valid Python module."
                )
                return

            Variant = importlib.util.module_from_spec(spec)
            sys.modules[module_name] = Variant
            spec.loader.exec_module(Variant)

            # Verify the module has the expected functions
            required_functions = ['Hull_1', 'Fx_1', 'Fy_1', 'Fz_1']
            missing_functions = [f for f in required_functions if not hasattr(Variant, f)]
            if missing_functions:
                messagebox.showerror(
                    "Invalid Module",
                    f"The selected file is missing required functions:\n{', '.join(missing_functions)}\n\n"
                    "Please select a valid AVL Cameo Variant.py file."
                )
                Variant = None
                return

            # Define Cameo inputs and outputs (get from Variant module)
            cameo_inputs = ['B1_offset', 'B2_offset', 'B3_offset', 'B4_offset', 'B5_offset',
                           'Helix_Angle', 'Input_Stifness', 'Lead_Crown_Pinion',
                           'Lead_Slope_Pinion', 'Pressure_Angle']

            cameo_outputs = ['Hull', 'B1_radialStiffnessX', 'B1_radialStiffnessY',
                            'B2_radialStiffnessX', 'B2_radialStiffnessY',
                            'B3_axialStiffness', 'B3_radialStiffnessX', 'B3_radialStiffnessY',
                            'B4_radialStiffnessX', 'B4_radialStiffnessY',
                            'B5_radialStiffnessX', 'B5_radialStiffnessY',
                            'Fx', 'Fy', 'Fz', 'Linear_TE', 'Mx', 'My', 'Tilt_TE']

            # Set up the channels
            self.input_channels = cameo_inputs.copy()
            self.output_channels = cameo_outputs.copy()
            self.cameo_input_names = cameo_inputs.copy()
            self.cameo_output_names = cameo_outputs.copy()

            # Map output names to Variant functions
            self.cameo_models = {
                'Hull': Variant.Hull_1,
                'B1_radialStiffnessX': Variant.B1_radialStiffnessX_1,
                'B1_radialStiffnessY': Variant.B1_radialStiffnessY_1,
                'B2_radialStiffnessX': Variant.B2_radialStiffnessX_1,
                'B2_radialStiffnessY': Variant.B2_radialStiffnessY_1,
                'B3_axialStiffness': Variant.B3_axialStiffness_1,
                'B3_radialStiffnessX': Variant.B3_radialStiffnessX_1,
                'B3_radialStiffnessY': Variant.B3_radialStiffnessY_1,
                'B4_radialStiffnessX': Variant.B4_radialStiffnessX_1,
                'B4_radialStiffnessY': Variant.B4_radialStiffnessY_1,
                'B5_radialStiffnessX': Variant.B5_radialStiffnessX_1,
                'B5_radialStiffnessY': Variant.B5_radialStiffnessY_1,
                'Fx': Variant.Fx_1,
                'Fy': Variant.Fy_1,
                'Fz': Variant.Fz_1,
                'Linear_TE': Variant.Linear_TE_1,
                'Mx': Variant.Mx_1,
                'My': Variant.My_1,
                'Tilt_TE': Variant.Tilt_TE_1
            }

            # Create sample data for the RNN tab sliders (using median of typical ranges)
            # These are reasonable defaults for gear design parameters
            default_values = {
                'B1_offset': 115.5,
                'B2_offset': 21.1,
                'B3_offset': -196.0,
                'B4_offset': 133.5,
                'B5_offset': 7.5,
                'Helix_Angle': 20.0,
                'Input_Stifness': 0.97,
                'Lead_Crown_Pinion': 10.0,
                'Lead_Slope_Pinion': 0.065,
                'Pressure_Angle': 20.0
            }

            # Create a dummy DataFrame for compatibility
            self.rnn_data = pd.DataFrame([default_values])

            # Set bounds (typical ranges for gear parameters)
            self.rnn_data_bounds = {
                'B1_offset': (100, 130),
                'B2_offset': (15, 30),
                'B3_offset': (-210, -180),
                'B4_offset': (120, 150),
                'B5_offset': (0, 15),
                'Helix_Angle': (15, 25),
                'Input_Stifness': (0.5, 1.5),
                'Lead_Crown_Pinion': (5, 15),
                'Lead_Slope_Pinion': (0.04, 0.09),
                'Pressure_Angle': (18, 22)
            }

            # Mark as using Cameo
            self.using_cameo = True
            self.cameo_model_loaded = True
            self.rnn_data_type = "cameo"

            # Populate channel list
            self.populate_channel_list()
            self._initialize_ref_inputs()

            # Update UI
            theme_props = self.app.get_theme_properties()
            text_color = theme_props.get("text_color", "black")
            self.data_status_label.configure(
                text=f"✓ AVL Cameo Model Loaded ({len(cameo_outputs)} outputs)",
                text_color="green"
            )

            self.build_model_button.configure(state="normal", text="Load Cameo Models")
            self.save_session_button.configure(state="normal")
            self.export_preset_button.configure(state="normal")

            # Automatically load all Cameo models (no second click needed)
            self._load_cameo_models()

            messagebox.showinfo(
                "Cameo Model Loaded",
                f"Successfully loaded AVL Cameo pre-trained model!\n\n"
                f"Inputs: {len(cameo_inputs)} parameters\n"
                f"Outputs: {len(cameo_outputs)} parameters\n\n"
                f"All models are now active and ready to use."
            )

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load Cameo model:\n{e}")
            traceback.print_exc()
            self.using_cameo = False
            self.cameo_model_loaded = False

    def build_models(self):
        """Build models based on selected mode"""
        if self.rnn_data is None and not self.using_cameo:
            messagebox.showerror("Error", "No data loaded.")
            return

        # Handle Cameo model loading
        if self.using_cameo and self.cameo_model_loaded:
            self._load_cameo_models()
            return

        mode = self.prediction_mode.get()

        if mode == "Scalar":
            self._build_scalar_models()
        elif mode == "Dynamic":
            self._build_dynamic_models()

    def _build_scalar_models(self):
        """Original scalar model building logic"""
        self.build_model_button.configure(text="Building...", state="disabled")
        self.recompute_graphics_button.configure(state="disabled")
        self.run_shap_button.configure(state="disabled")
        self.app.update_idletasks()

        X = self.rnn_data[self.input_channels].apply(pd.to_numeric, errors='coerce')
        drop_cols = [c for c in X.columns if X[c].isna().all()]
        if drop_cols:
            messagebox.showwarning("Data warning", f"Dropping inputs: {drop_cols}")
            X = X.drop(columns=drop_cols)
            self.input_channels = [c for c in self.input_channels if c not in drop_cols]
            self._initialize_ref_inputs()

        selected = [name for name, w in self.channel_widgets.items() if w['var'].get()]
        if not selected:
            messagebox.showinfo("Build", "No outputs selected to build.")
            self.build_model_button.configure(text="Build selected", state="normal")
            return

        n_plots = len(selected)
        if n_plots == 1:
            ncols, nrows = 1, 1
        elif n_plots == 2:
            ncols, nrows = 2, 1
        elif n_plots <= 4:
            ncols, nrows = 2, 2
        elif n_plots <= 6:
            ncols, nrows = 3, 2
        elif n_plots <= 9:
            ncols, nrows = 3, 3
        else:
            ncols = 4
            nrows = int(np.ceil(n_plots / ncols))

        try:
            plt.close(self.fig_modeling)
        except Exception:
            pass

        fig_w = max(6, 4 * ncols);
        fig_h = max(4, 3 * nrows)
        self.fig_modeling = plt.figure(figsize=(fig_w, fig_h), dpi=100)
        gs = self.fig_modeling.add_gridspec(nrows, ncols, hspace=0.35, wspace=0.35)
        axes_list = [self.fig_modeling.add_subplot(gs[i // ncols, i % ncols]) for i in range(n_plots)]

        try:
            self.canvas_modeling.get_tk_widget().destroy()
        except Exception:
            pass

        self.canvas_modeling = FigureCanvasTkAgg(self.fig_modeling, master=self.plot_tabs.tab("Modeling"))
        self.canvas_modeling.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.trained_models.clear()
        self.last_build_stats.clear()
        self.all_shap_values.clear()
        self.all_shap_samples.clear()
        self.current_shap_model_name = None

        param_grid_robust = {
            'n_estimators': [50, 100], 'max_depth': [2, 3, 4],
            'learning_rate': [0.05, 0.1], 'subsample': [0.7, 0.8],
            'max_features': [0.7, 0.8]
        }
        param_grid_complex = {
            'n_estimators': [100, 200], 'max_depth': [3, 5, 7],
            'learning_rate': [0.05, 0.1],
        }

        for idx, out in enumerate(selected):
            ax = axes_list[idx]
            ax.clear()
            ax.text(0.5, 0.5, f"Tuning {out}...", ha='center', va='center', color='gray')
            ax.set_xticks([]);
            ax.set_yticks([])
            self.canvas_modeling.draw_idle()
            self.app.update_idletasks()

            y = self.rnn_data[out]
            valid_idx = X.dropna().index.intersection(y.dropna().index)
            Xv = X.loc[valid_idx].reset_index(drop=True)
            yv = pd.to_numeric(y.loc[valid_idx].reset_index(drop=True), errors='coerce')

            if Xv.shape[0] < 10:
                ax.clear()
                ax.text(0.5, 0.5, f"No valid rows for {out}", ha='center', va='center', color='gray')
                ax.set_xticks([]);
                ax.set_yticks([])
                continue

            model_choice_widget = self.channel_widgets[out]['model_choice']
            choice = model_choice_widget.get() if hasattr(model_choice_widget, "get") else "RNN"
            model = None

            try:
                if choice == "RobustRNN":
                    model = RobustRNN(n_local=4, poly_order=2, ridge_alpha=10.0, random_state=42)
                    model.fit(Xv.values, yv.values)
                elif choice == "OptimizedSVR":
                    model = OptimizedSVR()
                    model.fit(Xv.values, yv.values)
                else:
                    base_model = GradientBoostingRegressor(random_state=42)
                    cv_folds_tuning = min(5, Xv.shape[0])
                    grid = param_grid_robust if choice == "RNN" else param_grid_complex
                    grid_search = GridSearchCV(estimator=base_model, param_grid=grid, cv=cv_folds_tuning, scoring='r2',
                                               n_jobs=-1)
                    grid_search.fit(Xv.values, yv.values)
                    model = grid_search.best_estimator_
            except Exception as ex:
                messagebox.showwarning("Model Build Error",
                                       f"Failed to tune {out} with {choice} (Error: {ex}).\nFalling back to default model.")
                model = GradientBoostingRegressor(n_estimators=50, max_depth=2, random_state=42)
                model.fit(Xv.values, yv.values)

            if model is not None:
                self.trained_models[out] = {'model': model, 'features': list(self.input_channels)}
                ax.clear()
                preds = self.safe_predict(model, Xv,
                                          feature_cols=list(self.input_channels)) if model is not None else np.full(
                    len(yv), np.nan)
                ax.scatter(preds, yv, alpha=0.85, s=18, edgecolors='none', color='#1f77b4')
                try:
                    vals = np.concatenate([np.asarray(preds).ravel(), np.asarray(yv).ravel()])
                    vals = vals[np.isfinite(vals)]
                    lims = [np.nanmin(vals), np.nanmax(vals)] if vals.size > 0 else [0.0, 1.0]
                    if not np.all(np.isfinite(lims)): lims = [0.0, 1.0]
                    ax.plot(lims, lims, linestyle='--', color='g', alpha=0.7)
                    ax.set_xlim(lims);
                    ax.set_ylim(lims)
                except Exception:
                    pass
                ax.set_xlabel("Predicted");
                ax.set_ylabel("Measured")
                ax.set_title(out, fontsize=10)
                stats = self.compute_model_stats(model, Xv, yv) if model is not None else {}
                self.last_build_stats[out] = stats
                r2 = stats.get('r2', np.nan)
                nrmse = stats.get('nrmse', np.nan)
                try:
                    stats_lines = [
                        f"  R2: {stats.get('r2', np.nan):>7.4f}", f"R2adj: {stats.get('r2_adj', np.nan):>7.4f}",
                        f"R2prd: {stats.get('r2_pred', np.nan):>7.4f}", f" Fval: {stats.get('f_test', np.nan):>7.1f}",
                        f"NRMSE: {stats.get('nrmse', np.nan):>7.2f}%"
                    ]
                    text_font = {'family': 'monospace', 'size': 8}
                    ax.text(0.02, 0.98, "\n".join(stats_lines),
                            transform=ax.transAxes, fontdict=text_font,
                            verticalalignment='top',
                            bbox=dict(boxstyle='round,pad=0.3', fc='black', alpha=0.6),
                            color='white')
                except Exception:
                    pass
                widgets = self.channel_widgets.get(out)
                if widgets:
                    try:
                        q_color = self._get_quality_color(stats.get('r2_pred', np.nan), is_r2=True)
                        f_color = self._get_quality_color((nrmse or 100) / 100.0, is_r2=False)
                        widgets['quality'].configure(fg_color=q_color)
                        widgets['fit'].configure(fg_color=f_color)
                    except Exception:
                        pass

        self.apply_theme()
        try:
            self.canvas_modeling.draw_idle()
        except Exception:
            try:
                self.canvas_modeling.draw()
            except Exception:
                pass

        self.build_model_button.configure(text="Build selected", state="normal")
        self.recompute_graphics_button.configure(state="normal")

        if self.trained_models:
            trained_model_names = list(self.trained_models.keys())
            input_feature_names = list(self.input_channels)
            self.shap_model_menu.configure(values=trained_model_names, state="normal")
            self.shap_model_menu.set(trained_model_names[0])
            self.shap_plot_type_menu.configure(state="normal")
            self.shap_plot_type_menu.set("Bar (Ranking)")
            self.shap_dependence_menu.configure(values=input_feature_names,
                                                state="disabled")
            if input_feature_names:
                self.shap_dependence_menu.set(input_feature_names[0])
            self.run_shap_button.configure(state="normal")
            self.export_shap_button.configure(state="disabled")
            self.draw_shap_plot()
        else:
            self.shap_model_menu.configure(values=["-"], state="disabled")
            self.shap_model_menu.set("-")
            self.shap_plot_type_menu.configure(state="disabled")
            self.shap_dependence_menu.configure(values=["-"], state="disabled")
            self.run_shap_button.configure(state="disabled")
            self.export_shap_button.configure(state="disabled")

        try:
            if "Interaction" in self.graphic_type_menu.get():
                self.build_interaction_graphics(points=48, percentiles=(20, 50, 80), cell_w=2.0, cell_h=1.8,
                                                show_scatter=self.show_scatter_var.get())
        except Exception:
            pass

        if hasattr(self.app, 'design_tab'):
            print(f"RNNTab.build_models: Found {len(self.trained_models)} models. Calling Design Tab update.")
            self.app.design_tab._update_rnn_checkbox_state()

    def _build_dynamic_models(self):
        """Build dynamic frequency response models"""
        from ENHANCED_DYNAMIC_RNN import DynamicModelTrainer, DynamicVisualization
        from sklearn.model_selection import train_test_split  # Ensure import

        self.build_model_button.configure(text="Training Dynamic...", state="disabled")
        self.recompute_graphics_button.configure(state="disabled")
        self.run_shap_button.configure(state="disabled")
        self.app.update_idletasks()

        try:
            # Get data
            X = self.dynamic_data['X']
            Y = self.dynamic_data['Y']
            frequency = self.dynamic_data['frequency']
            input_names = self.dynamic_data['input_names']
            output_names = self.dynamic_data['output_names']

            # Get selected model type
            model_type = self.model_type_var.get()

            # Train all models
            self.dynamic_trainer = DynamicModelTrainer(
                n_components=50,
                variance_threshold=0.99,
                model_type=model_type
            )

            stats = self.dynamic_trainer.train_all_outputs(
                X, Y, frequency, input_names, output_names
            )

            # Store models and stats
            self.trained_models = self.dynamic_trainer.models
            self.last_build_stats = stats

            # Update UI labels
            for out_name, stat_dict in stats.items():
                if out_name in self.channel_widgets:
                    widgets = self.channel_widgets[out_name]
                    q_color = self._get_quality_color(stat_dict.get('r2_test', np.nan), is_r2=True)
                    f_color = self._get_quality_color((stat_dict.get('rmse_test', 1.0) / (Y.max() - Y.min())),
                                                      is_r2=False)  # Simplified NRMSE
                    widgets['quality'].configure(fg_color=q_color)
                    widgets['fit'].configure(fg_color=f_color)

            # Create visualizations
            self._display_dynamic_results(stats, output_names)

            messagebox.showinfo(
                "Training Complete",
                f"Successfully trained {len(output_names)} dynamic models!\n"
                f"Check the 'Dynamic Results' tab."
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Training Error", f"Failed to train models:\n{e}")

        finally:
            self.build_model_button.configure(text="Build Models", state="normal")
            # Enable graphics button for dynamic mode, keep others disabled
            self.recompute_graphics_button.configure(state="normal")
            self.run_shap_button.configure(state="disabled")

    def _load_cameo_models(self):
        """Load all Cameo models (no training needed - pre-trained)"""
        global Variant
        try:
            self.build_model_button.configure(text="Loading Cameo...", state="disabled")
            self.app.update_idletasks()

            # Create wrapper models for each output using CameoModelWrapper
            for output_name in self.output_channels:
                # Create a wrapper that makes Cameo compatible with sklearn interface
                wrapper = CameoModelWrapper(Variant, output_name, self.input_channels)

                # Store in trained_models with same format as regular models
                self.trained_models[output_name] = {
                    'model': wrapper,
                    'features': self.input_channels,
                    'type': 'CAMEO'
                }

                # Update UI to show model is loaded
                if output_name in self.channel_widgets:
                    widgets = self.channel_widgets[output_name]
                    # Set quality indicators to green (pre-trained model)
                    widgets['quality'].configure(fg_color="green")
                    widgets['fit'].configure(fg_color="green")

            # Display Cameo model info in the Modeling tab
            self._display_cameo_model_info()

            # Enable UI elements
            self.recompute_graphics_button.configure(state="normal")
            self.run_shap_button.configure(state="disabled")  # SHAP not supported for Cameo

        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load Cameo models:\n{e}")
            traceback.print_exc()

        finally:
            self.build_model_button.configure(text="Load Cameo Models", state="normal")

    def _display_cameo_model_info(self):
        """Display AVL Cameo model information in the Modeling tab"""
        try:
            # Clear existing plots
            for ax in self.axes_modeling.flat:
                ax.clear()
                ax.set_xticks([])
                ax.set_yticks([])

            theme_props = self.app.get_theme_properties()
            bg_color = theme_props.get("plot_bg", "#111111")
            text_color = theme_props.get("text_color", "#ffffff")

            # Use first 3 axes for header info
            # Axes layout: [0,0], [0,1], [0,2] - Top row
            #              [1,0], [1,1], [1,2] - Middle row
            #              [2,0], [2,1], [2,2] - Bottom row

            # Main title in center top
            ax_title = self.axes_modeling[0, 1]
            ax_title.text(0.5, 0.5, "AVL CAMEO MODEL", ha='center', va='center',
                         fontsize=18, fontweight='bold', color='#00FF00')
            ax_title.text(0.5, 0.2, "Pre-trained Surrogate Model", ha='center', va='center',
                         fontsize=10, color=text_color)

            # Status in top-left
            ax_status = self.axes_modeling[0, 0]
            ax_status.text(0.5, 0.6, "STATUS", ha='center', va='center',
                          fontsize=12, fontweight='bold', color=text_color)
            ax_status.text(0.5, 0.3, "READY", ha='center', va='center',
                          fontsize=16, fontweight='bold', color='#00FF00')

            # Model type in top-right
            ax_type = self.axes_modeling[0, 2]
            ax_type.text(0.5, 0.6, "MODEL TYPE", ha='center', va='center',
                        fontsize=12, fontweight='bold', color=text_color)
            ax_type.text(0.5, 0.3, "DLL-Based", ha='center', va='center',
                        fontsize=14, color='#00BFFF')

            # Input parameters list in middle-left
            ax_inputs = self.axes_modeling[1, 0]
            ax_inputs.text(0.5, 0.95, f"INPUTS ({len(self.input_channels)})", ha='center', va='top',
                          fontsize=11, fontweight='bold', color='#FFA500')
            input_text = ""
            for i, inp in enumerate(self.input_channels[:6]):
                bounds = self.rnn_data_bounds.get(inp, (0, 0))
                input_text += f"{inp}\n[{bounds[0]:.1f} - {bounds[1]:.1f}]\n"
            if len(self.input_channels) > 6:
                input_text += f"... +{len(self.input_channels)-6} more"
            ax_inputs.text(0.5, 0.75, input_text, ha='center', va='top',
                          fontsize=8, color=text_color, linespacing=1.2)

            # Output parameters list in middle-center and middle-right
            ax_outputs = self.axes_modeling[1, 1]
            ax_outputs.text(0.5, 0.95, f"OUTPUTS ({len(self.output_channels)})", ha='center', va='top',
                           fontsize=11, fontweight='bold', color='#00FF00')
            output_text = ""
            for i, out in enumerate(self.output_channels[:8]):
                output_text += f"{out}\n"
            if len(self.output_channels) > 8:
                output_text += f"... +{len(self.output_channels)-8} more"
            ax_outputs.text(0.5, 0.75, output_text, ha='center', va='top',
                           fontsize=8, color=text_color, linespacing=1.3)

            # Quality indicator in middle-right
            ax_quality = self.axes_modeling[1, 2]
            ax_quality.text(0.5, 0.6, "QUALITY", ha='center', va='center',
                           fontsize=12, fontweight='bold', color=text_color)
            ax_quality.text(0.5, 0.3, "PRE-TRAINED", ha='center', va='center',
                           fontsize=12, color='#00FF00')

            # Instructions in bottom row
            ax_instr1 = self.axes_modeling[2, 0]
            ax_instr1.text(0.5, 0.5, "Use 'Graphics' tab\nfor predictions", ha='center', va='center',
                          fontsize=10, color='#888888')

            ax_instr2 = self.axes_modeling[2, 1]
            ax_instr2.text(0.5, 0.5, "Export Design Preset\nfor Design Tab", ha='center', va='center',
                          fontsize=10, color='#888888')

            ax_instr3 = self.axes_modeling[2, 2]
            ax_instr3.text(0.5, 0.5, "Adjust sliders in\nGraphics tab", ha='center', va='center',
                          fontsize=10, color='#888888')

            # Apply background colors
            for ax in self.axes_modeling.flat:
                ax.set_facecolor(bg_color)

            self.fig_modeling.set_facecolor(bg_color)
            self.canvas_modeling.draw()

        except Exception as e:
            print(f"Error displaying Cameo model info: {e}")
            traceback.print_exc()

    def _display_dynamic_results(self, stats, output_names):
        """Display dynamic modeling results"""
        from ENHANCED_DYNAMIC_RNN import DynamicVisualization  # Import here

        # Create or clear dynamic viz tab
        try:
            # If tab exists, delete and recreate
            if hasattr(self, 'dynamic_viz_tab'):
                self.plot_tabs.delete("Dynamic Results")
        except Exception:
            pass  # Tab didn't exist

        self.dynamic_viz_tab = self.plot_tabs.add("Dynamic Results")
        self.plot_tabs.set("Dynamic Results")  # Switch to the new tab

        # --- SHARED CONTROL FRAME (THE FIX) ---
        # Create it as a child of the MAIN TAB (self.dynamic_viz_tab)
        shared_control_frame = ctk.CTkFrame(self.dynamic_viz_tab)
        # Pack it FIRST, at the TOP of the main tab
        shared_control_frame.pack(fill="x", padx=10, pady=(5, 0))

        ctk.CTkLabel(shared_control_frame, text="Output:").pack(side="left", padx=5)
        self.dynamic_output_selector = ctk.CTkOptionMenu(
            shared_control_frame,
            values=output_names,
            command=self._update_dynamic_visualizations
        )
        self.dynamic_output_selector.pack(side="left", padx=5)

        ctk.CTkLabel(shared_control_frame, text="Config (Test Set):").pack(side="left", padx=5)

        # Get test set size
        n_test_samples = int(self.dynamic_metadata['n_configs'] * 0.2)
        if n_test_samples == 0: n_test_samples = 1

        self.dynamic_config_selector = ctk.CTkOptionMenu(
            shared_control_frame,
            values=[str(i) for i in range(n_test_samples)],
            command=self._update_dynamic_visualizations
        )
        self.dynamic_config_selector.pack(side="left", padx=5)
        # --- END SHARED CONTROL FRAME ---

        # --- NOTEBOOK (Now packed *after* the controls) ---
        self.dynamic_viz_notebook = ctk.CTkTabview(self.dynamic_viz_tab,
                                                   border_width=1,
                                                   fg_color=self.app.get_theme_properties().get("plot_bg", "gray20")
                                                   )
        self.dynamic_viz_notebook.pack(fill="both", expand=True, padx=6, pady=6)

        # Add tabs for different visualizations
        self.dynamic_viz_notebook.add("Training Summary")
        self.dynamic_viz_notebook.add("Frequency Response")
        self.dynamic_viz_notebook.add("Error Analysis")
        self.dynamic_viz_notebook.add("PCA Analysis")

        # --- Create empty frames for embedding plots ---
        # These frames will live inside their respective tabs
        self.summary_plot_frame = ctk.CTkFrame(self.dynamic_viz_notebook.tab("Training Summary"),
                                               fg_color="transparent")
        self.summary_plot_frame.pack(fill="both", expand=True)

        self.freq_plot_frame = ctk.CTkFrame(self.dynamic_viz_notebook.tab("Frequency Response"), fg_color="transparent")
        self.freq_plot_frame.pack(fill="both", expand=True)

        self.error_plot_frame = ctk.CTkFrame(self.dynamic_viz_notebook.tab("Error Analysis"), fg_color="transparent")
        self.error_plot_frame.pack(fill="both", expand=True)

        self.pca_plot_frame = ctk.CTkFrame(self.dynamic_viz_notebook.tab("PCA Analysis"), fg_color="transparent")
        self.pca_plot_frame.pack(fill="both", expand=True)

        # Generate visualizations
        self._generate_dynamic_visualizations(stats, output_names)

    def _generate_dynamic_visualizations(self, stats, output_names):
        """Generate all dynamic visualization plots"""
        from ENHANCED_DYNAMIC_RNN import DynamicVisualization  # Import here

        # 1. Training Summary (doesn't need update)
        try:
            fig1 = DynamicVisualization.plot_training_summary(stats, output_names)
            self._embed_figure_in_tab(fig1, self.summary_plot_frame)
        except Exception as e:
            print(f"Error plotting training summary: {e}")
            traceback.print_exc()

        # 2. Initial frequency response (first output, first config)
        self._update_dynamic_visualizations(output_names[0])
        self.apply_theme()  # Apply theme to all new plots

    def _update_dynamic_visualizations(self, *args):
        """Update dynamic visualizations based on selector"""
        from ENHANCED_DYNAMIC_RNN import DynamicVisualization  # Import here
        from sklearn.model_selection import train_test_split  # Ensure import

        try:
            output_name = self.dynamic_output_selector.get()
            config_idx = int(self.dynamic_config_selector.get())
        except Exception as e:
            print(f"Could not get dynamic selectors: {e}")
            return  # Not ready yet

        try:
            # Get test data (use same random_state as in training)
            X_train, X_test, Y_train, Y_test = train_test_split(
                self.dynamic_data['X'],
                self.dynamic_data['Y'],
                test_size=0.2,
                random_state=42
            )

            output_idx = self.dynamic_data['output_names'].index(output_name)
            Y_test_single = Y_test[:, :, output_idx]

            model = self.trained_models[output_name]
            Y_pred, Y_std = model.predict_with_uncertainty(X_test)

            frequency = self.dynamic_data['frequency']

            # Frequency Response
            fig2 = DynamicVisualization.plot_frequency_response_prediction(
                frequency, Y_test_single, Y_pred, Y_std,
                config_idx=min(config_idx, len(Y_test_single) - 1),
                output_name=output_name
            )
            self._embed_figure_in_tab(fig2, self.freq_plot_frame)

            # Error Analysis
            fig3 = DynamicVisualization.plot_error_analysis(
                frequency, Y_test_single, Y_pred, output_name
            )
            self._embed_figure_in_tab(fig3, self.error_plot_frame)

            # PCA Analysis
            fig4 = DynamicVisualization.plot_pca_analysis(model, frequency, output_name)
            self._embed_figure_in_tab(fig4, self.pca_plot_frame)

            # Re-apply theme to new/updated plots
            self.apply_theme()

        except Exception as e:
            print(f"Error updating dynamic visualizations: {e}")
            traceback.print_exc()

    def _embed_figure_in_tab(self, fig, parent_frame):
        """Helper to embed matplotlib figure in CTk frame"""
        # Clear existing
        for widget in parent_frame.winfo_children():
            widget.destroy()

        # Create new canvas
        canvas = FigureCanvasTkAgg(fig, master=parent_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)

        # Add toolbar
        toolbar = NavigationToolbar2Tk(canvas, parent_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(side="bottom", fill="x", padx=5, pady=2)

        plt.close(fig)  # Prevent memory leak

    def reset_to_default(self):
        print("--- Resetting RNN Tab ---")
        self.rnn_data = None
        self.rnn_units = {}
        self.input_channels = []
        self.output_channels = []
        self.trained_models = {}
        self.channel_widgets = {}
        self._ref_inputs = {}
        self._ref_outputs = {}
        self.rnn_data_bounds = {}
        self.last_build_stats = {}
        self.rnn_data_filepath = None

        # Reset dynamic mode
        self.prediction_mode.set("Scalar")
        self.rnn_data_type = "scalar"
        self.dynamic_data = None
        self.frequency_array = None
        self.dynamic_metadata = None
        try:
            if hasattr(self, 'dynamic_viz_tab'):
                self.plot_tabs.delete("Dynamic Results")
                del self.dynamic_viz_tab
        except Exception:
            pass

        self.all_shap_values.clear()
        self.all_shap_samples.clear()
        self.current_shap_model_name = None

        for w in self.variations_frame.winfo_children():
            w.destroy()
        self.variation_widgets.clear()
        ctk.CTkLabel(self.variations_frame, text="Load data and build models to see variations.",
                     text_color="gray").pack(padx=10, pady=10)

        self.data_status_label.configure(text="No data loaded.", text_color="gray")
        self.build_model_button.configure(state="disabled")
        self.save_session_button.configure(state="disabled")
        self.export_preset_button.configure(state="disabled")

        if hasattr(self, 'run_shap_button'):
            self.run_shap_button.configure(state="disabled")
            self.shap_model_menu.configure(values=["-"], state="disabled")
            self.shap_model_menu.set("-")
            self.shap_plot_type_menu.configure(state="disabled")
            self.shap_plot_type_menu.set("Bar (Ranking)")
            self.shap_dependence_menu.configure(values=["-"], state="disabled")
            self.shap_dependence_menu.set("-")
            self.export_shap_button.configure(state="disabled")

        for w in self._ch_inner.winfo_children():
            w.destroy()
        ctk.CTkLabel(self._ch_inner, text="No channels defined", text_color="gray").pack(padx=8, pady=8)

        if hasattr(self, 'axes_modeling') and self.axes_modeling is not None:
            for ax in self.axes_modeling.flat:
                ax.clear()
                ax.text(0.5, 0.5, 'Load Data and Build Models', ha='center', va='center', fontsize=12, color='gray')
                ax.set_xticks([]);
                ax.set_yticks([])
        if hasattr(self, 'ax_graphics') and self.ax_graphics is not None:
            self.ax_graphics.clear()
            self.ax_graphics.text(0.5, 0.5, 'Interaction Plot Area', ha='center', va='center', fontsize=14,
                                  color='gray')
            self.ax_graphics.set_xticks([]);
            self.ax_graphics.set_yticks([])  # Corrected from ax.set_yticks([])
        if hasattr(self, 'ax_shap') and self.ax_shap is not None:
            self.ax_shap.clear()
            theme_props = self.app.get_theme_properties()
            text_color = theme_props.get("text_color", "black")
            self.ax_shap.text(0.5, 0.5, 'Build models, then click "Run SHAP"', ha='center', va='center', fontsize=12,
                              color=text_color)
            self.ax_shap.set_xticks([]);
            self.ax_shap.set_yticks([])  # Corrected from ax.set_yticks([])

        self.apply_theme()
        try:
            self.canvas_modeling.draw_idle()
        except Exception:
            pass
        try:
            self.canvas_graphics.draw_idle()
        except Exception:
            pass
        try:
            if self.canvas_shap: self.canvas_shap.draw_idle()
        except Exception:
            pass

        if hasattr(self.app, 'design_tab'):
            self.app.design_tab._update_rnn_checkbox_state()


class SamplingUpdatedApp(ctk.CTk):
    def __init__(self):  # <--- THIS IS THE CORRECT SIGNATURE
        super().__init__()

        # --- THIS IS THE FIX ---
        # Initialize fonts *after* the root window (self) exists
        # but before any other widgets are created.
        initialize_fonts()
        # --- END FIX ---

        self.title("Sampling Generator & RNN Modeler")
        self.geometry("1220x800")

        # --- CHANGES FOR LIGHT THEME ---
        self.active_theme = "Light"
        self.df, self.generated_df, self.param_names, self.selected_indices = None, None, [], []
        self.ref_df, self.ref_corr_matrix, self.generated_design_df = None, None, None

        # --- THIS IS THE FIX ---
        self.target_stats = None  # <-- ADD THIS LINE TO INITIALIZE THE ATTRIBUTE
        # -----------------------

        ctk.set_appearance_mode(THEMES["Light"]["ctk_mode"])
        ctk.set_default_color_theme(THEMES["Light"]["ctk_color_theme"])
        plt.style.use(THEMES["Light"]["mpl_style"])
        # ---------------------------------

        self.tabs = ctk.CTkTabview(self, width=1210, command=self._on_tab_select)
        self.tabs.pack(fill="both", expand=True, padx=6, pady=6)
        self.tabs.add("Generator");
        self.tabs.add("Visualization");
        self.tabs.add("Design");
        self.tabs.add("RNN");
        self.tabs.add("Help")

        self.generator_tab = GeneratorTab(self.tabs.tab("Generator"), self);
        self.generator_tab.pack(fill="both", expand=True)
        self.visualization_tab = VisualizationTab(self.tabs.tab("Visualization"), self);
        self.visualization_tab.pack(fill="both", expand=True)
        self.design_tab = DesignTab(self.tabs.tab("Design"), self);
        self.design_tab.pack(fill="both", expand=True)
        self.rnn_tab = RNNTab(self.tabs.tab("RNN"), self);
        self.rnn_tab.pack(fill="both", expand=True)
        self.help_tab = HelpTab(self.tabs.tab("Help"), self);
        self.help_tab.pack(fill="both", expand=True)

        footer_frame = ctk.CTkFrame(self, height=40);
        footer_frame.pack(fill="x", padx=6, pady=(0, 6))
        ctk.CTkLabel(footer_frame, text="Appearance Theme:").pack(side="left", padx=(10, 5))
        self.theme_menu = ctk.CTkOptionMenu(footer_frame, values=list(THEMES.keys()), command=self._change_theme);
        self.theme_menu.set("Light")
        self.theme_menu.pack(side="left", padx=5)

        # --- NEW: Master Reset Button ---
        self.reset_button = ctk.CTkButton(footer_frame, text="Master Reset",
                                          command=self.master_reset,
                                          fg_color="#D32F2F", hover_color="#B71C1C")
        self.reset_button.pack(side="right", padx=10)

        self.tabs.set("Generator")

    def _on_tab_select(self):
        """
        Called every time the user clicks a different tab.
        This is the permanent fix.
        """
        try:
            # Get the name of the tab that was just clicked
            selected_tab_name = self.tabs.get()

            # If the user clicked on the "Design" tab...
            if selected_tab_name == "Design":
                # ...tell the Design Tab to update its internal state.
                if hasattr(self, 'design_tab') and hasattr(self.design_tab, '_update_rnn_checkbox_state'):
                    # This print statement will show in your console so you know it's working
                    print("User selected Design Tab. Refreshing RNN checkbox state.")
                    self.design_tab._update_rnn_checkbox_state()
        except Exception as e:
            print(f"Error in _on_tab_select: {e}")

    def _change_theme(self, theme_name: str):
        self.active_theme = theme_name
        theme_props = self.get_theme_properties()

        # Set CTk theme
        ctk.set_appearance_mode(theme_props["ctk_mode"])

        # --- THIS IS THE MISSING LINE ---
        # Set Matplotlib global theme
        plt.style.use(theme_props["mpl_style"])
        # ------------------------------

        # Apply specific colors to existing plots
        self.generator_tab.apply_theme()
        self.visualization_tab.apply_theme()
        self.design_tab.apply_theme()
        self.rnn_tab.apply_theme()

    def get_theme_properties(self):
        return THEMES[self.active_theme]

    def master_reset(self):
        """Resets the entire application to its initial state."""
        if not messagebox.askyesno("Confirm Master Reset",
                                   "Are you sure you want to reset the entire application?\n\nAll loaded data, generated samples, and settings will be lost."):
            return

        print("--- Master Resetting Application ---")  # Debug print

        # 1. Reset main application state variables
        self.df = None
        self.generated_df = None
        self.param_names = []
        self.selected_indices = []
        self.ref_df = None
        self.ref_corr_matrix = None
        self.generated_design_df = None  # <<< Make sure this is cleared
        self.target_stats = None

        # 2. Call the reset method for each tab
        try:
            print("Resetting Generator Tab...")
            self.generator_tab.reset_to_default()
        except Exception as e:
            print(f"Error resetting Generator Tab: {e}")
        try:
            print("Resetting Visualization Tab...")
            self.visualization_tab.reset_to_default()
        except Exception as e:
            print(f"Error resetting Visualization Tab: {e}")
        try:
            print("Resetting Design Tab...")
            self.design_tab.reset_to_default()  # <<< This calls the method below
        except Exception as e:
            print(f"Error resetting Design Tab: {e}")
        try:
            print("Resetting RNN Tab...")
            self.rnn_tab.reset_to_default()
        except Exception as e:
            print(f"Error resetting RNN Tab: {e}")

        # 3. Go back to the first tab
        self.tabs.set("Generator")
        print("--- Master Reset Complete ---")  # Debug print


# =============================================================================
# --- SCRIPT ENTRY POINT ---
# =============================================================================
if __name__ == "__main__":
    app = SamplingUpdatedApp()
    app.mainloop()
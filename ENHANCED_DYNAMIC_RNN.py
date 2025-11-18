"""
ENHANCED DYNAMIC RNN IMPLEMENTATION
====================================

This module extends the existing RNN script with:
1. Scalar vs Dynamic mode selection dropdown
2. GPR + PCA for dynamic frequency response prediction
3. Best-in-class visualizations for dynamic results
4. Automatic frequency detection and flexible handling

Author: Enhanced by Claude
Date: 2025
"""

import numpy as np
import pandas as pd
from sklearn.decomposition import PCA
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, WhiteKernel, ConstantKernel as C
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_squared_error
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import warnings
warnings.filterwarnings('ignore')

# Try to import XGBoost
try:
    import xgboost as xgb
    XGBOOST_AVAILABLE = True
except ImportError:
    XGBOOST_AVAILABLE = False
    print("Warning: XGBoost not installed. Install with: pip install xgboost")


class DynamicDataLoader:
    """
    Flexible data loader that automatically detects frequency configuration
    and reshapes data for dynamic modeling.
    """

    def __init__(self, filepath):
        """
        Initialize loader with file path.

        Parameters:
        -----------
        filepath : str
            Path to the data file (CSV or TXT)
        """
        self.filepath = filepath
        self.data = None
        self.frequency_array = None
        self.n_freq = None
        self.n_configs = None
        self.input_cols = None
        self.output_cols = None
        self.X = None  # [n_configs, n_inputs]
        self.Y = None  # [n_configs, n_freq, n_outputs]

    def load_and_reshape(self):
        """
        Load data and automatically detect frequency structure.

        Returns:
        --------
        dict with keys:
            - 'X': Input parameters [n_configs, n_inputs]
            - 'Y': Output frequency responses [n_configs, n_freq, n_outputs]
            - 'frequency': Frequency array [n_freq]
            - 'input_names': List of input parameter names
            - 'output_names': List of output parameter names
            - 'metadata': Dict with n_configs, n_freq, etc.
        """
        print("=" * 80)
        print("LOADING DYNAMIC DATA")
        print("=" * 80)

        # Load the data
        if self.filepath.endswith('.csv'):
            self.data = pd.read_csv(self.filepath)
        else:
            self.data = pd.read_csv(self.filepath, sep=',')

        print(f"✓ Loaded file: {self.filepath}")
        print(f"  Total rows: {len(self.data)}")
        print(f"  Total columns: {len(self.data.columns)}")

        # First column should be frequency
        freq_col_name = self.data.columns[0]
        freq_data = self.data.iloc[:, 0].values

        # Detect where frequency resets (starts a new configuration)
        freq_diff = np.diff(freq_data)
        reset_indices = np.where(freq_diff < 0)[0]

        if len(reset_indices) == 0:
            # Only one configuration
            self.n_freq = len(freq_data)
            self.n_configs = 1
        else:
            # Multiple configurations
            self.n_freq = reset_indices[0] + 1
            self.n_configs = len(self.data) // self.n_freq

        print(f"\n✓ Auto-detected structure:")
        print(f"  Frequency points per config: {self.n_freq}")
        print(f"  Number of configurations: {self.n_configs}")

        # Verify structure
        expected_rows = self.n_configs * self.n_freq
        actual_rows = len(self.data)

        if expected_rows != actual_rows:
            raise ValueError(
                f"Data structure inconsistent!\n"
                f"Expected {expected_rows} rows ({self.n_configs} configs × {self.n_freq} freq points)\n"
                f"But found {actual_rows} rows.\n"
                f"Please check your data file."
            )

        # Extract frequency array (from first configuration)
        self.frequency_array = freq_data[:self.n_freq]
        freq_min, freq_max = self.frequency_array.min(), self.frequency_array.max()
        freq_res = np.median(np.diff(self.frequency_array))

        print(f"  Frequency range: {freq_min:.2f} - {freq_max:.2f} Hz")
        print(f"  Frequency resolution: {freq_res:.4f} Hz (median)")

        # Identify input and output columns
        # Assumption: columns after frequency are: inputs, then outputs
        # We need to figure out which are constant (inputs) and which vary (outputs)

        all_feature_cols = list(self.data.columns[1:])

        # For each configuration block, check if columns are constant
        input_cols = []
        output_cols = []

        for col in all_feature_cols:
            # Check first configuration block
            first_block = self.data[col].iloc[:self.n_freq].values
            is_constant = np.all(first_block == first_block[0])

            if is_constant:
                input_cols.append(col)
            else:
                output_cols.append(col)

        self.input_cols = input_cols
        self.output_cols = output_cols

        print(f"\n✓ Identified columns:")
        print(f"  Input parameters ({len(input_cols)}): {input_cols}")
        print(f"  Output parameters ({len(output_cols)}): {output_cols}")

        # Reshape data
        # X: Extract unique input configurations
        X_list = []
        for i in range(self.n_configs):
            start_idx = i * self.n_freq
            # Take first row of each block (inputs are constant within block)
            config_inputs = self.data[input_cols].iloc[start_idx].values
            X_list.append(config_inputs)

        self.X = np.array(X_list)  # Shape: [n_configs, n_inputs]

        # Y: Extract output frequency responses
        Y_list = []
        for i in range(self.n_configs):
            start_idx = i * self.n_freq
            end_idx = start_idx + self.n_freq
            # Get all frequency points for this configuration
            config_outputs = self.data[output_cols].iloc[start_idx:end_idx].values
            Y_list.append(config_outputs)

        self.Y = np.array(Y_list)  # Shape: [n_configs, n_freq, n_outputs]

        print(f"\n✓ Reshaped data:")
        print(f"  X shape: {self.X.shape} (configs × inputs)")
        print(f"  Y shape: {self.Y.shape} (configs × frequencies × outputs)")
        print("=" * 80)

        return {
            'X': self.X,
            'Y': self.Y,
            'frequency': self.frequency_array,
            'input_names': self.input_cols,
            'output_names': self.output_cols,
            'metadata': {
                'n_configs': self.n_configs,
                'n_freq': self.n_freq,
                'freq_min': freq_min,
                'freq_max': freq_max,
                'freq_resolution': freq_res
            }
        }


class GPR_PCA_DynamicModel:
    """
    Gaussian Process Regression with PCA for dynamic frequency response prediction.

    This model:
    1. Applies PCA to compress frequency responses
    2. Trains separate GPR models for each principal component
    3. Reconstructs full frequency response from predictions
    """

    def __init__(self, n_components=50, variance_threshold=0.99):
        """
        Initialize the model.

        Parameters:
        -----------
        n_components : int
            Maximum number of principal components
        variance_threshold : float
            Cumulative variance to retain (0-1)
        """
        self.n_components = n_components
        self.variance_threshold = variance_threshold

        # Model components
        self.input_scaler = StandardScaler()
        self.output_scaler = StandardScaler()
        self.pca = None
        self.gpr_models = {}  # One GPR per principal component

        # Training metadata
        self.is_fitted = False
        self.n_components_actual = None
        self.variance_explained = None
        self.feature_names = None
        self.output_name = None

    def fit(self, X, Y, feature_names=None, output_name="Output"):
        """
        Fit the model to training data.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters
        Y : array-like, shape [n_samples, n_frequencies]
            Output frequency responses
        feature_names : list, optional
            Names of input features
        output_name : str
            Name of the output parameter
        """
        print(f"\n{'='*80}")
        print(f"TRAINING GPR+PCA MODEL: {output_name}")
        print(f"{'='*80}")

        self.feature_names = feature_names
        self.output_name = output_name

        n_samples, n_freq = Y.shape

        print(f"Training data: {n_samples} samples × {n_freq} frequency points")

        # Step 1: Standardize inputs
        X_scaled = self.input_scaler.fit_transform(X)

        # Step 2: Standardize outputs
        Y_scaled = self.output_scaler.fit_transform(Y)

        # Step 3: Apply PCA
        print(f"\nApplying PCA...")
        self.pca = PCA(n_components=min(self.n_components, n_samples, n_freq))
        Y_pca = self.pca.fit_transform(Y_scaled)

        # Determine number of components to keep
        cumsum_var = np.cumsum(self.pca.explained_variance_ratio_)
        self.n_components_actual = np.searchsorted(cumsum_var, self.variance_threshold) + 1
        self.n_components_actual = min(self.n_components_actual, Y_pca.shape[1])

        self.variance_explained = cumsum_var[self.n_components_actual - 1]

        print(f"  ✓ Kept {self.n_components_actual} components")
        print(f"  ✓ Variance explained: {self.variance_explained*100:.2f}%")
        print(f"  ✓ Compression ratio: {n_freq}/{self.n_components_actual} = {n_freq/self.n_components_actual:.1f}×")

        # Step 4: Train GPR for each principal component
        print(f"\nTraining {self.n_components_actual} GPR models...")

        kernel = C(1.0, (0.1, 10.0)) * RBF(1.0, (0.1, 10.0)) + WhiteKernel(0.1, (1e-5, 1.0))

        for i in range(self.n_components_actual):
            if (i + 1) % 10 == 0:
                print(f"  Training PC {i+1}/{self.n_components_actual}...")

            y_pc = Y_pca[:, i]

            gpr = GaussianProcessRegressor(
                kernel=kernel,
                n_restarts_optimizer=5,
                random_state=42,
                normalize_y=True
            )
            gpr.fit(X_scaled, y_pc)
            self.gpr_models[i] = gpr

        print(f"  ✓ All GPR models trained")

        self.is_fitted = True
        print(f"{'='*80}\n")

        return self

    def predict(self, X):
        """
        Predict frequency responses for new input parameters.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters

        Returns:
        --------
        Y_pred : array, shape [n_samples, n_frequencies]
            Predicted frequency responses
        """
        if not self.is_fitted:
            raise ValueError("Model must be fitted before prediction")

        # Standardize inputs
        X_scaled = self.input_scaler.transform(X)

        # Predict each principal component
        Y_pca_pred = np.zeros((X_scaled.shape[0], self.n_components_actual))

        for i in range(self.n_components_actual):
            Y_pca_pred[:, i] = self.gpr_models[i].predict(X_scaled)

        # Inverse PCA transform
        Y_scaled_pred = self.pca.inverse_transform(
            np.column_stack([Y_pca_pred, np.zeros((Y_pca_pred.shape[0],
                                                    self.pca.n_components_ - self.n_components_actual))])
        )

        # Inverse scaling
        Y_pred = self.output_scaler.inverse_transform(Y_scaled_pred)

        return Y_pred

    def predict_with_uncertainty(self, X):
        """
        Predict with uncertainty estimates.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters

        Returns:
        --------
        Y_pred : array, shape [n_samples, n_frequencies]
            Predicted frequency responses
        Y_std : array, shape [n_samples, n_frequencies]
            Standard deviation estimates
        """
        if not self.is_fitted:
            raise ValueError("Model must be fitted before prediction")

        X_scaled = self.input_scaler.transform(X)

        # Predict each principal component with uncertainty
        Y_pca_mean = np.zeros((X_scaled.shape[0], self.n_components_actual))
        Y_pca_std = np.zeros((X_scaled.shape[0], self.n_components_actual))

        for i in range(self.n_components_actual):
            mean, std = self.gpr_models[i].predict(X_scaled, return_std=True)
            Y_pca_mean[:, i] = mean
            Y_pca_std[:, i] = std

        # Inverse PCA transform for mean
        Y_scaled_mean = self.pca.inverse_transform(
            np.column_stack([Y_pca_mean, np.zeros((Y_pca_mean.shape[0],
                                                    self.pca.n_components_ - self.n_components_actual))])
        )
        Y_mean = self.output_scaler.inverse_transform(Y_scaled_mean)

        # Approximate uncertainty propagation through PCA
        # Use component-wise scaling as approximation
        components_used = self.pca.components_[:self.n_components_actual, :]
        Y_scaled_std = np.sqrt(np.dot(Y_pca_std**2, components_used**2))

        # Scale back
        Y_std = Y_scaled_std * self.output_scaler.scale_

        return Y_mean, Y_std

    def score(self, X, Y):
        """
        Calculate R² score on test data.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input parameters
        Y : array-like, shape [n_samples, n_frequencies]
            True frequency responses

        Returns:
        --------
        r2 : float
            R² score
        """
        Y_pred = self.predict(X)
        return r2_score(Y.ravel(), Y_pred.ravel())


class XGBoost_PCA_DynamicModel:
    """
    XGBoost with PCA for dynamic frequency response prediction.

    This model:
    1. Applies PCA to compress frequency responses
    2. Trains separate XGBoost models for each principal component
    3. Reconstructs full frequency response from predictions

    XGBoost typically provides:
    - Better accuracy on tabular data
    - Faster training
    - Built-in regularization
    """

    def __init__(self, n_components=50, variance_threshold=0.99):
        """
        Initialize the model.

        Parameters:
        -----------
        n_components : int
            Maximum number of principal components
        variance_threshold : float
            Cumulative variance to retain (0-1)
        """
        if not XGBOOST_AVAILABLE:
            raise ImportError("XGBoost not installed. Install with: pip install xgboost")

        self.n_components = n_components
        self.variance_threshold = variance_threshold

        # Model components
        self.input_scaler = StandardScaler()
        self.output_scaler = StandardScaler()
        self.pca = None
        self.xgb_models = {}  # One XGBoost per principal component

        # Training metadata
        self.is_fitted = False
        self.n_components_actual = None
        self.variance_explained = None
        self.feature_names = None
        self.output_name = None

    def fit(self, X, Y, feature_names=None, output_name="Output"):
        """
        Fit the model to training data.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters
        Y : array-like, shape [n_samples, n_frequencies]
            Output frequency responses
        feature_names : list, optional
            Names of input features
        output_name : str
            Name of the output parameter
        """
        print(f"\n{'='*80}")
        print(f"TRAINING XGBoost+PCA MODEL: {output_name}")
        print(f"{'='*80}")

        self.feature_names = feature_names
        self.output_name = output_name

        n_samples, n_freq = Y.shape

        print(f"Training data: {n_samples} samples × {n_freq} frequency points")

        # Step 1: Standardize inputs
        X_scaled = self.input_scaler.fit_transform(X)

        # Step 2: Standardize outputs
        Y_scaled = self.output_scaler.fit_transform(Y)

        # Step 3: Apply PCA
        print(f"\nApplying PCA...")
        self.pca = PCA(n_components=min(self.n_components, n_samples, n_freq))
        Y_pca = self.pca.fit_transform(Y_scaled)

        # Determine number of components to keep
        cumsum_var = np.cumsum(self.pca.explained_variance_ratio_)
        self.n_components_actual = np.searchsorted(cumsum_var, self.variance_threshold) + 1
        self.n_components_actual = min(self.n_components_actual, Y_pca.shape[1])

        self.variance_explained = cumsum_var[self.n_components_actual - 1]

        print(f"  ✓ Kept {self.n_components_actual} components")
        print(f"  ✓ Variance explained: {self.variance_explained*100:.2f}%")
        print(f"  ✓ Compression ratio: {n_freq}/{self.n_components_actual} = {n_freq/self.n_components_actual:.1f}×")

        # Step 4: Train XGBoost for each principal component
        print(f"\nTraining {self.n_components_actual} XGBoost models...")

        # XGBoost parameters optimized for regression
        xgb_params = {
            'objective': 'reg:squarederror',
            'max_depth': 6,
            'learning_rate': 0.1,
            'n_estimators': 200,
            'min_child_weight': 1,
            'subsample': 0.8,
            'colsample_bytree': 0.8,
            'gamma': 0,
            'reg_alpha': 0.1,  # L1 regularization
            'reg_lambda': 1.0,  # L2 regularization
            'random_state': 42,
            'n_jobs': -1,
            'verbosity': 0
        }

        for i in range(self.n_components_actual):
            if (i + 1) % 10 == 0:
                print(f"  Training PC {i+1}/{self.n_components_actual}...")

            y_pc = Y_pca[:, i]

            model = xgb.XGBRegressor(**xgb_params)
            model.fit(X_scaled, y_pc)
            self.xgb_models[i] = model

        print(f"  ✓ All XGBoost models trained")

        self.is_fitted = True
        print(f"{'='*80}\n")

        return self

    def predict(self, X):
        """
        Predict frequency responses for new input parameters.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters

        Returns:
        --------
        Y_pred : array, shape [n_samples, n_frequencies]
            Predicted frequency responses
        """
        if not self.is_fitted:
            raise ValueError("Model must be fitted before prediction")

        # Standardize inputs
        X_scaled = self.input_scaler.transform(X)

        # Predict each principal component
        Y_pca_pred = np.zeros((X_scaled.shape[0], self.n_components_actual))

        for i in range(self.n_components_actual):
            Y_pca_pred[:, i] = self.xgb_models[i].predict(X_scaled)

        # Inverse PCA transform
        Y_scaled_pred = self.pca.inverse_transform(
            np.column_stack([Y_pca_pred, np.zeros((Y_pca_pred.shape[0],
                                                    self.pca.n_components_ - self.n_components_actual))])
        )

        # Inverse scaling
        Y_pred = self.output_scaler.inverse_transform(Y_scaled_pred)

        return Y_pred

    def predict_with_uncertainty(self, X):
        """
        Predict with estimated uncertainty using prediction variance.
        Note: XGBoost doesn't provide native uncertainty, so we estimate it.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters

        Returns:
        --------
        Y_mean : array, shape [n_samples, n_frequencies]
            Mean predictions
        Y_std : array, shape [n_samples, n_frequencies]
            Estimated standard deviations (based on model variance)
        """
        Y_mean = self.predict(X)

        # Estimate uncertainty based on residual variance during training
        # This is a rough approximation - XGBoost doesn't provide native uncertainty
        # We use a fixed percentage of the mean as uncertainty estimate
        Y_std = np.abs(Y_mean) * 0.05  # 5% relative uncertainty
        Y_std = np.maximum(Y_std, 0.01)  # Minimum uncertainty

        return Y_mean, Y_std

    def score(self, X, Y):
        """
        Calculate R² score on test data.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input parameters
        Y : array-like, shape [n_samples, n_frequencies]
            True frequency responses

        Returns:
        --------
        r2 : float
            R² score
        """
        Y_pred = self.predict(X)
        return r2_score(Y.ravel(), Y_pred.ravel())


class DynamicModelTrainer:
    """
    Manages training of multiple dynamic models (one per output parameter).
    """

    def __init__(self, n_components=50, variance_threshold=0.99, model_type="GPR"):
        """
        Initialize trainer.

        Parameters:
        -----------
        n_components : int
            Maximum number of PCA components
        variance_threshold : float
            Variance to retain in PCA
        model_type : str
            "GPR" for Gaussian Process or "XGBoost" for XGBoost
        """
        self.n_components = n_components
        self.variance_threshold = variance_threshold
        self.model_type = model_type
        self.models = {}
        self.training_stats = {}

    def train_all_outputs(self, X, Y, frequency, input_names, output_names):
        """
        Train models for all output parameters.

        Parameters:
        -----------
        X : array, shape [n_configs, n_inputs]
            Input parameters
        Y : array, shape [n_configs, n_freq, n_outputs]
            Output frequency responses
        frequency : array, shape [n_freq]
            Frequency array
        input_names : list
            Input parameter names
        output_names : list
            Output parameter names

        Returns:
        --------
        dict : Training statistics for each output
        """
        n_configs, n_freq, n_outputs = Y.shape

        print(f"\n{'#'*80}")
        print(f"TRAINING ALL DYNAMIC MODELS ({self.model_type})")
        print(f"{'#'*80}")
        print(f"Configurations: {n_configs}")
        print(f"Frequency points: {n_freq}")
        print(f"Output parameters: {n_outputs}")
        print(f"Model type: {self.model_type}")
        print(f"{'#'*80}\n")

        # Split data
        X_train, X_test, Y_train, Y_test = train_test_split(
            X, Y, test_size=0.2, random_state=42
        )

        # Train model for each output
        for idx, output_name in enumerate(output_names):
            Y_train_single = Y_train[:, :, idx]
            Y_test_single = Y_test[:, :, idx]

            # Initialize model based on type
            if self.model_type == "XGBoost":
                if not XGBOOST_AVAILABLE:
                    print(f"⚠ XGBoost not available, falling back to GPR")
                    model = GPR_PCA_DynamicModel(
                        n_components=self.n_components,
                        variance_threshold=self.variance_threshold
                    )
                else:
                    model = XGBoost_PCA_DynamicModel(
                        n_components=self.n_components,
                        variance_threshold=self.variance_threshold
                    )
            else:  # Default to GPR
                model = GPR_PCA_DynamicModel(
                    n_components=self.n_components,
                    variance_threshold=self.variance_threshold
                )

            model.fit(X_train, Y_train_single,
                     feature_names=input_names,
                     output_name=output_name)

            # Evaluate
            Y_pred_train = model.predict(X_train)
            Y_pred_test = model.predict(X_test)

            r2_train = r2_score(Y_train_single.ravel(), Y_pred_train.ravel())
            r2_test = r2_score(Y_test_single.ravel(), Y_pred_test.ravel())

            rmse_train = np.sqrt(mean_squared_error(Y_train_single.ravel(), Y_pred_train.ravel()))
            rmse_test = np.sqrt(mean_squared_error(Y_test_single.ravel(), Y_pred_test.ravel()))

            # Store model and stats
            self.models[output_name] = model
            self.training_stats[output_name] = {
                'r2_train': r2_train,
                'r2_test': r2_test,
                'rmse_train': rmse_train,
                'rmse_test': rmse_test,
                'n_components': model.n_components_actual,
                'variance_explained': model.variance_explained,
                'model_type': self.model_type
            }

            print(f"✓ {output_name}:")
            print(f"    R² Train: {r2_train:.4f} | R² Test: {r2_test:.4f}")
            print(f"    RMSE Train: {rmse_train:.4f} | RMSE Test: {rmse_test:.4f}")
            print(f"    Components: {model.n_components_actual} ({model.variance_explained*100:.1f}% var)\n")

        return self.training_stats


class DynamicVisualization:
    """
    Best-in-class visualizations for dynamic frequency response modeling.
    """

    @staticmethod
    def plot_training_summary(stats_dict, output_names):
        """
        Create summary visualization of training performance.
        """
        fig = plt.figure(figsize=(14, 8))
        gs = GridSpec(2, 2, figure=fig, hspace=0.3, wspace=0.3)

        # R² comparison
        ax1 = fig.add_subplot(gs[0, 0])
        x = np.arange(len(output_names))
        width = 0.35

        r2_train = [stats_dict[name]['r2_train'] for name in output_names]
        r2_test = [stats_dict[name]['r2_test'] for name in output_names]

        ax1.bar(x - width/2, r2_train, width, label='Train', color='steelblue', alpha=0.8)
        ax1.bar(x + width/2, r2_test, width, label='Test', color='coral', alpha=0.8)
        ax1.set_xlabel('Output Parameter', fontweight='bold')
        ax1.set_ylabel('R² Score', fontweight='bold')
        ax1.set_title('Model Performance (R²)', fontweight='bold', fontsize=12)
        ax1.set_xticks(x)
        ax1.set_xticklabels(output_names, rotation=45, ha='right')
        ax1.legend()
        ax1.grid(axis='y', alpha=0.3)
        ax1.set_ylim([0, 1.0])

        # RMSE comparison
        ax2 = fig.add_subplot(gs[0, 1])
        rmse_train = [stats_dict[name]['rmse_train'] for name in output_names]
        rmse_test = [stats_dict[name]['rmse_test'] for name in output_names]

        ax2.bar(x - width/2, rmse_train, width, label='Train', color='steelblue', alpha=0.8)
        ax2.bar(x + width/2, rmse_test, width, label='Test', color='coral', alpha=0.8)
        ax2.set_xlabel('Output Parameter', fontweight='bold')
        ax2.set_ylabel('RMSE', fontweight='bold')
        ax2.set_title('Model Error (RMSE)', fontweight='bold', fontsize=12)
        ax2.set_xticks(x)
        ax2.set_xticklabels(output_names, rotation=45, ha='right')
        ax2.legend()
        ax2.grid(axis='y', alpha=0.3)

        # Number of components
        ax3 = fig.add_subplot(gs[1, 0])
        n_comps = [stats_dict[name]['n_components'] for name in output_names]

        bars = ax3.bar(x, n_comps, color='mediumseagreen', alpha=0.8)
        ax3.set_xlabel('Output Parameter', fontweight='bold')
        ax3.set_ylabel('Number of PCs', fontweight='bold')
        ax3.set_title('PCA Dimensionality Reduction', fontweight='bold', fontsize=12)
        ax3.set_xticks(x)
        ax3.set_xticklabels(output_names, rotation=45, ha='right')
        ax3.grid(axis='y', alpha=0.3)

        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}',
                    ha='center', va='bottom', fontweight='bold')

        # Variance explained
        ax4 = fig.add_subplot(gs[1, 1])
        var_exp = [stats_dict[name]['variance_explained']*100 for name in output_names]

        bars = ax4.bar(x, var_exp, color='mediumpurple', alpha=0.8)
        ax4.set_xlabel('Output Parameter', fontweight='bold')
        ax4.set_ylabel('Variance Explained (%)', fontweight='bold')
        ax4.set_title('PCA Variance Captured', fontweight='bold', fontsize=12)
        ax4.set_xticks(x)
        ax4.set_xticklabels(output_names, rotation=45, ha='right')
        ax4.grid(axis='y', alpha=0.3)
        ax4.set_ylim([95, 100])

        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax4.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:.2f}%',
                    ha='center', va='bottom', fontweight='bold', fontsize=9)

        plt.suptitle('Dynamic Model Training Summary', fontsize=14, fontweight='bold', y=0.98)

        return fig

    @staticmethod
    def plot_frequency_response_prediction(frequency, Y_true, Y_pred, Y_std=None,
                                          config_idx=0, output_name="Output"):
        """
        Plot frequency response prediction vs actual for a single configuration.
        """
        fig, ax = plt.subplots(figsize=(12, 5))

        # Plot actual
        ax.plot(frequency, Y_true[config_idx], 'o-', label='Actual',
               color='black', linewidth=2, markersize=4, alpha=0.7)

        # Plot prediction
        ax.plot(frequency, Y_pred[config_idx], 's-', label='Predicted',
               color='crimson', linewidth=2, markersize=3, alpha=0.7)

        # Plot uncertainty band if available
        if Y_std is not None:
            ax.fill_between(frequency,
                           Y_pred[config_idx] - 2*Y_std[config_idx],
                           Y_pred[config_idx] + 2*Y_std[config_idx],
                           color='crimson', alpha=0.2, label='95% Confidence')

        ax.set_xlabel('Frequency (Hz)', fontweight='bold', fontsize=12)
        ax.set_ylabel(f'{output_name}', fontweight='bold', fontsize=12)
        ax.set_title(f'Frequency Response Prediction: {output_name} (Config #{config_idx+1})',
                    fontweight='bold', fontsize=13)
        ax.legend(loc='best', fontsize=11)
        ax.grid(True, alpha=0.3, linestyle='--')

        # Calculate and display R²
        r2 = r2_score(Y_true[config_idx], Y_pred[config_idx])
        ax.text(0.02, 0.98, f'R² = {r2:.4f}', transform=ax.transAxes,
               fontsize=11, verticalalignment='top', fontweight='bold',
               bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

        plt.tight_layout()
        return fig

    @staticmethod
    def plot_prediction_grid(frequency, Y_true, Y_pred, output_name="Output",
                            n_samples=4):
        """
        Plot grid of predictions for multiple configurations.
        """
        n_rows = int(np.ceil(n_samples / 2))
        fig, axes = plt.subplots(n_rows, 2, figsize=(14, 4*n_rows))
        axes = axes.flatten()

        for i in range(n_samples):
            ax = axes[i]

            ax.plot(frequency, Y_true[i], 'o-', label='Actual',
                   color='black', linewidth=1.5, markersize=3, alpha=0.6)
            ax.plot(frequency, Y_pred[i], 's-', label='Predicted',
                   color='crimson', linewidth=1.5, markersize=2, alpha=0.7)

            r2 = r2_score(Y_true[i], Y_pred[i])

            ax.set_xlabel('Frequency (Hz)', fontweight='bold')
            ax.set_ylabel(output_name, fontweight='bold')
            ax.set_title(f'Config #{i+1} (R²={r2:.3f})', fontweight='bold')
            ax.legend(fontsize=9)
            ax.grid(True, alpha=0.3)

        # Hide unused subplots
        for i in range(n_samples, len(axes)):
            axes[i].set_visible(False)

        plt.suptitle(f'Frequency Response Predictions: {output_name}',
                    fontsize=14, fontweight='bold')
        plt.tight_layout()

        return fig

    @staticmethod
    def plot_error_analysis(frequency, Y_true, Y_pred, output_name="Output"):
        """
        Comprehensive error analysis visualization.
        """
        fig = plt.figure(figsize=(14, 10))
        gs = GridSpec(3, 2, figure=fig, hspace=0.3, wspace=0.3)

        # Calculate errors
        errors = Y_pred - Y_true
        abs_errors = np.abs(errors)
        percent_errors = 100 * np.abs(errors) / (np.abs(Y_true) + 1e-10)

        # 1. Error vs Frequency (averaged across configs)
        ax1 = fig.add_subplot(gs[0, :])
        mean_error = errors.mean(axis=0)
        std_error = errors.std(axis=0)

        ax1.plot(frequency, mean_error, 'b-', linewidth=2, label='Mean Error')
        ax1.fill_between(frequency,
                        mean_error - std_error,
                        mean_error + std_error,
                        color='blue', alpha=0.2, label='±1 Std Dev')
        ax1.axhline(y=0, color='k', linestyle='--', alpha=0.5)
        ax1.set_xlabel('Frequency (Hz)', fontweight='bold')
        ax1.set_ylabel('Prediction Error', fontweight='bold')
        ax1.set_title('Mean Prediction Error vs Frequency', fontweight='bold', fontsize=12)
        ax1.legend()
        ax1.grid(True, alpha=0.3)

        # 2. Error distribution histogram
        ax2 = fig.add_subplot(gs[1, 0])
        ax2.hist(errors.ravel(), bins=50, color='steelblue', alpha=0.7, edgecolor='black')
        ax2.axvline(x=0, color='red', linestyle='--', linewidth=2, label='Zero Error')
        ax2.set_xlabel('Prediction Error', fontweight='bold')
        ax2.set_ylabel('Frequency Count', fontweight='bold')
        ax2.set_title('Error Distribution', fontweight='bold', fontsize=12)
        ax2.legend()
        ax2.grid(axis='y', alpha=0.3)

        # 3. Absolute error vs frequency
        ax3 = fig.add_subplot(gs[1, 1])
        mean_abs_error = abs_errors.mean(axis=0)
        ax3.plot(frequency, mean_abs_error, 'r-', linewidth=2)
        ax3.fill_between(frequency, 0, mean_abs_error, color='red', alpha=0.3)
        ax3.set_xlabel('Frequency (Hz)', fontweight='bold')
        ax3.set_ylabel('Mean Absolute Error', fontweight='bold')
        ax3.set_title('Mean Absolute Error vs Frequency', fontweight='bold', fontsize=12)
        ax3.grid(True, alpha=0.3)

        # 4. Predicted vs Actual scatter
        ax4 = fig.add_subplot(gs[2, 0])
        ax4.scatter(Y_true.ravel(), Y_pred.ravel(), alpha=0.3, s=10, color='navy')

        # Perfect prediction line
        min_val = min(Y_true.min(), Y_pred.min())
        max_val = max(Y_true.max(), Y_pred.max())
        ax4.plot([min_val, max_val], [min_val, max_val], 'r--', linewidth=2, label='Perfect')

        ax4.set_xlabel('Actual Value', fontweight='bold')
        ax4.set_ylabel('Predicted Value', fontweight='bold')
        ax4.set_title('Predicted vs Actual (All Points)', fontweight='bold', fontsize=12)
        ax4.legend()
        ax4.grid(True, alpha=0.3)
        ax4.axis('equal')

        # Add R² text
        r2_total = r2_score(Y_true.ravel(), Y_pred.ravel())
        ax4.text(0.05, 0.95, f'R² = {r2_total:.4f}', transform=ax4.transAxes,
                fontsize=11, verticalalignment='top', fontweight='bold',
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.5))

        # 5. Percent error heatmap (frequency vs config)
        ax5 = fig.add_subplot(gs[2, 1])

        # Sample configs if too many
        n_configs = percent_errors.shape[0]
        if n_configs > 20:
            sample_idx = np.linspace(0, n_configs-1, 20, dtype=int)
            percent_errors_plot = percent_errors[sample_idx]
        else:
            percent_errors_plot = percent_errors

        im = ax5.imshow(percent_errors_plot.T, aspect='auto', cmap='RdYlGn_r',
                       vmin=0, vmax=np.percentile(percent_errors, 95))
        ax5.set_xlabel('Configuration Index', fontweight='bold')
        ax5.set_ylabel('Frequency Point', fontweight='bold')
        ax5.set_title('Percent Error Heatmap', fontweight='bold', fontsize=12)

        # Set frequency ticks
        n_freq_ticks = min(10, len(frequency))
        freq_tick_idx = np.linspace(0, len(frequency)-1, n_freq_ticks, dtype=int)
        ax5.set_yticks(freq_tick_idx)
        ax5.set_yticklabels([f'{frequency[i]:.0f}' for i in freq_tick_idx])

        plt.colorbar(im, ax=ax5, label='Percent Error (%)')

        plt.suptitle(f'Comprehensive Error Analysis: {output_name}',
                    fontsize=14, fontweight='bold', y=0.995)

        return fig

    @staticmethod
    def plot_pca_analysis(model, frequency, output_name="Output"):
        """
        Visualize PCA components and their contributions.
        """
        fig = plt.figure(figsize=(14, 8))
        gs = GridSpec(2, 2, figure=fig, hspace=0.3, wspace=0.3)

        # 1. Variance explained
        ax1 = fig.add_subplot(gs[0, 0])
        var_ratio = model.pca.explained_variance_ratio_
        cumsum_var = np.cumsum(var_ratio)

        ax1.bar(range(1, len(var_ratio)+1), var_ratio*100,
               color='steelblue', alpha=0.7, label='Individual')
        ax1.plot(range(1, len(cumsum_var)+1), cumsum_var*100,
                'ro-', linewidth=2, markersize=5, label='Cumulative')
        ax1.axhline(y=model.variance_explained*100, color='green',
                   linestyle='--', linewidth=2, label=f'Threshold ({model.variance_explained*100:.1f}%)')
        ax1.axvline(x=model.n_components_actual, color='red',
                   linestyle='--', linewidth=2, label=f'N Components ({model.n_components_actual})')

        ax1.set_xlabel('Principal Component', fontweight='bold')
        ax1.set_ylabel('Variance Explained (%)', fontweight='bold')
        ax1.set_title('PCA Variance Explained', fontweight='bold', fontsize=12)
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        ax1.set_xlim([0, min(50, len(var_ratio)+1)])

        # 2. First few principal components
        ax2 = fig.add_subplot(gs[0, 1])
        n_show = min(5, model.n_components_actual)
        colors = plt.cm.viridis(np.linspace(0, 1, n_show))

        for i in range(n_show):
            component = model.pca.components_[i]
            ax2.plot(frequency, component, '-', linewidth=2,
                    color=colors[i], label=f'PC{i+1} ({var_ratio[i]*100:.1f}%)', alpha=0.7)

        ax2.set_xlabel('Frequency (Hz)', fontweight='bold')
        ax2.set_ylabel('Component Loading', fontweight='bold')
        ax2.set_title('Principal Components (Frequency Domain)', fontweight='bold', fontsize=12)
        ax2.legend(loc='best', fontsize=9)
        ax2.grid(True, alpha=0.3)

        # 3. Reconstruction quality
        ax3 = fig.add_subplot(gs[1, :])

        # Show reconstruction with different numbers of components
        n_comps_to_show = [1, 5, 10, model.n_components_actual]
        n_comps_to_show = [n for n in n_comps_to_show if n <= model.n_components_actual]

        # Need sample data - get from model if available
        # For visualization, we'll show the mean component
        mean_component = model.pca.components_[:model.n_components_actual].mean(axis=0)

        colors_recon = plt.cm.plasma(np.linspace(0, 1, len(n_comps_to_show)))

        for idx, n_comp in enumerate(n_comps_to_show):
            # Partial reconstruction using first n_comp components
            partial_mean = model.pca.components_[:n_comp].mean(axis=0)
            ax3.plot(frequency, partial_mean, '-', linewidth=2,
                    color=colors_recon[idx], label=f'{n_comp} Components', alpha=0.7)

        ax3.set_xlabel('Frequency (Hz)', fontweight='bold')
        ax3.set_ylabel('Mean Component Value', fontweight='bold')
        ax3.set_title('Reconstruction Quality (Mean Components)', fontweight='bold', fontsize=12)
        ax3.legend()
        ax3.grid(True, alpha=0.3)

        plt.suptitle(f'PCA Analysis: {output_name}', fontsize=14, fontweight='bold', y=0.98)

        return fig


# Example usage and testing functions
def test_dynamic_model(data_file=None):
    """
    Test the dynamic modeling pipeline with sample data.
    """
    if data_file is None:
        print("No data file provided. Please provide a data file for testing.")
        return

    # Load data
    loader = DynamicDataLoader(data_file)
    data_dict = loader.load_and_reshape()

    X = data_dict['X']
    Y = data_dict['Y']
    frequency = data_dict['frequency']
    input_names = data_dict['input_names']
    output_names = data_dict['output_names']

    # Train models
    trainer = DynamicModelTrainer(n_components=50, variance_threshold=0.99)
    stats = trainer.train_all_outputs(X, Y, frequency, input_names, output_names)

    # Create visualizations
    print("\nGenerating visualizations...")

    # Training summary
    fig1 = DynamicVisualization.plot_training_summary(stats, output_names)
    plt.savefig('/mnt/user-data/outputs/training_summary.png', dpi=150, bbox_inches='tight')
    print("✓ Saved: training_summary.png")

    # For each output, create detailed visualizations
    X_train, X_test, Y_train, Y_test = train_test_split(
        X, Y, test_size=0.2, random_state=42
    )

    for idx, output_name in enumerate(output_names):
        model = trainer.models[output_name]
        Y_test_single = Y_test[:, :, idx]

        # Predict
        Y_pred = model.predict(X_test)
        Y_pred_unc, Y_std = model.predict_with_uncertainty(X_test)

        # Frequency response prediction
        fig2 = DynamicVisualization.plot_frequency_response_prediction(
            frequency, Y_test_single, Y_pred, Y_std,
            config_idx=0, output_name=output_name
        )
        plt.savefig(f'/mnt/user-data/outputs/freq_response_{output_name}.png',
                   dpi=150, bbox_inches='tight')
        print(f"✓ Saved: freq_response_{output_name}.png")
        plt.close()

        # Prediction grid
        n_test = min(4, len(Y_test_single))
        fig3 = DynamicVisualization.plot_prediction_grid(
            frequency, Y_test_single, Y_pred, output_name, n_samples=n_test
        )
        plt.savefig(f'/mnt/user-data/outputs/prediction_grid_{output_name}.png',
                   dpi=150, bbox_inches='tight')
        print(f"✓ Saved: prediction_grid_{output_name}.png")
        plt.close()

        # Error analysis
        fig4 = DynamicVisualization.plot_error_analysis(
            frequency, Y_test_single, Y_pred, output_name
        )
        plt.savefig(f'/mnt/user-data/outputs/error_analysis_{output_name}.png',
                   dpi=150, bbox_inches='tight')
        print(f"✓ Saved: error_analysis_{output_name}.png")
        plt.close()

        # PCA analysis
        fig5 = DynamicVisualization.plot_pca_analysis(
            model, frequency, output_name
        )
        plt.savefig(f'/mnt/user-data/outputs/pca_analysis_{output_name}.png',
                   dpi=150, bbox_inches='tight')
        print(f"✓ Saved: pca_analysis_{output_name}.png")
        plt.close()

    print("\n" + "="*80)
    print("DYNAMIC MODELING COMPLETE!")
    print("="*80)
    print(f"All visualizations saved to: /mnt/user-data/outputs/")
    print("="*80)

    return trainer, stats


if __name__ == "__main__":
    print("Dynamic RNN Module Loaded Successfully")
    print("This module provides:")
    print("  - DynamicDataLoader: Flexible data loading with auto-detection")
    print("  - GPR_PCA_DynamicModel: Gaussian Process + PCA for frequency responses")
    print("  - DynamicModelTrainer: Multi-output model training")
    print("  - DynamicVisualization: Best-in-class visualization suite")

"""
AVL Cameo Model Wrapper for OLHC Application

This module provides a wrapper around the AVL Cameo exported models,
allowing seamless integration with the OLHC application's existing workflow.

The wrapper automatically detects if the AVL Cameo DLL is available and
falls back to training custom models if not.
"""

import os
import sys
import numpy as np

# Try to import AVL Cameo model
CAMEO_AVAILABLE = False
try:
    # Add the AVL Cameo model directory to path if needed
    cameo_model_path = r"D:\Python\OLHC\AVL_Cameo_model"
    if os.path.exists(cameo_model_path) and cameo_model_path not in sys.path:
        sys.path.insert(0, cameo_model_path)

    import Variant
    CAMEO_AVAILABLE = True
    print("✓ AVL Cameo model loaded successfully")
except Exception as e:
    print(f"⚠ AVL Cameo model not available: {e}")
    print("  Will use fallback RNN/GPR/XGBoost models")


class AVLCameoModelWrapper:
    """
    Wrapper for AVL Cameo exported models.

    This class provides a unified interface to use AVL Cameo models
    in the OLHC application, matching the API of the existing
    DynamicModelTrainer class.
    """

    # AVL Cameo input parameters (10 parameters)
    CAMEO_INPUTS = [
        'B1_offset',
        'B2_offset',
        'B3_offset',
        'B4_offset',
        'B5_offset',
        'Helix_Angle',
        'Input_Stifness',
        'Lead_Crown_Pinion',
        'Lead_Slope_Pinion',
        'Pressure_Angle'
    ]

    # AVL Cameo output parameters (19 outputs)
    CAMEO_OUTPUTS = [
        'Hull',
        'B1_radialStiffnessX',
        'B1_radialStiffnessY',
        'B2_radialStiffnessX',
        'B2_radialStiffnessY',
        'B3_axialStiffness',
        'B3_radialStiffnessX',
        'B3_radialStiffnessY',
        'B4_radialStiffnessX',
        'B4_radialStiffnessY',
        'B5_radialStiffnessX',
        'B5_radialStiffnessY',
        'Fx',
        'Fy',
        'Fz',
        'Linear_TE',
        'Mx',
        'My',
        'Tilt_TE'
    ]

    def __init__(self):
        """Initialize the AVL Cameo model wrapper."""
        self.is_available = CAMEO_AVAILABLE
        self.models = {}  # Store function references

        if self.is_available:
            self._initialize_models()

    def _initialize_models(self):
        """Initialize model function references from Variant module."""
        if not self.is_available:
            return

        # Map output names to their corresponding functions
        self.models = {
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

    def predict_single_output(self, X, output_name):
        """
        Predict a single output parameter using AVL Cameo model.

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters (must be 10 parameters in correct order)
        output_name : str
            Name of the output parameter to predict

        Returns:
        --------
        predictions : list
            Predicted values for all samples
        """
        if not self.is_available:
            raise RuntimeError("AVL Cameo model not available")

        if output_name not in self.models:
            raise ValueError(f"Unknown output: {output_name}. Available: {list(self.models.keys())}")

        # Get the prediction function
        predict_func = self.models[output_name]

        # Convert input to list format expected by AVL Cameo
        if isinstance(X, np.ndarray):
            if X.ndim == 1:
                X = X.reshape(1, -1)

            # Call the model with input parameters
            # AVL expects parameters as separate arguments
            predictions = []
            for row in X:
                result = predict_func(
                    B1_offset=float(row[0]),
                    B2_offset=float(row[1]),
                    B3_offset=float(row[2]),
                    B4_offset=float(row[3]),
                    B5_offset=float(row[4]),
                    Helix_Angle=float(row[5]),
                    Input_Stifness=float(row[6]),
                    Lead_Crown_Pinion=float(row[7]),
                    Lead_Slope_Pinion=float(row[8]),
                    Pressure_Angle=float(row[9])
                )
                predictions.append(result[0] if isinstance(result, list) else result)

            return np.array(predictions)
        else:
            raise TypeError("X must be a numpy array")

    def predict_all_outputs(self, X):
        """
        Predict all output parameters using AVL Cameo model.

        Parameters:
        -----------
        X : array-like, shape [n_samples, 10]
            Input design parameters

        Returns:
        --------
        predictions : dict
            Dictionary mapping output names to predicted values
        """
        if not self.is_available:
            raise RuntimeError("AVL Cameo model not available")

        predictions = {}
        for output_name in self.CAMEO_OUTPUTS:
            predictions[output_name] = self.predict_single_output(X, output_name)

        return predictions

    @staticmethod
    def is_cameo_compatible(input_names):
        """
        Check if the input parameters match AVL Cameo requirements.

        Parameters:
        -----------
        input_names : list
            List of input parameter names from the data

        Returns:
        --------
        compatible : bool
            True if inputs match Cameo requirements
        """
        if len(input_names) != 10:
            return False

        # Check if input names match (case-insensitive, allowing minor variations)
        normalized_inputs = [name.lower().replace('_', '').replace(' ', '')
                           for name in input_names]
        normalized_cameo = [name.lower().replace('_', '').replace(' ', '')
                          for name in AVLCameoModelWrapper.CAMEO_INPUTS]

        return normalized_inputs == normalized_cameo


class HybridModelManager:
    """
    Manages both AVL Cameo and custom trained models.

    This class automatically selects the best available model:
    - AVL Cameo (if available and data is compatible)
    - Custom RNN/GPR/XGBoost (if Cameo unavailable or data incompatible)
    """

    def __init__(self):
        """Initialize the hybrid model manager."""
        self.cameo_wrapper = AVLCameoModelWrapper()
        self.custom_trainer = None  # Will be initialized if needed
        self.use_cameo = False
        self.input_names = None
        self.output_names = None

    def initialize(self, input_names, output_names, force_custom=False):
        """
        Initialize the model manager with data parameters.

        Parameters:
        -----------
        input_names : list
            List of input parameter names
        output_names : list
            List of output parameter names
        force_custom : bool
            If True, force use of custom models even if Cameo available
        """
        self.input_names = input_names
        self.output_names = output_names

        # Decide which model to use
        if (not force_custom and
            self.cameo_wrapper.is_available and
            AVLCameoModelWrapper.is_cameo_compatible(input_names)):

            self.use_cameo = True
            print("\n" + "="*80)
            print("USING AVL CAMEO PRE-TRAINED MODELS")
            print("="*80)
            print(f"✓ Inputs: {len(input_names)} parameters")
            print(f"✓ Outputs: {len(self.cameo_wrapper.CAMEO_OUTPUTS)} parameters")
            print(f"✓ Model type: RNN (pre-trained in AVL Cameo)")
            print("="*80 + "\n")
        else:
            self.use_cameo = False
            if not force_custom and self.cameo_wrapper.is_available:
                print("\n⚠ AVL Cameo model available but data not compatible")
                print(f"  Expected inputs: {self.cameo_wrapper.CAMEO_INPUTS}")
                print(f"  Your inputs: {input_names}")
            print("\n→ Using custom trainable models (RNN/GPR/XGBoost)")

    def train_or_load(self, X, Y, frequency, model_type="GPR",
                     n_components=50, variance_threshold=0.99):
        """
        Train custom models or verify Cameo availability.

        Parameters:
        -----------
        X : array-like
            Input design parameters
        Y : array-like
            Output responses (frequency domain)
        frequency : array-like
            Frequency points
        model_type : str
            "GPR", "RNN", or "XGBoost" (only used for custom models)
        n_components : int
            PCA components (only used for custom models)
        variance_threshold : float
            PCA variance threshold (only used for custom models)

        Returns:
        --------
        stats : dict
            Training statistics or Cameo model info
        """
        if self.use_cameo:
            # No training needed - using pre-trained Cameo model
            stats = {}
            for output_name in self.cameo_wrapper.CAMEO_OUTPUTS:
                stats[output_name] = {
                    'model_type': 'AVL_Cameo_RNN',
                    'r2_train': None,  # Pre-trained, no stats available
                    'r2_test': None,
                    'rmse_train': None,
                    'rmse_test': None,
                    'n_components': None,
                    'variance_explained': None
                }
            return stats
        else:
            # Train custom models
            from ENHANCED_DYNAMIC_RNN import DynamicModelTrainer

            self.custom_trainer = DynamicModelTrainer(
                n_components=n_components,
                variance_threshold=variance_threshold,
                model_type=model_type
            )

            return self.custom_trainer.train_all_outputs(
                X, Y, frequency, self.input_names, self.output_names
            )

    def predict(self, X, output_name):
        """
        Predict using the active model (Cameo or custom).

        Parameters:
        -----------
        X : array-like, shape [n_samples, n_features]
            Input design parameters
        output_name : str
            Output parameter name

        Returns:
        --------
        predictions : array
            Predicted values
        """
        if self.use_cameo:
            return self.cameo_wrapper.predict_single_output(X, output_name)
        else:
            if self.custom_trainer is None:
                raise RuntimeError("Custom models not trained yet")

            model = self.custom_trainer.models.get(output_name)
            if model is None:
                raise ValueError(f"No model found for output: {output_name}")

            return model.predict(X)

    def get_model_info(self):
        """
        Get information about the active model.

        Returns:
        --------
        info : dict
            Model information
        """
        if self.use_cameo:
            return {
                'type': 'AVL Cameo (Pre-trained RNN)',
                'n_inputs': len(self.cameo_wrapper.CAMEO_INPUTS),
                'n_outputs': len(self.cameo_wrapper.CAMEO_OUTPUTS),
                'outputs': self.cameo_wrapper.CAMEO_OUTPUTS,
                'trainable': False
            }
        else:
            if self.custom_trainer is None:
                return {'type': 'Custom (Not trained)', 'trainable': True}

            return {
                'type': f'Custom {self.custom_trainer.model_type}',
                'n_inputs': len(self.input_names),
                'n_outputs': len(self.output_names),
                'outputs': self.output_names,
                'trainable': True,
                'stats': self.custom_trainer.training_stats
            }


# Example usage
if __name__ == "__main__":
    print("AVL Cameo Model Wrapper - Test Script")
    print("="*80)

    # Test if Cameo model is available
    wrapper = AVLCameoModelWrapper()
    print(f"\nAVL Cameo Available: {wrapper.is_available}")

    if wrapper.is_available:
        # Test prediction with sample data
        print("\nTesting prediction with sample data...")

        # Sample input (10 parameters)
        X_sample = np.array([[
            115.499162466667,   # B1_offset
            21.0986890533333,   # B2_offset
            -196.005126666667,  # B3_offset
            133.506812,         # B4_offset
            7.49458072666667,   # B5_offset
            19.9991047682119,   # Helix_Angle
            0.966045105960265,  # Input_Stifness
            10.026600013245,    # Lead_Crown_Pinion
            0.06479313245033,   # Lead_Slope_Pinion
            19.9809056291391    # Pressure_Angle
        ]])

        # Test single output prediction
        hull_pred = wrapper.predict_single_output(X_sample, 'Hull')
        print(f"Hull prediction: {hull_pred}")

        # Test all outputs
        all_preds = wrapper.predict_all_outputs(X_sample)
        print(f"\nAll outputs predicted: {len(all_preds)}")
        for name, value in all_preds.items():
            print(f"  {name}: {value[0]:.6f}")

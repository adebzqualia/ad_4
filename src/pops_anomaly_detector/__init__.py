"""POPS structural anomaly detector."""

__version__ = "0.1.0"

from .analysis import analyze_directories
from .config import AnalysisConfig

__all__ = ["AnalysisConfig", "analyze_directories"]

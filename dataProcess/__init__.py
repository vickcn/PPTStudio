import sys
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    from .ContextParser import ppt_parser as ptp  # type: ignore
except Exception:
    from ContextParser import ppt_parser as ptp  # type: ignore

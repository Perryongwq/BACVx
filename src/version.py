"""
Version information for Block Cutting Automation (BAC)
"""
__version__ = "1.2.0"
__version_info__ = (1, 2, 0)
__build_date__ = "2026-01-26"
__author__ = "BAC Development Team"

def get_version():
    """Return the version string."""
    return __version__

def get_version_info():
    """Return the version as a tuple."""
    return __version_info__

def get_full_version():
    """Return full version information as a string."""
    return f"BAC v{__version__} (Build: {__build_date__})"

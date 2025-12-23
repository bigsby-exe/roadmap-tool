"""
Backward compatibility wrapper - use 'roadmap-ppt' command after installation.
This file allows direct execution: python main.py
"""

import sys
import os

# Add src to path for direct execution
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

try:
    from roadmap_ppt.cli import main
except ImportError:
    # Fallback if package structure isn't found
    print("Error: Could not import roadmap_ppt package.")
    print("Please install the package with: uv tool install .")
    sys.exit(1)

if __name__ == "__main__":
    main()

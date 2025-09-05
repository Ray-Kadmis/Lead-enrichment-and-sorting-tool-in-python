#!/usr/bin/env python3
"""
Simple launcher for Excel Sorter GUI
Double-click this file to run the application
"""

import sys
import os

# Add current directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

try:
    from excel_sorter_gui import main
    main()
except ImportError as e:
    print(f"Error importing required modules: {e}")
    print("Please install required dependencies:")
    print("pip install pandas openpyxl xlsxwriter")
    input("Press Enter to exit...")
except Exception as e:
    print(f"Error running application: {e}")
    input("Press Enter to exit...")

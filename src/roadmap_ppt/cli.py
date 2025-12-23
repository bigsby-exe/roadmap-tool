"""
Command-line interface for the roadmap PowerPoint generator.
"""

import argparse
import os
from .generator import generate_presentation


def main():
    """Main function to orchestrate PowerPoint generation."""
    parser = argparse.ArgumentParser(
        description='Generate PowerPoint presentation from Excel roadmap file'
    )
    parser.add_argument(
        'excel_file',
        type=str,
        help='Path to Excel file containing Objectives and Roadmap sheets'
    )
    parser.add_argument(
        '-o', '--output',
        type=str,
        default=None,
        help='Output PowerPoint file path (default: same as Excel file with .pptx extension)'
    )
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.excel_file):
        print(f"Error: Excel file '{args.excel_file}' not found.")
        return
    
    print(f"Reading Excel file: {args.excel_file}")
    
    # Generate presentation
    generate_presentation(args.excel_file, args.output)


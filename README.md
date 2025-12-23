# Excel to PowerPoint Roadmap Generator

A Python script that automatically converts Excel roadmap data into a professional, brandable PowerPoint presentation. Perfect for creating executive presentations from structured roadmap data.

## Features

- **Automatic PowerPoint Generation**: Converts Excel roadmap data into a polished presentation
- **Brandable Design**: Fully customizable colors, fonts, logos, and styling
- **Smart Data Parsing**: Automatically detects columns in your Excel sheets (case-insensitive)
- **Organized Slides**: Creates multiple slides grouped by timeline/phase
- **Professional Layout**: Modern design with rounded rectangles, proper spacing, and visual hierarchy

## Installation

### Prerequisites

- Python 3.7 or higher
- [UV](https://github.com/astral-sh/uv) (recommended) or pip

### Option 1: Install as UV Tool (Recommended)

Install the tool globally using UV:

```bash
uv tool install .
```

After installation, you can run it from anywhere:

```bash
roadmap-ppt your_roadmap.xlsx
roadmap-ppt your_roadmap.xlsx -o output.pptx
```

To install from a git repository:

```bash
uv tool install git+https://github.com/yourusername/roadmap-ppt-generator.git
```

### Option 2: Install with pip

```bash
pip install -r requirements.txt
```

Then run the script directly:

```bash
python main.py your_roadmap.xlsx
```

### Dependencies

This project requires:
- `python-pptx>=0.6.21` - For PowerPoint generation
- `pandas>=2.0.0` - For Excel file reading
- `openpyxl>=3.1.0` - For Excel file support

## Excel File Format

Your Excel file must contain at least two sheets:

### 1. Objectives Sheet

This sheet should have two columns:

| North star | Key elements |
|------------|--------------|
| Your main objective statement | Element 1 |
| | Element 2 |
| | Element 3 |

- **First column**: Contains your "North star" objective (typically one row)
- **Second column**: Contains key elements as a list (multiple rows)

**Note**: Column names are flexible - the script will find columns containing "north star" and "key elements" (case-insensitive). If not found, it uses the first two columns.

### 2. Roadmap Sheet

This sheet should have three columns:

| Timeline | Phase | Workpackage |
|----------|-------|-------------|
| Phase 1 | Foundation | Build core infrastructure |
| Phase 1 | Foundation | Set up development environment |
| Phase 2 | Transformation | Migrate legacy systems |
| Phase 2 | Growth | Launch new features |

- **Timeline column**: Groups roadmap items (e.g., "Phase 1", "Phase 2", "Q1 2024")
- **Phase column**: Sub-categorizes within each timeline (e.g., "Foundation", "Transformation")
- **Workpackage column**: Description of work to be done

**Note**: Column names are flexible - the script searches for "timeline", "phase", and "workpackage" (case-insensitive).

### 3. Workpackages Sheet (Optional)

This sheet is currently skipped by the script but can be added for future use.

## Usage

### If Installed as UV Tool

After installing with `uv tool install .`, you can run:

```bash
roadmap-ppt your_roadmap.xlsx
roadmap-ppt your_roadmap.xlsx -o my_presentation.pptx
```

### If Running Directly

```bash
python main.py your_roadmap.xlsx
python main.py your_roadmap.xlsx -o my_presentation.pptx
```

### Command-Line Options

- `excel_file` (required): Path to your Excel file
- `-o, --output` (optional): Custom output file path for the PowerPoint presentation

Both methods will create `your_roadmap.pptx` in the same directory as your Excel file (unless `-o` is specified).

## Customization

All branding and styling options are located at the top of `main.py` in the **CONFIGURATION SECTION**. Edit these values to match your corporate branding:

### Colors

```python
BRAND_PRIMARY_COLOR = RGBColor(0, 51, 102)      # Main brand color
BRAND_SECONDARY_COLOR = RGBColor(0, 102, 204)   # Secondary color
BRAND_ACCENT_COLOR = RGBColor(255, 153, 0)      # Accent color
BRAND_TEXT_COLOR = RGBColor(51, 51, 51)         # Text color
BRAND_BACKGROUND_COLOR = RGBColor(255, 255, 255) # Background color
```

**RGBColor values**: Use RGB values from 0-255. For example:
- `RGBColor(255, 0, 0)` = Red
- `RGBColor(0, 255, 0)` = Green
- `RGBColor(0, 0, 255)` = Blue

### Logo

```python
LOGO_PATH = "logo.png"  # Path to your logo file (or None to skip)
LOGO_POSITION = "top_right"  # Options: "top_left", "top_right", "bottom_left", "bottom_right", "center"
```

Place your logo file in the same directory as the script, or provide a relative/absolute path.

### Fonts

```python
TITLE_FONT_NAME = "Calibri"  # Font for titles
BODY_FONT_NAME = "Calibri"   # Font for body text
TITLE_FONT_SIZE = Pt(44)     # Title font size
BODY_FONT_SIZE = Pt(18)      # Body text font size
```

**Note**: Fonts must be installed on your system. Common options:
- Windows: Calibri, Arial, Times New Roman, Segoe UI
- Mac: Helvetica, Arial, Times
- Linux: DejaVu Sans, Liberation Sans

### Slide Layout

```python
SLIDE_WIDTH = Inches(10)   # Slide width
SLIDE_HEIGHT = Inches(7.5) # Slide height
```

Standard PowerPoint sizes:
- Standard (4:3): 10" x 7.5"
- Widescreen (16:9): 10" x 5.625"
- Custom: Adjust as needed

### Visual Style

```python
USE_SHAPES = True  # Use rounded rectangles for content boxes
CONTENT_BOX_COLOR = RGBColor(245, 245, 245)  # Color for content boxes
```

Set `USE_SHAPES = False` for a simpler text-only layout.

## Output Structure

The generated PowerPoint presentation includes:

1. **Title Slide**
   - Presentation title
   - North star objective as subtitle
   - Logo (if configured)

2. **Objectives Slide**
   - North star prominently displayed
   - Key elements listed as bullet points

3. **Roadmap Slides** (one per timeline)
   - Timeline as slide title
   - Content organized by phase (if phases are used)
   - Workpackages displayed as bullet points in organized boxes

## Examples

### Example 1: Basic Usage

```bash
python main.py roadmap.xlsx
```

Creates `roadmap.pptx` with default branding.

### Example 2: Custom Output Location

```bash
python main.py C:\Documents\roadmap.xlsx -o C:\Presentations\Q1_Roadmap.pptx
```

### Example 3: Using Different Excel File

```bash
python main.py my_company_roadmap_2024.xlsx -o company_presentation.pptx
```

## Troubleshooting

### Error: "Excel file not found"
- Check that the file path is correct
- Use absolute paths if relative paths don't work
- Ensure the file has a `.xlsx` extension

### Error: "Error reading Objectives sheet"
- Verify your Excel file has a sheet named "Objectives" (case-sensitive)
- Check that the sheet has at least one column
- Ensure the Excel file is not corrupted or password-protected

### Error: "Error reading Roadmap sheet"
- Verify your Excel file has a sheet named "Roadmap" (case-sensitive)
- Check that the sheet has at least one column
- Ensure data starts from the first row (no empty header rows)

### Logo Not Appearing
- Verify the logo file path is correct
- Check that the logo file exists at the specified location
- Supported formats: PNG, JPG, JPEG
- Ensure the file is not corrupted

### Fonts Not Applying
- Verify the font name matches exactly (case-sensitive)
- Check that the font is installed on your system
- Use system font names (e.g., "Arial" not "arial.ttf")

### Slides Look Empty
- Check that your Excel sheets contain data
- Verify column names match expected patterns
- Ensure cells are not empty where data is expected

## File Structure

```
roadmap/
├── main.py              # Main script
├── requirements.txt      # Python dependencies
├── README.md           # This file
└── your_roadmap.xlsx   # Your Excel file (example)
```

## License

This script is provided as-is for your use. Feel free to modify and customize as needed.

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Verify your Excel file format matches the requirements
3. Review the configuration section in `main.py` for customization options


"""
Core PowerPoint generation functions.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import pandas as pd

from . import config


def hex_to_rgb(hex_color):
    """Convert hex color string to RGBColor object."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def read_objectives(excel_file):
    """
    Read Objectives sheet from Excel file.
    Returns: DataFrame with 'North star' and 'Key elements' columns
    """
    try:
        df = pd.read_excel(excel_file, sheet_name='Objectives')
        # Handle case-insensitive column names
        df.columns = df.columns.str.strip()
        # Try to find columns (case-insensitive)
        north_star_col = None
        key_elements_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if 'north' in col_lower and 'star' in col_lower:
                north_star_col = col
            elif 'key' in col_lower and 'element' in col_lower:
                key_elements_col = col
        
        if north_star_col is None:
            # Assume first column is North star
            north_star_col = df.columns[0]
        if key_elements_col is None:
            # Assume second column is Key elements
            key_elements_col = df.columns[1] if len(df.columns) > 1 else None
        
        # Filter out empty rows
        df = df.dropna(subset=[north_star_col])
        
        return {
            'north_star': df[north_star_col].iloc[0] if len(df) > 0 else "",
            'key_elements': df[key_elements_col].dropna().tolist() if key_elements_col else []
        }
    except Exception as e:
        print(f"Error reading Objectives sheet: {e}")
        return {'north_star': '', 'key_elements': []}


def read_roadmap(excel_file):
    """
    Read Roadmap sheet from Excel file.
    Returns: DataFrame with Timeline, Phase, and Workpackage columns
    """
    try:
        df = pd.read_excel(excel_file, sheet_name='Roadmap')
        df.columns = df.columns.str.strip()
        
        # Try to identify columns (case-insensitive)
        timeline_col = None
        phase_col = None
        workpackage_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if 'timeline' in col_lower or 'phase' in col_lower and timeline_col is None:
                timeline_col = col
            elif 'phase' in col_lower and phase_col is None:
                phase_col = col
            elif 'workpackage' in col_lower or 'work package' in col_lower:
                workpackage_col = col
        
        # Fallback to positional columns
        if timeline_col is None:
            timeline_col = df.columns[0]
        if phase_col is None:
            phase_col = df.columns[1] if len(df.columns) > 1 else None
        if workpackage_col is None:
            workpackage_col = df.columns[2] if len(df.columns) > 2 else None
        
        # Filter out empty rows
        df = df.dropna(subset=[timeline_col])
        
        result_df = pd.DataFrame({
            'Timeline': df[timeline_col],
            'Phase': df[phase_col] if phase_col else [''] * len(df),
            'Workpackage': df[workpackage_col] if workpackage_col else [''] * len(df)
        })
        
        return result_df
    except Exception as e:
        print(f"Error reading Roadmap sheet: {e}")
        return pd.DataFrame(columns=['Timeline', 'Phase', 'Workpackage'])


def add_logo(slide, logo_path, position):
    """Add logo to slide at specified position."""
    if logo_path is None or not os.path.exists(logo_path):
        return
    
    try:
        left_map = {
            'top_left': Inches(0.5),
            'top_right': config.SLIDE_WIDTH - Inches(1.5),
            'bottom_left': Inches(0.5),
            'bottom_right': config.SLIDE_WIDTH - Inches(1.5),
            'center': (config.SLIDE_WIDTH - Inches(1)) / 2
        }
        top_map = {
            'top_left': Inches(0.3),
            'top_right': Inches(0.3),
            'bottom_left': config.SLIDE_HEIGHT - Inches(0.8),
            'bottom_right': config.SLIDE_HEIGHT - Inches(0.8),
            'center': Inches(0.3)
        }
        
        left = left_map.get(position, Inches(0.5))
        top = top_map.get(position, Inches(0.3))
        
        slide.shapes.add_picture(logo_path, left, top, height=Inches(0.7))
    except Exception as e:
        print(f"Warning: Could not add logo: {e}")


def calculate_text_height(text, width, font_size, min_height=None, max_height=None):
    """
    Calculate estimated height needed for text based on content length and width.
    
    Args:
        text: Text content
        width: Available width in inches
        font_size: Font size in points
        min_height: Minimum height in inches (optional)
        max_height: Maximum height in inches (optional)
    
    Returns:
        Estimated height in inches
    """
    if not text:
        return min_height or Inches(0.5)
    
    # Estimate characters per line based on width and font size
    # Rough estimate: 1 inch ≈ 12 characters at 18pt font
    chars_per_inch = max(8, 12 * (18 / font_size.pt))
    chars_per_line = int(width / Inches(1) * chars_per_inch)
    
    # Calculate estimated lines
    text_length = len(str(text))
    estimated_lines = max(1, (text_length // chars_per_line) + 1)
    
    # Add some buffer for word wrapping
    estimated_lines = int(estimated_lines * 1.2)
    
    # Calculate height (roughly 1.2x font size per line)
    line_height = font_size.pt * 1.2 / 72  # Convert points to inches
    estimated_height = Inches(estimated_lines * line_height)
    
    # Apply min/max constraints
    if min_height and estimated_height < min_height:
        estimated_height = min_height
    if max_height and estimated_height > max_height:
        estimated_height = max_height
    
    return estimated_height


def create_title_slide(prs, objectives_data):
    """Create title slide with branding."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Set background color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = config.BRAND_BACKGROUND_COLOR
    
    # Add logo if configured
    if config.LOGO_PATH:
        add_logo(slide, config.LOGO_PATH, config.LOGO_POSITION)
    
    # Add title
    title_box = slide.shapes.add_textbox(
        config.SIDE_MARGIN,
        config.TITLE_TOP_MARGIN,
        config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
        Inches(2)
    )
    title_frame = title_box.text_frame
    title_frame.text = "Roadmap Presentation"
    title_para = title_frame.paragraphs[0]
    title_para.font.name = config.TITLE_FONT_NAME
    title_para.font.size = config.TITLE_FONT_SIZE
    title_para.font.bold = True
    title_para.font.color.rgb = config.BRAND_PRIMARY_COLOR
    title_para.alignment = PP_ALIGN.CENTER
    
    # Add subtitle (North star)
    if objectives_data.get('north_star'):
        subtitle_box = slide.shapes.add_textbox(
            config.SIDE_MARGIN,
            config.TITLE_TOP_MARGIN + Inches(2.5),
            config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
            config.SLIDE_HEIGHT - config.TITLE_TOP_MARGIN - Inches(2.5) - config.BOTTOM_MARGIN
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = objectives_data['north_star']
        subtitle_frame.word_wrap = True
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.font.name = config.BODY_FONT_NAME
        subtitle_para.font.size = config.SUBTITLE_FONT_SIZE
        subtitle_para.font.color.rgb = config.BRAND_SECONDARY_COLOR
        subtitle_para.alignment = PP_ALIGN.CENTER


def create_objectives_slide(prs, objectives_data):
    """Create objectives slide(s) displaying North star and key elements with pagination."""
    key_elements = objectives_data.get('key_elements', [])
    has_north_star = bool(objectives_data.get('north_star'))
    
    # Calculate available space for key elements
    # Account for: slide title, north star section (if present), key elements title
    SLIDE_TITLE_HEIGHT = Inches(0.8)
    SLIDE_TITLE_TOP = Inches(0.5)
    KEY_ELEMENTS_TITLE_HEIGHT = Inches(0.6)
    KEY_ELEMENTS_TITLE_SPACING = Inches(0.8)
    
    # Calculate space after North Star (if present)
    if has_north_star:
        north_star_text = str(objectives_data['north_star'])
        box_width = config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN) - (2 * config.TEXT_BOX_MARGIN)
        north_star_height = calculate_text_height(
            north_star_text,
            box_width,
            config.BODY_FONT_SIZE,
            min_height=config.NORTH_STAR_MIN_HEIGHT,
            max_height=config.NORTH_STAR_MAX_HEIGHT
        )
        space_after_north_star = Inches(1.2) + north_star_height + Inches(0.3)  # Header + content + spacing
    else:
        space_after_north_star = 0
    
    # Calculate available height for key elements
    available_height = (
        config.SLIDE_HEIGHT 
        - SLIDE_TITLE_TOP 
        - SLIDE_TITLE_HEIGHT 
        - space_after_north_star
        - KEY_ELEMENTS_TITLE_HEIGHT 
        - KEY_ELEMENTS_TITLE_SPACING
        - config.BOTTOM_MARGIN
    )
    
    # Calculate how many key elements fit per slide
    items_per_slide = max(1, int(available_height / config.KEY_ELEMENT_HEIGHT_ESTIMATE))
    total_elements = len(key_elements)
    slides_needed = max(1, (total_elements + items_per_slide - 1) // items_per_slide) if total_elements > 0 else 1
    
    # Create slides
    for slide_num in range(slides_needed):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Set background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = config.BRAND_BACKGROUND_COLOR
        
        # Add logo
        if config.LOGO_PATH:
            add_logo(slide, config.LOGO_PATH, config.LOGO_POSITION)
        
        # Slide title - add page number if multiple slides
        title_text = "Objectives"
        if slides_needed > 1:
            title_text += f" (Page {slide_num + 1} of {slides_needed})"
        
        title_box = slide.shapes.add_textbox(
            config.SIDE_MARGIN,
            SLIDE_TITLE_TOP,
            config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
            SLIDE_TITLE_HEIGHT
        )
        title_frame = title_box.text_frame
        title_frame.text = title_text
        title_para = title_frame.paragraphs[0]
        title_para.font.name = config.TITLE_FONT_NAME
        title_para.font.size = config.HEADING_FONT_SIZE
        title_para.font.bold = True
        title_para.font.color.rgb = config.BRAND_PRIMARY_COLOR
        
        y_pos = config.CONTENT_TOP_MARGIN
        
        # North Star section - only show on first slide
        if has_north_star and slide_num == 0:
            north_star_box = slide.shapes.add_textbox(
                config.SIDE_MARGIN,
                y_pos,
                config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
                Inches(1)
            )
            north_star_frame = north_star_box.text_frame
            north_star_frame.text = "North Star"
            north_star_para = north_star_frame.paragraphs[0]
            north_star_para.font.name = config.BODY_FONT_NAME
            north_star_para.font.size = Pt(24)
            north_star_para.font.bold = True
            north_star_para.font.color.rgb = config.BRAND_SECONDARY_COLOR
            
            y_pos += Inches(1.2)
            
            # Calculate dynamic height for North Star content box
            north_star_text = str(objectives_data['north_star'])
            box_width = config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN) - (2 * config.TEXT_BOX_MARGIN)
            dynamic_height = calculate_text_height(
                north_star_text,
                box_width,
                config.BODY_FONT_SIZE,
                min_height=config.NORTH_STAR_MIN_HEIGHT,
                max_height=config.NORTH_STAR_MAX_HEIGHT
            )
            
            # North star content box
            if config.USE_SHAPES:
                shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    config.SIDE_MARGIN,
                    y_pos,
                    config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
                    dynamic_height
                )
                shape.fill.solid()
                shape.fill.fore_color.rgb = config.CONTENT_BOX_COLOR
                shape.line.color.rgb = config.BRAND_SECONDARY_COLOR
                shape.line.width = Pt(2)
                
                text_frame = shape.text_frame
                text_frame.text = north_star_text
                text_frame.word_wrap = True
                text_frame.margin_left = config.TEXT_BOX_MARGIN
                text_frame.margin_right = config.TEXT_BOX_MARGIN
                text_frame.margin_top = Inches(0.1)
                text_frame.margin_bottom = Inches(0.1)
                
                para = text_frame.paragraphs[0]
                para.font.name = config.BODY_FONT_NAME
                para.font.size = config.BODY_FONT_SIZE
                para.font.color.rgb = config.BRAND_TEXT_COLOR
            else:
                north_star_content = slide.shapes.add_textbox(
                    config.SIDE_MARGIN,
                    y_pos,
                    config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
                    dynamic_height
                )
                north_star_content_frame = north_star_content.text_frame
                north_star_content_frame.text = north_star_text
                north_star_content_frame.word_wrap = True
                north_star_content_para = north_star_content_frame.paragraphs[0]
                north_star_content_para.font.name = config.BODY_FONT_NAME
                north_star_content_para.font.size = config.BODY_FONT_SIZE
                north_star_content_para.font.color.rgb = config.BRAND_TEXT_COLOR
            
            y_pos += dynamic_height + Inches(0.3)
        
        # Key Elements section
        if total_elements > 0:
            key_elements_title = slide.shapes.add_textbox(
                config.SIDE_MARGIN,
                y_pos,
                config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
                KEY_ELEMENTS_TITLE_HEIGHT
            )
            key_elements_title_frame = key_elements_title.text_frame
            key_elements_title_frame.text = "Key Elements"
            key_elements_title_para = key_elements_title_frame.paragraphs[0]
            key_elements_title_para.font.name = config.BODY_FONT_NAME
            key_elements_title_para.font.size = Pt(24)
            key_elements_title_para.font.bold = True
            key_elements_title_para.font.color.rgb = config.BRAND_SECONDARY_COLOR
            
            y_pos += KEY_ELEMENTS_TITLE_SPACING
            
            # Calculate which elements to show on this slide
            start_idx = slide_num * items_per_slide
            end_idx = min(start_idx + items_per_slide, total_elements)
            elements_for_slide = key_elements[start_idx:end_idx]
            
            # Calculate available height for elements list
            elements_height = config.SLIDE_HEIGHT - y_pos - config.BOTTOM_MARGIN
            
            # Key elements list
            elements_box = slide.shapes.add_textbox(
                config.SIDE_MARGIN + Inches(0.3),
                y_pos,
                config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN) - Inches(0.3),
                elements_height
            )
            elements_frame = elements_box.text_frame
            elements_frame.word_wrap = True
            
            for i, element in enumerate(elements_for_slide):
                if i > 0:
                    elements_frame.add_paragraph()
                para = elements_frame.paragraphs[i]
                para.text = f"• {str(element)}"
                para.font.name = config.BODY_FONT_NAME
                para.font.size = config.BODY_FONT_SIZE
                para.font.color.rgb = config.BRAND_TEXT_COLOR
                para.space_after = Pt(8)
                para.level = 0


def create_roadmap_slides(prs, roadmap_df):
    """Create roadmap slides grouped by timeline/phase with pagination support."""
    if roadmap_df.empty:
        return
    
    # Constants for pagination calculations
    SLIDE_TITLE_HEIGHT = Inches(0.8)
    SLIDE_TITLE_TOP = Inches(0.5)
    PHASE_HEADER_HEIGHT = Inches(0.7)
    ITEM_HEIGHT_ESTIMATE = Inches(0.5)  # Estimated height per workpackage item
    MIN_ITEM_HEIGHT = Inches(0.3)
    MAX_CONTENT_HEIGHT = config.SLIDE_HEIGHT - config.CONTENT_TOP_MARGIN - config.BOTTOM_MARGIN
    
    # Group by Timeline
    grouped = roadmap_df.groupby('Timeline')
    
    for timeline, group in grouped:
        # Group by Phase within this timeline
        phase_groups = list(group.groupby('Phase') if 'Phase' in group.columns else [(None, group)])
        num_phases = len(phase_groups)
        
        # Calculate layout dimensions
        max_width = (config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN)) / num_phases if num_phases > 1 else (config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN))
        box_width = max_width - Inches(0.2)
        
        # Organize workpackages by phase with pagination info
        phase_data_list = []
        for phase, phase_data in phase_groups:
            workpackages = phase_data['Workpackage'].dropna().tolist()
            phase_data_list.append({
                'phase': phase,
                'workpackages': workpackages,
                'total_items': len(workpackages)
            })
        
        # Calculate how many items can fit per slide
        available_height = MAX_CONTENT_HEIGHT - PHASE_HEADER_HEIGHT
        items_per_slide = max(1, int(available_height / ITEM_HEIGHT_ESTIMATE))
        
        # Determine if we need multiple slides
        max_items_any_phase = max([pd['total_items'] for pd in phase_data_list], default=0)
        slides_needed = max(1, (max_items_any_phase + items_per_slide - 1) // items_per_slide)
        
        # Create slides for this timeline
        for slide_num in range(slides_needed):
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
            
            # Set background
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = config.BRAND_BACKGROUND_COLOR
            
            # Add logo
            if config.LOGO_PATH:
                add_logo(slide, config.LOGO_PATH, config.LOGO_POSITION)
            
            # Slide title (Timeline) - add page number if multiple slides
            title_text = f"Roadmap: {str(timeline)}"
            if slides_needed > 1:
                title_text += f" (Page {slide_num + 1} of {slides_needed})"
            
            title_box = slide.shapes.add_textbox(
                config.SIDE_MARGIN,
                SLIDE_TITLE_TOP,
                config.SLIDE_WIDTH - (2 * config.SIDE_MARGIN),
                SLIDE_TITLE_HEIGHT
            )
            title_frame = title_box.text_frame
            title_frame.text = title_text
            title_para = title_frame.paragraphs[0]
            title_para.font.name = config.TITLE_FONT_NAME
            title_para.font.size = config.HEADING_FONT_SIZE
            title_para.font.bold = True
            title_para.font.color.rgb = config.BRAND_PRIMARY_COLOR
            
            y_pos = config.CONTENT_TOP_MARGIN
            
            # Add content for each phase
            for phase_idx, phase_info in enumerate(phase_data_list):
                phase = phase_info['phase']
                workpackages = phase_info['workpackages']
                
                # Calculate which items to show on this slide
                start_idx = slide_num * items_per_slide
                end_idx = min(start_idx + items_per_slide, len(workpackages))
                
                items_for_this_slide = workpackages[start_idx:end_idx] if start_idx < len(workpackages) else []
                
                x_pos = config.SIDE_MARGIN + (phase_idx * max_width) + Inches(0.1)
                
                # Phase header - always show if phase name exists
                if phase and str(phase).strip():
                    phase_title = slide.shapes.add_textbox(
                        x_pos,
                        y_pos,
                        box_width,
                        Inches(0.6)
                    )
                    phase_title_frame = phase_title.text_frame
                    phase_title_frame.text = str(phase)
                    phase_title_frame.word_wrap = True  # Ensure phase titles wrap if too long
                    phase_title_para = phase_title_frame.paragraphs[0]
                    phase_title_para.font.name = config.BODY_FONT_NAME
                    phase_title_para.font.size = Pt(22)
                    phase_title_para.font.bold = True
                    phase_title_para.font.color.rgb = config.BRAND_SECONDARY_COLOR
                    y_pos_phase = y_pos + PHASE_HEADER_HEIGHT
                else:
                    y_pos_phase = y_pos
                
                # Only show content box if there are items for this slide
                if items_for_this_slide:
                    # Calculate content height based on number of items
                    content_height = min(
                        MAX_CONTENT_HEIGHT - (y_pos_phase - config.CONTENT_TOP_MARGIN),
                        max(MIN_ITEM_HEIGHT, len(items_for_this_slide) * ITEM_HEIGHT_ESTIMATE + Inches(0.3))
                    )
                    if config.USE_SHAPES:
                        shape = slide.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE,
                            x_pos,
                            y_pos_phase,
                            box_width,
                            content_height
                        )
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = config.CONTENT_BOX_COLOR
                        shape.line.color.rgb = config.BRAND_ACCENT_COLOR
                        shape.line.width = Pt(1.5)
                        
                        text_frame = shape.text_frame
                        text_frame.word_wrap = True
                        text_frame.margin_left = Inches(0.15)
                        text_frame.margin_right = Inches(0.15)
                        text_frame.margin_top = Inches(0.15)
                        text_frame.margin_bottom = Inches(0.15)
                        
                        for i, wp in enumerate(items_for_this_slide):
                            if i > 0:
                                text_frame.add_paragraph()
                            para = text_frame.paragraphs[i]
                            para.text = f"• {str(wp)}"
                            para.font.name = config.BODY_FONT_NAME
                            para.font.size = Pt(14)
                            para.font.color.rgb = config.BRAND_TEXT_COLOR
                            para.space_after = Pt(6)
                    else:
                        wp_box = slide.shapes.add_textbox(
                            x_pos,
                            y_pos_phase,
                            box_width,
                            content_height
                        )
                        wp_frame = wp_box.text_frame
                        wp_frame.word_wrap = True
                        
                        for i, wp in enumerate(items_for_this_slide):
                            if i > 0:
                                wp_frame.add_paragraph()
                            para = wp_frame.paragraphs[i]
                            para.text = f"• {str(wp)}"
                            para.font.name = config.BODY_FONT_NAME
                            para.font.size = Pt(14)
                            para.font.color.rgb = config.BRAND_TEXT_COLOR
                            para.space_after = Pt(6)


def generate_presentation(excel_file, output_path=None):
    """
    Generate PowerPoint presentation from Excel file.
    
    Args:
        excel_file: Path to Excel file
        output_path: Optional output path (defaults to same name as Excel with .pptx extension)
    
    Returns:
        Path to generated PowerPoint file
    """
    import os
    
    # Determine output file path
    if output_path is None:
        base_name = os.path.splitext(excel_file)[0]
        output_path = f"{base_name}.pptx"
    
    # Read Excel data
    objectives_data = read_objectives(excel_file)
    roadmap_df = read_roadmap(excel_file)
    
    print(f"Found {len(roadmap_df)} roadmap entries")
    print(f"North Star: {objectives_data.get('north_star', 'Not found')}")
    print(f"Key Elements: {len(objectives_data.get('key_elements', []))} items")
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = config.SLIDE_WIDTH
    prs.slide_height = config.SLIDE_HEIGHT
    
    # Generate slides
    print("Generating slides...")
    create_title_slide(prs, objectives_data)
    create_objectives_slide(prs, objectives_data)
    create_roadmap_slides(prs, roadmap_df)
    
    # Save presentation
    prs.save(output_path)
    print(f"PowerPoint presentation saved to: {output_path}")
    
    return output_path


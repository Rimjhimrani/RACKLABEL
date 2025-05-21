import subprocess
import sys

# Install pandas if not already installed
subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])

import pandas as pd
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import threading

# Style for bold part numbers - First version
bold_style_v1 = ParagraphStyle(
    name='Bold_v1',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=20,
    spaceBefore=2,
    spaceAfter=2
)

# Style for bold part numbers - IMPROVED ALIGNMENT - Second version
bold_style_v2 = ParagraphStyle(
    name='Bold_v2',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,  # Center alignment
    leading=12,  # Reduced leading for better spacing
    spaceBefore=0,  # No extra space before
    spaceAfter=15,  # Add space after to push text upward within the cell
)

# Style for wrapped descriptions - Second version
desc_style = ParagraphStyle(
    name='Description',
    fontName='Helvetica',
    fontSize=20,
    alignment=TA_LEFT,
    leading=16,
    spaceBefore=2,
    spaceAfter=2
)

class RedirectText:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")
        self.text_widget.update()

    def flush(self):
        pass

def format_part_no_v1(part_no):
    """Format part number with first 7 characters in 17pt font, rest in 22pt font."""
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)

    if len(part_no) > 5:
        split_point = len(part_no) - 5  # Calculate where to split based on total length
        part1 = part_no[:split_point]   # Everything except the last 5 characters
        part2 = part_no[-5:]            # Last 5 characters
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    else:
        # If part number is too short, just use one size
        return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    """Format part number with different font sizes to prevent overlapping."""
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)

    if len(part_no) > 5:
        split_point = len(part_no) - 5  # Calculate where to split based on total length
        part1 = part_no[:split_point]   # Everything except the last 5 characters
        part2 = part_no[-5:]
        # Add extra padding to ensure space between text and bottom line
        return Paragraph(f"<b><font size=34>{part1}</font><font size=40>{part2}</font></b><br/><br/>", bold_style_v2)
    else:
        # If part number is too short, just use one size
        return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description(desc):
    """Format description text with proper wrapping."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)

    # Prepare the description for proper wrapping in the PDF
    return Paragraph(desc, desc_style)

def parse_location_string_v1(location_str):
    """
    Parse a location string like "12M - LH -R-0-2-A-1" into its 7 components.
    Returns a list of 7 values.
    """
    # Initialize with empty values
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    # Remove any extra spaces
    location_str = location_str.strip()

    # Try to parse as "12M - LH -R-0-2-A-1" format
    # Define pattern: looking for parts separated by - or spaces
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    # Fill the available parts
    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    print(f"Parsed location '{location_str}' into: {location_parts}")
    return location_parts

def parse_location_string_v2(location_str):
    """
    Parse a location string like "12M_ST-140_R_0_2_A_1" into its 7 components.
    Returns a list of 7 values.
    """
    # Initialize with empty values
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    # Remove any extra spaces
    location_str = location_str.strip()

    # Try to parse as "12M - LH -R-0-2-A-1" format
    # Define pattern: looking for parts separated by - or spaces
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    # Fill the available parts
    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    print(f"Parsed location '{location_str}' into: {location_parts}")
    return location_parts

def generate_labels_from_excel_v1(excel_file_path, output_pdf_path):
    try:
        print(f"Attempting to read Excel file: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"Error: Excel file not found at {excel_file_path}")
            return None

        # Enhanced Excel reading with various engines to improve compatibility
        try:
            # Check if the file is CSV or Excel
            if excel_file_path.lower().endswith('.csv'):
                df = pd.read_csv(excel_file_path)
            else:
                df = pd.read_excel(excel_file_path)
        except Exception as first_error:
            try:
                print("First attempt failed, trying with engine='openpyxl'...")
                df = pd.read_excel(excel_file_path, engine='openpyxl')
            except Exception as second_error:
                try:
                    print("Second attempt failed, trying with engine='xlrd'...")
                    df = pd.read_excel(excel_file_path, engine='xlrd')
                except Exception as third_error:
                    # Final attempt: try csv with different encodings
                    try:
                        df = pd.read_csv(excel_file_path, encoding='utf-8')
                    except:
                        df = pd.read_csv(excel_file_path, encoding='latin1')

        print(f"Successfully read file with {len(df)} rows")
        print("Columns found:", df.columns.tolist())

        # Display first few rows to help with debugging
        print("\nFirst 2 rows of data:")
        print(df.head(2))

    except Exception as e:
        print(f"Error reading file: {e}")
        return None

    # Set up key measurements
    label_width = 15 * cm
    label_height = 5 * cm
    part_no_height = 1.3 * cm   # Height for part number rows
    desc_loc_height = 0.8 * cm  # Height for description and location rows

    # Identify column names in the file
    cols = df.columns.tolist()

    # Normalize column names (convert to uppercase)
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Standard column names to look for (case-insensitive)
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))

    desc_col = next((col for col in cols if 'DESC' in col), None)
    loc_col = next((col for col in cols if 'LOC' in col or 'POS' in col), None)

    if not part_no_col:
        print(f"Warning: Could not find part number column in {cols}")
        part_no_col = cols[0]  # Use first column as fallback

    if not desc_col:
        print(f"Warning: Could not find description column in {cols}")
        desc_col = cols[1] if len(cols) > 1 else part_no_col  # Use second column as fallback

    if not loc_col:
        print(f"Warning: Could not find location column in {cols}")
        loc_col = cols[2] if len(cols) > 2 else desc_col  # Use third column as fallback

    print(f"Using columns: Part No: {part_no_col}, Description: {desc_col}, Location: {loc_col}")

    # Group parts by location to create pairs
    df_grouped = df.groupby(loc_col)

    doc = SimpleDocTemplate(output_pdf_path, pagesize=A4)
    elements = []

    # Hard limit: maximum 4 labels per page (each label has 2 parts)
    MAX_LABELS_PER_PAGE = 4

    # Keep track of labels for pagination
    label_count = 0

    # Process records by location
    for location, group in df_grouped:
        try:
            # Get the first two parts for each location
            parts = group.head(2)

            if len(parts) < 2:
                # If only one part at this location, duplicate it (or keep as is)
                if len(parts) == 1:
                    print(f"Only one part found for location {location}. Proceeding with single part.")
                    part1 = parts.iloc[0]
                    part2 = parts.iloc[0]  # Use the same part twice if needed
                else:
                    print(f"No parts found for location {location}. Skipping.")
                    continue
            else:
                part1 = parts.iloc[0]
                part2 = parts.iloc[1]

            # Force a new page after every 4 labels
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            # Increment label counter
            label_count += 1

            # Extract details for both parts
            part_no_1 = str(part1[part_no_col])
            desc_1 = str(part1[desc_col])

            part_no_2 = str(part2[part_no_col])
            desc_2 = str(part2[desc_col])

            # Use location from the first part
            location_str = str(part1[loc_col])

            # Parse location string into components
            location_values = parse_location_string_v1(location_str)

            print(f"Creating label for location {location} with parts: {part_no_1} and {part_no_2}")

            # First part table
            part_table = Table(
                [['Part No', format_part_no_v1(part_no_1)],
                 ['Description', desc_1[:50]]],  # Limit length to prevent overflow
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),  # Part number label middle aligned
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),  # Part number value middle aligned
                ('VALIGN', (0, 1), (0, 1), 'TOP'),     # Description label top aligned
                ('VALIGN', (1, 1), (1, 1), 'TOP'),     # Description value top aligned
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),     # Font size for labels (left side)
                ('FONTSIZE', (1, 1), (1, 1), 16),      # Font size for description value
            ]))

            # Second part table (with different part number)
            part_table2 = Table(
                [['Part No', format_part_no_v1(part_no_2)],
                 ['Description', desc_2[:50]]],  # Limit length to prevent overflow
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table2.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),  # Part number label middle aligned
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),  # Part number value middle aligned
                ('VALIGN', (0, 1), (0, 1), 'TOP'),     # Description label top aligned
                ('VALIGN', (1, 1), (1, 1), 'TOP'),     # Description value top aligned
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),     # Font size for labels (left side)
                ('FONTSIZE', (1, 1), (1, 1), 16),      # Font size for description value
            ]))

            # Create location table with parsed location values
            location_data = [['Part Location'] + location_values]

            # Set the fixed width for the first column (e.g., "Part Location")
            first_col_width = 4 * cm
            location_widths = [first_col_width]

            # Define total remaining width for the rest of the columns
            remaining_width = 11 * cm  # total width available for remaining 7 columns

            # Define relative proportions for the remaining 7 columns
            # Example: more space for column like "ST-140"
            col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]  # Adjust as needed
            total_proportion = sum(col_proportions)

            # Calculate actual widths based on proportions and remaining width
            adjusted_widths = [w * remaining_width / total_proportion for w in col_proportions]
            location_widths.extend(adjusted_widths)

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=desc_loc_height
            )

            location_colors = [
                colors.HexColor('#E9967A'),  # Salmon
                colors.HexColor('#ADD8E6'),  # Light Blue
                colors.HexColor('#90EE90'),  # Light Green
                colors.HexColor('#FFD700'),  # Gold
                colors.HexColor('#ADD8E6'),  # Light Blue
                colors.HexColor('#E9967A'),  # Salmon
                colors.HexColor('#90EE90')   # Light Green
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),       # Part Location label top aligned
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),      # Location values top aligned
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),      # Font size for Part Location label (left side)
                ('FONTSIZE', (1, 0), (-1, -1), 14),    # Font size for location values
            ]

            for i, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (i+1, 0), (i+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(part_table2)
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

            # Add spacer between labels, but not if this is the last label on the page
            if (label_count % MAX_LABELS_PER_PAGE) < MAX_LABELS_PER_PAGE - 1 and label_count < len(df_grouped):
                elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            print(f"Error processing location {location}: {e}")
            import traceback
            traceback.print_exc()  # Print detailed stack trace for better debugging
            continue

    if elements:
        doc.build(elements)
        print(f"PDF generated successfully: {output_pdf_path}")
        return output_pdf_path
    else:
        print("No labels were generated. Check if the Excel file has the expected columns.")
        return None

def generate_labels_from_excel_v2(excel_file_path, output_pdf_path, status_callback=None, progress_callback=None):
    try:
        if status_callback:
            status_callback(f"Reading file: {excel_file_path}")
        
        if not os.path.exists(excel_file_path):
            if status_callback:
                status_callback(f"Error: File not found at {excel_file_path}")
            return None

        # Enhanced Excel reading with various engines to improve compatibility
        try:
            # Check if the file is CSV or Excel
            if excel_file_path.lower().endswith('.csv'):
                df = pd.read_csv(excel_file_path)
            else:
                df = pd.read_excel(excel_file_path)
        except Exception as first_error:
            try:
                if status_callback:
                    status_callback("First attempt failed, trying with engine='openpyxl'...")
                df = pd.read_excel(excel_file_path, engine='openpyxl')
            except Exception as second_error:
                try:
                    if status_callback:
                        status_callback("Second attempt failed, trying with engine='xlrd'...")
                    df = pd.read_excel(excel_file_path, engine='xlrd')
                except Exception as third_error:
                    # Final attempt: try csv with different encodings
                    try:
                        df = pd.read_csv(excel_file_path, encoding='utf-8')
                    except:
                        df = pd.read_csv(excel_file_path, encoding='latin1')

        if status_callback:
            status_callback(f"Successfully read file with {len(df)} rows")
            status_callback(f"Columns found: {df.columns.tolist()}")

    except Exception as e:
        if status_callback:
            status_callback(f"Error reading file: {e}")
        return None

    # Set up key measurements
    label_width = 15 * cm
    label_height = 5 * cm
    part_no_height = 1.9 * cm    # Increased height for part number rows
    desc_height = 2.1 * cm       # Height for description rows
    loc_height = 0.9 * cm        # Height for location rows

    # Identify column names in the file
    cols = df.columns.tolist()

    # Normalize column names (convert to uppercase)
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Standard column names to look for (case-insensitive)
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))

    desc_col = next((col for col in cols if 'DESC' in col), None)
    loc_col = next((col for col in cols if 'LOC' in col or 'POS' in col), None)

    if not part_no_col:
        if status_callback:
            status_callback(f"Warning: Could not find part number column in {cols}")
        part_no_col = cols[0]  # Use first column as fallback

    if not desc_col:
        if status_callback:
            status_callback(f"Warning: Could not find description column in {cols}")
        desc_col = cols[1] if len(cols) > 1 else part_no_col  # Use second column as fallback

    if not loc_col:
        if status_callback:
            status_callback(f"Warning: Could not find location column in {cols}")
        loc_col = cols[2] if len(cols) > 2 else desc_col  # Use third column as fallback

    if status_callback:
        status_callback(f"Using columns: Part No: {part_no_col}, Description: {desc_col}, Location: {loc_col}")

    # Group parts by location to create pairs
    df_grouped = df.groupby(loc_col)
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(output_pdf_path, pagesize=A4)
    elements = []

    # Hard limit: maximum 4 labels per page (each label has 2 parts)
    MAX_LABELS_PER_PAGE = 4

    # Keep track of labels for pagination
    label_count = 0
    
    # Process records by location
    for i, (location, group) in enumerate(df_grouped):
        try:
            # Update progress
            if progress_callback:
                progress_value = int((i / total_locations) * 100)
                progress_callback(progress_value)
            
            if status_callback:
                status_callback(f"Processing location {i+1}/{total_locations}: {location}")
                
            # Get the first two parts for each location
            parts = group.head(2)

            if len(parts) < 2:
                # If only one part at this location, duplicate it (or keep as is)
                if len(parts) == 1:
                    if status_callback:
                        status_callback(f"Only one part found for location {location}. Proceeding with single part.")
                    part1 = parts.iloc[0]
                    part2 = parts.iloc[0]  # Use the same part twice if needed
                else:
                    if status_callback:
                        status_callback(f"No parts found for location {location}. Skipping.")
                    continue
            else:
                part1 = parts.iloc[0]
                part2 = parts.iloc[1]

            # Force a new page after every 4 labels
            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            # Increment label counter
            label_count += 1

            # Extract details for both parts
            part_no = str(part1[part_no_col])
            desc = str(part1[desc_col])

            # Use location from the first part
            location_str = str(part1[loc_col])

            # Parse location string into components
            location_values = parse_location_string_v2(location_str)

            # First part table with formatted description for wrapping
            part_table = Table(
                [['Part No', format_part_no_v2(part_no)],
                 ['Description', format_description(desc)]],  # Using paragraph style for wrapping
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_height]  # Using the adjusted height
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),  # Center alignment for part number value
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),   # Left alignment for description
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),  # Part number label middle aligned
                ('VALIGN', (1, 0), (1, 0), 'TOP'),     # Changed to TOP alignment for part number value
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),     # Description label top aligned
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),     # Description value top aligned
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (1, 0), (1, 0), 10),    # Added top padding to move part number up
                ('BOTTOMPADDING', (1, 0), (1, 0), 5),  # Added bottom padding for part number
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),     # Font size for labels (left side)
                # Note: Font size for description is now controlled by the desc_style ParagraphStyle
            ]))

            # Create location table with parsed location values - ADJUSTED WIDTHS
            location_data = [['Part Location'] + location_values]

            # Adjust column widths to prevent overlapping - give more space to the second column (ST-140)
            # Total width: 11 cm to distribute among 7 columns
            location_widths = [4*cm]  # First column width (Part Location)

            # Distribute remaining width with more space for column with ST-140
            remaining_width = 11 * cm
            col_widths = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]  # Relative width proportions
            total_proportion = sum(col_widths)

            # Calculate actual widths based on proportions
            location_widths.extend([w * remaining_width / total_proportion for w in col_widths])

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=loc_height,
            )

            location_colors = [
                colors.HexColor('#E9967A'),  # Salmon
                colors.HexColor('#ADD8E6'),  # Light Blue
                colors.HexColor('#90EE90'),  # Light Green
                colors.HexColor('#FFD700'),  # Gold
                colors.HexColor('#ADD8E6'),  # Light Blue
                colors.HexColor('#E9967A'),  # Salmon
                colors.HexColor('#90EE90')   # Light Green
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),       # Part Location label top aligned
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),      # Location values top aligned
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),      # Font size for Part Location label (left side)
                ('FONTSIZE', (1, 0), (-1, -1), 16),    # Font size for location values
            ]

            for i, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (i+1, 0), (i+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

            # Add spacer between labels, but not if this is the last label on the page
            if (label_count % MAX_LABELS_PER_PAGE) < MAX_LABELS_PER_PAGE - 1 and label_count < len(df_grouped):
                elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_callback:
                status_callback(f"Error processing location {location}: {e}")
            import traceback
            traceback.print_exc()  # Print detailed stack trace for better debugging
            continue

    # Set progress to 100% when done
    if progress_callback:
        progress_callback(100)

    if elements:
        if status_callback:
            status_callback(f"Building PDF document...")
        doc.build(elements)
        if status_callback:
            status_callback(f"PDF generated successfully: {output_pdf_path}")
        return output_pdf_path
    else:
        if status_callback:
            status_callback("No labels were generated. Check if the Excel file has the expected columns.")
        return None

class CombinedLabelGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Combined Part Label Generator")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create the two tab frames
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="Enhanced Labels")
        self.notebook.add(self.tab2, text="Standard Labels")
        
        # Create GUI elements for each tab
        self.create_widgets_tab1()  # Enhanced style (Version 2)
        self.create_widgets_tab2()  # Standard style (Version 1)
    
    def create_widgets_tab1(self):
        """Create widgets for Enhanced Layout (Version 2)"""
        # Configure the grid layout
        self.tab1.columnconfigure(0, weight=1)
        self.tab1.rowconfigure(0, weight=0)  # Header
        self.tab1.rowconfigure(1, weight=1)  # Main content
        self.tab1.rowconfigure(2, weight=0)  # Buttons
        
        # Header frame
        header_frame = ttk.Frame(self.tab1)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ttk.Label(header_frame, text="Generate Enhanced Part Labels", font=("Helvetica", 14)).pack()
        
        # Main content frame
        content_frame = ttk.Frame(self.tab1)
        content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        content_frame.columnconfigure(0, weight=0)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(2, weight=1)  # Make the log area expandable
        
        # File selector
        ttk.Label(content_frame, text="Excel File:").grid(row=0, column=0, sticky="w", pady=5)
        self.file_path_var1 = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self.file_path_var1, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(content_frame, text="Browse...", command=self.browse_file_tab1).grid(row=0, column=2, padx=5)
        
        # Output file selector
        ttk.Label(content_frame, text="Output PDF:").grid(row=1, column=0, sticky="w", pady=5)
        self.output_path_var1 = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self.output_path_var1, width=50).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(content_frame, text="Browse...", command=self.browse_output_tab1).grid(row=1, column=2, padx=5)
        
        # Progress bar
        ttk.Label(content_frame, text="Progress:").grid(row=2, column=0, sticky="w", pady=5)
        self.progress_var1 = tk.IntVar()
        self.progress_bar1 = ttk.Progressbar(content_frame, variable=self.progress_var1, maximum=100)
        self.progress_bar1.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        # Log area
        ttk.Label(content_frame, text="Log:").grid(row=3, column=0, sticky="nw", pady=5)
        
        self.log_frame1 = ttk.Frame(content_frame)
        self.log_frame1.grid(row=3, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.log_frame1.columnconfigure(0, weight=1)
        self.log_frame1.rowconfigure(0, weight=1)
        
        self.log_text1 = scrolledtext.ScrolledText(self.log_frame1, height=12, width=70)
        self.log_text1.grid(row=0, column=0, sticky="nsew")
        self.log_text1.config(state="disabled")
        
        # Button frame
        button_frame1 = ttk.Frame(self.tab1)
        button_frame1.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        button_frame1.columnconfigure(1, weight=1)
        
        ttk.Button(button_frame1, text="Generate PDF", command=self.generate_pdf_tab1).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame1, text="Clear", command=self.clear_form_tab1).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame1, text="Exit", command=self.root.quit).grid(row=0, column=2, padx=(5, 0))
        
        # Set up stdout redirection for logging
        self.redirect1 = RedirectText(self.log_text1)

    def create_widgets_tab2(self):
        """Create widgets for Standard Layout (Version 1)"""
        # Configure the grid layout
        self.tab2.columnconfigure(0, weight=1)
        self.tab2.rowconfigure(0, weight=0)  # Header
        self.tab2.rowconfigure(1, weight=1)  # Main content
        self.tab2.rowconfigure(2, weight=0)  # Buttons
        
        # Header frame
        header_frame = ttk.Frame(self.tab2)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ttk.Label(header_frame, text="Generate Standard Part Labels", font=("Helvetica", 14)).pack()
        
        # Main content frame
        content_frame = ttk.Frame(self.tab2)
        content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        content_frame.columnconfigure(0, weight=0)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(2, weight=1)  # Make the log area expandable
        
        # File selector
        ttk.Label(content_frame, text="Excel File:").grid(row=0, column=0, sticky="w", pady=5)
        self.file_path_var2 = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self.file_path_var2, width=50).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(content_frame, text="Browse...", command=self.browse_file_tab2).grid(row=0, column=2, padx=5)
        
        # Output file selector
        ttk.Label(content_frame, text="Output PDF:").grid(row=1, column=0, sticky="w", pady=5)
        self.output_path_var2 = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self.output_path_var2, width=50).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(content_frame, text="Browse...", command=self.browse_output_tab2).grid(row=1, column=2, padx=5)
        
        # Log area
        ttk.Label(content_frame, text="Log:").grid(row=2, column=0, sticky="nw", pady=5)
        
        self.log_frame2 = ttk.Frame(content_frame)
        self.log_frame2.grid(row=2, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.log_frame2.columnconfigure(0, weight=1)
        self.log_frame2.rowconfigure(0, weight=1)
        
        self.log_text2 = scrolledtext.ScrolledText(self.log_frame2, height=14, width=70)
        self.log_text2.grid(row=0, column=0, sticky="nsew")
        self.log_text2.config(state="disabled")
        
        # Button frame
        button_frame2 = ttk.Frame(self.tab2)
        button_frame2.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        button_frame2.columnconfigure(1, weight=1)
        
        ttk.Button(button_frame2, text="Generate PDF", command=self.generate_pdf_tab2).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame2, text="Clear", command=self.clear_form_tab2).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame2, text="Exit", command=self.root.quit).grid(row=0, column=2, padx=(5, 0))
        
        # Set up stdout redirection for logging
        self.redirect2 = RedirectText(self.log_text2)

    def browse_file_tab1(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls *.csv"),
            ("All files", "*.*")
        ])
        if file_path:
            self.file_path_var1.set(file_path)
            # Auto-generate output path with same filename but .pdf extension
            output_path = os.path.splitext(file_path)[0] + "_enhanced.pdf"
            self.output_path_var1.set(output_path)

    def browse_output_tab1(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if output_path:
            self.output_path_var1.set(output_path)

    def browse_file_tab2(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls *.csv"),
            ("All files", "*.*")
        ])
        if file_path:
            self.file_path_var2.set(file_path)
            # Auto-generate output path with same filename but .pdf extension
            output_path = os.path.splitext(file_path)[0] + "_standard.pdf"
            self.output_path_var2.set(output_path)

    def browse_output_tab2(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if output_path:
            self.output_path_var2.set(output_path)

    def generate_pdf_tab1(self):
        # Enhanced layout (Version 2)
        file_path = self.file_path_var1.get()
        output_path = self.output_path_var1.get()
        
        if not file_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output files")
            return
        
        # Clear log and reset progress
        self.log_text1.config(state="normal")
        self.log_text1.delete(1.0, tk.END)
        self.log_text1.config(state="disabled")
        self.progress_var1.set(0)
        
        # Redirect stdout to our log widget
        old_stdout = sys.stdout
        sys.stdout = self.redirect1
        
        try:
            # Run the PDF generation in a separate thread to keep UI responsive
            def run_generation():
                try:
                    # Call version 2 of the generator with status and progress callbacks
                    result = generate_labels_from_excel_v2(
                        file_path, 
                        output_path,
                        status_callback=self.update_status_tab1,
                        progress_callback=self.update_progress_tab1
                    )
                    
                    # Show result in UI thread
                    self.root.after(0, lambda: self.show_result_tab1(result))
                except Exception as e:
                    self.root.after(0, lambda: self.show_error_tab1(str(e)))
                finally:
                    # Restore stdout
                    sys.stdout = old_stdout
            
            # Start the thread
            threading.Thread(target=run_generation, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            sys.stdout = old_stdout

    def generate_pdf_tab2(self):
        # Standard layout (Version 1)
        file_path = self.file_path_var2.get()
        output_path = self.output_path_var2.get()
        
        if not file_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output files")
            return
        
        # Clear log
        self.log_text2.config(state="normal")
        self.log_text2.delete(1.0, tk.END)
        self.log_text2.config(state="disabled")
        
        # Redirect stdout to our log widget
        old_stdout = sys.stdout
        sys.stdout = self.redirect2
        
        try:
            # Run the PDF generation in a separate thread to keep UI responsive
            def run_generation():
                try:
                    # Call version 1 of the generator
                    result = generate_labels_from_excel_v1(file_path, output_path)
                    
                    # Show result in UI thread
                    self.root.after(0, lambda: self.show_result_tab2(result))
                except Exception as e:
                    self.root.after(0, lambda: self.show_error_tab2(str(e)))
                finally:
                    # Restore stdout
                    sys.stdout = old_stdout
            
            # Start the thread
            threading.Thread(target=run_generation, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            sys.stdout = old_stdout

    def update_status_tab1(self, message):
        """Update status in log text widget for tab 1"""
        self.log_text1.config(state="normal")
        self.log_text1.insert(tk.END, message + "\n")
        self.log_text1.see(tk.END)
        self.log_text1.config(state="disabled")
        self.root.update_idletasks()

    def update_progress_tab1(self, value):
        """Update progress bar for tab 1"""
        self.progress_var1.set(value)
        self.root.update_idletasks()

    def show_result_tab1(self, result):
        """Show final result for tab 1"""
        if result:
            messagebox.showinfo("Success", f"PDF file has been generated:\n{result}")
            # Open the generated PDF if desired
            # os.startfile(result) # Windows-specific
            # For cross-platform:
            try:
                if sys.platform == 'win32':
                    os.startfile(result)
                elif sys.platform == 'darwin':  # macOS
                    import subprocess
                    subprocess.call(('open', result))
                else:  # Linux
                    import subprocess
                    subprocess.call(('xdg-open', result))
            except Exception as e:
                print(f"Could not open PDF automatically: {e}")
        else:
            messagebox.showerror("Error", "Failed to generate PDF file. Check the log for details.")

    def show_error_tab1(self, error_message):
        """Show error for tab 1"""
        messagebox.showerror("Error", f"An error occurred: {error_message}")

    def show_result_tab2(self, result):
        """Show final result for tab 2"""
        if result:
            messagebox.showinfo("Success", f"PDF file has been generated:\n{result}")
            # Open the generated PDF if desired
            try:
                if sys.platform == 'win32':
                    os.startfile(result)
                elif sys.platform == 'darwin':  # macOS
                    import subprocess
                    subprocess.call(('open', result))
                else:  # Linux
                    import subprocess
                    subprocess.call(('xdg-open', result))
            except Exception as e:
                print(f"Could not open PDF automatically: {e}")
        else:
            messagebox.showerror("Error", "Failed to generate PDF file. Check the log for details.")

    def show_error_tab2(self, error_message):
        """Show error for tab 2"""
        messagebox.showerror("Error", f"An error occurred: {error_message}")

    def clear_form_tab1(self):
        """Clear all form fields in tab 1"""
        self.file_path_var1.set("")
        self.output_path_var1.set("")
        self.progress_var1.set(0)
        self.log_text1.config(state="normal")
        self.log_text1.delete(1.0, tk.END)
        self.log_text1.config(state="disabled")

    def clear_form_tab2(self):
        """Clear all form fields in tab 2"""
        self.file_path_var2.set("")
        self.output_path_var2.set("")
        self.log_text2.config(state="normal")
        self.log_text2.delete(1.0, tk.END)
        self.log_text2.config(state="disabled")


# Main execution block
if __name__ == "__main__":
    # Set up the main application window
    root = tk.Tk()
    app = CombinedLabelGeneratorApp(root)
    
    # Add window icon if available
    try:
        # Check if running as script or frozen exe
        if getattr(sys, 'frozen', False):
            # If frozen, use sys._MEIPASS for resources
            app_path = sys._MEIPASS
        else:
            # If running as script, use the script's directory
            app_path = os.path.dirname(os.path.abspath(__file__))
            
        icon_path = os.path.join(app_path, "label_icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass  # Skip icon if not available
    
    # Start the application
    root.mainloop()

    
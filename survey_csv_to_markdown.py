#!/usr/bin/env python3
"""
Survey CSV to Markdown Converter - Question-Aware

This script reads a CSV file containing survey data and converts it to Markdown
while preserving the question-response table structure.


"""

import pandas as pd
import os
import re
from pathlib import Path


def sanitize_filename(filename):
    """
    Sanitize filename by removing/replacing invalid characters for Windows filesystem
    """
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    sanitized = sanitized.strip('. ')
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    return sanitized


def first_non_empty_cell(row):
    """
    Return the first non-empty, non-'nan' cell value from a pandas row as a
    stripped string. Returns an empty string if no cell contains text. Because the column headings start in B not A
    """
    for cell in row:
        cell_str = str(cell).strip()
        if cell_str and cell_str.lower() != 'nan':
            return cell_str
    return ''


def is_question_row(row):
    """
    Determine if a row contains a survey question
    """
    first_cell = first_non_empty_cell(row)
    
    # Check for question indicators
    question_patterns = [
        r'^[A-Z]\d+:',  # S1:, A1:, B1:, etc.
        r'^\[.*\]',     # [Age], [Gender], etc.
        r'\?$',         # Ends with question mark
        r'^CTP:',       # CTP questions
        r'^h[A-Z]',     # hSample, hB1_Flag, etc.
    ]
    
    for pattern in question_patterns:
        if re.search(pattern, first_cell):
            return True
    
    # Long descriptive text is likely a question (but NOT response options). The list in response_indicators is recycled from converting a different CSV, but
    # somehow it is not breaking this script. 
    if len(first_cell) > 50 and any(word in first_cell.lower() for word in ['what', 'how', 'which', 'please', 'following']):
        # Make sure it's not a response option that happens to be long
        response_indicators = ['marketing or advertising agency', 'business entity', 'educational institution', 'organization', 'government agency']
        if not any(indicator in first_cell.lower() for indicator in response_indicators):
            return True
    
    return False


def is_header_row(row):
    """
    Determine if a row contains column headers
    """
    first_cell = first_non_empty_cell(row)
    
    # Only "Total" rows that actually look like column headers
    if first_cell == 'Total':
        # Check if it has the structure of a header row (N=xxx values)
        row_text = ' '.join(str(cell) for cell in row)
        if 'N=' in row_text:
            return True
    
    # Also capture descriptive header rows that are specific to the data set. this will have to be updated for every .CSV that needs to be processed. 
    # Do this even if the first cell is empty
    row_text = ' '.join(str(cell) for cell in row if str(cell).strip() and str(cell).strip() != 'nan')

    # Exact pattern for this file, which is different in other files: "Total (A)" followed by the profile type
    # e.g. "Total (A)    Pro-to-Pro Switchers (B)    Software-to-Pro Switchers (C)    Pro-to-Software Switchers (D)"
    if 'Total (A)' in row_text:
        return True

    if 'Total (A)' in row_text and 'Pro-to-Pro Switchers (B)' in row_text and 'Software-to-Pro Switchers (C)' in row_text and 'Pro-to-Software Switchers (D)' in row_text:
        return True

    # Also treat specificheader labels as header rows so they're
    # recognized as descriptive header rows in this CSV.
    specific_headers = [
        "Pro-to-Pro Switchers (B)",
        "Software-to-Pro Switchers (C)",
        "Pro-to-Software Switchers (D)",
    ]
    if any(h in row_text for h in specific_headers):
        return True

    return False


def is_data_row(row):
    """
    Determine if a row contains actual response data
    """
    first_cell = first_non_empty_cell(row)
    
    # Skip completely empty rows
    if not first_cell or first_cell == 'nan':
        return False
    
    # Don't treat question rows as data
    if is_question_row(row):
        return False
    
    # Look for response categories or data
    response_patterns = [
        r'^\d+$',           # Just numbers
        r'%$',              # Ends with %
        r'^\d+\.\d+$',      # Decimal numbers
    ]
    
    # Check if this looks like response data
    non_empty_cells = [str(cell).strip() for cell in row if str(cell).strip() and str(cell).strip() != 'nan']
    
    # If multiple cells have percentage signs or numbers, it's likely data
    numeric_cells = sum(1 for cell in non_empty_cells if re.search(r'\d', cell))
    if numeric_cells >= 2:
        return True
    
    # Common response categories - these are definitely data rows. This was also recycled from previous conversion but isn't breaking this one. 
    response_categories = [
        'Total', 'Mean', 'Median', 'Standard Deviation',
        'Yes', 'No', 'Male', 'Female', 'Other',
        'Very important', 'Somewhat important', 'Not important',
        'Strongly agree', 'Agree', 'Disagree',
        'Currently use', 'Used', 'Never used',
        'A marketing or advertising agency',  # Add this specific case
        'A business entity', 'An educational institution',
        'An organization or non-profit', 'A government agency'
    ]
    
    if any(category in first_cell for category in response_categories):
        return True
    
    # Reasonable length for response option (not too long, not empty)
    return len(first_cell) > 0 and len(first_cell) < 120


def format_survey_csv_to_markdown(df, filename):
    """
    Convert survey CSV data to Markdown preserving question-response structure
    """
    if df.empty:
        return f"# {filename}\n\n*This file is empty.*\n"

    file_title = filename.replace('_', ' ').replace('.csv', '').title()
    markdown_content = f"# {file_title}\n\n"
    
    markdown_content += f"**Total Records:** {len(df)}  \n"
    markdown_content += f"**Total Columns:** {len(df.columns)}  \n"
    markdown_content += f"**Generated:** October 28, 2025  \n\n"
    
    markdown_content += "## Survey Questions and Responses\n\n"
    
    # Find all question boundaries first
    question_boundaries = []
    for index, row in df.iterrows():
        if is_question_row(row):
            question_boundaries.append(index)
    
    # Process each question section completely
    question_number = 0
    for i, question_start in enumerate(question_boundaries):
        # Determine end of this question section
        question_end = question_boundaries[i + 1] if i + 1 < len(question_boundaries) else len(df)
        
        # Get the question text
        question_text = str(df.iloc[question_start].iloc[0]).strip()
        
        # Find all headers and data within this question section
        section_headers = []
        all_header_rows = []  # Store multiple header rows
        section_data = []
        
        for row_idx in range(question_start + 1, question_end):
            row = df.iloc[row_idx]
            
            if is_header_row(row):
                headers = []
                for cell in row:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str != 'nan':
                        headers.append(cell_str)
                    else:
                        headers.append("")
                all_header_rows.append(headers)
                
            elif is_data_row(row):
                data_row = []
                for cell in row:
                    cell_str = str(cell).strip()
                    if cell_str and cell_str != 'nan':
                        data_row.append(cell_str)
                    else:
                        data_row.append("")
                
                if any(data_row):  # Only add if row has some data
                    section_data.append(data_row)
        
        # Combine multiple header rows intelligently. need to update this for this specific CSV .
        header_rows = []
        if all_header_rows:
            descriptive_headers = None
            total_headers = None
            
            # Look for a descriptive header row (e.g. the row with "Total (A)" and
            # the profile labels) and for the N= (sample size) row. Some CSVs put
            # the descriptive labels in a separate row from the N= row, so prefer
            # the descriptive labels as the primary visible header.
            # need to update this for this specific CSV
            switcher_labels = [
                "Pro-to-Pro Switchers (B)",
                "Software-to-Pro Switchers (C)",
                "Pro-to-Software Switchers (D)",
            ]
            for header_row in all_header_rows:
                header_text = ' '.join(header_row)

                # If the row contains the descriptive "Total (A)" or any of the
                # known switcher labels, treat it as the descriptive header.
                if 'Total (A)' in header_text or any(lbl in header_text for lbl in switcher_labels):
                    descriptive_headers = header_row

                # Separately capture rows that include sample-size markers like N=123
                if 'N=' in header_text or re.search(r'N=\d+', header_text):
                    total_headers = header_row

            # If we didn't find a total_headers from is_header_row matches,
            # search the section rows for any row that contains an N= pattern
            # (some CSVs place the N= values in a separate row that wasn't
            # flagged by is_header_row).
            if not total_headers:
                # Search within the section for N= rows
                for row_idx in range(question_start + 1, question_end):
                    row = df.iloc[row_idx]
                    cells = [str(c).strip() for c in row if str(c).strip() and str(c).strip() != 'nan']
                    if any(re.search(r'N=\d+', c) for c in cells):
                        # Build a header-like list from this row
                        th = []
                        for cell in row:
                            cell_str = str(cell).strip()
                            if cell_str and cell_str != 'nan':
                                th.append(cell_str)
                            else:
                                th.append("")
                        total_headers = th
                        break
            # If still not found, look just above the question start in case N= row precedes the question
            if not total_headers:
                start_above = max(0, question_start - 6)
                for row_idx in range(start_above, question_start):
                    row = df.iloc[row_idx]
                    cells = [str(c).strip() for c in row if str(c).strip() and str(c).strip() != 'nan']
                    if any(re.search(r'N=\d+', c) for c in cells):
                        th = []
                        for cell in row:
                            cell_str = str(cell).strip()
                            if cell_str and cell_str != 'nan':
                                th.append(cell_str)
                            else:
                                th.append("")
                        total_headers = th
                        break
            
            # Build header_rows in order: descriptive first (if present), then total/N=
            if descriptive_headers:
                header_rows.append(descriptive_headers)
            if total_headers:
                header_rows.append(total_headers)
        else:
            header_rows = []
        
        # Generate the question section if we have data
        if header_rows or section_data:
            question_number += 1
            markdown_content += format_question_section(
                question_number, question_text, header_rows, section_data
            )
    
    # Add summary
    markdown_content += f"## Summary\n\n"
    markdown_content += f"**Total Questions Processed:** {question_number}  \n"
    markdown_content += f"**Source File:** {filename}  \n"
    
    return markdown_content


def format_question_section(question_num, question, header_rows, data_rows):
    """
    Format a single question section with its response table
    """
    content = f"### Question {question_num}\n\n"
    content += f"**{question}**\n\n"
    
    if not header_rows and not data_rows:
        content += "*No response data available for this question.*\n\n"
        return content
    
    # Create the response table
    # header_rows is a list of header row lists (e.g. [descriptive_headers, total_headers])
    if header_rows:
        # Primary header (used to determine column count) is the first header row
        table_headers = header_rows[0]
    else:
        # Create generic headers based on data width
        max_cols = max(len(row) for row in data_rows) if data_rows else 0
        table_headers = [f"Column {i+1}" for i in range(max_cols)]
    
    if data_rows:
        content += "#### Response Data\n\n"
        
        # Create markdown table using the primary header row as the visible header
        content += "| " + " | ".join(table_headers) + " |\n"
        content += "|" + "---|" * len(table_headers) + "\n"

        # If there is a secondary header row (e.g., N= sample sizes), render it as the
        # first data row so the sample sizes appear directly under the descriptive headers.
        if header_rows and len(header_rows) > 1:
            secondary = header_rows[1]
            # Ensure secondary has correct length
            while len(secondary) < len(table_headers):
                secondary.append("")
            # Mark sample sizes as italic to distinguish them (optional)
            sec_row = [str(c) for c in secondary[:len(table_headers)]]
            content += "| " + " | ".join(sec_row) + " |\n"

        for data_row in data_rows:
            # Ensure row has enough columns
            while len(data_row) < len(table_headers):
                data_row.append("")
            
            # Clean up cell values for markdown
            clean_row = []
            for cell in data_row[:len(table_headers)]:
                clean_cell = str(cell).replace('|', '\\|').replace('\n', ' ')
                clean_row.append(clean_cell)
            
            content += "| " + " | ".join(clean_row) + " |\n"
        
        content += "\n"
    
    return content


def convert_survey_csv_to_markdown(csv_path, output_dir):
    """
    Convert a survey CSV file to Markdown format
    """
    csv_path = Path(csv_path)
    output_dir = Path(output_dir)
    
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Converting survey CSV: {csv_path}")
    print(f"Output directory: {output_dir}")
    
    try:
        print("Reading CSV file...")
        df = pd.read_csv(csv_path, encoding='utf-8')
        
        print(f"Data loaded: {df.shape[0]} rows × {df.shape[1]} columns")
        print("Analyzing survey structure...")
        
        filename = csv_path.name
        markdown_content = format_survey_csv_to_markdown(df, filename)
        
        output_filename = csv_path.stem + '_SURVEY.md'
        output_path = output_dir / output_filename
        

        ## this will overwrite existing files. 
        print(f"Writing survey markdown to: {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        print(f"✅ Successfully converted survey data!")
        print(f"📁 Saved as: {output_filename}")
        
        return output_path
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        raise


def main():
    """
    Main function to execute the conversion
    """
    csv_file_path = r"C:\Users\steph\OneDrive\Desktop\Code\Julia_Naomi\TurboTax Drivers of Switch Data.csv"
    output_directory = r"C:\Users\steph\OneDrive\Desktop\Code\Julia_Naomi"
    
    print("=== Survey CSV to Markdown Converter ===")
    print(f"CSV file: {csv_file_path}")
    print(f"Output folder: {output_directory}")
    print("=" * 50)
    print("🔍 This will analyze the CSV structure to identify:")
    print("   - Survey questions")
    print("   - Response tables")
    print("   - Data headers")
    print("=" * 50)
    
    try:
        output_file = convert_survey_csv_to_markdown(csv_file_path, output_directory)
        
        print(f"\n🎉 Survey conversion complete!")
        print(f"📄 Generated file: {output_file.name}")
        print("\n✅ This file contains:")
        print("   - Each survey question as a section")
        print("   - Response data tables with proper headers")
        print("   - Preserved question-answer relationships")
        print("   - Clean markdown formatting")
            
    except FileNotFoundError as e:
        print(f"\n❌ File not found: {e}")
        print("Please check that the CSV file path is correct.")
        
    except Exception as e:
        print(f"\n❌ An error occurred: {e}")
        print("Please check the file path and try again.")


if __name__ == "__main__":
    main()
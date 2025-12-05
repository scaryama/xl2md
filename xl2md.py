"""
Script to convert XLSX files to Markdown format
Runs safely locally and converts each sheet of Excel files to Markdown tables.

Required packages:
    pip install pandas openpyxl
"""

import os
import sys
import pandas as pd
from pathlib import Path

def sanitize_filename(filename):
    """
    Removes or converts invalid characters from a filename.
    
    Args:
        filename: Original filename
    
    Returns:
        str: Safe filename
    """
    # Remove characters that cannot be used in Windows
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Convert spaces to underscores (optional)
    filename = filename.replace(' ', '_')
    
    # Replace consecutive underscores with a single one
    while '__' in filename:
        filename = filename.replace('__', '_')
    
    return filename.strip('_')

def xlsx_to_markdown(xlsx_path, output_dir=None, log_callback=None):
    """
    Converts each sheet of an XLSX file into individual Markdown files.
    
    Args:
        xlsx_path: Path to the XLSX file to convert
        output_dir: Output directory path (None means same location as original file)
        log_callback: Callback function to output log messages (uses print if None)
    
    Returns:
        bool: Conversion success status
    """
    # Use default print if no callback provided
    log = log_callback if log_callback else print
    
    try:
        # Check if file exists
        if not os.path.exists(xlsx_path):
            log(f"❌ File not found: {xlsx_path}")
            return False
        
        # Set output directory
        if output_dir is None:
            output_dir = Path(xlsx_path).parent
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        # Read Excel file (all sheets)
        excel_file = pd.ExcelFile(xlsx_path)
        sheet_names = excel_file.sheet_names
        
        log(f"📄 Reading file: {xlsx_path}")
        log(f"📊 Number of sheets: {len(sheet_names)}")
        log(f"📁 Output directory: {output_dir}\n")
        
        base_name = Path(xlsx_path).stem
        converted_files = []
        
        # Convert each sheet to individual file
        for idx, sheet_name in enumerate(sheet_names, 1):
            log(f"  📋 Processing sheet {idx}/{len(sheet_names)}: {sheet_name}")
            
            try:
                # Read sheet
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Convert sheet name to filename
                safe_sheet_name = sanitize_filename(sheet_name)
                output_filename = f"{base_name}_{safe_sheet_name}.md"
                output_path = output_dir / output_filename
                
                # Generate Markdown content
                markdown_content = []
                markdown_content.append(f"# {sheet_name}\n")
                markdown_content.append(f"*Source file: {Path(xlsx_path).name}*\n")
                markdown_content.append(f"*Sheet name: {sheet_name}*\n")
                markdown_content.append("---\n\n")
                
                # Handle empty dataframe
                if df.empty:
                    markdown_content.append("*This sheet is empty.*\n")
                else:
                    # Convert NaN values to empty strings
                    df = df.fillna('')
                    
                    # Generate Markdown table
                    markdown_table = df_to_markdown_table(df)
                    markdown_content.append(markdown_table)
                    markdown_content.append("\n")
                
                # Save Markdown file
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(''.join(markdown_content))
                
                converted_files.append(output_path)
                log(f"    ✅ Saved: {output_filename}")
                
            except Exception as e:
                log(f"    ⚠️  Error processing sheet '{sheet_name}': {e}")
                import traceback
                traceback_str = traceback.format_exc()
                log(traceback_str)
        
        log(f"\n✅ Conversion complete: {len(converted_files)} files created")
        return True
        
    except Exception as e:
        log(f"❌ Error occurred: {e}")
        import traceback
        traceback_str = traceback.format_exc()
        log(traceback_str)
        return False

def df_to_markdown_table(df):
    """
    Converts a pandas DataFrame to Markdown table format.
    
    Args:
        df: pandas DataFrame
    
    Returns:
        str: Markdown table string
    """
    if df.empty:
        return "*Table is empty.*"
    
    lines = []
    
    # Generate header
    headers = [str(col) for col in df.columns]
    lines.append("| " + " | ".join(headers) + " |")
    
    # Generate separator line
    separator = "| " + " | ".join(["---"] * len(headers)) + " |"
    lines.append(separator)
    
    # Generate data rows
    for _, row in df.iterrows():
        # Convert each cell value to string and escape pipe characters
        cells = []
        for val in row:
            cell_str = str(val)
            # Handle characters that may cause issues in Markdown tables
            cell_str = cell_str.replace("|", "\\|")
            cell_str = cell_str.replace("\n", "<br>")
            cells.append(cell_str)
        
        lines.append("| " + " | ".join(cells) + " |")
    
    return "\n".join(lines)

def main():
    """Main function"""
    if len(sys.argv) <= 1:
        return
    xlsx_path = sys.argv[1]
    
    # Set output directory (optional)
    output_dir = None
    if len(sys.argv) > 2:
        output_dir = sys.argv[2]
    
    print("=" * 60)
    print("XLSX to Markdown Converter (Separated by Sheet)")
    print("=" * 60)
    
    success = xlsx_to_markdown(xlsx_path, output_dir, log_callback=None)
    
    if success:
        print("\n✨ Conversion completed successfully!")
    else:
        print("\n❌ Conversion failed.")
        sys.exit(1)

if __name__ == "__main__":
    main()


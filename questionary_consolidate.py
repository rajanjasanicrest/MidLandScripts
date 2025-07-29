import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from deep_translator import GoogleTranslator

def translate_to_english_if_needed(text):
    try:
        if not text or not isinstance(text, str):
            return text
        # Very basic heuristic: check for presence of Chinese characters
        if any('\u4e00' <= char <= '\u9fff' for char in text):
            translated = GoogleTranslator(source='auto', target='en').translate(text)
            return translated
        return text
    except Exception as e:
        print(f"Translation error: {e}")
        return text

def translate_table_data(answer_parts):
    """
    Translate tabular data while preserving structure
    """
    translated_parts = []
    for part in answer_parts:
        if '|' in part:
            # Split by separator, translate each cell, then rejoin
            cells = [cell.strip() for cell in part.split('|')]
            # translated_cells = [translate_to_english_if_needed(cell) for cell in cells]
            translated_cells = [cell for cell in cells]
            translated_parts.append(' | '.join(translated_cells))
        else:
            # translated_parts.append(translate_to_english_if_needed(part))
            translated_parts.append(part)
    return translated_parts

def is_tabular_data(answer_parts):
    """
    Determine if the answer contains tabular data based on patterns
    """
    if len(answer_parts) <= 1:
        return False
    
    # Check if multiple rows have similar structure (same number of separators)
    separator_counts = [part.count('|') for part in answer_parts]
    
    # If most rows have the same number of separators and > 1, likely tabular
    if len(set(separator_counts)) <= 2 and max(separator_counts) > 1:
        return True
    
    # Check for common table indicators
    table_keywords = ['header', 'column', 'row', 'total', 'sum', 'average']
    first_row = answer_parts[0].lower()
    
    return any(keyword in first_row for keyword in table_keywords)

def parse_tabular_data(answer_parts):
    """
    Parse tabular data into a 2D array for proper Excel placement
    """
    if not answer_parts:
        return []
    
    # First translate the data while preserving structure
    translated_parts = translate_table_data(answer_parts)
    
    # Split each row by separator
    rows = []
    max_cols = 0
    
    for part in translated_parts:
        cols = [col.strip() for col in part.split('|')]
        rows.append(cols)
        max_cols = max(max_cols, len(cols))
    
    # Pad rows to same length
    for row in rows:
        while len(row) < max_cols:
            row.append("")
    
    return rows

def format_regular_data(answer_parts):
    """
    Format non-tabular data
    """
    return "\n".join(answer_parts).strip()

def create_enhanced_excel(master_data, all_suppliers):
    """
    Create Excel file with reserved columns for each supplier (8 columns each)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidated Data"
    
    # Define styling
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_header_font = Font(bold=True, color="366092", size=11)
    data_font = Font(name="Calibri", size=10)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Calculate column layout
    COLS_PER_SUPPLIER = 8
    SPACING_COLS = 2  # Spacing columns between suppliers
    base_cols = 2  # Sheet Name + Question columns
    
    # Create headers
    current_row = 1
    
    # Main headers - merge across supplier columns
    ws.cell(row=current_row, column=1, value="Sheet Name")
    ws.cell(row=current_row, column=2, value="Question")
    
    col_start = base_cols + 1
    for i, supplier in enumerate(all_suppliers):
        # Merge cells for supplier header
        start_col = col_start
        end_col = col_start + COLS_PER_SUPPLIER - 1
        
        ws.merge_cells(f"{get_column_letter(start_col)}{current_row}:{get_column_letter(end_col)}{current_row}")
        ws.cell(row=current_row, column=start_col, value=supplier)
        
        # Move to next supplier position (including spacing)
        col_start = end_col + 1 + SPACING_COLS
    
    # Apply header styling
    for col in range(1, col_start):
        cell = ws.cell(row=current_row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    # Sub-headers for supplier columns
    current_row += 1
    ws.cell(row=current_row, column=1, value="")  # Empty for sheet name
    ws.cell(row=current_row, column=2, value="")  # Empty for question

    col_start = base_cols + 1
    for supplier in all_suppliers:
        for i in range(COLS_PER_SUPPLIER):
            col_num = col_start + i
            ws.cell(row=current_row, column=col_num, value=f"Col {i+1}")
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = sub_header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        # Add spacing columns (empty)
        for s in range(SPACING_COLS):
            col_num = col_start + COLS_PER_SUPPLIER + s
            ws.cell(row=current_row, column=col_num, value="")
            cell = ws.cell(row=current_row, column=col_num)
            cell.border = border
        # Move to next supplier position (including spacing)
        col_start += COLS_PER_SUPPLIER + SPACING_COLS

    # Style the first two columns in sub-header row
    for col in [1, 2]:
        cell = ws.cell(row=current_row, column=col)
        cell.font = sub_header_font
        cell.border = border
    
    # Data rows
    current_row += 1

    for (sheet, question), supplier_answers in master_data.items():
        row_start = current_row
        max_rows_needed = 1

        # First, determine how many rows we need for this question
        for supplier in all_suppliers:
            answer = supplier_answers.get(supplier, "")
            if answer:
                answer_parts = answer.split('\n') if isinstance(answer, str) else []
                # Check if it's tabular data that was processed
                if '|' in answer and ('┌' in answer or '├' in answer):
                    # This is a formatted table, count actual data rows
                    data_lines = [line for line in answer_parts if line.strip() and not line.startswith('┌') and not line.startswith('├') and not line.startswith('└') and '─' in line]
                    max_rows_needed = max(max_rows_needed, len(answer_parts))
                else:
                    # Check if original data was tabular
                    original_parts = answer.split('\n')
                    if any('|' in part for part in original_parts):
                        # Parse as table
                        table_data = parse_tabular_data([part for part in original_parts if '|' in part])
                        max_rows_needed = max(max_rows_needed, len(table_data))
                    else:
                        max_rows_needed = max(max_rows_needed, len(answer_parts))

        # Set sheet name and question (merge across all rows if multi-row)
        if max_rows_needed > 1:
            ws.merge_cells(f"A{row_start}:A{row_start + max_rows_needed - 1}")
            ws.merge_cells(f"B{row_start}:B{row_start + max_rows_needed - 1}")

        ws.cell(row=row_start, column=1, value=sheet)
        ws.cell(row=row_start, column=2, value=question)

        # Style sheet and question cells
        for col in [1, 2]:
            cell = ws.cell(row=row_start, column=col)
            cell.font = data_font
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            cell.border = border

        # Fill supplier data
        col_start = base_cols + 1
        for supplier in all_suppliers:
            answer = supplier_answers.get(supplier, "")

            if answer and isinstance(answer, str):
                # Check if this looks like processed tabular data or original tabular data
                if ('|' in answer and not ('┌' in answer or '├' in answer)) or answer.count('|') > answer.count('\n'):
                    # This is raw tabular data, parse it properly
                    original_parts = [part.strip() for part in answer.split('\n') if part.strip()]
                    table_parts = [part for part in original_parts if '|' in part]

                    if table_parts:
                        table_data = parse_tabular_data(table_parts)

                        # Place table data in grid
                        for row_idx, row_data in enumerate(table_data):
                            for col_idx, cell_value in enumerate(row_data[:COLS_PER_SUPPLIER]):
                                if cell_value.strip():
                                    target_row = row_start + row_idx
                                    target_col = col_start + col_idx
                                    ws.cell(row=target_row, column=target_col, value=cell_value)

                                    cell = ws.cell(row=target_row, column=target_col)
                                    cell.font = data_font
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                                    cell.border = border
                    else:
                        # Non-tabular multi-line data
                        lines = answer.split('\n')
                        for line_idx, line in enumerate(lines[:max_rows_needed]):
                            if line.strip():
                                target_row = row_start + line_idx
                                ws.merge_cells(f"{get_column_letter(col_start)}{target_row}:{get_column_letter(col_start + COLS_PER_SUPPLIER - 1)}{target_row}")
                                ws.cell(row=target_row, column=col_start, value=line)

                                cell = ws.cell(row=target_row, column=col_start)
                                cell.font = data_font
                                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                                cell.border = border
                else:
                    # Single value or simple text - merge across all supplier columns
                    ws.merge_cells(f"{get_column_letter(col_start)}{row_start}:{get_column_letter(col_start + COLS_PER_SUPPLIER - 1)}{row_start + max_rows_needed - 1}")
                    ws.cell(row=row_start, column=col_start, value=answer)

                    cell = ws.cell(row=row_start, column=col_start)
                    cell.font = data_font
                    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                    cell.border = border

            # Add borders to empty cells in this supplier's columns
            for col_offset in range(COLS_PER_SUPPLIER):
                for row_offset in range(max_rows_needed):
                    target_row = row_start + row_offset
                    target_col = col_start + col_offset
                    cell = ws.cell(row=target_row, column=target_col)
                    if not cell.border.left.style:  # Only add border if not already set
                        cell.border = border
            # Add borders to spacing columns
            for s in range(SPACING_COLS):
                for row_offset in range(max_rows_needed):
                    target_row = row_start + row_offset
                    target_col = col_start + COLS_PER_SUPPLIER + s
                    cell = ws.cell(row=target_row, column=target_col)
                    cell.border = border

            col_start += COLS_PER_SUPPLIER + SPACING_COLS

        current_row += max_rows_needed
    
    # Adjust column widths
    # Fixed width for first two columns
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 40

    # Set width for supplier columns and spacing columns
    col_start = base_cols + 1
    for supplier in all_suppliers:
        for i in range(COLS_PER_SUPPLIER):
            col_letter = get_column_letter(col_start + i)
            ws.column_dimensions[col_letter].width = 15
        # Set width for spacing columns
        for s in range(SPACING_COLS):
            col_letter = get_column_letter(col_start + COLS_PER_SUPPLIER + s)
            ws.column_dimensions[col_letter].width = 5
        col_start += COLS_PER_SUPPLIER + SPACING_COLS
    
    return wb

folder = "questionaries"
master_data = {}  # key: (sheet_name, question_text), value: {supplier: answer}

print("Processing files...")
for file in os.listdir(folder):
    if not file.endswith(".xlsx"):
        continue

    supplier = os.path.splitext(file)[0].split('--')[-1]
    filepath = os.path.join(folder, file)
    print(f"Processing {supplier}...")
    
    try:
        wb = load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"Error loading {file}: {e}")
        continue

    for sheetname in wb.sheetnames[1:]:  # Skip Sheet1 (index 0)
        ws = wb[sheetname]
        row = 1
        while row <= ws.max_row:
            q_cell = ws[f"C{row}"]
            if q_cell.value:
                question = str(q_cell.value).strip()

                # Fetch answer from merged D–L area starting from the same row
                answer_parts = []
                for ans_row in range(row, row + 10):  # Look ahead for merged block
                    if ans_row > ws.max_row:
                        break
                    row_values = [ws.cell(row=ans_row, column=col).value for col in range(4, 13)]
                    non_empty = [str(val).strip() for val in row_values if val not in [None, ""]]
                    if non_empty:
                        answer_parts.append(" | ".join(non_empty))
                    else:
                        # stop if empty row after some answer found
                        if answer_parts:
                            break

                # Smart formatting and translation
                if is_tabular_data(answer_parts):
                    # Keep raw tabular data for proper Excel placement
                    translated_parts = translate_table_data(answer_parts)
                    formatted_answer = "\n".join(translated_parts)
                else:
                    formatted_answer = format_regular_data(answer_parts)
                    # formatted_answer = translate_to_english_if_needed(formatted_answer)

                # store the data
                key = (sheetname, question)
                if key not in master_data:
                    master_data[key] = {}
                master_data[key][supplier] = formatted_answer
            row += 1

print("Building consolidated data...")

# Build data
all_suppliers = sorted({supplier for answers in master_data.values() for supplier in answers})

# Create enhanced Excel file
print("Creating enhanced Excel file with reserved columns...")
wb = create_enhanced_excel(master_data, all_suppliers)

# Save file
output_path = "Consolidated Questionaries WO Transation.xlsx"
wb.save(output_path)

print(f"Done! Enhanced file saved at: {output_path}")
print(f"Processed {len(master_data)} questions from {len(all_suppliers)} suppliers.")
print(f"Each supplier has {8} reserved columns for proper table display.")

# Create summary
print("Creating summary sheet...")
summary_data = []
for sheet_name in set(key[0] for key in master_data.keys()):
    sheet_questions = [key for key in master_data.keys() if key[0] == sheet_name]
    summary_data.append({
        'Sheet Name': sheet_name,
        'Number of Questions': len(sheet_questions),
        'Suppliers with Answers': len(set(supplier for key in sheet_questions for supplier in master_data[key].keys()))
    })

# Add summary sheet
summary_ws = wb.create_sheet("Summary")
summary_ws.append(['Sheet Name', 'Number of Questions', 'Suppliers with Answers'])

for data in summary_data:
    summary_ws.append([data['Sheet Name'], data['Number of Questions'], data['Suppliers with Answers']])

# Style summary sheet
for row in summary_ws.iter_rows():
    for cell in row:
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        if cell.row == 1:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")

wb.save(output_path)
print("Summary sheet added successfully!")
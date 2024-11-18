import re
from docx import Document
from docx.shared import Inches
import os

def detect_placeholders(doc):
    """Detect placeholders in the form of {placeholder} within paragraphs, processing top-to-bottom."""
    placeholders = set()  # Using a set to avoid duplicates
    for paragraph in doc.paragraphs:
        matches = re.findall(r'\{(.*?)\}', paragraph.text)  # Find placeholders inside curly braces
        for match in matches:
            placeholders.add(match)
    return list(placeholders)

def collect_input_for_placeholders(placeholders):
    """Prompt the user to input values for each placeholder."""
    replacements = {}
    print("\nEnter values for each placeholder:")
    for placeholder in placeholders:
        if placeholder != "logo":  # Skip logo placeholder as it will be handled separately
            value = input(f"  {placeholder}: ")
            replacements[placeholder] = value
    return replacements

def replace_placeholders_in_paragraphs(doc, replacements, logo_path):
    """Replace placeholders in paragraphs with the corresponding user input or images."""
    
    # Process body paragraphs
    for paragraph in doc.paragraphs:
        for placeholder, replacement in replacements.items():
            if f"{{{placeholder}}}" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"{{{placeholder}}}", replacement)

        # Handle image placeholder (logo) in body paragraphs
        if "{logo}" in paragraph.text:
            paragraph.clear()
            run = paragraph.add_run()
            run.add_picture(logo_path, width=Inches(1.5))
    
    # Process headers
    for section in doc.sections:
        header = section.header
        for paragraph in header.paragraphs:
            for placeholder, replacement in replacements.items():
                if f"{{{placeholder}}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace(f"{{{placeholder}}}", replacement)

            # Handle logo placeholder in header
            if "{logo}" in paragraph.text:
                paragraph.clear()
                run = paragraph.add_run()
                run.add_picture(logo_path, width=Inches(1.5))

def generate_document_with_placeholders(template_path, logo_path, output_path):
    """Generate a document by replacing placeholders with user input or images."""
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    doc = Document(template_path)

    # Replace placeholders inside the document (top to bottom)
    placeholders = detect_placeholders(doc)
    print("\nDetected placeholders:")
    for placeholder in placeholders:
        print(f"  - {placeholder}")
    
    # Collect replacements for placeholders (excluding image placeholders)
    replacements = collect_input_for_placeholders(placeholders)

    # Replace placeholders in paragraphs with user input or images
    replace_placeholders_in_paragraphs(doc, replacements, logo_path)

    # Process tables and add data
    if len(doc.tables) > 0:
        table = doc.tables[0]
        headers = get_table_headers(table)
        print(f"Detected table headers: {headers}")
        table_data = collect_data_from_user(headers)

        for i, row_data in enumerate(table_data):
            if i + 1 < len(table.rows):
                row = table.rows[i + 1]
            else:
                row = table.add_row()

            apply_borders_to_row(row)

            for j, cell_data in enumerate(row_data):
                if j < len(row.cells):
                    row.cells[j].text = cell_data
    else:
        raise ValueError("No tables found in the document template.")

    # Save the modified document
    doc.save(output_path)
    print(f"Document saved to: {output_path}")

# Example usage
template_path = ""
logo_path = ""
output_path = ""

generate_document_with_placeholders(template_path, logo_path, output_path)

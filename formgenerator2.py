from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import os

def detect_placeholders(doc):
    placeholders = []

    for paragraph in doc.paragraphs:
        matches = re.findall(r"\{.*?\}", paragraph.text)
        placeholders.extend(matches)  

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    matches = re.findall(r"\{.*?\}", paragraph.text)
                    placeholders.extend(matches)

    return list(dict.fromkeys(placeholders))


def replace_placeholder_with_image(cell, placeholder, image_path, width_in_inches=1):
    if placeholder in cell.text:
        for paragraph in cell.paragraphs:
            if placeholder in paragraph.text:
                parts = paragraph.text.split(placeholder)
                paragraph.text = parts[0] 
                run = paragraph.add_run()
                run.add_picture(image_path, width=Inches(width_in_inches))
                if len(parts) > 1:
                    paragraph.add_run(parts[1])
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


def replace_placeholder_with_text(cell, placeholder, value):
    if placeholder in cell.text:
        for paragraph in cell.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)


def generate_document_from_template(
    template_path, output_path, replacements, image_replacements
):
    doc = Document(template_path)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for placeholder, value in replacements.items():
                    replace_placeholder_with_text(cell, placeholder, value)
                for placeholder, image_path in image_replacements.items():
                    replace_placeholder_with_image(cell, placeholder, image_path)

    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)
        for placeholder, image_path in image_replacements.items():
            if placeholder in paragraph.text:
                parts = paragraph.text.split(placeholder)
                paragraph.text = parts[0]
                run = paragraph.add_run()
                run.add_picture(image_path, width=Inches(1))
                if len(parts) > 1:
                    paragraph.add_run(parts[1])

    doc.save(output_path)
    print(f"Document saved to: {output_path}")


def main():
    template_path = input("Enter the path to the template document (.docx): ").strip()

    doc = Document(template_path)
    placeholders = detect_placeholders(doc)
    print("\nDetected placeholders:")
    for placeholder in placeholders:
        print(f"  - {placeholder}")

    replacements = {}
    image_replacements = {}
    print("\nEnter values for each placeholder:")

    for placeholder in placeholders:
        if placeholder.lower().startswith("{logo") or "sign" in placeholder.lower():
            image_path = input(f"path to {placeholder.strip('{}')}: ").strip()
            image_replacements[placeholder] = image_path
        else:
            text_value = input(f"{placeholder.strip('{}')}: ").strip()
            replacements[placeholder] = text_value

    output_path = input("\nEnter the path to save the generated document: ").strip()

    generate_document_from_template(template_path, output_path, replacements, image_replacements)


if __name__ == "__main__":
    main()

import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import win32com.client  # For Word-to-PDF conversion


excel_path = os.path.abspath("data.xlsx")
word_template_path = os.path.abspath("template.docx")
output_dir = os.path.abspath("output_pdfs")
os.makedirs(output_dir, exist_ok=True)

print("Paths:")
print(f"Excel Path: {excel_path}")
print(f"Word Template Path: {word_template_path}")
print(f"Output Directory: {output_dir}")


try:
    data = pd.read_excel(excel_path)
    data.columns = data.columns.str.strip().str.lower()
    print("Excel Data Loaded Successfully!")
    print(data.head())
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()


def replace_placeholders(doc, replacements):
    for para in doc.paragraphs:
        for placeholder, replacement in replacements.items():
            if placeholder in para.text:
                for run in para.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, replacement)
                        run.bold = True

# Convert Word to PDF
def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        print(f"Converted to PDF: {pdf_path}")
    except Exception as e:
        print(f"Error converting Word to PDF: {e}")

# Replace placeholders and save PDFs
def replace_placeholders_and_save_pdf(row):
    try:
        print(f"Processing: Name={row['name']}, Company={row['company']}, Post={row['post']}")

        doc = Document(word_template_path)

        # Standardize name for comparison
        name_lower = row["name"].strip().lower()

        # Determine salutation based on name
        if "mr" in name_lower:
            salutation = "Dear Sir,"
        elif "ms" in name_lower:
            salutation = "Dear Ma'am,"
        else:
            salutation = "Dear Sir/Ma'am"

        replacements = {
            "{name}": row["name"],
            "{company}": row["company"],
            "{post}": row["post"],
            "{s}": salutation,
        }

        replace_placeholders(doc, replacements)

        temp_word_path = os.path.join(output_dir, f"{row['company']}_temp.docx")
        doc.save(temp_word_path)
        print(f"Temporary Word file saved at: {temp_word_path}")

        pdf_name = f"{row['company']}_Invitations_MMMUT.pdf"
        pdf_path = os.path.join(output_dir, pdf_name)
        convert_docx_to_pdf(temp_word_path, pdf_path)

        os.remove(temp_word_path)
        print(f"Temporary Word file deleted: {temp_word_path}")

    except Exception as e:
        print(f"Error processing row for {row['name']}: {e}")


print("\nStarting PDF generation...")
for _, row in data.iterrows():
    replace_placeholders_and_save_pdf(row)

print("\nAll PDFs generated successfully!")

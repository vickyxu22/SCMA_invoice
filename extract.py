from docx import Document
import os
import re
import pandas as pd

def extract_data(paragraph_text, keyword):
    segment = paragraph_text.split(keyword)
    return segment[1].strip() if len(segment) > 1 else None

def extract_information(doc):
    data_attributes = {
        "Invoice #:": "invoice",
        "Formats:": "formats",
        "Author:": "author",
        "Title of Publication:": "title",
        "Publisher:": "publisher",
        "Publication date:": "publication_year",
        "Print-run:": "print_num",
        "Distribution:": "distribution",
        "Language:": "language",
        "Type of Use:": "type_of_use",
    }

    extracted_data = {attr: None for attr in data_attributes.values()}

    for paragraph in doc.paragraphs:
        for keyword, attribute in data_attributes.items():
            if keyword in paragraph.text:
                extracted_data[attribute] = extract_data(paragraph.text, keyword)

                if keyword == "Publication date:":
                    match = re.search(r'\b\d{4}\b', extracted_data[attribute])
                    if match:
                        extracted_data[attribute] = match.group()

    return extracted_data

def create_note(row):
    note = (
        f"Print-run: {row['print_num']}. "
        f"Distribution: {row['distribution']}. "
        f"Language: {row['language']}. "
        f"Type of Use: {row['type_of_use']}."
    )
    return note

def process_documents(folder_path):
    all_data = []

    try:
        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            file_names = [file_name for file_name in os.listdir(folder_path) if file_name.endswith('.docx')]

            for file_name in file_names:
                file_path = os.path.join(folder_path, file_name)
                print(f"Processing file '{file_name}' at: {file_path}")

                doc = Document(file_path)
                extracted_info = extract_information(doc)
                all_data.append(extracted_info)

            df = pd.DataFrame(all_data)
            df['note'] = df.apply(create_note, axis=1)

            df = df.rename(columns={
                "invoice": "Invoice Number",
                "formats": "File Formats",
                "author": "Author",
                "title": "Publication Title",
                "publisher": "Publisher",
                "publication_year": "Publication Year",
                "print_num": "Print Run",
                "distribution": "Distribution",
                "language": "Language",
                "type_of_use": "Type of Use",
                "note": "Note"
            })

            output_file = os.path.join(folder_path, "combined_extracted_info.xlsx")
            df.to_excel(output_file, index=False)
            print(f"Combined extracted information has been exported to '{output_file}'.")

        else:
            print(f"Folder not found at path: '{folder_path}'.")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    desktop_path = os.path.expanduser("~/Desktop")
    folder_name = "SCMA_invoice"
    folder_path = os.path.join(desktop_path, folder_name)
    process_documents(folder_path)

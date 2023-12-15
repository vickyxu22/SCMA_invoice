from docx import Document
import os
import re
import pandas as pd

def extract_data(paragraph_text, keyword):
    segment = paragraph_text.split(keyword)
    return segment[1].strip() if len(segment) > 1 else None

def capitalize_except_prepositions(text):
    words = text.split()
    prepositions = ['a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'from', 'by', 'with', 'in', 'of']

    capitalized_text = ' '.join(word.capitalize() if word.lower() not in prepositions else word.lower() for word in words)
    return capitalized_text

def extract_information(doc):
    data_attributes = {
        "Invoice #:": "invoice",
        "Formats:": "formats",
        "Format:": "formats",
        "Type:": "formats",
        "Use:": "formats",
        "Author:": "author",
        "Authors:": "author",
        "Author/s:": "author",
        "Editor:": "editor",
        "Editors:": "editor",
        "Title of Publication:": "title",
        "Title of the Publication:": "title",
        "Title:": "title",
        "Publisher:": "publisher",
        "Permission granted to:": "publisher",
        "Publication date:": "publication_year",
        "Publication Date:": "publication_year",
        "Date of publication:": "publication_year",
        "Print Date:": "publication_year",
        "Print-run:": "print_num",
        "Print run:": "print_num",
        "Print run/number of units:": "print_num",
        "Print Run:": "print_num",
        "Size of print run:": "print_num",
        "Print run size：": "print_num", 
        "Distribution:": "distribution",
        "distribution:": "distribution",
        "distributed by:": "distribution",
        "Type of Use:": "type_of_use",
        "Language:": "language",
        "Languages:": "language",
        "Language of Publication:": "language",
        "Language of publication:": "language",
    }

    extracted_data = {attr: None for attr in data_attributes.values()}

    for paragraph in doc.paragraphs:
        for keyword, attribute in data_attributes.items():
            if keyword in paragraph.text:
                extracted_data[attribute] = capitalize_except_prepositions(extract_data(paragraph.text, keyword))

                if keyword == "Invoice #:":
                    match_invoice = re.search(r'\b\d{3}\b', extracted_data[attribute])
                    if match_invoice:
                        extracted_data[attribute] = match_invoice.group()

                if keyword == "Formats:" or keyword == "Format:" or keyword == "Use:" or keyword == "Type:":
                    match_print = re.search(r'\bPrint\b', extracted_data[attribute])
                    if match_print:
                        extracted_data[attribute] = extracted_data[attribute].replace('Print', 'Book')
                    lang_position = extracted_data[attribute].find('Language:')
                    if lang_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:lang_position]
                    loc_position = extracted_data[attribute].find('Location:')
                    if loc_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:loc_position]

                if keyword == "Author:" or keyword == "Authors:" or keyword == "Author/s:":
                    title_position = extracted_data[attribute].find('Title')
                    date_position = extracted_data[attribute].find('Date')
                    if title_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:title_position]
                    if date_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:date_position]

                if keyword == "Editor:" or keyword == "Editors:":
                    title_position = extracted_data[attribute].find('Title')
                    if title_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:title_position]
                
                if keyword == "Publisher:" or keyword == "Permission granted to:":
                    pub_position = extracted_data[attribute].find('Publication date')
                    if pub_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:pub_position]

                if keyword == "Title of Publication:" or keyword == "Title:" or keyword == "Title of the Publication:":
                    publisher_position = extracted_data[attribute].find('Publisher')
                    if publisher_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:publisher_position]

                if keyword == "Publication date:" or keyword == "Date of publication:" or keyword == "Publication Date:" or keyword == "Print Date:":
                    match = re.search(r'\b\d{4}\b', extracted_data[attribute])
                    if match:
                        extracted_data[attribute] = match.group()
                    dis_position = extracted_data[attribute].find('Distribution:')
                    if dis_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:dis_position]

                if keyword == "Print-run:" or keyword == "Print run:" or keyword == "Print Run:" or keyword == "Print run/number of units:" or keyword == "Size of print run:" or keyword == "Print run size:":
                    term_position = extracted_data[attribute].find('Term:')
                    if term_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:term_position]
                    loc_position = extracted_data[attribute].find('Location:')
                    if loc_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:loc_position]
                    format_position = extracted_data[attribute].find('Format:')
                    if format_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:format_position]
                
                if keyword == "Distribution:" or keyword == "distribution:" or keyword == "distributed by:":
                    lang_position = extracted_data[attribute].find('Language:')
                    if lang_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:lang_position]
                    langp_position = extracted_data[attribute].find('Language of Publication:')
                    if langp_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:langp_position]
                    pub_position = extracted_data[attribute].find('Publication Date:')
                    if pub_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:pub_position]
                        

                if keyword == "Language:" or keyword == "Languages:" or keyword == "Language of Publication" or keyword == "Language of publication":
                    print_position = extracted_data[attribute].find('Print-run:')
                    if print_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:print_position]
                    dis_position = extracted_data[attribute].find('Distribution:')
                    if dis_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:dis_position]
                
                if keyword == "Type of Use:":
                    print_position = extracted_data[attribute].find('Print-run:')
                    lang_position = extracted_data[attribute].find('Language:')
                    pub_position = extracted_data[attribute].find('Publication date')
                    if print_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:print_position]
                    if lang_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:lang_position]
                    if pub_position != -1:
                        extracted_data[attribute] = extracted_data[attribute][:pub_position]
                    if extracted_data[attribute] is not None:
                        if isinstance(extracted_data['distribution'], str) and isinstance(extracted_data[attribute], str):
                            extracted_data['distribution'] += f"; {extracted_data[attribute]}"
                        elif extracted_data['distribution'] is None:
                            extracted_data['distribution'] = extracted_data[attribute]

    return extracted_data

def create_note(row):
    print_run = f"Print-run: {row['print_num']}. " if row['print_num'] else ''
    distribution = f"Distribution: {row['distribution']}. " if row['distribution'] else ''
    language = f"Language: {row['language']}. " if row['language'] else ''

    note = f"{print_run}{distribution}{language}".strip()
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
                "editor": "Editor",
                "title": "Publication Title",
                "publisher": "Publisher",
                "publication_year": "Publication Year",
                "print_num": "Print Run",
                "distribution": "Distribution",
                "language": "Language",
                "type_of_use": "Type of Use",
                "note": "Note"
            })

            df = df.drop(columns=['Type of Use'])

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

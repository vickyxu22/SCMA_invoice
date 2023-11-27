from docx import Document
import os
import pandas as pd

# Directory path to search within the Desktop folder
desktop_path = os.path.expanduser("~/Desktop")  # Path to the Desktop folder

folder_name = "scma"  # Name of the folder within Desktop
file_names = ["Inv641.docx"]  # Names of the files to find within the 'scma' folder

folder_path = os.path.join(desktop_path, folder_name)  # Path to the 'scma' folder

# Lists to store extracted information
all_author_list = []
all_publication_list = []

try:
    # Check if the folder exists
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        for file_name in file_names:
            # Check if the desired file exists within the folder
            if file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)
                print(f"File '{file_name}' found at: {file_path}")

                # Extract information from each file
                author_list = []
                publication_list = []

                doc = Document(file_path)

                # Iterate through paragraphs in the document
                for paragraph in doc.paragraphs:
                    if "Author:" in paragraph.text:
                        author_info = paragraph.text.split("Author:", 1)[-1].strip()
                        author_list.append(author_info)
                    elif "Title of Publication:" in paragraph.text:
                        publication_info = paragraph.text.split("Title of Publication:", 1)[-1].strip()
                        publication_list.append(publication_info)

                # Append extracted data to the lists
                all_author_list.extend(author_list)
                all_publication_list.extend(publication_list)

            else:
                print(f"File '{file_name}' not found in the '{folder_name}' folder.")

        # Make sure 'Author' and 'Publication' lists have the same length
        max_length = max(len(all_author_list), len(all_publication_list))
        all_author_list += [''] * (max_length - len(all_author_list))
        all_publication_list += [''] * (max_length - len(all_publication_list))

        # Create a DataFrame from all extracted data
        data = {'Author': all_author_list, 'Publication': all_publication_list}
        df = pd.DataFrame(data)

        # Export DataFrame to Excel
        output_file = os.path.join(desktop_path, "combined_extracted_info.xlsx")
        df.to_excel(output_file, index=False)
        print(f"Combined extracted information has been exported to '{output_file}'.")

    else:
        print(f"Folder '{folder_name}' not found on the Desktop.")

except Exception as e:
    print(f"An error occurred: {e}")

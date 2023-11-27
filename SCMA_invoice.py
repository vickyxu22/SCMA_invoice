from docx import Document
import os
import pandas as pd

# Directory path to search within the Desktop folder
desktop_path = os.path.expanduser("~/Desktop")  # Path to the Desktop folder
folder_name = "scma"  # Name of the folder within Desktop
folder_path = os.path.join(desktop_path, folder_name)  # Path to the 'scma' folder

# List to store extracted information
all_data = []

try:
    # Check if the folder exists
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        file_names = os.listdir(folder_path)  # Get all files in the folder

        # Filter only .docx files
        file_names = [file_name for file_name in file_names if file_name.endswith('.docx')]

        for file_name in file_names:
            file_path = os.path.join(folder_path, file_name)
            print(f"Processing file '{file_name}' at: {file_path}")

            doc = Document(file_path)
            author = None
            publication = None

            # Extract information from each file
            for paragraph in doc.paragraphs:
                if "Author:" in paragraph.text:
                    # Split the paragraph into segments based on "Author:" and "Title of Publication:"
                    segments = paragraph.text.split("Author:")

                    # Process each segment
                    for segment in segments[1:]:
                        parts = segment.split("Title of Publication:")
                        if len(parts) == 2:
                            author = parts[0].strip()
                            publication = parts[1].strip()

                            # Append extracted data to the list as a tuple (author, publication)
                            all_data.append((author, publication))
                        else:
                            print(f"Error: Unable to extract Author and Title of Publication from '{file_name}'.")

        # Create a DataFrame from all extracted data
        data = {'Author': [entry[0] for entry in all_data], 'Publication': [entry[1] for entry in all_data]}
        df = pd.DataFrame(data)

        # Export DataFrame to Excel
        output_file = os.path.join(desktop_path, "combined_extracted_info.xlsx")
        df.to_excel(output_file, index=False)
        print(f"Combined extracted information has been exported to '{output_file}'.")

    else:
        print(f"Folder '{folder_name}' not found on the Desktop.")

except Exception as e:
    print(f"An error occurred: {e}")


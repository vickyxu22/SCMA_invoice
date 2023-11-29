# SCMA_docx_extractor

`docx_extractor` is a Python package for extracting publication text from Smith College Museum of Art invoice documents in DOCX format. The script parses the documents to retrieve attributions such as author, title, format, publication date, etc. The extracted information is then merged into an Excel file for convenient usage and integration into the Mimsy database.

## Installation

To install `docx_extractor`, use pip:

```bash
pip install docx-extractor
```

## Usage
The package provides a function process_documents(folder_path) to extract information from .docx files within a specified folder.

Example usage:
```
from docx_extractor.extract import process_documents

# Replace 'folder_path' with the path to your folder containing .docx files
folder_path = '/path/to/your/folder'
process_documents(folder_path)
```

## License
This project is licensed under the MIT License - see the LICENSE file for details.

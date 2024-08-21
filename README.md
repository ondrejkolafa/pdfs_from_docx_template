
# Document Generator README

## Overview
This Python script automates the generation of Word documents and PDF files by merging data from an Excel file with a Word template. The script processes each row of the Excel file, populates the Word template with the corresponding data, saves the populated template as a new document, and then converts it to a PDF. It also includes options for organizing the output into folders and cleaning up intermediate files.

## Features
- **Dynamic Document Generation**: Merge Excel data with a Word template to generate customized documents.
- **PDF Conversion**: Automatically convert generated Word documents into PDF format.
- **Foldered Output**: Optionally save the generated documents in individual folders.
- **Cleanup Option**: Optionally delete Word documents after PDF conversion to save space.

## Dependencies
The script relies on the following Python packages:
- `argparse`: For parsing command-line arguments.
- `glob`: For finding Word and Excel files.
- `docxtpl`: For populating Word templates.
- `docx2pdf`: For converting Word documents to PDF.
- `os`: For handling file and folder operations.
- `pandas`: For reading and processing Excel data.

Install the required packages using `pip`:
\`\`\`bash
pip install docxtpl docx2pdf pandas
\`\`\`

## Usage

### Command-Line Arguments
The script accepts the following command-line arguments:

- `-w, --word`: (Required) Name of the Word template file.
- `-e, --excel`: (Required) Name of the Excel data file.
- `-f, --foldered`: (Optional) Save the output files in separate folders.
- `-c, --cleanup`: (Optional) Delete the intermediate Word files after conversion to PDF.

### Example Commands
1. **Basic Usage:**
   \`\`\`bash
   python script.py -w template.docx -e data.xlsx
   \`\`\`
   This command will generate documents from `data.xlsx` using the `template.docx` file.

2. **Organizing Output in Folders:**
   \`\`\`bash
   python script.py -w template.docx -e data.xlsx -f
   \`\`\`
   This command will save the generated documents in separate folders.

3. **Cleaning Up Intermediate Files:**
   \`\`\`bash
   python script.py -w template.docx -e data.xlsx -c
   \`\`\`
   This command will delete the Word files after converting them to PDFs.

4. **Combining Foldered Output and Cleanup:**
   \`\`\`bash
   python script.py -w template.docx -e data.xlsx -f -c
   \`\`\`
   This command will save the generated documents in separate folders and then delete the Word files after conversion.

### Execution Flow
1. **Loading Data**: The script loads data from the provided Excel file into a dictionary. It also identifies a column with unique values to use as identifiers for naming files.
2. **Populating Templates**: Each row of data is used to populate the Word template, generating a new document.
3. **Saving Documents**: Depending on the `--foldered` option, the generated documents are saved either in the output folder or in separate folders.
4. **Converting to PDF**: The generated Word documents are converted to PDF format using `docx2pdf`.
5. **Cleanup (Optional)**: If the `--cleanup` flag is set, the intermediate Word files are deleted after conversion.

## Output
Generated documents and PDFs will be saved in the `output` folder. If the `--foldered` option is used, each document will be saved in a separate subfolder.

## Error Handling
The script handles the following errors:
- **File Not Found**: If the specified Word template or Excel data file is not found, an error message is displayed and the script exits.
- **Template Parsing Error**: If the script fails to populate the Word template, it will display an error message.

## Conclusion
This script provides an efficient way to generate and manage documents by combining data from Excel files with Word templates. The command-line options allow for flexibility in output organization and cleanup.

import argparse
import glob
from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
import docx2pdf
import os
import pandas as pd

OUTPUT_FOLDER = "output"


def main(word_template_name, data_file_name, foldered, cleanup):

    data_dict, unique_column_name = load_data(data_file_name)

    doc = DocxTemplate(word_template_name)

    for k, data in data_dict.items():
        context = {key: str(value) for key, value in data.items()}
        print(f"Parsing data line: {context}")

        try:
            doc.render(context)
        except PackageNotFoundError:
            print(f"\n--------------\nError: Word template file not found: {word_template_name}")
            exit()

        prepare_output_folder()

        unique_file_identifier = get_unique_file_identifier(unique_column_name, k, data)

        output_file_name = get_new_file_name(word_template_name, foldered, unique_file_identifier)

        doc.save(output_file_name)

        docx2pdf.convert(output_file_name)

        if cleanup:
            os.remove(output_file_name)
            print(f"Deleted file: {output_file_name}")

        print(f"File {output_file_name.replace('docx','pdf')} generated.\n")


def get_new_file_name(word_template_name, foldered, unique_file_identifier):
    if foldered:
        folder_name = f"{OUTPUT_FOLDER}/doc_{unique_file_identifier}"
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            print(f"Created folder: {folder_name}")
        else:
            print(f"Folder already exists: {folder_name}")
        output_file_name = f"{folder_name}/{word_template_name}.docx"
    else:
        output_file_name = f"{OUTPUT_FOLDER}/{word_template_name}_{unique_file_identifier}.docx"
    return output_file_name


def prepare_output_folder():
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)


def get_unique_file_identifier(unique_column_name, k, data):
    if unique_column_name:
        unique_file_identifier = data[str(unique_column_name)]
    else:
        unique_file_identifier = k
    return unique_file_identifier


def load_data(data_file_name):
    try:
        data = pd.read_excel(data_file_name)
        data_dict = data.to_dict("index")

        unique_columns = data.columns[data.nunique() == data.count()]

        print(f"Uniqness: {unique_columns}")
        if not unique_columns.empty:
            unique_column_name = unique_columns[0]

    except FileNotFoundError:
        print(f"\n--------------\nError: Excel data file not found: {data_file_name}")
        exit()
    print(f"Data sample: {data_dict.get(0)}")
    return data_dict, unique_column_name


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("-w", "--word", help="Name of the word template file")
    parser.add_argument("-e", "--excel", help="Name of the excel data file")
    parser.add_argument("-f", "--foldered", action="store_true", help="Output files will be saved in diferrent folders")
    parser.add_argument("-c", "--cleanup", action="store_true", help="Docx files will be deleted after conversion")
    args = parser.parse_args()

    foldered = False
    cleanup = False

    if args.word:
        print("Using Word template: % s" % args.word)
        word_template_name = args.word
    else:
        documents = glob.glob("*.docx", recursive=False)
        if len(documents) == 1:
            word_template_name = documents[0]
        else:
            print("\n--------------\nError: Excel data file not provided.")
            exit()

    if args.excel:
        print("Using Excel data: % s" % args.excel)
        data_file_name = args.excel
    else:
        excels = glob.glob("*.xlsx", recursive=False)
        if len(excels) == 1:
            data_file_name = excels[0]
        else:
            print("\n--------------\nError: Excel data file not provided.")
            exit()

    if args.foldered:
        foldered = True
        print("Output files will be saved in diferrent folders.")

    if args.cleanup:
        cleanup = True
        print("Docx files will be deleted after conversion.")

    print("\n")
    print(f"Creating documents from {data_file_name} using {word_template_name}")
    main(word_template_name, data_file_name, foldered, cleanup)

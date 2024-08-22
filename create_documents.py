import argparse
import glob
from docx.opc.exceptions import PackageNotFoundError
from docxtpl import DocxTemplate
import docx2pdf
import os
import pandas as pd

OUTPUT_FOLDER = "output"


def main(word_template_name, data_file_name, foldered, cleanup, manual):

    data_dict, unique_column_name = load_data(data_file_name, manual)

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


def get_default_file_name(suffix: str) -> str:
    documents = glob.glob(suffix, recursive=False)
    if len(documents) == 1:
        return documents[0]
    else:
        return ""


def get_possible_file_names(suffix: str) -> list:
    documents = glob.glob(suffix, recursive=False)
    return documents


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


def load_data(data_file_name, manual=False):
    try:
        data = pd.read_excel(data_file_name)
        data_dict = data.to_dict("index")

        unique_columns = data.columns[data.nunique() == data.count()]

        if not unique_columns.empty:
            unique_column_name = unique_columns[0]
        else:
            unique_column_name = ""

        if not manual:
            default_unique_column_name_yn = input(f"Can this column be used as a unique file name tag? *{unique_column_name}* (y/n): ")
            if default_unique_column_name_yn.lower() != "y":
                unique_column_name = input(
                    f"Which column should be used as unique identifier? There are those options: ({str(unique_columns)})"
                )
            else:
                unique_column_name = unique_columns[0]

    except FileNotFoundError:
        print(f"\n--------------\nError: Excel data file not found: {data_file_name}")
        exit()
    print(f"Data sample: {data_dict.get(0)}")
    return data_dict, unique_column_name


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("-m", "--manual", action="store_true", help="No interactive mode, all parameters must be entered.")
    parser.add_argument("-w", "--word", help="Name of the word template file")
    parser.add_argument("-e", "--excel", help="Name of the excel data file")
    parser.add_argument("-f", "--foldered", action="store_true", help="Output files will be saved in diferrent folders")
    parser.add_argument("-c", "--cleanup", action="store_true", help="Docx files will be deleted after conversion")
    args = parser.parse_args()

    foldered = False
    cleanup = False
    manual = False

    if not args.manual:
        print("\n--------------\nInteractive mode\n")
        print("Please provide the following parameters:")

        word_template_name = get_default_file_name("*.docx")
        default_word_template_name_yn = input(f"Do you want to use this word template file: {word_template_name} (y/n): ")
        if default_word_template_name_yn.lower() != "y":
            word_template_name = input(
                f"What is the name of the Word template file? There are those options: ({str(get_possible_file_names('*.docx'))})"
            )

        data_file_name = get_default_file_name("*.xlsx")
        default_data_file_name_yn = input(f"Do you want to use this word template file: {data_file_name} (y/n): ")
        if default_data_file_name_yn.lower() != "y":
            data_file_name = input(
                f"What is the name of the Word template file? There are those options: ({str(get_possible_file_names('*.xlsx'))})"
            )

        foldered_yn = input("Output files will be saved in diferrent folders (y/n): ")
        cleanup_yn = input("Docx files will be deleted after conversion (y/n): ")

        if foldered_yn.lower() == "y":
            foldered = True
        if cleanup_yn.lower() == "y":
            cleanup = True

    else:
        print("\n--------------\nManual mode\n")
        if args.word:
            print("Using Word template: % s" % args.word)
            word_template_name = args.word
        else:
            word_template_name = get_default_file_name("*.docx")
            if not word_template_name:
                print("\n--------------\nError: Word template file not provided.")
                exit()

        if args.word:
            print("Using Excel data file: % s" % args.excel)
            data_file_name = args.excel
        else:
            data_file_name = get_default_file_name("*.xlsx")
            if not data_file_name:
                print("\n--------------\nError: Excel data file not provided.")
                exit()

        if args.foldered:
            foldered = True
            print("Output files will be saved in diferrent folders.")

        if args.cleanup:
            cleanup = True
            print("Docx files will be deleted after conversion.")

        manual = True

    print("\n")
    print(f"Creating documents from {data_file_name} using {word_template_name}")
    main(word_template_name, data_file_name, foldered, cleanup, manual)

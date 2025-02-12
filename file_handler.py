import pandas
import os

def read_excel_data(xlsx_file, skip_rows):
    try:
        return pandas.read_excel(xlsx_file, engine="openpyxl", skiprows=[i for i in range(skip_rows)])
    except FileNotFoundError:
        print(f"Error: File '{xlsx_file}' not found.")
        exit()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        exit()


def write_snippet_file(output_directory, file_name, snippet_content):
    try:
        with open(f"{output_directory}/{file_name}", "w") as file:
            file.write(snippet_content)
    except Exception as e:
        print(f"Error writing snippet file: {e}")


def write_show_file(output_directory, show_file_name, shw_content):
     try:
         with open(f"{output_directory}/{show_file_name}.shw", "w") as file:
             file.write(shw_content)
     except Exception as e:
         print(f"Error writing show file: {e}")


def create_output_directory(output_directory):
    if output_directory not in os.listdir("."):
        os.mkdir(output_directory)


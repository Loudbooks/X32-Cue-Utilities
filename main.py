import pandas as pd
import os

def generate_snippets_from_xlsx(xlsx_file, output_directory, skip_rows, identifying_character):
    show_file_name = xlsx_file.replace(".xlsx", "")

    if (output_directory not in os.listdir(".")):
        os.mkdir(output_directory)

    data_frame = pd.read_excel(xlsx_file, engine="openpyxl", skiprows=[i for i in range(skip_rows)])

    first_integer_column = 1

    for column in data_frame.columns:
        if isinstance(column, (int, float)):
            first_integer_column = data_frame.columns.get_loc(column)
            break

    snippet_numbers = data_frame.columns[first_integer_column:]

    shw_content = '#4.0#\n'
    shw_content += f'show "{show_file_name}" 0 0 0 60 0 0 0 0 0 0 "X32-Edit 4.00"\n'
    
    snippet_list = ""

    cue_number = 0
    for snippet in snippet_numbers:
        snippet_index_formatted = str(cue_number).zfill(3)

        cue_index_formatted = ""

        if str(snippet).split(".").__len__() == 3:
            cue_index_formatted = str(snippet).replace(".", "")
        elif str(snippet).split(".").__len__() == 2:
            cue_index_formatted = (str(snippet) + ".0").replace(".", "")
        else:
            cue_index_formatted = (str(snippet) + ".0.0").replace(".", "")

        shw_content += f'cue/{snippet_index_formatted} {cue_index_formatted} "" 0 -1 {cue_number} 0 1 0 0\n'

        snippet_list += f'snippet/{snippet_index_formatted} "Q{snippet}" 128 131071 0 0 1\n'
        snippet_content = f'#4.0# "Q{snippet}" 128 131071 0 0 1\n'

        for _, row in data_frame.iterrows():
            if not row.iloc[0]:
                continue

            mic_num = int(row.iloc[0])
            unmuted = pd.notna(row[snippet]) and row[snippet] == identifying_character

            mute_state = "ON" if unmuted else "OFF"
            formatted_snippet = str(mic_num).zfill(2)

            snippet_content += f'/ch/{formatted_snippet}/mix/on {mute_state}\n'

        file_name = f"Q{snippet}.snp"
        with open(f"{output_directory}/{file_name}", "w") as file:
            file.write(snippet_content)

        print("                                                                         ", end="\r", flush=True)
        print(f"Snippet '{snippet}' saved as {file_name}", end="\r", flush=True)

        cue_number += 1

    shw_content += snippet_list
    
    with open(f"{output_directory}/{show_file_name}.shw", "w") as file:
        file.write(shw_content)

    print(f"\n{show_file_name}.shw saved")


def locate_xlsx_files():
    return [file for file in os.listdir(".") if file.endswith(".xlsx")]

xlsx_files = locate_xlsx_files()

print("")
if len(xlsx_files) == 0:
    print("No xlsx files found in the current directory.")
    exit()

if len(xlsx_files) > 1:
    print(f"Found {len(xlsx_files)} xlsx files.")

    for i, xlsx_file in enumerate(xlsx_files):
        print(f"{i + 1}. {xlsx_file}")


    print("")
    user_input = input("Choose a file to generate snippets from (1): ")
    try:
        if user_input == "":
            user_input = 1

        user_input = int(user_input)
    except ValueError:
        print("Invalid input.")
        exit()

    if user_input < 1 or user_input > len(xlsx_files):
        print("Invalid input.")
        exit()

    xlsx_file = xlsx_files[user_input - 1]
else:
    xlsx_file = xlsx_files[0]

print("")
skip_rows = input("Start at row (1): ")
try:
    if skip_rows == "":
        skip_rows = 1
    else:
        skip_rows = int(skip_rows) - 1

        if skip_rows < 0:
            skip_rows = 0
except ValueError:
    print("Invalid input.")
    exit()

print("")
identifying_character = input("Identifying character (X): ")
if identifying_character == "":
    identifying_character = "X"

generate_snippets_from_xlsx(xlsx_file, "output", skip_rows, identifying_character)
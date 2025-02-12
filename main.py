import pandas as pd
import os

def generate_snippets_from_xlsx(xlsx_file, output_directory):
    if (output_directory not in os.listdir(".")):
        os.mkdir(output_directory)

    df = pd.read_excel(xlsx_file, engine="openpyxl", skiprows=[0])

    first_integer_column = 1

    for column in df.columns:
        if isinstance(column, (int, float)):
            first_integer_column = df.columns.get_loc(column)
            break
            

    snippet_numbers = df.columns[first_integer_column:]

    for snippet in snippet_numbers:
        snp_content = f'#2.0# "Q{snippet}"\n- stat 1\n'

        for _, row in df.iterrows():
            if not row.iloc[0]:
                continue

            mic_num = int(row.iloc[0])
            unmuted = pd.notna(row[snippet])

            mute_state = 0 if unmuted else 1

            snp_content += f'- chan {mic_num}\n'
            snp_content += f'  - mix 0 {mute_state}\n'

        file_name = f"Q{snippet}.snp"
        with open(f"{output_directory}/{file_name}", "w") as file:
            file.write(snp_content)

        print(f"Snippet '{snippet}' saved as {file_name}")

def locate_xlsx_files():
    return [file for file in os.listdir(".") if file.endswith(".xlsx")]

print("Locating xlsx files...")
xlsx_files = locate_xlsx_files()

if len(xlsx_files) == 0:
    print("No xlsx files found in the current directory.")
    exit()

if len(xlsx_files) > 1:
    print("Enter the number of the xlsx file you want to use:")

    for i, xlsx_file in enumerate(xlsx_files):
        print(f"{i + 1}. {xlsx_file}")

    user_input = input()
    try:
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

print("Generating snippets...")
generate_snippets_from_xlsx(xlsx_file, "output")

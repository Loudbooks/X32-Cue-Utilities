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

def locate_xlsx_file():    
    for file in os.listdir("."):
        if file.endswith(".xlsx"):
            return file
    return None

generate_snippets_from_xlsx(locate_xlsx_file(), "snippets")

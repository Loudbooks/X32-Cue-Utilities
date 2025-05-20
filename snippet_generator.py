import sys
from file_handler import write_snippet_file, write_show_file, create_output_directory
import pandas

def generate_snippets(data_frame, show_file_name, output_directory, identifying_character, use_dca):
    create_output_directory(output_directory)

    first_integer_column = find_first_integer_column(data_frame)
    snippet_numbers = data_frame.columns[first_integer_column:]

    shw_content = generate_shw_header(show_file_name)
    snippet_list = ""

    total_cues = len(snippet_numbers)
    cue_number = 0

    for snippet in snippet_numbers:
        snippet_index_formatted = str(cue_number).zfill(3)
        cue_index_formatted = format_cue_index(snippet)

        shw_content += generate_cue_entry(snippet_index_formatted, cue_index_formatted, cue_number)
        snippet_list += generate_snippet_list_entry(snippet_index_formatted, snippet)

        snippet_content = generate_snippet_content(data_frame, snippet, identifying_character, use_dca)
        file_name = f"Q{pad_index(str(snippet))}.snp"
        write_snippet_file(output_directory, file_name, snippet_content)

        percent = round(((cue_number + 1) / total_cues) * 100)
        sys.stdout.write("\033[K") 
        print(f"Snippet '{snippet}' saved as {file_name}. ({percent}%)", end="\r", flush=True)

        cue_number += 1

    shw_content += snippet_list
    write_show_file(output_directory, show_file_name, shw_content)

    print(f"\n{show_file_name}.shw saved.")


def find_first_integer_column(data_frame):
    for column in data_frame.columns:
        if isinstance(column, (int, float)):
            return data_frame.columns.get_loc(column)
    return None


def generate_shw_header(show_file_name):
    return '#4.0#\n' + f'show "{show_file_name}" 0 0 0 60 0 0 0 0 0 0 "X32-Edit 4.00"\n'

def format_cue_index(snippet):
    parts = str(snippet).split(".")
    return "".join(parts + ["0"] * (3 - len(parts)))

def pad_index(version, width=2):
    parts = version.split('.')
    if len(parts) < 3:
        parts += ['0'] * (3 - len(parts))
    elif len(parts) > 3:
        parts = parts[:3]
        
    padded_parts = [part.zfill(width) for part in parts]
    return '.'.join(padded_parts)

def generate_cue_entry(snippet_index_formatted, cue_index_formatted, cue_number):
    return f'cue/{snippet_index_formatted} {cue_index_formatted} "" 0 -1 {cue_number} 0 1 0 0\n'

def generate_snippet_list_entry(snippet_index_formatted, snippet):
    return f'snippet/{snippet_index_formatted} "Q{pad_index(str(snippet))}" 128 131071 0 0 1\n'

def generate_snippet_content(data_frame, snippet, identifying_character, use_dca):
    snippet_content = f'#4.0# "Q{pad_index(str(snippet))}" 128 131071 0 0 1\n'
    for _, row in data_frame.iterrows():
        if not row.iloc[0]:
            continue

        mic_num = int(row.iloc[0])
        unmuted = (pandas.notna(row[snippet]) and row[snippet] == identifying_character) or (pandas.notna(row[snippet]) and row[snippet] == "-")

        isQuiet = pandas.notna(row[snippet]) and row[snippet] == "-"

        mute_state = "ON" if unmuted else "OFF"
        formatted_snippet = str(mic_num).zfill(2)

        if use_dca:
            if isQuiet:
                snippet_content += f'/ch/{formatted_snippet}/grp %00000011 %000000\n'
            else:
                snippet_content += f'/ch/{formatted_snippet}/grp %00000001 %000000\n'

            snippet_content += f'/ch/{formatted_snippet}/mix/on {mute_state}\n'
        
    return snippet_content

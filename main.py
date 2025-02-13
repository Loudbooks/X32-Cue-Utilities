from file_handler import read_excel_data
from snippet_generator import generate_snippets
from qlab_generator import generate_cues
import os

def locate_xlsx_files():
    return [file for file in os.listdir(".") if file.endswith(".xlsx")]

def get_user_input(prompt, default, validation_func=None):
    while True:
        user_input = input(prompt)
        if validation_func is None or validation_func(user_input):
            return user_input
        elif user_input == "":
            return default
        else:
            print("Invalid input. Please try again.")

def validate_int_input(input_str):
  try:
      int(input_str)
      return True
  except ValueError:
      return False

def main():
    xlsx_files = locate_xlsx_files()

    if not xlsx_files:
        print("No xlsx files found in the current directory.")
        return

    if len(xlsx_files) > 1:
        for i, xlsx_file in enumerate(xlsx_files):
            print(f"{i + 1}. {xlsx_file}")

        file_index = int(get_user_input("Choose a file to generate snippets from (1): ", "1", validate_int_input)) -1

        if 0 <= file_index < len(xlsx_files):
            xlsx_file = xlsx_files[file_index]
        else:
            print("Invalid file selection.")
            return
    else:
        xlsx_file = xlsx_files[0]

    skip_rows = int(get_user_input("Start at row (1): ", "1", validate_int_input)) - 1
    if skip_rows < 0:
        skip_rows = 0

    identifying_character = get_user_input("Identifying character (X): ", "X")
    if not identifying_character:
        identifying_character = "X"

    print(f"Choose cue generation method:")
    print("1. Generate X32 Snippets")
    print("2. Generate QLab Cues")
    method = get_user_input("Method (1): ", "1", validate_int_input)

    show_file_name = xlsx_file.replace(".xlsx", "")
    data_frame = read_excel_data(xlsx_file, skip_rows)

    if method == "1":
        generate_snippets(data_frame, show_file_name, "output", identifying_character)
    else:
        generate_cues(data_frame, identifying_character)


if __name__ == "__main__":
    main()


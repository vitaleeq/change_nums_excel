from random import choice
from string import ascii_uppercase as letters


def create_new_filename(filename):
    return filename if '/' not in filename else filename[filename.rfind('/')+1:]


def random_last_symbols(serial_number):
    quantity_symbols = -3
    if validate_string(serial_number, quantity_symbols):
        new_serial_number = serial_number[:quantity_symbols]      # -2 => -3
        for char in serial_number[quantity_symbols:]:             # -2 => -3
            if char.isnumeric():
                new_serial_number += choice('0123456789'.replace(char, ''))
            else:
                new_serial_number += choice(letters.replace(char, ''))
        return new_serial_number
    else:
        return serial_number


def create_cells_dictionary(wb_sheet, cell_names):
    column = wb_sheet[cell_names[0]].column
    start_row, finish_row = (int(''.join([i for i in elem if i.isnumeric()])) for elem in cell_names)
    cells_with_serial_numbers = {elem[0].coordinate: str(elem[1]).split(', ') for elem in ((cell, cell.value) for col \
        in wb_sheet.iter_cols(min_col=column, max_col=column, min_row=start_row, max_row=finish_row) for cell in col)}
    return cells_with_serial_numbers


def validate_string(string_vc, quantity_symbols):
    return len(string_vc) < quantity_symbols

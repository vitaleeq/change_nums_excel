from random import choice
from string import ascii_uppercase as letters


def validate_filename(file_name):
    pass


def create_new_filename(file_name):
    return file_name if '/' not in file_name else file_name[file_name.rfind('/')+1:]


def random_last_symbols(serial_number):
    new_serial_number = serial_number[:-3]
    for char in serial_number[-3:]:
        if char.isnumeric():
            new_serial_number += choice('0123456789')
        else:
            new_serial_number += choice(letters)
    return new_serial_number


def create_cells_dictionary(wb_sheet, cell_names):
    column = wb_sheet[cell_names[0]].column
    start_row, finish_row = (int(''.join([i for i in elem if i.isnumeric()])) for elem in cell_names)
    cells_with_serial_numbers = {elem[0].coordinate: str(elem[1]).split(', ') for elem in ((cell, cell.value) for col in wb_sheet.iter_cols(min_col=column, max_col=column, min_row=start_row, max_row=finish_row) for cell in col)}
    return cells_with_serial_numbers
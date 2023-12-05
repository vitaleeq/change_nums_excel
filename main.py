from openpyxl import load_workbook, Workbook
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
    cells_with_serial_numbers = {elem[0].coordinate: elem[1].split(', ') for elem in ((cell, cell.value) for col in wb_sheet.iter_cols(min_col=column, max_col=column, min_row=start_row, max_row=finish_row) for cell in col)}
    return cells_with_serial_numbers


if __name__ == "__main__":
    filename = input('Input the full filepath: ')     # /home/vitaly/Downloads/sample.xlsx
    # filename = '/home/vitaly/Downloads/test.xlsx'
    # validate filename here
    sheet_name = input('Input sheet name: ')
    # validate sheet_name here
    cell_names = list(map(str.strip, input('Input start_cell and finish_cell names, e.g. A1, A10: ').split(',')))
    # cell_names = ['Q10', 'Q124']
    # validate cell_names here

    # load file, activate sheet in it, select targeted cell's
    wb = load_workbook(filename=filename)
    wb_sheet = wb.active if sheet_name == '' else wb[sheet_name]
    cell_range = wb_sheet['Q10:Q124']

    # create dictionary with keys==cell's coords, values==cell's values
    cells_with_serial_numbers = create_cells_dictionary(wb_sheet, cell_names)

    # change serial numbers for each cell using randomiser, rewrite new serial numbers in file and save it
    for key, values in cells_with_serial_numbers.items():
        wb_sheet[key] = ', '.join([random_last_symbols(value) for value in values])

    wb.save(f"new_{create_new_filename(filename)}")

from openpyxl import load_workbook
import functional_tools as ft


# 'Спецификация Tissot Kelechek 09 24 16TIS-0524KL.xlsx'        E10:E85
if __name__ == "__main__":
    filename = input('Input the full filepath (allow only filename, if file is in a program root-directory): ')
    sheet_name = input("Input sheet name (allow to enter empty value if it's first sheet): ")
    cell_names = list(map(str.strip, input('Input start_cell and finish_cell names, e.g. A1:A10 (select target cells and copy name box): ').split(':')))

    # load file, activate sheet in it, select targeted cell's
    wb = load_workbook(filename=filename)
    wb_sheet = wb.active if sheet_name == '' else wb[sheet_name]

    # create dictionary with keys==cell's coords, values==cell's values
    cells_with_serial_numbers = ft.create_cells_dictionary(wb_sheet, cell_names)

    # change serial numbers for each cell using randomiser, rewrite new serial numbers in file and save it
    first_symbols = int(input('Сколько первых символов менять? Введи целое положительное число, не превышающее половину длины серийника: ').strip())
    last_symbols = int(input('Сколько последних символов менять? Введи целое положительное число, не превышающее половину длины серийника: ').strip())
    for key, values in cells_with_serial_numbers.items():
        wb_sheet[key] = ', '.join([ft.random_symbols(value, first_symbols, last_symbols) for value in values])

    wb.save(f"new_{ft.create_new_filename(filename)}")

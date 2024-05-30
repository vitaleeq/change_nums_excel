from openpyxl import load_workbook
import functional_tools as ft


# 'Спецификация FC Advant upd.xlsx'     E10:E44
# 'Спецификация Tissot Daghaya (2).xlsx'        F10, F51
if __name__ == "__main__":
    filename = input('Input the full filepath: ')     # /home/vitaly/Downloads/test.xlsx
    sheet_name = input('Input sheet name: ')
    cell_names = list(map(str.strip, input('Input start_cell and finish_cell names, e.g. A1, A10: ').split(',')))       # Q10:Q124

    # load file, activate sheet in it, select targeted cell's
    wb = load_workbook(filename=filename)
    wb_sheet = wb.active if sheet_name == '' else wb[sheet_name]

    # create dictionary with keys==cell's coords, values==cell's values
    cells_with_serial_numbers = ft.create_cells_dictionary(wb_sheet, cell_names)

    # change serial numbers for each cell using randomiser, rewrite new serial numbers in file and save it
    for key, values in cells_with_serial_numbers.items():
        wb_sheet[key] = ', '.join([ft.random_last_symbols(value) for value in values])

    wb.save(f"new_{ft.create_new_filename(filename)}")

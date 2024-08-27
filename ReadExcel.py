import openpyxl

def read_excel(filename):
    phones = []

    try:
        workbook = openpyxl.load_workbook(filename)
        for sheet in workbook.worksheets:
            for column in range(1, sheet.max_column + 1):
                if (str(sheet.cell(1, column).value).lower() == "телефон"):
                    for row in range(2, sheet.max_row + 1):
                        phones.append(sheet.cell(row, column).value)
                    break
    except Exception as e:
        print(e)

    return  phones
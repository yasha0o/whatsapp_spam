import logging
import openpyxl

def read_excel(filename):
    logging.info(f"read_excel filename = {filename}")
    phones = []

    try:
        workbook = openpyxl.load_workbook(filename)
        logging.info(f"workbook = {workbook}")
        for sheet in workbook.worksheets:
            logging.info(f"sheet = {sheet}")
            for column in range(1, sheet.max_column + 1):
                logging.debug(f"title cell = {sheet.cell(1, column).value}")
                if (str(sheet.cell(1, column).value).lower() == "телефон"):
                    for row in range(2, sheet.max_row + 1):
                        logging.debug(f"value cell = {sheet.cell(row, column).value}")
                        phones.append(sheet.cell(row, column).value)
                    break
    except Exception as e:
        logging.error("Exception",exc_info=True)

    return  phones
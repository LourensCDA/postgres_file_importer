import re, sys, logging, openpyxl, os, dotenv
from datetime import datetime

# function to format dictionary key
def fmt_key(key_name):
    # add leading " " before each capital
    re.sub(r"(\w)([A-Z])", r"\1 \2", "key_name")
    # replace spaces with underscores
    key_name = key_name.lower().replace(" ", "_")
    # remove leading and trailling spaces
    key_name = key_name.strip()
    # remove non alphanumeric characters
    key_name = re.sub(r"[^a-zA-Z0-9_]", "", key_name)
    # return value
    return key_name


# load file data
def load_file(var_loc):

    # open workbook, read only property allows for more consistent and faster read times esecially when working on larger excel files, data_only shows calculated values of formuals
    wb = openpyxl.load_workbook(var_loc, read_only=True, data_only=True)

    # set active sheet
    sheet = wb.active

    # get list of column names
    key_names = []
    for x in range(1, sheet.max_column + 1):
        if sheet.cell(1, x).value:
            logging.debug("Column name: " + sheet.cell(1, x).value)
            key_names.append(
                fmt_key(sheet.cell(1, x).value)
            )  # column names start at row 4

    # convert rows to dictionary
    # iterate over rows

    row_num = 2  # specify where rows starts at row 2
    max_rw = 10  # for testing, comment out for production
    mx_row = int(os.getenv("MAX_ROW"))
    data = []

    for row in sheet.iter_rows(min_row=row_num, values_only=True):
        res = {}
        logging.info(f"{row_num} of {mx_row}")
        # set dictinary key:value for each column
        res = {key_names[i]: row[i] for i in range(len(key_names))}

        # custom logic START

        # custom logic END

        # for testing
        if max_rw > 0 and row_num == max_rw:
            break

        row_num += 1

    return data


if __name__ == "__main__":
    dotenv.load_dotenv(verbose=True)

    # set logging level
    # debug, info, warning, error and critical
    if os.getenv("LOG_LEVEL") == "debug":
        log_lvl = logging.DEBUG
    elif os.getenv("LOG_LEVEL") == "info":
        log_lvl = logging.INFO
    elif os.getenv("LOG_LEVEL") == "warning":
        log_lvl = logging.WARNING
    elif os.getenv("LOG_LEVEL") == "error":
        log_lvl = logging.ERROR
    elif os.getenv("LOG_LEVEL") == "critical":
        log_lvl = logging.CRITICAL
    else:
        log_lvl = logging.INFO

    logging.basicConfig(stream=sys.stderr, level=log_lvl)

    try:
        file_data = load_file("Outbound Calls Data - 29 June 2022.xlsx")
        logging.debug(file_data[0])

    except Exception as e:
        logging.error("Error loading file")
        logging.debug(e)

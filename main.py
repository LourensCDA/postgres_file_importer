import re, sys, logging, openpyxl, os, dotenv, psycopg2
from datetime import datetime

#   @desc   function to format dictionary key
#   @author Lourens Botha
#   @date   2022-07-04
def fmt_key(key_name):
    if "_" not in key_name:
        # add leading " " before each capital
        key_name = re.sub(r"(\w)([A-Z])", r"\1 \2", key_name)
    # replace spaces with underscores
    key_name = key_name.lower().replace(" ", "_")
    # remove leading and trailling spaces
    key_name = key_name.strip()
    # remove non alphanumeric characters
    key_name = re.sub(r"[^a-zA-Z0-9_]", "", key_name)
    # return value
    return key_name


#   @desc   function to validate e-mail address structure
#   @author Lourens Botha
#   @date   2022-07-04
def valid_email(email):

    regex = re.compile(
        r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])"
    )

    if re.search(regex, email):
        return True
    return False


#   @desc   load excel file data into array of objects
#   @author Lourens Botha
#   @date   2022-07-04
def load_file(var_loc):

    # open workbook, read only property allows for more consistent and faster read times esecially when working on larger excel files, data_only shows calculated values of formuals
    wb = openpyxl.load_workbook(var_loc, read_only=True, data_only=True)

    # set active sheet
    sheet = wb.active

    # get list of column names
    key_names = []
    for x in range(1, sheet.max_column + 1):
        if sheet.cell(1, x).value:
            custom_value = None

            # Custom key name logic start

            if sheet.cell(1, x).value.upper() == "ID_NO":
                custom_value = "id_no"

            # Customer key name logic end

            if custom_value:
                key_names.append(custom_value)
            else:
                key_names.append(
                    fmt_key(sheet.cell(1, x).value)
                )  # column names start at row 1, column 1

    row_num = 2  # specify where rows starts at row 2
    max_rw = 0  # for testing, comment out for production
    mx_row = sheet.max_row
    data = []

    # convert rows to dictionary, iterate over rows
    for row in sheet.iter_rows(min_row=row_num, values_only=True):
        res = {}
        logging.info(f"{row_num} of {mx_row}")
        # set dictinary key:value for each column
        res = {key_names[i]: row[i] for i in range(len(key_names))}

        # custom logic START

        res["postal_code"] = (
            ("0000" + str(res["postal_code"]))[-4:] if res["postal_code"] else None
        )

        if res["postal_city"] == "-":
            res["postal_city"] = None

        if res["email_address"] and not (valid_email(res["email_address"])):
            logging.warning(f"Invalid email address: {res['email_address']}")
            res["email_address"] = None

        # custom logic END

        # append dictionary to list
        data.append(res)

        # for testing
        if max_rw > 0 and row_num == max_rw:
            break

        row_num += 1

    return data


#   @desc   insert data array of objects to table
#   @author Lourens Botha
#   @date   2022-07-04
def insert_data_many(qry, args=None):
    outcome = False
    try:
        # try to connect
        conn = psycopg2.connect(
            dbname=os.getenv("DB_NAME"),
            user=os.getenv("DB_USER"),
            host=os.getenv("DB_HOST"),
            password=os.getenv("DB_PASSWORD"),
            port=os.getenv("DB_PORT"),
        )
        cursor = conn.cursor()
        cursor.executemany(qry, args)
        conn.commit()
        outcome = True
    except Exception as error:
        logging.error(f"{error}")
    finally:
        if cursor:
            cursor.close()
    # return true if all ran good and false if not
    return outcome


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
        # specify file location
        file_data = load_file("files/fileName.xlsx")
        if file_data:
            logging.debug(file_data[0])

            # custom insert statement start

            insert_data_many(
                "INSERT INTO schema.table(column1, column2) values (%(column1)s, %(column2)s);",
                file_data,
            )

            # customer insert statement end

    except Exception as e:
        logging.error("Error loading file")
        logging.debug(e)

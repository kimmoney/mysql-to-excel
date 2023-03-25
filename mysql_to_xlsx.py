import xlsxwriter
import pymysql
import datetime
import os
login = []
db = {"gongzone_db":"db","gongzone_log":"log"}
try:
    f = open("C:/login",'r')
except:
    f = open("/Users/hoon/Documents/test", 'r')
for line in f:
    line = line.strip()
    login.append(line)
f.close()

mysql = pymysql.connect(host=login[0], 
                      user=login[1], password=login[2],
                       db=login[3], charset='utf8') # 한글처리 (charset = 'utf8')
def fetch_table_data(table_name):
    # The connect() constructor creates a connection to the MySQL server and returns a MySQLConnection object.
    cnx = pymysql.connect(
        host=login[0], 
                      user=login[1], password=login[2],
                       db=login[3], charset='utf8'
    )

    cursor = cnx.cursor()
    cursor.execute("SELECT * FROM "+table_name)

    header = [row[0] for row in cursor.description]

    rows = cursor.fetchall()

    # Closing connection
    cnx.close()

    return header, rows


def export():
    # Create an new Excel file and add a worksheet.
    try: os.mkdir('data/')
    except:pass
    workbook = xlsxwriter.Workbook("data/gongzone"+"_"+datetime.datetime.now().strftime('%Y_%m_%d-%H_%M_%S') + '.xlsx')

    for table_name in db:
        worksheet = workbook.add_worksheet(db[table_name])

        # Create style for cells
        header_cell_format = workbook.add_format({'bold': True, 'border': True, 'bg_color': 'yellow'})
        body_cell_format = workbook.add_format({'border': True})

        header, rows = fetch_table_data(table_name)

        row_index = 0
        column_index = 0

        for column_name in header:
            worksheet.write(row_index, column_index, column_name, header_cell_format)
            column_index += 1

        row_index += 1
        for row in rows:
            column_index = 0
            for column in row:
                worksheet.write(row_index, column_index, column, body_cell_format)
                column_index += 1
            row_index += 1

        print(str(row_index) + ' rows written successfully to ' + workbook.filename)

        # Closing workbook
    workbook.close()


# Tables to be exported
export()
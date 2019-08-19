# importing module
import cx_Oracle
import xlsxwriter
from xlsxwriter import Workbook
import argparse, pathlib
from config import config

cursor = ""
con = ""
data = ""
employee = {}

# Get Arguments
parser = argparse.ArgumentParser(description='UDS Remittance Calculator')
parser.add_argument('--output_type', help='Output Type, Values can be excel/text/database', required=True)
parser.add_argument('--output_file', help='File name when the Output type is excel/text', required=False)
parser.add_argument('--table_name', help='Table name when the Output type is database', required=False)

args = vars(parser.parse_args())

if args['output_type'] in ['excel', 'text']:
    if not args['output_file']:
        print("Missing output_file")
        parser.print_help()
        quit()
    else:
        filename = pathlib.Path(args['output_file'])
        if filename.exists():
            print("File already exists " + args['output_file'])
            quit()
elif args['output_type'] == 'database':
    if not args['table_name']:
        print("Missing table_name")
        parser.print_help()
        quit()

# Functions
def create_table(t_name, cursor):
    print("Creating table = " + t_name)
    query = 'CREATE TABLE ' + t_name + ' \
    (id NUMBER, \
    name VARCHAR2(20), \
    age NUMBER, \
    notes VARCHAR2(100))'
    cursor.execute(query)
    con.commit()

def drop_table(t_name, cursor):
    print("dropping table = " + t_name)

    cursor.execute( """
    SELECT table_name FROM user_tables where table_name = :myvar
    """, myvar=t_name)
    result = cursor.fetchall()
    #print(result)
    for row in result:
       print(row)
    if result:
        query = 'drop TABLE ' + t_name
        cursor.execute(query)
        con.commit()

def write_table(t_name, cursor, data):
    print("Inserting Rows in " + t_name)
    cursor.setinputsizes(int, 20, int, 100)
    cursor.executemany("insert into " + t_name + "(id, name, age, notes) values (:1, :2, :3, :4)", data)
    con.commit()

def fetch_data(t_name, cur):
    print("Fetching Rows from " + t_name)
    query = "SELECT * FROM " + t_name
    cur.execute(query)
    result = cur.fetchall()
    for row in result:
        #print(row)
        dict = {
            row[0] : {
                "name" : row[1],
                "age": row[2],
                "notes": row[3]
            }
        }
        employee.update(dict)
    return employee

def write_excel(f_name, data):
    print("Writing to Excel file..")
    row = 0

    wb = Workbook(f_name)
    sheet1 = wb.add_worksheet('Sheet 1')
    bold = wb.add_format({'bold': True, 'bg_color': 'silver'})

    # Add header
    for x in data.keys():
        column = 0
        for y in data[x].keys():
            sheet1.write(row, column, y, bold)
            column+=1
        break
    row+=1
    # Add data
    for x in data.keys():
        column = 0
        for y in data[x].keys():
            #sheet1.write(row, column, x)
            sheet1.write(row, column, data[x][y])
            column+=1
            print(x, "-", y, "-", data[x][y], row, column)
        row+=1
    wb.close()

mydict =     {
       "empl-1": {
            "name": "prem",
            "age": 33,
            "ESI": 120,
            "PF": 500
            },

       "empl-2": {
            "name": "ram",
            "age": 33,
            "ESI": 100,
            "PF": 450
            },

       "empl-3": {
            "name": "kumar",
            "age": 33,
            "ESI": 110,
            "PF": 400
            }
    }
write_excel(args['output_file'], mydict)

"""
try:
    user, password, host, port, database = [config[k] for k in ('user', 'password','host', 'port', 'database')]
    dsn_tns = cx_Oracle.makedsn(host, port, service_name=database)
    con = cx_Oracle.connect(user=user, password=password, dsn=dsn_tns)
    cursor = con.cursor()

    # Drop table
    drop_table("EMPLOYEE", cursor)

    # Create table
    create_table("employee", cursor)

    # Write data
    rows = [(3, 'rob', 35, 'I like dogs'), (4, 'jim', 27, 'I like birds')]
    write_table("employee", cursor, rows)

    # Read data
    fetched = fetch_data("employee", cursor)
    for x in fetched.keys():
        print("\nNew Key")
        for y in fetched[x].keys():
            print(x, "-", y, "-", fetched[x][y])

    # Write data to Excel
    write_excel(args['output_file'], fetched)

except cx_Oracle.DatabaseError as e:
    print("There is a problem with Oracle", e)

# by writing finally if any error occurs
# then also we can close the all database operation
finally:
    if cursor:
        cursor.close()
    if con:
        con.close()
"""

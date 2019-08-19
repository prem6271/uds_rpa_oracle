# importing module
import cx_Oracle
import xlwt
from xlwt import Workbook
import argparse
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
elif args['output_type'] == 'database':
    if not args['table_name']:
        print("Missing table_name")
        parser.print_help()

# Functions
def create_table(t_name, cursor):
    print("Create table Called.." + t_name)
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
    print("Fetching Rows " + t_name)
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

try:
    #con = cx_Oracle.connect("user/password@server/ServiceName")
    user, password, host, port, database = [config[k] for k in ('user', 'password','host', 'port', 'database')]
    dsn_tns = cx_Oracle.makedsn(host, port, service_name=database)
    con = cx_Oracle.connect(user=user, password=password, dsn=dsn_tns)

    #con = cx_Oracle.connect(user/password@host:port/database)

    #con = cx_Oracle.connect('admin/admin1234@mydbinstance.ctnch1lu7rxo.us-east-1.rds.amazonaws.com:1521/orcl')
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

#for x in mydict:
#   val = mydict[x]
#   print(x, val)


#for x in mydict.keys():
#    for y in mydict[x].keys():
#       print(x, "-", y)

#for x in mydict.keys():
#    for y in mydict[x].keys():
#           print(x, "-", y, "-", mydict[x][y])


if 'empl-4' not in mydict.keys():
    mydict['empl-4'] =  {
            "name": "rames",
            "age": 33,
            "ESI": 110,
            "PF": 490
            }


for key, val in mydict.items():
    for y in val.keys():
        print("Key", key, y, mydict[key][y])

"""

"""


# Program to show various ways to read and
# write data in a file.
file1 = open("myfile.txt","w")
L = ["This is Delhi \n","This is Paris \n","This is London \n"]

# \n is placed to indicate EOL (End of Line)
file1.write("Hello \n")
for x in L:
    file1.writelines(x)
file1.close() #to change file access modes
"""





"""

# xlwt is for writing; xlrd is for reading both are for old excel versions

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')

# row, column
sheet1.write(1, 0, 'prem')
sheet1.write(2, 0, 'kumar')

sheet1.write(0, 1, 'ISBT DEHRADUN')
sheet1.write(0, 2, 'SHASTRADHARA')

wb.save('xlwt example.xls')

"""

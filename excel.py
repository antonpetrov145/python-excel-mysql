import openpyxl
import MySQLdb

db = MySQLdb.connect(host='host', user='user', db='testdb', passwd='passwd')
cursor = db.cursor()

wb = openpyxl.load_workbook('file.xlsx')
ws = wb['sheet']

rows = []

for row in ws.iter_rows(min_row=1, min_col=1, values_only=True):
    number = ws.cell(1,1).value
    chip = ws.cell(1,3).value
    rows.append((number, chip))

db = MySQLdb.connect(host='host', user='user', db='testdb', passwd='passwd')
cursor = db.cursor()
cursor.executemany("""INSERT INTO table (col1,col2) VALUES (%s,%s)""", (rows))

db.commit()
db.close()

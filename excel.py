import openpyxl
import MySQLdb

db = MySQLdb.connect(host='host', user='user', db='testdb', passwd='passwd')
cursor = db.cursor()

wb = openpyxl.load_workbook('file.xlsx')
ws = wb['sheet']

for i in range(2, ws.max_row+1):
    row = [cell.value for cell in ws[i]]
    #print (row)
    cursor.execute("""INSERT INTO table (num,chip) VALUES (%s,%s)""", (str(row[1]), str(row[2]), ))

db.commit()
db.close()

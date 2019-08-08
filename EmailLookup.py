import pyodbc

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=PC\SQLEXPRESS;'
                      'Database=Scripting;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()

cursor.execute('SELECT * FROM Scripting.dbo.EmailLookup')

customers = []
emails = []
lookupDict = {}

for row in cursor:
    customers.append(row[0])
    emails.append(row[1])

lookupDict = dict(zip(emails, customers))
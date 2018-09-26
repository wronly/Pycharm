from dbfpy import dbf
table=dbf.DBF('C:\\people.dbf')

for record in table:
    print(record)
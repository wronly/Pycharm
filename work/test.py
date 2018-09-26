from collections import namedtuple
from dbfread import DBF

table = DBF('c:\\people.dbf', lowernames=True)

# Set record factory. This must be done after
# the table is opened because it needs the field
# names.
Record = namedtuple('Record', table.field_names)
factory = lambda lst: Record(**dict(lst))
table.recfactory = factory

for record in table:
    print(record.name)

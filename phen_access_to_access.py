import pyodbc
import pandas as pd
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

# grab the datasources we have access to
# pyodbc.dataSources()

# define components of our connection string
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\coron\OneDrive\Atom\Python\Webinars.accdb;'
)

# create a connection to the database
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
for table_info in crsr.tables(tableType='TABLE'):
    print(table_info.table_name)

# define the components of a query
table_name = 'Automation'

# define query
query = 'SELECT * FROM {}'.format(table_name)

data = pd.read_sql(query, cnxn)

print(data)


qry = data.query('FreeItemCode == "WBRLCCOVID20"')
print(qry)

data.head()


data['FreeItemCode']

webinar_title = qry.loc[[42], 'WebinarTitle']
print(webinar_title)

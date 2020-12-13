import win32com.client as client
import pyodbc
import pandas as pd
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

# Access section: Retrieve variable values

# define components of our connection string
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\coron\OneDrive\Atom\Python\Webinars.accdb;'
)

# create a connection to the database
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()

# define the components of a query
table_name = 'Automation'

# define query
query = 'SELECT * FROM {}'.format(table_name)

# define dataframe
data = pd.read_sql(query, cnxn, index_col="FreeItemCode")

# define selected FreeItemCode

webinar_identifier = "WBRLCCOVID20"

# Check data
print(data.columns)
# print(data.head())

# define variable values

webinartitle = data.loc[webinar_identifier, "WebinarTitle"]
presenter = data.loc[webinar_identifier, "Presenter"]
date = data.loc[webinar_identifier, "Date"]
time = data.loc[webinar_identifier, "Time"]
industrytechnique = data.loc[webinar_identifier, "IndustryTechnique"]
registration_sales = data.loc[webinar_identifier, "Registration Sales"]
whoshouldattend = data.loc[webinar_identifier, "WhoShouldAttend"]

print(webinartitle, presenter, date, time, industrytechnique, registration_sales, whoshouldattend)

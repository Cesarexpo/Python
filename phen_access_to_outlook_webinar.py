# Early version. Do not use.

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
data = pd.read_sql(query, cnxn)

print(data)


qry = data.query('FreeItemCode == "WBRENVPFAS20"')
print(qry)

data.head()


data['FreeItemCode']

webinartitle = qry.loc[[42], 'WebinarTitle'].to_string(index=False)
presenter = qry.loc[[42], 'Presenter'].to_string(index=False)
date = qry.loc[[42], 'Date'].to_string(index=False)
time = qry.loc[[42], 'Time'].to_string(index=False)
industrytechnique = qry.loc[[42], 'IndustryTechnique'].to_string(index=False)
registration_sales = qry.loc[[42], 'Registration Sales'].to_string(index=False)
whoshouldattend = qry.loc[[42], 'WhoShouldAttend'].to_string(index=False)
print(webinartitle)


# Outlook section
outlook = client.Dispatch("Outlook.Application")

message = outlook.CreateItem(0)
whoshouldattend = "Anyone"
learningobjectives = "You'll learm about HPLC"

message.Display()
message.to = "coronacorral@gmail.com"
message.CC = "ensaladacerveza@gmail.com"
message.BCC = "cesarcorona@cesarcorona.com"
message.Subject = "INVITE YOUR CUSTOMERS - Free Webinar: {webinartitle}".format(
    webinartitle=webinartitle)

html_body = """
    <html>
      <head></head>
      <body>
      <pstyle = font-family: Calibri; font-size: 14>Hi everyone,&nbsp;</p>
      <br></br>
      <pstyle = font-family: Calibri; font-size: 14>On {date}, at {time}, {presenter} will be hosting an {industrytechnique} webinar to North American customers:&nbsp;</p>
      <p><span style="font-family: 'Lucida Sans'; font-weight: bold; font-size: 28;">
        {webinartitle}&nbsp;
        </span></p>
        <pstyle = font-family: Calibri; font-size: 14>This is a great excuse to CALL your customers!&nbsp;</p>
         <br></br>
      <pstyle = font-family: Calibri; font-size: 14>Tell them to sign up at: {registration_sales}&nbsp;</p>
       <br></br>
      <pstyle = font-family: Calibri; font-size: 14>If they can't attend, tell them to register anyway: all registrants will get access to a recording of the webinar later.&nbsp;</p>
    	  <p>WHO SHOULD ATTEND:&nbsp;</p>
      <pstyle = font-family: Calibri; font-size: 14>• {whoshouldattend}&nbsp;</p>
       <br></br>
      <pstyle = font-family: Calibri; font-size: 14>OVERVIEW/LEARNING OBJECTIVES:&nbsp;</p>
      <pstyle = font-family: Calibri; font-size: 14>• {learningobjectives}&nbsp;</p>
         <br></br>
    <pstyle = font-family: Calibri; font-size: 14>Best, &nbsp;</p>
    <br></br>
    <pstyle = font-family: Calibri; font-size: 14>César &nbsp;</p>
    	</p>
      </body>
    </html>""".format(date=date, time=time, presenter=presenter, industrytechnique=industrytechnique, webinartitle=webinartitle, registration_sales=registration_sales, whoshouldattend=whoshouldattend, learningobjectives=learningobjectives)

message.HTMLBody = html_body

message.Save()
# message.Send()

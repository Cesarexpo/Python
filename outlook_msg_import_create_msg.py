import win32com.client as client
import pandas as pd

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)

df = pd.read_excel(r'C:\Users\coron\OneDrive\Atom\Python\For_Python.xlsx')
# print(df)

# Load variables from Excel spreadsheet
date = df["Date"][0]
time = df["Time"][0]
presenter = df["Presenter"][0]
industrytechnique = df["IndustryTechnique"][0]
webinartitle = df["WebinarTitle"][0]
registration_sales = df["Registration Sales"][0]
whoshouldattend = df["WhoShouldAttend"][0]
learningobjectives = df["Overview"][0]

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
</html>
""".format(date=date, time=time, presenter=presenter, industrytechnique=industrytechnique, webinartitle=webinartitle, registration_sales=registration_sales, whoshouldattend=whoshouldattend, learningobjectives=learningobjectives)

message.HTMLBody = html_body

message.Save()
# message.Send()

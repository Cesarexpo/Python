# Early version. Do not use.

import win32com.client as client

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
webinar_title = "COVID-19"
presenter = "Phil Koerner, PhD"

message.Display()
message.to = "coronacorral@gmail.com"
message.CC = "ensaladacerveza@gmail.com"
message.BCC = "cesarcorona@cesarcorona.com"
message.Subject = "Invite your customers: Webinar {webinar_title}".format(
    webinar_title=webinar_title)

# Message body in HTML format

html = """\
<html>
  <head></head>
  <body>
    <p>This week.<br>
       Invite your customers to this free webinar:<br>
       <br><br><h1>{webinar_title}</h1><br>
       <img src="https://images.squarespace-cdn.com/content/v1/58733716d1758e4368e000f2/1537756791469-RYHRI4W752VGHSYLUM8U/ke17ZwdGBToddI8pDm48kFQQgP34qnCpeHaeAOzTt7pZw-zPPgdn4jUwVcJE1ZvWQUxwkmyExglNqGp0IvTJZamWLI2zvYWH8K3-s_4yszcp2ryTI0HqTOaaUohrI8PICHnXC1b9smDvYLPdL-DS7U1pkhCtl83kemXd5r3C5ngKMshLAGzx4R3EDFOm1kBS/20100719+-+Cesar+Corona+-+_DSC0005.jpg?format=750w">
    </p>
  </body>
</html>
""".format(webinar_title=webinar_title)


html_body = """
    <div>
        <h1 style="font-family: 'Aharoni'; font-size: 56; font-weight: bold; color: #9eac9c;">
            Please watch our webinar {webinar_title} by {presenter}
        </h1>
        <span style="font-family: 'Aharoni'; font-size: 28; color: #8d395c;">
            Please watch our webinar {webinar_title} by {presenter}
        </span>
    </div><br>
    <div>
        <img src="https://hips.hearstapps.com/hmg-prod.s3.amazonaws.com/images/cute-birthday-instagram-captions-1584723902.jpg" width=50%>
    </div>
    """.format(webinar_title=webinar_title, presenter=presenter)
message.HTMLBody = html + html_body

message.Save()
# message.Send()

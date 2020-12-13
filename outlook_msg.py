import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.to = "coronacorral@gmail.com"
message.CC = "ensaladacerveza@gmail.com"
message.BCC = "cesarcorona@cesarcorona.com"
message.Subject = "This was created with Python"
message.Body = "Hello, this is a test"
message.HTMLBody = "<b>HTML text 1</b>"
html_body = """
    <div>
        <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;">
            Happy Birthday!!
        </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;">
            Wishing you all the best on your birthday!!
        </span>
    </div><br>
    <div>
        <img src="https://hips.hearstapps.com/hmg-prod.s3.amazonaws.com/images/cute-birthday-instagram-captions-1584723902.jpg" width=50%>
    </div>
    """
message.HTMLBody = html_body
message.Save()
# message.Send()

import win32com.client as client


free_item_code = "WBRRPMORPH20"
webinar_title = "Title"

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.to = "nicolee@phenomenex.com"
message.BCC = "cesarc@phenomenex.com"
message.Subject = "Request to add webinar to Featured Webinars page"
html_body = """
    <div>
        <span style="font-family: 'Arial'; font-size: 12; color: #000000;">
            Dear Data Integrity Team,<br>
            <br>
            Could you please add the following webinar to the Featured Webinars page?<br>
            <br>
            Thanks,<br>
            <br>
            César
        </span>
    </div>
    """
message.HTMLBody = html_body
message.Save()

# Operational version 20200917
import win32com.client as client

free_item_code = "WBRGCTRBS20"


def submit_blazing_hot():

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "Blazing Hot - Webinar - " + free_item_code
    message.Attachments.Add(r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - Blazing Hot.xlsx")
    html_body = """
        <div>
            <span style="font-family: 'Arial'; font-size: 12; color: #000000;">
                Dear Data Integrity Team,<br>
                <br>
                Could you please code the leads/contacts included in the attached spreadsheet?<br>
                <br>
                Thanks,<br>
                <br>
                César
            </span>
        </div>
        """
    message.HTMLBody = html_body
    message.Save()


def submit_toasty_warm():

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "Toasty Warm - Webinar - " + free_item_code
    message.Attachments.Add(r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - Toasty Warm.xlsx")
    html_body = """
        <div>
            <span style="font-family: 'Arial'; font-size: 12; color: #000000;">
                Dear Data Integrity Team,<br>
                <br>
                Could you please code the leads/contacts included in the attached spreadsheet?<br>
                <br>
                Thanks,<br>
                <br>
                César
            </span>
        </div>
        """
    message.HTMLBody = html_body
    message.Save()


def submit_new_project():

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "New Project - Webinar - " + free_item_code
    message.Attachments.Add(r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - New Project.xlsx")
    html_body = """
        <div>
            <span style="font-family: 'Arial'; font-size: 12; color: #000000;">
                Dear Data Integrity Team,<br>
                <br>
                Could you please code the leads/contacts included in the attached spreadsheet?<br>
                <br>
                Thanks,<br>
                <br>
                César
            </span>
        </div>
        """
    message.HTMLBody = html_body
    message.Save()

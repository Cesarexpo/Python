# Operational version 20201213
import win32com.client as client


def submit_blazing_hot(webinar_identifier):

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "Blazing Hot - Webinar - " + webinar_identifier
    message.Attachments.Add(
        r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - Blazing Hot.xlsx")
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


def submit_toasty_warm(webinar_identifier):

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "Toasty Warm - Webinar - " + webinar_identifier
    message.Attachments.Add(
        r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - Toasty Warm.xlsx")
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


def submit_new_project(webinar_identifier):

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()
    message.to = "helpdeskdataintegrity@phenomenex.com"
    message.BCC = "cesarc@phenomenex.com"
    message.Subject = "New Project - Webinar - " + webinar_identifier
    message.Attachments.Add(
        r"C:\Users\coron\OneDrive\Atom\Python\Project Charter - New Project.xlsx")
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

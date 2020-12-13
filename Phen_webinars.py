import win32com.client as client
import pyodbc
import pandas as pd
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]


class webinar:

    def __init__(self, webinar_identifier):
        self._webinar_identifier = webinar_identifier

    def retrieve_var(self):
        # MS Access section: Retrieve variable values
        # define components of the connection string
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
        # webinar_identifier = "WBRLCCOVID20"

        # Check data
        # print(data.columns)
        # print(data.head())

        # define variable values

        webinartitle = data.loc[webinar_identifier, "WebinarTitle"]
        brc = data.loc[webinar_identifier, "BRC"]
        # freeitemcode = data.loc[webinar_identifier, "FreeItemCode"]
        presenterfirstname = data.loc[webinar_identifier, "PresenterFirstName"]
        presenter = data.loc[webinar_identifier, "Presenter"]
        date = data.loc[webinar_identifier, "Date"]
        time = data.loc[webinar_identifier, "Time"]
        industrytechnique = data.loc[webinar_identifier, "IndustryTechnique"]
        registration_sales = data.loc[webinar_identifier, "Registration Sales"]
        overview = data.loc[webinar_identifier, "Overview"]
        registrantstotal = data.loc[webinar_identifier, "Registrants Total"]
        attendees = data.loc[webinar_identifier, "Attendees"]
        whoshouldattend = data.loc[webinar_identifier, "WhoShouldAttend"]
        registration_sales = data.loc[webinar_identifier, "Registration Sales"]
        learning_objectives = data.loc[webinar_identifier, "LearningObjectives"]

        # print(webinartitle, presenter, date, time, industrytechnique, registration_sales, whoshouldattend, brc, webinar_identifier, overview, registrantstotal, attendees)

        return webinartitle, brc, webinar_identifier, presenterfirstname, presenter, date, time, industrytechnique, registration_sales, overview, registrantstotal, attendees, whoshouldattend, registration_sales, learning_objectives

        # Test return
        # print(retrieve_var(webinar_identifier))

    def set_var(self):
        pass

    def email_sales(self):
        pass

    def email_invitation(self):
        pass


def main():
    pass

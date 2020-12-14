# Operational version 20201004
import pyodbc
import pandas as pd
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]


def phen_var(webinar_identifier):

    # Access section: Retrieve variable values

    # define components of our connection string
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\coron\OneDrive\Atom\Python\Webinars.accdb;'
    )

    # create a connection to the database
    cnxn = pyodbc.connect(conn_str)

    # define the components of a query
    table_name = 'Automation'

    # define query
    query = 'SELECT * FROM {}'.format(table_name)

    # define dataframe
    data = pd.read_sql(query, cnxn, index_col="FreeItemCode")

    # define selected FreeItemCode

    # Check data
    print(data.columns)
    print(data.head())

    # define variable values

    webinartitle = data.loc[webinar_identifier, "WebinarTitle"]
    brc = data.loc[webinar_identifier, "BRC"]
    # freeitemcode = data.loc[webinar_identifier, "FreeItemCode"]
    presenterfirstname = data.loc[webinar_identifier, "PresenterFirstName"]
    presenter = data.loc[webinar_identifier, "Presenter"]
    date = data.loc[webinar_identifier, "Date"]
    time = data.loc[webinar_identifier, "Time"]
    industry = data.loc[webinar_identifier, "Industry"]
    technique = data.loc[webinar_identifier, "Technique"]
    # registration_sales = data.loc[webinar_identifier, "Registration Sales"]
    overview = data.loc[webinar_identifier, "Overview"]
    registrantstotal = data.loc[webinar_identifier, "Registrants Total"]
    attendees = data.loc[webinar_identifier, "Attendees"]
    whoshouldattend = data.loc[webinar_identifier, "WhoShouldAttend"]
    learning_objectives = data.loc[webinar_identifier, "LearningObjectives"]
    print(webinartitle, presenter, date, time, whoshouldattend, brc,
          webinar_identifier, overview, registrantstotal, attendees)
    return webinartitle, presenter, date, time, whoshouldattend, brc, webinar_identifier, overview, registrantstotal, attendees

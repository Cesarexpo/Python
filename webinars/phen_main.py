# Operational version 20201213
from webinarDataCharter import *
from data_entry_charters_and_email_draft import *
from phen_variables import *
# from webinarDataCharter import *

free_item_code = "WBRBIONP18"

# Create data charters
createSpreadsheets(free_item_code)

# Create Blazing Hot email with attachment
submit_blazing_hot(free_item_code)

# Create Toasty Warm email with attachment
submit_toasty_warm(free_item_code)

# Create New Project email with attachment
submit_new_project(free_item_code)

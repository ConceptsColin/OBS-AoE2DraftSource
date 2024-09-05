#google authentication: https://github.com/burnash/gspread/blob/master/docs/oauth2.rst
#readme: Coming Soon

import os.path
import obspython as obs
import gspread
import requests
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#manual script run updates the Google Sheets Table and return
#user types in search criteria and clicks the update_filter button, the format table function is run.
#user types in civ_draft and map_draft numbers and clicks link_update button, the links are updated

def script_description():
  return '''Using Google Sheets and Captains Mode, save 30 drafts and return all the unique drafts that have been pulled. Then choose the draft to update the browser source URL based on the listed row number.'''

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

get_creds_token()
#provide the current drafts using unique values returned by update_draft_data
row_one_single = ["Title        ", "Draft ID", "Preset ID", "Ongoing", "Host", "Guest"]
unformated_draft = update_draft_data()
cur_draft = format_table(unformated_draft, row_one_single) #string formatted from nested list
#cur_draft = format_table(filter_table(unique_values, search_filter), row_one_single)

#press button to update row_listings with search_filter
#can I look through "Settings" and get the value of filter_text as needed?


#Go through Google Sheets and Google Cloud authorization process.
def get_creds_token():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("Your/Path/token.json"):
    creds = Credentials.from_authorized_user_file("Your/Path/token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "Your/Path/Oauth_API_MapCiv_Draft.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("Your/Path/token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)
    print(creds)

  except HttpError as err:
    print(err)

#Pull data from Captains Mode Spectate page and save it into a google sheets. Return the unique values.
def update_draft_data():
  gc = gspread.oauth(
    credentials_filename='Your/Path/Oauth_API_MapCiv_Draft.json', 
    authorized_user_filename='Your/Path/token.json'
  )

  ### Pull list of last 30 drafts ###
  url = "https://aoe2cm.net/api/recentdrafts"
  response = requests.get(url)

  if response.status_code == 200:
    data = response.json()
    print(data[1])
  else:
    print(f"Error: {response.status_code}")
  
  print("Updating draft with new data")

  #reformat JSON into list
  #BUG: FALSE is being entered into Google Sheets as 'FALSE.
  #This is avoided by converting ongoing into a string.
  ref_data = []
  for item in data:
      row = [item.get('title', ''), item.get('draftId', ''), item.get('presetId', ''),
            str(item.get('ongoing', '')), item.get('nameHost', ''), item.get('nameGuest', '')]
      ref_data.append(row)

  ### Create sheet or add data to Existing sheet ###
  date_str = datetime.datetime.now().strftime("%Y-%m-%d")
  sheet_name = f"Captains Mode {date_str}"
  print(sheet_name)

  #check if sheet_name already exists with today's date, else create one
  try:
    gs = gc.open(sheet_name)
  except gspread.exceptions.SpreadsheetNotFound:
    print(f"Creating new sheet: {sheet_name}")
    gs = gc.create(sheet_name)

  #get first worksheet
  gw = gs.get_worksheet(0)

  #first line with extra spaces behind "Title" for print formatting
  row_one = [["Title        ", "Draft ID", "Preset ID", "Ongoing", "Host", "Guest"]]
  row_one_single = ["Title        ", "Draft ID", "Preset ID", "Ongoing", "Host", "Guest"]

  #get all existing data
  values = gw.get_all_values()

  #append new data
  if values == []:
    values = row_one + ref_data
  else:
    values = row_one + ref_data + values

  # Check for duplicate rows and remove them
  unique_values = []
  for row in values:
      if not any(all(a == b for a, b in zip(row, existing_row)) for existing_row in unique_values):
          unique_values.append(row)

  # Clear the sheet and write the unique values
  gw.clear()
  gw.insert_rows(unique_values)
  print("Updated_Draft_Data:")
  print_table(unique_values, row_one_single)
  return(unique_values)

#Print the data row by row
def print_table(nested_list, column_names):
    num_cols = len(column_names)
    header_row = "| " + " | ".join(column_names) + " |"
    nice_horizontal_rule = ("|" + "-" * (len(header_row)-2)+"|")
    print(nice_horizontal_rule)
    print(header_row)
    print(nice_horizontal_rule)

    #enumerate so row numbers are present (rowdex)
    for rowdex, item in enumerate(nested_list):
        # writing each row to a string,
        # then printing the string, is better for performance:)
        s = "|"
        for i in range(num_cols):
            if i == 0:
              #only include first 12 characters in string
              entry = str(item[i])[:12]
              #check if string includes the phrase "civ"
              if "civ" in str(item[i]) or "Civ" in str(item[i]) or "CIV" in str(item[i]):
                entry += " civ"
              #check if string includes the phrase "map"
              elif "map" in str(item[i]) or "Map" in str(item[i]) or "MAP" in str(item[i]):
                entry += " map"
              else:
                entry += "      "
            else:
              entry = str(item[i])
            s += (" "*(len(column_names[i]) - len(entry)+2) +
                entry + "|")
        print(s + " " + str(rowdex))
    print(nice_horizontal_rule)

#Reformat the list data based on a search_term found in either columns 1, 5, or 6
def filter_table(nested_list, search_term):
  #take a nested list and only return rows that contain the search term in either the draft name, host, or guest slots
  filtered_list = []
  for row in nested_list:
        if search_term in row[0] or search_term in row[4] or search_term in row[5]:
            filtered_list.append(row)
  return filtered_list

#Reformat the data to be displayed row by row
def format_table(nested_list, column_names):
    num_cols = len(column_names)
    header_row = "| " + " | ".join(column_names) + " |"
    nice_horizontal_rule = ("|" + "-" * (len(header_row)-2)+"|")

    # Build the formatted string
    formatted_string = f"{nice_horizontal_rule}\n{header_row}\n{nice_horizontal_rule}\n"
    for rowdex, item in enumerate(nested_list):
        s = "|"
        for i in range(num_cols):
            if i == 0:
                entry = str(item[i])[:12]
                if "civ" in str(item[i]) or "Civ" in str(item[i]) or "CIV" in str(item[i]):
                    entry += " civ"
                elif "map" in str(item[i]) or "Map" in str(item[i]) or "MAP" in str(item[i]):
                    entry += " map"
                else:
                    entry += "      "
            else:
                entry = str(item[i])
            s += (" "*(len(column_names[i]) - len(entry)+2) +
                   entry + "|")
        formatted_string += f"{s} {rowdex}\n"
    formatted_string += nice_horizontal_rule

    return formatted_string

#Refreshes draft based on filter_text input
#not currently used
def update_filter():
  obs.obs_property_list_clear()
  print("using filter: " + search_filter)
  filtered_script = format_table(filter_table(unformated_draft, search_filter), row_one_single)
  print(filtered_script)

#Default values in each of the named objects in the properties pane on the OBS Script window
def script_defaults(settings):
  obs.obs_data_set_default_string(settings, "filter_text", "")
  obs.obs_data_set_default_string(settings, "draft_return", cur_draft) #should display formatted string of drafts with numbers to the far right
  obs.obs_data_set_default_int(settings, "map_row", 0)
  obs.obs_data_set_default_int(settings, "civ_row", 0)

#Replaces the current string shown in Drafts with a new string based on the filter.
# Called when the button is pressed as a callback in script_properties
def refresh_filter(draft_property):
  obs.obs_property_list_clear(draft_property)
  obs.obs_property_list_add_string(draft_property, "", "")
  obs.obs_property_list_add_string(draft_property, format_table(filter_table(unformated_draft, search_filter), row_one_single), "draft")

#updates draft links. Needs script_update to run first to get non-zero values
def draft_link_update(map_row, civ_row):
  map_link = "no link"
  civ_link = "no link"

  filtered_draft = filter_table(unformated_draft, search_filter)
  map_link = "https://aoe2cm.net/draft/{}".format(filtered_draft[map_row][1])
  civ_link = "https://aoe2cm.net/draft/{}".format(filtered_draft[civ_row][1])

  print("Maps: ", map_row, map_link)
  print("Civs: ", civ_row, civ_link)

  map_browser = obs.obs_get_source_by_name("Auto_MapDraft")
  civ_browser = obs.obs_get_source_by_name("Auto_CivDraft")

  map_settings = obs.obs_data_create()
  civ_settings = obs.obs_data_create()
  #assuming "url" is the name of the property in the OBS browser source.
  print("Changing map url")
  obs.obs_data_set_string(map_settings, "url", map_link)
  print("Changing civ url")
  obs.obs_data_set_string(civ_settings, "url", civ_link)
  obs.obs_source_update(map_browser, map_settings)
  obs.obs_source_update(civ_browser, civ_settings)

  obs.obs_data_release(map_settings)
  obs.obs_data_release(civ_settings)

  obs.obs_source_release(map_browser)
  obs.obs_source_release(civ_browser)

#Props creates properties objects in the OBS Script GUI
def script_properties():
  props = obs.obs_properties_create()
  
  draft_property = obs.obs_properties_add_text(props, "draft_return", "Drafts:", 3) # 3 refers to OBS_TEXT_INFO
  refresh_filter(draft_property)
  
  obs.obs_properties_add_text(props, "filter_text", "Search:", 0) # 0 refers to OBS_TEXT_DEFAULT
  
  # Button to refresh the drafts based on the filter
  obs.obs_properties_add_button(props, "refresh_button", "Refresh Filter",
    lambda props,prop: True if refresh_filter(draft_property) else True)
  
  obs.obs_properties_add_int_slider(props, "map_row", "Map Row", 0, 250, 1)
  obs.obs_properties_add_int_slider(props, "civ_row", "Civ Row", 0, 250, 1)
  obs.obs_properties_add_button(props, "link_button", "Update Links",
    lambda props,prop: True if draft_link_update(map_row, civ_row) else True)

  return props

#Update settings object with new user-input values in the named objects. Runs when GUI is altered.
def script_update(settings):
  global search_filter, map_row, civ_row
  search_filter = obs.obs_data_get_string(settings, "filter_text")
  draft_rows = obs.obs_data_get_string(settings, "draft_return")
  print("new filter: " + search_filter)
  filtered_draft = format_table(filter_table(unformated_draft, search_filter), row_one_single)
  print(filtered_draft)
  #draft_return will be overwritten by the default value cur_draft even after the button update
  #the next line sets the string again.
  obs.obs_data_set_string(settings, "draft_return", filtered_draft)
  map_row = obs.obs_data_get_int(settings, "map_row")
  civ_row = obs.obs_data_get_int(settings, "civ_row")
  print(str(map_row) + " " + str(civ_row))







  


  
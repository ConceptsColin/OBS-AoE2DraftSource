Running on Python 3.12.5

Requires installation of gspread `pip install gspread` and google api client `pip install google-api-python-client`

Captains Mode is a draft system for the Age of Empires series created by the group @SiegeEngineers found at https://aoe2cm.net/
It allows players to draft either maps or civilizations to play from in a tournament setting. An API page is provided with an Endpoint that allows the last 30 drafts to be retreived: GET https://aoe2cm.net/api/recentdrafts
The google spreadsheet python authorization requires signing up for a google cloud account, creating a project, and setting up the proper authorization. More information is available here: https://github.com/burnash/gspread/blob/master/docs/oauth2.rst

The script should be added to the OBS scripts window and reloaded to show the properties in the GUI.
![Screenshot 2024-09-05 173308](https://github.com/user-attachments/assets/e9f9abbe-2206-43e8-8917-05f47f7520ce)
Once you add a filter (CapsSensitive) and click the "Refresh Filter" button, you can narrow down the list to the draft you are looking to display.
![Screenshot 2024-09-05 173403](https://github.com/user-attachments/assets/fdb1af19-3ad2-45de-9b3c-2af78a240038)
Line up the integer sliders with the numbers shown to the right of the draft row, and click the "Update Links" button. The OBS Browser sources by default are named "Auto_MapDraft" and "Auto_CivDraft". Adjust the script and source names as necessary.

Pros:
- The draft data is saved to a new google sheet each day you reload the script. n a busy weekend, the 30 drafts shown on the spectate page can be replaced in a 15 minute span, so saving the draft links to a google spreadsheet is very helpful if the next matches' players are drafting but you are not ready for the link.
- Browser sources do not require additional windows to be open (as for window capture), and you can right click the browser source to Interact and adjust win/loss with the map/civ.
- No copy/paste of draft urls.
- Custom html/css within your browser source to display the draft as you want it.

Cons:
- Browser sources may be more resource intensive than some window capture sources depending on your internet browser.
- Set up and maintanance is more extensive.

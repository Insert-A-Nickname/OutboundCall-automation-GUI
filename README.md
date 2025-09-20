# OutboundCall-automation-GUI
As for the functioning, the app requires the next dependencies:
-pywinauto
-pandas
-tkinter
-sounddevice
-pyautogui
-phonenumbers
-numpy
-pyttsx3
-pypdf
The functioning of the buttons:
- The config button shows the configuration panel and allows to change the  values of the config.json files
- the click buttons enable the textbox to input the 'x' and 'y' coordinates to change or update them
- the itembox shows the currently open windows in the system, by selecting one of them and then clicking over, it
  it sets the json file to search for the app, it's currently using whatsapp for testing
start and record work smoothly
- the read button starts the audiobook reading (for optimization I would like to add a funtion to remove the '{3:1}'
  from the text, and split diferently, as there are some blank pages in the book, the function is easy but WhatsApp
  shuts the volume of the audio down in favor of the call microphone)
- Before starting the automation an item from the treebox has to be clicked
- Browse Sheet selects from the browser the spreadsheet to use
- pdf_status browses for the path of the pdf to read, this also means it can read other books

The other files:
- Generate ODS creates the testing files
- Generate config creates a blank or default configuration
- Decker file is untested

Ways to improve it:
I saw that you wanted to get some of the spreadsheets directly from the browser
this can be done with sql and get commands, but requires the links you want to get
the app can also be optimised for usage with another calling app, but I don't
have any license nor income to test it currently, same goes for getting the audio files, an
overhaul could be done with the links and the license from an app.
from the outbound calls app, which, in turn can also make the code run smoother as it removes
the need to record directly, since this one records directly, it doesn't have conversion.

possible bugs found through testing:
- extra feature to read book might crash the app
- without the proper wimdow-click-setup might skip a call
- at random might set a call as done while not having done it
- currently, before starting the automation an item from the treebox has to be clicked
- sets the text of the itembox to select the window to go in front as the path to the spreafsheet



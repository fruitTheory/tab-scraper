import time
import pywinauto
import urllib.request
import subprocess
from urllib.parse import urlparse
from os.path import splitext
import logging
import pyautogui
from win32gui import GetWindowText, GetForegroundWindow


# Basic logging writes new file each log
logging.basicConfig(filename='debug.log', encoding='utf-8', level=logging.DEBUG, filemode='w')

print("Script is running please wait until finished... ")

# Containers
tab_list = []
tab_count = 0

limit_reached = False

# While limit reached is not true continue program
while not limit_reached:
    # This delay matters, else the script will go out of range
    time.sleep(.25)
    # Grab the current window
    fg_window = GetWindowText(GetForegroundWindow())
    # I believe split by word/space
    split_name = fg_window.split()
    # Gets last two words in list and concatenates
    window_name = split_name[-2]+split_name[-1]

    # If window name has Microsoft Edge in it call it a hit
    edge_name = 'Microsoft Edge'
    if "Microsoft" and "Edge" in window_name:
        print("Edge selected!")
        edge_selected = True
    else:
        print("Not selected!")
        edge_selected = False



    while(edge_selected):

        print("Hey its me")

        # Object constructor
        app = pywinauto.Application(backend='uia')

        # Connects to any application named Microsoft Edge
        app.connect(title_re='.*Microsoft​ Edge.*', found_index=0)
        dlg = app.top_window()

        # Get the toolbar containing the address bar
        wrapper = dlg.child_window(title='App bar', control_type='ToolBar')

        # Store reference to child window and descendants are the controls contain in window
        # Returns list of all editable control types guessing [0] represents url edit box
        url = wrapper.descendants(control_type='Edit')[0]
        # Gets value of url edit box
        url_name = url.get_value()

        # Append url to list of urls
        tab_list.append(url_name)

        # This parses through url, and we choose to return 'path' in order to split extension
        parsed_url = urlparse(url_name)
        url_path = parsed_url.path
        root, ext = splitext(str(url_path))

        # If our list of items starts having duplicates we end flow
        if tab_list.count(url_name) > 1:
            # process.kill()
            print("Reached tab limit")
            limit_reached = True
            break

        # Logging tab names and how many
        logging.debug(tab_list)
        logging.debug("Number of items = " + str(len(tab_list)))
        # print(tab_list)

        # # if there is an extension available consider it writable
        # if ext:
        #     tab_count += 1
        #     # Write binary mode wb -- opening binary file writing img to it and closing
        #     f = open(str(tab_count)+ext, 'wb')
        #     # Try except to catch forbidden errors
        #     try:
        #         # Try requesting open url and writing the return to a binary file
        #         request = urllib.request.urlopen(url_name).read()
        #         f.write(request)
        #         f.close()
        #         print("finished")
        #     except urllib.error.HTTPError as forbid:
        #         if forbid.code == 403:
        #             logging.error("HTTP Error 403 Forbidden")
        #             print("403 junk")
        #             #process.kill()
        #             break

        # Simulate keyboard shortcut
        pyautogui.hotkey('ctrl', 'tab')

        # Break needed at end of this basically as a reset
        break

import pywinauto
import urllib.request
import subprocess
from urllib.parse import urlparse
from os.path import splitext
import logging

# Basic logging writes new file each log
logging.basicConfig(filename='debug.log', encoding='utf-8', level=logging.DEBUG, filemode='w')

print("Script is running please wait until finished... ")

# Popen allows other events to occur while the process is running
process = subprocess.Popen([r'C:\Program Files\AutoHotkey\v1.1.36.02\AutoHotkeyU64.exe',
                            r'tab_script.ahk'])

# Containers
tab_list = []
tab_count = 0

while(process):
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
        process.kill()
        break

    logging.debug(tab_list)
    logging.debug("Number of items = " + str(len(tab_list)))
    # print(tab_list)

    # if there is an extension available consider it writable
    if ext:
        tab_count += 1
        # Write binary mode wb -- opening binary file writing img to it and closing
        f = open(str(tab_count)+ext, 'wb')
        # Try except to catch forbidden errors
        try:
            # Try requesting open url and writing the return to a binary file
            request = urllib.request.urlopen(url_name).read()
            f.write(request)
            f.close()
        except urllib.error.HTTPError as forbid:
            if forbid.code == 403:
                logging.error("HTTP Error 403 Forbidden")
                process.kill()
                break

import os
import re
import time
import logging

import pywinauto
import pyautogui
import requests
import send2trash

from PIL import Image
from win32gui import GetWindowText, GetForegroundWindow


# Basic logging writes new file each log
logging.basicConfig(filename='debug.log', encoding='utf-8', level=logging.ERROR, filemode='w')

# Let user know its running
print("Script is running please wait until finished... ")

# Make image directory
if not os.path.exists('images'):
    os.makedirs('images')
image_dir = 'images/'

# Global Containers
tab_list = []

#* Find all numbers in a string, and return them as a list of floats or ints,
def find_numbers_in_string(input_string):
    # find float point
    numbers = re.findall(r'[-+]?\d*\.\d+|[-+]?\d+', input_string)
    # find just int
    #numbers = re.findall(r'\d+', input_string)
    return [float(num) if '.' in num else int(num) for num in numbers]

def delete():

    images_path = os.listdir(image_dir)
    for image in images_path:
        file_path = image_dir + image

        try:
            # Try seeing if file will open what PIL considers an image
            Image.open(file_path)
        except:
            # If error raised from PIL then send non image file to trash
            logging.error("Deleted " + file_path)
            send2trash.send2trash(file_path)

    
def save(tab_index):
    for tab in tab_list:
        # Get the image, and set the user agent, this is to avoid getting a 403 error,
        # some websites block requests coming outside a browser, so we need to set the user agent to a browser
        request = requests.get(tab, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/91.0.4472.124 Safari/537.36'}).content
        # Write the files
        with open(f"images/{tab_index}.png", "wb") as file:
            file.write(request)
        logging.error(f"Image {tab_index}.png saved from url - {tab}")
        tab_index += 1

    delete()
    print("")
    print("Complete!")

def check():
    # Start connection to edge
    app = pywinauto.Application(backend='uia').connect(title_re='.*Microsoft​ Edge.*', found_index=0)
    # Try to get the number of tabs
    try:
        # Splits the app window info at where it states the number of additional tabs open
        arr = str(app.windows()).split("'")[1].split("and")[1]
        # Extracts the number from this and adds 1 to account for active tab
        tabs = find_numbers_in_string(str(arr))[0] + 1
        print("Found " + str(tabs) + " active tabs")
    # If there is an IndexError, then there is only one tab, so set tabs to 1
    except IndexError:
        tabs = 1

    # Get info from forefround window
    fg_window = GetWindowText(GetForegroundWindow())
    # Split by word/space
    split_name = fg_window.split()
    # Gets last two words in list and concatenates
    window_name = split_name[-2] + split_name[-1]

    if "Microsoft" and "Edge" in window_name:
        print("Edge selected!")
        print("Please wait..")
        edge_selected = True
    else:
        # print("Not selected!")
        edge_selected = False

    if edge_selected:
        # Loop through tab amount
        for i in range(tabs):
            # Pull top window
            dlg = app.top_window()

            # Get toolbar containing the address bar
            wrapper = dlg.child_window(title='App bar', control_type='ToolBar')

            # Returns list of all editable control types [0] represents url edit box
            url = wrapper.descendants(control_type='Edit')[0]
            # Gets value of url edit box
            url_name = url.get_value()

            tab_list.append(url_name)

            # Temp extra guard for tab list
            if not tab_list.count(url_name) > 1:
                # Simulate keyboard shortcut
                pyautogui.hotkey('ctrl', 'tab')
        # Call save pass index 0
        save(0)

    elif not edge_selected:
        print("Please select an instance of Microsoft Edge\n")
        # The delay matters, else the script will go out of range
        time.sleep(2)
        # Another manual loop
        check()


# Basically a manual loop
check()

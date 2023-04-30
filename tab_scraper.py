import requests
import pygetwindow as gw
import pywinauto
import time
import re
import logging
import os
from win32gui import GetWindowText, GetForegroundWindow
from PIL import Image
import send2trash
import PIL
from os.path import splitext
from urllib.parse import urlparse

print("Script will start now, please wait until finished...\nThe script will always try to have the browser window in focus, try not to move the window while the script is running")
time.sleep(3)
# This is used to get the current time, and then use it as the file name for the log file
now = time.strftime("%Y-%m-%d %H-%M-%S")

# Create the images and logs folder if they don't exist
if not os.path.exists('images'):
    os.makedirs('images')
image_dir = 'images/'
if not os.path.exists('logs'):
    os.makedirs('logs')
    

# Set logging config, this will log all errors to a log file, the log file will be saved in the logs folder, and will be named with the current time
logging.basicConfig(filename= f'logs/{now}.log', level=logging.ERROR, format='%(asctime)s - %(message)s')

#* Find all numbers in a string, and return them as a list of floats or ints, this is used to get the number of tabs in Microsoft Edge
def find_numbers_in_string(input_string):
    numbers = re.findall(r'[-+]?\d*\.\d+|[-+]?\d+', input_string)
    return [float(num) if '.' in num else int(num) for num in numbers]

def delete():
    # List directory of our created images folder
    images_path = os.listdir(image_dir)
    for image in images_path:
        file_path = image_dir + image
        try:
            # Try seeing if file will open what PIL considers an image
            Image.open(file_path)
        except PIL.UnidentifiedImageError:
            # If error raised from PIL then send non image file to trash
            logging.error("Deleted " + file_path)
            send2trash.send2trash(file_path)
        except PermissionError:
            # A specific exception for folder permissions, also saves folder from being deleted
            logging.error("Folder Permission Error: ")

# Global Containers
tab_list = []
url_root_list = []
url_ext_list = []

def save():
    tab_index = 0
    for tab in tab_list:
        # Get the image, and set the user agent, this is to avoid getting a 403 error
        # some websites block requests coming outside of a browser, so we need to set the user agent to a browser
        try:
         request = requests.get(tab, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}).content
        # If the url is bad, or empty, then skip it
        except requests.exceptions.MissingSchema:
            continue
         # From the list of url roots grab [x] index
        root_name = url_root_list[tab_index]
        # Split the url text by /
        split_name = root_name.split('/')
        # Grab last one in list
        saved_name = split_name[-1]
        # Grab from list of saved url extensions
        saved_ext = url_ext_list[tab_index]

        try:
            #* Ive added a check to see if the extension is empty, if it is then save as .png
            #* Some websites dont have extensions in the url, so this is a fix for that
            if saved_ext != "":
                # Write the files
                with open(f"images/{saved_name}{saved_ext}", "wb") as file:
                    file.write(request)
                logging.error(f"Image {saved_name}{saved_ext} saved from url - {tab}")
                tab_index += 1
            else:
                # Write the files
                with open(f"images/{saved_name}.png", "wb") as file:
                    file.write(request)
                logging.error(f"Image {saved_name}.png saved from url - {tab}")
                logging.error (f"Extesion for {saved_name} was not found, saved as .png")
                tab_index += 1
        except PermissionError:
            print("Folder Permission Error: ")
            # This exception is needed to catch bad or empty urls
            with open(f"images/null_file", "wb") as file:
                file.write(request)
            logging.error(f"Image null_file saved from url - {tab}")
            tab_index += 1
        if tab_index == len(tab_list):
            break

    delete()
    print("")
    print("Complete!")

def get_open_tabs():
    # Get the window of Microsoft Edge
    edge_windows = gw.getWindowsWithTitle('Microsoft Edge')
    # If the window is not found, try using a zero-width space
    if not edge_windows:
        edge_windows = gw.getWindowsWithTitle('Microsoft\u200B Edge')  # using zero-width space

    # If the window is found, start the code
    if edge_windows:
        edge_window = edge_windows[0]
        edge_window.activate()
        # Wait for the window to be activated
        time.sleep(0.3)
        
        # check if the window is active
        fg_window = GetWindowText(GetForegroundWindow())
        # I believe split by word/space
        split_name = fg_window.split()
        # Gets last two words in list and concatenates
        try:
            window_name = split_name[-2]+split_name[-1]
        except IndexError:
            window_name = ""
            pass
        # If window name has Microsoft Edge in it call it a hit
        edge_name = 'Microsoft Edge'
        if "Microsoft" and "Edge" in window_name:
            edge_selected = True
        else:
            edge_selected = False

        while edge_selected:
            # Try to connect to the window
            try:
                app = pywinauto.Application(backend='uia').connect(title_re='.*Microsoft.*Edge.*')

            # If there are multiple windows open, print an error message and stop the program
            except pywinauto.findwindows.ElementAmbiguousError:
                print("Multiple Microsoft Edge windows are open. Please have only one window open.")
                # Stop the program
                return
            
            dlg = app.top_window()

            # Get the toolbar containing the address bar
            wrapper = dlg.child_window(title='App bar', control_type='ToolBar')

            #* Get the number of tabs, really janky way of doing it, but surprisingly it works
            # Try to get the number of tabs
            try:
                # Splits the app window info at where it states the number of additional tabs open
                arr = str(app.windows()).split("'")[1].split("and")[1]
                # Extracts the number from this and adds 1 to account for active tab
                tabs = find_numbers_in_string(str(arr))[0] + 1
            # If there is an IndexError, then there is only one tab, so set tabs to 1
            except IndexError:
                tabs = 1
            
            index = 1
            for i in range(tabs):
                
                print(f"Tab: {index}/{tabs}")
                # Gets the edit box for the url
                url = wrapper.descendants(control_type='Edit')[0]
                # Gets value of url edit box
                url_name = url.get_value()

                tab_list.append(url_name)

                # This parses through url, and we choose to return 'path' in order to split extension
                parsed_url = urlparse(url_name)
                url_path = parsed_url.path
                root, ext = splitext(str(url_path))

                # Checks if the name is already in the list, if it is, then add a number to the end of the name, 
                # this is to avoid overwriting files with the same name
                if root in url_root_list:
                    url_root_list.append(f"{root}({index})")
                else:
                    url_root_list.append(root)
                    
                url_ext_list.append(ext)

                # Press Ctrl+Tab to switch to the next tab
                dlg.type_keys('^({TAB})')
                time.sleep(0.1)
                index += 1
            if index == tabs + 1:
                save()
                break
            
            
    # If the window is not found, print an error message
    else:
        print("Microsoft Edge is not running.")

get_open_tabs()
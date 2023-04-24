import requests
import pygetwindow as gw
import pywinauto
import time
import re
import logging
import os

# This is used to get the current time, and then use it as the file name for the log file
now = time.strftime("%Y-%m-%d %H-%M-%S")

# Create the images and logs folder if they don't exist
if not os.path.exists('images'):
    os.makedirs('images')
if not os.path.exists('logs'):
    os.makedirs('logs')

# Set logging config, this will log all errors to a log file, the log file will be saved in the logs folder, and will be named with the current time
logging.basicConfig(filename= f'logs/{now}.log', level=logging.ERROR, format='%(asctime)s - %(message)s')

#* Find all numbers in a string, and return them as a list of floats or ints, this is used to get the number of tabs in Microsoft Edge
def find_numbers_in_string(input_string):
    numbers = re.findall(r'[-+]?\d*\.\d+|[-+]?\d+', input_string)
    return [float(num) if '.' in num else int(num) for num in numbers]

#* Main function
def get_open_tabs():
    open_tabs = []
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
            arr = str(app.windows()).split("'")[1].split("and")[1]
            tabs = find_numbers_in_string(str(arr))[0] + 1
        # If there is an IndexError, then there is only one tab, so set tabs to 1
        except IndexError:
            tabs = 1
        
        index = 1
        for i in range(tabs):
            
            print(f"Tab: {index}")
            # Gets the edit box for the url
            url = wrapper.descendants(control_type='Edit')[0]
            # Gets value of url edit box
            url_name = url.get_value()

            # Append url to list of urls, to be used to download the images later
            open_tabs.append(url_name)
            print(f"URL: {url_name}")

            # Press Ctrl+Tab to switch to the next tab
            dlg.type_keys('^({TAB})')
            time.sleep(0.1)
            index += 1
        tab_index = 1
        for tab in open_tabs:
            # Get the image, and set the user agent, this is to avoid getting a 403 error, some websites block requests coming outside of a browser, so we need to set the user agent to a browser
            request = requests.get(tab, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}).content

            #* The code below was used to remove special characters from the tab name, and then use it as the file name, but the file name was ugly, so I used the tab index instead
            # tab = str(tab).split("/")[-1]
            # tab = re.sub(r'\W+', '_', tab)[:200]

            # save the image to the images folder
            #? For now, all images are saved as png, not really a problem, but could be fixed later if needed
            #? Maybe try to get the file extension from the url, and then save it as that file extension, should be easy enough
            with open(f"images/{tab_index}.png", "wb") as file:
                file.write(request)
            logging.error(f"Image {tab_index}.png saved from url - {tab}")
            tab_index += 1
    # If the window is not found, print an error message
    else:
        print("Microsoft Edge is not running.")

get_open_tabs()
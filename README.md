# tab-scraper
Script for grabbing images from open tabs for those who have many tabs to gather photo reference. For now it just works with microsoft edge but could be extended if requested. After starting python script click on a browser with all tabs open and it will run through each window. If the window has an extension like .jpg or .png it will write that to disk and continue looping until all tabs have been saved in which it will stop both ahk and py scripts.

There is a couple caveats to this as mentioned its super jank
 --Will need autohotkey installed
 --Click on window browser shortly after running
 --The ahk runs regardless of what window its on so let it run through tabs
 --Possible that not all errors get caught so if ahk keeps running can go to bottom right tray and manually stop

Thats about it, definitely open to suggestions on a better workflow so please feel free to!



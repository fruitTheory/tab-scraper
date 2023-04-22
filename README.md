# tab-scraper
Script for grabbing images from open tabs for those who have to open many tabs to gather photo reference and such. For now it just works with microsoft edge but could be extended if requested. After starting the python script click on a browser with all tabs open and it will run through each window. If the window has an extension like .jpg or .png it was save that file to disk and continue looping until all tabs have been saved in which it will stop both ahk and py scripts.

If you end up using there is a couple caveats to this as mentioned its super jank
 --Will need autohotkey installed
 --Click on window browser somewhat quickly after running script
 --It error catches 403 forbidden so cant be used on sites that block
 --The autohotkey runs regardless of what window its on so have to let it run through all tabs to get each one
 --Possible that not all errors get caught so if ahk keeps running can go to bottom right tray and manually stop

Thats about it, open to suggestions for this jankyness and feel free to raise issues



Attribute VB_Name = "modComments"
Option Explicit

'----------------------------------------------------------------------------------------'
' Multi Downloader using multithreadings
' Created by Suk Yong Kim, 03/14/2001
'
' This project is my first project to upload to the PSC.
' Many persons contribute to create this project
' I really appreicate their efforts and codes and the great server PSC.
'
' if any question, mail to : techtrans@dreamwiz.com
'----------------------------------------------------------------------------------------'


'Planet Source Code Downloader... PSCdl
'
'I made this program to help my self downloade code from PSC. I often download code I think could be usefull later... I put it in a directory and try to manage the code by choosing the same name for the .zip file, the image file and the info/about-file...
'
'This program do it all!
'
'Choose destination directory...
'- By clicking on the little red image you will see a
'You mark the project name with the mouse and drag it to the first open box. If succesful the other boxes will open and the first one close and be green! That meens that you now have choosen the filename for each fiel you drop afterwards.
'
'Always on top!
'- The program will always be on the top and you can drag it around freely!
'
'Drag 'n'drop => download to destination directory
'- You can drag items from the browser window to the Downloader. By dropping a marked text in the second box (Text) you will automaticly save the text in a file named by the filename and endnig with _about.txt. You can drag a image to the third box (Image) and last you can drag the .zip fiel to the last box (.zip). Each drop result in a download and renaming of the file acording to the first drop - the filename.
'
'Loggin the actions...
'- Each action is logged. You can se the log while working or look at it in the file dl.log after closing the program!
'
'
'Use it... If you like!
'
'AND wote for me if you do - please!
'
'/Hannibal


'=========================================================================================
' PSCdl Update
' Secondary Author: John Baughman
' Email: johnb@atomiksolutions.com
' Date: 2001-11
'=========================================================================================
' Added:
'   Database to store data into
'   Base64 encoding to store data (Can someone fill in the code instead of the Base64.dll I found and am using?)
'   Settings moved into registry instead of an INI file. (INI is ooooold school. And these entries aren't in the VB/VBA section...)
'   Browse downloaded code
'   Save to file functionality
'
' Changes:
'   Modified original code to reflect cleanliness
'   Optimized objects by removing redundancies
'   Buttons were cool, but moved options to a popup menu.
'
' To Do:
'   Incorporate into Source Search 2.0 (Created by Casey Goodhew - goodhewc@ hotmail.com - goodhew.2y.net)
'   Make whole app an IE5 toolbar as well as standalone.
'   Port DB to PalmOS format. (I want this as an option for me.)
'
' Notes:
'   Acknowledgements:
'   Hannibal - Original idea
'   Suk Yong Kim - Multi Downloader code borrowed by Hannibal
'   Alvaro Redondo (sevillaonline.com/ActiveX/) - Base 64 DLL (Check out his site for a lot of great free DLL's and OCX's!)
'   Other PSC contributors - If you want your name included here send me a message!

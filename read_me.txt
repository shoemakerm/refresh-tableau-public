
Automated Tableau Desktop Extract Refresh and Republish to Tableau Public

Files are on Github
https://github.com/shoemakerm/refresh-tableau-public/tree/master

Windows Task Scheduler runs the batch file each day

VBScript file performs the automation
1.  Call cURL application to download the .tsv file from the CDC website to a local directory
	-cURL comes with Windows 10 or else install Git

2.  Open the Tableau workbook which launches the application
3.  Sends keystrokes with delays where needed
	-Refresh All Extracts
	-Publish to Tableau Public
	-Switch from browser window back to Tableau Desktop application window
	-Close Tableau Desktop

Why do I have to be logged in?  Only the interactive window station can accept keyboard input
Window Stations:  https://docs.microsoft.com/en-us/windows/win32/winstation/window-stations

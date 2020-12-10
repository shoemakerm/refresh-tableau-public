'Download File
App = "curl"
URL = "https://data.cdc.gov/api/views/9mfq-cb36/rows.tsv?accessType=DOWNLOAD&bom=true"
File = """C:\Users\Matt\Google Drive\git\refresh_tableau_public\United_States_COVID-19_Cases_and_Deaths_by_State_over_Time.tsv"""
CreateObject("WScript.Shell").Run App + " " + URL + " -o " + File

'Open Tableau Workbook
App = """C:\Users\Matt\Google Drive\git\refresh_tableau_public\Covid.twb"""
Set Shell_Object = WScript.CreateObject("WScript.Shell")
Shell_Object.Run App, 9
'Wait for workbook to open (adjust as needed)
WScript.Sleep 40000
Shell_Object.AppActivate("Tableau - Covid")
WScript.Sleep 2000

'Send keystrokes with waits as needed.

'Refresh extracts
Shell_Object.SendKeys "%d"
WScript.Sleep 1000
Shell_Object.SendKeys "x"
WScript.Sleep 1000
Shell_Object.SendKeys "{ENTER}"
WScript.Sleep 12000
Shell_Object.SendKeys "{ENTER}"
WScript.Sleep 2000

'Publish workbook
Shell_Object.SendKeys "%s"
WScript.Sleep 1000
Shell_Object.SendKeys "w"
WScript.Sleep 1000
Shell_Object.SendKeys "{TAB}"
WScript.Sleep 1000
Shell_Object.SendKeys "{TAB}"
WScript.Sleep 1000
Shell_Object.SendKeys Chr(32)
WScript.Sleep 1000
Shell_Object.SendKeys "{TAB}"
WScript.Sleep 1000
Shell_Object.SendKeys "{ENTER}"
WScript.Sleep 10000
'Turn NumLock back on because it gets turned off for some reason
WScript.Sleep 2000
Shell_Object.SendKeys "{NUMLOCK}"

'Close Tableau App
Shell_Object.AppActivate("Tableau - Covid")
WScript.Sleep 2000
Shell_Object.SendKeys "%{F4}"

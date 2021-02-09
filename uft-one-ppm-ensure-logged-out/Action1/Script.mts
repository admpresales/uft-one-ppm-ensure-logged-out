'===========================================================
'20201001 - Initial creation
'20201007 - Renamed action
'			Removed unused function
'			Cleaned up OR of unused objects
'			Cleaned up OR names
'20210209 - DJ: Updated to start the mediaserver service on the UFT One host machine if it isn't running
'===========================================================

Dim BrowserExecutable, oShell

Set oShell = CreateObject ("WSCript.shell")
oShell.run "powershell -command ""Start-Service mediaserver"""
Set oShell = Nothing

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM login page
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Logout if option is available to logout
'===========================================================================================
If AIUtil("search").Exist(1) Then
	Browser("PPM").Page("PPM Main Page").WebElement("menuUserIcon").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	AIUtil.FindText("Sign Out (").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
End If

AppContext.Close																			'Close the application at the end of your script


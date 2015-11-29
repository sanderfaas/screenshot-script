'* ****************************************************************************************************
'* screenshots.vbs
'* Sander Faas
'* Netherlands, 2015
'* 
'* Make screenshots on a client computer and store them on a network share.
'*
'* Requirements:
'* - A domain user (in this script called Screenshot Administrator) that will be used to 
'*   connect to the network share where the screenshots will be stored.
'* - A network share folder with full control permissions (or maybe just modify permissions)
'*   for the Screenshot Administrator user. This folder should contain 2 subfolders: one to
'*   store the screenshots and one where the tools psexec.exe and nircmd.exe can be found.
'* - PsExec.exe (https://technet.microsoft.com/en-us/sysinternals/psexec.aspx) in the Tools
'*   folder on the network share.
'* - NirCmd.exe (http://www.nirsoft.net/utils/nircmd.html) in the Tools folder on the network
'*   share
'*
'* ****************************************************************************************************

'* **************************************************
'* SETTINGS (CHANGE THESE TO YOUR OWN SITUATION)
'* **************************************************

'* SharePath refers to the share on the server reserved for this screenshot service.
'* Make sure to use the IP address of the fileserver instead of the UNC name
SharePath = "\\192.168.1.2\Screenshots$"
'* ToolsPath refers to the path on the server where psexec.exe and nircmd.exe reside
ToolsPath = SharePath & "\Tools"
'* ShotsPath refers to the path on the server where the screenshots will be stored
ShotsPath = SharePath & "\Shots"
'* Credentials of the Screenshot Administrator
scrU = "DOMAIN\USERNAME" 'User name of the Screenshot Administrator
scrP = "PASSWORD" 'Password of the Screenshot Administrator
'* Number of screenshots to take and time interval between them
'* These variables are strings (use quotes) instead of integers!
strShots = "360" 'Number of screenshots to take (pick a high number, the process will stop
                 'automatically when the user signs off)
strInterval = "30000" 'Interval between two screenshots in milliseconds


'* **************************************************
'* CREATE NECESSARY OBJECTS
'* **************************************************

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject( "WScript.Shell" )
Set oNetwork = CreateObject("WScript.Network")


'* **************************************************
'* CHECK FOR CORRECT NETWORK SHARE FOLDERS
'* **************************************************

'* Connect to network share with credentials of the Screenshot Administrator
On Error Resume Next
oNetwork.MapNetworkDrive "", SharePath, False, scrU, scrP 
If Err.Number <> 0 Then
	'Handle Error
	MsgBox "Error connecting to network share!", vbError, "Screenshot service"
	'Exit Sub
End If

'* Check for share path existence
If Not (oFSO.FolderExists(ShotsPath) and oFSO.FolderExists(ToolsPath)) Then
	MsgBox "Screenshotservice share paths cannot be found!",vbExclamation,"Screenshots service"
	'Exit Sub
End If

'* Check for PsExec.exe and nircmd.exe in the Tools path
strPsExec = ToolsPath & "\" & "psexec.exe"
strNircmd = ToolsPath & "\" & "nircmd.exe"
If Not (oFSO.FileExists(strPsExec) And oFSO.FileExists(strNircmd)) Then
	MsgBox "Screenshot tools not available.",vbExclamation,"Screenshot service"
	'Exit Sub
End If


'* **************************************************
'* CREATE APPROPRIATE FOLDERS TO STORE SCREENSHOTS
'* **************************************************

'* DatePath refers to the path where *today's* screenshots will be stored
strYear = NumToChar(Year(Now),4)
strMonth = NumToChar(Month(Now),2)
strDay = NumToChar(Day(Now),2)
strDate= strYear & "_" & strMonth & "_" & strDay
DatePath = ShotsPath & "\" & strDate

'* UserPath refers to the path where *this user's* screenshots will be stored    
strUser = oShell.ExpandEnvironmentStrings("%UserName%")
strComputer = oShell.ExpandEnvironmentStrings("%ComputerName%")
UserPath = DatePath & "\" & strUser & "(" & strComputer & ")"

'* Check for date path existence. Create folder if necessary
If Not oFSO.FolderExists(DatePath) Then
	oFSO.CreateFolder(DatePath)
End If

'* Check for user path existence. Create folder if necessary
If Not oFSO.FolderExists(UserPath) Then
	oFSO.CreateFolder(UserPath)
End If  


'* **************************************************
'* START MAKING SCREENSHOTS
'* **************************************************

'* Create the command to run: psexec.exe \\THISCOMPUTER -u "SCREENSHOTADMINISTRATOR" -p "PASSWORD" -i -d -c -e -accepteula nircmd.exe loop TIMES INTERVAL savescreenshot PATH\scr~$currtime.HH_mm_ss$.jpg
strCmd = "%Comspec% /c " & strPsExec & " \\%COMPUTERNAME% -u """ & scrU & """ -p """ & scrP & """ -i -d -c -e -accepteula " & strNirCmd & " loop " & strShots & " " & strInterval & " savescreenshot " & UserPath & "\scr~$currtime.HH_mm_ss$.jpg"
intRunError = oShell.Run(strCmd, 0, True)
If inRunError <> 0 Then
	MsgBox "Error making screenshots.",vbExclamation,"Screenshots service"
	'Exit Sub
End If    
	
REM MsgBox "Screenshot service started"


'* **************************************************
'* CLEAN UP
'* **************************************************

Set oFSO = Nothing
On Error Resume Next
oNetwork.RemoveNetworkDrive SharePath, True, False
Set oShell = Nothing
Set oNetwork = Nothing


'* **************************************************
'* FUNCTIONS
'* **************************************************

Function NumToChar(intValue, intLen)
    NumToChar = Right(String(intLen, "0") & intValue, intLen)
End Function

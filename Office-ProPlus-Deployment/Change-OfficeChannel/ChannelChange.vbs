'declare variables
Dim oReg
Dim sValue
Dim oWmiLocal
Dim f64
Dim Channels
Dim CurrentChannel
Dim CurrentChannelDisplay
Dim Selection
Dim UpdateToVersion
Dim SingleChannel

Set oWmiLocal   = GetObject("winmgmts:{(Debug)}\\.\root\cimv2")
Set oReg        = GetObject("winmgmts:\\.\root\default:StdRegProv")
Set Channels = CreateObject("Scripting.Dictionary")
Const HKLM          = &H80000002

'modify channels here
Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "CC/Current"
	.url = "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
End With
Channels.Add "1", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT CC/Current"
	.url = "http://officecdn.microsoft.com/pr/cbc7891e-9126-44de-8a56-2bd6d2e06c48"
End With
Channels.Add "2", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "FRCC/InsiderSlow"
	.url = "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
End With
Channels.Add "3", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "FRCC/InsiderFast"
	.url = "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
End With
Channels.Add "4", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "DC/Deferred"
	.url = "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
End With
Channels.Add "5", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT DC/Deferred"
	.url = "http://officecdn.microsoft.com/pr/f4f024c8-d611-4748-a7e0-02b6e754c0fe"
End With
Channels.Add "6", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT Elite"
	.url = "http://officecdn.microsoft.com/pr/2c508370-a266-4cfc-8877-af06fdeb0c24"
End With
Channels.Add "7", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "FRDC/Validation"
	.url = "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
End With
Channels.Add "8", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT FRDC/Validation"
	.url = "http://officecdn.microsoft.com/pr/9a3b7ff2-58ed-40fd-add5-1e5158059d1c"
End With
Channels.Add "9", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT Elite (DevMain)"
	.url = "http://officecdn.microsoft.com/pr/b61285dd-d9f7-41f2-9757-8f61cba4e9c8"
End With
Channels.Add "10", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "MSIT CC/Current (DevMain)"
	.url = "http://officecdn.microsoft.com/pr/5462EEE5-1E97-495B-9370-853CD873BB07"
End With
Channels.Add "11", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "Dogfood DevMain"
	.url = "http://officecdn.microsoft.com/pr/ea4a4090-de26-49d7-93c1-91bff9e53fc3"
End With
Channels.Add "12", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "Dogfood Canary DevMain"
	.url = "\\offdog\odfserver\V2\Releases_RTRel_Canary"
End With
Channels.Add "13", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "Dogfood CC"
	.url = "http://officecdn.microsoft.com/pr/f3260cf1-a92c-4c75-b02e-d64c0a86a968"
End With
Channels.Add "14", SingleChannel

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "Dogfood DevMain"
	.url = "http://officecdn.microsoft.com/pr/55C44C35-878E-4C43-83EE-B666BF4261A4"
End With
Channels.Add "15", SingleChannel
'end of modify channels

Set SingleChannel = New OfficeUpdateChannel
With SingleChannel
	.name = "Dogfood FRDC"
	.url = "http://officecdn.microsoft.com/pr/834504cc-dc55-4c6d-9e71-e024d0253f6d"
End With
Channels.Add "16", SingleChannel
'program begins with for next loop which checks if client machine is 64 bit or not
Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
For Each Item In ComputerItem
    f64 = Instr(Left(Item.SystemType,3),"64") > 0
    If f64 Then Exit For
Next

	'Program grabs UpdateUrl regkey from function, and displays it to user.
	CurrentChannel = FindRegKey	
	CurrentChannelDisplay = GetChannelAndUrl(CurrentChannel)
	Wscript.echo "Your Current channel and update path is: "+ vbCrLf +CurrentChannelDisplay
	Selection = GetUserSelection()
	UpdateChannelAndVersion(Selection)
	'WScript.Echo Channels.Item("1").name + vbCrLf + Channels.Item("2").name
	
	
	
	
	
	
Function FindRegKey()
        If RegKeyExists(HKLM,"SOFTWARE\Microsoft\Office\ClickToRun\Configuration\") Then
           FindRegKey = RegReadValue (HKLM,"SOFTWARE\Microsoft\Office\ClickToRun\Configuration\","UpdateUrl",sValue,"REG_SZ")
	End If
End Function


Function FindClientPath()
        If RegKeyExists(HKLM,"SOFTWARE\Microsoft\Office\ClickToRun\Configuration\") Then
           FindClientPath = RegReadValue (HKLM,"SOFTWARE\Microsoft\Office\ClickToRun\Configuration\","ClientFolder",sValue,"REG_SZ")
	End If
End Function


'Read the value of a given registry entry
Function RegReadValue(hDefKey, sSubKeyName, sName, sValue, sType)
    Dim RetVal
    Dim Item
    Dim arrValues
    
    Select Case UCase(sType)
        Case "REG_SZ"
            RetVal = oReg.GetStringValue(hDefKey,sSubKeyName,sName,sValue)
            If Not RetVal = 0 AND f64 Then 
            RetVal = oReg.GetStringValue(hDefKey,Wow64Key(hDefKey, sSubKeyName),sName,sValue)
        	End If        
        Case Else
            RetVal = -1
    End Select 'sValue
    
    RegReadValue = sValue
End Function 'RegReadValue

Function RegKeyExists(hDefKey,sSubKeyName)
    Dim arrKeys
    RegKeyExists = False
    If oReg.EnumKey(hDefKey,sSubKeyName,arrKeys) = 0 Then RegKeyExists = True
End Function

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos

    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case HKLM
            If Left(sSubKeyName,17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName,17) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName,"\")
                Wow64Key = Left(sSubKeyName,iPos) & "Wow6432Node\" & Right(sSubKeyName,Len(sSubKeyName)-iPos)
            End If
        
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
        
    End Select 'hDefKey
End Function 'Wow64Key

Function WriteRegValue(UpdateUrlValue)
	Dim wshShell
	Set wshShell = CreateObject("WScript.Shell")
	wshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\UpdateUrl", UpdateUrlValue, "REG_SZ"
End Function


'Find channel based on given URL from regkey
Function GetChannelAndUrl(CurrentUrl)
	Dim TempReturnString
	For Each ChannelItem In Channels.Items
		If InStr(LCase(CurrentUrl), LCase(ChannelItem.url)) Then
			TempReturnString = "Channel: " + ChannelItem.name + +vbCrLf+" URL: " + ChannelItem.url
			Exit For
			End If
	Next
	If (TempReturnString = "") Then
		If (CurrentUrl <> "") Then
			TempReturnString = "Channel: Other CDN "+vbCrLf+" URL: "+CurrentUrl
		Else 
			TempReturnString = "Channel: UpdateUrl is not present"	
		End If
	End If
	GetChannelAndUrl = TempReturnString
End Function

Function GetUserSelection()
	Dim UsersChoices 
	For Each choice In Channels
		UsersChoices = UsersChoices + choice + ". " + Channels(choice).name + vbCrlf
	Next
	
	GetUserSelection = InputBox("Please select the number of the channel you would like: " + vbCrLf + UsersChoices, "Get Selection")
End Function

Function UpdateChannelAndVersion(selection)
	Dim chosenUpdateUrl
	Dim version
	Dim UpdateUrlToUse
	
	If selection <> "" Then
		If IsNumeric(selection) Then
			chosenUpdateUrl = Channels(selection).url
			UpdateUrlToUse = chosenUpdateUrl
			version = GetVersion(chosenUpdateUrl)
	
			'update UpdateUrl here
			WriteRegValue(UpdateUrlToUse)
	
	
			'run update here
			RunUpdate(version)
		Else
			WScript.Echo "No item chosen, exiting with no further action"
		End If	
	Else
		WScript.Echo "No item chosen, exiting with no further action"
	End If
	
End Function

Function GetVersion(chsnUpdateUrl)
If InStr(LCase(chsnUpdateUrl), "http") Then
	chsnUpdateUrl = chsnUpdateUrl + "/Office/Data/v32.cab"
	dim xHttp: Set xHttp = createobject("Msxml2.XMLHttp.6.0")
	dim bStrm: Set bStrm = createobject("Adodb.Stream")
	Dim TempFolder: Set TempFolder = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
	Dim fso: Set fso = createobject("Scripting.FileSystemObject")
	
	'check for and delete existing files
	If(fso.FileExists(TempFolder+"\v32.cab")) Then
		fso.DeleteFile TempFolder+"\v32.cab"
	End If
	If(fso.FileExists(TempFolder+"\v32.hash")) Then
		fso.DeleteFile TempFolder+"\v32.hash"
	End If
	If(fso.FileExists(TempFolder+"\VersionDescriptor.xml")) Then
		fso.DeleteFile TempFolder+"\VersionDescriptor.xml"
	End If
	
	
	xHttp.Open "GET", chsnUpdateUrl, False
	xHttp.Send

	with bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile TempFolder+"\v32.cab", 2 '//overwrite
	end With
	
	Dim objShell 
	Set objShell = CreateObject("Shell.Application")
	Dim FilesInObject
	Set FilesInObject = objShell.NameSpace(TempFolder+"\v32.cab").Items	
	objShell.NameSpace(TempFolder+"").CopyHere FilesInObject
	Set objShell = Nothing
	
	Dim objXMLDoc
	Dim Node
	Dim Attribute
	
	Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False 
	objXMLDoc.load(TempFolder+"\VersionDescriptor.xml")

	Set Node = objXMLDoc.selectNodes("//Version/Available")
	GetVersion = Node(0).getAttribute("Build")
Else
	Dim fileso  
	Set fileso = CreateObject("Scripting.FileSystemObject")
	Dim exists
	
	exists = fileso.FolderExists(chsnUpdateUrl)
	If (exists) Then
	For Each subfolder In fileso.GetFolder(chsnUpdateUrl+"\Office\Data").SubFolders		
		GetVersion = Mid(subfolder, InStrRev(subfolder, CStr("\"))+1)
	Next
	End IF
End If
End Function

Function RunUpdate(VersionToUpdateTo)
	Dim ClientPath
	Dim oShell
	Set oShell = WScript.CreateObject ("WScript.Shell")
	
	ClientPath = FindClientPath
	ClientPath = ClientPath + "\OfficeC2RClient.exe"
	Dim TextToRun
	TextToRun = """"+ClientPath+"""" + " /update user displaylevel=false forceappshutdown=true updatepromptuser=false updatetoversion=" + VersionToUpdateTo 
	oShell.Run TextToRun
	Set oShell = Nothing
End Function

Class OfficeUpdateChannel
	Public name, url
End Class
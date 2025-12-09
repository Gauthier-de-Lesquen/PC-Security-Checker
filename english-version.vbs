Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colUsers = objWMI.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = True And Disabled = False")
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim UserCount
UserCount = 0
For Each user In colUsers
    If Not user.Name = "DefaultAccount" And Not user.Name = "Guest" Then
	UserCount = UserCount + 1
    End If
Next

If UserCount > 1 Then
    Dim Msg, Title, listusers
    listusers = ""
    Title = "hacker detection"
    For Each user In colUsers
        If Not user.Name = "DefaultAccount" And Not user.Name = "Guest" Then
	    listusers = listusers & user.Name & ", "
	End If
    Next
    Msg = "the users '" & listusers & "' have access to your PC. they're probably Hackers." & vbCrLf & "do you want to check their identity?" 
    
    Response = MsgBox(Msg, vbYesNo, Title)

    If Response = vbYes Then
        shell.Run "netplwiz"
    End If
Else
    MsgBox "No Hacker has been detected", vbOKOnly, "Hacker detection"
End If

folderPath = shell.ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup")

If fso.FolderExists(folderPath) Then
    Set folder = fso.GetFolder(folderPath)

    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        MsgBox "No suspect startup apps detected", vbOKOnly, "Startup apps detection"
    Else
        Msg = "Some suspect startup programs have been detected." & vbCrLf & "do you want to check them?"
	Title = "Startup apps detection"
	Response = MsgBox(Msg, vbYesNo, Title)

	If Response = vbYes Then
	    shell.Run "shell:startup"
	End If
    End If
End If

Dim registryPath, WshShell, objRegistry, valueNames, value, i

registryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
Set WshShell = CreateObject("WScript.Shell")
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

' Liste des valeurs de la clé Run
objRegistry.EnumValues &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", valueNames, Nothing

If IsArray(valueNames) Then
    For i = 0 To UBound(valueNames)
        objRegistry.GetStringValue &H80000002, _
            "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", valueNames(i), value
        
        If Not IsNull(value) Then
            ' Vérifie si le chemin ne contient pas "C:\Windows"
            If InStr(1, LCase(value), "c:\windows", vbTextCompare) = 0 Then
                Msg = "Some Startup apps have been detected on the regedit's key HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & valueNames(i) & " => " & value
		MsgBox Msg, vbOKOnly, "Startup apps detection"
	    Else
		MsgBox "Still no suspect startup apps detected", vbOKOnly, "Startup apps detection"
            End If
        End If
    Next
End If

shell.Run "mrt.exe /N", 1, True   ' 1 = fenêtre visible

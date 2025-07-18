' SaveWiFiProfiles.vbs
Option Explicit

Dim objShell, objFSO, objOutputFile, strCommand
Dim scriptPath, scriptFolder, outputPath
Dim strProfileList, arrProfiles, i, strProfileName, strProfileDetails

' Create required objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the folder path where the script is located
scriptPath = WScript.ScriptFullName
scriptFolder = objFSO.GetParentFolderName(scriptPath)

' Output file path in the same folder
outputPath = scriptFolder & "\wifi_profiles.txt"

Set objOutputFile = objFSO.CreateTextFile(outputPath, True)

' Step 1: Get list of Wi-Fi profiles
strCommand = "cmd /c netsh wlan show profiles"
Set strProfileList = objShell.Exec(strCommand).StdOut

Dim profileOutput : profileOutput = ""
Do Until strProfileList.AtEndOfStream
    profileOutput = profileOutput & strProfileList.ReadLine() & vbCrLf
Loop

' Extract profile names
Dim regex, matches
Set regex = New RegExp
With regex
    .Global = True
    .IgnoreCase = True
    .Pattern = "All User Profile\s*:\s*(.+)"
End With

Set matches = regex.Execute(profileOutput)

objOutputFile.WriteLine "=== Wi-Fi Profiles and Details ==="
objOutputFile.WriteLine "Exported on: " & Now
objOutputFile.WriteLine String(50, "-")

' Step 2: Loop through profiles and extract details
For i = 0 To matches.Count - 1
    strProfileName = Trim(matches(i).SubMatches(0))
    objOutputFile.WriteLine vbCrLf & "Profile: " & strProfileName
    objOutputFile.WriteLine String(40, "=")

    strCommand = "cmd /c netsh wlan show profile name=""" & strProfileName & """ key=clear"
    Set strProfileDetails = objShell.Exec(strCommand).StdOut

    Do Until strProfileDetails.AtEndOfStream
        objOutputFile.WriteLine strProfileDetails.ReadLine()
    Loop
Next

objOutputFile.Close

MsgBox "Wi-Fi profiles have been saved to:" & vbCrLf & outputPath, vbInformation, "Export Complete"

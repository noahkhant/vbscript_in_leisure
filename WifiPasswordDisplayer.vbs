Dim objShell, objExec, strProfile, arrProfiles
Dim strCommand, strResult, strPassword, count

Set objShell = CreateObject("WScript.Shell")

' Get the list of Wi-Fi profiles
strCommand = "cmd /c netsh wlan show profiles"
Set objExec = objShell.Exec(strCommand)
strResult = objExec.StdOut.ReadAll()

' Extract valid profile names and ignore non-profile sections
Set arrProfiles = CreateObject("Scripting.Dictionary")
For Each strProfile In Split(strResult, vbCrLf)
    ' Look for "All User Profile" which indicates a valid Wi-Fi profile
    If InStr(strProfile, "All User Profile") > 0 Then
        strProfile = Trim(Split(strProfile, ":")(1))
        
        ' Skip empty profile names and any other sections (like "Cost Setting")
        If Len(strProfile) > 0 And Not InStr(strProfile, "Cost") > 0 Then
            arrProfiles.Add strProfile, ""
        End If
    End If
Next

count = 0

' Loop through each profile and get its password
For Each strProfile In arrProfiles.Keys
    If count = 3 Then Exit For ' Stop after displaying 3 profiles
    
    ' Get profile details and extract the password
    strCommand = "cmd /c netsh wlan show profile name=""" & strProfile & """ key=clear"
    Set objExec = objShell.Exec(strCommand)
    strResult = objExec.StdOut.ReadAll()
    
    ' Search specifically for the Key Content line, ignoring other sections
    If InStr(strResult, "Key Content") > 0 Then
        strPassword = Trim(Split(Split(strResult, "Key Content")(1), ":")(1))
        strPassword = Trim(Split(strPassword, vbCrLf)(0))
    Else
        strPassword = "(No Password/Hidden)"
    End If
    
    ' Output the profile name and password
    WScript.Echo "Wi-Fi Profile: " & strProfile & vbCrLf & "Password: " & strPassword & vbCrLf
    count = count + 1
Next
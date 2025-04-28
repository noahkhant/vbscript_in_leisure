Dim objFSO, objFolder, objFile
Dim folderPath, fileTypeDict, fileType, fileCount, fileExtension
Dim result

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set fileTypeDict = CreateObject("Scripting.Dictionary")

' Ask the user to enter the folder path
folderPath = InputBox("Enter the folder path to collect file types:", "Folder Path")

' Check if the folder exists
If objFSO.FolderExists(folderPath) Then
    Set objFolder = objFSO.GetFolder(folderPath)

    ' Loop through each file in the folder
    For Each objFile In objFolder.Files
        ' Get the file extension and ensure it has a period (e.g., .html)
        fileExtension = LCase(objFSO.GetExtensionName(objFile.Name))
        If fileExtension <> "" Then
            fileExtension = "." & fileExtension
        End If
        
        ' Add the file type to the dictionary or increment the count
        If fileTypeDict.Exists(fileExtension) Then
            fileTypeDict(fileExtension) = fileTypeDict(fileExtension) + 1
        Else
            fileTypeDict.Add fileExtension, 1
        End If
    Next
    ' Initialize the result string
    result = "File types in folder: " & folderPath & vbCrLf & String(40, "-") & vbCrLf
    
    ' Build the result string with the desired format
    For Each fileType In fileTypeDict.Keys
        fileCount = fileTypeDict(fileType)
        
        ' Add each result to the string, with appropriate pluralization
        If fileCount = 1 Then
            result = result & fileType & " : " & fileCount & " file" & vbCrLf
        Else
            result = result & fileType & " : " & fileCount & " files" & vbCrLf
        End If
    Next
    WScript.Echo result
Else
    WScript.Echo "The folder path '" & folderPath & "' does not exist."
End If

' Clean up
Set fileTypeDict = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
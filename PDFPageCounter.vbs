Main()

Sub Main()
    Dim fso, f
    Dim path, name
    Dim pdfFile, fileContent

    path = InputBox("Location folder path of target pdf file")
    name = InputBox("Pdf file name")
    pdfFile = path & "\" & name

    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(pdfFile, 1)

    fileContent = f.ReadAll
    Call countPage(fileContent)
End Sub

Function countPage(content)
    Dim regEx, Match, Matches, count, pattern

    pattern = "/Count\s+(\d+)"
    count = 0

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = pattern
    regEx.IgnoreCase = True
    regEx.Global = True
    Set Matches = regEx.Execute(content)

    If Matches.Count > 0 Then
        For Each Match In Matches
            If CInt(Match.SubMatches(0)) > count Then
                count = CInt(Match.SubMatches(0))
            End If
        Next
    End If
    MsgBox("Total pages : " & count)
End Function

Dim objWord, objDoc, objFSO, objFile
Dim objInlineShape, objShape
Dim imgDict, filePath, docPath
Dim line, arrData, imgID, newText

' Ask user for Word file and CSV file
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objFSO = CreateObject("Scripting.FileSystemObject")

docPath = InputBox("Enter the full path of the Word document (.docm):", "Select Word File")
If docPath = "" Then WScript.Quit

filePath = objFSO.GetParentFolderName(docPath) & "\alt_text.csv"
If Not objFSO.FileExists(filePath) Then
    MsgBox "CSV file not found: " & filePath
    WScript.Quit
End If

Set objDoc = objWord.Documents.Open(docPath)
Set objFile = objFSO.OpenTextFile(filePath, 1)

' Read CSV file into dictionary
Set imgDict = CreateObject("Scripting.Dictionary")

' Skip header line
objFile.ReadLine

Do Until objFile.AtEndOfStream
    line = objFile.ReadLine
    arrData = Split(line, ",")
    If UBound(arrData) = 1 Then
        imgDict(arrData(0)) = arrData(1)
    End If
Loop
objFile.Close

' Update inline images
imgCounter = 1
For Each objInlineShape In objDoc.InlineShapes
    imgID = "Inline_" & imgCounter
    If imgDict.Exists(imgID) Then
        objInlineShape.AlternativeText = imgDict(imgID)
    End If
    imgCounter = imgCounter + 1
Next

' Update floating images
For Each objShape In objDoc.Shapes
    imgID = "Floating_" & imgCounter
    If imgDict.Exists(imgID) Then
        objShape.AlternativeText = imgDict(imgID)
    End If
    imgCounter = imgCounter + 1
Next

' Save document
objDoc.Save
objDoc.Close False
objWord.Quit

MsgBox "Alt text updated successfully in " & docPath

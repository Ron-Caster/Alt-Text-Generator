Dim objWord, objDoc, objFSO, objFile
Dim objInlineShape, objShape
Dim altText, imgCounter
Dim filePath, docPath

' Ask user to select Word file
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objFSO = CreateObject("Scripting.FileSystemObject")

docPath = InputBox("Enter the full path of the Word document (.docm):", "Select Word File")
If docPath = "" Then WScript.Quit

Set objDoc = objWord.Documents.Open(docPath)

' Create CSV file to store extracted alt-text
filePath = objFSO.GetParentFolderName(docPath) & "\alt_text.csv"
Set objFile = objFSO.CreateTextFile(filePath, True)

' Write CSV header
objFile.WriteLine "ImageID,AltText"

' Extract alt-text from inline images
imgCounter = 1
For Each objInlineShape In objDoc.InlineShapes
    If objInlineShape.AlternativeText = "" Then
        altText = "Missing"
    Else
        altText = objInlineShape.AlternativeText
    End If
    objFile.WriteLine "Inline_" & imgCounter & "," & altText
    imgCounter = imgCounter + 1
Next

' Extract alt-text from floating images
For Each objShape In objDoc.Shapes
    If objShape.AlternativeText = "" Then
        altText = "Missing"
    Else
        altText = objShape.AlternativeText
    End If
    objFile.WriteLine "Floating_" & imgCounter & "," & altText
    imgCounter = imgCounter + 1
Next

' Cleanup
objFile.Close
objDoc.Close False
objWord.Quit

MsgBox "Alt text extracted successfully! File saved as: " & filePath

## 27 January, 2025

- Got assigned to create a program for adding alt-text to word documents.
- Started creating LaTeX report for the program.
- Covered Program Planning.
- Tech used:
    - Python (Programming Language)
    - Groq (Free LPU Host)
    - Llama Vision (Model)
    - Openpyxl (Document Handling)
- Learnt and started using LaTeX for creating the report.

## 28 January, 2025

- Tried to create a python script to extract the alt-texts from the word file.
- Found out the word file provided wasn't having a single alt-text in the images.
- There's no need of an excel file preparation in the entire project.
- Removed excel file management from project as it's out of scope, unnecessary and not to the point of the successful completion of the project.
- Found VBA can be used to access alt-texts somehow.
    - [Link to Source](https://answers.microsoft.com/en-us/msoffice/forum/all/need-to-extract-alt-text-for-images-and-tables-in/0b46c9a6-4a4a-4243-b53d-e17adc824699)
- Used AI to Summarize the website content.
- Failed multiple times and reassured alt-text can't be extracted using python-docx or any other libraries in python as of now.
- Searching for other ways to mitigate/use VBA to resolve the issue.
- Found out input should be ".docm" (Macro Enabled Doc File) instead of ".docx"

### V1.0 Program

- The VBA code for the purpose:
  
    <details>
    <summary style="color: turquoise;">Click to see the VBA code!</summary>
  
  ```vba
  Sub ExportAltText()

      Dim strPictures As String
      Dim docPictures As Document
      Dim docTranslate As Document
      Dim objInlinePic As InlineShape
      Dim objFloatPic As Shape
      Dim objTable As Table ' in docPictures
      Dim strTblAlt As String
      Dim tblTranslate1 As Table ' in docTranslate
      Dim tblTranslate2 As Table
      Dim tblTranslate3 As Table
      Dim tblTranslate4 As Table
      Dim tblLoop As Table
      Dim rowCurrent As Row
      Dim oRg As Range

      MsgBox "In the next dialog, select the file containing " & _
             "the pictures whose alt text will be translated."

      strPictures = GetFileName()

      If strPictures = "" Then Exit Sub

      On Error GoTo BadInputFile
      Set docPictures = Documents.Open(FileName:=strPictures)
      Set docTranslate = Documents.Add

      With docTranslate
          ' Set up header and footer in translation document
          .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
              "Alt Text of " & docPictures.FullName
          Set oRg = .Sections(1).Footers(wdHeaderFooterPrimary).Range
          oRg.Text = vbTab
          oRg.Collapse wdCollapseEnd
          .Fields.Add Range:=oRg, Type:=wdFieldPage, PreserveFormatting:=False

          ' Create four 2x2 tables
          Set oRg = .Range
          oRg.InsertAfter "Inline Pictures" & vbCr
          Set oRg = .Range
          oRg.Collapse wdCollapseEnd
          Set tblTranslate1 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=2)

          Set oRg = .Range
          oRg.InsertAfter "Floating Pictures" & vbCr
          Set oRg = .Range
          oRg.Collapse wdCollapseEnd
          Set tblTranslate2 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=2)

          Set oRg = .Range
          oRg.InsertAfter "Tables" & vbCr
          Set oRg = .Range
          oRg.Collapse wdCollapseEnd
          Set tblTranslate3 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=2)

          Set oRg = .Range
          oRg.InsertAfter "Author and Title" & vbCr
          Set oRg = .Range
          oRg.Collapse wdCollapseEnd
          Set tblTranslate4 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=2)

          ' Save the docPictures path for future use
          .Variables("docPictures").Value = docPictures.FullName
      End With

      ' Set up the tables with headers
      For Each tblLoop In docTranslate.Tables
          With tblLoop
              .Cell(1, 1).Range.Text = "Original Alt Text"
              .Cell(1, 2).Range.Text = "Translated Alt Text"
              .Rows(1).Range.Font.Bold = True
              .Rows(1).HeadingFormat = True
              .Borders.InsideColor = wdColorAutomatic
              .Borders.InsideLineStyle = wdLineStyleSingle
              .Borders.OutsideColor = wdColorAutomatic
              .Borders.OutsideLineStyle = wdLineStyleSingle
          End With
      Next tblLoop

      ' Export alt text for inline pictures
      On Error Resume Next
      For Each objInlinePic In docPictures.InlineShapes
          If objInlinePic.AlternativeText <> "" Then
              tblTranslate1.Rows.Last.Cells(1).Range.Text = objInlinePic.AlternativeText
              tblTranslate1.Rows.Add
          End If
      Next objInlinePic
      tblTranslate1.Rows.Last.Delete

      ' Export alt text for floating pictures
      For Each objFloatPic In docPictures.Shapes
          If objFloatPic.AlternativeText <> "" Then
              tblTranslate2.Rows.Last.Cells(1).Range.Text = objFloatPic.AlternativeText
              tblTranslate2.Rows.Add
          End If
      Next objFloatPic
      tblTranslate2.Rows.Last.Delete

      ' Export alt text for tables
      For Each objTable In docPictures.Tables
          strTblAlt = ""
          If objTable.Descr <> "" Then strTblAlt = objTable.Descr
          If objTable.Title <> "" Then strTblAlt = objTable.Title & vbCr & strTblAlt
          If Len(strTblAlt) > 1 Then
              tblTranslate3.Rows.Last.Cells(1).Range.Text = strTblAlt
              tblTranslate3.Rows.Add
          End If
      Next objTable
      tblTranslate3.Rows.Last.Delete

      ' Export author and title
      With tblTranslate4
          .Rows.Last.Cells(1).Range.Text = "Author"
          .Rows.Add
          .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Author").Value
          .Rows.Add
          .Rows.Last.Cells(1).Range.Text = "Title"
          .Rows.Add
          .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Title").Value
      End With

      ' Save the new translation document
      docTranslate.SaveAs FileName:=Replace(strPictures, ".doc", " Alt Text.doc")
      MsgBox "Alt text has been exported and saved as 'Alt Text.doc'."
      docPictures.Close wdDoNotSaveChanges

      Exit Sub

  BadInputFile:
      MsgBox "The file " & strPictures & " could not be opened." & _
             vbCr & "Error " & Err.Number & vbCr & Err.Description
  End Sub

  ' Helper functions
  Function GetFileName() As String
      Dim dlg As FileDialog
      Set dlg = Application.FileDialog(msoFileDialogFilePicker)
      If dlg.Show <> -1 Then
          GetFileName = ""
      Else
          GetFileName = dlg.SelectedItems(1)
      End If
  End Function 
  ```
#### Steps to Follow:

- Open the *Macro Enabled Word Document*.
- Press *ALT + F11* to open *Microsoft Visual Basic Console* or (*Fn + Alt + F11*).
- Create a new module:
    - ![Image showing Module as the second dropdown in insert userform]({916839D4-03BD-4BC7-B0CA-845C80BA7A35}.png)
- In the created new module, paste the VBA Code.
- Press *F5* to run the VBA Code.
- Select the *.docm* file to extract the alt-text from the file dialog.
- Open the *.docm* file and press *ALT + F8*.
- Select *ExportAltText* from the Macros.
- Click Run.
- In the next dialogue box select the *.docm* file.
- Then *Filename Alt Text.docm* will be created in the same folder.
- *Filename Alt Text.docm* will contain the extracted alternative text.

### V1.1 Improvements

- Assign names to images and add it along with alt-text.
- If possible, get page numbers of the alt-texts extracted and add it in a new column.
- Updated VBA Code:

    <details>
    <summary style="color: turquoise;">Click to see the VBA code!</summary>
  
    ```vba
        Sub ExportAltTextWithNamesAndPageNumbers()

        Dim strPictures As String
        Dim docPictures As Document
        Dim docTranslate As Document
        Dim objInlinePic As InlineShape
        Dim objFloatPic As Shape
        Dim objTable As Table
        Dim strTblAlt As String
        Dim tblTranslate1 As Table
        Dim tblTranslate2 As Table
        Dim tblTranslate3 As Table
        Dim tblTranslate4 As Table
        Dim tblLoop As Table
        Dim rowCurrent As Row
        Dim oRg As Range
        Dim picIndex As Integer
        Dim pageNum As String

        MsgBox "In the next dialog, select the file containing the pictures whose alt text will be exported."

        strPictures = GetFileName()

        If strPictures = "" Then Exit Sub

        On Error GoTo BadInputFile
        Set docPictures = Documents.Open(FileName:=strPictures)
        Set docTranslate = Documents.Add

        With docTranslate
            ' Set up header and footer in translation document
            .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
                "Alt Text of " & docPictures.FullName
            Set oRg = .Sections(1).Footers(wdHeaderFooterPrimary).Range
            oRg.Text = vbTab
            oRg.Collapse wdCollapseEnd
            .Fields.Add Range:=oRg, Type:=wdFieldPage, PreserveFormatting:=False

            ' Create three-column tables (Image Name, Original Alt Text, Page Number)
            Set oRg = .Range
            oRg.InsertAfter "Inline Pictures" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate1 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Floating Pictures" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate2 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Tables" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate3 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Author and Title" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate4 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            ' Save the docPictures path for future use
            .Variables("docPictures").Value = docPictures.FullName
        End With

        ' Set up the tables with headers
        For Each tblLoop In docTranslate.Tables
            With tblLoop
                .Cell(1, 1).Range.Text = "Image Name"
                .Cell(1, 2).Range.Text = "Original Alt Text"
                .Cell(1, 3).Range.Text = "Page Number"
                .Rows(1).Range.Font.Bold = True
                .Rows(1).HeadingFormat = True
                .Borders.InsideColor = wdColorAutomatic
                .Borders.InsideLineStyle = wdLineStyleSingle
                .Borders.OutsideColor = wdColorAutomatic
                .Borders.OutsideLineStyle = wdLineStyleSingle
            End With
        Next tblLoop

        ' Export alt text for inline pictures
        picIndex = 1
        For Each objInlinePic In docPictures.InlineShapes
            If objInlinePic.AlternativeText <> "" Then
                pageNum = objInlinePic.Range.Information(wdActiveEndAdjustedPageNumber)
                tblTranslate1.Rows.Add
                tblTranslate1.Rows.Last.Cells(1).Range.Text = "Inline Picture " & picIndex
                tblTranslate1.Rows.Last.Cells(2).Range.Text = objInlinePic.AlternativeText
                tblTranslate1.Rows.Last.Cells(3).Range.Text = pageNum
                picIndex = picIndex + 1
            End If
        Next objInlinePic
        tblTranslate1.Rows.Last.Delete

        ' Export alt text for floating pictures
        picIndex = 1
        For Each objFloatPic In docPictures.Shapes
            If objFloatPic.AlternativeText <> "" Then
                pageNum = objFloatPic.Anchor.Information(wdActiveEndAdjustedPageNumber)
                tblTranslate2.Rows.Add
                tblTranslate2.Rows.Last.Cells(1).Range.Text = "Floating Picture " & picIndex
                tblTranslate2.Rows.Last.Cells(2).Range.Text = objFloatPic.AlternativeText
                tblTranslate2.Rows.Last.Cells(3).Range.Text = pageNum
                picIndex = picIndex + 1
            End If
        Next objFloatPic
        tblTranslate2.Rows.Last.Delete

        ' Export alt text for tables
        picIndex = 1
        For Each objTable In docPictures.Tables
            strTblAlt = ""
            If objTable.Descr <> "" Then strTblAlt = objTable.Descr
            If objTable.Title <> "" Then strTblAlt = objTable.Title & vbCr & strTblAlt
            If Len(strTblAlt) > 1 Then
                pageNum = objTable.Range.Information(wdActiveEndAdjustedPageNumber)
                tblTranslate3.Rows.Add
                tblTranslate3.Rows.Last.Cells(1).Range.Text = "Table " & picIndex
                tblTranslate3.Rows.Last.Cells(2).Range.Text = strTblAlt
                tblTranslate3.Rows.Last.Cells(3).Range.Text = pageNum
                picIndex = picIndex + 1
            End If
        Next objTable
        tblTranslate3.Rows.Last.Delete

        ' Export author and title
        With tblTranslate4
            .Rows.Last.Cells(1).Range.Text = "Author"
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Author").Value
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = "Title"
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Title").Value
        End With

        ' Save the new translation document
        docTranslate.SaveAs FileName:=Replace(strPictures, ".doc", " Alt Text.doc")
        MsgBox "Alt text, image names, and page numbers have been exported and saved as 'Alt Text.doc'."
        docPictures.Close wdDoNotSaveChanges

        Exit Sub

    BadInputFile:
        MsgBox "The file " & strPictures & " could not be opened." & _
            vbCr & "Error " & Err.Number & vbCr & Err.Description
    End Sub

    ' Helper functions
    Function GetFileName() As String
        Dim dlg As FileDialog
        Set dlg = Application.FileDialog(msoFileDialogFilePicker)
        If dlg.Show <> -1 Then
            GetFileName = ""
        Else
            GetFileName = dlg.SelectedItems(1)
        End If
    End Function
    ```

#### What This Code Does

1. Image Names:

    - Inline and floating images are automatically assigned sequential names, e.g., Inline Picture 1, Floating Picture 1, etc.
    - Tables are named sequentially as Table 1, Table 2, etc.

2. Page Numbers:

    - Retrieves the page number where each image or table appears in the document using wdActiveEndAdjustedPageNumber.
    - Adds the page number to a new column in the output document.

3. Output Structure:

    - Each table now has three columns:
        1. Image/Table Name
        2. Original Alt Text
        3. Page Number

4. Saves Output:

    - Creates a new document (OriginalFileName Alt Text.doc) with this updated structure.

#### Expected Output

- A Word document with the following tables:
    1. Inline Pictures:
        - Image Name, Alt Text, and Page Number.
    2. Floating Pictures:
        - Image Name, Alt Text, and Page Number.
    3. Tables:
        - Table Name, Alt Text (Title + Description), and     Page Number.
    4. Author and Title:
        - Metadata for the document.

### V1.2 Improvements

- The images without alt-text aren't considered in the current code.
- *Prompt given for updating code:*
    ```
    The images without alt-text aren't considered in the current code.

    So, this code pulls the images with alt-texts only. Therefore, update the code such that all the images (including those without alt-text) are also pulled and added to the list of inline images. In the column of Original Alt Text put "Missing" for the images without alt-text.
- Updated VBA Code:

    <details>
    <summary style="color: turquoise;">Click to see the VBA code!</summary>
  
    ```vba
        Sub ExportAllImagesWithAltTextAndPageNumbers()

        Dim strPictures As String
        Dim docPictures As Document
        Dim docTranslate As Document
        Dim objInlinePic As InlineShape
        Dim objFloatPic As Shape
        Dim objTable As Table
        Dim strTblAlt As String
        Dim tblTranslate1 As Table
        Dim tblTranslate2 As Table
        Dim tblTranslate3 As Table
        Dim tblTranslate4 As Table
        Dim tblLoop As Table
        Dim rowCurrent As Row
        Dim oRg As Range
        Dim picIndex As Integer
        Dim pageNum As String

        MsgBox "In the next dialog, select the file containing the pictures whose alt text will be exported."

        strPictures = GetFileName()

        If strPictures = "" Then Exit Sub

        On Error GoTo BadInputFile
        Set docPictures = Documents.Open(FileName:=strPictures)
        Set docTranslate = Documents.Add

        With docTranslate
            ' Set up header and footer in translation document
            .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
                "Alt Text of " & docPictures.FullName
            Set oRg = .Sections(1).Footers(wdHeaderFooterPrimary).Range
            oRg.Text = vbTab
            oRg.Collapse wdCollapseEnd
            .Fields.Add Range:=oRg, Type:=wdFieldPage, PreserveFormatting:=False

            ' Create three-column tables (Image Name, Original Alt Text, Page Number)
            Set oRg = .Range
            oRg.InsertAfter "Inline Pictures" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate1 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Floating Pictures" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate2 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Tables" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate3 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            Set oRg = .Range
            oRg.InsertAfter "Author and Title" & vbCr
            Set oRg = .Range
            oRg.Collapse wdCollapseEnd
            Set tblTranslate4 = .Tables.Add(Range:=oRg, numrows:=2, numcolumns:=3)

            ' Save the docPictures path for future use
            .Variables("docPictures").Value = docPictures.FullName
        End With

        ' Set up the tables with headers
        For Each tblLoop In docTranslate.Tables
            With tblLoop
                .Cell(1, 1).Range.Text = "Image Name"
                .Cell(1, 2).Range.Text = "Original Alt Text"
                .Cell(1, 3).Range.Text = "Page Number"
                .Rows(1).Range.Font.Bold = True
                .Rows(1).HeadingFormat = True
                .Borders.InsideColor = wdColorAutomatic
                .Borders.InsideLineStyle = wdLineStyleSingle
                .Borders.OutsideColor = wdColorAutomatic
                .Borders.OutsideLineStyle = wdLineStyleSingle
            End With
        Next tblLoop

        ' Export all inline pictures
        picIndex = 1
        For Each objInlinePic In docPictures.InlineShapes
            pageNum = objInlinePic.Range.Information(wdActiveEndAdjustedPageNumber)
            tblTranslate1.Rows.Add
            tblTranslate1.Rows.Last.Cells(1).Range.Text = "Inline Picture " & picIndex
            
            If objInlinePic.AlternativeText <> "" Then
                tblTranslate1.Rows.Last.Cells(2).Range.Text = objInlinePic.AlternativeText
            Else
                tblTranslate1.Rows.Last.Cells(2).Range.Text = "Missing"
            End If
            
            tblTranslate1.Rows.Last.Cells(3).Range.Text = pageNum
            picIndex = picIndex + 1
        Next objInlinePic
        tblTranslate1.Rows.Last.Delete

        ' Export all floating pictures
        picIndex = 1
        For Each objFloatPic In docPictures.Shapes
            pageNum = objFloatPic.Anchor.Information(wdActiveEndAdjustedPageNumber)
            tblTranslate2.Rows.Add
            tblTranslate2.Rows.Last.Cells(1).Range.Text = "Floating Picture " & picIndex
            
            If objFloatPic.AlternativeText <> "" Then
                tblTranslate2.Rows.Last.Cells(2).Range.Text = objFloatPic.AlternativeText
            Else
                tblTranslate2.Rows.Last.Cells(2).Range.Text = "Missing"
            End If
            
            tblTranslate2.Rows.Last.Cells(3).Range.Text = pageNum
            picIndex = picIndex + 1
        Next objFloatPic
        tblTranslate2.Rows.Last.Delete

        ' Export alt text for tables
        picIndex = 1
        For Each objTable In docPictures.Tables
            strTblAlt = ""
            If objTable.Descr <> "" Then strTblAlt = objTable.Descr
            If objTable.Title <> "" Then strTblAlt = objTable.Title & vbCr & strTblAlt
            If Len(strTblAlt) > 1 Then
                pageNum = objTable.Range.Information(wdActiveEndAdjustedPageNumber)
                tblTranslate3.Rows.Add
                tblTranslate3.Rows.Last.Cells(1).Range.Text = "Table " & picIndex
                tblTranslate3.Rows.Last.Cells(2).Range.Text = strTblAlt
                tblTranslate3.Rows.Last.Cells(3).Range.Text = pageNum
                picIndex = picIndex + 1
            End If
        Next objTable
        tblTranslate3.Rows.Last.Delete

        ' Export author and title
        With tblTranslate4
            .Rows.Last.Cells(1).Range.Text = "Author"
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Author").Value
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = "Title"
            .Rows.Add
            .Rows.Last.Cells(1).Range.Text = docPictures.BuiltInDocumentProperties("Title").Value
        End With

        ' Save the new translation document
        docTranslate.SaveAs FileName:=Replace(strPictures, ".doc", " Alt Text.doc")
        MsgBox "Alt text, image names, and page numbers (including missing ones) have been exported and saved as 'Alt Text.doc'."
        docPictures.Close wdDoNotSaveChanges

        Exit Sub

    BadInputFile:
        MsgBox "The file " & strPictures & " could not be opened." & _
            vbCr & "Error " & Err.Number & vbCr & Err.Description
    End Sub

    ' Helper functions
    Function GetFileName() As String
        Dim dlg As FileDialog
        Set dlg = Application.FileDialog(msoFileDialogFilePicker)
        If dlg.Show <> -1 Then
            GetFileName = ""
        Else
            GetFileName = dlg.SelectedItems(1)
        End If
    End Function
    ```
#### What This Update Does

1. Includes All Images:

    - Both inline and floating images are added to the output, regardless of whether they have alt-text.
    - Images without alt-text will show "Missing" in the Original Alt Text column.

2. Adds Page Numbers:

    - Retrieves the page number for all images, even those without alt-text.

3. Output Structure:

    - The tables for inline and floating images now include:
        - Image Name: Sequentially numbered (e.g., Inline Picture 1, Floating Picture 1).
        - Original Alt Text: Displays "Missing" for images without alt-text.
        - Page Number: Page where the image appears.
4. Saves Output:
    - Creates a Word document (OriginalFileName Alt Text.doc) with the updated data.

#### Expected Output

- A Word document with the following sections:
    1. Inline Pictures:
        - Lists all inline images, their alt-text (or "Missing"), and page numbers.
    2. Floating Pictures:
        - Lists all floating images, their alt-text (or "Missing"), and page numbers.
    3. Tables:
        - Includes table names, alt-text (from Title or Description), and page numbers.
    4. Author and Title:
        - Metadata from the document.

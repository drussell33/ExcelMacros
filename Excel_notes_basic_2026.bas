' --------------------------------------------------------
' ---------Search Everything in sheet button--------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub SearchAndHighlightNewest()
    Dim searchText As String
    Dim cell As range
    Dim ws As Worksheet
    Dim shp As shape
    Dim foundPos As Long
    Dim totalChars As Long
    Dim cellText As String
    Dim startPos As Long
    Dim lengthMatch As Long
    Dim hl As Hyperlink
    Dim totalCells As Long, totalShapes As Long
    Dim currentItem As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Count total cells and shapes to display the progress correctly
    totalCells = ws.UsedRange.Cells.Count
    totalShapes = ws.Shapes.Count
    currentItem = 0

    ' Disable user interaction while macro runs
    Application.ScreenUpdating = False
    Application.Interactive = False
    Application.EnableEvents = False

    ' ---- Clear previous highlights in cells ----
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            ' Reset entire cell's font color to black and remove bold, but preserve hyperlinks
            With cell.Font
                .Color = RGB(0, 0, 0) ' Set text color back to black
                .Bold = False ' Remove bold formatting
            End With
            ' Reapply hyperlink formatting if the cell contains a hyperlink
            If cell.Hyperlinks.Count > 0 Then
                For Each hl In cell.Hyperlinks
                    With hl.range.Font
                        .Color = RGB(5, 99, 193) ' Default hyperlink blue color
                        .Underline = xlUnderlineStyleSingle
                    End With
                Next hl
            End If
        End If
        ' Update progress bar for cell processing
        currentItem = currentItem + 1
        UpdateStatusBar currentItem, totalCells + totalShapes
    Next cell

    ' ---- Clear previous highlights in shapes ----
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            ' Reset the font color in the shape's text to black
            With shp.TextFrame2.TextRange
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' Set text color back to black
            End With
        End If
        ' Update progress bar for shape processing
        currentItem = currentItem + 1
        UpdateStatusBar currentItem, totalCells + totalShapes
    Next shp

    ' Ask the user for the search text
    searchText = InputBox("Enter the text to search:")

    ' Exit if search text is empty
    If searchText = "" Then
        Application.StatusBar = False ' Reset status bar
        Application.ScreenUpdating = True
        Application.Interactive = True
        Application.EnableEvents = True
        Exit Sub
    End If

    ' ---- Search text within all cells ----
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            cellText = cell.Value
            startPos = InStr(1, cellText, searchText, vbTextCompare) ' Find the start position of the match
            If startPos > 0 Then
                ' Change the font color of only the matched text, not the entire cell
                lengthMatch = Len(searchText)
                With cell.Characters(startPos, lengthMatch).Font
                    .Bold = True ' Make the matched text bold
                    .Color = RGB(255, 0, 0) ' Change the matched text color to red
                End With
                ' If the cell contains a hyperlink, keep the hyperlink formatting
                If cell.Hyperlinks.Count > 0 Then
                    For Each hl In cell.Hyperlinks
                        With hl.range.Characters(startPos, lengthMatch).Font
                            .Color = RGB(5, 99, 193) ' Default hyperlink blue color
                            .Underline = xlUnderlineStyleSingle
                        End With
                    Next hl
                End If
            End If
        End If
        ' Update progress bar for cell processing
        currentItem = currentItem + 1
        UpdateStatusBar currentItem, totalCells + totalShapes
    Next cell

    ' ---- Search text within all shapes ----
    For Each shp In ws.Shapes
        ' Check if the shape has text
        If shp.TextFrame2.HasText Then
            ' Search for the text inside the shape
            foundPos = InStr(1, shp.TextFrame2.TextRange.Text, searchText, vbTextCompare)
            If foundPos > 0 Then
                ' Highlight only the matched part of the text within the shape
                With shp.TextFrame2.TextRange
                    totalChars = Len(.Text)
                    Do While foundPos > 0
                        .Characters(foundPos, Len(searchText)).Font.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Change to red text
                        ' Find the next occurrence
                        foundPos = InStr(foundPos + Len(searchText), shp.TextFrame2.TextRange.Text, searchText, vbTextCompare)
                    Loop
                End With
            End If
        End If
        ' Update progress bar for shape processing
        currentItem = currentItem + 1
        UpdateStatusBar currentItem, totalCells + totalShapes
    Next shp

    ' Reset the status bar when done
    Application.StatusBar = False

    ' Re-enable user interaction and screen updates
    Application.ScreenUpdating = True
    Application.Interactive = True
    Application.EnableEvents = True
End Sub

' Subroutine to update the status bar
Sub UpdateStatusBar(current As Long, total As Long)
    Application.StatusBar = "Processing " & current & " of " & total & " (" & Format(current / total, "0%") & ")"
    DoEvents ' Allows the UI to refresh
End Sub

' --------------------------------------------------------
' ---------List Page Search Box-------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub SearchAndHighlight()
    Dim searchText As String
    Dim cell As range
    Dim firstAddress As String
    Dim ws As Worksheet
    Dim shp As shape
    Dim foundPos As Long
    Dim totalChars As Long
    Dim cellText As String
    Dim startPos As Long
    Dim lengthMatch As Long
    Dim hl As Hyperlink

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' ---- Clear previous highlights in cells ----
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            ' Reset entire cell's font color to black and remove bold, but preserve hyperlinks
            With cell.Font
                .Color = RGB(0, 0, 0) ' Set text color back to black
                .Bold = False ' Remove bold formatting
            End With
            ' Reapply hyperlink formatting if the cell contains a hyperlink
            If cell.Hyperlinks.Count > 0 Then
                For Each hl In cell.Hyperlinks
                    With hl.range.Font
                        .Color = RGB(5, 99, 193) ' Default hyperlink blue color
                        .Underline = xlUnderlineStyleSingle
                    End With
                Next hl
            End If
        End If
    Next cell

    ' ---- Clear previous highlights in shapes ----
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            ' Reset the font color in the shape's text to black
            With shp.TextFrame2.TextRange
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' Set text color back to black
            End With
        End If
    Next shp

    ' Ask the user for the search text
    searchText = InputBox("Enter the text to search:")

    ' Exit if search text is empty
    If searchText = "" Then Exit Sub

    ' Turn off screen updating to make the macro faster
    Application.ScreenUpdating = False

    ' ---- Search text within all cells ----
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            cellText = cell.Value
            startPos = InStr(1, cellText, searchText, vbTextCompare) ' Find the start position of the match
            If startPos > 0 Then
                ' Change the font color of only the matched text, not the entire cell
                lengthMatch = Len(searchText)
                With cell.Characters(startPos, lengthMatch).Font
                    .Bold = True ' Make the matched text bold
                    .Color = RGB(255, 0, 0) ' Change the matched text color to red
                End With
                ' If the cell contains a hyperlink, keep the hyperlink formatting
                If cell.Hyperlinks.Count > 0 Then
                    For Each hl In cell.Hyperlinks
                        With hl.range.Characters(startPos, lengthMatch).Font
                            .Color = RGB(5, 99, 193) ' Default hyperlink blue color
                            .Underline = xlUnderlineStyleSingle
                        End With
                    Next hl
                End If
            End If
        End If
    Next cell

    ' ---- Search text within all shapes ----
    For Each shp In ws.Shapes
        ' Check if the shape has text
        If shp.TextFrame2.HasText Then
            ' Search for the text inside the shape
            foundPos = InStr(1, shp.TextFrame2.TextRange.Text, searchText, vbTextCompare)
            If foundPos > 0 Then
                ' Highlight only the matched part of the text within the shape
                With shp.TextFrame2.TextRange
                    totalChars = Len(.Text)
                    Do While foundPos > 0
                        .Characters(foundPos, Len(searchText)).Font.Fill.ForeColor.RGB = RGB(255, 255, 0) ' Highlight in yellow
                        ' Find the next occurrence
                        foundPos = InStr(foundPos + Len(searchText), shp.TextFrame2.TextRange.Text, searchText, vbTextCompare)
                    Loop
                End With
            End If
        End If
    Next shp

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

' --------------------------------------------------------
' ---------Blank Note Page Button-------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub CreateSheetAndLink()
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim cell As range
    Dim sheetName As String ' Used for SheetName
    Dim sheetNameLong As String ' Used for Full Sheet Name at top of new sheet
    Dim maxLength As Integer ' Used for max chars allowed for sheet name
    Dim originalCellAddress As String ' Used to store the original cell address
    Dim originalSheetName As String ' Used to store the original sheet name

    ' Set the maximum length for a sheet name
    maxLength = 31

    ' Set the template worksheet name
    Set wsTemplate = ThisWorkbook.Sheets("Template_Note")

    ' Set the cell that is selected
    Set cell = Selection
    
    If Not IsSingleCellSelected() Then
        Exit Sub
    End If

    ' Store the address of the original cell and sheet
    originalCellAddress = cell.Address
    originalSheetName = cell.Parent.Name

    ' Get the sheet name from the selected cell text
    sheetName = cell.Value
    sheetNameLong = cell.Value

    ' Truncate if the sheet name exceeds the maximum length
    If Len(sheetName) > maxLength Then
        sheetName = Left(sheetName, maxLength)
        MsgBox "The sheet name was too long and has been truncated to: " & sheetName, vbExclamation
    End If
    
    ' Truncate if the sheet name exceeds the maximum length
    If Len(sheetName) = 0 Then
        MsgBox "You need to type some text first" & sheetName, vbExclamation
        Exit Sub
    End If

    ' Replace invalid characters in the sheet name
    sheetName = ReplaceInvalidCharacters(sheetName)

    ' Check if a sheet with that name already exists
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsNew Is Nothing Then
        ' Copy the template to create the new sheet
        On Error Resume Next
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        On Error GoTo 0
        Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsNew.Name = sheetName
        wsNew.range("$I$1").Value = sheetNameLong
        
        wsNew.Activate
        
        ' Set DisplayGridlines to False for the active window showing the worksheet
        ActiveWindow.DisplayGridlines = False

        ' Create a hyperlink in the selected cell to the new sheet
        cell.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:="'" & wsNew.Name & "'!A1", TextToDisplay:=sheetNameLong
        
        ' Formats the New HyperLink
        FormatHyperlinkText cell

        ' Store the address of the cell in a specific cell on the new sheet
        wsNew.range("CE2").Value = "'" & cell.Parent.Name & "'!" & cell.Address
        
        ' Reset the color of the sheet tab to nothing
        wsNew.Tab.ColorIndex = xlColorIndexNone

        ' Return to the location of the new hyperlink
        On Error Resume Next
        ThisWorkbook.Sheets(originalSheetName).Select
        ThisWorkbook.Sheets(originalSheetName).range(originalCellAddress).Select
        On Error GoTo 0
        
        ' Debugging output
        Debug.Print "Sheet created: " & wsNew.Name
    Else
        MsgBox "A sheet with this name already exists.", vbExclamation
    End If
End Sub

' --------------------------------------------------------
' ---------Blank List Page Button-------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub CreateListSheetAndLink()
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim cell As range
    Dim sheetName As String ' Used for SheetName
    Dim sheetNameLong As String ' Used for Full Sheet Name at top of new sheet
    Dim maxLength As Integer ' Used for max chars allowed for sheet name
    Dim originalCellAddress As String ' Used to store the original cell address
    Dim originalSheetName As String ' Used to store the original sheet name

    ' Set the maximum length for a sheet name
    maxLength = 31

    ' Set the template worksheet name
    Set wsTemplate = ThisWorkbook.Sheets("Template_List")

    ' Set the cell that is selected
    Set cell = Selection
    
    If Not IsSingleCellSelected() Then
        Exit Sub
    End If

    ' Store the address of the original cell and sheet
    originalCellAddress = cell.Address
    originalSheetName = cell.Parent.Name

    ' Get the sheet name from the selected cell text
    sheetName = cell.Value
    sheetNameLong = cell.Value

    ' Truncate if the sheet name exceeds the maximum length
    If Len(sheetName) > maxLength Then
        sheetName = Left(sheetName, maxLength)
        MsgBox "The sheet name was too long and has been truncated to: " & sheetName, vbExclamation
    End If
    
    ' Truncate if the sheet name exceeds the maximum length
    If Len(sheetName) = 0 Then
        MsgBox "You need to type some text first" & sheetName, vbExclamation
        Exit Sub
    End If

    ' Replace invalid characters in the sheet name
    sheetName = ReplaceInvalidCharacters(sheetName)

    ' Check if a sheet with that name already exists
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If wsNew Is Nothing Then
        ' Copy the template to create the new sheet
        On Error Resume Next
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        On Error GoTo 0
        Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsNew.Name = sheetName
        wsNew.range("$I$1").Value = sheetNameLong
        wsNew.range("$J$3").Value = sheetNameLong
        
        wsNew.Activate
        
        ' Set DisplayGridlines to False for the active window showing the worksheet
        ActiveWindow.DisplayGridlines = False

        ' Create a hyperlink in the selected cell to the new sheet
        cell.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:="'" & wsNew.Name & "'!A1", TextToDisplay:=sheetNameLong
        
        ' Formats the New HyperLink
        FormatHyperlinkText cell

        ' Store the address of the cell in a specific cell on the new sheet
        wsNew.range("CE2").Value = "'" & cell.Parent.Name & "'!" & cell.Address
        
        ' Reset the color of the sheet tab to nothing
        wsNew.Tab.ColorIndex = xlColorIndexNone

        ' Return to the location of the new hyperlink
        On Error Resume Next
        ThisWorkbook.Sheets(originalSheetName).Select
        ThisWorkbook.Sheets(originalSheetName).range(originalCellAddress).Select
        On Error GoTo 0
        
        ' Debugging output
        Debug.Print "Sheet created: " & wsNew.Name
    Else
        MsgBox "A sheet with this name already exists.", vbExclamation
    End If
End Sub

' --------------------------------------------------------
' ---------Blank Note Page Button-------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Function ReplaceInvalidCharacters(sheetName As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("/", "\", ":", "*", "?", "[", "]")
    
    Dim i As Integer
    For i = LBound(invalidChars) To UBound(invalidChars)
        sheetName = Replace(sheetName, invalidChars(i), "_")
    Next i
    
    ReplaceInvalidCharacters = sheetName
End Function



' --------------------------------------------------------
' ---------Back Button------------------------------------
' --------Generated When sheet Created--------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub GoToBackCell()
    Dim wsName As String
    Dim cellRef As String
    Dim parts As Variant
    
    ' Split the address stored in CE2
    parts = Split(ActiveSheet.range("CE2").Value, "'!")
    
    If UBound(parts) = 1 Then
        wsName = Replace(parts(0), "'", "")
        cellRef = parts(1)
        
        ' Select the sheet and range
        On Error Resume Next
        ThisWorkbook.Sheets(wsName).Select
        ThisWorkbook.Sheets(wsName).range(cellRef).Select
        On Error GoTo 0
    Else
        MsgBox "Stored address is not valid.", vbExclamation
    End If
End Sub



' --------------------------------------------------------
' ---------HyperLink Format Button------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub FormatHyperlinkTextButton()
    Dim cell As range
    Set cell = Selection
    ' Format the hyperlink text using the function
    FormatHyperlinkText cell
End Sub

' --------------------------------------------------------
' ---------HyperLink Format Function----------------------
' -----------Used by the Blank Note Page Button-----------
' --------------------------------------------------------
' --------------------------------------------------------
Function FormatHyperlinkText(rng As range)
    With rng
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 30
    End With
    With rng.Font
        .Name = "Times New Roman"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
        .Underline = xlUnderlineStyleSingle
        .Bold = True
    End With
End Function


' --------------------------------------------------------
' ---------Text Format Button-----------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub FormatNormalText()
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 23
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorHyperlink
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub


' --------------------------------------------------------
' ---------Toggle Zoom Button-----------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub NavagationZoomToSection()

     ' Dim myRange As range

    ' Set myRange = range(Selection.Address)
    ' myRange = range(Selection.Address).Select
    
    ' activewindow.ScrollRow =
    If ActiveWindow.Zoom <= 55 Then
    ActiveWindow.Zoom = 85
    ' CenterOnCell range(Selection.Address)
    Exit Sub
    End If
    
    If ActiveWindow.Zoom < 85 Then
    ActiveWindow.Zoom = 85
    ' CenterOnCell range(Selection.Address)
    Exit Sub
    End If
    
    
    If ActiveWindow.Zoom >= 85 Then
    ActiveWindow.Zoom = 55
    ' CenterOnCell range(Selection.Address)
    Exit Sub
    End If
    ' ActiveWindow.VisibleRange.
    
End Sub

' --------------------------------------------------------
' ---------Search Button----------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
' --------------------------------------------------------
Sub FindInShape1()
    Dim rStart As range
    Dim shp As shape
    Dim sFind As String
    Dim sTemp As String
    Dim Response

    sFind = InputBox("Search for?")
    If Trim(sFind) = "" Then
        MsgBox "Nothing entered"
        Exit Sub
    End If
    Set rStart = ActiveCell
    For Each shp In ActiveSheet.Shapes
        sTemp = shp.TextFrame.Characters.Text
        If InStr(LCase(sTemp), LCase(sFind)) <> 0 Then
            shp.Select
            Response = MsgBox( _
              prompt:=shp.Name & vbCrLf & _
              sTemp & vbCrLf & vbCrLf & _
              "Do you want to continue?", _
              Buttons:=vbYesNo, Title:="Continue?")
            If Response <> vbYes Then
                Set rStart = Nothing
                Exit Sub
            End If
        End If
    Next
    MsgBox "No more found"
    rStart.Select
    Set rStart = Nothing
End Sub

' --------------------------------------------------------
' ---------Ensure Single Cell Selection Helper------------
' -----------Used by all the add shapes and objects-------
' --------------------------------------------------------
' --------------------------------------------------------
Function IsSingleCellSelected() As Boolean
    ' Check if the selection is a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a single cell.", vbExclamation, "Invalid Selection"
        IsSingleCellSelected = False
        Exit Function
    End If

    ' Check if more than one cell is selected
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select only one cell.", vbExclamation, "Invalid Selection"
        IsSingleCellSelected = False
        Exit Function
    End If

    ' If we pass both checks, the selection is valid
    IsSingleCellSelected = True
End Function


' --------------------------------------------------------
' ---------Selects A shape on the hseet by name------------
' -----------Make a copy of that shape and pastes it-------
' ---------Extracts the text from the cell-----------------
' ----------Puts text in shape ----------------------------
Function AddShapeWithCellText(shapeName As String) As Boolean
    On Error GoTo ErrorHandler

    ' Call the helper function to check if the selection is a valid single cell
    If Not IsSingleCellSelected() Then
        AddShapeWithCellText = False
        Exit Function
    End If

    Dim myRange As range
    Dim shapeToCopy As shape
    Dim pastedShape As shape
    Dim cellText As String

    ' Set myRange to the selected single cell
    Set myRange = Selection

    ' Get the text from the selected cell
    cellText = myRange.Value

    ' Reference the shape by its name
    On Error Resume Next
    
    
    Set shapeToCopy = ActiveSheet.Shapes(shapeName)
    
    ' Set shapeToCopy = ActiveSheet.Shapes.range(Array(shapeName)).Select
    
    If shapeToCopy Is Nothing Then
        MsgBox "Shape '" & shapeName & "' not found.", vbExclamation, "Error"
        AddShapeWithCellText = False
        Exit Function
    End If
    On Error GoTo 0

    ' Copy the referenced shape
    shapeToCopy.Copy

    ' Paste the copied shape into the selected cell
    myRange.Select
    ActiveSheet.Paste

    ' Get the pasted shape (the most recent shape added to the sheet)
    Set pastedShape = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)

    ' Preserve formatting and set new text
    pastedShape.TextFrame2.TextRange.Text = cellText
    
        ' Set text formatting (Arial, size 22)
    With pastedShape.TextFrame2.TextRange.Font
        .Name = "Arial"
        .Size = 22
    End With

    ' Center the text horizontally and vertically
    With pastedShape.TextFrame2
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

    ' Ensure the text is centered within the shape
    pastedShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' Clear the text from the selected cell
    myRange.Value = ""

    AddShapeWithCellText = True
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    AddShapeWithCellText = False
End Function


' --------------------------------------------------------
' ---------Short Reminders Button-------------------------
' ---------Basic Shape Insert-----------------------------
' --------Requires Shape Called: ShortReminders-----------
' --------------------------------------------------------
Sub AddShortRemindersShape()
    Dim success As Boolean
    success = AddShapeWithCellText("ShortReminders")
    
    If Not success Then
        MsgBox "Failed to add the shape.", vbExclamation, "Error"
    End If
End Sub

' --------------------------------------------------------
' ---------Important Yellow Button-------------------------
' ---------Basic Shape Insert-----------------------------
' --------Requires Shape Called: ImportantYellow-----------
' --------------------------------------------------------
Sub AddImportantYellowShape()
    Dim success As Boolean
    success = AddShapeWithCellText("ImportantYellow")
    
    If Not success Then
        MsgBox "Failed to add the shape.", vbExclamation, "Error"
    End If
End Sub

' --------------------------------------------------------
' ---------Simple Pointer Button-------------------------
' ---------Basic Shape Copy-----------------------------
' --------Requires Shape in sheet: SimpleWideGreen--------
' --------------------------------------------------------
Sub AddSimpleWideGreenShape()
    Dim success As Boolean
    success = AddShapeWithCellText("SimpleWideGreen")
    
    If Not success Then
        MsgBox "Failed to add the shape.", vbExclamation, "Error"
    End If
End Sub


' --------------------------------------------------------
' ---------Complex Pointer Button-------------------------
' ---------Basic Shape Copy-----------------------------
' --------Requires Shape in sheet: SimpleWideGreen--------
' --------------------------------------------------------
Sub AddComplexNarrowShape()
    Dim success As Boolean
    success = AddShapeWithCellText("ComplexNarrow")
    
    If Not success Then
        MsgBox "Failed to add the shape.", vbExclamation, "Error"
    End If
End Sub


' --------------------------------------------------------
' ---------Lengthy Notes Button-------------------------
' ---------Basic Shape Copy-----------------------------
' --------Requires Shape in sheet: SimpleWideGreen--------
' --------------------------------------------------------
Sub AddLengthyNotesShape()
    Dim success As Boolean
    success = AddShapeWithCellText("LengthyNotes")
    
    If Not success Then
        MsgBox "Failed to add the shape.", vbExclamation, "Error"
    End If
End Sub



' --------------------------------------------------------
' ---------Insert Document Button-------------------------
' ---------Insert an embedded word doc--------------------
' --------Requires ExtractFilename() --------
' --------------------------------------------------------
Sub AAADocumentInsert()
    Dim oleObjFull As OLEObject
    Dim oleObjIcon As OLEObject
    Dim filePath As String
    Dim fileName As String
    Dim selectedCell As range
    Dim scaleFactor As Double
    Dim moveUp As Double
    Dim originalWidth As Double
    Dim originalHeight As Double
    Dim shapeArray As Variant
    Dim groupedShape As ShapeRange
    

    ' Set the scaling factor (e.g., 2 means 200%)
    scaleFactor = 2
    moveUp = 85

    ' Store the currently selected cell to position the objects accordingly
    Set selectedCell = ActiveCell

    ' Prompt user to select the Word document
    filePath = Application.GetOpenFilename("Word Files (*.doc; *.docx), *.doc; *.docx", , "Select Document")
    ' Exit if user cancels
    If filePath = "False" Then Exit Sub
    fileName = ExtractFilename(filePath)

    ' Add the full display OLE object
    Set oleObjFull = ActiveSheet.OLEObjects.Add(fileName:=filePath, Link:=False, DisplayAsIcon:=False)

    ' Add the icon OLE object
    Set oleObjIcon = ActiveSheet.OLEObjects.Add(fileName:=filePath, Link:=False, DisplayAsIcon:=True, _
        IconFileName:="C:\WINDOWS\Installer\{90160000-000F-0000-1000-0000000FF1CE}\wordicon.exe", _
        IconIndex:=0, IconLabel:=fileName)

    ' Set the position of the embedded icon to match the selected cell
    oleObjIcon.Top = selectedCell.Top - moveUp
    oleObjIcon.Left = selectedCell.Left - moveUp

    ' Store original width and height
    originalWidth = oleObjIcon.Width
    originalHeight = oleObjIcon.Height

    ' Scale the icon object by the scaling factor
    oleObjIcon.Width = originalWidth * scaleFactor
    oleObjIcon.Height = originalHeight * scaleFactor

    ' Move the icon object to the back
    oleObjIcon.ShapeRange.ZOrder msoSendToBack
    
    ' Group Both objects
    ActiveSheet.Shapes.range(Array(oleObjFull.Name, oleObjIcon.Name)).Select
    Selection.ShapeRange.Group.Select
    
End Sub

' --------------------------------------------------------
' ---------Helper function for AAADocumentInsert----------
' ---------Get the FileName------------------------------
' ----------------------------------------------- --------
' --------------------------------------------------------
Function ExtractFilename(fullPath As String) As String
    ' Function to extract the filename from a full path
    ExtractFilename = Right(fullPath, Len(fullPath) - InStrRev(fullPath, "\"))
End Function


' --------------------------------------------------------
' ---------Create a new folder named the sheet title------
' ---------button click------------------------------
' ---------then insert a shape that opens the folder----------
' ---------then insert a shpae that contains the foder contents---
Sub ListPage_CreateFolderAndHyperlink()
    Dim folderName As String
    Dim folderPath As String
    Dim ws As Worksheet
    Dim shape As shape
    Dim currentWorkbookPath As String
    Dim cell As range
    Dim cellcontents As range
    Dim updatebuttonlocation As range
    
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    
    'Set New shape location
    Set cell = ws.range("T1")
    Set cellcontents = ws.range("T2")
    Set updatebuttonlocation = ws.range("Z1")

    ' Get the value from cell E1 as the folder name
    folderName = ws.range("I1").Value
    
    ' Get the current workbook path (without workbook name)
    currentWorkbookPath = ThisWorkbook.Path
    
    ' Construct the full folder path
    folderPath = currentWorkbookPath & "\" & folderName
    
    ' Check if the folder already exists, if not, create it
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        MsgBox "Folder Created: " & folderPath
    Else
        MsgBox "Folder already exists: " & folderPath
    End If
    
    ' Insert a shape on the worksheet
    ' AddShape(dist from left,dist from top, width, height)
    Set shape = ws.Shapes.AddShape(msoShapeRectangle, 80, 80, 250, 60) ' Adjust position and size as necessary
    'Place shape over the button
    shape.Left = cell.Left + 5
    shape.Top = cell.Top + 5
    ' Set the text for the shape
    shape.TextFrame2.TextRange.Text = "Open: Folder" ' Text in the Folder Shape
    shape.TextFrame2.TextRange.Font.Size = 40 ' Adjust the size as needed
    shape.TextFrame2.VerticalAnchor = msoAnchorMiddle ' ' Center the text vertically
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter ' Center the text horizontally
    ' Create the hyperlink to the folder
    ws.Hyperlinks.Add Anchor:=shape, Address:=folderPath, TextToDisplay:="Open Folder"
    ' Inform user the operation is complete
    MsgBox "Hyperlink created to the folder."
    
    
    ' Insert the dynamic content shape (where the folder contents will be shown)
    Set contentShape = ws.Shapes.AddShape(msoShapeRectangle, 300, 50, 600, 900)
    contentShape.Name = "contentShape" ' Name the shape so it can be referenced later
    contentShape.TextFrame2.TextRange.Text = "Folder Contents Here"
    contentShape.Left = cellcontents.Left
    contentShape.Top = cellcontents.Top
    contentShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    contentShape.TextFrame2.VerticalAnchor = msoAnchorTop
    ' End Maybe
    
    ' Insert a button to the left of contentShape to update folder contents
    Set updateButton = ws.Shapes.AddShape(msoShapeRectangle, 80, 80, 250, 60) ' Position to the left
    updateButton.TextFrame2.TextRange.Text = "Update Files"
    updateButton.TextFrame2.TextRange.Font.Size = 20
    updateButton.TextFrame2.VerticalAnchor = msoAnchorMiddle
    updateButton.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    updateButton.Left = updatebuttonlocation.Left + 5
    updateButton.Top = updatebuttonlocation.Top + 5

    ' Assign the macro to the button
    updateButton.OnAction = "'" & ThisWorkbook.Name & "'!UpdateContentShapeWithFiles"
    
    MsgBox "Button to update folder contents created."
    
    
End Sub

' --------------------------------------------------------
' ---------Create a new folder named the sheet title------
' ---------button click------------------------------
' ---------then insert a shape that opens the folder----------
' ---------then insert a shpae that contains the foder contents---
Sub NotePage_CreateFolderAndHyperlink()
        Dim folderName As String
    Dim folderPath As String
    Dim ws As Worksheet
    Dim shape As shape
    Dim currentWorkbookPath As String
    Dim cell As range
    Dim cellcontents As range
    Dim updatebuttonlocation As range
    
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    
    'Set New shape location
    Set cell = ws.range("AY1")
    Set cellcontents = ws.range("AY2")
    Set updatebuttonlocation = ws.range("BE1")

    ' Get the value from cell E1 as the folder name
    folderName = ws.range("I1").Value
    
    ' Get the current workbook path (without workbook name)
    currentWorkbookPath = ThisWorkbook.Path
    
    ' Construct the full folder path
    folderPath = currentWorkbookPath & "\" & folderName
    
    ' Check if the folder already exists, if not, create it
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        MsgBox "Folder Created: " & folderPath
    Else
        MsgBox "Folder already exists: " & folderPath
    End If
    
    ' Insert a shape on the worksheet
    ' AddShape(dist from left,dist from top, width, height)
    Set shape = ws.Shapes.AddShape(msoShapeRectangle, 80, 80, 250, 60) ' Adjust position and size as necessary
    'Place shape over the button
    shape.Left = cell.Left + 5
    shape.Top = cell.Top + 5
    ' Set the text for the shape
    shape.TextFrame2.TextRange.Text = "Open: Folder" ' Text in the Folder Shape
    shape.TextFrame2.TextRange.Font.Size = 40 ' Adjust the size as needed
    shape.TextFrame2.VerticalAnchor = msoAnchorMiddle ' ' Center the text vertically
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter ' Center the text horizontally
    ' Create the hyperlink to the folder
    ws.Hyperlinks.Add Anchor:=shape, Address:=folderPath, TextToDisplay:="Open Folder"
    ' Inform user the operation is complete
    MsgBox "Hyperlink created to the folder."
    
    
    ' Insert the dynamic content shape (where the folder contents will be shown)
    Set contentShape = ws.Shapes.AddShape(msoShapeRectangle, 300, 50, 600, 900)
    contentShape.Name = "contentShape" ' Name the shape so it can be referenced later
    contentShape.TextFrame2.TextRange.Text = "Folder Contents Here"
    contentShape.Left = cellcontents.Left
    contentShape.Top = cellcontents.Top
    contentShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    contentShape.TextFrame2.VerticalAnchor = msoAnchorTop
    ' End Maybe
    
    ' Insert a button to the left of contentShape to update folder contents
    Set updateButton = ws.Shapes.AddShape(msoShapeRectangle, 80, 80, 250, 60) ' Position to the left
    updateButton.TextFrame2.TextRange.Text = "Update Files"
    updateButton.TextFrame2.TextRange.Font.Size = 20
    updateButton.TextFrame2.VerticalAnchor = msoAnchorMiddle
    updateButton.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    updateButton.Left = updatebuttonlocation.Left + 5
    updateButton.Top = updatebuttonlocation.Top + 5

    ' Assign the macro to the button
    updateButton.OnAction = "'" & ThisWorkbook.Name & "'!UpdateContentShapeWithFiles"
    
    MsgBox "Button to update folder contents created."
    
    
End Sub

' --------------------------------------------------------
' ---------Helper function for CreateFolderAndHyperlink---
' ---------Insert a shape that acts as update button-------
' --------To update the shape that has the file contents----
' --------------------------------------------------------
Sub UpdateContentShapeWithFiles()
    Dim ws As Worksheet
    Dim folderName As String
    Dim folderPath As String
    Dim fileName As String
    Dim fileList As String
    Dim folderList As String
    Dim fileCount As Long
    Dim folderCount As Long
    Dim contentShape As shape
    Dim currentWorkbookPath As String
    Dim totalFilesText As String
    Dim totalFilesLength As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(ActiveSheet.Name)
    
    ' Get the folder name from cell E1
    folderName = ws.range("I1").Value
    
    ' Get the current workbook path (without workbook name)
    currentWorkbookPath = ThisWorkbook.Path
    
    ' Construct the full folder path
    folderPath = currentWorkbookPath & "\" & folderName
    
    ' Reference the contentShape (the shape that will display folder contents)
    Set contentShape = ws.Shapes("contentShape") ' Now, this will correctly reference the shape by its name
    
    ' Initialize the file list and file counter
    fileList = ""
    folderList = ""
    fileCount = 0
    folderCount = 0
    
    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder does not exist: " & folderPath
        Exit Sub
    End If
    
    ' Loop through each file in the folder
    fileName = Dir(folderPath & "\*.*", vbDirectory)
    Do While fileName <> ""
        If fileName <> "." And fileName <> ".." Then ' Skip the current (.) and parent (..) directories
            If (GetAttr(folderPath & "\" & fileName) And vbDirectory) = vbDirectory Then
                ' It's a folder
                folderList = folderList & fileName & vbCrLf ' Add folder name to folder list
                folderCount = folderCount + 1 ' Increment folder count
            Else
                ' It's a file
                fileList = fileList & fileName & vbCrLf ' Add file name to file list
                fileCount = fileCount + 1 ' Increment file count
            End If
        End If
        fileName = Dir ' Get the next file or folder
    Loop
    
    ' Construct the "Total Files" and "Total Folders" text and get its length
    totalFilesText = "Total Files: " & fileCount & vbCrLf & "Total Folders: " & folderCount
    totalFilesLength = Len(totalFilesText)

    ' Update the content shape with the list of files and folders and their counts
    contentShape.TextFrame2.TextRange.Text = totalFilesText & vbCrLf & "Folders:" & vbCrLf & folderList & vbCrLf & "Files:" & vbCrLf & fileList
    
    contentShape.TextFrame2.TextRange.Font.Size = 20 '
    contentShape.TextFrame2.TextRange.Font.Bold = msoFalse '
    
    contentShape.TextFrame2.TextRange.Characters(1, totalFilesLength).Font.Size = 40 ' Smaller font for file and folder list
    contentShape.TextFrame2.TextRange.Characters(1, totalFilesLength).Font.Bold = msoTrue ' Smaller font for file and folder list

    ' Set text alignment and vertical anchor
    contentShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    contentShape.TextFrame2.VerticalAnchor = msoAnchorTop
    
    ' Set line spacing
    contentShape.TextFrame2.TextRange.ParagraphFormat.LineRuleWithin = msoLineSingleSpacing
    contentShape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 35 ' Adjust line spacing if needed

    MsgBox "Folder contents updated in shape!"
End Sub


' --------------------------------------------------------
' ---------Helper function for CreateFolderAndHyperlink---
' -------------------------------------------------------
' --------inserts folder contents into the new shape----
' --------------------------------------------------------
Sub ListFolderContents()
    Dim folderPath As String
    Dim fileName As String
    Dim rowNum As Long
    Dim ws As Worksheet

    ' Set the worksheet where you want the list to appear
    Set ws = ThisWorkbook.Sheets(1) ' Change to your specific sheet if needed

    ' Set the folder path
    folderPath = "C:\YourFolderPath" ' Replace with your folder path

    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder does not exist: " & folderPath
        Exit Sub
    End If

    ' Clear previous list of files (optional)
    ws.Cells.Clear

    ' Add header for the list
    ws.Cells(1, 1).Value = "Files in folder: " & folderPath

    ' Initialize the row counter
    rowNum = 2

    ' Loop through each file in the folder
    fileName = Dir(folderPath & "\*.*")
    Do While fileName <> ""
        ws.Cells(rowNum, 1).Value = fileName
        rowNum = rowNum + 1
        fileName = Dir
    Loop

    MsgBox "File list updated!"
End Sub



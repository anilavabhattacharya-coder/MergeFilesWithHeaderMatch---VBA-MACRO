'----------------------------------------------------------------------------------
' MACRO : MergeToBaseFile
' PURPOSE: Asks the user to select a BASE Excel file, then repeatedly asks for
'          additional files. Each additional file's data is matched to the BASE
'          file's headers and appended after the last used row. Continues until
'          the user cancels the file picker.
'----------------------------------------------------------------------------------

Option Explicit

Sub MergeToBaseFile()

    '--------------------------------------------------------------------------
    ' 0. DECLARATIONS
    '--------------------------------------------------------------------------
    Dim wbBase          As Workbook
    Dim wsBase          As Worksheet
    Dim wbNew           As Workbook
    Dim wsNew           As Worksheet

    Dim baseHeaders()   As String
    Dim newHeaders()    As String
    Dim baseColCount    As Long
    Dim newColCount     As Long

    Dim baseLastRow     As Long
    Dim newLastRow      As Long
    Dim newLastCol      As Long

    Dim baseColIdx      As Long
    Dim newColIdx       As Long

    Dim fileCount       As Long
    Dim rowCount        As Long
    Dim totalAdded      As Long

    Dim i               As Long
    Dim j               As Long
    Dim matchFound      As Boolean

    Dim fd              As FileDialog
    Dim filePath        As String
    Dim logMsg          As String

    fileCount  = 0
    totalAdded = 0

    '--------------------------------------------------------------------------
    ' 1. SELECT BASE FILE
    '--------------------------------------------------------------------------
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Step 1 – Select the BASE Excel file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
        .AllowMultiSelect = False
    End With

    If fd.Show <> -1 Then
        MsgBox "No base file selected. Macro cancelled.", vbExclamation, "Cancelled"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    Set wbBase = Workbooks.Open(fd.SelectedItems(1))
    Set wsBase = wbBase.Sheets(1)

    ' --- Read base headers from Row 1 ---
    baseColCount = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column

    If baseColCount = 0 Or wsBase.Cells(1, 1).Value = "" Then
        MsgBox "The base file appears to have no headers in Row 1. Macro cancelled.", _
               vbExclamation, "No Headers Found"
        wbBase.Close False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        Exit Sub
    End If

    ReDim baseHeaders(1 To baseColCount)
    For i = 1 To baseColCount
        baseHeaders(i) = Trim(LCase(CStr(wsBase.Cells(1, i).Value)))
    Next i

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    MsgBox "Base file loaded successfully!" & vbCrLf & vbCrLf & _
           "Base file : " & wbBase.Name & vbCrLf & _
           "Headers found : " & baseColCount & vbCrLf & vbCrLf & _
           "Now select files to merge. Click Cancel when done.", _
           vbInformation, "Base File Ready"

    '--------------------------------------------------------------------------
    ' 2. LOOP — KEEP ASKING FOR FILES UNTIL USER CANCELS
    '--------------------------------------------------------------------------
    Do

        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "Select file to merge (Cancel to finish)"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
            .AllowMultiSelect = False
        End With

        ' User cancelled — stop loop
        If fd.Show <> -1 Then Exit Do

        filePath = fd.SelectedItems(1)

        ' Prevent user from selecting the base file again
        If LCase(filePath) = LCase(wbBase.FullName) Then
            MsgBox "You selected the base file itself. Please choose a different file.", _
                   vbExclamation, "Same File"
            GoTo NextIteration
        End If

        Application.ScreenUpdating = False
        Application.DisplayAlerts  = False

        ' --- Open the new file ---
        On Error Resume Next
        Set wbNew = Workbooks.Open(filePath)
        On Error GoTo 0

        If wbNew Is Nothing Then
            Application.ScreenUpdating = True
            Application.DisplayAlerts  = True
            MsgBox "Could not open the file. Skipping." & vbCrLf & filePath, _
                   vbExclamation, "Open Error"
            GoTo NextIteration
        End If

        Set wsNew = wbNew.Sheets(1)

        newLastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
        newLastCol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column

        If newLastRow <= 1 Or wsNew.Cells(1, 1).Value = "" Then
            wbNew.Close False
            Application.ScreenUpdating = True
            Application.DisplayAlerts  = True
            MsgBox "The file '" & wbNew.Name & "' has no data or headers. Skipping.", _
                   vbExclamation, "No Data"
            GoTo NextIteration
        End If

        ' --- Read new file headers ---
        ReDim newHeaders(1 To newLastCol)
        For i = 1 To newLastCol
            newHeaders(i) = Trim(LCase(CStr(wsNew.Cells(1, i).Value)))
        Next i

        ' --- Find base last used row (where we will start pasting) ---
        baseLastRow = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

        ' --- Build column mapping: for each BASE column, find matching NEW column ---
        '     mapping(baseColIdx) = newColIdx   (0 = no match)
        Dim colMap() As Long
        ReDim colMap(1 To baseColCount)

        Dim matchedCols As Long
        Dim unmatchedList As String
        matchedCols   = 0
        unmatchedList = ""

        For baseColIdx = 1 To baseColCount
            colMap(baseColIdx) = 0          ' default = no match
            For newColIdx = 1 To newLastCol
                If baseHeaders(baseColIdx) = newHeaders(newColIdx) Then
                    colMap(baseColIdx) = newColIdx
                    matchedCols = matchedCols + 1
                    Exit For
                End If
            Next newColIdx
            If colMap(baseColIdx) = 0 Then
                unmatchedList = unmatchedList & "  • " & wsBase.Cells(1, baseColIdx).Value & vbCrLf
            End If
        Next baseColIdx

        ' --- Warn if some base columns had no match ---
        If unmatchedList <> "" Then
            Dim proceed As Integer
            proceed = MsgBox("File: " & wbNew.Name & vbCrLf & vbCrLf & _
                             "The following BASE columns have NO matching header in this file " & _
                             "(they will be left blank):" & vbCrLf & unmatchedList & vbCrLf & _
                             "Continue merging this file?", _
                             vbYesNo + vbQuestion, "Unmatched Columns")
            If proceed = vbNo Then
                wbNew.Close False
                Application.ScreenUpdating = True
                Application.DisplayAlerts  = True
                GoTo NextIteration
            End If
        End If

        ' --- Copy data row by row, column by column using mapping ---
        rowCount = 0
        For i = 2 To newLastRow                         ' skip header row
            Dim destRow As Long
            destRow = baseLastRow + rowCount + 1

            For baseColIdx = 1 To baseColCount
                If colMap(baseColIdx) <> 0 Then
                    ' Matched column — copy value & number format
                    With wsBase.Cells(destRow, baseColIdx)
                        .Value        = wsNew.Cells(i, colMap(baseColIdx)).Value
                        .NumberFormat = wsNew.Cells(i, colMap(baseColIdx)).NumberFormat
                    End With
                End If
                ' Unmatched columns stay empty (blank)
            Next baseColIdx

            rowCount = rowCount + 1
        Next i

        fileCount  = fileCount  + 1
        totalAdded = totalAdded + rowCount

        wbNew.Close False

        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True

        logMsg = "Merged: " & Mid(filePath, InStrRev(filePath, "\") + 1) & vbCrLf & _
                 "  Rows added   : " & rowCount & vbCrLf & _
                 "  Columns matched: " & matchedCols & " of " & baseColCount & vbCrLf & vbCrLf & _
                 "Select the next file, or click Cancel to finish."

        MsgBox logMsg, vbInformation, "File " & fileCount & " Merged"

NextIteration:
    Loop

    '--------------------------------------------------------------------------
    ' 3. SAVE BASE FILE & WRAP UP
    '--------------------------------------------------------------------------
    If fileCount > 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts  = False

        ' Auto-fit all columns in base sheet for readability
        wsBase.Cells.EntireColumn.AutoFit

        wbBase.Save

        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True

        MsgBox "All done!" & vbCrLf & vbCrLf & _
               "Files merged    : " & fileCount & vbCrLf & _
               "Total rows added: " & totalAdded & vbCrLf & vbCrLf & _
               "Base file saved : " & wbBase.FullName, _
               vbInformation, "Merge Complete"
    Else
        MsgBox "No files were merged. Base file was not modified.", _
               vbInformation, "Nothing Merged"
    End If

    ' Leave base workbook open for the user to review
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

End Sub

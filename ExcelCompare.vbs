Set fso = CreateObject("Scripting.FileSystemObject")
Dim sFilePath1, sFilePath2
If WScript.Arguments.Count = 2 then
	sFilePath1 = WScript.Arguments(0)
	sFilePath2 = WScript.Arguments(1)
Else
    MsgBox("Please drag and drop two excel files.")	    
    Wscript.Quit
End If

If fso.FileExists(sFilePath1) = False  Then
	MsgBox "File 1 is missing: " & sFilePath1
    Wscript.Quit
End If

If fso.FileExists(sFilePath2) = False Then
	MsgBox "File 2 is missing: " & sFilePath2
    Wscript.Quit
End If

Dim sMissingSheets: sMissingSheets = ""
Dim iDiffCount: iDiffCount = 0
Dim oExcel: Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
oExcel.DisplayAlerts = false
Set oWorkBook1 = oExcel.Workbooks.Open(sFilePath1)
Set oWorkBook2 = oExcel.Workbooks.Open(sFilePath2)

For Each oSheet in oWorkBook1.Worksheets
    If SheetExists(oWorkBook2, oSheet.Name) = False Then
        if sMissingSheets <> "" Then sMissingSheets = sMissingSheets & ","
        sMissingSheets = sMissingSheets & oSheet.Name
    Else
        oSheet.Activate
        Set oSheet2 = oWorkBook2.Worksheets(oSheet.Name)

        iColCount = GetLastCol(oSheet)
        iRowsCount = GetLastRowWithData(oSheet)

        For iRow = 1 to iRowsCount
            For iCol = 1 to iColCount
                If oSheet.Cells(iRow, iCol).Value <> oSheet2.Cells(iRow, iCol).Value Then
                    oSheet.Cells(iRow, iCol).Interior.Color = 65535
                    oSheet2.Cells(iRow, iCol).Interior.Color = 65535
                    iDiffCount = iDiffCount + 1
                End If
            Next
        Next
    
    End If
Next

For Each oSheet in oWorkBook2.Worksheets
    If SheetExists(oWorkBook1, oSheet.Name) = False Then
        if sMissingSheets <> "" Then sMissingSheets = sMissingSheets & ","
        sMissingSheets = sMissingSheets & oSheet.Name
    End If
Next

If iDiffCount = 0 Then
    MsgBox "Files match"
Else
    MsgBox "Found " & iDiffCount & " differences"
End If

If sMissingSheets <> "" Then
    MsgBox "Missing Worksheets: " & sMissingSheets
End If

Function GetLastRowWithData(oSheet)
    Dim iMaxRow: iMaxRow = oSheet.UsedRange.Rows.Count
    If iMaxRow > 500 Then
        iMaxRow = oSheet.Cells.Find("*", oSheet.Cells(1, 1),  -4163, , 1, 2).Row
    End If

    Dim iRow, iCol
    For iRow = iMaxRow to 1 Step -1
         For iCol = 1 to oSheet.UsedRange.Columns.Count
            If Trim(oSheet.Cells(iRow, iCol).Value) <> "" Then
                GetLastRowWithData = iRow
                Exit Function
            End If
         Next
    Next
    GetLastRowWithData = 1
End Function

Function GetLastCol(st)
    GetLastCol = st.Cells.Find("*", st.Cells(1, 1), , 2, 2, 2, False).Column
End Function

Function SheetExists(oWorkBook, sName)
    on error resume next
    Dim oSheet: Set oSheet = oWorkBook.Worksheets(sName) 
    If Err.number = 0 Then
        SheetExists = True
    Else
        SheetExists = False
        Err.Clear
    End If
End Function


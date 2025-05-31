Const sFirstColData = "Calendar"
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
Dim iDiffCell: iDiffCell = 0
Dim iDiffRow: iDiffRow = 0
Dim iDiffCol: iDiffCol = 0
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
        Set rs = GetExcelRecordset(oSheet)
        Set rs2 = GetExcelRecordset(oSheet2)
        CompareCells oSheet, rs, oSheet2, rs2
        CompareCells oSheet2, rs2, oSheet, rs
    End If
Next

For Each oSheet in oWorkBook2.Worksheets
    If SheetExists(oWorkBook1, oSheet.Name) = False Then
        if sMissingSheets <> "" Then sMissingSheets = sMissingSheets & ","
        sMissingSheets = sMissingSheets & oSheet.Name
    End If
Next

Dim sDiff: sDiff = ""

if iDiffCell <> 0 Then
    sDiff = sDiff & iDiffCell & " cell differences."
End If

if iDiffRow <> 0 Then
    if sDiff <> "" Then sDiff = sDiff & " "
    sDiff = sDiff & iDiffRow & " row differences."
End If

if iDiffCol <> 0 Then
    if iDiffCol <> "" Then sDiff = sDiff & " "
    sDiff = sDiff & iDiffCol & " column differences."
End If

If sMissingSheets <> "" Then
    if sDiff <> "" Then sDiff = sDiff & " "
    sDiff = sDiff & "Missing Worksheets: " & sMissingSheets & "."
End If

If sDiff = "" Then
    MsgBox "Files match"
Else
    MsgBox "Found " & sDiff
End If

'==============================================
Sub CompareCells(oSheet, rs, oSheet2, rs2)

    ResetRs rs
    ResetRs rs2

    Dim oColDiff: Set oColDiff = CreateObject("Scripting.Dictionary")
    Dim col: Set col = GetColDiff(oSheet,oSheet2)
    Dim iRow, iRow2

    While rs.EOF = False
        iRow = rs("RowNumber").Value
        sFirstCol = rs("c1").value & ""
        If sFirstCol <> "" Then
            rs2.Filter = "c1 = '" & sFirstCol & "'"
            If rs2.RecordCount = 0 Then
                oSheet.Rows(iRow & ":" & iRow).Interior.Color = RGB(219, 255, 0)
                iDiffRow = iDiffRow + 1

            ElseIf rs2.RecordCount = 1 Then
                iRow2 = rs2("RowNumber").Value
                    
                For iCol = 1 to rs.Fields.Count - 1
                    iCol2 = iCol
                    If col.Exists(iCol) Then
                        iCol2 = col(iCol)
                    End If

                    If iCol2 = -1 Then
                        'Col not found
                        If oColDiff.Exists(iCol) = False Then
                            oSheet.Columns(iCol).Interior.Color = RGB(219, 255, 51)
                            oColDiff(iCol) = True
                        End If

                    ElseIf iCol >= rs.Fields.Count Or iCol2 >= rs2.Fields.Count Then
                        'Out of range

                    ElseIf rs(iCol).Value & "" <> rs2(iCol2).Value & "" Then
                        oSheet.Cells(iRow, iCol ).Interior.Color = 65535
                        iDiffCell = iDiffCell + 1
                    End If
                Next
            End If
        End If
        rs.MoveNext
    Wend

    If oColDiff.Count > 0  Then
        iDiffCol = iDiffCol + oColDiff.Count 
    End If
End Sub

Sub ResetRs(rs)
    rs.Filter = ""
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
End Sub

Function GetColDiff(oSheet,oSheet2)
    Dim oRet: Set oRet = CreateObject("Scripting.Dictionary")
    Dim oCols: Set oCols = GetExcelColumns(oSheet)
    Dim oCols2: Set oCols2 = GetExcelColumns(oSheet2)
    Dim iCol: iCol = 0

    For Each sKey In oCols.Keys
        iCol = oCols(sKey)
        If oCols2.Exists(sKey) Then
            If iCol <> oCols2(sKey) Then
                oRet(iCol) = oCols2(sKey) 'Col 1 => 2 (column was moved for 1 to 2)
            End If
        Else
            oRet(iCol) = -1 'Col not found
        End If
    Next

    Set GetColDiff = oRet
End Function

Function GetExcelColumns(oSheet)
    Dim oCols: Set oCols = CreateObject("Scripting.Dictionary")
    Dim iHeaderRow: iHeaderRow = 1

    If sFirstColData <> "" Then
        For i = 1 to 100
            If oSheet.Cells(i, 1).Value = sFirstColData Then
                iHeaderRow = i -1
                Exit For
            End If
        Next
    End If

    Dim iColCount: iColCount = GetLastCol(oSheet)

    For iCol = 1 to iColCount
        sVal = oSheet.Cells(iHeaderRow, iCol).Value
        If sVal <> "" Then
            oCols(sVal) = iCol
        End If
    Next
    Set GetExcelColumns = oCols
End Function

Function GetExcelRecordset(oSheet)
    Dim iColCount: iColCount = GetLastCol(oSheet)
    Dim iRowsCount: iRowsCount = GetLastRowWithData(oSheet)

    Dim rs: Set rs= CreateObject("ADODB.recordset")
    rs.Fields.Append "RowNumber", 3 'adInteger

    For iCol = 1 to iColCount
        rs.Fields.Append "c" & iCol, 203, -1 'adVarChar
    Next

    rs.Open

    For iRow = 1 to iRowsCount
        rs.AddNew  
        rs("RowNumber") = iRow

        For iCol = 1 to iColCount
            rs("c" & iCol).Value = oSheet.Cells(iRow, iCol).Value & ""          
        Next
    Next

    rs.MoveFirst
    Set GetExcelRecordset = rs
End Function

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


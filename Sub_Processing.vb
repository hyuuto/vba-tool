'*****************************************************************************************************
'GLOBAL CONSTANT AND VARIABLES
'*****************************************************************************************************
Const String_InputSheet_Name As String = "InputData"
Const String_PivotSheet_Name As String = "PivotData"
Const String_ReportSheet_Name As String = "DataReport"

Dim ImportData As Variant
Dim ImportData_LastRow As Long

Dim ImpotData_Header_Dictionary As Object
Dim ImportDataColumnName(4) As Variant

'MainForm
Dim TempFolderPath As String


'*****************************************************************************************************
'FUNCTION NAME: GetDataFromInputFile
'*****************************************************************************************************
Function GetDataFromInputFile(String_FilePath As String)
    On Error GoTo continue1
    Dim Long_LastRow As Long
    Dim Long_LastColumn As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ws_collect_input As Worksheet
    Dim ws_refer As Worksheet
    Dim max_collect_input As Long
    Dim max_refer As Long
    Dim EachwsLastRow As Long, EachwsLastColumn As Long

    Application.DisplayAlerts = False
    For Each op_ws In ThisWorkbook.Sheets
        If op_ws.Name <> "MainMenu" Then
            op_ws.Delete
        End If
    Next op_ws
    Application.DisplayAlerts = True

    If InStr(LCase(String_FilePath), ".xls") <> 0 Then
        Set wb = Workbooks.Open(String_FilePath)
        For Each ws In wb.Sheets
            ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
  Next ws
        wb.Close False
  End If

    ThisWorkbook.Sheets.Add(After:=Sheets(1)).Name = String_InputSheet_Name
    For Each ws_refer In ThisWorkbook.Sheets
        max_collect_input = ThisWorkbook.Sheets(String_InputSheet_Name).Cells(ThisWorkbook.Sheets(String_InputSheet_Name).Rows.Count, "A").End(xlUp).Row
        If ws_refer.Name <> "MainMenu" And ws_refer.Name <> "InputData" And ws_refer.Name <> "PivotData" And ws_refer.Name <> "DataReport" Then
            Application.DisplayAlerts = False
            max_refer = ws_refer.Cells(ws_refer.Rows.Count, "A").End(xlUp).Row

            ThisWorkbook.Sheets(String_InputSheet_Name).Cells(1, 1) = ws_refer.Cells(1, 1)
            ThisWorkbook.Sheets(String_InputSheet_Name).Cells(1, 2) = ws_refer.Cells(1, 2)
            ThisWorkbook.Sheets(String_InputSheet_Name).Cells(1, 3) = ws_refer.Cells(1, 5)
            ThisWorkbook.Sheets(String_InputSheet_Name).Cells(1, 4) = ws_refer.Cells(1, 7)

            For i = 2 To max_refer
                If ws_refer.Cells(i, 1) = CStr(MainForm.YearMonth) Or CStr(MainForm.YearMonth) = "" Then
                    If ws_refer.Cells(i, 5) <> "PM/PL/PF" And ws_refer.Cells(i, 5) <> "Partner" And ws_refer.Cells(i, 5) <> "" Then
                        ThisWorkbook.Sheets(String_InputSheet_Name).Cells(i, 1) = ws_refer.Cells(i, 1)
                        ThisWorkbook.Sheets(String_InputSheet_Name).Cells(i, 2) = ws_refer.Cells(i, 2)
                        ThisWorkbook.Sheets(String_InputSheet_Name).Cells(i, 3) = ws_refer.Cells(i, 5)
                        ThisWorkbook.Sheets(String_InputSheet_Name).Cells(i, 4) = ws_refer.Cells(i, 7)
                    End If
                End If
            Next i

            Application.DisplayAlerts = True
        End If
    Next ws_refer
    
    Set ws = ThisWorkbook.Sheets(String_InputSheet_Name)
    Call DeleteEmptyRows(ws)

    ws.Range("A1").Select

    Long_LastRow = ws.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    Long_LastColumn = ws.Cells.Find("*", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    FileArray_Data = ws.Range(Cells(1, 1), Cells(Long_LastRow, Long_LastColumn)).Value

    If String_InputSheet_Name = String_InputSheet_Name Then
        ImportData = FileArray_Data
        ImportData_LastRow = Long_LastRow
    End If

continue1:
    Call DeleteQueries
    Call DeleteConnections
    If InStr(LCase(String_FilePath), ".xls") <> 0 Then
    End If
End Function
'*****************************************************************************************************
'FUNCTION NAME: PivotData
'*****************************************************************************************************
Function PivotData()
    Dim ws As Worksheet
    Dim maxRole As Integer
    Dim uBImportData As Integer
    maxRole = 12

    uBImportData = UBound(ImportData)

    'Add Pivot Sheet
    ThisWorkbook.Sheets.Add(After:=Sheets(2)).Name = String_PivotSheet_Name
    Set ws = ThisWorkbook.Sheets(String_PivotSheet_Name)
    ws.Cells(1, 1) = "Month"
    ws.Cells(1, 2) = "Role"
    ws.Cells(1, 3) = "Number"

    ws.Cells(1, 4) = "Utilitization == 0%"
    ws.Cells(1, 5) = "0% < Utilitization < 50%"

    ws.Cells(1, 6) = "Utilitization == 0% (Quantity)"
    ws.Cells(1, 7) = "0% < Utilitization < 50% (Quantity)"
    ws.Columns("A").NumberFormat = "@"

    Dim lastRoleIndex As Long
    lastRoleIndex = 2

    For i = 2 To uBImportData
        Dim roleExisting As Boolean
        roleExisting = False

        For rIndex = 2 To maxRole

            If ws.Cells(rIndex, 2) = ImportData(i, 3) And ws.Cells(rIndex, 1) = ImportData(i, 1) Then
                ws.Cells(rIndex, 3) = ws.Cells(rIndex, 3) + 1

                If ImportData(i, 4) = 0 Then
                    ws.Cells(rIndex, 6) = ws.Cells(rIndex, 6) + 1
                    If ws.Cells(rIndex, 4) = "" Then
                        ws.Cells(rIndex, 4) = ImportData(i, 2)
                    ElseIf InStr(ws.Cells(rIndex, 4), ImportData(i, 2)) < 1 Then

                        ws.Cells(rIndex, 4) = ws.Cells(rIndex, 4) & ", " & ImportData(i, 2)
                    End If
                End If

                If ImportData(i, 4) > 0 And ImportData(i, 4) < 0.5 Then
                    ws.Cells(rIndex, 7) = ws.Cells(rIndex, 7) + 1
                    If ws.Cells(rIndex, 5) = "" Then
                        ws.Cells(rIndex, 5) = ImportData(i, 2)
                    ElseIf InStr(ws.Cells(rIndex, 5), ImportData(i, 2)) < 1 Then
                        ws.Cells(rIndex, 5) = ws.Cells(rIndex, 5) & "," & ImportData(i, 2)
                    End If
                End If

                roleExisting = True
            End If
        Next rIndex

        If ws.Cells(lastRoleIndex, 2) = "" And roleExisting = False Then
            ws.Cells(lastRoleIndex, 1) = ImportData(i, 1)
            ws.Cells(lastRoleIndex, 2) = ImportData(i, 3)
            ws.Cells(lastRoleIndex, 3) = 1

            If ImportData(i, 4) = 0 Then
                ws.Cells(lastRoleIndex, 6) = 1
                ws.Cells(lastRoleIndex, 7) = 0
                ws.Cells(lastRoleIndex, 4) = ImportData(i, 2)
            End If

            If ImportData(i, 4) > 0 And ImportData(i, 4) < 0.5 Then
                ws.Cells(lastRoleIndex, 7) = 1
                ws.Cells(lastRoleIndex, 6) = 0
                ws.Cells(lastRoleIndex, 5) = ImportData(i, 2)
            End If
            lastRoleIndex = lastRoleIndex + 1
        End If
    Next i

End Function


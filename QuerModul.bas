Attribute VB_Name = "QuerModul"
Option Explicit
Public TblRange1 As Range
Public TblRange2 As Range
Public FieldType_table1() As New FieldType_Capt
Public FieldType_table2() As New FieldType_Capt
Public Command_Labels() As New CommandLabel_Class
Public OutHat As String
Public tableStruct2(), tableStruct()
Public startSheet As Worksheet
Sub StartQuerModule()
    Dim AutoCalcWas As Boolean, CalcWas
    Dim tempWsh As Worksheet, OutWsh As Worksheet
    Dim wb As Workbook
    Dim Tb1_Start As Long, Tb1_End As Long, Tb2_Start As Long, Tb2_End As Long
    Dim TableRows1 As Long, TableRows2 As Long
    Dim Table1RangeString As String, Table2RangeString As String
    'Vars for Query
    Dim Zs As String, Zs_origin As String
    Dim ADO As New ADO
    Dim i As Long, ArrAm(), r2 As Range, RowCount As Long, j As Long, k As Long, ColumnsCount As Long
    Dim splitArr() As String
    
    CalcWas = Application.Calculation
    AutoCalcWas = Application.CalculateBeforeSave
    
    Set startSheet = ActiveSheet
    Set wb = ActiveWorkbook
    QuerForm.Show
    TurnMeOff True, 0, 0
    Zs = QuerForm.TextBox1.Value
    Unload QuerForm
    'Stop
    Application.StatusBar = "Set data on query sheet"
    ThisWorkbook.Activate
    If WorksheetIsExist("tmp", 1) Then
        Set tempWsh = ActiveWorkbook.Worksheets("tmp")
        tempWsh.Activate
        tempWsh.cells.Clear
    Else
        Set tempWsh = ActiveWorkbook.Worksheets.Add
        tempWsh.Move Before:=Sheets(1)
        tempWsh.Name = "tmp"
    End If
    For i = 1 To UBound(tableStruct, 2)
        If tableStruct(2, i) = "Number" Then
            tempWsh.Columns(i).NumberFormat = "#,##0.00"
        Else
            If tableStruct(2, i) = "Date" Then
                tempWsh.Columns(i).NumberFormat = "m/d/yyyy"
            Else
                tempWsh.Columns(i).NumberFormat = "@"
            End If
        End If
    Next
    'tempWsh.Rows(1).NumberFormat = "@"
    wb.Activate
    If TblRange1.cells.CountLarge > 100000 Then
        TblRange1.Offset(1, 0).Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Copy
        tempWsh.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    Else
        tempWsh.Range("A1").Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value = TblRange1.Offset(1, 0).Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value2
    End If
    'tempWsh.Range("A1").Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value = TblRange1.Offset(1, 0).Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value
    Tb1_Start = 1
    Tb1_End = TblRange1.Columns.count
    TableRows1 = TblRange1.Rows.count - 1
    Table1RangeString = Range(cells(1, Tb1_Start), cells(TableRows1, Tb1_End)).Address(0, 0)
    If Not TblRange2 Is Nothing Then
        Tb2_Start = TblRange1.Columns.count + 2
        Tb2_End = TblRange1.Columns.count + 1 + TblRange2.Columns.count
        TableRows2 = TblRange2.Rows.count - 1
        Table2RangeString = Range(cells(1, Tb2_Start), cells(TableRows2, Tb2_End)).Address(0, 0)
        For i = 1 To UBound(tableStruct2, 2)
            If tableStruct2(2, i) = "Number" Then
                tempWsh.Columns(Tb2_Start + i - 1).NumberFormat = "#,##0.00"
            Else
                If tableStruct2(2, i) = "Date" Then
                    tempWsh.Columns(Tb2_Start + i - 1).NumberFormat = "m/d/yyyy"
                Else
                    tempWsh.Columns(Tb2_Start + i - 1).NumberFormat = "@"
                End If
            End If
        Next
        If TblRange2.cells.CountLarge > 100000 Then
            TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Copy
            tempWsh.Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Else
            tempWsh.Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value = TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value
        End If
        'tempWsh.Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value = TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value

    End If
    Call ParseQueryHat(Zs)
    
    Set TblRange1 = Nothing
    Set TblRange2 = Nothing     ' Здесь освободим память от диапазонов
    ADO.Header = False
    ADO.DataSource = ThisWorkbook.path & "\" & ThisWorkbook.Name
    
    Zs_origin = Zs
    Zs = UCase(Zs)
    Zs = Zs & ";"
    Zs = Replace(Zs, "TABLE1", "[" & Table1RangeString & "]", 1)
    Zs = Replace(Zs, "TABLE2", "[" & Table2RangeString & "]", 1)
    On Error Resume Next
    ThisWorkbook.RemovePersonalInformation = True
    ThisWorkbook.Save
    ThisWorkbook.Saved = True
    DoEvents
    Application.StatusBar = "Performing query"
    ADO.Query Trim(Zs)
    If err.Number > 0 Then
        MsgBox "Uncorrect Query, change query and try again"
        tempWsh.Delete
    Else
        On Error GoTo 0
        If WorksheetIsExist("Result") Then
            Set OutWsh = ActiveWorkbook.Worksheets("Result")
            OutWsh.UsedRange.cells.Delete
            OutWsh.cells.NumberFormat = "#,##0.00"
            OutWsh.Activate
        Else
            Set OutWsh = ActiveWorkbook.Worksheets.Add
            OutWsh.Name = "Result"
            OutWsh.cells.NumberFormat = "#,##0.00"
        End If
        If Not ADO.Recordset.EOF And ADO.Recordset.RecordCount > 0 Then
            DoEvents
            Application.StatusBar = "Put data on Result sheet"
            OutWsh.UsedRange.Clear
            On Error Resume Next
            err.Clear
            ArrAm = ADO.ToArray
            If err.Number > 0 Then
                On Error GoTo 0
                ADO.Recordset.MoveFirst
                err.Clear
                OutWsh.Range("A2").CopyFromRecordset ADO.Recordset 'вывод
                If err.Number > 0 Then
                    MsgBox "Не хватает памяти!"
                    GoTo AdoDone
                End If
                ColumnsCount = ADO.Recordset.Fields.count
            Else
                OutWsh.Range("A2").Resize(UBound(ArrAm, 1), UBound(ArrAm, 2)) = ArrAm
                ColumnsCount = UBound(ArrAm, 2)
            End If
            On Error GoTo 0
            
            RowCount = OutWsh.UsedRange.Rows.count + 1
            For i = 1 To ColumnsCount
                Set r2 = Range(OutWsh.cells(2, i), OutWsh.cells(RowCount, i))
                j = r2.cells.count - Application.count(r2) - Application.WorksheetFunction.CountBlank(r2)
                If j > 0 Then
                    r2.NumberFormat = "@"
                Else
                    If IsDate(r2.cells(1)) Then
                        r2.NumberFormat = "m/d/yyyy"
                    Else
                        r2.NumberFormat = "#,##0.00"
                    End If
                End If
            Next
            
            If OutHat <> "" Then
                splitArr() = Split(OutHat, vbTab)
                For i = 0 To UBound(splitArr, 1)
                    OutWsh.cells(1, 1 + i).Value = splitArr(i)
                Next
                For j = 0 To UBound(splitArr, 1)
                    k = 0
                    If Not ((Not tableStruct) = -1) Then
                        For i = 1 To UBound(tableStruct, 2)
                            If k = 1 Then Exit For
                            If tableStruct(1, i) = splitArr(j) Then
                                If tableStruct(2, i) = "Number" Then
                                    OutWsh.Columns(j + 1).NumberFormat = "#,##0.00"
                                Else
                                    If tableStruct(2, i) = "Date" Then
                                        OutWsh.Columns(j + 1).NumberFormat = "m/d/yyyy"
                                    Else
                                        OutWsh.Columns(j + 1).NumberFormat = "@"
                                    End If
                                End If
                                k = 1
                            End If
                        Next
                    End If
                    If Not ((Not tableStruct2) = -1) Then
                        For i = 1 To UBound(tableStruct2, 2)
                            If k = 1 Then Exit For
                            If tableStruct2(1, i) = splitArr(j) Then
                                If tableStruct2(2, i) = "Number" Then
                                    OutWsh.Columns(j + 1).NumberFormat = "#,##0.00"
                                Else
                                    If tableStruct2(2, i) = "Date" Then
                                        OutWsh.Columns(j + 1).NumberFormat = "m/d/yyyy"
                                    Else
                                        OutWsh.Columns(j + 1).NumberFormat = "@"
                                    End If
                                End If
                                k = 1
                            End If
                        Next
                    End If
                Next
            End If
            If WorksheetIsExist("UsedQueries", 1) Then
                i = ThisWorkbook.Worksheets("UsedQueries").UsedRange.Rows.count + 1
                ThisWorkbook.Worksheets("UsedQueries").Range("A" & CStr(i)) = Now()
                ThisWorkbook.Worksheets("UsedQueries").Range("B" & CStr(i)) = Zs_origin
                ThisWorkbook.Worksheets("UsedQueries").PivotTables("QueryPivot").RefreshTable
                ThisWorkbook.Worksheets("UsedQueries").PivotTables("QueryPivot").Update
            End If
            If OutWsh.AutoFilterMode = False Then
                Range(OutWsh.cells(1, 1), OutWsh.cells(1, OutWsh.UsedRange.Columns.count)).AutoFilter
                OutWsh.Range("A1").Select
                OutWsh.Rows(2).Select: ActiveWindow.FreezePanes = True
                OutWsh.Range("A1").Select
            End If
            With Range(OutWsh.cells(1, 1), OutWsh.cells(1, OutWsh.UsedRange.Columns.count))
                .Interior.Color = RGB(Int((100 * Rnd) + 150), Int((100 * Rnd) + 150), Int((100 * Rnd) + 150))
                .Font.Bold = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
            OutWsh.cells(1, OutWsh.UsedRange.Columns.count + 2) = "The Query was: " & Zs
        End If

AdoDone:
        ADO.Destroy
        tempWsh.Delete
        MsgBox "Done!"
    End If
    Application.StatusBar = False
    TurnMeOff False, CalcWas, AutoCalcWas
End Sub
Sub TurnMeOff(Typ As Boolean, CalcWas, AutoCalcWas)
    If Typ Then
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            .Calculation = xlManual
            .CalculateBeforeSave = False
        End With
    Else
        With Application
            .CutCopyMode = False
            .DisplayAlerts = True
            .ScreenUpdating = True
            .StatusBar = False
            .EnableEvents = True
            .Calculation = CalcWas
            .CalculateBeforeSave = AutoCalcWas
        End With
    End If
End Sub

Public Function TestQuer(Zs As String) As String
    Dim tempWsh As Worksheet, OutWsh As Worksheet
    Dim Tb1_Start As Long, Tb1_End As Long, Tb2_Start As Long, Tb2_End As Long
    Dim TableRows1 As Long, TableRows2 As Long, i As Long, j As Long
    Dim Table1RangeString As String, Table2RangeString As String
    
    'Vars for Query
    Dim ADO As New ADO
    Dim AutoCalcWas As Boolean, CalcWas
    CalcWas = Application.Calculation
    AutoCalcWas = Application.CalculateBeforeSave
    Set OutWsh = ActiveSheet
    Application.StatusBar = "Testing query"
    TurnMeOff True, 0, 0
    
    Set tempWsh = ActiveWorkbook.Worksheets.Add

    tempWsh.Move Before:=Sheets(1)
    For i = 1 To UBound(tableStruct, 2)
        If tableStruct(2, i) = "Number" Then
            tempWsh.Columns(i).NumberFormat = "#,##0.00"
        Else
            If tableStruct(2, i) = "Date" Then
                tempWsh.Columns(i).NumberFormat = "m/d/yyyy"
            Else
                tempWsh.Columns(i).NumberFormat = "@"
            End If
        End If
    Next
    'Stop
'    i = TblRange1.Rows.count - 1
'    j = TblRange1.Columns.count - 1
'    Set TblRange1 = Range(startSheet.cells(TblRange1.cells(1).Row + 1, TblRange1.cells(1).Column), startSheet.cells(TblRange1.cells(1).Row + 1 + i, TblRange1.cells(1).Column + j))
'
'    Range(tempWsh.cells(1, 1), tempWsh.cells(i, j)).Value = TblRange1.Value
    
    'Stop
    If TblRange1.cells.CountLarge > 100000 Then
        TblRange1.Offset(1, 0).Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Copy
        tempWsh.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    Else
        tempWsh.Range("A1").Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value = TblRange1.Offset(1, 0).Resize(TblRange1.Rows.count - 1, TblRange1.Columns.count).Value2
    End If
    
    Tb1_Start = 1
    Tb1_End = TblRange1.Columns.count
    TableRows1 = TblRange1.Rows.count - 1
    Table1RangeString = Range(cells(1, Tb1_Start), cells(TableRows1, Tb1_End)).Address(0, 0)
    If Not TblRange2 Is Nothing Then
        If TblRange2.cells.CountLarge > 100000 Then
            TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Copy
            Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        Else
            tempWsh.Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value = TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value
        End If
        'tempWsh.Range(tempWsh.cells(1, TblRange1.Columns.count + 2).Address).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value = TblRange2.Offset(1, 0).Resize(TblRange2.Rows.count - 1, TblRange2.Columns.count).Value
        Tb2_Start = TblRange1.Columns.count + 2
        Tb2_End = TblRange1.Columns.count + 1 + TblRange2.Columns.count
        TableRows2 = TblRange2.Rows.count - 1
        Table2RangeString = Range(cells(1, Tb2_Start), cells(TableRows2, Tb2_End)).Address(0, 0)
    End If
    ADO.Header = False
    ADO.DataSource = ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    
    Zs = UCase(Zs)
    Zs = Zs & ";"
    Zs = Replace(Zs, "TABLE1", "[" & Table1RangeString & "]", 1)
    Zs = Replace(Zs, "TABLE2", "[" & Table2RangeString & "]", 1)
    On Error Resume Next
    ActiveWorkbook.RemovePersonalInformation = True
    ActiveWorkbook.Save
    ActiveWorkbook.Saved = True
    ADO.Query Trim(Zs)
    If err <> 0 Then TestQuer = err.Description: err.Clear
    On Error GoTo 0

    ADO.Destroy
    tempWsh.Delete
    ActiveWorkbook.Save
    OutWsh.Activate
    Application.StatusBar = False
End Function
Sub ParseQueryHat(ByRef Zs As String)
    Dim myTxtval As String, splitArr() As String, myTxtVal2 As String
    Dim ReplString As String, ReplString2 As String
    Dim i As Long, j As Long, k As Long, maxColumn As Long, n As Long
    Dim Tabl1Syn As String, Tabl2Syn As String
    myTxtval = Trim(Replace(Zs, "select", ""))
    If InStr(1, myTxtval, "table1 as ", vbTextCompare) Then
        j = InStr(1, myTxtval, "table1 as ", vbTextCompare) + 10
        Tabl1Syn = ""
        For i = j To Len(myTxtval)
            If Mid(myTxtval, i, 1) = Chr(32) Then Exit For
            Tabl1Syn = Tabl1Syn & Mid(myTxtval, i, 1)
        Next
        If Tabl1Syn <> "" Then
            myTxtval = Replace(myTxtval, Tabl1Syn & ".", "table1.")
        End If
    End If
    If InStr(1, myTxtval, "table2 as ", vbTextCompare) Then
        j = InStr(1, myTxtval, "table2 as ", vbTextCompare) + 10
        Tabl2Syn = ""
        For i = j To Len(myTxtval)
            If Mid(myTxtval, i, 1) = Chr(32) Then Exit For
            Tabl2Syn = Tabl2Syn & Mid(myTxtval, i, 1)
        Next
        If Tabl2Syn <> "" Then
            myTxtval = Replace(myTxtval, Tabl2Syn & ".", "table2.")
        End If

    End If
    i = InStr(1, LCase(myTxtval), "from", vbTextCompare)
    j = InStr(i + 1, LCase(myTxtval), " table", vbTextCompare)
    myTxtval = Left(myTxtval, j + 7)
    myTxtval = Trim(Replace(myTxtval, "from", ""))
    myTxtval = Trim(Replace(myTxtval, "order by", ""))
    myTxtval = Trim(Replace(myTxtval, "group by", ""))
    myTxtval = Trim(Replace(myTxtval, "left join", ""))
    myTxtval = Trim(Replace(myTxtval, "on", ""))
    myTxtval = Trim(Replace(myTxtval, "(", ""))
    myTxtval = Trim(Replace(myTxtval, ")", ""))
    myTxtval = Trim(Replace(myTxtval, "where", ""))
    myTxtval = Replace(myTxtval, "=", " ")
    myTxtval = Replace(myTxtval, ",", " ")
    myTxtval = LCase(myTxtval)
    myTxtVal2 = myTxtval

    k = 1
    Do
        i = InStr(1, myTxtVal2, "*", vbTextCompare)
        If i > 0 Then
            j = InStr(i + 1, myTxtVal2, " table", vbTextCompare)
            If j > 0 Then
                ReplString = Mid(myTxtval, j + 1, 6)
                maxColumn = 0
                If ReplString = "table1" Then maxColumn = TblRange1.Columns.count
                If ReplString = "table2" Then maxColumn = TblRange2.Columns.count
                ReplString2 = ""
                For k = 1 To maxColumn
                    ReplString2 = ReplString2 & IIf(ReplString2 = "", "", " ") & ReplString & "." & "F" & CStr(k)
                Next
                myTxtVal2 = Left(myTxtVal2, i - 1) & " " & ReplString2 & " " & Right(myTxtVal2, Len(myTxtVal2) - i)
            Else
                Exit Do
            End If
        End If
    Loop Until i = 0
    myTxtval = myTxtVal2
    OutHat = ""
    splitArr = Split(myTxtval)

    For i = 0 To UBound(splitArr)
        If splitArr(i) Like "*.[f,F][0-9]*" Or splitArr(i) Like "*[f,F][0-9]*" Then     '"?.F[0-9]*"
            If InStr(1, splitArr(i), "table2", vbTextCompare) Then
                n = GetNumFromString(splitArr(i), 1)
                OutHat = OutHat & IIf(OutHat = "", "", vbTab) & TblRange2.cells(1, n)
            Else
                If InStr(1, splitArr(i), "table1", vbTextCompare) Then
                    n = GetNumFromString(splitArr(i), 1)
                    OutHat = OutHat & IIf(OutHat = "", "", vbTab) & Application.WorksheetFunction.Clean(TblRange1.cells(1, n))
                Else
                    If InStr(1, splitArr(i), "table", vbTextCompare) = 0 Then
                        On Error Resume Next
                        n = GetNumFromString(splitArr(i), 0)
                        k = 0
                        For j = i To UBound(splitArr)
                            If InStr(1, splitArr(j), "table1", vbTextCompare) Then k = 1: Exit For
                            If InStr(1, splitArr(j), "table2", vbTextCompare) Then k = 2: Exit For
                        Next
                        If k = 1 Then
                            OutHat = OutHat & IIf(OutHat = "", "", vbTab) & Application.WorksheetFunction.Clean(TblRange1.cells(1, n))
                        End If
                        If k = 2 Then
                            OutHat = OutHat & IIf(OutHat = "", "", vbTab) & Application.WorksheetFunction.Clean(TblRange2.cells(1, n))
                        End If
                        On Error GoTo 0
                    Else

                    End If
                End If
            End If
        End If
    Next
End Sub
Public Function WorksheetIsExist(iName$, Optional s_type As Byte = 0) As Boolean
    On Error Resume Next
    If s_type = 0 Then
        WorksheetIsExist = (TypeName(ActiveWorkbook.Worksheets(iName$)) = "Worksheet")
    Else
        WorksheetIsExist = (TypeName(ThisWorkbook.Worksheets(iName$)) = "Worksheet")
    End If
    On Error GoTo 0
End Function
Public Function GetNumFromString(myStr$, myInd As Byte) As Long
    With CreateObject("VBScript.RegExp")
        .Pattern = "\d+"
        .Global = True
        If .Test(myStr$) Then
            GetNumFromString = .Execute(myStr$)(myInd) '(1)
        End If
    End With
End Function

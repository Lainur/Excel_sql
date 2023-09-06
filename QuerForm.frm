VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuerForm 
   Caption         =   "Составление запроса"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   19380
   OleObjectBlob   =   "QuerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public T1_init As Boolean, T2_init As Boolean
Private Sub ComboBox1_Change()
    If Not (TblRange1 Is Nothing And TblRange2 Is Nothing) Then Me.TextBox1.Value = Me.ComboBox1.Value
End Sub
Private Sub CommandButton1_Click()
    Dim r As Range
    Dim wasleft As Long
    wasleft = QuerForm.Left
    QuerForm.Left = 20000
    On Error Resume Next

    Set r = Application.InputBox("Выберите первую таблицу", "Выберите первую таблицу", _
                ActiveCell.Address, Type:=8)
    If r.cells.count = 1 Then
        If r.CurrentRegion.cells.count > 3 Then
            Set TableRange1 = r.CurrentRegion.cells
            Set r = TableRange1
        End If
    Else
        Set TableRange1 = r
    End If
    
    If r.cells.count > 3 Then
        Tabletwo_uni r, 1 'TableOne_sel r
        Me.Label4.Caption = " " & r.Worksheet.Name & "!" & r.Address & ". Rows in Table: " & Format(CStr(r.Rows.count - 1), "# ###")
    End If
    Me.Repaint
    QuerForm.Left = wasleft
    On Error GoTo 0
    Me.TextBox1.SetFocus
End Sub
Private Sub CommandButton2_Click()
    Dim wasleft As Long
    wasleft = QuerForm.Left
    QuerForm.Left = 20000
    Dim r As Range
    On Error Resume Next
    Set r = Application.InputBox("Выберите вторую таблицу", "Выберите вторую таблицу", _
                ActiveCell.Address, Type:=8)
    If r.cells.count = 1 Then
        If r.CurrentRegion.cells.count > 3 Then
            Set TableRange2 = r.CurrentRegion.cells
            Set r = TableRange2
        End If
    Else
        Set TableRange2 = r
    End If
                
    If r.cells.count > 3 Then
        Tabletwo_uni r, 2
        Me.Label5.Caption = " " & r.Worksheet.Name & "!" & r.Address & ". Rows in Table: " & Format(CStr(r.Rows.count - 1), "# ###")
    End If
    Me.Repaint
    QuerForm.Left = wasleft
    On Error GoTo 0
    Me.TextBox1.SetFocus
End Sub
Private Sub CommandButton3_Click()
    Dim response As String
    If TblRange1 Is Nothing And TblRange2 Is Nothing Then Exit Sub
    If Me.CheckBox1.Value Then response = TestQuer(Me.TextBox1.Value)
    If Not (response = vbNullString) Then
        Me.Warning_lbl.Caption = response & "!"
    Else
        If Not (TblRange1 Is Nothing And TblRange2 Is Nothing) Then
            Me.Hide
        Else
            Me.Warning_lbl.Caption = "need to select at least one table for queries"
        End If
    End If
End Sub
Private Sub CommandButton4_Click()
    QuerHlpFrm.Show
End Sub
Private Sub CommandButton8_Click()
    Dim CurrTop As Long
    T1_init = False
    For Each ctr In Me.Controls
        If TypeName(ctr) = "Label" Then
            If (ctr.Name Like "*Tbl1_num*" Or ctr.Name Like "*Tbl1_hat*" Or ctr.Name Like "*Tbl1_data*" Or ctr.Name Like "*Tbl1_ind*") Then
                Me.Controls.Remove (ctr.Name)
            Else
                If CurrTop < ctr.Top + ctr.Height Then CurrTop = ctr.Top + ctr.Height
            End If
        End If
    Next
    CurrTop = CurrTop + 40
    Me.Height = CurrTop
    Set TblRange1 = Nothing
    Me.Label4.Caption = "Rows in Table: "
End Sub
Private Sub CommandButton7_Click()
    Dim CurrTop As Long
    T2_init = False
    Set TblRange2 = Nothing
    For Each ctr In Me.Controls
        If TypeName(ctr) = "Label" Then
            If (ctr.Name Like "*Tbl2_num*" Or ctr.Name Like "*Tbl2_hat*" Or ctr.Name Like "*Tbl2_data*" Or ctr.Name Like "*Tbl2_ind*") Then
                Me.Controls.Remove (ctr.Name)
            Else
                If CurrTop < ctr.Top + ctr.Height Then CurrTop = ctr.Top + ctr.Height
            End If
        End If
    Next
    CurrTop = CurrTop + 40
    Me.Height = CurrTop
    Me.Label5.Caption = "Rows in Table: "
End Sub
Private Sub Label6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.TextBox1 = Replace(Me.TextBox1, " GROUP BY ", "")
End Sub

Private Sub Label4_Click()
    Dim r As Variant, TableRange As Range
    On Error Resume Next

    r = Application.InputBox("Введите диапазон таблицы 1", "Введите диапазон таблицы 1", _
                TblRange1.Worksheet.Name & "!" & TblRange1.Address, Type:=2)
    If err.Number > 0 Then
        MsgBox "You've entered wrong value!"
    Else
        Set TableRange = Range(r)
        If err.Number > 0 Then
            MsgBox "Can't set such table range!"
        Else
            Tabletwo_uni TableRange, 1
        End If
    End If
    err.Clear
    On Error Resume Next
End Sub

Private Sub TabStrip1_Change()
    Select Case Me.TabStrip1.Value
        Case 1
            FillComboBox 0
            Me.Label16.Caption = "5 популярных запросов"
        Case 2
            FillComboBox 2
            Me.Label16.Caption = "Все запросы"
        Case 3
            FillComboBox 3
            Me.Label16.Caption = "Шаблоны"
        Case Else
            FillComboBox 1
            Me.Label16.Caption = "5 последних запросов"
    End Select
End Sub
Private Sub TextBox1_Change()
    Dim myTxtval As String
    Dim splitArr() As String
    Dim i As Long
    Dim j As Long, k As Long
    Dim Tabl1Syn As String, ctr
    myTxtval = LCase(Me.TextBox1.Value)
    
    If Len(myTxtval) > 5 Then
        For Each ctr In Me.Controls
            If TypeName(ctr) = "Label" Then
                If ctr.Name Like "Tbl1_hat*" Or ctr.Name Like "Tbl2_hat*" Then
                    ctr.BackColor = &HFFFFFF
                End If
            End If
        Next
        Me.Alias_txtb_1 = ""
        Me.Alias_txtb_2 = ""
        If InStr(1, myTxtval, "table1 as ", vbTextCompare) Then
            j = InStr(1, myTxtval, "table1 as ", vbTextCompare) + 10
            Tabl1Syn = ""
            For i = j To Len(myTxtval)
                If Mid(myTxtval, i, 1) = Chr(32) Then Exit For
                Tabl1Syn = Tabl1Syn & Mid(myTxtval, i, 1)
            Next
            If Tabl1Syn <> "" Then
                myTxtval = Replace(myTxtval, Tabl1Syn & ".", "table1.")
                Me.Alias_txtb_1 = Tabl1Syn
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
                Me.Alias_txtb_2 = Tabl2Syn
            End If

        End If
        
        myTxtval = Trim(Replace(myTxtval, "select", ""))
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
        myTxtval = Replace(myTxtval, "max", "")
        myTxtval = Replace(myTxtval, "min", "")
        myTxtval = Replace(myTxtval, "sum", "")
        myTxtval = Replace(myTxtval, "avg", "")
        myTxtval = Replace(myTxtval, "count", "")

        splitArr = Split(myTxtval)
        For i = 0 To UBound(splitArr)
            If splitArr(i) Like "*.[f,F][0-9]*" Or splitArr(i) Like "*[f,F][0-9]*" Then     '"?.F[0-9]*"
                'Stop
                If InStr(1, splitArr(i), "table2", vbTextCompare) And Not TblRange2 Is Nothing Then
                    On Error Resume Next
                    n = GetNumFromString(splitArr(i), 1)
                    Me.Controls("Tbl2_hat " & CStr(n)).BackColor = &H80FF80
                    On Error GoTo 0
                Else
                    If InStr(1, splitArr(i), "table1", vbTextCompare) And Not TblRange1 Is Nothing Then
                        On Error Resume Next
                        n = GetNumFromString(splitArr(i), 1)
                        Me.Controls("Tbl1_hat " & CStr(n)).BackColor = &H80FF80
                        On Error GoTo 0
                    Else
                        If InStr(1, splitArr(i), "table", vbTextCompare) = 0 Then
                            On Error Resume Next
                            n = GetNumFromString(splitArr(i), 0)
                            k = 0
                            For j = i To UBound(splitArr)
                                If InStr(1, splitArr(j), "table1", vbTextCompare) Then k = 1: Exit For
                                If InStr(1, splitArr(j), "table2", vbTextCompare) Then k = 2: Exit For
                            Next
                            If k = 1 And Not TblRange1 Is Nothing Then
                                Me.Controls("Tbl1_hat " & CStr(n)).BackColor = &H80FF80
                            End If
                            If k = 2 And Not TblRange2 Is Nothing Then
                                Me.Controls("Tbl2_hat " & CStr(n)).BackColor = &H80FF80
                            End If
                            On Error GoTo 0
                        Else

                        End If
                    End If
                End If
            End If
        Next
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim TableRange As Range, ctr
    Dim i As Long
    If ActiveCell.CurrentRegion.Rows.count > 3 And ActiveCell.CurrentRegion.Columns.count > 3 Then
        ActiveCell.CurrentRegion.cells.Select
        Set TableRange = Selection
        Call Tabletwo_uni(TableRange, 1)
        Me.Label4.Caption = " " & TableRange.Worksheet.Name & "!" & TableRange.Address & ". Rows in Table: " & Format(CStr(TableRange.Rows.count - 1), "# ###")
    End If
    For Each ctr In Me.Controls
        If ctr.Name Like "Cmd_Lbl*" Or ctr.Name Like "Cmd_a_Lbl*" Then
            i = i + 1
            ReDim Preserve Command_Labels(1 To i)
            Set Command_Labels(i).CommandLabel = ctr
        End If
    Next
    TextBox1_Change
    FillComboBox 1
End Sub
'Private Sub Reorder_Fs(val As Byte)
'    Dim i As Long
'    For i = 1 To 13
'        Me.Controls("Tbl" & CStr(val) & "_l" & CStr(i)).Visible = False
'    Next
'End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then End
End Sub
Private Sub Warning_lbl_Click()
    Me.Warning_lbl.Caption = vbNullString
End Sub
Public Sub Tabletwo_uni(ByRef r As Range, index As Byte)
    Dim i As Long
    Dim CellContent, ctr
    Dim MidWidth As Long, RightWidth As Long
    i = 0
    If index = 1 Then CurrLeft = 18 Else CurrLeft = 570
 
    MidWidth = 180
    RightWidth = 180
    CurrTop = 90
    If index = 1 Then tableStruct = r.Rows("1:3").Value
    If index = 2 Then tableStruct2 = r.Rows("1:3").Value
    For Each ctr In Me.Controls
        If TypeName(ctr) = "Label" Then
            If (ctr.Name Like "*Tbl2_num*" Or ctr.Name Like "*Tbl2_hat*" Or ctr.Name Like "*Tbl2_data*" Or ctr.Name Like "*Tbl2_ind*") And index = 2 Then
                Me.Controls.Remove (ctr.Name)
            End If
            If (ctr.Name Like "*Tbl1_num*" Or ctr.Name Like "*Tbl1_hat*" Or ctr.Name Like "*Tbl1_data*" Or ctr.Name Like "*Tbl1_ind*") And index = 1 Then
                Me.Controls.Remove (ctr.Name)
            End If
        End If
    Next

    For Each cell In r.Rows(1).cells
        i = i + 1
        CellContent = cell.Offset(1, 0)
'        If Not (CellContent = vbNullString) Then 'Not (CellContent = vbNullString)
            With Me.Controls.Add(bstrProgID:="Forms.Label.1", Name:="Tbl" & CStr(index) & "_num " & i, Visible:=True)
                .Height = 12
                .Left = CurrLeft - 10
                .Top = CurrTop
                .Width = 9
                .BackStyle = 1
                .BackColor = &HC0E0FF
                .TextAlign = 2
                .WordWrap = False
                .BorderStyle = 1
                Set r2 = r.Columns(i) 'Range(r.Cells(2, cell.Column), r.Cells(r.Rows.Count, cell.Column))
                Set r2 = r2.Offset(1, 0).Resize(r2.Rows.count - 1)
                j = r2.cells.count - Application.count(r2) - Application.WorksheetFunction.CountBlank(r2)
                If j > 0 Then
                    If index = 1 Then tableStruct(2, i) = "Characteristic"
                    If index = 2 Then tableStruct2(2, i) = "Characteristic"
                    .ControlTipText = "This column contains characters"
                    .Caption = "@"
                Else
                    If IsDate(r2.cells(1)) Then
                        If index = 1 Then tableStruct(2, i) = "Date"
                        If index = 2 Then tableStruct2(2, i) = "Date"
                        .BackColor = &H80000002
                        .ControlTipText = "This column contains date values"
                        .Caption = "d"
                    Else
                        If index = 1 Then tableStruct(2, i) = "Number"
                        If index = 2 Then tableStruct2(2, i) = "Number"
                        .BackColor = &HFFFFC0 '&H80000002&
                        .ControlTipText = "This column contains numeric values"
                        .Caption = "#"
                    End If
                End If
            End With
            With Me.Controls.Add(bstrProgID:="Forms.Label.1", Name:="Tbl" & CStr(index) & "_ind " & i, Visible:=True)
                .Height = 12
                .Left = CurrLeft
                .Top = CurrTop
                .Width = 24
                .BackStyle = 1
                .BackColor = &H80000003
                .TextAlign = 2
                .Caption = "F" & CStr(i)
                .WordWrap = False
                .BorderStyle = 1
            End With
            With Me.Controls.Add(bstrProgID:="Forms.Label.1", Name:="Tbl" & CStr(index) & "_hat " & i, Visible:=True)
                'Stop
                .Height = 12
                .Left = CurrLeft + 25
                .Top = CurrTop
                .Width = MidWidth
                .BackStyle = 1
                .BackColor = &HFFFFFF
                .TextAlign = 1
                .Caption = cell.Value
                .WordWrap = False
                .BorderStyle = 1
            End With
            With Me.Controls.Add(bstrProgID:="Forms.Label.1", Name:="Tbl" & CStr(index) & "_data " & i, Visible:=True)
                .Height = 12
                .Left = CurrLeft + MidWidth + 25 + 1
                .Top = CurrTop
                .Width = RightWidth
                .BackStyle = 1
                .BackColor = &HFFFFFF
                .TextAlign = 1
                .Caption = CellContent
                .WordWrap = False
                .BorderStyle = 1
            End With
            
'            With Me.Controls.Add(bstrProgID:="Forms.Checkbox.1", Name:="Checkbox1 " & i, Visible:=True)
'                .Height = 20
'                .Left = Currleft + 34
'                .Top = CurrTop + Me.Controls("NameRow1 " & i).Height + 4
'                .Width = 20
'                If IsNumeric(CellContent) And TypeName(CellContent) <> "String" Then
'                    .Enabled = False: .Visible = False
'                Else
'                    If j = 0 Then .Value = True: j = 1: ThisOneIsBad = True
'                End If
'            End With
'            With Me.Controls.Add(bstrProgID:="Forms.OptionButton.1", Name:="BaseOpt1 " & i, Visible:=True)
'                .Height = 20
'                .Left = Currleft + 34
'                .Top = CurrTop + Me.Controls("NameRow1 " & i).Height + 4 + 16
'                .Width = 20
'                .GroupName = "1"
'                If IsNumeric(CellContent) And TypeName(CellContent) <> "String" Then
'                    .Enabled = False: .Visible = False
'                Else
'                    If ThisOneIsBad Then
'                        .Enabled = False
'                        ThisOneIsBad = Not (ThisOneIsBad)
'                    Else
'                        .Value = True
'                    End If
'                End If
'            End With
'            With ABCform.Controls.Add(bstrProgID:="Forms.ToggleButton.1", Name:="UpControl " & i, Visible:=True)    'создание имен файлов
'            End With
            CurrTop = CurrTop + 14
            'me.Width = Currleft
        'End If
    Next
    CurrTop = 0
    For Each ctr In Me.Controls
        If ctr.Name Like "Tbl1_num*" Or ctr.Name Like "Tbl1_ind*" Then
            Cb_Count = Cb_Count + 1
            ReDim Preserve FieldType_table1(1 To Cb_Count)
            Set FieldType_table1(Cb_Count).FieldTypeLabel = ctr
        End If
        If ctr.Name Like "Tbl2_num*" Or ctr.Name Like "Tbl2_ind*" Then
            Cb_Count = Cb_Count + 1
            ReDim Preserve FieldType_table2(1 To Cb_Count)
            Set FieldType_table2(Cb_Count).FieldTypeLabel = ctr
        End If
        If CurrTop < ctr.Top + ctr.Height Then CurrTop = ctr.Top + ctr.Height
    Next ctr
    
    CurrTop = CurrTop + 40
    Me.Height = CurrTop

    If index = 1 Then Set TblRange1 = r
    If index = 2 Then Set TblRange2 = r
    If r.cells.count > 100000 Then Me.CheckBox1.Value = False
End Sub
Sub FillComboBox(MyType As Byte)
    Dim i As Long, j As Long
    Me.ComboBox1.Clear
    
    If WorksheetIsExist("UsedQueries", 1) Then
        Select Case MyType
            Case 1
                i = ThisWorkbook.Worksheets("UsedQueries").UsedRange.Rows.count
                For j = i To i - 4 Step -1
                    If j < 2 Then Exit For
                    Me.ComboBox1.AddItem (ThisWorkbook.Worksheets("UsedQueries").Range("B" & CStr(j)))
                Next
            Case 0
                For j = 2 To 6
                    If ThisWorkbook.Worksheets("UsedQueries").Range("N" & CStr(j)) = vbNullString Then Exit For
                    Me.ComboBox1.AddItem (ThisWorkbook.Worksheets("UsedQueries").Range("M" & CStr(j)))
                Next
            Case 3
                i = ThisWorkbook.Worksheets("UsedQueries").cells(Rows.count, 16).End(xlUp).row
                For j = 2 To i
                    Me.ComboBox1.AddItem (ThisWorkbook.Worksheets("UsedQueries").Range("P" & CStr(j)))
                Next
            Case Else
                i = ThisWorkbook.Worksheets("UsedQueries").cells(Rows.count, 14).End(xlUp).row
                For j = 2 To i
                    Me.ComboBox1.AddItem (ThisWorkbook.Worksheets("UsedQueries").Range("M" & CStr(j)))
                Next
        End Select

        If Me.ComboBox1.ListCount > 0 Then Me.ComboBox1.ListIndex = 0 Else Me.ComboBox1.ListIndex = -1
    End If
End Sub

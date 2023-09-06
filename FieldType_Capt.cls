VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldType_Capt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents FieldTypeLabel As MSForms.Label
Attribute FieldTypeLabel.VB_VarHelpID = -1
 
Private Sub FieldTypeLabel_Click()

    Dim myIndex As Long, myIndex2 As Long
      
    myIndex = CInt(Right(FieldTypeLabel.Name, Len(FieldTypeLabel.Name) - InStr(1, FieldTypeLabel.Name, " ", vbTextCompare)))
    If InStr(1, FieldTypeLabel.Name, "Tbl1", vbTextCompare) Then myIndex2 = 1 Else myIndex2 = 2
    If InStr(1, FieldTypeLabel.Name, "_num", vbTextCompare) > 0 Then
        If FieldTypeLabel.BackColor = &H80000002 Then
            FieldTypeLabel.BackColor = &HFFFFC0
            FieldTypeLabel.ControlTipText = "This column contains numeric values"
            If myIndex2 = 1 Then tableStruct(2, myIndex) = "Number"
            If myIndex2 = 2 Then tableStruct2(2, myIndex) = "Number"
            FieldTypeLabel.Caption = "#"
        Else
            If FieldTypeLabel.BackColor = &HFFFFC0 Then
                FieldTypeLabel.BackColor = &HC0E0FF
                FieldTypeLabel.ControlTipText = "This column contains characters"
                If myIndex2 = 1 Then tableStruct(2, myIndex) = "Characteristic"
                If myIndex2 = 2 Then tableStruct2(2, myIndex) = "Characteristic"
                FieldTypeLabel.Caption = "@"
            Else
                FieldTypeLabel.BackColor = &H80000002
                FieldTypeLabel.ControlTipText = "This column contains date values"
                If myIndex2 = 1 Then tableStruct(2, myIndex) = "Date"
                If myIndex2 = 2 Then tableStruct2(2, myIndex) = "Date"
                FieldTypeLabel.Caption = "d"
            End If
        End If
    Else
        Dim i As Long, j As Long, lpart$, rpart$, mpart$, k As Long
        j = Len(QuerForm.TextBox1)
        i = QuerForm.TextBox1.SelStart
        lpart$ = Left(QuerForm.TextBox1, i)
        If j - i > 0 Then rpart$ = Right(QuerForm.TextBox1, j - i)
        If myIndex2 = 1 And QuerForm.Alias_txtb_1 <> "" Then mpart$ = QuerForm.Alias_txtb_1 & "."
        If myIndex2 = 2 And QuerForm.Alias_txtb_2 <> "" Then mpart$ = QuerForm.Alias_txtb_2 & "."
        mpart$ = mpart$ & FieldTypeLabel.Caption
        QuerForm.TextBox1 = lpart$ & mpart$ & ", " & rpart$
        QuerForm.TextBox1.SelStart = Len(lpart$ & mpart$ & ", ")
        If Right(Trim(QuerForm.TextBox1), 1) = "," Then QuerForm.TextBox1 = Left(Trim(QuerForm.TextBox1), Len(Trim(QuerForm.TextBox1)) - 1)
    End If
End Sub

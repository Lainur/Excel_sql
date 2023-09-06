VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandLabel_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents CommandLabel As MSForms.Label
Attribute CommandLabel.VB_VarHelpID = -1

Private Sub CommandLabel_Click()
    Dim i As Long, j As Long, lpart$, rpart$, mpart$, k As Long
    j = Len(QuerForm.TextBox1)
    i = QuerForm.TextBox1.SelStart
    k = QuerForm.TextBox1.SelLength
    lpart$ = Left(QuerForm.TextBox1, i)
    If CommandLabel.Name Like "Cmd_a_Lbl*" Then
        If k > 0 Then mpart$ = Mid(QuerForm.TextBox1, i + 1, k)
        If j - i - k > 0 Then rpart$ = Right(QuerForm.TextBox1, j - i - k)
        mpart$ = CommandLabel.Caption & "(" & mpart$ & ")"
        QuerForm.TextBox1 = lpart$ & mpart$ & rpart$
        QuerForm.TextBox1.SelStart = Len(lpart$ & mpart$)
    Else
        If j - i > 0 Then rpart$ = Right(QuerForm.TextBox1, j - i)
        QuerForm.TextBox1 = lpart$ & " " & CommandLabel.Caption & " " & rpart$
        QuerForm.TextBox1.SelStart = Len(lpart$ & " " & CommandLabel.Caption & " ")
    End If
End Sub

Private Sub CommandLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If CommandLabel.Name Like "Cmd_Lbl*" Then
        QuerForm.TextBox1 = Trim(Replace(QuerForm.TextBox1, CommandLabel.Caption, ""))
    End If
End Sub

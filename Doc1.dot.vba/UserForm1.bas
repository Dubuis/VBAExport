Private msg As String

Private Sub CommandButton3_Click()
    If msg = "TOTO" Then
        CommandButton3.Caption = "TOTO"
        msg = "TATA"
    ElseIf msg = "TATA" Then
        CommandButton3.Caption = "TATA"
        msg = "TOTO"
    End If
End Sub

Private Sub UserForm_Initialize()
    If msg <> "TATA" Then
        msg = "TOTO"
    End If
End Sub

Private Sub CommandButton1_Click()
    UserForm1.Hide
End Sub

Private Sub CommandButton2_Click()
    MsgBox msg
End Sub

Public Function tata()
    msg = "TATA"
End Function

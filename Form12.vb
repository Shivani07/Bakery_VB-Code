Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
Label1.Caption = Val(Form10.Text3.Text)
End Sub

Private Sub Timer2_Timer()
Label3.Left = Label3.Left - 8
End Sub



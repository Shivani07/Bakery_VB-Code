Private Sub close_Click()
End
End Sub
Private Sub Command2_Click()
Form1.Show
Form12.Hide
End Sub
Private Sub Command4_Click()
If Text1.Text = "1234" And Text2.Text = "Shivani" And Text3.Text = "1234" Then
Form11.Show
Form12.Hide
Else
a = MsgBox("Wrong Information", vbOKOnly + vbCritical, "Error")
If a = 1 Then Form12.Show
End If
End Sub
Private Sub Form_paint()
Text1.SetFocus
End Sub


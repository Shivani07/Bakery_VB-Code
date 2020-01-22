
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New Command
Dim st As String
Dim st1 As String

Private Sub Form_Load()
con.ConnectionString = "driver={SQL Server};server=USER-PC\SQLEXPRESS;database=bakery"
con.Open
rs.Open "select * from Product", con, adOpenDynamic, adLockOptimistic
rs.ActiveConnection = con
cmd.ActiveConnection = con
rs.Open "select * from Product", con, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
List1.Clear
List3.Clear
List2.Clear
List1.Text = "    ITEMS PURCHASED"
Text2.Text = "0"
Text3.Text = "0"

End Sub

Private Sub Command3_Click()
For i = 0 To 1000
If List1.ListIndex = i Then
List1.RemoveItem (i)
List2.RemoveItem (i)
List3.RemoveItem (i)
End If
Next
If List3.ListCount = 0 Then
Text2.Text = "0": List1.Text = "    ITEMS PURCHASED"
Else
For i = 0 To List3.ListCount - 1
r = r + Val(List3.List(i))
c = c + Val(List2.List(i))
Text3.Text = c
Text2.Text = ""
Text2.Text = r
Next
List1.Text = "    ITEMS PURCHASED"
End If
End Sub

Private Sub Command4_Click()
Dim i As Integer, j As Integer, k As Integer
j = 0
i = List1.ListCount
For j = 0 To i
st = List1.List(j)
Label4.Caption = st
Text1.Text = List2.List(j)
k = Val(Text1.Text)

cmd.CommandText = "update Product set stock=32 where p_name='Butter Cookies';"
cmd.CommandText = "select * from Product where p_name=st"
rs.Update rs.Fields(3), 42
rs.MoveNext
cmd.CommandText = "update Product set stock=stock-st1 where p_name=st"
rs.Update(stock, 2) = rs.Fields(3)
MsgBox (rs.Fields(3))
Next
MsgBox ("Updated")
Form4.Show
Form10.Hide
End Sub

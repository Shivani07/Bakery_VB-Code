Private Sub Command1_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 30
Form10.List1.AddItem ("Veg Burger")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command15_Click(Index As Integer)
Form2.Show
Form8.Hide
End Sub
Private Sub Command2_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 65
Form10.List1.AddItem ("White Sauce Pasta")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command3_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 75
Form10.List1.AddItem ("Veggie Heaven")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub

Private Sub Command4_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 80
Form10.List1.AddItem ("Sizzle Burger")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command5_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 75
Form10.List1.AddItem ("Red Sauce Pasta")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command6_Click(Index As Integer)
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 99
Form10.List1.AddItem ("King Burger")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Timer1_Timer()
Image1.Left = Image1.Left - 5
End Sub
Private Sub VScroll1_Change()
Picture1.Top = VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
Picture1.Top = VScroll1.Value
End Sub

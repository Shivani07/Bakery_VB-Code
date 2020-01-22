Private Sub Command1_Click()
Form2.Show
Form3.Hide
End Sub

Private Sub Command10_Click()
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem (u * 600)
Form10.List1.AddItem ("French Sizzle")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command11_Click()
Dim u As Integer, cost As Integer
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem (u * 600)
Form10.List1.AddItem ("Fruit Cake")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command2_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Chocolate Treat")
u = 0
cost = 0
For c1 = 0 To Form10.List2.ListCount
cost = cost + Val(Form10.List2.List(c1))
u = u + Val(Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub

Private Sub Command3_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Fruity Delight")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command4_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 750
Form10.List1.AddItem ("Dutch Truffle")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command5_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Vanilla Treat")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command6_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Black Forest")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command7_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 750
Form10.List1.AddItem ("Red Velwet")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command8_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Fruit Treat")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub
Private Sub Command9_Click()
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem (Val(u))
Form10.List2.AddItem u * 650
Form10.List1.AddItem ("Choco Vanilla")
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

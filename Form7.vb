Private Sub Command1_Click(Index As Integer)
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 32
Form10.List1.AddItem ("Choco Delight")
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
Form7.Hide
End Sub

Private Sub Command2_Click(Index As Integer)
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 35
Form10.List1.AddItem ("Black Muffins")
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
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 55
Form10.List1.AddItem ("Dark Delight")
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
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 40
Form10.List1.AddItem ("Coco Nutties")
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
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 35
Form10.List1.AddItem ("Nutty Almo")
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
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 25
Form10.List1.AddItem ("Nutty Pine")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub

Private Sub Command7_Click(Index As Integer)
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 40
Form10.List1.AddItem ("Royal Bite")
u = 0
cost = 0
For c1 = 0 To Form10.List3.ListCount - 1
cost = cost + (Form10.List2.List(c1))
u = u + (Form10.List3.List(c1))
Next c1
Form10.Text3 = cost
Form10.Text2 = u
End Sub

Private Sub Command8_Click(Index As Integer)
u = InputBox("Enter the quantity", "Quantity")
Form10.List3.AddItem u
Form10.List2.AddItem u * 60
Form10.List1.AddItem ("Choco Hearts")
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
Image1.Left = Image1.Left - 8
End Sub

Private Sub VScroll1_Change()
Picture1.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = VScroll1.Value
End Sub










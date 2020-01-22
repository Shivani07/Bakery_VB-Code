Private Sub close_Click()
End
End Sub

Private Sub Combo1_GotFocus()
Combo1.ForeColor = vbBlack
End Sub

Private Sub Combo2_Click()
Label1.Visible = True
If Combo2.Text = "Agra" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "8, Handicraft Nagar, Fatehabad Road, Agra 0562-4064051-53"
End If
If Combo2.Text = "Ahmedabad" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Bake My Day, Shop No: 8 to 12, 3rd Floor, Alpha One Mall, Near Vastrapur Lake, Ahmedabad"
Text6.Text = "Bake My Day Shop No. 7, Ground Floor, 3rd Eye- Two, Nr. Panchwati Cross, C. G. Road, Ahmedabad"
Text7.Text = "Shapath III, Ground Floor, Satelite, Ahmedabad, 3988-3988 / 6000 9000"
End If
If Combo2.Text = "Allahabad" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Shop No 31/31, Ground floor, S.P Marg, Civil line, Allahabad"
End If
If Combo2.Text = "Amritsar" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "No. 8, Lawrence Road, Amritsar 0183-2560600"
End If
If Combo2.Text = "Bareilly" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "S-43, Phoenix United Mall, 2nd Floor, Philibhit Bye Pass Road, Bareilly"
End If
If Combo2.Text = "Bengaluru" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "# 124 Ground FloorSurya Chambers, Murgeshpalaya, Airport Road, Bengaluru 3988 3988 / 6000 9000"
Text6.Text = "No-8, 4th Floor Ascendas ParkSquare Mall, White Fields Main Road, Bengaluru 080-28026591"
Text7.Text = "Amarjyoti Colony, Domlur Village, Koramangla Intermediate Ring Road, Bengaluru 3988-3988 / 6000 9000"
End If
If Combo2.Text = "Bhopal" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "E - 5/15, Arera Colony, Vitthal Market, Bhopal 07552514187"
End If
If Combo2.Text = "Bhuvaneswar" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Forum Mart, Big Bazaar, Kharval Nagar, Unit Street, Bhuvaneswar 0674-2380811-12"
End If
If Combo2.Text = "Chandigarh" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Elante Mall, Shop no. 320, 3rd Floor, Industrial Area, Phase-1, Chandigarh"
Text6.Text = "PH Chandigarh 35 C - ACDR, SCO-473-474, Sector-35-C, Chandigarh 0172-4028777/4029777"
Text7.Text = "Sec -26, SCO No. 51, Sector-26, Chandigarh 0172-5088666/2793555"
End If
If Combo2.Text = "Chennai" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Flat#46,Door#61,3rd Main Road, Kasturibai Nagar, Chennai 3988-3988 / 60009000"
Text6.Text = "Shop Nos 4,floor 3,107/2, Nelson Manikam Road Aminjikarai, Chennai 3988-3988 / 60009000"
Text7.Text = "Old# 67, New# 41, 4th Avenue, Ashok Nagar, Chennai, Opposite Jawahar Vidyalaya, Chennai 3988-3988 / 60009000"
End If
If Combo2.Text = "Cochin" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "39/2198 Joyland Complex, Darbar Hall Road, Cochin 3988-3988 / 4044111"
Text6.Text = "Third Floor, Gold Souk Grande, Vyttila , Kochi, Cochin 0484-4044557"
Text7.Text = "34/1089, 34/1090 3rd Floor, LULU International Shopping Complex Edapally, Cochin"
End If
If Combo2.Text = "Coimbatore" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "2nd floor, Fun Republic mall, Avinashi Road, Coimbatore"
Text6.Text = "357, Ground Floor, Silver Montt Building, Opp Kennedy Theatre, D.B. Road, R.S. Puram, Coimbatore 3988-3988 / 6000900"
End If
If Combo2.Text = "Dehradun" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Amity Amusement, Eleganza Mall, 542 Rajpur Road, Dehradun 0135-2742740"
End If
If Combo2.Text = "Faridabad" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Ground Floor, Ansal Crowne Plaza, Sector 15 A,NH-28, Faridabad 011-3988-3988 / 6000 9000"
End If
If Combo2.Text = "Ghaziabad" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "Ground Floor, East Delhi Mall, Kaushambi, Ghaziabad 0120-3012200"
Text6.Text = "Shop No D-21 RDC, Raj Nagar, Ghaziabad, Ghaziabad 0120-4121034/0120-4569729"
End If
If Combo2.Text = "Goa" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Amara Bldg, Next to ST. Anthony Chapel, Naika Wad,Calangute main road Bardez, Goa"
End If
If Combo2.Text = "Gurgaon" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Shop No R1, 3rd Floor, Near NH-8 Toll Plaza, Ambience Mall, Gurgaon 0124 - 4665377 / 78"
Text6.Text = "Shop no. 1 and 101, DT City Centre Mall, Mehrauli Gurgaon Road, DLF Phase 2, Gurgaon, Gurgaon 011-3988-3988 / 6000 9000"
Text7.Text = "Ground Floor, Tower C, Unitech Cyber Park, Sec 39, Gurgaon 011-3988 3988/ 6000 9000"
End If
If Combo2.Text = "Guwahati" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Dona planet,Sp no- 1 & 2, Grnd FLR ABC , GS road, Guwahati, Guwahati 2467000 / 2467000"
End If
If Combo2.Text = "Hyderabad" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Road 1 & 12 Junction, Banjara Hills, Hyderabad 3988-3988 / 6000 9000"
Text6.Text = "2nd Floor, Inorbit Mall, Cyberabad, Hyderabad 040-6464 0202"
Text7.Text = "D.No. 1-98/6/51, Shristi Towers, Arunodaya Colony, Madhapur, Hyderabad 3988-3988 / 6000 9000"
End If
If Combo2.Text = "Indore" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "Bake My Day, Shop no 1, Second Floor, road facing, Malhar Mega Mall, AB Road, Vijay Nagar, Indore 3988-3988"
Text6.Text = "Treasure Island Mall, 11 Tukoganj, 11 MG Road, Indore-1, Indore 3988-3988"
End If
If Combo2.Text = "Jaipur" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "GF, Crystal Palm mall, Sardar Patel Marg, Near Sahakar Bhavan, C-scheme, Jaipur 0141-4030461/62"
Text6.Text = "G-3 & 4, Crystal Court, Indra Place, Jaipur 0141-4034779-81"
End If
If Combo2.Text = "Jalandhar" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Near B.M.C Chowk, Next to Kesar Petrol Pump,Jalandhar, Jalandhar 0181-4367777/2227777"
End If
If Combo2.Text = "Jammu" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Ground Floor,Indira Theatre, Canal Road,Opp State Guest House, Jammu 01912503519/20"
End If
If Combo2.Text = "Jodhpur" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Plot No.- 44/1, 5th Avenue, PWD Colony,Jodhpur, Jodhpur 5104500 / 5105500 / 5109600 / 5109700"
End If
If Combo2.Text = "Kanpur" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "3rd Floor, Z square Mall, Mall Road, Kanpur"
End If
If Combo2.Text = "Karnal" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "4th Floor, Sunshine Complex, Karnal, 0528-4500697 / 4500698"
End If
If Combo2.Text = "Kolkata" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "GF Block C, City Centre Mall 2, New Town, Rajarhat, Kolkata 033- 40620073"
Text6.Text = "22 Camac Street, Kolkata 033-22814343"
Text7.Text = "DC-1, Salt Lake City Centre, Sector 1, Kolkata 033-23580984"
End If
If Combo2.Text = "Lucknow" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "Shop no. 8/9/10, West End Mall, Vibhuti Khand, Gomti nagar., Lucknow 0522-6501893 / 6501894"
Text6.Text = "91, Durga Bhawan, Hazarat Ganj, M.G. Road, Lucknow 0522-2235801 / 02 / 03, 4071801, 4022819"
Text7.Text = "Shop no 101 ,First Floor,Saharaganj Mall, Shahnajaf Road, Lucknow +91-522-4082331,+91-522-4082332"
End If
If Combo2.Text = "Ludhiana" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Plot No. 2692, Gurdev Nagar Ferozpur Road, Next to Hotel Park Plaza, Ludhiana, Ludhiana 3988-3988 / 0161-2774555,2774454,2774455"
End If
If Combo2.Text = "Mangalore" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "Shop No. 4, Ground Floor, Bharath Mall Opp to KSRTC Bus Stand, Bejai, Mangalore 3988-3988 / 6000 900"
Text6.Text = "4th Floor, KS Rao Road, Mangalore"
End If
If Combo2.Text = "Manipal" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "No-5, Ground Floor Shiroor Tower, Syndicate Circle Laxminagar, Manipal"
End If
If Combo2.Text = "Mathura" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Hi-Way Plaza, Mathura-Agra Road, Near Govardhan crossing, Mathura 0565-2425424-25"
End If
If Combo2.Text = "Meerut" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "PVS Mall, I Block, Shastry Nagar, Meerut, 0121-2771160-62"
End If
If Combo2.Text = "Mohali" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Royal Mall, V Block, Vijaya Nagar, Mohali, 0134-2345360-34"
End If
If Combo2.Text = "Mumbai" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "Ground Floor,Shop no 1,Tandon Mall, Andheri Kurla Road, Andheri(East), Mumbai 3988 3988 / 6000 9000"
Text6.Text = "Ground Floor,Sumer Nagar Co-operative Society, Off S.V. Road, Borivili West, Mumbai 3988 3988 / 6000 9000"
End If
If Combo2.Text = "Mysore" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "#1/2 6th main Temple RoadV.V Mohalla, Opp to Loyal world super market, Mysore 3988-3988 / 6000 900"
End If
If Combo2.Text = "Nasik" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Opp. Wadnagre Bhavan, Thatte Nagar, College Road, Nasik"
End If
If Combo2.Text = "New Delhi" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text5.Text = "3rd Floor, Shop No. T-317, Ambience Mall, Vasant Kunj, New Delhi"
Text6.Text = "Shop No 14 & 15 Aggarwal Funcity Mall, Plot No 30, 31 CBD Shahdara, New Delhi 011-3988 3988/ 6000 9000 "
Text7.Text = "M-20, Outer Circle, Opposite Fire Station, Connaught Place, New Delhi 011-3988-3988 / 6000 9000"
End If
If Combo2.Text = "Noida" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "Ground Floor & First Floor, Spice World, Sector 25A, Noida 0120-4336498"
Text6.Text = "L-1, Centrestage Mall, Sector 18, Noida 9999391912"
End If
If Combo2.Text = "Pondicherry" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "No 4 Lal Bhahadur Shastri Street, (Bussy Street), Pondicherry, Pondicherry 4210999 / 6000900"
End If
If Combo2.Text = "Pune" Then
Text5.Visible = True
Text6.Visible = True
Text7.Visible = False
Text5.Text = "EB-FF-41 & 41D, Amonara Town Centre, Village Sadesatranali, Hadapsar, Taluka Haveli, Pune 3988-3988 / 6000-9000 "
Text6.Text = "Shop No.2, Supreme centre, ITI Road,Aundh, Pune 3988-3988 / 6000-9000 "
End If
If Combo2.Text = "Rajkot" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Shop # 6, Solitaire Building, 150 Ft. Ring Road, Nr. Big Bazaar, Rajkot 3988-3988 / 6000 900"
End If
If Combo2.Text = "Ranchi" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Shop No 101, Ground Floor, Spring City Mall, Hinoo, Ranchi"
End If
If Combo2.Text = "Surat" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Ganga House, Dumas Road, Piplod, opp Croma, Surat 3988-3988 / 6000 900"
End If
If Combo2.Text = "Thiruvananthapuram" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Heera highlife, Keston road,Kowdiar, Thiruvananthapuram, Kerala 695003"
End If
If Combo2.Text = "Tumkur" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "No. 2627/435/2536/4961, Ground Floor, Maharaja Square, Bangalore -Tumkur Road, Beside Karnataka Bank, Batawadi, Tumakuru, Karnataka 572101"
End If
If Combo2.Text = "Udaipur" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "15, 1st Floor, Lakecity Mall, Main Road, Near Ayad Bridge, Ashok Nagar, Udaipur, Rajasthan 313001"
End If
If Combo2.Text = "Vapi" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Ground Floor, NH 8, Golden Town, GIDC, Galaxy Hotel, Vapi, Gujarat 396145"
End If
If Combo2.Text = "Varanasi" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Shop No. G 1 & 2, Jhv Mall, Varanasi Cantt., Varanasi, Uttar Pradesh 221003"
End If
If Combo2.Text = "Vijayawada" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Lower Ground Floor,Next to Box Office,LEPL ICON Mall, Door No. 59A-1-4, Patamata, NH5 Frontage Road, K P Nagar, Benz Circle, Vijayawada, Andhra Pradesh 520008"
End If
If Combo2.Text = "Vishakapatnam" Then
Text5.Visible = True
Text6.Visible = False
Text7.Visible = False
Text5.Text = "Door No-47-10-23/3, Isnar Khazana Towers,Dwarkanagar, Second Lane, Visakhapatnam, Andhra Pradesh 530016"
End If
End Sub

Private Sub combo2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 59 Then
KeyAscii = 0
End If
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.MoveLast
Form2.Show
Form10.Show
Form1.Hide
End Sub

Private Sub Command2_Click()
Form12.Show
Form1.Hide
End Sub

Private Sub Command3_Click()
Label2.Visible = False
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Command1.Visible = True
Command2.Visible = False
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Combo1.Visible = True
Command3.Visible = False
Adodc1.Refresh
Adodc1.Recordset.AddNew
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
If Text1.Text = "" Or Text1.Text = "Name" Then
Text1.Text = ""
Text1.ForeColor = vbBlack
End If
End Sub

Private Sub Text2_GotFocus()
If Text2.Text = "" Or Text2.Text = "House/Flat Number" Then
Text2.Text = ""
Text2.ForeColor = vbBlack
End If
End Sub

Private Sub Text3_GotFocus()
If Text3.Text = "" Or Text3.Text = "Building Name/Street/Locality" Then
Text3.Text = ""
Text3.ForeColor = vbBlack
End If
End Sub
Private Sub Text4_GotFocus()
If Text4.Text = "" Or Text4.Text = "Mobile" Then
Text4.Text = ""
Text4.ForeColor = vbBlack
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97 And KeyAscii <= 122) Then
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Image1.Left = Image1.Left - 8
End Sub




















VERSION 5.00
Begin VB.Form reporting 
   Caption         =   "Reports"
   ClientHeight    =   5940
   ClientLeft      =   3570
   ClientTop       =   1320
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   11190
   Begin VB.Frame Frame5 
      Caption         =   "Interval report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      TabIndex        =   21
      Top             =   2640
      Width           =   4575
      Begin VB.CommandButton ipurchased 
         Caption         =   "Purchased"
         Height          =   495
         Left            =   3000
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton itrans 
         Caption         =   "Transactions"
         Height          =   495
         Left            =   1560
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton idetailed 
         Caption         =   "Detailed"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton ibar 
         Caption         =   "Bar"
         Height          =   495
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton iresto 
         Caption         =   "Restaurant"
         Height          =   495
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ibill 
         Caption         =   "Bills"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Yearly report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   4575
      Begin VB.CommandButton ybill 
         Caption         =   "Bills"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton yresto 
         Caption         =   "Restaurant"
         Height          =   495
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ybar 
         Caption         =   "Bar"
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton ydetailed 
         Caption         =   "Detailed"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton ytrans 
         Caption         =   "Transactions"
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton ypurchased 
         Caption         =   "Purchased"
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monthly report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5040
      TabIndex        =   7
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton mbill 
         Caption         =   "Bills"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton mresto 
         Caption         =   "Restaurant"
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton mbar 
         Caption         =   "Bar"
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton mdetailed 
         Caption         =   "Detailed"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton mlive 
         Caption         =   "Transactions"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton mpurchased 
         Caption         =   "Purchased"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Daily report"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton transreport 
         Caption         =   "Transactions"
         Height          =   495
         Left            =   3000
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton vpurchased 
         Caption         =   "Purchased"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton vlive 
         Caption         =   "Livestock"
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton billdetailed 
         Caption         =   "Detailed"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton billbar 
         Caption         =   "Bar"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton billresto 
         Caption         =   "Restaurant"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdbll 
         Caption         =   "Bills"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "reporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub billbar_Click()
status = "Bar"
frmbill.Show
End Sub

Private Sub billdetailed_Click()
Set conn = New connection
Dim item As New ADODB.Recordset
Dim display As Object
Dim sum, s As Long
Dim rse, rs, p As New ADODB.Recordset
s = 0
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(9, 4) = "---------------------------------"
Else
Me.Hide
company.Show
Exit Sub
End If
n = 11
Set rs = con.Execute("select * from billing where date='" + Format(Now, "d-m-yyyy") + "' order by id asc")
If Not rs.EOF Then
display.sheets("sheet1").cells(n, 3) = "Bill"
display.sheets("sheet1").cells(n, 4) = "Product"
display.sheets("sheet1").cells(n, 5) = "Unit price"
display.sheets("sheet1").cells(n, 6) = "Quantity"
display.sheets("sheet1").cells(n, 7) = "Total"
display.sheets("sheet1").cells(n, 8) = "Bloc"
rs.MoveFirst
n = n + 1
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
display.sheets("sheet1").cells(n, 3) = rs.Fields("bill_id")
display.sheets("sheet1").cells(n, 8) = rs.Fields("bloc")
item.MoveFirst
sum = 0
While Not item.EOF
Set p = con.Execute("select * from products_tb where prod_id='" + item.Fields("prod_id") + "' ")
display.sheets("sheet1").cells(n, 4) = p.Fields("name")
display.sheets("sheet1").cells(n, 5) = item.Fields("unit_price")
display.sheets("sheet1").cells(n, 6) = item.Fields("out_quantity")
display.sheets("sheet1").cells(n, 7) = item.Fields("total")
sum = sum + Val(item.Fields("total"))
item.MoveNext
n = n + 1
Wend
display.sheets("sheet1").cells(n, 6) = "Total of bill:"
display.sheets("sheet1").cells(n, 7) = sum
t = t + sum
s = s + t
n = n + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(n, 6) = "Total:"
display.sheets("sheet1").cells(n, 7) = t
Else
m = MsgBox("No data found", vbCritical + vbOKOnly, "Warning")
frmbill.Hide
Exit Sub
End If
End Sub

Private Sub billresto_Click()
status = "Resto"
frmbill.Show
End Sub

Private Sub cmdbll_Click()
status = "all"
frmbill.Show
End Sub


Private Sub Command1_Click()

End Sub

Private Sub ypurchased_Click()
Set conn = New connection
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
d = Format(Now, "d-m-yyyy")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Quantity"
display.sheets("sheet1").cells(9, 7) = "Total"
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from products_update where prod_id='" + rs.Fields("prod_id") + "' and date like'%" + Format(Now, "yyyy") + "%'")
If Not item.EOF Then
item.MoveFirst
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("new_quantity"))
sum = sum + (Val(item.Fields("unit_price")) * Val(item.Fields("new_quantity")))
item.MoveNext
Wend
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = rs.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = q
display.sheets("sheet1").cells(r, 7) = sum
n = n + 1
r = r + 1
t = t + sum
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 6) = "Total:"
display.sheets("sheet1").cells(r, 7) = t
End If

End Sub

Private Sub ytrans_Click()
Set conn = New connection
d = Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Bill NO"
display.sheets("sheet1").cells(9, 7) = "Quantity in store"
display.sheets("sheet1").cells(9, 8) = "Quantity sold"
display.sheets("sheet1").cells(9, 9) = "Rest Quantity"
rs.MoveFirst
n = 1
r = 11
sum = 0
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where (prod_id='" + rs.Fields("prod_id") + "' and date like'%" + Format(Now, "yyyy") + "%' and kind='Out')   order by bill_id")
If Not item.EOF Then
item.MoveFirst
While Not item.EOF
last = item.Fields("rest_quantity")
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = item.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = item.Fields("Bill_id")
display.sheets("sheet1").cells(r, 7) = item.Fields("actual_quantity")
display.sheets("sheet1").cells(r, 8) = item.Fields("out_quantity")
display.sheets("sheet1").cells(r, 9) = item.Fields("rest_quantity")
n = n + 1
r = r + 1
item.MoveNext
Wend
End If
rs.MoveNext
Wend
End If

End Sub

Private Sub ydetailed_Click()
Set conn = New connection
Dim item As New ADODB.Recordset
Dim display As Object
Dim sum, s As Long
Dim rse, rs, p As New ADODB.Recordset
s = 0
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(9, 4) = "---------------------------------"
Else
Me.Hide
company.Show
Exit Sub
End If
n = 11
Set rs = con.Execute("select * from billing where date like'%" + Format(Now, "yyyy") + "%' order by id asc")
If Not rs.EOF Then
display.sheets("sheet1").cells(n, 3) = "Bill"
display.sheets("sheet1").cells(n, 4) = "Product"
display.sheets("sheet1").cells(n, 5) = "Unit price"
display.sheets("sheet1").cells(n, 6) = "Quantity"
display.sheets("sheet1").cells(n, 7) = "Total"
display.sheets("sheet1").cells(n, 8) = "Bloc"
rs.MoveFirst
n = n + 1
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
display.sheets("sheet1").cells(n, 3) = rs.Fields("bill_id")
display.sheets("sheet1").cells(n, 8) = rs.Fields("bloc")
item.MoveFirst
sum = 0
While Not item.EOF
Set p = con.Execute("select * from products_tb where prod_id='" + item.Fields("prod_id") + "' ")
display.sheets("sheet1").cells(n, 4) = p.Fields("name")
display.sheets("sheet1").cells(n, 5) = item.Fields("unit_price")
display.sheets("sheet1").cells(n, 6) = item.Fields("out_quantity")
display.sheets("sheet1").cells(n, 7) = item.Fields("total")
sum = sum + Val(item.Fields("total"))
item.MoveNext
n = n + 1
Wend
display.sheets("sheet1").cells(n, 6) = "Total of bill:"
display.sheets("sheet1").cells(n, 7) = sum
t = t + sum
s = s + t
n = n + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(n, 6) = "Total:"
display.sheets("sheet1").cells(n, 7) = t
Else
m = MsgBox("No data found", vbCritical + vbOKOnly, "Warning")
frmbill.Hide
Exit Sub
End If

End Sub

Private Sub ybar_Click()
Set conn = New connection
status = "year_all"
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(8, 4) = "Bill report"
display.sheets("sheet1").cells(9, 4) = "---------------------------------"
display.sheets("sheet1").cells(10, 2) = "No"
display.sheets("sheet1").cells(10, 4) = "Bill No:"
display.sheets("sheet1").cells(10, 6) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Bar' and date like '%" + Format(Now, "yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
init = item.Fields("rest_quantity")
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("out_quantity"))
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 5) = "Total:"
display.sheets("sheet1").cells(r, 6) = t
End If

End Sub

Private Sub yresto_Click()
Set conn = New connection
status = "year_all"
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(8, 4) = "Bill report"
display.sheets("sheet1").cells(9, 4) = "---------------------------------"
display.sheets("sheet1").cells(10, 2) = "No"
display.sheets("sheet1").cells(10, 4) = "Bill No:"
display.sheets("sheet1").cells(10, 6) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto' and date like '%" + Format(Now, "yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
init = item.Fields("rest_quantity")
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("out_quantity"))
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 5) = "Total:"
display.sheets("sheet1").cells(r, 6) = t
End If

End Sub

Private Sub ybill_Click()
Set conn = New connection
status = "year_all"
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(8, 4) = "Bill report"
display.sheets("sheet1").cells(9, 4) = "---------------------------------"
display.sheets("sheet1").cells(10, 2) = "No"
display.sheets("sheet1").cells(10, 4) = "Bill No:"
display.sheets("sheet1").cells(10, 6) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where date like '%" + Format(Now, "yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
init = item.Fields("rest_quantity")
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("out_quantity"))
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 5) = "Total:"
display.sheets("sheet1").cells(r, 6) = t
End If
End Sub

Private Sub ibill_Click()
status = "All"
m = MsgBox("Choose start date", vbInformation, "Information")
frmcal.Show
End Sub

Private Sub iresto_Click()
status = "Resto"
m = MsgBox("Choose start date", vbInformation, "Information")
frmcal.Show
End Sub

Private Sub ibar_Click()
status = "Bar"
m = MsgBox("Choose start date", vbInformation, "Information")
frmcal.Show
End Sub

Private Sub idetailed_Click()
status = "detailed"
frmcal.Show
End Sub

Private Sub itrans_Click()
status = "transaction"
frmcal.Show
End Sub

Private Sub ipurchased_Click()
status = "purchased"
frmcal.Show
End Sub

Private Sub Form_Load()
'On Error GoTo connect
'connect:
'Call ShellAndWait("c:\wamp64\wampmanager.exe", vbNormalFocus)
End Sub

Private Sub mbar_Click()
status = "bar_month"
frmbill.Show
End Sub

Private Sub mbill_Click()
status = "month_all"
frmbill.Show
End Sub

Private Sub mdetailed_Click()
status = "month_detailed"
Set conn = New connection
Dim item As New ADODB.Recordset
Dim display As Object
Dim sum, s As Long
Dim rse, rs, p As New ADODB.Recordset
s = 0
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(9, 4) = "---------------------------------"
Else
Me.Hide
company.Show
Exit Sub
End If
n = 11
Set rs = con.Execute("select * from billing where date like'%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
display.sheets("sheet1").cells(n, 3) = "Bill"
display.sheets("sheet1").cells(n, 4) = "Product"
display.sheets("sheet1").cells(n, 5) = "Unit price"
display.sheets("sheet1").cells(n, 6) = "Quantity"
display.sheets("sheet1").cells(n, 7) = "Total"
display.sheets("sheet1").cells(n, 8) = "Bloc"
rs.MoveFirst
n = n + 1
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
display.sheets("sheet1").cells(n, 3) = rs.Fields("bill_id")
display.sheets("sheet1").cells(n, 8) = rs.Fields("bloc")
item.MoveFirst
sum = 0
While Not item.EOF
Set p = con.Execute("select * from products_tb where prod_id='" + item.Fields("prod_id") + "' ")
display.sheets("sheet1").cells(n, 4) = p.Fields("name")
display.sheets("sheet1").cells(n, 5) = item.Fields("unit_price")
display.sheets("sheet1").cells(n, 6) = item.Fields("out_quantity")
display.sheets("sheet1").cells(n, 7) = item.Fields("total")
sum = sum + Val(item.Fields("total"))
item.MoveNext
n = n + 1
Wend
display.sheets("sheet1").cells(n, 6) = "Total of bill:"
display.sheets("sheet1").cells(n, 7) = sum
t = t + sum
s = s + t
n = n + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(n, 6) = "Total:"
display.sheets("sheet1").cells(n, 7) = t
Else
m = MsgBox("No data found", vbCritical + vbOKOnly, "Warning")
frmbill.Hide
Exit Sub
End If

End Sub

Private Sub mlive_Click()
Set conn = New connection
d = Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Bill NO"
display.sheets("sheet1").cells(9, 7) = "Quantity in store"
display.sheets("sheet1").cells(9, 8) = "Quantity sold"
display.sheets("sheet1").cells(9, 9) = "Rest Quantity"
rs.MoveFirst
n = 1
r = 11
sum = 0
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where (prod_id='" + rs.Fields("prod_id") + "' and date like'%" + Format(Now, "m-yyyy") + "%' and kind='Out')   order by bill_id")
If Not item.EOF Then
item.MoveFirst
While Not item.EOF
last = item.Fields("rest_quantity")
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = item.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = item.Fields("Bill_id")
display.sheets("sheet1").cells(r, 7) = item.Fields("actual_quantity")
display.sheets("sheet1").cells(r, 8) = item.Fields("out_quantity")
display.sheets("sheet1").cells(r, 9) = item.Fields("rest_quantity")
n = n + 1
r = r + 1
item.MoveNext
Wend
End If
rs.MoveNext
Wend
End If
End Sub

Private Sub mpurchased_Click()
Set conn = New connection
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
d = Format(Now, "d-m-yyyy")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Quantity"
display.sheets("sheet1").cells(9, 7) = "Total"
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from products_update where prod_id='" + rs.Fields("prod_id") + "' and date like'%" + Format(Now, "m-yyyy") + "%'")
If Not item.EOF Then
item.MoveFirst
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("new_quantity"))
sum = sum + (Val(item.Fields("unit_price")) * Val(item.Fields("new_quantity")))
item.MoveNext
Wend
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = rs.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = q
display.sheets("sheet1").cells(r, 7) = sum
n = n + 1
r = r + 1
t = t + sum
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 6) = "Total:"
display.sheets("sheet1").cells(r, 7) = t
End If

End Sub

Private Sub mresto_Click()
status = "resto_month"
frmbill.Show
End Sub

Private Sub transreport_Click()
Set conn = New connection
d = Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Bill NO"
display.sheets("sheet1").cells(9, 7) = "Quantity in store"
display.sheets("sheet1").cells(9, 8) = "Quantity sold"
display.sheets("sheet1").cells(9, 9) = "Rest Quantity"
rs.MoveFirst
n = 1
r = 11
sum = 0
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where (prod_id='" + rs.Fields("prod_id") + "' and date='" + Format(Now, "d-m-yyyy") + "' and kind='Out')   order by bill_id")
If Not item.EOF Then
item.MoveFirst
While Not item.EOF
last = item.Fields("rest_quantity")
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = item.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = item.Fields("Bill_id")
display.sheets("sheet1").cells(r, 7) = item.Fields("actual_quantity")
display.sheets("sheet1").cells(r, 8) = item.Fields("out_quantity")
display.sheets("sheet1").cells(r, 9) = item.Fields("rest_quantity")
n = n + 1
r = r + 1
item.MoveNext
Wend
End If
rs.MoveNext
Wend
End If
End Sub

Private Sub vlive_Click()
Set conn = New connection
d = Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Quantity in store"
rs.MoveFirst
n = 1
r = 10
sum = 0
t = 0
While Not rs.EOF
Set item = con.Execute("select * from livestock where (prod_id='" + rs.Fields("prod_id") + "' and rest_quantity!=0)   order by transaction_id desc limit 1")
If Not item.EOF Then
last = item.Fields("rest_quantity")
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = item.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = item.Fields("rest_quantity")
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
End If

End Sub

Private Sub vpurchased_Click()
Set conn = New connection
Set rs = con.Execute("select * from products_tb order by name asc")
Set rse = con.Execute("select * from identification")
d = Format(Now, "d-m-yyyy")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.workbooks.Add
display.sheets("sheet1").cells(1, 1) = rse.Fields("name")
display.sheets("sheet1").cells(2, 1) = "REPUBLIC OF RWANDA"
display.sheets("sheet1").cells(3, 1) = "KIGALI CITY"
display.sheets("sheet1").cells(3, 2) = rse.Fields("district") & " District"
display.sheets("sheet1").cells(4, 1) = rse.Fields("sector") & " Sector"
display.sheets("sheet1").cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.sheets("sheet1").cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.sheets("sheet1").cells(6, 1) = "Email: " & rse.Fields("email")
'display.sheets("sheet1").cells(7, 4) = "---------------------------------"
display.sheets("sheet1").cells(9, 2) = "No"
display.sheets("sheet1").cells(9, 3) = "Product ID"
display.sheets("sheet1").cells(9, 4) = "Name"
display.sheets("sheet1").cells(9, 5) = "Category"
display.sheets("sheet1").cells(9, 6) = "Quantity"
display.sheets("sheet1").cells(9, 7) = "Total"
rs.MoveFirst
n = 1
r = 10
t = 0
While Not rs.EOF
Set item = con.Execute("select * from products_update where prod_id='" + rs.Fields("prod_id") + "' and date='" + d + "'")
If Not item.EOF Then
item.MoveFirst
q = 0
sum = 0
While Not item.EOF
q = q + Val(item.Fields("new_quantity"))
sum = sum + (Val(item.Fields("unit_price")) * Val(item.Fields("new_quantity")))
item.MoveNext
Wend
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 3) = rs.Fields("prod_id")
display.sheets("sheet1").cells(r, 4) = rs.Fields("name")
display.sheets("sheet1").cells(r, 5) = rs.Fields("category")
display.sheets("sheet1").cells(r, 6) = q
display.sheets("sheet1").cells(r, 7) = sum
n = n + 1
r = r + 1
t = t + sum
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 6) = "Total:"
display.sheets("sheet1").cells(r, 7) = t
End If
End Sub



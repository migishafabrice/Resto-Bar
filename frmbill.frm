VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmbill 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12375
   Begin VB.CommandButton reportexcel 
      Caption         =   "Report to excel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton back 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxbill 
      Height          =   7005
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   10
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image pci 
      Height          =   1695
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label lblcompany 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label lblrep 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Republic of Rwanda"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbldis 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label lblse 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sector"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lbltel 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label dte 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   1
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   2520
      Width           =   1170
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
Call billings
End Sub

Private Sub flxbill_Click()
back.Visible = True
If flxbill.Col = 1 Then
bill = flxbill.Text
Set rs = con.Execute("select * from billing where bill_id='" + bill + "'")
If Not rs.EOF Then
bloc = rs.Fields("bloc")
Label7.Caption = "Bill: " & bill & vbCr & "Bloc:" & bloc & vbCr & " Date:" & rs.Fields("date")
Set rs = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
rs.MoveLast
b = rs.RecordCount
flxbill.Clear
With flxbill
.Cols = 5
.Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2000
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Product"
.TextMatrix(0, 2) = "Quantity"
.TextMatrix(0, 3) = "Unit price"
.TextMatrix(0, 4) = "Total"
sum = 0
rs.MoveFirst
n = 1
While Not rs.EOF
Set item = con.Execute("select * from products_tb where  prod_id='" + rs.Fields("prod_id") + "'")
If Not item.EOF Then
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = item.Fields("name")
.TextMatrix(n, 2) = rs.Fields("out_quantity")
.TextMatrix(n, 3) = rs.Fields("unit_price")
.TextMatrix(n, 4) = rs.Fields("total")
sum = sum + rs.Fields("total")
End If
n = n + 1
rs.MoveNext
Wend
.TextMatrix(n, 3) = "Total:"
.TextMatrix(n, 4) = Str(sum)
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If
End If
End Sub

Private Sub Form_Load()
Call billings
End Sub

Public Sub billings()
back.Visible = False
Set conn = New connection
Me.Height = Screen.Height - 1000
Dim rse, rs As New ADODB.Recordset
n = 1
If status = "all" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where date='" + Format(Now, "d-m-yyyy") + "' order by id asc")
If Not rs.EOF Then
Label7.Caption = "Resto-Bar Bills: " & Format(Now, "dddd  d-m-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If
End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "Resto" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto' and date='" + Format(Now, "d-m-yyyy") + "' order by id asc")
If Not rs.EOF Then
Label7.Caption = "Resto Bills: " & Format(Now, "dddd  d-m-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If

End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "Bar" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where boc='Bar' and date='" + Format(Now, "d-m-yyyy") + "' order by id asc")
If Not rs.EOF Then
Label7.Caption = "Bar Bills: " & Format(Now, "dddd  d-m-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If

End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If status = "month_all" Then

Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where  date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")

If Not rs.EOF Then
Label7.Caption = "Resto-Bar Bills: " & Format(Now, "mmmm-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If
End If
''''''''''''''''''''''''''''''''''''''''''''''
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "resto_month" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto' and date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
Label7.Caption = "Resto Bills: " & Format(Now, "mmmm-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If

End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "bar_month" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
lblcompany.Caption = rse.Fields("name")
lbldis.Caption = rse.Fields("district") & " District"
lblse.Caption = rse.Fields("sector") & " Sector"
lbltel.Caption = "Telephone: " & rse.Fields("telephone")
lblmail.Caption = "Email: " & rse.Fields("email")
frmbill.Caption = rse.Fields("name")
pic = rse.Fields("logo")
pci.Picture = LoadPicture(App.Path & pic)
Else
Me.Hide
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Bar' and date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
Label7.Caption = "Bar Bills: " & Format(Now, "mmmm-yyyy")
flxbill.Visible = True
rs.MoveLast
b = rs.RecordCount
With flxbill
  .Rows = b + 2
For i = 0 To 3
.ColWidth(i) = 2500
Next
.TextMatrix(0, 0) = "NO"
.TextMatrix(0, 1) = "Bill No"
.TextMatrix(0, 2) = "Date"
.TextMatrix(0, 3) = "Total"
 
t = 0
rs.MoveFirst
While Not rs.EOF
Set item = con.Execute("select * from livestock where bill_id='" + rs.Fields("bill_id") + "' and kind='Out'")
If Not item.EOF Then
item.MoveFirst
sum = 0
While Not item.EOF
sum = sum + Val(item.Fields("total"))
item.MoveNext
Wend
t = t + sum
.TextMatrix(n, 0) = n
.TextMatrix(n, 1) = rs.Fields("bill_id")
.TextMatrix(n, 2) = rs.Fields("date")
.TextMatrix(n, 3) = sum
n = n + 1
End If
rs.MoveNext
Wend
.TextMatrix(n, 2) = "Total:"
.TextMatrix(n, 3) = t
Exit Sub
End With
Else
m = MsgBox("No data found", vbInformation, "warning")
End If

End If
''''''''''''''''''''''''''''''''''''''''''''''
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

End Sub

Private Sub reportexcel_Click()
Set conn = New connection
If status = "month_all" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If
'<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>
If status = "resto_month" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto' and date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "bar_month" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Bar' and date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If
'<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "all" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where date like '%" + Format(Now, "d-m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If
'<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>
If status = "Resto" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto' and date like '%" + Format(Now, "m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>
If status = "Bar" Then
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
Set display = CreateObject("excel.application")
display.Visible = True
display.Workbooks.Add
display.Sheets("sheet1").Cells(1, 1) = rse.Fields("name")
display.Sheets("sheet1").Cells(2, 1) = "REPUBLIC OF RWANDA"
display.Sheets("sheet1").Cells(3, 1) = "KIGALI CITY"
display.Sheets("sheet1").Cells(3, 2) = rse.Fields("district") & " District"
display.Sheets("sheet1").Cells(4, 1) = rse.Fields("sector") & " Sector"
display.Sheets("sheet1").Cells(5, 1) = "Telephone: " & rse.Fields("telephone")
display.Sheets("sheet1").Cells(6, 1) = "Email: " & rse.Fields("email")
display.Sheets("sheet1").Cells(1, 8) = "Date: " & Format(Now, "dddd  d-m-yyyy")
display.Sheets("sheet1").Cells(8, 4) = "Bill report"
display.Sheets("sheet1").Cells(9, 4) = "---------------------------------"
display.Sheets("sheet1").Cells(10, 2) = "No"
display.Sheets("sheet1").Cells(10, 4) = "Bill No:"
display.Sheets("sheet1").Cells(10, 6) = "Date"
display.Sheets("sheet1").Cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Bar' and date like '%" + Format(Now, "d-m-yyyy") + "%' order by id asc")
If Not rs.EOF Then
rs.MoveFirst
n = 1
r = 11
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
display.Sheets("sheet1").Cells(r, 2) = n
display.Sheets("sheet1").Cells(r, 4) = rs.Fields("bill_id")
display.Sheets("sheet1").Cells(r, 6) = rs.Fields("date")
display.Sheets("sheet1").Cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.Sheets("sheet1").Cells(r, 7) = "Total:"
display.Sheets("sheet1").Cells(r, 8) = t
End If
End If

End Sub

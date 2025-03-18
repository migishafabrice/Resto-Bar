VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmcal 
   Caption         =   "Calendar"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin MSACAL.Calendar calview 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _Version        =   524288
      _ExtentX        =   10821
      _ExtentY        =   7011
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   9
      Day             =   9
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmcal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calview_Click()
If startdate = "" Then
startdate = Format(calview.Value, "d-m-yyyy")
Me.Hide
m = MsgBox("Start date:" & startdate & vbCr & "Choose end date?", vbInformation, "Start date")
frmcal.Show
Exit Sub
End If
If enddate = "" Then
enddate = Format(calview.Value, "d-m-yyyy")
Me.Hide
m = MsgBox("End date:" & enddate & vbCr & "Generate  report?", vbYesNo + vbInformation, "End date")
If m = vbYes And status = "All" Then
Set conn = New connection
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
display.sheets("sheet1").cells(10, 4) = "Bill No"
display.sheets("sheet1").cells(10, 6) = "Date"
display.sheets("sheet1").cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') order by id asc")
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
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = rs.Fields("date")
display.sheets("sheet1").cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 7) = "Total:"
display.sheets("sheet1").cells(r, 8) = t
End If
Exit Sub
End If
If m = vbYes And status = "Resto" Then
Set conn = New connection
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
display.sheets("sheet1").cells(10, 4) = "Bill No"
display.sheets("sheet1").cells(10, 6) = "Date"
display.sheets("sheet1").cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Resto'  and str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') order by id asc")
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
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = rs.Fields("date")
display.sheets("sheet1").cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 7) = "Total:"
display.sheets("sheet1").cells(r, 8) = t
End If
Exit Sub
End If
If m = vbYes And status = "Bar" Then
Set conn = New connection
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
display.sheets("sheet1").cells(10, 4) = "Bill No"
display.sheets("sheet1").cells(10, 6) = "Date"
display.sheets("sheet1").cells(10, 8) = "Total"
Else
company.Show
Exit Sub
End If
Set rs = con.Execute("select * from billing where bloc='Bar'  and str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') order by id asc")
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
display.sheets("sheet1").cells(r, 2) = n
display.sheets("sheet1").cells(r, 4) = rs.Fields("bill_id")
display.sheets("sheet1").cells(r, 6) = rs.Fields("date")
display.sheets("sheet1").cells(r, 8) = sum
t = t + sum
n = n + 1
r = r + 1
End If
rs.MoveNext
Wend
display.sheets("sheet1").cells(r, 7) = "Total:"
display.sheets("sheet1").cells(r, 8) = t
End If
Exit Sub
End If
If m = vbYes And status = "detailed" Then
Set conn = New connection
'Dim item As New ADODB.Recordset
'Dim display As Object
'Dim sum, s As Long
'Dim rse, rs, p As New ADODB.Recordset
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
Set rs = con.Execute("select * from billing where str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') order by id asc")
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
Exit Sub
End If
Exit Sub
End If
If m = vbYes And status = "purchased" Then
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
Set item = con.Execute("select * from products_update where prod_id='" + rs.Fields("prod_id") + "' and str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') ")
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
Exit Sub
End If


If m = vbYes And status = "transaction" Then

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
Set item = con.Execute("select * from livestock where (prod_id='" + rs.Fields("prod_id") + "' and str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y') and kind='Out')   order by bill_id")
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
Exit Sub
End If
End If
If startdate <> "" And enddate <> "" Then
m = MsgBox("Reset date?", vbQuestion + vbYesNo, "Information")
If m = vbYes Then
startdate = ""
enddate = ""
frmcal.Show
End If
End If
End Sub


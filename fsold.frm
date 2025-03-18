VERSION 5.00
Begin VB.Form fsold 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7680
   ClientLeft      =   5865
   ClientTop       =   1920
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12165
   Begin VB.CommandButton rprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ProdID"
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
      Left            =   1005
      TabIndex        =   15
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "U.Price"
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
      Left            =   9765
      TabIndex        =   14
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
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
      Left            =   11115
      TabIndex        =   13
      Top             =   3480
      Width           =   570
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "R.Quantity"
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
      Left            =   8160
      TabIndex        =   12
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "S.Quantity"
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
      Left            =   6435
      TabIndex        =   11
      Top             =   3480
      Width           =   1185
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
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
      Left            =   4965
      TabIndex        =   10
      Top             =   3480
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Industry"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
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
      Left            =   2355
      TabIndex        =   8
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "No"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   1  'Blackness
      X1              =   4920
      X2              =   7575
      Y1              =   3240
      Y2              =   3255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Historic of sells period:"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   2760
      Width           =   2610
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
      Left            =   8160
      TabIndex        =   5
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Telephone:0788859419"
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
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nyambirambo Sector"
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
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nyarugenge District"
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
      Top             =   1320
      Width           =   1965
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIZAMA RESTO BAR"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   2880
   End
End
Attribute VB_Name = "fsold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim item As New ADODB.Recordset
Dim no, prod, nme, industry, category, actual, out, unit, total As Object
Dim sum As Long
Set conn = New connection
If report.soldyes.Value = True And startdate <> "" Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb order by name asc")
If Not rs.EOF Then
n = 1
nber = 20
sum = 0
a = 4240
t = 0
rs.MoveFirst
While Not rs.EOF
If enddate <> "" Then
Label7.Caption = "Historic of sells period: " & startdate & " to " & enddate
Set item = con.Execute("select * from livestock where prod_id='" + rs.Fields("prod_id") + "' and kind='Out' and (str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y'))")
Else
Label7.Caption = "Historic of sells: " & startdate
Set item = con.Execute("select * from livestock where prod_id='" + rs.Fields("prod_id") + "' and kind='Out' and (str_to_date(date,'%d-%m-%Y') =str_to_date('" + startdate + "','%d-%m-%Y'))")
End If
If Not item.EOF Then
item.MoveFirst
q = 0
b = 240
While Not item.EOF
q = q + Val(item.Fields("out_quantity"))
sum = sum + Val(item.Fields("total"))
t = t + sum
item.MoveNext
Wend
item.MoveLast
last = item.Fields("rest_quantity")
Set no = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
no.Caption = n
no.Left = Label8.Left
no.Top = a
no.Width = 300
no.Visible = True
no.BackColor = &HFFFFFF
no.FontSize = 10
no.FontName = "MS SERIF"
b = b + 1200
Set prod = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
prod.Caption = rs.Fields("prod_id")
prod.Left = Label17.Left
prod.Top = a
prod.Width = 1200
prod.Visible = True
prod.BackColor = &HFFFFFF
prod.FontSize = 10
prod.FontName = "MS SERIF"
Set nme = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
nme.Caption = rs.Fields("name")
nme.Left = Label9.Left
nme.Top = a
nme.Width = 1200
nme.Visible = True
nme.BackColor = &HFFFFFF
nme.FontSize = 10
prod.FontName = "MS SERIF"
Set industry = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
industry.Caption = rs.Fields("industry")
industry.Left = Label10.Left
industry.Top = a
industry.Width = 1200
industry.BackColor = &HFFFFFF
industry.Visible = True
industry.FontSize = 10
industry.FontName = "MS SERIF"
Set category = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
category.Caption = rs.Fields("category")
category.Left = Label11.Left
category.Top = a
category.Width = 1200
category.BackColor = &HFFFFFF
category.Visible = True
category.FontSize = 10
category.FontName = "MS SERIF"
Set actual = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
actual.Caption = last
actual.Left = Label14.Left
actual.Width = 1200
actual.Visible = True
actual.FontSize = 10
actual.BackColor = &HFFFFFF
actual.FontName = "MS SERIF"
actual.Top = a
b = b + 1200
Set out = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
out.Caption = q
out.Left = Label13.Left
out.Top = a
out.Width = 1200
out.Visible = True
out.FontSize = 10
out.BackColor = &HFFFFFF
out.FontName = "MS SERIF"
Set unit = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
unit.Caption = item.Fields("unit_price")
unit.Left = Label16.Left
unit.Top = a
unit.Width = 1200
unit.Visible = True
unit.FontSize = 10
unit.BackColor = &HFFFFFF
unit.FontName = "MS SERIF"
 Set total = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
total.Caption = sum
total.Left = Label15.Left
total.Top = a
total.Width = 1200
total.Visible = True
total.BackColor = &HFFFFFF
total.FontSize = 10
total.FontName = "MS SERIF"
n = n + 1
a = a + 500
End If
rs.MoveNext
Wend
End If
End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If report.scat.Text <> "None" And startdate <> "" Then
dte.Caption = "Date:" & Format(Now, "dddd  d-m-yyyy")
Set rs = con.Execute("select * from products_tb where category='" + report.scat.Text + "'order by name asc")
If Not rs.EOF Then
n = 1
nber = 20
sum = 0
a = 4240
t = 0
rs.MoveFirst
While Not rs.EOF
If enddate <> "" Then
Label7.Caption = "Historic of sells period: " & startdate & " to " & enddate
Set item = con.Execute("select * from livestock where prod_id='" + rs.Fields("prod_id") + "' and kind='Out' and (str_to_date(date,'%d-%m-%Y') between str_to_date('" + startdate + "','%d-%m-%Y') and str_to_date('" + enddate + "','%d-%m-%Y'))")
Else
Label7.Caption = "Historic of sells: " & startdate
Set item = con.Execute("select * from livestock where prod_id='" + rs.Fields("prod_id") + "' and kind='Out' and (str_to_date(date,'%d-%m-%Y') =str_to_date('" + startdate + "','%d-%m-%Y'))")
End If
If Not item.EOF Then
item.MoveFirst
q = 0
b = 240
While Not item.EOF
q = q + Val(item.Fields("out_quantity"))
sum = sum + Val(item.Fields("total"))
t = t + sum
Set no = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
no.Caption = n
no.Left = Label8.Left
no.Top = a
no.Width = 300
no.Visible = True
no.BackColor = &HFFFFFF
no.FontSize = 10
no.FontName = "MS SERIF"
b = b + 1200
Set prod = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
prod.Caption = rs.Fields("prod_id")
prod.Left = Label17.Left
prod.Top = a
prod.Width = 1200
prod.Visible = True
prod.BackColor = &HFFFFFF
prod.FontSize = 10
prod.FontName = "MS SERIF"
Set nme = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
nme.Caption = rs.Fields("name")
nme.Left = Label9.Left
nme.Top = a
nme.Width = 1200
nme.Visible = True
nme.BackColor = &HFFFFFF
nme.FontSize = 10
prod.FontName = "MS SERIF"
Set industry = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
industry.Caption = rs.Fields("industry")
industry.Left = Label10.Left
industry.Top = a
industry.Width = 1200
industry.BackColor = &HFFFFFF
industry.Visible = True
industry.FontSize = 10
industry.FontName = "MS SERIF"
Set category = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
category.Caption = rs.Fields("category")
category.Left = Label11.Left
category.Top = a
category.Width = 1200
category.BackColor = &HFFFFFF
category.Visible = True
category.FontSize = 10
category.FontName = "MS SERIF"
Set actual = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
actual.Caption = item.Fields("rest_quantity")
actual.Left = Label14.Left
actual.Width = 1200
actual.Visible = True
actual.FontSize = 10
actual.BackColor = &HFFFFFF
actual.FontName = "MS SERIF"
actual.Top = a
b = b + 1200
Set out = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
out.Caption = q
out.Left = Label13.Left
out.Top = a
out.Width = 1200
out.Visible = True
out.FontSize = 10
out.BackColor = &HFFFFFF
out.FontName = "MS SERIF"
Set unit = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
unit.Caption = item.Fields("unit_price")
unit.Left = Label16.Left
unit.Top = a
unit.Width = 1200
unit.Visible = True
unit.FontSize = 10
unit.BackColor = &HFFFFFF
unit.FontName = "MS SERIF"
Set total = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
total.Caption = sum
total.Left = Label15.Left
total.Top = a
total.Width = 1200
total.Visible = True
total.BackColor = &HFFFFFF
total.FontSize = 10
total.FontName = "MS SERIF"
n = n + 1
a = a + 500
item.MoveNext
Wend
End If
rs.MoveNext
Wend
Set unit = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
unit.Caption = "Total:"
unit.Left = Label16.Left
unit.Top = a
unit.Width = 1200
unit.Visible = True
unit.FontSize = 10
unit.BackColor = &HFFFFFF
unit.FontName = "MS SERIF"
unit.FontBold = True
Set total = Me.Controls.Add("VB.Label", "label" & nber)
nber = nber + 1
total.Caption = t
total.Left = Label15.Left
total.Top = a
total.Width = 1200
total.Visible = True
total.BackColor = &HFFFFFF
total.FontSize = 10
total.FontName = "MS SERIF"
total.FontBold = True
End If
End If

End Sub
Private Sub rprint_Click()
rprint.Visible = False
Me.PrintForm
End Sub

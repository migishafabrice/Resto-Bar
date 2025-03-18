VERSION 5.00
Begin VB.Form printing 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Price/unit"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Product"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label idlb 
      AutoSize        =   -1  'True
      Caption         =   "......"
      Height          =   240
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'If user = "" Or pass = "" Then
'm = MsgBox("Username or passpword incorrect", vbCritical, "Warning")
'Me.Hide
'login.Show
'End If
Dim mycontrol, price, quantity, total, all As Object
Dim rs, rse As New Recordset
Dim sum As Double
Set rse = con.Execute("select * from identification")
If Not rse.EOF Then
idlb.Caption = rse.Fields("name")
End If
Me.Caption = "Invoice:" & bill
a = 1440
Set conn = New connection
Set rs = con.Execute("select * from products_tb,livestock where products_tb.prod_id=livestock.prod_id and livestock.bill_id='" + bill + "'")
If Not rs.EOF Then
rs.MoveFirst
n = 6
sum = 0
While Not rs.EOF
Set mycontrol = Me.Controls.Add("VB.Label", "label" & n)
n = n + 1
Set price = Me.Controls.Add("VB.Label", "label" & n)
n = n + 1
Set quantity = Me.Controls.Add("VB.Label", "label" & n)
n = n + 1
Set total = Me.Controls.Add("VB.Label", "label1" & n)
mycontrol.Caption = rs.Fields("name")
price.Caption = rs.Fields("unit_price")
quantity.Caption = rs.Fields("out_quantity")
total.Caption = rs.Fields("total")
mycontrol.Left = 120
price.Left = 1560
quantity.Left = 2760
total.Left = 3720
mycontrol.Top = a
mycontrol.Width = 1300
mycontrol.Visible = True
price.Top = a
price.Width = 1000
price.Visible = True
quantity.Top = a
quantity.Width = 1000
quantity.Visible = True
total.Top = a
total.Width = 1500
total.Visible = True
a = a + 450
sum = sum + rs.Fields("total")
rs.MoveNext
Wend
Set all = Me.Controls.Add("VB.Label", "label" & n)
all.Caption = "---------"
all.Left = 3720
all.Top = a
all.Width = 1500
all.Visible = True
a = a + 450
n = n + 1
Set all = Me.Controls.Add("VB.Label", "label" & n)
all.Caption = sum
all.Left = 3720
all.Top = a
all.Width = 1500
all.Visible = True
Me.PrintForm
End If
End Sub

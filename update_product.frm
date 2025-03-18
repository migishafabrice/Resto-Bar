VERSION 5.00
Begin VB.Form update_product 
   Caption         =   "New quantity"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   10005
   Begin VB.Frame Frame1 
      Caption         =   "Update product  in stock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   7335
      Begin VB.TextBox expire 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   31
         Top             =   6360
         Width           =   4575
      End
      Begin VB.TextBox nquantity 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   30
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox man 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   28
         Top             =   4800
         Width           =   4575
      End
      Begin VB.TextBox buy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2280
         TabIndex        =   26
         Top             =   4320
         Width           =   4575
      End
      Begin VB.TextBox billnber 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   22
         Top             =   5280
         Width           =   4575
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtind 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label aquantity 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   29
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label lastprice 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   27
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Manufactured date:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Bill number:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Industry:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Last update:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lastdate 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Current quantity:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Last price/unit:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "New quantity:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Buy unit price:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Total price new quantity:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label totalprice 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   5760
         Width           =   4575
      End
      Begin VB.Label Label12 
         Caption         =   "Total quantity:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label tquantity 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   10
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label Label14 
         Caption         =   "Expiration date:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   6360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search product"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   7335
      Begin VB.ComboBox cboprod 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Enter or select  name/code:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.CommandButton updateprod 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Expiration date:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   5760
      Width           =   1695
   End
End
Attribute VB_Name = "update_product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, s As New ADODB.Recordset
Dim e As Integer
Dim id As String
Dim error As Boolean
Private Sub Command1_Click()

End Sub

Private Sub cmdsearch_Click()
Dim u As New ADODB.Recordset
Set conn = New connection
np = cboprod.Text
Set r = con.Execute("select distinct * from products_tb where prod_id='" + np + "' order by name asc")
If Not r.EOF Then
r.MoveFirst
txtname.Text = r.Fields("name")
txtind.Text = r.Fields("industry")
id = r.Fields("prod_id")
Set s = con.Execute("select * from livestock where prod_id='" + np + "' order by transaction_id desc limit 1")
If Not s.EOF Then
aquantity.Caption = s.Fields("rest_quantity")
Else
aquantity.Caption = 0
End If
Set u = con.Execute("select * from products_update where prod_id='" + np + "' order by update_id desc limit 1")
If Not u.EOF Then
lastdate.Caption = u.Fields("date")
lastprice.Caption = u.Fields("unit_price")
Else
lastdate.Caption = "Not found"
lastprice.Caption = "Not found"
End If
Else
e = MsgBox("No such record found", vbCritical + vbOKOnly, "Error")
End If
End Sub
Private Sub Command2_Click()

End Sub

Private Sub buy_LostFocus()
If IsNumeric(nquantity.Text) And IsNumeric(buy.Text) Then
totalprice.Caption = Val(nquantity.Text) * Val(buy.Text)
Else
totalprice.Caption = ""
End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub nquantity_LostFocus()
If IsNumeric(nquantity.Text) And IsNumeric(buy.Text) Then
totalprice.Caption = Val(nquantity.Text) * Val(buy.Text)
End If
If IsNumeric(nquantity.Text) Then
tquantity.Caption = Val(nquantity.Text) + aquantity.Caption
End If
End Sub
Private Sub updateprod_Click()
Dim ld, cuq, lp, nq, nprice, tot, tq, exp, sell, mdate, dt, bl As String
If lastdate.Caption = "" Then
ld = Format(Now, "d/m/yyyy")
Else
ld = lastdate.Caption
End If
If Val(aquantity.Caption) = 0 Then
cuq = 0 + Val(nquantity.Text)
cuq = Str(cuq)
Else
cuq = Str(aquantity.Caption)
End If
If lastprice.Caption = "" Then
lp = 0
Else
lp = lastprice.Caption
End If
If nquantity.Text = "" Then
nq = 0
Else
nq = Str(nquantity.Text)
End If
If buy.Text = "" Then
nprice = 0
Else
nprice = Str(buy.Text)
End If
mdate = man.Text
ex = expire.Text
bl = billnber.Text
tot = Str(tquantity.Caption)
dt = Format(Now, "d-m-yyyy")
Set conn = New connection
Sql = "INSERT INTO `restobar`.`products_update` (`prod_id`, `actual_quantity`, `new_quantity`, `tot_quantity`, `manufactured_date`, `expiration_date`, `unit_price`,  `user`, `date`, `billno`) VALUES ('" + id + "', '" + cuq + "', '" + nq + "', '" + tot + "', '" + mdate + "', '" + ex + "', '" + nprice + "', 'admin', '" + dt + "', '" + bl + "');"
con.Execute Sql
Sql = "insert into livestock(prod_id,actual_quantity,out_quantity,rest_quantity,unit_price,total,bill_id,kind,date) values('" + id + "','" + cuq + "','0','" + tot + "','None','None','None','Entry', '" + dt + "')"
con.Execute Sql
e = MsgBox("Update done", vbInformation + vbOKOnly, "Information")
Me.Hide
price.Show
End Sub

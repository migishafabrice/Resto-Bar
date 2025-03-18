VERSION 5.00
Begin VB.Form price 
   Caption         =   "Pricing"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   8670
   Begin VB.Frame Frame2 
      Caption         =   "Search product"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   7695
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
         Left            =   5880
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
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
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Enter or select  name/code:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update price"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   7695
      Begin VB.TextBox txtunit 
         Height          =   360
         Left            =   2040
         TabIndex        =   19
         Top             =   2520
         Width           =   4335
      End
      Begin VB.TextBox txtcat 
         Height          =   360
         Left            =   2040
         TabIndex        =   17
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CommandButton cmdprice 
         Caption         =   "Save"
         Height          =   480
         Left            =   4680
         TabIndex        =   15
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtnprice 
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtlprice 
         Height          =   360
         Left            =   2040
         TabIndex        =   12
         Top             =   3000
         Width           =   4335
      End
      Begin VB.TextBox txtlast 
         Height          =   360
         Left            =   2040
         TabIndex        =   10
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox txtind 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtname 
         Height          =   360
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "Category:"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Unit:"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "New price:"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Last price:"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Last update:"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Industry:"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String
Private Sub cmdprice_Click()
Dim u, d, l, n As String
Dim price As String
Dim m As Integer
u = Str(txtunit.Text)
d = Format(Now, "d-m-yyyy")
If txtlprice = "Not found" Then
l = "0"
Else
l = Str(txtlprice.Text)
End If
n = Str(txtnprice.Text)
Set conn = New connection
con.Execute ("insert into price(prod_id,unit,last_price,new_price,user,date) values('" + p + "','" + u + "','" + l + "','" + n + "','admin','" + d + "')")
m = MsgBox("Done", vbInformation, "Information")
Me.Hide
End Sub
Private Sub cmdsearch_Click()
Dim rs, r As New Recordset
Dim search As String
Dim m As Integer
search = cboprod.Text
Set conn = New connection
Set rs = con.Execute("select * from products_tb where prod_id='" + search + "'")
If Not rs.EOF Then
rs.MoveFirst
p = rs.Fields("prod_id")
txtname = rs.Fields("name")
txtind = rs.Fields("industry")
txtcat.Text = rs.Fields("category")
Set r = con.Execute("select * from price where prod_id='" + search + "' order by price_id desc limit 1")
If Not r.EOF Then
txtlast.Text = r.Fields("date")
txtunit.Text = r.Fields("unit")
txtlprice = r.Fields("new_price")
Else
txtlast.Text = "Not found"
txtunit.Text = "Not found"
txtlprice = "Not found"
End If
Else
m = MsgBox("No record found", vbCritical, "Warning")
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'If user = "" Or pass = "" Then
'm = MsgBox("Username or passpword incorrect", vbCritical, "Warning")
'Me.Hide
'login.Show
'End If
txtname.Enabled = False
txtind.Enabled = False
txtlast.Enabled = False
txtlprice.Enabled = False
txtcat.Enabled = False
End Sub


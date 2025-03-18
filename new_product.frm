VERSION 5.00
Begin VB.Form new_product 
   Caption         =   "New product"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8130
   Begin VB.Frame Frame1 
      Caption         =   "New product"
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   7095
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         Height          =   735
         Left            =   4080
         TabIndex        =   10
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add new"
         Height          =   735
         Left            =   1920
         TabIndex        =   9
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox des 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1920
         TabIndex        =   8
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox prodid 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   3975
      End
      Begin VB.ComboBox category 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox prodname 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Industry:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Category:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label lbcode 
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Code of product:"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Add new product"
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
      Left            =   2760
      TabIndex        =   11
      Top             =   240
      Width           =   2205
   End
End
Attribute VB_Name = "new_product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Integer
Private Sub cmdadd_Click()
Dim p As New ADODB.Recordset
Dim code, pname, pcat, pind, pdes As String
pname = prodname.Text
pcat = category.Text
pind = prodid.Text
pdes = des.Text
Set conn = New connection
Set p = con.Execute("select id from products_tb order by id desc limit 1")
If Not p.EOF Then
code = "prod-" & p.Fields("id") + 1 & "-" & Format(Now, "yyyy")
Else
code = "prod-1-" & Format(Now, "yyyy")
End If
d = Format(Now, "d-m-yyyy")
con.Execute ("insert into products_tb(prod_id,name,category,industry,description,date,user) values('" + code + "','" + pname + "','" + pcat + "','" + pind + "','" + pdes + "','" + d + "','" + user + "')")
e = MsgBox("Product saved", vbInformation, "Information")
Me.Hide
update_product.Show
End Sub

Private Sub Form_Load()
Dim cat As New ADODB.Recordset
'If user = "" Or pass = "" Then
'm = MsgBox("Username or passpword incorrect", vbCritical, "Warning")
'Me.Hide
'login.Show
'End If
Set conn = New connection
Set p = con.Execute("select id from products_tb order by id desc limit 1")
If Not p.EOF Then
code = "prod-" & p.Fields("id") + 1 & "-" & Format(Now, "yyyy")
lbcode.Caption = code
Else
code = "prod-1-" & Format(Now, "yyyy")
lbcode.Caption = code
End If
Set cat = con.Execute("select distinct category from products_tb order by name asc")
If Not cat.EOF Then
cat.MoveFirst
While Not cat.EOF
category.AddItem cat.Fields("category")
cat.MoveNext
Wend
End If
End Sub

VERSION 5.00
Begin VB.Form nlogin 
   Caption         =   "New user"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtuser 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   12
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtpass1 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "."
      TabIndex        =   8
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox txtpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "."
      TabIndex        =   6
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox txtmail 
      Height          =   405
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox txtemp 
      Height          =   405
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblerror 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Username"
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
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Retype password"
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
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
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
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Employee code"
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
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New user "
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
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "nlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set conn = New connection
Dim c, e, p, ps, u As String
u = txtuser.Text
c = txtemp.Text
e = txtmail.Text
p = txtpass.Text
ps = txtpass1.Text
If p <> ps Then
m = MsgBox("Passwords not matching", vbCritical, "warning")
txtpass.SetFocus
Exit Sub
End If
If c = "" Or e = "" Or p = "" Or u = "" Then
m = MsgBox("Empty field not allowed", vbCritical, "warning")
Exit Sub
End If
Dim rs As New Recordset
Set rs = con.Execute("select * from security_tb where userid='" + c + "'")
If rs.EOF Then
Set rs = con.Execute("select * from employee_tb where userid='" + c + "'")
If Not rs.EOF And rs.Fields("function") <> "Cooker" And rs.Fields("function") <> "Serveur" Then
con.Execute ("insert into security_tb (userid,username,password,type) values('" + c + "','" + u + "','" + p + "','" + rs.Fields("function") + "')")
m = MsgBox("User saved", vbInformation, "Warning")
Me.Hide
login.Show
Else
m = MsgBox("Not eligible to have an account", vbInformation, "Warning")
Exit Sub
End If
Else
m = MsgBox("Account already exists", vbInformation, "Warning")
Exit Sub
End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub txtuser_Change()
Dim rs As New Recordset
Set conn = New connection
Set rs = con.Execute("select username from security_tb where username like'" + txtuser.Text + "'")
If Not rs.EOF Then
lblerror.Caption = "Username exists"
Else
lblerror.Caption = ""
End If

End Sub

Private Sub txtuser_LostFocus()
Dim rs As New Recordset
Set conn = New connection
Set rs = con.Execute("select username from security_tb where username ='" + txtuser.Text + "'")
If Not rs.EOF Then
m = MsgBox("Username exists", vbCritical, "Warning")
txtuser.SetFocus
Exit Sub
End If
End Sub

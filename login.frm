VERSION 5.00
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtuser 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1920
      TabIndex        =   4
      Text            =   "Enter username"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1920
      TabIndex        =   3
      Text            =   "Enter  password"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdlog 
      Caption         =   "Login"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdreset 
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
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbllogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN TO ACCESS SYSTEM"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   4950
   End
   Begin VB.Label lbluser 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label lblpass 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblerroruser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username is compusolry"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblerrorpass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password is compusolry"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Integer
Private Sub cmdlog_Click()
Dim access As New ADODB.Recordset
Dim u, p As String
u = txtuser.Text
p = txtpass.Text
Set conn = New connection
Set access = con.Execute("select * from security_tb where username='" + u + "' and password='" + p + "' ")
If access.EOF Then
e = MsgBox("username or password not found", vbOKOnly + vbCritical, "Error")
Exit Sub
End If
If Not access.EOF And access.RecordCount = 1 Then
typeuser = access.Fields("type")
user = access.Fields("userid")
username = access.Fields("username")
Me.Hide
homepage.Show
Else
e = MsgBox("username or password not found", vbOKOnly + vbCritical, "Error")
Exit Sub
End If
End Sub

Private Sub cmdnew_Click()
nlogin.Show
End Sub

Private Sub cmdreset_Click()
txtuser.Text = ""
txtpass.Text = ""
End Sub

Private Sub Form_Load()
lblerroruser.Visible = False
lblerrorpass.Visible = False
End Sub

Private Sub txtpass_Click()
txtpass.Text = ""
End Sub

Private Sub txtpass_LostFocus()
If txtpass = "" Then
lblerrorpass.Visible = True
lblerrorpass.BackColor = vbRed
Else
lblerrorpass.Visible = False
End If
End Sub

Private Sub txtuser_Click()
txtuser.Text = ""
End Sub

Private Sub txtuser_LostFocus()
If txtuser.Text = "" Then
lblerroruser.Visible = True
lblerroruser.BackColor = vbRed
Else
lblerroruser.Visible = False
End If
End Sub


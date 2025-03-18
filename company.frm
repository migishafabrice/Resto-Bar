VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form company 
   Caption         =   "Company information"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   7050
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
   ScaleHeight     =   6510
   ScaleWidth      =   7050
   Begin MSComDlg.CommonDialog logo 
      Left            =   5400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add company information"
      Height          =   5535
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6495
      Begin VB.TextBox sector 
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox district 
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton cmdcompany 
         Caption         =   "Save"
         Height          =   495
         Left            =   3000
         TabIndex        =   14
         Top             =   4680
         Width           =   2655
      End
      Begin VB.CommandButton cmdlogo 
         Caption         =   "................."
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ListBox pro 
         Height          =   540
         Left            =   1440
         TabIndex        =   9
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox mail 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox tel 
         Height          =   360
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtnm 
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   4215
      End
      Begin VB.Image imglogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Browse Logo:"
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sector:"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "District:"
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Province:"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Telephone:"
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company Information"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
End
Attribute VB_Name = "company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

End Sub

Private Sub cmdcompany_Click()
Dim nm, t, m, prov, di, se, p As String
nm = txtnm.Text
t = tel.Text
ml = mail.Text
prov = pro.Text
di = district.Text
se = sector.Text
If logo.FileName <> "" Then
ext = Right(logo.FileName, Len(logo.FileName) - InStrRev(logo.FileName, "."))
p = "\icons\" & nm & "." & ext
Call FileCopy(logo.FileName, App.Path & p)
Else
m = MsgBox("Browse peicture", vbCritical, "warning")
Exit Sub
End If
Set conn = New connection
con.Execute ("insert into identification(comp_id,name,telephone,email,sector,district,province,logo) values('" + cd + "','" + nm + "','" + t + "','" + ml + "','" + se + "','" + di + "','" + prov + "','" + p + "')")
m = MsgBox("Company information saved", vbInformation, "Information")
End Sub

Private Sub cmdlogo_Click()
 logo.InitDir = App.Path
            logo.FileName = ""
            logo.Filter = "JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
            logo.DialogTitle = "Open Image"
            logo.ShowOpen
            If logo.FileName <> "" Then
            imglogo.Picture = LoadPicture(logo.FileName)
            End If
End Sub

Private Sub Form_Load()
pro.AddItem "KIGALI City"
pro.AddItem "NORTH"
pro.AddItem "SOUTH"
pro.AddItem "EST"
pro.AddItem "WEST"
End Sub

Private Sub ncomp_Click()

End Sub

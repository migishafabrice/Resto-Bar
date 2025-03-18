VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form new_user 
   Caption         =   "New employee"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11640
   Begin MSComDlg.CommonDialog picshow 
      Left            =   10800
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame info 
      Caption         =   "New user infomation"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   7215
      Begin VB.ListBox funct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   31
         Top             =   3000
         Width           =   4575
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtsurname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Width           =   4575
      End
      Begin VB.ListBox lblday 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox lblmonth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3120
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox lblyear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4800
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox lblq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   10
         Top             =   2520
         Width           =   4575
      End
      Begin VB.TextBox txtphone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   9
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox txtmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   4080
         Width           =   4575
      End
      Begin VB.TextBox txtid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Top             =   4560
         Width           =   4575
      End
      Begin VB.TextBox txtdes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   5040
         Width           =   3975
      End
      Begin VB.OptionButton optmale 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton optfemale 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Function:"
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
         TabIndex        =   30
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Surname:"
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
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Date of birth:"
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
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Qualification:"
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
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Telephone:"
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
         TabIndex        =   22
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Email:"
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
         Left            =   360
         TabIndex        =   21
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "ID number:"
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
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Description bout experience:"
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
         TabIndex        =   19
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Sex:"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdimage 
      Caption         =   "Browse picture"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7920
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   480
      Left            =   7920
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
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
      Height          =   480
      Left            =   7920
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "New employee"
      Height          =   315
      Left            =   3240
      TabIndex        =   29
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label lbcode 
      Alignment       =   2  'Center
      Caption         =   "Code"
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label14 
      Caption         =   "Employee new code:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   27
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "new_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim code, f As String
Private Sub List3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdimage_Click()
 picshow.InitDir = App.Path
            picshow.FileName = ""
            picshow.Filter = "JPEG Image (*.jpg)|*.jpg|All Files (*.*)|*.*"
            picshow.DialogTitle = "Open Image"
            picshow.ShowOpen
            If picshow.FileName <> "" Then
            pic.Picture = LoadPicture(picshow.FileName)
            End If
End Sub

Private Sub cmdsave_Click()
Dim n, s, d, sex, q, t, m, i, des, a As String
n = txtname.Text
s = txtsurname.Text
d = lblday.Text & "/" & lblmonth.Text & "/" & lblyear.Text
b = lblq.Text
t = txtphone.Text
m = txtmail.Text
q = lblq.Text
i = txtid.Text
des = txtdes.Text
r = Format(Now, "d/m/yyyy")
If picshow.FileName <> "" Then
ext = Right(picshow.FileName, Len(picshow.FileName) - InStrRev(picshow.FileName, "."))
a = "/images/" & lbcode.Caption & "." & ext
Call FileCopy(picshow.FileName, App.Path & a)
Else
m = MsgBox("Browse peicture", vbCritical, "warning")
Exit Sub
End If
If optmale.Value = True Then
sex = optmale.Caption
ElseIf optfemale.Value = True Then
sex = optfemale.Caption
Else
e = MsgBox("Choose sex", vbCritical + vbOKOnly, "Error")
Exit Sub
End If
e = MsgBox("You are about to save:" + vbCr + "Name:" + n + vbCr + "Surname" + s + vbCr + "Date of birth:" + d + vbCr + "Sex:" + sex + vbCr + "Qualification:" + q + vbCr + "Telephone:" + t + vbCr + "Email:" + m + vbCr + "ID number/PassportID:" + id + vbCr + "Experience:" + des, vbYesNo + vbInformation, "Save employee")
If e = vbYes Then
con.Execute ("insert into employee_tb(userid,name,surname,dob,sex,qualification,field,function,phone,mail,id_card,experience,reg_date,picture) values('" + code + "','" + n + "','" + s + "','" + d + "','" + sex + "','" + q + "','" + f + "','" + funct.Text + "','" + t + "','" + m + "','" + i + "','" + des + "','" + r + "','" + a + "')")
e = MsgBox("Employee saved", vbOKOnly + vbInformation, "Information")
Else
e = MsgBox("Employee not saved", vbOKOnly + vbCritical, "Information")
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim d, m, y As Integer
'If user = "" Or pass = "" Then
'm = MsgBox("Username or passpword incorrect", vbCritical, "Warning")
'Me.Hide
'login.Show
'End If
Set conn = New connection
Dim last As New ADODB.Recordset
lblday.AddItem "Day"
For d = 1 To 31
lblday.AddItem d
Next d
lblmonth.AddItem "Month"
For m = 1 To 12
lblmonth.AddItem m
Next m
lblyear.AddItem "Year"
For y = 1930 To Format(Now, "yyyy") - 18
lblyear.AddItem y
Next y
lblq.AddItem "No one"
lblq.AddItem "Certificate"
lblq.AddItem "Undergraduate"
lblq.AddItem "A2"
lblq.AddItem "A1"
lblq.AddItem "A0"
lblq.AddItem "Masters"
lblq.AddItem "PHD"
lblq.AddItem "Professorate"
funct.AddItem "Manager"
funct.AddItem "Store Manager"
funct.AddItem "Kitchen Manager"
funct.AddItem "Cashier"
funct.AddItem "Controller"
funct.AddItem "Cooker"
funct.AddItem "Serveur"
optmale.Value = False
optfemale.Value = False
Set last = con.Execute("select id from employee_tb order by id  desc limit 1")
id = last.Fields("id") + 1
code = "emp-" & id & "-" & Format(Now, "yyyy")
lbcode.Caption = code
End Sub
Private Sub lblq_Click()
If (lblq.Text <> "No one" And f = "") Or (lblq.Text <> "No one") Then
f = InputBox("Enter field of qualification", "Qualification", "")
Do While f = ""
If StrPtr(f) = 0 Then
Exit Do
End If
f = InputBox("Enter field of qualification", "Qualification", "")
Loop
End If
End Sub
Private Sub lblq_LostFocus()
If lblq.Text <> "No one" And f = "" Then
f = InputBox("Please enter field of qualification", "Qualification", "")
Do While f = ""
If StrPtr(f) = 0 Then
f = InputBox("Please enter field of qualification", "Qualification", "")
End If
Loop
End If

End Sub



Private Sub List1_Click()

End Sub

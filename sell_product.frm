VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form sell_product 
   Caption         =   "Selling"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   8220
   Begin VB.Frame Frame1 
      Caption         =   "Sell"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7695
      Begin VB.ListBox lbloc 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1800
         TabIndex        =   24
         Top             =   5040
         Width           =   5175
      End
      Begin VB.TextBox txtind 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   20
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtcat 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   18
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   17
         Top             =   1080
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid flprod 
         Height          =   855
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1508
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtnew 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox cbserver 
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
         Left            =   1800
         TabIndex        =   3
         Top             =   4560
         Width           =   5175
      End
      Begin VB.CommandButton cmdchart 
         Caption         =   "Add to chart"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   5640
         Width           =   2175
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print bill"
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
         Left            =   4440
         TabIndex        =   1
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Bloc:"
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
         TabIndex        =   23
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label lbtotal 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4920
         TabIndex        =   22
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label8 
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
         Left            =   3360
         TabIndex        =   21
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Industry"
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
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Category:"
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
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Select product:"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Price/unit:"
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
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lbunit 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity:"
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
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Server:"
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
         TabIndex        =   9
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Rest quantity:"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lbrest 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4920
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Total amount:"
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
         TabIndex        =   6
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label lbtotprice 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1800
         TabIndex        =   5
         Top             =   4080
         Width           =   5175
      End
   End
End
Attribute VB_Name = "sell_product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, pr As New Recordset
Dim id, bloc As String
Dim prodid(), quantity(), price(), rest(), total(), paid() As String
Private Sub cmdchart_Click()
If lbloc.Text = "--Choose bloc--" Or lbloc.Text = "" Then
m = MsgBox("Choose the bloc", vbCritical, "Warning")
lbloc.SetFocus
Exit Sub
End If
ReDim Preserve prodid(nber)
ReDim Preserve quantity(nber)
ReDim Preserve price(nber)
ReDim Preserve rest(nber)
ReDim Preserve total(nber)
ReDim Preserve paid(nber)
nber = nber + 1
i = nber - 1
If i > 0 Then
For u = 0 To i - 1
If prodid(u) = id Then
m = MsgBox("This item already exists in the item" + vbCr + "Choose a different item", vbCritical, "Warning")
nber = nber - 1
Exit Sub
Else
prodid(i) = id
quantity(i) = txtnew.Text
price(i) = lbunit.Caption
rest(i) = lbrest.Caption
total(i) = lbtotal.Caption
paid(i) = lbtotprice.Caption
End If
Next
Else
prodid(i) = id
quantity(i) = txtnew.Text
price(i) = lbunit.Caption
rest(i) = lbrest.Caption
total(i) = lbtotal.Caption
paid(i) = lbtotprice.Caption
End If
Call reset
End Sub

Private Sub cmdprint_Click()
Dim save, printed As New Recordset
Dim t As Integer
m = MsgBox("You are about to save data" + vbCr + "Proced?", vbInformation + vbYesNo, "About to save")
bloc = lbloc.Text
If m = vbYes Then
Set conn = New connection
Set save = con.Execute("select id from billing order by id desc limit 1")
If Not save.EOF Then
bill = "000" & save.Fields("id") + 1 & Format(Now, "yy")
Else
bill = "0001" & Format(Now, "yy")
End If
con.Execute ("insert into billing(bill_id,user,serveur,date,bloc) values('" + bill + "','','admin','" + Format(Now, "d-m-yyyy") + "','" + bloc + "')")
For i = 0 To nber - 1
t = Val(quantity(i)) * Val(price(i))
Set save = con.Execute("INSERT INTO `restobar`.`livestock` ( `prod_id`, `actual_quantity`, `out_quantity`, `rest_quantity`, `unit_price`, `total`, `bill_id`, `kind`, `date`) VALUES ('" + prodid(i) + "', '" + total(i) + "', '" + quantity(i) + "', '" + rest(i) + "','" + price(i) + "', '" + Str(t) + "', '" + bill + "', 'Out', '" + Format(Now, "d-m-yyyy") + "');")
Next
m = MsgBox("Print", vbQuestion + vbYesNo, "Information")
If m = vbYes Then
printing.Show
End If
nber = 0
End If
End Sub

Private Sub flprod_Click()
With flprod
id = .TextMatrix(flprod.Row, 0)
txtname.Enabled = False
txtind.Enabled = False
txtcat.Enabled = False
txtname.Text = .TextMatrix(flprod.Row, 1)
txtind.Text = .TextMatrix(flprod.Row, 2)
txtcat.Text = .TextMatrix(flprod.Row, 3)
Set conn = New connection
Set r = con.Execute("select * from livestock where prod_id='" + id + "' order by transaction_id desc limit 1")
Set pr = con.Execute("select * from price where prod_id='" + id + "' order by price_id desc limit 1")
If Not r.EOF And Not pr.EOF Then
lbunit.Caption = pr.Fields("new_price")
lbtotal.Caption = r.Fields("rest_quantity")
End If
End With
End Sub
Private Sub Form_Load()
Dim rs As New Recordset
Dim p, n As String
Dim t As Integer
t = 0
lbloc.AddItem "--Choose bloc--"
lbloc.AddItem "Bar"
lbloc.AddItem "Resto"
With flprod
.ColWidth(0) = .Width / 4
.ColWidth(1) = .Width / 4
.ColWidth(2) = .Width / 4
.ColWidth(3) = .Width / 4
End With
With flprod
.Rows = 1
.Cols = 4
.TextMatrix(0, 0) = "Product ID"
.TextMatrix(0, 1) = "Name"
.TextMatrix(0, 2) = "Industry"
.TextMatrix(0, 3) = "Category"
End With
Set conn = New connection
Set r = con.Execute("select * from products_tb")
If Not r.EOF Then
r.MoveFirst
While Not r.EOF
p = r.Fields("prod_id")
Set rs = con.Execute("select * from livestock where prod_id='" + p + "' and rest_quantity!='0' order by transaction_id desc limit 1")
If Not rs.EOF Then
With flprod
.Rows = t + 2
.Cols = 4
.TextMatrix(t + 1, 0) = r.Fields("prod_id")
.TextMatrix(t + 1, 1) = r.Fields("name")
.TextMatrix(t + 1, 2) = r.Fields("Industry")
.TextMatrix(t + 1, 3) = r.Fields("category")
t = t + 1
End With
End If
r.MoveNext
Wend
End If
End Sub

Private Sub txtnew_Change()
'Call txtnew_Validate(False)
If IsNumeric(Val(txtnew.Text)) Then
lbrest.Caption = Val(lbtotal.Caption) - Val(txtnew.Text)
Else
m = MsgBox("Enter only numbers", vbCritical, "Warning")
End If
If Val(lbrest.Caption) < 0 Then
m = MsgBox("You entered a greater quantity" + vbCr + " than the available in store", vbCritical, "Warning")
txtnew.Text = ""
lbrest.Caption = ""
lbtotprice.Caption = ""
Exit Sub
Else
lbtotprice.Caption = Val(txtnew.Text) * Val(lbunit.Caption)
End If
End Sub
Private Sub txtnew_Validate(Cancel As Boolean)
'If Not Val(txtnew.Text) Then
'm = MsgBox("Price must be an integer", vbCritical, "Warning")
'Cancel = True
'End If
End Sub


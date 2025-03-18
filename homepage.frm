VERSION 5.00
Begin VB.MDIForm homepage 
   BackColor       =   &H8000000C&
   Caption         =   "Home Page"
   ClientHeight    =   4455
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9255
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu new 
      Caption         =   "New"
      Begin VB.Menu nemployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu nprod 
         Caption         =   "Product"
      End
      Begin VB.Menu nprice 
         Caption         =   "Price"
      End
      Begin VB.Menu ncomp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu edemployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu edprod 
         Caption         =   "Product"
      End
      Begin VB.Menu edprice 
         Caption         =   "Price"
      End
      Begin VB.Menu edcomp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu nsell 
      Caption         =   "Sell"
   End
   Begin VB.Menu nreport 
      Caption         =   "Report"
   End
End
Attribute VB_Name = "homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub edprice_Click()
price.Show
End Sub

Private Sub edprod_Click()
update_product.Show
End Sub
Private Sub MDIForm_Load()
Dim rs As New Recordset
Set conn = New connection
Set rs = con.Execute("select * from identification")
pic = rs.Fields("logo")
Call check
End Sub
Private Sub Picture1_Click()

End Sub
Private Sub ncomp_Click()
company.Show
End Sub
Private Sub nemployee_Click()

new_user.Show
End Sub

Private Sub nprice_Click()
price.Show
End Sub

Private Sub nprod_Click()
new_product.Show
End Sub

Private Sub nreport_Click()
Call check
reporting.Show
End Sub

Private Sub nsell_Click()
sell_product.Show
End Sub

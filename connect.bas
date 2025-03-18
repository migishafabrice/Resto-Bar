Attribute VB_Name = "connect"
Public con As New ADODB.connection
Public m, nber, oldPos As Integer
Public bill, user, pass, typeuser, status, username As String
Public startdate, enddate As String
Public Function reset()
sell_product.txtname.Text = ""
sell_product.txtind.Text = ""
sell_product.txtcat.Text = ""
sell_product.lbunit.Caption = ""
sell_product.txtcat.Text = ""
sell_product.txtind.Text = ""
sell_product.txtnew.Text = ""
sell_product.lbrest.Caption = ""
sell_product.lbtotal.Caption = ""
sell_product.lbtotprice.Caption = ""
'sell_product.txtpaid.Text = ""
'sell_product.lbresting.Caption = ""
End Function
Public Function check()
If typeuser = "Store keeper" Then
reporting.cmdbll.Enabled = False
reporting.billbar.Enabled = False
reporting.mbill.Enabled = False
reporting.billresto = False
reporting.billdetailed = False
reporting.mresto.Enabled = False
reporting.mbar.Enabled = False
reporting.mdetailed.Enabled = False
End If
End Function

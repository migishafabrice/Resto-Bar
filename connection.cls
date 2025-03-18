VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub class_initialize()
DoEvents
Set con = New ADODB.connection

With con
.ConnectionString = "Driver={MySql ODBC 3.51 Driver};SERVER=localhost;PWD=;UID=root;port=3306;DATABASE=restobar;"
.CursorLocation = adUseClient
.Open
End With
End Sub
 


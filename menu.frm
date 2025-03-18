VERSION 5.00
Begin VB.Form mnfile 
   Caption         =   "Menu"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu nmenu 
      Caption         =   "New"
      Begin VB.Menu nuser 
         Caption         =   "User"
      End
      Begin VB.Menu nprod 
         Caption         =   "Product"
      End
      Begin VB.Menu ncomp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu nedit 
      Caption         =   "Edit"
      Begin VB.Menu euser 
         Caption         =   "User"
      End
      Begin VB.Menu eprod 
         Caption         =   "Product"
      End
      Begin VB.Menu ecomp 
         Caption         =   "Company"
      End
   End
   Begin VB.Menu nsell 
      Caption         =   "Sell"
   End
   Begin VB.Menu nreport 
      Caption         =   "Reports"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "mnfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Begin VB.Form welcome 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Timer1_Timer()
Unload welcome
End Sub

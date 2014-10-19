VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   216
   ClientTop       =   1380
   ClientWidth     =   7488
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4560
   ScaleWidth      =   7488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   1200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    frmSplash.Timer1.Interval = 800
    
    ''MDIForm1.Show
    ''MDIForm1.WindowState = 2
    ''Wenti_form.Width = MDIForm1.Width / 2
    ''Draw_form.Width = MDIForm1.Width / 2
    Unload Me
    
    Call Main
   
    ''MDIForm1
  
    ''MDIForm1.WindowState = 2
    ''Wenti_form.Width = MDIForm1.Width / 2
    ''Draw_form.Width = MDIForm1.Width / 2
    ''MDIForm1.WindowState = 2
    ''frmTip     ''frmSplash
    ''frmSplash.visible = True
    
End Sub

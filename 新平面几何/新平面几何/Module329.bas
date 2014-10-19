Attribute VB_Name = "Module329"
Option Explicit

Sub Main()
    'Display application window
    
     frmTip.WindowState = 0
    MDIForm1.Show
    MDIForm1.WindowState = 2
       
    
    Wenti_form.Width = MDIForm1.Width / 2
    Draw_form.Width = MDIForm1.Width / 2
    'Since Tip of the Day can unload itself,
    'trap the error
    On Error Resume Next
    'Display Tip of the Day
    frmTip.Show vbModeless, MDIForm1
    frmTip.WindowState = 0

    Wenti_form.Width = MDIForm1.Width / 2
    Draw_form.Width = MDIForm1.Width / 2
End Sub

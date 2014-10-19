Attribute VB_Name = "Module43"
Option Explicit

Sub Main()
    If event_statue <> exit_program Then
    MDIForm1.WindowState = 2
    'Since Tip of the Day can unload itself,
    'trap the error
   On Error Resume Next
    ''If frmTip.chkLoadTipsAtStartup.value = 1 Then
    'If regist_data.language = 1 Then
    If event_statue <> exit_program Then
    frmTip.Show vbModeless, MDIForm1
    End If
    'End If
    MDIForm1.WindowState = 2
    Wenti_form.width = MDIForm1.width / 2
    Draw_form.width = MDIForm1.width / 2
    End If
End Sub

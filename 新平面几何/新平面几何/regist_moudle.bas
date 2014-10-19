Attribute VB_Name = "regist_moudle"
Option Explicit
Type regist_data_type
study_progrss As Integer
th_chose(-10 To 200) As Integer
Tool_bar_visible As Boolean
projector_Checked As Boolean
yuantibudong_Checked As Boolean
treeview_Checked As Boolean
set_password_Checked As Boolean
mnuTip_Checked As Boolean
line_width  As Byte
condition_color As Byte
conclusion_color As Byte
fill_color  As Byte
server_IP As String * 15
server_port As Integer
teacher_password As String * 5
computer_id As String * 8
run_type As Byte
language As Byte
End Type
Global regist_data As regist_data_type
Public Sub set_regist()
   regist_data.condition_color = condition_color
   regist_data.conclusion_color = conclusion_color
   regist_data.fill_color = fill_color
   regist_data.line_width = line_width
   Open App.path & "\regist.dat" For Random As #3 Len = Len(regist_data)
     Put #3, 1, regist_data
   Close #3
End Sub
Public Sub get_regist()
Dim i%
   Open App.path & "\regist.dat" For Random As #5 Len = Len(regist_data)
     Get #5, 1, regist_data
   Close #5
   For i% = -6 To 180
      th_chose(i%).chose = regist_data.th_chose(i%)
   Next i%
   If regist_data.language = 0 Then
       regist_data.language = 1
   End If
   global_icon_char = LoadResString_(1225, "")

 End Sub
 Public Sub initial_set()
   condition_color = regist_data.condition_color
   conclusion_color = regist_data.conclusion_color
   fill_color = regist_data.fill_color
   line_width = regist_data.line_width
   MDIForm1.Toolbar1.visible = regist_data.Tool_bar_visible
   If regist_data.Tool_bar_visible Then
    set_form.Check1.value = 1
   Else
    set_form.Check1.value = 0
   End If
   '    Call init_project_(regist_data.projector_Checked)
    If regist_data.projector_Checked Then
     set_form.Check2.value = 1
    Else
     set_form.Check2.value = 0
    End If
    If regist_data.yuantibudong_Checked Then
      Call C_display_wenti.set_m_yuantibudong(True)
      set_form.Check6.value = 1
    Else
      Call C_display_wenti.set_m_yuantibudong(False)
      set_form.Check6.value = 0
    End If
'    Call init_tree_(regist_data.treeview_Checked)
   If regist_data.treeview_Checked Then
     set_form.Check3.value = 1
    Else
     set_form.Check3.value = 0
    End If
    If regist_data.mnuTip_Checked Then
     set_form.Check5.value = 1
    Else
     set_form.Check5.value = 0
    End If
    If regist_data.set_password_Checked Then
        If protect_data.pass_word_for_teacher <> "00000" Then
          set_form.Check4.value = 1
           regist_data.teacher_password = protect_data.pass_word_for_teacher
            'Call C_display_wenti.set_m_pass_word_for_teacher(False)
        ElseIf regist_data.teacher_password <> "00000" Then
           protect_data.pass_word_for_teacher = regist_data.teacher_password
            'Call C_display_wenti.set_m_pass_word_for_teacher(True)
           set_form.Check4.value = 1
        Else
           set_form.Check4.value = 0
         regist_data.set_password_Checked = False
        End If
    Else
         set_form.Check4.value = 0
         regist_data.teacher_password = "00000"
         protect_data.pass_word_for_teacher = "00000"
    End If
    If regist_data.run_type = 0 Then
       set_form.Check7.value = 1
       set_form.Check8.value = 0
       MDIForm1.Caption = Mdiform1_caption + "-" + LoadResString_(1135, "")
    ElseIf regist_data.run_type = 1 Then
       set_form.Check7.value = 0
       set_form.Check8.value = 1
       MDIForm1.Caption = MDIForm1.Caption + "-" + LoadResString_(1140, "") ' "版—传统证明方法"
    End If
End Sub


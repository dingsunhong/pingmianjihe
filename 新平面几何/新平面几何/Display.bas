Attribute VB_Name = "displ"
Option Explicit
'Dim empty_char As String * 1
Global display_string1 As String
Global display_string2 As String
Public Function display_char(ob As Object, ByVal A$, up_down As Integer, color As Byte, _
                     display_or_delete As Byte, is_display_icon As Boolean) As String
'ob 显示的平台，A显示的内容up_down显示上下标color%颜色，display_or_delete 显示或消除
Dim font As String
Dim old_currentx As Integer
Dim old_currenty As Integer
Dim new_currentx As Integer
Dim fontZ%
Dim i%
Dim op_ch$
Dim tA$
Dim ch$
Dim icon$
Dim up_modify%, abs_modify%, dwon_modify%, n%
display_char = ""
If ob.ScaleMode = 3 Then '初始化
      up_modify% = 1
       dwon_modify% = 5
         abs_modify% = 3
End If
'当前焦点坐标，字体字号
old_currentx = ob.CurrentX
old_currenty = ob.CurrentY
fontZ% = ob.FontSize
font = ob.font
'**************************************************
'代码转换
For i% = 1 To Len(A$)
ch$ = Mid$(A$, i%, 1)
            If ch$ = Chr(12) Then  '"," Then
             tA$ = tA$ + LoadResString_(1380, "")
            ElseIf ch$ = "$" Then
             tA$ = tA$ + "sin" + LoadResString_(1380, "")
            ElseIf ch = "&" Then
             tA$ = tA$ + "cos" + LoadResString_(1380, "")
            ElseIf ch = "`" Then
             tA$ = tA$ + "tan" + LoadResString_(1380, "")
            ElseIf ch = "\" Then
             tA$ = tA$ + "ctan" + LoadResString_(1380, "")
            ElseIf ch = "'" Then
             tA$ = tA$ + LoadResString_(1460, "")
            ElseIf ch = "@" Then '弧的记号
              tA$ = tA$ + "-"
            ElseIf ch = "#" Then
              tA$ = tA$ + "+"
            ElseIf ch = "^" Then
             ch$ = Mid$(A$, 1, i% - 1)
             display_char = display_char & _
               display_char(ob, ch$, up_down, color, display_or_delete, is_display_icon)
             ch$ = Mid$(A$, i% + 1, 1)
             display_char = display_char & _
               display_char(ob, ch$, up, color, display_or_delete, is_display_icon)
             If Len(A$) > i% Then
             ch$ = Mid$(A$, i% + 2, Len(A$) - i% - 1)
             display_char = display_char & _
              display_char(ob, ch$, up_down, color, display_or_delete, is_display_icon)
               tA$ = ""
             Else
              up_down = up
              Exit Function
             End If
               GoTo display_char_out
            ElseIf ch = global_icon_char Then
             ch$ = Mid$(A$, 1, i% - 1)
             display_char = display_char & _
              display_char(ob, ch$, up_down, color, display_or_delete, is_display_icon)
                  ob.font = LoadResString_(1231, "") ' "方正舒体"
             ob.Print global_icon_char;
             display_char = display_char & global_icon_char
             ob.font = font
             ch$ = Mid$(A$, i% + 1, Len(A$) - i%)
             display_char = display_char & _
              display_char(ob, ch$, up_down, color, display_or_delete, is_display_icon)
               tA$ = ""
               GoTo display_char_out
            ElseIf ch = "~" Or ch = "!" Then
             'Exit Sub
            Else
               tA$ = tA$ + ch
            End If
Next i%
'****************************************************************
Call SetTextColor_(ob, color, display_or_delete)
If Len(tA$) > 1 Then
     If up_down = up Then
     ob.CurrentY = ob.CurrentY + up_modify%
     ob.FontSize = 4       '下标
      ob.Print tA$;
     ob.CurrentY = ob.CurrentY - up_modify%
     ob.FontSize = 12       '下标
     ElseIf up_down = down Then
      ob.CurrentY = ob.CurrentY + dwon_modify%
       ob.FontSize = 4       '下标
        ob.Print tA$;
        ob.CurrentY = ob.CurrentY - dwon_modify%
         ob.FontSize = 12 '下标
     ElseIf up_down = arrow Then
      Call print_char_with_arrow(ob, color, tA$)
     Else
       ob.FontSize = 12
        ob.Print tA$;
        display_char = display_char & tA$
     End If
    'End If
    ob.font = font
Else
If tA$ = "" Then
  Exit Function
ElseIf Len(tA$) > 1 Then
 If Mid$(tA$, 1, 1) = "?" And Mid$(tA$, Len(tA$), 1) = "}" Then
    ob.FontSize = 12
     old_currentx = ob.CurrentX
      tA$ = Mid$(tA$, 2, Len(tA$) - 2)
      ob.Print tA$;
       new_currentx = ob.CurrentX
         old_currenty = ob.CurrentY
       If display_or_delete = 0 Then
       ob.Line (old_currentx, old_currenty + 3)-(new_currentx, old_currenty + 3), QBColor(15)
       ob.Line (new_currentx - 8, old_currenty + 1)-(new_currentx, old_currenty + 3), QBColor(15)
       ElseIf display_or_delete = 1 Then
       ob.Line (old_currentx, old_currenty + 3)-(new_currentx, old_currenty + 3), QBColor(9)
       ob.Line (new_currentx - 8, old_currenty + 1)-(new_currentx, old_currenty + 3), QBColor(9)
       ElseIf display_or_delete = 2 Then
       ob.Line (old_currentx, old_currenty + 3)-(new_currentx, old_currenty + 3), QBColor(7)
       ob.Line (new_currentx - 8, old_currenty + 1)-(new_currentx, old_currenty + 3), QBColor(7)
       End If
       Exit Function
 ElseIf Mid$(tA$, 1, 1) = "|" And Mid$(tA$, Len(tA$), 1) = "|" And Len(tA$) > 1 Then
        old_currenty = ob.CurrentY
        ob.CurrentY = old_currenty - abs_modify%
    ob.Print "|";
     ob.CurrentY = old_currenty
    tA$ = Mid$(tA$, 2, Len(tA$) - 2)
    tA$ = "?" & tA$ & "}"
    Call display_char(ob, tA$, up_down, color, display_or_delete, False)
            ob.CurrentY = old_currenty - abs_modify%
    ob.Print "|";
     ob.CurrentY = old_currenty
     Exit Function
 Else
  If Mid$(tA$, 1, 1) = "?" Then
   tA$ = Mid$(tA$, 2, Len(tA$) - 1)
   Call display_char(ob, Mid$(tA$, 1, 1), up_down, color, display_or_delete, False)
   Call display_char(ob, Mid$(tA$, 2, Len(tA$) - 1), up_down, color, _
                    display_or_delete, False)
  Else
       If up_down = up Then
        ob.CurrentY = ob.CurrentY + up_modify%
       ob.FontSize = 4       '下标
      ob.Print tA$;
     ob.CurrentY = ob.CurrentY - up_modify%
     ob.FontSize = 12       '下标
     ElseIf up_down = down Then
      ob.CurrentY = ob.CurrentY + dwon_modify%
       ob.FontSize = 4       '下标
        ob.Print tA$;
        ob.CurrentY = ob.CurrentY - dwon_modify%
         ob.FontSize = 12 '下标
     ElseIf up_down = arrow Then
      Call print_char_with_arrow(ob, color, tA$)
     Else
       ob.FontSize = 12
        ob.Print tA$;
     End If
    'End If
    ob.font = font
   End If
                      Exit Function
End If
   'Call SetTextColor_(Ob, color%, display_or_delete)
Else
If tA$ = "^" Or tA$ = "!" Or tA$ = ";" Then
 Exit Function
End If
If up_down = up Then
 ob.CurrentY = ob.CurrentY + up_modify%
  ob.FontSize = 4       '下标
   ob.Print tA$;
 ob.CurrentY = ob.CurrentY - up_modify%
  ob.FontSize = 12       '下标
ElseIf up_down = down Then
    ob.CurrentY = ob.CurrentY + dwon_modify%
     ob.FontSize = 4       '下标
      ob.Print tA$;
       ob.CurrentY = ob.CurrentY - dwon_modify%
        ob.FontSize = 12 '下标
ElseIf up_down = arrow Then
 Call print_char_with_arrow(ob, color, tA$)
Else
 If tA$ <> "" Then
 If Asc(tA$) > 0 And Asc(tA$) < 128 And tA$ <> Chr(17) Then
   If tA$ = "!" Or Asc(tA$) = 13 Or tA$ = "?" Then '控制符
     Exit Function
   End If
'   If tA$ = "_" Then
'    Call m_icon.set_m_icon("_", up_down, 0, display_or_delete)
'   End If
    ob.Print tA$; '英文
    display_char = display_char & tA$
   If tA$ = "_" Then 'LoadResString_(113) Then
         ob.CurrentX = ob.CurrentX + 2
   End If
 Else
   'Call SetTextColor_(Ob, 3, display_or_delete)
   If system_vision = 2 And tA$ = Chr(17) Then 'tA$ = LoadResString_(1226,"") Then
     ob.font = LoadResString_(1230, "") ' "方正舒体"
          tA$ = global_icon_char
      ob.Print tA$; '汉字
      display_char = display_char & tA$
       ob.font = font
   Else
     If tA$ = global_icon_char Then
        ob.font = LoadResString_(1230, "") ' "方正舒体"
     End If
        ob.Print tA$; '汉字
        display_char = display_char & tA$
        ob.font = font
   End If
 End If
End If
End If
End If
End If
display_char_out:
ob.CurrentY = old_currenty
ob.FontSize = fontZ%
 ob.font = font
If display_or_delete = 1 Then
   If ob.CurrentX > ob.ScaleWidth - 40 Then
   ob.CurrentY = old_currenty
   ob.CurrentX = old_currentx
   Call display_char(ob, A$, up_down, color, _
                     0, is_display_icon)
   display_char = change_L
'ob 显示的平台，A显示的内容up_down显示上下标color%颜色，display_or_delete 显示或消除
   ob.CurrentY = old_currenty
   ob.CurrentX = old_currentx
   End If
End If
End Function
Public Sub display_prove_inform(ByVal no%, ByVal ty As Boolean)
Dim s$
If prove_type = 0 Then
s$ = LoadResString_(1680, "")
ElseIf prove_type = 1 Then
s$ = LoadResString_(1685, "")
ElseIf prove_type = 1 Then
s$ = LoadResString_(1690, "")
ElseIf prove_type = 3 Then
s$ = LoadResString_(1695, "")
ElseIf prove_type = 4 Then
If problem_type = False Then
s$ = LoadResString_(1700, "") + "!"
End If
End If
Wenti_form.Picture1.CurrentY = 20 * (no% + 4)
Wenti_form.Picture1.CurrentX = 40
If ty = display Then
 Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(12))
ElseIf ty = delete Then
 Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(15))
End If
Wenti_form.Picture1.Print s$
 Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
End Sub
Public Sub display_problem_text(t As Boolean)
' t=false t=true
Dim temp_string As String
Dim i%, j%, k%
i% = Len(problem_text)
temp_string = "　　" + Mid$(problem_text, 1, 16)
Wenti_form.Picture1.Print temp_string
Wenti_form.Picture1.Print
j% = (i% - 16) Mod 18
For k% = 0 To j% - 1
temp_string = "　　" + Mid$(problem_text, 17 + 18 * k%, 16)
Wenti_form.Picture1.Print temp_string
Wenti_form.Picture1.Print
Next k%
temp_string = "　　" + Mid$(problem_text, 17 + 18 * j%, i% - (17 + 18 * j%))
Wenti_form.Picture1.Print temp_string
If t = False Then
Wenti_form.Picture1.Print global_icon_char
Else
Wenti_form.Picture1.Print
End If
End Sub
Public Sub clear_wenti_display()
If Wenti_form.List1.visible Then
   Wenti_form.List1.visible = False
   Wenti_form.SSTab1.Tab = 1
   Wenti_form.SSTab1.Caption = LoadResString_(4230, "")
   SSTab1_name_type = 0
   Wenti_form.Picture2.Cls
    Set C_display_picture = Nothing
    Set C_display_picture = New display_picture
End If
End Sub
Public Sub display_run(ByVal ty As Byte)
Dim ty1, ty2 As Boolean
Dim aid_statue As Boolean
Dim i%, j%, p%, k%, l%
Dim tp(1) As Integer
Dim t_y As Byte
Dim t As Boolean
Dim tid1 As Long
If finish_prove = 1 Then
 Exit Sub
ElseIf finish_prove = 0 Then
 finish_prove = 1
ElseIf finish_prove > 3 Then
 GoTo method_mark1
End If '优化证明
Call begin_prove
'Call C_display_wenti.m_display_start_prove(Wenti_form.Picture1)
'conclusion_no_wenti = wenti_no
If run_statue = 6 Then '12.10
 GoTo display_run_mark5
End If
run_statue = 1 '12.10
'old_wenti_no = C_display_wenti.m_last_input_wenti_no
'For i% = 1 To last_conditions.last_cond(1).circle_no
'Call simple_circle(i%)
'Next i%
''如果已有结论
'For i% = 1 To last_conditions.last_cond(1).point_no
'    m_poi(i%).data(0).no_reduce = 255
'Next i%
'For i% = 1 To last_conditions.last_cond(1).line_no
'    m_lin(i%).data(0).no_reduce = 255
'Next i%
'For i% = 1 To last_conditions.last_cond(1).circle_no
'    m_Circ(i%).data(0).no_reduce = 255
'Next i%
'For i% = 1 To last_conditions.last_cond(1).angle_no
'    angle(i%).data(0).no_reduce = 255
'Next i%
'For i% = 1 To last_conditions.last_cond(1).triangle_no
'    triangle(i%).data(0).no_reduce = 255
'Next i%
If find_conclusion1(0, 0, True) > 0 Then
 GoTo method_mark1
'Else
' Call pre_set_reduce
' Call set_point_relation_reduce(True)
' Call post_set_reduce
End If
run_statue = 2 '12.10
'Call CreateThread(0, 4096, start_prove, 0, 0, tid1)
   If chose_total_theorem = False Then
     th_chose(156).chose = 0
     th_chose(155).chose = 0
     th_chose(2).chose = 0
   End If
For i% = 0 To last_conclusion - 1
If conclusion_data(i%).ty = area_of_element_ Then
     th_chose(156).chose = 1
     th_chose(155).chose = 1
     th_chose(2).chose = 1
End If
Next i%
If last_conditions.last_cond(1).area_of_element_no > 0 Then
     th_chose(156).chose = 1
     th_chose(155).chose = 1
     th_chose(2).chose = 1
End If
If last_conditions.last_cond(1).tri_function_no > 0 Then
 th_chose(118).chose = 1
End If
prove_result = start_prove(0, reduce_level, 0)
 If error_of_wenti > 0 And error_of_wenti < 100 Then
  GoTo display_run_wenti_error
 End If
'暂时关闭Else
If prove_result = 2 Then
  GoTo method_mark1
ElseIf prove_result = 5 Then
display_run_mark0:
  MDIForm1.Toolbar1.Buttons(17).visible = False
  MDIForm1.Toolbar1.Buttons(18).visible = False
  MDIForm1.Toolbar1.Buttons(19).visible = False
  Exit Sub
  Else
If ty = 1 Or ty = 2 Then 'Call call_theorem(0)
'检查条件是否成立
'输出
Exit Sub
Else
run_type = 10 '进入加辅助点过程
run_statue = 10
t_y = add_aid_point(0, condition_data0)
If t_y = 2 Then
 If error_of_wenti > 0 And error_of_wenti < 100 Then
 GoTo display_run_wenti_error
 Else
 GoTo method_mark1
 End If
ElseIf t_y = 5 Then
GoTo display_run_mark0
Else
display_run_wenti_error:
Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_input_wenti_no + 2) '- display_wenti_v_position%
Wenti_form.Picture1.CurrentX = 20 - display_wenti_h_position%
Wenti_form.Picture1.ForeColor = QBColor(12)
If error_of_wenti = 0 Then '系统不能推出结论
 Wenti_form.Picture1.Print LoadResString_(1680, "")
ElseIf error_of_wenti = 1 Then '
 Wenti_form.Picture1.Print LoadResString_(1705, "")
ElseIf error_of_wenti = 2 Then '结论有错
 Wenti_form.Picture1.Print LoadResString_(1710, "")
ElseIf error_of_wenti = 3 Then
 Wenti_form.Picture1.Print LoadResString_(1715, "") '"问题的条件不充分(即现有条件不能推出结论)!"
ElseIf error_of_wenti = 4 Then
 Wenti_form.Picture1.Print "" 'LoadResString_(1530) '"问题的条件不充分(即现有条件不能推出结论)!"
End If
Call end_prove
run_type = 11
Exit Sub
End If
End If
End If
method_mark1:
If list_type_for_input = input_prove_by_hand Then
 Exit Sub
End If
 If finish_prove = 2 Then
  finish_prove = 3
    ' Call from_last_to_old(1)
 End If
MDIForm1.Inputcond.Enabled = False
MDIForm1.conclusion.Enabled = False
MDIForm1.method.Enabled = False
MDIForm1.method2.Enabled = False
'MDIForm1.method3.Enabled= False
run_statue = 6 '完成证明12.10
display_no = 0
event_statue = get_inform
If run_type = 1 Then
run_type = 2
End If
If is_new_result = False Then
' Call set_condition_tree(conclusion_data(i%).ty, conclusion_data(i%).no(0))
 Call set_display_string_no(0, 0, 0, 0)
 Call arrange_display_no
 Call set_display_string(True, 0, 0, 1, True)
 Call is_sufficient_condition
Else
is_new_result = False
End If
'Wenti_form.Picture1.CurrentY = 20 * (wenti_no + 1)
'Wenti_form.Picture1.CurrentX = 0
Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
display_run_mark5:
'Call C_display_wenti.display_result(Wenti_form.Picture1, False)
Draw_form.DrawMode = 10
run_type = 11 '.Picture1.visible = True
Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, &H8800C6)
      picture_copy = True
If protect_data.pass_word_for_teacher = "00000" Or _
     InStr(1, last_conditions.last_cond(1).pass_word_for_teacher, "*", 0) = 0 Then
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1660, "\\1\\" + str(last_conditions.last_cond(1).total_condition))
Else
'password.Text1.text = "*****"
'password.Show
'If protect_data.pass_word_for_teacher = "000000" Then
'password.Caption = loadresstring_(735)
'Else
'password.Caption = loadresstring_(736)
'End If
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1720, "")
'Call Wenti_form.SetFocus
End If
If finish_prove < 6 Then
  Call complete_prove_
Else
  Call end_prove
End If
'Call Wenti_form.SetFocus
End Sub
Public Sub delete_point_name(ByVal n%, ByVal m%)
Dim i%, j%
For i% = 1 To n% - 1
 For j% = 0 To 50
  If C_display_wenti.m_condition(i%, j%) = _
       C_display_wenti.m_condition(n%, m%) Then
   Exit Sub
  End If
Next j%
Next i%
For i% = n% + 1 To C_display_wenti.m_last_input_wenti_no
For j% = 0 To 50
  If C_display_wenti.m_condition(i%, j%) = _
     C_display_wenti.m_condition(n%, m%) Then
   Exit Sub
  End If
Next j%
Next i%
End Sub

Public Sub save_temp_input()
Dim i%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 Call C_display_wenti.Get_wenti(i%)
 temp_wenti_cond(i%) = wenti_cond0.data
Next i%
End Sub

Public Sub save_input()
Dim i%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 wenti_cond0.data = temp_wenti_cond(i%)
 wenti_cond0.data.wenti_no = i%
 wenti_cond0.is_set_data = True
  Call C_display_wenti.Set_wenti
Next i%
'wenti_no = temp_wenti_no
End Sub

Public Sub display_input_again()
  'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 0, 1, C_display_wenti.m_last_input_wenti_no, 0, False, 0)
   Call save_input
 ' Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, C_display_wenti.m_last_input_wenti_no, 0, False, 0)
End Sub

Public Sub display_char1(ByVal A$, ByVal up_down As Integer)
'显示上下标
If A$ = "^" Or A$ = "!" Then
 Exit Sub
End If
If up_down = up Then
 Printer.CurrentY = Wenti_form.Picture1.CurrentY - 135
  Printer.FontSize = 4       '下标
   Printer.Print A$;
Printer.CurrentY = Printer.CurrentY + 135
  Printer.FontSize = 12       '下标
   ElseIf up_down = down Then
    Printer.CurrentY = Printer.CurrentY + 135
     Printer.FontSize = 4       '下标
      Printer.Print A$;
       Printer.CurrentY = Printer.CurrentY - 135
  Printer.FontSize = 12 '下标
Else
 If Asc(A$) > 0 Then
   Printer.CurrentY = Printer.CurrentY - 45
    Printer.CurrentX = Printer.CurrentX + 45
     Printer.Print A$; '英文
      Printer.CurrentY = Printer.CurrentY + 45
   If A$ = "_" Then
          Printer.CurrentX = Printer.CurrentX + 45
   End If
 Else
     Printer.Print A$; '汉字
 End If
End If

End Sub






Public Sub put_name_to_cond(ByVal n$, ByVal p%)
Dim i%, j%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 For j% = 0 To 10
  If C_display_wenti.m_condition(i%, j%) <> "" And _
     C_display_wenti.m_condition(i%, j%) <> empty_char _
      And C_display_wenti.m_point_no(i%, j%) = p% Then
       Call C_display_wenti.set_m_condition(i%, n$, j%)
  End If
 Next j%
Next i%
End Sub



Public Function cond_to_string(ByVal n%, ByVal l%, ByVal l1%, last_p As Integer) As String
'从输入转化为标准数字串 从字符的第l%到l1%字读出数
Dim k%, i%, j%
Dim s$
 s$ = ""
 k% = l%
 While (Asc(C_display_wenti.m_condition(n%, k%)) <> 13 And l1% >= k%) And _
                       C_display_wenti.m_condition(n%, k%) <> empty_char And _
                        C_display_wenti.m_condition(n%, k%) <> ";"
'从l%开始
  If k% > 1 Then
   If C_display_wenti.m_condition(n%, k% - 1) = "^" And _
       C_display_wenti.m_condition(n%, k% - 2) > "A" Then
'乘方
    i% = val(C_display_wenti.m_condition(n%, k%))
      s$ = Mid$(s$, 1, Len(s$) - 1)
     For j% = 1 To i% - 1
      s$ = s$ + C_display_wenti.m_condition(n%, k% - 2)
     Next j%
     GoTo cond_to_string_mark1
   End If
  End If
'开方
  If C_display_wenti.m_condition(n%, k%) = LoadResString_(1460, "") Then '"'" Then
       s$ = s$ + "'"
  ElseIf C_display_wenti.m_condition(n%, k%) = ":" Then  'LoadResString_(1461,"") Then '"'" Then
       s$ = s$ + "/"
  Else
    If inp% <> 22 And inp% <> 67 Then
      Call C_display_wenti.set_m_condition(n%, _
           LCase(C_display_wenti.m_condition(n%, k%)), k%)
    End If
      s$ = s$ + C_display_wenti.m_condition(n%, k%)
   End If
cond_to_string_mark1:
   k% = k% + 1
 Wend
 If C_display_wenti.m_condition(n%, k%) = empty_char Then
      cond_to_string = s$ + "_"
 Else 'If Asc(c_display_wenti.m_display_string.item(n%).m_condition(k%)) = 13 Then
  last_p = k%
   cond_to_string = s$
 End If
End Function

 
Public Function set_display_inform(ByVal ts$, ty As Byte) As String
Dim i%
Dim ch As String
For i% = 1 To Len(ts$)
 ch = Mid$(ts$, i%, 1)
  If ch = "@" Then
   set_display_inform = set_display_inform + "-"
  ElseIf ch = "#" Then
   set_display_inform = set_display_inform + "+"
  ElseIf ch <> "~" And ch <> "!" Then
   set_display_inform = set_display_inform + ch
  End If
Next i%
If ty <> general_string_ Then
  Exit Function
End If
If set_display_inform <> "" Then
If Mid$(set_display_inform, 1, 1) = "=" Then
    set_display_inform = Mid$(set_display_inform, 2, Len(set_display_inform) - 1)
End If
End If
If set_display_inform <> "" Then
If Mid$(set_display_inform, Len(set_display_inform), 1) = "=" Then
 set_display_inform = set_display_inform + "0"
ElseIf InStr(1, set_display_inform, "=", 0) = 0 Then
 set_display_inform = set_display_inform + "=0"
End If
End If

End Function
Public Function display_information(ByVal ty_no%, button_ty As Integer, Title As String) As Integer
  display_information = MsgBox(display_information_string(ty_no%), button_ty, Title, "", 0)
End Function

Public Sub print_char_with_arrow(ob As Object, color As Byte, v As String)
Dim cu_y%, cu_x%, cu_y1%, cu_x1%, i%
Dim arrow_color As Integer
 If color = 12 Then
    arrow_color = 3
 Else
    arrow_color = 14
 End If
ob.FontSize = 12
cu_y% = ob.CurrentY
cu_x% = ob.CurrentX
ob.Print v;
cu_y1% = ob.CurrentY
cu_x1% = ob.CurrentX
If v = "v" Or v = "u" Then
ob.Line (cu_x%, cu_y% + 6)-(cu_x1%, cu_y% + 6), QBColor(arrow_color)
ob.Line (cu_x1% - 5, cu_y% + 4)-(cu_x1%, cu_y% + 6), QBColor(arrow_color)
Else
ob.Line (cu_x%, cu_y% + 3)-(cu_x1%, cu_y% + 3), QBColor(arrow_color)
ob.Line (cu_x1% - 5, cu_y% + 1)-(cu_x1%, cu_y% + 3), QBColor(arrow_color)
End If
ob.CurrentX = cu_x1%
ob.CurrentY = cu_y1%
End Sub
Public Function read_string_from_string_for_print(ByVal ts As String, s1 As String, S2 As String, _
                                 id_string As String, is_icon As Boolean) As String
Dim i%, j%
is_icon = False
If id_string <> "" Or id_string <> "true" Then
i% = InStr(1, ts, change_L, 0)
If i% > 0 Then
   s1$ = Mid$(ts, 1, i% - 1)
   S2$ = Mid$(ts, i% + 9, Len(ts) - i% - 8)
   id_string = change_L
Else
 i% = InStr(1, ts, "[", 0)
 j% = InStr(1, ts, "_", 0)
 If (i% > 0 And i% < j%) Or j% = 0 Then
  read_string_from_string_for_print = read_string_from_string(1, ts, "[", "]", 0, s1, S2)
 ElseIf j% > 0 Then
  read_string_from_string_for_print = read_string_from_string(1, ts, "_", "~", 0, s1, S2)
  S2 = "~" + S2
  'S2 = "_"
 End If
 '读出"[]"
If read_string_from_string_for_print <> "" Then
 If right(read_string_from_string_for_print, 3) = "\\]" And left(read_string_from_string_for_print, 3) = "[\\" Then '[\\0\\]
   id_string = Mid$(read_string_from_string_for_print, 2, Len(read_string_from_string_for_print) - 2) '\\0\\
    If S2 <> "" Then
       If left(S2, 1) = "_" Then
        read_string_from_string_for_print = "[_]"
        S2 = Mid$(S2, 2, Len(S2) - 1)
        is_icon = True
       Else
        read_string_from_string_for_print = "[" + global_icon_char + "]" '[□]
        is_icon = True
       End If
    Else
        read_string_from_string_for_print = "[" + global_icon_char + "]"
        is_icon = True
    End If
 ElseIf read_string_from_string_for_print = "_~" Then
            read_string_from_string_for_print = "[_]"
        S2 = Mid$(S2, 2, Len(S2) - 1)
        is_icon = True
 End If
Else
'*******************************
 i% = InStr(2, ts, " ", 0)
 j% = InStr(2, ts, ",", 0)
 If i% = 0 And j% = 0 Then
 ElseIf i% = 0 Then
 i% = j%
 ElseIf j% = 0 Then
 Else
 i% = min(i%, j%)
 End If
  j% = InStr(1, ts, "、", 0)
 If i% = 0 And j% = 0 Then
 ElseIf i% = 0 Then
 i% = j%
 ElseIf j% = 0 Then
 Else
 i% = min(i%, j%)
 End If
'*********************************************
 id_string = ""
 If i% > 0 And i% < Len(ts) + 1 Then
    s1 = Mid$(ts, 1, i%)
    If i% < Len(ts) Then
       S2 = Mid$(ts, i% + 1, Len(ts) - i%)
    Else
       S2 = ""
    End If
 ElseIf i% = 0 Then
       s1 = ts
       S2 = ""
 End If
End If
End If
End If
End Function

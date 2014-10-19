Attribute VB_Name = "IO"
Option Explicit
Global io_statue As Integer
Global wenti_data_type As Byte
Global Const new_ty = 1
Global Const old_ty = 0
Type problem_record
   wenti_no As Integer '问题语句数
    last_number_string As Byte
    input_last_point As Byte '
     number_string(1 To 8) As String
      wenti_cond(1 To 30) As wentitype
        poi(1 To 26) As io_point_data_type
         lin(1 To 20) As io_line_data_type
           circ(1 To 10) As io_circle_data_type
        last_line As Integer
         last_point As Integer '
          last_circle As Integer
End Type
'********************************************************
Type io_record_type
 name As String * 20 '例题的名
  pro_text As String * 256 '内容
   problem As problem_record
End Type
Type file_record_type
mark  As String * 8
record As io_record_type
End Type
'********************************************
Type problem_record_new
   wenti_no As Integer '问题语句数
    last_number_string As Byte
    input_last_point As Byte '
     number_string(1 To 8) As String
      wenti_cond(50) As wentitype
        '40 参数 41-45 直线 46-47 48-47
        poi(1 To 100) As point_data0_type
         lin(1 To 100) As line_data0_type
           circ(1 To 20) As circle_data_type
           Con_lin(1 To 20) As line_data0_type
       last_con_line As Integer
        last_line As Integer
         last_point As Integer '
          last_circle As Integer
End Type
Type io_record_type_new
 name As String * 20 '例题的名
  pro_text As String * 256 '内容
problem As problem_record
End Type
Type file_record_type_new
mark  As String * 8
record As io_record_type
End Type
'******************************
Global file_record As file_record_type
Global wenti_record As io_record_type
'Global wenti_record_0 As io_record_type_0
Global wenti_record0 As io_record_type
'Global file_record_new As file_record_type_new
Global wenti_record_new As io_record_type_new '初始化
Global wenti_record0_new As io_record_type
Global temp_record_poi(1 To 26, 1) As Integer
Global temp_problem(0 To 100) As problem_record
Global temp_problem_new(1 To 100) As problem_record_new
Global last_problem_input%
Public Sub put_wenti_to_record(ByVal n$) '
wenti_record = wenti_record0 '初始化
If io_statue = 1 Then
If wenti_type = 1 Then
wenti_record.name = IO_form.Text1.text + "~"
'用以记录选择题
Else
wenti_record.name = IO_form.Text1.text
End If
End If
wenti_record.problem = put_wenti_to_problem
End Sub
Public Sub get_wenti_from_record(wenti_record As io_record_type)
Dim i%
Dim t_name$
  For i% = 1 To Len(wenti_record.name) '读例题名
   If Mid$(wenti_record.name, i%, 1) <> empty_char Then
   Call change_example_name(wenti_record.name)
    t_name$ = t_name$ + Mid$(wenti_record.name, i%, 1)
   End If
   Next i%
   t_name$ = Trim(t_name$)
    wenti_form_title = LoadResString_(1960, "") + "-" + t_name$
    Draw_form.Caption = LoadResString_(2005, "") + "-" + t_name$
    Wenti_form.Caption = wenti_form_title & LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
MDIForm1.Inputcond.Enabled = False
MDIForm1.conclusion.Enabled = False
 If Mid$(wenti_record.name, Len(wenti_record.name), 1) = "~" Then '问题类型
  wenti_type = 1
Else
 wenti_type = 0
End If
Call get_wenti_from_problem(wenti_record.problem) '
     exam_form.Hide
      Draw_form.Show
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 Call set_initial_condition(i%, 0, True)
Next i%
 set_or_prove = 1
 draw_or_prove = 0
End Sub
Public Function put_wenti_to_problem() As problem_record
Dim i%, j%
'*********************************************************************
'记录问题的个数put_wenti_to_problem.wenti_no
If run_statue = 6 Then '12,10
put_wenti_to_problem.wenti_no = C_display_wenti.m_last_conclusion
Else
put_wenti_to_problem.wenti_no = C_display_wenti.m_last_input_wenti_no
End If
'**********************************************************************
For i% = 1 To put_wenti_to_problem.wenti_no '问题条件号
If C_display_wenti.m_no(i%) = -1 Then
 Call C_display_wenti.set_m_condition(i%, empty_char, 4) 'XX=XX的结束符
End If
'***********************************************************************
For j% = 1 To 4
 put_wenti_to_problem.wenti_cond(i%).line_no(j%) = _
       C_display_wenti.m_inner_lin(i%, j%)
 put_wenti_to_problem.wenti_cond(i%).poi(j%) = _
       C_display_wenti.m_inner_poi(i%, j%)
 put_wenti_to_problem.wenti_cond(i%).circ(j%) = _
       C_display_wenti.m_inner_circ(i%, j%)
Next j
For j% = 5 To 8
 put_wenti_to_problem.wenti_cond(i%).poi(j%) = _
       C_display_wenti.m_inner_poi(i%, j%)
Next j
 put_wenti_to_problem.wenti_cond(i%).inter_set_point_type = _
       C_display_wenti.m_inner_point_type(i%)

'将记录输入
For j% = 0 To 50
  put_wenti_to_problem.wenti_cond(i%).condition(j%) = _
          C_display_wenti.m_condition(i%, j%)
  put_wenti_to_problem.wenti_cond(i%).point_no(j%) = _
          C_display_wenti.m_point_no(i%, j%)
Next j%
  put_wenti_to_problem.wenti_cond(i%).no = _
          C_display_wenti.m_no(i%)
Next i%
'*********************************************************************
put_wenti_to_problem.last_point = last_conditions.last_cond(1).point_no
For i% = 1 To last_conditions.last_cond(1).point_no
 put_wenti_to_problem.poi(i%).point_data = m_poi(i%).data(0).data0 '.coordinate.X
 put_wenti_to_problem.poi(i%).condition = m_poi(i%).data(0).condition_ty
 put_wenti_to_problem.poi(i%).point_no = i%
 put_wenti_to_problem.poi(i%).depend_elemant = m_poi(i%).data(0).depend_element
Next i%
put_wenti_to_problem.last_line = 0
For i% = 1 To last_conditions.last_cond(1).line_no
 If m_lin(i%).data(0).data0.visible > 0 Then
   put_wenti_to_problem.last_line = put_wenti_to_problem.last_line + 1
    If put_wenti_to_problem.last_line <= 20 Then '为了不越界,限定
     put_wenti_to_problem.line_no(put_wenti_to_problem.last_line).line_no = i%
     put_wenti_to_problem.line_no(put_wenti_to_problem.last_line).depend_element = _
         m_lin(i%).data(0).depend_element
     put_wenti_to_problem.line_no(put_wenti_to_problem.last_line).condition = _
         m_lin(i%).data(0).condition
     put_wenti_to_problem.line_no(put_wenti_to_problem.last_line).line_data = _
         m_lin(i%).data(0).data0
    End If
  End If
Next i%
put_wenti_to_problem.last_circle = 0
For i% = 1 To C_display_picture.m_circle.Count
   put_wenti_to_problem.last_circle = put_wenti_to_problem.last_circle + 1
   put_wenti_to_problem.circ(put_wenti_to_problem.last_circle).circle_no = i%
   put_wenti_to_problem.circ(put_wenti_to_problem.last_circle).input_ty = m_Circ(i%).data(0).input_type
   put_wenti_to_problem.circ(put_wenti_to_problem.last_circle).circle_ty = m_Circ(i%).data(0).circle_type
   put_wenti_to_problem.circ(put_wenti_to_problem.last_circle).circle_data = m_Circ(i%).data(0).data0
Next i%
'For i% = 1 To last_conditions.last_cond(1).con_line_no
' put_wenti_to_problem.Con_lin(i%) = from_wenti_line_to_problem_line(m_Con_lin(i%))
'Next i%
End Function

Public Sub get_wenti_from_problem(pr As problem_record)
Dim i%, j%
Dim tc As circle_data_type
Dim l_d As line_data_type
Dim t_ld As line_data_type
'wenti_no = pr.wenti_no
input_last_point = pr.input_last_point
For i% = 1 To pr.last_number_string
    Call set_number_string(pr.number_string(i%), i%)
Next i%
last_conditions.last_cond(1).point_no = 0
For i% = 1 To pr.last_point
 Call set_point_data_from_input(pr.poi(i%))
Next i%
For i% = 1 To pr.last_line
Call set_line_data_from_input(pr.line_no(i%))
Next i%
For i% = 1 To pr.last_circle
Call set_circle_data_from_input(pr.circ(i%))
Next i%
For i% = 1 To pr.wenti_no
wenti_cond0.data = pr.wenti_cond(i%)
wenti_cond0.data.wenti_no = i%
wenti_cond0.is_set_data = True
 Call C_display_wenti.Set_wenti
Next i%
End Sub


Public Function from_line_to_string(l_data As line_data_type) As String
Dim temp_string
Dim i%
  from_line_to_string = ""
   from_line_to_string = Trim(str(l_data.data0.total_color))
  For i% = 0 To 10
   temp_string = value_to_string(l_data.in_point(i%), 3)
  from_line_to_string = from_line_to_string + temp_string
 Next i%
   temp_string = value_to_string(l_data.data0.type, 3)
  from_line_to_string = from_line_to_string + temp_string
     temp_string = value_to_string(l_data.degree, 3)
  from_line_to_string = from_line_to_string + temp_string
     temp_string = value_to_string(l_data.parent.element(0).no, 3)
  from_line_to_string = from_line_to_string + temp_string
     temp_string = value_to_string(l_data.parent.element(1).no, 3)
  from_line_to_string = from_line_to_string + temp_string
  from_line_to_string = from_line_to_string + Trim(str(l_data.data0.visible))
End Function

Public Function from_string_to_line(s As String) As line_data_type
Dim i%
If wenti_data_type <= 1 Then
from_string_to_line.data0.total_color = val(Mid$(s, 1, 1))
For i% = 0 To 10
from_string_to_line.in_point(i%) = string_to_value(Mid$(s, 2 + i * 3, 3))
from_string_to_line.data0.in_point(i%) = Abs(from_string_to_line.in_point(i%))
Next i%
from_string_to_line.data0.poi(0) = Abs(from_string_to_line.data0.in_point(1))
from_string_to_line.data0.poi(1) = Abs(from_string_to_line.data0.in_point( _
    from_string_to_line.data0.in_point(0)))
from_string_to_line.in_point(from_string_to_line.data0.in_point(0)) = _
    from_string_to_line.data0.in_point(from_string_to_line.data0.in_point(0))
from_string_to_line.data0.type = string_to_value(Mid$(s, 35, 3))
from_string_to_line.data0.visible = val(Mid$(s, 38, 1))
Else
from_string_to_line.data0.total_color = val(Mid$(s, 1, 1))
For i% = 0 To 10
from_string_to_line.in_point(i%) = string_to_value(Mid$(s, 2 + i * 3, 3))
from_string_to_line.data0.in_point(i%) = Abs(from_string_to_line.in_point(i%))
Next i%
from_string_to_line.data0.poi(0) = Abs(from_string_to_line.data0.in_point(1))
from_string_to_line.data0.poi(1) = Abs(from_string_to_line.data0.in_point( _
    from_string_to_line.data0.in_point(0)))
from_string_to_line.in_point(from_string_to_line.data0.in_point(0)) = _
    from_string_to_line.data0.in_point(from_string_to_line.data0.in_point(0))
from_string_to_line.data0.type = string_to_value(Mid$(s, 35, 3))
from_string_to_line.degree = string_to_value(Mid$(s, 38, 3))
from_string_to_line.parent.element(0).no = string_to_value(Mid$(s, 41, 3))
from_string_to_line.parent.element(1).no = string_to_value(Mid$(s, 44, 3))
from_string_to_line.data0.visible = val(Mid$(s, 47, 1))
End If
'*****************************
End Function

Public Function DoesFileExist(ByVal FileName As String) As Boolean
Dim Buffer As OFSTRUCT
Dim isThere As Integer
isThere = OpenFile(FileName, Buffer, &O4000)
If isThere = -1 Then
 DoesFileExist = False
Else
 DoesFileExist = True
End If
End Function

Public Function from_string_to_circle(s As String) As circle_data_type
Dim i%
Dim t_string
from_string_to_circle.data0.c_coord.X = string_to_value(Mid$(s, 1, 4))
from_string_to_circle.data0.c_coord.Y = string_to_value(Mid$(s, 5, 4))
from_string_to_circle.data0.center = string_to_value(Mid$(s, 9, 3))
from_string_to_circle.data0.color = string_to_value(Mid$(s, 12, 3))
For i% = 0 To 10
from_string_to_circle.data0.in_point(i%) = string_to_value(Mid$(s, 15 + i% * 3, 3))
Next i%
from_string_to_circle.data0.visible = val(Mid$(s, 48, 1))
If from_string_to_circle.data0.center > 0 Then
from_string_to_circle.data0.radii = _
   sqr((m_poi(from_string_to_circle.data0.in_point(1)).data(0).data0.coordinate.X - _
         m_poi(from_string_to_circle.data0.center).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(from_string_to_circle.data0.in_point(1)).data(0).data0.coordinate.Y - _
           m_poi(from_string_to_circle.data0.center).data(0).data0.coordinate.Y) ^ 2)

Else
from_string_to_circle.data0.radii = _
      sqr((m_poi(from_string_to_circle.data0.in_point(1)).data(0).data0.coordinate.X - _
        from_string_to_circle.data0.c_coord.X) ^ 2 + _
         (m_poi(from_string_to_circle.data0.in_point(1)).data(0).data0.coordinate.Y - _
           from_string_to_circle.data0.c_coord.Y) ^ 2)
End If
If wenti_data_type = 2 Then
from_string_to_circle.degree = string_to_value(Mid$(49, 3))
from_string_to_circle.parent.element(0).no = string_to_value(Mid$(52, 3))
from_string_to_circle.parent.element(1).no = string_to_value(Mid$(55, 3))
from_string_to_circle.parent.element(2).no = string_to_value(Mid$(58, 3))
End If
'**************
End Function

Public Function from_string_to_wenti(ByVal s As String) As wentitype
Dim i%, no%
Dim ch$
from_string_to_wenti.no = string_to_value(Mid$(s, 1, 3))
from_string_to_wenti.no_ = string_to_value(Mid$(s, 4, 3))
For i% = 0 To 50
          ch$ = Mid$(s, 7 + i% * 6, 1)
If ch$ = ";" Then
  ch$ = empty_char
ElseIf ch$ = "%" Then
  ch$ = Chr(13)
Else
  Call next_char(0, ch$, 0, 0)
End If
  from_string_to_wenti.condition(i%) = ch$
  from_string_to_wenti.point_no(i%) = string_to_value(Mid$(s, 8 + i% * 6, 5))
Next i%
For i% = 1 To 8
  from_string_to_wenti.poi(i%) = string_to_value(Mid$(s, 302 + i% * 5, 5))
Next i%
For i% = 1 To 4
  from_string_to_wenti.line_no(i%) = string_to_value(Mid$(s, 342 + i% * 5, 5))
Next i%
For i% = 1 To 2
  from_string_to_wenti.circ(i%) = string_to_value(Mid$(s, 362 + i% * 5, 5))
Next i%
  from_string_to_wenti.inter_set_point_type = string_to_value(Mid$(s, 387, 5))
End Function

Public Function from_point_to_string(p As point_data_type) As String
Dim temp_string
temp_string = value_to_string(p.data0.color, 3)
from_point_to_string = temp_string
temp_string = value_to_string(p.data0.coordinate.X, 4)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.data0.coordinate.Y, 4)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.degree, 2)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.degree_for_reduce, 2)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.parent.element(0).no, 3)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.parent.element(1).no, 3)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.parent.element(0).no, 3)
from_point_to_string = from_point_to_string + temp_string
temp_string = value_to_string(p.parent.element(1).no, 3)
from_point_to_string = from_point_to_string + temp_string
If p.data0.name <> empty_char Then
If Len(p.data0.name) = 2 Then
from_point_to_string = from_point_to_string + p.data0.name '两个字母的点名
Else
from_point_to_string = from_point_to_string + p.data0.name
from_point_to_string = from_point_to_string + "0"
End If
Else
from_point_to_string = from_point_to_string + "00"
End If
from_point_to_string = from_point_to_string + Trim(str(p.data0.visible))
End Function

Public Function from_string_to_point(s As String) As point_data_type
'ty=1 ty=2 ,二版
If wenti_data_type <= 1 Then
 from_string_to_point.data0.color = string_to_value(Mid$(s, 1, 3))
 from_string_to_point.data0.coordinate.X = string_to_value(Mid$(s, 4, 4))
 from_string_to_point.data0.coordinate.Y = string_to_value(Mid$(s, 8, 4))
 from_string_to_point.data0.name = Mid$(s, 12, 1)
  If from_string_to_point.data0.name = "0" Then
   from_string_to_point.data0.name = empty_char
  End If
 from_string_to_point.data0.visible = val(Mid$(s, 13, 1))
ElseIf wenti_data_type = 2 Then
 from_string_to_point.data0.color = string_to_value(Mid$(s, 1, 3))
 from_string_to_point.data0.coordinate.X = string_to_value(Mid$(s, 4, 4))
 from_string_to_point.data0.coordinate.Y = string_to_value(Mid$(s, 8, 4))
 from_string_to_point.degree = string_to_value(Mid$(s, 12, 2))
 from_string_to_point.degree_for_reduce = string_to_value(Mid$(s, 14, 2))
 from_string_to_point.parent.element(0).no = string_to_value(Mid$(s, 16, 3))
 from_string_to_point.parent.element(1).no = string_to_value(Mid$(s, 19, 3))
 'from_string_to_point.depend_cir(0) = string_to_value(Mid$(s, 22, 3))
 'from_string_to_point.depend_cir(1) = string_to_value(Mid$(s, 25, 3))
 from_string_to_point.data0.name = Mid$(s, 28, 2)
  If from_string_to_point.data0.name = "00" Then
   from_string_to_point.data0.name = empty_char
  ElseIf Mid$(from_string_to_point.data0.name, 2, 1) = "0" Then
   from_string_to_point.data0.name = Mid$(from_string_to_point.data0.name, 1, 1)
  End If
 from_string_to_point.data0.visible = val(Mid$(s, 30, 1))
End If
End Function
Public Function input_problem_from_file(file_name As String) As Byte
'从文件输入问题
Dim temp_string As String
Dim i%, n%, no%, no_%
Dim l_data0 As line_data_type
Dim t_last_conclusion_no%
Dim T_point_no%, t_line_no%, T_circle_no%, T_con_line_no%
'On Error GoTo input_problem_from_file_mark0
 Open file_name For Input As #2 '打开文件
  GoTo input_problem_from_file_mark1
input_problem_from_file_mark0:
   If MsgBox(Error$(Err), 2, "", "", 0) = 3 Then
    Close #2
    Exit Function
   End If
input_problem_from_file_mark1:
 Input #2, temp_string '输入文件
 If temp_string <> "DSH_PMJH" And temp_string <> "SHD_PMJH" And _
     temp_string <> "SHD_PMJH_V2" And temp_string <> "SHD_VMJH_V2" Then
   If MsgBox(LoadResString_(2105, ""), 4, "", "", 0) = 6 Then
    Close #2
     Exit Function
   End If
 ElseIf temp_string = "SHD_PMJH" Then '新文件文件长输入
  wenti_data_type = 1
   regist_data.run_type = 0
  ElseIf temp_string = "DSH_PMJH" Then
   wenti_data_type = 0
   regist_data.run_type = 0
 ElseIf temp_string = "SHD_PMJH_V2" Then
   regist_data.run_type = 0
   wenti_data_type = 2
 ElseIf temp_string = "SHD_VMJH_V2" Then
   wenti_data_type = 2
   regist_data.run_type = 1
 End If
 Do While Not EOF(2)
 Input #2, temp_string
 If n% = 0 Then
 '输入各记录长
 t_last_conclusion_no% = string_to_value(Mid$(temp_string, 1, 3))
 T_point_no% = string_to_value(Mid$(temp_string, 7, 3))
 t_line_no% = string_to_value(Mid$(temp_string, 10, 3))
 T_circle_no% = string_to_value(Mid$(temp_string, 13, 3))
 T_con_line_no% = string_to_value(Mid$(temp_string, 16, 3))
 If t_last_conclusion_no% = 0 Then
    t_last_conclusion_no% = string_to_value(Mid$(temp_string, 4, 3))
 End If
 ElseIf n% > 0 And n% <= t_last_conclusion_no% Then
 '输入已知
  wenti_cond0.data = from_string_to_wenti(temp_string)
  wenti_cond0.data.wenti_no = n%
  wenti_cond0.is_set_data = True
   Call C_display_wenti.Set_wenti
 ElseIf n% <= t_last_conclusion_no% + T_point_no% Then ' And n% <= old_wenti_no + last_conditions.last_cond(1).point_no Then
 '点记录
 Call set_point_data(n% - t_last_conclusion_no%, _
                                       from_string_to_point(temp_string), True)
 ElseIf n% <= t_last_conclusion_no% + T_point_no% + _
           t_line_no% Then '+ last_conditions.last_cond(1).point_no And n% <= old_wenti_no + last_conditions.last_cond(1).point_no + last_conditions.last_cond(1).line_no Then
 '线记录
 no% = n% - t_last_conclusion_no% - T_point_no%
 l_data0 = from_string_to_line(temp_string)
  Call set_line_data0(no%, l_data0, 0, 0) '显示状态由输入条件决定
 ElseIf n% <= t_last_conclusion_no% + T_point_no% + _
                                 t_line_no% + T_circle_no% Then
 '圆记录
 no% = n% - t_last_conclusion_no% - T_point_no% - t_line_no%
  m_input_circle_data0 = from_string_to_circle(temp_string)
  m_input_circle_data0.input_type = condition
  Call Set_m_circle_data(no%, m_input_circle_data0)
 ElseIf n% <= t_last_conclusion_no% + T_point_no% + _
                t_line_no% + T_circle_no% + T_con_line_no% Then
 '结论线记录
 no% = n% - t_last_conclusion_no% - T_point_no% - t_line_no% - T_circle_no%
 'Call set_con_line_data0(no%, from_string_to_line(temp_string), True)
 ElseIf temp_string <> "" Then
 '解题记录
'  Call C_display_wenti.m_display_string.item(no_%).set_m_ty(val(Mid$(temp_string, 1, 1)))
  no_% = no_% + 1
  Call C_display_wenti.set_m_conclusion_or_condition(no_%, val(Mid$(temp_string, 2, 1)))
  Call C_display_wenti.set_m_theorem_no(no_%, string_to_value(Mid$(temp_string, 3, 3)))
  Call C_display_wenti.set_m_string(no_%, _
             "", Mid$(temp_string, 6, Len(temp_string) - 5), "", "", "", 0, 1, 0)
 End If
 n% = n% + 1
Loop
Close #2
'*********************************************************
wenti_no_% = C_display_wenti.m_last_input_wenti_no
'初始化问题条件
  wenti_form_title = LoadResString_(1960, "") + "-" & path_and_file
   Draw_form.Caption = LoadResString_(2005, "") & "-" & path_and_file
    Wenti_form.Caption = wenti_form_title + LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
    For i% = 1 To C_display_wenti.m_last_input_wenti_no
      'Call C_display_wenti.m_display_string.item(i%).cond_to_display(1, -1)
    Next i%
 'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, C_display_wenti.m_last_input_wenti_no, 0, False, 0)
       Call draw_again1(Draw_form)
      Call set_name_for_draw_picture0
For i% = 1 To t_last_conclusion_no%
 Call set_initial_condition(i%, 0, False)
Next i%
 set_or_prove = 1
 draw_or_prove = 0
If C_display_wenti.m_last_input_wenti_no > t_last_conclusion_no% And _
                 t_last_conclusion_no% > 0 Then
  run_statue = 6 '12.10
End If
If wenti_data_type = 0 Then
wenti_data_type = 1
End If
End Function

Public Function value_to_string(v As Variant, n%) As String
Dim i%
 value_to_string = Trim(str(v))
For i% = Len(value_to_string) + 1 To n%
  value_to_string = "0" + value_to_string
Next i%
End Function

Public Function string_to_value(ByVal s As String) As Variant
Dim i%, l%
l% = Len(s)
For i% = 1 To l%
 If Mid$(s, i%, 1) <> "0" Then
  s = Mid$(s, i%, l% - i% + 1)
   string_to_value = val(s)
    Exit Function
 End If
Next i%
End Function

Public Function from_sqr_no_to_string(ByVal s As String) As String
Dim n1%, n2%, i%, j%
Dim ch As String * 1
Dim v$
n1% = InStr(1, s, "!", 0)
n2% = InStr(1, s, "~", 0)

If n1% < n2% - 1 Then
 from_sqr_no_to_string = Mid$(s, 1, n1%)
 For i% = n1% + 1 To n2% - 1
  ch = Mid$(s, i%, 1)
  If ch = "[" Then
   v$ = read_sqr_no_from_string(s, i%, i%, "")
   from_sqr_no_to_string = from_sqr_no_to_string + "[" + _
      v$ + "]"
  Else
   from_sqr_no_to_string = from_sqr_no_to_string + ch
  End If
 Next i%
from_sqr_no_to_string = from_sqr_no_to_string + _
     from_sqr_no_to_string(Mid$(s, n2%, Len(s) - n2% + 1))
Else
  from_sqr_no_to_string = s
End If
End Function

Public Function from_sqr_string_to_no(ByVal s As String) As String
Dim n1%, n2%
Call get_squre_brace_pair(s, n1%, n2%)
If n1% > 0 And n2% > 0 Then
from_sqr_string_to_no = Mid$(s, 1, n1 - 1)
from_sqr_string_to_no = from_sqr_string_to_no + _
     set_squre_root_string_(Mid$(s, n1% + 1, n2% - n1% - 1)) + _
         from_sqr_string_to_no(Mid$(s, n2% + 1, Len(s) - n2%))
Else
from_sqr_string_to_no = s
End If
End Function
Public Sub get_squre_brace_pair(ByVal s As String, n1%, n2%)
Dim i%, l%
Dim ch As String * 1
n1% = InStr(1, s, "[", 0)
If n1% = 0 Then
 n2% = 0
  Exit Sub
Else
 l% = 1
 For i% = n1% + 1 To Len(s)
  ch = Mid$(s, i%, 1)
  If ch = "]" Then
   l% = l% - 1
   If l% = 0 Then
    n2% = i%
     Exit Sub
   End If
  ElseIf ch = "[" Then
   l% = l% + 1
  End If
 Next i%
End If
End Sub

Public Sub input_wenti_from_problem(pr As problem_record)
Dim i%
Call clear_wenti_display
Call init_conditions(0)
Call get_wenti_from_problem(pr)
   For i% = 1 To C_display_wenti.m_last_input_wenti_no
      'Call C_display_wenti.m_display_string.item(i%).cond_to_display(1)
   Next i%
   'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, _
          C_display_wenti.m_last_input_wenti_no, 0, False, 0)
    'Call display_wenti
     exam_form.Hide
      Draw_form.Show
       Call draw_again1(Draw_form)
    Call set_name_for_draw_picture0
'For i% = 1 To last_conditions.last_cond(1).line_no
 'If lin(i%).data(0).other_no = 0 Then
 '   lin(i%).data(0).other_no = i%
 'End If
'Next i%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 Call set_initial_condition(i%, 0, False)
Next i%
 set_or_prove = 1
 draw_or_prove = 0

End Sub
'Public Sub from_old_type_wenti_to_new()
'Dim i%, j%
'If wenti_data_type = 0 Then
' wenti_record.name = wenti_record_0.name
' wenti_record.pro_text = wenti_record_0.pro_text
' For i% = 1 To 10
' wenti_record.problem.circ(i%) = wenti_record_0.problem.circ(i%)
' Next i%
' For i% = 1 To 26
' wenti_record.problem.poi(i%) = wenti_record_0.problem.poi(i%)
' Next i%
' For i% = 1 To 6
' wenti_record.problem.Con_lin(i%) = wenti_record_0.problem.Con_lin(i%)
' Next i%
' wenti_record.problem.input_last_point = wenti_record_0.problem.input_last_point
' wenti_record.problem.last_circle = wenti_record_0.problem.last_circle
' wenti_record.problem.last_con_line = wenti_record_0.problem.last_con_line
' wenti_record.problem.last_line = wenti_record_0.problem.last_line
' wenti_record.problem.last_point = wenti_record_0.problem.last_point
' For i% = 1 To 20
' wenti_record.problem.line_no(i%) = wenti_record_0.problem.line_no(i%)
' Next i%
' For j% = 1 To 30
' wenti_record.problem.wenti_cond(j%).no = wenti_record_0.problem.wenti_cond(j%).no
' For i% = 0 To 19
' wenti_record.problem.wenti_cond(j%).condition(i%) = wenti_record_0.problem.wenti_cond(j%).condition(i%)
' wenti_record.problem.wenti_cond(j%).point_no(i%) = wenti_record_0.problem.wenti_cond(j%).point_no(i%)
' Next i%
' For i% = 20 To 50
' wenti_record.problem.wenti_cond(j%).condition(i%) = empty_char
' wenti_record.problem.wenti_cond(j%).point_no(i%) = 0
' Next i%
' Next j%
' wenti_record.problem.wenti_no = wenti_record_0.problem.wenti_no
' If wenti_data_type = 0 Then
'  wenti_data_type = 1
' End If
'End If

'End Sub

Public Function from_input_line_data0_to_line_data(l_data0 As line_data0_type) As line_data_type
Dim i%
from_input_line_data0_to_line_data.data0 = l_data0
For i% = 0 To 10
from_input_line_data0_to_line_data.in_point(i%) = _
 from_input_line_data0_to_line_data.data0.in_point(i%)
 from_input_line_data0_to_line_data.data0.in_point(i%) = _
 Abs(from_input_line_data0_to_line_data.data0.in_point(i%))
Next i%
End Function

Private Function from_wenti_line_to_problem_line(l_d As line_data_type) As line_data0_type
Dim i%
from_wenti_line_to_problem_line = l_d.data0
For i% = 0 To l_d.data0.in_point(0)
from_wenti_line_to_problem_line.in_point(i%) = l_d.data0.in_point(i%)
Next i%
from_wenti_line_to_problem_line.in_point(10) = l_d.data0.in_point(10)
End Function
Private Function from_problem_line_to_wenti_line(l_d As line_data0_type) As line_data_type
Dim i%
For i% = 0 To l_d.in_point(0)
from_problem_line_to_wenti_line.in_point(i%) = l_d.in_point(i)
from_problem_line_to_wenti_line.data0.in_point(i%) = Abs(l_d.in_point(i))
Next i%
from_problem_line_to_wenti_line.in_point(l_d.in_point(0)) = _
         from_problem_line_to_wenti_line.data0.in_point(l_d.in_point(0))
from_problem_line_to_wenti_line.in_point(10) = l_d.in_point(10)
from_problem_line_to_wenti_line.data0.in_point(10) = Abs(l_d.in_point(10))
from_problem_line_to_wenti_line.data0.total_color = l_d.total_color
from_problem_line_to_wenti_line.data0.poi(0) = l_d.poi(0)
from_problem_line_to_wenti_line.data0.poi(1) = l_d.poi(1)
from_problem_line_to_wenti_line.data0.type = l_d.type
from_problem_line_to_wenti_line.data0.visible = l_d.visible
End Function


Public Sub change_example_name(t_wenti_name As String)
If regist_data.language = 1 Then
 If replace_string_by_string(t_wenti_name, "example", "例 ") = False Then '中文显示
    If replace_string_by_string(t_wenti_name, "Example", "例 ") = False Then
       Call replace_string_by_string(t_wenti_name, "例", "例 ")
    End If
 End If
ElseIf regist_data.language = 2 Then '英文显示
  If replace_string_by_string(t_wenti_name, "example", "Example ") = False Then
    If replace_string_by_string(t_wenti_name, "例", "Example ") = False Then
        Call replace_string_by_string(t_wenti_name, "例", "Example ")
    End If
  End If
Else '其他留作以后用
 If replace_string_by_string(t_wenti_name, "example", "例 ") = False Then
    If replace_string_by_string(t_wenti_name, "Example", "例 ") = False Then
       Call replace_string_by_string(t_wenti_name, "例", "例 ")
    End If
 End If
End If

End Sub


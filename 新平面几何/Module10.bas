Attribute VB_Name = "MODULE10"
Option Explicit
Global last_char As Integer
Global inpcond(-100 To 100) As inpcond0_type '供选用的输入语句
'Global inpcond0 As String
Type wentitype '???????
 wenti_no As Integer
 no As Integer
 no_ As Integer
 condition(50) As String * 1  '记录输入条件
  point_no(50)    As Integer  '记录输入点号
   inter_set_point_type As Integer '
    line_no(1 To 4) As Integer
     circ(1 To 4) As Integer
      poi(1 To 8) As Integer
End Type
Type display_input_cond
no As Integer
ty As Byte
'record As total_record_type
 condition_ty As Byte
 condition_no As Integer
 theorem_no As Integer
cond As String
conclusion_or_condition As Byte
End Type
Global list_type_for_input As Integer
'*******************************************
Type IO_pointtype
data As point_data0_type
is_set_data As Boolean
point_no As Integer
End Type
Type IO_linetype
data As line_data0_type
is_set_data As Boolean
line_no As Integer
End Type
Type IO_circletype
data As circle_data0_type
is_set_data As Boolean
circle_no As Integer
End Type
Type IO_wentitype
data As wentitype
is_set_data As Boolean
End Type
Global PointData0  As IO_pointtype '
Global LineData0  As IO_linetype '
Global CircleData0  As IO_circletype '
Global wenti_cond0  As IO_wentitype '
'**************************************************
Type aid_display_string_type
display_string As String
record_type As Byte
record_no As Integer
new_point As Integer
'display_no As Integer
record_ As record_type
End Type
'Global aid_display_string(16) As aid_display_string_type
'Global display_input_condition() As display_input_cond 'String '　输入显示
'Global display_input_condition1() As display_input_cond
'Global display_input_condition0() As display_input_cond
'*************************************
Global temp_wenti_cond(500) As wentitype '记录变化的输入
'Global c_display_wenti.m_display_string.item(500) As wentitype '输入问题
Global wenti_cond_no_reduce(100) As Boolean  '输入问题
'Global temp_wenti_no As Integer
Global wenti_condition_no As Integer '问题未解答前的条件数
Global draw_wenti_no As Integer
Global input_last_point As Byte
'记录已输入但尚未画的点的末号
'***************************************************
Global char_x As Single
Global char_y As Single
'Global icon_display As Integer  'icon是否显示
Global input_type As Integer  '区别是输入条件还是结论
'***********************************************
Global input_condition_no As Integer
Global operator As String '　画图操作
Global mouse_up_no_enabled As Boolean
Global draw_step As Integer
Global modify_statue  As Boolean
Global modify_wenti_no As Integer
Global last_condition As Integer
Global first_part$, second_part$, last_part$
Global char_position As Integer
Global write_wenti_no As Integer
Global input_statue_from_p As Integer
Global event_statue As Integer
Global temp_tangent_circle_no(1) As Integer
'记录输入状态
'Global draw_statue As Integer
Global up_down As Integer
Global Const up = 1
Global Const arrow = 2
Global Const middle = 0
Global Const down = -1
Global Const input_condition_statue = 0 '&H1
Global Const exit_program = -10
Global Const initial_condition = -1
Global Const ready = 0
Global Const mouse_is_moving = 1
Global Const wait_for_input_condition = 2
Global Const wait_for_draw_picture = 3
Global Const input_for_conclusion = 4 '&H2
Global Const complete_input = 5
Global Const proving = 6
Global Const input_add_point = 7
Global Const input_prove_by_hand = 8
Global Const set_measur = 9
Global Const wait_for_set_measur = 10
Global Const complete_prove = 11
Global Const drawing_picture = 12
Global Const choose_input_menu = 13
Global Const get_inform = 14
Global Const op_change_picture = 15
Global Const wait_for_input_char = 1001
Global Const wait_for_input_sentence = 1002
Global Const wait_for_modify_sentence = 1003
Global Const wait_for_modify_char = 1004
Global Const input_char_again = 1005
Global Const re_name = 1006
Global Const wenti_to_picture = 3000
Global Const picture_to_wenti = 4000
Global Const wait_for_draw_point = 4001
Global Const draw_point_down = 4002
Global Const draw_point_move = 4003
Global Const draw_point_up = 4004
Global Const input_char_in_draw = 4005  'draw_form中输入字符
Global Const postpone_ = 4006
Global Const stop_ = 4007
Global Const exit_prove = 4008
Global Const modify_picture = 5000
Global Const move_point = 5100
Global Const move_point_moude_down = 5101
Global Const move_point_moude_move = 5102
Global Const move_point_moude_up = 5103
Global Const new_free_point = 36
'*****************************************************
'combine_two_total_angle_type
Global Const A1_add_A2_equal_A3 = 1
Global Const A1_minus_A2_equal_A3 = 2
Global Const A2_minus_A1_equal_A3 = 3
Global Const A1_add_A2_add_A3_equal_360 = 4
Global Const A1_equal_A2 = 5
'Global Const A2_add_A1_add_A3_equal_180 = 5
'Global Const A2_add_A1_add_a3_equal_360 = 6
'Global Const A3_add_A1_minus_A2 = 7
'Global Const A3_minus_A1_add_sA2 = 8

'*******************************************************
'交点的类型
Global Const exist_point = 1
Global Const paral_ = 2
Global Const verti_ = 3
Global Const new_point_ = 4
Global Const two_point_line_ = 6004
Global Const new_point_on_line = 6005
Global Const new_point_on_circle = 6006
Global Const interset_point_line_line = 6007
'Global Const new_point_on_line_Tline = 8
Global Const new_point_on_line_circle = 6009
Global Const new_point_on_line_circle12 = 6010 '
Global Const new_point_on_line_circle21 = 6011
Global Const new_point_on_circle_circle = 6012
'Global Const new_point_on_circle_Tcircle = 24
Global Const new_point_on_circle_circle12 = 6013
Global Const new_point_on_circle_circle21 = 6014
Global Const exist_point_on_line_circle = 6015
Global Const Ratio_for_measure_ = 16
Global Const length_depended_by_two_points_ = 6017
Global Const ratio_point_on_line_ = 18
Global Const ratio_by_radii_of_circle_ = 6019
Global Const tangent_line_by_point_on_circle = verti_ '5
Global Const tangent_line_by_point_out_off_circle12 = -new_point_on_circle_circle12
Global Const tangent_line_by_point_out_off_circle21 = -new_point_on_circle_circle21
Global Const inner_tangent_line_by_two_circle00 = 6019
Global Const inner_tangent_line_by_two_circle12 = 6020
Global Const inner_tangent_line_by_two_circle21 = 6021
Global Const out_tangent_line_by_two_circle12 = 6022
Global Const out_tangent_line_by_two_circle21 = 6023
Global Const tangent_point_of_circle = 6025
Global Const tangent_point_of_line = 6026
Global Const tangent_point_ = 6024
Global Const circle_radio_relate_by_two_point_ = 6027
Global Const point_not_on_line = 0
Global Const point_on_segement = 26
Global Const point_out_segement = 27
'*************************************************************
'Global Const mid_point_line = 82
'**************************************************************************
Global Const set_new_wenti_cond = 35
Global Const no_modify = 0
Global Const modify = 1

Sub equal_char(ByVal A%, ByVal b%, ByVal c%) 'a 1st char,b 2nd,c input_cond
Dim m%, n%, f$, s$, FIS$, SEC$, THR$  ',n%
Dim i%, j%, k%, l% 'As Integer '第c%句中第a%个字于第b%个字相同
Dim ts$
k% = 0
ts$ = C_display_wenti.m_string(c%)
l% = Len(ts$) ' 输入语句长
For i = 1 To l%
If Asc(Mid$(ts$, i, 1)) < 160 Then
k% = k% + 2   '计算总长 ，每个字符为二份
Else
k% = k% + 1
End If

If k% = 2 * A% Or k% + 1 = 2 * A% Then
n% = i '两字节，计第一个
If k% + 1 = 2 * A% Then
k% = k% + 1
i = i + 1
End If

ElseIf k% = 2 * b% Or k% + 1 = 2 * b% Then
m% = i
If k% + 1 = 2 * b% Then
k% = k% + 1
i = i + 1
End If

End If

Next i
     
f$ = Mid$(ts$, n%, 1) '???????

If A% < b% Then
'修改的字在后
 FIS$ = Mid$(ts$, 1, n% - 1)
 If Asc(Mid(ts$, n%, 1)) > 159 Then
  SEC$ = Mid$(ts$, n% + 2, m% - n% - 2)
 Else
  SEC$ = Mid$(ts$, n% + 1, m% - n% - 1)
 End If
 If Asc(Mid(ts$, m%, 1)) > 159 Then
  THR$ = Mid$(ts$, m% + 2, l% - m% - 1)
 Else
  THR$ = Mid$(ts$, m% + 1, l% - m%)
 End If
Else
 FIS$ = Mid$(ts$, 1, m% - 1)
  SEC$ = Mid$(ts$, m% + 1, n% - m% - 1)
   THR$ = Mid$(ts$, n% + 1, l% - n%)
 End If
  If left(right(SEC$, 2), 1) = LoadResString_(1410, "") Then
  '修改圆后面的字符
  Call C_display_wenti.set_m_string("", FIS$ + f$ + SEC$ + _
     "(" + f$ + ")" + THR$, "", "", "", 0, c%, 0, 0)
  Else
  Call C_display_wenti.set_m_string("", FIS$ + f$ + SEC$ + f$ + THR$, "", "", "", 0, c%, 0, 0)
  End If


End Sub


Sub relation_char(select_wenti As Integer, m%)
Select Case C_display_wenti.m_no(select_wenti)
Case 8
 If m% = 3 Then
Call equal_char(3, 11, select_wenti)
End If
'Case 10
 'If m% = 4 Then
  'Call equal_char(4, 3, select_wenti)
   'End If
'Case 11
'If m% = 10 Then
'Call equal_char(m%, m% - 1, select_wenti)
'End If
'Case 12
'If m% = 3 Or m% = 8 Then
'Call equal_char(m%, m% - 1, select_wenti)
'End If
'Case 13
'If m% = 5 Or m% = 10 Then
'Call equal_char(m%, m% - 1, select_wenti)
'End If

Case 25
If m% = 3 Then
Call equal_char(3, 14, select_wenti)
ElseIf m% = 14 Then
Call equal_char(14, 3, select_wenti)
ElseIf m% = 4 Then

Call equal_char(4, 15, select_wenti)
ElseIf m% = 15 Then
Call equal_char(15, 4, select_wenti)
ElseIf m% = 6 Then
Call equal_char(6, 19, select_wenti)
ElseIf m% = 19 Then
Call equal_char(19, 6, select_wenti)
ElseIf m% = 7 Then
Call equal_char(7, 20, select_wenti)
ElseIf m% = 20 Then
Call equal_char(20, 7, select_wenti)
End If
Case 30
If m% = 1 Then
Call equal_char(1, 6, select_wenti)
ElseIf m% = 6 Then
Call equal_char(6, 1, select_wenti)
End If
Case 31
If m% = 3 Then
Call equal_char(3, 12, select_wenti)
ElseIf m% = 12 Then
Call equal_char(12, 3, select_wenti)
ElseIf m% = 4 Then
Call equal_char(4, 16, select_wenti)
ElseIf m% = 16 Then
Call equal_char(16, 4, select_wenti)
ElseIf m% = 9 Then
Call equal_char(9, 13, select_wenti)
Call equal_char(9, 15, select_wenti)
ElseIf m% = 13 Then
Call equal_char(13, 9, select_wenti)
Call equal_char(13, 15, select_wenti)


ElseIf m% = 15 Then
Call equal_char(15, 9, select_wenti)
Call equal_char(15, 13, select_wenti)
End If

End Select
End Sub

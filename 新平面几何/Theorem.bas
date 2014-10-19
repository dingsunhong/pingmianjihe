Attribute VB_Name = "theorem"
Option Explicit
'**********************************************************
Global last_right_angle_for_Pd As Integer ' 勾股定理
Global old_last_right_angle_for_Pd As Integer
'Global old_last_general_string_combine As Integer
'Global old_last_angle3_combine As Integer
'Global old_last_general_string_combine_for_aid As Integer
Public Function call_theorem(ByVal t As Byte, ByVal no_reduce As Byte) As Byte
Dim i%
Dim T_total_condition As Integer
If run_type = 1 Then
run_type = 2
End If
Do
T_total_condition = last_conditions.last_cond(1).total_condition
call_theorem = theorem1(no_reduce)
If call_theorem > 1 Then
 Exit Function
End If
Loop Until last_conditions.last_cond(1).total_condition = T_total_condition 'call_theorem = 0
End Function
Public Function set_prove_type(ty1 As Byte, no%, new_re As record_data_type, _
     old_re As record_data_type) As Boolean
      '化简证明过程
Dim i%
Dim index(7) As Integer
Dim temp_record As total_record_type
Dim re As record_data_type
If new_re.data0.condition_data.level >= old_re.data0.condition_data.level Then '新结论比结论步骤多
  Exit Function
End If
'拟相似和拟全等
If old_re.data0.condition_data.condition(1).ty = pseudo_total_equal_triangle_ Or _
        old_re.data0.condition_data.condition(1).ty = pseudo_similar_triangle_ Then
 If is_not_pseudo_record(new_re) Then
    For i = 1 To 8
    old_re.data0.condition_data.condition(i%).ty = new_re.data0.condition_data.condition(i%).ty
    old_re.data0.condition_data.condition(i%).no = new_re.data0.condition_data.condition(i%).no
    Next i%
    old_re.data0.condition_data.condition_no = new_re.data0.condition_data.condition_no
    old_re.data0.theorem_no = new_re.data0.theorem_no
      Exit Function
 End If
ElseIf new_re.data0.condition_data.condition(1).ty = pseudo_total_equal_triangle_ Or _
        new_re.data0.condition_data.condition(1).ty = pseudo_similar_triangle_ Then
         Exit Function
End If
re = new_re
 '化简
 If set_or_prove < 2 Then '自动证明
  If re.data0.condition_data.level < old_re.data0.condition_data.level Or _
   (re.data0.condition_data.condition_no < old_re.data0.condition_data.condition_no And _
     re.data0.condition_data.level = old_re.data0.condition_data.level) Then
  If re.data0.condition_data.level < old_re.data0.condition_data.level Then
   Call short_route(ty1, no%, re)
  End If
      For i% = 1 To 8
       old_re.data0.condition_data.condition(i%) = re.data0.condition_data.condition(i%)
        'old_re.data0.condition_data.condition(i%) = re.data0.condition_data.condition(i%)
      Next i%
        old_re.data0.condition_data.condition_no = re.data0.condition_data.condition_no
         old_re.data0.theorem_no = re.data0.theorem_no
          old_re.data0.condition_data.level = re.data0.condition_data.level
           old_re.data1.display_type = re.data1.display_type
          set_prove_type = True '证明路径有简化
  End If
'*****************************************************************
ElseIf set_or_prove = 2 And display_inform = 1 Then '手工证明用 直接推理
If old_re.data0.condition_data.condition_no = 0 And old_re.data1.is_proved = 0 Then 'And _
  old_re.is_proved = 1 '引用已知
ElseIf old_re.data1.is_proved = 0 Then
  For i% = 1 To old_re.data0.condition_data.condition_no
   Call record_no(old_re.data0.condition_data.condition(i%).ty, _
    old_re.data0.condition_data.condition(i%).no, temp_record, False, 0, 0)
     If temp_record.record_data.data1.is_proved = 0 And temp_record.record_data.data0.condition_data.condition_no > 0 Then
      prove_type = 0
         '结论正确但尚不能推出
       Exit Function
     End If
  Next i%
    prove_type = 2
ElseIf old_re.data1.is_proved = 1 Then '再次验证
  prove_type = 3
'ElseIf old_re.is_proved = 2 Then
 'prove_type = 4
End If
End If
End Function

Public Function T_cos(s1$, S2$, s3$, cal_float As Boolean) As String
Dim ts$
ts$ = cos_(s1$, 0)
If InStr(1, ts$, "F", 0) = 0 Then
ts$ = sqr_string(minus_string(add_string(time_string(S2$, S2$, False, cal_float), _
                time_string(s3$, s3$, False, cal_float), False, cal_float), _
                 time_string("2", time_string(time_string(S2$, s3$, False, cal_float), _
                  ts$, False, cal_float), False, cal_float), False, cal_float), True, cal_float)
If InStr(1, ts$, "F", 0) > 0 Then
 T_cos = "F"
Else
 T_cos = ts$
End If
Else
T_cos = ts$
End If
End Function

Public Function Th_cos(ByVal A As String, ByVal cos_A As String, _
          ByVal l1%, ByVal l2%, ByVal re_value$, cal_float As Boolean) As String
Dim ts$
If cos_A <> "" And InStr(1, cos_A, "F", 0) = 0 Then
ts$ = cos_A
Else
ts$ = cos_(A, 0)
End If
If InStr(1, ts$, "F", 0) > 0 Then
Th_cos = "F"
Else
If l1% > 0 And l2% > 0 Then
ts$ = sqr_string(minus_string(add_string(line_value(l1%).data(0).data0.squar_value, _
                 line_value(l2%).data(0).data0.squar_value, False, cal_float), _
                    time_string("2", time_string( _
                      time_string(line_value(l1%).data(0).data0.value, _
                       line_value(l2%).data(0).data0.value, False, cal_float), _
                          ts$, False, cal_float), False, cal_float), _
                           False, cal_float), True, cal_float)
If InStr(1, ts$, "F", 0) = 0 Then
 Th_cos = ts$
Else
 Th_cos = "F"
End If
ElseIf re_value$ <> "" Then
Th_cos = sqr_string(minus_string(add_string(time_string(re_value$, re_value$, False, False), _
                 "1", False, False), _
                    time_string("2", time_string(re_value$, _
                          ts$, False, False), False, False), False, False), True, False)
End If
End If
End Function

Public Function conclusion_from_new_point(ByVal p%, re As total_record_type, _
      ByVal l1%, ByVal l2%, ByVal n1%, ByVal n2%, ByVal c1%, ByVal c2%, _
        no_reduce As Byte) As Byte
Dim i%, j%, k%
Dim l%, m%, t%, no%
Dim A(1) As Integer
Dim n(1) As Integer
Dim tp(2) As Integer
Dim tl(2) As Integer
Dim dn(2) As Integer
Dim re_no%
Dim tA As angle3_value_type
Dim tn() As Integer
Dim last_tn%
Dim n_(1) As Integer
Dim temp_record As total_record_type
'On Error GoTo conclusion_from_new_point_error
temp_record = re
re_no% = re.record_data.data0.condition_data.condition_no
'************************************************
For i% = 1 To last_conditions.last_cond(1).eangle_no
no% = Deangle.av_no(i%).no
If angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l1% Or _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l2% Or _
     angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l1% Or _
      angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l2% Or _
   angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(0) = l1% Or _
    angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(0) = l2% Or _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(1) = l1% Or _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(1) = l2% Then
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
temp_record.record_data.data0.condition_data.condition(1).no = no%
If angle3_value(no%).record_.no_reduce < 4 Then
conclusion_from_new_point = _
set_total_equal_triangle_from_eangle(angle3_value(no%).data(0).data0.angle(0), _
   angle3_value(no%).data(0).data0.angle(1), temp_record, p%, l1, l2%, n1%, n2%, _
     no_reduce, 1)
If conclusion_from_new_point > 1 Then
Exit Function
End If
For j% = 0 To 1
If p% = inter_point_of_segment(angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(j%), _
       angle(angle3_value(no%).data(0).data0.angle(1)).data(0).te(j%)) And _
          th_chose(134).chose = 1 Then
  tp(0) = inter_point_of_segment(angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2), _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(no%).data(0).data0.angle(1)).data(0).te((j% + 1) Mod 2))
  temp_record.record_data.data0.theorem_no = 134
   If tp(0) > 0 And tp(0) <> p% Then
    conclusion_from_new_point = set_four_point_on_circle(p%, tp(0), _
     angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), 0, temp_record, 0, no_reduce)
      If conclusion_from_new_point > 1 Then
       Exit Function
      End If
       GoTo conclusion_from_new_point1
   End If
  End If
Next j%
conclusion_from_new_point1:
End If
End If
Next i%

'**************************************************
For i% = 2 To last_conditions.last_cond(1).angle_value_no
n(0) = angle_value.av_no(i%).no
If angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l1% Or _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l2% Or _
     angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l1% Or _
      angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l2% Then
If angle3_value(n(0)).record_.no_reduce < 4 Then
 For k% = 1 To i% - 1
  n(1) = angle_value.av_no(k%).no
   temp_record.record_data.data0.condition_data.condition_no = re_no% + 2
    temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = angle3_value_
     temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = n(0)
      temp_record.record_data.data0.condition_data.condition(re_no% + 2).ty = angle3_value_
       temp_record.record_data.data0.condition_data.condition(re_no% + 2).no = n(1)
  If angle3_value(n(1)).record_.no_reduce < 4 Then
   If angle3_value(n(0)).data(0).data0.value = angle3_value(n(1)).data(0).data0.value Then
    conclusion_from_new_point = _
     set_total_equal_triangle_from_eangle(angle3_value(n(0)).data(0).data0.angle(0), _
      angle3_value(n(1)).data(0).data0.angle(0), temp_record, p%, l1%, l2%, n1%, n2%, _
       no_reduce, 1)
   If conclusion_from_new_point > 1 Then
    Exit Function
   End If
  For j% = 0 To 1
If p% = inter_point_of_segment(angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no(j%), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).te(j%)) And _
        th_chose(134).chose = 1 Then
  tp(0) = inter_point_of_segment(angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
    angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2), _
     angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2))
If tp(0) > 0 And tp(0) <> p% Then
 temp_record.record_data.data0.theorem_no = 134
     conclusion_from_new_point = set_four_point_on_circle(p%, tp(0), _
      angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), 0, temp_record, 0, no_reduce)
      If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point2
 End If
 ElseIf angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%) = _
          angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no(j%) And _
           angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1) <> _
            angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1) And _
             th_chose(8).chose = 1 Then
     temp_record.record_data.data0.theorem_no = 8
 conclusion_from_new_point = set_dparal( _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
    angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
     temp_record, 0, no_reduce, False)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point2
 ElseIf p% = inter_point_of_segment(angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2)) And _
          angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2) = _
            angle(angle3_value(n(1)).data(0).data0.angle(1)).data(0).line_no(j%) _
             And th_chose(40).chose = 1 Then
   temp_record.record_data.data0.theorem_no = 40
   conclusion_from_new_point = set_equal_dline(p%, angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
     p%, angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), 0, 0, 0, 0, 0, 0, 0, temp_record, 0, 0, _
       0, 0, no_reduce, False)
      If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point2
  End If
Next j%
ElseIf add_string(angle3_value(n(0)).data(0).data0.value, angle3_value(n(1)).data(0).data0.value, _
               True, False) = "180" Then
 For j% = 0 To 1
  If p% = inter_point_of_segment(angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2)) And _
        th_chose(32).chose = 1 Then
  tp(0) = inter_point_of_segment(angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).te((j% + 1) Mod 2))
  If tp(0) > 0 And tp(0) <> p% Then
   temp_record.record_data.data0.theorem_no = 132
   conclusion_from_new_point = set_four_point_on_circle(p%, tp(0), _
     angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1), 0, temp_record, 0, no_reduce)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point3
  End If
 ElseIf angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no(j%) = _
      angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2) And _
        angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).poi(1) <> _
         angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).poi(1) And _
          th_chose(10).chose = 1 Then
  temp_record.record_data.data0.theorem_no = 10
 conclusion_from_new_point = set_dparal( _
   angle(angle3_value(n(0)).data(0).data0.angle(0)).data(0).line_no((j% + 1) Mod 2), _
    angle(angle3_value(n(1)).data(0).data0.angle(0)).data(0).line_no(j%), _
     temp_record, 0, no_reduce, False)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point3
 End If
conclusion_from_new_point2:
 Next j%
 End If
 End If
 Next k%
 End If
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).two_angle_value_no
 no% = Two_angle_value.av_no(i%).no
If angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l1% Or _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(0) = l2% Or _
     angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l1% Or _
      angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(1) = l2% Or _
   angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(0) = l1% Or _
    angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(0) = l2% Or _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(1) = l1% Or _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no(1) = l2% Then
If angle3_value(no%).record_.no_reduce < 4 Then
If angle3_value(no%).data(0).data0.value = "180" And _
    angle3_value(no%).data(0).data0.para(0) = "1" And _
     angle3_value(no%).data(0).data0.para(1) = "1" Then
For j% = 0 To 1
        temp_record.record_data.data0.condition_data.condition_no = re_no% + 1
      temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = angle3_value_
     temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = no%
  If p% = inter_point_of_segment(angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(no%).data(0).data0.angle(1)).data(0).te((j% + 1) Mod 2)) Then
  tp(0) = inter_point_of_segment(angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
   angle(angle3_value(no%).data(0).data0.angle(0)).data(0).line_no(j%), _
    angle(angle3_value(no%).data(0).data0.angle(0)).data(0).te(j%), _
     angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).line_no((j% + 1) Mod 2), _
       angle(angle3_value(no%).data(0).data0.angle(1)).data(0).te((j% + 1) Mod 2))
If tp(0) > 0 And tp(0) <> p% Then
   conclusion_from_new_point = set_four_point_on_circle(p%, tp(0), _
     angle(angle3_value(no%).data(0).data0.angle(0)).data(0).poi(1), _
      angle(angle3_value(no%).data(0).data0.angle(1)).data(0).poi(1), 0, temp_record, 0, no_reduce)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
      GoTo conclusion_from_new_point3
  End If
 End If
 Next j%
 End If
conclusion_from_new_point3:
End If
End If
Next i%
'*********************************************
For i% = 1 To last_conditions.last_cond(1).paral_no
 If Dparal(i%).data(0).data0.line_no(0) = l1% Or Dparal(i%).data(0).data0.line_no(0) = l2% Or _
     Dparal(i%).data(0).data0.line_no(1) = l2% Or Dparal(i%).data(0).data0.line_no(1) = l2% Then
 If Dparal(i%).record_.no_reduce < 4 Then
 For j% = 0 To 1
  temp_record.record_data.data0.condition_data.condition_no = re_no% + 1
  temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = paral_
   temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = i%
   If is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(j%)).data(0).data0, 0) Then
   For k% = 1 To m_lin(Dparal(i%).data(0).data0.line_no((j% + 1) Mod 2)).data(0).data0.in_point(0)
   ' tl(0) = line_number0(p%, _
     Lin(Dparal(i%).data(0).line_no((j% + 1) Mod 2)).data(0).data0.in_point(k), 0, 0)
    conclusion_from_new_point = set_angle_from_paral(Dparal(i%).data(0).data0.line_no(j%), _
     Dparal(i%).data(0).data0.line_no((j% + 1) Mod 2), p%, _
      m_lin(Dparal(i%).data(0).data0.line_no((j% + 1) Mod 2)).data(0).data0.in_point(k), temp_record.record_data, no_reduce)
    If conclusion_from_new_point > 1 Then
     Exit Function
    End If
   Next k%
  End If
  Next j%
  End If
  End If
  Next i%
  For i% = 1 To last_conditions.last_cond(1).eline_no
  If Deline(i%).data(0).data0.line_no(0) = l1% Or Deline(i%).data(0).data0.line_no(0) = l2% Or _
      Deline(i%).data(0).data0.line_no(1) = l1% Or Deline(i%).data(0).data0.line_no(1) = l2% Then
  temp_record.record_data.data0.condition_data.condition_no = re_no% + 1
   temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = eline_
    temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = i%
  conclusion_from_new_point = _
   set_total_equal_triangle_from_eline(Deline(i%).data(0).data0.poi(0), _
    Deline(i%).data(0).data0.poi(1), Deline(i%).data(0).data0.poi(2), _
     Deline(i%).data(0).data0.poi(3), temp_record, p%, no_reduce)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
  End If
 Next i%
  For i% = 1 To last_conditions.last_cond(1).mid_point_no
   If Dmid_point(i%).data(0).data0.line_no = l1% Or _
        Dmid_point(i%).data(0).data0.line_no = l2% Then
  temp_record.record_data.data0.condition_data.condition_no = re_no% + 1
   temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = midpoint_
    temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = i%
 conclusion_from_new_point = _
  set_total_equal_triangle_from_eline(Dmid_point(i%).data(0).data0.poi(0), _
   Dmid_point(i%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.poi(1), _
    Dmid_point(i%).data(0).data0.poi(2), temp_record, p%, no_reduce)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
  End If
 Next i%
  For i% = 2 To last_conditions.last_cond(1).line_value_no
   If line_value(i%).data(0).data0.line_no = l1% Or _
       line_value(i%).data(0).data0.line_no = l2% Then
   For j% = 1 To i% - 1
   If line_value(i%).data(0).data0.value = line_value(j%).data(0).data0.value Then
  temp_record.record_data.data0.condition_data.condition_no = re_no% + 2
   temp_record.record_data.data0.condition_data.condition(re_no% + 1).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(re_no% + 1).no = i%
       temp_record.record_data.data0.condition_data.condition(re_no% + 2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(re_no% + 2).no = j%
conclusion_from_new_point = _
  set_total_equal_triangle_from_eline(line_value(i%).data(0).data0.poi(0), _
   line_value(i%).data(0).data0.poi(1), line_value(j%).data(0).data0.poi(0), _
    line_value(j%).data(0).data0.poi(0), temp_record, p%, no_reduce)
       If conclusion_from_new_point > 1 Then
       Exit Function
      End If
 End If
  Next j%
  End If
  Next i%
For i% = 1 To last_conditions.last_cond(1).verti_no
 If Dverti(i%).data(0).line_no(0) = l1% Or Dverti(i%).data(0).line_no(0) = l2% Or _
     Dverti(i%).data(0).line_no(1) = l1% Or Dverti(i%).data(0).line_no(1) = l2% Then
 If is_line_line_intersect(Dverti(i%).data(0).line_no(0), _
         Dverti(i%).data(0).line_no(1), 0, 0, False) = p% Then
    temp_record = re
     Call add_conditions_to_record(verti_, i%, 0, 0, temp_record.record_data.data0.condition_data)
      conclusion_from_new_point = set_angle_value(Abs(angle_number( _
       m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0.poi(0), p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0.poi(0), 0, 0)), _
        "90", temp_record, 0, no_reduce, False)
    temp_record = re
     Call add_conditions_to_record(verti_, i%, 0, 0, temp_record.record_data.data0.condition_data)
      conclusion_from_new_point = set_angle_value(Abs(angle_number( _
       m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0.poi(0), p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0.poi(1), 0, 0)), _
         "90", temp_record, 0, no_reduce, False)
    temp_record = re
     Call add_conditions_to_record(verti_, i%, 0, 0, temp_record.record_data.data0.condition_data)
      conclusion_from_new_point = set_angle_value(Abs(angle_number( _
       m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0.poi(1), p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0.poi(0), 0, 0)), _
         "90", temp_record, 0, no_reduce, False)
    temp_record = re
     Call add_conditions_to_record(verti_, i%, 0, 0, temp_record.record_data.data0.condition_data)
     conclusion_from_new_point = set_angle_value(Abs(angle_number( _
       m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0.poi(1), p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0.poi(1), 0, 0)), _
         "90", temp_record, 0, no_reduce, False)
 End If
 End If
Next i%
Exit Function
conclusion_from_new_point_error:
End Function
Public Function set_angle_from_paral(ByVal l1%, ByVal l2%, ByVal p1%, _
           ByVal p2%, re As record_data_type, ByVal no_reduce As Byte) As Byte
Dim l3%, dr%
Dim A(7) As Integer
Dim n(1) As Integer
Dim temp_record As total_record_type
l3% = line_number0(p1%, p2%, n(0), n(1))
If n(0) > n(1) Then
 Call exchange_two_integer(p1%, p2%)
  Call exchange_two_integer(l1%, l2%)
End If
dr% = compare_two_point(m_poi(m_lin(l2).data(0).data0.poi(0)).data(0).data0.coordinate, _
        m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, _
         m_lin(l1%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(1), 8)
If dr% = 1 Then
A(0) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(0), p1%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(1) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(0), p1%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(2) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(1), p1%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(3) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(1), p1%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(4) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(0), p2%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(5) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(0), p2%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(6) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(1), p2%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(7) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(1), p2%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
ElseIf dr% = -1 Then
A(0) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(0), p1%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(1) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(0), p1%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(2) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(1), p1%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(3) = Abs(angle_number(m_lin(l1%).data(0).data0.poi(1), p1%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(4) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(1), p2%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(5) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(1), p2%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
A(6) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(0), p2%, m_lin(l3%).data(0).data0.poi(0), 0, 0))
A(7) = Abs(angle_number(m_lin(l2%).data(0).data0.poi(0), p2%, m_lin(l3%).data(0).data0.poi(1), 0, 0))
Else
 Exit Function
End If
If A(0) > 0 And A(4) > 0 Then
temp_record.record_data = re
 temp_record.record_data.data0.theorem_no = 11
 set_angle_from_paral = set_three_angle_value(A(0), A(4), 0, _
   "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
   Exit Function
 End If
  End If
If A(1) > 0 And A(5) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 11
set_angle_from_paral = set_three_angle_value(A(1), A(5), 0, _
   "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(2) > 0 And A(6) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 11
 set_angle_from_paral = set_three_angle_value(A(2), A(6), 0, _
   "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(3) > 0 And A(7) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 11
 set_angle_from_paral = set_three_angle_value(A(3), A(7), 0, _
   "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
  If set_angle_from_paral > 1 Then
  Exit Function
 End If
 End If
If A(0) > 0 And A(7) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 12
set_angle_from_paral = set_three_angle_value(A(0), A(7), 0, _
  "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(1) > 0 And A(6) > 0 Then
   temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 12
set_angle_from_paral = set_three_angle_value(A(1), A(6), 0, _
   "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(2) > 0 And A(5) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 12
set_angle_from_paral = set_three_angle_value(A(2), A(5), 0, _
  "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
  If set_angle_from_paral > 1 Then
  Exit Function
 End If
 End If
If A(3) > 0 And A(4) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 12
set_angle_from_paral = set_three_angle_value(A(3), A(4), 0, _
  "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(0) > 0 And A(5) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 13
set_angle_from_paral = set_three_angle_value(A(0), A(5), 0, "1", "1", "0", _
     "180", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(1) > 0 And A(4) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 13
set_angle_from_paral = set_three_angle_value(A(1), A(4), 0, "1", "1", "0", _
    "180", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(2) > 0 And A(7) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 13
set_angle_from_paral = set_three_angle_value(A(2), A(7), 0, "1", "1", "0", _
    "180", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If set_angle_from_paral > 1 Then
  Exit Function
 End If
  End If
If A(3) > 0 And A(6) > 0 Then
  temp_record.record_data = re
   temp_record.record_data.data0.theorem_no = 13
set_angle_from_paral = set_three_angle_value(A(3), A(6), 0, "1", "1", "0", _
  "180", 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 'If set_angle_from_paral > 1 Then
  'Exit Function
 'End If
  End If
End Function



Public Function arrange_four_point0(ByVal n1%, ByVal n2%, ByVal n3%, _
   ByVal n4%, on1%, on2%, on3%) As Byte
If n1% > n2% Then
 Call exchange_two_integer(n1%, n2%)
End If
If n3% > n4% Then
Call exchange_two_integer(n3%, n4%)
End If
If n1% = n3% Then
 If n2% < n4% Then
  arrange_four_point0 = 7
  on1% = n1%
   on2% = n2%
    on3% = n4%
 Else
  arrange_four_point0 = 8
   on1% = n1%
    on2% = n4%
     on3% = n2%
 End If
ElseIf n2% = n4% Then
 If n1% < n3% Then
  arrange_four_point0 = 4
  on1% = n1%
   on2% = n3%
    on3% = n4%
 Else
  arrange_four_point0 = 6
  on1% = n3%
   on2% = n1%
    on3% = n4%
 End If
ElseIf n2% = n3% Then
 arrange_four_point0 = 3
  on1% = n1%
   on2% = n2%
    on3% = n4%
ElseIf n4% = n1% Then
 arrange_four_point0 = 5
  on1% = n3%
   on2% = n4%
    on3% = n2%
End If
End Function


Public Function value_from_two_relation(ByVal i1%, ByVal v1$, ByVal v2$) As String
If i1% = 2 Then '0,1 由relation 定，2 由combine定
value_from_two_relation = divide_string(v2$, v1$, True, False)
ElseIf i1% = 0 Then
value_from_two_relation = divide_string(v2$, v1$, True, False)
ElseIf i1% = 1 Then
value_from_two_relation = divide_string("1", time_string(v2$, v1$, False, False), _
         True, False)
End If
End Function

Public Sub add_conditions_to_record(ByVal con_ty As Integer, _
     ByVal con_no1%, con_no2%, con_no3%, re As condition_data_type)
Dim i%, k%
'Dim tem_re(2) As record_type
'On Error GoTo add_conditions_to_record_mark0
If re.condition_no > 200 Then
    re.condition_no = 0
ElseIf re.condition_no > 5 Or con_ty = 0 Then
 Exit Sub
End If
k% = c_data_for_reduce.condition_no
c_data_for_reduce.condition_no = 0
For i% = 1 To k%
 Call add_conditions_to_record(c_data_for_reduce.condition(i%).ty, _
        c_data_for_reduce.condition(i%).no, 0, 0, re)
Next i%
Call add_condition_to_record(con_ty, con_no1, re, 0)
Call add_condition_to_record(con_ty, con_no2, re, 0)
Call add_condition_to_record(con_ty, con_no3, re, 0)
End Sub
Public Function set_right_triangle_from_mid_line(ByVal m_p1%, _
  ByVal m_p2%, ByVal p1%, ByVal p2%, re As total_record_type) As Byte '10.10
Dim n%, A%, n1%
Dim m_p(1) As Integer
Dim con_ty As Byte
Dim temp_record As total_record_type
Dim c_data As condition_data_type
If th_chose(130).chose = 1 Then
m_p(0) = m_p1%
m_p(1) = m_p2%
record_0.data0.condition_data.condition_no = 0 ' record0
temp_record = re
temp_record.record_data.data0.theorem_no = 130
c_data.condition_no = 0
If is_mid_point(p1%, m_p1%, p2%, 0, 0, 0, 0, n%, -1000, _
     0, 0, 0, 0, 0, 0, Dmid_point_data0, "", con_ty, n%, n1%, c_data) Then
Call add_conditions_to_record(con_ty, n%, n1%, 0, temp_record.record_data.data0.condition_data)
A% = angle_number(p1%, m_p2%, p2%, 0, 0)
If A% <> 0 Then
 set_right_triangle_from_mid_line = set_angle_value(Abs(A%), _
    "90", temp_record, 0, 1, False)
End If
Else
n% = 0
c_data.condition_no = 0
If is_mid_point(p1%, m_p2%, p2%, 0, 0, 0, 0, n%, -1000, _
     0, 0, 0, 0, 0, 0, Dmid_point_data0, "", con_ty, n%, n1%, c_data) Then
Call add_conditions_to_record(con_ty, n%, n1%, 0, temp_record.record_data.data0.condition_data)
A% = angle_number(p1%, m_p1%, p2%, 0, 0)
If A% <> 0 Then
set_right_triangle_from_mid_line = set_angle_value(Abs(A%), _
   "90", temp_record, 0, 1, False)
End If
End If
End If
End If
End Function


Public Sub read_ratio_from_relation(ByVal s0$, ty As Integer, v1$, v2$, _
     is_con_line As Boolean, con_line_ty As Byte)
If is_con_line And (con_line_ty = 3 Or con_line_ty = 5) Then
 If ty = 0 Then
 v1$ = s0$
 v2$ = divide_string(s0$, add_string("1", s0$, False, False), True, False)
 ElseIf ty = 1 Then
 v1$ = divide_string("1", add_string("1", s0$, False, False), True, False)
 v2$ = divide_string("1", s0$, True, False)
 ElseIf ty = 2 Then
 v1$ = divide_string(add_string("1", s0$, False, False), s0$, True, False)
 v2$ = add_string("1", s0$, True, False)
 End If
Else
 If ty = 0 Then
 v1$ = s0$
 v2$ = ""
 ElseIf ty = 1 Then
 v1$ = divide_string("1", s0$, True, False)
 v2$ = ""
 Else
 v1$ = ""
 v2$ = ""
 End If
End If
End Sub

Public Sub read_ratio_from_Drelation(ByVal re%, ty As Integer, v1$, v2$)
Dim is_coline As Boolean
If Drelation(re%).data(0).data0.poi(4) > 0 And Drelation(re%).data(0).data0.poi(5) > 0 Then
is_coline = True
Else
is_coline = False
End If
Call read_ratio_from_relation(Drelation(re%).data(0).data0.value, ty, v1$, v2$, is_coline, _
        Drelation(re%).data(0).data0.ty)
End Sub

Public Function set_diff_for_item(ByVal it%, k%, ByVal l%, no_reduce As Byte) As Byte
Dim i%
Dim n(1) As Integer
Dim re As condition_data_type
'Dim temp_record As record_type
If k% < 2 Then
 Exit Function
End If
re.condition_no = 1
 re.condition(1).ty = line_value_
  re.condition(1).no = l%
If item0(it%).data(0).sig = "*" Then
 If item0(it%).data(0).diff_type = 3 Or item0(it%).data(0).diff_type = 5 Then
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", item0(it%).data(0).n(0), _
      item0(it%).data(0).n(1), item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
       item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), "1", "1", "1", "", "1", 0, _
        record_data0.data0.condition_data, 0, n(0), no_reduce, 0, condition_data0, False)
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
      0, 0, item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, _
       record_data0.data0.condition_data, 0, n(1), no_reduce, 0, condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), n(1), "-1", _
      line_value(l%).data(0).data0.value, it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
ElseIf item0(it%).data(0).diff_type = 4 Or item0(it%).data(0).diff_type = 8 Then
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", _
      item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).n(0), _
       item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
        "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, n(0), no_reduce, _
           0, condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), 0, 0, _
      item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
        0, n(1), no_reduce, 0, condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), n(1), "1", _
      time_string("-1", line_value(l%).data(0).data0.value, True, False), it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
ElseIf item0(it%).data(0).diff_type = 6 Or item0(it%).data(0).diff_type = 7 Then
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", _
      item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).n(0), _
       item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
        "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, n(0), no_reduce, _
           0, condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
     0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
      0, 0, item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
        0, n(1), no_reduce, 0, condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), n(1), "1", _
      time_string("-1", line_value(l%).data(0).data0.value, True, False), it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
End If
ElseIf item0(it%).data(0).sig = "/" Then
 If item0(it%).data(0).diff_type = 3 Or item0(it%).data(0).diff_type = 5 Then
  set_diff_for_item = set_item0(0, 0, item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), "/", _
    0, 0, item0(it%).data(0).n(2), item0(it%).data(0).n(3), 0, item0(it%).data(0).line_no(1), _
     "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, n(0), no_reduce, 0, _
        condition_data0, False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), 0, "-1", _
      time_string("-1", line_value(l%).data(0).data0.value, True, False), it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
ElseIf item0(it%).data(0).diff_type = 4 Or item0(it%).data(0).diff_type = 8 Then
  set_diff_for_item = set_item0(0, 0, item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), "/", _
    0, 0, item0(it%).data(0).n(2), item0(it%).data(0).n(3), 0, item0(it%).data(0).line_no(1), _
     "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, n(0), no_reduce, 0, condition_data0, _
        False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), 0, "1", _
      time_string("-1", line_value(l%).data(0).data0.value, True, False), it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
ElseIf item0(it%).data(0).diff_type = 6 Or item0(it%).data(0).diff_type = 7 Then
  set_diff_for_item = set_item0(0, 0, item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), "/", _
    0, 0, item0(it%).data(0).n(2), item0(it%).data(0).n(3), 0, item0(it%).data(0).line_no(1), _
     "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, n(0), no_reduce, 0, condition_data0, _
        False)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
  set_diff_for_item = add_new_item_to_item(n(0), 0, "1", _
      line_value(l%).data(0).data0.value, it%, re)
  If set_diff_for_item > 1 Then
   Exit Function
  End If
End If
End If
End Function



Public Sub change_four_arrange_type(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
     ByVal ty As Byte, op1%, op2%, op3%, op4%)
If ty = 3 Then
op1% = p1%
op2% = p2%
op3% = p3%
op4% = p4%
ElseIf ty = 4 Then
op1% = p1%
op2% = p3%
op3% = p3%
op4% = p4%
ElseIf ty = 5 Then
op1% = p3%
op2% = p4%
op3% = p1%
op4% = p2%
ElseIf ty = 6 Then
op1% = p3%
op2% = p1%
op3% = p1%
op4% = p4%
ElseIf ty = 7 Then
op1% = p1%
op2% = p2%
op3% = p2%
op4% = p4%
ElseIf ty = 8 Then
op1% = p1%
op2% = p4%
op3% = p4%
op4% = p2%
End If
End Sub
Public Sub verse_change_four_arrange_type(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
     ByVal ty As Byte, op1%, op2%, op3%, op4%)
If ty = 3 Then
op1% = p1%
op2% = p2%
op3% = p3%
op4% = p4%
ElseIf ty = 4 Then
op1% = p1%
op2% = p4%
op3% = p2%
op4% = p4%
ElseIf ty = 5 Then
op1% = p2%
op2% = p4%
op3% = p1%
op4% = p2%
ElseIf ty = 6 Then
op1% = p2%
op2% = p4%
op3% = p1%
op4% = p4%
ElseIf ty = 7 Then
op1% = p1%
op2% = p2%
op3% = p1%
op4% = p4%
ElseIf ty = 8 Then
op1% = p1%
op2% = p4%
op3% = p1%
op4% = p2%
End If
End Sub
Public Function read_point_and_ratio_from_relation(ByVal ty As Byte, _
             ByVal no%, k%, p() As Integer, n() As Integer, _
                l() As Integer, v1$, v2$) As Boolean
Dim tp(5) As Integer
Dim tn(5) As Integer
Dim tl(2) As Integer
Dim m(2) As Integer
If k% > 2 Then
 Exit Function
End If
If ty = line_value_ Then
p(0) = line_value(no%).data(0).data0.poi(0)
p(1) = line_value(no%).data(0).data0.poi(1)
v1$ = line_value(no%).data(0).data0.value
v2$ = ""
Else
If ty = relation_ Then
tp(0) = Drelation(no%).data(0).data0.poi(0)
tp(1) = Drelation(no%).data(0).data0.poi(1)
tp(2) = Drelation(no%).data(0).data0.poi(2)
tp(3) = Drelation(no%).data(0).data0.poi(3)
tp(4) = Drelation(no%).data(0).data0.poi(4)
tp(5) = Drelation(no%).data(0).data0.poi(5)
tn(0) = Drelation(no%).data(0).data0.n(0)
tn(1) = Drelation(no%).data(0).data0.n(1)
tn(2) = Drelation(no%).data(0).data0.n(2)
tn(3) = Drelation(no%).data(0).data0.n(3)
tn(4) = Drelation(no%).data(0).data0.n(4)
tn(5) = Drelation(no%).data(0).data0.n(5)
tl(0) = Drelation(no%).data(0).data0.line_no(0)
tl(1) = Drelation(no%).data(0).data0.line_no(1)
tl(2) = Drelation(no%).data(0).data0.line_no(2)
Call read_ratio_from_Drelation(no%, k%, v1$, v2$)
ElseIf ty = eline_ Then
tp(0) = Deline(no%).data(0).data0.poi(0)
tp(1) = Deline(no%).data(0).data0.poi(1)
tp(2) = Deline(no%).data(0).data0.poi(2)
tp(3) = Deline(no%).data(0).data0.poi(3)
tp(4) = 0
tp(5) = 0
tn(0) = Deline(no%).data(0).data0.n(0)
tn(1) = Deline(no%).data(0).data0.n(1)
tn(2) = Deline(no%).data(0).data0.n(2)
tn(3) = Deline(no%).data(0).data0.n(3)
tn(4) = 0
tn(5) = 0
tl(0) = Deline(no%).data(0).data0.line_no(0)
tl(1) = Deline(no%).data(0).data0.line_no(1)
tl(2) = 0
Call read_ratio_from_relation("1", k%, v1$, v2$, False, 0)
ElseIf ty = midpoint_ Then
tp(0) = Dmid_point(no%).data(0).data0.poi(0)
tp(1) = Dmid_point(no%).data(0).data0.poi(1)
tp(2) = Dmid_point(no%).data(0).data0.poi(1)
tp(3) = Dmid_point(no%).data(0).data0.poi(2)
tp(4) = Dmid_point(no%).data(0).data0.poi(0)
tp(5) = Dmid_point(no%).data(0).data0.poi(2)
tn(0) = Dmid_point(no%).data(0).data0.n(0)
tn(1) = Dmid_point(no%).data(0).data0.n(1)
tn(2) = Dmid_point(no%).data(0).data0.n(1)
tn(3) = Dmid_point(no%).data(0).data0.n(2)
tn(4) = Dmid_point(no%).data(0).data0.n(0)
tn(5) = Dmid_point(no%).data(0).data0.n(2)
tl(0) = Dmid_point(no%).data(0).data0.line_no
tl(1) = Dmid_point(no%).data(0).data0.line_no
tl(2) = Dmid_point(no%).data(0).data0.line_no
Call read_ratio_from_relation("1", k%, v1$, v2$, True, 3)
Else
 read_point_and_ratio_from_relation = False
  Exit Function
End If
 read_point_and_ratio_from_relation = True
If tp(4) > 0 And tp(5) > 0 Then
m(0) = k%
m(1) = (k% + 1) Mod 3
m(2) = (k% + 2) Mod 3
Else
m(0) = k%
m(1) = (k% + 1) Mod 2
m(2) = 2
End If
p(0) = tp(2 * m(0))
p(1) = tp(2 * m(0) + 1)
p(2) = tp(2 * m(1))
p(3) = tp(2 * m(1) + 1)
p(4) = tp(2 * m(2))
p(5) = tp(2 * m(2) + 1)
n(0) = tn(2 * m(0))
n(1) = tn(2 * m(0) + 1)
n(2) = tn(2 * m(1))
n(3) = tn(2 * m(1) + 1)
n(4) = tn(2 * m(2))
n(5) = tn(2 * m(2) + 1)
l(0) = tl(m(0))
l(1) = tl(m(1))
l(2) = tl(m(2))
End If
End Function
Public Sub read_point_and_value_from_line_value(ByVal ty As Byte, _
          no%, k%, p() As Integer, n() As Integer, l() As Integer, _
            para() As String, v$)
Dim tp(5) As Integer
Dim tn(5) As Integer
Dim tl(2) As Integer
Dim Tpara(3) As String
Dim m(2) As Integer
If ty = line_value_ Then
tp(0) = line_value(no%).data(0).data0.poi(0)
tp(1) = line_value(no%).data(0).data0.poi(1)
tp(2) = 0
tp(3) = 0
tp(4) = 0
tp(5) = 0
Tpara(0) = "1"
Tpara(1) = "0"
Tpara(2) = "0"
v$ = line_value(no%).data(0).data0.value
ElseIf ty = relation_ Then
tp(0) = Drelation(no%).data(0).data0.poi(0)
tp(1) = Drelation(no%).data(0).data0.poi(1)
tp(2) = Drelation(no%).data(0).data0.poi(2)
tp(3) = Drelation(no%).data(0).data0.poi(3)
tp(4) = 0
tp(5) = 0
tn(0) = Drelation(no%).data(0).data0.n(0)
tn(1) = Drelation(no%).data(0).data0.n(1)
tn(2) = Drelation(no%).data(0).data0.n(2)
tn(3) = Drelation(no%).data(0).data0.n(3)
tn(4) = 0
tn(5) = 0
tl(0) = Drelation(no%).data(0).data0.line_no(0)
tl(1) = Drelation(no%).data(0).data0.line_no(1)
tl(2) = 0
Tpara(0) = "1"
Tpara(1) = time_string("-1", Drelation(no%).data(0).data0.value, True, False)
Tpara(2) = "0"
v$ = "0"
ElseIf ty = eline_ Then
tp(0) = Deline(no%).data(0).data0.poi(0)
tp(1) = Deline(no%).data(0).data0.poi(1)
tp(2) = Deline(no%).data(0).data0.poi(2)
tp(3) = Deline(no%).data(0).data0.poi(3)
tp(4) = 0
tp(5) = 0
tn(0) = Deline(no%).data(0).data0.n(0)
tn(1) = Deline(no%).data(0).data0.n(1)
tn(2) = Deline(no%).data(0).data0.n(2)
tn(3) = Deline(no%).data(0).data0.n(3)
tn(4) = 0
tn(5) = 0
tl(0) = Deline(no%).data(0).data0.line_no(0)
tl(1) = Deline(no%).data(0).data0.line_no(1)
tl(2) = 0
Tpara(0) = "1"
Tpara(1) = "-1"
Tpara(2) = "0"
v$ = "0"
ElseIf ty = midpoint_ Then
tp(0) = Dmid_point(no%).data(0).data0.poi(0)
tp(1) = Dmid_point(no%).data(0).data0.poi(1)
tp(2) = Dmid_point(no%).data(0).data0.poi(1)
tp(3) = Dmid_point(no%).data(0).data0.poi(2)
tp(4) = 0
tp(5) = 0
tn(0) = Dmid_point(no%).data(0).data0.n(0)
tn(1) = Dmid_point(no%).data(0).data0.n(1)
tn(2) = Dmid_point(no%).data(0).data0.n(1)
tn(3) = Dmid_point(no%).data(0).data0.n(2)
tn(4) = 0
tn(5) = 0
tl(0) = Dmid_point(no%).data(0).data0.lin
tl(1) = Dmid_point(no%).data(0).data0.lin
tl(2) = 0
Tpara(0) = "1"
Tpara(1) = "-1"
Tpara(2) = "0"
v$ = "0"
ElseIf ty = two_line_value_ Then
tp(0) = two_line_value(no%).data(0).data0.poi(0)
tp(1) = two_line_value(no%).data(0).data0.poi(1)
tp(2) = two_line_value(no%).data(0).data0.poi(2)
tp(3) = two_line_value(no%).data(0).data0.poi(3)
tp(4) = 0
tp(5) = 0
tn(0) = two_line_value(no%).data(0).data0.n(0)
tn(1) = two_line_value(no%).data(0).data0.n(1)
tn(2) = two_line_value(no%).data(0).data0.n(2)
tn(3) = two_line_value(no%).data(0).data0.n(3)
tn(4) = 0
tn(5) = 0
tl(0) = two_line_value(no%).data(0).data0.line_no(0)
tl(1) = two_line_value(no%).data(0).data0.line_no(1)
tl(2) = 0
Tpara(0) = two_line_value(no%).data(0).data0.para(0)
Tpara(1) = two_line_value(no%).data(0).data0.para(1)
Tpara(2) = "0"
v$ = two_line_value(no%).data(0).data0.value
ElseIf ty = line3_value_ Then
tp(0) = line3_value(no%).data(0).data0.poi(0)
tp(1) = line3_value(no%).data(0).data0.poi(1)
tp(2) = line3_value(no%).data(0).data0.poi(2)
tp(3) = line3_value(no%).data(0).data0.poi(3)
tp(4) = line3_value(no%).data(0).data0.poi(4)
tp(5) = line3_value(no%).data(0).data0.poi(5)
tn(0) = line3_value(no%).data(0).data0.n(0)
tn(1) = line3_value(no%).data(0).data0.n(1)
tn(2) = line3_value(no%).data(0).data0.n(2)
tn(3) = line3_value(no%).data(0).data0.n(3)
tn(4) = line3_value(no%).data(0).data0.n(4)
tn(5) = line3_value(no%).data(0).data0.n(5)
tl(0) = line3_value(no%).data(0).data0.line_no(0)
tl(1) = line3_value(no%).data(0).data0.line_no(1)
tl(2) = line3_value(no%).data(0).data0.line_no(2)
Tpara(0) = line3_value(no%).data(0).data0.para(0)
Tpara(1) = line3_value(no%).data(0).data0.para(1)
Tpara(2) = line3_value(no%).data(0).data0.para(2)
v$ = line3_value(no%).data(0).data0.value
End If
m(0) = k%
m(1) = (k% + 1) Mod 3
m(2) = (k% + 2) Mod 3
p(0) = tp(2 * m(0))
p(1) = tp(2 * m(0) + 1)
p(2) = tp(2 * m(1))
p(3) = tp(2 * m(1) + 1)
p(4) = tp(2 * m(2))
p(5) = tp(2 * m(2) + 1)
n(0) = tn(2 * m(0))
n(1) = tn(2 * m(0) + 1)
n(2) = tn(2 * m(1))
n(3) = tn(2 * m(1) + 1)
n(4) = tn(2 * m(2))
n(5) = tn(2 * m(2) + 1)
l(0) = tl(m(0))
l(1) = tl(m(1))
l(2) = tl(m(2))
para(0) = Tpara(m(0))
para(1) = Tpara(m(1))
para(2) = Tpara(m(2))
End Sub

Public Sub set_record_no_reduce(ByVal ty As Byte, no%, ty1, no1%, no_reduce)
If no% > 0 And no1% > 0 Then
If ty = general_string_ Then
 If ty1 = general_string_ Then
  If general_string(no%).data(0).value <> "" And _
       general_string(no1%).data(0).value <> "" Then
   If no% < no1% Then
    Call set_level_(general_string(no%).record_.no_reduce, 4)
   Else
    Call set_level_(general_string(no1%).record_.no_reduce, 4)
   End If
  End If
 Else
 general_string(no%).record_.no_reduce = no_reduce
 End If
ElseIf ty = line3_value_ Then
line3_value(no%).record_.no_reduce = no_reduce
ElseIf ty = two_line_value_ Then
two_line_value(no%).record_.no_reduce = no_reduce
ElseIf ty = angle3_value_ And ty1 = angle3_value_ Then
 If angle3_value(no%).data(0).data0.type = angle3_value_ And _
      angle3_value(no1%).data(0).data0.type = angle3_value_ Then
     If no% < no1% Then
      Call set_level_(angle3_value(no%).record_.no_reduce, 4)
     Else
      Call set_level_(angle3_value(no1%).record_.no_reduce, 4)
     End If
 ElseIf angle3_value(no%).data(0).data0.type = angle3_value_ Then
      Call set_level_(angle3_value(no%).record_.no_reduce, 4)
 ElseIf angle3_value(no1%).data(0).data0.type = angle3_value_ Then
      Call set_level_(angle3_value(no1%).record_.no_reduce, 4)
 Else
     If angle3_value(no%).data(0).data0.type = Two_angle_value_ And _
         angle3_value(no1%).data(0).data0.type = Two_angle_value_ Then
      If no% < no1% Then
      Call set_level_(angle3_value(no%).record_.no_reduce, 4)
      Else
      Call set_level_(angle3_value(no1%).record_.no_reduce, 4)
      End If
     ElseIf angle3_value(no%).data(0).data0.type = Two_angle_value_ Then
      Call set_level_(angle3_value(no%).record_.no_reduce, 4)
     ElseIf angle3_value(no1%).data(0).data0.type = Two_angle_value_ Then
      Call set_level_(angle3_value(no1%).record_.no_reduce, 4)
     Else
     End If
 End If
End If
End If
End Sub
Public Function solve_equation_for_angle3(ByVal n1%, ByVal n2%, _
     ByVal A0%, ByVal A1%, ByVal A2%, ByVal p0$, _
        ByVal p1$, ByVal p2$, ByVal v1$, ByVal b1%, ByVal b2%, ByVal q0$, _
          ByVal q1$, ByVal q2$, ByVal v2$, re As condition_data_type) As Byte
'两个方程A0=B0
Dim i%, j%, no%
Dim tA(5) As Integer
Dim tA_(2) As Integer
Dim ty As Byte
Dim tv(1) As String
Dim s(5) As String
Dim temp_record As total_record_type
Dim v$
Dim ty1 As Byte
 temp_record.record_data.data0.condition_data = re
'End If
If angle(A0%).data(0).no_reduce > 0 Or angle(A1%).data(0).no_reduce > 0 Or _
      angle(A2%).data(0).no_reduce > 0 Or angle(b1%).data(0).no_reduce > 0 Or _
        angle(b2%).data(0).no_reduce > 0 Then
       Exit Function
End If
If p0$ = q0$ Then
 s(0) = p1$
 s(1) = p2$
 s(2) = time_string("-1", q1$, True, False)
 s(3) = time_string("-1", q2$, True, False)
 tv(0) = v1$
 tv(1) = time_string("-1", v2$, True, False)
ElseIf p0$ = time_string("-1", q0$, True, False) Then
 s(0) = p1$
 s(1) = p2$
 s(2) = q1$
 s(3) = q2$
 tv(0) = v1$
 tv(1) = v2$
Else
 s(0) = time_string(p1$, q0$, True, False)
 s(1) = time_string(p2$, q0$, True, False)
 s(2) = time_string(q1$, p0$, True, False)
 s(3) = time_string(q2$, p0$, True, False)
 s(2) = time_string(s(2), "-1", True, False)
 s(3) = time_string(s(3), "-1", True, False)
 tv(0) = time_string(v1$, q0$, True, False)
 tv(1) = time_string(v2$, p0$, True, False)
 tv(1) = time_string("-1", tv(1), True, False)
End If
 tA(0) = A1%
 tA(1) = A2%
 tA(2) = b1%
 tA(3) = b2%

 v$ = add_string(tv(0), tv(1), True, False)
 'End If
 For i% = 0 To 3
  If s(i%) = "0" Then
     tA(i%) = 0
  End If
  Next i%
 For i = 0 To 1
 For j% = 2 To 3
  If tA(i%) = tA(j%) And tA(i%) > 0 Then
   s(i%) = add_string(s(i%), s(j%), True, False)
     tA(j%) = 0
      s(j%) = "0"
  End If
 Next j%
Next i%
Call remove_record_for_zero_para(s(), tA(), 4)
If s(3) = "0" Then
If s(0) = "0" And s(1) = "0" And s(2) = "0" Then
 Exit Function
Else
no% = 0
'last_angle3_value = last_angle3_value + 1
 ' ReDim Preserve angle3_value(last_angle3_value) As angle3_value_type
solve_equation_for_angle3 = set_three_angle_value(tA(0), _
    tA(1), tA(2), s(0), s(1), s(2), v$, 0, temp_record, _
     no%, 0, 0, False, 1, 0, False)
If solve_equation_for_angle3 > 1 Then
 Exit Function
End If
If n1% > n2% And n2% > 0 Then
 If angle3_value(n1%).data(0).data0.type = angle3_value_ Or angle3_value(n1%).data(0).data0.type = Two_angle_value_ Then
   Call set_level_(angle3_value(n1%).record_.no_reduce, 4)
 End If
ElseIf n1% < n2% And n1% > 0 Then
 If angle3_value(n2%).data(0).data0.type = angle3_value_ Or angle3_value(n2%).data(0).data0.type = Two_angle_value_ Then
  Call set_level_(angle3_value(n2%).record_.no_reduce, 4)
 End If
End If
End If
End If
 s(0) = p0$
 s(1) = p1$
 s(2) = p2$
 s(3) = q0$
 s(4) = q1$
 s(5) = q2$
 tv(0) = v1$
 tv(1) = v2$
 tA(0) = A0%
 tA(1) = A1%
 tA(2) = A2%
 tA(3) = A0%
 tA(4) = b1%
 tA(5) = b2%
 If (s(1) = "0" And s(2) = "0") Or s(1) <> "0" And s(2) <> "0" Then
  Exit Function
 ElseIf s(1) = "0" Then
  Call exchange_string(s(1), s(2))
  Call exchange_two_integer(tA(1), tA(2))
 End If
 If (s(4) = "0" And s(5) = "0") Or s(4) <> "0" And s(5) <> "0" Then
  Exit Function
 ElseIf s(4) = "0" Then
  Call exchange_string(s(4), s(5))
  Call exchange_two_integer(tA(4), tA(5))
 End If
 If combine_two_angle(tA(1), tA(4), tA_(0), 0, 0, tA_(1), 0, tA_(2), ty, 0, 0) Then
 s(0) = divide_string(s(0), s(1), True, False)
 tv(0) = divide_string(tv(0), s(1), True, False)
 s(1) = "1"
 s(3) = divide_string(s(3), s(4), True, False)
 tv(1) = divide_string(tv(1), s(4), True, False)
 s(4) = "1"
 If ty = 3 Or ty = 5 Then
  s(0) = add_string(s(0), s(3), True, False)
  tv(0) = add_string(tv(0), tv(1), True, False)
  If s(0) = "0" Then
  tA(0) = 0
  End If
  solve_equation_for_angle3 = set_three_angle_value(tA(0), tA_(2), 0, s(0), s(1), "0", tv(0), _
       0, temp_record, 0, 0, 0, 0, 0, 0, False)
  If solve_equation_for_angle3 > 1 Then
   Exit Function
  End If
 ElseIf ty = 4 Or ty = 8 Then
  If ty = 4 Then
  tA_(2) = tA_(0)
  ElseIf ty = 8 Then
  tA_(2) = tA_(1)
  End If
  s(0) = minus_string(s(0), s(3), True, False)
  tv(0) = minus_string(tv(0), tv(1), True, False)
  If s(0) = "0" Then
  tA(0) = 0
  End If
  solve_equation_for_angle3 = set_three_angle_value(tA(0), tA_(2), 0, s(0), s(1), "0", tv(0), _
       0, temp_record, 0, 0, 0, 0, 0, 0, False)
  If solve_equation_for_angle3 > 1 Then
   Exit Function
  End If
 ElseIf ty = 6 Or ty = 7 Then
  If ty = 6 Then
  tA_(2) = tA_(0)
  ElseIf ty = 7 Then
  tA_(2) = tA_(1)
  End If
  s(0) = minus_string(s(3), s(0), True, False)
  tv(0) = minus_string(tv(1), tv(0), True, False)
  If s(0) = "0" Then
  tA(0) = 0
  End If
  solve_equation_for_angle3 = set_three_angle_value(tA(0), tA_(2), 0, s(0), s(1), "0", tv(0), _
       0, temp_record, 0, 0, 0, 0, 0, 0, False)
  If solve_equation_for_angle3 > 1 Then
   Exit Function
  End If
 End If
 End If
End Function
Public Function add_new_item_to_item(ByVal it1%, ByVal it2%, _
         ByVal p1$, ByVal p2$, ByVal it0%, re As condition_data_type) As Byte
Dim i%, no%
'On Error GoTo add_new_item_to_item_error
If (it1% > it2% Or it1% = 0) And it2% > 0 Then
 Call exchange_two_integer(it1%, it2%)
 Call exchange_string(p1$, p2$)
End If
For i% = 1 To item0(it0%).data(0).record_for_trans.last_trans_to
 If item0(it0%).data(0).record_for_trans.record(i%).to_no(0) = it1% And _
       item0(it0%).data(0).record_for_trans.record(i%).to_no(1) = it2% Then
      add_new_item_to_item = 0
       Exit Function
 End If
Next i%
item0(it0%).data(0).record_for_trans.last_trans_to = _
    item0(it0%).data(0).record_for_trans.last_trans_to + 1
no% = item0(it0%).data(0).record_for_trans.last_trans_to
ReDim Preserve item0(it0%).data(0).record_for_trans.record(no%) As record_type0
 item0(it0%).data(0).record_for_trans.record(no%).condition_data = re
 item0(it0%).data(0).record_for_trans.record(no%).to_no(0) = it1%
 item0(it0%).data(0).record_for_trans.record(no%).to_no(1) = it2%
 item0(it0%).data(0).record_for_trans.record(no%).para(0) = p1$
 item0(it0%).data(0).record_for_trans.record(no%).para(1) = p2$
 add_new_item_to_item = combine_item_with_general_string_(it0%, _
      item0(it0%).data(0).record_for_trans.last_trans_to)
Exit Function
add_new_item_to_item_error:
End Function
Public Function theorem1(no_reduce As Byte) As Byte
Dim i%, tn%
Dim last(6) As Integer
last(0) = last_conditions.last_cond(0).last_general_string_combine
last_conditions.last_cond(0).last_general_string_combine = last_conditions.last_cond(1).general_string_no
For i% = last(0) + 1 To last_conditions.last_cond(0).last_general_string_combine
 If (general_string(i%).data(0).value = "" And general_string(i%).data(0).record.data0.condition_data.level > 9) Or _
      (general_string(i%).data(0).value <> "" And general_string(i%).data(0).record.data0.condition_data.level > 4) Then
theorem1 = combine_general_string_with_general_string(i%, no_reduce)
If theorem1 > 1 Then
 Exit Function
End If
theorem1 = combine_general_string_with_item(i%, no_reduce)
If theorem1 > 1 Then
 Exit Function
End If
End If
Next i%
last(0) = last_conditions.last_cond(0).last_angle3_value_combine
last_conditions.last_cond(0).last_angle3_value_combine = last_conditions.last_cond(1).angle3_value_no
For i% = last(0) + 1 To last_conditions.last_cond(0).last_angle3_value_combine
 If angle3_value(i%).record_.no_reduce < 2 And angle3_value(i%).data(0).data0.reduce = True Then
  If angle3_value(i%).data(0).record.data0.condition_data.level >= 4 Then
     Call set_level_(angle3_value(i%).record_.no_reduce, 2)
    If angle3_value(i%).data(0).data0.type = angle_value_ Or _
      angle3_value(i%).data(0).data0.type = eangle_ Or _
       angle3_value(i%).data(0).data0.type = angle_relation_ Or _
        angle3_value(i%).data(0).data0.type = two_angle_value_sum_ Then
   'no_reduce = 0
 theorem1 = combine_three_angle_with_three_angle(i%, no_reduce)
   If theorem1 > 1 Then
    Exit Function
   End If
   End If
 End If
 End If
Next i%
End Function

Public Sub set_level_(level As Byte, ByVal add_l As Byte)
If add_l = 1 Then
 If level = 0 Then
  level = 1
 ElseIf level = 2 Then
  level = 3
 ElseIf level = 4 Then
  level = 5
 End If
ElseIf add_l = 2 Then
 If level < 2 Then
  level = level + 2
 End If
ElseIf add_l = 4 Then
 If level < 4 Then
  level = level + 4
 End If
End If
End Sub

Public Function th_of_mid_point_for_triangle(ByVal no%, _
                      triA_ As triangle_data0_type, _
                       ByVal md1%, ByVal md2%, no_reduce As Byte) As Byte
Dim tl(1) As Integer
Dim tn(3) As Integer
Dim n_(1) As Integer
Dim md3%
Dim i%, j%, tp%
Dim temp_record As total_record_type
Dim triA As triangle_data0_type
triA = triA_
md3% = last_number_for_3number(md1%, md2%)
tl(0) = line_number0(triA.poi(md1%), triA.poi(md2%), tn(0), tn(1))
    temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
    temp_record.record_data.data0.condition_data.condition(1).no = triA.midpoint_no(md1%)
 If triA.midpoint_no(md2%) > 0 Then
   If th_chose(97).chose = 1 Then
    tl(1) = line_number0(Dmid_point(triA.midpoint_no(md1%)).data(0).data0.poi(1), _
        Dmid_point(triA.midpoint_no(md2%)).data(0).data0.poi(1), tn(2), tn(3))
    temp_record.record_data.data0.condition_data.condition_no = 2
    temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.midpoint_no(md2%)
    temp_record.record_data.data0.theorem_no = 97
     th_of_mid_point_for_triangle = set_dparal(tl(0), _
      tl(1), temp_record, 0, no_reduce, False)
     If th_of_mid_point_for_triangle > 1 Then
      Exit Function
     End If
     th_of_mid_point_for_triangle = set_Drelation(triA.poi(md1%), _
      triA.poi(md2%), Dmid_point(triA.midpoint_no(md1%)).data(0).data0.poi(1), _
       Dmid_point(triA.midpoint_no(md2%)).data(0).data0.poi(1), tn(0), tn(1), _
        tn(2), tn(3), tl(0), tl(1), "2", temp_record, 0, 0, 0, 0, no_reduce, False)
     If th_of_mid_point_for_triangle > 1 Then
      Exit Function
     End If
    End If
 Else
   If th_chose(100).chose = 1 Then
    temp_record.record_data.data0.theorem_no = 100
    'If md1% > 0 Then
      tl(1) = line_number0(triA.poi(md1%), triA.poi(md3%), tn(0), tn(2))
     ' tp(0) = p0%
      ' tp(2) = p2%
       ' tp(3) = Dmid_point(md1%).data(0).poi(1)
        ' temp_record.data0.condition_data.condition(1).ty = midpoint_
         '  temp_record.data0.condition_data.condition(1).no = TriA.midpoint_no(md1%)
    'ElseIf md2% > 0 Then
     ' tl(1) = line_number0(p0%, p1%, tn(0), tn(2))
      'tp(0) = p0%
       'tp(2) = p1%
        'tp(3) = Dmid_point(TriA.midpoint_no(md2%)).data(0).poi(1)
         'temp_record.data0.condition_data.condition(1).ty = midpoint_
          'temp_record.data0.condition_data.condition(1).no = TriA.midpoint_no(md1%)
    'Else
     'Exit Function
    'End If
    For i% = 1 To last_conditions.last_cond(1).paral_no
     If Dparal(i%).data(0).data0.line_no(0) = tl(0) Then
        If is_point_in_line3( _
         Dmid_point(triA.midpoint_no(md1%)).data(0).data0.poi(1), _
          m_lin(Dparal(i%).data(0).data0.line_no(1)).data(0).data0, 0) Then
           tp% = is_line_line_intersect(tl(1), _
             Dparal(i%).data(0).data0.line_no(1), tn(1), 0, False)
             If tp% > 0 Then
              temp_record.record_data.data0.condition_data.condition(2).ty = paral_
               temp_record.record_data.data0.condition_data.condition(2).no = i%
                temp_record.record_data.data0.condition_data.condition_no = 2
               GoTo th_of_mid_point_for_triangle_mark1
            End If
        End If
     ElseIf Dparal(i%).data(0).data0.line_no(1) = tl(0) Then
        If is_point_in_line3( _
         Dmid_point(triA.midpoint_no(md1%)).data(0).data0.poi(1), _
           m_lin(Dparal(i%).data(0).data0.line_no(0)).data(0).data0, 0) Then
         tp% = is_line_line_intersect(tl(1), _
            Dparal(i%).data(0).data0.line_no(0), tn(1), 0, False)
           If tp% > 0 Then
             temp_record.record_data.data0.condition_data.condition(2).ty = paral_
              temp_record.record_data.data0.condition_data.condition(2).no = i%
               temp_record.record_data.data0.condition_data.condition_no = 2
             GoTo th_of_mid_point_for_triangle_mark1
            End If
        End If
     End If
    Next i%
     Exit Function
th_of_mid_point_for_triangle_mark1:
th_of_mid_point_for_triangle = set_mid_point( _
     triA.poi(md1%), tp%, triA.poi(md3%), _
     tn(0), tn(1), tn(2), tl(1), 0, temp_record, 0, 0, 0, 0, no_reduce)
     If th_of_mid_point_for_triangle > 1 Then
      Exit Function
     End If
  End If
 End If
End Function
Public Function Th_sin(ByVal triA%, tri As triangle_data0_type, _
                         ByVal A1%, ByVal A2%, ByVal A3%, _
                          no_reduce As Byte) As Byte
Dim ts(1) As String
Dim n(2) As Integer
Dim triangle_data As triangle_data0_type
Dim temp_record As total_record_type
triangle_data = tri
If triangle_data.right_angle_no >= 0 Then
  If A1% = triangle_data.right_angle_no Then
   Th_sin = solve_right_triangle(triA%, triangle_data, _
         triangle_data.right_angle_no, _
          (triangle_data.right_angle_no + 1) Mod 3, 0, False, no_reduce)
  Else
   Th_sin = solve_right_triangle(triA%, triangle_data, _
         triangle_data.right_angle_no, _
          (triangle_data.right_angle_no + 1) Mod 3, 1, False, no_reduce)
  End If
Else
 If is_equal_angle(triangle_data.angle(A1%), triangle_data.angle(A2%), _
        n(0), n(1)) Then
  If th_chose(40).chose = 1 Then
   temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
    Call add_conditions_to_record(angle3_value_, n(0), n(1), 0, temp_record.record_data.data0.condition_data)
     temp_record.record_data.data0.theorem_no = 40
      Th_sin = set_equal_dline(triangle_data.poi(A3%), triangle_data.poi(A1%), _
       triangle_data.poi(A3%), triangle_data.poi(A2%), 0, 0, 0, 0, 0, 0, _
        0, temp_record, 0, 0, 0, 0, no_reduce, False)
         If Th_sin > 1 Then
          Exit Function
         End If
  End If
Else
 If th_chose(153).chose = 1 Then
    temp_record.record_data.data0.condition_data.condition_no = 1
     temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
      temp_record.record_data.data0.condition_data.condition(1).no = _
        angle(triangle_data.angle(A1%)).data(0).value_no
    Call add_conditions_to_record(angle3_value_, angle(triangle_data.angle(A2%)).data(0).value_no, _
           0, 0, temp_record.record_data.data0.condition_data)
    temp_record.record_data.data0.theorem_no = 153
    ts(0) = sin_(angle(triangle_data.angle(A1%)).data(0).value, 0)
    ts(1) = sin_(angle(triangle_data.angle(A2%)).data(0).value, 0)
  If InStr(1, ts(0), "F", 0) = 0 And InStr(1, ts(1), "F", 0) = 0 Then
   If triangle_data.line_value(A2%) > 0 And _
               triangle_data.line_value(A1%) = 0 Then
    Call add_conditions_to_record(line_value_, triangle_data.line_value(A2%), 0, 0, temp_record.record_data.data0.condition_data)
    ts(0) = divide_string(time_string( _
              line_value(triangle_data.line_value(A2%)).data(0).data0.value, _
               ts(0), False, False), ts(1), True, False)
    Th_sin = set_line_value(triangle_data.poi(A3%), _
       triangle_data.poi(A2%), ts(0), 0, 0, 0, temp_record.record_data, _
        0, no_reduce, False)
     If Th_sin > 1 Then
      Exit Function
     End If
   ElseIf triangle_data.line_value(A1%) > 0 And _
                       triangle_data.line_value(A2%) = 0 Then
    Call add_conditions_to_record(line_value_, triangle_data.line_value(A1%), 0, 0, temp_record.record_data.data0.condition_data)
    ts(0) = divide_string(time_string( _
              line_value(triangle_data.line_value(A1%)).data(0).data0.value, _
                ts(1), False, False), ts(0), True, False)
    Th_sin = set_line_value(triangle_data.poi(A3%), _
       triangle_data.poi(A1%), ts(0), 0, 0, 0, temp_record.record_data, _
        0, no_reduce, False)
     If Th_sin > 1 Then
      Exit Function
     End If
   ElseIf triangle_data.relation_no(A3%, 0).ty = 0 Then
    Th_sin = set_Drelation(triangle_data.poi(A3%), _
        triangle_data.poi(A1%), triangle_data.poi(A3%), _
         triangle_data.poi(A2%), 0, 0, 0, 0, 0, 0, _
          divide_string(ts(1), ts(0), True, False), temp_record, 0, _
            0, 0, 0, no_reduce, False)
    If Th_sin > 1 Then
     Exit Function
    End If
   End If
  End If
 End If
End If
End If
End Function

Public Function Th_area_of_triangle1(triA%, tri As triangle_data0_type, _
             ByVal p_k1%, ByVal p_k2%, ByVal p_k3%, no_reduce) _
              As Byte
Dim n(1) As Integer
Dim tri_data As triangle_data0_type
Dim temp_record As total_record_type
tri_data = tri
If th_chose(2).chose = 1 Then
'  If tri_data.right_angle_no = p_k2% Then
     
'  ElseIf tri_data.right_angle_no = p_k3% Then
  
'  Else 'If tri_data.right_angle_no >= 0 Then
 If tri_data.verti_no(p_k1%) > 0 Then
  If Dverti(tri_data.verti_no(p_k1%)).data(0).inter_poi > 0 Then '垂足
   temp_record.record_data.data0.theorem_no = 2
    If tri_data.area_no = 0 Then '不知道面积
      If tri_data.line_value(p_k1%) > 0 Then '底边
        If is_line_value(tri_data.poi(p_k1%), _
         Dverti(tri_data.verti_no(p_k1%)).data(0).inter_poi, _
          0, 0, 0, "", n(0), -1000, 0, 0, 0, line_value_data0) = 1 Then '高
         If tri_data.line_value(p_k1%) > 0 Then
         temp_record.record_data.data0.condition_data.condition_no = 2
         temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
         temp_record.record_data.data0.condition_data.condition(1).no = n(0)
         temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
         temp_record.record_data.data0.condition_data.condition(2).no = tri_data.line_value(p_k1%)
         n(1) = 0
       Th_area_of_triangle1 = set_area_of_triangle(triA%, _
          divide_string(time_string( _
           line_value(n(0)).data(0).data0.value, _
            line_value(tri_data.line_value(p_k1%)).data(0).data0.value, False, False), _
             "2", True, False), temp_record, n(1), no_reduce)
       triangle(triA%).data(0).area_no = n(1)
       If Th_area_of_triangle1 > 1 Then
        Exit Function
       End If
      End If
    End If
   End If
 Else '知道面积
   If is_line_value(tri_data.poi(p_k1%), _
       Dverti(tri_data.verti_no(p_k1%)).data(0).inter_poi, _
        0, 0, 0, "", n(0), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then '高
       If tri_data.line_value(p_k1%) = 0 Then
       temp_record.record_data.data0.condition_data.condition_no = 2
       temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
       temp_record.record_data.data0.condition_data.condition(1).no = n(0)
       temp_record.record_data.data0.condition_data.condition(2).ty = area_of_element_
       temp_record.record_data.data0.condition_data.condition(2).no = tri_data.area_no
       Th_area_of_triangle1 = set_line_value( _
         tri_data.poi(p_k2%), tri_data.poi(p_k3%), _
          divide_string(time_string("2", _
            area_of_element(tri_data.area_no).data(0).value, False, False), _
             line_value(n(0)).data(0).data0.value, True, False), _
              0, 0, 0, temp_record.record_data, 0, no_reduce, False) '底边
        If Th_area_of_triangle1 > 1 Then
         Exit Function
        End If
       End If
   ElseIf tri_data.line_value(p_k1%) > 0 Then '底边
       temp_record.record_data.data0.condition_data.condition_no = 2
       temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
       temp_record.record_data.data0.condition_data.condition(1).no = tri_data.line_value(p_k1%)
       temp_record.record_data.data0.condition_data.condition(2).ty = area_of_element_
       temp_record.record_data.data0.condition_data.condition(2).no = tri_data.area_no
       Th_area_of_triangle1 = set_line_value( _
         tri_data.poi(p_k1%), Dverti(tri_data.verti_no(p_k1%)).data(0).inter_poi, _
          divide_string(time_string("2", _
            area_of_element(tri_data.area_no).data(0).value, False, False), _
             line_value(tri_data.line_value(p_k1%)).data(0).data0.value, True, False), _
              0, 0, 0, temp_record.record_data, 0, no_reduce, False) '高
        If Th_area_of_triangle1 > 1 Then
         Exit Function
        End If
   End If
End If
End If
End If
End If
End Function
Public Function th_area_of_triangle3(tri_ As triangle_data0_type, _
                 ByVal triA%, ByVal no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim n1%, n2%, n3%, i%, j%
Dim p As String
Dim s As String
Dim ts As String
Dim tv(2)  As String
Dim tp(3) As Integer
Dim tl%
Dim tA%
Dim no%
Dim tri_f As tri_function_data_type
Dim triangle_data As triangle_data0_type
Dim tri As triangle_data0_type
tri = tri_
If triA% = 0 Then
 triA% = triangle_number_(tri)
ElseIf triA% > 0 Then
 tri = triangle(triA%).data(0)
Else
 Exit Function
End If
triangle_data = tri
n1% = triangle_data.line_value(0)
n2% = triangle_data.line_value(1)
n3% = triangle_data.line_value(2)
If n1% = 0 Or n2% = 0 Or n3% = 0 Then
 Exit Function
End If
temp_record.record_data.data0.condition_data.condition_no = 3
 temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
 temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
 temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
 temp_record.record_data.data0.condition_data.condition(1).no = n1%
 temp_record.record_data.data0.condition_data.condition(2).no = n2%
 temp_record.record_data.data0.condition_data.condition(3).no = n3%
 tv(0) = add_string(line_value(n2%).data(0).data0.squar_value, _
            line_value(n3%).data(0).data0.squar_value, True, False)
 tv(1) = add_string(line_value(n1%).data(0).data0.squar_value, _
            line_value(n3%).data(0).data0.squar_value, True, False)
 tv(2) = add_string(line_value(n2%).data(0).data0.squar_value, _
            line_value(n1%).data(0).data0.squar_value, True, False)
 tv(0) = minus_string(tv(0), _
            line_value(n1%).data(0).data0.squar_value, True, False)
 tv(1) = minus_string(tv(1), _
            line_value(n2%).data(0).data0.squar_value, True, False)
 tv(2) = minus_string(tv(2), _
            line_value(n3%).data(0).data0.squar_value, True, False)
 If tv(0) = "0" Then
   temp_record.record_data.data0.theorem_no = 51
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(0), 0, 0, "1", _
      "0", "0", "90", 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
 ElseIf tv(1) = "0" Then
   temp_record.record_data.data0.theorem_no = 51
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(1), 0, 0, "1", _
      "0", "0", "90", 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
 ElseIf tv(2) = "0" Then
   temp_record.record_data.data0.theorem_no = 51
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(2), 0, 0, "1", _
      "0", "0", "90", 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
 Else
  tv(0) = divide_string(tv(0), line_value(n2%).data(0).data0.value, False, False)
  tv(0) = divide_string(tv(0), line_value(n3%).data(0).data0.value, False, False)
  tv(0) = divide_string(tv(0), "2", True, False)
  p = accos_(tv(0))
  If p <> "F" Then
   temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(0), 0, 0, "1", _
      "0", "0", p, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  Else
   temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_tri_function(triangle_data.angle(0), "", tv(0), "", "", _
      0, temp_record, False, tri_f, 0)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  End If
  tv(1) = divide_string(tv(1), line_value(n1%).data(0).data0.value, False, False)
  tv(1) = divide_string(tv(1), line_value(n3%).data(0).data0.value, False, False)
  tv(1) = divide_string(tv(1), "2", True, False)
  p = accos_(tv(1))
  If p <> "F" Then
   temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(1), 0, 0, "1", _
      "0", "0", p, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  Else
     temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_tri_function(triangle_data.angle(1), "", tv(1), "", "", _
      0, temp_record, False, tri_f, 0)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  End If
  tv(2) = divide_string(tv(2), line_value(n2%).data(0).data0.value, False, False)
  tv(2) = divide_string(tv(2), line_value(n1%).data(0).data0.value, False, False)
  tv(2) = divide_string(tv(2), "2", True, False)
   p = accos_(tv(2))
  If p <> "F" Then
   temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_three_angle_value(triangle_data.angle(2), 0, 0, "1", _
      "0", "0", p, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  Else
     temp_record.record_data.data0.theorem_no = 154
   th_area_of_triangle3 = set_tri_function(triangle_data.angle(2), "", tv(2), "", "", _
      0, temp_record, False, tri_f, 0)
      If th_area_of_triangle3 > 1 Then
       Exit Function
      End If
  End If
 End If
               
'If triA% = 0 Or th_chose(156).chose = 0 Or no_reduce = 255 Or _
'         triangle_data.area_no > 0 Then
' Exit Function
'End If
 temp_record.record_data.data0.theorem_no = 156
If is_area_of_triangle(triA%, no%) Then
 If InStr(1, area_of_element(no%).data(0).value, ".", 0) = 0 Then
  Call set_level(temp_record.record_data.data0.condition_data)
   Call set_prove_type(area_of_element_, no%, temp_record.record_data, _
     area_of_element(no%).data(0).record)
  Exit Function
 End If
End If
p = add_string(line_value(n1%).data(0).data0.value, line_value(n2%).data(0).data0.value, True, False)
p = add_string(p, line_value(n3%).data(0).data0.value, True, False)
p = divide_string(p, "2", True, False)
s = time_string(p, minus_string(p, line_value(n1%).data(0).data0.value, False, False), True, False)
ts = sqr_string(s, True, False)
If InStr(1, ts, "F", 0) > 0 Then
 th_area_of_triangle3 = 0
  Exit Function
End If
s = time_string(minus_string(p, line_value(n2%).data(0).data0.value, False, False), _
        minus_string(p, line_value(n3%).data(0).data0.value, False, False), True, False)
s = sqr_string(s, True, False)
If InStr(1, s, "F", 0) > 0 Then
 th_area_of_triangle3 = 0
  Exit Function
End If
s = time_string(s, ts, True, False)
If s = "0" Then
 Exit Function
End If
th_area_of_triangle3 = set_area_of_triangle(triA%, _
   s, temp_record, triangle(triA%).data(0).area_no, no_reduce)
                   If th_area_of_triangle3 > 1 Then
                      Exit Function
                   End If
For i% = 0 To 2
tp(0) = triangle(triA%).data(0).poi(i%)
tp(1) = triangle(triA%).data(0).poi((i% + 1) Mod 3)
tp(2) = triangle(triA%).data(0).poi((i% + 2) Mod 3)
tl = line_number0(tp(1), tp(2), 0, 0)
 For j% = 1 To m_lin(tl%).data(0).data0.in_point(0)
  tp(3) = m_lin(tl%).data(0).data0.in_point(j%)
   If tp(3) > 0 Then
      If tp(2) <> tp(3) And tp(1) <> tp(3) Then
         tA% = Abs(angle_number(tp(0), tp(3), tp(1), "", 0))
         If angle(tA%).data(0).value = "30" Or angle(tA%).data(0).value = "45" Or _
             angle(tA%).data(0).value = "60" Or angle(tA%).data(0).value = "120" Or _
              angle(tA%).data(0).value = "135" Or angle(tA%).data(0).value = "150" Then
            tp(1) = tl%
            tp(2) = tl%
            tp(3) = 0
             Call C_wait_for_aid_point.set_wait_for_aid_point(verti_, tp(), 3)
              'If using_area_th = 0 Then
              'using_area_th = 1
              ' th_chose(156).chose = 1
                  th_area_of_triangle3 = set_area_of_triangle(triA%, _
                    s, temp_record, triangle(triA%).data(0).area_no, 255)
                   If th_area_of_triangle3 > 1 Then
                      Exit Function
                   End If
              ' End If
              GoTo th_area_of_triangle3_out
         End If
      End If
   End If
 Next j%
th_area_of_triangle3_out:
Next i%
End Function

Public Function Th_area_of_triangle2(ByVal triA%, _
         tri As triangle_data0_type, ByVal k1%, ByVal k2%, _
          ByVal k3%, ByVal no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim s As String
Dim ts$
Dim no%
Dim s1%, S2%, n%
Dim tri_data As triangle_data0_type
tri_data = tri
s1% = tri_data.line_value(k1%)
S2% = tri_data.line_value(k2%)
If triA% = 0 Or no_reduce = 255 Then
 Exit Function
End If
If tri_data.area_no = 0 Then
   If s1% > 0 And S2% > 0 And angle(tri_data.angle(k3%)).data(0).value <> "" Then
    temp_record.record_data.data0.condition_data.condition_no = 2
    temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(1).no = s1%
    temp_record.record_data.data0.condition_data.condition(2).no = S2%
    Call add_conditions_to_record(angle3_value_, angle(tri_data.angle(k3%)).data(0).value_no, _
           0, 0, temp_record.record_data.data0.condition_data)
   If tri_data.right_angle_no <> k1% Then
    If th_chose(155).chose = 1 Then
    temp_record.record_data.data0.theorem_no = 155
    s = time_string(line_value(s1%).data(0).data0.value, line_value(S2%).data(0).data0.value, True, False)
    ts$ = sin_(angle(tri_data.angle(k3%)).data(0).value, 0)
     If InStr(1, ts$, "F", 0) > 0 Then
      Exit Function
     End If
    s = time_string(s, sin_(angle(tri_data.angle(k3%)).data(0).value, 0), True, False)
    s = divide_string(s, "2", True, False)
    Th_area_of_triangle2 = set_area_of_triangle(triA%, _
     s, temp_record, n%, no_reduce)
     triangle(triA%).data(0).area_no = n%
     If Th_area_of_triangle2 > 1 Then
      Exit Function
     End If
     End If
    Else
     If th_chose(20).chose = 1 Then
      temp_record.record_data.data0.theorem_no = 20
      s = time_string(line_value(s1%).data(0).data0.value, line_value(S2%).data(0).data0.value, True, False)
      s = divide_string(s, "2", True, False)
       Th_area_of_triangle2 = set_area_of_triangle(triA%, _
         s, temp_record, n%, no_reduce)
          triangle(triA%).data(0).area_no = n%
       If Th_area_of_triangle2 > 1 Then
        Exit Function
       End If
     End If
    End If
    End If
 Else 'area=0
  If s1% = 0 And S2% > 0 And angle(tri_data.angle(k3%)).data(0).value <> "" Then
   temp_record.record_data.data0.condition_data.condition_no = 2
   temp_record.record_data.data0.condition_data.condition(1).ty = area_of_element_
   temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
   temp_record.record_data.data0.condition_data.condition(1).no = tri_data.area_no
   temp_record.record_data.data0.condition_data.condition(2).no = S2%
   Call add_conditions_to_record(angle3_value_, angle(tri_data.angle(k3%)).data(0).value_no, _
          0, 0, temp_record.record_data.data0.condition_data)
   If tri_data.right_angle_no <> k1% Then
    If th_chose(155).chose = 1 Then
     temp_record.record_data.data0.theorem_no = 155
      ts = sin_(angle(tri_data.angle(k3%)).data(0).value, 0)
     If InStr(1, ts, "F", 0) = 0 Then
       s = time_string(line_value(S2).data(0).data0.value, ts, True, False)
        s = divide_string(time_string("2", _
           area_of_element(tri_data.area_no).data(0).value, False, False), s, True, False)
         Th_area_of_triangle2 = set_line_value(tri_data.poi(k2%), _
           tri_data.poi(k3%), s, 0, 0, 0, temp_record.record_data, _
         0, no_reduce, False)
        If Th_area_of_triangle2 > 1 Then
         Exit Function
        End If
       End If
     End If 'If TH_CHOSE(155).chose = 1 Then
  Else 'tri_data.right_angle_no <> k1% Then
   If th_chose(20).chose = 1 Then
    temp_record.record_data.data0.theorem_no = 20
     s = divide_string(time_string("2", _
       area_of_element(tri_data.area_no).data(0).value, False, False), _
          line_value(S2).data(0).data0.value, True, False)
     Th_area_of_triangle2 = set_line_value(tri_data.poi(k2%), _
      tri_data.poi(k3%), s, 0, 0, 0, temp_record.record_data, _
       0, no_reduce, False)
      If Th_area_of_triangle2 > 1 Then
       Exit Function
      End If
   End If
  End If
 ElseIf s1% > 0 And S2% = 0 And angle(tri_data.angle(k3%)).data(0).value <> "" Then
  Th_area_of_triangle2 = Th_area_of_triangle2(no%, tri_data, _
      k1%, k3%, k2%, no_reduce)
  If Th_area_of_triangle2 > 1 Then
   Exit Function
  End If
 End If
End If
End Function
Public Function th_cos_(ByVal no%, tri As triangle_data0_type, _
          k%, no_reduce As Byte, cal_float As Boolean) As Byte
Dim ts(2) As String
Dim ts_ As String
Dim tn(2) As Integer
Dim tri_f As tri_function_data_type
Dim tri_data As triangle_data0_type
Dim temp_record As total_record_type
Dim temp_record_data As record_data_type
Dim it(1) As Integer
Dim cond_data As condition_data_type
tn(0) = k%
tn(1) = (k% + 1) Mod 3
tn(2) = (k% + 2) Mod 3
tri_data = tri
If tri_data.line_value(tn(0)) > 0 And _
     tri_data.line_value(tn(1)) > 0 And _
       tri_data.line_value(tn(2)) > 0 Then
'已知三边长
temp_record.record_data.data0.condition_data.condition_no = 3
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(1).no = tri_data.line_value(tn(0))
temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).no = tri_data.line_value(tn(1))
temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
temp_record.record_data.data0.condition_data.condition(3).no = tri_data.line_value(tn(2))
'计算勾股差
If is_x_in_string(line_value(tri_data.line_value(tn(0))).data(0).data0.value) = 1 Or _
    is_x_in_string(line_value(tri_data.line_value(tn(1))).data(0).data0.value) = 1 Or _
     is_x_in_string(line_value(tri_data.line_value(tn(2))).data(0).data0.value) = 1 Then
   If angle(tri_data.angle(tn(0))).data(0).value <> "" Or tri_data.tri_function(tn(0)) Then
    If tri_data.tri_function(tn(0)) = 0 Then
     ts_ = cos_(angle(tri_data.angle(tn(0))).data(0).value, 0)
     If InStr(1, ts_, "F", 0) = 0 Then
     If tri_data.tri_function(tn(0)) = 0 Then
     th_cos_ = set_tri_function(tri_data.angle(tn(0)), "", ts_, "", "", tri_data.tri_function(tn(0)), _
            temp_record, False, tri_f, 0)
     End If
     If th_cos_ > 1 Then
       Exit Function
     End If
     triangle(no%).data(0).tri_function(tn(0)) = tri_data.tri_function(tn(0))
     End If
    Else
    ts_ = tri_function(tri_data.tri_function(tn(0))).data(0).cos_value
    End If
     ts(0) = add_string(line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, _
             line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, False, False)
     ts(0) = minus_string(ts(0), _
             line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, False, False)
     ts(1) = time_string(line_value(tri_data.line_value(tn(1))).data(0).data0.value, _
             line_value(tri_data.line_value(tn(2))).data(0).data0.value, False, False)
     ts(1) = time_string(ts(1), "2", False, False)
     ts(1) = time_string(ts(1), ts_, False, False)
     ts(0) = minus_string(ts(0), ts(1), True, False)
      If InStr(1, ts(0), "F", 0) = 0 Then
      th_cos_ = set_equation(ts(0), 0, temp_record)
      If th_cos_ > 1 Then
       Exit Function
      End If
     End If
   ElseIf angle(tri_data.angle(tn(1))).data(0).value <> "" Or tri_data.tri_function(tn(1)) Then
    If tri_data.tri_function(tn(1)) = 0 Then
     ts_ = cos_(angle(tri_data.angle(tn(1))).data(0).value, 0)
     If InStr(1, ts_, "F", 0) = 0 Then
     If tri_data.tri_function(tn(1)) = 0 Then
     th_cos_ = set_tri_function(tri_data.angle(tn(1)), "", ts_, "", "", tri_data.tri_function(tn(1)), _
            temp_record, False, tri_f, 0)
     End If
     If th_cos_ > 1 Then
       Exit Function
     End If
     triangle(no%).data(0).tri_function(tn(1)) = tri_data.tri_function(tn(1))
     End If
    Else
     ts_ = tri_function(tri_data.tri_function(tn(1))).data(0).cos_value
    End If
     ts(0) = add_string(line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
             line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, False, False)
     ts(0) = minus_string(ts(0), _
             line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, False, False)
     ts(1) = time_string(line_value(tri_data.line_value(tn(0))).data(0).data0.value, _
             line_value(tri_data.line_value(tn(2))).data(0).data0.value, False, False)
     ts(1) = time_string(ts(1), "2", False, False)
     ts(1) = time_string(ts(1), ts_, False, False)
     ts(0) = minus_string(ts(0), ts(1), True, False)
     If InStr(1, ts(0), "F", 0) = 0 Then
     th_cos_ = set_equation(ts(0), 0, temp_record)
     If th_cos_ > 1 Then
      Exit Function
     End If
     End If
   ElseIf angle(tri_data.angle(tn(2))).data(0).value <> "" Or tri_data.tri_function(tn(2)) Then
     If tri_data.tri_function(tn(2)) = 0 Then
     ts_ = cos_(angle(tri_data.angle(tn(2))).data(0).value, 0)
     If InStr(1, ts_, "F", 0) = 0 Then
     If tri_data.tri_function(tn(2)) = 0 Then
     th_cos_ = set_tri_function(tri_data.angle(tn(2)), "", ts_, "", "", tri_data.tri_function(tn(2)), _
            temp_record, False, tri_f, 0)
     End If
     If th_cos_ > 1 Then
       Exit Function
     End If
     triangle(no%).data(0).tri_function(tn(2)) = tri_data.tri_function(tn(2))
     End If
     Else
      ts_ = tri_function(tri_data.tri_function(tn(2))).data(0).cos_value
     End If
     ts(0) = add_string(line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
             line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, False, False)
     ts(0) = minus_string(ts(0), _
             line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, False, False)
     ts(1) = time_string(line_value(tri_data.line_value(tn(0))).data(0).data0.value, _
             line_value(tri_data.line_value(tn(1))).data(0).data0.value, False, False)
     ts(1) = time_string(ts(1), "2", False, False)
     ts(1) = time_string(ts(1), ts_, False, False)
     ts(0) = minus_string(ts(0), ts(1), True, False)
     If InStr(1, ts(0), "F", 0) = 0 Then
     th_cos_ = set_equation(ts(0), 0, temp_record)
     If th_cos_ > 1 Then
      Exit Function
     End If
     End If
   End If
Else

ts(0) = minus_string(add_string( _
          line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, _
           line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, False, cal_float), _
            line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, True, cal_float)
ts(1) = minus_string(add_string( _
          line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, _
           line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, False, cal_float), _
            line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, True, cal_float)
ts(2) = minus_string(add_string( _
          line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
           line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, False, cal_float), _
            line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, True, cal_float)
 If th_chose(52).chose = 1 Then
  temp_record.record_data.data0.theorem_no = 52
   If ts(0) = "0" Then '直角三角形
    If angle(tri_data.angle(tn(0))).data(0).value = "" Then
     th_cos_ = set_angle_value(tri_data.angle(tn(0)), "90", temp_record, 0, _
        no_reduce, False)
     If th_cos_ > 1 Then
      Exit Function
     End If
    End If
   ElseIf ts(1) = "0" Then
    If angle(tri_data.angle(tn(1))).data(0).value = "" Then
     th_cos_ = set_angle_value(tri_data.angle(tn(1)), "90", temp_record, 0, _
      no_reduce, False)
      If th_cos_ > 1 Then
       Exit Function
      End If
    End If
   ElseIf ts(2) = "0" Then
     If angle(tri_data.angle(tn(1))).data(0).value = "" Then
       th_cos_ = set_angle_value(tri_data.angle(tn(2)), "90", temp_record, 0, _
        no_reduce, False)
       If th_cos_ > 1 Then
        Exit Function
       End If
     End If
   End If 'If ts(0) = "0" Then '直角三角形
 End If 'If TH_CHOSE(52).chose = 1 Then
'**************
'余弦定理
 If th_chose(154).chose = 1 Then
  temp_record.record_data.data0.theorem_no = 154
   If ts(0) <> "0" And angle(tri_data.angle(tn(0))).data(0).value = "" Then
      ts_ = time_string(time_string(line_value(tri_data.line_value(tn(1))).data(0).data0.value, _
        line_value(tri_data.line_value(tn(2))).data(0).data0.value, False, cal_float), "2", False, cal_float)
    ts(0) = divide_string(ts(0), ts_, True, cal_float)
    ts(0) = accos_(ts(0))
     If InStr(1, ts(0), "F", 0) = 0 Then
       th_cos_ = set_angle_value(tri_data.angle(tn(0)), ts(0), temp_record, 0, _
        no_reduce, False)
      If th_cos_ > 1 Then
       Exit Function
      End If
     End If 'If ts(0) <> "F" Then
  End If
'*****************
  If ts(1) <> "0" And angle(tri_data.angle(tn(1))).data(0).value = "" Then
   ts_ = time_string(time_string(line_value(tri_data.line_value(tn(0))).data(0).data0.value, _
        line_value(tri_data.line_value(tn(2))).data(0).data0.value, False, cal_float), "2", _
          False, cal_float)
    ts(1) = divide_string(ts(1), ts_, True, cal_float)
    ts(1) = accos_(ts(1))
    If InStr(1, ts(1), "F", 0) = 0 Then
     th_cos_ = set_angle_value(tri_data.angle(tn(1)), ts(1), temp_record, 0, _
      no_reduce, False)
      If th_cos_ > 1 Then
       Exit Function
      End If
    End If
  End If
'**************************
  If ts(2) <> "0" And angle(tri_data.angle(tn(2))).data(0).value = "" Then
   ts_ = time_string(time_string(line_value(tri_data.line_value(tn(1))).data(0).data0.value, _
        line_value(tri_data.line_value(tn(0))).data(0).data0.value, False, cal_float), "2", False, cal_float)
    ts(2) = divide_string(ts(2), ts_, True, cal_float)
    ts(2) = accos_(ts(2))
    If InStr(1, ts(2), "F", 0) = 0 Then
     th_cos_ = set_angle_value(tri_data.angle(tn(2)), ts(2), temp_record, 0, _
      no_reduce, False)
      If th_cos_ > 1 Then
       Exit Function
      End If
    End If
  End If
End If
End If
'***********
'两边一夹角
ElseIf angle(tri_data.angle(tn(0))).data(0).value <> "" Or tri_data.tri_function(tn(0)) > 0 Then
     If triangle(no%).data(0).tri_function(tn(0)) = 0 Then
     temp_record.record_data.data0.condition_data.condition_no = 1
     temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
     temp_record.record_data.data0.condition_data.condition(1).no = angle(tri_data.angle(tn(0))).data(0).value_no
     ts_ = cos_(angle(tri_data.angle(tn(0))).data(0).value, 0)
     If InStr(1, ts_, "F", 0) = 0 Then
     If tri_data.tri_function(tn(0)) = 0 Then
     th_cos_ = set_tri_function(tri_data.angle(tn(0)), "", ts_, "", "", tri_data.tri_function(tn(0)), _
            temp_record, False, tri_f, 0)
     End If
     If th_cos_ > 1 Then
       Exit Function
     End If
     triangle(no%).data(0).tri_function(tn(0)) = tri_data.tri_function(tn(0))
     End If
     Else
     ts_ = tri_function(tri_data.tri_function(tn(0))).data(0).cos_value
     End If
  If tri_data.line_value(tn(1)) > 0 And tri_data.line_value(tn(2)) > 0 Then
    If tri_data.line_value(tn(0)) = 0 Then       '已知两边长
    temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(1).no = tri_data.line_value(tn(1))
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = tri_data.line_value(tn(2))
    temp_record.record_data.data0.condition_data.condition_no = 2
    If tri_data.tri_function(tn(0)) > 0 Then
    Call add_record_to_record(tri_function(tri_data.tri_function(tn(0))).data(0).record.data0.condition_data, _
     temp_record.record_data.data0.condition_data)
    Else
    End If
    Call add_conditions_to_record(angle3_value_, angle(tri_data.angle(tn(0))).data(0).value_no, _
           0, 0, temp_record.record_data.data0.condition_data)
     If tri_data.right_angle_no >= 0 Then
    Else
     If th_chose(154).chose = 1 Then
      temp_record.record_data.data0.theorem_no = 154
       ts_ = Th_cos(angle(tri_data.angle(tn(0))).data(0).value, _
           tri_function(tri_data.tri_function(tn(0))).data(0).cos_value, _
                tri_data.line_value(tn(1)), tri_data.line_value(tn(2)), "", _
                  cal_float)
       th_cos_ = set_line_value(tri_data.poi(tn(1)), tri_data.poi(tn(2)), _
           ts_, 0, 0, 0, temp_record.record_data, 0, no_reduce, False)
     End If
    End If
    End If
'    if tria
   ElseIf tri_data.line_value(tn(1)) > 0 And tri_data.line_value(tn(0)) > 0 Then
     If tri_data.line_value(tn(2)) = 0 Then       '已知两边长
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tri_data.line_value(tn(0))
     temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).no = tri_data.line_value(tn(1))
     temp_record.record_data.data0.condition_data.condition_no = 2
     If tri_data.tri_function(tn(0)) > 0 Then
     Call add_record_to_record(tri_function(tri_data.tri_function(tn(0))).data(0).record.data0.condition_data, _
      temp_record.record_data.data0.condition_data)
     Else
     End If
      it(0) = 0
      it(1) = 0
      Call set_item0(tri_data.poi(tn(0)), tri_data.poi(tn(1)), tri_data.poi(tn(0)), tri_data.poi(tn(1)), _
        "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, cond_data, 0, it(0), 0, 0, condition_data0, False)
      Call set_item0(tri_data.poi(tn(0)), tri_data.poi(tn(1)), 0, 0, _
        "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, cond_data, 0, it(1), 0, 0, condition_data0, False)
      ts(0) = time_string("-2", line_value(tri_data.line_value(tn(1))).data(0).data0.value, False, False)
      ts(0) = time_string(ts(0), ts_, True, False)
      ts(1) = minus_string(line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
              line_value(tri_data.line_value(tn(1))).data(0).data0.squar_value, True, False)
      th_cos_ = set_general_string(it(0), it(1), 0, 0, "1", ts(0), "0", "0", ts(1), 0, 0, 0, temp_record, 0, 0)
      If th_cos_ > 1 Then
         Exit Function
      End If
     End If
   ElseIf tri_data.line_value(tn(2)) > 0 And tri_data.line_value(tn(0)) > 0 Then
     If tri_data.line_value(tn(1)) = 0 Then       '已知两边长
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tri_data.line_value(tn(0))
     temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).no = tri_data.line_value(tn(2))
     temp_record.record_data.data0.condition_data.condition_no = 2
     If tri_data.tri_function(tn(0)) > 0 Then
     Call add_record_to_record(tri_function(tri_data.tri_function(tn(0))).data(0).record.data0.condition_data, _
     temp_record.record_data.data0.condition_data)
     Else
      Call add_conditions_to_record(angle3_value_, angle(tri_data.angle(tn(0))).data(0).value_no, _
              0, 0, temp_record.record_data.data0.condition_data)
     End If
    If angle(tri_data.angle(tn(0))).data(0).value = "90" Then
      ts(0) = minus_string(line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
               line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, False, False)
       ts(0) = sqr_string(ts(0), True, False)
        th_cos_ = set_line_value(tri_data.poi(tn(2)), tri_data.poi(tn(0)), ts(0), _
              0, 0, 0, temp_record.record_data, 0, 0, False)
        If th_cos_ > 1 Then
         Exit Function
        End If
       ts(0) = divide_string(line_value(tri_data.line_value(tn(2))).data(0).data0.value, _
               line_value(tri_data.line_value(tn(0))).data(0).data0.value, True, False)
       th_cos_ = set_tri_function(tri_data.angle(tn(2)), ts(0), "", "", "", 0, temp_record, True, tri_f, 0)
        If th_cos_ > 1 Then
         Exit Function
        End If
      th_cos_ = set_tri_function(tri_data.angle(tn(1)), "", ts(0), "", "", 0, temp_record, True, tri_f, 0)
        If th_cos_ > 1 Then
         Exit Function
        End If
    Else
      it(0) = 0
      it(1) = 0
      Call set_item0(tri_data.poi(tn(0)), tri_data.poi(tn(2)), tri_data.poi(tn(0)), tri_data.poi(tn(2)), _
        "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, cond_data, 0, it(0), 0, 0, condition_data0, False)
      Call set_item0(tri_data.poi(tn(0)), tri_data.poi(tn(2)), 0, 0, _
        "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, cond_data, 0, it(1), 0, 0, condition_data0, False)
      ts(0) = time_string("-2", line_value(tri_data.line_value(tn(2))).data(0).data0.value, False, False)
      ts(0) = time_string(ts(0), ts_, True, False)
      ts(1) = minus_string(line_value(tri_data.line_value(tn(0))).data(0).data0.squar_value, _
              line_value(tri_data.line_value(tn(2))).data(0).data0.squar_value, True, False)
      th_cos_ = set_general_string(it(0), it(1), 0, 0, "1", ts(0), "0", "0", ts(1), 0, 0, 0, temp_record, 0, 0)
      If th_cos_ > 1 Then
         Exit Function
      End If
     End If
     
     End If
   ElseIf tri_data.relation_no(tn(0), 0).ty > 0 Then '已知两边比
   temp_record.record_data.data0.condition_data.condition_no = 1
   temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
   temp_record.record_data.data0.condition_data.condition(1).no = angle(tri_data.angle(tn(0))).data(0).value_no
    Call add_conditions_to_record(tri_data.relation_no(tn(0), 0).ty, tri_data.relation_no(tn(0), 0).no, 0, 0, temp_record.record_data.data0.condition_data)
    If tri_data.relation_no(tn(0), 1).ty > 0 Then
        Call add_conditions_to_record(tri_data.relation_no(tn(0), 1).ty, tri_data.relation_no(tn(0), 1).no, 0, 0, temp_record.record_data.data0.condition_data)
    End If
     'If TH_CHOSE(154).chose = 1
     If tri_data.relation_no(tn(1), 0).ty = 0 Then
      temp_record.record_data.data0.theorem_no = 154
       If read_direction(tn(0), tn(1), tn(2)) < 0 Then
        ts_ = tri_data.re_value(tn(0))
      Else
        ts_ = divide_string("1", tri_data.re_value(tn(0)), False, False)
      End If
        ts_ = Th_cos(angle(tri_data.angle(tn(0))).data(0).value, tri_function(tri_data.tri_function(tn(0))).data(0).cos_value, _
                0, 0, ts_, cal_float)
        If InStr(1, ts_, "F", 0) = 0 Then
         th_cos_ = set_Drelation(tri_data.poi(tn(1)), tri_data.poi(tn(2)), _
            tri_data.poi(tn(0)), tri_data.poi(tn(1)), 0, 0, 0, 0, 0, 0, ts_, temp_record, 0, 0, 0, 0, 0, False)
             If th_cos_ > 1 Then
              Exit Function
             End If
        End If
     End If
     'If tri_data.angle_value(tn(0)) > 0 Then
     'ts_ = cos_(angle3_value(tri_data.angle_value(tn(0))).data(0).data0.value, 0)
     'Else
     'ts_ = tri_function(tri_data.tri_function(tn(0))).data(0).cos_value
     'End If
    ' If ts_ <> "F" Then
     'If tri_data.re_value(tn(0)) = ts_ Then
     '   th_cos_ = set_angle_value(tri_data.angle(tn(2)), "90", temp_record, 0, 0)
     '         If th_cos_ > 1 Then
      '        Exit Function
     '        End If
     'ElseIf divide_string("1", tri_data.re_value(tn(0)), True, False) = ts_ Then
     '   th_cos_ = set_angle_value(tri_data.angle(tn(1)), "90", temp_record, 0, 0)
     '         If th_cos_ > 1 Then
     '         Exit Function
     '        End If
     'End If
     'End If
     End If
     End If
End Function

Public Function th_20(ByVal triA%, tri As triangle_data0_type, _
                        ByVal k1%, ByVal k2%, ByVal k3%, _
                         no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim tri_data As triangle_data0_type
If angle(tri.angle(k3%)).data(0).value = "" Then
tri_data = tri
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
temp_record.record_data.data0.condition_data.condition(1).no = angle(tri_data.angle(k1%)).data(0).value_no
Call add_conditions_to_record(angle3_value_, angle(tri_data.angle(k2%)).data(0).value_no, _
      0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 20
 th_20 = set_angle_value(tri_data.angle(k3%), _
     minus_string(minus_string("180", _
       angle(tri_data.angle(k1%)).data(0).value, False, False), _
        angle(tri_data.angle(k2%)).data(0).value, True, False), _
         temp_record, 0, no_reduce, False)
End If
End Function

Public Function th_160(ByVal no%, _
           triA_ As triangle_data0_type, _
            ByVal m1%, no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim tn_(2) As Integer
Dim triA As triangle_data0_type
triA = triA_
tn_(0) = m1%
 tn_(1) = (tn_(0) + 1) Mod 3
  tn_(2) = (tn_(0) + 2) Mod 3
If triA.center = 0 Then
   If triA.midpoint_no(tn_(1)) > 0 Then
    th_160 = set_center_of_triangle(no%, triA, tn_(0), tn_(1), no_reduce)
    If th_160 > 1 Then
     Exit Function
    End If
   ElseIf triA.midpoint_no(tn_(2)) > 0 Then
    th_160 = set_center_of_triangle(no%, triA, tn_(0), tn_(2), no_reduce)
    If th_160 > 1 Then
     Exit Function
    End If
   End If
Else
   temp_record.record_data.data0.condition_data.condition_no = 2
   temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
   temp_record.record_data.data0.condition_data.condition(1).no = triA.midpoint_no(tn_(1))
   temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
   temp_record.record_data.data0.condition_data.condition(2).no = triA.midpoint_no(tn_(2))
   temp_record.record_data.data0.theorem_no = 160
   th_160 = set_mid_point_from_center_of_tria(no%, _
        triA, tn_(0), tn_(1), tn_(2), temp_record, no_reduce)
   If th_160 > 1 Then
    Exit Function
   End If
End If
End Function

Public Function set_center_of_triangle(ByVal no%, triA_ As triangle_data0_type, _
              m1%, m2%, no_reduce As Byte) As Byte
Dim tp%, m3%
Dim temp_record As total_record_type
Dim triA As triangle_data0_type
triA = triA_
If th_chose(160).chose = 1 Then
m3% = last_number_for_3number(m1%, m2%)
If triA.center = 0 Then
 tp% = is_line_line_intersect(triA.mid_point_line(m1%), _
            triA.mid_point_line(m2%), 0, 0, False)
 If tp% > 0 Then
  triangle(no%).data(0).center = tp%
    temp_record.record_data.data0.condition_data.condition_no = 2
    temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
    temp_record.record_data.data0.condition_data.condition(1).no = triA.midpoint_no(m1%)
    temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.midpoint_no(m2%)
    temp_record.record_data.data0.theorem_no = 160
    set_center_of_triangle = set_Drelation(triA.poi(m1%), _
         triA.center, triA.center, _
          Dmid_point(triA.midpoint_no(m1%)).data(0).data0.poi(1), _
           0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, no_reduce, False)
    If set_center_of_triangle > 1 Then
       Exit Function
    End If
    set_center_of_triangle = set_Drelation(triA.poi(m2%), _
         triA.center, triA.center, _
          Dmid_point(triA.midpoint_no(m2%)).data(0).data0.poi(1), _
           0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, no_reduce, False)
    If set_center_of_triangle > 1 Then
       Exit Function
    End If
    set_center_of_triangle = set_mid_point_from_center_of_tria( _
      no%, triA, m3%, m1%, m2%, temp_record, no_reduce)
    If set_center_of_triangle > 1 Then
       Exit Function
    End If
 End If
Else 'center=0
    set_center_of_triangle = set_Drelation(triA.poi(m1%), _
         triA.center, triA.center, _
          Dmid_point(triA.midpoint_no(m1%)).data(0).data0.poi(1), _
           0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, no_reduce, False)
    If set_center_of_triangle > 1 Then
       Exit Function
    End If
    set_center_of_triangle = set_mid_point_from_center_of_tria( _
      no%, triA, m1%, m2%, m3%, temp_record, no_reduce)
    If set_center_of_triangle > 1 Then
       Exit Function
    End If
End If
End If
End Function

Public Function set_mid_point_from_center_of_tria(ByVal no%, _
       triA_ As triangle_data0_type, m1%, m2%, m3%, re As total_record_type, _
        no_reduce As Byte) As Byte
Dim tp%
Dim tl(1) As Integer
Dim n(2) As Integer
Dim triA As triangle_data0_type
triA = triA_
If triA.midpoint_no(m1%) = 0 Then
tl(0) = line_number0(triA.poi(m1%), triA.poi(m2%), n(0), n(1))
tl(1) = line_number0(triA.poi(m1%), triA.center, 0, 0)
tp% = is_line_line_intersect(tl(0), _
                 tl(1), n(1), 0, False)
If tp% > 0 Then
 set_mid_point_from_center_of_tria = set_mid_point( _
     triA.poi(m2%), tp%, triA.poi(m2%), n(0), n(1), n(2), _
      tl(0), 0, re, 0, 0, 0, 0, no_reduce)
 If set_mid_point_from_center_of_tria > 1 Then
  Exit Function
 End If
End If
Else
 set_mid_point_from_center_of_tria = set_three_point_on_line( _
     triA.poi(m1%), triA.center, _
      Dmid_point(triA.midpoint_no(m1%)).data(0).data0.poi(1), _
       re, 0, no_reduce, 1)
 If set_mid_point_from_center_of_tria > 1 Then
  Exit Function
 End If
End If
End Function

Public Function set_verti_center_of_triangle(ByVal no%, _
           triA_ As triangle_data0_type, ByVal v_n1%, _
            ByVal v_n2%, no_reduce As Byte) As Byte '10.10
Dim v_n3%
Dim tl(1) As Integer
Dim temp_record As total_record_type
Dim triA As triangle_data0_type
triA = triA_
v_n3% = last_number_for_3number(v_n1%, v_n2%)
If triA.verti_no(v_n3%) = 0 Or triangle(no%).data(0).verti_center = 0 Then
 triangle(no%).data(0).verti_center = is_line_line_intersect( _
       triA.verti_line(v_n1%), triA.verti_line(v_n2%), 0, 0, False)
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(verti_, triA.verti_no(v_n1%), _
                 triA.verti_no(v_n2%), 0, temp_record.record_data.data0.condition_data)
  If triangle(no%).data(0).verti_center > 0 Then
   If triangle(no%).data(0).verti_center = triA.poi(v_n1%) Then
       triangle(no%).data(0).right_angle_no = v_n1%
    'temp_record.record_data.data0.condition_data.condition_no = 2
     'temp_record.record_data.data0.condition_data.condition(1).ty = verti_
     ' temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(v_n1%)
     'temp_record.record_data.data0.condition_data.condition(1).ty = verti_
     ' temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(v_n2%)
       set_verti_center_of_triangle = set_angle_value( _
          triangle(no%).data(0).angle(v_n1%), "90", temp_record, _
            0, 0, False)
   ElseIf triangle(no%).data(0).verti_center = triA.poi(v_n2%) Then
      triangle(no%).data(0).right_angle_no = v_n2%
    'temp_record.record_data.data0.condition_data.condition_no = 2
     ' temp_record.record_data.data0.condition_data.condition(1).ty = verti_
     '  temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(v_n1%)
     ' temp_record.record_data.data0.condition_data.condition(1).ty = verti_
     '  temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(v_n2%)
       set_verti_center_of_triangle = set_angle_value( _
          triangle(no%).data(0).angle(v_n2%), "90", temp_record, _
            0, 0, False)
   ElseIf triangle(no%).data(0).verti_center = triA.poi(v_n3%) Then
      triangle(no%).data(0).right_angle_no = v_n3%
  '     temp_record.record_data.data0.condition_data.condition_no = 0
  '      Call add_conditions_to_record(verti_, triA.verti_no(v_n1%), _
  '               triA.verti_no(v_n2%), 0, temp_record.record_data.data0.condition_data,0)
       set_verti_center_of_triangle = set_angle_value( _
          triangle(no%).data(0).angle(v_n3%), "90", temp_record, _
            0, 0, False)
   Else
       temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_conditions_to_record(verti_, triA.verti_no(v_n1%), _
              triA.verti_no(v_n2%), 0, temp_record.record_data.data0.condition_data)
    temp_record.record_data.data0.theorem_no = 161
    set_verti_center_of_triangle = set_dverti( _
       line_number0(triA.poi(v_n1%), triA.poi(v_n2%), 0, 0), _
        line_number0(triA.poi(v_n3%), triA.verti_center, 0, 0), _
         temp_record, 0, no_reduce, False)
  End If
End If
End If
End Function

Public Function th_161(ByVal no%, triA_ As triangle_data0_type, _
             ByVal v_n%, no_reduce As Byte) As Byte
Dim tn_(3) As Integer
Dim tp%
Dim temp_record As total_record_type
Dim triA As triangle_data0_type
triA = triA_
tn_(0) = v_n%
tn_(1) = (tn_(0) + 1) Mod 3
tn_(2) = (tn_(0) + 2) Mod 3
If th_chose(161).chose = 1 Then
   temp_record.record_data.data0.condition_data.condition(1).ty = verti_
   temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(tn_(0))
   temp_record.record_data.data0.theorem_no = 161
  If triA.verti_center = 0 Then
   If triA.verti_no(tn_(1)) > 0 And _
              triA.verti_no(tn_(2)) = 0 Then
      th_161 = set_verti_center_of_triangle(no%, _
         triA, tn_(0), tn_(1), no_reduce)
      If th_161 > 1 Then
        Exit Function
      End If
   ElseIf triA.verti_no(tn_(1)) = 0 And _
              triA.verti_no(tn_(2)) > 0 Then
      th_161 = set_verti_center_of_triangle(no%, _
         triA, tn_(0), tn_(2), no_reduce)
      If th_161 > 1 Then
        Exit Function
      End If
   End If
  Else
     If triA.verti_no(tn_(0)) > 0 And triA.verti_no(tn_(1)) > 0 And _
         triA.verti_no(tn_(2)) > 0 Then
      Exit Function
     Else
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = verti_
     temp_record.record_data.data0.condition_data.condition(2).ty = verti_
     temp_record.record_data.data0.theorem_no = 161
     If triA.verti_no(tn_(0)) > 0 And triA.verti_no(tn_(1)) > 0 Then
      tn_(3) = tn_(2)
      temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(tn_(1))
       temp_record.record_data.data0.condition_data.condition(2).no = triA.verti_no(tn_(0))
     ElseIf triA.verti_no(tn_(1)) > 0 And triA.verti_no(tn_(2)) > 0 Then
      tn_(3) = tn_(0)
      temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(tn_(1))
       temp_record.record_data.data0.condition_data.condition(2).no = triA.verti_no(tn_(2))
     ElseIf triA.verti_no(tn_(0)) > 0 And triA.verti_no(tn_(2)) > 0 Then
      tn_(3) = tn_(1)
      temp_record.record_data.data0.condition_data.condition(1).no = triA.verti_no(tn_(0))
       temp_record.record_data.data0.condition_data.condition(2).no = triA.verti_no(tn_(2))
     Else
      Exit Function
     End If
     End If
     th_161 = set_verti_from_verti_center_of_triangle(no%, _
        triA, tn_(3), temp_record, no_reduce)
     If th_161 > 1 Then
      Exit Function
     End If
  End If
End If
End Function

Public Function set_verti_from_verti_center_of_triangle( _
          ByVal no%, triA_ As triangle_data0_type, ByVal v_n%, _
            re As total_record_type, no_reduce As Byte) As Byte
Dim tn_(2) As Integer
Dim triA As triangle_data0_type
triA = triA_
If triA.verti_center = triA.poi(0) Or _
      triA.verti_center = triA.poi(1) Or _
         triA.verti_center = triA.poi(2) Then
           Exit Function
Else
 tn_(0) = v_n%
   tn_(1) = (tn_(0) + 1) Mod 3
    tn_(2) = (tn_(0) + 2) Mod 3
 set_verti_from_verti_center_of_triangle = set_dverti( _
  line_number0(triA.poi(tn_(1)), triA.poi(tn_(2)), 0, 0), _
        line_number0(triA.poi(tn_(0)), triA.verti_center, 0, 0), _
          re, 0, no_reduce, False)
End If
End Function

Public Function th_70(ByVal no%, triA_ As triangle_data0_type, _
                  ByVal k1%, no_reduce As Byte) As Byte
Dim tn_(2) As Integer '重心loadresstring_(522)
Dim n(2) As Integer
Dim ts As String
Dim cond_type As Byte
Dim temp_record As total_record_type
Dim triA As triangle_data0_type
triA = triA_
If th_chose(70).chose = 1 Then
tn_(0) = k1%
tn_(1) = (tn_(0) + 1) Mod 3
tn_(2) = (tn_(0) + 2) Mod 3
If triA.right_angle_no = tn_(0) Then
    If triA.midpoint_no(tn_(0)) > 0 Then
     temp_record.record_data.data0.condition_data.condition_no = 1
     temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
     temp_record.record_data.data0.condition_data.condition(1).no = angle(triA.angle(tn_(0))).data(0).value_no
     Call add_conditions_to_record(midpoint_, triA.midpoint_no(tn_(0)), 0, 0, temp_record.record_data.data0.condition_data)
     temp_record.record_data.data0.theorem_no = 70
     th_70 = set_Drelation(triA.poi(tn_(1)), triA.poi(tn_(2)), triA.poi(tn_(0)), _
        Dmid_point(triA.midpoint_no(tn_(0))).data(0).data0.poi(1), _
         0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, no_reduce, False)
     If th_70 > 1 Then
      Exit Function
     End If
    End If
ElseIf triA.midpoint_no(tn_(0)) > 0 Then
    If is_relation(triA.poi(tn_(0)), triA.poi(tn_(2)), triA.poi(tn_(0)), _
            Dmid_point(triA.midpoint_no(tn_(0))).data(0).data0.poi(1), _
             0, 0, 0, 0, 0, 0, ts, n(0), -1000, 0, 0, 0, _
                relation_data0, n(1), n(2), cond_type, record_0.data0.condition_data, 0) Then
    If ts = "2" Then
     temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
     temp_record.record_data.data0.condition_data.condition(1).no = triA.midpoint_no(tn_(0))
     If n(0) > 0 Then
       temp_record.record_data.data0.condition_data.condition_no = 2
       temp_record.record_data.data0.condition_data.condition(2).ty = cond_type
       temp_record.record_data.data0.condition_data.condition(2).no = n(0)
     Else
       temp_record.record_data.data0.condition_data.condition_no = 3
       temp_record.record_data.data0.condition_data.condition(2).ty = cond_type
       temp_record.record_data.data0.condition_data.condition(2).no = n(1)
       temp_record.record_data.data0.condition_data.condition(3).ty = cond_type
       temp_record.record_data.data0.condition_data.condition(3).no = n(2)
     End If
       temp_record.record_data.data0.theorem_no = 70
       th_70 = set_angle_value(Abs(angle_number(triA.poi(tn_(1)), _
         triA.poi(tn_(0)), triA.poi(tn_(1)), 0, 0)), "90", _
          temp_record, 0, no_reduce, False)
       If th_70 > 1 Then
        Exit Function
       End If
    End If
    End If
End If
End If
End Function
Public Function th_51(ByVal no%, triA_ As triangle_data0_type, _
                         ByVal k1%, ByVal k2%, ByVal k3%, _
                          no_reduce As Byte) As Byte '勾股定理
Dim temp_record  As total_record_type
Dim it(1) As Integer
Dim ts As String
Dim triA As triangle_data0_type
triA = triA_
If triA.right_angle_no = k1% Then
 If th_chose(51).chose = 1 Then
    temp_record.record_data.data0.theorem_no = 51
    temp_record.record_data.data0.condition_data.condition_no = 1
    temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
    temp_record.record_data.data0.condition_data.condition(1).no = angle(triA.angle(k1%)).data(0).value_no
  If triA.line_value(k2%) > 0 And triA.line_value(k3%) > 0 Then
    Call add_conditions_to_record(line_value_, triA.line_value(k2%), triA.line_value(k3%), _
         0, temp_record.record_data.data0.condition_data)
    ts = add_string(line_value(triA.line_value(k3%)).data(0).data0.squar_value, _
            line_value(triA.line_value(k3%)).data(0).data0.squar_value, False, False)
    ts = sqr_string(ts, True, False)
    If InStr(1, ts, "F", 0) = 0 Then
    th_51 = set_line_value(triA.poi(k2%), triA.poi(k3%), ts, _
            0, 0, 0, temp_record.record_data, 0, no_reduce, False)
    If th_51 > 1 Then
     Exit Function
    End If
    End If
  ElseIf triA.line_value(k1%) > 0 And triA.line_value(k2%) > 0 Then
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k1%)
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k2%)
    temp_record.record_data.data0.condition_data.condition_no = 3
    ts = minus_string(line_value(triA.line_value(k1%)).data(0).data0.squar_value, _
            line_value(triA.line_value(k2%)).data(0).data0.squar_value, False, False)
    ts = sqr_string(ts, True, False)
    If InStr(1, ts, "F", 0) = 0 Then
    th_51 = set_line_value(triA.poi(k1%), triA.poi(k2%), ts, _
            0, 0, 0, temp_record.record_data, 0, no_reduce, False)
    If th_51 > 1 Then
     Exit Function
    End If
    End If
  ElseIf triA.line_value(k1%) > 0 And triA.line_value(k3%) > 0 Then
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k1%)
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k3%)
    temp_record.record_data.data0.condition_data.condition_no = 3
    ts = minus_string(line_value(triA.line_value(k1%)).data(0).data0.squar_value, _
            line_value(triA.line_value(k3%)).data(0).data0.squar_value, False, False)
    ts = sqr_string(ts, True, False)
    If InStr(1, ts, "F", 0) = 0 Then
    th_51 = set_line_value(triA.poi(k1%), triA.poi(k3%), ts, _
            0, 0, 0, temp_record.record_data, 0, no_reduce, False)
    If th_51 > 1 Then
     Exit Function
    End If
    End If
  ElseIf triA.line_value(k1%) > 0 Then
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k1%)
    temp_record.record_data.data0.condition_data.condition_no = 2
    th_51 = set_item0(triA.poi(k1%), triA.poi(k2%), _
           triA.poi(k1%), triA.poi(k2%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
              0, it(0), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_item0(triA.poi(k1%), triA.poi(k3%), _
           triA.poi(k1%), triA.poi(k3%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
              0, it(1), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_general_string(it(0), it(1), 0, 0, "1", "1", "0", "0", _
              line_value(triA.line_value(k1%)).data(0).data0.squar_value, _
               0, 0, 0, temp_record, 0, no_reduce)
    If th_51 > 1 Then
     Exit Function
    End If
  ElseIf triA.line_value(k2%) > 0 Then
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k2%)
    temp_record.record_data.data0.condition_data.condition_no = 2
    th_51 = set_item0(triA.poi(k2%), triA.poi(k1%), _
           triA.poi(k2%), triA.poi(k1%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
               it(0), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_item0(triA.poi(k2%), triA.poi(k3%), _
           triA.poi(k2%), triA.poi(k3%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
              it(1), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_general_string(it(1), it(0), 0, 0, "1", "-1", "0", "0", _
              line_value(triA.line_value(k2%)).data(0).data0.squar_value, _
               0, 0, 0, temp_record, 0, no_reduce)
    If th_51 > 1 Then
     Exit Function
    End If
  ElseIf triA.line_value(k3%) > 0 Then
    temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = triA.line_value(k3%)
    temp_record.record_data.data0.condition_data.condition_no = 2
    th_51 = set_item0(triA.poi(k3%), triA.poi(k2%), _
           triA.poi(k3%), triA.poi(k2%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
               it(0), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_item0(triA.poi(k1%), triA.poi(k3%), _
           triA.poi(k1%), triA.poi(k3%), "*", 0, 0, _
            0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
              it(1), no_reduce, 0, condition_data0, False)
    If th_51 > 1 Then
     Exit Function
    End If
    th_51 = set_general_string(it(0), it(1), 0, 0, "1", "-1", "0", "0", _
              line_value(triA.line_value(k3%)).data(0).data0.squar_value, _
               0, 0, 0, temp_record, 0, no_reduce)
    If th_51 > 1 Then
     Exit Function
    End If
  End If
 End If
End If
End Function

Public Function th_39(ByVal no%, triA_ As triangle_data0_type, _
              ByVal k1%, ByVal k2%, ByVal k3%, _
                no_reduce As Byte) As Byte '  等腰三角形
Dim temp_record(1) As total_record_type
Dim n(2) As Integer
Dim A(1) As Integer
Dim tn(2) As Integer
Dim tl%, tp%
Dim triA As triangle_data0_type
triA = triA_
If th_chose(39).chose = 1 Then
 If triA.relation_no(k1%, 0).ty > 0 And triA.re_value(k1%) = "1" Then
    temp_record(0).record_data.data0.condition_data.condition_no = _
       temp_record(0).record_data.data0.condition_data.condition_no + 1
     temp_record(0).record_data.data0.condition_data.condition(temp_record(0). _
           record_data.data0.condition_data.condition_no) = triA.relation_no(k1%, 0)
      If triA.relation_no(k1%, 1).ty > 0 Then
       temp_record(0).record_data.data0.condition_data.condition_no = _
         temp_record(0).record_data.data0.condition_data.condition_no + 1
           temp_record(0).record_data.data0.condition_data.condition(temp_record(0). _
            record_data.data0.condition_data.condition_no) = triA.relation_no(k1%, 1)
      End If
      temp_record(0).record_data.data0.theorem_no = 39
       tl = line_number0(triA.poi(k2%), triA.poi(k3%), tn(0), tn(2))
If triA.mid_point_line(k1%) > 0 Then
   temp_record(1) = temp_record(0)
    Call add_conditions_to_record(midpoint_, triA.midpoint_no(k1%), 0, 0, _
                                   temp_record(1).record_data.data0.condition_data)
   If triA.verti_line(k1%) = 0 Then
    th_39 = set_dverti(tl%, triA.mid_point_line(k1%), _
          temp_record(1), 0, no_reduce, False)
    If th_39 > 1 Then
     Exit Function
    End If
   End If
   If triA.eangle_line(k1%) = 0 Then
   A(0) = angle_number(triA.poi(k2%), triA.poi(k1%), _
          Dmid_point(triA.midpoint_no(k1%)).data(0).data0.poi(1), 0, 0)
   A(1) = angle_number(triA.poi(k3%), triA.poi(k1%), _
          Dmid_point(triA.midpoint_no(k1%)).data(0).data0.poi(1), 0, 0)
     If A(0) <> 0 And A(1) <> 0 Then
      th_39 = set_three_angle_value(Abs(A(0)), Abs(A(1)), 0, _
       "1", "-1", "0", "0", 0, temp_record(1), 0, 0, 0, no_reduce, _
         0, 0, False)
      If th_39 > 1 Then
       Exit Function
      End If
     End If
   End If
ElseIf triA.verti_line(k1%) > 0 Then
   If Dverti(triA.verti_no(k1%)).data(0).inter_poi > 0 Then
      Call is_point_in_line3( _
         Dverti(triA.verti_no(k1%)).data(0).inter_poi, _
           m_lin(tl%).data(0).data0, tn(1))
      temp_record(1) = temp_record(0)
       Call add_conditions_to_record(verti_, triA.verti_no(k1%), 0, 0, _
                                     temp_record(1).record_data.data0.condition_data)
     If triA.mid_point_line(k1%) = 0 Then
       th_39 = set_mid_point(triA.poi(k2%), _
          Dverti(triA.verti_no(k1%)).data(0).inter_poi, _
           triA.poi(k3%), tn(0), tn(1), tn(2), tl%, _
             0, temp_record(1), 0, 0, 0, 0, no_reduce)
      If th_39 > 1 Then
       Exit Function
      End If
     End If
   End If
   If triA.eangle_line(k1%) = 0 Then
     A(0) = angle_number(triA.poi(k2%), triA.poi(k1%), _
              m_lin(triA.eangle_line(k1%)).data(0).data0.poi(0), 0, 0)
     A(1) = angle_number(triA.poi(k3%), triA.poi(k1%), _
              m_lin(triA.eangle_line(k1%)).data(0).data0.poi(0), 0, 0)
    If A(0) <> 0 And A(1) <> 0 Then
     th_39 = set_three_angle_value(Abs(A(0)), Abs(A(1)), 0, _
      "1", "-1", "0", "0", 0, temp_record(1), 0, 0, 0, no_reduce, _
        0, 0, False)
     If th_39 > 1 Then
      Exit Function
     End If
    End If
     A(0) = angle_number(triA.poi(k2%), triA.poi(k1%), _
              m_lin(triA.eangle_line(k1%)).data(0).data0.poi(1), 0, 0)
     A(1) = angle_number(triA.poi(k3%), triA.poi(k1%), _
              m_lin(triA.eangle_line(k1%)).data(0).data0.poi(1), 0, 0)
    If A(0) <> 0 And A(1) <> 0 Then
     th_39 = set_three_angle_value(Abs(A(0)), Abs(A(1)), 0, _
      "1", "-1", "0", "0", 0, temp_record(1), 0, 0, 0, no_reduce, _
        0, 0, False)
     If th_39 > 1 Then
      Exit Function
     End If
    End If
   End If
ElseIf triA.eangle_line(k1%) > 0 Then
   temp_record(1) = temp_record(0)
   temp_record(1).record_data.data0.condition_data.condition_no = temp_record(1).record_data.data0.condition_data.condition_no + 1
   temp_record(1).record_data.data0.condition_data.condition( _
           temp_record(1).record_data.data0.condition_data.condition_no) = _
        triA.eangle_no(k1%, 0)
   'temp_record(1).record_data.data0.condition_data.condition(temp_record(1).record_data.data0.condition_data.condition_no).ty = _
        angle3_value_
   If triA.eangle_no(k1%, 1).no > 0 Then
    temp_record(1).record_data.data0.condition_data.condition_no = _
       temp_record(1).record_data.data0.condition_data.condition_no + 1
   temp_record(1).record_data.data0.condition_data.condition( _
           temp_record(1).record_data.data0.condition_data.condition_no) = _
        triA.eangle_no(k1%, 1)
   End If
   If triA.verti_line(k1%) = 0 Then
    th_39 = set_dverti(tl%, triA.eangle_line(k1%), temp_record(1), _
           0, no_reduce, False)
    If th_39 > 1 Then
     Exit Function
    End If
   End If
   tp% = is_line_line_intersect(tl%, triA.eangle_line(k1%), _
          tn(1), 0, False)
   If tp% > 0 Then
      If triA.mid_point_line(k1%) = 0 Then
       th_39 = set_mid_point(triA.poi(k2%), tp%, _
           triA.poi(k3%), tn(0), tn(1), tn(2), _
             tl%, 0, temp_record(1), 0, 0, 0, 0, no_reduce)
       If th_39 > 1 Then
        Exit Function
       End If
      End If
   End If
End If
End If
End If
End Function

Public Function th_120(ByVal verti_no%, no_reduce As Byte) As Byte '垂径定理
 th_120 = th_120_(Dverti(verti_no%).data(0).line_no(0), _
       Dverti(verti_no%).data(0).line_no(1), verti_no%, no_reduce)
 If th_120 > 1 Then
  Exit Function
 End If
 th_120 = th_120_(Dverti(verti_no%).data(0).line_no(1), _
       Dverti(verti_no%).data(0).line_no(0), verti_no%, no_reduce)
End Function

Public Function th_120_(l1%, l2%, v_n%, no_reduce As Byte) As Byte
Dim i%, j%, k%, l%
Dim n As Integer
Dim n_(1) As Integer
Dim m(1) As Integer
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = verti_
temp_record.record_data.data0.condition_data.condition(1).no = v_n%
temp_record.record_data.data0.theorem_no = 120
If th_chose(120).chose = 1 Then
 For i% = 1 To C_display_picture.m_circle.Count
  If m_Circ(i%).data(0).data0.center > 0 And m_Circ(i%).data(0).data0.in_point(0) > 2 Then
   n = 0
   For j% = 1 To m_lin(l1%).data(0).data0.in_point(0)
    For k% = 1 To m_Circ(i%).data(0).data0.in_point(0)
     If m_Circ(i%).data(0).data0.in_point(k%) = m_lin(l1%).data(0).data0.in_point(j%) Then
      n_(n) = k%
       n = n + 1
        If n = 2 Then
         GoTo th_120_mark0
        End If
     End If
    Next k%
  Next j%
GoTo th_120_last
th_120_mark0:
   If is_point_in_line3(m_Circ(i%).data(0).data0.center, m_lin(l2%).data(0).data0, 0) Then
      If Dverti(v_n%).data(0).inter_poi > 0 Then
       th_120_ = set_mid_point(m_Circ(i%).data(0).data0.in_point(n_(0)), _
        Dverti(v_n%).data(0).inter_poi, m_Circ(i%).data(0).data0.in_point(n_(1)), _
         0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, no_reduce)
        If th_120_ > 1 Then
         Exit Function
        End If
      End If
   For j% = 1 To m_lin(l2%).data(0).data0.in_point(0)
    For k% = 1 To m_Circ(i%).data(0).data0.in_point(0)
     If m_Circ(i%).data(0).data0.in_point(k%) = m_lin(l2%).data(0).data0.in_point(j%) Then
      th_120_ = set_equal_arc(arc_no(m_Circ(i%).data(0).data0.in_point(n_(0)), _
            i%, m_Circ(i%).data(0).data0.in_point(k%)), _
             arc_no(m_Circ(i%).data(0).data0.in_point(k%), i%, _
              m_Circ(i%).data(0).data0.in_point(n_(1))), temp_record, 0, _
               no_reduce)
      If th_120_ > 1 Then
       Exit Function
      End If
      th_120_ = set_equal_arc(arc_no(m_Circ(i%).data(0).data0.in_point(n_(0)), _
            i%, m_Circ(i%).data(0).data0.in_point(k%)), _
             arc_no(m_Circ(i%).data(0).data0.in_point(k%), i%, _
              m_Circ(i%).data(0).data0.in_point(n_(1))), temp_record, 0, _
               no_reduce)
      If th_120_ > 1 Then
       Exit Function
      End If
      th_120_ = set_equal_dline(m_Circ(i%).data(0).data0.in_point(n_(0)), _
             m_Circ(i%).data(0).data0.in_point(k%), m_Circ(i%).data(0).data0.in_point(k%), _
              m_Circ(i%).data(0).data0.in_point(n_(1)), 0, 0, 0, 0, 0, 0, 0, _
               temp_record, 0, 0, 0, 0, no_reduce, False)
      If th_120_ > 1 Then
       Exit Function
      End If
     End If
    Next k%
  Next j%
  End If
  End If
th_120_last:
 Next i%
End If
End Function

Public Function th_menei_(ByVal re1%, ty2 As Byte, ByVal re2%, ByVal k%, ByVal l%) As Byte
'美奈劳是斯定理
Dim tp1(2) As Integer
Dim tp2(2) As Integer
Dim tn1(2) As Integer
Dim tn2(2) As Integer
Dim tl(1) As Integer
Dim tp%
Dim m1(2) As Integer
Dim m2(2) As Integer
Dim ts1(2) As String
Dim ts2(2) As String
Dim temp_record As total_record_type
If k% > 2 Then
 k% = k% - 1
End If
If l% > 2 Then
 l% = l% - 1
End If
tp1(0) = Drelation(re1%).data(0).data0.poi(0)
tp1(1) = Drelation(re1%).data(0).data0.poi(1)
tp1(2) = Drelation(re1%).data(0).data0.poi(3)
tn1(0) = Drelation(re1%).data(0).data0.n(0)
tn1(1) = Drelation(re1%).data(0).data0.n(1)
tn1(2) = Drelation(re1%).data(0).data0.n(3)
ts1(0) = divide_string(Drelation(re1%).data(0).data0.value, _
          add_string("1", Drelation(re1%).data(0).data0.value, False, False), True, False) 'tp1(m1(1))tp1(m1(2))/tp1(m1(2))tp1(m1(0))
ts1(1) = divide_string("1", Drelation(re1%).data(0).data0.value, True, False)
ts1(2) = add_string("1", Drelation(re1%).data(0).data0.value, True, False)            'tp1(m1(1))tp1(m1(2))/tp1(m1(2))tp1(m1(0))
temp_record.record_data.data0.condition_data.condition(1).ty = relation_
temp_record.record_data.data0.condition_data.condition(1).no = re1%
If ty2 = relation_ Then
tp2(0) = Drelation(re2%).data(0).data0.poi(0)
tp2(1) = Drelation(re2%).data(0).data0.poi(1)
tp2(2) = Drelation(re2%).data(0).data0.poi(3)
tn2(0) = Drelation(re2%).data(0).data0.n(0)
tn2(1) = Drelation(re2%).data(0).data0.n(1)
tn2(2) = Drelation(re2%).data(0).data0.n(3)
ts2(0) = divide_string(Drelation(re2%).data(0).data0.value, _
          add_string("1", Drelation(re2%).data(0).data0.value, False, False), True, False) 'tp1(m1(1))tp1(m1(2))/tp1(m1(2))tp1(m1(0))
ts2(1) = divide_string("1", Drelation(re2%).data(0).data0.value, True, False)
ts2(2) = add_string("1", Drelation(re2%).data(0).data0.value, True, False)            'tp1(m1(1))tp1(m1(2))/tp1(m1(2))tp1(m1(0))
temp_record.record_data.data0.condition_data.condition(2).ty = relation_
temp_record.record_data.data0.condition_data.condition(2).no = re2%
Else
tp2(0) = Dmid_point(re2%).data(0).data0.poi(0)
tp2(1) = Dmid_point(re2%).data(0).data0.poi(1)
tp2(2) = Dmid_point(re2%).data(0).data0.poi(2)
tn2(0) = Dmid_point(re2%).data(0).data0.n(0)
tn2(1) = Dmid_point(re2%).data(0).data0.n(1)
tn2(2) = Dmid_point(re2%).data(0).data0.n(2)
ts2(0) = "1/2"
ts2(1) = "1"
ts2(2) = "2"
temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
temp_record.record_data.data0.condition_data.condition(2).no = re2%
End If
m1(0) = k%
m1(1) = (k% + 1) Mod 3
m1(2) = (k% + 2) Mod 3
m2(0) = l%
m2(1) = (l% + 1) Mod 3
m2(2) = (l% + 2) Mod 3
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.theorem_no = -5
tl(0) = line_number0(tp1(m1(1)), tp2(m2(1)), tn1(0), tn1(1))
tl(1) = line_number0(tp1(m1(2)), tp2(m2(2)), tn2(0), tn2(1))
tp% = is_line_line_intersect(tl(0), tl(1), tn1(2), tn2(2), False)
If tp% > 0 Then
  th_menei_ = max_for_byte(th_menei_, set_Drelation(tp%, tp1(m1(1)), tp%, tp2(m2(1)), tn1(2), tn1(0), _
        tn1(2), tn1(1), tl(0), tl(0), divide_string(ts2(m2(2)), ts1(m1(2)), True, False), _
         temp_record, 0, 0, 0, 0, 0, False))
  If th_menei_ > 1 Then
   Exit Function
  End If
  th_menei_ = max_for_byte(th_menei_, set_Drelation(tp%, tp1(m1(2)), tp%, tp2(m2(2)), tn2(2), tn2(0), _
        tn2(2), tn2(1), tl(1), tl(1), divide_string(ts1(m1(1)), ts2(m2(1)), True, False), _
         temp_record, 0, 0, 0, 0, 0, False))
  If th_menei_ > 1 Then
   Exit Function
  End If
End If
tl(0) = line_number0(tp1(m1(1)), tp2(m2(2)), tn1(0), tn1(1))
tl(1) = line_number0(tp1(m1(2)), tp2(m2(1)), tn2(0), tn2(1))
tp% = is_line_line_intersect(tl(0), tl(1), tn1(2), tn2(2), False)
If tp% > 0 Then
       th_menei_ = max_for_byte(th_menei_, set_Drelation(tp%, tp1(m1(1)), tp%, tp2(m2(2)), tn1(2), tn1(0), _
        tn1(2), tn1(1), tl(0), tl(0), divide_string(divide_string("1", _
         ts2(m2(1)), False, False), ts1(m1(2)), True, False), temp_record, 0, 0, 0, 0, 0, False))
  If th_menei_ > 1 Then
   Exit Function
  End If
  th_menei_ = max_for_byte(th_menei_, set_Drelation(tp%, tp1(m1(2)), tp%, tp2(m2(1)), tn2(2), tn2(0), _
        tn2(2), tn2(1), tl(1), tl(1), time_string(ts2(m2(2)), ts1(m1(1)), True, False), _
         temp_record, 0, 0, 0, 0, 0, False))
  If th_menei_ > 1 Then
   Exit Function
  End If
End If
End Function

Public Function th_menei() As Byte
Dim i%, j%
For i% = 1 To last_conditions.last_cond(1).relation_no - 1
 If Drelation(i%).data(0).data0.line_no(1) = Drelation(i%).data(0).data0.line_no(0) Then
  If Drelation(i%).data(0).data0.poi(1) = Drelation(i%).data(0).data0.poi(2) Then
  For j% = i% + 1 To last_conditions.last_cond(1).relation_no
   If Drelation(j%).data(0).data0.line_no(1) = Drelation(j%).data(0).data0.line_no(0) And _
       Drelation(i%).data(0).data0.line_no(0) <> Drelation(j%).data(0).data0.line_no(0) Then
    If Drelation(j%).data(0).data0.poi(1) = Drelation(j%).data(0).data0.poi(2) Then
     If Drelation(i%).data(0).data0.poi(0) = Drelation(j%).data(0).data0.poi(0) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 0, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(0) = Drelation(j%).data(0).data0.poi(1) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 0, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(0) = Drelation(j%).data(0).data0.poi(3) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 0, 3))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Drelation(j%).data(0).data0.poi(0) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 1, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Drelation(j%).data(0).data0.poi(1) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 1, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Drelation(j%).data(0).data0.poi(3) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 1, 3))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Drelation(j%).data(0).data0.poi(0) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 3, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Drelation(j%).data(0).data0.poi(1) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 3, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Drelation(j%).data(0).data0.poi(3) Then
     th_menei = max_for_byte(th_menei, th_menei_(i%, relation_, j%, 3, 3))
      If th_menei > 1 Then
       Exit Function
      End If
     End If
    End If
   End If
  Next j%
  For j% = 1 To last_conditions.last_cond(1).mid_point_no
    If Dmid_point(j%).data(0).data0.line_no <> Drelation(i%).data(0).data0.line_no(0) Then
     If Drelation(i%).data(0).data0.poi(0) = Dmid_point(j%).data(0).data0.poi(0) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 0, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(0) = Dmid_point(j%).data(0).data0.poi(1) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 0, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(0) = Dmid_point(j%).data(0).data0.poi(2) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 0, 2))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Dmid_point(j%).data(0).data0.poi(0) Then
       th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 1, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Dmid_point(j%).data(0).data0.poi(1) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 1, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(1) = Dmid_point(j%).data(0).data0.poi(2) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 1, 2))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Dmid_point(j%).data(0).data0.poi(0) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 3, 0))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Dmid_point(j%).data(0).data0.poi(1) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 3, 1))
      If th_menei > 1 Then
       Exit Function
      End If
     ElseIf Drelation(i%).data(0).data0.poi(3) = Dmid_point(j%).data(0).data0.poi(2) Then
      th_menei = max_for_byte(th_menei, th_menei_(i%, midpoint_, j%, 3, 2))
      If th_menei > 1 Then
       Exit Function
      End If
     End If
   End If
  Next j%
  End If
 End If
Next i%
End Function
Public Sub add_condition_to_record(ByVal con_ty As Integer, ByVal con_no%, re As condition_data_type, is_first As Byte)
Dim temp_record As record_data_type
Dim j%
Dim reduce_level As Integer
If con_ty = 0 Or con_no% = 0 Then
   Exit Sub
End If
 temp_record = get_record_data(con_ty, con_no%)
   If temp_record.data0.condition_data.condition_no = 1 And _
       temp_record.data0.condition_data.condition(1).ty = wenti_cond_ And _
        temp_record.data0.condition_data.condition(1).no < 0 Then
         con_ty = temp_record.data0.condition_data.condition(1).ty
          con_no% = -temp_record.data0.condition_data.condition(1).no
    '输入语句产生的数据，一包含在语句中(con_no%<0)，另一不在语句中，而是有输入推出的(con_no%>0)
   End If
    For j% = 1 To re.condition_no
     If con_ty = re.condition(j%).ty And _
                     con_no% = re.condition(j%).no Then
        Exit Sub
     End If
    Next j%
   re.condition_no = re.condition_no + 1
   re.condition(re.condition_no).ty = con_ty
   re.condition(re.condition_no).no = con_no%

End Sub
Public Sub combine_condition_to_record(ByVal con_ty As Integer, ByVal con_no%, re As condition_data_type, is_first As Byte)
Dim j%, k%
Dim temp_record As record_data_type
Dim reduce_level As Integer
If con_ty = 0 Or con_no% = 0 Then
   Exit Sub
End If
If con_ty = add_condition_ And is_first = 0 Then
 temp_record.data0.condition_data = re
 If new_point(con_no%).data(0).cond.no > 0 Then
  If is_condition_in_record(new_point(con_no%).data(0).cond.ty, _
       new_point(con_no%).data(0).cond.no, temp_record, 255) = False Then
    Call add_condition_to_record(new_point(con_no%).data(0).cond.ty, _
       new_point(con_no%).data(0).cond.no, re, 1)
  End If
 Else
   Call add_condition_to_record(con_ty, con_no%, re, 1)
 End If
Else
    For j% = 1 To re.condition_no
     If con_ty = re.condition(j%).ty And _
                     con_no% = re.condition(j%).no Then
        GoTo add_condition_to_record_mark0
     End If
    Next j%
 temp_record = get_record_data(con_ty, con_no%)
   If temp_record.data0.condition_data.condition_no = 1 And _
    temp_record.data0.condition_data.condition(1).ty = wenti_cond_ Then
     con_ty = wenti_cond_
      con_no% = temp_record.data0.condition_data.condition(1).no
       reduce_level = -1
   Else
       reduce_level = temp_record.data0.condition_data.level
   End If
If con_no% > 0 Then
   For j% = 1 To re.condition_no
    If con_ty = general_string_ And con_ty = re.condition(j%).ty Then
       If general_string(con_no%).record_.conclusion_no = 0 And _
           general_string(re.condition(j%).no).record_.conclusion_no = 0 Then
         If con_no% > re.condition(j%).no Then
            Call insert_record(con_ty, con_no%, j%, re)
             GoTo add_condition_to_record_mark0
         End If
       ElseIf general_string(con_no%).record_.conclusion_no = 0 Then
            Call insert_record(con_ty, con_no%, j%, re)
             GoTo add_condition_to_record_mark0
       End If
    ElseIf con_ty < re.condition(j%).ty Or (con_ty = re.condition(j%).ty And _
         con_no% > re.condition(j%).no) Then
          Call insert_record(con_ty, con_no%, j%, re) '插入记录
         GoTo add_condition_to_record_mark0
    End If
   Next j%
     re.condition_no = re.condition_no + 1 '插入最后
      re.condition(re.condition_no).ty = con_ty
       re.condition(re.condition_no).no = con_no%
        re.level = max(re.level, reduce_level + 1)
End If
add_condition_to_record_mark0:
End If
End Sub
Public Sub insert_record(ByVal con_ty As Byte, ByVal con_no%, ByVal insert_no, c_data As condition_data_type)
Dim k%
     c_data.condition_no = c_data.condition_no + 1
     For k% = c_data.condition_no To insert_no + 1 Step -1
         c_data.condition(k%).ty = _
               c_data.condition(k% - 1).ty
         c_data.condition(k%).no = _
                 c_data.condition(k% - 1).no
       Next k%
       c_data.condition(insert_no).ty = con_ty
        c_data.condition(insert_no).no = con_no%
End Sub

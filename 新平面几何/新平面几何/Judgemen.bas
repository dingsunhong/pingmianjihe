Attribute VB_Name = "judgement"
Option Explicit
Public Function is_same_point_by_coord(coord1 As POINTAPI, coord2 As POINTAPI) As Boolean
   If Abs(coord1.X - coord2.X) + Abs(coord1.Y - coord2.Y) < 9 Then
      is_same_point_by_coord = True
   Else
      is_same_point_by_coord = False
   End If
End Function
Public Function is_point_in_points(p1%, points() As Integer) As Integer
Dim i%
 For i% = 1 To points(0)
  If points(i%) = p1% Then
   is_point_in_points = i%
    Exit Function
  End If
 Next i%
  is_point_in_points = 0
End Function
Public Function is_con_general_string(no%) As Byte
Dim gs As general_string_data_type
Dim ts As String
If general_string(no%).data(0).value <> "" Then
If general_string(no%).record_.conclusion_no > 0 Then
If conclusion_data(general_string(no%).record_.conclusion_no - 1).no(0) = 0 Then
'是结论
gs = general_string(no%).data(0)
gs.value_ = gs.value '
  If gs.record.data0.condition_data.condition_no > 9 Then
   GoTo is_con_general_string_mark1
  End If
While gs.record.data0.condition_data.condition(gs.record.data0.condition_data.condition_no).ty = general_string_ And _
       gs.record.data0.condition_data.condition(gs.record.data0.condition_data.condition_no).no > 0
 general_string(gs.record.data0.condition_data.condition(gs.record.data0.condition_data.condition_no).no).data(0).value_ = _
    time_string(gs.value_, gs.trans_para, True, False)
 gs = general_string(gs.record.data0.condition_data.condition(gs.record.data0.condition_data.condition_no).no).data(0)
  If gs.record.data0.condition_data.condition_no > 9 Or _
        gs.record.data0.condition_data.condition_no = 0 Then
   GoTo is_con_general_string_mark1
  End If
Wend
is_con_general_string_mark1:
ts = time_string(gs.value_, gs.trans_para, True, False)
If InStr(1, ts, "F", 0) = 0 Then
If con_general_string(general_string(no%).record_.conclusion_no).data(0).value = "" Then
    con_general_string(general_string(no%).record_.conclusion_no).data(0).value = ts
     conclusion_data(general_string(no%).record_.conclusion_no - 1).no(0) = no%
      is_con_general_string = is_complete_prove
ElseIf con_general_string(general_string(no%).record_.conclusion_no).data(0).value = ts Then
     conclusion_data(general_string(no%).record_.conclusion_no - 1).no(0) = no%
      is_con_general_string = is_complete_prove
Else
     conclusion_data(general_string(no%).record_.conclusion_no - 1).no(0) = no%
      is_con_general_string = is_complete_prove
       error_of_wenti = 2
End If
End If
End If
End If
End If
End Function
Public Function is_dparal0(ByVal l1%, ByVal l2%) As Boolean
Dim i%
For i% = 1 To m_lin(l1%).data(0).in_paral(0).lin
 If m_lin(l1%).data(0).in_paral(i%).line_no = l2% Then
  is_dparal0 = True
   Exit Function
 End If
Next i%
End Function
Public Function is_dparal(ByVal l1%, ByVal l2%, n%, n1%, n2%, n3%, _
         outl1%, outl2%) As Boolean
Dim i%, tn%
Dim pl As two_line_type
If n1% = -1000 Then
 If l1% = 0 Or l2% = 0 Or last_conditions.last_cond(1).paral_no = 0 Then
    is_dparal = False
     Exit Function
 End If
Else
 If l1% = 0 Or l2% = 0 Then
    is_dparal = True
     Exit Function
 End If
End If
n% = 0
For i% = 1 To last_conditions.last_cond(1).same_three_lines_no
  If same_three_lines(i%).data(0).line_no(0) = l1% Or _
           same_three_lines(i%).data(0).line_no(1) = l1% Then
   l1% = same_three_lines(i%).data(0).line_no(2)
  End If
  If same_three_lines(i%).data(0).line_no(0) = l2% Or _
           same_three_lines(i%).data(0).line_no(1) = l2% Then
   l2% = same_three_lines(i%).data(0).line_no(2)
  End If
Next i%
If l1% = 0 Or l2% = 0 Then
 is_dparal = False
  Exit Function
ElseIf l1% = l2% Then
 is_dparal = True
  Exit Function
ElseIf last_conditions.last_cond(1).paral_no = 0 And n1% = -1000 Then
 is_dparal = False
   Exit Function
Else
 If l1 > l2% Then
  outl1% = l2%
   outl2% = l1%
 Else
 outl1% = l1%
  outl2% = l2%
 End If
End If
pl.line_no(0) = outl1%
 pl.line_no(1) = outl2%
If search_for_paral(pl, 0, n%, 1) Then '5.7
 If n1% = -5000 Then
   n1% = n%
 Else
   n% = Dparal(n% + 1).data(0).data0.record.data1.index.i(0)
   If set_or_prove = 2 Then '
    If Dparal(i%).data(0).data0.record.data1.is_proved = 1 Then
     is_dparal = True
    End If
   Else
     is_dparal = True
   End If
    Exit Function
 End If
Else
 If n1% = -1000 Then
  n% = 0
   Exit Function
 End If
  n1% = n%
End If
   Call search_for_paral(pl, 1, n2%, 1) '5.7
'   Call search_for_paral(pl, 2, n3%, 1) '5.7
    is_dparal = False
End Function
Public Function is_four_point_in_equal_side_tixing(ByVal p1%, _
     ByVal p2%, ByVal p3%, ByVal p4%) As Boolean
Dim tp(3) As Integer
Dim i%, j%, k%
For i% = 1 To last_conditions.last_cond(1).tixing_no
 If Dpolygon4(Dtixing(i%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
 For j% = 0 To 3
  For k% = 0 To 3
   If tp(j%) = Dtixing(i%).data(0).poi(k%) Then
    GoTo is_four_point_in_equal_side_tixing_mark0
   End If
  Next k%
GoTo is_four_point_in_equal_side_tixing_mark1
is_four_point_in_equal_side_tixing_mark0:
Next j%
is_four_point_in_equal_side_tixing = True
Exit Function
is_four_point_in_equal_side_tixing_mark1:
End If
Next i%
is_four_point_in_equal_side_tixing = False
End Function
Function is_four_point_on_circle(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
       no%, no1%, no2%, no3%, no4%, no5%, no6%, op1%, op2%, op3%, op4%, _
          p4_c As four_point_on_circle_data_type, set_or_judg As Boolean) As Boolean
'判断是否共圆，并排序
 '不共圆输出-1
Dim i%, j%, p%, c%
Dim tp(3) As Integer
Dim ang(3) As Integer
Dim l_v As line_value_data0_type
Dim temp_record As total_record_type
 no% = 0
If p1% = p2% Or p1% = p3% Or p1% = p4% Or p2% = p3% Or p2% = p4% Or p3% = p4% Then
 If no1% = -1000 Then
  is_four_point_on_circle = False
   Exit Function
 Else
    is_four_point_on_circle = True
   Exit Function
 End If
End If
 tp(0) = p1%
 tp(1) = p2%
 tp(2) = p3%
 tp(3) = p4%
If arrange_four_point_on_circle(tp(0), tp(1), tp(2), tp(3)) Then
   op1% = tp(0)
   op2% = tp(1)
   op3% = tp(2)
   op4% = tp(3)
Else
 If no1% = -1000 Then
  is_four_point_on_circle = False
 ElseIf no1% = 0 Then
  is_four_point_on_circle = True
   no% = 0
 End If
  Exit Function
End If
p4_c.poi(0) = op1%
 p4_c.poi(1) = op2%
  p4_c.poi(2) = op3%
   p4_c.poi(3) = op4%
If is_four_point_in_epolygon(tp(0), tp(1), tp(2), tp(3), 0, 0, "") Then
  is_four_point_on_circle = True
   no% = 0
    Exit Function
 End If
 If is_four_point_in_equal_side_tixing(tp(0), tp(1), tp(2), tp(3)) Then
  is_four_point_on_circle = True
   no% = 0
    Exit Function
 End If
 If last_conditions.last_cond(1).four_point_on_circle_no = 0 And no1% = -1000 Then
   p4_c.angle(0) = Abs(angle_number(p4_c.poi(3), p4_c.poi(0), p4_c.poi(1), "", 0))
  p4_c.angle(1) = Abs(angle_number(p4_c.poi(0), p4_c.poi(1), p4_c.poi(2), "", 0))
  p4_c.angle(2) = Abs(angle_number(p4_c.poi(1), p4_c.poi(2), p4_c.poi(3), "", 0))
  p4_c.angle(3) = Abs(angle_number(p4_c.poi(2), p4_c.poi(3), p4_c.poi(0), "", 0))
  Call is_line_value(p4_c.poi(0), p4_c.poi(1), 0, 0, 0, "", _
        p4_c.lin_value_no(0), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(1), p4_c.poi(2), 0, 0, 0, "", _
        p4_c.lin_value_no(1), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(2), p4_c.poi(3), 0, 0, 0, "", _
        p4_c.lin_value_no(2), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(3), p4_c.poi(0), 0, 0, 0, "", _
        p4_c.lin_value_no(3), -1000, 0, 0, 0, l_v)
   
     no% = 0
      is_four_point_on_circle = False
       GoTo is_four_point_on_circle_mark1
 End If
is_four_point_on_circle_next2:
If search_for_four_point_on_circle(p4_c, 1, 0, no%, 0) Then
   If set_or_prove = 2 Then '
    If four_point_on_circle(no%).data(0).record.data1.is_proved = 1 Then
     is_four_point_on_circle = True
    End If
   Else
         is_four_point_on_circle = True
   End If
            Exit Function
 End If
' Next i%
If set_or_judg = True Then
If no1% = -1000 Then
is_four_point_on_circle_mark1:
    c% = m_circle_number(1, 0, pointapi0, p1%, p2%, p3%, 0, 0, 0, 1, 0, 0, 0, False)
     If c% > 0 Then
        If is_point_in_circle(c%, 0, p4%, 0, 0) Then
           Call set_four_point_on_circle(p1%, p2%, p3%, p4%, c%, temp_record, no%, 0)
            is_four_point_on_circle = True
             Exit Function
        End If
     End If
 no% = 0
  Exit Function
End If
no1% = no%
  Call search_for_four_point_on_circle(p4_c, 1, 1, no2%, 1)
  Call search_for_four_point_on_circle(p4_c, 1, 2, no3%, 1)
  Call search_for_four_point_on_circle(p4_c, 1, 3, no4%, 1)
  Call search_for_four_point_on_circle(p4_c, 1, 4, no5%, 1)
  Call search_for_four_point_on_circle(p4_c, 1, 5, no6%, 1)
  p4_c.angle(0) = Abs(angle_number(p4_c.poi(3), p4_c.poi(0), p4_c.poi(1), "", 0))
  p4_c.angle(1) = Abs(angle_number(p4_c.poi(0), p4_c.poi(1), p4_c.poi(2), "", 0))
  p4_c.angle(2) = Abs(angle_number(p4_c.poi(1), p4_c.poi(2), p4_c.poi(3), "", 0))
  p4_c.angle(3) = Abs(angle_number(p4_c.poi(2), p4_c.poi(3), p4_c.poi(0), "", 0))
  Call is_line_value(p4_c.poi(0), p4_c.poi(1), 0, 0, 0, "", _
        p4_c.lin_value_no(0), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(1), p4_c.poi(2), 0, 0, 0, "", _
        p4_c.lin_value_no(1), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(2), p4_c.poi(3), 0, 0, 0, "", _
        p4_c.lin_value_no(2), -1000, 0, 0, 0, l_v)
  Call is_line_value(p4_c.poi(3), p4_c.poi(0), 0, 0, 0, "", _
        p4_c.lin_value_no(3), -1000, 0, 0, 0, l_v)
End If
is_four_point_on_circle = False
End Function
Public Function is_item0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
  ByVal sig$, ByVal in1%, ByVal in2%, ByVal in3%, ByVal in4%, _
    ByVal il1%, ByVal il2%, no%, no1%, no2%, no3%, _
     v$, it As item0_data_type) As Boolean
Dim i%
Dim ty As Byte
Dim l(1) As Integer
Dim n(1) As Integer
it = set_item0_(p1%, p2%, p3%, p4%, sig$, in1%, in2%, in3%, in4%, _
                 il1%, il2%, "", "", "", 0, v$, condition_data0)
is_item0 = is_item0_(it, no%, no1%, no2%, no3%)
End Function
Function is_three_coline(ByVal p1%, ByVal p2%, ByVal p3%, l%, _
   op1%, op2%, op3%) As Boolean
Dim i%, j%, n%
If p1% = p2% Or p2% = p3% Or p3% = p1% Then
is_three_coline = True
Exit Function
End If
l% = line_number0(p1%, p2%, op1, op2%)
If line_number0(p2%, p3%, op2%, op3%) = l% Then
 is_three_coline = True
Else
 l% = 0
 op1% = 0
 op2% = 0
 op3% = 0
 is_three_coline = False
End If
End Function
Function is_relation(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
      ByVal in1%, ByVal in2%, ByVal in3%, ByVal in4%, ByVal il1%, _
       ByVal il2%, value As String, no%, no1%, no2%, no3%, no4%, _
         dr As relation_data0_type, tn1%, tn2%, _
         ty As Byte, c_data As condition_data_type, con_ty As Byte) As Boolean
'ty 返回类型　t 是否要排序号,con_ty =0 condition,con_ty=1 conclusion
Dim i%, t_n%, tn3%, t_no%
Dim v1$, v2$
Dim tv$
Dim ts As String
Dim ty1 As Byte
Dim ty2 As Boolean
Dim ty3 As Byte
Dim t_y(2) As Byte
Dim t_re As relation_data0_type
Dim is_no_initial As Integer
Dim tc_data As condition_data_type
'Dim dr As relation_type
'Dim mp As mid_point_type
'Dim el As eline_type
Dim temp_record As record_data_type
temp_record.data0.condition_data = c_data
no% = 0
tn1% = 0
tn2% = 0
ty = 0
no2% = 0
no3% = 0
no4% = 0
ts = value
dr = t_re
prove_type = 0
dr = t_re
If ts = "" Or no1% = -1000 Then
 ts = "y"
End If
Call arrange_four_point(p1%, p2%, p3%, p4%, in1%, in2%, _
      in3%, in4%, il1%, il2%, dr.poi(0), dr.poi(1), _
       dr.poi(2), dr.poi(3), 0, 0, dr.n(0), dr.n(1), _
        dr.n(2), dr.n(3), 0, 0, dr.line_no(0), dr.line_no(1), 0, _
         ty1, tc_data, is_no_initial)
        If is_no_initial = 1 And no1% = 0 Then
         Call add_record_to_record(tc_data, c_data)
          temp_record.data0.condition_data = c_data
        End If
        If con_ty = 0 Then '条件
         If ty1 > 2 Then
          If con_ty = 0 Then
            dr.ty = 3
          Else
            dr.ty = ty1
          End If
         End If
        Else '结论
          dr.ty = ty1
        End If
If ts <> "" And ts <> "0" Then
 Call ratio_value1(ts, ty1, dr.value)
End If
If ty1 > 2 Then
 dr.poi(4) = dr.poi(0)
  dr.poi(5) = dr.poi(3)
   dr.n(4) = dr.n(0)
    dr.n(5) = dr.n(3)
     dr.line_no(2) = dr.line_no(0)
End If
If dr.poi(0) = dr.poi(2) And dr.poi(1) = dr.poi(3) Then
          Call ratio_value("1", ty1, value)
            tn1% = 0
             ty = 0
              is_relation = True
               Exit Function
End If
If InStr(1, dr.value, "F", 0) > 0 Then
 If no1% = -1000 Then '判断
  is_relation = False
 Else
  no% = 0
   is_relation = True
 End If
  Exit Function
End If
If search_for_relation(dr, 0, t_no%, 1) Then  '5.7
    no% = Drelation(t_no% + 1).data(0).record.data1.index.i(0)
    If InStr(1, dr.value, "y", 0) = 0 Then
               If dr.value <> Drelation(no%).data(0).data0.value And _
                  dr.value <> Drelation(no%).data(0).data0.value_ Then
                    no% = 0
                     is_relation = True
               End If
    Else
               value = solve_general_equation(dr.value, Drelation(no%).data(0).data0.value, _
                             "y")
               If value = "F" Then
                If no1% = 0 Then
                  is_relation = True
                   no% = 0
                    Exit Function
                Else
                   no1% = t_no%
                    GoTo is_relation_mark3:
                End If
               End If
    End If
      dr.value = Drelation(no%).data(0).data0.value
 If no1% = -5000 Then
    no1% = t_no%
     is_relation = True
     GoTo is_relation_mark6
 End If
             ty = relation_
              dr = Drelation(no%).data(0).data0
         'Call ratio_value(dr.value, ty1, value)
    If set_or_prove = 2 Then '
     If Drelation(i%).data(0).record.data1.is_proved = 2 Then
      is_relation = True
     ElseIf Drelation(no%).data(0).record.data1.is_proved = 1 Then
     ElseIf Drelation(no%).data(0).record.data1.is_proved = 0 Then
     End If
    Else
     is_relation = True
     ' Call set_level(re)
    End If
    'Call ratio_value(ovalue, ty1, Value)
     Exit Function
Else
 is_relation = False
 If no1% = -5000 Then
  no1% = t_no%
   GoTo is_relation_mark6
 End If
End If
If dr.value = "1" Or InStr(1, dr.value, "y", 0) > 0 Then
If ty1 > 2 Then
Dmid_point_data0.poi(0) = dr.poi(0)
 Dmid_point_data0.poi(1) = dr.poi(1)
  Dmid_point_data0.poi(2) = dr.poi(3)
If search_for_mid_point(Dmid_point_data0, 0, no%, 0) Then '4.7
          ty = midpoint_
    If set_or_prove = 2 Then '
     If Deline(no%).data(0).record.data1.is_proved = 1 Then
      is_relation = True
     End If
    Else
     is_relation = True
    End If
    dr.value = "1"
   Call ratio_value(dr.value, ty1, value)
    'Call set_level(re)
  Exit Function
ElseIf dr.value = "1" Then
 ty = midpoint_
  no% = 0
   Exit Function
End If
'Next i%
End If
eline_data0.poi(0) = dr.poi(0)
 eline_data0.poi(1) = dr.poi(1)
  eline_data0.poi(2) = dr.poi(2)
   eline_data0.poi(3) = dr.poi(3)
    eline_data0.line_no(0) = dr.line_no(0)
     eline_data0.line_no(1) = dr.line_no(1)
If search_for_eline(eline_data0, 0, no%, 0) Then
    'no% = t_no%
    ty = eline_
    If set_or_prove = 2 Then '
     If Deline(no%).data(0).record.data1.is_proved = 1 Then
      is_relation = True
     End If
    Else
     is_relation = True
    End If
   If InStr(1, dr.value, "y", 1) > 0 Then
    value = solve_general_equation(dr.value, "1", "y")
   End If   'Call set_level(re)
    Exit Function
 ElseIf dr.value = "1" Then
  ty = eline_
   no% = 0
    Exit Function
 End If
End If
 '*****************************
t_y(0) = is_line_value(dr.poi(0), dr.poi(1), dr.n(0), _
          dr.n(1), dr.line_no(0), "", tn1%, -1000, 0, 0, 0, _
             line_value_data0)
 t_y(1) = is_line_value(dr.poi(2), dr.poi(3), dr.n(2), _
            dr.n(3), dr.line_no(1), "", tn2%, -1000, 0, 0, 0, _
              line_value_data0)
  If ty1 > 2 Then
   t_y(2) = is_line_value(dr.poi(0), dr.poi(3), dr.n(0), _
         dr.n(3), dr.line_no(0), "", tn3%, -1000, 0, 0, 0, _
             line_value_data0)
  Else
   t_y(2) = 0
  End If
ty = line_value_
If t_y(0) = 1 Then
  If t_y(1) = 1 Or t_y(2) = 1 Then
   If t_y(1) = 1 Then 'And line_value(tn1%).data(0).data0.value = line_value(tn2%).data(0).data0.value Then
     tv$ = divide_string(line_value(tn1%).data(0).data0.value, _
                    line_value(tn2%).data(0).data0.value, True, False)
        If InStr(1, dr.value, "y", 0) > 0 Then
         value = solve_general_equation(dr.value, tv$, "y")
        End If   'Call set_level(re)
        dr.value = tv
       Call add_conditions_to_record(line_value_, tn1%, tn2%, 0, c_data)
 is_relation = True
       If value = "F" Then
        If no1% = -1000 Then
         is_relation = False
        End If
       End If
          no% = 0
         ' Call set_level(re)
        Exit Function
   ElseIf t_y(2) = 1 Then ' And line_value(tn3%).data(0).data0.value = _
                    time_string("2", line_value(tn1%).data(0).data0.value) Then
          ' tn1% = tn1%
          tn2% = tn3%
    tv$ = divide_string(line_value(tn1%).data(0).data0.value, _
                 line_value(tn2%).data(0).data0.value, True, False)
     tv$ = divide_string(tv$, minus_string("1", tv$, False, False), _
                   True, False)
     If InStr(1, dr.value, "y", 0) > 0 Then
      value = solve_general_equation(dr.value, tv$, "y")
     End If   'Call set_level(re)
        dr.value = tv$
     ' Call ratio_value(dr.value, ty1, value)
       Call add_conditions_to_record(line_value_, tn1%, tn3%, 0, c_data)
       is_relation = True
       is_relation = True
         If value = "F" Then
           If no1% = -1000 Then
            is_relation = False
           End If
          End If
        no% = 0
         'Call set_level(re)
         Exit Function
   End If
  End If
ElseIf t_y(1) = 1 And t_y(2) = 1 Then
      tv$ = minus_string(divide_string(line_value(tn3%).data(0).data0.value, _
         line_value(tn2%).data(0).data0.value, False, False), "1", True, False)
          If InStr(1, dr.value, "y", 0) > 0 Then
            value = solve_general_equation(dr.value, tv$, "y")
          End If
           dr.value = tv$
         ' Call ratio_value(dr.value, ty1, value)
           no% = 0
            tn1% = tn2%
             tn2% = tn3%
       Call add_conditions_to_record(line_value_, tn2%, tn3%, 0, c_data)
        is_relation = True
        ' Call set_level(re)
         Exit Function
End If
If no1% = 0 And ts <> "" Then '已知比值和线长,设置relation
 If t_y(0) = 1 Then
  Call add_conditions_to_record(line_value_, tn1%, 0, 0, c_data)
   dr.poi(0) = dr.poi(2)
    dr.poi(1) = dr.poi(3)
   dr.n(0) = dr.n(2)
    dr.n(1) = dr.n(3)
   dr.line_no(0) = dr.line_no(1)
     Call ratio_value1(ts, ty1, dr.value)
     Call add_conditions_to_record(line_value_, tn1%, 0, 0, c_data)
      is_relation = False
       dr.value = divide_string(line_value(tn1%).data(0).data0.value, dr.value, True, False)
        Exit Function
 ElseIf t_y(1) = 1 Then
    Call add_conditions_to_record(line_value_, tn2%, 0, 0, c_data)
     Call ratio_value1(ts, ty1, dr.value)
      dr.value = time_string(line_value(tn2%).data(0).data0.value, dr.value, True, False)
       tn1% = tn2%
       is_relation = False
       Call add_conditions_to_record(line_value_, tn2%, 0, 0, c_data)
        Exit Function
ElseIf t_y(2) = 1 Then
         Call add_conditions_to_record(line_value_, tn3%, 0, 0, c_data)
    'op1% = op1%
     ' op2% = op3%
       Call ratio_value1(ts, ty1, dr.value)
         dr.value = divide_string(dr.value, add_string(dr.value, "1", False, False), False, False)
      dr.value = time_string(line_value(tn3%).data(0).data0.value, dr.value, True, False)
       tn1% = tn3%
       is_relation = False
       Call add_conditions_to_record(line_value_, tn3%, 0, 0, c_data)
         Exit Function
End If
End If
'******************************************************************
'****************************************************************
ty = relation_
If no1% = -1000 Then
 no% = 0
  Exit Function
Else
 no1% = t_no%
End If
is_relation_mark3:
 no1% = t_no%
is_relation_mark6:
  Call search_for_relation(dr, 1, no2%, 1) '5.7
   Call search_for_relation(dr, 2, no3%, 1)
    Call search_for_relation(dr, 3, no4%, 1)
End Function
Public Function is_dverti0(ByVal l1%, ByVal l2%) As Boolean
Dim i%
For i% = 1 To m_lin(l1%).data(0).in_verti(0).lin
 If m_lin(l1%).data(0).in_verti(i%).line_no = l2% Then
  is_dverti0 = True
   Exit Function
 End If
Next i%
End Function
Public Function is_dverti(ByVal l1%, ByVal l2%, n%, n1%, n2%, _
         n3%, outl1%, outl2%) As Boolean
Dim i%, tn%
Dim jud As Long
Dim t_coord(1) As POINTAPI
Dim vert As two_line_type
If n1% <= -1000 Then
If last_conditions.last_cond(1).verti_no = 0 Or l1% = 0 Or l2% = 0 Or l1% = l2% Then
 is_dverti = False
  Exit Function
End If
Else
 If l1% = 0 Or l2% = 0 Or l1% = l2% Then
 
  is_dverti = True
   Exit Function
 End If
End If
n% = 0
For i% = 1 To last_conditions.last_cond(1).same_three_lines_no
  If same_three_lines(i%).data(0).line_no(0) = l1% Or _
           same_three_lines(i%).data(0).line_no(1) = l1% Then
   l1% = same_three_lines(i%).data(0).line_no(2)
  End If
  If same_three_lines(i%).data(0).line_no(0) = l2% Or _
           same_three_lines(i%).data(0).line_no(1) = l2% Then
   l2% = same_three_lines(i%).data(0).line_no(2)
  End If
Next i%
'If l1% = l2% Or l1% = 0 Or l2% = 0 Then
'If n1% = -1000 Then
' is_dverti = False
'Else
'  error_of_wenti = 1
'   is_dverti = True
'End If
'n1% = 0
'n2% = 0
'Exit Function
'Else
t_coord(0) = minus_POINTAPI(m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                           m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate)
t_coord(1) = minus_POINTAPI(m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                           m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate)
jud = time_POINTAPI(verti_POINTAPI(t_coord(0)), t_coord(1))
't_coord(0).X = m_poi(lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
                               m_poi(lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X
't_coord(0).Y = m_poi(lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
                               m_poi(lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y
't_coord(1).X = m_poi(lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
                               m_poi(lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.X
't_coord(1).Y = m_poi(lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
                               m_poi(lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.Y
'jud = t_coord(0).X * t_coord(1).Y - t_coord(1).X * t_coord(0).Y
If jud < 0 Then
 outl1% = l1%
 outl2% = l2%
ElseIf jud > 0 Then
 outl1% = l2%
 outl2% = l1%
Else
 If n1% = -1000 Then
  is_dverti = False
   Exit Function
 Else
  is_dverti = True
   Exit Function
 End If
End If
'End If
 vert.line_no(0) = outl1%
  vert.line_no(1) = outl2%
   vert.inter_poi = is_line_line_intersect(vert.line_no(0), vert.line_no(1), 0, 0, False)
 If search_for_verti(vert, 0, n%, 1) Then '5.7
  If n1% = -5000 Then
   n1% = n%
  Else
   n% = Dverti(n% + 1).data(0).record.data1.index.i(0)
   If set_or_prove = 2 Then '
    If Dverti(n%).data(0).record.data1.is_proved = 1 Then
    is_dverti = True
    End If
   Else
    is_dverti = True
   End If
    Exit Function
  End If
  is_dverti = True
Else
   is_dverti = False
If n1% = -1000 Then
 n% = 0
  Exit Function
End If
  n1% = n%
End If
   Call search_for_verti(vert, 1, n2%, 1) '5.7
'   Call search_for_verti(vert, 2, n3%, 1)

End Function
Public Function is_equal_dline(ByVal p1%, ByVal p2%, ByVal p3%, _
  ByVal p4%, ByVal in1%, ByVal in2%, ByVal in3%, ByVal in4%, _
   ByVal il1%, ByVal il2%, no%, no1%, no2%, no3%, no4%, _
    el_data As eline_data0_type, dn1%, dn2%, con_type As Byte, _
     value$, c_data As condition_data_type) As Boolean
'判断线段相等，true and no%=0 ,表示相同线段,-表示中点,t 是否要排序
Dim stri(1) As String
Dim ty As Boolean
Dim t_y(1) As Boolean
Dim ty1 As Byte
Dim d_n(1) As Integer
Dim n%, i%
Dim s$
Dim t_el_data As eline_data0_type
Dim mp As mid_point_data0_type
Dim is_no_initial As Integer
Dim tc_data As condition_data_type
'If no1% = -1000 Then
' c_data.condition_no = 0
'End If
el_data = t_el_data
no% = 0
dn1% = 0
dn2% = 0
'con_type = 0
prove_type = 0
If is_same_two_point(p1%, p2%, p3%, p4%) Then
 is_equal_dline = True
  n% = 0
   con_type = 0
    Exit Function
End If
If arrange_four_point(p1%, p2%, _
  p3%, p4%, in1%, in2%, in3%, in4%, il1%, il2%, el_data.poi(0), el_data.poi(1), _
    el_data.poi(2), el_data.poi(3), 0, 0, el_data.n(0), el_data.n(1), _
     el_data.n(2), el_data.n(3), 0, 0, el_data.line_no(0), el_data.line_no(1), _
      0, ty1, tc_data, is_no_initial) And ty1 >= 3 Then
      If is_no_initial = 1 And no1% = 0 Then
       Call add_record_to_record(tc_data, c_data)
      End If
If el_data.line_no(0) = el_data.line_no(1) Then
  If el_data.n(1) < el_data.n(3) Then
   If el_data.n(1) > el_data.n(2) Then
      Call exchange_two_integer(el_data.poi(1), el_data.poi(2))
      Call exchange_two_integer(el_data.n(1), el_data.n(2))
   End If
  End If
End If
If ty1 > 3 And ty1 <> 5 Then
If no1% = -1000 Then
 is_equal_dline = False
Else
 is_equal_dline = True
  no% = 0
  con_type = eline_
End If
   Exit Function
End If
Else
 If el_data.line_no(0) = 0 Or el_data.line_no(1) = 0 Then
  If no1% = -1000 Then
   is_equal_dline = False
  Else
   no% = 0
    is_equal_dline = True
  End If
   Exit Function
 End If
End If
  If no1% = -5000 Then
   GoTo is_equal_dline_mark3
  End If
If (el_data.line_no(0) = 0 Or el_data.line_no(1) = 0 Or _
    el_data.poi(0) = 0 Or el_data.poi(1) = 0 Or _
     el_data.poi(2) = 0 Or el_data.poi(3) = 0) And _
      no1% <> -1000 Then
 is_equal_dline = True
  no% = 0
  con_type = eline_
   Exit Function
End If
If no1% = -2000 Then
 con_type = eline_
  GoTo is_equal_dline_mark6
  'Exit Function
'ElseIf no1% = -5000 Then
'If t_el_data = el_data Then
 'is_equal_dline = True
'Else
'Call search_for_eline(el_data, 1, 0, no%, 1)
' no1% = 0
'  GoTo is_equal_dline_mark3
'End If
End If
t_y(0) = is_line_value(el_data.poi(0), el_data.poi(1), el_data.n(0), _
              el_data.n(1), el_data.line_no(0), "", _
               d_n(0), -1000, 0, 0, 0, line_value_data0)
 t_y(1) = is_line_value(el_data.poi(2), el_data.poi(3), el_data.n(2), _
              el_data.n(3), el_data.line_no(1), "", _
          d_n(1), -1000, 0, 0, 0, line_value_data0)
If t_y(0) And t_y(1) Then
 If line_value(d_n(0)).data(0).data0.value = line_value(d_n(1)).data(0).data0.value Then
  is_equal_dline = True
  dn1% = d_n(0)
   dn2% = d_n(1)
    If no1% < 0 Then
     Call add_conditions_to_record(line_value_, d_n(0), d_n(1), 0, c_data)
    End If
   con_type = line_value_
    Exit Function
 End If
End If
dn1% = 0
 dn2% = 0
p1% = el_data.poi(0)
p2% = el_data.poi(1)
p3% = el_data.poi(2)
p4% = el_data.poi(3)
in1% = el_data.n(0)
in2% = el_data.n(1)
in3% = el_data.n(2)
in4% = el_data.n(3)
is_equal_dline_mark6:
If el_data.line_no(0) = el_data.line_no(1) Then
 If ty1 = 3 Or ty1 = 5 Then
         con_type = midpoint_
If no1% = -2000 Then
 cond_type = midpoint_
  Exit Function
End If
  mp.poi(0) = el_data.poi(0)
   mp.poi(1) = el_data.poi(1)
    mp.poi(2) = el_data.poi(3)
   If search_for_mid_point(mp, 0, no%, 0) Then   '4.7
          is_equal_dline = True
           'Call set_level(re)
     Exit Function
   Else
   is_equal_dline = False
    no% = 0
     Exit Function
   End If
 ElseIf el_data.n(1) > el_data.n(2) Then
 Call exchange_two_integer(el_data.n(1), el_data.n(2))
 Call exchange_two_integer(el_data.poi(1), el_data.poi(2))
  'Next i%
 End If
End If
con_type = eline_
is_equal_dline_mark3:
If search_for_eline(el_data, 0, no%, 0) Then '5.7
    is_equal_dline = True
     'Call set_level(re)
 If no1% = -5000 Then
 no1% = 0
  Call search_for_eline(el_data, 0, no1%, 1)
   GoTo is_equal_dline_mark5
 Else
     Exit Function
 End If
Else
 'If no1% = -5000 Then
  no1% = no%
   GoTo is_equal_dline_mark5
 If no1% = -1000 Then
  Exit Function
 End If
End If
If t_y(0) And t_y(1) = False Then
If no1% >= 0 Then
Call add_conditions_to_record(line_value_, d_n(0), 0, 0, c_data)
End If
 con_type = line_value_
  is_equal_dline = False
el_data.poi(0) = p3%
 el_data.poi(1) = p4%
el_data.n(0) = in3%
 el_data.n(1) = in4%
el_data.line_no(0) = el_data.line_no(1)
  value = line_value(d_n(0)).data(0).data0.value
   dn1% = d_n(0)
   Exit Function
ElseIf t_y(1) And t_y(0) = False Then
If no1% >= 0 Then
Call add_conditions_to_record(line_value_, d_n(1), 0, 0, c_data)
End If
 con_type = line_value_
  is_equal_dline = False
   el_data.poi(0) = p1%
    el_data.poi(1) = p2%
     el_data.n(0) = in1%
      el_data.n(1) = in2%
  value = line_value(d_n(1)).data(0).data0.value
      dn1% = d_n(1)
   Exit Function
End If
If ty1 = 3 Then
con_type = midpoint_
Else
con_type = eline_
End If
If no1% <> -1000 Then
is_equal_dline_mark5:
'no1% = no%
 Call search_for_eline(el_data, 1, no2%, 1)
 Call search_for_eline(el_data, 2, no3%, 1)
 Call search_for_eline(el_data, 3, no4%, 1)
 'Call search_for_eline(el_data, 4, no5%, 1)
 'Else
 'no% = 0
End If
'is_equal_dline = False
End Function
Function is_point_in_line1(in_coord As POINTAPI, p1%, p2%, _
     start%, t As Boolean, out_coord As POINTAPI, out_p%, nu!, is_change As Boolean) As Boolean
  'in_x%,in_y% 输入点的座标，p1%,p2%线的端点，out_x输出座标
  'nu! 点的分比,start% 出发的端点，t=0 平行t=1垂直
Dim s&, r&
Dim s_coord As POINTAPI
 r& = (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) ^ 2 + _
       (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) ^ 2
     '线长
 If r& < 4 Then
  is_point_in_line1 = False
 Else
  is_point_in_line1 = True
  If start% = 0 Then
    s_coord = m_poi(p1%).data(0).data0.coordinate
  Else
    s_coord = m_poi(start%).data(0).data0.coordinate
  End If
If t = False Then '平行
  s& = (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) * (in_coord.X - s_coord.X) + _
        (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) * (in_coord.Y - s_coord.Y)
  nu! = s& / r&
  t_coord.X = s_coord.X + (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) * nu!
  t_coord.Y = s_coord.Y + (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) * nu!
 ElseIf t Then
 
  s& = (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) * (in_coord.X - s_coord.X) - _
         (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) * (in_coord.Y - s_coord.Y)
  nu! = s& / r&
  t_coord.X = s_coord.X + (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) * nu!
  t_coord.Y = s_coord.Y - (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) * nu!
End If
    out_coord = t_coord
   If out_p% > 0 Then
    Call set_point_coordinate(out_p%, out_coord, is_change)
   End If
  End If
End Function





Public Function is_same_three_point(p1%, p2%, p3%, _
  p4%, p5%, p6%) As Boolean
If (p1% = p4% And is_same_two_point(p2%, p3%, p5%, p6%)) Or _
(p1% = p5% And is_same_two_point(p2%, p3%, p4%, p6%)) Or _
(p1% = p6% And is_same_two_point(p2%, p3%, p5%, p4%)) Then
is_same_three_point = True
Else
is_same_three_point = False
End If
End Function
Public Function is_same_four_point(p1%, p2%, p3%, _
  p4%, p5%, p6%, p7%, p8%) As Boolean
If (p1% = p5% And is_same_three_point(p2%, p3%, p4%, p6%, p7%, p8%)) Or _
(p1% = p6% And is_same_three_point(p2%, p3%, p4%, p5%, p7%, p8%)) Or _
(p1% = p7% And is_same_three_point(p2%, p3%, p4%, p6%, p5%, p8%)) Or _
(p1% = p8% And is_same_three_point(p2%, p3%, p4%, p6%, p7%, p5%)) Then
is_same_four_point = True
Else
is_same_four_point = False
End If
End Function

Public Function is_same_two_point(p1%, p2%, p3%, p4%) As Boolean
If (p1% = p3% And p2% = p4%) Or (p1% = p4% And p2% = p3%) Then
is_same_two_point = True
Else
is_same_two_point = False
End If

End Function
Public Function is_same_point(in_coord As POINTAPI, point_no%) As Boolean
 is_same_point = is_same_point_by_coord(in_coord, m_poi(point_no%).data(0).data0.coordinate)
End Function



Public Function is_point_pair(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
       ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, _
        ByVal in1%, ByVal in2%, ByVal in3%, ByVal in4%, ByVal in5%, _
         ByVal in6%, ByVal in7%, ByVal in8%, ByVal il1%, ByVal il2%, _
          ByVal il3%, ByVal il4%, no%, no1%, no2%, no3%, no4%, no5%, _
           no6%, dp As point_pair_data0_type, ty As Byte, n1%, _
            n2%, ty1 As Byte, ty2 As Byte, n3%, n4%, n5%, n6%, _
              re_v1 As String, re_v2 As String, re As record_data_type) As Boolean
'n3%,n4% 记录line_value 导出的相等
Dim t(9) As Boolean
Dim i%, k%
Dim tn(9) As Integer
Dim tn2(9) As Integer
Dim v(4) As String
Dim con_ty(9) As Byte
Dim t_dp As point_pair_data0_type
Dim tc_data(4) As condition_data_type
If p1% <= 0 Or p2% <= 0 Or p3% <= 0 Or p4% <= 0 Or _
    p5% <= 0 Or p6% <= 0 Or p7% <= 0 Or p8% <= 0 Then
 If no1% <= -1000 Then
  is_point_pair = False
 Else
  is_point_pair = True
 End If
  Exit Function
End If
no% = 0
ty = 0
ty1 = 0
ty2 = 0
n1 = 0
n2 = 0
n3 = 0
n4 = 0
n5 = 0
n6 = 0
dp = t_dp
'ty=1 1=2;ty =2 1=3;ty=3 3=4; ty =4 4=2;ty=5 1=2,3=4;by=6 1=3,2=4
If il1% = 0 Or run_type = 10 Then
dp.line_no(0) = line_number0(p1%, p2%, in1%, in2%)
Else
dp.line_no(0) = il1%
End If
If in1% <= in2% Then
dp.n(0) = in1%
dp.n(1) = in2%
dp.poi(0) = p1%
dp.poi(1) = p2%
Else
dp.n(0) = in2%
dp.n(1) = in1%
dp.poi(0) = p2%
dp.poi(1) = p1%
'call exchange_two_integer(p1%, p2%)
'call exchange_two_integer(n1%, n2%)
End If
If il2% = 0 Or run_type = 10 Then
dp.line_no(1) = line_number0(p3%, p4%, in3%, in4%)
Else
dp.line_no(1) = il2%
End If
If in3% <= in4% Then
dp.n(2) = in3%
dp.n(3) = in4%
dp.poi(2) = p3%
dp.poi(3) = p4%
Else
dp.n(2) = in4%
dp.n(3) = in3%
dp.poi(2) = p4%
dp.poi(3) = p3%
'call exchange_two_integer(p3%, p4%)
'call exchange_two_integer(n3%, n4%)
End If
If il3% = 0 Or run_type = 10 Then
dp.line_no(2) = line_number0(p5%, p6%, in5%, in6%)
Else
dp.line_no(2) = il3%
End If
If in5% <= in6% Then
dp.n(4) = in5%
dp.n(5) = in6%
dp.poi(4) = p5%
dp.poi(5) = p6%
Else
dp.n(4) = in6%
dp.n(5) = in5%
dp.poi(4) = p6%
dp.poi(5) = p5%
'call exchange_two_integer(p5%, p6%)
'call exchange_two_integer(n5%, n6%)
End If
If il4% = 0 Or run_type = 10 Then
dp.line_no(3) = line_number0(p7%, p8%, in7%, in8%)
Else
dp.line_no(3) = il4%
End If
If in7% <= in8% Then
dp.n(6) = in7%
dp.n(7) = in8%
dp.poi(6) = p7%
dp.poi(7) = p8%
Else
dp.n(6) = in8%
dp.n(7) = in7%
dp.poi(6) = p8%
dp.poi(7) = p7%
'call exchange_two_integer(p7%, p8%)
'call exchange_two_integer(n7%, n8%)
End If
If no1% = -3000 Then
 GoTo is_point_pair_mark6
End If
Call simple_point_pair(dp, dp)
If no1% = -2000 Then
 Exit Function '只排序
End If
is_point_pair_mark6:
If no1% = -5000 Then
 no1% = 0
End If
If no1% = -5000 Then 'simple_data_for_line
is_point_pair = search_for_point_pair(dp, 0, no1%, 1)
Call search_for_point_pair(dp, 0, no1%, 1)
Call search_for_point_pair(dp, 1, no2%, 1)
Call search_for_point_pair(dp, 2, no3%, 1)
Call search_for_point_pair(dp, 3, no4%, 1)
Call search_for_point_pair(dp, 4, no5%, 1)
Call search_for_point_pair(dp, 5, no6%, 1)
Exit Function
End If
If search_for_point_pair(dp, 0, no%, 0) Then '5.7
       ty = dpoint_pair_
 If set_or_prove = 2 And _
   Ddpoint_pair(i%).data(0).record.data1.is_proved = 2 Then '
 is_point_pair = True
 Else
 is_point_pair = True
 End If
  Call set_level(re.data0.condition_data)
  Exit Function
Else
 If no1% = -5000 Then
  no1% = no%
   GoTo is_point_pair_mark7
 End If
  'no1% = no%
End If
'******************************************************
'*************************************************************
t(0) = is_relation(dp.poi(0), dp.poi(1), dp.poi(2), dp.poi(3), _
        dp.n(0), dp.n(1), dp.n(2), dp.n(3), dp.line_no(0), dp.line_no(1), _
         v(0), tn(0), -1000, 0, 0, 0, _
          relation_data0, tn2(0), tn2(1), con_ty(0), tc_data(0), 0)
t(1) = is_relation(dp.poi(4), dp.poi(5), dp.poi(6), dp.poi(7), _
         dp.n(4), dp.n(5), dp.n(6), dp.n(7), dp.line_no(2), dp.line_no(3), _
          v(1), tn(1), -1000, 0, 0, 0, _
            relation_data0, tn2(2), tn2(3), con_ty(1), tc_data(1), 0)
t(2) = is_relation(dp.poi(0), dp.poi(1), dp.poi(4), dp.poi(5), _
         dp.n(0), dp.n(1), dp.n(4), dp.n(5), dp.line_no(0), dp.line_no(2), _
           v(2), tn(2), -1000, 0, 0, 0, _
            relation_data0, tn2(4), tn2(5), con_ty(2), tc_data(2), 0)
t(3) = is_relation(dp.poi(2), dp.poi(3), dp.poi(6), dp.poi(7), _
          dp.n(2), dp.n(3), dp.n(6), dp.n(7), dp.line_no(1), dp.line_no(3), _
           v(3), tn(3), -1000, 0, 0, 0, _
            relation_data0, tn2(6), tn2(7), con_ty(3), tc_data(3), 0)
If (dp.con_line_type(0) = 3 Or dp.con_line_type(0) = 5) And _
     (dp.con_line_type(1) = 3 Or dp.con_line_type(1) = 5) Then
t(4) = is_relation(dp.poi(8), dp.poi(9), dp.poi(10), dp.poi(11), _
          dp.n(8), dp.n(9), dp.n(10), dp.n(11), dp.line_no(0), dp.line_no(2), _
           v(4), tn(4), -1000, 0, 0, 0, _
            relation_data0, tn2(8), tn2(9), con_ty(4), tc_data(4), 0)
End If
'***********************************
If t(0) And t(1) And v(0) = v(1) And v(0) <> "" And v(0) <> "F" Then
 '1/2=3/4
 ty1 = con_ty(0)
 ty2 = con_ty(1)
  ty = 0
    n3% = tn2(0)
     n4% = tn2(1)
      n1% = tn(0)
    n5% = tn2(2)
     n6% = tn2(3)
      n2% = tn(1)
 is_point_pair = True
 Call add_conditions_to_record(con_ty(0), tn(0), tn2(0), tn2(1), re.data0.condition_data)
 Call add_conditions_to_record(con_ty(1), tn(1), tn2(2), tn2(3), re.data0.condition_data)
  Exit Function
ElseIf t(0) And t(4) And v(0) = v(4) And v(0) <> "" And v(0) <> "F" Then
 '1/2=3/4
 ty1 = con_ty(0)
 ty2 = con_ty(4)
  ty = 0
    n3% = tn2(0)
     n4% = tn2(1)
      n1% = tn(0)
    n5% = tn2(8)
     n6% = tn2(9)
      n2% = tn(4)
 is_point_pair = True
 Call add_conditions_to_record(con_ty(0), tn(0), tn2(0), tn2(1), re.data0.condition_data)
 Call add_conditions_to_record(con_ty(4), tn(4), tn2(8), tn2(9), re.data0.condition_data)
  Exit Function
ElseIf t(4) And t(1) And v(4) = v(1) And v(1) <> "" Then
 '1/2=3/4
 ty1 = con_ty(4)
 ty2 = con_ty(1)
  ty = 0
    n3% = tn2(8)
     n4% = tn2(9)
      n1% = tn(4)
    n5% = tn2(2)
     n6% = tn2(3)
      n2% = tn(1)
 is_point_pair = True
 Call add_conditions_to_record(con_ty(4), tn(4), tn2(8), tn2(9), re.data0.condition_data)
 Call add_conditions_to_record(con_ty(1), tn(1), tn2(2), tn2(3), re.data0.condition_data)
  Exit Function
ElseIf t(2) And t(3) And v(2) = v(3) And v(2) <> "" And v(2) <> "F" Then
'1/3=2/4
ty1 = con_ty(2)
ty2 = con_ty(3)
ty = 0
    n3% = tn2(4)
     n4% = tn2(5)
    n1% = tn(2)
'ty1 = con_ty(3)
    n5% = tn2(6)
     n6% = tn2(7)
    n2% = tn(3)
    is_point_pair = True
 Call add_conditions_to_record(con_ty(2), tn(2), tn2(4), tn2(5), re.data0.condition_data)
 Call add_conditions_to_record(con_ty(3), tn(3), tn2(6), tn2(7), re.data0.condition_data)
  Exit Function
ElseIf t(0) And v(0) <> "" Then '
 ty1 = con_ty(0)
  ty = 1
    n3% = tn2(0)
     n4% = tn2(1)
       n1% = tn(0)
     re_v1 = v(0)
      Call add_conditions_to_record(con_ty(0), tn(0), tn2(0), tn2(1), re.data0.condition_data)
   is_point_pair = False
 If t(2) = True And v(2) <> "" And v(2) <> "F" Then
 ty2 = con_ty(2)
  ty = 5
    n5% = tn2(4)
     n6% = tn2(5)
    n2% = tn(2)
    re_v2 = v(2)
 ElseIf t(3) And v(3) <> "" And v(3) <> "F" Then
  ty2 = con_ty(3)
  ty = 6
    n5% = tn2(6)
     n6% = tn2(7)
    n2% = tn(3)
     re_v2 = v(3)
 End If
ElseIf t(1) And v(1) <> "" And v(1) <> "F" Then
ty1 = con_ty(1)
 ty = 2
    n3% = tn2(2)
     n4% = tn2(3)
    n1% = tn(1)
    re_v1 = v(1)
      Call add_conditions_to_record(con_ty(1), tn(1), tn2(2), tn2(3), re.data0.condition_data)
 is_point_pair = False
 If t(2) = True And v(2) <> "" And v(2) <> "F" Then
 ty2 = con_ty(2)
  ty = 7
    n5% = tn2(4)
     n6% = tn2(5)
    n2% = tn(2)
    re_v2 = v(2)
 ElseIf t(3) And v(3) <> "" And v(3) <> "F" Then
  ty2 = con_ty(3)
  ty = 8
    n5% = tn2(6)
     n6% = tn2(7)
    n2% = tn(3)
     re_v2 = v(3)
 End If
ElseIf t(2) And v(2) <> "" And v(2) <> "F" Then
ty1 = con_ty(2)
 ty = 3
    n3% = tn2(4)
     n4% = tn2(5)
    n1% = tn(2)
    re_v1 = v(2)
      Call add_conditions_to_record(con_ty(2), tn(2), tn2(4), tn2(5), re.data0.condition_data)
is_point_pair = False
ElseIf t(3) And v(3) <> "" And v(3) <> "F" Then
ty1 = con_ty(3)
 ty = 3
    n3% = tn2(6)
     n4% = tn2(7)
    n1% = tn(3)
    re_v1 = v(3)
      Call add_conditions_to_record(con_ty(3), tn(3), tn2(6), tn2(7), re.data0.condition_data)
 is_point_pair = False
ElseIf t(4) And v(4) <> "" And v(4) <> "F" Then
ty1 = con_ty(4)
 ty = 9
    n3% = tn2(8)
     n4% = tn2(9)
      n1% = tn(4)
    re_v1 = v(4)
      Call add_conditions_to_record(con_ty(4), tn(5), tn2(8), tn2(9), re.data0.condition_data)
 is_point_pair = False
End If
If no1% > -1000 Then
 no1% = no%
is_point_pair_mark7:
  If dp.line_no(0) = dp.line_no(1) And dp.line_no(0) = dp.line_no(2) And dp.line_no(0) = dp.line_no(3) Then
        If dp.n(0) = dp.n(2) And dp.n(5) = dp.n(7) And _
          dp.n(1) = dp.n(4) And dp.n(5) = dp.n(6) Then
          dp.is_h_ratio = 1
        ElseIf dp.n(1) = dp.n(2) And dp.n(5) = dp.n(6) Then
          dp.n(8) = dp.n(0)
          dp.n(9) = dp.n(3)
          dp.n(10) = dp.n(4)
          dp.n(11) = dp.n(7)
          dp.poi(8) = dp.poi(0)
          dp.poi(9) = dp.poi(3)
          dp.poi(10) = dp.poi(4)
          dp.poi(11) = dp.poi(7)
          dp.line_no(4) = dp.line_no(0)
          dp.line_no(5) = dp.line_no(2)
          dp.is_h_ratio = 2
        ElseIf dp.n(1) = dp.n(2) And dp.n(4) = dp.n(7) Then
          dp.n(8) = dp.n(0)
          dp.n(9) = dp.n(3)
          dp.n(10) = dp.n(6)
          dp.n(11) = dp.n(5)
          dp.poi(8) = dp.poi(0)
          dp.poi(9) = dp.poi(3)
          dp.poi(10) = dp.poi(6)
          dp.poi(11) = dp.poi(5)
          dp.line_no(4) = dp.line_no(0)
          dp.line_no(5) = dp.line_no(2)
          dp.is_h_ratio = 2
        End If
  ElseIf dp.line_no(0) = dp.line_no(1) And dp.line_no(2) = dp.line_no(3) Then
        If dp.n(1) = dp.n(2) And dp.n(5) = dp.n(6) Then
          dp.n(8) = dp.n(0)
          dp.n(9) = dp.n(3)
          dp.n(10) = dp.n(4)
          dp.n(11) = dp.n(7)
          dp.poi(8) = dp.poi(0)
          dp.poi(9) = dp.poi(3)
          dp.poi(10) = dp.poi(4)
          dp.poi(11) = dp.poi(7)
          dp.line_no(4) = dp.line_no(0)
          dp.line_no(5) = dp.line_no(2)
          dp.is_h_ratio = 2
        ElseIf dp.n(1) = dp.n(2) And dp.n(4) = dp.n(7) Then
          dp.n(8) = dp.n(0)
          dp.n(9) = dp.n(3)
          dp.n(10) = dp.n(6)
          dp.n(11) = dp.n(5)
          dp.poi(8) = dp.poi(0)
          dp.poi(9) = dp.poi(3)
          dp.poi(10) = dp.poi(6)
          dp.poi(11) = dp.poi(5)
          dp.line_no(4) = dp.line_no(0)
          dp.line_no(5) = dp.line_no(2)
          dp.is_h_ratio = 2
        End If
  End If
   Call search_for_point_pair(dp, 1, no2%, 1) '5.7
   Call search_for_point_pair(dp, 2, no3%, 1)
    Call search_for_point_pair(dp, 3, no4%, 1)
     Call search_for_point_pair(dp, 4, no5%, 1)
      Call search_for_point_pair(dp, 5, no6%, 1)
Else
 no% = 0
End If
 is_point_pair = False
End Function
Public Function is_same_angle(A1 As angle_data_type, A2 As angle_data_type) As Boolean
Dim p(5) As Integer
If A1.poi(1) = A2.poi(1) And A1.line_no(0) = A2.line_no(0) And _
   A1.line_no(1) = A2.line_no(1) And A1.te(0) = A2.te(0) And A1.te(1) = A2.te(1) Then
      is_same_angle = True
Else
is_same_angle = False
End If
End Function

Public Function is_equal_angle(ByVal A1%, ByVal A2%, _
         n1%, n2%) As Boolean
'判断两角是否相等,true and no%=0  表示相同角
'no% = 0
Dim A3_v As angle3_value_data0_type
If A1% = 0 Or A2% = 0 Then
 is_equal_angle = False
  Exit Function
End If
n1% = 0
n2% = 0
record_0.data0.condition_data.condition_no = 0 ' record0
is_equal_angle = is_three_angle_value(A1%, A2%, 0, "1", "-1", "0", _
  "0", "0", n1%, n2%, 0, -1000, 0, 0, 0, 0, 0, 0, 0, A3_v, record_0.data0.condition_data, 0)
'*********************************************************
End Function

Public Function is_total_equal_Triangle(ByVal triAngle1%, _
     ByVal triAngle2%, ByVal di1%, ByVal di2%, n%, n1%, n2%, _
        n3%, t_triA As two_triangle_type, re As record_data_type, _
         is_find_conclusion As Byte) As Boolean
Dim tn%
Dim triA(1) As Integer
Dim i%
Dim temp_record
n% = 0
If triAngle1% = triAngle2% Then
 is_total_equal_Triangle = True
  n% = 0
  Exit Function
ElseIf triAngle1% < triAngle2% Then
   t_triA.triangle(0) = triAngle1%
    t_triA.triangle(1) = triAngle2%
     t_triA.direction = set_direction(di1%, di2%)
ElseIf triAngle1% > triAngle2% Then
   t_triA.triangle(0) = triAngle2%
    t_triA.triangle(1) = triAngle1%
     t_triA.direction = set_direction(di2%, di1%)
End If
If search_for_total_equal_triangle(t_triA, 0, n%, 0, is_find_conclusion) Then
   If n1% <> -1000 Then
    If t_triA.direction <> Dtotal_equal_triangle(n%).data(0).direction Then
       Call simple_two_two_triangle(t_triA, _
          Dtotal_equal_triangle(n%).data(0), re, _
           Dtotal_equal_triangle(n%).data(0).record, 0)
    End If
   Else
      If is_find_conclusion = 0 Then
       If t_triA.direction <> Dtotal_equal_triangle(n%).data(0).direction Then
        is_total_equal_Triangle = False
         Exit Function
        End If
      Else
               is_total_equal_Triangle = True
         Exit Function
      End If
    End If
   If set_or_prove = 2 Then
    If Dtotal_equal_triangle(i%).data(0).record.data1.is_proved = 1 Then '
    is_total_equal_Triangle = True
    End If
   Else
    is_total_equal_Triangle = True
   End If
      Exit Function
End If
If n1% <> -1000 Then
 n1% = n%
  Call search_for_total_equal_triangle(t_triA, 1, n2%, 1, is_find_conclusion)
'  Call search_for_total_equal_triangle(t_triA, 2, n3%, 1)
Else
 n% = 0
End If
'Next i
    is_total_equal_Triangle = False

End Function
Public Function is_verti_mid_line(ByVal p1%, ByVal p2%, ByVal p3%, _
         ByVal l%, no%, n1%, n2%, v_m_line As verti_mid_line_data0_type) As Boolean
Dim i%
no% = 0
v_m_line.line_no(0) = l%
v_m_line.line_no(1) = line_number0(p1%, p3%, v_m_line.n(0), v_m_line.n(2))
v_m_line.poi(1) = p2%
Call is_point_in_line3(p2%, m_lin(v_m_line.line_no(1)).data(0).data0, v_m_line.n(1))
If v_m_line.n(0) < v_m_line.n(2) Then
v_m_line.poi(0) = p1%
v_m_line.poi(2) = p3%
Else
Call exchange_two_integer(v_m_line.n(0), v_m_line.n(2))
v_m_line.poi(0) = p3%
v_m_line.poi(2) = p1%
End If
         If search_for_verti_mid_line(v_m_line, no%, 0, 0) Then
           is_verti_mid_line = True '5.7
         Else
           n1% = no%
           Call search_for_verti_mid_line(v_m_line, n2%, 1, 1)
           is_verti_mid_line = False '5.7
         End If
End Function

Public Function is_point_in_line3(ByVal p%, l As line_data0_type, n%) As Boolean
Dim i%
For i% = 1 To l.in_point(0)
If Abs(l.in_point(i%)) = p% Then
 is_point_in_line3 = True
  n% = i%
   Exit Function
End If
Next i%
is_point_in_line3 = False
End Function
Public Function is_point_in_paral(ByVal p%, ByVal l1%, l2%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).paral_no
If Dparal(i%).data(0).data0.line_no(0) = l1% Then
   is_point_in_paral = is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(1)).data(0).data0, 0)
    If is_point_in_paral = True Then
      l2% = Dparal(i%).data(0).data0.line_no(1)
        Exit Function
    End If
ElseIf Dparal(i%).data(0).data0.line_no(1) = l1% Then
   is_point_in_paral = is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(0)).data(0).data0, 0)
    If is_point_in_paral = True Then
      l2% = Dparal(i%).data(0).data0.line_no(0)
        Exit Function
    End If
End If
Next i%
is_point_in_paral = False
End Function
Public Function is_point_in_verti_line(ByVal p%, ByVal l1%, n%, l2%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).verti_no
If l1% = 0 Then
   If is_point_in_line3(p%, m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0, 0) Then
      is_point_in_verti_line = True
      n% = i%
      l2% = Dverti(i%).data(0).line_no(0)
      Exit Function
   ElseIf is_point_in_line3(p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0, 0) Then
      is_point_in_verti_line = True
      n% = i%
      l2% = Dverti(i%).data(0).line_no(1)
      Exit Function
   End If
Else
If Dverti(i%).data(0).line_no(0) = l1% Then
   If is_point_in_line3(p%, m_lin(Dverti(i%).data(0).line_no(1)).data(0).data0, 0) Then
    is_point_in_verti_line = True
      n% = i%
      l2% = Dverti(i%).data(0).line_no(1)
        Exit Function
    End If
ElseIf Dverti(i%).data(0).line_no(1) = l1% Then
   If is_point_in_line3(p%, m_lin(Dverti(i%).data(0).line_no(0)).data(0).data0, 0) Then
      is_point_in_verti_line = True
      n% = i%
      l2% = Dverti(i%).data(0).line_no(0)
        Exit Function
    End If
End If
End If
Next i%
is_point_in_verti_line = False
End Function
Public Function is_mid_point(ByVal p1%, p2%, ByVal p3%, _
    ByVal in1%, ByVal in2%, ByVal in3%, ByVal il%, no%, _
     no_1%, no_2%, no_3%, no_4%, no_5%, no_6%, no_7%, _
      mp As mid_point_data0_type, v As String, _
       con_ty As Byte, no1%, no2%, c_data As condition_data_type) As Boolean
Dim i%, tn1%, tn2%, tn3%
Dim ty(2) As Boolean
Dim tl(2) As Integer
Dim md As mid_point_data0_type
If last_conditions.last_cond(1).mid_point_no = 0 And no_1% = -1000 Then
 is_mid_point = False
  Exit Function
End If
mp = md
no% = 0
'mp.poi(1) = p2%
'mp.n(1) = in2%
If il% = 0 Or run_type = 10 Then
 tl(0) = line_number0(p1%, p3%, in1%, in3%)
 tl(1) = line_number0(p1%, p2%, in1%, in2%)
 If tl(0) > 0 And tl(1) > 0 Then
   If tl(0) = tl(1) Then
    il% = tl(0)
   Else
    If no_1% = -1000 Then
     is_mid_point = False
    Else 'if no_1%
     is_mid_point = True
    End If
     Exit Function
    End If
  Else
  If p2% > 0 Then
    is_mid_point = False
     Exit Function
  End If
 End If
Else
 If il% = 0 Then
 il% = tl(0)
 End If
 mp.line_no = il%
 If in1% < in3% Then
  mp.poi(0) = p1%
  mp.poi(1) = p2%
  mp.poi(2) = p3%
  mp.n(0) = in1%
  mp.n(1) = in2%
  mp.n(2) = in3%
 Else
  mp.poi(0) = p3%
  mp.poi(1) = p2%
  mp.poi(2) = p1%
  mp.n(0) = in3%
  mp.n(1) = in2%
  mp.n(2) = in1%
 End If
End If
'******************
If tl(0) > 0 And tl(1) > 0 Then
 il% = tl(0)
 mp.line_no = il%
 If in1% < in3% Then
  mp.poi(0) = p1%
  mp.poi(1) = p2%
  mp.poi(2) = p3%
  mp.n(0) = in1%
  mp.n(1) = in2%
  mp.n(2) = in3%
 Else
  mp.poi(0) = p3%
  mp.poi(1) = p2%
  mp.poi(2) = p1%
  mp.n(0) = in3%
  mp.n(1) = in2%
  mp.n(2) = in1%
 End If
ElseIf tl(0) > 0 Then
 il% = tl(0)
 mp.line_no = il%
 If in1% < in3% Then
  mp.poi(0) = p1%
  mp.poi(2) = p3%
  mp.n(0) = in1%
  mp.n(2) = in3%
   If search_for_mid_point(mp, 2, no%, 2) Then '5.7
       mp = Dmid_point(no%).data(0).data0
        p1% = mp.poi(0)
        p2% = mp.poi(1)
        p3% = mp.poi(2)
        con_ty = midpoint_
        is_mid_point = True
   End If
   Exit Function
 Else
  mp.poi(0) = p3%
  mp.poi(2) = p1%
  mp.n(0) = in3%
  mp.n(2) = in1%
   If search_for_mid_point(mp, 2, no%, 2) Then  '5.7
       mp = Dmid_point(no%).data(0).data0
        p1% = mp.poi(0)
        p2% = mp.poi(1)
        p3% = mp.poi(2)
         is_mid_point = True
   End If
   Exit Function
 End If
  If mp.n(2) - mp.n(0) = 2 Then
   mp.n(1) = mp.n(0) + 1
   mp.poi(1) = m_lin(mp.line_no).data(0).data0.in_point(mp.n(1))
   p2% = mp.poi(1)
   in2% = mp.n(1)
  ElseIf mp.n(2) - mp.n(0) = 1 Then
   If no_1% = -1000 Then
    is_mid_point = False
   Else
    is_mid_point = True
   End If
   Exit Function
  End If
 ElseIf tl(1) > 0 Then
 il% = tl(1)
 mp.line_no = il%
 If in1% < in2% Then
  mp.poi(0) = p1%
  mp.poi(1) = p2%
  mp.n(0) = in1%
  mp.n(1) = in2%
 Else
  mp.poi(1) = p2%
  mp.poi(2) = p1%
  mp.n(1) = in2%
  mp.n(2) = in1%
 End If
 End If
If mp.poi(1) = 0 Then
 GoTo is_mid_point_mark3
End If
If no_1% = -5000 Then
  no_1% = 0
   GoTo is_mid_point_mark3
End If

is_mid_point_mark4:
If no1% = -2000 Then
 Exit Function
ElseIf no1% = -3000 Then
 GoTo is_mid_point_mark3
End If
ty(0) = is_line_value(mp.poi(0), mp.poi(1), mp.n(0), mp.n(1), _
             mp.line_no, "", tn1%, -1000, 0, 0, 0, _
              line_value_data0)
 ty(1) = is_line_value(mp.poi(1), mp.poi(2), mp.n(1), mp.n(2), _
             mp.line_no, "", tn2%, -1000, 0, 0, 0, _
            line_value_data0)
  ty(2) = is_line_value(mp.poi(0), mp.poi(2), mp.n(0), mp.n(2), _
             mp.line_no, "", tn3%, -1000, 0, 0, 0, _
             line_value_data0)
 con_ty = line_value_
If ty(0) = 1 And tl(0) = 1 Then
  If line_value(tn1%).data(0).data0.value = line_value(tn2%).data(0).data0.value Then
    no1% = tn1%
     no2% = tn2%
       is_mid_point = True
        If no1% < 0 Then
         Call add_conditions_to_record(line_value_, tn1%, tn2%, 0, c_data)
        End If
        Exit Function
  End If
ElseIf ty(0) = 1 And ty(2) = 1 Then
    If line_value(tn3%).data(0).data0.value = _
                    time_string("2", line_value(tn1%).data(0).data0.value, True, False) Then
       no1% = tn1%
        no2% = tn3%
       is_mid_point = True
       If no1% < 0 Then
         Call add_conditions_to_record(line_value_, tn1%, tn3%, 0, c_data)
       End If
       Exit Function
    End If
ElseIf ty(1) = 1 And ty(2) = 1 Then
    If line_value(tn3%).data(0).data0.value = _
                    time_string("2", line_value(tn2%).data(0).data0.value, True, False) Then
       no1% = tn2%
        no2% = tn3%
       is_mid_point = True
       If no1% < 0 Then
         Call add_conditions_to_record(line_value_, tn3%, tn2%, 0, c_data)
       End If
       Exit Function
    End If
 End If
is_mid_point_mark3:
con_ty = midpoint_
 If search_for_mid_point(mp, 0, no%, 0) Then '5.7
  If mp.poi(0) > 0 And mp.poi(1) > 0 And mp.poi(2) > 0 Then
   If Dmid_point(no%).data(0).data0.poi(0) <> mp.poi(0) Or _
        Dmid_point(no%).data(0).data0.poi(1) <> mp.poi(1) Or _
          Dmid_point(no%).data(0).data0.poi(2) <> mp.poi(2) Then
           If no_1% = 0 Then
            no% = 0
              is_mid_point = True
           Else
              is_mid_point = False
           End If
           Exit Function
    End If
  ElseIf mp.poi(0) > 0 And mp.poi(2) > 0 Then
      If Dmid_point(no%).data(0).data0.poi(0) = mp.poi(0) And _
          Dmid_point(no%).data(0).data0.poi(2) = mp.poi(2) Then
            mp = Dmid_point(no%).data(0).data0
             p2% = Dmid_point(no%).data(0).data0.poi(1)
     Else
            If no_1% = 0 Then
            no% = 0
              is_mid_point = True
           Else
              is_mid_point = False
           End If
           Exit Function
      End If
  ElseIf mp.poi(0) > 0 And mp.poi(1) > 0 Then
      If Dmid_point(no%).data(0).data0.poi(0) = mp.poi(0) And _
          Dmid_point(no%).data(0).data0.poi(1) = mp.poi(1) Then
           mp = Dmid_point(no%).data(0).data0
            If p1% = 0 Then
             p1% = Dmid_point(no%).data(0).data0.poi(2)
            Else
             p3% = Dmid_point(no%).data(0).data0.poi(2)
            End If
      Else
            If no_1% = 0 Then
            no% = 0
              is_mid_point = True
           Else
              is_mid_point = False
           End If
           Exit Function
      End If
  ElseIf mp.poi(1) > 0 And mp.poi(2) > 0 Then
         If Dmid_point(no%).data(0).data0.poi(1) = mp.poi(1) And _
          Dmid_point(no%).data(0).data0.poi(2) = mp.poi(2) Then
           mp = Dmid_point(no%).data(0).data0
            If p1% = 0 Then
             p1% = Dmid_point(no%).data(0).data0.poi(0)
            Else
             p3% = Dmid_point(no%).data(0).data0.poi(0)
            End If
         Else
            If no_1% = 0 Then
            no% = 0
              is_mid_point = True
           Else
              is_mid_point = False
           End If
           Exit Function
      End If
  End If
  If no_1% = -5000 Then
  Call search_for_mid_point(mp, 0, no_1%, 1)
   GoTo is_mid_point_mark6
  End If
   If Dmid_point(no%).data(0).data0.poi(2) = mp.poi(2) Then
         mp.poi(1) = Dmid_point(no%).data(0).data0.poi(1)
     If set_or_prove = 2 Then '
      If Dmid_point(no%).data(0).record.data1.is_proved = 1 Or _
       Dmid_point(no%).data(0).record.data0.condition_data.condition_no = 0 Then
        is_mid_point = True
      End If
     Else
       is_mid_point = True
     End If
       Exit Function
  End If
 End If
If run_type = 1 Then
 If ty(0) = 1 Then
   mp.poi(0) = mp.poi(1)
    mp.n(0) = mp.n(1)
   mp.poi(1) = mp.poi(2)
    mp.n(1) = mp.n(2)
      If no1% >= 0 Then
              Call add_conditions_to_record(line_value_, tn1%, 0, 0, c_data)
      End If
     is_mid_point = False
      v = line_value(tn1%).data(0).data0.value
       con_ty = line_value_
        no1% = tn1%
        Exit Function
ElseIf ty(1) = 1 Then
      v = line_value(tn2%).data(0).data0.value
       If no1% >= 0 Then
               Call add_conditions_to_record(line_value_, tn2%, 0, 0, c_data)
       End If
      is_mid_point = False
        con_ty = line_value_
         no1% = tn2%
         Exit Function
ElseIf ty(2) = 1 Then
      v = divide_string(line_value(tn3%).data(0).data0.value, "2", True, False)
       If no1% >= 0 Then
              Call add_conditions_to_record(line_value_, tn3%, 0, 0, c_data)
       End If
       con_ty = line_value_
        is_mid_point = False
         no1% = tn3%
         Exit Function
End If
End If
If no_1% <= -1000 Then
 no% = 0
 Exit Function
End If
no_1% = no%
is_mid_point_mark6:
Call search_for_mid_point(mp, 1, no_2%, 1)
Call search_for_mid_point(mp, 2, no_3%, 1)
 'End If
End Function

Public Function is_similar_triangle(ByVal p1%, ByVal p2%, ByVal p3%, _
   ByVal p4%, ByVal p5%, ByVal p6%, no%, no1%, no2%, no3%, _
    t_triA As two_triangle_type, re As record_data_type, is_find_conclusion As Byte) As Boolean
Dim D1%, D2% 'ty As Boolean
Dim i%, A1%, A2%
'Dim t_triA As two_triangle_type
 A1% = triangle_number(p1%, p2%, p3%, 0, 0, 0, 0, 0, 0, D1%)
'Call initial_record(record_0)
 A2% = triangle_number(p4%, p5%, p6%, 0, 0, 0, 0, 0, 0, D2%)
is_similar_triangle = is_similar_triangle0(A1%, A2%, D1%, D2%, no%, _
      no1%, no2%, no3%, t_triA, re, 0, is_find_conclusion)
End Function
Public Function is_similar_triangle0(ByVal triAngle1%, _
    ByVal triAngle2%, ByVal D1%, ByVal D2%, n%, n1%, n2%, _
     n3%, t_triA As two_triangle_type, re As record_data_type, _
        ty As Byte, is_find_conclusion As Byte) As Boolean
'Dim triA(1) As Integer
Dim i%
Dim temp_record As record_data_type
If triAngle1% = triAngle2% Then
 is_similar_triangle0 = True
  n% = 0
  Exit Function '同一三角形
ElseIf triAngle1% < triAngle2% Then
 t_triA.triangle(0) = triAngle1%
  t_triA.triangle(1) = triAngle2%
   t_triA.direction = set_direction(D1%, D2%)
ElseIf triAngle1% > triAngle2% Then
 t_triA.triangle(0) = triAngle2%
  t_triA.triangle(1) = triAngle1%
   t_triA.direction = set_direction(D2%, D1%)
End If
If search_for_total_equal_triangle(t_triA, 0, n%, 0, is_find_conclusion) Then '是否全等
    If n1% <> -1000 Then
    If t_triA.direction <> Dtotal_equal_triangle(n%).data(0).direction Then
       Call simple_two_two_triangle(t_triA, _
          Dtotal_equal_triangle(n%).data(0), re, _
           Dtotal_equal_triangle(n%).data(0).record, 0)
    End If
    n% = 0
    End If
     is_similar_triangle0 = True
      ty = total_equal_triangle_
      Exit Function
Else
If search_for_similar_triangle(t_triA, 0, n%, 0, is_find_conclusion) Then '是否相似
    is_similar_triangle0 = True
    ty = similar_triangle_
 If n1% <> -1000 Then
  temp_record = re
   Call add_conditions_to_record(similar_triangle_, n%, 0, 0, temp_record.data0.condition_data)
    If t_triA.direction <> Dsimilar_triangle(n%).data(0).direction Then
       Call simple_two_two_triangle(t_triA, _
          Dsimilar_triangle(n%).data(0), temp_record, _
           Dsimilar_triangle(n%).data(0).record, 0)
    End If
 Else
    If is_find_conclusion = 0 Then
    If t_triA.direction <> Dsimilar_triangle(n%).data(0).direction Then
     is_similar_triangle0 = False
      Exit Function
    End If
    Else
         is_similar_triangle0 = True
      Exit Function
    End If
 End If
Else
If n1% = -1000 Then
 n% = 0
 Exit Function
End If
n1% = n%
Call search_for_similar_triangle(t_triA, 1, n2%, 1, is_find_conclusion)
'Call search_for_similar_triangle(t_triA, 2, n3%, 1)
 is_similar_triangle0 = False
End If
End If
End Function
Public Function is_angle_value(ByVal A%, v As String, v_ As String, n%, re As condition_data_type) As Boolean
Dim A3_v As angle3_value_data0_type
n% = 0
re.condition_no = 0
If A% = 0 Then
   is_angle_value = False
    Exit Function
End If
If angle(A%).data(0).value <> "" Then
 If v <> "" Then
    If angle(A%).data(0).value = v Then
     n% = angle(A%).data(0).value_no
         re = angle3_value(angle(A%).data(0).value_no).data(0).record.data0.condition_data
          is_angle_value = True
    Else
          is_angle_value = False
    End If
 Else
  v = angle(A%).data(0).value
   n% = angle(A%).data(0).value_no
    re = angle3_value(angle(A%).data(0).value_no).data(0).record.data0.condition_data
     is_angle_value = True
  End If
Else
is_angle_value = is_three_angle_value(A%, 0, 0, "1", "0", "0", _
   v, v_, n%, 0, 0, -1000, 0, 0, 0, 0, 0, 0, 0, A3_v, re, 0)
'a3_v.angle(0) = A%
'a3_v.value = V
End If
End Function
Public Function is_parallelogram(ByVal p1%, ByVal p2%, _
  ByVal p3%, ByVal p4%, tn%, n1%, poly4_no%, cond_ty As Byte) As Boolean
Dim i%
If last_conditions.last_cond(1).parallelogram_no = 0 And n1% = -1000 Then
 tn% = 0
   is_parallelogram = False
    Exit Function
End If
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, 0)
If poly4_no% = 0 Then
  If n1% <> -1000 Then
   is_parallelogram = True
  Else 'If n1% = -1000 Then
   is_parallelogram = False
  End If
  Exit Function
Else
If is_long_squre0(poly4_no%, tn%, n1%, cond_ty) Then
 is_parallelogram = True
  Exit Function
End If
If is_rhombus0(poly4_no%, tn%, n1%, cond_ty) Then
 is_parallelogram = True
  Exit Function
End If
cond_ty = parallelogram_
 Dpolygon4(poly4_no%).data(0).ty = parallelogram_
If search_for_parallelogram(poly4_no%, tn%, 0) Then
  is_parallelogram = True
Else
  is_parallelogram = False
  n1% = tn%
   tn% = 0
End If
End If
End Function

Public Function get_midpoint(ByVal p1%, ByVal p2%, ByVal p3%, _
        n1%, n0%, n2%, l%, no%) As Integer
Dim i%
Dim md As mid_point_data0_type
If p2% = 0 Then
l% = line_number0(p1%, p3%, n1%, n2%)
If n1 > n2% Then
Call exchange_two_integer(p1%, p3%)
End If
md.poi(0) = p1%
md.poi(2) = p3%
If search_for_mid_point(md, 2, no%, 2) Then '5.7原5
get_midpoint = Dmid_point(no%).data(0).data0.poi(1)
n0% = Dmid_point(no%).data(0).data0.n(1)
Exit Function
Else
get_midpoint = 0
End If
ElseIf p3% = 0 Then
l% = line_number0(p1%, p2%, n1, n0%)
If n1 > n0% Then
md.poi(1) = p2%
md.poi(2) = p1%
If search_for_mid_point(md, 1, no%, 2) Then '5.7原4
get_midpoint = Dmid_point(no%).data(0).data0.poi(0)
n2% = Dmid_point(no%).data(0).data0.n(0)
Exit Function
Else
get_midpoint = 0
End If
Else
md.poi(0) = p1%
md.poi(1) = p2%
If search_for_mid_point(md, 0, no%, 2) Then '5.7原3
get_midpoint = Dmid_point(no%).data(0).data0.poi(2)
l% = Dmid_point(no%).data(0).data0.line_no
n2% = Dmid_point(no%).data(0).data0.n(2)
Exit Function
Else
get_midpoint = 0
End If
End If
End If
no% = 0
  get_midpoint = 0
End Function

Public Function is_line_line_intersect(ByVal l1 As Integer, _
        ByVal l2 As Integer, n1%, n2%, is_set_reduce As Boolean) As Integer
'判断两直线是否有焦点，并读出点的序号
Dim tl(1) As line_data0_type
If l1 = 0 Or l2 = 0 Then
 Exit Function
ElseIf l1 > last_conditions.last_cond(1).line_no Or l2 > last_conditions.last_cond(1).line_no Then
 Exit Function
ElseIf m_lin(l1).data(0).other_no = m_lin(l2).data(0).other_no And m_lin(l1).data(0).other_no > 0 Then
Exit Function
Else '正常情况
tl(0) = m_lin(l1).data(0).data0
tl(1) = m_lin(l2).data(0).data0
For n1% = 1 To tl(0).in_point(0) '扫描两直线所有点，找出公共点，并不添加新交点
 For n2% = 1 To tl(1).in_point(0)
  If tl(0).in_point(n1%) = tl(1).in_point(n2%) Then
   is_line_line_intersect = tl(0).in_point(n1%) '输出点的序号
    Exit Function
  End If
Next n2%
Next n1%
n1% = 0
n2% = 0
End If
End Function

Public Function get_inter_point_line_line0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%) As Integer
Dim tl(1) As Integer
tl(0) = line_number0(p1%, p2%, 0, 0)
tl(1) = line_number0(p3%, p4%, 0, 0)
get_inter_point_line_line0 = is_line_line_intersect(tl(0), tl(1), 0, 0, False)
End Function

Public Function is_three_angle_value(ByVal A1%, _
  ByVal A2%, ByVal A3%, ByVal para1 As String, _
    ByVal para2 As String, ByVal para3 As String, v As String, _
      v_ As String, no1%, no2%, no3%, n1%, n2%, n3%, n4%, n5%, n6%, _
        n7%, n8%, A3_v As angle3_value_data0_type, re As condition_data_type, _
        t_y As Byte) As Boolean
         'a_A3_v.angle(0), a_A3_v.angle(1), a_A3_v.para(0) As String, a_A3_v.para(1) As String, _
           a_A3_v.value As String, con_ty As Byte) As Boolean
'A1%,para1,v 输入，A3_v.angle(0),opear1,A3_v.angle(0),输出，con_ty, 类型 t_y=0 输入三角和，t_y=1 导出三角和
Dim i%, tn%
Dim tA(2) As Integer
Dim ts As String
Dim ts1 As String
Dim tp(2) As String
Dim t_p(5) As String
Dim tp1(2) As String
Dim t_n(2) As Integer
Dim tv As String
Dim total_para As String
Dim ty As Byte
Dim max_no%
Dim tA3_v(1) As angle3_value_data0_type
Dim tA0%, tA1%, tA2%
Dim depend_no As Integer
Dim temp_record1 As record_data_type
Dim temp_record As record_data_type
Dim temp_record2 As total_record_type
Dim temp_v$
temp_record1.data0.condition_data = re
If n1% < 0 Or t_y = 10 Then
temp_record1.data0.condition_data.condition_no = 0
End If
no1% = 0
no2% = 0
no3% = 0
'排序
tA3_v(0).angle(0) = angle(A1%).data(0).other_no
 tA3_v(0).angle(1) = angle(A2%).data(0).other_no
  tA3_v(0).angle(2) = angle(A3%).data(0).other_no
tA3_v(0).para(0) = para1
 tA3_v(0).para(1) = para2
  tA3_v(0).para(2) = para3
If v = "" And n1% <> -2000 Then '不是设置结论
tA3_v(0).value = "y"
tA3_v(0).value_ = "y"
Else
tA3_v(0).value = v
tA3_v(0).value_ = v_
End If
A3_v = tA3_v(0)
'If need_arrange Then
Call reduce_to_used_angle(A3_v.angle(0), A3_v.para(0), A3_v.value, A3_v.value_, temp_record1, 1)
Call reduce_to_used_angle(A3_v.angle(1), A3_v.para(1), A3_v.value, A3_v.value_, temp_record1, 1)
Call reduce_to_used_angle(A3_v.angle(2), A3_v.para(2), A3_v.value, A3_v.value_, temp_record1, 1)
If A3_v.para(0) = "0" And A3_v.para(1) = "0" And A3_v.para(2) = "0" Then
   If InStr(1, A3_v.value, "y", 0) > 0 Then '
    v = solve_first_order_equation(A3_v.value, "0", "y")
     re = temp_record1.data0.condition_data
      If n1% = -1000 Then
      If temp_record1.data0.condition_data.condition_no = 1 Then
       no1% = temp_record1.data0.condition_data.condition(1).no
       no2% = 0
       no3% = 0
      ElseIf temp_record1.data0.condition_data.condition_no = 2 Then
       no1% = temp_record1.data0.condition_data.condition(1).no
       no2% = temp_record1.data0.condition_data.condition(2).no
       no3% = 0
      ElseIf temp_record1.data0.condition_data.condition_no = 3 Then
       no1% = temp_record1.data0.condition_data.condition(1).no
       no2% = temp_record1.data0.condition_data.condition(2).no
       no3% = temp_record1.data0.condition_data.condition(3).no
      End If
      End If
    is_three_angle_value = True
     Exit Function
   End If
End If
ts = ""
If (A3_v.angle(0) = 0 Or A3_v.angle(0) > A3_v.angle(1)) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
 Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If (A3_v.angle(1) = 0 Or A3_v.angle(1) > A3_v.angle(2)) And A3_v.angle(2) > 0 Then
 Call exchange_two_integer(A3_v.angle(1), A3_v.angle(2))
 Call exchange_string(A3_v.para(1), A3_v.para(2))
End If
If (A3_v.angle(0) = 0 Or A3_v.angle(0) > A3_v.angle(1)) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
 Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If A3_v.para(0) <> "0" And A3_v.para(0) <> "" Then
Call simple_multi_string0(A3_v.para(0), A3_v.para(1), A3_v.para(2), "0", ts, True)
Else
ts = "1"
End If
If A3_v.value <> "" And ts <> "1" Then
 A3_v.value = divide_string(A3_v.value, ts, True, False)
 A3_v.value_ = divide_string(A3_v.value_, ts, True, False)
End If
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no _
      Or A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If (angle(A3_v.angle(1)).data(0).total_no > angle(A3_v.angle(2)).data(0).total_no Or _
   A3_v.angle(1) = 0) And A3_v.angle(2) > 0 Then
Call exchange_two_integer(A3_v.angle(1), A3_v.angle(2))
 Call exchange_string(A3_v.para(1), A3_v.para(2))
End If
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no Or _
     A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If A3_v.angle(1) > 0 Then
Call combine_two_angle_with_para(A3_v.angle(0), A3_v.angle(1), A3_v.angle(3), _
        A3_v.angle_(3), A3_v.para(0), A3_v.para(1), A3_v.value, A3_v.value_, A3_v.ty(0), _
          A3_v.ty_(0), 1, temp_record1)
If A3_v.angle(2) > 0 Then
Call combine_two_angle_with_para(A3_v.angle(1), A3_v.angle(2), A3_v.angle(4), _
       A3_v.angle_(4), A3_v.para(1), A3_v.para(2), A3_v.value, A3_v.value_, ty, 0, 1, temp_record1)
Call combine_two_angle_with_para(A3_v.angle(0), A3_v.angle(2), 0, _
       0, A3_v.para(0), A3_v.para(2), A3_v.value, A3_v.value_, ty, 0, 1, temp_record1)
Call combine_two_angle_with_para(A3_v.angle(0), A3_v.angle(1), A3_v.angle(3), _
        A3_v.angle_(3), A3_v.para(0), A3_v.para(1), A3_v.value, A3_v.value_, A3_v.ty(0), _
          A3_v.ty_(0), 1, temp_record1)
End If
End If
   Call add_record_to_record(temp_record1.data0.condition_data, re)
 If A3_v.angle(0) = 0 And A3_v.angle(1) = 0 And A3_v.angle(2) = 0 Then
  If A3_v.value = "0" Then
          If n1% < 0 Then
           If re.condition_no = 1 Then
           no1% = re.condition(1).no
           no2% = 0
           n3% = 0
           End If
           If re.condition_no = 2 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           n3% = 0
           End If
           If re.condition_no = 3 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           no2% = re.condition(3).no
           End If
          End If
   is_three_angle_value = True
      Exit Function
  Else
     If InStr(1, A3_v.value, "y", 0) > 0 Then
      v = solve_first_order_equation(A3_v.value, "0", "y")
     ElseIf A3_v.value <> "0" Then
     temp_record2.record_data = temp_record1
        If n1% < 0 Then
           If re.condition_no = 1 Then
           no1% = re.condition(1).no
           n2% = 0
           n3% = 0
           End If
           If re.condition_no = 2 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           n3% = 0
           End If
           If re.condition_no = 3 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           no3% = re.condition(3).no
           End If
        ElseIf n1% >= 0 Then
          Call set_equation(A3_v.value, 0, temp_record2)
           is_three_angle_value = True
        End If
       Exit Function
     End If
      If InStr(1, v, "F", 0) > 0 Then
       is_three_angle_value = False
      Else
          If n1% < 0 Then
           If re.condition_no = 1 Then
           no1% = re.condition(1).no
           End If
           If re.condition_no = 2 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           End If
           If re.condition_no = 3 Then
           no1% = re.condition(1).no
           no2% = re.condition(2).no
           no3% = re.condition(3).no
           End If
          End If
       is_three_angle_value = True
      End If
     Exit Function
  End If
 End If
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no _
      Or A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If (angle(A3_v.angle(1)).data(0).total_no > angle(A3_v.angle(2)).data(0).total_no Or _
   A3_v.angle(1) = 0) And A3_v.angle(2) > 0 Then
Call exchange_two_integer(A3_v.angle(1), A3_v.angle(2))
 Call exchange_string(A3_v.para(1), A3_v.para(2))
End If
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no Or _
     A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
'设置首项系数
'***********************
ts = ""
Call simple_multi_string0(A3_v.para(0), A3_v.para(1), A3_v.para(2), "0", ts, True)
If A3_v.value <> "" And ts <> "1" Then
 A3_v.value = divide_string(A3_v.value, ts, True, False)
 A3_v.value_ = divide_string(A3_v.value_, ts, True, False)
End If
'is_three_angle_value_mark_0:
'Call reduce_to_used_angle(a3_v.angle(0), a3_v.para(0), a3_v.value, temp_record1)
'Call reduce_to_used_angle(a3_v.angle(1), a3_v.para(1), a3_v.value, temp_record1)
'Call reduce_to_used_angle(a3_v.angle(2), a3_v.para(2), a3_v.value, temp_record1)
'Call add_record_to_record(temp_record1.data0.condition_data, re.data0.condition_data,0)
'ts = ""
'Call simple_multi_string0(a3_v.para(0), a3_v.para(1), a3_v.para(2), "0", ts, True)

'If a3_v.value <> "" And ts <> "1" Then
 'a3_v.value = divide_string(a3_v.value, ts, True, False)
'End If
If A3_v.para(1) = "0" And A3_v.para(0) <> "0" Then
 If n1% = -1000 And v <> "" Then
  If angle(A3_v.angle(0)).data(0).value = divide_string(A3_v.value, A3_v.para(0), True, False) Then
   Call add_conditions_to_record(angle3_value_, angle(A3_v.angle(0)).data(0).value_no, 0, 0, _
         re)
   is_three_angle_value = True
    Exit Function
  Else
   is_three_angle_value = False
    Exit Function
  End If
 End If
End If
If InStr(1, A3_v.para(0), "F", 0) > 0 Or InStr(1, A3_v.para(1), "F", 0) > 0 And _
      InStr(1, A3_v.para(2), "F", 0) > 0 Or InStr(1, A3_v.value, "F", 0) > 0 Then
 If n1% = -1000 Or n1% = -5000 Then
  is_three_angle_value = False
 Else
  is_three_angle_value = True
   no1% = 0
   no2% = 0
   no3% = 0
 End If
  Exit Function
End If
If search_for_three_angle_value(A3_v, 0, no1%, 0) Then '5.7
 If InStr(1, A3_v.value, "y", 0) > 0 Or _
       InStr(1, angle3_value(no1%).data(0).data0.value, "y", 0) > 0 Then
  v = solve_first_order_equation(A3_v.value, angle3_value(no1%).data(0).data0.value, "y")
 Else
  If angle3_value(no1%).data(0).data0.para(0) <> A3_v.para(0) Or _
      angle3_value(no1%).data(0).data0.para(1) <> A3_v.para(1) Or _
       angle3_value(no1%).data(0).data0.para(2) <> A3_v.para(2) Or _
        (angle3_value(no1%).data(0).data0.value <> A3_v.value And _
          A3_v.value <> "") Then
    If n1 = -1000 Then
      is_three_angle_value = False
       Exit Function
    Else
      Call add_conditions_to_record(angle3_value_, no1%, 0, 0, re)
       Call solve_equation_for_angle3(no1%, 0, A3_v.angle(0), A3_v.angle(1), _
        A3_v.angle(2), angle3_value(no1%).data(0).data0.para(0), angle3_value(no1%).data(0).data0.para(1), _
         angle3_value(no1%).data(0).data0.para(2), angle3_value(no1%).data(0).data0.value, _
          A3_v.angle(1), A3_v.angle(2), A3_v.para(0), A3_v.para(1), _
           A3_v.para(2), A3_v.value, re)
            is_three_angle_value = True
             'Call set_level(re)
                Exit Function
     End If
   Else
     Call add_conditions_to_record(angle3_value_, no1%, 0, 0, re)
      is_three_angle_value = True
      ' Call set_level(re)
        Exit Function
   End If
   End If
  Else
   'no1% = 0
    is_three_angle_value = False
  End If
' Next i%
If n1% = -1000 Then
 no1% = 0
Else
 n1% = no1%
 'If a3_v.para(2) <> "0" Then
 Call combine_two_angle_with_para(A3_v.angle(0), A3_v.angle(1), A3_v.angle(3), _
        A3_v.angle_(3), A3_v.para(0), A3_v.para(1), A3_v.value, A3_v.value_, A3_v.ty(0), _
          A3_v.ty_(0), 1, temp_record1)
'Call combine_two_Tangle(a3_v.angle(0), a3_v.angle(1), a3_v.angle(3), a3_v.angle_(3),
'        a3_v.ty(0), a3_v.ty_(0), 0, 1)
 'End If
'If a3_v.angle(3) > 0 Then
'   a3_v.angle_(3) = a3_v.angle(3)
'  If reduce_to_used_angle(a3_v.angle(3), "", "", temp_record1, 1) = 0 Then
 '  a3_v.angle_(3) = a3_v.angle(3)
 ' End If
' End If
 'If a3_v.angle(4) > 0 Then
'   a3_v.angle_(4) = a3_v.angle(4)
 ' If reduce_to_used_angle(a3_v.angle(4), "", "", temp_record1, 1) = 0 Then
 '  a3_v.angle_(4) = a3_v.angle(4)
 ' End If
 'End If
 'If a3_v.angle(5) > 0 Then
  ' a3_v.angle_(5) = a3_v.angle(5)
  'If reduce_to_used_angle(a3_v.angle(5), "", "", temp_record1, 1) = 0 Then
  ' a3_v.angle_(5) = a3_v.angle(5)
  'End If
 'End If

  'If a3_v.ty(0) <> 3 And a3_v.ty(0) <> 5 Then
   ' a3_v.ty(0) = 0
    'a3_v.angle(3) = 0
  'End If
 'If a3_v.angle(2) > 0 Then
 'Call combine_two_angle(a3_v.angle(1), a3_v.angle(2), 0, 0, 0, a3_v.angle(4), _
        a3_v.ty(1), 0, 0)
  'If a3_v.ty(1) <> 3 And a3_v.ty(1) <> 5 Then
   ' a3_v.ty(1) = 0
    'a3_v.angle(4) = 0
  'End If
 'Call combine_two_angle(a3_v.angle(2), a3_v.angle(0), 0, 0, 0, a3_v.angle(5), _
  '      a3_v.ty(2), 0, 0)
  'If a3_v.ty(2) <> 3 And a3_v.ty(2) <> 5 Then
   ' a3_v.ty(2) = 0
    'a3_v.angle(5) = 0
  'End If
 'End If
 If A3_v.para(1) = "0" Then
    A3_v.no_zero_angle = 1
 ElseIf A3_v.para(2) = "0" Then
    A3_v.no_zero_angle = 2
 Else
    A3_v.no_zero_angle = 3
 End If
 Call search_for_three_angle_value(A3_v, 1, n2%, 1) '5.7
 Call search_for_three_angle_value(A3_v, 2, n3%, 1)
 Call search_for_three_angle_value(A3_v, 3, n4%, 1)
 Call search_for_three_angle_value(A3_v, 4, n5%, 1) '5.7
 Call search_for_three_angle_value(A3_v, 5, n6%, 1)
 Call search_for_three_angle_value(A3_v, 6, n7%, 1)
 End If
End Function


Public Function is_three_point_on_line(ByVal p1%, ByVal p2%, _
        ByVal p3%, n%, n1%, n2%, n3%, c_data As condition_data_type, _
         op1%, op2%, op3%) As Boolean
Dim i%, j%, tl%, is_jud%
Dim ts$
Dim p3_l As three_point_on_line_data_type
Dim temp_record As record_data_type
is_jud% = n1%
 If p1% = 0 Or p2% = 0 Or p3% = 0 Then
  is_three_point_on_line = False
   Exit Function
 End If
'*******************************************
op1% = p1%
 op2% = p2%
  op3% = p3%
If line_number0(op1%, op2%, n1%, n2%, True) = line_number0(op2%, op3%, n2%, n3%, True) Then
If n1% < n3% And n3% < n2% Then
op1% = p1%
 op2% = p3%
  op3% = p2%
ElseIf n2% < n1% And n1% < n3% Then
op1% = p2%
 op2% = p1%
  op3% = p3%
ElseIf n2% < n3% And n3% < n1% Then
op1% = p2%
 op2% = p3%
  op3% = p1%
ElseIf n3% < n1% And n1 < n2% Then
op1% = p3%
 op2% = p1%
  op3% = p2%
ElseIf n3% < n2% And n2 < n1% Then
op1% = p3%
 op2% = p2%
  op3% = p1%
ElseIf n1% < n2% And n2% < n3% Then
op1% = p1%
 op2% = p2%
  op3% = p3%
End If
Else
If compare_two_point(m_poi(op1%).data(0).data0.coordinate, m_poi(op2%).data(0).data0.coordinate, _
      0, 0, 5) < 0 Then
      Call exchange_two_integer(op1%, op2%)
End If
If compare_two_point(m_poi(op2%).data(0).data0.coordinate, m_poi(op3%).data(0).data0.coordinate, _
      0, 0, 5) < 0 Then
      Call exchange_two_integer(op2%, op3%)
End If
If compare_two_point(m_poi(op1%).data(0).data0.coordinate, m_poi(op2%).data(0).data0.coordinate, _
      0, 0, 5) < 0 Then
      Call exchange_two_integer(op1%, op2%)
End If
End If
p3_l.poi(0) = op1%
 p3_l.poi(1) = op2%
  p3_l.poi(2) = op3%
If search_for_three_point_on_line(p3_l, 1, 0, n%, 0) Then
    If set_or_prove = 2 Then '
     If three_point_on_line(n%).data(0).record.data1.is_proved = 1 Then
            is_three_point_on_line = True
     End If
    Else
     is_three_point_on_line = True
    End If
    Exit Function
End If
If n1% = -1000 Then
 n% = 0
 Exit Function
End If
n1% = n%
 Call search_for_three_point_on_line(p3_l, 1, 1, n2%, 1)
 Call search_for_three_point_on_line(p3_l, 1, 2, n3%, 1)
End Function
Public Function is_line_value(ByVal p1%, ByVal p2%, _
            ByVal in1%, ByVal in2%, ByVal il%, v As String, _
             n%, n1%, n2%, n3%, n4%, lv As line_value_data0_type) As Byte
Dim tn1%
Dim t_lv As line_value_data0_type
lv = t_lv
n% = 0
If p1% = p2% Or p1% = 0 Or p2% = 0 Then
is_line_value = 0
Exit Function
'ElseIf last_conditions.last_cond(1).line_no_value = 0 And n1% = -1000 Then
 '   is_line_value = False
  '   Exit Function
Else
 If il% = 0 Or run_type = 10 Then
  il% = line_number0(p1%, p2%, in1%, in2%)
 End If
  lv.line_no = il%
   If in1% > in2% Then
     lv.n(0) = in2%
      lv.n(1) = in1%
     lv.poi(0) = p2%
      lv.poi(1) = p1%
   Else
     lv.n(0) = in1%
      lv.n(1) = in2%
     lv.poi(0) = p1%
      lv.poi(1) = p2%
   End If
     lv.value = v
     lv.value_ = v
   If n1% = -5000 Then
     n1% = 0
    GoTo is_line_value_mark3
   ElseIf n1% = -2000 Then
    Exit Function
   End If
End If
is_line_value_mark3:
If InStr(1, lv.value, "F", 0) > 0 Then
 If n1% = -1000 Or n1% = -5000 Then
    is_line_value = 0
 Else
    is_line_value = 1
     n% = 0
 End If
 Exit Function
End If
If search_for_line_value(lv, 0, n%, 0) Then '5.7
 If n1% = -5000 Then
  Call search_for_line_value(lv, 0, n1%, 1)
   GoTo is_line_value_mark6
 End If
  If v = "" Then
   v = line_value(n%).data(0).data0.value
  End If
  If set_or_prove = 2 Then '
   If line_value(n%).data(0).record.data1.is_proved = 1 Then
  is_line_value = 1
  End If
  Else
  is_line_value = 1
  End If
   Exit Function
Else
 If n1% = -1000 Then
  n% = 0
   Exit Function
 End If
n1% = n%
is_line_value_mark6:
     lv.squar_value = time_string(v, v, True, False)
     If InStr(1, lv.squar_value, "F", 0) > 0 Then
      n% = 0
       is_line_value = 1
        Exit Function
     End If
Call search_for_line_value(lv, 1, n2%, 1)
Call search_for_line_value(lv, 2, n3%, 1)
Call search_for_line_value(lv, 3, n4%, 1)
End If
End Function
Public Function is_V_line_value(ByVal p1%, ByVal p2%, _
            ByVal in1%, ByVal in2%, ByVal il%, v As String, _
             n%, n1%, n2%, n3%, n4%, lv As V_line_value_data0_type, is_initial As Boolean) As Byte
Dim i%, vl%, tn%
Dim dir As String
Dim s(3) As String
Dim t_lv As V_line_value_data0_type
lv = t_lv
n% = 0
If p1% = p2% Or p1% = 0 Or p2% = 0 Then
 is_V_line_value = 0
 Exit Function
Else
 vl% = vector_number(p1%, p2%, dir)
  lv.v_line = vl%
  lv.v_poi(0) = Dtwo_point_line(vl%).data(0).v_poi(0)
  lv.v_poi(1) = Dtwo_point_line(vl%).data(0).v_poi(1)
  If Dtwo_point_line(vl%).data(0).v_value <> "" Then
        v = time_string(Dtwo_point_line(vl%).data(0).v_value, _
           dir, True, False)
     n% = Dtwo_point_line(vl%).data(0).v_line_value_no
          is_V_line_value = True
           Exit Function
  End If
  If n1% <> -1000 Then
  If dir = "1" Or is_initial Then
   lv.value = v
  Else
   lv.value = time_string(v, "-1", True, False)
  End If
  End If
    If n1% = -5000 Then
     n1% = 0
    GoTo is_V_line_value_mark3
   ElseIf n1% = -2000 Then
    Exit Function
   End If
End If
is_V_line_value_mark3:
If InStr(1, lv.value, "F", 0) > 0 Then
 If n1% = -1000 Or n1% = -5000 Then
    is_V_line_value = 0
 Else
    is_V_line_value = 1
     n% = 0
 End If
 Exit Function
End If
 If search_for_V_line_value(lv, 0, n%, 0) Then  '5.7
       is_V_line_value = True
       lv = V_line_value(n%).data(0)
        If dir = 1 Then
         v = lv.value
        Else
         v = time_string(lv.value, "-1", True, False)
        End If
         Exit Function
  Else
   If n1% <> -1000 Then
   If string_type(lv.value, "", s(0), s(1), "") = 3 Then
      lv.unit_value = s(0)
   Else
      lv.unit_value = lv.value
   End If
   Call remove_brace(lv.unit_value)
   Call do_factor1(lv.unit_value, s(0), s(1), s(2), s(3), tn%)
   For i% = 0 To 3
    If InStr(1, s(i%), "U", 0) > 0 Or InStr(1, s(i%), "V", 0) > 0 Then
       lv.unit_value = s(i%)
        GoTo is_v_line_value_mark5
    End If
   Next i%
is_v_line_value_mark5:
   n1% = n%
   n% = 0
   Call search_for_V_line_value(lv, 1, n2%, 1)
   Call search_for_V_line_value(lv, 2, n3%, 1)
  End If
  End If
End Function
Public Function is_two_angle_value(ByVal A1%, ByVal A2%, _
        ByVal s1 As String, ByVal S2 As String, ByVal value$, ByVal value_$, no1%, no2%, _
           oA1%, oA2%, os1 As String, os2 As String, _
           ovalue$) As Boolean
'当成立时,且是角值0a1%,oa2%输出角值号
Dim A3_v As angle3_value_data0_type
record_0.data0.condition_data.condition_no = 0 ' record0
'***********************************
is_two_angle_value = is_three_angle_value(A1%, A2%, 0, s1, S2, "0", value$, value_$, no1%, _
   no2%, 0, -1000, 0, 0, 0, 0, 0, 0, 0, A3_v, record_0.data0.condition_data, 0)
oA1% = A3_v.angle(0)
oA2% = A3_v.angle(1)
os1 = A3_v.para(0)
os2 = A3_v.para(1)
ovalue$ = A3_v.value
End Function

Public Function is_total_equal_triangle1(ByVal p1%, ByVal p2%, ByVal p3%, _
   ByVal p4%, ByVal p5%, ByVal p6%, no%, n1%, n2%, n3%, _
     t_triA As two_triangle_type, re As record_data_type, is_find_conclusion As Byte) As Boolean
'Dim ty As Boolean
Dim i%, D1%, D2%, A1%, A2%
no% = 0
 A1% = triangle_number(p1%, p2%, p3%, 0, 0, 0, 0, 0, 0, D1)
 A2% = triangle_number(p4%, p5%, p6%, 0, 0, 0, 0, 0, 0, D2)
If A1 = 0 Or A2 = 0 Then
 is_total_equal_triangle1 = False
  Exit Function
ElseIf A1 = A2% Then
 is_total_equal_triangle1 = True
  Exit Function
End If
If A1 > A2 Then
Call exchange_two_integer(A1, A2)
 Call exchange_two_integer(D1, D2)
End If
is_total_equal_triangle1 = is_total_equal_Triangle(A1%, A2%, D1, D2, _
  no%, n1%, n2%, n3%, t_triA, re, is_find_conclusion)
 '　只能判断不是已知相似形
End Function

Public Function is_arc_value(ByVal Ar%, value As String, n%) As Boolean
Dim i%, tn%
Dim arc_v As arc_value_data_type
 If Ar% = 0 Then
   is_arc_value = False
    Exit Function
 End If
 arc_v.arc = Ar%
 If search_for_arc_value(arc_v, 1, n%, 0) Then
     n% = tn%
      value = arc_value(n%).data(0).value
       is_arc_value = True
        Exit Function
  End If
End Function

Public Function is_equal_arc(ByVal Ar1%, ByVal Ar2%, _
  oAr1%, oAr2%, n%, n1%, n2%)
Dim i%, A1%, A2%
Dim A As angle_type
Dim e_arc As equal_arc_data_type
If last_conditions.last_cond(1).equal_arc_no = 0 And n1% = -1000 Then
 is_equal_arc = False
  Exit Function
End If
If Ar1% = Ar2% Then
 is_equal_arc = True
  Exit Function
End If

'A1% = three_point_direction(p1%, Circ(C%).data(0).center, p2%)
 'A2% = three_point_direction(p3%, Circ(C%).data(0).center, p4%)
'If A1% = -1 Then
 'call exchange_two_integer(p1%, p2%)
'End If
'If A2% = -1 Then
 'call exchange_two_integer(p3%, p4%)
'End If
'If A1% = 0 Or A2% = 0 Then
 'is_equal_arc = False
  'Exit Function
'End If
If Ar1% > Ar2% Then
 oAr1 = Ar2%
  oAr2% = Ar1%
Else
 oAr1 = Ar1%
  oAr2% = Ar2%
End If
e_arc.arc(0) = oAr1%
 e_arc.arc(1) = oAr2%
If arc(oAr1%).data(0).cir = arc(oAr2%).data(0).cir Then
   If arc(oAr1%).data(0).poi(0) = arc(oAr2%).data(0).poi(1) Then
      e_arc.arc(2) = arc_no(arc(oAr2%).data(0).poi(0), arc(oAr1%).data(0).cir, arc(oAr1%).data(0).poi(1))
   ElseIf arc(oAr1%).data(0).poi(1) = arc(oAr2%).data(0).poi(0) Then
      e_arc.arc(2) = arc_no(arc(oAr2%).data(0).poi(1), arc(oAr1%).data(0).cir, arc(oAr1%).data(0).poi(0))
   End If
End If
If search_for_equal_arc(e_arc, 1, 0, n%, 0) Then
'For i% = 1 To last_equal_arc
'If equal_arc(i%).data(0).poi(0) = op1% And equal_arc(i%).data(0).poi(1) = op2% And _
    equal_arc(i%).data(0).poi(2) = op3% And equal_arc(i%).data(0).poi(3) = op4% Then
     'equal_arc(i%).Circ = C% Then
 '     n% = i%
      is_equal_arc = True
        Exit Function
End If
If n1% <> -1000 Then
n1% = n%
Call search_for_equal_arc(e_arc, 1, 1, n2%, 1)
End If
'Next i%
End Function

Public Function P_diffrence(p1 As POINTAPI, p2 As POINTAPI, p3 As POINTAPI) As Long
P_diffrence = (p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2 + _
               (p3.X - p2.X) ^ 2 + (p3.Y - p2.Y) ^ 2 - _
                (p1.X - p3.X) ^ 2 - (p1.Y - p3.Y) ^ 2
End Function
Public Function is_two_line_value(ByVal p1%, ByVal p2%, _
 ByVal p3%, ByVal p4%, ByVal in1%, ByVal in2%, ByVal in3%, _
   ByVal in4%, ByVal il1%, ByVal il2%, ByVal para1 As String, _
     ByVal para2 As String, ByVal v As String, no%, n1%, n2%, n3%, n4%, _
         t_l_value As two_line_value_data0_type, cond_ty As Byte, _
           c_data As condition_data_type) As Byte
Dim i%, n%, tn_%, tn1%
Dim tn(3) As Integer
Dim tl(1) As Integer
Dim ts As String
Dim ty As Byte
Dim tp(3) As Integer
Dim depend_no As Integer
Dim temp_record As total_record_type
Dim t_l2_value As two_line_value_data0_type
Dim l_value(1) As line_value_data0_type
Dim re As relation_data0_type
Dim is_no_initial As Integer
Dim tc_data As condition_data_type
If p2% > 0 Then
tn1% = n1%
t_l_value = t_l2_value
If in1% = 0 Or in2% = 0 Then
il1% = line_number0(p1%, p2%, in1%, in2%)
If in1% > in2% Then
 Call exchange_two_integer(in1%, in2%)
 Call exchange_two_integer(p1%, p2%)
End If
End If
If in3% = 0 Or in4% = 0 Then
il2% = line_number0(p3%, p4%, in3%, in4%)
If in3% > in4% Then
 Call exchange_two_integer(in3%, in4%)
 Call exchange_two_integer(p3%, p4%)
End If
End If
'If last_conditions.last_cond(1).two_line_value_no = 0 And n1% = -1000 Then
 'is_two_line_value = False
  'Exit Function
'End If
If p1% = p3% And p2% = p4% Then
   para1 = add_string(para1, para2, True, False)
    p3% = 0
     p4% = 0
    n3% = 0
     n4% = 0
    il2% = 0
      para2 = "0"
   If para1 = "0" Then
    p1% = 0
     p2% = 0
    n1% = 0
     n2% = 0
    il1% = 0
   End If
End If
If para1 = "0" Then
 If para2 <> "0" Or p1% = p2% Then
 t_l_value.value = divide_string(v, para2, True, False)
 is_two_line_value = is_line_value(p3%, p4%, in3%, in4%, il2%, _
    t_l_value.value, no%, n1%, 0, 0, 0, line_value_data0)
       t_l_value.poi(0) = line_value_data0.poi(0)
        t_l_value.poi(1) = line_value_data0.poi(1)
       t_l_value.poi(2) = 0
        t_l_value.poi(3) = 0
       t_l_value.n(0) = line_value_data0.n(0)
        t_l_value.n(1) = line_value_data0.n(1)
       t_l_value.n(2) = 0
        t_l_value.n(3) = 0
       t_l_value.line_no(0) = line_value_data0.line_no
        t_l_value.line_no(1) = 0
       t_l_value.para(0) = "1"
        t_l_value.para(1) = "0"
       t_l_value.value = line_value_data0.value
      cond_ty = line_value_
       Exit Function
 Else
  If v = "0" Then
   no% = 0
    is_two_line_value = 1
     Exit Function
  End If
 End If
ElseIf para2 = "0" Or p3% = p4% Then
If para1 <> "0" Then
 t_l_value.value = divide_string(v, para1, True, False)
 is_two_line_value = is_line_value(p1%, p2%, in1%, in2%, il1%, _
       t_l_value.value, no%, n1%, 0, 0, 0, line_value_data0)
       t_l_value.poi(0) = line_value_data0.poi(0)
        t_l_value.poi(1) = line_value_data0.poi(1)
       t_l_value.poi(2) = 0
        t_l_value.poi(3) = 0
       t_l_value.n(0) = line_value_data0.n(0)
        t_l_value.n(1) = line_value_data0.n(1)
       t_l_value.n(2) = 0
        t_l_value.n(3) = 0
       t_l_value.line_no(0) = line_value_data0.line_no
        t_l_value.line_no(1) = 0
       t_l_value.para(0) = "1"
        t_l_value.para(1) = "0"
       t_l_value.value = line_value_data0.value
       cond_ty = line_value_
        Exit Function
Else
 If v = "0" Then
 no% = 0
 is_two_line_value = 1
  Exit Function
 End If
End If
End If
Call arrange_four_point(p1%, p2%, p3%, p4%, _
         in1%, in2%, in3%, in4%, il1%, il2%, _
          t_l_value.poi(0), t_l_value.poi(1), t_l_value.poi(2), _
           t_l_value.poi(3), 0, 0, t_l_value.n(0), t_l_value.n(1), _
            t_l_value.n(2), t_l_value.n(3), 0, 0, t_l_value.line_no(0), _
             t_l_value.line_no(1), 0, ty, tc_data, is_no_initial)
    If is_no_initial = 1 And n1% = 0 Then
     Call add_record_to_record(tc_data, c_data)
    End If
 t_l_value.value = v
 If ty = 2 Then
    t_l_value.para(0) = add_string(para1, para2, True, False)
    t_l_value.poi(2) = 0
    t_l_value.poi(3) = 0
    t_l_value.line_no(1) = 0
    t_l_value.para(1) = "0"
 ElseIf ty = 1 Or ty = 5 Then
   t_l_value.para(0) = para2
    t_l_value.para(1) = para1
   If ty = 5 And para1 = para2 Then
    t_l_value.poi(1) = t_l_value.poi(3)
    t_l_value.n(1) = t_l_value.n(3)
    t_l_value.poi(2) = 0
    t_l_value.poi(3) = 0
    t_l_value.line_no(1) = 0
    t_l_value.para(1) = "0"
   End If
 ElseIf ty = 3 Or ty = 0 Then
   t_l_value.para(0) = para1
    t_l_value.para(1) = para2
   If ty = 3 And para1 = para2 Then
    t_l_value.poi(1) = t_l_value.poi(3)
    t_l_value.n(1) = t_l_value.n(3)
    t_l_value.poi(2) = 0
    t_l_value.poi(3) = 0
    t_l_value.line_no(1) = 0
    t_l_value.para(1) = "0"
   End If
 ElseIf ty = 4 Then
   t_l_value.para(0) = para1
    t_l_value.para(1) = add_string(para1, para2, True, False)
 ElseIf ty = 6 Then
   t_l_value.para(0) = para2
    t_l_value.para(1) = add_string(para1, para2, True, False)
 ElseIf ty = 7 Then
    t_l_value.para(0) = add_string(para1, para2, True, False)
    t_l_value.para(1) = para2
 ElseIf ty = 8 Then
    t_l_value.para(0) = add_string(para1, para2, True, False)
    t_l_value.para(1) = para1
 End If
If t_l_value.para(1) = "0" Then
   t_l_value.line_no(1) = 0
   t_l_value.poi(2) = 0
   t_l_value.poi(3) = 0
   t_l_value.n(2) = 0
   t_l_value.n(3) = 0
End If
If t_l_value.para(0) = "0" Then
   t_l_value.para(0) = t_l_value.para(1)
   t_l_value.line_no(0) = t_l_value.line_no(1)
   t_l_value.poi(0) = t_l_value.poi(2)
   t_l_value.poi(1) = t_l_value.poi(3)
   t_l_value.n(0) = t_l_value.n(2)
   t_l_value.n(1) = t_l_value.n(3)
   t_l_value.para(1) = "0"
   t_l_value.line_no(1) = 0
   t_l_value.poi(2) = 0
   t_l_value.poi(3) = 0
   t_l_value.n(2) = 0
   t_l_value.n(3) = 0
End If
If n1% = -5000 Then
'  n1% = 0
   GoTo is_two_line_value_mark1
End If
If t_l_value.para(0) = "0" Then
 Exit Function
 '排除相容
Else
ts = ""
Call simple_multi_string0(t_l_value.para(0), t_l_value.para(1), "0", "0", _
          ts, True)

End If
If t_l_value.value = "" Then
 If t_l_value.para(1) = "0" Then
  is_two_line_value = is_line_value(t_l_value.poi(0), t_l_value.poi(1), _
       t_l_value.n(0), t_l_value.n(1), t_l_value.line_no(0), _
     "", no%, -1000, 0, 0, 0, line_value_data0)
   cond_ty = line_value_
    Exit Function
 Else
  GoTo is_two_line_value_mark1
 End If
End If
t_l_value.value = divide_string(t_l_value.value, ts, True, False)
'***********************************
  If t_l_value.para(0) <> "0" And t_l_value.para(1) = "0" Then
    t_l_value.value = divide_string(t_l_value.value, _
         t_l_value.para(0), True, False)
    t_l_value.para(0) = "1"
      is_two_line_value = is_line_value(t_l_value.poi(0), t_l_value.poi(1), _
          t_l_value.n(0), t_l_value.n(1), t_l_value.line_no(0), t_l_value.value, _
           no%, -1000, 0, 0, 0, line_value_data0)
         cond_ty = line_value_
      Exit Function
   ElseIf t_l_value.value = "0" Then
     If t_l_value.para(0) = "1" And t_l_value.para(1) <> "-1" And t_l_value.para(1) <> "@1" Then
      cond_ty = eline_
       is_two_line_value = is_equal_dline(t_l_value.poi(0), _
        t_l_value.poi(1), t_l_value.poi(2), t_l_value.poi(3), _
         t_l_value.n(0), t_l_value.n(1), t_l_value.n(2), _
          t_l_value.n(3), t_l_value.line_no(0), t_l_value.line_no(1), _
            no%, -1000, 0, 0, 0, eline_data0, 0, 0, _
             cond_ty, "", record_0.data0.condition_data)
        Exit Function
     Else
       t_l_value.para(1) = divide_string(t_l_value.para(1), _
        t_l_value.para(0), True, False)
       t_l_value.para(0) = "1"
       is_two_line_value = is_relation(t_l_value.poi(0), _
        t_l_value.poi(1), t_l_value.poi(2), t_l_value.poi(3), _
         t_l_value.n(0), t_l_value.n(1), t_l_value.n(2), _
          t_l_value.n(3), t_l_value.line_no(0), t_l_value.line_no(1), _
            time_string("-1", t_l_value.para(1), True, False), no%, _
             -1000, 0, 0, 0, _
              relation_data0, 0, 0, cond_ty, record_0.data0.condition_data, 0)
     Exit Function
     End If
   End If
'****************************
If is_line_value(t_l_value.poi(0), t_l_value.poi(1), _
      t_l_value.n(0), t_l_value.n(1), t_l_value.line_no(0), "", _
       n%, -1000, 0, 0, 0, line_value_data0) Then
     cond_ty = line_value_
       Call add_conditions_to_record(line_value_, n%, 0, 0, c_data)
   t_l_value.value = divide_string(minus_string(t_l_value.value, _
      time_string(t_l_value.para(0), line_value(n%).data(0).data0.value, False, False), False, False), _
       t_l_value.para(1), True, False)
    is_two_line_value = is_line_value(t_l_value.poi(2), t_l_value.poi(3), _
       t_l_value.n(2), t_l_value.n(3), t_l_value.line_no(1), _
         t_l_value.value, n%, -1000, 0, 0, 0, line_value_data0)
       t_l_value.poi(0) = t_l_value.poi(2)
        t_l_value.poi(1) = t_l_value.poi(3)
       t_l_value.n(0) = t_l_value.n(2)
        t_l_value.n(1) = t_l_value.n(3)
       t_l_value.line_no(0) = t_l_value.line_no(1)
        t_l_value.poi(2) = 0
       t_l_value.poi(3) = 0
        t_l_value.n(2) = 0
         t_l_value.n(3) = 0
        t_l_value.line_no(1) = 0
        t_l_value.para(0) = "1"
        t_l_value.para(1) = "0"
       Exit Function
ElseIf is_line_value(t_l_value.poi(2), t_l_value.poi(3), _
      t_l_value.n(2), t_l_value.n(3), t_l_value.line_no(1), "", _
        n%, -1000, 0, 0, 0, line_value_data0) Then
     cond_ty = line_value_
  Call add_conditions_to_record(line_value_, n%, 0, 0, c_data)
   t_l_value.value = divide_string(minus_string(t_l_value.value, _
        time_string(t_l_value.para(1), line_value(n%).data(0).data0.value, False, False), False, False), _
         t_l_value.para(0), True, False)
    is_two_line_value = is_line_value(t_l_value.poi(0), _
       t_l_value.poi(1), t_l_value.n(0), t_l_value.n(1), _
        t_l_value.line_no(0), t_l_value.value, n%, _
              -1000, 0, 0, 0, line_value_data0)
       t_l_value.poi(2) = 0
       t_l_value.poi(3) = 0
       t_l_value.n(2) = 0
       t_l_value.n(3) = 0
       t_l_value.line_no(1) = 0
       t_l_value.para(0) = "1"
       t_l_value.para(1) = "0"
      Exit Function
End If
is_two_line_value_mark1:
If InStr(1, t_l_value.para(0), "F", 0) > 0 Or _
    InStr(1, t_l_value.para(1), "F", 0) > 0 Or _
     InStr(1, t_l_value.value, "F", 0) > 0 Then
 If no% <= -1000 Then
  is_two_line_value = 0
 Else
  is_two_line_value = 1
   no% = 0
 End If
 Exit Function
End If
cond_ty = two_line_value_
ElseIf p2% = -10 Then
 If Dtwo_point_line(p1%).data(0).value <> "" Then
    Call add_conditions_to_record(V_line_value_, _
       Dtwo_point_line(p1%).data(0).v_line_value_no, 0, 0, c_data)
        t_l_value.poi(0) = p3%
         t_l_value.poi(1) = -10
          t_l_value.para(0) = "1"
           t_l_value.value = time_string(para1, Dtwo_point_line(p1%).data(0).value, False, False)
            t_l_value.value = minus_string(v, t_l_value.value, False, False)
             t_l_value.value = divide_string(t_l_value.value, para2, True, False)
         cond_ty = line_value_
         Exit Function
 ElseIf Dtwo_point_line(p1%).data(0).value <> "" Then
    Call add_conditions_to_record(V_line_value_, _
       Dtwo_point_line(p3%).data(0).v_line_value_no, 0, 0, c_data)
        t_l_value.poi(0) = p1%
         t_l_value.poi(1) = -10
          t_l_value.para(0) = "1"
           t_l_value.value = time_string(para1, Dtwo_point_line(p3%).data(0).value, False, False)
            t_l_value.value = minus_string(v, t_l_value.value, False, False)
             t_l_value.value = divide_string(t_l_value.value, para1, True, False)
         cond_ty = line_value_
         Exit Function
 Else
  If p1% > p3% Then
   Call exchange_two_integer(p1%, p3%)
    Call exchange_string(para1, para2)
  End If
   Call simple_multi_string0(para1, para2, v, "", "", True)
         t_l_value.poi(0) = p1%
         t_l_value.poi(1) = -10
         t_l_value.poi(2) = p3%
         t_l_value.poi(3) = -10
          t_l_value.para(0) = para1
           t_l_value.para(1) = para2
           t_l_value.value = v
         cond_ty = two_line_value_
End If
End If
If n1% = -5000 Then
  If search_for_two_line_value(t_l_value, 0, n1%, 1) Then
    is_two_line_value = 1
 Call search_for_two_line_value(t_l_value, 1, n2%, 1) '5.7
 Call search_for_two_line_value(t_l_value, 2, n3%, 1)
 Call search_for_two_line_value(t_l_value, 3, n4%, 1)
  Exit Function
  Else
    is_two_line_value = 0
     Exit Function
  End If
End If
If search_for_two_line_value(t_l_value, 0, no%, 0) Then
 If minus_string(two_line_value(no%).data(0).data0.para(0), t_l_value.para(0), True, False) <> "0" Or _
      minus_string(two_line_value(no%).data(0).data0.para(1), t_l_value.para(1), True, False) <> "0" Then
      If n1% = -1000 Then
      is_two_line_value = 0
       Exit Function
     Else
     Call add_conditions_to_record(two_line_value_, n%, 0, 0, c_data)
     temp_record.record_data.data0.condition_data = c_data
     If solve_multi_varity_equations(t_l_value.para(0), t_l_value.para(1), _
           "0", "0", t_l_value.value, two_line_value(no%).data(0).data0.para(0), _
            two_line_value(no%).data(0).data0.para(1), "0", "0", two_line_value(no%).data(0).data0.value, _
              t_l_value.para(0), "", "", t_l_value.value) = False Then
                Call set_equation(minus_string("x", t_l_value.value, True, False), 0, temp_record)
                 is_two_line_value = 1
              Exit Function
     End If
        Call set_line_value(t_l_value.poi(2), t_l_value.poi(3), _
              divide_string(t_l_value.value, t_l_value.para(0), True, False), _
                0, 0, 0, temp_record, 0, 0, False)
           no% = 0
            is_two_line_value = 1
             Exit Function
     End If
   Else
      cond_ty = two_line_value_
       is_two_line_value = 1
   If v = "" Or (v <> "" And minus_string(two_line_value(i%).data(0).data0.value, v, True, False) = "0") Then
    If set_or_prove = 2 Then '
     If two_line_value(no%).data(0).record.data1.is_proved = 1 Then
      is_two_line_value = 1
     End If
    Else
     is_two_line_value = 1
    End If
    End If
    Exit Function
    End If
  Else
   If is_line_value(t_l_value.poi(0), t_l_value.poi(1), t_l_value.n(0), t_l_value.n(1), _
              t_l_value.line_no(0), "", tn_%, -1000, 0, 0, 0, l_value(0)) = 1 Then
               Call add_conditions_to_record(line_value_, tn_%, 0, 0, c_data)
        If is_line_value(t_l_value.poi(2), t_l_value.poi(3), t_l_value.n(2), t_l_value.n(3), _
             t_l_value.line_no(1), divide_string(minus_string(t_l_value.value, time_string(t_l_value.para(0), _
                l_value(0).value, False, False), False, False), t_l_value.para(1), True, False), _
                 tn_%, -1000, 0, 0, 0, l_value(1)) = 1 Then
                is_two_line_value = 1
               Call add_conditions_to_record(line_value_, tn_%, 0, 0, c_data)
               If n1% = -5000 Then
                GoTo is_two_line_value_out
               End If
            End If
        cond_ty = line_value_
        t_l_value.poi(0) = l_value(1).poi(0)
        t_l_value.poi(1) = l_value(1).poi(1)
        t_l_value.poi(2) = 0
        t_l_value.poi(3) = 0
        t_l_value.n(0) = l_value(1).n(0)
        t_l_value.n(1) = l_value(1).n(1)
        t_l_value.n(2) = 0
        t_l_value.n(3) = 0
        t_l_value.line_no(0) = l_value(1).line_no
        t_l_value.para(0) = "1"
        t_l_value.para(0) = "0"
        t_l_value.value = l_value(1).value
        Exit Function
   ElseIf is_line_value(t_l_value.poi(2), t_l_value.poi(3), t_l_value.n(2), t_l_value.n(3), _
              t_l_value.line_no(1), "", tn_%, -1000, 0, 0, 0, l_value(0)) = 1 Then
              cond_ty = line_value_
        t_l_value.poi(0) = t_l_value.poi(2)
        t_l_value.poi(1) = t_l_value.poi(3)
        t_l_value.poi(2) = 0
        t_l_value.poi(3) = 0
        t_l_value.n(0) = t_l_value.n(2)
        t_l_value.n(1) = t_l_value.n(3)
        t_l_value.n(2) = 0
        t_l_value.n(3) = 0
        t_l_value.line_no(0) = t_l_value.line_no(1)
        t_l_value.value = divide_string(minus_string(l_value(1).value, time_string( _
           l_value(0).value, t_l_value.para(1), False, True), False, True), _
              t_l_value.para(1), True, False)
        t_l_value.para(0) = "1"
        t_l_value.para(0) = "0"
        Exit Function
   ElseIf t_l_value.value = "0" Then
    If is_relation(t_l_value.poi(0), t_l_value.poi(1), t_l_value.poi(2), t_l_value.poi(3), _
       t_l_value.n(0), t_l_value.n(1), t_l_value.n(2), t_l_value.n(3), t_l_value.line_no(0), _
        t_l_value.line_no(1), divide_string(time_string("-1", t_l_value.para(1), False, False), _
         t_l_value.para(0), True, False), tn_%, -1000, 0, 0, 0, re, 0, 0, cond_ty, c_data, 0) Then
         is_two_line_value = 1
          If n1% = -5000 Then
           GoTo is_two_line_value_out
          End If
         Exit Function
    End If
   'Else
   'depend_no = depend_no + 1
    'If depend_no = 1 Then
     '    is_two_line_value = True
      '    Exit Function
    'End If
  End If
 End If
If n1% = -1000 Then
no% = 0
Else
is_two_line_value_out:
 n1% = no%
 Call search_for_two_line_value(t_l_value, 1, n2%, 1) '5.7
 Call search_for_two_line_value(t_l_value, 2, n3%, 1)
 Call search_for_two_line_value(t_l_value, 3, n4%, 1)
End If
is_two_line_value = False
'End If
'Next i%
  
End Function

Public Function is_tangent_line(ByVal l%, p1%, c1 As condition_type, _
         p2%, c2 As condition_type, tan_l As tangent_line_data_type, _
          n%, n1%, n2%, re As record_data_type) As Byte
Dim i%, j%, k%
Dim tn() As Integer
Dim last_tn As Integer
Dim n_(1) As Integer
Dim tan_L1 As tangent_line_data_type
Dim temp_record As record_data_type
temp_record = re
For i% = 1 To last_conditions.last_cond(1).tangent_line_no
 If tangent_line(i%).data(0).line_no = l% Then
  If c1.no > 0 And c2.no > 0 Then
   If is_same_two_pair_condition(c1, c2, tangent_line(i%).data(0).ele(0), _
          tangent_line(i%).data(0).ele(1)) Then
      n% = i%
       is_tangent_line = 2
        Exit Function
   ElseIf is_same_condition(c1, tangent_line(i%).data(0).ele(0)) Or _
            is_same_condition(c2%, tangent_line(i%).data(0).ele(0)) Then
    If p1% <> p2% Then
     Call line_number0(p1%, p2%, tan_l.n(0), tan_l.n(1))
    Else
     Call is_point_in_line3(p1%, m_lin(l%).data(0).data0, tan_l.n(0))
      tan_l.n(1) = tan_l.n(0)
    End If
     tan_l.ele(0).no = c1.no
     tan_l.ele(1).no = c2.no
     tan_l.ele(0).ty = c1.ty
     tan_l.ele(1).ty = c2.ty
     tan_l.poi(0) = p1%
     tan_l.poi(1) = p2%
    End If
   tangent_line(i%).data(0) = tan_l
    Call add_record_to_record(re.data0.condition_data, tangent_line(i%).data(0).record.data0.condition_data)
     n% = i%
     is_tangent_line = 1
      Exit Function
  ElseIf c1% > 0 Then
   If c1% = tangent_line(i%).data(0).ele(0).no Then
         If tangent_line(i%).data(0).ele(1).no > 0 Then
          n% = i%
           is_tangent_line = 2
            Exit Function
         Else
          n% = i%
           is_tangent_line = 2
            Exit Function
         End If
   ElseIf c1% = tangent_line(i%).data(0).ele(1).no Then
        n% = i%
         is_tangent_line = 2
          Exit Function
   End If
  Else 'c2%>0
   If c2% = tangent_line(i%).data(0).ele(0).no Then
         If tangent_line(i%).data(0).ele(1).no > 0 Then
          n% = i%
           is_tangent_line = 2
            Exit Function
         Else
          n% = i%
           is_tangent_line = 2
            Exit Function
         End If
   ElseIf c2% = tangent_line(i%).data(0).ele(1).no Then
        n% = i%
         is_tangent_line = 2
          Exit Function
   End If
  End If
 End If
Next i%
If p1% > 0 And p2% > 0 Then
If p1% <> p2% Then
Call line_number0(p1%, p2%, tan_l.n(0), tan_l.n(1))
Else
Call is_point_in_line3(p1%, m_lin(l%).data(0).data0, tan_l.n(0))
tan_l.n(1) = tan_l.n(0)
End If
tan_l.ele(0).no = c1%
tan_l.ele(1).no = c2%
tan_l.poi(0) = p1%
tan_l.poi(1) = p2%
If tan_l.n(0) > tan_l.n(1) Then
 Call exchange_two_integer(tan_l.n(0), tan_l.n(1))
 Call exchange_two_integer(tan_l.ele(0).no, tan_l.ele(1).no)
 Call exchange_two_integer(tan_l.poi(0), tan_l.poi(1))
End If
Else
 If p1% > 0 Then
 
  Call is_point_in_line3(p1%, m_lin(l%).data(0).data0, tan_l.n(0))
  tan_l.poi(0) = p1%
  tan_l.ele(0).no = c1%
 ElseIf p2% > 0 Then
  Call is_point_in_line3(p2%, m_lin(l%).data(0).data0, tan_l.n(0))
  tan_l.poi(0) = p2%
  tan_l.ele(0).no = c2%
 End If
End If
tan_l.line_no = l%
End Function
Public Function triangle_number(ByVal p1%, ByVal p2%, _
     ByVal p3%, A1%, A2%, A3%, lv1%, lv2%, lv3%, director As Integer) As Integer
Dim i%, j%, k%, n%
Dim triA As triangle_data0_type
Dim temp_record As total_record_type
If set_triangle_(p1%, p2%, p3%, triA, director) = 0 Then
 triangle_number = 0
  Exit Function
End If
If search_for_triangle(triA, 0, n%, 0) = False Then
 record_0.data0.condition_data.condition_no = 0 'record0
  n% = 0
   Call set_triangle(0, 0, 0, triA, n%, _
    A1%, A2%, A3%, director%, temp_record, 0)
     triangle_number = n%
Else
 Call read_triangle_element(n%, director%, 0, 0, 0, _
            A1%, A2%, A3%, lv1%, lv2%, lv3%, 0, 0, 0)
End If
 triangle_number = n%
End Function
Public Function is_angle_relation(ByVal A1%, ByVal A2%, ByVal v1$, v2$, _
                  no1%, no2%, ov1$, ov2$, oA1%, oA2%, re As condition_data_type) As Boolean
Dim A3_v As angle3_value_data0_type
If A1% = 0 Or A2% = 0 Then
   is_angle_relation = False
    Exit Function
End If
no1% = 0
no2% = 0
is_angle_relation = is_three_angle_value(A1%, A2%, 0, "1", "-1", "0", "0", _
    "0", no1%, no2%, 0, -1000, 0, 0, 0, 0, 0, 0, 0, A3_v, re, 0)
oA1% = A3_v.angle(0)
oA2% = A3_v.angle(1)
ov1$ = A3_v.para(1)
ov2$ = time_string("-1", A3_v.para(0), False, False)
End Function
Public Sub arrange_four_point_for_input_order(ByVal p1%, ByVal p2%, _
    ByVal p3%, ByVal p4%, op1%, op2%, op3%, op4%)
If p1% > p2% Then
   op1% = p2%
   op2% = p1%
Else
   op1% = p1%
   op2% = p2%
End If
If p3% > p4% Then
   op3% = p4%
   op4% = p3%
Else
   op3% = p3%
   op4% = p4%
End If
If op1% > op3% Then
  Call exchange_two_integer(op1%, op3%)
  Call exchange_two_integer(op2%, op4%)
End If
End Sub
Public Function arrange_four_point(ByVal p1%, ByVal p2%, _
    ByVal p3%, ByVal p4%, ByVal n1%, ByVal n2%, ByVal n3%, _
     ByVal n4%, ByVal l1%, ByVal l2%, op1%, op2%, op3%, op4%, _
      op5%, op6%, tn1%, tn2%, tn3%, tn4%, tn5%, tn6%, _
        ol1%, ol2%, ol3%, ty As Byte, cond_data As condition_data_type, _
         is_no_initial As Integer) As Boolean
        '共线条件
Dim n(3) As Integer
Dim tp(3) As Integer
Dim p3_con_l(3) As Integer
Dim c_data As condition_data_type
cond_data.condition_no = 0
If l1% = 0 Or run_type = 10 Then
l1% = line_number0(p1%, p2%, n1%, n2%)
End If
If l1% > 0 Then
If n1% < n2% Then
 tp(0) = p1%
  tp(1) = p2%
 n(0) = n1%
  n(1) = n2%
Else
 tp(0) = p2%
  tp(1) = p1%
 n(0) = n2%
  n(1) = n1%
End If
End If
If l2% = 0 Or run_type = 10 Then
l2% = line_number0(p3%, p4%, n3%, n4%)
End If
If l2% > 0 Then
If n3% < n4% Then
 tp(2) = p3%
  tp(3) = p4%
 n(2) = n3%
  n(3) = n4%
Else
 tp(2) = p4%
  tp(3) = p3%
 n(2) = n4%
  n(3) = n3%
End If
End If
If l1% = 0 And l2% = 0 Then
 arrange_four_point = False
  Exit Function
ElseIf l2% = 0 Then
 arrange_four_point = True
 ty = 0
 op1% = tp(0)
  op2% = tp(1)
   op3% = 0
    op4% = 0
 tn1% = n(0)
  tn2% = n(1)
   tn3% = 0
    tn4% = 0
  ol1% = l1%
   ol2% = 0
ElseIf l1% = 0 Then
 ty = 1
 op1% = tp(2)
  op2% = tp(3)
   op3% = 0
    op4% = 0
 tn1% = n(2)
  tn2% = n(3)
   tn3% = 0
    tn4% = 0
  ol1% = l2%
   ol2% = 0
Else
'*************************************
If l1% < l2% Then
 ty = 0
 op1% = tp(0)
  op2% = tp(1)
   op3% = tp(2)
    op4% = tp(3)
 tn1% = n(0)
  tn2% = n(1)
   tn3% = n(2)
    tn4% = n(3)
  ol1% = l1%
   ol2% = l2%
     arrange_four_point = False
ElseIf l1% > l2% Then
ty = 1
 op1% = tp(2)
  op2% = tp(3)
   op3% = tp(0)
    op4% = tp(1)
tn1% = n(2)
 tn2% = n(3)
  tn3% = n(0)
   tn4% = n(1)
ol1% = l2%
 ol2% = l1%
     arrange_four_point = False
Else 'l1%=l2%
 ol1% = l1%
  ol2% = l2%
   Call is_three_point_on_line(p1%, p2%, p3%, p3_con_l(0), -1000, 0, 0, condition_data0, 0, 0, 0)
   Call is_three_point_on_line(p1%, p2%, p4%, p3_con_l(1), -1000, 0, 0, condition_data0, 0, 0, 0)
   Call is_three_point_on_line(p1%, p4%, p3%, p3_con_l(2), -1000, 0, 0, condition_data0, 0, 0, 0)
   Call is_three_point_on_line(p2%, p4%, p3%, p3_con_l(3), -1000, 0, 0, condition_data0, 0, 0, 0)
   If three_point_on_line(p3_con_l(0)).data(0).is_no_initial = 1 And p3_con_l(0) > 0 Then
      cond_data = three_point_on_line(p3_con_l(0)).data(0).record.data0.condition_data
          is_no_initial = 1
   ElseIf three_point_on_line(p3_con_l(1)).data(0).is_no_initial = 1 And p3_con_l(1) > 0 Then
      cond_data = three_point_on_line(p3_con_l(1)).data(0).record.data0.condition_data
          is_no_initial = 1
   ElseIf three_point_on_line(p3_con_l(2)).data(0).is_no_initial = 1 And p3_con_l(2) > 0 Then
      cond_data = three_point_on_line(p3_con_l(2)).data(0).record.data0.condition_data
          is_no_initial = 1
   ElseIf three_point_on_line(p3_con_l(3)).data(0).is_no_initial = 1 And p3_con_l(3) > 0 Then
      cond_data = three_point_on_line(p3_con_l(3)).data(0).record.data0.condition_data
          is_no_initial = 1
   End If
 If n(0) = n(2) And n(1) = n(3) Then
     ty = 2
   op1% = tp(0)
    op2% = tp(1)
     op3% = tp(2)
      op4% = tp(3)
   tn1% = n(0)
    tn2% = n(1)
     tn3% = n(2)
      tn4% = n(3)
     arrange_four_point = False
  ElseIf n(1) = n(2) Then
     ty = 3
  op1% = tp(0)
    op2% = tp(1)
     op3% = tp(2)
      op4% = tp(3)
       op5% = op1%
        op6% = op4%
   tn1% = n(0)
    tn2% = n(1)
     tn3% = n(2)
      tn4% = n(3)
       tn5% = tn1%
        tn6% = tn4%
   ol3% = ol1%
     arrange_four_point = True
 ElseIf n(1) = n(3) And n(0) < n(2) Then
     ty = 4
    op1% = tp(0)
     op2% = tp(2)
      op3% = tp(2)
       op4% = tp(1)
    tn1% = n(0)
     tn2% = n(2)
      tn3% = n(2)
       tn4% = n(1)
      arrange_four_point = True
ElseIf n(3) = n(0) Then
   ty = 5
    op1% = tp(2)
     op2% = tp(3)
      op3% = tp(0)
       op4% = tp(1)
        op5% = op1%
         op6% = op4%
   tn1% = n(2)
    tn2% = n(3)
     tn3% = n(0)
      tn4% = n(1)
       tn5% = tn1%
        tn6% = tn4%
         ol3% = ol1%
      arrange_four_point = True
ElseIf n(3) = n(1) And n(0) > n(2) Then
    ty = 6
     op1% = tp(2)
      op2% = tp(0)
       op3% = tp(0)
        op4% = tp(1)
      tn1% = n(2)
       tn2% = n(0)
        tn3% = n(0)
         tn4% = n(1)
     arrange_four_point = True
ElseIf n(0) = n(2) And n(1) < n(3) Then
    ty = 7
     op1% = tp(0)
      op2% = tp(1)
       op3% = tp(1)
        op4% = tp(3)
      tn1% = n(0)
       tn2% = n(1)
        tn3% = n(1)
         tn4% = n(3)
     arrange_four_point = True
ElseIf n(0) = n(2) And n(1) > n(3) Then
   ty = 8
     op1% = tp(0)
      op2% = tp(3)
       op3% = tp(3)
        op4% = tp(1)
      tn1% = n(0)
       tn2% = n(3)
        tn3% = n(3)
         tn4% = n(1)
     arrange_four_point = True
Else
    If n(0) > n(2) Or (n(0) = n(2) And _
        n(1) < n(3)) Then
     Call exchange_two_integer(n(0), n(2))
     Call exchange_two_integer(n(1), n(3))
     Call exchange_two_integer(tp(0), tp(2))
     Call exchange_two_integer(tp(1), tp(3))
     ty = 1 '反序
    Else
     ty = 0
    End If
     op1% = tp(0)
      op2% = tp(1)
       op3% = tp(2)
        op4% = tp(3)
      tn1% = n(0)
       tn2% = n(1)
        tn3% = n(2)
         tn4% = n(3)
     arrange_four_point = False
End If
End If
End If
'ol1% = l1%
 'ol2% = l2%
End Function


Public Sub ratio_value(v1$, ty As Byte, v2$)
'以序后表序前
If v1$ = "" Then
 v2$ = ""
 Exit Sub
End If
Select Case ty
Case 0, 3
 v2$ = v1$
Case 1, 5
 v2$ = divide_string("1", v1$, True, False)
Case 2
 v1$ = "1"
 v2$ = "1"
Case 4
v2$ = add_string(v1$, "1", True, False)
Case 6
v2$ = divide_string("1", add_string(v1$, "1", False, False), True, False)
Case 7
v2$ = divide_string("1", add_string("1", divide_string("1", v1$, False, False), _
         False, False), True, False)
Case 8
v2$ = add_string("1", divide_string("1", v1$, False, False), True, False)
End Select

End Sub


Public Sub ratio_value1(v1$, ty As Byte, v2$)
'排序前表排序后
If v1$ = "" Then
 v2$ = ""
  Exit Sub
End If
Select Case ty
Case 0, 3
 v2$ = v1$
Case 1, 5
 v2$ = divide_string("1", v1$, True, False)
Case 2
 v1$ = "1"
 v2$ = "1"
Case 4
v2$ = minus_string(v1$, "1", True, False)
Case 6
v2$ = minus_string(divide_string("1", v1$, False, False), "1", True, False)
Case 7
v2$ = divide_string("1", minus_string(divide_string("1", v1$, False, False), _
          "1", False, False), True, False)
Case 8
v2$ = divide_string("1", minus_string(v1$, "1", False, False), True, False)
End Select
End Sub

Public Function is_mid_point_line(ByVal p1%, ByVal p2%, _
        ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, op1%, _
         op2%, op3%, op4%, op5%, op6%, no%)
Dim tp(3) As Integer
Dim i%
If p1% < p2% Then
 tp(0) = p1%
  tp(1) = p2%
Else
 tp(0) = p2%
  tp(1) = p1%
End If
If p3% < p4% Then
 tp(2) = p3%
  tp(3) = p4%
Else
 tp(2) = p4%
  tp(0) = p3%
End If
If p5% < p6% Then
 op5% = p5%
  op6% = p6%
Else
 op5% = p6%
  op6% = p5%
End If
 If tp(0) < tp(2) Or (tp(0) = tp(2) And tp(1) <= tp(3)) Then
  op1% = tp(0)
   op2% = tp(1)
    op3% = tp(2)
     op4% = tp(3)
 Else
  op1% = tp(1)
   op2% = tp(0)
    op3% = tp(3)
     op4% = tp(2)
End If
For i% = 1 To last_conditions.last_cond(1).mid_point_line_no
 If mid_point_line(i%).data(0).poi(0) = op1% And _
  mid_point_line(i%).data(0).poi(1) = op2% And _
   mid_point_line(i%).data(0).poi(2) = op3% And _
    mid_point_line(i%).data(0).poi(3) = op4% And _
     mid_point_line(i%).data(0).poi(4) = op5% And _
       mid_point_line(i%).data(0).poi(5) = op6% Then
   no% = i%
    If set_or_prove = 2 And _
     mid_point_line(i%).data(0).record.data1.is_proved = 2 Then '
      is_mid_point_line = True
    ElseIf set_or_prove < 2 Then '
     is_mid_point_line = False
    End If
          Exit Function

  End If
Next i%
 
End Function

Public Sub arrange_four_index_(ByVal p1%, ByVal p2%, _
    ByVal p3%, ByVal p4%, ByVal ty As Byte, _
     op1%, op2%, op3%, op4%)
If (ty = 3 Or ty = 5) And p2% = p3% Then
 op1% = p1%
  op2% = p2%
   op3% = p3%
    op4% = p4%
ElseIf (ty = 4 Or ty = 8) And p2% = p3% Then
 op1% = p1%
  op2% = p4%
   op3% = p2%
    op4% = p4%
ElseIf (ty = 7 Or ty = 6) And p2% = p3% Then
 op1% = p1%
 op2% = p2%
 op3% = p1%
 op4% = p4%
ElseIf (ty = 5 Or ty = 3) Then
 op1% = p2%
 op2% = p4%
 op3% = p1%
 op4% = p2%
ElseIf (ty = 6 Or ty = 7) Then
 op1% = p2%
 op2% = p4%
 op3% = p1%
 op4% = p4%
ElseIf ty = 8 Or ty = 4 Then
 op1% = p1%
 op2% = p4%
 op3% = p1%
 op4% = p2%
End If

End Sub

Public Function is_area_relation(triA1 As condition_type, triA2 As condition_type, _
                     ByVal value As String, n%, n1%, n2%, n3%, otriA1 As condition_type, _
                      otriA2 As condition_type, otriA3 As condition_type, _
                       ovalue As String, ty As Byte, tn1%, tn2%) As Boolean
Dim i%
Dim ty1 As Byte
Dim t_A(3) As condition_type
Dim triA_r As area_relation_data_type
n% = 0
ty = 0
If triA1.no = 0 Or triA2.no = 0 Then
 is_area_relation = False
  Exit Function
ElseIf triA1.no = triA2.no And triA1.ty = triA2.ty Then
 is_area_relation = True
  ty = 0
   otriA1 = triA1
    otriA2 = triA2
     otriA3.no = 0
      ovalue = "1"
   Exit Function
End If
'value = "x"
If triA1.ty > triA2.ty Then
   ty1 = triA1.ty
   triA1.ty = triA2.ty
   triA2.ty = ty1
   Call exchange_two_integer(triA1.no, triA2.no)
    value = divide_string("1", value, True, False)
ElseIf triA1.no > triA2.no Then
   Call exchange_two_integer(triA1.no, triA2.no)
    value = divide_string("1", value, True, False)
End If
ty1 = combine_two_area_elemenet(triA1, triA2, t_A(0), t_A(1), t_A(2))
If t_A(0).no = 0 Or t_A(1).no = 0 Then
   If n1% = -1000 Then
      is_area_relation = False
   Else
      is_area_relation = True
       n% = 0
   End If
     Exit Function
End If
If value <> "" Then
Call ratio_value1(value, ty1, ovalue$)
End If
If (t_A(0).ty > t_A(1).ty) Or (t_A(0).ty = t_A(1).ty Or t_A(0).no > t_A(1).no) Then
 t_A(3) = t_A(0)
  t_A(0) = t_A(1)
   t_A(1) = t_A(3)
  If ovalue <> "" And ovalue <> "0" Then
   ovalue = divide_string("1", ovalue, True, False)
  End If
If ty1 = 0 Then
 ty1 = 1
ElseIf ty1 = 3 Then
 ty1 = 5
ElseIf ty1 = 5 Then
 ty1 = 3
ElseIf ty1 = 4 Then
 ty1 = 8
ElseIf ty1 = 8 Then
 ty1 = 4
ElseIf ty1 = 6 Then
 ty1 = 7
ElseIf ty1 = 7 Then
 ty1 = 6
End If
End If  'ty1 = 1
triA_r.area_element(0) = t_A(0)
 triA_r.area_element(1) = t_A(1)
  triA_r.area_element(2) = t_A(2)
If InStr(1, triA_r.value, "F", 0) > 0 Then
 If n1% < -1000 Then
  is_area_relation = False
 Else
  is_area_relation = True
   n% = 0
 End If
 Exit Function
End If
If search_for_area_relation(triA_r, 0, n%, 0) Then
   '         n% = i%
             otriA1 = t_A(0)
              otriA2 = t_A(1)
               otriA3 = t_A(2)
 '               call solve_equation_(
             ty = area_relation_
              tn1% = 0
               tn2% = 0
              ovalue = Darea_relation(n%).data(0).value
               If ty1 = 1 Then
                value = divide_string("1", ovalue$, True, False)
               ElseIf ty1 > 2 Then
                Call ratio_value(ovalue$, ty1, value)
               Else
                value = ovalue
               End If
    If set_or_prove = 2 And _
      Darea_relation(n%).data(0).record.data1.is_proved = 2 Then '
     is_area_relation = True
    Else
     is_area_relation = True
    End If
     Exit Function
End If
If is_area_of_element(triA1.ty, triA1.no, tn1%, -1000) And _
     is_area_of_element(triA2.ty, triA2.no, tn2%, -1000) Then
   ty = area_of_element_
    n% = 0
     ovalue = divide_string(area_of_element(tn1%).data(0).value, _
       area_of_element(tn2%).data(0).value, True, False)
  If ty1 = 0 Then
    value = ovalue
    Else
    value = divide_string("1", ovalue, True, False)
    End If
    is_area_relation = True
    Exit Function
End If
If n1% <> -1000 Then
 n1% = n%
 Call search_for_area_relation(triA_r, 1, n2%, 1)
 Call search_for_area_relation(triA_r, 2, n3%, 1)
Else
 n% = 0
  Exit Function
End If
 otriA1 = t_A(0)
  otriA2 = t_A(1)
   otriA3 = t_A(2)
   'ovalue = value
End Function

Public Function is_general_string(ByVal i1%, ByVal i2%, ByVal i3%, _
          ByVal I4%, ByVal pA1$, ByVal pA2$, ByVal pa3$, ByVal pa4$, _
           ByVal v$, no%, no1%, no2%, no3%, no4%, g_s As general_string_data_type, _
            concl_no As Byte, cond_ty As Byte, re As record_data_type, _
              ByVal no_reduce As Byte) As Byte
Dim i%, j%, t_n%, tn_%
'Dim rA1 As String
Dim rA2 As String
Dim l(3) As Integer
Dim tn(7) As Integer
Dim insert_no%
Dim ty As Byte
Dim t_g_s As general_string_data_type
Dim temp_record As total_record_type
temp_record.record_data = re
g_s.item(0) = i1%
 g_s.item(1) = i2%
  g_s.item(2) = i3%
   g_s.item(3) = I4%
g_s.para(0) = pA1$
 g_s.para(1) = pA2$
  g_s.para(2) = pa3$
   g_s.para(3) = pa4$
    g_s.value = v$
If i1% > 0 Then
 If item0(i1%).data(0).value <> "" Then
  Call add_record_to_record(item0(i1%).data(0).record_for_value.data0.condition_data, _
                                            temp_record.record_data.data0.condition_data)
  Call add_record_to_record(item0(i1%).data(0).record_for_value.data0.condition_data, _
                                            re.data0.condition_data)
   g_s.para(0) = time_string(g_s.para(0), item0(i1%).data(0).value, True, False)
    g_s.item(0) = 0
     If g_s.value <> "" Then
      g_s.value = minus_string(g_s.value, g_s.para(0), True, False)
       g_s.para(0) = "0"
     End If
 End If
End If
If i2% > 0 Then
 If item0(i2%).data(0).value <> "" Then
  Call add_record_to_record(item0(i2%).data(0).record_for_value.data0.condition_data, _
                  temp_record.record_data.data0.condition_data)
  Call add_record_to_record(item0(i2%).data(0).record_for_value.data0.condition_data, _
                                           re.data0.condition_data)
   g_s.para(1) = time_string(g_s.para(1), item0(i2%).data(0).value, True, False)
    g_s.item(1) = 0
     If g_s.value <> "" Then
      g_s.value = minus_string(g_s.value, g_s.para(1), True, False)
       g_s.para(1) = "0"
     End If
 End If
End If
If i3% > 0 Then
 If item0(i3%).data(0).value <> "" Then
  Call add_record_to_record(item0(i3%).data(0).record_for_value.data0.condition_data, _
                                              temp_record.record_data.data0.condition_data)
  Call add_record_to_record(item0(i3%).data(0).record_for_value.data0.condition_data, _
                                                              re.data0.condition_data)
   g_s.para(2) = time_string(g_s.para(2), item0(i3%).data(0).value, True, False)
    g_s.item(2) = 0
     If g_s.value <> "" Then
      g_s.value = minus_string(g_s.value, g_s.para(2), True, False)
       g_s.para(2) = "0"
     End If
 End If
End If
If I4% > 0 Then
 If item0(I4%).data(0).value <> "" Then
  Call add_record_to_record(item0(I4%).data(0).record_for_value.data0.condition_data, _
                                                    temp_record.record_data.data0.condition_data)
  Call add_record_to_record(item0(I4%).data(0).record_for_value.data0.condition_data, _
                                                     re.data0.condition_data)
   g_s.para(3) = time_string(g_s.para(3), item0(I4%).data(0).value, True, False)
    g_s.item(3) = 0
     If g_s.value <> "" Then
      g_s.value = minus_string(g_s.value, g_s.para(3), True, False)
       g_s.para(3) = "0"
     End If
 End If
End If

     'g_s.record_.conclusion_no = concl_no
If g_s.value <> "" Then
For i% = 3 To 0 Step -1
If g_s.item(i%) = 0 And g_s.para(i%) <> "" And g_s.para(i%) <> "0" Then
 g_s.value = minus_string(g_s.value, g_s.para(i%), True, False)
  g_s.para(i%) = "0"
End If
Next i%
ElseIf g_s.value = "0" Then
If (g_s.para(0) = "0" Or g_s.para(0) = "") And (g_s.para(1) = "0" Or g_s.para(1) = "") And _
     (g_s.para(2) = "0" Or g_s.para(2) = "") And (g_s.para(3) = "0" Or g_s.para(3) = "") Then
       is_general_string = 1
        no% = 0
         re = temp_record.record_data
         Exit Function
End If
End If
If ((g_s.item(0) > g_s.item(1) Or g_s.item(0) = 0) And g_s.item(1) > 0) Or _
     (g_s.item(0) = g_s.item(1) And g_s.para(0) = "0" And g_s.para(1) <> "0") Then
 Call exchange_two_integer(g_s.item(0), g_s.item(1))
  Call exchange_two_string(g_s.para(0), g_s.para(1))
End If
If ((g_s.item(1) > g_s.item(2) Or g_s.item(1) = 0) And g_s.item(2) > 0) Or _
     (g_s.item(1) = g_s.item(2) And g_s.para(1) = "0" And g_s.para(2) <> "0") Then
 Call exchange_two_integer(g_s.item(1), g_s.item(2))
  Call exchange_two_string(g_s.para(1), g_s.para(2))
End If
If ((g_s.item(2) > g_s.item(3) Or g_s.item(2) = 0) And g_s.item(3) > 0) Or _
     (g_s.item(2) = g_s.item(3) And g_s.para(2) = "0" And g_s.para(3) <> "0") Then
 Call exchange_two_integer(g_s.item(2), g_s.item(3))
  Call exchange_two_string(g_s.para(2), g_s.para(3))
End If
If ((g_s.item(0) > g_s.item(1) Or g_s.item(0) = 0) And g_s.item(1) > 0) Or _
     (g_s.item(0) = g_s.item(1) And g_s.para(0) = "0" And g_s.para(1) <> "0") Then
 Call exchange_two_integer(g_s.item(0), g_s.item(1))
  Call exchange_two_string(g_s.para(0), g_s.para(1))
End If
If ((g_s.item(1) > g_s.item(2) Or g_s.item(1) = 0) And g_s.item(2) > 0) Or _
     (g_s.item(1) = g_s.item(2) And g_s.para(1) = "0" And g_s.para(2) <> "0") Then
 Call exchange_two_integer(g_s.item(1), g_s.item(2))
  Call exchange_two_string(g_s.para(1), g_s.para(2))
End If
If ((g_s.item(0) > g_s.item(1) Or g_s.item(0) = 0) And g_s.item(1) > 0) Or _
   (g_s.item(0) = g_s.item(1) And g_s.para(0) = "0" And g_s.para(1) <> "0") Then
 Call exchange_two_integer(g_s.item(0), g_s.item(1))
  Call exchange_two_string(g_s.para(0), g_s.para(1))
End If
Do While g_s.item(0) = g_s.item(1) And g_s.para(1) <> "0"
 g_s.para(0) = add_string(g_s.para(0), g_s.para(1), True, False)
  g_s.para(1) = g_s.para(2)
   g_s.para(2) = g_s.para(3)
    g_s.para(3) = "0"
  g_s.item(1) = g_s.item(2)
   g_s.item(2) = g_s.item(3)
    g_s.item(3) = 0
     If g_s.para(0) = "0" Then
      g_s.item(0) = 0
     End If
Loop
Do While g_s.item(1) = g_s.item(2) And g_s.para(2) <> "0"
 g_s.para(1) = add_string(g_s.para(1), g_s.para(2), True, False)
   g_s.para(2) = g_s.para(3)
    g_s.para(3) = "0"
      g_s.item(2) = g_s.item(3)
    g_s.item(3) = 0
     If g_s.para(1) = "0" Then
      g_s.item(1) = 0
     End If
Loop
If g_s.item(2) = g_s.item(3) And g_s.para(3) <> "0" Then
 g_s.para(2) = add_string(g_s.para(2), g_s.para(3), True, False)
    g_s.para(3) = "0"
    g_s.item(3) = 0
     If g_s.para(2) = "0" Then
      g_s.item(2) = 0
     End If
Else
 GoTo is_general_string_mark2
End If
If g_s.item(0) = g_s.item(1) And g_s.item(0) > 0 Then
 g_s.para(0) = add_string(g_s.para(0), g_s.para(1), True, False)
  g_s.para(1) = g_s.para(2)
   g_s.para(2) = "0"
  g_s.item(1) = g_s.item(2)
   g_s.item(2) = 0
    If g_s.para(0) = "0" Then
     g_s.item(0) = 0
    End If
ElseIf g_s.item(1) = g_s.item(2) And g_s.item(1) > 0 Then
 g_s.para(1) = add_string(g_s.para(1), g_s.para(2), True, False)
   g_s.para(2) = "0"
      g_s.item(2) = 0
       If g_s.para(1) = "0" Then
        g_s.item(1) = 0
       End If
Else
GoTo is_general_string_mark2
End If
If g_s.item(0) = g_s.item(1) And g_s.item(0) > 0 Then
 g_s.para(0) = add_string(g_s.para(0), g_s.para(1), True, False)
  g_s.para(1) = "0"
     g_s.item(1) = 0
      If g_s.para(0) = "0" Then
       g_s.item(0) = 0
      End If
End If
is_general_string_mark2:
If g_s.para(0) = "0" And v$ = "0" Then
 is_general_string = 1
  Call set_level(temp_record.record_data.data0.condition_data)
   no% = 0
    re = temp_record.record_data
     Exit Function
End If
If v$ <> "" Then
If g_s.item(0) = 0 And g_s.para(0) <> "0" Then
 v$ = minus_string(v$, g_s.para(0), True, False)
  g_s.para(0) = "0"
End If
If g_s.item(1) = 0 And g_s.para(1) <> "0" Then
 v$ = minus_string(v$, g_s.para(1), True, False)
  g_s.para(1) = "0"
End If
If g_s.item(2) = 0 And g_s.para(2) <> "0" Then
 v$ = minus_string(v$, g_s.para(2), True, False)
  g_s.para(2) = "0"
End If
If g_s.item(3) = 0 And g_s.para(3) <> "0" Then
 v$ = minus_string(v$, g_s.para(3), True, False)
  g_s.para(3) = "0"
End If
End If
If g_s.para(0) <> "0" Then 'And g_s.value <> "" Then
rA2 = ""
Call simple_multi_string0(g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), _
        rA2, True)
        g_s.trans_para = rA2
Else
g_s.trans_para = "1"
  ' g_s.trans_para(1) = "1"
End If
If g_s.value <> "" And g_s.value <> "0" And rA2 <> "" Then
 g_s.value = divide_string(g_s.value, rA2, True, False)
End If
    If g_s.value = "0" And g_s.para(0) = "0" Then
     is_general_string = 1
      Call set_level(temp_record.record_data.data0.condition_data)
      no% = 0
                re = temp_record.record_data
       Exit Function
     End If
If g_s.item(0) = 0 Then
  If g_s.value <> "0" And g_s.value <> "" Then
   If no1% = -1000 Then
    is_general_string = 0
   ElseIf no1% = 0 Then
    is_general_string = 1
     no% = 0
   End If
  End If
   Call simple_equation(g_s.value, temp_record, g_s.value)
   'g_s.value_ = simple_equation(g_s.value_)
End If
'*************************************************************************************
If InStr(1, g_s.para(0), "F", 0) > 0 Or InStr(1, g_s.para(1), "F", 0) > 0 Or _
     InStr(1, g_s.para(2), "F", 0) > 0 Or InStr(1, g_s.para(3), "F", 0) > 0 Or _
      InStr(1, g_s.value, "F", 0) > 0 Then
        If no1% = -1000 Then
         is_general_string = 0
        Else
         is_general_string = 1
        End If
          re = temp_record.record_data
       Exit Function
'ElseIf simple_general_string0(g_s) Then
 '  is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
       g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, _
         no4%, g_s, concl_no, cond_ty, temp_record.record_data, no_reduce)
  '        re = temp_record.record_data
  '         Exit Function
End If
If search_for_general_string(g_s, 0, no%, 0) Then
 If no1 = -2000 Then
   Call search_for_general_string(g_s, 0, no1%, 1)
    GoTo is_general_string_mark6
 End If
 If g_s.value <> "" Then
  If g_s.para(0) <> general_string(no%).data(0).para(0) Or _
      g_s.para(1) <> general_string(no%).data(0).para(1) Or _
       g_s.para(2) <> general_string(no%).data(0).para(2) Or _
        g_s.para(3) <> general_string(no%).data(0).para(3) Or _
         g_s.value <> general_string(no%).data(0).value Then
          re = temp_record.record_data
    If no1% = -1000 Then
     is_general_string = 0
      Exit Function
    End If
   'Else
    Call add_conditions_to_record(general_string_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    t_g_s = g_s
   ty = solve_multi_varity_equations(t_g_s.para(0), t_g_s.para(1), _
     t_g_s.para(2), t_g_s.para(3), t_g_s.value, _
      general_string(no%).data(0).para(0), general_string(no%).data(0).para(1), _
       general_string(no%).data(0).para(2), general_string(no%).data(0).para(3), _
        general_string(no%).data(0).value, t_g_s.para(1), t_g_s.para(2), _
         t_g_s.para(3), t_g_s.value)
    If ty = 1 Then
    If t_g_s.para(0) = "0" Then
     t_g_s.item(0) = 0
    End If
    If t_g_s.para(1) = "0" Then
     t_g_s.item(1) = 0
    End If
    If t_g_s.para(2) = "0" Then
     t_g_s.item(2) = 0
    End If
     If t_g_s.item(1) = 0 And t_g_s.item(2) = 0 And t_g_s.item(3) = 0 Then
        If t_g_s.value = "0" Then
         is_general_string = 1
        Else
         is_general_string = 0
        End If
          re = temp_record.record_data
       Exit Function
     Else
      no% = 0
      is_general_string = is_general_string(t_g_s.item(1), t_g_s.item(2), t_g_s.item(3), _
        0, t_g_s.para(1), t_g_s.para(2), t_g_s.para(3), "0", t_g_s.value, _
            no%, no1%, no2%, no3%, no4%, g_s, concl_no, cond_ty, temp_record.record_data, no_reduce)
            Call set_level(temp_record.record_data.data0.condition_data)
          re = temp_record.record_data
             Exit Function
     End If
    ElseIf ty = 2 Then
     error_of_wenti = 3
      If no1% = -1000 Then
       is_general_string = 0
      Else
       no% = 0
        is_general_string = 2
      End If
         re = temp_record.record_data
      Exit Function
    End If
   End If
  End If
  cond_ty = general_string_
  is_general_string = 1
   Call set_level(temp_record.record_data.data0.condition_data)
         re = temp_record.record_data
   Exit Function
Else
insert_no% = no%
no% = 0
End If
If g_s.item(3) > 0 Then
 t_g_s = g_s
  If combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(1), _
           t_g_s.item(2), 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
       g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(1), _
           t_g_s.item(3), 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(2), _
           t_g_s.item(3), 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(1), t_g_s.item(2), _
           t_g_s.item(3), 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(1), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
           re = temp_record.record_data
        Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(2), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(3), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(1), t_g_s.item(2), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
         re = temp_record.record_data
          Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(1), t_g_s.item(3), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
         re = temp_record.record_data
          Exit Function
   ElseIf combine_general_string_with_general_string0(t_g_s.item(2), t_g_s.item(3), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
         re = temp_record.record_data
          Exit Function
   End If
 ElseIf g_s.item(2) > 0 Then
   t_g_s = g_s
   If combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(1), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
         re = temp_record.record_data
          Exit Function
    ElseIf combine_general_string_with_general_string0(t_g_s.item(0), t_g_s.item(2), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
         re = temp_record.record_data
          Exit Function
    ElseIf combine_general_string_with_general_string0(t_g_s.item(1), t_g_s.item(2), _
          0, 0, t_g_s, _
                    re.data0.condition_data, g_s) Then
     is_general_string = is_general_string(g_s.item(0), g_s.item(1), g_s.item(2), g_s.item(3), _
        g_s.para(0), g_s.para(1), g_s.para(2), g_s.para(3), g_s.value, no%, no1%, no2%, no3%, no4%, g_s, _
         concl_no, cond_ty, re, 0)
          re = temp_record.record_data
         Exit Function
    End If
End If
  
'Next i%
'

If g_s.value <> "" Then
'If g_s.para(3) = "0" And g_s.para(3) <> "0" Then
 'If item0(g_s.item(0)).data(0).sig = "~" And item0(g_s.item(1)).data(0).sig = "~" And _
    item0(g_s.item(2)).data(0).sig = "~" And item0(g_s.item(0)).data(0).poi(1) > 0 And _
     item0(g_s.item(1)).data(0).poi(1) > 0 And item0(g_s.item(2)).data(0).poi(1) > 0 Then
  'Call initial_record(record_0)
  't_g_s = g_s
  'is_general_string = set_three_line_value(item0(t_g_s.item(0)).data(0).poi(0), item0(t_g_s.item(0)).data(0).poi(1), _
   item0(t_g_s.item(1)).data(0).poi(0), item0(t_g_s.item(1)).data(0).poi(1), item0(t_g_s.item(2)).data(0).poi(0), _
    item0(t_g_s.item(2)).data(0).poi(1), item0(t_g_s.item(0)).data(0).n(0), item0(t_g_s.item(0)).data(0).n(1), _
     item0(t_g_s.item(1)).data(0).n(0), item0(t_g_s.item(1)).data(0).n(1), item0(t_g_s.item(2)).data(0).n(0), _
      item0(t_g_s.item(2)).data(0).n(1), item0(t_g_s.item(0)).data(0).line_no(0), item0(t_g_s.item(1)).data(0).line_no(0), _
       item0(t_g_s.item(2)).data(0).line_no(0), t_g_s.para(0), t_g_s.para(1), t_g_s.para(2), t_g_s.value, temp_record, 0, 0, 0)
   '   If t_n% > 0 Then
   '    Call set_level_(general_string(t_n%).record_.no_reduce, 4)
   '   End If
   '   If is_general_string = 0 Then
   '    is_general_string = 1
   '   End If
   '   Call set_level(temp_record.record_data)
   '   no% = 0
   '    Exit Function
  'End If
ElseIf g_s.para(3) = "0" And g_s.para(2) = "0" And g_s.para(1) <> "0" Then
 'If item0(g_s.item(0)).data(0).sig = "~" And item0(g_s.item(1)).data(0).sig = "~" And _
    item0(g_s.item(0)).data(0).poi(1) > 0 And item0(g_s.item(1)).data(0).poi(1) > 0 Then
 't_g_s = g_s
 ' is_general_string = set_two_line_value(item0(t_g_s.item(0)).data(0).poi(0), item0(t_g_s.item(0)).data(0).poi(1), _
    item0(t_g_s.item(1)).data(0).poi(0), item0(t_g_s.item(1)).data(0).poi(1), item0(t_g_s.item(0)).data(0).n(0), _
     item0(t_g_s.item(0)).data(0).n(1), item0(t_g_s.item(1)).data(0).n(0), item0(t_g_s.item(1)).data(0).n(1), _
      item0(t_g_s.item(0)).data(0).line_no(0), item0(t_g_s.item(1)).data(0).line_no(0), t_g_s.para(0), _
       t_g_s.para(1), t_g_s.value, temp_record, 0, no_reduce)
 '       If is_general_string = 0 Then
 '         is_general_string = 1
 '       End If
 '  Call set_level(temp_record.record_data)
 '     If t_n% > 0 Then
 '      Call set_level_(general_string(t_n%).record_.no_reduce, 4)
 '     End If
 '  no% = 0
 '   Exit Function
 ' End If
 'End If
End If
 If t_n% > 0 Then
  is_general_string = 1
   no% = t_n%
 Else
  If no1% = -1000 Then
   no% = 0
    Exit Function
  End If
  no1% = insert_no%
is_general_string_mark6:
  re = temp_record.record_data
  Call search_for_general_string(g_s, 1, no2%, 1)
  Call search_for_general_string(g_s, 2, no3%, 1)
  Call search_for_general_string(g_s, 3, no4%, 1)
  End If
  is_general_string = 0
  cond_ty = general_string_
End Function
Public Function is_same_item0(i1 As item0_data_type, i2 As item0_data_type) As Boolean
If i1.poi(0) = i2.poi(0) And i1.poi(1) = i2.poi(1) And _
    i1.poi(2) = i2.poi(2) And i1.poi(3) = i2.poi(3) And _
      i1.sig = i2.sig Then
        is_same_item0 = True
End If
End Function


Public Function is_epolygon(pol As polygon, no%, EPol As epolygon_data_type) As Boolean
Dim i%, j%
Dim t_pol As polygon
Dim oPol As polygon
j% = 0
t_pol = pol
For i% = 1 To t_pol.total_v - 1
 If t_pol.v(j%) > t_pol.v(i%) Then
  j% = i%
 End If
Next i%
For i% = 0 To t_pol.total_v - 1
 oPol.v(i%) = t_pol.v((j% + i%) Mod t_pol.total_v)
Next i%
If oPol.v(1) > oPol.v(t_pol.total_v - 1) Then
For i% = 1 To t_pol.total_v - 1
 oPol.v(t_pol.total_v - i%) = t_pol.v((j% + i%) Mod t_pol.total_v)
Next i%
End If
oPol.total_v = t_pol.total_v
EPol.p = oPol
If t_pol.total_v = 3 Then
EPol.no = triangle_number(t_pol.v(0), t_pol.v(1), t_pol.v(2), 0, 0, 0, 0, 0, 0, 0)
End If
 is_epolygon = search_for_epolygon(EPol, 0, no%, 0)
End Function

Public Function arrange_two_arc(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, c%, _
      op1%, op2%, op3%, ty As Byte) As Boolean
Dim d%
 d% = three_point_direction(p1%, m_Circ(c%).data(0).data0.center, p2%)
 If d% = -1 Then
 Call exchange_two_integer(p1%, p2%)
 End If
 d% = three_point_direction(p3%, m_Circ(c%).data(0).data0.center, p4%)
  If d% = -1 Then
  Call exchange_two_integer(p4%, p4%)
  End If
 If p2% = p3% And three_point_direction(p1%, m_Circ(c%).data(0).data0.center, p4%) = 1 Then
  ty = 3
   op1% = p1%
    op2% = p2%
      op3% = p4%
  arrange_two_arc = True
 ElseIf p2% = p4% And three_point_direction(p1%, m_Circ(c%).data(0).data0.center, p3%) = 1 Then
  ty = 4
   op1% = p1%
    op2% = p3%
      op3% = p4%
 arrange_two_arc = True
 ElseIf p1% = p4% And three_point_direction(p3%, m_Circ(c%).data(0).data0.center, p2%) = 1 Then
  ty = 5
   op1% = p3%
    op2% = p4%
      op3% = p2%
 arrange_two_arc = True
 ElseIf p2% = p4% And three_point_direction(p3%, m_Circ(c%).data(0).data0.center, p1%) = 1 Then
  ty = 6
   op1% = p3%
    op2% = p1%
      op3% = p2%
 arrange_two_arc = True
 ElseIf p1% = p3% And three_point_direction(p2%, m_Circ(c%).data(0).data0.center, p4%) = 1 Then
  ty = 7
   op1% = p1%
    op2% = p2%
      op3% = p4%
 arrange_two_arc = True
 ElseIf p1% = p3% And three_point_direction(p4%, m_Circ(c%).data(0).data0.center, p2%) = 1 Then
  ty = 8
   op1% = p1%
    op2% = p4%
      op3% = p2%
 arrange_two_arc = True
 Else
  ty = 0
 arrange_two_arc = False
 End If
 End Function

Public Function is_diameter(ByVal p1%, ByVal p2%, ByVal p3%, c%, c_data As condition_data_type) As Boolean
Dim i%, n%
If c% = 0 Then
record_0.data0.condition_data.condition_no = 0 'record0
 If is_three_point_on_line(p1%, p2%, p3%, n%, -1000, 0, 0, c_data, _
         0, 0, 0) = 1 Then
  If c% = 0 Then
  c% = m_circle_number(1, p2%, pointapi0, p1%, p3%, 0, 0, 0, 0, 1, 0, 0, 0, False)
   If c% > 0 Then
    If is_point_in_circle(c%, 0, p3%, 0, 0) = True Then
     is_diameter = True
      Exit Function
   End If
   End If
  End If
 End If
 n% = 0
Else
' If m_poi(p2%).data(0).no_reduce = 0 Then
  If is_three_point_on_line(p1%, p2%, p3%, 0, -1000, 0, 0, c_data, _
        0, 0, 0) = 1 Then
         is_diameter = True
  End If
' Else
 For n% = 1 To m_Circ(c%).data(0).last_Diameter
     If is_same_two_point(p1%, p3%, m_Circ(c%).data(0).Diameter(n%).poi(0), _
           m_Circ(c%).data(0).Diameter(n%).poi(1)) Then
            c_data = m_Circ(c%).data(0).Diameter(n%).cond
            is_diameter = True
             Exit Function
     End If
 Next n%
  n% = 0
 End If
'End If
End Function


Public Function is_new_length(ByVal p1%, ByVal p2%, n%) As Boolean
Dim i%
If p1% > p2% Then
 Call exchange_two_integer(p1%, p2%)
End If
For i% = 1 To last_length
  If p1% = length_(i%).poi(0) And p2% = length_(i%).poi(1) Then
    n% = i%
     is_new_length = False
       Exit Function
  End If
Next i%
is_new_length = True
last_length = last_length + 1
n% = last_length
length_(n%).poi(0) = p1%
length_(n%).poi(1) = p2%
End Function

Public Function is_new_length_point_to_line(ByVal p1%, ByVal p2%, ByVal p3%, n%) As Boolean
Dim i%
For i% = 1 To last_length_point_to_line
 If p1% = length_point_to_line(i%).poi(0) And _
       line_number0(p2%, p3%, 0, 0) = length_point_to_line(i%).line_no Then
   is_new_length_point_to_line = False
    Exit Function
 End If
Next i%
is_new_length_point_to_line = True
last_length_point_to_line = last_length_point_to_line + 1
n% = last_length_point_to_line
length_point_to_line(n%).poi(0) = p1%
length_point_to_line(n%).poi(1) = p2%
length_point_to_line(n%).poi(2) = p3%
length_point_to_line(n%).line_no = line_number0(p2%, p3%, 0, 0)

End Function

Public Function is_new_angle_value_for_measur(ByVal p1%, ByVal p2%, ByVal p3%, n%)
Dim i%, A%
A% = Abs(angle_number(p1%, p2%, p3%, 0, 0))
For i% = 1 To last_angle_value_for_measur '(last_angle_value_for_measur).angle = tA
 If A% = angle_value_for_measur(i%).angle Then
  n% = i%
   is_new_angle_value_for_measur = False
    Exit Function
 End If
Next i%
is_new_angle_value_for_measur = True
last_angle_value_for_measur = last_angle_value_for_measur + 1
n% = last_angle_value_for_measur
angle_value_for_measur(n%).poi(0) = p1%
angle_value_for_measur(n%).poi(1) = p2%
angle_value_for_measur(n%).poi(2) = p3%
angle_value_for_measur(n%).angle = A%

End Function

Public Function is_new_area_polygon(po As polygon, n%)
Dim i%, j%, k%, l%
For i% = 1 To last_Area_polygon
 If po.total_v = Area_polygon(i%).p.total_v Then
  For j% = 0 To po.total_v - 1
   For k% = 0 To po.total_v - 1
    If po.v(j%) = Area_polygon(i%).p.v(k%) Then
       For l% = 0 To po.total_v - 1
        If po.v((j% + l%) Mod po.total_v) <> _
            Area_polygon(i%).p.v((k% + l%) Mod po.total_v) Then
             GoTo is_area_polygon_mark1
        End If
       Next l%
       n% = i%
        is_new_area_polygon = False
         Exit Function
is_area_polygon_mark1:
       For l% = 0 To po.total_v - 1
        If po.v((j% + l%) Mod po.total_v) <> _
            Area_polygon(i%).p.v((k% + po.total_v - l%) Mod po.total_v) Then
          GoTo is_area_polygon_mark2
        End If
       Next l%
           n% = k%
        is_new_area_polygon = False
         Exit Function
is_area_polygon_mark2:
    End If
   Next k%
  Next j%
End If
  Next i%
is_new_area_polygon = True
last_Area_polygon = last_Area_polygon + 1
n% = last_Area_polygon
Area_polygon(n%).p = po
End Function
Public Function arrange_four_point_(ByVal p1%, ByVal p2%, ByVal p3%, _
      ByVal p4%, ty As Byte, op1%, op2%, op3%, op4%) As Boolean
' 根据ty安排点      '
Dim l(1) As Integer
l(0) = line_number0(p1%, p2%, 0, 0)
l(1) = line_number0(p3%, p4%, 0, 0)
   If l(0) = 0 Or l(1) = 0 Or l(0) <> l(1) Or p2% <> p3% Or _
       ty < 3 Then
   Exit Function
   End If
If ty = 3 Then
op1% = p1%
 op2% = p2%
  op3% = p3%
   op4% = p4%
    arrange_four_point_ = True
     Exit Function
ElseIf ty = 4 Then
op1% = p1%
 op2% = p4%
  op3% = p2%
   op4% = p4%
    arrange_four_point_ = True
     Exit Function
ElseIf ty = 5 Then
op1% = p3%
 op2% = p4%
  op3% = p1%
   op4% = p2%
    arrange_four_point_ = True
     Exit Function
ElseIf ty = 6 Then
op1% = p3%
 op2% = p4%
  op3% = p1%
   op4% = p4%
    arrange_four_point_ = True
     Exit Function
ElseIf ty = 7 Then
op1% = p1%
 op2% = p2%
  op3% = p1%
   op4% = p4%
    arrange_four_point_ = True
     Exit Function
ElseIf ty = 8 Then
op1% = p1%
 op2% = p4%
  op3% = p1%
   op4% = p2%
    arrange_four_point_ = True
     Exit Function
End If
End Function
Public Function combine_two_angle_with_para(A1%, A2%, A3%, A3_%, _
                para1 As String, para2 As String, v$, v_$, _
                  ty1 As Byte, ty2 As Byte, ty_ As Byte, re0 As record_data_type) As Boolean 'ty_=0 , 合并ty_=1 推理
'1 是否对顶角
'2 已知关系
'3 同顶点角合并
'4 共边角合并
'Dim tA(3) As Integer
Dim last As Byte
Dim tn%
Dim A3_v  As angle3_value_data0_type
Dim tv As String
'Dim tv_ As String
 If A1% = 0 Or A2% = 0 Then
  Exit Function
 End If
 A3% = 0
 If A1% = A2% Then  '合并
      para1 = add_string(para2, para1, True, False)
       A2% = 0
        para2 = "0"
     If para1 = "0" Then
       A1% = 0
     End If
 Else
 '排序
 If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no And A2% > 0 Then
  Call exchange_two_integer(A1%, A2%)
  Call exchange_string(para1, para2)
 End If
      '*****************************
     A3_v.angle(0) = A1%
     A3_v.angle(1) = A2%
     A3_v.angle(2) = 0
     A3_v.para(0) = para1
     A3_v.para(1) = para2
     A3_v.para(2) = "0"
 If search_for_three_angle_value(A3_v, 0, tn%, 0) And A2% > 0 Then
   '已知关系
       Call add_conditions_to_record(angle3_value_, tn%, 0, 0, re0.data0.condition_data)
        If v$ = v_$ Then 'value=value_
        v$ = minus_string(v$, divide_string(time_string(angle3_value(tn%).data(0).data0.value, A3_v.para(1), _
                  False, False), angle3_value(tn%).data(0).data0.para(1), False, False), True, False)
        v_$ = v$
        Else
        v$ = minus_string(v$, divide_string(time_string(angle3_value(tn%).data(0).data0.value, A3_v.para(1), _
                  False, False), angle3_value(tn%).data(0).data0.para(1), False, False), True, False)
        v_$ = minus_string(v_$, divide_string(time_string(angle3_value(tn%).data(0).data0.value, A3_v.para(1), _
                  False, False), angle3_value(tn%).data(0).data0.para(1), False, False), True, False)
        End If
         para1 = minus_string(A3_v.para(0), divide_string(time_string(angle3_value(tn%).data(0).data0.para(0), _
                    A3_v.para(1), False, False), angle3_value(tn%).data(0).data0.para(1), False, False), True, False)
         para2 = "0"
          A2% = 0
          A3% = 0
        If para1 = "0" Then
          A1% = 0
        Else
          A1% = angle3_value(tn%).data(0).data0.angle(0)
           combine_two_angle_with_para = combine_two_angle_with_para(A1%, 0, 0, 0, para1, "0", v$, v_$, ty1, ty2, ty_, re0)
        End If
          combine_two_angle_with_para = True
         Exit Function
   End If
 If combine_two_Tangle(A1%, A2%, A3%, A3_%, ty1, ty2, last, ty_) Then
   If (ty1 = 15 Or ty1 = 17) And v$ <> "" Then    '平角'有缝A+B+C=180
     If ty_ = 1 Then
          If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
           Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
           If v$ = v_$ Then
            v$ = minus_string("180", angle(A1%).data(0).value, False, False)
             v$ = minus_string(v$, angle(A2%).data(0).value, True, False)
              v_$ = v$
           Else
            v$ = minus_string("180", angle(A1%).data(0).value, False, False)
             v$ = minus_string(v$, angle(A2%).data(0).value, True, False)
              v_$ = minus_string("180", angle(A1%).data(0).value, False, False)
               v_$ = minus_string(v_$, angle(A2%).data(0).value, True, False)
           End If
           A1% = A3%
           para1 = "1"
           para2 = "0"
           A3% = 0
           A2% = 0
           combine_two_angle_with_para = True
           Exit Function
           ElseIf para1 = para2 Then
           A1% = A3%
           A2% = 0
           A3% = 0
           A3_% = 0
           ty1 = 0
           ty2 = 0
           para1 = time_string("-1", para1, True, False)
           para2 = "0"
           v$ = add_string(time_string(para1, "180", False, False), v$, True, False)
           v_$ = add_string(time_string(para1, "180", False, False), v_$, True, False)
           combine_two_angle_with_para = True
           Exit Function
           End If
     Else
     If angle(A3%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
             0, 0, re0.data0.condition_data)
       tv = minus_string("180", angle(A3%).data(0).value, True, False)
       A3% = 0
     Else
     tv = "180"
     End If
    If para1 = para2 Then '可以合并
      A1% = A3%
        A2% = 0
         A3% = 0
           v = minus_string(v, time_string(para1, tv, False, False), True, False)
         para2 = "0"
          If A1% > 0 Then
          para1 = time_string(para1, "-1", True, False)
          Else
          para1 = "0"
          End If
             ty1 = 0
             ty2 = 0
             combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
             angle(A2%).data(0).total_no > angle(A2%).data(0).total_no Then 'tA% 序号最大
             combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
             angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = minus_string(v, time_string(para2, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para1 = minus_string(para1, para2, True, False)
           para2 = time_string(para2, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A2%, A3%)
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
           End If
           Else
           A2% = 0
           para2 = "0"
           ty1 = 0
           End If
            combine_two_angle_with_para = True
             Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A3%).data(0).total_no And _
         angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
           v = minus_string(v, time_string(para1, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para2 = minus_string(para2, para1, True, False)
           para1 = time_string(para1, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A1%, A3%)
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
           End If
           Else
           A1% = 0
           para1 = "0"
           ty1 = 0
           End If
             combine_two_angle_with_para = True
             Exit Function
     End If
     End If
   ElseIf (ty1 = 16 Or ty1 = 18) And v$ <> "" Then    '平角,有重合A+B-C=180
    If ty_ = 1 Then
     If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
         If v$ = v_$ Then
         v$ = minus_string(angle(A1%).data(0).value, "180", False, False)
         v$ = add_string(angle(A2%).data(0).value, v$, True, False)
         v_$ = v$
         Else
         v$ = minus_string(angle(A1%).data(0).value, "180", False, False)
         v$ = add_string(angle(A2%).data(0).value, v$, True, False)
         v_$ = minus_string(angle(A1%).data(0).value, "180", False, False)
         v_$ = add_string(angle(A2%).data(0).value, v_$, True, False)
         End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
     ElseIf para1 = para2 Then
      para2 = "0"
      A1% = A3%
      A2% = 0
      A3% = 0
      A3_% = 0
      ty1 = 0
      ty2 = 0
      v$ = minus_string(v$, time_string(para1, "180", False, False), True, False)
      v_$ = minus_string(v_$, time_string(para1, "180", False, False), True, False)
         combine_two_angle_with_para = True
         Exit Function
     End If
    Else
     If angle(A3%).data(0).value <> "" Then
        tv = add_string("180", angle(A3%).data(0).value, True, False)
         A3% = 0
          Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
            0, 0, re0.data0.condition_data)
     Else
      tv = "180"
     End If
      If para1 = para2 Then
      A1% = A3%
       A3% = 0
          If v$ = v_$ Then
          v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = v$
          Else
          v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = minus_string(v_$, time_string(para1, tv, False, False), True, False)
          End If
       para2 = "0"
        A2% = 0
         ty1 = 0
           If A1% = 0 Then
            para1 = "0"
           End If
            combine_two_angle_with_para = True
             Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then 'tA% 序号最大
              combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = minus_string(v, time_string(para2, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para1 = minus_string(para1, para2, True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A2%, A3%)
           ty1 = 23
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
             ty1 = 24
            End If
            Else
            A2% = 0
            para2 = "0"
            ty1 = 0
            End If
              combine_two_angle_with_para = True
            Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A2%).data(0).total_no And _
         angle(A1%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = minus_string(v, time_string(para1, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para2 = minus_string(para2, para1, True, False)
           'para1 =  para1, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A1%, A3%)
             ty1 = 24
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
             ty1 = 23
           End If
           Else
           A1% = 0
           para1 = "0"
           ty1 = 0
           End If
             combine_two_angle_with_para = True
             Exit Function
     End If
     End If
    ElseIf ty1 = 23 And v$ <> "" Then     '平角,有重合A-B+C=180
     If ty_ = 1 Then
     If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
         If v$ = v_$ Then
         v$ = minus_string("180", angle(A1%).data(0).value, False, False)
         v$ = add_string(angle(A2%).data(0).value, v$, True, False)
         v_$ = v$
         Else
         End If
         v$ = minus_string("180", angle(A1%).data(0).value, False, False)
         v$ = add_string(angle(A2%).data(0).value, v$, True, False)
         v_$ = minus_string("180", angle(A1%).data(0).value, False, False)
         v_$ = add_string(angle(A2%).data(0).value, v_$, True, False)
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
        ElseIf para1 = time_string("-1", para2, True, False) Then
         para1 = para2
         para2 = 0
         A1% = A3%
         A2% = 0
         A3% = 0
         A3_% = 0
         ty1 = 0
         ty2 = 0
         v$ = add_string(v$, time_string(para1, "180", False, False), True, False)
         v_$ = add_string(v_$, time_string(para1, "180", False, False), True, False)
        End If
     Else
     If angle(A3%).data(0).value <> "" Then
      tv = minus_string("180", angle(A3).data(0).value, True, False)
       Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, 0, 0, _
           re0.data0.condition_data)
        A3% = 0
     Else
        tv = "180"
     End If
      If para1 = time_string(para2, "-1", True, False) Then
      A1% = A3%
       A2% = 0
        A3% = 0
      para1 = para2
       para2 = "0"
          If v$ = v_$ Then
          v$ = add_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = v$
          Else
          v$ = add_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = add_string(v_$, time_string(para1, tv, False, False), True, False)
          End If
             ty1 = 0
        If A1% = 0 Then
         para1 = 0
        End If
             combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then 'tA% 序号最大
               Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = add_string(v, time_string(para2, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para1 = add_string(para1, para2, True, False)
           If A3% > 0 Then
            Call exchange_two_integer(A2%, A3%)
           ty1 = 16
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
            End If
           Else
            A2% = 0
            para2 = "0"
            ty1 = 0
           End If
             combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A1%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = minus_string(v, time_string(para1, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para2 = add_string(para2, para1, True, False)
           para1 = time_string(para1, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A1%, A3%)
             If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
            ty1 = 24
            End If
            Else
            A1% = 0
            para1 = "0"
            ty1 = 0
            End If
             combine_two_angle_with_para = True
              Exit Function
     End If
     End If
    ElseIf ty1 = 24 And v$ <> "" Then     '平角,有重合A-B-C=-180
    If ty_ = 1 Then
     If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
         If v$ = v_$ Then
         v$ = minus_string("180", angle(A2%).data(0).value, False, False)
         v$ = add_string(angle(A1%).data(0).value, v$, True, False)
         v_$ = v$
         Else
         v$ = minus_string("180", angle(A2%).data(0).value, False, False)
         v$ = add_string(angle(A1%).data(0).value, v$, True, False)
         v_$ = minus_string("180", angle(A2%).data(0).value, False, False)
         v_$ = add_string(angle(A1%).data(0).value, v_$, True, False)
         End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
       ElseIf para1 = time_string("-1", para2, True, False) Then
        A1% = A3%
        A2% = 0
        A3% = 0
        A3_% = 0
        ty1 = 0
        ty2 = 0
        v$ = add_string(v$, time_string(para1, "180", False, False), True, False)
        v_$ = add_string(v_$, time_string(para1, "180", False, False), True, False)
         combine_two_angle_with_para = True
         Exit Function
       End If
    Else
     If angle(A3%).data(0).value <> "" Then
      tv = minus_string(angle(A3%).data(0).value, "180", True, False)
       Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
            0, 0, re0.data0.condition_data)
       A3% = 0
     Else
      tv = "-180"
     End If
      If para1 = time_string(para2, "-1", True, False) Then
      A1% = A3%
       A2% = 0
        A3% = 0
      para2 = "0"
         If v$ = v_$ Then
          v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = v$
          Else
          v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
          v_$ = minus_string(v_$, time_string(para1, tv, False, False), True, False)
          End If
             combine_two_angle_with_para = True
             If A1% = 0 Then
              para1 = "0"
             End If
              ty1 = 0
              Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then 'tA% 序号最大
               combine_two_angle_with_para = True
              Exit Function
      ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
           v = add_string(v, time_string(para2, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para1 = add_string(para1, para2, True, False)
            para2 = time_string(para2, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A2%, A3%)
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
              ty1 = 23
            End If
            Else
             A2% = 0
              para2 = "0"
               ty1 = 0
            End If
             combine_two_angle_with_para = True
              Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A1%).data(0).total_no And _
         angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
           v = minus_string(v, time_string(para1, tv, False, False), True, False)
             combine_two_angle_with_para = True
           para2 = add_string(para2, para1, True, False)
           'para1 =  para1, "-1", True, False)
           If A3% > 0 Then
           Call exchange_two_integer(A1%, A3%)
           ty1 = 16
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
            Call exchange_string(para1, para2)
             End If
            Else
            A1% = 0
            para1 = "0"
            ty1 = 0
            End If
              combine_two_angle_with_para = True
             Exit Function
     End If
     End If
   ElseIf ty1 = 20 Or ty1 = 19 Then 'A+B=180
      If para1 = para2 Then
       A1% = 0
        A2% = 0
         If v$ = v_$ Then
         v$ = minus_string(v$, time_string(para1, "180", False, False), True, False)
         v_$ = v$
         Else
         v$ = minus_string(v$, time_string(para1, "180", False, False), True, False)
         v_$ = minus_string(v_$, time_string(para1, "180", False, False), True, False)
         End If
          para1 = "0"
           para2 = "0"
            ty1 = 0
            combine_two_angle_with_para = True
             Exit Function
      ElseIf angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
       A2% = 0
        A3% = 0
         If v$ = v_$ Then
         v$ = minus_string(v$, time_string(para2, "180", False, False), True, False)
         v_$ = v$
         Else
         v$ = minus_string(v$, time_string(para2, "180", False, False), True, False)
         v_$ = minus_string(v_$, time_string(para2, "180", False, False), True, False)
         End If
          para1 = minus_string(para1, para2, True, False)
           para2 = "0"
           ty1 = 0
            combine_two_angle_with_para = True
             Exit Function
      Else
        A1% = A2%
        A2% = 0
        A3% = 0
        If v$ = v_$ Then
         v$ = minus_string(v$, time_string(para1, "180", False, False), True, False)
         v_$ = v$
        Else
         v$ = minus_string(v$, time_string(para1, "180", False, False), True, False)
         v_$ = minus_string(v_$, time_string(para1, "180", False, False), True, False)
        End If
          para1 = minus_string(para2, para1, True, False)
           para2 = "0"
           ty1 = 0
            combine_two_angle_with_para = True
             Exit Function
     End If
    ElseIf ty1 = 21 Then '对顶角
      para1 = add_string(para2, para1, True, False)
       If para1 = "0" Then
        A1% = 0
       End If
        A2% = 0
         para2 = "0"
          combine_two_angle_with_para = True
             Exit Function
    ElseIf ty1 = 3 Or ty1 = 5 Then '和A+B-C=0
     If ty_ = 1 Then
     If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
        If v$ = v_$ Then
         v$ = add_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
         v_$ = v$
        Else
         v$ = add_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
         v_$ = add_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
        End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
     End If
     Else
     If angle(A3%).data(0).value <> "" Then
     tv = angle(A3%).data(0).value
     Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, 0, _
           0, re0.data0.condition_data)
     A3% = 0
     Else
     tv = "0"
     End If
      If para2 = para1 Then
       If v$ = v_$ Then
       v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
       v_$ = v$
       Else
       v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
       v_$ = minus_string(v_$, time_string(tv, para1, False, True), True, False)
       End If
       A1% = A3%
        A2% = 0
         A3% = 0 ' tA2%
          para2 = "0"
           ty1 = 0
           If A1% = 0 Then
           para1 = "0"
           End If
           combine_two_angle_with_para = True
             Exit Function
      ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then
               combine_two_angle_with_para = True
                 Exit Function
      ElseIf angle(A1%).data(0).total_no > angle(A3%).data(0).total_no And _
          angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
         If v$ = v_$ Then
         v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
         v_$ = v$
         Else
         v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
         v_$ = minus_string(v_$, time_string(tv, para1, False, True), True, False)
         End If
         para2 = minus_string(para2, para1, True, False)
          If A3% > 0 Then
            Call exchange_two_integer(A1%, A3%)
             ty1 = 4
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            ty1 = 3
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
          Else
          A1% = 0
          para1 = "0"
          ty1 = 0
          End If
            combine_two_angle_with_para = True
         Exit Function
      ElseIf angle(A2%).data(0).total_no > angle(A3%).data(0).total_no And _
          angle(A2%).data(0).total_no > angle(A1%).data(0).total_no Then
       If v$ = v_$ Then
       v$ = minus_string(v$, time_string(tv, para2, False, True), True, False)
       v_$ = v$
       Else
       v$ = minus_string(v$, time_string(tv, para2, False, True), True, False)
       v_$ = minus_string(v_$, time_string(tv, para2, False, True), True, False)
       End If
          para1 = minus_string(para1, para2, True, False)
          If A3% > 0 Then
          Call exchange_two_integer(A2%, A3%)
             ty1 = 3
          If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            ty1 = 4
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
           Else
           A2% = 0
           para2 = "0"
           End If
            combine_two_angle_with_para = True
          Exit Function
        End If
     End If
    ElseIf ty1 = 4 Or ty1 = 8 Then 'A-B-C=0
    If ty_ = 1 Then
      If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
         If v$ = v_$ Then
         v$ = minus_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
         v_$ = v$
         Else
         v$ = minus_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
         v_$ = minus_string(angle(A1%).data(0).value, angle(A2%).data(0).value, True, False)
         End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
      ElseIf para1 = time_string("-1", para2, True, False) Then
      A1% = A3%
      A2% = 0
      A3% = 0
      A3_% = 0
      ty1 = 0
      ty2 = 0
      para2 = "0"
          combine_two_angle_with_para = True
         Exit Function
      Else
      A1% = A3%
      A3% = 0
      ty1 = 0
      ty2 = 0
      para2 = add_string(para1, para2, True, False)
          combine_two_angle_with_para = True
         Exit Function
      End If
   Else
     If angle(A3%).data(0).value <> "" Then
      tv = angle(A3%).data(0).value
       Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
          0, 0, re0.data0.condition_data)
      A3% = 0
     Else
      tv = "0"
     End If
     If para1 = time_string(para2, "-1", True, False) Then
      If v$ = v_$ Then
      v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
      v_$ = v$
      Else
      v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
      v_$ = minus_string(v_$, time_string(tv, para1, False, True), True, False)
      End If
       para2 = 0
        A1% = A3%
         A2% = 0
          A3% = 0
       ty1 = 0
       If A1% = 0 Then
        para1 = "0"
       End If
       combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then
       combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
       If v$ = v_$ Then
       v$ = add_string(v$, time_string(tv, para2, False, True), True, False)
       v_$ = v$
       Else
       v$ = add_string(v$, time_string(tv, para2, False, True), True, False)
       v_$ = add_string(v_$, time_string(tv, para2, False, True), True, False)
       End If
        para1 = add_string(para1, para2, True, False)
        para2 = time_string(para2, "-1", True, False)
         If A3% > 0 Then
                Call exchange_two_integer(A2%, A3%)
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            ty1 = 6
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
         Else
         A2% = 0
         para2 = "0"
         End If
       combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A2%).data(0).total_no And _
              angle(A1%).data(0).total_no > angle(A3%).data(0).total_no Then
              If v$ = v_$ Then
               v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
               v_$ = v$
              Else
               v$ = minus_string(v$, time_string(tv, para1, False, True), True, False)
               v_$ = minus_string(v_$, time_string(tv, para1, False, True), True, False)
              End If
        para2 = add_string(para1, para2, True, False)
        If A3% > 0 Then
         Call exchange_two_integer(A1%, A3%)
          ty1 = 3
           If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
         Else
         A1% = 0
         para1 = "0"
         ty1 = 0
         End If
         combine_two_angle_with_para = True
       Exit Function
     End If
     End If
     ElseIf ty1 = 6 Or ty1 = 7 Then 'B-A=C
     If ty_ = 1 Then
          If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
        If v$ = v_$ Then
         v$ = minus_string(angle(A2%).data(0).value, angle(A1%).data(0).value, True, False)
         v_$ = v$
        Else
         v$ = minus_string(angle(A2%).data(0).value, angle(A1%).data(0).value, True, False)
         v_$ = minus_string(angle(A2%).data(0).value, angle(A1%).data(0).value, True, False)
        End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
        ElseIf para1 = time_string("-1", para2, True, False) Then
         para1 = para2
         para2 = 0
         A1% = A3%
         A2% = 0
         A3% = 0
         A3_% = 0
         ty1 = 0
         ty2 = 0
         combine_two_angle_with_para = True
         Exit Function
        Else
         A2% = A3%
         A3% = 0
         ty1 = 0
         ty2 = 0
         para1 = add_string(para1, para2, True, False)
         combine_two_angle_with_para = True
        End If
     Else
     If angle(A3%).data(0).value <> "" Then
     tv = angle(A3%).data(0).value
     Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
          0, 0, re0.data0.condition_data)
     A3% = 0
     Else
     tv = "0"
     End If
      If para1 = time_string(para2, "-1", True, False) Then
      If v$ = v_$ Then
      v$ = minus_string(v$, time_string(tv, para2, False, False), True, False)
      v_$ = v$
      Else
      v$ = minus_string(v$, time_string(tv, para2, False, False), True, False)
      v_$ = minus_string(v_$, time_string(tv, para2, False, False), True, False)
      End If
       para1 = para2
       para2 = "0"
       A1% = A3%
       A3% = 0
       A2% = 0
       ty1 = 0
       If A1% = 0 Then
       para1 = "0"
       End If
       combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then
       combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
      If v$ = v_$ Then
      v$ = minus_string(v$, time_string(tv, para2, False, False), True, False)
      v_$ = v$
      Else
      v$ = minus_string(v$, time_string(tv, para2, False, False), True, False)
      v_$ = minus_string(v_$, time_string(tv, para2, False, False), True, False)
      End If
               para1 = add_string(para1, para2, True, False)
          If A3% > 0 Then
               Call exchange_two_integer(A3%, A2%)
               ty1 = 3
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
           Else
            A2% = 0
            para2 = "0"
            ty1 = 0
           End If
         combine_two_angle_with_para = True
       Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A2%).data(0).total_no And _
              angle(A1%).data(0).total_no > angle(A3%).data(0).total_no Then
              If v$ = v_$ Then
                    v$ = add_string(v$, time_string(tv, para1, False, False), True, False)
                    v_$ = v$
              Else
                    v$ = add_string(v$, time_string(tv, para1, False, False), True, False)
                    v_$ = add_string(v_$, time_string(tv, para1, False, False), True, False)
              End If
              para2 = add_string(para1, para2, True, False)
              para1 = time_string(para1, "-1", True, False)
            If A3% > 0 Then
              Call exchange_two_integer(A2%, A3%)
            If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
             ty1 = 3
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
           Else
           A1% = 0
           para1 = "0"
           ty1 = 0
           End If
         combine_two_angle_with_para = True
       Exit Function
     End If
     End If
    ElseIf ty1 = 9 Or ty1 = 10 Then 'A+B+C=360
    If ty_ = 1 Then
     If angle(A1%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
       Call add_conditions_to_record(angle3_value_, angle(A1%).data(0).value_no, _
              angle(A2%).data(0).value_no, 0, re0.data0.condition_data)
         If v$ = v_$ Then
         v$ = minus_string("360", angle(A1%).data(0).value, False, False)
         v$ = minus_string(v$, angle(A2%).data(0).value, True, False)
         v_$ = v$
         Else
         v$ = minus_string("360", angle(A1%).data(0).value, False, False)
         v$ = minus_string(v$, angle(A2%).data(0).value, True, False)
         v_$ = minus_string("360", angle(A1%).data(0).value, False, False)
         v_$ = minus_string(v_$, angle(A2%).data(0).value, True, False)
         End If
         A1% = A3%
         para1 = "1"
         para2 = "0"
         A3% = 0
         A2% = 0
         combine_two_angle_with_para = True
         Exit Function
     ElseIf para1 = para2 Then
      para1 = time_string("-1", para1, True, False)
      para2 = "0"
      A1% = A3%
      A2% = 0
      A3% = 0
      A3_% = 0
      ty1 = 0
      ty2 = 0
      v$ = add_string(v$, time_string("360", para1, False, False), True, False)
      v_$ = add_string(v_$, time_string("360", para1, False, False), True, False)
         combine_two_angle_with_para = True
         Exit Function
     End If
    Else
    If angle(A3%).data(0).value <> "" Then
    tv = minus_string("360", angle(A3%).data(0).value, True, False)
    Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, _
          0, 0, re0.data0.condition_data)
          A3% = 0
    Else
    tv = "360"
    End If
     If para1 = para2 Then
     A1% = A3%
      A2% = 0
       A3% = 0
       para2 = "0"
        para1 = time_string(para1, "-1", True, False)
         v = add_string(v, time_string(para1, tv, False, False), True, False)
          ty1 = 0
          If A1% = 0 Then
           para1 = "0"
          End If
         combine_two_angle_with_para = True
          Exit Function
     ElseIf angle(A3%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A3%).data(0).total_no > angle(A2%).data(0).total_no Then
         combine_two_angle_with_para = True
          Exit Function
     ElseIf angle(A2%).data(0).total_no > angle(A1%).data(0).total_no And _
              angle(A2%).data(0).total_no > angle(A3%).data(0).total_no Then
        If v$ = v_$ Then
        v$ = minus_string(v$, time_string(para2, tv, False, False), True, False)
        v_$ = v$
        Else
        v$ = minus_string(v$, time_string(para2, tv, False, False), True, False)
        v_$ = minus_string(v_$, time_string(para2, tv, False, False), True, False)
        End If
        para1 = minus_string(para1, para2, True, False)
        para2 = time_string("-1", para2, True, False)
        If A3% > 0 Then
        Call exchange_two_integer(A2%, A3%)
             If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
        Else
        A2% = 0
        para2 = "0"
        ty1 = 0
        End If
         combine_two_angle_with_para = True
          Exit Function
     ElseIf angle(A1%).data(0).total_no > angle(A2%).data(0).total_no And _
              angle(A1%).data(0).total_no > angle(A3%).data(0).total_no Then
        If v$ = v_$ Then
        v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
        v_$ = v$
        Else
        v$ = minus_string(v$, time_string(para1, tv, False, False), True, False)
        v_$ = minus_string(v_$, time_string(para1, tv, False, False), True, False)
        End If
        para2 = minus_string(para2, para1, True, False)
        para1 = time_string("-1", para1, True, False)
        If A3% > 0 Then
        Call exchange_two_integer(A1%, A3%)
             If angle(A1%).data(0).total_no > angle(A2%).data(0).total_no Then
            Call exchange_two_integer(A1%, A2%)
             Call exchange_string(para1, para2)
           End If
        Else
        A1% = 0
        para1 = "0"
        ty1 = 0
        End If
         combine_two_angle_with_para = True
          Exit Function
    End If
    End If
    End If
   End If
   End If
End Function

Public Function is_four_point_in_epolygon(ByVal p1%, _
     ByVal p2%, ByVal p3%, ByVal p4%, e%, _
        ty As Byte, s$) As Boolean
'ty 类型，s$, 比
Dim i%
Dim t_y(1) As Byte
If is_same_two_point(p1%, p2%, p3%, p4%) Then
 s$ = "1"
  Exit Function
End If
For i% = 1 To last_conditions.last_cond(1).squre_no
    If Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(0) = p1% And _
         Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(1) = p2% And _
          Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(2) = p3% And _
           Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(3) = p4% Then
      is_four_point_in_epolygon = True
      Exit Function
    End If
Next i%
For i% = 1 To last_conditions.last_cond(1).epolygon_no
 t_y(0) = is_two_point_in_polygon(p1%, p2%, epolygon(i%).data(0).p)
  t_y(1) = is_two_point_in_polygon(p3%, p4%, epolygon(i%).data(0).p)
 If epolygon(i%).data(0).p.total_v > 3 Then
 If epolygon(i%).data(0).p.total_v = 4 Then
  If t_y(0) = 1 And t_y(1) = 1 Then
   ty = eline_
    s$ = "1"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 2 And t_y(1) = 2 Then
   ty = eline_
    s$ = "1"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 1 And t_y(1) = 2 Then
   ty = relation_
    s$ = "'2&2" 'LoadResString_(1127)
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 2 And t_y(1) = 1 Then
   ty = relation_
    s$ = "'2" ' LoadResString_(1128)
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  End If
ElseIf epolygon(i%).data(0).p.total_v = 5 Then
  If (t_y(0) = 1 And t_y(1) = 1) Or (t_y(0) = 2 And t_y(1) = 2) Then
   ty = eline_
    s$ = "1"
     is_four_point_in_epolygon = True
      e% = i%
       Exit Function
  ElseIf t_y(0) = 1 And t_y(1) = 2 Then
   ty = relation_
      s$ = divide_string(minus_string("'5", "1", False, False), "2", True, False)
       e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 2 And t_y(1) = 1 Then
   ty = relation_
      s$ = divide_string(add_string("1", "'5", False, False), "2", True, False)
      e% = i%
     is_four_point_in_epolygon = True
      Exit Function
   End If
ElseIf epolygon(i%).data(0).p.total_v = 6 Then
  If (t_y(0) = 1 And t_y(1) = 1) Or (t_y(0) = 2 And t_y(1) = 2) Or _
        (t_y(0) = 3 And t_y(1) = 3) Then
   ty = eline_
    s$ = "1"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 1 And t_y(1) = 2 Then
   ty = relation_
      s$ = "'3/3"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 2 And t_y(1) = 1 Then
   ty = relation_
    s$ = "'3"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 1 And t_y(1) = 3 Then
   ty = relation_
      s$ = "1/2"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 1 And t_y(1) = 1 Then
   ty = relation_
    s$ = "2"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
    ElseIf t_y(0) = 2 And t_y(1) = 3 Then
   ty = relation_
      s$ = "'3&2"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
  ElseIf t_y(0) = 3 And t_y(1) = 2 Then
   ty = relation_
    s$ = "2'3&3"
     e% = i%
     is_four_point_in_epolygon = True
      Exit Function
 End If
End If
End If
Next i%
End Function
Public Function is_equal_side_tixing(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
     no%, op1%, op2%, op3%, op4%, poly4_no%, n1%, cond_ty As Byte) As Boolean
Dim tn(3) As Integer
Dim i%, l1%, l2%
End Function
Public Function is_long_squre(ByVal p1%, ByVal p2%, ByVal p3%, _
       ByVal p4%, no%, poly4_no%, n1%, cond_ty As Byte) As Boolean
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, 0)
If poly4_no% = 0 Then
 If n1% <> -1000 Then
  no% = 0
  is_long_squre = True
 Else
  is_long_squre = False
 End If
 no% = 0
 Exit Function
Else
 If Dpolygon4(poly4_no%).data(0).ty = Squre Or Dpolygon4(poly4_no%).data(0).ty = long_squre_ Then
  cond_ty = Dpolygon4(poly4_no%).data(0).ty
  no% = Dpolygon4(poly4_no%).data(0).no
  is_long_squre = True
  Exit Function
 End If
End If
is_long_squre = is_long_squre0(poly4_no%, no%, n1%, cond_ty)
End Function

Public Function is_condition_in_record(con_ty As Integer, con_no As Integer, _
       re As record_data_type, level As Byte) As Boolean
Dim i%
Dim temp_record As total_record_type
If re.data0.condition_data.condition_no > 8 Or re.data0.condition_data.condition_no = 0 Or level = 0 Then
  Exit Function
Else
For i% = 1 To re.data0.condition_data.condition_no
 If re.data0.condition_data.condition(i%).ty = con_ty And re.data0.condition_data.condition(i%).no = con_no And _
         re.data0.theorem_no = 1 Then
  is_condition_in_record = True
   Exit Function
 Else
  Call record_no(re.data0.condition_data.condition(i%).ty, _
         re.data0.condition_data.condition(i%).no, temp_record, False, 0, 0)
   is_condition_in_record = is_condition_in_record(con_ty, con_no, temp_record.record_data, _
       level - 1)
    If is_condition_in_record = True Then
     Exit Function
    End If
 End If
Next i%
End If
 'is_condition_in_record = False
End Function
Public Function is_condition_in_record_(con_ty As Byte, con_no As Integer, _
    re As record_data_type) As Byte
Dim i%
Dim temp_record As total_record_type
If re.data0.condition_data.condition_no > 8 Or re.data0.condition_data.condition_no < 1 Then
  Exit Function
Else
For i% = 1 To re.data0.condition_data.condition_no
 If re.data0.condition_data.condition(i%).ty = con_ty And re.data0.condition_data.condition(i%).no = con_no Then ' And _
         re.data0.theorem_no = 1 Then
  is_condition_in_record_ = 1
   Exit Function
 Else
  Call record_no(re.data0.condition_data.condition(i%).ty, _
     re.data0.condition_data.condition(i%).no, temp_record, False, 0, 0)
   is_condition_in_record_ = is_condition_in_record_(con_ty, con_no, temp_record.record_data)
    If is_condition_in_record_ = 1 Then
     Exit Function
    End If
 End If
Next i%
End If
 'is_condition_in_record = False
End Function
Public Function is_rhombus(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
        n%, poly4_no%, tn%, ty As Byte) As Boolean
Dim i%
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, 0)
If poly4_no% = 0 Then
  If tn% = -1000 Then
  n% = 0
  is_rhombus = False
  Else
   is_rhombus = True
  End If
  n% = 0
Else
 If Dpolygon4(poly4_no%).data(0).ty = Squre Or Dpolygon4(poly4_no%).data(0).ty = rhombus_ Then
  n% = Dpolygon4(poly4_no%).data(0).no
   ty = Dpolygon4(poly4_no%).data(0).ty
    is_rhombus = True
 Else
 is_rhombus = is_rhombus0(poly4_no%, n%, tn%, ty)
  End If
End If
End Function

Public Function is_tixing(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, no%, _
        op1%, op2%, op3%, op4%, poly4_no%, n1%, paral_no%, cond_ty As Byte, _
         set_or_reduce As Boolean) As Boolean
Dim tn(3) As Integer
Dim i%, l1%, l2%, dir%
Dim tl(1) As Integer
Dim temp_record As record_data_type
'*******************************
'判断是否平行
paral_no% = 0
l1% = line_number0(p1%, p2%, tn(0), tn(1))
l2% = line_number0(p4%, p3%, tn(2), tn(3))
If (tn(0) > tn(1) And tn(2) < tn(3)) Or _
       (tn(0) < tn(1) And tn(2) > tn(3)) Or l1% = l2% Or _
         is_dparal(l1%, l2%, paral_no%, -1000, 0, 0, 0, 0) = False Then
 If n1% <> -1000 Then
 is_tixing = True
 Else
 is_tixing = False
   ' no% = 0 ' 一定不是梯形
     Exit Function
 End If
End If
'***********************************************************
'读出四边形
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, dir%)
If poly4_no% = 0 Then
 If n1% <> -1000 Then
  no% = 0
  is_tixing = True
 Else
  is_tixing = False
 End If
 Exit Function
End If
'*******************************************
cond_ty = Dpolygon4(poly4_no%).data(0).ty
If dir% = 1 Then
op1% = Dpolygon4(poly4_no%).data(0).poi(0)
op2% = Dpolygon4(poly4_no%).data(0).poi(1)
op3% = Dpolygon4(poly4_no%).data(0).poi(2)
op4% = Dpolygon4(poly4_no%).data(0).poi(3)
ElseIf dir = -1 Then
op1% = Dpolygon4(poly4_no%).data(0).poi(0)
op2% = Dpolygon4(poly4_no%).data(0).poi(3)
op3% = Dpolygon4(poly4_no%).data(0).poi(2)
op4% = Dpolygon4(poly4_no%).data(0).poi(1)
'ElseIf dir% = -1 Then
'op1% = Dpolygon4(poly4_no%).data(0).poi(0)
'op2% = Dpolygon4(poly4_no%).data(0).poi(3)
'op3% = Dpolygon4(poly4_no%).data(0).poi(2)
'op4% = Dpolygon4(poly4_no%).data(0).poi(1)
'ElseIf dir% = -2 Then
'op1% = Dpolygon4(poly4_no%).data(0).poi(3)
'op2% = Dpolygon4(poly4_no%).data(0).poi(2)
'op3% = Dpolygon4(poly4_no%).data(0).poi(1)
'op4% = Dpolygon4(poly4_no%).data(0).poi(0)
End If
If Dpolygon4(poly4_no%).data(0).ty > 0 Then
 no% = Dpolygon4(poly4_no%).data(0).no
 is_tixing = True
  Exit Function
End If
'**************************************************************
If Dpolygon4(poly4_no%).data(0).line_no(0) = l1% Or _
      Dpolygon4(poly4_no%).data(0).line_no(0) = l2% Then
op1% = Dpolygon4(poly4_no%).data(0).poi(0)
 op2% = Dpolygon4(poly4_no%).data(0).poi(1)
  op3% = Dpolygon4(poly4_no%).data(0).poi(2)
   op4% = Dpolygon4(poly4_no%).data(0).poi(3)
Else
op1% = Dpolygon4(poly4_no%).data(0).poi(1)
 op2% = Dpolygon4(poly4_no%).data(0).poi(2)
  op3% = Dpolygon4(poly4_no%).data(0).poi(3)
   op4% = Dpolygon4(poly4_no%).data(0).poi(0)
    If n1% <> -1000 Then
     Dpolygon4(poly4_no%).data(0).start_poi = 1
    End If
End If
cond_ty = tixing_
no% = 0
is_tixing = False
End Function

Public Function is_three_line_value(ByVal p1%, ByVal p2%, ByVal p3%, _
     ByVal p4%, ByVal p5%, ByVal p6%, ByVal in1%, ByVal in2%, ByVal in3%, _
      ByVal in4%, ByVal in5%, ByVal in6%, ByVal il1%, ByVal il2%, _
       ByVal il3%, ByVal para1 As String, ByVal para2 As String, _
        ByVal para3 As String, ByVal v As String, n%, n1%, n2%, n3%, _
         n4%, n5%, n6%, l3_value As line3_value_data0_type, con_ty As Byte, _
          c_data As condition_data_type, find_conclusion As Byte) As Byte
         'a_oA1%, a_oA2%, a_opara1 As String, a_opara2 As String, _
           a_ov As String, con_ty As Byte) As Boolean
'A1%,para1,v 输入，oA1%,opear1,oA1%,输出，con_ty, 类型
Dim i%, j%, tn%
Dim ts As String
Dim tl(2) As Integer
Dim t_n(5) As Integer
Dim tp(4) As Integer
Dim t_para(3) As String
Dim t_ppara0(2) As String
Dim ty As Byte
Dim depend_no As Integer
Dim t_l3_value As line3_value_data0_type
Dim l2_value  As two_line_value_data0_type
Dim l_value As line_value_data0_type
Dim temp_v$
Dim temp_record As total_record_type
Dim is_no_initial As Integer
Dim tc_data As condition_data_type
'If n1% = -5000 Then 'simple_dbase_for_line
' l3_value.poi(0) = p1%
'  l3_value.poi(1) = p2%
'   l3_value.poi(2) = p3%
'    l3_value.poi(3) = p4%
'     l3_value.poi(4) = p5%
'      l3_value.poi(5) = p6%
' l3_value.n(0) = n1%
'  l3_value.n(1) = n2%
'   l3_value.n(2) = n3%
'    l3_value.n(3) = n4%
'     l3_value.n(4) = n4%
'      l3_value.n(5) = n6%
' l3_value.line_no(0) = il1%
'  l3_value.line_no(2) = il2%
'   l3_value.line_no(3) = il3%
' l3_value.para(0) = para1
'  l3_value.para(1) = para2
'   l3_value.para(2) = para3
' l3_value.value = v
'   GoTo search_for_three_line_value
'End If
l3_value = t_l3_value
If in1% = 0 Or in2% = 0 Then
 il1% = line_number0(p1%, p2%, in1%, in2%)
  If in1% > in2% Then
   Call exchange_two_integer(in1%, in2%)
    Call exchange_two_integer(p1%, p2%)
  End If
End If
If in3% = 0 Or in4% = 0 Then
 il2% = line_number0(p3%, p4%, in3%, in4%)
  If in3% > in4% Then
   Call exchange_two_integer(in3%, in4%)
    Call exchange_two_integer(p3%, p4%)
  End If
End If
If in5% = 0 Or in6% = 0 Then
 il3% = line_number0(p5%, p6%, in5%, in6%)
  If in5% > in6% Then
   Call exchange_two_integer(in5%, in6%)
    Call exchange_two_integer(p5%, p6%)
  End If
End If
If p1% = p3% And p2% = p4% Then
 para1 = add_string(para1, para2, True, False)
  p3% = 0
   p4% = 0
  in3% = 0
   in4% = 0
  il2% = 0
  If para1 = "0" Then
   p1% = 0
    p2% = 0
   n1% = 0
    n2% = 0
     il1% = 0
  End If
ElseIf p1% = p5% And p2% = p6% Then
 para1 = add_string(para1, para3, True, False)
  p5% = 0
   p6% = 0
  in5% = 0
   in6% = 0
  il3% = 0
  If para1 = "0" Then
   p1% = 0
    p2% = 0
   n1% = 0
    n2% = 0
     il1% = 0
  End If
ElseIf p3% = p5% And p4% = p6% Then
 para2 = add_string(para2, para3, True, False)
  p5% = 0
   p6% = 0
  in5% = 0
   in6% = 0
  il3% = 0
  If para2 = "0" Then
   p3% = 0
    p4% = 0
   n3% = 0
    n4% = 0
     il2% = 0
   End If
End If
n% = 0
'排序
  If para1 = "0" Or p1% = p2% Then
  is_three_line_value = is_two_line_value(p3%, p4%, p5%, p6%, _
       in3%, in4%, in5%, in6%, il2%, il3%, para2, para3, v, n%, _
        n1%, 0, 0, 0, two_line_value_data0, con_ty, c_data)
     l3_value.poi(0) = two_line_value_data0.poi(0)
      l3_value.poi(1) = two_line_value_data0.poi(1)
     l3_value.poi(2) = two_line_value_data0.poi(2)
      l3_value.poi(3) = two_line_value_data0.poi(3)
     l3_value.n(0) = two_line_value_data0.n(0)
      l3_value.n(1) = two_line_value_data0.n(1)
     l3_value.n(2) = two_line_value_data0.n(2)
      l3_value.n(3) = two_line_value_data0.n(3)
     l3_value.line_no(0) = two_line_value_data0.line_no(0)
      l3_value.line_no(1) = two_line_value_data0.line_no(1)
     l3_value.para(0) = two_line_value_data0.para(0)
      l3_value.para(1) = two_line_value_data0.para(1)
     l3_value.value = two_line_value_data0.value
  ElseIf para2 = "0" Or p3% = p4% Then
  is_three_line_value = is_two_line_value(p1%, p2%, p5%, p6%, _
       in1%, in2%, in5%, in6%, il1%, il3%, para1, para3, v, n%, _
         n1%, 0, 0, 0, two_line_value_data0, con_ty, c_data)
     l3_value.poi(0) = two_line_value_data0.poi(0)
      l3_value.poi(1) = two_line_value_data0.poi(1)
     l3_value.poi(2) = two_line_value_data0.poi(2)
      l3_value.poi(3) = two_line_value_data0.poi(3)
     l3_value.n(0) = two_line_value_data0.n(0)
      l3_value.n(1) = two_line_value_data0.n(1)
     l3_value.n(2) = two_line_value_data0.n(2)
      l3_value.n(3) = two_line_value_data0.n(3)
     l3_value.line_no(0) = two_line_value_data0.line_no(0)
      l3_value.line_no(1) = two_line_value_data0.line_no(1)
     l3_value.para(0) = two_line_value_data0.para(0)
      l3_value.para(1) = two_line_value_data0.para(1)
     l3_value.value = two_line_value_data0.value
          Exit Function
  ElseIf para3 = "0" Or p5% = p6% Then
  is_three_line_value = is_two_line_value(p1%, p2%, p3%, p4%, _
    in1%, in2%, in3%, in4%, il1%, il2%, para1, para2, v, n%, _
      n1%, 0, 0, 0, two_line_value_data0, con_ty, c_data)
     l3_value.poi(0) = two_line_value_data0.poi(0)
      l3_value.poi(1) = two_line_value_data0.poi(1)
     l3_value.poi(2) = two_line_value_data0.poi(2)
      l3_value.poi(3) = two_line_value_data0.poi(3)
     l3_value.n(0) = two_line_value_data0.n(0)
      l3_value.n(1) = two_line_value_data0.n(1)
     l3_value.n(2) = two_line_value_data0.n(2)
      l3_value.n(3) = two_line_value_data0.n(3)
     l3_value.line_no(0) = two_line_value_data0.line_no(0)
      l3_value.line_no(1) = two_line_value_data0.line_no(1)
     l3_value.para(0) = two_line_value_data0.para(0)
      l3_value.para(1) = two_line_value_data0.para(1)
     l3_value.value = two_line_value_data0.value
          Exit Function
  End If
Call arrange_four_point(p1%, p2%, p3%, p4%, in1%, in2%, in3%, in4%, _
        il1%, il2%, l3_value.poi(0), l3_value.poi(1), l3_value.poi(2), _
         l3_value.poi(3), 0, 0, l3_value.n(0), l3_value.n(1), l3_value.n(2), _
          l3_value.n(3), 0, 0, l3_value.line_no(0), l3_value.line_no(1), 0, ty, tc_data, _
           is_no_initial)
   If is_no_initial = 1 And n1% = 0 Then
    Call add_record_to_record(tc_data, c_data)
   End If
If ty = 2 Then
 l3_value.para(0) = add_string(para1, para2, True, False)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
ElseIf ty = 0 Or ty = 3 Then
l3_value.para(0) = para1
l3_value.para(1) = para2
If ty = 3 And para1 = para2 Then
 l3_value.poi(1) = l3_value.poi(3)
 l3_value.n(1) = l3_value.n(3)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
End If
ElseIf ty = 1 Or ty = 5 Then
l3_value.para(0) = para2
l3_value.para(1) = para1
If ty = 5 And para1 = para2 Then
 l3_value.poi(1) = l3_value.poi(3)
 l3_value.n(1) = l3_value.n(3)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
End If
ElseIf ty = 4 Then
l3_value.para(0) = para1
l3_value.para(1) = add_string(para1, para2, True, False)
ElseIf ty = 6 Then
l3_value.para(0) = para2
l3_value.para(1) = add_string(para1, para2, True, False)
ElseIf ty = 7 Then
l3_value.para(0) = add_string(para1, para2, True, False)
l3_value.para(1) = para2
ElseIf ty = 8 Then
l3_value.para(0) = add_string(para1, para2, True, False)
l3_value.para(1) = para1
'Else
'l3_value.para(0) = para1
'l3_value.para(1) = para2
End If
If l3_value.para(0) = "0" Then
   l3_value.line_no(0) = 0
   l3_value.poi(0) = 0
   l3_value.poi(1) = 0
   l3_value.n(0) = 0
   l3_value.n(1) = 0
End If
If l3_value.para(1) = "0" Then
   l3_value.line_no(1) = 0
   l3_value.poi(2) = 0
   l3_value.poi(3) = 0
   l3_value.n(2) = 0
   l3_value.n(3) = 0
End If
Call arrange_four_point(l3_value.poi(2), l3_value.poi(3), p5%, p6%, _
        l3_value.n(2), l3_value.n(3), in5%, in6%, l3_value.line_no(1), _
         il3%, l3_value.poi(2), l3_value.poi(3), l3_value.poi(4), _
          l3_value.poi(5), 0, 0, l3_value.n(2), l3_value.n(3), l3_value.n(4), _
           l3_value.n(5), 0, 0, l3_value.line_no(1), l3_value.line_no(2), 0, ty, tc_data, is_no_initial)
    If is_no_initial = 1 And n1% = 0 Then
     Call add_record_to_record(tc_data, c_data)
    End If
If ty = 2 Then
 l3_value.para(1) = add_string(l3_value.para(1), para3, True, False)
 l3_value.poi(4) = 0
 l3_value.poi(5) = 0
 l3_value.n(4) = 0
 l3_value.n(5) = 0
 l3_value.line_no(2) = 0
 l3_value.para(2) = "0"
ElseIf ty = 0 Or ty = 3 Then
'l3_value.para(1) = l3_value.para(1)
l3_value.para(2) = para3
If ty = 3 And l3_value.para(1) = l3_value.para(2) Then
 l3_value.poi(3) = l3_value.poi(5)
 l3_value.n(3) = l3_value.n(5)
 l3_value.poi(4) = 0
 l3_value.poi(5) = 0
 l3_value.n(4) = 0
 l3_value.n(5) = 0
 l3_value.line_no(2) = 0
 l3_value.para(2) = "0"
End If
ElseIf ty = 1 Or ty = 5 Then
l3_value.para(2) = l3_value.para(1)
l3_value.para(1) = para3
If ty = 5 And l3_value.para(1) = l3_value.para(2) Then
 l3_value.poi(3) = l3_value.poi(5)
 l3_value.n(3) = l3_value.n(5)
 l3_value.poi(4) = 0
 l3_value.poi(5) = 0
 l3_value.n(4) = 0
 l3_value.n(5) = 0
 l3_value.line_no(2) = 0
 l3_value.para(2) = "0"
End If
ElseIf ty = 4 Then
'l3_value.para(1) = l3_value.para(1)
l3_value.para(2) = add_string(l3_value.para(1), para3, True, False)
ElseIf ty = 6 Then
l3_value.para(2) = add_string(l3_value.para(1), para3, True, False)
l3_value.para(1) = para3
ElseIf ty = 7 Then
l3_value.para(1) = add_string(l3_value.para(1), para3, True, False)
l3_value.para(2) = para3
ElseIf ty = 8 Then
l3_value.para(2) = l3_value.para(1)
l3_value.para(1) = add_string(l3_value.para(1), para3, True, False)
End If
If l3_value.para(1) = "0" Then
   l3_value.line_no(1) = 0
   l3_value.poi(2) = 0
   l3_value.poi(3) = 0
   l3_value.n(2) = 0
   l3_value.n(3) = 0
End If
If l3_value.para(2) = "0" Then
   l3_value.line_no(2) = 0
   l3_value.poi(4) = 0
   l3_value.poi(5) = 0
   l3_value.n(4) = 0
   l3_value.n(5) = 0
End If
Call arrange_four_point(l3_value.poi(0), l3_value.poi(1), _
      l3_value.poi(2), l3_value.poi(3), l3_value.n(0), _
       l3_value.n(1), l3_value.n(2), l3_value.n(3), l3_value.line_no(0), _
        l3_value.line_no(1), l3_value.poi(0), l3_value.poi(1), l3_value.poi(2), _
          l3_value.poi(3), 0, 0, l3_value.n(0), l3_value.n(1), l3_value.n(2), _
           l3_value.n(3), 0, 0, l3_value.line_no(0), l3_value.line_no(1), 0, ty, tc_data, is_no_initial)
   If is_no_initial = 1 And n1% = 0 Then
    Call add_record_to_record(tc_data, c_data)
   End If
If ty = 2 Then
 l3_value.para(0) = add_string(l3_value.para(0), l3_value.para(1), True, False)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
ElseIf ty = 0 Or ty = 3 Then
'l3_value.para(0) = l3_value.para(0)
'l3_value.para(1) = l3_value.para(1)
If ty = 3 And l3_value.para(0) = l3_value.para(1) Then
 l3_value.poi(1) = l3_value.poi(3)
 l3_value.n(1) = l3_value.n(3)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
End If
ElseIf ty = 1 Or ty = 5 Then
Call exchange_two_string(l3_value.para(0), l3_value.para(1))
If ty = 5 And l3_value.para(0) = l3_value.para(1) Then
 l3_value.poi(1) = l3_value.poi(3)
 l3_value.n(1) = l3_value.n(3)
 l3_value.poi(2) = 0
 l3_value.poi(3) = 0
 l3_value.n(2) = 0
 l3_value.n(3) = 0
 l3_value.line_no(1) = 0
 l3_value.para(1) = "0"
End If
ElseIf ty = 4 Then
'l3_value.para(0) = l3_value.para(0)
l3_value.para(1) = add_string(l3_value.para(0), l3_value.para(1), True, False)
ElseIf ty = 6 Then
Call exchange_two_string(l3_value.para(0), l3_value.para(1))
l3_value.para(1) = add_string(l3_value.para(0), l3_value.para(1), True, False)
ElseIf ty = 7 Then
'l3_value.para(1) = l3_value.para(1)
l3_value.para(0) = add_string(l3_value.para(0), l3_value.para(1), True, False)
ElseIf ty = 8 Then
Call exchange_two_string(l3_value.para(0), l3_value.para(1))
l3_value.para(0) = add_string(l3_value.para(0), l3_value.para(1), True, False)
End If
If l3_value.para(0) = "0" Then
   l3_value.line_no(0) = 0
   l3_value.poi(0) = 0
   l3_value.poi(1) = 0
   l3_value.n(0) = 0
   l3_value.n(1) = 0
End If
If l3_value.para(1) = "0" Then
   l3_value.line_no(1) = 0
   l3_value.poi(2) = 0
   l3_value.poi(3) = 0
   l3_value.n(2) = 0
   l3_value.n(3) = 0
End If
l3_value.value = v
If n1% = -5000 Then
   GoTo is_three_line_value_mark5
End If
'***********************************
'设置首项系数
If l3_value.para(1) = "0" Then
 l3_value.para(1) = l3_value.para(2)
 l3_value.poi(2) = l3_value.poi(4)
 l3_value.poi(3) = l3_value.poi(5)
 l3_value.n(2) = l3_value.n(4)
 l3_value.n(3) = l3_value.n(5)
 l3_value.line_no(1) = l3_value.line_no(2)
 l3_value.para(2) = "0"
 l3_value.poi(4) = 0
 l3_value.poi(5) = 0
 l3_value.n(4) = 0
 l3_value.n(5) = 0
 l3_value.line_no(2) = 0
End If
If l3_value.para(0) = "0" Then
 l3_value.para(0) = l3_value.para(1)
 l3_value.poi(0) = l3_value.poi(2)
 l3_value.poi(1) = l3_value.poi(3)
 l3_value.n(0) = l3_value.n(2)
 l3_value.n(1) = l3_value.n(3)
 l3_value.line_no(0) = l3_value.line_no(1)
 '****
 l3_value.para(1) = l3_value.para(2)
 l3_value.poi(2) = l3_value.poi(4)
 l3_value.poi(3) = l3_value.poi(5)
 l3_value.n(2) = l3_value.n(4)
 l3_value.n(3) = l3_value.n(5)
 l3_value.line_no(1) = l3_value.line_no(2)
 '***
 l3_value.para(2) = "0"
 l3_value.poi(4) = 0
 l3_value.poi(5) = 0
 l3_value.n(4) = 0
 l3_value.n(5) = 0
 l3_value.line_no(2) = 0
End If
If l3_value.para(0) = "0" Then
Exit Function
Else
ts = ""
Call simple_multi_string0(l3_value.para(0), l3_value.para(1), _
      l3_value.para(2), "0", ts, True)
End If
If l3_value.value <> "" Then
l3_value.value = divide_string(l3_value.value, ts, True, False)
End If
'***********************************
'排除相容
'If re.condition_data.condition_no > 0 Then
If is_line_value(l3_value.poi(0), l3_value.poi(1), l3_value.n(0), _
   l3_value.n(1), l3_value.line_no(0), "", tn%, -1000, _
     0, 0, 0, line_value_data0) = 1 Then
 Call add_conditions_to_record(line_value_, tn%, 0, 0, c_data)
   l3_value.value = minus_string(l3_value.value, _
      time_string(l3_value.para(0), line_value(tn%).data(0).data0.value, False, False), True, False)
 is_three_line_value = is_two_line_value(l3_value.poi(2), l3_value.poi(3), _
      l3_value.poi(4), l3_value.poi(5), l3_value.n(2), l3_value.n(3), _
       l3_value.n(4), l3_value.n(5), l3_value.line_no(1), l3_value.line_no(2), _
        l3_value.para(1), l3_value.para(2), l3_value.value, n%, -1000, _
         0, 0, 0, two_line_value_data0, con_ty, c_data)
          l3_value.poi(0) = two_line_value_data0.poi(0)
           l3_value.poi(1) = two_line_value_data0.poi(1)
          l3_value.poi(2) = two_line_value_data0.poi(2)
           l3_value.poi(3) = two_line_value_data0.poi(3)
          l3_value.n(0) = two_line_value_data0.n(0)
           l3_value.n(1) = two_line_value_data0.n(1)
          l3_value.n(2) = two_line_value_data0.n(2)
           l3_value.n(3) = two_line_value_data0.n(3)
          l3_value.line_no(0) = two_line_value_data0.line_no(0)
           l3_value.line_no(1) = two_line_value_data0.line_no(1)
          l3_value.para(0) = two_line_value_data0.para(0)
           l3_value.para(1) = two_line_value_data0.para(1)
          l3_value.value = two_line_value_data0.value
          l3_value.para(2) = "0"
           l3_value.poi(4) = 0
            l3_value.poi(5) = 0
           l3_value.n(4) = 0
            l3_value.n(5) = 0
           l3_value.line_no(2) = 0
        Exit Function
ElseIf is_line_value(l3_value.poi(2), l3_value.poi(3), l3_value.n(2), _
            l3_value.n(3), l3_value.line_no(1), "", tn%, -1000, _
             0, 0, 0, line_value_data0) = 1 Then
 Call add_conditions_to_record(line_value_, tn%, 0, 0, c_data)
    l3_value.value = minus_string(l3_value.value, _
      time_string(l3_value.para(1), line_value(tn%).data(0).data0.value, _
         False, False), True, False)
     is_three_line_value = is_two_line_value(l3_value.poi(0), l3_value.poi(1), _
        l3_value.poi(4), l3_value.poi(5), l3_value.n(0), l3_value.n(1), _
         l3_value.n(4), l3_value.n(5), l3_value.line_no(0), l3_value.line_no(2), _
          l3_value.para(0), l3_value.para(2), l3_value.value, n%, -1000, _
           0, 0, 0, two_line_value_data0, con_ty, c_data)
          l3_value.poi(0) = two_line_value_data0.poi(0)
           l3_value.poi(1) = two_line_value_data0.poi(1)
          l3_value.poi(2) = two_line_value_data0.poi(2)
           l3_value.poi(3) = two_line_value_data0.poi(3)
          l3_value.n(0) = two_line_value_data0.n(0)
           l3_value.n(1) = two_line_value_data0.n(1)
          l3_value.n(2) = two_line_value_data0.n(2)
           l3_value.n(3) = two_line_value_data0.n(3)
          l3_value.line_no(0) = two_line_value_data0.line_no(0)
           l3_value.line_no(1) = two_line_value_data0.line_no(1)
          l3_value.para(0) = two_line_value_data0.para(0)
           l3_value.para(1) = two_line_value_data0.para(1)
          l3_value.value = two_line_value_data0.value
          l3_value.para(2) = "0"
           l3_value.poi(4) = 0
            l3_value.poi(5) = 0
           l3_value.n(4) = 0
            l3_value.n(5) = 0
           l3_value.line_no(2) = 0
        Exit Function
ElseIf is_line_value(l3_value.poi(4), l3_value.poi(5), l3_value.n(4), _
          l3_value.n(5), l3_value.line_no(2), "", tn%, -1000, _
            0, 0, 0, line_value_data0) = 1 Then
 Call add_conditions_to_record(line_value_, tn%, 0, 0, c_data)
    l3_value.value = minus_string(l3_value.value, _
       time_string(l3_value.para(2), line_value(tn%).data(0).data0.value, False, False), True, False)
 is_three_line_value = is_two_line_value(l3_value.poi(0), l3_value.poi(1), _
     l3_value.poi(2), l3_value.poi(3), l3_value.n(0), l3_value.n(1), _
      l3_value.n(2), l3_value.n(3), l3_value.line_no(0), l3_value.line_no(1), _
       l3_value.para(0), l3_value.para(1), l3_value.value, 0, -1000, 0, 0, 0, _
        two_line_value_data0, con_ty, c_data)
          l3_value.poi(0) = two_line_value_data0.poi(0)
           l3_value.poi(1) = two_line_value_data0.poi(1)
          l3_value.poi(2) = two_line_value_data0.poi(2)
           l3_value.poi(3) = two_line_value_data0.poi(3)
          l3_value.n(0) = two_line_value_data0.n(0)
           l3_value.n(1) = two_line_value_data0.n(1)
          l3_value.n(2) = two_line_value_data0.n(2)
           l3_value.n(3) = two_line_value_data0.n(3)
          l3_value.line_no(0) = two_line_value_data0.line_no(0)
           l3_value.line_no(1) = two_line_value_data0.line_no(1)
          l3_value.para(0) = two_line_value_data0.para(0)
           l3_value.para(1) = two_line_value_data0.para(1)
          l3_value.value = two_line_value_data0.value
          l3_value.para(2) = "0"
           l3_value.poi(4) = 0
            l3_value.poi(5) = 0
           l3_value.n(4) = 0
            l3_value.n(5) = 0
           l3_value.line_no(2) = 0
        Exit Function
End If

 '***************************
is_three_line_value_mark1:
con_ty = line3_value_
is_three_line_value_mark5:
If InStr(1, l3_value.para(0), "F", 0) > 0 Or _
     InStr(1, l3_value.para(1), "F", 0) > 0 Or _
      InStr(1, l3_value.para(2), "F", 0) > 0 Or _
        InStr(1, l3_value.value, "F", 0) > 0 Then
 If n1% <= -1000 Then
  is_three_line_value = 0
 Else
  is_three_line_value = 1
   n% = 0
 End If
  Exit Function
End If
search_for_three_line_value:
If search_for_line3_value(l3_value, 0, n%, 0) Then '5.7
 If minus_string(l3_value.para(0), line3_value(n%).data(0).data0.para(0), True, False) <> "0" Or _
     minus_string(l3_value.para(1), line3_value(n%).data(0).data0.para(1), True, False) <> "0" Or _
      minus_string(l3_value.para(2), line3_value(n%).data(0).data0.para(2), True, False) <> "0" Or _
       (minus_string(l3_value.value, line3_value(n%).data(0).data0.value, True, False) <> "0" And _
          l3_value.value <> "") Then
   If n1% = -1000 Then
    is_three_line_value = 0
     Exit Function
   Else
   Call add_conditions_to_record(line3_value_, n%, 0, 0, c_data)
   If solve_multi_varity_equations(l3_value.para(0), l3_value.para(1), _
    l3_value.para(2), "0", l3_value.value, line3_value(n%).data(0).data0.para(0), _
     line3_value(n%).data(0).data0.para(1), line3_value(n%).data(0).data0.para(2), "0", _
      line3_value(n%).data(0).data0.value, l3_value.para(0), l3_value.para(1), _
       "", l3_value.value) = False Then
      is_three_line_value = set_equation(minus_string("x", l3_value.value, True, False), 0, temp_record)
         Exit Function
    End If
    If l3_value.para(0) = "0" Then
     l3_value.poi(2) = 0
     l3_value.poi(3) = 0
    End If
    If l3_value.para(1) = "0" Then
     l3_value.poi(4) = 0
     l3_value.poi(5) = 0
    End If
    temp_record.record_data.data0.condition_data = c_data
    Call set_two_line_value(l3_value.poi(2), l3_value.poi(3), _
        l3_value.poi(4), l3_value.poi(5), l3_value.n(2), _
         l3_value.n(3), l3_value.n(4), l3_value.n(5), _
          l3_value.line_no(1), l3_value.line_no(2), l3_value.para(0), _
           l3_value.para(1), l3_value.value, temp_record, 0, 0)
             n% = 0
    is_three_line_value = 1
'     Call set_level(re)
     Exit Function
   End If
 Else
     If set_or_prove = 2 Then
      If line3_value(n%).data(0).record.data1.is_proved = 1 Then '
      is_three_line_value = 1
      End If
     Else
      is_three_line_value = 1
     End If
'      Call set_level(re)
      Exit Function
 End If
Else
 If n1% = -5000 Then
    GoTo is_three_line_value_out
 End If
 t_n(0) = 0
 If is_line_value(l3_value.poi(0), l3_value.poi(1), l3_value.n(0), l3_value.n(1), _
     l3_value.line_no(0), "", t_n(0), -1000, 0, 0, 0, l_value) = 1 Then
      Call add_conditions_to_record(line_value_, t_n(0), 0, 0, c_data)
       If is_l3_value_from_l_l2_value(l3_value, 0, -1000, l_value, c_data, con_ty) Then
            is_three_line_value = 1
            If n1% = -5000 Then
              GoTo is_three_line_value_out
            Else
                        n% = 0
            End If
       End If
             Exit Function
 ElseIf is_line_value(l3_value.poi(2), l3_value.poi(3), l3_value.n(2), l3_value.n(3), _
     l3_value.line_no(1), "", t_n(0), -1000, 0, 0, 0, l_value) = 1 Then
      Call add_conditions_to_record(line_value_, t_n(0), 0, 0, c_data)
       If is_l3_value_from_l_l2_value(l3_value, 1, -1000, l_value, c_data, con_ty) Then
             is_three_line_value = 1
            If n1% = -5000 Then
              GoTo is_three_line_value_out
            Else
                        n% = 0
            End If
       End If
             Exit Function
 ElseIf is_line_value(l3_value.poi(4), l3_value.poi(5), l3_value.n(4), l3_value.n(5), _
     l3_value.line_no(2), "", t_n(0), -1000, 0, 0, 0, l_value) = 1 Then
      Call add_conditions_to_record(line_value_, t_n(0), 0, 0, c_data)
       If is_l3_value_from_l_l2_value(l3_value, 2, -1000, l_value, c_data, con_ty) Then
            is_three_line_value = 1
            If n1% = -5000 Then
              GoTo is_three_line_value_out
            Else
                        n% = 0
            End If
       End If
             Exit Function
 ElseIf is_two_line_value_(l3_value.poi(0), l3_value.poi(1), l3_value.poi(2), l3_value.poi(3), _
         l3_value.n(0), l3_value.n(1), l3_value.n(2), l3_value.n(3), l3_value.line_no(0), l3_value.line_no(1), _
           c_data, l2_value) = 1 Then
          If is_l3_value_from_l2_l_value(l3_value, l2_value, 0, -1000, c_data, con_ty) Then
            is_three_line_value = 1
            If n1% = -5000 Then
              GoTo is_three_line_value_out
            Else
                        n% = 0
            End If
           End If
           Exit Function
 ElseIf is_two_line_value_(l3_value.poi(2), l3_value.poi(3), l3_value.poi(4), l3_value.poi(5), _
         l3_value.n(2), l3_value.n(3), l3_value.n(4), l3_value.n(5), l3_value.line_no(1), l3_value.line_no(2), _
          c_data, l2_value) = 1 Then
           If is_l3_value_from_l2_l_value(l3_value, l2_value, 1, -1000, c_data, con_ty) Then
             is_three_line_value = 1
            If n1% = -5000 Then
             GoTo is_three_line_value_out
             Else
                        n% = 0
           End If
           End If
           Exit Function
 ElseIf is_two_line_value_(l3_value.poi(0), l3_value.poi(1), l3_value.poi(4), l3_value.poi(5), _
         l3_value.n(0), l3_value.n(1), l3_value.n(4), l3_value.n(5), l3_value.line_no(0), l3_value.line_no(2), _
          c_data, l2_value) = 1 Then
           If is_l3_value_from_l2_l_value(l3_value, l2_value, 2, -1000, c_data, con_ty) Then
             is_three_line_value = 1
            If n1% = -5000 Then
             GoTo is_three_line_value_out
            Else
                        n% = 0
            End If
          End If
          Exit Function
 Else
    depend_no = depend_no + 1
     If depend_no = 2 Then
            is_three_line_value = 1
         '    Call set_level(re)
      Exit Function
     End If
 End If
End If
is_three_line_value_out:
If n1% = -1000 Then
 n% = 0
  Exit Function
End If
n1% = n%
 Call search_for_line3_value(l3_value, 1, n2%, 1) '5.7
 Call search_for_line3_value(l3_value, 2, n3%, 1)
 Call search_for_line3_value(l3_value, 3, n4%, 1)
 Call search_for_line3_value(l3_value, 4, n5%, 1)
 Call search_for_line3_value(l3_value, 5, n6%, 1)
End Function


Public Function is_area_of_triangle(tri As Integer, n%) As Boolean
Dim i%, j%
Dim ln(2) As Integer
Dim A(2) As Integer
Dim l(2) As Integer
Dim vl(2) As Integer
Dim temp_record As record_data_type
Dim value As String
Dim value1 As String
Dim area_A As area_of_element_data_type
Dim insert_no%
area_A.element.ty = triangle_
area_A.element.no = tri
If search_for_area_element(area_A, 1, n%, 0) Then
'For i% = 1 To last_area_of_triangle
 'If area_of_triangle(i%).triangle = tri Then 'And Area_of_triangle(i%).data(0).value <> "" Then
  'n% = i%
   is_area_of_triangle = True
    Exit Function
Else
 insert_no% = n%
End If
'Next i%
    l(0) = line_number0(triangle(tri).data(0).poi(1), triangle(tri).data(0).poi(2), 0, 0)
    l(1) = line_number0(triangle(tri).data(0).poi(2), triangle(tri).data(0).poi(0), 0, 0)
    l(2) = line_number0(triangle(tri).data(0).poi(0), triangle(tri).data(0).poi(1), 0, 0)

If is_line_value(triangle(tri).data(0).poi(0), triangle(tri).data(0).poi(1), _
        0, 0, 0, "", ln(2), -1000, 0, 0, 0, line_value_data0) = 0 Then
    ln(2) = 0
End If
If is_line_value(triangle(tri).data(0).poi(2), triangle(tri).data(0).poi(1), _
        0, 0, 0, "", ln(0), -1000, 0, 0, 0, line_value_data0) = 0 Then
    ln(0) = 0
End If
If is_line_value(triangle(tri).data(0).poi(2), triangle(tri).data(0).poi(0), _
        0, 0, 0, "", ln(1), -1000, 0, 0, 0, line_value_data0) = 0 Then
    ln(1) = 0
End If
If is_angle_value(triangle(tri).data(0).angle(0), "", "", A(0), record_0.data0.condition_data) = False Then
    A(0) = 0
End If
If is_angle_value(triangle(tri).data(0).angle(1), "", "", A(1), record_0.data0.condition_data) = False Then
      A(1) = 0
End If
If is_angle_value(triangle(tri).data(0).angle(2), "", "", A(2), record_0.data0.condition_data) = False Then
      A(2) = 0
End If
For i% = 0 To 2
 For j% = 1 To m_lin(l(i%)).data(0).data0.in_point(0)
  If is_dverti(line_number0(triangle(tri).data(0).poi(i%), _
          m_lin(l(i%)).data(0).data0.in_point(j%), 0, 0), _
            l(i%), 0, -1000, 0, 0, 0, 0) Then
   If is_line_value(triangle(tri).data(0).poi(i%), m_lin(l(i%)).data(0).data0.in_point(j%), _
      0, 0, 0, "", vl(i%), -1000, 0, 0, 0, line_value_data0) = 0 Then
        vl(i%) = 0
   End If
    GoTo is_area_of_triangle_mark0
  End If
 Next j%
is_area_of_triangle_mark0:
Next i%
'For i% = 0 To 2
 'If ln(i%) > 0 And vl(i%) > 0 Then
  'temp_record.condition_data.condition_no = 2
   'temp_record.condition_data.condition(1).ty = line_value_
   'temp_record.condition_data.condition(2).ty = line_value_
   'temp_record.condition_data.condition(1).no= ln(i%)
   'temp_record.condition_data.condition(2).no= vl(i%)
    'is_area_of_triangle = True
    'value = divide_string(time_string(line_value(ln(i%)).data(0).value, _
         line_value(vl(i%)).data(0).value), "2")
     'Call set_area_of_triangle(tri, value, temp_record, n%)
     ' Exit Function
 'End If
'Next i%
'For i% = 0 To 2
 'If ln(i%) > 0 And ln((i% + 2) Mod 3) > 0 And _
    A((i% + 1) Mod 3) > 0 Then
  ' temp_record.condition_data.condition_no = 3
   'temp_record.condition_data.condition(1).ty = line_value_
   'temp_record.condition_data.condition(2).ty = line_value_
   'temp_record.condition_data.condition(3).ty = angle_value_
   'temp_record.condition_data.condition(1).no= ln(i%)
   'temp_record.condition_data.condition(2).no= ln((i% + 2) Mod 3)
   'temp_record.condition_data.condition(3).no= A((i% + 1) Mod 3)
   'value = divide_string(time_string( _
      sin_(angle_value(A((i% + 1) Mod 3)).data(0).value, 0), _
      time_string(line_value(ln(i%)).data(0).value, _
        line_value(ln((i% + 2) Mod 3)).data(0).value)), "2")
   '  is_area_of_triangle = True
    ' Call set_area_of_triangle(tri, value, temp_record, n%)
     ' Exit Function
 'End If
'Next i%
'If ln(0) > 0 And ln(1) > 0 And ln(2) > 0 Then
 '  temp_record.condition_data.condition_no = 3
  ' temp_record.condition_data.condition(1).ty = line_value_
   'temp_record.condition_data.condition(2).ty = line_value_
   't'emp_record.data.condition_data.condition(3).ty = line_value_
   'temp_record.condition_data.condition(1).no= ln(0)
   'temp_record.condition_data.condition(2).no= ln(1)
   'temp_record.condition_data.condition(3).no= ln(2)
   'value = add_string(line_value(ln(0)).data(0).value, line_value(ln(1)).data(0).value)
   'value = add_string(value, line_value(ln(2)).data(0).value)
   'value1 = divide_string(value, "2")
   'value = value1
   'For i% = 0 To 2
   '  value1 = time_string(value1, _
   '      minus_string(value, line_value(ln(i%)).data(0).value))
   'Next i%
   'value1 = sqr_string(value1)
   '   is_area_of_triangle = True
   '  Call set_area_of_triangle(tri, value1, temp_record, n%)
   '   Exit Function
'End If
is_area_of_triangle = False
n% = insert_no%
End Function

Public Function is_area_of_circle(c%, n%) As Boolean
Dim i%, tn%
 For i% = 1 To last_conditions.last_cond(1).area_of_circle_no
  If area_of_circle(i%).data(0).circ = c% Then
   is_area_of_circle = True
    Exit Function
  End If
 Next i%
If is_line_value(m_Circ(c%).data(0).data0.center, m_Circ(i%).data(0).data0.in_point(1), _
      0, 0, 0, "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
If last_conditions.last_cond(1).area_of_circle_no Mod 10 = 0 Then
ReDim Preserve area_of_circle(last_conditions.last_cond(1).area_of_circle_no + 10) As area_of_circle_type
End If
last_conditions.last_cond(1).area_of_circle_no = last_conditions.last_cond(1).area_of_circle_no + 1
 n% = last_conditions.last_cond(1).area_of_circle_no
area_of_circle(n%).data(0) = area_of_circle_data_0
 area_of_circle(n%).data(0).circ = c%
  area_of_circle(n%).data(0).value = _
     time_string(line_value(tn%).data(0).data0.value, _
       line_value(tn%).data(0).data0.value, True, False)
  If InStr(1, area_of_circle(n%).data(0).value, "+", 0) = 0 And _
       InStr(1, area_of_circle(n%).data(0).value, "-", 0) = 0 Then
        area_of_circle(n%).data(0).value = area_of_circle(n%).data(0).value + "\" 'LoadResString_(1456,"")
  Else
        area_of_circle(n%).data(0).value = "(" + area_of_circle(n%).data(0).value + ")" + LoadResString_(1455, "")
  End If
  area_of_circle(n%).data(0).record.data0.condition_data.condition_no = 1
   area_of_circle(n%).data(0).record.data0.condition_data.condition(1).ty = line_value_
    area_of_circle(n%).data(0).record.data0.condition_data.condition(1).no = tn%
     'area_of_circle(n%).record.other_no = n%
If run_type = 10 Then
'If new_result_from_add = False Then
If m_Circ(c%).data(0).data0.center <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no And _
     m_Circ(c%).data(0).data0.in_point(1) <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no Then
last_conditions_for_aid(last_conditions_for_aid_no).new_result_from_add = False
End If
'End If
End If
     is_area_of_circle = True
      Exit Function
End If
is_area_of_circle = False
End Function

Public Function is_area_of_fan(ByVal p1%, ByVal p2%, ByVal p3%, n%) As Boolean
Dim i%
Dim tn(1) As Integer
Dim temp_record As record_data_type
Dim v As String
Dim temp_record1 As record_data_type
If p1% > p3% Then
Call exchange_two_integer(p1%, p3%)
End If
For i% = 1 To last_conditions.last_cond(1).area_of_fan_no
 If Area_of_fan(i%).data(0).poi(0) = p1% And _
      Area_of_fan(i%).data(0).poi(1) = p2% And _
       Area_of_fan(i%).data(0).poi(2) = p3% Then
        is_area_of_fan = True
         Exit Function
 End If
Next i%
v = ""
If is_angle_value(Abs(angle_number(p1%, p2%, p3%, 0, 0)), v, "", tn(0), temp_record1.data0.condition_data) Then
 temp_record = temp_record1
   If is_line_value(p1%, p2%, 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
  temp_record.data0.condition_data.condition(2).ty = line_value_
  temp_record.data0.condition_data.condition(2).no = tn(1)
   ElseIf is_line_value(p2%, p3%, 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
  temp_record.data0.condition_data.condition(2).ty = line_value_
  temp_record.data0.condition_data.condition(2).no = tn(1)
   Else
    is_area_of_fan = False
     Exit Function
   End If
   temp_record.data0.condition_data.condition_no = 2
 If last_conditions.last_cond(1).area_of_fan_no Mod 10 = 0 Then
 ReDim Preserve Area_of_fan(last_conditions.last_cond(1).area_of_fan_no + 10) As area_of_fan_type
 End If
 last_conditions.last_cond(1).area_of_fan_no = last_conditions.last_cond(1).area_of_fan_no + 1
  n% = last_conditions.last_cond(1).area_of_fan_no
 Area_of_fan(n%).data(0) = area_of_fan_data_0
  Area_of_fan(n%).data(0).poi(0) = p1%
  Area_of_fan(n%).data(0).poi(1) = p2%
  Area_of_fan(n%).data(0).poi(2) = p3%
  Area_of_fan(n%).data(0).record = temp_record
  'Area_of_fan(n%).record.other_no = n%
  Area_of_fan(n%).data(0).value = calcutete_area_of_fan(angle(angle3_value(tn(0)).data(0).data0.angle(0)).data(0).value, _
       line_value(tn(1)).data(0).data0.value)
   If run_type = 10 Then
  'If new_result_from_add = False Then
  If p1% <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no And p2% <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no And _
      p3% <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no Then
  last_conditions_for_aid(last_conditions_for_aid_no).new_result_from_add = True
  End If
  'End If
  End If
   is_area_of_fan = True
     Exit Function
  End If
  is_area_of_fan = False

End Function

Public Function is_area_of_polygon(ByVal p1%, ByVal p2%, _
        ByVal p3%, ByVal p4%, n%, value As String, poly4_no%) As Byte
Dim i%, j%
Dim t_p%
Dim tp(3) As Integer
Dim tn(3) As Integer
Dim l(2) As Integer
Dim con_ty As Byte
Dim v1 As String
Dim area_ele As area_of_element_data_type
Dim ts$
Dim triA(1) As Integer
Dim insert_no%
'Dim area_p As polygon4_data_type
Dim temp_record As total_record_type
Dim temp_record1 As record_data_type
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, 0)
If poly4_no% = 0 Then
'If is_polygon4(p1%, p2%, p3%, p4%, poly4_data, tn%, 0) = False Then
 Exit Function
End If
Dpolygon4(poly4_no%).data(0).area_value = value
area_ele.element.no = poly4_no%
area_ele.element.ty = polygon_
area_ele.value = value
area_ele.value_ = value
If search_for_area_element(area_ele, 1, n%, 0) Then
 is_area_of_polygon = 1
  Exit Function
Else
insert_no% = n%
End If
'****************
is_area_of_polygon = is_parallelogram0(poly4_no%, tn(0), -1000, con_ty)
  If is_area_of_polygon > 0 Then
 If con_ty = parallelogram_ Then
  temp_record.record_data.data0.condition_data.condition_no = 1
   temp_record.record_data.data0.condition_data.condition(1).ty = parallelogram_
    temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
  is_area_of_polygon = is_area_of_parallelogram(tn(0), temp_record.record_data, n%)
     Exit Function
 ElseIf con_ty = epolygon_ Then
 temp_record.record_data.data0.condition_data.condition(1).ty = long_squre_
  temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
  If is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(1), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark0
  ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(2), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark0
  ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(1), Dpolygon4(poly4_no%).data(0).poi(2), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark0
  ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark0
  Else
   GoTo is_area_of_polygon_mark1
  End If
is_area_of_polygon_mark0:
  temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
   temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
    temp_record.record_data.data0.condition_data.condition_no = 2
    n% = 0
is_area_of_polygon = set_area_of_polygon0(poly4_no%, time_string( _
   line_value(tn(1)).data(0).data0.value, line_value(tn(1)).data(0).data0.value, True, False), _
    temp_record, n%, 0)
  '  is_area_of_polygon = True
     Exit Function
 ElseIf con_ty = long_squre_ Then
  temp_record.record_data.data0.condition_data.condition(1).ty = long_squre_
   temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
  If is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(1), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark2
  ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(2), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark2
  Else
   GoTo is_area_of_polygon_mark1
  End If
is_area_of_polygon_mark2:
temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
 temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
  If is_line_value(Dpolygon4(poly4_no%).data(0).poi(1), Dpolygon4(poly4_no%).data(0).poi(2), 0, 0, 0, "", tn(2), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark3
  ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(2), -1000, 0, 0, 0, _
         line_value_data0) = 1 Then
   GoTo is_area_of_polygon_mark3
  Else
   GoTo is_area_of_polygon_mark1
  End If
is_area_of_polygon_mark3:
temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
 temp_record.record_data.data0.condition_data.condition(3).no = tn(2)
  temp_record.record_data.data0.condition_data.condition_no = 3
 n% = 0
 is_area_of_polygon = set_area_of_polygon0(poly4_no%, time_string( _
   line_value(tn(1)).data(0).data0.value, line_value(tn(2)).data(0).data0.value, True, False), _
    temp_record, n%, 0)
    'is_area_of_polygon = True
     Exit Function
  End If
ElseIf is_rhombus(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(1), _
                   Dpolygon4(poly4_no%).data(0).poi(2), Dpolygon4(poly4_no%).data(0).poi(3), _
                    tn(0), 0, -1000, con_ty) Then
'If tn(0) > 0 Then
 If con_ty = rhombus_ Then
  temp_record.record_data.data0.condition_data.condition(1).ty = rhombus_
  temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
 If is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(1), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
 GoTo is_area_of_polygon_mark4
 ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(1), Dpolygon4(poly4_no%).data(0).poi(2), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
 GoTo is_area_of_polygon_mark4
 ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(2), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
 GoTo is_area_of_polygon_mark4
 ElseIf is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
 GoTo is_area_of_polygon_mark4
 Else
 GoTo is_area_of_polygon_mark5
 End If
is_area_of_polygon_mark4:
temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
temp_record.record_data.data0.condition_data.condition_no = 2
ts$ = ""
 If is_angle_value(Dpolygon4(poly4_no%).data(0).angle(0), ts$, "", tn(2), temp_record1.data0.condition_data) Then
  GoTo is_area_of_polygon_mark6
 Else
  ts$ = ""
  If is_angle_value(Dpolygon4(poly4_no%).data(0).angle(1), ts$, "", tn(2), temp_record1.data0.condition_data) Then
   GoTo is_area_of_polygon_mark6
  Else
   ts$ = ""
   If is_angle_value(Dpolygon4(poly4_no%).data(0).angle(2), ts$, "", tn(2), temp_record1.data0.condition_data) Then
    GoTo is_area_of_polygon_mark6
   Else
    ts$ = ""
    If is_angle_value(Dpolygon4(poly4_no%).data(0).angle(3), ts$, "", tn(2), temp_record1.data0.condition_data) Then
     GoTo is_area_of_polygon_mark6
    End If
   End If
  End If
End If
End If
  GoTo is_area_of_polygon_mark5
is_area_of_polygon_mark6:
Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
 ts$ = sin_(ts$, 0)
If InStr(1, ts$, "F", 0) = 0 Then
 is_area_of_polygon = set_area_of_polygon0(poly4_no%, time_string(time_string( _
   line_value(tn(1)).data(0).data0.value, line_value(tn(1)).data(0).data0.value, False, False), _
         ts$, True, False), _
      temp_record, n%, 0)
    'is_area_of_polygon = True
     Exit Function
End If
'End If
is_area_of_polygon_mark5:
If is_line_value(Dpolygon4(poly4_no%).data(0).poi(0), Dpolygon4(poly4_no%).data(0).poi(2), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
       line_value_data0) = 1 Then
 If is_line_value(Dpolygon4(poly4_no%).data(0).poi(1), Dpolygon4(poly4_no%).data(0).poi(3), 0, 0, 0, "", tn(2), -1000, 0, 0, 0, _
       line_value_data0) = 1 Then
  temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
   temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
  temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
   temp_record.record_data.data0.condition_data.condition(3).no = tn(2)
    temp_record.record_data.data0.condition_data.condition_no = 3
 n% = 0
is_area_of_polygon = set_area_of_polygon0(poly4_no%, divide_string(time_string( _
   line_value(tn(1)).data(0).data0.value, line_value(tn(2)).data(0).data0.value, False, False), _
          "2", True, False), _
    temp_record, n%, 0)
    'is_area_of_polygon = True
     Exit Function
    
 End If
End If
Else
If is_equal_side_tixing0(poly4_no%, tn(0), tp(0), tp(1), tp(2), tp(3), con_ty) Then
temp_record.record_data.data0.condition_data.condition(1).ty = con_ty
temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
GoTo is_area_of_polygon_mark9
Else
GoTo is_area_of_polygon_mark1
End If
is_area_of_polygon_mark9:
If is_line_value(tp(0), tp(1), 0, 0, 0, "", tn(1), -1000, 0, 0, 0, _
        line_value_data0) = 1 And _
  is_line_value(tp(2), tp(3), 0, 0, 0, "", tn(2), -1000, 0, 0, 0, _
        line_value_data0) = 1 Then
v1 = add_string(line_value(tn(1)).data(0).data0.value, _
        line_value(tn(2)).data(0).data0.value, True, False)
temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
temp_record.record_data.data0.condition_data.condition(3).no = tn(2)
temp_record.record_data.data0.condition_data.condition_no = 3
 GoTo is_area_of_polygon_mark8
ElseIf is_two_line_value(tp(0), tp(1), tp(2), tp(3), _
    0, 0, 0, 0, 0, 0, "1", "1", "", tn(1), -1000, 0, 0, 0, _
     two_line_value_data0, 0, temp_record1.data0.condition_data) = 1 Then
v1 = two_line_value(tn(1)).data(0).data0.value
temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
 temp_record.record_data.data0.condition_data.condition_no = 2
GoTo is_area_of_polygon_mark8
Else
GoTo is_area_of_polygon_mark1
End If
is_area_of_polygon_mark8:
l(0) = line_number0(tp(0), tp(1), 0, 0)
 l(1) = line_number0(tp(2), tp(3), 0, 0)
  For i% = 1 To m_lin(l(0)).data(0).data0.in_point(0)
   For j% = 1 To m_lin(l(1)).data(0).data0.in_point(0)
    l(2) = line_number0(m_lin(l(0)).data(0).data0.in_point(i%), m_lin(l(1)).data(0).data0.in_point(j%), 0, 0)
     If is_dverti(l(2), l(0), tn(3), -1000, 0, 0, 0, 0) Then
      GoTo is_area_of_polygon_mark7
     ElseIf is_dverti(l(2), l(1), tn(3), -1000, 0, 0, 0, 0) Then
      GoTo is_area_of_polygon_mark7
    End If
    Next j%
    Next i%
    GoTo is_area_of_polygon_mark1
is_area_of_polygon_mark7:
 Call add_conditions_to_record(verti_, tn(3), 0, 0, temp_record.record_data.data0.condition_data)
     If is_line_value(m_lin(l(0)).data(0).data0.in_point(i%), _
            m_lin(l(1)).data(0).data0.in_point(j%), 0, 0, 0, "", tn(0), -1000, _
             0, 0, 0, line_value_data0) = 1 Then
 Call add_conditions_to_record(line_value_, tn(0), 0, 0, temp_record.record_data.data0.condition_data)
 n% = 0
 is_area_of_polygon = set_area_of_polygon0(poly4_no%, _
  time_string(line_value(tn(0)).data(0).data0.value, divide_string(v1, "2", False, False), _
       True, False), temp_record, n%, 0)
     'is_area_of_polygon = True
     Exit Function
    End If
'
 '  Next j%
'  Next i%
End If
'Else
'GoTo is_area_of_polygon_mark
'End If
'End If
is_area_of_polygon_mark1:
If Dpolygon4(poly4_no%).data(0).triAngle1(0) > 0 And Dpolygon4(poly4_no%).data(0).triAngle1(1) > 0 Then
If triangle(Dpolygon4(poly4_no%).data(0).triAngle1(0)).data(0).area_no > 0 Then
 If triangle(Dpolygon4(poly4_no%).data(0).triAngle1(1)).data(0).area_no Then
 temp_record.record_data.data0.condition_data.condition_no = 2
  temp_record.record_data.data0.condition_data.condition(1).ty = area_of_element_
   temp_record.record_data.data0.condition_data.condition(2).ty = area_of_element_
    temp_record.record_data.data0.condition_data.condition(1).no = _
           triangle(Dpolygon4(poly4_no%).data(0).triAngle1(0)).data(0).area_no
     temp_record.record_data.data0.condition_data.condition(2).no = _
         triangle(Dpolygon4(poly4_no%).data(0).triAngle1(1)).data(0).area_no
n% = 0
is_area_of_polygon = _
  set_area_of_polygon0(poly4_no%, add_string(triangle(Dpolygon4(poly4_no%).data(0).triAngle1(0)).data(0).Area, _
     triangle(Dpolygon4(poly4_no%).data(0).triAngle1(1)).data(0).Area, True, False), temp_record, n%, 0)
      'is_area_of_polygon = True
       Exit Function
 End If
End If
End If
If Dpolygon4(poly4_no%).data(0).triAngle2(0) > 0 And Dpolygon4(poly4_no%).data(0).triAngle2(1) > 0 Then
If triangle(Dpolygon4(poly4_no%).data(0).triAngle2(0)).data(0).area_no > 0 Then
 If triangle(Dpolygon4(poly4_no%).data(0).triAngle2(1)).data(0).area_no Then
 temp_record.record_data.data0.condition_data.condition_no = 2
  temp_record.record_data.data0.condition_data.condition(1).ty = area_of_element_
   temp_record.record_data.data0.condition_data.condition(2).ty = area_of_element_
    temp_record.record_data.data0.condition_data.condition(1).no = _
      triangle(Dpolygon4(poly4_no%).data(0).triAngle2(0)).data(0).area_no
     temp_record.record_data.data0.condition_data.condition(2).no = _
       triangle(Dpolygon4(poly4_no%).data(0).triAngle2(1)).data(0).area_no
n% = 0
is_area_of_polygon = _
  set_area_of_polygon0(poly4_no%, add_string(triangle(Dpolygon4(poly4_no%).data(0).triAngle2(0)).data(0).Area, _
     triangle(Dpolygon4(poly4_no%).data(0).triAngle2(1)).data(0).Area, True, False), temp_record, n%, 0)
      'is_area_of_polygon = True
       Exit Function
 End If
End If
End If
If is_area_of_polygon = False Then
n% = insert_no%
End If
End Function

Public Function is_sides_length_of_circle(c%, n%) As Boolean
Dim i%, tn%
 For i% = 1 To last_conditions.last_cond(1).sides_length_of_circle_no
  If Sides_length_of_circle(i%).data(0).circ = c% Then
   n% = i%
   is_sides_length_of_circle = True
    Exit Function
  End If
 Next i%
If is_line_value(m_Circ(c%).data(0).data0.center, m_Circ(i%).data(0).data0.in_point(1), _
     0, 0, 0, "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
If last_conditions.last_cond(1).sides_length_of_circle_no Mod 10 = 0 Then
ReDim Preserve Sides_length_of_circle(last_conditions.last_cond(1).sides_length_of_circle_no + 10) As sides_length_of_circle_type
End If
last_conditions.last_cond(1).sides_length_of_circle_no = last_conditions.last_cond(1).sides_length_of_circle_no + 1
 n% = last_conditions.last_cond(1).sides_length_of_circle_no
 Sides_length_of_circle(n%).data(0) = sides_length_of_circle_data_0
 Sides_length_of_circle(n%).data(0).circ = c%
  Sides_length_of_circle(n%).data(0).value = time_string(PI, _
     time_string("2", _
       line_value(tn%).data(0).data0.value, False, False), True, False)
  Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition_no = 1
   Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition(1).ty = line_value_
    Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition(1).no = tn%
If run_type = 10 Then
'If new_result_from_add = False Then
If m_Circ(c%).data(0).data0.center <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no And _
      m_Circ(c%).data(0).data0.in_point(1) <= last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no Then
last_conditions_for_aid(last_conditions_for_aid_no).new_result_from_add = True
End If
End If
'End If
     is_sides_length_of_circle = True
      Exit Function
End If
is_sides_length_of_circle = False


End Function

Public Function is_sides_length_of_triangle(tri%, n%) As Boolean
Dim i%
Dim temp_record As total_record_type
Dim temp_record1 As record_data_type
Dim tn(2) As Integer
Dim s_l_A As sides_length_of_triangle_data_type
Dim insert_no%
s_l_A.triangle = tri%
triangle_data0 = triangle(tri%).data(0)
If search_for_sides_length_of_triangle(s_l_A, 1, n%, 0) Then
'For i% = 1 To last_sides_length_of_triangle
 'If Sides_length_of_triangle(i%).triangle = tri% Then
  'n% = i%
   is_sides_length_of_triangle = True
    Exit Function
Else
insert_no% = n%
End If
'Next i%
If is_line_value(triangle_data0.poi(0), triangle_data0.poi(1), _
     0, 0, 0, "", tn(0), -1000, 0, 0, 0, line_value_data0) = 1 Then
 If is_line_value(triangle_data0.poi(1), triangle_data0.poi(2), _
     0, 0, 0, "", tn(1), -1000, 0, 0, 0, line_value_data0) = 1 And _
      is_line_value(triangle_data0.poi(0), triangle_data0.poi(2), _
     0, 0, 0, "", tn(2), -1000, 0, 0, 0, line_value_data0) = 1 Then
     temp_record.record_data.data0.condition_data.condition_no = 3
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
     temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
     temp_record.record_data.data0.condition_data.condition(3).no = tn(2)
     n% = 0
     Call set_sides_length_of_triangle(tri%, add_string(add_string( _
        line_value(tn(0)).data(0).data0.value, line_value(tn(1)).data(0).data0.value, False, False), _
         line_value(tn(2)).data(0).data0.value, True, False), n%, temp_record, 0)
     is_sides_length_of_triangle = True
      Exit Function
 ElseIf is_two_line_value(triangle(tri%).data(0).poi(2), triangle(tri%).data(0).poi(0), _
    triangle(tri%).data(0).poi(2), triangle(tri%).data(0).poi(1), 0, 0, 0, 0, 0, 0, _
     "1", "1", "", tn(1), -1000, 0, 0, 0, _
       two_line_value_data0, 0, temp_record1.data0.condition_data) = 1 Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
     temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
     Call set_sides_length_of_triangle(tri%, add_string( _
       two_line_value(tn(1)).data(0).data0.value, line_value(tn(0)).data(0).data0.value, True, False), _
        n%, temp_record, 0)
     is_sides_length_of_triangle = True
      Exit Function
       
 End If
ElseIf is_line_value(triangle_data0.poi(2), triangle_data0.poi(1), _
      0, 0, 0, "", tn(0), -1000, 0, 0, 0, line_value_data0) = 1 Then
 If is_two_line_value(triangle_data0.poi(2), triangle_data0.poi(0), _
    triangle_data0.poi(0), triangle_data0.poi(1), 0, 0, 0, 0, 0, 0, _
     "1", "1", "", tn(1), -1000, 0, 0, 0, _
       two_line_value_data0, 0, temp_record1.data0.condition_data) = 1 Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
     temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
     Call set_sides_length_of_triangle(tri%, add_string( _
       two_line_value(tn(1)).data(0).data0.value, line_value(tn(0)).data(0).data0.value, True, False), _
        n%, temp_record, 0)
      Exit Function
 End If
ElseIf is_line_value(triangle_data0.poi(0), triangle_data0.poi(2), _
        0, 0, 0, "", tn(0), -1000, 0, 0, 0, line_value_data0) = 1 Then
 If is_two_line_value(triangle_data0.poi(1), triangle_data0.poi(0), _
    triangle_data0.poi(2), triangle_data0.poi(1), 0, 0, 0, 0, 0, 0, _
     "1", "1", "", tn(1), -1000, 0, 0, 0, _
       two_line_value_data0, 0, temp_record1.data0.condition_data) = 1 Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = tn(0)
     temp_record.record_data.data0.condition_data.condition(2).no = tn(1)
     Call set_sides_length_of_triangle(tri%, add_string( _
       two_line_value(tn(1)).data(0).data0.value, line_value(tn(0)).data(0).data0.value, True, False), _
        n%, temp_record, 0)
      Exit Function
 End If
End If
is_sides_length_of_triangle = False
n% = insert_no%
End Function

Public Function is_area_of_parallelogram(ByVal p%, re As record_data_type, no%) As Byte
'0 不成立,1成立,2推理完成
Dim i%
Dim tn(1) As Integer
Dim tn_(3) As Integer
Dim tp(1, 3) As Integer
Dim temp_record As total_record_type
 For i% = last_conditions.last_cond(0).area_of_element_no + 1 To _
          last_conditions.last_cond(1).area_of_element_no
  If area_of_element(i%).data(0).element.ty = polygon_ Then
  If area_of_element(i%).data(0).element.no = Dparallelogram(p%).data(0).polygon4_no Then
           is_area_of_parallelogram = 1
           Exit Function
  End If
  End If
Next i%
For i% = 0 To 3
temp_record.record_data = re
If is_line_value(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
      Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), 0, 0, 0, "", tn(0), _
        -1000, 0, 0, 0, line_value_data0) = 1 Then
 Call add_conditions_to_record(line_value_, _
  tn(0), 0, 0, temp_record.record_data.data0.condition_data)
 If is_line_value(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
    Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4), 0, 0, 0, "", tn(1), _
     -1000, 0, 0, 0, line_value_data0) = 1 Then
 ElseIf is_line_value(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
    Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4), 0, 0, 0, "", tn(1), _
     -1000, 0, 0, 0, line_value_data0) = 1 Then
 End If
 If find_verti_foot(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
  line_number0(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4), _
      Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4), 0, 0), _
        tp(0, 1), 0, tn_(0)) Then
  tp(0, 0) = Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%)
 End If
 If find_verti_foot(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), _
  line_number0(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4), _
      Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4), 0, 0), _
        tp(0, 1), 0, tn_(1)) Then
  tp(1, 1) = Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4)
 End If
 If find_verti_foot(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4), _
  line_number0(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
      Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), 0, 0), _
        tp(0, 2), 0, tn_(2)) Then
  tp(1, 2) = Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4)
 End If
 If find_verti_foot(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4), _
  line_number0(Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi(i%), _
      Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), 0, 0), _
       tp(0, 3), 0, tn_(3)) Then
   tp(1, 3) = Dpolygon4(Dparallelogram(p%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4)
 End If
 If is_line_value(tp(1, 0), tp(0, 0), 0, 0, 0, "", tn(1), -1000, _
        0, 0, 0, line_value_data0) = 1 Then
  Call add_conditions_to_record(line_value_, tn(1), 0, 0, _
    temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(verti_, tn_(0), 0, 0, _
    temp_record.record_data.data0.condition_data)
 no% = 0
    Call set_level(temp_record.record_data.data0.condition_data)
 is_area_of_parallelogram = set_area_of_polygon0(Dparallelogram(p%).data(0).polygon4_no, _
     time_string(line_value(tn(0)).data(0).data0.value, _
      line_value(tn(1)).data(0).data0.value, True, False), temp_record, no%, 0)
    Exit Function
 ElseIf is_line_value(tp(1, 1), tp(0, 1), 0, 0, 0, "", _
      tn(1), -1000, 0, 0, 0, line_value_data0) = 1 Then
  Call add_conditions_to_record(line_value_, tn(1), 0, 0, _
    temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(verti_, tn_(1), 0, 0, _
    temp_record.record_data.data0.condition_data)
no% = 0
    Call set_level(temp_record.record_data.data0.condition_data)
 is_area_of_parallelogram = set_area_of_polygon0(Dparallelogram(p%).data(0).polygon4_no, _
     time_string(line_value(tn(0)).data(0).data0.value, _
     line_value(tn(1)).data(0).data0.value, True, False), temp_record, no%, 0)
    Exit Function
 ElseIf is_line_value(tp(1, 2), tp(0, 2), 0, 0, 0, _
        "", tn(1), -1000, 0, 0, 0, line_value_data0) = 1 Then
  Call add_conditions_to_record(line_value_, tn(1), 0, 0, _
    temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(verti_, tn_(2), 0, 0, _
    temp_record.record_data.data0.condition_data)
 no% = 0
    Call set_level(temp_record.record_data.data0.condition_data)
 is_area_of_parallelogram = set_area_of_polygon0(Dparallelogram(p%).data(0).polygon4_no, _
    time_string(line_value(tn(0)).data(0).data0.value, _
     line_value(tn(1)).data(0).data0.value, True, False), temp_record, no%, 0)
    Exit Function
 ElseIf is_line_value(tp(1, 3), tp(0, 3), 0, 0, 0, _
         "", tn(1), -1000, 0, 0, 0, line_value_data0) = 1 Then
  Call add_conditions_to_record(line_value_, tn(1), 0, 0, temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(verti_, tn_(3), 0, 0, _
    temp_record.record_data.data0.condition_data)
 no% = 0
    Call set_level(temp_record.record_data.data0.condition_data)
is_area_of_parallelogram = set_area_of_polygon0(Dparallelogram(p%).data(0).polygon4_no, _
    time_string(line_value(tn(0)).data(0).data0.value, _
     line_value(tn(1)).data(0).data0.value, True, False), temp_record, no%, 0)
     Exit Function
 End If
End If
Next i%
End Function
Public Function is_point_in_paral_line(ByVal p%, ByVal in_l%, n%, l%) As Boolean
Dim i% 'in_l%指定的直线,
If is_point_in_line3(p%, m_lin(in_l%).data(0).data0, 0) Then
    l% = in_l% 'p%在in_l%上
    n% = 0
    is_point_in_paral_line = True
Else
For i% = 1 To last_conditions.last_cond(1).paral_no
 If in_l% = 0 Then 'p%在 一对平行线上
   If is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(0)).data(0).data0, 0) Then
    n% = i%
    l% = Dparal(i%).data(0).data0.line_no(0)
    is_point_in_paral_line = True
    Exit Function
   ElseIf is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(1)).data(0).data0, 0) Then
    n% = i%
    l% = Dparal(i%).data(0).data0.line_no(1)
    is_point_in_paral_line = True
    Exit Function
   End If
 Else 'p% 点in_l%平行线上的直线,
  If Dparal(i%).data(0).data0.line_no(0) = in_l% Then
     If is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(1)).data(0).data0, 0) Then
          n% = i%
         l% = Dparal(i%).data(0).data0.line_no(1)
         is_point_in_paral_line = True
         Exit Function
     End If
  ElseIf Dparal(i%).data(0).data0.line_no(1) = in_l% Then
     If is_point_in_line3(p%, m_lin(Dparal(i%).data(0).data0.line_no(1)).data(0).data0, 0) Then
          n% = i%
         l% = Dparal(i%).data(0).data0.line_no(0)
         is_point_in_paral_line = True
         Exit Function
     End If
  End If
 End If
Next i%
End If
End Function


Public Function is_relation0(ByVal p1%, ByVal p2%, ByVal num1%, ByVal num2%) As Boolean
'输入线段比是判断是否已有
Dim i%, l%
Dim tn(1) As Integer
Dim value$
l% = line_number0(p1%, p2%, tn(0), tn(1))
'读取直线
If tn(0) > tn(1) Then
 Call exchange_two_integer(p1%, p2%)
  Call exchange_two_integer(num1%, num2%)
End If
'排序
Call simple_two_int(num1%, num2%)
If num1% > 0 Then
 If num1% = 1 And num2% = 1 Then
 '比值等于一
 For i% = 1 To last_conditions.last_cond(1).mid_point_no
  If Dmid_point(i%).data(0).data0.poi(0) = p1% And Dmid_point(i%).data(0).data0.poi(2) = p2% Then
    is_relation0 = True
     Exit Function
  End If
 Next i%
 Else
  If num2% = 1 Then
   value$ = Trim(str(num1%))
  Else
  value$ = Trim(str(num1%)) + "/" + Trim(str(num2%))
  End If
For i% = 1 To last_conditions.last_cond(1).relation_no
  If Drelation(i%).data(0).data0.line_no(0) = l% And Drelation(i%).data(0).data0.line_no(1) = l% Then
   If Drelation(i%).data(0).data0.poi(0) = p1% And Drelation(i%).data(0).data0.poi(3) = p2% And _
        Drelation(i%).data(0).data0.poi(1) = Drelation(i%).data(0).data0.poi(2) Then
    If Drelation(i%).data(0).data0.value = value$ Then
    is_relation0 = True
     Exit Function
    End If
   End If
  End If
 Next i%
 End If
Else
 num1% = Abs(num1%)
  If num1% > num2% Then
      num1% = num1% - num2%
   If num2% = 1 Then
    value$ = Trim(str(num1%))
   Else
    value$ = Trim(str(num1%)) + "/" + Trim(str(num2%))
   End If
 For i% = 1 To last_conditions.last_cond(1).relation_no
  If Drelation(i%).data(0).data0.line_no(0) = l% And Drelation(i%).data(0).data0.line_no(1) = l% Then
   If Drelation(i%).data(0).data0.poi(0) = p1% And Drelation(i%).data(0).data0.poi(1) = p2% And _
        Drelation(i%).data(0).data0.poi(2) = p2% Then
    If Drelation(i%).data(0).data0.value = value$ Then
    is_relation0 = True
     Exit Function
    End If
   End If
  End If
 Next i%
ElseIf num1% = num2 Then
 is_relation0 = True
  Exit Function
Else
 num2% = num2% - num1%
   If num2% = 1 Then
    value$ = Trim(str(num1%))
   Else
    value$ = Trim(str(num1%)) + "/" + Trim(str(num2%))
   End If
 For i% = 1 To last_conditions.last_cond(1).relation_no
  If Drelation(i%).data(0).data0.line_no(0) = l% And Drelation(i%).data(0).data0.line_no(1) = l% Then
   If Drelation(i%).data(0).data0.poi(1) = p1% And Drelation(i%).data(0).data0.poi(1) = p1% And _
        Drelation(i%).data(0).data0.poi(2) = p2% Then
    If Drelation(i%).data(0).data0.value = value$ Then
    is_relation0 = True
     Exit Function
    End If
   End If
  End If
 Next i%
End If
End If
End Function
Public Function is_two_point_in_polygon(p1%, p2%, ep As polygon) As Byte
Dim i%
If ep.total_v = 3 Then
If is_same_two_point(p1%, p2%, ep.v(0), ep.v(1)) Or _
      is_same_two_point(p1%, p2%, ep.v(0), ep.v(2)) Or _
        is_same_two_point(p1%, p2%, ep.v(2), ep.v(1)) Then
  is_two_point_in_polygon = 1
End If
ElseIf ep.total_v = 4 Then
If is_same_two_point(p1%, p2%, ep.v(0), ep.v(1)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(2)) Or _
        is_same_two_point(p1%, p2%, ep.v(2), ep.v(3)) Or _
          is_same_two_point(p1%, p2%, ep.v(0), ep.v(3)) Then
  is_two_point_in_polygon = 1
ElseIf is_same_two_point(p1%, p2%, ep.v(0), ep.v(2)) = True Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(3)) = True Then
  is_two_point_in_polygon = 2
End If
ElseIf ep.total_v = 5 Then
If is_same_two_point(p1%, p2%, ep.v(0), ep.v(1)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(2)) Or _
        is_same_two_point(p1%, p2%, ep.v(2), ep.v(3)) Or _
          is_same_two_point(p1%, p2%, ep.v(3), ep.v(4)) Or _
            is_same_two_point(p1%, p2%, ep.v(0), ep.v(4)) Then
  is_two_point_in_polygon = 1
ElseIf is_same_two_point(p1%, p2%, ep.v(0), ep.v(2)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(3)) Or _
       is_same_two_point(p1%, p2%, ep.v(2), ep.v(4)) Or _
        is_same_two_point(p1%, p2%, ep.v(3), ep.v(0)) Or _
         is_same_two_point(p1%, p2%, ep.v(4), ep.v(2)) Then
  is_two_point_in_polygon = 2
End If
ElseIf ep.total_v = 6 Then
If is_same_two_point(p1%, p2%, ep.v(0), ep.v(1)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(2)) Or _
        is_same_two_point(p1%, p2%, ep.v(2), ep.v(3)) Or _
          is_same_two_point(p1%, p2%, ep.v(3), ep.v(4)) Or _
            is_same_two_point(p1%, p2%, ep.v(4), ep.v(5)) Or _
              is_same_two_point(p1%, p2%, ep.v(0), ep.v(5)) Then
  is_two_point_in_polygon = 1
ElseIf is_same_two_point(p1%, p2%, ep.v(0), ep.v(2)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(3)) Or _
       is_same_two_point(p1%, p2%, ep.v(2), ep.v(4)) Or _
        is_same_two_point(p1%, p2%, ep.v(3), ep.v(5)) Or _
         is_same_two_point(p1%, p2%, ep.v(4), ep.v(0)) Or _
          is_same_two_point(p1%, p2%, ep.v(5), ep.v(1)) Then
  is_two_point_in_polygon = 2
ElseIf is_same_two_point(p1%, p2%, ep.v(0), ep.v(3)) Or _
      is_same_two_point(p1%, p2%, ep.v(1), ep.v(4)) Or _
       is_same_two_point(p1%, p2%, ep.v(2), ep.v(5)) Then
  is_two_point_in_polygon = 3
End If
End If
End Function

Public Function is_two_line_same(ByVal l1%, ByVal l2%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).same_three_lines_no
 If is_same_two_point(l1%, l2%, same_three_lines(i%).data(0).line_no(0), _
          same_three_lines(i%).data(0).line_no(1)) Or _
    is_same_two_point(l1%, l2%, same_three_lines(i%).data(0).line_no(1), _
          same_three_lines(i%).data(0).line_no(2)) Or _
    is_same_two_point(l1%, l2%, same_three_lines(i%).data(0).line_no(0), _
          same_three_lines(i%).data(0).line_no(2)) Then
     is_two_line_same = True
      Exit Function
 End If
Next i%
End Function
Public Function is_two_line_same0(l_data1 As line_data0_type, l_data2 As line_data0_type) As Byte
'0=不同 1 相同,端点不同2 端点同,所含的点不同3全同
Dim i%, j%, k%
If l_data1.poi(0) = l_data2.poi(0) And l_data1.poi(1) = l_data2.poi(1) Then
   is_two_line_same0 = 2 '端点同
  If l_data1.in_point(0) = l_data2.in_point(0) Then
     For i% = 2 To l_data1.in_point(0) - 1
         If l_data1.in_point(i%) <> l_data2.in_point(i%) Then
          Exit Function '端点同,所含的点不同
         End If
     Next i%
        is_two_line_same0 = 3 '全同
          Exit Function
  Else
     Exit Function '端点同,所含的点不同
  End If
Else
  For i% = 1 To l_data1.in_point(0)
   For j% = 1 To l_data2.in_point(0)
     If l_data1.in_point(i%) = l_data2.in_point(j%) Then
        k% = k + 1
         If k% = 2 Then
            is_two_line_same0 = 1 '1 相同
             Exit Function
         End If
     End If
   Next j%
  Next i%
End If
End Function
Public Function compare_two_record(re1 As record_data_type, re2 As record_data_type) As Integer
Dim re As total_record_type
 Call set_level(re1.data0.condition_data)
  Call set_level(re2.data0.condition_data)
If re1.data0.condition_data.condition_no = 1 Or re2.data0.condition_data.condition_no = 1 Then
 If re1.data0.condition_data.condition_no = 1 Then
  Call record_no(re1.data0.condition_data.condition(1).ty, re1.data0.condition_data.condition(1).no, re, False, 0, 0)
   If re.record_data.data0.condition_data.level + 1 <= re2.data0.condition_data.level Then
    compare_two_record = 1
   Else
    compare_two_record = -1
   End If
 Else
  Call record_no(re2.data0.condition_data.condition(1).ty, re2.data0.condition_data.condition(1).no, re, False, 0, 0)
   If re2.data0.condition_data.level <= re.record_data.data0.condition_data.level + 1 Then
    compare_two_record = 1
   Else
    compare_two_record = -1
   End If
 End If
Else
 If re1.data0.condition_data.condition_no < re2.data0.condition_data.condition_no Then
  compare_two_record = 1
 Else
  If re1.data0.condition_data.level <= re2.data0.condition_data.level Then
   compare_two_record = 1
  Else
   compare_two_record = -1
  End If
 End If
End If
End Function

Public Function is_known_line(ByVal p1%, ByVal p2%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).line_value_no
If is_same_two_point(p1%, p2%, line_value(i%).data(0).data0.poi(0), _
        line_value(i%).data(0).data0.poi(1)) Then
     is_known_line = True
      Exit Function
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no
If is_same_two_point(p1%, p2%, Dmid_point(i%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(1)) Or _
  is_same_two_point(p1%, p2%, Dmid_point(i%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.poi(2)) Or _
   is_same_two_point(p1%, p2%, Dmid_point(i%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(2)) Then
     is_known_line = True
      Exit Function
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no
If is_same_two_point(p1%, p2%, Drelation(i%).data(0).data0.poi(0), Drelation(i%).data(0).data0.poi(1)) Or _
  is_same_two_point(p1%, p2%, Drelation(i%).data(0).data0.poi(2), Drelation(i%).data(0).data0.poi(3)) Then
     is_known_line = True
      Exit Function
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no
If is_same_two_point(p1%, p2%, Deline(i%).data(0).data0.poi(0), Deline(i%).data(0).data0.poi(1)) Or _
  is_same_two_point(p1%, p2%, Deline(i%).data(0).data0.poi(2), Deline(i%).data(0).data0.poi(3)) Then
     is_known_line = True
      Exit Function
End If
Next i%
End Function


Public Function is_condition_in_record1(ByVal ty As Integer, _
           ByVal no%, ByVal level As Byte, old_level As Byte, _
            re As record_data_type) As Boolean
If re.data0.condition_data.level > level And re.data0.condition_data.level = old_level + 1 Then
 is_condition_in_record1 = is_condition_in_record(ty, no%, re, 1)
Else
 is_condition_in_record1 = False
End If
End Function

'Public Function is_two_record_related(con_ty1 As Byte, con_no1 As Integer, _
               re1 As record_data_type, con_ty2 As Byte, con_no2 As Integer, _
                re2 As record_data_type) As Boolean
' If re1.data0.level < re2.data0.level Then
'  is_two_record_related = is_condition_in_record(con_ty1, con_no1, re2, 3)
' ElseIf re1.data0.level < re2.data0.level Then
'  is_two_record_related = is_condition_in_record(con_ty2, con_no2, re1, 3)
' Else
'  If re1.data0.condition_data.condition_no = 1 And re2.data0.condition_data.condition_no = 1 Then
'   If re1.data0.condition_data.condition(1).ty = re2.data0.condition_data.condition(1).ty And _
          re1.data0.condition_data.condition(1).no = re2.data0.condition_data.condition(1).no Then
'    is_two_record_related = True
'   End If
'  End If
' End If
'End Function

Public Function triangle_number_(triA_ As triangle_data0_type) As Integer
Dim triA As triangle_data0_type
triA = triA_
If search_for_triangle(triA, 0, triangle_number_, 0) = False Then
  triangle_number_ = 0
End If
End Function

Public Function is_string_value(str As String, v As String, no%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).string_value_no
 If string_value(i%).data(0).s = str Then
  no% = i%
   is_string_value = no%
    Exit Function
 End If
Next i%
End Function
Public Function is_three_circle_co_point(ByVal c1%, ByVal c2%, _
               ByVal c3%) As Integer
Dim i%, j%
For i% = 1 To m_Circ(c1%).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(c2%).data(0).data0.in_point(0)
  If m_Circ(c1%).data(0).data0.in_point(i%) = m_Circ(c2%).data(0).data0.in_point(j%) Then
     If is_point_in_circle(c3%, 0, m_Circ(c1%).data(0).data0.in_point(i%), 0, 0) Then
           is_three_circle_co_point = m_Circ(c1%).data(0).data0.in_point(i%)
       Exit Function
     End If
  End If
 Next j%
Next i%
End Function

Public Function get_two_circle_inter_point(ByVal ep%, ByVal c1%, _
         ByVal c2%) As Integer
Dim i%, j%
For i% = 1 To m_Circ(c1%).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(c2%).data(0).data0.in_point(0)
  If m_Circ(c1%).data(0).data0.in_point(i%) = m_Circ(c2%).data(0).data0.in_point(j%) And _
       m_Circ(c1%).data(0).data0.in_point(i%) <> ep% Then
      get_two_circle_inter_point = m_Circ(c1%).data(0).data0.in_point(i%)
       Exit Function
  End If
 Next j%
Next i%
End Function
Public Function is_point_inner_circle(ByVal p%, ByVal c%) As Integer
Dim r As Long
If is_point_in_circle(c%, 0, p%, 0, 0) Then
 is_point_inner_circle = 0
Else
r = sqr((m_poi(p%).data(0).data0.coordinate.X - m_Circ(c%).data(0).data0.c_coord.X) ^ 2 + _
         (m_poi(p%).data(0).data0.coordinate.Y - m_Circ(c%).data(0).data0.c_coord.Y) ^ 2)
 If r < m_Circ(c%).data(0).data0.radii Then
 is_point_inner_circle = 1
 Else
 is_point_inner_circle = -1
 End If
End If
End Function

Public Function is_dparal_for_conclusion(ByVal l1%, ByVal l2%, n%) As Boolean
Dim i%
For i% = 0 To 3
 If is_same_two_point( _
        con_paral(i%).data(0).line_no(0), con_paral(i%).data(0).line_no(1), l1%, l2%) Then
         n% = i%
          is_dparal_for_conclusion = True
           Exit Function
 End If
Next i%
End Function
Public Function is_point_in_circle_(p As POINTAPI, c As circle_data0_type) As Integer
Dim r&
r& = sqr((p.X - c.c_coord.X) ^ 2 + (p.Y - c.c_coord.Y) ^ 2)
 r& = r& - c.radii
If r& >= -5 And r& <= 5 Then
 is_point_in_circle_ = 0
ElseIf r& > 5 Then
 is_point_in_circle_ = 1
Else
 is_point_in_circle_ = -1
End If
End Function

Public Function is_complete_prove() As Byte
Dim i%
For i% = 0 To last_conclusion - 1
 If conclusion_data(i%).no(0) = 0 Then
  is_complete_prove = 0
   Exit Function
 End If
Next i%
  is_complete_prove = 2
   If finish_prove = 1 Then
   finish_prove = 2
   End If
End Function

Public Function is_item0_(item_data As item0_data_type, _
             no%, no1%, no2%, no3%) As Boolean
Dim i%
Dim tp(3) As Integer
 If search_for_item0(item_data, 0, no%, 0) Then
  If no1% = -5000 Then
   Call search_for_item0(item_data, 0, no%, 1)
  Else
   is_item0_ = True
      Exit Function
  End If
 Else
  If no1% = -1000 Then
   no% = 0
    Exit Function
  Else
   no1% = no%
    no% = 0
  End If
For i% = 0 To 3
If item_data.poi(i%) > 0 Then
tp(i%) = item_data.poi(i%)
End If
Next i%
  If item_data.sig$ = "*" Or item_data.sig$ = "/" Or item_data.sig$ = "~" Then
         item_data.is_const = True
  End If
 Call search_for_item0(item_data, 1, no2%, 1)
 Call search_for_item0(item_data, 2, no3%, 1)
 End If
End Function

Public Function is_tri_function(ByVal A%, _
          tri_f As tri_function_data_type, no%, n%) As Byte
tri_f.A = A%
If search_for_tri_function(1, tri_f, no%) Then
  is_tri_function = 1
Else
  n% = no%
  is_tri_function = 0
End If
End Function

Public Function is_element_value(ByVal it%, n%, no%, ty As Byte, v$) As Byte
Dim l_v As line_value_data0_type
Dim tr_f As tri_function_data_type
If item0(it%).data(0).poi(2 * n% + 1) > 0 Then
 ty = line_value_
  is_element_value = is_line_value(item0(it%).data(0).poi(2 * n%), item0(it%).data(0).poi(2 * n% + 1), _
    -1000, 0, 0, "", no%, -1000, 0, 0, 0, l_v)
      v$ = line_value(no%).data(0).data0.value
Else
 ty = tri_function_
 is_element_value = is_tri_function(item0(it%).data(0).poi(2 * n%), tr_f, no%, 0)
  If is_element_value = 1 Then
  If item0(it%).data(0).poi(2 * n% + 1) = -2 Then
   v$ = tri_function(no%).data(0).sin_value
  ElseIf item0(it%).data(0).poi(2 * n% + 1) = -3 Then
   v$ = tri_function(no%).data(0).cos_value
  ElseIf item0(it%).data(0).poi(2 * n% + 1) = -4 Then
   v$ = tri_function(no%).data(0).tan_value
  ElseIf item0(it%).data(0).poi(2 * n% + 1) = -5 Then
   v$ = tri_function(no%).data(0).ctan_value
  End If
  End If
End If
End Function

Public Function is_same_record_data(re_data1 As record_data_type, _
           re_data2 As record_data_type) As Boolean
Dim i%
If re_data1.data0.condition_data.condition_no = re_data2.data0.condition_data.condition_no Then
 For i% = 1 To re_data1.data0.condition_data.condition_no
   If re_data1.data0.condition_data.condition(i%).no <> re_data2.data0.condition_data.condition(i%).no Then
    Exit Function
   End If
 Next i%
 is_same_record_data = True
End If
End Function

Public Function is_inter_point_line_circle(l As line_data_type, c As circle_data0_type, p1%, p2%) As Byte
Dim i%, j%, k%
Dim tp(1) As Integer
For i% = 1 To l.data0.in_point(0)
 For j% = 1 To c.in_point(0)
  If l.data0.in_point(i%) = c.in_point(j%) Then
   tp(k%) = l.data0.in_point(i%)
    k% = k% + 1
     is_inter_point_line_circle = k%
      If k% = 2 Then
       p1% = tp(0)
        p2% = tp(1)
       Exit Function
      End If
  End If
 Next j%
Next i%
       p1% = tp(0)
        p2% = tp(1)
End Function

Public Function is_total_angle(ByVal l1%, ByVal l2%, n%, n1%, n2%, t_A As total_angle_data_type) As Boolean
   If two_time_area_triangle(m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
                               m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.Y, _
                                 m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
                                   m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y, _
                              m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
                               m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.X + _
                                m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
                              m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
                               m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.Y + _
                                m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y) > 0 Then
      t_A.line_no(0) = l1%
      t_A.line_no(1) = l2%
  Else
      t_A.line_no(0) = l2%
      t_A.line_no(1) = l1%
  End If
 If search_for_total_angle(t_A, n%, 0, 0) Then
  Exit Function
 Else
  n1% = n%
   Call search_for_total_angle(t_A, n2%, 1, 1)
 End If
End Function

Public Function is_four_sides_fig(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, n%, _
                        f_s_fig As four_sides_fig_data_type) As Boolean
Dim tp(3) As Integer
Dim tn(3) As Integer
Dim i%
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
tn(0) = 0
For i% = 0 To 3
 If tp(tn(0)) > tp(i%) Then
  tn(0) = i%
 End If
Next i%
tn(1) = (tn(0) + 1) Mod 4
tn(2) = (tn(0) + 2) Mod 4
tn(3) = (tn(0) + 3) Mod 4
f_s_fig.poi(0) = tp(tn(0))
f_s_fig.poi(1) = tp(tn(1))
f_s_fig.poi(2) = tp(tn(2))
f_s_fig.poi(3) = tp(tn(3))
If f_s_fig.poi(1) > f_s_fig.poi(3) Then
 Call exchange_two_integer(f_s_fig.poi(1), f_s_fig.poi(3))
End If
 If search_for_four_sides_fig(f_s_fig, 1, 0, n%, 0) Then
     Exit Function
 Else
  f_s_fig.index(0) = n%
  Call search_for_four_sides_fig(f_s_fig, 1, 1, f_s_fig.index(1), 1)
  Call search_for_four_sides_fig(f_s_fig, 1, 2, f_s_fig.index(2), 1)
  Call search_for_four_sides_fig(f_s_fig, 1, 2, f_s_fig.index(3), 1)
 End If
End Function

Public Function is_two_angle_value_180(ByVal A1%, ByVal A2%, con_data As condition_data_type, ty As Boolean) As Boolean
'ty=1 相邻互补
Dim tA3 As angle3_value_data0_type
Dim re As record_data_type
If angle(A1%).data(0).poi(1) = angle(A2%).data(0).poi(1) Then
  If angle(A1%).data(0).line_no(0) = angle(A2%).data(0).line_no(1) And _
      angle(A1%).data(0).line_no(1) = angle(A2%).data(0).line_no(0) Then
       ty = 1
        is_two_angle_value_180 = True
  End If
End If
  is_two_angle_value_180 = is_three_angle_value(A1%, A2%, 0, "1", "1", "0", "180", "180", 0, 0, 0, -1000, 0, 0, 0, 0, 0, 0, _
                       0, tA3, re.data0.condition_data, 0)
   con_data = re.data0.condition_data
End Function
Public Function reduce_to_used_angle(A%, para$, v$, v_$, re As record_data_type, ty As Byte) As Byte
 '相同全角,选顶其中一个ty=0 不计值
Dim n%
If A% = 0 Then
 Exit Function
End If
If angle(A%).data(0).value <> "" And ty = 1 Then
 v$ = minus_string(v, time_string(para$, angle(A%).data(0).value, False, False), True, False)
 Call add_conditions_to_record(angle3_value_, angle(A%).data(0).value_no, _
        0, 0, re.data0.condition_data)
 A% = 0
 para$ = "0"
Else
 n% = T_angle(angle(A%).data(0).total_no).data(0).is_used_no
If n% = -1 Then
   T_angle(angle(A%).data(0).total_no).data(0).is_used_no = angle(A%).data(0).total_no_
Else
   If n% <> angle(A%).data(0).total_no_ Then
    If (Abs(n% - angle(A%).data(0).total_no_)) Mod 2 = 1 Then
     reduce_to_used_angle = 1
      v$ = minus_string(v$, time_string(para$, "180", False, False), True, False)
      v_$ = minus_string(v_$, time_string(para$, "180", False, False), True, False)
      para$ = time_string("-1", para$, True, False)
    End If
      A% = T_angle(angle(A%).data(0).total_no). _
                 data(0).angle_no(n%).no
    
   End If
End If
End If
End Function

Public Function is_not_pseudo_record(re As record_data_type) As Boolean
Dim temp_record As record_data_type
Dim temp_re As total_record_type
Dim i%
temp_record = re
If temp_record.data0.condition_data.condition_no = 1 Then
 If temp_record.data0.condition_data.condition(1).ty = pseudo_similar_triangle_ Or _
      temp_record.data0.condition_data.condition(1).ty = pseudo_total_equal_triangle_ Then
  is_not_pseudo_record = False
 Else
  Call record_no(temp_record.data0.condition_data.condition(1).ty, _
        temp_record.data0.condition_data.condition(1).no, temp_re, False, 0, 0)
   is_not_pseudo_record = is_not_pseudo_record(temp_re.record_data)
  End If
Else
 If temp_record.data0.condition_data.condition_no < 9 Then
  For i% = 1 To temp_record.data0.condition_data.condition_no
    Call record_no(temp_record.data0.condition_data.condition(i%).ty, _
        temp_record.data0.condition_data.condition(i%).no, temp_re, False, 0, 0)
     If is_not_pseudo_record(temp_re.record_data) = False Then
      is_not_pseudo_record = False
      Exit Function
     End If
  Next i%
  is_not_pseudo_record = True
 Else
  is_not_pseudo_record = True
 End If
End If
End Function

Public Function is_same_display_condition(ByVal ty1 As Byte, ByVal n1%, ByVal ty2 As Byte, ByVal n2%) As Boolean
Dim temp_record(1) As total_record_type
Dim i%, j%, k%
If ty1 = ty2 Then
Call record_no(ty1, n1%, temp_record(0), True, 0, 0)
Call record_no(ty2, n2%, temp_record(1), True, 0, 0)
If temp_record(0).record_data.data0.condition_data.condition_no < 8 And _
     temp_record(1).record_data.data0.condition_data.condition_no < 8 Then
If temp_record(0).record_data.data0.condition_data.condition_no = _
    temp_record(1).record_data.data0.condition_data.condition_no Then
   If temp_record(0).record_data.data0.condition_data.condition_no = 0 Then
    is_same_display_condition = True
   Else
    For i% = 1 To temp_record(0).record_data.data0.condition_data.condition_no
     For j% = i% To temp_record(0).record_data.data0.condition_data.condition_no
      If is_same_display_condition(temp_record(0).record_data.data0.condition_data.condition(i%).ty, _
           temp_record(0).record_data.data0.condition_data.condition(i%).no, _
            temp_record(1).record_data.data0.condition_data.condition(j%).ty, _
             temp_record(1).record_data.data0.condition_data.condition(j%).no) Then
              For k% = j% To 2 Step -1
               temp_record(1).record_data.data0.condition_data.condition(k%).ty = _
                  temp_record(1).record_data.data0.condition_data.condition(k% - 1).ty
               temp_record(1).record_data.data0.condition_data.condition(k%).no = _
                  temp_record(1).record_data.data0.condition_data.condition(k% - 1).no
              Next k%
        GoTo is_same_display_condition_mark0
      End If
     Next j%
 is_same_display_condition = False
  Exit Function
is_same_display_condition_mark0:
    Next i%
   End If
 End If
End If
End If
End Function

Public Function get_verti_foot(ByVal p%, ByVal l%, verti_n%, point_n%) As Integer
Dim i%
For i% = 1 To m_lin(l%).data(0).data0.in_point(0)
 If is_dverti(l%, line_number0(p%, m_lin(l%).data(0).data0.in_point(i%), 0, 0), verti_n%, -1000, _
      0, 0, 0, 0) Then
       point_n% = i%
       get_verti_foot = m_lin(l%).data(0).data0.in_point(i%)
        Exit Function
 End If
Next i%
End Function

Public Function is_two_line_value_(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal n1%, ByVal n2%, _
                 ByVal n3%, ByVal n4%, ByVal l1%, ByVal l2%, c_data As condition_data_type, _
                  t_l_value As two_line_value_data0_type) As Byte
Dim re As relation_data0_type
Dim e_l As eline_data0_type
Dim md As mid_point_data0_type
Dim no%
t_l_value.poi(0) = p1%
t_l_value.poi(1) = p2%
t_l_value.poi(2) = p3%
t_l_value.poi(3) = p4%
t_l_value.n(0) = n1%
t_l_value.n(1) = n2%
t_l_value.n(2) = n3%
t_l_value.n(3) = n4%
t_l_value.line_no(0) = l1%
t_l_value.line_no(1) = l2%
If search_for_two_line_value(t_l_value, 0, no%, 0) Then
 Call add_conditions_to_record(two_line_value_, no%, 0, 0, c_data)
 t_l_value = two_line_value(no%).data(0).data0
 is_two_line_value_ = 1
  Exit Function
End If
'************************
re.poi(0) = p1%
re.poi(1) = p2%
re.poi(2) = p3%
re.poi(3) = p4%
re.n(0) = n1%
re.n(1) = n2%
re.n(2) = n3%
re.n(3) = n4%
re.line_no(0) = l1%
re.line_no(1) = l2%
If search_for_relation(re, 0, no%, 0) Then
 Call add_conditions_to_record(relation_, no%, 0, 0, c_data)
  t_l_value.poi(0) = p1%
  t_l_value.poi(1) = p2%
  t_l_value.poi(2) = p3%
  t_l_value.poi(3) = p4%
  t_l_value.n(0) = n1%
  t_l_value.n(1) = n2%
  t_l_value.n(2) = n3%
  t_l_value.n(3) = n4%
  t_l_value.line_no(0) = l1%
  t_l_value.line_no(1) = l2%
  t_l_value.para(0) = "1"
  t_l_value.para(1) = time_string("-1", Drelation(no%).data(0).data0.value, True, False)
  t_l_value.value = "0"
 is_two_line_value_ = 1
  Exit Function
End If
'*********************
e_l.poi(0) = p1%
e_l.poi(1) = p2%
e_l.poi(2) = p3%
e_l.poi(3) = p4%
e_l.n(0) = n1%
e_l.n(1) = n2%
e_l.n(2) = n3%
e_l.n(3) = n4%
e_l.line_no(0) = l1%
e_l.line_no(1) = l2%
If search_for_eline(e_l, 0, no%, 0) Then
 Call add_conditions_to_record(eline_, no%, 0, 0, c_data)
  t_l_value.poi(0) = p1%
  t_l_value.poi(1) = p2%
  t_l_value.poi(2) = p3%
  t_l_value.poi(3) = p4%
  t_l_value.n(0) = n1%
  t_l_value.n(1) = n2%
  t_l_value.n(2) = n3%
  t_l_value.n(3) = n4%
  t_l_value.line_no(0) = l1%
  t_l_value.line_no(1) = l2%
  t_l_value.para(0) = "1"
  t_l_value.para(1) = "-1"
  t_l_value.value = "0"
 is_two_line_value_ = 1
  Exit Function
End If
If l1% = l2% And p2% = p3% Then
md.poi(0) = p1%
md.poi(1) = p2%
md.poi(2) = p4%
md.n(0) = n1%
md.n(1) = n2%
md.n(2) = n4%
md.line_no = l1%
If search_for_mid_point(md, 0, no%, 0) Then
 Call add_conditions_to_record(midpoint_, no%, 0, 0, c_data)
  t_l_value.poi(0) = p1%
  t_l_value.poi(1) = p2%
  t_l_value.poi(2) = p3%
  t_l_value.poi(3) = p4%
  t_l_value.n(0) = n1%
  t_l_value.n(1) = n2%
  t_l_value.n(2) = n3%
  t_l_value.n(3) = n4%
  t_l_value.line_no(0) = l1%
  t_l_value.line_no(1) = l2%
  t_l_value.para(0) = "1"
  t_l_value.para(1) = "-1"
  t_l_value.value = "0"
 is_two_line_value_ = 1
  Exit Function
End If
End If
End Function


Public Function is_l3_value_from_l_l2_value(l3_value As line3_value_data0_type, ByVal k%, ByVal n1%, _
           l_value As line_value_data0_type, c_data As condition_data_type, cond_ty As Byte) As Boolean
Dim t_l3_value As line3_value_data0_type
Dim t_l2_value As two_line_value_data0_type
Dim tn%, tn1%
tn1% = n1%
If k% = 0 Then
t_l3_value = l3_value
ElseIf k% = 1 Then
t_l3_value.poi(0) = l3_value.poi(2)
t_l3_value.poi(1) = l3_value.poi(3)
t_l3_value.poi(2) = l3_value.poi(0)
t_l3_value.poi(3) = l3_value.poi(1)
t_l3_value.poi(4) = l3_value.poi(4)
t_l3_value.poi(5) = l3_value.poi(5)
t_l3_value.n(0) = l3_value.n(2)
t_l3_value.n(1) = l3_value.n(3)
t_l3_value.n(2) = l3_value.n(0)
t_l3_value.n(3) = l3_value.n(1)
t_l3_value.n(4) = l3_value.n(4)
t_l3_value.n(5) = l3_value.n(5)
t_l3_value.line_no(0) = l3_value.line_no(1)
t_l3_value.line_no(1) = l3_value.line_no(0)
t_l3_value.line_no(2) = l3_value.line_no(2)
t_l3_value.value = l3_value.value
t_l3_value.para(0) = l3_value.para(1)
t_l3_value.para(1) = l3_value.para(0)
t_l3_value.para(2) = l3_value.para(2)
Else
t_l3_value.poi(0) = l3_value.poi(4)
t_l3_value.poi(1) = l3_value.poi(5)
t_l3_value.poi(2) = l3_value.poi(0)
t_l3_value.poi(3) = l3_value.poi(1)
t_l3_value.poi(4) = l3_value.poi(2)
t_l3_value.poi(5) = l3_value.poi(3)
t_l3_value.n(0) = l3_value.n(4)
t_l3_value.n(1) = l3_value.n(5)
t_l3_value.n(2) = l3_value.n(0)
t_l3_value.n(3) = l3_value.n(1)
t_l3_value.n(4) = l3_value.n(2)
t_l3_value.n(5) = l3_value.n(3)
t_l3_value.line_no(0) = l3_value.line_no(2)
t_l3_value.line_no(1) = l3_value.line_no(0)
t_l3_value.line_no(2) = l3_value.line_no(1)
t_l3_value.para(0) = l3_value.para(2)
t_l3_value.para(1) = l3_value.para(0)
t_l3_value.para(2) = l3_value.para(1)
t_l3_value.value = l3_value.value
End If
      If is_two_line_value(t_l3_value.poi(2), t_l3_value.poi(3), t_l3_value.poi(4), _
               t_l3_value.poi(5), t_l3_value.n(2), t_l3_value.n(3), t_l3_value.n(4), t_l3_value.n(5), _
                t_l3_value.line_no(1), t_l3_value.line_no(2), t_l3_value.para(1), t_l3_value.para(2), _
                  minus_string(t_l3_value.value, time_string(l_value.value, t_l3_value.para(0), _
                   False, False), True, False), tn%, tn1%, 0, 0, 0, t_l2_value, cond_ty, c_data) = 1 Then
            is_l3_value_from_l_l2_value = True
             Exit Function
      Else
       l3_value.poi(0) = t_l2_value.poi(0)
       l3_value.poi(1) = t_l2_value.poi(1)
       l3_value.poi(2) = t_l2_value.poi(2)
       l3_value.poi(3) = t_l2_value.poi(3)
       l3_value.poi(4) = 0
       l3_value.poi(5) = 0
       l3_value.n(0) = t_l2_value.n(0)
       l3_value.n(0) = t_l2_value.n(1)
       l3_value.n(0) = t_l2_value.n(2)
       l3_value.n(0) = t_l2_value.n(3)
       l3_value.n(0) = 0
       l3_value.n(0) = 0
       l3_value.line_no(0) = t_l2_value.line_no(0)
       l3_value.line_no(1) = t_l2_value.line_no(1)
       l3_value.line_no(2) = 0
       l3_value.para(0) = t_l2_value.para(0)
       l3_value.para(1) = t_l2_value.para(1)
       l3_value.para(2) = "0"
       l3_value.value = t_l2_value.value
       Exit Function
      End If
End Function

Public Function is_l3_value_from_l2_l_value(l3_value As line3_value_data0_type, l2_value As two_line_value_data0_type, _
                 ByVal k%, ByVal n1%, c_data As condition_data_type, con_ty As Byte) As Boolean
Dim t_l3_value As line3_value_data0_type
Dim t_l2_value(1) As two_line_value_data0_type
Dim t_l_value As line_value_data0_type
Dim tn%, tn1%
Dim ts$
tn1% = n1%
If k% = 0 Then
t_l3_value = l3_value
ElseIf k% = 1 Then
t_l3_value.poi(0) = l3_value.poi(2)
t_l3_value.poi(1) = l3_value.poi(3)
t_l3_value.poi(2) = l3_value.poi(0)
t_l3_value.poi(3) = l3_value.poi(1)
t_l3_value.poi(4) = l3_value.poi(4)
t_l3_value.poi(5) = l3_value.poi(5)
t_l3_value.n(0) = l3_value.n(2)
t_l3_value.n(1) = l3_value.n(3)
t_l3_value.n(2) = l3_value.n(0)
t_l3_value.n(3) = l3_value.n(1)
t_l3_value.n(4) = l3_value.n(4)
t_l3_value.n(5) = l3_value.n(5)
t_l3_value.line_no(0) = l3_value.line_no(1)
t_l3_value.line_no(1) = l3_value.line_no(0)
t_l3_value.line_no(2) = l3_value.line_no(2)
t_l3_value.value = l3_value.value
t_l3_value.para(0) = l3_value.para(1)
t_l3_value.para(1) = l3_value.para(0)
t_l3_value.para(2) = l3_value.para(2)
Else
t_l3_value.poi(0) = l3_value.poi(4)
t_l3_value.poi(1) = l3_value.poi(5)
t_l3_value.poi(2) = l3_value.poi(0)
t_l3_value.poi(3) = l3_value.poi(1)
t_l3_value.poi(4) = l3_value.poi(2)
t_l3_value.poi(5) = l3_value.poi(3)
t_l3_value.n(0) = l3_value.n(4)
t_l3_value.n(1) = l3_value.n(5)
t_l3_value.n(2) = l3_value.n(0)
t_l3_value.n(3) = l3_value.n(1)
t_l3_value.n(4) = l3_value.n(2)
t_l3_value.n(5) = l3_value.n(3)
t_l3_value.line_no(0) = l3_value.line_no(2)
t_l3_value.line_no(1) = l3_value.line_no(0)
t_l3_value.line_no(2) = l3_value.line_no(1)
t_l3_value.para(0) = l3_value.para(2)
t_l3_value.para(1) = l3_value.para(0)
t_l3_value.para(2) = l3_value.para(1)
t_l3_value.value = l3_value.value
End If
'If n1% = 0 Then
t_l2_value(0) = l2_value
'Else
't_l2_value(0).poi(0) = l2_value.poi(2)
't_l2_value(0).poi(1) = l2_value.poi(3)
't_l2_value(0).poi(2) = l2_value.poi(0)
't_l2_value(0).poi(3) = l2_value.poi(1)
't_l2_value(0).n(0) = l2_value.n(2)
't_l2_value(0).n(1) = l2_value.n(3)
't_l2_value(0).n(2) = l2_value.n(0)
't_l2_value(0).n(3) = l2_value.n(1)
't_l2_value(0).line_no(0) = l2_value.line_no(1)
't_l2_value(0).line_no(1) = l2_value.line_no(0)
't_l2_value(0).para(0) = l2_value.para(1)
't_l2_value(0).para(1) = l2_value.para(0)
't_l2_value(0).value = l2_value.value
'End If
ts$ = t_l3_value.para(0)
t_l3_value.para(0) = time_string(t_l3_value.para(0), t_l2_value(0).para(0), True, False)
t_l3_value.para(1) = time_string(t_l3_value.para(1), t_l2_value(0).para(0), True, False)
t_l3_value.para(2) = time_string(t_l3_value.para(2), t_l2_value(0).para(0), True, False)
t_l3_value.value = time_string(t_l3_value.value, t_l2_value(0).para(0), True, False)
t_l2_value(0).para(0) = time_string(t_l2_value(0).para(0), ts$, True, False)
t_l2_value(0).para(1) = time_string(t_l2_value(0).para(1), ts$, True, False)
t_l2_value(0).value = time_string(t_l2_value(0).value, ts$, True, False)
If t_l2_value(0).poi(2) = t_l3_value.poi(2) And t_l2_value(0).poi(3) = t_l3_value.poi(3) Then
   t_l2_value(1).poi(0) = t_l2_value(0).poi(2)
   t_l2_value(1).poi(1) = t_l2_value(0).poi(3)
   t_l2_value(1).n(0) = t_l2_value(0).n(2)
   t_l2_value(1).n(1) = t_l2_value(0).n(3)
   t_l2_value(1).line_no(0) = t_l2_value(0).line_no(1)
   t_l2_value(1).para(0) = minus_string(t_l3_value.para(1), t_l2_value(0).para(1), True, False)
   t_l2_value(1).poi(2) = t_l3_value.poi(4)
   t_l2_value(1).poi(3) = t_l3_value.poi(5)
   t_l2_value(1).n(2) = t_l3_value.n(4)
   t_l2_value(1).n(3) = t_l3_value.n(5)
   t_l2_value(1).line_no(1) = t_l3_value.line_no(2)
   t_l2_value(1).para(1) = t_l3_value.para(2)
   t_l2_value(1).value = minus_string(t_l3_value.value, t_l2_value(0).value, True, False)
ElseIf t_l2_value(0).poi(2) = t_l3_value.poi(4) And t_l2_value(0).poi(3) = t_l3_value.poi(5) Then
   t_l2_value(1).poi(0) = t_l3_value.poi(2)
   t_l2_value(1).poi(1) = t_l3_value.poi(3)
   t_l2_value(1).n(0) = t_l3_value.n(2)
   t_l2_value(1).n(1) = t_l3_value.n(3)
   t_l2_value(1).line_no(0) = t_l3_value.line_no(1)
   t_l2_value(1).para(0) = t_l3_value.para(1)
   t_l2_value(1).poi(2) = t_l3_value.poi(4)
   t_l2_value(1).poi(3) = t_l3_value.poi(5)
   t_l2_value(1).n(2) = t_l3_value.n(4)
   t_l2_value(1).n(3) = t_l3_value.n(5)
   t_l2_value(1).line_no(1) = t_l3_value.line_no(2)
   t_l2_value(1).para(1) = minus_string(t_l3_value.para(2), t_l2_value(0).para(1), True, False)
   t_l2_value(1).value = minus_string(t_l3_value.value, t_l2_value(0).value, True, False)
Else
 Exit Function
End If
          If is_two_line_value(t_l2_value(1).poi(0), t_l2_value(1).poi(1), t_l2_value(1).poi(2), t_l2_value(1).poi(3), _
               t_l2_value(1).n(0), t_l2_value(1).n(1), t_l2_value(1).n(2), t_l2_value(1).n(3), t_l2_value(1).line_no(0), _
                t_l2_value(1).line_no(1), t_l2_value(1).para(0), t_l2_value(1).para(1), t_l2_value(1).value, tn%, -1000, _
                 0, 0, 0, t_l2_value(1), con_ty, c_data) = 1 Then
                  is_l3_value_from_l2_l_value = True
          End If
           l3_value.poi(0) = t_l2_value(1).poi(0)
           l3_value.poi(1) = t_l2_value(1).poi(1)
           l3_value.poi(2) = t_l2_value(1).poi(2)
           l3_value.poi(3) = t_l2_value(1).poi(3)
           l3_value.poi(4) = 0
           l3_value.poi(5) = 0
           l3_value.n(0) = t_l2_value(1).n(0)
           l3_value.n(1) = t_l2_value(1).n(1)
           l3_value.n(2) = t_l2_value(1).n(2)
           l3_value.n(3) = t_l2_value(1).n(3)
           l3_value.n(4) = 0
           l3_value.n(5) = 0
           l3_value.line_no(0) = t_l2_value(1).line_no(0)
           l3_value.line_no(1) = t_l2_value(1).line_no(1)
           l3_value.line_no(2) = t_l2_value(1).line_no(2)
           l3_value.para(0) = t_l2_value(1).para(0)
           l3_value.para(1) = t_l2_value(1).para(1)
           l3_value.para(2) = "0"
           l3_value.value = t_l2_value(1).value
End Function

Public Function is_polygon4_(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
          poly4_data As polygon4_data_type, no%, n1%, dir%) As Boolean
Dim A(1) As Integer
A(0) = angle_number(p1%, p2%, p3%, "", 0)
A(1) = angle_number(p3%, p4%, p1%, "", 0)
If A(0) * A(1) < 0 Then
 Call exchange_two_integer(p1%, p2%)
End If
is_polygon4_ = is_polygon4(p1%, p2%, p3%, p4%, poly4_data, no%, n1%, dir%)
End Function
Public Function is_polygon4(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
          poly4_data As polygon4_data_type, no%, n1%, dir%) As Boolean
Dim tp(3) As Integer
Dim i%, k%, l%, m%
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
k% = 0
m% = 0
For i% = 1 To 3
 For l% = 0 To i% - 1
  If tp(l%) = tp(i%) Then
   Exit Function '有相同的点
  End If
 Next l%
If tp(k%) > tp(i%) Then
   k% = i%
End If
If tp(m%) < tp(i%) Then
   m% = i%
End If
Next i%
is_polygon4 = True
If angle_number(tp(k%), tp((k% + 1) Mod 4), tp((k% + 2) Mod 4), "", 0) < 0 Then
poly4_data.poi(0) = tp(k%)
poly4_data.poi(1) = tp((4 + k% - 1) Mod 4)
poly4_data.poi(2) = tp((4 + k% - 2) Mod 4)
poly4_data.poi(3) = tp((4 + k% - 3) Mod 4)
If k% = 0 Or k% = 2 Then
dir% = -1
Else
dir% = 1
End If
Else
poly4_data.poi(0) = tp(k%)
poly4_data.poi(1) = tp((k% + 1) Mod 4)
poly4_data.poi(2) = tp((k% + 2) Mod 4)
poly4_data.poi(3) = tp((k% + 3) Mod 4)
If k% = 0 Or k% = 2 Then
dir% = 1
Else
dir% = -1
End If
End If
poly4_data.poi(4) = tp(m%) '最后一点
If search_for_polygon4(poly4_data, n1%, 0) Then
   is_polygon4 = True
    no% = n1%
Else
  poly4_data.angle(0) = Abs(angle_number(poly4_data.poi(3), poly4_data.poi(0), poly4_data.poi(1), "", 0))
  poly4_data.angle(1) = Abs(angle_number(poly4_data.poi(0), poly4_data.poi(1), poly4_data.poi(2), "", 0))
  poly4_data.angle(2) = Abs(angle_number(poly4_data.poi(1), poly4_data.poi(2), poly4_data.poi(3), "", 0))
  poly4_data.angle(3) = Abs(angle_number(poly4_data.poi(2), poly4_data.poi(3), poly4_data.poi(0), "", 0))
   is_polygon4 = False
End If
End Function
Public Function is_equal_side_tixing0(poly4_no%, _
     no%, op1%, op2%, op3%, op4%, cond_ty As Byte) As Boolean
Dim i%
If poly4_no% = 0 Then
 is_equal_side_tixing0 = True
  no% = 0
Else
 If Dpolygon4(poly4_no%).data(0).start_poi = 0 Then
  op1% = Dpolygon4(poly4_no%).data(0).poi(0)
  op2% = Dpolygon4(poly4_no%).data(0).poi(1)
  op3% = Dpolygon4(poly4_no%).data(0).poi(2)
  op4% = Dpolygon4(poly4_no%).data(0).poi(3)
 Else
  op1% = Dpolygon4(poly4_no%).data(0).poi(1)
  op2% = Dpolygon4(poly4_no%).data(0).poi(2)
  op3% = Dpolygon4(poly4_no%).data(0).poi(3)
  op4% = Dpolygon4(poly4_no%).data(0).poi(0)
 End If
For i% = 1 To last_conditions.last_cond(0).tixing_no
 If Dtixing(i%).data(0).poly4_no = poly4_no% Then
  If Dpolygon4(Dtixing(i%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
  no% = i%
   is_equal_side_tixing0 = True
    Exit Function
  End If
 End If
Next i%
End If
End Function
Public Function is_tixing0(poly4_data As polygon4_data_type, no%, _
       op1%, op2%, op3%, op4%, cond_ty As Byte, set_or_reduce As Boolean) As Boolean
If is_dparal(line_number0(poly4_data.poi(0), poly4_data.poi(1), 0, 0), _
    line_number0(poly4_data.poi(2), poly4_data.poi(3), 0, 0), 0, 0, 0, 0, 0, 0) Then
      op1% = poly4_data.poi(0)
      op2% = poly4_data.poi(1)
      op3% = poly4_data.poi(2)
      op4% = poly4_data.poi(3)
      is_tixing0 = True
ElseIf is_dparal(line_number0(poly4_data.poi(1), poly4_data.poi(2), 0, 0), _
    line_number0(poly4_data.poi(3), poly4_data.poi(0), 0, 0), 0, 0, 0, 0, 0, 0) Then
      op1% = poly4_data.poi(1)
      op2% = poly4_data.poi(2)
      op3% = poly4_data.poi(3)
      op4% = poly4_data.poi(0)
      poly4_data.start_poi = 1
      is_tixing0 = True
Else
      is_tixing0 = False
End If
End Function
Public Function is_long_squre0(ByVal poly4_no%, _
                        no%, tn%, cond_ty As Byte) As Boolean
Dim i%
If poly4_no% = 0 Then
 If tn% <> -1000 Then
 is_long_squre0 = True
 Else
 is_long_squre0 = False
 End If
  no% = 0
Else
 For i% = 1 To last_conditions.last_cond(1).squre_no
      If Dsqure(i%).data(0).polygon4_no = poly4_no% Then
       no% = i%
       cond_ty = Squre
        is_long_squre0 = True
         Exit Function
      End If
 Next i%
 For i% = 1 To last_conditions.last_cond(1).long_squre_no
   If Dlong_squre(i%).data(0).polygon4_no = poly4_no% Then
         no% = i%
    cond_ty = long_squre_
     is_long_squre0 = True
      Exit Function
   End If
Next i%
 is_long_squre0 = False
End If
End Function
Public Function is_rhombus0(ByVal poly4_no%, n%, tn%, cond_ty As Byte) As Boolean
 Dim i%
 If poly4_no% = 0 Then
  If tn% <> -1000 Then
  is_rhombus0 = True
  Else
   is_rhombus0 = False
  End If
   Exit Function
 Else
 For i% = 1 To last_conditions.last_cond(1).rhombus_no
 If rhombus(i%).data(0).polygon4_no = poly4_no% Then
   n% = i%
    cond_ty = rhombus_
     is_rhombus0 = True
      Exit Function
 End If
Next i%
   n% = 0
    cond_ty = 0
     is_rhombus0 = False
End If
End Function
Public Function is_parallelogram0(poly4_no%, _
                    tn%, n1%, cond_ty As Byte) As Boolean
Dim i%
If last_conditions.last_cond(1).parallelogram_no = 0 And n1% = -1000 Then
 tn% = 0
   is_parallelogram0 = False
    Exit Function
End If
 If search_for_parallelogram(poly4_no%, tn%, 0) Then
    cond_ty = parallelogram_
   If set_or_prove = 2 Then '
    If Dparallelogram(tn%).data(0).record.data1.is_proved = 1 Then
   is_parallelogram0 = True
    End If
   Else
   is_parallelogram0 = True
   End If
    'tn% = i%
     Exit Function
 End If
If n1% <> -1000 Then
 n1% = tn%
End If
End Function
Public Function is_squre0(poly4_no%, no%, tn%) As Boolean
Dim i%
If poly4_no% = 0 Then
 If no% <> -1000 Then
 is_squre0 = True
 Else
 is_squre0 = False
 End If
  no% = 0
Else
For i% = 1 To last_conditions.last_cond(1).squre_no
 If poly4_no% = Dsqure(i%).data(0).polygon4_no Then
  is_squre0 = True
   Exit Function
 End If
Next i%
is_squre0 = False
End If
End Function

Public Function is_squre(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, no%, tn%, poly4_no%) As Boolean
poly4_no% = polygon4_number(p1%, p2%, p3%, p4%, 0)
If poly4_no% = 0 Then
 If tn% <> -1000 Then
  no% = 0
  is_squre = True
 Else
  is_squre = False
 End If
 no% = 0
 Exit Function
Else
 If Dpolygon4(poly4_no%).data(0).ty = Squre Then
  no% = Dpolygon4(poly4_no%).data(0).no
  is_squre = True
  Exit Function
 End If
End If
is_squre = is_squre0(poly4_no%, no%, tn%)
End Function
Public Function is_squre_length(ByVal s_no%, ByVal l_no%) As Boolean
Dim i%
Dim lv As line_value_data0_type
If Dsqure(s_no%).data(0).length_of_diag_no = 0 And _
     Dsqure(s_no%).data(0).length_of_side_no = 0 And _
       Dsqure(s_no%).data(0).radii_no = 0 Then
 If l_no% = 0 Then
 For i% = 0 To 3
  is_squre_length = is_line_value(Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(i%), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), _
            0, 0, 0, "", Dsqure(s_no%).data(0).length_of_side_no, -1000, 0, _
              0, 0, lv)
   If is_squre_length Then
     Exit Function
   End If
 Next i%
   is_squre_length = is_line_value(Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(0), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(2), _
             0, 0, 0, "", Dsqure(s_no%).data(0).length_of_diag_no, -1000, 0, _
              0, 0, lv)
   If is_squre_length Then
     Exit Function
   End If
  is_squre_length = is_line_value(Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(1), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(3), _
            0, 0, 0, "", Dsqure(s_no%).data(0).length_of_diag_no, -1000, 0, _
              0, 0, lv)
   If is_squre_length Then
     Exit Function
   End If
  If Dsqure(s_no%).data(0).four_point_on_circle_no > 0 Then
     If m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center > 0 Then
      If m_poi(m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center).data(0).data0.visible > 0 Then
   For i% = 0 To 3
     is_squre_length = is_line_value(Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(i%), _
          m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center, _
            0, 0, 0, "", Dsqure(s_no%).data(0).radii_no, -1000, 0, _
              0, 0, lv)
      If is_squre_length Then
       Exit Function
      End If
    Next i%
    End If
    End If
   End If
 Else 'l_no%>0
  For i% = 0 To 3
   If is_same_two_point(line_value(l_no%).data(0).data0.poi(0), _
         line_value(l_no%).data(0).data0.poi(1), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(i%), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4)) Then
          Dsqure(s_no%).data(0).length_of_side_no = l_no%
      is_squre_length = True
       Exit Function
   End If
 Next i%
   If is_same_two_point(line_value(l_no%).data(0).data0.poi(0), _
         line_value(l_no%).data(0).data0.poi(1), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(0), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(2)) Then
          Dsqure(s_no%).data(0).length_of_diag_no = l_no%
      is_squre_length = True
       Exit Function
   End If
   If is_same_two_point(line_value(l_no%).data(0).data0.poi(0), _
         line_value(l_no%).data(0).data0.poi(1), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(1), _
          Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(3)) Then
          Dsqure(s_no%).data(0).length_of_diag_no = l_no%
      is_squre_length = True
       Exit Function
   End If
  If Dsqure(s_no%).data(0).four_point_on_circle_no > 0 Then
     If m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center > 0 Then
      If m_poi(m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center).data(0).data0.visible > 0 Then
   For i% = 0 To 3
      If is_same_two_point(line_value(l_no%).data(0).data0.poi(0), _
         line_value(l_no%).data(0).data0.poi(1), _
              Dpolygon4(Dsqure(s_no%).data(0).polygon4_no).data(0).poi(i%), _
          m_Circ(Dsqure(s_no%).data(0).four_point_on_circle_no).data(0).data0.center) Then
          Dsqure(s_no%).data(0).radii_no = l_no%
         is_squre_length = True
       Exit Function
      End If
    Next i%
    End If
    End If
   End If
 End If
Else
 is_squre_length = True
End If
End Function

Public Function is_general_string_satis_midpoint(ByVal gs%, p1%, p2%, p3%, p4%, _
         p_1%, p_2%, pa3$, it1%, it2%, pA1$, pA2$) As Byte '0 false,1,liang bian,2 mid and bian
         '判断三项和是否满足中位线定理'p1%,p2%,p3%,p4%,否满足中位线定理可以合并的点'p_1%,p_2%分母,
         ' it1%,it2% p21$,pa2$选择的两项,pa3$,新项的系数
Dim tl(1) As Integer
Dim ty(2) As Boolean
Dim sig As String
Dim c_data0 As condition_data_type
 If general_string(gs%).record_.conclusion_no > 0 Then '结论
  If general_string(gs%).data(0).value = "" And general_string(gs%).data(0).para(2) <> "0" Then '三项
   If (item0(general_string(gs%).data(0).item(0)).data(0).sig = "~" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "~" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "~") Or _
     (item0(general_string(gs%).data(0).item(0)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "/") Then
     '三项是线段和或比值和
     If item0(general_string(gs%).data(0).item(0)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "/" Then
     '如果是比值和
     ty(0) = is_same_two_point(item0(general_string(gs%).data(0).item(0)).data(0).poi(2), _
          item0(general_string(gs%).data(0).item(0)).data(0).poi(3), _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(2), _
            item0(general_string(gs%).data(0).item(1)).data(0).poi(3))
     ty(1) = is_same_two_point(item0(general_string(gs%).data(0).item(1)).data(0).poi(2), _
          item0(general_string(gs%).data(0).item(1)).data(0).poi(3), _
           item0(general_string(gs%).data(0).item(2)).data(0).poi(2), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(3))
     ty(2) = is_same_two_point(item0(general_string(gs%).data(0).item(0)).data(0).poi(2), _
          item0(general_string(gs%).data(0).item(0)).data(0).poi(3), _
           item0(general_string(gs%).data(0).item(2)).data(0).poi(2), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(3))
     sig = "/"
     If ty(0) = True Then '第一二的分母同
      p_1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(2)
      p_2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(3)
     ElseIf ty(1) = True Then
      p_1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(2)
      p_2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(3)
     ElseIf ty(2) = True Then
      p_1% = item0(general_string(gs%).data(0).item(2)).data(0).poi(2)
      p_2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(3) '分母
     End If
     Else
     ty(0) = True
     ty(1) = True
     ty(2) = True
     sig = "~"
     End If
    'End If
      it2% = general_string(gs%).data(0).item(3) '
      pA2$ = general_string(gs%).data(0).para(3) '
      If general_string(gs%).data(0).para(0) = "1" And (ty(0) Or ty(2)) Then '第一项
       If general_string(gs%).data(0).para(1) = "1" And ty(0) Then '第二项
         it1% = general_string(gs%).data(0).item(2)  '第三项
         pA1$ = general_string(gs%).data(0).para(2)
         pa3$ = "2"
          p1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          is_general_string_satis_midpoint = 1
        ElseIf general_string(gs%).data(0).para(2) = "1" And ty(2) Then
         it1% = general_string(gs%).data(0).item(1)
         pA1$ = general_string(gs%).data(0).para(1)
         pa3$ = "2"
          p1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 1
       ElseIf general_string(gs%).data(0).para(1) = "-2" And ty(0) Then
         it1% = general_string(gs%).data(0).item(2)
         pA1$ = general_string(gs%).data(0).para(2)
         pa3$ = "-1"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
       ElseIf general_string(gs%).data(0).para(2) = "-2" And ty(2) Then
         it1% = general_string(gs%).data(0).item(1)
         pA1$ = general_string(gs%).data(0).para(1)
         pa3$ = "-1"
          p1% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
       Else
         Exit Function
       End If
      ElseIf general_string(gs%).data(0).para(0) = "2" And (ty(1) Or ty(2)) Then
       If (general_string(gs%).data(0).para(1) = "-1" Or _
            general_string(gs%).data(0).para(1) = "@1") And _
             (general_string(gs%).data(0).para(2) = "-1" Or _
               general_string(gs%).data(0).para(2) = "@1") And ty(1) Then
         it1% = general_string(gs%).data(0).item(0)
         pA1$ = general_string(gs%).data(0).para(0)
         pa3$ = "-2"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
           is_general_string_satis_midpoint = 1
       ElseIf (general_string(gs%).data(0).para(1) = "-1" Or _
                general_string(gs%).data(0).para(1) = "@1") And ty(0) Then
         it1% = general_string(gs%).data(0).item(2)
         pA1$ = general_string(gs%).data(0).para(2)
         pa3$ = "1"
          p1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
        ElseIf (general_string(gs%).data(0).para(2) = "-1" Or _
                general_string(gs%).data(0).para(2) = "@1") And ty(1) Then
          it1% = general_string(gs%).data(0).item(1)
         pA1$ = general_string(gs%).data(0).para(1)
         pa3$ = "1"
          p1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
        Else
         Exit Function
        End If
      ElseIf general_string(gs%).data(0).para(1) = "1" And _
                  (ty(0) Or ty(1)) Then
         it1% = general_string(gs%).data(0).item(0)
         pA1$ = general_string(gs%).data(0).para(0)
          If general_string(gs%).data(0).para(2) = "1" Then
          pa3$ = "2"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 1
            ElseIf general_string(gs%).data(0).para(2) = "-2" Then
          pa3$ = "-1"
          p1% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
            End If
      ElseIf (general_string(gs%).data(0).para(1) = "-1" Or _
               general_string(gs%).data(0).para(1) = "@1") And _
                  (ty(0) Or ty(1)) Then
           it1% = general_string(gs%).data(0).item(0)
         pA1$ = general_string(gs%).data(0).para(0)
          If (general_string(gs%).data(0).para(2) = "-1" Or _
                general_string(gs%).data(0).para(2) = "@1") Then
           pa3$ = "-2"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 1
           ElseIf general_string(gs%).data(0).para(2) = "2" Then
           pa3$ = "1"
          p1% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
           End If

      ElseIf general_string(gs%).data(0).para(1) = "2" And _
                 (ty(0) Or ty(1)) Then
             If (general_string(gs%).data(0).para(2) = "-1" Or _
                  general_string(gs%).data(0).para(2) = "@1") Then
              it1% = general_string(gs%).data(0).item(0)
              pA1$ = general_string(gs%).data(0).para(0)
           pa3$ = "1"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
              End If
     ElseIf (general_string(gs%).data(0).para(1) = "-2" Or _
              general_string(gs%).data(0).para(1) = "@2") And _
            (ty(0) Or ty(1)) Then
            If general_string(gs%).data(0).para(2) = "1" Then
              it1% = general_string(gs%).data(0).item(0)
              pA1$ = general_string(gs%).data(0).para(0)
           pa3$ = "-1"
          p1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(0)
          p2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(1)
          p3% = item0(general_string(gs%).data(0).item(2)).data(0).poi(0)
          p4% = item0(general_string(gs%).data(0).item(2)).data(0).poi(1)
          is_general_string_satis_midpoint = 2
            End If
      Else
        Exit Function
      End If
    'is_general_string_satis_midpoint = True
  End If
 End If
End If
End Function

Public Function is_equation(ByVal s As String, e As Equation_data0_type, no%, n1%, re As total_record_type) As Byte
Dim n%
Dim ts$
If string_type(s, "", ts$, "", "") = 3 Then
   s = ts$
End If
is_equation = simple_equation(s, re, ts$)
If is_equation >= 1 Then
 Exit Function
ElseIf InStr(1, ts$, "F", 0) > 0 Or ts$ = "" Then
   is_equation = 1
    no% = 0
   Exit Function
End If
If read_para_from_equation(ts$, e) Then
   If search_for_equation(e, no%, 0) Then
    is_equation = 1
   Else
    no% = 0
     n1% = no%
    is_equation = 0
'     Call solve_equation(e.para(0), e.para(1), e.para(2), e.root(0), e.root(1), False)
   End If
Else
 If n1% = -1000 Then
  is_equation = 0
 Else
  is_equation = 1
  no% = 0
 End If
End If
End Function
'Public Function is_equal_side_triangle(p1%, p2%, p3%, _
                             tri As triangle_data0_type, dir As Integer, no%) As Boolean
'If p2% > p3% Then
'call exchange_two_integer(p2%, p3%)
'End If
'End Function

Public Function is_general_string_contain_squ_sum(gs As general_string_data_type, _
              ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%) As Boolean
If is_item0_contain_squ_sum(gs.item(0), p1%, p2%, p3%, p4%, p5%, p6%) Then
    is_general_string_contain_squ_sum = True
     Exit Function
ElseIf is_item0_contain_squ_sum(gs.item(1), p1%, p2%, p3%, p4%, p5%, p6%) Then
    is_general_string_contain_squ_sum = True
     Exit Function
ElseIf is_item0_contain_squ_sum(gs.item(2), p1%, p2%, p3%, p4%, p5%, p6%) Then
    is_general_string_contain_squ_sum = True
     Exit Function
ElseIf is_item0_contain_squ_sum(gs.item(3), p1%, p2%, p3%, p4%, p5%, p6%) Then
    is_general_string_contain_squ_sum = True
     Exit Function
End If
End Function

Public Function is_item0_contain_squ_sum(ByVal it%, ByVal p1%, ByVal p2%, ByVal p3%, _
                  ByVal p4%, ByVal p5%, ByVal p6%) As Boolean
If item0(it%).data(0).sig = "*" Then
   If item0(it%).data(0).poi(0) = item0(it%).data(0).poi(2) And _
       item0(it%).data(0).poi(1) = item0(it%).data(0).poi(3) Then
        If is_same_two_point(p1%, p2%, item0(it%).data(0).poi(0), _
              item0(it%).data(0).poi(1)) Then
           is_item0_contain_squ_sum = True
            Exit Function
        ElseIf is_same_two_point(p3%, p4%, item0(it%).data(0).poi(0), _
              item0(it%).data(0).poi(1)) Then
           is_item0_contain_squ_sum = True
            Exit Function
        ElseIf is_same_two_point(p5%, p6%, item0(it%).data(0).poi(0), _
              item0(it%).data(0).poi(1)) Then
           is_item0_contain_squ_sum = True
            Exit Function
        End If
   End If
End If

End Function

Public Function is_item0_squ(it As item0_data_type, p1%, p2%) As Boolean
If it.sig = "*" Then
 If it.poi(0) = it.poi(2) And it.poi(1) = it.poi(3) Then
  p1% = it.poi(0)
  p2% = it.poi(1)
  is_item0_squ = True
 End If
End If
End Function

Public Function is_general_string_sum_squ(gs As general_string_data_type, it1%, it2%, _
        para1$, para2$) As Boolean
Dim squ_item(4) As Integer
Dim tp(3) As Integer
Dim tp_(1) As Integer
Dim i%
Dim ty As Byte
Dim Tpara As String
Dim c_data0 As condition_data_type
squ_item(0) = -1
squ_item(1) = -1
squ_item(2) = -1
squ_item(3) = -1

For i% = 0 To 3
 If gs.para(i%) = "2" Or gs.para(i%) = "-2" Then
  If item0(gs.item(i%)).data(0).sig = "*" Then
   If item0(gs.item(i%)).data(0).line_no(0) = item0(gs.item(i%)).data(0).line_no(1) Then
    If item0(gs.item(i%)).data(0).n(0) = item0(gs.item(i%)).data(0).n(2) Then
     squ_item(0) = i%
     ty = 1
GoTo is_general_string_sum_squ_mark0
    ElseIf item0(gs.item(i%)).data(0).n(1) = item0(gs.item(i%)).data(0).n(3) Then
     ty = 1
     squ_item(0) = i%
GoTo is_general_string_sum_squ_mark0
    ElseIf item0(gs.item(i%)).data(0).n(1) = item0(gs.item(i%)).data(0).n(2) Then
     ty = 3
      squ_item(0) = i%
GoTo is_general_string_sum_squ_mark0
    Else
     GoTo is_general_string_sum_squ_next
    End If
   End If
  End If
 End If
is_general_string_sum_squ_next:
Next i%
Exit Function
is_general_string_sum_squ_mark0:
tp(0) = item0(gs.item(squ_item(0))).data(0).poi(0)
tp(1) = item0(gs.item(squ_item(0))).data(0).poi(1)
tp(2) = item0(gs.item(squ_item(0))).data(0).poi(2)
tp(3) = item0(gs.item(squ_item(0))).data(0).poi(3)
For i% = 0 To 3
 If i% <> squ_item(0) Then
  Tpara = time_string(gs.para(i%), gs.para(squ_item(0)), True, False)
  If (ty = 1 And Tpara = "-2") Or (ty = 3 And Tpara = "2") Then
   If is_item0_squ(item0(gs.item(i%)).data(0), tp_(0), tp_(1)) Then
    If is_same_two_point(tp_(0), tp_(1), tp(0), tp(1)) Then
     tp(0) = 0
     tp(1) = 0
      GoTo is_general_string_sum_squ_next1
    ElseIf is_same_two_point(tp_(0), tp_(1), tp(2), tp(3)) Then
     tp(2) = 0
     tp(3) = 0
      GoTo is_general_string_sum_squ_next1
    End If
   End If
  End If
squ_item(1) = i%
 End If
is_general_string_sum_squ_next1:
Next i%
If tp(0) = 0 And tp(1) = 0 And tp(2) = 0 And tp(3) = 0 Then
'Exit Function
'Else
 If squ_item(1) >= 0 Then
   it2% = gs.item(squ_item(1))
   para2$ = gs.para(squ_item(1))
 Else
   it2% = 0
   para2$ = "0"
 End If
 If ty = 1 Then
  para1$ = divide_string(gs.para(squ_item(0)), "-2", True, False)
   If item0(gs.item(squ_item(0))).data(0).poi(0) = item0(gs.item(squ_item(0))).data(0).poi(2) Then
  Call set_item0(item0(gs.item(squ_item(0))).data(0).poi(1), item0(gs.item(squ_item(0))).data(0).poi(3), _
            item0(gs.item(squ_item(0))).data(0).poi(1), item0(gs.item(squ_item(0))).data(0).poi(3), _
             "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, c_data0, 0, it1%, 0, 0, condition_data0, False)
  ElseIf item0(gs.item(squ_item(0))).data(0).poi(1) = item0(gs.item(squ_item(0))).data(0).poi(3) Then
 Call set_item0(item0(gs.item(squ_item(0))).data(0).poi(0), item0(gs.item(squ_item(0))).data(0).poi(2), _
            item0(gs.item(squ_item(0))).data(0).poi(0), item0(gs.item(squ_item(0))).data(0).poi(2), _
             "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, c_data0, 0, it1%, 0, 0, condition_data0, False)
   End If
 Else
  para1$ = divide_string(gs.para(squ_item(0)), "2", True, False)
  Call set_item0(item0(gs.item(squ_item(0))).data(0).poi(4), item0(gs.item(squ_item(0))).data(0).poi(5), _
            item0(gs.item(squ_item(0))).data(0).poi(4), item0(gs.item(squ_item(0))).data(0).poi(5), _
             "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, c_data0, 0, it1%, 0, 0, condition_data0, False)
             
 End If
 is_general_string_sum_squ = True
End If
End Function


Public Function calcutete_area_of_fan(ByVal a_v As String, ByVal r As String) As String
Dim arc As String
If InStr(1, a_v, ".", 0) = 0 Then
 arc = divide_string(a_v, "180", True, False)
   calcutete_area_of_fan = time_string(arc, time_string(r, r, False, False), False, False)
     calcutete_area_of_fan = divide_string(calcutete_area_of_fan, "2", True, False)
      If InStr(1, calcutete_area_of_fan, "/", 0) > 0 Then
       calcutete_area_of_fan = "(" + calcutete_area_of_fan + ")" + "\" 'LoadResString_(1456,"")
      Else
       calcutete_area_of_fan = calcutete_area_of_fan + "\" 'LoadResString_(1456,"")
      End If
Else
 arc = divide_string(a_v, "180", True, True)
  arc = time_string(arc, PI, False, False)
   calcutete_area_of_fan = time_string(arc, time_string(r, r, False, False), False, False)
     calcutete_area_of_fan = divide_string(calcutete_area_of_fan, "2", True, False)
End If
End Function

Public Function chose_line_value_from_two_line_value(ByVal p1%, ByVal p2%, ByVal s1$, ByVal S2$, ByVal sum$) As String
Dim i%, j%, k%, no%
'Dim n(1) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim l(1)  As Long
Dim n_(1) As Integer
Dim t_line As two_line_value_data0_type
Dim temp_record As total_record_type
l(0) = (m_poi(p1%).data(0).data0.coordinate.X - _
        m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p1%).data(0).data0.coordinate.Y - _
           m_poi(p2%).data(0).data0.coordinate.Y) ^ 2
For j% = 0 To 1
  m(0) = j%
   m(1) = (j% + 1) Mod 2
t_line.poi(2 * m(0)) = p1%
t_line.poi(2 * m(0) + 1) = p2%
t_line.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_line, m(0), n_(0), 1)
t_line.poi(2 * m(1)) = 30000
Call search_for_two_line_value(t_line, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
If two_line_value(no%).data(0).data0.para(0) = "1" And _
      two_line_value(no%).data(0).data0.para(1) = "1" And _
       two_line_value(no%).data(0).data0.value = sum$ Then
 l(1) = (m_poi(two_line_value(no%).data(0).data0.poi(2 * m(1))).data(0).data0.coordinate.X - _
        m_poi(two_line_value(no%).data(0).data0.poi(2 * m(1) + 1)).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(two_line_value(no%).data(0).data0.poi(2 * m(1))).data(0).data0.coordinate.Y - _
           m_poi(two_line_value(no%).data(0).data0.poi(2 * m(1) + 1)).data(0).data0.coordinate.Y) ^ 2
 If l(0) > l(1) Then
  chose_line_value_from_two_line_value = s1$
 Else
  chose_line_value_from_two_line_value = S2$
 End If
  Exit Function
 End If
Next k%
Next j%
chose_line_value_from_two_line_value = s1$
End Function
Public Function is_equal_sides_triangle(triA%, no%, re As condition_data_type) As Boolean
Dim i%
Dim con_ty As Byte
Dim con_ty_ As Byte
Dim dn(2) As Integer
Dim dn_(2) As Integer
Dim el_data0 As eline_data0_type
Dim temp_triA As triangle_data0_type
Dim a3_v0 As angle3_value_data0_type
Dim re_data0 As record_data0_type
For i% = 1 To last_conditions.last_cond(1).epolygon_no
 If epolygon(i%).data(0).p.total_v = 3 And epolygon(i%).data(0).no = triA% Then
  no% = i%
   is_equal_sides_triangle = True
    Exit Function
 End If
Next i%
re.condition_no = 0
temp_triA = triangle(triA%).data(0)
 If angle(temp_triA.angle(0)).data(0).value = "60" And angle(temp_triA.angle(1)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(0)).data(0).value_no, _
        angle(temp_triA.angle(1)).data(0).value_no, 0, re)
   is_equal_sides_triangle = True
    Exit Function
 End If
 If angle(temp_triA.angle(1)).data(0).value = "60" And angle(temp_triA.angle(2)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(1)).data(0).value_no, _
        angle(temp_triA.angle(2)).data(0).value_no, 0, re)
   is_equal_sides_triangle = True
    Exit Function
 End If
 If angle(temp_triA.angle(0)).data(0).value = "60" And angle(temp_triA.angle(2)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(0)).data(0).value_no, _
        angle(temp_triA.angle(2)).data(0).value_no, 0, re)
   is_equal_sides_triangle = True
    Exit Function
 End If
If angle(temp_triA.angle(0)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(0)).data(0).value_no, _
        0, 0, re)
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(2), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(0), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
       Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
       Exit Function
  ElseIf is_three_angle_value(temp_triA.angle(1), temp_triA.angle(2), 0, "1", "-1", "0", _
           "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
  
ElseIf angle(temp_triA.angle(1)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(1)).data(0).value_no, _
        0, 0, re)
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(2), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(0), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
       Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
       Exit Function
  ElseIf is_three_angle_value(temp_triA.angle(0), temp_triA.angle(2), 0, "1", "-1", "0", _
          "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
ElseIf angle(temp_triA.angle(2)).data(0).value = "60" Then
  Call add_conditions_to_record(angle3_value_, angle(temp_triA.angle(2)).data(0).value_no, _
        0, 0, re)
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(2), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
     is_equal_sides_triangle = True
      Exit Function
  ElseIf is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(0), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
      is_equal_sides_triangle = True
       Exit Function
  ElseIf is_three_angle_value(temp_triA.angle(1), temp_triA.angle(0), 0, "1", "-1", "0", _
          "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
End If
re.condition_no = 0
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(2), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) And _
    is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(0), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn_(0), -1000, 0, 0, 0, el_data0, dn_(1), dn_(2), con_ty_, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(con_ty_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
re.condition_no = 0
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(2), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) And _
    is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn_(0), -1000, 0, 0, 0, el_data0, dn_(1), dn_(2), con_ty_, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(con_ty_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
re.condition_no = 0
  If is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(1), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, el_data0, dn(1), dn(2), con_ty, "", re_data0.condition_data) And _
    is_equal_dline(temp_triA.poi(0), temp_triA.poi(1), temp_triA.poi(0), temp_triA.poi(2), _
     0, 0, 0, 0, 0, 0, dn_(0), -1000, 0, 0, 0, el_data0, dn_(1), dn_(2), con_ty_, "", re_data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(con_ty_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
re.condition_no = 0
  If is_three_angle_value(temp_triA.angle(0), temp_triA.angle(1), 0, "1", "-1", "0", _
          "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) And _
          is_three_angle_value(temp_triA.angle(1), temp_triA.angle(2), 0, "1", "-1", "0", _
          "0", "0", dn_(0), dn_(1), dn_(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(angle3_value_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
re.condition_no = 0
  If is_three_angle_value(temp_triA.angle(0), temp_triA.angle(1), 0, "1", "-1", "0", _
          "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) And _
          is_three_angle_value(temp_triA.angle(0), temp_triA.angle(2), 0, "1", "-1", "0", _
          "0", "0", dn_(0), dn_(1), dn_(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(angle3_value_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
re.condition_no = 0
  If is_three_angle_value(temp_triA.angle(0), temp_triA.angle(2), 0, "1", "-1", "0", _
          "0", "0", dn(0), dn(1), dn(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) And _
          is_three_angle_value(temp_triA.angle(1), temp_triA.angle(2), 0, "1", "-1", "0", _
          "0", "0", dn_(0), dn_(1), dn_(2), -1000, 0, 0, 0, 0, 0, 0, 0, a3_v0, re_data0.condition_data, 0) Then
      Call add_conditions_to_record(angle3_value_, dn(0), dn(1), dn(2), re)
      Call add_conditions_to_record(angle3_value_, dn_(0), dn_(1), dn_(2), re)
      is_equal_sides_triangle = True
       Exit Function
  End If
  
End Function

Public Function set_point_pair_data( _
    ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
       ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, _
        ByVal in1%, ByVal in2%, ByVal in3%, ByVal in4%, ByVal in5%, _
         ByVal in6%, ByVal in7%, in8%, il1%, il2%, il3%, il4%, _
            dp As point_pair_data0_type) As Boolean
 'n3%,n4% 记录line_value 导出的相等
Dim t(7) As Boolean
Dim i%, k%
Dim tn(7) As Integer
Dim tn2(7) As Integer
Dim v(3) As String
Dim con_ty(7) As Byte
Dim t_dp As point_pair_data0_type
Dim tc_data(3) As condition_data_type
If p1% <= 0 Or p2% <= 0 Or p3% <= 0 Or p4% <= 0 Or _
    p5% <= 0 Or p6% <= 0 Or p7% <= 0 Or p8% <= 0 Then
 set_point_pair_data = False
End If
dp = t_dp
'ty=1 1=2;ty =2 1=3;ty=3 3=4; ty =4 4=2;ty=5 1=2,3=4;by=6 1=3,2=4
If il1% = 0 Or run_type = 10 Then
dp.line_no(0) = line_number0(p1%, p2%, in1%, in2%)
Else
dp.line_no(0) = il1%
End If
If in1% <= in2% Then
dp.n(0) = in1%
dp.n(1) = in2%
dp.poi(0) = p1%
dp.poi(1) = p2%
Else
dp.n(0) = in2%
dp.n(1) = in1%
dp.poi(0) = p2%
dp.poi(1) = p1%
'call exchange_two_integer(p1%, p2%)
'call exchange_two_integer(n1%, n2%)
End If
If il2% = 0 Or run_type = 10 Then
dp.line_no(1) = line_number0(p3%, p4%, in3%, in4%)
Else
dp.line_no(1) = il2%
End If
If in3% <= in4% Then
dp.n(2) = in3%
dp.n(3) = in4%
dp.poi(2) = p3%
dp.poi(3) = p4%
Else
dp.n(2) = in4%
dp.n(3) = in3%
dp.poi(2) = p4%
dp.poi(3) = p3%
'call exchange_two_integer(p3%, p4%)
'call exchange_two_integer(n3%, n4%)
End If
If il3% = 0 Or run_type = 10 Then
dp.line_no(2) = line_number0(p5%, p6%, in5%, in6%)
Else
dp.line_no(2) = il3%
End If
If in5% <= in6% Then
dp.n(4) = in5%
dp.n(5) = in6%
dp.poi(4) = p5%
dp.poi(5) = p6%
Else
dp.n(4) = in6%
dp.n(5) = in5%
dp.poi(4) = p6%
dp.poi(5) = p5%
'call exchange_two_integer(p5%, p6%)
'call exchange_two_integer(n5%, n6%)
End If
If il4% = 0 Or run_type = 10 Then
dp.line_no(3) = line_number0(p7%, p8%, in7%, in8%)
Else
dp.line_no(3) = il4%
End If
If in7% <= in8% Then
dp.n(6) = in7%
dp.n(7) = in8%
dp.poi(6) = p7%
dp.poi(7) = p8%
Else
dp.n(6) = in8%
dp.n(7) = in7%
dp.poi(6) = p8%
dp.poi(7) = p7%
'call exchange_two_integer(p7%, p8%)
'call exchange_two_integer(n7%, n8%)
End If
Call simple_point_pair(dp, dp)
set_point_pair_data = True
End Function
Public Function is_pseudo_dpoint_pair(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                                       ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, no%, _
                                        dp As point_pair_data0_type) As Boolean
Dim i%
dp.poi(0) = p1%
dp.poi(1) = p2%
dp.poi(2) = p3%
dp.poi(3) = p4%
dp.poi(4) = p5%
dp.poi(5) = p6%
dp.poi(6) = p7%
dp.poi(7) = p8%
dp.line_no(0) = line_number0(dp.poi(0), dp.poi(1), dp.n(0), dp.n(1))
dp.line_no(1) = line_number0(dp.poi(2), dp.poi(3), dp.n(2), dp.n(3))
dp.line_no(2) = line_number0(dp.poi(4), dp.poi(5), dp.n(4), dp.n(5))
dp.line_no(3) = line_number0(dp.poi(6), dp.poi(7), dp.n(6), dp.n(7))
If dp.n(0) > dp.n(1) Then
 Call exchange_two_integer(dp.poi(0), dp.poi(1))
 Call exchange_two_integer(dp.n(0), dp.n(1))
End If
If dp.n(2) > dp.n(3) Then
 Call exchange_two_integer(dp.poi(2), dp.poi(3))
 Call exchange_two_integer(dp.n(2), dp.n(3))
End If
If dp.n(4) > dp.n(5) Then
 Call exchange_two_integer(dp.poi(4), dp.poi(5))
 Call exchange_two_integer(dp.n(4), dp.n(5))
End If
If dp.n(6) > dp.n(7) Then
 Call exchange_two_integer(dp.poi(6), dp.poi(7))
 Call exchange_two_integer(dp.n(6), dp.n(7))
End If
Call simple_point_pair(dp, dp)
For i% = 1 To last_conditions.last_cond(1).pseudo_dpoint_pair_no
 If pseudo_dpoint_pair(i%).data(0).data0.poi(0) = dp.poi(0) And _
     pseudo_dpoint_pair(i%).data(0).data0.poi(1) = dp.poi(1) And _
      pseudo_dpoint_pair(i%).data(0).data0.poi(2) = dp.poi(2) And _
       pseudo_dpoint_pair(i%).data(0).data0.poi(3) = dp.poi(3) And _
        pseudo_dpoint_pair(i%).data(0).data0.poi(4) = dp.poi(4) And _
         pseudo_dpoint_pair(i%).data(0).data0.poi(5) = dp.poi(5) And _
          pseudo_dpoint_pair(i%).data(0).data0.poi(6) = dp.poi(6) And _
           pseudo_dpoint_pair(i%).data(0).data0.poi(7) = dp.poi(7) Then
            no% = i%
             is_pseudo_dpoint_pair = True
              Exit Function
 End If
Next i%
 no% = 0
  is_pseudo_dpoint_pair = False
End Function

Public Function is_pseudo_eline(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, no%, _
           el As eline_data0_type) As Boolean
Dim i%
Call arrange_four_point(p1%, p2%, _
  p3%, p4%, 0, 0, 0, 0, 0, 0, el.poi(0), el.poi(1), _
    el.poi(2), el.poi(3), 0, 0, el.n(0), el.n(1), _
     el.n(2), el.n(3), 0, 0, el.line_no(0), el.line_no(1), 0, 0, condition_data0, 0)
For i% = 1 To last_conditions.last_cond(1).pseudo_eline_no
 If pseudo_eline(i%).data(0).data0.poi(0) = el.poi(0) And _
     pseudo_eline(i%).data(0).data0.poi(1) = el.poi(1) And _
      pseudo_eline(i%).data(0).data0.poi(2) = el.poi(2) And _
       pseudo_eline(i%).data(0).data0.poi(3) = el.poi(3) Then
        no% = i%
        is_pseudo_eline = True
         Exit Function
 End If
Next i%
is_pseudo_eline = False
 no% = 0
End Function

Public Function simple_three_angle_value0(angle3_v As angle3_value_data0_type, ty As Byte) As angle3_value_data0_type
Dim A3_v As angle3_value_data0_type 'ty=0 condition,ty=1 conclusion
Dim ts As String
Dim re_data As record_data_type
A3_v = angle3_v
A3_v.angle(3) = 0
A3_v.ty(0) = 0
A3_v.ty_(0) = 0
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no _
      Or A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
If (angle(A3_v.angle(1)).data(0).total_no > angle(A3_v.angle(2)).data(0).total_no Or _
   A3_v.angle(1) = 0) And A3_v.angle(2) > 0 Then
Call exchange_two_integer(A3_v.angle(1), A3_v.angle(2))
 Call exchange_string(A3_v.para(1), A3_v.para(2))
End If
If (angle(A3_v.angle(0)).data(0).total_no > angle(A3_v.angle(1)).data(0).total_no Or _
     A3_v.angle(0) = 0) And A3_v.angle(1) > 0 Then
 Call exchange_two_integer(A3_v.angle(0), A3_v.angle(1))
  Call exchange_string(A3_v.para(0), A3_v.para(1))
End If
'设置首项系数
'***********************
ts = ""
Call simple_multi_string0(A3_v.para(0), A3_v.para(1), A3_v.para(2), "0", ts, True)
If A3_v.value <> "" And ts <> "1" Then
 A3_v.value = divide_string(A3_v.value, ts, True, False)
 A3_v.value_ = divide_string(A3_v.value_, ts, True, False)
End If
If ty = 1 Then
 simple_three_angle_value0 = A3_v
 Exit Function
End If
 Call combine_two_angle_with_para(A3_v.angle(0), A3_v.angle(1), A3_v.angle(3), _
        A3_v.angle_(3), A3_v.para(0), A3_v.para(1), A3_v.value, A3_v.value_, A3_v.ty(0), _
          A3_v.ty_(0), 1, re_data)
 ts = ""
  Call simple_multi_string0(A3_v.para(0), A3_v.para(1), A3_v.para(2), "0", ts, True)
   If A3_v.value <> "" And ts <> "1" Then
    A3_v.value = divide_string(A3_v.value, ts, True, False)
     A3_v.value_ = divide_string(A3_v.value_, ts, True, False)
   End If
simple_three_angle_value0 = A3_v
End Function
Public Sub is_depend_condition(record_data As record_data_type)
'判断相关条件'相关条件condition_data.condition_no=100
Dim i%
Dim re As total_record_type
Dim t_record_data As record_data_type
t_record_data = record_data
If t_record_data.data0.condition_data.condition_no < 9 Then
  For i% = 1 To t_record_data.data0.condition_data.condition_no
    Call record_no(t_record_data.data0.condition_data.condition(i%).ty, _
           t_record_data.data0.condition_data.condition(i%).no, re, False, 0, 0)
     If re.record_data.data0.condition_data.condition_no = 0 Then
            Call set_depend_condition(t_record_data.data0.condition_data.condition(i%).ty, _
                        t_record_data.data0.condition_data.condition(i%).no)
     Else
            Call is_depend_condition(re.record_data)
     End If
  Next i%
End If
End Sub
Public Sub is_sufficient_condition()
Dim n%
For n% = 1 To C_display_wenti.m_last_condition
 If C_display_wenti.m_depend_no(n%) = 0 Then
      error_condition_no = n%
        error_of_wenti = 4
         Exit Sub
 End If
Next n%
End Sub
Public Function is_contain_two_nukwon_element(ByVal s$, ele1 As String) As Boolean
Dim ty As Byte
Dim ts(3) As String
Dim i%
Dim ch As String
ty = string_type(s$, ts(0), ts(1), ts(2), ts(3))
If ty = 3 Then
   If is_contain_two_nukwon_element(ts(1), ele1) = False Then
       is_contain_two_nukwon_element = is_contain_two_nukwon_element(ts(1), ele1)
   End If
Else
 If ts(3) = "" Then
    For i% = 1 To Len(ts(1))
        ch = Mid$(ts(1), i%, 1)
         If ch >= "A" And ch <= "z" Then
            If ele1 = "" Then
            ele1 = ch
            ElseIf ele1 <> ch Then
             is_contain_two_nukwon_element = True
               Exit Function
            End If
         End If
    Next i%
    For i% = 1 To Len(ts(2))
        ch = Mid$(ts(2), i%, 1)
         If ch >= "A" And ch <= "z" Then
            If ele1 = "" Then
            ele1 = ch
            ElseIf ele1 <> ch Then
             is_contain_two_nukwon_element = True
               Exit Function
            End If
         End If
    Next i%
 Else
   If is_contain_two_nukwon_element(ts(0), ele1) = False Then
       is_contain_two_nukwon_element = is_contain_two_nukwon_element(ts(3), ele1)
   End If
 End If
End If
End Function
Public Function is_area_of_element(ty As Integer, ele%, no%, n1%) As Boolean
Dim area_ele As area_of_element_data_type
'If ty = triangle_ Then
' is_area_of_element = is_area_of_triangle(ele%, no%)
'ElseIf ty = polygon_ Then
 area_ele.element.ty = ty
 area_ele.element.no = ele%
 If ele% = 0 Then
   If n1% = -1000 Then
      is_area_of_element = False
   Else
      is_area_of_element = True
      no% = 0
   End If
    Exit Function
 End If
 is_area_of_element = search_for_area_element(area_ele, 1, no%, 0)
'End If
End Function
Public Function is_area_of_element0(element As condition_type, no%, n1%) As Boolean
  is_area_of_element0 = is_area_of_element(element.ty, element.no, no%, n1%)
 End Function
Public Function is_uselly_para(para$) As Boolean
 If para$ = "1" Or para = "+1" Or para = "#1" Then
    para$ = "1"
    is_uselly_para = True
 ElseIf para$ = "-1" Or para = "@1" Then
    para$ = "-1"
    is_uselly_para = True
 ElseIf para$ = "2" Or para = "+2" Or para = "#2" Then
    para$ = "2"
    is_uselly_para = True
 ElseIf para$ = "-2" Or para = "@2" Then
    para$ = "-2"
    is_uselly_para = True
 ElseIf para$ = "0" Or para = "+0" Or para = "#0" Or para$ = "-0" Or para = "@0" Then
    para$ = "0"
    is_uselly_para = True
 Else
    is_uselly_para = False
 End If
End Function
Public Function arrange_four_point_on_circle(p1%, p2%, p3%, p4%) As Boolean
Dim tp(4) As Integer
Dim i%, j%, p%
Dim ang(3) As Integer
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
For i% = 2 To 0 Step -1
 For j% = 0 To i%
  If tp(j%) > tp(j% + 1) Or (tp(j%) = 0 And tp(j% + 1) > 0) Then
   Call exchange_two_integer(tp(j%), tp(j% + 1))
  ElseIf tp(j%) = tp(j% + 1) Then
   tp(j% + 1) = 0
    For p% = j% + 1 To 2
     tp(p%) = tp(p% + 1)
    Next p%
     tp(3) = 0
  End If
 Next j%
Next i%
If tp(3) = 0 Then
 arrange_four_point_on_circle = False
Else
ang(0) = angle_number(tp(0), tp(2), tp(1), 0, 0)
ang(1) = angle_number(tp(0), tp(3), tp(1), 0, 0)
If ang(0) > 0 And ang(1) > 0 Then
 p1% = tp(0)
  p4% = tp(1)
  ang(2) = angle_number(tp(2), tp(0), tp(3), 0, 0)
  ang(3) = angle_number(tp(2), tp(1), tp(3), 0, 0)
   If ang(2) > 0 And ang(3) > 0 Then
    p2% = tp(3)
     p3% = tp(2)
      arrange_four_point_on_circle = True
   ElseIf ang(2) < 0 And ang(3) < 0 Then
    p2% = tp(2)
     p3% = tp(3)
      arrange_four_point_on_circle = True
   Else
       arrange_four_point_on_circle = False
   End If
ElseIf ang(0) > 0 And ang(1) < 0 Then
   ang(2) = angle_number(tp(2), tp(0), tp(3), 0, 0)
   ang(3) = angle_number(tp(2), tp(1), tp(3), 0, 0)
 If ang(2) < 0 And ang(3) > 0 Then
  p1% = tp(0)
   p2% = tp(2)
    p3% = tp(1)
     p4% = tp(3)
      arrange_four_point_on_circle = True
 Else
       arrange_four_point_on_circle = False
 End If
ElseIf ang(0) < 0 And ang(1) > 0 Then
   ang(2) = angle_number(tp(2), tp(0), tp(3), 0, 0)
   ang(3) = angle_number(tp(2), tp(1), tp(3), 0, 0)
 If ang(2) > 0 And ang(3) < 0 Then
  p1% = tp(0)
   p2% = tp(3)
    p3% = tp(1)
     p4% = tp(2)
      arrange_four_point_on_circle = True
 Else
       arrange_four_point_on_circle = False
 End If
ElseIf ang(0) < 0 And ang(1) < 0 Then
p1% = tp(0)
 p2% = tp(1)
   ang(2) = angle_number(tp(2), tp(0), tp(3), 0, 0)
    ang(3) = angle_number(tp(2), tp(1), tp(3), 0, 0)
   If ang(2) < 0 And ang(3) < 0 Then
    p3% = tp(2)
     p4% = tp(3)
      arrange_four_point_on_circle = True
   ElseIf ang(2) > 0 And ang(3) > 0 Then
    p3% = tp(3)
     p4% = tp(2)
      arrange_four_point_on_circle = True
   Else
      arrange_four_point_on_circle = False
   End If
Else
  arrange_four_point_on_circle = False
End If
End If
End Function
Public Function vector_number(ByVal p1%, ByVal p2%, dir As String) As Integer
Dim tn%
Dim re As record_data_type
If m_poi(p1%).data(0).data0.visible = 0 Or _
       m_poi(p2%).data(0).data0.visible = 0 Then
      Exit Function
End If
 If search_for_two_point_line(p1%, p2%, tn%, 0) Then
  vector_number = tn%
   If Dtwo_point_line(tn%).data(0).v_poi(0) = p1% Then
      dir = "1"
   Else
      dir = "-1"
   End If
 Else
  Call set_line_from_two_point(p1%, p2%, _
           0, 0, 0, tn%, dir, re)
   vector_number = tn%
 End If
End Function

Public Function is_distance_of_paral(lv%, pl%, vei1%, vei2%) As Byte
 Dim i%, tp%
 If lv% > 0 Then
  For i% = last_conditions.last_cond(0).paral_no + 1 To last_conditions.last_cond(1).paral_no
    pl% = Dparal(i%).data(0).data0.record.data1.index.i(0)
     If Dparal(pl%).data(0).distance_no > 0 Then
           If Dparal(pl%).data(0).distance_no = lv% Then
                       is_distance_of_paral = 1 '
                  Exit Function
           End If
     Else
      If is_dverti(Dparal(pl%).data(0).data0.line_no(0), line_value(lv%).data(0).data0.line_no, _
          vei1%, -1000, 0, 0, 0, 0) Then '
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1).data(0).inter_poi, Dparal(pl%).data(0).data0.line_no(1), _
                      line_value(lv%).data(0).data0.line_no, pl%, vei1%, vei2%)
             If is_distance_of_paral > 1 Then
                 Exit Function
             End If
     ElseIf is_dverti(Dparal(pl%).data(0).data0.line_no(1), line_value(lv%).data(0).data0.line_no, _
          vei1%, -1000, 0, 0, 0, 0) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1).data(0).inter_poi, Dparal(pl%).data(0).data0.line_no(0), _
                      line_value(lv%).data(0).data0.line_no, pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
    End If
   End If
 Next i%
ElseIf pl% > 0 Then
 If Dparal(pl%).data(0).distance_no > 0 Then
    is_distance_of_paral = 1
     Exit Function
Else
For i% = last_conditions.last_cond(0).verti_no + 1 To last_conditions.last_cond(1).verti_no
  vei1% = Dverti(i%).data(0).record.data1.index.i(0)
       If Dverti(i%).data(0).inter_poi > 0 Then
          If Dparal(pl%).data(0).data0.line_no(0) = Dverti(vei1%).data(0).line_no(0) Then
          is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(1), _
                      Dverti(vei1%).data(0).line_no(1), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           ElseIf Dparal(pl%).data(0).data0.line_no(1) = Dverti(i%).data(0).line_no(0) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(0), _
                      Dverti(vei1%).data(0).line_no(1), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           ElseIf Dparal(pl%).data(0).data0.line_no(0) = Dverti(vei1%).data(0).line_no(1) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(i%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(1), _
                      Dverti(vei1%).data(0).line_no(0), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           ElseIf Dparal(pl%).data(0).data0.line_no(1) = Dverti(vei1%).data(0).line_no(1) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(i%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(0), _
                      Dverti(vei1%).data(0).line_no(0), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           End If
        End If
       Next i%
 End If
ElseIf vei1% > 0 Then
 If Dverti(vei1%).data(0).inter_poi > 0 Then
 For i% = last_conditions.last_cond(0).paral_no + 1 To last_conditions.last_cond(1).paral_no
     pl% = Dparal(i%).data(0).data0.record.data1.index.i(0)
           If Dparal(pl%).data(0).data0.line_no(0) = Dverti(vei1%).data(0).line_no(0) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(1), _
                      Dverti(vei1%).data(0).line_no(1), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           ElseIf Dparal(pl%).data(0).data0.line_no(1) = Dverti(vei1%).data(0).line_no(0) Then
             is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(0), _
                      Dverti(vei1%).data(0).line_no(1), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           ElseIf Dparal(pl%).data(0).data0.line_no(0) = Dverti(vei1%).data(0).line_no(1) Then
            is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(1), _
                      Dverti(vei1%).data(0).line_no(0), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
          ElseIf Dparal(pl%).data(0).data0.line_no(1) = Dverti(vei1%).data(0).line_no(1) Then
           is_distance_of_paral = is_distance_of_paral0(lv%, Dverti(vei1%).data(0).inter_poi, _
                 Dparal(pl%).data(0).data0.line_no(0), _
                      Dverti(vei1%).data(0).line_no(0), pl%, vei1%, vei2%)
             If is_distance_of_paral = 1 Then
                 Exit Function
             End If
           End If
Next i%
 End If
End If
End Function
Public Function is_distance_of_paral0(ByVal lv%, ByVal tp1%, ByVal l1%, ByVal l2%, ByVal pl%, _
            ByVal vei1%, distance_of_paral_line_no%) As Byte
Dim tp%, i%
Dim l_v As line_value_data0_type
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).distance_of_paral_line_no
    If lv% > 0 And pl% > 0 Then
    If Ddistance_of_paral_line(i%).data(0).paral_no = pl% And _
        Ddistance_of_paral_line(i%).data(0).lv_no = lv% Then
        distance_of_paral_line_no% = i%
         is_distance_of_paral0 = 1
          Exit Function
    End If
    If Ddistance_of_paral_line(i%).data(0).paral_no = pl% Then
       lv% = Ddistance_of_paral_line(i%).data(0).paral_no
    ElseIf Ddistance_of_paral_line(i%).data(0).paral_no = lv% Then
    End If
    End If
Next i%
tp% = is_line_line_intersect(l1%, l2%, 0, 0, False)
If tp1% > 0 And tp% > 0 Then
If lv% > 0 Then
If is_same_two_point(tp%, tp1%, line_value(lv%).data(0).data0.poi(0), _
       line_value(lv%).data(0).data0.poi(1)) Then
      temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(paral_, pl%, 0, 0, temp_record.record_data.data0.condition_data)
      Call add_conditions_to_record(verti_, vei1%, 0, 0, temp_record.record_data.data0.condition_data)
      Call add_conditions_to_record(line_value_, lv%, 0, 0, temp_record.record_data.data0.condition_data)
      is_distance_of_paral0 = set_distance_of_paral_line(pl%, lv%, "", distance_of_paral_line_no%, _
                                             temp_record.record_data)
      If is_distance_of_paral0 > 1 Then
         Exit Function
      End If
End If
ElseIf is_line_value(tp1%, tp%, 0, 0, 0, "", lv%, -1000, 0, 0, 0, l_v) Then
      temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(paral_, pl%, 0, 0, temp_record.record_data.data0.condition_data)
      Call add_conditions_to_record(verti_, vei1%, 0, 0, temp_record.record_data.data0.condition_data)
      Call add_conditions_to_record(line_value_, lv%, 0, 0, temp_record.record_data.data0.condition_data)
      is_distance_of_paral0 = set_distance_of_paral_line(pl%, lv%, "", distance_of_paral_line_no%, _
                                            temp_record.record_data)
      If is_distance_of_paral0 > 1 Then
         Exit Function
      End If
End If
End If
End Function
Public Function is_two_paral_same(ByVal l1%, ByVal l2%, ByVal no2%, outl1%, outl2%) As Boolean
 If l1% > l2% Then
  Call exchange_two_integer(l1%, l2%)
 End If
 If l1% = Dparal(no2%).data(0).data0.line_no(0) Then
    If l2% = Dparal(no2%).data(0).data0.line_no(1) Then
       outl1% = 0
       outl2% = 0
       is_two_paral_same = True
    ElseIf is_line_line_intersect(l2%, Dparal(no2%).data(0).data0.line_no(1), 0, 0, False) Then
       is_two_paral_same = True
       outl1% = l2%
       outl2% = Dparal(no2%).data(0).data0.line_no(1)
    End If
 ElseIf l1% = Dparal(no2%).data(0).data0.line_no(1) Then
    If is_line_line_intersect(l2%, Dparal(no2%).data(0).data0.line_no(0), 0, 0, False) Then
       is_two_paral_same = True
       outl1% = l2%
       outl2% = Dparal(no2%).data(0).data0.line_no(0)
    End If
 ElseIf l2% = Dparal(no2%).data(0).data0.line_no(0) Then
    If is_line_line_intersect(l1%, Dparal(no2%).data(0).data0.line_no(1), 0, 0, False) Then
       is_two_paral_same = True
       outl1% = l1%
       outl2% = Dparal(no2%).data(0).data0.line_no(1)
    End If
 ElseIf l2% = Dparal(no2%).data(0).data0.line_no(1) Then
    If is_line_line_intersect(l1%, Dparal(no2%).data(0).data0.line_no(0), 0, 0, False) Then
       is_two_paral_same = True
       outl1% = l1%
       outl2% = Dparal(no2%).data(0).data0.line_no(0)
    End If
 End If
End Function
Public Function is_two_paral_same_(ByVal l1%, ByVal l2%, no%, outl1%, outl2%) As Boolean
For no% = last_conditions.last_cond(0).paral_no + 1 To _
                            last_conditions.last_cond(1).paral_no
    If is_two_paral_same(l1%, l2%, Dparal(no%).data(0).data0.record.data1.index.i(0), _
           outl1%, outl2%) Then
            is_two_paral_same_ = True
             Exit Function
    End If
Next no%
End Function
Public Function is_kwon_radii_for_circle(ByVal c%, ByVal lv_no%) As Boolean
Dim i%
Dim lv As line_value_data0_type
Dim no%
If m_Circ(c%).data(0).data0.center > 0 And _
    m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.visible > 0 Then
If m_Circ(c%).data(0).radii_no > 0 Then
 is_kwon_radii_for_circle = True
  Exit Function
Else
If lv_no% = 0 Then
For i% = 1 To m_Circ(c%).data(0).data0.in_point(0)
  If is_line_value(m_Circ(c%).data(0).data0.center, _
                m_Circ(c%).data(0).data0.in_point(i%), 0, 0, 0, "", _
                  no%, -1000, 0, 0, 0, _
                   lv) > 0 Then
                   Call C_display_picture.set_circle_radii_no(c%, no%)
                    Exit Function
  End If
Next i%
Else
End If
End If
End If
End Function


Public Function is_old_conclusion(ByVal wenti_no%) As Boolean
Dim i%
For i% = 0 To 3
    If conclusion_data(i%).wenti_no = wenti_no% Then
      is_old_conclusion = True
    Exit Function
End If
Next i%
End Function

Public Function is_two_circle_3point_same(c1 As circle_data0_type, c2 As circle_data0_type) As Boolean
Dim i%, j%, k%
If c1.in_point(0) >= 3 And c2.in_point(0) >= 3 Then
For i% = 1 To c1.in_point(0)
 For j% = 1 To c2.in_point(0)
  If c1.in_point(i%) = c2.in_point(j%) Then
     k% = k% + 1
      If k% = 3 Then
         is_two_circle_3point_same = True
      End If
  End If
 Next j%
Next i%
End If
End Function
Public Function is_two_circle_1point_same(c1 As circle_data0_type, c2 As circle_data0_type) As Boolean
Dim i%, j%
For i% = 1 To c1.in_point(0)
 For j% = 1 To c2.in_point(0)
  If c1.in_point(i%) = c2.in_point(j%) Then
         is_two_circle_1point_same = True
          Exit Function
  End If
 Next j%
Next i%
End Function

Public Sub is_conclusion_no(ByVal ty As Byte, n() As Integer)
Dim i%
For i% = 0 To last_conclusion - 1
   If conclusion_data(i%).ty = ty And conclusion_data(i%).no(0) = 0 Then
      n(0) = i%
   Else
      n(0) = -1
   End If
Next i%
End Sub
Public Function is_there_diameter_in_circle(ByVal c%) As Boolean
Dim i%, j%
Dim c_data0 As condition_data_type
For i% = 2 To m_Circ(c%).data(0).data0.in_point(0)
 For j% = 1 To j% - 1
  If is_three_point_on_line(m_Circ(c%).data(0).data0.in_point(i%), _
       m_Circ(c%).data(0).data0.center, m_Circ(c%).data(0).data0.in_point(j%), _
        0, 0, 0, 0, c_data0, 0, 0, 0) = 1 Then
   is_there_diameter_in_circle = True
     Exit Function
   End If
 Next j%
Next i%
End Function
Public Function is_initial_data_for_general_string(ByVal g_no%) As Boolean
Dim i%
is_initial_data_for_general_string = True
For i% = 0 To 3
  If general_string(g_no%).data(0).item(i%) > 0 Then
     If is_initial_data_for_item(general_string(g_no%).data(0).item(i%)) = False Then
       is_initial_data_for_general_string = False
         Exit Function
     End If
  End If
Next i%
End Function
Public Function is_initial_data_for_item(ByVal it%) As Boolean
 If is_initial_data_for_element(item0(it%).data(0).poi(0), _
                                     item0(it%).data(0).poi(1)) Then
  If is_initial_data_for_element(item0(it%).data(0).poi(2), _
                                      item0(it%).data(0).poi(3)) Then
        is_initial_data_for_item = True
  End If
 End If
End Function
Public Function is_initial_data_for_element(ByVal p1%, ByVal p2%) As Boolean
 If p1% = 0 And p2% = 0 Then
  is_initial_data_for_element = True
 Else
  If p2% > 0 Then
     If m_poi(p1%).data(0).degree_for_reduce = 0 And _
         m_poi(p2%).data(0).degree_for_reduce = 0 Then
         is_initial_data_for_element = True
          Exit Function
     End If
  Else
     If m_poi(angle(p1%).data(0).poi(0)).data(0).degree_for_reduce = 0 And _
          m_poi(angle(p1%).data(0).poi(1)).data(0).degree_for_reduce = 0 And _
           m_poi(angle(p1%).data(0).poi(2)).data(0).degree_for_reduce = 0 Then
         is_initial_data_for_element = True
          Exit Function
     End If
  End If
 End If
End Function

Public Function is_paral_v_line(vl1%, vl2%, re_v As String) As Boolean
If is_dparal(Dtwo_point_line(V_line_value(vl2%).data(0).v_line).data(0).line_no, _
       Dtwo_point_line(V_line_value(vl2%).data(0).v_line).data(0).line_no, 0, -1000, _
         0, 0, 0, 0) Then
       is_paral_v_line = False
Else
If V_line_value(vl2%).data(0).value <> "0" Then
   re_v = divide_string(V_line_value(vl1%).data(0).value, _
            V_line_value(vl2%).data(0).value, True, False)
   If re_v <> "F" Then
          is_paral_v_line = True
   End If
ElseIf V_line_value(vl2%).data(0).value <> "0" Then
   re_v = divide_string(V_line_value(vl1%).data(0).value, _
         V_line_value(vl2%).data(0).value, True, False)
   If re_v <> "F" Then
          is_paral_v_line = True
   End If
End If
End If
End Function
Public Function simple_general_string0(gs As general_string_data_type) As Boolean
Dim i%

End Function
Public Function simple_two_item0(ByVal it1%, ByVal it2%, ByVal para1$, ByVal para2$, _
                      out_it1%, out_it2%, out_para1$, out_para2$) As Boolean
out_it1% = it1%
out_it2% = it2%
out_para1$ = para1$
out_para2$ = para2$

End Function
Public Function is_v_relation(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                       re_v As String, no%, cond_ty As Byte, _
                        v_re_data As V_relation_data0_type, _
                         c_data As condition_data_type) As Boolean
Dim v_l(2) As Integer
Dim dir(1) As String
Dim dr As relation_data0_type
Dim c_data_  As condition_data_type
v_l(0) = vector_number(p1%, p2%, dir(0))
v_l(1) = vector_number(p3%, p4%, dir(1))
dir(0) = dir(0) * dir(1)
re_v = time_string(re_v, dir(0), True, False)
If Dtwo_point_line(v_l(0)).data(0).line_no <> Dtwo_point_line(v_l(1)).data(0).line_no Then
 If v_l(0) > v_l(1) Then
  Call exchange_two_integer(v_l(0), v_l(1))
  Call exchange_string(dir(0), dir(1))
  re_v = divide_string("1", re_v, True, False)
 End If
Else
 Call is_relation(Dtwo_point_line(v_l(0)).poi(0), Dtwo_point_line(v_l(0)).poi(1), _
         Dtwo_point_line(v_l(1)).poi(0), Dtwo_point_line(v_l(1)).poi(1), _
          Dtwo_point_line(v_l(0)).data(0).n(0), Dtwo_point_line(v_l(0)).data(0).n(1), _
           Dtwo_point_line(v_l(1)).data(0).n(0), Dtwo_point_line(v_l(1)).data(0).n(1), _
            Dtwo_point_line(v_l(0)).data(0).line_no, Dtwo_point_line(v_l(1)).data(0).line_no, _
             re_v, 0, 0, 0, 0, 0, dr, 0, 0, 0, c_data_, 0)
v_l(0) = vector_number(dr.poi(0), dr.poi(1), dir(0))
v_l(1) = vector_number(dr.poi(2), dr.poi(3), dir(1))
v_l(2) = vector_number(dr.poi(4), dr.poi(5), 0)
re_v = dr.value
End If
v_re_data.v_line(0) = v_l(0)
v_re_data.v_line(1) = v_l(1)
v_re_data.v_line(2) = v_l(2)
v_re_data.value = re_v
If Dtwo_point_line(v_l(0)).data(0).value <> "" And Dtwo_point_line(v_l(1)).data(0).value <> "" Then
   is_v_relation = True
    Exit Function
ElseIf Dtwo_point_line(v_l(0)).data(0).value <> "" Then
    v_re_data.v_line(0) = v_l(1)
    v_re_data.value = divide_string(Dtwo_point_line(v_l(0)).data(0).value, _
          v_re_data.value, True, False)
    cond_ty = V_line_value_
    Call add_conditions_to_record(V_line_value_, Dtwo_point_line(v_l(0)).data(0).v_line_value_no, _
          0, 0, c_data)
    Exit Function
ElseIf Dtwo_point_line(v_l(1)).data(0).value <> "" Then
    v_re_data.value = time_string(Dtwo_point_line(v_l(1)).data(0).value, _
          v_re_data.value, True, False)
    cond_ty = V_line_value_
    Call add_conditions_to_record(V_line_value_, Dtwo_point_line(v_l(0)).data(0).v_line_value_no, _
          0, 0, c_data)
    Exit Function
ElseIf Dtwo_point_line(v_l(2)).data(0).value <> "" And v_l(2) > 0 Then
    v_re_data.v_line(0) = v_l(1)
    v_re_data.value = divide_string(Dtwo_point_line(v_l(2)).data(0).value, _
          add_string("1", v_re_data.value, False, False), True, False)
    cond_ty = V_line_value_
    Call add_conditions_to_record(V_line_value_, Dtwo_point_line(v_l(0)).data(0).v_line_value_no, _
          0, 0, c_data)
    Exit Function
End If
For no% = 1 To last_conditions.last_cond(1).v_relation_no
 If v_Drelation(no%).data(0).data0.v_line(0) = v_l(0) And _
     v_Drelation(no%).data(0).data0.v_line(1) = v_l(1) Then
      is_v_relation = True
       re_v = time_string(v_Drelation(no%).data(0).data0.value, dir(0), True, False)
        v_re_data = v_Drelation(no%).data(0).data0
         Exit Function
 End If
Next no%
cond_ty = v_relation_
no% = 0

End Function
Public Function is_verti_v_line(ByVal vl0 As Integer, ByVal vl1 As Integer, _
          c_data As condition_data_type) As Boolean
Dim ts$
Dim i%
If is_dverti(Dtwo_point_line(V_line_value(vl0).data(0).v_line).data(0).line_no, _
                  Dtwo_point_line(V_line_value(vl1).data(0).v_line).data(0).line_no, _
                   0, -1000, 0, 0, 0, 0) Then
                     is_verti_v_line = False
Else
  ts$ = time_string(Dtwo_point_line(V_line_value(vl0).data(0).v_line).data(0).v_value, _
           Dtwo_point_line(V_line_value(vl1).data(0).v_line).data(0).v_value, _
                     True, False)
     Call add_conditions_to_record(0, 0, 0, 0, c_data)
     If ts$ = "0" Then
      is_verti_v_line = True
        Exit Function
     End If
End If
End Function

Public Function is_same_two_pair_condition(ele11 As condition_type, _
                  ele12 As condition_type, ele21 As condition_type, _
                     ele22 As condition_type) As Boolean
If is_same_condition(ele11, ele21) Or is_same_condition(ele11, ele22) Or _
     is_same_condition(ele12, ele21) Or is_same_condtion(ele12, ele22) Then
      is_same two + pair_condition = True
 End If
End Function

Public Function is_same_condition(ele1 As condition_type, ele2 As condition_type)
  If elel.ty = ele2.ty And ele1.no = ele2.no Then
      is_same_condition = True
  Else
      is_same_condition = False
  End If
End Function

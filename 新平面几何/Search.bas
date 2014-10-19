Attribute VB_Name = "search"
Option Explicit
Type element_from_data_type
  element As condition_type
  data0 As condition_data_type
End Type
Type element_from_data
  last_element As Integer
  data(16) As element_from_data_type
End Type
Public Function find_same_point_from_two_points(p1() As Integer, p2() As Integer, same_points() As Integer, p_no As Integer) As Integer    '两个序列中找到唯一的相同点
Dim i%, j%
Dim temp_p%
Dim t_points1() As Integer
Dim t_points2() As Integer
t_points1 = p1
t_points2 = p2
same_points(0) = 0
If t_points1(0) > 0 Then '点在圆周上
'************************合并前两个序列
 For i% = 1 To t_points1(0)
  For j% = 1 To t_points2(0)
   If t_points1(i%) = t_points2(j%) Then
    same_points(0) = same_points(0) + 1
     same_points(same_points(0)) = t_points1(i%)
   End If
  Next j%
 Next i%
'*************************
Else '否则，点在圆心，在same_points中选择圆
If p_no > 0 Then
For i% = 1 To t_points2(0)
 If m_Circ(t_points2(i%)).data(0).data0.center = p_no% Then
    same_points(0) = same_points(0) + 1
     same_points(same_points(0)) = t_points2(i%)
 End If
Next i%
Else
 same_points = t_points2
End If
End If
    find_same_point_from_two_points = same_points(0)
End Function
Public Function find_point_from_points(p1 As Integer, p2() As Integer) As Boolean '两个序列中找到唯一的相同点
Dim i%
 For i% = 1 To p2(0)
  If p1 = p2(i%) Then '找到与p1相同的点
     find_point_from_points = True
      Exit Function
  End If
 Next i%
End Function

Public Function search_for_line_number_from_two_point(ByVal p1%, ByVal p2%, n1%, n2%, _
          Optional index As Integer = 0) As Integer
'确定过两点的直线 输出直线序号和p1p2在直线上的位置,不建立新线,仅供推理
Dim i%, k%, tn%
Dim temp_record As record_data_type
If p1% = 0 Or p2% = 0 Or p1% = p2% Then 'p1,p2均为空 直线序号=0
search_for_line_number_from_two_point = 0
  Exit Function
'ElseIf p2% = 0 Or p1% = p2% Then 'p2=0 或p2 =p1>0 同一点,但不包含p1=p2=0,
' For i% = last_conditions.last_cond(1).line_no To 1 Step -1
'     If m_lin(i%).data(0).data0.in_point(0) = 1 And _
'         m_lin(i%).data(0).data0.in_point(1) = p1% Then
'           line_number_ = i%
'            Exit Function
'     End If
' Next i%
Else
 If search_for_two_point_line(p1%, p2%, tn%, 0) Then '搜索过p1,p2的直线
  search_for_line_number_from_two_point = read_other_line(Dtwo_point_line(tn%).data(0).line_no) '输出直线序号
 If search_for_line_number_from_two_point > 0 Then '如果直线>0,搜索p1p2在zhixian上的位置
  For i% = 1 To m_lin(search_for_line_number_from_two_point).data(0).data0.in_point(0)
  If p1% = m_lin(search_for_line_number_from_two_point).data(0).data0.in_point(i%) Then
   n1% = i%
  ElseIf p2% = m_lin(search_for_line_number_from_two_point).data(0).data0.in_point(i%) Then
   n2% = i%
  End If
  Next i%
 End If
End If
End If
End Function

Public Function find_conclusion(ByVal n%, con_type As Byte, _
           ByVal con_no%, is_set_reduce As Boolean) As Byte
Dim i%, j%
Dim is_x As Boolean
Dim no(3) As Integer
Dim n_(8) As Integer
Dim no1%
Dim ts$
Dim dn(2) As Integer
Dim tp(5) As Integer
Dim tl(3) As Integer
Dim tA(2) As Integer
Dim tn(7) As Integer
Dim t_l As tangent_line_data_type
Dim temp_record As total_record_type
Dim gs As general_string_type
Dim con_ty(1) As Byte
'Dim p As polygon
Dim poly4_no%
Dim lv_data0 As line_value_data0_type
Dim l2v_data0 As two_line_value_data0_type
Dim l3v_data0 As line3_value_data0_type
Dim c_data As condition_data_type
Dim dr_data As relation_data0_type
Dim ele As condition_type
If prove_or_set_dbase = True Then
  find_conclusion = 0
   Exit Function
End If
If conclusion_data(n%).ty = 0 Then
 find_conclusion = 0
  Exit Function
End If
Select Case conclusion_data(n%).ty
Case equation_
   If con_type = equation_ Then
    If equation(con_no%).data(0).root(0) <> "" Then
      conclusion_data(n%).no(0) = con_no%
       find_conclusion = 1
    End If
   End If
Case V_line_value_
 If con_type = V_line_value_ Then
      If con_V_line_value(n%).data(0).v_line = V_line_value(con_no%).data(0).v_line Then
         conclusion_data(n%).no(0) = con_no%
          find_conclusion = 1
      End If
 ElseIf con_type = 0 And is_set_reduce Then
   Call set_conclusion_point(n%, con_V_line_value(n%).data(0).v_poi(0))
   Call set_conclusion_point(n%, con_V_line_value(n%).data(0).v_poi(1))
   'Call set_condition_reduce(point_, con_V_line_value(n%).data(0).v_poi(0), 0, n% + 1)
   'Call set_condition_reduce(point_, con_V_line_value(n%).data(0).v_poi(1), 0, n% + 1)
 End If
Case area_of_element_
If con_type = area_of_element_ Then
 If is_contain_x(area_of_element(con_no%).data(0).value_, "x", 1) = False Then  '不含未知数
  If con_Area_of_element(n%).data(0).element.ty = area_of_element(con_no%).data(0).element.ty And _
     con_Area_of_element(n%).data(0).element.no = area_of_element(con_no%).data(0).element.no Then
   If area_of_element(con_no%).data(0).value = area_of_element(con_no%).data(0).value_ Then '未设未知数
        conclusion_data(n%).no(0) = con_no%
   Else
        conclusion_data(n%).no(0) = -con_no% '设有未知数
   End If
   area_of_element(con_no%).data(0).record.data1.is_proved = 3
    find_conclusion = 1
  End If
 Else '
  If con_Area_of_element(n%).data(0).element.ty = area_of_element(con_no%).data(0).element.ty And _
      con_Area_of_element(n%).data(0).element.no = area_of_element(con_no%).data(0).element.no Then
        conclusion_data(n%).no(1) = con_no% '
  End If
 End If
ElseIf con_type = 0 Then
 If is_set_reduce Then
  Call set_conclusion_point_for_area_element(n%, con_Area_of_element(n%).data(0).element)
  'Call set_area_of_element_reduce(con_Area_of_element(n%).data(0).element, n% + 1)
 End If
 If is_area_of_element(con_Area_of_element(n%).data(0).element.ty, _
        con_Area_of_element(n%).data(0).element.no, no(0), -1000) Then
  If is_contain_x(area_of_element(no(0)).data(0).value_, "x", 1) = False Then
   If area_of_element(no(0)).data(0).value = area_of_element(no(0)).data(0).value_ Then
    conclusion_data(n%).no(0) = no(0)
   Else
    conclusion_data(n%).no(0) = -no(0)
   End If
   line_value(no(0)).data(0).record.data1.is_proved = 3
    find_conclusion = 1
  End If
 End If
End If
Case sides_length_of_triangle_
If con_type = 0 Then
     If is_set_reduce Then
      ele.ty = triangle_
      ele.no = con_Sides_length_of_triangle(n%).data(0).triangle
      Call set_conclusion_point_for_area_element(n%, ele)
      'Call set_area_of_element_reduce(ele, n% + 1)
     End If
     ts$ = ""
     temp_record.record_data.data0.condition_data.condition_no = 0
     If is_sides_length_of_triangle(con_Sides_length_of_triangle(n%).data(0).triangle, _
      no(0)) Then
      If is_contain_x(Sides_length_of_triangle(no(0)).data(0).value_, "x", 1) = False Then
      conclusion_data(n%).no(0) = no(0)
      find_conclusion = 1
      Exit Function
      End If
     End If
 'ElseIf con_type = line_value_ Then
  '  if is_same_two_point(line_value(con_no%).data(0).data0(
 'ElseIf con_type = line2_value_ Then
 ElseIf con_type = line3_value_ Then
    If is_three_line_value(triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(0), _
              triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(1), _
               triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(1), _
                triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(2), _
                 triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(2), _
                  triangle(con_Sides_length_of_triangle(n%).data(0).triangle).data(0).poi(0), _
                   0, 0, 0, 0, 0, 0, 0, 0, 0, "1", "1", "1", ts$, dn(0), -1000, 0, 0, 0, 0, 0, l3v_data0, 0, _
                     temp_record.record_data.data0.condition_data, 0) = 1 Then
     If dn(0) > 0 Then
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = line3_value_
        temp_record.record_data.data0.condition_data.condition(1).no = dn(0)
         ts$ = line3_value(dn(0)).data(0).data0.value
     End If
     dn(0) = 0
      Call set_sides_length_of_triangle(con_Sides_length_of_triangle(n%).data(0).triangle, ts$, dn(0), temp_record, 0)
             conclusion_data(n%).no(0) = dn(0)
      find_conclusion = 1
      Exit Function
    End If
ElseIf con_type = sides_length_of_triangle_ Then
    If con_Sides_length_of_triangle(n%).data(0).triangle = _
                 Sides_length_of_triangle(con_no%).data(0).triangle Then
      If is_contain_x(Sides_length_of_triangle(con_no%).data(0).value_, "x", 1) = False Then
      conclusion_data(n%).no(0) = con_no%
      find_conclusion = 1
      Exit Function
      End If
    End If
End If
Case sides_length_of_circle_
If con_type = 0 Then
    If is_set_reduce Then
     Call set_conclusion_point_for_circle(n, con_Sides_length_of_circle(n%).data(0).circ)
     'Call set_condition_reduce(circle_, con_Sides_length_of_circle(n%).data(0).circ, 0, n% + 1)
    End If
    If is_sides_length_of_circle(con_Sides_length_of_circle(n%).data(0).circ, _
        no(0)) = True Then
     If is_contain_x(Sides_length_of_circle(no(0)).data(0).value_, "x", 1) = False Then
     conclusion_data(n%).no(0) = no(0)
     find_conclusion = 1
     Exit Function
     End If
    End If
ElseIf con_type = sides_length_of_circle_ Then
    If con_Sides_length_of_circle(n%).data(0).circ = Sides_length_of_circle(con_no%).data(0).circ Then
     If is_contain_x(Sides_length_of_circle(con_no%).data(0).value_, "x", 1) = False Then
     conclusion_data(n%).no(0) = con_no%
     find_conclusion = 1
     Exit Function
     End If
    End If
End If
Case area_of_circle_
If con_type = 0 Then
      If is_set_reduce Then
        Call set_conclusion_point_for_circle(n%, con_Area_of_circle(n%).data(0).circ)
        'Call set_condition_reduce(circle_, con_Area_of_circle(n%).data(0).circ, 0, n% + 1)
      End If
    If is_area_of_circle(con_Area_of_circle(n%).data(0).circ, no(0)) Then
     If is_contain_x(area_of_circle(no(0)).data(0).value_, "x", 1) = False Then
     conclusion_data(n%).no(0) = no(0)
     find_conclusion = 1
     Exit Function
     End If
    End If
ElseIf con_type = area_of_circle_ Then
   If con_Area_of_circle(n%).data(0).circ = area_of_circle(con_no%).data(0).circ Then
     If is_contain_x(area_of_circle(con_no%).data(0).value_, "x", 1) = False Then
      If con_Area_of_circle(n%).data(0).value <> "" And _
          con_Area_of_circle(n%).data(0).value <> _
           area_of_circle(con_no%).data(0).value Then
            error_of_wenti = 2
      Else
       conclusion_data(n%).no(0) = con_no%
      End If
     find_conclusion = 1
     Exit Function
     End If
   End If
End If
Case area_of_fan_
If con_type = 0 Then
  If is_set_reduce Then
   Call set_conclusion_point(n%, con_Area_of_fan(n%).data(0).poi(0))
   Call set_conclusion_point(n%, con_Area_of_fan(n%).data(0).poi(1))
   Call set_conclusion_point(n%, con_Area_of_fan(n%).data(0).poi(2))
   'Call set_condition_reduce(point_, con_Area_of_fan(n%).data(0).poi(0), 0, n% + 1)
   'Call set_condition_reduce(point_, con_Area_of_fan(n%).data(0).poi(1), 0, n% + 1)
   'Call set_condition_reduce(point_, con_Area_of_fan(n%).data(0).poi(2), 0, n% + 1)
  End If
  If is_area_of_fan(con_Area_of_fan(n%).data(0).poi(0), _
     con_Area_of_fan(n%).data(0).poi(1), _
       con_Area_of_fan(n%).data(0).poi(2), no(0)) Then
     If is_contain_x(Area_of_fan(no(0)).data(0).value_, "x", 1) = False Then
     conclusion_data(n%).no(0) = no(0)
     find_conclusion = 1
     End If
     Exit Function
   End If
ElseIf con_type = area_of_fan_ Then
 If con_Area_of_fan(n%).data(0).poi(0) = Area_of_fan(con_no%).data(0).poi(0) And _
     con_Area_of_fan(n%).data(0).poi(1) = Area_of_fan(con_no%).data(0).poi(1) And _
      con_Area_of_fan(n%).data(0).poi(2) = Area_of_fan(con_no%).data(0).poi(2) Then
     If is_contain_x(Area_of_fan(con_no%).data(0).value_, "x", 1) = False Then
      If con_Area_of_fan(n%).data(0).value <> "" And _
           Area_of_fan(n%).data(0).value <> _
           con_Area_of_fan(con_no%).data(0).value Then
            error_of_wenti = 2
      Else
       conclusion_data(n%).no(0) = con_no%
      End If
     find_conclusion = 1
     Exit Function
     End If
 End If
End If
Case angle3_value_
If con_type = angle3_value_ Or con_type = 0 Then
 If con_type = 0 Then
  If is_set_reduce Then
   Call set_conclusion_point_for_angle(n%, con_angle3_value(n%).data(0).data0.angle(0))
   Call set_conclusion_point_for_angle(n%, con_angle3_value(n%).data(0).data0.angle(1))
   Call set_conclusion_point_for_angle(n%, con_angle3_value(n%).data(0).data0.angle(2))
   'Call set_condition_reduce(angle_, con_angle3_value(n%).data(0).data0.angle(0), 0, n% + 1)
   'Call set_condition_reduce(angle_, con_angle3_value(n%).data(0).data0.angle(1), 0, n% + 1)
   'Call set_condition_reduce(angle_, con_angle3_value(n%).data(0).data0.angle(2), 0, n% + 1)
  End If
 End If
 record_0.data0.condition_data.condition_no = 0
 ts$ = con_angle3_value(n%).data(0).data0.value
  If is_three_angle_value(con_angle3_value(n%).data(0).data0.angle(0), _
   con_angle3_value(n%).data(0).data0.angle(1), con_angle3_value(n%).data(0).data0.angle(2), _
    con_angle3_value(n%).data(0).data0.para(0), con_angle3_value(n%).data(0).data0.para(1), _
      con_angle3_value(n%).data(0).data0.para(2), con_angle3_value(n%).data(0).data0.value, _
       con_angle3_value(n%).data(0).data0.value, no(0), dn(0), dn(1), -1000, 0, 0, 0, 0, 0, 0, 0, angle3_value_data0, _
        temp_record.record_data.data0.condition_data, 10) Then
         find_conclusion = 1
          If temp_record.record_data.data0.condition_data.condition_no = 1 Then
           no(0) = temp_record.record_data.data0.condition_data.condition(1).no
           dn(0) = 0
           dn(1) = 0
          ElseIf temp_record.record_data.data0.condition_data.condition_no = 2 Then
           no(0) = temp_record.record_data.data0.condition_data.condition(1).no
           dn(0) = temp_record.record_data.data0.condition_data.condition(2).no
           dn(1) = 0
          ElseIf temp_record.record_data.data0.condition_data.condition_no = 3 Then
           no(0) = temp_record.record_data.data0.condition_data.condition(1).no
           dn(0) = temp_record.record_data.data0.condition_data.condition(2).no
           dn(1) = temp_record.record_data.data0.condition_data.condition(2).no
         End If
    If dn(0) = 0 And dn(1) = 0 Then
       If is_contain_x(angle3_value(no(0)).data(0).data0.value_, "x", 1) = False Then
          If con_angle3_value(n%).data(1).data0.value = "y" Then
             con_angle3_value(n%).data(1).data0.value = _
              solve_general_equation(con_angle3_value(n%).data(0).data0.value, _
                   con_angle3_value(n%).data(0).data0.value_, "y")
          Else
              con_angle3_value(n%).data(1).data0.value = _
                  con_angle3_value(n%).data(0).data0.value
              con_angle3_value(n%).data(0).data0.value_ = _
                  con_angle3_value(n%).data(0).data0.value
          End If
         If (con_angle3_value(n%).data(0).data0.angle(0) = angle3_value(no(0)).data(0).data0.angle(0) And _
            con_angle3_value(n%).data(0).data0.angle(1) = angle3_value(no(0)).data(0).data0.angle(1) And _
            con_angle3_value(n%).data(0).data0.angle(2) = angle3_value(no(0)).data(0).data0.angle(2) And _
            con_angle3_value(n%).data(0).data0.para(0) = angle3_value(no(0)).data(0).data0.para(0) And _
            con_angle3_value(n%).data(0).data0.para(1) = angle3_value(no(0)).data(0).data0.para(1) And _
            con_angle3_value(n%).data(0).data0.para(2) = angle3_value(no(0)).data(0).data0.para(2)) And _
              con_angle3_value(n%).data(0).record.data0.condition_data.condition_no = 0 Then
            conclusion_data(n%).no(0) = no(0)
         Else
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 0, n_(0), 1) '5.7
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 1, n_(1), 1) '5.7
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 2, n_(2), 1)
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 3, n_(3), 1)
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 4, n_(4), 1)
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 5, n_(5), 1)
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 6, n_(6), 1)
       Call search_for_three_angle_value(con_angle3_value(n%).data(1).data0, 7, n_(7), 1)
       If last_conditions.last_cond(1).angle3_value_no Mod 10 = 0 Then
       ReDim Preserve angle3_value(last_conditions.last_cond(1).angle3_value_no) As angle3_value_type
       End If
       last_conditions.last_cond(1).angle3_value_no = last_conditions.last_cond(1).angle3_value_no + 1
        conclusion_data(n%).no(0) = last_conditions.last_cond(1).angle3_value_no
         If con_angle3_value(n%).data(0).record.data0.condition_data.condition_no > 0 Then '共线,化简结论
          Call add_record_to_record(con_angle3_value(n%).data(0).record.data0.condition_data, _
                  temp_record.record_data.data0.condition_data)
         End If
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0).record = temp_record.record_data
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0).data0 = con_angle3_value(n%).data(1).data0
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0).record = temp_record.record_data
       For i% = 0 To 6
        For j% = last_conditions.last_cond(1).angle3_value_no To n_(i%) + 2 Step -1
         angle3_value(j%).data(0).record.data1.index.i(i%) = _
          angle3_value(j% - 1).data(0).record.data1.index.i(i%)
        Next j%
         angle3_value(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).angle3_value_no
       Next i%
        'conclusion_no(n%,0) = last_conditions.last_cond(1).angle3_value_no
      End If
         find_conclusion = 1
          Exit Function
    'Else '
     '   conclusion_no(n%,0) = no(0)
     '    find_conclusion = 1
     '     Exit Function
     '   End If
     '  Else
      '   find_conclusion = 0
      '    Exit Function
      End If
    Else '"x"
     If search_for_three_angle_value(angle3_value_data0, 0, n_(0), 1) Then '5.7
        If is_contain_x(angle3_value(angle3_value(n_(0) + 1).data(0).record.data1.index.i(0)).data(0).data0.value_, _
           "x", 1) = False Then
        conclusion_data(n%).no(0) = angle3_value(n_(0) + 1).data(0).record.data1.index.i(0)
        Else
        conclusion_data(n%).no(1) = angle3_value(n_(0) + 1).data(0).record.data1.index.i(0)
         find_conclusion = 0
          Exit Function
        End If
     Else
      If is_contain_x(angle3_value(no(0)).data(0).data0.value_, "x", 1) Or _
          is_contain_x(angle3_value(dn(0)).data(0).data0.value_, "x", 1) Or _
           is_contain_x(angle3_value(dn(1)).data(0).data0.value_, "x", 1) Then
            find_conclusion = 0
             Exit Function
      Else
       Call search_for_three_angle_value(angle3_value_data0, 1, n_(1), 1) '5.7
       Call search_for_three_angle_value(angle3_value_data0, 2, n_(2), 1)
       Call search_for_three_angle_value(angle3_value_data0, 3, n_(3), 1)
       Call search_for_three_angle_value(angle3_value_data0, 4, n_(4), 1)
       Call search_for_three_angle_value(angle3_value_data0, 5, n_(5), 1)
       Call search_for_three_angle_value(angle3_value_data0, 6, n_(6), 1)
       Call search_for_three_angle_value(angle3_value_data0, 7, n_(7), 1)
        temp_record.record_data.data0.condition_data.condition_no = 0
          temp_record.record_data.data0.theorem_no = 1
       Call add_conditions_to_record(angle3_value_, no(0), dn(0), dn(1), _
        temp_record.record_data.data0.condition_data)
         Call set_level(temp_record.record_data.data0.condition_data)
       If last_conditions.last_cond(1).angle3_value_no Mod 10 = 0 Then
       ReDim Preserve angle3_value(last_conditions.last_cond(1).angle3_value_no + 10) As angle3_value_type
       End If
       last_conditions.last_cond(1).angle3_value_no = last_conditions.last_cond(1).angle3_value_no + 1
        conclusion_data(n%).no(0) = last_conditions.last_cond(1).angle3_value_no
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0) = angle3_value_data_0
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0).data0 = angle3_value_data0
       angle3_value(last_conditions.last_cond(1).angle3_value_no).data(0).record = temp_record.record_data
       For i% = 0 To 6
        For j% = last_conditions.last_cond(1).angle3_value_no To n_(i%) + 2 Step -1
         angle3_value(j%).data(0).record.data1.index.i(i%) = _
          angle3_value(j% - 1).data(0).record.data1.index.i(i%)
        Next j%
         angle3_value(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).angle3_value_no
       Next i%
     End If
    End If
   End If
   End If
End If
'Case equal_area_triangle_
' If con_type = equal_area_triangle_ Then
'  If is_equal_area_triangle(con_equal_area_triangle(n%).data(0).triangle(0), _
'     con_equal_area_triangle(n%).data(0).triangle(1), dn(0), -1000, 0, 0, 0, 0, con_ty(0)) Then
'       conclusion_data(n%).no(0) = dn(0)
'  find_conclusion = 1
'   Exit Function
'  End If
' End If
Case epolygon_
 If con_type = epolygon_ Then
  If con_Epolygon(n%).data(0).no = epolygon(con_no%).data(0).no Then
  conclusion_data(n%).no(0) = con_no%
   find_conclusion = 1
   Exit Function
  End If
 ElseIf con_type = 0 Or con_type = angle3_value_ Or con_type = eline_ Or con_type = line_value_ Then
  find_conclusion = find_conclusion_for_epolygon(epolygon_, n%, con_type, con_no, is_set_reduce)
 ElseIf con_type = Squre Then
  If con_Epolygon(n%).data(0).no = Dsqure(con_no%).data(0).polygon4_no Then
  conclusion_data(n%).ty = Squre
  conclusion_data(n%).no(0) = con_no%
   find_conclusion = 1
   Exit Function
  End If
 End If
Case Squre
   If con_type = Squre Then
    If con_squre(n%).data(0).polygon4_no = Dsqure(con_no%).data(0).polygon4_no Then
     conclusion_data(n%).no(0) = con_no%
      find_conclusion = 1
    Exit Function
    End If
   ElseIf con_type = epolygon_ Then
    If epolygon(con_no%).data(0).p.total_v = 4 Then
       If epolygon(con_no%).data(0).no = con_squre(n%).data(0).polygon4_no Then
        conclusion_data(n%).ty = epolygon_
        conclusion_data(n%).no(0) = con_no%
       End If
    End If
   ElseIf con_type = 0 Or con_type = angle3_value_ Or con_type = eline_ Or con_type = line_value_ Or _
           con_type = rhombus_ Or con_type = long_squre_ Then
    find_conclusion = find_conclusion_for_epolygon(Squre, n%, con_type, con_no, is_set_reduce)
   End If
Case equal_side_tixing_
If con_type = equal_side_tixing_ Then
   If Dtixing(con_no%).data(0).poly4_no = con_Dtixing(n%).data(0).poly4_no Then
      conclusion_data(n%).no(0) = con_no%
       find_conclusion = 1
      Exit Function
   End If
ElseIf con_type = 0 Then
If is_set_reduce Then
Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(0))
Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(1))
Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(2))
Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(3))
'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(0), 0, n% + 1)
'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(1), 0, n% + 1)
'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(2), 0, n% + 1)
'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(3), 0, n% + 1)
End If
For i% = 1 To last_conditions.last_cond(1).tixing_no
If Dpolygon4(Dtixing(i%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
If con_Dtixing(n%).data(0).poi(0) = Dtixing(i%).data(0).poi(0) And _
     con_Dtixing(n%).data(0).poi(1) = Dtixing(i%).data(0).poi(1) And _
      con_Dtixing(n%).data(0).poi(2) = Dtixing(i%).data(0).poi(2) And _
       con_Dtixing(n%).data(0).poi(3) = Dtixing(i%).data(0).poi(3) Then
 conclusion_data(n%).no(0) = i%
  find_conclusion = 1
   Exit Function
End If
End If
Next i%
End If
Case equal_side_triangle_
  If con_type = 0 And is_set_reduce Then
    Call set_conclusion_point_for_triangle(n%, con_equal_side_triangle(n%).data(0).triangle)
    'Call set_triangle_reduce(con_equal_side_triangle(n%).data(0).triangle, 0, n% + 1)
  ElseIf con_type = eline_ Or con_type = line_value_ Then
    'Call read_triangle_element(con_equal_side_triangle(n%).data(0).triangle, con_equal_side_triangle(n%).data(0).direction, _
          tp(0), tp(1), tp(2), tA(0), tA(1), tA(2), 0, 0, 0)
          tp(0) = triangle(con_equal_side_triangle(n%).data(0).triangle).data(0).poi(0)
          tp(1) = triangle(con_equal_side_triangle(n%).data(0).triangle).data(0).poi(1)
          tp(2) = triangle(con_equal_side_triangle(n%).data(0).triangle).data(0).poi(2)
    If is_equal_dline(tp(0), tp(1), tp(0), tp(2), _
          0, 0, 0, 0, 0, 0, no(0), -1000, 0, 0, 0, _
         eline_data0, dn(0), dn(1), cond_type, "", record_0.data0.condition_data) Then
           con_equal_side_triangle(n%).data(0).direction = 1
           ' (con_equal_side_triangle(n%).data(0).direction - 1) Mod 3 + 1
    ElseIf is_equal_dline(tp(1), tp(0), tp(1), tp(2), _
          0, 0, 0, 0, 0, 0, no(0), -1000, 0, 0, 0, _
         eline_data0, dn(0), dn(1), cond_type, "", record_0.data0.condition_data) Then
         con_equal_side_triangle(n%).data(0).direction = 2
    ElseIf is_equal_dline(tp(2), tp(1), tp(2), tp(0), _
          0, 0, 0, 0, 0, 0, no(0), -1000, 0, 0, 0, _
         eline_data0, dn(0), dn(1), cond_type, "", record_0.data0.condition_data) Then
         con_equal_side_triangle(n%).data(0).direction = 3
    Else
     Exit Function
    End If
     temp_record.record_data.data0.condition_data.condition_no = 0
      temp_record.record_data.data0.theorem_no = 0
      Call add_conditions_to_record(cond_type, no(0), dn(0), dn(1), temp_record.record_data.data0.condition_data)
If last_conditions.last_cond(1).equal_side_triangle_no Mod 10 = 0 Then
 ReDim Preserve equal_side_triangle(last_conditions.last_cond(1).equal_side_triangle_no + 1) _
     As one_triangle_type
End If
 last_conditions.last_cond(1).equal_side_triangle_no = last_conditions.last_cond(1).equal_side_triangle_no + 1
  no(0) = last_conditions.last_cond(1).equal_side_triangle_no
' equal_side_triangle(no(0)).data(0) = one_triangle_data_0
  equal_side_triangle(no(0)).data(0) = con_equal_side_triangle(n%).data(0)
  Call set_level(temp_record.record_data.data0.condition_data)
  equal_side_triangle(no(0)).data(0).record = temp_record.record_data
'  equal_side_triangle(no(0)).record.other_no = no(0)
  conclusion_data(n%).no(0) = no(0)
   find_conclusion = 1
    Exit Function
 'End If
 End If
 If con_type = angle3_value_ Then
  If is_equal_angle(tA(1), tA(2), no(0), no(1)) Then
       Call add_conditions_to_record(angle3_value_, no(0), no(1), 0, temp_record.record_data.data0.condition_data)
 If last_conditions.last_cond(1).equal_side_triangle_no Mod 10 = 0 Then
   ReDim Preserve equal_side_triangle(last_conditions.last_cond(1).equal_side_triangle_no + 1) As one_triangle_type
 End If
  last_conditions.last_cond(1).equal_side_triangle_no = last_conditions.last_cond(1).equal_side_triangle_no + 1
     no(0) = last_conditions.last_cond(1).equal_side_triangle_no
   equal_side_triangle(no(0)).data(0) = one_triangle_data_0
     equal_side_triangle(no(0)).data(0).triangle = con_equal_side_triangle(n%).data(0).triangle
       equal_side_triangle(no(0)).data(0).direction = con_equal_side_triangle(n%).data(0).direction
      Call set_level(temp_record.record_data.data0.condition_data)
     equal_side_triangle(no(0)).data(0).record = temp_record.record_data
     'equal_side_triangle(no(0)).record.other_no = no(0)
  conclusion_data(n%).no(0) = no(0)
   find_conclusion = 1
    Exit Function
 End If
 End If
Case equal_side_right_triangle_
find_conclusion = find_conclusion_for_equal_side_right_triangle(n%, con_type, is_set_reduce)
Case general_string_
 If con_type = general_string_ Then
  If general_string(con_no%).record_.conclusion_no = n% + 1 Then
    If general_string(con_no%).data(0).value <> "" Then
     If is_contain_x(general_string(con_no%).data(0).value_, "x", 1) = False Then
      conclusion_data(n%).no(0) = con_no%
       find_conclusion = 1
     Else
      conclusion_data(n%).no(1) = con_no%
     End If
    ElseIf general_string(con_no%).record_.conclusion_ty = 73 Or _
            general_string(con_no%).record_.conclusion_ty = 75 Or _
             general_string(con_no%).record_.conclusion_ty = 76 Then
     If is_initial_data_for_general_string(con_no%) Then
            conclusion_data(n%).no(0) = con_no%
       find_conclusion = 1
     End If
    End If
  Else
   If general_string(con_no%).data(0).item(0) = con_general_string(n%).data(0).item(0) And _
       general_string(con_no%).data(0).item(1) = con_general_string(n%).data(0).item(1) And _
        general_string(con_no%).data(0).item(2) = con_general_string(n%).data(0).item(2) And _
         general_string(con_no%).data(0).item(3) = con_general_string(n%).data(0).item(3) And _
      general_string(con_no%).data(0).para(0) = con_general_string(n%).data(0).para(0) And _
       general_string(con_no%).data(0).para(1) = con_general_string(n%).data(0).para(1) And _
        general_string(con_no%).data(0).para(2) = con_general_string(n%).data(0).para(2) And _
         general_string(con_no%).data(0).para(3) = con_general_string(n%).data(0).para(3) And _
          general_string(con_no%).data(0).value = con_general_string(n%).data(0).value Then
     If is_contain_x(general_string(con_no%).data(0).value_, "x", 1) = False Then
           conclusion_data(n%).no(0) = con_no%
            find_conclusion = 1
             Exit Function
     Else
           conclusion_data(n%).no(1) = con_no%
     End If
   End If
  End If
 Else
  If con_type = 0 And is_set_reduce Then
     Call set_conclusion_point_for_item(n%, con_general_string(n%).data(0).item(0))
     Call set_conclusion_point_for_item(n%, con_general_string(n%).data(0).item(1))
     Call set_conclusion_point_for_item(n%, con_general_string(n%).data(0).item(2))
     Call set_conclusion_point_for_item(n%, con_general_string(n%).data(0).item(3))
     'Call set_item_reduce(con_general_string(n%).data(0).item(0), 0, n% + 1)
     'Call set_item_reduce(con_general_string(n%).data(0).item(1), 0, n% + 1)
     'Call set_item_reduce(con_general_string(n%).data(0).item(2), 0, n% + 1)
     'Call set_item_reduce(con_general_string(n%).data(0).item(3), 0, n% + 1)
  End If
  If search_for_general_string(con_general_string(n%).data(0), 0, no1%, 0) Then
   If general_string(no1%).record_.conclusion_no = con_no% + 1 Then
    If general_string(no1%).data(0).value <> "" Or _
     general_string(no1%).data(0).value_ <> "" Then
     Call add_record_to_record(general_string(n%).data(0).record.data0.condition_data, general_string(no1%).data(0).record.data0.condition_data)
     If is_contain_x(general_string(no1%).data(0).value_, "x", 1) = False Then
       conclusion_data(n%).no(0) = con_no%
        find_conclusion = 1
     Else
       conclusion_data(n%).no(1) = con_no%
     End If
     End If
   End If
  End If
 End If
Case point3_on_line_
If con_type = point3_on_line_ Then
temp_record.record_data.data0.condition_data.condition_no = 0
If is_three_point_on_line(con_Three_point_on_line(n%).data(0).poi(0), _
  con_Three_point_on_line(n%).data(0).poi(1), _
   con_Three_point_on_line(n%).data(0).poi(2), no(0), -1000, 0, 0, _
        0, 0, 0) Then
 conclusion_data(n%).no(0) = no(0)
  three_point_on_line(no(0)).data(0).record.data1.is_proved = 3
   find_conclusion = 1
End If
ElseIf con_no% = 0 And is_set_reduce Then
 Call set_conclusion_point(n%, con_Three_point_on_line(n%).data(0).poi(0))
 Call set_conclusion_point(n%, con_Three_point_on_line(n%).data(0).poi(1))
 Call set_conclusion_point(n%, con_Three_point_on_line(n%).data(0).poi(2))
 'Call set_condition_reduce(point_, con_Three_point_on_line(n%).data(0).poi(0), 0, n% + 1)
 'Call set_condition_reduce(point_, con_Three_point_on_line(n%).data(0).poi(1), 0, n% + 1)
' Call set_condition_reduce(point_, con_Three_point_on_line(n%).data(0).poi(2), 0, n% + 1)
End If
Case point4_on_circle_
If con_type = point4_on_circle_ Or con_type = 0 Then
 If con_type = 0 And is_set_reduce Then
  Call set_conclusion_point(n%, con_Four_point_on_circle(n%).data(0).poi(0))
  Call set_conclusion_point(n%, con_Four_point_on_circle(n%).data(0).poi(1))
  Call set_conclusion_point(n%, con_Four_point_on_circle(n%).data(0).poi(2))
  Call set_conclusion_point(n%, con_Four_point_on_circle(n%).data(0).poi(3))
  'Call set_condition_reduce(point_, con_Four_point_on_circle(n%).data(0).poi(0), 0, n% + 1)
  'Call set_condition_reduce(point_, con_Four_point_on_circle(n%).data(0).poi(1), 0, n% + 1)
  'Call set_condition_reduce(point_, con_Four_point_on_circle(n%).data(0).poi(2), 0, n% + 1)
  'Call set_condition_reduce(point_, con_Four_point_on_circle(n%).data(0).poi(3), 0, n% + 1)
 End If
If is_four_point_on_circle(con_Four_point_on_circle(n%).data(0).poi(0), _
   con_Four_point_on_circle(n%).data(0).poi(1), _
    con_Four_point_on_circle(n%).data(0).poi(2), _
     con_Four_point_on_circle(n%).data(0).poi(3), no(0), p4_on_C, False) Then
 conclusion_data(n%).no(0) = no(0)
  four_point_on_circle(no(0)).data(0).record.data1.is_proved = 3
 find_conclusion = 1
End If
End If
Case paral_
If con_type = paral_ Or con_type = 0 Then
If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point_for_line(n%, con_paral(n%).data(0).line_no(0))
   Call set_conclusion_point_for_line(n%, con_paral(n%).data(0).line_no(1))
   'Call set_condition_reduce(line_, con_paral(n%).data(0).line_no(0), 0, n% + 1)
   'Call set_condition_reduce(line_, con_paral(n%).data(0).line_no(1), 0, n% + 1)
End If
If is_dparal(con_paral(n%).data(0).line_no(0), con_paral(n%).data(0).line_no(1), no(0), _
    -1000, 0, 0, 0, 0) Then
 Dparal(no(0)).data(0).data0.record.data1.is_proved = 3
  conclusion_data(n%).no(0) = no(0)
   find_conclusion = 1
End If
'ElseIf con_type = equal_area_triangle_ Then

End If
Case verti_
If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point_for_line(n%, con_verti(n%).data(0).line_no(0))
   Call set_conclusion_point_for_line(n%, con_verti(n%).data(0).line_no(1))
   'Call set_condition_reduce(line_, con_verti(n%).data(0).line_no(0), 0, n% + 1)
   'Call set_condition_reduce(line_, con_verti(n%).data(0).line_no(1), 0, n% + 1)
If is_dverti(con_verti(n%).data(0).line_no(0), con_verti(n%).data(0).line_no(1), no(0), _
      -1000, 0, 0, 0, 0) Then
 conclusion_data(n%).no(0) = no(0)
  Dverti(no(0)).data(0).record.data1.is_proved = 3
   find_conclusion = 1
End If
ElseIf con_type = verti_ Then
 If con_verti(n%).data(0).line_no(0) = Dverti(con_no%).data(0).line_no(0) And _
     con_verti(n%).data(0).line_no(1) = Dverti(con_no%).data(0).line_no(1) Then
 conclusion_data(n%).no(0) = con_no%
  Dverti(no(0)).data(0).record.data1.is_proved = 3
   find_conclusion = 1
 End If
ElseIf con_type = angle3_value_ Then
 If angle3_value(con_no%).data(0).data0.angle(1) = 0 And angle3_value(con_no%).data(0).data0.value = "90" Then
  If is_same_two_point(con_verti(n%).data(0).line_no(0), con_verti(n%).data(0).line_no(1), _
        angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0), _
         angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1)) Then
          temp_record.record_data = angle3_value(con_no%).data(0).record
   Call set_dverti(angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0), _
           angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1), temp_record, _
            conclusion_data(n%).no(0), 0, False)
             find_conclusion = 1
  End If
 End If
End If
Case relation_
If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point(n%, con_relation(n%).data(0).poi(0))
   Call set_conclusion_point(n%, con_relation(n%).data(0).poi(1))
   Call set_conclusion_point(n%, con_relation(n%).data(0).poi(2))
   Call set_conclusion_point(n%, con_relation(n%).data(0).poi(3))
   'Call set_condition_reduce(point_, con_relation(n%).data(0).poi(0), 0, n% + 1)
   'Call set_condition_reduce(point_, con_relation(n%).data(0).poi(1), 0, n% + 1)
   'Call set_condition_reduce(point_, con_relation(n%).data(0).poi(2), 0, n% + 1)
   'Call set_condition_reduce(point_, con_relation(n%).data(0).poi(3), 0, n% + 1)
ElseIf con_type = line_value_ Or con_type = midpoint_ Or con_type = eline_ Or _
   con_type = relation_ Then
   c_data.condition_no = 0
 If is_relation(con_relation(n%).data(0).poi(0), _
       con_relation(n%).data(0).poi(1), _
        con_relation(n%).data(0).poi(2), _
         con_relation(n%).data(0).poi(3), _
       con_relation(n%).data(0).n(0), _
        con_relation(n%).data(0).n(1), _
         con_relation(n%).data(0).n(2), _
          con_relation(n%).data(0).n(3), _
       con_relation(n%).data(0).line_no(0), _
        con_relation(n%).data(0).line_no(1), _
          con_relation(n%).data(0).value, no(0), -1000, 0, 0, 0, _
            relation_data0, dn(0), dn(1), cond_type, c_data, 1) Then
            Call ratio_value(relation_data0.value, _
                  con_relation(n%).data(0).ty, con_relation(n%).data(1).value)
     find_conclusion = 1
  If cond_type = relation_ Then
         Drelation(no(0)).data(0).record.data1.is_proved = 3
          Drelation(no(0)).record_.conclusion_no = n% + 1
           If is_contain_x(Drelation(no(0)).data(0).data0.value, "x", 1) = False Then
            If con_relation(n%).data(0).value <> "" And _
                con_relation(n%).data(0).value <> _
                 Drelation(no(0)).data(0).data0.value Then
             error_of_wenti = 2
            Else
             conclusion_data(n%).no(0) = Abs(no(0))
            End If
             Exit Function
           Else
            conclusion_data(n%).no(1) = Abs(no(0))
             Exit Function
           End If
  Else
     If search_for_relation(con_relation(n%).data(0), 0, n_(0), 1) Then '5.7
      If is_contain_x(Drelation(Drelation(n_(0) + 1).data(0).record.data1.index.i(0)).data(0).data0.value_, _
        "x", 1) = False Then
        conclusion_data(n%).no(0) = Drelation(n_(0) + 1).data(0).record.data1.index.i(0)
        If con_relation(n%).data(0).value <> "" And _
             con_relation(n%).data(0).value <> _
              Drelation(conclusion_data(n%).no(0)).data(0).data0.value_ Then
               error_of_wenti = 2
                conclusion_data(n%).no(0) = 0
        End If
      Else
        conclusion_data(n%).no(1) = Drelation(n_(0) + 1).data(0).record.data1.index.i(0)
            find_conclusion = 0
             Exit Function
      End If
     Else
       If no(0) > 0 And cond_type = relation_ Then
        is_x = is_contain_x(Drelation(no(0)).data(0).data0.value_, "x", 1)
       End If
       If dn(0) > 0 Then
        is_x = is_contain_x(line_value(dn(0)).data(0).data0.value_, "x", 1)
       End If
       If dn(1) > 0 Then
         is_x = is_contain_x(line_value(dn(1)).data(0).data0.value_, "x", 1)
      End If
       If (no(0) > 0 _
             Or (dn(0) > 0 And dn(1) > 0)) And is_x = False Then
        Call search_for_relation(con_relation(n%).data(0), 1, n_(1), 1) '5.7
        Call search_for_relation(con_relation(n%).data(0), 2, n_(2), 1)
        Call search_for_relation(con_relation(n%).data(0), 3, n_(3), 1)
      temp_record.record_data.data0.condition_data.condition_no = 0
      temp_record.record_data.data0.theorem_no = 1
  Call add_conditions_to_record(cond_type, no(0), dn(0), dn(1), temp_record.record_data.data0.condition_data)
  If last_conditions.last_cond(1).relation_no Mod 10 = 0 Then
  ReDim Preserve Drelation(last_conditions.last_cond(1).relation_no + 10) As relation_type
  End If
  last_conditions.last_cond(1).relation_no = last_conditions.last_cond(1).relation_no + 1
  Drelation(last_conditions.last_cond(1).relation_no).data(0) = relation_data_0
  Drelation(last_conditions.last_cond(1).relation_no).data(0).data0 = _
      con_relation(n%).data(0)
  Call set_level(temp_record.record_data.data0.condition_data)
  Drelation(last_conditions.last_cond(1).relation_no).data(0).record = temp_record.record_data
  'Drelation(last_conditions.last_cond(1).relation_no).record.other_no = last_conditions.last_cond(1).relation_no
   conclusion_data(n%).no(0) = last_conditions.last_cond(1).relation_no
   Drelation(last_conditions.last_cond(1).relation_no).record_.conclusion_no = n% + 1
  For i% = 0 To 7
  For j% = last_conditions.last_cond(1).relation_no To n_(i%) + 2 Step -1
   Drelation(j%).data(0).record.data1.index.i(i%) = Drelation(j% - 1).data(0).record.data1.index.i(i%)
  Next j%
   Drelation(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).relation_no
  Next i%
  End If
 End If
 End If
 End If
 End If
Case total_equal_triangle_
  If con_type = total_equal_triangle_ Or con_type = 0 Then
  If con_type = 0 And is_set_reduce Then
     Call set_conclusion_point_for_triangle(n%, con_total_equal_triangle(n%).data(0).triangle(0))
     Call set_conclusion_point_for_triangle(n%, con_total_equal_triangle(n%).data(0).triangle(1))
     'Call set_triangle_reduce(con_total_equal_triangle(n%).data(0).triangle(0), 0, n% + 1)
     'Call set_triangle_reduce(con_total_equal_triangle(n%).data(0).triangle(1), 0, n% + 1)
  End If
  If is_total_equal_Triangle(con_total_equal_triangle(n%).data(0).triangle(0), _
    con_total_equal_triangle(n%).data(0).triangle(1), 1, _
      con_total_equal_triangle(n%).data(0).direction, no(0), -1000, _
         0, 0, two_triangle0, record_0, 1) Then
     Dtotal_equal_triangle(no(0)).data(0).record.data1.is_proved = 3
      conclusion_data(n%).no(0) = no(0)
    find_conclusion = 1
  End If
  End If
Case similar_triangle_
  If con_type = similar_triangle_ Or con_type = 0 Then
    If con_type = 0 And is_set_reduce Then
     Call set_conclusion_point_for_triangle(n%, con_similar_triangle(n%).data(0).triangle(0))
     Call set_conclusion_point_for_triangle(n%, con_similar_triangle(n%).data(0).triangle(1))
     'Call set_triangle_reduce(con_similar_triangle(n%).data(0).triangle(0), 0, n% + 1)
     'Call set_triangle_reduce(con_similar_triangle(n%).data(0).triangle(1), 0, n% + 1)
    End If
   If is_similar_triangle0(con_similar_triangle(n%).data(0).triangle(0), _
    con_similar_triangle(n%).data(0).triangle(1), 1, _
     con_similar_triangle(n%).data(0).direction, no(0), -1000, 0, _
      0, two_triangle0, record_0, 0, 1) Then
     Dsimilar_triangle(no(0)).data(0).record.data1.is_proved = 3
      conclusion_data(n%).no(0) = no(0)
    find_conclusion = 1
  End If
  End If
Case eline_
If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point(n%, con_eline(n%).data(0).data0.poi(0))
   Call set_conclusion_point(n%, con_eline(n%).data(0).data0.poi(1))
   Call set_conclusion_point(n%, con_eline(n%).data(0).data0.poi(2))
   Call set_conclusion_point(n%, con_eline(n%).data(0).data0.poi(3))
'Call set_condition_reduce(point_, con_eline(n%).data(0).data0.poi(0), 0, n% + 1)
'Call set_condition_reduce(point_, con_eline(n%).data(0).data0.poi(1), 0, n% + 1)
'Call set_condition_reduce(point_, con_eline(n%).data(0).data0.poi(2), 0, n% + 1)
'Call set_condition_reduce(point_, con_eline(n%).data(0).data0.poi(3), 0, n% + 1)
ElseIf con_type = eline_ Or con_type = line_value_ Or _
     con_type = midpoint_ Or con_type = epolygon_ Then
temp_record.record_data.data0.condition_data.condition_no = 0
If is_equal_dline(con_eline(n%).data(0).data0.poi(0), con_eline(n%).data(0).data0.poi(1), _
    con_eline(n%).data(0).data0.poi(2), con_eline(n%).data(0).data0.poi(3), _
     con_eline(n%).data(0).data0.n(0), con_eline(n%).data(0).data0.n(1), _
      con_eline(n%).data(0).data0.n(2), con_eline(n%).data(0).data0.n(3), _
       con_eline(n%).data(0).data0.line_no(0), con_eline(n%).data(0).data0.line_no(1), _
       no(0), -1000, 0, 0, 0, eline_data0, dn(0), dn(1), cond_type, _
        "", temp_record.record_data.data0.condition_data) Then
   find_conclusion = 1
If cond_type = eline_ Or cond_type = midpoint_ Then
  conclusion_data(n%).no(0) = no(0)
ElseIf cond_type = midpoint_ Then
 conclusion_data(n%).ty = midpoint_
  conclusion_data(n%).no(0) = no(0)
   Dmid_point(no(0)).data(0).record.data1.is_proved = 3
Else
 If search_for_eline(con_eline(n%).data(0).data0, 0, n_(0), 1) Then
     conclusion_data(n%).no(0) = Deline(no(0) + 1).data(0).record.data1.index.i(0)
 Else
 Call search_for_eline(con_eline(n%).data(0).data0, 1, n_(1), 1)
 Call search_for_eline(con_eline(n%).data(0).data0, 2, n_(2), 1)
 Call search_for_eline(con_eline(n%).data(0).data0, 3, n_(3), 1)
 'Call search_for_eline(con_eline(n%).data(0), 4, n_(4), 1)
 'temp_record.record_data.data0.condition_data.condition_no = 0
 temp_record.record_data.data0.theorem_no = 1
 'Call add_conditions_to_record(cond_type, no(0), dn(0), dn(1), temp_record.record_data.data0.condition_data)
 Call set_level(temp_record.record_data.data0.condition_data)
 If last_conditions.last_cond(1).eline_no Mod 10 = 0 Then
 ReDim Preserve Deline(last_conditions.last_cond(1).eline_no + 10) As eline_type
 End If
 last_conditions.last_cond(1).eline_no = last_conditions.last_cond(1).eline_no + 1
 conclusion_data(n%).no(0) = last_conditions.last_cond(1).eline_no
  Deline(last_conditions.last_cond(1).eline_no).data(0) = eline_data_0
 Deline(last_conditions.last_cond(1).eline_no).data(0).data0 = con_eline(n%).data(0).data0
 Deline(last_conditions.last_cond(1).eline_no).data(0).record = temp_record.record_data
 For i% = 0 To 3
 For j% = last_conditions.last_cond(1).eline_no To n_(i%) + 2 Step -1
  Deline(j%).data(0).record.data1.index.i(i%) = Deline(j% - 1).data(0).record.data1.index.i(i%)
 Next j%
  Deline(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).eline_no
 Next i%
 End If
End If
End If
End If
Case midpoint_
If con_type = 0 And is_set_reduce Then
 Call set_conclusion_point(n%, con_mid_point(n%).data(0).poi(0))
 Call set_conclusion_point(n%, con_mid_point(n%).data(0).poi(1))
 Call set_conclusion_point(n%, con_mid_point(n%).data(0).poi(2))
 'Call set_condition_reduce(point_, con_mid_point(n%).data(0).poi(0), 0, n% + 1)
 'Call set_condition_reduce(point_, con_mid_point(n%).data(0).poi(1), 0, n% + 1)
 'Call set_condition_reduce(point_, con_mid_point(n%).data(0).poi(2), 0, n% + 1)
ElseIf con_type = line_value_ Or con_type = midpoint_ Then
c_data.condition_no = 0
If is_mid_point(con_mid_point(n%).data(0).poi(0), con_mid_point(n%).data(0).poi(1), _
    con_mid_point(n%).data(0).poi(2), con_mid_point(n%).data(0).n(0), _
     con_mid_point(n%).data(0).n(1), con_mid_point(n%).data(0).n(2), _
      con_mid_point(n%).data(0).line_no, no(0), -3000, 0, 0, 0, 0, 0, 0, _
       Dmid_point_data0, "", cond_type, no(0), no1%, c_data) Then
find_conclusion = 1
If cond_type = midpoint_ Then
conclusion_data(n%).no(0) = no(0)
Dmid_point(no(0)).data(0).record.data1.is_proved = 3
Else
If search_for_mid_point(con_mid_point(n%).data(0), 0, n_(0), 1) Then '5.7
 conclusion_data(n%).no(0) = Dmid_point(n_(0) + 1).data(0).record.data1.index.i(0)
Else
 Call search_for_mid_point(con_mid_point(n%).data(0), 1, n_(1), 1) '5.7
 Call search_for_mid_point(con_mid_point(n%).data(0), 2, n_(2), 1)
temp_record.record_data.data0.condition_data.condition_no = 0
temp_record.record_data.data0.theorem_no = 1
Call add_conditions_to_record(cond_type, no(0), no1%, 0, temp_record.record_data.data0.condition_data)
Call set_level(temp_record.record_data.data0.condition_data)
If last_conditions.last_cond(1).mid_point_no Mod 10 = 0 Then
ReDim Preserve Dmid_point(last_conditions.last_cond(1).mid_point_no + 10) As mid_point_type
End If
last_conditions.last_cond(1).mid_point_no = last_conditions.last_cond(1).mid_point_no + 1
conclusion_data(n%).no(0) = last_conditions.last_cond(1).mid_point_no
 Dmid_point(conclusion_data(n%).no(0)).data(0) = mid_point_data_0
Dmid_point(conclusion_data(n%).no(0)).data(0).data0 = con_mid_point(n%).data(0)
Call set_level(temp_record.record_data.data0.condition_data)
Dmid_point(conclusion_data(n%).no(0)).data(0).record.data0.condition_data = c_data
'Dmid_pointconclusion_no(n%,0)).data(1).record.other_no = conclusion_no(n%,0)
Dmid_point(conclusion_data(n%).no(0)).data(0).record.data1.is_proved = 3
For i% = 0 To 2
For j% = last_conditions.last_cond(1).eline_no To n_(i%) + 2 Step -1
Deline(j%).data(0).record.data1.index.i(i%) = Deline(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Deline(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).eline_no
Next i%
End If
End If
End If
ElseIf con_type = relation_ Then
 If Drelation(con_no%).data(0).data0.poi(1) = Drelation(con_no%).data(0).data0.poi(2) Then
   If con_mid_point(n%).data(0).poi(0) = Drelation(con_no%).data(0).data0.poi(0) And _
        con_mid_point(n%).data(0).poi(1) = Drelation(con_no%).data(0).data0.poi(2) And _
          con_mid_point(n%).data(0).poi(2) = Drelation(con_no%).data(0).data0.poi(3) Then
           error_of_wenti = 2
            find_conclusion = 1
             Exit Function
   End If
 End If
End If
Case length_of_polygon_
If con_type = 0 And is_set_reduce Then
   ele.ty = polygon_
   ele.no = con_length_of_polygon(n%).polygon_no
   Call set_conclusion_point_for_area_element(n%, ele)
   'Call set_area_of_element_reduce(ele, n% + 1)
ElseIf con_type = length_of_polygon_ Then
   If length_of_polygon(con_no%).record_.conclusion_no = n% + 1 Then
   If length_of_polygon(con_no%).data(0).last_segment = 0 And _
        InStr(1, length_of_polygon(con_no%).data(0).value, "x", 0) = 0 Then
         conclusion_data(n%).no(0) = con_no%
          find_conclusion = 1
            Exit Function
   End If
   End If
ElseIf con_type = sides_length_of_triangle_ Then
  If con_length_of_polygon(n%).polygon_ty = triangle_ And _
       con_length_of_polygon(n%).polygon_no = Sides_length_of_triangle(con_no%).data(0).triangle Then
    If InStr(1, Sides_length_of_triangle(con_no%).data(0).value, "x", 0) Then
       conclusion_data(n%).ty = sides_length_of_triangle_
        conclusion_data(n%).no(0) = con_no%
        find_conclusion = 1
          Exit Function
    End If
  End If
End If
Case line_value_
If con_type = line_value_ Then
 If is_contain_x(line_value(con_no%).data(0).data0.value_, "x", 1) = False Then '不含未知数
  If is_same_two_point(con_line_value(n%).data(0).data0.poi(0), con_line_value(n%).data(0).data0.poi(1), _
      line_value(con_no%).data(0).data0.poi(0), line_value(con_no%).data(0).data0.poi(1)) Then
   If line_value(con_no%).data(0).data0.value = line_value(con_no%).data(0).data0.value_ Then '未设未知数
        conclusion_data(n%).no(0) = con_no%
   Else
        conclusion_data(n%).no(0) = -con_no% '设有未知数
   End If
   line_value(con_no%).data(0).record.data1.is_proved = 3
    find_conclusion = 1
  End If
 Else '
  If is_same_two_point(con_line_value(n%).data(0).data0.poi(0), con_line_value(n%).data(0).data0.poi(1), _
      line_value(con_no%).data(0).data0.poi(0), line_value(con_no%).data(0).data0.poi(1)) Then
        conclusion_data(n%).no(1) = con_no% '
  End If
 End If
ElseIf con_type = 0 Then
 If is_set_reduce Then
 Call set_conclusion_point(n%, con_line_value(n%).data(0).data0.poi(0))
 Call set_conclusion_point(n%, con_line_value(n%).data(0).data0.poi(1))
 'Call set_condition_reduce(point_, con_line_value(n%).data(0).data0.poi(0), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line_value(n%).data(0).data0.poi(1), 0, n% + 1)
 End If
 If is_line_value(con_line_value(n%).data(0).data0.poi(0), _
        con_line_value(n%).data(0).data0.poi(1), con_line_value(n%).data(0).data0.n(0), _
         con_line_value(n%).data(0).data0.n(1), con_line_value(n%).data(0).data0.line_no, _
          "", no(0), -1000, 0, 0, 0, line_value_data0) = 1 Then
  If is_contain_x(line_value(no(0)).data(0).data0.value_, "x", 1) = False Then
   If line_value(no(0)).data(0).data0.value = line_value(no(0)).data(0).data0.value_ Then
    conclusion_data(n%).no(0) = no(0)
   Else
    conclusion_data(n%).no(0) = -no(0)
   End If
   line_value(no(0)).data(0).record.data1.is_proved = 3
    find_conclusion = 1
  End If
 End If
End If
Case area_relation_
If con_type = 0 And is_set_reduce Then
 Call set_conclusion_point_for_area_element(n%, con_area_relation(n%).data(0).area_element(0))
 Call set_conclusion_point_for_area_element(n%, con_area_relation(n%).data(0).area_element(1))
 'Call set_area_of_element_reduce(con_area_relation(n%).data(0).area_element(0), n% + 1)
 'Call set_area_of_element_reduce(con_area_relation(n%).data(0).area_element(1), n% + 1)
End If
If is_area_relation(con_area_relation(n%).data(0).area_element(0), _
     con_area_relation(n%).data(0).area_element(1), "", no(0), -1000, 0, 0, _
       condition_type0, condition_type0, condition_type0, "", con_ty(0), tn(0), tn(1)) Then
 If con_ty(0) = area_relation_ Then
        conclusion_data(n%).no(0) = no(0)
 Else 'If is_area_of_triangle(con_area_relation(n%).data(0).area_element(0).no, no(0)) And _
     is_area_of_triangle(con_area_relation(n%).data(0).area_element(1).no, no(1)) Then
If last_conditions.last_cond(1).area_relation_no Mod 10 = 0 Then
ReDim PreserveDarea_relation(last_conditions.last_cond(1).area_relation_no + 10) As area_relation_type
End If
last_conditions.last_cond(1).area_relation_no = last_conditions.last_cond(1).area_relation_no + 1
Darea_relation(last_conditions.last_cond(1).area_relation_no) = con_area_relation(n%)
Darea_relation(last_conditions.last_cond(1).area_relation_no).data(0).value = _
  divide_string(area_of_element(tn(0)).data(0).value_, area_of_element(tn(1)).data(0).value_, True, False)
Call add_conditions_to_record(area_of_element_, tn(0), tn(1), 0, _
Darea_relation(last_conditions.last_cond(1).area_relation_no).data(0).record.data0.condition_data)
End If
End If
Case line3_value_
If con_type = 0 And is_set_reduce Then
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(0))
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(1))
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(2))
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(3))
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(4))
 Call set_conclusion_point(n, con_line3_value(n%).data(0).poi(5))
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(0), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(1), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(2), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(3), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(4), 0, n% + 1)
 'Call set_condition_reduce(point_, con_line3_value(n%).data(0).poi(5), 0, n% + 1)
ElseIf con_type = line3_value_ Then
   If con_line3_value(n%).data(0).poi(0) = line3_value(con_no%).data(0).data0.poi(0) And _
       con_line3_value(n%).data(0).poi(1) = line3_value(con_no%).data(0).data0.poi(1) And _
        con_line3_value(n%).data(0).poi(2) = line3_value(con_no%).data(0).data0.poi(2) And _
         con_line3_value(n%).data(0).poi(3) = line3_value(con_no%).data(0).data0.poi(3) And _
          con_line3_value(n%).data(0).poi(4) = line3_value(con_no%).data(0).data0.poi(4) And _
           con_line3_value(n%).data(0).poi(5) = line3_value(con_no%).data(0).data0.poi(5) And _
     con_line3_value(n%).data(0).para(0) = line3_value(con_no%).data(0).data0.para(0) And _
      con_line3_value(n%).data(0).para(1) = line3_value(con_no%).data(0).data0.para(1) And _
       con_line3_value(n%).data(0).para(2) = line3_value(con_no%).data(0).data0.para(2) Then
      conclusion_data(n%).no(0) = con_no%
   line3_value(con_no%).data(0).record.data1.is_proved = 3
    find_conclusion = 1
 Else
n_(0) = -5000
 If is_three_line_value( _
    con_line3_value(n%).data(0).poi(0), con_line3_value(n%).data(0).poi(1), _
     con_line3_value(n%).data(0).poi(2), con_line3_value(n%).data(0).poi(3), _
      con_line3_value(n%).data(0).poi(4), con_line3_value(n%).data(0).poi(5), _
    con_line3_value(n%).data(0).n(0), con_line3_value(n%).data(0).n(1), _
     con_line3_value(n%).data(0).n(2), con_line3_value(n%).data(0).n(3), _
      con_line3_value(n%).data(0).n(4), con_line3_value(n%).data(0).n(5), _
    con_line3_value(n%).data(0).line_no(0), con_line3_value(n%).data(0).line_no(1), _
     con_line3_value(n%).data(0).line_no(2), con_line3_value(n%).data(0).para(0), _
      con_line3_value(n%).data(0).para(1), con_line3_value(n%).data(0).para(2), _
       con_line3_value(n%).data(0).value, no(0), n_(0), n_(1), n_(2), n_(3), _
        n_(4), n_(5), line3_value_data0, 0, c_data, 0) = 1 Then
If no(0) > 0 And is_contain_x(line3_value(no(0)).data(0).data0.value_, "x", 1) = False Then
   conclusion_data(n%).no(0) = no(0)
   line3_value(no(0)).data(0).record.data1.is_proved = 3
    find_conclusion = 1
 Else
  If is_contain_x(line3_value_data0.value_, "x", 1) = False Then
    For i% = 0 To 5
     Call search_for_line3_value(line3_value_data0, i%, n_(i%), 1)
    Next i%
 If last_conditions.last_cond(1).line3_value_no Mod 10 = 0 Then
 ReDim Preserve line3_value(last_conditions.last_cond(1).line3_value_no + 10) As line3_value_type
 End If
 last_conditions.last_cond(1).line3_value_no = last_conditions.last_cond(1).line3_value_no + 1
 line3_value(last_conditions.last_cond(1).line3_value_no).data(0).data0 = line3_value_data0
 line3_value(last_conditions.last_cond(1).line3_value_no).data(0).record.data0.condition_data = _
  c_data
  For i% = 0 To 5
   For j% = last_conditions.last_cond(1).line3_value_no To n_(i%) + 2 Step -1
    line3_value(j%).data(0).record.data1.index.i(i%) = line3_value(j% - 1).data(0).record.data1.index.i(i%)
   Next j%
   line3_value(n_(i%) + 1).data(0).record.data1.index.i(i%) = _
         last_conditions.last_cond(1).line3_value_no
  Next i%
    conclusion_data(n%).no(0) = last_conditions.last_cond(1).line3_value_no
   line3_value(last_conditions.last_cond(1).line3_value_no).data(0).record.data1.is_proved = 3
    find_conclusion = 1
 End If
 End If
 End If
 End If
 End If
Case dpoint_pair_
If con_type = 0 And is_set_reduce Then
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(0))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(1))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(2))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(3))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(4))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(5))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(6))
 Call set_conclusion_point(n%, con_dpoint_pair(n%).data(0).poi(7))
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(0), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(1), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(2), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(3), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(4), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(5), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(6), 0, n% + 1)
 'Call set_condition_reduce(point_, con_dpoint_pair(n%).data(0).poi(7), 0, n% + 1)
ElseIf con_type = line_value_ Or con_type = eline_ Or con_type = midpoint_ Or _
 con_type = relation_ Or con_type = dpoint_pair_ Or con_type = 0 Then
record_0.data0.condition_data.condition_no = 0 'record0
Dim dp As point_pair_data0_type
If con_type = dpoint_pair_ Then
 If compare_two_point_pair(con_dpoint_pair(n%).data(0), _
       Ddpoint_pair(con_no%).data(0).data0, 0) = 0 Then
  find_conclusion = 1
   conclusion_data(n%).no(0) = con_no%
    Ddpoint_pair(con_no%).data(0).record.data1.is_proved = 3
 End If
Else
temp_record.record_data.data0.condition_data.condition_no = 0
If is_point_pair(con_dpoint_pair(n%).data(0).poi(0), _
    con_dpoint_pair(n%).data(0).poi(1), _
    con_dpoint_pair(n%).data(0).poi(2), _
    con_dpoint_pair(n%).data(0).poi(3), _
    con_dpoint_pair(n%).data(0).poi(4), _
    con_dpoint_pair(n%).data(0).poi(5), _
    con_dpoint_pair(n%).data(0).poi(6), _
    con_dpoint_pair(n%).data(0).poi(7), _
    con_dpoint_pair(n%).data(0).n(0), _
    con_dpoint_pair(n%).data(0).n(1), _
    con_dpoint_pair(n%).data(0).n(2), _
    con_dpoint_pair(n%).data(0).n(3), _
    con_dpoint_pair(n%).data(0).n(4), _
    con_dpoint_pair(n%).data(0).n(5), _
    con_dpoint_pair(n%).data(0).n(6), _
    con_dpoint_pair(n%).data(0).n(7), _
    con_dpoint_pair(n%).data(0).line_no(0), _
    con_dpoint_pair(n%).data(0).line_no(1), _
    con_dpoint_pair(n%).data(0).line_no(2), _
    con_dpoint_pair(n%).data(0).line_no(3), _
     no(0), -3000, 0, 0, 0, 0, 0, dp, cond_type, tn(0), tn(1), _
        con_ty(0), con_ty(1), tn(2), tn(3), tn(4), tn(5), "", "", _
          temp_record.record_data) Then
  find_conclusion = 1
  If no(0) > 0 Then
  conclusion_data(n%).no(0) = no(0)
  Ddpoint_pair(no(0)).data(0).record.data1.is_proved = 3
  Else
   If search_for_point_pair(con_dpoint_pair(n%).data(0), 0, n_(0), 1) Then
    conclusion_data(n%).no(0) = Ddpoint_pair(n_(0) + 1).data(0).record.data1.index.i(0)
     Ddpoint_pair(n_(0)).data(0).record.data1.is_proved = 3
   Else
   temp_record.record_data.data0.theorem_no = 1
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 1, n_(1), 1)
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 2, n_(2), 1)
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 3, n_(3), 1)
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 4, n_(4), 1)
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 5, n_(5), 1)
   Call search_for_point_pair(con_dpoint_pair(n%).data(0), 6, n_(6), 1)
    temp_record.record_data.data0.condition_data.condition_no = 0
    temp_record.record_data.data0.theorem_no = 1
    Call set_record_for_point_pair(temp_record.record_data, cond_type, con_ty(0), con_ty(1), _
   no(0), tn(0), tn(1), tn(2), tn(3), tn(4), tn(5))
 If last_conditions.last_cond(1).dpoint_pair_no Mod 10 = 0 Then
  ReDim Preserve Ddpoint_pair(last_conditions.last_cond(1).dpoint_pair_no + 10) As Dpoint_pair_type
 End If
    last_conditions.last_cond(1).dpoint_pair_no = last_conditions.last_cond(1).dpoint_pair_no + 1
      conclusion_data(n%).no(0) = last_conditions.last_cond(1).dpoint_pair_no
      no(0) = last_conditions.last_cond(1).dpoint_pair_no
  Ddpoint_pair(last_conditions.last_cond(1).dpoint_pair_no).data(0) = dpoint_pair_data_0
  For i% = 0 To 6
   Ddpoint_pair(no(0)).data(0).data0 = _
    con_dpoint_pair(n%).data(0)
  Next i%
  Call set_level(temp_record.record_data.data0.condition_data)
  Ddpoint_pair(no(0)).data(0).record = temp_record.record_data
 ' Ddpoint_pair(no(0)).record.other_no = no(0)
  Ddpoint_pair(no(0)).data(0).record.data1.is_proved = 3
  For i% = 0 To 7
  For j% = last_conditions.last_cond(1).dpoint_pair_no To n_(i%) + 2 Step -1
  Ddpoint_pair(j%).data(0).record.data1.index.i(i%) = Ddpoint_pair(j% - 1).data(0).record.data1.index.i(i%)
  Next j%
  Ddpoint_pair(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).dpoint_pair_no
  Next i%
  End If
  End If
  End If
  End If
End If
Case verti_mid_line_
If con_type = verti_mid_line_ Or con_type = 0 Then
   If con_type = 0 And is_set_reduce Then
    Call set_conclusion_point(n%, con_verti_mid_line(n%).data(0).data0.poi(0))
    Call set_conclusion_point(n%, con_verti_mid_line(n%).data(0).data0.poi(1))
    Call set_conclusion_point(n%, con_verti_mid_line(n%).data(0).data0.poi(2))
    Call set_conclusion_point(n%, con_verti_mid_line(n%).data(0).data0.poi(3))
    'Call set_condition_reduce(point_, con_verti_mid_line(n%).data(0).data0.poi(0), 0, n% + 1)
    'Call set_condition_reduce(point_, con_verti_mid_line(n%).data(0).data0.poi(1), 0, n% + 1)
    'Call set_condition_reduce(point_, con_verti_mid_line(n%).data(0).data0.poi(2), 0, n% + 1)
    'Call set_condition_reduce(point_, con_verti_mid_line(n%).data(0).data0.poi(3), 0, n% + 1)
   End If
If is_verti_mid_line(con_verti_mid_line(n%).data(0).data0.poi(0), _
       con_verti_mid_line(n%).data(0).data0.poi(1), _
         con_verti_mid_line(n%).data(0).data0.poi(2), _
      con_verti_mid_line(n%).data(0).data0.line_no(0), no(0), 0, 0, _
       verti_mid_line_data0) = False Then
       conclusion_data(n%).no(0) = no(0)
 find_conclusion = 1
 End If
 End If
Case rhombus_
If con_type = 0 And is_set_reduce Then
ele.ty = polygon_
ele.no = con_rhombus(n%).data(0).polygon4_no
Call set_conclusion_point_for_area_element(n%, ele)
'Call set_area_of_element_reduce(ele, n% + 1)
ElseIf con_type = rhombus_ Or con_type = epolygon_ Then
If is_rhombus(Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(0), _
              Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(1), _
              Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(2), _
              Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(3), no(0), _
              poly4_no%, -1000, cond_type) = True Then
    'If no(0) > 0 Then
    If cond_type = rhombus_ Then
     conclusion_data(n%).no(0) = no(0)
      find_conclusion = 1
       rhombus(no(0)).data(0).record.data1.is_proved = 3
    ElseIf cond_type = epolygon_ Then
         temp_record.record_data.data0.condition_data.condition(1).ty = epolygon_
          temp_record.record_data.data0.condition_data.condition(1).no = no(0)
      tp(0) = epolygon(no(0)).data(0).p.v(0)
      tp(1) = epolygon(no(0)).data(0).p.v(1)
      tp(2) = epolygon(no(0)).data(0).p.v(2)
      tp(3) = epolygon(no(0)).data(0).p.v(3)
 If last_conditions.last_cond(1).rhombus_no = last_conditions.last_cond(2).rhombus_no Then
    ReDim Preserve rhombus(last_conditions.last_cond(2).rhombus_no + 10) As rhombus_type
    last_conditions.last_cond(2).rhombus_no = last_conditions.last_cond(2).rhombus_no + 10
 End If
         last_conditions.last_cond(1).rhombus_no = last_conditions.last_cond(1).rhombus_no + 1
   conclusion_data(n%).no(0) = last_conditions.last_cond(1).rhombus_no
     rhombus(conclusion_data(n%).no(0)).data(0) = dpolygon4_data_0
      find_conclusion = 1
     Dpolygon4(rhombus(conclusion_data(n%).no(0)).data(0).polygon4_no).data(0).poi(0) = tp(0)
      Dpolygon4(rhombus(conclusion_data(n%).no(0)).data(0).polygon4_no).data(0).poi(1) = tp(1)
       Dpolygon4(rhombus(conclusion_data(n%).no(0)).data(0).polygon4_no).data(0).poi(2) = tp(2)
        Dpolygon4(rhombus(conclusion_data(n%).no(0)).data(0).polygon4_no).data(0).poi(3) = tp(3)
         Call set_level(temp_record.record_data.data0.condition_data)
          rhombus(conclusion_data(n%).no(0)).data(1).record = temp_record.record_data
         ' rhombusconclusion_no(n%,0)).data(1).record.other_no = conclusion_no(n%,0)
     End If
  End If
End If
'End If
Case long_squre_
If con_type = 0 And is_set_reduce Then
ele.ty = polygon_
ele.no = con_long_squre(no(0)).data(0).polygon4_no
Call set_conclusion_point_for_area_element(n%, ele)
'Call set_area_of_element_reduce(ele, n% + 1)
ElseIf con_type = long_squre_ Or con_type = epolygon_ Then
 If is_long_squre(Dpolygon4(con_long_squre(no(0)).data(0).polygon4_no).data(0).poi(0), _
                  Dpolygon4(con_long_squre(no(0)).data(0).polygon4_no).data(0).poi(1), _
                  Dpolygon4(con_long_squre(no(0)).data(0).polygon4_no).data(0).poi(2), _
                  Dpolygon4(con_long_squre(no(0)).data(0).polygon4_no).data(0).poi(3), _
                  no(0), 0, -1000, cond_type) = True Then
    If cond_type = epolygon_ Then
 If last_conditions.last_cond(1).long_squre_no = last_conditions.last_cond(1).long_squre_no Then
 ReDim Preserve Dlong_squre(last_conditions.last_cond(2).long_squre_no + 10) As long_squre_type
 last_conditions.last_cond(2).long_squre_no = last_conditions.last_cond(2).long_squre_no + 10
 End If
   last_conditions.last_cond(1).long_squre_no = last_conditions.last_cond(1).long_squre_no + 1
    i% = last_conditions.last_cond(1).long_squre_no
 Dlong_squre(i%).data(0) = dpolygon4_data_0
    Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(0) = epolygon(no(0)).data(0).p.v(0)
    Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(1) = epolygon(no(0)).data(0).p.v(1)
    Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(2) = epolygon(no(0)).data(0).p.v(2)
    Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(3) = epolygon(no(0)).data(0).p.v(3)
    conclusion_data(n%).no(0) = i%
     Dlong_squre(i%).data(0).record.data1.is_proved = 3
     'Dlong_squre(i%).record.other_no = i%
    Else
    conclusion_data(n%).no(0) = no(0)
     Dlong_squre(no(0)).data(0).record.data1.is_proved = 3
     'Dlong_squre(no(0)).record.other_no = no(0)
    End If
 End If
 End If
Case tangent_line_
  If con_type = tangent_line_ Then
    If con_tangent_line(n%).data(0).line_no = tangent_line(con_no%).data(0).line_no Then
       If (con_tangent_line(n%).data(0).ele(0).no = tangent_line(con_no%).data(0).ele(0).no And _
            con_tangent_line(n%).data(0).ele(0).ty = tangent_line(con_no%).data(0).ele(0).ty) Or _
         (con_tangent_line(n%).data(0).ele(1).no = tangent_line(con_no%).data(0).ele(1).no And _
            con_tangent_line(n%).data(0).ele(1).ty = tangent_line(con_no%).data(0).ele(1).ty) Then
                conclusion_data(n%).no(0) = con_no%
                 tangent_line(con_no%).data(0).record.data1.is_proved = 3
           find_conclusion = 1
       End If
    End If
  ElseIf con_type = 0 Then
   If is_set_reduce Then
      Call set_conclusion_point(n%, con_tangent_line(n%).data(0).poi(0))
      Call set_conclusion_point(n%, con_tangent_line(n%).data(0).poi(1))
      Call set_conclusion_point_for_circle(n%, con_tangent_line(n%).data(0).ele(0).no)
      Call set_conclusion_point_for_circle(n%, con_tangent_line(n%).data(0).ele(1).no)
      Call set_conclusion_point_for_line(n%, con_tangent_line(n%).data(0).line_no)
      'Call set_condition_reduce(point_, con_tangent_line(n%).data(0).poi(0), 0, n% + 1)
      'Call set_condition_reduce(point_, con_tangent_line(n%).data(0).poi(1), 0, n% + 1)
      'Call set_condition_reduce(circle_, con_tangent_line(n%).data(0).circ(0), 0, n% + 1)
      'Call set_condition_reduce(circle_, con_tangent_line(n%).data(0).circ(1), 0, n% + 1)
      'Call set_condition_reduce(line_, con_tangent_line(n%).data(0).line_no, 0, n% + 1)
   End If
   no(0) = 0
   If is_tangent_line(con_tangent_line(n%).data(0).line_no, con_tangent_line(n%).data(0).poi(0), _
       con_tangent_line(n%).data(0).ele(0), con_tangent_line(n%).data(0).poi(1), _
        con_tangent_line(n%).data(0).ele(1), t_l, no(0), 0, 0, record_0) Then
                conclusion_data(n%).no(0) = no(0)
                 tangent_line(no(0)).data(0).record.data1.is_proved = 3
                  find_conclusion = 1
   End If
  End If
Case two_line_value_
 If con_type = 0 And is_set_reduce Then
  Call set_conclusion_point(n%, con_two_line_value(n%).data(0).poi(0))
  Call set_conclusion_point(n%, con_two_line_value(n%).data(0).poi(1))
  Call set_conclusion_point(n%, con_two_line_value(n%).data(0).poi(2))
  Call set_conclusion_point(n%, con_two_line_value(n%).data(0).poi(3))
  'Call set_condition_reduce(point_, con_two_line_value(n%).data(0).poi(0), 0, n% + 1)
  'Call set_condition_reduce(point_, con_two_line_value(n%).data(0).poi(1), 0, n% + 1)
  'Call set_condition_reduce(point_, con_two_line_value(n%).data(0).poi(2), 0, n% + 1)
  'Call set_condition_reduce(point_, con_two_line_value(n%).data(0).poi(3), 0, n% + 1)
 ElseIf con_type = two_line_value_ Then
  If con_two_line_value(n%).data(0).value <> "" Then
  If is_two_line_value(con_two_line_value(n%).data(0).poi(0), con_two_line_value(n%).data(0).poi(1), _
      con_two_line_value(n%).data(0).poi(1), con_two_line_value(n%).data(0).poi(3), _
       con_two_line_value(n%).data(0).n(0), con_two_line_value(n%).data(0).n(1), _
         con_two_line_value(n%).data(0).n(1), con_two_line_value(n%).data(0).n(3), _
          con_two_line_value(n%).data(0).line_no(0), con_two_line_value(n%).data(0).line_no(1), _
           con_two_line_value(n%).data(0).para(0), con_two_line_value(n%).data(0).para(1), _
            con_two_line_value(n%).data(0).value, no(0), -1000, 0, 0, 0, l2v_data0, 0, c_data) = 1 Then
             conclusion_data(n%).no(0) = no(0)
              find_conclusion = 1
  End If
  Else
    If con_two_line_value(n%).data(0).poi(0) = two_line_value(con_no%).data(0).data0.poi(0) And _
       con_two_line_value(n%).data(0).poi(1) = two_line_value(con_no%).data(0).data0.poi(1) And _
       con_two_line_value(n%).data(0).poi(2) = two_line_value(con_no%).data(0).data0.poi(2) And _
       con_two_line_value(n%).data(0).poi(3) = two_line_value(con_no%).data(0).data0.poi(3) And _
       con_two_line_value(n%).data(0).para(0) = two_line_value(con_no%).data(0).data0.para(0) And _
       con_two_line_value(n%).data(0).para(1) = two_line_value(con_no%).data(0).data0.para(1) Then
             conclusion_data(n%).no(0) = con_no%
              find_conclusion = 1
    End If
  End If
 ElseIf con_type = eline_ And con_two_line_value(n%).data(0).value = "" Then
    If con_two_line_value(n%).data(0).poi(0) = Deline(con_no%).data(0).data0.poi(0) And _
       con_two_line_value(n%).data(0).poi(1) = Deline(con_no%).data(0).data0.poi(1) And _
       con_two_line_value(n%).data(0).poi(2) = Deline(con_no%).data(0).data0.poi(2) And _
       con_two_line_value(n%).data(0).poi(3) = Deline(con_no%).data(0).data0.poi(3) And _
       con_two_line_value(n%).data(0).para(0) = "1" And _
       con_two_line_value(n%).data(0).para(1) = "0" Then
             conclusion_data(n%).no(0) = con_no%
             conclusion_data(n%).ty = eline_
              find_conclusion = 1
    End If
 ElseIf con_type = relation_ And con_two_line_value(n%).data(0).value = "" Then
    If is_relation(con_two_line_value(n%).data(0).poi(0), con_two_line_value(n%).data(0).poi(1), _
                    con_two_line_value(n%).data(0).poi(2), con_two_line_value(n%).data(0).poi(3), _
                     con_two_line_value(n%).data(0).n(0), con_two_line_value(n%).data(0).n(1), _
                   con_two_line_value(n%).data(0).n(2), con_two_line_value(n%).data(0).n(3), _
                    con_two_line_value(n%).data(0).line_no(0), con_two_line_value(n%).data(0).line_no(1), _
                     divide_string(con_two_line_value(n%).data(0).para(1), time_string("-1", _
                   con_two_line_value(n%).data(0).para(0), False, False), True, False), _
                    no(0), -1000, 0, 0, 0, dr_data, tn(1), tn(2), con_ty(0), record_0.data0.condition_data, 0) Then
        If no(0) > 0 Then
            conclusion_data(n%).no(0) = no(0)
             conclusion_data(n%).ty = relation_
              find_conclusion = 1
        Else
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_conditions_to_record(line_value_, no(1), no(2), 0, temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         If last_conditions.last_cond(1).relation_no Mod 10 = 0 Then
            ReDim Preserve Drelation(last_conditions.last_cond(1).relation_no + 10) As relation_type
         End If
         last_conditions.last_cond(1).relation_no = last_conditions.last_cond(1).relation_no + 1
            conclusion_data(n%).no(0) = last_conditions.last_cond(1).relation_no
             conclusion_data(n%).ty = relation_
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.poi(0) = line_value(no(1)).data(0).data0.poi(0)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.poi(1) = line_value(no(1)).data(0).data0.poi(1)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.n(0) = line_value(no(1)).data(0).data0.n(0)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.n(1) = line_value(no(1)).data(0).data0.n(1)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.line_no(0) = line_value(no(1)).data(0).data0.line_no
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.poi(2) = line_value(no(2)).data(0).data0.poi(2)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.poi(3) = line_value(no(2)).data(0).data0.poi(3)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.n(2) = line_value(no(2)).data(0).data0.n(2)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.n(3) = line_value(no(2)).data(0).data0.n(3)
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.line_no(1) = line_value(no(2)).data(0).data0.line_no
            Drelation(last_conditions.last_cond(1).relation_no).data(0).data0.value = _
              divide_string(line_value(no(1)).data(0).data0.value, line_value(no(2)).data(0).data0.value, _
                  True, False)
              find_conclusion = 1
        End If
    End If
 Else
  If is_line_value(con_two_line_value(n%).data(0).poi(0), con_two_line_value(n%).data(0).poi(1), _
       con_two_line_value(n%).data(0).n(0), con_two_line_value(n%).data(0).n(1), _
        con_two_line_value(n%).data(0).line_no(0), "", no(0), -1000, 0, 0, 0, lv_data0) = 1 And _
      is_line_value(con_two_line_value(n%).data(0).poi(1), con_two_line_value(n%).data(0).poi(3), _
       con_two_line_value(n%).data(0).n(2), con_two_line_value(n%).data(0).n(3), _
        con_two_line_value(n%).data(0).line_no(1), "", no(1), -1000, 0, 0, 0, lv_data0) = 1 Then
         If is_contain_x(line_value(no(0)).data(0).data0.value_, "x", 1) = False And _
              is_contain_x(line_value(no(1)).data(0).data0.value_, "x", 1) = False Then
     temp_record.record_data.data0.condition_data.condition_no = 0
     Call add_conditions_to_record(line_value_, no(0), no(1), 0, temp_record.record_data.data0.condition_data)
 If last_conditions.last_cond(1).two_line_value_no Mod 10 = 0 Then
         ReDim Preserve two_line_value(last_conditions.last_cond(1).two_line_value_no + 10) As two_line_value_type
 End If
   last_conditions.last_cond(1).two_line_value_no = last_conditions.last_cond(1).two_line_value_no + 1
      no1% = last_conditions.last_cond(1).two_line_value_no
          two_line_value(no1%).data(0).data0 = con_two_line_value(n%).data(0)
           two_line_value(no1%).data(0).record = temp_record.record_data
            two_line_value(no1%).data(0).data0.value = _
              add_string(time_string(line_value(no(0)).data(0).data0.value_, two_line_value(no1%).data(0).data0.para(0), _
                False, False), time_string(line_value(no(1)).data(0).data0.value_, two_line_value(no1%).data(0).data0.para(0), _
                False, False), True, False)
            conclusion_data(n%).no(0) = no1%
          find_conclusion = 1
         End If
  End If
 End If
Case tixing_
 If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(0))
   Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(1))
   Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(2))
   Call set_conclusion_point(n%, con_Dtixing(n%).data(0).poi(3))
   'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(0), 0, n% + 1)
   'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(1), 0, n% + 1)
   'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(2), 0, n% + 1)
   'Call set_condition_reduce(point_, con_Dtixing(n%).data(0).poi(3), 0, n% + 1)
 End If
 no(0) = 0
 If is_tixing(con_Dtixing(n%).data(0).poi(0), con_Dtixing(n%).data(0).poi(1), con_Dtixing(n%).data(0).poi(2), _
     con_Dtixing(n%).data(0).poi(3), no(0), 0, 0, 0, 0, 0, 0, 0, 0, True) Then
      If Dpolygon4(con_Dtixing(n%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
        If Dpolygon4(Dtixing(no(0)).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
         conclusion_data(n%).no(0) = no(0)
           Dtixing(no(0)).data(0).record.data1.is_proved = 3
            find_conclusion = 1
        ElseIf is_equal_dline(con_Dtixing(n%).data(0).poi(0), con_Dtixing(n%).data(0).poi(3), _
          con_Dtixing(n%).data(0).poi(1), con_Dtixing(n%).data(0).poi(2), 0, 0, 0, 0, 0, 0, _
             0, no(1), -1000, 0, 0, eline_data0, no(2), no(3), con_ty(0), "", c_data) Then
             Call add_conditions_to_record(con_ty(0), no(1), no(2), no(3), _
                    Dtixing(no(0)).data(0).record.data0.condition_data)
                     Dpolygon4(Dtixing(no(0)).data(0).poly4_no).data(0).ty = equal_side_tixing_
         conclusion_data(n%).no(0) = no(0)
           Dtixing(no(0)).data(0).record.data1.is_proved = 3
            find_conclusion = 1
        End If
      End If
 End If
Case parallelogram_
If con_type = parallelogram_ Or con_type = long_squre_ Or _
    con_type = rhombus_ Or con_type = epolygon_ Then
If is_parallelogram(con_parallelogram(n%).data(0).poi(0), _
 con_parallelogram(n%).data(0).poi(1), con_parallelogram(n%).data(0).poi(2), _
   con_parallelogram(n%).data(0).poi(3), no(0), -1000, _
    0, cond_type) = True Then
 If cond_type = parallelogram_ Then
  conclusion_data(n%).no(0) = no(0)
   Dparallelogram(no(0)).data(0).record.data1.is_proved = 3
 Else
    temp_record.record_data.data0.condition_data.condition_no = 1
     temp_record.record_data.data0.condition_data.condition(1).no = no(0)
'      Call set_level(temp_record.record_data)
 If cond_type = long_squre_ Then
     temp_record.record_data.data0.condition_data.condition(1).ty = long_squre_
      tp(0) = Dpolygon4(Dlong_squre(no(0)).data(0).polygon4_no).data(0).poi(0)
      tp(1) = Dpolygon4(Dlong_squre(no(0)).data(0).polygon4_no).data(0).poi(1)
      tp(2) = Dpolygon4(Dlong_squre(no(0)).data(0).polygon4_no).data(0).poi(2)
      tp(3) = Dpolygon4(Dlong_squre(no(0)).data(0).polygon4_no).data(0).poi(3)
 ElseIf cond_type = rhombus_ Then
    temp_record.record_data.data0.condition_data.condition(1).ty = rhombus_
      tp(0) = Dpolygon4(rhombus(no(0)).data(0).polygon4_no).data(0).poi(0)
      tp(1) = Dpolygon4(rhombus(no(0)).data(0).polygon4_no).data(0).poi(1)
      tp(2) = Dpolygon4(rhombus(no(0)).data(0).polygon4_no).data(0).poi(2)
      tp(3) = Dpolygon4(rhombus(no(0)).data(0).polygon4_no).data(0).poi(3)
 ElseIf cond_type = epolygon_ Then
    temp_record.record_data.data0.condition_data.condition(1).ty = epolygon_
      tp(0) = epolygon(no(0)).data(0).p.v(0)
      tp(1) = epolygon(no(0)).data(0).p.v(1)
      tp(2) = epolygon(no(0)).data(0).p.v(2)
      tp(3) = epolygon(no(0)).data(0).p.v(3)
 End If
 If last_conditions.last_cond(1).parallelogram_no Mod 10 = 0 Then
    ReDim Preserve Dparallelogram(last_conditions.last_cond(1).parallelogram_no + 10) As parallelogram_type
 End If
  last_conditions.last_cond(1).parallelogram_no = last_conditions.last_cond(1).parallelogram_no + 1
   conclusion_data(n%).no(0) = last_conditions.last_cond(1).parallelogram_no
     Dparallelogram(conclusion_data(n%).no(0)).data(0) = dpolygon4_data_0
     Dpolygon4(Dparallelogram(conclusion_data(n%).no(0)).data(1).polygon4_no).data(0).poi(0) = tp(0)
      Dpolygon4(Dparallelogram(conclusion_data(n%).no(0)).data(1).polygon4_no).data(0).poi(1) = tp(1)
       Dpolygon4(Dparallelogram(conclusion_data(n%).no(0)).data(1).polygon4_no).data(0).poi(2) = tp(2)
        Dpolygon4(Dparallelogram(conclusion_data(n%).no(0)).data(1).polygon4_no).data(0).poi(3) = tp(3)
         Call set_level(temp_record.record_data.data0.condition_data)
         Dparallelogram(conclusion_data(n%).no(0)).data(0).record = temp_record.record_data
          'Dparallelogramconclusion_no(n%,0)).data(1).record.other_no = conclusion_no(n%,0)
 End If
  find_conclusion = 1
 End If
 ElseIf con_type = 0 And is_set_reduce Then
  Call set_conclusion_point(n%, con_parallelogram(n%).data(0).poi(0))
  Call set_conclusion_point(n%, con_parallelogram(n%).data(0).poi(1))
  Call set_conclusion_point(n%, con_parallelogram(n%).data(0).poi(2))
  Call set_conclusion_point(n%, con_parallelogram(n%).data(0).poi(3))
  'Call set_condition_reduce(point_, con_parallelogram(n%).data(0).poi(0), 0, n% + 1)
  'Call set_condition_reduce(point_, con_parallelogram(n%).data(0).poi(1), 0, n% + 1)
  'Call set_condition_reduce(point_, con_parallelogram(n%).data(0).poi(2), 0, n% + 1)
  'Call set_condition_reduce(point_, con_parallelogram(n%).data(0).poi(3), 0, n% + 1)
 End If
End Select
End Function
Public Function next_char(ByVal p%, ByVal ch$, ty As Byte, no%) As String    '
'确定下一个字母,P%>0 为点P% 确定名称,p%=0 and ch$="" 选取下一个小写字母,
 'p%=0 and ch$<>"" 将一个小写字母ch$添加到used_char()数组
Dim i%, j%
If p% > 0 Then
'************************************************
If m_poi(p%).data(0).data0.visible = 0 Then
   next_char = ""
    Exit Function
End If
For j% = 65 To 90
 For i% = 1 To last_conditions.last_cond(1).point_no
 If p% <> i% Then
 If m_poi(i%).data(0).data0.name = Chr(j%) Then
  GoTo next_char_mark
 End If
 End If
 Next i%
next_char = Chr(j%) '未使用的字母
 Exit Function
next_char_mark:
Next j%
Else '
If ch$ = "" Then
 For j% = 97 To 119 '小写字母
  For i% = 1 To last_used_char
   If used_char(i%).name = Chr(j%) Then
     GoTo next_char_next '已使用
   End If
  Next i%
  next_char = Chr(j%) '新字母
   last_used_char = last_used_char + 1
    used_char(last_used_char).name = Chr(j%)
     used_char(last_used_char).cond.ty = ty
      used_char(last_used_char).cond.no = no%
     Exit Function
next_char_next:
 Next j%
ElseIf ch$ >= "a" And ch$ <= "w" Or ch$ >= "A" And ch$ < "Z" Then
 For i% = 1 To last_used_char
  If used_char(i%).name = ch$ Then
   next_char = ch$
    If used_char(i%).cond.no > 0 Then
      ty = used_char(i%).cond.ty
      no% = used_char(i%).cond.no
    Else
     used_char(last_used_char).cond.ty = ty
      used_char(last_used_char).cond.no = no%
    End If
    Exit Function
  End If
 Next i%
last_used_char = last_used_char + 1
  used_char(last_used_char).name = ch$
   used_char(last_used_char).cond.ty = ty
    used_char(last_used_char).cond.no = no%
     next_char = ch$
      Exit Function
End If
End If
End Function

Public Function read_mid_point(ByVal p0%, ByVal p2%, p1%, n%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no
If is_same_two_point(p0%, p2%, Dmid_point(i%).data(0).data0.poi(0), _
   Dmid_point(i%).data(0).data0.poi(2)) = True Then
    n% = i%
     p1% = Dmid_point(i%).data(0).data0.poi(1)
      read_mid_point = True
      Exit Function
End If
Next i%

End Function
Public Function find_conclusion1(con_ty As Byte, con_no%, is_set_reduce As Boolean) As Byte
Dim i%
Dim ty As Byte
If con_ty = 255 Then
 Exit Function
End If
If finish_prove = 3 Or finish_prove = 4 Then
  Exit Function
End If
If prove_or_set_dbase = True Then
  find_conclusion1 = 0
    Exit Function
End If
If last_conclusion > 0 Then
 If wenti_type = 0 Then
  ty = 1
  For i% = 0 To last_conclusion - 1
   If finish_prove = 1 And conclusion_data(i%).no(0) = 0 Then ' conclusion_data(i%)未搜索到
      If con_ty = 0 Or con_ty = conclusion_data(i%).ty Then '
        If find_conclusion(i%, con_ty, con_no%, is_set_reduce) = 0 Then '未搜寻到符合结论的数据
         ty = 0
        ElseIf error_of_wenti > 0 Then
           find_conclusion1 = 2 '问题有错
        Exit Function
      End If
    Else
        ty = 0
    End If
   ElseIf conclusion_data(i%).no(0) = 0 Then
     ty = 0
   End If
  Next i%
 find_conclusion1 = ty '推出结论
 If find_conclusion1 = 1 Then
   find_conclusion1 = 2
    finish_prove = 2
     Exit Function
    End If
'  Exit Function
 Else
  ty = 0
  'For i% = 0 To last_conclusion - 1
 '   If find_conclusion(i%, con_ty, con_no%, is_set_reduce) = 1 Then
 '    ty = 1
 '       find_conclusion1 = ty
 ' If find_conclusion1 = 1 Then
 '  find_conclusion1 = 2
 '   finish_prove = 2
 '    Exit Function
 '  End If
 '         Exit Function
 '   End If
 ' Next i%
  find_conclusion1 = 0
 End If
End If
End Function

Public Function find_conclusion_for_epolygon(ByVal conclu_ty As Byte, ByVal n%, _
       con_type As Byte, con_no%, is_set_reduce) As Byte
Dim i%, j%, k%, no%
Dim A(3) As Integer
Dim dn(2) As Integer
Dim last_eline As Integer
Dim ty As Boolean
Dim ty_ As Boolean
Dim v(3) As Integer
Dim temp_record As total_record_type
Dim temp_record1 As record_data_type
If conclu_ty = Squre Then
v(0) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).poi(0)
v(1) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).poi(1)
v(2) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).poi(2)
v(3) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).poi(3)
A(0) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).angle(0)
A(1) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).angle(1)
A(2) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).angle(2)
A(3) = Dpolygon4(con_squre(n%).data(0).polygon4_no).data(0).angle(3)
Else
v(0) = con_Epolygon(n%).data(0).p.v(0)
v(1) = con_Epolygon(n%).data(0).p.v(1)
v(2) = con_Epolygon(n%).data(0).p.v(2)
v(3) = con_Epolygon(n%).data(0).p.v(3)
 A(0) = Abs(angle_number(v(3), v(0), v(1), 0, 0))
 A(1) = Abs(angle_number(v(0), v(1), v(2), 0, 0))
 A(2) = Abs(angle_number(v(1), v(2), v(3), 0, 0))
 A(3) = Abs(angle_number(v(2), v(3), v(0), 0, 0))
End If
If con_type = 0 And is_set_reduce Then
 Call set_conclusion_point(n%, v(0))
 Call set_conclusion_point(n%, v(1))
 Call set_conclusion_point(n%, v(2))
 Call set_conclusion_point(n%, v(3))
 'Call set_condition_reduce(point_, v(0), 0, n% + 1)
 'Call set_condition_reduce(point_, v(1), 0, n% + 1)
 'Call set_condition_reduce(point_, v(2), 0, n% + 1)
 'Call set_condition_reduce(point_, v(3), 0, n% + 1)
ElseIf con_type = epolygon_ Or con_type = Squre Then
 If v(0) = epolygon(con_no%).data(0).p.v(0) = v(0) And _
     v(1) = epolygon(con_no%).data(0).p.v(0) = v(1) And _
      v(2) = epolygon(con_no%).data(0).p.v(0) = v(2) And _
       v(3) = epolygon(con_no%).data(0).p.v(0) = v(3) Then
  conclusion_data(n%).ty = epolygon_
  conclusion_data(n%).no(0) = no%
  find_conclusion_for_epolygon = 1
 End If
ElseIf con_type = line_value_ Or con_type = eline_ Or _
   con_type = angle3_value_ Then
If con_Epolygon(n%).data(0).p.total_v = 3 And con_type = epolygon_ Then
 If is_equal_sides_triangle(con_Epolygon(n%).data(0).no, no%, _
          temp_record.record_data.data0.condition_data) Then
  If no% > 0 Then
   conclusion_data(n%).ty = no%
  Else
   ty = True
  End If
 End If
Else 'If con_Epolygon(n%).data(0).p.total_v = 4 Then
   ty = False
    temp_record1.data0.condition_data.condition_no = 0
    If angle(A(0)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(0)).data(0).value_no, 0, 0, _
                                                       temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark1
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark1
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark1
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark1
    End If
find_conclusion_for_epolygon_mark1:
ty_ = False
last_eline = 0
  record_0.data0.condition_data.condition_no = 0 ' record0
For i% = 0 To 2
 For j% = i% + 1 To 3
    If is_equal_dline(v(i%), v((i% + 1) Mod 4), v(j%), v((j% + 1) Mod 4), _
             0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, _
              0, eline_data0, dn(1), dn(2), _
               cond_type, "", record_0.data0.condition_data) Then
         Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
          last_eline = last_eline + 1
           If last_eline = 3 Then
            ty_ = True
             GoTo find_conclusion_for_epolygon_next
           End If
   End If
 Next j%
Next i%
find_conclusion_for_epolygon_next:
If ty_ Then
 If ty Then
  Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
   If last_conditions.last_cond(1).epolygon_no Mod 10 = 0 Then
    ReDim Preserve epolygon(last_conditions.last_cond(1).epolygon_no + 10) As epolygon_type
   End If
last_conditions.last_cond(1).epolygon_no = last_conditions.last_cond(1).epolygon_no + 1
 no% = last_conditions.last_cond(1).epolygon_no
 epolygon(no%).data(0) = epolygon_data_0
epolygon(no%).data(0).p.v(0) = v(0) 'con_Epolygon(n%).data(0).p
epolygon(no%).data(0).p.v(1) = v(1)
epolygon(no%).data(0).p.v(2) = v(2)
epolygon(no%).data(0).p.v(3) = v(3)
epolygon(no%).data(0).p.total_v = 4
Call set_level(temp_record.record_data.data0.condition_data)
epolygon(no%).data(0).record = temp_record.record_data
 conclusion_data(n%).ty = epolygon_
 conclusion_data(n%).no(0) = no%
 find_conclusion_for_epolygon = 1
  Exit Function
 Else
 find_conclusion_for_epolygon = set_rhombus(v(0), v(1), v(2), v(3), temp_record, 0, 0)
  find_conclusion_for_epolygon = 0
 End If
End If
End If
ElseIf con_type = rhombus_ And con_Epolygon(n%).data(0).p.total_v = 4 Then
  If Dpolygon4(rhombus(con_no%).data(0).polygon4_no).data(0).poi(0) = v(0) And _
      Dpolygon4(rhombus(con_no%).data(0).polygon4_no).data(0).poi(1) = v(1) And _
        Dpolygon4(rhombus(con_no%).data(0).polygon4_no).data(0).poi(2) = v(2) And _
         Dpolygon4(rhombus(con_no%).data(0).polygon4_no).data(0).poi(3) = v(3) Then
   ty = False
    temp_record1.data0.condition_data.condition_no = 0
    If angle(A(0)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(0)).data(0).value_no, 0, 0, _
                                                       temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark2
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark2
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark2
    ElseIf angle(A(1)).data(0).value = "90" Then
        Call add_conditions_to_record(angle3_value_, angle(A(1)).data(0).value_no, 0, 0, _
                                                        temp_record1.data0.condition_data)
         ty = True
          GoTo find_conclusion_for_epolygon_mark2
    End If
find_conclusion_for_epolygon_mark2:
If ty Then
    temp_record.record_data.data0.condition_data.condition_no = 1
    temp_record.record_data.data0.condition_data.condition(1).ty = rhombus_
    temp_record.record_data.data0.condition_data.condition(1).no = no%
Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
If last_conditions.last_cond(1).epolygon_no Mod 10 = 0 Then
ReDim Preserve epolygon(last_conditions.last_cond(1).epolygon_no + 10) As epolygon_type
End If
last_conditions.last_cond(1).epolygon_no = last_conditions.last_cond(1).epolygon_no + 1
 no% = last_conditions.last_cond(1).epolygon_no
 epolygon(no%).data(0) = epolygon_data_0
epolygon(no%).data(0).p.v(0) = v(0)
epolygon(no%).data(0).p.v(1) = v(1)
epolygon(no%).data(0).p.v(2) = v(2)
epolygon(no%).data(0).p.v(3) = v(3)
epolygon(no%).data(0).p.total_v = 4
Call set_level(temp_record.record_data.data0.condition_data)
epolygon(no%).data(0).record = temp_record.record_data
 conclusion_data(n%).ty = epolygon_
 conclusion_data(n%).no(0) = no%
 find_conclusion_for_epolygon = 1
  Exit Function
 End If
 End If
ElseIf con_type = angle3_value_ Then
 If angle3_value(con_no%).data(0).data0.angle(1) = 0 And _
      angle3_value(con_no%).data(0).data0.value = "90" Then
        If angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).total_no = _
             angle(A(0)).data(0).total_no Or _
           angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).total_no = _
             angle(A(1)).data(0).total_no Or _
           angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).total_no = _
             angle(A(2)).data(0).total_no Or _
           angle(angle3_value(con_no%).data(0).data0.angle(0)).data(0).total_no = _
             angle(A(0)).data(0).total_no Then
 ty_ = False
 temp_record.record_data.data0.condition_data.condition_no = 0
 If is_rhombus(v(0), v(1), v(2), v(3), no%, 0, 0, 0) Then
  Call add_conditions_to_record(rhombus_, no%, 0, 0, _
     temp_record.record_data.data0.condition_data)
   ty_ = True
 Else
 last_eline = 0
  record_0.data0.condition_data.condition_no = 0 ' record0
  For i% = 0 To 2
   For j% = i% To 3
    If is_equal_dline(v(i%), v((i% + 1) Mod 4), v(j%), v((j% + 1) Mod 4), _
             0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, _
              0, eline_data0, dn(1), dn(2), _
               cond_type, "", record_0.data0.condition_data) Then
         Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
          last_eline = last_eline + 1
           If last_eline = 3 Then
            ty_ = True
             GoTo find_conclusion_for_epolygon_next
           End If
   End If
 Next j%
Next i%
  End If
  If ty_ And ty Then
         Call add_conditions_to_record(con_type, con_no%, 0, 0, temp_record.record_data.data0.condition_data)
   If last_conditions.last_cond(1).epolygon_no Mod 10 = 0 Then
    ReDim Preserve epolygon(last_conditions.last_cond(1).epolygon_no + 10) As epolygon_type
   End If
last_conditions.last_cond(1).epolygon_no = last_conditions.last_cond(1).epolygon_no + 1
 no% = last_conditions.last_cond(1).epolygon_no
 epolygon(no%).data(0) = epolygon_data_0
epolygon(no%).data(0).p.v(0) = v(0)
epolygon(no%).data(0).p.v(1) = v(1)
epolygon(no%).data(0).p.v(2) = v(2)
epolygon(no%).data(0).p.v(3) = v(3)
epolygon(no%).data(0).p.total_v = 4
Call set_level(temp_record.record_data.data0.condition_data)
epolygon(no%).data(0).record = temp_record.record_data
 conclusion_data(n%).ty = epolygon_
 conclusion_data(n%).no(0) = no%
 find_conclusion_for_epolygon = 1
  End If
 End If
End If
End If
End Function

Public Function find_conclusion_for_rhombus(ByVal n%, _
     con_type As Byte) As Byte
'rhombus_
Dim i%, j%, k%, no%
Dim dn(2) As Integer
Dim ty As Boolean
Dim temp_record As total_record_type
If con_type = line_value_ Or con_type = eline_ Or _
     con_type = angle3_value_ Then
For i% = 2 To 3
 For j% = 1 To i% - 1
  For k% = 0 To j% - 1
  record_0.data0.condition_data.condition_no = 0 ' record0
  temp_record.record_data.data0.condition_data.condition_no = 0
   If is_equal_dline(Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(k%), _
       Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((k% + 1) Mod 4), _
        Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((k% + 2) Mod 4), _
         Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((k% + 3) Mod 4), _
          0, 0, 0, 0, 0, 0, dn(0), _
           -1000, 0, 0, 0, eline_data0, dn(1), dn(2), cond_type, _
            "", temp_record.record_data.data0.condition_data) Then
           'temp_record.record_data.data0.condition_data.condition_no = 0
           'Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), _
            temp_record.record_data.data0.condition_data)
   'record_0 = record0
 If is_equal_dline(Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(j%), _
       Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((j% + 1) Mod 4), _
        Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((j% + 2) Mod 4), _
         Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((j% + 3) Mod 4), _
          0, 0, 0, 0, 0, 0, dn(0), _
          -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
            cond_type, "", temp_record.record_data.data0.condition_data) Then
           'temp_record.record_data.data0.condition_data.condition_no = 0
           'Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), _
             temp_record.record_data.data0.condition_data)
   'record_0 = record0
If is_equal_dline(Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi(i%), _
       Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((i% + 1) Mod 4), _
        Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((i% + 2) Mod 4), _
          Dpolygon4(con_rhombus(n%).data(0).polygon4_no).data(0).poi((i% + 3) Mod 4), _
           0, 0, 0, 0, 0, 0, dn(0), _
            -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
             cond_type, "", temp_record.record_data.data0.condition_data) Then
If last_conditions.last_cond(1).epolygon_no Mod 10 = 0 Then
ReDim Preserve rhombus(last_conditions.last_cond(1).epolygon_no + 10) As rhombus_type
End If
 last_conditions.last_cond(1).rhombus_no = last_conditions.last_cond(1).rhombus_no + 1
     no% = last_conditions.last_cond(1).rhombus_no
 rhombus(no%).data(0) = dpolygon4_data_0
rhombus(no%).data(0).polygon4_no = con_rhombus(n%).data(0).polygon4_no
Call set_level(temp_record.record_data.data0.condition_data)
rhombus(no%).data(0).record = temp_record.record_data
'rhombus(no%).record.other_no = no%
 conclusion_data(n%).no(0) = no%
 find_conclusion_for_rhombus = 1
  Exit Function

   End If
   End If
   End If
  Next k%
 Next j%
Next i%
End If
End Function

Public Function find_conclusion_for_long_squre(ByVal n%, _
      con_type As Byte) As Byte
Dim i%, j%, k%, no%
Dim dn(2) As Integer
Dim A(3) As Integer
Dim ty As Boolean
Dim temp_record As total_record_type
Dim temp_record1 As record_data_type
If con_type = line_value_ Or con_type = eline_ Or _
     con_type = angle3_value_ Then
 A(0) = Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).angle(0)
                         'Abs(angle_number(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(3), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(0), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(1), 0, 0))
 A(1) = Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).angle(1)
                         'Abs(angle_number(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(0), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(1), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(2), 0, 0))
 A(2) = Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).angle(2)
                         'Abs(angle_number(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(1), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(2), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(3), 0, 0))
 A(3) = Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).angle(3)
                         'Abs(angle_number(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(2), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(3), _
                         Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(0), 0, 0))
'record_0 = record0
temp_record.record_data.data0.condition_data.condition_no = 0
   If is_equal_dline(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(0), _
                     Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(1), _
                     Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(2), _
                     Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(3), _
                     0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
            cond_type, "", temp_record.record_data.data0.condition_data) Then
           ''temp_record.record_data.data0.condition_data.condition_no = 0
           'Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), _
            temp_record.record_data.data0.condition_data)
            'record_0 = record0
If is_equal_dline(Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(1), _
                  Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(2), _
                  Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(3), _
                  Dpolygon4(con_long_squre(n%).data(0).polygon4_no).data(0).poi(0), _
                  0, 0, 0, 0, 0, 0, dn(0), -1000, 0, 0, 0, eline_data0, _
                  dn(1), dn(2), cond_type, "", temp_record.record_data.data0.condition_data) Then
           'temp_record.record_data.data0.condition_data.condition_no = 0
           'Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), _
            temp_record.record_data.data0.condition_data)
    If is_angle_value(A(0), "90", "", dn(0), temp_record1.data0.condition_data) Then
        Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
         ty = True
    ElseIf is_angle_value(A(1), "90", "", dn(0), temp_record1.data0.condition_data) Then
        Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
         ty = True
    ElseIf is_angle_value(A(2), "90", "", dn(0), temp_record1.data0.condition_data) Then
        Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
         ty = True
    ElseIf is_angle_value(A(3), "90", "", dn(0), temp_record1.data0.condition_data) Then
        Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
         ty = True
    End If
    End If
    End If
If ty = True Then
If last_conditions.last_cond(1).long_squre_no Mod 10 = 0 Then
ReDim Preserve Dlong_squre(last_conditions.last_cond(1).long_squre_no + 10) As long_squre_type
End If
    last_conditions.last_cond(1).long_squre_no = last_conditions.last_cond(1).long_squre_no + 1
     no% = last_conditions.last_cond(1).long_squre_no
Dlong_squre(no%).data(0) = dpolygon4_data_0
Dlong_squre(no%).data(0).polygon4_no = con_long_squre(n%).data(0).polygon4_no
Call set_level(temp_record.record_data.data0.condition_data)
Dlong_squre(no%).data(0).record = temp_record.record_data
'Dlong_squre(no%).record.other_no = no%
 conclusion_data(n%).no(0) = no%
 find_conclusion_for_long_squre = 1
  Exit Function
End If
End If
End Function

Public Function find_conclusion_for_equal_side_triangle(ByVal n%) As Byte
Dim no%
Dim dn(2) As Integer
Dim tp(2) As Integer
Dim temp_record As total_record_type
  ' record_0 = record0
   temp_record.record_data.data0.condition_data.condition_no = 0
 Call read_triangle_element(con_equal_side_triangle(n%).data(0).triangle, con_equal_side_triangle(n%).data(0).direction, _
        tp(0), tp(1), tp(2), 0, 0, 0, 0, 0, 0, 0, 0, 0)
If is_equal_dline(tp(0), tp(1), tp(0), tp(2), _
               0, 0, 0, 0, 0, 0, dn(0), _
          -1000, 0, 0, 0, eline_data0, _
           dn(1), dn(2), cond_type, "", temp_record.record_data.data0.condition_data) Then
If last_conditions.last_cond(1).equal_side_triangle_no Mod 10 = 0 Then
  ReDim Preserve equal_side_triangle(last_conditions.last_cond(1).equal_side_triangle_no + 10) _
     As one_triangle_type
End If
last_conditions.last_cond(1).equal_side_triangle_no = last_conditions.last_cond(1).equal_side_triangle_no + 1
 no% = last_conditions.last_cond(1).equal_side_triangle_no
   equal_side_triangle(no%).data(0) = one_triangle_data_0
   equal_side_triangle(no%).data(0).triangle = con_equal_side_triangle(no%).data(0).triangle
   equal_side_triangle(no%).data(0).direction = con_equal_side_triangle(no%).data(0).direction
    Call set_level(temp_record.record_data.data0.condition_data)
   equal_side_triangle(no%).data(0).record = temp_record.record_data
   'equal_side_triangle(no%).record.other_no = n%
   conclusion_data(n%).no(0) = no%
   find_conclusion_for_equal_side_triangle = 1
    Exit Function
  End If
End Function

Public Function find_conclusion_for_equal_side_right_triangle(ByVal n%, _
       con_type As Byte, is_set_reduce As Boolean) As Byte
Dim no%
Dim dn(2) As Integer
Dim tp(2) As Integer
Dim A%
Dim temp_record As total_record_type
Dim temp_record1 As record_data_type
If con_type = 0 And is_set_reduce Then
   Call set_conclusion_point_for_triangle(n%, con_equal_side_right_triangle(n%).data(0).triangle)
   'Call set_triangle_reduce(con_equal_side_right_triangle(n%).data(0).triangle, 0, n% + 1)
ElseIf con_type = line_value_ Or con_type = eline_ Or con_type = angle3_value_ Then
  temp_record.record_data.data0.condition_data.condition_no = 0
   Call read_triangle_element(con_equal_side_right_triangle(n%).data(0).triangle, _
            con_equal_side_right_triangle(n%).data(0).direction, tp(0), tp(1), tp(2), A%, 0, _
             0, 0, 0, 0, 0, 0, 0)
 If is_equal_dline(tp(0), tp(1), tp(0), tp(2), _
              0, 0, 0, 0, 0, 0, dn(0), _
              -1000, 0, 0, 0, eline_data0, _
                 dn(1), dn(2), cond_type, "", temp_record.record_data.data0.condition_data) Then
           'Call add_conditions_to_record(cond_type, dn(0), dn(1), dn(2), _
              temp_record.record_data.data0.condition_data)
   If is_angle_value(A%, "90", "", dn(0), temp_record1.data0.condition_data) Then
         Call add_record_to_record(temp_record1.data0.condition_data, temp_record.record_data.data0.condition_data)
If last_conditions.last_cond(1).equal_side_right_triangle_no Mod 10 = 0 Then
  ReDim Preserve equal_side_right_triangle(last_conditions.last_cond(1).equal_side_right_triangle_no + 10) As one_triangle_type
End If
last_conditions.last_cond(1).equal_side_right_triangle_no = last_conditions.last_cond(1).equal_side_right_triangle_no + 1
 no% = last_conditions.last_cond(1).equal_side_right_triangle_no
   equal_side_right_triangle(no%).data(0) = one_triangle_data_0
   equal_side_right_triangle(no%).data(0).triangle = con_equal_side_right_triangle(no%).data(0).triangle
   equal_side_right_triangle(no%).data(0).direction = con_equal_side_right_triangle(no%).data(0).direction
   Call set_level(temp_record.record_data.data0.condition_data)
   equal_side_right_triangle(no%).data(0).record = temp_record.record_data
  ' equal_side_right_triangle(no%).record.other_no = n%
   conclusion_data(n%).no(0) = no%
   find_conclusion_for_equal_side_right_triangle = 1
    Exit Function
End If
   End If
End If
End Function

Public Function find_verti_foot(ByVal p1%, ByVal l%, op%, n%, no%) As Boolean
 'op% 垂足,n% 在l% 上编号op% no% 垂线号
Dim p%, k%
 For k% = 1 + last_conditions.last_cond(0).verti_no To last_conditions.last_cond(1).verti_no
  no% = Dverti(k%).data(0).record.data1.index.i(0)
  If Dverti(no%).data(0).line_no(0) = l% Then
   If is_point_in_line3(p1%, m_lin(Dverti(no%).data(0).line_no(1)).data(0).data0, 0) Then
   p% = is_line_line_intersect(l%, Dverti(no%).data(0).line_no(1), n%, 0, False)
    If p% > 0 Then
    op% = p%
     find_verti_foot = True
      Exit Function
     End If
    End If
  ElseIf Dverti(no%).data(0).line_no(1) = l% Then
   If is_point_in_line3(p1%, m_lin(Dverti(no%).data(0).line_no(0)).data(0).data0, 0) = True Then
    p% = is_line_line_intersect(l%, Dverti(no%).data(0).line_no(0), n%, 0, False)
    If p% > 0 Then
    op% = p%
     find_verti_foot = True
      Exit Function
     End If
    End If
    End If
 Next k%
 no% = 0
End Function
Public Function compare_two_three_angle_value(t1 As angle3_value_data0_type, _
    t2 As angle3_value_data0_type, ByVal k As Byte) As Integer
'Dim tl1(0)
Dim ty(1) As Integer
Dim n(3) As Integer
If k < 3 Then
n(0) = k
 n(1) = (k + 1) Mod 3
  n(2) = (k + 2) Mod 3
If t1.angle(n(0)) = t2.angle(n(0)) Then
 If t1.angle(n(1)) = t2.angle(n(1)) Then
  If t1.angle(n(2)) = t2.angle(n(2)) Then
   compare_two_three_angle_value = 0
  ElseIf t1.angle(n(2)) < t2.angle(n(2)) Then
   compare_two_three_angle_value = 1
  Else
   compare_two_three_angle_value = -1
  End If
 ElseIf t1.angle(n(1)) < t2.angle(n(1)) Then
  compare_two_three_angle_value = 1
 Else
  compare_two_three_angle_value = -1
 End If
ElseIf t1.angle(n(0)) < t2.angle(n(0)) Then
 compare_two_three_angle_value = 1
Else
 compare_two_three_angle_value = -1
End If
ElseIf k > 2 And k < 6 Then
If t1.angle(k) = t2.angle(k) Then
 If t1.angle(0) = t2.angle(0) Then
  If t1.angle(1) = t2.angle(1) Then
   If t1.angle(2) = t2.angle(2) Then
    compare_two_three_angle_value = 0
   ElseIf t1.angle(2) < t2.angle(2) Then
    compare_two_three_angle_value = 1
   Else
    compare_two_three_angle_value = -1
   End If
  ElseIf t1.angle(1) < t2.angle(1) Then
   compare_two_three_angle_value = 1
  Else
   compare_two_three_angle_value = -1
  End If
 ElseIf t1.angle(0) < t2.angle(0) Then
  compare_two_three_angle_value = 1
 Else
  compare_two_three_angle_value = -1
 End If
ElseIf t1.angle(3) < t2.angle(3) Then
 compare_two_three_angle_value = 1
Else
 compare_two_three_angle_value = -1
End If
End If
'ElseIf k = 4 Then
'If t1.value = t2.value Then
'If t1.angle(0) = t2.angle(0) Then
'If t1.angle(1) = t2.angle(1) Then
'If t1.angle(2) = t2.angle(2) Then
'compare_two_three_angle_value = 0
'ElseIf t1.angle(2) < t2.angle(2) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf t1.angle(1) < t2.angle(1) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf t1.angle(0) < t2.angle(0) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf t1.value < t2.value Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'Else
'n(0) = k - 5
' n(1) = (n(0) + 1) Mod 2
'If t1.para(1) = "0" And t1.para(2) = "0" Then
'ty(0) = compare_two_segment(angle(t1.angle(0)).data(0).poi(1), _
         angle(t1.angle(0)).data(0).line_no(n(0)), angle(t1.angle(0)).data(0).te(n(0)), _
          angle(t2.angle(0)).data(0).poi(1), angle(t2.angle(0)).data(0).line_no(n(0)), _
           angle(t2.angle(0)).data(0).te(n(0)))
'ty(1) = compare_two_segment(angle(t1.angle(0)).data(0).poi(1), _
         angle(t1.angle(0)).data(0).line_no(n(1)), angle(t1.angle(0)).data(0).te(n(1)), _
          angle(t2.angle(0)).data(0).poi(1), angle(t2.angle(0)).data(0).line_no(n(1)), _
           angle(t2.angle(0)).data(0).te(n(1)))
'If ty(0) = 0 Then
 'compare_two_three_angle_value = ty(1)
'Else
 'compare_two_three_angle_value = ty(0)
'End If
'Else
' compare_two_three_angle_value = -1
'End If

'If angle(t1.data.angle(0)).line_no(n(0)) < angle(t1.data.angle(1)).line_no(n(0)) Then

'    tl1(0) = angle(t1.data.angle(0)).line_no(n(0))
'    tl1(1) = angle(t1.data.angle(1)).line_no(n(0))
'  Else
'    tl1(0) = angle(t1.data.angle(1)).line_no(n(0))
'    tl1(1) = angle(t1.data.angle(0)).line_no(n(0))
'  End If
'  If angle(t2.data.angle(0)).line_no(n(0)) < angle(t2.data.angle(1)).line_no(n(0)) Then
'    tl2(0) = angle(t2.data.angle(0)).line_no(n(0))
'    tl2(1) = angle(t2.data.angle(1)).line_no(n(0))
'  Else
'    tl2(0) = angle(t2.data.angle(0)).line_no(n(0))
'    tl2(1) = angle(t2.data.angle(1)).line_no(n(0))
'  End If
'If tl1(0) = tl2(0) Then
'If tl1(2) = tl2(1) Then
'If t1.data.angle(0) = t2.data.angle(0) Then
'If t1.data.angle(1) = t2.data.angle(1) Then
'If t1.data.angle(2) = t2.data.angle(2) Then
'compare_two_three_angle_value = 0
'ElseIf t1.data.angle(2) < t2.data.angle(2) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf t1.data.angle(1) < t2.data.angle(1) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf t1.data.angle(0) < t2.data.angle(0) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf tl1(1) < tl2(1) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'ElseIf tl1(0) < tl2(0) Then
'compare_two_three_angle_value = 1
'Else
'compare_two_three_angle_value = -1
'End If
'End If
End Function
Public Function search_for_three_angle_value(t As angle3_value_data0_type, _
         ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2% 'k=7 and ty_=1 search_for_total_angle
'Dim k1 As Byte
Dim ty As Integer
'If k = 0 Then
'n1% = 1
'Else
n1% = 1 + last_conditions.last_cond(0).angle3_value_no
 'k1 = k - 1
'End If
n2% = last_conditions.last_cond(1).angle3_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_three_angle_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) Mod 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_three_angle_value(t, _
     angle3_value(angle3_value(n%).data(0).record.data1.index.i(k)).data(0).data0, k)  'k1=6
If ty = 0 Then
If ty_ = 0 Then
 n% = angle3_value(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_three_angle_value = True
   Exit Function
Else
 search_for_three_angle_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_three_angle_value = True Then
   search_for_three_angle_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_three_point_on_line(p3_l As three_point_on_line_data_type, _
   ByVal start%, ByVal k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).three_point_on_line_no
If n2% = 0 Then
 n% = 0
  search_for_three_point_on_line = False
   Exit Function
End If
While three_point_on_line(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) Mod 2
If n% = 0 Or n1% > n2% Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_three_point_on_line(p3_l, three_point_on_line(three_point_on_line(n%).data(0).record.data1.index.i(k%)).data(0), k%)
If ty = 0 Then
 If ty_ = 0 Then
 n% = three_point_on_line(n%).data(0).record.data1.index.i(k%)
 Else
 n% = n% - 1
 End If
  search_for_three_point_on_line = True
   Exit Function
Else
 search_for_three_point_on_line = judge_loop(n%, n1%, n2%, ty)
  If search_for_three_point_on_line = True Then
   search_for_three_point_on_line = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_general_string(g1 As general_string_data_type, _
     g2 As general_string_data_type, ByVal k%) As Integer
Dim n(3) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 4
  n(2) = (k% + 2) Mod 4
   n(3) = (k% + 3) Mod 4
If g1.item(n(0)) = g2.item(n(0)) Then
If g1.item(n(1)) = g2.item(n(1)) Then
If g1.item(n(2)) = g2.item(n(2)) Then
If g1.item(n(3)) = g2.item(n(3)) Then
If g1.value = g2.value Then
  compare_two_general_string = 0
ElseIf g1.value < g2.value Then
 compare_two_general_string = 1
Else
  'compare_two_general_string = -1
End If
ElseIf g1.item(n(3)) < g2.item(n(3)) Then
compare_two_general_string = 1
Else
compare_two_general_string = -1
End If
ElseIf g1.item(n(2)) < g2.item(n(2)) Then
compare_two_general_string = 1
Else
compare_two_general_string = -1
End If
ElseIf g1.item(n(1)) < g2.item(n(1)) Then
compare_two_general_string = 1
Else
compare_two_general_string = -1
End If
ElseIf g1.item(n(0)) < g2.item(n(0)) Then
compare_two_general_string = 1
Else
compare_two_general_string = -1
End If
End Function
Public Function search_for_general_string(g As general_string_data_type, _
                     ByVal k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).general_string_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_general_string = False
   Exit Function
End If
While general_string(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_general_string(g, general_string(general_string(n%).data(0).record.data1.index.i(k%)).data(0), k%)
If ty = 0 Then
If ty_ = 0 Then
 n% = general_string(n%).data(0).record.data1.index.i(k%)
Else
 n% = n% - 1
End If
  search_for_general_string = True
   Exit Function
Else
 search_for_general_string = judge_loop(n%, n1%, n2%, ty)
  If search_for_general_string = True Then
   search_for_general_string = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_triangle(triA_ As triangle_data0_type, _
     k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
Dim triA As triangle_data0_type
triA = triA_
n1% = 1 + last_conditions.last_cond(0).triangle_no
n2% = last_conditions.last_cond(1).triangle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_triangle = False
   Exit Function
End If
While triangle(n1%).data(0).index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_triangle(triA, triangle(triangle(n%).data(0).index.i(k%)).data(0), k%)
If ty = 0 Then
 If ty_ = 0 Then
  n% = triangle(n%).data(0).index.i(0)
 Else
  n% = n% - 1
 End If
  search_for_triangle = True
   Exit Function
Else
 search_for_triangle = judge_loop(n%, n1%, n2%, ty)
  If search_for_triangle = True Then
   search_for_triangle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_area_relation(triA_r As area_relation_data_type, _
              ByVal k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).area_relation_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_area_relation = False
   Exit Function
End If
While Darea_relation(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_area_relation(triA_r, Darea_relation(Darea_relation(n%).data(0).record.data1.index.i(k%)).data(0), k%)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Darea_relation(n%).data(0).record.data1.index.i(k%)
 Else
 n% = n% - 1
 End If
  search_for_area_relation = True
   Exit Function
Else
 search_for_area_relation = judge_loop(n%, n1%, n2%, ty)
  If search_for_area_relation = True Then
   search_for_area_relation = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_two_line_value(t_l_value As two_line_value_data0_type, _
           ByVal k As Byte, n%, ty_ As Byte) As Integer
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).two_line_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_two_line_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_two_line_value(t_l_value, _
    two_line_value(two_line_value(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = two_line_value(n%).data(0).record.data1.index.i(k)
 Else
 n% = n% - 1
 End If
  search_for_two_line_value = True
   Exit Function
Else
 search_for_two_line_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_two_line_value = True Then
   search_for_two_line_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_verti(vert As two_line_type, _
                       ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(1).verti_no
'k1 = k - 1
End If
n2% = last_conditions.last_cond(1).verti_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_verti = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_verti(vert, Dverti(Dverti(n%).data(0).record.data1.index.i(k)).data(0), k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Dverti(n%).data(0).record.data1.index.i(k)
 Else
 n% = n% - 1
 End If
  search_for_verti = True
   Exit Function
Else
 search_for_verti = judge_loop(n%, n1%, n2%, ty)
  If search_for_verti = True Then
   search_for_verti = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_arc_value(arc_v As arc_value_data_type, _
   ByVal start%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).arc_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_arc_value = False
   Exit Function
End If
While arc_value(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_arc_value(arc_v, arc_value(arc_value(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
 If ty_ = 0 Then
 n% = arc_value(n%).data(0).record.data1.index.i(0)
 Else
 n% = n% - 1
 End If
  search_for_arc_value = True
   Exit Function
Else
 search_for_arc_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_arc_value = True Then
   search_for_arc_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_arc(arc_ As arc_data_type, _
         ByVal start%, k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).arc_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_arc = False
   Exit Function
End If
While arc(n1%).data(0).index(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_arc(arc_, arc(arc(n%).data(0).index(k%)).data(0), k%)
If ty = 0 Then
If ty_ = 0 Then
 n% = arc(n%).data(0).index(k%)
Else
 n% = n% - 1
End If
  search_for_arc = True
   Exit Function
Else
 search_for_arc = judge_loop(n%, n1%, n2%, ty)
  If search_for_arc = True Then
   search_for_arc = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_area_element(area_ele As area_of_element_data_type, _
           ByVal start%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).area_of_element_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_area_element = False
   Exit Function
End If
While area_of_element(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_area_of_element(area_ele, area_of_element(area_of_element(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
If ty_ = 0 Then
 n% = area_of_element(n%).data(0).record.data1.index.i(0)
Else
 n% = n% - 1
End If
  search_for_area_element = True
   Exit Function
Else
 search_for_area_element = judge_loop(n%, n1%, n2%, ty)
  If search_for_area_element = True Then
   search_for_area_element = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_angle(A As angle_data_type, _
        n%, k%, ty_ As Byte, total_angle_no%, t_A As total_angle_data_type, insert_no%) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
'Dim t_A As total_angle_data_type
If A.line_no(0) < A.line_no(1) Then
 t_A.line_no(0) = A.line_no(0)
 t_A.line_no(1) = A.line_no(1)
Else
 t_A.line_no(0) = A.line_no(1)
 t_A.line_no(1) = A.line_no(0)
End If
total_angle_no% = 0
If search_for_total_angle(t_A, total_angle_no%) Then
   If total_angle_no% > 0 Then
      n% = T_angle(total_angle_no%).data(0).angle_no(A.total_no_).no
       If n% > 0 Then
        search_for_angle = True
       End If
   End If
Else
  t_A.line_no_constr_t_angle(0) = A.line_no(0)
  t_A.line_no_constr_t_angle(1) = A.line_no(1)
   insert_no% = total_angle_no%
   total_angle_no% = 0
End If
'n1% = 1 + last_conditions.last_cond(0).angle_no
'n2% = last_conditions.last_cond(1).angle_no
'If n2% = 0 Or n1% > n2% Then
' n% = 0
'  search_for_angle = False
'   Exit Function
'End If
'Do
'searh_for_angle_mark0:
'n% = n1% + (n2% - n1%) \ 2
'If n% = 0 Then
' If n2% = 0 Then
'  Exit Function
' Else
'  n% = 1
' End If
'End If
'ty = compare_two_angle(A, angle(angle(n%).data(0).index(k%)).data(0), k%)
'If ty = 0 Then
' If ty_ = 0 Then
'  n% = angle(n%).data(0).index(k%)
' Else
'  n% = n% - 1
' End If
'  search_for_angle = True
'   Exit Function
'Else
' search_for_angle = judge_loop(n%, n1%, n2%, ty)
'  If search_for_angle = True Then
'   search_for_angle = False
'    Exit Function
'  End If
'End If
'Loop
End Function
Public Function search_for_mid_point(mp As mid_point_data0_type, _
          ByVal k As Byte, n%, ty_ As Byte) As Boolean
           'ty_=0,1 核对三点,2核对两点
Dim n1%, n2%
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).mid_point_no
End If
n2% = last_conditions.last_cond(1).mid_point_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_mid_point = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_mid_point(mp, _
        Dmid_point(Dmid_point(n%).data(0).record.data1.index.i(k)).data(0).data0, k, ty_)
If ty = 0 Then
If ty_ = 0 Or ty_ = 2 Then
 n% = Dmid_point(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_mid_point = True
   Exit Function
Else
 search_for_mid_point = judge_loop(n%, n1%, n2%, ty)
  If search_for_mid_point = True Then
   search_for_mid_point = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_line3_value(l3_value As line3_value_data0_type, _
            ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).line3_value_no
'k1 = k - 1
End If
n2% = last_conditions.last_cond(1).line3_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_line3_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_line3_value(l3_value, _
       line3_value(line3_value(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = line3_value(n%).data(0).record.data1.index.i(k)
 Else
 n% = n% - 1
 End If
  search_for_line3_value = True
   Exit Function
Else
 search_for_line3_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_line3_value = True Then
   search_for_line3_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_item0(it As item0_data_type, _
          ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).item0_no
End If
n2% = last_conditions.last_cond(1).item0_no
If n2% = 0 Or _
    (it.poi(0) = 0 And it.poi(1) = 0 And _
        it.poi(2) = 0 And it.poi(3) = 0) Or n1% > n2% Then
 n% = 0
  search_for_item0 = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_item0(it, item0(item0(n%).data(0).index(k)).data(0), k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = item0(n%).data(0).index(k)
 Else
 n% = n% - 1
 End If
  search_for_item0 = True
   Exit Function
Else
 search_for_item0 = judge_loop(n%, n1%, n2%, ty)
  If search_for_item0 = True Then
   search_for_item0 = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_line_value(l_value As line_value_data0_type, _
            ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).line_value_no
'k1 = k - 1
End If
n2% = last_conditions.last_cond(1).line_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_line_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_line_value(l_value, _
       line_value(line_value(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
If ty_ = 0 Then
 n% = line_value(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_line_value = True
   Exit Function
Else
 search_for_line_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_line_value = True Then
   search_for_line_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_paral(pl As two_line_type, _
              ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1 = 1
Else
n1% = 1 + last_conditions.last_cond(0).paral_no
'k1 = k - 1
End If
n2% = last_conditions.last_cond(1).paral_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_paral = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_paral(pl, Dparal(Dparal(n%).data(0).data0.record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
If ty_ = 0 Then
 n% = Dparal(n%).data(0).data0.record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_paral = True
   Exit Function
Else
 search_for_paral = judge_loop(n%, n1%, n2%, ty)
  If search_for_paral = True Then
   search_for_paral = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_parallelogram(poly4_no As Integer, _
             n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).parallelogram_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_parallelogram = False
   Exit Function
End If
While Dparallelogram(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_integer(poly4_no, _
     Dparallelogram(Dparallelogram(n%).data(0).record.data1.index.i(0)).data(0).polygon4_no)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Dparallelogram(n%).data(0).record.data1.index.i(0)
 Else
 n% = n% - 1
 End If
  search_for_parallelogram = True
   Exit Function
Else
 search_for_parallelogram = judge_loop(n%, n1%, n2%, ty)
  If search_for_parallelogram = True Then
   search_for_parallelogram = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_polygon4(pal_gram As polygon4_data_type, _
             n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).polygon4_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_polygon4 = False
   Exit Function
End If
While Dpolygon4(n1%).data(0).index = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_polygon4(pal_gram, _
     Dpolygon4(Dpolygon4(n%).data(0).index).data(0), 0)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Dpolygon4(n%).data(0).index
 Else
 n% = n% - 1
 End If
  search_for_polygon4 = True
   Exit Function
Else
 search_for_polygon4 = judge_loop(n%, n1%, n2%, ty)
  If search_for_polygon4 = True Then
   search_for_polygon4 = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_point_pair(dp As point_pair_data0_type, _
                       ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).dpoint_pair_no
End If
n2% = last_conditions.last_cond(1).dpoint_pair_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_point_pair = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_point_pair(dp, _
    Ddpoint_pair(Ddpoint_pair(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
If ty_ = 0 Then
 n% = Ddpoint_pair(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_point_pair = True
   Exit Function
Else
 search_for_point_pair = judge_loop(n%, n1%, n2%, ty)
  If search_for_point_pair = True Then
   search_for_point_pair = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_angle_value(value As String, n%) As Boolean
Dim n1%, n2%
Dim k1 As Byte
Dim ty As Integer
'Dim V As Single
n1% = 1 + last_conditions.last_cond(0).angle_value_no
n2% = last_conditions.last_cond(1).angle_value_no
'V = Val(value)
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_angle_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_value(value, _
               angle3_value(angle_value.av_no(angle_value.av_no(n%).index).no).data(0).data0.value)
If ty = 0 Then
 n% = angle_value.av_no(n%).index
  search_for_angle_value = True
   Exit Function
Else
 search_for_angle_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_angle_value = True Then
   search_for_angle_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_relation(D_r As relation_data0_type, _
         ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).relation_no
End If
n2% = last_conditions.last_cond(1).relation_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_relation = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_relation(D_r, _
       Drelation(Drelation(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
If ty_ = 0 Then
 n% = Drelation(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_relation = True
   Exit Function
Else
 search_for_relation = judge_loop(n%, n1%, n2%, ty)
  If search_for_relation = True Then
   search_for_relation = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_total_equal_triangle(T_E_triA As two_triangle_type, _
           ByVal k%, n%, ty_ As Byte, is_find_conclusion As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
'Dim k1%
If k% = 0 Then
n1% = 1
'k1% = k%
Else
n1% = 1 + last_conditions.last_cond(0).total_equal_triangle_no
'k1% = k% - 1
End If
n2% = last_conditions.last_cond(1).total_equal_triangle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_total_equal_triangle = False
   Exit Function
End If
While Dtotal_equal_triangle(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_two_triangle(T_E_triA, Dtotal_equal_triangle( _
        Dtotal_equal_triangle(n%).data(0).record.data1.index.i(k%)).data(0), k%, is_find_conclusion)
If ty = 0 Then
If ty_ = 0 Then
 n% = Dtotal_equal_triangle(n%).data(0).record.data1.index.i(k%)
Else
 n% = n% - 1
End If
  search_for_total_equal_triangle = True
   Exit Function
Else
 search_for_total_equal_triangle = judge_loop(n%, n1%, n2%, ty)
  If search_for_total_equal_triangle = True Then
   search_for_total_equal_triangle = False
    Exit Function
  End If
End If
Loop
End Function

Public Function search_for_sides_length_of_triangle(s_l_triA As sides_length_of_triangle_data_type, _
       ByVal start%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).sides_length_of_triangle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_sides_length_of_triangle = False
   Exit Function
End If
While Sides_length_of_triangle(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_sides_length_of_triangle(s_l_triA, Sides_length_of_triangle(Sides_length_of_triangle(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
If ty_ = 0 Then
 n% = Sides_length_of_triangle(n%).data(0).record.data1.index.i(0)
Else
 n% = n% - 1
End If
  search_for_sides_length_of_triangle = True
   Exit Function
Else
 search_for_sides_length_of_triangle = judge_loop(n%, n1%, n2%, ty)
  If search_for_sides_length_of_triangle = True Then
   search_for_sides_length_of_triangle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_similar_triangle(s_triA As two_triangle_type, _
             ByVal k%, n%, ty_ As Byte, is_find_conclusion As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
'Dim k1%
If k% = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).similar_triangle_no
'k1% = k% - 1
End If
n2% = last_conditions.last_cond(1).similar_triangle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_similar_triangle = False
   Exit Function
End If
While Dsimilar_triangle(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_two_triangle(s_triA, Dsimilar_triangle( _
       Dsimilar_triangle(n%).data(0).record.data1.index.i(k%)).data(0), k%, is_find_conclusion)
If ty = 0 Then
If ty_ = 0 Then
 n% = Dsimilar_triangle(n%).data(0).record.data1.index.i(k%)
Else
 n% = n% - 1
End If
  search_for_similar_triangle = True
   Exit Function
Else
 search_for_similar_triangle = judge_loop(n%, n1%, n2%, ty)
  If search_for_similar_triangle = True Then
   search_for_similar_triangle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_epolygon(E_p As epolygon_data_type, _
    ByVal start%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).epolygon_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_epolygon = False
   Exit Function
End If
While epolygon(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_epolygon(E_p, epolygon(epolygon(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
If ty_ = 0 Then
 n% = epolygon(n%).data(0).record.data1.index.i(0)
Else
 n% = n% - 1
End If
  search_for_epolygon = True
   Exit Function
Else
 search_for_epolygon = judge_loop(n%, n1%, n2%, ty)
  If search_for_epolygon = True Then
   search_for_epolygon = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_equal_arc(e_arc As equal_arc_data_type, _
    ByVal start%, ByVal k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).equal_arc_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_equal_arc = False
   Exit Function
End If
While equal_arc(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_equal_arc(e_arc, equal_arc(equal_arc(n%).data(0).record.data1.index.i(k%)).data(0), k%)
If ty = 0 Then
 If ty_ = 0 Then
 n% = equal_arc(n%).data(0).record.data1.index.i(k%)
 Else
 n% = n% - 1
 End If
  search_for_equal_arc = True
   Exit Function
Else
 search_for_equal_arc = judge_loop(n%, n1%, n2%, ty)
  If search_for_equal_arc = True Then
   search_for_equal_arc = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_eline(el As eline_data0_type, _
            ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
n1% = 1 + last_conditions.last_cond(0).eline_no
n2% = last_conditions.last_cond(1).eline_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_eline = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_eline(el, Deline(Deline(n%).data(0).record.data1.index.i(k)).data(0).data0, k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Deline(n%).data(0).record.data1.index.i(k)
 Else
 n% = n% - 1
 End If
  search_for_eline = True
   Exit Function
Else
 search_for_eline = judge_loop(n%, n1%, n2%, ty)
  If search_for_eline = True Then
   search_for_eline = False
    Exit Function
  End If
End If
Loop
End Function
'Public Function search_for_equal_area_triangle(E_A_triangle As equal_area_triangle_data_type, _
'    ByVal start%, ByVal k%, n%, ty_ As Byte) As Boolean
'Dim n1%, n2%
'Dim ty As Integer
'n1% = start%
'n2% = last_conditions.last_cond(1).equal_area_triangle_no
'If n2% = 0 Or n1% > n2% Then
' n% = 0
'  search_for_equal_area_triangle = False
'   Exit Function
'End If
'While equal_area_triangle(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
' n1% = n1% + 1
'Wend
'Do
'n% = n1% + (n2% - n1%) \ 2
'If n% = 0 Then
' If n2% = 0 Then
'  Exit Function
' Else
'  n% = 1
' End If
'End If
'ty = compare_two_equal_area_triangle(E_A_triangle, equal_area_triangle(equal_area_triangle(n%).data(0).record.data1.index.i(k%)).data(0), k%)
'If ty = 0 Then
' If ty_ = 0 Then
' n% = equal_area_triangle(n%).data(0).record.data1.index.i(k%)
' Else
' n% = n% - 1
' End If
'  search_for_equal_area_triangle = True
'   Exit Function
'Else
' search_for_equal_area_triangle = judge_loop(n%, n1%, n2%, ty)
'  If search_for_equal_area_triangle = True Then
'   search_for_equal_area_triangle = False
'    Exit Function
'  End If
'End If
'Loop
'End Function
Public Function search_for_four_point_on_circle(p4_c As four_point_on_circle_data_type, _
     n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).four_point_on_circle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_four_point_on_circle = False
   Exit Function
End If
While four_point_on_circle(n1%).data(0).index = 0 '.record.data1.index.i(k%) = 0 'And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_four_point_on_circle(p4_c, _
        four_point_on_circle(four_point_on_circle(n%).data(0).index).data(0))
If ty = 0 Then
'If ty_ = 0 Then
 n% = four_point_on_circle(n%).data(0).index '.record.data1.index.i(k%)
'Else
' n% = n% - 1
'End If
  search_for_four_point_on_circle = True
   Exit Function
Else
 search_for_four_point_on_circle = judge_loop(n%, n1%, n2%, ty)
  If search_for_four_point_on_circle = True Then
   search_for_four_point_on_circle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_four_sides_fig(f_s_fig As four_sides_fig_data_type, _
   ByVal start%, ByVal k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).four_sides_fig_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_four_sides_fig = False
   Exit Function
End If
While four_sides_fig(n1%).data(0).index(k%) = 0 'And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_four_sides_fig(f_s_fig, _
        four_sides_fig(four_sides_fig(n%).data(0).index(k%)).data(0), k%)
If ty = 0 Then
If ty_ = 0 Then
 n% = four_sides_fig(n%).data(0).index(k%)
Else
 n% = n% - 1
End If
  search_for_four_sides_fig = True
   Exit Function
Else
 search_for_four_sides_fig = judge_loop(n%, n1%, n2%, ty)
  If search_for_four_sides_fig = True Then
   search_for_four_sides_fig = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_point_pair(dp1 As point_pair_data0_type, _
      dp2 As point_pair_data0_type, ByVal k%) As Integer
Dim n(3) As Integer
Dim i%
Dim ty(4) As Integer
If k% < 4 Then
 n(0) = k%
   n(1) = (k% + 1) Mod 4
     n(2) = (k% + 2) Mod 4
       n(3) = (k% + 3) Mod 4
ElseIf k% = 4 Then
 n(0) = 4
   n(1) = 5
     n(2) = 0
       n(3) = 2
ElseIf k% = 5 Then
 n(0) = 5
   n(1) = 4
     n(2) = 2
       n(3) = 0
End If
ty(0) = compare_two_point_(dp1.poi(2 * n(0)), dp1.poi(2 * n(0) + 1), _
    dp2.poi(2 * n(0)), dp2.poi(2 * n(0) + 1))
If ty(0) = 0 Then
 ty(1) = compare_two_point_(dp1.poi(2 * n(1)), dp1.poi(2 * n(1) + 1), _
     dp2.poi(2 * n(1)), dp2.poi(2 * n(1) + 1))
 If ty(1) = 0 Then
  ty(2) = compare_two_point_(dp1.poi(2 * n(2)), dp1.poi(2 * n(2) + 1), _
     dp2.poi(2 * n(2)), dp2.poi(2 * n(2) + 1))
  If ty(2) = 0 Then
    ty(3) = compare_two_point_(dp1.poi(2 * n(3)), dp1.poi(2 * n(3) + 1), _
     dp2.poi(2 * n(3)), dp2.poi(2 * n(3) + 1))
      compare_two_point_pair = ty(3)
  Else
      compare_two_point_pair = ty(2)
  End If
 Else
      compare_two_point_pair = ty(1)
 End If
Else
      compare_two_point_pair = ty(0)
End If
End Function

Public Function compare_two_angle(A1 As angle_data_type, _
    A2 As angle_data_type, k%) As Integer
If k% = 0 Then
If A1.poi(1) = A2.poi(1) Then '顶点
If A1.line_no(0) = A2.line_no(0) Then '两边
If A1.line_no(1) = A2.line_no(1) Then
If A1.te(0) = A2.te(0) Then '两点
If A1.te(1) = A2.te(1) Then
compare_two_angle = 0
ElseIf A1.te(1) < A2.te(1) Then
compare_two_angle = 1
Else
compare_two_angle = -1
End If
ElseIf A1.te(0) < A2.te(0) Then
compare_two_angle = 1
Else
compare_two_angle = -1
End If
ElseIf A1.line_no(1) < A2.line_no(1) Then
compare_two_angle = 1
Else
compare_two_angle = -1
End If
ElseIf A1.line_no(0) < A2.line_no(0) Then
compare_two_angle = 1
Else
compare_two_angle = -1
End If
ElseIf A1.poi(1) < A2.poi(1) Then
compare_two_angle = 1
Else
compare_two_angle = -1
End If
Else
 If A1.value = "" And A2.value <> "" Then
  compare_two_angle = 1
 ElseIf A2.value = "" And A1.value <> "" Then
  compare_two_angle = -1
 Else
   If A1.value < A2.value Then
    compare_two_angle = 1
   ElseIf A1.value > A2.value Then
    compare_two_angle = -1
   Else
    compare_two_angle = 0
   End If
 End If
End If
End Function

Public Function compare_two_triangle(triA1_ As triangle_data0_type, _
    triA2_ As triangle_data0_type, k%) As Integer
Dim n(2)
Dim triA1 As triangle_data0_type
Dim triA2 As triangle_data0_type
triA1 = triA1_
triA2 = triA2_
If k% < 3 Then
n(0) = k%
 n(1) = (n(0) + 1) Mod 3
  n(2) = (n(0) + 2) Mod 3
If triA1.poi(n(0)) < triA2.poi(n(0)) Then
 compare_two_triangle = 1
ElseIf triA1.poi(n(0)) > triA2.poi(n(0)) Then
 compare_two_triangle = -1
Else
If triA1.poi(n(1)) < triA2.poi(n(1)) Then
 compare_two_triangle = 1
ElseIf triA1.poi(n(1)) > triA2.poi(n(1)) Then
 compare_two_triangle = -1
Else
If triA1.poi(n(2)) < triA2.poi(n(2)) Then
 compare_two_triangle = 1
ElseIf triA1.poi(n(2)) > triA2.poi(n(2)) Then
 compare_two_triangle = -1
Else
 compare_two_triangle = 0
End If
End If
End If
Else
n(0) = k% - 3
 n(1) = (n(0) + 1) Mod 3
  n(2) = (n(0) + 2) Mod 3
 If triA1.angle(n(0)) = triA2.angle(n(0)) Then
  If triA1.angle(n(1)) = triA2.angle(n(1)) Then
   If triA1.angle(n(2)) = triA2.angle(n(2)) Then
      compare_two_triangle = 0
   ElseIf triA1.angle(n(0)) < triA2.angle(n(0)) Then
     compare_two_triangle = 1
   Else
     compare_two_triangle = -1
   End If
  ElseIf triA1.angle(n(1)) < triA2.angle(n(1)) Then
     compare_two_triangle = 1
  Else
     compare_two_triangle = -1
  End If
 ElseIf triA1.angle(n(2)) < triA2.angle(n(2)) Then
     compare_two_triangle = 1
 Else
     compare_two_triangle = -1
 End If
End If
End Function
Public Function compare_two_area_relation(triA_r1 As area_relation_data_type, _
    triA_r2 As area_relation_data_type, ByVal k%) As Integer
Dim n(2) As Integer
Dim ty(2) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 3
  n(2) = (k% + 2) Mod 3
ty(0) = compare_two_condition_type(triA_r1.area_element(n(0)), _
            triA_r2.area_element(n(0)))
ty(1) = compare_two_condition_type(triA_r1.area_element(n(1)), _
            triA_r2.area_element(n(1)))
ty(2) = compare_two_condition_type(triA_r1.area_element(n(2)), _
            triA_r2.area_element(n(2)))
If ty(0) = 0 Then
 If ty(1) = 0 Then
   compare_two_area_relation = ty(2)
 Else
   compare_two_area_relation = ty(1)
 End If
Else
 compare_two_area_relation = ty(0)
End If
End Function
Public Function compare_two_eline(el1 As eline_data0_type, _
           el2 As eline_data0_type, ByVal k%) As Integer
Dim n(3) As Integer
Dim ty(1) As Integer
If k% < 2 Then
 n(0) = k%
 n(1) = (k% + 1) Mod 2
 ty(0) = compare_two_point_(el1.poi(2 * n(0)), el1.poi(2 * n(0) + 1), _
                 el2.poi(2 * n(0)), el2.poi(2 * n(0) + 1))
 ty(1) = compare_two_point_(el1.poi(2 * n(1)), el1.poi(2 * n(1) + 1), _
                 el2.poi(2 * n(1)), el2.poi(2 * n(1) + 1))
 If ty(0) = 0 Then
  compare_two_eline = ty(1)
 Else
  compare_two_eline = ty(0)
 End If
ElseIf k% < 4 Then
 n(0) = k% - 2
 n(1) = (n(0) + 1) Mod 2
If el1.line_no(n(0)) = el2.line_no(n(0)) Then
If el1.line_no(n(1)) = el2.line_no(n(1)) Then
 ty(0) = compare_two_point_(el1.poi(2 * n(0)), el1.poi(2 * n(0) + 1), _
                 el2.poi(2 * n(0)), el2.poi(2 * n(0) + 1))
 ty(1) = compare_two_point_(el1.poi(2 * n(1)), el1.poi(2 * n(1) + 1), _
                 el2.poi(2 * n(1)), el2.poi(2 * n(1) + 1))
 If ty(0) = 0 Then
  compare_two_eline = ty(1)
 Else
  compare_two_eline = ty(0)
 End If
ElseIf el1.line_no(n(1)) < el2.line_no(n(1)) Then
  compare_two_eline = 1
Else
  compare_two_eline = -1
End If
ElseIf el1.line_no(n(0)) < el2.line_no(n(0)) Then
  compare_two_eline = 1
Else
  compare_two_eline = -1
End If
End If
End Function
Public Function compare_two_mid_point(mp1 As mid_point_data0_type, _
      mp2 As mid_point_data0_type, ByVal k%, ty As Byte) As Integer
'k=0全序 k=1 ,k=2 ,k=3 线段k=4,=5,=6, 点
If ty = 2 Then
 If k = 0 Then
 compare_two_mid_point = compare_two_point_(mp1.poi(0), mp1.poi(1), _
      mp2.poi(0), mp2.poi(1))
 ElseIf k% = 1 Then
  compare_two_mid_point = compare_two_point_(mp1.poi(1), mp1.poi(2), _
      mp2.poi(1), mp2.poi(2))
 ElseIf k% = 2 Then
  compare_two_mid_point = compare_two_point_(mp1.poi(0), mp1.poi(2), _
      mp2.poi(0), mp2.poi(2))
 End If
Else
If k% = 0 Then
 compare_two_mid_point = compare_two_point_(mp1.poi(0), mp1.poi(1), _
      mp2.poi(0), mp2.poi(1))
  If compare_two_mid_point = 0 Then
   If mp1.poi(2) < mp2.poi(2) Then
    compare_two_mid_point = 1
   ElseIf mp1.poi(2) > mp2.poi(2) Then
    compare_two_mid_point = -1
   End If
  End If
ElseIf k% = 1 Then
compare_two_mid_point = compare_two_point_(mp1.poi(1), mp1.poi(2), _
      mp2.poi(1), mp2.poi(2))
  If compare_two_mid_point = 0 Then
   If mp1.poi(0) < mp2.poi(0) Then
    compare_two_mid_point = 1
   ElseIf mp1.poi(0) > mp2.poi(0) Then
    compare_two_mid_point = -1
   End If
  End If
ElseIf k% = 2 Then
compare_two_mid_point = compare_two_point_(mp1.poi(0), mp1.poi(2), _
      mp2.poi(0), mp2.poi(2))
  If compare_two_mid_point = 0 Then
   If mp1.poi(1) < mp2.poi(1) Then
    compare_two_mid_point = 1
   ElseIf mp1.poi(1) > mp2.poi(1) Then
    compare_two_mid_point = -1
   End If
  End If
End If
End If
End Function

Public Function compare_two_four_point_on_circle( _
    p4_c1 As four_point_on_circle_data_type, _
      p4_c2 As four_point_on_circle_data_type) As Integer
Dim n(3) As Integer
If p4_c1.poi(0) = p4_c2.poi(0) Then
If p4_c1.poi(1) = p4_c2.poi(1) Then
If p4_c1.poi(2) = p4_c2.poi(2) Then
If p4_c1.poi(3) = p4_c2.poi(3) Then
compare_two_four_point_on_circle = 0
ElseIf p4_c1.poi(3) < p4_c2.poi(3) Then
compare_two_four_point_on_circle = 1
Else
compare_two_four_point_on_circle = -1
End If
ElseIf p4_c1.poi(2) < p4_c2.poi(2) Then
compare_two_four_point_on_circle = 1
Else
compare_two_four_point_on_circle = -1
End If
ElseIf p4_c1.poi(1) < p4_c2.poi(1) Then
compare_two_four_point_on_circle = 1
Else
compare_two_four_point_on_circle = -1
End If
ElseIf p4_c1.poi(0) < p4_c2.poi(0) Then
compare_two_four_point_on_circle = 1
Else
compare_two_four_point_on_circle = -1
End If

End Function
Public Function compare_two_paral(p1 As two_line_type, _
    p2 As two_line_type, ByVal k%) As Integer
Dim n(1) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 2
If p1.line_no(n(0)) = p2.line_no(n(0)) Then
If p1.line_no(n(1)) = p2.line_no(n(1)) Then
 compare_two_paral = 0
ElseIf p1.line_no(n(1)) < p2.line_no(n(1)) Then
 compare_two_paral = 1
Else
 compare_two_paral = -1
End If
ElseIf p1.line_no(n(0)) < p2.line_no(n(0)) Then
 compare_two_paral = 1
Else
 compare_two_paral = -1
End If
End Function
Public Function compare_two_polygon4(pal_gram1 As polygon4_data_type, _
       pal_gram2 As polygon4_data_type, k%) As Integer
Dim n(3) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 4
  n(2) = (k% + 2) Mod 4
   n(3) = (k% + 3) Mod 4
If pal_gram1.poi(n(0)) = pal_gram2.poi(n(0)) Then
If pal_gram1.poi(n(1)) = pal_gram2.poi(n(1)) Then
If pal_gram1.poi(n(2)) = pal_gram2.poi(n(2)) Then
If pal_gram1.poi(n(3)) = pal_gram2.poi(n(3)) Then
 compare_two_polygon4 = 0
ElseIf pal_gram1.poi(n(3)) < pal_gram2.poi(n(3)) Then
 compare_two_polygon4 = 1
Else
 compare_two_polygon4 = -1
End If
ElseIf pal_gram1.poi(n(2)) < pal_gram2.poi(n(2)) Then
 compare_two_polygon4 = 1
Else
 compare_two_polygon4 = -1
End If
ElseIf pal_gram1.poi(n(1)) < pal_gram2.poi(n(1)) Then
 compare_two_polygon4 = 1
Else
 compare_two_polygon4 = -1
End If
ElseIf pal_gram1.poi(n(0)) < pal_gram2.poi(n(0)) Then
 compare_two_polygon4 = 1
Else
 compare_two_polygon4 = -1
End If
End Function
Public Function compare_two_relation(d_r1 As relation_data0_type, _
      d_r2 As relation_data0_type, ByVal k%) As Integer
Dim n(2) As Integer
Dim ty(2) As Integer
Dim i%
'k%=0,比较poi(0),poi(1);k%=1,比较poi(2),poi(3);k%=2,比较poi(4),poi(5)(共线);
'k%=6;比较value
If k% <= 2 Then
  n(0) = k%
   n(1) = (k% + 1) Mod 3
    n(2) = (k% + 2) Mod 3
If d_r1.poi(2 * n(0)) < d_r2.poi(2 * n(0)) Or (d_r1.poi(2 * n(0)) = d_r2.poi(2 * n(0)) And _
     d_r1.poi(2 * n(0) + 1) < d_r2.poi(2 * n(0) + 1)) Then
     compare_two_relation = 1
ElseIf d_r1.poi(2 * n(0)) > d_r2.poi(2 * n(0)) Or (d_r1.poi(2 * n(0)) = d_r2.poi(2 * n(0)) And _
     d_r1.poi(2 * n(0) + 1) > d_r2.poi(2 * n(0) + 1)) Then
     compare_two_relation = -1
Else
If d_r1.poi(2 * n(1)) < d_r2.poi(2 * n(1)) Or (d_r1.poi(2 * n(1)) = d_r2.poi(2 * n(1)) And _
     d_r1.poi(2 * n(1) + 1) < d_r2.poi(2 * n(1) + 1)) Then
     compare_two_relation = 1
ElseIf d_r1.poi(2 * n(1)) > d_r2.poi(2 * n(1)) Or (d_r1.poi(2 * n(1)) = d_r2.poi(2 * n(1)) And _
     d_r1.poi(2 * n(1) + 1) > d_r2.poi(2 * n(1) + 1)) Then
     compare_two_relation = -1
Else
If d_r1.poi(2 * n(2)) < d_r2.poi(2 * n(2)) Or (d_r1.poi(2 * n(2)) = d_r2.poi(2 * n(2)) And _
     d_r1.poi(2 * n(2) + 1) < d_r2.poi(2 * n(2) + 1)) Then
     compare_two_relation = 1
ElseIf d_r1.poi(2 * n(2)) > d_r2.poi(2 * n(2)) Or (d_r1.poi(2 * n(2)) = d_r2.poi(2 * n(2)) And _
     d_r1.poi(2 * n(2) + 1) > d_r2.poi(2 * n(2) + 1)) Then
     compare_two_relation = -1
Else
     compare_two_relation = 0
End If
End If
End If
'For i% = 0 To 2
'ty(i%) = compare_two_point_(d_r1.poi(2 * n(i%)), d_r1.poi(2 * n(i%) + 1), _
 '   d_r2.poi(2 * n(i%)), d_r2.poi(2 * n(i%) + 1))
'Next i%
'If ty(0) = 0 Then
'   If ty(1) = 0 Then
'    compare_two_relation = ty(2)
'   Else
'    compare_two_relation = ty(1)
'   End If
'Else
'  compare_two_relation = ty(0)
'End If
ElseIf k% = 3 Then
If d_r1.value = d_r2.value Then
If d_r1.poi(0) = d_r2.poi(0) Then
If d_r1.poi(1) = d_r2.poi(1) Then
If d_r1.poi(2) = d_r2.poi(2) Then
If d_r1.poi(3) = d_r2.poi(3) Then
 compare_two_relation = 0
ElseIf d_r1.poi(3) < d_r2.poi(3) Then
 compare_two_relation = 1
Else
 compare_two_relation = -1
End If
ElseIf d_r1.poi(2) < d_r2.poi(2) Then
 compare_two_relation = 1
Else
 compare_two_relation = -1
End If
ElseIf d_r1.poi(1) < d_r2.poi(1) Then
 compare_two_relation = 1
Else
 compare_two_relation = -1
End If
ElseIf d_r1.poi(0) < d_r2.poi(0) Then
 compare_two_relation = 1
Else
 compare_two_relation = -1
End If
ElseIf d_r1.value < d_r2.value Then
 compare_two_relation = 1
Else
 compare_two_relation = -1
End If
End If
End Function

Public Function compare_two_two_triangle(S_triA1 As two_triangle_type, _
       S_triA2 As two_triangle_type, ByVal k%, is_find_conclusion As Byte) As Integer
Dim n(1) As Integer
 n(0) = k%
  n(1) = (k% + 1) Mod 2
If S_triA1.triangle(n(0)) = S_triA2.triangle(n(0)) Then
If S_triA1.triangle(n(1)) = S_triA2.triangle(n(1)) Then
If is_find_conclusion = 0 Then
 If S_triA1.direction = S_triA2.direction Then
  compare_two_two_triangle = 0
 ElseIf S_triA1.direction < S_triA2.direction Then
  compare_two_two_triangle = 1
 Else
  compare_two_two_triangle = -1
 End If
 Else
  compare_two_two_triangle = 0
 End If
ElseIf S_triA1.triangle(n(1)) < S_triA2.triangle(n(1)) Then
 compare_two_two_triangle = 1
Else
 compare_two_two_triangle = -1
End If
ElseIf S_triA1.triangle(n(0)) < S_triA2.triangle(n(0)) Then
 compare_two_two_triangle = 1
Else
 compare_two_two_triangle = -1
End If
End Function
Public Function compare_two_two_area_of_element(t_A_ele1 As two_area_element_value_data_type, _
       t_A_ele2 As two_area_element_value_data_type, ByVal k%) As Integer
Dim n(1) As Integer
 n(0) = k%
  n(1) = (k% + 1) Mod 2
If t_A_ele1.area_element(n(0)).element.no = t_A_ele2.area_element(n(0)).element.no And _
      t_A_ele1.area_element(n(0)).element.ty = t_A_ele2.area_element(n(0)).element.ty Then
 If t_A_ele1.area_element(n(1)).element.no = t_A_ele2.area_element(n(1)).element.no And _
     t_A_ele1.area_element(n(1)).element.ty = t_A_ele2.area_element(n(1)).element.ty Then
      compare_two_two_area_of_element = 0
 ElseIf t_A_ele1.area_element(n(1)).element.ty < t_A_ele2.area_element(n(1)).element.ty Then
      compare_two_two_area_of_element = 1
 ElseIf t_A_ele1.area_element(n(1)).element.ty > t_A_ele2.area_element(n(1)).element.ty Then
      compare_two_two_area_of_element = -1
 ElseIf t_A_ele1.area_element(n(1)).element.no < t_A_ele2.area_element(n(1)).element.no Then
      compare_two_two_area_of_element = 1
 ElseIf t_A_ele1.area_element(n(1)).element.no > t_A_ele2.area_element(n(1)).element.no Then
      compare_two_two_area_of_element = -1
 End If
ElseIf t_A_ele1.area_element(n(0)).element.ty < t_A_ele2.area_element(n(0)).element.ty Then
      compare_two_two_area_of_element = 1
ElseIf t_A_ele1.area_element(n(0)).element.ty > t_A_ele2.area_element(n(0)).element.ty Then
      compare_two_two_area_of_element = -1
ElseIf t_A_ele1.area_element(n(0)).element.no < t_A_ele2.area_element(n(0)).element.no Then
      compare_two_two_area_of_element = 1
ElseIf t_A_ele1.area_element(n(0)).element.no > t_A_ele2.area_element(n(0)).element.no Then
      compare_two_two_area_of_element = -1
End If
End Function

Public Function compare_two_three_point_on_line(p3_l1 As three_point_on_line_data_type, _
     p3_l2 As three_point_on_line_data_type, ByVal k%) As Integer
Dim n(2) As Integer
 n(0) = k%
  n(1) = (k% + 1) Mod 3
   n(2) = (k% + 2) Mod 3
If p3_l1.poi(n(0)) = p3_l2.poi(n(0)) Then
If p3_l1.poi(n(1)) = p3_l2.poi(n(1)) Then
If p3_l1.poi(n(2)) = p3_l2.poi(n(2)) Then
compare_two_three_point_on_line = 0
ElseIf p3_l1.poi(n(2)) < p3_l2.poi(n(2)) Then
compare_two_three_point_on_line = 1
Else
compare_two_three_point_on_line = -1
End If
ElseIf p3_l1.poi(n(1)) < p3_l2.poi(n(1)) Then
compare_two_three_point_on_line = 1
Else
compare_two_three_point_on_line = -1
End If
ElseIf p3_l1.poi(n(0)) < p3_l2.poi(n(0)) Then
compare_two_three_point_on_line = 1
Else
compare_two_three_point_on_line = -1
End If
End Function
Public Function compare_two_two_line_value(t_l_value1 As two_line_value_data0_type, _
       t_l_value2 As two_line_value_data0_type, ByVal k%) As Integer
Dim n(1) As Integer
Dim i%
Dim ty(1) As Integer
If k% < 2 Then
n(0) = k%
 n(1) = (k% + 1) Mod 2
For i% = 0 To 1
ty(i%) = compare_two_point_(t_l_value1.poi(2 * n(i%)), t_l_value1.poi(2 * n(i%) + 1), _
    t_l_value2.poi(2 * n(i%)), t_l_value2.poi(2 * n(i%) + 1))
Next i%
If ty(0) = 0 Then
  compare_two_two_line_value = ty(1)
Else
  compare_two_two_line_value = ty(0)
End If
Else
k% = k% - 2
n(0) = k%
 n(1) = (k% + 1) Mod 2
If t_l_value1.line_no(n(0)) = t_l_value2.line_no(n(0)) Then
If t_l_value1.line_no(n(1)) = t_l_value2.line_no(n(1)) Then
If t_l_value1.poi(2 * n(0)) = t_l_value2.poi(2 * n(0)) Then
If t_l_value1.poi(2 * n(0) + 1) = t_l_value2.poi(2 * n(0) + 1) Then
If t_l_value1.poi(2 * n(1)) = t_l_value2.poi(2 * n(1)) Then
If t_l_value1.poi(2 * n(1) + 1) = t_l_value2.poi(2 * n(1) + 1) Then
compare_two_two_line_value = 0
ElseIf t_l_value1.poi(2 * n(1) + 1) < t_l_value2.poi(2 * n(1) + 1) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
ElseIf t_l_value1.poi(2 * n(1)) < t_l_value2.poi(2 * n(1)) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
ElseIf t_l_value1.poi(2 * n(0) + 1) < t_l_value2.poi(2 * n(0) + 1) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
ElseIf t_l_value1.poi(2 * n(0)) < t_l_value2.poi(2 * n(0)) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
ElseIf t_l_value1.line_no(n(1)) < t_l_value2.line_no(n(1)) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
ElseIf t_l_value1.line_no(0) < t_l_value2.line_no(n(0)) Then
compare_two_two_line_value = 1
Else
compare_two_two_line_value = -1
End If
End If
End Function

Public Function compare_two_line3_value(l3_value1 As line3_value_data0_type, _
   l3_value2 As line3_value_data0_type, ByVal k%) As Integer
Dim n(2) As Integer
Dim i%
Dim ty(2) As Integer
If k% < 3 Then
n(0) = k%
 n(1) = (k% + 1) Mod 3
  n(2) = (k% + 2) Mod 3
For i% = 0 To 2
ty(i%) = compare_two_point_(l3_value1.poi(2 * n(i%)), l3_value1.poi(2 * n(i%) + 1), _
    l3_value2.poi(2 * n(i%)), l3_value2.poi(2 * n(i%) + 1))
Next i%
If ty(0) = 0 Then
 If ty(1) = 0 Then
  compare_two_line3_value = ty(2)
 Else
  compare_two_line3_value = ty(1)
 End If
Else
  compare_two_line3_value = ty(0)
End If
ElseIf k% = 3 Then
If l3_value1.poi(0) = l3_value2.poi(0) Then
If l3_value1.poi(1) = l3_value2.poi(1) Then
If l3_value1.poi(4) = l3_value2.poi(4) Then
If l3_value1.poi(5) = l3_value2.poi(5) Then
If l3_value1.poi(2) = l3_value2.poi(2) Then
If l3_value1.poi(3) = l3_value2.poi(3) Then
If l3_value1.para(0) = l3_value2.para(0) Then
If l3_value1.para(2) = l3_value2.para(2) Then
If l3_value1.para(1) = l3_value2.para(1) Then
compare_two_line3_value = 0
ElseIf l3_value1.para(n(2)) < l3_value2.para(n(2)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.para(n(1)) < l3_value2.para(n(1)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.para(n(0)) < l3_value2.para(n(0)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(2) + 1) < l3_value2.poi(2 * n(2) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(2)) < l3_value2.poi(2 * n(2)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(1) + 1) < l3_value2.poi(2 * n(1) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(1)) < l3_value2.poi(2 * n(1)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(0) + 1) < l3_value2.poi(2 * n(0) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(0)) < l3_value2.poi(2 * n(0)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
Else
k% = k% - 4
n(0) = k%
 n(1) = (k% + 1) Mod 3
  n(2) = (k% + 2) Mod 3
If l3_value1.line_no(n(0)) = l3_value2.line_no(n(0)) Then
If l3_value1.line_no(n(1)) = l3_value2.line_no(n(1)) Then
If l3_value1.line_no(n(2)) = l3_value2.line_no(n(2)) Then
If l3_value1.poi(2 * n(0)) = l3_value2.poi(2 * n(0)) Then
If l3_value1.poi(2 * n(0) + 1) = l3_value2.poi(2 * n(0) + 1) Then
If l3_value1.poi(2 * n(1)) = l3_value2.poi(2 * n(1)) Then
If l3_value1.poi(2 * n(1) + 1) = l3_value2.poi(2 * n(1) + 1) Then
If l3_value1.poi(2 * n(2)) = l3_value2.poi(2 * n(2)) Then
If l3_value1.poi(2 * n(2) + 1) = l3_value2.poi(2 * n(2) + 1) Then
compare_two_line3_value = 0
ElseIf l3_value1.poi(2 * n(2) + 1) < l3_value2.poi(2 * n(2) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(2)) < l3_value2.poi(2 * n(2)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(1) + 1) < l3_value2.poi(2 * n(1) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(1)) < l3_value2.poi(2 * n(1)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(0) + 1) < l3_value2.poi(2 * n(0) + 1) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.poi(2 * n(0)) < l3_value2.poi(2 * n(0)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.line_no(n(2)) < l3_value2.line_no(n(2)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.line_no(n(1)) < l3_value2.line_no(n(1)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
ElseIf l3_value1.line_no(n(0)) < l3_value2.line_no(n(0)) Then
compare_two_line3_value = 1
Else
compare_two_line3_value = -1
End If
End If
End Function
Public Function compare_two_line_value(l_value1 As line_value_data0_type, _
    l_value2 As line_value_data0_type, ByVal k%) As Integer
Dim n(1) As Integer
If k% < 2 Then
n(0) = k%
 n(1) = (k% + 1) Mod 2
If l_value1.poi(n(0)) = l_value2.poi(n(0)) Then
If l_value1.poi(n(1)) = l_value2.poi(n(1)) Then
 compare_two_line_value = 0
ElseIf l_value1.poi(n(1)) < l_value2.poi(n(1)) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf l_value1.poi(n(0)) < l_value2.poi(n(0)) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf k% = 2 Then
If l_value1.line_no = l_value2.line_no Then
If l_value1.poi(0) = l_value2.poi(0) Then
If l_value1.poi(1) = l_value2.poi(1) Then
 compare_two_line_value = 0
ElseIf l_value1.poi(1) < l_value2.poi(1) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf l_value1.poi(0) < l_value2.poi(0) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf l_value1.line_no < l_value2.line_no Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf k% = 3 Then
If l_value1.value = l_value2.value Then
If l_value1.poi(0) = l_value2.poi(0) Then
If l_value1.poi(1) = l_value2.poi(1) Then
 compare_two_line_value = 0
ElseIf l_value1.poi(1) < l_value2.poi(1) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf l_value1.poi(0) < l_value2.poi(0) Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
ElseIf l_value1.value < l_value2.value Then
 compare_two_line_value = 1
Else
 compare_two_line_value = -1
End If
End If
End Function
Public Function compare_two_V_line_value(l_value1 As V_line_value_data0_type, _
    l_value2 As V_line_value_data0_type, ByVal k%) As Integer
If k% < 2 Then
If l_value1.v_poi(k%) < l_value2.v_poi(k%) Then
 compare_two_V_line_value = 1
ElseIf l_value1.v_poi(k%) > l_value2.v_poi(k%) Then
 compare_two_V_line_value = -1
Else
 If l_value1.v_poi((k% + 1) Mod 2) < l_value2.v_poi((k% + 1) Mod 2) Then
  compare_two_V_line_value = 1
 ElseIf l_value1.v_poi((k% + 1) Mod 2) > l_value2.v_poi((k% + 1) Mod 2) Then
  compare_two_V_line_value = -1
 Else
  compare_two_V_line_value = 0
 End If
End If
ElseIf k% = 2 Then
 If l_value1.unit_value < l_value2.unit_value Then
  compare_two_V_line_value = 1
 ElseIf l_value1.unit_value > l_value2.unit_value Then
  compare_two_V_line_value = -1
 Else
  compare_two_V_line_value = 0
 End If
End If
End Function
Public Function compare_two_verti(vert1 As two_line_type, _
    vert2 As two_line_type, ByVal k%) As Integer
Dim n(1) As Integer
 n(0) = k%
  n(1) = (k% + 1) Mod 2
If vert1.line_no(n(0)) = vert2.line_no(n(0)) Then
If vert1.line_no(n(1)) = vert2.line_no(n(1)) Then
compare_two_verti = 0
ElseIf vert1.line_no(n(1)) < vert2.line_no(n(1)) Then
compare_two_verti = 1
Else
compare_two_verti = -1
End If
ElseIf vert1.line_no(n(0)) < vert2.line_no(n(0)) Then
compare_two_verti = 1
Else
compare_two_verti = -1
End If
End Function
Public Function compare_two_value(v1$, v2$) As Integer
If v1$ = v2$ Then
compare_two_value = 0
ElseIf v1$ < v2$ Then
compare_two_value = 1
Else
compare_two_value = -1
End If
End Function

Public Function compare_two_arc(arc_1 As arc_data_type, _
   arc_2 As arc_data_type, k%) As Integer
Dim n(1) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 2
If arc_1.cir = arc_2.cir Then
If arc_1.poi(n(0)) = arc_2.poi(n(0)) Then
If arc_1.poi(n(1)) = arc_2.poi(n(1)) Then
compare_two_arc = 0
ElseIf arc_1.poi(n(1)) < arc_2.poi(n(1)) Then
compare_two_arc = 1
Else
compare_two_arc = -1
End If
ElseIf arc_1.poi(n(0)) < arc_2.poi(n(0)) Then
compare_two_arc = 1
Else
compare_two_arc = -1
End If
ElseIf arc_1.cir < arc_2.cir Then
compare_two_arc = 1
Else
compare_two_arc = -1
End If
End Function
Public Function compare_two_arc_value(arc_v1 As arc_value_data_type, _
   arc_v2 As arc_value_data_type) As Integer
If arc_v1.arc = arc_v2.arc Then
compare_two_arc_value = 0
ElseIf arc_v1.arc < arc_v2.arc Then
compare_two_arc_value = 1
Else
compare_two_arc_value = -1
End If
End Function

Public Function compare_two_equal_arc(e_arc1 As equal_arc_data_type, _
       e_arc2 As equal_arc_data_type, ByVal k%) As Integer
Dim n(3) As Integer
n(0) = k%
 n(1) = (k% + 1) Mod 2
If e_arc1.arc(n(0)) = e_arc2.arc(n(0)) Then
If e_arc1.arc(n(1)) = e_arc2.arc(n(1)) Then
compare_two_equal_arc = 0
ElseIf e_arc1.arc(n(1)) < e_arc2.arc(n(1)) Then
compare_two_equal_arc = 1
Else
compare_two_equal_arc = -1
End If
ElseIf e_arc1.arc(n(0)) < e_arc2.arc(n(0)) Then
compare_two_equal_arc = 1
Else
compare_two_equal_arc = -1
End If
End Function
Public Function compare_two_epolygon(e_p1 As epolygon_data_type, _
       e_p2 As epolygon_data_type) As Integer
If e_p1.p.v(0) = e_p2.p.v(0) Then
If e_p1.p.v(1) = e_p2.p.v(1) Then
If e_p1.p.v(2) = e_p2.p.v(2) Then
If e_p1.p.v(3) = e_p2.p.v(3) Then
If e_p1.p.v(4) = e_p2.p.v(4) Then
If e_p1.p.v(5) = e_p2.p.v(5) Then
compare_two_epolygon = 0
ElseIf e_p1.p.v(5) < e_p2.p.v(5) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
ElseIf e_p1.p.v(4) < e_p2.p.v(4) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
ElseIf e_p1.p.v(3) < e_p2.p.v(3) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
ElseIf e_p1.p.v(2) < e_p2.p.v(2) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
ElseIf e_p1.p.v(1) < e_p2.p.v(1) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
ElseIf e_p1.p.v(0) < e_p2.p.v(0) Then
compare_two_epolygon = 1
Else
compare_two_epolygon = -1
End If
End Function

Public Function compare_two_area_of_element(area_A1 As area_of_element_data_type, _
     area_A2 As area_of_element_data_type) As Integer
compare_two_area_of_element = compare_two_condition_type(area_A1.element, _
     area_A2.element)
End Function

Public Function compare_two_sides_length_of_triangle(s_l_A1 As sides_length_of_triangle_data_type, _
    s_l_A2 As sides_length_of_triangle_data_type) As Integer
If s_l_A1.triangle = s_l_A2.triangle Then
compare_two_sides_length_of_triangle = 0
ElseIf s_l_A1.triangle < s_l_A2.triangle Then
compare_two_sides_length_of_triangle = 1
Else
compare_two_sides_length_of_triangle = -1
End If
End Function
Public Function compare_two_item0(it1 As item0_data_type, it2 As item0_data_type, _
      k As Byte) As Integer
Dim n(2) As Integer
n(0) = k
n(1) = (k + 1) Mod 3
n(2) = (k + 2) Mod 3
If k < 3 Then
If it1.poi(2 * n(0)) < it2.poi(2 * n(0)) Then
compare_two_item0 = 1
ElseIf it1.poi(2 * n(0)) > it2.poi(2 * n(0)) Then
compare_two_item0 = -1
Else 'If it1.poi(2 * n(0)) < it2.poi(2 * n(0)) Then
 If it1.poi(2 * n(0) + 1) < it2.poi(2 * n(0) + 1) Then
  compare_two_item0 = 1
 ElseIf it1.poi(2 * n(0) + 1) > it2.poi(2 * n(0) + 1) Then
  compare_two_item0 = -1
 Else 'If it1.poi(2 * n(0) + 1) = it2.poi(2 * n(0) + 1) Then
  If it1.poi(2 * n(1)) < it2.poi(2 * n(1)) Then
   compare_two_item0 = 1
  ElseIf it1.poi(2 * n(1)) > it2.poi(2 * n(1)) Then
   compare_two_item0 = -1
  Else 'If it1.poi(2 * n(1)) = it2.poi(2 * n(1)) Then
   If it1.poi(2 * n(1) + 1) < it2.poi(2 * n(1) + 1) Then
    compare_two_item0 = 1
   ElseIf it1.poi(2 * n(1) + 1) > it2.poi(2 * n(1) + 1) Then
    compare_two_item0 = -1
   Else 'If it1.poi(2 * n(1) + 1) = it2.poi(2 * n(1) + 1) Then
    If it1.poi(2 * n(2)) < it2.poi(2 * n(2)) Then
    compare_two_item0 = 1
    ElseIf it1.poi(2 * n(2)) > it2.poi(2 * n(2)) Then
    compare_two_item0 = -1
    Else 'If it1.poi(2 * n(2)) = it2.poi(2 * n(2)) Then
     If it1.poi(2 * n(2) + 1) < it2.poi(2 * n(2) + 1) Then
      compare_two_item0 = 1
     ElseIf it1.poi(2 * n(2) + 1) > it2.poi(2 * n(2) + 1) Then
      compare_two_item0 = -1
     Else ' If it1.poi(2 * n(2) + 1) = it2.poi(2 * n(2) + 1) Then
      If it1.sig <> "" And it2.sig <> "" Then
       If it1.sig < it2.sig Then
        compare_two_item0 = 1
       ElseIf it1.sig > it2.sig Then
        compare_two_item0 = -1
       Else
        compare_two_item0 = 0
       End If
      ElseIf it2.sig <> "" Then
       compare_two_item0 = 1
      ElseIf it1.sig <> "" Then
       compare_two_item0 = -1
      Else
       compare_two_item0 = 0
      End If
     End If
    End If
  End If
 End If
 End If
 End If
Else
If it1.sig = it2.sig Then
If it1.poi(0) = it2.poi(0) Then
If it1.poi(1) = it2.poi(1) Then
If it1.poi(2) = it2.poi(2) Then
If it1.poi(3) = it2.poi(3) Then
compare_two_item0 = 0
ElseIf it1.poi(3) < it2.poi(3) Then
compare_two_item0 = 1
Else
compare_two_item0 = -1
End If
ElseIf it1.poi(2) < it2.poi(2) Then
compare_two_item0 = 1
Else
compare_two_item0 = -1
End If
ElseIf it1.poi(1) < it2.poi(1) Then
compare_two_item0 = 1
Else
compare_two_item0 = -1
End If
ElseIf it1.poi(0) < it2.poi(0) Then
compare_two_item0 = 1
Else
compare_two_item0 = -1
End If
ElseIf it1.sig < it2.sig Then
compare_two_item0 = 1
Else
compare_two_item0 = -1
End If
End If
End Function
Public Function compare_two_line_from_two_point(L_tp1 As line_from_two_point, _
  L_tp2 As line_from_two_point) As Integer
If L_tp1.poi(0) = L_tp2.poi(0) Then
If L_tp1.poi(1) = L_tp2.poi(1) Then
compare_two_line_from_two_point = 0
ElseIf L_tp1.poi(1) < L_tp2.poi(1) Then
compare_two_line_from_two_point = 1
Else
compare_two_line_from_two_point = -1
End If
ElseIf L_tp1.poi(0) < L_tp2.poi(0) Then
compare_two_line_from_two_point = 1
Else
compare_two_line_from_two_point = -1
End If
End Function
Public Function search_for_two_point_line(ByVal p1%, ByVal p2%, _
                     n%, ty_ As Byte) As Boolean '
Dim n1%, n2%
Dim ty As Integer
Dim L_two_p As line_from_two_point
If p1% > 90 Or p2% > 90 Then '用于辅助线，只有一个实点
 For n% = 1 To last_conditions.last_cond(1).line_no
  For n1% = 1 To m_lin(n%).data(0).data0.in_point(0) + 1
  For n2% = n1% + 1 To m_lin(n%).data(0).data0.in_point(0) + 1
  If is_same_two_point(p1%, p2%, _
   m_lin(n%).data(0).data0.in_point(n1%), m_lin(n%).data(0).data0.in_point(n2%)) Then
    search_for_two_point_line = n%
     Exit Function
  End If
  Next n2%
  Next n1%
 Next n%
 n% = 0
 Exit Function
Else
If p1% < p2% Then
L_two_p.poi(0) = p1%
 L_two_p.poi(1) = p2%
Else
L_two_p.poi(0) = p2%
 L_two_p.poi(1) = p1%
End If
n1% = 1 + last_conditions.last_cond(0).line_from_two_point_no
n2% = last_conditions.last_cond(1).line_from_two_point_no
If n2% = 0 Then
 n% = 0
  search_for_two_point_line = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_line_from_two_point(L_two_p, _
                    Dtwo_point_line( _
                     Dtwo_point_line(n%).data(0).index))
If ty = 0 Then
 If ty_ = 0 Then
 n% = Dtwo_point_line(n%).data(0).index
   search_for_two_point_line = Dtwo_point_line(n%).data(0).line_no
 Else
 n% = n% - 1
 End If
   search_for_two_point_line = True
    Exit Function
Else
  If judge_loop(n%, n1%, n2%, ty) Then
  'If search_for_two_point_line > 0 Then
   search_for_two_point_line = False
    Exit Function
  End If
End If
Loop
End If

End Function
Public Function judge_loop(n%, n1%, n2%, ty%) As Boolean
If ty = 1 Then
 If n1% = n2% Or n% = n1% Then
  n% = n% - 1
   judge_loop = True 'n%之前
 Else
  n2% = n% - 1
   judge_loop = False
 End If
ElseIf ty = -1 Then
 If n1% = n2% Or n% = n2% Then
    judge_loop = True
 Else
  n1% = n% + 1
    judge_loop = False
 End If
End If
End Function
Private Function compare_two_point_(ByVal p1%, ByVal p2%, ByVal p3%, _
          ByVal p4%) As Integer
If p1% < p3% Then
 compare_two_point_ = 1
ElseIf p1% = p3% Then
 If p2% < p4% Then
  compare_two_point_ = 1
 ElseIf p2% = p4% Then
  compare_two_point_ = 0
 Else
  compare_two_point_ = -1
 End If
Else
 compare_two_point_ = -1
End If
End Function

Public Function compare_two_Rtriangle(Rtriangle1 As Rtriangle_data_type, _
          Rtriangle2 As Rtriangle_data_type, k%) As Integer
If k% = 0 Then
If Rtriangle1.triangle = Rtriangle2.triangle Then
compare_two_Rtriangle = 0
ElseIf Rtriangle1.triangle < Rtriangle2.triangle Then
compare_two_Rtriangle = 1
Else
compare_two_Rtriangle = -1
End If
End If
End Function
Public Function search_for_Rtriangle(RtriA As Rtriangle_data_type, _
       start%, k%, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).rtriangle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_Rtriangle = False
   Exit Function
End If
While Rtriangle(n1%).data(0).record.data1.index.i(k%) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_Rtriangle(RtriA, Rtriangle(Rtriangle(n%).data(0).record.data1.index.i(k%)).data(0), k%)
If ty = 0 Then
If ty_ = 0 Then
 n% = Rtriangle(n%).data(0).record.data1.index.i(k%)
Else
 n% = n% - 1
End If
  search_for_Rtriangle = True
   Exit Function
Else
 search_for_Rtriangle = judge_loop(n%, n1%, n2%, ty)
  If search_for_Rtriangle = True Then
   search_for_Rtriangle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_segment(ByVal p1%, ByVal l1%, ByVal te1 As Byte, _
      ByVal p2%, ByVal l2%, ByVal te2 As Byte) As Integer
If p1% = p2% Then
If l1% = l2% Then
If te1 = te2 Then
 compare_two_segment = 0
ElseIf te1 < te2 Then
 compare_two_segment = 1
Else
 compare_two_segment = -1
End If
ElseIf l1% < l2% Then
 compare_two_segment = 1
Else
 compare_two_segment = -1
End If
ElseIf p1% < p2% Then
 compare_two_segment = 1
Else
 compare_two_segment = -1
End If
End Function
Public Function compare_two_condition(con1 As condition_type, _
          con2 As condition_type) As Integer
If con1.ty = con2.ty Then
If con1.no = con2.no Then
 compare_two_condition = 0
ElseIf con1.no < con2.no Then
 compare_two_condition = 1
Else
 compare_two_condition = -1
End If
ElseIf con1.ty < con2.ty Then
 compare_two_condition = 1
Else
 compare_two_condition = -1
End If
End Function
Public Function compare_two_record_(re1 As record_data_type, re2 As record_data_type) As Integer
Dim ty(8) As Integer
Dim i%
If re1.data0.condition_data.condition_no = re1.data0.condition_data.condition_no Then
If re1.data0.condition_data.condition_no = 0 Then
If re1.data0.theorem_no = re2.data0.theorem_no Then
 compare_two_record_ = 0
ElseIf re1.data0.theorem_no < re2.data0.theorem_no Then
 compare_two_record_ = 1
Else
 compare_two_record_ = -1
End If
Else
For i% = 1 To re1.data0.condition_data.condition_no
 ty(i%) = compare_two_condition(re1.data0.condition_data.condition(i%), re2.data0.condition_data.condition(i%))
 If ty(i%) <> 0 Then
  compare_two_record_ = ty(i%)
   Exit Function
  End If
Next i%
  compare_two_record_ = 0
End If
ElseIf re1.data0.condition_data.condition_no < re2.data0.condition_data.condition_no Then
 compare_two_record_ = 1
Else
 compare_two_record_ = -1
End If
End Function
Public Function find_conclusion_for_line3_value(con_ty As Byte, no%) As Byte
End Function



Public Function compare_two_two_circle(A_c1 As add_point_for_two_circle_type, _
      A_c2 As add_point_for_two_circle_type) As Integer
If A_c1.circ(0) < A_c2.circ(0) Then
compare_two_two_circle = 1
ElseIf A_c1.circ(0) > A_c2.circ(0) Then
compare_two_two_circle = -1
Else
If A_c1.circ(1) < A_c2.circ(1) Then
compare_two_two_circle = 1
ElseIf A_c1.circ(1) > A_c2.circ(1) Then
compare_two_two_circle = -1
Else
compare_two_two_circle = 0
End If
End If
End Function

Public Function compare_two_line_circle(A_lc1 As add_point_for_line_circle_type, _
        A_lc2 As add_point_for_line_circle_type) As Integer
If A_lc1.line_no < A_lc2.line_no Then
compare_two_line_circle = 1
ElseIf A_lc1.line_no > A_lc2.line_no Then
compare_two_line_circle = -1
Else
If A_lc1.circ < A_lc2.circ Then
compare_two_line_circle = 1
ElseIf A_lc1.circ > A_lc2.circ Then
compare_two_line_circle = -1
Else
compare_two_line_circle = 0
End If
End If
End Function
Public Function search_for_line_circle(t_l_c As add_point_for_line_circle_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_line_circle
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_line_circle = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_line_circle(t_l_c, _
       add_aid_point_for_line_circle_(add_aid_point_for_line_circle_(n%).index))
If ty = 0 Then
  search_for_line_circle = True
   Exit Function
Else
 search_for_line_circle = judge_loop(n%, n1%, n2%, ty)
  If search_for_line_circle = True Then
   search_for_line_circle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function search_for_two_circle(t_c As add_point_for_two_circle_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_two_circle
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_two_circle = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_two_circle(t_c, _
       add_aid_point_for_two_circle_(add_aid_point_for_two_circle_(n%).index))
If ty = 0 Then
  search_for_two_circle = True
   Exit Function
Else
 search_for_two_circle = judge_loop(n%, n1%, n2%, ty)
  If search_for_two_circle = True Then
   search_for_two_circle = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_eline_for_aid(A_el1 As add_point_for_eline_type, _
        A_el2 As add_point_for_eline_type) As Integer
If A_el1.poi(0) < A_el2.poi(0) Then
 compare_two_eline_for_aid = 1
ElseIf A_el1.poi(0) > A_el2.poi(0) Then
 compare_two_eline_for_aid = -1
Else
 If A_el1.poi(1) < A_el2.poi(1) Then
  compare_two_eline_for_aid = 1
 ElseIf A_el1.poi(1) > A_el2.poi(1) Then
  compare_two_eline_for_aid = -1
 Else
  If A_el1.line_no < A_el2.line_no Then
   compare_two_eline_for_aid = 1
  ElseIf A_el1.line_no > A_el2.line_no Then
   compare_two_eline_for_aid = -1
  Else
   If A_el1.te < A_el2.te Then
    compare_two_eline_for_aid = 1
   ElseIf A_el1.te > A_el2.te Then
    compare_two_eline_for_aid = -1
   Else
    compare_two_eline_for_aid = 0
   End If
  End If
 End If
End If
End Function
Public Function search_for_aid_mid_point(A_mp As add_point_for_mid_point_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
If last_conditions.last_cond(1).new_midpoint_no > 0 Then
 Exit Function
End If
n1% = 1
n2% = last_add_aid_point_for_mid_point
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_aid_mid_point = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_mid_point_for_aid(A_mp, _
       add_aid_point_for_mid_point_(add_aid_point_for_mid_point_(n%).index))
If ty = 0 Then
  search_for_aid_mid_point = True
   Exit Function
Else
 search_for_aid_mid_point = judge_loop(n%, n1%, n2%, ty)
  If search_for_aid_mid_point = True Then
   search_for_aid_mid_point = False
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_mid_point_for_aid(A_mp1 As add_point_for_mid_point_type, _
        A_mp2 As add_point_for_mid_point_type) As Integer
If A_mp1.poi(0) < A_mp2.poi(0) Then
 compare_two_mid_point_for_aid = 1
ElseIf A_mp1.poi(0) > A_mp2.poi(0) Then
 compare_two_mid_point_for_aid = -1
Else
 If A_mp1.poi(1) < A_mp2.poi(1) Then
  compare_two_mid_point_for_aid = 1
 ElseIf A_mp1.poi(1) > A_mp2.poi(1) Then
  compare_two_mid_point_for_aid = -1
 Else
  If A_mp1.poi(2) < A_mp2.poi(2) Then
   compare_two_mid_point_for_aid = 1
  ElseIf A_mp1.poi(2) > A_mp2.poi(2) Then
   compare_two_mid_point_for_aid = -1
  Else
   compare_two_mid_point_for_aid = 0
  End If
 End If
End If
End Function
'Public Function search_for_aid_eline(A_el As add_point_for_eline_type, n%) As Boolean
'Dim n1%, n2%
'Dim ty As Integer
'n1% = 1
'n2% = last_add_aid_point_for_eline
'If n2% = 0 Or n1% > n2% Then
' n% = 0
'  search_for_aid_eline = False
'   Exit Function
'End If
'Do
'n% = n1% + (n2% - n1%) \ 2
'If n% = 0 Then
' If n2% = 0 Then
'  Exit Function
' Else
'  n% = 1
' End If
'End If
'ty = compare_two_eline_for_aid(A_el, _
'       add_aid_point_for_eline_(add_aid_point_for_eline_(n%).index))
'If ty = 0 Then
'  search_for_aid_eline = True
'   Exit Function
'Else
' search_for_aid_eline = judge_loop(n%, n1%, n2%, ty)
'  If search_for_aid_eline = True Then
'   search_for_aid_eline = False
'    Exit Function
'  End If
'End If
'Loop
'End Function
Public Function compare_two_verti_mid_line(v_m_l1 As verti_mid_line_data0_type, _
     v_m_l2 As verti_mid_line_data0_type) As Integer
If v_m_l1.poi(0) < v_m_l2.poi(0) Then
compare_two_verti_mid_line = 1
ElseIf v_m_l1.poi(0) > v_m_l2.poi(0) Then
compare_two_verti_mid_line = -1
Else
If v_m_l1.poi(2) < v_m_l2.poi(2) Then
compare_two_verti_mid_line = 1
ElseIf v_m_l1.poi(2) > v_m_l2.poi(2) Then
compare_two_verti_mid_line = -1
Else
If v_m_l1.line_no(0) < v_m_l2.line_no(0) Then
compare_two_verti_mid_line = 1
ElseIf v_m_l1.line_no(0) > v_m_l2.line_no(0) Then
compare_two_verti_mid_line = -1
Else
compare_two_verti_mid_line = 0
End If
End If
End If
End Function
Public Function search_for_verti_mid_line(v_m_l As verti_mid_line_data0_type, _
             n%, ByVal k As Byte, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
If k = 0 Then
n1% = 0
Else
n1% = 1 + last_conditions.last_cond(0).verti_mid_line_no
k = k - 1
End If
n2% = last_conditions.last_cond(1).verti_mid_line_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_verti_mid_line = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_verti_mid_line(v_m_l, _
        verti_mid_line(verti_mid_line(n%).data(0).record.data1.index.i(0)).data(0).data0)
If ty = 0 Then
 If ty_ = 0 Then
 n% = verti_mid_line(n%).data(0).record.data1.index.i(0)
 Else
 n% = n% - 1
 End If
  search_for_verti_mid_line = True
   Exit Function
Else
 search_for_verti_mid_line = judge_loop(n%, n1%, n2%, ty)
  If search_for_verti_mid_line = True Then
   search_for_verti_mid_line = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_add_aid_point_for_mid_point( _
      mp1 As add_point_for_mid_point_type, _
       mp2 As add_point_for_mid_point_type) As Integer
If mp1.poi(0) < mp2.poi(0) Then
 compare_two_add_aid_point_for_mid_point = 1
ElseIf mp1.poi(0) > mp2.poi(0) Then
 compare_two_add_aid_point_for_mid_point = -1
Else
 If mp1.poi(1) < mp2.poi(1) Then
  compare_two_add_aid_point_for_mid_point = 1
 ElseIf mp1.poi(1) > mp2.poi(1) Then
  compare_two_add_aid_point_for_mid_point = -1
 Else
   If mp1.poi(2) < mp2.poi(2) Then
    compare_two_add_aid_point_for_mid_point = 1
   ElseIf mp1.poi(2) > mp2.poi(2) Then
    compare_two_add_aid_point_for_mid_point = -1
   Else
    compare_two_add_aid_point_for_mid_point = 0
   End If
 End If
End If
End Function

Public Function search_for_add_aid_point_for_mid_point(mp As add_point_for_mid_point_type, _
             n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_mid_point
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_add_aid_point_for_mid_point = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_add_aid_point_for_mid_point(mp, _
        add_aid_point_for_mid_point_(add_aid_point_for_mid_point_(n%).index))
If ty = 0 Then
 n% = add_aid_point_for_mid_point_(n%).index
  search_for_add_aid_point_for_mid_point = True
   Exit Function
Else
 search_for_add_aid_point_for_mid_point = judge_loop(n%, n1%, n2%, ty)
  If search_for_add_aid_point_for_mid_point = True Then
   search_for_add_aid_point_for_mid_point = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_add_aid_point_for_two_line( _
      wl1 As add_point_for_two_line_type, _
       wl2 As add_point_for_two_line_type) As Integer
If wl1.line_no(0) < wl2.line_no(0) Then
 compare_two_add_aid_point_for_two_line = 1
ElseIf wl1.line_no(0) > wl2.line_no(0) Then
 compare_two_add_aid_point_for_two_line = -1
Else
 If wl1.line_no(1) < wl2.line_no(1) Then
  compare_two_add_aid_point_for_two_line = 1
 ElseIf wl1.line_no(1) > wl2.line_no(1) Then
  compare_two_add_aid_point_for_two_line = -1
 Else
  compare_two_add_aid_point_for_two_line = 0
 End If
End If
End Function

Public Function search_for_add_aid_point_for_two_line( _
        wl As add_point_for_two_line_type, _
             n%, ByVal k As Byte, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_two_line
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_add_aid_point_for_two_line = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_add_aid_point_for_two_line(wl, _
        add_aid_point_for_two_line_(add_aid_point_for_two_line_(n%).index))
If ty = 0 Then
 If ty_ = 0 Then
 n% = add_aid_point_for_two_line_(n%).index
 Else
 n% = n% - 1
 End If
  search_for_add_aid_point_for_two_line = True
   Exit Function
Else
 search_for_add_aid_point_for_two_line = judge_loop(n%, n1%, n2%, ty)
  If search_for_add_aid_point_for_two_line = True Then
   search_for_add_aid_point_for_two_line = False
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_add_aid_point_for_two_circle( _
      wc1 As add_point_for_two_circle_type, _
       wc2 As add_point_for_two_circle_type) As Integer
If wc1.circ(0) < wc2.circ(0) Then
 compare_two_add_aid_point_for_two_circle = 1
ElseIf wc1.circ(0) > wc2.circ(0) Then
 compare_two_add_aid_point_for_two_circle = -1
Else
 If wc1.circ(1) < wc2.circ(1) Then
  compare_two_add_aid_point_for_two_circle = 1
 ElseIf wc1.circ(1) > wc2.circ(1) Then
  compare_two_add_aid_point_for_two_circle = -1
 Else
  compare_two_add_aid_point_for_two_circle = 0
 End If
End If
End Function

Public Function search_for_add_aid_point_for_two_circle( _
        wc As add_point_for_two_circle_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_two_circle
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_add_aid_point_for_two_circle = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_add_aid_point_for_two_circle(wc, _
        add_aid_point_for_two_circle_(add_aid_point_for_two_circle_(n%).index))
If ty = 0 Then
 n% = add_aid_point_for_two_circle_(n%).index
  search_for_add_aid_point_for_two_circle = True
   Exit Function
Else
 search_for_add_aid_point_for_two_circle = judge_loop(n%, n1%, n2%, ty)
  If search_for_add_aid_point_for_two_circle = True Then
   search_for_add_aid_point_for_two_circle = False
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_add_aid_point_for_line_circle( _
      lc1 As add_point_for_line_circle_type, _
       lc2 As add_point_for_line_circle_type) As Integer
If lc1.line_no < lc2.line_no Then
 compare_two_add_aid_point_for_line_circle = 1
ElseIf lc1.line_no > lc2.line_no Then
 compare_two_add_aid_point_for_line_circle = -1
Else
 If lc1.circ < lc2.circ Then
  compare_two_add_aid_point_for_line_circle = 1
 ElseIf lc1.circ > lc2.circ Then
  compare_two_add_aid_point_for_line_circle = -1
 Else
  If lc1.poi = 0 And lc2.poi = 0 Then
   compare_two_add_aid_point_for_line_circle = 0
  ElseIf lc1.poi = 0 Then
   compare_two_add_aid_point_for_line_circle = 1
  ElseIf lc2.poi = 0 Then
   compare_two_add_aid_point_for_line_circle = -1
  Else
   If lc1.poi < lc2.poi Then
    compare_two_add_aid_point_for_line_circle = 1
   ElseIf lc1.poi > lc2.poi Then
    compare_two_add_aid_point_for_line_circle = -1
   Else
    If lc1.paral_or_verti < lc2.paral_or_verti Then
     compare_two_add_aid_point_for_line_circle = 1
    ElseIf lc1.paral_or_verti > lc2.paral_or_verti Then
     compare_two_add_aid_point_for_line_circle = -1
    Else
     compare_two_add_aid_point_for_line_circle = 0
    End If
   End If
  End If
 End If
End If
End Function

Public Function search_for_add_aid_point_for_line_circle( _
        wc As add_point_for_line_circle_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_line_circle
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_add_aid_point_for_line_circle = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_add_aid_point_for_line_circle(wc, _
        add_aid_point_for_line_circle_(add_aid_point_for_line_circle_(n%).index))
If ty = 0 Then
 n% = add_aid_point_for_line_circle_(n%).index
  search_for_add_aid_point_for_line_circle = True
   Exit Function
Else
 search_for_add_aid_point_for_line_circle = judge_loop(n%, n1%, n2%, ty)
  If search_for_add_aid_point_for_line_circle = True Then
   search_for_add_aid_point_for_line_circle = False
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_add_aid_point_for_eline( _
      el1 As add_point_for_eline_type, _
       el2 As add_point_for_eline_type) As Integer
If el1.poi(0) < el2.poi(0) Then
 compare_two_add_aid_point_for_eline = 1
ElseIf el1.poi(0) > el2.poi(0) Then
 compare_two_add_aid_point_for_eline = -1
Else
 If el1.poi(1) < el2.poi(1) Then
  compare_two_add_aid_point_for_eline = 1
 ElseIf el1.poi(1) > el2.poi(1) Then
  compare_two_add_aid_point_for_eline = -1
 Else
  If el1.poi(2) < el2.poi(2) Then
   compare_two_add_aid_point_for_eline = 1
  ElseIf el1.poi(2) > el2.poi(2) Then
   compare_two_add_aid_point_for_eline = -1
  Else
   If el1.line_no < el2.line_no Then
    compare_two_add_aid_point_for_eline = 1
   ElseIf el1.line_no > el2.line_no Then
    compare_two_add_aid_point_for_eline = -1
   Else
    If el1.te < el2.te Then
     compare_two_add_aid_point_for_eline = 1
    ElseIf el1.te > el2.te Then
     compare_two_add_aid_point_for_eline = -1
    Else
     compare_two_add_aid_point_for_eline = 0
   End If
   End If
  End If
 End If
End If
End Function

Public Function search_for_add_aid_point_for_eline( _
        el As add_point_for_eline_type, n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_add_aid_point_for_eline
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_add_aid_point_for_eline = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_add_aid_point_for_eline(el, _
        add_aid_point_for_eline_(add_aid_point_for_eline_(n%).index))
If ty = 0 Then
 n% = add_aid_point_for_eline_(n%).index
  search_for_add_aid_point_for_eline = True
   Exit Function
Else
 search_for_add_aid_point_for_eline = judge_loop(n%, n1%, n2%, ty)
  If search_for_add_aid_point_for_eline = True Then
   search_for_add_aid_point_for_eline = False
    Exit Function
  End If
End If
Loop
End Function


Public Function compare_two_tri_function(tri_f1 As tri_function_data_type, _
     tri_f2 As tri_function_data_type) As Integer
 If tri_f1.A < tri_f2.A Then
  compare_two_tri_function = 1
 ElseIf tri_f1.A > tri_f2.A Then
  compare_two_tri_function = -1
 Else
   compare_two_tri_function = 0
 End If
End Function

Public Function search_for_tri_function(start%, tri_f As tri_function_data_type, _
              n%) As Boolean
Dim n1%, n2%
Dim ty As Integer
n1% = start%
n2% = last_conditions.last_cond(1).tri_function_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_tri_function = False
   Exit Function
End If
While tri_function(n1%).data(0).record.data1.index.i(0) = 0 And n1% < n2%
 n1% = n1% + 1
Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_tri_function(tri_f, tri_function(tri_function(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
 n% = tri_function(n%).data(0).record.data1.index.i(0)
  search_for_tri_function = True
   Exit Function
Else
 search_for_tri_function = judge_loop(n%, n1%, n2%, ty)
  If search_for_tri_function = True Then
   search_for_tri_function = False
    Exit Function
  End If
End If
Loop

End Function

Public Function compare_two_total_angle(t_A1 As total_angle_data_type, _
                                      t_A2 As total_angle_data_type) As Integer
If t_A1.line_no(0) < t_A2.line_no(0) Then
 compare_two_total_angle = 1
ElseIf t_A1.line_no(0) > t_A2.line_no(0) Then
 compare_two_total_angle = -1
Else
 If t_A1.line_no(1) < t_A2.line_no(1) Then
  compare_two_total_angle = 1
 ElseIf t_A1.line_no(1) > t_A2.line_no(1) Then
  compare_two_total_angle = -1
 Else
  compare_two_total_angle = 0
 End If
End If
End Function

Public Function compare_two_angle_(A1%, A2%)
Dim angle_data0  As angle_data_type
If A1% = -1 Then
 angle_data0.line_no(0) = -1
 angle_data0.line_no(1) = -1
ElseIf A1% = 30000 Then
 angle_data0.line_no(0) = 30000
 angle_data0.line_no(1) = 30000
Else
 angle_data0 = angle(A1%).data(0)
End If
  If angle_data0.line_no(0) = angle(A2%).data(0).line_no(0) Or angle_data0.line_no(0) = angle(A2%).data(0).line_no(1) Or _
       angle_data0.line_no(1) = angle(A2%).data(0).line_no(0) Or angle_data0.line_no(1) = angle(A2%).data(0).line_no(1) Then
   compare_two_angle_ = 0
  ElseIf angle_data0.line_no(0) < angle(A2%).data(0).line_no(0) Then
   compare_two_angle_ = 1
  ElseIf angle_data0.line_no(0) > angle(A2%).data(0).line_no(0) Then
   compare_two_angle_ = -1
  End If
End Function
Public Function search_for_total_angle(t_A As total_angle_data_type, n%) As Boolean ', k%, ty_ As Byte) As Boolean
Dim n1%, n2%
Dim ty As Integer
 n1% = last_conditions.last_cond(0).total_angle_no
 n2% = last_conditions.last_cond(1).total_angle_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_total_angle = False
   Exit Function
End If
'While T_angle(n1%).data(0).index(k%) = 0 And n1% < n2%
' n1% = n1% + 1
'Wend
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_total_angle(t_A, T_angle(T_angle(n%).data(0).index).data(0))
If ty = 0 Then
' If ty_ = 0 Then
  n% = T_angle(n%).data(0).index
' Else
 ' n% = n% - 1
' End If
  search_for_total_angle = True
    Exit Function
Else 'If ty <>1 Then
  If judge_loop(n%, n1%, n2%, ty) Then
      search_for_total_angle = False
       'If ty = 1 Then  '输出新数据插入点，在n%后
       '  n% = n% - 1
       'End If
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_four_sides_fig(F_s_fig1 As four_sides_fig_data_type, _
                 F_s_fig2 As four_sides_fig_data_type, k%)
Dim tn(3) As Integer
tn(0) = k%
tn(1) = (tn(0) + 1) Mod 4
tn(2) = (tn(0) + 2) Mod 4
tn(3) = (tn(0) + 3) Mod 4
If F_s_fig1.poi(tn(0)) < F_s_fig2.poi(tn(0)) Then
 compare_two_four_sides_fig = 1
ElseIf F_s_fig1.poi(tn(0)) > F_s_fig2.poi(tn(0)) Then
 compare_two_four_sides_fig = -1
Else
If F_s_fig1.poi(tn(1)) < F_s_fig2.poi(tn(1)) Then
 compare_two_four_sides_fig = 1
ElseIf F_s_fig1.poi(tn(1)) > F_s_fig2.poi(tn(1)) Then
 compare_two_four_sides_fig = -1
Else
If F_s_fig1.poi(tn(2)) < F_s_fig2.poi(tn(2)) Then
 compare_two_four_sides_fig = 1
ElseIf F_s_fig1.poi(tn(2)) > F_s_fig2.poi(tn(2)) Then
 compare_two_four_sides_fig = -1
Else
If F_s_fig1.poi(tn(3)) < F_s_fig2.poi(tn(3)) Then
 compare_two_four_sides_fig = 1
ElseIf F_s_fig1.poi(tn(3)) > F_s_fig2.poi(tn(3)) Then
 compare_two_four_sides_fig = -1
Else
 compare_two_four_sides_fig = 0
End If
End If
End If
End If
End Function

Public Function compare_two_integer(int1%, int2%) As Integer
If int1% < int2% Then
 compare_two_integer = 1
ElseIf int1% > int2% Then
 compare_two_integer = -1
Else
 compare_two_integer = 0
End If
End Function

Public Function compare_two_equation(E1 As Equation_data0_type, _
      E2 As Equation_data0_type) As Integer
If E1.para_xx < E2.para_xx Then
 compare_two_equation = 1
ElseIf E1.para_xx > E2.para_xx Then
 compare_two_equation = -1
Else
 If E1.para_yy < E2.para_yy Then
  compare_two_equation = 1
 ElseIf E1.para_yy > E2.para_yy Then
  compare_two_equation = -1
 Else
  If E1.para_xy < E2.para_xy Then
   compare_two_equation = 1
  ElseIf E1.para_xy > E2.para_xy Then
   compare_two_equation = -1
  Else
     If E1.para_x < E2.para_x Then
      compare_two_equation = 1
     ElseIf E1.para_x > E2.para_x Then
      compare_two_equation = -1
     Else
        If E1.para_y < E2.para_y Then
          compare_two_equation = 1
        ElseIf E1.para_y > E2.para_y Then
         compare_two_equation = -1
        Else
           If E1.para_c < E2.para_c Then
            compare_two_equation = 1
           ElseIf E1.para_c > E2.para_c Then
            compare_two_equation = -1
           Else
            compare_two_equation = 0
           End If
        End If
     End If
  End If
 End If
End If
End Function
Public Function search_for_equation(e As Equation_data0_type, _
                n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
n1 = 1
n2% = last_conditions.last_cond(1).equation_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_equation = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_equation(e, equation(equation(n%).data(0).record.data1.index.i(0)).data(0))
If ty = 0 Then
If ty_ = 0 Then
 n% = equation(n%).data(0).record.data1.index.i(0)
Else
 n% = n% - 1
End If
  search_for_equation = True
   Exit Function
Else
 search_for_equation = judge_loop(n%, n1%, n2%, ty)
  If search_for_equation = True Then
   search_for_equation = False
    Exit Function
  End If
End If
Loop
End Function
Public Function compare_two_value_string(v_s1 As value_string0_type, v_s2 As value_string0_type, k%) As Integer
Dim i%, l%
If k% = 0 Then
 If Len(v_s1.value) < Len(v_s2.value) Then
  compare_two_value_string = 1
 ElseIf Len(v_s1.value) > Len(v_s2.value) Then
  compare_two_value_string = -1
 Else
  If v_s1.value < v_s2.value Then
   compare_two_value_string = 1
  ElseIf v_s1.value > v_s2.value Then
   compare_two_value_string = -1
  Else
   compare_two_value_string = 0
  End If
 End If
Else
 
 If v_s1.factor.data(0).last_factor < v_s2.factor.data(0).last_factor Then
  compare_two_value_string = 1
 ElseIf v_s1.factor.data(0).last_factor < v_s2.factor.data(0).last_factor Then
  compare_two_value_string = -1
 Else
   If v_s1.factor.data(1).last_factor < v_s2.factor.data(1).last_factor Then
    compare_two_value_string = 1
   ElseIf v_s1.factor.data(1).last_factor < v_s2.factor.data(1).last_factor Then
    compare_two_value_string = -1
   Else
    l% = compare_two_factor(v_s1.factor.data(0), v_s2.factor.data(0))
     If l% = 0 Then
     compare_two_value_string = compare_two_factor(v_s1.factor.data(1), v_s2.factor.data(1))
     Else
     compare_two_value_string = l%
     End If
   End If
 End If
 End If
End Function

Public Function search_for_value_string(v_s As value_string0_type, k%, n%, ty_ As Byte) As Integer
Dim n1%, n2%
Dim ty As Integer
n1% = 1 + last_conditions.last_cond(0).value_string_no
n2% = last_conditions.last_cond(1).value_string_no
If n2% = 0 Then
 n% = 0
  search_for_value_string = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_value_string(v_s, _
                    Dvalue_string(Dvalue_string(n%).data(0).index(k%)).data(0), _
                     k%)
If ty = 0 Then
 If ty_ = 0 Then
 n% = Dvalue_string(n%).data(0).index(k%)
 Else
 n% = n% - 1
 End If
   search_for_value_string = True
    Exit Function
Else
  If judge_loop(n%, n1%, n2%, ty) Then
  'If search_for_two_point_line > 0 Then
   search_for_value_string = False
    Exit Function
  End If
End If
Loop
End Function

Public Function compare_two_factor(fa1 As factor0_type, fa2 As factor0_type) As Integer
Dim i%
If fa1.last_factor < fa2.last_factor Then
  compare_two_factor = 1
 ElseIf fa1.last_factor < fa2.last_factor Then
  compare_two_factor = -1
 Else
  If fa1.para < fa2.para Then
   compare_two_factor = 1
  ElseIf fa1.para > fa2.para Then
   compare_two_factor = -1
  Else
   For i% = 1 To fa2.last_factor
    If fa1.para = fa2.para Then
     If fa1.order(i%) = fa2.order(i%) Then
      If i% = fa2.last_factor Then
      compare_two_factor = 0
       Exit Function
      End If
     ElseIf fa1.order(i%) < fa2.order(i%) Then
      compare_two_factor = 1
       Exit Function
     Else
      compare_two_factor = -1
       Exit Function
     End If
    ElseIf fa1.para < fa2.para Then
     compare_two_factor = 1
      Exit Function
    Else
    compare_two_factor = -1
      Exit Function
    End If
   Next i%
  End If
 End If

End Function

Public Function mid_no(n1%, n2%, dr%) As Integer
If n2% - n1% > 1 Then
  mid_no = (n1% + n2%) \ 2
Else
  If dr% = -1 Then
   mid_no = n1%
  ElseIf dr% = 1 Then
   mid_no = n2%
  End If
End If
End Function
Public Function compare_two_area_element_for_new(area_ele1 As condition_type, _
                    area_ele2 As condition_type) As Integer
Dim tp(1) As Integer
If area_ele1.ty = triangle_ Then
 tp(0) = triangle(area_ele1.no).data(0).poi(3)
ElseIf area_ele1.ty = polygon_ Then
 tp(0) = Dpolygon4(area_ele1.no).data(0).poi(4)
Else
 Exit Function
End If
If area_ele2.ty = triangle_ Then
 tp(1) = triangle(area_ele2.no).data(0).poi(3)
ElseIf area_ele1.ty = polygon_ Then
 tp(1) = Dpolygon4(area_ele2.no).data(0).poi(4)
Else
 Exit Function
End If
If tp(0) > tp(1) Then
 compare_two_area_element_for_new = 1
ElseIf tp(0) < tp(1) Then
 compare_two_area_element_for_new = -1
Else
 compare_two_area_element_for_new = 0
End If
 

End Function

Public Function compare_two_condition_type(cond1 As condition_type, _
           cond2 As condition_type) As Integer
 If cond1.ty < cond2.ty Then
  compare_two_condition_type = 1
 ElseIf cond1.ty > cond2.ty Then
  compare_two_condition_type = -1
 Else
  If cond1.no < cond2.no Then
   compare_two_condition_type = 1
  ElseIf cond1.no > cond2.no Then
   compare_two_condition_type = -1
  Else
    compare_two_condition_type = 0
  End If
 End If
End Function
Public Function searh_for_two_area_of_element(t_A_ele As two_area_element_value_data_type, _
           ByVal k As Byte, n%, ty_ As Byte) As Integer
Dim n1%, n2%
Dim ty As Integer
n1% = 1
n2% = last_conditions.last_cond(1).two_area_of_element_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  searh_for_two_area_of_element = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_two_area_of_element(t_A_ele, _
       two_area_of_element_value(two_area_of_element_value(n%).data(0).record.data1.index.i(k)).data(0), k)
If ty = 0 Then
 If ty_ = 0 Then
 n% = two_area_of_element_value(n%).data(0).record.data1.index.i(k)
 Else
 n% = n% - 1
 End If
  searh_for_two_area_of_element = True
   Exit Function
Else
 searh_for_two_area_of_element = judge_loop(n%, n1%, n2%, ty)
  If searh_for_two_area_of_element = True Then
   searh_for_two_area_of_element = False
    Exit Function
  End If
End If
Loop
End Function

Public Function search_for_V_line_value(l_value As V_line_value_data0_type, _
            ByVal k As Byte, n%, ty_ As Byte) As Boolean
Dim n1%, n2%
'Dim k1 As Byte
Dim ty As Integer
If k = 0 Then
n1% = 1
Else
n1% = 1 + last_conditions.last_cond(0).v_line_value_no
'k1 = k - 1
End If
n2% = last_conditions.last_cond(1).v_line_value_no
If n2% = 0 Or n1% > n2% Then
 n% = 0
  search_for_V_line_value = False
   Exit Function
End If
Do
n% = n1% + (n2% - n1%) \ 2
If n% = 0 Then
 If n2% = 0 Then
  Exit Function
 Else
  n% = 1
 End If
End If
ty = compare_two_V_line_value(l_value, _
       V_line_value(V_line_value(n%).data(0).record.data1.index.i(k)).data(0), k)
If ty = 0 Then
If ty_ = 0 Then
 n% = V_line_value(n%).data(0).record.data1.index.i(k)
Else
 n% = n% - 1
End If
  search_for_V_line_value = True
   Exit Function
Else
 search_for_V_line_value = judge_loop(n%, n1%, n2%, ty)
  If search_for_V_line_value = True Then
   search_for_V_line_value = False
    Exit Function
  End If
End If
Loop
End Function
Public Sub set_conclusion_point(ByVal conc_no%, ByVal p%)
Dim i%, j%
If p% > 0 Then
If conclusion_point(conc_no%).poi(0) > 0 Then
 For i% = 1 To conclusion_point(conc_no%).poi(0)
  If conclusion_point(conc_no%).poi(i%) > p% Then
     conclusion_point(conc_no%).poi(0) = conclusion_point(conc_no%).poi(0) + 1
     For j% = conclusion_point(conc_no%).poi(0) To i% + 1 Step -1
         conclusion_point(conc_no%).poi(j%) = conclusion_point(conc_no%).poi(j% - 1)
     Next j%
     conclusion_point(conc_no%).poi(i%) = p%
     Exit Sub
  ElseIf conclusion_point(conc_no%).poi(i%) = p% Then
     Exit Sub
  End If
 Next i%
      conclusion_point(conc_no%).poi(0) = conclusion_point(conc_no%).poi(0) + 1
      conclusion_point(conc_no%).poi(conclusion_point(conc_no%).poi(0)) = _
            p%
Else
     conclusion_point(conc_no%).poi(0) = conclusion_point(conc_no%).poi(0) + 1
     conclusion_point(conc_no%).poi(1) = p%
End If
End If
End Sub
Public Sub set_conclusion_point_for_circle(ByVal conc_no%, ByVal c%)
 Call set_conclusion_point(conc_no%, m_Circ(c%).data(0).parent.element(0).no)
 Call set_conclusion_point(conc_no%, m_Circ(c%).data(0).parent.element(1).no)
 Call set_conclusion_point(conc_no%, m_Circ(c%).data(0).parent.element(2).no)
End Sub
Public Sub set_conclusion_point_for_angle(ByVal conc_no%, ByVal A%)
 Call set_conclusion_point(conc_no%, angle(A%).data(0).poi(0))
 Call set_conclusion_point(conc_no%, angle(A%).data(0).poi(1))
 Call set_conclusion_point(conc_no%, angle(A%).data(0).poi(2))
End Sub
Public Sub set_conclusion_point_for_line(ByVal conc_no%, ByVal l%)
 If m_lin(l%).data(0).parent.element(1).ty = point_ And _
      m_lin(l%).data(0).parent.element(2).ty = point_ Then
 Call set_conclusion_point(conc_no%, m_lin(l%).data(0).parent.element(1).no)
 Call set_conclusion_point(conc_no%, m_lin(l%).data(0).parent.element(1).ty)
 Else
 Call set_conclusion_point(conc_no%, m_lin(l%).data(0).data0.poi(0))
 Call set_conclusion_point(conc_no%, m_lin(l%).data(0).data0.poi(1))
 End If
End Sub
Public Sub set_conclusion_point_for_triangle(ByVal conc_no%, ByVal tA%)
 Call set_conclusion_point(conc_no%, triangle(tA%).data(0).poi(0))
 Call set_conclusion_point(conc_no%, triangle(tA%).data(0).poi(1))
 Call set_conclusion_point(conc_no%, triangle(tA%).data(0).poi(2))
End Sub
Public Sub set_conclusion_point_for_area_element(ByVal conc_n%, ele As condition_type)
  If ele.ty = triangle_ Then
   Call set_conclusion_point(conc_n%, triangle(ele.no).data(0).poi(0))
   Call set_conclusion_point(conc_n%, triangle(ele.no).data(0).poi(1))
   Call set_conclusion_point(conc_n%, triangle(ele.no).data(0).poi(2))
  Else
   Call set_conclusion_point(conc_n%, Dpolygon4(ele.no).data(0).poi(0))
   Call set_conclusion_point(conc_n%, Dpolygon4(ele.no).data(0).poi(1))
   Call set_conclusion_point(conc_n%, Dpolygon4(ele.no).data(0).poi(2))
   Call set_conclusion_point(conc_n%, Dpolygon4(ele.no).data(0).poi(3))
  End If
End Sub
Public Sub set_conclusion_point_for_item(ByVal conc_n%, i%)
Call set_conclusion_point_for_element_of_item(conc_n%, item0(i%).data(0).poi(0), item0(i%).data(0).poi(1))
Call set_conclusion_point_for_element_of_item(conc_n%, item0(i%).data(0).poi(2), item0(i%).data(0).poi(3))
End Sub
Public Sub set_conclusion_point_for_element_of_item(ByVal conc_n%, ByVal p1%, ByVal p2%)
If p1% > 0 And p2% > 0 Then
 Call set_conclusion_point(conc_n%, p1%)
 Call set_conclusion_point(conc_n%, p2%)
ElseIf p2% = -10 Then
 Call set_conclusion_point(conc_n%, Dtwo_point_line(p1%).data(0).v_poi(0))
 Call set_conclusion_point(conc_n%, Dtwo_point_line(p1%).data(0).v_poi(1))
ElseIf p2% = -1 Or p2% = -2 Or p2% = -3 Or p2% = -4 Or p2 = -6 Then
 Call set_conclusion_point_for_angle(conc_n%, p1%)
ElseIf p2% = -5 Then
 Call set_conclusion_point_for_item(conc_n%, p1%)
End If
End Sub

Public Function read_other_line(ByVal line_no%) As Integer
 If m_lin(line_no%).data(0).other_no <> line_no% Then
  read_other_line = read_other_line(m_lin(line_no%).data(0).other_no)
 Else
  read_other_line = line_no%
 End If
End Function

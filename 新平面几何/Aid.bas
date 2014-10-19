Attribute VB_Name = "aid"
Option Explicit
Global last_aid_point As Integer
Global add_point_name(7) As String * 1
Global last_aid_point_name As Integer
'*****************************
Type aid_line_data
  display_aid_line As display_line_data
  end_point(1) As POINTAPI '标准直线的端点
  start_point As Integer '辅助线的起点
  related_circle_no(1)  As Integer  '用于圆的切线
End Type
'*****************************
Type conditions_data_type
branch_data_no As Integer
init_v_line_no As Integer
pass_word_for_teacher As String * 5
set_branch As Boolean
new_midpoint_no As Integer
value_string_no As Integer
aid_point_data1_no As Integer
aid_point_data2_no As Integer
aid_point_data3_no As Integer
aid_point_no As Integer '15
four_sides_fig_no As Integer
angle3_value_no As Integer '32
angle_less_angle_no As Integer '28
angle_no As Integer '1
angle_relation_no As Integer
angle_value_90_no As Integer
angle_value_no  As Integer
area_of_circle_no As Integer '44
area_of_fan_no As Integer '46
area_of_element_no As Integer '43
sides_length_of_triangle_no As Integer '47
arc_value_no As Integer '25
arc_no As Integer
circle_no As Integer
con_line_no As Integer
'aid_circle_no As Integer
change_picture_type As Byte
change_picture_step As Integer
dangle_no As Integer
distance_of_paral_line_no As Integer
distance_of_point_line_no As Integer
dline1_no As Integer
dpoint_pair_no As Integer '3
epolygon_no As Integer '40
eangle_no As Integer
equal_3angle_no As Integer
equal_arc_no As Integer '26
equal_side_right_triangle_no As Integer
'equal_area_triangle_no As Integer '35
'equal_side_tixing_no As Integer
equal_side_triangle_no As Integer
equation_no As Integer
eline_no As Integer '8
general_string_no  As Integer '36
general_angle_string_no As Integer '37
function_of_angle_no As Integer
four_point_on_circle_no As Integer '9
three_point_on_circle_no As Integer '9
general_string_combine_no As Integer
item0_no As Integer
last_angle3_value_combine As Integer
last_general_string_combine As Integer
length_of_polygon_no As Integer
line_from_two_point_no As Integer
line2_less_line2_no As Integer '31
line3_value_no As Integer
line_less_line2_no As Integer '30
line_less_line_no As Integer '29
line_no As Integer '11
line_value_no As Integer '33
long_squre_no As Integer '42
squre_no As Integer '42
mid_point_line_no As Integer '7
mid_point_no As Integer '12
new_point_no As Integer
note_space_no As Integer
unkown_element_no As Integer
'point_pair_for_simlilar_no As Integer
parallelogram_no As Integer '14
paral_no As Integer '13
point_no As Integer
poly_no As Integer
polygon4_no As Integer
point_pair_for_similar_no As Integer
pre_add_condition_no As Integer
pseudo_dpoint_pair_no As Integer
pseudo_eline_no As Integer
pseudo_midpoint_no As Integer
pseudo_relation_no As Integer
pseudo_line3_value_no As Integer
relation_from_line_to_triangle_no As Integer '3
relation_from_triangle_to_line_no As Integer '3
relation_no As Integer '16
v_relation_no As Integer '16
relation_on_line_no As Integer
relation_string_no As Integer
ratio_of_two_arc_no As Integer '27
rhombus_no As Integer '41
right_angle_for_Pd_no As Integer
rtriangle_no As Integer
same_three_lines_no As Integer
sides_length_of_circle_no As Integer '48
similar_triangle_no As Integer '17
squ_sum_no As Integer '50
string_value_no As Integer
tangent_circle_no As Integer
tangent_line_no As Integer '34
'three_angle_value_sum_no  As Integer
three_angle_value0_no As Integer
total_angle_no As Integer '1
total_condition As Integer
trajectory_no As Integer
area_relation_no As Integer '4
tri_function_no As Integer
three_angle_value_sum_no As Integer
two_area_of_element_value_no As Integer
two_angle_value_sum_no As Integer
two_angle_value_180_no As Integer
two_angle_value_90_no As Integer
two_angle_value_no  As Integer
two_angle_value0_no As Integer
two_order_eqution_no As Integer
tixing_no As Integer '39
total_equal_triangle_no As Integer '18
pseudo_total_equal_triangle_no As Integer '18
pseudo_similar_triangle_no As Integer '18
triangle_no As Integer '19
three_point_on_line_no As Integer '20
two_point_conset_no As Integer
two_line_value_no As Integer '21
verti_mid_line_no As Integer '49
verti_no As Integer '24
last_view_point_no As Integer
last_dot_line As Integer
'**************************************************
v_line_value_no As Integer
V_two_line_time_no As Integer
End Type
Type last_conditions_type
  last_cond(2) As conditions_data_type
End Type
Type last_conditions_for_aid_type
last_cond(1) As conditions_data_type
'temp_cond(1) As temp_conditions_data_type
new_result_from_add As Boolean
End Type
Global last_conditions As last_conditions_type
Global t_condition As last_conditions_type
Global last_conditions_for_aid() As last_conditions_for_aid_type
Global last_conditions_for_aid_no As Byte '辅助点个数
'****************************************************
Public Sub display_aid_line(m_ob As Object, aid_line_dis_data As display_line_data)
   If aid_line_dis_data.display > 0 Then
 Call Drawline(m_ob, QBColor(fill_color), 0, aid_line_dis_data.end_point_coord(0), _
                    aid_line_dis_data.end_point_coord(1), 0, 1)
                         aid_line_dis_data.display = ( _
                          aid_line_dis_data.display + 1) Mod 3
   End If
End Sub
Public Function read_line_0(in_pointapi, p_coord1 As POINTAPI, p_coord2 As POINTAPI) As Byte '0不在线上,1在线上 ,2在辅助线上，3=1+2，4=在平行线上，5=在垂直线上

End Function
Public Function read_line_(m_object As Object, in_coord As POINTAPI, _
                                  m_aid_line0 As aid_line_data, out_coordinate As POINTAPI, _
                                   temp_end_point_coord0 As POINTAPI, temp_end_point_coord1 As POINTAPI, _
                                     Optional ty As Integer = 0) As Integer '0不在线上,1在线上 ,2在辅助线上，3=1+2，4=在平行线上，5=在垂直线上
Dim i%
'Dim in_coord As POINTAPI
Dim out_coord As POINTAPI
Dim temp_aid_line As display_line_data
'Dim temp_aid_line As aid_line_data
Dim is_ty As Integer
 If (ty <= condition Or ty >= aid_condition) And (ty <> paral_ And ty <> verti_) Then
    temp_end_point_coord0 = m_aid_line0.end_point(0)
    temp_end_point_coord1 = m_aid_line0.end_point(1)
  ElseIf distance_of_two_POINTAPI(temp_end_point_coord0, _
                                m_aid_line0.end_point(1)) > 6 Then
     If ty = paral_ Then
      temp_end_point_coord0 = m_poi(m_aid_line0.start_point).data(0).data0.coordinate
      temp_end_point_coord1 = add_POINTAPI(temp_end_point_coord0, _
                                minus_POINTAPI(m_aid_line0.end_point(1), _
                                    m_aid_line0.end_point(0)))
     ElseIf ty = verti_ Then
      temp_end_point_coord0 = m_poi(m_aid_line0.start_point).data(0).data0.coordinate
      temp_end_point_coord1 = add_POINTAPI(temp_end_point_coord0, _
                                verti_POINTAPI(minus_POINTAPI(m_aid_line0.end_point(1), _
                                    m_aid_line0.end_point(0))))
    End If
 Else
    read_line_ = 0
    Exit Function
 End If
  is_ty = is_point_on_line(in_coord, _
                             temp_end_point_coord0, _
                              temp_end_point_coord1, _
                                out_coord, _
                                  temp_aid_line.end_point_coord(0), _
                                   temp_aid_line.end_point_coord(1), _
                                   ty)
                        out_coordinate = out_coord
  If is_ty > 0 Then
     If ty <= 1 Then
      read_line_ = 1
       If is_ty > 1 Then
          read_line_ = 2
       End If
     ElseIf ty = 2 Then
       read_line_ = 4
     ElseIf ty = 3 Then
       read_line_ = 5
     End If
     
     If is_ty = point_out_segement Then '在直线外
                     read_line_ = is_ty
           If ty >= 1 Then
                  If m_aid_line0.display_aid_line.display = 2 Then
                   Call display_aid_line(m_object, m_aid_line0.display_aid_line)
                  End If
                  If m_aid_line0.display_aid_line.display = 0 Then
                   m_aid_line0.display_aid_line = temp_aid_line
                   m_aid_line0.display_aid_line.color = fill_color
                   m_aid_line0.display_aid_line.display = 1
                   Call display_aid_line(m_object, m_aid_line0.display_aid_line)
                  End If
           End If
      Else 'if is_ty = point_on_segement
           If ty >= 1 Then
                If m_aid_line0.display_aid_line.display = 2 Then
                   Call display_aid_line(m_object, m_aid_line0.display_aid_line)
                End If
           End If
      End If
   Else 'is_ty=0
           'If ty = 1 Then
                If m_aid_line0.display_aid_line.display = 2 Then
                   Call display_aid_line(m_object, m_aid_line0.display_aid_line)
                End If
           'End If
            read_line_ = 0
   End If
   ' m_aid_line0.display_aid_line.end_point_coord(0) = temp_end_point_coord(0)
  '  m_aid_line0.display_aid_line.end_point_coord(1) = temp_end_point_coord(1)
End Function

Public Function add_aid_point(is_no_initial As Byte, c_data0 As condition_data_type) As Byte
Dim t_n(7) As Integer
Dim t_l(1) As Integer
Dim tem_p(1) As Integer
Dim i%, j%, k%, tn%, n%, l%, tl%, m%, o%, p%, t_p%, tn_%
Dim r&, s&
Dim t!
Dim tv$
Dim t_y As Byte
Dim r1, r2 As Long
Dim A(3) As Integer
Dim tp(4) As Integer
Dim ty As Boolean
Dim lv As line_value_data0_type
Dim temp_record As record_type
Dim temp_record1 As total_record_type
Dim unkown_number As Byte
Dim triA(1) As temp_triangle_data_type
Dim cond_data(1) As condition_data_type
Dim t_triA(1) As temp_triangle_data_type
' 进入辅助点
'平行四边形,加对角线交点
'On Error GoTo add_aid_point_error
'add_aid_point = add_unknown_value
'If add_aid_point > 1 Then
 'Exit Function
'End If
run_type = 5
'add_aid_point = add_condition_for_no_reduce0
'    If add_aid_point > 1 Then
'        Exit Function
'     End If
add_aid_point = C_wait_for_aid_point.get_wait_for_aid_point
   If add_aid_point > 1 Then
    Exit Function
   End If
For i% = 1 To last_conditions.last_cond(1).pseudo_eline_no
 temp_record1.record_data = pseudo_eline(i%).data(0).record
  add_aid_point = set_equal_dline(pseudo_eline(i%).data(0).data0.poi(0), pseudo_eline(i%).data(0).data0.poi(1), _
   pseudo_eline(i%).data(0).data0.poi(2), pseudo_eline(i%).data(0).data0.poi(3), pseudo_eline(i%).data(0).data0.n(0), _
    pseudo_eline(i%).data(0).data0.n(1), pseudo_eline(i%).data(0).data0.n(2), pseudo_eline(i%).data(0).data0.n(3), _
     pseudo_eline(i%).data(0).data0.line_no(0), pseudo_eline(i%).data(0).data0.line_no(1), False, temp_record1, 0, 0, 0, 0, 0, False)
      If add_aid_point > 1 Then
       Exit Function
      End If
  add_aid_point = start_prove(0, 1, 1)
      If add_aid_point > 1 Then
       Exit Function
      End If
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_dpoint_pair_no
  temp_record1.record_data = pseudo_dpoint_pair(i%).data(0).record
  add_aid_point = set_dpoint_pair(pseudo_dpoint_pair(i%).data(0).data0.poi(0), pseudo_dpoint_pair(i%).data(0).data0.poi(1), _
   pseudo_dpoint_pair(i%).data(0).data0.poi(2), pseudo_dpoint_pair(i%).data(0).data0.poi(3), pseudo_dpoint_pair(i%).data(0).data0.poi(4), _
    pseudo_dpoint_pair(i%).data(0).data0.poi(5), pseudo_dpoint_pair(i%).data(0).data0.poi(6), pseudo_dpoint_pair(i%).data(0).data0.poi(7), _
     pseudo_dpoint_pair(i%).data(0).data0.n(0), pseudo_dpoint_pair(i%).data(0).data0.n(1), pseudo_dpoint_pair(i%).data(0).data0.n(2), _
      pseudo_dpoint_pair(i%).data(0).data0.n(3), pseudo_dpoint_pair(i%).data(0).data0.n(4), pseudo_dpoint_pair(i%).data(0).data0.n(5), _
       pseudo_dpoint_pair(i%).data(0).data0.n(6), pseudo_dpoint_pair(i%).data(0).data0.n(7), pseudo_dpoint_pair(i%).data(0).data0.line_no(0), _
        pseudo_dpoint_pair(i%).data(0).data0.line_no(1), pseudo_dpoint_pair(i%).data(0).data0.line_no(2), pseudo_dpoint_pair(i%).data(0).data0.line_no(3), _
         0, temp_record1, False, 0, 0, 0, 0, False)
      If add_aid_point > 1 Then
       Exit Function
      End If
  add_aid_point = start_prove(0, 1, 1)
      If add_aid_point > 1 Then
       Exit Function
      End If
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_midpoint_no
  add_aid_point = add_mid_point(pseudo_mid_point(i%).data(0).data0.poi(0), _
                                 pseudo_mid_point(i%).data(0).data0.poi(1), _
                                  pseudo_mid_point(i%).data(0).data0.poi(2), 0)
   If add_aid_point > 1 Then
        Exit Function
   End If
Next i%
'For i% = 1 To last_conditions.last_cond(1).pseudo_total_equal_triangle_no
'   add_aid_point = set_pseudo_total_equal_triangle(triA(0), triA(1), 0, 0, 0, 0, i%, cond_data(0), cond_data(1))
'   If add_aid_point > 1 Then
'       Exit Function
'    End If
'Next i%
For i% = 0 To last_conclusion - 1
  If conclusion_data(i%).no(0) = 0 Then
   add_aid_point = add_aid_point_from_conclusion(i%, unkown_number)
    If add_aid_point > 1 Then
     Exit Function
    End If
  End If
'End If
Next i%
run_type = 10
'*************
'add_aid_point = add_point_from_aid_point_data
For i% = 1 To last_conditions.last_cond(1).parallelogram_no
t_l(0) = line_number0(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(0), _
    Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(2), 0, 0)
t_l(1) = line_number0(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(1), _
     Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(3), 0, 0)
 add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 1, 0, 0, 0, c_data0)
  If add_aid_point > 1 Then
   Exit Function
  End If
Next i%
'矩形,加对角线交点
For i% = 1 To last_conditions.last_cond(1).long_squre_no
t_l(0) = line_number0(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(0), _
    Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(2), 0, 0)
t_l(1) = line_number0(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(1), _
     Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(3), 0, 0)
 add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 1, 0, 0, 0, c_data0)
  If add_aid_point > 1 Then
   Exit Function
  End If
Next i%
'菱形,加对角线交点
For i% = 1 To last_conditions.last_cond(1).rhombus_no
t_l(0) = line_number0(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(0), _
    Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(2), 0, 0)
t_l(1) = line_number0(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(1), _
     Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(3), 0, 0)
 add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 1, 0, 0, 0, c_data0)
  If add_aid_point > 1 Then
   Exit Function
  End If
Next i%
'正四边形,加对角线交点
For i% = 1 To last_conditions.last_cond(1).epolygon_no
If epolygon(i%).data(0).p.total_v = 4 Then
t_l(0) = line_number0(epolygon(i%).data(0).p.v(0), _
                      epolygon(i%).data(0).p.v(2), 0, 0)
t_l(1) = line_number0(epolygon(i%).data(0).p.v(1), _
                      epolygon(i%).data(0).p.v(3), 0, 0)
 add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 1, 0, 0, 0, c_data0)
  If add_aid_point > 1 Then
   Exit Function
  End If
End If
Next i%
 add_aid_point = add_point_for_tixing
  If add_aid_point > 1 Then
   Exit Function
  End If
'结论中, 有general_string
add_aid_point_mark_conclusiosn:
'************
For i% = 1 To last_conditions.last_cond(1).four_point_on_circle_no
 add_aid_point = add_point_from_paral_and_circle(0, i%)
 If add_aid_point > 1 Then
    Exit Function
 End If
 tp(1) = four_point_on_circle(i%).data(0).poi(1)
 tp(2) = four_point_on_circle(i%).data(0).poi(2)
 tp(3) = four_point_on_circle(i%).data(0).poi(3)
 A(0) = Abs(angle_number(tp(0), tp(1), tp(2), 0, 0))
 A(1) = Abs(angle_number(tp(1), tp(2), tp(3), 0, 0))
  If (angle(A(0)).data(0).value = "90" And angle(A(1)).data(0).value <> "") Or _
       (angle(A(1)).data(0).value = "90" And angle(A(0)).data(0).value <> "") Then
     If is_line_value(tp(1), tp(0), 0, 0, 0, "", 0, 0, 0, 0, 0, lv) = 1 Or _
       is_line_value(tp(2), tp(3), 0, 0, 0, "", 0, 0, 0, 0, 0, lv) = 1 Then
     t_l(0) = line_number0(tp(1), tp(2), 0, 0)
     t_l(1) = line_number0(tp(0), tp(3), 0, 0)
     If is_line_line_intersect(t_l(0), t_l(1), 0, 0, False) = 0 Then
      add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 0, 0, 0, 0, c_data0)
       If add_aid_point > 1 Then
        Exit Function
       End If
     End If
    ElseIf is_line_value(tp(0), tp(3), 0, 0, 0, "", 0, 0, 0, 0, 0, lv) = 1 Or _
       is_line_value(tp(1), tp(2), 0, 0, 0, "", 0, 0, 0, 0, 0, lv) = 1 Then
     t_l(0) = line_number0(tp(0), tp(1), 0, 0)
     t_l(1) = line_number0(tp(2), tp(3), 0, 0)
     If is_line_line_intersect(t_l(0), t_l(1), 0, 0, False) = 0 Then
      add_aid_point = add_interset_point_line_line(t_l(0), t_l(1), 0, 0, 0, 0, 0, c_data0)
       If add_aid_point > 1 Then
        Exit Function
       End If
     End If
     
     End If
  End If
Next i%
'************
'圆的对称点
For i% = 1 To last_conditions.last_cond(1).tangent_line_no - 1
 For j% = i% + 1 To last_conditions.last_cond(1).tangent_line_no
  add_aid_point = add_aid_point_for_com_tangent_line(i%, j%)
  If add_aid_point > 1 Then
   Exit Function
  End If
Next j%
Next i%
For i% = 1 To C_display_picture.m_circle.Count
add_aid_point = add_aid_point_for_circle(i%, 0)
If add_aid_point > 1 Then
 Exit Function
End If
 For j% = i% + 1 To C_display_picture.m_circle.Count
  If inter_point_circle_circle0(m_Circ(j%).data(0).data0, m_Circ(i%).data(0).data0, 0, 0) > 0 Then
   t_l(0) = line_number0(m_Circ(i%).data(0).data0.center, m_Circ(j%).data(0).data0.center, 0, 0)
    If t_l(0) > 0 Then
     add_aid_point = add_interset_point_line_circle( _
         m_Circ(i%).data(0).data0.center, m_Circ(j%).data(0).data0.center, t_l(0), j%, 0, cond_data(0), 0)
      If add_aid_point > 1 Then
       Exit Function
      End If
     add_aid_point = add_interset_point_line_circle( _
         m_Circ(j%).data(0).data0.center, m_Circ(i%).data(0).data0.center, t_l(0), i%, 0, cond_data(0), 0)
      If add_aid_point > 1 Then
       Exit Function
      End If
    End If
  End If
 Next j%
Next i%
'********************************************************************
p% = 1 + last_conditions.last_cond(0).mid_point_no
While p% <= last_conditions.last_cond(1).mid_point_no
' i% = Dmid_point(p%).record.data1.index(0)
add_aid_point = add_aid_point_for_midpoint(Dmid_point(p%).data(0).data0.poi(0), _
    Dmid_point(p%).data(0).data0.poi(1), Dmid_point(p%).data(0).data0.poi(2))
If add_aid_point > 1 Then
      Exit Function
End If
p% = p% + 1
Wend
For i% = 2 To last_conditions.last_cond(1).line_no '添加两线的交点
  If m_lin(i%).data(0).data0.visible = 2 Then
   For j% = 1 To i% - 1
    tp(0) = m_lin(i%).data(0).data0.poi(0)
     tp(1) = m_lin(i%).data(0).data0.poi(1)
      tp(2) = m_lin(j%).data(0).data0.poi(0)
       tp(3) = m_lin(j%).data(0).data0.poi(1)
 '    A(0) = angle_number(tp(0), tp(2), tp(1), 0, 0)
  '    A(1) = angle_number(tp(1), tp(3), tp(0), 0, 0)
   '    A(2) = angle_number(tp(2), tp(1), tp(3), 0, 0)
    '    A(3) = angle_number(tp(3), tp(0), tp(2), 0, 0)
    'If ((A(0) > 0 And A(1) > 0) Or _
           (A(0) < 0 And A(1) < 0)) And _
          ((A(2) > 0 And A(3) > 0) Or _
           (A(2) < 0 And A(3) < 0)) Then '
    If m_lin(j%).data(0).data0.visible = 2 And is_dparal(i%, j%, 0, -1000, _
          0, 0, 0, 0) = False Then
     'If is_line_line_intersect(Lin(i%), Lin(j%), 0, 0) = 0 Then
    '************
     add_aid_point = add_interset_point_line_line(i%, j%, 0, 0, 0, 0, 0, c_data0)
     If add_aid_point > 1 Then
      Exit Function
     End If
    '*****************8
    'End If
     End If
     'End If
     Next j%
    End If
    Next i%
'****************************************
'******************************************

'******************
'共线的结论
'************
' 角平分
'For i% = 1 To last_conditions.last_cond(1).eangle_no
' tn% = Deangle.av_no(i%).no
' If angle3_value(tn%).data(0).data0.angle_(3) > 0 And (angle3_value(tn%).data(0).data0.ty_(0) = 3 Or _
       angle3_value(tn%).data(0).data0.ty_(0) = 5) Then
' If angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0) = _
      angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1) Then
'     add_aid_point = add_point_for_eangle(angle3_value(tn%).data(0).data0.angle(1), _
                          angle3_value(tn%).data(0).data0.angle(0))
'          If add_aid_point > 1 Then
'           Exit Function
'          End If
' ElseIf angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1) = _
      angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0) Then
'      add_aid_point = add_point_for_eangle(angle3_value(tn%).data(0).data0.angle(0), _
                            angle3_value(tn%).data(0).data0.angle(1))
'          If add_aid_point > 1 Then
'           Exit Function
'          End If
'End If
' End If
'Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no
If Drelation(i%).data(0).data0.value = "2" Then
   add_aid_point = add_mid_point(Drelation(i%).data(0).data0.poi(0), 0, _
            Drelation(i%).data(0).data0.poi(1), 0)
    If add_aid_point > 1 Then
       Exit Function
    End If
ElseIf Drelation(i%).data(0).data0.value = "1/2" Then
   add_aid_point = add_mid_point(Drelation(i%).data(0).data0.poi(2), 0, _
            Drelation(i%).data(0).data0.poi(3), 0)
    If add_aid_point > 1 Then
       Exit Function
    End If
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).angle3_value_no
tn% = angle3_value(i%).data(0).record.data1.index.i(0)
If angle3_value(tn%).data(0).data0.type = angle_relation_ Then
If angle3_value(tn%).data(0).data0.para(0) = "2" And _
     angle3_value(tn%).data(0).data0.para(1) = "-1" Then
      If angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0) = _
           angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1) Then
          p% = is_line_line_intersect(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1), _
                angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0), 0, 0, False)
                  If p% > 0 Then
                   add_aid_point = add_aid_point_for_eangle_(p%, _
                       angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
                          angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).poi(1), _
                             angle3_value(tn%).data(0).data0.angle(1), 0, 0, 0, 0, 0, 0, 0)
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                   '在直线p2%,p3%上取一点使得∠p1%pp2%=∠p1%p2%p3%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
                   End If
      ElseIf angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1) = _
          angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0) Then
           p% = is_line_line_intersect(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0), _
                angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1), 0, 0, False)
                  If p% > 0 Then
                   add_aid_point = add_aid_point_for_eangle_(p%, _
                       angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
                          angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).poi(1), _
                             angle3_value(tn%).data(0).data0.angle(1), 0, 0, 0, 0, 0, 0, 0)
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                  End If
      End If
ElseIf angle3_value(tn%).data(0).data0.para(0) = "1" And _
    angle3_value(tn%).data(0).data0.para(1) = "-2" Then
      If angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0) = _
           angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1) Then
          p% = is_line_line_intersect(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1), _
                angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0), 0, 0, False)
                  If p% > 0 Then
                   add_aid_point = add_aid_point_for_eangle_(p%, _
                       angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).poi(1), _
                          angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
                             angle3_value(tn%).data(0).data0.angle(1), 0, 0, 0, 0, 0, 0, 0)
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                   '在直线p2%,p3%上取一点使得∠p1%pp2%=∠p1%p2%p3%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
                   End If
      ElseIf angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1) = _
          angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0) Then
           p% = is_line_line_intersect(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0), _
                angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1), 0, 0, False)
                  If p% > 0 Then
                   add_aid_point = add_aid_point_for_eangle_(p%, _
                       angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).poi(1), _
                          angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
                             angle3_value(tn%).data(0).data0.angle(1), 0, 0, 0, 0, 0, 0, 0)
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                   If add_aid_point > 1 Then
                      Exit Function
                   End If
                  End If
      End If
End If
ElseIf angle3_value(tn%).data(0).data0.type = eangle_ Then
'For i% = 1 To last_conditions.last_cond(1).eangle_no

 add_aid_point = set_total_triangle_from_eangle_( _
     angle3_value(tn%).data(0).data0.angle(0), angle3_value(tn%).data(0).data0.angle(1))
 If add_aid_point > 1 Then
  Exit Function
 End If
If angle3_value(tn%).data(0).data0.angle_(3) > 0 And (angle3_value(tn%).data(0).data0.ty_(0) = 3 Or _
       angle3_value(tn%).data(0).data0.ty_(0) = 5) Then
 If angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0) = _
      angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1) Then
     add_aid_point = add_point_for_eangle(angle3_value(tn%).data(0).data0.angle(1), _
                          angle3_value(tn%).data(0).data0.angle(0))
          If add_aid_point > 1 Then
           Exit Function
          End If
 ElseIf angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1) = _
      angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0) Then
      add_aid_point = add_point_for_eangle(angle3_value(tn%).data(0).data0.angle(0), _
                            angle3_value(tn%).data(0).data0.angle(1))
          If add_aid_point > 1 Then
           Exit Function
          End If
End If
End If
 If angle3_value(tn%).data(0).data0.angle_(3) > 0 Then
 If angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0) = _
      angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1) Then
  t_l(0) = angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1)
    t_l(1) = angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0)
  Call line_number0(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
     m_lin(t_l(0)).data(0).data0.poi(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).te(1)), t_n(0), t_n(1))
  Call line_number0(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
     m_lin(t_l(1)).data(0).data0.poi(angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).te(0)), t_n(2), t_n(3))
ElseIf angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(1) = _
         angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(0) Then
  t_l(0) = angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).line_no(0)
    t_l(1) = angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).line_no(1)
  Call line_number0(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
     m_lin(t_l(0)).data(0).data0.poi(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).te(0)), t_n(0), t_n(1))
  Call line_number0(angle(angle3_value(tn%).data(0).data0.angle(0)).data(0).poi(1), _
     m_lin(t_l(1)).data(0).data0.poi(angle(angle3_value(tn%).data(0).data0.angle(1)).data(0).te(1)), t_n(2), t_n(3))
Else
 GoTo add_aid_point_for_eangle
End If
add_aid_point = add_point_for_eangle1(t_n(0), t_n(1), t_l(0), _
                 t_n(2), t_n(3), t_l(1))
If add_aid_point > 1 Then
 Exit Function
End If
add_aid_point = add_point_for_eangle1(t_n(2), t_n(3), t_l(1), _
                 t_n(0), t_n(1), t_l(0))
If add_aid_point > 1 Then
 Exit Function
End If
add_aid_point_for_eangle:
End If
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).verti_no
 '垂线垂足
 If Dverti(i%).data(0).inter_poi > 0 Then '垂直相交
  For j% = 1 To last_conditions.last_cond(1).angle_value_no
   tv$ = val_(angle3_value(angle_value.av_no(j%).no).data(0).data0.value)
   If tv$ < 90 And tv$ > 0 Then
    '构造rt三角形
    If angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(0) = _
                      Dverti(i%).data(0).line_no(0) Then
           add_aid_point = add_interset_point_line_line( _
                angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(1), _
                Dverti(i%).data(0).line_no(1), 0, 0, 0, 0, 0, c_data0)
          If add_aid_point > 1 Then
           Exit Function
          End If
        'End If
    ElseIf angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(0) = _
                    Dverti(i%).data(0).line_no(1) Then
           add_aid_point = add_interset_point_line_line( _
                angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(1), _
                Dverti(i%).data(0).line_no(0), 0, 0, 0, 0, 0, c_data0)
          If add_aid_point > 1 Then
           Exit Function
          End If
      '  End If
    ElseIf angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(1) = _
                     Dverti(i%).data(0).line_no(0) Then
           add_aid_point = add_interset_point_line_line( _
                angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(0), _
                 Dverti(i%).data(0).line_no(1), 0, 0, 0, 0, 0, c_data0)
          If add_aid_point > 1 Then
           Exit Function
          End If
        'End If
    ElseIf angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(1) = _
                     Dverti(i%).data(0).line_no(1) Then
            add_aid_point = add_interset_point_line_line( _
                angle(angle3_value(angle_value.av_no(j%).no).data(0).data0.angle(0)).data(0).line_no(0), _
                Dverti(i%).data(0).line_no(0), 0, 0, 0, 0, 0, c_data0)
          If add_aid_point > 1 Then
           Exit Function
          End If
       'End If
    End If
   Else
    '加垂线垂足
            add_aid_point = add_interset_point_line_line(Dverti(i%).data(0).line_no(0), _
                  Dverti(i%).data(0).line_no(1), Dverti(i%).data(0).inter_poi, 0, 0, 0, 0, c_data0)
          If add_aid_point > 1 Then
           Exit Function
          End If
   End If
  Next j%
  End If
Next i%
'*****************************
For p% = 1 + last_conditions.last_cond(0).eline_no To last_conditions.last_cond(1).eline_no
i% = Deline(p%).data(0).record.data1.index.i(0)
add_aid_point = add_aid_point_from_eline(Deline(i%).data(0).data0.poi(0), _
         Deline(i%).data(0).data0.poi(1), Deline(i%).data(0).data0.poi(2), _
          Deline(i%).data(0).data0.poi(3))
If add_aid_point > 1 Then
 Exit Function
End If
Next p%
'*********************************8
For p% = 1 + last_conditions.last_cond(0).mid_point_no To last_conditions.last_cond(1).mid_point_no
i% = Dmid_point(p%).data(0).record.data1.index.i(0)
add_aid_point = add_aid_point2(Dmid_point(i%).data(0).data0.poi(0), _
 Dmid_point(i%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.poi(2), _
  line_number0(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(2), 0, 0))
If add_aid_point > 1 Then
 Exit Function
End If
Next p%
'***********************
'过线段比的分点作平行线
For p% = 2 + last_conditions.last_cond(0).line_value_no To last_conditions.last_cond(1).line_value_no
 i% = line_value(p%).data(0).record.data1.index.i(0)
 If line_value(i%).data(0).record.data0.condition_data.condition_no = 0 Then
   tn_% = m_circle_number(1, 0, pointapi0, line_value(i%).data(0).data0.poi(0), _
                             line_value(i%).data(0).data0.poi(1), 0, 0, 0, 0, 0, _
                              0, 1, 0, False)
   If tn_% > 0 Then
     If m_Circ(tn_%).data(0).data0.center > 0 And m_poi(m_Circ(tn_%).data(0).data0.center).data(0).data0.visible > 0 Then
      add_aid_point = _
           add_mid_point(line_value(i%).data(0).data0.poi(0), 0, line_value(i%).data(0).data0.poi(1), 2)
      If add_aid_point > 1 Then
         Exit Function
      End If
     End If
    End If
End If
 For o% = 1 To p% - 1
 j% = line_value(o%).data(0).record.data1.index.i(0)
  If line_value(i%).data(0).data0.value = line_value(j%).data(0).data0.value Then
   add_aid_point = add_aid_point_from_eline(line_value(i%).data(0).data0.poi(0), _
     line_value(i%).data(0).data0.poi(1), line_value(j%).data(0).data0.poi(0), _
      line_value(j%).data(0).data0.poi(1))
      If add_aid_point > 1 Then
       Exit Function
      End If
  End If
  tl% = line_number0(line_value(i%).data(0).data0.poi(0), line_value(i%).data(0).data0.poi(1), 0, 0)
  If line_value(i%).data(0).data0.line_no = line_value(j%).data(0).data0.line_no Then
  If line_value(i%).data(0).data0.poi(1) = line_value(j%).data(0).data0.poi(0) Then
   tp(0) = line_value(i%).data(0).data0.poi(0)
    tp(1) = line_value(i%).data(0).data0.poi(1)
     tp(2) = line_value(j%).data(0).data0.poi(1)
  ElseIf line_value(j%).data(0).data0.poi(1) = line_value(i%).data(0).data0.poi(0) Then
   tp(0) = line_value(j%).data(0).data0.poi(0)
    tp(1) = line_value(j%).data(0).data0.poi(1)
     tp(2) = line_value(i%).data(0).data0.poi(1)
  Else
   GoTo add_aid_point_mark2
 End If
  add_aid_point = add_aid_point2(tp(0), tp(1), tp(2), tl%)
    If add_aid_point > 1 Then
     Exit Function
    End If
   End If
add_aid_point_mark2:
 Next o%
Next p%
For k% = 0 To last_conclusion - 1
If conclusion_data(k%).ty = relation_ Then
If con_relation(k%).data(0).line_no(0) = con_relation(k%).data(0).line_no(1) Then
 If con_relation(k%).data(0).poi(1) = con_relation(k%).data(0).poi(2) Then
  tp(0) = con_relation(k%).data(0).poi(0)
   tp(1) = con_relation(k%).data(0).poi(1)
    tp(2) = con_relation(k%).data(0).poi(3)
 Else
  GoTo add_aid_point_mark4
 End If
add_aid_point = add_aid_point2(tp(0), tp(1), tp(2), tl%)
   If add_aid_point > 1 Then
    Exit Function
   End If
 End If
ElseIf conclusion_data(k%).ty = eline_ Then
 add_aid_point = add_point_for_con_eline(k%)
End If
add_aid_point_mark4:
Next k%
For p% = 1 + last_conditions.last_cond(0).relation_no To last_conditions.last_cond(1).relation_no
 i% = Drelation(p%).data(0).record.data1.index.i(0)
 If Drelation(i%).data(0).data0.line_no(0) = Drelation(i%).data(0).data0.line_no(1) And _
      Drelation(i%).data(0).data0.poi(1) = Drelation(i%).data(0).data0.poi(2) Then
  tp(0) = Drelation(i%).data(0).data0.poi(0)
   tp(1) = Drelation(i%).data(0).data0.poi(1)
    tp(2) = Drelation(i%).data(0).data0.poi(3)
  add_aid_point = add_aid_point2(tp(0), tp(1), tp(2), Drelation(i%).data(0).data0.line_no(0))
   If add_aid_point > 1 Then
    Exit Function
   End If
 End If
Next p%
For p% = 1 + last_conditions.last_cond(0).dpoint_pair_no To last_conditions.last_cond(1).dpoint_pair_no
 i% = Ddpoint_pair(p%).data(0).record.data1.index.i(0)
 If Ddpoint_pair(i%).data(0).data0.con_line_type(0) = 3 Then
 ElseIf Ddpoint_pair(i%).data(0).data0.con_line_type(1) = 3 Then
 ElseIf Ddpoint_pair(i%).data(0).data0.con_line_type(1) = 5 Then
 End If
Next p%
 add_aid_point = add_aid_point_for_point3_on_line(0, c_data0)
       If add_aid_point > 1 Then
        Exit Function
       End If
For i% = 1 To last_conditions.last_cond(1).point_no
 For j% = 1 To last_conditions.last_cond(1).line_no
  For k% = 1 To last_conditions.last_cond(1).line_no
   If is_point_in_paral_line(i%, 0, 0, tl%) Then
    If j% <> k% And tl% <> j% And tl% <> k% Then
       add_aid_point = add_paral_line(i%, j%, k%, 0, 0, 0, 0, 0, 0)
       If add_aid_point > 1 Then
        Exit Function
       End If
     'End If
   End If
   End If
   Next k%
  Next j%
 Next i%
 For i% = 1 To last_conditions.last_cond(1).aid_point_data1_no
 t_triA(0) = aid_point_data1(i%).data(0).triA(0)
  t_triA(1) = aid_point_data1(i%).data(0).triA(1)
 add_aid_point = add_aid_point_for_t_e_triangle1(t_triA(0), t_triA(1))
 If add_aid_point > 1 Then
  Exit Function
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data2_no
 t_triA(0) = aid_point_data2(i%).data(0).triA(0)
 t_triA(1) = aid_point_data2(i%).data(0).triA(1)
 add_aid_point = add_aid_point_for_t_e_triangle2(t_triA(0), t_triA(1))
 If add_aid_point > 1 Then
  Exit Function
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data3_no
 t_triA(0) = aid_point_data3(i%).data(0).triA(0)
 t_triA(1) = aid_point_data3(i%).data(0).triA(1)
 add_aid_point = add_aid_point_for_t_e_triangle3(t_triA(0), t_triA(1))
 If add_aid_point > 1 Then
  Exit Function
 End If
Next i%
 
For i% = 1 To last_conditions.last_cond(1).point_no
 'If m_poi(i%).data(0).no_reduce = 0 Then
  For j% = 1 To last_conditions.last_cond(1).line_no
   For k% = 1 To last_conditions.last_cond(1).line_no
   'If is_point_in_paral_line(i%, 0, 0) = False Then
    If j% <> k% And is_point_in_line3(i%, m_lin(j%).data(0).data0, 0) = False And _
     is_point_in_line3(i%, m_lin(k%).data(0).data0, 0) = False Then
            add_aid_point = add_paral_line(i%, j%, k%, 0, 0, 0, 0, 0, 0)
        If add_aid_point > 1 Then
         Exit Function
        End If
      ' End If
   End If
   'End If
    Next k%
   Next j%
  'End If
 Next i%
For i% = 1 To last_conditions.last_cond(1).point_no
 For j% = 1 To last_conditions.last_cond(1).line_no
  For k% = 1 To last_conditions.last_cond(1).line_no
    add_aid_point = add_aid_point_for_verti0(i%, j%, k%, 0, cond_data(0), 1)
     If add_aid_point > 1 Then
      Exit Function
     End If
  Next k%
 Next j%
Next i%
Exit Function
add_aid_point_error:
End Function

Public Function from_old_to_aid() As Byte '=1 exit
Dim i%, j%
'On Error GoTo from_old_to_aid_error
If last_conditions_for_aid_no = 8 Then
 from_old_to_aid = 1
  Exit Function
End If
last_conditions_for_aid_no = last_conditions_for_aid_no + 1
ReDim Preserve last_conditions_for_aid(last_conditions_for_aid_no) As last_conditions_for_aid_type
 last_conditions_for_aid(last_conditions_for_aid_no).last_cond(0) = last_conditions.last_cond(0)
 last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1) = last_conditions.last_cond(1)
'**************************************************
'For i% = 1 To last_conditions.last_cond(1).point_no
  'ReDim Preserve poi(i%).data(last_conditions_for_aid_no) As point_data_type
   Call move_point_inner_data(0, last_conditions_for_aid_no)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).circle_no
  Call move_circle_inner_data(0, last_conditions_for_aid_no)
   
 'Circ(i%).data(last_conditions_for_aid_no) = Circ(i%).data(0)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).line_no
  Call move_line_inner_data(0, last_conditions_for_aid_no)
 'lin(i%).data(last_conditions_for_aid_no) = m_lin(i%)
'Next i%
Call from_old_to_aid0
from_old_to_aid_error:
'last_conditions.last_cond(0) = last_conditions.last_cond(1)
End Function

Public Sub from_aid_to_old()
Dim i%, j%, k%, p%
Dim c As circle_data0_type
'****************
'On Error GoTo from_aid_to_old_error
If last_conditions.last_cond(1).point_no > last_conditions.last_cond(0).point_no Then
    p% = last_conditions.last_cond(1).point_no
Else
    p% = 0
End If
If last_conditions_for_aid(last_conditions_for_aid_no).new_result_from_add Then
     last_conditions_for_aid(last_conditions_for_aid_no).new_result_from_add = False
      For i% = last_conditions_for_aid_no - 1 To 0
       last_conditions_for_aid(i%).new_result_from_add = True
      Next i%
Else
 'For i% = last_conditions.last_cond(1).circle_no To last_conditions.last_cond(0).circle_no
  'Circ(i%).data(0).data0 = c
 'Next i%
 'End If
 '*******************************************
'For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_no
Call move_point_inner_data(last_conditions_for_aid_no, 0)
'poi(i%).data(0) = poi(i%).data(last_conditions_for_aid_no)
'Next i%
'For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).circle_no
Call move_circle_inner_data(last_conditions_for_aid_no, 0)
'Circ(i%).data(0) = Circ(i%).data(last_conditions_for_aid_no)
'Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).circle_no + 1 To _
         last_conditions.last_cond(1).circle_no
'Call init_circle0(i%)
Next i%
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_no
'lin(i%).data(0) = lin(i%).data(last_conditions_for_aid_no)
Call move_line_inner_data(last_conditions_for_aid_no, 0)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_no + 1 To _
       last_conditions.last_cond(1).line_no
'Call init_line0(lin(i%).data(0))
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).branch_data_no + 1 To _
       last_conditions.last_cond(1).branch_data_no
'Call delete_branch_data(i%)
Next i%

'If p% > 0 Then
'If poi(p%).data(0).data0.visible = 2 Then
' Call draw_point(Draw_form, poi(p%), 0, delete)
'End If
'End If
Call from_aid_to_old0
 last_conditions.last_cond(0) = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(0)
 last_conditions.last_cond(1) = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1)
 t_condition.last_cond(0) = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(0)
 t_condition.last_cond(1) = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1)
last_conditions_for_aid_no = last_conditions_for_aid_no - 1
End If
from_aid_to_old_error:
End Sub


Public Function add_line(ByVal is_no_initial As Byte) As Byte
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Dim t_n(1) As Integer
Dim tp As POINTAPI
Dim i%, j%, k%, l%, m%, t%, tn%, n%, tn_%, p%, tp_%, tl1%, tl2%
Dim X&, Y&
Dim ty As Boolean
Dim ty1 As Integer
Dim temp_record As total_record_type
Dim c_data0 As condition_data_type
'On Error GoTo add_line_mark1
If from_old_to_aid = 1 Then '超过8个辅助点
    Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_line = 6
  Exit Function
End If
 Call set_point_name(last_conditions.last_cond(1).point_no, _
      next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
 For i% = 1 To last_conditions.last_cond(1).point_no - 1
  For j% = 1 To last_conditions.last_cond(1).line_no
   For k% = 1 To last_conditions.last_cond(1).line_no
       tp_% = i%
       tl1% = j%
       tl2% = k%
    For t% = 0 To 1
'************************************
     If t% = 0 Then
      If set_add_aid_point_for_two_line(tp_%, tl1%, paral_, tl2%, 0) = False Then
          GoTo add_line_next1
      Else
       ty1 = paral_
      End If
     Else
      If set_add_aid_point_for_two_line(tp_%, tl1%, verti_, tl2%, 0) = False Then
          GoTo add_line_next1
      Else
      ty1 = verti_
      End If
     End If
   '********************************************************
If inter_point_line_line3(tp_%, ty1, tl1%, m_lin(tl2%).data(0).data0.poi(0), paral_, tl2%, tp, p%, False, c_data0, True) Then
 '　作平行或垂直的交点
  n% = line_number(last_conditions.last_cond(1).point_no, i%, pointapi0, pointapi0, _
                     depend_condition(point_, last_conditions.last_cond(1).point_no), _
                      depend_condition(point_, i%), condition, condition_color, 1, 0)
   If n% = j% Then
    GoTo add_line_mark1
   End If
  Call add_point_to_line(last_conditions.last_cond(1).point_no, j%, tn_%, no_display, False, 0, temp_record)
    Call set_two_point_line_for_line(j%, temp_record.record_data)
     Call arrange_data_for_new_point(j%, 0)
     If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
       ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
    End If
        last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_ ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(0).new_point_no ' record0
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = p%
      'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
      'new_point(last_conditions.last_cond(1).new_point_no).record_.data1.aid_condition = last_conditions.last_cond(1).new_point_no
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = k%
'*********************************************************
 n% = 0
 If ty1 = paral_ Then
  ty = set_dparal(n%, j%, temp_record, n%, 0, False)
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1325, "\\1\\" + m_poi(i%).data(0).data0.name + _
                        "\\2\\" + m_poi(m_lin(j%).data(0).data0.poi(0)).data(0).data0.name + _
                                 m_poi(m_lin(j%).data(0).data0.poi(1)).data(0).data0.name + _
                        "\\3\\" + m_poi(m_lin(k%).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(k%).data(0).data0.poi(1)).data(0).data0.name + _
                        "\\4\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
        temp_record.record_data.data0.condition_data.condition(1).ty = paral_
         temp_record.record_data.data0.condition_data.condition(1).no = n%
          temp_record.record_data.data0.condition_data.condition_no = 1 'Dparal(n%).record_.data1.aid_condition = last_conditions.last_cond(1).new_point_no
 Else
  ty = set_dverti(n%, j%, temp_record, n%, 0, False)
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1330, "\\1\\" + m_poi(i%).data(0).data0.name + _
                         "\\2\\" + m_poi(m_lin(j%).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(j%).data(0).data0.poi(1)).data(0).data0.name + _
                         "\\3\\" + m_poi(m_lin(k%).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(k%).data(0).data0.poi(1)).data(0).data0.name + _
                         "\\4\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
        temp_record.record_data.data0.condition_data.condition(1).ty = verti_ 'paral
         temp_record.record_data.data0.condition_data.condition(1).no = n%
          temp_record.record_data.data0.condition_data.condition_no = 1
 End If
 '***********************************
 add_line = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
 If add_line > 1 Then
   Exit Function
 Else
 
 'poi(last_conditions.last_cond(1).point_no).data(0).in_line(0) = 0
 'poi(last_conditions.last_cond(1).point_no).data(0).in_circle(0) = 0
 End If
 '***********************************************
  End If

add_line_mark1:
     Call from_aid_to_old
   Next t%
  Next k%
 Next j%
add_line_next1:
Next i%

'%%%%%%%%%%%%%%%%%%%%%%%%%
End Function
Public Function add_aid_point_for_midpoint(ByVal p1%, ByVal p2%, ByVal p3%) As Byte
Dim t_n(1) As Integer
Dim i%, j%, k%, tn%, n%
Dim ty As Boolean
Dim tl(2) As Integer
Dim temp_record As record_type
'On Error GoTo add_aid_point_for_midpoint_error
last_aid_point = last_conditions.last_cond(1).point_no
  tl(0) = line_number0(p1%, p3%, 0, 0)
 For j% = 1 To last_conditions.last_cond(1).point_no
  If j% <> p1% And j% <> p2% And j% <> p3% Then
    tl(1) = line_number0(p2%, j%, 0, 0)
If tl(0) <> tl(1) And m_lin(tl(1)).data(0).data0.visible > 0 And _
   get_midpoint(j%, p2%, 0, 0, 0, 0, 0, 0) = 0 Then
    add_aid_point_for_midpoint = add_mid_point(j%, p2%, 0, 0)
 If add_aid_point_for_midpoint > 1 Then
  Exit Function
 End If
End If
'******
tl(1) = line_number0(p1%, j%, 0, 0)
If tl(0) <> tl(1) And m_lin(tl(1)).data(0).data0.visible > 0 And _
       get_midpoint(p1%, 0, j%, 0, 0, 0, 0, 0) = 0 Then
If is_known_line(p1%, j%) Then
add_aid_point_for_midpoint = add_mid_point(j%, 0, p3%, 0)
 If add_aid_point_for_midpoint > 1 Then
  Exit Function
 End If
End If
add_aid_point_for_midpoint = add_mid_point(p1%, 0, j%, 0)
If add_aid_point_for_midpoint > 1 Then
 Exit Function
 End If
End If
'*****
    tl(1) = line_number0(p3%, j%, 0, 0)
      If tl(0) <> tl(1) And m_lin(tl(1)).data(0).data0.visible > 0 And _
        get_midpoint(p3%, 0, j%, 0, 0, 0, 0, 0) = 0 Then
If is_known_line(p3%, j%) Then
add_aid_point_for_midpoint = add_mid_point(p1%, 0, j%, 0)
If add_aid_point_for_midpoint > 1 Then
 Exit Function
 End If
End If
add_aid_point_for_midpoint = add_mid_point(p3%, 0, j%, 0)
If add_aid_point_for_midpoint > 1 Then
 Exit Function
 End If
End If
End If
Next j%
Exit Function
add_aid_point_for_midpoint_error:
End Function


Public Function add_paral_line(ByVal p%, ByVal l1%, ByVal l2%, p1%, p2%, con_no%, _
                    no%, n_p%, ByVal is_remove As Byte) As Byte
'p1%,P2% 与 新点构成三角形,与梯形面积
Dim temp_record As total_record_type
Dim tn_(1) As Integer
Dim tl(1) As Integer
Dim tp%, l3%
Dim n%, i%, o%, tri_no%
Dim t_p(3) As Integer
Dim c_data0 As condition_data_type
If set_add_aid_point_for_two_line(p%, l1%, 1, l2%, 0) Then
t_p(0) = m_lin(l1%).data(0).data0.poi(0)
t_p(1) = m_lin(l1%).data(0).data0.poi(1)
t_p(2) = m_lin(l2%).data(0).data0.poi(0)
t_p(3) = m_lin(l2%).data(0).data0.poi(1)
For i% = 1 To m_lin(l2%).data(0).data0.in_point(0)
 If is_dparal(l1%, line_number0(p%, m_lin(l2%).data(0).data0.in_point(i%), 0, 0), 0, _
  -1000, 0, 0, 0, 0) Then
   Exit Function
 End If
Next i%
If is_point_in_line3(p%, m_lin(l2%).data(0).data0, 0) Then
   Exit Function
ElseIf is_point_in_line3(p%, m_lin(l1%).data(0).data0, 0) Then
    add_paral_line = add_interset_point_line_line(l1%, l2%, 0, 0, 0, 0, 0, c_data0)
      Exit Function
ElseIf is_point_in_paral(p%, l1%, l3%) Then
    add_paral_line = add_interset_point_line_line(l3%, l2%, 0, 0, 0, 0, 0, c_data0)
      Exit Function
Else
'********************
'On Error GoTo add_paral_line_error
If from_old_to_aid = 1 Then
   Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_paral_line = 6
  Exit Function
End If
 last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 'MDIForm1.Toolbar1.Buttons(21).Image = 33
 n_p% = last_conditions.last_cond(1).point_no
 Call set_point_name(n_p%, next_char(n_p%, "", 0, 0))
If inter_point_line_line3(p%, paral_, l1%, _
       m_lin(l2%).data(0).data0.poi(0), paral_, l2%, t_coord, n_p%, False, c_data0, True) = False Then
        Call from_aid_to_old
Else
 l3% = line_number0(n_p%, p%, tn_(1), 0)
    Call add_point_to_line(n_p%, l2%, tn_(0), no_display, False, 0, temp_record)
     Call set_two_point_line_for_line(l2%, temp_record.record_data)
      Call arrange_data_for_new_point(l2%, 0)
      If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
       ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
      End If
      last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
       temp_record.record_data.data0.condition_data.condition_no = 1 'record0
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
        temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
         temp_record.record_data.data0.theorem_no = 0
    new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = n_p%
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l2%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(1) = l3%
       no% = 0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1325, _
         "\\1\\" + m_poi(p%).data(0).data0.name + _
         "\\2\\" + m_poi(t_p(0)).data(0).data0.name + m_poi(t_p(1)).data(0).data0.name + _
         "\\3\\" + m_poi(t_p(2)).data(0).data0.name + m_poi(t_p(3)).data(0).data0.name + _
         "\\4\\" + m_poi(n_p%).data(0).data0.name)
          add_paral_line = set_dparal(l1%, l3%, temp_record, no%, 0, False)
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = paral_
        temp_record.record_data.data0.condition_data.condition(1).no = no%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = paral_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = no%
      add_paral_line = set_New_point(n_p%, temp_record, l2%, l3%, _
            tn_(0), tn_(1), 0, 0, 0, 1)
       If add_paral_line > 1 Then
        Exit Function
       End If
      If is_remove < 2 Then
      add_paral_line = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
       If add_paral_line > 1 Then
         Exit Function
       End If
       End If
      If p1% > 0 Then
      tri_no% = triangle_number(p1%, p2%, n_p%, 0, 0, 0, 0, 0, 0, 0)
       n% = 0
       If is_area_of_element(triangle_, tri_no, n%, -1000) Then
        If con_Area_of_element(con_no%).data(0).element.ty = polygon_ Then
         temp_record.record_data.data0.condition_data.condition_no = 1
          temp_record.record_data.data0.condition_data.condition(1).ty = area_of_element_
           temp_record.record_data.data0.condition_data.condition(1).no = n%
        add_paral_line = set_area_of_polygon0(con_Area_of_element(con_no%).data(0).element.no, _
           area_of_element(n%).data(0).value, temp_record, 0, 0)
           If add_paral_line > 1 Then
            Exit Function
           End If
         End If
       End If
      End If
add_paral_line_error:
  If is_remove = 0 Then
         Call from_aid_to_old
  End If
End If
End If
End If
End Function

Public Function add_aid_point2(ByVal p1%, ByVal p2%, _
    ByVal p3%, ByVal l%) As Byte 'paral
Dim i%, j%, k%, tl% 'p1,p2,p3位于l%上
Dim m_i%, m_j%, tn_%
Dim tp(2) As Integer
tp(0) = p1%
 tp(1) = p2%
  tp(2) = p3%
For i% = 0 To 2 '中点结论取一点
 For j% = 1 To m_poi(tp(i%)).data(0).in_line(0) 'last_conditions.last_cond(1).line_no 'poi(tp(i%)).data(0).in_line(0)
  '过选取点的任意直线
  If m_lin(m_poi(tp(i%)).data(0).in_line(j%)).data(0).data0.visible > 0 Then '可见直线
   If m_poi(tp(i%)).data(0).in_line(j%) <> l% Then  '点在线上
   For k% = 1 To m_lin(m_poi(tp(i%)).data(0).in_line(j%)).data(0).data0.in_point(0)
    If m_lin(m_poi(tp(i%)).data(0).in_line(j%)).data(0).data0.in_point(k%) <> tp(i%) Then '在线上取另一点
  '过点tp(i%)平行poi(tp(i%)).data(0).in_line(j%)的直线交tl%直线于点
    m_i% = (i% + 1) Mod 3 '中点结论取第二点
     tl% = line_number0(m_lin(m_poi(tp(i%)).data(0).in_line(j%)).data(0).data0.in_point(k%), _
                             tp(m_i%), 0, 0)  '第二点与过第一点的直线上等某点连线
    If tl% > 0 And m_lin(tl%).data(0).data0.visible > 0 Then
    m_i% = (i% + 2) Mod 3 '中点结论取第三点
    add_aid_point2 = add_paral_line(tp(m_i%), m_poi(tp(i%)).data(0).in_line(j%), _
            tl%, 0, 0, 0, 0, 0, 0) '
    If add_aid_point2 > 1 Then
    Exit Function
    End If
    '********************************
    add_aid_point2 = add_paral_line(tp(m_i%), tl%, m_poi(tp(i%)).data(0).in_line(j%), 0, 0, 0, 0, 0, 0)
    If add_aid_point2 > 1 Then
    Exit Function
    End If
    'End If
    End If
    End If
   Next k%
  End If
 End If
 Next j%
Next i%
End Function

Public Function add_aid_point0(ByVal gs%) As Byte
Dim it(6) As item0_data_type
Dim pA(5) As String
Dim tp(5) As Integer
Dim i%, n%, tl%, tn_%, tp_%
Dim p As POINTAPI
Dim temp_con_no%
Dim r1&, r2&, r3&
Dim c_data0 As condition_data_type
Dim temp_record As total_record_type
For i% = 0 To 3
  it(i%) = item0(general_string(gs%).data(0).item(i%)).data(0)
   pA(i%) = general_string(gs%).data(0).para(i%)
Next i%
If (pA(1) = "-1" Or pA(1) = "@1") And (pA(2) = "-1" Or pA(2) = "@1") Then
    it(4) = it(0)
     it(5) = it(1)
      pA(4) = pA(0)
       pA(5) = pA(1)
    it(6) = it(2)
 ElseIf pA(1) = "1" And (pA(2) = "-1" Or pA(2) = "@1") Then
    it(4) = it(2)
     it(5) = it(0)
      pA(4) = pA(2)
       pA(5) = pA(0)
     it(6) = it(1)
 ElseIf (pA(1) = "-1" Or pA(1) = "@1") And pA(2) = "1" Then
    it(4) = it(1)
     it(5) = it(0)
      pA(4) = pA(1)
       pA(5) = pA(0)
        it(6) = it(2)
Else
 Exit Function
End If
If pA(4) = time_string("-1", pA(5), True, False) Then
If it(0).sig = "~" And it(1).sig = "~" And it(2).sig = "~" And _
     it(3).sig = "~" Then
 If pA(0) = "1" Then
    
 End If
ElseIf it(0).sig = "~" And it(1).sig = "~" And it(2).sig = "~" And _
     pA(3) = "0" Then
r1& = sqr((m_poi(it(5).poi(0)).data(0).data0.coordinate.X - m_poi(it(5).poi(1)).data(0).data0.coordinate.X) ^ 2 + _
           (m_poi(it(5).poi(0)).data(0).data0.coordinate.Y - m_poi(it(5).poi(1)).data(0).data0.coordinate.Y) ^ 2)
r2& = sqr((m_poi(it(4).poi(0)).data(0).data0.coordinate.X - m_poi(it(4).poi(1)).data(0).data0.coordinate.X) ^ 2 + _
           (m_poi(it(4).poi(0)).data(0).data0.coordinate.Y - m_poi(it(4).poi(1)).data(0).data0.coordinate.Y) ^ 2)
If r1& < r2& Then
 If it(4).poi(0) = it(5).poi(0) Then
  tp(0) = it(4).poi(0)
   tp(1) = it(4).poi(1)
  tp(2) = it(5).poi(0)
   tp(3) = it(5).poi(1)
 ElseIf it(4).poi(0) = it(5).poi(1) Then
  tp(0) = it(4).poi(0)
   tp(1) = it(4).poi(1)
  tp(2) = it(5).poi(1)
   tp(3) = it(5).poi(0)
 ElseIf it(4).poi(1) = it(5).poi(0) Then
  tp(0) = it(4).poi(1)
   tp(1) = it(4).poi(0)
  tp(2) = it(5).poi(0)
   tp(3) = it(5).poi(1)
 Else
  tp(0) = it(4).poi(1)
   tp(1) = it(4).poi(0)
  tp(2) = it(5).poi(1)
   tp(3) = it(5).poi(0)
 End If
Else
 r3& = r1&
  r1& = r2&
   r2& = r3&
 If it(4).poi(0) = it(5).poi(0) Then
  tp(0) = it(5).poi(0)
   tp(1) = it(5).poi(1)
  tp(2) = it(4).poi(0)
   tp(3) = it(4).poi(1)
 ElseIf it(4).poi(0) = it(5).poi(1) Then
  tp(0) = it(5).poi(1)
   tp(1) = it(5).poi(0)
  tp(2) = it(4).poi(0)
   tp(3) = it(4).poi(1)
 ElseIf it(4).poi(1) = it(5).poi(0) Then
  tp(0) = it(5).poi(1)
   tp(1) = it(5).poi(0)
  tp(2) = it(4).poi(0)
   tp(3) = it(4).poi(1)
 Else
  tp(0) = it(5).poi(1)
   tp(1) = it(5).poi(0)
  tp(2) = it(4).poi(1)
   tp(3) = it(4).poi(0)
 End If
End If
'On Error GoTo add_aid_point0_error
If from_old_to_aid = 1 Then
   Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point0 = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
tp_% = last_conditions.last_cond(1).point_no
     Call get_new_char(tp_%)
 p.X = m_poi(tp(0)).data(0).data0.coordinate.X + _
            (m_poi(tp(1)).data(0).data0.coordinate.X - m_poi(tp(0)).data(0).data0.coordinate.X) * r1& / r2&
 p.Y = m_poi(tp(0)).data(0).data0.coordinate.Y + _
            (m_poi(tp(1)).data(0).data0.coordinate.Y - m_poi(tp(0)).data(0).data0.coordinate.Y) * r1& / r2&
  If read_point(p, 0) > 0 Then
   GoTo add_aid_point0_mark1
  End If
  Call set_point_coordinate(tp_%, p, False)
  tl% = line_number0(tp(0), tp(1), 0, 0)
    Call add_point_to_line(tp_%, tl%, tn_%, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
     If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
     ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
     End If
       last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
    temp_record.record_data.data0.condition_data.condition_no = 1 'record0
    temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = tp_%
       new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = tl%
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
         LoadResString_(1350, "\\1\\" + m_poi(tp(0)).data(0).data0.name + m_poi(tp(1)).data(0).data0.name + _
                             "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
                             "\\3\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
            m_poi(tp(0)).data(0).data0.name + "=" + m_poi(tp(2)).data(0).data0.name + m_poi(tp(3)).data(0).data0.name)
   Call set_equal_dline(tp(0), last_conditions.last_cond(1).point_no, tp(2), tp(3), _
           0, 0, 0, 0, 0, 0, 0, temp_record, n%, 0, 0, 0, 0, False)
      new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = eline_
      new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = eline_
        temp_record.record_data.data0.condition_data.condition(1).no = n%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = eline_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
         'temp_record.data1.aid_condition = 0 'last_conditions.last_cond(1).new_point_no
         temp_record.record_data.data0.theorem_no = 0
 add_aid_point0 = set_New_point(last_conditions.last_cond(1).point_no, temp_record, tl%, 0, _
      tn_%, 0, 0, 0, 0, 1)
   If add_aid_point0 > 1 Then
    Exit Function
   End If
 add_aid_point0 = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point0 > 1 Then
    Exit Function
   End If
add_aid_point0_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
 add_aid_point0 = add_point_from_aid_point_data(tp(2), tp(3), it(6).poi(0), _
         it(6).poi(1), tp(0), tp(1))
ElseIf it(0).sig = "*" And it(1).sig = "*" And it(2).sig = "*" Then
If from_old_to_aid = 1 Then
   Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point0 = 6
  Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point0 = 6
  Exit Function
End If
 last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 'MDIForm1.Toolbar1.Buttons(21).Image = 33
    Call get_new_char(last_conditions.last_cond(1).point_no)
  add_aid_point0 = add_point_from_point_pair(it(4).poi(0), it(4).poi(1), it(5).poi(0), _
    it(5).poi(1), it(5).poi(2), it(5).poi(3), it(4).poi(2), it(4).poi(3), _
      last_conditions.last_cond(1).point_no)
  If add_aid_point0 > 1 Then
   Exit Function
  End If
add_aid_point0_error:
    Call from_aid_to_old
End If
End If
End Function

Public Function add_point_from_point_pair(ByVal p0%, ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
    ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%) As Byte
Dim tl%, tn_%
Dim p As POINTAPI
Dim A(1) As Integer
Dim tp(7) As Integer
Dim r(3) As Single
Dim l(3) As Integer
Dim n(7) As Integer
Dim dn(3) As Integer
Dim ty As Byte
Dim c_data0 As condition_data_type
Dim temp_record As total_record_type
'temp_record.condition_data.condition_no = 255
'共线的三
If p0% = p2% Then
tp(0) = p1%
tp(1) = p0%
tp(2) = p3%
ElseIf p0% = p3% Then
tp(0) = p1%
tp(1) = p0%
tp(2) = p2%
ElseIf p1% = p2% Then
tp(0) = p0%
tp(1) = p1%
tp(2) = p3%
ElseIf p1% = p3% Then
tp(0) = p0%
tp(1) = p1%
tp(2) = p2%
Else
 GoTo add_point_from_point_pair_mark0
End If

If p4% = p6% Then
tp(3) = p5%
tp(4) = p4%
tp(5) = p7%
ElseIf p4% = p7% Then
tp(3) = p5%
tp(4) = p4%
tp(5) = p6%
ElseIf p5% = p6% Then
tp(3) = p4%
tp(4) = p5%
tp(5) = p7%
ElseIf p5% = p7% Then
tp(3) = p4%
tp(4) = p5%
tp(5) = p6%
Else
 GoTo add_point_from_point_pair_mark0
End If
 GoTo add_point_from_point_pair_mark1
add_point_from_point_pair_mark0:
If p0% = p4% Then
tp(0) = p1%
tp(1) = p0%
tp(2) = p5%
ElseIf p0% = p5% Then
tp(0) = p1%
tp(1) = p0%
tp(2) = p4%
ElseIf p1% = p4% Then
tp(0) = p0%
tp(1) = p1%
tp(2) = p5%
ElseIf p1% = p5% Then
tp(0) = p0%
tp(1) = p1%
tp(2) = p4%
Else
Exit Function
End If
If p2% = p6% Then
tp(3) = p3%
tp(4) = p2%
tp(5) = p7%
ElseIf p2% = p7% Then
tp(3) = p2%
tp(4) = p3%
tp(5) = p6%
ElseIf p3% = p6% Then
tp(3) = p2%
tp(4) = p3%
tp(5) = p7%
ElseIf p3% = p7% Then
tp(3) = p2%
tp(4) = p3%
tp(5) = p6%
Else
Exit Function
End If
add_point_from_point_pair_mark1:
'On Error GoTo add_point_from_point_pair_mark2
If from_old_to_aid = 1 Then
   Exit Function
End If
If is_equal_angle(Abs(angle_number(tp(0), tp(1), tp(2), 0, 0)), _
   Abs(angle_number(tp(3), tp(4), tp(5), 0, 0)), dn(1), dn(2)) Then
    If tp(2) = tp(3) Then
     Call exchange_two_integer(tp(0), tp(2))
      Call exchange_two_integer(tp(3), tp(5))
    End If
 r(0) = sqr((m_poi(tp(0)).data(0).data0.coordinate.X - m_poi(tp(1)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(tp(0)).data(0).data0.coordinate.Y - m_poi(tp(1)).data(0).data0.coordinate.Y) ^ 2)
  r(1) = sqr((m_poi(tp(2)).data(0).data0.coordinate.X - m_poi(tp(1)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(tp(2)).data(0).data0.coordinate.Y - m_poi(tp(1)).data(0).data0.coordinate.Y) ^ 2)
   r(2) = sqr((m_poi(tp(3)).data(0).data0.coordinate.X - m_poi(tp(4)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(tp(3)).data(0).data0.coordinate.Y - m_poi(tp(4)).data(0).data0.coordinate.Y) ^ 2)
    r(3) = sqr((m_poi(tp(5)).data(0).data0.coordinate.X - m_poi(tp(4)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(tp(5)).data(0).data0.coordinate.Y - m_poi(tp(4)).data(0).data0.coordinate.Y) ^ 2)
 p.X = m_poi(tp(4)).data(0).data0.coordinate.X + _
     ((m_poi(tp(5)).data(0).data0.coordinate.X - m_poi(tp(4)).data(0).data0.coordinate.X) * r(1) / r(3)) * r(2) / r(0)
 p.Y = m_poi(tp(4)).data(0).data0.coordinate.Y + _
     ((m_poi(tp(5)).data(0).data0.coordinate.Y - m_poi(tp(4)).data(0).data0.coordinate.Y) * r(1) / r(3)) * r(2) / r(0)
If read_point(p, 0) > 0 Then
 GoTo add_point_from_point_pair_mark2
End If
 Call set_point_coordinate(p8%, p, False)
tl% = line_number0(tp(4), tp(5), 0, 0)
Call add_point_to_line(p8%, tl%, tn_%, no_display, False, 0)
     Call set_two_point_line_for_line(tl%, temp_record.record_data)
      Call arrange_data_for_new_point(tl%, 0)
If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
End If
last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
'ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no).data(1) As new_point_data_type
new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = tl%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
            "\\1\\" + m_poi(tp(4)).data(0).data0.name + m_poi(tp(5)).data(0).data0.name + _
            "\\2\\" + m_poi(p8%).data(0).data0.name + _
            "\\3\\" + m_poi(p8%).data(0).data0.name + m_poi(tp(4)).data(0).data0.name + ":" + _
            m_poi(tp(3)).data(0).data0.name + m_poi(tp(4)).data(0).data0.name + "=" + _
            m_poi(tp(2)).data(0).data0.name + m_poi(tp(1)).data(0).data0.name + ":" + _
             m_poi(tp(0)).data(0).data0.name + m_poi(tp(1)).data(0).data0.name)
 dn(0) = 0
  temp_record.record_data.data0.condition_data.condition_no = 1 'record0
   temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
     add_point_from_point_pair = set_dpoint_pair(p8%, tp(4), tp(3), tp(4), tp(2), tp(1), tp(0), tp(1), _
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
           temp_record, True, dn(0), 0, 0, 0, False)
    If add_point_from_point_pair > 1 Then
     Exit Function
    End If
       new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = dpoint_pair_
       new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = dn(0)
temp_record.record_data.data0.condition_data.condition_no = 1
 temp_record.record_data.data0.condition_data.condition(1).ty = dpoint_pair_
  temp_record.record_data.data0.condition_data.condition(1).no = dn(0)
  temp_record.record_data.data0.theorem_no = 0
 add_point_from_point_pair = set_New_point(p8%, temp_record, tl%, 0, _
    tn_%, 0, 0, 0, 0, 1)
    If add_point_from_point_pair > 1 Then
     Exit Function
    End If
 add_point_from_point_pair = start_prove(0, 1, 1)
    If add_point_from_point_pair > 1 Then
     Exit Function
    End If
add_point_from_point_pair_mark2:
Call from_aid_to_old
End If
End Function

Public Function add_point_from_general_string() As Byte
Dim i%, temp_con_no%
'On Error GoTo add_point_from_general_string_error
 For i% = 1 To last_conclusion
  If conclusion_data(i% - 1).no(0) = 0 And _
   conclusion_data(i% - 1).ty = general_string_ Then
   temp_con_no% = i%
    GoTo add_point_form_general_string0
 End If
 Exit Function
 Next i%
add_point_form_general_string0:
For i% = last_conditions.last_cond(1).general_string_no To 1 Step -1
 If general_string(i%).data(0).record.data0.condition_data.condition_no = 254 Then
 add_point_from_general_string = add_aid_point0(i%)
  If add_point_from_general_string > 1 Then
   Exit Function
 End If
 End If
Next i%
add_point_from_general_string_error:
End Function

Public Function add_point_for_tixing() As Byte
Dim i%, tp%, n_p%
Dim tl(4) As Integer
Dim tn_(1) As Integer
Dim temp_record As total_record_type
Dim c_data0 As condition_data_type
'temp_record.condition_data.condition_no = 255
'On Error GoTo add_point_for_tixing_error
For i% = 1 To last_conditions.last_cond(1).tixing_no
 tl(0) = line_number0(Dtixing(i%).data(0).poi(0), Dtixing(i%).data(0).poi(1), 0, 0)
 tl(1) = line_number0(Dtixing(i%).data(0).poi(1), Dtixing(i%).data(0).poi(2), 0, 0)
 tl(2) = line_number0(Dtixing(i%).data(0).poi(2), Dtixing(i%).data(0).poi(3), 0, 0)
 tl(3) = line_number0(Dtixing(i%).data(0).poi(3), Dtixing(i%).data(0).poi(0), 0, 0)
  add_point_for_tixing = add_interset_point_line_line(tl(1), tl(3), 0, 0, 0, 0, 0, c_data0)
   If add_point_for_tixing > 1 Then
    Exit Function
   End If
 tp% = get_midpoint(Dtixing(i%).data(0).poi(0), 0, Dtixing(i%).data(0).poi(3), _
        0, 0, 0, 0, 0)
  If tp% > 0 Then
   add_point_for_tixing = add_interset_point_line_line(tl(0), _
     line_number0(tp%, Dtixing(i%).data(0).poi(2), 0, 0), 0, 0, 0, 0, 0, c_data0)
   If add_point_for_tixing > 1 Then
    Exit Function
   End If
   add_point_for_tixing = add_interset_point_line_line(tl(2), _
     line_number0(tp%, Dtixing(i%).data(0).poi(1), 0, 0), 0, 0, 0, 0, 0, c_data0)
   If add_point_for_tixing > 1 Then
    Exit Function
   End If
  End If
 tp% = get_midpoint(Dtixing(i%).data(0).poi(1), 0, Dtixing(i%).data(0).poi(2), _
        0, 0, 0, 0, 0)
  If tp% > 0 Then
   add_point_for_tixing = add_interset_point_line_line(tl(0), _
     line_number0(tp%, Dtixing(i%).data(0).poi(3), 0, 0), 0, 0, 0, 0, 0, c_data0)
   If add_point_for_tixing > 1 Then
    Exit Function
   End If
   add_point_for_tixing = add_interset_point_line_line(tl(2), _
     line_number0(tp%, Dtixing(i%).data(0).poi(2), 0, 0), 0, 0, 0, 0, 0, c_data0)
   If add_point_for_tixing > 1 Then
    Exit Function
   End If
  End If
  tl(4) = line_number0(Dtixing(i%).data(0).poi(0), Dtixing(i%).data(0).poi(2), 0, 0)
  '过一个顶点平行一条对角线交另一底边
  add_point_for_tixing = add_paral_line(Dtixing(i%).data(0).poi(0), tl(4), tl(2), 0, 0, 0, 0, 0, 0)
     If add_point_for_tixing > 1 Then
    Exit Function
   End If
Next i%
add_point_for_tixing_error:
End Function

Public Sub from_aid_to_old0()
Dim i%, j%
'On Error GoTo from_aid_to_old0_error
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).value_string_no
 Dvalue_string(i%).data(0).index(0) = _
  Dvalue_string(i%).data(last_conditions_for_aid_no).index(0)
 Dvalue_string(i%).data(0).index(1) = _
  Dvalue_string(i%).data(last_conditions_for_aid_no).index(1)
 Dvalue_string(i%).data(0).value = _
  Dvalue_string(i%).data(last_conditions_for_aid_no).value
 Call copy_factor_to_factor(Dvalue_string(i%).data(last_conditions_for_aid_no).factor, _
  Dvalue_string(i%).data(0).factor)
Next i%
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data1_no
 aid_point_data1(i%).data(0) = _
  aid_point_data1(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data1_no + 1 To _
               last_conditions.last_cond(1).aid_point_data1_no
 aid_point_data1(i%) = _
    aid_point_data1(0)
Next i%
'***********************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data2_no
 aid_point_data2(i%).data(0) = _
  aid_point_data2(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data2_no + 1 To _
               last_conditions.last_cond(1).aid_point_data2_no
 aid_point_data2(i%) = _
    aid_point_data2(0)
Next i%
'********************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data3_no
 aid_point_data3(i%).data(0) = _
  aid_point_data3(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).aid_point_data3_no + 1 To _
               last_conditions.last_cond(1).aid_point_data3_no
 aid_point_data3(i%) = _
    aid_point_data3(0)
Next i%
'************************************************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).same_three_lines_no
 same_three_lines(i%).data(0) = _
  same_three_lines(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).same_three_lines_no + 1 To _
               last_conditions.last_cond(1).same_three_lines_no
 same_three_lines(i%) = same_three_lines(0)
Next i%
'************************************************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).distance_of_paral_line_no
 Ddistance_of_paral_line(i%).data(0) = _
  Ddistance_of_paral_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).distance_of_paral_line_no + 1 To _
               last_conditions.last_cond(1).distance_of_paral_line_no
 Ddistance_of_paral_line(i%) = Ddistance_of_paral_line(0)
Next i%
'************************************************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).distance_of_point_line_no
 Ddistance_of_point_line(i%).data(0) = _
  Ddistance_of_point_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).distance_of_point_line_no + 1 To _
               last_conditions.last_cond(1).distance_of_point_line_no
 Ddistance_of_point_line(i%) = Ddistance_of_point_line(0)
Next i%
'*************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).length_of_polygon_no
 length_of_polygon(i%).data(0) = _
  length_of_polygon(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).length_of_polygon_no + 1 To _
        last_conditions.last_cond(1).length_of_polygon_no
 length_of_polygon(i%) = length_of_polygon(0)
Next i%
'***********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_from_two_point_no
 Dtwo_point_line(i%).data(0) = _
  Dtwo_point_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_from_two_point_no + 1 To _
        last_conditions.last_cond(1).line_from_two_point_no
 Dtwo_point_line(i%) = Dtwo_point_line(0)
Next i%
'***********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle_no
   angle(i%).data(0) = angle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle_no + 1 To _
                last_conditions.last_cond(1).angle_no
   angle(i%) = angle(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).total_angle_no
   T_angle(i%).data(0) = T_angle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).total_angle_no + 1 To _
      last_conditions.last_cond(1).total_angle_no
   T_angle(i%) = T_angle(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).item0_no
   item0(i%).data(0) = item0(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).item0_no + 1 To _
       last_conditions.last_cond(1).item0_no
   item0(i%) = item0(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dline1_no '3
    Dline1(i%).data(0) = _
     Dline1(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dline1_no + 1 To _
           last_conditions.last_cond(1).dline1_no '3
    Dline1(i%) = Dline1(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dpoint_pair_no '3
    Ddpoint_pair(i%).data(0) = _
     Ddpoint_pair(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dpoint_pair_no + 1 To _
        last_conditions.last_cond(1).dpoint_pair_no '3
    Ddpoint_pair(i%) = Ddpoint_pair(0)
Next i%
'***********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dangle_no '3
    Dangle(i%).data(0) = _
     Dangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).dangle_no + 1 To _
      last_conditions.last_cond(1).dangle_no '3
    Dangle(i%) = Dangle(0)
Next i%
'***********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).triangle_no '4
    triangle(i%).data(0) = _
     triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).triangle_no + 1 To _
       last_conditions.last_cond(1).triangle_no '4
    triangle(i%) = triangle(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).rtriangle_no '4
    Rtriangle(i%).data(0) = _
     Rtriangle(i%).data(last_conditions_for_aid_no)
 Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).rtriangle_no + 1 To _
        last_conditions.last_cond(1).rtriangle_no '4
    Rtriangle(i%) = Rtriangle(0)
 Next i%
'***********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_relation_no '4
    Darea_relation(i%).data(0) = _
     Darea_relation(i%).data(last_conditions_for_aid_no)
 Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_relation_no + 1 To _
      last_conditions.last_cond(1).area_relation_no '4
    Darea_relation(i%) = Darea_relation(0)
 Next i%
'***********************
'Deangle.last_no(1) = last_conditions.last_cond(1).Deangle.old_last_no
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).mid_point_line_no '7
      mid_point_line(i%).data(0) = _
       mid_point_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).mid_point_line_no + 1 To _
     last_conditions.last_cond(1).mid_point_line_no '7
      mid_point_line(i%) = mid_point_line(0)
Next i%
'************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).eline_no '8
     Deline(i%).data(0) = _
      Deline(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).eline_no + 1 To _
      last_conditions.last_cond(1).eline_no '8
     Deline(i%) = Deline(0)
Next i%
'**********************************
'************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_3angle_no '8
     equal_3angle(i%).data(0) = _
      equal_3angle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_3angle_no + 1 To _
      last_conditions.last_cond(1).equal_3angle_no '8
     equal_3angle(i%) = equal_3angle(0)
Next i%

'*************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).function_of_angle_no
    function_of_angle(i%).data(0) = _
     function_of_angle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).function_of_angle_no + 1 To _
       last_conditions.last_cond(1).function_of_angle_no
    function_of_angle(i%) = function_of_angle(0)
Next i%
'*************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).four_point_on_circle_no '9
    four_point_on_circle(i%).data(0) = _
     four_point_on_circle(i%).data(last_conditions_for_aid_no)
Next i%

For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).four_point_on_circle_no + 1 To _
      last_conditions.last_cond(1).four_point_on_circle_no '9
    four_point_on_circle(i%) = four_point_on_circle(0)
Next i%
'*************************
'***************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).three_point_on_circle_no '9
    three_point_on_circle(i%).data(0) = _
     three_point_on_circle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).three_point_on_circle_no + 1 To _
      last_conditions.last_cond(1).three_point_on_circle_no '9
    three_point_on_circle(i%) = three_point_on_circle(0)
Next i%

'For i% = 1 To last_angle_relation '10
 '     angle_relation(i%).record = _
  '     angle_relation(i%).record_0
'Next i%
'angle_relation.last_no(1) = last_conditions.last_cond(1).angle_relation.old_last_no
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).mid_point_no '12
     Dmid_point(i%).data(0) = _
      Dmid_point(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).mid_point_no + 1 To _
      last_conditions.last_cond(1).mid_point_no '12
     Dmid_point(i%) = Dmid_point(0)
Next i%
'***************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).paral_no '13
     Dparal(i%).data(0) = _
      Dparal(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).paral_no + 1 To _
      last_conditions.last_cond(1).paral_no '13
     Dparal(i%) = Dparal(0)
Next i%
'*************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).parallelogram_no
     Dparallelogram(i%).data(0) = _
      Dparallelogram(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).parallelogram_no + 1 To _
       last_conditions.last_cond(1).parallelogram_no
     Dparallelogram(i%) = Dparallelogram(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_no '16
     Drelation(i%).data(0) = _
      Drelation(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_no + 1 To _
     last_conditions.last_cond(1).relation_no '16
     Drelation(i%) = Drelation(0)
Next i%
'*********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).v_relation_no  '16
     v_Drelation(i%).data(0) = _
       v_Drelation(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).v_relation_no + 1 To _
     last_conditions.last_cond(1).v_relation_no  '16
     v_Drelation(i%) = v_Drelation(0)
Next i%
'*********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_on_line_no '16
     relation_on_line(i%).data(0) = _
       relation_on_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_on_line_no + 1 To _
     last_conditions.last_cond(1).relation_on_line_no '16
     relation_on_line(i%) = relation_on_line(0)
Next i%
'*********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_string_no '16
     relation_string(i%).data(0) = _
       relation_string(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).relation_string_no + 1 To _
     last_conditions.last_cond(1).relation_string_no '16
     relation_string(i%) = relation_string(0)
Next i%
'********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).similar_triangle_no '17
     Dsimilar_triangle(i%).data(0) = _
      Dsimilar_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).similar_triangle_no + 1 To _
      last_conditions.last_cond(1).similar_triangle_no '17
     Dsimilar_triangle(i%) = Dsimilar_triangle(0)
Next i%
'******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).total_equal_triangle_no '18
     Dtotal_equal_triangle(i%).data(0) = _
      Dtotal_equal_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).total_equal_triangle_no + 1 To _
      last_conditions.last_cond(1).total_equal_triangle_no '18
     Dtotal_equal_triangle(i%) = Dtotal_equal_triangle(0)
Next i%
'********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_total_equal_triangle_no '18
     pseudo_total_equal_triangle(i%).data(0) = _
      pseudo_total_equal_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_total_equal_triangle_no + 1 To _
      last_conditions.last_cond(1).pseudo_total_equal_triangle_no '18
     pseudo_total_equal_triangle(i%) = _
      pseudo_total_equal_triangle(0)
Next i%
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_similar_triangle_no '18
     pseudo_similar_triangle(i%).data(0) = _
      pseudo_similar_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_similar_triangle_no + 1 To _
     last_conditions.last_cond(1).pseudo_similar_triangle_no '18
     pseudo_similar_triangle(i%) = _
      pseudo_similar_triangle(0)
Next i%
'*****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).three_point_on_line_no '20
     three_point_on_line(i%).data(0) = _
      three_point_on_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).three_point_on_line_no + 1 To _
     last_conditions.last_cond(1).three_point_on_line_no '20
     three_point_on_line(i%) = three_point_on_line(0)
Next i%
'******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tri_function_no  '20
     tri_function(i%).data(0) = _
      tri_function(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tri_function_no + 1 To _
     last_conditions.last_cond(1).tri_function_no '20
     tri_function(i%) = _
      tri_function(0)
Next i%
'******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_point_conset_no '20
     two_point_conset(i%).data(0) = _
      two_point_conset(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_point_conset_no + 1 To _
     last_conditions.last_cond(1).two_point_conset_no '20
     two_point_conset(i%) = _
      two_point_conset(0)
Next i%
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_line_value_no
     two_line_value(i%).data(0) = _
      two_line_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_line_value_no + 1 To _
     last_conditions.last_cond(1).two_line_value_no
     two_line_value(i%) = _
      two_line_value(0)
Next i%
'****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line3_value_no
    line3_value(i%).data(0) = _
      line3_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line3_value_no + 1 To _
    last_conditions.last_cond(1).line3_value_no
    line3_value(i%) = _
      line3_value(0)
Next i%
'************
'For i% = 1 To Last_two_angle_value '22
 '    Two_angle_value(i%).record = _
  '    Two_angle_value(i%).record_0
'Next i%
'two_angle_value_sum.last_no(1) = _
 'two_angle_value_sum.old_last_no
'Two_angle_value_90.last_no(1) = _
 'Two_angle_value_90.old_last_no
'Two_angle_value_180.last_no(1) = _
 'Two_angle_value_180.old_last_no
'For i% = 1 To last_conditions.last_cond(1).two_angle_value_180_no '24
    ' Dverti(i%).data(0) = _
      Dverti(i%).data(1)
'Next i%
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).verti_no '24
     Dverti(i%).data(0) = _
      Dverti(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).verti_no + 1 To _
     last_conditions.last_cond(1).verti_no '24
     Dverti(i%) = Dverti(0)
Next i%
'****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).arc_no
 arc(i%).data(0) = arc(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).arc_no + 1 To _
      last_conditions.last_cond(1).arc_no
 arc(i%) = arc(0)
Next i%
'**************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).arc_value_no '25
    arc_value(i%).data(0) = _
      arc_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).arc_value_no + 1 To _
     last_conditions.last_cond(1).arc_value_no '25
    arc_value(i%) = arc_value(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_arc_no '26
     equal_arc(i%).data(0) = _
      equal_arc(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_arc_no + 1 To _
          last_conditions.last_cond(1).equal_arc_no '26
     equal_arc(i%) = equal_arc(0)
Next i%
'**************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).ratio_of_two_arc_no '27
     ratio_of_two_arc(i%).data(0) = _
      ratio_of_two_arc(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).ratio_of_two_arc_no + 1 To _
     last_conditions.last_cond(1).ratio_of_two_arc_no '27
     ratio_of_two_arc(i%) = _
      ratio_of_two_arc(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle_less_angle_no '28
     angle_less_angle(i%).data(0) = _
      angle_less_angle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle_less_angle_no + 1 To _
     last_conditions.last_cond(1).angle_less_angle_no '28
     angle_less_angle(i%) = _
      angle_less_angle(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_less_line_no '29
     line_less_line(i%).data(0) = _
      line_less_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_less_line_no + 1 To _
     last_conditions.last_cond(1).line_less_line_no '29
     line_less_line(i%) = _
      line_less_line(0)
Next i%
'*************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_less_line2_no '30
     line_less_line2(i%).data(0) = _
      line_less_line2(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_less_line2_no + 1 To _
     last_conditions.last_cond(1).line_less_line2_no '30
     line_less_line2(i%) = _
      line_less_line2(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line2_less_line2_no '31
     line2_less_line2(i%).data(0) = _
      line2_less_line2(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line2_less_line2_no + 1 _
    To last_conditions.last_cond(1).line2_less_line2_no '31
     line2_less_line2(i%) = _
      line2_less_line2(0)
Next i%
'**************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle3_value_no '32
    angle3_value(i%).data(0) = _
     angle3_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).angle3_value_no + 1 To _
    last_conditions.last_cond(1).angle3_value_no '32
    angle3_value(i%) = _
     angle3_value(0)
Next i%
'*************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_value_no '33
     line_value(i%).data(0) = _
      line_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).line_value_no + 1 To _
     last_conditions.last_cond(1).line_value_no '33
     line_value(i%) = _
      line_value(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tangent_line_no '34
     tangent_line(i%).data(0) = _
      tangent_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tangent_line_no + 1 To _
     last_conditions.last_cond(1).tangent_line_no '34
     tangent_line(i%) = _
      tangent_line(0)
Next i%
'******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tangent_circle_no '34
     m_tangent_circle(i%).data(0) = _
      m_tangent_circle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tangent_circle_no + 1 To _
    last_conditions.last_cond(1).tangent_circle_no '34
     m_tangent_circle(i%) = _
      m_tangent_circle(0)
Next i%
'****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_right_triangle_no '35
    equal_side_right_triangle(i%).data(0) = _
     equal_side_right_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_right_triangle_no + 1 To _
    last_conditions.last_cond(1).equal_side_right_triangle_no '35
    equal_side_right_triangle(i%) = _
     equal_side_right_triangle(0)
Next i%
'****************
'For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_area_triangle_no '35
 '   equal_area_triangle(i%).data(0) = _
  '   equal_area_triangle(i%).data(last_conditions_for_aid_no)
'Next i%
'For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_area_triangle_no + 1 To _
'    last_conditions.last_cond(1).equal_area_triangle_no '35
'    equal_area_triangle(i%) = _
'     equal_area_triangle(0)
'Next i%
'****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).general_string_no '36
     general_string(i%).data(0) = _
      general_string(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).general_string_no + 1 To _
     last_conditions.last_cond(1).general_string_no '36
     general_string(i%) = _
      general_string(0)
Next i%
'***************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).general_angle_string_no  '37
     general_angle_string(i%).data(0) = _
      general_angle_string(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).general_angle_string_no + 1 To _
     last_conditions.last_cond(1).general_angle_string_no '37
     general_angle_string(i%) = _
      general_angle_string(0)
Next i%
'For i% = 1 To last_conditions.last_cond(1).poly_no
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_pair_for_similar_no   '38
     point_pair_for_similar(i%).data(0) = _
      point_pair_for_similar(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).point_pair_for_similar_no + 1 To _
     last_conditions.last_cond(1).point_pair_for_similar_no '38
     point_pair_for_similar(i%) = _
      point_pair_for_similar(0)
Next i%
'*********************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_dpoint_pair_no   '38
     pseudo_dpoint_pair(i%).data(0) = _
      pseudo_dpoint_pair(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_dpoint_pair_no + 1 To _
     last_conditions.last_cond(1).pseudo_dpoint_pair_no '38
     pseudo_dpoint_pair(i%) = _
      pseudo_dpoint_pair(0)
Next i%
'*********************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_midpoint_no   '38
     pseudo_mid_point(i%).data(0) = _
      pseudo_mid_point(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_midpoint_no + 1 To _
     last_conditions.last_cond(1).pseudo_midpoint_no '38
     pseudo_mid_point(i%) = _
      pseudo_mid_point(0)
Next i%

'*******************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_relation_no   '38
     pseudo_relation(i%).data(0) = _
      pseudo_relation(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_relation_no + 1 To _
     last_conditions.last_cond(1).pseudo_relation_no '38
     pseudo_relation(i%) = _
      pseudo_relation(0)
Next i%
'*********************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_line3_value_no   '38
     pseudo_line3_value(i%).data(0) = _
      pseudo_line3_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_line3_value_no + 1 To _
     last_conditions.last_cond(1).pseudo_line3_value_no '38
     pseudo_line3_value(i%) = _
      pseudo_line3_value(0)
Next i%
'************************************************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_eline_no   '38
     pseudo_eline(i%).data(0) = _
      pseudo_eline(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).pseudo_eline_no + 1 To _
     last_conditions.last_cond(1).pseudo_eline_no '38
     pseudo_eline(i%) = _
      pseudo_eline(0)
Next i%
'**********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).polygon4_no   '38
     Dpolygon4(i%).data(0) = _
       Dpolygon4(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).polygon4_no + 1 To _
     last_conditions.last_cond(1).polygon4_no '38
     Dpolygon4(i%) = Dpolygon4(0)
Next i%
'*******************
'For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_tixing_no '38
 '    Dequal_side_tixing(i%).data(0) = _
      Dequal_side_tixing(i%).data(1)
'Next i%
'For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_tixing_no + 1 To _
 '    last_conditions.last_cond(1).equal_side_tixing_no '38
  '   Dequal_side_tixing(i%) = _
   '   Dequal_side_tixing(0)
'Next i%
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_triangle_no '38
     equal_side_triangle(i%).data(0) = _
      equal_side_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equal_side_triangle_no + 1 To _
    last_conditions.last_cond(1).equal_side_triangle_no '38
     equal_side_triangle(i%) = _
      equal_side_triangle(0)
Next i%
'**************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equation_no '38
     equation(i%).data(0) = _
      equation(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).equation_no + 1 To _
   last_conditions.last_cond(1).equation_no '38
     equation(i%) = _
      equation(0)
Next i%
'**************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).epolygon_no '39
     epolygon(i%).data(0) = _
      epolygon(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).epolygon_no + 1 To last_conditions.last_cond(1).epolygon_no '39
     epolygon(i%) = epolygon(0)
Next i%
'******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tixing_no '40
     Dtixing(i%).data(0) = _
      Dtixing(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).tixing_no + 1 To last_conditions.last_cond(1).tixing_no '40
     Dtixing(i%) = Dtixing(0)
Next i%
'*****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).rhombus_no '41
    rhombus(i%).data(0) = _
      rhombus(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).rhombus_no + 1 To _
                   last_conditions.last_cond(1).rhombus_no '41
    rhombus(i%) = rhombus(0)
Next i%
'*****************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).long_squre_no '42
     Dlong_squre(i%).data(0) = _
      Dlong_squre(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).long_squre_no + 1 To _
                  last_conditions.last_cond(1).long_squre_no '42
     Dlong_squre(i%) = Dlong_squre(0)
Next i%
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).squre_no '42
     Dsqure(i%).data(0) = _
      Dsqure(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).squre_no + 1 To _
               last_conditions.last_cond(1).squre_no '42
     Dsqure(i%) = Dsqure(0)
Next i%
'***********************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_element_no '43
     area_of_element(i%).data(0) = _
      area_of_element(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_element_no + 1 To _
                 last_conditions.last_cond(1).area_of_element_no '43
     area_of_element(i%) = area_of_element(0)
Next i%
'********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_area_of_element_value_no '43
     two_area_of_element_value(i%).data(0) = _
      two_area_of_element_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).two_area_of_element_value_no + 1 To _
                 last_conditions.last_cond(1).two_area_of_element_value_no '43
     two_area_of_element_value(i%) = two_area_of_element_value(0)
Next i%
'***************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_circle_no '44
     area_of_circle(i%).data(0) = _
      area_of_circle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_circle_no + 1 To _
       last_conditions.last_cond(1).area_of_circle_no '44
     area_of_circle(i%) = area_of_circle(0)
Next i%
'*************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_fan_no '46
     Area_of_fan(i%).data(0) = _
      Area_of_fan(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).area_of_fan_no + 1 To _
      last_conditions.last_cond(1).area_of_fan_no '46
     Area_of_fan(i%) = Area_of_fan(0)
Next i%
'*********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).sides_length_of_triangle_no '47
     Sides_length_of_triangle(i%).data(0) = _
      Sides_length_of_triangle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).sides_length_of_triangle_no + 1 To _
      last_conditions.last_cond(1).sides_length_of_triangle_no '47
     Sides_length_of_triangle(i%) = Sides_length_of_triangle(0)
Next i%
'*********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).sides_length_of_circle_no '48
     Sides_length_of_circle(i%).data(0) = _
      Sides_length_of_circle(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).sides_length_of_circle_no + 1 To _
      last_conditions.last_cond(1).sides_length_of_circle_no '48
     Sides_length_of_circle(i%) = Sides_length_of_circle(0)
Next i%
'********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).verti_mid_line_no '49
     verti_mid_line(i%).data(0) = _
      verti_mid_line(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).verti_mid_line_no + 1 To _
     last_conditions.last_cond(1).verti_mid_line_no '49
     verti_mid_line(i%) = verti_mid_line(0)
Next i%
'**************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).v_line_value_no '49
     V_line_value(i%).data(0) = _
      V_line_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).v_line_value_no + 1 To _
     last_conditions.last_cond(1).v_line_value_no '49
     V_line_value(i%) = V_line_value(0)
Next i%
'***************************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).V_two_line_time_no '49
    V_two_line_time_value(i%).data(0) = _
     V_two_line_time_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).V_two_line_time_no + 1 To _
     last_conditions.last_cond(1).V_two_line_time_no '49
     V_two_line_time_value(i%) = V_two_line_time_value(0)
Next i%
'*******************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).squ_sum_no '50
     Squ_sum(i%).data(0) = _
      Squ_sum(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).squ_sum_no + 1 To _
     last_conditions.last_cond(1).squ_sum_no '50
     Squ_sum(i%) = Squ_sum(0)
Next i%
'********************
For i% = 1 To last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).string_value_no
    string_value(i%).data(0) = _
     string_value(i%).data(last_conditions_for_aid_no)
Next i%
For i% = last_conditions_for_aid(last_conditions_for_aid_no).last_cond(1).string_value_no + 1 To _
    last_conditions.last_cond(1).string_value_no
    string_value(i%) = string_value(0)
Next i%
from_aid_to_old0_error:
End Sub

Public Sub draw_aid_point(aid_point_no As Integer)
Dim i%, p%
p% = new_point(aid_point_no).data(0).poi(0)
If p% > 0 Then
If m_poi(p%).data(0).data0.visible = 0 Then
Call set_point_visible(p%, 1, False)
'Call draw_point(Draw_form, poi(p%), 0, display)
End If
For i% = 0 To 1
If new_point(aid_point_no).data(0).add_to_line(i%) > 0 Then
 Call C_display_picture.set_dot_line(0, 0, _
    new_point(aid_point_no).data(0).add_to_line(i%), p%)
End If
Next i%
End If
Select Case new_point(aid_point_no).data(0).cond.ty
Case angle3_value_
For i% = 0 To 2
 If angle3_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.angle(i%) > 0 Then
   Call draw_aid_angle( _
       angle3_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.angle(i%))
 End If
Next i%
Case line_value_
   Call C_display_picture.set_dot_line(line_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(0), _
         line_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(1), 0, 0)
Case two_line_value_
For i% = 0 To 1
    Call C_display_picture.set_dot_line( _
       two_line_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i%), _
         two_line_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i% + 1), 0, 0)
Next i%
Case line3_value_
For i% = 0 To 2
   Call C_display_picture.set_dot_line(line3_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i%), _
                     line3_value(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i% + 1), 0, 0)
Next i%
Case relation_
For i% = 0 To 1
   Call C_display_picture.set_dot_line(Drelation(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i%), _
                     Drelation(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i% + 1), 0, 0)
Next i%
Case dpoint_pair_
For i% = 0 To 3
  Call C_display_picture.set_dot_line(Ddpoint_pair(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i%), _
                     Ddpoint_pair(new_point(aid_point_no).data(0).cond.no).data(0).data0.poi(2 * i% + 1), 0, 0)
Next i%
Case general_string_
For i% = 0 To 3
 If general_string(new_point(aid_point_no).data(0).cond.no).data(0).item(i%) = 0 Then
   Call draw_aid_item( _
         general_string(new_point(aid_point_no).data(0).cond.no).data(0).item(i%))
 End If
Next i%
End Select
End Sub
Public Function add_mid_point(ByVal p1%, p2%, p3%, is_remove_aid_point As Byte) As Byte
Dim tl%, n%, tn_% 'is_remove_aid_point=0 =1, =2
Dim md As mid_point_data0_type
Dim temp_record As total_record_type
Dim c_data0 As condition_data_type
'On Error GoTo add_mid_point_error
If p2% = 0 Then '中点
If p1% = p3% Or p1% = p2% Then
Exit Function
ElseIf is_mid_point(p1%, p2%, p3%, 0, 0, 0, 0, 0, 1000, _
     0, 0, 0, 0, 0, 0, md, "", 0, 0, 0, temp_record.record_data.data0.condition_data) Then
       Exit Function
End If
If set_add_aid_point_for_mid_point(p1%, p2%, p3%) Then
If from_old_to_aid = 1 Then
   Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
If last_conditions.last_cond(1).point_no = 26 Then
 add_mid_point = 6
  Exit Function
End If
     Call set_point_name(last_conditions.last_cond(1).point_no, _
        next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
 p2% = last_conditions.last_cond(1).point_no
If m_poi(p1%).data(0).degree_for_reduce = 0 And m_poi(0).data(p3%).degree_for_reduce = 0 Then
   m_poi(p2%).data(0).degree_for_reduce = 0
ElseIf m_poi(p1%).data(0).degree_for_reduce > 0 And m_poi(0).data(p3%).degree_for_reduce > 0 Then
   If m_lin(line_number0(p1%, p2%, 0, 0)).data(0).degree = 0 Then
      m_poi(p2%).data(0).degree_for_reduce = 1
   Else
      m_poi(p2%).data(0).degree_for_reduce = 2
   End If
Else
   m_poi(p2%).data(0).degree_for_reduce = 1
End If
    t_coord.X = (m_poi(p1%).data(0).data0.coordinate.X + _
                            m_poi(p3%).data(0).data0.coordinate.X) / 2
    t_coord.Y = (m_poi(p1%).data(0).data0.coordinate.Y + _
                            m_poi(p3%).data(0).data0.coordinate.Y) / 2
    Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
       tl% = line_number0(p1%, p3%, 0, 0)
    Call add_point_to_line(last_conditions.last_cond(1).point_no, tl%, 0, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
  last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 'record0
    temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
     temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = tl%
n% = 0
new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
         LoadResString_(1365, "\\1\\" + m_poi(p1%).data(0).data0.name + m_poi(p3%).data(0).data0.name + _
                              "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
 add_mid_point = set_mid_point(p1%, last_conditions.last_cond(1).point_no, _
       p3%, 0, 0, 0, 0, 0, temp_record, n%, 0, 0, 0, 0)
        new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = midpoint_
        new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
   If add_mid_point > 1 Then
    Exit Function
   End If
    add_mid_point = set_New_point(last_conditions.last_cond(1).point_no, temp_record, tl%, _
     tn_%, 0, 0, 0, 0, 0, 1)
      If add_mid_point > 1 Then
       Exit Function
      End If
    If is_remove_aid_point < 2 Then
    add_mid_point = start_prove(0, 1, 1)
     If add_mid_point > 1 Then
      Exit Function
     End If
    End If
   'If new_result_from_add = False Then
   If is_remove_aid_point = 0 Then
    Call from_aid_to_old
   End If
   'Else
    ' new_result_from_add = False
   'End If
Else '已有中点
  Exit Function
End If
ElseIf p3% = 0 Then '延长成中点
If set_add_aid_point_for_mid_point(p1%, p2%, 2) Then
If from_old_to_aid = 1 Then
   Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
If last_conditions.last_cond(1).point_no = 26 Then
 add_mid_point = 6
  Exit Function
End If
     Call set_point_name(last_conditions.last_cond(1).point_no, _
        next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
p3% = last_conditions.last_cond(1).point_no
   t_coord.X = 2 * m_poi(p2%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X
   t_coord.Y = 2 * m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y
     Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
    tl% = line_number0(p1%, p2%, 0, 0)
     c_data0.condition_no = 0
     Call add_point_to_line(last_conditions.last_cond(1).point_no, tl%, 0, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
    If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
    End If
        last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
       new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = tl%
        ' poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
n% = 0
new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
   LoadResString_(1350, "\\1\\" + m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name + _
                       "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
                       "\\3\\" + m_poi(p2%).data(0).data0.name + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
                        "=" + m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name)
add_mid_point = set_mid_point(p1%, p2%, last_conditions.last_cond(1).point_no, _
       0, 0, 0, 0, 0, temp_record, n%, 0, 0, 0, 0)
 If add_mid_point > 1 Then
  Exit Function
 End If
 new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = midpoint_
 new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
 temp_record.record_data.data0.condition_data.condition_no = 1
  temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
   temp_record.record_data.data0.condition_data.condition(1).no = n%
add_mid_point = set_three_point_on_line(p1%, p2%, p3%, _
       temp_record, 0, 0, 1)
 If add_mid_point > 1 Then
 Exit Function
 End If
 temp_record.record_data.data0.condition_data.condition_no = 1
  temp_record.record_data.data0.condition_data.condition(1).ty = midpoint_
   temp_record.record_data.data0.condition_data.condition(1).no = n%
add_mid_point = set_New_point(last_conditions.last_cond(1).point_no, temp_record, tl%, _
    tn_%, 0, 0, 0, 0, 0, 1)
 If add_mid_point > 1 Then
  Exit Function
 End If
 If is_remove_aid_point < 2 Then
add_mid_point = start_prove(0, 1, 1)
If add_mid_point > 1 Then
 Exit Function
End If
End If
If is_remove_aid_point = 0 Then
add_mid_point_error:
Call from_aid_to_old
End If
End If
End If
End Function

Public Sub from_old_to_aid0()
Dim i%, j% '将数据复制到last_conditions_for_aid_no
'On Error GoTo from_old_to_aid0_error
t_condition.last_cond(0).point_no = 0
t_condition.last_cond(1).point_no = 0
For i% = 1 To last_conditions.last_cond(1).value_string_no
  Dvalue_string(i%).data(last_conditions_for_aid_no).value = _
               Dvalue_string(i%).data(0).value
  Dvalue_string(i%).data(last_conditions_for_aid_no).index(0) = _
               Dvalue_string(i%).data(0).index(0)
  Dvalue_string(i%).data(last_conditions_for_aid_no).index(1) = _
               Dvalue_string(i%).data(0).index(1)
  Call copy_factor_to_factor(Dvalue_string(i%).data(0).factor, _
               Dvalue_string(i%).data(last_conditions_for_aid_no).factor)
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data1_no
   aid_point_data1(i%).data(last_conditions_for_aid_no) = _
              aid_point_data1(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data2_no
  aid_point_data2(i%).data(last_conditions_for_aid_no) = _
              aid_point_data2(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data3_no
  aid_point_data3(i%).data(last_conditions_for_aid_no) = _
              aid_point_data3(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).same_three_lines_no
  same_three_lines(i%).data(last_conditions_for_aid_no) = _
               same_three_lines(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).length_of_polygon_no
  length_of_polygon(i%).data(last_conditions_for_aid_no) = _
   length_of_polygon(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line_from_two_point_no
  Dtwo_point_line(i%).data(last_conditions_for_aid_no) = _
   Dtwo_point_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).total_angle_no
  T_angle(i%).data(last_conditions_for_aid_no) = T_angle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).angle_no
 'ReDim Preserve angle(i%).data(last_conditions_for_aid_no) As angle_data_type
  angle(i%).data(last_conditions_for_aid_no) = angle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).item0_no
 'ReDim Preserve item0(i%).data(last_conditions_for_aid_no) As item0_data_type
   item0(i%).data(last_conditions_for_aid_no) = item0(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).dline1_no '3
 'ReDim Preserve Dline1(i%).data(last_conditions_for_aid_no) As Dline1_data_type
    Dline1(i%).data(last_conditions_for_aid_no) = _
     Dline1(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).distance_of_paral_line_no '3
 'ReDim Preserve Dline1(i%).data(last_conditions_for_aid_no) As Dline1_data_type
    Ddistance_of_paral_line(i%).data(last_conditions_for_aid_no) = _
     Ddistance_of_paral_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).distance_of_point_line_no '3
 'ReDim Preserve Dline1(i%).data(last_conditions_for_aid_no) As Dline1_data_type
    Ddistance_of_point_line(i%).data(last_conditions_for_aid_no) = _
     Ddistance_of_point_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no '3
' ReDim Preserve Ddpoint_pair(i%).data(last_conditions_for_aid_no) As point_pair_data_type
    Ddpoint_pair(i%).data(last_conditions_for_aid_no) = _
     Ddpoint_pair(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).dangle_no '3
 'ReDim Preserve Dangle(i%).data(last_conditions_for_aid_no) As Dangle_data_type
    Dangle(i%).data(last_conditions_for_aid_no) = _
     Dangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).triangle_no '4
' ReDim Preserve triangle(i%).data(last_conditions_for_aid_no) As triangle_data_type
   triangle(i%).data(last_conditions_for_aid_no) = _
     triangle(i%).data(0)
 Next i%
For i% = 1 To last_conditions.last_cond(1).rtriangle_no '4
' ReDim Preserve Rtriangle(i%).data(last_conditions_for_aid_no) As Rtriangle_data_type
    Rtriangle(i%).data(last_conditions_for_aid_no) = _
     Rtriangle(i%).data(0)
 Next i%
For i% = 1 To last_conditions.last_cond(1).area_relation_no '4
 'ReDim PreserveDarea_relation(i%).data(last_conditions_for_aid_no) As area_relation_data_type
    Darea_relation(i%).data(last_conditions_for_aid_no) = _
     Darea_relation(i%).data(0)
 Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_line_no '7
' ReDim Preserve mid_point_line(i%).data(last_conditions_for_aid_no) As mid_point_line_data_type
      mid_point_line(i%).data(last_conditions_for_aid_no) = _
       mid_point_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no '8
' ReDim Preserve Deline(i%).data(last_conditions_for_aid_no) As eline_data_type
     Deline(i%).data(last_conditions_for_aid_no) = _
     Deline(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).equal_3angle_no '8
' ReDim Preserve Deline(i%).data(last_conditions_for_aid_no) As eline_data_type
     equal_3angle(i%).data(last_conditions_for_aid_no) = _
     equal_3angle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).four_point_on_circle_no '9
  'ReDim Preserve four_point_on_circle(i%).data(last_conditions_for_aid_no) As four_point_on_circle_data_type
     four_point_on_circle(i%).data(last_conditions_for_aid_no) = _
     four_point_on_circle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_circle_no '9
  'ReDim Preserve four_point_on_circle(i%).data(last_conditions_for_aid_no) As four_point_on_circle_data_type
     three_point_on_circle(i%).data(last_conditions_for_aid_no) = _
        three_point_on_circle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).function_of_angle_no  '9
 ' ReDim Preserve function_of_angle(i%).data(last_conditions_for_aid_no) As function_of_angle_data_type
   function_of_angle(i%).data(last_conditions_for_aid_no) = _
     function_of_angle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no '12
 'ReDim Preserve Dmid_point(i%).data(last_conditions_for_aid_no) As mid_point_data_type
     Dmid_point(i%).data(last_conditions_for_aid_no) = _
      Dmid_point(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).paral_no '13
 'ReDim Preserve Dparal(i%).data(last_conditions_for_aid_no) As two_line_type
     Dparal(i%).data(last_conditions_for_aid_no) = _
      Dparal(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).parallelogram_no
'  ReDim Preserve Dparallelogram(i%).data(last_conditions_for_aid_no) As parallelogram_data_type
     Dparallelogram(i%).data(last_conditions_for_aid_no) = _
      Dparallelogram(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no '16
 'ReDim Preserve Drelation(i%).data(last_conditions_for_aid_no) As relation_data_type
     Drelation(i%).data(last_conditions_for_aid_no) = _
      Drelation(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).v_relation_no '16
 'ReDim Preserve v_Drelation(i%).data(last_conditions_for_aid_no) As v_relationn_data_type
     v_Drelation(i%).data(last_conditions_for_aid_no) = _
      v_Drelation(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_on_line_no '16
 'ReDim Preserve relation_on_line(i%).data(last_conditions_for_aid_no) As relation_on_line_type
     relation_on_line(i%).data(last_conditions_for_aid_no) = _
      relation_on_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_string_no '16
 'ReDim Preserve relation_string(i%).data(last_conditions_for_aid_no) As relation_string_data_type
     relation_string(i%).data(last_conditions_for_aid_no) = _
      relation_string(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).similar_triangle_no '17
 'ReDim Preserve Dsimilar_triangle(i%).data(last_conditions_for_aid_no) As two_triangle_type
     Dsimilar_triangle(i%).data(last_conditions_for_aid_no) = _
      Dsimilar_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).total_equal_triangle_no '18
 'ReDim Preserve Dtotal_equal_triangle(i%).data(last_conditions_for_aid_no) As two_triangle_type
     Dtotal_equal_triangle(i%).data(last_conditions_for_aid_no) = _
      Dtotal_equal_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_total_equal_triangle_no '18
 'ReDim Preserve Dtotal_equal_triangle(i%).data(last_conditions_for_aid_no) As two_triangle_type
     pseudo_total_equal_triangle(i%).data(last_conditions_for_aid_no) = _
      pseudo_total_equal_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_similar_triangle_no '18
 'ReDim Preserve Dtotal_equal_triangle(i%).data(last_conditions_for_aid_no) As two_triangle_type
     pseudo_similar_triangle(i%).data(last_conditions_for_aid_no) = _
      pseudo_similar_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_line_no '20
 '  ReDim Preserve three_point_on_line(i%).data(last_conditions_for_aid_no) As three_point_on_line_data_type
     three_point_on_line(i%).data(last_conditions_for_aid_no) = _
      three_point_on_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).tri_function_no '20
 '  ReDim Preserve three_point_on_line(i%).data(last_conditions_for_aid_no) As three_point_on_line_data_type
     tri_function(i%).data(last_conditions_for_aid_no) = _
      tri_function(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).two_point_conset_no '20
 '  ReDim Preserve three_point_on_line(i%).data(last_conditions_for_aid_no) As three_point_on_line_data_type
     two_point_conset(i%).data(last_conditions_for_aid_no) = _
      two_point_conset(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).two_line_value_no
   'ReDim Preserve two_line_value(i%).data(last_conditions_for_aid_no) As two_line_value_data_type
     two_line_value(i%).data(last_conditions_for_aid_no) = _
      two_line_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line3_value_no
'ReDim Preserve line3_value(i%).data(last_conditions_for_aid_no) As line3_value_data_type
    line3_value(i%).data(last_conditions_for_aid_no) = _
      line3_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).verti_no '24
'ReDim Preserve Dverti(i%).data(last_conditions_for_aid_no) As two_line_type
     Dverti(i%).data(last_conditions_for_aid_no) = _
      Dverti(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).arc_no
'ReDim Preserve Arc(i%).data(last_conditions_for_aid_no) As arc_data_type
arc(i%).data(last_conditions_for_aid_no) = arc(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).arc_value_no '25
'ReDim Preserve arc_value(i%).data(last_conditions_for_aid_no) As arc_value_data_type
    arc_value(i%).data(last_conditions_for_aid_no) = _
      arc_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).equal_arc_no '26
'ReDim Preserve equal_arc(i%).data(last_conditions_for_aid_no) As equal_arc_data_type
     equal_arc(i%).data(last_conditions_for_aid_no) = _
      equal_arc(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).ratio_of_two_arc_no '27
'ReDim Preserve ratio_of_two_arc(i%).data(last_conditions_for_aid_no) As ratio_of_two_arc_data_type
     ratio_of_two_arc(i%).data(last_conditions_for_aid_no) = _
      ratio_of_two_arc(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).angle_less_angle_no '28
'ReDim Preserve angle_less_angle(i%).data(last_conditions_for_aid_no) As angle_less_angle_data_type
     angle_less_angle(i%).data(last_conditions_for_aid_no) = _
      angle_less_angle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line_less_line_no '29
'ReDim Preserve line_less_line(i%).data(last_conditions_for_aid_no) As line_less_line_data_type
     line_less_line(i%).data(last_conditions_for_aid_no) = _
      line_less_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line_less_line2_no '30
'ReDim Preserve line_less_line2(i%).data(last_conditions_for_aid_no) As line_less_line2_data_type
     line_less_line2(i%).data(last_conditions_for_aid_no) = _
      line_less_line2(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line2_less_line2_no '31
'ReDim Preserve line2_less_line2(i%).data(last_conditions_for_aid_no) As line2_less_line2_data_type
     line2_less_line2(i%).data(last_conditions_for_aid_no) = _
      line2_less_line2(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).angle3_value_no '32
'ReDim Preserve angle3_value(i%).data(last_conditions_for_aid_no) As angle3_value_data_type
    angle3_value(i%).data(last_conditions_for_aid_no) = _
     angle3_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).line_value_no '33
 'ReDim Preserve line_value(i%).data(last_conditions_for_aid_no) As line_value_data_type
     line_value(i%).data(last_conditions_for_aid_no) = _
      line_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).tangent_line_no '34
'ReDim Preserve tangent_line(i%).data(last_conditions_for_aid_no) As tangent_line_data_type
     tangent_line(i%).data(last_conditions_for_aid_no) = _
      tangent_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).tangent_circle_no '34
'ReDim Preserve tangent_circle(i%).data(last_conditions_for_aid_no) As tangent_circle_data_type
     m_tangent_circle(i%).data(last_conditions_for_aid_no) = _
      m_tangent_circle(i%).data(0)
Next i%
'For i% = 1 To last_conditions.last_cond(1).equal_area_triangle_no '35
'ReDim Preserve equal_area_triangle(i%).data(last_conditions_for_aid_no) As equal_area_triangle_data_type
'    equal_area_triangle(i%).data(last_conditions_for_aid_no) = _
'     equal_area_triangle(i%).data(0)
'Next i%
For i% = 1 To last_conditions.last_cond(1).general_string_no '36
'ReDim Preserve general_string(i%).data(last_conditions_for_aid_no) As general_string_data_type
     general_string(i%).data(last_conditions_for_aid_no) = _
      general_string(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).general_angle_string_no  '37
'ReDim Preserve general_angle_string(i%).data(last_conditions_for_aid_no) As general_angle_string_data_type
     general_angle_string(i%).data(last_conditions_for_aid_no) = _
      general_angle_string(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).point_pair_for_similar_no '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     point_pair_for_similar(i%).data(last_conditions_for_aid_no) = _
      point_pair_for_similar(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_dpoint_pair_no  '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     pseudo_dpoint_pair(i%).data(last_conditions_for_aid_no) = _
      pseudo_dpoint_pair(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_midpoint_no  '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     pseudo_mid_point(i%).data(last_conditions_for_aid_no) = _
      pseudo_mid_point(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_relation_no  '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     pseudo_relation(i%).data(last_conditions_for_aid_no) = _
      pseudo_relation(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_line3_value_no  '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     pseudo_line3_value(i%).data(last_conditions_for_aid_no) = _
      pseudo_line3_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_eline_no  '38
'ReDim Preserve point_pair_for_similar(i%).data(last_conditions_for_aid_no) As point_pair_for_similar_data_type
     pseudo_eline(i%).data(last_conditions_for_aid_no) = _
      pseudo_eline(i%).data(0)
Next i%
'For i% = 1 To last_conditions.last_cond(1).equal_side_tixing_no '38
'     Dequal_side_tixing(i%).data(1last_conditions_for_aid_no) = _
'      Dequal_side_tixing(i%).data(0)
'Next i%
For i% = 1 To last_conditions.last_cond(1).equal_side_triangle_no '38
     equal_side_triangle(i%).data(last_conditions_for_aid_no) = _
      equal_side_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).equation_no '38
     equation(i%).data(1) = _
      equation(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).polygon4_no '39
'ReDim Preserve epolygon(i%).data(1) As epolygon_data_type
     Dpolygon4(i%).data(last_conditions_for_aid_no) = _
      Dpolygon4(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).epolygon_no '39
'ReDim Preserve epolygon(i%).data(1) As epolygon_data_type
     epolygon(i%).data(last_conditions_for_aid_no) = _
      epolygon(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).tixing_no '40
' ReDim Preserve Dtixing(i%).data(1) As tixing_data_type
    Dtixing(i%).data(last_conditions_for_aid_no) = _
      Dtixing(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).rhombus_no '41
'ReDim Preserve rhombus(i%).data(1) As rhombus_data_type
    rhombus(i%).data(last_conditions_for_aid_no) = _
      rhombus(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).squre_no '42
'ReDim Preserve Dlong_squre(i%).data(1) As long_squre_data_type
     Dsqure(i%).data(last_conditions_for_aid_no) = _
      Dsqure(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).long_squre_no '42
'ReDim Preserve Dlong_squre(i%).data(1) As long_squre_data_type
     Dlong_squre(i%).data(last_conditions_for_aid_no) = _
      Dlong_squre(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_element_no '43
'ReDim Preserve area_of_triangle(i%).data(1) As area_of_triangle_data_type
     area_of_element(i%).data(last_conditions_for_aid_no) = _
      area_of_element(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).two_area_of_element_value_no '43
'ReDim Preserve area_of_triangle(i%).data(1) As area_of_triangle_data_type
     two_area_of_element_value(i%).data(last_conditions_for_aid_no) = _
      two_area_of_element_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_circle_no '44
'ReDim Preserve area_of_circle(i%).data(1) As area_of_circle_data_type
     area_of_circle(i%).data(last_conditions_for_aid_no) = _
      area_of_circle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_fan_no '46
'ReDim Preserve Area_of_fan(i%).data(1) As area_of_fan_data_type
     Area_of_fan(i%).data(last_conditions_for_aid_no) = _
      Area_of_fan(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).sides_length_of_triangle_no '47
'ReDim Preserve Sides_length_of_triangle(i%).data(last_conditions_for_aid_no) As sides_length_of_triangle_data_type
     Sides_length_of_triangle(i%).data(last_conditions_for_aid_no) = _
      Sides_length_of_triangle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).sides_length_of_circle_no '48
'ReDim Preserve Sides_length_of_circle(i%).data(last_conditions_for_aid_no) As sides_length_of_circle_data_type
     Sides_length_of_circle(i%).data(last_conditions_for_aid_no) = _
      Sides_length_of_circle(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).verti_mid_line_no '49
'ReDim Preserve verti_mid_line(i%).data(last_conditions_for_aid_no) As verti_mid_line_data_type
     verti_mid_line(i%).data(last_conditions_for_aid_no) = _
      verti_mid_line(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).v_line_value_no  '49
'ReDim Preserve  V_line_value(i%).data(last_conditions_for_aid_no) As V_line_value_data_type
     V_line_value(i%).data(last_conditions_for_aid_no) = _
      V_line_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).V_two_line_time_no '49
'ReDim Preserve  V_two_line_time_value(i%).data(last_conditions_for_aid_no) As V_two_line_time_data_type
     V_two_line_time_value(i%).data(last_conditions_for_aid_no) = _
      V_two_line_time_value(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).squ_sum_no '50
'ReDim Preserve Squ_sum(i%).data(last_conditions_for_aid_no) As squ_sum_data_type
     Squ_sum(i%).data(last_conditions_for_aid_no) = _
      Squ_sum(i%).data(0)
Next i%
For i% = 1 To last_conditions.last_cond(1).string_value_no '50
'ReDim Preserve Squ_sum(i%).data(last_conditions_for_aid_no) As squ_sum_data_type
     string_value(i%).data(last_conditions_for_aid_no) = _
       string_value(i%).data(0)
Next i%
from_old_to_aid0_error:
End Sub
Public Function add_point_for_con_line3_value(ByVal no%) As Byte
add_point_for_con_line3_value = add_point_for_con_line3_value0(con_line3_value(no%).data(0).poi(0), _
        con_line3_value(no%).data(0).poi(1), con_line3_value(no%).data(0).poi(2), _
         con_line3_value(no%).data(0).poi(3), con_line3_value(no%).data(0).poi(4), _
          con_line3_value(no%).data(0).poi(5), con_line3_value(no%).data(0).para(0), _
           con_line3_value(no%).data(0).para(1), con_line3_value(no%).data(0).para(2), _
            con_line3_value(no%).data(0).value)
End Function
Public Function add_point_for_con_line3_value0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, _
                           ByVal pA1$, ByVal pA2$, ByVal pa3$, ByVal v$) As Byte
Dim tl%, n%, tn_%
Dim tp(5) As Integer
Dim r1&
Dim r2&
Dim temp_record As record_type
If v$ = "0" And pA1$ = "1" Then
If pA2$ = "1" And (pa3$ = "-1" Or pa3$ = "@1") Then
 tp(0) = p5% 'con_line3_value(no%).data(0).poi(4)
 tp(1) = p6% 'con_line3_value(no%).data(0).poi(5)
 tp(2) = p1% 'con_line3_value(no%).data(0).poi(0)
 tp(3) = p2% 'con_line3_value(no%).data(0).poi(1)
 tp(4) = p3% 'con_line3_value(no%).data(0).poi(2)
 tp(5) = p4% 'con_line3_value(no%).data(0).poi(3)
ElseIf (pA2$ = "-1" Or pA2$ = "@1") And (pa3$ = "-1" Or pa3$ = "@1") Then
 tp(0) = p1% 'con_line3_value(no%).data(0).poi(0)
 tp(1) = p2% 'con_line3_value(no%).data(0).poi(1)
 tp(2) = p3% 'con_line3_value(no%).data(0).poi(2)
 tp(3) = p4% 'con_line3_value(no%).data(0).poi(3)
 tp(4) = p5% 'con_line3_value(no%).data(0).poi(4)
 tp(5) = p6% 'con_line3_value(no%).data(0).poi(5)
ElseIf (pA2$ = "-1" Or pA2$ = "@1") And pa3$ = "1" Then
 tp(0) = p3% 'con_line3_value(no%).data(0).poi(2)
 tp(1) = p4% 'con_line3_value(no%).data(0).poi(3)
 tp(2) = p1% 'con_line3_value(no%).data(0).poi(0)
 tp(3) = p2% 'con_line3_value(no%).data(0).poi(1)
 tp(4) = p5% 'con_line3_value(no%).data(0).poi(4)
 tp(5) = p6% 'con_line3_value(no%).data(0).poi(5)
Else
 Exit Function
End If
add_point_for_con_line3_value0 = add_aid_point_for_eline(tp(0), tp(1), tp(2), tp(3), tp(2))
If add_point_for_con_line3_value0 > 1 Then
 Exit Function
End If
add_point_for_con_line3_value0 = add_aid_point_for_eline(tp(1), tp(0), tp(2), tp(3), tp(2))
If add_point_for_con_line3_value0 > 1 Then
 Exit Function
End If
add_point_for_con_line3_value0 = add_aid_point_for_eline(tp(0), tp(1), tp(4), tp(5), tp(4))
If add_point_for_con_line3_value0 > 1 Then
 Exit Function
End If
add_point_for_con_line3_value0 = add_aid_point_for_eline(tp(1), tp(0), tp(4), tp(5), tp(4))
End If

End Function

Public Function add_interset_point_line_line(ByVal l1%, ByVal l2%, t_p%, _
         ty As Integer, ty1 As Integer, is_no_initial As Integer, new_p%, c_data As condition_data_type) As Byte
Dim tn_(1) As Integer   'ty=0 消点 ty=1 不消点,ty=2 不作后继推理 ty1=1 其他辅助信息用
Dim tp(3) As Integer
Dim tl As add_point_for_two_line_type
Dim temp_record As total_record_type
Dim c_data0 As condition_data_type
is_no_initial = 0
  t_p% = is_line_line_intersect(l1%, l2%, 0, 0, False)
   If t_p% > 0 Then '已相交
     Exit Function
 End If
If set_add_aid_point_for_two_line(0, l1%, 0, l2%, 0) Then
   If from_old_to_aid = 1 Then
      Exit Function
   End If
    If last_conditions.last_cond(1).point_no = 26 Then
     add_interset_point_line_line = 6
      Exit Function
    End If
    'MDIForm1.Toolbar1.Buttons(21).Image = 33
     ' t_p% = last_conditions.last_cond(1).point_no
     '  Call set_point_name(last_conditions.last_cond(1).point_no, _
     '     next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
      'tp(0) = m_lin(l1%).data(0).data0.poi(0)
      'tp(1) = m_lin(l1%).data(0).data0.poi(1)
      'tp(2) = m_lin(l2%).data(0).data0.poi(0)
      'tp(3) = m_lin(l2%).data(0).data0.poi(1)
      t_p% = 0
      If inter_point_line_line(0, 0, 0, 0, _
               l1%, l2%, t_p%, pointapi0, True, False, False) > 0 Then
      Call set_two_point_line_for_line(l1%, temp_record.record_data)
      Call set_two_point_line_for_line(l2%, temp_record.record_data)
      Call arrange_data_for_new_point(l1%, l2%)
      If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
       ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
      End If
      last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 'record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      c_data.condition_no = 1
      c_data.condition(1).ty = new_point_
      c_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      new_p% = last_conditions.last_cond(1).new_point_no
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = t_p%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l1%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(1) = l2%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
        LoadResString_(1335, "\\1\\" + m_poi(tp(0)).data(0).data0.name + _
                            m_poi(tp(1)).data(0).data0.name + _
                            "\\2\\" + m_poi(tp(2)).data(0).data0.name + m_poi(tp(3)).data(0).data0.name + _
                            "\\3\\" + m_poi(t_p%).data(0).data0.name)
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
       add_interset_point_line_line = set_New_point(last_conditions.last_cond(1).point_no, temp_record, _
           l1%, l2%, tn_(0), tn_(1), 0, 0, 0, 1)
       If add_interset_point_line_line > 1 Or ty = 2 Then
        Exit Function
       End If
       If ty < 2 Then
       add_interset_point_line_line = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
       If add_interset_point_line_line > 1 Then
        Exit Function
       End If
       End If
       If ty1 = 1 Then
        Exit Function
       End If
add_interset_point_line_line_error:
     End If 'inter_point_line_line1
     If ty < 1 Then
     Call from_aid_to_old
     End If
      End If
End Function
Public Function add_point_for_paral(ByVal l1%, ByVal l2%) As Byte
Dim tl(2) As Integer
Dim i%, j%, k%, l%
Dim tn(1) As Integer
'On Error GoTo add_point_for_paral_error
tl(0) = l1%
tl(1) = l2%
For i% = 1 To m_lin(tl(0)).data(0).data0.in_point(0)
For j% = 1 To m_lin(tl(1)).data(0).data0.in_point(0)
tl(2) = line_number0(m_lin(tl(0)).data(0).data0.in_point(i%), m_lin(tl(1)).data(0).data0.in_point(j%), tn(0), tn(1))
add_point_for_paral = add_interset_point_three_line(tl(0), tl(1), tl(2))
If add_point_for_paral > 1 Then
 Exit Function
End If
add_point_for_paral = add_interset_point_three_line(tl(1), tl(2), tl(0))
If add_point_for_paral > 1 Then
 Exit Function
End If
add_point_for_paral = add_interset_point_three_line(tl(2), tl(0), tl(1))
If add_point_for_paral > 1 Then
 Exit Function
End If
Next j%
Next i%
add_point_for_paral_error:
End Function

Public Function add_interset_point_three_line(l1%, l2%, l3%) _
      As Byte
Dim i%, j%, l4%
Dim c_data0 As condition_data_type
For i% = 1 To m_lin(l1%).data(0).data0.in_point(0)
For j% = 1 To m_lin(l2%).data(0).data0.in_point(0)
l4% = line_number0(m_lin(l1%).data(0).data0.in_point(i%), m_lin(l2%).data(0).data0.in_point(j%), 0, 0)
If l4% > 0 And l3% <> l4% Then
'If is_dparal(l3%, l4%, 0, -1000, 0, 0, 0) = False Then
'If is_line_line_intersect(m_lin(l4%), m_lin(l3%), 0, 0) = 0 Then
add_interset_point_three_line = _
 add_interset_point_line_line(l3%, l4%, 0, 0, 0, 0, 0, c_data0)
If add_interset_point_three_line > 1 Then
 Exit Function
End If
'End If
'End If
End If
Next j%
Next i%
End Function
Public Function add_symmetry_point_of_line(p1%, l%, out_p1%, out_p2%, c_data0 As condition_data_type) As Byte
Dim t_coord(1)  As POINTAPI
Dim tp(1) As Integer
Dim temp_record As total_record_type
 add_symmetry_point_of_line = add_aid_point_for_verti0(p1%, l%, l%, out_p1%, c_data0, 1)
  If add_symmetry_point_of_line > 1 Then
     Exit Function
  End If
  tp(0) = last_conditions.last_cond(1).point_no
   out_p2% = 0
   add_symmetry_point_of_line = add_mid_point(p1%, tp(0), out_p2%, 2)
   If add_symmetry_point_of_line > 1 Then
     Exit Function
   End If
  '  m_poi(out_p2%).data(0).degree = m_poi(p1%).data(0).degree
     add_symmetry_point_of_line = start_prove(0, 1, 1)
End Function
Public Function add_point_for_eangle(ByVal A1%, ByVal A2%) As Byte
Dim i%, j%, t_p%, o%
Dim tp(1) As Integer
Dim tl(3) As Integer
Dim n(1) As Integer
Dim tA(1) As Integer
Dim c_data0 As condition_data_type
'On Error GoTo add_point_for_eangle_error
If last_conditions.last_cond(1).point_no = 26 Then
 add_point_for_eangle = 6
  Exit Function
End If
     tl(1) = angle(A1%).data(0).line_no(1) '角平分线
 For o% = 1 + last_conditions.last_cond(0).verti_no To last_conditions.last_cond(1).verti_no
     i% = Dverti(o%).data(0).record.data1.index.i(0)
      If tl(1) = Dverti(i%).data(0).line_no(0) Then
       tp(0) = is_line_line_intersect(Dverti(i%).data(0).line_no(1), _
                 angle(A1%).data(0).line_no(0), 0, 0, False)
       tp(1) = is_line_line_intersect(Dverti(i%).data(0).line_no(1), _
                 angle(A2%).data(0).line_no(1), 0, 0, False)
       If tp(0) > 0 And tp(1) = 0 Then
        add_point_for_eangle = add_interset_point_line_line(Dverti(i%).data(0).line_no(1), _
          angle(A2%).data(0).line_no(1), 0, 1, 0, 0, 0, c_data0)
          If add_point_for_eangle > 1 Then
           Exit Function
          End If
       ElseIf tp(0) = 0 And tp(1) > 0 Then
        add_point_for_eangle = add_interset_point_line_line(Dverti(i%).data(0).line_no(1), _
          angle(A1%).data(0).line_no(0), 0, 0, 0, 0, 0, c_data0)
          If add_point_for_eangle > 1 Then
           Exit Function
          End If
       End If
      ElseIf tl(1) = Dverti(i%).data(0).line_no(1) Then
       tp(0) = is_line_line_intersect(Dverti(i%).data(0).line_no(0), _
              angle(A1%).data(0).line_no(0), 0, 0, False)
       tp(1) = is_line_line_intersect(Dverti(i%).data(0).line_no(0), _
              angle(A2%).data(0).line_no(1), 0, 0, False)
       If tp(0) > 0 And tp(1) = 0 Then
        add_point_for_eangle = add_interset_point_line_line(Dverti(i%).data(0).line_no(0), _
          angle(A2%).data(0).line_no(1), 0, 1, 0, 0, 0, c_data0)
          If add_point_for_eangle > 1 Then
           Exit Function
          End If
       ElseIf tp(0) = 0 And tp(1) > 0 Then
        add_point_for_eangle = add_interset_point_line_line(Dverti(i%).data(0).line_no(0), _
          angle(A1%).data(0).line_no(0), 0, 1, 0, 0, 0, c_data0)
          If add_point_for_eangle > 1 Then
           Exit Function
          End If
       End If
      Else
      End If
     Next o%
     For i% = 1 To C_display_picture.m_circle.Count
      If m_Circ(i%).data(0).data0.visible > 0 Then
      If is_point_in_circle(i%, 0, angle(A1%).data(0).poi(1), _
                                    angle(A2%).data(0).poi(1), 0) Then
           If from_old_to_aid = 1 Then
              Exit Function
           End If
           add_point_for_eangle = add_interset_point_line_circle( _
            angle(A1%).data(0).poi(1), angle(A1%).data(0).poi(0), angle(A1%).data(0).line_no(0), _
             i%, 0, c_data0, 0)
           If add_point_for_eangle > 1 Then
            Exit Function
           End If
           add_point_for_eangle = add_interset_point_line_circle( _
            angle(A1%).data(0).poi(1), angle(A1%).data(0).poi(2), angle(A1%).data(0).line_no(1), _
             i%, 0, c_data0, 0)
           If add_point_for_eangle > 1 Then
            Exit Function
           End If
           add_point_for_eangle = add_interset_point_line_circle( _
            angle(A2%).data(0).poi(1), angle(A2%).data(0).poi(0), angle(A2%).data(0).line_no(0), _
             i%, 0, c_data0, 0)
           If add_point_for_eangle > 1 Then
            Exit Function
           End If
           add_point_for_eangle = add_interset_point_line_circle( _
            angle(A2%).data(0).poi(1), angle(A2%).data(0).poi(2), angle(A2%).data(0).line_no(1), _
             i%, 0, c_data0, 0)
           If add_point_for_eangle > 1 Then
            Exit Function
           End If
      End If
      End If
     Next i%
add_point_for_eangle_error:
End Function

Public Function add_point_for_eangle1(ByVal n1%, ByVal n2%, _
             ByVal l1%, ByVal n3%, ByVal n4%, ByVal l2%) As Byte
Dim j%, k%, tn_%, n%
Dim t_y As Byte
Dim r1 As Long
Dim r2 As Long
Dim tn(7)
Dim temp_record As total_record_type
Dim c_data As condition_data_type
'On Error GoTo add_point_for_eangle1_mark1
 tn(4) = n2%
  tn(5) = n4%
   tn(6) = n1%
    tn(7) = n3%
 If n1% <= n2% Then
     tn(0) = n1% + 1
      tn(1) = n2%
 ElseIf n1 > n2% Then
     tn(0) = n2%
      tn(1) = n1% - 1
 End If
 If n3% <= n4% Then
     tn(2) = n3% + 1
      tn(3) = n4%
 ElseIf n3% > n4% Then
     tn(2) = n4%
      tn(3) = n3% - 1
 End If
 For j% = tn(0) To tn(1)
  For k% = tn(2) To tn(3)
   c_data.condition_no = 0
   If is_equal_dline(m_lin(l1%).data(0).data0.in_point(tn(6)), m_lin(l1%).data(0).data0.in_point(j%), _
        m_lin(l2%).data(0).data0.in_point(tn(7)), m_lin(l2%).data(0).data0.in_point(k%), _
         tn(6), j%, tn(7), k%, l1%, l2%, 0, -1000, 0, 0, 0, eline_data0, _
          0, 0, 0, "", c_data) Then
            GoTo add_point_for_eangle1_mark1
   End If
  Next k%
  add_point_for_eangle1 = add_aid_point_for_eline(m_lin(l1%).data(0).data0.in_point(tn(6)), _
       m_lin(l1%).data(0).data0.in_point(j%), m_lin(l1%).data(0).data0.in_point(tn(6)), _
         m_lin(l2%).data(0).data0.in_point(tn(5)), m_lin(l1%).data(0).data0.in_point(tn(6)))
        If add_point_for_eangle1 > 1 Then
         Exit Function
        End If
add_point_for_eangle1_mark1:
Next j%
End Function

Public Function set_total_triangle_from_eangle_(ByVal A1%, _
                      ByVal A2%) As Byte
'等角,邻边,对边
Dim j%, k%
Dim triA(1) As temp_triangle_type
Dim c_data As condition_data_type
'On Error GoTo set_total_angle_triangle_from_eangle_error
triA(0).last_T = 0
triA(1).last_T = 0
Call set_temp_triangle_from_angle(A1%, 0, _
  triA(0), True)
Call set_temp_triangle_from_angle(A2%, 0, _
    triA(1), False)
For j% = 1 To triA(0).last_T
 For k% = 1 To triA(1).last_T
'**************************************************************
    If is_total_equal_Triangle(triA(0).data(j%).no, triA(1).data(k%).no, _
     triA(0).data(j%).direction, triA(1).data(k%).direction, 0, -1000, 0, _
       0, two_triangle0, record_0, 0) Then
      GoTo set_total_equal_triangle_from_eangle0_mark2
    End If
   c_data.condition_no = 0
     If is_equal_dline(triA(0).data(j%).poi(0), triA(0).data(j%).poi(1), _
        triA(1).data(k%).poi(0), triA(1).data(k%).poi(1), _
           0, 0, 0, 0, 0, 0, 0, -1000, 0, 0, 0, eline_data0, _
             0, 0, 0, "", c_data) Then '一邻边等
      If is_equal_dline(triA(0).data(j%).poi(1), triA(0).data(j%).poi(2), _
        triA(1).data(k%).poi(1), triA(1).data(k%).poi(2), _
           0, 0, 0, 0, 0, 0, 0, -1000, 0, 0, 0, eline_data0, _
            0, 0, 0, "", c_data) = False Then
       set_total_triangle_from_eangle_ = add_aid_point_for_eline( _
          triA(0).data(j%).poi(0), triA(0).data(j%).poi(2), _
           triA(1).data(k%).poi(0), triA(1).data(k%).poi(2), triA(1).data(k%).poi(0))
       If set_total_triangle_from_eangle_ > 1 Then
        Exit Function
       End If
       set_total_triangle_from_eangle_ = add_aid_point_for_eline( _
          triA(1).data(k%).poi(0), triA(1).data(k%).poi(2), _
           triA(0).data(j%).poi(0), triA(0).data(j%).poi(2), triA(0).data(j%).poi(0))
       If set_total_triangle_from_eangle_ > 1 Then
        Exit Function
       End If
      End If
    ElseIf is_equal_dline(triA(0).data(j%).poi(0), triA(0).data(j%).poi(2), _
        triA(1).data(k%).poi(0), triA(1).data(k%).poi(2), _
           0, 0, 0, 0, 0, 0, 0, -1000, 0, 0, 0, eline_data0, _
             0, 0, 0, "", c_data) Then '一邻边等
      If is_equal_dline(triA(0).data(j%).poi(1), triA(0).data(j%).poi(2), _
        triA(1).data(k%).poi(1), triA(1).data(k%).poi(2), _
           0, 0, 0, 0, 0, 0, 0, -1000, 0, 0, 0, eline_data0, _
             0, 0, 0, "", c_data) Then
       set_total_triangle_from_eangle_ = add_aid_point_for_eline( _
          triA(0).data(j%).poi(0), triA(0).data(j%).poi(1), _
           triA(1).data(k%).poi(0), triA(1).data(k%).poi(1), triA(1).data(k%).poi(0))
       If set_total_triangle_from_eangle_ > 1 Then
        Exit Function
       End If
       set_total_triangle_from_eangle_ = add_aid_point_for_eline( _
          triA(1).data(k%).poi(0), triA(1).data(k%).poi(1), _
           triA(0).data(j%).poi(0), triA(0).data(j%).poi(1), triA(0).data(j%).poi(0))
       If set_total_triangle_from_eangle_ > 1 Then
        Exit Function
       End If
        End If
    End If
set_total_equal_triangle_from_eangle0_mark2:
   Next k%
  Next j%
set_total_angle_triangle_from_eangle_error:
End Function

Public Function add_aid_point_for_eline(ByVal p1%, ByVal p2%, _
      ByVal p3%, ByVal p4%, ByVal p5%) As Byte
       '在直线p3%p4%上取点p使得p5%p=p1%p2%
Dim r1!
Dim r2!
Dim tl%, i%, j%
Dim con_ty As Byte
Dim p As POINTAPI
Dim tn%, n%, tp%
Dim t_n(1) As Integer
Dim el_data0 As eline_data0_type
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
Dim c_data0 As condition_data_type
'On Error GoTo add_aid_point_for_eline_mark1
If p3% = p4% Then
 Exit Function
ElseIf p5% = p3% Then
 '起点相同,且p1%p2%=p3%p4%
 If is_equal_dline(p1%, p2%, p3%, p4%, 0, 0, 0, 0, 0, 0, 0, -1000, 0, 0, 0, el_data0, 0, 0, 0, "", c_data0) Then
  Exit Function
 End If
End If
tl% = line_number0(p3%, p4%, t_n(0), t_n(1))
If t_n(0) < t_n(1) Then
    For j% = t_n(0) + 1 To m_lin(tl%).data(0).data0.in_point(0)
     If m_lin(tl%).data(0).data0.in_point(j%) = p5% Then
     t_n(0) = j% + 1
     End If
    Next j%
    t_n(1) = m_lin(tl%).data(0).data0.in_point(0)
Else
    For j% = m_lin(tl%).data(0).data0.in_point(0) To t_n(0) - 1
     If m_lin(tl%).data(0).data0.in_point(j%) = p5% Then
     t_n(1) = j% - 1
     End If
    Next j%
    t_n(0) = 1
End If
For i% = t_n(0) To t_n(1)
If is_equal_dline(p5%, _
     m_lin(tl%).data(0).data0.in_point(i%), p1%, p2%, 0, 0, 0, 0, 0, 0, 0, -1000, _
       0, 0, 0, el_data0, 0, 0, 0, "", c_data0) Then
        Exit Function
End If
Next i%
If from_old_to_aid = 1 Then
   Exit Function
End If
If p1% < p2% Then
el.poi(0) = p1%
el.poi(1) = p2%
Else
el.poi(0) = p2%
el.poi(1) = p1%
End If
el.poi(2) = p3%
el.line_no = line_number0(p3%, p4%, t_n(0), t_n(1))
If t_n(0) < t_n(1) Then
el.te = 0
Else
el.te = 1
End If
If search_for_add_aid_point_for_eline(el, 0) Then
 Exit Function
End If
'p1%p2%的长
r1! = sqr((m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
'p3%p4%的长
r2! = sqr((m_poi(p3%).data(0).data0.coordinate.X - m_poi(p4%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p4%).data(0).data0.coordinate.Y) ^ 2)
'Call from_old_to_aid
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_eline = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
Call get_new_char(last_conditions.last_cond(1).point_no)
'p3%p4%上取一点使得p5%,p=p1%,p2%
 p.X = m_poi(p5%).data(0).data0.coordinate.X + _
            (m_poi(p4%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) * r1! / r2!
 p.Y = m_poi(p5%).data(0).data0.coordinate.Y + _
            (m_poi(p4%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) * r1! / r2!
  tp% = read_point(p, 0)
  If tp% > 0 And is_equal_dline(p1%, p2%, p5%, tp%, 0, 0, 0, 0, 0, 0, 0, _
          -1000, 0, 0, 0, eline_data0, 0, 0, 0, "", temp_record.record_data.data0.condition_data) Then
   GoTo add_aid_point_for_eline_mark1
  End If
  Call set_point_coordinate(last_conditions.last_cond(1).point_no, p, False)
  tl% = line_number0(p3%, p4%, 0, 0)
    c_data0.condition_no = 0
    Call add_point_to_line(last_conditions.last_cond(1).point_no, tl%, tn%, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
   If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
   End If
       last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
       'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
       new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = tl%
        'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
             "\\1\\" + m_poi(p3%).data(0).data0.name + m_poi(p4%).data(0).data0.name + _
             "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
             "\\3\\" + m_poi(p5%).data(0).data0.name + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + "=" + _
             m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name)
   Call set_equal_dline(p5%, last_conditions.last_cond(1).point_no, p1%, p2%, _
           0, 0, 0, 0, 0, 0, 0, temp_record, n%, con_ty, t_n(0), t_n(1), 0, False)
           new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = eline_
           new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
      temp_record.record_data.data0.condition_data.condition_no = 0
       Call add_conditions_to_record(con_ty, n%, t_n(0), t_n(1), _
                  temp_record.record_data.data0.condition_data)
 add_aid_point_for_eline = set_New_point(last_conditions.last_cond(1).point_no, temp_record, tl%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_eline > 1 Then
    Exit Function
   End If
 add_aid_point_for_eline = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_eline > 1 Then
    Exit Function
   End If
add_aid_point_for_eline_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
 'Else
 ' new_result_from_add = False
' End If
End Function

Public Function add_interset_point_line_circle(ByVal p%, _
      ByVal p1%, ByVal l%, ByVal c%, out_p%, c_data0 As condition_data_type, _
        is_remove_new_point As Byte) As Byte
Dim tp As Integer
Dim tn%
Dim t_p(1) As Integer
Dim ty As Integer
Dim p_coord As POINTAPI
Dim temp_record As total_record_type
Dim i%
'On Error GoTo add_interset_point_line_circle_error
If from_old_to_aid = 1 Then
   Exit Function
End If
  tn% = inter_point_line_circle0(m_lin(l%).data(0).data0, m_Circ(c%).data(0).data0, t_p(0), t_p(1))
   If tn% = 2 Then
    Exit Function
   ElseIf tn% = 1 Then
    If t_p(0) > 0 And t_p(0) <> p% And t_p(0) <> p1% Then
     add_interset_point_line_circle = add_interset_point_line_circle(t_p(0), m_Circ(c%).data(0).data0.center, _
          l%, c%, out_p%, c_data0, is_remove_new_point)
           Exit Function
    ElseIf t_p(1) > 0 And t_p(1) <> p% And t_p(1) <> p1% Then
     add_interset_point_line_circle = add_interset_point_line_circle(t_p(1), m_Circ(c%).data(0).data0.center, _
           l%, c%, out_p%, c_data0, is_remove_new_point)
           Exit Function
    End If
   ty = compare_two_point(m_poi(p%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate, 0, 0, 0)
    Call inter_point_line_circle2(l%, p%, c%, p_coord, 0)
      If ty = compare_two_point(p_coord, m_poi(t_p(0)).data(0).data0.coordinate, 0, 0, 0) Then
       Exit Function
      End If
   End If
ty = compare_two_point(m_poi(p%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate, 0, 0, 0)
  Call inter_point_line_circle2(l%, p%, c%, p_coord, 0)
  If ty = compare_two_point(m_poi(p%).data(0).data0.coordinate, p_coord, 0, 0, 0) Then
    If last_conditions.last_cond(1).point_no = 26 Then
     add_interset_point_line_circle = 6
      Exit Function
    End If
     last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
       Call set_point_name(last_conditions.last_cond(1).point_no, _
          next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
        tp = last_conditions.last_cond(1).point_no
         out_p% = tp%
         Call set_point_coordinate(tp, p_coord, False)
     record_0.data0.condition_data.condition_no = 0 ' record0
    Call add_point_to_line(last_conditions.last_cond(1).point_no, l%, tn%, no_display, False, 0)
     Call set_two_point_line_for_line(l%, temp_record.record_data)
      Call arrange_data_for_new_point(l%, 0)
       Call add_point_to_m_circle(last_conditions.last_cond(1).point_no, c%, record0, 0)
      If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
      End If
        last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = tp
      'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l%
      c_data0.condition_no = 1
      c_data0.condition(1).no = last_conditions.last_cond(1).new_point_no
      c_data0.condition(1).ty = new_point_
      '*new_point(last_conditions.last_cond(1).new_point_no).data(0).record.data0.condition_data.condition_no = 1
      '*new_point(last_conditions.last_cond(1).new_point_no).data(0).record.data0.condition_data.condition(1).no = _
             last_conditions.last_cond(1).new_point_no
      '*new_point(last_conditions.last_cond(1).new_point_no).data(0).record.data0.condition_data.condition(1).ty = _
              add_condition_
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
        LoadResString_(1340, "\\1\\" + m_poi(p%).data(0).data0.name + _
                            m_poi(p1%).data(0).data0.name + _
                            "\\2\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + "(" + _
                            m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                            "\\3\\" + m_poi(tp).data(0).data0.name)
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
         'temp_record.record_data.data1.aid_condition = 0 ' last_conditions.last_cond(1).new_point_no
      add_interset_point_line_circle = set_New_point(tp, temp_record, l%, 0, tn%, 0, c%, 0, _
             0, 1)
       If add_interset_point_line_circle > 1 Then
        Exit Function
       End If
       If is_remove_new_point < 2 Then
       add_interset_point_line_circle = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
       If add_interset_point_line_circle > 1 Then
        Exit Function
       End If
       End If
add_interset_point_line_circle_error:
       If is_remove_new_point = 0 Then
       Call from_aid_to_old
       End If
 End If
End Function
Public Function add_aid_point_for_paral_or_verti(ByVal p%, paral_or_verti_ As Integer, _
          ByVal l1%, ByVal l2%, _
           ByVal remove_add_point As Byte) As Byte
Dim tp%
Dim p_coord As POINTAPI
Dim c_data0 As condition_data_type

add_aid_point_for_paral_or_verti = inter_point_line_line3(p%, paral_or_verti_, l1%, _
                              m_lin(l2%).data(0).data0.poi(0), True, l2%, p_coord, tp%, False, c_data0, True)
add_aid_point_for_verti_error:
End Function
Public Function set_add_point_data(ByVal p%, ByVal l1%, ByVal l2%, _
                                    ByVal c1%, ByVal c2%, _
                                     ByVal remove_add_point As Byte) As Byte
          '过p%点垂直l1%的直线交直线l2%
Dim n%
Dim temp_record As total_record_type
'*************************************************************************
'设置新的辅助点
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
   ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
  '*********************************************************************
  'last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
       temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
         temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
          temp_record.record_data.data0.theorem_no = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l1%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(1) = l2%
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_circle(0) = c1%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_circle(1) = c2%
      set_add_point_data = set_New_point(last_conditions.last_cond(1).point_no, temp_record, l1%, l2%, _
            0, 0, c1%, c2%, 0, 1)
       If set_add_point_data > 1 Then
        Exit Function
       End If
      If remove_add_point < 2 Then
      set_add_point_data = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
       If set_add_point_data > 1 Then
         Exit Function
       End If
      ElseIf remove_add_point = 0 Then
        Call from_aid_to_old
      End If
add_aid_point_for_verti0_error:
End Function
Public Function add_aid_point_for_verti0(ByVal p%, _
         ByVal l1%, ByVal l2%, out_p%, c_data0 As condition_data_type, _
          ByVal remove_add_point As Byte) As Byte
          '过p%点垂直l1%的直线交直线l2%
Dim tl%, n%, i%, l3%
Dim tn(1) As Integer
Dim p_coord As POINTAPI
Dim tp(3) As Integer
Dim temp_record As total_record_type
For i% = 1 To m_lin(l2%).data(0).data0.in_point(0)
 If is_dverti(l1%, line_number0(p%, m_lin(l2%).data(0).data0.in_point(i%), 0, 0), 0, _
      -1000, 0, 0, 0, 0) Then
        Exit Function
 End If
Next i%
If is_point_in_line3(p%, m_lin(l2%).data(0).data0, 0) Then
    Exit Function
ElseIf is_point_in_verti_line(p%, l2%, 0, l3%) Then
    add_aid_point_for_verti0 = add_interset_point_line_line(l3%, l2%, 0, 0, 0, 0, 0, c_data0)
    Exit Function
End If
If set_add_aid_point_for_two_line(p%, l1%, 0, l2%, 0) Then
tp(0) = m_lin(l1%).data(0).data0.poi(0)
tp(1) = m_lin(l1%).data(0).data0.poi(1)
tp(2) = m_lin(l2%).data(0).data0.poi(0)
tp(3) = m_lin(l2%).data(0).data0.poi(1)
If inter_point_line_line3(p%, verti_, l1%, _
         m_lin(l2%).data(0).data0.poi(0), paral_, l2%, p_coord, 0, False, _
          c_data0, True) Then
'On Error GoTo add_aid_point_for_verti0_error
If from_old_to_aid = 1 Then
   Exit Function
End If
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_verti0 = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
out_p% = last_conditions.last_cond(1).point_no
Call set_point_coordinate(last_conditions.last_cond(1).point_no, p_coord, False)
 Call set_point_name(last_conditions.last_cond(1).point_no, _
  next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
   'if (m_poi(p%).data(0).degree=2 or m_poi(p%).data(0).degree=0) and _
        m_lin(l1%).data(0).degree=2  and m_lin(l2%).data(0).degree=2 _
   'm_poi(out_p%).data (0).
  tl% = line_number0(last_conditions.last_cond(1).point_no%, p%, tn(1), 0)
  record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(last_conditions.last_cond(1).point_no%, l2%, _
       tn(0), no_display, False, 0, temp_record)
     Call set_two_point_line_for_line(l2%, temp_record.record_data)
      Call arrange_data_for_new_point(l2%, 0)
      
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
   ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
      last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
       temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
         temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
         temp_record.record_data.data0.theorem_no = 0
   new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l2%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(1) = tl%
          'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
          'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1330, _
          "\\1\\" + m_poi(p%).data(0).data0.name + _
          "\\2\\" + m_poi(tp(0)).data(0).data0.name + m_poi(tp(1)).data(0).data0.name + _
          "\\3\\" + m_poi(tp(2)).data(0).data0.name + m_poi(tp(3)).data(0).data0.name + _
          "\\4\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
     add_aid_point_for_verti0 = set_dverti(l1%, tl%, temp_record, n%, 0, False)
       If add_aid_point_for_verti0 > 1 Then
        Exit Function
       End If
       temp_record.record_data.data0.condition_data.condition_no = 1
        temp_record.record_data.data0.condition_data.condition(1).ty = verti_
         temp_record.record_data.data0.condition_data.condition(1).no = n%
       c_data0 = temp_record.record_data.data0.condition_data
      new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = verti_
       new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
      add_aid_point_for_verti0 = set_New_point(last_conditions.last_cond(1).point_no, temp_record, l2%, tl%, _
            tn(0), tn(1), 0, 0, 0, 1)
       If add_aid_point_for_verti0 > 1 Then
        Exit Function
       End If
      If remove_add_point < 2 Then
      add_aid_point_for_verti0 = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
       If add_aid_point_for_verti0 > 1 Then
         Exit Function
       End If
      End If
add_aid_point_for_verti0_error:
      If remove_add_point = 0 Then
        Call from_aid_to_old
      End If
       'Else
       ' new_result_from_add = False
      ' End If
     End If
     End If
    End Function
Public Function add_aid_point_for_circle(ByVal i%, no_reduce As Byte) As Byte
Dim j%, k%, l%, m%, n%, o%, t_p%, tn_%, tl%
Dim r&, s&
Dim t!
Dim tem_p(1) As Integer
Dim t_n(1) As Integer
Dim temp_record As total_record_type

'On Error GoTo add_aid_point_for_circle_error
If m_Circ(i%).data(0).data0.visible = 0 Then
 Exit Function
End If
 If m_Circ(i%).data(0).data0.center > 0 Then '圆i%是有心圆
For m% = 1 To last_conditions.last_cond(1).verti_no
 If Dverti(m%).data(0).inter_poi > 0 Then '相互垂直的直线有交点
  For o% = 0 To 1
   If is_point_in_line3(m_Circ(i%).data(0).data0.center, m_lin(Dverti(m%).data(0).line_no(o%)).data(0).data0, 0) Then
    '圆心在一条垂线上
    l% = Dverti(m%).data(0).line_no((o% + 1) Mod 2)
           add_aid_point_for_circle = add_aid_point_for_circle1(i%, l%) '
    If add_aid_point_for_circle > 1 Then
     Exit Function
    Else
     GoTo add_aid_point_for_circle_mark20
    End If
   End If
  Next o%
 End If
Next m%
End If
add_aid_point_for_circle_mark20:
 For j% = 1 To m_Circ(i%).data(0).data0.in_point(0)
  For k% = 1 To m_Circ(i%).data(0).data0.in_point(0)
  record_0.data0.condition_data.condition_no = 0 ' record0
   l% = line_number0(m_Circ(i%).data(0).data0.in_point(j%), m_Circ(i%).data(0).data0.center, 0, 0)  '过j%的半径
    If k% <> j% And is_three_point_on_line(m_Circ(i%).data(0).data0.in_point(j%), _
      m_Circ(i%).data(0).data0.in_point(k%), m_Circ(i%).data(0).data0.center, 0, -1000, 0, 0, _
        0, 0, 0) Then    'j%,k%是直径
        GoTo add_aid_point_for_circle_mark10 '是直径
    'ElseIf lin(l%).data(0).data0.visible > 0 Then
     '   GoTo add_aid_point_for_circle_mark10
    ElseIf set_add_aid_point_for_line_circle(0, l%, i%, 1, 0) = False Then
        GoTo add_aid_point_for_circle_mark10
    End If
    Next k%
   'If m_poi(m_Circ(i%).data(0).data0.center).data(0).no_reduce = 0 Then
    If m_lin(l%).data(0).data0.visible > 0 Then
      If from_old_to_aid = 1 Then
         Exit Function
      End If
     If last_conditions.last_cond(1).point_no = 26 Then
      add_aid_point_for_circle = 6
       Exit Function
     End If
      last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
     ' MDIForm1.Toolbar1.Buttons(21).Image = 33
       Call set_point_name(last_conditions.last_cond(1).point_no, _
        next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
        t_p% = last_conditions.last_cond(1).point_no
       '可见半径延长为直径
       t_coord.X = 2 * m_Circ(i%).data(0).data0.c_coord.X - _
           m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).data0.coordinate.X
       t_coord.Y = 2 * m_Circ(i%).data(0).data0.c_coord.Y - _
           m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).data0.coordinate.Y
       Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
        record_0.data0.condition_data.condition_no = 0 ' record0
    Call add_point_to_line(last_conditions.last_cond(1).point_no, l%, tn_%, no_display, False, 0, temp_record.record_data)
     Call set_two_point_line_for_line(l%, temp_record.record_data)
      Call arrange_data_for_new_point(l%, 0)
       Call add_point_to_m_circle(last_conditions.last_cond(1).point_no, i%, temp_record, 255)
    If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
    End If
        last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = t_p%
      'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
        LoadResString_(1340, "\\1\\" + m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).data0.name + _
                            m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.name + _
                            "\\2\\" + m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.name + "(" + _
                            m_poi(m_Circ(i%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                            "\\3\\" + m_poi(t_p%).data(0).data0.name)
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
'         temp_record.record_data.data1.aid_condition = 0 ' last_conditions.last_cond(1).new_point_no
      add_aid_point_for_circle = set_New_point(t_p%, temp_record, l%, 0, tn_%, 0, i%, 0, _
             0, 1)
       If add_aid_point_for_circle > 1 Then
        Exit Function
       End If
       add_aid_point_for_circle = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
       If add_aid_point_for_circle > 1 Then
        Exit Function
       Else
        'If new_result_from_add = False Then
         Call from_aid_to_old
        'Else
        ' new_result_from_add = False
        'End If
      End If
    End If
 'End If
add_aid_point_for_circle_mark10:
  For k% = 1 To m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(0) '过第一点的直线
   If m_lin(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)).data(0).data0.visible > 0 Then
    If is_point_in_line3(m_Circ(i%).data(0).data0.center, _
          m_lin(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)).data(0).data0, 0) = False Then '不过圆心
     For l% = 1 To m_Circ(i%).data(0).data0.in_point(0) '第三点
      If m_Circ(i%).data(0).data0.in_point(l%) <> m_Circ(i%).data(0).data0.in_point(j%) Then '与第一点不同
       If is_point_in_line3(m_Circ(i%).data(0).data0.in_point(l%), _
          m_lin(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)).data(0).data0, 0) Then '不在直线上
           GoTo add_aid_point_for_circle_mark11
       End If
      End If
     Next l%
     record_0.data0.condition_data.condition_no = 0 ' record0
     If is_tangent_line(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%), _
          m_Circ(i%).data(0).data0.in_point(j%), depend_condition(circle_, i%), _
            m_Circ(i%).data(0).data0.in_point(j%), depend_condition(0, 0), tangent_line_data0, 0, _
             0, 0, record_0) Then   '判断是否切线
            GoTo add_aid_point_for_circle_mark11 '是
     End If
    If from_old_to_aid = 1 Then
       Exit Function
    End If
    If last_conditions.last_cond(1).point_no = 26 Then
     add_aid_point_for_circle = 6
      Exit Function
    End If
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'    MDIForm1.Toolbar1.Buttons(21).Image = 33
      Call set_point_name(last_conditions.last_cond(1).point_no, _
        next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
        t_p% = last_conditions.last_cond(1).point_no
 '  tem_p(0) = m_lin(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)).data(0).data0.poi(0)
 '  tem_p(1) = m_lin(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)).data(0).data0.poi(1)
Call inter_point_line_circle2(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%), _
          m_Circ(i%).data(0).data0.in_point(j%), _
       i%, t_coord, last_conditions.last_cond(1).point_no)
   If read_point(m_poi(last_conditions.last_cond(1).point_no).data(0).data0.coordinate, 0) _
           = m_Circ(i%).data(0).data0.in_point(j%) Then
         GoTo add_aid_point_for_circle_mark13
   End If
   record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(last_conditions.last_cond(1).point_no, _
       m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%), tn_%, no_display, False, 0, temp_record.record_data)
     Call set_two_point_line_for_line(m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%), temp_record.record_data)
      Call arrange_data_for_new_point(m_Circ(i%).data(0).data0.in_point(k%), 0)
      Call add_point_to_m_circle(last_conditions.last_cond(1).point_no, i%, temp_record, 255)
      If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
      End If
        last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 'record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = t_p%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = _
             m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%)
      If m_Circ(i%).data(0).data0.center > 0 Then
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
        LoadResString_(1340, "\\1\\" + m_poi(tem_p(0)).data(0).data0.name + _
                            m_poi(tem_p(1)).data(0).data0.name + _
                            "\\2\\" + m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.name + "(" + _
                            m_poi(m_Circ(i%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                            "\\3\\" + m_poi(t_p%).data(0).data0.name)
      Else
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
        LoadResString_(1340, "\\1\\" + m_poi(tem_p(0)).data(0).data0.name + _
                            m_poi(tem_p(1)).data(0).data0.name + _
                            "\\2\\" + m_poi(m_Circ(i%).data(0).data0.in_point(1)).data(0).data0.name + _
                            m_poi(m_Circ(i%).data(0).data0.in_point(2)).data(0).data0.name + _
                            m_poi(m_Circ(i%).data(0).data0.in_point(3)).data(0).data0.name + _
                            "\\3\\" + m_poi(t_p%).data(0).data0.name)
      End If
      temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      add_aid_point_for_circle = set_New_point(t_p%, temp_record, _
         m_poi(m_Circ(i%).data(0).data0.in_point(j%)).data(0).in_line(k%), 0, tn_%, 0, i%, 0, _
           0, 1)
       If add_aid_point_for_circle > 1 Then
        Exit Function
       End If
       add_aid_point_for_circle = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
       If add_aid_point_for_circle > 1 Then
        Exit Function
       Else
add_aid_point_for_circle_mark13:
       'If new_result_from_add = False Then
        Call from_aid_to_old
        'Else
        'new_result_from_add = False
        'End If
       End If
    End If
   End If
add_aid_point_for_circle_mark11:
  Next k%
  For k% = 1 To j% - 1
   l% = line_number0(m_Circ(i%).data(0).data0.in_point(j%), m_Circ(i%).data(0).data0.in_point(k%), _
            t_n(0), t_n(1))
     If t_n(0) > t_n(1) Then
       Call exchange_two_integer(t_n(0), t_n(1))
     End If
     For m% = t_n(0) + 1 To t_n(1) - 1
      For o% = m% To t_n(1) - 1
       If is_equal_dline(m_lin(l%).data(0).data0.in_point(t_n(0)), _
        m_lin(l%).data(0).data0.in_point(m%), m_lin(l%).data(0).data0.in_point(o%), _
         m_lin(l%).data(0).data0.in_point(t_n(1)), t_n(0), m%, o%, t_n(1), _
          l%, l%, 0, -1000, 0, 0, 0, eline_data0, 0, 0, 0, "", record0.record_data.data0.condition_data) Then
           add_aid_point_for_circle = add_aid_point_for_circle1(i%, l%)
            If add_aid_point_for_circle > 1 Then
             Exit Function
            Else
             GoTo add_aid_point_for_circle_mark21
            End If
        End If
      Next o%
     Next m%
add_aid_point_for_circle_mark21:
  Next k%
  Next j%
 '****************
 'If m_poi(m_Circ(i%).data(0).data0.center).data(0).no_reduce = 2 Then
  Call set_point_no_reduce(m_Circ(i%).data(0).data0.center, False)
   If from_old_to_aid = 1 Then
       Exit Function
   End If
    temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
    For j% = 2 To m_Circ(i%).data(0).data0.in_point(0)
     For k% = 1 To j% - 1
     temp_record.record_data.data0.condition_data.condition_no = 0 'record0
     add_aid_point_for_circle = set_equal_dline(m_Circ(i%).data(0).data0.center, _
        m_Circ(i%).data(0).data0.in_point(j%), m_Circ(i%).data(0).data0.center, _
         m_Circ(i%).data(0).data0.in_point(k%), 0, 0, 0, 0, 0, 0, 0, _
          temp_record, 0, 0, 0, 0, 0, False)
        If add_aid_point_for_circle > 1 Then
         Exit Function
        End If
     Next k%
    Next j%
     add_aid_point_for_circle = start_prove(0, 1, 1)
        If add_aid_point_for_circle > 1 Then
         Exit Function
        End If
        'If new_result_from_add = False Then
         Call from_aid_to_old
        'Else
        'new_result_from_add = False
        'End If
  Call set_point_no_reduce(m_Circ(i%).data(0).data0.center, True)
 'End If
 If m_Circ(i%).data(0).data0.center > 0 Then
     For j% = 2 To m_Circ(i%).data(0).data0.in_point(0)
     For k% = 1 To j% - 1
     l% = line_number0(m_Circ(i%).data(0).data0.in_point(j%), m_Circ(i%).data(0).data0.in_point(k%), 0, 0)
     If m_lin(l%).data(0).data0.visible > 0 Then
      add_aid_point_for_circle = add_aid_point_for_paral_or_verti(m_Circ(i%).data(0).data0.center, _
               verti_, l%, l%, 0)
           If add_aid_point_for_circle > 1 Then
            Exit Function
           End If
     End If
     If is_three_point_on_line(m_Circ(i%).data(0).data0.in_point(j%), m_Circ(i%).data(0).data0.in_point(k%), _
          m_Circ(i%).data(0).data0.center, 0, 0, 0, 0, 0, 0, 0) Then
     GoTo add_aid_point_for_circle_next
     End If
     add_aid_point_for_circle = add_aid_point_for_circle2(i%, m_Circ(i%).data(0).data0.in_point(k%))
     If add_aid_point_for_circle > 1 Then
       Exit Function
     End If
add_aid_point_for_circle_next:
     Next k%
    Next j%
 End If
add_aid_point_for_circle_error:
 '*******
End Function

Public Function add_aid_point_for_tangent_line(no%) As Byte
Dim i%
Dim l%, tn%
Dim temp_record As total_record_type
If m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center Then
l% = line_number0(con_tangent_line(no%).data(0).poi(0), _
        m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center, 0, 0)
 For i% = 1 To m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(0)
   If m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(i%) <> _
            con_tangent_line(no%).data(0).poi(0) Then
     If l% = line_number0(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(i%), _
        m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center, 0, 0) Then
         Exit Function
     End If
    End If
 Next i%
 'On Error GoTo add_aid_point_for_tangent_line_error
   If from_old_to_aid = 1 Then
      Exit Function
   End If
 '***
 last_conditions.last_cond(1).point_no% = last_conditions.last_cond(1).point_no% + 1
 t_coord.X = _
     2 * m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center).data(0).data0.coordinate.X - _
        m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.coordinate.X
 t_coord.Y = _
     2 * m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center).data(0).data0.coordinate.Y - _
        m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.coordinate.Y
    Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
    record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(last_conditions.last_cond(1).point_no%, l%, tn%, False, False, 0, temp_record)
     Call set_two_point_line_for_line(l%, temp_record.record_data)
      Call arrange_data_for_new_point(l%, 0)
Else
 t_coord.X = _
     2 * m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.c_coord.X - _
        m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.coordinate.X
 t_coord.Y = _
     2 * m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.c_coord.Y - _
        m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.coordinate.Y
        Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
    l% = line_number0(last_conditions.last_cond(1).point_no%, con_tangent_line(no%).data(0).poi(0), tn%, 0)
End If
 Call set_point_name(last_conditions.last_cond(1).point_no, _
   next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
   Call add_point_to_m_circle( _
             last_conditions.last_cond(1).point_no%, con_tangent_line(no%).data(0).circ(0), temp_record, 255)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
   ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
  last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
       temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
         temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
         temp_record.record_data.data0.theorem_no = 0
   new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = l%
         'new_point(last_conditions.last_cond(1).new_point_no). .add_to_line(1) = tl%
          'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
          'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       'n% = 0
       If m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center > 0 Then
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1370, _
          "\\1\\" + m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.name + _
          "\\2\\" + m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.center).data(0).data0.name + _
          "\\3\\" + m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.name + _
            m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
       Else
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1370, _
         "\\1\\" + m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.name + _
         "\\2\\" + m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(1)).data(0).data0.name + _
          m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(2)).data(0).data0.name + _
           m_poi(m_Circ(con_tangent_line(no%).data(0).circ(0)).data(0).data0.in_point(3)).data(0).data0.name + _
         "\\3\\" + m_poi(con_tangent_line(no%).data(0).poi(0)).data(0).data0.name + _
             m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
       End If
      temp_record.record_data.data0.condition_data.condition_no = 1
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      temp_record.record_data.data0.condition_data.condition(i%).no = last_conditions.last_cond(1).new_point_no
      add_aid_point_for_tangent_line = set_New_point(last_conditions.last_cond(1).point_no, temp_record, l%, 0, _
             tn%, 0, con_tangent_line(no%).data(0).circ(0), 0, 0, 1)
       If add_aid_point_for_tangent_line > 1 Then
        Exit Function
       End If
      add_aid_point_for_tangent_line = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
       If add_aid_point_for_tangent_line > 1 Then
         Exit Function
       End If
       'If ty = 0 Then
add_aid_point_for_tangent_line_error:
        Call from_aid_to_old
       'Else
       'new_result_from_add = False
       'End If
      '***
End Function
Public Function add_aid_point_for_circle1(c%, l%) As Byte
Dim m%, o%, tl%, n%, tn_%
Dim r&
Dim s&
Dim t!
Dim temp_record As total_record_type
      For m% = 1 To m_Circ(c%).data(0).data0.in_point(0)
       If is_point_in_line3(m_Circ(c%).data(0).data0.in_point(m%), m_lin(l%).data(0).data0, 0) = False Then 'm%,不在l%上
           If is_dverti(line_number0(m_Circ(c%).data(0).data0.in_point(m%), _
                m_Circ(c%).data(0).data0.center, 0, 0), l%, 0, -1000, 0, 0, 0, 0) Then  '经垂经
                 GoTo add_aid_point_to_circle
'           ElseIf is_dverti(l%, line_number0(m_circ(c%).data(0).data0.in_point(m%), _
 '               m_circ(c%).data(0).data0.center, 0, 0), 0, -1000, 0, 0, 0, 0) Then
 '                 GoTo add_aid_point_to_circle
           Else
            For o% = m% + 1 To m_Circ(c%).data(0).data0.in_point(0)
  '确定o%没有对称点
               If is_point_in_line3(m_Circ(c%).data(0).data0.in_point(o%), m_lin(l%).data(0).data0, 0) = False Then
                tl% = line_number0(m_Circ(c%).data(0).data0.in_point(m%), m_Circ(c%).data(0).data0.in_point(o%), 0, 0)
               If tl% <> l% Then
                 If is_dparal(l%, line_number0( _
                    m_Circ(c%).data(0).data0.in_point(m%), m_Circ(c%).data(0).data0.in_point(o%), 0, 0), _
                        0, -1000, 0, 0, 0, 0) Then
                  GoTo add_aid_point_to_circle
                 End If
               ElseIf tl% = l% Then
                 GoTo add_aid_point_to_circle
               End If
              End If
             Next o%
            End If
  If from_old_to_aid = 1 Then
     Exit Function
  End If
  'On Error GoTo add_aid_point_for_circle1_error
         r& = (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
               m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
               (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
                m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) ^ 2
         s& = (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
               m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X) * _
                (m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.coordinate.X - _
                  m_Circ(c%).data(0).data0.c_coord.X) + _
                   (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
                    m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) * _
                     (m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.coordinate.Y - _
                       m_Circ(c%).data(0).data0.c_coord.Y)
        t! = -2 * CSng(s&) / r&
        If last_conditions.last_cond(1).point_no = 26 Then
         add_aid_point_for_circle1 = 6
          Exit Function
        End If
        last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'        MDIForm1.Toolbar1.Buttons(21).Image = 33
         Call set_point_name(last_conditions.last_cond(1).point_no, _
             next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
          'poi(last_conditions.last_cond(1).point_no).data(0).data0.visible = 1
         t_coord.X = m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.coordinate.X + _
             (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
                m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X) * t!
         t_coord.Y = m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.coordinate.Y + _
             (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
                m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) * t!
         Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
        record_0.data0.condition_data.condition_no = 0 'record0
       tl% = line_number0(m_Circ(c%).data(0).data0.in_point(m%), last_conditions.last_cond(1).point_no, 0, tn_%)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
    last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
      ' new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
      ' poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
   n% = 0
   If m_Circ(c%).data(0).data0.center > 0 Then
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1430, "\\1\\" + m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.name + _
                        "\\2\\" + m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.name + _
                        "\\3\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + "(" + _
                                  m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                        "\\4\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
   Else
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1435, "\\1\\" + m_poi(m_Circ(c%).data(0).data0.in_point(m%)).data(0).data0.name + _
                        "\\2\\" + m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.name + _
                        "\\3\\" + m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + _
                                  m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + _
                                  m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).data0.name + _
                        "\\4\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
   End If
     'Call set_dparal(l%, tl%, temp_record, n%, 0, False)
          temp_record.record_data.data0.condition_data.condition_no = 1
            temp_record.record_data.data0.condition_data.condition(1).ty = paral_
             temp_record.record_data.data0.condition_data.condition(1).no = n%
              new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = paral_
               new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
               temp_record.record_data.data0.theorem_no = 0
           Call set_dparal(l%, tl%, temp_record, n%, 0, False)
      add_aid_point_for_circle1 = set_New_point(last_conditions.last_cond(1).point_no, temp_record, _
           tl%, 0, tn_%, 0, c%, 0, 0, 1)
       If add_aid_point_for_circle1 > 1 Then
         Exit Function
       End If
      add_aid_point_for_circle1 = start_prove(0, 1, 1)   'call_theorem(0)
       If add_aid_point_for_circle1 > 1 Then
         Exit Function
      Else
add_aid_point_for_circle1_error:
 Call from_aid_to_old
      End If
       End If
add_aid_point_to_circle:
     Next m%

End Function
Public Function add_aid_point_for_circle2(ByVal c%, ByVal p%) As Byte
Dim l%, n%, p1%, p2%
Dim temp_record As total_record_type
'On Error GoTo add_aid_point_for_circle2_error
l% = line_number0(p%, m_Circ(c%).data(0).data0.center, 0, 0)
n% = inter_point_line_circle0(m_lin(l%).data(0).data0, m_Circ(c%).data(0).data0, p1%, p2%)
If n% = 2 Then
 Exit Function
ElseIf n% = 1 Then
 If compare_two_point(m_poi(p).data(0).data0.coordinate, m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate, 0, 0, 0) = _
     compare_two_point(m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate, 0, 0, 0) Then
      Exit Function
 End If
End If
  If from_old_to_aid = 1 Then
     Exit Function
  End If
        last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'        MDIForm1.Toolbar1.Buttons(21).Image = 33
         Call set_point_name(last_conditions.last_cond(1).point_no, _
            next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
          'poi(last_conditions.last_cond(1).point_no).data(0).data0.visible = 1
         t_coord.X = 2 * m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate.X - _
                         m_poi(p%).data(0).data0.coordinate.X
         t_coord.Y = 2 * m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate.Y - _
                         m_poi(p%).data(0).data0.coordinate.Y
         Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
'         poi(last_conditions.last_cond(1).point_no)..data(0).
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
    last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
       'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
      ' poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
      l% = line_number0(m_Circ(c%).data(0).data0.center, p%, 0, 0)
    record_0.data0.condition_data.condition_no = 0
   Call add_point_to_line(last_conditions.last_cond(1).point_no, l%, n%, False, False, 0, temp_record)
   Call add_point_to_m_circle(last_conditions.last_cond(1).point_no, c%, temp_record, 255)
   If p% <> m_Circ(c%).data(0).data0.in_point(1) Then
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
     LoadResString_(1440, "\\1\\" + m_poi(p%).data(0).data0.name + _
                                   m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                         "\\2\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                                  "(" + m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                         "\\3\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
   Else
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1440, "\\1\\" + m_poi(p%).data(0).data0.name + _
                                 m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                        "\\2\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                               "(" + m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + ")" + _
                        "\\3\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
   End If
      temp_record.record_data.data0.condition_data.condition_no = 1
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      add_aid_point_for_circle2 = set_New_point(last_conditions.last_cond(1).point_no, temp_record, _
           l%, 0, n%, 0, c%, 0, 0, 1)
       If add_aid_point_for_circle2 > 1 Then
         Exit Function
       End If
      add_aid_point_for_circle2 = start_prove(0, 1, 1)   'call_theorem(0)
       If add_aid_point_for_circle2 > 1 Then
         Exit Function
      Else
add_aid_point_for_circle2_error:
 Call from_aid_to_old
      End If
       
 
End Function
Public Function add_aid_point_for_com_tangent_line(ByVal t_l1%, ByVal t_l2%) As Byte
Dim tp%
Dim c(1) As Integer
Dim p(3) As Integer
Dim temp_record As record_type
Dim c_data0 As condition_data_type
'On Error GoTo add_aid_point_for_com_tangent_line_error
 If tangent_line(t_l1%).data(0).circ(0) = tangent_line(t_l2%).data(0).circ(0) Then
      c(0) = tangent_line(t_l1%).data(0).circ(0)
       p(0) = tangent_line(t_l1%).data(0).poi(0)
       p(1) = tangent_line(t_l2%).data(0).poi(0)
 ElseIf tangent_line(t_l1%).data(0).circ(0) = tangent_line(t_l2%).data(0).circ(1) Then
      c(0) = tangent_line(t_l1%).data(0).circ(0)
       p(0) = tangent_line(t_l1%).data(0).poi(0)
       p(1) = tangent_line(t_l2%).data(0).poi(1)
End If
If tangent_line(t_l1%).data(0).circ(1) > 0 Then
 If tangent_line(t_l1%).data(0).circ(1) = tangent_line(t_l2%).data(0).circ(0) Then
      c(1) = tangent_line(t_l1%).data(0).circ(1)
        p(2) = tangent_line(t_l1%).data(0).poi(1)
        p(3) = tangent_line(t_l2%).data(0).poi(0)
ElseIf tangent_line(t_l1%).data(0).circ(1) = tangent_line(t_l2%).data(0).circ(1) Then
      c(1) = tangent_line(t_l1%).data(0).circ(1)
        p(2) = tangent_line(t_l1%).data(0).poi(1)
        p(3) = tangent_line(t_l2%).data(0).poi(1)
End If
End If
 If c(0) > 0 Or c(1) > 0 Then
 'If is_line_line_intersect(Lin(tangent_line(t_l1%).data(0).line_no), _
                     Lin(tangent_line(t_l2%).data(0).line_no), 0, 0) Then
          Exit Function
 'Else
  'Call from_old_to_aid
   add_aid_point_for_com_tangent_line = _
    add_interset_point_line_line(tangent_line(t_l1%).data(0).line_no, _
              tangent_line(t_l2%).data(0).line_no, tp%, 0, 1, 0, 0, c_data0)
    If add_aid_point_for_com_tangent_line > 1 Then
     Exit Function
    End If
 End If
add_aid_point_for_com_tangent_line_error:
End Function

Public Function add_aid_point_for_point3_on_line(is_no_initial As Integer, c_data0 As condition_data_type) As Byte
Dim i%, j%, p%, k%, l%
Dim tp(4) As Integer
Dim tem_p(2) As Integer
Dim para(2) As String
Dim temp_record1 As total_record_type
'On Error GoTo add_aid_point_for_point3_on_line_error
c_data0.condition_no = 0
  For i% = 1 To last_conclusion
   If conclusion_data(i% - 1).no(0) = 0 Then
    If conclusion_data(i% - 1).ty = point3_on_line_ Then
     p% = con_Three_point_on_line(i% - 1).data(0).poi(0)
   tp(0) = 0
   For j% = 1 To 2
    If p% < con_Three_point_on_line(i% - 1).data(0).poi(j%) Then
     p% = con_Three_point_on_line(i% - 1).data(0).poi(j%)
      tp(0) = j%
    End If
   Next j%
   If tp(0) = 0 Then
    tp(1) = con_Three_point_on_line(i% - 1).data(0).poi(1)
    tp(2) = con_Three_point_on_line(i% - 1).data(0).poi(2)
   ElseIf tp(0) = 1 Then
    tp(1) = con_Three_point_on_line(i% - 1).data(0).poi(0)
    tp(2) = con_Three_point_on_line(i% - 1).data(0).poi(2)
   Else
    tp(1) = con_Three_point_on_line(i% - 1).data(0).poi(0)
    tp(2) = con_Three_point_on_line(i% - 1).data(0).poi(1)
   End If
   For j% = 1 To m_poi(p%).data(0).in_line(0)
    If m_lin(m_poi(p%).data(0).in_line(j%)).data(0).data0.in_point(0) > 2 Then
     l% = m_poi(p%).data(0).in_line(j%)
      tp(3) = m_lin(l%).data(0).data0.poi(0)
      tp(4) = m_lin(l%).data(0).data0.poi(1)
       For k% = 2 To m_lin(l%).data(0).data0.in_point(0) - 1
        If m_lin(l%).data(0).data0.in_point(k%) < tp(3) Then
         tp(3) = m_lin(l%).data(0).data0.in_point(k%)
        ElseIf m_lin(l%).data(0).data0.in_point(k%) < tp(4) Then
         tp(4) = m_lin(l%).data(0).data0.in_point(k%)
        End If
       Next k%
    If tp(3) < p% And tp(3) < tp(1) And tp(3) < tp(2) And _
        tp(4) < p% And tp(4) < tp(1) And tp(4) < tp(2) Then
     tp(0) = 0
      add_aid_point_for_point3_on_line = add_interset_point_line_line(line_number0(tp(1), tp(2), 0, 0), _
        line_number0(tp(3), tp(4), 0, 0), tp(0), 0, 1, is_no_initial, 0, c_data0)
        add_aid_point_for_point3_on_line = set_item0(tp(3), p%, p%, tp(4), "/", 0, 0, 0, 0, 0, 0, _
            "1", "1", "1", "", para(0), i%, record_0.data0.condition_data, 0, tem_p(0), 0, _
               is_no_initial, c_data0, False) '0310
        item0(tem_p(0)).data(0).conclusion_no = i%
        If add_aid_point_for_point3_on_line > 1 Then
         Exit Function
        End If
        add_aid_point_for_point3_on_line = set_item0(tp(3), tp(0), tp(0), tp(4), "/", 0, 0, 0, 0, 0, 0, _
            "1", "1", "1", "", para(1), i%, record_0.data0.condition_data, 0, tem_p(1), 0, is_no_initial, _
              c_data0, False) '0310
         item0(tem_p(1)).data(0).conclusion_no = i%
         temp_record1.record_.conclusion_no = i%
         If add_aid_point_for_point3_on_line > 1 Then
           Exit Function
         End If
        add_aid_point_for_point3_on_line = set_general_string(tem_p(0), tem_p(1), 0, 0, para(0), _
            time_string("-1", para(1), True, False), "0", "0", _
              "", i%, 0, 0, temp_record1, 0, 0)
        If add_aid_point_for_point3_on_line > 1 Then
         Exit Function
        End If
        add_aid_point_for_point3_on_line = start_prove(0, 1, 1)
         If add_aid_point_for_point3_on_line > 1 Then
          Exit Function
         End If
    End If
  End If
  Next j%
  End If
  End If
  Next i%
add_aid_point_for_point3_on_line_error:
End Function

Public Function add_aid_point_for_eangle0(ByVal A%, ByVal p1%, ByVal p2%, ByVal p3%, k As Single) As Byte
'在直线p2%p3%上取一点p%使得∠pp1%p2%=A%一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
Dim r1!
Dim r2!
Dim new_a%, tp%, l%, tn%, new_p%, n%
Dim p As POINTAPI
Dim temp_record As total_record_type
r1! = (m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2
r1! = sqr(r1!)
r1! = (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) ^ 2
r1! = sqr(r1!)
p.X = m_poi(p2%).data(0).data0.coordinate.X + _
    (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) * k * r1! / r2!
p.Y = m_poi(p2%).data(0).data0.coordinate.Y + _
    (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) * k * r1! / r2!
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_eangle0 = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 tp% = last_conditions.last_cond(1).point_no
   Call set_point_coordinate(tp%, p, False)
Call get_new_char(tp%)
l% = line_number0(p2%, p3%, 0, 0)
    record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(tp%, l%, tn%, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(l%, temp_record.record_data)
       Call arrange_data_for_new_point(l%, 0)
new_a% = Abs(angle_number(tp%, p1%, p2%, "", 0))
'************************
 If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
 End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
          new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       new_point(new_p%).data(0).add_to_line(0) = l%
       n% = 0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
         "\\1\\" + m_poi(p2%).data(0).data0.name + m_poi(p3%).data(0).data0.name + _
         "\\2\\" + m_poi(tp%).data(0).data0.name + _
         "\\3\\" + set_display_angle0(m_poi(tp%).data(0).data0.name + m_poi(p1%).data(0).data0.name + _
            m_poi(p2%).data(0).data0.name) + "=" + set_display_angle(A%, False))
      n% = 0
   Call set_three_angle_value(A%, new_a%, 0, "1", "-1", "0", "0", 0, temp_record, n%, 0, 0, 0, 0, 0, False)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
     l% = line_number0(tp%, p2%, tn%, 0)
  add_aid_point_for_eangle0 = set_New_point(tp%, temp_record, l%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_eangle0 > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_eangle0 = start_prove(0, 1, 1) 'call_theorem(0, no_reduce)
   If add_aid_point_for_eangle0 > 1 Then
    Exit Function
   End If
add_aid_point_for_eangle0_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
'End If ' new_result_from_add = False
' End If

End Function

Public Function add_aid_point_for_eangle_(ByVal p1%, ByVal p2%, ByVal p3%, ByVal A%, tp%, n%, _
        cond_ty As Byte, n1%, n2%, n3%, is_pseudo As Byte) As Byte
'在直线p2%,p3%上取一点使得∠p1%pp2%=∠p1%p2%p3%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
Dim r1!
Dim r2!
Dim tA(1) As Integer
Dim tl%, new_p%
Dim p As POINTAPI
Dim tn%
Dim t_n(1) As Integer
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
'On Error GoTo add_aid_point_for_eangle_mark1
If is_pseudo = 0 Then
If from_old_to_aid = 1 Then
    Exit Function
End If
End If
r1! = (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) ^ 2
r2! = (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) * _
        (m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) + _
          (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) * _
           (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y)
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_eangle_ = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
 tp% = last_conditions.last_cond(1).point_no
Call get_new_char(tp%)
 r1! = 2 * r2! / r1!
 p.X = m_poi(p2%).data(0).data0.coordinate.X + _
            (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) * r1!
 p.Y = m_poi(p2%).data(0).data0.coordinate.Y + _
            (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) * r1!
  If read_point(p, 0) > 0 Then
   tA(0) = Abs(angle_number(p1%, tp%, p2%, 0, 0))
    If is_equal_angle(tA(0), A%, 0, 0) Then
     GoTo add_aid_point_for_eangle_mark1
   End If
  End If
  Call set_point_coordinate(tp%, p, False)
  tl% = line_number0(p2%, p3%, 0, 0)
  add_aid_point_for_eangle_ = set_New_point(tp%, temp_record, tl%, 0, _
        tn%, 0, 0, 0, 0, 1)
  record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(tp%, tl%, tn%, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
        tA(0) = Abs(angle_number(p1%, tp%, p2%, 0, 0))
        tA(1) = Abs(angle_number(p1%, p2%, tp%, 0, 0))
 If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
 End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
          new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       new_point(new_p%).data(0).add_to_line(0) = tl%
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
          "\\1\\" + m_poi(p2%).data(0).data0.name + m_poi(p3%).data(0).data0.name + _
          "\\2\\" + m_poi(tp%).data(0).data0.name + _
          "\\3\\" + set_display_angle0(m_poi(p1%).data(0).data0.name + m_poi(tp%).data(0).data0.name + _
              m_poi(p2%).data(0).data0.name) + "=" + set_display_angle(tA(1), False))
      n% = 0
   Call set_three_angle_value(tA(0), tA(1), 0, "1", "-1", "0", "0", 0, temp_record, n%, n1%, n2%, 0, 0, 0, False)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
     temp_record.record_data.data0.condition_data.condition_no = 0
       Call add_conditions_to_record(angle3_value_, n%, n1%, n2%, _
           temp_record.record_data.data0.condition_data)
          temp_record.record_data.data0.theorem_no = 0
    If is_pseudo = 1 Then
     temp_record.record_data.data0.theorem_no = 40
      n1% = 0
       add_aid_point_for_eangle_ = _
           set_equal_dline(p1%, p2%, p1%, tp%, 0, 0, 0, 0, 0, 0, 0, temp_record, n1%, cond_ty, n2%, n3%, 0, False)
      Exit Function
    End If
   If add_aid_point_for_eangle_ > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_eangle_ = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_eangle_ > 1 Then
    Exit Function
   End If
add_aid_point_for_eangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
'End If ' new_result_from_add = False
' End If
End Function

Public Function add_aid_point_for_double_angle_(ByVal p1%, ByVal p2%, ByVal p3%, _
                     ByVal p4%, ByVal p5%, ByVal p6%) As Byte
'在直线p2%,p3%上取一点使得∠p%p1%p2%=∠p5%p4%p6%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
Dim r1!
Dim r2!
Dim r_(3) As Single
Dim tA(1) As Integer
Dim tl%, new_p%, tp%, n%
Dim p As POINTAPI
Dim tn%, n1%, n2%
Dim t_n(1) As Integer
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
'On Error GoTo add_aid_point_for_double_angle_
If from_old_to_aid = 1 Then
   Exit Function
End If
r_(0) = (m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2
r_(1) = (m_poi(p4%).data(0).data0.coordinate.X - m_poi(p5%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p4%).data(0).data0.coordinate.Y - m_poi(p5%).data(0).data0.coordinate.Y) ^ 2
r_(2) = (m_poi(p5%).data(0).data0.coordinate.X - m_poi(p6%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p5%).data(0).data0.coordinate.Y - m_poi(p6%).data(0).data0.coordinate.Y) ^ 2
r_(3) = r_(0) * r_(2) / r_(1)
r1! = (m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) ^ 2
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_double_angle_ = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
 tp% = last_conditions.last_cond(1).point_no
Call get_new_char(tp%)
 r1! = r_(3) / r1!
 p.X = m_poi(p2%).data(0).data0.coordinate.X + _
            (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) * r1!
 p.Y = m_poi(p2%).data(0).data0.coordinate.Y + _
            (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) * r1!
  If read_point(p, 0) > 0 Then
        tA(0) = Abs(angle_number(p2%, p1%, tp%, 0, 0))
        tA(1) = Abs(angle_number(p5%, p4%, p6%, 0, 0))
    If is_equal_angle(tA(0), tA(1), 0, 0) Then
     GoTo add_aid_point_for_eangle_mark1
    End If
  End If
  Call set_point_coordinate(tp%, p, False)
  tl% = line_number0(p2%, p3%, 0, 0)
    record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(tp%, tl%, tn%, no_display, False, 0, temp_record)
      Call set_two_point_line_for_line(tl%, temp_record.record_data)
       Call arrange_data_for_new_point(tl%, 0)
 If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
 End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
          new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       'new_point(new_p%).data(0).record = temp_record.record_data
       new_point(new_p%).data(0).add_to_line(0) = tl%
        'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
         "\\1\\" + m_poi(p2%).data(0).data0.name + m_poi(p3%).data(0).data0.name + _
         "\\2\\" + m_poi(tp%).data(0).data0.name + _
         "\\3\\" + set_display_angle0(m_poi(p2%).data(0).data0.name + m_poi(p1%).data(0).data0.name + _
               m_poi(p2%).data(0).data0.name) + "=" + set_display_angle(tA(1), False))
   Call set_three_angle_value(tA(0), tA(1), 0, "1", "-1", "0", "0", 0, temp_record, n%, n1%, n2%, 0, 0, 0, False)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
     temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(angle3_value_, n%, n1%, n2%, _
              temp_record.record_data.data0.condition_data)
          temp_record.record_data.data0.theorem_no = 0
  add_aid_point_for_double_angle_ = set_New_point(tp%, temp_record, tl%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_double_angle_ > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_double_angle_ = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_double_angle_ > 1 Then
    Exit Function
   End If
add_aid_point_for_eangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
'End If ' new_result_from_add = False
' End If
add_aid_point_for_double_angle_:
End Function
Public Function add_aid_point_for_t_e_triangle2(triA1 As temp_triangle_data_type, triA2 As temp_triangle_data_type) As Byte '
'在直线p2%,p3%上取一点使得∠p1%pp2%=∠p1%p2%p3%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
'作 两三角形全等
 add_aid_point_for_t_e_triangle2 = _
        add_aid_point_for_eline(triA1.poi(0), triA1.poi(1), _
         triA2.poi(0), triA2.poi(1), triA2.poi(1))
       '在直线p3%p4%上取点p使得p3%p=p1%p2%

End Function
Public Function add_aid_point_for_t_e_triangle1(triA1 As temp_triangle_data_type, triA2 As temp_triangle_data_type) As Byte '
 add_aid_point_for_t_e_triangle1 = _
        add_aid_point_for_eline(triA1.poi(0), triA1.poi(1), _
         triA2.poi(1), triA2.poi(0), triA2.poi(1))
       '在直线p3%p4%上取点p使得p3%p=p1%p2%
If add_aid_point_for_t_e_triangle1 > 1 Then
 Exit Function
End If
If angle(triA1.angle(1)).data(0).value = "90" And angle(triA2.angle(1)).data(0).value = "90" Then
  add_aid_point_for_t_e_triangle1 = _
        add_aid_point_for_eline(triA1.poi(0), triA1.poi(1), _
         triA2.poi(0), triA2.poi(1), triA2.poi(1))
End If
End Function
Public Function add_aid_point_for_double_angle0(ByVal A%, ByVal A2%) As Byte
'在直线p2%,p3%上取一点使得∠p%p1%p2%=∠p5%p4%p6%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
Dim r1!
Dim r_(1) As Single
Dim tl%, dr%, dr1%, tp%, n%, tA%
Dim p(0) As POINTAPI
Dim tn%, i%, j%, new_p%, n1%, n2%
Dim t_n(1) As Integer
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
Dim tp_(2) As Integer '角A%的三个点
'On Error GoTo add_aid_point_for_double_angle0
If angle(A%).data(0).te(0) = 0 Then
 tp_(0) = m_lin(angle(A%).data(0).line_no(0)).data(0).data0.poi(0)
Else
 tp_(0) = m_lin(angle(A%).data(0).line_no(0)).data(0).data0.poi(1)
End If
If angle(A%).data(0).te(1) = 0 Then
 tp_(1) = m_lin(angle(A%).data(0).line_no(1)).data(0).data0.poi(0)
Else
 tp_(1) = m_lin(angle(A%).data(0).line_no(1)).data(0).data0.poi(1)
End If
tp_(2) = angle(A%).data(0).poi(1)
'*************************************************************************************
r_(0) = distance_of_two_POINTAPI(m_poi(tp_(2)).data(0).data0.coordinate, _
                            m_poi(tp_(0)).data(0).data0.coordinate) '线段t_p(0),t_p(2)(第一邻边)的长
r_(1) = distance_of_two_POINTAPI(m_poi(tp_(2)).data(0).data0.coordinate, _
                            m_poi(tp_(1)).data(0).data0.coordinate) '线段t_p(1),t_p(2)(第二邻边)的长
p(0) = add_POINTAPI(m_poi(tp_(0)).data(0).data0.coordinate, _
                    time_POINTAPI_by_number(minus_POINTAPI( _
                       m_poi(tp_(1)).data(0).data0.coordinate, m_poi(tp_(0)).data(0).data0.coordinate), _
                         r_(0) / (r_(0) + r_(1))))
'p 位于角A的平分线上
For i% = 1 To m_lin(angle(A%).data(0).line_no(0)).data(0).data0.in_point(0)
     tp_(0) = m_lin(angle(A%).data(0).line_no(0)).data(0).data0.in_point(i%) '第一邻边上选一点
     If tp_(0) <> tp_(2) Then
        For j% = 1 To m_poi(tp_(0)).data(0).in_line(0)
                tl% = m_poi(tp_(0)).data(0).in_line(j%)
            If tl% <> angle(A%).data(0).line_no(0) Then '与第一邻边不同的直线
              If from_old_to_aid = 1 Then
                  Exit Function
              End If
              If m_lin(tl%).data(0).data0.visible > 0 Then
                    dr1% = angle_number(m_lin(tl%).data(0).data0.poi(0), tp_(2), _
                                   m_lin(tl%).data(0).data0.poi(1), 0, 0)
                    If dr1% = 0 Then
                       GoTo add_aid_point_for_eangle_mark1
                    End If
               If calculate_line_line_intersect_point(m_poi(tp_(2)).data(0).data0.coordinate, p(0), _
                   m_poi(m_lin(tl%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                    m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.coordinate, p(1), False) Then
                    tp% = m_point_number(p(1), condition, 1, condition_color, "", condition_type0, _
                      condition_type0, 0, True)
                   record_0.data0.condition_data.condition_no = 0
         Call add_point_to_line(tp%, tl%, 0, no_display, False, 0, temp_record)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
          new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       'new_point(new_p%).data(0).record = temp_record.record_data
       new_point(new_p%).data(0).add_to_line(0) = tl%
        'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
         "\\1\\" + m_poi(m_lin(tl%).data(0).data0.poi(0)).data(0).data0.name + m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.name + _
         "\\2\\" + m_poi(tp%).data(0).data0.name + _
         "\\3\\" + set_display_angle0(m_poi(tp_(0)).data(0).data0.name + m_poi(tp_(2)).data(0).data0.name + _
               m_poi(tp%).data(0).data0.name) + "=" + set_display_angle(A2%, False))
      n% = 0
      tA% = angle_number(tp_(0), tp_(2), tp%, 0, 0)
      tA% = Abs(tA%)
   Call set_three_angle_value(tA, A2%, 0, "1", "-1", "0", "0", 0, temp_record, n%, n1%, n2%, 0, 0, 0, False)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
     temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(angle3_value_, n%, n1%, n2%, _
                        temp_record.record_data.data0.condition_data)
          temp_record.record_data.data0.theorem_no = 0
  add_aid_point_for_double_angle0 = set_New_point(tp%, temp_record, tl%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_double_angle0 > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_double_angle0 = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_double_angle0 > 1 Then
    Exit Function
   End If
'add_aid_point_for_eangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old

     End If
    End If
   End If
 Next j%
 End If
Next i%
'***************************************
For i% = 1 To m_lin(angle(A%).data(0).line_no(1)).data(0).data0.in_point(0)
tp_(0) = m_lin(angle(A%).data(0).line_no(1)).data(0).data0.in_point(i%)
If tp_(0) <> tp_(2) Then
 For j% = 1 To m_poi(i%).data(0).in_line(0)
 tl% = m_poi(i%).data(0).in_line(j%)
  If tl% <> angle(A%).data(0).line_no(1) Then
  If from_old_to_aid = 1 Then
   Exit Function
 End If
    If m_lin(tl%).data(0).data0.visible > 0 Then
       dr1% = angle_number(m_lin(tl%).data(0).data0.poi(0), tp_(2), _
                                   m_lin(tl%).data(0).data0.poi(1), 0, 0)
          If dr1% = 0 Then
               GoTo add_aid_point_for_eangle_mark1
          End If
     If calculate_line_line_intersect_point(m_poi(tp_(2)).data(0).data0.coordinate, p(0), _
            m_poi(m_lin(tl%).data(0).data0.poi(0)).data(0).data0.coordinate, _
             m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.coordinate, p(1), False) Then
      tp% = m_point_number(p(1), condition, 1, condition_color, "", condition_type0, _
            condition_type0, 0, True)
       record_0.data0.condition_data.condition_no = 0
         Call add_point_to_line(tp%, tl%, 0, no_display, False, 0, temp_record)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
          new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       'new_point(new_p%).data(0).record = temp_record.record_data
       new_point(new_p%).data(0).add_to_line(0) = tl%
        'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1350, _
         "\\1\\" + m_poi(m_lin(tl%).data(0).data0.poi(0)).data(0).data0.name + m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.name + _
         "\\2\\" + m_poi(tp%).data(0).data0.name + _
         "\\3\\" + set_display_angle0(m_poi(tp_(0)).data(0).data0.name + m_poi(tp_(1)).data(0).data0.name + _
               m_poi(tp%).data(0).data0.name) + "=" + set_display_angle(A2%, False))
      n% = 0
      tA% = angle_number(tp_(0), tp_(2), tp%, 0, 0)
      tA% = Abs(tA%)
   Call set_three_angle_value(tA, A2%, 0, "1", "-1", "0", "0", 0, temp_record, n%, n1%, n2%, 0, 0, 0, False)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
     temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(angle3_value_, n%, n1%, n2%, _
                         temp_record.record_data.data0.condition_data)
          temp_record.record_data.data0.theorem_no = 0
  add_aid_point_for_double_angle0 = set_New_point(tp%, temp_record, tl%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_double_angle0 > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_double_angle0 = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_double_angle0 > 1 Then
    Exit Function
   End If
add_aid_point_for_eangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old

     End If
    End If
  End If
 Next j%
 End If
Next i%

add_aid_point_for_double_angle0:
End Function

Public Function add_aid_point_for_t_e_triangle3(triA1 As temp_triangle_data_type, triA2 As temp_triangle_data_type) As Byte '
'在直线p2%,p3%上取一点使得∠p1%pp2%=∠p1%p2%p3%ty=0 一般辅助点,ty=1 pseudo_triangle,n%返回等角,n1%返回等线段
'作 两三角形全等
Dim n%
Dim r1!
Dim r2!
Dim s!
Dim c!
Dim a_v!
Dim tA(1) As Integer
Dim tl%, new_p%
Dim p As POINTAPI
Dim tp%
Dim tn%
Dim t_n(1) As Integer
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
'On Error GoTo add_aid_point_for_t_e_triangle_mark1
 Call val0(angle(triA1.angle(1)).data(0).value, a_v!)
a_v! = a_v! * PI / 180
s! = Sin(a_v)
c! = Cos(a_v)
'等边长
If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_t_e_triangle3 = 6
  Exit Function
End If
If from_old_to_aid = 1 Then
   Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
 tp% = last_conditions.last_cond(1).point_no
Call get_new_char(tp%)
r1! = (m_poi(triA1.poi(1)).data(0).data0.coordinate.X - m_poi(triA1.poi(0)).data(0).data0.coordinate.X) ^ 2 + _
     (m_poi(triA1.poi(1)).data(0).data0.coordinate.Y - m_poi(triA1.poi(0)).data(0).data0.coordinate.Y) ^ 2
r2! = (m_poi(triA1.poi(1)).data(0).data0.coordinate.X - m_poi(triA1.poi(2)).data(0).data0.coordinate.X) ^ 2 + _
     (m_poi(triA1.poi(1)).data(0).data0.coordinate.Y - m_poi(triA1.poi(2)).data(0).data0.coordinate.Y) ^ 2
r1! = sqr(r1!)
r2! = sqr(r2!)
tA(0) = angle_number(triA1.poi(2), triA1.poi(1), triA1.poi(0), "", 0)
If tA(0) < 0 Then
 s! = -s!
 c! = c!
End If
tA(1) = angle_number(triA2.poi(2), triA2.poi(1), triA2.poi(0), "", 0)
If tA(0) * tA(1) > 0 Then
 p.X = m_poi(triA2.poi(1)).data(0).data0.coordinate.X + _
            ((m_poi(triA2.poi(2)).data(0).data0.coordinate.X - m_poi(triA2.poi(1)).data(0).data0.coordinate.X) * c! - _
              (m_poi(triA2.poi(2)).data(0).data0.coordinate.Y - m_poi(triA2.poi(1)).data(0).data0.coordinate.Y) * s!) * r1! / r2!
 p.Y = m_poi(triA2.poi(1)).data(0).data0.coordinate.Y + _
            ((m_poi(triA2.poi(2)).data(0).data0.coordinate.X - m_poi(triA2.poi(1)).data(0).data0.coordinate.X) * s! + _
             (m_poi(triA2.poi(2)).data(0).data0.coordinate.Y - m_poi(triA2.poi(1)).data(0).data0.coordinate.Y) * c!) * r1! / r2!
Else
 p.X = m_poi(triA2.poi(1)).data(0).data0.coordinate.X + _
           ((m_poi(triA2.poi(2)).data(0).data0.coordinate.X - m_poi(triA2.poi(1)).data(0).data0.coordinate.X) * c! + _
             (m_poi(triA2.poi(2)).data(0).data0.coordinate.Y - m_poi(triA2.poi(1)).data(0).data0.coordinate.Y) * s!) * r1! / r2!
 p.Y = m_poi(triA2.poi(1)).data(0).data0.coordinate.Y + _
            (-(m_poi(triA2.poi(2)).data(0).data0.coordinate.X - m_poi(triA2.poi(1)).data(0).data0.coordinate.X) * s! + _
              (m_poi(triA2.poi(2)).data(0).data0.coordinate.Y - m_poi(triA2.poi(1)).data(0).data0.coordinate.Y) * c!) * r1! / r2!
End If
  If read_point(p, 0) > 0 Then
   tA(0) = Abs(angle_number(triA1.poi(0), triA1.poi(1), triA1.poi(2), 0, 0))
    If is_equal_angle(tA(0), triA1.poi(1), 0, 0) Then
     GoTo add_aid_point_for_t_e_triangle_mark1
   End If
  End If
  Call set_point_coordinate(tp%, p, False)
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
       new_p% = last_conditions.last_cond(1).new_point_no
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(new_p%).data(0).poi(0) = last_conditions.last_cond(1).point_no
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = "作△" + _
        m_poi(tp%).data(0).data0.name + m_poi(triA2.poi(1)).data(0).data0.name + m_poi(triA2.poi(2)).data(0).data0.name + _
          "≌△" + m_poi(triA1.poi(0)).data(0).data0.name + _
            m_poi(triA1.poi(1)).data(0).data0.name + m_poi(triA1.poi(2)).data(0).data0.name
      n% = 0
   Call set_total_equal_triangle(tp%, triA2.poi(1), triA2.poi(2), _
              triA1.poi(0), triA1.poi(1), triA1.poi(2), temp_record, n%, 0)
     temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = total_equal_triangle_
        temp_record.record_data.data0.condition_data.condition(1).no = n%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = total_equal_triangle_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
          temp_record.record_data.data0.theorem_no = 0
  add_aid_point_for_t_e_triangle3 = set_New_point(tp%, temp_record, tl%, 0, _
      tn%, 0, 0, 0, 0, 1)
   If add_aid_point_for_t_e_triangle3 > 1 Then
    Exit Function
   End If
 'If ty = 0 Then
 add_aid_point_for_t_e_triangle3 = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_t_e_triangle3 > 1 Then
    Exit Function
   End If
add_aid_point_for_t_e_triangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
'End If ' new_result_from_add = False
' End If
End Function
 

Private Function add_aid_point_from_eline(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%) As Byte
'等腰三角形
Dim l%
Dim tp(2) As Integer
If p1% = p3% Then
tp(0) = p1%
tp(1) = p2%
tp(2) = p4%
ElseIf p1% = p4% Then
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
ElseIf p2% = p3% Then
tp(0) = p2%
tp(1) = p1%
tp(2) = p4%
ElseIf p2% = p4% Then
tp(0) = p2%
tp(1) = p1%
tp(2) = p3%
Else
Exit Function
End If
l% = line_number0(tp(1), tp(2), 0, 0)
add_aid_point_from_eline = add_aid_point_for_paral_or_verti(tp(0), l%, verti_, l%, 0)
End Function

Public Function add_unknown_value() As Byte
Dim i%, n%
For i% = 1 To last_conclusion
 If conclusion_data(i% - 1).ty = line_value_ And conclusion_data(i% - 1).no(0) = 0 Then
   add_unknown_value = add_new_value_for_line0(con_line_value(i% - 1).data(0).data0.poi(0), _
                         con_line_value(i% - 1).data(0).data0.poi(1), "x", 1, 0)
   If add_unknown_value > 1 Then
      Exit Function
   End If
 End If
Next i%
End Function
Public Function add_new_value_for_line0(ByVal p1%, ByVal p2%, ch$, ty As Byte, no%) As Byte
Dim temp_record As total_record_type
Dim l_v As line_value_data0_type
Dim n%
If is_line_value(p1%, p2%, 0, 0, 0, "", 0, -1000, 0, 0, 0, l_v) = 0 Then
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = 0
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = 0
           n% = 0
     new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1445, _
        "\\1\\" + m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name + "=" + ch$)
        temp_record.record_data.data0.condition_data.condition_no = 1
        temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
        temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
          Call set_line_value(p1%, p2%, ch$, 0, 0, 0, temp_record.record_data, n%, 0, False)
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = line_value_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
         If ty = 1 Then
          add_new_value_for_line0 = start_prove(0, 1, 1)
         End If
End If
End Function
Public Function add_new_value_for_relation0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
          ch$, ty As Byte, no%) As Byte
Dim temp_record As total_record_type
Dim re_v As relation_data0_type
If is_relation(p1%, p2%, p3%, p4%, 0, 0, 0, 0, 0, 0, "", 0, -1000, 0, 0, 0, re_v, _
      0, 0, 0, record_0.data0.condition_data, 0) = False Then
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = 0
        'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = 0
           no% = 0
     new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
         LoadResString_(1445, "\\1\\" + m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name + "/" + _
          m_poi(p3%).data(0).data0.name + m_poi(p4%).data(0).data0.name + "=" + ch$)
       temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
       temp_record.record_data.data0.condition_data.condition(1).no = _
          last_conditions.last_cond(1).new_point_no
         temp_record.record_data.data0.theorem_no = 0
          Call set_Drelation(p1%, p2%, p3%, p4%, 0, 0, 0, 0, 0, 0, ch$, temp_record, no%, 0, 0, 0, 0, False)
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = relation_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = no%
         If ty = 1 Then
          add_new_value_for_relation0 = start_prove(0, 1, 1)
         End If
End If
End Function
Public Function add_new_value_for_angle0(ByVal A%, ch$, ty As Byte, n%) As Byte
Dim temp_record As total_record_type
Dim re_v As relation_data0_type
If angle(A%).data(0).value = "" Then
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
    temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = 0
        'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_line(0) = 0
           n% = 0
     new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
         LoadResString_(1445, "\\1\\" + set_display_angle(A%, False) + "=" + ch$)
        temp_record.record_data.data0.condition_data.condition_no = 1
        temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
        temp_record.record_data.data0.condition_data.condition(1).no = _
         last_conditions.last_cond(1).new_point_no
         n% = 0
          Call set_angle_value(A%, ch$, temp_record, n%, 0, False)
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = angle3_value_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
         If ty = 1 Then
          add_new_value_for_angle0 = start_prove(0, 1, 1)
         End If
End If
End Function
Public Function conclusion_four_point_on_circle(ByVal con_n%) As Byte
Dim tA(3) As Integer
 tA(0) = con_Four_point_on_circle(con_n%).data(0).angle(0)
 tA(1) = con_Four_point_on_circle(con_n%).data(0).angle(1)
 tA(2) = con_Four_point_on_circle(con_n%).data(0).angle(2)
 tA(3) = con_Four_point_on_circle(con_n%).data(0).angle(3)
 Call set_three_angle_value_from_2eangle(tA(0))
 Call set_three_angle_value_from_2eangle(tA(2))
  conclusion_four_point_on_circle = conclusion_two_angle_pi(tA(0), tA(2))
   If conclusion_four_point_on_circle > 1 Then
    Exit Function
   End If
Call set_three_angle_value_from_2eangle(tA(1))
 Call set_three_angle_value_from_2eangle(tA(3))
  conclusion_four_point_on_circle = conclusion_two_angle_pi(tA(1), tA(3))
End Function
Public Function conclusion_two_angle_pi(ByVal A1%, ByVal A2%) As Byte
Dim tn1(1) As Integer
Dim tn2(1) As Integer
Dim dn1_no(1) As Integer
Dim dn2_no(1) As Integer
Dim m1_(2) As Integer
Dim m2_(2) As Integer
Dim dn1() As Integer
Dim dn2() As Integer
Dim dn1_() As Integer
Dim dn2_() As Integer
Dim a3_v(1) As angle3_value_data0_type
Dim temp_record As total_record_type
Dim i%, j%, k%, l%, m%, n%, no1%, no2%
For i% = 0 To 2
 m1_(0) = i%
  m1_(1) = (i% + 1) Mod 3
   m1_(2) = (i% + 2) Mod 3
a3_v(0).angle(m1_(0)) = A1%
a3_v(0).angle(m1_(1)) = -1
Call search_for_three_angle_value(a3_v(0), m1_(0), tn1(0), 1)
a3_v(0).angle(m1_(1)) = 30000
Call search_for_three_angle_value(a3_v(0), m1_(0), tn2(0), 1)
 For j% = 0 To 2
   m2_(0) = j%
    m2_(1) = (j% + 1) Mod 3
     m2_(2) = (j% + 2) Mod 3
  a3_v(1).angle(m2_(0)) = A2%
  a3_v(1).angle(m2_(1)) = -1
  Call search_for_three_angle_value(a3_v(1), m2_(0), tn1(1), 1)
  a3_v(1).angle(m2_(1)) = 30000
  Call search_for_three_angle_value(a3_v(1), m2_(0), tn2(1), 1)
   dn1_no(0) = 0
   dn1_no(1) = 0
   For k% = tn1(0) + 1 To tn2(0)
     no1% = angle3_value(k%).data(0).record.data1.index.i(m1_(0))
      If angle3_value(no1%).data(0).data0.value = "360" Then
      If angle3_value(no1%).data(0).data0.para(m1_(0)) = "2" And _
          angle3_value(no1%).data(0).data0.para(m1_(1)) = "1" And _
           angle3_value(no1%).data(0).data0.para(m1_(2)) = "1" Then
       dn1_no(1) = dn1_no(1) + 1
        ReDim Preserve dn2(dn1_no(1)) As Integer
        dn2(dn1_no(1)) = no1%
      ElseIf (angle3_value(no1%).data(0).data0.para(m1_(0)) = "1" And _
          angle3_value(no1%).data(0).data0.para(m1_(1)) = "-1" And _
           angle3_value(no1%).data(0).data0.para(m1_(2)) = "-1") Or _
       (angle3_value(no1%).data(0).data0.para(m1_(0)) = "-1" And _
          angle3_value(no1%).data(0).data0.para(m1_(1)) = "1" And _
           angle3_value(no1%).data(0).data0.para(m1_(2)) = "1") Then
       dn1_no(0) = dn1_no(0) + 1
       ReDim Preserve dn1(dn1_no(0)) As Integer
        dn1(dn1_no(0)) = no1%
      End If
      End If
   Next k%
   '*********************
      dn2_no(0) = 0
   dn2_no(1) = 0
   For k% = tn1(1) + 1 To tn2(1)
     no1% = angle3_value(k%).data(0).record.data1.index.i(m2_(0))
      If angle3_value(no1%).data(0).data0.value = "360" Then
      If angle3_value(no1%).data(0).data0.para(m2_(0)) = "2" And _
          angle3_value(no1%).data(0).data0.para(m2_(1)) = "1" And _
           angle3_value(no1%).data(0).data0.para(m2_(2)) = "1" Then
       dn2_no(1) = dn2_no(dn2_no(1)) + 1
       ReDim Preserve dn2_(1) As Integer
       dn2_(dn2_no(1)) = no1%
      ElseIf (angle3_value(no1%).data(0).data0.para(m2_(0)) = "1" And _
          (angle3_value(no1%).data(0).data0.para(m2_(1)) = "-1" Or _
            angle3_value(no1%).data(0).data0.para(m2_(1)) = "@1") And _
            (angle3_value(no1%).data(0).data0.para(m2_(2)) = "-1" Or _
              angle3_value(no1%).data(0).data0.para(m2_(2)) = "-1")) Or _
       ((angle3_value(no1%).data(0).data0.para(m2_(0)) = "-1" Or _
          angle3_value(no1%).data(0).data0.para(m2_(0)) = "@1") And _
          angle3_value(no1%).data(0).data0.para(m2_(1)) = "1" And _
           angle3_value(no1%).data(0).data0.para(m2_(2)) = "1") Then
       dn2_no(0) = dn2_no(0) + 1
       ReDim Preserve dn1_(dn2_no(0)) As Integer
       dn1_(dn2_no(0)) = no1%
      End If
      End If
   Next k%
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
     temp_record.record_data.data0.condition_data.condition(2).ty = angle3_value_
   For k% = 1 To dn1_no(0)
    For l% = 1 To dn2_no(0)
     no1% = dn1(k%)
      no2% = dn1_(l%)
     temp_record.record_data.data0.condition_data.condition(1).no = no1%
      temp_record.record_data.data0.condition_data.condition(2).no = no2%
    If add_four_angle_for_four_point_on_circle(angle3_value(no1%).data(0).data0.angle(m1_(1)), _
      angle3_value(no1%).data(0).data0.angle(m1_(2)), angle3_value(no2%).data(0).data0.angle(m2_(1)), _
       angle3_value(no2%).data(0).data0.angle(m2_(2)), 0) = "180" Then
       conclusion_two_angle_pi = set_three_angle_value(angle3_value(no1%).data(0).data0.angle(m1_(0)), _
         angle3_value(no2%).data(0).data0.angle(m2_(0)), 0, "1", "1", "0", "180", 0, temp_record, 0, 0, 0, 0, _
           0, 0, False)
    If conclusion_two_angle_pi > 1 Then
      Exit Function
    End If
    End If
    Next l%
   Next k%
   For k% = 1 To dn1_no(1)
    For l% = 1 To dn2_no(1)
     no1% = dn2(k%)
      no2% = dn2_(l%)
     temp_record.record_data.data0.condition_data.condition(1).no = no1%
      temp_record.record_data.data0.condition_data.condition(2).no = no2%
      If add_four_angle_for_four_point_on_circle(angle3_value(no1%).data(0).data0.angle(m1_(1)), _
      angle3_value(no1%).data(0).data0.angle(m1_(2)), angle3_value(no2%).data(0).data0.angle(m2_(1)), _
       angle3_value(no2%).data(0).data0.angle(m2_(2)), 1) = "360" Then
       conclusion_two_angle_pi = set_three_angle_value(angle3_value(no1%).data(0).data0.angle(m1_(0)), _
         angle3_value(no2%).data(0).data0.angle(m2_(0)), 0, "1", "1", "0", "180", 0, temp_record, 0, 0, 0, 0, _
           0, 0, False)
         If conclusion_two_angle_pi > 1 Then
          Exit Function
         End If
      End If
    If conclusion_two_angle_pi > 1 Then
      Exit Function
    End If
    Next l%
   Next k%
   Next j%
   Next i%
End Function

Public Function add_four_angle_for_four_point_on_circle(ByVal A1%, ByVal A2%, ByVal A3%, _
                                        ByVal A4%, c_ty As Byte) As String
Dim tA(3) As Integer
Dim T_A(1) As Integer
Dim ty As Byte
Dim i%
Dim dn(2) As Integer
Dim cond_ty As Byte
Dim a3_v As angle3_value_data0_type
Dim p_4 As polygon4_data_type
Dim temp_record As total_record_type
tA(0) = A1%
tA(1) = A2%
tA(2) = A3%
tA(3) = A4%
If c_ty = 0 Then 'sum=120
If combine_two_angle(tA(0), tA(2), 0, 0, 0, 0, 0, T_A(0), ty, 0, 0) = True Then
 If ty = 3 Or ty = 5 Then
  If combine_two_angle(tA(1), tA(3), 0, 0, 0, 0, 0, T_A(1), ty, 0, 0) = True Then
   If ty = 3 Or ty = 5 Then
   angle3_value_data0 = a3_v
    If is_three_angle_value(T_A(0), T_A(1), 0, "1", "1", "0", "180", "180", 0, 0, 0, -1000, _
     0, 0, 0, 0, 0, 0, 0, angle3_value_data0, temp_record.record_data.data0.condition_data, 0) Then
       add_four_angle_for_four_point_on_circle = "180"
          Exit Function
    End If
   End If
  End If
 End If
ElseIf combine_two_angle(tA(0), tA(3), 0, 0, 0, 0, 0, T_A(0), ty, 0, 0) = True Then
 If ty = 3 Or ty = 5 Then
  If combine_two_angle(tA(1), tA(2), 0, 0, 0, 0, 0, T_A(1), ty, 0, 0) = True Then
   If ty = 3 Or ty = 5 Then
   angle3_value_data0 = a3_v
    If is_three_angle_value(T_A(0), T_A(1), 0, "1", "1", "0", "180", "180", 0, 0, 0, -1000, _
     0, 0, 0, 0, 0, 0, 0, angle3_value_data0, temp_record.record_data.data0.condition_data, 0) Then
       add_four_angle_for_four_point_on_circle = "180"
          Exit Function
    End If
   End If
  End If
 End If
End If
Else 'c_ty=1
Call is_polygon4_(angle(A1%).data(0).poi(1), angle(A2%).data(0).poi(1), angle(A3%).data(0).poi(1), _
         angle(A4%).data(0).poi(1), p_4, 0, 0, 0)
For i% = 0 To 3
 If p_4.angle(i%) <> A1% And p_4.angle(i%) <> A2% And _
  p_4.angle(i%) <> A3% And p_4.angle(i%) <> A4% Then
   Exit Function
 End If
Next i%
       add_four_angle_for_four_point_on_circle = "360"
          Exit Function
End If
add_four_angle_for_four_point_on_circle = "F"
End Function
Public Function set_three_angle_value_from_2eangle(ByVal A%) As Byte
Dim i%, j%, n1%, n2%
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).eangle_no
 n1% = Deangle.av_no(i%).no
  If angle3_value(n1%).data(0).data0.ty(0) = 3 Or angle3_value(n1%).data(0).data0.ty(0) = 5 Then
   If angle(angle3_value(n1%).data(0).data0.angle(0)).data(0).poi(1) = _
        angle(angle3_value(n1%).data(0).data0.angle(1)).data(0).poi(1) Then
   For j% = i% + 1 To last_conditions.last_cond(1).eangle_no
    n2% = Deangle.av_no(j%).no
    If angle3_value(n2%).data(0).data0.ty(0) = 3 Or angle3_value(n2%).data(0).data0.ty(0) = 5 Then
      If angle(angle3_value(n2%).data(0).data0.angle(0)).data(0).poi(1) = _
        angle(angle3_value(n2%).data(0).data0.angle(1)).data(0).poi(1) Then
         If Abs(angle_number(angle(angle3_value(n1%).data(0).data0.angle(0)).data(0).poi(1), _
              angle(A%).data(0).poi(1), angle(angle3_value(n2%).data(0).data0.angle(0)).data(0).poi(1), _
               "", 0)) = A% Then
         temp_record.record_data.data0.condition_data.condition_no = 2
         temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
         temp_record.record_data.data0.condition_data.condition(2).ty = angle3_value_
         temp_record.record_data.data0.condition_data.condition(1).no = n1%
         temp_record.record_data.data0.condition_data.condition(2).no = n2%
         temp_record.record_data.data0.theorem_no = 1
          set_three_angle_value_from_2eangle = set_three_angle_value(A%, angle3_value(n1%).data(0).data0.angle(3), _
            angle3_value(n2%).data(0).data0.angle(3), "2", "1", "1", "360", 0, temp_record, 0, 0, 0, 0, 0, _
                  0, False)
          Exit Function
         End If
      End If
    End If
 Next j%
   End If
  End If
Next i%
End Function

Public Function conclusion_for_eline(i%)
Dim triA(1) As temp_triangle_type
Dim p1%, p2%, p3%, p4%, k%, j%, l%, tn%
Dim p4_on_circle As four_point_on_circle_data_type
p1% = con_eline(i%).data(0).data0.poi(0)
p2% = con_eline(i%).data(0).data0.poi(1)
p3% = con_eline(i%).data(0).data0.poi(2)
p4% = con_eline(i%).data(0).data0.poi(3)
If is_four_point_on_circle(p1%, p2%, p3%, p4%, tn%, p4_on_circle, False) Then
   conclusion_for_eline = add_two_point_for_mid_point(p1%, p2%, p3%, p4%)
    If conclusion_for_eline > 1 Then
     Exit Function
    End If
End If
 Call set_temp_triangle_from_lin(p1%, p2%, 0, triA(0), True)
  Call set_temp_triangle_from_lin(p3%, p4%, 0, triA(1), False)
   '设置与p1,p2及p3,p4有关的三角形,
For j% = 1 To triA(0).last_T
 For k% = 1 To triA(1).last_T
  If is_equal_angle(triA(0).data(j%).angle(1), triA(1).data(k%).angle(1), 0, 0) Then
   If angle(triA(0).data(j%).angle(2)).data(0).value = "90" Then
    l% = line_number0(triA(1).data(k%).poi(0), triA(1).data(k%).poi(2), 0, 0)
    conclusion_for_eline = add_aid_point_for_paral_or_verti(triA(1).data(k%).poi(1), _
                                                        l%, verti_, l%, 0)
    If conclusion_for_eline > 1 Then
     Exit Function
    End If
   ElseIf angle(triA(1).data(k%).angle(2)).data(0).value = "90" Then
    l% = line_number0(triA(0).data(j%).poi(0), triA(1).data(k%).poi(2), 0, 0)
    conclusion_for_eline = add_aid_point_for_paral_or_verti(triA(0).data(j%).poi(1), _
                                      l%, verti_, l%, 0)
    If conclusion_for_eline > 1 Then
     Exit Function
    End If
   End If
  ElseIf is_equal_angle(triA(0).data(j%).angle(2), triA(1).data(k%).angle(2), 0, 0) Then
   If angle(triA(0).data(j%).angle(1)).data(0).value = "90" Then
    l% = line_number0(triA(1).data(k%).poi(0), triA(1).data(k%).poi(1), 0, 0)
    conclusion_for_eline = add_aid_point_for_paral_or_verti(triA(1).data(k%).poi(2), _
                                                   l%, verti_, l%, 0)
    If conclusion_for_eline > 1 Then
     Exit Function
    End If
   ElseIf angle(triA(1).data(k%).angle(1)).data(0).value = "90" Then
    l% = line_number0(triA(0).data(j%).poi(0), triA(1).data(k%).poi(1), 0, 0)
    conclusion_for_eline = add_aid_point_for_paral_or_verti(triA(0).data(j%).poi(2), _
                                                   l%, verti_, l%, 0)
    If conclusion_for_eline > 1 Then
     Exit Function
    End If
   End If
  End If
 Next k%
Next j%
If m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(0).line_no > 0 Then
 For j% = 1 To m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(0).line_no
  If is_point_in_line3(con_eline(i%).data(0).data0.poi(0), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
     If is_point_in_line3(con_eline(i%).data(0).data0.poi(2), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(3), _
                 verti_, m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, _
                         m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(3), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(2), _
                        verti_, m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, _
                          m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     End If
  ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(1), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
     If is_point_in_line3(con_eline(i%).data(0).data0.poi(2), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(3), _
                               verti_, m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, _
                                 m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(3), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(2), _
                              verti_, m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, _
                                m_lin(con_eline(i%).data(0).data0.line_no(0)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     End If
  End If
 Next j%
End If
If m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(0).line_no > 0 Then
 For j% = 1 To m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(0).line_no
  If is_point_in_line3(con_eline(i%).data(0).data0.poi(2), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
     If is_point_in_line3(con_eline(i%).data(0).data0.poi(0), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(1), _
                             verti_, m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, _
                               m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(1), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
        conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(0), _
                                  verti_, m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, _
                                     m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
     End If
  ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(3), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
     If is_point_in_line3(con_eline(i%).data(0).data0.poi(0), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
         conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(1), _
                                   verti_, m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, _
                                    m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
    ElseIf is_point_in_line3(con_eline(i%).data(0).data0.poi(1), _
        m_lin(m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no).data(0).data0, 0) Then
         conclusion_for_eline = add_aid_point_for_paral_or_verti(con_eline(i%).data(0).data0.poi(0), _
                                      verti_, m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, _
                                        m_lin(con_eline(i%).data(0).data0.line_no(1)).data(0).in_verti(j%).line_no, 0)
        If conclusion_for_eline > 1 Then
         Exit Function
        End If
    End If
  End If
 Next j%
End If
End Function

Public Function conclusion_for_line_value(i%) As Byte
Dim tn(2) As Integer
Dim ts$
Dim ts1$
Dim k%
Dim tp(3) As Integer
Dim l_v0 As line_value_data0_type
Dim temp_record As total_record_type
For k% = 1 To last_conditions.last_cond(1).tixing_no
 If Dpolygon4(Dtixing(i%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
 If is_same_two_point(con_line_value(i%).data(0).data0.poi(0), _
         con_line_value(i%).data(0).data0.poi(1), Dtixing(k%).data(0).poi(0), _
          Dtixing(k%).data(0).poi(2)) Or _
     is_same_two_point(con_line_value(i%).data(0).data0.poi(0), _
         con_line_value(i%).data(0).data0.poi(1), Dtixing(k%).data(0).poi(1), _
          Dtixing(k%).data(0).poi(3)) Then
          temp_record.record_data.data0.condition_data.condition(1).ty = tixing_
          temp_record.record_data.data0.condition_data.condition(1).no = k%
  If is_line_value(Dtixing(k%).data(0).poi(0), _
       Dtixing(k%).data(0).poi(1), 0, 0, 0, "", _
        tn(0), -1000, 0, 0, 0, l_v0) = 1 And _
    is_line_value(Dtixing(k%).data(0).poi(1), _
       Dtixing(k%).data(0).poi(2), 0, 0, 0, "", _
        tn(1), -1000, 0, 0, 0, l_v0) = 1 And _
    is_line_value(Dtixing(k%).data(0).poi(2), _
       Dtixing(k%).data(0).poi(3), 0, 0, 0, "", _
        tn(2), -1000, 0, 0, 0, l_v0) = 1 Then
          temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
          temp_record.record_data.data0.condition_data.condition(2).no = tn(0)
          temp_record.record_data.data0.condition_data.condition(3).ty = line_value_
          temp_record.record_data.data0.condition_data.condition(3).no = tn(1)
          temp_record.record_data.data0.condition_data.condition(4).ty = line_value_
          temp_record.record_data.data0.condition_data.condition(4).no = tn(2)
          temp_record.record_data.data0.condition_data.condition_no = 4
          temp_record.record_data.data0.theorem_no = 1
          ts$ = minus_string(line_value(tn(0)).data(0).data0.value_, line_value(tn(2)).data(0).data0.value_, False, False)
          ts$ = divide_string(ts$, "2", False, False)
          ts$ = time_string(ts$, ts$, False, False)
          ts$ = minus_string(time_string(line_value(tn(1)).data(0).data0.value_, _
                       line_value(tn(1)).data(0).data0.value, False, False), ts$, False, False)
          ts1$ = add_string(line_value(tn(0)).data(0).data0.value_, line_value(tn(2)).data(0).data0.value_, False, False)
          ts1$ = divide_string(ts1$, "2", False, False)
          ts$ = add_string(time_string(ts1$, ts1$, False, False), ts$, False, False)
          ts$ = sqr_string(ts$, True, False)
          conclusion_for_line_value = set_line_value(con_line_value(i%).data(0).data0.poi(0), _
               con_line_value(i%).data(0).data0.poi(1), ts$, con_line_value(i%).data(0).data0.n(0), _
                con_line_value(i%).data(0).data0.n(1), _
                  con_line_value(i%).data(0).data0.line_no, temp_record.record_data, 0, 0, False)
          If conclusion_for_line_value > 1 Then
             Exit Function
          End If
  End If
End If
End If
Next k%
End Function

Public Function conclusion_for_sides_length_of_triangle(ByVal s_tri%) As Byte
Dim tn0(1) As Integer
Dim tn1(1) As Integer
Dim tn2(1) As Integer
Dim tp(2) As Integer
Dim tl(2) As Integer
Dim cond_data(1) As condition_data_type
Dim value(1) As String
Dim i%
Dim temp_record As total_record_type
tp(0) = triangle(con_Sides_length_of_triangle(s_tri%).data(0).triangle).data(0).poi(0)
tp(1) = triangle(con_Sides_length_of_triangle(s_tri%).data(0).triangle).data(0).poi(1)
tp(2) = triangle(con_Sides_length_of_triangle(s_tri%).data(0).triangle).data(0).poi(2)
tl(0) = line_number0(tp(0), tp(1), tn0(0), tn0(1))
tl(1) = line_number0(tp(1), tp(2), tn1(0), tn1(1))
tl(2) = line_number0(tp(2), tp(0), tn2(0), tn2(1))
If Abs(tn0(0) - tn0(1)) > 1 Then
    If tn0(0) > tn0(1) Then
    Call exchange_two_integer(tn0(0), tn0(1))
    End If
    For i% = tn0(0) + 1 To tn0(1) - 1
     If set_two_segment_value(tp(0), m_lin(tl(0)).data(0).data0.in_point(i%), tp(0), tp(2), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(1), m_lin(tl(0)).data(0).data0.in_point(i%), tp(1), tp(2), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
     ElseIf set_two_segment_value(tp(0), m_lin(tl(0)).data(0).data0.in_point(i%), tp(1), tp(2), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(1), m_lin(tl(0)).data(0).data0.in_point(i%), tp(0), tp(2), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
    End If
    Next i%
End If
If Abs(tn1(0) - tn1(1)) > 1 Then
   If tn1(0) > tn1(1) Then
   Call exchange_two_integer(tn1(0), tn1(1))
   End If
    For i% = tn1(0) + 1 To tn1(1) - 1
     If set_two_segment_value(tp(1), m_lin(tl(1)).data(0).data0.in_point(i%), tp(1), tp(0), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(2), m_lin(tl(1)).data(0).data0.in_point(i%), tp(2), tp(0), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
     ElseIf set_two_segment_value(tp(1), m_lin(tl(1)).data(0).data0.in_point(i%), tp(2), tp(0), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(2), m_lin(tl(1)).data(0).data0.in_point(i%), tp(1), tp(0), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
    End If
    Next i%
End If
If Abs(tn2(0) - tn2(1)) > 1 Then
   If tn2(0) > tn2(1) Then
   Call exchange_two_integer(tn2(0), tn2(1))
   End If
    For i% = tn2(0) + 1 To tn2(1) - 1
     If set_two_segment_value(tp(0), m_lin(tl(2)).data(0).data0.in_point(i%), tp(0), tp(1), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(2), m_lin(tl(2)).data(0).data0.in_point(i%), tp(2), tp(1), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
     ElseIf set_two_segment_value(tp(0), m_lin(tl(2)).data(0).data0.in_point(i%), tp(2), tp(1), 0, 0, 0, _
         0, 0, 0, value(0), cond_data(0), 1) And _
        set_two_segment_value(tp(2), m_lin(tl(2)).data(0).data0.in_point(i%), tp(0), tp(1), 0, 0, 0, _
         0, 0, 0, value(1), cond_data(1), 1) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(cond_data(0), temp_record.record_data.data0.condition_data)
         Call add_record_to_record(cond_data(1), temp_record.record_data.data0.condition_data)
         temp_record.record_data.data0.theorem_no = 1
         conclusion_for_sides_length_of_triangle = set_sides_length_of_triangle( _
            con_Sides_length_of_triangle(s_tri%).data(0).triangle, _
              add_string(value(0), value(1), True, False), 0, temp_record, 0)
         If conclusion_for_sides_length_of_triangle > 1 Then
         Exit Function
         End If
    End If
    Next i%
End If
End Function

Public Function set_two_segment_value(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
        ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, ByVal l1%, ByVal l2%, _
             value$, cond_data As condition_data_type, ty As Byte) As Boolean
Dim e_line As eline_data0_type
Dim l2_v As two_line_value_data0_type
Dim i%, k%, l%, last_tn%, no%
Dim tn() As Integer
Dim n_(1) As Integer
Dim m(1) As Integer
Dim cond_ty As Byte
If ty = 1 Then
l1% = line_number0(p1%, p2%, n1%, n2%)
 If n1% > n2% Then
 Call exchange_two_integer(n1%, n2%)
 Call exchange_two_integer(p1%, p2%)
 End If
 l2% = line_number0(p3%, p4%, n3%, n4%)
 If n1% > n2% Then
 Call exchange_two_integer(n3%, n4%)
 Call exchange_two_integer(p3%, p4%)
 End If
End If
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
e_line.poi(2 * m(0)) = p1%
e_line.poi(2 * m(0) + 1) = p2%
e_line.poi(2 * m(1)) = -1
Call search_for_eline(e_line, m(0), n_(0), 1)
e_line.poi(2 * m(1)) = 30000
Call search_for_eline(e_line, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 l% = Deline(k%).data(0).record.data1.index.i(m(0))
 last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = l%
Next k%
For k% = 1 To last_tn%
l% = tn(k%)
value$ = ""
cond_data.condition_no = 0
no% = 0
If is_two_line_value(Deline(l%).data(0).data0.poi(2 * m(1)), Deline(l%).data(0).data0.poi(2 * m(1) + 1), _
       p3%, p4%, Deline(l%).data(0).data0.n(2 * m(1)), Deline(l%).data(0).data0.n(2 * m(1) + 1), _
        n3%, n4%, Deline(l%).data(0).data0.line_no(m(1)), l2%, "1", "1", value$, no%, -1000, _
         0, 0, 0, l2_v, cond_ty, cond_data) = 1 Then
         cond_data.condition_no = 2
         cond_data.condition(1).ty = cond_ty
         cond_data.condition(1).no = no%
         cond_data.condition(2).ty = eline_
         cond_data.condition(2).no = l%
         value$ = line_value(no%).data(0).data0.value_
          set_two_segment_value = True
           Exit Function
End If
Next k%
Next i%

If ty = 1 Then
value$ = ""
cond_data.condition_no = 0
set_two_segment_value = set_two_segment_value(p3%, p4%, p1%, p2%, _
         n3%, n4%, n1%, n2%, l2%, l1%, _
             value$, cond_data, 0)
End If
End Function
Public Function add_point_for_general_string_satis_midpoint(ByVal gs%)
Dim tl(1) As Integer
Dim ty(2) As Boolean
Dim sig As String
Dim c_data As condition_data_type
 If general_string(gs%).record_.conclusion_no > 0 Then '结论
  If general_string(gs%).data(0).value = "" And general_string(gs%).data(0).para(2) <> "0" Then '三项
   If (item0(general_string(gs%).data(0).item(0)).data(0).sig = "~" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "~" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "~") Or _
     (item0(general_string(gs%).data(0).item(0)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "/") Then
     If item0(general_string(gs%).data(0).item(0)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(1)).data(0).sig = "/" And _
     item0(general_string(gs%).data(0).item(2)).data(0).sig = "/" Then
     'If item0(general_string(gs%).data(0).item(0)).data(0).sig = "/" Then '比
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
     If ty(0) = False And ty(1) = False And ty(2) = False Then
      Exit Function
     End If
     ' p_1% = item0(general_string(gs%).data(0).item(0)).data(0).poi(2)
     ' p_2% = item0(general_string(gs%).data(0).item(0)).data(0).poi(3)
     'ElseIf ty(1) = True Then
      'p_1% = item0(general_string(gs%).data(0).item(1)).data(0).poi(2)
      'p_2% = item0(general_string(gs%).data(0).item(1)).data(0).poi(3)
     'ElseIf ty(2) = True Then
      'p_1% = item0(general_string(gs%).data(0).item(2)).data(0).poi(2)
      'p_2% = item0(general_string(gs%).data(0).item(2)).data(0).poi(3) '分母
     'End If
     Else
     ty(0) = True
     ty(1) = True
     ty(2) = True
     sig = "~"
  End If
  If sig = "~" Then
    If is_three_point_on_line(item0(general_string(gs%).data(0).item(0)).data(0).poi(1), _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(1), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(1), _
              0, 0, 0, 0, 0, 0, 0) Then
    If general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(1) Then
       add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(1) = general_string(gs%).data(0).para(2) Then
       add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(2) Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(2)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
   End If
    ElseIf is_three_point_on_line(item0(general_string(gs%).data(0).item(0)).data(0).poi(0), _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(0), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(0), _
             0, 0, 0, 0, 0, 0, 0) Then
    If general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(1) = 1 Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(1) = general_string(gs%).data(0).para(2) Then
         add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
   ElseIf general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(2) Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    End If
    End If
  Else
    If is_three_point_on_line(item0(general_string(gs%).data(0).item(0)).data(0).poi(1), _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(1), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(1), _
             0, 0, 0, 0, 0, 0, 0) Then
    If general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(1) And _
          ty(0) Then
       add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(1) = general_string(gs%).data(0).para(2) And _
          ty(1) Then
       add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(2) And _
          ty(2) Then
       add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(0), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(0), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    End If
    ElseIf is_three_point_on_line(item0(general_string(gs%).data(0).item(0)).data(0).poi(0), _
           item0(general_string(gs%).data(0).item(1)).data(0).poi(0), _
            item0(general_string(gs%).data(0).item(2)).data(0).poi(0), _
             0, 0, 0, 0, 0, 0, 0) Then
    If general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(1) And _
          ty(0) Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(1) = general_string(gs%).data(0).para(2) And _
          ty(1) Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(2)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(1)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    ElseIf general_string(gs%).data(0).para(0) = general_string(gs%).data(0).para(2) And _
          ty(2) Then
        add_point_for_general_string_satis_midpoint = add_mid_point( _
           item0(general_string(gs%).data(0).item(0)).data(0).poi(1), 0, _
              item0(general_string(gs%).data(0).item(2)).data(0).poi(1), 0)
       If add_point_for_general_string_satis_midpoint > 1 Then
          Exit Function
       End If
    End If
    End If
  End If
 End If
End If
End If
End Function
Public Sub add_condition_for_relation(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal v$)
Dim n%
Dim temp_record As total_record_type
If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 'record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1445, _
        "\\1\\" + m_poi(p1%).data(0).data0.name + m_poi(p2%).data(0).data0.name + "/" + _
           m_poi(p3%).data(0).data0.name + m_poi(p4%).data(0).data0.name + "=" + v$)
        n% = 0
   Call set_Drelation(p1%, p2%, p3%, p4%, 0, 0, 0, 0, 0, 0, _
         v$, temp_record, n%, 0, 0, 0, 0, False)
     new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = new_point_
     new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
 End Sub

Public Sub set_drelation_for_add(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, v As String)
Dim n%
Dim tp(3) As Integer
Dim tn(3) As Integer
Dim tl(1) As Integer
Dim temp_record1 As total_record_type
Call arrange_four_point(p1%, p2%, p3%, p4%, 0, 0, 0, 0, 0, 0, _
      tp(0), tp(1), tp(2), tp(3), 0, 0, tn(0), tn(1), tn(2), tn(3), _
        0, 0, tl(0), tl(1), 0, 0, record_0.data0.condition_data, 0)
 If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
 End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record1.record_data.data0.condition_data.condition_no = 1 'record0
    temp_record1.record_data.data0.condition_data.condition(1).ty = new_point_
     temp_record1.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
                      v = next_char(0, "", 0, 0)
                      temp_record1.record_data.data0.condition_data.condition_no = 254
             n% = 0
      Call set_Drelation(tp(0), tp(1), tp(2), tp(3), tn(0), tn(1), tn(2), tn(3), tl(0), tl(1), v, temp_record1, n%, 0, 0, 0, 0, False)
         Call next_char(0, v, relation_, n%)
         new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
           LoadResString_(1445, "\\1\\\" + _
           m_poi(tp(0)).data(0).data0.name + m_poi(tp(1)).data(0).data0.name + "/" + _
            m_poi(tp(2)).data(0).data0.name + m_poi(tp(3)).data(0).data0.name + "=" + v)
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = relation_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n% 'last_conditions.last_cond(1).new_point_no
End Sub

Public Function add_point_from_aid_point_data(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
             ByVal p5%, ByVal p6%) As Byte 'p1%p2%+p2$p3%=p4%p5%
Dim i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data1_no
If is_same_two_point(aid_point_data1(i%).data(0).triA(0).poi(0), aid_point_data1(i%).data(0).triA(0).poi(1), _
     p3%, p4%) Then
 If is_same_two_point(aid_point_data1(i%).data(0).triA(1).poi(0), aid_point_data1(i%).data(0).triA(1).poi(1), _
     p1%, p2%) Then
      GoTo add_point_from_aid_point_data_mark0
 Else
     GoTo add_point_from_aid_point_data_next0
 End If
ElseIf is_same_two_point(aid_point_data1(i%).data(0).triA(0).poi(0), aid_point_data1(i%).data(0).triA(0).poi(1), _
     p1%, p2%) Then
  If is_same_two_point(aid_point_data1(i%).data(0).triA(1).poi(0), aid_point_data1(i%).data(0).triA(1).poi(1), _
     p3%, p4%) Then
    GoTo add_point_from_aid_point_data_mark0
  Else
     GoTo add_point_from_aid_point_data_next0
  End If

Else
 GoTo add_point_from_aid_point_data_next0
End If
add_point_from_aid_point_data_mark0:
add_point_from_aid_point_data = add_aid_point_for_eline(aid_point_data1(i%).data(0).triA(0).poi(0), _
        aid_point_data1(i%).data(0).triA(0).poi(1), aid_point_data1(i%).data(0).triA(1).poi(0), _
         aid_point_data1(i%).data(1).triA(1).poi(1), aid_point_data1(i%).data(1).triA(1).poi(1))
If add_point_from_aid_point_data > 1 Then
 Exit Function
End If
       '在直线p3%p4%上取点p使得p3%p=p1%p2%
add_point_from_aid_point_data_next0:
Next i%
For i% = 1 To last_conditions.last_cond(1).aid_point_data2_no
If is_same_two_point(aid_point_data1(i%).data(0).triA(0).poi(0), aid_point_data1(i%).data(0).triA(0).poi(1), _
     p5%, p6%) Then
 If is_same_two_point(aid_point_data1(i%).data(0).triA(1).poi(0), aid_point_data1(i%).data(0).triA(1).poi(1), _
     p1%, p2%) Then
     GoTo add_point_from_aid_point_data_mark1
 ElseIf is_same_two_point(aid_point_data1(i%).data(0).triA(1).poi(0), aid_point_data1(i%).data(0).triA(1).poi(1), _
     p3%, p4%) Then
     GoTo add_point_from_aid_point_data_mark1
 Else
    GoTo add_point_from_aid_point_data_next1
 End If
     
ElseIf is_same_two_point(aid_point_data1(i%).data(0).triA(1).poi(0), aid_point_data1(i%).data(0).triA(1).poi(1), _
     p5%, p6%) Then
  If is_same_two_point(aid_point_data1(i%).data(0).triA(0).poi(0), aid_point_data1(i%).data(0).triA(0).poi(1), _
     p1%, p2%) Then
     GoTo add_point_from_aid_point_data_mark1
  ElseIf is_same_two_point(aid_point_data1(i%).data(0).triA(0).poi(0), aid_point_data1(i%).data(0).triA(0).poi(1), _
     p3%, p4%) Then
     GoTo add_point_from_aid_point_data_mark1
  Else
    GoTo add_point_from_aid_point_data_next1
  End If
Else
 GoTo add_point_from_aid_point_data_next1
End If
add_point_from_aid_point_data_mark1:
add_point_from_aid_point_data = add_aid_point_for_eline(aid_point_data1(i%).data(0).triA(0).poi(0), _
        aid_point_data1(i%).data(0).triA(0).poi(1), aid_point_data1(i%).data(0).triA(1).poi(0), _
         aid_point_data1(i%).data(1).triA(1).poi(1), aid_point_data1(i%).data(1).triA(1).poi(0))
If add_point_from_aid_point_data > 1 Then
 Exit Function
End If
add_point_from_aid_point_data_next1:
       '在直线p3%p4%上取点p使得p3%p=p1%p2%
Next i%

End Function
Public Function add_point_for_chord(ByVal p1%, ByVal p2%, ByVal c%) As Byte
Dim i%, j%
Dim k%
If c% = 0 Then
For i% = 1 To C_display_picture.m_circle.Count
 k% = 0
  If m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.visible > 0 And m_Circ(i%).data(0).data0.center > 0 Then
 For j% = 1 To m_Circ(i%).data(0).data0.in_point(0)
  If m_Circ(i%).data(0).data0.in_point(j%) = p1% Or m_Circ(i%).data(0).data0.in_point(j%) = p2% Then
   k% = k% + 1
    If k% = 2 Then
     add_point_for_chord = add_mid_point(p1%, 0, p2%, 1)
      Exit Function
    End If
  End If
  Next j%
 End If
Next i%
Else
  If m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.visible > 0 Then
 For j% = 1 To m_Circ(c%).data(0).data0.in_point(0)
  If m_Circ(c%).data(0).data0.in_point(j%) = p1% Or m_Circ(c%).data(0).data0.in_point(j%) = p2% Then
   k% = k% + 1
    If k% = 2 Then
     add_point_for_chord = add_mid_point(p1%, 0, p2%, 1)
      Exit Function
    End If
  End If
  Next j%
 End If
End If
End Function

Public Function add_point_for_circle_center_angle(ByVal A%) As Byte
Dim i%, k%, dr%
Dim p(1) As Integer
Dim tp(1) As Integer
For i% = 0 To C_display_picture.m_circle.Count
 If m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.visible > 0 Then
  If m_Circ(i%).data(0).data0.center = angle(A%).data(0).poi(1) Then
   k% = inter_point_line_circle0(m_lin(angle(A%).data(0).line_no(0)).data(0).data0, _
     m_Circ(i%).data(0).data0, tp(0), tp(1))
   If k% = 0 Then
    Exit Function
   ElseIf k% = 1 Then
    p(0) = tp(0)
   Else
    dr% = compare_two_point(m_poi(angle(A%).data(0).poi(1)).data(0).data0.coordinate, _
                   m_poi(tp(0)).data(0).data0.coordinate, angle(A%).data(0).poi(1), _
                     angle(A%).data(0).poi(0), 6)
    If dr% = 1 Then
         p(0) = tp(0)
    ElseIf dr% = -1 Then
         p(0) = tp(1)
    Else
     Exit Function
    End If
   End If
   '****************************************************
   k% = inter_point_line_circle0(m_lin(angle(A%).data(0).line_no(1)).data(0).data0, _
     m_Circ(i%).data(0).data0, tp(0), tp(1))
   If k% = 0 Then
    Exit Function
   ElseIf k% = 1 Then
    p(1) = tp(0)
   Else
    dr% = compare_two_point(m_poi(angle(A%).data(0).poi(1)).data(0).data0.coordinate, _
            m_poi(tp(0)).data(0).data0.coordinate, angle(A%).data(0).poi(1), _
               angle(A%).data(0).poi(2), 6)
    If dr% = 1 Then
         p(1) = tp(0)
    ElseIf dr% = -1 Then
         p(1) = tp(1)
    Else
         Exit Function
    End If
   End If
    add_point_for_circle_center_angle = add_point_for_chord(p(0), p(1), i%)
     Exit Function
 End If
 End If
Next i%
End Function
Public Function add_point_for_point_pair(dp As point_pair_data0_type, is_remove_add_point As Boolean) As Byte
Dim tp As point_pair_data0_type
tp = dp
If tp.line_no(0) = tp.line_no(3) Then
 add_point_for_point_pair = add_point_for_chord_tangent_line(tp.poi(0), tp.poi(1), tp.poi(6), tp.poi(7), _
                                                     is_remove_add_point)
   If add_point_for_point_pair > 1 Then
       Exit Function
   End If
End If
If tp.line_no(1) = tp.line_no(2) Then
 add_point_for_point_pair = add_point_for_chord_tangent_line(tp.poi(2), tp.poi(3), tp.poi(4), tp.poi(5), _
                               is_remove_add_point)
End If

End Function

Public Function add_point_for_chord_tangent_line(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                                               is_remove_add_point As Boolean) As Byte
'弦切
Dim no%
Dim tp(2) As Integer
Dim inter_p(1) As Integer
Dim ty As Byte
Dim c%, i%, k%
Dim l(2) As Integer
Dim t_p(1) As POINTAPI
Dim tangent_point  As POINTAPI
Dim circ_data  As circle_data0_type
Dim temp_record  As total_record_type
'On Error GoTo add_point_for_chord_tangent_line_error
If p1% = p3% And p2% <> p4% Then
   tp(0) = p1%
   tp(1) = p2%
   tp(2) = p4%
   ty = 0
   '割线定理
ElseIf p1% = p4% Then
   tp(0) = p1%
   tp(1) = p2%
   tp(2) = p3%
   ty = 1
   '相交弦定理
ElseIf p2% = p3% Then
   tp(0) = p2%
   tp(1) = p1%
   tp(2) = p4%
   ty = 1
ElseIf p2% = p4% Then
   tp(0) = p2%
   tp(1) = p1%
   tp(2) = p3%
   ty = 0
Else
   Exit Function
End If
c% = read_circle_from_chord(tp(1), tp(2), 0)
If c% = 0 Then
   Exit Function
End If
'由弦得圆
If ty = 0 Then
 For i% = 1 To last_conditions.last_cond(1).tangent_line_no
 If tangent_line(i%).data(0).ele(0).no = c% Or tangent_line(i%).data(0).ele(1).no = c% Then
  If is_point_in_line3(tp(0), m_lin(tangent_line(i%).data(0).line_no).data(0).data0, 0) Then
    '在切线上
        Exit Function
  End If
 End If
Next i%
circ_data.c_coord = mid_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, _
                                        m_Circ(c%).data(0).data0.c_coord)
'圆心位置 旧到弦的端点的中点
circ_data.radii = abs_POINTAPI(minus_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, _
                                circ_data.c_coord))
'
Call inter_point_circle_circle_(m_Circ(c%).data(0).data0, circ_data, tangent_point, 0, t_p(0), 0, 0, 0, False)
If from_old_to_aid = 1 Then
   Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
If last_conditions.last_cond(1).point_no = 26 Then
 add_point_for_chord_tangent_line = 6
  Exit Function
End If
     Call set_point_name(last_conditions.last_cond(1).point_no, _
       next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
Call set_point_coordinate(last_conditions.last_cond(1).point_no, tangent_point, False)
Call add_point_to_m_circle( _
                last_conditions.last_cond(1).point_no, c%, temp_record, 255)
If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
End If
 last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 'record0
   temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
   temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
        'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
         new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_circle(0) = c%
If m_Circ(c%).data(0).data0.center > 0 Then
 new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
LoadResString_(1535, "\\1\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
                    "\\2\\" + set_display_circle0(m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + "(" + _
                              m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + ")") + _
                    "\\3\\" + m_poi(tp(0)).data(0).data0.name + _
                              m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
Else
 new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
LoadResString_(1535, "\\1\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name + _
                    "\\2\\" + set_display_circle0(m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + _
                              m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + _
                              m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).data0.name) + _
                    "\\3\\" + m_poi(tp(0)).data(0).data0.name + _
                              m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
End If
no% = 0
  add_point_for_chord_tangent_line = set_tangent_line(line_number0(tp(0), last_conditions.last_cond(1).point_no, 0, 0), _
      last_conditions.last_cond(1).point_no, c%, 0, 0, temp_record, no%, 0)
      new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = tangent_line_
      new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = no%
   If add_point_for_chord_tangent_line > 1 Then
    Exit Function
   End If
   add_point_for_chord_tangent_line = start_prove(0, 1, 1)
     If add_point_for_chord_tangent_line > 1 Then
      Exit Function
     End If
   'If new_result_from_add = False Then
   If is_remove_add_point = True Then
    Call from_aid_to_old
   End If
 Else
 l(0) = line_number0(m_Circ(c%).data(0).data0.center, tp(0), 0, 0)
 k% = 0
 For i% = 1 To m_Circ(c%).data(0).data0.in_point(0)
  l(1) = line_number0(m_Circ(c%).data(0).data0.in_point(i%), tp(0), 0, 0)
   If k% = 0 Then
    If is_dverti(l(0), l(1), 0, -1000, 0, 0, 0, 0) Then
     inter_p(k%) = m_Circ(c%).data(0).data0.in_point(i%)
      k% = k% + 1
       If k% = 1 Then
       l(2) = l(1)
       End If
    End If
   Else
    If l(1) = l(2) Then
     inter_p(k%) = m_Circ(c%).data(0).data0.in_point(i%)
      Exit Function
    End If
   End If
 Next i%
 Call inter_point_line_circle3(m_poi(tp(0)).data(0).data0.coordinate, False, _
            m_poi(tp(0)).data(0).data0.coordinate, _
              m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate, _
               m_Circ(c%).data(0).data0, t_p(0), 0, t_p(1), 0, 0, False)
  If inter_p(0) > 0 Then
   If Abs(m_poi(inter_p(0)).data(0).data0.coordinate.X - t_p(0).X) < 5 And _
      Abs(m_poi(inter_p(0)).data(0).data0.coordinate.Y - t_p(0).Y) < 5 Then
       t_p(0).X = t_p(1).X
        t_p(0).Y = t_p(1).Y
   End If
  End If
If inter_p(0) > 0 Then
Else
If from_old_to_aid = 1 Then
   Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
If last_conditions.last_cond(1).point_no = 26 Then
 add_point_for_chord_tangent_line = 6
  Exit Function
End If
     Call set_point_name(last_conditions.last_cond(1).point_no, _
        next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
     Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_p(0), False)
     Call add_point_to_m_circle( _
                 last_conditions.last_cond(1).point_no, c%, temp_record, 255)
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
If last_conditions.last_cond(1).point_no = 26 Then
 add_point_for_chord_tangent_line = 6
  Exit Function
End If
     Call set_point_name(last_conditions.last_cond(1).point_no, _
       next_char(last_conditions.last_cond(1).point_no, "", 0, 0))
Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_p(1), False)
Call add_point_to_m_circle(last_conditions.last_cond(1).point_no, c%, 255)
l(1) = line_number0(last_conditions.last_cond(1).point_no, last_conditions.last_cond(1).point_no - 1, 0, 0)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(tp(0), l(1), 0, False, False, 0, temp_record)
If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
End If
 last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 ' last_conditions.last_cond(1).new_point_no 'record0
   temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
   temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no - 1
       new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(1) = last_conditions.last_cond(1).point_no
       'new_point(last_conditions.last_cond(1).new_point_no).data(0).record = temp_record.record_data
        new_point(last_conditions.last_cond(1).new_point_no).data(0).add_to_circle(0) = c%
        ' poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
         'poi(p2%).data(0).data0.visible
If m_Circ(c%).data(0).data0.center > 0 Then
new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1555, "\\1\\" + m_poi(tp(0)).data(0).data0.name + _
                        "\\2\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                                  m_poi(tp(0)).data(0).data0.name + _
                        "\\3\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + "(" + _
                                  m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + ")" + _
                        "\\4\\" + m_poi(last_conditions.last_cond(1).point_no - 1).data(0).data0.name + _
                                  m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
Else
new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1555, "\\1\\" + m_poi(tp(0)).data(0).data0.name + _
                       "\\2\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                                 m_poi(tp(0)).data(0).data0.name + _
                       "\\3\\" + m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + _
                                 m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + _
                                 m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).data0.name + _
                       "\\4\\" + m_poi(last_conditions.last_cond(1).point_no - 1).data(0).data0.name + _
                                 m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name)
End If
no% = 0
 add_point_for_chord_tangent_line = set_dverti(l(0), l(1), temp_record, _
          no%, 0, False)
   If add_point_for_chord_tangent_line > 1 Then
    Exit Function
   End If
        new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = verti_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = no%
   add_point_for_chord_tangent_line = set_New_point(last_conditions.last_cond(1).point_no, temp_record, 0, _
           0, 0, 0, c%, 0, 0, 1)
   If add_point_for_chord_tangent_line > 1 Then
    Exit Function
   End If
   add_point_for_chord_tangent_line = start_prove(0, 1, 1)
     If add_point_for_chord_tangent_line > 1 Then
      Exit Function
     End If
   'If new_result_from_add = False Then
   If is_remove_add_point = True Then
    Call from_aid_to_old
   End If
End If
End If
add_point_for_chord_tangent_line_error:
End Function

Public Function read_circle_from_chord(ByVal p1%, ByVal p2%, s%) As Integer
Dim i%, j%, k%
For i% = s% + 1 To C_display_picture.m_circle.Count
 For j% = 1 To m_Circ(i%).data(0).data0.in_point(0)
  If m_Circ(i%).data(0).data0.in_point(j%) = p1% Or m_Circ(i%).data(0).data0.in_point(j%) = p2% Then
   k% = k% + 1
    If k% = 2 Then
     read_circle_from_chord = i%
      Exit Function
    End If
  End If
 Next j%
Next i%
End Function
Private Function add_aid_point_from_tri_function_to_Rt(ByVal A%, ty) As Byte
Dim i%, j%
Dim tl(1) As Integer
Dim c_data0 As condition_data_type
For i% = 0 To 1
 For j% = 1 To last_conditions.last_cond(1).verti_no
  If Dverti(j%).data(0).line_no(0) = angle(A%).data(0).line_no(i%) Then
   If Dverti(j%).data(0).inter_poi > 0 Then
      If is_line_line_intersect(angle(A%).data(0).line_no(1), Dverti(j%).data(0).line_no(1), 0, 0, False) = 0 Then
        add_aid_point_from_tri_function_to_Rt = add_interset_point_line_line( _
         angle(A%).data(0).line_no(1), Dverti(j%).data(0).line_no(1), 0, 1, 0, 0, 0, c_data0)
      End If
   Else
      If is_line_line_intersect(angle(A%).data(0).line_no(1), Dverti(j%).data(0).line_no(1), 0, 0, False) = 0 Then
      End If
   End If
  ElseIf Dverti(j%).data(0).line_no(1) = angle(A%).data(0).line_no(i%) Then
   If Dverti(j%).data(0).inter_poi > 0 Then
      If is_line_line_intersect(angle(A%).data(0).line_no(1), Dverti(j%).data(0).line_no(1), 0, 0, False) = 0 Then
      End If
   Else
      If is_line_line_intersect(angle(A%).data(0).line_no(1), Dverti(j%).data(0).line_no(1), 0, 0, False) = 0 Then
      End If
   End If
  End If
 Next j%
Next i%
End Function

Public Function add_aid_point_from_conclusion(con_no%, unkown_number As Byte) As Byte
Dim i%, n%, k%, j%, tn_%, tl%, p%
Dim t_l(2) As Integer
Dim t_n(3) As Integer
Dim tp(2) As Integer
Dim para(1) As String
Dim unkown_char As String
Dim temp_record As total_record_type
Dim c_data0 As condition_data_type
'On Error GoTo add_aid_point_from_conclusion_error
If regist_data.run_type = 0 Then
  If conclusion_data(con_no%).ty = length_of_polygon_ Then
     For i% = 1 To last_conditions.last_cond(1).length_of_polygon_no
         For j% = 1 To length_of_polygon(i%).data(0).last_segment - 1
          If length_of_polygon(i%).data(0).segment(j%).line_no <> _
              length_of_polygon(i%).data(0).segment(j% + 1).line_no Or _
               length_of_polygon(i%).data(0).segment(j%).poi(1) <> _
                length_of_polygon(i%).data(0).segment(j% + 1).poi(0) Or _
                 length_of_polygon(i%).data(0).segment(j%).para <> _
                  length_of_polygon(i%).data(0).segment(j% + 1).para Then
                  GoTo add_aid_point_from_conclusion_next1
          End If
        Next j%
         add_aid_point_from_conclusion = add_new_value_for_line0( _
             length_of_polygon(i%).data(0).segment(1).poi(0), _
               length_of_polygon(i%).data(0).segment(length_of_polygon(i%).data(0).last_segment).poi(1), "x", 1, 0)
         If add_aid_point_from_conclusion > 1 Then
            Exit Function
         End If
add_aid_point_from_conclusion_next1:
     Next i%
  ElseIf conclusion_data(con_no%).ty = line_value_ Then
     add_aid_point_from_conclusion = conclusion_for_line_value(con_no%)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
   add_aid_point_from_conclusion = add_point_for_chord(con_line_value(con_no%).data(0).data0.poi(0), _
        con_line_value(con_no%).data(0).data0.poi(1), 0)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
'***********
  If unkown_number = 0 Then
     unkown_number = 1
   If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
   End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 'record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
       temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1345, _
         "\\1\\" + m_poi(con_line_value(con_no%).data(0).data0.poi(0)).data(0).data0.name + _
         m_poi(con_line_value(con_no%).data(0).data0.poi(1)).data(0).data0.name + "= x")
    n% = 0
   Call set_line_value(con_line_value(con_no%).data(0).data0.poi(0), con_line_value(con_no%).data(0).data0.poi(1), _
         "x", con_line_value(con_no%).data(0).data0.n(0), con_line_value(con_no%).data(0).data0.n(1), _
         con_line_value(con_no%).data(0).data0.line_no, temp_record.record_data, n%, 0, False)
       new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = line_value_
       new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
  add_aid_point_from_conclusion = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
   End If
   '**************
  For p% = 1 + last_conditions.last_cond(0).line_value_no To last_conditions.last_cond(1).line_value_no
   i% = line_value(p%).data(0).record.data1.index.i(0)
    If is_contain_x(line_value(i%).data(0).data0.value, "x", 1) = False Then
    If con_line_value(k%).data(0).data0.line_no = line_value(i%).data(0).data0.line_no Then
     If line_value(i%).data(0).data0.poi(1) = con_line_value(k%).data(0).data0.poi(0) Or _
      line_value(i%).data(0).data0.poi(0) = con_line_value(k%).data(0).data0.poi(0) Then
       tp(0) = line_value(i%).data(0).data0.poi(0)
        tp(1) = line_value(i%).data(0).data0.poi(1)
         tp(2) = con_line_value(k%).data(0).data0.poi(1)
     ElseIf line_value(i%).data(0).data0.poi(1) = con_line_value(k%).data(0).data0.poi(1) Or _
             line_value(i%).data(0).data0.poi(0) = con_line_value(k%).data(0).data0.poi(1) Then
      tp(0) = line_value(i%).data(0).data0.poi(0)
       tp(1) = line_value(i%).data(0).data0.poi(1)
        tp(2) = con_line_value(k%).data(0).data0.poi(0)
    Else
     GoTo add_aid_point_mark3
    End If
    add_aid_point_from_conclusion = add_aid_point2(tp(0), tp(1), tp(2), tl%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
     End If
    End If
add_aid_point_mark3:
   End If
Next p%
   '****************
  ElseIf conclusion_data(con_no%).ty = two_line_value_ Then
    If is_dparal(con_two_line_value(con_no%).data(0).line_no(0), con_two_line_value(con_no%).data(0).line_no(1), _
                 0, -1000, 0, 0, 0, 0) Then
       If con_two_line_value(con_no%).data(0).para(0) = "1" And _
            con_two_line_value(con_no%).data(0).para(1) = "1" Then
          add_aid_point_from_conclusion = add_mid_point(con_two_line_value(con_no%).data(0).poi(0), 0, _
                    con_two_line_value(con_no%).data(0).poi(2), 1)
             If add_aid_point_from_conclusion > 1 Then
              Exit Function
             End If
          add_aid_point_from_conclusion = add_mid_point(con_two_line_value(con_no%).data(0).poi(1), 0, _
                    con_two_line_value(con_no%).data(0).poi(3), True)
             If add_aid_point_from_conclusion > 1 Then
              Exit Function
             End If
       ElseIf con_two_line_value(con_no%).data(0).para(0) = "1" And _
            con_two_line_value(con_no%).data(0).para(1) = "-1" Then
          add_aid_point_from_conclusion = add_mid_point(con_two_line_value(con_no%).data(0).poi(0), 0, _
                    con_two_line_value(con_no%).data(0).poi(3), 1)
             If add_aid_point_from_conclusion > 1 Then
              Exit Function
             End If
          add_aid_point_from_conclusion = add_mid_point(con_two_line_value(con_no%).data(0).poi(1), 0, _
                    con_two_line_value(con_no%).data(0).poi(2), 0)
             If add_aid_point_from_conclusion > 1 Then
              Exit Function
             End If
       End If
    End If
  ElseIf conclusion_data(con_no%).ty = area_of_element_ Then
  For i% = 1 To last_conditions.last_cond(1).tixing_no
      If Dtixing(i%).data(0).area_value_no = 0 Then
         If Dtixing(i%).data(0).line_value_no(0) > 0 And Dtixing(i%).data(0).line_value_no(2) > 0 Then
            If Dtixing(i%).data(0).line_value_no(1) > 0 Or Dtixing(i%).data(0).line_value_no(3) > 0 Then
             add_aid_point_from_conclusion = add_paral_line(Dtixing(i%).data(0).poi(0), _
                line_number0(Dtixing(i%).data(0).poi(1), Dtixing(i%).data(0).poi(2), 0, 0), _
                 line_number0(Dtixing(i%).data(0).poi(2), Dtixing(i%).data(0).poi(3), 0, 0), _
                   0, 0, 0, 0, 0, 0)
                   If add_aid_point_from_conclusion > 1 Then
                      Exit Function
                   End If
            End If
         End If
      End If
  Next i%
  If unkown_number = 0 Then
     unkown_number = 1
  If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
  End If
   last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
   temp_record.record_data.data0.condition_data.condition_no = 1 'record0
    temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
       new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
        If con_Area_of_element(con_no%).data(0).element.ty = triangle_ Then
          new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(475, _
            set_triangle_display_string(con_Area_of_element(con_no%).data(0).element.no, 1, no_display, False, 1, 0)) + _
            "= x"
             n% = 0
               Call set_area_of_triangle(con_Area_of_element(con_no%).data(0).element.no, _
                "x", temp_record, n%, 0)
               new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = new_point_
               new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
      Else 'con_Area_of_element(con_no%).data(0).element.ty = polygon_
         new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(420, _
                set_display_polygon4(Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0), 0, False, 1, 0)) + _
                 "= x"
              n% = 0
               Call set_area_of_polygon0(con_Area_of_element(con_no%).data(0).element.no, _
                "x", temp_record, n%, 0)
               new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = new_point_
              new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
      End If
  add_aid_point_from_conclusion = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
   End If
  
  ElseIf conclusion_data(con_no%).ty = eline_ Then
   add_aid_point_from_conclusion = conclusion_for_eline(con_no%)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
   add_aid_point_from_conclusion = add_point_for_chord(con_eline(con_no%).data(0).data0.poi(0), _
       con_eline(con_no%).data(0).data0.poi(1), 0)
    If add_aid_point_from_conclusion > 1 Then
     Exit Function
    End If
   add_aid_point_from_conclusion = add_point_for_chord(con_eline(con_no%).data(0).data0.poi(2), _
       con_eline(con_no%).data(0).data0.poi(3), 0)
    If add_aid_point_from_conclusion > 1 Then
     Exit Function
    End If
  ElseIf conclusion_data(con_no%).ty = tixing_ Then
   add_aid_point_from_conclusion = add_interset_point_line_line(line_number0(con_Dtixing(con_no%).data(0).poi(0), _
      con_Dtixing(con_no%).data(0).poi(2), 0, 0), line_number0(con_Dtixing(con_no%).data(0).poi(1), _
        con_Dtixing(con_no%).data(0).poi(3), 0, 0), 0, 0, 0, 0, 0, c_data0)
        If add_aid_point_from_conclusion > 1 Then
          Exit Function
        End If
  ElseIf conclusion_data(con_no%).ty = point4_on_circle_ Then
   If C_display_picture.m_circle.Count > 4 Then
     For k% = 1 To last_conditions.last_cond(1).two_angle_value0_no
      n% = Two_angle_value0.av_no(k%).no
       If minus_string(angle3_value(n%).data(0).data0.value, "180", True, False) = "0" Then
        For j% = 1 + con_no% To last_conditions.last_cond(1).two_angle_value0_no
         tn_% = Two_angle_value0.av_no(j%).no
          If minus_string(angle3_value(tn_%).data(0).data0.value, "180", True, False) = "0" Then
           add_aid_point_from_conclusion = combine_three_angle_with_three_angle_(n%, tn_%, 0)
            If add_aid_point_from_conclusion > 1 Then
             Exit Function
            End If
           End If
        Next j%
       End If
      Next k%
For k% = 1 + last_conditions.last_cond(0).angle3_value_no To last_conditions.last_cond(1).angle3_value_no
 n% = angle3_value(k%).data(0).record.data1.index.i(0)
   If minus_string(angle3_value(n%).data(0).data0.para(0), "1", True, False) = "0" Then
     If (minus_string(angle3_value(n%).data(0).data0.value, "0", True, False) = "0" And _
       (minus_string(angle3_value(n%).data(0).data0.para(1), "1", True, False) = "0" Or _
         minus_string(angle3_value(n%).data(0).data0.para(1), "-1", True, False) = "0") And _
         (minus_string(angle3_value(n%).data(0).data0.para(2), "1", True, False = "0") Or _
           minus_string(angle3_value(n%).data(0).data0.para(2), "-1", True, False) = "0")) Or _
           (minus_string(angle3_value(n%).data(0).data0.value, "180", True, False) = "0" And _
             minus_string(angle3_value(n%).data(0).data0.para(1), "1", True, False) = "0") Then
      For j% = 1 + con_no% To last_conditions.last_cond(1).angle3_value_no
            tn_% = angle3_value(j%).data(0).record.data1.index.i(0)
        If minus_string(angle3_value(tn_%).data(0).data0.para(0), "1", True, False) = "0" Then
         If (angle3_value(n%).data(0).data0.value = "0" And _
          (minus_string(angle3_value(tn_%).data(0).data0.para(1), "1", True, False) = "0" Or _
            minus_string(angle3_value(tn_%).data(0).data0.para(1), "-1", True, False) = "0") And _
            (minus_string(angle3_value(tn_%).data(0).data0.para(2), "1", True, False) = "0" Or _
              minus_string(angle3_value(tn_%).data(0).data0.para(2), "-1", True, False) = "0")) Or _
            (minus_string(angle3_value(tn_%).data(0).data0.value, "180", True, False) = "0" And _
               minus_string(angle3_value(tn_%).data(0).data0.para(1), "1", True, False) = "0") Then
           add_aid_point_from_conclusion = combine_three_angle_with_three_angle_(n%, tn_%, 0)
            If add_aid_point_from_conclusion > 1 Then
             Exit Function
            End If
         End If
        End If
      Next j%
     End If
  End If
Next k%
Else
    add_aid_point_from_conclusion = conclusion_four_point_on_circle(con_no%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
    End If
End If
 add_aid_point_from_conclusion = start_prove(0, 1, 1)   'call_theorem(0, no_reduce)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
ElseIf conclusion_data(con_no%).ty = sides_length_of_triangle_ Then
   add_aid_point_from_conclusion = conclusion_for_sides_length_of_triangle(con_no%)
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
'End If
  'If conclusion_no(con_no% - 1) = 0 Then
   'If conclusion_data(con_no% - 1) = point4_on_circle_ Then
ElseIf conclusion_data(con_no%).ty = general_string_ Then
   For j% = last_conditions.last_cond(1).general_string_no To 1 Step -1
     If general_string(j%).record_.conclusion_no > 0 Then
      For k% = 0 To 3
       If general_string(j%).data(0).item(0) > 0 Then
          If item0(general_string(j%).data(0).item(k%)).data(0).sig = "~" And k% = 0 Then
             'If item0(general_string(j%).data(0).item(1)).data(0).sig = "~" And _
              '  item0(general_string(j%).data(0).item(2)).data(0).sig = "~" And _
               '   general_string(j%).data(0).item(3) = 0 Then
                ' add_aid_point_from_conclusion = add_point_for_con_line3_value0( _
                 '  item0(general_string(j%).data(0).item(0)).data(0).poi(0), _
                  '  item0(general_string(j%).data(0).item(0)).data(0).poi(1), _
                   '  item0(general_string(j%).data(0).item(1)).data(0).poi(0), _
                    '  item0(general_string(j%).data(0).item(1)).data(0).poi(1), _
                     '  item0(general_string(j%).data(0).item(2)).data(0).poi(0), _
                      '  item0(general_string(j%).data(0).item(2)).data(0).poi(1), _
                       '   general_string(j%).data(0).para(0), _
                        '    general_string(j%).data(0).para(1), _
                         '    general_string(j%).data(0).para(2), "0")
                  'Exit Function
             'End If
          ElseIf item0(general_string(j%).data(0).item(k%)).data(0).sig = "*" Then
            If item0(general_string(j%).data(0).item(k%)).data(0).poi(1) > 0 Then
             If item0(general_string(j%).data(0).item(k%)).data(0).line_no(0) > 0 And _
                 item0(general_string(j%).data(0).item(k%)).data(0).line_no(0) = _
                  item0(general_string(j%).data(0).item(k%)).data(0).line_no(1) Then
               add_aid_point_from_conclusion = add_point_for_chord_tangent_line( _
                 item0(general_string(j%).data(0).item(k%)).data(0).poi(0), _
                  item0(general_string(j%).data(0).item(k%)).data(0).poi(1), _
                   item0(general_string(j%).data(0).item(k%)).data(0).poi(2), _
                    item0(general_string(j%).data(0).item(k%)).data(0).poi(3), False)
                    If add_aid_point_from_conclusion > 1 Then
                     Exit Function
                    End If
             End If
            ElseIf item0(general_string(j%).data(0).item(k%)).data(0).poi(1) < 0 And _
                   item0(general_string(j%).data(0).item(k%)).data(0).poi(1) > -5 Then
            If item0(general_string(j%).data(0).item(k%)).data(0).poi(0) > 0 And _
                 item0(general_string(j%).data(0).item(k%)).data(0).poi(2) > 0 Then
                  add_aid_point_from_conclusion = _
                     add_point_from_two_angle_for_Rtriangle( _
                      item0(general_string(j%).data(0).item(k%)).data(0).poi(0), _
                       item0(general_string(j%).data(0).item(k%)).data(0).poi(2))
                  If add_aid_point_from_conclusion > 1 Then
                     Exit Function
                  End If
            Else
                  add_aid_point_from_conclusion = _
                     add_point_from_angle_for_Rtriangle( _
                         item0(general_string(j%).data(0).item(k%)).data(0).poi(0))
                  If add_aid_point_from_conclusion > 1 Then
                     Exit Function
                  End If
            End If
            End If
           End If
       End If
      Next k%
       add_aid_point_from_conclusion = add_aid_point0(j%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
     End If
'      add_aid_point_from_conclusion = add_point_for_general_string_satis_midpoint(j%)
'     If add_aid_point_from_conclusion > 1 Then
'      Exit Function
'     End If
    End If
   Next j%
  ElseIf conclusion_data(con_no%).ty = line3_value_ Then
    add_aid_point_from_conclusion = add_point_for_con_line3_value(con_no%)
    If add_aid_point_from_conclusion > 1 Then
     Exit Function
    End If
  ElseIf conclusion_data(con_no%).ty = angle3_value_ Then
    If con_angle3_value(con_no%).data(0).data0.angle(0) > 0 Then
       add_aid_point_from_conclusion = add_point_for_circle_center_angle( _
              con_angle3_value(con_no%).data(0).data0.angle(0))
        If add_aid_point_from_conclusion > 1 Then
        Exit Function
        End If
    End If
    If con_angle3_value(con_no%).data(0).data0.angle(1) > 0 Then
       add_aid_point_from_conclusion = add_point_for_circle_center_angle( _
              con_angle3_value(con_no%).data(0).data0.angle(1))
        If add_aid_point_from_conclusion > 1 Then
        Exit Function
        End If
    End If
    If con_angle3_value(con_no%).data(0).data0.angle(2) > 0 Then
       add_aid_point_from_conclusion = add_point_for_circle_center_angle( _
              con_angle3_value(con_no%).data(0).data0.angle(2))
        If add_aid_point_from_conclusion > 1 Then
         Exit Function
        End If
    End If
    If con_angle3_value(con_no%).data(0).data0.para(0) = "1" And _
        (con_angle3_value(con_no%).data(0).data0.para(1) = "1" And _
          con_angle3_value(con_no%).data(0).data0.para(2) = "1") Then
       If angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(1) Then
       If angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).line_no(0) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1) And _
          angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).line_no(1) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(0) Then
         add_aid_point_from_conclusion = add_aid_point_for_triangle( _
           angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
            angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).poi(1), _
              angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).poi(1))
         If add_aid_point_from_conclusion > 1 Then
          Exit Function
         End If
       End If
       ElseIf angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).line_no(1) Then
       If angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(0) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1) And _
          angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(1) = _
           angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).line_no(0) Then
         add_aid_point_from_conclusion = add_aid_point_for_triangle( _
           angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
            angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).poi(1), _
              angle(con_angle3_value(con_no%).data(0).data0.angle(2)).data(0).poi(1))
         If add_aid_point_from_conclusion > 1 Then
          Exit Function
         End If
       End If
       End If
    ElseIf con_angle3_value(con_no%).data(0).data0.para(0) = "1" And _
        (con_angle3_value(con_no%).data(0).data0.para(1) = "-1" Or _
          con_angle3_value(con_no%).data(0).data0.para(1) = "@1") And _
         con_angle3_value(con_no%).data(0).data0.para(2) = "0" And _
         con_angle3_value(con_no%).data(0).data0.value = "0" Then
         add_aid_point_from_conclusion = add_point_for_eangle( _
            con_angle3_value(con_no%).data(0).data0.angle(0), _
             con_angle3_value(con_no%).data(0).data0.angle(1))
         If add_aid_point_from_conclusion > 1 Then
          Exit Function
         End If
    '**********
'    If con_angle3_value(con_no%).data(0).angle_(3) > 0 Then
     If angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0) = _
          angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(1) Then
         t_l(0) = angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1)
         t_l(1) = angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(0)
     Call line_number0(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
      m_lin(t_l(0)).data(0).data0.poi(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).te(1)), t_n(0), t_n(1))
     Call line_number0(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
      m_lin(t_l(1)).data(0).data0.poi(angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).te(0)), t_n(2), t_n(3))
    ElseIf angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(1) = _
         angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(0) Then
        t_l(0) = angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).line_no(0)
        t_l(1) = angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).line_no(1)
     Call line_number0(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
      m_lin(t_l(0)).data(0).data0.poi(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).te(0)), t_n(0), t_n(1))
     Call line_number0(angle(con_angle3_value(con_no%).data(0).data0.angle(0)).data(0).poi(1), _
      m_lin(t_l(1)).data(0).data0.poi(angle(con_angle3_value(con_no%).data(0).data0.angle(1)).data(0).te(1)), t_n(2), t_n(3))
    End If
   add_aid_point_from_conclusion = add_point_for_eangle1(t_n(0), t_n(1), t_l(0), _
                 t_n(2), t_n(3), t_l(1))
    If add_aid_point_from_conclusion > 1 Then
    Exit Function
    End If
    add_aid_point_from_conclusion = add_point_for_eangle1(t_n(2), t_n(3), t_l(1), _
                 t_n(0), t_n(1), t_l(0))
    If add_aid_point_from_conclusion > 1 Then
    Exit Function
    End If
   'End If
    '**********
    ElseIf con_angle3_value(con_no%).data(0).data0.para(0) = "1" And _
        con_angle3_value(con_no%).data(0).data0.para(1) = "1" And _
         con_angle3_value(con_no%).data(0).data0.para(1) = "0" And _
          con_angle3_value(con_no%).data(0).data0.value = "180" Then
        add_aid_point_from_conclusion = conclusion_two_angle_pi(con_angle3_value(con_no%).data(0).data0.angle(0), _
          con_angle3_value(con_no% - 1).data(0).data0.angle(1))
           If add_aid_point_from_conclusion > 1 Then
              Exit Function
           End If
    ElseIf con_angle3_value(con_no%).data(0).data0.para(0) = "1" And _
        con_angle3_value(con_no%).data(0).data0.para(1) = "-2" And _
         con_angle3_value(con_no%).data(0).data0.para(2) = "0" And _
          con_angle3_value(con_no%).data(0).data0.value = "0" Then
        add_aid_point_from_conclusion = add_aid_point_for_double_angle( _
              con_angle3_value(con_no%).data(0).data0.angle(0), _
                con_angle3_value(con_no%).data(0).data0.angle(1))
                 If add_aid_point_from_conclusion > 1 Then
                    Exit Function
                 End If
    ElseIf con_angle3_value(con_no%).data(0).data0.para(0) = "2" And _
        con_angle3_value(con_no%).data(0).data0.para(1) = "-1" And _
         con_angle3_value(con_no%).data(0).data0.para(2) = "0" And _
          con_angle3_value(con_no%).data(0).data0.value = "0" Then
        add_aid_point_from_conclusion = add_aid_point_for_double_angle( _
              con_angle3_value(con_no%).data(0).data0.angle(1), _
                con_angle3_value(con_no%).data(0).data0.angle(2))
                 If add_aid_point_from_conclusion > 1 Then
                  Exit Function
                 End If
  End If
  ElseIf conclusion_data(con_no%).ty = tangent_line_ Then
     add_aid_point_from_conclusion = add_aid_point_for_tangent_line(con_no%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
     End If
  ElseIf conclusion_data(con_no%).ty = relation_ Then
     add_aid_point_from_conclusion = add_point_for_chord(con_relation(con_no%).data(0).poi(0), _
          con_relation(con_no%).data(0).poi(1), 0)
      If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
     add_aid_point_from_conclusion = add_point_for_chord(con_relation(con_no%).data(0).poi(2), _
          con_relation(con_no%).data(0).poi(3), 0)
      If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
      If con_relation(con_no%).data(0).value = "2" Then
         add_aid_point_from_conclusion = add_mid_point(con_relation(con_no%).data(0).poi(0), _
                0, con_relation(con_no%).data(0).poi(1), 0)
      If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
      ElseIf con_relation(con_no%).data(0).value = "1/2" Then
         add_aid_point_from_conclusion = add_mid_point(con_relation(con_no%).data(0).poi(2), _
                0, con_relation(con_no%).data(0).poi(3), 0)
      If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
      End If
  '    n% = -1
  '    If con_relation(con_no%).data(0).poi(4) > 0 And con_relation(con_no%).data(0).poi(5) > 0 Then
  '    For i% = 1 To last_conditions.last_cond(1).relation_no
  '        If Drelation(i%).data(0).record.data0.condition_data.condition_no = 0 Then
  '           If Drelation(i%).data(0).data0.poi(4) > 0 And Drelation(i%).data(0).data0.poi(5) > 0 Then
  '             n% = n% + 1
  '              t_n(n%) = i%
  '           End If
  '        End If
  '    Next i%
  '    If n% >= 0 Then
  '       t_l(0) = Drelation(t_n(0)).data(0).data0.line_no(2)
  '       p% = 0
  '     If n% > 0 Then
  '       t_l(1) = Drelation(t_n(1)).data(0).data0.line_no(2)
  '        If Drelation(t_n(0)).data(0).data0.poi(0) = Drelation(t_n(1)).data(0).data0.poi(0) Then
  '         p% = is_line_line_intersect(line_number0(Drelation(t_n(0)).data(0).data0.poi(3), _
               Drelation(t_n(1)).data(0).data0.poi(1), 0, 0), line_number0(Drelation(t_n(1)).data(0).data0.poi(3), _
               Drelation(t_n(0)).data(0).data0.poi(1), 0, 0), 0, 0)
  '        ElseIf Drelation(t_n(0)).data(0).data0.poi(3) = Drelation(t_n(1)).data(0).data0.poi(0) Then
  '         p% = is_line_line_intersect(line_number0(Drelation(t_n(0)).data(0).data0.poi(0), _
               Drelation(t_n(1)).data(0).data0.poi(1), 0, 0), line_number0(Drelation(t_n(1)).data(0).data0.poi(3), _
               Drelation(t_n(0)).data(0).data0.poi(1), 0, 0), 0, 0)
  '        ElseIf Drelation(t_n(0)).data(0).data0.poi(0) = Drelation(t_n(1)).data(0).data0.poi(3) Then
  '         p% = is_line_line_intersect(line_number0(Drelation(t_n(0)).data(0).data0.poi(3), _
               Drelation(t_n(1)).data(0).data0.poi(1), 0, 0), line_number0(Drelation(t_n(1)).data(0).data0.poi(0), _
               Drelation(t_n(0)).data(0).data0.poi(1), 0, 0), 0, 0)
  '        ElseIf Drelation(t_n(0)).data(0).data0.poi(3) = Drelation(t_n(1)).data(0).data0.poi(3) Then
  '         p% = is_line_line_intersect(line_number0(Drelation(t_n(0)).data(0).data0.poi(0), _
               Drelation(t_n(1)).data(0).data0.poi(1), 0, 0), line_number0(Drelation(t_n(1)).data(0).data0.poi(0), _
               Drelation(t_n(0)).data(0).data0.poi(1), 0, 0), 0, 0)
  '        End If
  '     Else
  '       t_l(1) = 0
  '     End If
  '     tl% = line_number0(con_relation(con_no%).data(0).poi(0), Drelation(t_n(0)).data(0).data0.poi(0), 0, 0)
  '       If tl% = 0 Or tl% = con_relation(con_no%).data(0).line_no(0) Or tl% = t_l(0) Or tl% = t_l(1) Then
  '        tl% = line_number0(con_relation(con_no%).data(0).poi(0), Drelation(t_n(0)).data(0).data0.poi(1), 0, 0)
  '         If tl% = 0 Or tl% = con_relation(con_no%).data(0).line_no(0) Or tl% = t_l(0) Or tl% = t_l(1) Then
  '          tl% = line_number0(con_relation(con_no%).data(0).poi(0), Drelation(t_n(0)).data(0).data0.poi(2), 0, 0)
  '         End If
  '       End If
  '    End If
  '      If tl% > 0 And tl% <> con_relation(con_no%).data(0).line_no(0) Then
  '         add_aid_point_from_conclusion = add_paral_line(p%, _
             tl%, Drelation(t_n(0)).data(0).data0.line_no(0), 0, 0, 0, 0, 0, 2)
  '          If add_aid_point_from_conclusion > 1 Then
  '             Exit Function
  '          End If
  '      End If
  '      For i% = 1 To n%
  '         add_aid_point_from_conclusion = add_paral_line(Drelation(t_n(i%)).data(0).data0.poi(1), _
            tl%, Drelation(t_n(0)).data(0).data0.line_no(0), 0, 0, 0, 0, 0, 1)
  '          If add_aid_point_from_conclusion > 1 Then
  '             Exit Function
  '          End If
  '      Next i%
  ' End If
 ElseIf conclusion_data(con_no%).ty = dpoint_pair_ Then
      add_aid_point_from_conclusion = add_point_for_point_pair(con_dpoint_pair(con_no%).data(0), False)
     If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
  ElseIf conclusion_data(con_no%).ty = verti_ Then
   'If is_line_line_intersect(Lin(con_verti(con_no%).data(0).line_no(0)), _
                     Lin(con_verti(con_no%).data(0).line_no(1)), 0, 0) = 0 Then
   add_aid_point_from_conclusion = add_interset_point_line_line(con_verti(con_no%).data(0).line_no(0), _
         con_verti(con_no%).data(0).line_no(1), 0, 1, 0, 0, 0, c_data0)
          If add_aid_point_from_conclusion > 1 Then
            Exit Function
          End If
 ElseIf conclusion_data(con_no%).ty = midpoint_ Then
  tp(0) = con_mid_point(con_no%).data(0).poi(0)
   tp(1) = con_mid_point(con_no%).data(0).poi(1)
     tp(2) = con_mid_point(con_no%).data(0).poi(2)
    add_aid_point_from_conclusion = add_aid_point2(tp(0), tp(1), tp(2), tl%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
     End If
  ElseIf conclusion_data(con_no%).ty = paral_ Then
   add_aid_point_from_conclusion = add_point_for_paral(con_paral(con_no%).data(0).line_no(0), _
     con_paral(con_no%).data(0).line_no(1))
   If add_aid_point_from_conclusion > 1 Then
    Exit Function
   End If
 ElseIf conclusion_data(con_no%).ty = area_of_element_ Then
  If con_Area_of_element(con_no%).data(0).element.ty = polygon_ Then
   If Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).ty > 0 Then
    If Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).start_poi = 0 Then
     add_aid_point_from_conclusion = add_aid_point_from_conclusion_area_of_polygon( _
            Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(0), _
             Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(1), _
              Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(2), _
               Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(3), con_no%)
     If add_aid_point_from_conclusion > 1 Then
      Exit Function
     End If
    Else 'Dpolygon4(con_Area_of_polygon(con_no%).data(0).polygon4_no).data(0).start_poi=1
     add_aid_point_from_conclusion = add_aid_point_from_conclusion_area_of_polygon( _
            Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(1), _
             Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(2), _
              Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(3), _
               Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).poi(0), con_no%)
      If add_aid_point_from_conclusion > 1 Then
       Exit Function
      End If
     End If
    End If
  Else 'if con_Area_of_element(con_no%).data(0).element.ty = triangle_
  End If
 End If
 Else 'run_type=1
  If conclusion_data(con_no%).ty = relation_ Then
     conclusion_data(con_no%).ty = equation_
     If last_conditions.last_cond(1).unkown_element_no = 0 Then
      last_conditions.last_cond(1).unkown_element_no = 1
       unkown_element(1).char = "x"
        unkown_element(1).conclusion_no = con_no%
     ElseIf unkown_number = 1 Then
      unkown_number = 2
       unkown_char = "y"
     End If
     temp_record.record_data.data0.condition_data.condition_no = 0
     temp_record.record_data.data0.condition_data.condition(8).ty = new_point_
     tn_% = 0
     add_aid_point_from_conclusion = set_Drelation(con_relation(con_no%).data(0).poi(0), _
            con_relation(con_no%).data(0).poi(1), con_relation(con_no%).data(0).poi(2), _
             con_relation(con_no%).data(0).poi(3), con_relation(con_no%).data(0).n(0), _
              con_relation(con_no%).data(0).n(1), con_relation(con_no%).data(0).n(2), _
               con_relation(con_no%).data(0).n(3), con_relation(con_no%).data(0).line_no(0), _
                con_relation(con_no%).data(0).line_no(1), "x", temp_record, tn_%, 0, 0, 0, 0, False)
     temp_record.record_data.data0.condition_data.condition_no = 1
     temp_record.record_data.data0.condition_data.condition(1).ty = relation_
     temp_record.record_data.data0.condition_data.condition(1).no = tn_%
     c_data0.condition_no = 0
    t_l(0) = vector_number(con_relation(con_no%).data(0).poi(0), con_relation(con_no%).data(0).poi(1), "")
     t_l(1) = vector_number(con_relation(con_no%).data(0).poi(2), con_relation(con_no%).data(0).poi(3), "")
     If Dtwo_point_line(t_l(0)).data(0).line_no = Dtwo_point_line(t_l(1)).data(0).line_no Or _
           is_dparal(Dtwo_point_line(t_l(0)).data(0).line_no, Dtwo_point_line(t_l(1)).data(0).line_no, _
                      0, -1000, 0, 0, 0, 0) Then
       If Dtwo_point_line(t_l(0)).data(0).v_value <> "" Then
         add_aid_point_from_conclusion = set_V_line_value(Dtwo_point_line(t_l(1)).data(0).v_poi(0), _
            Dtwo_point_line(t_l(1)).data(0).v_poi(1), 0, 0, 0, divide_string( _
                      Dtwo_point_line(t_l(0)).data(0).v_value, "x", _
                       True, False), temp_record, 0, False)
           If add_aid_point_from_conclusion > 1 Then
              Exit Function
           End If
       ElseIf Dtwo_point_line(t_l(1)).data(0).v_value <> "" Then
         add_aid_point_from_conclusion = set_V_line_value(Dtwo_point_line(t_l(0)).data(0).v_poi(0), _
            Dtwo_point_line(t_l(0)).data(0).v_poi(1), 0, 0, 0, time_string("x", _
                      Dtwo_point_line(t_l(1)).data(0).v_value, _
                       True, False), temp_record, 0, False)
           If add_aid_point_from_conclusion > 1 Then
              Exit Function
           End If
       Else
        If con_relation(con_no%).data(0).poi(1) = con_relation(con_no%).data(0).poi(2) Then
          t_l(2) = vector_number(con_relation(con_no%).data(0).poi(0), con_relation(con_no%).data(0).poi(3), "")
          If Dtwo_point_line(t_l(2)).data(0).v_value <> "" Then
         add_aid_point_from_conclusion = set_V_line_value(Dtwo_point_line(t_l(1)).data(0).v_poi(0), _
            Dtwo_point_line(t_l(1)).data(0).v_poi(1), 0, 0, 0, divide_string( _
                      Dtwo_point_line(t_l(2)).data(0).v_value, "1+x", _
                       True, False), temp_record, 0, False)
           If add_aid_point_from_conclusion > 1 Then
              Exit Function
           End If
         Else
           GoTo add_aid_point_from_conclusion_mark0
          End If
        Else
add_aid_point_from_conclusion_mark0:
         Call set_item0(t_l(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, c_data0, 0, t_n(0), 0, _
                  0, c_data0, False) '0310
         Call set_item0(t_l(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, c_data0, 0, t_n(1), 0, _
                  0, c_data0, False) '0310
         add_aid_point_from_conclusion = set_general_string(t_n(0), t_n(1), 0, 0, para(0), time_string("-1", _
                time_string("x", para(0), False, False), True, False), _
                 "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
         If add_aid_point_from_conclusion > 1 Then
          Exit Function
         End If
       End If
       End If
     Else
     Call set_item0(t_l(0), -10, t_l(0), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, c_data0, 0, t_n(0), 0, _
                  0, c_data0, False) '0310
     Call set_item0(t_l(1), -10, t_l(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, c_data0, 0, t_n(1), 0, _
                  0, c_data0, False) '0310
     add_aid_point_from_conclusion = set_general_string(t_n(0), t_n(1), 0, 0, para(0), time_string("-1", _
                time_string("xx", para(0), False, False), True, False), _
                 "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
     If add_aid_point_from_conclusion > 1 Then
     Exit Function
     End If
     End If
  End If
 End If
 add_aid_point_from_conclusion = start_prove(0, 1, 1)
add_aid_point_from_conclusion_error:
End Function

Public Function add_aid_point_from_conclusion_area_of_polygon(ByVal p1%, ByVal p2%, _
        ByVal p3%, ByVal p4%, con_no%) As Byte
Dim tp(3) As Integer
Dim t_l(1) As Integer
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
If con_Area_of_element(con_no%).data(0).element.ty <> polygon_ Then
    Exit Function
End If
    t_l(0) = line_number0(tp(1), tp(3), 0, 0)
    t_l(1) = line_number0(tp(0), tp(2), 0, 0)
    add_aid_point_from_conclusion_area_of_polygon = _
         add_paral_line(tp(0), t_l(0), _
          Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).line_no(2), _
            tp(0), tp(2), con_no%, 0, 0, 0)
    If add_aid_point_from_conclusion_area_of_polygon > 1 Then
     Exit Function
    End If
    add_aid_point_from_conclusion_area_of_polygon = _
         add_paral_line(tp(1), _
           t_l(1), Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).line_no(2), _
             tp(1), tp(3), con_no%, 0, 0, 0)
    If add_aid_point_from_conclusion_area_of_polygon > 1 Then
     Exit Function
    End If
        add_aid_point_from_conclusion_area_of_polygon = _
         add_paral_line(tp(2), _
           t_l(0), Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).line_no(0), _
             tp(2), tp(0), con_no%, 0, 0, 0)
    If add_aid_point_from_conclusion_area_of_polygon > 1 Then
     Exit Function
    End If
    add_aid_point_from_conclusion_area_of_polygon = _
      add_paral_line(tp(3), _
           t_l(1), Dpolygon4(con_Area_of_element(con_no%).data(0).element.no).data(0).line_no(0), _
             tp(3), tp(1), con_no%, 0, 0, 0)
    If add_aid_point_from_conclusion_area_of_polygon > 1 Then
     Exit Function
    End If

End Function

Public Function add_aid_point_for_triangle(ByVal p1%, ByVal p2%, ByVal p3%)
Dim tp(1) As Integer
Dim tl%, n%
Dim temp_record As total_record_type
Dim el As add_point_for_eline_type 'paral_type
'On Error GoTo add_aid_point_for_triangle_mark1
 If last_conditions.last_cond(1).point_no = 26 Then
 add_aid_point_for_triangle = 6
  Exit Function
End If
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 tp(0) = last_conditions.last_cond(1).point_no
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 tp(1) = last_conditions.last_cond(1).point_no
Call get_new_char(tp(0))
Call get_new_char(tp(1))
t_coord.X = m_poi(p1%).data(0).data0.coordinate.X + _
     m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X
t_coord.Y = m_poi(p1%).data(0).data0.coordinate.Y + _
     m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y
     Call set_point_coordinate(tp(0), t_coord, False)
t_coord.X = m_poi(p1%).data(0).data0.coordinate.X + _
     m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X
t_coord.Y = m_poi(p1%).data(0).data0.coordinate.Y + _
     m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y
     Call set_point_coordinate(tp(1), t_coord, False)
   tl% = line_number(p2%, p3%, pointapi0, pointapi0, _
                     depend_condition(point_, p2%), depend_condition(point_, p3%), _
                     condition, condition_color, 1, 0)
   record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(p1, tl%, 0, no_display, False, 0, temp_record)
       Call arrange_data_for_new_point(tl%, 0)
   If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
   End If
    last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
     temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
     temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
     temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(0).new_point_no).data(0) = new_point_data_0
       new_point(last_conditions.last_cond(0).new_point_no).data(0).poi(0) = _
              last_conditions.last_cond(1).point_no - 1
       new_point(last_conditions.last_cond(0).new_point_no).data(0).poi(1) = _
              last_conditions.last_cond(1).point_no
       'new_point(last_conditions.last_cond(0).new_point_no).data(0).record = temp_record.record_data
       new_point(last_conditions.last_cond(0).new_point_no).data(0).add_to_line(0) = tl%
        'poi(last_conditions.last_cond(1).point_no).old_data = poi(last_conditions.last_cond(1).point_no).data
       n% = 0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1325, _
           "\\1\\" + m_poi(p1%).data(0).data0.name + _
           "\\2\\\" + m_poi(p2%).data(0).data0.name + _
                     m_poi(p3%).data(0).data0.name + _
           "\\3\\" + m_poi(tp(0)).data(0).data0.name + _
                    m_poi(tp(1)).data(0).data0.name + _
           "\\4\\" + set_display_line(tl%)) 'LoadResString_(782) + _
      n% = 0
   Call set_dparal(tl%, line_number0(p2%, p3%, 0, 0), temp_record, n%, 0, False)
     temp_record.record_data.data0.condition_data.condition_no = 1
       temp_record.record_data.data0.condition_data.condition(1).ty = paral_
        temp_record.record_data.data0.condition_data.condition(1).no = n%
         new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.ty = paral_
          new_point(last_conditions.last_cond(1).new_point_no).data(0).cond.no = n%
          temp_record.record_data.data0.theorem_no = 0
   add_aid_point_for_triangle = set_New_point(tp(0), temp_record, tl%, 0, _
      0, 0, 0, 0, 0, 1)
   If add_aid_point_for_triangle > 1 Then
    Exit Function
   End If
    add_aid_point_for_triangle = set_New_point(tp(1), temp_record, tl%, 0, _
      0, 0, 0, 0, 0, 1)
   If add_aid_point_for_triangle > 1 Then
    Exit Function
   End If

 'If ty = 0 Then
 add_aid_point_for_triangle = start_prove(0, 1, 1)  'call_theorem(0, no_reduce)
   If add_aid_point_for_triangle > 1 Then
    Exit Function
   End If
add_aid_point_for_triangle_mark1:
 'If new_result_from_add = False Then
  Call from_aid_to_old
'End If ' new_result_from_add = False
' End If
End Function
Public Function add_aid_point_for_double_angle(ByVal A1%, ByVal A2%) As Byte
'结论中A1%=2A2%,在A1%中取一半,使之与A2%所在的三角形全等或相似
Dim i%, j%, l%, k%
Dim triA(1) As temp_triangle_type
Dim r1!, r2!, k_!
Call set_temp_triangle_from_angle(A1%, 0, triA(0), True)
Call set_temp_triangle_from_angle(A2%, 0, triA(1), False)
For i% = 1 To triA(0).last_T
 For j% = 1 To triA(1).last_T
  If is_equal_angle(triA(0).data(i%).angle(1), triA(1).data(j%).angle(1), 0, 0) Then
   add_aid_point_for_double_angle = add_aid_point_for_double_angle_( _
    triA(0).data(i%).poi(0), triA(0).data(i%).poi(1), triA(0).data(i%).poi(2), _
      triA(1).data(j%).poi(0), triA(1).data(j%).poi(1), triA(1).data(j%).poi(2))
       If add_aid_point_for_double_angle > 1 Then
              Exit Function
       End If
  ElseIf is_equal_angle(triA(0).data(i%).angle(1), triA(1).data(j%).angle(2), 0, 0) Then
   add_aid_point_for_double_angle = add_aid_point_for_double_angle_( _
    triA(0).data(i%).poi(0), triA(0).data(i%).poi(1), triA(0).data(i%).poi(2), _
      triA(1).data(j%).poi(0), triA(1).data(j%).poi(2), triA(1).data(j%).poi(1))
       If add_aid_point_for_double_angle > 1 Then
              Exit Function
       End If
  ElseIf is_equal_angle(triA(0).data(i%).angle(2), triA(1).data(j%).angle(2), 0, 0) Then
   add_aid_point_for_double_angle = add_aid_point_for_double_angle_( _
    triA(0).data(i%).poi(0), triA(0).data(i%).poi(2), triA(0).data(i%).poi(3), _
      triA(1).data(j%).poi(0), triA(1).data(j%).poi(2), triA(1).data(j%).poi(1))
       If add_aid_point_for_double_angle Then
          Exit Function
       End If
  ElseIf is_equal_angle(triA(0).data(i%).angle(2), triA(1).data(j%).angle(1), 0, 0) Then
   add_aid_point_for_double_angle = add_aid_point_for_double_angle_( _
    triA(0).data(i%).poi(0), triA(0).data(i%).poi(2), triA(0).data(i%).poi(3), _
      triA(1).data(j%).poi(0), triA(1).data(j%).poi(1), triA(1).data(j%).poi(2))
       If add_aid_point_for_double_angle Then
          Exit Function
       End If
  End If
 Next j%
Next i%
add_aid_point_for_double_angle = add_aid_point_for_double_angle0(A1%, A2%)
If add_aid_point_for_double_angle > 1 Then
   Exit Function
End If
For j% = 1 To triA(1).last_T
 For i% = 1 To 2
  If angle(triA(1).data(j%).angle(i%)).data(0).value <> "" Then
   For k% = 1 To last_conditions.last_cond(1).angle_no
        If angle(k%).data(0).other_no = k% And k% <> triA(1).data(0).angle(i%) Then
         If angle(k%).data(0).value = angle(triA(1).data(j%).angle(i%)).data(0).value Then
             If angle(A1%).data(0).line_no(0) = angle(k%).data(0).line_no(1) And _
                angle(A1%).data(0).poi(1) <> angle(k%).data(0).poi(1) Then
                 If angle_number(angle(A1%).data(0).poi(1), angle(k%).data(0).poi(1), _
                         angle(A1%).data(0).poi(2), "", 0) * angle_number( _
                          angle(A1%).data(0).poi(1), angle(A1%).data(0).poi(0), _
                           angle(A1%).data(0).poi(2), "", 0) > 0 Then
                r1! = (m_poi(triA(1).data(j%).poi(1)).data(0).data0.coordinate.X - _
                      m_poi(triA(1).data(j%).poi(2)).data(0).data0.coordinate.X) ^ 2 + _
                    (m_poi(triA(1).data(j%).poi(1)).data(0).data0.coordinate.Y - _
                      m_poi(triA(1).data(j%).poi(2)).data(0).data0.coordinate.Y) ^ 2
                r2! = (m_poi(triA(1).data(j%).poi(0)).data(0).data0.coordinate.X - _
                      m_poi(triA(1).data(j%).poi(i%)).data(0).data0.coordinate.X) ^ 2 + _
                    (m_poi(triA(1).data(j%).poi(0)).data(0).data0.coordinate.Y - _
                      m_poi(triA(1).data(j%).poi(i%)).data(0).data0.coordinate.Y) ^ 2
                r1! = sqr(r1!)
                r2! = sqr(r2!)
                k_! = r1! / r2!
                add_aid_point_for_double_angle = add_aid_point_for_eangle0( _
                    A2%, angle(A1%).data(0).poi(1), angle(k%).data(0).poi(1), angle(k%).data(0).poi(0), k_!)
                     If add_aid_point_for_double_angle > 1 Then
                        Exit Function
                     End If
                End If
             ElseIf angle(A1%).data(0).line_no(1) = angle(k%).data(0).line_no(0) And _
                angle(A1%).data(0).poi(1) <> angle(k%).data(0).poi(1) Then
                If angle_number(angle(A1%).data(0).poi(1), angle(k%).data(0).poi(1), _
                         angle(A1%).data(0).poi(0), "", 0) * angle_number( _
                          angle(A1%).data(0).poi(1), angle(A1%).data(0).poi(2), _
                           angle(A1%).data(0).poi(0), "", 0) > 0 Then
                End If
             End If
         End If
        End If
   Next k%
  Else
  End If
 Next i%
Next j%
End Function
Public Function add_item0(ByVal it1%, ByVal para1$, _
             ByVal it2%, ByVal para2$, o_it%, o_para$, no_reduce As Byte) As Byte
Dim tp(1) As Integer
Dim is_no_initial As Integer
Dim c_data As condition_data_type
Dim para$
record_0.data0.condition_data.condition_no = 0 ' record0
o_it% = 0
If it1% = it2% Then
 o_it% = it1%
  o_para$ = add_string(para1$, para2$, True, False)
   If o_para$ = "0" Then
    o_it% = 0
   End If
   add_item0 = 1
Else
 If item0(it1%).data(0).sig = "~" And item0(it2%).data(0).sig = "~" Then
  If add_line_with_para(item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
        para1$, item0(it2%).data(0).poi(0), item0(it2%).data(0).poi(1), para2$, tp(0), _
         tp(1), o_para$, is_no_initial, c_data) Then
   Call set_item0(tp(0), tp(1), 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para$, 0, _
           record_data0.data0.condition_data, 0, o_it%, no_reduce, is_no_initial, c_data, False) '0310
            o_para$ = time_string(o_para$, para$, True, False)
             add_item0 = 1
  End If
 ElseIf item0(it1%).data(0).sig = "*" And item0(it2%).data(0).sig = "*" Then
 If item0(it1%).data(0).poi(0) = item0(it2%).data(0).poi(0) And _
       item0(it1%).data(0).poi(1) = item0(it2%).data(0).poi(1) Then
   If add_line_with_para(item0(it1%).data(0).poi(2), item0(it1%).data(0).poi(3), _
         para1$, item0(it2%).data(0).poi(2), item0(it2%).data(0).poi(3), para2$, tp(0), _
           tp(1), o_para$, is_no_initial, c_data) Then
    Call set_item0(tp(0), tp(1), item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
         "*", 0, 0, item0(it1%).data(0).n(0), item0(it1%).data(0).n(1), 0, _
          item0(it1%).data(0).line_no(0), "1", "1", "1", "", para$, 0, record_data0.data0.condition_data, _
            0, o_it%, no_reduce, is_no_initial, c_data, False) '0310
             o_para$ = time_string(o_para$, para$, True, False)
              add_item0 = 1
   End If
 ElseIf item0(it1%).data(0).poi(2) = item0(it2%).data(0).poi(0) And _
       item0(it1%).data(0).poi(3) = item0(it2%).data(0).poi(1) Then
   If add_line_with_para(item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
          para1$, item0(it2%).data(0).poi(2), item0(it2%).data(0).poi(3), para2$, tp(0), _
            tp(1), o_para$, is_no_initial, c_data) Then
    Call set_item0(tp(0), tp(1), item0(it1%).data(0).poi(2), item0(it1%).data(0).poi(3), _
        "*", 0, 0, item0(it1%).data(0).n(2), item0(it1%).data(0).n(3), 0, item0(it1%).data(0).line_no(1), _
           "1", "1", "1", "", para$, 0, record_data0.data0.condition_data, 0, o_it%, no_reduce, is_no_initial, _
              c_data, False) '0310
           o_para$ = time_string(o_para$, para$, True, False)
              add_item0 = 1
   End If
 ElseIf item0(it1%).data(0).poi(0) = item0(it2%).data(0).poi(2) And _
       item0(it1%).data(0).poi(1) = item0(it2%).data(0).poi(3) Then
   If add_line_with_para(item0(it1%).data(0).poi(2), item0(it1%).data(0).poi(3), _
         para1$, item0(it2%).data(0).poi(0), item0(it2%).data(0).poi(1), para2$, tp(0), _
           tp(1), o_para$, is_no_initial, c_data) Then
    Call set_item0(tp(0), tp(1), item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
            "*", 0, 0, item0(it1%).data(0).n(0), item0(it1%).data(0).n(1), 0, _
              item0(it1%).data(0).line_no(0), "1", "1", "1", "", para$, 0, record_data0.data0.condition_data, _
               0, o_it%, no_reduce, is_no_initial, c_data, False) '0310
                o_para$ = time_string(o_para$, para$, True, False)
           add_item0 = 1
   End If
 ElseIf item0(it1%).data(0).poi(2) = item0(it2%).data(0).poi(2) And _
       item0(it1%).data(0).poi(3) = item0(it2%).data(0).poi(3) Then
    If add_line_with_para(item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
          para1$, item0(it2%).data(0).poi(0), item0(it2%).data(0).poi(1), para2$, tp(0), _
             tp(1), o_para$, is_no_initial, c_data) Then
    Call set_item0(tp(0), tp(1), item0(it1%).data(0).poi(2), item0(it1%).data(0).poi(3), _
           "*", 0, 0, item0(it1%).data(0).n(2), item0(it1%).data(0).n(3), 0, _
             item0(it1%).data(0).line_no(1), "1", "1", "1", "", para$, 0, record_data0.data0.condition_data, _
              0, o_it%, no_reduce, is_no_initial, c_data, False) '0310
               o_para$ = time_string(o_para$, para$, True, False)
               add_item0 = 1
   End If
 End If
 ElseIf item0(it1%).data(0).sig = "/" And item0(it2%).data(0).sig = "/" Then
 If item0(it1%).data(0).poi(2) = item0(it2%).data(0).poi(2) And _
       item0(it1%).data(0).poi(3) = item0(it2%).data(0).poi(3) Then
    If add_line_with_para(item0(it1%).data(0).poi(0), item0(it1%).data(0).poi(1), _
           para1$, item0(it2%).data(0).poi(0), item0(it2%).data(0).poi(1), para2$, tp(0), _
            tp(1), o_para$, is_no_initial, c_data) Then
    If tp(0) = item0(it1%).data(0).poi(2) And tp(1) = item0(it1%).data(0).poi(3) Then
         para$ = "1"
        o_it% = 0
    Else
        Call set_item0(tp(0), tp(1), item0(it1%).data(0).poi(2), item0(it1%).data(0).poi(3), _
              "/", 0, 0, item0(it1%).data(0).n(2), item0(it1%).data(0).n(3), 0, _
                item0(it1%).data(0).line_no(1), "1", "1", "1", "", para$, 0, record_data0.data0.condition_data, _
                  0, o_it%, no_reduce, is_no_initial, c_data, False) '0310
    End If
     o_para$ = time_string(o_para$, para$, True, False)
        add_item0 = 1
   End If
 End If
 End If
End If
End Function
Private Function add_line_with_para(p1%, p2%, pA1$, _
              p3%, p4%, pA2$, op1%, op2%, opa$, is_no_initial As Integer, _
               c_data As condition_data_type) As Byte
Dim ty As Byte
Dim tp(3) As Integer
Call arrange_four_point(p1%, p2%, p3%, p4%, 0, 0, 0, 0, _
     0, 0, tp(0), tp(1), tp(2), tp(3), 0, 0, 0, 0, 0, 0, _
       0, 0, 0, 0, 0, ty, c_data, is_no_initial)
If pA1$ = pA2$ Then
 If ty = 3 Or ty = 5 Then
  op1% = tp(0)
   op2% = tp(3)
    opa$ = pA1$
     add_line_with_para = 1
 End If
ElseIf time_string("-1", pA1$, True, False) = pA2$ Then
 If ty = 4 Then
  op1% = tp(0)
   op2% = tp(1)
    opa$ = pA1$
     add_line_with_para = 1
 ElseIf ty = 6 Then
  op1% = tp(0)
   op2% = tp(1)
    opa$ = pA2$
     add_line_with_para = 1
 ElseIf ty = 7 Then
  op1% = tp(2)
   op2% = tp(3)
    opa$ = pA2$
     add_line_with_para = 1
 ElseIf ty = 8 Then
   op1% = tp(2)
    op2% = tp(3)
     opa$ = pA1$
      add_line_with_para = 1
 End If
Else
 Exit Function
End If
End Function
Public Sub draw_aid_angle(ByVal A%)
 Call C_display_picture.set_dot_line(angle(A%).data(0).poi(0), angle(A%).data(0).poi(1), 0, 0)
 Call C_display_picture.set_dot_line(angle(A%).data(0).poi(2), angle(A%).data(0).poi(1), 0, 0)
End Sub
Public Sub draw_aid_item(it%)
Dim i%
For i% = 0 To 1
 If item0(it%).data(0).poi(2 * i%) > 0 And _
     item0(it%).data(0).poi(2 * i% + 1) > 0 Then
     Call C_display_picture.set_dot_line(item0(it%).data(0).poi(2 * i%), item0(it%).data(0).poi(2 * i% + 1), 0, 0)
 ElseIf item0(it%).data(0).poi(2 * i% + 1) = -6 Then
     Call draw_aid_angle(item0(it%).data(0).poi(2 * i%))
 End If
Next i%
End Sub
Public Function add_point_for_con_eline(ByVal n%) As Byte
Dim tp(1) As Integer
Dim l%, i%
Dim tn(1) As Integer
Dim md_data As mid_point_data0_type
Dim con_data As condition_data_type
If con_eline(n%).data(0).data0.poi(0) = con_eline(n%).data(0).data0.poi(2) Then
   tp(0) = con_eline(n%).data(0).data0.poi(1)
   tp(1) = con_eline(n%).data(0).data0.poi(3)
ElseIf con_eline(n%).data(0).data0.poi(0) = con_eline(n%).data(0).data0.poi(3) Then
   tp(0) = con_eline(n%).data(0).data0.poi(1)
   tp(1) = con_eline(n%).data(0).data0.poi(2)
ElseIf con_eline(n%).data(0).data0.poi(1) = con_eline(n%).data(0).data0.poi(2) Then
   tp(0) = con_eline(n%).data(0).data0.poi(0)
   tp(1) = con_eline(n%).data(0).data0.poi(3)
ElseIf con_eline(n%).data(0).data0.poi(1) = con_eline(n%).data(0).data0.poi(3) Then
   tp(0) = con_eline(n%).data(0).data0.poi(0)
   tp(1) = con_eline(n%).data(0).data0.poi(2)
Else
   Exit Function
End If
l% = line_number0(tp(0), tp(1), tn(0), tn(1))
If tn(0) > tn(1) Then
   Call exchange_two_integer(tp(0), tp(1))
   Call exchange_two_integer(tn(0), tn(1))
End If
If m_lin(l%).data(0).data0.in_point(0) > 2 Then
 For i% = 1 To m_lin(l%).data(0).data0.in_point(0)
     If i% <> tn(0) And i% <> tn(1) Then
        If is_mid_point(tp(0), m_lin(l%).data(0).data0.in_point(i%), tp(1), _
             tn(0), i%, tn(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, md_data, _
               "", 0, 0, 0, con_data) = False Then
           If i% < tn(0) Then
            add_point_for_con_eline = add_aid_point_for_eline( _
              m_lin(l%).data(0).data0.in_point(i%), tp(0), tp(0), tp(1), _
                 tp(1))
            If add_point_for_con_eline > 1 Then
               Exit Function
            End If
           ElseIf tn(0) < i% And i% < tn(1) Then
            add_point_for_con_eline = add_aid_point_for_eline( _
              tp(0), m_lin(l%).data(0).data0.in_point(i%), tp(1), tp(0), _
                 tp(1))
            If add_point_for_con_eline > 1 Then
               Exit Function
            End If
           Else
            add_point_for_con_eline = add_aid_point_for_eline( _
              m_lin(l%).data(0).data0.in_point(i%), tp(1), tp(1), tp(0), _
                 tp(0))
            If add_point_for_con_eline > 1 Then
               Exit Function
            End If
           End If
        End If
     End If
  Next i%
 End If
End Function

Public Function add_point_from_two_angle_for_Rtriangle(ByVal A1%, ByVal A2%) As Byte
Dim tp As Integer
Dim tl(2) As Integer
'作高
If angle(A1%).data(0).line_no(0) = angle(A2%).data(0).line_no(0) Then
    tl(0) = angle(A1%).data(0).line_no(0)
    tl(1) = angle(A1%).data(0).line_no(1)
    tl(2) = angle(A2%).data(0).line_no(1)
ElseIf angle(A1%).data(0).line_no(0) = angle(A2%).data(0).line_no(1) Then
    tl(0) = angle(A1%).data(0).line_no(0)
    tl(1) = angle(A1%).data(0).line_no(1)
    tl(2) = angle(A2%).data(0).line_no(0)
ElseIf angle(A1%).data(0).line_no(1) = angle(A2%).data(0).line_no(0) Then
    tl(0) = angle(A1%).data(0).line_no(1)
    tl(1) = angle(A1%).data(0).line_no(0)
    tl(2) = angle(A2%).data(0).line_no(1)
ElseIf angle(A1%).data(0).line_no(1) = angle(A2%).data(0).line_no(1) Then
    tl(0) = angle(A1%).data(0).line_no(1)
    tl(1) = angle(A1%).data(0).line_no(0)
    tl(2) = angle(A2%).data(0).line_no(0)
Else
Exit Function
End If
tp = is_line_line_intersect(tl(1), tl(2), 0, 0, False)
If tp > 0 Then
 add_point_from_two_angle_for_Rtriangle = _
   add_aid_point_for_paral_or_verti(tp, verti_, tl(0), tl(0), 0)
End If
End Function
Public Function add_point_from_angle_for_Rtriangle(ByVal A%) As Byte
Dim tn(1) As Integer
Dim tn_(1) As Integer
Dim i%
If angle(A%).data(0).te(0) = 0 Then
   tn(0) = 0
   Call is_point_in_line3(angle(A%).data(0).poi(1), m_lin(angle(A%).data(0).line_no(0)).data(0).data0, tn(1))
Else
   Call is_point_in_line3(angle(A%).data(0).poi(1), m_lin(angle(A%).data(0).line_no(0)).data(0).data0, tn(0))
   tn(1) = m_lin(angle(A%).data(0).line_no(0)).data(0).data0.in_point(0)
End If
If angle(A%).data(0).te(1) = 0 Then
   tn_(0) = 0
   Call is_point_in_line3(angle(A%).data(0).poi(1), m_lin(angle(A%).data(0).line_no(1)).data(0).data0, tn_(1))
Else
   Call is_point_in_line3(angle(A%).data(0).poi(1), m_lin(angle(A%).data(0).line_no(1)).data(0).data0, tn_(0))
   tn_(1) = m_lin(angle(A%).data(0).line_no(1)).data(0).data0.in_point(0)
End If
For i% = tn(0) + 1 To tn(1)
   add_point_from_angle_for_Rtriangle = _
        add_aid_point_for_paral_or_verti(m_lin(angle(A%).data(0).line_no(0)).data(0).data0.in_point(i%), _
               verti_, angle(A%).data(0).line_no(1), angle(A%).data(0).line_no(1), 0)
   If add_point_from_angle_for_Rtriangle > 0 Then
      Exit Function
   End If
Next i%
For i% = tn_(0) + 1 To tn_(1)
   add_point_from_angle_for_Rtriangle = _
        add_aid_point_for_paral_or_verti(m_lin(angle(A%).data(0).line_no(1)).data(0).data0.in_point(i%), _
               verti_, angle(A%).data(0).line_no(0), angle(A%).data(0).line_no(0), 0)
   If add_point_from_angle_for_Rtriangle > 0 Then
      Exit Function
   End If
Next i%
End Function

Public Function add_two_point_for_mid_point(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%) As Byte
add_two_point_for_mid_point = add_mid_point(p1%, 0, p2%, 1)
  If add_two_point_for_mid_point > 1 Then
     Exit Function
  End If
add_two_point_for_mid_point = add_mid_point(p3%, 0, p4%, 0)
  If add_two_point_for_mid_point > 1 Then
     Exit Function
  End If
End Function
Public Function add_point_from_paral_and_circle(ByVal pl%, ByVal p4_on_circle%) As Byte
Dim i%, j%, k%, l%
Dim p4_c As four_point_on_circle_data_type
If pl% > 0 Then
 If Dparal(pl%).data(0).data0.record.data0.condition_data.condition_no = 0 Then
   For i% = 2 To m_lin(Dparal(pl%).data(0).data0.line_no(0)).data(0).data0.in_point(0)
   For j% = 1 To i% - 1
   For k% = 2 To m_lin(Dparal(pl%).data(0).data0.line_no(1)).data(0).data0.in_point(0)
   For l% = 1 To k% - 1
    If is_four_point_on_circle(m_lin(Dparal(pl%).data(0).data0.line_no(0)).data(0).data0.in_point(i%), _
         m_lin(Dparal(pl%).data(0).data0.line_no(0)).data(0).data0.in_point(j%), _
           m_lin(Dparal(pl%).data(0).data0.line_no(1)).data(0).data0.in_point(k), _
            m_lin(Dparal(pl%).data(0).data0.line_no(1)).data(0).data0.in_point(l%), _
              p4_on_circle%, p4_c, True) Then
     If m_Circ(p4_on_circle%).data(0).data0.center = 0 Or _
         m_poi(m_Circ(p4_on_circle%).data(0).data0.center).data(0).data0.visible = 0 Then
          Exit Function
     End If
     add_point_from_paral_and_circle = add_mid_point( _
          m_lin(Dparal(pl%).data(0).data0.line_no(0)).data(0).data0.in_point(i%), 0, _
            m_lin(Dparal(pl%).data(0).data0.line_no(0)).data(0).data0.in_point(j%), 2)
             If add_point_from_paral_and_circle > 1 Then
                Exit Function
             End If
     add_point_from_paral_and_circle = add_mid_point( _
          m_lin(Dparal(pl%).data(0).data0.line_no(1)).data(0).data0.in_point(k%), 0, _
            m_lin(Dparal(pl%).data(0).data0.line_no(1)).data(0).data0.in_point(l%), 2)
             Exit Function
    End If
   Next l%
   Next k%
   Next j%
   Next i%
 End If
Else
 i% = line_number0(four_point_on_circle(p4_on_circle%).data(0).poi(0), _
                        four_point_on_circle(p4_on_circle%).data(0).poi(1), 0, 0)
 j% = line_number0(four_point_on_circle(p4_on_circle%).data(0).poi(2), _
                        four_point_on_circle(p4_on_circle%).data(0).poi(3), 0, 0)
 If is_dparal(i%, j%, pl%, -1000, 0, 0, 0, 0) Then
   If Dparal(pl%).data(0).data0.record.data0.condition_data.condition_no = 0 Then
     add_point_from_paral_and_circle = add_mid_point( _
          four_point_on_circle(p4_on_circle%).data(0).poi(0), 0, _
            four_point_on_circle(p4_on_circle%).data(0).poi(1), 2)
             If add_point_from_paral_and_circle > 1 Then
                Exit Function
             End If
     add_point_from_paral_and_circle = add_mid_point( _
          four_point_on_circle(p4_on_circle%).data(0).poi(2), 0, _
            four_point_on_circle(p4_on_circle%).data(0).poi(3), 2)
             Exit Function
   End If
 End If
 i% = line_number0(four_point_on_circle(p4_on_circle%).data(0).poi(0), _
                        four_point_on_circle(p4_on_circle%).data(0).poi(2), 0, 0)
 j% = line_number0(four_point_on_circle(p4_on_circle%).data(0).poi(1), _
                        four_point_on_circle(p4_on_circle%).data(0).poi(3), 0, 0)
 If is_dparal(i%, j%, pl%, -1000, 0, 0, 0, 0) Then
   If Dparal(pl%).data(0).data0.record.data0.condition_data.condition_no = 0 Then
     add_point_from_paral_and_circle = add_mid_point( _
          four_point_on_circle(p4_on_circle%).data(0).poi(0), 0, _
            four_point_on_circle(p4_on_circle%).data(0).poi(2), 2)
             If add_point_from_paral_and_circle > 1 Then
                Exit Function
             End If
     add_point_from_paral_and_circle = add_mid_point( _
          four_point_on_circle(p4_on_circle%).data(0).poi(2), 0, _
            four_point_on_circle(p4_on_circle%).data(0).poi(3), 2)
             Exit Function
   End If
 End If
End If
End Function

Public Function add_point_from_tixing_for_condition(ByVal tx%) As Byte
Dim l1%, l2%, tA%, i%, no%, tp%
Dim can_add_point As Boolean
Dim cond_data(1) As condition_type
Dim temp_record As total_record_type
If run_type < 5 Then
'If Dtixing(tx%).data(0).record.data0.condition_data.condition_no = 0 Then
l1% = line_number0(Dtixing(tx%).data(0).poi(0), Dtixing(tx%).data(0).poi(2), 0, 0)
l2% = line_number0(Dtixing(tx%).data(0).poi(1), Dtixing(tx%).data(0).poi(3), 0, 0)
tA% = total_angle_no(l1%, l2%)
For i% = 0 To 3
  If T_angle(tA%).data(0).angle_no(i%).no > 0 Then
      If angle(T_angle(tA%).data(0).angle_no(i%).no).data(0).value <> "" Then
       can_add_point = True
      End If
  End If
Next i%
 If can_add_point = False Then
 If is_dverti(l1%, l2%, 0, -1000, 0, 0, 0, 0) Then
  can_add_point = True
 End If
 End If
  If can_add_point Then
        add_point_from_tixing_for_condition = add_paral_line(Dtixing(tx%).data(0).poi(1), _
             l1%, line_number0(Dtixing(tx%).data(0).poi(2), Dtixing(tx%).data(0).poi(3), 0, 0), _
               0, 0, 0, no%, tp%, 2)
       If add_point_from_tixing_for_condition > 1 Then
          Exit Function
       End If
       cond_data(0).ty = polygon_
       cond_data(0).no = Dtixing(tx%).data(0).poly4_no
       cond_data(1).ty = triangle_
       cond_data(1).no = triangle_number(Dtixing(tx%).data(0).poi(1), _
             Dtixing(tx%).data(0).poi(3), tp%, 0, 0, 0, 0, 0, 0, 0)
       temp_record.record_data.data0.condition_data.condition_no = 0
       Call add_conditions_to_record(paral_, no%, 0, 0, temp_record.record_data.data0.condition_data)
       Call add_conditions_to_record(tixing_, tx%, 0, 0, temp_record.record_data.data0.condition_data)
       add_point_from_tixing_for_condition = _
          set_area_relation(cond_data(0), cond_data(1), "1", temp_record, 0, 0, 0)
       If add_point_from_tixing_for_condition > 1 Then
          Exit Function
       End If
   End If
End If
'End If
End Function
Public Function add_element_value(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal sig) As Byte
If sig = "*" Then

ElseIf sig = "/" Then
 
Else
If p2% > 0 Then
ElseIf p2 = -1 Then
ElseIf p2 = -2 Then
ElseIf p2 = -3 Then
ElseIf p2 = -4 Then
ElseIf p2 = -5 Then
ElseIf p2 = -6 Then
 
ElseIf p2% = -7 Then
End If

End If
End Function
Public Function add_aid_point_for_circle_center(ByVal c%) As Byte
Dim temp_record As total_record_type
If m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.visible > 0 Then
  Exit Function
Else
   If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
      ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
   End If
    last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
      temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
      temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no
      temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
      new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
      new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = last_conditions.last_cond(1).point_no
    Call set_point_name(m_Circ(c%).data(0).data0.center, "")
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = _
    LoadResString_(1450, "\\1\\" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name + _
                         "\\2\\" + m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + _
                                   m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + _
                                   m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).data0.name)
End If
End Function

Public Sub add_depend_display_no(dis_no%, dep_no() As Integer)
Dim i%, j%
For i% = 1 To dep_no(0)
    If dep_no(i%) = dis_no Then
       Exit Sub
    End If
Next i%
For i% = 1 To dep_no(0)
    If dep_no(i%) > dis_no Then
        For j% = i% + 1 To 1
          dep_no(j%) = dep_no(j% - 1)
        Next j%
        dep_no(i%) = dis_no
        Exit Sub
    End If
Next i%
   dep_no(0) = dep_no(0) + 1
    dep_no(dep_no(0)) = dis_no
End Sub

Public Function add_condition_for_no_reduce() As Byte
Dim i%
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).line_value_no
   If line_value(i%).record_.no_reduce > 0 Then
     line_value(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_line_value(0, 0, "", 0, 0, 0, temp_record.record_data, i%, 0, False)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no
   If Drelation(i%).record_.no_reduce > 0 Then
     Drelation(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_Drelation(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", temp_record, i%, _
            0, 0, 0, 0, False)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no
   If Deline(i%).record_.no_reduce > 0 Then
      Deline(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_equal_dline(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, temp_record, i%, _
                              0, 0, 0, 0, False)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%

For i% = 1 To last_conditions.last_cond(1).two_line_value_no
   If two_line_value(i%).record_.no_reduce > 0 Then
     two_line_value(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_two_line_value(0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, "", "", "", temp_record, i%, 0)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%
For i% = 1 To last_conditions.last_cond(1).line3_value_no
   If line3_value(i%).record_.no_reduce > 0 Then
      line3_value(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_three_line_value(0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", "", temp_record, i%, 0, 0)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no
   If Ddpoint_pair(i%).record_.no_reduce > 0 Then
      Ddpoint_pair(i%).record_.no_reduce = 0
    add_condition_for_no_reduce = set_dpoint_pair(0, 0, 0, 0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, temp_record, False, i%, 0, 0, 0, False)
     If add_condition_for_no_reduce > 1 Then
        Exit Function
     End If
   End If
Next i%


End Function

Public Function read_tangent_line(initial_coord As POINTAPI, in_point_no%, out_coord_ As POINTAPI, tangent_type As Integer, _
        is_set_data As Boolean, Optional visible As Byte = 1, Optional line_type As Byte = condition, _
           Optional color As Byte = 3) As Integer
Dim i%, j%, k%
Dim out_ty As Integer
Dim in_coord As POINTAPI
Dim temp_tangent_line As tangent_line_type
Dim out_coord(1) As POINTAPI
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = wenti_cond_
temp_record.record_data.data0.condition_data.condition(1).no = _
 C_display_wenti.m_display_string.Count + 1
k% = last_conditions.last_cond(0).tangent_line_no
in_coord = initial_coord '输入鼠标的位置
  For i% = last_conditions.last_cond(1).tangent_line_no + 1 To last_conditions.last_cond(0).tangent_line_no '搜索切线
'        If i% <> k% Then
       If tangent_line(i%).data(0).visible >= 4 Then
        out_ty = is_point_on_line(in_coord, tangent_line(i%).data(0).coordinate(0), _
            tangent_line(i%).data(0).coordinate(1), out_coord_, out_coord(0), out_coord(1), aid_condition)  '在线外
       Else
        out_ty = is_point_on_line(in_coord, tangent_line(i%).data(0).coordinate(0), _
            tangent_line(i%).data(0).coordinate(1), out_coord_, out_coord(0), out_coord(1), condition)  '在线外
       End If
            Call set_new_coordinate_for_tangent_line(i%, out_coord(0), out_coord(1), out_ty) '输出线端点的坐标
          If out_ty = point_on_segement Or out_ty = point_out_segement Then
          If is_set_data Then
                        Call draw_tangent_line(i%, 0, is_set_data) '消线（选中的切线）
                 tangent_type = tangent_line(i%).tangent_type
                last_conditions.last_cond(1).tangent_line_no = last_conditions.last_cond(1).tangent_line_no + 1
            If tangent_line(i%).data(0).visible >= 4 Then
               tangent_line(i%).data(0).poi(0) = m_point_number(out_coord(0), condition, 1, condition_color, "", _
                       depend_condition(0, 0), depend_condition(0, 0), 0, True) '建立新点
             If tangent_line(i%).tangent_type = tangent_line_by_point_on_circle Then
                tangent_line(i%).data(0).line_no = line_number(tangent_line(i%).data(0).poi(0), _
                   0, tangent_line(i%).data(0).coordinate(0), _
                       tangent_line(i%).data(0).coordinate(1), depend_condition(point_, tangent_line(i%).data(0).poi(0)), _
                        depend_condition(0, 0), line_type, color, visible, 0) '建立新线
               in_point_no% = m_point_number(initial_coord, condition, 1, condition_color, "", _
                   depend_condition(0, 0), depend_condition(0, 0), 0, True)
               Call add_point_to_line(in_point_no%, tangent_line(i%).data(0).line_no, 0, False, False, 0, temp_record.record_data)
               tangent_line(i%).data(0).poi(1) = in_point_no%
             Else
              tangent_line(i%).data(0).line_no = line_number(tangent_line(i%).data(0).poi(0), _
                   tangent_line(i%).data(0).poi(1), tangent_line(i%).data(0).coordinate(0), _
                       tangent_line(i%).data(0).coordinate(1), depend_condition(point_, tangent_line(i%).data(0).poi(0)), _
                        depend_condition(point_, tangent_line(i%).data(0).poi(1)), line_type, color, visible, 0) '建立新线
             End If
            Else
              tangent_line(i%).data(0).poi(0) = m_point_number(tangent_line(i%).data(0).coordinate(0), condition, 1, condition_color, "", _
                   depend_condition(circle_, tangent_line(i%).data(0).circ(0)), depend_condition(circle_, tangent_line(i%).data(0).circ(1)), 0, True) '建立新点
              tangent_line(i%).data(0).poi(1) = m_point_number(tangent_line(i%).data(0).coordinate(1), condition, 1, condition_color, "", _
                   depend_condition(circle_, tangent_line(i%).data(0).circ(0)), depend_condition(circle_, tangent_line(i%).data(0).circ(1)), 0, True) '
              tangent_line(i%).data(0).line_no = line_number(tangent_line(i%).data(0).poi(0), _
                              tangent_line(i%).data(0).poi(1), tangent_line(i%).data(0).coordinate(0), _
                       tangent_line(i%).data(0).coordinate(1), depend_condition(point_, tangent_line(i%).data(0).poi(0)), _
                        depend_condition(point_, tangent_line(i%).data(0).poi(1)), line_type, color, visible, 0)
              in_point_no% = tangent_line(i%).data(0).poi(1) '
              initial_coord = m_poi(tangent_line(i%).data(0).poi(1)).data(0).data0.coordinate
             If tangent_line(i%).data(0).circ(0) > 0 And tangent_line(i%).data(0).circ(1) > 0 Then '两圆的公切线
               Call add_point_to_m_circle(tangent_line(i%).data(0).poi(0), _
                       tangent_line(i%).data(0).circ(0), record0, 0)
                   Call set_parent(circle_, tangent_line(i%).data(0).circ(0), point_, _
                       tangent_line(i%).data(0).poi(0), tangent_point_, 0)  '
                   Call set_parent(circle_, tangent_line(i%).data(0).circ(1), point_, _
                       tangent_line(i%).data(0).poi(0), tangent_point_, last_conditions.last_cond(1).tangent_line_no)  '
               Call add_point_to_m_circle(tangent_line(i%).data(0).poi(1), _
                       tangent_line(i%).data(0).circ(1), record0, 0)
                   'Call set_parent(circle_, tangent_line(i%).data(0).circ(0), point_, _
                       tangent_line(i%).data(0).poi(1), tangent_line(i%).tangent_type)
                       
             ElseIf tangent_line(i%).data(0).circ(0) > 0 Then '过圆上一点做切线
               Call add_point_to_m_circle(tangent_line(i%).data(0).poi(0), _
                       tangent_line(i%).data(0).circ(0), record0, 0)
             ElseIf tangent_line(i%).data(0).circ(1) > 0 Then '过圆外一点做切线
               Call add_point_to_m_circle(tangent_line(i%).data(0).poi(1), _
                       tangent_line(i%).data(0).circ(1), temp_record, 0)
               Call set_parent(point_, tangent_line(i%).data(0).poi(0), point_, _
                       tangent_line(i%).data(0).poi(1), tangent_line(i%).tangent_type)
             End If
             
            End If
            
            '*******************************************************************************
                 tangent_line(i%).data(0).visible = 0
            temp_tangent_line = tangent_line(i%) '切线的直线序号
               k% = last_conditions.last_cond(1).tangent_line_no
               'read_tangent_line = k% '读出线号
               'm_lin(k%).data(0).tangent_line_no =
              For j% = i% - 1 To last_conditions.last_cond(1).tangent_line_no Step -1 '移动未进入数据库的切线
                  tangent_line(j% + 1) = tangent_line(j%)
              Next j%
                  tangent_line(k%) = temp_tangent_line
                  read_tangent_line = tangent_line(k%).data(0).line_no
                  m_lin(read_tangent_line).data(0).tangent_line_no = k%
                  
    '**************************************************************************************************************************
                If tangent_line(k%).data(0).ele(1).ty = circle_ And tangent_line(k%).data(0).ele(0).ty = circle_ Then '两圆的公切线
                 If tangent_line(k%).data(0).poi(0) = tangent_line(k%).data(0).poi(1) Then '切点重合，两圆相切
                 Else
                         If out_ty = 2 Then
                         ' initial_coord = tangent_line(k%).data(0).coordinate(0) '进入read_inter_point为exist_point
                         '  in_point_no% = tangent_line(k%).data(0).poi(0)
                         End If
                 End If
               'Else
                'If out_ty = 2 Then
                '    in_point_no% = tangent_line(k%).data(0).poi(0)
                'End If
                ' Call set_wenti_cond_33_44(tangent_line(k%).data(0).poi(1),
                '               tangent_line(k%).data(0).poi(1), tangent_line(k%).data(0).ele(0).no, 0, _
                                 tangent_line(k%).data(0).poi(0), tangent_line(k%).data(0).poi(1), k%)
               End If
 '************************************************************************************************************************
         Else '
              Call draw_tangent_line(i%, 1, 0)
         End If
      Else '
                Call draw_tangent_line(i%, 1, 0)
      End If
'    End If
  Next i%
             For j% = k% + 1 To last_conditions.last_cond(0).tangent_line_no
              Call draw_tangent_line(j%, 0, 0)
               tangent_line(j%).data(0).visible = 3
           Next j%
End Function
Public Function read_tangent_circle(initial_coord As POINTAPI, out_coord As POINTAPI, out_point_no%, is_set_data As Boolean) As Integer
Dim i%, j%, k%, tl%
Dim out_ty As Byte
Dim temp_tangent_circle As tangent_circle_type
Dim read_tangent_circle_no%
Dim dis&
Dim is_new_point As Boolean
Dim temp_record As total_record_type
If is_set_data = False Then
   Exit Function
End If
  For i% = last_conditions.last_cond(1).tangent_circle_no + 1 To last_conditions.last_cond(0).tangent_circle_no
       dis& = distance_of_two_POINTAPI(initial_coord, m_tangent_circle(i%).data(0).data0(0).circle_center) '计算鼠标到圆心的距离
        If Abs(dis& - m_tangent_circle(i%).data(0).data0(0).circle_radii) < 4 Then 'Or _
               is_same_POINTAPI(initial_coord, m_tangent_circle(i%).data(0).data0(0).circle_center) Then '鼠标点落在圆周或圆心上
         read_tangent_circle_no% = i% '选中的切圆序号
          Call draw_tangent_circle(i%, True) '消初临时显示的圆
          m_tangent_circle(i%).data(0).circle_no = m_circle_number(1, m_tangent_circle(i%).data(0).center, _
            m_poi(m_tangent_circle(i%).data(0).center).data(0).data0.coordinate, 0, 0, 0, _
             m_tangent_circle(i%).data(0).data0(0).circle_radii, 0, 0, 0, 1, 1, condition_color, True) '建立切圆
          If m_tangent_circle(i%).data(0).ele(0).no > 0 And m_tangent_circle(i%).data(0).ele(0).ty > 0 Then
            m_tangent_circle(i%).data(0).tangent_poi(0) = m_point_number( _
                   m_tangent_circle(i%).data(0).data0(0).tangent_coord(0), condition, 1, condition_color, "", _
                    depend_condition(point_, m_Circ(m_tangent_circle(i%).data(0).circle_no).data(0).data0.center), _
                     m_tangent_circle(i%).data(0).ele(0), tangent_point_of_circle, True, is_new_point)
                      '建立切点,切点由被切圆和切圆的圆心确定
               Call add_point_to_m_circle(m_tangent_circle(i%).data(0).tangent_poi(0), _
                      m_tangent_circle(i%).data(0).circle_no, temp_record, True)
               Call set_parent(point_, m_tangent_circle(i%).data(0).tangent_poi(0), _
                          circle_, m_tangent_circle(i%).data(0).circle_no, 0) '
               If is_new_point = False Then '相切于已知点
                 If m_tangent_circle(i%).data(0).ele(0).ty = circle_ Then
                  tl% = line_number(m_tangent_circle(i%).data(0).tangent_poi(0), _
                      m_Circ(m_tangent_circle(i%).data(0).ele(0).no).data(0).data0.center, _
                       m_tangent_circle(i%).data(0).data0(0).tangent_coord(0), _
                        m_Circ(m_tangent_circle(i%).data(0).ele(0).no).data(0).data0.c_coord, _
                         depend_condition(point_, m_tangent_circle(i%).data(0).tangent_poi(0)), _
                          depend_condition(point_, m_Circ(m_tangent_circle(i%).data(0).ele(0).no).data(0).data0.center), _
                           condition, condition_color, 1, 0)  '切点和圆心连线
                  If m_poi(m_tangent_circle(i%).data(0).center).data(0).parent.last_element = 0 Then '自由点
                    m_tangent_circle(i%).data(0).data0(0).tangent_coord(0) = _
                     m_poi(m_tangent_circle(i%).data(0).tangent_poi(0)).data(0).data0.coordinate
                    Call distance_point_to_line(m_tangent_circle(i%).data(0).data0(0).circle_center, _
                          m_tangent_circle(i%).data(0).data0(0).tangent_coord(0), paral_, _
                            m_tangent_circle(i%).data(0).data0(0).tangent_coord(0), _
                          m_Circ(m_tangent_circle(i%).data(0).ele(0).no).data(0).data0.c_coord, 0, _
                            m_tangent_circle(i%).data(0).data0(0).circle_center)
                         m_poi(m_Circ(m_tangent_circle(i%).data(0).circle_no).data(0).data0.center). _
                              data(0).data0.coordinate = _
                               m_tangent_circle(i%).data(0).data0(0).circle_center
                          'm_poi(m_Circ(m_tangent_circle(i%).data(0).circle_no).data(0).data0.center).
                          '   data(0).is_change = True
                         Call change_m_point(m_Circ(m_tangent_circle(i%).data(0).circle_no).data(0).data0.center, True)
                      Call add_point_to_line(m_Circ(m_tangent_circle(i%).data(0).circle_no).data(0).data0.center, _
                             tl%, 0, False, False, 0, temp_record.record_data)
                  ElseIf m_poi(m_tangent_circle(i%).data(0).center).data(0).parent.last_element = 1 Then
                    If m_poi(m_tangent_circle(i%).data(0).center).data(0).parent.element(1).ty = line_ Then '直线上的点
                    ElseIf m_poi(m_tangent_circle(i%).data(0).center).data(0).parent.element(1).ty = circle_ Then '圆上的点
                    End If
                  End If
                 ElseIf m_tangent_circle(i%).data(0).ele(0).ty = line_ Then
                 End If
               End If
               out_coord = m_tangent_circle(i%).data(0).data0(0).tangent_coord(0)
               out_point_no% = m_tangent_circle(i%).data(0).tangent_poi(0)
               Call add_point_to_m_circle(m_tangent_circle(i%).data(0).tangent_poi(0), _
                m_tangent_circle(i%).data(0).circle_no, record0, 0)
             If m_tangent_circle(i%).data(0).ele(0).ty = circle_ Then
               Call add_point_to_m_circle(m_tangent_circle(i%).data(0).tangent_poi(0), _
                      m_tangent_circle(i%).data(0).ele(0).no, record0, 0)
             ElseIf m_tangent_circle(i%).data(0).ele(0).ty = line_ Then
               Call add_point_to_line(m_tangent_circle(i%).data(0).tangent_poi(0), m_tangent_circle(i%).data(0).ele(0).no, _
                     0, 0, True, 0, temp_record.record_data)
             End If
          End If
          If m_tangent_circle(i%).data(0).ele(1).no > 0 And m_tangent_circle(i%).data(0).ele(1).ty > 0 Then
            m_tangent_circle(i%).data(0).tangent_poi(1) = m_point_number( _
                   m_tangent_circle(i%).data(0).data0(0).tangent_coord(1), condition, 1, condition_color, "", _
                    depend_condition(circle_, m_tangent_circle(i%).data(0).circle_no), _
                     m_tangent_circle(i%).data(0).ele(1), tangent_point_of_circle, True)
                 Call add_point_to_m_circle(m_tangent_circle(i%).data(0).tangent_poi(1), _
                     m_tangent_circle(i%).data(0).circle_no, record0, 0)
             If m_tangent_circle(i%).data(0).ele(1).ty = circle_ Then
                 Call add_point_to_m_circle(m_tangent_circle(i%).data(0).tangent_poi(1), _
                      m_tangent_circle(i%).data(0).ele(1).no, record0, 0)
             ElseIf m_tangent_circle(i%).data(0).ele(1).ty = line_ Then
                 Call add_point_to_line(m_tangent_circle(i%).data(0).tangent_poi(1), m_tangent_circle(i%).data(0).ele(1).no, _
                     0, 0, True, 0, temp_record.record_data)
             End If

          End If
          read_tangent_circle = m_tangent_circle(i%).data(0).circle_no
          temp_tangent_circle = m_tangent_circle(i%)
   End If
   Next i%
   If read_tangent_circle_no% > 0 Then
          For j% = last_conditions.last_cond(0).tangent_circle_no To last_conditions.last_cond(1).tangent_circle_no + 1 Step -1
           If j% <> read_tangent_circle_no% Then
            Call draw_tangent_circle(j%, True)
           If j% > read_tangent_circle_no% Then
            m_tangent_circle(j% - 1) = m_tangent_circle(j%)
            Call draw_tangent_circle(i%, True)
           End If
           End If
          Next j%
          last_conditions.last_cond(1).tangent_circle_no = last_conditions.last_cond(1).tangent_circle_no + 1
           m_tangent_circle(last_conditions.last_cond(1).tangent_circle_no) = _
             temp_tangent_circle
   End If
End Function
Private Sub set_new_coordinate_for_tangent_line(n%, coord1 As POINTAPI, coord2 As POINTAPI, point_on_line_ty As Integer)
Dim t_coord(1) As POINTAPI
t_coord(0) = coord1
t_coord(1) = coord2
If point_on_line_ty = point_on_segement Then '点在线内
   If tangent_line(n%).data(0).visible = 2 Then
      tangent_line(n%).data(0).new_coordinate(0).X = 10000
      tangent_line(n%).data(0).new_coordinate(0).Y = 10000
      tangent_line(n%).data(0).new_coordinate(1).X = 10000
      tangent_line(n%).data(0).new_coordinate(1).Y = 10000
   Else
      tangent_line(n%).data(0).new_coordinate(0) = tangent_line(n%).data(0).coordinate(0)
      tangent_line(n%).data(0).new_coordinate(1) = tangent_line(n%).data(0).coordinate(1)
   End If
ElseIf point_on_line_ty = point_out_segement Then
   If tangent_line(n%).data(0).visible = 2 Or tangent_line(n%).data(0).visible >= 4 Then
      tangent_line(n%).data(0).new_coordinate(0) = t_coord(0)
      tangent_line(n%).data(0).new_coordinate(1) = t_coord(1)
   Else
   If is_same_POINTAPI(t_coord(0), tangent_line(n%).data(0).coordinate(0)) Then
      tangent_line(n%).data(0).new_coordinate(0) = t_coord(1)
      tangent_line(n%).data(0).new_coordinate(1) = tangent_line(n%).data(0).coordinate(1)
   ElseIf is_same_POINTAPI(t_coord(0), tangent_line(n%).data(0).coordinate(1)) Then
      tangent_line(n%).data(0).new_coordinate(0) = t_coord(1)
      tangent_line(n%).data(0).new_coordinate(1) = tangent_line(n%).data(0).coordinate(0)
   ElseIf is_same_POINTAPI(t_coord(1), tangent_line(n%).data(0).coordinate(0)) Then
       tangent_line(n%).data(0).new_coordinate(0) = t_coord(0)
      tangent_line(n%).data(0).new_coordinate(1) = tangent_line(n%).data(0).coordinate(1)
  ElseIf is_same_POINTAPI(t_coord(1), tangent_line(n%).data(0).coordinate(1)) Then
      tangent_line(n%).data(0).new_coordinate(0) = t_coord(0)
      tangent_line(n%).data(0).new_coordinate(1) = tangent_line(n%).data(0).coordinate(0)
   End If
   End If
Else
      tangent_line(n%).data(0).new_coordinate(0).X = 10000
      tangent_line(n%).data(0).new_coordinate(0).Y = 10000
      tangent_line(n%).data(0).new_coordinate(1).X = 10000
      tangent_line(n%).data(0).new_coordinate(1).Y = 10000
End If
End Sub
Public Function get_ratio_of_two_lines(coord10 As POINTAPI, coord11 As POINTAPI, _
                        coord20 As POINTAPI, coord21 As POINTAPI, Optional paral_or_verti As Integer = paral_) As Single
Dim tcoord(1) As POINTAPI
tcoord(0) = minus_POINTAPI(coord11, coord10)
tcoord(1) = minus_POINTAPI(coord21, coord20)
If paral_or_verti = verti_ Then
  tcoord(1) = verti_POINTAPI(tcoord(1))
End If
If Abs(tcoord(1).X) > 5 Then
   get_ratio_of_two_lines = tcoord(0).X / tcoord(1).X
ElseIf Abs(tcoord(1).Y) > 5 Then
   get_ratio_of_two_lines = tcoord(0).Y / tcoord(1).Y
Else
End If

End Function

Public Function get_ratio_of_point_on_line(point_no%, line_no%, related_p1%, related_p2%, start_point%) As Single
Dim t_coord(1) As POINTAPI '计算（point_no%--start_point%)/related_p1%--related-p2%)
If related_p1% > 0 And related_p2% > 0 Then
   t_coord(0) = m_poi(related_p1%).data(0).data0.coordinate
   t_coord(1) = m_poi(related_p2%).data(0).data0.coordinate
Else
   t_coord(0) = m_poi(m_lin(line_no%).data(0).data0.depend_poi(0)).data(0).data0.coordinate
   t_coord(1) = second_end_point_coordinate(line_no%)
End If
   If start_point% = 0 Then
    get_ratio_of_point_on_line = _
     get_ratio_of_two_lines(m_poi(point_no%).data(0).data0.coordinate, t_coord(0), _
         t_coord(1), t_coord(0), m_poi(point_no%).data(0).parent.inter_type)
   Else
    get_ratio_of_point_on_line = _
     get_ratio_of_two_lines(m_poi(point_no%).data(0).data0.coordinate, m_poi(start_point%).data(0).data0.coordinate, _
         t_coord(1), t_coord(0), m_poi(point_no%).data(0).parent.inter_type)
   End If
End Function

Public Function get_coordinate_of_point_on_line(point_no%, line_no%) As POINTAPI
Dim t_coord(1) As POINTAPI
If m_poi(point_no%).data(0).parent.related_point(0) > 0 And _
       m_poi(point_no%).data(0).parent.related_point(1) > 0 Then
   t_coord(0) = m_poi(m_poi(point_no%).data(0).parent.related_point(2)).data(0).data0.coordinate
   t_coord(1) = minus_POINTAPI(m_poi(m_poi(point_no%).data(0).parent.related_point(1)).data(0).data0.coordinate, _
                   m_poi(m_poi(point_no%).data(0).parent.related_point(0)).data(0).data0.coordinate)
   If m_poi(point_no%).data(0).parent.inter_type = verti_ Then
     t_coord(1) = verti_POINTAPI(t_coord(1))
   End If
Else
   t_coord(0) = m_poi(m_lin(line_no%).data(0).data0.depend_poi(0)).data(0).data0.coordinate
   t_coord(1) = minus_POINTAPI(second_end_point_coordinate(line_no%), _
                       t_coord(0))
End If
 get_coordinate_of_point_on_line = add_POINTAPI(t_coord(0), time_POINTAPI_by_number( _
       t_coord(1), m_poi(point_no%).data(0).parent.ratio))
     
End Function

Public Function tangent_line_no_from_temp_no(temp_no%) As Integer

End Function

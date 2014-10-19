Attribute VB_Name = "setwenti0"

Public Sub set_wenti_cond_71_70_69(tangent_circle_no%, c1%, c2%)
Dim inputcond_no%
Dim w_n%
Dim tangent_pointapi1 As POINTAPI
Dim tangent_pointapi2 As POINTAPI
Dim tangent_point_no(1) As Integer

If m_Circ(c1%).data(0).data0.center > 0 And m_Circ(c2%).data(0).data0.center > 0 Then
    inputcond_no% = -71
ElseIf m_Circ(c1%).data(0).data0.center > 0 And m_Circ(c2%).data(0).data0.center = 0 Then
    inputcond_no% = -69
ElseIf m_Circ(c1%).data(0).data0.center = 0 And m_Circ(c2%).data(0).data0.center > 0 Then
    Call exchange_two_integer(c1%, c2%)
    inputcond_no% = -69
Else
    inputcond_no% = -70
End If

If inputcond_no = -71 Then
   tangent_pointapi1 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_Circ(c1%).data(0).data0.center).data(0).data0.coordinate, _
            m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate), _
             (m_Circ(tangent_circle_no%).data(0).data0.radii / _
              (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c1%).data(0).data0.radii))))
   tangent_point_no(0) = set_point(tangent_pointapi1, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(0), c1%, record0, True)
   Call add_point_to_m_circle(tangent_point_no(0), tangent_circle_no%, record0, 1)
'*******************************************************************************************************
   tangent_pointapi2 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_Circ(c2%).data(0).data0.center).data(0).data0.coordinate, _
            m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate), _
             (m_Circ(tangent_circle_no%).data(0).data0.radii / _
              (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c2%).data(0).data0.radii))))
    tangent_point_no(1) = set_point(tangent_pointapi2, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(1), c2%, record0, 1)
   Call add_point_to_m_circle(tangent_point_no(1), tangent_circle_no%, record0, 1)
'****************************************************************************************************************
   Call C_display_wenti.set_m_no(0, -71, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.in_point(1), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 2, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 3, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.center, 4, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 5, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(0), 6, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(1), 7, True)
ElseIf inputcond_no% = -70 Then
   tangent_pointapi1 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
            m_Circ(tangent_circle_no%).data(0).data0.c_coord), (m_Circ(tangent_circle_no%).data(0).data0.radii / _
             (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c1%).data(0).data0.radii))))
   tangent_point_no(0) = set_point(tangent_pointapi1, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(0), c1%, record0, 1)
   Call add_point_to_m_circle(tangent_point_no(0), tangent_circle_no%, record0, 1)
'*******************************************************************************************************
   tangent_pointapi2 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(tangent_circle_no%).data(0).data0.c_coord), (m_Circ(tangent_circle_no%).data(0).data0.radii / _
             (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c2%).data(0).data0.radii))))
    tangent_point_no(1) = set_point(tangent_pointapi2, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(1), c2%, record0, 1)
   Call add_point_to_m_circle(tangent_point_no(1), tangent_circle_no%, record0, 1)
'****************************************************************************************************************
   Call C_display_wenti.set_m_no(0, -70, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.in_point(1), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 2, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(2), 3, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(3), 4, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 5, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 6, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 7, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(0), 8, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(1), 9, True)
'***************************************************************************************************************
ElseIf inputcond_no% = -69 Then
   tangent_pointapi1 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_Circ(c1%).data(0).data0.center).data(0).data0.coordinate, _
            m_Circ(tangent_circle_no%).data(0).data0.c_coord), (m_Circ(tangent_circle_no%).data(0).data0.radii / _
             (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c1%).data(0).data0.radii))))
   tangent_point_no(0) = set_point(tangent_pointapi1, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(0), c1%, record0, True)
   Call add_point_to_m_circle(tangent_point_no(0), tangent_circle_no%, record0, 1)
'*******************************************************************************************************
   tangent_pointapi2 = add_POINTAPI(m_poi(m_Circ(tangent_circle_no%).data(0).data0.center).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(tangent_circle_no%).data(0).data0.c_coord), (m_Circ(tangent_circle_no%).data(0).data0.radii / _
             (m_Circ(tangent_circle_no%).data(0).data0.radii + m_Circ(c2%).data(0).data0.radii))))
    tangent_point_no(1) = set_point(tangent_pointapi2, 1, condition_color, 0, "")
   Call add_point_to_m_circle(tangent_point_no(1), c2%, record0, True)
   Call add_point_to_m_circle(tangent_point_no(1), tangent_circle_no%, record0, 1)
'****************************************************************************************************************
   Call C_display_wenti.set_m_no(0, -69, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(tangent_circle_no%).data(0).data0.in_point(1), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 2, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 3, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 4, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 5, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 6, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(0), 7, True)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_point_no(1), 8, True)

End If
End Sub
Public Sub set_wenti_cond2_3(p0%, p1%, p2%, p3%, paral_or_verti As Integer, w_n%)
'2 □□∥□□
'3 □□⊥□□
Dim A!
Dim r!
Dim i%, b1_x%, b1_y%, b2_x%, b2_y%
Dim tp(3) As Integer
Dim tp1(3) As Integer
Dim l(1) As Integer
Dim p_coord(1) As POINTAPI
Dim t_coord As POINTAPI
Dim chose_degree%
Dim paral_or_verti_(1) As Integer
Dim wenti_ty As Byte
Dim c_cond As condition_data_type
Dim temp_record As total_record_type
'On Error GoTo set_wenti_cond2_3_error
    l(1) = line_number0(p1%, p2%, 0, 0)
    l(0) = line_number0(p0%, p3%, 0, 0)
If m_poi(p0%).data(0).parent.inter_type = interset_point_line_line Then
    Call set_wenti_cond_22_23(p3%, p1%, p2%, m_poi(p0%).data(0).parent.element(1).no, m_poi(p0%).data(0).parent.element(2).no, p0%, 0, w_n%)
ElseIf m_poi(p0%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
         m_poi(p0%).data(0).parent.inter_type = new_point_on_line_circle21 Then
    Call set_wenti_cond10_16(p3%, p1%, p2%, m_poi(p0%).data(0).parent.element(2).no, p0%, 0, paral_or_verti, _
         m_poi(p0%).data(0).parent.inter_type, w_n%)
Else
'End If
If w_n% = 0 Then
tp1(0) = p0%
tp1(1) = p3%
tp1(2) = p1%
tp1(3) = p2%
chose_degree = 1
set_wenti_cond2_3_back:
 If arrange_points_by_degree(tp1(), tp(), 4, chose_degree%) = False Then '调整输入顺序
  If tp(0) = 0 And chose_degree = 1 Then
     chose_degree = 2
     GoTo set_wenti_cond2_3_back
  End If
 End If
Else
   tp(0) = p3%
   tp(1) = p0%
   tp(2) = p1%
   tp(3) = p2%
End If
   If tp(2) < tp(3) Then
      Call exchange_two_integer(tp(2), tp(3))
   End If
If w_n% = 0 Then
    Call C_display_wenti.set_m_no(0, paral_or_verti, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(3)
    Call C_display_wenti.set_m_point_no(w_n%, p3%, 1, True) 'temp_point(0)
    Call C_display_wenti.set_m_point_no(w_n%, p1%, 2, True) 'temp_point(2)
    Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True) 'temp_point(5)
    Call C_display_wenti.set_m_inner_lin(w_n%, l(1), 1)
    Call C_display_wenti.set_m_inner_lin(w_n%, l(0), 2)
    Call C_display_wenti.set_m_inner_poi(w_n%, tp(0), 1)
    Call C_display_wenti.set_m_inner_poi(w_n%, tp(1), 2)
    Call C_display_wenti.set_m_inner_poi(w_n%, tp(2), 3)
    Call C_display_wenti.set_m_inner_poi(w_n%, tp(3), 4)
'        Call C_display_wenti.set_m_inner_lin(w_n%, l(0), 3)
   temp_record.record_data.data0.condition_data.condition_no = 0
   temp_record.record_.display_no = w_n%
'******************************************************************************************
 '******************************************************************************************
 
    wenti_ty = 1 '第一次输入
Else
    l(1) = line_number(p1%, p2%, pointapi0, pointapi0, _
                       depend_condition(point_, p1%), _
                       depend_condition(point_, p2%), _
                       condition, condition_color, 1, 0)
    l(0) = line_number(p0%, p3%, pointapi0, pointapi0, _
                       depend_condition(point_, p0%), _
                       depend_condition(point_, p3%), _
                       condition, condition_color, 1, 0)
    Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 1)  '表示新点的位置
    Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 2)
    Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 3)
    Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 4)
    tp(0) = p0%
    tp(1) = p1%
    tp(2) = p2%
    tp(3) = p3%
End If
'******************************************************************************************
'   If wenti_ty = 0 Then ' 首次输入
       ' Call C_display_wenti.set_m_inner_point_type(w_n%, 0)
       'If C_display_wenti.m_no(C_display_wenti.m_last_input_wenti_no) = 3 Then '垂直
       '            If (b1_y% >= 0 And b2_x% >= 0) Or (b1_y% <= 0 And b2_x% >= 0) Then
       '             A! = -A!
       '            End If
       ' Call vertical_line(l(0), l(1), True, True)
       'Else '平行
       '            If (b1_x% >= 0 And b2_x <= 0) Or (b1_x% <= 0 And b2_x% >= 0) Then
       '             A! = -A!
       '            End If
       ' Call paral_line(l(0), l(1), True, True)
       'End If
    'Else 'wenti_ty=0 有调整过程
         'If C_display_wenti.m_no(C_display_wenti.m_last_input_wenti_no) = paral_ Then '平行
         '     paral_or_verti_(0) = paral_
         '     paral_or_verti_(1) = verti_
         '     Call paral_line(l(0), l(1), True, True)
         'Else '垂直
         '     paral_or_verti_(0) = verti_
         '     paral_or_verti_(1) = paral_
         '     Call vertical_line(l(0), l(1), True, True)
         'End If
             If C_display_wenti.m_inner_point_type(w_n%) > 0 And _
                   m_poi(tp(0)).data(0).parent.co_degree <= 2 Then '有交点
                If m_poi(tp(0)).data(0).parent.element(1).ty = circle_ And _
                   m_poi(tp(0)).data(0).parent.element(1).no > 0 Then '与圆相交
                 Call C_display_wenti.set_m_inner_circ(w_n%, _
                      m_poi(tp(0)).data(0).parent.element(1).no, 1)
                 If inter_point_line_circle3(m_poi(tp(1)).data(0).data0.coordinate, _
                     paral_or_verti_(0), m_poi(tp(2)).data(0).data0.coordinate, _
                     m_poi(tp(3)).data(0).data0.coordinate, _
                      m_Circ(m_poi(tp(0)).data(0).parent.element(1).no).data(0).data0, _
                         p_coord(0), 0, p_coord(1), 0, 0, True) Then
                   If distance_of_two_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, p_coord(0)) < _
                       distance_of_two_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, p_coord(1)) Then
                      Call set_point_coordinate(tp(0), p_coord(0), True)
                      Call C_display_wenti.set_m_inner_point_type(w_n%, 1)
                   Else
                      Call set_point_coordinate(tp(0), p_coord(1), True)
                      Call C_display_wenti.set_m_inner_point_type(w_n%, 2)
                   End If
                 Else
                  If chose_degree = 1 Then
                   chose_degree = 2
                    GoTo set_wenti_cond2_3_back
                  Else
                    Exit Sub
                  End If
                End If
            ElseIf m_poi(tp(0)).data(0).parent.element(1).ty = line_ And _
                      m_poi(tp(0)).data(0).parent.element(1).no > 0 Then '与线相交
                Call C_display_wenti.set_m_inner_lin(w_n%, _
                      m_poi(tp(0)).data(0).parent.element(1).no, 1)
                Call inter_point_line_line3(tp(1), paral_or_verti_(0), l(1), _
                   m_lin(m_poi(tp(0)).data(0).parent.element(1).no).data(0).data0.poi(0), _
                        paral_, m_poi(tp(0)).data(0).parent.element(1).no, p_coord(0), tp(0), _
                         True, c_cond, False)
            Else
                If inter_point_line_line3(tp(1), paral_or_verti_(0), l(1), _
                               tp(0), paral_or_verti_(1), l(1), p_coord(0), 0, True, _
                                c_cond, False) Then
                Else
                End If
                b1_x% = p_coord(0).X - m_poi(tp(1)).data(0).data0.coordinate.X
                b1_y% = p_coord(0).Y - m_poi(tp(1)).data(0).data0.coordinate.Y
                b2_x% = m_poi(tp(2)).data(0).data0.coordinate.X - _
                          m_poi(tp(3)).data(0).data0.coordinate.X
                b2_y% = m_poi(tp(2)).data(0).data0.coordinate.Y - _
                          m_poi(tp(3)).data(0).data0.coordinate.Y
                If paral_or_verti_(0) = True Then
                   If (b1_x% >= 0 And b2_x <= 0) Or (b1_x% <= 0 And b2_x% >= 0) Then
                    A! = -A!
                   End If
                   p_coord(1) = add_POINTAPI(m_poi(tp(1)).data(0).data0.coordinate, _
                         time_POINTAPI_by_number(minus_POINTAPI( _
                           m_poi(tp(2)).data(0).data0.coordinate, _
                            m_poi(tp(3)).data(0).data0.coordinate), A!))
                Else
                   If (b1_y% >= 0 And b2_x% >= 0) Or (b1_y% <= 0 And b2_x% >= 0) Then
                    A! = -A!
                   End If
                    p_coord(1) = add_POINTAPI(m_poi(tp(1)).data(0).data0.coordinate, _
                                     verti_POINTAPI(time_POINTAPI_by_number(minus_POINTAPI( _
                                      m_poi(tp(2)).data(0).data0.coordinate, _
                                       m_poi(tp(3)).data(0).data0.coordinate), A!)))
               End If
                  Call set_point_coordinate(tp(0), p_coord(1), True)
                  Call C_display_wenti.set_m_inner_point_type(w_n%, 0)
                  
       End If
                  Call draw_again0(Draw_form, 1)
     End If
 If wenti_ty = 1 Then
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
        draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End If
set_wenti_cond2_3_error:
End If
          temp_record.record_data.data0.condition_data.condition_no = 1
          temp_record.record_data.data0.condition_data.condition(1).ty = wenti_cond_
          temp_record.record_data.data0.condition_data.condition(1).no = w_n%
          temp_record.record_.display_no = w_n%
           If paral_or_verti = paral_ Then
            Call set_dparal(l(0), l(1), temp_record, 0, 0, False)
          ElseIf paral_or_verti = verti_ Then
            Call set_dverti(l(0), l(1), temp_record, 0, 0, False)
          End If

End Sub
Public Sub set_wenti_cond_16_12_9_8(list_type%, p1%, p2%, p3%, p4%, p5%, p6%)
Dim w_n%, i%, j%
Dim tp(5) As Integer
Dim input_ty_no%
Dim temp_record As total_record_type
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
tp(4) = p5%
tp(5) = p6%
   temp_record.record_data.data0.condition_data.condition_no = 1
   temp_record.record_data.data0.condition_data.condition(1).ty = wenti_cond_
   'temp_record.record_data.data0.condition_data.condition(1).no = w_n%
   'temp_record.record_.display_no = w_n%
If list_type% = 1 Then
   input_ty_no% = -16
    Call C_display_wenti.set_m_no(0, -16, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
       temp_record.record_data.data0.condition_data.condition(1).no = w_n%
For i% = 0 To 2
    Call set_equal_dline(tp(i%), tp((i% + 1) Mod 3), tp((i% + 1) Mod 3), tp((i% + 2) Mod 3), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
    Call set_angle_value(Abs(angle_number(tp(i%), tp((i% + 1) Mod 3), tp((i% + 2) Mod 3), 0, 0)), "60", temp_record, 0, 0, True)
    'Call set_angle_value(Abs(angle_number(p2%, p3%, p1%, 0, 0)), "60", temp_record, 0, 0, True)
    'Call set_angle_value(Abs(angle_number(p3%, p1%, p2%, 0, 0)), "60", temp_record, 0, 0, True)
Next i%
ElseIf list_type% = 2 Then
   input_ty_no% = -12
    Call C_display_wenti.set_m_no(0, -12, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(3), 3, True)
       temp_record.record_data.data0.condition_data.condition(1).no = w_n%
  For i% = 0 To 3
      For j% = 1 To 2
       Call set_equal_dline(tp(i%), tp((i% + 1) Mod 4), tp((i% + j% + 1) Mod 4), tp((i% + j% + 2) Mod 4), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
     Next j%
    Call set_angle_value(Abs(angle_number(tp(i%), tp((i% + 1) Mod 4), tp((i% + 2) Mod 4), 0, 0)), "90", temp_record, 0, 0, True)
'********************************************************************************************************
  Next i%
  For i% = 0 To 3
  For j% = 1 To 2
    Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 4), tp(i%), tp((i% + j% + 1) Mod 4), 0, 0)), "45", temp_record, 0, 0, True)
   Next j%
  Next i%
    Call set_dverti(line_number0(p1%, p3%, 0, 0, False), line_number0(p2%, p4%, 0, 0, False), temp_record, 0, 0, False)
    Call set_equal_dline(tp(0), tp(2), tp(1), tp(3), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
ElseIf list_type% = 3 Then
    Call C_display_wenti.set_m_no(0, -9, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(3), 3, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(4), 4, True)
       temp_record.record_data.data0.condition_data.condition(1).no = w_n%
For i% = 0 To 4
  For j% = 1 To 3
    Call set_equal_dline(tp(i%), tp((i% + j%) Mod 5), tp((i% + j%) Mod 5), tp((i% + j% + 1) Mod 5), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
  Next j%
    Call set_angle_value(Abs(angle_number(tp(i%), tp((i% + 1) Mod 5), tp((i% + 2) Mod 5), 0, 0)), "108", temp_record, 0, 0, True)
Next i%
For i% = 0 To 4
 For j% = 1 To 3
      Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 5), tp(i%), tp((i% + j% + 1) Mod 5), 0, 0)), "36", temp_record, 0, 0, True)
 Next j%
Next i%
For i% = 0 To 4
For j% = 1 To 2
    Call set_equal_dline(tp(i%), tp((i% + 2) Mod 5), tp((i% + j%) Mod 5), tp((i% + j% + 2) Mod 5), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
    Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 5), tp(i%), tp((i% + j% + 2) Mod 5), 0, 0)), "72", temp_record, 0, 0, True)
 Next j%
Next i%

ElseIf list_type% = 4 Then
    Call C_display_wenti.set_m_no(0, -8, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(3), 3, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(4), 4, True)
    Call C_display_wenti.set_m_point_no(w_n%, tp(5), 5, True)
       temp_record.record_data.data0.condition_data.condition(1).no = w_n%
For i% = 0 To 5
  For j% = 1 To 5
    Call set_equal_dline(tp(i%), tp((i% + j%) Mod 6), tp((i% + j%) Mod 6), tp((i% + j% + 1) Mod 6), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
  Next j%
    Call set_angle_value(Abs(angle_number(tp(i%), tp((i% + 1) Mod 6), tp((i% + 2) Mod 6), 0, 0)), "120", temp_record, 0, 0, True)
Next i%
'********************************************************************************************************
For i% = 0 To 5
 For j% = 1 To 4
      Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 6), tp(i%), tp((i% + j% + 1) Mod 6), 0, 0)), "30", temp_record, 0, 0, True)
 Next j%
Next i%
For i% = 0 To 5
For j% = 1 To 3
    Call set_equal_dline(tp(i%), tp((i% + 2) Mod 6), tp((i% + j%) Mod 6), tp((i% + j% + 2) Mod 6), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
    Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 6), tp(i%), tp((i% + j% + 2) Mod 6), 0, 0)), "60", temp_record, 0, 0, True)
 Next j%
Next i%
For i% = 0 To 5
For j% = 1 To 2
    Call set_equal_dline(tp(i%), tp((i% + 3) Mod 6), tp((i% + j%) Mod 6), tp((i% + j% + 3) Mod 6), 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
    Call set_angle_value(Abs(angle_number(tp((i% + j%) Mod 6), tp(i%), tp((i% + j% + 3) Mod 6), 0, 0)), "90", temp_record, 0, 0, True)
 Next j%
Next i%
'    Call set_dverti(line_number0(p1%, p3%, 0, 0, False), line_number0(p2%, p4%, 0, 0, False), temp_record, 0, 0, False)
End If
End Sub

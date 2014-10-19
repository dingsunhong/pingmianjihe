Attribute VB_Name = "setwenti"
Option Explicit
Public Sub set_wenti_cond6_6(Wno%, ByVal p1%, ByVal p2%)
'6 □是线段□□上分比为!_~的分点
Dim i%, k%, s$, l%
Dim tn(3) As Integer
Dim ts$
Dim r1%, w_n%
Dim r2%, tp1%, tp2%
record_0.data0.condition_data.condition_no = 0 'record0
'If m_poi(p2%).data(0).parent.inter_type = new_free_point Then
' Wno% = 6
'End If
tp1% = p1%
tp2% = p2%
If p1% > 0 And p2% > 0 Then
  If Wno% <> -6 Then
        Wno% = 6 '
  Else
    If Wno% = -6 Then  '
      If m_poi(p1%).data(0).parent.co_degree < m_poi(p2%).data(0).parent.co_degree Then
       Call input_from_point_ty(p2%)  '因为要交换p1,p2所以要先显示p2%
        Call exchange_two_integer(p1%, p2%)
      Else
         If m_poi(p2%).data(0).parent.inter_type = new_point_on_line Then
           Wno% = -43
         ElseIf m_poi(p2%).data(0).parent.inter_type = new_point_on_circle Then
          If m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.center > 0 Then
            Wno% = -57
          Else
            Wno% = -42
          End If
        End If
        If m_poi(p2%).data(0).parent.inter_type = interset_point_line_line Or _
           m_poi(p2%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
             m_poi(p2%).data(0).parent.inter_type = new_point_on_line_circle21 Or _
          m_poi(p2%).data(0).parent.inter_type = new_point_on_circle_circle12 Or _
           m_poi(p2%).data(0).parent.inter_type = new_point_on_circle_circle21 Then
            Ratio_for_measure.is_fixed_ratio = True
         End If
      End If
  End If
  End If
 If Wno% = -6 Then
        Call set_wenti_cond_6(p1%, p2%, "", 0, 0)
          Exit Sub
 ElseIf Wno% = -57 Then '
         Call C_display_wenti.set_m_no(0, -57, w_n%)
         Call C_display_wenti.set_m_point_no(w_n%, _
              m_Circ(m_poi(p2%).data(0).in_circle(1)).data(0).parent.element(0).no, 0, True)
         Call C_display_wenti.set_m_point_no(w_n%, _
              m_Circ(m_poi(p2%).data(0).in_circle(1)).data(0).parent.element(1).no, 1, True)
         Call C_display_wenti.set_m_point_no(w_n%, _
              m_Circ(m_poi(p2%).data(0).in_circle(1)).data(0).parent.element(2).no, 1, True)
         Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True)
         Call C_display_wenti.set_m_point_no(w_n%, p1%, 4, True)
         Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1)
         Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2)
         Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(p2%).data(0).in_circle(1), 1)
 ElseIf Wno% = -43 Then '
         Call C_display_wenti.set_m_no(0, -43, w_n%)
        Call C_display_wenti.set_m_point_no(w_n%, _
              m_lin(m_poi(p2%).data(0).in_line(1)).data(0).data0.poi(0), 0, True)
        Call C_display_wenti.set_m_point_no(w_n%, _
              m_lin(m_poi(p2%).data(0).in_line(1)).data(0).data0.poi(1), 1, True)
        Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)
        Call C_display_wenti.set_m_point_no(w_n%, p1%, 3, True)
        Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1)
        Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2)
        Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(p2%).data(0).in_line(1), 1)
 ElseIf Wno% = -42 Then '
          Call C_display_wenti.set_m_no(0, -42, w_n%)
         Call C_display_wenti.set_m_point_no(w_n%, _
              m_Circ(m_poi(p2%).data(0).in_circle(1)).data(0).parent.element(0).no, 0, True)
         Call C_display_wenti.set_m_point_no(w_n%, _
              m_Circ(m_poi(p2%).data(0).in_circle(1)).data(0).parent.element(1).no, 1, True)
         Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)
         Call C_display_wenti.set_m_point_no(w_n%, p1%, 3, True)
         Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1)
         Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2)
         Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(p2%).data(0).in_circle(1), 1)
 ElseIf Wno% = 6 Then
   If tp2% = p2% Then
    Call input_from_point_ty(tp2%)
   End If
   Call C_display_wenti.set_m_no(0, 6, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_point_number(divide_POINTAPI_by_number( _
                  add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                    m_poi(p2%).data(0).data0.coordinate), 2), condition, 0, condition_color, _
                     "", depend_condition(point_, p1%), depend_condition(point_, p2%), _
                      0, True), 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)
  ' Call set_son_data(wenti_cond_, w_n%, m_poi(p1%).data(0).sons)
  ' Call set_son_data(wenti_cond_, w_n%, m_poi(p2%).data(0).sons)
   Call C_display_wenti.set_m_inner_poi(w_n%, last_conditions.last_cond(1).point_no, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 2)
   Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
   l% = line_number0(p1%, p2%, 0, 0)
   Call C_display_wenti.set_m_inner_lin(w_n%, l%, 1)
 End If
 End If
Call Wenti_form.Picture1.SetFocus
End Sub
Public Sub set_wenti_cond7(ByVal c%, ByVal p%)
'7 ⊙□[down\\(_)]上任取一点□
'-61⊙□□□上任取一点□
Dim w_n%
If c% > 0 And p% > 0 Then
If m_Circ(c%).data(0).data0.center = p% Then
   Exit Sub
'ElseIf is_point_in_circle(c%, 0, p%, 0, 0) Then
 '  Exit Sub
Else
   If m_Circ(c%).data(0).data0.center = 0 Then  '无心圆
   Call C_display_wenti.set_m_no(0, -61, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 2, True)
   Call C_display_wenti.set_m_point_no(w_n%, p%, 3, True)
   ElseIf m_Circ(c%).data(0).data0.in_point(0) > 0 Then
   Call C_display_wenti.set_m_no(0, 7, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, p%, 2, True)
   Else
    Exit Sub
   End If
   Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p%, 1)
Call Wenti_form.Picture1.SetFocus
End If
End If
End Sub

Public Sub set_wenti_cond_22_23(ByVal p0%, ByVal p1%, ByVal p2%, _
    ByVal l1%, ByVal l2%, ByVal p5%, ty As Integer)
'-22 过□点平行□□的直线交□□于□
'-23过□点垂直□□的直线交□□于□
'Dim tl(1) As Integer
Dim tl(1) As Integer
Dim p3%, p4%, w_n%
Dim i%
If is_point_in_points(p0%, m_lin(l1%).data(0).data0.in_point) > 0 Then
   Call exchange_two_integer(l1%, l2%)
End If
If m_lin(l1%).data(0).data0.poi(1) > 0 Then
 p3% = m_lin(l1%).data(0).data0.poi(0)
 p4% = m_lin(l1%).data(0).data0.poi(1)
Else
  p3% = m_lin(l1%).data(0).data0.in_point(1)
 p4% = m_lin(l1%).data(0).data0.in_point(m_lin(l1%).data(0).data0.in_point(0))
End If
'***************************************************************************
If ty = paral_ Then
  Call C_display_wenti.set_m_no(0, -22, w_n%) '15
Else
  Call C_display_wenti.set_m_no(0, -23, w_n%)
End If
  Call C_display_wenti.set_m_inner_lin(w_n%, l1, 1) '标准线
  Call C_display_wenti.set_m_inner_lin(w_n%, l2%, 2) '交线
  Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p1%, p2%, 0, 0), 3) '连线
  Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 1)
  Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 2)
  Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(3)
  Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) 'temp_point(0)
  Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'temp_point(2)
  Call C_display_wenti.set_m_point_no(w_n%, p3%, 3, True) 'm_lin(temp_line(2)).poi(0)
  Call C_display_wenti.set_m_point_no(w_n%, p4%, 4, True) 'm_lin(temp_line(2)).poi(1)
  Call C_display_wenti.set_m_point_no(w_n%, p5%, 5, True) 'temp_point(5)
  'Call set_son_data(wenti_cond_, w_n%, m_poi(p0%).data(0).sons)
  'Call set_son_data(wenti_cond_, w_n%, m_lin(l1%).data(0).sons)
  'Call set_son_data(wenti_cond_, w_n%, m_lin(line_number0(p1%, p2%, 0, 0)).data(0).sons)
  'call set_parent(
If ty = -22 Then
  Call paral_line(tl(0), tl(1), True, True) '
Else
  Call vertical_line(tl(0), tl(1), True, True) '
End If
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
          draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond_3(ByVal c%, ByVal p2%, ByVal p5%, ByVal c1%, ty As Integer)
'-3 与⊙□[down\\(_)]相切于点□的切线交⊙□[down\\(_)]
'-29 与⊙□□□相切于□的切线交⊙□[down\\(_)]于□
'-28 与⊙□□□相切于□的切线交⊙□□□于□
'-26 与⊙□[down\\(_)]相切于□的切线交⊙□□□于□
Dim i%, w_n%, t_p%, tl%
Dim c_ty(1) As Integer
Dim tp(2) As Integer
Dim c_data0 As condition_data_type
Dim tp1(2) As Integer
If m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name >= "A" And _
    m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name <= "Z" Then
     tp(0) = m_Circ(c%).data(0).data0.center
     tp(1) = m_Circ(c%).data(0).data0.in_point(1)
Else
     tp(0) = m_Circ(c%).data(0).data0.in_point(1)
     tp(1) = m_Circ(c%).data(0).data0.in_point(2)
     tp(2) = m_Circ(c%).data(0).data0.in_point(3)
     c_ty(0) = 1
End If
If m_poi(m_Circ(c1%).data(0).data0.center).data(0).data0.name >= "A" And _
    m_poi(m_Circ(c1%).data(0).data0.center).data(0).data0.name <= "Z" Then
     tp1(0) = m_Circ(c1%).data(0).data0.center
     tp1(1) = m_Circ(c1%).data(0).data0.in_point(1)
Else
     tp1(0) = m_Circ(c1%).data(0).data0.in_point(1)
     tp1(1) = m_Circ(c1%).data(0).data0.in_point(2)
     tp1(2) = m_Circ(c1%).data(0).data0.in_point(3)
     c_ty(1) = 1
End If
If c_ty(0) = 0 And c_ty(0) = 0 Then
Call C_display_wenti.set_m_no(0, -3, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'circ(temp_circle(0)).data(0).data0.in_point(1)
Call C_display_wenti.set_m_point_no(w_n%, tp1(0), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(1), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 5, True)
ElseIf c_ty(0) = 1 And c_ty(1) = 0 Then
Call C_display_wenti.set_m_no(0, -29, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True) 'circ(temp_circle(0)).data(0).data0.in_point(1)
Call C_display_wenti.set_m_point_no(w_n%, tp1(0), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(1), 5, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 6, True)
ElseIf c_ty(0) = 0 And c_ty(1) = 1 Then
Call C_display_wenti.set_m_no(0, -26, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'circ(temp_circle(0)).data(0).data0.in_point(1)
Call C_display_wenti.set_m_point_no(w_n%, tp1(0), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(1), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(2), 5, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 6, True)
ElseIf c_ty(0) = 1 And c_ty(1) = 1 Then
Call C_display_wenti.set_m_no(0, -28, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True) 'circ(temp_circle(0)).data(0).data0.in_point(1)
Call C_display_wenti.set_m_point_no(w_n%, tp1(0), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(1), 5, True)
Call C_display_wenti.set_m_point_no(w_n%, tp1(2), 6, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 7, True)
End If
Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1) '第一圆
Call C_display_wenti.set_m_inner_circ(w_n%, c1%, 2) '第二圆
Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 1) '第二圆
Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2) '第二圆
tl% = line_number0(p2%, p5%, 0, 0)
Call C_display_wenti.set_m_inner_lin(w_n%, tl%, 1) '切线
t_coord = time_POINTAPI_by_number(m_poi(p2%).data(0).data0.coordinate, 2)
t_coord = minus_POINTAPI(t_coord, m_poi(p5%).data(0).data0.coordinate)
Call set_point_coordinate(t_p%, t_coord, False)
Call C_display_wenti.set_m_inner_poi(w_n%, t_p%, 3) '第二圆
Call add_point_to_line(t_p%, tl%, 0, False, False, 0)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
 draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond14(p0%, p1%, p2%, p3%)
'14 过□作直线□□的垂线垂足为□
Dim l(1) As Integer
Dim i%, w_n%
         l(0) = line_number0(p0%, p3%, 0, 0)
         l(1) = line_number0(p1%, p2%, 0, 0)
     Call C_display_wenti.set_m_no(0, 14, w_n%)
     Call C_display_wenti.set_m_inner_lin(w_n%, l(1), 1)
     Call C_display_wenti.set_m_inner_lin(w_n%, l(0), 2)
     Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 1)
     Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True)  'temp_point(3)
     Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)
     Call C_display_wenti.set_m_point_no(w_n%, p3%, 3, True)   'temp_point(5)
      ' For i% = 0 To 3
      ' If m_poi(C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no, i%)).data(0).parent.co_degree < 0 Then
      '  m_poi(C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no, i%)).data(0).degree = _
      '   m_poi(C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no, i%)).data(0).degree - 3
           End If
      ' Next i%
        Call vertical_line(l(0), l(1), True, True) '
       ' record0.condition_data.condition_no = 0
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
            draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond_33_44(ByVal p0%, _
      ByVal c%, ty As Integer, ByVal p4%, tangent_line_no%) '作切线
'c%,p4%切点，p4%,p5% 切线 ; ty , 切点类型0， 切点固定
'-33过□作⊙□[down\\(_)]的切线□□
'-44过□作⊙□□□的切线□□
Dim i%, tn%, w_n%
Dim A!
'On Error GoTo set_wenti_cond_33_error
If m_Circ(c%).data(0).data0.center > 0 Then
Call C_display_wenti.set_m_no(0, -33, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(0)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 1, True) 'Circ(temp_circle(0)).data(0).center
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 2, True) ' Circ(temp_circle(0)).data(0).data0.in_point(1)
Call C_display_wenti.set_m_point_no(w_n%, p4%, 3, True) 'temp_point(2)
Else
Call C_display_wenti.set_m_no(0, -44, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(0)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 1, True) 'Circ(temp_circle(0)).data(0).center
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 2, True) 'Circ(temp_circle(0)).data(0).center
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 3, True) 'Circ(temp_circle(0)).data(0).center
Call C_display_wenti.set_m_point_no(w_n%, p4%, 4, True) 'temp_point(2)
End If
Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1) '圆
Call C_display_wenti.set_m_inner_point_type(w_n%, ty)  '切线类型
Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 1) '
Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 2)
Call set_parent(circle_, c%, wenti_cond_, w_n%, 0)
Call set_parent(point_, p0%, wenti_cond_, w_n%, 0)
'Call C_display_wenti.set_m_point_no(w_n%, p4%, 35, True) '切点
'Call C_display_wenti.set_m_point_no(w_n%, p5%, 36, True)
Call C_display_wenti.set_m_inner_lin(w_n%, tangent_line(tangent_line_no%).data(0).line_no, 1)  '切线
If ty = tangent_line_by_point_on_circle Then
 Call set_parent(point_, p0%, wenti_cond_, w_n%, 0)
 Call set_parent(circle_, c%, wenti_cond_, w_n%, 0)
End If

'     m_poi(p5%).data(0).parent.related_circle = c%
'     m_poi(p5%).data(0).parent.ratio = _
      distance_of_two_POINTAPI(m_poi(p5%).data(0).data0.coordinate, m_poi(p4%).data(0).data0.coordinate) / _
       distance_of_two_POINTAPI(m_Circ(c%).data(0).data0.c_coord, m_poi(p0%).data(0).data0.coordinate)
' Call C_display_wenti.set_m_point_no(C_display_wenti.m_last_input_wenti_no, _
     Int(1000 * A!), 7, False)
't_coord = time_POINTAPI_by_number(m_poi(p4%).data(0).data0.coordinate, 2)
't_coord = minus_POINTAPI(t_coord, m_poi(p5%).data(0).data0.coordinate)
'Call set_point_coordinate(tn%, t_coord, False)
'Call get_new_char(tn%)
'Call C_display_wenti.set_m_inner_poi(w_n%, tn%, 3) 'temp_point(2)
record_0.data0.condition_data.condition_no = 0
Call add_tangent_line_to_circle(tangent_line_no%, c%, record0)
'Call add_point_to_line(tn%, line_number0(p4%, p5%, 0, 0), 0, False, False, 0)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
      draw_wenti_no = C_display_wenti.m_last_input_wenti_no
set_wenti_cond_33_error:
End Sub

Public Sub set_wenti_cond5_15(ByVal p0%, p3%, ByVal p1%, ByVal c%, ty%, w_n%)
'inter_point_ty :p2%的类型
'5 取线段□□的中点□
'以□□为直径作⊙□[down\\(_)]
Dim tl%
Dim p2%
Dim i%
Dim n%
Dim temp_record As total_record_type
Dim c_data As condition_data_type
Call input_from_point_ty(p1%)
If ty = 15 Then '输入园的直径
  p2% = m_Circ(c%).data(0).data0.center '取圆心点的序号
    Call get_new_char(p2%) '给圆心点命名
Else
    p2% = 0
End If
c_data.condition_no = 0
If is_mid_point(p0%, p2%, p1%, 0, 0, 0, 0, n%, -1000, 0, 0, 0, _
     0, 0, 0, Dmid_point_data0, "", 0, 0, 0, _
       c_data) Then '如果已经是中点
        p2% = Dmid_point(n%).data(0).data0.poi(1)
   If ty% = 5 Then '输入中点,退出
    Exit Sub
   End If
End If
'************************************************************
      MDIForm1.Toolbar1.Buttons(21).Image = 33
If p2% = 0 Then '设置线段中点
          t_coord = mid_POINTAPI( _
              m_poi(p0%).data(0).data0.coordinate, _
              m_poi(p1%).data(0).data0.coordinate) '读中点坐标
   p2% = m_point_number(t_coord, condition, 1, condition_color, "", _
                      depend_condition(point_, p0%), depend_condition(point_, p1%), _
                       0, True) '设置中点
   tl% = line_number(p0%, p1%, pointapi0, pointapi0, _
                       depend_condition(point_, p0%), _
                       depend_condition(point_, p1%), _
                       condition, condition_color, 1, 0)
  Call add_point_to_line(p2%, tl%, 0, display, True, 0)
End If
' Call add_point_to_line(p2%, line_number0(p0%, p1%, 0, 0), 0, True, True, 0, c_data) '将中点p2%加入线段p0%p1%
' m_poi(p2%).data(0).parent.co_degree = 2
  Call C_display_wenti.set_m_no(0, ty%, w_n%) '设语句号w_n%
 Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True)  '设置输入语句
 Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True)    '
 Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)    '
 Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 1)    '保存输入信息
 Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 2)    '
 Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 3)    '
 '*************************************************************************************
  temp_record.record_data.data0.condition_data.condition_no = 0
 temp_record.record_.display_no = -w_n%
 Call set_mid_point(p0%, p1%, p2%, 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 0)
 'Call set_initial_condition(w_n%, 0, True)
 If ty% = 15 Then '以为直径作圆
    Call C_display_wenti.set_m_point_no(w_n%, p0%, 3, True)
    m_Circ(c%).data(0).data0.center = p2%
    Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
 End If
     Call set_line_visible(tl%, 1)
     Call C_display_wenti.set_m_inner_lin(w_n%, tl%, 1)
     record_0.data0.condition_data.condition_no = 0
  Call set_parent(line_, tl%, point_, p2%, paral_, p0%, p1%, p0%)
     C_display_wenti.complete_set_inner_data (w_n%)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub



Public Sub set_wenti_cond_1(ByVal p0%, ByVal p1%, _
                                ByVal p2%, ByVal p3%, w_n%, _
                                 ByVal ele1_ty As Integer, ByVal ele1_no%, _
                                  ByVal ele2_ty As Integer, ByVal ele2_no%, _
                                    inter_point%)
'□□＝□□
'-54 □□的垂直平分线交□□于□
 '-31 在□□上取一点□使得□□＝□□
  '-32 与⊙□[down\\(_)]相切于点□的切线交直线□□于□
Dim i%, tc%, tl%, midpoint%
Dim A!
Dim tp(3) As Integer
Dim tp1(3) As Integer
Dim el As eline_data0_type
Dim cd As condition_data_type
Dim p_coord(1) As POINTAPI
Dim mid_coord As POINTAPI
Dim circ_data As circle_data_type
Dim c_data As condition_data_type
Dim chose_degree As Integer
Dim temp_record As total_record_type
If m_poi(p0%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
        m_poi(p0%).data(0).parent.inter_type = new_point_on_line_circle21 Then '-31
   Call C_display_wenti.set_m_no(0, -31, w_n%)
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.poi(0), 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.poi(1), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, p0%, 2, True) '设置输入语句中几何数据(点)
   Call C_display_wenti.set_m_point_no(w_n%, p1%, 3, True)
   Call C_display_wenti.set_m_point_no(w_n%, p2%, 4, True)
   Call C_display_wenti.set_m_point_no(w_n%, p3%, 5, True)
ElseIf m_poi(p0%).data(0).parent.inter_type = new_point_on_circle_circle12 Or _
        m_poi(p0%).data(0).parent.inter_type = new_point_on_circle_circle21 Then '-36
    If m_Circ(m_poi(p0%).data(0).parent.element(2).no).data(0).data0.center = 0 Then
    Call C_display_wenti.set_m_no(0, -58, w_n%)
     Call C_display_wenti.set_m_point_no(w_n%, _
                   m_Circ(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.in_point(1), 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, _
                   m_Circ(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.in_point(2), 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, _
                   m_Circ(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.in_point(3), 2, True)
     Call C_display_wenti.set_m_point_no(w_n%, p0%, 3, True) '设置输入语句中几何数据(点)
     Call C_display_wenti.set_m_point_no(w_n%, p1%, 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, p2%, 5, True)
     Call C_display_wenti.set_m_point_no(w_n%, p3%, 6, True)
    Else '-32
     Call C_display_wenti.set_m_no(0, -30, w_n%)
     Call C_display_wenti.set_m_point_no(w_n%, _
                   m_Circ(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.center, 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, _
                   m_Circ(m_poi(p0%).data(0).parent.element(1).no).data(0).data0.in_point(1), 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, p0%, 2, True) '设置输入语句中几何数据(点)
     Call C_display_wenti.set_m_point_no(w_n%, p1%, 3, True)
     Call C_display_wenti.set_m_point_no(w_n%, p2%, 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, p3%, 5, True)
     End If
Else 'If w_n% = 0 Then '第一次输入,即由图形输入
Call C_display_wenti.set_m_no(0, -1, w_n%) '设置输入语句号
Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) '设置输入语句中几何数据(点)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 3, True)
End If
    'Call C_display_wenti.set_m_point_no(w_n%, p3%, 5, True)
    'Call C_display_wenti.set_m_inner_lin(w_n%, ele1_no%, 1)
    'Call C_display_wenti.set_m_inner_circ(w_n%, ele2_no%, 1)
    Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 1)
    Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 2)
    Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
    Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 4)

'm_poi(p0%).data(0).degree = m_poi(p0%).data(0).degree - 1 '登长线段，端点，位于定圆上，自由度=1，若同时，位于直线（或圆上）自由度=0
    '        Call line_number(p0%, p1%, pointapi0, pointapi0, _
                             depend_condition(point_, p0%), _
                             depend_condition(point_, p1%), _
                             condition, condition_color, 1, 0) '连线
    '        Call line_number(p2%, p3%, pointapi0, pointapi0, _
                             depend_condition(point_, p2%), _
                             depend_condition(point_, p3%), _
                             condition, condition_color, 1, 0)
  Call set_parent(point_, p1%, ele1_ty, ele1_no, length_depended_by_two_points_, p2%, p3%) 'p1%圆心，半径等于线段p2%p3%的长
  Call line_number(temp_point(draw_step - 1).no, temp_point(draw_step).no, _
                                  pointapi0, pointapi0, _
                                  depend_condition(point_, temp_point(draw_step - 1).no), _
                                  depend_condition(point_, temp_point(draw_step).no), _
                                  condition, condition_color, 1, 0) '建立直线

   temp_record.record_data.data0.condition_data.condition_no = 0
   temp_record.record_.display_no = -w_n%
     Call set_equal_dline(p0%, p1%, p2%, p3%, 0, 0, 0, 0, 0, 0, 0, _
                 temp_record, 0, eline_, 0, 0, 0, False)
'************************************************
   draw_wenti_no = C_display_wenti.m_last_input_wenti_no
       Call C_display_wenti.complete_set_inner_data(w_n%)


End Sub
Public Sub set_wenti_cond_31(ByVal ele1_no%, ByVal ele2_no%, ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ty As Integer)
Dim m_no% ' 在□□上取一点□使得□□＝□□
Dim w_n%
If p1% < p2% Then
   Call exchange_two_integer(p1%, p2%)
End If

   Call C_display_wenti.set_m_no(0, -31, w_n%)
   Call C_display_wenti.set_m_inner_lin(w_n%, ele1_no%, 1)
   Call C_display_wenti.set_m_inner_circ(w_n%, ele2_no%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2)
   Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 3)
   Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 4)
   Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
   If m_poi(m_lin(ele1_no).data(0).data0.in_point(m_lin(ele1_no).data(0).data0.in_point(0))).data(0).data0.visible = 0 Then
   Call C_display_wenti.set_m_point_no(w_n%, _
         m_lin(ele1_no).data(0).data0.in_point(m_lin(ele1_no).data(0).data0.in_point(0) - 1), 0, True)
   Else
   Call C_display_wenti.set_m_point_no(w_n%, _
            m_lin(ele1_no).data(0).data0.in_point(m_lin(ele1_no).data(0).data0.in_point(0)), 0, True)
   End If
   If m_poi(m_lin(ele1_no).data(0).data0.in_point(1)).data(0).data0.visible = 0 Then
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(ele1_no).data(0).data0.in_point(2), 1, True)
   Else
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(ele1_no).data(0).data0.in_point(1), 1, True)
   End If
   Call C_display_wenti.set_m_point_no(w_n%, p1%, 2, True)
   Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True)
   Call C_display_wenti.set_m_point_no(w_n%, p3%, 4, True)
   Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
   Call set_parent(point_, p3%, wenti_cond_, w_n%, 0)
   Call set_parent(point_, p4%, wenti_cond_, w_n%, 0)
   Call set_parent(circle_, ele2_no, wenti_cond_, w_n%, 0)
   'Call set_son_data(wenti_cond_, w_n%, son_data0, line_, ele1_no, m_lin(ele1_no).data(0).sons)
   'Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, ele2_no, m_Circ(ele2_no).data(0).sons)

  ' Call C_display_wenti.set_m_point_no(w_n%, p4%, 6, True)
'            Call line_number(p1%, p2%, pointapi0, pointapi0, _
                             depend_condition(point_, p1%), _
                             depend_condition(point_, p2%), _
                             condition, condition_color, 1, 0)
'            Call line_number(p3%, p4%, pointapi0, pointapi0, _
                             depend_condition(point_, p3%), _
                             depend_condition(point_, p4%), _
                             condition, condition_color, 1, 0)
   draw_wenti_no = C_display_wenti.m_last_input_wenti_no
       Call C_display_wenti.complete_set_inner_data(w_n%)
   
End Sub
Public Sub set_wenti_cond_30_58(ByVal ele1%, ByVal ele2%, ByVal p1%, _
    ByVal p2%, ByVal p3%, ByVal p4%, ty As Integer) 'ele1 第一圆，ele2第二圆，ty交点类
'-30 在⊙□[down\\(_)]上取一点□使得□□＝□□
'-58 在⊙□□□上取一点□使得□□＝□□型
'-31 在□□上取一点□使得□□＝□□
Dim i%, w_n%, n%
If p1% < p2% Then
   Call exchange_two_integer(p1%, p2%)
End If
   If m_Circ(ele2%).data(0).data0.center > 0 Then
     Call C_display_wenti.set_m_no(0, -30, w_n%)
     n% = -30
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(ele2%).data(0).data0.center, 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(ele2%).data(0).data0.in_point(1), 1, True)
       Call C_display_wenti.set_m_point_no(w_n%, p1%, 2, True)
       Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True)
       Call C_display_wenti.set_m_point_no(w_n%, p3%, 4, True)
       Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
       'Call C_display_wenti.set_m_point_no(w_n%, p4%, 6, True)
   Else
     Call C_display_wenti.set_m_no(0, -58, w_n%)
     n% = -58
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(ele2%).data(0).data0.in_point(1), 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(ele2%).data(0).data0.in_point(2), 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(ele2%).data(0).data0.in_point(3), 2, True)
       Call C_display_wenti.set_m_point_no(w_n%, p1%, 3, True)
       Call C_display_wenti.set_m_point_no(w_n%, p2%, 4, True)
       Call C_display_wenti.set_m_point_no(w_n%, p1%, 5, True)
       Call C_display_wenti.set_m_point_no(w_n%, p3%, 6, True)
       Call C_display_wenti.set_m_point_no(w_n%, p4%, 7, True)
   End If
   Call C_display_wenti.set_m_inner_circ(w_n%, ele1%, 1)
   Call C_display_wenti.set_m_inner_circ(w_n%, ele2%, 2)
   Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 2)
   Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 3)
   Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 4)
     Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
     Call C_display_wenti.complete_set_inner_data(w_n%)
   Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p2%, m_poi(p2%).data(0).sons)
   Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p3%, m_poi(p3%).data(0).sons)
   Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p4%, m_poi(p4%).data(0).sons)
   Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, ele1%, m_Circ(ele1%).data(0).sons)
   Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, ele2%, m_Circ(ele2%).data(0).sons)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
    draw_wenti_no = C_display_wenti.m_last_input_wenti_no
 

End Sub

Public Sub set_wenti_cond8_71(ByVal c%, w_n%, Optional input_ty As Byte = 0)
'input_ty=0 由图输入生成语句 =1由语句输入生成图
'8 过点□、□、□作⊙
'-71 过点□、□、□作⊙□
Dim i%, wenti_no%, tl%, tc%
Dim tp(3) As Integer
Dim ty As Byte
If w_n% = 0 Then
   ty = 0
Else
   ty = 1
End If
  tp(0) = m_Circ(c%).data(0).data0.center
  tp(1) = m_Circ(c%).data(0).data0.in_point(1)
  tp(2) = m_Circ(c%).data(0).data0.in_point(2)
  tp(3) = m_Circ(c%).data(0).data0.in_point(3)
 If ty = 0 Then
 If m_Circ(c%).data(0).data0.center = 0 Then
   Call C_display_wenti.set_m_no(0, -71, w_n%)
   Call C_display_wenti.set_m_no_(w_n%, -71)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 2, True)
   wenti_no% = -71
 Else
   Call C_display_wenti.set_m_no(0, 8, w_n%)
   Call C_display_wenti.set_m_no_(w_n%, 8)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 1, True)
    wenti_no% = 8
 End If
 End If
 Exit Sub
 '************************************************************************************
  If ty = 0 Then
    operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
         draw_wenti_no = C_display_wenti.m_last_input_wenti_no
  Else
      Call C_display_wenti.set_m_no_(w_n%, -71)
   If m_poi(tp(0)).data(0).parent.co_degree = 0 Then
         Call C_display_wenti.set_m_inner_circ(w_n%, _
                m_circle_number(1, 0, pointapi0, _
           C_display_wenti.m_inner_poi(w_n%, 2), _
             C_display_wenti.m_inner_poi(w_n%, 3), _
              C_display_wenti.m_inner_poi(w_n%, 4), _
               0, 0, 0, 1, 1, condition, condition_color, True), 1)
   Else
    Call reduce_degree_for_point(tp(0), tp(1), tp(2), tp(3))
    If m_poi(tp(0)).data(0).parent.co_degree = 0 Then
        Call C_display_wenti.set_m_inner_circ(w_n%, _
             m_circle_number(1, 0, pointapi0, _
             C_display_wenti.m_inner_poi(w_n%, 2), _
               C_display_wenti.m_inner_poi(w_n%, 3), _
                C_display_wenti.m_inner_poi(w_n%, 4), _
                 0, 0, 0, 1, 1, condition, condition_color, True), 1)
    ElseIf m_poi(tp(0)).data(0).parent.co_degree = 1 Then
     If m_poi(tp(1)).data(0).degree > 0 Then
        Call exchange_two_integer(tp(1), tp(3))
     ElseIf m_poi(tp(2)).data(0).parent.co_degree <= 2 Then
        Call exchange_two_integer(tp(2), tp(3))
     ElseIf m_poi(tp(3)).data(0).parent.co_degree <= 2 Then
     Else
      Call C_display_wenti.set_m_no_(w_n%, 0)
      Exit Sub
     End If
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(1), 2)  'temp_point(0)
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(2), 3)  'temp_point(0)
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(3), 4)  'temp_point(0)
     '*****************************************************************************
     If m_poi(tp(0)).data(0).parent.element(1).ty = line_ Then
        tl% = m_poi(tp(0)).data(0).parent.element(1).no
        Call C_display_wenti.set_m_inner_lin(w_n%, tl%, 1)
        Call C_display_wenti.set_m_no_(w_n%, -7101)
        '1-tp(0)
        '3- tp(3)
        '2-circle
     ElseIf m_poi(tp(0)).data(0).parent.element(1).ty = line_ Then
        tc% = m_poi(tp(0)).data(0).parent.element(1).no
        Call C_display_wenti.set_m_inner_lin(w_n%, tc%, 2)
        Call C_display_wenti.set_m_no_(w_n%, -7101)
     End If
      Call change_picture(w_n%, depend_condition(0, 0), 0)
     If m_poi(tp(3)).data(0).parent.co_degree = 1 Then
        If m_poi(tp(3)).data(0).parent.element(1).ty = line_ Then
         tl% = m_poi(tp(0)).data(0).parent.element(0).no
          Call C_display_wenti.set_m_inner_lin(w_n%, tl%, 2)
        ElseIf m_poi(tp(3)).data(0).parent.element(1).ty = circle_ Then
         tc% = m_poi(tp(3)).data(0).parent.element(1).no
          Call C_display_wenti.set_m_inner_lin(w_n%, tc%, 3)
        End If
     End If
     '******************************************************************************
   Else 'm_poi(tp(0)).data(0).degree=0
     If m_poi(tp(1)).data(0).parent.co_degree > 0 And m_poi(tp(2)).data(0).parent.co_degree > 0 Then
        Call exchange_two_integer(tp(1), tp(3))
     ElseIf m_poi(tp(1)).data(0).parent.co_degree <= 2 And m_poi(tp(3)).data(0).parent.co_degree <= 2 Then
        Call exchange_two_integer(tp(2), tp(1))
     ElseIf m_poi(tp(2)).data(0).parent.co_degree <= 2 And m_poi(tp(3)).data(0).parent.co_degree <= 2 Then
     Else
      Call C_display_wenti.set_m_no_(w_n%, 0)
      Exit Sub
     End If
     If tp(2) > tp(3) Then
             Call exchange_two_integer(tp(2), tp(3))
     End If
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(1), 2)  'temp_point(0)
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(2), 3)  'temp_point(0)
     Call C_display_wenti.set_m_inner_poi(w_n%, tp(3), 4)  'temp_point(0)
'***************************************************************************************
     
'*********************************************************************************
   End If
   End If
     Call draw_again0(Draw_form, 1)
  End If
End Sub

Public Sub set_wenti_cond9(ByVal l1%, ByVal l2%, ByVal p4%)
'9 直线□□和直线□□交于点□
Dim w_n%
Dim i%, p0%, p1%, p2%, p3%
If m_lin(l1%).data(0).data0.in_point(1) <> p4% Then
   p0% = m_lin(l1%).data(0).data0.in_point(1)
Else
   p0% = m_lin(l1%).data(0).data0.in_point(2)
End If
If m_lin(l1%).data(0).data0.in_point(m_lin(l1%).data(0).data0.in_point(0)) <> p4% Then
   p1% = m_lin(l1%).data(0).data0.in_point(m_lin(l1%).data(0).data0.in_point(0))
Else
   p1% = m_lin(l1%).data(0).data0.in_point(m_lin(l1%).data(0).data0.in_point(0) - 1)
End If
'***********************
If m_lin(l2%).data(0).data0.in_point(1) <> p4% Then
   p2% = m_lin(l2%).data(0).data0.in_point(1)
Else
   p2% = m_lin(l2%).data(0).data0.in_point(2)
End If
If m_lin(l2%).data(0).data0.in_point(m_lin(l2%).data(0).data0.in_point(0)) <> p4% Then
   p3 = m_lin(l2%).data(0).data0.in_point(m_lin(l2%).data(0).data0.in_point(0))
Else
   p3 = m_lin(l2%).data(0).data0.in_point(m_lin(l2%).data(0).data0.in_point(0) - 1)
End If
If p0% = p1% Then
   Call set_wenti_cond1(l2%, p4%)
ElseIf p2% = p3% Then
   Call set_wenti_cond1(l1%, p4%)
Else
  Call C_display_wenti.set_m_no(0, 9, w_n%) '建立输入语句
  'set_m_point 编辑输入语句用
  Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) '前四点是两直线的端点)
  Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) '
  Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) '
  Call C_display_wenti.set_m_point_no(w_n%, p3%, 3, True) '
  Call C_display_wenti.set_m_point_no(w_n%, p4%, 4, True) '两直线的交点
  Call C_display_wenti.set_m_inner_lin(w_n%, l1%, 1)  ' 第一条直线的序号
  Call C_display_wenti.set_m_inner_lin(w_n%, l2%, 2)  '第二条直线的序号
  Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 1) '两直线的交点
  'Call set_son_data(wenti_cond_, w_n%, m_lin(l1%).data(0).sons)
  'Call set_son_data(wenti_cond_, w_n%, m_lin(l2%).data(0).sons)
   operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
          draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End If
End Sub
Private Sub set_wenti_interset_point(ByVal w_n%, ByVal ele1_ty As Byte, _
                ByVal ele1_no%, ByVal ele2_ty As Byte, ele2_no%, ByVal inter_point%, inter_ty As Integer)
If ele1_ty = line_ And ele2_ty = line_ Then
ElseIf ele1_ty = line_ And ele2_ty = circle_ Then
  Call C_display_wenti.set_m_inner_poi(w_n%, inter_point%, 1)
 If inter_ty = new_point_on_line_circle Then
  Call C_display_wenti.set_m_inner_point_type(w_n%, 0)
 ElseIf inter_ty = new_point_on_line_circle12 Then
  Call C_display_wenti.set_m_inner_point_type(w_n%, 1)
 ElseIf inter_ty = new_point_on_line_circle21 Then
  Call C_display_wenti.set_m_inner_point_type(w_n%, 2)
 End If
ElseIf ele1_ty = circle_ And ele2_ty = circle_ Then
ElseIf ele1_ty = circle_ And ele2_ty = line_ Then
 Call set_wenti_interset_point(w_n%, ele2_ty, ele2_no%, ele1_ty, ele1_no, inter_point%, inter_ty)
Else
  Exit Sub
End If
End Sub

Public Sub set_wenti_cond11(ByVal l%, ByVal cir%, ByVal p3%, p4%)
'11 □是直线□□与⊙□[down\\(_)]的一个交点
'-63 □是直线□□与⊙□□□的一个交点
Dim l1%, l2%, w_n%, p1%, p2%
Dim i%
If m_lin(l%).data(0).data0.poi(0) <> p3% Then
   p1% = m_lin(l%).data(0).data0.poi(0)
Else
   p1% = m_lin(l%).data(0).data0.in_point(2)
End If
If m_lin(l%).data(0).data0.poi(1) <> p3% Then
   p2% = m_lin(l%).data(0).data0.poi(1)
Else
   p2% = m_lin(l%).data(0).data0.in_point(m_lin(l%).data(0).data0.in_point(0) - 1)
End If
  If m_Circ(cir%).data(0).data0.center > 0 Then
  Call C_display_wenti.set_m_no(0, 11, w_n%)
  Call C_display_wenti.set_m_point_no(w_n%, p3%, 0, True) 'temp_point(3)
  Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) 'temp_point(0)
  Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'temp_point(2)
  Call C_display_wenti.set_m_point_no(w_n%, _
       m_Circ(cir%).data(0).data0.center, 3, True)  'm_circ(ele2%).data(0).center
  Call C_display_wenti.set_m_point_no(w_n%, _
       m_Circ(cir%).data(0).data0.in_point(1), 4, True) 'm_circ(ele2%).data(0).data0.in_point(1)
  Else
  Call C_display_wenti.set_m_no(0, -63, w_n%)
  Call C_display_wenti.set_m_point_no(w_n%, p3%, 0, True) 'temp_point(3)
  Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) 'temp_point(0)
  Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'temp_point(2)
  Call C_display_wenti.set_m_point_no(w_n%, _
       m_Circ(cir%).data(0).data0.in_point(1), 3, True) 'm_circ(ele2%).data(0).data0.in_point(1)
  Call C_display_wenti.set_m_point_no(w_n%, _
       m_Circ(cir%).data(0).data0.in_point(2), 4, True) 'm_circ(ele2%).data(0).data0.in_point(1)
  Call C_display_wenti.set_m_point_no(w_n%, _
       m_Circ(cir%).data(0).data0.in_point(3), 5, True) 'm_circ(ele2%).data(0).data0.in_point(1)
  End If
  Call C_display_wenti.set_m_inner_lin(w_n%, l%, 1)
  Call C_display_wenti.set_m_inner_circ(w_n%, cir%, 1)
  Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 1)
  Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 2)
  Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
 ' Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
  'Call set_son_data(wenti_cond_, w_n%, son_data0, line_, l%, m_lin(l%).data(0).sons)
  'Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, cir%, m_Circ(cir%).data(0).sons)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
          draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond10_16(ByVal p0%, ByVal p1%, _
  ByVal p2%, c%, ByVal p5%, cond_no%, paral_or_verti As Integer, ByVal ty As Integer)
'10 过□平行□□的直线交⊙□[down\\(_)]于□
'16 过□垂直□□的直线交⊙□[down\\(_)]于□
'-68 过□垂直□□的直线交⊙□□□于□
'-62 过□平行□□的直线交⊙□□□于□
Dim l1%, l2%, w_n%
Dim i%
   If m_Circ(c%).data(0).data0.center > 0 Then '两点圆
         If paral_or_verti = paral_ Then
                  Call C_display_wenti.set_m_no(0, 10, w_n%)
         Else 'If ty = verti_ Then
                  Call C_display_wenti.set_m_no(0, 16, w_n%)
         End If
         Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(0)
         Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) 'temp_point(0)
         Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'temp_point(2)
         Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 3, True) 'Circ(ele2%).data(0).center
         Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 4, True) 'Circ(ele2%).data(0).data0.in_point(1)
         Call C_display_wenti.set_m_point_no(w_n%, p5%, 5, True) 'temp_point(5)
   Else '三点圆
        If paral_or_verti = paral_ Then
         Call C_display_wenti.set_m_no(0, -62, w_n%)
        Else 'If ty = verti_ Then
         Call C_display_wenti.set_m_no(0, -68, w_n%)
        End If
         Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(0)
         Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True) 'temp_point(0)
         Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) 'temp_point(2)
         Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 3, True) 'Circ(ele2%).data(0).data0.in_point(1)
         Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 4, True) 'Circ(ele2%).data(0).data0.in_point(1)
         Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 5, True) 'Circ(ele2%).data(0).data0.in_point(1)
         Call C_display_wenti.set_m_point_no(w_n%, p5%, 6, True) 'temp_point(5)
   End If
         Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p0%, p5%, 0, 0), 2)
         Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p1%, p2%, 0, 0), 1)
         Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
         Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 1)
         Call C_display_wenti.set_m_inner_poi(w_n%, p0%, 2)
         Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 3)
         Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 4)
         Call C_display_wenti.set_m_inner_poi(w_n%, paral_or_verti, 5)
         Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
         'Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p0%, m_poi(p0%).data(0).sons)
         'Call set_son_data(wenti_cond_, w_n%, son_data0, Line, _
                         line_number0(p1%, p2%, 0, 0), m_lin(line_number0(p1%, p2%, 0, 0)).data(0).sons)
         'Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, c%, m_Circ(c%).data(0).sons)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
          draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond_6(ByVal p0%, ByVal p1%, ByVal value$, ByVal c%, w_n%)
'□□=!_~
Dim l%, i%
Dim temp_record As total_record_type
Dim t_coord As POINTAPI
 temp_record.record_data.data0.condition_data.condition_no = 0
  temp_record.record_.display_no = -(C_display_wenti.m_last_input_wenti_no)
 Call C_display_wenti.set_m_no(0, -6, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, p0%, 0, True) 'temp_point(0)
 Call C_display_wenti.set_m_point_no(w_n%, p1%, 1, True)   'temp_point(1)
 Call C_display_wenti.set_m_inner_lin(w_n%, _
      line_number(p0%, p1%, t_coord, t_coord, depend_condition(point_, p0%), depend_condition(point_, p1%), condition, _
            condition_color, 1, 0), 1)
   c% = m_circle_number(1, p0%, m_poi(p0).data(0).data0.coordinate, 0, 0, 0, _
                   distance_of_two_POINTAPI(m_poi(p0%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate), _
                    0, 0, 0, 0, 0, 0, False)
 Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
 Call add_point_to_m_circle(p1%, c%, record0, True)
 If value$ <> "" Then '未输入长度值
 Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1) '
 Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1) '
   Call C_display_wenti.set_m_point_no(w_n%, _
        m_poi(p1%).data(0).data0.coordinate.X - m_poi(p0%).data(0).data0.coordinate.X, 17, False) '原线长的横坐标
   Call C_display_wenti.set_m_point_no(w_n%, _
        m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p0%).data(0).data0.coordinate.Y, 18, False) '
l% = Len(value$) '输入的线长值
For i% = 1 To l%
 Call C_display_wenti.set_m_condition(w_n%, _
        Mid$(value$, i%, 1), i% + 1) '记录长度值的字符串
Next i%
 Call C_display_wenti.set_m_condition(w_n%, Chr(13), i% + 1) '字符串结束符
 Call line_number(p0%, p1%, pointapi0, pointapi0, _
                  depend_condition(point_, p0%), _
                  depend_condition(point_, p1%), _
                  condition, condition_color, 1, 0) '建立直线数据
record0.data0.condition_data.condition_no = 0
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
      draw_wenti_no = C_display_wenti.m_last_input_wenti_no
Else
' Call Wenti_form.Picture1.SetFocus
End If
End Sub

Public Sub set_wenti_cond1(ByVal line_no%, ByVal p2%)
'1 直线□□上任取一点□
Dim i%, w_n%
Dim r!
Dim p1%, p0%
'*******************************************************************
If m_poi(p2%).data(0).parent.co_degree = 1 And m_poi(p2%).data(0).parent.element(1).ty = line_ And _
               m_poi(p2%).data(0).parent.element(1).no = line_no% Then
'line_no%上除p2%外的两点
p0% = m_lin(line_no%).data(0).data0.poi(0)
p1% = m_lin(line_no%).data(0).data0.poi(1)

If p0% = p1% Or p2% = p0% Or p2 = p1% Or p1% = 0 Or p2% = 0 Or p0% = 0 Then
'p0% = p1% 直线是一点,p2% = p0% Or p2 = p1% 输入的新点与直线的端点重合
 Exit Sub
Else
       Call C_display_wenti.set_m_no(0, 1, w_n%) '设置输入语句
       Call C_display_wenti.set_m_point_no(w_n%, m_lin(line_no%).data(0).data0.poi(0), 0, True) '记录直线端点
       Call C_display_wenti.set_m_point_no(w_n%, m_lin(line_no%).data(0).data0.poi(1), 1, True)
       Call C_display_wenti.set_m_point_no(w_n%, p2%, 2, True) '记录直线上的任意点
       Call C_display_wenti.set_m_inner_lin(w_n%, line_no%, 1) '记录直线序号
       Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 1) '记录输出点
'       Call set_parent_data(line_, line_no%, m_poi(p2%).data(0).parent) '理论点上的父辈数据
       '记录点在直线上的相对位置（比值），若任意点变动此值，会变化，并重新记录
       ''If Abs(m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
               m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate.X) > 5 Then
       ''     m_poi(p2%).data(0).parent.ratio = _
              (m_poi(p2%).data(0).data0.coordinate.X - m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.X) / _
               (m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
                 m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.X)
       'Else
       '     m_poi(p2%).data(0).parent.ratio = _
              (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) / _
               (m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
                 m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.Y)
       'End If
       '添加点的后辈数据
       'Call set_son_data(wenti_cond_, w_n%, m_lin(line_no%).data(0).sons)
       '记录当前操作的点号
       operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
       '记录当前输入语句的序号
           draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End If
End If
End Sub
Public Sub set_wenti_cond_32(ByVal tangent_p%, ByVal l%, ByVal p%, c%)
'-32 与⊙□[down\\(_)]相切于点□的切线交直线□□于□
'-27 与⊙□□□相切于□的切线交直线□□于□
Dim i%
Dim tn%, w_n%
Dim tp(2) As Integer
If m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name >= "A" And _
     m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name <= "Z" Then
Call C_display_wenti.set_m_no(0, -32, w_n%)
   Call C_display_wenti.set_m_inner_lin(w_n%, l%, 1)
   Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, tangent_p%, 2)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 1, True)
   tn% = 2
Else
Call C_display_wenti.set_m_no(0, -27, w_n%)
   Call C_display_wenti.set_m_inner_lin(w_n%, l%, 1)
   Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, p%, 1)
   Call C_display_wenti.set_m_inner_poi(w_n%, tangent_p%, 2)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 2, True)
tn% = 3
End If
   Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(tangent_p%, p%, 0, 0), 2)
   Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(tangent_p%, _
                          m_Circ(c%).data(0).data0.center, 0, 0), 3)
   Call C_display_wenti.set_m_point_no(w_n%, tangent_p%, tn%, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(0), tn% + 1, True)
   Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(1), tn% + 2, True)
   Call C_display_wenti.set_m_point_no(w_n%, p%, tn% + 3, True)
 t_coord = minus_POINTAPI(time_POINTAPI_by_number( _
           m_poi(tangent_p%).data(0).data0.coordinate, 2), _
                                m_poi(p%).data(0).data0.coordinate)
       tn% = 0
       Call set_point_coordinate(tn%, t_coord, False)
       Call C_display_wenti.set_m_inner_poi(w_n%, tn%, 3)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
           draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond_2(ByVal p4%, ByVal p5%, ByVal c1%, _
         ByVal c2%, ByVal line_no%, ByVal tangent_line_no%)  'tl%,切线,=6 两园相切
'-2 作⊙□[down\\(_)]和⊙□[down\\(_)]的公切线□□
'-60 作⊙□□□和⊙□□□的公切线□□
'-59 作⊙□□□和⊙□[down\\(_)]的公切线□□
Dim i%, j%, l%, tp%, w_n%
Dim p(1) As Integer
Dim c_ty(1) As Byte
If p4% > 90 Then
  p4% = set_point_from_aid_point(p4%)
End If
If p5% > 90 Then
  p4% = set_point_from_aid_point(p4%)
End If
If tangent_line_no% <> 6 Then
Call add_point_to_m_circle(p4%, c1%, record0, True)
Call add_point_to_m_circle(p5%, c2%, record0, False)
End If
If m_Circ(c1%).data(0).data0.center = 0 Then
     c_ty(0) = 1
End If
If m_Circ(c2%).data(0).data0.center = 0 Then
     c_ty(1) = 1
End If
If c_ty(0) = 0 And c_ty(1) = 0 Then
 Call C_display_wenti.set_m_no(0, -2, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.center, 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, p4%, 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, p5%, 5, True)
ElseIf c_ty(0) = 0 And c_ty(1) = 1 Then
 Call C_display_wenti.set_m_no(0, -59, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
 Call C_display_wenti.set_m_point_no(w_n%, p5%, 6, True)
ElseIf c_ty(0) = 1 And c_ty(1) = 0 Then
Call C_display_wenti.set_m_no(0, -59, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.center, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
 Call C_display_wenti.set_m_point_no(w_n%, p5%, 6, True)
ElseIf c_ty(0) = 1 And c_ty(1) = 1 Then
 Call C_display_wenti.set_m_no(0, -60, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 5, True)
 Call C_display_wenti.set_m_point_no(w_n%, p4%, 6, True)
 Call C_display_wenti.set_m_point_no(w_n%, p5%, 7, True)
End If
Call C_display_wenti.set_m_inner_circ(w_n%, c1%, 1)
Call C_display_wenti.set_m_inner_circ(w_n%, c2%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 2)
Call C_display_wenti.set_m_inner_lin(w_n%, tangent_line_no%, 1)
Call C_display_wenti.set_m_inner_lin(w_n%, line_no%, 2)
'Call set_parent(circle_, c1%, wenti_cond_, w_n%, 0)
'Call set_parent(circle_, c2%, wenti_cond_, w_n%, 0)
Call C_display_wenti.set_m_inner_point_type(w_n%, tangent_line_no%)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
         draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond13(ByVal c1%, ByVal c2%, ByVal tp%, ByVal ty As Integer)
'13 □是⊙□[down\\(_)]和⊙□[down\\(_)]一个交点
'-66 □是⊙□□□和⊙□[down\\(_)]一个交点
'-67  □是⊙□□□和⊙□□□一个交点
Dim i%, j%, w_n%
If m_Circ(c1%).data(0).data0.center > 0 And m_Circ(c2%).data(0).data0.center > 0 Then
 Call C_display_wenti.set_m_no(0, 13, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, tp%, 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.center, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 4, True)
ElseIf m_Circ(c1%).data(0).data0.center = 0 And m_Circ(c2%).data(0).data0.center > 0 Then
    Call C_display_wenti.set_m_no(0, -66, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp%, 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(2), 2, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(3), 3, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.center, 4, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 5, True)
ElseIf m_Circ(c1%).data(0).data0.center > 0 And m_Circ(c2%).data(0).data0.center = 0 Then
    Call C_display_wenti.set_m_no(0, -66, w_n%)
    Call C_display_wenti.set_m_point_no(w_n%, tp%, 0, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 1, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 2, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 3, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.center, 4, True)
    Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 5, True)
Else
  If m_Circ(c1%).data(0).data0.center = 0 And m_Circ(c2%).data(0).data0.center = 0 Then
     Call C_display_wenti.set_m_no(0, -67, w_n%)
     Call C_display_wenti.set_m_point_no(w_n%, tp%, 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(1), 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(2), 2, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c1%).data(0).data0.in_point(3), 3, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(1), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(2), 5, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c2%).data(0).data0.in_point(3), 6, True)
 Else
     Exit Sub
 End If
End If
'If c1% > c2% Then
'   call exchange_two_integer(c1%, c2%)
'End If
  Call C_display_wenti.set_m_inner_circ(w_n%, c1%, 1)
  Call C_display_wenti.set_m_inner_circ(w_n%, c2%, 2)
  Call C_display_wenti.set_m_inner_poi(w_n%, tp%, 1)
  Call C_display_wenti.complete_set_inner_data(w_n%)
         Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
     'Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, c1%, m_Circ(c1%).data(0).sons)
     'Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, c2%, m_Circ(c2%).data(0).sons)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
          draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond12_64_65(ByVal c1%, ByVal c2%, ByVal tp%, ByVal ty As Integer)
'12 ⊙□[down\\(_)]和⊙□[down\\(_)]相切于点□
'-65 ⊙□□□和⊙□□□相切于点□
'-64 ⊙□□□和⊙□[down\\(_)]相切于点□
Dim p(1) As Integer
Dim l%, m%, i%, w_n%
Dim circ_type(1) As Integer
Dim c_data0 As condition_data_type
MDIForm1.Toolbar1.Buttons(21).Image = 33
t_coord.X = m_poi(tp%).data(0).data0.coordinate.X + _
     (m_Circ(c2%).data(0).data0.c_coord.Y - m_Circ(c1%).data(0).data0.c_coord.Y)
t_coord.Y = m_poi(tp%).data(0).data0.coordinate.Y + _
     (m_Circ(c1%).data(0).data0.c_coord.X - m_Circ(c2%).data(0).data0.c_coord.X)
     Call set_point_coordinate(p(0), t_coord, False)
t_coord.X = m_poi(tp%).data(0).data0.coordinate.X + _
     (m_Circ(c1%).data(0).data0.c_coord.Y - m_Circ(c2%).data(0).data0.c_coord.Y)
t_coord.Y = m_poi(tp%).data(0).data0.coordinate.Y + _
     (m_Circ(c2%).data(0).data0.c_coord.X - m_Circ(c1%).data(0).data0.c_coord.X)
     Call set_point_coordinate(p(1), t_coord, False)
m% = line_number(p(0), p(1), pointapi0, pointapi0, _
                 depend_condition(point_, p(0)), _
                 depend_condition(point_, p(1)), _
                 condition, condition_color, 1, 0)      '公切线
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(tp%, m%, 0, no_display, False, 0)
Call get_new_char(p(0))
Call get_new_char(p(1))
If m_Circ(c1%).data(0).data0.in_point(0) >= 3 And _
    m_Circ(c1%).data(0).data0.in_point(3) < m_Circ(c1%).data(0).data0.center Then
     circ_type(0) = 1
End If
If m_Circ(c2%).data(0).data0.in_point(0) >= 3 And _
    m_Circ(c2%).data(0).data0.in_point(3) < m_Circ(c2%).data(0).data0.center Then
     circ_type(1) = 1
End If
If circ_type(0) = 0 And circ_type(1) = 0 Then '两个有心圆
Call C_display_wenti.set_m_no(0, 12, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.center, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.center, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, tp%, 4, True)
ElseIf circ_type(0) = 1 And circ_type(1) = 0 Then
 Call C_display_wenti.set_m_no(0, -64, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.center, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, tp%, 5, True)
ElseIf circ_type(0) = 0 And circ_type(1) = 1 Then
 Call C_display_wenti.set_m_no(0, -64, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.center, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, tp%, 5, True)
Else
 Call C_display_wenti.set_m_no(0, -65, w_n%)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 0, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(2), 1, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(3), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(2), 4, True)
 Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(3), 5, True)
 Call C_display_wenti.set_m_point_no(w_n%, tp%, 6, True)
End If
Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
Call C_display_wenti.set_m_inner_circ(w_n%, c1%, 1)
Call C_display_wenti.set_m_inner_circ(w_n%, c2%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, tp%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, p(0), 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p(1), 3)
Call C_display_wenti.set_m_inner_lin(w_n%, m%, 1)
 If m_Circ(c1%).data(0).data0.center > 0 And _
      m_Circ(c2%).data(0).data0.center > 0 Then
   m% = line_number(m_Circ(c1%).data(0).data0.center, _
                    m_Circ(c2%).data(0).data0.center, _
                    pointapi0, pointapi0, _
                    depend_condition(point_, m_Circ(c1%).data(0).data0.center), _
                    depend_condition(point_, m_Circ(c2%).data(0).data0.center), _
                    condition, condition_color, 1, 0) '连心线
   Call add_point_to_line(tp%, m%, 0, no_display, True, 0)
 Else
   m% = 0
 End If
   Call C_display_wenti.set_m_inner_lin(w_n%, m%, 2)
' For i% = 0 To 4
' If m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree > 2 Then
'  m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree = _
'    m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree - 3
' End If
 Next i%
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond12_27(ByVal c1%, ByVal c2%, ByVal tp%, m_no%, ByVal ty As Integer)
'12 ⊙□[down\\(_)]和⊙□[down\\(_)]相切于点□
'-27与⊙□□□相切于□的切线交直线□□于□
Dim p(1) As Integer
Dim l%, m%, i%, w_n%
If m_Circ(c1%).data(0).data0.center > 0 And m_Circ(c2%).data(0).data0.center > 0 Then
l% = line_number0(m_Circ(c1%).data(0).data0.center, m_Circ(c2%).data(0).data0.center, 0, 0)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(tp%, l%, 0, no_display, True, 0)
End If
MDIForm1.Toolbar1.Buttons(21).Image = 33
t_coord.X = m_poi(tp%).data(0).data0.coordinate.X + _
     (m_Circ(c2%).data(0).data0.c_coord.Y - m_Circ(c1%).data(0).data0.c_coord.Y)
t_coord.Y = m_poi(tp%).data(0).data0.coordinate.Y + _
     (m_Circ(c1%).data(0).data0.c_coord.X - m_Circ(c2%).data(0).data0.c_coord.X)
     Call set_point_coordinate(p(0), t_coord, False)
t_coord.X = m_poi(tp%).data(0).data0.coordinate.X + _
     (m_Circ(c1%).data(0).data0.c_coord.Y - m_Circ(c2%).data(0).data0.c_coord.Y)
t_coord.Y = m_poi(tp%).data(0).data0.coordinate.Y + _
     (m_Circ(c2%).data(0).data0.c_coord.X - m_Circ(c1%).data(0).data0.c_coord.X)
     Call set_point_coordinate(p(1), t_coord, False)
m% = line_number0(p(0), p(1), 0, 0)  '公切线
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(tp%, m%, 0, no_display, True, 0)
'lin(m%).data(0).data0.visible = 10
Call get_new_char(p(0))
Call get_new_char(p(1))
Call vertical_line(m%, l%, True, True)
'Call set_tangent_circle(c1%, c2%, tp%, m%, 0, temp_record)
Call C_display_wenti.set_m_no(0, m_no%, w_n%)
If m_no% = 12 Then
'Call C_display_wenti.set_m_no(0,12, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.center, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.center, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, tp%, 4, True)
Else
'Call C_display_wenti.set_m_no(0,m_no%)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(1), 0, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(2), 1, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c1%).data(0).data0.in_point(3), 2, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.center, 3, True)
Call C_display_wenti.set_m_point_no(w_n%, _
     m_Circ(c2%).data(0).data0.in_point(1), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, tp%, 5, True)
End If
Call C_display_wenti.set_m_point_no(w_n%, c1%, 12, False)
Call C_display_wenti.set_m_point_no(w_n%, c2%, 13, False)
 For i% = 0 To 4
 'If m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree > 2 Then
 ' m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree = _
 '   m_poi(C_display_wenti.m_point_no(w_n%, i%)).data(0).degree - 3
 'End If
 Next i%
Call C_display_wenti.set_m_point_no(w_n%, ty, 9, False)
Call C_display_wenti.set_m_point_no(w_n%, m%, 10, False)
Call C_display_wenti.set_m_point_no(w_n%, p(0), 15, True)
Call C_display_wenti.set_m_point_no(w_n%, p(1), 16, True)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond_6_42_43_57(ByVal p1%, ByVal p2%, _
       ele1 As condition_type, ele2 As condition_type, ty%)
'-6□□=!_~
'-42 在⊙□[down\\(_)]上取一点□使得□□＝!_~
'-43 在□□上取一点□使得□□＝_"
'-57 在⊙□□□上取一点□使得□□＝!_~
Dim tp(2) As Integer
Dim i%, j%, w_n%, tp_%
Dim val As String
Dim tc%, tl%
Dim value_line_no%
Dim value_circle_no%
Dim temp_record As total_record_type
Dim n%
 temp_record.record_data.data0.condition_data.condition_no = 0
  temp_record.record_.display_no = -(C_display_wenti.m_last_input_wenti_no)
If m_poi(p2%).data(0).parent.inter_type = exist_point Or m_poi(p2%).data(0).parent.inter_type = new_free_point Then
 n% = -6
ElseIf m_poi(p2%).data(0).parent.inter_type = new_point_on_line Then
 n% = -43
ElseIf m_poi(p2%).data(0).parent.inter_type = new_point_on_circle Then
 If m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.center = 0 Then
  n% = -57
 Else
  n% = 42
 End If
End If
'*************************************************************************************
If n% = -43 Then
    tp(0) = m_lin(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.poi(0)
     tp(1) = m_lin(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.poi(1)
ElseIf n% = -42 Then
       tp(0) = m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.center
       tp(1) = m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.in_point(1)
ElseIf n% = -57 Then
         tp(0) = m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.in_point(1)
         tp(1) = m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.in_point(2)
         tp(2) = m_Circ(m_poi(p2%).data(0).parent.element(1).no).data(0).data0.in_point(3)
End If
'**************************************************************************
Call C_display_wenti.set_m_no(0, n%, w_n%)
value_line_no% = line_number(p1%, p2%, pointapi0, pointapi0, depend_condition(point_, p1%), depend_condition(point_, p2%), condition, _
            condition_color, 1, 0)
value_circle_no% = m_circle_number(1, p1%, m_poi(p1%).data(0).data0.coordinate, 0, 0, 0, _
                   distance_of_two_POINTAPI(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate), _
                    0, 0, 0, 0, 0, 0, False)
'***************************************************************************************************************************
If n% = -6 Then
   Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
   Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)     'temp_point(1)
   Call C_display_wenti.set_m_inner_lin(w_n%, value_line_no%, 1)
  ' value_circle_no% = m_circle_number(1, p1%, m_poi(p1%).data(0).data0.coordinate, 0, 0, 0, _
                   distance_of_two_POINTAPI(m_poi(p0%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate), _
                    0, 0, 0, 0, 0, 0, False)
   Call C_display_wenti.set_m_inner_circ(w_n%, value_circle_no%, 1)
 Call add_point_to_m_circle(p2%, value_circle_no%, record0, True)
 'If value$ <> "" Then '未输入长度值
 'Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1) '
 'Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 1) '
 '  Call C_display_wenti.set_m_point_no(w_n%, _
        m_poi(p1%).data(0).data0.coordinate.X - m_poi(p0%).data(0).data0.coordinate.X, 17, False) '原线长的横坐标
 '  Call C_display_wenti.set_m_point_no(w_n%, _
        m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p0%).data(0).data0.coordinate.Y, 18, False) '
 'End If
ElseIf n% = -43 Or n% = -42 Then '-43 在□□上取一点□使得□□＝_"'-42 在⊙□[down\\(_)]上取一点□使得□□＝!_~
  Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)    'temp_point(0)
  Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True) 'temp_point(1)
   Call C_display_wenti.set_m_point_no(w_n%, p1%, 2, True)
   Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True)     'temp_point(1)
  Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(p2%).data(0).parent.element(1).no, 1) '选点的直线
     tc% = m_circle_number(0, p1%, m_poi(p1%).data(0).data0.coordinate, _
                0, 0, 0, distance_of_two_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                 m_poi(p2%).data(0).data0.coordinate), 0, 0, 0, 0, 0, 0, False)
     If n% = -43 Then
       m_poi(p2%).data(0).parent.inter_type = new_inter_point_type(m_poi(p2%).data(0).parent.inter_type, _
                inter_point_line_circle(m_poi(p2%).data(0).parent.element(1).no, tc%, m_poi(p2%).data(0).data0.coordinate, p2%, False, True))
     ElseIf n% = -42 Then
        Call C_display_wenti.set_m_inner_circ(w_n%, m_poi(p2%).data(0).parent.element(1).no, 2)
          'm_poi(p2%).data(0).parent.inter_type = new_inter_point_type(, m_poi(p2%).data(0).parent.inter_type, _
                inter_point_circle_circle(m_poi(p2%).data(0).parent.element(1).no, _
                    tc%, m_poi(p2%).data(0).data0.coordinate))
     End If
  Call C_display_wenti.set_m_inner_circ(w_n%, tc%, 1) '定长线段确定的圆
  Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 1)
'*****************************************************************
ElseIf n% = -57 Then '-57 在⊙□□□上取一点□使得□□＝!_~
 Call C_display_wenti.set_m_point_no(w_n%, tp(0), 0, True)    'temp_point(0)
 Call C_display_wenti.set_m_point_no(w_n%, tp(1), 1, True) 'temp_point(1)
 Call C_display_wenti.set_m_point_no(w_n%, tp(2), 2, True)
 Call C_display_wenti.set_m_point_no(w_n%, p2%, 3, True)
 Call C_display_wenti.set_m_point_no(w_n%, p1%, 4, True)     'temp_point(1)
'*****************************************************************************
End If
If val <> "" Then
   For i% = 1 To Len(val)
    Call C_display_wenti.set_m_condition(w_n%, _
     Mid$(val, i%, 1), 4 + i%)
   Next i%
    Call C_display_wenti.set_m_condition(w_n%, Chr(13), 4 + i%)
 End If

'Call C_display_wenti.set_m_point_no(w_n%, value * 100, 11, False)      '线长
 'Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), _
                  depend_condition(point_, p2%), _
                  condition, condition_color, 1, 0)
record0.data0.condition_data.condition_no = 0
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
'****
End Sub

Public Sub set_wenti_cond_50_51_52_56(ByVal p1%, ByVal p2%, ByVal p3%, _
          ByVal p4%, w_n%)
'-50 □□是∠□□□的平分线
Dim l%, change_p%
Dim ty As Integer
Dim r(1) As Integer
Dim tl(2) As Integer
Dim A!
Dim tp(4) As Integer
Dim t_coord(1) As POINTAPI
Dim wenti_ty As Byte
Dim wenti_no As Integer
Dim temp_record As total_record_type
Dim triA(1) As Integer
tp(0) = p1%
tp(1) = p2%
tp(2) = p3%
tp(3) = p4%
'输入四个点
If m_poi(p4%).data(0).parent.inter_type = interset_point_line_line Then
   wenti_no = -51
ElseIf m_poi(p4%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
        m_poi(p4%).data(0).parent.inter_type = new_point_on_line_circle21 Then
        If m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.center > 0 Then
           wenti_no = -52
        Else
           wenti_no = -56
        End If
Else
   wenti_no = -50
End If
 Call C_display_wenti.set_m_no(0, wenti_no, w_n%)
     Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
     Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
     Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
  If wenti_no = -50 Then
     Call C_display_wenti.set_m_point_no(w_n%, p4%, 3, True)
  ElseIf wenti_no = -51 Then
     Call C_display_wenti.set_m_point_no(w_n%, m_lin(m_poi(p4%).data(0).parent.element(1).no).data(0).data0.poi(0), 3, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_lin(m_poi(p4%).data(0).parent.element(1).no).data(0).data0.poi(1), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
  ElseIf wenti_no = -52 Then
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.center, 3, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.in_point(1), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
  ElseIf wenti_no = -56 Then
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.in_point(1), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.in_point(2), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, m_Circ(m_poi(p4%).data(0).parent.element(2).no).data(0).data0.in_point(3), 4, True)
     Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
  End If
l% = abs_POINTAPI(minus_POINTAPI(m_poi(tp(1)).data(0).data0.coordinate, _
            m_poi(tp(3)).data(0).data0.coordinate))
Call C_display_wenti.set_m_point_no(w_n%, l%, 40, False)
tl(0) = line_number(tp(1), tp(3), pointapi0, pointapi0, _
                    depend_condition(point_, tp(1)), _
                    depend_condition(point_, tp(3)), _
                    condition, condition_color, 1, 0)
tl(1) = line_number(tp(1), tp(0), pointapi0, pointapi0, _
                    depend_condition(point_, tp(1)), _
                    depend_condition(point_, tp(0)), _
                    condition, condition_color, 1, 0)
tl(2) = line_number(tp(1), tp(2), pointapi0, pointapi0, _
                    depend_condition(point_, tp(1)), _
                    depend_condition(point_, tp(2)), _
                    condition, condition_color, 1, 0)
Call set_parent(line_, tl(1), wenti_cond_, w_n%, 0)
Call set_parent(line_, tl(2), wenti_cond_, w_n%, 0)
Call C_display_wenti.set_m_inner_lin(w_n%, tl(0), 2)
Call C_display_wenti.set_m_inner_lin(w_n%, tl(1), 3)
Call C_display_wenti.set_m_inner_lin(w_n%, tl(2), 4)
Call C_display_wenti.set_m_inner_poi(w_n%, tp(0), 1)
Call C_display_wenti.set_m_inner_poi(w_n%, tp(1), 2)
Call C_display_wenti.set_m_inner_poi(w_n%, tp(2), 3)
Call C_display_wenti.set_m_inner_poi(w_n%, tp(3), 4)
 temp_record.record_data.data0.condition_data.condition_no = 0
 temp_record.record_.display_no = -w_n%
    triA(0) = Abs(angle_number(p1%, p2%, p4%, 0, 0))
    triA(1) = Abs(angle_number(p3%, p2%, p4%, 0, 0))
If triA(0) > 0 And triA(1) > 0 Then
       Call set_three_angle_value(triA(0), triA(1), 0, "1", "-1", "0", "0", _
              0, temp_record, 0, 0, 0, 0, 0, 0, True)
 'Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If

'***********************************************************************************
'*************************************************************************************
'If read_free_point_in_line(tl(0), tp(1), tp(3), tp(4)) Then
'   tp(3) = tp(4)
'End If
'If read_free_point_in_line(tl(1), tp(1), tp(0), tp(4)) Then
'   tp(0) = tp(4)
'End If
'If read_free_point_in_line(tl(2), tp(1), tp(2), tp(4)) Then
'   tp(2) = tp(4)
'End If

'Call C_display_wenti.set_m_inner_poi(w_n%, tp(3), 1)
'Call C_display_wenti.set_m_inner_poi(w_n%, tp(0), 2)
'Call C_display_wenti.set_m_inner_poi(w_n%, tp(1), 3)
'Call C_display_wenti.set_m_inner_poi(w_n%, tp(2), 4)
'Call C_display_wenti.set_m_no_(w_n%, -50)
'If wenti_type = 1 Then
'operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
'     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
'Else
'   If m_poi(tp(3)).data(0).parent.co_degree = 1 Then
'       If m_poi(tp(3)).data(0).parent.element(1).ty = line_ Then
'         If m_poi(tp(3)).data(0).parent.element(1).no = tl(0) Then
 '        Else
'          Call C_display_wenti.set_m_no_(w_n%, -51)
'          Call C_display_wenti.set_m_inner_lin( _
                w_n%, m_poi(tp(3)).data(0).parent.element(1).no, 1)
'          Call inter_point_line_sp_angle_with_line(tp(0), _
            tp(1), tp(2), m_lin(m_poi(tp(3)).data(0).parent.element(1).no).data(0).data0.poi(0), _
            m_lin(m_poi(tp(3)).data(0).parent.element(1).no).data(0).data0.poi(1), _
             t_coord(0), True, 0)
'          Call set_point_coordinate(tp(3), t_coord(0), True)
'          change_p% = tp(3)
'         End If
'       ElseIf m_poi(tp(3)).data(0).parent.element(1).ty = circle_ Then
'          Call C_display_wenti.set_m_no_(w_n%, -52)
'                    Call C_display_wenti.set_m_no_(w_n%, -52)
'          Call C_display_wenti.set_m_inner_circ( _
                w_n%, m_poi(tp(3)).data(0).parent.element(1).no, 1)
'          Call inter_point_line_sp_angle_with_circle( _
            tp(0), tp(1), tp(2), _
                   m_poi(tp(3)).data(0).parent.element(1).no, _
                    t_coord(0), 0, t_coord(1), 0)
'          If distance_of_two_POINTAPI(t_coord(0), m_poi(tp(3)).data(0).data0.coordinate) < _
              distance_of_two_POINTAPI(t_coord(1), m_poi(tp(3)).data(0).data0.coordinate) Then
'              Call C_display_wenti.set_m_inner_point_type(w_n%, 1)
'              Call set_point_coordinate(tp(3), t_coord(0), True)
'          Else
'              Call C_display_wenti.set_m_inner_point_type(w_n%, 2)
'              Call set_point_coordinate(tp(3), t_coord(1), True)
'          End If
'              change_p% = tp(3)
'       End If
'    ElseIf m_poi(tp(3)).data(0).parent.co_degree = 0 Then
'      r(0) = distance_of_two_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, _
'                              m_poi(tp(1)).data(0).data0.coordinate)
'      r(1) = distance_of_two_POINTAPI(m_poi(tp(2)).data(0).data0.coordinate, _
'                              m_poi(tp(1)).data(0).data0.coordinate)
'      A! = r(0) / r(1)
'      t_coord(0) = add_POINTAPI(m_poi(tp(1)).data(0).data0.coordinate, _
'                     time_POINTAPI_by_number(minus_POINTAPI(m_poi(tp(2)).data(0).data0.coordinate, _
'                      m_poi(tp(1)).data(0).data0.coordinate), A!))
'      t_coord(0) = mid_POINTAPI(t_coord(0), m_poi(tp(0)).data(0).data0.coordinate)
'      r(0) = distance_of_two_POINTAPI(m_poi(tp(3)).data(0).data0.coordinate, _
                              m_poi(tp(1)).data(0).data0.coordinate)
'      r(1) = distance_of_two_POINTAPI(t_coord(0), _
'                              m_poi(tp(1)).data(0).data0.coordinate)
'      A! = r(0) / r(1)
'      t_coord(1) = add_POINTAPI(m_poi(tp(1)).data(0).data0.coordinate, _
'                     time_POINTAPI_by_number(minus_POINTAPI(t_coord(0), _
                      m_poi(tp(1)).data(0).data0.coordinate), A!))
'      Call set_point_coordinate(tp(3), t_coord(1), True)
'      change_p% = tp(3)
   'End If
' Else 'm_poi(tp(3)).data(0).parent.co_degree=2
'   If m_lin(tl(0)).data(0).parent.co_degree <= 2 And _
'           read_free_point_in_line(tl(0), tp(1), tp(3), tp(4)) Then
'      Call C_display_wenti.set_m_inner_poi(w_n%, tp(4), 1)
'      Call set_wenti_cond_50(tp(0), tp(1), tp(2), tp(4), w_n%)
'      change_p% = tp(4)
'   Else
'   If m_poi(tp(0)).data(0).parent.co_degree = 0 Then
'      ty = 0
'      Call C_display_wenti.set_m_no_(w_n%, -501)
'      Call C_display_wenti.set_m_inner_lin(w_n%, 0, 1)
'      Call draw_equal_angle(tp(3), tp(1), tp(2), t_coord(0), t_coord(1), tp(1), tp(3), 0, 0, ty)
'      Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
'      Call set_point_coordinate(tp(0), t_coord(0), True)
'      change_p% = tp(0)
'   ElseIf m_poi(tp(2)).data(0).parent.co_degree = 0 Then
'      Call C_display_wenti.set_m_no_(w_n%, -502)
'      ty = 0
'      Call draw_equal_angle(tp(3), tp(1), tp(0), t_coord(0), t_coord(1), tp(1), tp(3), 0, 0, ty)
'      Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
'      Call set_point_coordinate(tp(2), t_coord(0), True)
'      change_p% = tp(2)
'   ElseIf m_poi(tp(0)).data(0).parent.co_degree = 1 Then
'      If m_poi(tp(0)).data(0).parent.element(1).ty = line_ Then
'         If m_poi(tp(0)).data(0).parent.element(1).no <> tl(1) Then
'          Call C_display_wenti.set_m_no_(w_n%, -501)
'          Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(tp(0)).data(0).parent.element(1).no, 1)
'          Call C_display_wenti.set_m_inner_circ(w_n%, 0, 1)
'          Call draw_equal_angle(tp(3), tp(1), tp(2), t_coord(0), t_coord(1), tp(1), tp(3), _
'               m_poi(tp(0)).data(0).parent.element(1).no, 0, ty)
'          Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 2)
'          Call set_point_coordinate(tp(0), t_coord(0), True)
'          change_p% = tp(0)
'         End If
'      ElseIf m_poi(tp(0)).data(0).parent.element(1).ty = circle_ Then
'          Call C_display_wenti.set_m_no_(w_n%, -501)
'          Call C_display_wenti.set_m_inner_lin(w_n%, 0, 1)
'          Call C_display_wenti.set_m_inner_circ(w_n%, m_poi(tp(0)).data(0).parent.element(1).no, 1)
'          Call draw_equal_angle(tp(3), tp(1), tp(2), t_coord(0), t_coord(1), tp(1), tp(3), _
'               0, m_poi(tp(0)).data(0).parent.element(1).no, ty)
'          If distance_of_two_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, t_coord(0)) < _
'               distance_of_two_POINTAPI(m_poi(tp(0)).data(0).data0.coordinate, t_coord(1)) Then
'           Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 2)
'           Call set_point_coordinate(tp(0), t_coord(0), True)
'          Else
'           Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 4)
'           Call set_point_coordinate(tp(0), t_coord(0), True)
'          End If
'          change_p% = tp(0)
'      End If
'   ElseIf m_poi(tp(2)).data(0).parent.co_degree = 1 Then
'      If m_poi(tp(2)).data(0).parent.element(1).ty = line_ Then
'         If m_poi(tp(2)).data(0).parent.element(1).no <> tl(1) Then
'          Call C_display_wenti.set_m_no_(w_n%, -502)
'          Call C_display_wenti.set_m_inner_lin(w_n%, m_poi(tp(0)).data(0).parent.element(1).no, 1)
'          Call C_display_wenti.set_m_inner_circ(w_n%, 0, 1)
'          Call draw_equal_angle(tp(3), tp(1), tp(0), t_coord(0), t_coord(1), tp(1), tp(3), _
'                m_poi(tp(2)).data(0).parent.element(1).no, 0, ty)
'          Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 2)
'          Call set_point_coordinate(tp(2), t_coord(0), True)
'          change_p% = tp(2)
'         End If
'      ElseIf m_poi(tp(2)).data(0).parent.element(1).ty = circle_ Then
'       Call C_display_wenti.set_m_no_(w_n%, -502)
'       Call C_display_wenti.set_m_inner_lin(w_n%, 0, 1)
'       Call C_display_wenti.set_m_inner_circ(w_n%, m_poi(tp(0)).data(0).parent.element(1).no, 1)
'       Call draw_equal_angle(tp(3), tp(1), tp(0), t_coord(0), t_coord(1), tp(1), tp(3), 0, _
'             m_poi(tp(2)).data(0).parent.element(1).no, ty)
'          If distance_of_two_POINTAPI(m_poi(tp(2)).data(0).data0.coordinate, t_coord(0)) < _
'               distance_of_two_POINTAPI(m_poi(tp(2)).data(0).data0.coordinate, t_coord(1)) Then
'           Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 2)
'           Call set_point_coordinate(tp(2), t_coord(0), True)
'          Else
'           Call C_display_wenti.set_m_inner_point_type(w_n%, ty + 4)
'           Call set_point_coordinate(tp(2), t_coord(0), True)
'          End If
'       change_p% = tp(0)
'      End If
'   End If
'   End If
' End If
    Call C_display_wenti.complete_set_inner_data(w_n%)

   'Call change_picture_(change_p%, 0)
'   Call draw_again0(Draw_form, 1)
'End If
End Sub
Public Sub set_wenti_cond_51(ByVal p1%, ByVal p2%, ByVal p3%, _
          ByVal p4%, ByVal l%)
'-51 ∠□□□的平分线交□□于□
Dim w_n%
Call C_display_wenti.set_m_no(0, -51, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
If m_lin(l%).data(0).data0.poi(0) = p4% Then
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.in_point(2), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(1), 4, True)
ElseIf m_lin(l%).data(0).data0.poi(1) = p4% Then
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(0), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.in_point(m_lin(l%).data(0).data0.in_point(0) - 1), 4, True)
Else
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(0), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_lin(l%).data(0).data0.poi(1), 4, True)
End If
Call C_display_wenti.set_m_point_no(w_n%, p4%, 5, True)
Call C_display_wenti.set_m_inner_poi(w_n%, p4%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 4)
Call C_display_wenti.set_m_inner_lin(w_n%, l%, 1) '相交线
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p2%, p4%, 0, 0), 2)  '角平分线
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p2%, p1%, 0, 0), 3)  '角的边
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p2%, p3%, 0, 0), 4)  '角的边
C_display_wenti.complete_set_inner_data (w_n%)
'Call set_son_data(wenti_cond_, w_n%, son_data0, line_, _
              line_number0(p2%, p1%, 0, 0), m_lin(line_number0(p2%, p1%, 0, 0)).data(0).sons)
'Call set_son_data(wenti_cond_, w_n%, son_data0, line_, _
              line_number0(p2%, p3%, 0, 0), m_lin(line_number0(p2%, p3%, 0, 0)).data(0).sons)
'Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p2%, m_poi(p2%).data(0).sons)
'Call set_son_data(wenti_cond_, w_n%, son_data0, line_, l%, m_lin(l%).data(0).sons)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_wenti_cond_52(ByVal p1%, ByVal p2%, ByVal p3%, _
       ByVal p6%, ByVal c%, ty As Integer)
'-52∠□□□的平分线交⊙□[down\\(_)]于□
'-56∠□□□的平分线交⊙□□□于□
Dim w_n%
Dim l1%, l2%
If m_Circ(c%).data(0).data0.center > 0 Then
Call C_display_wenti.set_m_no(0, -52, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, p6%, 5, True)
Else
Call C_display_wenti.set_m_no(0, -56, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 5, True)
Call C_display_wenti.set_m_point_no(w_n%, p6%, 6, True)
End If
l1% = line_number0(p2%, p1%, 0, 0)
l2% = line_number0(p2%, p3%, 0, 0)
Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p2%, p6%, 0, 0), 2)  '角平分线
Call C_display_wenti.set_m_inner_lin(w_n%, l1%, 3)  '角的边
Call C_display_wenti.set_m_inner_lin(w_n%, l2%, 4)  '角的边
Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, p6%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 4)
Call set_son_data(wenti_cond_, son_data0, line_, l1%, w_n%, m_lin(l1%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, data0, line_, l2%, m_lin(l2%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, son_data0, circle_, c%, m_Circ(c%).data(0).sons)
record0.data0.condition_data.condition_no = 0
    Call C_display_wenti.complete_set_inner_data(w_n%)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond_41(ByVal p1%, ByVal p2%, ByVal p3%, _
         ByVal p4%, ByVal p5%, p6%, value$, w_n%)
If value$ = "1" Then
   Call set_wenti_cond_4(p1%, p2%, p3%, p4%, p5%, p6%, w_n%)
End If
End Sub
Public Sub set_wenti_cond_4(ByVal p1%, ByVal p2%, ByVal p3%, _
         ByVal p4%, ByVal p5%, p6%, w_n%)
'-4∠□□□=∠□□□
Dim midpoint%
Dim tl(3) As Integer
If w_n% = 0 Then
Call C_display_wenti.set_m_no(0, -4, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, p4%, 3, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 4, True)
Call C_display_wenti.set_m_point_no(w_n%, p6%, 5, True)
Call C_display_wenti.set_m_point_no(w_n%, midpoint%, 8, True)
End If
tl(0) = line_number(p1%, p2%, pointapi0, pointapi0, _
                    depend_condition(point_, p1%), _
                    depend_condition(point_, p2%), _
                    condition, condition_color, 1, 0)
tl(1) = line_number(p3%, p2%, pointapi0, pointapi0, _
                    depend_condition(point_, p3%), _
                    depend_condition(point_, p2%), _
                    condition, condition_color, 1, 0)
tl(2) = line_number(p4%, p5%, pointapi0, pointapi0, _
                    depend_condition(point_, p4%), _
                    depend_condition(point_, p5%), _
                    condition, condition_color, 1, 0)
tl(3) = line_number(p6%, p5%, pointapi0, pointapi0, _
                    depend_condition(point_, p6%), _
                    depend_condition(point_, p5%), _
                    condition, condition_color, 1, 0)
If p2% = p5% Then
   If tl(0) = tl(2) Then
      Call C_display_wenti.set_m_no_(w_n%, -50)
      Call set_wenti_cond_50(p3%, p2%, p6%, p1%, w_n%)
   ElseIf tl(0) = tl(3) Then
      Call C_display_wenti.set_m_no_(w_n%, -50)
      Call set_wenti_cond_50(p3%, p2%, p4%, p1%, w_n%)
   ElseIf tl(1) = tl(2) Then
      Call C_display_wenti.set_m_no_(w_n%, -50)
      Call set_wenti_cond_50(p1%, p2%, p6%, p3%, w_n%)
   ElseIf tl(1) = tl(3) Then
      Call C_display_wenti.set_m_no_(w_n%, -50)
      Call set_wenti_cond_50(p1%, p2%, p4%, p3%, w_n%)
   End If
 End If
End Sub
Public Sub set_wenti_cond4(ByVal p1%, ByVal p2%, ByVal p3%, _
         ByVal midpoint%, ByVal l%, w_n%)
'4 在□□的垂直平分线上任取一点□
Dim A!
Dim tn%, tl%, tc%
Dim mp As mid_point_data0_type
Dim c_data As condition_data_type
Dim p_coord(1) As POINTAPI
Dim t_coord As POINTAPI
 'l% = line_number(p3%, midpoint%, condition, True, 0)
If w_n% = 0 Then
Call C_display_wenti.set_m_no(0, 4, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, midpoint%, 8, True)
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p3%, midpoint%, 0, 0), 2)
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p1%, p2%, 0, 0), 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, midpoint%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 4)
Call is_point_in_line1(m_poi(p3%).data(0).data0.coordinate, _
        p1%, p2%, midpoint%, True, t_coord, 0, A!, False)
'Call C_display_wenti.set_m_point_no(w_n%, Int(1000 * A!), 11, False) '长比
'Call set_son_data(wenti_cond_, w_n%, m_poi(p1%).data(0).sons)
'Call set_son_data(wenti_cond_, w_n%, m_poi(p2%).data(0).sons)
'Call set_son_data(wenti_cond_, w_n%, m_poi(p3%).data(0).sons)
t_coord = verti_POINTAPI(minus_POINTAPI(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate)) '4-3
If Abs(t_coord.X) > 5 Then
 m_poi(p3%).data(0).parent.ratio = (m_poi(p3%).data(0).data0.coordinate.X - _
                             m_poi(midpoint%).data(0).data0.coordinate.X) / t_coord.X '1-2/4-3
 
Else
 m_poi(p3%).data(0).parent.ratio = (m_poi(p3%).data(0).data0.coordinate.Y - _
                             m_poi(midpoint%).data(0).data0.coordinate.Y) / t_coord.Y
End If
record0.data0.condition_data.condition_no = 0
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
Else
 If is_mid_point(p1%, midpoint%, p2%, 0, 0, 0, 0, tn%, -1000, 0, 0, 0, 0, 0, 0, _
                    mp, "", 0, 0, 0, c_data) Then
 Else
    p_coord(0) = mid_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                          m_poi(p2%).data(0).data0.coordinate)
    midpoint% = set_point(p_coord(0), 0, 0, "", 0)
 End If
    Call add_point_to_line(midpoint%, line_number(p1%, p2%, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, p1%), _
                           depend_condition(point_, p2%), _
                           condition, condition_color, 1, 0), 0, False, False, 0)
    Call C_display_wenti.set_m_inner_poi(w_n%, p3%, 1)
    Call C_display_wenti.set_m_inner_poi(w_n%, midpoint%, 2)
    Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 3)
    Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 4)
   If m_poi(p3%).data(0).parent.co_degree = 1 Then
     If m_poi(p3%).data(0).parent.element(0).ty = line_ Then
        tl% = m_poi(p3%).data(0).parent.element(0).no
      Call C_display_wenti.set_m_no_(w_n%, -54)
          Call C_display_wenti.set_m_inner_lin(w_n%, tl%, 1)
         Call inter_point_line_line2(m_poi(midpoint%).data(0).data0.coordinate, _
             verti_, m_poi(p1%).data(0).data0.coordinate, _
               m_poi(p2%).data(0).data0.coordinate, _
                m_poi(m_lin(tl%).data(0).data0.poi(0)). _
                 data(0).data0.coordinate, _
                  paral_, m_poi(m_lin(tl%).data(0).data0.poi(0)). _
                   data(0).data0.coordinate, _
                    m_poi(m_lin(tl%).data(0).data0.poi(1)). _
                     data(0).data0.coordinate, p_coord(0), p3%, True, True)
     ElseIf m_poi(p3%).data(0).parent.element(0).ty = circle_ Then
         tc% = m_poi(p3%).data(0).parent.element(0).no
      Call C_display_wenti.set_m_no_(w_n%, -53)
          Call C_display_wenti.set_m_inner_circ(w_n%, tc%, 1)
          Call inter_point_line_circle3(m_poi(midpoint%).data(0).data0.coordinate, _
              False, m_poi(p1%).data(0).data0.coordinate, _
               m_poi(p2%).data(0).data0.coordinate, _
                 m_Circ(tc%).data(0).data0, _
                  p_coord(0), 0, p_coord(1), 0, 0, True)
          If distance_of_two_POINTAPI(p_coord(0), m_poi(p3%).data(0).data0.coordinate) < _
              distance_of_two_POINTAPI(p_coord(1), m_poi(p3%).data(0).data0.coordinate) Then
               Call set_point_coordinate(p3%, p_coord(0), True)
               Call C_display_wenti.set_m_inner_point_type(w_n%, 1)
          Else
               Call set_point_coordinate(p3%, p_coord(1), True)
               Call C_display_wenti.set_m_inner_point_type(w_n%, 2)
          End If
     End If
    ElseIf m_poi(p3%).data(0).parent.co_degree = 0 Then
      Call is_point_in_line1(m_poi(p3%).data(0).data0.coordinate, _
        p1%, p2%, midpoint%, True, t_coord, 0, A!, False)
         Call C_display_wenti.set_m_point_no(w_n%, Int(1000 * A!), 11, False) '长比
          p_coord(0) = add_POINTAPI(m_poi(midpoint%).data(0).data0.coordinate, _
              verti_POINTAPI(time_POINTAPI_by_number(minus_POINTAPI( _
               m_poi(p2%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate), _
                A!)))
                        Call set_point_coordinate(p3%, p_coord(0), True)
    End If
     Call draw_again0(Draw_form, 1)
End If
End Sub

Public Sub set_wenti_cond_54(ByVal p1%, ByVal p2%, ByVal midpoint%, ByVal p5%, ByVal l1%, ByVal l2%) 'l1% 标准,l2%垂线,
'-54 □□的垂直平分线交□□于□
Dim w_n%, p3%, p4%
If is_point_in_points(midpoint%, m_lin(l1%).data(0).data0.in_point) > 0 Then '是垂直平分线
    Call exchange_two_integer(l1%, l2%) '设定l2%是垂直平分线
End If
If m_lin(l1%).data(0).data0.poi(1) > 0 Then
 p3% = m_lin(l1%).data(0).data0.poi(0)
 p4% = m_lin(l1%).data(0).data0.poi(1)
Else
  p3% = m_lin(l1%).data(0).data0.in_point(1)
 p4% = m_lin(l1%).data(0).data0.in_point(m_lin(l1%).data(0).data0.in_point(0))
End If
Call C_display_wenti.set_m_no(0, -54, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, p3%, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, p4%, 3, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 4, True)

Call C_display_wenti.set_m_inner_lin(w_n%, l1%, 1)
Call C_display_wenti.set_m_inner_lin(w_n%, l2%, 2)
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p1%, p2%, 0, 0), 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, midpoint%, 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p1%, 3)
Call C_display_wenti.set_m_inner_poi(w_n%, p2%, 4)
'*********************************************************
Call set_son_data(wenti_cond_, w_n%, m_poi(p1%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, m_poi(p2%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, m_lin(l1%).data(0).sons)
record0.data0.condition_data.condition_no = 0
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub

Public Sub set_wenti_cond_53(ByVal p1%, ByVal p2%, ByVal p5%, ByVal midpoint%, ByVal c%, ty As Integer)
'-53 □□的垂直平分线交⊙□(_)□
'-25 □□的垂直平分线交⊙□□□于□
Dim w_n%
If m_Circ(c%).data(0).data0.center > 0 Then
Call C_display_wenti.set_m_no(0, -53, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.center, 2, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 4, True)
Else
Call C_display_wenti.set_m_no(0, -25, w_n%)
Call C_display_wenti.set_m_point_no(w_n%, p1%, 0, True)
Call C_display_wenti.set_m_point_no(w_n%, p2%, 1, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(1), 2, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(2), 3, True)
Call C_display_wenti.set_m_point_no(w_n%, m_Circ(c%).data(0).data0.in_point(3), 4, True)
Call C_display_wenti.set_m_point_no(w_n%, p5%, 5, True)
End If
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p1, p2%, 0, 0), 1)
Call C_display_wenti.set_m_inner_lin(w_n%, line_number0(p5%, midpoint%, 0, 0), 2)
Call C_display_wenti.set_m_inner_poi(w_n%, p5%, 1)
Call C_display_wenti.set_m_inner_circ(w_n%, c%, 1)
Call C_display_wenti.set_m_inner_poi(w_n%, midpoint%, 2)
Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p1%, m_poi(p1%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, son_data0, point_, p2%, m_poi(p2%).data(0).sons)
Call set_son_data(wenti_cond_, w_n%, son_data, circle_, c%, m_Circ(c%).data(0).sons)
record0.data0.condition_data.condition_no = 0
    Call C_display_wenti.set_m_inner_point_type(w_n%, ty)
     draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End Sub
Public Sub set_temp_problem()
Dim no%, i%
last_problem_input% = last_problem_input% + 1
no% = last_problem_input%
Call get_wenti_from_problem(temp_problem(no%))
End Sub
Public Function arrange_points_by_degree(in_p() As Integer, out_p() As Integer, total_p%, chose_degree%) As Boolean
'in_p() 排序 输出out_p()
Dim tp(7) As Integer
Dim i%, j%
For i% = 0 To total_p% - 1 '复制输入,初始化out_p()
tp(i%) = in_p(i%) '
out_p(i%) = 0
Next i%
'*********************************************************************
For i% = total_p% - 1 To 0 Step -1
 If m_poi(tp(2 * i% + 1)).data(0).degree = chose_degree Then '如果选中一点的自由度,排在前
    out_p(0) = tp(2 * i% + 1)
    out_p(1) = tp(2 * i%)
 ElseIf m_poi(tp(2 * i%)).data(0).degree = chose_degree Then
    out_p(0) = tp(2 * i%)
    out_p(1) = tp(2 * i% + 1)
 Else
    out_p(0) = tp(0)
    out_p(1) = tp(1)
    out_p(2) = tp(2)
    out_p(3) = tp(3)
    GoTo arrange_points_by_degree
 End If
 i% = 2 * i% + 1
 For j% = 1 To total_p% - 2
  out_p(j% + 1) = tp((i% + j%) Mod total_p%)
 Next j%
 arrange_points_by_degree = True
 Exit Function
arrange_points_by_degree:
Next i%
End Function



Public Function read_free_point_in_line(ByVal l_n%, ByVal start_p%, ByVal direct_p%, p%) As Boolean
Dim tn(1) As Integer
Dim ste_p%, i%
Call line_number0(start_p%, direct_p%, tn(0), tn(1))
If tn(0) < tn(1) Then
tn(1) = m_lin(l_n%).data(0).data0.in_point(i%)
ste_p% = 1
Else
tn(1) = 1
ste_p% = -1
End If
p% = 0
For i% = tn(0) + ste_p% To tn(1) Step ste_p%
    If m_lin(l_n%).data(0).data0.in_point(i%) = _
                 m_lin(l_n%).data(0).data0.poi(0) And _
                  m_poi(m_lin(l_n%).data(0).data0.poi(0)).data(0).degree > 0 Then
                   p% = m_lin(l_n%).data(0).data0.poi(0)
                    read_free_point_in_line = True
    ElseIf m_lin(l_n%).data(0).data0.in_point(i%) = _
                 m_lin(l_n%).data(0).data0.poi(1) And _
                  m_poi(m_lin(l_n%).data(0).data0.poi(1)).data(0).degree > 0 Then
                    p% = m_lin(l_n%).data(0).data0.poi(1)
                      read_free_point_in_line = True
   End If
Next i%
End Function
Public Function reduce_degree_for_point(ByVal p%, ByVal p1%, ByVal p2%, ByVal p3%)
'降低的约束
If m_poi(p%).data(0).parent.co_degree = 1 Then
   If m_poi(p%).data(0).parent.element(0).ty = line_ Then
      Call reduce_degree_for_point_line(p%, _
               m_poi(p%).data(0).parent.element(0).no, p1%, p2%, p3%)
   ElseIf m_poi(p%).data(0).parent.element(0).ty = circle_ Then
      Call reduce_degree_for_point_circle(p%, _
               m_poi(p%).data(0).parent.element(0).no, p1%, p2%, p3%)
   End If
ElseIf m_poi(p%).data(0).parent.co_degree = 2 Then
   If m_poi(p%).data(0).parent.element(0).ty = line_ Then
      If reduce_degree_for_point_line(p%, _
               m_poi(p%).data(0).parent.element(0).no, p1%, p2%, p3%) Then
            Call reduce_degree_for_point(ByVal p%, ByVal p1%, ByVal p2%, ByVal p3%)
      End If
   ElseIf m_poi(p%).data(0).parent.element(0).ty = circle_ Then
      Call reduce_degree_for_point_circle(p%, _
               m_poi(p%).data(0).parent.element(0).no, p1%, p2%, p3%)
   End If
End If
End Function
Public Function reduce_degree_for_point_line(ByVal p%, ByVal l%, ByVal p1%, ByVal p2%, ByVal p3%) As Boolean
Dim i_n%, tp%, tl%, tc%
'转化直线的生成条件,ByVal p1%, ByVal p2%, ByVal p3%排除的点
i_n% = -1
' 确定上的可转化点
  If m_lin(l%).data(0).data0.poi(1) <> p1% And _
       m_lin(l%).data(0).data0.poi(1) <> p2% And _
        m_lin(l%).data(0).data0.poi(1) <> p3% Then
     If m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).parent.co_degree <= 2 Then
        i_n% = 1
     End If
  ElseIf m_lin(l%).data(0).data0.poi(0) = p2% And _
       m_lin(l%).data(0).data0.poi(0) <> p1% And _
        m_lin(l%).data(0).data0.poi(0) <> p3% Then
     If m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).parent.co_degree <= 2 Then
        i_n% = 0
     End If
  'ElseIf m_lin(l%).data(0).data0.poi(0) = p3% And _
       m_lin(l%).data(0).data0.poi(0) <> p2% And _
        m_lin(l%).data(0).data0.poi(0) <> p1% Then
   '  If m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).degree > 0 Then
     '   i_n% = 1
   '  End If
  'ElseIf m_lin(l%).data(0).data0.poi(1) = p1% And _
   '    m_lin(l%).data(0).data0.poi(1) <> p2% And _
        m_lin(l%).data(0).data0.poi(1) <> p3% Then
    ' If m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).degree > 0 Then
      '  i_n% = 0
     'End If
  'ElseIf m_lin(l%).data(0).data0.poi(1) = p2% And _
       m_lin(l%).data(0).data0.poi(1) <> p1% And _
        m_lin(l%).data(0).data0.poi(1) <> p3% Then
   '  If m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).degree > 0 Then
    '    i_n% = 0
    ' End If
  'ElseIf m_lin(l%).data(0).data0.poi(1) = p3% And _
       m_lin(l%).data(0).data0.poi(1) <> p2% And _
        m_lin(l%).data(0).data0.poi(1) <> p1% Then
   '  If m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).degree > 0 Then
   '     i_n% = 0
   '  End If
  ElseIf m_lin(l%).data(0).data0.poi(0) <> p1% And _
          m_lin(l%).data(0).data0.poi(0) <> p2% And _
           m_lin(l%).data(0).data0.poi(0) <> p3% And _
         m_lin(l%).data(0).data0.poi(1) <> p1% And _
          m_lin(l%).data(0).data0.poi(1) <> p2% And _
           m_lin(l%).data(0).data0.poi(1) <> p3% Then
     If m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).parent.co_degree <= 2 Then
        i_n% = 0
     ElseIf m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).parent.co_degree <= 2 Then
        i_n% = 1
     End If
  End If
     '可转化的端点
     If i_n% = 0 Then
      tp% = m_lin(l%).data(0).data0.poi(0)
      m_lin(l%).data(0).data0.poi(0) = p%
     Else
      tp% = m_lin(l%).data(0).data0.poi(1)
      m_lin(l%).data(0).data0.poi(1) = p%
     End If
  'p%与tp%转化
  If m_poi(p%).data(0).parent.element(0).no = l% And _
        m_poi(p%).data(0).parent.element(0).ty = line_ Then
     m_poi(p%).data(0).parent.element(0) = m_poi(p%).data(0).parent.element(1)
  'ElseIf m_poi(p%).data(0).parent.element(1) = l% Then
  End If
    m_poi(p%).data(0).parent.element(1).ty = 0
     m_poi(p%).data(0).parent.element(1).no = 0
      m_poi(p%).data(0).p.degree = m_poi(p%).data(0).degree + 1
    m_lin(l%).data(0).depend_poi(i_n%) = p%
  If m_poi(tp%).data(0).degree = 2 Then
     m_poi(tp%).data(0).degree = 1
     m_poi(tp%).data(0).parent.element(0).ty = line_
     m_poi(tp%).data(0).parent.element(0).no = l%
     Call C_display_wenti.set_m_no_(m_poi(tp%).data(0).from_wenti_no, -1000)
        reduce_degree_for_point_line = True
  ElseIf m_poi(tp%).data(0).degree = 1 Then
     If m_poi(tp%).data(0).parent.element(0).ty = line_ Then
        tl% = m_poi(tp%).data(0).parent.element(0).no
        m_poi(tp%).data(0).degree = 0
         m_poi(tp%).data(0).parent.element(0).ty = line_
          m_poi(tp%).data(0).parent.element(0).no = l%
         m_poi(tp%).data(0).parent.element(1).ty = line_
          m_poi(tp%).data(0).parent.element(1).no = tl%
     ElseIf m_poi(tp%).data(0).parent.element(0).ty = circle_ Then
        tc% = m_poi(tp%).data(0).parent.element(0).no
        m_poi(tp%).data(0).degree = 0
         m_poi(tp%).data(0).parent.element(0).ty = line_
          m_poi(tp%).data(0).parent.element(0).no = l%
         m_poi(tp%).data(0).parent.element(1).ty = circle_
          m_poi(tp%).data(0).parent.element(1).no = tc%
     End If
          Call C_display_wenti.set_m_no_(m_poi(tp%).data(0).from_wenti_no, -1000)
        reduce_degree_for_point_line = True
  End If
End Function
Public Function reduce_degree_for_point_circle(ByVal p%, ByVal c%, ByVal p1%, ByVal p2%, ByVal p3%) As Boolean
    'if m_circ(c%).
End Function

Public Sub from_draw_to_input(ByVal ty As Integer, ByVal p1%, _
                        ele1 As condition_type, ele2 As condition_type, tangent_type As Integer, is_no_need_pre_input As Boolean)
Dim LastInput0No%, lastInput0ConditionNo%
Dim i%, k%, j%, l%, w_n%
Dim ch$
Dim s$
Dim D1&
Dim tp(1) As POINTAPI
Dim di_r(1) As Integer
Dim vf As POINTAPI
 If m_temp_line_for_input.is_using Then
    m_temp_line_for_input.data(0).poi(1) = p1%
                 Call draw_temp_line_for_input(1) '消除临时输入数据
     temp_line(draw_line_no) = line_number(m_temp_line_for_input.data(0).poi(0), m_temp_line_for_input.data(0).poi(1), _
           m_temp_line_for_input.data(0).end_point_coord(0), m_temp_line_for_input.data(0).end_point_coord(1), _
            depend_condition(point_, m_temp_line_for_input.data(0).poi(0)), _
             depend_condition(point_, m_temp_line_for_input.data(0).poi(1)), _
              condition, condition_color, 1, 0) '建立输入直线的数据
     l% = temp_line(draw_line_no)
             m_temp_line_for_input.data(0) = m_lin(0).data(0).data0
             m_temp_line_for_input.data(1) = m_lin(0).data(0).data0
                   m_temp_line_for_input.is_using = False
              If m_poi(p1%).data(0).parent.inter_type > 0 Then
              Call from_draw_to_input(ty, p1%, ele1, ele2, tangent_type, is_no_need_pre_input) '重新调用，建立输入语句
             End If
 ElseIf m_temp_circle_for_input.is_using And ((list_type_for_draw = 1 And draw_step = 1) Or (list_type_for_draw = 2 And draw_step = 2)) Then
    i% = last_conditions.last_cond(1).circle_no
   k% = m_circle_number(1, m_temp_circle_for_input.data(0).center, m_temp_circle_for_input.data(0).c_coord, _
        p1%, m_temp_circle_for_input.data(0).in_point(1), m_temp_circle_for_input.data(0).in_point(2), _
            m_temp_circle_for_input.data(0).radii, 0, 0, 0, 1, condition, condition_color, True)
      m_temp_circle_for_input.data(0) = m_Circ(0).data(0).data0
                m_temp_circle_for_input.is_using = False
     Call from_draw_to_input(ty, p1%, ele1, ele2, tangent_type, is_no_need_pre_input) '重新调用，建立输入语句
Else
Select Case operator
         ' Call set_wenti_cond0(p1%, ty, 0)
  Case "draw_point_and_line"
           If list_type_for_draw = 1 And (draw_step = 0 Or draw_step = 1) Then
             Call set_wenti_cond0(p1%, ty, ele1.no, is_no_need_pre_input)
           ElseIf list_type_for_draw = 2 And draw_step >= 1 Then '中点
              Call set_wenti_cond5_15(temp_point(0).no, 0, temp_point(1).no, 0, 5, 0)
           ElseIf list_type_for_draw = 3 And draw_step = 1 Then '定比 '
              Call set_wenti_cond6_6(6, temp_point(0).no, temp_point(1).no)
              Exit Sub
           ElseIf list_type_for_draw = 4 And draw_step >= 1 Then '定长
               Call set_wenti_cond_6_42_43_57(temp_point(0).no, temp_point(1).no, ele1, ele2, 0)
               ' Call set_wenti_cond_43_42(temp_point(0).no, temp_point(1).no, ele1, ele2, 0)
           ElseIf list_type_for_draw = 5 And draw_step = 3 Then '等长
             Call set_wenti_cond_1(temp_point(3).no, temp_point(2).no, temp_point(0).no, temp_point(1).no, _
                0, ele1.ty, ele1.no, ele2.ty, ele2.no, temp_point(3).no) '□□＝□□
           ElseIf list_type_for_draw = 6 And draw_step = 4 Then '画角平分线
              Call set_wenti_cond_50_51_52_56(temp_point(0).no, temp_point(1).no, temp_point(3).no, _
                temp_point(4).no, 0)
           'ElseIf draw_step Mod 2 = 0 Then
           '  Call set_wenti_cond0(p1%, ty, 0)
           End If
   Case "draw_circle"
           Call set_wenti_cond0(p1%, ty, 0, is_no_need_pre_input)
           If (list_type_for_draw = 1 And draw_step = 1) Or (list_type_for_draw = 2 And draw_step = 2) Then
                    Call set_wenti_cond7(temp_circle(0), temp_point(draw_step).no)
           ElseIf list_type_for_draw = 3 And draw_step = 4 Then
             'If draw_step = 3 Then
                    'Call set_wenti_cond0(p1%, ty, 0)
             'If draw_step = 4 Then
                    If ele1.ty = line_ Then
                     Call set_wenti_cond_2_33_44(m_lin(ele1.no).data(0).tangent_line_no)
                    ElseIf ele2.ty = line_ Then
                     Call set_wenti_cond_2_33_44(m_lin(ele2.no).data(0).tangent_line_no)
                    End If
             'End If
           ElseIf list_type_for_draw = 4 And draw_step = 6 Then
                    If ele1.ty = line_ Then
                     Call set_wenti_cond_2_33_44(m_lin(ele1.no).data(0).tangent_line_no)
                    ElseIf ele2.ty = line_ Then
                     Call set_wenti_cond_2_33_44(m_lin(ele2.no).data(0).tangent_line_no)
                    End If
           ElseIf list_type_for_draw = 5 Then
               If draw_step = 4 Then
                If ele1.ty = circle_ Then
                    Call set_wenti_cond12_64_65(ele1.no, temp_circle(0), p1%, 12)
                ElseIf ele1.ty = line_ Then
                End If
               End If
           End If
   Case "paral_verti"
       Call set_wenti_cond0(p1%, ty, 0, is_no_need_pre_input)
      If list_type_for_draw = 1 And draw_step = 4 Then '平行垂直
        Call set_wenti_cond2_3(temp_point(4).no, temp_point(0).no, temp_point(1).no, temp_point(2).no, paral_or_verti, 0)
      ElseIf list_type_for_draw = 2 And draw_step = 2 Then '垂直平分线
        Call set_wenti_cond4(temp_point(0).no, temp_point(1).no, temp_point(2).no, temp_point(3).no, temp_line(0), 0)
      End If
   Case "epolygon"
      Call set_wenti_cond0(p1%, ty, 0, is_no_need_pre_input)
  End Select
Select Case ty
'****************************************************************************************************************
Case interset_point_line_line
'***********************************************************************************************************************
'*************************************************************************************************
Case new_point_on_line_circle, _
    new_point_on_line_circle12, new_point_on_line_circle21
'****************************************************************************************************8
If operator = "draw_point_and_line" Then '
  If list_type_for_draw = 1 And draw_step < 2 Then
 If list_type_for_draw = 4 And draw_step = 2 Then
    ' Call set_wenti_cond_43_42(temp_point(2).no, temp_point(0).no, ele1, ele2, ty)
  ElseIf list_type_for_draw = 6 And draw_step = 4 Then
    Call set_wenti_cond_52(temp_point(0).no, temp_point(1).no, temp_point(2).no, temp_point(4).no, ele2.no, ty) 'ty 交点的类型
  Else
             'Call set_wenti_cond9(ele1.no, ele2.no, p1%)
   End If
 ElseIf operator = "paral_verti" Then
     If list_type_for_draw = 1 And draw_step = 3 Then

     ElseIf list_type_for_draw = 2 And draw_step = 2 Then
       Call set_wenti_cond_53(temp_point(0).no, temp_point(1).no, temp_point(2).no, temp_point(3).no, ele2.no, ty)        '-53 □□的垂直平分线交⊙□(_)□
                                          '  temp_point(3).no线段的中点                                                       '-25 □□的垂直平分线交⊙□□□于□
     Else
     End If
 ElseIf operator = "draw_circle" Then
Else
'Call C_display_wenti.set_display_string("", -1, 0)
'inpcond(10) = "过⊙□_外一点□和圆上一点□的直线交圆于□"
'inpcond(11) = □是直线□□与⊙□(_)的一个交点
Call distance_point_to_line(tp(0), _
                 m_poi(m_lin(ele1.no).data(0).data0.poi(0)).data(0).data0.coordinate, paral_, _
                  m_poi(m_lin(ele1.no).data(0).data0.poi(0)).data(0).data0.coordinate, _
                   m_poi(m_lin(ele1.no).data(0).data0.poi(1)).data(0).data0.coordinate, D1&, vf, 1)
 D1& = Abs(D1&) '切线
If Abs(D1& - m_Circ(ele2.no).data(0).data0.radii) < 5 Then
   If MsgBox(LoadResString_(1375, "\\1\\" + m_poi(m_lin(ele1.no).data(0).data0.poi(0)).data(0).data0.name + _
      m_poi(m_lin(ele1.no).data(0).data0.poi(1)).data(0).data0.name + _
           "\\2\\" + s$), 4, "", "", 0) = 6 Then
  ' Call set_wenti_cond_33_44(m_lin(ele1.no).data(0).data0.poi(0), _
          m_lin(ele1.no).data(0).data0.poi(1), ele2.no, 1, 0, 0, 0)
'****************************************************
 GoTo from_draw_to_input_mark10
End If
End If
'*******************************************
'**********************************************
from_draw_to_input_mark10:
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
   draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End If
End If
'****************************************************************************************
Case new_point_on_circle_circle, _
  new_point_on_circle_circle12, new_point_on_circle_circle21
'************************************************************************************************
  If operator = "draw_point_and_line" Then 'And draw_step = 0 Then '圆与圆的交点 'inpcond(13) = □是⊙□(_)和⊙□(_)的一个交点   '14
    ' Call set_wenti_cond13(ele1.no, ele2.no, temp_point(0).no, ty)
   'inpcond(12) = "⊙□_和⊙□_相交于点□取另一个交点□"  '13
 'inpcond(13) = □是⊙□(_)和⊙□(_)的一个交点   '14
   ElseIf operator = "draw_circle" Then
         If list_type_for_draw = 1 And draw_step = 1 Then
             'If aid_circle_for_input > 0 Then
             '  Call set_wenti_cond13(ele1.no, ele2.no, p1%, ty)
             '  Call set_wenti_cond7(aid_circle_for_input, p1%)
             'Else
               Call set_wenti_cond7(ele1.no, p1%) '7 ⊙□[down\\(_)]上任取一点□'-61⊙□□□上任取一点□
               Call set_wenti_cond7(ele2.no, p1%)
             'End If
         ElseIf list_type_for_draw = 2 And draw_step = 2 Then
        End If
    If operator = "draw_circle" And list_type_for_draw = 2 And draw_step = 2 Then

    ElseIf list_type_for_draw = 4 And draw_step = 2 Then
   ' Call set_wenti_cond_43_42(temp_point(0).no, temp_point(2).no, ele1, ele2, ty)
    ElseIf list_type_for_draw = 5 And draw_step = 3 Then '等长
      If draw_step = 5 Then
       temp_point(4).no = temp_point(5).no
      End If
    'If ele1.no = temp_circle(0) Then
      Call set_wenti_cond_30_58(ele1.no, ele2.no, _
         temp_point(3).no, temp_point(2).no, temp_point(0).no, temp_point(1).no, ty)
    'ElseIf ele2.no = temp_circle(0) Then
       'Call set_wenti_cond_31_30(ele1.no, _
         temp_point(4).no, temp_point(3).no, temp_point(0).no, temp_point(1).no, -30, ele2.no, ty)
   'End If
     End If
ElseIf operator = "draw_circle" Then
If m_Circ(ele1.no).data(0).data0.center > 0 Then
tp(0) = m_poi(m_Circ(ele1.no).data(0).data0.center).data(0).data0.coordinate
Else
tp(0) = m_Circ(ele1.no).data(0).data0.c_coord
End If
If m_Circ(ele2.no).data(0).data0.center > 0 Then
tp(1) = m_poi(m_Circ(ele2.no).data(0).data0.center).data(0).data0.coordinate
Else
tp(1) = m_Circ(ele2.no).data(0).data0.c_coord
End If
D1& = sqr((tp(0).X - tp(1).X) ^ 2 + (tp(0).Y - tp(1).Y) ^ 2)
di_r(0) = m_Circ(ele1.no).data(0).data0.radii + m_Circ(ele2.no).data(0).data0.radii
di_r(1) = Abs(m_Circ(ele1.no).data(0).data0.radii - m_Circ(ele2.no).data(0).data0.radii)
If Abs(D1& - di_r(0)) < 5 Or Abs(D1& - di_r(1)) < 5 Then
  If MsgBox(LoadResString_(1420, "\\1\\" + m_poi(m_Circ(ele1.no).data(0).data0.center).data(0).data0.name + _
       "\\2\\" + m_poi(m_Circ(ele2.no).data(0).data0.center).data(0).data0.name), 4, "", "", 0) = 6 Then
'Call set_wenti_cond12_64_65(ele1.no, ele2.no, p1%, 1)
 Exit Sub
  End If
End If
'Call set_wenti_cond13(ele1.no, ele2.no, p1%, ty)
'13 □是⊙□[down\\(_)]和⊙□[down\\(_)]一个交点
'-66 □是⊙□□□和⊙□[down\\(_)]一个交点
'-67  □是⊙□□□和⊙□□□一个交点
    draw_wenti_no = C_display_wenti.m_last_input_wenti_no
End If

'********************************************************************************************************
Case new_point_on_circle
    If list_type_for_draw = 6 And operator = "draw_point_and_line" Then
        Call set_wenti_cond_1(temp_point(3).no, temp_point(2).no, temp_point(0).no, temp_point(1).no, _
                0, ele1.ty, ele1.no, ele2.ty, ele2.no, temp_point(3).no)
    End If
   If ele1.ty = line_ Then
   'If list_type_for_draw = 4 Then
    Call set_wenti_cond_6(m_Circ(ele1.no).data(0).data0.center, p1%, Trim(MDIForm1.Text2.text), ele1.no, 0)
   ElseIf ele1.ty = 0 Then 'list_type_for_draw = 5 Then
    Call from_draw_to_input(new_point_on_circle, p1%, ele2, ele1, tangent_type, is_no_need_pre_input)
   End If
'******************************************************************************************************************
'*************************************************************************************************
  End Select
  End If
End Sub

Public Sub set_wenti_cond_70(ByVal p%, ByVal l%)
'-70 直线□□上取一定点□
End Sub

Private Sub input_from_point_ty(in_p%)
If m_poi(in_p%).data(0).parent.inter_type = exist_point Then
  Exit Sub
ElseIf m_poi(in_p%).data(0).parent.inter_type = new_point_on_circle_circle12 Or _
                m_poi(in_p%).data(0).parent.inter_type = new_point_on_circle_circle21 Then
  Call set_wenti_cond13(m_poi(in_p%).data(0).parent.element(1).no, m_poi(in_p%).data(0).parent.element(2).no, _
                 in_p%, m_poi(in_p%).data(0).parent.inter_type) '
ElseIf m_poi(in_p%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
                m_poi(in_p%).data(0).parent.inter_type = new_point_on_line_circle21 Then
 Call set_wenti_cond11(m_poi(in_p%).data(0).parent.element(1).no, m_poi(in_p%).data(0).parent.element(2).no, in_p%, 0)
                        'm_poi(in_p%).data(0).parent.inter_type, 0)  '
ElseIf m_poi(in_p%).data(0).parent.inter_type = interset_point_line_line Then
 Call set_wenti_cond9(m_poi(in_p%).data(0).parent.element(1).no, m_poi(in_p%).data(0).parent.element(2).no, in_p%) '
ElseIf m_poi(in_p%).data(0).parent.inter_type = new_point_on_circle Then
 Call set_wenti_cond7(m_poi(in_p%).data(0).parent.element(1).no, in_p%) '
ElseIf m_poi(in_p%).data(0).parent.inter_type = new_point_on_line Then
 Call set_wenti_cond1(m_poi(in_p%).data(0).parent.element(1).no, in_p%) '
End If

End Sub

Public Sub set_wenti_cond0(ByVal p%, ty As Integer, tangent_line_no%, Optional is_no_need_pre_input As Boolean = False)
Dim Wn%
If ty = exist_point Or is_no_need_pre_input Then ' t_line_no%<0 输入不作预处理，如角的平分线交支线于
   Exit Sub
ElseIf m_poi(p%).data(0).parent.inter_type = tangent_point_ Then
     Call set_wenti_cond_2_33_44(m_lin(tangent_line_no%).data(0).tangent_line_no)
ElseIf m_poi(p%).data(0).parent.inter_type = new_point_on_line Then
  Call set_wenti_cond1(m_poi(p%).data(0).parent.element(1).no, p%) '直线上任取一点
ElseIf m_poi(p%).data(0).parent.inter_type = new_point_on_circle Then
  Call set_wenti_cond7(m_poi(p%).data(0).parent.element(1).no, p%)  '7 ⊙□[down\\(_)]上任取一点□'-61⊙□□□上任取一点□
ElseIf m_poi(p%).data(0).parent.inter_type = interset_point_line_line Then
  Call set_wenti_cond9(m_poi(p%).data(0).parent.element(1).no, _
             m_poi(p%).data(0).parent.element(2).no, p%)
ElseIf m_poi(p%).data(0).parent.inter_type = new_point_on_line_circle12 Or _
         m_poi(p%).data(0).parent.inter_type = new_point_on_line_circle21 Then
      Call set_wenti_cond11(m_poi(p%).data(0).parent.element(1).no, m_poi(p%).data(0).parent.element(2).no, _
            p%, m_poi(p%).data(0).parent.inter_type)
ElseIf m_poi(p%).data(0).parent.inter_type = new_point_on_circle_circle12 Or _
         m_poi(p%).data(0).parent.inter_type = new_point_on_circle_circle21 Then
    Call set_wenti_cond13(m_poi(p%).data(0).parent.element(1).no, m_poi(p%).data(0).parent.element(2).no, _
            p%, m_poi(p%).data(0).parent.inter_type)
End If
End Sub

Public Sub set_wenti_cond_2_33_44(tangent_line_no%)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).parent.inter_type = _
  Abs(m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).parent.inter_type)
 m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).parent.inter_type = _
  Abs(m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).parent.inter_type)
If tangent_line(tangent_line_no%).is_display_in_wenti_data = False Then
 If tangent_line(tangent_line_no%).data(0).circ(0) > 0 And tangent_line(tangent_line_no%).data(0).circ(1) > 0 Then
    Call set_wenti_cond_2(tangent_line(tangent_line_no%).data(0).poi(0), _
                           tangent_line(tangent_line_no%).data(0).poi(1), _
                            tangent_line(tangent_line_no%).data(0).circ(0), _
                             tangent_line(tangent_line_no%).data(0).circ(1), _
                              tangent_line(tangent_line_no%).data(0).line_no, _
                               tangent_line_no%)
         tangent_line(tangent_line_no%).is_display_in_wenti_data = True
 ElseIf tangent_line(tangent_line_no%).data(0).circ(0) > 0 Then
    Call set_wenti_cond_33_44(tangent_line(tangent_line_no%).data(0).poi(0), _
         tangent_line(tangent_line_no%).data(0).circ(0), tangent_line(tangent_line_no%).tangent_type, _
          tangent_line(tangent_line_no%).data(0).poi(1), tangent_line_no%)
        tangent_line(tangent_line_no%).is_display_in_wenti_data = True
 ElseIf tangent_line(tangent_line_no%).data(0).circ(1) > 0 Then
     Call set_wenti_cond_33_44(tangent_line(tangent_line_no%).data(0).poi(0), _
         tangent_line(tangent_line_no%).data(0).circ(1), tangent_line(tangent_line_no%).tangent_type, _
          tangent_line(tangent_line_no%).data(0).poi(1), tangent_line(tangent_line_no%).data(0).line_no)
        tangent_line(tangent_line_no%).is_display_in_wenti_data = True
 End If
End If
End Sub

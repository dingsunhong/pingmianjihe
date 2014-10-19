Attribute VB_Name = "DATA_BASE"
Option Explicit
Global input_record0 As record_data0_type

Function inter_point_circle_circle(ByVal c1%, _
  ByVal c2%, in_coord As POINTAPI, Optional inter_ty As Integer = 0) As Integer
Dim i%, j%, n%, p% '???
Dim r!, s&, sr!, temp_num0!
Dim temp_k1!, temp_k2!, temp_x!, temp_y!
Dim di_r(1) As Long
Dim cir1 As circle_data0_type
Dim cir2 As circle_data0_type
'计算新点座标
If c1% > c2% Then
 Call exchange_two_integer(c1%, c2%)
End If
If c1% > 0 Then
cir1 = m_Circ(c1%).data(0).data0
End If
If c2% > 0 Then
cir2 = m_Circ(c2%).data(0).data0
End If
  s& = ((cir1.c_coord.X) - _
          cir2.c_coord.X) ^ 2 + _
          (cir1.c_coord.Y - _
             cir2.c_coord.Y) ^ 2
              sr! = distance_of_two_POINTAPI(cir1.c_coord, cir2.c_coord) '圆心距
              di_r(0) = cir1.radii + cir2.radii '两圆半径和
               di_r(1) = cir1.radii - cir2.radii '两圆半径差
If sr! > di_r(0) + 5 Or _
    sr! < Abs(di_r(1)) - 5 Then '圆心距大于半径和或小于半径差，两圆无交点
     inter_point_circle_circle = -1
      Exit Function
'两圆分离或包含，无交点
ElseIf sr! <= di_r(0) + 5 And sr! > di_r(0) Then '等于和
 sr! = di_r(0)
   t_coord = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
         time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(c1%).data(0).data0.c_coord), m_Circ(c1%).data(0).data0.radii / di_r(0)))
     If inter_ty = 0 Then
        If distance_of_two_POINTAPI(t_coord, in_coord) < 5 Then
           in_coord = t_coord
           inter_point_circle_circle = new_point_on_circle_circle12
        End If
     Else
         in_coord = t_coord
     End If
ElseIf sr! >= Abs(di_r(1)) - 5 And sr! < Abs(di_r(1)) Then '等于差
 sr! = di_r(1) '
    t_coord = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
         time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(c1%).data(0).data0.c_coord), m_Circ(c1%).data(0).data0.radii / di_r(1)))
     If inter_ty = 0 Then
        If distance_of_two_POINTAPI(t_coord, in_coord) < 5 Then
           in_coord = t_coord
           inter_point_circle_circle = new_point_on_circle_circle12
        End If
     Else
         in_coord = t_coord
     End If
Else
'************************************************************
 If inter_point_circle_circle_(cir1, cir2, t_coord1, 0, t_coord2, 0, _
                                   s&, sr!, False) > 0 Then
  If inter_ty = 0 Then
  If Abs(in_coord.X - t_coord1.X) < 5 And Abs(in_coord.Y - t_coord1.Y) < 5 Then
   If c1% > 0 And c2% > 0 Then
   inter_point_circle_circle = new_point_on_circle_circle12
   End If
    t_coord = t_coord1
 ElseIf Abs(in_coord.X - t_coord2.X) < 5 And Abs(in_coord.Y - t_coord2.Y) < 5 Then
   If c1% > 0 And c2% > 0 Then
   inter_point_circle_circle = new_point_on_circle_circle21
   End If
   t_coord = t_coord2
End If
   in_coord = t_coord
Else
 inter_point_circle_circle = inter_ty
 If inter_ty = new_point_on_circle_circle12 Then
    in_coord = t_coord1
 ElseIf inter_ty = new_point_on_circle_circle21 Then
    in_coord = t_coord2
 End If
End If
                'last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
               MDIForm1.Toolbar1.Buttons(21).Image = 33 '***
'**************************************************************************
End If
End If
End Function

Function inter_point_circle_circle_for_change(ByVal c1%, _
  ByVal c2%, in_coord As POINTAPI, Optional inter_ty As Integer = 0) As Integer
Dim i%, j%, n%, p% '???
Dim r!, s&, sr!, temp_num0!
Dim temp_k1!, temp_k2!, temp_x!, temp_y!
Dim di_r(1) As Long
Dim cir1 As circle_data0_type
Dim cir2 As circle_data0_type
'计算新点座标
If c1% > c2% Then
 Call exchange_two_integer(c1%, c2%)
End If
If c1% > 0 Then
  cir1 = m_Circ(c1%).data(0).data0
End If
If c2% > 0 Then
  cir2 = m_Circ(c2%).data(0).data0
End If
'  s& = ((cir1.c_coord.X) - _
          cir2.c_coord.X) ^ 2 + _
          (cir1.c_coord.Y - _
             cir2.c_coord.Y) ^ 2
''              sr! = distance_of_two_POINTAPI(cir1.c_coord, cir2.c_coord) '圆心距
'              di_r(0) = cir1.radii + cir2.radii '两圆半径和
'               di_r(1) = cir1.radii - cir2.radii '两圆半径差
'If sr! > di_r(0) + 5 Or _
'    sr! < Abs(di_r(1)) - 5 Then '圆心距大于半径和或小于半径差，两圆无交点
'     inter_point_circle_circle_for_change = -1
'      Exit Function
'两圆分离或包含，无交点
'ElseIf sr! <= di_r(0) + 5 And sr! > di_r(0) Then '等于和
' sr! = di_r(0)
'   t_coord = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
         time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(c1%).data(0).data0.c_coord), m_Circ(c1%).data(0).data0.radii / di_r(0)))
'     If inter_ty = 0 Then
'        If distance_of_two_POINTAPI(t_coord, in_coord) < 5 Then
'           in_coord = t_coord
'           inter_point_circle_circle_for_change = new_point_on_circle_circle12
 '       End If
'     Else
'         in_coord = t_coord
'     End If
'ElseIf sr! >= Abs(di_r(1)) - 5 And sr! < Abs(di_r(1)) Then '等于差
' sr! = di_r(1) '
'    t_coord = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
         time_POINTAPI_by_number(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
            m_Circ(c1%).data(0).data0.c_coord), m_Circ(c1%).data(0).data0.radii / di_r(1)))
'     If inter_ty = 0 Then
'        If distance_of_two_POINTAPI(t_coord, in_coord) < 5 Then
'           in_coord = t_coord
'           inter_point_circle_circle_for_change = new_point_on_circle_circle12
'        End If
'     Else
'         in_coord = t_coord
'     End If
'Else
'************************************************************
 If inter_point_circle_circle_(cir1, cir2, t_coord1, 0, t_coord2, 0, _
                                   s&, sr!, False) > 0 Then
  'If inter_ty = 0 Then
  'If Abs(in_coord.X - t_coord1.X) < 5 And Abs(in_coord.Y - t_coord1.Y) < 5 Then
  ' If c1% > 0 And c2% > 0 Then
  '  inter_point_circle_circle_for_change = new_point_on_circle_circle12
  ' End If
  '  t_coord = t_coord1
  'ElseIf Abs(in_coord.X - t_coord2.X) < 5 And Abs(in_coord.Y - t_coord2.Y) < 5 Then
  ' If c1% > 0 And c2% > 0 Then
  '  inter_point_circle_circle_for_change = new_point_on_circle_circle21
  '  End If
  '  t_coord = t_coord2
  'End If
  '  in_coord = t_coord
'Else
 inter_point_circle_circle_for_change = inter_ty
 If inter_ty = new_point_on_circle_circle12 Then
    in_coord = t_coord1
 ElseIf inter_ty = new_point_on_circle_circle21 Then
    in_coord = t_coord2
 End If
'End If
                'last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'              MDIForm1.Toolbar1.Buttons(21).Image = 33 '***
'**************************************************************************
End If
'End If
End Function

Function inter_point_circle_circle_(Circle1 As circle_data0_type, _
                                     Circle2 As circle_data0_type, _
                                      out_coord1 As POINTAPI, ByVal out_p1%, _
     out_coord2 As POINTAPI, ByVal out_p2%, s&, sr!, is_change As Boolean) As Integer
Dim r!, temp_num0!, temp_k1!, temp_k2!, temp_x!, temp_y!
Dim t_coord As POINTAPI
Dim c1 As POINTAPI
Dim c2 As POINTAPI
Dim Ra1&, rA2&
c1 = Circle1.c_coord
c2 = Circle2.c_coord
Ra1 = Circle1.radii
rA2 = Circle2.radii
       s& = (c1.X - c2.X) ^ 2 + (c1.Y - c2.Y) ^ 2
       sr! = sqr(s&)
If sr! > Ra1& + rA2& Or _
     sr! < Abs(Ra1& - rA2&) _
       Or sr! = 0 Then
  If sr! = 0 Then '同心圆
  out_coord1.X = 10000
  out_coord1.Y = 10000
  out_coord2.X = 10000
  out_coord2.Y = 10000
  ElseIf sr! > Ra1& + rA2& Then '
  t_coord = minus_POINTAPI(c2, c1)
  out_coord1 = add_POINTAPI(c1, time_POINTAPI_by_number(t_coord, Ra1& / sr!))
  out_coord2 = minus_POINTAPI(c2, time_POINTAPI_by_number(t_coord, rA2& / sr!))
  ElseIf sr! < Abs(Ra1& - rA2&) Then
   If Ra1& - rA2& > 0 Then
      t_coord = minus_POINTAPI(c2, c1)
      out_coord1 = add_POINTAPI(c1, time_POINTAPI_by_number(t_coord, Ra1& / sr!))
      out_coord2 = add_POINTAPI(c2, time_POINTAPI_by_number(t_coord, rA2& / sr!))
   Else
         t_coord = minus_POINTAPI(c1, c2)
      out_coord1 = add_POINTAPI(c1, time_POINTAPI_by_number(t_coord, Ra1& / sr!))
      out_coord2 = add_POINTAPI(c2, time_POINTAPI_by_number(t_coord, rA2& / sr!))
   End If
  End If
 inter_point_circle_circle_ = 0     '无交点
 Exit Function
ElseIf (sr! = Ra1& + rA2& Or _
     sr! = Abs(Ra1& - rA2&)) And sr! <> 0 Then
  out_coord1 = add_POINTAPI(c1, time_POINTAPI_by_number(minus_POINTAPI(c2, c1), Ra1& / sr!))
  out_coord2 = add_POINTAPI(c2, time_POINTAPI_by_number(minus_POINTAPI(c1, c2), rA2& / sr!))
  inter_point_circle_circle_ = 1    '无交点
Else
   temp_num0! = (s& + Ra1& ^ 2 - _
       rA2& ^ 2) / 2 / sr!
         's! = Sqr(s!)
          temp_k1! = (c2.X - c1.X) / sr!
           temp_k2! = (c2.Y - c1.Y) / sr!
          temp_x! = temp_k1! * temp_num0! + c1.X
          temp_y! = temp_k2! * temp_num0! + c1.Y
          temp_num0 = CLng((Ra1& + rA2& + sr!) / 2)
                '会出负数
               r! = (temp_num0! - Ra1&) * _
                     (temp_num0! - rA2&) * _
                       (temp_num0! - sr!) * temp_num0!
    If r! < 0 Then
     inter_point_circle_circle_ = 0
     Exit Function
    End If
                 temp_num0 = sqr(r!) * 2 / sr!
                   out_coord1.X = CInt(-temp_k2! * temp_num0! + temp_x!)
                   out_coord1.Y = CInt(temp_k1! * temp_num0! + temp_y!)
                   out_coord2.X = CInt(temp_k2! * temp_num0! + temp_x!)
                   out_coord2.Y = CInt(-temp_k1! * temp_num0! + temp_y!)
                   inter_point_circle_circle_ = 2
   If out_p1% > 0 Then
     Call set_point_coordinate(out_p1%, out_coord1, is_change)
   End If
   If out_p2% > 0 Then
     Call set_point_coordinate(out_p2%, out_coord2, is_change)
   End If
End If
End Function

Sub inter_point_circle_circle2(c1 As circle_data0_type, _
      c2 As circle_data0_type, ByVal p%, out_coord As POINTAPI, out_p%)
Dim s&, r&, t!
'两圆已交于一点
s& = (c1.c_coord.X - _
       c2.c_coord.X) ^ 2 + _
        (c1.c_coord.Y - _
          c2.c_coord.Y) ^ 2
If s& > 0 Then
r& = (c1.c_coord.X - _
       c2.c_coord.X) * _
          (m_poi(p%).data(0).data0.coordinate.Y - c1.c_coord.Y) - _
             (c1.c_coord.Y - _
                c2.c_coord.Y) * _
                   (m_poi(p%).data(0).data0.coordinate.X - _
                     c1.c_coord.X)
t! = 2 * r& / s&
out_coord.X = m_poi(p%).data(0).data0.coordinate.X + (c1.c_coord.Y - _
        c2.c_coord.Y) * t!
out_coord.Y = m_poi(p%).data(0).data0.coordinate.Y - (c1.c_coord.X - _
        c2.c_coord.X) * t!
        If out_p% > 0 Then
         Call set_point_coordinate(out_p%, out_coord, False)
        End If
End If
End Sub
Private Sub remove_point_from_eline(ByVal p%)
Dim i%, j%, k%
i% = 1
While i% <= last_conditions.last_cond(1).eline_no And i% > 0
  If Deline(i%).data(0).data0.poi(0) = p% Or Deline(i%).data(0).data0.poi(0) = p% Or _
       Deline(i%).data(0).data0.poi(2) = p% Or Deline(i%).data(0).data0.poi(3) = p% Then
      last_conditions.last_cond(1).eline_no = last_conditions.last_cond(1).eline_no - 1
  For j% = i% To last_conditions.last_cond(1).eline_no
   Deline(j%) = Deline(j% + 1)
  Next j%
  Else
  i% = i% - 1
  End If
Wend
End Sub

Private Sub remove_point_from_point_pair(ByVal p%)
Dim i%, j%, k%
i% = 1
While i% <= last_conditions.last_cond(1).dpoint_pair_no
   If Ddpoint_pair(i%).data(0).data0.poi(0) = p% Or Ddpoint_pair(i%).data(0).data0.poi(1) = p% Or _
       Ddpoint_pair(i%).data(0).data0.poi(2) = p% Or Ddpoint_pair(i%).data(0).data0.poi(3) = p% Or _
        Ddpoint_pair(i%).data(0).data0.poi(4) = p% Or Ddpoint_pair(i%).data(0).data0.poi(5) = p% Or _
         Ddpoint_pair(i%).data(0).data0.poi(6) = p% Or Ddpoint_pair(i%).data(0).data0.poi(7) = p% Then
   last_conditions.last_cond(1).dpoint_pair_no = last_conditions.last_cond(1).dpoint_pair_no - 1
   For j% = i% To last_conditions.last_cond(1).dpoint_pair_no
    Ddpoint_pair(j%) = Ddpoint_pair(j% + 1)
   Next j%
   Else
    i% = i% + 1
   End If
Wend
End Sub

Function right_triangle1(p1 As POINTAPI, p2 As POINTAPI, _
 dis&, x1&, y1&, x2&, y2&) As Integer
Dim r&, s&, sr&, temp_k1!, temp_k2!
Dim co!, si!
       s& = (p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2
        sr& = sqr(s&)
If sr& <= dis& Then
 x1& = 10000
  y1& = 10000
 x2& = 10000
  y2& = 10000
 right_triangle1 = 0 '无交点
Else
right_triangle1 = 1    '无交点
   '射影定理
         temp_k1! = CSng((p2.X - p1.X) / sr&)
           temp_k2! = CSng((p2.Y - p1.Y) / sr&)
         co! = CSng(dis& / sr&)
         si! = sqr(s& - dis& ^ 2) / sr&
          
                   x1& = p1.X + dis& * (co! * temp_k1! - si! * temp_k2!)
                    y1& = p1.Y + dis& * (si! * temp_k1! + co! * temp_k2!)
                   x2& = p1.X + dis& * (co! * temp_k1! + si! * temp_k2!)
                    y2& = p1.Y + dis& * (-si! * temp_k1! + co! * temp_k2!)



End If

End Function
Function right_triangle2(c%, p2 As POINTAPI, _
  ByVal input_x&, ByVal input_y&, p%) As Integer
Dim i%, j%, n% '???
Dim r!, s&, sr!, temp_num0!
Dim temp_k1!, temp_k2!, temp_x!, temp_y!
'Dim X1&, Y1&, X2&, Y2&
'计算新点座标
If right_triangle1(m_Circ(c%).data(0).data0.c_coord, p2, m_Circ(c%).data(0).data0.radii, _
               t_coord1.X, t_coord1.Y, t_coord2.X, t_coord2.Y) = 0 Then
right_triangle2 = 0
 '************************************************************
Else
 
If Abs(input_x& - t_coord1.X) < 5 And Abs(input_y& - t_coord1.Y) < 5 Then
   right_triangle2 = new_point_on_circle_circle12
   t_coord = t_coord1
   ' X& = X1&
   '  Y& = Y1&
ElseIf Abs(input_x& - t_coord2.X) < 5 And Abs(input_y& - t_coord2.Y) < 5 Then
   right_triangle2 = new_point_on_circle_circle21
   t_coord = t_coord2
  'X& = X2&
   'Y& = Y2&
End If
 If p% = 0 Then
                last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
               MDIForm1.Toolbar1.Buttons(21).Image = 33 '***
                  p% = last_conditions.last_cond(1).point_no
End If
                  Call set_point_coordinate(p%, t_coord, False)
                  ' poi(p%).data(0).data0.coordinate.X = X&
                  '  poi(p%).data(0).data0.coordinate.Y = Y&
                  Call set_point_visible(p%, 1, False)
                  Call C_display_picture.set_m_point_color(p%, 0)
                  m_poi(p%).data(0).degree = 0
                  Call add_point_to_m_circle(p%, c%, 255)
              ' Call add_point_to_circle(p%, C2%)
'**************************************************************************

End If

End Function
Function inter_point_line_circle(line_no%, circle_no%, in_coord As POINTAPI, _
                       p1%, is_draw As Boolean, is_set_data As Boolean, Optional point_type As Integer = 0) As Integer
Dim i%, j%, k%, n%, s&, A!, b!, c&, r&, d&, radii&
Dim p(1) As Integer
Dim c_p As POINTAPI
Dim p_coord(2) As POINTAPI
Dim tc As circle_data0_type
Dim tl As line_data_type
Dim center_coord As POINTAPI
Dim vf As POINTAPI
Dim t_ele As condition_type
Dim temp_record As total_record_type
'   Call C_display_picture.get_end_point(ele1.no, p(0), 0, p_coord(0).X, p_coord(0).Y, _
                                        p_coord(1).X, p_coord(1).Y) ' 读直线端点的坐标
'*******************************************************************************************
       
      'If m_lin(line_no%).data(0).data0.poi(0) > 0 And _
       '   m_lin(line_no%).data(0).data0.poi(1) > 0 Then
        p_coord(0) = m_poi(m_lin(line_no%).data(0).data0.depend_poi(0)).data(0).data0.coordinate
        p_coord(1) = second_end_point_coordinate(line_no%)
      'Else
      '  p_coord(0) = m_lin(line_no%).data(0).data0.end_point_coord(0)
      '  p_coord(1) = m_lin(line_no%).data(0).data0.end_point_coord(1)
      'End If
      If m_Circ(circle_no%).data(0).data0.center > 0 Then
        center_coord = m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.coordinate
      Else
         center_coord = m_Circ(circle_no%).data(0).data0.c_coord
      End If
      radii& = m_Circ(circle_no%).data(0).data0.radii
      tl = m_lin(line_no%).data(0)
      tc = m_Circ(circle_no%).data(0).data0
'*******************************************************************
'计算新点座标
 '***********************************************************************
 '判断是否已有交点
If point_type = 0 Then
n% = 0
For i% = 1 To tl.data0.in_point(0)
 For j% = 1 To tc.in_point(0)
  If tl.data0.in_point(i%) = tc.in_point(j%) Then '已有交点
   If m_poi(tl.data0.in_point(i%)).data(0).parent.inter_type <> new_point_on_line_circle12 Or _
       m_poi(tl.data0.in_point(i%)).data(0).parent.inter_type <> new_point_on_line_circle21 Then '已有交点，单未记录交点类型
        If is_same_point(in_coord, tc.in_point(j%)) Then
           p1% = tl.data0.in_point(i%)
            'Exit Function
           GoTo inter_point_line_circle_mark1
        Else
           If (p(0)) = 0 Then
            p(0) = tl.data0.in_point(i%) '有一交点
           Else
            p(1) = tl.data0.in_point(i%) '有第二个交点
             Exit Function
           End If
        End If
   Else
   If n% = 0 Then
    p(0) = tl.data0.in_point(i%)
   ElseIf n% = 1 Then
    p(1) = tl.data0.in_point(i%)
   End If
     If (m_poi(tl.data0.in_point(i%)).data(0).data0.coordinate.X - in_coord.X) ^ 2 + _
        (m_poi(tl.data0.in_point(i%)).data(0).data0.coordinate.Y - in_coord.Y) ^ 2 < 5 Then '交点是输入点
         If n% = 0 Then
          inter_point_line_circle = new_point_on_line_circle12
         Else
          inter_point_line_circle = new_point_on_line_circle21
         End If
     End If
       n% = n% + 1
   End If
  End If
 Next j%
Next i%
End If
inter_point_line_circle_mark1:
  If distance_point_to_line(center_coord, p_coord(0), paral_, _
      p_coord(0), p_coord(1), r&, vf, 1) = False Then
      Exit Function
  End If
 If Abs(r&) >= radii& Then
   inter_point_line_circle = 0
    in_coord = add_POINTAPI(center_coord, time_POINTAPI_by_number(minus_POINTAPI(vf, center_coord), radii& / r&))
    Exit Function
    
 Else
    t_coord = minus_POINTAPI(p_coord(1), p_coord(0))
     b! = time_POINTAPI(t_coord, t_coord)
     If b! > 0 And tc.radii > r& Then
       b! = sqr((tc.radii ^ 2 - r& ^ 2) / b!)
      Else
       Exit Function
     End If
 End If
  p_coord(2) = time_POINTAPI_by_number(t_coord, b!)
  t_coord1 = add_POINTAPI(vf, p_coord(2))
  t_coord2 = minus_POINTAPI(vf, p_coord(2))
 If point_type = 0 Then
  If Abs(t_coord1.X - in_coord.X) < 5 And Abs(t_coord1.Y - in_coord.Y) < 5 Then
       in_coord = t_coord1 '新交点座标
       If line_no <> 0 And circle_no% > 0 Then
          inter_point_line_circle = new_point_on_line_circle12 '新交点类型
       'ElseIf line_no% = 0 And circle_no% > 0 Then
       '   inter_point_line_circle = new_point_on_Tline_circle12 '新交点类型
       'ElseIf line_no% <> 0 And circle_no% = 0 Then
       '   inter_point_line_circle = new_point_on_line_Tcircle12 '新交点类型
       'ElseIf line_no% = 0 And circle_no% = 0 Then
       '   inter_point_line_circle = new_point_on_Tline_Tcircle12 '新交点类型
       End If
  ElseIf Abs(t_coord2.X - in_coord.X) < 5 And Abs(t_coord2.Y - in_coord.Y) < 5 Then
       If line_no% <> 0 And circle_no% > 0 Then
          inter_point_line_circle = new_point_on_line_circle21 '新交点类型
       'ElseIf line_no% = 0 And circle_no% > 0 Then
       '   inter_point_line_circle = new_point_on_Tline_circle21 '新交点类型
       'ElseIf line_no% <> 0 And circle_no% = 0 Then
       '   inter_point_line_circle = new_point_on_line_Tcircle21 '新交点类型
       'ElseIf line_no% = 0 And circle_no% = 0 Then
       '   inter_point_line_circle = new_point_on_Tline_Tcircle21 '新交点类型
       End If
       in_coord = t_coord2 '新交点座标
  Else
           inter_point_line_circle = -1
           Exit Function
  End If
 'Else
  '          inter_point_line_circle = -1
   '          Exit Function
 'End If
'End If
          MDIForm1.Toolbar1.Buttons(21).Image = 33
        If is_set_data Then
        If p1% = 0 Then 'And is_set_data Then '设置新数据
                p1% = m_point_number(in_coord, condition, 1, condition_color, "", _
                    depend_condition(line_, line_no%), depend_condition(circle_, circle_no%), _
                    inter_point_line_circle, is_set_data)
                     Call add_point_to_line(p1%, line_no%, 0, display, is_draw, 0, temp_record.record_data)
                      Call add_point_to_m_circle(p1%, circle_no%, record0, 0)

           'If ele1.no > last_conditions.last_cond(1).line_no Then
           '   Call set_line_from_aid_line(ele1.no, p1%, in_coord, ele1.no) '从辅助线设置新线数据
           'End If
           'If p1% = 0 Then
           '   p1% = m_point_number(in_coord, condition, 1, condition_color, "", _
                    depend_condition(line_, ele1.no), depend_condition(circle_, ele2.no), _
                    inter_point_line_circle, is_set_data)
          '          If ele1.no = 0 Then
          '             inter_point_line_circle = new_point_on_circle
          '          End If
'                  m_poi(p1%).data(0).degree = 0
                   Call add_point_to_line(p1%, line_no%, 0, display, is_draw, 0, temp_record.record_data)
                   Call add_point_to_m_circle(p1%, circle_no%, record0, 0)
                   m_poi(p1%).data(0).parent.inter_type = new_inter_point_type(m_poi(p1%).data(0).parent.inter_type, inter_point_line_circle)
           '**************************************************************************************
           'm_poi(p1%).data(0).degree = 0
           If m_Circ(circle_no%).data(0).parent.co_degree = 2 And tl.parent.co_degree = 2 Then
              m_poi(p1%).data(0).degree_for_reduce = 0
           ElseIf m_Circ(circle_no%).data(0).parent.co_degree <= 1 And tl.parent.co_degree <= 1 Then
              m_poi(p1%).data(0).degree_for_reduce = 2
           Else
              m_poi(p1%).data(0).degree_for_reduce = 1
           End If
           Call set_two_point_line_for_line(line_no%, record_0)
        ElseIf p1% > 0 Then
           'Call set_point_coordinate(p1%, in_coord, False)
                Call add_point_to_line(p1%, line_no%, 0, display, is_draw, 0, temp_record.record_data)
                   Call add_point_to_m_circle(p1%, circle_no%, record0, 0)
                    m_poi(p1%).data(0).parent.inter_type = new_inter_point_type(m_poi(p1%).data(0).parent.inter_type, inter_point_line_circle)
        End If
        End If
  End If
   If point_type = new_point_on_line_circle12 Then
       in_coord = t_coord1
       'm_poi(p1%).data(0).data0.coordinate = t_coord1
   ElseIf point_type = new_point_on_line_circle21 Then
       in_coord = t_coord2
       ' m_poi(p1%).data(0).data0.coordinate = t_coord2
   ElseIf point_type > 0 Then
       in_coord = t_coord1
       'm_poi(p1%).data(0).data0.coordinate = t_coord1
   End If
End Function
Function inter_point_line_circle1(p0 As POINTAPI, para_or_verti As Integer, line_data As line_data_type, _
                                     c As circle_data0_type, out_coord1 As POINTAPI, _
                                     out_p1%, out_coord2 As POINTAPI, out_p2%) As Integer
If inter_point_line_circle3(p0, _
  para_or_verti, m_poi(line_data.data0.poi(0)).data(0).data0.coordinate, _
       m_poi(line_data.data0.poi(1)).data(0).data0.coordinate, _
        c, out_coord1, out_p1%, out_coord2, out_p2%, 0, False) Then
   If out_coord1.X = 10000 And out_coord1.Y = 10000 And _
     out_coord2.X = 10000 And out_coord2.Y = 10000 Then
    inter_point_line_circle1 = 0
   Else
     inter_point_line_circle1 = 1
   End If
End If
End Function

Function inter_point_line_circle2(ByVal l%, ByVal p0%, ByVal c%, _
                           coord As POINTAPI, point_no%) As Boolean '圆与直线已交于一点
Dim t&, r&, p%, p1%, p2%
Dim c_p As POINTAPI
Dim k!
p1% = m_lin(l%).data(0).data0.poi(0)
p2% = m_lin(l%).data(0).data0.poi(1)
If p1% <> p0% Then
 p% = p1%
ElseIf p2% <> p0% Then
 p% = p2%
End If
c_p = m_Circ(c%).data(0).data0.c_coord
r& = (m_poi(p0%).data(0).data0.coordinate.X - m_poi(p%).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(p0%).data(0).data0.coordinate.Y - m_poi(p%).data(0).data0.coordinate.Y) ^ 2
t& = 2 * ((m_poi(p0%).data(0).data0.coordinate.X - m_poi(p%).data(0).data0.coordinate.X) * _
       (c_p.X - m_poi(p0%).data(0).data0.coordinate.X) + _
        (m_poi(p0%).data(0).data0.coordinate.Y - m_poi(p%).data(0).data0.coordinate.Y) * _
           (c_p.Y - m_poi(p0%).data(0).data0.coordinate.Y))
k! = t& / r&
coord.X = m_poi(p0%).data(0).data0.coordinate.X + _
   CInt(k! * (m_poi(p0%).data(0).data0.coordinate.X - m_poi(p%).data(0).data0.coordinate.X))
coord.Y = m_poi(p0%).data(0).data0.coordinate.Y + _
    CInt(k! * (m_poi(p0%).data(0).data0.coordinate.Y - m_poi(p%).data(0).data0.coordinate.Y))
If point_no% > 0 Then
 Call set_point_coordinate(point_no%, coord, True)
End If
End Function
Sub remove_point(ByVal p%, dis As Boolean, ty As Byte) 'As Integer
Dim i%, j%, k%, l%, m%, o% 'ty=0 作图过程,1 更改数据
Dim t_line%, t_circle%, t_point% '因为过程传递的变量地址，
Dim ch$
'If p% > last_conditions.last_cond(1).point_no Then
' Exit Sub
'End If
ch$ = m_poi(p%).data(0).data0.name
If p% = 0 Then
 Exit Sub
End If
t_point% = p%
 For i% = 1 To C_display_picture.m_circle.Count '***
  If m_Circ(i%).data(0).data0.center = t_point% Then
   If m_Circ(i%).data(0).data0.in_point(0) > 2 Then
    Exit Sub '圆心不能先消除
  Else
   Call C_display_picture.remove_circle(i%)
    GoTo remove_point_mark0
  End If
 End If
Next i%
remove_point_mark0:
If m_poi(t_point%).data(0).data0.visible > 0 And dis = display Then
 Call set_point_visible(t_point%, 0, True)
 'Call draw_point(Draw_form, poi(t_point%), 0, delete)   '消点的像
End If
  i% = 1
 Do While i% <= last_conditions.last_cond(1).line_no '***
  If remove_point_from_lin(t_point%, i%, dis, ty) = False Then
    i% = i% + 1
  End If
 Loop
' i% = 1
For i% = 1 To C_display_picture.m_circle.Count
  Call remove_point_from_Circ(t_point%, i%) ', o%)
Next i%
Call remove_point_from_relation(t_point%)
Call remove_point_from_point_pair(t_point%)
Call remove_point_from_eline(t_point%)
'++++++++++++++++++++++++++++++++++++++++++++++++++++
 If t_point% <= last_conditions.last_cond(1).point_no Then
  For i% = t_point% To last_conditions.last_cond(1).point_no - 1
  Call move_point_data(i% + 1, i%)
 Next i%
 last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no - 1
 End If 'Else
 '*******
 For i% = 1 To C_display_wenti.m_last_input_wenti_no
  Call remove_point_from_senten(t_point%, i%)
 Next i%
 For i% = 0 To 7
 If temp_last_point(i%) > t_point% Then
 
  temp_last_point(i%) = temp_last_point(i%) - 1
  ElseIf temp_last_point(i%) = t_point% Then
  temp_last_point(i%) = 0
 End If
 Next i%

 For i% = 0 To 15
  If temp_point(i%).no < 95 Then
  If temp_point(i%).no > t_point% Then
   temp_point(i%).no = temp_point(i%).no - 1
  ElseIf temp_point(i%).no = t_point% Then
   temp_point(i%).no = 0
  End If
  End If
Next i%
i% = 1
While i% <= last_conditions.last_cond(1).line_no
If m_lin(i%).data(0).data0.in_point(0) < 2 Then
Call delete_line(i%)
End If
i% = i% + 1
Wend
If t_point% < 95 Then
For i% = 1 To last_conditions.last_cond(1).line_no
 For j% = 1 To m_lin(i%).data(0).data0.in_point(0)
  'If m_lin(i%).data(0).data0.in_point(j%) > t_point% Then
  ' lin(i%).data(0).data0.in_point(j%) = lin(i%).data(0).data0.in_point(j%) - 1
  'End If
 Next j%
 'lin(i%).data(0).data0.poi(0) = lin(i%).data(0).data0.in_point(1)
 'lin(i%).data(0).data0.poi(1) = lin(i%).data(0).data0.in_point(lin(i%).data(0).data0.in_point(0))
 Next i%
End If
If ch$ <> "" And ch$ <> empty_char Then
 Call delete_name(ch$)
End If
For i% = 1 To last_conditions.last_cond(1).verti_no
 If Dverti(i%).data(0).inter_poi = p% Then
     Dverti(i%).data(0).inter_poi = 0
 End If
Next i%
Call C_display_picture.delete_point(p%)
'Call remove_point_from_wenti(t_point%)
'Call display_input_again

End Sub

Private Sub remove_point_from_Circ(ByVal p%, ByVal c%) ', o%)
'消p%改号
Dim i%, j%, k%
Dim t_point%, t_circle%
t_point% = p%
t_circle% = c%
'*****************************************************
'圆上消去点
If p% = 0 Then
 Exit Sub
End If
k% = 0
For i% = 1 To m_Circ(t_circle%).data(0).data0.in_point(0)
 If m_Circ(t_circle%).data(0).data0.in_point(i%) = t_point% Then
  k% = i%
      GoTo remove_point_from_circle_mark0
 End If
Next i%
GoTo remove_point_from_circle_mark1
 'End If
remove_point_from_circle_mark0:
 For i% = k% To m_Circ(t_circle%).data(0).data0.in_point(0) - 1
'   Call set_circle_in_point(t_circle%, i%, m_Circ(t_circle%).data(0).data0.in_point(i% + 1), condition)
 Next i%
'   Call set_circle_in_point(t_circle%, 0, m_Circ(t_circle%).data(0).data0.in_point(0) - 1, condition)
remove_point_from_circle_mark1:
For i% = 1 To m_Circ(t_circle%).data(0).data0.in_point(0)
 If m_Circ(t_circle%).data(0).data0.in_point(i%) > t_point% Then
 ' Call set_circle_in_point(t_circle%, i%, m_Circ(t_circle%).data(0).data0.in_point(i%) - 1, condition)
 End If
Next i%


'******************************************************
'*******************************************************

'圆上只有一个点圆心是Ｐ或圆上没有点
End Sub

Private Function remove_point_from_lin(ByVal p%, ByVal l%, ByVal dis As Boolean, ty As Byte) As Boolean
'　消点不改号
Dim i%, j%, k%, r!, n%
Dim t_point%, t_line%
t_point% = p%
t_line% = l%
'n% = i%
For i% = 1 To m_lin(t_line%).data(0).data0.in_point(0)
 If m_lin(t_line).data(0).data0.in_point(i%) = t_point% Then
  n% = i%
   GoTo remove_point_form_line_mark0
 End If
Next i%
GoTo remove_point_from_line_mark1
remove_point_form_line_mark0:
For i% = n% + 1 To m_lin(t_line%).data(0).data0.in_point(0)
 m_lin(t_line%).data(0).data0.in_point(i% - 1) = _
 m_lin(t_line%).data(0).data0.in_point(i%)
Next i%
 m_lin(t_line%).data(0).data0.in_point(0) = m_lin(t_line%).data(0).data0.in_point(0) - 1
remove_point_from_line_mark1:
If m_lin(t_line%).data(0).data0.in_point(0) < 2 Then
Call delete_line(t_line%)
last_conditions.last_cond(1).line_no = last_conditions.last_cond(1).line_no - 1
For i% = l% To last_conditions.last_cond(1).line_no
  m_lin(i%) = m_lin(i% + 1)
     Call C_display_picture.draw_line(i% + 1, 0, 0)
Next i%
    Call C_display_picture.delete_line(i%)
     remove_point_from_lin = True
Else
 m_lin(t_line%).data(0).data0.poi(0) = m_lin(t_line%).data(0).data0.in_point(1)
  m_lin(t_line%).data(0).data0.poi(1) = m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0))
m_lin(t_line%).data(0).is_change = 255
Call C_display_picture.set_m_line_data0(t_line%, 0, 0)
End If
End Function
 
Private Sub remove_point_from_relation(ByVal p%)
Dim i%, j%, k%
i% = 1
While i% <= last_conditions.last_cond(1).relation_no
  If Drelation(i%).data(0).data0.poi(0) = p% Or Drelation(i%).data(0).data0.poi(1) = p% Or _
      Drelation(i%).data(0).data0.poi(1) = p% Or Drelation(i%).data(0).data0.poi(2) = p% Then
   last_conditions.last_cond(1).relation_no = last_conditions.last_cond(1).relation_no - 1
   For j% = 0 To last_conditions.last_cond(1).relation_no
   Drelation(j%) = Drelation(j% + 1)
   Next j%
  Else
   i% = i% + 1
  End If
Wend
End Sub

Function two_time_area_triangle(ByVal x1&, ByVal y1&, ByVal x2&, _
     ByVal y2&, ByVal X3&, ByVal Y3&) As Long
Dim s%
two_time_area_triangle = x1& * y2& - x2& * y1& + _
   x2& * Y3& - X3& * y2& + X3& * y1& - x1& * Y3&
End Function

Function two_before_point_in_line(ByVal p%, ByVal l%, p1%, p2%) As Boolean
Dim i%, n%
n% = 0
For i% = 1 To m_lin(l%).data(0).data0.in_point(0)
If m_lin(l%).data(0).data0.in_point(i%) < p% Then
n% = n% + 1
 If n% = 1 Then
  p1% = m_lin(l%).data(0).data0.in_point(i%)
 ElseIf i% = 2 Then
  p2% = m_lin(l%).data(0).data0.in_point(i%)
   two_before_point_in_line = True
   Exit Function
 End If
End If
Next i%
two_before_point_in_line = False
End Function

Private Sub remove_point_from_senten(ByVal p%, ByVal s%)
 Dim i%
 Dim tp%
 tp% = p%
  For i% = 0 To 50
   If C_display_wenti.m_condition(s%, i%) >= "A" And _
        C_display_wenti.m_condition(s%, i%) <= "Z" And _
         C_display_wenti.m_point_no(s%, i%) > tp% Then
        Call C_display_wenti.set_m_point_no(s%, _
           C_display_wenti.m_point_no(s%, i%) - 1, i%, True)
   End If
  Next i%
End Sub
Public Function inter_point_two_dline(l1 As line_type, l2 As line_type) As Integer
Dim i%, j%
For i% = i To l1.data(0).data0.in_point(0)
 For j% = 1 To l2.data(0).data0.in_point(0)
  If l1.data(0).data0.in_point(i%) = l2.data(0).data0.in_point(j%) Then
   inter_point_two_dline = l1.data(0).data0.in_point(i%)
    Exit Function
   End If
 Next j%
 Next i%
End Function

Public Function inter_point_line_circle3(p1_coord As POINTAPI, paral_or_verti_ As Integer, _
            p2_coord As POINTAPI, p3_coord As POINTAPI, cir As circle_data0_type, _
              out_coord1 As POINTAPI, out_p1%, out_coord2 As POINTAPI, out_p2%, _
                ty As Integer, is_change As Boolean, Optional point_type As Byte = 0) As Boolean
 '过一点平行（垂直）的直线交圆
Dim i%, j%, k%, n%, s!, A&, r&, b!, b1!, b2!, c!, d!, radii!
Dim c_coord As POINTAPI
Dim Cp_coord As POINTAPI
'*********
  If cir.center > 0 Then
      c_coord = m_poi(cir.center).data(0).data0.coordinate
  Else
      c_coord = cir.c_coord
  End If
  If cir.radii = 0 And cir.in_point(0) > 0 Then
      Cp_coord = m_poi(cir.in_point(1)).data(0).data0.coordinate
       radii = (Cp_coord.X - c_coord.X) ^ 2 + (Cp_coord.Y - c_coord.Y) ^ 2
  Else
   radii = cir.radii ^ 2
  End If
       A& = (p3_coord.X - p2_coord.X) ^ 2 + (p3_coord.Y - p2_coord.Y) ^ 2
       If A& = 0 Then
          Exit Function
       End If
         r& = sqr(A&) 'p2p3的长
 '************************************************
      b1! = CSng((p1_coord.X - c_coord.X) * (p3_coord.X - p2_coord.X) + _
             (p1_coord.Y - c_coord.Y) * (p3_coord.Y - p2_coord.Y)) / A&
      b2! = CSng((p1_coord.X - c_coord.X) * (p3_coord.Y - p2_coord.Y) - _
             (p1_coord.Y - c_coord.Y) * (p3_coord.X - p2_coord.X)) / A&
 If ty < 3 Then
  If paral_or_verti_ = paral_ Then
   b! = b1!
    b1! = b2!
     'b2! = b!
  ElseIf paral_or_verti_ = verti_ Then
    b! = b2!
  End If
   If Abs(b1! * r&) < 4 Then '过圆心
    s! = cir.radii / r&
    If paral_or_verti_ = paral_ Then
    out_coord1.X = c_coord.X + (p3_coord.X - p2_coord.X) * s!
    out_coord1.Y = c_coord.Y + (p3_coord.Y - p2_coord.Y) * s!
    out_coord2.X = c_coord.X - (p3_coord.X - p2_coord.X) * s!
    out_coord2.Y = c_coord.Y - (p3_coord.Y - p2_coord.Y) * s!
    ElseIf paral_or_verti_ = verti_ Then
    out_coord1.X = c_coord.X + (p3_coord.Y - p2_coord.Y) * s!
    out_coord1.Y = c_coord.Y - (p3_coord.X - p2_coord.X) * s!
    out_coord2.X = c_coord.X - (p3_coord.Y - p2_coord.Y) * s!
    out_coord2.Y = c_coord.Y + (p3_coord.X - p2_coord.X) * s!

    End If
   Else
  '***************************************************
       c! = CSng((c_coord.X - p1_coord.X) ^ 2 + (c_coord.Y - p1_coord.Y) ^ 2 - _
          radii!) / A&
          
        d! = b! ^ 2 - c!
         If d! < 0 Then
          out_coord1.X = 10000
          out_coord1.Y = 10000
          out_coord2.X = 10000
          out_coord2.Y = 10000
           inter_point_line_circle3 = False
              Exit Function
         'ElseIf D! >= -5 And D! <= 0 Then
         'D! = 0
         End If
         d! = sqr(d!)
If paral_or_verti_ = paral_ Then
   s! = -b! + d!
    out_coord1.X = p1_coord.X + (p3_coord.X - p2_coord.X) * s!
    out_coord1.Y = p1_coord.Y + (p3_coord.Y - p2_coord.Y) * s!
   s! = -b! - d!
    out_coord2.X = p1_coord.X + (p3_coord.X - p2_coord.X) * s!
    out_coord2.Y = p1_coord.Y + (p3_coord.Y - p2_coord.Y) * s!
ElseIf paral_or_verti_ = verti_ Then
   s! = -b! + d!
    out_coord1.X = p1_coord.X + (p3_coord.Y - p2_coord.Y) * s!
    out_coord1.Y = p1_coord.Y - (p3_coord.X - p2_coord.X) * s!
   s! = -b! - d!
    out_coord2.X = p1_coord.X + (p3_coord.Y - p2_coord.Y) * s!
    out_coord2.Y = p1_coord.Y - (p3_coord.X - p2_coord.X) * s!
End If
End If
Else 'ty=2
 If paral_or_verti_ = paral_ Then
    b1 = 2 * b1
    out_coord1.X = p1_coord.X - (p3_coord.X - p2_coord.X) * b1
    out_coord1.Y = p1_coord.Y - (p3_coord.Y - p2_coord.Y) * b1
 ElseIf paral_or_verti_ = verti_ Then
    b2 = 2 * b2
    out_coord1.X = p1_coord.X - (p3_coord.Y - p2_coord.Y) * b2
    out_coord1.Y = p1_coord.Y + (p3_coord.X - p2_coord.X) * b2
 End If
End If
inter_point_line_circle3 = True
If out_p1% > 0 Then
   Call set_point_coordinate(out_p1%, out_coord1, is_change)
End If
If out_p2% > 0 Then
   Call set_point_coordinate(out_p2%, out_coord2, is_change)
End If
End Function

Public Sub exchange_two_int(x1%, x2%)
Dim k%
k% = x1%
 x1% = x2%
  x2% = k%
End Sub

Public Function length_of_line(l As length_type) As Single
l.len = sqr((m_poi(l.poi(0)).data(0).data0.coordinate.X - m_poi(l.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
   (m_poi(l.poi(0)).data(0).data0.coordinate.Y - m_poi(l.poi(1)).data(0).data0.coordinate.Y) ^ 2)
End Function

Public Function C_Area_polygon(Polyg As polygon)
Dim i As Byte
C_Area_polygon = 0
If Polyg.total_v < 3 Then
 Exit Function
End If
For i = 1 To Polyg.total_v - 2
 C_Area_polygon = C_Area_polygon + _
   area_triangle(m_poi(Polyg.v(0)).data(0).data0.coordinate, _
     m_poi(Polyg.v(i)).data(0).data0.coordinate, m_poi(Polyg.v(i + 1)).data(0).data0.coordinate)
Next i
C_Area_polygon = Abs(C_Area_polygon)
End Function
Public Sub floodfill_polygon(p As polygon)
Dim i%, X&, Y&
If p.total_v > 2 Then
 Draw_form.Line (m_poi(p.v(0)).data(0).data0.coordinate.X, _
   m_poi(p.v(0)).data(0).data0.coordinate.Y)-( _
    m_poi(p.v(p.total_v - 1)).data(0).data0.coordinate.X, _
      m_poi(p.v(p.total_v - 1)).data(0).data0.coordinate.Y), QBColor(14)
For i% = 0 To p.total_v - 1
X& = X& + m_poi(p.v(i%)).data(0).data0.coordinate.X
Y& = Y& + m_poi(p.v(i%)).data(0).data0.coordinate.Y
Next i%
X& = X& / p.total_v
Y& = Y& / p.total_v
'Call FloodFill(Draw_form.hdc, X&, Y&, QBColor(12))
End If
End Sub

Public Function area_triangle(p1 As POINTAPI, _
   p2 As POINTAPI, p3 As POINTAPI) As Long
Dim x1&, y1&, x2&, y2&, X3&, Y3&
x1& = p1.X
x2& = p2.X
X3& = p3.X
y1& = p1.Y
y2& = p2.Y
Y3& = p3.Y

area_triangle = (x1& * y2& + x2& * Y3& + X3& * y1& - _
     y1& * x2& - y2& * X3& - Y3& * x1&) / 2
End Function

Public Function distance_point_point(p1 As POINTAPI, _
  p2 As POINTAPI, sq As Long) As Single
sq = (p1.X - p2.X) ^ 2 + _
    (p1.Y - p2.Y) ^ 2
distance_point_point = sqr(sq)
End Function
Public Function squre_distance_point_point(p1 As POINTAPI, _
  p2 As POINTAPI) As Single
squre_distance_point_point = (p1.X - p2.X) ^ 2 + _
    (p1.Y - p2.Y) ^ 2
End Function

Public Function value_of_angle(A As angle_value_for_measur_type) As Boolean
Dim sq(2) As Long
Dim s(2) As Single
Dim a_v As Single
Dim v(1) As Integer
Dim X As Single
If A.angle = 0 Then
 A.angle = Abs(angle_number(A.poi(0), A.poi(1), A.poi(2), 0, 0))
End If
If A.eangle > 0 Then
 A.value = angle_value_for_measur(A.eangle).value
Else 'End If

s(0) = distance_point_point(m_poi(A.poi(0)).data(0).data0.coordinate, _
    m_poi(A.poi(1)).data(0).data0.coordinate, sq(0))
s(1) = distance_point_point(m_poi(A.poi(1)).data(0).data0.coordinate, _
    m_poi(A.poi(2)).data(0).data0.coordinate, sq(1))
s(2) = distance_point_point(m_poi(A.poi(2)).data(0).data0.coordinate, _
   m_poi(A.poi(0)).data(0).data0.coordinate, sq(2))
If s(0) = 0 Or s(1) = 0 Then
 A.value = ""
  value_of_angle = False
 Exit Function
Else
 X = (sq(0) + sq(1) - sq(2)) / 2 / s(0) / s(1)
  If X <> 0 And Abs(X) <> 1 Then
   a_v = Atn(-X / sqr(-X * X + 1)) + PI / 2 '2 * Atn(1#) '
    'If X < 0 Then
    'a_v = Pi - a_v
    'End If
   ElseIf X = 1 Then
   a_v = 0
   ElseIf X = -1 Then
   a_v = PI
   Else
    a_v = PI / 2
   End If
  a_v = a_v * 180 / PI
  v(0) = Int(a_v)
   a_v = (a_v - v(0)) * 60
    v(1) = Int(a_v)
A.value = LoadResString_(1675, "\\1\\" + str(v(0)) + "\\2\\" + str(v(1)))
 value_of_angle = True
End If
End If
End Function

Public Sub value_of_angle1(p1 As POINTAPI, p2 As POINTAPI, _
   p3 As POINTAPI, a_v As Single, si!, co!)
Dim sq(2) As Long
Dim s(2) As Long
Dim X As Single
Dim t As Boolean
Dim vf As POINTAPI
t = distance_point_to_line(p1, p3, paral_, p3, p2, s(0), vf, 1)
s(1) = distance_point_point(p1, p2, sq(1))
si! = Abs(s(0)) / s(1) 'sin
co! = sqr(1 - si! ^ 2) 'cos
If co! <> 0 Then
a_v = Atn(si! / co!) 'arctg
Else 'cos=0
a_v = PI / 2
End If

End Sub
Public Sub remove_point_from_wenti(ByVal p%, ByVal n%)
Dim j%
'从输入语句删除点
 For j% = 0 To 50
  If C_display_wenti.m_condition(n%, j%) <> "" And _
      C_display_wenti.m_condition(n%, j%) <> empty_char And _
       C_display_wenti.m_point_no(n%, j%) > p% Then
    Call C_display_wenti.set_m_point_no(n%, _
        C_display_wenti.m_point_no(n%, j%) - 1, j%, True)
  End If
 Next j%
'Next i%
End Sub
Public Sub remove_init()
Dim i%
'初始化画图数据
For i% = 0 To 7
fill_color_line(i%) = 0
fill_color_line(i%) = 0
Next i%
For i% = 0 To 15
red_line(i%) = 0
Next i%
End Sub

Public Sub set_red_line(ByVal l%)
Dim i%
For i% = 0 To 15
If red_line(i%) = l% Then
Exit Sub
End If
Next i%
For i% = 0 To 15
If red_line(i%) = 0 Then
red_line(i%) = l%
Exit Sub
End If
Next i%

End Sub

Public Sub set_fill_color_line(ByVal l%)
Dim i%
For i% = 0 To 7
If fill_color_line(i%) = l% Then
Exit Sub
End If
Next i%
For i% = 0 To 7
If fill_color_line(i%) = 0 Then
fill_color_line(i%) = l%
Exit Sub
End If
Next i%
End Sub

Public Function inter_point_circle_circle0(cir1 As circle_data0_type, cir2 As circle_data0_type, p1%, p2%) As Integer
Dim i%, j%
p1% = 0
p2% = 0
For i% = 1 To cir1.in_point(0)
 For j% = 1 To cir2.in_point(0)
  If cir1.in_point(i%) = cir2.in_point(j%) Then
   If inter_point_circle_circle0 = 0 Then
       inter_point_circle_circle0 = 1
        p1% = cir1.in_point(i%)
   Else
       inter_point_circle_circle0 = 2
        p1% = cir1.in_point(i%)
         Exit Function
   End If
  End If
 Next j%
Next i%
 
End Function

Public Function inter_point_line_circle0(l As line_data0_type, ci As circle_data0_type, _
       p1%, p2%) As Integer
Dim i%, j%
p1% = 0
p2% = 0
For i% = 1 To l.in_point(0)
 For j% = 1 To ci.in_point(0)
  If l.in_point(i%) = ci.in_point(j%) Then
   If inter_point_line_circle0 = 0 Then
    p1% = l.in_point(i%)
     inter_point_line_circle0 = 1
   Else
    p1% = l.in_point(i%)
     inter_point_line_circle0 = 2
       Exit Function
   End If
  End If
 Next j%
Next i%
End Function

Public Function inter_point_circle_circle_by_pointapi(center1 As POINTAPI, radii1 As Long, _
                 center2 As POINTAPI, radii2 As Long, out_coord1 As POINTAPI, out_coord2 As POINTAPI, _
                    Optional inter_point_type As Integer = 0) As POINTAPI
Dim r!, sr!, temp_num0!
Dim s&
Dim temp_k1!, temp_k2!, temp_x!, temp_y!
Dim temp_coord As POINTAPI
'********************************************************************************
       s& = (center1.X - center2.X) ^ 2 + (center1.Y - center2.Y) ^ 2
       sr! = sqr(s&)
If sr! > radii1 + radii2 Or _
     sr! < Abs(radii1 - radii2) _
       Or sr! = 0 Then
  If sr! = 0 Then '同心圆
  out_coord1.X = 10000
  out_coord1.Y = 10000
  out_coord2.X = 10000
  out_coord2.Y = 10000
  ElseIf sr! > radii1 + radii2 Then '
  t_coord = minus_POINTAPI(center2, center1)
  out_coord1 = add_POINTAPI(center1, time_POINTAPI_by_number(t_coord, radii1 / sr!))
  out_coord2 = minus_POINTAPI(center2, time_POINTAPI_by_number(t_coord, radii2 / sr!))
  ElseIf sr! < Abs(radii1 - radii2) Then
   If radii1 - radii2 > 0 Then
      t_coord = minus_POINTAPI(center2, center1)
      out_coord1 = add_POINTAPI(center1, time_POINTAPI_by_number(t_coord, radii1 / sr!))
      out_coord2 = add_POINTAPI(center2, time_POINTAPI_by_number(t_coord, radii2 / sr!))
   Else
         t_coord = minus_POINTAPI(center1, center2)
      out_coord1 = add_POINTAPI(center1, time_POINTAPI_by_number(t_coord, radii1 / sr!))
      out_coord2 = add_POINTAPI(center2, time_POINTAPI_by_number(t_coord, radii2 / sr!))
   End If
  End If
  inter_point_circle_circle_by_pointapi.X = -10000   '无交点
 Exit Function
ElseIf (sr! = radii1 + radii2 Or _
     sr! = Abs(radii1 - radii2)) And sr! <> 0 Then
  out_coord1 = add_POINTAPI(center1, time_POINTAPI_by_number(minus_POINTAPI(center2, center1), radii1 / sr!))
  out_coord2 = add_POINTAPI(center2, time_POINTAPI_by_number(minus_POINTAPI(center1, center2), radii2 / sr!))
  'inter_point_circle_circle_by_pointapi = 1    '无交点
Else
   temp_num0! = (s& + radii1 ^ 2 - _
       radii2 ^ 2) / 2 / sr!
         's! = Sqr(s!)
          temp_k1! = (center2.X - center1.X) / sr!
           temp_k2! = (center2.Y - center1.Y) / sr!
          temp_x! = temp_k1! * temp_num0! + center1.X
          temp_y! = temp_k2! * temp_num0! + center1.Y
          temp_num0 = CLng((radii1 + radii2 + sr!) / 2)
                '会出负数
               r! = (temp_num0! - radii1) * _
                     (temp_num0! - radii2) * _
                       (temp_num0! - sr!) * temp_num0!
    If r! < 0 Then
     inter_point_circle_circle_by_pointapi.X = -10000
     Exit Function
    End If
                 temp_num0 = sqr(r!) * 2 / sr!
                   out_coord1.X = CInt(-temp_k2! * temp_num0! + temp_x!)
                   out_coord1.Y = CInt(temp_k1! * temp_num0! + temp_y!)
                   out_coord2.X = CInt(temp_k2! * temp_num0! + temp_x!)
                   out_coord2.Y = CInt(-temp_k1! * temp_num0! + temp_y!)
                  ' inter_point_circle_circle_by_pointapi = 2
   'If out_p1% > 0 Then
   '  Call set_point_coordinate(out_p1%, out_coord1, is_change)
   'End If
   'If out_p2% > 0 Then
   '  Call set_point_coordinate(out_p2%, out_coord2, is_change)
   'End If
End If
'*************************************************************************************
           If inter_point_type = new_point_on_circle_circle12 Then
            inter_point_circle_circle_by_pointapi = out_coord1
           ElseIf inter_point_type = new_point_on_circle_circle21 Then
            inter_point_circle_circle_by_pointapi = out_coord2
           End If
                   'out_coord2.Y = CInt(-temp_k1! * temp_num0! + temp_y!)


End Function

Public Sub init_data_base()
last_conditions.last_cond(1).aid_point_data1_no = 0
last_conditions.last_cond(1).aid_point_data2_no = 0
last_conditions.last_cond(1).aid_point_data3_no = 0
last_conditions.last_cond(1).angle_less_angle_no = 0
last_conditions.last_cond(1).angle_no = 0
last_conditions.last_cond(1).angle_relation_no = 0
last_conditions.last_cond(1).angle_value_90_no = 0
last_conditions.last_cond(1).angle_value_no = 0
last_conditions.last_cond(1).angle3_value_no = 0
last_conditions.last_cond(1).arc_no = 0
last_conditions.last_cond(1).arc_value_no = 0
last_conditions.last_cond(1).area_of_circle_no = 0
last_conditions.last_cond(1).area_of_element_no = 0
last_conditions.last_cond(1).area_of_fan_no = 0
last_conditions.last_cond(1).area_relation_no = 0
last_conditions.last_cond(1).branch_data_no = 0
last_conditions.last_cond(1).change_picture_step = 0
last_conditions.last_cond(1).change_picture_type = 0
last_conditions.last_cond(1).con_line_no = 0
last_conditions.last_cond(1).dangle_no = 0
last_conditions.last_cond(1).distance_of_paral_line_no = 0
last_conditions.last_cond(1).distance_of_point_line_no = 0
last_conditions.last_cond(1).dline1_no = 0
last_conditions.last_cond(1).dpoint_pair_no = 0
last_conditions.last_cond(1).eangle_no = 0
last_conditions.last_cond(1).eline_no = 0
last_conditions.last_cond(1).epolygon_no = 0
last_conditions.last_cond(1).equal_3angle_no = 0
last_conditions.last_cond(1).equal_arc_no = 0
last_conditions.last_cond(1).equal_side_right_triangle_no = 0
last_conditions.last_cond(1).equal_side_triangle_no = 0
last_conditions.last_cond(1).equation_no = 0
last_conditions.last_cond(1).four_point_on_circle_no = 0
last_conditions.last_cond(1).four_sides_fig_no = 0
last_conditions.last_cond(1).function_of_angle_no = 0
last_conditions.last_cond(1).general_angle_string_no = 0
last_conditions.last_cond(1).general_string_combine_no = 0
last_conditions.last_cond(1).general_string_no = 0
last_conditions.last_cond(1).init_v_line_no = 0
last_conditions.last_cond(1).item0_no = 0
last_conditions.last_cond(1).length_of_polygon_no = 0
last_conditions.last_cond(1).line_from_two_point_no = 0
last_conditions.last_cond(1).line_less_line_no = 0
last_conditions.last_cond(1).line_less_line2_no = 0
last_conditions.last_cond(1).line_value_no = 0
last_conditions.last_cond(1).line2_less_line2_no = 0
last_conditions.last_cond(1).line3_value_no = 0
last_conditions.last_cond(1).long_squre_no = 0


End Sub

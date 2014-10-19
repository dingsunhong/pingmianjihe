Attribute VB_Name = "circle_module"
Option Explicit

Global last_m_circle As Integer
Global m_Circ() As circle_type
Global m_input_circle_data0 As circle_data_type
Global m_output_circle_data0 As circle_data_type
Public Sub set_point_inform(ByVal p%, inf$)
If m_poi(p%).data(0).inform = "" Or m_poi(p%).data(0).inform <> "自由点" Then
   m_poi(p%).data(0).inform = inf$
End If
End Sub
Public Function m_circle_number0(ByVal center%, ByVal radii&, ByVal p1%, ByVal p2%, ByVal p3%, _
                                     c_data0 As circle_data_type) As Boolean '判断满足条件的圆是否是c_data0，是=true
 Dim i%, j%, last_p%
 Dim tp(1 To 3) As Integer
' If p1% = 0 And p2% = 0 And p3% = 0 Then
'    m_circle_number0 = False
  '    Exit Function
 'End If
  If p1% > 0 Then
     last_p% = last_p% + 1
     tp(last_p%) = p1%
  End If
  If p2% > 0 Then
     last_p% = last_p% + 1
     tp(last_p%) = p2%
  End If
  If p3% > 0 Then
     last_p% = last_p% + 1
     tp(last_p%) = p3%
  End If
 If center% = 0 Then '无圆心，三点确定圆
   For i% = 1 To last_p%
    For j% = 1 To c_data0.data0.in_point(0)
         If c_data0.data0.in_point(j%) = tp(i%) Then
              GoTo m_circle0_next '一点重合
         End If
    Next j%
  Exit Function '无重合点
m_circle0_next:
  Next i%
  m_circle_number0 = True
ElseIf center% > 0 Then '有圆心
     If center% = c_data0.data0.center Then '圆心同
       If tp(1) = 0 And Abs(radii& - c_data0.data0.radii) < 5 Then '第一点空，半径同
        m_circle_number0 = True
       Else 'If tp(1) > 0 Then
        For j% = 1 To c_data0.data0.in_point(0)
         If c_data0.data0.in_point(j%) = tp(1) Then
              m_circle_number0 = True
               Exit Function
         End If
        Next j%
                           
       End If
     End If
End If
End Function
Public Function m_circle_number(ByVal start%, ByVal center As Integer, cen_coord As POINTAPI, _
                                  ByVal p1%, ByVal p2%, ByVal p3%, ByVal r&, _
                                    ByVal dp1%, ByVal dp2%, ByVal d_para As Single, _
                                     visible As Byte, ty As Byte, color As Byte, is_set As Boolean) As Integer
Dim i%, j%, no%
Dim temp_record As total_record_type
Dim tc_data As circle_data_type
'*******************************************************************
'判断是否是已有记录
m_circle_number = is_old_m_circle(start%, center, r&, p1%, p2%, p3%)
If m_circle_number > 0 Then
    Exit Function
Else
'*****************************************************************
  If center > 0 Then
      tc_data.data0.c_coord = m_poi(center).data(0).data0.coordinate
  Else
      tc_data.data0.c_coord = cen_coord
  End If
  tc_data.data0.center = center
  tc_data.data0.color = color
  tc_data.data0.visible = visible
  tc_data.data0.radii = r&
  tc_data.input_type = ty
  If p1% > 0 And p2% > 0 And p3% > 0 Then
    tc_data.data0.in_point(0) = 3
    tc_data.data0.in_point(1) = p1%
    tc_data.data0.in_point(2) = p2%
    tc_data.data0.in_point(3) = p3%
    tc_data.circle_type = 2
  ElseIf p1% > 0 And p2% > 0 Then
    tc_data.data0.in_point(0) = 2
    tc_data.data0.in_point(1) = p1%
    tc_data.data0.in_point(2) = p2%
    tc_data.circle_type = 1
  ElseIf p2% > 0 And p3% > 0 Then
    tc_data.data0.in_point(0) = 2
    tc_data.data0.in_point(1) = p2%
    tc_data.data0.in_point(2) = p3%
    tc_data.circle_type = 1
  ElseIf p1% > 0 And p3% > 0 Then
    tc_data.data0.in_point(0) = 2
    tc_data.data0.in_point(1) = p1%
    tc_data.data0.in_point(2) = p3%
    tc_data.circle_type = 1
  ElseIf p1% > 0 Then
    tc_data.data0.in_point(0) = 1
    tc_data.data0.in_point(1) = p1%
  ElseIf p2% > 0 Then
    tc_data.data0.in_point(0) = 1
    tc_data.data0.in_point(1) = p2%
  ElseIf p3 > 0 Then
    tc_data.data0.in_point(0) = 1
    tc_data.data0.in_point(1) = p3%
  End If
    tc_data.radii_depend_poi(0) = dp1%
    tc_data.radii_depend_poi(1) = dp2%
    tc_data.depend_para = d_para
       m_circle_number = Set_m_circle_data(0, tc_data)
       If m_Circ(m_circle_number).data(0).data0.center > 0 Then
          Call set_parent(point_, m_Circ(m_circle_number).data(0).data0.center, circle_, m_circle_number, 0)
         'If m_Circ(m_circle_number).data(0).data0.in_point(1) > 0 Then
          Call set_parent(point_, m_Circ(m_circle_number).data(0).data0.in_point(1), circle_, m_circle_number, 0)
         'End If
       Else
         Call set_parent(point_, m_Circ(m_circle_number).data(0).data0.in_point(1), circle_, m_circle_number, 0)
         Call set_parent(point_, m_Circ(m_circle_number).data(0).data0.in_point(2), circle_, m_circle_number, 0)
         Call set_parent(point_, m_Circ(m_circle_number).data(0).data0.in_point(3), circle_, m_circle_number, 0)
       End If
    'If is_set Then
     '        Call set_wenti_cond8_71(m_circle_number, 0)
    'End If
  End If
End Function
Public Function m_circle_radii(c_data0 As circle_data_type) As Long
   If c_data0.radii_depend_poi(0) > 0 And c_data0.radii_depend_poi(1) > 0 Then
        m_circle_radii = abs_POINTAPI(minus_POINTAPI(m_poi(c_data0.radii_depend_poi(0)).data(0).data0.coordinate, _
                                   m_poi(c_data0.radii_depend_poi(1)).data(0).data0.coordinate))
        m_circle_radii = m_circle_radii * c_data0.depend_para
   Else
    If c_data0.circle_type = 3 Then
      m_circle_radii = circle_radii0(m_poi(c_data0.data0.in_point(1)).data(0).data0.coordinate, _
                                 m_poi(c_data0.data0.in_point(2)).data(0).data0.coordinate, _
                                 m_poi(c_data0.data0.in_point(3)).data(0).data0.coordinate, _
                                  c_data0.data0.c_coord)
      Call next_char(c_data0.data0.in_point(2), "", 0, 0)
      Call next_char(c_data0.data0.in_point(3), "", 0, 0)
     ' If c_data0.data0.center = 0 Then
      '   c_data0.data0.center = set_point(c_data0.data0.c_coord, 1, condition_color)
      'Else
      '   Call set_point_coordinate(c_data0.data0.center, c_data0.data0.c_coord, True)
      'End If
    ElseIf c_data0.circle_type = 2 Then
         m_circle_radii = abs_POINTAPI(minus_POINTAPI(m_poi(c_data0.data0.in_point(1)).data(0).data0.coordinate, _
                                   m_poi(c_data0.data0.in_point(2)).data(0).data0.coordinate)) / 2
         c_data0.data0.c_coord = divide_POINTAPI_by_number(add_POINTAPI(m_poi(c_data0.data0.in_point(1)).data(0).data0.coordinate, _
                                   m_poi(c_data0.data0.in_point(2)).data(0).data0.coordinate), 2)
         Call next_char(c_data0.data0.in_point(2), "", 0, 0)
         If c_data0.data0.center >= 0 Then
            Call set_point_coordinate(c_data0.data0.center, c_data0.data0.c_coord, True)
         End If
    ElseIf c_data0.circle_type = 1 Then
      If c_data0.data0.in_point(1) > 0 Then
       m_circle_radii = abs_POINTAPI(minus_POINTAPI(m_poi(c_data0.data0.center).data(0).data0.coordinate, _
                                   m_poi(c_data0.data0.in_point(1)).data(0).data0.coordinate))
      Else
       m_circle_radii = c_data0.data0.radii
      End If
    End If
   End If
End Function
Public Function circle_radii0(p_0 As POINTAPI, p_1 As POINTAPI, p_2 As POINTAPI, c_coord As POINTAPI) As Long
 Dim p_coord(2) As POINTAPI
Dim tn%
'On Error GoTo read_three_circle0_error
If p_2.X = -30000 And p_2.Y = -30000 Then
   If p_0.X = p_1.X And p_0.Y = p_1.Y Then
   c_coord = p_0
   circle_radii0 = 0
   Else
   c_coord = divide_POINTAPI_by_number(add_POINTAPI(p_0, p_1), 2)
   circle_radii0 = abs_POINTAPI(minus_POINTAPI(p_0, p_1)) / 2
   End If
Else
If Abs(p_0.X - p_1.X) + Abs(p_0.Y - p_1.Y) < 5 Then
   p_coord(0) = p_0
   p_coord(1) = p_2
   p_coord(2).X = -30000
   p_coord(2).Y = -30000
   circle_radii0 = circle_radii0(p_coord(0), p_coord(1), p_coord(2), c_coord)
ElseIf Abs(p_0.X - p_2.X) + Abs(p_0.Y - p_2.Y) < 5 Then
   p_coord(0) = p_0
   p_coord(1) = p_1
   p_coord(2).X = -30000
   p_coord(2).Y = -30000
   circle_radii0 = circle_radii0(p_coord(0), p_coord(1), p_coord(2), c_coord)
ElseIf Abs(p_1.X - p_2.X) + Abs(p_1.Y - p_2.Y) < 5 Then
   p_coord(0) = p_0
   p_coord(1) = p_1
   p_coord(2).X = -30000
   p_coord(2).Y = -30000
   circle_radii0 = circle_radii0(p_coord(0), p_coord(1), p_coord(2), c_coord)
Else
p_coord(0) = mid_POINTAPI(p_0, p_1) '中点
p_coord(1) = mid_POINTAPI(p_1, p_2)
Call inter_point_line_line2(p_coord(0), verti_, p_0, p_1, _
            p_coord(1), verti_, p_1, p_2, c_coord, 0, True, False)  '外心
 If c_coord.X > 30000 Then
  c_coord.X = 30000
  c_coord.Y = (c_coord.Y / c_coord.X) * 30000
 ElseIf c_coord.Y > 30000 Then
  c_coord.Y = 30000
  c_coord.X = (c_coord.X / c_coord.Y) * 30000
 ElseIf c_coord.X < -30000 Then
  c_coord.X = -30000
  c_coord.Y = (c_coord.Y / c_coord.X) * -30000
 ElseIf c_coord.Y < -30000 Then
  c_coord.Y = -30000
  c_coord.X = (c_coord.X / c_coord.Y) * -30000
 End If
 circle_radii0 = distance_of_two_POINTAPI(c_coord, p_0) '半径
 End If
End If
End Function
Public Function read_three_circle0(circle_no%, p_0 As POINTAPI, p_1 As POINTAPI, _
             p_2 As POINTAPI, c_coord As POINTAPI, r&, last_point%) As Boolean
             'last_point%>0 第三点已定
Dim A&, b&, c&, d&, p&, q&, delta&
Dim p_coord(1) As POINTAPI
Dim tn%
'On Error GoTo read_three_circle0_error
p_coord(0) = mid_POINTAPI(p_0, p_1) '中点
p_coord(1) = mid_POINTAPI(p_1, p_2)
Call inter_point_line_line2(p_coord(0), verti_, p_0, p_1, _
            p_coord(1), verti_, p_1, p_2, c_coord, 0, True, False) '外心
 If c_coord.X > 30000 Then
  c_coord.X = 30000
  c_coord.Y = (c_coord.Y / c_coord.X) * 30000
 ElseIf c_coord.Y > 30000 Then
  c_coord.Y = 30000
  c_coord.X = (c_coord.X / c_coord.Y) * 30000
 ElseIf c_coord.X < -30000 Then
  c_coord.X = -30000
  c_coord.Y = (c_coord.Y / c_coord.X) * -30000
 ElseIf c_coord.Y < -30000 Then
  c_coord.Y = -30000
  c_coord.X = (c_coord.X / c_coord.Y) * -30000
 End If
 r& = distance_of_two_POINTAPI(c_coord, p_0) '半径
 read_three_circle0 = True '
   If circle_no% > 0 Then
    If m_Circ(circle_no%).data(0).data0.center > 0 Then
     Call set_point_coordinate(m_Circ(circle_no%).data(0).data0.center, c_coord, True)
      Call set_point_coordinate(m_Circ(circle_no%).data(0).data0.center, c_coord, True)
    End If
     m_Circ(circle_no%).data(0).data0.c_coord = c_coord
     m_Circ(circle_no%).data(0).data0.radii = r&
     m_input_circle_data0 = m_Circ(circle_no%).data(0)
     If last_point% > 0 Then
        m_input_circle_data0.in_point(0) = 3
        m_input_circle_data0.in_point(3) = last_point%
     End If
     m_input_circle_data0.is_change = True
     Call C_display_picture.set_m_circle_data0(circle_no%)
   'ElseIf circle_no% < 0 Then
     'm_aid_Circ(-circle_no%).data(0).data0.c_coord = c_coord
     'm_aid_Circ(-circle_no%).data(0).data0.radii = r&
     'm_circle_data0 = m_aid_Circ(-circle_no%).data(0)
     'm_circle_data0.is_change = True
     'Call C_display_picture.set_m_circle_data0(circle_no%, aid_condition)
  End If
read_three_circle0_error:
End Function
Function circumcenter(p1%, p2%, p3%, p%) As Integer
Dim tcoord As POINTAPI
Call m_circle_number(1, p1%, tcoord, p1%, p2%, p3%, 0, 0, 0, 1, 1, condition, condition_color, True)
Call draw_triangle(p1%, p2%, p3, condition)
End Function


Public Function set_circle_from_four_point(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%) As Integer
Dim i%
   i% = m_circle_number(1, 0, pointapi0, p1%, p2%, p3%, 0, 0, 0, 1, 0, condition, condition_color, False)
   If i% > 0 Then
        set_circle_from_four_point = m_circle_number(1, 0, pointapi0, p1%, p2%, p4%, _
                                                        0, 0, 0, 1, 0, condition, condition_color, False)
        If i% = set_circle_from_four_point Then
            Exit Function
        Else
         If set_circle_from_four_point > i% Then
            Call exchange_two_integer(set_circle_from_four_point, i%)
         End If
            Call combine_two_circle(set_circle_from_four_point, i%)
        End If
     Call add_point_to_m_circle(p4%, set_circle_from_four_point, record0, True)
     set_circle_from_four_point = i%
   Else
     set_circle_from_four_point = m_circle_number(1, 0, pointapi0, p1%, p2%, p4%, _
                                                        0, 0, 0, 1, 0, condition, condition_color, True)
     Call add_point_to_m_circle(p3%, set_circle_from_four_point, record0, True)
   End If
 End Function
Public Sub move_circle_inner_data(s As Byte, t As Byte)
Dim i%
For i% = 1 To last_conditions.last_cond(1).circle_no
  m_Circ(i%).data(t) = m_Circ(i%).data(s)
Next i%
End Sub
Public Function Set_m_circle_data(ByVal circle_no%, c_data0 As circle_data_type) As Integer
Dim i%
For i% = 1 To last_conditions.last_cond(1).circle_no
    If is_same_POINTAPI(m_Circ(i%).data(0).data0.c_coord, c_data0.data0.c_coord) And _
         Abs(m_Circ(i%).data(0).data0.radii - c_data0.data0.radii) < 5 Then
           circle_no% = i%
            Exit Function
    End If
Next i%
If last_conditions.last_cond(1).circle_no Mod 10 = 0 Or _
        circle_no% > last_conditions.last_cond(1).circle_no Then
ReDim Preserve m_Circ(last_conditions.last_cond(1).circle_no + 10) As circle_type
End If
If circle_no% = 0 Then
    For i% = 1 To last_conditions.last_cond(1).circle_no
     If m_Circ(i%).is_set_data = False Then
      Set_m_circle_data = i%
       GoTo set_m_circle_data_mark1
     End If
    Next i%
     last_conditions.last_cond(1).circle_no = _
      last_conditions.last_cond(1).circle_no + 1
     Set_m_circle_data = last_conditions.last_cond(1).circle_no
Else
     Set_m_circle_data = circle_no%
End If
set_m_circle_data_mark1:
      m_Circ(Set_m_circle_data).is_set_data = True
 If c_data0.data0.in_point(3) > 0 And (c_data0.data0.in_point(3) < c_data0.data0.center Or _
                     c_data0.data0.center = 0) Then
    c_data0.circle_type = 3
 ElseIf c_data0.data0.in_point(2) > 0 And (c_data0.data0.in_point(2) < c_data0.data0.center Or _
                     c_data0.data0.center = 0) Then
    c_data0.circle_type = 2
 ElseIf c_data0.data0.in_point(0) > 0 Then
    c_data0.circle_type = 1
 End If
   m_Circ(Set_m_circle_data).data(0) = c_data0

 '**************************************************************************************
    If m_Circ(Set_m_circle_data).data(0).input_type = 1 Then
     If m_Circ(Set_m_circle_data).data(0).circle_type = 1 Then
      m_Circ(Set_m_circle_data).data(0).parent.element(1).no = _
             m_Circ(Set_m_circle_data).data(0).data0.center
      m_Circ(Set_m_circle_data).data(0).parent.element(2).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(1)
      m_Circ(Set_m_circle_data).data(0).parent.element(1).ty = point_
      m_Circ(Set_m_circle_data).data(0).parent.element(2).ty = point_
    ElseIf m_Circ(Set_m_circle_data).data(0).circle_type = 2 Then
      m_Circ(Set_m_circle_data).data(0).parent.element(1).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(1)
      m_Circ(Set_m_circle_data).data(0).parent.element(2).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(2)
      m_Circ(Set_m_circle_data).data(0).parent.element(1).ty = point_
      m_Circ(Set_m_circle_data).data(0).parent.element(2).ty = point_
    ElseIf m_Circ(Set_m_circle_data).data(0).circle_type = 3 Then
     m_Circ(Set_m_circle_data).data(0).parent.element(1).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(1)
     m_Circ(Set_m_circle_data).data(0).parent.element(2).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(2)
     m_Circ(Set_m_circle_data).data(0).parent.element(3).no = _
             m_Circ(Set_m_circle_data).data(0).data0.in_point(3)
     m_Circ(Set_m_circle_data).data(0).parent.element(1).ty = point_
     m_Circ(Set_m_circle_data).data(0).parent.element(2).ty = point_
     m_Circ(Set_m_circle_data).data(0).parent.element(3).ty = point_
    End If
    'If m_Circ(Set_m_circle_data).data(0).data0.center > 0 Then
    '  Call set_son_data(circle_, Set_m_circle_data, m_Circ(Set_m_circle_data).data(0).sons, _
                        point_, c_data0.data0.center, m_poi(c_data0.data0.center).data(0).sons)
    'End If
    Call add_m_circle_to_point(Set_m_circle_data, c_data0.data0.in_point(1))
    Call add_m_circle_to_point(Set_m_circle_data, c_data0.data0.in_point(2))
    Call add_m_circle_to_point(Set_m_circle_data, c_data0.data0.in_point(3))
    Call set_circle_inform(Set_m_circle_data, "")
    End If
    If m_Circ(Set_m_circle_data).data(0).data0.radii = 0 Then
    m_Circ(Set_m_circle_data).data(0).data0.radii = _
        m_circle_radii(m_Circ(Set_m_circle_data).data(0))
    End If
    Call C_display_picture.set_m_circle_data0(Set_m_circle_data)
End Function

Public Function is_point_in_circle(ByVal circle_no%, ByVal c_p%, ByVal p0%, ByVal p1%, ByVal p2%) As Boolean
If circle_no > 0 And circle_no% <= last_conditions.last_cond(1).circle_no Then
 is_point_in_circle = m_circle_number0(c_p%, 0, p0%, p1%, p2%, m_Circ(circle_no%).data(0))
End If
End Function

Public Function is_two_circles_inter_two_point(ByVal c1%, ByVal c2%, _
   p1%, p2%) As Boolean
If is_point_in_circle(c1%, 0, p1%, p2%, 0) And _
     is_point_in_circle(c2%, 0, p1%, p2%, 0) Then
      is_two_circles_inter_two_point = True
End If
End Function
Public Function add_point_to_m_circle(n_point%, circle_no%, record As total_record_type, is_set_data As Boolean) As Byte
Dim i%, j%, k%, tl%
Dim ele_(1) As condition_type
If n_point% = 0 Or circle_no% = 0 Then
 Exit Function
End If
If n_point% <> 0 Then
If m_Circ(circle_no%).data(0).input_type = aid_condition Then
   If m_Circ(circle_no%).data(0).data0.in_point(0) = 0 Then '圆上无点
     m_Circ(circle_no%).data(0).data0.c_coord = divide_POINTAPI_by_number(add_POINTAPI(m_Circ(circle_no%).data(0).data0.c_coord, _
            m_poi(n_point%).data(0).data0.coordinate), 2) '圆心坐标
     m_Circ(circle_no%).data(0).data0.in_point(0) = 2
     add_point_to_m_circle = 2
     m_Circ(circle_no%).data(0).data0.in_point(2) = n_point%
     m_Circ(circle_no%).data(0).data0.in_point(1) = m_Circ(circle_no%).data(0).data0.center '设置圆上的点
     m_Circ(circle_no%).data(0).data0.radii = m_Circ(circle_no%).data(0).data0.radii / 2 '半径
     m_Circ(circle_no%).data(0).data0.center = 0 '圆心点的序号
     Call add_m_circle_to_point(circle_no%, m_Circ(circle_no%).data(0).data0.in_point(1))
     Call add_m_circle_to_point(circle_no%, m_Circ(circle_no%).data(0).data0.in_point(2))
      m_Circ(circle_no%).data(0).circle_type = 2
      Call C_display_picture.draw_circle(circle_no%, 0, 0) '显示
'******************************************************************************************************************
   ElseIf m_Circ(circle_no%).data(0).data0.in_point(0) = 2 Then '圆上有两点
     If n_point% <> m_Circ(circle_no%).data(0).data0.in_point(1) And _
                  n_point% <> m_Circ(circle_no%).data(0).data0.in_point(2) Then '新点不同于已有点
           m_Circ(circle_no%).data(0).data0.in_point(0) = 3
           add_point_to_m_circle = 3
           m_Circ(circle_no%).data(0).data0.in_point(3) = n_point% '
           Call add_m_circle_to_point(circle_no%, n_point%)
      '*****************************************************************************************************
     ElseIf n_point% = m_Circ(circle_no%).data(0).data0.in_point(1) Or _
                  n_point% = m_Circ(circle_no%).data(0).data0.in_point(2) Then '
      m_Circ(circle_no%).data(0).data0.c_coord = divide_POINTAPI_by_number( _
          add_POINTAPI(m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                    m_poi(m_Circ(circle_no%).data(0).data0.in_point(2)).data(0).data0.coordinate), 2)
      m_Circ(circle_no%).data(0).data0.radii = distance_of_two_POINTAPI(m_Circ(circle_no%).data(0).data0.c_coord, _
                           m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate)
      ele_(0).ty = point_
      ele_(0).no = m_Circ(circle_no%).data(0).data0.in_point(1)
      ele_(1).ty = point_
      ele_(1).no = m_Circ(circle_no%).data(0).data0.in_point(2)
            Call m_point_number(m_Circ(circle_no%).data(0).data0.c_coord, condition, 1, condition_color, "", _
                      ele_(0), ele_(1), 0, True)
            Call C_display_picture.draw_circle(circle_no%, 0, 0)
    End If
           m_Circ(circle_no%).data(0).input_type = condition
           Call Set_m_circle_data(circle_no%, m_Circ(circle_no%).data(0))
   End If
 Else
    If is_point_in_points(n_point%, m_Circ(circle_no%).data(0).data0.in_point) > 0 Then '判断输入点是否已在圆上
     Exit Function
    End If
     m_Circ(circle_no%).data(0).data0.in_point(0) = m_Circ(circle_no%).data(0).data0.in_point(0) + 1 '设置圆上的新点
      'add_point_to_m_circle = m_Circ(circle_no%).data(0).data0.in_point(0)
     m_Circ(circle_no%).data(0).data0.in_point(m_Circ(circle_no%).data(0).data0.in_point(0)) = n_point%
        Call set_parent(circle_, circle_no%, point_, n_point%, new_point_on_circle)
         Call add_m_circle_to_point(circle_no, n_point%)
   End If
 End If
' Call set_point_inform_for_circle(n_point%, circle_no%)
 If is_set_data Then
  add_point_to_m_circle = add_point_to_circle_(n_point%, circle_no%)
   If m_Circ(circle_no%).data(0).data0.in_point(0) >= 4 Then
      For i% = 1 To m_Circ(circle_no%).data(0).data0.in_point(0) - 3
       For j% = i% + 1 To m_Circ(circle_no%).data(0).data0.in_point(0) - 2
        For k% = j% + 1 To m_Circ(circle_no%).data(0).data0.in_point(0) - 1
        Call set_four_point_on_circle(m_Circ(circle_no%).data(0).data0.in_point(i%), _
               m_Circ(circle_no%).data(0).data0.in_point(j%), m_Circ(circle_no%).data(0).data0.in_point(k%), _
                 n_point%, circle_no%, record, 0, 0)
        Next k%
       Next j%
      Next i%
   End If
  If m_Circ(circle_no%).data(0).data0.tangent_line.element_no > 0 Then
     For i% = 1 To m_Circ(circle_no%).data(0).data0.tangent_line.element_no
       For j% = 1 To m_Circ(circle_no%).data(0).data0.in_point(0)
           If m_Circ(circle_no%).data(0).data0.in_point(j%) <> n_point% And _
               m_Circ(circle_no%).data(0).data0.in_point(j%) <> _
                 m_Circ(circle_no%).data(0).data0.tangent_line.tangent_celement(i).tangent_point Then
                  Call add_tangent_line_to_circle0( _
                    m_Circ(circle_no%).data(0).data0.tangent_line.tangent_celement(i).tangent_element_no, _
                     m_Circ(circle_no%).data(0).data0.tangent_line.tangent_celement(i).tangent_point, _
                      m_Circ(circle_no%).data(0).data0.in_point(j%), n_point%, record)
           End If
       Next j%
     Next i%
  End If
 Else
 If m_Circ(circle_no%).data(0).data0.in_point(0) = 0 Then
   m_Circ(circle_no%).data(0).data0.radii = abs_POINTAPI(minus_POINTAPI(m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.coordinate, _
                     m_poi(n_point%).data(0).data0.coordinate))
   Call C_display_picture.draw_circle(circle_no%, 0, 0)
 Else
  If m_Circ(circle_no%).data(0).data0.in_point(0) = 1 Then
   m_Circ(circle_no%).data(0).data0.in_point(0) = 2
   m_Circ(circle_no%).data(0).data0.in_point(2) = m_Circ(circle_no%).data(0).data0.in_point(1)
   m_Circ(circle_no%).data(0).data0.in_point(1) = m_Circ(circle_no%).data(0).data0.center
   m_Circ(circle_no%).data(0).data0.center = 0
   m_Circ(circle_no%).data(0).circle_type = 2
  End If
        m_Circ(circle_no%).data(0).data0.radii = _
          m_circle_radii(m_Circ(circle_no%).data(0))
   Call C_display_picture.draw_circle(circle_no%, 0, 0)
 End If
End If
End Function
Private Function add_point_to_circle_(ByVal p%, circle_no%) As Byte
Dim i%, j%, k%, o%, l%
Dim A(1) As Integer
Dim temp_record As total_record_type
Dim p_d(2) As Long
temp_record.record_data.data0 = input_record0
If is_point_in_circle(circle_no%, 0, p%, 0, 0) Then  '判断点是否在圆上
   Exit Function '在圆上
End If
If m_Circ(circle_no%).data(0).data0.in_point(0) > 1 Then
 If m_Circ(circle_no%).data(0).data0.in_point(0) > 3 Then
  For i% = 3 To m_Circ(circle_no%).data(0).data0.in_point(0) '
   If m_Circ(circle_no%).data(0).data0.in_point(i%) < p% Then
   For j% = 2 To i% - 1
    If m_Circ(circle_no%).data(0).data0.in_point(j%) < p% Then
     For k% = 1 To j% - 1
      If m_Circ(circle_no%).data(0).data0.in_point(k%) < p% Then
   add_point_to_circle_ = set_four_point_on_circle( _
  m_Circ(circle_no%).data(0).data0.in_point(i%), m_Circ(circle_no%).data(0).data0.in_point(j%), _
   m_Circ(circle_no%).data(0).data0.in_point(k%), p%, circle_no%, temp_record, 0, 0)
 If add_point_to_circle_ > 1 Then
  Exit Function
 End If
      End If
     Next k%
    End If
   Next j%
   End If
  Next i%
 End If
 For i% = 1 To m_Circ(circle_no%).data(0).data0.in_point(0)
  If m_Circ(circle_no%).data(0).data0.in_point(i%) <> p% And m_Circ(circle_no%).data(0).data0.center > 0 Then
  Call set_equal_dline(m_Circ(circle_no%).data(0).data0.center, m_Circ(circle_no%).data(0).data0.in_point(i%), _
   m_Circ(circle_no%).data(0).data0.center, p%, 0, 0, 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 0, False)
   For j% = 1 To i% - 1
    If m_Circ(circle_no%).data(0).data0.in_point(j%) <> p% Then
         add_point_to_circle_ = set_three_point_on_circle(m_Circ(circle_no%).data(0).data0.in_point(i%), _
           m_Circ(circle_no%).data(0).data0.in_point(j%), p%, 0, circle_no%, temp_record)
            If add_point_to_circle_ > 1 Then
             Exit Function
            End If
    End If
   Next j%
  End If
 Next i%
End If
For o% = 1 + last_conditions.last_cond(0).tangent_line_no To last_conditions.last_cond(1).tangent_line_no
i% = tangent_line(o%).data(0).record.data1.index.i(0)
temp_record.record_data.data0.condition_data.condition_no = 1
 temp_record.record_data.data0.condition_data.condition(1).ty = tangent_line_
  temp_record.record_data.data0.condition_data.condition(1).no = i%
 For k% = 0 To 1
 If tangent_line(i%).data(0).ele(k%).no = circle_no% And tangent_line(i%).data(0).ele(k%).ty = circle_ Then
   For l% = 0 To 1
    For j% = 1 To m_Circ(circle_no%).data(0).data0.in_point(0)
     If m_Circ(circle_no%).data(0).data0.in_point(j%) <> tangent_line(i%).data(0).poi(k%) And _
      m_Circ(circle_no%).data(0).data0.in_point(j%) <> p% Then
   A(0) = angle_number(m_lin(tangent_line(i%).data(0).line_no).data(0).data0.poi(l%), _
               tangent_line(i%).data(0).poi(k%), m_Circ(circle_no%).data(0).data0.in_point(j%), 0, 0)
   A(1) = angle_number(tangent_line(i%).data(0).poi(k%), p%, _
                m_Circ(circle_no%).data(0).data0.in_point(j%), 0, 0)
   If (A(0) > 0 And A(1) > 0) Or (A(0) < 0 And A(1) < 0) Then
    add_point_to_circle_ = set_three_angle_value( _
     Abs(A(0)), Abs(A(1)), 0, "1", "-1", "0", "0", _
       0, temp_record, 0, 0, 0, 0, 0, 0, False)
    If add_point_to_circle_ > 1 Then
     Exit Function
    End If
   End If
  End If
  Next j%
  Next l%
  End If
  Next k%
Next o%
End Function
Private Sub add_m_circle_to_point(ByVal circle_no%, ByVal point_no%)
Dim i%
If circle_no% = 0 Or point_no% = 0 Then
   Exit Sub
End If
For i% = 1 To m_poi(point_no%).data(0).in_circle(0)
    If m_poi(point_no%).data(0).in_circle(i%) = circle_no% Then
       Exit Sub
    End If
Next i%
m_poi(point_no%).data(0).in_circle(0) = _
   m_poi(point_no%).data(0).in_circle(0) + 1
    m_poi(point_no%).data(0).in_circle(m_poi(point_no%).data(0).in_circle(0)) = _
     circle_no%
     Call set_point_inform(point_no%, m_Circ(circle_no%).data(0).inform + _
           "上的点")
For i% = 1 To m_Circ(circle_no%).data(0).data0.in_point(0)
  If m_Circ(circle_no%).data(0).data0.in_point(i%) = point_no% Then
   Exit Sub
  End If
Next i%
 m_Circ(circle_no%).data(0).data0.in_point(0) = _
 m_Circ(circle_no%).data(0).data0.in_point(0) + 1
 m_Circ(circle_no%).data(0).data0.in_point(m_Circ(circle_no%).data(0).data0.in_point(0)) = _
    point_no%
End Sub

Private Sub set_point_inform_for_circle(ByVal point_no%, ByVal circle_no%)
If m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.visible > 0 Then
 Call set_point_inform(point_no%, "⊙" + _
             m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.name + "上的点")
ElseIf m_Circ(circle_no%).data(0).circle_type <> 1 Then
 Call set_point_inform(point_no%, "⊙" + _
             m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.name + _
               m_poi(m_Circ(circle_no%).data(0).data0.in_point(2)).data(0).data0.name + _
                m_poi(m_Circ(circle_no%).data(0).data0.in_point(3)).data(0).data0.name + "上的点")
End If
End Sub

Public Function combine_circle_with_circle(ByVal c%, no_reduce As Byte) As Byte
Dim i%
For i% = 1 To last_conditions.last_cond(1).circle_no
 If i% < c% Then
  combine_circle_with_circle = combine_two_circle(i%, c%)
   If combine_circle_with_circle > 1 Then
    Exit Function
   End If
 End If
Next i%
End Function
Public Function combine_two_circle(c1%, c2%) As Byte
Dim i%, j%, k%
Dim p(2) As Integer
Dim temp_record  As total_record_type
If m_Circ(c2%).data(0).data0.center = m_Circ(c1%).data(0).data0.center And m_Circ(c1%).data(0).circle_type = 1 Then
    If m_Circ(c2%).data(0).data0.in_point(1) = _
      m_Circ(c1%).data(0).data0.in_point(1) Then
        GoTo combine_two_circle_mark1
Else
For i% = 1 To m_Circ(c1%).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(c2%).data(0).data0.in_point(0)
  If m_Circ(c1%).data(0).data0.in_point(i%) = m_Circ(c2%).data(0).data0.in_point(j%) Then
     p(k%) = m_Circ(c1%).data(0).data0.in_point(i%)
      k% = k% + 1
       If k% = 3 Then
           combine_two_circle = True
            GoTo combine_two_circle_mark1
       End If
  End If
 Next j%
Next i%
GoTo combine_two_circle_mark0
combine_two_circle_mark1:
For i% = 1 To m_Circ(c2%).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(c1%).data(0).data0.in_point(0)
  If m_Circ(c2%).data(0).data0.in_point(i%) = m_Circ(c1%).data(0).data0.in_point(j%) Then
     GoTo combine_two_circle_mark2
  End If
 Next j%
combine_two_circle_mark2:
 Call add_point_to_m_circle(m_Circ(c2%).data(0).data0.in_point(i%), c1%, record0, True)
        combine_two_circle = True
Next i%
Exit Function
combine_two_circle_mark0:
If k% = 2 Then
record_0.data0.condition_data.condition_no = 0 ' record0
combine_two_circle = set_dverti(line_number0(p(0), p(1), 0, 0), _
   line_number0(m_Circ(c1%).data(0).data0.center, m_Circ(c2%).data(0).data0.center, 0, 0), _
      temp_record, 0, 0, False)
End If
End If
End If
End Function

Private Function is_old_m_circle(ByVal start%, ByVal cp%, radii&, ByVal p1%, ByVal p2%, ByVal p3%) As Integer
Dim i%
For i% = start% To last_conditions.last_cond(1).circle_no
    If m_circle_number0(cp%, radii&, p1%, p2%, p3%, m_Circ(i%).data(0)) Then
       is_old_m_circle = i%
        Exit Function
    End If
Next i%
End Function


Public Sub set_circle_data_from_input(io_circle_data As io_circle_data_type)

End Sub

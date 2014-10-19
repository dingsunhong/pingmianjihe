Attribute VB_Name = "reduce"
Option Explicit
Public Function inter_point_of_segment(p1%, ByVal l1%, ByVal t1 As Byte, _
      p2%, l2%, t2 As Byte) As Integer
'射线的交点
Dim i%, j%
Dim n(3) As Integer
Dim tl(1) As Integer
Dim start(1), en(1) As Integer
Dim tp%
'*********************************************************
'两直线重合
'**********************************************************
If l1% = l2% Then
 inter_point_of_segment = 0
  Exit Function
End If
'***************************************************
'确定起点的位置
'******************************************************
For i% = 1 To m_lin(l1%).data(0).data0.in_point(0)
  If m_lin(l1%).data(0).data0.in_point(i%) = p1% Then
   n(0) = i%
  GoTo inter_point_of_segment_mark1
  End If
Next i%
inter_point_of_segment_mark1:
For i% = 1 To m_lin(l2%).data(0).data0.in_point(0)
  If m_lin(l2%).data(0).data0.in_point(i%) = p2% Then
   n(1) = i%
    GoTo inter_point_of_segment_mark2
  End If
Next i%
inter_point_of_segment_mark2:
'*************************************************************
'确定循的上下限
'****************************************************************
If t1 = 1 Then
start(0) = n(0) + 1
  en(0) = m_lin(l1%).data(0).data0.in_point(0)
Else
start(0) = 1
 en(0) = n(0) - 1
End If
If t2 = 1 Then
 start(1) = n(1) + 1
  en(1) = m_lin(l2%).data(0).data0.in_point(0)
Else
start(1) = 1
 en(1) = n(1) - 1
End If
'**********************************************************
' 求交点
'************************************************************
For i% = start(0) To en(0)
 For j% = start(1) To en(1)
  If m_lin(l1%).data(0).data0.in_point(i%) = m_lin(l2%).data(0).data0.in_point(j%) Then
   inter_point_of_segment = m_lin(l1%).data(0).data0.in_point(i%)
    Exit Function
  End If
 Next j%
Next i%
End Function

'Public Function add_point_to_circle1(ByVal p%, ByVal c%, _
                   re As total_record_type, ByVal no_reduce As Byte) As Byte
'input_record0 = re.record_data.data0
'Call add_point_to_m_circle(p%, c%, record0, True)
'Dim A(1) As Integer
'Dim p_d(2) As Long
'End Function
Public Function angle_number(ByVal p1%, ByVal p2%, ByVal p3%, _
         degree As String, ByVal total_angle_no%) As Integer
'确定角的编号，负号表示于标准角相反
Dim i%, j%, no%
Dim n(1) As Integer
Dim ts As Long
Dim A As angle_data_type
 Dim t As Integer
  t = set_angle(p1%, p2%, p3%, A, degree)
   If t = 0 Then
    angle_number = 0
     Exit Function
   End If
angle_number = angle_number0(A, t, total_angle_no%)
'***********
End Function

Public Sub find_verti_center(ByVal A%, is_draw As Boolean)
Dim i%, j%, l%
Dim ty As Byte
Dim t(2) As Integer
If triangle(A%).data(0).verti_center = 0 Then
For i% = 0 To 2
 l% = m_poi(triangle(A%).data(0).poi(i%)).data(0).in_line(0)
 For j% = 1 To l%
   If is_dverti(line_number0(triangle(A%).data(0).poi((j% + 1) Mod 3), _
    triangle(A%).data(0).poi((j% + 2) Mod 3), 0, 0), j%, 0, -1000, 0, _
      0, 0, 0) Then
     t(i%) = j%
    End If
 Next j%
Next i%
For i% = 0 To 2
 If t(i%) > 0 And t((i% + 1) Mod 3) > 0 Then
ty = inter_point_line_line(m_lin(t(i%)).data(0).data0.poi(0), m_lin(t(i%)).data(0).data0.poi(1), _
 m_lin(t((i% + 1) Mod 3)).data(0).data0.poi(0), m_lin(t((i% + 1) Mod 3)).data(0).data0.poi(1), _
  0, 0, triangle(A%).data(0).verti_center, pointapi0, False, is_draw, False)
   If ty >= 0 And t((i% + 2) Mod 3) > 0 Then
    Call set_point_name(triangle(A%).data(0).verti_center, find_new_char)
    record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(triangle(A%).data(0).verti_center, t((i% + 2) Mod 3), _
      0, display, is_draw, 0)
     Exit Sub
    End If
  End If
Next i%
End If

End Sub
Public Function arc_no(ByVal p1%, ByVal c%, ByVal p2%) As Integer
'ty=0 确定优劣，ty=1  劣ty=2优
Dim Ag As angle_type
Dim t!
Dim Ar As arc_data_type
Dim no%, i%, j%
Dim n_(1) As Integer
't = set_angle(p1%, m_circ(c%).data(0).center, p2%, Ag, 0)
'If p1% = 0 Or p2% = 0 Or m_circ(c%).data(0).center = 0 Then
 'Arc_no = 0
  'Exit Function
'End If
If m_Circ(c%).data(0).data0.center > 0 Then
If read_line1(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate, _
   m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate, _
       t_coord, 0, 0, 6, False) Then
  Exit Function
End If
Else
If read_line1(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate, _
     m_Circ(c%).data(0).data0.c_coord, t_coord, 0, 0, 6, False) Then
  Exit Function
End If
End If
t! = (m_poi(p1%).data(0).data0.coordinate.X - m_Circ(c%).data(0).data0.c_coord.X) * _
       (m_poi(p2%).data(0).data0.coordinate.Y - m_Circ(c%).data(0).data0.c_coord.Y) - _
        (m_poi(p1%).data(0).data0.coordinate.Y - m_Circ(c%).data(0).data0.c_coord.Y) * _
         (m_poi(p2%).data(0).data0.coordinate.X - m_Circ(c%).data(0).data0.c_coord.X)
If t! > 0 Then
  Ar.cir = c%
   Ar.poi(0) = p1%
    Ar.poi(1) = p2%
     Ar.small_or_big = False
Else 'If t_% = -1 Then
    Ar.cir = c%
   Ar.poi(0) = p2%
    Ar.poi(1) = p1%
     Ar.small_or_big = True
End If
If search_for_arc(Ar, 1, 0, no%, 0) Then
arc_no = no%
Else
n_(0) = no%
Call search_for_arc(Ar, 1, 1, n_(1), 1)
If last_conditions.last_cond(1).arc_no Mod 10 = 0 Then
ReDim Preserve arc(last_conditions.last_cond(1).arc_no + 10) As arc_type
End If
last_conditions.last_cond(1).arc_no = last_conditions.last_cond(1).arc_no + 1
'arc(last_conditions.last_cond(1).arc_no).data(1) = arc_data_0
    arc(last_conditions.last_cond(1).arc_no).data(0) = Ar
 For j% = 0 To 1
 For i% = last_conditions.last_cond(1).arc_no To n_(j%) + 2 Step -1
 arc(i%).data(0).index(j%) = arc(i% - 1).data(0).index(j%)
 Next i%
 arc(n_(j%) + 1).data(0).index(j%) = last_conditions.last_cond(1).arc_no
 Next j%
  arc_no = last_conditions.last_cond(1).arc_no
End If
End Function
Public Function angle_number0(A As angle_data_type, t As Integer, ByVal total_angle_no%) As Integer
Dim i%, j%, no%, t_A%
Dim n(2) As Integer
'Dim l_2(2) As Single
Dim set_total_A As Boolean
'Dim te(2) As Byte
'D'im tl0(1) As Integer
'Dim te0(2) As Byte
Dim total_A As total_angle_data_type
Dim tA As angle_data_type
Dim ts!
'***********
If search_for_angle(A, no%, 0, 0) Then
       angle_number0 = no% * t
Else '新角
n(0) = no%
Call search_for_angle(A, n(1), 1, 1)
    If A.te(0) = 1 And A.te(1) = 1 Then
     A.total_no_ = 0
    ElseIf A.te(0) = 1 Then
     A.total_no_ = 1
    ElseIf A.te(1) = 1 Then
     A.total_no_ = 3
    Else
     A.total_no_ = 2
    End If
'l_2(0) = (m_poi(A.poi(0)).data(0).data0.coordinate.X - m_poi(A.poi(2)).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(A.poi(0)).data(0).data0.coordinate.Y - m_poi(A.poi(2)).data(0).data0.coordinate.Y) ^ 2
'l_2(1) = (m_poi(A.poi(0)).data(0).data0.coordinate.X - m_poi(A.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(A.poi(0)).data(0).data0.coordinate.Y - m_poi(A.poi(1)).data(0).data0.coordinate.Y) ^ 2
'l_2(2) = (m_poi(A.poi(1)).data(0).data0.coordinate.X - m_poi(A.poi(2)).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(A.poi(1)).data(0).data0.coordinate.Y - m_poi(A.poi(2)).data(0).data0.coordinate.Y) ^ 2
If total_angle_no% = 0 Then
 For i% = last_conditions.last_cond(0).total_angle_no + 1 To last_conditions.last_cond(1).total_angle_no
 t_A% = T_angle(i%).data(0).index(0)
  If is_same_two_point(T_angle(t_A%).data(0).line_no(0), T_angle(t_A%).data(0).line_no(1), A.line_no(0), A.line_no(1)) Then
    A.total_no = t_A% '已有全角
     set_total_A = False
      GoTo angle_number0_mark0
  End If
Next i%
  'If A.line_no(0) < A.line_no(1) Then
   total_A.line_no(0) = A.line_no(0)
   total_A.line_no(1) = A.line_no(1)
   total_A.inter_point = A.poi(1)
  'Else
   'total_A.line_no(3) = A.line_no(1)
   'total_A.line_no(2) = A.line_no(0)
  'End If
  set_total_A = True
Else
  t_A% = total_angle_no%
   set_total_A = False
End If
'*******************************
angle_number0_mark0:
If last_conditions.last_cond(1).angle_no Mod 10 = 0 Then
    ReDim Preserve angle(last_conditions.last_cond(1).angle_no + 10) As angle_type
End If
   last_conditions.last_cond(1).angle_no = last_conditions.last_cond(1).angle_no + 1
    angle(last_conditions.last_cond(1).angle_no).data(0) = angle_data_0
      angle(last_conditions.last_cond(1).angle_no).data(0) = A
      angle(last_conditions.last_cond(1).angle_no).data(0).other_no = last_conditions.last_cond(1).angle_no
    '******************
     For j% = 0 To 1
     For i = last_conditions.last_cond(1).angle_no To n(j%) + 2 Step -1
      angle(i%).data(0).index(j%) = angle(i% - 1).data(0).index(j%)
     Next i%
    angle(n(j%) + 1).data(0).index(j%) = last_conditions.last_cond(1).angle_no
     Next j%
    '****************************
    no% = last_conditions.last_cond(1).angle_no
    If set_total_A Then
     Call set_total_angle(total_A, -1, t_A%, False) '-1 表示未建立全角
    End If
      angle(no%).data(0).total_no = t_A%
       If T_angle(t_A%).data(0).value_no > 0 Then
        angle(no%).data(0).value_no = T_angle(t_A%).data(0).value_no
         If angle(no%).data(0).total_no_ = 0 Or _
               angle(no%).data(0).total_no_ = 2 Then
               angle(no%).data(0).value = T_angle(t_A%).data(0).value
         Else
               angle(no%).data(0).value = minus_string("180", _
                   T_angle(t_A%).data(0).value, True, False)
         End If
       End If
    '****************************
        angle_number0 = no% * t
       T_angle(t_A%).data(0).angle_no(angle(no%).data(0).total_no_).no = no%
       If T_angle(t_A%).data(0).is_used_no = -1 Then
          T_angle(t_A%).data(0).is_used_no = angle(no%).data(0).total_no_
       End If
       ''If l_2(1) + l_2(2) > l_2(0) + 10 Then
       '  T_angle(t_A%).data(0).angle_no(angle(no%).data(0).total_no_).sh = True
       'End If
  ts = (CSng(m_poi(A.poi(0)).data(0).data0.coordinate.X) - m_poi(A.poi(2)).data(0).data0.coordinate.X) * _
       (m_poi(A.poi(0)).data(0).data0.coordinate.X - m_poi(A.poi(2)).data(0).data0.coordinate.X) - _
        (CSng(m_poi(A.poi(0)).data(0).data0.coordinate.X) - m_poi(A.poi(1)).data(0).data0.coordinate.X) * _
         (m_poi(A.poi(0)).data(0).data0.coordinate.X - m_poi(A.poi(1)).data(0).data0.coordinate.X) - _
           (CSng(m_poi(A.poi(2)).data(0).data0.coordinate.X) - m_poi(A.poi(1)).data(0).data0.coordinate.X) * _
            (m_poi(A.poi(2)).data(0).data0.coordinate.X - m_poi(A.poi(1)).data(0).data0.coordinate.X)
  ts = ts + (CSng(m_poi(A.poi(0)).data(0).data0.coordinate.Y) - m_poi(A.poi(2)).data(0).data0.coordinate.Y) * _
       (m_poi(A.poi(0)).data(0).data0.coordinate.Y - m_poi(A.poi(2)).data(0).data0.coordinate.Y) - _
        (m_poi(A.poi(0)).data(0).data0.coordinate.Y - m_poi(A.poi(1)).data(0).data0.coordinate.Y) * _
         (CSng(m_poi(A.poi(0)).data(0).data0.coordinate.Y) - m_poi(A.poi(1)).data(0).data0.coordinate.Y) - _
           (CSng(m_poi(A.poi(2)).data(0).data0.coordinate.Y) - m_poi(A.poi(1)).data(0).data0.coordinate.Y) * _
            (m_poi(A.poi(2)).data(0).data0.coordinate.Y - m_poi(A.poi(1)).data(0).data0.coordinate.Y)
 End If

End Function
